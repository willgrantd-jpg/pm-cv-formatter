"""
app.py — PM CV Formatter local web app.

Run with:  python app.py
Then open: http://localhost:5000

Drag-drop a CV (docx or pdf) → get back a formatted PM Profile .docx.
"""

import json
import os
import re
import threading
import webbrowser
from pathlib import Path

from flask import Flask, jsonify, render_template, request, send_file

from extract  import process_cv
from populate import populate_template

# ── Config ────────────────────────────────────────────────────────────────────
BASE_DIR      = Path(__file__).parent
TEMPLATE_PATH = BASE_DIR / "assets" / "PM_CV_Template_v2.docx"
LOGO_PATH     = BASE_DIR / "assets" / "pm-logo-ui.png"
OUTPUTS_DIR   = BASE_DIR / "outputs"
CONFIG_PATH   = BASE_DIR / "config.json"

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024   # 20 MB upload limit


# ── Config helpers ────────────────────────────────────────────────────────────

def load_config() -> dict:
    if CONFIG_PATH.exists():
        with open(CONFIG_PATH, "r") as f:
            return json.load(f)
    return {}


def save_config(cfg: dict):
    with open(CONFIG_PATH, "w") as f:
        json.dump(cfg, f, indent=2)


def get_api_key() -> str:
    # Priority: env var → config file
    return os.environ.get("ANTHROPIC_API_KEY") or load_config().get("api_key", "")


# ── Routes ────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    cfg = load_config()
    has_key = bool(get_api_key())
    return render_template("index.html", has_key=has_key, api_key=cfg.get("api_key", ""))


@app.route("/logo")
def serve_logo():
    if LOGO_PATH.exists():
        return send_file(str(LOGO_PATH), mimetype="image/png")
    return "", 404


@app.route("/linkedin-badge")
def serve_linkedin_badge():
    badge_path = BASE_DIR / "assets" / "linkedin-badge.png"
    if badge_path.exists():
        return send_file(str(badge_path), mimetype="image/png")
    return "", 404


@app.route("/save-key", methods=["POST"])
def save_key():
    key = request.json.get("api_key", "").strip()
    if not key:
        return jsonify({"error": "No API key provided"}), 400
    cfg = load_config()
    cfg["api_key"] = key
    save_config(cfg)
    return jsonify({"ok": True})


@app.route("/format", methods=["POST"])
def format_cv():
    api_key = get_api_key()
    if not api_key:
        return jsonify({"error": "No API key configured. Enter your Anthropic API key in Settings."}), 400

    if "cv_file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    cv_file = request.files["cv_file"]
    if not cv_file.filename:
        return jsonify({"error": "Empty filename"}), 400

    ext = Path(cv_file.filename).suffix.lower()
    if ext not in (".docx", ".pdf", ".txt"):
        return jsonify({"error": f"Unsupported file type '{ext}'. Upload a .docx or .pdf file."}), 400

    # Save upload to a temp path
    import tempfile
    tmp = tempfile.NamedTemporaryFile(suffix=ext, delete=False)
    cv_file.save(tmp.name)
    tmp.close()

    extra_notes      = request.form.get("extra_notes",      "").strip()
    structural_notes = request.form.get("structural_notes", "").strip()

    # Extract text from any supplementary document and append to notes
    supp_tmp = None
    if "supplementary_doc" in request.files:
        supp_file = request.files["supplementary_doc"]
        if supp_file.filename:
            supp_ext = Path(supp_file.filename).suffix.lower()
            if supp_ext in (".docx", ".pdf", ".txt"):
                import tempfile as _tf
                supp_tmp = _tf.NamedTemporaryFile(suffix=supp_ext, delete=False)
                supp_file.save(supp_tmp.name)
                supp_tmp.close()
                try:
                    from extract import read_cv_text
                    supp_text = read_cv_text(supp_tmp.name)
                    if supp_text.strip():
                        extra_notes = (extra_notes + "\n\n" if extra_notes else "") + \
                            f"SUPPLEMENTARY DOCUMENT ({supp_file.filename}):\n{supp_text}"
                except Exception:
                    pass  # supplementary doc extraction failure is non-fatal

    try:
        # Step 1: Extract structured JSON via Claude API
        data = process_cv(tmp.name, api_key, extra_notes=extra_notes)

        # Step 2: Build output filename
        name = data.get("name", "Unknown")
        safe_name = re.sub(r"[^\w\s-]", "", name).strip().replace(" ", "_")
        output_filename = f"{safe_name}_PM_Profile.docx"
        output_path = OUTPUTS_DIR / output_filename
        OUTPUTS_DIR.mkdir(exist_ok=True)

        # Step 3: Populate template (XML-first, no Word COM)
        populate_template(str(TEMPLATE_PATH), data, str(output_path),
                          structural_notes=structural_notes)

        # Save source CV for re-render (keep alongside the output docx)
        import shutil
        source_name = f"{safe_name}_source{ext}"
        source_path = OUTPUTS_DIR / source_name
        shutil.copy2(tmp.name, str(source_path))

        # Parse structural flags to send to preview
        from populate import _parse_structural
        s_flags = _parse_structural(structural_notes)

        return jsonify({
            "ok":               True,
            "filename":         output_filename,
            "source_filename":  source_name,
            "structural_flags": s_flags,
            "data":             data,
        })

    except json.JSONDecodeError as e:
        return jsonify({"error": f"Claude returned invalid JSON: {e}"}), 500
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        try:
            os.unlink(tmp.name)
        except Exception:
            pass
        if supp_tmp:
            try:
                os.unlink(supp_tmp.name)
            except Exception:
                pass


@app.route("/download/<filename>")
def download(filename):
    # Sanitise — only allow simple filenames from the outputs folder
    safe = re.sub(r"[^\w\-. ]", "", filename)
    path = OUTPUTS_DIR / safe
    if not path.exists():
        return "File not found", 404
    return send_file(
        str(path),
        as_attachment=True,
        download_name=safe,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )


@app.route("/rerender", methods=["POST"])
def rerender_cv():
    api_key = get_api_key()
    if not api_key:
        return jsonify({"error": "No API key configured"}), 400

    body             = request.json or {}
    source_filename  = re.sub(r"[^\w\-. ]", "", body.get("source_filename", ""))
    extra_notes      = body.get("extra_notes",      "").strip()
    structural_notes = body.get("structural_notes", "").strip()

    source_path = OUTPUTS_DIR / source_filename
    if not source_path.exists():
        return jsonify({"error": "Original CV not found — please re-upload the file."}), 404

    try:
        data = process_cv(str(source_path), api_key, extra_notes=extra_notes)

        name        = data.get("name", "Unknown")
        safe_name   = re.sub(r"[^\w\s-]", "", name).strip().replace(" ", "_")
        out_name    = f"{safe_name}_PM_Profile.docx"
        output_path = OUTPUTS_DIR / out_name
        OUTPUTS_DIR.mkdir(exist_ok=True)

        populate_template(str(TEMPLATE_PATH), data, str(output_path),
                          structural_notes=structural_notes)

        from populate import _parse_structural
        s_flags = _parse_structural(structural_notes)

        return jsonify({
            "ok":               True,
            "filename":         out_name,
            "source_filename":  source_filename,
            "structural_flags": s_flags,
            "data":             data,
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/clear-outputs", methods=["POST"])
def clear_outputs():
    OUTPUTS_DIR.mkdir(exist_ok=True)
    for f in OUTPUTS_DIR.iterdir():
        if f.suffix in (".docx", ".pdf", ".txt"):
            try:
                f.unlink()
            except Exception:
                pass
    return jsonify({"ok": True})


@app.route("/outputs")
def list_outputs():
    OUTPUTS_DIR.mkdir(exist_ok=True)
    files = sorted(
        [
            {
                "name":    f.name,
                "size_kb": round(f.stat().st_size / 1024, 1),
                "mtime":   f.stat().st_mtime,
            }
            for f in OUTPUTS_DIR.iterdir()
            if f.suffix == ".docx"
        ],
        key=lambda x: x["mtime"],
        reverse=True   # most recent first
    )
    resp = jsonify(files)
    resp.headers["Cache-Control"] = "no-store"
    return resp


# ── Launch ────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    port = 5000
    url  = f"http://localhost:{port}"
    print(f"\n  PM CV Formatter running at {url}\n")
    # Open browser after a short delay so Flask has time to start
    threading.Timer(1.2, lambda: webbrowser.open(url)).start()
    app.run(host="127.0.0.1", port=port, debug=False)
