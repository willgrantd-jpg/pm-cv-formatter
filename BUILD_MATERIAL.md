# PM CV Formatter — Complete Build Material

> Feed this document to Claude to generate user manuals, onboarding guides,
> technical documentation, deployment runbooks, or any other build artefact.
> Everything needed to understand, maintain, and extend the application is here.

---

## 1. What the App Does

**PM CV Formatter** is a local Windows web application built for Patrick Morgan, a recruitment firm. It takes a candidate's raw CV (.docx, .pdf, or .txt) and produces a branded Word document (.docx) that matches the Patrick Morgan "PM Profile" template design — automatically, using the Anthropic Claude AI API.

The recruiter opens the app in their browser at `http://localhost:5000`, drags a CV onto the page, optionally adds consultant notes or attaches a supplementary document, hits **Format CV**, and receives a ready-to-send branded profile in seconds.

**Key outputs:**
- A `{CandidateName}_PM_Profile.docx` file in the `outputs/` folder
- An in-browser live preview of the formatted profile
- One-click download from the Recent Outputs panel

---

## 2. Technology Stack

| Layer | Technology |
|---|---|
| Backend | Python 3.10+, Flask 3.x |
| AI extraction | Anthropic Claude API (`claude-haiku-4-5`) |
| Word generation | python-docx + direct OOXML manipulation |
| PDF reading | PyMuPDF (fitz) |
| Frontend | Vanilla HTML/CSS/JS (no framework), Inter + JetBrains Mono fonts |
| Launcher | Windows Batch (.bat) + VBScript (.vbs) + PowerShell (.ps1) |
| Packaging | ZIP archive with silent launcher |

---

## 3. File Structure

```
PM CV Formatter/                   ← top-level folder (extracted from ZIP)
│
├── app.py                         ← Flask server (routes, config, upload handling)
├── extract.py                     ← CV text extraction + Claude API call
├── populate.py                    ← Word document builder (OOXML)
│
├── start.bat                      ← Primary launcher (Python check, pip install, shortcut)
├── launch.vbs                     ← Silent launcher (used by Desktop shortcut)
├── _create_shortcut.ps1           ← Creates branded Desktop shortcut on first run
│
├── requirements.txt               ← Python dependencies
├── config.json                    ← Stores API key (created on first save)
│
├── templates/
│   └── index.html                 ← Full single-page UI (HTML + CSS + JS)
│
├── static/
│   ├── favicon.png                ← Browser tab icon
│   └── favicon.ico                ← Desktop shortcut icon (multi-resolution)
│
├── assets/
│   ├── PM_CV_Template_v2.docx     ← Base Word template (header logo, footer, margins)
│   ├── pm-logo.png                ← Full Patrick Morgan logo (4957×1150 px)
│   └── pm-logo-ui.png             ← Transparent-background version for the UI header
│
├── outputs/                       ← Generated profiles saved here (git-ignored)
│
└── make_transparent_logo.py       ← Utility: strips white bg from pm-logo.png
```

---

## 4. Architecture & Data Flow

```
User drops CV onto browser
        │
        ▼
[index.html JS]
  builds FormData:
    cv_file        ← the CV
    extra_notes    ← consultant chat notes (joined)
    supplementary_doc ← optional attached PDF/DOCX
        │
        ▼ POST /format
[app.py — format_cv()]
  1. Saves CV to temp file
  2. Extracts supplementary doc text (if any), appends to extra_notes
  3. Calls process_cv() → extract.py
        │
        ▼
[extract.py — process_cv()]
  1. read_cv_text() — extracts plain text from .docx / .pdf / .txt
  2. extract_cv_data() — calls Claude API with system prompt
  3. Claude returns structured JSON
  4. Post-processing: enforce position field, inject English language,
     normalise LinkedIn URL, clean nulls
  Returns: dict
        │
        ▼
[app.py]
  Calls populate_template() → populate.py
        │
        ▼
[populate.py — populate_template()]
  1. Opens PM_CV_Template_v2.docx (preserves header/footer/logo)
  2. Clears body content
  3. Rebuilds document body from structured data:
       - Candidate name + LinkedIn badge
       - "CANDIDATE PROFILE" sub-label with navy rule
       - Summary paragraph (Georgia italic)
       - Selected Impact box (cream bg, gold left border)
       - Experience table (two-column: dates | content)
       - Education table
       - Languages & Skills table
       - Certifications table
  4. Fixes footer logo (swaps image bytes, sets size)
  5. Saves .docx to outputs/
  6. Post-save: injects spell/grammar check suppress via ZIP manipulation
        │
        ▼
[app.py]
  Returns JSON: { ok: true, filename: "...", data: { full structured dict } }
        │
        ▼
[index.html JS]
  - Shows success banner
  - Renders live preview from data dict
  - Refreshes Recent Outputs panel
  - Download link becomes active
```

---

## 5. Source Code — Complete Files

### 5.1 `app.py`

```python
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


def load_config() -> dict:
    if CONFIG_PATH.exists():
        with open(CONFIG_PATH, "r") as f:
            return json.load(f)
    return {}


def save_config(cfg: dict):
    with open(CONFIG_PATH, "w") as f:
        json.dump(cfg, f, indent=2)


def get_api_key() -> str:
    return os.environ.get("ANTHROPIC_API_KEY") or load_config().get("api_key", "")


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

    import tempfile
    tmp = tempfile.NamedTemporaryFile(suffix=ext, delete=False)
    cv_file.save(tmp.name)
    tmp.close()

    extra_notes = request.form.get("extra_notes", "").strip()

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
                    pass

    try:
        data = process_cv(tmp.name, api_key, extra_notes=extra_notes)

        name = data.get("name", "Unknown")
        safe_name = re.sub(r"[^\w\s-]", "", name).strip().replace(" ", "_")
        output_filename = f"{safe_name}_PM_Profile.docx"
        output_path = OUTPUTS_DIR / output_filename
        OUTPUTS_DIR.mkdir(exist_ok=True)

        populate_template(str(TEMPLATE_PATH), data, str(output_path))

        return jsonify({
            "ok":       True,
            "filename": output_filename,
            "data":     data,
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


@app.route("/clear-outputs", methods=["POST"])
def clear_outputs():
    OUTPUTS_DIR.mkdir(exist_ok=True)
    for f in OUTPUTS_DIR.iterdir():
        if f.suffix == ".docx":
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
        reverse=True
    )
    resp = jsonify(files)
    resp.headers["Cache-Control"] = "no-store"
    return resp


if __name__ == "__main__":
    port = 5000
    url  = f"http://localhost:{port}"
    print(f"\n  PM CV Formatter running at {url}\n")
    threading.Timer(1.2, lambda: webbrowser.open(url)).start()
    app.run(host="127.0.0.1", port=port, debug=False)
```

---

### 5.2 `extract.py`

```python
"""
extract.py — CV text extraction + Claude API JSON parsing.
"""

import json
import os
import re
import tempfile
import zipfile

import anthropic


SYSTEM_PROMPT = """You are a CV data extractor for Patrick Morgan, a recruitment firm.
Extract structured data from a candidate CV following these rules exactly:

RULES:
0. Language: If the CV is written in a language other than English, translate all content
   into English during extraction. Preserve exactly as written — do NOT translate:
   company names, institution/university names, certification names, city and country names,
   and any other proper nouns. Translate everything else — job titles, bullet points,
   summaries, degree subject names, skill labels.
1. Zero hallucination — every word must come from the CV or supplied notes. If not present → null.
2. Position field: ALWAYS set to "Candidate Profile" — never the candidate's job title.
3. Third person only. Convert first-person ("I managed") to third person. Infer pronoun from CV;
   if unclear use "They". This is the ONLY permitted change to any text.
4. Dates: 4-digit years only ("2019", "Present"). Never ranges in a single field.
5. Bullets: copy the candidate's exact words verbatim. Do NOT paraphrase, condense, summarise,
   embellish, or improve the language in any way. Only convert first-person to third-person
   where necessary — every other word must be preserved exactly as written.
6. Experience: include EVERY role listed on the CV, most recent first. Do not omit any position
   regardless of how many there are.
7. Summary: 3-5 sentences, third person, factual and neutral in tone. Cover: seniority level,
   functional expertise, sector/industry focus, one standout differentiator. Draw only from facts
   on the CV — do not add colour, marketing language, or interpretation. Do NOT invent details.
   - If ADDITIONAL NOTES are supplied, weave in any relevant factual context from them.
8. Key Achievement: single most impressive quantified result from the CV. Headline 6-8 words.
   Up to 3 supporting bullets — copied verbatim (third person only).
   - If ADDITIONAL NOTES contain a stronger or more specific achievement, use it.
9. Skills: up to 6, explicitly listed only. No inferred skills.
10. Languages: extract only additional languages beyond English. Normalise fluency to:
    Native, Fluent, Conversational, or Basic. Do NOT include English here — it is added automatically.
11. LinkedIn: extract the candidate's LinkedIn profile URL if present anywhere in the CV
    (header, footer, contact section). Return as a full URL. If the URL is missing "https://",
    prepend it. If not found → null.
12. Max slots: all experience (no limit), 3 education, 5 languages, 6 skills, 6 certifications.
13. Promotions at same company: one entry per distinct title.
14. Certifications: professional certifications, licences, or accreditations explicitly listed.
    Include name, issuing body, and year if stated. If not present → empty array [].

Return ONLY valid JSON matching this exact schema — no markdown, no commentary:

{
  "name": "Full Name",
  "position": "Candidate Profile",
  "linkedin_url": "https://linkedin.com/in/username or null",
  "summary": "3-5 sentence third-person summary.",
  "achievement": {
    "title": "Six to eight word headline",
    "bullet1": "Quantified supporting detail",
    "bullet2": "Second detail or null",
    "bullet3": "Third detail or null"
  },
  "languages": [
    {"name": "French", "fluency": "Fluent"}
  ],
  "experience": [
    {
      "startYear": "2020",
      "endYear": "Present",
      "company": "Company Name",
      "position": "Job Title",
      "bullet1": "Key responsibility or achievement",
      "bullet2": "Second bullet or null",
      "bullet3": "Third bullet or null"
    }
  ],
  "education": [
    {
      "startYear": "2015",
      "endYear": "2018",
      "institution": "University Name",
      "degree": "BSc Economics",
      "bullet1": "Honours or notable detail, or null",
      "bullet2": null
    }
  ],
  "skills": ["Skill 1", "Skill 2"],
  "certifications": [
    {
      "name": "Certification Name",
      "issuer": "Issuing Body or null",
      "year": "2022 or null"
    }
  ]
}"""


def read_docx_text(path: str) -> str:
    with zipfile.ZipFile(path, "r") as z:
        with z.open("word/document.xml") as f:
            xml_bytes = f.read()
    text = re.sub(rb"<[^>]+>", b" ", xml_bytes)
    text = text.decode("utf-8", errors="replace")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def read_pdf_text(path: str) -> str:
    try:
        import fitz
        doc = fitz.open(path)
        pages = [page.get_text() for page in doc]
        return "\n\n".join(pages)
    except ImportError:
        raise RuntimeError("PyMuPDF is not installed. Run: pip install pymupdf")


def read_cv_text(path: str) -> str:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".docx":
        return read_docx_text(path)
    elif ext == ".pdf":
        return read_pdf_text(path)
    elif ext in (".txt", ".md"):
        with open(path, "r", encoding="utf-8", errors="replace") as f:
            return f.read()
    else:
        raise ValueError(f"Unsupported file type: {ext}. Upload a .docx or .pdf file.")


def extract_cv_data(cv_text: str, api_key: str, extra_notes: str = "") -> dict:
    client = anthropic.Anthropic(api_key=api_key)

    user_content = f"Extract structured data from this CV:\n\n{cv_text}"
    if extra_notes and extra_notes.strip():
        user_content += (
            f"\n\n---\nADDITIONAL NOTES (call transcript / consultant additions):\n"
            f"{extra_notes.strip()}\n\n"
            "Use these notes to enrich the summary and key achievement where relevant. "
            "All other fields (experience, education, skills, languages) must still come "
            "from the CV only."
        )

    message = client.messages.create(
        model="claude-haiku-4-5",
        max_tokens=4096,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_content}]
    )

    raw = message.content[0].text.strip()
    raw = re.sub(r"^```(?:json)?\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)

    data = json.loads(raw)

    # Enforce position field
    data["position"] = "Candidate Profile"

    # Always include English as first language; deduplicate if already present
    languages = [l for l in (data.get("languages") or []) if l]
    has_english = any(
        isinstance(l, dict) and l.get("name", "").strip().lower() == "english"
        for l in languages
    )
    if not has_english:
        languages.insert(0, {"name": "English", "fluency": "Native"})
    data["languages"] = languages

    # Normalise LinkedIn URL
    linkedin = (data.get("linkedin_url") or "").strip()
    if linkedin and not linkedin.startswith("http"):
        linkedin = "https://" + linkedin.lstrip("/")
    data["linkedin_url"] = linkedin or None

    def clean(obj):
        if isinstance(obj, dict):
            return {k: clean(v) for k, v in obj.items()}
        if isinstance(obj, list):
            return [clean(v) for v in obj]
        return obj

    return clean(data)


def process_cv(cv_path: str, api_key: str, extra_notes: str = "") -> dict:
    cv_text = read_cv_text(cv_path)
    if len(cv_text.strip()) < 100:
        ext = os.path.splitext(cv_path)[1].lower()
        if ext == ".pdf":
            raise ValueError(
                "Could not extract text from this PDF — it may be a scanned image. "
                "Please export it as a searchable PDF or supply a Word (.docx) version."
            )
        raise ValueError("The uploaded file appears to be empty or unreadable.")
    return extract_cv_data(cv_text, api_key, extra_notes=extra_notes)
```

---

### 5.3 `populate.py`

Full source — see the file directly. Key sections summarised:

**Brand constants:**
```python
NAVY  = "1B2A4E"   # dark navy — headings, name
GOLD  = "C8A961"   # PM gold — accents, bullets, linkedin badge
GRAY  = "6B7280"   # muted gray — dates, sub-labels
BODY  = "1A1A1A"   # near-black — body text
CREAM = "F4F1EA"   # warm cream — Selected Impact box background
```

**Column widths (twips):**
```python
COL_DATE    = 1800   # ~3.17 cm — left date column
COL_CONTENT = 7560   # ~13.3 cm — right content column
TBL_WIDTH   = 9360   # total body width
```

**Document sections built in order:**
1. Candidate name (Calibri 28pt bold navy) + optional LinkedIn badge
2. "CANDIDATE PROFILE" sub-label (Calibri 9pt gray, 200 tracking) + navy bottom rule
3. Summary paragraph (Georgia 11pt italic)
4. Selected Impact box (cream fill, gold left border, gold "SELECTED IMPACT" label)
5. EXPERIENCE section — two-column table, gold em-dash bullets
6. EDUCATION section — two-column table
7. LANGUAGES & SKILLS section — two-column table, dot-separated values
8. CERTIFICATIONS section — two-column table

**LinkedIn badge (`_add_linkedin_badge`):**
- Creates a `w:hyperlink` element with `r:id` pointing to external URL
- Run styled: Calibri 11pt bold, PM gold colour, gold character border, superscript, no underline
- Text: `'   in'` (3 spaces + "in") — appears as a gold outlined badge next to the name

**Footer fix (`_fix_footer`):**
- Swaps the image bytes in the footer's image Part with the full `pm-logo.png`
- Sets display size to 1.2" wide (maintaining 4957:1150 aspect ratio)
- Removes "Patrick Morgan  ·  " text run from footer text paragraph
- Grows bottom margin to 1800 twips

**Proofing suppression (`_suppress_proofing`):**
- Directly manipulates the saved .docx ZIP
- Injects `<w:hideSpellingErrors w:val="1"/>` and `<w:hideGrammaticalErrors w:val="1"/>` into `word/settings.xml`

---

### 5.4 `start.bat`

```bat
@echo off
setlocal EnableDelayedExpansion
cd /d "%~dp0"
title PM CV Formatter

echo.
echo  ============================================
echo   Patrick Morgan  ^|  CV Formatter
echo  ============================================
echo.

REM ── Guard: must be run from an extracted folder, not inside a ZIP ─────────
if not exist "%~dp0app.py" (
    echo  [ERROR] Cannot find app files.
    echo  Please extract the ZIP first, then run start.bat from the extracted folder.
    pause
    exit /b 1
)

REM ── Python check ─────────────────────────────────────────────────────────
python --version >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Python not found. Install from https://www.python.org/downloads/
    echo  Tick "Add Python to PATH" during install.
    start https://www.python.org/downloads/
    pause
    exit /b 1
)

for /f "tokens=2" %%v in ('python --version 2^>^&1') do set PYVER=%%v
echo  Python !PYVER! detected.

REM ── Install dependencies (first run only) ────────────────────────────────
pip show flask >nul 2>&1
if errorlevel 1 (
    echo  Installing dependencies -- please wait, this only happens once...
    pip install -r requirements.txt --quiet
    if errorlevel 1 (
        echo  [ERROR] Dependency installation failed. Try running as Administrator.
        pause
        exit /b 1
    )
    echo  Dependencies installed.
)

REM ── Ensure pymupdf is present ─────────────────────────────────────────────
pip show pymupdf >nul 2>&1
if errorlevel 1 (
    pip install pymupdf --quiet
)

REM ── Create desktop shortcut on first run ─────────────────────────────────
if not exist "%APPDATA%\pm-cv-formatter.flag" (
    echo  Creating desktop shortcut...
    set "APPDIR=%~dp0"
    if "!APPDIR:~-1!"=="\" set "APPDIR=!APPDIR:~0,-1!"
    powershell -NoProfile -ExecutionPolicy Bypass -File "!APPDIR!\_create_shortcut.ps1" "!APPDIR!" >nul 2>&1
    if %ERRORLEVEL% EQU 0 (
        echo 1> "%APPDATA%\pm-cv-formatter.flag"
        echo  Shortcut added to your Desktop.
    )
)

REM ── Launch ────────────────────────────────────────────────────────────────
echo  Starting PM CV Formatter...
echo  Your browser will open automatically.
echo.

set PYTHONUTF8=1
start "" pythonw app.py
```

**Critical design notes:**
- Uses `pythonw.exe` (not `python.exe`) so no console window remains open
- `setlocal EnableDelayedExpansion` is required for `!APPDIR!` variable expansion
- Trailing backslash stripping (`if "!APPDIR:~-1!"=="\"`) prevents `\"` escaping bugs in quoted paths
- Uses `if %ERRORLEVEL% EQU 0` (not `if not errorlevel 1`) to detect exact success, as PowerShell failure returns negative exit codes
- `PYTHONUTF8=1` prevents encoding errors on non-UTF-8 Windows terminals

---

### 5.5 `launch.vbs`

```vbscript
Set fso   = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
shell.Run """" & scriptDir & "\start.bat""", 0, False
```

- Window style `0` = completely hidden — no terminal flashes on screen
- The Desktop shortcut points to `wscript.exe launch.vbs`, not directly to `start.bat`
- `WScript.Shell.Run` with style 0 suppresses the cmd.exe window entirely

---

### 5.6 `_create_shortcut.ps1`

```powershell
param([string]$AppDir)

$AppDir    = $AppDir.TrimEnd('\')
$LaunchVbs = Join-Path $AppDir "launch.vbs"
$IconPath  = Join-Path $AppDir "static\favicon.ico"
$Desktop   = [Environment]::GetFolderPath("Desktop")
$Shortcut  = Join-Path $Desktop "PM CV Formatter.lnk"

try {
    $WshShell             = New-Object -ComObject WScript.Shell
    $lnk                  = $WshShell.CreateShortcut($Shortcut)
    $lnk.TargetPath       = "wscript.exe"
    $lnk.Arguments        = """$LaunchVbs"""
    $lnk.WorkingDirectory = $AppDir
    $lnk.Description      = "PM CV Formatter"
    $lnk.WindowStyle      = 1
    if (Test-Path $IconPath) { $lnk.IconLocation = "$IconPath,0" }
    $lnk.Save()
} catch {
    exit 1
}

exit 0
```

- Creates shortcut at `%USERPROFILE%\Desktop\PM CV Formatter.lnk`
- Icon sourced from `static\favicon.ico` (multi-resolution: 256, 64, 48, 32, 16 px)
- Shortcut calls `wscript.exe "C:\path\to\launch.vbs"` — not the .bat directly
- Flag file `%APPDATA%\pm-cv-formatter.flag` prevents re-running on subsequent launches

---

### 5.7 `requirements.txt`

```
flask>=3.0.0
anthropic>=0.40.0
python-docx>=1.1.0
pymupdf>=1.24.0
lxml>=5.0.0
```

---

## 6. UI Design System

**Design language:** Linear.app dark aesthetic — near-black backgrounds, subtle borders, minimal chrome.

**Design tokens (CSS variables):**
```
--bg-base:        #08090a   ← page background
--bg-elevated:    #101113   ← card backgrounds
--bg-overlay:     #18191b   ← nested elements
--bg-hover:       #1e2022   ← hover states
--border-subtle:  rgba(255,255,255,0.05)
--border-default: rgba(255,255,255,0.09)
--border-focus:   rgba(255,255,255,0.22)
--text-primary:   #e8e9ea
--text-secondary: #8b8d91
--text-tertiary:  #55575b
--pm-blue:        #4a7fb8   ← primary action colour
--pm-gold:        #c9a84c   ← brand accent
--success:        #3d9e6e
--error:          #c0443a
--font-sans:      'Inter'
--font-mono:      'JetBrains Mono'
```

**Layout:** Three-column CSS grid at max-width 1180px
```
grid-template-columns: 260px 1fr 230px
```
- Left 260px: Consultant Notes panel (sticky)
- Centre flexible: Upload card + Preview card (stacked)
- Right 230px: Recent Outputs panel (sticky)

**Background texture:** `radial-gradient` dot grid at 24px spacing, 2% white opacity.

---

## 7. UI Panels & Features

### Left Panel — Consultant Notes
- Chat-bubble display of added notes
- Textarea for typing new notes
- "Add note" button — pushes text into `chatNotes[]` array, renders bubble
- Per-bubble remove button (×)
- "Attach supplementary doc" dashed button — opens hidden file input
- Attached file shown as chip with filename + remove button
- All notes and supplementary doc sent with every Format CV call in session

### Centre — Upload Card
- Drag-and-drop zone + click-to-browse (`<input type="file" multiple>`)
- Accepts `.docx`, `.pdf`, `.txt`
- File list with per-file status: `pending` / `busy` (pulsing blue dot) / `ok` / `error`
- Error status shows tooltip with the actual error message on hover
- Download button per completed file
- "Format CV" button — processes all pending files sequentially
- Progress bar + spinner with "Processing N of M: filename" text
- Result banner (success/error) after completion

### Centre — Preview Card
- Shows live in-browser preview rendered from the structured JSON returned by Claude
- White document background, Calibri/Georgia fonts matching the real Word output
- Sections: Name + LinkedIn badge, Summary, Achievement box, Experience, Education, Skills/Languages, Certifications
- Navigation arrows for multi-CV sessions (1 / N counter)
- Download button linking to the actual .docx file

### Right Panel — Recent Outputs
- Lists all `.docx` files in `outputs/` folder, most recent first
- Click filename or download icon to download
- "Clear" button (with confirm dialog) deletes all .docx from outputs folder
- Auto-refreshes after each format run and on page load

### Settings Modal
- Gear icon (⚙) in header
- Password input for Anthropic API key
- Key saved to `config.json` locally — never sent anywhere except Anthropic
- API key also readable from `ANTHROPIC_API_KEY` environment variable (takes priority)

### Header Logo
- `pm-logo-ui.png` — white-background stripped version of `pm-logo.png`
- CSS filter: `brightness(0) invert(1)` → pure white mark, then:
  `drop-shadow(0 0 6px rgba(201,168,76,0.75)) drop-shadow(0 0 16px rgba(201,168,76,0.35))`
  → PM gold inner halo + wider ambient glow
- No white pill background — transparent, floats directly on dark header

---

## 8. Claude AI Integration

**Model:** `claude-haiku-4-5`
**Max tokens:** 4096
**API library:** `anthropic` Python SDK

**System prompt enforces (14 rules):**
1. Non-English CVs → translated to English (proper nouns preserved)
2. Zero hallucination — only facts from the CV
3. `position` field always "Candidate Profile"
4. Third-person voice (only permitted change to candidate's words)
5. Verbatim bullets — no paraphrasing, embellishment, or improvement
6. ALL experience entries — no cap
7. Summary: 3-5 sentences, neutral, factual
8. Key Achievement: best quantified result, 6-8 word headline
9. Skills: up to 6, only explicitly listed
10. Languages: additional beyond English only
11. LinkedIn URL: extracted and normalised
12. Slot limits: 3 education, 5 languages, 6 skills, 6 certifications
13. Promotions at same company: one entry per title
14. Certifications: professional certs with name, issuer, year

**Post-processing (Python, not Claude):**
- `position` forced to "Candidate Profile" regardless of Claude output
- English injected as first language if not present
- LinkedIn URL normalised to `https://` prefix
- All JSON nulls cleaned

---

## 9. Word Document Design

**Template:** `assets/PM_CV_Template_v2.docx`
- Provides: page size (A4), margins, header (PM diamond logo), footer (PM logo + page numbers)
- Body content is cleared and fully rebuilt on each run

**Typography:**
- Name: Calibri 28pt bold navy (`#1B2A4E`)
- Sub-label: Calibri 9pt gray, 200 tracking, uppercase
- Summary: Georgia 11pt italic
- Section headings: Calibri 10pt bold navy, uppercase, 60 tracking, navy bottom rule
- Company/Institution: Calibri 11pt bold navy
- Role/Degree: Calibri 9pt italic gold (`#C8A961`)
- Bullet text: Calibri 9pt body (`#1A1A1A`), gold em-dash prefix
- Dates: Calibri 9pt gray

**Selected Impact box:**
- Background fill: cream (`#F4F1EA`)
- Left border: gold, 18 half-points wide
- "SELECTED IMPACT" label: gold 8pt bold, 150 tracking
- Title row: bold navy headline + gray supporting detail on same line
- Additional rows: bold navy standalone bullets

**Two-column table structure:**
- No borders (invisible grid)
- Date column: 1800 twips (~3.2 cm)
- Content column: 7560 twips (~13.3 cm)
- Thin warm-gray divider lines between entries

**Footer:**
- Full PM logo image (pm-logo.png), 1.2" wide
- Page number: `PAGE / NUMPAGES` field
- "Patrick Morgan · " text run removed programmatically

---

## 10. Distribution & Installation

**Packaging:**
- ZIP file: `PM CV Formatter.zip`
- Single top-level folder: `PM CV Formatter\`
- All app files included; `outputs/` folder empty; `config.json` can include a pre-baked API key

**First-run experience:**
1. User extracts ZIP to any location
2. User double-clicks `start.bat`
3. Batch checks Python is installed (prompts download if not)
4. Installs Python dependencies from `requirements.txt` (one-time, ~30 seconds)
5. Creates branded Desktop shortcut (`PM CV Formatter.lnk`) pointing to `wscript.exe launch.vbs`
6. Writes flag file `%APPDATA%\pm-cv-formatter.flag` so steps 4-5 are skipped on future runs
7. Launches Flask server silently with `pythonw.exe`
8. Opens browser at `http://localhost:5000`

**Subsequent runs (via Desktop shortcut):**
- Double-click shortcut → `wscript.exe` runs `launch.vbs` with window style 0 → `start.bat` runs hidden → `pythonw app.py` starts → browser opens
- No terminal window ever visible

**Requirements:**
- Windows 10 or 11
- Python 3.10+ (installer at python.org — tick "Add Python to PATH")
- Internet connection (first run for pip install; every run for Claude API)
- Anthropic API key (entered once in Settings)

---

## 11. API Key Management

**Storage:** `config.json` in the app folder:
```json
{
  "api_key": "sk-ant-..."
}
```

**Priority order:**
1. `ANTHROPIC_API_KEY` environment variable (for IT-managed deployments)
2. `config.json` `api_key` field (for user-entered keys)

**Pre-baked key:** For team distribution, the `config.json` can be included in the ZIP with the firm's API key already set. Users never need to touch Settings.

**Where the key is used:** Only sent to `api.anthropic.com` in the `x-api-key` header. Never logged, never sent anywhere else.

---

## 12. Known Limitations & Edge Cases

| Issue | Behaviour | Mitigation |
|---|---|---|
| Scanned PDF (image-only) | PyMuPDF extracts no text | Error message shown: "may be a scanned image — export as searchable PDF" |
| Very long CVs (10+ pages) | May approach 4096 token output limit | Claude will truncate gracefully; some later entries may be omitted |
| Non-Latin characters in filename | Saved with `re.sub(r"[^\w\s-]", "")` stripping | Unicode letters preserved; only special characters removed |
| Multiple simultaneous uploads | Processed sequentially in JS for-loop | Reliable; each file gets its own API call |
| App already running on port 5000 | Second instance fails silently | Kill with `Get-Process pythonw | Stop-Process -Force` in PowerShell |
| Running .bat from inside ZIP | `app.py` not found guard triggers | Clear error message with extraction instructions |

---

## 13. Structured Data Schema (Claude Output)

```json
{
  "name": "Full Name",
  "position": "Candidate Profile",
  "linkedin_url": "https://linkedin.com/in/username",
  "summary": "3-5 sentence third-person summary.",
  "achievement": {
    "title": "Six to eight word headline",
    "bullet1": "Quantified supporting detail",
    "bullet2": "Second detail or null",
    "bullet3": "Third detail or null"
  },
  "languages": [
    {"name": "English", "fluency": "Native"},
    {"name": "French",  "fluency": "Fluent"}
  ],
  "experience": [
    {
      "startYear": "2020",
      "endYear":   "Present",
      "company":   "Company Name",
      "position":  "Job Title",
      "bullet1":   "Key responsibility or achievement",
      "bullet2":   "Second bullet or null",
      "bullet3":   "Third bullet or null"
    }
  ],
  "education": [
    {
      "startYear":   "2015",
      "endYear":     "2018",
      "institution": "University Name",
      "degree":      "BSc Economics",
      "bullet1":     "Honours or notable detail, or null",
      "bullet2":     null
    }
  ],
  "skills": ["Skill 1", "Skill 2"],
  "certifications": [
    {
      "name":   "Certification Name",
      "issuer": "Issuing Body or null",
      "year":   "2022 or null"
    }
  ]
}
```

---

## 14. Flask Routes Reference

| Method | Route | Purpose |
|---|---|---|
| GET | `/` | Serve main UI (`index.html`) |
| GET | `/logo` | Serve `pm-logo-ui.png` |
| POST | `/save-key` | Save API key to `config.json` |
| POST | `/format` | Process CV — main endpoint |
| GET | `/download/<filename>` | Download a generated .docx |
| POST | `/clear-outputs` | Delete all .docx from `outputs/` |
| GET | `/outputs` | List generated files (JSON, no-cache) |

---

## 15. Utility Scripts (Not Part of Runtime)

| Script | Purpose |
|---|---|
| `make_transparent_logo.py` | Strips white background from `pm-logo.png` → `pm-logo-ui.png` using PyMuPDF + stdlib PNG writer. Run once when the logo changes. |
| `make_ico.py` | Generates `static/favicon.ico` multi-resolution ICO from `static/favicon.png`. Run once. |
| `test_populate.py` | Smoke test for the Word document builder |
| `test_v2.py` | Integration test for the v2 template |
| `validate_output.py` | Validates a generated .docx structure |

---

*End of build material.*
