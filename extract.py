"""
extract.py — CV text extraction + Claude API JSON parsing.

Reads a CV file (docx or PDF) and calls the Claude API to produce
the structured JSON the populate module needs.
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
0. Language: If the CV is written in a language other than English, translate all content into English during extraction. Preserve exactly as written — do NOT translate: company names, institution/university names, certification names, city and country names, and any other proper nouns. Translate everything else — job titles, bullet points, summaries, degree subject names, skill labels.
1. Zero hallucination — every word must come from the CV or supplied notes. If not present → null.
2. Position field: ALWAYS set to "Candidate Profile" — never the candidate's job title.
3. Third person only. Convert first-person ("I managed") to third person. Infer pronoun from CV; if unclear use "They". This is the ONLY permitted change to any text.
4. Dates: 4-digit years only ("2019", "Present"). Never ranges in a single field.
5. Bullets: copy the candidate's exact words verbatim. Do NOT paraphrase, condense, summarise, embellish, or improve the language in any way. Only convert first-person to third-person where necessary — every other word must be preserved exactly as written.
   - Include EVERY bullet point listed under each role — there is no limit. Never drop, merge, or summarise bullets to save space.
   - Wrapped lines: PDF and docx extraction often splits a single long bullet across two or more lines with no bullet marker on the continuation lines. Join these into one string. A new bullet only begins when the source has a new leading marker (•, –, -, *, or a fresh numbered item). Continuation lines with no marker must be appended to the previous bullet.
6. Experience: include EVERY role listed on the CV, most recent first. Do not omit any position regardless of how many there are.
   - Single role at a company: use the flat entry format with "position" and "bullets" at the top level.
   - Multiple distinct roles at the same company (promotions/progression): use the grouped format with a "roles" array. Set the overall date span at the top level and nest each role inside "roles" (most recent first). Each role has its own "startYear", "endYear", "position", and "bullets".
7. Summary: 3-5 sentences, third person, factual and neutral in tone. Cover: seniority level, functional expertise, sector/industry focus, one standout differentiator. Draw only from facts on the CV — do not add colour, marketing language, or interpretation. Do NOT invent details.
   - If ADDITIONAL NOTES are supplied, weave in any relevant factual context from them.
   - If CONSULTANT INSTRUCTIONS include a Target role, frame the summary to position the candidate for that role.
   - If CONSULTANT INSTRUCTIONS include a Sector focus, lead with that sector and make it prominent.
   - If CONSULTANT INSTRUCTIONS include a Tone instruction, adjust the register of the summary accordingly (e.g. "more senior" = emphasise leadership and strategic scope; "more technical" = lead with technical depth).
8. Key Achievement: single most impressive quantified result from the CV. Headline 6-8 words. Up to 3 supporting bullets — copied verbatim (third person only).
   - If ADDITIONAL NOTES contain a stronger or more specific achievement, use it.
   - If CONSULTANT INSTRUCTIONS include a Sector focus or Target role, prefer achievements relevant to those.
9. Skills: up to 6, explicitly listed only. No inferred skills.
10. Languages: extract only additional languages beyond English. Normalise fluency to: Native, Fluent, Conversational, or Basic. Do NOT include English here — it is added automatically.
11. LinkedIn: extract the candidate's LinkedIn profile URL if present anywhere in the CV (header, footer, contact section). Return as a full URL. If the URL is missing "https://", prepend it. If not found → null.
12. Max slots: all experience (no limit), 3 education, 5 languages, 6 skills, 6 certifications.
13. Promotions at same company: use the grouped "roles" format described in Rule 6. One role object per distinct title — do not merge separate roles into one.
14. Certifications: professional certifications, licences, or accreditations explicitly listed. Include name, issuing body, and year if stated. If not present → empty array [].
15. CONSULTANT INSTRUCTIONS (power commands): If the prompt contains a CONSULTANT INSTRUCTIONS block, treat each instruction as a binding directive:
    - "Target role: X"     → frame summary and achievement to position the candidate for role X
    - "Sector focus: X"    → lead with sector X in the summary; prioritise related achievements
    - "Tone: X"            → adjust summary register (e.g. more senior, more technical, more commercial)
    - "Exclude: X"         → omit the named company or role from the experience section entirely
    - "Emphasise: X"       → give extra weight to X in the summary and achievement selection
    - Any other instruction → apply it as best you can using only facts already present in the CV or notes.

Return ONLY valid JSON matching this exact schema — no markdown, no commentary:

{
  "name": "Full Name",
  "position": "Candidate Profile",
  "linkedin_url": "https://linkedin.com/in/username or null",
  "summary": "3-5 sentence third-person summary.",
  "achievement": {
    "title": "Six to eight word headline",
    "bullets": [
      "Quantified supporting detail — verbatim",
      "Second detail or omit if not present",
      "Third detail or omit if not present"
    ]
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
      "bullets": [
        "First bullet copied verbatim",
        "Second bullet copied verbatim",
        "All bullets — no limit, include every one from the CV"
      ]
    },
    {
      "startYear": "2015",
      "endYear": "Present",
      "company": "Company With Multiple Roles",
      "roles": [
        {
          "startYear": "2020",
          "endYear": "Present",
          "position": "Director",
          "bullets": ["Bullet verbatim", "Another bullet verbatim"]
        },
        {
          "startYear": "2015",
          "endYear": "2020",
          "position": "Manager",
          "bullets": ["Bullet verbatim"]
        }
      ]
    }
  ],
  "education": [
    {
      "startYear": "2015",
      "endYear": "2018",
      "institution": "University Name",
      "degree": "BSc Economics",
      "bullets": ["Honours or notable detail — omit if nothing to add"]
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
    """Extract plain text from a .docx file without Word COM."""
    with zipfile.ZipFile(path, "r") as z:
        with z.open("word/document.xml") as f:
            xml_bytes = f.read()

    # Strip XML tags, normalise whitespace
    text = re.sub(rb"<[^>]+>", b" ", xml_bytes)
    text = text.decode("utf-8", errors="replace")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def read_pdf_text(path: str) -> str:
    """Extract text from a PDF using PyMuPDF (fitz)."""
    try:
        import fitz  # PyMuPDF
        doc = fitz.open(path)
        pages = [page.get_text() for page in doc]
        return "\n\n".join(pages)
    except ImportError:
        raise RuntimeError(
            "PyMuPDF is not installed. Run: pip install pymupdf"
        )


def read_cv_text(path: str) -> str:
    """Read CV text from docx or pdf."""
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



# Power command prefixes recognised in consultant notes
_POWER_COMMANDS = [
    "target role:",
    "sector focus:",
    "tone:",
    "exclude:",
    "emphasise:",
    "emphasize:",
]


def _parse_notes(raw_notes: str) -> tuple[list[str], list[str]]:
    """
    Split raw consultant notes into:
      - power_commands : lines that start with a recognised prefix
      - free_notes     : everything else
    Returns (power_commands, free_notes) as lists of strings.
    """
    power_commands = []
    free_notes = []
    for line in raw_notes.splitlines():
        stripped = line.strip()
        if not stripped:
            continue
        lower = stripped.lower()
        if any(lower.startswith(p) for p in _POWER_COMMANDS):
            power_commands.append(stripped)
        else:
            free_notes.append(stripped)
    return power_commands, free_notes


def extract_cv_data(cv_text: str, api_key: str, extra_notes: str = "") -> dict:
    """
    Call Claude API with CV text and return parsed JSON dict.
    Uses claude-haiku-4-5 for speed and cost efficiency.
    extra_notes: optional call transcript or consultant additions.
    """
    client = anthropic.Anthropic(api_key=api_key)

    user_content = f"Extract structured data from this CV:\n\n{cv_text}"

    if extra_notes and extra_notes.strip():
        power_commands, free_notes = _parse_notes(extra_notes)

        # Inject power commands as a clearly labelled binding block
        if power_commands:
            user_content += (
                "\n\n---\nCONSULTANT INSTRUCTIONS (binding — apply these exactly):\n"
                + "\n".join(f"• {cmd}" for cmd in power_commands)
                + "\n"
            )

        # Inject free-form notes as contextual additions
        if free_notes:
            user_content += (
                "\n\n---\nADDITIONAL NOTES (call transcript / consultant context):\n"
                + "\n".join(free_notes)
                + "\n\n"
                "Use these notes to enrich the summary and key achievement where relevant. "
                "All other fields (experience, education, skills, languages) must still come "
                "from the CV only."
            )

    message = client.messages.create(
        model="claude-haiku-4-5",
        max_tokens=4096,
        system=SYSTEM_PROMPT,
        messages=[
            {
                "role": "user",
                "content": user_content
            }
        ]
    )

    raw = message.content[0].text.strip()

    # Strip markdown code fences if Claude wrapped the JSON
    raw = re.sub(r"^```(?:json)?\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)

    data = json.loads(raw)

    # Enforce position field is always "Candidate Profile"
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

    # Normalise LinkedIn URL — ensure https:// prefix
    linkedin = (data.get("linkedin_url") or "").strip()
    if linkedin and not linkedin.startswith("http"):
        linkedin = "https://" + linkedin.lstrip("/")
    data["linkedin_url"] = linkedin or None

    # Normalise bullets — convert old bullet1/2/3 format to bullets array
    def _normalise_bullets(entry: dict) -> dict:
        if 'bullets' not in entry:
            bullets = []
            for i in range(1, 10):
                b = entry.pop(f'bullet{i}', None)
                if b and str(b).strip() and str(b).strip().lower() != 'null':
                    bullets.append(str(b).strip())
            entry['bullets'] = bullets
        else:
            entry['bullets'] = [
                b for b in (entry.get('bullets') or [])
                if b and str(b).strip() and str(b).strip().lower() != 'null'
            ]
        return entry

    # Normalise experience entries (flat and multi-role)
    for entry in (data.get('experience') or []):
        if not entry:
            continue
        if 'roles' in entry:
            for role in (entry.get('roles') or []):
                if role:
                    _normalise_bullets(role)
        else:
            _normalise_bullets(entry)

    # Normalise education bullets
    for entry in (data.get('education') or []):
        if entry:
            _normalise_bullets(entry)

    # Normalise achievement bullets
    ach = data.get('achievement')
    if isinstance(ach, dict):
        _normalise_bullets(ach)

    # Normalise nulls — convert JSON null to Python None throughout
    def clean(obj):
        if isinstance(obj, dict):
            return {k: clean(v) for k, v in obj.items()}
        if isinstance(obj, list):
            return [clean(v) for v in obj]
        return obj

    return clean(data)


def process_cv(cv_path: str, api_key: str, extra_notes: str = "") -> dict:
    """Read CV file and extract structured data. Returns dict."""
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


# ── Stax screening-question mapping ───────────────────────────────────────────

# Fixed screening questions defined by the retained client (Stax). Order matters.
STAX_SCREENING_QUESTIONS = [
    {"group": "", "question": "Service line, summary of functional responsibilities, reporting lines"},
    {"group": "", "question": "Buy-side, sell-side, or both?"},
    {"group": "", "question": "Are they owning workstreams? If so, which workstreams?"},
    {"group": "", "question": "Are they leading projects end-to-end? How many at one time? How many led in total?"},
    {"group": "", "question": "Typical team size, structure, and duration for their projects?"},
    {"group": "", "question": "How long have they been in a project management role?"},
    {"group": "", "question": "Experience managing junior team members"},
    {"group": "", "question": "Experience leading client readouts"},
    {"group": "", "question": "Experience on proposals"},
    {"group": "", "question": "Description of client engagement responsibilities"},
    {"group": "", "question": "Vertical / industry exposure"},
    {"group": "Experience", "question": "Project mix (CDD % / Strategy % / Other %)"},
    {"group": "Experience", "question": "CDD volume (total / annual / typical length in weeks)"},
    {"group": "Motivation", "question": "Why are they looking to move or explore opportunities?"},
    {"group": "Motivation", "question": "Compensation expectations and current salary"},
    {"group": "Logistics", "question": "Location"},
    {"group": "Logistics", "question": "Availability (notice period) and availability for interviews"},
    {"group": "Logistics", "question": "Visa requirements / US citizen?"},
]


SCREENING_SYSTEM_PROMPT = """You map a recruiter's raw screening-call notes onto a fixed list of client screening questions for Patrick Morgan.

You are given N numbered questions and the recruiter's notes. For EACH question, produce the answer using ONLY the notes.

RULES:
1. Facts only from the notes — never invent or infer facts, numbers, names, dates, or claims. Keep every figure, percentage and number exactly as written. If the notes do not address a question, return an empty list for it.
2. Beef it up: the consultant captures notes quickly, so the source is terse shorthand. Rewrite each point into a clear, well-formed, professional bullet — complete the sentence, fix grammar, and spell out obvious shorthand — WITHOUT adding any fact that is not in the notes.
3. Neutral, impersonal voice: write subject-less bullets that start with the verb or noun. NEVER use first person ("I led", "I own") and NEVER use third-person pronouns ("they led", "the candidate manages"). Just "Led...", "Owns...", "Manages...". Examples: note "i led 3 cdd projects end to end" -> "Led three commercial due diligence projects end-to-end."; note "owns 2 workstreams buy side" -> "Owns two workstreams, focused on buy-side."
4. Each answer is a list of bullet strings. Usually one or two bullets; split into several only when the notes contain genuinely distinct points.
5. Do not move a fact to a question it does not answer. If unsure where a fact belongs, place it under the single closest-matching question only.

Return ONLY valid JSON, no markdown, no commentary, in exactly this shape:
{"answers": [["bullet", "bullet"], [], ["bullet"], ...]}
The "answers" array MUST have exactly N inner lists, in the same order as the questions."""


def extract_screening(notes: str, api_key: str) -> list:
    """Map raw screening notes onto the fixed Stax question list.

    Returns a list of {"group", "question", "bullets"} in canonical order.
    Answers come only from the notes; questions not covered get empty bullets.
    Never raises — on any failure returns the question skeleton with no answers.
    """
    questions = STAX_SCREENING_QUESTIONS
    skeleton = [
        {"group": q["group"], "question": q["question"], "bullets": []}
        for q in questions
    ]

    if not notes or not notes.strip():
        return skeleton

    numbered = "\n".join(f"{i+1}. {q['question']}" for i, q in enumerate(questions))
    user_content = (
        f"There are {len(questions)} questions.\n\n"
        f"QUESTIONS:\n{numbered}\n\n"
        f"RECRUITER SCREENING NOTES:\n{notes.strip()}"
    )

    try:
        client = anthropic.Anthropic(api_key=api_key)
        message = client.messages.create(
            model="claude-haiku-4-5",
            max_tokens=4096,
            system=SCREENING_SYSTEM_PROMPT,
            messages=[{"role": "user", "content": user_content}],
        )
        raw = message.content[0].text.strip()
        raw = re.sub(r"^```(?:json)?\s*", "", raw)
        raw = re.sub(r"\s*```$", "", raw)
        parsed = json.loads(raw)
        answers = parsed.get("answers", []) if isinstance(parsed, dict) else []
    except Exception:
        return skeleton

    result = []
    for i, q in enumerate(questions):
        bullets = []
        if i < len(answers) and isinstance(answers[i], list):
            bullets = [str(b).strip() for b in answers[i] if b and str(b).strip()]
        result.append({"group": q["group"], "question": q["question"], "bullets": bullets})
    return result
