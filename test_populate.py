"""
test_populate.py — run this directly to verify populate.py works without Flask.

Usage:  python test_populate.py
Output: outputs/TEST_PM_Profile.docx
"""

import sys
from pathlib import Path

BASE = Path(__file__).parent
sys.path.insert(0, str(BASE))

from populate import populate_template

TEMPLATE = str(BASE / "assets" / "PM_CV_Template_1.docx")
OUTPUT   = str(BASE / "outputs" / "TEST_PM_Profile.docx")

# Minimal realistic data — 2 roles, 1 language, 2 education
TEST_DATA = {
    "name":    "Alessandro Ricci",
    "position":"Candidate Profile",
    "summary": "Senior finance professional with 15 years of experience across investment banking and private equity. Proven track record of driving value creation in complex transactions.",
    "achievement": {
        "title":   "Led €450m Cross-Border Acquisition",
        "bullet1": "Managed end-to-end due diligence across five jurisdictions",
        "bullet2": "Delivered 18% IRR above fund hurdle rate",
        "bullet3": "",
    },
    "languages": [
        {"name": "English", "fluency": "Native"},
        {"name": "Italian", "fluency": "Fluent"},
    ],
    "experience": [
        {
            "startYear": "2019",
            "endYear":   "Present",
            "company":   "Goldman Sachs",
            "position":  "Managing Director",
            "bullet1":   "Led origination of £2.1bn in M&A mandates across TMT sector",
            "bullet2":   "Built and managed a team of 12 analysts and associates",
            "bullet3":   "Advised on 8 cross-border transactions exceeding $500m each",
        },
        {
            "startYear": "2014",
            "endYear":   "2019",
            "company":   "Morgan Stanley",
            "position":  "Vice President",
            "bullet1":   "Executed 14 leveraged buyout transactions totalling $3.4bn",
            "bullet2":   "Developed proprietary deal sourcing network across DACH region",
            "bullet3":   "",
        },
    ],
    "education": [
        {
            "startYear":   "2001",
            "endYear":     "2004",
            "institution": "London School of Economics",
            "degree":      "BSc Economics (First Class Honours)",
            "bullet1":     "",
            "bullet2":     "",
        },
        {
            "startYear":   "2004",
            "endYear":     "2005",
            "institution": "INSEAD",
            "degree":      "MBA",
            "bullet1":     "",
            "bullet2":     "",
        },
    ],
    "skills": [
        "M&A Transaction Execution",
        "LBO Modelling",
        "Capital Markets",
        "Cross-Border Deals",
        "Team Leadership",
        "Due Diligence",
    ],
}

def main():
    print(f"\n  Template : {TEMPLATE}")
    print(f"  Output   : {OUTPUT}\n")

    try:
        result = populate_template(TEMPLATE, TEST_DATA, OUTPUT)
        print(f"  OK  Success: {result}")
        print()
        print("  Open the file in Word to verify it looks correct.")
        print("  If Word opens it without an error dialog, the build is working.\n")
    except Exception as e:
        import traceback
        print(f"\n  FAIL  Error: {e}\n")
        traceback.print_exc()

if __name__ == "__main__":
    main()
