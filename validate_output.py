"""
validate_output.py — quick structural check on a generated docx.
Usage: python validate_output.py [path_to_docx]
"""
import sys
import zipfile
from pathlib import Path
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def _w(t):
    return "{%s}%s" % (W, t)

path = sys.argv[1] if len(sys.argv) > 1 else str(
    Path(__file__).parent / "outputs" / "TEST_PM_Profile.docx"
)

print(f"\nValidating: {path}\n")

with zipfile.ZipFile(path) as z:
    xml = z.read("word/document.xml")

root = etree.fromstring(xml)
body = root.find(_w("body"))

errors = []

# Check: no <w:p> directly inside <w:sdtContent> when sdtContent is inside a <w:p>
for outer_p in body.iter(_w("p")):
    for sdt in outer_p.findall(".//" + _w("sdt")):
        content = sdt.find(_w("sdtContent"))
        if content is not None:
            if content.find(_w("p")) is not None:
                pr = sdt.find(_w("sdtPr"))
                tag = ""
                if pr is not None:
                    tg = pr.find(_w("tag"))
                    if tg is not None:
                        tag = tg.get(_w("val"), "")
                errors.append(
                    "INLINE SDT '%s' has <w:p> inside sdtContent -- INVALID OOXML" % tag
                )

# Check: every sdt has sdtContent
for sdt in body.iter(_w("sdt")):
    if sdt.find(_w("sdtContent")) is None:
        pr = sdt.find(_w("sdtPr"))
        tag = ""
        if pr is not None:
            tg = pr.find(_w("tag"))
            if tg is not None:
                tag = tg.get(_w("val"), "")
        errors.append("SDT '%s' missing sdtContent" % tag)

body_count  = len(list(body))
table_count = len(body.findall(_w("tbl")))
sdt_count   = sum(1 for _ in body.iter(_w("sdt")))

print("  Body elements : %d" % body_count)
print("  Tables        : %d" % table_count)
print("  SDTs (total)  : %d" % sdt_count)
print()

if errors:
    print("  ERRORS (%d):" % len(errors))
    for e in errors:
        print("    !! " + e)
    sys.exit(1)
else:
    print("  OK -- no structural violations detected.")
    print("  The file should open in Word without an error dialog.\n")
