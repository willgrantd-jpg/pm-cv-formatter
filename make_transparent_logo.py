"""
Strip white background from pm-logo.png → pm-logo-ui.png (transparent RGBA PNG).
Uses PyMuPDF (already a dependency) + stdlib only.
"""
import struct, zlib, fitz

SRC = r"C:\Users\W.Dier\Downloads\pm-cv-formatter-app\assets\pm-logo.png"
DST = r"C:\Users\W.Dier\Downloads\pm-cv-formatter-app\assets\pm-logo-ui.png"

def write_rgba_png(samples_rgba: bytearray, w: int, h: int, path: str):
    """Write a raw RGBA bytearray as a valid PNG file."""
    def chunk(tag: bytes, data: bytes) -> bytes:
        body = tag + data
        return struct.pack(">I", len(data)) + body + struct.pack(">I", zlib.crc32(body) & 0xFFFFFFFF)

    signature = b"\x89PNG\r\n\x1a\n"
    ihdr = chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 6, 0, 0, 0))

    raw = bytearray()
    for y in range(h):
        raw.append(0)  # filter type: None
        raw += samples_rgba[y * w * 4 : (y + 1) * w * 4]

    idat = chunk(b"IDAT", zlib.compress(bytes(raw), 9))
    iend = chunk(b"IEND", b"")

    with open(path, "wb") as f:
        f.write(signature + ihdr + idat + iend)


doc = fitz.open(SRC)
pix = doc[0].get_pixmap(alpha=True)
assert pix.n == 4, f"Expected RGBA, got {pix.n} channels"

samples = bytearray(pix.samples)
THRESHOLD = 238          # pixels this bright or brighter → transparent
EDGE_SMOOTH = 245        # semi-transparent near-edge pixels

for i in range(0, len(samples), 4):
    r, g, b = samples[i], samples[i+1], samples[i+2]
    brightness = (r + g + b) / 3
    if brightness >= THRESHOLD:
        # Fully transparent
        samples[i+3] = 0
    elif brightness >= EDGE_SMOOTH - 10:
        # Soft edge: partial transparency
        samples[i+3] = int((THRESHOLD - brightness) / (THRESHOLD - EDGE_SMOOTH + 10) * 255)

write_rgba_png(samples, pix.width, pix.height, DST)
print(f"Done → {DST}  ({pix.width}×{pix.height})")
