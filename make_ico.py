"""Generate a multi-resolution favicon.ico from favicon.png using PyMuPDF."""
import struct
import fitz

src = r"C:\Users\W.Dier\Downloads\pm-cv-formatter-app\static\favicon.png"
dst = r"C:\Users\W.Dier\Downloads\pm-cv-formatter-app\static\favicon.ico"

sizes = [256, 64, 48, 32, 16]

doc = fitz.open(src)
page = doc[0]

entries = []
for sz in sizes:
    mat = fitz.Matrix(sz / page.rect.width, sz / page.rect.height)
    pix = page.get_pixmap(matrix=mat, alpha=True)
    png_bytes = pix.tobytes("png")
    entries.append((sz, png_bytes))

num_images = len(entries)
header = struct.pack("<HHH", 0, 1, num_images)

dir_size = num_images * 16
data_offset = 6 + dir_size

directory = b""
image_data = b""
for (sz, png_bytes) in entries:
    w = 0 if sz >= 256 else sz
    h = 0 if sz >= 256 else sz
    directory += struct.pack("<BBBBHHII",
        w, h, 0, 0, 1, 32,
        len(png_bytes),
        data_offset + len(image_data)
    )
    image_data += png_bytes

with open(dst, "wb") as f:
    f.write(header + directory + image_data)

print(f"Done: {dst}  ({len(entries)} sizes: {[s for s,_ in entries]})")
