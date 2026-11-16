"""presentation/*.pptx -> site/slides/N.jpg + manifest.json (with real speaker notes)."""
import glob, json, os, shutil, subprocess
from pptx import Presentation

src = sorted(glob.glob("presentation/*.pptx"))[0]
shutil.rmtree("build", ignore_errors=True); shutil.rmtree("site", ignore_errors=True)
os.makedirs("build"); os.makedirs("site/slides")
subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", "build", src], check=True)
pdf = glob.glob("build/*.pdf")[0]
subprocess.run(["pdftoppm", "-jpeg", "-jpegopt", "quality=88", "-r", "130", pdf, "build/p"], check=True)
imgs = sorted(glob.glob("build/p-*.jpg"))
for i, f in enumerate(imgs, 1):
    shutil.copy(f, f"site/slides/{i}.jpg")

prs = Presentation(src)
notes = []
for s in prs.slides:
    if s._element.get("show") == "0":  # hidden slides are not exported
        continue
    notes.append(s.notes_slide.notes_text_frame.text if s.has_notes_slide else "")
notes = (notes + [""] * len(imgs))[:len(imgs)]
title = prs.core_properties.title or os.path.splitext(os.path.basename(src))[0]
json.dump({"title": title, "count": len(imgs), "w": prs.slide_width, "h": prs.slide_height, "notes": notes},
          open("site/slides/manifest.json", "w", encoding="utf-8"), ensure_ascii=False)
shutil.copy("src/index.html", "site/index.html")
print("OK:", len(imgs), "slides")
