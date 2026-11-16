"""presentation/*.pptx -> site/slides/N.jpg + manifest.json + index.html
یادداشت‌های ویرایش‌شده‌ی سایت (presentation/notes.json) داخل خود فایل pptx هم نوشته می‌شوند."""
import glob, hashlib, json, os, shutil, subprocess, time
from pptx import Presentation

TITLE = "PptxSARA Presentation"   # نام سایت
PASSWORD = "1611"                  # رمز ورود به سایت (برای تغییر، فقط همین‌جا را عوض کنید)

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
vis = [s for s in prs.slides if s._element.get("show") != "0"]  # hidden slides are not exported
notes = [s.notes_slide.notes_text_frame.text if s.has_notes_slide else "" for s in vis]
notes = (notes + [""] * len(imgs))[:len(imgs)]

# notes edited from the website (key = visible slide number, starting at 1)
edits, changed = "presentation/notes.json", False
if os.path.exists(edits):
    try:
        for k, v in json.load(open(edits, encoding="utf-8")).items():
            if k.isdigit() and 1 <= int(k) <= len(notes) and notes[int(k) - 1] != v:
                notes[int(k) - 1] = v; changed = True
    except Exception as e:
        print("WARN: notes.json ignored:", e)

# write the edited notes back into the original pptx (committed by the workflow)
if changed and len(vis) == len(imgs):
    for s, t in zip(vis, notes):
        if (s.notes_slide.notes_text_frame.text if s.has_notes_slide else "") != t:
            s.notes_slide.notes_text_frame.text = t
    prs.save(src)
    print("notes folded into", src)

json.dump({"title": TITLE, "count": len(imgs), "w": prs.slide_width, "h": prs.slide_height, "notes": notes,
           "repo": os.environ.get("GITHUB_REPOSITORY", ""), "branch": os.environ.get("GITHUB_REF_NAME", "main"),
           "v": int(time.time())},
          open("site/slides/manifest.json", "w", encoding="utf-8"), ensure_ascii=False)
ph = hashlib.sha256(("sara|" + PASSWORD).encode()).hexdigest()
html = open("src/index.html", encoding="utf-8").read().replace("__PH__", ph)
open("site/index.html", "w", encoding="utf-8").write(html)
print("OK:", len(imgs), "slides")
