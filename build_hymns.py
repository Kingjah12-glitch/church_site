"""Reads every PowerPoint in songs/<LETTER>/ and builds hymns.json for the website."""
import json, re, shutil, subprocess, tempfile
from pathlib import Path
from pptx import Presentation

ROOT = Path(__file__).parent
SONGS = ROOT / "songs"
OUT = ROOT / "hymns.json"


def find_soffice():
    """LibreOffice is only needed for old .ppt files."""
    for name in ("soffice", "libreoffice"):
        p = shutil.which(name)
        if p:
            return p
    for p in (r"C:\Program Files\LibreOffice\program\soffice.exe",
              r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"):
        if Path(p).exists():
            return p
    return None


def convert_all_ppt(files, soffice, tmp):
    """Convert .ppt to .pptx in batches, one output folder per source folder."""
    result = {}
    by_folder = {}
    for f in files:
        by_folder.setdefault(f.parent, []).append(f)

    total = len(by_folder)
    for i, (folder, group) in enumerate(by_folder.items(), start=1):
        print(f"  Converting folder {i}/{total}: {folder.name} ({len(group)} files)")
        out = tmp / f"c{i}"
        out.mkdir()
        for j in range(0, len(group), 30):
            chunk = group[j:j + 30]
            subprocess.run([soffice, "--headless", "--convert-to", "pptx",
                            "--outdir", str(out), *map(str, chunk)],
                           capture_output=True)
        for f in group:
            converted = out / (f.stem + ".pptx")
            if converted.exists():
                result[f] = converted
    return result


def shape_lines(shape):
    lines = []
    if shape.shape_type == 6:  # grouped shapes
        for s in shape.shapes:
            lines += shape_lines(s)
    elif shape.has_text_frame:
        for para in shape.text_frame.paragraphs:
            t = "".join(r.text for r in para.runs).strip()
            if t:
                lines.append(t)
    elif getattr(shape, "has_table", False) and shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                t = cell.text.strip()
                if t:
                    lines.append(t)
    return lines


def read_slides(pptx_path):
    prs = Presentation(str(pptx_path))
    slides = []
    for slide in prs.slides:
        lines = []
        for shape in slide.shapes:
            lines += shape_lines(shape)
        if lines:
            slides.append("\n".join(lines))
    return slides


def slugify(text):
    return re.sub(r"[^a-z0-9]+", "-", text.lower()).strip("-") or "song"


def main():
    soffice = find_soffice()
    hymns, problems, used_ids = [], [], set()

    files = sorted(p for p in SONGS.rglob("*")
                   if p.suffix.lower() in (".ppt", ".pptx") and not p.name.startswith("~$"))
    ppts = [f for f in files if f.suffix.lower() == ".ppt"]
    print(f"Found {len(files)} song files ({len(ppts)} old .ppt).")

    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)

        converted = {}
        skipped_ppt = 0
        if ppts:
            if not soffice:
                skipped_ppt = len(ppts)
                print(f"\n!! LibreOffice not found, so {skipped_ppt} old .ppt files will be skipped.")
                print("   Install LibreOffice from libreoffice.org, reopen VS Code, and run again.\n")
            else:
                print("Converting old .ppt files. This can take a few minutes...")
                converted = convert_all_ppt(ppts, soffice, tmp)

        for f in files:
            try:
                src = f
                if f.suffix.lower() == ".ppt":
                    if not soffice:
                        continue
                    src = converted.get(f)
                    if not src:
                        problems.append(f"{f.name}: conversion failed")
                        continue

                slides = read_slides(src)
                if not slides:
                    problems.append(f"{f.name}: no text found (maybe the words are images)")
                    continue

                title = f.stem.replace("_", " ").strip()
                folder = f.parent.name.strip()
                letter = folder.upper() if len(folder) == 1 else title[:1].upper()

                sid, n = slugify(title), 2
                while sid in used_ids:
                    sid = f"{slugify(title)}-{n}"
                    n += 1
                used_ids.add(sid)

                hymns.append({"id": sid, "title": title, "letter": letter, "slides": slides})
            except Exception as e:
                problems.append(f"{f.name}: {e}")

    hymns.sort(key=lambda h: h["title"].lower())
    OUT.write_text(json.dumps(hymns, ensure_ascii=False), encoding="utf-8")

    print(f"\nDone: {len(hymns)} songs written to {OUT.name}")
    if skipped_ppt:
        print(f"Skipped {skipped_ppt} old .ppt files (LibreOffice missing).")
    if problems:
        print(f"\n{len(problems)} problem(s):")
        for p in problems[:40]:
            print("  -", p)
        if len(problems) > 40:
            print(f"  ... and {len(problems) - 40} more")


main()