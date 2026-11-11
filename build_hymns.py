"""Reads every PowerPoint in songs/<LETTER>/ and builds a clean hymns.json."""
import json, re, shutil, subprocess, tempfile
from difflib import SequenceMatcher
from pathlib import Path
from pptx import Presentation

ROOT = Path(__file__).parent
SONGS = ROOT / "songs"
CACHE = ROOT / "_converted"
OUT = ROOT / "hymns.json"

TITLE_RE = re.compile(r"^\s*\d+\s*[–—-]\s*(\S.*)$")      # "1 – A Purple Robe"
NUM_RE = re.compile(r"^\s*\d+\.?\s*$")                   # "1." on its own line
CREDIT_RES = [re.compile(p, re.I) for p in (
    r"^\s*(words|music|text|lyrics|tune|arr\.?|arranged|translation|translated|harmony|meter|"
    r"written|composed|author|composer)(\s+(and|&)\s+(music|words|lyrics))?(\s+by)?\s*[:©]",
    r"©", r"\(c\)\s*\d{4}", r"^\s*\(?admin\.?", r"all rights reserved",
    r"used by permission", r"paperless hymnal", r"\bccli\b", r"public domain",
)]


def is_credit(line):
    return any(r.search(line) for r in CREDIT_RES)


def key(s):
    return re.sub(r"[^a-z0-9]", "", s.lower())


def base_title(stem):
    """File name without copy markers like (1), '- Copy' and the SFP tag."""
    s = re.sub(r"\s*\(\d+\)\s*$", "", stem)
    s = re.sub(r"\s*[-–]\s*copy(\s*\(\d+\))?\s*$", "", s, flags=re.I)
    s = re.sub(r"[-_ ]*SFP$", "", s, flags=re.I)
    return s.strip()


def find_soffice():
    for name in ("soffice", "libreoffice"):
        p = shutil.which(name)
        if p:
            return p
    for p in (r"C:\Program Files\LibreOffice\program\soffice.exe",
              r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"):
        if Path(p).exists():
            return p
    return None


def cached_path(f):
    return CACHE / f.parent.name / (f.stem + ".pptx")


def run_convert(soffice, files, out, timeout):
    return subprocess.run([soffice, "--headless", "--convert-to", "pptx",
                           "--outdir", str(out), *map(str, files)],
                          capture_output=True, timeout=timeout)


def convert_missing(ppts, soffice):
    by_folder = {}
    for f in ppts:
        if not cached_path(f).exists():
            by_folder.setdefault(f.parent, []).append(f)
    total = len(by_folder)
    for i, (folder, group) in enumerate(by_folder.items(), start=1):
        print(f"  Converting folder {i}/{total}: {folder.name} ({len(group)} files)")
        out = CACHE / folder.name
        out.mkdir(parents=True, exist_ok=True)
        for j in range(0, len(group), 30):
            with tempfile.TemporaryDirectory() as tmpdir:
                copies = []
                for f in group[j:j + 30]:
                    c = Path(tmpdir) / f.name
                    shutil.copyfile(f, c)
                    copies.append(c)
                try:
                    run_convert(soffice, copies, out, 900)
                except subprocess.TimeoutExpired:
                    for c in copies:
                        if not (out / (c.stem + ".pptx")).exists():
                            try:
                                run_convert(soffice, [c], out, 120)
                            except subprocess.TimeoutExpired:
                                print("    Skipped (took too long):", c.name)


def shape_lines(shape):
    lines = []
    if shape.shape_type == 6:
        for s in shape.shapes:
            lines += shape_lines(s)
    elif shape.has_text_frame:
        for para in shape.text_frame.paragraphs:
            for t in para.text.replace("\x0b", "\n").split("\n"):
                t = t.strip()
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
    out = []
    for slide in prs.slides:
        lines = []
        for shape in slide.shapes:
            lines += shape_lines(shape)
        out.append(lines)
    return out


def clean_slides(raw_slides, title_key, unclear):
    """Remove the title line and the credits from every slide."""
    starts = sum(1 for lines in raw_slides if lines and key(lines[0]) == title_key)
    strip_title = starts >= 2

    slides, slide_title = [], None
    for lines in raw_slides:
        keep, in_footer = [], False
        for idx, line in enumerate(lines):
            if (idx == 0 and key(line) == title_key and
                    (strip_title or (len(lines) > 1 and NUM_RE.match(lines[1])))):
                continue
            m = TITLE_RE.match(line)
            if m:
                if slide_title is None:
                    slide_title = m.group(1).strip()
                continue
            if is_credit(line):
                if keep:
                    in_footer = True
                continue
            if in_footer:
                unclear.append(line)
                continue
            keep.append(line)
        if keep:
            slides.append("\n".join(keep))
    return slides, slide_title


def slugify(text):
    return re.sub(r"[^a-z0-9]+", "-", text.lower()).strip("-") or "song"


def dedupe(cands):
    """Keep one copy of each song. Best version first: most words, non-SFP."""
    kept, merged = [], []
    by_text, by_base = {}, {}
    for c in sorted(cands, key=lambda c: (-len(c["text"]), c["sfp"], c["title"].lower())):
        dup = by_text.get(c["text"])
        if not dup:
            for k in by_base.get(key(c["base"]), []):
                if SequenceMatcher(None, k["text"], c["text"]).ratio() >= 0.6:
                    dup = k
                    break
        if dup:
            merged.append(f'{c["src"]}  ->  kept {dup["src"]}')
        else:
            kept.append(c)
            by_text[c["text"]] = c
            by_base.setdefault(key(c["base"]), []).append(c)
    return kept, merged


def main():
    soffice = find_soffice()
    files = sorted(p for p in SONGS.rglob("*")
                   if p.suffix.lower() in (".ppt", ".pptx") and not p.name.startswith("~$"))
    ppts = [f for f in files if f.suffix.lower() == ".ppt"]
    print(f"Found {len(files)} song files ({len(ppts)} old .ppt).")

    if ppts:
        if not soffice:
            print("\n!! LibreOffice not found. Old .ppt files without a saved copy will be skipped.\n")
        else:
            missing = [f for f in ppts if not cached_path(f).exists()]
            if missing:
                print(f"Converting {len(missing)} .ppt files (only needed once)...")
                convert_missing(missing, soffice)

    cands, problems, unclear, empty_names = [], [], [], []

    for f in files:
        try:
            src = f
            if f.suffix.lower() == ".ppt":
                src = cached_path(f)
                if not src.exists():
                    problems.append(f"{f.parent.name}/{f.name}: could not be converted")
                    continue

            stem = f.stem.replace("_", " ").strip()
            base = base_title(stem)
            sfp = base != stem and bool(re.search(r"SFP$", stem, re.I))

            local_unclear = []
            slides, slide_title = clean_slides(read_slides(src), key(base), local_unclear)
            if not slides:
                empty_names.append(f"{f.parent.name}/{f.name}")
                continue

            title = base
            if slide_title and key(slide_title) == key(base):
                title = slide_title
            if sfp:
                title += " (SFP)"

            if local_unclear:
                unclear.append(f"{title}: {local_unclear[0]}")

            folder = f.parent.name.strip()
            cands.append({
                "base": base, "title": title, "sfp": sfp,
                "letter": folder.upper() if len(folder) == 1 else title[:1].upper(),
                "slides": slides, "text": key("".join(slides)),
                "src": f"{f.parent.name}/{f.name}",
            })
        except Exception as e:
            problems.append(f"{f.parent.name}/{f.name}: {e}")

    cands, merged = dedupe(cands)

    hymns, used_ids = [], set()
    for c in sorted(cands, key=lambda c: c["title"].lower()):
        sid, n = slugify(c["title"]), 2
        while sid in used_ids:
            sid = f"{slugify(c['title'])}-{n}"
            n += 1
        used_ids.add(sid)
        hymns.append({"id": sid, "title": c["title"], "letter": c["letter"], "slides": c["slides"]})

    OUT.write_text(json.dumps(hymns, ensure_ascii=False), encoding="utf-8")
    size = OUT.stat().st_size / 1_000_000
    print(f"\nDone: {len(hymns)} songs written to {OUT.name} ({size:.1f} MB)")
    print(f"Removed: {len(empty_names)} empty entries, {len(merged)} duplicates")

    if merged:
        print(f"\nDuplicates merged (first one removed, second one kept):")
        for m in merged[:40]:
            print("  -", m)
        if len(merged) > 40:
            print(f"  ... and {len(merged) - 40} more")
    if empty_names:
        print(f"\n{len(empty_names)} file(s) with no lyrics found:")
        for n in empty_names[:25]:
            print("  -", n)
    if problems:
        print(f"\n{len(problems)} problem(s):")
        for p in problems[:25]:
            print("  -", p)
    if unclear:
        print(f"\n{len(unclear)} song(s) where text after the footer was dropped. Please check a few:")
        for u in unclear[:15]:
            print("  -", u)


main()