"""Build the Law Study Desk page from the Markdown notes in law-course/notes.

Usage: python3 law-course/build_study_desk.py
Writes law-course/study-desk.html, which is published as the study artifact.
Every notes/*.md file becomes a class in the page, so a new class only needs
its Markdown file and a rebuild.
"""
import json
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent


def group_and_order(stem):
    if stem.startswith("00_"):
        return "Course map", (0, stem)
    if stem.startswith("F"):
        return "Phase 0: Foundations", (1, stem)
    return "Phase 1: Civil practice", (2, stem)


def short_name(md, stem):
    title = next((l[2:].strip() for l in md.splitlines() if l.startswith("# ")), stem)
    return re.sub(r"\*", "", title)


def main():
    classes = []
    for f in (ROOT / "notes").glob("*.md"):
        md = f.read_text(encoding="utf-8")
        group, order = group_and_order(f.stem)
        cid = re.sub(r"[^A-Za-z0-9]+", "-", f.stem.split("_")[0] + "-" + f.stem.split("_")[1]).lower()
        classes.append({"id": cid, "short": short_name(md, f.stem), "group": group,
                        "md": md, "_order": order})
    classes.sort(key=lambda c: c.pop("_order"))
    # Put the course map last in the list so the first class opens by default.
    classes = [c for c in classes if c["group"] != "Course map"] + \
              [c for c in classes if c["group"] == "Course map"]
    exams = {}
    for f in sorted((ROOT / "exams").glob("*.json")):
        exam = json.loads(f.read_text(encoding="utf-8"))
        for q in exam["questions"]:
            if q["type"] == "written":
                points = sum(p["m"] for p in q["points"])
                assert abs(points - q["marks"]) < 1e-9, f"{f.name} {q['id']}: points add to {points}, not {q['marks']}"
        exams[exam["classId"]] = exam
    data = json.dumps(classes, ensure_ascii=False).replace("</", "<\\/")
    exam_data = json.dumps(exams, ensure_ascii=False).replace("</", "<\\/")
    template = (ROOT / "study_desk_template.html").read_text(encoding="utf-8")
    out = ROOT / "study-desk.html"
    out.write_text(template.replace("/*__CLASSES__*/", data).replace("/*__EXAMS__*/", exam_data),
                   encoding="utf-8")
    print("built", out.relative_to(ROOT.parent), "with", len(classes), "classes and", len(exams), "exams")


if __name__ == "__main__":
    main()
