"""Read-only DOCX format/coverage gate; does not grade answers or inspect pixels."""
import argparse
import hashlib
import json
import re
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = "{" + NS["w"] + "}"
FLAGS = ("keepNext", "keepLines", "pageBreakBefore", "suppressLineNumbers")
OFF = {"0", "false", "off"}
SCORE = re.compile(r"[+＋]\d+(?:\.\d+)?(?:分)?$")
BOX_PREFIX = re.compile(r"[\s□■▪▫☐]*$")
COMMENT = re.compile(r"未完成评分点|扣\s*\d+(?:\.\d+)?\s*分|未作答[，,：:\s]*(?:得)?\s*0\s*分|^(?:扣分说明|扣分原因|批注)[：:]")


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def text(element):
    return "".join(t.text or "" for t in element.iter(W + "t"))


def scan_docx(path, expected_scores, score_color=None):
    """Return every detected format issue; never rewrite the package."""
    result = {"file": str(Path(path).resolve()), "sha256": sha(path),
              "score_fields": 0, "paragraphs": 0, "errors": []}
    errors = result["errors"]
    try:
        with ZipFile(path) as archive:
            broken = archive.testzip()
            if broken:
                raise ValueError("Broken ZIP member: " + broken)
            names = archive.namelist()
            if "word/document.xml" not in names:
                raise ValueError("Missing word/document.xml")
            for name in names:
                paragraph_part = (name == "word/document.xml" or
                                  name.startswith(("word/header", "word/footer")))
                style_part = name == "word/styles.xml"
                if not name.endswith(".xml") or not (paragraph_part or style_part):
                    continue
                root = ET.fromstring(archive.read(name))
                if style_part:
                    for flag in FLAGS:
                        for node in root.iter(W + flag):
                            if node.get(W + "val", "1") not in OFF:
                                errors.append({"code": "STYLE_CONTROL_ENABLED", "part": name, "flag": flag})
                    continue
                for index, paragraph in enumerate(root.iter(W + "p"), 1):
                    result["paragraphs"] += 1
                    loc = {"part": name, "paragraph": index}
                    prop = paragraph.find(W + "pPr")
                    for flag in FLAGS:
                        node = None if prop is None else prop.find(W + flag)
                        if node is None or node.get(W + "val", "1") not in OFF:
                            errors.append(dict(loc, code="PARAGRAPH_CONTROL_NOT_EXPLICIT_OFF", flag=flag))
                    numbering = None if prop is None else prop.find(W + "numPr")
                    if numbering is not None:
                        num_id = numbering.find(W + "numId")
                        if num_id is None or num_id.get(W + "val") != "0":
                            errors.append(dict(loc, code="REAL_LIST_NUMBERING"))
                    if any(paragraph.find(".//w:" + n, NS) is not None
                           for n in ("commentReference", "commentRangeStart", "commentRangeEnd")):
                        errors.append(dict(loc, code="WORD_COMMENT_MARKER"))
                    value = text(paragraph)
                    content_role = re.match(r"^\s*(?:\d+[.、．]|答[：:]|作答[：:])", value)
                    if COMMENT.search(value) and not content_role:
                        errors.append(dict(loc, code="BODY_DEDUCTION_COMMENT", text=value))
                    if name != "word/document.xml":
                        continue
                    match = SCORE.search(value.rstrip())
                    if not match:
                        continue
                    # Score color is a task parameter, not a global red preference.
                    pos = 0
                    owner = None
                    for run in paragraph.iter(W + "r"):
                        run_text = text(run)
                        if pos <= match.start() < pos + len(run_text):
                            owner = run
                            break
                        pos += len(run_text)
                    if owner is None:
                        errors.append(dict(loc, code="UNRESOLVED_SCORE_FIELD"))
                        continue
                    color = owner.find("./w:rPr/w:color", NS)
                    color_value = None if color is None else color.get(W + "val")
                    if score_color and (color_value or "").upper() != score_color.upper():
                        continue
                    if not score_color:
                        # Without an explicit color, only isolated score runs count;
                        # avoid treating the '+1' at the end of an answer as a mark.
                        if not re.fullmatch(r"[\s□■▪▫☐]*[+＋]\d+(?:\.\d+)?(?:分)?\s*", text(owner)):
                            continue
                    result["score_fields"] += 1
                    prefix = BOX_PREFIX.search(value[:match.start()]).group()
                    if "\u3000" in prefix or any(c in prefix for c in "□■▪▫☐"):
                        errors.append(dict(loc, code="SCORE_PREFIX_BOX_OR_IDEOGRAPHIC_SPACE", prefix=prefix))
            if result["score_fields"] != expected_scores:
                errors.append({"code": "SCORE_FIELD_COUNT", "expected": expected_scores,
                               "actual": result["score_fields"]})
    except Exception as exc:
        errors.append({"code": "UNREADABLE_DOCX", "reason": str(exc)})
    return result


def run_gate(roots, expected_count, expected_scores, score_color=None,
             render_report=None, structure_only=False):
    files = set()
    errors = []
    for root_value in roots:
        root = Path(root_value).resolve()
        if not root.is_dir():
            errors.append({"code": "MISSING_ROOT", "root": str(root)})
            continue
        for path in root.rglob("*"):
            if not path.is_file() or path.name.startswith("~$"):
                continue
            if path.suffix.lower() == ".docx":
                files.add(path.resolve())
            elif path.suffix.lower() in {".doc", ".docm", ".dotx", ".dotm"}:
                errors.append({"code": "UNSUPPORTED_WORD_FORMAT", "file": str(path)})
    files = sorted(files, key=str)
    if len(files) != expected_count:
        errors.append({"code": "DOCUMENT_COUNT", "expected": expected_count, "actual": len(files)})
    records = [scan_docx(path, expected_scores, score_color) for path in files]
    evidence = {}
    if not structure_only:
        try:
            raw = json.loads(Path(render_report).read_text(encoding="utf-8-sig"))
            for item in raw["records"]:
                key = str(Path(item["file"]).resolve())
                if key in evidence:
                    errors.append({"code": "DUPLICATE_RENDER_EVIDENCE", "file": key})
                evidence[key] = item
        except Exception as exc:
            errors.append({"code": "MISSING_OR_INVALID_RENDER_REPORT", "reason": str(exc)})
        for record in records:
            item = evidence.get(record["file"])
            if not item:
                record["errors"].append({"code": "MISSING_RENDER_EVIDENCE"})
                continue
            if item.get("docx_sha256") != record["sha256"]:
                record["errors"].append({"code": "STALE_RENDER_HASH"})
            pages, images = item.get("pages"), item.get("images", [])
            if type(pages) is not int or pages < 1 or len(images) != pages:
                record["errors"].append({"code": "RENDER_PAGE_COVERAGE"})
            if item.get("layout_checked") is not True or item.get("unexplained_differences") != 0:
                record["errors"].append({"code": "LAYOUT_REVIEW_INCOMPLETE"})
            seen = set()
            for img in images:
                try:
                    image_path = Path(img["file"]).resolve()
                    if image_path in seen:
                        raise ValueError("Duplicate page image")
                    seen.add(image_path)
                    if not image_path.is_file() or sha(image_path) != img["sha256"]:
                        raise ValueError("Missing image or stale image hash")
                except Exception as exc:
                    record["errors"].append({"code": "INVALID_PAGE_IMAGE", "reason": str(exc)})
    failed = sum(bool(r["errors"]) for r in records)
    status = "FAIL" if errors or failed else ("STRUCTURE_ONLY_PASS" if structure_only else "FORMAT_AND_COVERAGE_PASS")
    return {"status": status, "expected_documents": expected_count, "checked_documents": len(records),
            "failed_documents": failed, "structure_only": structure_only,
            "limitations": "Does not grade answers, compare Excel, inspect image pixels, or prove native Word UI behavior.",
            "errors": errors, "records": records}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", action="append", required=True)
    parser.add_argument("--expected-count", type=int, required=True)
    parser.add_argument("--expected-score-count", type=int, required=True)
    parser.add_argument("--score-color")
    parser.add_argument("--render-report")
    parser.add_argument("--structure-only", action="store_true")
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    if args.expected_count < 1 or args.expected_score_count < 0:
        parser.error("Expected document count must be positive; score count cannot be negative")
    if not args.structure_only and not args.render_report:
        parser.error("Final gate requires --render-report; --structure-only is debugging only")
    report = run_gate(args.root, args.expected_count, args.expected_score_count,
                      args.score_color, args.render_report, args.structure_only)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({k: v for k, v in report.items() if k != "records"}, ensure_ascii=False))
    return 1 if report["status"] == "FAIL" else 0


if __name__ == "__main__":
    raise SystemExit(main())
