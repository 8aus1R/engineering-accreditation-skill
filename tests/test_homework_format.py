"""Behavior tests for the format guard, not a Word renderer or grading test."""
import base64
import json
import os
import shutil
import uuid
import unittest
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET

import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "skills" / "engineering-accreditation" / "scripts"))
from check_homework_format import FLAGS, W, run_gate, sha


class HomeworkFormatTests(unittest.TestCase):
    def setUp(self):
        self.base = Path(os.environ.get("HOMEWORK_FORMAT_TEST_ROOT", str(Path.cwd() / "tmp/homework_format_tests"))).resolve()
        self.base.mkdir(parents=True, exist_ok=True)
        self.root = self.base / ("case_" + uuid.uuid4().hex)
        self.root.mkdir()
        self.docs = self.root / "docs"
        self.docs.mkdir()
        self.file = self.docs / "student.docx"
        self.make_doc()

    def tearDown(self):
        target = self.root.resolve()
        assert target.is_relative_to(self.base) and target.name.startswith("case_")
        shutil.rmtree(target)

    def make_doc(self, prefix="", comment=None, bad_flag=None, inherited=False,
                 style_true=False, numbered=False, normal_question=None):
        document = ET.Element(W + "document")
        body = ET.SubElement(document, W + "body")

        def paragraph():
            p = ET.SubElement(body, W + "p")
            prop = ET.SubElement(p, W + "pPr")
            for flag in FLAGS:
                if inherited and flag == "keepNext":
                    continue
                ET.SubElement(prop, W + flag).set(W + "val", "1" if flag == bad_flag else "0")
            return p

        p = paragraph()
        if numbered:
            prop = p.find(W + "pPr")
            ET.SubElement(ET.SubElement(prop, W + "numPr"), W + "numId").set(W + "val", "1")
        run = ET.SubElement(p, W + "r")
        ET.SubElement(run, W + "t").text = normal_question or "1. 简述原理。"
        if prefix:  # Separate run makes the regression cross the run boundary.
            ET.SubElement(ET.SubElement(p, W + "r"), W + "t").text = prefix
        run = ET.SubElement(p, W + "r")
        ET.SubElement(ET.SubElement(run, W + "rPr"), W + "color").set(W + "val", "C00000")
        ET.SubElement(run, W + "t").text = "+4"
        if comment:
            ET.SubElement(ET.SubElement(paragraph(), W + "r"), W + "t").text = comment
        styles = ET.Element(W + "styles")
        if style_true:
            prop = ET.SubElement(ET.SubElement(styles, W + "style"), W + "pPr")
            ET.SubElement(prop, W + "keepNext")  # Omitted val is true, not false.
        with ZipFile(self.file, "w") as z:
            z.writestr("word/document.xml", ET.tostring(document))
            z.writestr("word/styles.xml", ET.tostring(styles))

    def gate(self, **kw):
        return run_gate([self.docs], kw.pop("expected_count", 1), 1,
                        score_color="C00000", structure_only=True, **kw)

    def codes(self, report):
        return {e["code"] for r in report["records"] for e in r["errors"]} | {e["code"] for e in report["errors"]}

    def test_clean_and_read_only(self):
        before = sha(self.file)
        result = self.gate()
        self.assertEqual(result["status"], "STRUCTURE_ONLY_PASS")
        self.assertEqual(before, sha(self.file))

    def test_space_in_previous_run_and_literal_box_are_caught(self):
        for prefix in ["\u3000", "□"]:
            self.make_doc(prefix=prefix)
            self.assertIn("SCORE_PREFIX_BOX_OR_IDEOGRAPHIC_SPACE", self.codes(self.gate()))

    def test_body_comments_are_caught(self):
        for comment in ["未完成评分点8、9、10，扣3分。", "最终计算有误，扣1分。", "未作答，得0分。"]:
            self.make_doc(comment=comment)
            self.assertIn("BODY_DEDUCTION_COMMENT", self.codes(self.gate()))

    def test_question_content_is_not_deleted_or_misclassified(self):
        self.make_doc(normal_question="1. 某测验答错扣1分，求最终成绩。")
        self.assertNotIn("BODY_DEDUCTION_COMMENT", self.codes(self.gate()))

    def test_direct_true_and_missing_override_fail(self):
        for flag in FLAGS:
            self.make_doc(bad_flag=flag)
            self.assertIn("PARAGRAPH_CONTROL_NOT_EXPLICIT_OFF", self.codes(self.gate()))
        self.make_doc(inherited=True, style_true=True)
        self.assertIn("PARAGRAPH_CONTROL_NOT_EXPLICIT_OFF", self.codes(self.gate()))
        self.assertIn("STYLE_CONTROL_ENABLED", self.codes(self.gate()))

    def test_real_numbering_fails(self):
        self.make_doc(numbered=True)
        self.assertIn("REAL_LIST_NUMBERING", self.codes(self.gate()))

    def test_coverage_mismatch_and_unsupported_doc_fail(self):
        self.assertIn("DOCUMENT_COUNT", self.codes(self.gate(expected_count=2)))
        (self.docs / "legacy.doc").write_bytes(b"not a DOCX")
        self.assertIn("UNSUPPORTED_WORD_FORMAT", self.codes(self.gate()))

    def test_all_documents_are_checked_after_first_failure(self):
        self.make_doc(comment="未完成评分点1，扣1分。")
        (self.docs / "second.docx").write_bytes(self.file.read_bytes())
        result = self.gate(expected_count=2)
        self.assertEqual(result["checked_documents"], 2)
        self.assertEqual(result["failed_documents"], 2)

    def render_evidence(self):
        image = self.root / "page-1.png"
        image.write_bytes(base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl6uN8AAAAASUVORK5CYII="))
        return {"records": [{"file": str(self.file), "docx_sha256": sha(self.file),
                             "pages": 1, "layout_checked": True, "unexplained_differences": 0,
                             "images": [{"file": str(image), "sha256": sha(image)}]}]}

    def final_gate(self, evidence):
        path = self.root / "render.json"
        path.write_text(json.dumps(evidence), encoding="utf-8")
        return run_gate([self.docs], 1, 1, "C00000", path)

    def test_final_gate_needs_current_render_and_all_pages(self):
        evidence = self.render_evidence()
        self.assertEqual(self.final_gate(evidence)["status"], "FORMAT_AND_COVERAGE_PASS")
        evidence["records"][0]["docx_sha256"] = "stale"
        self.assertIn("STALE_RENDER_HASH", self.codes(self.final_gate(evidence)))
        evidence = self.render_evidence()
        evidence["records"][0]["pages"] = 2
        self.assertIn("RENDER_PAGE_COVERAGE", self.codes(self.final_gate(evidence)))
        self.assertIn("MISSING_RENDER_EVIDENCE", self.codes(self.final_gate({"records": []})))

    def test_unreviewed_layout_and_stale_image_fail(self):
        evidence = self.render_evidence()
        evidence["records"][0]["layout_checked"] = False
        evidence["records"][0]["images"][0]["sha256"] = "stale"
        codes = self.codes(self.final_gate(evidence))
        self.assertIn("LAYOUT_REVIEW_INCOMPLETE", codes)
        self.assertIn("INVALID_PAGE_IMAGE", codes)


if __name__ == "__main__":
    unittest.main()
