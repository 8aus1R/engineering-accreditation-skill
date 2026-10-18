"""Exercise public helpers using synthetic scores and isolated temporary files."""

import copy
import csv
import hashlib
import importlib.util
import json
import shutil
import subprocess
import sys
import unittest
import uuid
import zipfile
from contextlib import contextmanager
from decimal import Decimal
from pathlib import Path


REPO = Path(__file__).resolve().parents[1]
SCRIPTS = REPO / "skills" / "engineering-accreditation" / "scripts"
EXAMPLE = REPO / "examples" / "attainment"
sys.dont_write_bytecode = True


def load_helper(filename):
    spec = importlib.util.spec_from_file_location(filename.removesuffix(".py"), SCRIPTS / filename)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


calculator = load_helper("check_attainment.py")
delivery = load_helper("verify_homework_delivery.py")


@contextmanager
def temporary_workspace():
    # Keep disposable data inside the repository's ignored workspace, including
    # when a sandbox disallows the operating system's default temp directory.
    root = (REPO / "work" / "test-temp").resolve()
    if not root.is_relative_to(REPO):
        raise ValueError("Test workspace escapes repository")
    root.mkdir(parents=True, exist_ok=True)
    # Inherit the workspace ACL. Python 3.13's restrictive tempfile ACL may
    # exclude the sandbox identity on Windows even inside a writable folder.
    directory = root / ("check-" + uuid.uuid4().hex)
    directory.mkdir()
    try:
        yield str(directory)
    finally:
        resolved = directory.resolve()
        if resolved.parent != root or not resolved.is_relative_to(REPO):
            raise ValueError("Refusing cleanup outside the test workspace")
        shutil.rmtree(resolved)


class AttainmentChecks(unittest.TestCase):
    def setUp(self):
        self.rules = json.loads((EXAMPLE / "rules.json").read_text(encoding="utf-8"))
        with (EXAMPLE / "scores.csv").open(encoding="utf-8", newline="") as handle:
            self.rows = list(csv.DictReader(handle))

    def test_hand_calculated_results_and_real_zero(self):
        result = calculator.calculate(self.rows, self.rules)
        self.assertEqual(result["student_count"], 3)
        self.assertEqual(result["students"][0]["attainment"]["goal1"]["rounded"], "0.77")
        self.assertEqual(result["students"][1]["attainment"]["goal1"]["rounded"], "0.84")
        self.assertEqual(result["students"][2]["attainment"]["goal1"]["rounded"], "0.00")
        self.assertLess(abs(Decimal(result["aggregate"]["goal1"]["exact"]) - Decimal(161) / Decimal(300)), Decimal("1e-26"))
        self.assertEqual(result["aggregate"]["goal1"]["rounded"], "0.54")
        self.assertEqual(result["aggregate"]["goal2"]["rounded"], "0.50")

    def test_missing_score_rejected(self):
        self.rows[0]["coursework"] = ""
        with self.assertRaisesRegex(ValueError, "missing"):
            calculator.calculate(self.rows, self.rules)

    def test_duplicate_identity_rejected(self):
        self.rows[1]["student_id"] = self.rows[0]["student_id"]
        with self.assertRaisesRegex(ValueError, "duplicate"):
            calculator.calculate(self.rows, self.rules)

    def test_score_out_of_range_rejected(self):
        self.rows[0]["exam_part_a"] = "41"
        with self.assertRaisesRegex(ValueError, "maximum|outside"):
            calculator.calculate(self.rows, self.rules)

    def test_wrong_expected_count_rejected(self):
        self.rules["expected_count"] = 4
        with self.assertRaisesRegex(ValueError, "count"):
            calculator.calculate(self.rows, self.rules)

    def test_changed_observed_value_rejected(self):
        expected = calculator.calculate(self.rows, self.rules)
        changed = copy.deepcopy(expected)
        changed["aggregate"]["goal1"]["exact"] = "0.9"
        with self.assertRaisesRegex(ValueError, "exact attainment differs"):
            calculator.verify(expected, changed)

    def test_cli_demo_and_input_hashes(self):
        paths = [EXAMPLE / name for name in ("scores.csv", "rules.json", "expected-results.json")]
        before = {p: p.read_bytes() for p in paths}
        with temporary_workspace() as directory:
            output = Path(directory) / "result.json"
            run = subprocess.run([
                sys.executable, "-X", "utf8", str(SCRIPTS / "check_attainment.py"),
                "--csv", str(paths[0]), "--config", str(paths[1]),
                "--observed", str(paths[2]), "--output", str(output),
            ], capture_output=True, text=True, encoding="utf-8")
            self.assertEqual(run.returncode, 0, run.stderr)
            result = json.loads(output.read_text(encoding="utf-8"))
            self.assertEqual(result["verification"], "excel_numeric_results_matched")
            self.assertEqual(result["input_sha256"], {str(p): hashlib.sha256(before[p]).hexdigest() for p in paths})
        for path in paths:
            self.assertEqual(path.read_bytes(), before[path])

    def test_cli_does_not_overwrite_input(self):
        with temporary_workspace() as directory:
            source = Path(directory) / "input.csv"
            source.write_bytes((EXAMPLE / "scores.csv").read_bytes())
            before = source.read_bytes()
            run = subprocess.run([
                sys.executable, "-X", "utf8", str(SCRIPTS / "check_attainment.py"),
                "--csv", str(source), "--config", str(EXAMPLE / "rules.json"),
                "--output", str(source),
            ], capture_output=True, text=True, encoding="utf-8")
            self.assertNotEqual(run.returncode, 0)
            self.assertIn("output must differ", run.stderr)
            self.assertEqual(source.read_bytes(), before)


class DeliveryChecks(unittest.TestCase):
    def setUp(self):
        self.temp = temporary_workspace()
        self.root = Path(self.temp.__enter__())
        self.addCleanup(self.temp.__exit__, None, None, None)
        folder = self.root / "assignment-01"
        folder.mkdir()
        self.document = folder / "DEMO-001.docx"
        # Small synthetic OOXML fixture. Tests concern inventory, not Word rendering.
        with zipfile.ZipFile(self.document, "w") as package:
            package.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
            package.writestr("_rels/.rels", '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
            package.writestr("word/document.xml", '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Synthetic fixture</w:t></w:r></w:p></w:body></w:document>')
        self.inventory = {
            "documents_inventory": [{"file": "assignment-01/DEMO-001.docx", "sha256": hashlib.sha256(self.document.read_bytes()).hexdigest()}],
            "documents": 1,
            "assignment_folders": [{"folder": "assignment-01", "documents": 1}],
        }
        self.save_inventory()
        (self.root / "scores.csv").write_text("file,score\nassignment-01/DEMO-001.docx,77\n", encoding="utf-8")

    def save_inventory(self):
        (self.root / "inventory.json").write_text(json.dumps(self.inventory), encoding="utf-8")

    def test_complete_delivery_passes_with_explicit_limits(self):
        result = delivery.verify(self.root, "inventory.json", ["scores.csv"], "file")
        self.assertEqual(result["status"], "PASS")
        self.assertEqual(result["documents"], 1)
        self.assertFalse(result["grades_and_rendering_checked_by_this_script"])

    def test_changed_file_rejected(self):
        self.document.write_bytes(self.document.read_bytes() + b"changed")
        with self.assertRaisesRegex(ValueError, "SHA-256 mismatch"):
            delivery.verify(self.root, "inventory.json")

    def test_extra_document_rejected(self):
        (self.root / "extra.docx").write_bytes(self.document.read_bytes())
        with self.assertRaisesRegex(ValueError, "Missing or extra"):
            delivery.verify(self.root, "inventory.json")

    def test_path_escape_rejected(self):
        self.inventory["documents_inventory"][0]["file"] = "../outside.docx"
        self.save_inventory()
        with self.assertRaisesRegex(ValueError, "escapes delivery root"):
            delivery.verify(self.root, "inventory.json")

    def test_incomplete_csv_rejected(self):
        (self.root / "scores.csv").write_text("file,score\n", encoding="utf-8")
        with self.assertRaisesRegex(ValueError, "does not cover"):
            delivery.verify(self.root, "inventory.json", ["scores.csv"], "file")

    def test_docx_does_not_satisfy_binary_doc_requirement(self):
        with self.assertRaisesRegex(ValueError, "Not binary Word DOC"):
            delivery.verify(self.root, "inventory.json", require_binary_doc=True)


if __name__ == "__main__":
    unittest.main()
