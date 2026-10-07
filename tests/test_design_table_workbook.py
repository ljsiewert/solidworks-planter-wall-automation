"""Opt-in Excel integration checks using the user-provided workbook, read-only."""

import hashlib
import os
from pathlib import Path
import shutil
import subprocess
import unittest


ROOT = Path(__file__).resolve().parents[1]
WORKBOOK = ROOT / "PLANTER_ASSEMBLY.xlsx"


@unittest.skipUnless(
    os.environ.get("PLANTER_TEST_EXCEL") == "1"
    and WORKBOOK.exists() and shutil.which("powershell"),
    "Set PLANTER_TEST_EXCEL=1 with Excel installed and the reference workbook present",
)
class DesignTableWorkbookTests(unittest.TestCase):
    def test_real_excel_formulas_and_restoration(self):
        digest = hashlib.sha256(WORKBOOK.read_bytes()).digest()
        result = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File",
             str(ROOT / "tests" / "verify_design_table.ps1"),
             "-WorkbookPath", str(WORKBOOK)],
            capture_output=True, text=True, check=False, timeout=180,
        )
        self.assertEqual(hashlib.sha256(WORKBOOK.read_bytes()).digest(), digest,
                         "The reference workbook must not be modified")
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn("PASS: 240 Excel formula scenarios", result.stdout)


if __name__ == "__main__":
    unittest.main()
