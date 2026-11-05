from __future__ import annotations

import hashlib
import os
import shutil
import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

from llamamientos.core import import_month


def sheet_signature(worksheet) -> str:
    digest = hashlib.sha256()
    digest.update((worksheet.title + worksheet.calculate_dimension()).encode())
    for row in worksheet.iter_rows():
        for cell in row:
            digest.update(repr(cell.value).encode("utf-8", "backslashreplace"))
            digest.update(b"\0")
    return digest.hexdigest()


@unittest.skipUnless(os.environ.get("RUN_REAL_EXCEL_TESTS") == "1", "prueba real desactivada")
class RealWorkbookIntegrationTests(unittest.TestCase):
    def test_october_import_on_temporary_copy(self) -> None:
        root = Path(__file__).resolve().parents[1]
        source = root / "10. OCTUBRE LLAMAMIENTOS 2026.xlsm"
        original = root / "HERRAMIENTA LLAMAMIENTOS 2026.xlsx"

        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / original.name
            shutil.copy2(original, target)
            original_before_import = target.read_bytes()

            before = load_workbook(target, data_only=False, keep_links=True)
            try:
                before_names = before.sheetnames[:]
                before_other = {
                    name: sheet_signature(before[name])
                    for name in before_names
                    if name != "OCTUBRE 2026"
                }
                before_tables = {
                    name: {table.name: table.ref for table in before[name].tables.values()}
                    for name in before_names
                }
                before_calculation = (
                    before.calculation.calcMode,
                    before.calculation.fullCalcOnLoad,
                    before.calculation.forceFullCalc,
                )
            finally:
                before.close()

            result = import_month(source, target)
            self.assertEqual(result.imported_records, 4812)
            self.assertTrue(result.backup_path.is_file())
            self.assertEqual(result.backup_path.read_bytes(), original_before_import)
            self.assertNotEqual(target.read_bytes(), original_before_import)

            after = load_workbook(target, data_only=False, keep_links=True)
            try:
                self.assertEqual(after.sheetnames, before_names)
                after_other = {
                    name: sheet_signature(after[name])
                    for name in after.sheetnames
                    if name != "OCTUBRE 2026"
                }
                self.assertEqual(after_other, before_other)

                for name in after.sheetnames:
                    if name == "OCTUBRE 2026":
                        continue
                    refs = {table.name: table.ref for table in after[name].tables.values()}
                    self.assertEqual(refs, before_tables[name])

                sheet = after["OCTUBRE 2026"]
                self.assertEqual(sheet.tables["Octubre26"].ref, "A1:K4813")
                keys = set()
                for row in range(2, 4814):
                    value = sheet.cell(row, 1).value
                    work_date = value.date() if hasattr(value, "date") else value
                    key = (work_date, sheet.cell(row, 2).value)
                    self.assertNotIn(key, keys)
                    keys.add(key)
                    for column in range(6, 12):
                        formula = sheet.cell(row, column).value
                        self.assertIsInstance(formula, str)
                        self.assertTrue(formula.startswith("="))
                        self.assertNotIn("#REF!", formula.upper())
                self.assertEqual(len(keys), 4812)
                self.assertEqual(
                    (
                        after.calculation.calcMode,
                        after.calculation.fullCalcOnLoad,
                        after.calculation.forceFullCalc,
                    ),
                    before_calculation,
                )
            finally:
                after.close()


if __name__ == "__main__":
    unittest.main()
