from __future__ import annotations

import tempfile
import unittest
from subprocess import CompletedProcess
from datetime import datetime
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

from llamamientos.core import LlamamientosError, _ensure_excel_closed, analyze_files, import_month


HEADERS = [
    "FECHA",
    "Nº OPERARIO",
    "NOMBRE",
    "CONTRATO",
    "ESTADO",
    "SACOS",
    "CAJAS",
    "FRIO",
    "ELECTRODOM.",
    "OP. GENERALES",
    "VEHICULO",
]


def create_source(path: Path, duplicate: bool = False) -> None:
    workbook = Workbook()
    workbook.remove(workbook.active)
    for day in range(1, 4):
        sheet = workbook.create_sheet(f"{day:02d}_10")
        sheet["D9"], sheet["E9"], sheet["F9"], sheet["G9"] = (
            "Nº",
            datetime(2026, 10, day),
            "CONTRATO",
            "ESTADO",
        )
        if day == 1:
            sheet["D10"], sheet["E10"], sheet["F10"], sheet["G10"] = 101, "Ana", "200 7 H", 1
            sheet["D11"], sheet["E11"], sheet["F11"], sheet["G11"] = 102, "Luis", 100, "ND"
            sheet["D12"], sheet["E12"], sheet["F12"], sheet["G12"] = 103, "Sin estado", 100, None
            if duplicate:
                sheet["D13"], sheet["E13"], sheet["F13"], sheet["G13"] = 101, "Ana", "200 7 H", 1
        elif day == 2:
            sheet["D10"], sheet["E10"], sheet["F10"], sheet["G10"] = 104, "Marta", "300 6 H", "E"
    x_sheet = workbook.create_sheet("X")
    x_sheet.append(["ID", "NOMBRE", "OTRO", "SACOS", "CAJAS", "FRIO", "ELECTRO", "GENERALES", "VEHICULO"])
    x_sheet.append([101, "Ana", "", "X", "X", "", "", "", "SI"])
    workbook.save(path)


def create_target(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "OCTUBRE 2026"
    sheet.append(HEADERS)
    for row in range(2, 7):
        sheet.cell(row, 1).value = datetime(2026, 10, 1) if row == 2 else None
        sheet.cell(row, 2).value = 999 if row == 2 else None
        for column in range(6, 12):
            sheet.cell(row, column).value = (
                f'=IFERROR(VLOOKUP([@[Nº OPERARIO]],X!$A:$I,{column - 2},FALSE),"")'
            )
    table = Table(displayName="Octubre26", ref="A1:K6")
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    sheet.add_table(table)

    x_sheet = workbook.create_sheet("X")
    x_sheet.append(["ID", "NOMBRE", "OTRO", "SACOS", "CAJAS", "FRIO", "ELECTRO", "GENERALES", "VEHICULO"])
    x_sheet.append([101, "Ana", "", "X", "X", "", "", "", "SI"])
    workbook.save(path)


class CoreTests(unittest.TestCase):
    @patch("llamamientos.core.shutil.which", return_value="tasklist.exe")
    @patch("llamamientos.core.subprocess.run")
    def test_import_is_blocked_while_excel_is_open(self, run_mock, _which_mock) -> None:
        run_mock.return_value = CompletedProcess(
            args=[], returncode=0, stdout='"EXCEL.EXE","1234","Console","1","100 K"\n', stderr=""
        )
        with self.assertRaisesRegex(LlamamientosError, "Excel está abierto"):
            _ensure_excel_closed()

    def test_analysis_keeps_nd_and_ignores_blank_status(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source, target = root / "origen.xlsx", root / "destino.xlsx"
            create_source(source)
            create_target(target)
            result = analyze_files(source, target)
            self.assertTrue(result.can_import)
            self.assertEqual(len(result.records), 3)
            self.assertIn("ND", [record.status for record in result.records])
            self.assertEqual([day.record_count for day in result.days], [2, 1, 0])

    def test_duplicate_in_source_blocks_import(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source, target = root / "origen.xlsx", root / "destino.xlsx"
            create_source(source, duplicate=True)
            create_target(target)
            result = analyze_files(source, target)
            self.assertFalse(result.can_import)
            self.assertTrue(any("Duplicado en el origen" in error for error in result.errors))
            with self.assertRaises(LlamamientosError):
                import_month(source, target)

    def test_analysis_recognizes_target_table_and_formula_columns(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source, target = root / "origen.xlsx", root / "destino.xlsx"
            create_source(source)
            create_target(target)
            result = analyze_files(source, target)
            self.assertTrue(result.can_import)
            self.assertEqual(result.target_sheet, "OCTUBRE 2026")
            self.assertEqual(result.table_name, "Octubre26")
            self.assertEqual(result.table_ref, "A1:K6")
            self.assertEqual(result.target_capacity, 5)


if __name__ == "__main__":
    unittest.main()
