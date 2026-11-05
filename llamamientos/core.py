from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import subprocess
import unicodedata
import warnings
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime, time
from pathlib import Path
from typing import Any, Iterable

from openpyxl import load_workbook
from openpyxl.utils import range_boundaries


DAILY_SHEET_RE = re.compile(r"^(?P<day>\d{2})_(?P<month>\d{2})$")
YEAR_RE = re.compile(r"(?<!\d)(20\d{2})(?!\d)")

MONTH_NAMES = {
    1: "ENERO",
    2: "FEBRERO",
    3: "MARZO",
    4: "ABRIL",
    5: "MAYO",
    6: "JUNIO",
    7: "JULIO",
    8: "AGOSTO",
    9: "SEPTIEMBRE",
    10: "OCTUBRE",
    11: "NOVIEMBRE",
    12: "DICIEMBRE",
}

EXPECTED_TARGET_HEADERS = (
    "FECHA",
    "OPERARIO",
    "NOMBRE",
    "CONTRATO",
    "ESTADO",
    "SACOS",
    "CAJAS",
    "FRIO",
    "ELECTRODOM",
    "OP GENERALES",
    "VEHICULO",
)


class LlamamientosError(RuntimeError):
    """Error controlado que puede mostrarse directamente al usuario."""


@dataclass(frozen=True)
class Record:
    work_date: date
    operator_id: int
    name: Any
    contract: Any
    status: Any
    source_sheet: str
    source_row: int

    @property
    def key(self) -> tuple[date, int]:
        return self.work_date, self.operator_id

    def target_values(self) -> tuple[Any, ...]:
        return (
            datetime.combine(self.work_date, time.min),
            self.operator_id,
            self.name,
            self.contract,
            self.status,
        )


@dataclass(frozen=True)
class DaySummary:
    work_date: date
    sheet_name: str
    record_count: int
    sheet_state: str


@dataclass
class AnalysisResult:
    source_path: Path
    target_path: Path
    month: int
    year: int
    target_sheet: str
    table_name: str
    table_ref: str
    records: list[Record]
    days: list[DaySummary]
    existing_record_count: int
    target_capacity: int
    warnings: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)

    @property
    def can_import(self) -> bool:
        return bool(self.records) and not self.errors

    @property
    def period_label(self) -> str:
        return f"{MONTH_NAMES[self.month].title()} {self.year}"


@dataclass(frozen=True)
class ImportResult:
    target_path: Path
    backup_path: Path
    target_sheet: str
    imported_records: int


@dataclass(frozen=True)
class _TargetInfo:
    sheet_name: str
    table_name: str
    table_ref: str
    min_col: int
    min_row: int
    max_col: int
    max_row: int
    formula_row: int
    formula_templates: tuple[str, ...]
    existing_count: int
    duplicate_keys: tuple[tuple[Any, Any], ...]

    @property
    def capacity(self) -> int:
        return self.max_row - self.min_row


def analyze_files(source_path: str | os.PathLike[str], target_path: str | os.PathLike[str]) -> AnalysisResult:
    source = _validated_path(source_path, "origen", {".xlsx", ".xlsm"})
    target = _validated_path(target_path, "destino", {".xlsx"})
    if source.resolve() == target.resolve():
        raise LlamamientosError("El archivo de origen y el de destino no pueden ser el mismo.")

    records, days, month, year, source_warnings, source_errors, source_x = _read_source(source)
    target_sheet = f"{MONTH_NAMES[month]} {year}"
    target_info, target_warnings, target_errors, target_x = _read_target(target, target_sheet)

    warnings_list = source_warnings + target_warnings
    errors = source_errors + target_errors

    if source_x and target_x and source_x != target_x:
        warnings_list.append(
            "La tabla auxiliar X no coincide entre el origen y el destino. "
            "Las fórmulas F:K usarán la versión existente en la herramienta."
        )

    if target_info.duplicate_keys:
        examples = ", ".join(_format_key(key) for key in target_info.duplicate_keys[:5])
        warnings_list.append(
            f"El mes de destino contiene {len(target_info.duplicate_keys)} duplicado(s) "
            f"por fecha y operario ({examples}). La sustitución mensual los eliminará."
        )

    return AnalysisResult(
        source_path=source,
        target_path=target,
        month=month,
        year=year,
        target_sheet=target_sheet,
        table_name=target_info.table_name,
        table_ref=target_info.table_ref,
        records=records,
        days=days,
        existing_record_count=target_info.existing_count,
        target_capacity=target_info.capacity,
        warnings=_deduplicate(warnings_list),
        errors=_deduplicate(errors),
    )


def import_month(
    source_path: str | os.PathLike[str],
    target_path: str | os.PathLike[str],
) -> ImportResult:
    """Analiza de nuevo, crea una copia de seguridad e importa el mes completo."""
    _ensure_excel_closed()
    analysis = analyze_files(source_path, target_path)
    if analysis.errors:
        raise LlamamientosError("No se puede importar:\n\n" + "\n".join(f"- {x}" for x in analysis.errors))
    if not analysis.records:
        raise LlamamientosError("No se han encontrado registros válidos para importar.")

    # Se comprueba otra vez justo antes de copiar y escribir por si Excel se ha
    # abierto mientras se analizaban los libros.
    _ensure_excel_closed()

    target = analysis.target_path
    backup_dir = target.parent / "Copias de seguridad llamamientos"
    backup_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = backup_dir / f"{target.stem}_antes_{analysis.year}_{analysis.month:02d}_{timestamp}{target.suffix}"
    shutil.copy2(target, backup)

    data_file = target.parent / f".{target.stem}.llamamientos-importacion.json"
    try:
        # Excel abre y guarda directamente el libro seleccionado por el usuario.
        # La copia anterior ya esta protegida en la carpeta de copias de seguridad.
        _write_import_with_excel(analysis, target, data_file)
        _validate_written_file(target, analysis)
    except PermissionError as exc:
        raise LlamamientosError(
            "No se ha podido guardar la herramienta original. Comprueba que el archivo no esté "
            "abierto en Excel ni bloqueado por OneDrive. La copia de seguridad se conserva en:\n"
            f"{backup}"
        ) from exc
    except Exception:
        raise
    finally:
        for disposable in (data_file,):
            if disposable.exists():
                try:
                    disposable.unlink()
                except PermissionError:
                    pass

    return ImportResult(
        target_path=target,
        backup_path=backup,
        target_sheet=analysis.target_sheet,
        imported_records=len(analysis.records),
    )


def _ensure_excel_closed() -> None:
    """Evita que otra instancia de Excel provoque copias de conflicto de OneDrive."""
    if os.name != "nt":
        return

    tasklist = shutil.which("tasklist.exe") or shutil.which("tasklist")
    if not tasklist:
        raise LlamamientosError(
            "No se ha podido comprobar si Microsoft Excel está abierto. "
            "Cierra Excel e inténtalo de nuevo."
        )

    creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    completed = subprocess.run(
        [tasklist, "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"],
        capture_output=True,
        text=True,
        errors="replace",
        creationflags=creationflags,
        check=False,
    )
    if completed.returncode != 0:
        raise LlamamientosError(
            "No se ha podido comprobar si Microsoft Excel está abierto. "
            "Cierra Excel e inténtalo de nuevo."
        )
    if "EXCEL.EXE" in (completed.stdout or "").upper():
        raise LlamamientosError(
            "Microsoft Excel está abierto. Cierra TODAS las ventanas de Excel antes de importar.\n\n"
            "No se ha creado ninguna copia de seguridad y no se ha modificado el archivo original."
        )


def _validated_path(
    value: str | os.PathLike[str], label: str, allowed_suffixes: set[str]
) -> Path:
    path = Path(value).expanduser()
    if not path.is_file():
        raise LlamamientosError(f"No se encuentra el archivo de {label}:\n{path}")
    if path.suffix.lower() not in allowed_suffixes:
        accepted = ", ".join(sorted(allowed_suffixes))
        raise LlamamientosError(f"El archivo de {label} debe tener formato {accepted}.")
    return path


def _read_source(
    path: Path,
) -> tuple[list[Record], list[DaySummary], int, int, list[str], list[str], str | None]:
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        workbook = load_workbook(
            path,
            data_only=True,
            read_only=False,
            keep_vba=path.suffix.lower() == ".xlsm",
            keep_links=False,
        )

    try:
        daily = []
        for worksheet in workbook.worksheets:
            match = DAILY_SHEET_RE.fullmatch(worksheet.title)
            if match:
                daily.append((worksheet, int(match.group("day")), int(match.group("month"))))

        if not daily:
            raise LlamamientosError("El origen no contiene hojas diarias con el formato 01_10, 02_10, etc.")

        months = {month for _, _, month in daily}
        if len(months) != 1:
            raise LlamamientosError(
                "El origen contiene hojas diarias de más de un mes: "
                + ", ".join(f"{month:02d}" for month in sorted(months))
            )
        month = months.pop()
        if month not in MONTH_NAMES:
            raise LlamamientosError(f"El mes {month:02d} de las hojas diarias no es válido.")

        year = _detect_year(path, daily)
        warnings_list: list[str] = []
        errors: list[str] = []
        records: list[Record] = []
        days: list[DaySummary] = []
        seen: dict[tuple[date, int], Record] = {}
        textual_statuses: defaultdict[str, set[str]] = defaultdict(set)

        for worksheet, day, sheet_month in sorted(daily, key=lambda item: item[1]):
            try:
                work_date = date(year, sheet_month, day)
            except ValueError:
                errors.append(f"La hoja {worksheet.title} no representa una fecha válida de {year}.")
                continue

            header_date = _as_date(worksheet["E9"].value)
            if header_date is None:
                warnings_list.append(f"{worksheet.title}: E9 no contiene una fecha válida.")
            elif header_date != work_date:
                warnings_list.append(
                    f"{worksheet.title}: E9 contiene {header_date.strftime('%d/%m/%Y')} "
                    f"pero el nombre de la hoja corresponde a {work_date.strftime('%d/%m/%Y')}. "
                    "Se usará la fecha del nombre de la hoja."
                )

            day_records = 0
            for row in range(10, worksheet.max_row + 1):
                cells = [worksheet.cell(row, column) for column in range(4, 8)]
                raw = [cell.value for cell in cells]
                status = raw[3]
                if _is_blank(status):
                    continue

                error_cells = [cell.coordinate for cell in cells if _is_excel_error(cell)]
                if error_cells:
                    errors.append(
                        f"{worksheet.title}, fila {row}: hay errores de Excel en "
                        f"{', '.join(error_cells)}. La fila no se importará."
                    )
                    continue

                try:
                    operator_id = _normalize_operator_id(raw[0])
                except ValueError:
                    errors.append(
                        f"{worksheet.title}, fila {row}: el nº de operario {raw[0]!r} no es válido."
                    )
                    continue

                if _is_blank(raw[1]) or _is_blank(raw[2]):
                    errors.append(
                        f"{worksheet.title}, fila {row}: nombre o contrato vacío con un estado informado."
                    )
                    continue

                record = Record(
                    work_date=work_date,
                    operator_id=operator_id,
                    name=raw[1],
                    contract=raw[2],
                    status=status,
                    source_sheet=worksheet.title,
                    source_row=row,
                )
                previous = seen.get(record.key)
                if previous:
                    errors.append(
                        f"Duplicado en el origen: {work_date.strftime('%d/%m/%Y')}, operario "
                        f"{operator_id}, en {previous.source_sheet}!{previous.source_row} y "
                        f"{worksheet.title}!{row}."
                    )
                    continue

                seen[record.key] = record
                records.append(record)
                day_records += 1
                if isinstance(status, str):
                    cleaned = status.strip()
                    textual_statuses[cleaned.casefold()].add(cleaned)

            days.append(
                DaySummary(
                    work_date=work_date,
                    sheet_name=worksheet.title,
                    record_count=day_records,
                    sheet_state=worksheet.sheet_state,
                )
            )

        for variants in textual_statuses.values():
            if len(variants) > 1:
                warnings_list.append(
                    "El origen contiene variantes de mayúsculas/minúsculas para un mismo estado: "
                    + ", ".join(sorted(repr(item) for item in variants))
                    + ". Se conservarán tal como están."
                )

        if not records:
            errors.append("No hay filas con un estado informado en la columna G.")

        records.sort(key=lambda item: (item.work_date, item.source_row))
        return records, days, month, year, warnings_list, errors, _sheet_signature(workbook, "X")
    finally:
        workbook.close()


def _detect_year(path: Path, daily: list[tuple[Any, int, int]]) -> int:
    header_years = []
    for worksheet, _, _ in daily:
        value = _as_date(worksheet["E9"].value)
        if value:
            header_years.append(value.year)

    if header_years:
        counts = Counter(header_years)
        year, count = counts.most_common(1)[0]
        if len(counts) == 1 or count >= max(2, len(header_years) // 2):
            return year

    filename_years = [int(value) for value in YEAR_RE.findall(path.stem)]
    if filename_years:
        return filename_years[-1]
    raise LlamamientosError(
        "No se ha podido determinar el año. Debe aparecer en E9 o en el nombre del archivo."
    )


def _read_target(
    path: Path, target_sheet: str
) -> tuple[_TargetInfo, list[str], list[str], str | None]:
    workbook = load_workbook(path, data_only=False, read_only=False, keep_links=True)
    try:
        errors: list[str] = []
        warnings_list: list[str] = []
        if target_sheet not in workbook.sheetnames:
            available = ", ".join(workbook.sheetnames)
            raise LlamamientosError(
                f"No existe la hoja '{target_sheet}' en la herramienta. Hojas disponibles: {available}"
            )
        worksheet = workbook[target_sheet]
        table = _find_month_table(worksheet)
        min_col, min_row, max_col, max_row = range_boundaries(table.ref)

        if max_col - min_col + 1 != 11:
            errors.append(
                f"La tabla {table.name} de {target_sheet} debe tener 11 columnas A:K; "
                f"actualmente ocupa {table.ref}."
            )

        headers = [worksheet.cell(min_row, column).value for column in range(min_col, max_col + 1)]
        if not _headers_match(headers):
            errors.append(
                f"Las cabeceras de la tabla {table.name} no coinciden con FECHA, Nº OPERARIO, "
                "NOMBRE, CONTRATO, ESTADO y las seis columnas calculadas."
            )

        formula_row = 0
        formula_templates: tuple[str, ...] = ()
        for row in range(min_row + 1, max_row + 1):
            formulas = tuple(worksheet.cell(row, column).value for column in range(min_col + 5, min_col + 11))
            if len(formulas) == 6 and all(_is_formula(value) for value in formulas):
                formula_row = row
                formula_templates = formulas
                break
        if not formula_row:
            errors.append(f"No se ha encontrado una fila modelo con fórmulas F:K en {target_sheet}.")

        existing_count = 0
        duplicates = []
        seen: set[tuple[Any, Any]] = set()
        for row in range(min_row + 1, max_row + 1):
            values = [worksheet.cell(row, column).value for column in range(min_col, min_col + 5)]
            if not any(not _is_blank(value) for value in values):
                continue
            existing_count += 1
            key = (_as_date(values[0]) or values[0], values[1])
            if key in seen:
                duplicates.append(key)
            else:
                seen.add(key)

        info = _TargetInfo(
            sheet_name=target_sheet,
            table_name=table.name,
            table_ref=table.ref,
            min_col=min_col,
            min_row=min_row,
            max_col=max_col,
            max_row=max_row,
            formula_row=formula_row,
            formula_templates=formula_templates,
            existing_count=existing_count,
            duplicate_keys=tuple(duplicates),
        )
        return info, warnings_list, errors, _sheet_signature(workbook, "X")
    finally:
        workbook.close()


def _write_import_with_excel(
    analysis: AnalysisResult, workbook_path: Path, data_path: Path
) -> None:
    if os.name != "nt":
        raise LlamamientosError(
            "La importación requiere Windows con Microsoft Excel instalado. El análisis sí puede "
            "ejecutarse en otros sistemas."
        )

    script_path = Path(__file__).resolve().parents[1] / "excel_import.ps1"
    if not script_path.is_file():
        raise LlamamientosError(f"No se encuentra el componente de Excel:\n{script_path}")

    payload = {
        "records": [
            [
                record.work_date.isoformat(),
                record.operator_id,
                _json_value(record.name),
                _json_value(record.contract),
                _json_value(record.status),
            ]
            for record in analysis.records
        ]
    }
    data_path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")

    powershell = shutil.which("powershell.exe") or shutil.which("pwsh.exe")
    if not powershell:
        raise LlamamientosError("No se encuentra PowerShell, necesario para controlar Excel.")

    command = [
        powershell,
        "-NoProfile",
        "-NonInteractive",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(script_path),
        "-WorkbookPath",
        str(workbook_path),
        "-SheetName",
        analysis.target_sheet,
        "-TableName",
        analysis.table_name,
        "-DataPath",
        str(data_path),
    ]
    creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    try:
        completed = subprocess.run(
            command,
            capture_output=True,
            text=True,
            errors="replace",
            timeout=600,
            creationflags=creationflags,
            check=False,
        )
    except subprocess.TimeoutExpired as exc:
        raise LlamamientosError(
            "Excel ha tardado más de diez minutos en completar la importación. "
            "Revisa el archivo original y, si fuera necesario, restáuralo desde la copia de seguridad."
        ) from exc

    if completed.returncode != 0:
        detail = (completed.stderr or completed.stdout or "Error desconocido de Excel").strip()
        raise LlamamientosError(
            "Excel no ha podido guardar la importación en la herramienta original. "
            f"La copia de seguridad anterior se conserva.\n\n{detail}"
        )


def _validate_written_file(path: Path, analysis: AnalysisResult) -> None:
    workbook = load_workbook(path, data_only=False, read_only=False, keep_links=True)
    try:
        worksheet = workbook[analysis.target_sheet]
        table = worksheet.tables[analysis.table_name]
        min_col, min_row, max_col, max_row = range_boundaries(table.ref)
        expected_max_row = min_row + len(analysis.records)
        if max_row != expected_max_row or max_col - min_col + 1 != 11:
            raise LlamamientosError(
                f"La validación de la tabla ha fallado: {table.ref}; se esperaba terminar en la fila "
                f"{expected_max_row}."
            )

        seen: set[tuple[date, int]] = set()
        for offset, expected in enumerate(analysis.records, start=1):
            row = min_row + offset
            actual_date = _as_date(worksheet.cell(row, min_col).value)
            actual_id = _normalize_operator_id(worksheet.cell(row, min_col + 1).value)
            actual = (
                actual_date,
                actual_id,
                worksheet.cell(row, min_col + 2).value,
                worksheet.cell(row, min_col + 3).value,
                worksheet.cell(row, min_col + 4).value,
            )
            wanted = (
                expected.work_date,
                expected.operator_id,
                expected.name,
                expected.contract,
                expected.status,
            )
            if actual != wanted:
                raise LlamamientosError(
                    f"La validación de datos ha fallado en {analysis.target_sheet}, fila {row}."
                )
            if expected.key in seen:
                raise LlamamientosError(
                    f"La validación ha encontrado un duplicado en {analysis.target_sheet}, fila {row}."
                )
            seen.add(expected.key)

            formulas = [worksheet.cell(row, column).value for column in range(min_col + 5, min_col + 11)]
            if not all(_is_formula(value) and "#REF!" not in value.upper() for value in formulas):
                raise LlamamientosError(
                    f"La validación de fórmulas F:K ha fallado en {analysis.target_sheet}, fila {row}."
                )
    finally:
        workbook.close()


def _find_month_table(worksheet: Any) -> Any:
    tables = list(worksheet.tables.values())
    if not tables:
        raise LlamamientosError(f"La hoja {worksheet.title} no contiene una tabla de Excel.")
    for table in tables:
        min_col, min_row, max_col, _ = range_boundaries(table.ref)
        if min_col == 1 and min_row == 1 and max_col == 11:
            return table
    if len(tables) == 1:
        return tables[0]
    raise LlamamientosError(
        f"No se ha podido identificar la tabla mensual de A:K en la hoja {worksheet.title}."
    )


def _find_formula_templates(
    worksheet: Any, first_row: int, last_row: int, first_col: int, last_col: int
) -> tuple[int, tuple[str, ...]]:
    for row in range(first_row, last_row + 1):
        formulas = tuple(worksheet.cell(row, column).value for column in range(first_col, last_col + 1))
        if formulas and all(_is_formula(value) for value in formulas):
            return row, formulas
    return 0, ()


def _headers_match(headers: Iterable[Any]) -> bool:
    normalized = [_normalize_text(value) for value in headers]
    if len(normalized) != len(EXPECTED_TARGET_HEADERS):
        return False
    return all(expected in actual for actual, expected in zip(normalized, EXPECTED_TARGET_HEADERS))


def _normalize_text(value: Any) -> str:
    text = "" if value is None else str(value)
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^A-Z0-9]+", " ", text.upper()).strip()


def _normalize_operator_id(value: Any) -> int:
    if isinstance(value, bool) or value is None:
        raise ValueError
    if isinstance(value, int):
        return value
    if isinstance(value, float) and value.is_integer():
        return int(value)
    if isinstance(value, str):
        stripped = value.strip()
        if stripped.isdigit():
            return int(stripped)
    raise ValueError


def _is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def _is_formula(value: Any) -> bool:
    return isinstance(value, str) and value.startswith("=")


def _is_excel_error(cell: Any) -> bool:
    value = cell.value
    return cell.data_type == "e" or (isinstance(value, str) and value.lstrip().startswith("#"))


def _as_date(value: Any) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return None


def _sheet_signature(workbook: Any, sheet_name: str) -> str | None:
    if sheet_name not in workbook.sheetnames:
        return None
    worksheet = workbook[sheet_name]
    digest = hashlib.sha256()
    digest.update(f"{worksheet.max_row}:{worksheet.max_column}".encode())
    for row in worksheet.iter_rows():
        for cell in row:
            digest.update(repr(cell.value).encode("utf-8", "backslashreplace"))
            digest.update(b"\0")
    return digest.hexdigest()


def _format_key(key: tuple[Any, Any]) -> str:
    value_date, operator = key
    if isinstance(value_date, date):
        return f"{value_date.strftime('%d/%m/%Y')} · {operator}"
    return f"{value_date} · {operator}"


def _json_value(value: Any) -> Any:
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    return str(value)


def _deduplicate(items: Iterable[str]) -> list[str]:
    return list(dict.fromkeys(items))
