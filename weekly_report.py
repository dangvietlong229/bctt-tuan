#!/usr/bin/env python3
"""Build the weekly MBS market report from the existing project assets.

The workflow deliberately stops at a review gate before the final PPTX/PDF:

1. route/process/prepare -> draft deck + facts + GPT review prompt
2. review/edit review/commentary.json
3. finalize -> final deck + PDF

PowerPoint and Excel are controlled through their native Windows or macOS
automation APIs so the current deck's tables, typography, bullets, and layout
stay intact.
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import importlib.util
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


if os.name == "nt":
    try:
        import truststore

        # requests/urllib should validate HTTPS through the Windows certificate
        # store so managed corporate root certificates continue to be trusted.
        truststore.inject_into_ssl()
    except ImportError:
        pass

    # Windows PowerShell 5 often starts Python with a legacy console encoding.
    # The workflow logs Vietnamese text, so force a predictable Unicode stream.
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")

import urllib.error
import urllib.request


ROOT = Path(__file__).resolve().parent
CONFIG_PATH = ROOT / "report_config.json"
SCRIPT_DIR = ROOT / "scripts"
UNIT_SEPARATOR = "\x1f"
RECORD_SEPARATOR = "\x1e"
USE_EXCEL_FALLBACK = False
REPORT_FONT = "Arial"
FORBIDDEN_REPORT_FONTS = {"calibri", "aptos"}


def log(message: str) -> None:
    print(message, flush=True)


def write_text_atomic(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(f".{path.name}.{os.getpid()}.publishing")
    try:
        temporary.write_text(text, encoding="utf-8")
        os.replace(temporary, path)
    finally:
        if temporary.exists():
            temporary.unlink()


def write_json_atomic(path: Path, value: Any) -> None:
    write_text_atomic(path, json.dumps(value, ensure_ascii=False, indent=2))


def load_config() -> dict[str, Any]:
    with CONFIG_PATH.open(encoding="utf-8") as stream:
        return json.load(stream)


def load_dotenv(path: Path) -> None:
    if not path.exists():
        return
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value


def run_applescript(script_name: str, *args: object, timeout: int = 240) -> str:
    script_path = SCRIPT_DIR / script_name
    command = ["osascript", str(script_path), *(str(arg) for arg in args)]
    result = subprocess.run(
        command,
        cwd=ROOT,
        text=True,
        capture_output=True,
        timeout=timeout,
    )
    if result.returncode != 0:
        detail = (result.stderr or result.stdout).strip()
        raise RuntimeError(f"Office automation failed ({script_name}): {detail}")
    return result.stdout.rstrip("\n")


def run_windows_office(action: str, timeout: int = 300, **arguments: object) -> str:
    powershell = shutil.which("powershell.exe") or shutil.which("powershell") or shutil.which("pwsh")
    if not powershell:
        raise FileNotFoundError("Không tìm thấy PowerShell để điều khiển Microsoft Office trên Windows.")
    script_path = SCRIPT_DIR / "windows_office.ps1"
    if not script_path.exists():
        raise FileNotFoundError(f"Thiếu script tự động hóa Windows: {script_path}")
    command = [
        powershell,
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(script_path),
        "-Action",
        action,
    ]
    for key, value in arguments.items():
        if value is None:
            continue
        command.extend([f"-{key}", str(value)])
    result = subprocess.run(
        command,
        cwd=ROOT,
        text=True,
        encoding="utf-8",
        errors="replace",
        capture_output=True,
        timeout=timeout,
    )
    if result.returncode != 0:
        detail = (result.stderr or result.stdout).strip()
        raise RuntimeError(f"Windows Office automation failed ({action}): {detail}")
    return result.stdout.strip()


def assert_presentation_fonts(pptx_path: Path) -> None:
    """Fail publishing if a generated slide still explicitly requests Calibri/Aptos."""
    offenders: set[str] = set()
    with zipfile.ZipFile(pptx_path) as archive:
        for name in archive.namelist():
            if not name.startswith("ppt/slides/slide") or not name.endswith(".xml"):
                continue
            xml = archive.read(name).decode("utf-8", errors="ignore")
            for font in re.findall(r'<a:(?:latin|ea|cs)[^>]*typeface="([^"]*)"', xml):
                if font.strip().lower() in FORBIDDEN_REPORT_FONTS:
                    offenders.add(font)
    if offenders:
        raise RuntimeError(
            f"PowerPoint vẫn chứa font không được phép: {', '.join(sorted(offenders))}."
        )


def assert_pdf_fonts(pdf_path: Path) -> None:
    """Verify all exported, visible PDF text uses Arial (excluding symbol fonts)."""
    import pdfplumber

    offenders: set[str] = set()
    with pdfplumber.open(pdf_path) as document:
        for page in document.pages:
            for character in page.chars:
                font = str(character.get("fontname") or "")
                lowered = font.lower()
                allowed_symbol = any(name in lowered for name in ("symbol", "wingdings"))
                if font and "arial" not in lowered and not allowed_symbol:
                    offenders.add(font)
    if offenders:
        raise RuntimeError(f"PDF vẫn chứa font không phải Arial: {', '.join(sorted(offenders))}.")


def export_pdf(pptx_path: Path, pdf_path: Path) -> None:
    if os.name == "nt":
        run_windows_office(
            "ExportPdf",
            PresentationPath=pptx_path.resolve(),
            PdfPath=pdf_path.resolve(),
        )
        if not pdf_path.exists():
            raise RuntimeError(f"PowerPoint không tạo được PDF: {pdf_path}")
        assert_pdf_fonts(pdf_path)
        return

    soffice = shutil.which("soffice")
    if soffice:
        with tempfile.TemporaryDirectory(prefix="weekly_report_pdf_") as temporary_dir:
            result = subprocess.run(
                [
                    soffice,
                    "--headless",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    temporary_dir,
                    str(pptx_path.resolve()),
                ],
                cwd=ROOT,
                text=True,
                capture_output=True,
                timeout=300,
            )
            generated = Path(temporary_dir) / f"{pptx_path.stem}.pdf"
            if result.returncode == 0 and generated.exists():
                shutil.copy2(generated, pdf_path)
                assert_pdf_fonts(pdf_path)
                return
            log("LibreOffice không xuất được PDF; chuyển sang bộ render dự phòng.")

    node_binary, node_modules = locate_artifact_tool()
    python_candidates = [
        node_modules.parent.parent / "python" / "bin" / "python3",
        node_modules.parent.parent / "python" / "python.exe",
    ]
    bundled_python = next((path for path in python_candidates if path.exists()), None)
    if bundled_python is None:
        raise FileNotFoundError("Không tìm thấy Python đi kèm Codex để tạo PDF dự phòng.")
    with tempfile.TemporaryDirectory(prefix="weekly_report_pdf_render_") as temporary_dir:
        image_dir = Path(temporary_dir) / "slides"
        render = subprocess.run(
            [
                str(node_binary),
                str(SCRIPT_DIR / "render_pptx_images.mjs"),
                str(node_modules),
                str(pptx_path.resolve()),
                str(image_dir),
            ],
            cwd=ROOT,
            text=True,
            capture_output=True,
            timeout=300,
        )
        images = sorted(image_dir.glob("slide-*.png"))
        if render.returncode != 0 or not images:
            detail = (render.stderr or render.stdout).strip()
            raise RuntimeError(f"Không thể render slide để tạo PDF: {detail}")
        result = subprocess.run(
            [
                str(bundled_python),
                str(SCRIPT_DIR / "images_to_pdf.py"),
                str(pdf_path),
                *(str(image) for image in images),
            ],
            cwd=ROOT,
            text=True,
            capture_output=True,
            timeout=300,
        )
        if result.returncode != 0 or not pdf_path.exists():
            detail = (result.stderr or result.stdout).strip()
            raise RuntimeError(f"Không thể ghép PDF dự phòng: {detail}")
        assert_pdf_fonts(pdf_path)


def extract_excel_range(
    workbook: Path,
    sheet: str,
    start_row: int,
    start_column: int,
    row_count: int,
    column_count: int,
    mode: str = "display",
) -> list[list[str]]:
    global USE_EXCEL_FALLBACK
    matrix: list[list[str]] | None = None
    if not USE_EXCEL_FALLBACK:
        try:
            if os.name == "nt":
                with tempfile.TemporaryDirectory(prefix="weekly_report_excel_") as temporary_dir:
                    output_path = Path(temporary_dir) / "range.json"
                    run_windows_office(
                        "ExtractExcelRange",
                        WorkbookPath=workbook.resolve(),
                        WorksheetName=sheet,
                        StartRow=start_row,
                        StartColumn=start_column,
                        RowCount=row_count,
                        ColumnCount=column_count,
                        OutputMode=mode,
                        OutputPath=output_path,
                    )
                    matrix = json.loads(output_path.read_text(encoding="utf-8-sig"))
            elif sys.platform == "darwin":
                raw = run_applescript(
                    "extract_excel_range.applescript",
                    workbook.resolve(),
                    sheet,
                    start_row,
                    start_column,
                    row_count,
                    column_count,
                    mode,
                )
                rows = raw.split(RECORD_SEPARATOR) if raw else []
                matrix = [row.split(UNIT_SEPARATOR) for row in rows]
            else:
                USE_EXCEL_FALLBACK = True
        except Exception as exc:
            USE_EXCEL_FALLBACK = True
            log(
                "Không thể đọc dữ liệu qua ứng dụng Excel "
                f"({exc}); chuyển sang bộ đọc dự phòng cho phần dữ liệu còn lại."
            )

    if matrix is None:
        book = openpyxl.load_workbook(workbook, read_only=True, data_only=True)
        worksheet = book[sheet]
        matrix = []
        for row in range(start_row, start_row + row_count):
            values: list[str] = []
            for column in range(start_column, start_column + column_count):
                cell = worksheet.cell(row=row, column=column)
                value = cell.value
                if mode == "raw":
                    values.append("" if value is None else str(value))
                elif value is None:
                    values.append("")
                elif isinstance(value, (dt.datetime, dt.date)):
                    values.append(value.strftime("%d/%m/%Y"))
                elif isinstance(value, (int, float)) and "%" in (cell.number_format or ""):
                    percent_match = re.search(r"0(?:\.(0+))?%", cell.number_format or "")
                    decimals = len(percent_match.group(1) or "") if percent_match else 1
                    values.append(format_vi_percent(value, decimals))
                elif isinstance(value, (int, float)) and any(token in (cell.number_format or "") for token in ("0", "#")):
                    values.append(format_vi_integer(value))
                else:
                    values.append(str(value))
            matrix.append(values)
        book.close()
    while len(matrix) < row_count:
        matrix.append([""] * column_count)
    for row in matrix:
        if len(row) < column_count:
            row.extend([""] * (column_count - len(row)))
    return [row[:column_count] for row in matrix[:row_count]]


from report_files import date_from_filename, parse_module_exclusions


def latest_file(patterns: list[str], as_of: dt.date | None = None) -> Path | None:
    matches: list[Path] = []
    for pattern in patterns:
        matches.extend(ROOT.glob(pattern))
    matches = [path for path in matches if path.is_file() and not path.name.startswith("~$")]
    if as_of is not None:
        dated = [(path, date_from_filename(path)) for path in matches]
        eligible = [(path, day) for path, day in dated if day is not None and day <= as_of]
        if eligible:
            return max(eligible, key=lambda item: (item[1], item[0].stat().st_mtime))[0]
        # Do not silently use a future-dated file. Undated legacy files remain a
        # fallback and are reported by validate_sources().
        matches = [path for path, day in dated if day is None]
    if not matches:
        return None
    return max(matches, key=lambda path: path.stat().st_mtime)


def locate_artifact_tool() -> tuple[Path, Path]:
    package_files = list(
        Path.home().glob(
            ".cache/codex-runtimes/**/dependencies/node/node_modules/@oai/artifact-tool/package.json"
        )
    )
    if not package_files:
        raise FileNotFoundError(
            "Không tìm thấy bộ công cụ xử lý Excel đi kèm Codex. Hãy mở dự án bằng Codex rồi chạy lại."
        )
    package_file = max(package_files, key=lambda path: path.stat().st_mtime)
    node_modules = package_file.parents[2]
    node_binary = node_modules.parent / "bin" / "node"
    if not node_binary.exists():
        node_from_path = shutil.which("node")
        if not node_from_path:
            raise FileNotFoundError("Không tìm thấy Node.js để xử lý file tự doanh.")
        node_binary = Path(node_from_path)
    return node_binary, node_modules


def ensure_proprietary_workbook(as_of: dt.date, config: dict[str, Any]) -> Path | None:
    patterns = config.get("proprietary_raw_patterns", {})
    buy = latest_file(patterns.get("buy", []), as_of)
    sell = latest_file(patterns.get("sell", []), as_of)
    if not buy or not sell:
        return None
    buy_date, sell_date = date_from_filename(buy), date_from_filename(sell)
    if buy_date and sell_date and buy_date != sell_date:
        raise ValueError(f"Dữ liệu tự doanh mua/bán không cùng kỳ: {buy.name} / {sell.name}")

    output = ROOT / "gd tu doanh" / f"Thong ke Tu doanh Co phieu_update {as_of:%d%m%y}.xlsx"
    if output.exists() and output.stat().st_mtime >= max(buy.stat().st_mtime, sell.stat().st_mtime):
        return output

    def read_top_rows(source_path: Path, side: str) -> list[list[object]]:
        source_book = openpyxl.load_workbook(source_path, read_only=True, data_only=True)
        source_sheet = source_book["Sheet1"]
        header = (
            ["Mã", "Giá", "Bán ròng (GT)", "Tỷ trọng bán của Khối Tự doanh"]
            if side == "sell"
            else ["Mã", "Giá", "Mua ròng (GT)", "Tỷ trọng mua của Khối Tự doanh"]
        )
        rows: list[list[object]] = [header]
        for row in range(6, 16):
            rows.append(
                [
                    source_sheet.cell(row=row, column=1).value,
                    source_sheet.cell(row=row, column=2).value,
                    source_sheet.cell(row=row, column=6).value,
                    source_sheet.cell(row=row, column=7).value,
                ]
            )
        source_book.close()
        return rows

    def style_top_sheet(worksheet: Any, matrix: list[list[object]]) -> None:
        worksheet.sheet_view.showGridLines = False
        worksheet.freeze_panes = "A2"
        dark_blue = PatternFill("solid", fgColor="17365D")
        stripe = PatternFill("solid", fgColor="F4F7FB")
        header_font = Font(name=REPORT_FONT, size=10, bold=True, color="FFFFFF")
        body_font = Font(name=REPORT_FONT, size=10, color="1F2937")
        ticker_font = Font(name=REPORT_FONT, size=10, bold=True, color="17365D")
        thin = Side(style="thin", color="D9E2F3")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        for row_index, row_values in enumerate(matrix, start=1):
            worksheet.row_dimensions[row_index].height = 34 if row_index == 1 else 21
            for column_index, value in enumerate(row_values, start=1):
                cell = worksheet.cell(row=row_index, column=column_index, value=value)
                cell.border = border
                cell.alignment = Alignment(
                    horizontal="center" if row_index == 1 or column_index == 1 else "right",
                    vertical="center",
                    wrap_text=row_index == 1,
                )
                if row_index == 1:
                    cell.fill = dark_blue
                    cell.font = header_font
                else:
                    cell.font = ticker_font if column_index == 1 else body_font
                    if row_index % 2 == 0:
                        cell.fill = stripe
                    if column_index in (2, 3):
                        cell.number_format = "#,##0"
                    elif column_index == 4:
                        cell.number_format = "0.00%"
        for column, width in {"A": 15, "B": 14, "C": 24, "D": 38}.items():
            worksheet.column_dimensions[column].width = width

    sold = read_top_rows(sell, "sell")
    bought = read_top_rows(buy, "buy")
    workbook = openpyxl.Workbook()
    sold_sheet = workbook.active
    sold_sheet.title = "Top_10_Ban_Rong"
    style_top_sheet(sold_sheet, sold)
    bought_sheet = workbook.create_sheet("Top_10_Mua_Rong")
    style_top_sheet(bought_sheet, bought)

    source_sheet = workbook.create_sheet("Nguon_Du_Lieu")
    source_sheet.sheet_view.showGridLines = False
    source_rows = [
        ["Nguồn", "Tệp FiinProX"],
        ["Top bán ròng 1 tuần", sell.name],
        ["Top mua ròng 1 tuần", buy.name],
    ]
    source_fill = PatternFill("solid", fgColor="17365D")
    source_border_side = Side(style="thin", color="D9E2F3")
    source_border = Border(
        left=source_border_side,
        right=source_border_side,
        top=source_border_side,
        bottom=source_border_side,
    )
    for row_index, row_values in enumerate(source_rows, start=1):
        for column_index, value in enumerate(row_values, start=1):
            cell = source_sheet.cell(row=row_index, column=column_index, value=value)
            cell.font = Font(
                name=REPORT_FONT,
                size=10,
                bold=row_index == 1,
                color="FFFFFF" if row_index == 1 else "1F2937",
            )
            if row_index == 1:
                cell.fill = source_fill
            else:
                cell.border = source_border
    source_sheet.column_dimensions["A"].width = 25
    source_sheet.column_dimensions["B"].width = 78

    output.parent.mkdir(parents=True, exist_ok=True)
    temporary_output = output.with_name(f".{output.stem}.{os.getpid()}.publishing{output.suffix}")
    try:
        workbook.save(temporary_output)
        os.replace(temporary_output, output)
    finally:
        workbook.close()
        if temporary_output.exists():
            temporary_output.unlink()
    log(f"Đã tạo dữ liệu tự doanh: {output.relative_to(ROOT)}")
    return output


def resolve_vnindex_image(config: dict[str, Any]) -> Path | None:
    preferred = ROOT / config.get("vnindex_chart_image", "incoming/vnindex_chart.png")
    if preferred.exists():
        return preferred
    incoming = ROOT / config["incoming_dir"]
    if not incoming.exists():
        return None
    candidates = [
        path
        for path in incoming.iterdir()
        if path.is_file() and path.suffix.lower() == ".png" and not path.name.startswith(".")
    ]
    return max(candidates, key=lambda path: path.stat().st_mtime) if candidates else None


def resolve_sources(config: dict[str, Any], as_of: dt.date | None = None) -> dict[str, Path | None]:
    return {
        name: latest_file(patterns, as_of)
        for name, patterns in config["source_patterns"].items()
    }


def validate_sources(
    sources: dict[str, Path | None], as_of: dt.date, config: dict[str, Any]
) -> tuple[list[str], list[str]]:
    """Return blocking errors and non-blocking warnings for selected source files."""
    errors: list[str] = []
    warnings: list[str] = []
    critical = set(config.get("critical_sources", sources))
    max_age = int(config.get("source_max_age_days", 7))
    for name, path in sources.items():
        if path is None or not path.is_file():
            message = f"Thiếu nguồn dữ liệu bắt buộc: {name}."
            (errors if name in critical else warnings).append(message)
            continue
        file_date = date_from_filename(path)
        if file_date is None:
            warnings.append(f"Không xác định được ngày dữ liệu từ tên tệp {name}: {path.name}.")
            continue
        age = (as_of - file_date).days
        if age < 0:
            errors.append(f"Nguồn {name} có ngày tương lai {file_date:%d/%m/%Y}: {path.name}.")
        elif age > max_age:
            message = f"Nguồn {name} đã cũ {age} ngày (ngày {file_date:%d/%m/%Y}): {path.name}."
            (errors if name in critical else warnings).append(message)

    for left, right, label in (
        ("foreign_buy", "foreign_sell", "khối ngoại"),
    ):
        left_path, right_path = sources.get(left), sources.get(right)
        if left_path and right_path:
            left_date, right_date = date_from_filename(left_path), date_from_filename(right_path)
            if left_date and right_date and left_date != right_date:
                errors.append(
                    f"Nguồn mua/bán {label} không cùng kỳ: {left_path.name} / {right_path.name}."
                )
    return errors, warnings


def friday_on_or_before(day: dt.date) -> dt.date:
    return day - dt.timedelta(days=(day.weekday() - 4) % 7)


def infer_as_of(sources: dict[str, Path | None], override: str | None) -> dt.date:
    if override:
        return dt.date.fromisoformat(override)
    candidates: list[dt.date] = [friday_on_or_before(dt.date.today())]
    incoming_dir = ROOT / "incoming"
    if incoming_dir.exists():
        for path in incoming_dir.iterdir():
            if path.is_file() and not path.name.startswith("."):
                day = date_from_filename(path)
                if day:
                    candidates.append(day)
    for path in sources.values():
        if path and path.is_file():
            day = date_from_filename(path)
            if day:
                candidates.append(day)
    return max(candidates)


def parse_date(value: str) -> dt.date | None:
    parts = str(value).strip().split()
    if not parts:
        return None
    text = parts[0]
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%Y/%m/%d"):
        try:
            return dt.datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def display_date(value: str) -> str:
    parsed = parse_date(value)
    return parsed.strftime("%d/%m/%Y") if parsed else str(value).strip()


def normalize_whitespace(value: object) -> str:
    return re.sub(r"\s+", " ", "" if value is None else str(value)).strip()


def parse_localized_number(value: object) -> float | None:
    text = normalize_whitespace(value)
    if not text:
        return None
    try:
        return float(text.replace(",", "."))
    except ValueError:
        return None


def format_vi_integer(value: object) -> str:
    number = parse_localized_number(value)
    if number is None:
        return normalize_whitespace(value)
    return f"{round(number):,}".replace(",", ".")


def format_vi_percent(value: object, decimal_places: int = 1) -> str:
    number = parse_localized_number(value)
    if number is None:
        return normalize_whitespace(value)
    return f"{number * 100:.{decimal_places}f}%".replace(".", ",")


def pad_rows(rows: list[list[str]], row_count: int, column_count: int) -> list[list[str]]:
    normalized = [
        [normalize_whitespace(value) for value in row[:column_count]]
        + [""] * max(0, column_count - len(row))
        for row in rows[:row_count]
    ]
    while len(normalized) < row_count:
        normalized.append([""] * column_count)
    return normalized


def applescript_string(value: object) -> str:
    text = normalize_whitespace(value)
    text = text.replace("\\", "\\\\").replace('"', '\\"')
    return f'"{text}"'


def applescript_matrix(matrix: list[list[str]]) -> str:
    rows = ["{" + ", ".join(applescript_string(cell) for cell in row) + "}" for row in matrix]
    return "{" + ", ".join(rows) + "}"


def make_table_operation(
    slide: int,
    shape_name: str,
    matrix: list[list[str]],
    start_row: int,
    start_column: int = 1,
    font_colors: list[list[str]] | None = None,
) -> dict[str, Any]:
    operation = {
        "kind": "table",
        "slide": slide,
        "shape": shape_name,
        "matrix": [[normalize_whitespace(cell) for cell in row] for row in matrix],
        "start_row": start_row,
        "start_column": start_column,
    }
    if font_colors is not None:
        operation["font_colors"] = font_colors
    return operation


def portfolio_font_colors(matrix: list[list[str]]) -> list[list[str]]:
    percentage_columns = {3, 4, 5, 6, 7, 10, 11}
    colors: list[list[str]] = []
    for row in matrix:
        color_row: list[str] = []
        for index, value in enumerate(row):
            if index not in percentage_columns or not str(value).strip().endswith("%"):
                color_row.append("")
                continue
            number = parse_localized_number(str(value).strip().rstrip("%"))
            if number is None:
                color_row.append("")
            elif number > 0:
                color_row.append("00B050")
            elif number < 0:
                color_row.append("C00000")
            else:
                color_row.append("FFC000")
        colors.append(color_row)
    return colors


def make_paragraph_operation(slide: int, shape: str, paragraph: int, text: str) -> dict[str, Any]:
    return {
        "kind": "paragraph",
        "slide": slide,
        "shape": shape,
        "paragraph": paragraph,
        "text": normalize_whitespace(text),
    }


def make_text_operation(slide: int, shape: str, text: str) -> dict[str, Any]:
    return {
        "kind": "text",
        "slide": slide,
        "shape": shape,
        "text": normalize_whitespace(text),
    }


def make_paragraph_list_operation(slide: int, shape: str, paragraphs: list[str]) -> dict[str, Any]:
    return {
        "kind": "paragraph_list",
        "slide": slide,
        "shape": shape,
        "paragraphs": [normalize_whitespace(text) for text in paragraphs],
    }


def powerpoint_operation_to_applescript(operation: dict[str, Any]) -> str:
    kind = operation["kind"]
    if kind == "table":
        return "\n".join(
            [
                f"set targetTable to table object of shape {applescript_string(operation['shape'])} "
                f"of slide {operation['slide']} of pres",
                f"my writeTable(targetTable, {applescript_matrix(operation['matrix'])}, "
                f"{operation['start_row']}, {operation['start_column']})",
            ]
        )
    if kind == "paragraph":
        return (
            f"my setParagraphText(shape {applescript_string(operation['shape'])} "
            f"of slide {operation['slide']} of pres, {operation['paragraph']}, "
            f"{applescript_string(operation['text'])})"
        )
    if kind == "text":
        return (
            f"my setShapeText(shape {applescript_string(operation['shape'])} "
            f"of slide {operation['slide']} of pres, {applescript_string(operation['text'])})"
        )
    if kind == "paragraph_list":
        joined = " & return & ".join(applescript_string(text) for text in operation["paragraphs"])
        return (
            f"my setShapeText(shape {applescript_string(operation['shape'])} "
            f"of slide {operation['slide']} of pres, {joined})"
        )
    raise ValueError(f"Unsupported PowerPoint operation: {kind}")


def build_powerpoint_operations(
    as_of: dt.date,
    table_operations: list[dict[str, Any]],
    commentary: dict[str, Any] | None = None,
    include_dates: bool = True,
) -> list[dict[str, Any]]:
    week_start = as_of - dt.timedelta(days=as_of.weekday())
    next_start = as_of + dt.timedelta(days=(7 - as_of.weekday()))
    next_end = next_start + dt.timedelta(days=4)
    data_range = f"{week_start:%d/%m} – {as_of:%d/%m/%Y}"
    next_range = f"{next_start:%d/%m/%Y} – {next_end:%d/%m/%Y}"

    # Slides 5 and 6 must contain no generated commentary. Clear the legacy
    # commentary inherited from the reference deck in both draft and finalize.
    operations: list[dict[str, Any]] = [
        make_paragraph_operation(5, "Text Placeholder 7", 3, ""),
        make_paragraph_operation(6, "Text Placeholder 7", 16, ""),
        make_paragraph_operation(2, "TextBox 5", 11, ""),
    ]
    if include_dates:
        operations.extend(
            [
                make_paragraph_operation(1, "TextBox 4", 2, f"NHẬN ĐỊNH THỊ TRƯỜNG TUẦN {next_range}"),
                make_paragraph_operation(2, "TextBox 5", 1, f"DIỄN BIẾN THỊ TRƯỜNG TUẦN {data_range}"),
                make_paragraph_operation(2, "TextBox 5", 6, f"CÁC SỰ KIỆN DIỄN RA TRONG TUẦN {next_range}"),
                make_paragraph_operation(2, "TextBox 5", 8, f"NHẬN ĐỊNH XU HƯỚNG THỊ TRƯỜNG TUẦN {next_range}"),
                make_text_operation(3, "Title 5", f"DIỄN BIẾN THỊ TRƯỜNG TUẦN {data_range}"),
                make_text_operation(4, "Title 5", f"DIỄN BIẾN THỊ TRƯỜNG TUẦN {data_range}"),
                make_text_operation(5, "Title 5", f"DIỄN BIẾN THỊ TRƯỜNG TUẦN {data_range}"),
                make_text_operation(6, "Title 5", f"DIỄN BIẾN DANH MỤC THEO DÕI TUẦN {data_range}"),
                make_text_operation(7, "Title 5", f"CÁC SỰ KIỆN DIỄN RA TRONG TUẦN {next_range}"),
                make_text_operation(8, "Title 5", f"NHẬN ĐỊNH THỊ TRƯỜNG TUẦN {next_range}"),
            ]
        )

    if commentary:
        operations.extend(
            [
                make_paragraph_operation(3, "Text Placeholder 7", 3, commentary["slide3"][0]),
                make_paragraph_operation(3, "Text Placeholder 7", 4, commentary["slide3"][1]),
                make_paragraph_operation(3, "Text Placeholder 7", 5, commentary["slide3"][2]),
                make_paragraph_operation(4, "Text Placeholder 7", 3, commentary["slide4"][0]),
                make_paragraph_operation(4, "Text Placeholder 7", 4, commentary["slide4"][1]),
                make_paragraph_list_operation(8, "Text Placeholder 7", commentary["slide8"]),
                # Remove legacy commentary paragraphs inherited from the source
                # deck after the current paragraphs have been written.
                make_paragraph_operation(3, "Text Placeholder 7", 8, ""),
                make_paragraph_operation(3, "Text Placeholder 7", 7, ""),
                make_paragraph_operation(4, "Text Placeholder 7", 6, ""),
            ]
        )
    operations.extend(table_operations)
    return operations


def generate_powerpoint_script(
    ppt_path: Path,
    as_of: dt.date,
    table_operations: list[dict[str, Any]],
    commentary: dict[str, Any] | None = None,
    include_dates: bool = True,
) -> str:
    operations = build_powerpoint_operations(as_of, table_operations, commentary, include_dates)
    body = "\n\n        ".join(powerpoint_operation_to_applescript(operation) for operation in operations)
    return f'''on writeTable(targetTable, tableData, startRow, startColumn)
    tell application "Microsoft PowerPoint"
        repeat with rowOffset from 1 to count of tableData
            set rowData to item rowOffset of tableData
            repeat with columnOffset from 1 to count of rowData
                set targetCell to get cell from targetTable row (startRow + rowOffset - 1) column (startColumn + columnOffset - 1)
                set targetRange to text range of text frame of shape of targetCell
                set content of targetRange to item columnOffset of rowData
                set font name of font of targetRange to "Arial"
            end repeat
        end repeat
    end tell
end writeTable

on setParagraphText(targetShape, paragraphNumber, newText)
    tell application "Microsoft PowerPoint"
        set targetRange to text range of text frame of targetShape
        set content of paragraph paragraphNumber of targetRange to newText
        set font name of font of paragraph paragraphNumber of targetRange to "Arial"
    end tell
end setParagraphText

on setShapeText(targetShape, newText)
    tell application "Microsoft PowerPoint"
        set targetRange to text range of text frame of targetShape
        set content of targetRange to newText
        set font name of font of targetRange to "Arial"
    end tell
end setShapeText

on run argv
    tell application "Microsoft PowerPoint"
        open {applescript_string(ppt_path.resolve())}
        set pres to active presentation
        {body}
        save pres
        close pres saving yes
    end tell
end run
'''


def apply_powerpoint_updates(
    ppt_path: Path,
    as_of: dt.date,
    table_operations: list[dict[str, Any]],
    commentary: dict[str, Any] | None,
    include_dates: bool,
    artifact_base: Path,
) -> None:
    operations = build_powerpoint_operations(as_of, table_operations, commentary, include_dates)
    if os.name == "nt":
        operations_path = artifact_base.with_suffix(".operations.json")
        write_json_atomic(operations_path, operations)
        run_windows_office(
            "UpdatePresentation",
            PresentationPath=ppt_path.resolve(),
            OperationsPath=operations_path.resolve(),
        )
        assert_presentation_fonts(ppt_path)
        return
    if sys.platform == "darwin":
        script_path = artifact_base.with_suffix(".applescript")
        write_text_atomic(
            script_path,
            generate_powerpoint_script(ppt_path, as_of, table_operations, commentary, include_dates),
        )
        result = subprocess.run(
            ["osascript", str(script_path)],
            cwd=ROOT,
            text=True,
            capture_output=True,
            timeout=300,
        )
        if result.returncode != 0:
            raise RuntimeError((result.stderr or result.stdout).strip())
        assert_presentation_fonts(ppt_path)
        return
    raise RuntimeError("Chỉ hỗ trợ tự động hóa Microsoft Office trên Windows hoặc macOS.")


def replace_excel_chart(
    workbook_path: Path,
    worksheet_name: str,
    chart_number: int,
    ppt_path: Path,
    slide_number: int,
    target_shape_name: str,
) -> None:
    if os.name == "nt":
        run_windows_office(
            "ReplaceExcelChart",
            WorkbookPath=workbook_path.resolve(),
            WorksheetName=worksheet_name,
            ChartNumber=chart_number,
            PresentationPath=ppt_path.resolve(),
            SlideNumber=slide_number,
            TargetShapeName=target_shape_name,
        )
        return
    if sys.platform == "darwin":
        run_applescript(
            "replace_excel_chart.applescript",
            workbook_path.resolve(),
            worksheet_name,
            chart_number,
            ppt_path.resolve(),
            slide_number,
            target_shape_name,
        )
        return
    raise RuntimeError("Chart replacement requires Microsoft Office on Windows or macOS.")


def replace_presentation_image(
    image_path: Path,
    ppt_path: Path,
    slide_number: int,
    target_shape_name: str,
) -> None:
    if os.name == "nt":
        run_windows_office(
            "ReplaceImage",
            ImagePath=image_path.resolve(),
            PresentationPath=ppt_path.resolve(),
            SlideNumber=slide_number,
            TargetShapeName=target_shape_name,
        )
        return
    if sys.platform == "darwin":
        run_applescript(
            "replace_image.applescript",
            image_path.resolve(),
            ppt_path.resolve(),
            slide_number,
            target_shape_name,
        )
        return
    raise RuntimeError("Image replacement requires Microsoft PowerPoint on Windows or macOS.")


def validate_commentary(payload: dict[str, Any]) -> dict[str, Any]:
    expected_lists = {"slide3": 3, "slide4": 2}
    for key, count in expected_lists.items():
        value = payload.get(key)
        if not isinstance(value, list) or len(value) != count or not all(isinstance(item, str) and item.strip() for item in value):
            raise ValueError(f"{key} must contain exactly {count} non-empty strings")
    normalized: dict[str, list[str]] = {}
    for key, values in ((key, payload[key]) for key in expected_lists):
        normalized[key] = [re.sub(r"\s+", " ", item).strip() for item in values]
        for item in normalized[key]:
            if item.startswith(("-", "•", "*")) or "\n" in item:
                raise ValueError(f"{key} must use plain prose paragraphs without bullets")
    payload["slide3"] = normalized["slide3"]
    payload["slide4"] = normalized["slide4"]
    if not payload["slide3"][0].startswith("Trong tuần "):
        raise ValueError('slide3 paragraph 1 must start with "Trong tuần "')
    if not payload["slide4"][0].startswith("Xét về các nhóm ngành,"):
        raise ValueError('slide4 paragraph 1 must start with "Xét về các nhóm ngành,"')
    length_limits = {"slide3": (80, 230), "slide4": (35, 130)}
    for key, paragraphs in (("slide3", payload["slide3"][:2]), ("slide4", payload["slide4"])):
        low, high = length_limits[key]
        for index, paragraph in enumerate(paragraphs, start=1):
            count = len(paragraph.split())
            if not low <= count <= high:
                raise ValueError(f"{key} paragraph {index} has {count} words; expected {low}-{high}")

    # Slide 5 and slide 6 deliberately retain the template's existing text.
    # Ignore legacy keys so old reviewed JSON files remain compatible without
    # ever writing those values into the presentation.
    payload.pop("slide5_summary", None)
    payload.pop("slide6_summary", None)
    # Slide 8 intentionally repeats the first two Slide 3 paragraphs.  Deriving it
    # here (instead of asking the model to rewrite it) guarantees identical wording
    # across both slides and always excludes the Vingroup contribution paragraph.
    payload["slide8"] = payload["slide3"][:2]
    return payload


def format_vi_number(value: float, places: int = 1) -> str:
    return f"{value:,.{places}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def compute_vingroup_contribution(path: Path | None, as_of: dt.date) -> dict[str, Any] | None:
    """Calculate index-point contribution from source rows instead of cached Excel formulas."""
    if not path:
        return None
    workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    sheet = workbook.active
    records: list[dict[str, Any]] = []
    for row_values in sheet.iter_rows(min_row=3, max_col=13, values_only=True):
        raw_date = row_values[0]
        day = raw_date.date() if isinstance(raw_date, dt.datetime) else raw_date
        if not isinstance(day, dt.date) or day > as_of:
            continue
        values = [row_values[column - 1] for column in (2, 3, 6, 7, 8, 9, 10, 11, 12, 13)]
        if all(isinstance(value, (int, float)) for value in values):
            index_value, market_cap = float(values[0]), float(values[1])
            stocks = [float(value) for value in values[2:]]
            prices = stocks[0::2]
            shares = stocks[1::2]
            # FiinProX valuation exports (P/E, P/B) can have the same ticker
            # layout as the required price/share export. Reject those rows rather
            # than publishing implausible index-point contributions.
            if not (500 <= index_value <= 5000 and market_cap > 100_000):
                continue
            if any(price < 1_000 or price > 5_000_000 for price in prices):
                continue
            if any(share < 1_000_000 or share > 50_000_000_000 for share in shares):
                continue
            records.append({"date": day, "index": index_value, "market_cap": market_cap, "stocks": stocks})
    workbook.close()
    records.sort(key=lambda item: item["date"])
    if len(records) < 2 or (as_of - records[-1]["date"]).days > 3:
        return None
    return compute_vingroup_metrics(records, as_of)


def compute_vingroup_metrics(records: list[dict[str, Any]], as_of: dt.date) -> dict[str, Any] | None:
    """Compute period metrics from already validated, chronological price/share records."""
    import math
    records = [dict(record) for record in sorted(records, key=lambda item: item["date"]) if record["date"] <= as_of]
    if len(records) < 2:
        return None
    if len({record["date"] for record in records}) != len(records):
        return None
    for record in records:
        values = [record["index"], record["market_cap"], *record["stocks"]]
        if len(record["stocks"]) != 8 or any(
            isinstance(value, bool) or not isinstance(value, (int, float))
            or not math.isfinite(value) or value <= 0 for value in values
        ):
            return None
    for current, previous in zip(records[1:], records[:-1]):
        cap_change = sum(
            float(current["stocks"][offset]) * float(current["stocks"][offset + 1])
            - float(previous["stocks"][offset]) * float(previous["stocks"][offset + 1])
            for offset in range(0, 8, 2)
        ) / 1_000_000_000
        daily_contribution = cap_change / previous["market_cap"] * previous["index"]
        index_change = current["index"] - previous["index"]
        if abs(daily_contribution) > 150 or abs(daily_contribution) > abs(index_change) * 4 + 50:
            return None
        current["daily_contribution"] = daily_contribution
    records[0]["daily_contribution"] = 0.0

    def period(label: str, start: dt.date) -> dict[str, Any]:
        selected = [item for item in records if start <= item["date"] <= as_of]
        if not selected:
            return {"label": label, "from": None, "to": None, "group_points": 0.0, "vnindex_points": 0.0}
        baseline = max((item for item in records if item["date"] < selected[0]["date"]), key=lambda item: item["date"], default=None)
        return {
            "label": label,
            "from": selected[0]["date"].isoformat(),
            "to": selected[-1]["date"].isoformat(),
            "group_points": sum(item["daily_contribution"] for item in selected),
            "vnindex_points": selected[-1]["index"] - (baseline or selected[0])["index"],
        }

    return {
        "method": "sum_daily_market_cap_contribution",
        "tickers": ["VIC", "VHM", "VRE", "VPL"],
        "year": period("từ đầu năm", dt.date(as_of.year, 1, 1)),
        "month": period("từ đầu tháng", as_of.replace(day=1)),
        "week": period("trong tuần", as_of - dt.timedelta(days=as_of.weekday())),
    }


def build_vingroup_commentary(metrics: dict[str, Any] | None) -> str:
    if not metrics:
        return "Xét về đóng góp của nhóm Vingroup, dữ liệu nguồn kỳ này chưa đủ để tính chính xác mức đóng góp vào VNINDEX theo từng giai đoạn. Vì vậy, báo cáo không đưa ra ước tính thay thế nhằm tránh sai lệch số liệu. Diễn biến của chỉ số cần được đánh giá cùng mức độ tập trung khi dữ liệu được bổ sung đầy đủ."

    def signed(value: float) -> str:
        return f"{'tăng' if value >= 0 else 'giảm'} {format_vi_number(abs(value))} điểm"

    year, month, week = metrics["year"], metrics["month"], metrics["week"]
    return (
        f"Xét về đóng góp của nhóm Vingroup, từ đầu năm nhóm VIC, VHM, VRE và VPL đóng góp {signed(year['group_points'])} "
        f"trong khi VNINDEX {signed(year['vnindex_points'])}. Từ đầu tháng, nhóm đóng góp {signed(month['group_points'])}, "
        f"so với mức {signed(month['vnindex_points'])} của VNINDEX. Riêng trong tuần, nhóm đóng góp {signed(week['group_points'])}, "
        f"trong khi VNINDEX {signed(week['vnindex_points'])}. Chênh lệch giữa đóng góp của nhóm và biến động toàn chỉ số cho thấy mức độ tập trung của diễn biến thị trường trong từng giai đoạn."
    )


def sum_net_values(path: Path | None) -> float | None:
    """Sum the net-trading-value column from a weekly FiinProX ranking export."""
    if not path:
        return None
    workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    sheet = workbook.active
    total = 0.0
    found = False
    for values in sheet.iter_rows(min_row=6, min_col=6, max_col=6, values_only=True):
        value = values[0]
        if isinstance(value, (int, float)):
            total += float(value)
            found = True
    workbook.close()
    return total if found else None


def compute_weekly_net_values(
    sources: dict[str, Path | None], config: dict[str, Any], as_of: dt.date
) -> dict[str, Any]:
    proprietary_buy = latest_file(config.get("proprietary_raw_patterns", {}).get("buy", []), as_of)
    proprietary_sell = latest_file(config.get("proprietary_raw_patterns", {}).get("sell", []), as_of)

    def combine(*values: float | None) -> float | None:
        return sum(value for value in values if value is not None) if any(value is not None for value in values) else None

    return {
        "foreign_vnd": combine(sum_net_values(sources.get("foreign_buy")), sum_net_values(sources.get("foreign_sell"))),
        "proprietary_vnd": combine(sum_net_values(proprietary_buy), sum_net_values(proprietary_sell)),
        "foreign_sources": [str(path) for path in (sources.get("foreign_buy"), sources.get("foreign_sell")) if path],
        "proprietary_sources": [str(path) for path in (proprietary_buy, proprietary_sell) if path],
    }


def describe_net_value(value: float | None) -> str:
    if value is None:
        return "KHÔNG TÍNH ĐƯỢC (thiếu dữ liệu)"
    action = "mua ròng" if value >= 0 else "bán ròng"
    return f"{action} {format_vi_number(abs(value) / 1_000_000_000)} tỷ đồng"


def copy_to_archive(as_of: dt.date, config: dict[str, Any]) -> Path:
    archive_root = ROOT / config["archive_dir"] / as_of.isoformat() / "before_process"
    for folder in config["feature_folders"]:
        source_dir = ROOT / folder
        if not source_dir.exists():
            continue
        for source in source_dir.rglob("*"):
            if not source.is_file() or source.name.startswith("~$"):
                continue
            relative = source.relative_to(ROOT)
            target = archive_root / relative
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(source, target)
    return archive_root


def inspect_company_export(path: Path, config: dict[str, Any]) -> list[str]:
    try:
        workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
        sheet = workbook.active
        tickers = {
            str(sheet.cell(row=8, column=column).value).strip().upper()
            for column in range(1, sheet.max_column + 1)
            if sheet.cell(row=8, column=column).value
        }
        workbook.close()
    except Exception:
        return []

    vingroup = set(config["routing_ticker_sets"]["vingroup"])
    banks = set(config["routing_ticker_sets"]["banks"])
    destinations: list[str] = []
    if len(tickers & vingroup) >= 3:
        destinations.append("dong gop cua vingroup")
    if len(tickers & banks) >= 3:
        destinations.append("nn ban rong")
    if len(tickers) >= 10:
        destinations.append("danh muc mbs")
    return destinations


def route_inputs(config: dict[str, Any]) -> list[str]:
    incoming = ROOT / config["incoming_dir"]
    incoming.mkdir(parents=True, exist_ok=True)
    messages: list[str] = []
    for source in sorted(incoming.iterdir()):
        if not source.is_file() or source.name.startswith((".", "~$")):
            continue
        if source.name == "README.txt":
            continue
        lower = source.name.lower()
        destinations: list[str] = []

        if "oneday" in lower or "oneweek" in lower or "onemonth" in lower:
            destinations = ["nganh"]
        elif source.suffix.lower() == ".png":
            messages.append(f"Giữ ảnh {source.name} trong incoming để chèn vào biểu đồ VNINDEX.")
            continue
        elif (
            ("to chuc trong nuoc" in lower or "tổ chức trong nước" in lower)
            and ("top gia tri rong" in lower or "top giá trị ròng" in lower)
            and source.suffix.lower() == ".xlsx"
        ):
            destinations = ["gd tu doanh"]
        elif ("top gia tri rong" in lower or "top giá trị ròng" in lower) and source.suffix.lower() == ".xlsx":
            destinations = ["gd nuoc ngoai"]
        elif lower.startswith("room") and source.suffix.lower() == ".pdf":
            destinations = ["gd nuoc ngoai"]
        elif "chi_so_&_nganh" in lower or ("chi_so" in lower and "nganh" in lower):
            destinations = ["thanh khoan tt", "dong gop cua vingroup"]
        elif "de_doanh_nghiep" in lower and "du_lieu_giao_dich" not in lower:
            destinations = ["danh muc mbs"]
        elif "du_lieu_giao_dich_doanh_nghiep" in lower and source.suffix.lower() == ".xlsx":
            destinations = inspect_company_export(source, config)
        elif source.suffix.lower() == ".pdf":
            messages.append(
                f"Skipped PDF {source.name}: gd tu doanh now uses only the two FiinProX Top 10 Excel files."
            )
            continue
        else:
            messages.append(f"Unclassified input: {source.name}")
            continue

        if not destinations:
            messages.append(f"Ambiguous company export (not copied): {source.name}")
            continue

        for destination in destinations:
            target_dir = ROOT / destination
            target_dir.mkdir(parents=True, exist_ok=True)
            target = target_dir / source.name
            shutil.copy2(source, target)
            messages.append(f"Copied {source.name} -> {destination}/")
    return messages


def run_processor(as_of: dt.date, config: dict[str, Any], excluded_modules: str = "", skip_failed: bool = False) -> Path:
    archive = copy_to_archive(as_of, config)
    log(f"Archived current feature files to {archive.relative_to(ROOT)}")
    command = [sys.executable, str(ROOT / config["processor"]), "all"]
    if excluded_modules.strip():
        command.append(excluded_modules.strip())
    if skip_failed:
        command.append("--skip-failed")
    processor_env = os.environ.copy()
    processor_env["PYTHONUTF8"] = "1"
    processor_env["PYTHONIOENCODING"] = "utf-8"
    processor_env["REPORT_AS_OF"] = as_of.isoformat()
    process = subprocess.run(command, cwd=ROOT, text=True, env=processor_env)
    if process.returncode != 0:
        raise RuntimeError(f"{config['processor']} exited with status {process.returncode}")
    return archive


def read_news_rows(path: Path | None) -> list[list[str]]:
    if not path:
        return []
    try:
        matrix = extract_excel_range(path, "Tin_Chinh_Thong_Loc", 2, 1, 20, 9)
    except Exception:
        matrix = extract_excel_range(path, "Tin_Chinh_Thong_Toan_Bo", 2, 1, 20, 9)
    if not any(normalize_whitespace(row[0]) and normalize_whitespace(row[4]) for row in matrix):
        matrix = extract_excel_range(path, "Tin_Chinh_Thong_Toan_Bo", 2, 1, 20, 9)
    rows: list[list[str]] = []
    for row in matrix:
        if not normalize_whitespace(row[0]) or not normalize_whitespace(row[4]):
            continue
        rows.append([display_date(row[1]), row[0], row[4]])
        if len(rows) == 3:
            break
    return rows


def read_portfolio_rows(path: Path) -> list[list[str]]:
    workbook = openpyxl.load_workbook(path, read_only=True, data_only=False)
    portfolio = workbook["Danh mục"]
    prices = workbook["Cập nhật giá"]
    lookup: dict[str, list[object]] = {}
    for values in prices.iter_rows(min_row=9, max_col=17, values_only=True):
        ticker = normalize_whitespace(values[1]).upper()
        if ticker:
            lookup[ticker] = list(values)

    def ratio(numerator: object, denominator: object) -> float | None:
        try:
            top = float(numerator)
            bottom = float(denominator)
            return top / bottom - 1 if bottom else None
        except (TypeError, ValueError, ZeroDivisionError):
            return None

    def vi_decimal(value: object, places: int = 1) -> str:
        try:
            return f"{float(value):.{places}f}".replace(".", ",")
        except (TypeError, ValueError):
            return ""

    rows: list[list[str]] = []
    for index_value, ticker_value in portfolio.iter_rows(
        min_row=4, max_col=2, values_only=True
    ):
        ticker = normalize_whitespace(ticker_value).upper()
        if not ticker:
            continue
        source = lookup.get(ticker)
        if not source:
            rows.append([normalize_whitespace(index_value), ticker, *([""] * 10)])
            continue
        current = source[4]
        pe = source[11]
        pb = source[12]
        metrics = [
            normalize_whitespace(index_value),
            ticker,
            format_vi_integer(current),
            format_vi_percent(ratio(current, source[5]), 2),
            format_vi_percent(ratio(current, source[6]), 2),
            format_vi_percent(ratio(current, source[7]), 2),
            format_vi_percent(ratio(current, source[13]), 2),
            format_vi_percent(ratio(current, source[14]), 2),
            vi_decimal(pe),
            vi_decimal(pb),
            format_vi_percent(ratio(pe, source[15]), 2),
            format_vi_percent(ratio(pb, source[16]), 2),
        ]
        rows.append(metrics)
    workbook.close()
    return rows


def build_tables(
    sources: dict[str, Path | None],
    as_of: dt.date,
) -> tuple[list[dict[str, Any]], dict[str, Any], list[str]]:
    operations: list[dict[str, Any]] = []
    facts: dict[str, Any] = {"as_of": as_of.isoformat()}
    warnings: list[str] = []

    if sources.get("sectors"):
        matrix = extract_excel_range(sources["sectors"], "nganh", 3, 1, 18, 4)
        matrix = pad_rows(matrix, 18, 4)
        operations.append(make_table_operation(4, "Table 1", matrix, 3))
        facts["sectors"] = matrix
    else:
        warnings.append("Missing processed sector workbook; slide 4 table was preserved.")

    if sources.get("foreign_buy"):
        display = extract_excel_range(sources["foreign_buy"], "Sheet1", 6, 1, 10, 7)
        raw = extract_excel_range(sources["foreign_buy"], "Sheet1", 6, 1, 10, 7, "raw")
        matrix = pad_rows(
            [[shown[0], shown[1], format_vi_integer(values[5]), shown[6]] for shown, values in zip(display, raw)],
            10,
            4,
        )
        operations.append(make_table_operation(5, "Table 24", matrix, 2))
        facts["foreign_top_buy"] = matrix
    else:
        warnings.append("Missing foreign-buy workbook; slide 5 foreign-buy table was preserved.")

    if sources.get("foreign_sell"):
        display = extract_excel_range(sources["foreign_sell"], "Sheet1", 6, 1, 10, 7)
        raw = extract_excel_range(sources["foreign_sell"], "Sheet1", 6, 1, 10, 7, "raw")
        matrix = pad_rows(
            [[shown[0], shown[1], format_vi_integer(values[5]), shown[6]] for shown, values in zip(display, raw)],
            10,
            4,
        )
        operations.append(make_table_operation(5, "Table 25", matrix, 2))
        facts["foreign_top_sell"] = matrix
    else:
        warnings.append("Missing foreign-sell workbook; slide 5 foreign-sell table was preserved.")

    if sources.get("proprietary"):
        sold = pad_rows(extract_excel_range(sources["proprietary"], "Top_10_Ban_Rong", 1, 1, 11, 4), 11, 4)
        bought = pad_rows(extract_excel_range(sources["proprietary"], "Top_10_Mua_Rong", 1, 1, 11, 4), 11, 4)
        operations.append(make_table_operation(5, "Table 16", sold, 1))
        operations.append(make_table_operation(5, "Table 23", bought, 1))
        facts["proprietary_top_sell"] = sold[1:]
        facts["proprietary_top_buy"] = bought[1:]
    else:
        warnings.append("Missing proprietary-trading workbook; both proprietary tables were preserved.")

    news_rows = pad_rows(read_news_rows(sources.get("company_news")), 4, 3)
    operations.append(make_table_operation(6, "Table 2", news_rows, 2))
    facts["company_events"] = [row for row in news_rows if any(row)]

    if sources.get("foreign_watchlist"):
        summary = extract_excel_range(sources["foreign_watchlist"], "Sheet2", 4, 7, 16, 5)
        month_rows = [row for row in summary if re.fullmatch(r"Tháng\s+\d{1,2}", normalize_whitespace(row[0]), re.I)]
        month_rows.sort(key=lambda row: int(re.search(r"\d+", row[0]).group()))
        ytd_rows = [row for row in summary if normalize_whitespace(row[0]).lower() == "từ đầu năm"]
        selected = month_rows[-3:] + (ytd_rows[:1] if ytd_rows else [["Từ đầu năm", "", "", "", ""]])
        room_raw = extract_excel_range(sources["foreign_watchlist"], "Sheet1", 6, 7, 1, 16, "raw")
        room = room_raw[0] if room_raw else [""] * 16
        room_values = [format_vi_percent(room[index]) if index < len(room) else "" for index in (0, 5, 10, 15)]
        selected.append(["Room NN còn lại", *room_values])
        matrix = pad_rows(selected, 5, 5)
        operations.append(make_table_operation(6, "Table 6", matrix, 2))
        facts["foreign_watchlist"] = matrix
    else:
        warnings.append("Missing watchlist foreign-flow workbook; slide 6 table was preserved.")

    if sources.get("calendar"):
        calendar = extract_excel_range(sources["calendar"], "Lịch sự kiện", 2, 1, 5, 2)
        calendar = pad_rows([[display_date(row[0]), row[1]] for row in calendar if any(row)], 5, 2)
        operations.append(make_table_operation(7, "Table 2", calendar, 2))
        facts["next_week_calendar"] = [row for row in calendar if any(row)]
    else:
        warnings.append("Missing calendar workbook; slide 7 table was preserved.")

    if sources.get("portfolio"):
        portfolio_rows = read_portfolio_rows(sources["portfolio"])
        slide9 = pad_rows(portfolio_rows[:25], 25, 12)
        slide10 = pad_rows(portfolio_rows[25:50], 25, 12)
        slide11 = pad_rows(portfolio_rows[50:70], 20, 12)
        operations.append(make_table_operation(9, "Table 3", slide9, 4, font_colors=portfolio_font_colors(slide9)))
        operations.append(make_table_operation(10, "Table 8", slide10, 4, font_colors=portfolio_font_colors(slide10)))
        operations.append(make_table_operation(11, "Table 4", slide11, 4, font_colors=portfolio_font_colors(slide11)))
        date_cell = [[as_of.strftime("%d/%m/%Y")]]
        operations.append(make_table_operation(9, "Table 3", date_cell, 2, 2))
        operations.append(make_table_operation(10, "Table 8", date_cell, 2, 2))
        operations.append(make_table_operation(11, "Table 4", date_cell, 2, 2))
        facts["portfolio"] = portfolio_rows
    else:
        warnings.append("Missing portfolio workbook; slides 9-11 were preserved.")

    facts.update(compute_market_metrics(sources, as_of))
    return operations, facts, warnings


def compute_market_metrics(sources: dict[str, Path | None], as_of: dt.date) -> dict[str, Any]:
    result: dict[str, Any] = {}
    liquidity = sources.get("liquidity")
    if liquidity:
        workbook = openpyxl.load_workbook(liquidity, read_only=True, data_only=True)
        sheet = workbook.active
        records: list[dict[str, Any]] = []
        for row in range(1, sheet.max_row + 1):
            day_value = sheet.cell(row=row, column=2).value
            if isinstance(day_value, dt.datetime):
                day = day_value.date()
            elif isinstance(day_value, dt.date):
                day = day_value
            elif isinstance(day_value, str):
                day = parse_date(day_value)
            else:
                day = None
            if not day or day > as_of:
                continue
            values = [sheet.cell(row=row, column=column).value for column in (3, 4, 5, 6)]
            if all(isinstance(value, (int, float)) for value in values):
                records.append(
                    {
                        "date": day,
                        "vnindex": float(values[0]),
                        "vnindex_value": float(values[1]),
                        "vn30": float(values[2]),
                        "vn30_value": float(values[3]),
                    }
                )
        workbook.close()
        records.sort(key=lambda item: item["date"])
        week_start = as_of - dt.timedelta(days=as_of.weekday())
        previous_start = week_start - dt.timedelta(days=7)
        current = [item for item in records if week_start <= item["date"] <= as_of]
        previous = [item for item in records if previous_start <= item["date"] < week_start]
        if len(current) >= 2:
            comparison_base = max(
                (item for item in records if item["date"] < current[0]["date"]),
                key=lambda item: item["date"],
                default=current[0],
            )
            current_value = sum(item["vnindex_value"] for item in current)
            previous_value = sum(item["vnindex_value"] for item in previous) if previous else None
            result["market_week"] = {
                "from": current[0]["date"].isoformat(),
                "to": current[-1]["date"].isoformat(),
                "comparison_from": comparison_base["date"].isoformat(),
                "vnindex_start": comparison_base["vnindex"],
                "vnindex_end": current[-1]["vnindex"],
                "vnindex_change_points": current[-1]["vnindex"] - comparison_base["vnindex"],
                "vnindex_change_pct": current[-1]["vnindex"] / comparison_base["vnindex"] - 1,
                "vn30_start": comparison_base["vn30"],
                "vn30_end": current[-1]["vn30"],
                "vn30_change_points": current[-1]["vn30"] - comparison_base["vn30"],
                "vn30_change_pct": current[-1]["vn30"] / comparison_base["vn30"] - 1,
                "vnindex_matched_value_vnd": current_value,
                "vnindex_liquidity_change_pct": (current_value / previous_value - 1) if previous_value else None,
                "daily": [
                    {**item, "date": item["date"].isoformat()}
                    for item in current
                ],
            }
    return result


def write_review_package(
    run_dir: Path,
    facts: dict[str, Any],
    as_of: dt.date,
    config: dict[str, Any],
) -> tuple[Path, Path | None]:
    review_dir = run_dir / "review"
    review_dir.mkdir(parents=True, exist_ok=True)
    facts_path = review_dir / "analysis_data.json"
    write_json_atomic(facts_path, facts)

    context_path = ROOT / config["market_context_file"]
    context = context_path.read_text(encoding="utf-8").strip() if context_path.exists() else ""
    reference_path = ROOT / config.get("commentary_reference_file", "")
    reference = json.loads(reference_path.read_text(encoding="utf-8")) if reference_path.is_file() else {}
    prompt = f"""Vai trò: Chuyên viên chiến lược thị trường chứng khoán Việt Nam tại MBS.

Mục tiêu: Viết nhận định ngắn gọn cho báo cáo tuần kết thúc ngày {as_of:%d/%m/%Y}, chỉ sử dụng dữ liệu được cung cấp bên dưới.

Tiêu chuẩn:
- Không phát minh số liệu, tin vĩ mô, độ rộng thị trường hoặc nguyên nhân nếu dữ liệu không có.
- Phân biệt rõ dữ kiện và suy luận; giọng văn trung tính, chuyên nghiệp, không quảng cáo.
- Số liệu dùng định dạng Việt Nam; nêu đơn vị; làm tròn vừa đủ.
- Lấy commentary mẫu ở cuối prompt làm chuẩn bắt buộc về cấu trúc, thứ tự luận điểm, cách nối ý, giọng văn và mức độ chi tiết. Tuyệt đối không sao chép ngày, số liệu, sự kiện, dự báo hoặc kết luận của kỳ mẫu.
- Chỉ viết commentary cho Slide 3 và Slide 4. Tuyệt đối không tạo nội dung cho Slide 5 hoặc Slide 6.
- Nội dung phải vừa các hộp văn bản hiện có. Slide 3 đoạn 1 dài 125-180 từ, đoạn 2 dài 100-150 từ; các đoạn Slide 4 dài 45-100 từ. Viết với nhịp câu, độ cô đọng và mức độ diễn giải tương đương báo cáo mẫu ngày 24/08/2026.
- Dùng thống nhất các thuật ngữ: "VNINDEX", "VN30", "thanh khoản khớp lệnh", "tăng/giảm so với tuần trước". Không đổi sang từ đồng nghĩa chỉ để làm mới câu chữ.
- Mỗi đoạn là một khối văn xuôi; không thêm tiêu đề, bullet, ký hiệu đầu dòng, câu hỏi tu từ, lời khuyên mua/bán hoặc ngôi thứ nhất.
- Câu đầu nêu dữ kiện hoặc chủ đề chính, các câu giữa đưa số liệu và diễn giải, câu cuối chốt hàm ý hoặc triển vọng. Không lặp cùng một số liệu ở nhiều đoạn nếu không cần thiết.

Quy cách bắt buộc cho Slide 3:
- Đoạn 1 luôn mở bằng "Trong tuần [khoảng ngày], ..." và gồm 5-7 câu theo đúng thứ tự của mẫu: (1) VNINDEX: số điểm tăng/giảm, tỷ lệ, mức đóng cửa và thanh khoản; (2) VN30 với cùng hệ quy chiếu; (3) diễn biến đáng chú ý trong các phiên; (4) đánh giá vai trò nhóm vốn hóa lớn, độ lan tỏa và dòng tiền; (5) tín hiệu cuối tuần và hàm ý cần theo dõi. Chỉ nêu độ rộng khi dữ liệu có.
- Đoạn 2 đi thẳng vào yếu tố vĩ mô, quốc tế, hàng hóa hoặc địa chính trị quan trọng nhất giống cách mở trực tiếp bằng dầu Brent trong mẫu. Sắp xếp các sự kiện theo quan hệ nguyên nhân–tác động, sau đó nêu lịch dữ liệu tuần tới và kết luận bằng triển vọng VNINDEX nếu context hỗ trợ. Chỉ được nêu dự báo điểm số, hỗ trợ/kháng cự khi context do người dùng cung cấp có chính xác nội dung đó. Nếu context trống, mở bằng "Về bối cảnh thị trường, ...", nói rõ chưa đủ dữ liệu vĩ mô và chỉ đánh giá từ tín hiệu giao dịch.
- Đoạn 3 phải là một chuỗi không rỗng nhưng hệ thống sẽ tự thay bằng đoạn Vingroup tính toán theo mẫu cố định khi lưu và khi xuất báo cáo. Không tự suy luận hoặc sao chép bất kỳ số Vingroup nào từ commentary mẫu.
- Cách viết số: điểm và nghìn tỷ đồng dùng tối đa một chữ số thập phân; tỷ lệ phần trăm dùng một chữ số thập phân, trừ khi số liệu nguồn cần hai chữ số để không bị hiểu thành 0. Dùng dấu phân cách theo chuẩn Việt Nam.
- Slide 8 được hệ thống tự động sao chép nguyên văn hai đoạn đầu của Slide 3; đoạn Vingroup thứ ba bị loại bỏ. Không tạo nội dung Slide 8 riêng.
- Slide 4 đoạn 1 luôn mở bằng "Xét về các nhóm ngành, ...", nêu số nhóm tăng/giảm, hai nhóm dẫn đầu và nhóm kém tích cực nhất theo đúng nhịp của mẫu. Đoạn 2 đánh giá độ lan tỏa, đối chiếu diễn biến một tuần với một tháng và kết luận đây là xu hướng hay nhịp phục hồi; không nêu cổ phiếu riêng nếu dữ liệu không có.

Trả về JSON thuần, đúng schema sau, không dùng markdown:
{{
  "slide3": ["...", "...", "..."],
  "slide4": ["...", "..."]
}}

Context do người dùng cung cấp (có thể trống):
{context or '[trống]'}

Dữ liệu máy tính đã tổng hợp:
{json.dumps(facts, ensure_ascii=False, indent=2)}

Commentary mẫu tham chiếu từ file {reference.get('source', '[không có]')}:
LƯU Ý: chỉ học cấu trúc và văn phong; mọi ngày, số liệu và nhận định trong mẫu đều thuộc kỳ cũ và bị cấm sao chép.
{json.dumps(reference, ensure_ascii=False, indent=2)}
"""
    prompt_path = review_dir / "review_prompt.txt"
    write_text_atomic(prompt_path, prompt)

    commentary_path: Path | None = None
    load_dotenv(ROOT / ".env")
    if os.environ.get("OPENAI_API_KEY"):
        try:
            commentary = call_openai(prompt, config)
            commentary = validate_commentary(commentary)
            commentary["slide3"][2] = build_vingroup_commentary(facts.get("vingroup_contribution"))
            commentary_path = review_dir / "commentary.json"
            write_json_atomic(commentary_path, commentary)
        except Exception as exc:
            write_text_atomic(review_dir / "openai_error.txt", str(exc))
            log(f"OpenAI draft was not generated: {exc}")
    else:
        example = {
            "slide3": ["", "", ""],
            "slide4": ["", ""],
        }
        write_json_atomic(review_dir / "commentary.template.json", example)
    return prompt_path, commentary_path


def call_openai(prompt: str, config: dict[str, Any]) -> dict[str, Any]:
    schema = {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "slide3": {"type": "array", "minItems": 3, "maxItems": 3, "items": {"type": "string"}},
            "slide4": {"type": "array", "minItems": 2, "maxItems": 2, "items": {"type": "string"}},
        },
        "required": ["slide3", "slide4"],
    }
    payload = {
        "model": os.environ.get("OPENAI_MODEL", config["openai"]["model"]),
        "input": [
            {
                "role": "user",
                "content": [{"type": "input_text", "text": prompt}],
            }
        ],
        "reasoning": {"effort": config["openai"]["reasoning_effort"]},
        "text": {
            "verbosity": "medium",
            "format": {
                "type": "json_schema",
                "name": "weekly_market_commentary",
                "strict": True,
                "schema": schema,
            },
        },
        "store": False,
    }
    request = urllib.request.Request(
        "https://api.openai.com/v1/responses",
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {os.environ['OPENAI_API_KEY']}",
            "Content-Type": "application/json",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=180) as response:
            body = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"OpenAI API returned HTTP {exc.code}: {detail[:1000]}") from exc

    for item in body.get("output", []):
        if item.get("type") != "message":
            continue
        for content in item.get("content", []):
            if content.get("type") == "output_text" and content.get("text"):
                return json.loads(content["text"])
    raise RuntimeError("OpenAI response did not contain output_text")


def prepare_report(as_of: dt.date, config: dict[str, Any], force: bool = False) -> dict[str, Any]:
    ensure_proprietary_workbook(as_of, config)
    sources = resolve_sources(config, as_of)
    source_errors, source_warnings = validate_sources(sources, as_of, config)
    run_dir = ROOT / config["output_dir"] / as_of.isoformat()
    run_dir.mkdir(parents=True, exist_ok=True)
    next_monday = as_of + dt.timedelta(days=(7 - as_of.weekday()))
    stem = f"MBS Dau Tu - BC Thi truong Tuan - {next_monday:%d.%m.%Y}"
    draft_pptx = run_dir / f"{stem} - DRAFT.pptx"
    final_pptx = run_dir / f"{stem}.pptx"
    final_pdf = run_dir / f"{stem}.pdf"
    template = ROOT / config["template_pptx"]

    if draft_pptx.exists() and not force:
        raise FileExistsError(f"Draft already exists: {draft_pptx}. Use --force to rebuild it.")
    shutil.copy2(template, draft_pptx)

    operations, facts, warnings = build_tables(sources, as_of)
    warnings = source_warnings + warnings
    facts["vingroup_contribution"] = compute_vingroup_contribution(sources.get("vingroup"), as_of)
    if sources.get("vingroup") and facts["vingroup_contribution"] is None:
        source_errors.append(
            "Dữ liệu Vingroup không đạt kiểm tra hợp lý (giá/cổ phiếu hoặc mức đóng góp bất thường)."
        )
    facts["weekly_net_values"] = compute_weekly_net_values(sources, config, as_of)
    apply_powerpoint_updates(
        draft_pptx,
        as_of,
        operations,
        commentary=None,
        include_dates=True,
        artifact_base=run_dir / "update_draft",
    )

    if sources.get("vingroup"):
        try:
            replace_excel_chart(
                sources["vingroup"].resolve(),
                "Sheet1",
                1,
                draft_pptx.resolve(),
                3,
                "Chart 6",
            )
        except Exception as exc:
            warnings.append(f"Vingroup chart was preserved because replacement failed: {exc}")
    else:
        warnings.append("Missing Vingroup workbook; slide 3 lower chart was preserved.")

    chart_image = resolve_vnindex_image(config)
    if chart_image:
        try:
            replace_presentation_image(
                chart_image.resolve(),
                draft_pptx.resolve(),
                3,
                "Picture 1",
            )
        except Exception as exc:
            warnings.append(f"VNINDEX screenshot was preserved because replacement failed: {exc}")
    else:
        warnings.append("Không tìm thấy ảnh PNG trong incoming; ảnh VNINDEX cũ được giữ nguyên.")

    prompt_path, commentary_path = write_review_package(run_dir, facts, as_of, config)
    state = {
        "as_of": as_of.isoformat(),
        "draft_pptx": str(draft_pptx),
        "final_pptx": str(final_pptx),
        "final_pdf": str(final_pdf),
        "review_prompt": str(prompt_path),
        "commentary": str(commentary_path) if commentary_path else None,
        "warnings": warnings,
        "data_quality": {"errors": source_errors, "warnings": source_warnings},
        "sources": {name: str(path) if path else None for name, path in sources.items()},
        "weekly_net_values": facts["weekly_net_values"],
        "vnindex_chart_image": str(chart_image) if chart_image else None,
    }
    state_path = run_dir / "run_state.json"
    write_json_atomic(state_path, state)
    write_json_atomic(ROOT / config["output_dir"] / "latest_run.json", {"state": str(state_path)})
    log(f"Draft created: {draft_pptx}")
    log(f"Review prompt: {prompt_path}")
    if commentary_path:
        log(f"GPT draft created for review: {commentary_path}")
    else:
        log("No OPENAI_API_KEY found. Use review_prompt.txt in ChatGPT and save the JSON as review/commentary.json.")
    for warning in warnings:
        log(f"WARNING: {warning}")
    for error in source_errors:
        log(f"ERROR: {error}")
    log("\n[TỔNG GIÁ TRỊ RÒNG TRONG TUẦN]")
    log(f"(1) Khối nước ngoài: {describe_net_value(facts['weekly_net_values']['foreign_vnd'])}")
    log(f"(2) Tự doanh (tổ chức trong nước): {describe_net_value(facts['weekly_net_values']['proprietary_vnd'])}")
    return state


def load_latest_state(config: dict[str, Any]) -> tuple[Path, dict[str, Any]]:
    latest = ROOT / config["output_dir"] / "latest_run.json"
    if not latest.exists():
        raise FileNotFoundError("No prepared run found. Run prepare/all first.")
    pointer = json.loads(latest.read_text(encoding="utf-8"))
    state_path = Path(pointer["state"])
    state = json.loads(state_path.read_text(encoding="utf-8"))
    return state_path, state


def finalize_report(config: dict[str, Any], commentary_override: str | None, keep_existing: bool) -> dict[str, Any]:
    state_path, state = load_latest_state(config)
    draft = Path(state["draft_pptx"])
    final_pptx = Path(state["final_pptx"])
    final_pdf = Path(state["final_pdf"])
    as_of = dt.date.fromisoformat(state["as_of"])
    run_dir = state_path.parent

    if keep_existing:
        raise ValueError(
            "--keep-existing-commentary đã bị vô hiệu hóa vì bản nháp chứa nội dung kỳ cũ. "
            "Hãy duyệt review/commentary.json rồi chạy finalize."
        )
    state_sources = {
        name: Path(state.get("sources", {}).get(name)) if state.get("sources", {}).get(name) else None
        for name in config.get("source_patterns", {})
    }
    runtime_errors, _ = validate_sources(state_sources, as_of, config)
    vingroup_source = state_sources.get("vingroup")
    current_vingroup = compute_vingroup_contribution(
        vingroup_source if vingroup_source and vingroup_source.is_file() else None,
        as_of,
    )
    if vingroup_source and vingroup_source.is_file() and current_vingroup is None:
        runtime_errors.append(
            "Dữ liệu Vingroup không đạt kiểm tra hợp lý (giá/cổ phiếu hoặc mức đóng góp bất thường)."
        )
    quality_errors = list(
        dict.fromkeys(state.get("data_quality", {}).get("errors", []) + runtime_errors)
    )
    if quality_errors:
        raise ValueError("Không thể xuất bản vì lỗi chất lượng dữ liệu:\n- " + "\n- ".join(quality_errors))

    commentary_path = Path(commentary_override) if commentary_override else run_dir / "review" / "commentary.json"
    if not commentary_path.exists():
        raise FileNotFoundError(f"Commentary not found: {commentary_path}. Save the reviewed JSON there.")
    commentary = validate_commentary(json.loads(commentary_path.read_text(encoding="utf-8")))
    facts_path = run_dir / "review" / "analysis_data.json"
    facts = json.loads(facts_path.read_text(encoding="utf-8")) if facts_path.exists() else {}
    facts["vingroup_contribution"] = current_vingroup
    commentary["slide3"][2] = build_vingroup_commentary(current_vingroup)

    temporary_pptx = run_dir / f".{final_pptx.stem}.publishing.pptx"
    temporary_pdf = run_dir / f".{final_pdf.stem}.publishing.pdf"
    try:
        shutil.copy2(draft, temporary_pptx)
        apply_powerpoint_updates(
            temporary_pptx,
            as_of,
            [],
            commentary=commentary,
            include_dates=False,
            artifact_base=run_dir / "finalize_text",
        )
        export_pdf(temporary_pptx, temporary_pdf)
        if not temporary_pdf.exists() or temporary_pdf.stat().st_size < 10_000:
            raise RuntimeError("PDF xuất ra không hợp lệ hoặc trống.")
        if final_pptx.exists():
            timestamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
            backup = final_pptx.with_name(f"{final_pptx.stem}.backup_{timestamp}{final_pptx.suffix}")
            shutil.copy2(final_pptx, backup)
        os.replace(temporary_pptx, final_pptx)
        os.replace(temporary_pdf, final_pdf)
    finally:
        for temporary in (temporary_pptx, temporary_pdf):
            if temporary.exists():
                temporary.unlink()
    state["finalized_at"] = dt.datetime.now().isoformat(timespec="seconds")
    write_json_atomic(state_path, state)
    log(f"Final PPTX: {final_pptx}")
    log(f"Final PDF:  {final_pdf}")
    return state


def run_doctor(config: dict[str, Any]) -> None:
    required_modules = {
        "openpyxl": "openpyxl",
        "pandas": "pandas",
        "numpy": "numpy",
        "pdfplumber": "pdfplumber",
        "requests": "requests",
        "bs4": "beautifulsoup4",
        "reportlab": "reportlab",
    }
    if os.name == "nt":
        required_modules["truststore"] = "truststore"
    missing = [package for module, package in required_modules.items() if importlib.util.find_spec(module) is None]
    if missing:
        raise RuntimeError(
            "Thiếu thư viện Python: " + ", ".join(missing) + ". Hãy chạy 0_Cai_dat_Windows.bat."
        )

    required_files = [
        ROOT / config["template_pptx"],
        ROOT / config["processor"],
        SCRIPT_DIR / "windows_office.ps1" if os.name == "nt" else SCRIPT_DIR / "replace_image.applescript",
    ]
    missing_files = [str(path) for path in required_files if not path.exists()]
    if missing_files:
        raise FileNotFoundError("Thiếu file bắt buộc: " + "; ".join(missing_files))

    if os.name == "nt":
        office_status = run_windows_office("CheckOffice", timeout=120)
        log(f"Microsoft Office COM: {office_status}")
    elif sys.platform == "darwin":
        if not shutil.which("osascript"):
            raise FileNotFoundError("Không tìm thấy osascript trên macOS.")
        log("macOS AppleScript: OK")
    else:
        raise RuntimeError("Hệ điều hành này chưa được hỗ trợ. Hãy dùng Windows 11 hoặc macOS.")
    log(f"Python: {sys.executable}")
    log("Kiểm tra môi trường hoàn tất: OK")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Automate the weekly MBS market report.")
    subparsers = parser.add_subparsers(dest="command", required=True)

    subparsers.add_parser("doctor", help="Check Python dependencies and Microsoft Office automation.")

    route_parser = subparsers.add_parser("route", help="Copy files from incoming/ to feature folders.")
    route_parser.add_argument("--as-of")

    process_parser = subparsers.add_parser("process", help="Archive current files and run process_data.py.")
    process_parser.add_argument("--as-of")
    process_parser.add_argument("--exclude-modules", default="", help="Comma-separated module numbers 2-10 to skip.")
    process_parser.add_argument("--skip-failed", action="store_true", help="Bỏ qua các module lỗi và tiếp tục.")

    prepare_parser = subparsers.add_parser("prepare", help="Build draft deck and review package.")
    prepare_parser.add_argument("--as-of")
    prepare_parser.add_argument("--force", action="store_true")

    all_parser = subparsers.add_parser("all", help="Route inputs, process data, and build the draft.")
    all_parser.add_argument("--as-of")
    all_parser.add_argument("--skip-process", action="store_true")
    all_parser.add_argument("--force", action="store_true")
    all_parser.add_argument("--exclude-modules", default="", help="Comma-separated module numbers 2-10 to skip.")
    all_parser.add_argument("--skip-failed", action="store_true", help="Bỏ qua các module lỗi và tiếp tục tạo bản nháp.")

    finalize_parser = subparsers.add_parser("finalize", help="Apply reviewed commentary and export PPTX/PDF.")
    finalize_parser.add_argument("--commentary")
    finalize_parser.add_argument(
        "--keep-existing-commentary",
        action="store_true",
        help="Deprecated and rejected: finalization always requires reviewed commentary.",
    )
    return parser


def main() -> int:
    args = build_parser().parse_args()
    if hasattr(args, "exclude_modules"):
        parse_module_exclusions(args.exclude_modules)
    config = load_config()
    if args.command == "doctor":
        run_doctor(config)
        return 0
    sources = resolve_sources(config)
    as_of = infer_as_of(sources, getattr(args, "as_of", None))
    statuses: list[tuple[str, str]] = []

    def step(name: str, action: Any) -> Any:
        log(f"\n[ĐANG CHẠY] {name}")
        try:
            result = action()
        except Exception as exc:
            statuses.append((name, "KHÔNG ĐẠT"))
            log(f"[KHÔNG ĐẠT] {name}: {exc}")
            failed = ", ".join(item for item, status in statuses if status == "KHÔNG ĐẠT")
            log(f"[TỔNG HỢP] CÓ MODULE KHÔNG ĐẠT: {failed}")
            raise
        statuses.append((name, "THÀNH CÔNG"))
        log(f"[THÀNH CÔNG] {name}")
        return result

    if args.command in {"route", "all"}:
        messages = step("Phân loại dữ liệu đầu vào", lambda: route_inputs(config))
        for message in messages:
            log(message)

    if args.command in {"process", "all"} and not getattr(args, "skip_process", False):
        excluded_modules = getattr(args, "exclude_modules", "")
        skip_failed = getattr(args, "skip_failed", False)
        try:
            step("Xử lý các module dữ liệu", lambda: run_processor(as_of, config, excluded_modules, skip_failed=skip_failed))
        except Exception as exc:
            if skip_failed:
                log(f"[BỎ QUA] Đã bỏ qua lỗi xử lý dữ liệu theo tùy chọn --skip-failed: {exc}")
            else:
                should_skip = False
                if os.environ.get("REPORT_NONINTERACTIVE") == "1" or not sys.stdin.isatty():
                    raise
                try:
                    print("\n" + "=" * 65)
                    print(f"CẢNH BÁO: Bước xử lý dữ liệu gặp lỗi: {exc}")
                    print("=" * 65)
                    ans = input("👉 Bạn có muốn BỎ QUA lỗi này và tiếp tục tạo bản nháp với dữ liệu hiện có? (y/n) [mặc định: y]: ").strip().lower()
                    if ans in {'', 'y', 'yes', 'c', 'co', 'ok'}:
                        should_skip = True
                except (KeyboardInterrupt, EOFError):
                    should_skip = False

                if should_skip:
                    log("[TIẾP TỤC] Người dùng chọn tiếp tục tạo bản nháp.")
                else:
                    raise

        sources = resolve_sources(config)
        as_of = infer_as_of(sources, getattr(args, "as_of", None))
        try:
            step("Tạo bảng giao dịch tự doanh", lambda: ensure_proprietary_workbook(as_of, config))
        except Exception as exc:
            log(f"WARNING: Không thể tạo bảng tự doanh ({exc}); tiếp tục tạo bản nháp.")

    if args.command in {"prepare", "all"}:
        step("Tạo bản nháp và gói commentary", lambda: prepare_report(as_of, config, force=getattr(args, "force", False)))
    elif args.command == "finalize":
        step("Cập nhật commentary và xuất PowerPoint/PDF", lambda: finalize_report(config, args.commentary, args.keep_existing_commentary))
    if statuses:
        if any(status != "THÀNH CÔNG" for _, status in statuses):
            log("[TỔNG HỢP] Hoàn tất với lỗi đã bỏ qua; cần kiểm tra dữ liệu bản nháp.")
        else:
            log("[TỔNG HỢP] TẤT CẢ BƯỚC ĐÃ CHỌN ĐỀU CHẠY THÀNH CÔNG")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        log("Cancelled.")
        raise SystemExit(130)
    except Exception as exc:
        log(f"ERROR: {exc}")
        raise SystemExit(1)
