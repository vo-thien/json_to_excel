# -*- coding: utf-8 -*-
"""Logic chuyển đổi JSON -> Excel.

Module này KHÔNG import tkinter, nên có thể chạy và kiểm thử độc lập:

    python json_excel_core.py aaaa.json -o out.xlsx

Điểm cốt lõi: record trong file JSON thường KHÔNG đồng nhất (record có key này,
record khác lại không). Vì vậy cột của mỗi sheet là HỢP tất cả key theo thứ tự
xuất hiện lần đầu, record thiếu key thì để ô trống.
"""

from __future__ import annotations

import datetime
import json
import os
import re
import sys
from dataclasses import dataclass, field
from typing import Any, Callable, Sequence

import xlsxwriter

# ---------------------------------------------------------------- giới hạn Excel

EXCEL_MAX_ROWS = 1_048_576
EXCEL_MAX_COLS = 16_384
EXCEL_MAX_CELL_CHARS = 32_767
EXCEL_MAX_SHEET_NAME = 31

#: Ký tự Excel không cho phép trong tên sheet.
SHEET_BAD_CHARS = re.compile(r"[\[\]:*?/\\]")
#: Ký tự điều khiển mà xlsxwriter từ chối ghi.
CONTROL_CHARS = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")

TRUNCATION_SUFFIX = "…[đã cắt]"
KV_PATH = "(thông tin chung)"
ROOT_PATH = "(gốc)"


# ------------------------------------------------------------------- dữ liệu ra


@dataclass
class ArrayCandidate:
    """Một mảng trong JSON có thể xuất thành một sheet."""

    path: str
    rows: list
    sheet_name: str = ""
    columns: list[str] = field(default_factory=list)

    @property
    def count(self) -> int:
        return len(self.rows)


@dataclass
class SheetReport:
    name: str
    rows: int
    cols: int


@dataclass
class Report:
    out_path: str
    sheets: list[SheetReport] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    @property
    def total_rows(self) -> int:
        return sum(s.rows for s in self.sheets)

    def summary(self) -> str:
        lines = [
            f"Đã ghi {len(self.sheets)} sheet, {self.total_rows} dòng vào:\n{self.out_path}"
        ]
        for s in self.sheets:
            lines.append(f"  • {s.name}: {s.rows} dòng × {s.cols} cột")
        if self.warnings:
            lines.append(f"\nCảnh báo ({len(self.warnings)}):")
            lines += [f"  • {w}" for w in self.warnings]
        return "\n".join(lines)


# ------------------------------------------------------------------- nạp JSON


def load_json(path: str) -> Any:
    """Đọc file JSON, tự xử lý BOM và vài encoding thường gặp."""
    last_err: Exception | None = None
    for enc in ("utf-8-sig", "utf-16", "cp1258", "latin-1"):
        try:
            with open(path, "r", encoding=enc) as fh:
                return json.load(fh)
        except UnicodeDecodeError as exc:
            last_err = exc
        except json.JSONDecodeError:
            raise  # file đọc được nhưng JSON sai -> báo thẳng, không thử encoding khác
    raise ValueError(f"Không đọc được file (encoding không hỗ trợ): {last_err}")


# ----------------------------------------------------------------- làm phẳng


def _put(out: dict, key: str, value: Any) -> None:
    """Ghi vào dict, không bao giờ đè key đã có (thêm hậu tố _2, _3...)."""
    if key not in out:
        out[key] = value
        return
    i = 2
    while f"{key}_{i}" in out:
        i += 1
    out[f"{key}_{i}"] = value


def json_text(value: Any) -> str:
    """Chuyển object/list thành chuỗi JSON gọn, không mất dữ liệu."""
    try:
        return json.dumps(value, ensure_ascii=False, separators=(",", ":"), default=str)
    except (TypeError, ValueError):
        return str(value)


def _flatten_into(value: Any, prefix: str, sep: str, max_depth: int, depth: int, out: dict) -> None:
    if isinstance(value, dict):
        if not value:
            _put(out, prefix, "{}")
        elif depth >= max_depth:
            _put(out, prefix, json_text(value))
        else:
            for k, v in value.items():
                child = f"{prefix}{sep}{k}" if prefix else str(k)
                _flatten_into(v, child, sep, max_depth, depth + 1, out)
    elif isinstance(value, list):
        # List luôn thành chuỗi JSON: nối bằng "; " sẽ hỏng nếu phần tử chứa "; ".
        _put(out, prefix, json_text(value))
    else:
        _put(out, prefix, value)


def flatten_record(rec: Any, sep: str = ".", max_depth: int = 20) -> dict:
    """Làm phẳng object lồng nhau thành cột kiểu 'user.address.city'."""
    out: dict = {}
    _flatten_into(rec, "", sep, max_depth, 0, out)
    return out


def shallow_record(rec: dict) -> dict:
    """Chế độ KHÔNG làm phẳng: giữ key cấp 1, object/list lồng thành chuỗi JSON."""
    return {str(k): json_text(v) if isinstance(v, (dict, list)) else v for k, v in rec.items()}


def normalize_record(rec: Any, flatten: bool, sep: str = ".", max_depth: int = 20) -> dict:
    """Chuẩn hoá 1 dòng thành dict phẳng (hoặc gần phẳng) để ghi ra Excel."""
    if isinstance(rec, dict):
        return flatten_record(rec, sep, max_depth) if flatten else shallow_record(rec)
    # Mảng của các giá trị vô hướng (hoặc mảng của mảng) -> một cột "value".
    return {"value": json_text(rec) if isinstance(rec, list) else rec}


def scan_columns(rows: Sequence, flatten: bool = True, sep: str = ".", max_depth: int = 20) -> list[str]:
    """Hợp tất cả key của mọi dòng, giữ thứ tự xuất hiện lần đầu."""
    cols: list[str] = []
    seen: set[str] = set()
    for rec in rows:
        for k in normalize_record(rec, flatten, sep, max_depth):
            if k not in seen:
                seen.add(k)
                cols.append(k)
    return cols


# ------------------------------------------------------------------ tìm mảng


def find_arrays(data: Any, max_depth: int = 8) -> list[ArrayCandidate]:
    """Duyệt JSON, tìm mọi mảng có thể xuất thành sheet.

    Không đi sâu vào bên trong một mảng đã nhận (tránh nổ số sheet).
    """
    found: list[ArrayCandidate] = []

    if isinstance(data, list):
        found.append(ArrayCandidate(path=ROOT_PATH, rows=data))
        return found
    if not isinstance(data, dict):
        return found

    def walk(node: Any, path: str, depth: int) -> None:
        if depth > max_depth:
            return
        if isinstance(node, dict):
            for k, v in node.items():
                walk(v, f"{path}.{k}" if path else str(k), depth + 1)
        elif isinstance(node, list):
            # Mọi list đều là ứng viên; list rỗng vẫn liệt kê để người dùng biết.
            found.append(ArrayCandidate(path=path, rows=node))

    walk(data, "", 0)

    # Gom các giá trị vô hướng ở cấp cao nhất (generatedAt, total...) thành một
    # bảng key/value, để phần metadata không bị mất khi xuất.
    scalars = [(k, v) for k, v in data.items() if not isinstance(v, (dict, list))]
    if scalars:
        found.append(
            ArrayCandidate(
                path=KV_PATH,
                rows=[{"key": k, "value": v} for k, v in scalars],
            )
        )
    return found


# --------------------------------------------------------------- tên sheet


def sanitize_sheet_name(name: str, taken: set[str]) -> str:
    """Làm sạch tên sheet: bỏ ký tự cấm, cắt 31 ký tự, tránh trùng."""
    s = SHEET_BAD_CHARS.sub("_", str(name)).strip().strip("'")
    if not s:
        s = "Sheet"
    s = s[:EXCEL_MAX_SHEET_NAME]
    base = s
    i = 2
    while s.casefold() in taken:
        suffix = f" ({i})"
        s = base[: EXCEL_MAX_SHEET_NAME - len(suffix)] + suffix
        i += 1
    taken.add(s.casefold())
    return s


def _default_sheet_name(path: str) -> str:
    if path == KV_PATH:
        return "ThongTin"
    if path == ROOT_PATH:
        return "DuLieu"
    return path


# ----------------------------------------------------------------- giá trị ô


def safe_cell(value: Any) -> tuple[Any, bool]:
    """Chuẩn hoá một giá trị thành thứ xlsxwriter ghi được.

    Trả về (giá_trị, bị_cắt_bớt). Giá trị None nghĩa là ô để trống.
    """
    if value is None:
        return None, False
    if isinstance(value, bool):
        return value, False
    if isinstance(value, (int, float)):
        # NaN / Infinity không ghi được vào xlsx -> chuyển thành chuỗi.
        if isinstance(value, float) and (value != value or value in (float("inf"), float("-inf"))):
            return str(value), False
        return value, False
    if isinstance(value, datetime.datetime):
        # xlsxwriter không nhận datetime có timezone -> ghi dạng chuỗi ISO.
        return (value.isoformat(), False) if value.tzinfo else (value, False)
    if isinstance(value, datetime.date):
        return datetime.datetime(value.year, value.month, value.day), False

    text = value if isinstance(value, str) else str(value)
    text = CONTROL_CHARS.sub("", text)
    if len(text) > EXCEL_MAX_CELL_CHARS:
        keep = EXCEL_MAX_CELL_CHARS - len(TRUNCATION_SUFFIX)
        return text[:keep] + TRUNCATION_SUFFIX, True
    return text, False


# ------------------------------------------------------------------ ghi Excel


def _write_cell(ws, row: int, col: int, value: Any, fmt=None) -> None:
    """Ghi đúng theo kiểu Python, không để xlsxwriter tự đoán."""
    if value is None:
        return  # ô trống thật, khác với chuỗi ""
    if isinstance(value, bool):
        ws.write_boolean(row, col, value, fmt)
    elif isinstance(value, (int, float)):
        ws.write_number(row, col, value, fmt)
    elif isinstance(value, datetime.datetime):
        ws.write_datetime(row, col, value, fmt)
    else:
        ws.write_string(row, col, str(value), fmt)


def _autofit(ws, columns: list[str], rows: Sequence[dict], limit: int = 1000) -> None:
    """Ước lượng độ rộng cột từ dữ liệu (chỉ xét `limit` dòng đầu cho nhanh)."""
    widths = [len(c) for c in columns]
    for i, rec in enumerate(rows):
        if i >= limit:
            break
        for j, c in enumerate(columns):
            v = rec.get(c)
            if v is None:
                continue
            width = len(str(v))
            if width > widths[j]:
                widths[j] = width
    for j, w in enumerate(widths):
        ws.set_column(j, j, min(max(w + 2, 8), 40))


def export_xlsx(
    candidates: Sequence[ArrayCandidate],
    out_path: str,
    flatten: bool = True,
    sep: str = ".",
    max_depth: int = 20,
    progress_cb: Callable[[int, int], None] | None = None,
) -> Report:
    """Ghi các mảng đã chọn ra file .xlsx.

    progress_cb(số_dòng_đã_ghi, tổng_số_dòng) được gọi theo lô để cập nhật GUI.
    """
    if not candidates:
        raise ValueError("Chưa chọn mảng dữ liệu nào để xuất.")

    total_rows = sum(c.count for c in candidates)
    # constant_memory giữ bộ nhớ thấp cho file lớn; chỉ an toàn vì ta ghi dòng
    # theo đúng thứ tự tăng dần trong từng sheet.
    options = {
        "strings_to_formulas": False,  # đừng biến chuỗi "=..." thành công thức
        "strings_to_urls": False,      # đừng biến "http://..." thành hyperlink
        "constant_memory": total_rows > 50_000,
    }

    report = Report(out_path=out_path)
    taken_names: set[str] = set()
    done_rows = 0

    wb = xlsxwriter.Workbook(out_path, options)
    header_fmt = wb.add_format({"bold": True, "bg_color": "#DDEBF7", "border": 1, "text_wrap": False})
    date_fmt = wb.add_format({"num_format": "yyyy-mm-dd hh:mm:ss"})
    truncated = 0

    try:
        for cand in candidates:
            columns = scan_columns(cand.rows, flatten, sep, max_depth)
            if not columns:
                report.warnings.append(f"Bỏ qua “{cand.path}”: không có cột nào.")
                continue
            if len(columns) > EXCEL_MAX_COLS:
                raise ValueError(
                    f"“{cand.path}” có {len(columns)} cột, vượt giới hạn {EXCEL_MAX_COLS} cột của Excel."
                )

            base_name = sanitize_sheet_name(cand.sheet_name or _default_sheet_name(cand.path), taken_names)
            # Chừa 1 dòng cho header.
            per_sheet = EXCEL_MAX_ROWS - 1
            chunks = [cand.rows[i : i + per_sheet] for i in range(0, len(cand.rows), per_sheet)] or [[]]
            if len(chunks) > 1:
                report.warnings.append(
                    f"“{cand.path}” có {cand.count} dòng, vượt giới hạn 1 sheet — đã tách thành {len(chunks)} sheet."
                )

            for ci, chunk in enumerate(chunks):
                name = base_name if ci == 0 else sanitize_sheet_name(
                    f"{base_name} ({ci + 1})", taken_names
                )
                ws = wb.add_worksheet(name)

                for j, col in enumerate(columns):
                    ws.write_string(0, j, col, header_fmt)

                normalized = [normalize_record(r, flatten, sep, max_depth) for r in chunk]
                _autofit(ws, columns, normalized)

                for i, rec in enumerate(normalized, start=1):
                    for j, col in enumerate(columns):
                        value, was_cut = safe_cell(rec.get(col))
                        if was_cut:
                            truncated += 1
                        _write_cell(ws, i, j, value, date_fmt if isinstance(value, datetime.datetime) else None)

                    done_rows += 1
                    if progress_cb and done_rows % 1000 == 0:
                        progress_cb(done_rows, total_rows)

                ws.freeze_panes(1, 0)
                if chunk:
                    ws.autofilter(0, 0, len(chunk), len(columns) - 1)

                report.sheets.append(SheetReport(name=name, rows=len(chunk), cols=len(columns)))
    finally:
        wb.close()

    if not report.sheets:
        # Không ghi được sheet nào. Nếu để nguyên, xlsxwriter sẽ tạo một "Sheet1"
        # trống trông như lỗi — xoá file và báo rõ nguyên nhân.
        try:
            os.remove(out_path)
        except OSError:
            pass
        detail = " ".join(report.warnings) or "Các bảng được chọn đều không có dữ liệu."
        raise ValueError(f"Không có gì để xuất. {detail}")

    if truncated:
        report.warnings.append(
            f"{truncated} ô dài quá {EXCEL_MAX_CELL_CHARS} ký tự đã bị cắt bớt (giới hạn của Excel)."
        )
    if progress_cb:
        progress_cb(total_rows, total_rows)
    return report


# ------------------------------------------------------------------------ CLI


def _setup_console() -> None:
    """Console Windows mặc định là cp1252, in tiếng Việt sẽ lỗi -> chuyển UTF-8."""
    if sys.platform == "win32":
        try:
            import ctypes

            ctypes.windll.kernel32.SetConsoleOutputCP(65001)
        except Exception:  # noqa: BLE001 - không phải console thật (IDE, pipe...)
            pass
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError, OSError):
            pass


def _main(argv: list[str] | None = None) -> int:
    import argparse

    _setup_console()
    parser = argparse.ArgumentParser(description="Chuyển đổi JSON sang Excel (không giao diện).")
    parser.add_argument("input", help="File JSON nguồn")
    parser.add_argument("-o", "--output", help="File .xlsx đích (mặc định: cùng tên với file nguồn)")
    parser.add_argument("--sheet", action="append", help="Chỉ xuất mảng có đường dẫn này (lặp lại được)")
    parser.add_argument("--no-flatten", action="store_true", help="Không làm phẳng JSON lồng nhau")
    parser.add_argument("--list", action="store_true", help="Chỉ liệt kê các mảng tìm được rồi thoát")
    args = parser.parse_args(argv)

    data = load_json(args.input)
    candidates = find_arrays(data)
    flatten = not args.no_flatten

    if args.list or not args.output:
        print(f"{'Đường dẫn':<40} {'Số dòng':>8}  Số cột")
        for c in candidates:
            cols = scan_columns(c.rows, flatten)
            print(f"{c.path:<40} {c.count:>8}  {len(cols)}")
        if not args.output:
            return 0

    if args.sheet:
        wanted = set(args.sheet)
        candidates = [c for c in candidates if c.path in wanted]
        if not candidates:
            parser.error(f"Không tìm thấy mảng nào khớp: {', '.join(sorted(wanted))}")

    out = args.output
    try:
        report = export_xlsx(candidates, out, flatten=flatten)
    except ValueError as exc:
        print(f"Lỗi: {exc}", file=sys.stderr)
        return 2
    print(report.summary())
    return 0


if __name__ == "__main__":
    raise SystemExit(_main())
