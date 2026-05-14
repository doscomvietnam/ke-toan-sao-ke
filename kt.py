# -*- coding: utf-8 -*-
"""
Tự động đọc sao kê ngân hàng -> xuất 1 file Excel kế toán gồm các sheet:
  1. Đầu ra - Phiếu thu
  2. Đầu ra - Phiếu chi
  3. LOG
(Có thêm sheet ẩn _ma_thong_ke để công thức VLOOKUP cột "Mã thống kê" hoạt động.)
"""

from __future__ import annotations

import io
import os
import re
import sys
import unicodedata
from datetime import datetime, date

import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet


# ============================================================
# CẤU HÌNH ĐƯỜNG DẪN
# ============================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_STATEMENT = os.path.join(BASE_DIR, "lich-su-giao-dich-tai-khoan (17).xlsx")
TEMPLATE_FILE = os.path.join(BASE_DIR, "Template-dau-ra.xlsx")
NGAN_HANG_FILE = os.path.join(BASE_DIR, "ngan-hang.xlsx")
MA_THONG_KE_FILE = os.path.join(BASE_DIR, "ma-thong-ke.xlsx")
OUTPUT_FILE = os.path.join(BASE_DIR, "Ket-qua-ke-toan.xlsx")

TK_NO_PHIEU_THU = "1121"
TK_CO_PHIEU_CHI = "1121.1"
MA_DON_VI_PHIEU_CHI = "CPC"

WARN_FILL = PatternFill(start_color="FFFFF2CC", end_color="FFFFF2CC", fill_type="solid")
ERROR_FILL = PatternFill(start_color="FFFFC7CE", end_color="FFFFC7CE", fill_type="solid")


# ============================================================
# UTILITIES
# ============================================================
def strip_accents(s: str) -> str:
    nfkd = unicodedata.normalize("NFD", s)
    return "".join(c for c in nfkd if not unicodedata.combining(c)).replace("đ", "d").replace("Đ", "D")


def norm_key(s) -> str:
    if s is None:
        return ""
    s = str(s)
    s = strip_accents(s)
    s = s.lower()
    # Bỏ dấu câu và ký tự đặc biệt thường gây nhiễu khi so khớp header
    for ch in [":", "/", ".", ",", "(", ")", "*", "-", "\t", "\n", "\r"]:
        s = s.replace(ch, " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def header_match(cell_value: str, candidates: list[str]) -> bool:
    """Match một chiều: pattern (đã chuẩn hóa) phải là substring của cell.
    Đồng thời cell phải đủ dài (>=2 ký tự) để tránh false positive với giá trị 1 ký tự."""
    n = norm_key(cell_value)
    if not n or len(n) < 2:
        return False
    for c in candidates:
        cn = norm_key(c)
        if not cn:
            continue
        if cn in n:
            return True
    return False


def clean_account_str(value) -> str:
    """Số tài khoản dạng text, không mất số 0 đầu, bỏ khoảng trắng và ký tự không phải số."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    s = str(value).strip()
    s = re.sub(r"\s+", "", s)
    # Bỏ ký tự không phải số nhưng giữ nguyên nếu cần text
    digits = re.sub(r"\D", "", s)
    return digits if digits else s


def to_number(value):
    """Chuẩn hóa số tiền về kiểu số (float/int). Trả về None nếu rỗng/không parse được. Giữ giá trị 0."""
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return value
    s = str(value).strip()
    if not s:
        return None
    s = s.replace(",", "").replace(" ", "")
    s = s.replace("VND", "").replace("vnd", "")
    try:
        return float(s)
    except ValueError:
        return None


_DATE_PATTERNS = [
    "%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%Y/%m/%d",
    "%d/%m/%Y %H:%M:%S", "%d-%m-%Y %H:%M:%S",
    "%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S",
    "%d/%m/%y", "%d-%m-%y",
]


def parse_date_only(value):
    """Trả về chuỗi dd/mm/yyyy hoặc None nếu không parse được."""
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.strftime("%d/%m/%Y")
    if isinstance(value, date):
        return value.strftime("%d/%m/%Y")
    s = str(value).strip()
    if not s:
        return None
    # Ưu tiên 1: thử parse nguyên chuỗi (dùng cho Techcombank "2026-05-14" hoặc datetime string)
    for fmt in _DATE_PATTERNS:
        try:
            d = datetime.strptime(s, fmt)
            return d.strftime("%d/%m/%Y")
        except ValueError:
            continue
    # Ưu tiên 2: tìm pattern yyyy-mm-dd trong chuỗi (trước pattern dd/mm/yy để tránh nhầm)
    m = re.search(r"(\d{4}[/\-]\d{1,2}[/\-]\d{1,2})", s)
    if m:
        candidate = m.group(1).replace("-", "/")
        try:
            d = datetime.strptime(candidate, "%Y/%m/%d")
            return d.strftime("%d/%m/%Y")
        except ValueError:
            pass
    # Ưu tiên 3: tìm pattern dd/mm/yyyy (dùng cho Vietcombank "22/04/2026\n5284 - 37679")
    m = re.search(r"(\d{1,2}[/\-]\d{1,2}[/\-]\d{4})", s)
    if m:
        candidate = m.group(1).replace("-", "/")
        try:
            d = datetime.strptime(candidate, "%d/%m/%Y")
            return d.strftime("%d/%m/%Y")
        except ValueError:
            pass
    # Ưu tiên 4: dd/mm/yy (năm 2 chữ số) — cuối cùng vì dễ nhầm
    m = re.search(r"(\d{1,2}[/\-]\d{1,2}[/\-]\d{2})\b", s)
    if m:
        candidate = m.group(1).replace("-", "/")
        try:
            d = datetime.strptime(candidate, "%d/%m/%y")
            return d.strftime("%d/%m/%Y")
        except ValueError:
            pass
    return None


def format_phieu_thu_date(value):
    """
    Phiếu thu — quy tắc:
    - Nếu cell chỉ chứa 1 ngày (Techcombank "2026-05-14", datetime) → chuẩn hóa dd/mm/yyyy
    - Nếu cell chứa thêm thông tin khác (Vietcombank "22/04/2026\\n5284-37679") → giữ nguyên giá trị gốc
    """
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%d/%m/%Y")
    if isinstance(value, date):
        return value.strftime("%d/%m/%Y")
    s = str(value).strip()
    if not s:
        return ""
    # Nếu nguyên chuỗi đúng là 1 ngày → chuẩn hóa
    for fmt in _DATE_PATTERNS:
        try:
            d = datetime.strptime(s, fmt)
            return d.strftime("%d/%m/%Y")
        except ValueError:
            continue
    # Có thêm thông tin → giữ nguyên (Vietcombank case)
    return s


# ============================================================
# ĐỌC FILE NGÂN HÀNG
# ============================================================
def load_bank_lookup(path: str, log_rows: list) -> dict[str, str]:
    if not os.path.exists(path):
        log_rows.append([_now(), "", "ngân hàng", "Thiếu file", f"Không tìm thấy file ngân hàng: {path}"])
        return {}
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    header_row_idx = None
    col_acc = None
    col_bank = None
    for r in range(1, min(ws.max_row, 5) + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if header_match(v, ["so tai khoan", "số tài khoản", "account number", "stk"]):
                col_acc = c
                header_row_idx = r
            elif header_match(v, ["ten ngan hang", "tên ngân hàng", "bank name", "ngan hang"]):
                col_bank = c
                if header_row_idx is None:
                    header_row_idx = r
    if header_row_idx is None or col_acc is None or col_bank is None:
        log_rows.append([_now(), "", "ngân hàng", "Thiếu cột", "Không tìm thấy cột số tài khoản / tên ngân hàng"])
        return {}
    lookup = {}
    for r in range(header_row_idx + 1, ws.max_row + 1):
        acc = clean_account_str(ws.cell(r, col_acc).value)
        bank = ws.cell(r, col_bank).value
        if acc and bank:
            lookup[acc] = str(bank).strip()
    return lookup


# ============================================================
# ĐỌC FILE MÃ THỐNG KÊ
# ============================================================
def load_stat_codes(path: str, log_rows: list) -> list[tuple[str, str, str]]:
    if not os.path.exists(path):
        log_rows.append([_now(), "", "mã thống kê", "Thiếu file", f"Không tìm thấy file mã thống kê: {path}"])
        return []
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    header_row_idx = None
    col_tkno = None
    col_kmcp = None
    col_mtk = None
    for r in range(1, min(ws.max_row, 5) + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if header_match(v, ["tk no", "tk nợ"]):
                col_tkno = c
                header_row_idx = r
            elif header_match(v, ["ma kmcp", "mã kmcp", "ma khoan muc chi phi", "mã khoản mục chi phí"]):
                col_kmcp = c
                if header_row_idx is None:
                    header_row_idx = r
            elif header_match(v, ["ma thong ke", "mã thống kê"]):
                col_mtk = c
                if header_row_idx is None:
                    header_row_idx = r
    if header_row_idx is None or col_tkno is None or col_kmcp is None or col_mtk is None:
        log_rows.append([_now(), "", "mã thống kê", "Thiếu cột", "Không tìm thấy đủ cột TK Nợ / Mã KMCP / Mã thống kê"])
        return []
    rows = []
    for r in range(header_row_idx + 1, ws.max_row + 1):
        tkno = ws.cell(r, col_tkno).value
        kmcp = ws.cell(r, col_kmcp).value
        mtk = ws.cell(r, col_mtk).value
        if tkno is None and kmcp is None:
            continue
        rows.append((_to_text(tkno), _to_text(kmcp), _to_text(mtk)))
    return rows


def _to_text(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


# ============================================================
# ĐỌC FILE SAO KÊ NGÂN HÀNG
# ============================================================
def read_bank_statement(source, log_rows: list):
    """source có thể là đường dẫn (str) hoặc file-like (BytesIO / bytes)."""
    if isinstance(source, (bytes, bytearray)):
        source = io.BytesIO(source)
    elif isinstance(source, str):
        if not os.path.exists(source):
            log_rows.append([_now(), "", "Đầu vào", "Thiếu file", f"Không tìm thấy file sao kê: {source}"])
            return None, []
    wb = openpyxl.load_workbook(source, data_only=True)
    ws = wb[wb.sheetnames[0]]

    # Tìm số tài khoản của chủ sao kê
    account_number = ""
    for r in range(1, min(ws.max_row, 15) + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if v and header_match(v, ["so tai khoan account number", "so tai khoan", "account number"]):
                # tìm giá trị bên phải trên cùng dòng
                for c2 in range(c + 1, ws.max_column + 1):
                    v2 = ws.cell(r, c2).value
                    if v2 not in (None, ""):
                        account_number = clean_account_str(v2)
                        break
                if account_number:
                    break
        if account_number:
            break

    # Tìm dòng header của bảng giao dịch (hỗ trợ Vietcombank, Techcombank...)
    header_row_idx = None
    col_idx = {"ngay": None, "credit": None, "debit": None, "balance": None,
               "noi_dung": None, "stt": None}
    for r in range(1, ws.max_row + 1):
        row_values = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        if not any(row_values):
            continue
        # heuristic: dòng có chứa ít nhất "ngay" và một trong credit/debit
        has_date = any(header_match(v, ["ngay tnx date so ct doc no", "ngay tnx date",
                                        "tnx date", "ngay giao dich", "transaction date",
                                        "doc no"]) for v in row_values if v)
        has_credit = any(header_match(v, ["so tien ghi co", "credit", "ghi co"]) for v in row_values if v)
        has_debit = any(header_match(v, ["so tien ghi no", "debit", "ghi no"]) for v in row_values if v)
        if has_date and (has_credit or has_debit):
            header_row_idx = r
            for c, v in enumerate(row_values, start=1):
                if v is None:
                    continue
                if col_idx["ngay"] is None and header_match(v, [
                        "ngay tnx date so ct doc no", "ngay tnx date", "tnx date",
                        "ngay giao dich", "transaction date", "doc no", "ngay"]):
                    col_idx["ngay"] = c
                elif col_idx["credit"] is None and header_match(v, ["so tien ghi co", "credit", "ghi co"]):
                    col_idx["credit"] = c
                elif col_idx["debit"] is None and header_match(v, ["so tien ghi no", "debit", "ghi no"]):
                    col_idx["debit"] = c
                elif col_idx["balance"] is None and header_match(v, ["so du", "balance"]):
                    col_idx["balance"] = c
                elif col_idx["noi_dung"] is None and header_match(v, [
                        "noi dung chi tiet", "transactions in detail",
                        "dien giai", "description", "noi dung"]):
                    col_idx["noi_dung"] = c
                elif col_idx["stt"] is None and header_match(v, ["stt", "no"]):
                    col_idx["stt"] = c
            break

    if header_row_idx is None:
        log_rows.append([_now(), "", "Đầu vào", "Thiếu cột", "Không tìm thấy dòng header bảng giao dịch"])
        return None, []

    required = ["ngay", "credit", "debit", "noi_dung"]
    missing = [k for k in required if col_idx[k] is None]
    if missing:
        log_rows.append([_now(), "", "Đầu vào", "Thiếu cột", f"Không tìm thấy cột: {missing}"])

    # Đọc các dòng giao dịch
    transactions = []
    for r in range(header_row_idx + 1, ws.max_row + 1):
        ngay_raw = ws.cell(r, col_idx["ngay"]).value if col_idx["ngay"] else None
        noi_dung = ws.cell(r, col_idx["noi_dung"]).value if col_idx["noi_dung"] else None
        credit_raw = ws.cell(r, col_idx["credit"]).value if col_idx["credit"] else None
        debit_raw = ws.cell(r, col_idx["debit"]).value if col_idx["debit"] else None

        # Bỏ qua dòng trống
        if all(v in (None, "") for v in (ngay_raw, noi_dung, credit_raw, debit_raw)):
            continue
        # Bỏ qua dòng tổng / dòng chú thích / số dư cuối kỳ
        full_row_text = " ".join(str(ws.cell(r, c).value)
                                 for c in range(1, ws.max_column + 1)
                                 if ws.cell(r, c).value is not None)
        n_full = norm_key(full_row_text)
        if any(k in n_full for k in ["tong so", "total", "so du cuoi ky", "closing balance",
                                      "ghi chu note", "ghi chu giay"]):
            continue
        # Cần có ít nhất một trong credit/debit là số > 0 hoặc có ngày
        credit_num = to_number(credit_raw)
        debit_num = to_number(debit_raw)
        # Một số ngân hàng (Techcombank) ghi Debit dạng số âm — luôn lấy giá trị tuyệt đối
        if credit_num is not None:
            credit_num = abs(credit_num)
        if debit_num is not None:
            debit_num = abs(debit_num)
        if credit_num is None and debit_num is None and not ngay_raw:
            continue

        transactions.append({
            "src_row": r,
            "ngay_raw": ngay_raw,
            "noi_dung": noi_dung,
            "credit": credit_num,
            "debit": debit_num,
            "account_number": account_number,
        })

    return account_number, transactions


# ============================================================
# GHI FILE KẾT QUẢ
# ============================================================
def _now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def find_template_header_row(ws: Worksheet) -> tuple[int, dict[str, int]]:
    """Tìm dòng header trong sheet template, trả về (row_index, {normalized_name: col_index})."""
    for r in range(1, min(ws.max_row, 10) + 1):
        row_values = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        non_empty = [v for v in row_values if v is not None and str(v).strip()]
        if len(non_empty) < 5:
            continue
        # Có chứa "Ngày hạch toán" hoặc "Số tiền" ?
        if any(header_match(v, ["ngay hach toan", "so tien", "dien giai hach toan", "phuong thuc thanh toan"]) for v in row_values if v):
            mapping = {}
            for c, v in enumerate(row_values, start=1):
                if v is None:
                    continue
                mapping[norm_key(v)] = c
            return r, mapping
    return 1, {}


def find_col(mapping: dict[str, int], candidates: list[str]) -> int | None:
    """Tìm cột theo header: thử lần lượt từng pattern (theo thứ tự ưu tiên),
    với mỗi pattern thử exact match trước rồi mới substring."""
    cand_norms = [norm_key(c) for c in candidates if norm_key(c)]
    # exact match trước
    for cn in cand_norms:
        for k, v in mapping.items():
            if k == cn:
                return v
    # substring: pattern là substring của key cột
    for cn in cand_norms:
        for k, v in mapping.items():
            if cn in k:
                return v
    return None


def build_output(transactions: list, bank_lookup: dict, stat_codes: list, log_rows: list,
                 output=None, template_path: str | None = None):
    """
    output: đường dẫn (str) hoặc BytesIO. Nếu None, dùng OUTPUT_FILE mặc định.
    template_path: đường dẫn template. Nếu None dùng TEMPLATE_FILE mặc định.
    Trả về đường dẫn output (nếu output là str) hoặc BytesIO (nếu output là BytesIO).
    """
    if output is None:
        output = OUTPUT_FILE
    if template_path is None:
        template_path = TEMPLATE_FILE
    # Mở template
    wb = openpyxl.load_workbook(template_path)
    # Sheet phiếu thu / phiếu chi: tìm theo tên hiện có
    sn_thu = None
    sn_chi = None
    for sn in wb.sheetnames:
        n = norm_key(sn)
        if "thu" in n:
            sn_thu = sn
        elif "chi" in n:
            sn_chi = sn
    if sn_thu is None or sn_chi is None:
        raise RuntimeError("Template thiếu sheet Phiếu thu / Phiếu chi")

    ws_thu = wb[sn_thu]
    ws_chi = wb[sn_chi]

    thu_header_row, thu_map = find_template_header_row(ws_thu)
    chi_header_row, chi_map = find_template_header_row(ws_chi)

    # ============ MAP CỘT PHIẾU THU ============
    thu_cols = {
        "ngay_hach_toan": find_col(thu_map, ["ngay hach toan"]),
        "ngay_chung_tu": find_col(thu_map, ["ngay chung tu"]),
        "dien_giai": find_col(thu_map, ["dien giai hach toan", "dien giai"]),
        "nop_vao_tk": find_col(thu_map, ["nop vao tk"]),
        "mo_tai_nh": find_col(thu_map, ["mo tai ngan hang"]),
        "tk_no": find_col(thu_map, ["tk no"]),
        "so_tien": find_col(thu_map, ["so tien"]),
    }

    # ============ MAP CỘT PHIẾU CHI ============
    chi_cols = {
        "ngay_hach_toan": find_col(chi_map, ["ngay hach toan"]),
        "ngay_chung_tu": find_col(chi_map, ["ngay chung tu"]),
        "dien_giai": find_col(chi_map, ["dien giai hach toan", "dien giai"]),
        "stk_chi": find_col(chi_map, ["so tai khoan chi"]),
        "ten_nh_chi": find_col(chi_map, ["ten ngan hang chi"]),
        "tk_co": find_col(chi_map, ["tk co"]),
        "tk_no": find_col(chi_map, ["tk no"]),
        "so_tien": find_col(chi_map, ["so tien"]),
        "noi_dung_tt": find_col(chi_map, ["noi dung thanh toan"]),
        "ma_don_vi": find_col(chi_map, ["ma don vi"]),
        "ma_kmcp": find_col(chi_map, ["ma khoan muc chi phi"]),
        "ma_thong_ke": find_col(chi_map, ["ma thong ke"]),
    }

    # ============ KIỂM TRA TRÙNG ============
    dup_key_count = {}
    dup_first_row = {}
    for tx in transactions:
        key = (
            norm_key(str(tx["ngay_raw"])),
            norm_key(str(tx["noi_dung"])),
            tx["account_number"],
            tx["credit"] or 0,
            tx["debit"] or 0,
        )
        if key in dup_first_row:
            dup_key_count[key] = dup_key_count.get(key, 1) + 1
            tx["_dup_with"] = dup_first_row[key]
        else:
            dup_first_row[key] = tx["src_row"]

    # ============ ĐIỀN DỮ LIỆU ============
    thu_start = thu_header_row + 1
    chi_start = chi_header_row + 1
    thu_row = thu_start
    chi_row = chi_start

    for tx in transactions:
        notes_thu = []
        notes_chi = []
        # Debit/Credit = 0 vẫn coi là có giao dịch (cột Số tiền sẽ để trống)
        has_credit = tx["credit"] is not None
        has_debit = tx["debit"] is not None

        # Cảnh báo trùng
        if "_dup_with" in tx:
            log_rows.append([_now(), tx["src_row"], "Đầu vào",
                             "Giao dịch trùng",
                             f"Trùng với dòng {tx['_dup_with']}"])

        # Cảnh báo có cả credit + debit
        if has_credit and has_debit:
            log_rows.append([_now(), tx["src_row"], "Đầu vào",
                             "Cả Credit và Debit",
                             "Dòng có cả Credit và Debit, sẽ tạo cả 2 phiếu"])

        # Không có credit và debit
        if not has_credit and not has_debit:
            log_rows.append([_now(), tx["src_row"], "Đầu vào",
                             "Bỏ qua",
                             "Không có Credit/Debit"])
            continue

        # ============ PHIẾU THU ============
        if has_credit:
            ngay_raw_str = format_phieu_thu_date(tx["ngay_raw"])

            if thu_cols["ngay_hach_toan"]:
                ws_thu.cell(thu_row, thu_cols["ngay_hach_toan"], ngay_raw_str)
            if thu_cols["ngay_chung_tu"]:
                ws_thu.cell(thu_row, thu_cols["ngay_chung_tu"], ngay_raw_str)
            if thu_cols["dien_giai"]:
                ws_thu.cell(thu_row, thu_cols["dien_giai"], tx["noi_dung"])
            if thu_cols["nop_vao_tk"]:
                cell = ws_thu.cell(thu_row, thu_cols["nop_vao_tk"], tx["account_number"])
                cell.number_format = "@"
            if thu_cols["tk_no"]:
                ws_thu.cell(thu_row, thu_cols["tk_no"], TK_NO_PHIEU_THU)
            if thu_cols["so_tien"]:
                # Credit = 0 → để trống cột Số tiền theo yêu cầu
                if tx["credit"] and tx["credit"] != 0:
                    c = ws_thu.cell(thu_row, thu_cols["so_tien"], tx["credit"])
                    c.number_format = "#,##0"

            # Tra ngân hàng
            bank_name = bank_lookup.get(tx["account_number"], "")
            if thu_cols["mo_tai_nh"]:
                cell = ws_thu.cell(thu_row, thu_cols["mo_tai_nh"], bank_name or "")
                if not bank_name:
                    cell.fill = WARN_FILL
                    log_rows.append([_now(), tx["src_row"], "Phiếu thu",
                                     "Không tìm thấy ngân hàng",
                                     f"Số TK {tx['account_number']} không có trong sheet ngân hàng"])
                    notes_thu.append("Không tìm thấy ngân hàng")

            thu_row += 1

        # ============ PHIẾU CHI ============
        if has_debit:
            ngay_dmy = parse_date_only(tx["ngay_raw"])
            if ngay_dmy is None:
                # Giữ nguyên giá trị gốc và log
                ngay_keep = str(tx["ngay_raw"]) if tx["ngay_raw"] is not None else ""
                log_rows.append([_now(), tx["src_row"], "Phiếu chi",
                                 "Không parse được ngày",
                                 f"Giá trị gốc: {ngay_keep!r}"])
                notes_chi.append("Không parse được ngày")
                date_cell_value = ngay_keep
                noi_dung_tt = f"Các khoản chi ngày {ngay_keep}"
            else:
                date_cell_value = ngay_dmy
                noi_dung_tt = f"Các khoản chi ngày {ngay_dmy}"

            if chi_cols["ngay_hach_toan"]:
                c = ws_chi.cell(chi_row, chi_cols["ngay_hach_toan"], date_cell_value)
                if ngay_dmy is None:
                    c.fill = WARN_FILL
            if chi_cols["ngay_chung_tu"]:
                c = ws_chi.cell(chi_row, chi_cols["ngay_chung_tu"], date_cell_value)
                if ngay_dmy is None:
                    c.fill = WARN_FILL
            if chi_cols["dien_giai"]:
                ws_chi.cell(chi_row, chi_cols["dien_giai"], tx["noi_dung"])
            if chi_cols["stk_chi"]:
                cell = ws_chi.cell(chi_row, chi_cols["stk_chi"], tx["account_number"])
                cell.number_format = "@"
            if chi_cols["tk_co"]:
                ws_chi.cell(chi_row, chi_cols["tk_co"], TK_CO_PHIEU_CHI)
            if chi_cols["so_tien"]:
                # Debit = 0 → để trống cột Số tiền theo yêu cầu
                if tx["debit"] and tx["debit"] != 0:
                    c = ws_chi.cell(chi_row, chi_cols["so_tien"], tx["debit"])
                    c.number_format = "#,##0"
            if chi_cols["noi_dung_tt"]:
                ws_chi.cell(chi_row, chi_cols["noi_dung_tt"], noi_dung_tt)
            if chi_cols["ma_don_vi"]:
                ws_chi.cell(chi_row, chi_cols["ma_don_vi"], MA_DON_VI_PHIEU_CHI)
            # TK Nợ và Mã KMCP để trống để nhập tay

            # Tra ngân hàng
            bank_name = bank_lookup.get(tx["account_number"], "")
            if chi_cols["ten_nh_chi"]:
                cell = ws_chi.cell(chi_row, chi_cols["ten_nh_chi"], bank_name or "")
                if not bank_name:
                    cell.fill = WARN_FILL
                    log_rows.append([_now(), tx["src_row"], "Phiếu chi",
                                     "Không tìm thấy ngân hàng",
                                     f"Số TK {tx['account_number']} không có trong sheet ngân hàng"])
                    notes_chi.append("Không tìm thấy ngân hàng")

            # Công thức Mã thống kê
            if chi_cols["ma_thong_ke"] and chi_cols["tk_no"] and chi_cols["ma_kmcp"]:
                tkno_letter = get_column_letter(chi_cols["tk_no"])
                kmcp_letter = get_column_letter(chi_cols["ma_kmcp"])
                formula = (
                    f'=IF(AND({tkno_letter}{chi_row}<>"",{kmcp_letter}{chi_row}<>""),'
                    f'IFERROR(VLOOKUP({tkno_letter}{chi_row}&"|"&{kmcp_letter}{chi_row},'
                    f'_ma_thong_ke!$A:$D,4,FALSE),""),"")'
                )
                ws_chi.cell(chi_row, chi_cols["ma_thong_ke"], formula)

            chi_row += 1

    # ============ SHEET VERY HIDDEN: _ma_thong_ke ============
    # Sheet này phục vụ công thức VLOOKUP ở cột "Mã thống kê" — kế toán không thấy
    if "_ma_thong_ke" in wb.sheetnames:
        del wb["_ma_thong_ke"]
    ws_mtk = wb.create_sheet("_ma_thong_ke")
    ws_mtk.append(["Key", "TK Nợ", "Mã KMCP", "Mã thống kê"])
    for tkno, kmcp, mtk in stat_codes:
        key = f"{tkno}|{kmcp}"
        ws_mtk.append([key, tkno, kmcp, mtk])
    ws_mtk.sheet_state = "veryHidden"

    # ============ ĐỔI TÊN SHEET ============
    new_thu_name = "Phiếu thu tiền gửi"
    new_chi_name = "Phiếu chi tiền gửi"
    if ws_thu.title != new_thu_name:
        ws_thu.title = new_thu_name
    if ws_chi.title != new_chi_name:
        ws_chi.title = new_chi_name

    # Xóa sheet LOG cũ nếu có (nhật ký giờ in ra terminal, không gắn vào file)
    if "LOG" in wb.sheetnames:
        del wb["LOG"]

    # ============ CHỈNH ĐỘ RỘNG CỘT (tránh ######) ============
    thu_widths = {
        "ngay_hach_toan": 24, "ngay_chung_tu": 24,
        "dien_giai": 50,
        "nop_vao_tk": 16, "mo_tai_nh": 36,
        "tk_no": 10, "so_tien": 18,
    }
    chi_widths = {
        "ngay_hach_toan": 14, "ngay_chung_tu": 14,
        "dien_giai": 50,
        "stk_chi": 16, "ten_nh_chi": 36,
        "tk_co": 10, "tk_no": 10, "so_tien": 18,
        "noi_dung_tt": 30, "ma_don_vi": 10,
        "ma_kmcp": 12, "ma_thong_ke": 14,
    }
    for key, col_idx in thu_cols.items():
        w = thu_widths.get(key)
        if col_idx and w:
            ws_thu.column_dimensions[get_column_letter(col_idx)].width = w
    for key, col_idx in chi_cols.items():
        w = chi_widths.get(key)
        if col_idx and w:
            ws_chi.column_dimensions[get_column_letter(col_idx)].width = w

    # Thứ tự sheet: Phiếu thu, Phiếu chi, (_ma_thong_ke veryHidden)
    order = [new_thu_name, new_chi_name, "_ma_thong_ke"]
    wb._sheets = [wb[name] for name in order if name in wb.sheetnames]

    wb.save(output)
    return output


# ============================================================
# PUBLIC API (dùng cho web / GUI)
# ============================================================
def process(input_source, ngan_hang_path: str | None = None,
            ma_thong_ke_path: str | None = None,
            template_path: str | None = None) -> tuple[bytes, list]:
    """
    Xử lý 1 file sao kê và trả về (output_bytes, log_rows).
    input_source: đường dẫn file sao kê (str) hoặc bytes.
    """
    if ngan_hang_path is None:
        ngan_hang_path = NGAN_HANG_FILE
    if ma_thong_ke_path is None:
        ma_thong_ke_path = MA_THONG_KE_FILE
    if template_path is None:
        template_path = TEMPLATE_FILE

    log_rows: list = []
    _, transactions = read_bank_statement(input_source, log_rows)
    if transactions is None:
        # Trả về workbook rỗng với LOG để người dùng biết lỗi
        bank_lookup = {}
        stat_codes = []
        transactions = []
    else:
        bank_lookup = load_bank_lookup(ngan_hang_path, log_rows)
        stat_codes = load_stat_codes(ma_thong_ke_path, log_rows)

    buf = io.BytesIO()
    build_output(transactions, bank_lookup, stat_codes, log_rows,
                 output=buf, template_path=template_path)
    return buf.getvalue(), log_rows


# ============================================================
# MAIN
# ============================================================
def main():
    log_rows: list = []

    print(f"Đang đọc sao kê: {INPUT_STATEMENT}")
    account_number, transactions = read_bank_statement(INPUT_STATEMENT, log_rows)
    if transactions is None:
        print("LỖI: Không đọc được sao kê. Xem LOG.")
        return
    print(f"  - Số tài khoản: {account_number}")
    print(f"  - Số giao dịch: {len(transactions)}")

    print(f"Đang đọc ngân hàng: {NGAN_HANG_FILE}")
    bank_lookup = load_bank_lookup(NGAN_HANG_FILE, log_rows)
    print(f"  - Số dòng ngân hàng: {len(bank_lookup)}")

    print(f"Đang đọc mã thống kê: {MA_THONG_KE_FILE}")
    stat_codes = load_stat_codes(MA_THONG_KE_FILE, log_rows)
    print(f"  - Số dòng mã thống kê: {len(stat_codes)}")

    print(f"Đang ghi kết quả vào: {OUTPUT_FILE}")
    out = build_output(transactions, bank_lookup, stat_codes, log_rows)
    print(f"Xong: {out}")
    print_log(log_rows)


def print_log(log_rows: list, stream=None) -> None:
    """In nhật ký lỗi/cảnh báo ra terminal."""
    if stream is None:
        stream = sys.stderr
    if not log_rows:
        print(f"[LOG] Không có cảnh báo.", file=stream)
        return
    print(f"[LOG] {len(log_rows)} dòng cảnh báo:", file=stream)
    for row in log_rows:
        # row = [thoi_gian, dong_dau_vao, sheet, loai, noi_dung]
        ts, src, sheet, kind, detail = (row + [""] * 5)[:5]
        print(f"  - [{ts}] dòng={src} sheet={sheet} | {kind}: {detail}", file=stream)


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    main()
