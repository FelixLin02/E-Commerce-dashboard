"""京鴻 ERP「商品銷售總表列印」Excel 解析器（ETL）。

處理流程：
1. 讀取 .xls / .xlsx 為純字串矩陣（不設表頭）。
2. 掃描頁首資訊：製表日期、日期區間；定位「商品型號 / 段 / 數量 / 合計」欄位。
3. 跨頁重複表頭濾除：只保留「型號非空 + 段碼非空 + 數量為數字」的資料列。
4. 段碼對照矩陣動態擷取（檔頭中的段碼 1/2/3 列），作為內建對照表之備援。
5. 17 欄尺寸寬轉長（Unpivot），輸出符合 `daily_sales_items` 之明細 DataFrame。
"""

from __future__ import annotations

import io
import re
from dataclasses import dataclass, field
from datetime import datetime

import pandas as pd

from .brand import BrandMatcher
from .size_codes import SIZE_SLOTS, is_known_code, normalize_label, resolve_size

HEADER_KEYWORDS = ("總公司", "商品銷售總表列印", "製表日期", "商品型號", "品牌區間", "分店區間")
DATE_RE = re.compile(r"^(20\d{2})(\d{2})(\d{2})$")
DEFAULT_GAP_OFFSET = 15  # ERP 版面在第 16 個尺寸欄位後有一個空白合併欄

OUTPUT_COLUMNS = [
    "sale_date", "sku", "brand", "duan_code", "size",
    "quantity", "unit_price", "row_total_amount",
]


@dataclass
class ParseResult:
    items: pd.DataFrame                     # 尺寸層級明細（寫入資料庫）
    rows: pd.DataFrame                      # 型號層級（未轉置）資料，供檢核顯示
    report_date: datetime | None = None     # 製表日期
    period_start: datetime | None = None    # 日期區間起
    period_end: datetime | None = None      # 日期區間迄
    file_size_codes: dict[str, list[str]] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)
    file_total_qty: float | None = None     # 報表最末「合計」列數量
    file_total_amount: float | None = None  # 報表最末「合計」列金額

    @property
    def default_sale_date(self) -> datetime | None:
        """預設銷售日期：日期區間起日（單日報表即為銷售日），否則用製表日期。"""
        return self.period_start or self.report_date


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------

def _cell(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = str(v).strip()
    if s.lower() == "nan":
        return ""
    if re.fullmatch(r"-?\d+\.0", s):
        s = s[:-2]
    return s


def _num(s: str) -> float | None:
    s = s.replace(",", "").strip()
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _parse_date(s: str) -> datetime | None:
    m = DATE_RE.match(s.replace("/", "").replace("-", ""))
    if not m:
        return None
    try:
        return datetime(int(m[1]), int(m[2]), int(m[3]))
    except ValueError:
        return None


def read_raw(file) -> pd.DataFrame:
    """讀取 ERP 匯出檔為字串矩陣。支援路徑、bytes 或 file-like（Streamlit UploadedFile）。"""
    if isinstance(file, (bytes, bytearray)):
        file = io.BytesIO(file)
    name = getattr(file, "name", str(file) if isinstance(file, str) else "")
    engine = "xlrd" if str(name).lower().endswith(".xls") else None
    try:
        df = pd.read_excel(file, header=None, dtype=str, engine=engine)
    except Exception:
        if hasattr(file, "seek"):
            file.seek(0)
        df = pd.read_excel(file, header=None, dtype=str)
    return df.map(_cell) if hasattr(df, "map") else df.applymap(_cell)


# ---------------------------------------------------------------------------
# main parser
# ---------------------------------------------------------------------------

def parse_erp(file, matcher: BrandMatcher, sale_date: datetime | None = None) -> ParseResult:
    raw = read_raw(file)
    warnings: list[str] = []
    ncols = raw.shape[1]

    # --- 1. 頁首資訊 & 欄位定位 -------------------------------------------------
    report_date = period_start = period_end = None
    col_sku, col_code, col_qty, col_amt = 0, 2, 3, 5
    header_found = False
    for _, r in raw.iterrows():
        vals = list(r)
        for j, v in enumerate(vals):
            if v.startswith("製表日期") and report_date is None:
                report_date = next((d for d in map(_parse_date, vals[j + 1:]) if d), None)
            elif v.startswith("日期區間") and period_start is None:
                dates = [d for d in map(_parse_date, vals[j + 1:]) if d]
                if dates:
                    period_start, period_end = dates[0], dates[-1]
        if not header_found and "商品型號" in vals:
            col_sku = vals.index("商品型號")
            for j, v in enumerate(vals):
                if v == "段":
                    col_code = j
                elif v == "數量":
                    col_qty = j
                elif v == "合計":
                    col_amt = j
            header_found = True
    if not header_found:
        warnings.append("找不到「商品型號」表頭列，改用預設欄位位置解析。")

    size_cols_all = list(range(col_amt + 1, ncols))

    # --- 2. 段碼對照列（檔頭）＆ 資料列 --------------------------------------------
    file_codes: dict[str, list[str]] = {}
    data_rows: list[list[str]] = []
    total_row: list[str] | None = None
    for _, r in raw.iterrows():
        vals = list(r)
        sku, code, qty = vals[col_sku], vals[col_code], vals[col_qty]
        if any(k in v for v in vals for k in HEADER_KEYWORDS):
            continue  # 跨頁重複表頭
        if sku == "合計":
            total_row = vals
            continue
        if not sku and code:
            # 段碼定義列：例如 [ , , '2', , , , '4', '-4', '5', ...]
            labels = [normalize_label(vals[c]) for c in size_cols_all if vals[c]]
            if labels and code.upper() not in file_codes:
                file_codes[code.upper()] = labels[:SIZE_SLOTS]
            continue
        if sku and code and _num(qty) is not None:
            data_rows.append(vals)

    if not data_rows:
        raise ValueError("檔案中找不到任何商品資料列，請確認是否為京鴻 ERP「商品銷售總表列印」匯出檔。")

    # --- 3. 決定 17 個尺寸欄位（排除版面空白欄）-----------------------------------
    size_cols = _detect_size_columns(size_cols_all, data_rows, total_row, warnings)

    # --- 4. 寬轉長 ---------------------------------------------------------------
    if sale_date is None:
        sale_date = period_start or report_date or datetime.now()
        if period_start and period_end and period_start != period_end:
            warnings.append(
                f"日期區間為 {period_start:%Y-%m-%d} ~ {period_end:%Y-%m-%d}（多日），"
                f"所有明細將記錄於 {sale_date:%Y-%m-%d}。"
            )

    items: list[dict] = []
    row_records: list[dict] = []
    unknown_codes: set[str] = set()
    unmatched: set[str] = set()
    for vals in data_rows:
        sku = vals[col_sku]
        code = vals[col_code].upper()
        qty_total = _num(vals[col_qty]) or 0
        amount = _num(vals[col_amt]) or 0
        brand, method = matcher.match(sku)
        if method == "none":
            unmatched.add(sku)
        if not is_known_code(code, file_codes):
            unknown_codes.add(code)
        unit_price = round(amount / qty_total, 2) if qty_total else 0.0

        size_qty: list[tuple[str, int]] = []
        for slot, c in enumerate(size_cols):
            q = _num(vals[c]) if c < len(vals) else None
            if q and q > 0:
                size_qty.append((resolve_size(code, slot, file_codes), int(q)))

        if not size_qty and qty_total > 0:
            # 17 欄皆空但合計數量 > 0 → 視為單一尺碼
            size_qty = [(resolve_size(code, 0, file_codes), int(qty_total))]
        else:
            s = sum(q for _, q in size_qty)
            if s != qty_total:
                warnings.append(f"型號 {sku}：尺寸數量加總 {s} 與合計數量 {qty_total:g} 不一致。")

        for size, q in size_qty:
            items.append({
                "sale_date": sale_date,
                "sku": sku,
                "brand": brand,
                "duan_code": code,
                "size": size,
                "quantity": q,
                "unit_price": unit_price,
                "row_total_amount": amount,
            })
        row_records.append({
            "型號": sku, "品牌": brand, "判定方式": method, "段碼": code,
            "數量": qty_total, "合計金額": amount, "單價": unit_price,
            "尺寸明細": ", ".join(f"{s}×{q}" for s, q in size_qty),
        })

    if unknown_codes:
        warnings.append(
            "下列段碼尚未建立尺寸對照，尺寸以「段碼X-第N欄」暫記，"
            f"請於 etl/size_codes.py 補齊：{', '.join(sorted(unknown_codes))}"
        )
    if unmatched:
        warnings.append(f"{len(unmatched)} 個型號無法判定品牌，已標記為「其他/未分類」。")

    file_total_qty = file_total_amount = None
    if total_row is not None:
        file_total_qty = _num(total_row[col_qty])
        file_total_amount = _num(total_row[col_amt])
        parsed_qty = sum(r["數量"] for r in row_records)
        parsed_amt = sum(r["合計金額"] for r in row_records)
        if file_total_qty is not None and abs(parsed_qty - file_total_qty) > 1e-6:
            warnings.append(f"解析總數量 {parsed_qty:g} 與報表合計 {file_total_qty:g} 不符。")
        if file_total_amount is not None and abs(parsed_amt - file_total_amount) > 0.5:
            warnings.append(f"解析總金額 {parsed_amt:,.0f} 與報表合計 {file_total_amount:,.0f} 不符。")

    return ParseResult(
        items=pd.DataFrame(items, columns=OUTPUT_COLUMNS),
        rows=pd.DataFrame(row_records),
        report_date=report_date,
        period_start=period_start,
        period_end=period_end,
        file_size_codes=file_codes,
        warnings=warnings,
        file_total_qty=file_total_qty,
        file_total_amount=file_total_amount,
    )


def _detect_size_columns(cols: list[int], data_rows: list[list[str]],
                         total_row: list[str] | None, warnings: list[str]) -> list[int]:
    """從合計欄之後的欄位中挑出 17 個實際尺寸欄。

    ERP 列印版面在尺寸欄中夾有空白合併欄（資料列從不填值），
    以「資料列與合計列皆為空」判定；判定不出時退回預設位移。
    """
    if len(cols) <= SIZE_SLOTS:
        return cols
    extra = len(cols) - SIZE_SLOTS

    def used(c: int) -> bool:
        if any(c < len(v) and _num(v[c]) for v in data_rows):
            return True
        return bool(total_row is not None and c < len(total_row) and _num(total_row[c]))

    empty = [c for c in cols if not used(c)]
    if len(empty) == extra:
        return [c for c in cols if c not in empty]
    if extra == 1 and len(cols) > DEFAULT_GAP_OFFSET and cols[DEFAULT_GAP_OFFSET] in empty:
        # 資料太少導致多個空欄時，採用 ERP 標準版面的空白欄位置
        gap = cols[DEFAULT_GAP_OFFSET]
        return [c for c in cols if c != gap]
    warnings.append("尺寸欄位數量異常，無法判定版面空白欄，取前 17 欄解析，請人工確認。")
    return cols[:SIZE_SLOTS]
