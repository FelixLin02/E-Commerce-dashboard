"""品牌與商品主檔判定模組。

判定順序：
1. 精確比對：資料庫 `product_catalog`（或 `品牌/*.xlsx`）之「主商品貨號」（正規化：大寫、去除 - 與空白）。
2. 前綴比對：ERP 型號與主檔貨號互為前綴（例如 ERP `AC4562WBKF` ↔ 主檔 `AC4562WBK`），
   取最長者，且長度需 ≥ 6 以避免誤判。
3. 正規表示式規則：依各品牌貨號編碼規則（BRAND_REGEX_RULES，依序比對，先命中者優先）。
4. 皆未命中 → `其他/未分類`。
"""

from __future__ import annotations

import io
import re
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd

UNCLASSIFIED = "其他/未分類"
MIXED_FILES = {"配件", "其他品牌"}
MIN_PREFIX_LEN = 6

# 商品名稱 → 品牌 關鍵字（大寫比對，取最先出現者）
NAME_KEYWORDS: dict[str, list[str]] = {
    "NIKE": ["NIKE", "JORDAN"],
    "adidas": ["ADIDAS"],
    "NEW BALANCE": ["NEW BALANCE", "NB "],
    "ASICS": ["ASICS"],
    "SKECHERS": ["SKECHERS"],
    "PALLADIUM": ["PALLADIUM"],
    "REEBOK": ["REEBOK"],
    "FILA": ["FILA"],
    "Converse": ["CONVERSE"],
    "Havaianas": ["HAVAIANAS", "哈瓦仕"],
    "Teva": ["TEVA"],
    "Native": ["NATIVE"],
    "Pony": ["PONY"],
    "MERRELL": ["MERRELL"],
    "Spalding": ["SPALDING", "斯伯丁"],
    "BN3TH": ["BN3TH"],
    "IFME": ["IFME"],
    "G.P": ["G.P", "GP "],
    "LP SUPPORT": ["LP SUPPORT", "LP "],
    "EVERLAST": ["EVERLAST"],
    "K-SWISS": ["K-SWISS"],
    "DIADORA": ["DIADORA", "迪亞多那"],
    "LACOSTE": ["LACOSTE"],
    "ROYAL": ["ROYAL"],
    "NBA": ["NBA "],
}

# 各品牌貨號規則（正規化後之型號；依序比對，越特定的規則放越前面）
BRAND_REGEX_RULES: list[tuple[str, str]] = [
    ("Havaianas", r"^4\d{6}\d{4}[A-Z]$"),            # 4145741-3624U
    ("EVERLAST", r"^4[0245][25]5\d{6}$"),            # 4255162245 / 4525162225 / 4025220100
    ("LACOSTE", r"^4\d[A-Z]{3}\d{4}[A-Z0-9]{3}$"),   # 42SMA0075454
    ("FILA", r"^\d[A-Z]\d{3}[A-Z]\d{3}$"),           # 3-J414A-188
    ("IFME", r"^IF[A-Z0-9]{2}\d{6}$"),               # IF20-481302
    ("MERRELL", r"^ML\d{5,8}[A-Z]?$"),               # ML036326
    ("Spalding", r"^SP[AB][A-Z0-9]{5,7}$"),          # SPA84363 / SPB5321N62
    ("Pony", r"^\d{2,3}[A-Z]\d[A-Z]{2}\d{2}[A-Z]{2}$"),  # 262U1SA08BE
    ("ASICS", r"^\d{4}[A-Z]\d{6}$"),                 # 1011B881006 / 3033C109101
    ("Native", r"^(\d{12}|F\d{11})$"),               # 111046004201 / F00002329270
    ("REEBOK", r"^100\d{6}$"),                       # 100249414
    ("PALLADIUM", r"^\d{8}$"),                       # 74474008 / 99323080
    ("Teva", r"^(1\d{6}[A-Z]{3,5}|\d{4}[A-Z]{3,4})$"),   # 1124070BCBK / 6154SLA
    ("Converse", r"^(\d{6}C|A\d{5}C|M\d{4}C)$"),     # 162058C / A12975C
    ("SKECHERS", r"^\d{5,6}[A-Z]{1,6}$"),            # 108268BLK / 419030LWNVG
    ("BN3TH", r"^M\d{6}[A-Z0-9]\d{3}$"),             # M2110191739
    ("G.P", r"^([DGP]\d{3,4}[MWBUK]\d{2}|A5\d{4}[A-Z0-9]{2,3}|G95\d{4})$"),  # G3816B35 / A535020
    ("LP SUPPORT", r"^(\d{3}(CA|CN|XT|Z|R1|CAR1)?|IM[57]10|TM\d{3}R)$"),  # 633 / 768CN / 757CAR1
    ("DIADORA", r"^RE\d{5}$"),                       # RE71388（DA 開頭由主檔精確/前綴比對，避免與 NIKE DAxxxx 衝突）
    ("NIKE", r"^([A-Z]{2}\d{7}|[A-Z]{2}\d{7}[A-Z]{2}\d{2,3}|[NJ]\d{10,12}(OS|NS|SL|MD|LG|XL)?|NRN\d+[A-Z]*|NKS\d+[A-Z]*|NNN\d+[A-Z]*|NMS\d+|NER\d+[A-Z]*|NLGB\d+|SX\d{4,7})$"),
    ("adidas", r"^[A-Z]{2}\d{4}$"),                  # EG1758 / JP7149
    ("NEW BALANCE", r"^(7\d{9}F?|396\d{7}F?|LA[BHS]\d+[A-Z]*|AC\d{4,5}[A-Z0-9]+|[MWUPYGLSI][A-Z0-9]{5,9})$"),
]

_COMPILED_RULES: list[tuple[str, re.Pattern]] = [(b, re.compile(p)) for b, p in BRAND_REGEX_RULES]


def normalize_sku(sku: str) -> str:
    return re.sub(r"[\s\-_/]", "", str(sku)).upper()


def clean_product_name(name) -> str:
    if name is None or (isinstance(name, float) and pd.isna(name)):
        return ""
    return str(name).lstrip("\ufffd").strip()


def _brand_from_name(name: str) -> str | None:
    if not isinstance(name, str) or not name.strip():
        return None
    up = f" {name.upper()} "
    best: tuple[int, str] | None = None
    for brand, kws in NAME_KEYWORDS.items():
        for kw in kws:
            pos = up.find(kw)
            if pos >= 0 and (best is None or pos < best[0]):
                best = (pos, brand)
    return best[1] if best else None


def infer_brand(sku: str, product_name: str | None = None, fallback_brand: str | None = None) -> str:
    """由商品名稱關鍵字或商品 ID 正規表示式推斷品牌。"""
    by_name = _brand_from_name(product_name or "")
    if by_name:
        return by_name
    if fallback_brand and fallback_brand not in MIXED_FILES:
        return fallback_brand
    key = normalize_sku(sku)
    for brand, pat in _COMPILED_RULES:
        if pat.match(key):
            return brand
    return fallback_brand or UNCLASSIFIED


def _extract_skus(cell: str, name: str | None = None, expand_variants: bool = True) -> list[str]:
    """從主商品貨號（及商品名稱末尾之同款多色貨號）拆出所有正規化貨號。"""
    tokens = re.findall(r"[A-Za-z0-9][A-Za-z0-9\-]{1,}", str(cell))
    out = [normalize_sku(t) for t in tokens if re.search(r"\d", t)]
    if expand_variants and isinstance(name, str) and out:
        primary = out[0]
        if len(primary) >= 5:
            for t in re.findall(r"[A-Za-z0-9][A-Za-z0-9\-]{4,}", name):
                norm = normalize_sku(t)
                # 同長度且前 2 碼相同（例如 黑KY9952 白KY9953、NKS05010SL NKS05010MD）
                if (len(norm) == len(primary) and norm[:2] == primary[:2]
                        and re.search(r"\d", norm) and norm not in out):
                    out.append(norm)
    return out


def parse_stock_dataframe(
    df: pd.DataFrame,
    fallback_brand: str | None = None,
    expand_variants: bool = True,
) -> tuple[pd.DataFrame, int]:
    """解析商品主檔 DataFrame（如 SHOPPEE STOCK.xlsx），自動排除重複 ID。

    回傳 `(clean_df, duplicate_count)`，其中 `clean_df` 欄位為 `sku, product_name, brand`，
    `duplicate_count` 為原始主商品貨號列的重複筆數。
    """
    if df.empty:
        return pd.DataFrame(columns=["sku", "product_name", "brand"]), 0

    sku_col = df.columns[0]
    name_col = df.columns[1] if len(df.columns) > 1 else None
    brand_col = df.columns[2] if len(df.columns) > 2 else None

    records: dict[str, dict] = {}
    dup_rows = 0
    for _, row in df.iterrows():
        cell = row[sku_col]
        if pd.isna(cell) or not str(cell).strip():
            continue
        raw_name = row[name_col] if name_col is not None else ""
        pname = clean_product_name(raw_name)
        explicit_brand = str(row[brand_col]).strip() if brand_col is not None and pd.notna(row[brand_col]) else None

        skus = _extract_skus(str(cell), name=pname, expand_variants=expand_variants)
        if not skus:
            continue
        primary = skus[0]
        brand = explicit_brand or infer_brand(primary, pname, fallback_brand=fallback_brand)
        if primary in records:
            dup_rows += 1
        for idx, s in enumerate(skus):
            if s not in records:
                records[s] = {"sku": s, "product_name": pname, "brand": brand}
            elif idx == 0 and not records[s]["product_name"] and pname:
                records[s]["product_name"] = pname

    out_df = pd.DataFrame(list(records.values()), columns=["sku", "product_name", "brand"])
    return out_df, dup_rows


def read_stock_excel(file) -> tuple[pd.DataFrame, int]:
    """讀取上傳或本機之商品主檔 Excel（.xlsx / .xls）並去重。"""
    if isinstance(file, (bytes, bytearray)):
        file = io.BytesIO(file)
    df = pd.read_excel(file, dtype=str)
    return parse_stock_dataframe(df, expand_variants=True)


@dataclass
class BrandMatcher:
    sku_to_brand: dict[str, str] = field(default_factory=dict)
    sku_to_name: dict[str, str] = field(default_factory=dict)
    rules: list[tuple[str, re.Pattern]] = field(default_factory=lambda: list(_COMPILED_RULES))
    _sorted_keys: list[str] = field(default_factory=list)

    @classmethod
    def from_dataframe(cls, df: pd.DataFrame) -> "BrandMatcher":
        """由資料庫 `product_catalog` DataFrame（欄位 sku, product_name, brand）建立比對器。"""
        mapping: dict[str, str] = {}
        names: dict[str, str] = {}
        if not df.empty:
            for _, row in df.iterrows():
                key = normalize_sku(row["sku"])
                if not key:
                    continue
                brand = str(row.get("brand") or "").strip() or UNCLASSIFIED
                pname = clean_product_name(row.get("product_name", ""))
                mapping.setdefault(key, brand)
                if pname:
                    names.setdefault(key, pname)
        m = cls(sku_to_brand=mapping, sku_to_name=names, rules=list(_COMPILED_RULES))
        m._sorted_keys = sorted((k for k in mapping if len(k) >= MIN_PREFIX_LEN), key=len, reverse=True)
        return m

    @classmethod
    def from_stock_file(cls, stock_path: str | Path) -> "BrandMatcher":
        df, _ = read_stock_excel(Path(stock_path))
        return cls.from_dataframe(df)

    @classmethod
    def from_directory(cls, brand_dir: str | Path) -> "BrandMatcher":
        brand_dir = Path(brand_dir)
        frames: list[pd.DataFrame] = []
        files = sorted(brand_dir.glob("*.xlsx"), key=lambda p: p.stem in MIXED_FILES)
        for f in files:
            if f.name.startswith("~$"):
                continue
            raw = pd.read_excel(f, dtype=str)
            parsed, _ = parse_stock_dataframe(
                raw,
                fallback_brand=f.stem,
                expand_variants=(f.stem in MIXED_FILES),
            )
            frames.append(parsed)
        combined = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=["sku", "product_name", "brand"])
        return cls.from_dataframe(combined)

    def match(self, sku: str) -> tuple[str, str]:
        """回傳 (brand, method)，method ∈ exact / prefix / regex / none。"""
        key = normalize_sku(sku)
        if not key:
            return UNCLASSIFIED, "none"
        if key in self.sku_to_brand:
            return self.sku_to_brand[key], "exact"
        if len(key) >= MIN_PREFIX_LEN:
            for k in self._sorted_keys:  # 由長至短
                if key.startswith(k) or k.startswith(key):
                    return self.sku_to_brand[k], "prefix"
        for brand, pat in self.rules:
            if pat.match(key):
                return brand, "regex"
        return UNCLASSIFIED, "none"

    def brand_of(self, sku: str) -> str:
        return self.match(sku)[0]

    def name_of(self, sku: str) -> str:
        """從商品主檔查詢商品名稱（精確或前綴比對），未建檔時回傳 '—'。"""
        key = normalize_sku(sku)
        if not key:
            return "—"
        if key in self.sku_to_name:
            return self.sku_to_name[key]
        if len(key) >= MIN_PREFIX_LEN:
            for k in self._sorted_keys:
                if (key.startswith(k) or k.startswith(key)) and k in self.sku_to_name:
                    return self.sku_to_name[k]
        return "—"
