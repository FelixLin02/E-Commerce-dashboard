"""資料庫存取層（Supabase PostgreSQL；本機開發可退回 SQLite）。

所有報表查詢皆在資料庫端以 GROUP BY / SUM 聚合，只回傳輕量結果。
SQL 僅使用 PostgreSQL 與 SQLite 共通語法，方便本機測試。
"""

from __future__ import annotations

from datetime import date, datetime, timedelta

import pandas as pd
from sqlalchemy import (BigInteger, Column, DateTime, Index, Integer, MetaData, Numeric,
                        String, Table, Text, create_engine, func, text)
from sqlalchemy.engine import Engine

from .brand import clean_product_name, infer_brand, normalize_sku

metadata = MetaData()

daily_sales_items = Table(
    "daily_sales_items",
    metadata,
    Column("id", BigInteger().with_variant(Integer, "sqlite"), primary_key=True, autoincrement=True),
    Column("sale_date", DateTime, nullable=False),
    Column("sku", String(64), nullable=False),
    Column("brand", String(64), nullable=False),
    Column("duan_code", String(32)),
    Column("size", String(32), nullable=False),
    Column("quantity", Integer, nullable=False, server_default="1"),
    Column("unit_price", Numeric(10, 2), server_default="0"),
    Column("row_total_amount", Numeric(10, 2), server_default="0"),
    Column("created_at", DateTime, server_default=func.now()),
    Index("idx_sales_date", "sale_date"),
    Index("idx_sales_brand", "brand"),
    Index("idx_sales_sku", "sku"),
    Index("idx_sales_size", "size"),
)

product_catalog = Table(
    "product_catalog",
    metadata,
    Column("sku", String(64), primary_key=True),
    Column("product_name", Text, nullable=False),
    Column("brand", String(64), nullable=False),
    Column("updated_at", DateTime, server_default=func.now()),
    Index("idx_catalog_brand", "brand"),
)

TABLE = "daily_sales_items"
CATALOG_TABLE = "product_catalog"
REVENUE_EXPR = "SUM(quantity * unit_price)"


def make_engine(url: str) -> Engine:
    """建立連線。Supabase 連線字串請使用 Pooler 並加上 sslmode=require。"""
    if url.startswith("postgres://"):
        url = "postgresql://" + url[len("postgres://"):]
    kwargs: dict = {"pool_pre_ping": True}
    if url.startswith("postgresql"):
        if "sslmode=" not in url:
            url += ("&" if "?" in url else "?") + "sslmode=require"
        kwargs.update(pool_size=3, max_overflow=2, pool_recycle=1800)
    engine = create_engine(url, **kwargs)
    metadata.create_all(engine)  # CREATE TABLE / INDEX IF NOT EXISTS
    return engine


def _range_params(start: date, end: date) -> dict:
    """[start, end] 含首尾日期 → 半開區間 [start 00:00, end+1 00:00)。"""
    return {
        "start": datetime.combine(start, datetime.min.time()),
        "end": datetime.combine(end + timedelta(days=1), datetime.min.time()),
    }


def _query(engine: Engine, sql: str, params: dict | None = None) -> pd.DataFrame:
    with engine.connect() as conn:
        return pd.read_sql(text(sql), conn, params=params or {})


# ---------------------------------------------------------------------------
# 匯入
# ---------------------------------------------------------------------------

def count_rows_on_date(engine: Engine, d: date) -> int:
    df = _query(engine, f"SELECT COUNT(*) AS n FROM {TABLE} "
                        "WHERE sale_date >= :start AND sale_date < :end", _range_params(d, d))
    return int(df.iloc[0, 0])


def insert_items(engine: Engine, items: pd.DataFrame, overwrite_dates: bool = True) -> int:
    """寫入明細；overwrite_dates=True 時先刪除同日期既有資料（避免重複匯入）。"""
    if items.empty:
        return 0
    df = items.copy()
    df["sale_date"] = pd.to_datetime(df["sale_date"])
    dates = sorted(df["sale_date"].dt.date.unique())
    # 轉為原生 Python 型別（psycopg2 無法直接處理 numpy.int64）
    records = df.astype(object).where(df.notna(), None).to_dict(orient="records")
    for r in records:
        r["sale_date"] = pd.Timestamp(r["sale_date"]).to_pydatetime()
        r["quantity"] = int(r["quantity"])
        r["unit_price"] = float(r["unit_price"] or 0)
        r["row_total_amount"] = float(r["row_total_amount"] or 0)
    with engine.begin() as conn:
        if overwrite_dates:
            for d in dates:
                conn.execute(text(f"DELETE FROM {TABLE} WHERE sale_date >= :start AND sale_date < :end"),
                             _range_params(d, d))
        conn.execute(daily_sales_items.insert(), records)
    return len(records)


# ---------------------------------------------------------------------------
# 報表查詢
# ---------------------------------------------------------------------------

def kpis(engine: Engine, start: date, end: date) -> dict:
    df = _query(engine, f"""
        SELECT COALESCE({REVENUE_EXPR}, 0) AS revenue,
               COALESCE(SUM(quantity), 0)   AS qty,
               COUNT(DISTINCT sku)          AS sku_count,
               COUNT(DISTINCT brand)        AS brand_count
        FROM {TABLE}
        WHERE sale_date >= :start AND sale_date < :end
    """, _range_params(start, end))
    r = df.iloc[0]
    return {k: float(r[k] or 0) for k in df.columns}


def brand_summary(engine: Engine, start: date, end: date) -> pd.DataFrame:
    return _query(engine, f"""
        SELECT brand, SUM(quantity) AS qty, {REVENUE_EXPR} AS revenue, COUNT(DISTINCT sku) AS sku_count
        FROM {TABLE}
        WHERE sale_date >= :start AND sale_date < :end
        GROUP BY brand
        ORDER BY revenue DESC
    """, _range_params(start, end))


def daily_trend(engine: Engine, start: date, end: date) -> pd.DataFrame:
    df = _query(engine, f"""
        SELECT DATE(sale_date) AS day, SUM(quantity) AS qty, {REVENUE_EXPR} AS revenue
        FROM {TABLE}
        WHERE sale_date >= :start AND sale_date < :end
        GROUP BY DATE(sale_date)
        ORDER BY day
    """, _range_params(start, end))
    if not df.empty:
        df["day"] = pd.to_datetime(df["day"])
    return df


def top_products(engine: Engine, start: date, end: date, limit: int = 10,
                 brand: str | None = None, order_by: str = "qty") -> pd.DataFrame:
    order_col = "revenue" if order_by == "revenue" else "qty"
    params = _range_params(start, end) | {"limit": int(limit)}
    brand_sql = ""
    if brand:
        brand_sql = "AND brand = :brand"
        params["brand"] = brand
    return _query(engine, f"""
        SELECT sku, brand, SUM(quantity) AS qty, {REVENUE_EXPR} AS revenue
        FROM {TABLE}
        WHERE sale_date >= :start AND sale_date < :end {brand_sql}
        GROUP BY sku, brand
        ORDER BY {order_col} DESC, sku
        LIMIT :limit
    """, params)


def size_distribution(engine: Engine, start: date, end: date, sku: str) -> pd.DataFrame:
    return _query(engine, f"""
        SELECT duan_code, size, SUM(quantity) AS qty, {REVENUE_EXPR} AS revenue
        FROM {TABLE}
        WHERE sale_date >= :start AND sale_date < :end AND sku = :sku
        GROUP BY duan_code, size
    """, _range_params(start, end) | {"sku": sku})


def search_skus(engine: Engine, start: date, end: date, keyword: str, limit: int = 50) -> pd.DataFrame:
    return _query(engine, f"""
        SELECT sku, brand, SUM(quantity) AS qty
        FROM {TABLE}
        WHERE sale_date >= :start AND sale_date < :end AND UPPER(sku) LIKE :kw
        GROUP BY sku, brand
        ORDER BY qty DESC
        LIMIT :limit
    """, _range_params(start, end) | {"kw": f"%{keyword.upper()}%", "limit": int(limit)})


def brand_size_matrix(engine: Engine, start: date, end: date, brand: str | None = None) -> pd.DataFrame:
    params = _range_params(start, end)
    brand_sql = ""
    if brand:
        brand_sql = "AND brand = :brand"
        params["brand"] = brand
    return _query(engine, f"""
        SELECT duan_code, size, SUM(quantity) AS qty
        FROM {TABLE}
        WHERE sale_date >= :start AND sale_date < :end {brand_sql}
        GROUP BY duan_code, size
    """, params)


# ---------------------------------------------------------------------------
# 資料維護
# ---------------------------------------------------------------------------

def import_batches(engine: Engine) -> pd.DataFrame:
    df = _query(engine, f"""
        SELECT DATE(sale_date) AS day, COUNT(*) AS item_rows, COUNT(DISTINCT sku) AS sku_count,
               SUM(quantity) AS qty, {REVENUE_EXPR} AS revenue, MAX(created_at) AS imported_at
        FROM {TABLE}
        GROUP BY DATE(sale_date)
        ORDER BY day DESC
    """)
    if not df.empty:
        df["day"] = pd.to_datetime(df["day"]).dt.date
    return df


def delete_date(engine: Engine, d: date) -> int:
    with engine.begin() as conn:
        res = conn.execute(text(f"DELETE FROM {TABLE} WHERE sale_date >= :start AND sale_date < :end"),
                           _range_params(d, d))
        return res.rowcount or 0


def distinct_skus_by_brand(engine: Engine, brand: str | None = None) -> pd.DataFrame:
    if brand:
        return _query(engine, f"SELECT sku, brand, SUM(quantity) AS qty FROM {TABLE} "
                              "WHERE brand = :brand GROUP BY sku, brand ORDER BY qty DESC", {"brand": brand})
    return _query(engine, f"SELECT sku, brand, SUM(quantity) AS qty FROM {TABLE} "
                          "GROUP BY sku, brand ORDER BY qty DESC")


def update_brand(engine: Engine, sku: str, brand: str) -> int:
    with engine.begin() as conn:
        res = conn.execute(text(f"UPDATE {TABLE} SET brand = :brand WHERE sku = :sku"),
                           {"brand": brand, "sku": sku})
        return res.rowcount or 0


def list_brands(engine: Engine) -> list[str]:
    df = _query(engine, f"SELECT DISTINCT brand FROM {TABLE} ORDER BY brand")
    return df["brand"].tolist()


# ---------------------------------------------------------------------------
# 商品主檔 (product_catalog) CRUD 與批次匯入
# ---------------------------------------------------------------------------

def catalog_count(engine: Engine) -> int:
    df = _query(engine, f"SELECT COUNT(*) AS n FROM {CATALOG_TABLE}")
    return int(df.iloc[0, 0])


def load_catalog(engine: Engine) -> pd.DataFrame:
    return _query(engine, f"SELECT sku, product_name, brand, updated_at FROM {CATALOG_TABLE} ORDER BY brand, sku")


def search_catalog(engine: Engine, keyword: str = "", brand: str | None = None, limit: int = 500) -> pd.DataFrame:
    conds = ["1=1"]
    params: dict = {"limit": int(limit)}
    if brand:
        conds.append("brand = :brand")
        params["brand"] = brand
    if keyword.strip():
        conds.append("(UPPER(sku) LIKE :kw OR UPPER(product_name) LIKE :kw_raw)")
        params["kw"] = f"%{normalize_sku(keyword)}%"
        params["kw_raw"] = f"%{keyword.strip().upper()}%"
    where = " AND ".join(conds)
    return _query(engine, f"""
        SELECT sku, product_name, brand, updated_at
        FROM {CATALOG_TABLE}
        WHERE {where}
        ORDER BY updated_at DESC, sku
        LIMIT :limit
    """, params)


def add_catalog_item(
    engine: Engine,
    sku: str,
    product_name: str,
    brand: str | None = None,
    overwrite: bool = False,
) -> tuple[bool, str]:
    """逐一新增商品主檔。

    回傳 `(changed, status)`，status ∈ {'inserted', 'updated', 'duplicate'}。
    當 `overwrite=False` 且商品 ID 已存在時，自動排除重複不覆蓋。
    """
    norm = normalize_sku(sku)
    if not norm:
        raise ValueError("商品 ID 不可為空")
    pname = clean_product_name(product_name)
    resolved_brand = (brand or "").strip() or infer_brand(norm, pname)
    now = datetime.now()

    exists_df = _query(engine, f"SELECT sku FROM {CATALOG_TABLE} WHERE sku = :sku", {"sku": norm})
    if not exists_df.empty:
        if not overwrite:
            return False, "duplicate"
        with engine.begin() as conn:
            conn.execute(
                text(f"UPDATE {CATALOG_TABLE} SET product_name = :pname, brand = :brand, updated_at = :now WHERE sku = :sku"),
                {"pname": pname, "brand": resolved_brand, "now": now, "sku": norm},
            )
        return True, "updated"

    with engine.begin() as conn:
        conn.execute(
            product_catalog.insert(),
            [{"sku": norm, "product_name": pname, "brand": resolved_brand, "updated_at": now}],
        )
    return True, "inserted"


def delete_catalog_item(engine: Engine, sku: str) -> int:
    norm = normalize_sku(sku)
    with engine.begin() as conn:
        res = conn.execute(text(f"DELETE FROM {CATALOG_TABLE} WHERE sku = :sku"), {"sku": norm})
        return res.rowcount or 0


def import_catalog(engine: Engine, clean_df: pd.DataFrame, mode: str = "append") -> dict:
    """批次匯入商品主檔（clean_df 需含 sku, product_name, brand 且已排除檔案內重複）。

    - mode='append'：批次匯入新增（自動排除資料庫中已存在的重複 ID，僅寫入新 ID）
    - mode='replace'：整批匯入覆蓋（清空舊主檔，寫入去重後的全部品項）
    """
    now = datetime.now()
    df = clean_df[["sku", "product_name", "brand"]].copy()
    df["sku"] = df["sku"].map(normalize_sku)
    df = df[df["sku"] != ""].drop_duplicates(subset=["sku"], keep="first")

    if mode == "replace":
        records = [
            {
                "sku": str(r["sku"]),
                "product_name": clean_product_name(r["product_name"]),
                "brand": str(r["brand"] or "").strip() or infer_brand(str(r["sku"]), str(r["product_name"])),
                "updated_at": now,
            }
            for _, r in df.iterrows()
        ]
        with engine.begin() as conn:
            conn.execute(text(f"DELETE FROM {CATALOG_TABLE}"))
            if records:
                conn.execute(product_catalog.insert(), records)
        return {"inserted": len(records), "skipped_existing": 0, "total_in_file": len(df), "mode": "replace"}

    # mode == "append"：排除資料庫既有 ID
    existing_df = _query(engine, f"SELECT sku FROM {CATALOG_TABLE}")
    existing_set = set(existing_df["sku"].map(normalize_sku)) if not existing_df.empty else set()
    new_df = df[~df["sku"].isin(existing_set)]
    records = [
        {
            "sku": str(r["sku"]),
            "product_name": clean_product_name(r["product_name"]),
            "brand": str(r["brand"] or "").strip() or infer_brand(str(r["sku"]), str(r["product_name"])),
            "updated_at": now,
        }
        for _, r in new_df.iterrows()
    ]
    if records:
        with engine.begin() as conn:
            conn.execute(product_catalog.insert(), records)
    return {
        "inserted": len(records),
        "skipped_existing": len(df) - len(new_df),
        "total_in_file": len(df),
        "mode": "append",
    }
