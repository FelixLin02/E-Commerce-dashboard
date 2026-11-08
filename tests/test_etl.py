from datetime import date, datetime
from pathlib import Path

import pytest

from etl import UNCLASSIFIED, BrandMatcher, db, parse_erp
from etl.size_codes import SIZE_CODES, SIZE_SLOTS, display_size, normalize_label, resolve_size

ROOT = Path(__file__).resolve().parents[1]
SAMPLE = ROOT / "Sample" / "1103銷單.xls"


@pytest.fixture(scope="session")
def matcher():
    return BrandMatcher.from_directory(ROOT / "品牌")


@pytest.fixture(scope="session")
def result(matcher):
    if not SAMPLE.exists():
        pytest.skip("sample file not available")
    return parse_erp(SAMPLE, matcher)


# --- size codes ---------------------------------------------------------------

def test_all_codes_have_17_slots():
    for code, labels in SIZE_CODES.items():
        assert len(labels) == SIZE_SLOTS, code


@pytest.mark.parametrize("code,slot,expected", [
    ("1", 0, "F"), ("2", 0, "4"), ("2", 9, "8-"), ("2", 16, "12"),
    ("3", 5, "2XL"), ("4", 16, "30"), ("7", 16, "50"), ("A", 11, "12-"),
    ("C", 16, "10"), ("D", 11, "11"), ("D", 16, "13.5"), ("H", 16, "11"),
    ("I", 0, "11.5"), ("J", 12, "7Y"), ("N", 7, "714"), ("O", 12, "21-"),
])
def test_resolve_size(code, slot, expected):
    assert resolve_size(code, slot) == expected


def test_unknown_code_placeholder():
    assert resolve_size("6", 8) == "段碼6-第9欄"


def test_label_helpers():
    assert normalize_label("-4") == "4-"
    assert normalize_label("-10.0") == "10-"
    assert display_size("8-") == "8.5"
    assert display_size("XL") == "XL"


# --- brand ----------------------------------------------------------------------

@pytest.mark.parametrize("sku,brand", [
    ("1011B977004", "ASICS"),          # exact
    ("AC4562WBKF", "NEW BALANCE"),     # prefix (配件主檔 AC4562WBK)
    ("DB0456010", "NIKE"),             # regex
    ("EG1758", "adidas"),
    ("74474008", "PALLADIUM"),
    ("108268BLK", "SKECHERS"),
    ("3-J414A-188", "FILA"),
    ("4135045-2022W", "Havaianas"),
    ("111046004201", "Native"),
    ("M2110191739", "BN3TH"),
    ("1019180BNGC", "Teva"),           # 放錯檔(BN3TH)但名稱為 TEVA
    ("633", "LP SUPPORT"),
    ("768CN", "LP SUPPORT"),
    ("4255162245", "EVERLAST"),
    ("08526133", "K-SWISS"),
    ("DA7617230", "DIADORA"),
    ("UNKNOWN_SKU_XYZ", UNCLASSIFIED),
])
def test_brand(matcher, sku, brand):
    assert matcher.brand_of(sku) == brand


def test_product_name_lookup(matcher):
    assert "ASICS" in matcher.name_of("1011B977004")
    assert "New Balance" in matcher.name_of("AC4562WBKF")  # prefix match
    assert matcher.name_of("UNKNOWN_SKU_XYZ") == "—"


# --- parser ---------------------------------------------------------------------

def test_sample_totals_match_report(result):
    assert result.file_total_qty == 263
    assert result.items["quantity"].sum() == 263
    assert result.rows["合計金額"].sum() == result.file_total_amount == 499216


def test_sample_header_info(result):
    assert result.report_date == datetime(2026, 11, 3)
    assert result.period_start == result.period_end == datetime(2026, 10, 27)
    assert result.file_size_codes["2"] == SIZE_CODES["2"]


def test_sample_no_header_rows_leak(result):
    assert not result.items["sku"].isin(["商品型號", "合計", ""]).any()
    assert len(result.rows) == 196


def test_sample_specific_rows(result):
    it = result.items
    r = it[it["sku"] == "111001001999"].sort_values("size")
    assert sorted(r["size"]) == ["11", "6", "9"]
    assert (r["unit_price"].round(2) == round(4898 / 3, 2)).all()
    assert it[it["sku"] == "1014A338009"]["size"].tolist() == ["21-"]
    assert it[it["sku"] == "DR2695002"]["size"].tolist() == ["11"]
    assert it[it["sku"] == "N1009194205OS"]["quantity"].tolist() == [5]


# --- db round trip (SQLite) -----------------------------------------------------

def test_db_roundtrip(result, tmp_path):
    engine = db.make_engine(f"sqlite:///{(tmp_path / 't.sqlite').as_posix()}")
    items = result.items.copy()
    items["sale_date"] = datetime(2026, 10, 27)
    assert db.insert_items(engine, items) == len(items)
    db.insert_items(engine, items)  # 重複匯入 → 覆蓋，不應倍增
    d = date(2026, 10, 27)
    k = db.kpis(engine, d, d)
    assert k["qty"] == 263
    assert abs(k["revenue"] - 499216) < 5  # 單價四捨五入誤差
    assert db.brand_summary(engine, d, d)["qty"].sum() == 263
    top = db.top_products(engine, d, d, 10)
    assert len(top) == 10 and top["qty"].is_monotonic_decreasing
    assert db.size_distribution(engine, d, d, "111001001999")["qty"].sum() == 3
    assert db.count_rows_on_date(engine, d) == len(items)
    assert db.delete_date(engine, d) == len(items)
    assert db.kpis(engine, d, d)["qty"] == 0


def test_product_catalog_crud_and_import(tmp_path):
    import pandas as pd
    from etl.brand import parse_stock_dataframe

    engine = db.make_engine(f"sqlite:///{(tmp_path / 'cat.sqlite').as_posix()}")
    assert db.catalog_count(engine) == 0

    # 1. 逐一新增 & 排除重複
    ok, status = db.add_catalog_item(engine, "NEW-001", "NIKE 測試新款跑鞋 NEW-001")
    assert ok and status == "inserted"
    ok2, status2 = db.add_catalog_item(engine, "NEW001", "重複ID不應覆蓋", overwrite=False)
    assert not ok2 and status2 == "duplicate"
    ok3, status3 = db.add_catalog_item(engine, "NEW001", "NIKE 已更新名稱", overwrite=True)
    assert ok3 and status3 == "updated"

    # 2. 批次匯入新增（含檔案內重複 + 資料庫既有重複）
    raw = pd.DataFrame({
        "主商品貨號": ["NEW001", "1011B977004", "1011B977-004", "HQ2448"],
        "商品名稱": ["應被排除", "ASICS 慢跑鞋", "檔案內重複應被排除", "adidas 拖鞋"],
    })
    clean_df, file_dups = parse_stock_dataframe(raw)
    assert file_dups == 1 and len(clean_df) == 3
    res_append = db.import_catalog(engine, clean_df, mode="append")
    assert res_append["inserted"] == 2 and res_append["skipped_existing"] == 1
    assert db.catalog_count(engine) == 3

    # 3. 從資料庫建立 BrandMatcher 查詢商品名稱與品牌
    m = BrandMatcher.from_dataframe(db.load_catalog(engine))
    assert m.name_of("NEW-001") == "NIKE 已更新名稱"
    assert m.brand_of("NEW-001") == "NIKE"
    assert m.name_of("1011B977004") == "ASICS 慢跑鞋"

    # 4. 整批匯入覆蓋
    res_replace = db.import_catalog(engine, clean_df.iloc[:2], mode="replace")
    assert res_replace["inserted"] == 2 and db.catalog_count(engine) == 2
    assert db.delete_catalog_item(engine, "NEW001") == 1
    assert db.catalog_count(engine) == 1

