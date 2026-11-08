"""京鴻 ERP 運動鞋服銷售數據分析與樞紐報表系統（Streamlit 前端）。

執行：streamlit run app.py
資料庫連線：Streamlit Secrets 之 DATABASE_URL（Supabase Pooler 連線字串）。
未設定時自動使用本機 SQLite（local_dev.sqlite），僅供開發測試。
"""

from __future__ import annotations

import hmac
import os
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

from etl import UNCLASSIFIED, BrandMatcher, db, parse_erp
from etl.brand import NAME_KEYWORDS, infer_brand, normalize_sku, read_stock_excel
from etl.size_codes import display_size, size_sort_key

ROOT = Path(__file__).resolve().parent
BRAND_DIR = ROOT / "品牌"
STOCK_FILE = ROOT / "補充" / "SHOPPEE STOCK.xlsx"
LOCAL_DB = ROOT / "local_dev.sqlite"

st.set_page_config(page_title="你我商行_銷售分析", page_icon="👟", layout="wide")


# ---------------------------------------------------------------------------
# 資源 / 設定
# ---------------------------------------------------------------------------

def _secret(key: str, default=None):
    try:
        return st.secrets.get(key, default)
    except Exception:  # 本機沒有 secrets.toml
        return default


def _database_url() -> tuple[str, bool]:
    url = _secret("DATABASE_URL") or os.environ.get("DATABASE_URL")
    if url:
        return url, False
    return f"sqlite:///{LOCAL_DB.as_posix()}", True


@st.cache_resource(show_spinner="連線資料庫中…")
def get_engine(url: str):
    engine = db.make_engine(url)
    _ensure_catalog_seeded(engine)
    return engine


def _ensure_catalog_seeded(engine) -> None:
    """若資料庫 `product_catalog` 為空，自動由 SHOPPEE STOCK.xlsx（或 品牌/*.xlsx）初始化匯入。"""
    if db.catalog_count(engine) > 0:
        return
    if STOCK_FILE.exists():
        clean_df, _ = read_stock_excel(STOCK_FILE)
        db.import_catalog(engine, clean_df, mode="replace")
    elif BRAND_DIR.exists():
        m = BrandMatcher.from_directory(BRAND_DIR)
        rows = [
            {"sku": k, "product_name": m.sku_to_name.get(k, ""), "brand": b}
            for k, b in m.sku_to_brand.items()
        ]
        if rows:
            db.import_catalog(engine, pd.DataFrame(rows), mode="replace")


@st.cache_data(ttl=600, show_spinner="載入商品主檔資料庫中…")
def _load_catalog_df(url: str) -> pd.DataFrame:
    engine = get_engine(url)
    _ensure_catalog_seeded(engine)
    return db.load_catalog(engine)


def get_matcher(url: str) -> BrandMatcher:
    return BrandMatcher.from_dataframe(_load_catalog_df(url))


def check_password() -> bool:
    """若 Secrets 設有 APP_PASSWORD，需輸入密碼才能使用。"""
    pwd = _secret("APP_PASSWORD")
    if not pwd or st.session_state.get("auth_ok"):
        return True
    with st.form("login"):
        st.subheader("🔒 請輸入存取密碼")
        entered = st.text_input("密碼", type="password")
        if st.form_submit_button("登入"):
            if hmac.compare_digest(entered, str(pwd)):
                st.session_state["auth_ok"] = True
                st.rerun()
            st.error("密碼錯誤")
    return False


# ---------------------------------------------------------------------------
# 快取查詢（匯入或修改資料後呼叫 clear_cache()）
# ---------------------------------------------------------------------------

@st.cache_data(ttl=600, show_spinner=False)
def q_kpis(url, s, e):
    return db.kpis(get_engine(url), s, e)


@st.cache_data(ttl=600, show_spinner=False)
def q_brand(url, s, e):
    return db.brand_summary(get_engine(url), s, e)


@st.cache_data(ttl=600, show_spinner=False)
def q_trend(url, s, e):
    return db.daily_trend(get_engine(url), s, e)


@st.cache_data(ttl=600, show_spinner=False)
def q_top(url, s, e, n, brand, order_by):
    return db.top_products(get_engine(url), s, e, n, brand, order_by)


@st.cache_data(ttl=600, show_spinner=False)
def q_size(url, s, e, sku):
    return db.size_distribution(get_engine(url), s, e, sku)


def clear_cache():
    for f in (q_kpis, q_brand, q_trend, q_top, q_size, _load_catalog_df):
        f.clear()


def fmt_money(v: float) -> str:
    return f"${v:,.0f}"


def sort_sizes(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["_k"] = [size_sort_key(c, s) for c, s in zip(df["duan_code"], df["size"])]
    df = df.sort_values(["duan_code", "_k"]).drop(columns="_k")
    df["尺寸"] = [display_size(s) if c != "N" else s for c, s in zip(df["duan_code"], df["size"])]
    return df


# ---------------------------------------------------------------------------
# 頁面：匯入 ERP 銷售檔
# ---------------------------------------------------------------------------

def page_import(url: str):
    st.header("📥 京鴻 ERP 銷售檔匯入")
    st.caption("上傳 ERP「商品銷售總表列印」匯出之 Excel（.xls / .xlsx），系統自動清洗跨頁表頭、透過商品主檔資料庫判定品牌與名稱，並將 17 欄尺寸轉置為明細。")

    files = st.file_uploader("拖曳或選擇檔案（可多選）", type=["xls", "xlsx"], accept_multiple_files=True)
    if not files:
        return

    matcher = get_matcher(url)
    engine = get_engine(url)
    for idx, f in enumerate(files):
        with st.expander(f"📄 {f.name}", expanded=True):
            try:
                t0 = datetime.now()
                res = parse_erp(f, matcher)
                elapsed = (datetime.now() - t0).total_seconds()
            except Exception as ex:  # noqa: BLE001
                st.error(f"解析失敗：{ex}")
                continue

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("型號數", f"{len(res.rows):,}")
            c2.metric("尺寸明細筆數", f"{len(res.items):,}")
            c3.metric("總件數", f"{res.items['quantity'].sum():,}")
            c4.metric("總金額", fmt_money(res.rows["合計金額"].sum()) if not res.rows.empty else "$0")
            info = []
            if res.period_start:
                info.append(f"日期區間：{res.period_start:%Y-%m-%d} ~ {res.period_end:%Y-%m-%d}")
            if res.report_date:
                info.append(f"製表日期：{res.report_date:%Y-%m-%d}")
            if res.file_total_qty is not None:
                info.append(f"報表合計：{res.file_total_qty:,.0f} 件 / {fmt_money(res.file_total_amount or 0)}")
            info.append(f"解析耗時 {elapsed:.2f} 秒")
            st.caption("　｜　".join(info))

            for w in res.warnings:
                st.warning(w)

            # 銷售日期選擇
            options = {}
            if res.period_start:
                options[f"日期區間起日（{res.period_start:%Y-%m-%d}）"] = res.period_start.date()
            if res.report_date:
                options[f"製表日期（{res.report_date:%Y-%m-%d}）"] = res.report_date.date()
            options["手動指定"] = None
            choice = st.radio("寫入之銷售日期", list(options), horizontal=True, key=f"date_choice_{idx}")
            sale_day = options[choice]
            if sale_day is None:
                sale_day = st.date_input("指定銷售日期", value=date.today(), key=f"date_manual_{idx}")

            tab1, tab2 = st.tabs(["型號彙總", "尺寸明細（將寫入資料庫）"])
            with tab1:
                st.dataframe(res.rows, width="stretch", hide_index=True)
            with tab2:
                st.dataframe(res.items.drop(columns="sale_date"), width="stretch", hide_index=True)

            existing = db.count_rows_on_date(engine, sale_day)
            overwrite = True
            if existing:
                st.info(f"資料庫已有 {sale_day:%Y-%m-%d} 共 {existing} 筆明細。")
                overwrite = st.checkbox("覆蓋該日既有資料（建議，避免重複匯入）", value=True, key=f"ow_{idx}")

            if st.button("✅ 確認寫入資料庫", type="primary", key=f"commit_{idx}"):
                items = res.items.copy()
                items["sale_date"] = datetime.combine(sale_day, datetime.min.time())
                with st.spinner("寫入中…"):
                    n = db.insert_items(engine, items, overwrite_dates=overwrite)
                clear_cache()
                st.success(f"已寫入 {n} 筆明細（銷售日期 {sale_day:%Y-%m-%d}）。")


# ---------------------------------------------------------------------------
# 商品主檔管理區塊（併入「🛠️ 資料維護」頁籤）
# ---------------------------------------------------------------------------

def _render_catalog_section(url: str, engine):
    st.subheader("📦 商品主檔資料庫管理")
    st.caption("基於 `SHOPPEE STOCK.xlsx` 建立之商品主檔資料庫（`product_catalog`），提供由商品 ID 搜尋商品名稱與品牌，支援單筆新增、批次新增與整批覆蓋，並自動排除重複 ID。")

    cat_df = _load_catalog_df(url)
    c1, c2 = st.columns(2)
    c1.metric("目前主檔總品項數（不重複 ID）", f"{len(cat_df):,}")
    c2.metric("涵蓋品牌數", f"{cat_df['brand'].nunique():,}" if not cat_df.empty else "0")

    tab_single, tab_batch, tab_browse = st.tabs([
        "➕ 逐一新增品項",
        "📤 批次匯入（新增／覆蓋）",
        "🔍 搜尋與管理主檔",
    ])

    # --- Tab 1：逐一新增 ---
    with tab_single:
        known_brands = sorted(set(NAME_KEYWORDS.keys()) | set(cat_df["brand"].unique().tolist()))
        with st.form("single_add_form", clear_on_submit=False):
            col1, col2 = st.columns([1, 2])
            sku_in = col1.text_input("商品 ID（主商品貨號）*", placeholder="例如：1011B977004")
            name_in = col2.text_input("商品名稱*", placeholder="例如：ASICS GEL-SONOMA 8 GTX 男生 黑色 慢跑鞋 1011B977-004")
            col3, col4 = st.columns([1, 1])
            brand_choice = col3.selectbox("品牌", ["⚡ 自動依名稱／貨號判別"] + known_brands)
            dup_policy = col4.radio(
                "當商品 ID 已存在時",
                ["🚫 排除重複（略過不覆蓋）", "✏️ 更新既有品項"],
                horizontal=True,
            )
            submitted = st.form_submit_button("💾 寫入商品主檔", type="primary")
            if submitted:
                norm_sku = normalize_sku(sku_in)
                if not norm_sku or not name_in.strip():
                    st.error("請填寫「商品 ID」與「商品名稱」。")
                else:
                    explicit_brand = None if brand_choice.startswith("⚡") else brand_choice
                    resolved_brand = explicit_brand or infer_brand(norm_sku, name_in.strip())
                    overwrite = dup_policy.startswith("✏️")
                    changed, status = db.add_catalog_item(
                        engine, norm_sku, name_in.strip(), resolved_brand, overwrite=overwrite
                    )
                    if status == "duplicate":
                        st.warning(f"商品 ID `{norm_sku}` 已存在於主檔中，已自動排除重複（未覆蓋）。")
                    elif status == "updated":
                        clear_cache()
                        st.success(f"已更新商品 `{norm_sku}`（品牌：{resolved_brand}）。")
                        st.rerun()
                    else:
                        clear_cache()
                        st.success(f"已新增商品 `{norm_sku}`（品牌：{resolved_brand}）。")
                        st.rerun()

    # --- Tab 2：批次匯入（新增 / 覆蓋） ---
    with tab_batch:
        st.caption("上傳格式與 `SHOPPEE STOCK.xlsx` 相同（第 1 欄：主商品貨號、第 2 欄：商品名稱）。系統會先自動排除檔案內重複的商品 ID。")

        mode_label = st.radio(
            "選擇匯入模式",
            [
                "➕ 批次匯入新增（僅新增資料庫尚未存在的 ID，自動排除重複 ID）",
                "🔄 整批匯入覆蓋（清空舊主檔，以新檔案去重後完整覆蓋）",
            ],
            key="cat_import_mode",
        )
        mode = "append" if mode_label.startswith("➕") else "replace"

        up_file = st.file_uploader("選擇商品主檔 Excel（.xlsx / .xls）", type=["xlsx", "xls"], key="cat_uploader")
        if up_file is not None:
            try:
                clean_df, file_dups = read_stock_excel(up_file)
            except Exception as ex:  # noqa: BLE001
                st.error(f"讀取檔案失敗：{ex}")
                clean_df = pd.DataFrame()
                file_dups = 0

            if not clean_df.empty:
                existing_set = set(cat_df["sku"].map(normalize_sku)) if not cat_df.empty else set()
                in_db_count = int(clean_df["sku"].isin(existing_set).sum())
                new_count = len(clean_df) - in_db_count

                m1, m2, m3, m4 = st.columns(4)
                m1.metric("檔案去重後有效 ID 數", f"{len(clean_df):,}")
                m2.metric("已排除檔案內重複列", f"{file_dups:,}")
                if mode == "append":
                    m3.metric("資料庫已存在（將排除）", f"{in_db_count:,}")
                    m4.metric("預計新增筆數", f"{new_count:,}")
                else:
                    m3.metric("原主檔筆數（將被覆蓋）", f"{len(cat_df):,}")
                    m4.metric("預計寫入總筆數", f"{len(clean_df):,}")

                st.dataframe(
                    clean_df.rename(columns={"sku": "商品ID", "product_name": "商品名稱", "brand": "判別品牌"}).head(100),
                    hide_index=True,
                    width="stretch",
                )
                if st.button("✅ 確認執行批次匯入", type="primary", key="btn_cat_import"):
                    with st.spinner("匯入商品主檔中…"):
                        res = db.import_catalog(engine, clean_df, mode=mode)
                    clear_cache()
                    if mode == "append":
                        st.success(
                            f"批次匯入新增完成：新增 **{res['inserted']:,}** 筆，"
                            f"自動排除已存在重複 ID **{res['skipped_existing']:,}** 筆（檔案內重複已排除 {file_dups:,} 筆）。"
                        )
                    else:
                        st.success(
                            f"整批匯入覆蓋完成：共寫入 **{res['inserted']:,}** 筆不重複商品（檔案內重複已排除 {file_dups:,} 筆）。"
                        )
                    st.rerun()

    # --- Tab 3：搜尋與管理 ---
    with tab_browse:
        s1, s2 = st.columns([2, 1])
        kw = s1.text_input("搜尋商品 ID 或商品名稱關鍵字", key="cat_search_kw")
        b_list = ["全部品牌"] + sorted(cat_df["brand"].unique().tolist()) if not cat_df.empty else ["全部品牌"]
        b_filter = s2.selectbox("篩選品牌", b_list, key="cat_search_brand")

        found = db.search_catalog(
            engine,
            keyword=kw,
            brand=None if b_filter == "全部品牌" else b_filter,
            limit=500,
        )
        st.caption(f"顯示前 {len(found):,} 筆符合結果（主檔共 {len(cat_df):,} 筆）")
        st.dataframe(
            found.rename(columns={"sku": "商品ID", "product_name": "商品名稱", "brand": "品牌", "updated_at": "更新時間"}),
            hide_index=True,
            width="stretch",
        )

        with st.expander("🗑️ 刪除指定商品 ID"):
            with st.form("del_sku_form"):
                del_sku = st.text_input("輸入欲刪除之商品 ID")
                if st.form_submit_button("刪除品項", type="secondary") and del_sku.strip():
                    n = db.delete_catalog_item(engine, del_sku.strip())
                    if n:
                        clear_cache()
                        st.success(f"已刪除商品 `{normalize_sku(del_sku)}`。")
                        st.rerun()
                    else:
                        st.warning("查無此商品 ID。")


# ---------------------------------------------------------------------------
# 頁面：報表
# ---------------------------------------------------------------------------

def date_range_picker() -> tuple[date, date]:
    today = date.today()
    mode = st.sidebar.radio("時間區間", ["單日", "本週", "本月", "前 30 天", "上個月", "自訂區間"], index=3)
    if mode == "單日":
        d = st.sidebar.date_input("日期", value=today)
        return d, d
    if mode == "本週":
        s = today - timedelta(days=today.weekday())
        return s, today
    if mode == "本月":
        return today.replace(day=1), today
    if mode == "前 30 天":
        return today - timedelta(days=29), today
    if mode == "上個月":
        e = today.replace(day=1) - timedelta(days=1)
        return e.replace(day=1), e
    rng = st.sidebar.date_input("起訖日期", value=(today.replace(day=1), today))
    if isinstance(rng, (list, tuple)):
        if len(rng) == 2:
            return rng[0], rng[1]
        if len(rng) == 1:
            return rng[0], rng[0]
    return rng, rng


def page_report(url: str):
    s, e = date_range_picker()
    st.header("📊 營運報表")
    st.caption(f"查詢區間：{s:%Y-%m-%d} ~ {e:%Y-%m-%d}")

    k = q_kpis(url, s, e)
    if k["qty"] == 0:
        st.info("此區間尚無銷售資料。請調整時間區間，或至「匯入資料」上傳 ERP 檔案。")
        return

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("總銷售額", fmt_money(k["revenue"]))
    c2.metric("總銷售件數", f"{k['qty']:,.0f}")
    c3.metric("涵蓋品項數（型號）", f"{k['sku_count']:,.0f}")
    c4.metric("平均件單價", fmt_money(k["revenue"] / k["qty"]) if k["qty"] else "-")

    trend = q_trend(url, s, e)
    if len(trend) > 1:
        fig = px.bar(trend, x="day", y="revenue", hover_data=["qty"],
                     labels={"day": "日期", "revenue": "銷售額", "qty": "件數"}, title="每日銷售額")
        fig.update_layout(height=300, margin=dict(t=40, b=10))
        st.plotly_chart(fig, width="stretch")

    # --- 品牌樞紐 ---
    st.subheader("🏷️ 品牌樞紐分析")
    bdf = q_brand(url, s, e)
    col_a, col_b = st.columns([3, 2])
    with col_a:
        show = bdf.copy()
        show["營收佔比"] = show["revenue"] / show["revenue"].sum()
        show["件數佔比"] = show["qty"] / show["qty"].sum()
        show.insert(0, "排名", range(1, len(show) + 1))
        st.dataframe(
            show.rename(columns={"brand": "品牌", "qty": "銷售件數", "revenue": "銷售額", "sku_count": "品項數"}),
            hide_index=True, width="stretch",
            column_config={
                "銷售額": st.column_config.NumberColumn(format="$%d"),
                "營收佔比": st.column_config.ProgressColumn(format="percent", min_value=0, max_value=1),
                "件數佔比": st.column_config.NumberColumn(format="percent"),
            },
        )
    with col_b:
        metric = st.radio("圓餅圖指標", ["銷售額", "銷售件數"], horizontal=True, key="pie_metric")
        fig = px.pie(bdf, names="brand", values="revenue" if metric == "銷售額" else "qty", hole=0.45)
        fig.update_traces(textposition="inside", textinfo="percent+label")
        fig.update_layout(height=380, margin=dict(t=10, b=10), showlegend=False)
        st.plotly_chart(fig, width="stretch")

    # --- 熱銷排行 ---
    st.subheader("🔥 熱門商品銷售排行")
    f1, f2, f3 = st.columns([2, 1, 1])
    brand_opt = f1.selectbox("品牌", ["全館"] + bdf["brand"].tolist())
    top_n = f2.selectbox("排行數", [10, 20, 50], index=0)
    order = f3.selectbox("排序依據", ["件數", "銷售額"])
    top = q_top(url, s, e, top_n, None if brand_opt == "全館" else brand_opt,
                "revenue" if order == "銷售額" else "qty")
    if not top.empty:
        matcher = get_matcher(url)
        top_table = pd.DataFrame({
            "商品ID": top["sku"],
            "商品名稱": [matcher.name_of(sku) for sku in top["sku"]],
            "銷售件數": top["qty"],
            "銷售額": top["revenue"],
        })
        st.dataframe(
            top_table,
            hide_index=True,
            width="stretch",
            column_config={
                "銷售額": st.column_config.NumberColumn(format="$%d"),
            },
        )

    # --- 尺寸深度分析 ---
    st.subheader("📐 商品尺寸深度分析")
    t1, t2 = st.columns([2, 2])
    picked = t1.selectbox("從排行榜選擇型號", ["—"] + top["sku"].tolist())
    keyword = t2.text_input("或搜尋型號（部分字串）")
    sku = None
    if keyword.strip():
        found = db.search_skus(get_engine(url), s, e, keyword.strip())
        if found.empty:
            st.warning("查無符合的型號。")
        else:
            sku = st.selectbox("搜尋結果", found["sku"].tolist(),
                               format_func=lambda x: f"{x}（{found.set_index('sku').loc[x, 'brand']}）")
    elif picked != "—":
        sku = picked

    if sku:
        sz = q_size(url, s, e, sku)
        if sz.empty:
            st.info("此型號在區間內無銷售。")
        else:
            sz = sort_sizes(sz)
            st.dataframe(
                sz[["duan_code", "尺寸", "size", "qty"]].rename(
                    columns={"duan_code": "段碼", "size": "ERP 尺寸", "qty": "件數"}),
                hide_index=True,
                width="stretch",
            )


# ---------------------------------------------------------------------------
# 頁面：資料維護
# ---------------------------------------------------------------------------

def page_admin(url: str):
    st.header("🛠️ 資料維護")
    engine = get_engine(url)
    st.caption("管理商品主檔資料庫（新增品項、批次匯入）、銷售匯入批次與未分類品牌修正。")

    _render_catalog_section(url, engine)
    st.divider()

    st.subheader("🗂️ 銷售匯入批次（依銷售日期）")
    batches = db.import_batches(engine)
    if batches.empty:
        st.info("資料庫目前沒有銷售明細資料。")
    else:
        st.dataframe(batches.rename(columns={"day": "銷售日期", "item_rows": "明細筆數", "sku_count": "型號數",
                                             "qty": "件數", "revenue": "銷售額", "imported_at": "最後寫入"}),
                     hide_index=True, width="stretch")
        with st.form("del_form"):
            d = st.selectbox("刪除指定日期之全部資料", batches["day"].tolist())
            confirm = st.checkbox("我確認要刪除（無法復原）")
            if st.form_submit_button("🗑️ 刪除", type="secondary") and confirm:
                n = db.delete_date(engine, d)
                clear_cache()
                st.success(f"已刪除 {d} 共 {n} 筆。")
                st.rerun()

    st.divider()
    st.subheader(f"🏷️ 品牌修正（{UNCLASSIFIED}）")
    matcher = get_matcher(url)
    unc = db.distinct_skus_by_brand(engine, UNCLASSIFIED)
    if unc.empty:
        st.success("目前沒有未分類型號 🎉")
    else:
        if st.button("🔄 以最新商品主檔 / 規則重新判定"):
            fixed = 0
            for sku in unc["sku"]:
                b, method = matcher.match(sku)
                if method != "none":
                    db.update_brand(engine, sku, b)
                    fixed += 1
            clear_cache()
            st.success(f"重新判定完成，修正 {fixed} 個型號。")
            st.rerun()
        brands = sorted((set(matcher.sku_to_brand.values()) | set(db.list_brands(engine))) - {UNCLASSIFIED})
        edited = st.data_editor(
            unc.assign(new_brand=UNCLASSIFIED).rename(columns={"sku": "型號", "brand": "目前品牌", "qty": "累計件數"}),
            column_config={"new_brand": st.column_config.SelectboxColumn("修正為", options=[UNCLASSIFIED] + brands)},
            disabled=["型號", "目前品牌", "累計件數"], hide_index=True, width="stretch",
        )
        if st.button("💾 儲存品牌修正", type="primary"):
            changes = edited[edited["new_brand"] != UNCLASSIFIED]
            for _, r in changes.iterrows():
                db.update_brand(engine, r["型號"], r["new_brand"])
            clear_cache()
            st.success(f"已更新 {len(changes)} 個型號。")
            st.rerun()

    with st.expander("單一型號品牌修改"):
        with st.form("one_brand"):
            sku = st.text_input("型號")
            new_b = st.text_input("新品牌名稱")
            if st.form_submit_button("更新") and sku and new_b:
                n = db.update_brand(engine, sku.strip(), new_b.strip())
                clear_cache()
                st.success(f"已更新 {n} 筆明細。")


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

def main():
    if not check_password():
        return
    url, is_local = _database_url()
    st.sidebar.title("👟 京鴻 ERP 銷售分析")
    page = st.sidebar.radio("功能", ["📊 營運報表", "📥 匯入資料", "🛠️ 資料維護"])
    if is_local:
        st.sidebar.warning("未設定 DATABASE_URL，目前使用本機 SQLite（僅供測試）。")
    st.sidebar.divider()
    try:
        get_engine(url)
    except Exception as ex:  # noqa: BLE001
        st.error(f"資料庫連線失敗：{ex}")
        st.stop()

    if page == "📥 匯入資料":
        page_import(url)
    elif page == "🛠️ 資料維護":
        page_admin(url)
    else:
        page_report(url)


main()
