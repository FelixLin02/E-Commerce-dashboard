# 京鴻 ERP 運動鞋服銷售數據分析與樞紐報表系統

依 [erp_prd.md](erp_prd.md) 實作：Streamlit（前端 + ETL）＋ Supabase PostgreSQL。

## 專案結構

```
app.py                  Streamlit 主程式（營運報表 / 匯入資料 / 商品主檔 / 資料維護）
etl/
  parser.py             京鴻 ERP「商品銷售總表列印」解析（跨頁表頭濾除、17 欄尺寸轉置）
  brand.py              商品主檔與品牌判定（product_catalog 資料庫精確 → 前綴 → 正規表示式 → 其他/未分類）
  size_codes.py         段碼 → 17 欄尺寸對照表（來源：段碼/斷碼*.jpg）
  db.py                 資料庫存取與聚合查詢（daily_sales_items 與 product_catalog）
品牌/*.xlsx             各品牌主商品貨號（雲端初始化備援主檔）
sql/schema.sql          建表 DDL（含 daily_sales_items 與 product_catalog，app 啟動也會自動建立）
tests/                  pytest 單元測試（43 項測試）
```

## 部署（M1 / M4）

1. **Supabase**：建立專案 → SQL Editor 執行 `sql/schema.sql`。
   Project Settings → Database → Connection string → 複製 **Session pooler** URI。
2. **GitHub**：建立 *私有* 倉庫並推送本資料夾（`.gitignore` 已排除 secrets、Sample、補充）。
3. **Streamlit Community Cloud**：New app → 選倉庫、主檔 `app.py` →
   Advanced settings → Secrets 貼上（參考 `.streamlit/secrets.toml.example`）：
   ```toml
   DATABASE_URL = "postgresql://postgres.<ref>:<PASSWORD>@aws-0-<region>.pooler.supabase.com:5432/postgres?sslmode=require"
   APP_PASSWORD = "自訂密碼"   # 選填
   ```

## 本機測試（選用）

```powershell
pip install -r requirements.txt pytest
python -m pytest tests -q          # 單元測試（含樣本檔對帳 263 件 / $499,216 與商品主檔 CRUD）
streamlit run app.py               # 未設 DATABASE_URL 時使用 local_dev.sqlite
```

## 維護說明

* **商品主檔管理（📦 商品主檔）**：
  * 首次啟動自動由 `補充/SHOPPEE STOCK.xlsx`（或 `品牌/*.xlsx`）初始化匯入至 `product_catalog` 資料表。
  * 支援於網頁介面 **逐一新增品項**、**批次匯入新增（自動排除重複 ID）**、**整批匯入覆蓋** 與 **關鍵字搜尋／刪除**。
* **新增段碼**：於 `etl/size_codes.py` 的 `SIZE_CODES` 加入 17 個元素的清單。
  未建檔段碼的尺寸會暫記為 `段碼X-第N欄`，並於匯入時提示。
* **尺寸標籤**：資料庫保留 ERP 原始標籤（`8-` = US 8.5），報表顯示時轉為 `8.5`。
* **銷售日期**：匯入時可選「日期區間起日」（預設）、「製表日期」或手動指定；
  同日期重複匯入預設覆蓋，避免重複計算。
* **銷售額**：以 `SUM(quantity × unit_price)` 計算，單價 = 合計金額 / 數量（四捨五入至小數 2 位）。
