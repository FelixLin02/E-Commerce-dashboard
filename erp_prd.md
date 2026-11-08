# 產品需求文件 (PRD)：京鴻 ERP 運動鞋服銷售數據分析與樞紐報表系統

---

## 1. 專案背景與痛點 (Problem Statement)
* **營運模式**：每日需處理平均約 70 筆運動鞋與服飾之銷售訂單，資料來源為「京鴻 ERP 系統」匯出之「商品銷售總表列印」Excel 檔。
* **痛點 1（分頁格式雜訊）**：ERP 匯出之 Excel 採列印分頁排版，每頁頂部皆夾帶重複之標頭（抬頭、日期、頁次、段碼定義、欄位名），無法直接用傳統資料庫或樞紐分析表導入。
* **痛點 2（二維矩陣排版）**：尺寸資料橫向分散於 17 個欄位，並以「段碼（如 1、2、3）」決定各欄位代表之實際尺寸（如 F、半號 4~12、XS~3XL），需先解碼與寬轉長（Unpivot）後方可進行單品與尺寸層級之統計。
* **痛點 3（硬體與資安限制）**：公司電腦受限無法部署後台服務，個人筆電 C 槽空間有限，需採取**本地零負擔（Zero-footprint）**、全雲端託管之輕量架構。

---

## 2. 產品目標 (Product Goals)
1. **自動化 ETL 解析**：使用者僅需上傳京鴻 ERP 原檔，系統自動清洗跨頁重複表頭、比對型號品牌、解析段碼與 17 欄尺寸，並轉置為標準關聯式明細。
2. **多維度動態樞紐**：提供前端時間軸自訂（日/月/自訂區間），即時查詢品牌總銷量、銷售額、各品牌排行榜及特定單品之尺寸銷售分佈。
3. **雲端輕量與長期儲存**：採用雲端 PostgreSQL（Supabase）+ 前端介面（Streamlit Cloud），在免費額度內穩定支撐 8~10 年以上資料量。

---

## 3. 使用者角色與使用情境 (User Persona & User Journey)
* **使用者**：賣場營運管理員。
* **使用旅程**：
  1. **日常匯入**：每日營業結束後，自京鴻 ERP 匯出銷售 Excel，開啟雲端報表網址，拖曳上傳檔案完成同步。
  2. **營運調閱**：隨時開啟網頁，挑選「本月」、「前 30 天」或特定自訂日期。
  3. **策略分析**：
     * 檢視全店總營收與總件數。
     * 檢視各品牌佔比（Nike、Jordan、adidas 等）與排名。
     * 查看 TOP 10 爆款單品，並點選深入查看該商品之各尺碼斷碼/熱銷分佈。

---

## 4. 系統架構與技術選型 (System Architecture)

```text
┌─────────────────────────┐
│ 京鴻 ERP 匯出 Excel      │
└────────────┬────────────┘
             │ (瀏覽器上傳)
             ▼
┌────────────────────────────────────────────────────────┐
│ Streamlit Community Cloud (前端介面 + ETL 運算引擎)     │
│  - 跨頁重複表頭過濾器                                    │
│  - 品牌正規表示式比對 (Regex Matcher)                    │
│  - 段碼 17 欄動態映射器 (Wide-to-Long Unpivot)          │
│  - 記憶體快取 (@st.cache_data)                          │
└────────────┬───────────────────────────────────────────┘
             │ (SQLAlchemy + Pooler SSL 連線)
             ▼
┌────────────────────────────────────────────────────────┐
│ Supabase (PostgreSQL 雲端託管資料庫)                     │
│  - 儲存商品主檔、銷售明細與索引                           │
│  - Table Editor 線上增修刪改 (CRUD)                     │
│  - 執行 GROUP BY / SUM 聚合運算                          │
└────────────────────────────────────────────────────────┘
```

* **前端與運算層**：Streamlit Community Cloud（完全免費、免自建伺服器）。
* **資料儲存層**：Supabase Free Tier（PostgreSQL，提供 500 MB 儲存空間與 Web 管理後台）。
* **部署方式**：GitHub 私有倉庫連動 Streamlit Cloud，金鑰以 Streamlit Secrets 管理，本機零殘留。

---

## 5. 資料規格與資料庫架構 (Database Schema)

### 5.1 邏輯明細表 (`daily_sales_items`)
所有 ERP 匯入之列經清洗與尺寸橫轉直後，寫入此明細表：

| 欄位名稱 | 型態 | 說明 | 範例 |
| :--- | :--- | :--- | :--- |
| `id` | BIGSERIAL | 主鍵（流水號） | `1` |
| `sale_date` | TIMESTAMP | 銷售/製表日期（加索引） | `2026-11-03 00:00:00` |
| `sku` | VARCHAR(64) | 商品型號（加索引） | `79542008` |
| `brand` | VARCHAR(64) | 系統判讀之品牌（加索引） | `自訂品牌A` / `Nike` |
| `duan_code` | VARCHAR(32) | 原始段碼代號 | `2` |
| `size` | VARCHAR(32) | 映射出之實際尺寸（加索引） | `8-` (US 8.5) |
| `quantity` | INT | 該尺寸銷售數量 | `1` |
| `unit_price` | NUMERIC(10,2) | 單價（合計金額 / 總件數） | `3947.00` |
| `row_total_amount` | NUMERIC(10,2) | 該型號銷售合計金額 | `11841.00` |
| `created_at` | TIMESTAMP | 寫入時間 | `2026-11-07 18:14:00` |

### 5.2 資料庫建表 DDL (SQL)

```sql
CREATE TABLE IF NOT EXISTS daily_sales_items (
    id BIGSERIAL PRIMARY KEY,
    sale_date TIMESTAMP NOT NULL,
    sku VARCHAR(64) NOT NULL,
    brand VARCHAR(64) NOT NULL,
    duan_code VARCHAR(32),
    size VARCHAR(32) NOT NULL,
    quantity INT NOT NULL DEFAULT 1,
    unit_price NUMERIC(10, 2) DEFAULT 0,
    row_total_amount NUMERIC(10, 2) DEFAULT 0,
    created_at TIMESTAMP DEFAULT NOW()
);

-- 索引優化：加速報表的時間區間與品牌聚合
CREATE INDEX IF NOT EXISTS idx_sales_date ON daily_sales_items(sale_date);
CREATE INDEX IF NOT EXISTS idx_sales_brand ON daily_sales_items(brand);
CREATE INDEX IF NOT EXISTS idx_sales_sku ON daily_sales_items(sku);
CREATE INDEX IF NOT EXISTS idx_sales_size ON daily_sales_items(size);
```

---

## 6. 功能需求詳細規格 (Functional Specifications)

### 6.1 模組一：京鴻 ERP 匯入與 ETL 解析模組
1. **跨頁重複表頭濾除**：
   * 識別包含「總公司」、「商品銷售總表列印」、「製表日期」、「商品型號」等標頭列，自動略過處理。
2. **段碼對照矩陣動態擷取**：
   * 逐列辨識開頭包含段碼代號（如 `1`、`2`、`3`）且後方伴隨尺寸標籤的資料列。
   * 自動截取其後 17 欄尺寸清單（如段碼 2 對應 `4, 4-, 5, 5-, ..., 11-, 12`）並暫存映射表。
3. **品牌判定**：
   * 根據型號字串透過正規表示式庫比對品牌（如 `DD/CW/DH` 判定為 Nike、`EG/GW` 判定為 adidas 等）；若皆未命中則標記為「其他/未分類」。
4. **二維橫轉直（Unpivot）**：
   * 讀取資料列之 17 欄數值，若 $數量 > 0$ 則提取對應尺寸生成一筆明細。
   * 若 17 欄皆無數值但「合計數量」$> 0$，自動視為單一尺碼（如段碼 1 之 `F`）。

### 6.2 模組二：自訂時間軸動態報表模組
1. **時間軸選擇器**：支援「單日」、「本週」、「本月」及「自訂起訖日期」篩選。
2. **營運核心 KPI 卡片**：即時計算選定區間之「總銷售額」、「總銷售件數」、「涵蓋訂單/品項數」。
3. **品牌樞紐分析**：
   * 各品牌之銷售件數與營收排行表。
   * 品牌銷售佔比圓餅圖 / 環形圖。
4. **熱門商品銷售排行 (TOP 10 / TOP 20)**：
   * 支援全館排行或依「指定品牌」下拉篩選該品牌的熱銷商品長條圖與清單。
5. **商品尺寸深度分析**：
   * 點選或搜尋特定型號，展開其在該期間內「各尺寸（如 US 8.5、US 9）銷售數量與比例」，協助補貨與斷碼決策。

### 6.3 模組三：資料增修刪改 (CRUD) 管理
* **報表端查詢**：所有查詢皆在資料庫端執行 `GROUP BY` 與 `SUM`，僅傳輸輕量聚合結果。
* **資料維護**：管理員可直接登入 Supabase 網頁後台 Table Editor：
  * 手動修正誤判之品牌或未分類型號。
  * 刪除重複匯入之日期批次或錯誤資料列。

---

## 7. 非功能性需求 (Non-Functional Requirements)
1. **儲存容量與可擴展性**：
   * 日均 70 筆訂單（約 105 筆尺寸明細），每年累積數據量小於 25 MB。
   * Supabase 500 MB 免費額度可滿足 8~10 年以上使用，無需提前升級付費。
2. **效能與延遲**：
   * 單一 Excel 匯入解析時間需在 3 秒內完成。
   * 樞紐報表查詢時間需控制在 1 秒以內（利用日期與品牌索引優化）。
3. **資料安全性**：
   * 資料庫連線一律使用 SSL 加密連線（`sslmode=require`）。
   * 資料庫密鑰與網址嚴格存放於 Streamlit Secrets，不簽入 GitHub 公開代碼。

---

## 8. 里程碑與交付時程 (Milestones)

| 階段 | 項目內容 | 預估工時 |
| :--- | :--- | :--- |
| **M1: 雲端環境初始化** | 建立 Supabase 專案、執行 Schema DDL、建立 GitHub 倉庫 | 0.5 天 |
| **M2: ETL 解析器開發** | 開發京鴻 ERP 跨頁重複表頭清洗、段碼 17 欄轉置與品牌正則比對函式 | 1.5 天 |
| **M3: 前端視覺化與報表** | 開發 Streamlit 儀表板、日期篩選器、品牌樞紐與 TOP 10 圖表 | 1.5 天 |
| **M4: 部署與驗收測試** | 上線 Streamlit Cloud、上傳真實歷史 Excel 進行驗收測試與除錯 | 0.5 天 |