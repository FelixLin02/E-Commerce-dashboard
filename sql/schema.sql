-- 京鴻 ERP 銷售明細與商品主檔（於 Supabase SQL Editor 執行）
-- 註：app 啟動時也會自動建立（CREATE IF NOT EXISTS），此檔供手動初始化使用。

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

CREATE INDEX IF NOT EXISTS idx_sales_date ON daily_sales_items(sale_date);
CREATE INDEX IF NOT EXISTS idx_sales_brand ON daily_sales_items(brand);
CREATE INDEX IF NOT EXISTS idx_sales_sku ON daily_sales_items(sku);
CREATE INDEX IF NOT EXISTS idx_sales_size ON daily_sales_items(size);

-- 商品主檔（基於 SHOPPEE STOCK.xlsx，提供由 ID 查詢商品名稱與品牌，支援單筆新增與批次匯入）
CREATE TABLE IF NOT EXISTS product_catalog (
    sku VARCHAR(64) PRIMARY KEY,
    product_name TEXT NOT NULL,
    brand VARCHAR(64) NOT NULL,
    updated_at TIMESTAMP DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_catalog_brand ON product_catalog(brand);

-- 建議：啟用 RLS 並不建立任何 policy，阻擋透過 Supabase 公開 API (anon key) 存取；
-- app 以資料庫使用者直連（postgres 角色）不受影響。
ALTER TABLE daily_sales_items ENABLE ROW LEVEL SECURITY;
ALTER TABLE product_catalog ENABLE ROW LEVEL SECURITY;
