"""段碼 (size-run code) → 17 個尺寸欄位之實際尺寸對照表。

資料來源：專案資料夾 `段碼/斷碼*.jpg`（京鴻 ERP 畫面截圖）。
每個段碼固定對應 17 個尺寸欄位，空字串代表該欄位未使用。
ERP 以「8-」表示半號（US 8.5），此處保留 ERP 原始標籤。

如需新增段碼（例如目前樣本中出現但尚無截圖的段碼 6），
直接在 SIZE_CODES 補上一筆 17 個元素的清單即可。
"""

from __future__ import annotations

SIZE_SLOTS = 17


def _half_run(start: int, end: int) -> list[str]:
    """產生 ERP 半號序列，例如 4, 4-, 5, 5-, ..., 12。"""
    out: list[str] = []
    for n in range(start, end + 1):
        out.append(str(n))
        if n != end:
            out.append(f"{n}-")
    return out


def _decimal_run(start: float, end: float, step: float = 0.5, suffix: str = "") -> list[str]:
    out: list[str] = []
    v = start
    while v <= end + 1e-9:
        label = str(int(v)) if float(v).is_integer() else f"{v:g}"
        out.append(f"{label}{suffix}")
        v += step
    return out


def _pad(labels: list[str]) -> list[str]:
    if len(labels) > SIZE_SLOTS:
        raise ValueError(f"段碼尺寸超過 {SIZE_SLOTS} 欄：{labels}")
    return labels + [""] * (SIZE_SLOTS - len(labels))


SIZE_CODES: dict[str, list[str]] = {
    # 斷碼1：單一尺寸（配件 / Free size）
    "1": _pad(["F"]),
    # 斷碼2：US 4 ~ 12（含半號）
    "2": _pad(_half_run(4, 12)),
    # 斷碼3：服飾尺寸
    "3": _pad(["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL", "5XL", "XXS", "3XS"]),
    # 斷碼4：22 ~ 30（cm，含半號）
    "4": _pad(_half_run(22, 30)),
    # 斷碼7：歐碼 34 ~ 50
    "7": _pad([str(n) for n in range(34, 51)]),
    # 斷碼A：7 ~ 15（含半號）
    "A": _pad(_half_run(7, 15)),
    # 斷碼C：2 ~ 10（含半號）
    "C": _pad(_half_run(2, 10)),
    # 斷碼D：1 ~ 6（0.5 進位）+ 童鞋 11 ~ 13.5
    "D": _pad(_decimal_run(1, 6) + _decimal_run(11, 13.5)),
    # 斷碼H：3 ~ 11（0.5 進位）
    "H": _pad(_decimal_run(3, 11)),
    # 斷碼I：11.5 ~ 19.5（0.5 進位）
    "I": _pad(_decimal_run(11.5, 19.5)),
    # 斷碼J：童鞋 1Y ~ 7Y（0.5 進位）
    "J": _pad(_decimal_run(1, 7, suffix="Y")),
    # 斷碼N：S/M、M/L、L/XL，第 8~13 欄為特殊代碼
    "N": _pad(["S/M", "M/L", "L/XL", "", "", "", "", "714", "738", "712", "758", "734", "778"]),
    # 斷碼O：依 ERP 畫面原樣保留
    "O": _pad(["11", "11.5", "1175", "12", "12.5", "1275", "13", "1125", "1225",
               "20", "20-", "21", "21-"]),
}


def normalize_label(raw: str) -> str:
    """把 Excel 轉成負數的半號（如 '-4'、'-4.0'）還原為 ERP 標籤 '4-'。"""
    s = str(raw).strip()
    if s.endswith(".0"):
        s = s[:-2]
    if s.startswith("-") and s[1:].replace(".", "", 1).isdigit():
        s = f"{s[1:]}-"
    return s


def resolve_size(code: str, slot: int, file_codes: dict[str, list[str]] | None = None) -> str:
    """依段碼與欄位序號（0-based）取得尺寸標籤。

    優先使用內建對照表；若段碼不在表內，改用從 ERP 檔頭動態擷取的對照；
    兩者都沒有時回傳 `段碼{code}-第{n}欄`，方便事後於資料庫修正。
    """
    code = str(code).strip().upper()
    for table in (SIZE_CODES, file_codes or {}):
        labels = table.get(code)
        if labels and slot < len(labels) and labels[slot]:
            return labels[slot]
    return f"段碼{code}-第{slot + 1}欄"


def is_known_code(code: str, file_codes: dict[str, list[str]] | None = None) -> bool:
    code = str(code).strip().upper()
    return code in SIZE_CODES or code in (file_codes or {})


def size_sort_key(code: str, size: str) -> tuple[int, int, str]:
    """報表排序用：依段碼對照表中的欄位順序排列尺寸。"""
    labels = SIZE_CODES.get(str(code).strip().upper(), [])
    if size in labels:
        return (0, labels.index(size), size)
    return (1, 0, str(size))


def display_size(size: str) -> str:
    """ERP 半號標籤轉為易讀格式：'8-' → '8.5'。其他尺寸原樣回傳。"""
    s = str(size)
    if s.endswith("-") and s[:-1].isdigit():
        return f"{s[:-1]}.5"
    return s
