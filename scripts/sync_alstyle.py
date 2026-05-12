"""
Скачивает прайс Al-Style и конвертирует в JSON для Nemo B2B
URL: https://b2b.al-style.kz/export/Al-Style_price.xlsx
Запускается: каждые 30 минут через GitHub Actions
Результат: cloud-data/alstyle.json
"""
 
import os, json, sys, io
from datetime import datetime, timezone, timedelta
import requests
import pandas as pd
 
PRICE_URL  = "https://b2b.al-style.kz/export/Al-Style_price.xlsx"
TOKEN      = os.environ.get("ALSTYLE_TOKEN", "78uoSPj9NHdGKTDlaXhYaUd8htymrs8q")
OUT_FILE   = "cloud-data/alstyle.json"
ALMATY     = timezone(timedelta(hours=5))
 
# ── Скачать файл ──────────────────────────────────────────────
def download():
    print(f"⬇️  Скачиваем: {PRICE_URL}")
 
    # Пробуем разные способы передачи токена
    attempts = [
        {"headers": {"Authorization": f"Bearer {TOKEN}", "Accept": "*/*"}},
        {"headers": {"X-Api-Key": TOKEN, "Accept": "*/*"}},
        {"params":  {"token": TOKEN}},
        {"params":  {"api_key": TOKEN}},
        {},  # без авторизации (вдруг публичная ссылка)
    ]
 
    for opts in attempts:
        try:
            r = requests.get(PRICE_URL, timeout=60, **opts)
            print(f"   → {r.status_code} ({len(r.content)} байт) | opts: {list(opts.keys())}")
            if r.status_code == 200 and len(r.content) > 1000:
                print("✅ Файл получен")
                return r.content
        except Exception as e:
            print(f"   ⚠️  {e}")
 
    sys.exit("❌ Не удалось скачать файл ни одним способом")
 
# ── Парсинг Excel ─────────────────────────────────────────────
def parse(xlsx_bytes):
    print("📊 Парсим Excel...")
    xl = pd.ExcelFile(io.BytesIO(xlsx_bytes))
    print(f"   Листов: {xl.sheet_names}")
 
    all_items = []
    for sheet in xl.sheet_names:
        df = pd.read_excel(io.BytesIO(xlsx_bytes), sheet_name=sheet, header=None, engine="openpyxl")
        items = parse_sheet(df, sheet)
        print(f"   Лист «{sheet}»: {len(items)} товаров")
        all_items.extend(items)
 
    print(f"✅ Итого: {len(all_items)} товаров")
    return all_items
 
def parse_sheet(df, sheet_name):
    """Авто-определяет строку с заголовком и парсит данные"""
    # Ищем строку-заголовок в первых 20 строках
    header_row = None
    for i in range(min(20, len(df))):
        row = [str(c).lower().strip() for c in df.iloc[i]]
        has_name  = any(k in ' '.join(row) for k in ['наименован', 'название', 'name', 'товар'])
        has_price = any(k in ' '.join(row) for k in ['цена', 'price', 'стоимость', 'прайс'])
        if has_name and has_price:
            header_row = i
            break
 
    if header_row is None:
        # Попробуем иерархический формат
        return parse_hierarchical(df, sheet_name)
 
    # Переименовываем с заголовком
    df.columns = [str(c).strip() for c in df.iloc[header_row]]
    df = df.iloc[header_row + 1:].reset_index(drop=True)
    df = df.dropna(how='all')
 
    col = find_columns(df.columns.tolist())
    items = []
    current_cat = sheet_name
 
    for _, row in df.iterrows():
        name = str(row.get(col['name'], '') or '').strip()
        if not name or name.lower() in ('nan', 'none', ''):
            continue
 
        price_raw = row.get(col['price'], 0) if col['price'] else 0
        price = safe_float(price_raw)
 
        currency = 'KZT'
        if col['currency']:
            c = str(row.get(col['currency'], '') or '').strip().upper()
            if c in ('USD', 'EUR', 'RUB'):
                currency = c
 
        qty = safe_float(row.get(col['qty'], 0)) if col['qty'] else 0
        art = str(row.get(col['article'], '') or '').strip() if col['article'] else ''
        cat = str(row.get(col['category'], '') or current_cat).strip() if col['category'] else current_cat
        vendor = str(row.get(col['vendor'], '') or 'Al-Style').strip() if col['vendor'] else 'Al-Style'
 
        if not name or price <= 0:
            continue
 
        items.append({
            "article":  art,
            "name":     name[:200],
            "category": cat[:100],
            "vendor":   vendor[:80],
            "price":    price,
            "currency": currency,
            "qty":      int(qty),
        })
 
    return items
 
def parse_hierarchical(df, sheet_name):
    """Для прайсов где категория — отдельная строка (как PriceMC)"""
    items = []
    current_cat = sheet_name
    for _, row in df.iterrows():
        vals = [str(v).strip() for v in row if str(v).strip() not in ('', 'nan', 'None')]
        if not vals:
            continue
        if len(vals) == 1:
            current_cat = vals[0]
            continue
        if len(vals) >= 2:
            name  = vals[0] if len(vals[0]) > 5 else vals[1]
            price = next((safe_float(v) for v in vals[1:] if safe_float(v) > 0), 0)
            if name and price > 0:
                items.append({"article": "", "name": name[:200], "category": current_cat,
                              "vendor": "Al-Style", "price": price, "currency": "KZT", "qty": 0})
    return items
 
def find_columns(headers):
    """Находит нужные колонки по ключевым словам"""
    def find(keywords):
        for h in headers:
            hl = str(h).lower()
            if any(k in hl for k in keywords):
                return h
        return None
 
    return {
        'name':     find(['наименован', 'название', 'товар', 'name', 'product']),
        'article':  find(['артикул', 'article', 'sku', 'код', 'part']),
        'price':    find(['цена', 'price', 'стоимость']),
        'currency': find(['валют', 'currency']),
        'qty':      find(['кол-во', 'количест', 'остат', 'qty', 'stock', 'свободно']),
        'category': find(['категор', 'группа', 'раздел', 'category', 'group']),
        'vendor':   find(['бренд', 'вендор', 'произво', 'vendor', 'brand']),
    }
 
def safe_float(v):
    try:
        return float(str(v).replace(' ', '').replace(',', '.').replace('₸', '').replace('$', ''))
    except:
        return 0.0
 
# ── Сохранить JSON ────────────────────────────────────────────
def save(items):
    os.makedirs("cloud-data", exist_ok=True)
    now_almaty = datetime.now(ALMATY).strftime("%d.%m.%Y %H:%M")
    output = {
        "source":     "Al-Style B2B",
        "url":        PRICE_URL,
        "updated_at": now_almaty,
        "count":      len(items),
        "items":      items,
    }
    with open(OUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    size_kb = os.path.getsize(OUT_FILE) // 1024
    print(f"💾 Сохранено: {OUT_FILE} ({size_kb} KB, {len(items)} товаров)")
 
# ── Точка входа ───────────────────────────────────────────────
if __name__ == "__main__":
    xlsx = download()
    items = parse(xlsx)
    if not items:
        sys.exit("❌ Товары не найдены — проверьте формат файла")
    save(items)
    print("🎉 Готово!")
