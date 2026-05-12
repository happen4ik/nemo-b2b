"""
Скачивает прайс Al-Style и конвертирует в JSON для Nemo B2B
URL: https://b2b.al-style.kz/export/Al-Style_price.xlsx
"""
 
import os, json, sys, io
from datetime import datetime, timezone, timedelta
import requests
import pandas as pd
 
PRICE_URL   = "https://b2b.al-style.kz/export/Al-Style_price.xlsx"
TOKEN       = os.environ.get("ALSTYLE_TOKEN", "")
OUT_DIR     = "cloud-data"
OUT_FILE    = f"{OUT_DIR}/alstyle.json"
ALMATY      = timezone(timedelta(hours=5))
 
def download():
    print(f"⬇️  Скачиваем: {PRICE_URL}")
    print(f"   Токен: {'есть (' + TOKEN[:8] + '...)' if TOKEN else 'НЕТ — добавьте ALSTYLE_TOKEN в Secrets'}")
 
    # Все варианты авторизации
    attempts = [
        ("Bearer header",    {"headers": {"Authorization": f"Bearer {TOKEN}"}}),
        ("Token header",     {"headers": {"Authorization": f"Token {TOKEN}"}}),
        ("X-Api-Key header", {"headers": {"X-Api-Key": TOKEN}}),
        ("?token= param",    {"params":  {"token": TOKEN}}),
        ("?api_key= param",  {"params":  {"api_key": TOKEN}}),
        ("Без авторизации",  {}),
    ]
 
    for name, opts in attempts:
        try:
            r = requests.get(
                PRICE_URL,
                timeout=60,
                allow_redirects=True,
                headers={"User-Agent": "Mozilla/5.0", "Accept": "*/*"},
                **opts
            )
            print(f"   [{name}] → HTTP {r.status_code}, размер: {len(r.content)} байт")
 
            if r.status_code == 200 and len(r.content) > 500:
                # Проверяем что это Excel
                if r.content[:4] in (b'PK\x03\x04', b'\xd0\xcf\x11\xe0'):
                    print(f"✅ Файл скачан ({name})")
                    return r.content
                else:
                    print(f"   ⚠️  Ответ не Excel: {r.content[:100]}")
            elif r.status_code in (401, 403):
                print(f"   ⛔ Доступ запрещён — IP не в whitelist или неверный токен")
            elif r.status_code == 404:
                print(f"   ❌ Файл не найден по этому URL")
 
        except Exception as e:
            print(f"   ⚠️  Ошибка: {e}")
 
    return None
 
def parse(xlsx_bytes):
    print("📊 Парсим Excel...")
    xl = pd.ExcelFile(io.BytesIO(xlsx_bytes))
    print(f"   Листов: {xl.sheet_names}")
 
    all_items = []
    for sheet in xl.sheet_names:
        df_raw = pd.read_excel(io.BytesIO(xlsx_bytes), sheet_name=sheet, header=None, engine="openpyxl")
        items = parse_sheet(df_raw, sheet)
        print(f"   Лист «{sheet}»: {len(items)} товаров")
        all_items.extend(items)
 
    return all_items
 
def parse_sheet(df, sheet_name):
    header_row = None
    for i in range(min(20, len(df))):
        row = [str(c).lower().strip() for c in df.iloc[i]]
        joined = ' '.join(row)
        has_name  = any(k in joined for k in ['наименован', 'название', 'name', 'товар', 'продукт'])
        has_price = any(k in joined for k in ['цена', 'price', 'стоимость'])
        if has_name and has_price:
            header_row = i
            break
 
    if header_row is None:
        return parse_hierarchical(df, sheet_name)
 
    df2 = df.copy()
    df2.columns = [str(c).strip() for c in df.iloc[header_row]]
    df2 = df2.iloc[header_row + 1:].reset_index(drop=True).dropna(how='all')
 
    col = find_columns(df2.columns.tolist())
    items = []
    current_cat = sheet_name
 
    for _, row in df2.iterrows():
        name = str(row.get(col['name'], '') or '').strip()
        if not name or name.lower() in ('nan', 'none', ''):
            continue
 
        price = safe_float(row.get(col['price'], 0) if col['price'] else 0)
        if price <= 0:
            continue
 
        currency = 'KZT'
        if col['currency']:
            c = str(row.get(col['currency'], '') or '').strip().upper()
            if c in ('USD', 'EUR', 'RUB'):
                currency = c
 
        items.append({
            "article":  str(row.get(col['article'], '') or '').strip()[:100] if col['article'] else '',
            "name":     name[:200],
            "category": str(row.get(col['category'], '') or current_cat).strip()[:100] if col['category'] else current_cat,
            "vendor":   str(row.get(col['vendor'], '') or 'Al-Style').strip()[:80] if col['vendor'] else 'Al-Style',
            "price":    round(price, 2),
            "currency": currency,
            "qty":      int(safe_float(row.get(col['qty'], 0))) if col['qty'] else 0,
        })
 
    return items
 
def parse_hierarchical(df, sheet_name):
    items, current_cat = [], sheet_name
    for _, row in df.iterrows():
        vals = [str(v).strip() for v in row if str(v).strip() not in ('', 'nan', 'None')]
        if not vals:
            continue
        if len(vals) == 1 and len(vals[0]) < 80:
            current_cat = vals[0]
            continue
        if len(vals) >= 2:
            name  = vals[0] if len(vals[0]) > 3 else (vals[1] if len(vals) > 1 else '')
            price = next((safe_float(v) for v in vals[1:] if safe_float(v) > 0), 0)
            if name and price > 0:
                items.append({"article": "", "name": name[:200], "category": current_cat,
                              "vendor": "Al-Style", "price": round(price, 2), "currency": "KZT", "qty": 0})
    return items
 
def find_columns(headers):
    def find(keywords):
        for h in headers:
            if any(k in str(h).lower() for k in keywords):
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
        return float(str(v).replace(' ', '').replace(',', '.').replace('₸','').replace('$',''))
    except:
        return 0.0
 
def save(items, success):
    os.makedirs(OUT_DIR, exist_ok=True)
    now = datetime.now(ALMATY).strftime("%d.%m.%Y %H:%M")
 
    # Если скачать не удалось — сохраняем статус ошибки, старые данные не трогаем
    if not success:
        status_file = f"{OUT_DIR}/alstyle_status.json"
        with open(status_file, "w") as f:
            json.dump({"error": True, "checked_at": now, "message": "Не удалось скачать файл — IP не в whitelist"}, f)
        print(f"⚠️  Статус ошибки сохранён в {status_file}")
        return
 
    output = {
        "source":     "Al-Style B2B",
        "url":        PRICE_URL,
        "updated_at": now,
        "count":      len(items),
        "items":      items,
    }
    with open(OUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, separators=(',', ':'))
    print(f"💾 Сохранено: {OUT_FILE} ({os.path.getsize(OUT_FILE)//1024} KB, {len(items)} товаров)")
 
if __name__ == "__main__":
    xlsx = download()
    if not xlsx:
        print("\n" + "="*60)
        print("❌ ФАЙЛ НЕ СКАЧАН")
        print("Причина: GitHub Actions IP не в белом списке Al-Style")
        print("Решение: напишите в Al-Style чтобы добавили в whitelist IP GitHub Actions")
        print("ИЛИ: попросите Al-Style выдать публичную ссылку без whitelist")
        print("="*60)
        save([], success=False)
        sys.exit(0)   # exit 0 — не считать ошибкой workflow
 
    items = parse(xlsx)
    if not items:
        print("⚠️  Товары не найдены — возможно изменился формат файла")
        save([], success=False)
        sys.exit(0)
 
    save(items, success=True)
    print(f"🎉 Готово! {len(items)} товаров обновлено")
