"""
Скачивает YML-фид Al-Style и конвертирует в JSON для Nemo B2B
Фид: https://apifeed.al-style.kz/feed.xml  (формат YML/Яндекс.Маркет)
Резерв: https://b2b.al-style.kz/export/Al-Style_price.xlsx
"""
 
import os, json, sys, io
from datetime import datetime, timezone, timedelta
import requests
 
try:
    from lxml import etree as ET
    HAS_LXML = True
except ImportError:
    import xml.etree.ElementTree as ET
    HAS_LXML = False
 
TOKEN    = os.environ.get("ALSTYLE_TOKEN", "78uoSPj9NHdGKTDlaXhYaUd8htymrs8q")
YML_URL  = "https://apifeed.al-style.kz/feed.xml"
XLS_URL  = "https://b2b.al-style.kz/export/Al-Style_price.xlsx"
OUT_DIR  = "cloud-data"
OUT_FILE = f"{OUT_DIR}/alstyle.json"
ALMATY   = timezone(timedelta(hours=5))
 
# ── Скачать файл ──────────────────────────────────────────────
def download(url):
    print(f"⬇️  {url}")
    attempts = [
        {"params":  {"token": TOKEN}},
        {"params":  {"api_key": TOKEN}},
        {"headers": {"Authorization": f"Bearer {TOKEN}"}},
        {"headers": {"X-Api-Key": TOKEN}},
        {},
    ]
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0", "Accept": "*/*"})
 
    for opts in attempts:
        try:
            r = session.get(url, timeout=60, allow_redirects=True, **opts)
            print(f"   HTTP {r.status_code} | {len(r.content)} байт | {list(opts.keys()) or 'no-auth'}")
            if r.status_code == 200 and len(r.content) > 200:
                return r.content
            if r.status_code == 403:
                print(f"   ⛔ {r.text[:60]}")
        except Exception as e:
            print(f"   ⚠️  {e}")
    return None
 
# ── Парсинг YML (Яндекс.Маркет XML) ──────────────────────────
def parse_yml(content):
    print("📊 Парсим YML-фид...")
    try:
        root = ET.fromstring(content)
    except Exception as e:
        print(f"❌ Ошибка XML: {e}")
        return []
 
    # Пространства имён — иногда есть, иногда нет
    ns = ''
    if root.tag.startswith('{'):
        ns = root.tag.split('}')[0] + '}'
 
    def tag(name): return f"{ns}{name}"
 
    shop = root.find(tag('shop'))
    if shop is None:
        shop = root  # некоторые фиды без <shop>
 
    # Категории → dict {id: name}
    categories = {}
    cats_el = shop.find(tag('categories'))
    if cats_el is not None:
        for cat in cats_el.findall(tag('category')):
            categories[cat.get('id', '')] = (cat.text or '').strip()
 
    # Валюты → dict {id: rate}
    currencies = {'KZT': 1.0}
    curr_el = shop.find(tag('currencies'))
    if curr_el is not None:
        for cur in curr_el.findall(tag('currency')):
            try:
                currencies[cur.get('id')] = float(cur.get('rate', 1) or 1)
            except: pass
 
    # Офферы (товары)
    offers_el = shop.find(tag('offers'))
    if offers_el is None:
        print("⚠️  Тег <offers> не найден")
        return []
 
    items = []
    for offer in offers_el.findall(tag('offer')):
        def g(t):
            el = offer.find(tag(t))
            return (el.text or '').strip() if el is not None and el.text else ''
 
        name     = g('name') or g('model') or g('typePrefix')
        price_s  = g('price')
        currency = g('currencyId') or 'KZT'
        cat_id   = g('categoryId')
        vendor   = g('vendor') or g('manufacturer') or 'Al-Style'
        article  = g('vendorCode') or g('article') or offer.get('id', '')
        qty_s    = g('count') or g('quantity') or g('stock')
        available = offer.get('available', 'true').lower() != 'false'
 
        if not name:
            continue
        try:
            price = float(price_s.replace(' ', '').replace(',', '.')) if price_s else 0
        except: price = 0
        if price <= 0: continue
 
        # Конвертируем в KZT если нужно
        rate = currencies.get(currency, 1.0)
        price_kzt = round(price * rate, 2) if currency != 'KZT' else price
 
        try: qty = int(float(qty_s)) if qty_s else 0
        except: qty = 0
 
        items.append({
            "article":  article[:100],
            "name":     name[:200],
            "category": categories.get(cat_id, cat_id)[:100],
            "vendor":   vendor[:80],
            "price":    round(price, 2),
            "currency": currency,
            "price_kzt": price_kzt,
            "qty":      qty,
            "available": available,
        })
 
    print(f"✅ YML: {len(items)} товаров")
    return items
 
# ── Парсинг Excel (резерв) ────────────────────────────────────
def parse_xlsx(content):
    try:
        import pandas as pd
    except ImportError:
        print("⚠️  pandas не установлен")
        return []
 
    print("📊 Парсим Excel (резерв)...")
    xl = pd.ExcelFile(io.BytesIO(content))
    items = []
    for sheet in xl.sheet_names:
        df = pd.read_excel(io.BytesIO(content), sheet_name=sheet, header=None, engine="openpyxl")
        items.extend(parse_sheet(df, sheet))
    print(f"✅ Excel: {len(items)} товаров")
    return items
 
def parse_sheet(df, sheet_name):
    import pandas as pd
    header_row = None
    for i in range(min(20, len(df))):
        row = [str(c).lower().strip() for c in df.iloc[i]]
        joined = ' '.join(row)
        if any(k in joined for k in ['наименован','name']) and any(k in joined for k in ['цена','price']):
            header_row = i; break
    if header_row is None:
        return []
    df.columns = [str(c).strip() for c in df.iloc[header_row]]
    df = df.iloc[header_row + 1:].reset_index(drop=True).dropna(how='all')
    col = find_cols(df.columns.tolist())
    items = []
    for _, row in df.iterrows():
        name = str(row.get(col['name'], '') or '').strip()
        price = safe_float(row.get(col['price'], 0) if col['price'] else 0)
        if not name or price <= 0: continue
        items.append({
            "article":  str(row.get(col['article'], '') or '').strip()[:100] if col['article'] else '',
            "name":     name[:200],
            "category": str(row.get(col['category'], '') or sheet_name).strip()[:100] if col['category'] else sheet_name,
            "vendor":   str(row.get(col['vendor'], '') or 'Al-Style').strip()[:80] if col['vendor'] else 'Al-Style',
            "price":    round(price, 2), "currency": "KZT", "price_kzt": round(price, 2),
            "qty":      int(safe_float(row.get(col['qty'], 0))) if col['qty'] else 0,
            "available": True,
        })
    return items
 
def find_cols(headers):
    def f(kw): return next((h for h in headers if any(k in str(h).lower() for k in kw)), None)
    return {
        'name':    f(['наименован','название','товар','name']),
        'article': f(['артикул','article','sku','код']),
        'price':   f(['цена','price','стоимость']),
        'currency':f(['валют','currency']),
        'qty':     f(['кол-во','количест','остат','qty','stock','свободно']),
        'category':f(['категор','группа','category']),
        'vendor':  f(['бренд','вендор','произво','vendor']),
    }
 
def safe_float(v):
    try: return float(str(v).replace(' ','').replace(',','.').replace('₸',''))
    except: return 0.0
 
# ── Сохранить ─────────────────────────────────────────────────
def save(items, source_url):
    os.makedirs(OUT_DIR, exist_ok=True)
    now = datetime.now(ALMATY).strftime("%d.%m.%Y %H:%M")
    out = {
        "source":     "Al-Style",
        "feed_url":   source_url,
        "updated_at": now,
        "count":      len(items),
        "items":      items,
    }
    with open(OUT_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, separators=(',', ':'))
    kb = os.path.getsize(OUT_FILE) // 1024
    print(f"💾 {OUT_FILE}: {kb} KB, {len(items)} товаров")
 
# ── Главная логика ────────────────────────────────────────────
if __name__ == "__main__":
    items, source = [], None
 
    # 1) Пробуем YML-фид
    data = download(YML_URL)
    if data and (data[:5] in (b'<?xml', b'<yml_') or b'<offer' in data[:2000]):
        items = parse_yml(data)
        source = YML_URL
 
    # 2) Резерв — Excel
    if not items:
        print("⚠️  YML недоступен, пробуем Excel...")
        data = download(XLS_URL)
        if data and data[:4] in (b'PK\x03\x04', b'\xd0\xcf\x11\xe0'):
            import io as _io
            items = parse_xlsx(data)
            source = XLS_URL
 
    # 3) Оба недоступны
    if not items:
        print("\n" + "="*60)
        print("❌ Оба источника недоступны (IP не в whitelist)")
        print("Напишите в Al-Style: попросите whitelist для GitHub Actions")
        print("="*60)
        # Не падаем — сохраняем статус
        os.makedirs(OUT_DIR, exist_ok=True)
        with open(f"{OUT_DIR}/alstyle_status.json", "w") as f:
            json.dump({"error": True, "checked_at": datetime.now(ALMATY).strftime("%d.%m.%Y %H:%M")}, f)
        sys.exit(0)
 
    save(items, source)
    print(f"🎉 Готово! Источник: {source}")
 
 
# ═══════════════════════════════════════════════════════════════
# КУРС НБК — дополнительно сохраняем в cloud-data/nbk_rate.json
# ═══════════════════════════════════════════════════════════════
def fetch_nbk_rate():
    print("\n💱 Получаем курс USD/KZT с НБК...")
    sources = [
        # RSS лента НБК
        {
            "name": "НБК RSS",
            "url":  "https://nationalbank.kz/rss/rates_all.xml",
            "type": "xml"
        },
        # Open Data API НБК
        {
            "name": "НБК Open Data",
            "url":  "https://data.nationalbank.kz/api/v1/rates?currencyId=USD&language=ru",
            "type": "json"
        },
    ]
 
    for src in sources:
        try:
            r = requests.get(src["url"], timeout=20,
                             headers={"User-Agent": "Mozilla/5.0", "Accept": "*/*"})
            print(f"   [{src['name']}] HTTP {r.status_code}")
            if r.status_code != 200:
                continue
 
            if src["type"] == "xml":
                rate = parse_nbk_xml(r.text)
            else:
                rate = parse_nbk_json(r.json())
 
            if rate and rate > 100:
                print(f"   ✅ Курс USD/KZT = {rate}")
                return rate
 
        except Exception as e:
            print(f"   ⚠️  {src['name']}: {e}")
 
    print("   ❌ Не удалось получить курс НБК")
    return None
 
 
def parse_nbk_xml(xml_text):
    """Парсит RSS XML НБК: находит USD и извлекает курс"""
    try:
        root = ET.fromstring(xml_text.encode("utf-8"))
        for item in root.iter("item"):
            # Ищем <index>USD</index>
            idx = item.find("index")
            if idx is not None and idx.text and idx.text.strip() == "USD":
                for tag in ["description", "quant", "exchangedate"]:
                    el = item.find(tag)
                    if el is not None and el.text:
                        val = float(''.join(c for c in el.text if c.isdigit() or c == '.'))
                        if val > 100:
                            return val
            # Или ищем в <title>: "1 USD = 487.35 KZT"
            title = item.find("title")
            if title is not None and title.text and "USD" in title.text:
                import re
                m = re.search(r'=\s*([\d.,]+)', title.text)
                if m:
                    val = float(m.group(1).replace(",", "."))
                    if val > 100:
                        return val
    except Exception as e:
        print(f"   XML parse error: {e}")
    return None
 
 
def parse_nbk_json(data):
    """Парсит JSON ответ API НБК"""
    try:
        arr = data.get("rates") or data.get("data") or data.get("items") or []
        if isinstance(arr, list):
            for item in arr:
                cur = (item.get("currency") or item.get("currencyId") or "").upper()
                if cur == "USD":
                    val = item.get("rateToTenge") or item.get("rate") or item.get("value")
                    if val:
                        return float(val)
    except Exception as e:
        print(f"   JSON parse error: {e}")
    return None
 
 
def save_nbk_rate(rate):
    os.makedirs(OUT_DIR, exist_ok=True)
    now = datetime.now(ALMATY).strftime("%d.%m.%Y %H:%M")
    out = {"currency": "USD", "rate_kzt": rate, "updated_at": now, "source": "nationalbank.kz"}
    with open(f"{OUT_DIR}/nbk_rate.json", "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False)
    print(f"   💾 Сохранено: cloud-data/nbk_rate.json → {rate} ₸")
 
 
# Вызываем после основного синка
rate = fetch_nbk_rate()
if rate:
    save_nbk_rate(rate)
