# -*- coding: utf-8 -*-
"""
AliExpress Product Scraper
==========================
أداة لاستخراج بيانات المنتجات من أي صفحة داخل AliExpress
(SuperDeals / Bundle Deals / Choice / Flash Deals / نتائج البحث / أي صفحة قائمة منتجات).

الاستخدام:
    python scraper.py
    ثم أدخل رابط الصفحة وعدد المنتجات المطلوب.

يعتمد على Playwright لأنه ينفذ JavaScript ويتعامل مع Lazy Loading / Infinite Scroll.
"""

import csv
import json
import re
import sys
import time

from playwright.sync_api import sync_playwright

try:
    from openpyxl import Workbook
except ImportError:  # openpyxl مطلوب فقط لملف Excel
    Workbook = None

# ---------------------------------------------------------------------------
# الإعدادات العامة
# ---------------------------------------------------------------------------

# أسماء ملفات الإخراج
CSV_FILE = "products.csv"
XLSX_FILE = "products.xlsx"
JSON_FILE = "products.json"

# أعمدة السجل النهائي (بنفس الترتيب المطلوب)
COLUMNS = [
    "Product Name",
    "Price",
    "Product URL",
    "Image URL",
    "Store",
    "Discount",
    "Orders",
    "Rating",
]

# تشغيل المتصفح بواجهة مرئية (False) أفضل لتجنب حماية البوتات في AliExpress.
# يمكن تمرير --headless من سطر الأوامر لتشغيله مخفياً.
HEADLESS = "--headless" in sys.argv

# أقصى عدد من محاولات التمرير قبل الاستسلام (حماية من الحلقات اللانهائية)
MAX_SCROLL_ROUNDS = 60

# User-Agent واقعي لتقليل احتمالية الحظر
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"
)

# ---------------------------------------------------------------------------
# كود JavaScript يُنفَّذ داخل الصفحة لاستخراج كل المنتجات دفعة واحدة.
# الفكرة: لا نعتمد على Selector واحد، بل:
#   1) قائمة Selectors احتياطية لبطاقات المنتجات المعروفة.
#   2) خطة بديلة عامة: أي رابط يحتوي على "/item/" يعتبر منتجاً،
#      ثم نصعد للحاوية الأقرب ونستخرج منها البيانات.
# لذلك تبقى الأداة تعمل حتى لو غيّرت AliExpress أسماء الكلاسات.
# ---------------------------------------------------------------------------

EXTRACT_JS = r"""
() => {
    // --- معرّفات موجودة في رابط الصفحة نفسها (معرّف الحملة مثل 300000512) ---
    // نستبعدها حتى لا تُستخدم كمعرّف منتج فيتحول رابط كل المنتجات إلى رابط الحملة الموحّد.
    const pageIds = new Set(
        ((location.pathname + location.search).match(/\d{5,}/g) || [])
    );

    // --- استخراج معرّف المنتج من رابط: يدعم /item/ و /i/ ومعاملات productId ---
    const ITEM_RE = /(?:\/item\/|\/i\/)(\d{6,})|[?&](?:productId|itemId|objectId|product_id|item_id)=(\d{6,})/i;
    const idFromHref = (href) => {
        if (!href) return null;
        const m = href.match(ITEM_RE);
        const id = m ? (m[1] || m[2]) : null;
        return (id && !pageIds.has(id)) ? id : null;
    };

    // --- إيجاد معرّف منتج حقيقي لبطاقة واحدة من عدة مصادر بالترتيب ---
    const DATA_ID_ATTRS = [
        'data-product-id', 'data-productid', 'data-item-id', 'data-itemid',
        'data-p-id', 'data-pid', 'data-id', 'data-spm-anchor-id',
    ];
    const findProductId = (card) => {
        if (!card) return null;
        // 1) البطاقة نفسها إن كانت رابطاً
        if (card.tagName === 'A') {
            const id = idFromHref(card.href);
            if (id) return id;
        }
        // 2) أي رابط منتج داخل البطاقة
        for (const a of card.querySelectorAll('a[href]')) {
            const id = idFromHref(a.href);
            if (id) return id;
        }
        // 3) خصائص data-* التي تحمل المعرّف
        const candidates = [card, ...card.querySelectorAll('[' + DATA_ID_ATTRS.join('],[') + ']')];
        for (const el of candidates) {
            if (!el.getAttribute) continue;
            for (const attr of DATA_ID_ATTRS) {
                const v = (el.getAttribute(attr) || '').match(/\d{6,}/);
                if (v && !pageIds.has(v[0])) return v[0];
            }
        }
        // 4) مسح HTML البطاقة بحثاً عن productId/itemId (بيانات مضمّنة في JSON)
        const html = card.outerHTML || '';
        const re = /(?:"?(?:productId|itemId|product_id|item_id)"?\s*[:=]\s*"?|\/item\/)(\d{9,})/gi;
        let m;
        while ((m = re.exec(html))) {
            if (!pageIds.has(m[1])) return m[1];
        }
        return null;
    };

    // للتوافق مع الكود القديم: معرّف مبني على الرابط فقط
    const getId = (el) => idFromHref((el && el.href) || '');
    const productLinks = (root) =>
        [...root.querySelectorAll('a[href]')].filter(a => getId(a));

    // --- Selectors احتياطية لبطاقات المنتجات (تُجرَّب بالترتيب) ---
    const CARD_SELECTORS = [
        'div.search-item-card-wrapper-gallery',   // نتائج البحث (شبكة)
        'a.search-card-item',                      // نتائج البحث (بطاقة رابط)
        'div[class*="multi--outWrapper"]',         // تصميم قديم لنتائج البحث
        'div[class*="card-out-wrapper"]',          // SuperDeals / حملات
        'div[class*="productContainer"]',          // صفحات الحملات
        'div[class*="product-card"]',              // عام
        'div[class*="ProductCard"]',               // عام
        'li[class*="product"]',                    // قوائم
    ];

    // --- Selectors احتياطية للحقول داخل البطاقة ---
    const NAME_SELECTORS = [
        'h1', 'h2', 'h3',
        '[class*="titleText"]', '[class*="title--"]',
        '[class*="product-title"]', '[class*="name"]', '[title]',
    ];
    const PRICE_SELECTORS = [
        '[class*="price-sale"]', '[class*="salePrice"]',
        '[class*="price--current"]', '[class*="currentPrice"]',
        '[class*="Price"]', '[class*="price"]',
    ];
    const STORE_SELECTORS = [
        'a[href*="/store/"]',
        '[class*="storeName"]', '[class*="store-name"]', '[class*="shopName"]',
    ];

    // نص العنصر الأول المطابق لأي Selector من القائمة
    const pickText = (root, selectors) => {
        for (const sel of selectors) {
            try {
                const el = root.querySelector(sel);
                if (el) {
                    const t = (el.innerText || el.getAttribute('title') || '').trim();
                    if (t) return t;
                }
            } catch (e) { /* Selector غير صالح -> نجرب التالي */ }
        }
        return '';
    };

    // --- جمع البطاقات: أولاً بالـ Selectors المعروفة ---
    let cards = [];
    for (const sel of CARD_SELECTORS) {
        const found = document.querySelectorAll(sel);
        // نتأكد أن البطاقات تحتوي فعلاً على معرّف منتج قابل للاستخراج
        const valid = [...found].filter(c => findProductId(c));
        if (valid.length >= 3) { cards = valid; break; }
    }

    // --- خطة بديلة عامة: نبدأ من روابط المنتجات ونصعد للحاوية ---
    if (cards.length === 0) {
        const seen = new Set();
        for (const a of productLinks(document)) {
            // نصعد حتى نجد حاوية معقولة الحجم (بطاقة منتج)
            let node = a, card = a;
            for (let i = 0; i < 6 && node.parentElement; i++) {
                node = node.parentElement;
                // إذا أصبحت الحاوية تضم أكثر من منتج فقد تجاوزنا حدود البطاقة
                const ids = new Set(productLinks(node).map(getId));
                if (ids.size > 1) break;
                card = node;
            }
            if (!seen.has(card)) { seen.add(card); cards.push(card); }
        }
    }

    // --- استخراج البيانات من كل بطاقة ---
    const results = [];
    for (const card of cards) {
        try {
            // معرّف المنتج الحقيقي (وليس معرّف الحملة) ثم نبني منه رابطاً مباشراً
            const pid = findProductId(card);
            if (!pid) continue;
            const href = 'https://www.aliexpress.com/item/' + pid + '.html';

            const text = card.innerText || '';

            // تجاهل الإعلانات والعناصر الترويجية
            if (/\b(ad|sponsored|إعلان|ممول)\b/i.test(text.slice(0, 60))) continue;

            // الصورة الرئيسية (مع دعم التحميل الكسول data-src / srcset)
            const img = card.querySelector('img');
            let image = '';
            if (img) {
                image = img.currentSrc || img.src ||
                        img.getAttribute('data-src') ||
                        (img.getAttribute('srcset') || '').split(' ')[0] || '';
            }

            // الاسم: من الـ Selectors ثم من alt الصورة ثم من أطول سطر نصي
            let name = pickText(card, NAME_SELECTORS);
            if (!name && img) name = (img.alt || '').trim();
            if (!name) {
                const lines = text.split('\n').map(s => s.trim()).filter(Boolean);
                name = lines.sort((a, b) => b.length - a.length)[0] || '';
            }

            results.push({
                id: pid,
                name: name,
                url: href,
                image: image,
                price_raw: pickText(card, PRICE_SELECTORS),
                store: pickText(card, STORE_SELECTORS),
                text: text,   // النص الكامل لتحليل السعر/الخصم/الطلبات/التقييم في بايثون
            });
        } catch (e) { /* منتج به مشكلة -> نتجاوزه ونكمل */ }
    }
    return results;
}
"""

# ---------------------------------------------------------------------------
# دوال مساعدة لتنظيف وتحليل النصوص (تعمل على النص الكامل للبطاقة)
# ---------------------------------------------------------------------------

# نمط سعر: عملة + رقم مثل "US $12.99" أو "12,99 €" أو "SAR 45.00"
PRICE_RE = re.compile(
    r"(?:US\s?\$|\$|€|£|¥|₽|SAR|AED|EGP|USD|EUR|ر\.س|د\.إ|ج\.م)\s?"
    r"\d[\d.,]*|\d[\d.,]*\s?(?:US\s?\$|\$|€|£|ر\.س|د\.إ|ج\.م)",
    re.IGNORECASE,
)
# نمط خصم مثل "-30%" أو "30% off" أو "خصم 30%"
DISCOUNT_RE = re.compile(r"-?\s?(\d{1,2})\s?%")
# نمط عدد الطلبات/المبيعات مثل "1,000+ sold" أو "500 orders" أو "تم بيع 100"
# اللاحقة (?<![\d.,]) تمنع الالتقاط الخاطئ عندما يلتصق رقم التقييم برقم المبيعات
ORDERS_RE = re.compile(
    r"(?<![\d.,])(\d[\d.,]*\s?[KkMm]?\+?)\s*(?:sold|orders|pcs sold|قطعة|طلب|تم البيع|مبيع)",
    re.IGNORECASE,
)
# نمط تقييم مثل "4.8" (من 0.0 إلى 5.0)
RATING_RE = re.compile(r"(?<![\d.])([0-5](?:\.\d))(?![\d.%])")


def parse_price(raw_price: str, full_text: str) -> str:
    """استخراج السعر: من حقل السعر إن وجد، وإلا من النص الكامل للبطاقة."""
    for source in (raw_price, full_text):
        if not source:
            continue
        m = PRICE_RE.search(source.replace("\n", " "))
        if m:
            return m.group(0).strip()
    # أحياناً يكون السعر رقماً بدون رمز عملة في حقل السعر
    if raw_price:
        m = re.search(r"\d[\d.,]*", raw_price)
        if m:
            return m.group(0)
    return ""


def parse_discount(full_text: str) -> str:
    """استخراج نسبة الخصم (إن وجدت)."""
    m = DISCOUNT_RE.search(full_text)
    return f"{m.group(1)}%" if m else ""


def parse_orders(full_text: str) -> str:
    """استخراج عدد الطلبات/المبيعات (إن وجد) — سطراً بسطر لتفادي التداخل مع أرقام أخرى."""
    for line in full_text.split("\n"):
        m = ORDERS_RE.search(line)
        if m:
            return m.group(1).strip()
    return ""


def parse_rating(full_text: str) -> str:
    """استخراج التقييم (إن وجد) — رقم بين 0.0 و 5.0."""
    for line in full_text.split("\n"):
        line = line.strip()
        # سطر قصير يحتوي رقم تقييم فقط مثل "4.8" هو الأكثر موثوقية
        if re.fullmatch(r"[0-5]\.\d", line):
            return line
    m = RATING_RE.search(full_text)
    return m.group(1) if m else ""


def normalize_url(url: str) -> str:
    """توحيد رابط المنتج (إزالة معاملات التتبع) لاستخدامه في إزالة التكرار."""
    url = url.split("?")[0]
    if url.startswith("//"):
        url = "https:" + url
    return url


def build_product_url(raw: dict) -> str:
    """
    بناء رابط مباشر لكل منتج.
    في صفحات الحملات (Bundle/SuperDeals) يكون رابط <a> هو رابط الحملة الموحّد نفسه،
    لذلك نعتمد على معرّف المنتج (id) لبناء رابط /item/{id}.html خاص بكل منتج.
    نعود إلى الرابط الخام فقط إذا لم يتوفر المعرّف.
    """
    pid = str(raw.get("id", "") or "")
    if re.fullmatch(r"\d{6,}", pid):
        return f"https://www.aliexpress.com/item/{pid}.html"
    url = normalize_url(raw.get("url", ""))
    # كخطة أخيرة: حاول انتزاع معرّف منتج من الرابط الخام
    m = re.search(r"(?:/item/|/i/)(\d{6,})", url)
    if m:
        return f"https://www.aliexpress.com/item/{m.group(1)}.html"
    return url


def clean_record(raw: dict) -> dict:
    """تحويل البيانات الخام من المتصفح إلى سجل نهائي منظم."""
    text = raw.get("text", "")
    return {
        "Product Name": raw.get("name", "").strip(),
        "Price": parse_price(raw.get("price_raw", ""), text),
        "Product URL": build_product_url(raw),
        "Image URL": ("https:" + raw["image"]) if raw.get("image", "").startswith("//") else raw.get("image", ""),
        "Store": raw.get("store", "").strip(),
        "Discount": parse_discount(text),
        "Orders": parse_orders(text),
        "Rating": parse_rating(text),
    }


# ---------------------------------------------------------------------------
# الطبقة الثالثة: التقاط بيانات المنتجات من استجابات الشبكة (JSON)
# صفحات الحملات (SuperDeals / BundleDeals / Choice / best.aliexpress.com)
# ترسم البطاقات بجافاسكريبت وتجلب البيانات عبر طلبات API داخلية،
# لذلك نلتقط تلك الاستجابات مباشرة — وهذا يعمل حتى لو لم توجد روابط في الصفحة.
# ---------------------------------------------------------------------------

_ID_KEYS = {"productid", "itemid", "productidstr", "itemidstr", "objectid"}
_TITLE_KEYS = {"title", "subject", "producttitle", "itemtitle", "displaytitle", "name"}


def parse_maybe_jsonp(body: str):
    """تحويل نص JSON أو JSONP (مثل mtopjsonp1({...})) إلى كائن بايثون."""
    body = body.strip()
    try:
        return json.loads(body)
    except (ValueError, TypeError):
        pass
    m = re.match(r"^[\w$.]+\s*\(\s*(\{.*\})\s*\)\s*;?$", body, re.S)
    if m:
        try:
            return json.loads(m.group(1))
        except (ValueError, TypeError):
            return None
    return None


def harvest_json_products(obj, bucket: list, depth: int = 0):
    """البحث المتكرر داخل JSON عن كائنات تشبه المنتجات (معرّف + عنوان)."""
    if depth > 12:
        return
    if isinstance(obj, dict):
        pid = title = None
        for k, v in obj.items():
            kl = k.lower()
            if kl in _ID_KEYS and isinstance(v, (str, int)) and re.fullmatch(r"\d{5,}", str(v)):
                pid = str(v)
            elif kl in _TITLE_KEYS and isinstance(v, str) and len(v.strip()) > 5:
                title = v.strip()
        if pid and title:
            bucket.append((pid, title, obj))
        for v in obj.values():
            if isinstance(v, (dict, list)):
                harvest_json_products(v, bucket, depth + 1)
    elif isinstance(obj, list):
        for v in obj:
            if isinstance(v, (dict, list)):
                harvest_json_products(v, bucket, depth + 1)


def _iter_pairs(obj, depth: int = 0):
    """المرور على كل أزواج (اسم الحقل، القيمة) داخل كائن JSON متداخل."""
    if depth > 10:
        return
    if isinstance(obj, dict):
        for k, v in obj.items():
            if isinstance(v, (dict, list)):
                yield from _iter_pairs(v, depth + 1)
            elif isinstance(v, (str, int, float)) and not isinstance(v, bool):
                yield k.lower(), str(v)
    elif isinstance(obj, list):
        for v in obj:
            if isinstance(v, (dict, list)):
                yield from _iter_pairs(v, depth + 1)
            elif isinstance(v, (str, int, float)) and not isinstance(v, bool):
                yield "", str(v)


def sniffed_to_record(pid: str, title: str, obj: dict) -> dict:
    """تحويل كائن منتج ملتقط من الشبكة إلى سجل نهائي بنفس الأعمدة."""
    pairs = list(_iter_pairs(obj))

    def first(cond):
        for k, v in pairs:
            v = v.strip()
            if v and cond(k, v):
                return k, v
        return "", ""

    # الصورة: أي رابط لخوادم صور AliExpress
    _, image = first(lambda k, v: "alicdn.com" in v)
    if image.startswith("//"):
        image = "https:" + image

    # السعر: نفضل الحقول التي يوحي اسمها بالسعر، ثم أي قيمة بنمط سعر
    _, price = first(lambda k, v: ("price" in k or "amount" in k) and PRICE_RE.search(v))
    if not price:
        _, price = first(lambda k, v: bool(PRICE_RE.search(v)))
    if price:
        price = PRICE_RE.search(price).group(0).strip()

    # الخصم
    _, discount = first(lambda k, v: "discount" in k and re.fullmatch(r"-?\d{1,2}%?", v))
    if discount:
        discount = discount.lstrip("-")
        if not discount.endswith("%"):
            discount += "%"

    # الطلبات / المبيعات
    _, orders = first(lambda k, v: any(s in k for s in ("sold", "order", "trade")) and re.search(r"\d", v))
    if orders:
        m = ORDERS_RE.search(orders)
        if m:
            orders = m.group(1).strip()
        else:
            m = re.search(r"\d[\d.,]*\s?[KkMm]?\+?", orders)
            orders = m.group(0).strip() if m else ""

    # التقييم
    _, rating = first(
        lambda k, v: any(s in k for s in ("rating", "star", "evaluat"))
        and re.fullmatch(r"[0-5](\.\d+)?", v)
    )

    # المتجر
    _, store = first(lambda k, v: any(s in k for s in ("storename", "shopname", "sellername")))

    return {
        "Product Name": title,
        "Price": price,
        "Product URL": f"https://www.aliexpress.com/item/{pid}.html",
        "Image URL": image,
        "Store": store,
        "Discount": discount,
        "Orders": orders,
        "Rating": rating,
    }


def attach_network_sniffer(page, bucket: list):
    """مراقبة استجابات الشبكة والتقاط أي بيانات منتجات فيها."""

    def on_response(response):
        try:
            if response.request.resource_type not in ("xhr", "fetch", "script"):
                return
            body = response.text()
            if not body or len(body) > 3_000_000:
                return
            # تصفية سريعة قبل التحليل الكامل
            if "roductId" not in body and "temId" not in body:
                return
            data = parse_maybe_jsonp(body)
            if data is not None:
                harvest_json_products(data, bucket)
        except Exception:
            pass  # فشل قراءة استجابة واحدة لا يوقف البرنامج

    page.on("response", on_response)


def product_key(url: str) -> str:
    """مفتاح إزالة التكرار: معرّف المنتج إن وجد في الرابط."""
    m = re.search(r"(\d{5,})", url)
    return m.group(1) if m else url


# ---------------------------------------------------------------------------
# إدخال المستخدم
# ---------------------------------------------------------------------------

def ask_user_inputs():
    """طلب رابط الصفحة وعدد المنتجات من المستخدم."""
    url = input("أدخل رابط الصفحة: ").strip()
    while not url.startswith("http"):
        print("⚠️  الرابط غير صالح، يجب أن يبدأ بـ http أو https")
        url = input("أدخل رابط الصفحة: ").strip()

    while True:
        try:
            count = int(input("أدخل عدد المنتجات: ").strip())
            if count > 0:
                return url, count
            print("⚠️  أدخل رقماً أكبر من صفر")
        except ValueError:
            print("⚠️  أدخل رقماً صحيحاً")


# ---------------------------------------------------------------------------
# التصفح والاستخراج
# ---------------------------------------------------------------------------

def launch_browser(playwright):
    """
    تشغيل المتصفح مع خطط بديلة:
      1) متصفح Chromium الخاص بـ Playwright (إن كان منزّلاً).
      2) Google Chrome المثبت على الجهاز.
      3) Microsoft Edge (موجود افتراضياً في كل أجهزة ويندوز).
    بهذا يعمل البرنامج حتى لو فشل أمر: playwright install chromium
    """
    attempts = [
        ("Playwright Chromium", {}),
        ("Google Chrome", {"channel": "chrome"}),
        ("Microsoft Edge", {"channel": "msedge"}),
    ]
    last_error = None
    for label, opts in attempts:
        try:
            browser = playwright.chromium.launch(headless=HEADLESS, **opts)
            print(f"🌐 المتصفح المستخدم: {label}")
            return browser
        except Exception as exc:
            last_error = exc
    raise RuntimeError(
        "تعذر تشغيل أي متصفح. ثبّت Google Chrome أو Microsoft Edge، "
        "أو نفّذ: playwright install chromium\n"
        f"تفاصيل آخر خطأ: {last_error}"
    )


def open_page(playwright, url: str):
    """فتح المتصفح والانتقال إلى الصفحة المطلوبة مع تفعيل التقاط بيانات الشبكة."""
    browser = launch_browser(playwright)
    context = browser.new_context(
        user_agent=USER_AGENT,
        viewport={"width": 1366, "height": 900},
        locale="en-US",
    )
    page = context.new_page()
    # التقاط بيانات المنتجات من طلبات JSON الداخلية (مهم لصفحات الحملات)
    sniffed = []
    attach_network_sniffer(page, sniffed)
    print("⏳ جاري فتح الصفحة ...")
    page.goto(url, timeout=90_000, wait_until="domcontentloaded")
    # ننتظر قليلاً حتى تنفذ الصفحة الـ JavaScript وتعرض المنتجات
    page.wait_for_timeout(6000)
    return browser, page, sniffed


def collect_products(page, target_count: int, sniffed: list = None) -> list:
    """
    جمع المنتجات من ثلاث طبقات:
      1) بطاقات الصفحة (DOM) في الصفحة الرئيسية وكل الإطارات (iframes).
      2) بيانات JSON الملتقطة من طلبات الشبكة (sniffed).
    مع التمرير التلقائي (Lazy Loading / Infinite Scroll)
    حتى الوصول للعدد المطلوب أو انتهاء محتوى الصفحة.
    """
    products = {}          # المفتاح: معرّف المنتج -> إزالة التكرار تلقائياً
    last_reported = 0      # لعرض التقدم دون تكرار نفس السطر
    stagnant_rounds = 0    # عدد جولات التمرير بدون منتجات جديدة

    def add_record(record):
        """إضافة سجل مع إزالة التكرار حسب معرّف المنتج."""
        if not record["Product Name"] or not record["Product URL"]:
            return
        key = product_key(record["Product URL"])
        if key not in products and len(products) < target_count:
            products[key] = record

    for round_no in range(MAX_SCROLL_ROUNDS):
        before = len(products)

        # الطبقة 1: استخراج بطاقات الصفحة من كل الإطارات
        raw_items = []
        for frame in page.frames:
            try:
                raw_items.extend(frame.evaluate(EXTRACT_JS))
            except Exception:
                continue  # إطار لم يكتمل تحميله -> نتجاوزه

        for raw in raw_items:
            if len(products) >= target_count:
                break
            try:
                add_record(clean_record(raw))
            except Exception:
                # خطأ في منتج واحد -> نتجاوزه ونكمل
                continue

        # الطبقة 2: المنتجات الملتقطة من استجابات الشبكة (JSON)
        if sniffed:
            seen_ids = set()
            for pid, title, obj in list(sniffed):
                if len(products) >= target_count:
                    break
                if pid in seen_ids:
                    continue
                seen_ids.add(pid)
                try:
                    add_record(sniffed_to_record(pid, title, obj))
                except Exception:
                    continue

        # عرض التقدم كل 20 منتجاً جديداً تقريباً
        current = len(products)
        if current - last_reported >= 20 or (current and current >= target_count):
            print(f"تم استخراج {min(current, target_count)} من {target_count}")
            last_reported = current

        if current >= target_count:
            break

        # هل أضفنا منتجات جديدة في هذه الجولة؟
        stagnant_rounds = stagnant_rounds + 1 if current == before else 0
        if current == 0 and stagnant_rounds == 3:
            print("⏳ الصفحة بطيئة أو ما زالت تُحمّل — نواصل الانتظار والتمرير ...")
        if stagnant_rounds >= 8:
            print("ℹ️  لا توجد منتجات جديدة بعد عدة محاولات تمرير — نكتفي بما جُمع.")
            break

        # التمرير للأسفل لتحميل المزيد من المنتجات
        page.mouse.wheel(0, 2500)
        page.wait_for_timeout(2000)

        # بعض الصفحات تستخدم زر "عرض المزيد" بدلاً من التمرير اللانهائي
        try:
            more = page.locator(
                "button:has-text('View more'), button:has-text('Show more'), "
                "button:has-text('عرض المزيد'), [class*='loadMore'], [class*='load-more']"
            ).first
            if more.is_visible(timeout=500):
                more.click()
                page.wait_for_timeout(2000)
        except Exception:
            pass  # لا يوجد زر — نكمل بالتمرير العادي

    # نقتطع العدد المطلوب بالضبط
    return list(products.values())[:target_count]


# ---------------------------------------------------------------------------
# حفظ النتائج
# ---------------------------------------------------------------------------

def save_csv(products: list, path: str = CSV_FILE):
    """حفظ النتائج في ملف CSV."""
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=COLUMNS)
        writer.writeheader()
        writer.writerows(products)
    print(f"✅ تم الحفظ في {path}")


def save_xlsx(products: list, path: str = XLSX_FILE):
    """حفظ النتائج في ملف Excel."""
    if Workbook is None:
        print("⚠️  مكتبة openpyxl غير مثبتة — تم تخطي ملف Excel. ثبّتها بـ: pip install openpyxl")
        return
    wb = Workbook()
    ws = wb.active
    ws.title = "Products"
    ws.append(COLUMNS)
    for p in products:
        ws.append([p.get(col, "") for col in COLUMNS])
    wb.save(path)
    print(f"✅ تم الحفظ في {path}")


def save_json(products: list, path: str = JSON_FILE):
    """حفظ النتائج في ملف JSON."""
    with open(path, "w", encoding="utf-8") as f:
        json.dump(products, f, ensure_ascii=False, indent=2)
    print(f"✅ تم الحفظ في {path}")


# ---------------------------------------------------------------------------
# نقطة البداية
# ---------------------------------------------------------------------------

def main():
    print("=" * 50)
    print("   AliExpress Product Scraper 🛒")
    print("=" * 50)

    # يمكن تمرير الرابط والعدد مباشرة لتخطي الأسئلة:
    #   python scraper.py <الرابط> <العدد>
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if len(args) >= 2 and args[0].startswith("http") and args[1].isdigit():
        url, count = args[0], int(args[1])
    else:
        url, count = ask_user_inputs()

    start = time.time()
    with sync_playwright() as playwright:
        browser, page, sniffed = open_page(playwright, url)
        try:
            products = collect_products(page, count, sniffed)
        finally:
            browser.close()

    if not products:
        print("❌ لم يتم العثور على أي منتجات. تأكد من الرابط أو جرّب بدون --headless.")
        sys.exit(1)

    print(f"\n📦 إجمالي المنتجات المستخرجة: {len(products)} (خلال {time.time() - start:.0f} ثانية)")

    # الحفظ في الملفات الثلاثة
    save_csv(products)
    save_xlsx(products)
    save_json(products)

    print("\n🎉 انتهت العملية بنجاح!")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n⛔ تم إيقاف البرنامج من قبل المستخدم.")
