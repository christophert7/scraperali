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
        // نتأكد أن البطاقات تحتوي فعلاً على روابط منتجات
        const valid = [...found].filter(c =>
            (c.matches && c.matches('a[href*="/item/"]')) ||
            c.querySelector('a[href*="/item/"]'));
        if (valid.length >= 3) { cards = valid; break; }
    }

    // --- خطة بديلة عامة: نبدأ من روابط "/item/" ونصعد للحاوية ---
    if (cards.length === 0) {
        const seen = new Set();
        for (const a of document.querySelectorAll('a[href*="/item/"]')) {
            // نصعد حتى نجد حاوية معقولة الحجم (بطاقة منتج)
            let node = a, card = a;
            for (let i = 0; i < 6 && node.parentElement; i++) {
                node = node.parentElement;
                const links = node.querySelectorAll('a[href*="/item/"]');
                // إذا أصبحت الحاوية تضم أكثر من منتج فقد تجاوزنا حدود البطاقة
                const ids = new Set([...links].map(l => (l.href.match(/item\/(\d+)/) || [])[1]));
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
            // رابط المنتج
            const link = card.matches && card.matches('a[href*="/item/"]')
                ? card
                : card.querySelector('a[href*="/item/"]');
            if (!link) continue;
            const href = link.href || '';
            const idMatch = href.match(/item\/(\d+)/);
            if (!idMatch) continue;

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
                id: idMatch[1],
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


def clean_record(raw: dict) -> dict:
    """تحويل البيانات الخام من المتصفح إلى سجل نهائي منظم."""
    text = raw.get("text", "")
    return {
        "Product Name": raw.get("name", "").strip(),
        "Price": parse_price(raw.get("price_raw", ""), text),
        "Product URL": normalize_url(raw.get("url", "")),
        "Image URL": ("https:" + raw["image"]) if raw.get("image", "").startswith("//") else raw.get("image", ""),
        "Store": raw.get("store", "").strip(),
        "Discount": parse_discount(text),
        "Orders": parse_orders(text),
        "Rating": parse_rating(text),
    }


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

def open_page(playwright, url: str):
    """فتح المتصفح والانتقال إلى الصفحة المطلوبة."""
    browser = playwright.chromium.launch(headless=HEADLESS)
    context = browser.new_context(
        user_agent=USER_AGENT,
        viewport={"width": 1366, "height": 900},
        locale="en-US",
    )
    page = context.new_page()
    print("⏳ جاري فتح الصفحة ...")
    page.goto(url, timeout=90_000, wait_until="domcontentloaded")
    # ننتظر قليلاً حتى تنفذ الصفحة الـ JavaScript وتعرض المنتجات
    page.wait_for_timeout(5000)
    return browser, page


def collect_products(page, target_count: int) -> list:
    """
    جمع المنتجات مع التمرير التلقائي (Lazy Loading / Infinite Scroll)
    حتى الوصول للعدد المطلوب أو انتهاء محتوى الصفحة.
    """
    products = {}          # المفتاح: الرابط الموحّد -> إزالة التكرار تلقائياً
    last_reported = 0      # لعرض التقدم دون تكرار نفس السطر
    stagnant_rounds = 0    # عدد جولات التمرير بدون منتجات جديدة

    for _ in range(MAX_SCROLL_ROUNDS):
        # استخراج كل المنتجات الظاهرة حالياً في الصفحة
        try:
            raw_items = page.evaluate(EXTRACT_JS)
        except Exception as exc:
            print(f"⚠️  خطأ أثناء الاستخراج، سنعيد المحاولة: {exc}")
            page.wait_for_timeout(2000)
            continue

        before = len(products)
        for raw in raw_items:
            if len(products) >= target_count:
                break
            try:
                record = clean_record(raw)
                # نتجاهل السجلات بدون اسم أو رابط (عناصر غير منتجات)
                if not record["Product Name"] or not record["Product URL"]:
                    continue
                key = record["Product URL"]
                if key not in products:
                    products[key] = record
            except Exception:
                # خطأ في منتج واحد -> نتجاوزه ونكمل
                continue

        # عرض التقدم كل 20 منتجاً جديداً تقريباً
        current = len(products)
        if current - last_reported >= 20 or current >= target_count:
            print(f"تم استخراج {min(current, target_count)} من {target_count}")
            last_reported = current

        if current >= target_count:
            break

        # هل أضفنا منتجات جديدة في هذه الجولة؟
        stagnant_rounds = stagnant_rounds + 1 if current == before else 0
        if stagnant_rounds >= 5:
            print("ℹ️  لا توجد منتجات جديدة بعد عدة محاولات تمرير — نكتفي بما جُمع.")
            break

        # التمرير للأسفل لتحميل المزيد من المنتجات
        page.mouse.wheel(0, 2500)
        page.wait_for_timeout(1500)

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

    url, count = ask_user_inputs()

    start = time.time()
    with sync_playwright() as playwright:
        browser, page = open_page(playwright, url)
        try:
            products = collect_products(page, count)
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
