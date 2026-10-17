import os, re, json, asyncio, uuid, io, datetime
import httpx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from PIL import Image, ImageDraw, ImageFont
from telegram import (Update, InlineKeyboardButton, InlineKeyboardMarkup,
                      InputMediaPhoto)
from telegram.ext import (Application, CommandHandler, MessageHandler,
    CallbackQueryHandler, ConversationHandler, ContextTypes, filters)

# ============ SOZLAMALAR ============
BOT_TOKEN = os.environ.get("BOT_TOKEN")
ADMIN_ID = os.environ.get("ADMIN_ID")
GEMINI_KEY = os.environ.get("GEMINI_KEY")
PEXELS_KEY = "58019301-89bcfc740e2f87b23496dd056"          # ixtiyoriy: Pixabay tekin API kaliti (rasmlar uchun)
CARD = "9860 3501 0918 8481 (Abdulbosid Khamidullaev)"
NARX_BET = 1500
MODELS = ["gemini-flash-latest", "gemini-3-flash-preview", "gemini-2.5-flash"]
# =====================================

LANGS = {
    "uz": ("🇺🇿 O'zbekcha", "Uzbek (Latin script)"),
    "ru": ("🇷🇺 Русский", "Russian"),
    "en": ("🇬🇧 English", "English"),
}
NOTE = {
    "uz": "Ma'lumotlar taxminiy va tasviriy",
    "ru": "Данные приблизительные и иллюстративные",
    "en": "Figures are approximate and illustrative",
}

# nomi, fon, sarlavha rangi, matn rangi, aksent rangi
THEMES = {
    "1": ("Okean",         "12263F", "FFFFFF", "DCE6F5", "FFD23F"),
    "2": ("Oq-minimal",    "FFFFFF", "12263F", "3A4A5F", "2F5BFF"),
    "3": ("Tabiat",        "0F3D2E", "FFFFFF", "D6EBDD", "7BD389"),
    "4": ("Quyosh",        "FFF4E0", "7A2E00", "5A3A1E", "FF8A00"),
    "5": ("Binafsha",      "2A1458", "FFFFFF", "E3D9F7", "B28DFF"),
    "6": ("Qora-oltin",    "111111", "F5D77A", "E6E6E6", "D4AF37"),
    "7": ("Qizil-kulrang", "ECEFF3", "8C1C2B", "2E3440", "C8102E"),
}

NAME, TOPIC, PAGES, LANG, THEME, CHECK = range(6)
DB = "orders.json"
PREVIEWS = {}
STATUS_ICON = {"pending": "⏳", "working": "⚙️", "done": "✅",
               "rejected": "❌", "error": "⚠️"}


# ============ BUYURTMALAR BAZASI ============
def load():
    try:
        with open(DB, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save(d):
    with open(DB, "w", encoding="utf-8") as f:
        json.dump(d, f, ensure_ascii=False, indent=1)


def now_str():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def parse_ts(s):
    try:
        return datetime.datetime.strptime(s, "%Y-%m-%d %H:%M:%S")
    except Exception:
        return None


def money(n):
    return f"{int(n):,}".replace(",", " ")


# ============ SHABLON NAMUNALARI ============
def _font(size):
    for n in ("arial.ttf", "DejaVuSans.ttf"):
        try:
            return ImageFont.truetype(n, size)
        except Exception:
            pass
    return ImageFont.load_default()


def preview(key):
    if key in PREVIEWS:
        return PREVIEWS[key]
    name, bg, tc, xc, ac = THEMES[key]
    im = Image.new("RGB", (640, 360), "#" + bg)
    d = ImageDraw.Draw(im)
    d.rectangle([50, 110, 130, 120], fill="#" + ac)
    d.text((50, 140), "Slayd mavzusi", fill="#" + tc, font=_font(44))
    d.text((50, 215), "Ism Familiya", fill="#" + xc, font=_font(24))
    for i in range(3):
        d.rectangle([50, 275 + i * 20, 350 - i * 50, 283 + i * 20], fill="#" + xc)
    buf = io.BytesIO()
    im.save(buf, "PNG")
    PREVIEWS[key] = buf.getvalue()
    return PREVIEWS[key]


# ============ MIJOZ SUHBATI ============
async def start(u: Update, c: ContextTypes.DEFAULT_TYPE):
    await u.message.reply_text("Assalomu alaykum! F.I.Sh ni kiriting:")
    return NAME


async def get_name(u, c):
    c.user_data["name"] = u.message.text
    await u.message.reply_text("Slayd mavzusi:")
    return TOPIC


async def get_topic(u, c):
    c.user_data["topic"] = u.message.text
    await u.message.reply_text("Necha betlik? (1 dan 30 gacha)")
    return PAGES


async def get_pages(u, c):
    try:
        n = int(u.message.text)
        assert 1 <= n <= 30
    except Exception:
        await u.message.reply_text("1 dan 30 gacha raqam kiriting.")
        return PAGES
    c.user_data["pages"] = n
    c.user_data["summa"] = n * NARX_BET
    kb = InlineKeyboardMarkup([[InlineKeyboardButton(v[0], callback_data=f"lang:{k}")]
                               for k, v in LANGS.items()])
    await u.message.reply_text("Slayd qaysi tilda bo'lsin?", reply_markup=kb)
    return LANG


async def pick_lang(u, c):
    q = u.callback_query
    await q.answer()
    k = q.data.split(":")[1]
    c.user_data["lang"] = k
    await q.edit_message_text(f"Til: {LANGS[k][0]}")
    chat = u.effective_chat.id
    try:
        media = [InputMediaPhoto(preview(t), caption=f"{t}. {THEMES[t][0]}")
                 for t in THEMES]
        await c.bot.send_media_group(chat, media)
    except Exception:
        await c.bot.send_message(
            chat, "Shablonlar: " + ", ".join(f"{t}. {v[0]}" for t, v in THEMES.items()))
    keys = list(THEMES)
    rows = []
    for i in range(0, len(keys), 2):
        rows.append([InlineKeyboardButton(f"{t}. {THEMES[t][0]}", callback_data=f"theme:{t}")
                     for t in keys[i:i + 2]])
    await c.bot.send_message(chat, "Qaysi shablonda yaratamiz?",
                             reply_markup=InlineKeyboardMarkup(rows))
    return THEME


async def pick_theme(u, c):
    q = u.callback_query
    await q.answer()
    t = q.data.split(":")[1]
    c.user_data["theme"] = t
    await q.edit_message_text(f"Shablon: {t}. {THEMES[t][0]}")
    d = c.user_data
    await c.bot.send_message(
        u.effective_chat.id,
        f"Summa: {money(d['summa'])} so'm\n\nKarta: {CARD}\n\n"
        "To'lovni qilib, chek rasmini shu yerga yuboring.")
    return CHECK


def admin_kb(oid, retry=False):
    return InlineKeyboardMarkup([[
        InlineKeyboardButton("🔁 Qayta urinish" if retry else "✅ Tasdiqlash",
                             callback_data=f"ok:{oid}"),
        InlineKeyboardButton("❌ Rad etish", callback_data=f"no:{oid}")]])


async def get_check(u, c):
    if not u.message.photo:
        await u.message.reply_text("Iltimos, chekni rasm sifatida yuboring.")
        return CHECK
    oid = uuid.uuid4().hex[:8]
    d = dict(c.user_data)
    usr = u.effective_user
    d.update({
        "kind": "slide",           # kelajakda: referat, kurs ishi...
        "status": "pending",
        "created": now_str(),
        "chat_id": u.effective_chat.id,
        "user_id": usr.id,
        "username": usr.username or "",
    })
    orders = load()
    orders[oid] = d
    save(orders)
    who = f"@{usr.username}" if usr.username else f"id {usr.id}"
    await c.bot.send_photo(
        ADMIN_ID, u.message.photo[-1].file_id,
        caption=(f"Yangi buyurtma #{oid}\n{d['name']} ({who})\nMavzu: {d['topic']}\n"
                 f"Bet: {d['pages']}\nTil: {LANGS[d['lang']][0]}\n"
                 f"Shablon: {THEMES[d['theme']][0]}\nSumma: {money(d['summa'])} so'm"),
        reply_markup=admin_kb(oid))
    await u.message.reply_text("Chek adminga yuborildi. Tasdiqlangach slayd keladi ⏳")
    return ConversationHandler.END


async def cancel(u, c):
    await u.message.reply_text("Bekor qilindi. /start")
    return ConversationHandler.END


# ============ SLAYD MATNI (GEMINI) ============
def build_prompt(topic, pages, lang):
    return f"""You are an expert presentation designer. Create a presentation on the topic: '{topic}'.
Exactly {pages} slides. Write ALL text in {LANGS[lang][1]}.
Return ONLY a JSON array. Each item is one slide object with a "layout" field:
- "title": first slide only. Fields: title, subtitle (short).
- "bullets": title, points (3 to 5 short bullets, max 14 words each), image (2 to 4 English words for a stock photo search, concrete and visual).
- "chart": title, chart {{"type": "bar" or "pie" or "line", "labels": [3 to 6 strings], "values": [numbers]}}, unit (short, e.g. "%" or "mln"), points (2 to 3 short takeaways).
- "steps": title, steps [3 to 5 items: {{"head": short, "text": max 10 words}}].
- "stats": title, stats [3 to 4 items: {{"value": "short number like 45% or 2.5 mln", "label": "max 6 words"}}], points (1 short takeaway).
Rules: slide 1 uses "title". The last slide is a conclusion using "bullets". If there are 5 or more slides, include at least one "chart", one "steps" and one "stats" slide, and use varied layouts. Numbers must be realistic and plausible. Keep all text concise."""


def clean(slides, pages):
    if isinstance(slides, dict):
        slides = slides.get("slides", [])
    out = []
    for i, s in enumerate(slides[:pages]):
        if not isinstance(s, dict):
            continue
        lay = s.get("layout", "bullets")
        d = {"title": str(s.get("title", "")).strip() or "-",
             "layout": lay,
             "points": [str(p) for p in (s.get("points") or []) if p],
             "image": str(s.get("image", "")).strip()}
        if i == 0:
            d["layout"] = "title"
            d["subtitle"] = str(s.get("subtitle", "")).strip()
        elif lay == "chart":
            try:
                ch = s.get("chart") or {}
                labels = [str(x) for x in ch.get("labels", [])]
                vals = [float(x) for x in ch.get("values", [])]
                assert len(labels) == len(vals) >= 2
                t = ch.get("type", "bar")
                t = t if t in ("bar", "pie", "line") else "bar"
                if t == "pie":
                    vals = [abs(v) for v in vals]
                d["chart"] = {"type": t, "labels": labels[:8], "values": vals[:8]}
                d["unit"] = str(s.get("unit", "")).strip()
            except Exception:
                d["layout"] = "bullets"
        elif lay == "steps":
            st = [x for x in (s.get("steps") or []) if isinstance(x, dict)]
            if len(st) >= 2:
                d["steps"] = [{"head": str(x.get("head", "")), "text": str(x.get("text", ""))}
                              for x in st[:5]]
            else:
                d["layout"] = "bullets"
        elif lay == "stats":
            st = [x for x in (s.get("stats") or []) if isinstance(x, dict)]
            if len(st) >= 2:
                d["stats"] = [{"value": str(x.get("value", "")), "label": str(x.get("label", ""))}
                              for x in st[:4]]
            else:
                d["layout"] = "bullets"
        else:
            d["layout"] = "bullets"
        out.append(d)
    if not out:
        raise Exception("Gemini bo'sh javob qaytardi")
    return out


async def ai_slides(topic, pages, lang):
    body = {"contents": [{"parts": [{"text": build_prompt(topic, pages, lang)}]}],
            "generationConfig": {"responseMimeType": "application/json",
                                 "maxOutputTokens": 12000}}
    r = None
    async with httpx.AsyncClient(timeout=180) as cl:
        for model in MODELS:
            url = ("https://generativelanguage.googleapis.com/v1beta/models/"
                   f"{model}:generateContent")
            for _ in range(3):
                r = await cl.post(url, headers={"x-goog-api-key": GEMINI_KEY}, json=body)
                if r.status_code not in (429, 503):
                    break
                await asyncio.sleep(8)
            if r.status_code != 404:
                break
    if r.status_code != 200:
        raise Exception(f"API xato {r.status_code}: {r.text[:300]}")
    parts = r.json()["candidates"][0]["content"]["parts"]
    text = "".join(p.get("text", "") for p in parts)
    m = re.search(r"\[.*\]", text, re.S)
    return clean(json.loads(m.group(0)), pages)


# ============ RASMLAR (Pexels) ============
IMG_W, IMG_H = 4.7, 4.4


def crop_to(data, ratio):
    im = Image.open(io.BytesIO(data)).convert("RGB")
    w, h = im.size
    if w / h > ratio:
        nw = int(h * ratio)
        im = im.crop(((w - nw) // 2, 0, (w - nw) // 2 + nw, h))
    else:
        nh = int(w / ratio)
        im = im.crop((0, (h - nh) // 2, w, (h - nh) // 2 + nh))
    im = im.resize((900, int(900 / ratio)))
    buf = io.BytesIO()
    im.save(buf, "JPEG", quality=85)
    return buf.getvalue()


async def fetch_image(cl, query):
    try:
        r = await cl.get("https://pixabay.com/api/",
                         params={"key": PEXELS_KEY, "q": query[:90], "image_type": "photo",
                                 "orientation": "horizontal", "per_page": 3,
                                 "safesearch": "true"})
        hits = r.json().get("hits") or []
        if not hits:
            return None
        img = await cl.get(hits[0].get("largeImageURL") or hits[0]["webformatURL"])
        return crop_to(img.content, IMG_W / IMG_H)
    except Exception:
        return None


async def attach_images(slides):
    if not PEXELS_KEY:
        return
    targets = [s for s in slides if s["layout"] == "bullets" and s.get("image")]
    async with httpx.AsyncClient(timeout=25, follow_redirects=True) as cl:
        res = await asyncio.gather(*[fetch_image(cl, s["image"]) for s in targets])
    for s, img in zip(targets, res):
        s["_img"] = img


# ============ PPTX YIG'ISH ============
def blend(a, b, t):
    x = [int(a[i:i + 2], 16) for i in (0, 2, 4)]
    y = [int(b[i:i + 2], 16) for i in (0, 2, 4)]
    return "%02X%02X%02X" % tuple(round(p * (1 - t) + q * t) for p, q in zip(x, y))


def make_pptx(slides, path, author, theme, lang="uz"):
    _, bg, tc, xc, ac = THEMES.get(theme, THEMES["1"])
    rgb = RGBColor.from_string
    card = blend(bg, tc, 0.08)
    muted = blend(bg, tc, 0.55)
    pal = [ac, blend(ac, tc, 0.4), blend(ac, bg, 0.45), blend(tc, bg, 0.45),
           blend(ac, tc, 0.75), blend(ac, bg, 0.75), blend(tc, bg, 0.7), blend(ac, "808080", 0.5)]

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    def new_slide():
        sl = prs.slides.add_slide(prs.slide_layouts[6])
        f = sl.background.fill
        f.solid()
        f.fore_color.rgb = rgb(bg)
        return sl

    def box(sl, shape, x, y, w, h, fill, radius=None):
        s = sl.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
        s.fill.solid()
        s.fill.fore_color.rgb = rgb(fill)
        s.line.fill.background()
        s.shadow.inherit = False
        if radius is not None:
            s.adjustments[0] = radius
        return s

    def text(sl, x, y, w, h, lines, size, color, bold=False, gap=0,
             align=None, anchor=None):
        tb = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        if anchor:
            tf.vertical_anchor = anchor
        for i, line in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            if gap:
                p.space_after = Pt(gap)
            if align:
                p.alignment = align
            r = p.add_run()
            r.text = line
            r.font.size = Pt(size)
            r.font.bold = bold
            r.font.color.rgb = rgb(color)
        return tb

    def head(sl, title, num):
        box(sl, MSO_SHAPE.RECTANGLE, 0.8, 0.7, 1.2, 0.12, ac)
        size = 32 if len(title) <= 55 else 26
        text(sl, 0.8, 0.95, 11.7, 1.2, [title], size, tc, bold=True)
        text(sl, 12.0, 6.95, 0.8, 0.4, [str(num)], 12, muted, align=PP_ALIGN.RIGHT)

    def foot(sl, unit=""):
        note = NOTE.get(lang, NOTE["en"])
        text(sl, 0.8, 6.95, 9, 0.4, [(unit + " · " if unit else "") + note], 11, muted)

    for i, s in enumerate(slides):
        sl = new_slide()
        lay = s["layout"]

        if lay == "title":
            box(sl, MSO_SHAPE.RECTANGLE, 1, 2.0, 1.6, 0.18, ac)
            n = len(s["title"])
            size = 44 if n <= 45 else (38 if n <= 70 else 32)
            text(sl, 1, 2.4, 11.3, 2.2, [s["title"]], size, tc, bold=True)
            if s.get("subtitle"):
                text(sl, 1, 4.8, 11.3, 0.9, [s["subtitle"]], 22, xc)
            text(sl, 1, 6.2, 11, 0.6, [author], 18, muted)

        elif lay == "chart":
            head(sl, s["title"], i + 1)
            ch = s["chart"]
            cd = CategoryChartData()
            cd.categories = ch["labels"]
            cd.add_series(s.get("unit") or "-", ch["values"])
            ctype = {"bar": XL_CHART_TYPE.COLUMN_CLUSTERED, "pie": XL_CHART_TYPE.PIE,
                     "line": XL_CHART_TYPE.LINE_MARKERS}[ch["type"]]
            gf = sl.shapes.add_chart(ctype, Inches(0.8), Inches(2.2),
                                     Inches(7.6), Inches(4.6), cd)
            chart = gf.chart
            chart.has_title = False
            chart.font.size = Pt(14)
            chart.font.color.rgb = rgb(xc)
            plot = chart.plots[0]
            plot.has_data_labels = True
            dl = plot.data_labels
            dl.font.size = Pt(13)
            dl.font.color.rgb = rgb(tc)
            ser = plot.series[0]
            try:
                if ch["type"] == "pie":
                    dl.show_percentage = True
                    dl.show_value = False
                    dl.number_format = "0%"
                    dl.number_format_is_linked = False
                    dl.position = XL_LABEL_POSITION.OUTSIDE_END
                    chart.has_legend = True
                    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
                    chart.legend.include_in_layout = False
                    chart.legend.font.color.rgb = rgb(xc)
                    for k in range(len(ch["values"])):
                        pt = ser.points[k]
                        pt.format.fill.solid()
                        pt.format.fill.fore_color.rgb = rgb(pal[k % len(pal)])
                        pt.format.line.color.rgb = rgb(bg)
                else:
                    chart.has_legend = False
                    va = chart.value_axis
                    va.has_major_gridlines = True
                    va.major_gridlines.format.line.color.rgb = rgb(blend(bg, tc, 0.15))
                    va.format.line.fill.background()
                    chart.category_axis.format.line.color.rgb = rgb(blend(bg, tc, 0.4))
                    if ch["type"] == "bar":
                        plot.gap_width = 60
                        plot.vary_by_categories = False
                        dl.position = XL_LABEL_POSITION.OUTSIDE_END
                        ser.format.fill.solid()
                        ser.format.fill.fore_color.rgb = rgb(ac)
                    else:
                        dl.position = XL_LABEL_POSITION.ABOVE
                        ser.smooth = False
                        ser.format.line.color.rgb = rgb(ac)
                        ser.format.line.width = Pt(3)
                        ser.marker.size = 9
                        ser.marker.format.fill.solid()
                        ser.marker.format.fill.fore_color.rgb = rgb(ac)
            except Exception:
                pass
            for k, p in enumerate(s["points"][:3]):
                y = 2.3 + k * 1.5
                box(sl, MSO_SHAPE.ROUNDED_RECTANGLE, 8.8, y, 3.8, 1.3, card, 0.08)
                box(sl, MSO_SHAPE.RECTANGLE, 8.8, y + 0.15, 0.08, 1.0, ac)
                text(sl, 9.05, y, 3.45, 1.3, [p], 16, tc, anchor=MSO_ANCHOR.MIDDLE)
            foot(sl, s.get("unit", ""))

        elif lay == "steps":
            head(sl, s["title"], i + 1)
            steps = s["steps"]
            n = len(steps)
            gap, total = 0.3, 11.7
            cw = (total - gap * (n - 1)) / n
            box(sl, MSO_SHAPE.RECTANGLE, 0.8 + cw / 2, 2.77, total - cw, 0.06,
                blend(bg, tc, 0.2))
            for k, st in enumerate(steps):
                x = 0.8 + k * (cw + gap)
                box(sl, MSO_SHAPE.ROUNDED_RECTANGLE, x, 3.4, cw, 3.2, card, 0.06)
                c = box(sl, MSO_SHAPE.OVAL, x + cw / 2 - 0.4, 2.4, 0.8, 0.8, ac)
                tf = c.text_frame
                tf.margin_left = tf.margin_right = 0
                tf.vertical_anchor = MSO_ANCHOR.MIDDLE
                p = tf.paragraphs[0]
                p.alignment = PP_ALIGN.CENTER
                r = p.add_run()
                r.text = str(k + 1)
                r.font.size = Pt(22)
                r.font.bold = True
                r.font.color.rgb = rgb(bg)
                text(sl, x + 0.15, 3.65, cw - 0.3, 1.0, [st["head"]],
                     20 if n <= 4 else 17, tc, bold=True, align=PP_ALIGN.CENTER)
                text(sl, x + 0.15, 4.7, cw - 0.3, 1.8, [st["text"]],
                     16 if n <= 4 else 14, xc, align=PP_ALIGN.CENTER)

        elif lay == "stats":
            head(sl, s["title"], i + 1)
            stats = s["stats"]
            n = len(stats)
            gap, total = 0.3, 11.7
            cw = (total - gap * (n - 1)) / n
            for k, st in enumerate(stats):
                x = 0.8 + k * (cw + gap)
                box(sl, MSO_SHAPE.ROUNDED_RECTANGLE, x, 2.5, cw, 3.0, card, 0.06)
                box(sl, MSO_SHAPE.RECTANGLE, x + 0.3, 2.5, cw - 0.6, 0.08, ac)
                size = 44 if n <= 3 else 38
                if len(st["value"]) > 8:
                    size -= 10
                text(sl, x + 0.1, 2.8, cw - 0.2, 1.3, [st["value"]], size, ac,
                     bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
                text(sl, x + 0.2, 4.2, cw - 0.4, 1.2, [st["label"]], 18, xc,
                     align=PP_ALIGN.CENTER)
            if s["points"]:
                text(sl, 0.8, 5.85, 11.7, 0.9, [s["points"][0]], 20, tc)
            foot(sl)

        else:  # bullets
            head(sl, s["title"], i + 1)
            img = s.get("_img")
            pts = ["•  " + p for p in s["points"]]
            total_chars = sum(len(p) for p in pts)
            if img:
                size = 20 if total_chars < 230 else 17
                width = 6.6
                box(sl, MSO_SHAPE.RECTANGLE, 7.9 + 0.15, 2.2 + 0.15, IMG_W, IMG_H, ac)
                sl.shapes.add_picture(io.BytesIO(img), Inches(7.9), Inches(2.2),
                                      Inches(IMG_W), Inches(IMG_H))
            else:
                size = 22 if total_chars < 380 else 19
                width = 11.7
            text(sl, 0.8, 2.4, width, 4.4, pts, size, xc, gap=12)

    prs.save(path)


# ============ ADMIN: TASDIQLASH ============
def set_caption(q, status_text, kb=None):
    base = (q.message.caption or "").split("\n\n")[0]
    return q.edit_message_caption(caption=(base + "\n\n" + status_text)[:1000],
                                  reply_markup=kb)


async def decision(u: Update, c: ContextTypes.DEFAULT_TYPE):
    q = u.callback_query
    if u.effective_user.id != ADMIN_ID:
        await q.answer("Ruxsat yo'q", show_alert=True)
        return
    act, oid = q.data.split(":")
    orders = load()
    o = orders.get(oid)
    if not o:
        await q.answer()
        await set_caption(q, "Buyurtma topilmadi.")
        return
    if o.get("status") in ("done", "rejected", "working"):
        await q.answer("Bu buyurtma allaqachon ko'rib chiqilgan yoki ishlanmoqda",
                       show_alert=True)
        return
    await q.answer()

    if act == "no":
        o["status"] = "rejected"
        o["done_at"] = now_str()
        save(orders)
        await c.bot.send_message(o["chat_id"], "To'lov tasdiqlanmadi ❌. Admin bilan bog'laning.")
        await set_caption(q, "❌ Rad etildi")
        return

    o["status"] = "working"
    save(orders)
    await set_caption(q, "⏳ Slayd yaratilmoqda...")
    path = f"{oid}.pptx"
    try:
        slides = await ai_slides(o["topic"], o["pages"], o.get("lang", "uz"))
        await attach_images(slides)
        await asyncio.to_thread(make_pptx, slides, path, o["name"],
                                o.get("theme", "1"), o.get("lang", "uz"))
        with open(path, "rb") as f:
            await c.bot.send_document(o["chat_id"], f,
                filename=f"{o['topic'][:40]}.pptx", caption="Tayyor! ✅")
        o["status"] = "done"
        o["done_at"] = now_str()
        save(orders)
        await set_caption(q, "✅ Tasdiqlandi va yuborildi")
    except Exception as e:
        o["status"] = "error"
        o["error"] = str(e)[:300]
        save(orders)
        await set_caption(q, f"⚠️ Xato: {e}"[:600], admin_kb(oid, retry=True))
    finally:
        if os.path.exists(path):
            os.remove(path)


# ============ ADMIN: STATISTIKA ============
async def stats(u: Update, c: ContextTypes.DEFAULT_TYPE):
    if u.effective_user.id != ADMIN_ID:
        return
    orders = list(load().values())
    n = datetime.datetime.now()
    today = n.date()

    def period(pred):
        done = [o for o in orders if o.get("status") == "done"
                and parse_ts(o.get("done_at") or o.get("created", "")) is not None
                and pred(parse_ts(o.get("done_at") or o.get("created")))]
        return len(done), sum(o.get("summa", 0) for o in done)

    rows = [
        ("Bugun", period(lambda t: t.date() == today)),
        ("Oxirgi 7 kun", period(lambda t: (n - t).days < 7)),
        ("Shu oy", period(lambda t: t.year == n.year and t.month == n.month)),
        ("Hammasi", period(lambda t: True)),
    ]
    cnt = {}
    for o in orders:
        cnt[o.get("status", "pending")] = cnt.get(o.get("status", "pending"), 0) + 1
    langs, themes = {}, {}
    for o in orders:
        if o.get("status") == "done":
            langs[o.get("lang", "?")] = langs.get(o.get("lang", "?"), 0) + 1
            th = THEMES.get(o.get("theme"), ("?",))[0]
            themes[th] = themes.get(th, 0) + 1
    lines = ["📊 Statistika (tasdiqlangan buyurtmalar)", ""]
    for name, (k, s) in rows:
        lines.append(f"{name}: {k} ta, {money(s)} so'm")
    lines += ["", "Holatlar: " + ", ".join(
        f"{STATUS_ICON.get(k, '•')} {v}" for k, v in cnt.items()) if cnt else "Buyurtma yo'q"]
    if langs:
        lines.append("Tillar: " + ", ".join(f"{k.upper()} {v}" for k, v in
                                            sorted(langs.items(), key=lambda x: -x[1])))
    if themes:
        lines.append("Shablonlar: " + ", ".join(f"{k} {v}" for k, v in
                                               sorted(themes.items(), key=lambda x: -x[1])[:3]))
    lines += ["", "So'nggi buyurtmalar: /orders"]
    await u.message.reply_text("\n".join(lines))


async def orders_cmd(u: Update, c: ContextTypes.DEFAULT_TYPE):
    if u.effective_user.id != ADMIN_ID:
        return
    items = sorted(load().items(), key=lambda kv: kv[1].get("created", ""), reverse=True)[:10]
    if not items:
        await u.message.reply_text("Hali buyurtma yo'q.")
        return
    lines = ["🧾 So'nggi 10 ta buyurtma", ""]
    for oid, o in items:
        who = f"@{o['username']}" if o.get("username") else o.get("name", "?")
        lines.append(
            f"{STATUS_ICON.get(o.get('status', 'pending'), '•')} #{oid} "
            f"{o.get('created', '')[5:16]}\n"
            f"   {o.get('name', '?')} ({who})\n"
            f"   {o.get('topic', '')[:35]} | {o.get('pages', '?')} bet | "
            f"{money(o.get('summa', 0))} so'm")
    await u.message.reply_text("\n".join(lines))


def main():
    app = Application.builder().token(BOT_TOKEN).build()
    conv = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_name)],
            TOPIC: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_topic)],
            PAGES: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_pages)],
            LANG: [CallbackQueryHandler(pick_lang, pattern="^lang:")],
            THEME: [CallbackQueryHandler(pick_theme, pattern="^theme:")],
            CHECK: [MessageHandler(filters.ALL & ~filters.COMMAND, get_check)],
        },
        fallbacks=[CommandHandler("cancel", cancel), CommandHandler("start", start)])
    app.add_handler(conv)
    app.add_handler(CommandHandler("stats", stats))
    app.add_handler(CommandHandler("orders", orders_cmd))
    app.add_handler(CallbackQueryHandler(decision, pattern="^(ok|no):"))
    if not PEXELS_KEY:
        print("Eslatma: rasm kaliti yo'q, slaydlarga rasm qo'shilmaydi (grafik va infografika ishlaydi).")
    print("Bot ishga tushdi.")
    app.run_polling()


if __name__ == "__main__":
    main()
