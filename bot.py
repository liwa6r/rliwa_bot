"""
بوت تيليغرام لتقارير التفتيش - النسخة المحسّنة
"""

import os
import re
import json
import base64
import logging
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from datetime import datetime
from io import BytesIO

import anthropic
import requests
from telegram import Update, ReplyKeyboardMarkup, KeyboardButton, ReplyKeyboardRemove
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    filters, ContextTypes, ConversationHandler
)
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

logging.basicConfig(format="%(asctime)s - %(levelname)s - %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)

# ── المتغيرات البيئية
TELEGRAM_TOKEN   = os.environ.get("TELEGRAM_TOKEN", "")
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
ADMIN_EMAIL      = "afra.6r@gmail.com"
GMAIL_USER       = os.environ.get("GMAIL_USER", "")
GMAIL_APP_PASS   = os.environ.get("GMAIL_APP_PASS", "")

# ── ألوان النموذج
BROWN    = RGBColor(0x8B, 0x25, 0x00)
LIGHT_BG = RGBColor(0xF2, 0xF0, 0xED)
WHITE    = RGBColor(0xFF, 0xFF, 0xFF)
BLACK    = RGBColor(0x1A, 0x1A, 0x1A)

# ── مراحل المحادثة
(WAITING_PHOTOS, WAITING_NAME, CONFIRM_NAME,
 WAITING_LOCATION, WAITING_NOTES, CONFIRM_SEND) = range(6)

user_sessions = {}


# ══════════════════════════════════════════
#  استخراج الموقع من رابط جوجل ماب
# ══════════════════════════════════════════

def extract_location_from_url(url: str) -> str:
    """يستخدم Claude لاستخراج اسم الموقع المختصر من رابط جوجل ماب"""
    try:
        client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
        response = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=100,
            messages=[{
                "role": "user",
                "content": (
                    f"هذا رابط جوجل ماب: {url}\n"
                    "استخرج اسم المنطقة والشارع فقط باللغة العربية، "
                    "بدون ذكر الدولة أو المدينة الكبيرة. "
                    "مثال: 'منطقة المرور، شارع الكورنيش'. "
                    "أجب بجملة واحدة قصيرة فقط."
                )
            }]
        )
        return response.content[0].text.strip()
    except Exception as e:
        logger.error(f"Location extraction error: {e}")
        return url


# ══════════════════════════════════════════
#  بداية المحادثة
# ══════════════════════════════════════════

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    user_sessions[user_id] = {
        "photos": [],
        "name": update.effective_user.first_name or "",
        "location": "",
        "notes": "",
        "date": datetime.now().strftime("%Y/%m/%d")
    }
    await update.message.reply_text(
        "👷 أهلاً! بوت تقارير التفتيش\n\nابدأ بإرسال الصور 📸",
        reply_markup=ReplyKeyboardRemove()
    )
    return WAITING_PHOTOS


# ══════════════════════════════════════════
#  استقبال الصور
# ══════════════════════════════════════════

async def receive_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if user_id not in user_sessions:
        await start(update, context)
        return WAITING_PHOTOS

    session = user_sessions[user_id]
    photo = update.message.photo[-1]
    file = await context.bot.get_file(photo.file_id)
    photo_bytes = await file.download_as_bytearray()
    session["photos"].append(bytes(photo_bytes))

    count = len(session["photos"])
    keyboard = [[KeyboardButton("✅ انتهيت من الصور")]]
    await update.message.reply_text(
        f"📸 {count} {'صورة' if count == 1 else 'صور'} — أرسل المزيد أو اضغط ✅",
        reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    )
    return WAITING_PHOTOS


async def done_photos(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    session = user_sessions.get(user_id, {})

    if not session.get("photos"):
        await update.message.reply_text("❌ أرسل صورة واحدة على الأقل.")
        return WAITING_PHOTOS

    # سؤال الاسم مع الاسم الافتراضي
    name = session.get("name", "")
    keyboard = [[KeyboardButton(f"✅ {name}")]] if name else []
    await update.message.reply_text(
        f"👤 ما اسمك؟\n(أو اضغط للتأكيد: {name})" if name else "👤 ما اسمك؟",
        reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True) if keyboard else ReplyKeyboardRemove()
    )
    return WAITING_NAME


# ══════════════════════════════════════════
#  الاسم
# ══════════════════════════════════════════

async def receive_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    text = update.message.text.strip()

    # إذا ضغط على زر التأكيد السريع
    if text.startswith("✅ "):
        text = text[2:].strip()

    user_sessions[user_id]["name"] = text

    keyboard = [[KeyboardButton("✅ نعم"), KeyboardButton("❌ تعديل")]]
    await update.message.reply_text(
        f"اسمك: *{text}*؟",
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    )
    return CONFIRM_NAME


async def confirm_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    if "تعديل" in text or "❌" in text:
        await update.message.reply_text("👤 أعد إدخال اسمك:", reply_markup=ReplyKeyboardRemove())
        return WAITING_NAME

    keyboard = [
        [KeyboardButton("📍 مشاركة الموقع", request_location=True)],
        [KeyboardButton("⏭️ تخطي")]
    ]
    await update.message.reply_text(
        "📍 شارك موقعك أو أرسل رابط جوجل ماب\n(أو اضغط تخطي)",
        reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    )
    return WAITING_LOCATION


# ══════════════════════════════════════════
#  الموقع
# ══════════════════════════════════════════

async def receive_location(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """استقبال موقع GPS من تيليغرام"""
    user_id = update.effective_user.id
    loc = update.message.location
    lat, lng = loc.latitude, loc.longitude

    # تحويل الإحداثيات لاسم مختصر
    location_text = f"https://maps.google.com/?q={lat},{lng}"
    if ANTHROPIC_API_KEY:
        location_text = extract_location_from_url(location_text)
    else:
        location_text = f"{lat:.4f}, {lng:.4f}"

    user_sessions[user_id]["location"] = location_text
    user_sessions[user_id]["maps_link"] = f"https://maps.google.com/?q={lat},{lng}"

    keyboard = [[KeyboardButton("⏭️ بدون ملاحظات")]]
    await update.message.reply_text(
        f"✅ الموقع: {location_text}\n\n📝 ملاحظات إضافية؟",
        reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    )
    return WAITING_NOTES


async def receive_location_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """استقبال رابط أو نص الموقع"""
    user_id = update.effective_user.id
    text = update.message.text.strip()

    if "تخطي" in text or "⏭️" in text:
        user_sessions[user_id]["location"] = ""
    elif "maps" in text.lower() or "goo.gl" in text.lower() or "http" in text.lower():
        # رابط جوجل ماب
        if ANTHROPIC_API_KEY:
            await update.message.reply_text("⏳ جاري تحديد الموقع...")
            location_name = extract_location_from_url(text)
        else:
            location_name = text
        user_sessions[user_id]["location"] = location_name
        user_sessions[user_id]["maps_link"] = text
    else:
        user_sessions[user_id]["location"] = text

    keyboard = [[KeyboardButton("⏭️ بدون ملاحظات")]]
    await update.message.reply_text(
        "📝 ملاحظات إضافية؟",
        reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    )
    return WAITING_NOTES


# ══════════════════════════════════════════
#  الملاحظات والتأكيد
# ══════════════════════════════════════════

async def receive_notes(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    text = update.message.text.strip()
    session = user_sessions[user_id]
    session["notes"] = "" if "بدون ملاحظات" in text else text

    count = len(session["photos"])
    layout = "صورتان في الصفحة" if count >= 2 else "صورة في الصفحة"

    summary = (
        f"📋 *ملخص:*\n\n"
        f"👤 {session['name']}\n"
        f"📍 {session.get('location') or '—'}\n"
        f"📸 {count} صورة | {layout}\n"
        f"📅 {session['date']}\n"
    )
    if session["notes"]:
        summary += f"📝 {session['notes']}\n"

    keyboard = [[KeyboardButton("✅ إرسال"), KeyboardButton("❌ إلغاء")]]
    await update.message.reply_text(
        summary,
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    )
    return CONFIRM_SEND


async def confirm_send(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    text = update.message.text

    if "إلغاء" in text:
        await update.message.reply_text("❌ تم الإلغاء. اكتب /start للبدء.", reply_markup=ReplyKeyboardRemove())
        return ConversationHandler.END

    await update.message.reply_text("⏳ جاري إنشاء التقرير...", reply_markup=ReplyKeyboardRemove())
    session = user_sessions[user_id]
    await generate_and_send(update, context, session)
    user_sessions.pop(user_id, None)
    return ConversationHandler.END


# ══════════════════════════════════════════
#  إنشاء ملف PowerPoint
# ══════════════════════════════════════════

def create_pptx_report(session: dict) -> BytesIO:
    prs = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)
    W = prs.slide_width
    H = prs.slide_height

    photos   = session["photos"]
    name     = session.get("name", "")
    location = session.get("location", "") or "—"
    notes    = session.get("notes", "")
    date     = session.get("date", "")

    # ── دوال مساعدة ──────────────────────────────────────────────

    def add_bg(slide):
        bg = slide.shapes.add_shape(1, 0, 0, W, H)
        bg.fill.solid()
        bg.fill.fore_color.rgb = RGBColor(0xF5, 0xF3, 0xF0)
        bg.line.fill.background()

    def add_corners(slide):
        sz = Inches(0.45)
        for (x, y) in [(0, 0), (W-sz, 0), (0, H-sz), (W-sz, H-sz)]:
            s = slide.shapes.add_shape(1, x, y, sz, sz)
            s.fill.solid()
            s.fill.fore_color.rgb = BROWN
            s.line.fill.background()

    def add_text(slide, text, x, y, w, h, size=10, bold=False,
                 color=BLACK, align=PP_ALIGN.RIGHT, italic=False):
        tb = slide.shapes.add_textbox(x, y, w, h)
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = align
        run = p.add_run()
        run.text = str(text)
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.italic = italic
        run.font.color.rgb = color
        run.font.name = "Arial"

    def add_photo_card(slide, photo_bytes, x, y, w, h, loc_text):
        """بطاقة صورة احترافية مع بيانات"""
        # إطار خارجي
        frame = slide.shapes.add_shape(1, x, y, w, h)
        frame.fill.solid()
        frame.fill.fore_color.rgb = WHITE
        frame.line.color.rgb = BROWN
        frame.line.width = Pt(1.5)

        padding    = Inches(0.12)
        info_h     = Inches(0.95)
        img_h      = h - info_h - padding * 3

        # الصورة
        img_stream = BytesIO(photo_bytes)
        try:
            slide.shapes.add_picture(
                img_stream,
                x + padding, y + padding,
                w - padding * 2, img_h
            )
        except Exception as e:
            logger.warning(f"Photo error: {e}")

        # خلفية المعلومات
        info_y = y + padding + img_h + padding * 0.5
        info_bg = slide.shapes.add_shape(
            1,
            x + padding, info_y,
            w - padding * 2, info_h
        )
        info_bg.fill.solid()
        info_bg.fill.fore_color.rgb = RGBColor(0xF8, 0xF5, 0xF2)
        info_bg.line.color.rgb = RGBColor(0xD0, 0xC0, 0xB0)
        info_bg.line.width = Pt(0.5)

        # خط بني في الأعلى
        accent = slide.shapes.add_shape(
            1,
            x + padding, info_y,
            w - padding * 2, Inches(0.05)
        )
        accent.fill.solid()
        accent.fill.fore_color.rgb = BROWN
        accent.line.fill.background()

        lh   = Inches(0.28)
        lpad = Inches(0.18)
        rows = [
            ("👤", name),
            ("📍", loc_text),
        ]
        for i, (icon, val) in enumerate(rows):
            ly = info_y + Inches(0.1) + i * lh
            add_text(slide, f"{icon} {val}",
                     x + lpad, ly,
                     w - lpad * 2, lh,
                     size=9, color=BLACK, align=PP_ALIGN.RIGHT)

    # ── صفحة الغلاف ───────────────────────────────────────────────

    def make_cover():
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        add_bg(slide)
        add_corners(slide)

        # شريط العنوان
        bar = slide.shapes.add_shape(1, Inches(1), Inches(2.5), W - Inches(2), Inches(1.3))
        bar.fill.solid()
        bar.fill.fore_color.rgb = BROWN
        bar.line.fill.background()

        add_text(slide, "تقرير تفتيش موقع",
                 Inches(1), Inches(2.6), W - Inches(2), Inches(0.7),
                 size=28, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

        add_text(slide, "Site Inspection Report",
                 Inches(1), Inches(3.2), W - Inches(2), Inches(0.4),
                 size=14, color=RGBColor(0xCC, 0xAA, 0x88),
                 align=PP_ALIGN.CENTER, italic=True)

        # البيانات
        info_items = [
            ("👤 المفتش:", name),
            ("📍 الموقع:", location),
            ("📅 التاريخ:", date),
            ("📸 عدد الصور:", str(len(photos))),
        ]
        for i, (lbl, val) in enumerate(info_items):
            iy = Inches(4.1) + i * Inches(0.42)
            add_text(slide, lbl, Inches(7), iy, Inches(2), Inches(0.38),
                     size=11, bold=True, color=BROWN, align=PP_ALIGN.RIGHT)
            add_text(slide, val, Inches(3), iy, Inches(4.2), Inches(0.38),
                     size=11, color=BLACK, align=PP_ALIGN.RIGHT)

        if notes:
            add_text(slide, f"📝 {notes}",
                     Inches(1), H - Inches(0.7), W - Inches(2), Inches(0.4),
                     size=10, color=BROWN, align=PP_ALIGN.CENTER, italic=True)

        add_text(slide, date,
                 Inches(0.6), H - Inches(0.45), Inches(3), Inches(0.35),
                 size=10, color=BROWN, align=PP_ALIGN.LEFT)

    # ── صفحة صورة واحدة ───────────────────────────────────────────

    def make_single_slide(photo_bytes, photo_num, total):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        add_bg(slide)
        add_corners(slide)

        add_text(slide, f"صورة {photo_num} من {total}",
                 Inches(0.6), Inches(0.2), Inches(4), Inches(0.4),
                 size=11, color=BROWN, align=PP_ALIGN.RIGHT, bold=True)
        add_text(slide, date,
                 W - Inches(4.6), Inches(0.2), Inches(4), Inches(0.4),
                 size=10, color=BROWN, align=PP_ALIGN.LEFT)

        mx, my = Inches(1.0), Inches(0.7)
        bw = W - mx * 2
        bh = H - my - Inches(0.4)
        add_photo_card(slide, photo_bytes, mx, my, bw, bh, location)

    # ── صفحة صورتين ───────────────────────────────────────────────

    def make_double_slide(p1, p2, page_num):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        add_bg(slide)
        add_corners(slide)

        add_text(slide, f"صفحة {page_num}",
                 Inches(0.6), Inches(0.15), Inches(4), Inches(0.35),
                 size=10, color=BROWN, align=PP_ALIGN.RIGHT, bold=True)
        add_text(slide, date,
                 W - Inches(4.6), Inches(0.15), Inches(4), Inches(0.35),
                 size=10, color=BROWN, align=PP_ALIGN.LEFT)

        mx   = Inches(0.6)
        my   = Inches(0.6)
        gap  = Inches(0.25)
        bw   = (W - mx * 2 - gap) / 2
        bh   = H - my - Inches(0.35)

        add_photo_card(slide, p1, mx, my, bw, bh, location)
        add_photo_card(slide, p2, mx + bw + gap, my, bw, bh, location)

    # ── صفحة الختام ───────────────────────────────────────────────

    def make_closing():
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        add_bg(slide)
        add_corners(slide)

        # إطار وسط
        fw, fh = Inches(5), Inches(2.5)
        fx = (W - fw) / 2
        fy = (H - fh) / 2
        frame = slide.shapes.add_shape(1, fx, fy, fw, fh)
        frame.fill.solid()
        frame.fill.fore_color.rgb = WHITE
        frame.line.color.rgb = BROWN
        frame.line.width = Pt(2)

        add_text(slide, "شكـــراً",
                 fx, fy + Inches(0.5), fw, Inches(1.2),
                 size=40, bold=True, color=BROWN, align=PP_ALIGN.CENTER)

        add_text(slide, "Thank You",
                 fx, fy + Inches(1.6), fw, Inches(0.6),
                 size=14, color=RGBColor(0xAA, 0x66, 0x44),
                 align=PP_ALIGN.CENTER, italic=True)

        add_text(slide, date,
                 Inches(0.6), H - Inches(0.45), Inches(3), Inches(0.35),
                 size=10, color=BROWN, align=PP_ALIGN.LEFT)

    # ── بناء الصفحات ─────────────────────────────────────────────

    make_cover()

    if len(photos) == 1:
        make_single_slide(photos[0], 1, 1)
    else:
        page = 1
        for i in range(0, len(photos), 2):
            if i + 1 < len(photos):
                make_double_slide(photos[i], photos[i+1], page)
            else:
                make_single_slide(photos[i], i+1, len(photos))
            page += 1

    make_closing()

    buf = BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf


# ══════════════════════════════════════════
#  إرسال الإيميل
# ══════════════════════════════════════════

def send_email(buf: BytesIO, session: dict) -> bool:
    if not GMAIL_USER or not GMAIL_APP_PASS:
        logger.warning("Gmail not configured")
        return False
    try:
        msg = MIMEMultipart()
        msg["From"]    = GMAIL_USER
        msg["To"]      = ADMIN_EMAIL
        msg["Subject"] = f"تقرير تفتيش — {session.get('name','')} — {session.get('date','')}"

        body = (
            f"تقرير تفتيش جديد\n\n"
            f"المفتش: {session.get('name','—')}\n"
            f"الموقع: {session.get('location') or '—'}\n"
            f"التاريخ: {session.get('date','—')}\n"
            f"عدد الصور: {len(session.get('photos',[]))}\n"
        )
        if session.get("maps_link"):
            body += f"رابط الموقع: {session['maps_link']}\n"
        if session.get("notes"):
            body += f"الملاحظات: {session['notes']}\n"

        msg.attach(MIMEText(body, "plain", "utf-8"))

        fname = f"تقرير_{session.get('name','')}_{session.get('date','')}.pptx"
        part = MIMEBase("application", "octet-stream")
        part.set_payload(buf.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f'attachment; filename="{fname}"')
        msg.attach(part)

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
            s.login(GMAIL_USER, GMAIL_APP_PASS)
            s.send_message(msg)
        return True
    except Exception as e:
        logger.error(f"Email error: {e}")
        return False


# ══════════════════════════════════════════
#  إنشاء وإرسال التقرير
# ══════════════════════════════════════════

async def generate_and_send(update, context, session):
    try:
        buf = create_pptx_report(session)

        fname = f"تقرير_{session.get('name','')}_{session.get('date','')}.pptx"

        # إرسال للموظف
        buf.seek(0)
        await update.message.reply_document(
            document=buf,
            filename=fname,
            caption=(
                f"✅ التقرير جاهز\n"
                f"👤 {session.get('name','')}\n"
                f"📸 {len(session.get('photos',[]))} صورة\n"
                f"📅 {session.get('date','')}"
            )
        )

        # إرسال بالإيميل
        buf.seek(0)
        ok = send_email(buf, session)
        if ok:
            await update.message.reply_text("📧 تم الإرسال إلى البريد ✅\n\nاكتب /start لتقرير جديد")
        else:
            await update.message.reply_text("⚠️ الملف جاهز هنا — راجع إعداد الإيميل\n\nاكتب /start لتقرير جديد")

    except Exception as e:
        logger.error(f"Error: {e}")
        await update.message.reply_text(f"❌ خطأ: {str(e)}")


# ══════════════════════════════════════════
#  تشغيل البوت
# ══════════════════════════════════════════

def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()

    conv = ConversationHandler(
        entry_points=[
            CommandHandler("start", start),
            MessageHandler(filters.PHOTO, receive_photo),
        ],
        states={
            WAITING_PHOTOS: [
                MessageHandler(filters.PHOTO, receive_photo),
                MessageHandler(filters.Regex("^✅ انتهيت من الصور$"), done_photos),
            ],
            WAITING_NAME: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, receive_name),
            ],
            CONFIRM_NAME: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, confirm_name),
            ],
            WAITING_LOCATION: [
                MessageHandler(filters.LOCATION, receive_location),
                MessageHandler(filters.TEXT & ~filters.COMMAND, receive_location_text),
            ],
            WAITING_NOTES: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, receive_notes),
            ],
            CONFIRM_SEND: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, confirm_send),
            ],
        },
        fallbacks=[CommandHandler("start", start)],
        allow_reentry=True,
    )

    app.add_handler(conv)
    print("🤖 البوت يعمل...")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
