"""
بوت تيليغرام لتقارير التفتيش الإنشائي
النسخة النهائية — مع تصميم النموذج الرسمي
"""

import os
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

from telegram import Update, ReplyKeyboardMarkup, KeyboardButton, ReplyKeyboardRemove
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    filters, ContextTypes, ConversationHandler
)
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ── إعداد الـ Logging
logging.basicConfig(format="%(asctime)s - %(levelname)s - %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)

# ── المتغيرات البيئية
TELEGRAM_TOKEN    = os.environ.get("TELEGRAM_TOKEN", "")
ADMIN_EMAIL       = "afra.6r@gmail.com"
GMAIL_USER        = os.environ.get("GMAIL_USER", "")
GMAIL_APP_PASS    = os.environ.get("GMAIL_APP_PASS", "")

# ── ألوان النموذج
BROWN    = RGBColor(0x8B, 0x25, 0x00)
LIGHT_BG = RGBColor(0xF2, 0xF0, 0xED)
BLACK    = RGBColor(0x00, 0x00, 0x00)

# ── مراحل المحادثة
(WAITING_PHOTOS, WAITING_LAYOUT, WAITING_NAME,
 CONFIRM_NAME, WAITING_PHONE, WAITING_ADDRESS,
 WAITING_NOTES, CONFIRM_SEND) = range(8)

user_sessions = {}


# ══════════════════════════════════════════
#  بداية المحادثة
# ══════════════════════════════════════════

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    user_sessions[user_id] = {
        "photos": [], "name": "", "phone": "",
        "address": "", "notes": "", "layout": "single",
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

    count = len(session["photos"])

    if count > 1:
        keyboard = [
            [KeyboardButton("🖼️ صورة واحدة في كل صفحة")],
            [KeyboardButton("🖼️🖼️ صورتان في نفس الصفحة")]
        ]
        await update.message.reply_text(
            f"استلمت {count} صور 📷\nكيف تريد ترتيبها؟",
            reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
        )
        return WAITING_LAYOUT
    else:
        session["layout"] = "single"
        await update.message.reply_text("👤 ما اسمك الكامل؟", reply_markup=ReplyKeyboardRemove())
        return WAITING_NAME


async def receive_layout(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    text = update.message.text
    user_sessions[user_id]["layout"] = "double" if "صورتان" in text else "single"
    await update.message.reply_text("👤 ما اسمك الكامل؟", reply_markup=ReplyKeyboardRemove())
    return WAITING_NAME


# ══════════════════════════════════════════
#  جمع البيانات
# ══════════════════════════════════════════

async def receive_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    name = update.message.text.strip()
    user_sessions[user_id]["name"] = name

    keyboard = [[KeyboardButton("✅ نعم، صحيح"), KeyboardButton("❌ لا، عدّل")]]
    await update.message.reply_text(
        f"هل اسمك هو:\n*{name}*؟",
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    )
    return CONFIRM_NAME


async def confirm_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    if "لا" in text or "عدّل" in text:
        await update.message.reply_text("👤 أعد إدخال اسمك:", reply_markup=ReplyKeyboardRemove())
        return WAITING_NAME

    await update.message.reply_text(
        "📱 رقم هاتفك؟\n(أو أرسل — للتخطي)",
        reply_markup=ReplyKeyboardRemove()
    )
    return WAITING_PHONE


async def receive_phone(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    text = update.message.text.strip()
    user_sessions[user_id]["phone"] = "" if text in ["-", "—"] else text
    await update.message.reply_text("📍 العنوان / الموقع؟\n(أو أرسل — للتخطي)")
    return WAITING_ADDRESS


async def receive_address(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    text = update.message.text.strip()
    user_sessions[user_id]["address"] = "" if text in ["-", "—"] else text

    keyboard = [[KeyboardButton("⏭️ بدون ملاحظات")]]
    await update.message.reply_text(
        "📝 ملاحظات إضافية؟",
        reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    )
    return WAITING_NOTES


async def receive_notes(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    text = update.message.text.strip()
    session = user_sessions[user_id]
    session["notes"] = "" if "بدون ملاحظات" in text else text

    layout_text = "صورتان في الصفحة" if session["layout"] == "double" else "صورة واحدة في الصفحة"
    summary = (
        f"📋 *ملخص التقرير:*\n\n"
        f"👤 الاسم: {session['name']}\n"
        f"📱 الهاتف: {session.get('phone') or '—'}\n"
        f"📍 العنوان: {session.get('address') or '—'}\n"
        f"📸 الصور: {len(session['photos'])} | {layout_text}\n"
        f"📅 التاريخ: {session['date']}\n"
    )
    if session["notes"]:
        summary += f"📝 الملاحظات: {session['notes']}\n"

    keyboard = [[KeyboardButton("✅ أرسل التقرير"), KeyboardButton("❌ إلغاء")]]
    await update.message.reply_text(
        summary + "\nهل تريد إرسال التقرير؟",
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    )
    return CONFIRM_SEND


async def confirm_send(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    text = update.message.text

    if "إلغاء" in text:
        await update.message.reply_text("❌ تم الإلغاء.\nاكتب /start لبدء من جديد", reply_markup=ReplyKeyboardRemove())
        return ConversationHandler.END

    await update.message.reply_text("⏳ جاري إنشاء التقرير...", reply_markup=ReplyKeyboardRemove())

    session = user_sessions[user_id]
    success = await generate_and_send(update, context, session)

    if success:
        await update.message.reply_text("✅ تم! اكتب /start لتقرير جديد")
    else:
        await update.message.reply_text("❌ حدث خطأ، حاول مجدداً")

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

    photos  = session["photos"]
    layout  = session["layout"]
    name    = session.get("name", "")
    phone   = session.get("phone", "") or "—"
    address = session.get("address", "") or "—"
    notes   = session.get("notes", "")
    date    = session.get("date", "")

    def add_corners(slide):
        sz = Inches(0.5)
        for (x, y) in [(0,0), (W-sz,0), (0,H-sz), (W-sz,H-sz)]:
            s = slide.shapes.add_shape(1, x, y, sz, sz)
            s.fill.solid()
            s.fill.fore_color.rgb = BROWN
            s.line.fill.background()

    def add_bg(slide):
        bg = slide.shapes.add_shape(1, 0, 0, W, H)
        bg.fill.solid()
        bg.fill.fore_color.rgb = RGBColor(0xF5, 0xF3, 0xF0)
        bg.line.fill.background()

    def add_label(slide, text, x, y, w, h, size=10, bold=False, color=BLACK, align=PP_ALIGN.RIGHT):
        tb = slide.shapes.add_textbox(x, y, w, h)
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = align
        run = p.add_run()
        run.text = text
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = color
        run.font.name = "Arial"

    def add_photo_block(slide, photo_bytes, x, y, w, h):
        info_h = Inches(1.3)
        img_h  = h - info_h - Inches(0.08)

        # الصورة
        img_stream = BytesIO(photo_bytes)
        try:
            slide.shapes.add_picture(img_stream, x, y, w, img_h)
        except Exception as e:
            logger.warning(f"Photo error: {e}")

        # خلفية البيانات
        iy = y + img_h + Inches(0.05)
        bg = slide.shapes.add_shape(1, x, iy, w, info_h)
        bg.fill.solid()
        bg.fill.fore_color.rgb = LIGHT_BG
        bg.line.color.rgb = BROWN
        bg.line.width = Pt(0.75)

        lh = Inches(0.38)
        rows = [
            ("الاسم:", name),
            ("الهاتف:", phone),
            ("العنوان:", address),
        ]
        for i, (lbl, val) in enumerate(rows):
            ly = iy + Inches(0.06) + i * lh
            add_label(slide, lbl, x + Inches(0.1), ly, Inches(1.1), lh,
                      bold=True, color=BROWN)
            add_label(slide, val, x + Inches(1.15), ly, w - Inches(1.3), lh,
                      color=BLACK)

    def make_slide_single(photo_bytes):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        add_bg(slide)
        add_corners(slide)
        add_label(slide, date, Inches(0.65), H - Inches(0.42),
                  Inches(2.5), Inches(0.35), size=11, color=BROWN, align=PP_ALIGN.LEFT)
        if notes:
            add_label(slide, f"ملاحظات: {notes}",
                      Inches(0.65), H - Inches(0.75),
                      W - Inches(1.3), Inches(0.35),
                      size=9, color=BROWN, align=PP_ALIGN.RIGHT)

        mx, my = Inches(1.3), Inches(0.55)
        bw = W - mx * 2
        bh = H - my * 2 - Inches(0.5)
        add_photo_block(slide, photo_bytes, mx, my, bw, bh)

    def make_slide_double(p1, p2):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        add_bg(slide)
        add_corners(slide)
        add_label(slide, date, Inches(0.65), H - Inches(0.42),
                  Inches(2.5), Inches(0.35), size=11, color=BROWN, align=PP_ALIGN.LEFT)

        mx, my = Inches(0.7), Inches(0.5)
        gap    = Inches(0.25)
        bw     = (W - mx * 2 - gap) / 2
        bh     = H - my * 2 - Inches(0.5)

        add_photo_block(slide, p1, mx, my, bw, bh)
        add_photo_block(slide, p2, mx + bw + gap, my, bw, bh)

    # إنشاء الصفحات
    if layout == "double":
        for i in range(0, len(photos), 2):
            if i + 1 < len(photos):
                make_slide_double(photos[i], photos[i+1])
            else:
                make_slide_single(photos[i])
    else:
        for p in photos:
            make_slide_single(p)

    buf = BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf


# ══════════════════════════════════════════
#  إرسال الإيميل
# ══════════════════════════════════════════

def send_email(pptx_buf: BytesIO, session: dict) -> bool:
    if not GMAIL_USER or not GMAIL_APP_PASS:
        logger.warning("Gmail credentials not set")
        return False
    try:
        msg = MIMEMultipart()
        msg["From"]    = GMAIL_USER
        msg["To"]      = ADMIN_EMAIL
        msg["Subject"] = f"تقرير تفتيش — {session.get('name','')} — {session.get('date','')}"

        body = (
            f"تقرير تفتيش جديد\n\n"
            f"الاسم: {session.get('name','—')}\n"
            f"الهاتف: {session.get('phone') or '—'}\n"
            f"العنوان: {session.get('address') or '—'}\n"
            f"التاريخ: {session.get('date','—')}\n"
            f"عدد الصور: {len(session.get('photos',[]))}\n"
        )
        if session.get("notes"):
            body += f"الملاحظات: {session['notes']}\n"

        msg.attach(MIMEText(body, "plain", "utf-8"))

        fname = f"تقرير_{session.get('name','')}_{session.get('date','')}.pptx"
        part = MIMEBase("application", "octet-stream")
        part.set_payload(pptx_buf.read())
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
#  إنشاء وإرسال
# ══════════════════════════════════════════

async def generate_and_send(update, context, session):
    try:
        buf = create_pptx_report(session)

        # إرسال للموظف عبر تيليغرام
        buf.seek(0)
        fname = f"تقرير_{session.get('name','')}_{session.get('date','')}.pptx"
        await update.message.reply_document(
            document=buf,
            filename=fname,
            caption=(
                f"✅ تقرير جاهز\n"
                f"👤 {session.get('name','')}\n"
                f"📸 {len(session.get('photos',[]))} صورة\n"
                f"📅 {session.get('date','')}"
            )
        )

        # إرسال بالإيميل
        buf.seek(0)
        ok = send_email(buf, session)
        if ok:
            await update.message.reply_text("📧 تم الإرسال إلى البريد الإلكتروني ✅")
        else:
            await update.message.reply_text("⚠️ الملف متاح هنا لكن الإيميل يحتاج إعداد")

        return True
    except Exception as e:
        logger.error(f"Generate error: {e}")
        return False


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
            WAITING_PHOTOS:  [MessageHandler(filters.PHOTO, receive_photo),
                              MessageHandler(filters.Regex("^✅ انتهيت من الصور$"), done_photos)],
            WAITING_LAYOUT:  [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_layout)],
            WAITING_NAME:    [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_name)],
            CONFIRM_NAME:    [MessageHandler(filters.TEXT & ~filters.COMMAND, confirm_name)],
            WAITING_PHONE:   [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_phone)],
            WAITING_ADDRESS: [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_address)],
            WAITING_NOTES:   [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_notes)],
            CONFIRM_SEND:    [MessageHandler(filters.TEXT & ~filters.COMMAND, confirm_send)],
        },
        fallbacks=[CommandHandler("start", start)],
        allow_reentry=True,
    )

    app.add_handler(conv)
    print("🤖 البوت يعمل...")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
