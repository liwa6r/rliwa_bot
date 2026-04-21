"""
بوت ليوا - نظام تقارير التفتيش الذكي
Liwa Inspection Bot - Smart Version
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
from telegram import Update, ReplyKeyboardMarkup, KeyboardButton, ReplyKeyboardRemove
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    filters, ContextTypes, ConversationHandler
)
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

logging.basicConfig(format="%(asctime)s - %(levelname)s - %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)

# ── المتغيرات البيئية
TELEGRAM_TOKEN    = os.environ.get("TELEGRAM_TOKEN", "")
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
ADMIN_EMAIL       = "afra.6r@gmail.com"
GMAIL_USER        = os.environ.get("GMAIL_USER", "")
GMAIL_APP_PASS    = os.environ.get("GMAIL_APP_PASS", "")

# ── نطاق منطقة ليوا الجغرافي
LIWA_BOUNDS = {
    "lat_min": 22.8, "lat_max": 23.5,
    "lng_min": 53.5, "lng_max": 54.5
}

# ── ألوان النموذج
BROWN    = RGBColor(0x8B, 0x25, 0x00)
LIGHT_BG = RGBColor(0xF2, 0xF0, 0xED)
WHITE    = RGBColor(0xFF, 0xFF, 0xFF)
BLACK    = RGBColor(0x1A, 0x1A, 0x1A)

# ── مراحل المحادثة
(MAIN_MENU, WAITING_PHOTO, WAITING_LOCATION,
 WAITING_NOTE, WAITING_BEFORE_PHOTO, WAITING_AFTER_PHOTO,
 WAITING_BEFORE_LOC, WAITING_AFTER_LOC,
 ADD_MORE, CONFIRM_SEND, CHAT_MODE) = range(11)

user_sessions = {}
report_stats = {"total": 0}  # إحصائيات التقارير


# ══════════════════════════════════════════════════
#  Claude AI - المساعد الذكي
# ══════════════════════════════════════════════════

def ai_chat(user_message: str, session: dict) -> str:
    """محادثة ذكية مع Claude للإجابة على الاستفسارات"""
    try:
        client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
        context = f"""أنت "بوت ليوا" — مساعد ذكي لإنشاء تقارير تفتيش المواقع في منطقة ليوا.

معلومات النظام:
- إجمالي التقارير المُنشأة: {report_stats['total']}
- الجلسة الحالية: {len(session.get('photos', []))} صورة مضافة
- نوع التقرير: {session.get('report_type', 'عادي')}

قواعد الإجابة:
- أجب بالعربية بشكل ودي وقصير
- إذا سأل عن عدد التقارير: أخبره بالرقم
- إذا سأل عن تقرير قبل/بعد: اشرح له كيف يبدأ
- إذا كتب خطأً إملائياً واضحاً: صحح له وأجب
- إذا طلب تغيير اسم أو بيانات: أخبره أنك ستضيفها
- إذا سأل عن الموقع: اشرح له طرق الإرسال
- لا تتجاوز 3 جمل في إجابتك"""

        response = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=200,
            system=context,
            messages=[{"role": "user", "content": user_message}]
        )
        return response.content[0].text.strip()
    except Exception as e:
        logger.error(f"AI chat error: {e}")
        return "عذراً، حدث خطأ. جرب مرة أخرى أو اكتب /start"


def ai_extract_location(lat: float, lng: float, photo_bytes: bytes = None) -> dict:
    """استخراج اسم الموقع والتحقق من نطاق ليوا"""
    try:
        # التحقق من نطاق ليوا
        in_liwa = (
            LIWA_BOUNDS["lat_min"] <= lat <= LIWA_BOUNDS["lat_max"] and
            LIWA_BOUNDS["lng_min"] <= lng <= LIWA_BOUNDS["lng_max"]
        )

        client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

        content = []
        if photo_bytes:
            b64 = base64.standard_b64encode(photo_bytes).decode()
            content.append({
                "type": "image",
                "source": {"type": "base64", "media_type": "image/jpeg", "data": b64}
            })

        content.append({
            "type": "text",
            "text": (
                f"الإحداثيات: {lat:.4f}, {lng:.4f}\n"
                f"هذا الموقع {'ضمن' if in_liwa else 'خارج'} نطاق منطقة ليوا.\n\n"
                "استخرج اسم المنطقة والموقع بالعربية فقط، مختصراً (بدون ذكر الإمارات أو أبوظبي).\n"
                "مثال: 'واحة ليوا، طريق الرمال' أو 'منطقة المرور'\n"
                "أجب بسطر واحد فقط."
            )
        })

        response = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=60,
            messages=[{"role": "user", "content": content}]
        )

        location_name = response.content[0].text.strip()
        return {"name": location_name, "in_liwa": in_liwa, "lat": lat, "lng": lng}

    except Exception as e:
        logger.error(f"Location extraction error: {e}")
        return {
            "name": f"{lat:.4f}, {lng:.4f}",
            "in_liwa": True,
            "lat": lat, "lng": lng
        }


def parse_coordinates(text: str) -> tuple:
    """تحليل الإحداثيات من نص"""
    patterns = [
        r'(-?\d+\.?\d*)[,،\s]+(-?\d+\.?\d*)',
        r'q=(-?\d+\.?\d*),(-?\d+\.?\d*)',
        r'@(-?\d+\.?\d*),(-?\d+\.?\d*)',
        r'll=(-?\d+\.?\d*),(-?\d+\.?\d*)',
    ]
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            lat, lng = float(match.group(1)), float(match.group(2))
            if -90 <= lat <= 90 and -180 <= lng <= 180:
                return lat, lng
    return None, None


# ══════════════════════════════════════════════════
#  بداية المحادثة
# ══════════════════════════════════════════════════

def init_session(user_id: int, name: str):
    user_sessions[user_id] = {
        "photos": [],           # [{photo, location, note, type}]
        "name": name,
        "report_type": "normal",
        "date": datetime.now().strftime("%Y/%m/%d"),
        "last_location": None,  # آخر موقع لإعادة استخدامه
        "pending_photo": None,  # الصورة المعلقة حتى يُضاف الموقع
        "sent": False,
    }


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    name = update.effective_user.first_name or "مستخدم"
    init_session(user_id, name)

    keyboard = [
        [KeyboardButton("📸 تقرير عادي")],
        [KeyboardButton("🔄 تقرير قبل وبعد")],
    ]
    await update.message.reply_text(
        f"👷 أهلاً {name}! بوت ليوا للتفتيش\n\nاختر نوع التقرير:",
        reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    )
    return MAIN_MENU


async def main_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    text = update.message.text

    if "عادي" in text or "📸" in text:
        user_sessions[user_id]["report_type"] = "normal"
        await update.message.reply_text(
            "📸 أرسل الصورة الأولى:",
            reply_markup=ReplyKeyboardRemove()
        )
        return WAITING_PHOTO

    elif "قبل وبعد" in text or "🔄" in text:
        user_sessions[user_id]["report_type"] = "before_after"
        await update.message.reply_text(
            "📸 أرسل صورة **قبل** التدخل:",
            parse_mode="Markdown",
            reply_markup=ReplyKeyboardRemove()
        )
        return WAITING_BEFORE_PHOTO

    else:
        # ردّ ذكي
        session = user_sessions.get(user_id, {})
        reply = ai_chat(text, session)
        keyboard = [
            [KeyboardButton("📸 تقرير عادي")],
            [KeyboardButton("🔄 تقرير قبل وبعد")],
        ]
        await update.message.reply_text(
            reply,
            reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
        )
        return MAIN_MENU


# ══════════════════════════════════════════════════
#  التقرير العادي
# ══════════════════════════════════════════════════

async def receive_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    session = user_sessions.get(user_id)
    if not session:
        await start(update, context)
        return WAITING_PHOTO

    photo = update.message.photo[-1]
    file = await context.bot.get_file(photo.file_id)
    photo_bytes = await file.download_as_bytearray()
    session["pending_photo"] = bytes(photo_bytes)

    # إذا عنده موقع سابق نسأله هل يستخدمه
    if session.get("last_location"):
        loc = session["last_location"]["name"]
        keyboard = [
            [KeyboardButton(f"✅ نفس الموقع: {loc[:30]}")],
            [KeyboardButton("📍 موقع جديد")],
        ]
        await update.message.reply_text(
            "📍 الموقع؟",
            reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
        )
    else:
        await ask_location(update)

    return WAITING_LOCATION


async def ask_location(update):
    keyboard = [
        [KeyboardButton("📍 مشاركة موقعي الآن", request_location=True)],
    ]
    await update.message.reply_text(
        "📍 *الموقع إجباري* — أرسله بإحدى الطرق:\n\n"
        "1️⃣ اضغط 'مشاركة موقعي الآن'\n"
        "2️⃣ أرسل رابط جوجل ماب\n"
        "3️⃣ أرسل الإحداثيات مثل: `23.1234, 53.6789`\n\n"
        "⚠️ يجب أن يكون الموقع ضمن نطاق منطقة ليوا",
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardMarkup(
            [[KeyboardButton("📍 مشاركة موقعي الآن", request_location=True)]],
            resize_keyboard=True
        )
    )


async def receive_location_gps(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """موقع GPS من تيليغرام"""
    user_id = update.effective_user.id
    session = user_sessions.get(user_id)
    if not session:
        return await start(update, context)

    loc = update.message.location
    await process_location(update, context, loc.latitude, loc.longitude, session)
    return WAITING_NOTE


async def receive_location_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """موقع نصي أو رابط أو إعادة استخدام"""
    user_id = update.effective_user.id
    session = user_sessions.get(user_id)
    if not session:
        return await start(update, context)

    text = update.message.text.strip()

    # إعادة استخدام نفس الموقع
    if "نفس الموقع" in text and session.get("last_location"):
        loc_data = session["last_location"]
        photo_entry = {
            "photo": session["pending_photo"],
            "location": loc_data,
            "note": "",
            "type": "normal"
        }
        session["photos"].append(photo_entry)
        session["pending_photo"] = None

        keyboard = [[KeyboardButton("⏭️ بدون ملاحظة")]]
        await update.message.reply_text(
            f"✅ موقع: {loc_data['name']}\n\n📝 أضف ملاحظة للصورة:",
            reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
        )
        return WAITING_NOTE

    # موقع جديد
    if "موقع جديد" in text:
        await ask_location(update)
        return WAITING_LOCATION

    # تحليل الإحداثيات أو الرابط
    lat, lng = parse_coordinates(text)
    if lat and lng:
        await process_location(update, context, lat, lng, session)
        return WAITING_NOTE
    else:
        await update.message.reply_text(
            "❌ لم أتعرف على الموقع.\n\n"
            "جرب:\n"
            "• رابط جوجل ماب\n"
            "• إحداثيات: `23.1234, 53.6789`\n"
            "• أو اضغط زر المشاركة",
            parse_mode="Markdown"
        )
        return WAITING_LOCATION


async def process_location(update, context, lat: float, lng: float, session: dict):
    """معالجة الموقع والتحقق من نطاق ليوا"""
    await update.message.reply_text("⏳ جاري التحقق من الموقع...")

    loc_data = ai_extract_location(lat, lng, session.get("pending_photo"))

    if not loc_data["in_liwa"]:
        await update.message.reply_text(
            f"⚠️ *الموقع خارج نطاق منطقة ليوا!*\n\n"
            f"الإحداثيات المُرسلة: `{lat:.4f}, {lng:.4f}`\n\n"
            "تأكد من أنك في منطقة ليوا وأعد مشاركة موقعك.",
            parse_mode="Markdown"
        )
        await ask_location(update)
        return

    # حفظ الصورة مع الموقع
    photo_entry = {
        "photo": session["pending_photo"],
        "location": loc_data,
        "note": "",
        "type": "normal"
    }
    session["photos"].append(photo_entry)
    session["pending_photo"] = None
    session["last_location"] = loc_data

    keyboard = [[KeyboardButton("⏭️ بدون ملاحظة")]]
    await update.message.reply_text(
        f"✅ الموقع: *{loc_data['name']}*\n\n📝 أضف ملاحظة للصورة:",
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    )


async def receive_note(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    session = user_sessions[user_id]
    text = update.message.text.strip()

    # حفظ الملاحظة على آخر صورة
    note = "" if "بدون ملاحظة" in text else text
    if session["photos"]:
        session["photos"][-1]["note"] = note

    count = len(session["photos"])

    keyboard = [
        [KeyboardButton("📸 إضافة صورة أخرى")],
        [KeyboardButton("✅ إنشاء التقرير")],
    ]
    await update.message.reply_text(
        f"✅ تمت إضافة الصورة {count}\n\nهل تريد إضافة المزيد؟",
        reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    )
    return ADD_MORE


async def add_more(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    text = update.message.text

    if "إضافة صورة" in text or "📸" in text:
        await update.message.reply_text(
            "📸 أرسل الصورة التالية:",
            reply_markup=ReplyKeyboardRemove()
        )
        return WAITING_PHOTO

    elif "إنشاء التقرير" in text or "✅" in text:
        return await confirm_and_send(update, context)

    else:
        # رد ذكي
        session = user_sessions.get(user_id, {})
        reply = ai_chat(text, session)
        keyboard = [
            [KeyboardButton("📸 إضافة صورة أخرى")],
            [KeyboardButton("✅ إنشاء التقرير")],
        ]
        await update.message.reply_text(
            reply,
            reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
        )
        return ADD_MORE


# ══════════════════════════════════════════════════
#  تقرير قبل وبعد
# ══════════════════════════════════════════════════

async def receive_before_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    session = user_sessions[user_id]

    photo = update.message.photo[-1]
    file = await context.bot.get_file(photo.file_id)
    photo_bytes = await file.download_as_bytearray()
    session["before_photo"] = bytes(photo_bytes)

    await update.message.reply_text(
        "✅ صورة 'قبل' مُستلمة\n\n📸 الآن أرسل صورة **بعد** التدخل:",
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardRemove()
    )
    return WAITING_AFTER_PHOTO


async def receive_after_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    session = user_sessions[user_id]

    photo = update.message.photo[-1]
    file = await context.bot.get_file(photo.file_id)
    photo_bytes = await file.download_as_bytearray()
    session["after_photo"] = bytes(photo_bytes)

    await ask_location(update)
    return WAITING_BEFORE_LOC


async def receive_before_after_location(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    session = user_sessions[user_id]

    if update.message.location:
        loc = update.message.location
        lat, lng = loc.latitude, loc.longitude
    else:
        text = update.message.text.strip()
        lat, lng = parse_coordinates(text)
        if not lat:
            await update.message.reply_text("❌ أرسل الموقع بشكل صحيح")
            return WAITING_BEFORE_LOC

    await update.message.reply_text("⏳ جاري التحقق من الموقع...")
    loc_data = ai_extract_location(lat, lng)

    if not loc_data["in_liwa"]:
        await update.message.reply_text(
            "⚠️ الموقع خارج نطاق ليوا! أعد الإرسال.",
            parse_mode="Markdown"
        )
        await ask_location(update)
        return WAITING_BEFORE_LOC

    session["photos"].append({
        "photo": session["before_photo"],
        "photo_after": session["after_photo"],
        "location": loc_data,
        "note": "",
        "type": "before_after"
    })
    session["last_location"] = loc_data

    keyboard = [[KeyboardButton("⏭️ بدون ملاحظة")]]
    await update.message.reply_text(
        f"✅ الموقع: *{loc_data['name']}*\n\n📝 ملاحظة على هذا الموقع:",
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    )
    return WAITING_NOTE


# ══════════════════════════════════════════════════
#  الاستفسارات الذكية (أي رسالة نصية غير متوقعة)
# ══════════════════════════════════════════════════

async def smart_reply(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    session = user_sessions.get(user_id, {})
    text = update.message.text

    reply = ai_chat(text, session)
    await update.message.reply_text(reply)
    return CHAT_MODE


# ══════════════════════════════════════════════════
#  إنشاء وإرسال التقرير
# ══════════════════════════════════════════════════

async def confirm_and_send(update, context):
    user_id = update.effective_user.id
    session = user_sessions[user_id]

    if not session.get("photos"):
        await update.message.reply_text("❌ لا توجد صور بعد!", reply_markup=ReplyKeyboardRemove())
        return ConversationHandler.END

    await update.message.reply_text("⏳ جاري إنشاء التقرير...", reply_markup=ReplyKeyboardRemove())

    try:
        buf = create_pptx_report(session)
        fname = f"تقرير_ليوا_{session['name']}_{session['date']}.pptx"

        # إرسال للموظف
        buf.seek(0)
        count = len(session["photos"])
        await update.message.reply_document(
            document=buf,
            filename=fname,
            caption=(
                f"✅ تقرير جاهز\n"
                f"👤 {session['name']}\n"
                f"📸 {count} صورة\n"
                f"📅 {session['date']}"
            )
        )

        # إرسال بالإيميل
        buf.seek(0)
        ok = send_email(buf, session, fname)
        if ok:
            await update.message.reply_text("📧 تم الإرسال للبريد ✅")
        else:
            await update.message.reply_text("⚠️ الملف جاهز هنا — راجع إعداد الإيميل")

        # تحديث الإحصائيات
        report_stats["total"] += 1
        session["sent"] = True

        await update.message.reply_text(
            f"🎉 تم! هذا هو التقرير رقم {report_stats['total']}\n\n"
            "اكتب /start لتقرير جديد"
        )

    except Exception as e:
        logger.error(f"Report error: {e}")
        await update.message.reply_text(
            f"❌ خطأ في إنشاء التقرير:\n{str(e)}\n\nاكتب /start وحاول مجدداً"
        )

    user_sessions.pop(user_id, None)
    return ConversationHandler.END


# ══════════════════════════════════════════════════
#  إنشاء ملف PowerPoint
# ══════════════════════════════════════════════════

def create_pptx_report(session: dict) -> BytesIO:
    prs = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)
    W = prs.slide_width
    H = prs.slide_height

    name  = session.get("name", "")
    date  = session.get("date", "")
    photos = session.get("photos", [])

    def add_bg(slide):
        bg = slide.shapes.add_shape(1, 0, 0, W, H)
        bg.fill.solid()
        bg.fill.fore_color.rgb = RGBColor(0xF5, 0xF3, 0xF0)
        bg.line.fill.background()

    def add_corners(slide):
        sz = Inches(0.45)
        for (x, y) in [(0,0),(W-sz,0),(0,H-sz),(W-sz,H-sz)]:
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

    def add_photo_card(slide, photo_bytes, x, y, w, h, location_name, note):
        """بطاقة صورة مع موقع وملاحظة"""
        info_lines = []
        if location_name:
            info_lines.append(("📍", location_name))
        if note:
            info_lines.append(("📝", note))

        info_h = Inches(0.38 * max(len(info_lines), 1) + 0.2)
        img_h  = h - info_h - Inches(0.1)

        # إطار
        frame = slide.shapes.add_shape(1, x, y, w, h)
        frame.fill.solid()
        frame.fill.fore_color.rgb = WHITE
        frame.line.color.rgb = BROWN
        frame.line.width = Pt(1.5)

        padding = Inches(0.1)

        # الصورة
        try:
            img_stream = BytesIO(photo_bytes)
            slide.shapes.add_picture(
                img_stream, x+padding, y+padding,
                w-padding*2, img_h-padding
            )
        except Exception as e:
            logger.warning(f"Photo: {e}")

        # خلفية المعلومات
        iy = y + img_h
        if info_lines:
            info_bg = slide.shapes.add_shape(1, x+padding, iy, w-padding*2, info_h)
            info_bg.fill.solid()
            info_bg.fill.fore_color.rgb = RGBColor(0xF8, 0xF5, 0xF2)
            info_bg.line.color.rgb = RGBColor(0xD0, 0xC0, 0xB0)
            info_bg.line.width = Pt(0.5)

            # خط بني
            accent = slide.shapes.add_shape(1, x+padding, iy, w-padding*2, Inches(0.04))
            accent.fill.solid()
            accent.fill.fore_color.rgb = BROWN
            accent.line.fill.background()

            lh = Inches(0.35)
            for i, (icon, val) in enumerate(info_lines):
                ly = iy + Inches(0.08) + i * lh
                add_text(slide, f"{icon} {val}",
                         x+Inches(0.15), ly, w-Inches(0.3), lh,
                         size=9, color=BLACK, align=PP_ALIGN.RIGHT)

    def make_cover():
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        add_bg(slide)
        add_corners(slide)

        bar = slide.shapes.add_shape(1, Inches(1), Inches(2.2), W-Inches(2), Inches(1.4))
        bar.fill.solid()
        bar.fill.fore_color.rgb = BROWN
        bar.line.fill.background()

        add_text(slide, "بوت ليوا — تقرير تفتيش",
                 Inches(1), Inches(2.35), W-Inches(2), Inches(0.75),
                 size=26, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_text(slide, "Liwa Inspection Report",
                 Inches(1), Inches(3.05), W-Inches(2), Inches(0.4),
                 size=13, color=RGBColor(0xCC,0xAA,0x88),
                 align=PP_ALIGN.CENTER, italic=True)

        info = [
            ("👤 المفتش:", name),
            ("📅 التاريخ:", date),
            ("📸 عدد الصور:", str(len(photos))),
            ("📋 نوع التقرير:", "قبل وبعد" if session.get("report_type") == "before_after" else "تفتيش عادي"),
        ]
        for i, (lbl, val) in enumerate(info):
            iy = Inches(3.8) + i * Inches(0.45)
            add_text(slide, lbl, Inches(7.5), iy, Inches(1.8), Inches(0.4),
                     size=11, bold=True, color=BROWN)
            add_text(slide, val, Inches(3.5), iy, Inches(4.2), Inches(0.4),
                     size=11, color=BLACK)

        add_text(slide, date, Inches(0.6), H-Inches(0.42),
                 Inches(3), Inches(0.35), size=10, color=BROWN, align=PP_ALIGN.LEFT)

    def make_normal_slide(entry, idx, total):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        add_bg(slide)
        add_corners(slide)

        add_text(slide, f"صورة {idx} من {total}",
                 Inches(0.6), Inches(0.15), Inches(4), Inches(0.35),
                 size=10, bold=True, color=BROWN, align=PP_ALIGN.RIGHT)
        add_text(slide, date, W-Inches(4.6), Inches(0.15),
                 Inches(4), Inches(0.35), size=10, color=BROWN, align=PP_ALIGN.LEFT)

        mx, my = Inches(0.8), Inches(0.55)
        add_photo_card(
            slide, entry["photo"],
            mx, my, W-mx*2, H-my-Inches(0.4),
            entry["location"]["name"], entry.get("note", "")
        )

    def make_before_after_slide(entry, idx):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        add_bg(slide)
        add_corners(slide)

        add_text(slide, f"قبل وبعد — موقع {idx}",
                 Inches(0.6), Inches(0.15), Inches(5), Inches(0.35),
                 size=11, bold=True, color=BROWN, align=PP_ALIGN.RIGHT)
        add_text(slide, date, W-Inches(4.6), Inches(0.15),
                 Inches(4), Inches(0.35), size=10, color=BROWN, align=PP_ALIGN.LEFT)

        mx, my = Inches(0.5), Inches(0.55)
        gap  = Inches(0.2)
        bw   = (W - mx*2 - gap) / 2
        bh   = H - my - Inches(0.4)

        loc_name = entry["location"]["name"]
        note     = entry.get("note", "")

        # قبل
        add_text(slide, "◄ قبل", mx, my-Inches(0.3), bw, Inches(0.28),
                 size=11, bold=True, color=BROWN, align=PP_ALIGN.RIGHT)
        add_photo_card(slide, entry["photo"], mx, my, bw, bh, loc_name, "")

        # بعد
        add_text(slide, "بعد ►", mx+bw+gap, my-Inches(0.3), bw, Inches(0.28),
                 size=11, bold=True, color=BROWN, align=PP_ALIGN.LEFT)
        add_photo_card(slide, entry["photo_after"], mx+bw+gap, my, bw, bh, loc_name, note)

    def make_closing():
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        add_bg(slide)
        add_corners(slide)

        fw, fh = Inches(5.5), Inches(2.8)
        fx = (W-fw)/2
        fy = (H-fh)/2
        frame = slide.shapes.add_shape(1, fx, fy, fw, fh)
        frame.fill.solid()
        frame.fill.fore_color.rgb = WHITE
        frame.line.color.rgb = BROWN
        frame.line.width = Pt(2)

        add_text(slide, "شكـراً",
                 fx, fy+Inches(0.4), fw, Inches(1.3),
                 size=44, bold=True, color=BROWN, align=PP_ALIGN.CENTER)
        add_text(slide, "Thank You",
                 fx, fy+Inches(1.7), fw, Inches(0.6),
                 size=16, color=RGBColor(0xAA,0x66,0x44),
                 align=PP_ALIGN.CENTER, italic=True)
        add_text(slide, date, Inches(0.6), H-Inches(0.42),
                 Inches(3), Inches(0.35), size=10, color=BROWN, align=PP_ALIGN.LEFT)

    # بناء الصفحات
    make_cover()

    for i, entry in enumerate(photos, 1):
        if entry.get("type") == "before_after":
            make_before_after_slide(entry, i)
        else:
            make_normal_slide(entry, i, len(photos))

    make_closing()

    buf = BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf


# ══════════════════════════════════════════════════
#  إرسال الإيميل
# ══════════════════════════════════════════════════

def send_email(buf: BytesIO, session: dict, fname: str) -> bool:
    if not GMAIL_USER or not GMAIL_APP_PASS:
        return False
    try:
        msg = MIMEMultipart()
        msg["From"]    = GMAIL_USER
        msg["To"]      = ADMIN_EMAIL
        msg["Subject"] = f"تقرير ليوا — {session.get('name','')} — {session.get('date','')}"

        locations = list(set([
            p["location"]["name"] for p in session.get("photos", [])
            if p.get("location")
        ]))
        body = (
            f"تقرير تفتيش جديد من بوت ليوا\n\n"
            f"المفتش: {session.get('name','—')}\n"
            f"التاريخ: {session.get('date','—')}\n"
            f"عدد الصور: {len(session.get('photos',[]))}\n"
            f"المواقع: {', '.join(locations) or '—'}\n"
            f"النوع: {'قبل وبعد' if session.get('report_type')=='before_after' else 'عادي'}\n"
        )
        msg.attach(MIMEText(body, "plain", "utf-8"))

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


# ══════════════════════════════════════════════════
#  تشغيل البوت
# ══════════════════════════════════════════════════

def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()

    conv = ConversationHandler(
        entry_points=[
            CommandHandler("start", start),
            MessageHandler(filters.PHOTO, receive_photo),
        ],
        states={
            MAIN_MENU: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, main_menu),
            ],
            WAITING_PHOTO: [
                MessageHandler(filters.PHOTO, receive_photo),
            ],
            WAITING_LOCATION: [
                MessageHandler(filters.LOCATION, receive_location_gps),
                MessageHandler(filters.TEXT & ~filters.COMMAND, receive_location_text),
            ],
            WAITING_NOTE: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, receive_note),
            ],
            ADD_MORE: [
                MessageHandler(filters.PHOTO, receive_photo),
                MessageHandler(filters.TEXT & ~filters.COMMAND, add_more),
            ],
            WAITING_BEFORE_PHOTO: [
                MessageHandler(filters.PHOTO, receive_before_photo),
            ],
            WAITING_AFTER_PHOTO: [
                MessageHandler(filters.PHOTO, receive_after_photo),
            ],
            WAITING_BEFORE_LOC: [
                MessageHandler(filters.LOCATION, receive_before_after_location),
                MessageHandler(filters.TEXT & ~filters.COMMAND, receive_before_after_location),
            ],
            CHAT_MODE: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, smart_reply),
            ],
        },
        fallbacks=[
            CommandHandler("start", start),
            MessageHandler(filters.TEXT & ~filters.COMMAND, smart_reply),
        ],
        allow_reentry=True,
    )

    app.add_handler(conv)
    print("🤖 بوت ليوا يعمل...")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
