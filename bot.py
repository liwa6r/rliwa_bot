"""
بوت ليوا v3 - نظام تقارير التفتيش والتوعية
"""

import os
import re
import base64
import logging
import smtplib
import requests as req
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

# ── الإصدار
BOT_VERSION = "v3.0"
BOT_NAME    = "بوت ليوا"

# ── أكواد الدخول
CODE_ADMIN      = "9090"   # مدير النظام — فحص شامل
CODE_AWARENESS  = "9091"   # قسم التوعية
CODE_INSPECTION = "9092"   # قسم التفتيش

# ── المتغيرات البيئية
TELEGRAM_TOKEN    = os.environ.get("TELEGRAM_TOKEN", "")
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
ADMIN_EMAIL       = "bot.6rr@gmail.com"
GMAIL_USER        = os.environ.get("GMAIL_USER", "")
GMAIL_APP_PASS    = os.environ.get("GMAIL_APP_PASS", "")
RESEND_API_KEY    = os.environ.get("RESEND_API_KEY", "")
GOOGLE_MAPS_KEY   = os.environ.get("GOOGLE_MAPS_KEY", "")

# ── نطاق ليوا والمنطقة الغربية
LIWA_BOUNDS = {
    "lat_min": 22.0, "lat_max": 24.5,
    "lng_min": 51.5, "lng_max": 55.5
}

# ── ألوان النموذج
BROWN    = RGBColor(0x8B, 0x25, 0x00)
LIGHT_BG = RGBColor(0xF2, 0xF0, 0xED)
WHITE    = RGBColor(0xFF, 0xFF, 0xFF)
BLACK    = RGBColor(0x1A, 0x1A, 0x1A)
TEAL     = RGBColor(0x00, 0x6B, 0x6B)   # لون تقارير التوعية

# ── مراحل المحادثة
(WAITING_CODE, MAIN_MENU,
 WAITING_PHOTO, WAITING_LOCATION, WAITING_AWARENESS_TYPE,
 WAITING_NOTE, ADD_MORE,
 WAITING_BEFORE_PHOTO, WAITING_AFTER_PHOTO,
 WAITING_BA_LOC, WAITING_BA_NOTE,
 CONFIRM_SEND) = range(12)

user_sessions = {}
report_stats  = {"total": 0, "inspection": 0, "awareness": 0}


# ══════════════════════════════════════════════════
#  Google Geocoding
# ══════════════════════════════════════════════════

def get_address(lat: float, lng: float) -> dict:
    in_liwa = (
        LIWA_BOUNDS["lat_min"] <= lat <= LIWA_BOUNDS["lat_max"] and
        LIWA_BOUNDS["lng_min"] <= lng <= LIWA_BOUNDS["lng_max"]
    )
    location_name = f"{lat:.4f}, {lng:.4f}"

    if GOOGLE_MAPS_KEY:
        try:
            url = (
                f"https://maps.googleapis.com/maps/api/geocode/json"
                f"?latlng={lat},{lng}&language=ar&key={GOOGLE_MAPS_KEY}"
            )
            r = req.get(url, timeout=6)
            data = r.json()
            if data.get("status") == "OK" and data.get("results"):
                components = data["results"][0].get("address_components", [])
                route = sublocality = locality = ""
                for c in components:
                    types = c.get("types", [])
                    name  = c.get("long_name", "")
                    if "route" in types and not route:
                        route = name
                    elif ("sublocality" in types or "sublocality_level_1" in types) and not sublocality:
                        sublocality = name
                    elif "locality" in types and not locality:
                        locality = name
                parts = [p for p in [route, sublocality] if p]
                if not parts and locality:
                    parts = [locality]
                if parts:
                    location_name = "، ".join(parts)
                else:
                    full = data["results"][0].get("formatted_address", "")
                    full = re.sub(r"الإمارات.*|Abu Dhabi.*|UAE.*", "", full)
                    location_name = full.strip().strip(",").strip()
        except Exception as e:
            logger.error(f"Geocoding: {e}")

    return {"name": location_name, "in_liwa": in_liwa, "lat": lat, "lng": lng}


def parse_coords(text: str):
    patterns = [
        r'(-?\d+\.?\d*)[,،\s]+(-?\d+\.?\d*)',
        r'[?&]q=(-?\d+\.?\d*),(-?\d+\.?\d*)',
        r'@(-?\d+\.?\d*),(-?\d+\.?\d*)',
        r'll=(-?\d+\.?\d*),(-?\d+\.?\d*)',
        r'!3d(-?\d+\.?\d*)!4d(-?\d+\.?\d*)',
    ]
    for p in patterns:
        m = re.search(p, text)
        if m:
            la, ln = float(m.group(1)), float(m.group(2))
            if -90 <= la <= 90 and -180 <= ln <= 180:
                return la, ln
    return None, None


# ══════════════════════════════════════════════════
#  Claude AI
# ══════════════════════════════════════════════════

def ai_reply(msg: str, session: dict) -> str:
    try:
        client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
        dept   = "التوعية" if session.get("dept") == "awareness" else "التفتيش"
        sys = (
            f"أنت {BOT_NAME} {BOT_VERSION} — مساعد تقارير {dept} في منطقة ليوا.\n"
            f"إجمالي التقارير: {report_stats['total']} "
            f"(تفتيش: {report_stats['inspection']} — توعية: {report_stats['awareness']})\n"
            "أجب بالعربية بشكل ودي ومختصر (جملتان فقط)."
        )
        r = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=150,
            system=sys,
            messages=[{"role": "user", "content": msg}]
        )
        return r.content[0].text.strip()
    except Exception:
        return f"أنا {BOT_NAME} {BOT_VERSION} — كيف أساعدك؟"


# ══════════════════════════════════════════════════
#  فحص النظام — كود المدير 9090
# ══════════════════════════════════════════════════

async def run_system_check(update: Update):
    await update.message.reply_text("🔍 جاري فحص النظام... لحظة")

    results = []
    all_ok  = True

    # 1. Claude AI
    try:
        client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
        client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=5,
            messages=[{"role": "user", "content": "ping"}]
        )
        results.append("✅ Claude AI — يعمل")
    except Exception as e:
        results.append(f"❌ Claude AI — {str(e)[:50]}")
        all_ok = False

    # 2. Google Maps
    try:
        url = (
            f"https://maps.googleapis.com/maps/api/geocode/json"
            f"?latlng=23.085,54.016&language=ar&key={GOOGLE_MAPS_KEY}"
        )
        r    = req.get(url, timeout=6)
        data = r.json()
        if data.get("status") == "OK":
            addr = data["results"][0].get("formatted_address", "")
            addr = re.sub(r"الإمارات.*|UAE.*", "", addr).strip()[:45]
            results.append(f"✅ Google Maps — يعمل\n    📍 {addr}")
        else:
            results.append(f"❌ Google Maps — {data.get('status')} (تحقق من Billing)")
            all_ok = False
    except Exception as e:
        results.append(f"❌ Google Maps — {str(e)[:50]}")
        all_ok = False

    # 3. Resend API
    try:
        r = req.get(
            "https://api.resend.com/domains",
            headers={"Authorization": f"Bearer {RESEND_API_KEY}"},
            timeout=8
        )
        if r.status_code == 200:
            results.append(f"✅ Resend — متصل ✓")
        else:
            results.append(f"❌ Resend — خطأ {r.status_code}")
            all_ok = False
    except Exception as e:
        results.append(f"❌ Resend — {str(e)[:50]}")
        all_ok = False

    # 4. إيميل تجريبي حقيقي
    if RESEND_API_KEY:
        try:
            r = req.post(
                "https://api.resend.com/emails",
                headers={
                    "Authorization": f"Bearer {RESEND_API_KEY}",
                    "Content-Type": "application/json"
                },
                json={
                    "from":    "بوت ليوا <onboarding@resend.dev>",
                    "to":      [ADMIN_EMAIL],
                    "subject": f"🔧 فحص النظام — {BOT_NAME} {BOT_VERSION}",
                    "text": (
                        f"تقرير فحص النظام\n\n"
                        f"البوت: {BOT_NAME} {BOT_VERSION}\n"
                        f"التاريخ: {datetime.now().strftime('%Y/%m/%d %H:%M')}\n\n"
                        f"إجمالي التقارير: {report_stats['total']}\n"
                        f"تقارير التفتيش: {report_stats['inspection']}\n"
                        f"تقارير التوعية: {report_stats['awareness']}\n\n"
                        + "\n".join(results)
                    )
                },
                timeout=15
            )
            if r.status_code == 200:
                results.append(f"✅ إيميل تجريبي — وصل إلى {ADMIN_EMAIL}")
            else:
                results.append(f"❌ إيميل تجريبي — {r.text[:50]}")
                all_ok = False
        except Exception as e:
            results.append(f"❌ إيميل تجريبي — {str(e)[:50]}")
            all_ok = False

    status = "🟢 كل شيء يعمل بشكل صحيح" if all_ok else "🔴 يوجد مشاكل تحتاج مراجعة"

    report = (
        f"━━━━━━━━━━━━━━━━━━━━\n"
        f"🔧 {BOT_NAME} {BOT_VERSION}\n"
        f"📊 تقرير صحة النظام\n"
        f"━━━━━━━━━━━━━━━━━━━━\n\n"
        + "\n".join(results) +
        f"\n\n━━━━━━━━━━━━━━━━━━━━\n"
        f"📈 إجمالي التقارير: {report_stats['total']}\n"
        f"🔍 تفتيش: {report_stats['inspection']}  |  📢 توعية: {report_stats['awareness']}\n"
        f"🕐 {datetime.now().strftime('%Y/%m/%d — %H:%M')}\n\n"
        f"{status}"
    )
    await update.message.reply_text(report)


# ══════════════════════════════════════════════════
#  بداية المحادثة — إدخال الكود
# ══════════════════════════════════════════════════

def new_session(user_id, name, dept):
    user_sessions[user_id] = {
        "photos":        [],
        "name":          name,
        "dept":          dept,
        "report_type":   "normal",
        "date":          datetime.now().strftime("%Y/%m/%d"),
        "last_location": None,
        "pending_photo": None,
    }


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        f"👷 أهلاً! أنا {BOT_NAME} {BOT_VERSION}\n\n"
        "🔐 أدخل كودك للدخول:",
        reply_markup=ReplyKeyboardRemove()
    )
    return WAITING_CODE


async def receive_code(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    code    = update.message.text.strip()
    name    = update.effective_user.first_name or "مستخدم"

    # ── مدير النظام
    if code == CODE_ADMIN:
        await update.message.reply_text(
            f"🔓 مرحباً بك مدير النظام!\n{BOT_NAME} {BOT_VERSION}",
            reply_markup=ReplyKeyboardRemove()
        )
        await run_system_check(update)
        return ConversationHandler.END

    # ── قسم التفتيش
    elif code == CODE_INSPECTION:
        new_session(user_id, name, "inspection")
        kb = [
            [KeyboardButton("📸 تقرير عادي")],
            [KeyboardButton("🔄 تقرير قبل وبعد")],
        ]
        await update.message.reply_text(
            f"✅ أهلاً {name}! قسم التفتيش\nاختر نوع التقرير:",
            reply_markup=ReplyKeyboardMarkup(kb, resize_keyboard=True)
        )
        return MAIN_MENU

    # ── قسم التوعية
    elif code == CODE_AWARENESS:
        new_session(user_id, name, "awareness")
        await update.message.reply_text(
            f"✅ أهلاً {name}! قسم التوعية\n\n📸 أرسل الصورة:",
            reply_markup=ReplyKeyboardRemove()
        )
        return WAITING_PHOTO

    # ── كود خاطئ
    else:
        await update.message.reply_text(
            "❌ الكود غير صحيح\nأعد المحاولة أو تواصل مع المسؤول"
        )
        return WAITING_CODE


# ══════════════════════════════════════════════════
#  القائمة الرئيسية (تفتيش فقط)
# ══════════════════════════════════════════════════

async def main_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    text    = update.message.text
    session = user_sessions.get(user_id, {})

    if "عادي" in text or "📸" in text:
        session["report_type"] = "normal"
        await update.message.reply_text("📸 أرسل الصورة:", reply_markup=ReplyKeyboardRemove())
        return WAITING_PHOTO
    elif "قبل وبعد" in text or "🔄" in text:
        session["report_type"] = "before_after"
        await update.message.reply_text(
            "📸 أرسل صورة *قبل* التدخل:", parse_mode="Markdown",
            reply_markup=ReplyKeyboardRemove()
        )
        return WAITING_BEFORE_PHOTO
    else:
        reply = ai_reply(text, session)
        kb = [[KeyboardButton("📸 تقرير عادي")], [KeyboardButton("🔄 تقرير قبل وبعد")]]
        await update.message.reply_text(reply, reply_markup=ReplyKeyboardMarkup(kb, resize_keyboard=True))
        return MAIN_MENU


# ══════════════════════════════════════════════════
#  استقبال الصور
# ══════════════════════════════════════════════════

async def receive_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    session = user_sessions.get(user_id)
    if not session:
        await start(update, context)
        return WAITING_CODE

    photo = update.message.photo[-1]
    file  = await context.bot.get_file(photo.file_id)
    pb    = await file.download_as_bytearray()
    session["pending_photo"] = bytes(pb)

    if session.get("last_location"):
        loc = session["last_location"]["name"]
        kb  = [[KeyboardButton("✅ نفس الموقع")], [KeyboardButton("📍 موقع جديد")]]
        await update.message.reply_text(
            f"📍 الموقع السابق: *{loc}*\nهل نفس الموقع؟",
            parse_mode="Markdown",
            reply_markup=ReplyKeyboardMarkup(kb, resize_keyboard=True)
        )
    else:
        await ask_location(update)
    return WAITING_LOCATION


async def ask_location(update):
    kb = [[KeyboardButton("📍 مشاركة موقعي", request_location=True)]]
    await update.message.reply_text(
        "📍 *الموقع إجباري* — أرسله بإحدى الطرق:\n\n"
        "1️⃣ اضغط 'مشاركة موقعي'\n"
        "2️⃣ أرسل رابط جوجل ماب\n"
        "3️⃣ أرسل إحداثيات: `23.085, 54.016`\n\n"
        "⚠️ يجب أن يكون في نطاق منطقة ليوا",
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardMarkup(kb, resize_keyboard=True)
    )


async def receive_location_gps(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    session = user_sessions.get(user_id)
    if not session:
        return await start(update, context)
    loc = update.message.location
    await process_location(update, session, loc.latitude, loc.longitude)
    return await next_after_location(session, update)


async def receive_location_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    session = user_sessions.get(user_id)
    if not session:
        return await start(update, context)

    text = update.message.text.strip()

    if "نفس الموقع" in text and session.get("last_location"):
        session["photos"].append({
            "photo":    session["pending_photo"],
            "location": session["last_location"],
            "note":     "",
            "awareness_type": "",
            "type":     "normal"
        })
        session["pending_photo"] = None
        return await next_after_location(session, update)

    if "موقع جديد" in text:
        await ask_location(update)
        return WAITING_LOCATION

    lat, lng = parse_coords(text)
    if lat:
        await process_location(update, session, lat, lng)
        return await next_after_location(session, update)

    await update.message.reply_text(
        "❌ لم أتعرف على الموقع.\nجرب رابط جوجل ماب أو إحداثيات مثل:\n`23.085, 54.016`",
        parse_mode="Markdown"
    )
    return WAITING_LOCATION


async def process_location(update, session, lat, lng):
    await update.message.reply_text("⏳ جاري تحديد الموقع...")
    loc_data = get_address(lat, lng)

    if not loc_data["in_liwa"]:
        await update.message.reply_text(
            f"⚠️ *الموقع خارج نطاق ليوا!*\n`{lat:.4f}, {lng:.4f}`\n\nتأكد من موقعك وأعد الإرسال.",
            parse_mode="Markdown"
        )
        await ask_location(update)
        return

    session["photos"].append({
        "photo":         session["pending_photo"],
        "location":      loc_data,
        "note":          "",
        "awareness_type": "",
        "type":          "normal"
    })
    session["pending_photo"]  = None
    session["last_location"]  = loc_data

    await update.message.reply_text(f"✅ *{loc_data['name']}*", parse_mode="Markdown")


async def next_after_location(session, update=None):
    """بعد تسجيل الموقع — توعية تسأل عن النوع، تفتيش تسأل عن الملاحظة"""
    if session.get("dept") == "awareness":
        if update:
            await update.message.reply_text(
                "📋 ما نوع التوعية؟",
                reply_markup=ReplyKeyboardRemove()
            )
        return WAITING_AWARENESS_TYPE
    else:
        if update:
            kb = [[KeyboardButton("⏭️ بدون ملاحظة")]]
            await update.message.reply_text(
                "📝 ملاحظة للصورة؟",
                reply_markup=ReplyKeyboardMarkup(kb, resize_keyboard=True)
            )
        return WAITING_NOTE


# ══════════════════════════════════════════════════
#  نوع التوعية (قسم التوعية فقط)
# ══════════════════════════════════════════════════

async def receive_awareness_type(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    session = user_sessions[user_id]
    text    = update.message.text.strip()

    if session["photos"]:
        session["photos"][-1]["awareness_type"] = text

    kb = [[KeyboardButton("⏭️ بدون ملاحظة")]]
    await update.message.reply_text(
        f"✅ نوع التوعية: *{text}*\n\n📝 ملاحظات إضافية؟",
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardMarkup(kb, resize_keyboard=True)
    )
    return WAITING_NOTE


# ══════════════════════════════════════════════════
#  الملاحظة
# ══════════════════════════════════════════════════

async def receive_note(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    session = user_sessions[user_id]
    text    = update.message.text.strip()

    note = "" if "بدون ملاحظة" in text else text
    if session["photos"]:
        session["photos"][-1]["note"] = note

    count = len(session["photos"])
    kb = [[KeyboardButton("📸 إضافة صورة")], [KeyboardButton("✅ إنشاء التقرير")]]
    await update.message.reply_text(
        f"✅ الصورة {count} مضافة\n\nإضافة المزيد أو إنشاء التقرير؟",
        reply_markup=ReplyKeyboardMarkup(kb, resize_keyboard=True)
    )
    return ADD_MORE


async def add_more(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    session = user_sessions.get(user_id, {})
    text    = update.message.text

    if "إضافة صورة" in text or "📸" in text:
        await update.message.reply_text("📸 أرسل الصورة:", reply_markup=ReplyKeyboardRemove())
        return WAITING_PHOTO
    elif "إنشاء التقرير" in text or "✅" in text:
        return await do_send(update, context)
    else:
        reply = ai_reply(text, session)
        kb = [[KeyboardButton("📸 إضافة صورة")], [KeyboardButton("✅ إنشاء التقرير")]]
        await update.message.reply_text(reply, reply_markup=ReplyKeyboardMarkup(kb, resize_keyboard=True))
        return ADD_MORE


# ══════════════════════════════════════════════════
#  تقرير قبل وبعد (تفتيش فقط)
# ══════════════════════════════════════════════════

async def receive_before_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    session = user_sessions[user_id]
    photo   = update.message.photo[-1]
    file    = await context.bot.get_file(photo.file_id)
    pb      = await file.download_as_bytearray()
    session["before_photo"] = bytes(pb)
    await update.message.reply_text(
        "✅ صورة 'قبل' مستلمة\n\n📸 أرسل صورة *بعد*:",
        parse_mode="Markdown", reply_markup=ReplyKeyboardRemove()
    )
    return WAITING_AFTER_PHOTO


async def receive_after_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    session = user_sessions[user_id]
    photo   = update.message.photo[-1]
    file    = await context.bot.get_file(photo.file_id)
    pb      = await file.download_as_bytearray()
    session["after_photo"] = bytes(pb)
    await ask_location(update)
    return WAITING_BA_LOC


async def receive_ba_location_gps(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    session = user_sessions[user_id]
    loc     = update.message.location
    await process_ba_location(update, session, loc.latitude, loc.longitude)
    return WAITING_BA_NOTE


async def receive_ba_location_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    session = user_sessions[user_id]
    lat, lng = parse_coords(update.message.text.strip())
    if lat:
        await process_ba_location(update, session, lat, lng)
        return WAITING_BA_NOTE
    await update.message.reply_text("❌ أرسل رابط أو إحداثيات صحيحة")
    return WAITING_BA_LOC


async def process_ba_location(update, session, lat, lng):
    await update.message.reply_text("⏳ جاري تحديد الموقع...")
    loc_data = get_address(lat, lng)
    if not loc_data["in_liwa"]:
        await update.message.reply_text("⚠️ الموقع خارج نطاق ليوا! أعد الإرسال.")
        await ask_location(update)
        return
    session["photos"].append({
        "photo":       session["before_photo"],
        "photo_after": session["after_photo"],
        "location":    loc_data,
        "note":        "",
        "type":        "before_after"
    })
    session["last_location"] = loc_data
    kb = [[KeyboardButton("⏭️ بدون ملاحظة")]]
    await update.message.reply_text(
        f"✅ *{loc_data['name']}*\n\n📝 ملاحظة؟",
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardMarkup(kb, resize_keyboard=True)
    )


async def receive_ba_note(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    session = user_sessions[user_id]
    text    = update.message.text.strip()
    if session["photos"]:
        session["photos"][-1]["note"] = "" if "بدون ملاحظة" in text else text
    kb = [[KeyboardButton("📸 إضافة موقع آخر")], [KeyboardButton("✅ إنشاء التقرير")]]
    await update.message.reply_text(
        "✅ تمت الإضافة\n\nإضافة موقع آخر أو إنشاء التقرير؟",
        reply_markup=ReplyKeyboardMarkup(kb, resize_keyboard=True)
    )
    return ADD_MORE


# ══════════════════════════════════════════════════
#  إنشاء PowerPoint
# ══════════════════════════════════════════════════

def create_pptx(session: dict) -> BytesIO:
    prs = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)
    W = prs.slide_width
    H = prs.slide_height

    is_awareness = session.get("dept") == "awareness"
    name         = session.get("name", "")
    date         = session.get("date", "")
    photos       = session.get("photos", [])
    header_color = TEAL if is_awareness else BROWN
    report_title = "تقرير توعية" if is_awareness else "تقرير تفتيش"

    def add_bg(slide):
        s = slide.shapes.add_shape(1, 0, 0, W, H)
        s.fill.solid(); s.fill.fore_color.rgb = RGBColor(0xF5, 0xF3, 0xF0)
        s.line.fill.background()

    def add_corners(slide):
        sz = Inches(0.45)
        for x, y in [(0,0),(W-sz,0),(0,H-sz),(W-sz,H-sz)]:
            s = slide.shapes.add_shape(1, x, y, sz, sz)
            s.fill.solid(); s.fill.fore_color.rgb = header_color
            s.line.fill.background()

    def txt(slide, text, x, y, w, h, size=10, bold=False,
            color=BLACK, align=PP_ALIGN.RIGHT, italic=False):
        tb = slide.shapes.add_textbox(x, y, w, h)
        tf = tb.text_frame; tf.word_wrap = True
        p  = tf.paragraphs[0]; p.alignment = align
        run = p.add_run(); run.text = str(text)
        run.font.size = Pt(size); run.font.bold = bold
        run.font.italic = italic; run.font.color.rgb = color
        run.font.name = "Arial"

    def photo_card(slide, photo_bytes, x, y, w, h, loc_name, note, awareness_type=""):
        # حساب ارتفاع البيانات — كل عنصر له سطرين (تسمية + قيمة)
        items = []
        if loc_name:       items.append(("📍 الموقع", loc_name))
        if awareness_type: items.append(("📋 نوع التوعية", awareness_type))
        if note:           items.append(("📝 الملاحظات", note))

        row_h  = Inches(0.55)  # ارتفاع كل عنصر (تسمية + قيمة)
        info_h = Inches(row_h.inches * max(len(items), 1) + 0.2)
        img_h  = h - info_h - Inches(0.12)
        pad    = Inches(0.12)

        # إطار الكارد
        frame = slide.shapes.add_shape(1, x, y, w, h)
        frame.fill.solid(); frame.fill.fore_color.rgb = WHITE
        frame.line.color.rgb = header_color; frame.line.width = Pt(1.5)

        # الصورة — أصغر قليلاً
        try:
            slide.shapes.add_picture(
                BytesIO(photo_bytes),
                x + pad, y + pad,
                w - pad*2, img_h - pad
            )
        except Exception as e:
            logger.warning(f"Photo: {e}")

        # خلفية البيانات
        iy = y + img_h
        if items:
            bg = slide.shapes.add_shape(1, x+pad, iy, w-pad*2, info_h)
            bg.fill.solid(); bg.fill.fore_color.rgb = RGBColor(0xF8,0xF5,0xF2)
            bg.line.color.rgb = RGBColor(0xD0,0xC0,0xB0); bg.line.width = Pt(0.5)

            # خط ملون في الأعلى
            ac = slide.shapes.add_shape(1, x+pad, iy, w-pad*2, Inches(0.04))
            ac.fill.solid(); ac.fill.fore_color.rgb = header_color; ac.line.fill.background()

            for i, (label, val) in enumerate(items):
                base_y = iy + Inches(0.08) + i * row_h
                # التسمية
                txt(slide, label,
                    x+Inches(0.15), base_y, w-Inches(0.3), Inches(0.22),
                    size=8, bold=True, color=header_color, align=PP_ALIGN.RIGHT)
                # القيمة
                txt(slide, val,
                    x+Inches(0.15), base_y+Inches(0.22), w-Inches(0.3), Inches(0.3),
                    size=9, color=BLACK, align=PP_ALIGN.RIGHT)

    def make_cover():
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        add_bg(slide); add_corners(slide)

        bar = slide.shapes.add_shape(1, Inches(1), Inches(2.0), W-Inches(2), Inches(1.5))
        bar.fill.solid(); bar.fill.fore_color.rgb = header_color; bar.line.fill.background()

        txt(slide, f"{BOT_NAME} — {report_title}",
            Inches(1), Inches(2.15), W-Inches(2), Inches(0.8),
            size=26, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        txt(slide, f"Liwa Bot  |  {BOT_VERSION}",
            Inches(1), Inches(2.92), W-Inches(2), Inches(0.4),
            size=12, color=RGBColor(0xCC,0xAA,0x88),
            align=PP_ALIGN.CENTER, italic=True)

        info = [
            ("👤 المفتش:" if not is_awareness else "👤 المسؤول:", name),
            ("📅 التاريخ:", date),
            ("📸 عدد الصور:", str(len(photos))),
            ("📋 القسم:", "التوعية" if is_awareness else "التفتيش"),
        ]
        for i, (lbl, val) in enumerate(info):
            iy = Inches(3.85) + i * Inches(0.46)
            txt(slide, lbl, Inches(7.5), iy, Inches(1.8), Inches(0.4),
                size=11, bold=True, color=header_color)
            txt(slide, val, Inches(3.5), iy, Inches(4.2), Inches(0.4),
                size=11, color=BLACK)

        txt(slide, date, Inches(0.6), H-Inches(0.42), Inches(3), Inches(0.35),
            size=10, color=header_color, align=PP_ALIGN.LEFT)
        txt(slide, BOT_VERSION, W-Inches(2), H-Inches(0.42), Inches(1.8), Inches(0.35),
            size=9, color=header_color, align=PP_ALIGN.LEFT, italic=True)

    def make_single_slide(entry, idx, total):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        add_bg(slide); add_corners(slide)
        txt(slide, f"صورة {idx} من {total}", Inches(0.6), Inches(0.15),
            Inches(4), Inches(0.35), size=10, bold=True, color=header_color, align=PP_ALIGN.RIGHT)
        txt(slide, date, W-Inches(4.6), Inches(0.15), Inches(4), Inches(0.35),
            size=10, color=header_color, align=PP_ALIGN.LEFT)

        mx, my = Inches(3.2), Inches(0.65)
        bw     = W - mx * 2
        bh     = H - my - Inches(0.5)
        photo_card(slide, entry["photo"], mx, my, bw, bh,
                   entry["location"]["name"],
                   entry.get("note",""),
                   entry.get("awareness_type",""))

    def make_double_slide(e1, e2, page_num):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        add_bg(slide); add_corners(slide)
        txt(slide, f"صفحة {page_num}", Inches(0.6), Inches(0.15),
            Inches(4), Inches(0.35), size=10, bold=True, color=header_color, align=PP_ALIGN.RIGHT)
        txt(slide, date, W-Inches(4.6), Inches(0.15), Inches(4), Inches(0.35),
            size=10, color=header_color, align=PP_ALIGN.LEFT)

        mx, my = Inches(0.5), Inches(0.55)
        gap    = Inches(0.22)
        bw     = (W - mx*2 - gap) / 2
        bh     = H - my - Inches(0.38)

        photo_card(slide, e1["photo"], mx, my, bw, bh,
                   e1["location"]["name"], e1.get("note",""), e1.get("awareness_type",""))
        photo_card(slide, e2["photo"], mx+bw+gap, my, bw, bh,
                   e2["location"]["name"], e2.get("note",""), e2.get("awareness_type",""))

    def make_ba_slide(entry, idx):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        add_bg(slide); add_corners(slide)
        txt(slide, f"قبل وبعد — موقع {idx}", Inches(0.6), Inches(0.15),
            Inches(6), Inches(0.35), size=11, bold=True, color=header_color, align=PP_ALIGN.RIGHT)
        txt(slide, date, W-Inches(4.6), Inches(0.15), Inches(4), Inches(0.35),
            size=10, color=header_color, align=PP_ALIGN.LEFT)

        mx, my = Inches(0.5), Inches(0.55)
        gap    = Inches(0.22)
        bw     = (W - mx*2 - gap) / 2
        bh     = H - my - Inches(0.38)

        txt(slide, "◄ قبل", mx, my-Inches(0.3), bw, Inches(0.28),
            size=11, bold=True, color=header_color, align=PP_ALIGN.RIGHT)
        photo_card(slide, entry["photo"], mx, my, bw, bh,
                   entry["location"]["name"], "")

        txt(slide, "بعد ►", mx+bw+gap, my-Inches(0.3), bw, Inches(0.28),
            size=11, bold=True, color=header_color, align=PP_ALIGN.LEFT)
        photo_card(slide, entry["photo_after"], mx+bw+gap, my, bw, bh,
                   entry["location"]["name"], entry.get("note",""))

    def make_closing():
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        add_bg(slide); add_corners(slide)
        fw, fh = Inches(5.5), Inches(2.8)
        fx, fy = (W-fw)/2, (H-fh)/2
        frame = slide.shapes.add_shape(1, fx, fy, fw, fh)
        frame.fill.solid(); frame.fill.fore_color.rgb = WHITE
        frame.line.color.rgb = header_color; frame.line.width = Pt(2)
        txt(slide, "شكـراً", fx, fy+Inches(0.35), fw, Inches(1.4),
            size=44, bold=True, color=header_color, align=PP_ALIGN.CENTER)
        txt(slide, "Thank You", fx, fy+Inches(1.75), fw, Inches(0.6),
            size=16, color=RGBColor(0xAA,0x66,0x44),
            align=PP_ALIGN.CENTER, italic=True)
        txt(slide, date, Inches(0.6), H-Inches(0.42), Inches(3), Inches(0.35),
            size=10, color=header_color, align=PP_ALIGN.LEFT)
        txt(slide, BOT_VERSION, W-Inches(2), H-Inches(0.42), Inches(1.8), Inches(0.35),
            size=9, color=header_color, align=PP_ALIGN.LEFT, italic=True)

    # بناء الصفحات
    make_cover()

    normal  = [e for e in photos if e.get("type") != "before_after"]
    ba_list = [e for e in photos if e.get("type") == "before_after"]

    page = 1
    i = 0
    while i < len(normal):
        e1 = normal[i]
        if i+1 < len(normal) and normal[i+1]["location"]["name"] == e1["location"]["name"]:
            make_double_slide(e1, normal[i+1], page)
            i += 2
        else:
            make_single_slide(e1, i+1, len(normal))
            i += 1
        page += 1

    for j, entry in enumerate(ba_list, 1):
        make_ba_slide(entry, j)

    make_closing()

    buf = BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf


# ══════════════════════════════════════════════════
#  إرسال الإيميل
# ══════════════════════════════════════════════════

def send_email(buf: BytesIO, session: dict, fname: str) -> bool:
    if not RESEND_API_KEY:
        return False
    try:
        is_awareness = session.get("dept") == "awareness"
        dept_ar      = "توعية" if is_awareness else "تفتيش"

        locs = list(dict.fromkeys([
            p["location"]["name"] for p in session.get("photos",[]) if p.get("location")
        ]))
        body = (
            f"تقرير {dept_ar} جديد — {BOT_NAME} {BOT_VERSION}\n\n"
            f"المسؤول: {session.get('name','—')}\n"
            f"التاريخ: {session.get('date','—')}\n"
            f"الصور: {len(session.get('photos',[]))}\n"
            f"المواقع: {', '.join(locs) or '—'}\n"
        )

        # تحويل الملف لـ base64
        file_content = base64.b64encode(buf.read()).decode()

        # إرسال عبر Resend API
        response = req.post(
            "https://api.resend.com/emails",
            headers={
                "Authorization": f"Bearer {RESEND_API_KEY}",
                "Content-Type": "application/json"
            },
            json={
                "from":    "بوت ليوا <onboarding@resend.dev>",
                "to":      [ADMIN_EMAIL],
                "subject": f"تقرير {dept_ar} — {session.get('name','')} — {session.get('date','')}",
                "text":    body,
                "attachments": [{
                    "filename": fname,
                    "content":  file_content
                }]
            },
            timeout=15
        )

        if response.status_code == 200:
            return True
        else:
            logger.error(f"Resend error: {response.text}")
            return False

    except Exception as e:
        logger.error(f"Email: {e}")
        return False


# ══════════════════════════════════════════════════
#  إنشاء وإرسال التقرير
# ══════════════════════════════════════════════════

async def do_send(update, context):
    user_id = update.effective_user.id
    session = user_sessions.get(user_id, {})

    if not session.get("photos"):
        await update.message.reply_text("❌ لا توجد صور!", reply_markup=ReplyKeyboardRemove())
        return ConversationHandler.END

    await update.message.reply_text("⏳ جاري إنشاء التقرير...", reply_markup=ReplyKeyboardRemove())

    try:
        is_awareness = session.get("dept") == "awareness"
        dept_ar      = "توعية" if is_awareness else "تفتيش"
        buf   = create_pptx(session)
        fname = f"تقرير_{dept_ar}_{session['name']}_{session['date']}.pptx"

        buf.seek(0)
        await update.message.reply_document(
            document=buf, filename=fname,
            caption=(
                f"✅ تقرير {dept_ar} جاهز — {BOT_VERSION}\n"
                f"👤 {session['name']}\n"
                f"📸 {len(session['photos'])} صورة\n"
                f"📅 {session['date']}"
            )
        )

        # إرسال الإيميل والتحقق
        buf.seek(0)
        await update.message.reply_text("📧 جاري التحقق من الإرسال...")
        ok = send_email(buf, session, fname)
        if ok:
            await update.message.reply_text(
                f"✅ تم إرسال التقرير إلى البريد الإلكتروني بنجاح!\n📩 {ADMIN_EMAIL}"
            )
        else:
            await update.message.reply_text(
                "⚠️ الإيميل لا يعمل حالياً — سيتم الإصلاح قريباً\n"
                "الملف متاح هنا للتحميل 👆"
            )

        # تحديث الإحصائيات
        report_stats["total"] += 1
        if is_awareness:
            report_stats["awareness"] += 1
        else:
            report_stats["inspection"] += 1

        await update.message.reply_text(
            f"🎉 تقرير رقم {report_stats['total']} مكتمل!\n\nاكتب /start لتقرير جديد"
        )

    except Exception as e:
        logger.error(f"Send error: {e}")
        await update.message.reply_text(f"❌ خطأ: {e}\n\nاكتب /start وحاول مجدداً")

    user_sessions.pop(user_id, None)
    return ConversationHandler.END


# ══════════════════════════════════════════════════
#  تشغيل البوت
# ══════════════════════════════════════════════════

def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()

    conv = ConversationHandler(
        entry_points=[
            CommandHandler("start", start),
        ],
        states={
            WAITING_CODE:        [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_code)],
            MAIN_MENU:           [MessageHandler(filters.TEXT & ~filters.COMMAND, main_menu)],
            WAITING_PHOTO:       [MessageHandler(filters.PHOTO, receive_photo)],
            WAITING_LOCATION:    [
                MessageHandler(filters.LOCATION, receive_location_gps),
                MessageHandler(filters.TEXT & ~filters.COMMAND, receive_location_text),
            ],
            WAITING_AWARENESS_TYPE: [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_awareness_type)],
            WAITING_NOTE:        [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_note)],
            ADD_MORE:            [
                MessageHandler(filters.PHOTO, receive_photo),
                MessageHandler(filters.TEXT & ~filters.COMMAND, add_more),
            ],
            WAITING_BEFORE_PHOTO: [MessageHandler(filters.PHOTO, receive_before_photo)],
            WAITING_AFTER_PHOTO:  [MessageHandler(filters.PHOTO, receive_after_photo)],
            WAITING_BA_LOC:      [
                MessageHandler(filters.LOCATION, receive_ba_location_gps),
                MessageHandler(filters.TEXT & ~filters.COMMAND, receive_ba_location_text),
            ],
            WAITING_BA_NOTE:     [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_ba_note)],
        },
        fallbacks=[CommandHandler("start", start)],
        allow_reentry=True,
    )

    app.add_handler(conv)
    print(f"🤖 {BOT_NAME} {BOT_VERSION} يعمل...")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
