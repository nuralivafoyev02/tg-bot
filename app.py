import os
import json
import telebot
import requests
from telebot import types
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

# 🔐 Tokenlar (bularni o'zing bilan almashtir)
BOT_TOKEN = "8142445503:AAHoF0S0of-bCQmPkOitpHHg9KXTL51bPIg"
GEMINI_API_KEY = "AIzaSyAs3WE0NyNn_uuNQG2KVUA_deQbYGEOG-8"

bot = telebot.TeleBot(BOT_TOKEN)


# ========== GEMINI AI FUNKSIYASI ==========
def generate_ppt_content(topic):
    """Gemini AI yordamida JSON formatda slaydlar kontenti yaratish"""
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key={GEMINI_API_KEY}"

    prompt = f"""
    Sen professional taqdimot dizaynerisan.
    Mavzu: "{topic}"
    Kamida 5 ta, ko‘pi bilan 9 ta slayd bo‘lsin.
    Quyidagi formatda JSON qaytar:
    [
      {{
        "sarlavha": "Kirish",
        "matn": "Mavzuga kirish, umumiy tushuncha va e’tibor jalb qiluvchi qisqa gaplar."
      }},
      {{
        "sarlavha": "Asosiy qism",
        "matn": "Asosiy g‘oyalar, faktlar, tahlillar va muhim ma’lumotlar."
      }},
      {{
        "sarlavha": "Xulosa",
        "matn": "Yakuniy fikr, tavsiya va motivatsion xulosa."
      }}
    ]
    Har bir sarlavha 2–5 so‘zdan iborat bo‘lsin. Matn to'liq, mano va mazmunli. chiroyli va foydali bo‘lsin. Imloviy xatoliklarga yo'l qoyilmasin.
    """

    headers = {"Content-Type": "application/json"}
    payload = {"contents": [{"parts": [{"text": prompt}]}]}

    try:
        response = requests.post(url, headers=headers, json=payload, timeout=30)
        response.raise_for_status()
        data = response.json()
        text = data["candidates"][0]["content"]["parts"][0]["text"]
        text = text.replace("```json", "").replace("```", "").strip()
        slides = json.loads(text)
        return slides
    except Exception as e:
        print("❌ AI kontent xatosi:", e)
        print("🔍 Gemini javobi:", response.text if 'response' in locals() else "javob yo‘q")
        return None


# ========== DIZAYNLI POWERPOINT YARATISH ==========
def create_ppt(topic, slides):
    """AI kontent asosida dizaynli PowerPoint yaratish"""
    prs = Presentation()

    for i, s in enumerate(slides):
        slide = prs.slides.add_slide(prs.slide_layouts[6])  # Bo‘sh slayd

        # 🎨 Fon rangi (gradient-style)
        background = slide.background
        fill = background.fill
        fill.solid()
        if i % 2 == 0:
            fill.fore_color.rgb = RGBColor(245, 247, 255)  # och moviy
        else:
            fill.fore_color.rgb = RGBColor(255, 248, 240)  # och sariq

        # 🖋️ Sarlavha joylashuvi
        left = Inches(1)
        top = Inches(1)
        width = Inches(8)
        height = Inches(1.5)

        title_box = slide.shapes.add_textbox(left, top, width, height)
        tf = title_box.text_frame
        tf.text = s.get("sarlavha", "").title()

        title_run = tf.paragraphs[0].runs[0]
        title_run.font.bold = True
        title_run.font.size = Pt(42)
        title_run.font.color.rgb = RGBColor(25, 25, 112)  # navy

        # 📘 Matn qismi
        body_box = slide.shapes.add_textbox(Inches(1), Inches(2.5), Inches(8), Inches(3))
        tf_body = body_box.text_frame
        tf_body.text = s.get("matn", "")

        for paragraph in tf_body.paragraphs:
            for run in paragraph.runs:
                run.font.name = "Calibri (Body)"
                run.font.size = Pt(24)
                run.font.color.rgb = RGBColor(40, 40, 40)

        # 💠 Dekorativ element (minimalist gradient panel)
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(10), Inches(0.4))
        fill = shape.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(0, 120, 212)  # ko‘k chiziq
        shape.line.fill.background()

    filename = f"{topic}.pptx"
    prs.save(filename)
    return filename


# ========== TELEGRAM BOT INTERFACE ==========
@bot.message_handler(commands=['start'])
def start(message):
    user_name = message.from_user.first_name
    welcome_text = (
        f"👋 Salom, *{user_name}!* \n\n"
        "Men — yorug‘lik tezligida ishlaydigan, sun’iy intellekt asosidagi *slayd tayyorlovchi bot*man. ⚡\n\n"
        "💡 Nimalarga qodirman:\n"
        "• Mavzuni tahlil qilib, foydali va chiroyli matn yarataman.\n"
        "• Har bir slaydni professional dizayn bilan bezayman.\n"
        "• Listlar 5 tadan 9 tagacha bo'ladi.\n"
        "• Tayyor PowerPoint faylni darhol yuboraman.\n\n"
        "Boshlaymizmi? 👇"
    )

    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    slide_btn = types.KeyboardButton("📑 Slayd tayyorlash")
    markup.add(slide_btn)

    bot.send_message(message.chat.id, welcome_text, parse_mode="Markdown", reply_markup=markup)


@bot.message_handler(func=lambda msg: msg.text == "📑 Slayd tayyorlash")
def ask_topic(message):
    bot.send_message(message.chat.id, "🧠 Qaysi mavzu bo‘yicha slayd tayyorlay? Mavzuni yozing ✍️")
    bot.register_next_step_handler(message, process_topic)


def process_topic(message):
    topic = message.text
    bot.send_message(message.chat.id, f"Bu atigi bir necha soniya vaqt oladi⚡\n\n'{topic}' mavzusi bo‘yicha slaydlar tayyorlanmoqda...")

    slides = generate_ppt_content(topic)
    if not slides:
        bot.send_message(message.chat.id, "⚠️ Kechirasiz, AI kontent yaratishda xatolik bo‘ldi.")
        return

    ppt_file = create_ppt(topic, slides)
    with open(ppt_file, "rb") as f:
        bot.send_document(message.chat.id, f)

    os.remove(ppt_file)
    bot.send_message(message.chat.id, "✅ Tayyor! Yana biror mavzu xohlaysizmi?")


# ========== BOTNI ISHGA TUSHURISH ==========
print("🚀 Bot ishga tushdi — dizaynli slaydlar yaratishga tayyor!")
bot.infinity_polling()
