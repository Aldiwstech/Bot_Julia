import os
import re

# PyTelegramBotAPI
try:
    import telebot
except ImportError as exc:
    raise RuntimeError(
        "PyTelegramBotAPI belum terinstall. Tambahkan pyTelegramBotAPI ke requirements.txt."
    ) from exc
import pandas as pd
from PIL import Image, ImageDraw, ImageFont

TOKEN = os.getenv("TOKEN")
if not TOKEN:
    raise ValueError("Token belum diset di Environment Variables Railway!")

bot = telebot.TeleBot(TOKEN, parse_mode=None)

EXCEL_CANDIDATES = ["Excel_master.xlsx", "Excel_master(1).xlsx"]
MOCKUP_CANDIDATES = ["Mokup.png", "mokup2.png", "mokup(1).png", "Mokup(1).png"]

# Mockup final yang Anda upload: 1672 x 941 px
BASE_W = 1672
BASE_H = 941

def get_font(size, bold=False):
    candidates = []
    env_path = os.getenv("FONT_PATH")
    if env_path:
        candidates.append(env_path)

    if bold:
        candidates += [
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
            "/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf",
            "DejaVuSans-Bold.ttf",
            "LiberationSans-Bold.ttf",
        ]
    else:
        candidates += [
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
            "DejaVuSans.ttf",
            "LiberationSans-Regular.ttf",
        ]

    for path in candidates:
        if path and os.path.exists(path):
            try:
                return ImageFont.truetype(path, size)
            except Exception:
                pass
    return ImageFont.load_default()

def find_existing(candidates):
    for path in candidates:
        if os.path.exists(path):
            return path
    return None

def clean_filename(text):
    return re.sub(r"[^A-Za-z0-9._-]+", "_", str(text))

def safe_float(value):
    try:
        if pd.isna(value):
            return None
        return float(str(value).strip().replace(",", "."))
    except Exception:
        return None

def is_empty(value):
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except Exception:
        pass
    return str(value).strip().lower() in ("", "nan", "none", "-")

def display_value(value, default="-"):
    if is_empty(value):
        return default
    return str(value).strip()

def fit_font(draw, text, max_width, size=21, minimum=12, bold=True):
    current = size
    while current > minimum:
        font = get_font(current, bold=bold)
        bbox = draw.textbbox((0, 0), str(text), font=font)
        if bbox[2] - bbox[0] <= max_width:
            return font
        current -= 1
    return get_font(minimum, bold=bold)

def dynamic_color(value):
    text = str(value).strip().upper()

    red_words = (
        "CRITICAL", "DOWN", "NOT AVAILABLE", "NOT_AVAILABLE",
        "NEED", "TRIP", "PERGANTIAN", "FAULT", "FAILED",
        "ERROR", "NO BACKUP",
    )
    orange_words = (
        "WARNING", "POTENSIAL", "POTENTIAL", "MEDIUM",
        "FAIR", "LOW", "CHECK",
    )
    green_words = (
        "NORMAL", "SECURED", "OK", "VALID", "AVAILABLE",
        "MONITOR", "SAFE",
    )

    if any(word in text for word in red_words):
        return (211, 42, 50)
    if any(word in text for word in orange_words):
        return (232, 142, 18)
    if any(word in text for word in green_words):
        return (20, 142, 68)
    return (24, 55, 105)

def load_site(site_id):
    excel_path = find_existing(EXCEL_CANDIDATES)
    if not excel_path:
        raise FileNotFoundError("Excel_master.xlsx tidak ditemukan.")

    df = pd.read_excel(excel_path, sheet_name="Rectifire&battery")
    if "Site ID" not in df.columns:
        raise KeyError("Kolom 'Site ID' tidak ditemukan.")

    wanted = str(site_id).strip().upper()
    ids = df["Site ID"].astype(str).str.strip().str.upper()
    result = df.loc[ids == wanted]

    return None if result.empty else result.iloc[0]

def value(row, column, default="-"):
    if column not in row.index:
        return default
    return display_value(row[column], default)

def build_actions(row):
    actions = []

    for col in ("Action", "Activity (SOW) Actual"):
        if col in row.index and not is_empty(row[col]):
            text = str(row[col]).strip()
            if text != "-":
                actions.append(text)
                break

    eas = value(row, "EAS Validation").upper()
    neteco = value(row, "NETECO Status").upper()
    rect = value(row, "Rectifier Condition").upper()

    if "DOWN" in rect or "CRITICAL" in rect:
        actions.append("NEED REPLACE RECTIFIER")
    if "NEED" in eas or "VALIDATION" in eas:
        actions.append("NEED VALIDATE")
    if "NOT AVAILABLE" in neteco or "DOWN" in neteco:
        actions.append("NEED CHECK NETECO")

    unique = []
    for item in actions:
        if item not in unique:
            unique.append(item)
    return unique or ["NORMAL"]

def generate_site_card(site_id):
    row = load_site(site_id)
    if row is None:
        return None

    mockup_path = find_existing(MOCKUP_CANDIDATES)
    if not mockup_path:
        raise FileNotFoundError("Mokup.png tidak ditemukan.")

    img = Image.open(mockup_path).convert("RGB")
    sx = img.width / BASE_W
    sy = img.height / BASE_H
    draw = ImageDraw.Draw(img)

    base_size = max(16, round(21 * sx))
    small_size = max(14, round(18 * sx))

    def xy(x, y):
        return (round(x * sx), round(y * sy))

    def write_value(text, x, y, max_width, size=base_size):
        text = display_value(text)
        font = fit_font(
            draw, text, round(max_width * sx),
            size=round(size * sx), minimum=12, bold=True
        )
        draw.text(
            xy(x, y), text, font=font,
            fill=dynamic_color(text), anchor="lm"
        )

    # Mockup sudah berisi logo, icon, label, garis, colon, card, dll.
    # Python hanya mengisi VALUE.

    site_rows = [
        value(row, "Site ID"),
        value(row, "Site Name"),
        value(row, "Regional"),
        value(row, "NOP_1"),
        value(row, "TO"),
        value(row, "ROH"),
        value(row, "Site Owner"),
        f"{value(row, 'Lat')} / {value(row, 'Long')}",
    ]
    for text, y in zip(site_rows, [236, 292, 347, 403, 459, 515, 570, 625]):
        write_value(text, 247, y, 215)

    rect_rows = [
        value(row, "ID PLN"),
        f"{value(row, 'Daya PLN (KVA)')} kVA",
        value(row, "Rectifier Brand (1)"),
        value(row, "Rectifier Model (1)"),
        value(row, "Module Capacity (1)"),
        value(row, "Inserted Module Qty (1)"),
        value(row, "Load System (1)"),
    ]
    for text, y in zip(rect_rows, [212, 257, 302, 346, 391, 436, 479]):
        write_value(text, 783, y, 287)

    bbt = safe_float(row["BBT H (1)"]) if "BBT H (1)" in row.index else None
    batt_rows = [
        value(row, "Battery Brand (1)"),
        value(row, "Battery Type (1)"),
        value(row, "Battery Capacity (1)"),
        value(row, "Battery Bank (1)"),
        f"{bbt:.2f} Hours" if bbt is not None else "-",
        value(row, "BBT Category (1)"),
    ]
    for text, y in zip(batt_rows, [609, 650, 691, 731, 771, 812]):
        write_value(text, 783, y, 287, size=small_size)

    utility = safe_float(row["Rectifier Utility"]) if "Rectifier Utility" in row.index else None
    health_rows = [
        value(row, "Rectifier Condition"),
        f"{utility * 100:.1f} %" if utility is not None else "-",
        value(row, "Rectifier Config (Category)"),
        value(row, "Capacity Status"),
        value(row, "Potensial Trip"),
        value(row, "Activity (SOW) Actual"),
        value(row, "EAS Validation"),
        value(row, "NETECO Status"),
    ]
    for text, y in zip(health_rows, [226, 281, 335, 390, 443, 497, 549, 604]):
        write_value(text, 1394, y, 225)

    # Action berada di area kosong setelah NETECO.
    actions = build_actions(row)
    for index, action in enumerate(actions[:3]):
        write_value(action, 1394, 658 + index * 36, 225, size=small_size)

    # Footer mockup
    update = value(row, "Last Check Update")
    footer_font = fit_font(
        draw, f"Data Update : {update}", 300,
        size=18, minimum=12, bold=True
    )
    draw.text(
        xy(72, 890),
        f"Data Update : {update}",
        font=footer_font,
        fill=(255, 255, 255),
        anchor="lm",
    )

    output_path = f"output_{clean_filename(site_id)}.png"
    img.save(output_path, format="PNG", dpi=(150, 150))
    return output_path

@bot.message_handler(commands=["site"])
def handle_site(message):
    args = message.text.split()
    if len(args) < 2:
        bot.reply_to(
            message,
            "⚠️ Format salah!\nGunakan: `/site <Site_ID>`",
            parse_mode="Markdown",
        )
        return

    site_id = args[1].strip()
    status_msg = bot.reply_to(
        message,
        f"⏳ Sedang memproses Site ID: *{site_id}*...",
        parse_mode="Markdown",
    )

    try:
        img_path = generate_site_card(site_id)

        if not img_path or not os.path.exists(img_path):
            bot.edit_message_text(
                f"❌ Site ID *{site_id}* tidak ditemukan.",
                message.chat.id,
                status_msg.message_id,
                parse_mode="Markdown",
            )
            return

        with open(img_path, "rb") as photo:
            bot.send_photo(
                message.chat.id,
                photo,
                caption=f"✅ Status Report Site ID: *{site_id}*",
                parse_mode="Markdown",
            )

        os.remove(img_path)

        try:
            bot.delete_message(message.chat.id, status_msg.message_id)
        except Exception:
            pass

    except Exception as exc:
        print(f"Error generate report {site_id}: {exc}")
        try:
            bot.edit_message_text(
                f"❌ Terjadi error saat membuat report untuk *{site_id}*.\n`{exc}`",
                message.chat.id,
                status_msg.message_id,
                parse_mode="Markdown",
            )
        except Exception:
            bot.reply_to(
                message,
                f"❌ Terjadi error saat membuat report untuk *{site_id}*.",
                parse_mode="Markdown",
            )

if __name__ == "__main__":
    print("Bot Telegram siap dijalankan...")
    bot.infinity_polling(skip_pending=True)
