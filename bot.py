import os
import re
import math
import telebot
import pandas as pd
from PIL import Image, ImageDraw, ImageFont

TOKEN = os.getenv("TOKEN")
if not TOKEN:
    raise ValueError("Token belum diset di Environment Variables Railway!")

bot = telebot.TeleBot(TOKEN)


def load_font(size):
    """Load font with safe fallbacks."""
    for path in [
        "DejaVuSans.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "arial.ttf",
        "Arial.ttf",
    ]:
        try:
            return ImageFont.truetype(path, size)
        except (OSError, IOError):
            continue
    return ImageFont.load_default()


def safe_float(value):
    """Convert Excel numeric values safely, including decimal comma."""
    if value is None or pd.isna(value):
        return None
    text = str(value).strip().replace(",", ".")
    try:
        return float(text)
    except (ValueError, TypeError):
        return None


def generate_site_card(site_id):
    excel_path = "Excel_master.xlsx"
    mockup_path = "Mokup.png"

    if not os.path.exists(excel_path) or not os.path.exists(mockup_path):
        return None

    requested_id = str(site_id).strip().upper()
    site_data = None

    # Cari Site ID di seluruh sheet Excel.
    xls = pd.ExcelFile(excel_path)
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name)
        id_col = None

        for col in df.columns:
            col_text = str(col).lower()
            if "site" in col_text and "id" in col_text:
                id_col = col
                break

        if id_col is None:
            continue

        normalized_ids = df[id_col].astype(str).str.strip().str.upper()
        filtered = df[normalized_ids == requested_id]

        if not filtered.empty:
            site_data = filtered.iloc[0]
            break

    if site_data is None:
        return None

    img = Image.open(mockup_path).convert("RGB")
    draw = ImageDraw.Draw(img)
    W, H = img.size

    # Ukuran font dibuat relatif terhadap ukuran mockup.
    # Untuk mockup sekitar 4200 px, hasilnya sekitar 60-65 px.
    base_font_size = max(24, round(W * 0.0155))
    font = load_font(base_font_size)
    icon_radius = max(5, round(base_font_size * 0.13))

    COLOR_LABEL = (80, 80, 80)
    COLOR_TEXT = (20, 20, 20)
    COLOR_GREEN = (16, 124, 65)
    COLOR_BLUE = (0, 120, 212)
    COLOR_RED = (209, 52, 56)

    def val(col, default="-"):
        if col not in site_data:
            return default
        value = site_data.get(col)
        if pd.isna(value):
            return default
        text = str(value).strip()
        return text if text and text.lower() != "nan" else default

    def get_dynamic_color(text):
        t = str(text).upper()
        if any(w in t for w in [
            "DOWN", "CRITICAL", "NO BACKUP", "NOT AVAILABLE",
            "NEED VALIDATION", "NOT_AVAILABLE", "POTENSIAL TRIP",
            "TRIP", "NEED", "PERGANTIAN"
        ]):
            return COLOR_RED
        if any(w in t for w in [
            "NORMAL", "SECURED", "VALID", "OK", "AVAILABLE",
            "VIP", "MONITOR"
        ]):
            return COLOR_GREEN
        if any(w in t for w in [
            "SILVER", "GOLD", "LITHIUM", "HUAWEI", "TELKOM"
        ]):
            return COLOR_BLUE
        return COLOR_TEXT

    # ------------------------------------------------------------------
    # ACTION
    # ------------------------------------------------------------------
    action_list = []
    for col_name in ["Action", "Activity (SOW) Actual", "SOW"]:
        if col_name in site_data:
            excel_action = site_data.get(col_name)
            if pd.notna(excel_action):
                action = str(excel_action).strip()
                if action and action not in ["-", "nan"]:
                    action_list.append(action)
                    break

    if not action_list:
        rect_cond = val("Rectifier Condition").upper()
        pot_trip = val("Potensial Trip").upper()
        eas_val = val("EAS Validation").upper()
        neteco = val("NETECO Status").upper()

        if "DOWN" in rect_cond or "CRITICAL" in rect_cond:
            action_list.append("NEED REPLACE RECTIFIER")
        if "TRIP" in pot_trip or "POTENSIAL" in pot_trip:
            action_list.append("NEED CHECK LOAD")
        if "NEED" in eas_val or "VALIDATION" in eas_val:
            action_list.append("NEED VALIDATE")
        if "NOT AVAILABLE" in neteco or "DOWN" in neteco:
            action_list.append("NEED CHECK NETECO")

        if not action_list:
            action_list = ["NORMAL"]

    # ------------------------------------------------------------------
    # DATA
    # ------------------------------------------------------------------
    col_site = [
        ("Site ID", val("Site ID")),
        ("Site Name", val("Site Name")),
        ("Regional", val("Regional")),
        ("NOP", val("NOP_1")),
        ("TO", val("TO")),
        ("ROH", val("ROH")),
        ("Site Owner", val("Site Owner")),
        ("Lat / Long", f"{val('Lat')} / {val('Long')}"),
    ]

    col_rect = [
        ("ID PLN", val("ID PLN")),
        ("Daya PLN", f"{val('Daya PLN (KVA)')} kVA"),
        ("Brand", val("Rectifier Brand (1)")),
        ("Model", val("Rectifier Model (1)")),
        ("Capacity", val("Module Capacity (1)")),
        ("Module Qty", val("Inserted Module Qty (1)")),
        ("Load System", val("Load System (1)")),
    ]

    bbt_value = safe_float(site_data.get("BBT H (1)")) if "BBT H (1)" in site_data else None
    bbt_str = f"{bbt_value:.2f} Hours" if bbt_value is not None else "-"

    col_batt = [
        ("Brand", val("Battery Brand (1)")),
        ("Type", val("Battery Type (1)")),
        ("Capacity", val("Battery Capacity (1)")),
        ("Bank Qty", val("Battery Bank (1)")),
        ("BBT Backup", bbt_str),
        ("Category", val("BBT Category (1)")),
    ]

    utility = safe_float(site_data.get("Rectifier Utility")) if "Rectifier Utility" in site_data else None
    util_str = f"{utility * 100:.1f} %" if utility is not None else "-"

    col_health = [
        ("Rect. Cond", val("Rectifier Condition")),
        ("Utility", util_str),
        ("Config", val("Rectifier Config (Category)")),
        ("Cap. Status", val("Capacity Status")),
        ("Pot. Trip", val("Potensial Trip")),
        ("SOW Act.", val("Activity (SOW) Actual")),
        ("EAS Valid.", val("EAS Validation")),
        ("NETECO Stat", val("NETECO Status")),
    ]

    # ------------------------------------------------------------------
    # LAYOUT
    # ------------------------------------------------------------------
    # Semua posisi dibuat dalam persentase ukuran gambar agar tetap pas
    # walaupun resolusi Mokup.png berubah.
    configs = {
        "site": {
            "icon_x": 0.027,
            "label_x": 0.049,
            "colon_x": 0.119,
            "value_x": 0.129,
            "start_y": 0.250,
            "spacing": 0.041,
            "value_right": 0.255,
        },
        "rect": {
            "icon_x": 0.323,
            "label_x": 0.346,
            "colon_x": 0.444,
            "value_x": 0.461,
            "start_y": 0.250,
            "spacing": 0.041,
            "value_right": 0.603,
        },
        "batt": {
            "icon_x": 0.323,
            "label_x": 0.346,
            "colon_x": 0.444,
            "value_x": 0.461,
            "start_y": 0.642,
            "spacing": 0.041,
            "value_right": 0.603,
        },
        "health": {
            "icon_x": 0.612,
            "label_x": 0.634,
            "colon_x": 0.729,
            "value_x": 0.747,
            "start_y": 0.250,
            "spacing": 0.041,
            "value_right": 0.950,
        },
    }

    def fit_font(text, max_width, start_size=None):
        """Shrink value font only when necessary, preserving readability."""
        size = start_size or base_font_size
        text = str(text)
        while size > 18:
            candidate = load_font(size)
            bbox = draw.textbbox((0, 0), text, font=candidate)
            if bbox[2] - bbox[0] <= max_width:
                return candidate
            size -= 2
        return load_font(size)

    def draw_icon(cx, cy):
        # Jangan gunakan karakter '•' karena pada beberapa environment
        # Telegram/Railway/font fallback dapat berubah menjadi kotak/simbol aneh.
        r = icon_radius
        draw.ellipse((cx - r, cy - r, cx + r, cy + r), fill=COLOR_TEXT)

    def draw_row(cfg, row_index, label, value):
        x_icon = round(W * cfg["icon_x"])
        x_label = round(W * cfg["label_x"])
        x_colon = round(W * cfg["colon_x"])
        x_value = round(W * cfg["value_x"])
        y = round(H * (cfg["start_y"] + row_index * cfg["spacing"]))

        # Gunakan anchor 'lm' agar semua elemen berada pada garis tengah yang sama.
        draw_icon(x_icon, y)
        draw.text((x_label, y), str(label), fill=COLOR_LABEL, font=font, anchor="lm")
        draw.text((x_colon, y), ":", fill=COLOR_TEXT, font=font, anchor="mm")

        max_width = max(100, round(W * cfg["value_right"]) - x_value)
        value_font = fit_font(value, max_width)
        draw.text(
            (x_value, y),
            str(value),
            fill=get_dynamic_color(value),
            font=value_font,
            anchor="lm",
        )

    def render_box(cfg, items):
        for i, (label, value) in enumerate(items):
            draw_row(cfg, i, label, value)

    render_box(configs["site"], col_site)
    render_box(configs["rect"], col_rect)
    render_box(configs["batt"], col_batt)
    render_box(configs["health"], col_health)

    # ------------------------------------------------------------------
    # ACTION - posisinya mengikuti baris terakhir Health Check
    # ------------------------------------------------------------------
    health_cfg = configs["health"]
    action_y_ratio = health_cfg["start_y"] + len(col_health) * health_cfg["spacing"] + 0.018
    action_y = round(H * action_y_ratio)

    draw_icon(round(W * health_cfg["icon_x"]), action_y)
    draw.text(
        (round(W * health_cfg["label_x"]), action_y),
        "Action",
        fill=COLOR_LABEL,
        font=font,
        anchor="lm",
    )
    draw.text(
        (round(W * health_cfg["colon_x"]), action_y),
        ":",
        fill=COLOR_TEXT,
        font=font,
        anchor="mm",
    )

    x_value = round(W * health_cfg["value_x"])
    max_width = max(100, round(W * health_cfg["value_right"]) - x_value)
    action_spacing = round(H * 0.041)

    for index, action_item in enumerate(action_list):
        action_font = fit_font(action_item, max_width)
        draw.text(
            (x_value, action_y + index * action_spacing),
            action_item,
            fill=get_dynamic_color(action_item),
            font=action_font,
            anchor="lm",
        )

    # Nama file dibuat aman untuk Windows/Linux/Railway.
    safe_site_id = re.sub(r"[^A-Za-z0-9_.-]+", "_", str(site_id).strip())
    output_path = f"output_{safe_site_id}.png"
    img.save(output_path, format="PNG", dpi=(300, 300))
    return output_path


@bot.message_handler(commands=["site"])
def handle_site(message):
    args = message.text.split()
    if len(args) < 2:
        bot.reply_to(
            message,
            "⚠️ Format salah! Gunakan: `/site <Site_ID>`,",
            parse_mode="Markdown",
        )
        return

    site_id = args[1]
    bot.reply_to(
        message,
        f"⏳ Sedang memproses Site ID: *{site_id}*...",
        parse_mode="Markdown",
    )

    try:
        img_path = generate_site_card(site_id)

        if img_path and os.path.exists(img_path):
            with open(img_path, "rb") as doc_file:
                bot.send_document(
                    message.chat.id,
                    doc_file,
                    caption=f"✅ Status Report Site ID: *{site_id}*",
                    parse_mode="Markdown",
                )
            os.remove(img_path)
        else:
            bot.reply_to(
                message,
                f"❌ Maaf, Site ID *{site_id}* tidak ditemukan.",
                parse_mode="Markdown",
            )
    except Exception as exc:
        print(f"ERROR generate_site_card({site_id}): {exc}")
        bot.reply_to(
            message,
            "❌ Terjadi error saat membuat report. Silakan cek log Railway.",
        )


print("Bot Telegram siap dijalankan...")
bot.infinity_polling()
