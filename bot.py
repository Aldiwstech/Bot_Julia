import os
import re
import telebot
import pandas as pd
from PIL import Image, ImageDraw, ImageFont

TOKEN = os.getenv("TOKEN")
if not TOKEN:
    raise ValueError("Token belum diset di Environment Variables Railway!")

bot = telebot.TeleBot(TOKEN)

EXCEL_CANDIDATES = ["Excel_master(1).xlsx", "Excel_master.xlsx"]
MOCKUP_CANDIDATES = [
    # Final Power mockup: separate HEALTHY CHECK + ACTION panels.
    "wide_clean_infographic_dashboard_ui_mockup_on_a_l.png",
    "Mokup(2).png",
    "Mokup(1).png", "Mokup.png", "mokup.png", "mokup(1).png",
    "a_clean_flat_vector_infographic_dashboard_templat.png",
    "Blank Telkomsel Huawei Dashboard Template.png"
]

# Final mockup coordinate system: 1672 x 941.
BASE_W = 1683
BASE_H = 935

# Text colors are deliberately explicit; Action is always red.
NAVY = (24, 55, 105)
RED = (211, 42, 50)
GREEN = (20, 142, 68)
ORANGE = (232, 142, 18)
WHITE = (255, 255, 255)


def get_font(size, bold=False):
    candidates = []
    env_path = os.getenv("FONT_PATH")
    if env_path:
        candidates.append(env_path)

    candidates += (
        [
            # Liberation Sans is metrically close to Arial and matches the mockup
            # much better than DejaVu Sans for the data values.
            "/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
            "LiberationSans-Bold.ttf",
            "DejaVuSans-Bold.ttf",
        ]
        if bold else
        [
            "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "LiberationSans-Regular.ttf",
            "DejaVuSans.ttf",
        ]
    )

    for path in candidates:
        if path and os.path.exists(path):
            try:
                return ImageFont.truetype(path, size)
            except Exception:
                pass
    return ImageFont.load_default()


def find_existing(candidates):
    # Keep one source of truth. EXCEL_PATH can explicitly pin the workbook;
    # otherwise the uploaded Excel_master(1).xlsx is preferred when present.
    env_excel = os.getenv("EXCEL_PATH")
    if env_excel and os.path.exists(env_excel):
        return env_excel
    for path in candidates:
        if os.path.exists(path):
            return path
    return None


def resolve_sheet_name(excel_path, wanted):
    """Resolve worksheet names case-insensitively (e.g. Rectifire&Battery)."""
    with pd.ExcelFile(excel_path) as book:
        sheets = book.sheet_names
    target = str(wanted).strip().casefold()
    for sheet in sheets:
        if str(sheet).strip().casefold() == target:
            return sheet
    raise ValueError(f"Worksheet named '{wanted}' not found. Available: {', '.join(sheets)}")


def clean_filename(text):
    return re.sub(r"[^A-Za-z0-9._-]+", "_", str(text))


def is_empty(value):
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except Exception:
        pass
    return str(value).strip().lower() in ("", "nan", "none", "-", "nat")


def display_value(value, default="-"):
    return default if is_empty(value) else str(value).strip()


def safe_float(value):
    try:
        if pd.isna(value):
            return None
        return float(str(value).strip().replace(",", "."))
    except Exception:
        return None


def fit_font(draw, text, max_width, size=21, minimum=11, bold=True):
    text = str(text)
    current = int(size)
    minimum = int(minimum)
    while current >= minimum:
        font = get_font(current, bold=bold)
        bbox = draw.textbbox((0, 0), text, font=font)
        if bbox[2] - bbox[0] <= max_width:
            return font
        current -= 1
    return get_font(minimum, bold=bold)


def wrap_text(draw, text, font, max_width, max_lines=3):
    """Word-wrap text without allowing it to collide with the panel edge."""
    words = str(text).split()
    if not words:
        return []
    lines = []
    current = words[0]
    for word in words[1:]:
        candidate = current + " " + word
        if draw.textbbox((0, 0), candidate, font=font)[2] <= max_width:
            current = candidate
        else:
            lines.append(current)
            current = word
    lines.append(current)
    if len(lines) <= max_lines:
        return lines
    # Keep the first max_lines-1 lines and fit the remainder with an ellipsis.
    kept = lines[:max_lines-1]
    remainder = " ".join(lines[max_lines-1:])
    while remainder and draw.textbbox((0, 0), remainder + "...", font=font)[2] > max_width:
        remainder = remainder[:-1].rstrip()
    kept.append((remainder + "...") if remainder else "..." )
    return kept


def draw_wrapped(draw, text, x, y, max_width, font, color, line_gap=4, anchor="la"):
    lines = wrap_text(draw, text, font, max_width, max_lines=3)
    bbox = draw.textbbox((0, 0), "Ag", font=font)
    line_h = bbox[3] - bbox[1]
    for i, line in enumerate(lines):
        draw.text((x, y + i * (line_h + line_gap)), line, font=font, fill=color, anchor=anchor)
    return len(lines)


def dynamic_color(value):
    """Status color with negative states checked before positive substrings."""
    text = str(value).strip().upper()
    if text in ("UNMONITOR", "NOT AVAILABLE", "NOT_AVAILABLE"):
        return RED
    if any(w in text for w in (
        "UNBALANCE", "UNBALANCED", "NEED VALIDATION", "NEED CHECK",
        "CRITICAL", "DOWN", "TRIP", "FAULT", "FAILED", "ERROR",
        "NOT SAFE", "NOT AUTO", "BROKEN", "PROBLEM", "OFFLINE"
    )):
        return RED
    if any(w in text for w in ("POTENSIAL", "POTENTIAL", "WARNING", "CHECK")):
        return ORANGE
    if any(w in text for w in (
        "NORMAL", "SECURED", "OK", "VALID", "AVAILABLE", "MONITOR",
        "SAFE", "BALANCE", "BALANCED", "ACTIVE", "AUTO", "CLOSED"
    )):
        return GREEN
    return NAVY


def load_dataframe():
    excel_path = find_existing(EXCEL_CANDIDATES)
    if not excel_path:
        raise FileNotFoundError("Excel_master.xlsx tidak ditemukan.")

    sheet = resolve_sheet_name(excel_path, "Rectifire&Battery")
    df = pd.read_excel(excel_path, sheet_name=sheet)
    if "Site ID" not in df.columns:
        raise KeyError("Kolom 'Site ID' tidak ditemukan.")
    return df


def load_site(site_id):
    df = load_dataframe()
    wanted = str(site_id).strip().upper()
    ids = df["Site ID"].astype(str).str.strip().str.upper()
    result = df.loc[ids == wanted]
    return None if result.empty else result.iloc[0]


def value(row, column, default="-"):
    if column not in row.index:
        return default
    return display_value(row[column], default)


def excel_column_value(row, column_letter, default="-"):
    """Read a value by its physical Excel column letter.

    This is intentional for fields whose source position is fixed in the
    workbook and should not change when the header/order changes.
    """
    try:
        import openpyxl.utils
        idx = openpyxl.utils.column_index_from_string(column_letter) - 1
        if idx < 0 or idx >= len(row):
            return default
        return display_value(row.iloc[idx], default)
    except Exception:
        return default


def format_voltage(value_):
    n = safe_float(value_)
    if n is None:
        return "-"
    if abs(n - round(n)) < 1e-9:
        return f"{int(round(n))} V"
    return f"{n:.1f} V"


def phase_condition(row):
    """Rect. Cond is sourced from physical Excel column BL."""
    return excel_column_value(row, "BL")


def rect_status(row):
    """Direct source field for the Rect. Status row."""
    return value(row, "RECT1_Status")


def build_actions(row):
    """
    Action is a generated technical instruction, not an Excel field.
    Multiple independent issues are retained.
    """
    actions = []

    # 1) PLN phase imbalance -> inspect individual phase voltage.
    phase = value(row, "Phase Balanced").upper()
    if phase in ("UNBALANCE", "UNBALANCED"):
        voltages = {
            "R": safe_float(row["Voltage R"]) if "Voltage R" in row.index else None,
            "S": safe_float(row["Voltage S"]) if "Voltage S" in row.index else None,
            "T": safe_float(row["Voltage T"]) if "Voltage T" in row.index else None,
        }
        zero_phases = [p for p, v in voltages.items() if v is not None and abs(v) < 1e-9]

        if zero_phases:
            phases = " & ".join(zero_phases)
            actions.append(f"Check PLN Phase {phases} - Voltage 0V")
        else:
            actions.append("Check PLN Phase Balance / Voltage")

    # 2) EAS.
    eas = value(row, "EAS Validation").upper()
    if eas == "NEED VALIDATION":
        actions.append("Need Check Onsite")

    # 3) NetEco. Only the agreed UnMonitor rule is applied here.
    neteco = value(row, "NETECO Status").upper()
    if neteco == "UNMONITOR":
        actions.append("Need Check Onsite Connection NetEco")

    # 4) Rectifier.
    rect_cond = excel_column_value(row, "BL").upper()
    rect1 = value(row, "RECT1_Status").upper()
    rect_config = value(row, "Rectifier Config (Category)").upper()

    if "CRITICAL" in rect_cond:
        actions.append("Need Check / Upgrade Rectifier")
    elif "NEED CHECK" in rect1 or "NEED CHECK" in rect_config:
        actions.append("Need Check Onsite Rectifier")

    # De-duplicate while preserving rule order.
    unique = []
    for action in actions:
        if action not in unique:
            unique.append(action)

    return unique


def generate_site_card(site_id):
    row = load_site(site_id)
    if row is None:
        return None

    mockup_path = find_existing(MOCKUP_CANDIDATES)
    if not mockup_path:
        raise FileNotFoundError("Mockup template tidak ditemukan.")

    img = Image.open(mockup_path).convert("RGB")
    sx = img.width / BASE_W
    sy = img.height / BASE_H
    draw = ImageDraw.Draw(img)

    def xy(x, y):
        return round(x * sx), round(y * sy)

    def write_value(text, x, y, max_width, size=21, color=None, bold=True):
        text = display_value(text)
        font = fit_font(
            draw, text, round(max_width * sx),
            size=round(size * min(sx, sy)), minimum=10, bold=bold
        )
        draw.text(
            xy(x, y), text, font=font,
            fill=color if color is not None else dynamic_color(text),
            anchor="lm",
        )

    # -------------------------
    # SITE INFO
    # -------------------------
    site_rows = [
        value(row, "Site ID"),
        value(row, "Site Name"),
        value(row, "Regional"),
        value(row, "NOP"),
        value(row, "TO"),
        value(row, "ROH"),
        value(row, "Site Owner"),
        f"{value(row, 'Lat')} / {value(row, 'Long')}",
    ]
    for i, (text, y) in enumerate(zip(site_rows, [231, 290, 349, 409, 468, 526, 610, 676])):
        site_size = 20
        site_width = 202
        if i == 7:  # Lat / Long is intentionally one point smaller.
            site_size = 17
        write_value(text, 270, y, site_width, size=site_size)

    # -------------------------
    # RECTIFIER & PLN
    # -------------------------
    rect_rows = [
        value(row, "ID PLN"),
        f"{value(row, 'Daya PLN (KVA)')} kVA",
        value(row, "Rectifier Brand (1)"),
        value(row, "Rectifier Model (1)"),
        value(row, "Module Capacity (1)"),
        value(row, "Inserted Module Qty (1)"),
        value(row, "Load System (1)"),
    ]
    for text, y in zip(rect_rows, [211, 257, 302, 348, 394, 440, 486]):
        write_value(text, 800, y, 270, size=19)

    # -------------------------
    # BATTERY STATUS
    # -------------------------
    bbt = safe_float(row["BBT H (1)"]) if "BBT H (1)" in row.index else None
    batt_rows = [
        value(row, "Battery Brand (1)"),
        value(row, "Battery Type (1)"),
        value(row, "Battery Capacity (1)"),
        value(row, "Battery Bank (1)"),
        f"{bbt:.2f} Hours" if bbt is not None else "-",
        value(row, "BBT Category (1)"),
    ]
    for text, y in zip(batt_rows, [618, 661, 704, 746, 789, 830]):
        write_value(text, 800, y, 270, size=18)

    # -------------------------
    # HEALTHY CHECK & ACTION
    # 9 rows in the final mockup:
    # Rect. Cond / Utility / Config / Cap. Status / PLN Voltage /
    # EAS Valid. / NETECO Stat / Rect. Status / Action
    # -------------------------
    utility = safe_float(row["Rectifier Utility"]) if "Rectifier Utility" in row.index else None
    utility_text = f"{utility * 100:.1f} %" if utility is not None else "-"

    voltage_text = " / ".join([
        f"R {format_voltage(row['Voltage R'])}" if "Voltage R" in row.index else "R -",
        f"S {format_voltage(row['Voltage S'])}" if "Voltage S" in row.index else "S -",
        f"T {format_voltage(row['Voltage T'])}" if "Voltage T" in row.index else "T -",
    ])

    health_values = [
        phase_condition(row),
        utility_text,
        value(row, "Rectifier Config (Category)"),
        value(row, "Capacity Status"),
        voltage_text,
        value(row, "EAS Validation"),
        value(row, "NETECO Status"),
        rect_status(row),
    ]

    # The mockup's 9th row is reserved for generated Action.
    health_y = [210, 257, 304, 351, 398, 445, 492, 539]
    for text, y in zip(health_values, health_y):
        write_value(text, 1429, y, 200, size=16)

    # ACTION is a dedicated panel in the final mockup. The template already
    # provides the panel, so only the generated instructions are drawn here.
    actions = build_actions(row)
    action_x = round(1175 * sx)
    action_y = round(704 * sy)
    action_w = round(410 * sx)
    action_font = fit_font(draw, "Need Check Onsite Connection NetEco", action_w,
                           size=18, minimum=13, bold=True)
    if not actions:
        actions = ["No action required"]
        action_color = GREEN
    else:
        action_color = RED

    for action in actions[:3]:
        lines = wrap_text(draw, f"• {action}", action_font, action_w, max_lines=2)
        bbox = draw.textbbox((0, 0), "Ag", font=action_font)
        line_h = bbox[3] - bbox[1]
        for line in lines:
            draw.text((action_x, action_y), line, font=action_font, fill=action_color, anchor="la")
            action_y += line_h + 4
        action_y += 8
        if action_y > round(825 * sy):
            break

    # Intentionally no "Data Update / Last Check" footer text.

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


from genset_module import register_genset_handler
register_genset_handler(bot)

print("Bot Telegram siap dijalankan...")
print("Commands aktif: /site <Site_ID> dan /genset <Site_ID>")
bot.infinity_polling(skip_pending=True)
