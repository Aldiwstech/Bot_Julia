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
    # Power template: allow explicit override first.
    "power2.png",
]

# Power v5 uses the current 10-row mockup exactly.
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
    """Status color rules for the Power dashboard.

    Green: Normal, Safe, Monitor, Valid, Balance/Balanced, etc.
    Red: Warning, Unmonitor, Unavailable, Need Validation, faults, etc.
    Utility percentage is handled by utility_color() below.
    """
    text = str(value).strip().upper()

    # Explicit negative states first so MONITOR never overrides UNMONITOR.
    red_exact = {
        "UNMONITOR", "UNAVAILABLE", "NOT AVAILABLE", "NOT_AVAILABLE",
        "WARNING", "WARN", "NEED VALIDATION", "NEED CHECK",
        "UNBALANCE", "UNBALANCED", "CRITICAL", "DOWN", "TRIP",
        "FAULT", "FAILED", "ERROR", "NOT SAFE", "NOT AUTO",
        "BROKEN", "PROBLEM", "OFFLINE",
    }
    if text in red_exact:
        return RED

    if any(w in text for w in (
        "UNMONITOR", "UNAVAILABLE", "NOT AVAILABLE", "WARNING",
        "NEED VALIDATION", "NEED CHECK", "UNBALANCE", "UNBALANCED",
        "CRITICAL", "DOWN", "TRIP", "FAULT", "FAILED", "ERROR",
        "NOT SAFE", "NOT AUTO", "BROKEN", "PROBLEM", "OFFLINE",
    )):
        return RED

    if any(w in text for w in (
        "NORMAL", "SECURED", "SAFE", "VALID", "AVAILABLE",
        "MONITOR", "BALANCE", "BALANCED", "ACTIVE", "AUTO", "OK",
    )):
        return GREEN

    return NAVY


def utility_color(value):
    """0-60% green; above 60% red, per dashboard rule."""
    n = safe_float(value)
    if n is None:
        return NAVY
    # Excel may store either 0.549 or 54.9.
    pct = n * 100 if 0 <= n <= 1 else n
    return GREEN if 0 <= pct <= 60 else RED


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
    """Return Rect. Cond from the workbook's Rectifier Condition field.

    The current workbook exposes this as the named column "Rectifier Condition".
    If a future workbook moves the field and keeps the agreed physical ER
    position, ER is used as a fallback. This prevents silent column drift.
    """
    # ER adalah sumber utama sesuai posisi kolom yang sudah dikunci.
    er_value = excel_column_value(row, "ER")
    if er_value not in ("", "-"):
        return er_value
    # Fallback hanya bila ER kosong.
    return value(row, "Rectifier Condition")


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
    rect_cond = phase_condition(row).upper()
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

    def write_value(text, x, y, max_width, size=21, color=None, bold=True,
                    min_size=17, compress=True):
        """Draw an Excel value from one locked X anchor.

        Rules:
        - X never follows the label length.
        - X is the same for every row inside a panel.
        - Keep text visually large; if a value is long, compress it
          horizontally before reducing the font size too much.
        - Never add ellipsis and never wrap normal dashboard values.
        """
        text = display_value(text)
        if not text:
            return

        pixel_width = max(20, round(max_width * sx))
        font_size = max(min_size, round(size * min(sx, sy)))
        font = get_font(font_size, bold=bold)
        bbox = draw.textbbox((0, 0), text, font=font)
        text_w = bbox[2] - bbox[0]
        text_h = max(1, bbox[3] - bbox[1])

        # Prefer keeping the requested font height. Compress only the width.
        if compress and text_w > pixel_width:
            layer_w = text_w + 8
            layer_h = text_h + 8
            layer = Image.new("RGBA", (layer_w, layer_h), (0, 0, 0, 0))
            ld = ImageDraw.Draw(layer)
            ld.text((4 - bbox[0], 4 - bbox[1]), text, font=font,
                    fill=color if color is not None else dynamic_color(text))
            target_w = max(8, pixel_width)
            layer = layer.resize((target_w, layer_h), Image.Resampling.LANCZOS)
            # x is the LEFT edge of the value, exactly after the colon.
            img.paste(layer, (round(x * sx), round((y - text_h / 2 - 4) * sy)), layer)
            return

        # Only reduce font height as a last resort for extreme strings.
        if text_w > pixel_width:
            font = fit_font(draw, text, pixel_width, size=font_size,
                            minimum=min_size, bold=bold)

        draw.text(
            xy(x, y), text, font=font,
            fill=color if color is not None else dynamic_color(text),
            anchor="lm",
        )

    # =========================================================
    # LOCKED VALUE COLUMNS
    # Semua value dimulai pada X tetap setelah separator ":".
    # Jangan memakai X berbeda per label.
    # =========================================================
    # =========================================================
    # PRECISION ANCHORS — measured from the current 1672/941 mockup
    # =========================================================
    # Colon is the reference point. Values begin only ~14–16 px after it.
    # NEVER calculate these X positions from label length.
    SITE_X = 286
    SITE_W = 190

    MID_X = 800
    MID_W = 260

    HEALTH_X = 1400
    HEALTH_W = 215

    # -------------------------
    # SITE INFO
    # -------------------------
    # Template power saat ini memakai 8 row Site Info.
    # Jika template 10-row (Class Site + VIP) dipasang, kedua row
    # tambahan otomatis diisi tanpa menggeser row lain.
    site_rows = [
        value(row, "Site ID"),
        value(row, "Site Name"),
        value(row, "Regional"),
        value(row, "NOP_1"),
        value(row, "TO"),
        value(row, "ROH"),
        value(row, "Site Owner"),
        value(row, "Class Site "),
        value(row, "VIP"),
        f"{value(row, 'Lat')} / {value(row, 'Long')}",
    ]

    # Layout is selected explicitly. This avoids painting Class/VIP onto an
    # older 8-row template. Set POWER_SITE_ROWS=10 when using the new mockup.
    site_rows_mode = os.getenv("POWER_SITE_ROWS", "10").strip()
    if site_rows_mode == "10":
        site_y = [230, 286, 342, 398, 454, 510, 566, 622, 678, 734]
    else:
        site_rows = site_rows[:8]
        site_y = [230, 286, 343, 400, 457, 514, 571, 628]

    for i, (text, y) in enumerate(zip(site_rows, site_y)):
        if i == 9:
            size = 15
        elif i in (1, 7, 8):
            size = 18
        else:
            size = 19
        write_value(text, SITE_X, y, SITE_W, size=size, color=NAVY)

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
    for text, y in zip(rect_rows, [216, 263, 310, 357, 404, 451, 498]):
        write_value(text, MID_X, y, MID_W, size=18)

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
    for text, y in zip(batt_rows, [624, 669, 713, 758, 802, 846]):
        write_value(text, MID_X, y, MID_W, size=17)

    # -------------------------
    # HEALTHY CHECK
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
    health_y = [216, 264, 311, 359, 406, 454, 501, 548]
    for idx, (text, y) in enumerate(zip(health_values, health_y)):
        if idx == 1:
            write_value(text, HEALTH_X, y, HEALTH_W, size=16, color=utility_color(utility))
        else:
            write_value(text, HEALTH_X, y, HEALTH_W, size=16)

    # -------------------------
    # ACTION
    # -------------------------
    actions = build_actions(row)
    if not actions:
        actions = ["No action required"]
        action_color = GREEN
    else:
        action_color = RED

    action_x = round(1208 * sx)
    action_y = round(733 * sy)
    action_w = round(365 * sx)
    for action in actions[:4]:
        action_font = fit_font(
            draw, f"• {action}", action_w, size=17, minimum=12, bold=True
        )
        lines = wrap_text(draw, f"• {action}", action_font, action_w, max_lines=2)
        bbox = draw.textbbox((0, 0), "Ag", font=action_font)
        line_h = bbox[3] - bbox[1]
        for line in lines:
            draw.text((action_x, action_y), line, font=action_font, fill=action_color, anchor="la")
            action_y += line_h + 3
        action_y += 6
        if action_y > round(825 * sy):
            break

    # Intentionally no "Data Update / Last Check" footer text.

    output_path = f"output_precision_{clean_filename(site_id)}.png"
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
