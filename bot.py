import os
import re
import pandas as pd
from PIL import Image, ImageDraw, ImageFont

# ============================================================
# POWER MODULE - PRECISION V2
# ============================================================
# /site only. Genset and Area are NOT touched by this module.
#
# Source:
#   Excel_master(1).xlsx / Excel_master.xlsx
#   sheet: Rectifire&Battery
#
# Mockup:
#   Mokup2.png  <-- fixed Power mockup, do not replace automatically
#
# Main fixes:
#   1. All value X positions are locked after the existing ':'
#   2. No wrapping for normal Excel values
#   3. NO artificial "..." truncation
#   4. Long values are fitted by reducing font slightly; if still too long,
#      text is horizontally compressed so the full Excel value remains visible
#   5. Y positions are fixed per row
#   6. Rectifier Condition uses the named Excel header, NOT physical ER
#      (in the current workbook ER is "Battery Type (6)")
# ============================================================

EXCEL_CANDIDATES = [
    "Excel_master(1).xlsx",
    "Excel_master.xlsx",
]

MOCKUP_CANDIDATES = [
    "Mokup2.png",
]

BASE_W = 1683
BASE_H = 935

NAVY = (24, 55, 105)
RED = (211, 42, 50)
GREEN = (20, 142, 68)
ORANGE = (232, 142, 18)
WHITE = (255, 255, 255)


# ------------------------------------------------------------
# File / Excel helpers
# ------------------------------------------------------------

def find_existing(candidates):
    for path in candidates:
        if os.path.exists(path):
            return path
    return None


def resolve_sheet_name(excel_path, wanted):
    with pd.ExcelFile(excel_path) as book:
        sheets = book.sheet_names

    target = str(wanted).strip().casefold()
    for sheet in sheets:
        if str(sheet).strip().casefold() == target:
            return sheet

    raise ValueError(
        f"Worksheet '{wanted}' tidak ditemukan. Available: {', '.join(sheets)}"
    )


def is_empty(value):
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except Exception:
        pass
    return str(value).strip().lower() in ("", "nan", "none", "nat", "-")


def clean_text(value, default="-"):
    if is_empty(value):
        return default

    text = str(value)
    text = text.replace("\r\n", " ").replace("\n", " ").replace("\r", " ")
    text = text.replace("\t", " ")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def value(row, column, default="-"):
    if column not in row.index:
        return default
    return clean_text(row[column], default)


def safe_float(value_):
    try:
        if pd.isna(value_):
            return None
        return float(str(value_).strip().replace(",", "."))
    except Exception:
        return None


def load_dataframe():
    excel_path = find_existing(EXCEL_CANDIDATES)
    if not excel_path:
        raise FileNotFoundError(
            "Excel_master(1).xlsx / Excel_master.xlsx tidak ditemukan."
        )

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


# ------------------------------------------------------------
# Font / rendering helpers
# ------------------------------------------------------------

def get_font(size, bold=True, narrow=False):
    env_path = os.getenv("FONT_PATH")
    candidates = []
    if env_path:
        candidates.append(env_path)

    if narrow:
        candidates += (
            [
                "/usr/share/fonts/truetype/liberation/LiberationSansNarrow-Bold.ttf",
                "/usr/share/fonts/truetype/dejavu/DejaVuSansCondensed-Bold.ttf",
            ]
            if bold
            else [
                "/usr/share/fonts/truetype/liberation/LiberationSansNarrow-Regular.ttf",
                "/usr/share/fonts/truetype/dejavu/DejaVuSansCondensed.ttf",
            ]
        )
    else:
        candidates += (
            [
                "/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf",
                "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
                "LiberationSans-Bold.ttf",
                "DejaVuSans-Bold.ttf",
            ]
            if bold
            else [
                "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
                "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
                "LiberationSans-Regular.ttf",
                "DejaVuSans.ttf",
            ]
        )

    for path in candidates:
        if path and os.path.exists(path):
            try:
                return ImageFont.truetype(path, int(size))
            except Exception:
                pass

    return ImageFont.load_default()


def text_width(draw, text, font):
    bbox = draw.textbbox((0, 0), str(text), font=font)
    return bbox[2] - bbox[0]


def fit_font(draw, text, max_width, start_size, min_size):
    """Fit without truncating: normal font first, then narrow font."""
    for size in range(int(start_size), int(min_size) - 1, -1):
        font = get_font(size, bold=True, narrow=False)
        if text_width(draw, text, font) <= max_width:
            return font, False

    # Condensed font preserves visual height better than making text tiny.
    for size in range(int(start_size), int(min_size) - 1, -1):
        font = get_font(size, bold=True, narrow=True)
        if text_width(draw, text, font) <= max_width:
            return font, True

    return get_font(min_size, bold=True, narrow=True), True


def draw_text_fit(draw, text, x, y, max_width, start_size=19, min_size=15, color=NAVY, anchor="lm"):
    """Draw one complete value on one line.

    If it cannot fit at min_size, the text bitmap is horizontally compressed.
    No ellipsis and no wrapping.
    """
    text = clean_text(text, "")
    if not text:
        return

    font, _ = fit_font(draw, text, max_width, start_size, min_size)
    width = text_width(draw, text, font)

    if width <= max_width:
        draw.text((x, y), text, font=font, fill=color, anchor=anchor)
        return

    # Render text on a transparent strip, then compress only its WIDTH.
    bbox = draw.textbbox((0, 0), text, font=font)
    pad_x = 4
    pad_y = 4
    src_w = max(1, bbox[2] - bbox[0] + pad_x * 2)
    src_h = max(1, bbox[3] - bbox[1] + pad_y * 2)

    layer = Image.new("RGBA", (src_w, src_h), (255, 255, 255, 0))
    ld = ImageDraw.Draw(layer)
    ld.text(
        (pad_x - bbox[0], pad_y - bbox[1]),
        text,
        font=font,
        fill=color + (255,),
    )

    target_w = max(1, int(max_width))
    layer = layer.resize((target_w, src_h), Image.Resampling.LANCZOS)

    # Convert anchor position into a top-left placement.
    if anchor == "lm":
        px = int(x)
        py = int(y - src_h / 2)
    elif anchor == "la":
        px = int(x)
        py = int(y)
    else:
        px = int(x)
        py = int(y)

    # Clamp to the intended value area.
    if px < 0:
        px = 0
    if px + target_w > draw.im.size[0]:
        px = max(0, draw.im.size[0] - target_w)

    draw._image.paste(layer, (px, py), layer)


# ------------------------------------------------------------
# Status / business rules
# ------------------------------------------------------------

def dynamic_color(text):
    t = clean_text(text, "").upper()

    red_words = (
        "UNMONITOR", "UNAVAILABLE", "NOT AVAILABLE", "WARNING",
        "NEED VALIDATION", "NEED CHECK", "UNBALANCE", "UNBALANCED",
        "CRITICAL", "DOWN", "TRIP", "FAULT", "FAILED", "ERROR",
        "NOT SAFE", "NOT AUTO", "BROKEN", "PROBLEM", "OFFLINE",
    )
    if any(word in t for word in red_words):
        return RED

    green_words = (
        "NORMAL", "SECURED", "SAFE", "VALID", "AVAILABLE", "MONITOR",
        "BALANCE", "BALANCED", "ACTIVE", "AUTO", "OK",
    )
    if any(word in t for word in green_words):
        return GREEN

    return NAVY


def utility_color(value_):
    n = safe_float(value_)
    if n is None:
        return NAVY
    pct = n * 100 if 0 <= n <= 1 else n
    return GREEN if 0 <= pct <= 60 else RED


def format_voltage(value_):
    n = safe_float(value_)
    if n is None:
        return "-"
    if abs(n - round(n)) < 1e-9:
        return f"{int(round(n))} V"
    return f"{n:.1f} V"


def rectifier_condition(row):
    """Use the actual named Rectifier Condition field.

    DO NOT read ER here. In the current Excel_master workbook, ER is
    'Battery Type (6)', so using ER would put the wrong value in Rect. Cond.
    """
    if "Rectifier Condition" in row.index:
        return value(row, "Rectifier Condition")

    # Safe fallback: locate a header containing both words.
    for col in row.index:
        normalized = re.sub(r"\s+", " ", str(col).strip()).casefold()
        if normalized == "rectifier condition":
            return value(row, col)

    return "-"


def rect_status(row):
    return value(row, "RECT1_Status")


def build_actions(row):
    actions = []

    phase = value(row, "Phase Balanced").upper()
    if phase in ("UNBALANCE", "UNBALANCED"):
        voltages = {
            "R": safe_float(row["Voltage R"]) if "Voltage R" in row.index else None,
            "S": safe_float(row["Voltage S"]) if "Voltage S" in row.index else None,
            "T": safe_float(row["Voltage T"]) if "Voltage T" in row.index else None,
        }
        zero = [p for p, v in voltages.items() if v is not None and abs(v) < 1e-9]
        if zero:
            actions.append(f"Check PLN Phase {' & '.join(zero)} - Voltage 0V")
        else:
            actions.append("Check PLN Phase Balance / Voltage")

    eas = value(row, "EAS Validation").upper()
    if eas == "NEED VALIDATION":
        actions.append("Need Check Onsite EAS")

    neteco = value(row, "NETECO Status").upper()
    if neteco == "UNMONITOR":
        actions.append("Need Check Onsite Connection NetEco")

    rect_cond = rectifier_condition(row).upper()
    rect1 = value(row, "RECT1_Status").upper()
    rect_config = value(row, "Rectifier Config (Category)").upper()

    if any(word in rect_cond for word in ("CRITICAL", "BROKEN", "FAULT", "FAILED", "ERROR")):
        actions.append("Need Check / Upgrade Rectifier")
    elif "NEED CHECK" in rect1 or "NEED CHECK" in rect_config:
        actions.append("Need Check Onsite Rectifier")

    unique = []
    for action in actions:
        if action not in unique:
            unique.append(action)
    return unique


# ------------------------------------------------------------
# Renderer
# ------------------------------------------------------------

def generate_site_card(site_id):
    row = load_site(site_id)
    if row is None:
        return None

    mockup_path = find_existing(MOCKUP_CANDIDATES)
    if not mockup_path:
        raise FileNotFoundError(
            "Mokup2.png tidak ditemukan. Taruh mockup Power di folder yang sama dengan bot."
        )

    img = Image.open(mockup_path).convert("RGB")
    sx = img.width / BASE_W
    sy = img.height / BASE_H
    draw = ImageDraw.Draw(img)

    def xy(x, y):
        return round(x * sx), round(y * sy)

    # ---------------------------------------------------------
    # LOCKED X/Y COORDINATES
    # ---------------------------------------------------------
    # These coordinates are measured from the actual Mokup2.png.
    # VALUE X is the left edge immediately after the printed ':' — never
    # calculated from label length.
    #
    # Site Info colon ~= x 269 -> value starts x 283.
    # Battery/Rectifier colon ~= x 778 -> value starts x 792.
    SITE_X = 283
    SITE_W = 182

    MID_X = 792
    MID_W = 251

    HEALTH_X = 1393
    HEALTH_W = 225

    # -------------------------
    # SITE INFO - 10 rows
    # -------------------------
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
    site_y = [230, 286, 342, 398, 454, 510, 566, 622, 678, 734]

    for i, (text, y) in enumerate(zip(site_rows, site_y)):
        if i == 9:
            # Long coordinate value: slightly smaller, but still complete.
            draw_text_fit(
                draw, text, *xy(SITE_X, y),
                round(SITE_W * sx), start_size=16, min_size=14,
                color=NAVY, anchor="lm"
            )
        else:
            draw_text_fit(
                draw, text, *xy(SITE_X, y),
                round(SITE_W * sx), start_size=20, min_size=16,
                color=NAVY, anchor="lm"
            )

    # -------------------------
    # RECTIFIER & PLN - 7 rows
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
    rect_y = [216, 263, 310, 357, 404, 451, 498]

    for text, y in zip(rect_rows, rect_y):
        draw_text_fit(
            draw, text, *xy(MID_X, y),
            round(MID_W * sx), start_size=19, min_size=16,
            color=NAVY, anchor="lm"
        )

    # -------------------------
    # BATTERY STATUS - 6 rows
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
    # Battery values in Mokup2.png are vertically centered against
    # their labels/colons. The previous y values put values visibly too low.
    batt_y = [591, 635, 679, 723, 767, 811]

    for text, y in zip(batt_rows, batt_y):
        draw_text_fit(
            draw, text, *xy(MID_X, y),
            round(MID_W * sx), start_size=19, min_size=16,
            color=NAVY, anchor="lm"
        )

    # -------------------------
    # HEALTHY CHECK - 8 rows
    # -------------------------
    utility = safe_float(row["Rectifier Utility"]) if "Rectifier Utility" in row.index else None
    utility_text = f"{utility * 100:.1f} %" if utility is not None else "-"

    voltage_text = " / ".join([
        f"R {format_voltage(row['Voltage R'])}" if "Voltage R" in row.index else "R -",
        f"S {format_voltage(row['Voltage S'])}" if "Voltage S" in row.index else "S -",
        f"T {format_voltage(row['Voltage T'])}" if "Voltage T" in row.index else "T -",
    ])

    health_values = [
        rectifier_condition(row),
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
        color = utility_color(utility) if idx == 1 else dynamic_color(text)
        if idx == 4:
            draw_text_fit(
                draw, text, *xy(HEALTH_X, y),
                round(HEALTH_W * sx), start_size=16, min_size=14,
                color=color, anchor="lm"
            )
        else:
            draw_text_fit(
                draw, text, *xy(HEALTH_X, y),
                round(HEALTH_W * sx), start_size=17, min_size=15,
                color=color, anchor="lm"
            )

    # -------------------------
    # ACTION
    # -------------------------
    actions = build_actions(row)
    if not actions:
        actions = ["No action required"]
        action_color = GREEN
    else:
        action_color = RED

    action_x, action_y = xy(1218, 733)
    action_width = round(365 * sx)
    action_font = get_font(round(17 * min(sx, sy)), bold=True)

    for action in actions[:3]:
        # Action is allowed to wrap because it is a dedicated instruction box.
        words = str(action).split()
        line = "•"
        lines = []
        for word in words:
            candidate = f"{line} {word}" if line != "•" else f"• {word}"
            if text_width(draw, candidate, action_font) <= action_width:
                line = candidate
            else:
                lines.append(line)
                line = word
        if line:
            lines.append(line)

        line_h = draw.textbbox((0, 0), "Ag", font=action_font)[3]
        for current in lines[:2]:
            draw.text((action_x, action_y), current, font=action_font, fill=action_color, anchor="la")
            action_y += line_h + 3
        action_y += 7
        if action_y > round(820 * sy):
            break

    output_path = f"output_{re.sub(r'[^A-Za-z0-9._-]+', '_', str(site_id))}.png"
    img.save(output_path, format="PNG", dpi=(150, 150))
    return output_path


# ------------------------------------------------------------
# Optional Telegram registration
# ------------------------------------------------------------

def register_power_handler(bot):
    """Register only /site. Does not register /genset or /area."""
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

            try:
                os.remove(img_path)
            except Exception:
                pass

            try:
                bot.delete_message(message.chat.id, status_msg.message_id)
            except Exception:
                pass

        except Exception as exc:
            print(f"Error generate Power report {site_id}: {exc}")
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
    # Simple local test:
    #   python power_module_precision.py CBN234
    import sys
    if len(sys.argv) > 1:
        result = generate_site_card(sys.argv[1])
        print(result or "SITE_NOT_FOUND")
