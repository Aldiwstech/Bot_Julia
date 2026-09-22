import os
import re
import pandas as pd
from PIL import Image, ImageDraw, ImageFont

# ============================================================
# GENSET MODULE V7 - PRECISION ALIGNMENT / CORRECT EXCEL MAPPING
# ============================================================
# Source columns are read by PHYSICAL EXCEL COLUMN LETTER:
#   Genset Autorate  : I, J, K, Q
#   Genset WarmingUp : I, M
#   BBM GensetFix    : T, X, Y
#
# IMPORTANT:
# - Every value is forced to ONE LINE.
# - No text wrapping.
# - Font size is fixed; long values are shortened with "...".
# - X and Y positions are fixed per card.
# - /site and Power renderer are not touched.
# ============================================================

EXCEL_CANDIDATES = [
    "Excel_master(1).xlsx",
    "Excel_master.xlsx",
]

MOCKUP_CANDIDATES = [
    "mokupgensetnew.png",
]

BASE_W = 1670
BASE_H = 942

NAVY = (24, 55, 105)
RED = (211, 42, 50)
GREEN = (20, 142, 68)
ORANGE = (232, 142, 18)


def find_existing(candidates):
    env_excel = os.getenv("EXCEL_PATH")
    if env_excel and os.path.exists(env_excel):
        return env_excel

    for p in candidates:
        if os.path.exists(p):
            return p
    return None


def resolve_sheet_name(excel_path, wanted):
    with pd.ExcelFile(excel_path) as book:
        sheets = book.sheet_names

    target = str(wanted).strip().casefold()
    for sheet in sheets:
        if str(sheet).strip().casefold() == target:
            return sheet

    raise ValueError(
        f"Worksheet named '{wanted}' not found. "
        f"Available: {', '.join(sheets)}"
    )


def is_empty(v):
    if v is None:
        return True
    try:
        if pd.isna(v):
            return True
    except Exception:
        pass
    return str(v).strip().lower() in ("", "nan", "none", "nat", "-")


def clean_text(v, default=""):
    """Force every Excel value into ONE visual line."""
    if is_empty(v):
        return default

    s = str(v)
    s = s.replace("\r\n", " ").replace("\n", " ").replace("\r", " ")
    s = s.replace("\t", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def format_display_value(value, kind=None):
    """Normalize a few Excel types so the visual stays clean and one-line."""
    if is_empty(value):
        return ""

    # Dates should never render as '00:00:00'.
    if kind == "date":
        try:
            dt = pd.to_datetime(value, errors="coerce")
            if not pd.isna(dt):
                return dt.strftime("%d-%m-%Y")
        except Exception:
            pass

    if kind == "latlong":
        try:
            num = float(str(value).replace(",", "."))
            return f"{num:.5f}"
        except Exception:
            return clean_text(value)

    return clean_text(value)


def get_font(size, bold=True):
    paths = (
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

    for p in paths:
        if os.path.exists(p):
            try:
                return ImageFont.truetype(p, size)
            except Exception:
                pass

    return ImageFont.load_default()


def shorten_one_line(draw, text, max_width, font):
    text = clean_text(text)
    if not text:
        return ""

    if draw.textbbox((0, 0), text, font=font)[2] <= max_width:
        return text

    suffix = "..."
    lo, hi = 0, len(text)

    while lo < hi:
        mid = (lo + hi + 1) // 2
        candidate = text[:mid].rstrip() + suffix
        if draw.textbbox((0, 0), candidate, font=font)[2] <= max_width:
            lo = mid
        else:
            hi = mid - 1

    return text[:lo].rstrip() + suffix


def display(v, default=""):
    return clean_text(v, default)


def dynamic_color(value):
    text = clean_text(value)
    upper = text.upper()

    if not text:
        return NAVY

    red_words = (
        "WARNING", "UNMONITOR", "UNAVAILABLE", "NOT AVAILABLE",
        "BROKEN", "CRITICAL", "DOWN", "FAULT", "FAILED", "ERROR",
        "NOT OK", "NOT SAFE", "UNBALANCE", "UNBALANCED", "PROBLEM",
        "NEED CHECK", "NEED VALIDATION", "VALIDATION", "WAITING",
        "POWER OFF", "NOT MONITOR", "NOT AUTO",
    )
    if any(w in upper for w in red_words):
        return RED

    match = re.search(r"(-?\d+(?:[.,]\d+)?)\s*%", text)
    if match:
        try:
            pct = float(match.group(1).replace(",", "."))
            if 0 <= pct <= 60:
                return GREEN
            if 60 < pct <= 90:
                return RED
        except ValueError:
            pass

    green_words = (
        "NORMAL", "SAFE", "ACTIVE", "AUTO", "OK", "SECURED",
        "VALID", "AVAILABLE", "MONITOR", "MONITORING",
        "BALANCE", "BALANCED", "CLOSED",
    )
    if any(w in upper for w in green_words):
        return GREEN

    return NAVY


def load_sheet(sheet_name):
    path = find_existing(EXCEL_CANDIDATES)
    if not path:
        raise FileNotFoundError(
            "Excel_master.xlsx / Excel_master(1).xlsx tidak ditemukan."
        )

    sheet = resolve_sheet_name(path, sheet_name)
    return pd.read_excel(path, sheet_name=sheet)


def find_site(df, site_id):
    if "Site ID" not in df.columns:
        raise KeyError("Kolom 'Site ID' tidak ditemukan.")

    wanted = str(site_id).strip().upper()
    ids = df["Site ID"].astype(str).str.strip().str.upper()
    found = df.loc[ids == wanted]

    return None if found.empty else found.iloc[0]


def row_value(row, column, default=""):
    if row is None or column not in row.index:
        return default
    return display(row[column], default)


def excel_column_value(row, column_letter, default=""):
    """Read by physical Excel column, not by header name."""
    if row is None:
        return default

    try:
        from openpyxl.utils import column_index_from_string
        idx = column_index_from_string(column_letter) - 1
        if idx < 0 or idx >= len(row):
            return default
        return display(row.iloc[idx], default)
    except Exception:
        return default


def load_genset_data(site_id):
    autorate = find_site(load_sheet("Genset Autorate"), site_id)
    warming = find_site(load_sheet("Genset WarmingUp"), site_id)
    bbm = find_site(load_sheet("BBM GensetFix"), site_id)

    if autorate is None and warming is None and bbm is None:
        return None

    return autorate, warming, bbm


def site_info(site_id):
    info = {
        "Site ID": str(site_id).upper(),
        "Site Name": "",
        "Regional": "",
        "NOP": "",
        "TO": "",
        "ROH": "",
        "Site Owner": "",
        "Lat / Long": "",
    }

    # Primary identity from Genset Autorate.
    try:
        autorate = find_site(load_sheet("Genset Autorate"), site_id)
        if autorate is not None:
            info["Site ID"] = row_value(
                autorate, "Site ID", str(site_id).upper()
            )
            info["Site Name"] = row_value(autorate, "Site Name")
            info["Regional"] = row_value(autorate, "Region")
            info["NOP"] = row_value(autorate, "NOP")
            info["TO"] = row_value(autorate, "TO")
    except Exception:
        pass

    # Metadata fallback from Power sheet.
    try:
        power = find_site(load_sheet("Rectifire&battery"), site_id)
        if power is not None:
            if not info["Site Name"]:
                info["Site Name"] = row_value(power, "Site Name")
            if not info["Regional"]:
                info["Regional"] = row_value(power, "Regional")
            if not info["NOP"]:
                info["NOP"] = row_value(power, "NOP")
            if not info["TO"]:
                info["TO"] = row_value(power, "TO")

            info["ROH"] = row_value(power, "ROH")
            info["Site Owner"] = row_value(power, "Site Owner")

            lat = row_value(power, "Lat")
            lon = row_value(power, "Long")
            if lat and lon:
                info["Lat / Long"] = (
                    f"{format_display_value(lat, 'latlong')} / "
                    f"{format_display_value(lon, 'latlong')}"
                )
    except Exception:
        pass

    return info


def is_bad(value):
    text = clean_text(value).upper()
    return any(
        w in text
        for w in (
            "BROKEN", "NOT AUTO", "NOT OK", "FAULT", "FAILED",
            "ERROR", "PROBLEM", "UNAVAILABLE", "DOWN",
            "UNMONITOR", "NOT SAFE",
        )
    )


def build_genset_health(autorate, warming):
    dg_status = excel_column_value(autorate, "I")
    dg_condition = excel_column_value(autorate, "J")
    week_condition = excel_column_value(autorate, "K")
    warm_status = excel_column_value(warming, "I")

    if (
        dg_status.upper() == "ACTIVE"
        and not is_bad(dg_condition)
        and not is_bad(week_condition)
        and not is_bad(warm_status)
    ):
        return "Normal"

    return "Need Check"


def build_genset_actions(autorate, warming, bbm):
    actions = []

    dg_status = excel_column_value(autorate, "I").upper()
    dg_condition = excel_column_value(autorate, "J").upper()
    week_condition = excel_column_value(autorate, "K").upper()

    warm_status = excel_column_value(warming, "I").upper()
    warm_week = excel_column_value(warming, "M").upper()

    bbm_status = excel_column_value(bbm, "X").upper()
    bbm_advice = excel_column_value(bbm, "Y").upper()

    if dg_condition == "NOT AUTO" or dg_status not in ("ACTIVE", "OK"):
        actions.append("Need Check Genset Auto")

    if is_bad(week_condition):
        actions.append("Need Check Genset Condition")

    if is_bad(warm_status) or is_bad(warm_week):
        actions.append("Need Check Genset Warming Up")

    if (
        "NOT SAFE" in bbm_status
        or "NOT SAFE" in bbm_advice
        or "UNAVAILABLE" in bbm_status
        or "UNAVAILABLE" in bbm_advice
    ):
        actions.append("Need Check BBM Genset")

    return list(dict.fromkeys(actions))


def generate_genset_card(site_id):
    data = load_genset_data(site_id)
    if data is None:
        return None

    autorate, warming, bbm = data

    mockup = find_existing(MOCKUP_CANDIDATES)
    if not mockup:
        raise FileNotFoundError(
            "Mockup Genset tidak ditemukan. "
            "Upload MokupGenset.png ke Railway."
        )

    img = Image.open(mockup).convert("RGB")

    sx = img.width / BASE_W
    sy = img.height / BASE_H
    scale = min(sx, sy)

    draw = ImageDraw.Draw(img)

    def xy(x, y):
        return round(x * sx), round(y * sy)

    def write(text, x, y, width, size=18, color=None, bold=True):
        text = clean_text(text)
        if not text:
            return

        # FIX: fixed font size. Never shrink the font.
        font = get_font(round(size * scale), bold=bold)
        max_width = round(width * sx)

        text = shorten_one_line(draw, text, max_width, font)
        if not text:
            return

        draw.text(
            xy(x, y),
            text,
            font=font,
            fill=color if color is not None else dynamic_color(text),
            anchor="lm",
        )

    def write_action(text, x, y, width, size=17, color=RED):
        text = clean_text(text)
        if not text:
            return

        font = get_font(round(size * scale), bold=True)
        max_width = round(width * sx)
        text = shorten_one_line(draw, text, max_width, font)

        draw.text(
            xy(x, y),
            text,
            font=font,
            fill=color,
            anchor="lm",
        )

    # ========================================================
    # LOCKED VALUE X
    # ========================================================
    # Measured against the current MokupGenset.png.
    # Value starts only a small gap after the printed ":".
    # IMPORTANT: these X values are FIXED and never depend on label length.
    SITE_VALUE_X = 265
    SITE_VALUE_W = 198

    MID_VALUE_X = 815
    MID_VALUE_W = 255

    HEALTH_VALUE_X = 1408
    HEALTH_VALUE_W = 220

    ACTION_VALUE_X = 1210
    ACTION_VALUE_W = 350

    # ========================================================
    # SITE INFO
    # ========================================================
    info = site_info(site_id)

    site_rows = [
        info["Site ID"],
        info["Site Name"],
        info["Regional"],
        info["NOP"],
        info["TO"],
        info["ROH"],
        info["Site Owner"],
        info["Lat / Long"],
    ]

    site_y = [231, 293, 355, 416, 478, 539, 600, 661]

    for i, (text, y) in enumerate(zip(site_rows, site_y)):
        # Coordinates get a compact but still readable 16px font so the
        # complete value stays inside the Site Info panel.
        write(
            text,
            SITE_VALUE_X,
            y,
            SITE_VALUE_W,
            size=16 if i == 7 else 18,
            color=NAVY,
        )

    # ========================================================
    # GENSET AUTORATE
    # Physical columns: I, J, K, Q
    # ========================================================
    autorate_rows = [
        excel_column_value(autorate, "I"),
        excel_column_value(autorate, "J"),
        excel_column_value(autorate, "K"),
        excel_column_value(autorate, "Q"),
    ]

    autorate_y = [214, 269, 323, 377]

    for text, y in zip(autorate_rows, autorate_y):
        write(
            text,
            MID_VALUE_X,
            y,
            MID_VALUE_W,
            size=18,
        )

    # ========================================================
    # GENSET WARMING UP
    # Physical columns: I, M
    # ========================================================
    # Mockup order is: Warming Up Status, then Current Week Condition.
    # Excel physical columns: M = Warming Up Weekly Status, I = Current Week Condition.
    warming_rows = [
        excel_column_value(warming, "M"),
        excel_column_value(warming, "I"),
    ]

    warming_y = [523, 577]

    for text, y in zip(warming_rows, warming_y):
        write(
            text,
            MID_VALUE_X,
            y,
            MID_VALUE_W,
            size=18,
        )

    # ========================================================
    # BBM GENSETFIX
    # Physical columns: T, X, Y
    # ========================================================
    bbm_rows = [
        excel_column_value(bbm, "T"),
        excel_column_value(bbm, "X"),
        excel_column_value(bbm, "Y"),
    ]

    bbm_y = [719, 771, 815]

    for text, y in zip(bbm_rows, bbm_y):
        write(
            text,
            MID_VALUE_X,
            y,
            MID_VALUE_W,
            size=18,
        )

    # ========================================================
    # HEALTHY CHECK
    # ========================================================
    # Healthy Check mirrors the labels already printed in the mockup.
    # All source positions are physical Excel columns so column/header changes
    # cannot silently shift the data into the wrong row.
    health_rows = [
        build_genset_health(autorate, warming),   # Genset Health
        excel_column_value(autorate, "I"),        # DG Status
        excel_column_value(autorate, "J"),        # DG Condition
        excel_column_value(autorate, "K"),        # Current Week Condition
        excel_column_value(autorate, "M"),        # Problem Genset
        excel_column_value(autorate, "N"),        # RCA
        excel_column_value(autorate, "O"),        # Plan Action
        excel_column_value(autorate, "Q"),        # Current Progress
        excel_column_value(autorate, "R"),        # PIC
        format_display_value(excel_column_value(autorate, "T"), "date"),  # Auto Date
    ]

    health_y = [
        210, 254, 298, 342, 386,
        430, 476, 523, 570, 616
    ]

    for text, y in zip(health_rows, health_y):
        if text:
            write(
                text,
                HEALTH_VALUE_X,
                y,
                HEALTH_VALUE_W,
                size=17,
            )

    # ========================================================
    # ACTION
    # ========================================================
    actions = build_genset_actions(autorate, warming, bbm)

    if not actions:
        actions = ["No action required"]
        action_color = GREEN
    else:
        action_color = RED

    action_y = 765

    for action in actions[:4]:
        write_action(
            action,
            ACTION_VALUE_X,
            action_y,
            ACTION_VALUE_W,
            size=17,
            color=action_color,
        )
        action_y += 30

    # ========================================================
    # SAVE
    # ========================================================
    safe_site = re.sub(
        r"[^A-Za-z0-9._-]+",
        "_",
        str(site_id),
    )

    output = f"output_genset_{safe_site}.png"

    img.save(
        output,
        format="PNG",
        dpi=(150, 150),
    )

    return output


def register_genset_handler(bot):
    @bot.message_handler(commands=["genset"])
    def handle_genset(message):
        args = message.text.split()

        if len(args) < 2:
            bot.reply_to(
                message,
                "⚠️ Format salah!\n"
                "Gunakan: `/genset <Site_ID>`",
                parse_mode="Markdown",
            )
            return

        site_id = args[1].strip().upper()

        status_msg = bot.reply_to(
            message,
            f"⏳ Sedang memproses Genset Site ID: *{site_id}*...",
            parse_mode="Markdown",
        )

        try:
            img_path = generate_genset_card(site_id)

            if not img_path or not os.path.exists(img_path):
                bot.edit_message_text(
                    f"❌ Data Genset Site ID *{site_id}* tidak ditemukan.",
                    message.chat.id,
                    status_msg.message_id,
                    parse_mode="Markdown",
                )
                return

            with open(img_path, "rb") as photo:
                bot.send_photo(
                    message.chat.id,
                    photo,
                    caption=f"✅ Genset Report Site ID: *{site_id}*",
                    parse_mode="Markdown",
                )

            os.remove(img_path)

            try:
                bot.delete_message(
                    message.chat.id,
                    status_msg.message_id,
                )
            except Exception:
                pass

        except Exception as exc:
            print(f"Error generate genset report {site_id}: {exc}")

            try:
                bot.edit_message_text(
                    f"❌ Error Genset *{site_id}*.\n`{exc}`",
                    message.chat.id,
                    status_msg.message_id,
                    parse_mode="Markdown",
                )
            except Exception:
                bot.reply_to(
                    message,
                    f"❌ Error saat membuat report Genset *{site_id}*.",
                    parse_mode="Markdown",
                )
