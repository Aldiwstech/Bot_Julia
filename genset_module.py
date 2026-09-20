import os
import re
import pandas as pd
from PIL import Image, ImageDraw, ImageFont

# Genset module is intentionally isolated from the Power/Rectifier renderer.
EXCEL_CANDIDATES = ["Excel_master(1).xlsx", "Excel_master.xlsx"]
MOCKUP_CANDIDATES = [
    "Mokupgenset.png",
    "MokupGenset.png",
    "mokupgenset.png",
]

BASE_W = 1671
BASE_H = 941

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
    raise ValueError(f"Worksheet named '{wanted}' not found. Available: {', '.join(sheets)}")


def is_empty(v):
    if v is None:
        return True
    try:
        if pd.isna(v):
            return True
    except Exception:
        pass
    return str(v).strip().lower() in ("", "nan", "none", "-", "nat")


def display(v, default="-"):
    return default if is_empty(v) else str(v).strip()


def get_font(size, bold=True):
    paths = (
        [
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
    for p in paths:
        if os.path.exists(p):
            try:
                return ImageFont.truetype(p, size)
            except Exception:
                pass
    return ImageFont.load_default()


def fit_font(draw, text, max_width, size=20, minimum=11, bold=True):
    for n in range(int(size), int(minimum) - 1, -1):
        f = get_font(n, bold=bold)
        bbox = draw.textbbox((0, 0), str(text), font=f)
        if bbox[2] - bbox[0] <= max_width:
            return f
    return get_font(minimum, bold=bold)


def wrap_text(draw, text, font, max_width, max_lines=2):
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
    kept = lines[:max_lines-1]
    remainder = " ".join(lines[max_lines-1:])
    while remainder and draw.textbbox((0, 0), remainder + "...", font=font)[2] > max_width:
        remainder = remainder[:-1].rstrip()
    kept.append((remainder + "...") if remainder else "..." )
    return kept


def shorten(draw, text, max_width, font):
    text = str(text)
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


def dynamic_color(v):
    t = str(v).strip().upper()
    if t in ("UNMONITOR", "NOT AVAILABLE", "NOT_AVAILABLE"):
        return RED
    # Negative phrases must be checked before positive substrings (e.g.
    # NOT AUTO must not inherit the green AUTO rule).
    if any(x in t for x in (
        "NOT AUTO", "BROKEN", "CRITICAL", "DOWN", "FAULT", "FAILED",
        "ERROR", "NOT OK", "NOT SAFE", "UNBALANCE", "UNBALANCED",
        "PROBLEM", "NEED CHECK", "NEED VALIDATION"
    )):
        return RED
    if any(x in t for x in ("WARNING", "CHECK", "FAIR", "LOW", "MEDIUM")):
        return ORANGE
    if any(x in t for x in (
        "ACTIVE", "AUTO", "CLOSED", "OK", "SAFE", "NORMAL", "SECURED",
        "VALID", "AVAILABLE", "MONITOR", "BALANCE", "BALANCED"
    )):
        return GREEN
    return NAVY


def load_sheet(sheet_name):
    path = find_existing(EXCEL_CANDIDATES)
    if not path:
        raise FileNotFoundError("Excel_master.xlsx tidak ditemukan.")
    sheet = resolve_sheet_name(path, sheet_name)
    return pd.read_excel(path, sheet_name=sheet)


def find_site(df, site_id):
    if "Site ID" not in df.columns:
        raise KeyError("Kolom 'Site ID' tidak ditemukan.")
    wanted = str(site_id).strip().upper()
    ids = df["Site ID"].astype(str).str.strip().str.upper()
    found = df.loc[ids == wanted]
    return None if found.empty else found.iloc[0]


def load_genset_data(site_id):
    autorate = find_site(load_sheet("Genset Autorate"), site_id)
    warming = find_site(load_sheet("Genset WarmingUp"), site_id)
    bbm = find_site(load_sheet("BBM GensetFix"), site_id)

    # Site must exist in the genset source. Missing individual sheet data is allowed.
    if autorate is None and warming is None and bbm is None:
        return None

    return autorate, warming, bbm


def row_value(row, column, default="-"):
    if row is None or column not in row.index:
        return default
    return display(row[column], default)


def site_info(site_id):
    # Site Info uses the existing Power source only for identity/location metadata.
    # This does not change the Power renderer.
    try:
        df = load_sheet("Rectifire&battery")
        row = find_site(df, site_id)
        if row is not None:
            return {
                "Site ID": row_value(row, "Site ID"),
                "Site Name": row_value(row, "Site Name"),
                "Regional": row_value(row, "Regional"),
                "NOP": row_value(row, "NOP"),
                "TO": row_value(row, "TO"),
                "ROH": row_value(row, "ROH"),
                "Site Owner": row_value(row, "Site Owner"),
                "Lat / Long": f"{row_value(row, 'Lat')} / {row_value(row, 'Long')}",
            }
    except Exception:
        pass

    # Fallback to Autorate identity fields without using them as the selected metrics.
    df = load_sheet("Genset Autorate")
    row = find_site(df, site_id)
    if row is None:
        return {"Site ID": str(site_id).upper()}
    return {
        "Site ID": str(site_id).upper(),
        "Site Name": row_value(row, "Site Name"),
        "Regional": row_value(row, "Region"),
        "NOP": row_value(row, "NOP"),
        "TO": row_value(row, "TO"),
    }


def build_genset_actions(autorate, warming, bbm):
    actions = []
    dg_status = row_value(autorate, "DG Status").upper()
    dg_condition = row_value(autorate, "DG Condition").upper()
    week_condition = row_value(autorate, "Current Week Genset Condition").upper()
    warm_status = row_value(warming, "Warming Up Weekly Status").upper()
    warm_week = row_value(warming, "Current Week Genset Condition").upper()

    if dg_condition == "NOT AUTO" or dg_status not in ("ACTIVE", "OK"):
        actions.append("Need Check Genset Auto")
    if week_condition in ("BROKEN", "NOT OK", "FAULT", "PROBLEM"):
        actions.append("Need Check Genset Condition")
    if warm_status in ("BROKEN", "NOT OK", "FAULT", "PROBLEM") or warm_week in ("BROKEN", "NOT OK", "FAULT", "PROBLEM"):
        actions.append("Need Check Genset Warming Up")

    # Keep future BBM rules explicit; do not invent an action from a safe status.
    return list(dict.fromkeys(actions))


def generate_genset_card(site_id):
    data = load_genset_data(site_id)
    if data is None:
        return None

    autorate, warming, bbm = data
    mockup = find_existing(MOCKUP_CANDIDATES)
    if not mockup:
        raise FileNotFoundError(
            "Mokupgenset.png tidak ditemukan. Simpan mockup dengan nama Mokupgenset.png."
        )

    img = Image.open(mockup).convert("RGB")
    sx, sy = img.width / BASE_W, img.height / BASE_H
    draw = ImageDraw.Draw(img)

    def xy(x, y):
        return round(x * sx), round(y * sy)

    def write(text, x, y, width, size=20, color=None, bold=True):
        text = display(text)
        px = round(width * sx)
        font = fit_font(draw, text, px, size=round(size * min(sx, sy)), minimum=10, bold=bold)
        text = shorten(draw, text, px, font)
        draw.text(
            xy(x, y), text, font=font,
            fill=color if color is not None else dynamic_color(text),
            anchor="lm",
        )

    def write_wrapped(text, x, y, width, size=15, color=RED, max_lines=3):
        px = round(width * sx)
        font = fit_font(draw, text, px, size=round(size * min(sx, sy)), minimum=11, bold=True)
        lines = wrap_text(draw, text, font, px, max_lines=max_lines)
        line_h = draw.textbbox((0, 0), "Ag", font=font)[3] - draw.textbbox((0, 0), "Ag", font=font)[1]
        py = round(y * sy)
        for line in lines:
            draw.text((round(x * sx), py), line, font=font, fill=color, anchor="la")
            py += line_h + round(4 * sy)

    # -------------------------
    # SITE INFO
    # -------------------------
    info = site_info(site_id)
    vals = [
        info.get("Site ID", "-"),
        info.get("Site Name", "-"),
        info.get("Regional", "-"),
        info.get("NOP", "-"),
        info.get("TO", "-"),
        info.get("ROH", "-"),
        info.get("Site Owner", "-"),
        info.get("Lat / Long", "-"),
    ]
    for i, (text, y) in enumerate(zip(vals, [231, 290, 349, 409, 468, 526, 610, 676])):
        write(text, 270, y, 202, size=17 if i == 7 else (19 if i == 1 else 20))

    # -------------------------
    # 3 SOURCE CATEGORIES
    # Only the columns explicitly requested are populated:
    # Autorate: I, J, K, Q
    # WarmingUp: I, M
    # BBM GensetFix: T, X, Y
    # -------------------------
    autorate_vals = [
        row_value(autorate, "DG Status"),
        row_value(autorate, "DG Condition"),
        row_value(autorate, "Current Week Genset Condition"),
        row_value(autorate, "Current Progress"),
    ]
    for text, y in zip(autorate_vals, [213, 266, 320, 374]):
        write(text, 830, y, 220, size=17)

    warming_vals = [
        row_value(warming, "Warming Up Weekly Status"),
        row_value(warming, "Current Week Genset Condition"),
    ]
    for text, y in zip(warming_vals, [522, 576]):
        write(text, 830, y, 220, size=17)

    bbm_vals = [
        row_value(bbm, "Perkiraan Sisa Fuel"),
        row_value(bbm, "Status"),
        row_value(bbm, "Saran Pengisian"),
    ]
    for text, y in zip(bbm_vals, [720, 774, 817]):
        write(text, 830, y, 220, size=16)

    # -------------------------
    # HEALTHY CHECK
    # No unsupported source columns are invented.
    # Directly sourced selected values are echoed here only where
    # the mockup has matching labels. Other rows remain blank.
    # -------------------------
    health = [
        row_value(autorate, "DG Status"),
        row_value(autorate, "DG Condition"),
        row_value(autorate, "Current Week Genset Condition"),
        "-",  # Problem Genset is not one of the requested columns.
        "-",  # RCA is not one of the requested columns.
        "-",  # Plan Action is not one of the requested columns.
        row_value(autorate, "Current Progress"),
        "-",  # PIC is not one of the requested columns.
        "-",  # Auto Date is not one of the requested columns.
    ]
    for text, y in zip(health, [212, 266, 320, 374, 428, 482, 536, 590, 634]):
        # Blank unsupported rows instead of inserting "-" into the visual.
        if text == "-":
            continue
        write(text, 1412, y, 205, size=15)

    # -------------------------
    # ACTION
    # Separate action panel. Larger red text + wrapping keeps the instruction readable.
    # -------------------------
    actions = build_genset_actions(autorate, warming, bbm)
    if not actions:
        actions = ["No action required"]
        action_color = GREEN
    else:
        action_color = RED
    action_y = 754
    for action in actions[:3]:
        write_wrapped(f"• {action}", 1230, action_y, 330, size=16, color=action_color, max_lines=1)
        action_y += 30

    output = f"output_genset_{re.sub(r'[^A-Za-z0-9._-]+', '_', str(site_id))}.png"
    img.save(output, format="PNG", dpi=(150, 150))
    return output


def register_genset_handler(bot):
    @bot.message_handler(commands=["genset"])
    def handle_genset(message):
        args = message.text.split()
        if len(args) < 2:
            bot.reply_to(
                message,
                "⚠️ Format salah!\nGunakan: `/genset <Site_ID>`",
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
                bot.delete_message(message.chat.id, status_msg.message_id)
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
