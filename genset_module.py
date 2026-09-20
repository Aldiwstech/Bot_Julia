import os
import re
import pandas as pd
from PIL import Image, ImageDraw, ImageFont

# Genset module is intentionally isolated from the Power/Rectifier renderer.
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

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
    explicit = os.getenv("EXCEL_PATH")
    if explicit:
        if not os.path.isabs(explicit):
            explicit = os.path.join(BASE_DIR, explicit)
        if os.path.exists(explicit):
            return explicit

    existing = []
    for name in candidates:
        path = name if os.path.isabs(name) else os.path.join(BASE_DIR, name)
        if os.path.exists(path):
            existing.append(path)
    if not existing:
        return None
    return max(existing, key=os.path.getmtime)


def resolve_sheet_name(excel_path, wanted):
    import openpyxl
    wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
    names = wb.sheetnames
    target = str(wanted).strip().casefold()
    for name in names:
        if str(name).strip().casefold() == target:
            return name
    raise KeyError(f"Worksheet named '{wanted}' not found. Available: {names}")


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


def fit_font(draw, text, max_width, size=20, minimum=11):
    for n in range(int(size), int(minimum) - 1, -1):
        f = get_font(n, bold=True)
        if draw.textbbox((0, 0), str(text), font=f)[2] <= max_width:
            return f
    return get_font(minimum, bold=True)


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
    if any(x in t for x in (
        "NEED REFUEL", "NOT SAFE", "CRITICAL", "DOWN", "FAULT",
        "FAILED", "ERROR", "NOT OK", "UNBALANCE", "UNBALANCED",
        "PROBLEM",
    )):
        return RED
    if any(x in t for x in ("WARNING", "CHECK", "FAIR", "LOW", "MEDIUM")):
        return ORANGE
    if any(x in t for x in (
        "ACTIVE", "AUTO", "CLOSED", "OK", "SAFE", "NORMAL",
        "SECURED", "VALID", "AVAILABLE", "MONITOR",
    )):
        return GREEN
    return NAVY


def load_sheet(sheet_name):
    path = find_existing(EXCEL_CANDIDATES)
    if not path:
        raise FileNotFoundError("Excel_master.xlsx tidak ditemukan.")
    return pd.read_excel(path, sheet_name=resolve_sheet_name(path, sheet_name))


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
    """Generate concise onsite instructions from the three selected sources."""
    actions = []

    dg_condition = row_value(autorate, "DG Condition", "").upper()
    if dg_condition and dg_condition not in ("AUTO", "NORMAL", "OK", "SAFE"):
        actions.append("Need Check Genset Auto")

    dg_week = row_value(autorate, "Current Week Genset Condition", "").upper()
    if dg_week and dg_week not in ("OK", "NORMAL", "SAFE", "GOOD", "ACTIVE"):
        actions.append("Need Check Genset Condition")

    warm_status = row_value(warming, "Warming Up Weekly Status", "").upper()
    warm_week = row_value(warming, "Current Week Genset Condition", "").upper()
    if warm_status and warm_status not in ("OK", "NORMAL", "SAFE", "GOOD", "ACTIVE"):
        actions.append("Need Check Genset Warming Up")
    elif warm_week and warm_week not in ("OK", "NORMAL", "SAFE", "GOOD", "ACTIVE"):
        actions.append("Need Check Genset Warming Up")

    bbm_status = row_value(bbm, "Status", "").upper()
    if bbm_status and bbm_status not in ("SAFE", "NORMAL", "OK", "GOOD"):
        suggestion = row_value(bbm, "Saran Pengisian", "").strip()
        if suggestion and suggestion.upper() not in ("SAFE", "NORMAL", "OK", "GOOD", "-"):
            actions.append(f"BBM: {suggestion}")
        else:
            actions.append("Need Check / Refill BBM Genset")

    # Preserve order and remove duplicates.
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

    def write(text, x, y, width, size=20, color=None):
        text = display(text)
        px = round(width * sx)
        font = fit_font(draw, text, px, size=round(size * min(sx, sy)), minimum=10)
        text = shorten(draw, text, px, font)
        draw.text(
            xy(x, y), text, font=font,
            fill=color if color is not None else dynamic_color(text),
            anchor="lm",
        )

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
    for i, (text, y) in enumerate(zip(vals, [236, 291, 350, 414, 477, 543, 610, 675])):
        write(text, 270, y, 205, size=19 if i in (1, 7) else 20)

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
        write(text, 800, y, 270, size=18)

    warming_vals = [
        row_value(warming, "Warming Up Weekly Status"),
        row_value(warming, "Current Week Genset Condition"),
    ]
    for text, y in zip(warming_vals, [522, 576]):
        write(text, 800, y, 270, size=18)

    bbm_vals = [
        row_value(bbm, "Perkiraan Sisa Fuel"),
        row_value(bbm, "Status"),
        row_value(bbm, "Saran Pengisian"),
    ]
    for text, y in zip(bbm_vals, [720, 774, 817]):
        write(text, 800, y, 270, size=17)

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
        write(text, 1390, y, 220, size=16)

    # -------------------------
    # ACTION
    # Separate action card, matching the Power layout.
    # -------------------------
    actions = build_genset_actions(autorate, warming, bbm)
    action_y = 765
    for action in actions[:4]:
        write(action, 1220, action_y, 360, size=15, color=RED)
        action_y += 27

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
