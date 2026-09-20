import os
import re
import telebot
import pandas as pd
from PIL import Image, ImageDraw, ImageFont

TOKEN = os.getenv("TOKEN")
if not TOKEN:
    raise ValueError("Token belum diset di Environment Variables Railway!")

bot = telebot.TeleBot(TOKEN)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

EXCEL_CANDIDATES = ["Excel_master(1).xlsx", "Excel_master.xlsx"]
MOCKUP_CANDIDATES = [
    "Mokup.png", "mokup.png", "mokup(1).png", "Mokup(1).png",
    "a_clean_flat_vector_infographic_dashboard_templat.png",
    "Blank Telkomsel Huawei Dashboard Template.png"
]

# Final mockup coordinate system: 1672 x 941.
BASE_W = 1672
BASE_H = 941

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
    # Prefer an explicitly configured workbook. Otherwise search next to
    # this script, so launching Python from another working directory is safe.
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
    # Excel sheet names are case-sensitive in pandas/openpyxl lookup.
    # Resolve them case-insensitively so Rectifire&Battery and
    # Rectifire&battery are treated as the same source sheet.
    import openpyxl
    wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
    names = wb.sheetnames
    target = str(wanted).strip().casefold()
    for name in names:
        if str(name).strip().casefold() == target:
            return name
    raise KeyError(f"Worksheet named '{wanted}' not found. Available: {names}")


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


def dynamic_color(value):
    """Color for ordinary values. UNMONITOR is checked before MONITOR."""
    text = str(value).strip().upper()

    # IMPORTANT: UNMONITOR contains MONITOR, so check it first.
    if text == "UNMONITOR":
        return RED
    if text in ("NOT AVAILABLE", "NOT_AVAILABLE"):
        return RED

    if any(w in text for w in ("CRITICAL", "DOWN", "NEED", "TRIP", "FAULT",
                               "FAILED", "ERROR", "NO BACKUP", "UNBALANCE",
                               "UNBALANCED")):
        return RED
    if any(w in text for w in ("POTENSIAL", "POTENTIAL", "WARNING", "CHECK")):
        return ORANGE
    if any(w in text for w in ("NORMAL", "SECURED", "OK", "VALID",
                               "AVAILABLE", "MONITOR", "SAFE", "BALANCE",
                               "BALANCED")):
        return GREEN
    return NAVY


def load_dataframe():
    excel_path = find_existing(EXCEL_CANDIDATES)
    if not excel_path:
        raise FileNotFoundError("Excel_master.xlsx tidak ditemukan.")

    df = pd.read_excel(excel_path, sheet_name=resolve_sheet_name(excel_path, "Rectifire&Battery"))
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


def format_voltage(value_):
    n = safe_float(value_)
    if n is None:
        return "-"
    if abs(n - round(n)) < 1e-9:
        return f"{int(round(n))} V"
    return f"{n:.1f} V"


def phase_condition(row):
    """
    Rect. Cond is based on Phase Balanced, as agreed.
    The source field remains the source of truth.
    """
    phase = value(row, "Phase Balanced")
    if phase == "-":
        return "-"
    return phase


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
    rect_cond = value(row, "Rectifier Condition").upper()
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

    def write_value(text, x, y, max_width, size=21, color=None):
        text = display_value(text)
        font = fit_font(
            draw,
            text,
            round(max_width * sx),
            size=round(size * min(sx, sy)),
            minimum=10,
            bold=True,
        )
        draw.text(
            xy(x, y),
            text,
            font=font,
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
    for i, (text, y) in enumerate(zip(site_rows, [236, 291, 346, 400, 455, 510, 575, 634])):
        site_size = 20
        site_width = 205
        if i == 7:  # Lat / Long is intentionally one point smaller.
            site_size = 18
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
    for text, y in zip(rect_rows, [212, 257, 302, 346, 391, 436, 479]):
        write_value(text, 798, y, 270, size=20)

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
    for text, y in zip(batt_rows, [610, 651, 692, 732, 772, 812]):
        write_value(text, 798, y, 270, size=18)

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
    health_y = [229, 283, 337, 391, 445, 499, 553, 611]
    for text, y in zip(health_values, health_y):
        write_value(text, 1394, y, 220, size=18)

    # Action is ALWAYS red, per the agreed design.
    # Keep each instruction inside the right-panel value area.
    actions = build_actions(row)
    if not actions:
        actions = ["-"]

    action_y = 671
    action_max_width = 225
    for action in actions[:3]:
        write_value(
            f"- {action}",
            1394,
            action_y,
            action_max_width,
            size=12,
            color=RED,
        )
        action_y += 22

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
