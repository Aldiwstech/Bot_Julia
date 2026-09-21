import os
import re
import pandas as pd
from PIL import Image, ImageDraw, ImageFont

# ============================================================
# AREA MODULE
# Source : Master Data PBI_A2_Rev1.xlsx
# Sheet  : Area 2
# Command: /area <SITE_ID>
#
# IMPORTANT:
# This module NEVER reads Excel_master.xlsx.
# It is intentionally isolated from /site and /genset.
# ============================================================

AREA_EXCEL_CANDIDATES = [
    "Master Data PBI_A2_Rev1.xlsx",
]

AREA_MOCKUP_CANDIDATES = [
    "power_mockup_latest.png",
    "Mokupnew.png",
    "MokupPowerAction.png",
]

BASE_W = 1683
BASE_H = 935

NAVY = (24, 55, 105)
RED = (211, 42, 50)
GREEN = (20, 142, 68)
ORANGE = (232, 142, 18)


def find_area_excel():
    env = os.getenv("AREA_EXCEL_PATH")
    if env and os.path.exists(env):
        return env

    for path in AREA_EXCEL_CANDIDATES:
        if os.path.exists(path):
            return path

    raise FileNotFoundError(
        "Master Data PBI_A2_Rev1.xlsx tidak ditemukan."
    )


def find_mockup():
    env = os.getenv("AREA_MOCKUP_PATH")
    if env and os.path.exists(env):
        return env

    for path in AREA_MOCKUP_CANDIDATES:
        if os.path.exists(path):
            return path

    raise FileNotFoundError(
        "Mockup Power untuk /area tidak ditemukan."
    )


def normalize_col(name):
    return re.sub(r"\s+", " ", str(name).strip()).casefold()


def find_column(df, *names):
    lookup = {
        normalize_col(col): col
        for col in df.columns
    }

    for name in names:
        key = normalize_col(name)
        if key in lookup:
            return lookup[key]

    return None


def value(row, *columns, default=""):
    if row is None:
        return default

    col = find_column(row.to_frame().T, *columns)
    if not col:
        return default

    val = row[col]

    try:
        if pd.isna(val):
            return default
    except Exception:
        pass

    text = str(val).strip()

    if text.lower() in ("", "nan", "none", "nat", "-"):
        return default

    return text


def load_area_row(site_id):
    excel = find_area_excel()

    df = pd.read_excel(
        excel,
        sheet_name="Area 2",
    )

    site_col = find_column(df, "Site ID")

    if not site_col:
        raise KeyError(
            "Kolom Site ID tidak ditemukan pada sheet Area 2."
        )

    wanted = str(site_id).strip().upper()

    ids = (
        df[site_col]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    result = df.loc[ids == wanted]

    if result.empty:
        return None

    return result.iloc[0]


def font(size, bold=True):
    paths = (
        [
            "/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        ]
        if bold
        else [
            "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        ]
    )

    for path in paths:
        if os.path.exists(path):
            return ImageFont.truetype(path, size)

    return ImageFont.load_default()


def fit_font(draw, text, max_width, size=18, minimum=11, bold=True):
    for n in range(size, minimum - 1, -1):
        f = font(n, bold=bold)
        if draw.textbbox((0, 0), str(text), font=f)[2] <= max_width:
            return f

    return font(minimum, bold=bold)


def dynamic_color(text):
    t = str(text).strip()
    u = t.upper()

    if not t:
        return NAVY

    # User's current color rules.
    red_words = (
        "WARNING",
        "UNMONITOR",
        "UNAVAILABLE",
        "NOT AVAILABLE",
        "NEED VALIDATION",
        "VALIDATION",
        "UNBALANCE",
        "UNBALANCED",
        "BROKEN",
        "FAULT",
        "FAILED",
        "ERROR",
        "DOWN",
        "NOT AUTO",
        "PROBLEM",
        "NEED CHECK",
        "0 V",
        "0V",
    )

    if any(x in u for x in red_words):
        return RED

    pct = re.search(r"(-?\d+(?:[.,]\d+)?)\s*%", t)
    if pct:
        try:
            number = float(pct.group(1).replace(",", "."))
            if 0 <= number <= 60:
                return GREEN
            if 60 < number <= 90:
                return RED
        except ValueError:
            pass

    green_words = (
        "NORMAL",
        "SAFE",
        "OK",
        "ACTIVE",
        "VALID",
        "AVAILABLE",
        "MONITOR",
        "MONITORING",
        "BALANCED",
        "BALANCE",
    )

    if any(x in u for x in green_words):
        return GREEN

    return NAVY


def draw_value(draw, text, x, y, width, size=18, color=None):
    text = str(text).strip()

    if not text:
        return

    max_width = int(width)

    f = fit_font(
        draw,
        text,
        max_width,
        size=size,
        minimum=11,
        bold=True,
    )

    # Never move X according to text length.
    # The value is always anchored at the same X.
    while (
        draw.textbbox((0, 0), text, font=f)[2] > max_width
        and len(text) > 8
    ):
        text = text[:-4].rstrip() + "..."

    draw.text(
        (int(x), int(y)),
        text,
        font=f,
        fill=color if color is not None else dynamic_color(text),
        anchor="lm",
    )


def build_area_data(row):
    # --------------------------------------------------------
    # SITE INFO
    # --------------------------------------------------------
    site = {
        "Site ID": value(row, "Site ID"),
        "Site Name": value(row, "Site Name"),
        "Regional": value(row, "Regional"),
        "NOP": value(row, "NOP"),
        "TO": value(row, "TO"),
        "ROH": value(row, "ROH"),
        "Site Owner": value(row, "Site Owner"),
        "Class Site": value(row, "Class Site", "New Class Site"),
        "VIP": value(row, "VIP", "New VIP/NON VIP"),
        "Lat / Long": "",
    }

    lat = value(row, "Lat")
    lon = value(row, "Long")

    if lat and lon:
        site["Lat / Long"] = f"{lat} / {lon}"

    # --------------------------------------------------------
    # RECTIFIER & PLN
    # --------------------------------------------------------
    power = {
        "ID PLN": value(row, "ID PLN"),
        "Daya PLN": value(row, "Daya PLN (KVA)"),
        "Phase": value(row, "Phase"),
        "MCB": value(row, "MCB"),
        "Voltage R": value(row, "Voltage R"),
        "Voltage S": value(row, "Voltage S"),
        "Voltage T": value(row, "Voltage T"),
        "Utility PLN": value(row, "Utility PLN"),
        "Capacity Status": value(row, "Capacity Status"),
        "Phase Balanced": value(row, "Phase Balanced"),
    }

    # --------------------------------------------------------
    # BATTERY
    # --------------------------------------------------------
    battery = {
        "Battery Type": value(row, "Battery Type (1)"),
        "Battery Brand": value(row, "Battery Brand (1)"),
        "Battery Capacity": value(row, "Battery Capacity (1)"),
        "Battery Bank": value(row, "Battery Bank (1)"),
        "BBT H": value(row, "BBT H (1)"),
        "BBT Category": value(row, "BBT Category (1)"),
        "Battery Install": value(row, "Battery Install (1)"),
    }

    # --------------------------------------------------------
    # HEALTHY CHECK
    # --------------------------------------------------------
    health = {
        "EAS Validation": value(row, "EAS Validation"),
        "NETECO Status": value(row, "NETECO Status"),
        "Rectifier Condition": value(row, "Rectifier Condition"),
        "Capacity Status": power["Capacity Status"],
        "Phase Balanced": power["Phase Balanced"],
        "Rectifier Utility": value(row, "Rectifier Utility"),
        "Last Check": value(row, "Last Check Update"),
    }

    # --------------------------------------------------------
    # ACTION
    # --------------------------------------------------------
    actions = []

    eas = health["EAS Validation"].upper()
    neteco = health["NETECO Status"].upper()
    rect_cond = health["Rectifier Condition"].upper()
    capacity = health["Capacity Status"].upper()
    phase = health["Phase Balanced"].upper()

    # EAS rule from the agreed workflow.
    if "NEED VALIDATION" in eas or "VALIDATION" in eas:
        actions.append("Need Check Onsite EAS")

    # NetEco rule from the agreed workflow.
    if "UNMONITOR" in neteco or "UNMONITOR" in neteco.replace(" ", ""):
        actions.append("Need Check Onsite Connection NetEco")

    # PLN / phase rule:
    # If unbalanced, inspect voltage. Any phase at 0 means
    # the immediate onsite check should focus on that phase.
    if "UNBALANCE" in phase or "UNBALANCED" in phase:
        zeros = []
        for label in ("R", "S", "T"):
            v = power[f"Voltage {label}"]
            try:
                num = float(
                    re.sub(r"[^0-9.\-]", "", v)
                )
                if num == 0:
                    zeros.append(label)
            except Exception:
                pass

        if zeros:
            actions.append(
                "Check PLN Voltage Phase "
                + "/".join(zeros)
            )
        else:
            actions.append(
                "Check PLN Phase Balance & Voltage"
            )

    # Capacity status.
    if (
        "UNBALANCE" in capacity
        or "UNBALANCED" in capacity
        or "WARNING" in capacity
        or "NOT OK" in capacity
    ):
        actions.append("Check PLN Capacity & Load")

    # Rectifier condition.
    if any(
        x in rect_cond
        for x in (
            "BROKEN",
            "FAULT",
            "FAILED",
            "ERROR",
            "BAD",
            "PROBLEM",
            "UNAVAILABLE",
        )
    ):
        actions.append("Check Rectifier Condition")

    # If there is no trigger, keep the card explicit.
    if not actions:
        actions = ["No action required"]

    return site, power, battery, health, actions


def render_area(site_id):
    row = load_area_row(site_id)

    if row is None:
        return None

    site, power, battery, health, actions = build_area_data(row)

    mockup = find_mockup()
    img = Image.open(mockup).convert("RGB")
    draw = ImageDraw.Draw(img)

    sx = img.width / BASE_W
    sy = img.height / BASE_H

    def P(x, y):
        return (round(x * sx), round(y * sy))

    # --------------------------------------------------------
    # FIXED ANCHORS
    #
    # These are card anchors, not label-dependent positions.
    # All values in the same panel share one X.
    # --------------------------------------------------------

    SITE_X = 305
    SITE_W = 255

    POWER_X = 845
    POWER_W = 270

    HEALTH_X = 1410
    HEALTH_W = 230

    ACTION_X = 1215
    ACTION_W = 390

    # --------------------------------------------------------
    # SITE INFO - 10 rows
    # --------------------------------------------------------
    site_values = [
        site["Site ID"],
        site["Site Name"],
        site["Regional"],
        site["NOP"],
        site["TO"],
        site["ROH"],
        site["Site Owner"],
        site["Class Site"],
        site["VIP"],
        site["Lat / Long"],
    ]

    site_y = [
        208,
        258,
        308,
        358,
        408,
        458,
        508,
        558,
        608,
        658,
    ]

    for text, y in zip(site_values, site_y):
        draw_value(
            draw,
            text,
            SITE_X * sx,
            y * sy,
            SITE_W * sx,
            size=17,
            color=NAVY,
        )

    # --------------------------------------------------------
    # RECTIFIER & PLN
    # --------------------------------------------------------
    power_values = [
        power["ID PLN"],
        power["Daya PLN"],
        power["Phase"],
        power["MCB"],
        power["Voltage R"],
        power["Voltage S"],
        power["Voltage T"],
        power["Utility PLN"],
        power["Capacity Status"],
        power["Phase Balanced"],
    ]

    power_y = [
        205,
        250,
        295,
        340,
        385,
        430,
        475,
        520,
        565,
        610,
    ]

    for text, y in zip(power_values, power_y):
        draw_value(
            draw,
            text,
            POWER_X * sx,
            y * sy,
            POWER_W * sx,
            size=16,
        )

    # --------------------------------------------------------
    # BATTERY STATUS
    # --------------------------------------------------------
    battery_values = [
        battery["Battery Type"],
        battery["Battery Brand"],
        battery["Battery Capacity"],
        battery["Battery Bank"],
        battery["BBT H"],
        battery["BBT Category"],
        battery["Battery Install"],
    ]

    battery_y = [
        695,
        740,
        785,
        830,
        875,
        920,
        965,
    ]

    # Clip rows that are outside a particular mockup version.
    for text, y in zip(battery_values, battery_y):
        if y * sy >= img.height - 8:
            continue

        draw_value(
            draw,
            text,
            POWER_X * sx,
            y * sy,
            POWER_W * sx,
            size=15,
        )

    # --------------------------------------------------------
    # HEALTHY CHECK
    # --------------------------------------------------------
    health_values = [
        health["EAS Validation"],
        health["NETECO Status"],
        health["Rectifier Condition"],
        health["Capacity Status"],
        health["Phase Balanced"],
        health["Rectifier Utility"],
        health["Last Check"],
    ]

    health_y = [
        205,
        255,
        305,
        355,
        405,
        455,
        505,
    ]

    for text, y in zip(health_values, health_y):
        draw_value(
            draw,
            text,
            HEALTH_X * sx,
            y * sy,
            HEALTH_W * sx,
            size=15,
        )

    # --------------------------------------------------------
    # ACTION
    # --------------------------------------------------------
    action_y = 690

    for action in actions[:5]:
        draw_value(
            draw,
            "• " + action,
            ACTION_X * sx,
            action_y * sy,
            ACTION_W * sx,
            size=16,
            color=GREEN if action == "No action required" else RED,
        )
        action_y += 34

    safe = re.sub(
        r"[^A-Za-z0-9._-]+",
        "_",
        str(site_id).upper(),
    )

    output = f"output_area_{safe}.png"
    img.save(output, format="PNG", dpi=(150, 150))

    return output


def register_area_handler(bot):
    @bot.message_handler(commands=["area"])
    def handle_area(message):
        args = message.text.split()

        if len(args) < 2:
            bot.reply_to(
                message,
                "⚠️ Format:\n`/area <SITE_ID>`",
                parse_mode="Markdown",
            )
            return

        site_id = args[1].strip().upper()

        status = bot.reply_to(
            message,
            f"⏳ Memproses Area 2: *{site_id}*...",
            parse_mode="Markdown",
        )

        try:
            output = render_area(site_id)

            if not output or not os.path.exists(output):
                bot.edit_message_text(
                    f"❌ Site ID *{site_id}* tidak ditemukan "
                    f"di sheet Area 2.",
                    message.chat.id,
                    status.message_id,
                    parse_mode="Markdown",
                )
                return

            with open(output, "rb") as photo:
                bot.send_photo(
                    message.chat.id,
                    photo,
                    caption=f"✅ Area 2 Report: *{site_id}*",
                    parse_mode="Markdown",
                )

            os.remove(output)

            try:
                bot.delete_message(
                    message.chat.id,
                    status.message_id,
                )
            except Exception:
                pass

        except Exception as exc:
            print(f"Area error {site_id}: {exc}")

            try:
                bot.edit_message_text(
                    f"❌ Error Area 2 *{site_id}*:\n`{exc}`",
                    message.chat.id,
                    status.message_id,
                    parse_mode="Markdown",
                )
            except Exception:
                bot.reply_to(
                    message,
                    f"❌ Error Area 2 *{site_id}*.",
                    parse_mode="Markdown",
                )
