import os
import re
import pandas as pd
from PIL import Image, ImageDraw, ImageFont

# ============================================================
# GENSET MODULE — V5 FIXED LAYOUT
# ------------------------------------------------------------
# Module ini sengaja berdiri sendiri dari renderer Power/Rectifier.
# Tidak mengubah /site atau logic Power.
#
# Source yang dipakai:
#   Genset Autorate   : I, J, K, Q
#   Genset WarmingUp  : I, M
#   BBM GensetFix     : T, X, Y
# ============================================================

EXCEL_CANDIDATES = [
    "Excel_master(1).xlsx",
    "Excel_master.xlsx",
]

MOCKUP_CANDIDATES = [
    "Mokupgensetnew.png",
    "MokupGenset.png",
    "mokupgenset.png",
]

# Reference = ukuran Mokupgenset.png yang sekarang.
# Kalau ukuran mockup berubah, renderer mengikuti ukuran file aktual.
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

    return str(v).strip().lower() in (
        "",
        "nan",
        "none",
        "-",
        "nat",
    )


def display(v, default=""):
    if is_empty(v):
        return default
    return str(v).strip()


def get_font(size, bold=True):
    paths = (
        [
            "/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
            "LiberationSans-Bold.ttf",
            "DejaVuSans-Bold.ttf",
        ]
        if bold
        else
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
    text = str(text)

    for n in range(int(size), int(minimum) - 1, -1):
        font = get_font(n, bold=bold)
        bbox = draw.textbbox((0, 0), text, font=font)

        if (bbox[2] - bbox[0]) <= max_width:
            return font

    return get_font(minimum, bold=bold)


def shorten(draw, text, max_width, font):
    text = str(text)

    if draw.textbbox((0, 0), text, font=font)[2] <= max_width:
        return text

    suffix = "..."

    lo = 0
    hi = len(text)

    while lo < hi:
        mid = (lo + hi + 1) // 2
        candidate = text[:mid].rstrip() + suffix

        if draw.textbbox((0, 0), candidate, font=font)[2] <= max_width:
            lo = mid
        else:
            hi = mid - 1

    return text[:lo].rstrip() + suffix


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

    kept = lines[:max_lines - 1]
    remainder = " ".join(lines[max_lines - 1:])

    while remainder:
        candidate = remainder.rstrip() + "..."

        if draw.textbbox((0, 0), candidate, font=font)[2] <= max_width:
            break

        remainder = remainder[:-1].rstrip()

    kept.append((remainder + "...") if remainder else "...")
    return kept


# ============================================================
# COLOR RULE
# ------------------------------------------------------------
# User rule:
#   Normal       -> GREEN
#   Safe         -> GREEN
#   Warning      -> RED
#   UnMonitor    -> RED
#   Unavailable  -> RED
#   0-60%        -> GREEN
#   60-90%       -> RED
#
# Site Info tetap NAVY supaya identitas tidak ikut berubah warna.
# ============================================================

def dynamic_color(value):
    text = str(value).strip()
    upper = text.upper()

    if not text:
        return NAVY

    # Explicit red states.
    red_words = (
        "WARNING",
        "UNMONITOR",
        "UNAVAILABLE",
        "NOT AVAILABLE",
        "NOT_AVAILABLE",
        "NOT AUTO",
        "BROKEN",
        "CRITICAL",
        "DOWN",
        "FAULT",
        "FAILED",
        "ERROR",
        "NOT OK",
        "NOT SAFE",
        "UNBALANCE",
        "UNBALANCED",
        "PROBLEM",
        "NEED CHECK",
        "NEED VALIDATION",
        "VALIDATION",
        "PROPOSED",
        "WAITING",
        "POWER OFF",
        "NOT MONITOR",
    )

    if any(word in upper for word in red_words):
        return RED

    # Percentage rule.
    # Hanya diterapkan bila value memang bertanda %.
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

    # Explicit green states.
    green_words = (
        "NORMAL",
        "SAFE",
        "ACTIVE",
        "AUTO",
        "OK",
        "SECURED",
        "VALID",
        "AVAILABLE",
        "MONITOR",
        "MONITORING",
        "BALANCE",
        "BALANCED",
        "CLOSED",
    )

    if any(word in upper for word in green_words):
        return GREEN

    # Warning/check state tetap merah sesuai rule terbaru.
    if any(word in upper for word in ("CHECK", "LOW", "MEDIUM")):
        return RED

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

    ids = (
        df["Site ID"]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    found = df.loc[ids == wanted]

    if found.empty:
        return None

    return found.iloc[0]


def row_value(row, column, default=""):
    if row is None or column not in row.index:
        return default

    return display(row[column], default)


def load_genset_data(site_id):
    autorate = find_site(
        load_sheet("Genset Autorate"),
        site_id,
    )

    warming = find_site(
        load_sheet("Genset WarmingUp"),
        site_id,
    )

    bbm = find_site(
        load_sheet("BBM GensetFix"),
        site_id,
    )

    if autorate is None and warming is None and bbm is None:
        return None

    return autorate, warming, bbm


# ============================================================
# SITE INFO
# ------------------------------------------------------------
# Genset Autorate menjadi source utama identitas:
# Site ID / Site Name / Regional / NOP / TO.
#
# ROH / Site Owner / Lat / Long tidak tersedia di Autorate,
# sehingga dicari dari Rectifire&battery sebagai metadata saja.
# Tidak mengubah renderer Power.
# ============================================================

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

    # Primary identity source = Genset Autorate.
    try:
        autorate = find_site(
            load_sheet("Genset Autorate"),
            site_id,
        )

        if autorate is not None:
            info["Site ID"] = row_value(
                autorate,
                "Site ID",
                str(site_id).upper(),
            )
            info["Site Name"] = row_value(
                autorate,
                "Site Name",
            )
            info["Regional"] = row_value(
                autorate,
                "Region",
            )
            info["NOP"] = row_value(
                autorate,
                "NOP",
            )
            info["TO"] = row_value(
                autorate,
                "TO",
            )
    except Exception:
        pass

    # Metadata fallback only.
    try:
        power = find_site(
            load_sheet("Rectifire&battery"),
            site_id,
        )

        if power is not None:
            if not info["Site Name"]:
                info["Site Name"] = row_value(power, "Site Name")

            if not info["Regional"]:
                info["Regional"] = row_value(
                    power,
                    "Regional",
                )

            if not info["NOP"]:
                info["NOP"] = row_value(power, "NOP")

            if not info["TO"]:
                info["TO"] = row_value(power, "TO")

            info["ROH"] = row_value(power, "ROH")
            info["Site Owner"] = row_value(
                power,
                "Site Owner",
            )

            lat = row_value(power, "Lat")
            lon = row_value(power, "Long")

            if lat and lon:
                info["Lat / Long"] = f"{lat} / {lon}"
    except Exception:
        pass

    # Last fallback for coordinates from BBM GensetFix.
    if not info["Lat / Long"]:
        try:
            bbm = find_site(
                load_sheet("BBM GensetFix"),
                site_id,
            )

            if bbm is not None:
                lat = row_value(bbm, "Lat")
                lon = row_value(bbm, "Long")

                if lat and lon:
                    info["Lat / Long"] = f"{lat} / {lon}"
        except Exception:
            pass

    return info


# ============================================================
# HEALTH / ACTION
# ============================================================

def is_bad(value):
    text = str(value).strip().upper()

    return any(
        word in text
        for word in (
            "BROKEN",
            "NOT AUTO",
            "NOT OK",
            "FAULT",
            "FAILED",
            "ERROR",
            "PROBLEM",
            "UNAVAILABLE",
            "DOWN",
            "UNMONITOR",
            "NOT SAFE",
        )
    )


def build_genset_health(autorate, warming, bbm):
    """
    Genset Health adalah hasil evaluasi sederhana dari source yang
    memang tersedia, bukan kolom Excel baru.

    Active + tidak ada kondisi rusak -> Active
    Selain itu -> Need Check
    """

    dg_status = row_value(
        autorate,
        "DG Status",
    )

    dg_condition = row_value(
        autorate,
        "DG Condition",
    )

    week_condition = row_value(
        autorate,
        "Current Week Genset Condition",
    )

    warm_status = row_value(
        warming,
        "Warming Up Weekly Status",
    )

    if (
        dg_status.upper() == "ACTIVE"
        and not is_bad(dg_condition)
        and not is_bad(week_condition)
        and not is_bad(warm_status)
    ):
        return "Active"

    return "Need Check"


def build_genset_actions(autorate, warming, bbm):
    actions = []

    dg_status = row_value(
        autorate,
        "DG Status",
    ).upper()

    dg_condition = row_value(
        autorate,
        "DG Condition",
    ).upper()

    week_condition = row_value(
        autorate,
        "Current Week Genset Condition",
    ).upper()

    warm_status = row_value(
        warming,
        "Warming Up Weekly Status",
    ).upper()

    warm_week = row_value(
        warming,
        "Current Week Genset Condition",
    ).upper()

    bbm_status = row_value(
        bbm,
        "Status",
    ).upper()

    bbm_advice = row_value(
        bbm,
        "Saran Pengisian",
    ).upper()

    # Autorate.
    if (
        dg_condition == "NOT AUTO"
        or dg_status not in ("ACTIVE", "OK")
    ):
        actions.append("Need Check Genset Auto")

    if is_bad(week_condition):
        actions.append("Need Check Genset Condition")

    # Warming Up.
    if (
        is_bad(warm_status)
        or is_bad(warm_week)
    ):
        actions.append("Need Check Genset Warming Up")

    # BBM only creates action when the source explicitly says
    # the fuel status/advice is not safe.
    if (
        "NOT SAFE" in bbm_status
        or "NOT SAFE" in bbm_advice
        or "UNAVAILABLE" in bbm_status
        or "UNAVAILABLE" in bbm_advice
    ):
        actions.append("Need Check BBM Genset")

    # Remove duplicates while preserving order.
    return list(dict.fromkeys(actions))


# ============================================================
# RENDER
# ============================================================

def generate_genset_card(site_id):
    data = load_genset_data(site_id)

    if data is None:
        return None

    autorate, warming, bbm = data

    mockup = find_existing(MOCKUP_CANDIDATES)

    if not mockup:
        raise FileNotFoundError(
            "Mokupgenset.png tidak ditemukan. "
            "Simpan mockup dengan nama Mokupgenset.png."
        )

    img = Image.open(mockup).convert("RGB")

    # Gunakan ukuran mockup aktual supaya tidak ada drift 1-2 px.
    sx = img.width / BASE_W
    sy = img.height / BASE_H

    draw = ImageDraw.Draw(img)

    def xy(x, y):
        return (
            round(x * sx),
            round(y * sy),
        )

    def write(
        text,
        x,
        y,
        width,
        size=20,
        color=None,
        bold=True,
        minimum=11,
    ):
        text = display(text)

        if not text:
            return

        pixel_width = round(width * sx)

        font = fit_font(
            draw,
            text,
            pixel_width,
            size=round(size * min(sx, sy)),
            minimum=minimum,
            bold=bold,
        )

        text = shorten(
            draw,
            text,
            pixel_width,
            font,
        )

        draw.text(
            xy(x, y),
            text,
            font=font,
            fill=color if color is not None else dynamic_color(text),
            anchor="lm",
        )

    def write_action(
        text,
        x,
        y,
        width,
        size=17,
        color=RED,
    ):
        """
        Action bukan row biasa.
        Tidak menggunakan anchor/value X yang sama dengan Healthy Check.
        """
        text = display(text)

        if not text:
            return

        pixel_width = round(width * sx)

        font = fit_font(
            draw,
            text,
            pixel_width,
            size=round(size * min(sx, sy)),
            minimum=12,
            bold=True,
        )

        text = shorten(
            draw,
            text,
            pixel_width,
            font,
        )

        draw.text(
            xy(x, y),
            text,
            font=font,
            fill=color,
            anchor="lm",
        )

    # --------------------------------------------------------
    # FIXED VALUE X
    # --------------------------------------------------------
    #
    # Mockup:
    # Site Info      ":" sekitar x=270 -> value x=280
    # Autorate       ":" sekitar x=800 -> value x=815
    # Warming Up     ":" sekitar x=800 -> value x=815
    # BBM            ":" sekitar x=800 -> value x=815
    # Healthy Check  ":" sekitar x=1390 -> value x=1405
    #
    # Semua row memakai X yang sama dalam card masing-masing.
    # --------------------------------------------------------

    # Site Info: posisi value dibuat konsisten tepat setelah ":".
    # Jangan terlalu jauh dari colon, dan jangan berubah-ubah antar row.
    # ========================================================
    # LOCKED VALUE COLUMNS
    # --------------------------------------------------------
    # Semua nilai dimulai dari X yang SAMA pada card masing-
    # masing. X tidak boleh mengikuti panjang label.
    # Posisi ini dikunci terhadap colon pada Mokupgenset.
    #
    # SITE INFO      colon ~ 238 -> value X 270
    # GENSET MID     colon ~ 798 -> value X 830
    # HEALTHY CHECK  colon ~ 1380 -> value X 1412
    # ACTION         text X 1230
    # ========================================================
    # ========================================================
    # V5 — VALUE ANCHORS LOCKED TO THE ACTUAL Mokupgenset.png
    # ========================================================
    # Colon positions measured from the current mockup:
    #   Site Info      ~ X 254
    #   Middle cards   ~ X 812
    #   Healthy Check  ~ X 1410
    #
    # Value starts at a fixed offset after the colon.
    # IMPORTANT: do NOT calculate X from label length.
    # This keeps every row vertically aligned.
    SITE_VALUE_X = 276
    SITE_VALUE_W = 190

    MID_VALUE_X = 834
    MID_VALUE_W = 235

    HEALTH_VALUE_X = 1432
    HEALTH_VALUE_W = 205

    # Action is a free-text list inside its own pink box.
    ACTION_VALUE_X = 1230
    ACTION_VALUE_W = 330

    # --------------------------------------------------------
    # SITE INFO
    # --------------------------------------------------------

    info = site_info(site_id)

    site_rows = [
        info.get("Site ID", ""),
        info.get("Site Name", ""),
        info.get("Regional", ""),
        info.get("NOP", ""),
        info.get("TO", ""),
        info.get("ROH", ""),
        info.get("Site Owner", ""),
        info.get("Lat / Long", ""),
    ]

    # Center tiap row mengikuti posisi ":" pada mockup.
    # Dibuat per-row karena spacing mockup Site Info tidak benar-benar
    # seragam, terutama pada ROH -> Site Owner -> Lat/Long.
    # Y juga dikunci ke center setiap row pada mockup.
    # Exact row centers from the current mockup.
    site_y = [
        231,  # Site ID
        293,  # Site Name
        355,  # Regional
        416,  # NOP
        478,  # TO
        539,  # ROH
        600,  # Site Owner
        661,  # Lat / Long
    ]

    for index, (text, y) in enumerate(
        zip(site_rows, site_y)
    ):
        # Identitas selalu NAVY.
        # Tidak ikut berubah hijau/merah hanya karena isi text.
        write(
            text,
            SITE_VALUE_X,
            y,
            SITE_VALUE_W,
            size=17 if index == 7 else (19 if index == 1 else 20),
            color=NAVY,
            minimum=12,
        )

    # --------------------------------------------------------
    # GENSET AUTORATE
    # Source: I, J, K, Q
    # --------------------------------------------------------

    autorate_rows = [
        row_value(autorate, "DG Status"),
        row_value(autorate, "DG Condition"),
        row_value(
            autorate,
            "Current Week Genset Condition",
        ),
        row_value(
            autorate,
            "Current Progress",
        ),
    ]

    autorate_y = [
        214,
        269,
        323,
        377,
    ]

    for text, y in zip(
        autorate_rows,
        autorate_y,
    ):
        write(
            text,
            MID_VALUE_X,
            y,
            MID_VALUE_W,
            size=18,
            minimum=12,
        )

    # --------------------------------------------------------
    # GENSET WARMING UP
    # Source: I, M
    # --------------------------------------------------------

    warming_rows = [
        row_value(
            warming,
            "Warming Up Weekly Status",
        ),
        row_value(
            warming,
            "Current Week Genset Condition",
        ),
    ]

    warming_y = [
        523,
        577,
    ]

    for text, y in zip(
        warming_rows,
        warming_y,
    ):
        write(
            text,
            MID_VALUE_X,
            y,
            MID_VALUE_W,
            size=18,
            minimum=12,
        )

    # --------------------------------------------------------
    # BBM GENSETFIX
    # Source: T, X, Y
    # --------------------------------------------------------

    bbm_rows = [
        row_value(
            bbm,
            "Perkiraan Sisa Fuel",
        ),
        row_value(
            bbm,
            "Status",
        ),
        row_value(
            bbm,
            "Saran Pengisian",
        ),
    ]

    bbm_y = [
        719,
        771,
        815,
    ]

    for text, y in zip(
        bbm_rows,
        bbm_y,
    ):
        write(
            text,
            MID_VALUE_X,
            y,
            MID_VALUE_W,
            size=17,
            minimum=12,
        )

    # --------------------------------------------------------
    # HEALTHY CHECK
    # --------------------------------------------------------
    #
    # Yang memang punya source dari scope Genset:
    #   Genset Health   -> hasil evaluasi
    #   DG Status       -> Autorate I
    #   DG Condition    -> Autorate J
    #   Current Week    -> Autorate K
    #   Current Progress-> Autorate Q
    #
    # Problem/RCA/Plan Action/PIC/Auto Date tidak dipaksa
    # masuk karena bukan bagian source yang kita sepakati.
    # --------------------------------------------------------

    genset_health = build_genset_health(
        autorate,
        warming,
        bbm,
    )

    health_rows = [
        genset_health,
        row_value(autorate, "DG Status"),
        row_value(autorate, "DG Condition"),
        row_value(
            autorate,
            "Current Week Genset Condition",
        ),
        "",  # Problem Genset
        "",  # RCA
        "",  # Plan Action
        row_value(
            autorate,
            "Current Progress",
        ),
        "",  # PIC
        "",  # Auto Date
    ]

    # 10 rows — one Y anchor for EVERY Healthy Check row.
    # The previous v4 had only 9 Y values, so Auto Date could
    # silently lose its intended position.
    health_y = [
        210,  # Genset Health
        254,  # DG Status
        298,  # DG Condition
        342,  # Current Week Condition
        386,  # Problem Genset
        430,  # RCA
        476,  # Plan Action
        523,  # Current Progress
        570,  # PIC
        616,  # Auto Date
    ]

    for text, y in zip(
        health_rows,
        health_y,
    ):
        if not text:
            continue

        write(
            text,
            HEALTH_VALUE_X,
            y,
            HEALTH_VALUE_W,
            size=16,
            minimum=11,
        )

    # --------------------------------------------------------
    # ACTION
    # --------------------------------------------------------

    actions = build_genset_actions(
        autorate,
        warming,
        bbm,
    )

    if not actions:
        actions = ["No action required"]
        action_color = GREEN
    else:
        action_color = RED

    # Action tidak lagi menggunakan x=1230 yang membuat text
    # terlihat seperti menabrak colon mockup.
    ACTION_X = ACTION_VALUE_X
    ACTION_W = ACTION_VALUE_W

    action_y = 765

    for action in actions[:4]:
        write_action(
            action,
            ACTION_X,
            action_y,
            ACTION_W,
            size=16,
            color=action_color,
        )
        action_y += 30

    # --------------------------------------------------------
    # SAVE
    # --------------------------------------------------------

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


# ============================================================
# TELEGRAM HANDLER
# ============================================================

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
            print(
                f"Error generate genset report "
                f"{site_id}: {exc}"
            )

            try:
                bot.edit_message_text(
                    f"❌ Error Genset *{site_id}*.\n"
                    f"`{exc}`",
                    message.chat.id,
                    status_msg.message_id,
                    parse_mode="Markdown",
                )
            except Exception:
                bot.reply_to(
                    message,
                    f"❌ Error saat membuat report Genset "
                    f"*{site_id}*.",
                    parse_mode="Markdown",
                )
