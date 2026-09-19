import os
import glob
import re
import telebot
import pandas as pd
from PIL import Image, ImageDraw, ImageFont

TOKEN = os.getenv('TOKEN')
if not TOKEN:
    raise ValueError('Token belum diset di Environment Variables Railway!')

bot = telebot.TeleBot(TOKEN)

# ============================================================
# FONT
# ============================================================
def find_font(preferred=None, emoji=False, size=18):
    candidates = []
    if preferred:
        candidates.append(preferred)

    # Font yang diletakkan di root repository akan ikut dicari.
    candidates += glob.glob('*.ttf') + glob.glob('*.TTF')
    candidates += [
        '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
        '/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf',
        'DejaVuSans.ttf',
        'arial.ttf',
    ]

    if emoji:
        # Prioritaskan nama font yang biasanya berisi glyph emoji/symbol.
        candidates = sorted(
            candidates,
            key=lambda p: 0 if any(k in os.path.basename(p).lower()
                                   for k in ['emoji', 'symbol', 'segui', 'noto']) else 1
        )

    for path in candidates:
        if path and os.path.exists(path):
            try:
                return ImageFont.truetype(path, size)
            except Exception:
                pass

    return ImageFont.load_default()


FONT_PATH = os.getenv('FONT_PATH') or None
EMOJI_FONT_PATH = os.getenv('EMOJI_FONT_PATH') or None


# ============================================================
# HELPER
# ============================================================
def safe_float(value):
    try:
        if pd.isna(value):
            return None
        text = str(value).strip().replace(',', '.')
        return float(text)
    except Exception:
        return None


def clean_filename(text):
    return re.sub(r'[^A-Za-z0-9._-]+', '_', str(text))


def generate_site_card(site_id):
    excel_path = 'Excel_master.xlsx'
    mockup_path = 'Mokup.png'

    if not os.path.exists(excel_path):
        print('Excel_master.xlsx tidak ditemukan')
        return None
    if not os.path.exists(mockup_path):
        print('Mokup.png tidak ditemukan')
        return None

    wanted_id = str(site_id).strip().upper()
    site_data = None

    try:
        xls = pd.ExcelFile(excel_path)
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(excel_path, sheet_name=sheet_name)
            id_col = None

            for col in df.columns:
                name = str(col).lower()
                if 'site' in name and 'id' in name:
                    id_col = col
                    break

            if id_col is not None:
                ids = df[id_col].astype(str).str.strip().str.upper()
                filtered = df[ids == wanted_id]
                if not filtered.empty:
                    site_data = filtered.iloc[0]
                    break
    except Exception as exc:
        print(f'Error membaca Excel: {exc}')
        return None

    if site_data is None:
        return None

    img = Image.open(mockup_path).convert('RGB')
    draw = ImageDraw.Draw(img)
    W, H = img.size

    # ========================================================
    # UKURAN FONT UNTUK MOKUP 1479 x 772
    # Jangan pakai 65 px. Itu terlalu besar untuk canvas ini.
    # ========================================================
    BASE_W = 1479
    scale = W / BASE_W
    FONT_SIZE = max(16, round(19 * scale))
    EMOJI_SIZE = max(18, round(21 * scale))
    SMALL_FONT_SIZE = max(14, round(17 * scale))
    LINE_GAP = max(28, round(31 * scale))

    font = find_font(FONT_PATH, emoji=False, size=FONT_SIZE)
    small_font = find_font(FONT_PATH, emoji=False, size=SMALL_FONT_SIZE)
    emoji_font = find_font(EMOJI_FONT_PATH, emoji=True, size=EMOJI_SIZE)

    COLOR_LABEL = (80, 80, 80)
    COLOR_TEXT = (30, 30, 30)
    COLOR_GREEN = (16, 124, 65)
    COLOR_BLUE = (0, 120, 212)
    COLOR_RED = (209, 52, 56)

    def val(col, default='-'):
        if col not in site_data:
            return default
        v = site_data.get(col)
        if pd.notnull(v) and str(v).strip() not in ['', 'nan', 'None']:
            return str(v).strip()
        return default

    def get_dynamic_color(text):
        t = str(text).upper()
        if any(w in t for w in [
            'DOWN', 'CRITICAL', 'NO BACKUP', 'NOT AVAILABLE',
            'NEED VALIDATION', 'NOT_AVAILABLE', 'POTENSIAL TRIP',
            'TRIP', 'NEED', 'PERGANTIAN'
        ]):
            return COLOR_RED
        if any(w in t for w in [
            'NORMAL', 'SECURED', 'VALID', 'OK', 'AVAILABLE',
            'VIP', 'MONITOR'
        ]):
            return COLOR_GREEN
        if any(w in t for w in ['SILVER', 'GOLD', 'LITHIUM', 'HUAWEI', 'TELKOMSEL', 'TELKOM']):
            return COLOR_BLUE
        return COLOR_TEXT

    # ========================================================
    # ACTION
    # ========================================================
    action_list = []
    for col_name in ['Action', 'Activity (SOW) Actual', 'SOW']:
        if col_name in site_data:
            action = site_data.get(col_name)
            if pd.notnull(action) and str(action).strip() not in ['', '-', 'nan']:
                action_list.append(str(action).strip())
                break

    if not action_list:
        rect_cond = val('Rectifier Condition').upper()
        pot_trip = val('Potensial Trip').upper()
        eas_val = val('EAS Validation').upper()
        neteco = val('NETECO Status').upper()

        if 'DOWN' in rect_cond or 'CRITICAL' in rect_cond:
            action_list.append('NEED REPLACE RECTIFIER')
        if 'TRIP' in pot_trip or 'POTENSIAL' in pot_trip:
            action_list.append('NEED CHECK LOAD')
        if 'NEED' in eas_val or 'VALIDATION' in eas_val:
            action_list.append('NEED VALIDATE')
        if 'NOT AVAILABLE' in neteco or 'DOWN' in neteco:
            action_list.append('NEED CHECK NETECO')
        if not action_list:
            action_list = ['NORMAL']

    # ========================================================
    # DATA
    # ========================================================
    bbt = safe_float(site_data.get('BBT H (1)')) if 'BBT H (1)' in site_data else None
    bbt_str = f'{bbt:.2f} Hours' if bbt is not None else '-'

    util = safe_float(site_data.get('Rectifier Utility')) if 'Rectifier Utility' in site_data else None
    util_str = f'{util * 100:.1f} %' if util is not None else '-'

    col_site = [
        ('🆔', 'Site ID', val('Site ID')),
        ('🏷️', 'Site Name', val('Site Name')),
        ('🌐', 'Regional', val('Regional')),
        ('📡', 'NOP', val('NOP_1')),
        ('📍', 'TO', val('TO')),
        ('👤', 'ROH', val('ROH')),
        ('🏢', 'Site Owner', val('Site Owner')),
        ('📌', 'Lat / Long', f"{val('Lat')} / {val('Long')}"),
    ]

    col_rect = [
        ('🆔', 'ID PLN', val('ID PLN')),
        ('⚡', 'Daya PLN', f"{val('Daya PLN (KVA)')} kVA"),
        ('🔌', 'Brand', val('Rectifier Brand (1)')),
        ('⚙️', 'Model', val('Rectifier Model (1)')),
        ('📊', 'Capacity', val('Module Capacity (1)')),
        ('🔢', 'Module Qty', val('Inserted Module Qty (1)')),
        ('📈', 'Load System', val('Load System (1)')),
    ]

    col_batt = [
        ('🏷️', 'Brand', val('Battery Brand (1)')),
        ('🔋', 'Type', val('Battery Type (1)')),
        ('📊', 'Capacity', val('Battery Capacity (1)')),
        ('🔢', 'Bank Qty', val('Battery Bank (1)')),
        ('⏱️', 'BBT Backup', bbt_str),
        ('📂', 'Category', val('BBT Category (1)')),
    ]

    col_health = [
        ('🩺', 'Rect. Cond', val('Rectifier Condition')),
        ('📶', 'Utility', util_str),
        ('🛡️', 'Config', val('Rectifier Config (Category)')),
        ('📊', 'Cap. Status', val('Capacity Status')),
        ('⚠️', 'Pot. Trip', val('Potensial Trip')),
        ('🛠️', 'SOW Act.', val('Activity (SOW) Actual')),
        ('✅', 'EAS Valid.', val('EAS Validation')),
        ('🌐', 'NETECO Stat', val('NETECO Status')),
    ]

    # ========================================================
    # LAYOUT
    # Berdasarkan Mokup 1479 x 772.
    # ========================================================
    boxes = {
        'site':   {'x_icon': 42,  'x_label': 70,  'x_colon': 165,  'x_value': 182,  'y': 198, 'max_value': 175},
        'rect':   {'x_icon': 442, 'x_label': 470, 'x_colon': 600,  'x_value': 618,  'y': 198, 'max_value': 220},
        'batt':   {'x_icon': 442, 'x_label': 470, 'x_colon': 600,  'x_value': 618,  'y': 493, 'max_value': 220},
        'health': {'x_icon': 832, 'x_label': 860, 'x_colon': 1012, 'x_value': 1030, 'y': 198, 'max_value': 300},
    }

    def fit_font(text, max_width, start_font=FONT_SIZE, minimum=12):
        size = start_font
        while size > minimum:
            f = find_font(FONT_PATH, emoji=False, size=size)
            bbox = draw.textbbox((0, 0), str(text), font=f)
            if bbox[2] - bbox[0] <= max_width:
                return f
            size -= 1
        return find_font(FONT_PATH, emoji=False, size=minimum)

    def draw_row(cfg, y, icon, label, value):
        # Emoji/icon di font khusus; text selalu pakai font biasa.
        icon_bbox = draw.textbbox((0, 0), icon, font=emoji_font)
        icon_h = icon_bbox[3] - icon_bbox[1]
        draw.text(
            (cfg['x_icon'], y - icon_h / 2),
            icon,
            fill=COLOR_TEXT,
            font=emoji_font,
            anchor='lm'
        )

        draw.text((cfg['x_label'], y), label, fill=COLOR_LABEL, font=font, anchor='lm')
        draw.text((cfg['x_colon'], y), ':', fill=COLOR_TEXT, font=font, anchor='lm')

        value_font = fit_font(value, cfg['max_value'])
        draw.text(
            (cfg['x_value'], y),
            str(value),
            fill=get_dynamic_color(value),
            font=value_font,
            anchor='lm'
        )

    def render_rows(cfg, items):
        y = cfg['y']
        for icon, label, value in items:
            draw_row(cfg, y, icon, label, value)
            y += LINE_GAP
        return y

    render_rows(boxes['site'], col_site)
    render_rows(boxes['rect'], col_rect)
    render_rows(boxes['batt'], col_batt)
    health_end = render_rows(boxes['health'], col_health)

    # ========================================================
    # ACTION
    # ========================================================
    action_y = health_end + 4
    draw_row(boxes['health'], action_y, '🎯', 'Action', action_list[0])

    # Action tambahan dibuat di baris berikutnya.
    for action in action_list[1:]:
        action_y += LINE_GAP
        draw_row(boxes['health'], action_y, '  ', '', action)

    output_path = f"output_{clean_filename(site_id)}.png"
    img.save(output_path, dpi=(300, 300), quality=95)
    return output_path


@bot.message_handler(commands=['site'])
def handle_site(message):
    args = message.text.split()
    if len(args) < 2:
        bot.reply_to(message, '⚠️ Format salah! Gunakan: `/site <Site_ID>`', parse_mode='Markdown')
        return

    site_id = args[1]
    bot.reply_to(message, f'⏳ Sedang memproses Site ID: *{site_id}*...', parse_mode='Markdown')

    try:
        img_path = generate_site_card(site_id)

        if img_path and os.path.exists(img_path):
            with open(img_path, 'rb') as doc_file:
                bot.send_document(
                    message.chat.id,
                    doc_file,
                    caption=f'✅ Status Report Site ID: *{site_id}*',
                    parse_mode='Markdown'
                )
            os.remove(img_path)
        else:
            bot.reply_to(message, f'❌ Maaf, Site ID *{site_id}* tidak ditemukan.', parse_mode='Markdown')
    except Exception as exc:
        print(f'Error generate report {site_id}: {exc}')
        bot.reply_to(message, f'❌ Terjadi error saat membuat report untuk *{site_id}*.', parse_mode='Markdown')


print('Bot Telegram siap dijalankan...')
bot.infinity_polling(skip_pending=True)
