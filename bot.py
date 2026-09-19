import os
import glob
import re
import textwrap
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
def find_font(preferred=None, size=18, bold=False):
    candidates = []
    if preferred:
        candidates.append(preferred)

    # Jika TTF diletakkan di root repository, ikut dicari.
    if bold:
        candidates += glob.glob('*Bold*.ttf') + glob.glob('*bold*.ttf')
        candidates += [
            '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf',
            '/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf',
            '/usr/share/fonts/truetype/msttcorefonts/Arial_Bold.ttf',
        ]
    else:
        candidates += glob.glob('*.ttf') + glob.glob('*.TTF')
        candidates += [
            '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
            '/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf',
            'DejaVuSans.ttf',
            'arial.ttf',
        ]

    # Hilangkan duplikat sambil mempertahankan urutan prioritas.
    seen = set()
    for path in candidates:
        if not path or path in seen:
            continue
        seen.add(path)
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size)
            except Exception:
                pass

    return ImageFont.load_default()


FONT_PATH = os.getenv('FONT_PATH') or None
BOLD_FONT_PATH = os.getenv('BOLD_FONT_PATH') or os.getenv('FONT_BOLD_PATH') or None


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


def clean_value(value, default='-'):
    if value is None:
        return default
    try:
        if pd.isna(value):
            return default
    except Exception:
        pass
    text = str(value).strip()
    if text.lower() in ('', 'nan', 'none', 'nat'):
        return default
    return text


def wrap_text(draw, text, font, max_width):
    """Wrap berdasarkan ukuran pixel, bukan jumlah karakter."""
    text = str(text)
    if not text:
        return ['-']

    words = text.split()
    lines = []
    current = ''

    for word in words:
        trial = word if not current else current + ' ' + word
        bbox = draw.textbbox((0, 0), trial, font=font)
        if bbox[2] - bbox[0] <= max_width:
            current = trial
        else:
            if current:
                lines.append(current)
            current = word

    if current:
        lines.append(current)

    # Kalau satu kata saja lebih panjang dari area, biarkan fit_font yang mengecilkan.
    return lines or [text]


def get_dynamic_color(text):
    t = str(text).upper()

    if any(w in t for w in [
        'DOWN', 'CRITICAL', 'NO BACKUP', 'NOT AVAILABLE',
        'NEED VALIDATION', 'NOT_AVAILABLE', 'POTENSIAL TRIP',
        'TRIP', 'NEED', 'PERGANTIAN', 'REPLACE'
    ]):
        return (211, 47, 47)

    if any(w in t for w in [
        'NORMAL', 'SECURED', 'VALID', 'OK', 'AVAILABLE', 'MONITOR'
    ]):
        return (18, 137, 72)

    # Value umum dibuat navy agar konsisten dengan mockup.
    return (24, 55, 96)


# ============================================================
# GENERATE REPORT
# ============================================================
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
    # MOCKUP FINAL: 1672 x 941
    # Jangan menggambar ulang icon/label karena semuanya sudah
    # menjadi bagian dari Mokup.png.
    # Script hanya mengisi VALUE dan Data Update.
    # ========================================================
    BASE_W, BASE_H = 1672, 941
    sx = W / BASE_W
    sy = H / BASE_H

    # Font value dibuat bold agar sama dengan contoh output yang diinginkan.
    VALUE_SIZE = round(21 * sx)
    VALUE_SMALL = round(18 * sx)
    FOOTER_SIZE = round(18 * sx)

    value_font = find_font(BOLD_FONT_PATH, size=VALUE_SIZE, bold=True)
    footer_font = find_font(BOLD_FONT_PATH, size=FOOTER_SIZE, bold=True)

    def val(col, default='-'):
        return clean_value(site_data.get(col), default) if col in site_data else default

    # ========================================================
    # DATA
    # ========================================================
    bbt = safe_float(site_data.get('BBT H (1)')) if 'BBT H (1)' in site_data else None
    bbt_str = f'{bbt:.2f} Hours' if bbt is not None else '-'

    util = safe_float(site_data.get('Rectifier Utility')) if 'Rectifier Utility' in site_data else None
    util_str = f'{util * 100:.1f} %' if util is not None else '-'

    col_site = [
        ('Site ID', val('Site ID')),
        ('Site Name', val('Site Name')),
        ('Regional', val('Regional')),
        ('NOP', val('NOP_1')),
        ('TO', val('TO')),
        ('ROH', val('ROH')),
        ('Site Owner', val('Site Owner')),
        ('Lat / Long', f"{val('Lat')} / {val('Long')}"),
    ]

    col_rect = [
        ('ID PLN', val('ID PLN')),
        ('Daya PLN', f"{val('Daya PLN (KVA)')} kVA"),
        ('Brand', val('Rectifier Brand (1)')),
        ('Model', val('Rectifier Model (1)')),
        ('Capacity', val('Module Capacity (1)')),
        ('Module Qty', val('Inserted Module Qty (1)')),
        ('Load System', val('Load System (1)')),
    ]

    col_batt = [
        ('Brand', val('Battery Brand (1)')),
        ('Type', val('Battery Type (1)')),
        ('Capacity', val('Battery Capacity (1)')),
        ('Bank Qty', val('Battery Bank (1)')),
        ('BBT Backup', bbt_str),
        ('Category', val('BBT Category (1)')),
    ]

    col_health = [
        ('Rect. Cond', val('Rectifier Condition')),
        ('Utility', util_str),
        ('Config', val('Rectifier Config (Category)')),
        ('Cap. Status', val('Capacity Status')),
        ('Pot. Trip', val('Potensial Trip')),
        ('SOW Act.', val('Activity (SOW) Actual')),
        ('EAS Valid.', val('EAS Validation')),
        ('NETECO Stat', val('NETECO Status')),
    ]

    # ========================================================
    # ACTION
    # ========================================================
    action_list = []
    for col_name in ['Action', 'Activity (SOW) Actual', 'SOW']:
        if col_name in site_data:
            action = clean_value(site_data.get(col_name), '')
            if action not in ('', '-'):
                action_list.append(action)
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
    # KOORDINAT VALUE SAJA
    # Koordinat mengikuti titik ':' yang sudah ada di mockup.
    # ========================================================
    boxes = {
        'site': {
            'x': 273, 'y': 236, 'gap': 54,
            'max_width': 185, 'max_height': 42,
        },
        'rect': {
            'x': 780, 'y': 211, 'gap': 45,
            'max_width': 285, 'max_height': 40,
        },
        'batt': {
            'x': 780, 'y': 609, 'gap': 40,
            'max_width': 285, 'max_height': 38,
        },
        'health': {
            'x': 1375, 'y': 229, 'gap': 54,
            'max_width': 215, 'max_height': 42,
        },
    }

    # Scale koordinat jika gambar ternyata berubah ukuran.
    for cfg in boxes.values():
        cfg['x'] = round(cfg['x'] * sx)
        cfg['y'] = round(cfg['y'] * sy)
        cfg['gap'] = round(cfg['gap'] * sy)
        cfg['max_width'] = round(cfg['max_width'] * sx)
        cfg['max_height'] = round(cfg['max_height'] * sy)

    def fit_bold_font(text, max_width, start_size=VALUE_SIZE, minimum=13):
        size = max(start_size, minimum)
        while size > minimum:
            f = find_font(BOLD_FONT_PATH, size=size, bold=True)
            bbox = draw.textbbox((0, 0), str(text), font=f)
            if bbox[2] - bbox[0] <= max_width:
                return f
            size -= 1
        return find_font(BOLD_FONT_PATH, size=minimum, bold=True)

    def draw_value(x, y, value, max_width, center_y=True):
        text = clean_value(value)
        color = get_dynamic_color(text)
        f = fit_bold_font(text, max_width)
        bbox = draw.textbbox((0, 0), text, font=f)
        text_h = bbox[3] - bbox[1]
        yy = y - text_h / 2 if center_y else y
        draw.text((x, yy), text, fill=color, font=f)

    def draw_wrapped_value(x, y, value, max_width, max_lines=2, line_gap=4):
        text = clean_value(value)
        # Coba satu baris dulu.
        f = fit_bold_font(text, max_width)
        bbox = draw.textbbox((0, 0), text, font=f)
        if bbox[2] - bbox[0] <= max_width and len(text) <= 34:
            draw.text((x, y - (bbox[3] - bbox[1]) / 2), text,
                      fill=get_dynamic_color(text), font=f)
            return

        # Untuk SOW/Action panjang, wrap maksimal 2 baris dengan font tetap terbaca.
        f = find_font(BOLD_FONT_PATH, size=max(VALUE_SMALL, 15), bold=True)
        lines = wrap_text(draw, text, f, max_width)
        lines = lines[:max_lines]
        if len(lines) == max_lines and len(wrap_text(draw, text, f, max_width)) > max_lines:
            last = lines[-1]
            while last and draw.textbbox((0, 0), last + '...', font=f)[2] - draw.textbbox((0, 0), last + '...', font=f)[0] > max_width:
                last = last[:-1]
            lines[-1] = (last.rstrip() + '...') if last else '...'

        line_h = f.size + line_gap
        start_y = y - ((len(lines) - 1) * line_h) / 2
        for i, line in enumerate(lines):
            draw.text((x, start_y + i * line_h), line,
                      fill=get_dynamic_color(text), font=f)

    # Site / Rectifier / Battery: satu baris per field.
    for i, (_, value) in enumerate(col_site):
        draw_value(boxes['site']['x'], boxes['site']['y'] + i * boxes['site']['gap'],
                   value, boxes['site']['max_width'])

    for i, (_, value) in enumerate(col_rect):
        draw_value(boxes['rect']['x'], boxes['rect']['y'] + i * boxes['rect']['gap'],
                   value, boxes['rect']['max_width'])

    for i, (_, value) in enumerate(col_batt):
        draw_value(boxes['batt']['x'], boxes['batt']['y'] + i * boxes['batt']['gap'],
                   value, boxes['batt']['max_width'])

    # Health. SOW dibuat wrap agar tidak menabrak tepi kanan.
    for i, (_, value) in enumerate(col_health):
        y = boxes['health']['y'] + i * boxes['health']['gap']
        if i == 5:  # SOW Act.
            draw_wrapped_value(boxes['health']['x'], y, value,
                               boxes['health']['max_width'], max_lines=2)
        else:
            draw_value(boxes['health']['x'], y, value, boxes['health']['max_width'])

    # Action di area kosong bawah health panel.
    action_y = boxes['health']['y'] + len(col_health) * boxes['health']['gap']
    action_y += round(2 * sy)
    draw_wrapped_value(boxes['health']['x'], action_y, action_list[0],
                       boxes['health']['max_width'], max_lines=3)

    # ========================================================
    # FOOTER DATA UPDATE
    # ========================================================
    update_value = '-'
    for col_name in ['Data Update', 'DATA UPDATE', 'Update Date', 'Last Update', 'Last Check Update']:
        if col_name in site_data:
            update_value = clean_value(site_data.get(col_name))
            break

    if update_value == '-':
        # Jangan menulis tanggal hari ini secara otomatis jika Excel tidak punya data.
        update_value = '-'

    # Label 'Data Update :' sudah ada di mockup. Jangan gambar ulang
    # label tersebut karena akan menghasilkan teks dobel. Jika Excel
    # punya tanggal update, hanya area nilai setelah ':' yang ditimpa.
    if update_value != '-':
        # Mockup sudah punya label dan tanda '-'. Hapus hanya tanda '-'
        # dengan warna background lokal, lalu isi tanggal sesudah ':' .
        cover_color = img.getpixel((round(220 * sx), round(892 * sy)))
        draw.rectangle(
            [round(188 * sx), round(878 * sy), round(216 * sx), round(906 * sy)],
            fill=cover_color
        )
        draw.text((round(196 * sx), round(891 * sy)), str(update_value),
                  fill=(255, 255, 255), font=footer_font)

    output_path = f'output_{clean_filename(site_id)}.png'
    img.save(output_path, dpi=(300, 300), quality=96)
    return output_path


# ============================================================
# TELEGRAM
# ============================================================
@bot.message_handler(commands=['site'])
def handle_site(message):
    args = message.text.split()
    if len(args) < 2:
        bot.reply_to(message, '⚠️ Format salah! Gunakan: `/site <Site_ID>`', parse_mode='Markdown')
        return

    site_id = args[1]
    bot.reply_to(message,
                 f'⏳ Sedang memproses Site ID: *{site_id}*...',
                 parse_mode='Markdown')

    try:
        img_path = generate_site_card(site_id)

        if img_path and os.path.exists(img_path):
            with open(img_path, 'rb') as photo:
                bot.send_photo(
                    message.chat.id,
                    photo,
                    caption=f'✅ Status Report Site ID: *{site_id}*',
                    parse_mode='Markdown'
                )
            os.remove(img_path)
        else:
            bot.reply_to(message,
                         f'❌ Maaf, Site ID *{site_id}* tidak ditemukan.',
                         parse_mode='Markdown')
    except Exception as exc:
        print(f'Error generate report {site_id}: {exc}')
        bot.reply_to(message,
                     f'❌ Terjadi error saat membuat report untuk *{site_id}*.',
                     parse_mode='Markdown')


print('Bot Telegram siap dijalankan...')
bot.infinity_polling(skip_pending=True)
