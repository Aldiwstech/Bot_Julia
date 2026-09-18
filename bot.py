import os
import telebot
from PIL import Image, ImageDraw, ImageFont
import pandas as pd

TOKEN = '8955027973:AAH1dm2tOBRMf85Pxy0N9MMKM-fw8ekrZnM'
bot = telebot.TeleBot(TOKEN)


def generate_site_card(site_id):
  excel_path = 'Excel_master.xlsx'
  if not os.path.exists(excel_path):
    return None

  xls = pd.ExcelFile(excel_path)
  filtered = pd.DataFrame()
  site_data = None

  for sheet_name in xls.sheet_names:
    df = pd.read_excel(excel_path, sheet_name=sheet_name)
    id_col = None
    for col in df.columns:
      if 'site' in str(col).lower() and 'id' in str(col).lower():
        id_col = col
        break

    if id_col:
      df[id_col] = df[id_col].astype(str).str.strip().str.upper()
      filtered = df[df[id_col] == str(site_id).strip().upper()]
      if not filtered.empty:
        site_data = filtered.iloc[0]
        break

  if site_data is None or filtered.empty:
    return None

  mockup_path = 'Mokup.png'
  if not os.path.exists(mockup_path):
    return None

  img = Image.open(mockup_path).convert('RGB')
  draw = ImageDraw.Draw(img)

  font_size = 21
  try:
    font = ImageFont.truetype('seguiemj.ttf', font_size)
  except:
    try:
      font = ImageFont.truetype('arial.ttf', font_size)
    except:
      font = ImageFont.load_default()

  COLOR_LABEL = (80, 80, 80)
  COLOR_TEXT = (20, 20, 20)
  COLOR_GREEN = (16, 124, 65)
  COLOR_BLUE = (0, 120, 212)
  COLOR_RED = (209, 52, 56)

  def val(col, default='-'):
    if col not in site_data:
      return default
    v = site_data.get(col)
    return str(v) if pd.notnull(v) and str(v).strip() != '' else default

  def get_dynamic_color(text):
    t = str(text).upper()
    if any(
        w in t
        for w in [
            'DOWN',
            'CRITICAL',
            'NO BACKUP',
            'NOT AVAILABLE',
            'NEED VALIDATION',
            'NOT_AVAILABLE',
            'POTENSIAL TRIP',
            'TRIP',
            'NEED',
        ]
    ):
      return COLOR_RED
    elif any(
        w in t
        for w in ['NORMAL', 'SECURED', 'VALID', 'OK', 'AVAILABLE', 'VIP']
    ):
      return COLOR_GREEN
    elif any(w in t for w in ['SILVER', 'GOLD', 'LITHIUM', 'HUAWEI', 'TELKOM']):
      return COLOR_BLUE
    return COLOR_TEXT

  excel_action = site_data.get('Action')
  if pd.notnull(excel_action) and str(excel_action).strip() not in ['', '-']:
    action_list = [str(excel_action)]
  else:
    action_list = []
    rect_cond = str(site_data.get('Rectifier Condition', '')).upper()
    pot_trip = str(site_data.get('Potensial Trip', '')).upper()
    eas_val = str(site_data.get('EAS Validation', '')).upper()
    neteco = str(site_data.get('NETECO Status', '')).upper()

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

  site_cfg = {
      'start_x': 45,
      'start_y': 225,
      'colon_x': 200,
      'val_x': 215,
      'spacing': 38,
  }
  rect_cfg = {
      'start_x': 540,
      'start_y': 225,
      'colon_x': 740,
      'val_x': 765,
      'spacing': 38,
  }
  batt_cfg = {
      'start_x': 540,
      'start_y': 625,
      'colon_x': 740,
      'val_x': 765,
      'spacing': 38,
  }
  health_cfg = {
      'start_x': 1020,
      'start_y': 225,
      'colon_x': 1220,
      'val_x': 1245,
      'spacing': 34,
  }

  def render_box(cfg, items):
    y = cfg['start_y']
    spacing = cfg['spacing']
    for emoji, label, value in items:
      if label == '':
        y += int(spacing * 0.4)
        continue
      draw.text((cfg['start_x'], y), emoji, fill=COLOR_TEXT, font=font)
      draw.text(
          (cfg['start_x'] + 38, y), label, fill=COLOR_LABEL, font=font
      )
      draw.text((cfg['colon_x'], y), ':', fill=COLOR_TEXT, font=font)
      val_color = get_dynamic_color(value)
      draw.text((cfg['val_x'], y), f' {value}', fill=val_color, font=font)
      y += spacing

  col_site = [
      ('🆔', 'Site ID', val('Site ID')),
      ('📍', 'Site Name', val('Site Name')),
      ('🌐', 'Regional', val('Regional')),
      ('🏢', 'NOP', val('NOP_1')),
      ('📡', 'TO', val('TO')),
      ('🏠', 'ROH', val('ROH')),
      ('👤', 'Site Owner', val('Site Owner')),
      ('🧭', 'Lat / Long', f"{val('Lat')} / {val('Long')}"),
  ]

  col_rect = [
      ('⚡', 'ID PLN', val('ID PLN')),
      ('💡', 'Daya PLN', f"{val('Daya PLN (KVA)')} kVA"),
      ('🔌', 'Brand', val('Rectifier Brand (1)')),
      ('⚙️', 'Model', val('Rectifier Model (1)')),
      ('🔋', 'Capacity', val('Module Capacity (1)')),
      ('📦', 'Module Qty', val('Inserted Module Qty (1)')),
      ('📈', 'Load System', val('Load System (1)')),
  ]

  bbt_val = site_data.get('BBT H (1)') if 'BBT H (1)' in site_data else None
  bbt_str = (
      f'{round(float(bbt_val), 2)} Hours' if pd.notnull(bbt_val) else '-'
  )
  col_batt = [
      ('🔋', 'Brand', val('Battery Brand (1)')),
      ('🧪', 'Type', val('Battery Type (1)')),
      ('⚡', 'Capacity', val('Battery Capacity (1)')),
      ('📦', 'Bank Qty', val('Battery Bank (1)')),
      ('⏱️', 'BBT Backup', bbt_str),
      ('📊', 'Category', val('BBT Category (1)')),
  ]

  util_val = (
      site_data.get('Rectifier Utility')
      if 'Rectifier Utility' in site_data
      else None
  )
  util_str = (
      f'{round(float(util_val) * 100, 1)} %' if pd.notnull(util_val) else '-'
  )

  col_health = [
      ('🔧', 'Rect. Cond', val('Rectifier Condition')),
      ('📊', 'Utility', util_str),
      ('🔒', 'Config', val('Rectifier Config (Category)')),
      ('✔️', 'Cap. Status', val('Capacity Status')),
      ('⚠️', 'Pot. Trip', val('Potensial Trip')),
      ('📋', 'SOW Act.', val('Activity (SOW) Actual')),
      ('✅', 'EAS Valid.', val('EAS Validation')),
      ('🌐', 'NETECO Stat', val('NETECO Status')),
  ]

  render_box(site_cfg, col_site)
  render_box(rect_cfg, col_rect)
  render_box(batt_cfg, col_batt)
  render_box(health_cfg, col_health)

  line_y = (
      health_cfg['start_y'] + len(col_health) * health_cfg['spacing'] - 6
  )
  draw.text(
      (health_cfg['colon_x'], line_y),
      '===============',
      fill=COLOR_LABEL,
      font=font,
  )

  act_start_y = line_y + 28
  draw.text(
      (health_cfg['start_x'], act_start_y), '🛠️', fill=COLOR_TEXT, font=font
  )
  draw.text(
      (health_cfg['start_x'] + 38, act_start_y),
      'Action',
      fill=COLOR_LABEL,
      font=font,
  )
  draw.text(
      (health_cfg['colon_x'], act_start_y), ':', fill=COLOR_TEXT, font=font
  )

  current_y = act_start_y
  for idx, action_item in enumerate(action_list):
    if idx > 0:
      current_y += 28
    val_color = get_dynamic_color(action_item)
    draw.text(
        (health_cfg['val_x'], current_y),
        f' {action_item}',
        fill=val_color,
        font=font,
    )

  output_path = f'output_{site_id}.png'
  img.save(output_path)
  return output_path


@bot.message_handler(commands=['start', 'help'])
def send_welcome(message):
  bot.reply_to(
      message,
      'Halo! Ketik /site [Site ID] untuk generate kartu status site Telkomsel.',
  )


@bot.message_handler(commands=['site'])
def handle_site(message):
  text_parts = message.text.split()
  if len(text_parts) < 2:
    bot.reply_to(message, 'Format salah! Contoh penggunaan: `/site CKR207`')
    return

  site_id = text_parts[1].upper()
  bot.reply_to(message, f'Sedang memproses Site ID: {site_id}...')

  img_path = generate_site_card(site_id)
  if img_path and os.path.exists(img_path):
    with open(img_path, 'rb') as photo:
      bot.send_photo(
          message.chat.id, photo, caption=f'✅ Status untuk Site ID: {site_id}'
      )
    os.remove(img_path)
  else:
    bot.reply_to(
        message, f'❌ Maaf, Site ID {site_id} tidak ditemukan di database Excel!'
    )


print('Bot Telegram siap berjalan...')
bot.infinity_polling()
