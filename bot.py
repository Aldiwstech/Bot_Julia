import os
import telebot
import pandas as pd
from PIL import Image, ImageDraw, ImageFont

TOKEN = os.getenv('TOKEN')
if not TOKEN:
  raise ValueError("Token belum diset di Environment Variables Railway!")
  
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

  # UKURAN FONT DINAIKKAN JADI 40 SUPAYA JELAS DAN BESAR
  font_size = 40
  font_path = 'DejaVuSans.ttf'
  
  try:
    font = ImageFont.truetype(font_path, font_size)
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
    if any(w in t for w in ['DOWN', 'CRITICAL', 'NO BACKUP', 'NOT AVAILABLE', 'NEED VALIDATION', 'NOT_AVAILABLE', 'POTENSIAL TRIP', 'TRIP', 'NEED', 'PERGANTIAN']):
      return COLOR_RED
    elif any(w in t for w in ['NORMAL', 'SECURED', 'VALID', 'OK', 'AVAILABLE', 'VIP', 'MONITOR']):
      return COLOR_GREEN
    elif any(w in t for w in ['SILVER', 'GOLD', 'LITHIUM', 'HUAWEI', 'TELKOM']):
      return COLOR_BLUE
    return COLOR_TEXT

  action_list = []
  for col_name in ['Action', 'Activity (SOW) Actual', 'SOW']:
    if col_name in site_data:
      excel_action = site_data.get(col_name)
      if pd.notnull(excel_action) and str(excel_action).strip() not in ['', '-', 'nan']:
        action_list.append(str(excel_action).strip())
        break

  if not action_list:
    rect_cond = str(val('Rectifier Condition')).upper()
    pot_trip = str(val('Potensial Trip')).upper()
    eas_val = str(val('EAS Validation')).upper()
    neteco = str(val('NETECO Status')).upper()

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

  # KOORDINAT DIKOREKSI SUPAYA PAS DI DALAM KOTAK MOCKUP & JARAK ANTAR BARIS AMAN
  site_cfg = {'start_x': 60, 'start_y': 240, 'label_x': 110, 'colon_x': 380, 'val_x': 410, 'spacing': 55}
  rect_cfg = {'start_x': 560, 'start_y': 240, 'label_x': 610, 'colon_x': 880, 'val_x': 910, 'spacing': 55}
  batt_cfg = {'start_x': 560, 'start_y': 670, 'label_x': 610, 'colon_x': 880, 'val_x': 910, 'spacing': 55}
  health_cfg = {'start_x': 1060, 'start_y': 240, 'label_x': 1110, 'colon_x': 1380, 'val_x': 1410, 'spacing': 48}

  def render_box(cfg, items):
    y = cfg['start_y']
    spacing = cfg['spacing']
    for emoji, label, value in items:
      if label == '':
        y += int(spacing * 0.4)
        continue
      # Render ikon / nomor list kecil di sebelah kiri
      draw.text((cfg['start_x'], y), emoji, fill=COLOR_TEXT, font=font)
      # Render Label teks
      draw.text((cfg['label_x'], y), label, fill=COLOR_LABEL, font=font)
      # Titik dua
      draw.text((cfg['colon_x'], y), ':', fill=COLOR_TEXT, font=font)
      # Nilai dari Excel
      val_color = get_dynamic_color(value)
      draw.text((cfg['val_x'], y), f' {value}', fill=val_color, font=font)
      y += spacing

  col_site = [
      ('-', 'Site ID', val('Site ID')),
      ('-', 'Site Name', val('Site Name')),
      ('-', 'Regional', val('Regional')),
      ('-', 'NOP', val('NOP_1')),
      ('-', 'TO', val('TO')),
      ('-', 'ROH', val('ROH')),
      ('-', 'Site Owner', val('Site Owner')),
      ('-', 'Lat / Long', f"{val('Lat')} / {val('Long')}"),
  ]

  col_rect = [
      ('-', 'ID PLN', val('ID PLN')),
      ('-', 'Daya PLN', f"{val('Daya PLN (KVA)')} kVA"),
      ('-', 'Brand', val('Rectifier Brand (1)')),
      ('-', 'Model', val('Rectifier Model (1)')),
      ('-', 'Capacity', val('Module Capacity (1)')),
      ('-', 'Module Qty', val('Inserted Module Qty (1)')),
      ('-', 'Load System', val('Load System (1)')),
  ]

  bbt_val = site_data.get('BBT H (1)') if 'BBT H (1)' in site_data else None
  bbt_str = f'{round(float(bbt_val), 2)} Hours' if pd.notnull(bbt_val) and str(bbt_val).replace('.','',1).isdigit() else '-'
  col_batt = [
      ('-', 'Brand', val('Battery Brand (1)')),
      ('-', 'Type', val('Battery Type (1)')),
      ('-', 'Capacity', val('Battery Capacity (1)')),
      ('-', 'Bank Qty', val('Battery Bank (1)')),
      ('-', 'BBT Backup', bbt_str),
      ('-', 'Category', val('BBT Category (1)')),
  ]

  util_val = site_data.get('Rectifier Utility') if 'Rectifier Utility' in site_data else None
  util_str = f'{round(float(util_val) * 100, 1)} %' if pd.notnull(util_val) and str(util_val).replace('.','',1).isdigit() else '-'

  col_health = [
      ('-', 'Rect. Cond', val('Rectifier Condition')),
      ('-', 'Utility', util_str),
      ('-', 'Config', val('Rectifier Config (Category)')),
      ('-', 'Cap. Status', val('Capacity Status')),
      ('-', 'Pot. Trip', val('Potensial Trip')),
      ('-', 'SOW Act.', val('Activity (SOW) Actual')),
      ('-', 'EAS Valid.', val('EAS Validation')),
      ('-', 'NETECO Stat', val('NETECO Status')),
  ]

  render_box(site_cfg, col_site)
  render_box(rect_cfg, col_rect)
  render_box(batt_cfg, col_batt)
  render_box(health_cfg, col_health)

  line_y = health_cfg['start_y'] + (len(col_health) * health_cfg['spacing']) - 10
  draw.text((health_cfg['colon_x'], line_y), '===============', fill=COLOR_LABEL, font=font)

  act_start_y = line_y + 40
  draw.text((health_cfg['start_x'], act_start_y), '-', fill=COLOR_TEXT, font=font)
  draw.text((health_cfg['label_x'], act_start_y), 'Action', fill=COLOR_LABEL, font=font)
  draw.text((health_cfg['colon_x'], act_start_y), ':', fill=COLOR_TEXT, font=font)

  current_y = act_start_y
  for idx, action_item in enumerate(action_list):
    if idx > 0:
      current_y += 45
    val_color = get_dynamic_color(action_item)
    draw.text((health_cfg['val_x'], current_y), f' {action_item}', fill=val_color, font=font)

  output_path = f'output_{site_id}.png'
  img.save(output_path, dpi=(300, 300), quality=95)
  return output_path


@bot.message_handler(commands=['site'])
def handle_site(message):
  args = message.text.split()
  if len(args) < 2:
    bot.reply_to(message, "⚠️ Format salah! Gunakan: `/site <Site_ID>`", parse_mode='Markdown')
    return

  site_id = args[1]
  bot.reply_to(message, f"⏳ Sedang memproses Site ID: *{site_id}*...", parse_mode='Markdown')

  img_path = generate_site_card(site_id)

  if img_path and os.path.exists(img_path):
    with open(img_path, 'rb') as doc_file:
      bot.send_document(
          message.chat.id,
          doc_file,
          caption=f"✅ Status Report Site ID: *{site_id}*",
          parse_mode='Markdown'
      )
    os.remove(img_path)
  else:
    bot.reply_to(message, f"❌ Maaf, Site ID *{site_id}* tidak ditemukan.", parse_mode='Markdown')


print("Bot Telegram siap dijalankan...")
bot.infinity_polling()
