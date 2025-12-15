import re
from datetime import datetime

def sanitize_filename(name):
    """Membersihkan karakter ilegal untuk nama file"""
    if not name: return "Unknown"
    clean = re.sub(r'[\\/*?:"<>|]', "", str(name)).strip()
    return clean if clean else "Unknown"

def clean_currency(value):
    if value in [None, ""]: return 0
    if isinstance(value, (int, float)): return value
    str_val = str(value).replace("Rp", "").replace(".", "").replace(",", ".").strip()
    try: return float(str_val)
    except: return 0

def detect_currency_from_text(text):
    """Mendeteksi currency dari kalimat 'Sales price in XXX excl. VAT'"""
    if not text: return "IDR"
    match = re.search(r"Sales price in\s+([A-Za-z]{3})", str(text), re.IGNORECASE)
    if match:
        return match.group(1).upper()
    return "IDR"

def extract_year_from_date(date_str):
    try:
        if not date_str: return str(datetime.now().year)
        parts = date_str.split('-')
        if len(parts) == 3:
            yy = parts[2]
            return f"20{yy}" if len(yy) == 2 else yy
    except:
        pass
    return str(datetime.now().year)

def addr_to_index(addr):
    """
    Mengubah alamat Excel ('K4') menjadi index tuple 0-based (row, col).
    Contoh: 'A1' -> (0, 0), 'K4' -> (3, 10)
    """
    if not addr: return None, None
    match = re.match(r"([A-Z]+)([0-9]+)", str(addr).upper())
    if not match: return None, None
    
    col_str, row_str = match.groups()
    row = int(row_str) - 1 # Excel row 1 -> index 0
    
    col = 0
    for char in col_str:
        col = col * 26 + (ord(char) - ord('A') + 1)
    col -= 1 # Excel col A -> index 0
    
    return row, col