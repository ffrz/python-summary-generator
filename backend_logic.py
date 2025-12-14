import xlrd
import os
import re

def addr_to_idx(address):
    """
    Mengubah 'K4' menjadi (3, 10)
    Mengubah 'AB10' menjadi (9, 27)
    """
    # Pisahkan Huruf dan Angka (misal: "K" dan "4")
    match = re.match(r"([A-Z]+)([0-9]+)", address.upper())
    if not match:
        raise ValueError(f"Format alamat sel salah: {address}")
    
    col_str, row_str = match.groups()
    
    # 1. Konversi Row (Excel 1-based -> Python 0-based)
    row_idx = int(row_str) - 1
    
    # 2. Konversi Column (A->0, B->1, ... AA->26)
    col_idx = 0
    for char in col_str:
        col_idx = col_idx * 26 + (ord(char) - ord('A') + 1)
    col_idx -= 1 # Jadikan 0-based
    
    return row_idx, col_idx

def get_val(sheet, address):
    """
    Helper sakti: Ambil nilai pakai kode Excel 'K4'
    """
    try:
        r, c = addr_to_idx(address)
        return sheet.cell_value(r, c)
    except IndexError:
        return None # Return None kalau sel diluar range
    
def clean_currency(value):
    """Sama seperti sebelumnya, membersihkan format"""
    if value in [None, ""]:
        return 0
    if isinstance(value, (int, float)):
        return value
    str_val = str(value).replace("Rp", "").replace(".", "").replace(",", ".").strip()
    try:
        return float(str_val)
    except:
        return 0

def extract_one_file(filepath):
    try:
        wb = xlrd.open_workbook(filepath)
        sheet = wb.sheet_by_index(0) 

        # --- 1. AMBIL HEADER (Pakai "Bahasa Manusia") ---
        # Tinggal sesuaikan string ini kalau nanti klien minta geser
        
        exchange_rate = get_val(sheet, "B4") # Exchange Rate
        date_updated = get_val(sheet, "B3") # Date Updated
        proj_value = get_val(sheet, "B5")  # Value (Sales price)
        cust_name  = get_val(sheet, "K3")  # Customer Name
        project_no = get_val(sheet, "K4")  # Project No
        
        # --- 2. AMBIL FOOTER (Search Logic tetap diperlukan) ---
        # Karena barisnya dinamis, kita tetap perlu looping cari kata kunci
        
        sub_total = 0
        warranty = 0
        cm_booked = 0
        cr_booked = 0
        penalty = 0
        total_cost = 0
        
        nrows = sheet.nrows
        limit = min(nrows, 100) # Scan max 100 baris
        
        for r in range(9, limit): # Mulai scan dari baris 10 (index 9)
            # Ambil text di kolom A baris ke-r
            # Kita bisa pakai cell_value biasa disini karena lagi looping
            cell_text = str(sheet.cell_value(r, 0)).upper() 
            
            # Logic: Jika ketemu Keyword, ambil nilai di Kolom E (Col index 4)
            # Kolom E = Index 4. 
            if sub_total == 0 and "SUB TOTAL" in cell_text:
                sub_total = sheet.cell_value(r, 4) 

            elif penalty == 0 and "PENALTY" in cell_text:
                penalty = sheet.cell_value(r, 4)
            
            elif warranty == 0 and "WARRANTTY" in cell_text:
                warranty = sheet.cell_value(r, 4)
            
            elif total_cost == 0 and "TOTAL COST" in cell_text:
                total_cost = sheet.cell_value(r, 4)

            elif cm_booked == 0 and "CM BOOKED" in cell_text:
                cm_booked = sheet.cell_value(r, 4)

            elif cr_booked == 0 and "CR BOOKED" in cell_text:
                cr_booked = sheet.cell_value(r, 4)

        return {
            "status": "OK",
            "Project No": project_no,
            "Cust Name": cust_name,
            "Date Updated": date_updated,
            "Exchange Rate": exchange_rate,
            "Project Value": clean_currency(proj_value),
            "Sub Total": clean_currency(sub_total),
            "Penalty": clean_currency(penalty),
            "Warranty": clean_currency(warranty),
            "Total Cost": clean_currency(total_cost),
            "CM Booked": clean_currency(cm_booked),
            "CR Booked": clean_currency(cr_booked)
        }

    except Exception as e:
        return {"status": "ERROR", "msg": str(e)}

# --- TEST AREA ---
if __name__ == "__main__":
    test_file = "2025 P1051 CUST1.xls"   # Pastikan ekstensi .xls
    
    if os.path.exists(test_file):
        print(f"Membaca {test_file} dengan XLRD...")
        result = extract_one_file(test_file)
        print("\nHASIL:")
        for k, v in result.items():
            print(f"{k} : {v}")
    else:
        print("File tidak ditemukan")