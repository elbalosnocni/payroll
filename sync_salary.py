import datetime
import os
import requests
import json
from pyxlsb import open_workbook

# ---------------------------------------------------------
# CONFIGURATION & CONSTANTS
# ---------------------------------------------------------
GAS_WEB_APP_URL = "https://script.google.com/macros/s/AKfycbzCHDkrlhr4ZBzZUXGQe4P6RImV4YEe-IicO2W6PHWc0Fcmm9yblZ3GyCEa78KyCyf8/exec"
EXCEL_PASSWORD = "1234"

# Chuyển đổi tên cột Excel sang chỉ số Cột (0-indexed)
def col2num(col_str):
    num = 0
    for c in col_str.upper():
        num = num * 26 + (ord(c) - ord('A')) + 1
    return num - 1

# Mapping Cột theo yêu cầu
# Sheet Salary
COL_SAL = {
    'ho_ten': col2num('CF'),
    'ma_nv': col2num('D'),
    'luong_co_ban': col2num('F'),
    'ngay_lam_viec': col2num('H'),
    'ngay_le': col2num('I'),
    'ngay_nghi_huong_luong': col2num('J'),
    'ngay_nghi_khong_luong': col2num('K'),
    'gio_ot': col2num('L'),
    'gio_ngay_nghi': col2num('M'),
    'gio_ot_ngay_nghi': col2num('N'),
    'ngay_huong_luong_ttv': col2num('O'),
    'gio_tc_dem_ngay_nghi': col2num('P'),
    'gio_tc_dem': col2num('Q'),
    'gio_ngay_le': col2num('R'),
    'gio_ot_ngay_le': col2num('S'),
    'gio_tc_dem_ngay_le': col2num('T'),
    'ngay_ca_dem': col2num('U'),
    'tien_khac': col2num('X'),
    'tien_ky_luat': col2num('Y'),
    'gan_bo_2n': col2num('Z'),
    'gan_bo_5n': col2num('AA'),
    'gan_bo_10n': col2num('AB'),
    'nhan_o': col2num('AC'),
    'di_lai': col2num('AD'),
    'chuyen_can': col2num('AE'),
    'tro_cap_thoi_viec_pnam': col2num('AF'),
    'bhxh': col2num('AG'),
    'bhyt': col2num('AH'),
    'bhtn': col2num('AI'),
    'khau_tru_khac': col2num('AJ'),
    'tam_ung': col2num('AK'),
    'luong_thang': col2num('AN'),
    'luong_ngoai_gio': col2num('AO'),
    'hoa_hong_thuong_dm': col2num('AQ'),
    'tong_thu_nhap': col2num('AR'),
    'thue_tncn': col2num('AT'),
    'thuc_linh': col2num('AW')
}

# Sheet DSCNV
COL_DS = {
    'ho_ten': col2num('B'),
    'cccd': col2num('H'),
    'phong_ban': col2num('AF'),
    'bo_phan': col2num('AG'),
    'chuc_vu': col2num('AH')
}

def get_target_path():
    """Tự động tính ngày 10 hàng tháng lấy lùi 1 tháng"""
    now = datetime.datetime.now()
    # Lùi lại 1 tháng
    first_day_this_month = now.replace(day=1)
    last_month = first_day_this_month - datetime.timedelta(days=1)
    
    year_str = last_month.strftime('%Y')
    month_year_str = last_month.strftime('%m-%Y')
    
    path_snack = rf"\\192.168.0.253\vn hr\SALARY - 2014 - 2015\VNLWW\{year_str}\SALARY {month_year_str}.xlsb"
    path_flexible = rf"\\192.168.0.253\vn hr\SALARY - 2014 - 2015\Printing line\{year_str}\PRINTING LINE {month_year_str}.xlsb"
    
    return month_year_str, path_snack, path_flexible

def process_xlsb(file_path, xuong_name):
    records = []
    if not os.path.exists(file_path):
        print(f"[-] Không tìm thấy file: {file_path}")
        return records

    print(f"[+] Đang xử lý file: {file_path}")
    try:
        with open_workbook(file_path, password=EXCEL_PASSWORD) as wb:
            # 1. Đọc sheet DSCNV để lấy CCCD, Phòng ban, Bộ phận, Chức vụ
            dscnv_data = {}
            if 'DSCNV' in wb.sheets:
                with wb.get_sheet('DSCNV') as sheet:
                    for row in sheet.rows():
                        if len(row) <= COL_DS['chuc_vu']: continue
                        ho_ten_val = str(row[COL_DS['ho_ten']].v or '').strip()
                        cccd_val = str(row[COL_DS['cccd']].v or '').strip()
                        
                        # Giữ nguyên định dạng chuỗi, đảm bảo không mất số 0 ở đầu CCCD
                        if cccd_val.startswith("'"):
                            cccd_val = cccd_val[1:]
                        if cccd_val.endswith('.0'):
                            cccd_val = cccd_val[:-2]
                        cccd_val = cccd_val.zfill(12) if len(cccd_val) > 0 and len(cccd_val) < 12 else cccd_val

                        if ho_ten_val:
                            dscnv_data[ho_ten_val] = {
                                'cccd': cccd_val,
                                'phong_ban': str(row[COL_DS['phong_ban']].v or '').strip(),
                                'bo_phan': str(row[COL_DS['bo_phan']].v or '').strip(),
                                'chuc_vu': str(row[COL_DS['chuc_vu']].v or '').strip()
                            }

            # 2. Đọc sheet Salary
            if 'Salary' in wb.sheets:
                with wb.get_sheet('Salary') as sheet:
                    for idx, row in enumerate(sheet.rows()):
                        if idx < 3: continue # Bỏ qua dòng tiêu đề
                        if len(row) <= COL_SAL['thuc_linh']: continue
                        
                        ma_nv = str(row[COL_SAL['ma_nv']].v or '').strip()
                        ho_ten = str(row[COL_SAL['ho_ten']].v or '').strip()
                        
                        if not ma_nv or ma_nv.upper() == 'NONE': continue

                        # Lấy thông tin tương ứng từ DSCNV
                        info_ds = dscnv_data.get(ho_ten, {'cccd': '', 'phong_ban': '', 'bo_phan': '', 'chuc_vu': ''})

                        # Safe float extract helper
                        def get_val(col_idx, is_round=False):
                            try:
                                v = row[col_idx].v
                                if v is None: return 0
                                val = float(v)
                                return round(val, 2) if is_round else round(val)
                            except:
                                return 0

                        record = {
                            'xuong': xuong_name,
                            'ma_nv': ma_nv,
                            'ho_ten': ho_ten,
                            'cccd': str(info_ds['cccd']),
                            'phong_ban': info_ds['phong_ban'],
                            'bo_phan': info_ds['bo_phan'],
                            'chuc_vu': info_ds['chuc_vu'],
                            
                            # Thu nhập
                            'tong_thu_nhap': get_val(COL_SAL['tong_thu_nhap']),
                            'luong_thang': get_val(COL_SAL['luong_thang']),
                            'luong_co_ban': get_val(COL_SAL['luong_co_ban']),
                            'ngay_lam_viec': get_val(COL_SAL['ngay_lam_viec'], True),
                            'ngay_le': get_val(COL_SAL['ngay_le'], True),
                            'ngay_nghi_huong_luong': get_val(COL_SAL['ngay_nghi_huong_luong'], True),
                            'ngay_nghi_khong_luong': get_val(COL_SAL['ngay_nghi_khong_luong'], True),
                            'ngay_huong_luong_ttv': get_val(COL_SAL['ngay_huong_luong_ttv'], True),
                            
                            'luong_ngoai_gio': get_val(COL_SAL['luong_ngoai_gio']),
                            'gio_ot': get_val(COL_SAL['gio_ot'], True),
                            'gio_ngay_nghi': get_val(COL_SAL['gio_ngay_nghi'], True),
                            'gio_ot_ngay_nghi': get_val(COL_SAL['gio_ot_ngay_nghi'], True),
                            'gio_tc_dem_ngay_nghi': get_val(COL_SAL['gio_tc_dem_ngay_nghi'], True),
                            'gio_ngay_le': get_val(COL_SAL['gio_ngay_le'], True),
                            'gio_ot_ngay_le': get_val(COL_SAL['gio_ot_ngay_le'], True),
                            'gio_tc_dem_ngay_le': get_val(COL_SAL['gio_tc_dem_ngay_le'], True),
                            'ngay_ca_dem': get_val(COL_SAL['ngay_ca_dem'], True),
                            'gio_tc_dem': get_val(COL_SAL['gio_tc_dem'], True),
                            
                            'tien_khac': get_val(COL_SAL['tien_khac']),
                            'tien_ky_luat': get_val(COL_SAL['tien_ky_luat']),
                            'gan_bo_2n': get_val(COL_SAL['gan_bo_2n']),
                            'gan_bo_5n': get_val(COL_SAL['gan_bo_5n']),
                            'gan_bo_10n': get_val(COL_SAL['gan_bo_10n']),
                            'nhan_o': get_val(COL_SAL['nhan_o']),
                            'di_lai': get_val(COL_SAL['di_lai']),
                            'chuyen_can': get_val(COL_SAL['chuyen_can']),
                            'hoa_hong_thuong_dm': get_val(COL_SAL['hoa_hong_thuong_dm']),
                            'tro_cap_thoi_viec_pnam': get_val(COL_SAL['tro_cap_thoi_viec_pnam']),
                            
                            # Khấu trừ
                            'bhxh': get_val(COL_SAL['bhxh']),
                            'bhyt': get_val(COL_SAL['bhyt']),
                            'bhtn': get_val(COL_SAL['bhtn']),
                            'thue_tncn': get_val(COL_SAL['thue_tncn']),
                            'tam_ung': get_val(COL_SAL['tam_ung']),
                            'khau_tru_khac': get_val(COL_SAL['khau_tru_khac']),
                            
                            # Thực lĩnh
                            'thuc_linh': get_val(COL_SAL['thuc_linh'])
                        }
                        records.append(record)
    except Exception as e:
        print(f"[-] Lỗi khi xử lý {file_path}: {str(e)}")
    
    return records

def sync_to_google_sheet():
    period_str, path_snack, path_flexible = get_target_path()
    print(f"=== Đang thực hiện đồng bộ lương Kỳ: {period_str} ===")
    
    all_data = []
    all_data.extend(process_xlsb(path_snack, "Xưởng Snack"))
    all_data.extend(process_xlsb(path_flexible, "Xưởng Flexible"))
    
    if not all_data:
        print("[-] Không có dữ liệu để đồng bộ!")
        return

    payload = {
        'action': 'syncData',
        'period': period_str,
        'data': all_data
    }

    print(f"[+] Đang gửi {len(all_data)} dòng dữ liệu lên Google Apps Script...")
    res = requests.post(GAS_WEB_APP_URL, json=payload)
    print("[+] KẾT QUẢ API:", res.text)

if __name__ == "__main__":
    sync_to_google_sheet()
