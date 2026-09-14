Attribute VB_Name = "Config"
Option Explicit

' ========================= KET NOI GAS =========================
Public Const GAS_URL As String = "https://script.google.com/macros/s/AKfycbzCHDkrlhr4ZBzZUXGQe4P6RImV4YEe-IicO2W6PHWc0Fcmm9yblZ3GyCEa78KyCyf8/exec"

' PHẢI TRÙNG với Script Property SYNC_API_KEY trong GAS.
' Đổi chuỗi này thành key bí mật của bạn.
Public Const SYNC_API_KEY As String = "CHANGE_ME_SYNC_KEY"

' ========================= FILE XLSB =========================
Public Const XLSB_PASSWORD As String = "1234"

Public Const SNACK_ROOT As String = "\\192.168.0.253\vn hr\SALARY - 2014 - 2015\VNLWW"
Public Const SNACK_FILE_PREFIX As String = "SALARY"
Public Const SNACK_LABEL As String = "Snack"

Public Const FLEXIBLE_ROOT As String = "\\192.168.0.253\vn hr\SALARY - 2014 - 2015\Printing line"
Public Const FLEXIBLE_FILE_PREFIX As String = "PRINTING LINE"
Public Const FLEXIBLE_LABEL As String = "Flexible"

Public Const SHEET_SALARY As String = "Salary"
Public Const SHEET_DSCNV As String = "DSCNV"
Public Const DATA_START_ROW As Long = 7

' ========================= CỘT SALARY =========================
Public Const COL_SALARY_HOTEN As String = "CF"
Public Const COL_SALARY_MANV As String = "D"
Public Const COL_SALARY_LUONGCOBAN As String = "F"
Public Const COL_SALARY_SONGAYLAMVIEC As String = "H"
Public Const COL_SALARY_SONGAYLE As String = "I"
Public Const COL_SALARY_SONGAYNGHIHUONGLUONG As String = "J"
Public Const COL_SALARY_SONGAYNGHIKHONGLUONG As String = "K"
Public Const COL_SALARY_SOGIONGOAIGIO As String = "L"
Public Const COL_SALARY_SOGIONGAYNGHI As String = "M"
Public Const COL_SALARY_SOGIONGOAIGIONGAYNGHI As String = "N"
Public Const COL_SALARY_SONGAYNGHIHUONGLUONGTOITHIEUVUNG As String = "O"
Public Const COL_SALARY_SOGIOTANGCADEMNGAYNGHI As String = "P"
Public Const COL_SALARY_SOGIOTANGCADEM As String = "Q"
Public Const COL_SALARY_SOGIOLAMNGAYLE As String = "R"
Public Const COL_SALARY_SOGIONGOAIGIONGAYLE As String = "S"
Public Const COL_SALARY_SOGIOTANGCADEMNGAYLE As String = "T"
Public Const COL_SALARY_SONGAYLAMCADEM As String = "U"
Public Const COL_SALARY_LUONGTHANG As String = "AN"
Public Const COL_SALARY_LUONGNGOAIGIO As String = "AO"
Public Const COL_SALARY_TIENKHAC As String = "X"
Public Const COL_SALARY_TIENKYLUAT As String = "Y"
Public Const COL_SALARY_TIENGANBO2NAM As String = "Z"
Public Const COL_SALARY_TIENGANBO5NAM As String = "AA"
Public Const COL_SALARY_TIENGANBO10NAM As String = "AB"
Public Const COL_SALARY_TIENNHAO As String = "AC"
Public Const COL_SALARY_TIENDILAI As String = "AD"
Public Const COL_SALARY_TIENTHUONGCHUYENCAN As String = "AE"
Public Const COL_SALARY_TROCAPTHOIVIECPHEPNAM As String = "AF"
Public Const COL_SALARY_BHXH As String = "AG"
Public Const COL_SALARY_BHYT As String = "AH"
Public Const COL_SALARY_BHTN As String = "AI"
Public Const COL_SALARY_KHAUTRUKHAC As String = "AJ"
Public Const COL_SALARY_TAMUNG As String = "AK"
Public Const COL_SALARY_HOAHONGTHUONGVUOTDINHMUC As String = "AQ"
Public Const COL_SALARY_TONGKHOANTHUNHAP As String = "AR"
Public Const COL_SALARY_THUETHUNHAP As String = "AT"
Public Const COL_SALARY_LUONGTHUCLINH As String = "AW"

' ========================= CỘT DSCNV =========================
Public Const COL_DSCNV_HOTEN As String = "B"
Public Const COL_DSCNV_CCCD As String = "H"
Public Const COL_DSCNV_PHONGBAN As String = "AF"
Public Const COL_DSCNV_BOPHAN As String = "AG"
Public Const COL_DSCNV_CHUCVU As String = "AH"
