import gspread
import os
from oauth2client.service_account import ServiceAccountCredentials

def clear_screen():
  
    if os.name == 'nt':
        _ = os.system('cls')
    
    else:
        _ = os.system('clear')

# เรียกใช้ฟังก์ชัน clear_screen() เพื่อเคลียร์หน้าจอ
clear_screen()

class colors:
    GREEN = '\033[92m'
    END = '\033[0m'

text = """
     ________________________
    |          V.0.1         |
    |      GOOGLE SHEETS     | 
    |    BY : Shadow522023   |
    |________________________|
""" 

text2 = "กรอกID ชีต : "
text3 ="รหัสแผ่นงาน : "
text4 = "พิมพ์ช่องแรกที่ต้องการลบ : "
text5 = "พิมช่องที่สองที่ต้องการลบถึง : "
colored_text = colors.GREEN + text + colors.END
print(colored_text)
colored_text2 = colors.GREEN + text2 + colors.END
colored_text3 = colors.GREEN + text3 + colors.END
colored_text4 = colors.GREEN + text4 + colors.END
colored_text5 = colors.GREEN + text5 + colors.END


id_sheets = input(colored_text2)
pwd_sheets = input(colored_text3)
delete_blox1 = input(colored_text4)
delete_blox2 = input(colored_text5)
# กำหนดข้อมูลใบรับรองความถูกต้อง (credentials)
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('google_sheets\gifted-bongo-387415-53a6fdc1be8d.json', scope)

# เชื่อมต่อกับ Google Sheets
client = gspread.authorize(credentials)

# เปิดชีท
sheet = client.open_by_key(id_sheets).worksheet(pwd_sheets)

# รับข้อมูลในช่วง C5:D57
data_range = sheet.range(delete_blox1 + ':' + delete_blox2)

# ลบข้อมูลในช่วง C5:D57
for cell in data_range:
    cell.value = ''
    end = ("ทำการลบเรียบร้อย")
    end_color = colors.GREEN + end + colors.END
    print(end_color)
    

# อัปเดตช่วงเซลล์ที่มีการลบข้อมูล
sheet.update_cells(data_range)
