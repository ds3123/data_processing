
'''
 * 篩選 _ 寵物資料
 * 輸出成 Excel 檔案
'''


from tool.database_connection import engine_2018  # 資料庫連結
from data_type.Pets_Data import Pet_Data
from data_format.Excel import Excel
from data_format.Excel_Styles import Excel_Styles


# 寵物
pet = Pet_Data()
pets_2018 = pet.read_Pet_Data( 'SELECT * FROM pet_data' , engine_2018 )


# 樣式
style = Excel_Styles()

# 取得 _ 工作簿、工作表( 第一個 )
excel      = Excel()
wb , ws_1  = excel.get_Workbook_Sheet1()
ws_1.title = '寵物篩選'  # 修改資料表名稱





# 存檔
#wb.save('../../data_files/../../data_files/篩選清單_寵物.xlsx')

