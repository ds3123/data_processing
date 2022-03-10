

'''
 * 篩選 _ 洗澡資料 ( JOIN 寵物資料 )
 * 輸出成 Excel 檔案
'''

import pandas as pd
from tool.database_connection import engine_2018  # 資料庫連結

from data_format.Excel import Excel
from data_format.Excel_Styles import Excel_Styles


# 讀取 _ 寵物、客戶資料
bath_pets_2018 = pd.read_sql( 'SELECT a.bathe_id , a.pet_id , a.arrive_date , b.p_name , b.custom_id FROM bathe_history AS a LEFT JOIN pet_data AS b ON a.pet_id = b.pet_id' , engine_2018 )


# 樣式
style = Excel_Styles()

# 取得 _ 工作簿、工作表( 第一個 )
excel      = Excel()
wb , ws_1  = excel.get_Workbook_Sheet1()
ws_1.title = '洗澡寵物篩選'  # 修改資料表名稱


column_List  = [ 'A' , 'B' , 'C'  , 'D' , 'E' ]
column_Title = [ '序號', '洗澡單 ID' , '寵物名字' , '寵物 custom_id' , '到店日期'  ]  # 標題
ws_1.append( column_Title )


for idx , data in bath_pets_2018.iterrows():

    bath_Id , pet_Name , pet_Id , arrive_Date = (
        data['bathe_id'] ,     # 洗澡單 id
        data['p_name'] ,      # 寵物名字
        data['custom_id'] ,   # 寵物資料表 id
        data['arrive_date']   # 到店日期
    )

    # print( f'{ bath_Id } / { pet_Name } / { pet_Id } / { arrive_Date } \n' )

    # 新增資料
    ws_1.append( [ ( idx + 1 ) , bath_Id , pet_Name , pet_Id , arrive_Date  ] )

    #  設定 _ 預設樣式 ( 置中 )
    style.set_Default_Style( column_List , ws_1 , idx )


# 調整欄位寬度
ws_1.column_dimensions['B'].width = 15
ws_1.column_dimensions['C'].width = 30
ws_1.column_dimensions['D'].width = 20
ws_1.column_dimensions['E'].width = 15

# 存檔
wb.save( '../../data_files/洗澡單與寵物.xlsx' )
print( '存檔成功' )
