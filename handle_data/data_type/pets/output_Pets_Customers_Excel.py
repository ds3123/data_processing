

'''
 * 篩選 _ 寵物資料 ( JOIN 客戶資料 )
 * 輸出成 Excel 檔案
'''


import pandas as pd
from tool.database_connection import engine_2018  # 資料庫連結

from data_format.Excel import Excel
from data_format.Excel_Styles import Excel_Styles


# 讀取 _ 寵物、客戶資料
pets_customers_2018 = pd.read_sql( '''
                                    SELECT a.p_name , a.custom_id , b.name , b.master_id , b.phone 
                                    FROM pet_data AS a 
                                    LEFT JOIN master AS b 
                                    ON a.master_id = b.master_id
                                   ''' , engine_2018 )


# 樣式
style = Excel_Styles()

# 取得 _ 工作簿、工作表( 第一個 )
excel      = Excel()
wb , ws_1  = excel.get_Workbook_Sheet1()
ws_1.title = '寵物客戶篩選'  # 修改資料表名稱


column_List  = [ 'A' , 'B' , 'C'  , 'D' , 'E' , 'F' ]
column_Title = [ '序號', '寵物名字' , '寵物 custom_id' , '客戶姓名' , '客戶 master_id' , '手機號碼' ]  # 標題
ws_1.append( column_Title )


pets_customers_2018 = pets_customers_2018.dropna() ; # 刪除 NaN ( float NaN )

for idx , data in pets_customers_2018.iterrows():

    pet_Name , pet_Id , customer_Name , customer_Id , mobilePhone = (
        data['p_name'] ,    # 寵物名字
        data['custom_id'] , # 寵物資料表 id
        data['name'] ,      # 客戶姓名
        data['master_id'] , # 客戶資料表 id
        data['phone']       # 客戶手機號碼
    )

    # print( f'{pet_Name} / {customer_Name} / {int(customer_Id)} / {mobilePhone} \n' )

    # 新增資料
    ws_1.append( [ ( idx+1 ) , pet_Name , pet_Id , customer_Name , customer_Id , mobilePhone ] )

    #  設定 _ 預設樣式 ( 置中 )
    style.set_Default_Style( column_List , ws_1 , idx )

# 調整欄位寬度
ws_1.column_dimensions['B'].width = 20
ws_1.column_dimensions['C'].width = 15
ws_1.column_dimensions['D'].width = 15
ws_1.column_dimensions['E'].width = 15
ws_1.column_dimensions['F'].width = 15

# 存檔
wb.save( '../../data_files/寵物與客戶.xlsx' )
print( '存檔成功' )
