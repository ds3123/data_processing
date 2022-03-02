'''
   * 篩選客戶資料
   * 輸出成 Excel 檔案
'''

from tool.database_connection import engine_2018  # 資料庫連結
from data_type.Customers_Data import Customers_Data
from data_format.Excel import Excel
from data_format.Excel_Styles import Excel_Styles
from filter_condition.Filter_Customers import Filter_Customers


# 客戶
cus = Customers_Data()
customers_2018 = cus.read_Customer_Data('SELECT * FROM master', engine_2018)

# 取得 _ 工作簿、工作表( 第一個 )
exe = Excel()
wb , ws = exe.get_Workbook_Sheet1()

# 樣式
style = Excel_Styles()

# 篩選錯誤條件
filer = Filter_Customers()

# 寫入資料
column_List  = ['A', 'B', 'C']
column_Title = ['客戶姓名', 'ID', '手機號碼']  # 標題
ws.append(column_Title)


for idx, data in customers_2018.iterrows():

    customer_Name, customer_Id, mobilePhone = (
        data['name'],  # 客戶姓名
        data['master_id'],  # id
        data['phone']  # 手機號碼
    )

    # 新增資料
    ws.append([customer_Name, customer_Id, mobilePhone])

    #  設定 _ 預設樣式 ( 置中 )
    style.set_Default_Style( column_List , ws , idx )


    # 判斷 _ 是否為錯誤
    if ( filer.is_Error_Customer_Name( customer_Name ) ):
        style.set_Error_Style( column_List , ws , idx )


# 存檔
wb.save('../../data_files/88.xlsx')
