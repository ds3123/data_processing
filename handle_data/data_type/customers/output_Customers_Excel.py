'''
   * 篩選 _ 客戶資料
   * 輸出成 Excel 檔案
'''
from tool.database_connection import engine_2018  # 資料庫連結
from data_type.Customers_Data import Customers_Data
from data_format.Excel import Excel
from data_format.Excel_Styles import Excel_Styles
from filter_condition.Filter_Customers import Filter_Customers
from openpyxl.styles import Font , Color , PatternFill , NamedStyle , Alignment
import math
import pymysql



# 客戶
cus = Customers_Data()

# 客戶
customers_2018 = cus.read_Customer_Data( '''
                                           SELECT name , master_id , phone 
                                           FROM master
                                         ''' , engine_2018 )


# 客戶 JOIN 寵物
customers_pets_2018 = cus.read_Customer_Data( '''
                                                SELECT a.name , a.master_id , a.phone , b.p_name , b.custom_id 
                                                FROM master AS a 
                                                RIGHT JOIN pet_data AS b 
                                                ON a.master_id = b.master_id
                                               ''' , engine_2018 )



# 取得 _ 有寵物客戶 ( 不重複 )
cus_Has_Pets_Id = set()

for idx , data in customers_pets_2018.iterrows():

    customer_Id = data['master_id']  # 客人 master_id
    customer_Id = '' if math.isnan( customer_Id ) else int( customer_Id ) # 排除 NaN、轉為 INT

    cus_Has_Pets_Id.add( customer_Id )



# 樣式
style = Excel_Styles()

# 取得 _ 工作簿、工作表( 第一個 )
excel      = Excel()
wb , ws_1  = excel.get_Workbook_Sheet1()
ws_1.title = '客戶篩選'  # 修改資料表名稱


# 篩選錯誤條件
filer = Filter_Customers()

# 寫入資料
column_List  = [ 'A' , 'B' , 'C' , 'D' ]
column_Title = [ '客戶姓名' , '客戶 master_id' , '手機號碼' , '是否有寵物' ]  # 標題
ws_1.append( column_Title )



# 從 DataFrame 分離出所需欄位
def get_Columns( data ) :

    customer_Name , customer_Id , mobilePhone = (
        data['name'] ,      # 客戶姓名
        data['master_id'] , # id
        data['phone']       # 手機號碼
    )

    return customer_Name , customer_Id , mobilePhone


for idx , data in customers_2018.iterrows() :

    # 取得所需欄位
    customer_Name , customer_Id , mobilePhone = get_Columns( data )


    # 前一個客戶姓名
    pre_Index         = idx - 1 if idx > 0 else 0
    customer_Name_Pre = customers_2018.loc[ pre_Index ]["name"]

    # 去除左右空格
    customer_Name_Pre = customer_Name_Pre.strip()
    customer_Name     = customer_Name.strip()
    mobilePhone       = mobilePhone.strip()

    # 是否有寵物
    has_Pets = '有' if customer_Id in cus_Has_Pets_Id else '無'

    # 新增資料
    ws_1.append( [ customer_Name , customer_Id , mobilePhone , has_Pets ] )


    # 1.沒有寵物
    if not customer_Id in cus_Has_Pets_Id :
        style.set_Error_Style( column_List , ws_1 , idx , 'f1b103' )

    # 2.姓名中有 : 測試 ( 灰色標示 )
    if '測試' in customer_Name or 'test' in customer_Name :
        style.set_Error_Style( column_List , ws_1 , idx , '817d80' )

    # 3.姓名中有 : 拒接 ( 黑色標示 )
    if '拒接' in customer_Name :
        style.set_Error_Style( column_List , ws_1 , idx , '000000' )

    # 4.姓名中有 : 停用、不用、改號、已換門號 ( 桃色標示 )
    if '停用' in customer_Name or '不用' in customer_Name or '號' in customer_Name :
        style.set_Error_Style( column_List , ws_1 , idx , 'f815c5' )

    # 5.列舉清單錯誤 ( 紫色標示 Ex. '先生' , '小姐' , '先生小姐' ....  )
    if filer.is_Error_Customer_Name( customer_Name )  :
        style.set_Error_Style( column_List , ws_1 , idx , 'ca0bec' )

    # 6.姓名重複( 藍色標示 )
    if idx > 2 and customer_Name == customer_Name_Pre  :
        style.set_Error_Style( column_List , ws_1 , idx-1 , '2d28f4' )
        style.set_Error_Style( column_List , ws_1 , idx , '2d28f4' )

    # 7.沒有手機號碼 ( 紅色標示 )
    if mobilePhone == ''  :
        style.set_Error_Style( column_List , ws_1 , idx , 'FF0000' )

    # 8.客戶姓名為數字 ( 綠色標示 )
    if customer_Name.isnumeric() :
        style.set_Error_Style( column_List , ws_1 , idx , '3dd731' )

    # 設定 _ 預設樣式 ( 置中 )
    # style.set_Default_Style( column_List , ws_1 , idx)


# 調整欄位寬度
ws_1.column_dimensions['A'].width = 25
ws_1.column_dimensions['B'].width = 17
ws_1.column_dimensions['C'].width = 15
ws_1.column_dimensions['D'].width = 15


# 存檔
wb.save( '../../data_files/客戶.xlsx' )
print('存檔成功')



