

'''
 * 篩選 _ 客戶資料 ( JOIN 寵物資料 )
 * 輸出成 Excel 檔案
'''
from tool.database_connection import engine_2018  # 資料庫連結
from data_type.Customers_Data import Customers_Data
from data_format.Excel import Excel
from data_format.Excel_Styles import Excel_Styles
from filter_condition.Filter_Customers import Filter_Customers
from openpyxl.styles import Font , Color , PatternFill , NamedStyle , Alignment


# 讀取 _ 寵物、客戶資料
cus = Customers_Data()
customers_pets_2018 = cus.read_Customer_Data( '''
                                                SELECT a.name , a.master_id , a.phone , b.p_name , b.custom_id 
                                                FROM master AS a 
                                                RIGHT JOIN pet_data AS b 
                                                ON a.master_id = b.master_id
                                               ''' , engine_2018 )

# 樣式
style = Excel_Styles()

# 取得 _ 工作簿、工作表( 第一個 )
excel      = Excel()
wb , ws_1  = excel.get_Workbook_Sheet1()
ws_1.title = '客戶_寵物篩選'  # 修改資料表名稱

# 篩選錯誤條件
filer = Filter_Customers()

# 寫入資料
column_List  = [ 'A' , 'B' , 'C' , 'D' , 'E' ]
column_Title = [ '客戶姓名' , '客戶 master_id' , '手機號碼' , '寵物名字' , '寵物 custom_id' ]  # 標題
ws_1.append( column_Title )

# 不重複資料
cus_Multi_Pets = set()
cus_Has_Pets   = set() # 有寵物的客戶

for idx , data in customers_pets_2018.iterrows():

    customer_Name , customer_Id , mobilePhone , pet_Name , pet_Id = (
        data['name'] ,       # 客戶姓名
        data['master_id'] ,  # id
        data['phone'] ,      # 手機號碼
        data['p_name'] ,     # 寵物名字
        data['custom_id']    # 寵物客制 id
    )

    cus_Has_Pets.add( customer_Name )


    # 前一個 :
    pre_Index         = idx - 1 if idx > 0 else 0
    customer_Name_Pre = customers_pets_2018.loc[ pre_Index ]["name"]   # 客戶姓名
    pet_Name_Pre      = customers_pets_2018.loc[ pre_Index ]["p_name"] # 寵物名字

    # 去除左右空格
    customer_Name_Pre = customer_Name_Pre.strip() if customer_Name_Pre is not None else ''
    customer_Name     = customer_Name.strip() if customer_Name is not None else ''
    mobilePhone       = mobilePhone.strip() if mobilePhone is not None else ''

    # 新增資料
    ws_1.append( [ customer_Name , customer_Id , mobilePhone , pet_Name , pet_Id ] )

    # 1.姓名中有 : 測試 ( 灰色標示 )
    if ( '測試' in customer_Name or 'test' in customer_Name ):
        style.set_Error_Style( column_List , ws_1 , idx , '817d80')

    # 2.姓名中有 : 拒接 ( 黑色標示 )
    if ( '拒接' in customer_Name ):
        style.set_Error_Style( column_List , ws_1 , idx , '000000')

    # 3.姓名中有 : 停用、不用、改號碼 ( 桃色標示 )
    if( '停用' in customer_Name or '不用' in customer_Name ) :
        style.set_Error_Style( column_List , ws_1, idx, 'f815c5')

    # 4.列舉清單錯誤 ( 紫色標示 Ex. '先生' , '小姐' , '先生小姐' ....  )
    if( filer.is_Error_Customer_Name( customer_Name ) ) :
        style.set_Error_Style( column_List , ws_1 , idx , 'ca0bec' )

    # 5.姓名重複( 藍色標示 )
    if( idx > 2 and customer_Name == customer_Name_Pre ) :
        style.set_Error_Style( column_List , ws_1 , idx-1 , '2d28f4' )
        style.set_Error_Style( column_List , ws_1 , idx , '2d28f4' )

        cus_Multi_Pets.add( customer_Name )


    # 6.沒有手機號碼 ( 紅色標示 )
    if ( mobilePhone == '' ) :
        style.set_Error_Style( column_List , ws_1 , idx , 'FF0000' )

    # 7.客戶姓名為數字 ( 綠色標示 )
    if( customer_Name.isnumeric() ) :
        style.set_Error_Style( column_List , ws_1 , idx , '3dd731' )

    # 設定 _ 預設樣式 ( 置中 )
    # style.set_Default_Style( column_List , ws_1 , idx)



print( f'多隻寵物的客人數 ：{ len( cus_Multi_Pets ) } '  )
print( f'有寵物的客人數 ： { len( cus_Has_Pets ) } ' )





# 調整欄位寬度
ws_1.column_dimensions['A'].width = 25
ws_1.column_dimensions['B'].width = 17
ws_1.column_dimensions['C'].width = 15
ws_1.column_dimensions['D'].width = 25
ws_1.column_dimensions['E'].width = 15

# 存檔
#wb.save( '../../data_files/客戶_2.xlsx' )
print('存檔成功')
