from data_type.Customers_Data import Customers_Data
from tool.database_connection import engine_2018  # 資料庫連結
from filter_condition.Filter_Customers import Filter_Customers
from data_format.Excel import Excel
from data_format.Excel_Styles import Excel_Styles
from filter_condition.Filter_Customers import Filter_Customers
from openpyxl.styles import Font , Color , PatternFill , NamedStyle , Alignment
import math
import pymysql
import sys


'''

  @ 輸入 _ 客戶資料
    
'''

# 客戶
cus = Customers_Data()

# 客戶
customers_2018 = cus.read_Customer_Data( '''
                                           SELECT name , gender , master_id , phone , 
                                                  telephone , address , note ,
                                                  mergency_contact , mergency_contact_phone , mergency_contact_telephone
                                           FROM master
                                         ''' , engine_2018 )


# 客戶 JOIN 寵物
customers_pets_2018 = cus.read_Customer_Data( '''
                                                SELECT a.name , a.master_id , a.phone , b.p_name , b.custom_id 
                                                FROM master AS a 
                                                RIGHT JOIN pet_data AS b 
                                                ON a.master_id = b.master_id
                                               ''' , engine_2018 )


# 連接 _ "新" 資料庫
conn_Gogopark_2021 = pymysql.Connect(
                                      host     = '127.0.0.1' ,
                                      user     = 'root' ,
                                      password = 'root' ,
                                      port     = 8889 ,
                                      db       = 'gogopark_ts' ,
                                      charset  = 'utf8'
                                    )


# 篩選錯誤條件
filer = Filter_Customers()

# 從 DataFrame 分離出所需欄位
def get_Columns( data ) :

    customer_Name , customer_Sex ,  customer_Id , mobilePhone , \
    telePhone , address , note , \
    contact_Name , contact_MobilePhone , contact_TelePhone = (

        data['name'] ,            # 客戶姓名
        data['gender'] ,          # 客戶性別
        data['master_id'] ,       # id
        data['phone'] ,           # 手機號碼
        data['telephone'] ,       # 家用號碼
        data['address'] ,         # 通訊地址
        data['note'] ,            # 主人備註

        data['mergency_contact'] ,          # 緊急連絡人姓名
        data['mergency_contact_phone'] ,    # 緊急連絡人手機號碼
        data['mergency_contact_telephone']  # 緊急連絡人家用電話

    )

    # 將雙引號 "，替換為單引號 '，避免輸入時造成 SQL 語法錯誤
    address = address.replace( '\"' , '\'' ) if address is not None else ''
    note    = note.replace( '\"' , '\'' )    if note is not None else ''

    return  customer_Name , customer_Sex , customer_Id , \
            mobilePhone , telePhone , address , note , \
            contact_Name , contact_MobilePhone , contact_TelePhone



# 輸入資料至客戶資料表
def insert_Customers_Table( data ) :

    # 取得所需欄位
    customer_Name , customer_Sex , customer_Id , mobilePhone ,\
    telePhone , address , note , \
    contact_Name , contact_MobilePhone , contact_TelePhone \
    = get_Columns( data )

    # 輸入 2021 資料表 : 客戶
    sql_1 = f'''
                INSERT INTO `customer`
                ( `name` , `id` , `mobile_phone` , `tel_phone` , `address`, `sex` , `note` )
                VALUES( "{ customer_Name }" , "{ customer_Id }" , "{ mobilePhone }" , "{ telePhone }" , "{ address }" , "{ customer_Sex }" , "{ note }" )
             '''

    # 輸入 2021 資料表 : 客戶關係人
    sql_2 = f'''
                INSERT INTO `customer_relation`
                ( `customer_id` , `type` , `name` , `mobile_phone` , `tel_phone` )
                VALUES( "{ customer_Id }" , "緊急連絡人" , "{ contact_Name }" , "{ contact_MobilePhone  }" , "{ contact_TelePhone }"  )
             '''

    # 輸入 2021 資料表 : 客戶關係人
    cursor = conn_Gogopark_2021.cursor()
    cursor.execute( sql_1 )
    cursor.execute( sql_2 )
    conn_Gogopark_2021.commit()



# 判斷 _ 是否輸入至資料庫
def is_Insert_Customer( data ) :

    customer_Name , customer_Sex , customer_Id, mobilePhone, \
    telePhone, address, note, \
    contact_Name, contact_MobilePhone, contact_TelePhone \
    = get_Columns(data)

    # 2.姓名中有 : 測試 ( 灰色標示 )
    if '測試' in customer_Name or 'test' in customer_Name: return False

    # 3.姓名中有 : 停用、不用、改號、已換門號 ( 桃色標示 )
    if '停用' in customer_Name or '不用' in customer_Name or '號' in customer_Name : return False

    # 5.列舉清單錯誤 ( 紫色標示 Ex. '先生' , '小姐' , '先生小姐' ....  )
    if filer.is_Error_Customer_Name( customer_Name ) : return False

    # 7.沒有手機號碼 ( 紅色標示 )
    if mobilePhone == '' : return False

    # 8.客戶姓名為數字 ( 綠色標示 )
    if customer_Name.isnumeric() : return False

    return True


# 取得 _ "有"寵物客戶 ( 不重複 )
cus_Has_Pets_Id = set()

for idx , data in customers_pets_2018.iterrows() :
    customer_Id = data['master_id']  # 客人 master_id
    customer_Id = '' if math.isnan( customer_Id ) else int( customer_Id ) # 排除 NaN、轉為 INT
    cus_Has_Pets_Id.add( customer_Id )


# 不重複客戶
not_Duplicate_Customers = set()
for idx , data in customers_2018.iterrows() :
    not_Duplicate_Customers.add( data['name'] )


for idx , data in customers_2018.iterrows() :

    # 取得所需欄位
    customer_Name , customer_Sex , customer_Id, mobilePhone, \
    telePhone, address, note, \
    contact_Name, contact_MobilePhone, contact_TelePhone \
    = get_Columns(data)


    # 前一個客戶姓名
    pre_Index = idx - 1 if idx > 0 else 0
    customer_Name_Pre = customers_2018.loc[pre_Index]["name"]

    # 下一個客戶姓名
    next_Index         = idx + 1
    customer_Name_Next = customers_2018.loc[ next_Index ]["name"] if next_Index < len( customers_2018 ) else None

    # 去除左右空格
    customer_Name_Pre  = customer_Name_Pre.strip()
    customer_Name_Next = customer_Name_Next.strip() if customer_Name_Next is not None else ''
    customer_Name      = customer_Name.strip()
    mobilePhone        = mobilePhone.strip()

    if customer_Name in not_Duplicate_Customers and customer_Id not in cus_Has_Pets_Id :
        continue

    # 6.姓名重複、沒有寵物 --> 略過
    # is_Duplicate_NoPets = ( customer_Name == customer_Name_Pre or customer_Name == customer_Name_Next ) \
    #                        and \
    #                        customer_Id not in cus_Has_Pets_Id
    #
    # if is_Duplicate_NoPets : continue

    # 通過篩選條件 --> Insert 資料
    if is_Insert_Customer( data ):
        insert_Customers_Table( data )



print( '客戶資料，輸入成功' )




