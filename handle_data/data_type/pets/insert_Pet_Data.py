
'''
    @ 輸入 _ 客戶資料

'''

from tool.database_connection import engine_2018  # 資料庫連結
import pandas as pd
import pymysql


pets_2018 = pd.read_sql( '''
                            SELECT a.pet_id , a.custom_id , a.master_id , a.kind_id , a.p_name , 
                                   a.gender , a.age , a.color , a.note_1 , a.note_2 , a.note_3 ,
                                   b.kind_name
                            FROM pet_data AS a LEFT JOIN pet_kind AS b
                            ON a.kind_id = b.kind_id
                         '''  , engine_2018 )


# 連接 _ "新" 資料庫
conn_Gogopark_2021 = pymysql.Connect(
                                      host     = '127.0.0.1' ,
                                      user     = 'root' ,
                                      password = 'root' ,
                                      port     = 8889 ,
                                      db       = 'gogopark_ts' ,
                                      charset  = 'utf8'
                                    )

# 輸入資料至寵物資料表
def insert_Pets_Table( data ) :

    # 取出欄位
    pet_Id , custom_Id , master_Id , species_id , species_name , pet_Name \
    , sex , age , color , note_1 , note_2 , note_3 = (

        data['pet_id'],     # 寵物資料表 id
        data['custom_id'],  # 寵物自訂 id
        data['master_id'],  # 主人 id
        data['kind_id'],    # 品種 id
        data['kind_name'],  # 品種名稱
        data['p_name'],     # 寵物名字
        data['gender'],     # 公/母
        data['age'],        # 年紀
        data['color'],      # 毛色
        data['note_1'],     # 備註 : 洗澡美容
        data['note_2'],     # 備註 : 住宿
        data['note_3'],     # 備註 : 客訴及其他

    )

    # 將雙引號 "，替換為單引號 '，避免輸入時造成 SQL 語法錯誤
    note_1 = note_1.replace( '\"' , '\'' ) if note_1 is not None else ''

    # 去除空格
    custom_Id = custom_Id.strip()


    # 輸入 2021 資料表
    sql = f'''
             INSERT INTO `pet`
             ( `customer_id` , `serial` ,  `species` , `name` , `sex` , `color` , `note` )
             VALUES(
                     "{ master_Id }" , "{ custom_Id }" , "{ species_name }" , "{ pet_Name }" ,
                     "{ sex }" , "{ color }"  , "{ note_1 }"
                   )
           '''

    cursor = conn_Gogopark_2021.cursor()
    cursor.execute( sql )
    conn_Gogopark_2021.commit()


# 執行輸入
for idx , data in pets_2018.iterrows() :

    insert_Pets_Table( data )



print( '寵物資料，輸入成功' )
