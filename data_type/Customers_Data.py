

'''

  @ 客戶資料

'''


import pandas as pd


class Customers_Data():

    # 讀取 _ 客戶資料
    def read_Customer_Data( self , read_Sql , engine_Type ) :
        return pd.read_sql( read_Sql , engine_Type )



