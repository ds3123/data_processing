




'''

  @ 寵物資料

'''


import pandas as pd


class Pet_Data():

    # 讀取 _ 寵物資料
    def read_Pet_Data( self , read_Sql , engine_Type ) :
        return pd.read_sql( read_Sql , engine_Type )

