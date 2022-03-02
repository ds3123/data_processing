

'''

  @ 客戶資料

'''


import pandas as pd


class Customers_Data():

    def __init__( self ):
        pass

    def read_Customer_Data( self , read_Sql , engine_Type ) :
        return pd.read_sql( read_Sql , engine_Type )



