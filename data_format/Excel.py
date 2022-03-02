
import openpyxl

''' 
  Excel 格式
'''

class Excel():

    workbook = None # 工作簿
    sheet_1  = None # 工作表 ( 第一個 )

    def __init__( self ) :
        # 新增 _ 工作簿、工作表
        wb = openpyxl.Workbook()
        self.workbook = wb
        self.sheet_1  = wb.worksheets[0]


    # 輸出 _ 客戶資料
    def output_Customers_Excel( self , file_Name ):
        pass

    def get_Workbook_Sheet1( self ):
        return self.workbook , self.sheet_1