
'''

   @ 共同操作

'''

from tool.Common import Common
import os
import tool


# 取得 _ 工作簿、工作表( 第一個 )
# exe = Excel()
# wb , ws_1  = exe.get_Workbook_Sheet1()
# ws_1.title = '客戶篩選'  # 修改資料表名稱


# 讀取 _ Excel 檔案 ：工作簿、工作表
wb   = load_workbook( '../../data_files/狗狗公園資料篩選清單_2022.03.03.xlsx' )
ws_1 = wb.worksheets[ 0 ]

