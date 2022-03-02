'''

  @ 客戶篩選條件

'''



class Filter_Customers():
    
    # 客戶姓名錯誤類型
    name_Error = [ ' ' , '先生', '小姐', '先生小姐' , '先生及小姐' , '盧素雪' , '張順發' ]


    # 判斷 _ 客戶姓名是否有誤
    def is_Error_Customer_Name(self, name):

        _name = name.strip()  # 去除空格

        # 比對錯誤情況
        error_List = []
        for err in self.name_Error:
            error_List.append( _name == err )

        if True in error_List:
            return True
        else:
            return False
