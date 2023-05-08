
import os
import QueryDatabase
import GenerateFile
import PrintFile

# 程序参数设置
# 资源文件路径
# current_file='./assets/example.xlsx'
current_file = os.getcwd()+'\\assets\\example.xlsx'

# 该位置写入查询数据库返回数据,根据自身业务补全代码
database_Pack=QueryDatabase.getPack()
database_SN=QueryDatabase.getSN()
#生成需要打印的文件
GenerateFile.GenerateExcel(current_file,database_Pack,database_SN)
#用系统默认打印机打印
PrintFile.filePrint(current_file)
print('end')