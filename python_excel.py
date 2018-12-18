import xlrd
import pymysql
# importlib.reload(sys) #出现呢reload错误使用
'''
把Excel文件和本python文件放在同一目录下

需要安装的第三方库[pymysql,xlrd]

'''
class Excel_data_import_database(object):
    def __init__(self):
        self.excel_name = "test.xlsx"  # excel文件名
        self.sheet_name = "test"  # excelsheet名
        self.database_ip = "localhost"  # 数据库ip
        self.user = "root"  # 数据库用户
        self.password = "642936557"  # 数据库密码
        self.database = "pythontest"  # 库名
    def open_excel(self):
        try:
            book = xlrd.open_workbook(self.excel_name)  # 打开Excel文件
        except:
            print("打开Excel文件失败!")
        try:
            sheet = book.sheet_by_name(self.sheet_name)  # 读取sheet
            return sheet
        except:
            print("读取Excelsheet失败!")
    def connect_mysql(self):
        # 连接数据库
        try:
            db = pymysql.connect(host=self.database_ip, user=self.user,passwd=self.password,db=self.database,charset='utf8')
            return db
        except:
            print("连接数据库失败!")
    def search_count(self,table_name):
        db = self.connect_mysql()
        cursor = db.cursor()
        select = "select count(name) from "+table_name # 获取table_name表的记录数
        cursor.execute(select)  # 执行sql语句
        line_count = cursor.fetchone()
        print("一共"+str(line_count[0])+"条数据!")
    def insert_deta(self):
        sheet = self.open_excel()
        db = self.connect_mysql()
        cursor = db.cursor()
        data_list = []
        for i in range(1,sheet.nrows):  # 第一行是标题,不读
            name = sheet.cell(i, 0).value  # 取第i行第0列
            age = sheet.cell(i, 1).value  # 取第i行第1列，下面依次类推
            phone = sheet.cell(i,2).value
            sex = sheet.cell(i,3).value
            hobby = sheet.cell(i,4).value
            value = [name,age,phone,sex,hobby]
            select_sql = "select name from exceltest;"
            cursor = db.cursor()
            cursor.execute(select_sql)
            data = cursor.fetchall()
            for database_data in data:
                for list_data in database_data:
                    data_list.append(list_data)
            try:
                if str(name) in data_list:
                    sql_update = "UPDATE exceltest SET age=%s,phone=%s,sex=%s,hobby=%s where name=%s;"
                    cursor.execute(sql_update,[age,phone,sex,hobby,name])
                    db.commit()
                    print("数据更新成功",value)
                else:
                    sql_insert = "INSERT INTO exceltest(name,age,phone,sex,hobby)VALUES(%s,%s,%s,%s,%s);"
                    cursor.execute(sql_insert,value)  # 执行sql语句
                    db.commit()
                    print("数据插入成功",value)
            except:
                print("数据插入错误，事物回滚")
                db.rollback()
        cursor.close()  # 关闭连接
        db.close()  # 关闭数据
if __name__ == "__main__":
    Excel = Excel_data_import_database()
    Excel.insert_deta()
    Excel.search_count("exceltest")
