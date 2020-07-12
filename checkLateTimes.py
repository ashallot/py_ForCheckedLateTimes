import xlrd
import xlwt
import os
import datetime

def print_name(excel_path):
    files = os.listdir(excel_path)
    name_dict = {}
    num = 1
    for x in files:
        # print (x)
        name_dict[str(num)] = x
        print(num,x)
        num += 1
    choice = input("请输入需要操作文件的选项：")
    return name_dict[choice]


def rw_excel(file_path):
    # 打开Excel文件
    file_path = "\\" + file_path
    open_path = root_path + file_path
    workbook = xlrd.open_workbook(open_path)
    # 月度汇总
    month_sheet = workbook.sheets()[0]
    rows = month_sheet.nrows
    cols = month_sheet.ncols
    book = xlwt.Workbook() #创建一个Excel
    sheet1 = book.add_sheet('迟到汇总') 
    # 表头
    sheet1.write(0,0,'姓名') 
    sheet1.write(0,1,'迟到（11-30m）') 
    sheet1.write(0,2,'迟到（31-60m）') 
    sheet1.write(0,3,'迟到（61-90m）') 
    new_table_row = 0
    for i in range(rows-1):
        name = month_sheet.cell(i+1,0).value
        col_index = 14
        late10Nums = 0
        late30Nums = 0
        late60Nums = 0
        for j in range(cols-14):
            used_col_index = col_index + j
            checkedStr = month_sheet.cell(i+1,used_col_index).value
            if checkedStr != '':
                checkedArr = checkedStr.split("分钟")
                if len(checkedArr) == 2:
                    lateStr = checkedArr[0]
                    lateStrRes = lateStr.split("上班迟到")[1]
                    if len(lateStrRes)>2:
                        late60Nums = late60Nums + 1
                    else :
                        lateTime = int(lateStrRes)
                        if (lateTime>10) & (lateTime<31):
                            late10Nums = late10Nums + 1
                        if (lateTime>31) & (lateTime<60):
                            late30Nums = late30Nums + 1
                        
        if (late10Nums != 0) | (late30Nums != 0) | (late60Nums != 0):
            # 内容
            new_table_row = new_table_row+1
            sheet1.write(new_table_row,0,name) #如果label是中文一定要解码
            sheet1.write(new_table_row,1,late10Nums) 
            sheet1.write(new_table_row,2,late30Nums) 
            sheet1.write(new_table_row,3,late30Nums) 
    # 保存新表格
    new_name = file_path.split("月")[0]
    book.save(str(root_path)+'\\'+str(new_name)+'月迟到人次表.xls')
    print("运行结束！")
    
if __name__ == "__main__":
    root_path = os.getcwd()
    file_path = print_name(root_path)
    rw_excel(file_path)