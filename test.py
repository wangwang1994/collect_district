from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string

list_shizhongqu=[]
wb_shizhongqu=load_workbook('市中区.xlsx')
shizhongqu_xiangmushu=int(input('请输入市中区有多少个项目：'))
shizhongqu_lieshu=input('请输入市中区表格中最后一列有效数据列数。如AJ：')
ws_shizhongqu=wb_shizhongqu.active
for row in ws_shizhongqu['B5':str(shizhongqu_lieshu)+'{0}'.format(4+shizhongqu_xiangmushu)]:
    list_shizhongqu.append(row)
shizhongqu_lieshu=column_index_from_string(shizhongqu_lieshu)
wb_zongbiao=load_workbook('总表.xlsx')
ws_zongbiao=wb_zongbiao.active
for row in range(6,shizhongqu_xiangmushu+6):
    for col in range(2,shizhongqu_lieshu+1):
        if row==6:
            ws_zongbiao.cell(row=6,column=col).value=list_shizhongqu[0][col-2].value
        if row==7:
            ws_zongbiao.cell(row=7, column=col).value = list_shizhongqu[1][col-2].value
        if row == 8:
            ws_zongbiao.cell(row=8, column=col).value = list_shizhongqu[2][col-2].value
# wb_zongbiao.save('总表_市中区添加后.xlsx')


print(len(list_shizhongqu))
print(list_shizhongqu)
print(list_shizhongqu[0])
print(list_shizhongqu[0][0])
print(list_shizhongqu[0][0].value)