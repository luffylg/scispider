import xlrd
# xlrd模块读取xls文件
# 调用read方法返回以name和address为key的字典组成的列表
def read():
    data = xlrd.open_workbook('cncp-speaker list_20150427.xlsx')
    table = data.sheets()[0]
    nrows = table.nrows
    ncols = table.ncols
    ctype = 1
    xf = 0
    datalist=[]
    # for i in range(1:nrows):
    for i in range(1,nrows):
        datadict={}
        name=table.cell(i,0).value
        address=table.cell(i,5).value
        email=table.cell(i,4).value
        datadict["name"]=name
        datadict["address"]=address
        datadict["email"]=email
        #print("FirstName:"+FirstName+"  LastName:"+LastName)
        # table.put_cell(i, 9, ctype, i, xf)
        # print(table.cell(i,9))
        # print(table.cell(i,9).value)
        datalist.append(datadict)
    return datalist
if __name__ == '__main__':
    datalist=read()
    print(datalist)
    print(len(datalist))
