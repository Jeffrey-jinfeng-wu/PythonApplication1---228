import pandas as pd
import pyodbc
import sqlalchemy
import math
import datetime
import numpy as np
import os.path as path

conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=DESKTOP-E7J4R0R;'
                      'Database=example;'
                      'Trusted_Connection=yes;')
cursor = conn.cursor()

sql = "select * from \
(select *, ROW_NUMBER() OVER(PARTITION BY category,date order by Category, logdate desc, budgetcost) AS row from otb) t1 \
where row=1 order by category,date,row"

cursor.execute(sql)
row = cursor.fetchall()
record = pd.DataFrame(row)



#### 1. ItemList ####
#if path.exists("C:\Document\Solutions\Aging and Out of Stock in Stores\Analysis\\backup\\test\itemlist.xls"):
#    print(str(datetime.datetime.now())[11:19])
#    print("ItemList")

#    itemlist = pd.ExcelFile("C:\Document\Solutions\Aging and Out of Stock in Stores\Analysis\\backup\\test\itemlist.xls")
#    sheet_num = len(itemlist.sheet_names)
#    print('Read ItemList')
    
#    cursor.execute("delete item")
#    conn.commit()
#    print('Delete Item')

#    for numofsheet in range(0, sheet_num):
#        for index, row in itemlist.parse(numofsheet).iterrows():
#            brandname = row['Brand Name']
#            if pd.isnull(brandname):
#                brandname = ''
#            if pd.isnull(row['Item Type']) :
#                continue
#            cursor.execute(
#                "INSERT INTO Item (Category, CategoryID, ItemCode, BrandName, Phaseout, Level) VALUES (?,?,?,?,?,?)",
#                row['TopCategory'], row['TopCategoryId'], row['Item Code'], brandname, row['Phase Out'], row['Level'])
#            cursor.execute("INSERT INTO ItemLevel ([itemcode],[date],[level]) VALUES (?,?,?)", row['Item Code'], today, row['Level'])
#            conn.commit()
#    print('Write Item')
    
#    #### 3. ItemStock ####
#if path.exists("C:\Document\Solutions\Aging and Out of Stock in Stores\Analysis\\backup\\test\ItemStockList.xls"):
#    print(datetime.datetime.now())
#    print('ItemStock')

#    itemstock = pd.ExcelFile(
#        "C:\Document\Solutions\Aging and Out of Stock in Stores\Analysis\\backup\\test\ItemStockList.xls")
#    print('Read ItemStockList')

#    cursor.execute("delete stock")
#    conn.commit()
#    print('Delete Stock')

#    sheet_num = len(itemstock.sheet_names)
#    for numofsheet in range(0, sheet_num):
#        for index, row in itemstock.parse(numofsheet).iterrows():
#            if pd.isnull(row['Item']) or row['Primary']==0 or row['Type']=='Unknown':
#                continue
#            if math.isnan(row['AP Cost']):
#               if math.isnan(row['LP Cost']):
#                   if (math.isnan(row['Market Cost'])):
#                       continue
#                   else:
#                       cost = row['Market Cost']
#               else:
#                   cost = row['LP Cost']
#            else:
#                cost = row['AP Cost']

#            sql = "INSERT INTO stock (ItemCode, oh, AvgLDCost,Phaseout,Saleable,Type) VALUES (?,?,?,?,?,?)"
#            cursor.execute(sql,row['Item'],row['Primary'],cost,row['PhaseOut'],row['Saleable'],row['Type'])
#            conn.commit()
#    print('Write Stock')


##### 6. PO ####
#if path.exists("C:\Document\Solutions\Aging and Out of Stock in Stores\Analysis\\backup\\test\ExportPOItemList.xls"):
#    sql = "delete OnOrder "
#    cursor.execute(sql)
#    conn.commit()
#    print('Delete OnOrder')

#    po = pd.read_excel("C:\Document\Solutions\Aging and Out of Stock in Stores\Analysis\\backup\\test\ExportPOItemList.xls")
#    print('Read PO')
#    sql = 'INSERT INTO OnOrder ([ItemCode], [Category], [WareHouseName], [SupplierCode], [PONo], [POPNo], [Status], [PODate], [Type], \
#    [OrderQty], [VoidQty], [SpareQty], [Cost], [Cur], [Currency], [BO], [ETADate], [ReceiveNo], [RecTime]) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)'
#    for index, row in po.iterrows():
#        rt = row['RecTime']
#        if pd.isnull(rt):
#            rectime = ''
#        else:
#            rectime = str(rt)[:10]

#        eta = row['ETADate']
#        if pd.isnull(eta):
#            eta = ''
#        else:
#            eta = str(eta)[:10]

#        receiveno = row['ReceiveNo']
#        if pd.isnull(receiveno):
#            receiveno = ''

#        popno = row['POPNo']
#        if pd.isnull(popno):
#            popno = ''

#        topcategory = row['TopCategory']
#        if pd.isnull(topcategory):
#            topcategory = ''

#        spareqty = row['SpareQty']
#        if pd.isnull(spareqty):
#            spareqty = 0

#        cursor.execute(sql, row['ItemCode'], topcategory, row['WareHouseName'], row['SupplierCode'], row['PONo'],
#                       popno, row['Status'], row['PODate'], row['Type'], row['OrderQty'], row['VoidQty'], spareqty, row['Cost'], row['Cur'],
#                       row['Currency Rate'], row['BO'], eta, receiveno, rectime)
#        conn.commit()

#    sql = "delete po where status = 'Not Submitted' or status = 'not approved' or status = 'deleted'"
#    cursor.execute(sql)
#    conn.commit()

#    print('Write PO')


cursor.close()
conn.close()
