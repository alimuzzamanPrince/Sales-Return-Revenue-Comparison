import pyodbc as db
import numpy as np
import pandas as pd
import locale
locale.setlocale(locale.LC_ALL, 'en_US')


connection = db.connect('DRIVER={SQL Server};'
                        'SERVER=10.168.2.241;'
                        'DATABASE=ARCHIVESKF;'
                        'UID=sa;PWD=erp')#####Add trusted_connection=yes and remove uid and pwd while connecting to local

cursor = connection.cursor()

query = """
DECLARE @TotalItem VARCHAR(20)
SET @TotalItem = (SELECT count( distinct  ITEMNO) as TotalMovingItem FROM ICStockStatusCurrentLOT
              where  left(ITEMNO,1)<>'9' and AUDTORG<>'SKFDAT')


SELECT NDMNAME, LEFT(AUDTORG,3) AS AUDTORG,
Sum( CASE WHEN [Days]<=15 THEN 1 END) AS SUS
,Sum( CASE WHEN [Days]>15 AND [Days]<=35 THEN 1 END) AS US
,Sum( CASE WHEN [Days]>35 AND [Days]<=45 THEN 1 END) AS NS
,Sum( CASE WHEN [Days]>45 AND [Days]<=60 THEN 1 END) AS OS
,Sum( CASE WHEN [Days]>60 THEN 1 END) AS SOS
,@TotalItem-(Sum( CASE WHEN [Days]<=15 THEN 1 END)+Sum( CASE WHEN [Days]>15 AND [Days]<=35 THEN 1 END)+Sum( CASE WHEN [Days]>35 AND [Days]<=45 THEN 1 END)
+Sum( CASE WHEN [Days]>45 AND [Days]<=60 THEN 1 END)+Sum( CASE WHEN [Days]>60 THEN 1 END)) AS NIL
,Sum( CASE WHEN [Days]>=0 THEN 1 END) AS BranchItem
,@TotalItem AS TotalItem

FROM
--------------------------------------
(SELECT NDMNAME, Stock.ITEMNO, Stock.AUDTORG, Sum(QTYONHAND) as QTYONHAND, Sum(QTYSHIPPED)as QTYSHIPPED
,CAST(ISNULL(CASE WHEN SUM(QTYSHIPPED)>0 THEN (SUM(QTYONHAND)+ISNULL(SUM(SIT),0))/(SUM(QTYSHIPPED)/90) END,0) AS INT) AS [Days]FROM
--------------------------------------
(SELECT ITEMNO, AUDTORG, SUM(QTYONHAND) AS QTYONHAND FROM ICStockStatusCurrentLOT
       WHERE AUDTORG = 'KHLSKF' AND LEN(LOCATION)>'3' AND left(ITEMNO,1)<>'9'
       GROUP BY AUDTORG, ITEMNO) AS Stock
LEFT JOIN
(SELECT ITEM, AUDTORG, SUM(QTYSHIPPED) AS QTYSHIPPED FROM OESalesDetails
       WHERE TRANSDATE BETWEEN CONVERT(varchar(8), GETDATE()-90,112) AND CONVERT(varchar(8), GETDATE(),112)
       AND AUDTORG = 'KHLSKF'
	   GROUP BY ITEM, AUDTORG) AS Sales
       ON RTRIM(Stock.ITEMNO) = RTRIM(Sales.ITEM) AND RTRIM(Stock.AUDTORG) = RTRIM(Sales.AUDTORG)
LEFT JOIN
(SELECT ITEMNO, AUDTORG,SUM(QTY) AS SIT FROM GIT WHERE 
		AUDTORG = 'KHLSKF' AND OPENINGDATE = convert(varchar, getdate(), 23) GROUP BY ITEMNO, AUDTORG) as GIT
       ON Stock.ITEMNO = GIT.ITEMNO AND Stock.AUDTORG=GIT.AUDTORG
LEFT JOIN
(select distinct  BRANCHNAME,BRANCH,NDMNAME from NDM ) as NDM
       ON (Stock.AUDTORG=NDM.BRANCH)
---------------------------------
Group BY NDMNAME,Stock.ITEMNO, Stock.AUDTORG) AS TX
---------------------------------
GROUP BY NDMNAME, AUDTORG
"""

cursor.execute(query)
data = list(cursor.fetchall())

ndm = str(data[0][0])
branch = str(data[0][1])
sus = int(data[0][2])
us = int(data[0][3])
ns = int(data[0][4])
os = int(data[0][5])
sos = int(data[0][6])
nil = int(data[0][7])
branchitem = int(data[0][8])
totalitem = int(data[0][9])

statusdata = [sus, us, ns, os, sos, nil]
statusdata1 = [nil, sus, us, ns, os, sos]
missingitem = totalitem - branchitem


linequery = """ SELECT [NDMNAME],[AUDTORG],[SUS],[US],[NS],[OS],[SOS],[NIL],[TotalItem], [AllItem]
      ,RIGHT([AUDTDATE],2) AS [AUDTDATE]FROM [ARCHIVESKF].[dbo].[StockChart]
  WHERE AUDTORG = 'KHL' """

cursor.execute(linequery)

linedata = list(cursor.fetchall())

branchlist = []
SUS = []
US = []
NS = []
OS = []
SOS = []
NIL = []
BranchItem = []
TotalItem = []
AUDTDATE = []
rownumber = 0
for row in linedata:
        branchname=linedata[rownumber][1]
        branchlist.append(branchname)
        sus = linedata[rownumber][2]
        SUS.append(sus)
        us = linedata[rownumber][3]
        US.append(us)
        ns = linedata[rownumber][4]
        NS.append(ns)
        os = linedata[rownumber][5]
        OS.append(os)
        sos = linedata[rownumber][6]
        SOS.append(sos)
        nil = linedata[rownumber][7]
        NIL.append(nil)
        bi = linedata[rownumber][8]
        BranchItem.append(bi)
        ti = linedata[rownumber][9]
        TotalItem.append(ti)
        audtdate = linedata[rownumber][10]
        AUDTDATE.append(audtdate)
        rownumber = rownumber+1

SUS.reverse()
US.reverse()
NS.reverse()
OS.reverse()
SOS.reverse()
NIL.reverse()
BranchItem.reverse()
TotalItem.reverse()
AUDTDATE.reverse()

lostquery = """
SELECT T1.ITEMNO, AUDTDATE, ISNULL(ORDERDATE,0) AS ORDERDATE
, CAST(ISNULL(BACKORDERQUNATITY,0) AS INT) AS BACKORDERQUNATITY, CAST(TRADEPRICE AS float) AS TRADEPRICE  FROM

(select ITEMNO, MAX(AUDTDATE)AS AUDTDATE
from ICHistoricalStock 
where  ITEMNO IN (SELECT DISTINCT ITEMNO FROM 
					ICStockStatusCurrentLOT where left(itemno,1)<>9 AND AUDTORG <> 'SKFDAT'
					and ITEMNO not in 
					(SELECT distinct ITEMNO FROM ICStockStatusCurrentLOT 
					where  left(ITEMNO,1)<>'9' and AUDTORG ='KHLSKF'))
AND AUDTORG ='KHLSKF'
GROUP BY ITEMNO) AS T1

LEFT JOIN
(SELECT ITEM, MAX(ORDERDATE) AS ORDERDATE, SUM(QTYORDERED) AS BACKORDERQUNATITY FROM OEOrderDetails
WHERE AUDTORG = 'KHLSKF' AND ORDERDATE between FORMAT(GETDATE()-31,'yyyyMMdd') and FORMAT(GETDATE()-1,'yyyyMMdd')
GROUP BY ITEM)AS T2
ON T1.ITEMNO=T2.ITEM AND T1.AUDTDATE < T2.ORDERDATE

LEFT JOIN
(SELECT ITEMNO, CAST(TRADEPRICE AS float) AS TRADEPRICE FROM PRINFOSKF) AS T3
ON T1.ITEMNO = T3.ITEMNO
"""

###df = pd.read_excel("Book1.xlsx")
df = pd.read_sql(lostquery, connection)

backorder = df['BACKORDERQUNATITY']
tp = df['TRADEPRICE']
df['BACKORDERVALUE'] = backorder*tp


newquery = """SELECT FORMAT(GETDATE()-1,'yyyyMMdd') AS LASTDAY"""
datedf = pd.read_sql(newquery, connection)

yesterday = int(datedf['LASTDAY'])
yesterday = str(yesterday)
print(yesterday)

###Last day count
lastdayitem = df.query("ORDERDATE == "+yesterday)['BACKORDERVALUE'].count()
###Last day sum
lastdayvalue = df.query("ORDERDATE == "+yesterday)['BACKORDERVALUE'].sum()

###Last month count
lastmonthitem = df.shape[0]
###Last month sum
lastmonthvalue = df['BACKORDERVALUE'].sum()

lastdayvalue = locale.format("%d", lastdayvalue, grouping=True)
lastmonthvalue = locale.format("%d", lastmonthvalue, grouping=True)

#lastdayvalue = round(lastdayvalue, 2)
#lastmonthvalue = round(lastmonthvalue, 2)

print(lastdayitem)
print(lastdayvalue)
print(lastmonthitem)
print(lastmonthvalue)

print("I'm done with data")