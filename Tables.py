import pyodbc as db
import numpy as np
import pandas as pd
import os

dirpath = os.path.dirname(os.path.realpath(__file__))

font = {'family': 'serif',
        'color':  '#bb20dd',
        'weight': 400,
        'size': 20,
        }
font1 = {'family': 'serif',
        'color':  '#3b5998',
        'weight': 700,
        'size': 17,
        }

connection = db.connect('DRIVER={SQL Server};'
                        'SERVER=10.168.2.241;'
                        'DATABASE=ARCHIVESKF;'
                        'UID=sa;PWD=erp')#####Add trusted_connection=yes and remove uid and pwd while connecting to local

cursor = connection.cursor()

revenuequery1 = """
SELECT COUNT(DISTINCT INVNUMBER) AS LASTLASTINVOICE, SUM(EXTINVMISC) as LASTLASTSALES, FORMAT(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE())-2, 0),'MMMM') AS LASTLASTMONTH
                                   FROM OESalesDetails
                                   WHERE transdate between FORMAT(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE())-2, 0),'yyyyMMdd')
								   AND FORMAT(DATEADD(MONTH, DATEDIFF(MONTH, -1, GETDATE())-2, -1),'yyyyMMdd') AND AUDTORG = 'KHLSKF' AND TRANSTYPE=1
"""

revenuequery2 = """
SELECT COUNT(DISTINCT INVNUMBER) AS LASTINVOICE, SUM(EXTINVMISC) as LASTSALES, FORMAT(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE())-1, 0),'MMMM') AS LASTMONTH
                                    FROM OESalesDetails
                                    WHERE transdate between FORMAT(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE())-1, 0),'yyyyMMdd')
									AND FORMAT(DATEADD(MONTH, DATEDIFF(MONTH, -1, GETDATE())-1, -1),'yyyyMMdd') AND AUDTORG = 'KHLSKF' AND TRANSTYPE=1
"""
rev1 = pd.read_sql_query(revenuequery1, connection)
rev2 = pd.read_sql_query(revenuequery2, connection)

lastlastinv = int(rev1['LASTLASTINVOICE'])
lastlastsales = float(rev1['LASTLASTSALES'])
lastlastmonth = str(rev1['LASTLASTMONTH'].item())

lastinv = int(rev2['LASTINVOICE'])
lastsales = float(rev2['LASTSALES'])
lastmonth = str(rev2['LASTMONTH'].item())

invgrowth = format(((lastinv-lastlastinv)/lastlastinv)*100,',.2f')+"%"
salesgrowth = format(((lastsales-lastlastsales)/lastlastsales)*100,',.2f')+"%"
salesgrowthnum = ((lastsales-lastlastsales)/lastlastsales)*100
invgrowthnum = ((lastsales-lastlastsales)/lastlastsales)*100
#invgrowthnum = -0.9
#salesgrowthnum = -0.35
#salesgrowth = int(salesgrowth)
#print(salesgrowth)
#print(salesgrowthnum)
#lastlastinv = format(lastlastinv, ',d')
#lastlastsales = format(lastlastsales, ',.2f')
#lastinv = format(lastinv, ',d')
#lastsales = format(lastsales, ',.2f')

mtdquery1 = """
SELECT COUNT(DISTINCT INVNUMBER) AS LASTINVOICE, SUM(EXTINVMISC) as LASTSALES, FORMAT(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE())-1, 0),'MMMM') AS LASTMONTH
                                   FROM OESalesDetails
                                   WHERE transdate between FORMAT(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE())-1, 0),'yyyyMMdd')
								   AND FORMAT(DATEADD(month, -1, GETDATE()-1),'yyyyMMdd') AND AUDTORG = 'KHLSKF' AND TRANSTYPE=1
"""

mtdquery2 = """
SELECT COUNT(DISTINCT INVNUMBER) AS THISINVOICE, SUM(EXTINVMISC) as THISSALES, FORMAT(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()), 0),'MMMM') AS THISMONTH
                                    FROM OESalesDetails
                                    WHERE transdate between FORMAT(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()), 0),'yyyyMMdd')
									AND FORMAT(GETDATE()-1,'yyyyMMdd') AND AUDTORG = 'KHLSKF' AND TRANSTYPE=1
"""

revv1 = pd.read_sql_query(mtdquery1, connection)
revv2 = pd.read_sql_query(mtdquery2, connection)

mtdlastinv = int(revv1['LASTINVOICE'])
mtdlastsales = float(revv1['LASTSALES'])
mtdlastmonth = str(revv1['LASTMONTH'].item())

mtdthisinv = int(revv2['THISINVOICE'])
mtdthissales = float(revv2['THISSALES'])
mtdthismonth = str(revv2['THISMONTH'].item())

mtdinvgrowth = format(((mtdthisinv-mtdlastinv)/mtdlastinv)*100,',.2f')+"%"
mtdinvgrowthnum = ((mtdthisinv-mtdlastinv)/mtdlastinv)*100
mtdsalesgrowth = format(((mtdthissales-mtdlastsales)/mtdlastsales)*100,',.2f')+"%"
mtdsalesgrowthnum = ((mtdthissales-mtdlastsales)/mtdlastsales)*100

#mtdlastinv = format(mtdlastinv, ',d')
#mtdlastsales = format(mtdlastsales, ',.2f')
#mtdthisinv = format(mtdthisinv, ',d')
#mtdthissales = format(mtdthissales, ',.2f')


returnquery1 = """
SELECT COUNT(DISTINCT INVNUMBER) AS LASTLASTINVOICE, SUM(EXTINVMISC) as LASTLASTSALES, FORMAT(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE())-2, 0),'MMMM') AS LASTLASTMONTH
                                   FROM OESalesDetails
                                   WHERE transdate between FORMAT(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE())-1, 0),'yyyyMMdd')
								   AND FORMAT(DATEADD(month, -1, GETDATE()-1),'yyyyMMdd') AND AUDTORG = 'KHLSKF' AND TRANSTYPE=2
"""

returnquery2 = """
SELECT COUNT(DISTINCT INVNUMBER) AS LASTINVOICE, SUM(EXTINVMISC) as LASTSALES, FORMAT(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE())-1, 0),'MMMM') AS LASTMONTH
                                    FROM OESalesDetails
                                    WHERE transdate between FORMAT(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()), 0),'yyyyMMdd')
									AND FORMAT(GETDATE()-1,'yyyyMMdd') AND AUDTORG = 'KHLSKF' AND TRANSTYPE=2
"""

ret1 = pd.read_sql_query(returnquery1, connection)
ret2 = pd.read_sql_query(returnquery2, connection)

mtdlastlastretinv = int(ret1['LASTLASTINVOICE'])
mtdlastlastretsales = float(ret1['LASTLASTSALES'])*(-1)
mtdlastlastretmonth = str(ret1['LASTLASTMONTH'].item())

mtdlastretinv = int(ret2['LASTINVOICE'])
mtdlastretsales = float(ret2['LASTSALES'])*(-1)
mtdlastretmonth = str(ret2['LASTMONTH'].item())

mtdretinvgrowth = format(((mtdlastretinv-mtdlastlastretinv)/mtdlastlastretinv)*100,',.2f')+"%"
mtdretinvgrowthnum = ((mtdlastretinv-mtdlastlastretinv)/mtdlastlastretinv)*100
mtdretsalesgrowth = format(((mtdlastretsales-mtdlastlastretsales)/mtdlastlastretsales)*100,',.2f')+"%"
mtdretsalesgrowthnum = ((mtdlastretsales-mtdlastlastretsales)/mtdlastlastretsales)*100

#mtdlastlastretinv = format(mtdlastlastretinv, ',d')
#mtdlastlastretsales = format(mtdlastlastretsales, ',.2f')
#mtdlastretinv = format(mtdlastretinv, ',d')
#mtdlastretsales = format(mtdlastretsales, ',.2f')

mtdreturnquery1 = """
SELECT COUNT(DISTINCT INVNUMBER) AS LASTLASTINVOICE, SUM(EXTINVMISC) as LASTLASTSALES, FORMAT(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE())-2, 0),'MMMM') AS LASTLASTMONTH
                                   FROM OESalesDetails
                                   WHERE transdate between FORMAT(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE())-2, 0),'yyyyMMdd')
								   AND FORMAT(DATEADD(MONTH, DATEDIFF(MONTH, -1, GETDATE())-2, -1),'yyyyMMdd') AND AUDTORG = 'KHLSKF' AND TRANSTYPE=2
"""

mtdreturnquery2 = """
SELECT COUNT(DISTINCT INVNUMBER) AS LASTINVOICE, SUM(EXTINVMISC) as LASTSALES, FORMAT(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE())-1, 0),'MMMM') AS LASTMONTH
                                    FROM OESalesDetails
                                    WHERE transdate between FORMAT(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE())-1, 0),'yyyyMMdd')
									AND FORMAT(DATEADD(MONTH, DATEDIFF(MONTH, -1, GETDATE())-1, -1),'yyyyMMdd') AND AUDTORG = 'KHLSKF' AND TRANSTYPE=2
"""

rett1 = pd.read_sql_query(mtdreturnquery1, connection)
rett2 = pd.read_sql_query(mtdreturnquery2, connection)

lastlastretinv = int(ret1['LASTLASTINVOICE'])
lastlastretsales = float(ret1['LASTLASTSALES'])*(-1)
lastlastretmonth = str(ret1['LASTLASTMONTH'].item())

lastretinv = int(ret2['LASTINVOICE'])
lastretsales = float(ret2['LASTSALES'])*(-1)
lastretmonth = str(ret2['LASTMONTH'].item())

retinvgrowth = format(((lastretinv-lastlastretinv)/lastlastretinv)*100,',.2f')+"%"
retinvgrowthnum = ((lastretinv-lastlastretinv)/lastlastretinv)*100
retsalesgrowth = format(((lastretsales-lastlastretsales)/lastlastretsales)*100,',.2f')+"%"
retsalesgrowthnum = ((lastretsales-lastlastretsales)/lastlastretsales)*100

#lastlastretinv = format(lastlastretinv, ',d')
#lastlastretsales = format(lastlastretsales, ',.2f')
#lastretinv = format(lastretinv, ',d')
#lastretsales = format(lastretsales, ',.2f')

####-----------------------------------------------
month = [lastlastmonth, lastmonth]
sales = [float(lastlastsales/1000000),float(lastsales/1000000)]
invoice = [int(lastlastinv), int(lastinv)]


from matplotlib import pyplot as plt
import numpy as np

plt.figure(100)
ind = np.arange(len(month))
p1 = plt.bar(ind, sales, alpha=1, label = 'Sales', width=0.15, color="#5c20dd")
for bar in p1:
    yval = round(bar.get_height(),2)
    plt.text(bar.get_x()+0.07, yval+0.7, yval, ha="center", va="center", fontsize=12)


if (salesgrowthnum> 0):
    bbox_props = dict(boxstyle="rarrow,pad=1",fc="#70b31a",lw=0.5)    
    plt.text(0.5, max(sales)/2, " Growth  "+str(salesgrowth)+"   " , ha="center", va="center", rotation = 90, color='white'
                    , size=14
                    , bbox=bbox_props)
else:
    bbox_props = dict(boxstyle="rarrow,pad=1", fc="#701046", lw=0.5)    
    plt.text(0.5, max(sales)/2, " Shrinking  "+str(salesgrowth)+"   ", ha="center", va="center", rotation=270, color='white'
                 , size=14
                 , bbox=bbox_props)


#p2 = plt.bar(ind, invoice, alpha=0.5, label = 'Invoice')
plt.title("Sales", fontdict = font1)
plt.xticks(ind, month, color="#dd5c20", fontsize=14)
plt.ylabel("Sales (Million)", fontdict = font)
plt.legend(loc=8)
plt.tight_layout()
plt.savefig("Sales.png")

plt.figure(200)
ind = np.arange(len(month))
p1 = plt.bar(ind, invoice, alpha=1, label = 'Invoice', width=0.15, color="#ddbb20")
for bar in p1:
    yval = round(bar.get_height(),2)
    plt.text(bar.get_x()+0.07, yval+250, yval, ha="center", va="center", fontsize=12)


if (invgrowthnum> 0):
    bbox_props = dict(boxstyle="rarrow,pad=1",fc="#70b31a",lw=0.5)    
    plt.text(0.5, max(invoice)/2, " Growth  "+str(invgrowth)+"   " , ha="center", va="center", rotation = 90, color='white'
                    , size=14
                    , bbox=bbox_props)
else:
    bbox_props = dict(boxstyle="rarrow,pad=1", fc="#701046", lw=0.5)    
    plt.text(0.5, max(invoice)/2, " Shrinking  "+str(invgrowth)+"   ", ha="center", va="center", rotation=270, color='white'
                 , size=14
                 , bbox=bbox_props)


#p2 = plt.bar(ind, invoice, alpha=0.5, label = 'Invoice')
plt.title("Invoice", fontdict = font1)
plt.xticks(ind, month, color="#dd5c20", fontsize=14)
plt.ylabel("Invoice", fontdict = font)
plt.legend(loc=8)
plt.tight_layout()
#plt.show()
plt.savefig("Invoice.png")

###--------**MTD**---------
month = [mtdlastmonth, mtdthismonth]
sales = [float(mtdlastsales/1000000),float(mtdthissales/1000000)]
invoice = [int(mtdlastinv), int(mtdthisinv)]

plt.figure(300)
ind = np.arange(len(month))
p1 = plt.bar(ind, sales, alpha=1, label = 'Sales', width=0.15, color="#5c20dd")
for bar in p1:
    yval = round(bar.get_height(),2)
    plt.text(bar.get_x()+0.07, yval+0.22, yval, ha="center", va="center", fontsize=12)


if (mtdsalesgrowthnum> 0):
    bbox_props = dict(boxstyle="rarrow,pad=1",fc="#70b31a",lw=0.5)    
    plt.text(0.5, max(sales)/2, " Growth  "+str(mtdsalesgrowth)+"   ", ha="center", va="center", rotation = 90, color='white'
                    , size=12
                    , bbox=bbox_props)
else:
    bbox_props = dict(boxstyle="rarrow,pad=1", fc="#701046", lw=0.5)    
    plt.text(0.5, max(sales)/2, " Shrinking  "+str(mtdsalesgrowth)+"   ", ha="center", va="center", rotation=270, color='white'
                 , size=12
                 , bbox=bbox_props)


#p2 = plt.bar(ind, invoice, alpha=0.5, label = 'Invoice')
plt.title("Sales-MTD", fontdict = font1)
plt.xticks(ind, month, color="#dd5c20", fontsize=14)
plt.ylabel("Sales (Million)", fontdict = font)
plt.legend(loc=8)
plt.tight_layout()
plt.savefig("mtdSales.png")

plt.figure(400)
ind = np.arange(len(month))
p1 = plt.bar(ind, invoice, alpha=1, label = 'Invoice', width=0.15, color="#ddbb20")
for bar in p1:
    yval = round(bar.get_height(),2)
    plt.text(bar.get_x()+0.07, yval+90, yval, ha="center", va="center", fontsize=12)


if (mtdinvgrowthnum> 0):
    bbox_props = dict(boxstyle="rarrow,pad=1",fc="#70b31a",lw=0.5)    
    plt.text(0.5, max(invoice)/2, " Growth  "+str(mtdinvgrowth)+"   ", ha="center", va="center", rotation = 90, color='white'
                    , size=12
                    , bbox=bbox_props)
else:
    bbox_props = dict(boxstyle="rarrow,pad=1", fc="#701046", lw=0.5)    
    plt.text(0.5, max(invoice)/2, " Shrinking  "+str(mtdinvgrowth)+"   ", ha="center", va="center", rotation=270, color='white'
                 , size=12
                 , bbox=bbox_props)


#p2 = plt.bar(ind, invoice, alpha=0.5, label = 'Invoice')
plt.title("Invoice-MTD", fontdict = font1)
plt.xticks(ind, month, color="#dd5c20", fontsize=14)
plt.ylabel("Invoice", fontdict = font)
plt.legend(loc=8)
plt.tight_layout()
#plt.show()
plt.savefig("mtdInvoice.png")


###--------**Return**---------
month = [lastlastretmonth, lastretmonth]
sales = [float(lastlastretsales/1000),float(lastretsales/1000)]
invoice = [int(lastlastretinv), int(lastretinv)]

plt.figure(401)
ind = np.arange(len(month))
p1 = plt.bar(ind, sales, alpha=1, label = 'Sales', width=0.15, color="#5c20dd")
for bar in p1:
    yval = round(bar.get_height(),2)
    plt.text(bar.get_x()+0.07, yval+0.07, yval, ha="center", va="center", fontsize=12)


if (retsalesgrowthnum> 0):
    bbox_props = dict(boxstyle="rarrow,pad=1",fc="#70b31a",lw=0.5)    
    plt.text(0.5, max(sales)/2, " Growth  "+str(mtdsalesgrowth)+"   ", ha="center", va="center", rotation = 90, color='white'
                    , size=12
                    , bbox=bbox_props)
else:
    bbox_props = dict(boxstyle="rarrow,pad=1", fc="#701046", lw=0.5)    
    plt.text(0.5, max(sales)/2, " Shrinking  "+str(mtdsalesgrowth)+"   ", ha="center", va="center", rotation=270, color='white'
                 , size=12
                 , bbox=bbox_props)


#p2 = plt.bar(ind, invoice, alpha=0.5, label = 'Invoice')
plt.title("Sales Return", fontdict = font1)
plt.xticks(ind, month, color="#dd5c20", fontsize=14)
plt.ylabel("Sales (Thousand)", fontdict = font)
plt.legend(loc=8)
plt.tight_layout()
plt.savefig("retSales.png")

plt.figure(402)
ind = np.arange(len(month))
p1 = plt.bar(ind, invoice, alpha=1, label = 'Invoice', width=0.15, color="#ddbb20")
for bar in p1:
    yval = round(bar.get_height(),2)
    plt.text(bar.get_x()+0.07, yval+0.30, yval, ha="center", va="center", fontsize=12)


if (retinvgrowthnum> 0):
    bbox_props = dict(boxstyle="rarrow,pad=1",fc="#70b31a",lw=0.5)    
    plt.text(0.5, max(invoice)/2, " Growth  "+str(mtdinvgrowth)+"   ", ha="center", va="center", rotation = 90, color='white'
                    , size=12
                    , bbox=bbox_props)
else:
    bbox_props = dict(boxstyle="rarrow,pad=1", fc="#701046", lw=0.5)    
    plt.text(0.5, max(invoice)/2, " Shrinking  "+str(mtdinvgrowth)+"   ", ha="center", va="center", rotation=270, color='white'
                 , size=12
                 , bbox=bbox_props)


#p2 = plt.bar(ind, invoice, alpha=0.5, label = 'Invoice')
plt.title("Invoice Return", fontdict = font1)
plt.xticks(ind, month, color="#dd5c20", fontsize=14)
plt.ylabel("Invoice", fontdict = font)
plt.legend(loc=8)
plt.tight_layout()
#plt.show()
plt.savefig("retInvoice.png")


from PIL import Image
image1 = Image.new('RGB', (1281, 480*3+2))

im1 = Image.open(dirpath+"/Sales.png")
im2 = Image.open(dirpath+"/Invoice.png")
im3 = Image.open(dirpath+"/mtdSales.png")
im4 = Image.open(dirpath+"/mtdInvoice.png")
im5 = Image.open(dirpath+"/retSales.png")
im6 = Image.open(dirpath+"/retInvoice.png")

width, height = im1.size

image1.paste(im1,(0,0))
image1.paste(im2,(width+1,0))
image1.paste(im3,(0,height+1))
image1.paste(im4,(width+1,height+1))
image1.paste(im5,(0,height*2+2))
image1.paste(im6,(width+1,height*2+2))

image1.save(dirpath+"/new.png")