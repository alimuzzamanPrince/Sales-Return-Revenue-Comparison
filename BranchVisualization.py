import data as dt
import Tables as tbl
from matplotlib import pyplot as plt
import numpy as np
import os
from PIL import Image
import smtplib, ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from datetime import datetime
import locale
locale.setlocale(locale.LC_ALL, 'en_US')

#----------------------------------------------------
dirpath = os.path.dirname(os.path.realpath(__file__))
x = 0
stockstatus = ['SUS', 'US', 'NS', 'OS', 'SOS', 'NIL']
stockstatus1 = ['NIL', 'SUS', 'US', 'NS', 'OS', 'SOS']
font = {'family': 'serif',
        'color':  '#3b5998',
        'weight': 400,
        'size': 15,
        }

font1 = {'family': 'serif',
        'color':  '#3b5998',
        'weight': 700,
        'size': 17,
        }

colors = ['#ff0000', '#e6e600', '#008000', '#00ff16', 'darkorange', '#990000']
days = np.arange(1,31,1)
#----------------------------------------------------

plt.figure(x)
wtf = plt.pie(dt.statusdata, colors=colors, autopct='%1.1f%%', pctdistance=1.2)
plt.title('Khulna Branch', fontdict=font1)
####plt.legend(labels = stockstatus, title="Status", loc="lower center",bbox_to_anchor=(1, 0, 0.5, 1))
###plt.legend(labels = stockstatus, title="Status", loc="lower left", mode = "expand", ncol = 12)
###plt.legend(labels = stockstatus, bbox_to_anchor=(1, 0, 0.5, 1), loc='1')
plt.legend(labels = stockstatus, bbox_to_anchor=(1, 0, 0.5, 1), loc=3)

fraction_text_list = wtf[2]
for text in fraction_text_list:
    text.set_rotation(0)
    text.set_size(12)
plt.savefig(dirpath+'/branchpie.png')
x = x+1

plt.figure(x)
y_pos = np.arange(len(stockstatus))

bars = plt.bar(y_pos, dt.statusdata1, align='center', alpha=1)
for bar in bars:
    yval = bar.get_height()
    plt.text(bar.get_x()+0.42, yval + 10.25, yval, horizontalalignment='center', verticalalignment='center', fontsize=14)

plt.xticks(y_pos, stockstatus1, rotation=0, fontsize=12)
plt.yticks(np.arange(0, 400, 20), fontsize=12)
plt.ylabel('No. of Item', fontdict=font)
plt.title('Khulna Branch - Stock Count('+str(dt.branchitem)+' of '+str(dt.totalitem)+')', fontdict=font1)
plt.tight_layout()
bars[0].set_color('#990000')
bars[1].set_color('#ff0000')
bars[2].set_color('#e6e600')
bars[3].set_color('#008000')
bars[4].set_color('#00ff16')
bars[5].set_color('darkorange')
plt.savefig(dirpath+'/branchbar.png')
x = x+1

plt.figure(x)

plt.plot(days, dt.SUS, color = '#ff0000', marker = '*', markerfacecolor='green')
plt.plot(days, dt.US, color = '#e6e600', marker = '*', markerfacecolor='green')
plt.plot(days, dt.NS, color = '#008000', marker = '*', markerfacecolor='green')
plt.plot(days, dt.OS, color = '#00ff16', marker = '*', markerfacecolor='green')
plt.plot(days, dt.SOS, color = 'darkorange', marker = '*', markerfacecolor='green')
plt.plot(days, dt.NIL, color = '#990000', marker = '*', markerfacecolor='green')

plt.xticks(range(1,31), dt.AUDTDATE, rotation=90)
plt.ylabel('No. of Item', fontdict=font)
plt.xlabel('Days', fontdict=font)
plt.title('Khulna Branch\'s Last 30 Days Status', fontdict=font1)
plt.legend(['SUS','US','NS','OS','SOS','NIL'], loc='best')
plt.savefig(dirpath+'/branchlinenew.png')

x = x+1

plt.figure(x)
plt.plot(days, dt.SUS, color = '#ff0000', marker = '*', markerfacecolor='green')

plt.xticks(range(1,31), dt.AUDTDATE, rotation=90)
plt.yticks(np.arange(0,145,5))
plt.ylabel('No. of Item', fontdict=font)
plt.xlabel('Days', fontdict=font)
plt.title('Khulna Branch\'s Last 30 Days SUS Status', fontdict=font1)
plt.legend(['SUS'], loc='best')
plt.savefig(dirpath+'/susline.png')
x=x+1

plt.figure(x)
plt.plot(days, dt.US, color = '#e6e600', marker = '*', markerfacecolor='green')

plt.xticks(range(1,31), dt.AUDTDATE, rotation=90)
plt.yticks(np.arange(0,145,5))
plt.ylabel('No. of Item', fontdict=font)
plt.xlabel('Days', fontdict=font)
plt.title('Khulna Branch\'s Last 30 Days US Status', fontdict=font1)
plt.legend(['US'], loc='best')
plt.savefig(dirpath+'/usline.png')
x=x+1

plt.figure(x)
plt.plot(days, dt.NS, color = '#008000', marker = '*', markerfacecolor='red')

plt.xticks(range(1,31), dt.AUDTDATE, rotation=90)
plt.yticks(np.arange(0,145,5))
plt.ylabel('No. of Item', fontdict=font)
plt.xlabel('Days', fontdict=font)
plt.title('Khulna Branch\'s Last 30 Days NS Status', fontdict=font1)
plt.legend(['NS'], loc='best')
plt.savefig(dirpath+'/nsline.png')
x=x+1

plt.figure(x)
plt.plot(days, dt.OS, color = '#00ff16', marker = '*', markerfacecolor='red')

plt.xticks(range(1,31), dt.AUDTDATE, rotation=90)
plt.yticks(np.arange(0,145,5))
plt.ylabel('No. of Item', fontdict=font)
plt.xlabel('Days', fontdict=font)
plt.title('Khulna Branch\'s Last 30 Days OS Status', fontdict=font1)
plt.legend(['OS'], loc='best')
plt.savefig(dirpath+'/osline.png')
x=x+1

plt.figure(x)
plt.plot(days, dt.SOS, color = 'darkorange', marker = '*', markerfacecolor='red')

plt.xticks(range(1,31), dt.AUDTDATE, rotation=90)
plt.yticks(np.arange(200,510,10))
plt.ylabel('No. of Item', fontdict=font)
plt.xlabel('Days', fontdict=font)
plt.title('Khulna Branch\'s Last 30 Days SOS Status', fontdict=font1)
plt.legend(['SOS'], loc='best')
plt.savefig(dirpath+'/sosline.png')
x=x+1

plt.figure(x)
plt.plot(days, dt.NIL, color = '#990000', marker = '*', markerfacecolor='green')

plt.xticks(range(1,31), dt.AUDTDATE, rotation=90)
plt.yticks(np.arange(0,145,5))
plt.ylabel('No. of Item', fontdict=font)
plt.xlabel('Days', fontdict=font)
plt.title('Khulna Branch\'s Last 30 Days NIL Status', fontdict=font1)
plt.legend(['NIL'], loc='best')
plt.savefig(dirpath+'/nilline.png')
x=x+1

image1 = Image.new('RGB', (1281, 480))
im1 = Image.open(dirpath+"/branchpie.png")
im2 = Image.open(dirpath+"/branchbar.png")
image1.paste(im1,(0,0))
image1.paste(im2,(641,0))
image1.save(dirpath+"/branchpiebar.png")

width, height = im1.size
image2 = Image.new('RGB', (1281, 1442))
im3 = Image.open(dirpath+"/susline.png")
im4 = Image.open(dirpath+"/usline.png")
im5 = Image.open(dirpath+"/nsline.png")
im6 = Image.open(dirpath+"/osline.png")
im7 = Image.open(dirpath+"/sosline.png")
im8 = Image.open(dirpath+"/nilline.png")
image2.paste(im8,(0,0))
image2.paste(im5,(width+1,0))
image2.paste(im4,(0,height+1))
image2.paste(im3,(width+1,height+1))
image2.paste(im6,(0,height*2+2))
image2.paste(im7,(width+1,height*2+2))
image2.save(dirpath+"/branchline.png")

#im9 = Image.open(dirpath+"/branchlinenew.png")
#im9 = im9.resize((1281,480))
#print(im9.size)
#im9.save(dirpath+"/branchlinenew.png")


me  = 'erp-bi.service@transcombd.com'
recipient = 'm.alimuzzaman@transcombd.com'

date = datetime.today()
today = str(date.day)+'-'+str(date.strftime("%b"))+'-'+str(date.year)+' '+date.strftime("%I:%M %p")

today1 = str(date.day)+' '+str(date.strftime("%B"))+' '+str(date.year)+' at '+date.strftime("%I:%M %p")

#####subject = 'Graph Report'
subject="SK+F Formulation – Stock Status Report - "+today

email_server_host = 'mail.transcombd.com'
port = 25

msgRoot = MIMEMultipart('related')
msgRoot['Subject'] = subject
msgRoot['From'] = me
msgRoot['To'] = recipient
msgRoot.preamble = 'This is a multi-part message in MIME format.'

## Encapsulate the plain and HTML versions of the message body in an
## 'alternative' part, so message agents can decide which they want to display.
msgAlternative = MIMEMultipart('alternative')
msgRoot.attach(msgAlternative)

msgText = MIMEText('This is the alternative plain text message.')
msgAlternative.attach(msgText)

## We reference the image in the IMG SRC attribute by the ID we give it below
msgText = MIMEText("""<b>Dear Sir,</b><br><br>Enclosed, Please find herewith the Stock Status report of """
                   + today1 + """ for <b>Khulna</b> Branch.<br><br>"""

"""
<table border="0" width=100%>
  <tr>
    <font size="2" face="Tahoma" color="blue"><b><td width=50%>Total Item Available in TDCL: """+str(dt.totalitem)+"""</td></b></font>
    <font size="2" face="Tahoma" color = "#49396c"><b><td>Total No. of back order item yesterday: """+str(dt.lastdayitem)+"""</td></b></font>
  </tr>
  <tr>
    <font size="2" face="Tahoma" color = "green"><b><td>Total Item Available in Khulna: """+str(dt.branchitem)+"""</td></b></font>
    <font size="2" face="Tahoma" color = "#e2350a"><b><td>Total sales opportunity lost yesterday: """+str(dt.lastdayvalue)+"""</td></b></font>
  </tr>
   <tr>
    <font size="2" face="Tahoma" color = "red"><b><td>Total No. of item not available in Khulna: """+str(dt.missingitem)+"""</td></b></font>
    <font size="2" face="Tahoma" color = "#242e3c"><b><td>Total No. of back order item in last month: """+str(dt.lastmonthitem)+"""</td></b></font>
  </tr>
  <tr>
    <td></td>
    <font size="2" face="Tahoma" color = "#990000"><b><td>Total sales opportunity lost in last month: """+str(dt.lastmonthvalue)+"""</td></b></font>
  </tr>
</table>
"""
                   """<br><br><img src="cid:image4"><br><br>"""
                   """<br><br><img src="cid:image1"><br><br>"""
                   """<img src="cid:image3"><br><br>  
                   <img src="cid:image2"><br><br>   
                   If there is any inconvenience, you are requested to communicate"""
                   """with the ERP BI Service:<br><b>(Mobile: 01713-389972, 01713-380499)</b><br>"""
                   """<br>Regards<br><b>ERP BI Service</b><br>Information System Automation (ISA)<br> <br>
                   <i><font color="blue">****This is a system generated stock report ****</i></font>""", 'html')

msgAlternative.attach(msgText)

## Assigning the image directory
fp = open(dirpath+'/new.png', 'rb')
msgImage4 = MIMEImage(fp.read())
fp.close()

## Define the image's ID as referenced above
msgImage4.add_header('Content-ID', '<image4>')
msgRoot.attach(msgImage4)

## Assigning the image directory
fp = open(dirpath+'/branchpiebar.png', 'rb')
msgImage1 = MIMEImage(fp.read())
fp.close()

## Define the image's ID as referenced above
msgImage1.add_header('Content-ID', '<image1>')
msgRoot.attach(msgImage1)

# Assigning the image directory
fp = open(dirpath+'/branchline.png', 'rb')
msgImage2 = MIMEImage(fp.read())
fp.close()

# Define the image's ID as referenced above
msgImage2.add_header('Content-ID', '<image2>')
msgRoot.attach(msgImage2)

# Assigning the image directory
fp = open(dirpath+'/branchlinenew.png', 'rb')
msgImage3 = MIMEImage(fp.read())
fp.close()

# Define the image's ID as referenced above
msgImage3.add_header('Content-ID', '<image3>')
msgRoot.attach(msgImage3)

server = smtplib.SMTP(email_server_host, port)
server.ehlo()
print('Sending Mail')
server.sendmail(me, recipient, msgRoot.as_string())
print('Mail Sent')
server.close()