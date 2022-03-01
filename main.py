import os
import requests
from bs4 import BeautifulSoup as BS
import smtplib
import xlsxwriter
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


#constants
template="https://nithp.herokuapp.com/result/student?roll="
sender=os.environ.get('test_email')
sender_password=os.environ.get('test_password')
location=os.environ.get('test_location')
file_name='Result.xlsx'

#function to build the url based on the roll number given
def build_url():
    print("Roll. No.: ")
    x=input()
    roll=x.upper()
    return template+roll
 

#function to create a excel workbook and write the result data to it
def write_workbook(table1,table2,table3):
    workbook=xlsxwriter.Workbook(file_name)
    sheet=workbook.add_worksheet('Result')
    sheet.set_column_pixels(0,100,120)
    merge_format_heading=workbook.add_format({
        'align':'center',
        'valign':'vcenter',
        'fg_color':'#75a3a3',
        'size':20,
        'border':1,
        'bold':True,
    })
    sheet.merge_range('H4:M6','Student Information',merge_format_heading)

    merge_format_title=workbook.add_format({
        'align':'center',
        'valign':'vcenter',
        'fg_color':'#a55b35',
        'size':16,
        'border':1,
        'bold':True,
    })
    sheet.merge_range('H7:I8','Roll. No.',merge_format_title)
    sheet.merge_range('J7:K8','Name',merge_format_title)
    sheet.merge_range('L7:M8','Department',merge_format_title)

    merge_format_info=workbook.add_format({
        'align':'center',
        'valign':'vcenter',
        'fg_color':'#feebb8',
        'size':12,
        'border':1, 
    })    
    sheet.merge_range('H9:I10',table1[0],merge_format_info)
    sheet.merge_range('J9:K10',table1[1],merge_format_info)
    sheet.merge_range('L9:M10',table1[2],merge_format_info)

    sheet.merge_range('H14:M16','Semester-wise Result',merge_format_heading)
    sheet.merge_range('H17:I18','Semester',merge_format_title)
    sheet.merge_range('J17:K18','SGPI',merge_format_title)
    sheet.merge_range('L17:M18','CGPI',merge_format_title)

    start=19
    for i in range(len(table2)):
        sheet.merge_range('H'+str(start)+':I'+str(start+1),float(table2[i][0]),merge_format_info)
        sheet.merge_range('J'+str(start)+':K'+str(start+1),float(table2[i][1]),merge_format_info)
        sheet.merge_range('L'+str(start)+':M'+str(start+1),float(table2[i][2]),merge_format_info)
        start+=2

    start+=3
    # print(start)
    sheet.merge_range('G'+str(start)+':N'+str(start+2),'Subject-wise Result',merge_format_heading)
    # sheet.merge_range()
    start+=3
    # print(start)
    sheet.merge_range('G'+str(start)+':G'+str(start+1),'Semester',merge_format_title)
    sheet.merge_range('H'+str(start)+':H'+str(start+1),'Subject Code',merge_format_title)
    sheet.merge_range('I'+str(start)+':K'+str(start+1),'Subject',merge_format_title)
    sheet.merge_range('L'+str(start)+':L'+str(start+1),'Grade',merge_format_title)
    sheet.merge_range('M'+str(start)+':M'+str(start+1),'Subject GP',merge_format_title)
    sheet.merge_range('N'+str(start)+':N'+str(start+1),'Subject Point',merge_format_title)
    start+=2
    # print(start)
    for i in range(len(table3)):
        sheet.merge_range('G'+str(start)+':G'+str(start+1),float(table3[i][0]),merge_format_info)
        sheet.merge_range('H'+str(start)+':H'+str(start+1),table3[i][1],merge_format_info)
        sheet.merge_range('I'+str(start)+':K'+str(start+1),table3[i][2],merge_format_info)
        sheet.merge_range('L'+str(start)+':L'+str(start+1),table3[i][3],merge_format_info)
        sheet.merge_range('M'+str(start)+':M'+str(start+1),float(table3[i][4]),merge_format_info)
        sheet.merge_range('N'+str(start)+':N'+str(start+1),float(table3[i][5]),merge_format_info)
        start+=2

    workbook.close()


#function to write into the email and send the "Result" excel file to the required email address
def send_mail():
    print("Email: ")
    x=input()
    receiver=x.lower()
    msg=MIMEMultipart()
    msg['From']=sender
    msg['To']=receiver
    msg['Subject']="Result"
    body="Your result is attached with this mail"
    msg.attach(MIMEText(body,'plain'))
    filename="Result.xlsx"
    attachment=open(f"{location}{file_name}",'rb')
    part=MIMEBase('application','octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition','attachment;filename="%s"'%filename)
    msg.attach(part)
    s=smtplib.SMTP('smtp.gmail.com',587)
    s.starttls()

    s.login(sender,sender_password)
    s.sendmail(sender,receiver,msg.as_string())
    s.quit()



if __name__=='__main__':
    url=build_url()
    print(url)
    page=requests.get(url)
    status=page.status_code
    if(status!=200):
        print("Invalid Student!!!")
        exit()

    html_content=requests.get(url).text
    soup=BS(html_content, 'lxml')
    table=soup.find('table',attrs={'class': 'table-responsive'})
    table1_info=[]
    tds=table.tbody.tr.find_all('td')
    for i in range(0,tds.__len__()):{
        table1_info.append(tds[i].text)
    }

    
    table=soup.find('table',attrs={'class': 'table-striped'})
    table2_info=[]
    trs=table.tbody.find_all('tr')
    for i in range(0,trs.__len__()):
        tds=trs[i].find_all('td')
        table2_info.append([tds[0].text,tds[1].text,tds[2].text])
    

    table=soup.find('table',attrs={'id': 'showResult'})
    table3_info=[]
    trs=table.tbody.find_all('tr')
    for i in range(0,trs.__len__()):
        tds=trs[i].find_all('td')
        temp=[]
        for j in range(0,tds.__len__()):
            temp.append(tds[j].text)
        table3_info.append(temp)

    table3_info.sort()
    table3_info.reverse()

    write_workbook(table1_info,table2_info,table3_info)
    send_mail()
