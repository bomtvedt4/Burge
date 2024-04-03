""" AR Issues
The goals of this program are to save time and create a structured
protocol for following up with customers regarding orders on hold
for due to need for prepayment of if an open account has gone past due.
Follow-ups should be every 2-3 business days.
"""
import numpy as np
import time
from openpyxl import Workbook, load_workbook
import types
import os
import warnings
import datetime
import pyperclip
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
import win32com.client as win32
from itoolkit.transport import DatabaseTransport
import pyodbc
#imports

def jobs():#this function pulls in all of the jobs that match prepayment status
    #our sql query
    query = """

    SELECT
        OCNHDRPF.ACCTNO, OCNHDRPF.CONTRL, OCNHDRPF.PONUM, OCNHDRPF.OHRDAT,
        ACTFILPF.NAME, ACTFILPF.STRDAT

        FROM DTABUD.OCNHDRPF OCNHDRPF, DTABUD.ACTFILPF ACTFILPF

        WHERE OCNHDRPF.ACCTNO = ACTFILPF.ACCTNO
            AND OCNHDRPF.CRDSTA = 'C'
            AND OCNHDRPF.ACCTNO NOT IN (99999,99996,9999103,99995)
            AND ACTFILPF.CRDCOD IN ('P','W','B')
            AND OCNHDRPF.OHBUD = 1

        ORDER BY OCNHDRPF.CONTRL DESC
"""
    #iniatiate our connection and run our query code from above
    conn = pyodbc.connect(REDACTED)
    itransport = DatabaseTransport(conn)
    cursor = conn.cursor()
    cursor = cursor.execute(query)
    #cleans our query info and places it into a python list
    for row in cursor:
        final = list(row)
        for i in range(0,len(final)):
            if isinstance(final[i],str):
                while final[i].endswith(' '):
                    final[i] = final[i][:-1]
        master.append(final)


def arEmail(account,index):#this function pulls in AR contact emails
    #our sql query
    query = """
    SELECT
        ACCTCNTPF.EMAILADR

        FROM DTABUD.ACCTCNTPF ACCTCNTPF

        WHERE ACCTCNTPF.USEDFOR = 'AR'
            AND ACCTCNTPF.ACCTNBR = """+str(account)
    #iniatiate our connection and run our query code from above
    conn = pyodbc.connect(REDACTED)
    itransport = DatabaseTransport(conn)
    cursor = conn.cursor()
    cursor = cursor.execute(query)
    #cleans our query info and places it into our master list
    for row in cursor:
        emails.append(row[0].replace(' ',''))


def ordEmail(control,index):
    #our sql query
    query = """
    SELECT
        ORDCNTPF.EMAILADR

        FROM DTABUD.ORDCNTPF ORDCNTPF

        WHERE ORDCNTPF.ORDERNBR = """+str(control)
    #iniatiate our connection and run our query code from above
    conn = pyodbc.connect(REDACTED)
    itransport = DatabaseTransport(conn)
    cursor = conn.cursor()
    cursor = cursor.execute(query)
    #cleans our query info and places it into our master list
    for row in cursor:
        if '@3mpromote.com' not in row[0]:
            emails.append(row[0].replace(' ',''))


def sendEmail(masterList):
    #function that sends emails based on "flagged" orders
    import datetime#coming back to this after the fact, not sure if this line is neccessary
    #get proper date
    current_day = datetime.date.today()
    formatted_date = datetime.date.strftime(current_day, "%m/%d/%Y")
    date = formatted_date.replace('/','-')
    email = 'REDACTED'
    server = smtplib.SMTP(REDACTED)
    pb1 = '\n'
    pb2 = '\n' + '\n'
    #formatted email body that goes out
    body1 = 'Thank you for your order and we appreciate your business!'+pb2+'We have not received prepayment for your order(PO#'
    body2 = '). Please call REDACTED with the prepayment information indicating your method of payment. Upon receiving the prepayment, we will release the credit hold and proceed with your order.'+pb2
    body3 = 'Please respond promptly so that we can prevent any unnecessary delays to your order. Any questions/concerns please contact us at REDACTED.'+pb2
    signature = 'Kind Regards,'+pb2+'Accounts Receivable Department'+pb1+'REDACTED'+pb1+'REDACTED'+pb1+'REDACTED'
    for item in masterList:
        #for every item that is flagged, get emails and send em out!
        if item[7] == True:
            emailList = ''
            for Email in item[6]:
                emailList = emailList + Email + ','
            acct = item[0]
            control = item[1]
            po = item[2]
            acctname = item[4]
            msg = MIMEMultipart()
            msg['From'] = email
            msg['To'] = emailList
            #msg['To'] = 'REDACTED'
            msg['Cc'] = email
            msg['Subject'] = 'Prepayment followup - PO# ' + po + ' Order# ' + str(control)
            msg.attach(MIMEText(date+pb1+'Acct #'+str(acct)+' - '+acctname+pb2+body1+po+body2+body3+signature, 'plain'))
            server.send_message(msg)
            del msg
    server.quit()                  

def dateCalc():
    #function responsible for getting calculated dates
    from datetime import datetime
    now = datetime.now()
    file = open(r'REDACTED', 'r')
    holidaysList = file.read()
    holidaysList = holidaysList.replace('\n','').split('|')
    file.close()
    #misrepresentation of order can happen if account is not live before order
    for item in master:
        if item[3] < item[5]:
            item[3] = item[5]
    total = len(master)
    step = 0
    #iterate through our list and decide whether or not it should be sent a followup
    while step != total:
        for item in master:
            count = np.busday_count(item[3],now.date(),holidays=holidaysList)
            if count != 0 and count%2 == 0:
                pass
            else:
                item[7] = False
        step = step + 1
            

def holidayEdit(month,day,year):
    #add a date date
    #needs to be converted to a proper format when used elsewhere
    convert = year+'-'+month+'-'+day
    f = open(r'REDACTED', 'a')
    f.write('|\n'+convert)
    f.close()


def main():
    #main function that returns a default list to Burge
    global master,emails
    master = []
    jobs()
    for item in master:
        emails = []
        control = item[1]
        account = item[0]
        index = master.index(item)
        arEmail(account,index)
        ordEmail(control,index)
        emails = list(set(emails))
        master[index].append(emails)
        master[index].append(True)
    dateCalc()
    return master

