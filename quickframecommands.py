import os
from misc_commands import write_options,get_options,get_apps,get_links
from tkinter import filedialog
from customtkinter import StringVar
import pyperclip
from fpdf import FPDF
from itoolkit.transport import DatabaseTransport
import pyodbc
import decimal
from datetime import datetime
from datetime import timedelta
from datetime import date
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import secrets
import pyperclip
#imports
pb = '\n'
#class for the functions of quick apps
class quickAppsSetup:
    def launch_app_all(as400_username):
        #grabs the entire list of the given username's apps and starts them
        apps = get_apps(get_options(as400_username))
        for item in apps:
            os.startfile(item)

    def launch_app_wrapper(quickAppsLaunch):
        #wrapper to handle passing an arguement
        def launch_app(choice):
            #launches one app of the users choice
            launch = StringVar(value='Launch')
            quickAppsLaunch.configure(variable=launch)
            os.startfile(choice)
        return launch_app
    
    def add_app(as400_username,quickAppsRemove,quickAppsLaunch,prompt):
        #when the user hits "add app" this will copy a useless string and wait
        #for the user to copy a different string, pyperclip library can't properly
        #copy a path to an app so we used tkinters, it then adds it to the user txt file
        pyperclip.copy('burgezxcv,Copy')
        pyperclip.waitForNewPaste()
        file = prompt.clipboard_get()
        
        if file != '' and ',' not in file:
            options = get_options(as400_username)
            app = file
            for i in range(0,len(options)):
                if 'quick_apps=' in options[i]:
                    if options[i].replace('quick_apps=','') == '':
                        options[i] = 'quick_apps=' + app
                    else:
                        options[i] = options[i] + ',' + app
            write_options(as400_username,options)
            apps = get_apps(get_options(as400_username))
            quickAppsRemove.configure(values=apps)
            quickAppsLaunch.configure(values=apps)
            
    def remove_app_wrapper(as400_username,quickAppsRemove,quickAppsLaunch):
        #wrapper to pass args
        def remove_app(choice):
            #removes the app by adjusting the user txt file
            options = get_options(as400_username)
            for i in range(0,len(options)):
                if 'quick_apps=' in options[i]:
                    options[i] = options[i].replace('quick_apps=','')
                    options[i] = options[i].split(',')
                    options[i].remove(choice)
                    options[i] = ','.join(options[i])
                    options[i] = 'quick_apps=' + options[i]
                write_options(as400_username,options)
                apps = get_apps(get_options(as400_username))
                delapp = StringVar(value='Del App')
                quickAppsRemove.configure(values=apps,variable=delapp)
                quickAppsLaunch.configure(values=apps)
        return remove_app

#class for functions of quick links
class quickLinksSetup:
    def launch_link_all(as400_username):
        #launches every link from user txt file
        links = get_links(get_options(as400_username))
        for item in links:
            os.startfile(item)

    def launch_link_wrapper(quickLinksLaunch):
        #wrapper to pass arg
        def launch_link(choice):
            #launches a single selected link from list
            launch = StringVar(value='Launch')
            quickLinksLaunch.configure(variable=launch)
            os.startfile(choice)
        return launch_link
    
    def add_link(as400_username,
                 quickLinksAddEntry,
                 quickLinksRemove,
                 quickLinksLaunch):
        #adds a link by taking user input when button is activated
        file = quickLinksAddEntry.get()
        quickLinksAddEntry.delete(0, 'end')
        options = get_options(as400_username)
        if file != '' and ',' not in file and '\n' not in file: 
            link = file
            for i in range(0,len(options)):
                if 'quick_links=' in options[i]:
                    if options[i].replace('quick_links=','') == '':
                        options[i] = 'quick_links=' + link
                    else:
                        options[i] = options[i] + ',' + link
            write_options(as400_username,options)
            links = get_links(get_options(as400_username))
            quickLinksRemove.configure(values=links)
            quickLinksLaunch.configure(values=links)
            
    def remove_link_wrapper(as400_username,quickLinksRemove,quickLinksLaunch):
        #wrapper to pass args
        def remove_link(choice):
            #removes link from list and edits txt file
            options = get_options(as400_username)
            for i in range(0,len(options)):
                if 'quick_links=' in options[i]:
                    options[i] = options[i].replace('quick_links=','')
                    options[i] = options[i].split(',')
                    options[i].remove(choice)
                    options[i] = ','.join(options[i])
                    options[i] = 'quick_links=' + options[i]
                write_options(as400_username,options)
                dellink = StringVar(value='Del Link')
                links = get_links(get_options(as400_username))
                quickLinksRemove.configure(values=links,variable=dellink)
                quickLinksLaunch.configure(values=links)
        return remove_link

#class of functions for setting mode and text color
class setup:
    def __init__(self,as400_username=None):
        self.as400_username = as400_username

    def set_mode_wrapper(options,as400_username,menu_root,mainMenu):
        #wrapper to pass args
        def set_mode(choice):
            #changes theme to dark or light, JSON files somewhere control this
            #could possibly add to this later
            if choice == 'Dark':
                for i in range(0,len(options)):
                    if 'set_appearance_mode=' in options[i]:
                        options[i] = 'set_appearance_mode=dark' 
            elif choice == 'Light':
                for i in range(0,len(options)):
                    if 'set_appearance_mode=' in options[i]:
                        options[i] = 'set_appearance_mode=light'
            write_options(as400_username,options)
            menu_root.destroy()        
            mainMenu()
        return set_mode

    def set_text_color_wrapper(options,as400_username,menu_root,mainMenu):
        #wrapper to pass args
        def set_text_color(choice):
            #edits the txt file and refreshes the menu with the new text
            default_text = choice
            for i in range(0,len(options)):
                if 'default_text=' in options[i]:
                    options[i] = 'default_text='+default_text       
            write_options(as400_username,options)
            menu_root.destroy()        
            mainMenu()
        return set_text_color

    def set_random_color(options,as400_username,menu_root,mainMenu):
        #a fun little function I added that uses the secrets library to get
        #a random color, after it edits the user default and selects that
        #color as the default text
        default_text = '#'+secrets.token_hex(3)
        for i in range(0,len(options)):
            if 'default_text=' in options[i]:
                options[i] = 'default_text='+default_text       
        write_options(as400_username,options)
        menu_root.destroy()        
        mainMenu()

#class for an invoice generation straight from Burge rather than as400
class invoiceGen:
    def __init__(self,invoice,invoiceYear,as400_username,as400_password):
        #these are all defined by user input earlier
        self.invoice = invoice
        self.invoiceYear = invoiceYear
        self.as400_username = as400_username
        self.as400_password = as400_password
        
    def creditBalance(self):
        def getPaymentInfo(payment_number,payment_year):
            types = [['ACH_DISTR','ACH'],
                     ['ACH_AR','ACH'],
                     ['CHECK','Check'],
                     ['REMOTE','Check'],
                     ['WIRE','Wire'],
                     ['CARD','Card']]
            conn = pyodbc.connect('REDACTED',
                                  uid=self.as400_username,
                                  password=self.as400_password)
            itransport = DatabaseTransport(conn)
            cursor = conn.cursor()
            paymentQuery = f"""
            SELECT
                ARCASHPF.ACHKN, ARCASHPF.PYMTTYPE
                
            FROM
                DTABUD.ARCASHPF ARCASHPF
                
            WHERE
                ARCASHPF.TRANN = {payment_number}
                AND ARCASHPF.TRNYR = {payment_year}
            """
            cursor.execute(paymentQuery)
            for row in cursor:
                final = [elem for elem in row]
                for i in range(0,len(final)):
                    while final[i].endswith(' '):
                        final[i] = final[i][:-1]
                for pType in types:
                    if final[1] == pType[0]:
                        final.append(pType[1])
                if final[2] == 'ACH':
                    final.append(final[2])
                elif final[2] == 'Check':
                    final.append('check #'+final[0])
                elif final[2] == 'Wire':
                    final.append('wire')
                elif final[2] == 'Card':
                    final.append('card '+final[0])
            return final
        #function that runs a query that determines the amount of credit on account
        #applied against this invoice as well as any payments made
        conn = pyodbc.connect('REDACTED',
                              uid=self.as400_username,
                              password=self.as400_password)
        itransport = DatabaseTransport(conn)
        cursor = conn.cursor()
        creditQuery = """
        SELECT
        ARTRAKPF.DELETE, ARTRAKPF.CRDT, ARTRAKPF.APPAMT,
        ARTRAKPF.CRDN, ARTRAKPF.CRDYR
        
        FROM
        DTABUD.ARTRAKPF ARTRAKPF

        WHERE
        ARTRAKPF.DBTT = 'I'
        AND ARTRAKPF.DBTN = """+str(self.invoice)+"""
        AND ARTRAKPF.DBTYR = """+str(self.invoiceYear)+"""
        AND ARTRAKPF.CRDT IN ('C','P')
        AND ARTRAKPF.DELETE = ' '
        """
        cursor = cursor.execute(creditQuery)
        creditValues = []
        creditNumbers = []
        paymentNumbers = []
        creditTotal = 0
        paymentTotal = 0
        for row in cursor:
            final =[elem for elem in row]
            creditValues.append(final)
        for item in creditValues:
            if item[1] == 'P':
                paymentTotal += item[2]
                paymentNumbers.append(getPaymentInfo(item[3],item[4]))
            else:
                creditTotal += item[2]
                creditNumbers.append(str(item[3]))
                
        return float(paymentTotal),float(creditTotal),paymentNumbers,creditNumbers

    def exportCharges(self,invoiceCounter):
        #function that runs a query and calculates if there is any
        #export charges on a counter
        conn = pyodbc.connect('REDACTED',
                              uid=self.as400_username,
                              password=self.as400_password)
        itransport = DatabaseTransport(conn)
        cursor = conn.cursor()
        exportQuery = """
        SELECT
        AININTLEXT.PRODSALES, AININTLEXT.DUTIES,
        AININTLEXT.PRODTXPCT
        
        FROM
        DTABUD.AININTLEXT AININTLEXT

        WHERE
        AININTLEXT.INVNBR = """+str(self.invoice)+"""
        AND AININTLEXT.INVYR = """+str(self.invoiceYear)+"""
        AND AININTLEXT.INVJOB = """+str(invoiceCounter)+"""
        """
        cursor = cursor.execute(exportQuery)
        final = None
        for row in cursor:
            final = [elem for elem in row]
        return final

    def shippingInfo(self):
        #the first function is to get all the packslips and their info
        def shippingDetail(packslip):
            #this function takes the packslip and gets box count
            #whichever address has the highest box count shows on the invoice
            #below the address shows the number of drops or ship to locations
            conn = pyodbc.connect('REDACTED',
                                  uid=self.as400_username,
                                  password=self.as400_password)
            itransport = DatabaseTransport(conn)
            cursor = conn.cursor()

            detailQuery = """

            SELECT
            OPSDTLPF.BOXCNT
            
            FROM
            DTABUD.OPSDTLPF OPSDTLPF
            
            WHERE
            OPSDTLPF.PCKSLP = """+str(packslip)+"""
            AND OPSDTLPF.BOXCNT <> 0
            
            
            """
            shipBoxCount = 0
            cursor = cursor.execute(detailQuery)
            for row in cursor:
                final = [elem for elem in row]
                shipBoxCount += final[0]
            return shipBoxCount
            
        conn = pyodbc.connect('REDACTED',
                              uid=self.as400_username,
                              password=self.as400_password)
        itransport = DatabaseTransport(conn)
        cursor = conn.cursor()

        shipQuery = """

        SELECT
        OPSHDRPF.PCKSLP, OPSHDRPF.TO_NAME, OPSHDRPF.TO_STR1,
        OPSHDRPF.TO_STR2, OPSHDRPF.TO_STR3, OPSHDRPF.TO_CITY,
        OPSHDRPF.TO_SORP, OPSHDRPF.TO_ZIP, OPSHDRPF.PUDATE
        
        FROM
        DTABUD.OPSHDRPF OPSHDRPF
        
        WHERE
        OPSHDRPF.INVNBR = """+str(self.invoice)+"""
        AND OPSHDRPF.INVYR = """+str(self.invoiceYear)+"""
        AND OPSHDRPF.ADRCTR <> 0

        ORDER BY
        OPSHDRPF.PCKSLP
        
        """
        shipHeaders = []
        cursor = cursor.execute(shipQuery)
        for row in cursor:
            final = [elem for elem in row]
            shipHeaders.append(final)
        for j in range(0,len(shipHeaders)):
            for i in range(0,len(shipHeaders[j])):
                if isinstance(shipHeaders[j][i],str):
                    while True:
                        if shipHeaders[j][i].endswith(' '):
                            shipHeaders[j][i] = shipHeaders[j][i][:-1]
                        elif shipHeaders[j][i] == '':
                            break
                        elif shipHeaders[j][i].startswith(' '):
                            shipHeaders[j][i] = shipHeaders[j][i][2:]
                        else:
                            break
        prioBox = 0
        prioChoice = 1
        if shipHeaders != []:
            for item in shipHeaders:
                item.append(int(shippingDetail(item[0])))
            for i in range(0,len(shipHeaders)):
                if shipHeaders[i][9] > prioBox:
                    prioChoice = i
                    prioBox = shipHeaders[i][9]
            return 'Drop Shipments = '+str(len(shipHeaders)),shipHeaders[prioChoice],shipHeaders
        else:
            return 'Drop Shipments = N/A',None

    def header(self):
        #a function that runs a query that returns the majority of our primary
        #information, after doing some cleaning up of the info, it returns a dictionary
        #for later invoice use
        conn = pyodbc.connect('REDACTED',
                              uid=self.as400_username,
                              password=self.as400_password)
        itransport = DatabaseTransport(conn)
        cursor = conn.cursor()

        headerQuery = """

        SELECT
        AINHDRPF.INVNO, AINHDRPF.PONUM, AINHDRPF.ACCTNO,
        AINHDRPF.APOSTD, AINHDRPF.OHRDAT, AINHDRPF.MERCH,
        AINHDRPF.SHIP, AINHDRPF.TAX, AINHDRPF.DISC,
        ACTFILPF.NAME, ACTFILPF.POBOX, ACTFILPF.STR1,
        ACTFILPF.STR2, ACTFILPF.CITY, ACTFILPF.SORP,
        ACTFILPF.ZIP, ACTFILPF.CTYCOD, ACTFILPF.LPHON,
        AINHDRPF.CONTRL, AINHDRPF.AIBUD, CSRREP2PF.CLOUDID,
        ACTFILPF.CRDCOD
        
        FROM
        DTABUD.AINHDRPF AINHDRPF, DTABUD.ACTFILPF ACTFILPF,
        DTABUD.CSRREP2PF CSRREP2PF
        
        WHERE
        AINHDRPF.INVNO = """+str(self.invoice)+"""
        AND AINHDRPF.INVYR = """+str(self.invoiceYear)+"""
        AND AINHDRPF.ACCTNO = ACTFILPF.ACCTNO
        AND CSRREP2PF.REPNO = AINHDRPF.CSR
        AND CSRREP2PF.REPTYP = 'C'
        
        """
        
        cursor = cursor.execute(headerQuery)

        for row in cursor:
            headerInfo = [elem for elem in row]
        if headerInfo[10] != '          ':
            headerInfo[10] = 'PO Box: '+headerInfo[10]
        for i in range(0,len(headerInfo)):
            if isinstance(headerInfo[i],str):
                while True:
                    if headerInfo[i].endswith(' '):
                        headerInfo[i] = headerInfo[i][:-1]
                    elif headerInfo[i] == '':
                        break
                    elif headerInfo[i].startswith(' '):
                        headerInfo[i] = headerInfo[i][2:]
                    else:
                        break
        acctNum = str(headerInfo[2])
        acctName = headerInfo[9]
        pobox = headerInfo[10]
        street1 = headerInfo[11]
        street2 = headerInfo[12]
        if street2 != '':
            street = street1 + pb + street2
        else:
            street = street1
        if pobox != '':
            street = pobox + pb + street
        city = headerInfo[13]
        state = headerInfo[14]
        postal = headerInfo[15]
        csz = city + ', ' + state + ' ' + postal
        phone = str(headerInfo[16]) + '-' + str(headerInfo[17])[0:3]+'-'+str(headerInfo[17])[3:7]
        address = 'Sold to:'+pb+'Acct #'+acctNum+pb+acctName+pb+street+pb+csz+pb+phone
        info = {
                'invoice':str(headerInfo[0]),'po':headerInfo[1],
                'acctNum':str(headerInfo[2]),
                'postDate':str(headerInfo[3].strftime("%m/%d/%y")),
                'ordDate':str(headerInfo[4].strftime("%m/%d/%y")),
                'merch':float(headerInfo[5]),'ship':float(headerInfo[6]),
                'tax':float(headerInfo[7]),'discount':float(headerInfo[8]),
                'acctName':headerInfo[9],'billTo':address,
                'control':str(headerInfo[18]),'bud':headerInfo[19],
                'csrphone':headerInfo[20],'terms':headerInfo[21]
                }

        return info

    def invoiceItems(self):
        #the primary layout of the invoice shows each item counter followed
        #by its associated options/upcharges/adders however you want to call them
        #this function gets all of the items and then runs the options query
        #on each item and returns the two lists
        def getOptions(counter):
            #ran for each counter
            optionsList = []
            optionsQuery = """

            SELECT
            OITOPTPF.OTDSC, AINOPTPF.OUNIT, AINOPTPF.OEPRIC,
            AINOPTPF.ADDSUB
            
            FROM
            DTABUD.AINOPTPF AINOPTPF, DTABUD.OITOPTPF OITOPTPF

            WHERE
            AINOPTPF.INVYR = """+str(self.invoiceYear)+"""
            AND AINOPTPF.INVNO = """+str(self.invoice)+"""
            AND AINOPTPF.ICNTER = """+str(counter)+"""
            AND OITOPTPF.PRCOPT = AINOPTPF.PRCOPT
            AND OITOPTPF.ITMNTP = AINOPTPF.OITYPE
            AND AINOPTPF.INVPRT = 'Y'
            """
            conn = pyodbc.connect('REDACTED',
                                  uid=self.as400_username,
                                  password=self.as400_password)
            itransport = DatabaseTransport(conn)
            cursor = conn.cursor()
            cursor = cursor.execute(optionsQuery)
            for row in cursor:
                final = [elem for elem in row]
                optionsList.append(final)
            for j in range(0,len(optionsList)):
                for i in range(0,len(optionsList[j])):
                    if isinstance(optionsList[j][i],str):
                        while True:
                            if optionsList[j][i].endswith(' '):
                                optionsList[j][i] = optionsList[j][i][:-1]
                            elif optionsList[j][i] == '':
                                break
                            elif optionsList[j][i].startswith(' '):
                                optionsList[j][i] = optionsList[j][i][2:]
                            else:
                                break
            for item in optionsList:
                if item[3] == 'S':
                    item[2] *= -1
            return optionsList
            
        itemsQuery = """

        SELECT
        AINDTLPF.ICNTER, OITEMSPF.ITMDSC, AINDTLPF.MLINE,
        AINDTLPF.UBILL, AINDTLPF.EPRICE, AINDTLPF.ITEMNO,
        AINDTLPF.SHTCNT, AINDTLPF.COLCNT
        
        FROM
        DTABUD.AINDTLPF AINDTLPF, DTABUD.OITEMSPF OITEMSPF

        WHERE
        AINDTLPF.INVYR = """+str(self.invoiceYear)+"""
        AND AINDTLPF.INVNO = """+str(self.invoice)+"""
        AND AINDTLPF.ITEMNO = OITEMSPF.ITEMNO
        AND AINDTLPF.DETPRT = 'Y'
        """
        conn = pyodbc.connect('REDACTED',
                              uid=self.as400_username,
                              password=self.as400_password)
        itransport = DatabaseTransport(conn)
        cursor = conn.cursor()
        cursor = cursor.execute(itemsQuery)
        itemsList = []
        optionsList = []
        for row in cursor:
            final = [elem for elem in row]
            itemsList.append(final)
        for j in range(0,len(itemsList)):
            for i in range(0,len(itemsList[j])):
                if isinstance(itemsList[j][i],str):
                    while True:
                        if itemsList[j][i].endswith(' '):
                            itemsList[j][i] = itemsList[j][i][:-1]
                        elif itemsList[j][i] == '':
                            break
                        elif itemsList[j][i].startswith(' '):
                            itemsList[j][i] = itemsList[j][i][2:]
                        else:
                            break
        for item in itemsList:
            optionsList.append(getOptions(item[0]))

        return itemsList,optionsList
                
    def buildInvoice(self,info):
        #this function starts it off with a global, yeah for sure could have coded this better
        #note to self, is that not the point of making this a class? smh
        #sets up our pdf that we are building
        global pdf
        pdf = FPDF()
        #this is how packslip/shipping information is generated
        #after the invoice portion is done, this will fire
        def genShippingPages():
            def getShipMethod(packslip):
                smQuery = f"""
                SELECT
                SHPMTHPF.CODDESC, SHPTRKPF.CARTRK
                
                FROM
                DTABUD.SHPMTHPF SHPMTHPF,
                DTABUD.OPSHDRPF OPSHDRPF,
                DTABUD.SHPTRKPF SHPTRKPF

                WHERE
                OPSHDRPF.PCKSLP = {int(packslip)}
                AND SHPMTHPF.CARRIER = OPSHDRPF.CARRIER
                AND SHPMTHPF.SERVICE = OPSHDRPF.SERVICE
                AND SHPTRKPF.PSLPNBR = {int(packslip)}

                """
                conn = pyodbc.connect('REDACTED',
                                      uid=self.as400_username,
                                      password=self.as400_password)
                cursor = conn.cursor()
                cursor = cursor.execute(smQuery)
                for row in cursor:
                    final = [str(elem) for elem in row]
                while final[0].endswith(' '):
                    final[0] = final[0][:-1]
                while final[1].endswith(' '):
                    final[1] = final[1][:-1]
                return final[0],final[1]
            def shipDefaults():
                #gen new page element but for shipping exclusively
                global y
                if y >= 270:
                    y = 95
                    pdf.add_page()
                    pdf.set_margins(top=0,left=0,right=0)
                    pdf.set_auto_page_break(0,0)
                    pdf.set_font('Courier','B', 13)
                    pdf.set_xy(0,15)
                    pdf.multi_cell(0,5,'Shipping & Tracking'+pb+pb+'Invoice: '+info['invoice']+' - PO: '+info['po'],align='C')
                    pdf.set_font('Courier','B', 10)
                    pdf.set_text_color(0,0,0)
                    pdf.set_fill_color(236, 236, 236)
                    #title of shipping page plus both companies addresses
                    pdf.set_xy(15,35)
                    pdf.cell(180,0,border=1)
                    pdf.set_xy(17.5,40)
                    pdf.multi_cell(85,5,mmmaddress.replace('Remit to','Remit'), border=1, align='L',fill=False)
                    pdf.set_xy(107.5,40)
                    pdf.multi_cell(85,5,info['billTo'], border=1, align='L',fill=False)
                    #titles for packslip info
                    pdf.set_xy(15,90)
                    pdf.cell(25,5,'Packslip', border=1, align='C',fill=True)
                    pdf.set_xy(40,90)
                    pdf.cell(25,5,'Ship Date', border=1, align='C',fill=True)
                    pdf.set_xy(65,90)
                    pdf.cell(17,5,'Box Ct.', border=1, align='C',fill=True)
                    pdf.set_xy(82,90)
                    pdf.cell(113,5,'Location, Tracking & Method', border=1, align='C',fill=True)
                    pdf.set_xy(0,285)
                    pdf.cell(0,5,'Refer to carriers website for additional track numbers',align='C')

                    
            global y
            packslipList = (self.shippingInfo())[2]
            y = 999
            shipDefaults()
            #for every packslip get its ship method
            for record in packslipList:
                record.append(getShipMethod(record[0])[0])
                record.append(getShipMethod(record[0])[1])
                shipDefaults()
                pdf.set_xy(15,y)
                pdf.cell(25,10,str(record[0]), border=1, align='C',fill=False)
                pdf.set_xy(40,y)
                pdf.cell(25,10,record[8].strftime("%m/%d/%y"), border=1, align='C',fill=False)
                pdf.set_xy(65,y)
                pdf.cell(17,10,str(record[9]), border=1, align='C',fill=False)
                pdf.set_xy(82,y)
                pdf.multi_cell(113,5,record[1]+' '+record[5]+', '+record[6]+pb+'Tracking: '+record[11]+' - '+record[10], border=1, align='C',fill=False)
                y += 10 
        
        def amounts(sub,ship,tax,discount,paid,credit,total):
            #gets each of the different final amount values
            totals = [ship,tax,discount,paid,credit]
            final = f'{sub:,.2f}'
            finalLabel = 'Sub Total:'
            finalLabelList = ['Shipping & Handling:',
                              'Sales Tax:',
                              'Discount:',
                              'Payment Received:',
                              'Credit Applied:',
                              'Total Amount Due:']
            for item in totals:
                final += pb + f'{item:,.2f}'
            for item in finalLabelList:
                finalLabel += pb + item
            pdf.set_xy(120,240)
            pdf.multi_cell(80,5,final, border=1, align='R')
            pdf.set_xy(120,240)
            pdf.multi_cell(50,5,finalLabel, border=0, align='R')
            pdf.set_xy(120,270)
            pdf.cell(80,5,f'{total:,.2f}'+' '+currency, border=1, align='R')
        def genNewPage(info):
            #this function is ran several times as a check to the program
            #if the y value gets too long, it will fire and create a fresh page
            global y,firstPage,currency,ourName,email,ourPhone,mmmaddress
            if y >= 230:
                try:
                    if not firstPage:
                        firstPage = False
                    else:
                        amounts(subTotal,0,0,0,0,0,0)
                except Exception as e:
                    firstPage = True
                    #print(e)
                y = 105
                #each of the primary business units that would use this have unique info
                #such as contact, currency and remit to
                if info['bud'] == 1:
                    mmmaddress = REDACTED'
                    logo = r'REDACTED.png'
                    logoW = 90
                    logoH = 9
                    link = 'REDACTED'
                    ourPhone = 'REDACTED'
                    email = 'REDACTED'
                    ourName = 'REDACTED'
                    currency = 'USD'
                elif info['bud'] == 5:
                    mmmaddress = REDACTED'
                    logo = r'REDACTED.png'
                    logoW = 90
                    logoH = 9
                    link = 'REDACTED'
                    ourPhone = 'REDACTED'
                    email = 'REDACTED'
                    ourName = 'REDACTED'
                    currency = 'USD'
                elif info['bud'] == 9:
                    mmmaddress = REDACTED'
                    logo = r'REDACTED.png'
                    logoW = 90
                    logoH = 9
                    link = 'REDACTED'
                    ourPhone = 'REDACTED'
                    email = 'REDACTED'
                    ourName = 'REDACTED'
                    currency = 'CAD'
                elif info['bud'] == 10:
                    mmmaddress = REDACTED'
                    logo = r'REDACTED.png'
                    logoW = 90
                    logoH = 9
                    link = 'REDACTED'
                    ourPhone = 'REDACTED'
                    email = 'REDACTED'
                    ourName = 'REDACTED'
                    currency = 'USD'
                #all the code to setup our new page is below, lots of cells and positions
                pdf.add_page()
                pdf.set_margins(top=0,left=0,right=0)
                pdf.set_auto_page_break(0,0)
                pdf.set_font('Courier','B', 13)
                pdf.set_xy(0,15)
                pdf.multi_cell(0,5,'Invoice'+'\n('+currency+')',align='C')
                pdf.set_font('Courier','B', 10)
                pdf.set_text_color(0,0,0)
                pdf.set_fill_color(236, 236, 236)

                pdf.set_xy(145,12.5)
                prepays = ['F','P','W','B']
                if info['terms'] in prepays:
                    dueDate = info['postDate']
                else:
                    dueDate = (datetime.strptime(info['postDate'], '%m/%d/%y') + timedelta(days=30)).strftime("%m/%d/%y")
                invoiceInfoTitles = 'Invoice:'+pb+'PO Ref:'+pb+'Order:'+pb+'Post Date:'+pb+'Due Date:'+pb+'Total Due:'
                invoiceInfo = info['invoice']+pb+info['po']+pb+info['control']+pb+info['postDate']+pb+dueDate+pb+pb
                pdf.set_xy(145,12.5)
                pdf.multi_cell(23.5,5,invoiceInfoTitles, border=0, align='R')
                pdf.set_xy(145,12.5)
                pdf.multi_cell(55,5,invoiceInfo, border=1, align='R')

                pdf.set_xy(10,12.5)
                pdf.multi_cell(0,5,mmmaddress,align='L')
                
                pdf.set_xy(10,45)
                pdf.cell(190,5,'',border=1,align='C',fill=True)
                pdf.set_xy(10,45)
                pdf.cell(25,5,'Acct #',border=1,align='L',fill=True)
                pdf.set_xy(35,45)
                pdf.cell(35,5,'Purchase Order',border=1,align='L',fill=True)
                pdf.set_xy(70,45)
                pdf.cell(30,5,'Receive Date',border=1,align='L',fill=True)
                pdf.set_xy(100,45)
                pdf.cell(30,5,'Invoice Date',border=1,align='L',fill=True)
                pdf.set_xy(130,45)
                pdf.cell(35,5,'CSR Contact',border=1,align='L',fill=True)
                pdf.set_xy(165,45)
                pdf.cell(35,5,'AR Contact',border=1,align='L',fill=True)

                pdf.set_xy(10,50)
                pdf.cell(190,5,'',border=1,align='C')
                pdf.set_xy(10,50)
                pdf.cell(25,5,info['acctNum'],border=1,align='L')
                pdf.set_xy(35,50)
                pdf.cell(35,5,info['po'],border=1,align='L')
                pdf.set_xy(70,50)
                pdf.cell(30,5,info['ordDate'],border=1,align='L')
                pdf.set_xy(100,50)
                pdf.cell(30,5,info['postDate'],border=1,align='L')
                pdf.set_xy(130,50)
                pdf.cell(35,5,info['csrphone'],border=1,align='L')
                pdf.set_xy(165,50)
                pdf.cell(35,5,ourPhone,border=1,align='L')

                pdf.set_xy(17.5,60)
                pdf.multi_cell(85,5,info['billTo'], border=1, align='L',fill=True)
                pdf.set_xy(107.5,60)
                drops,shipTo,blank = self.shippingInfo()
                if shipTo != None:
                    shipToAddress = 'Ship To:'+pb+shipTo[1]+pb+shipTo[2]+shipTo[3]+shipTo[4]+pb+shipTo[5]+', '+shipTo[6]+' '+shipTo[7]+pb+pb+drops
                else:
                    shipToAddress = 'Ship To:'+pb+'~Missing Ship Info'+pb+pb+drops
                pdf.multi_cell(85,5,shipToAddress, border=1, align='L',fill=True)
                
                pdf.set_xy(10,100)
                pdf.cell(110,5,'Item Description', border=1, align='C',fill=True)
                pdf.set_xy(10,105)
                pdf.cell(110,135,'', border=1, align='C')
                pdf.set_xy(120,100)
                pdf.cell(25,5,'Bill Qty', border=1, align='C',fill=True)
                pdf.set_xy(120,105)
                pdf.cell(25,135,'', border=1, align='C')
                pdf.set_xy(145,100)
                pdf.cell(30,5,'Unit Price', border=1, align='C',fill=True)
                pdf.set_xy(145,105)
                pdf.cell(30,135,'', border=1, align='C')
                pdf.set_xy(175,100)
                pdf.cell(25,5,'Total', border=1, align='C',fill=True)
                pdf.set_xy(175,105)
                pdf.cell(25,135,'', border=1, align='C')
                
                pdf.set_xy(10,240)
                pdf.cell(110,35,'', border=1, align='L')
                pdf.set_xy(120,240)
                pdf.cell(80,35,'', border=1, align='C')
        #some setup stuff before getting our items and options from a
        #function above
        global subTotal,eproofTotal,y
        eproofTotal = 0
        subTotal = 0
        y = 240
        items,options = self.invoiceItems()
        
        for i in range(0,len(items)):
            #for every item in the items list format, place, calculate etc.
            genNewPage(info)
            #item description
            desc = items[i][1]
            if items[i][6] != 0:
                desc += ' '+str(items[i][6])+' SHEET'
            if items[i][7] != 0:
                desc += ' '+str(items[i][7])+' COLOR'
            pdf.set_xy(10,y)
            pdf.set_font('Courier','BU', 10)
            pdf.cell(110,5,desc, border=0, align='L',fill=False)
            pdf.set_font('Courier','B', 10)
            #item quantity
            pdf.set_xy(120,y)
            pdf.cell(25,5,str(items[i][3]), border=0, align='R',fill=False)
            #price per
            pdf.set_xy(145,y)
            if items[i][3] != 0:
                pdf.cell(30,5,f'{(items[i][4]/items[i][3]):,.2f}', border=0, align='R',fill=False)
            else:
                pdf.cell(30,5,'0.00', border=0, align='R',fill=False)
            #total price
            pdf.set_xy(175,y)
            pdf.cell(25,5,f'{items[i][4]:,.2f}', border=0, align='R',fill=False)
            subTotal += items[i][4]
            genNewPage(info)
            y += 5
            #the reason for seperating eproofs is they will not have any export charges
            #not a physical product after all
            if items[i][5] != 'PFBW':
                #item mainline
                pdf.set_xy(10,y)
                pdf.cell(110,5,items[i][2], border=0, align='L',fill=False)
                y += 7.5
            else:
                eproofTotal += float(items[i][4])
                y += 2.5
            for option in options[i]:
                genNewPage(info)
                #option description
                pdf.set_xy(10,y)
                pdf.cell(110,5,'    '+option[0], border=0, align='L',fill=False)
                #option quantity
                pdf.set_xy(120,y)
                pdf.cell(25,5,str(int(option[1])), border=0, align='R',fill=False)
                #price per
                pdf.set_xy(145,y)
                if option[1] != 0:
                    pdf.cell(30,5,f'{(option[2]/option[1]):,.2f}', border=0, align='R',fill=False)
                else:
                    pdf.cell(30,5,'0.00', border=0, align='R',fill=False)
                #total price
                pdf.set_xy(175,y)
                pdf.cell(25,5,f'{option[2]:,.2f}', border=0, align='R',fill=False)
                subTotal += option[2]
                y += 7.5
            #all things export, have to be formatted this way for regulations I believe
            export = self.exportCharges(items[i][0])
            if export != None:
                if y+27.5 >= 230:
                    genNewPage(info)
                duty = round(export[0]*export[1]*decimal.Decimal(.01),2)
                line = round(export[0]*export[2]*decimal.Decimal(.01),2)
                pdf.set_xy(10,y)
                pdf.set_font('Courier','BU', 10)
                pdf.cell(110,5,'Export Charges by Product Line', border=0, align='L',fill=False)
                y += 5
                pdf.set_xy(10,y)
                pdf.set_font('Courier','B', 10)
                pdf.cell(110,5,'Product Line Tax = Product('+str(export[0])+') x GST('+str(export[2])+'%)', border=0, align='L',fill=False)
                pdf.set_xy(175,y)
                pdf.cell(25,5,str(round(line,2)), border=0, align='R',fill=False)
                y += 5
                pdf.set_xy(10,y)
                pdf.cell(110,5,'Duty Tax = Product('+str(export[0])+') x ('+str(export[1])+'%)', border=0, align='L',fill=False)
                pdf.set_xy(175,y)
                pdf.cell(25,5,str(round(duty,2)), border=0, align='R',fill=False)
                subTotal += duty + line
                
                y += 5
                pdf.set_xy(10,y)
                pdf.cell(110,5,'GST reported under REDACTED',border=0,align='L',fill=False)
                y += 7.5
        #get credit amounts and then run the amounts function
        #now returns more infomation to customer on how payment was made
        #and also what credits were applied
        payment,credit,paymentValues,creditValues = self.creditBalance()
        payNotes = ''
        credNotes = ''
        pFirst = True
        cFirst = True
        for paid in paymentValues:
            if pFirst:
                payNotes += '**Payment was made via '+paid[3]
                pFirst = False
            else:
                payNotes += pb+'**Payment was made via '+paid[3]
        for cred in creditValues:
            if cFirst:
                credNotes += '*Credit #'+cred+' was applied'
                cFirst = False
            else:
                credNotes += pb+'*Credit #'+cred+' was applied'
        finalNotes = ''
        if payNotes != '' and credNotes != '':
            finalNotes = payNotes+pb+credNotes
            pdf.set_xy(120,260)
            pdf.cell(5,5,'   **')
            pdf.set_xy(120,265)
            pdf.cell(5,5,'      *')
        elif payNotes != '':
            finalNotes = payNotes
            pdf.set_xy(120,260)
            pdf.cell(5,5,'**')
        elif credNotes != '':
            finalNotes = credNotes
            pdf.set_xy(120,260)
            pdf.cell(5,5,'*')
        #places our notes
        pdf.set_xy(10,240)
        pdf.multi_cell(110,5,finalNotes,border=0,align='L')
        total = float(subTotal)+info['ship']+info['tax']-info['discount']-payment-credit
        amounts(subTotal,
                info['ship'],
                info['tax'],
                (-1*info['discount']),
                (-1*payment),
                (-1*credit),
                total
                )
        #iterate over each page and add a little trademark based on business unit
        for i in range(0,len(pdf.pages)):
            pdf.page = i + 1
            pdf.set_xy(145,37.5)
            pdf.cell(55,5,f'{total:,.2f}'+' '+currency, border=0, align='R')
            if info['bud'] == 1 or info['bud'] == 9:
                pdf.set_xy(0,287.5)
                pdf.cell(0,5,"REDACTED",align='C')
        #add some more basic things    
        if len(pdf.pages) > 1:
            for i in range(0,len(pdf.pages)-1):
                pdf.page = i + 1
                pdf.set_xy(0,280)
                pdf.cell(0,5,'CONTINUED ON NEXT PAGE',align='C')
            pdf.set_xy(0,280)
            pdf.page = len(pdf.pages)
            pdf.cell(0,5,'TO REORDER, PLEASE REFERENCE THE PURCHASE ORDER '+info['po']+' DATED '+info['postDate'],align='C')
        else:
            pdf.set_xy(0,280)
            pdf.cell(0,5,'TO REORDER, PLEASE REFERENCE THE PURCHASE ORDER '+info['po']+' DATED '+info['postDate'],align='C')
        #time to call the shipping portion of the invoice 
        genShippingPages()
        #finishing touches and viola, its saved and started on the users computer
        #it is also copied to their clipboard to be pasteable into an email response
        #if they do not want to send it from Burge
        pdf.page = len(pdf.pages)
        pdf.set_title(str(self.invoice)+' - '+str(self.invoiceYear))
        pdf.set_author(ourName+' - User: '+self.as400_username)
        pdf.output(r'REDACTED\burge\Generated Invoices/'+str(self.invoice)+' - '+str(self.invoiceYear)+'.pdf')
        os.startfile(r'REDACTED\burge\Generated Invoices/'+str(self.invoice)+' - '+str(self.invoiceYear)+'.pdf')
        output = r"'REDACTED\burge\Generated Invoices/"+str(self.invoice)+" - "+str(self.invoiceYear)+".pdf'"
        os.system('powershell Set-Clipboard -LiteralPath '+output)
            
    def sendEmail(self,requester):
        #function to send an email based on users input
        #a lot of static and dynamic info based on invoice/customer
        today = date.today()
        today = str(today.strftime("%m/%d/%y"))
        info = self.header()
        server = smtplib.SMTP(REDACTED)
        msg = MIMEMultipart()
        msg['From'] = email
        msg['To'] = requester

        msg['Subject'] = 'Invoice #'+info['invoice']+' for PO '+info['po']
        body1 = 'Acct: '+info['acctNum']+' - '+info['acctName']+pb+'PO: '+info['po']+pb+pb
        body2 = 'See the attached document for your invoice with '+ourName+'.'+pb
        body5 = 'If you have any questions or concerns, please let us know.'+pb+pb+'Kind regards,'+pb+pb
        body6 = 'Accounts Receivable Department'+pb+ourName+pb+'AR Direct: '+ourPhone+pb
        body7 = 'REDACTED'
        body = body1+body2+body5+body6+body7
        msg.attach(MIMEText(body, 'plain'))
        
        filename = str(self.invoice)+' - '+str(self.invoiceYear)+'.pdf'
        path = r'M:\burge\Generated Invoices/'
        
        with open(path+filename, 'rb') as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
            
        encoders.encode_base64(part)
        part.add_header("Content-Disposition",f"attachment; filename = {filename}")
        msg.attach(part)

        server.send_message(msg)
        del msg
        print('Invoice #'+str(self.invoice)+' - '+str(self.invoiceYear)+ ' sent to '+requester)
        server.quit()

#invClass = invoiceGen(str(51802),str(2024),'REDACTED','REDACTED')
#invClass.buildInvoice(invClass.header())


class CMClass(FPDF):
    def db_queries(self,credit,creditYear,as400_username,as400_password):
        conn = pyodbc.connect('REDACTED',
                              uid=as400_password,
                              password=as400_username)
        itransport = DatabaseTransport(conn)
        cursor = conn.cursor()
        creditSQLQuery = '''
        SELECT
            ACDHDRPF.TRANN, ACDHDRPF.TRNYR, ACDHDRPF.INVNO,
            ACDHDRPF.INVYR, AINHDRPF.PONUM, ACTFILPF.NAME,
            ACTFILPF.POBOX, ACTFILPF.STR1, ACTFILPF.STR2,
            ACTFILPF.CITY, ACTFILPF.SORP, ACTFILPF.ZIP,
            ACTFILPF.ACCTNO, ACDHDRPF.CRTDAT, ACDHDRPF.APVDAT,
            ACDHDRPF.TOTCRD, ACDHDRPF.TTBUD
            
        FROM
            DTABUD.ACDHDRPF ACDHDRPF
            LEFT JOIN DTABUD.AINHDRPF AINHDRPF ON
                AINHDRPF.INVNO = ACDHDRPF.INVNO
                AND AINHDRPF.INVYR = ACDHDRPF.INVYR
            LEFT JOIN DTABUD.ACTFILPF ACTFILPF ON
                ACTFILPF.ACCTNO = ACDHDRPF.ACCTNO
        
        WHERE
            ACDHDRPF.TRNYR = '''+creditYear+'''
            AND ACDHDRPF.TRANN = '''+credit+'''
        '''
        cursor = cursor.execute(creditSQLQuery)
        for row in cursor:
            headerInfo = [elem for elem in row]
        for i in range(0,len(headerInfo)):
            if isinstance(headerInfo[i],str):
                while headerInfo[i].endswith(' '):
                    headerInfo[i] = headerInfo[i][:-1]
        if headerInfo[6] != '':
            headerInfo[6] = 'PO Box: '+headerInfo[6]
        global acctNum,acctName
        acctNum = str(headerInfo[12])
        acctName = headerInfo[5]
        pobox = headerInfo[6]
        street1 = headerInfo[7]
        street2 = headerInfo[8]
        if street2 != '':
            street = street1 + pb + street2
        else:
            street = street1
        if pobox != '':
            street = pobox + pb + street
        city = headerInfo[9]
        state = headerInfo[10]
        postal = headerInfo[11]
        csz = city + ', ' + state + ' ' + postal
        address = 'To:'+pb+'Acct #'+acctNum+pb+acctName+pb+street+pb+csz
        info = [headerInfo[0],headerInfo[1],headerInfo[2],
                headerInfo[3],str(headerInfo[4]),address,
                headerInfo[13],headerInfo[14],headerInfo[15],
                headerInfo[16]
                ]

        return info
    
    def buildCM(self,credit,creditYear,as400_username,as400_password):
        info = self.db_queries(credit,creditYear,as400_username,as400_password)
        global email,ourName,ourPhone
        if info[9] == 1:
            mmmaddress = 'REDACTED'
            logo = r'REDACTED.png'
            logoW = 90
            logoH = 9
            link = 'REDACTED'
            ourPhone = 'REDACTED'
            email = 'REDACTED'
            ourName = 'REDACTED'
            currency = 'USD'
        elif info[9] == 5:
            mmmaddress = 'REDACTED'
            logo = r'REDACTED.png'
            logoW = 90
            logoH = 9
            link = 'REDACTED'
            ourPhone = 'REDACTED'
            email = 'REDACTED'
            ourName = 'REDACTED'
            currency = 'USD'
        elif info[9] == 9:
            mmmaddress = 'REDACTED'
            logo = r'REDACTED.png'
            logoW = 90
            logoH = 9
            link = 'REDACTED'
            ourPhone = 'REDACTED'
            email = 'REDACTED'
            ourName = 'REDACTED'
            currency = 'CAD'
        elif info[9] == 10:
            mmmaddress = 'REDACTED'
            logo = r'REDACTED.png'
            logoW = 90
            logoH = 9
            link = 'REDACTED'
            ourPhone = 'REDACTED'
            email = 'REDACTED'
            ourName = 'REDACTED'
            currency = 'USD'
        today = datetime.today()
        today = today.strftime("%m/%d/%y")
        self.add_page()
        self.set_margins(top=0,left=0,right=0)
        self.set_auto_page_break(0,0)
        self.set_font('Courier','B', 13)
        self.set_xy(20,12.5)
        self.cell(0,5,'Date: '+today)
        self.set_xy(20,25)
        self.multi_cell(0,5,mmmaddress)
        self.set_xy(120,25)
        self.multi_cell(0,5,info[5])
        self.set_xy(0,75)
        self.cell(0,5,'*** C R E D I T  M E M O ***',align='C')
        self.set_xy(0,90)
        self.multi_cell(50,5,'Credit #: '+pb+'Ref. Invoice: ',align='R')
        self.set_xy(47.5,90)
        self.multi_cell(50,5,str(info[0])+' '+str(info[1])+pb+str(info[2])+' '+str(info[3]),align='L')
        self.set_xy(0,95)
        self.cell(200,5,'PO: '+info[4],align='C')
        self.set_xy(120,90)
        self.multi_cell(80,5,f'Date Created: {info[6].strftime("%m/%d/%y")}'+pb+f'Date Approved: {info[7].strftime("%m/%d/%y")}',align='R') 
        textBody = f"""
Greetings{pb}
We have credited the following amount to your account: {pb+pb+'$'+f'{info[8]:,.2f}'+' '+currency+pb+pb}
For more information contact Accounts Receivables at {ourPhone}
        """
        self.set_xy(0,115)
        self.multi_cell(0,5,textBody,align='C')
        self.set_title(str(credit)+' - '+str(creditYear))
        self.set_author(ourName+' - User: '+as400_username)
        self.output(r'M:\Burge\Generated Credits/'+str(credit)+' - '+creditYear+'.pdf')
        os.startfile(r'M:\Burge\Generated Credits/'+str(credit)+' - '+creditYear+'.pdf')
        output = r"'M:\burge\Generated Credits/"+str(credit)+" - "+str(creditYear)+".pdf'"
        os.system('powershell Set-Clipboard -LiteralPath '+output)

    def sendEmail(self,requester,credit,creditYear,as400_username,as400_password):########
        #function to send an email based on users input
        #a lot of static and dynamic info based on credit/customer
        today = datetime.today()
        today = str(today.strftime("%m/%d/%y"))
        info = self.db_queries(credit,creditYear,as400_username,as400_password)
        server = smtplib.SMTP(host='smtp.taylorcorp.com',port=25)
        msg = MIMEMultipart()
        msg['From'] = email
        msg['To'] = requester

        msg['Subject'] = 'Credit #'+str(info[0])+' from '+ourName
        body1 = 'Acct: '+acctNum+' - '+acctName+pb+pb
        body2 = 'See the attached document for your credit memo with '+ourName+'.'+pb
        body5 = 'If you have any questions or concerns, please let us know.'+pb+pb+'Kind regards,'+pb+pb
        body6 = 'Accounts Receivable Department'+pb+ourName+pb+'AR Direct: '+ourPhone+pb
        body7 = 'REDACTED'
        body = body1+body2+body5+body6+body7
        msg.attach(MIMEText(body, 'plain'))
        
        filename = str(credit)+' - '+str(creditYear)+'.pdf'
        path = r'REDACTED\burge\Generated Credits/'
        
        with open(path+filename, 'rb') as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
            
        encoders.encode_base64(part)
        part.add_header("Content-Disposition",f"attachment; filename = {filename}")
        msg.attach(part)

        server.send_message(msg)
        del msg
        print('Credit #'+str(credit)+' - '+str(creditYear)+ ' sent to '+requester)
        server.quit()

#pdf = CMClass()
#pdf.buildCM(str(826),str(2024),'REDACTED','REDACTED')
#pdf.sendEmail('REDACTED',str(826),str(2024),'REDACTED','REDACTED')

