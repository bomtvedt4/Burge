import customtkinter as ctk
from misc_commands import get_time, as400_quick_log,exit_command_wrapper
import accounting
import billing
import quickframeprompts
from quickframecommands import *
#imports
#some default text colors that a user can select for their text
text_colors = ['White','Black','Red','Blue','Green',
               'Yellow','Orange','Purple','Pink','Gray','Violet',
               'Maroon','Deep Sky Blue']
#a few setup variables
status = 'Activate'
hidden = True
#class for the frame found on the left of the main menu, these are functions/programs available to any user
class quick_frame:
    def __init__(self,quickFrame,default_text,as400_username,as400_password,menu_root,options,login,mainMenu):
        self.quickFrame = quickFrame
        self.default_text = default_text
        self.menu_root = menu_root
        self.as400_username = as400_username
        self.as400_password = as400_password
        self.options = options
        self.login = login
        self.mainMenu = mainMenu
    #one function button that opens all saved apps and another that creates a prompt to edit them
    def quickApps(self):
        quickAppsLabel = ctk.CTkLabel(master=self.quickFrame,
                     text='Quick Apps',
                     font=('Calibri',20,'bold'),
                     text_color=self.default_text
                     ).place(relx=.5,rely=.05,anchor='center')
        launchAllApps = ctk.CTkButton(master=self.quickFrame,
                      text='Launch All',
                      width=90,
                      height=35,
                      text_color=self.default_text,
                      command=lambda: quickAppsSetup.launch_app_all(self.as400_username)
                      ).place(relx=.33,rely=.1,anchor='center')
        editApps = ctk.CTkButton(master=self.quickFrame,
                      text='Edit Apps',
                      width=90,
                      height=35,
                      text_color=self.default_text,
                      command=lambda:quickframeprompts.quickAppsPrompt(self.quickFrame,
                                                      self.default_text,
                                                      self.as400_username)
                      ).place(relx=.67,rely=.1,anchor='center')
    #similar to quickapps, however just keeping a seperate list
    def quickLinks(self):
        quicklinksLabel = ctk.CTkLabel(master=self.quickFrame,
                     text='Quick Links',
                     font=('Calibri',20,'bold'),
                     text_color=self.default_text
                     ).place(relx=.5,rely=.15,anchor='center')
        launchAllLinks = ctk.CTkButton(master=self.quickFrame,
                      text='Launch All',
                      width=90,
                      height=35,
                      text_color=self.default_text,
                      command=lambda: quickLinksSetup.launch_link_all(self.as400_username)
                      ).place(relx=.33,rely=.2,anchor='center')
        editLinks = ctk.CTkButton(master=self.quickFrame,
                      text='Edit Links',
                      width=90,
                      height=35,
                      text_color=self.default_text,
                      command=lambda:quickframeprompts.quickLinksPrompt(self.quickFrame,
                                                      self.default_text,
                                                      self.as400_username)
                      ).place(relx=.67,rely=.2,anchor='center')
    #two different buttons that lead to either a credit or invoice prompt, these can be sent out from Burge
    def invoice_credit(self):
        invoiceCreditLabel = ctk.CTkLabel(master=self.quickFrame,
                     text='Send Invoice/Credit',
                     font=('Calibri',20,'bold'),
                     text_color=self.default_text
                     ).place(relx=.5,rely=.25,anchor='center')
        invoicePromptButton = ctk.CTkButton(master=self.quickFrame,
                      text='Send Invoice',
                      width=90,
                      height=35,
                      text_color=self.default_text,
                      command=lambda:quickframeprompts.invoicePrompt(self.quickFrame,
                                                                     self.default_text,
                                                                     self.as400_username,
                                                                     self.as400_password)
                      ).place(relx=.33,rely=.3,anchor='center')
        
        creditPromptButton = ctk.CTkButton(master=self.quickFrame,
                      text='Send Credit',
                      width=90,
                      height=35,
                      text_color=self.default_text,
                      command=lambda:quickframeprompts.creditPrompt(self.quickFrame,
                                                                     self.default_text,
                                                                     self.as400_username,
                                                                     self.as400_password)
                      )
        creditPromptButton.place(relx=.67,rely=.3,anchor='center')
        validCreditUsers = ['arburns','jlsargen','tljohnso','bomtved','sallen2']
        nameVar = (self.as400_username).lower()
        if nameVar not in validCreditUsers:
            creditPromptButton.configure(state='disabled')
            
        
    #user can change text color or produce a random color
    def modeChoices(self):
        modeLabel = ctk.CTkLabel(master=self.quickFrame,
                             text='Appearance/Text Modes',
                             font=('Calibri',20,'bold'),
                             text_color=self.default_text
                             ).place(relx=.5,rely=.85,anchor='center')
        mode = ctk.CTkOptionMenu(master=self.quickFrame,
                             values=['Dark','Light'],
                             width=90,
                             height=35,
                             text_color=self.default_text,
                             variable=ctk.StringVar(value='Mode'),
                             command=setup.set_mode_wrapper(self.options,self.as400_username,self.menu_root,self.mainMenu)
                             ).place(relx=.33,rely=.9,anchor='center')
        textMode = ctk.CTkOptionMenu(master=self.quickFrame,
                                 values=text_colors,
                                 width=90,
                                 height=35,
                                 text_color=self.default_text,
                                 variable=ctk.StringVar(value='Colors'),
                                 command=setup.set_text_color_wrapper(self.options,self.as400_username,self.menu_root,self.mainMenu)
                                 ).place(relx=.66,rely=.9,anchor='center')
        randomColor = ctk.CTkButton(master=self.quickFrame,
                                    text='Random?',
                                    width=90,
                                    height=35,
                                    text_color=self.default_text,
                                    command=lambda:setup.set_random_color(self.options,self.as400_username,self.menu_root,self.mainMenu)
                                    ).place(relx=.5,rely=.96,anchor='center')

#class for the frame that has exit options, a timer and a quick log option
class user_frame:
    def __init__(self,userFrame,as400_username,as400_password,default_text,menu_root,login):
        self.userFrame = userFrame
        self.as400_username = as400_username
        self.as400_password = as400_password
        self.default_text = default_text
        self.menu_root = menu_root
        self.login = login
    #every one second we are calling the current time and updating the label
    def time_label(self):
        def reset(time):
            time.destroy()
            self.time_label()
        text = get_time()

        time = ctk.CTkLabel(master=self.userFrame,
                            text=text,
                            font=('Calibri',15),
                            text_color=self.default_text
                            )
        time.place(relx=.5,rely=.1,anchor='center')
        self.userFrame.after(1000,lambda: reset(time))
    #basic label showing logged in user
    def user_label(self):
        user = ctk.CTkLabel(master=self.userFrame,
                            text='User: '+self.as400_username.upper(),
                            font=('Calibri',15,'underline'),
                            text_color=self.default_text,
                            ).place(relx=.5,rely=.2,anchor='center')
    #button that uses user login information on each open as400 session
    def quick_log(self):
        def changeStatus(quickLogButton):
            global status
            status = 'Activated'
            quickLogButton.configure(text=status)
            quickLogButton.configure(state='disabled')
        quickLogLabel = ctk.CTkLabel(master=self.userFrame,
                                     text='AS400 Quick Log',
                                     font=('Calibri',15,'bold'),
                                     text_color=self.default_text
                                     ).place(relx=.5,rely=.375,anchor='center')
        quickLogButton = ctk.CTkButton(master=self.userFrame,
                                       text=status,
                                       width=75,
                                       height=30,
                                       text_color=self.default_text,
                                       command=lambda: [changeStatus(quickLogButton),
                                                        as400_quick_log(self.as400_username,self.as400_password)]
                                       )
        quickLogButton.place(relx=.5,rely=.5,anchor='center')

    #user can either log out or exit completely using this list
    def exit_options(self):
        exitLabel = ctk.CTkLabel(master=self.userFrame,
                             text='Exit Options',
                             font=('Calibri',15,'bold'),
                             text_color=self.default_text
                             ).place(relx=.5,rely=.65,anchor='center')
        exitOptions = ctk.CTkOptionMenu(master=self.userFrame,
                             values=['Options','Log Out','Exit'],
                             width=75,
                             height=30,
                             text_color=self.default_text,
                             command=exit_command_wrapper(self.menu_root,self.login)
                             ).place(relx=.5,rely=.775,anchor='center')

#class for a frame that is basically filling space, may show changelogs/help info and current version
class des_frame:
    def __init__(self,desFrame,as400_username,default_text,version):
        self.desFrame = desFrame
        self.as400_username = as400_username
        self.default_text = default_text
        self.version = version
    #function that alternates between showing text versus not
    def welcome(self):
        def clear():
            global hidden
            for item in self.desFrame.winfo_children():
                item.destroy()
            if hidden == True:
                hidden = False
            else:
                hidden = True
            self.welcome()
        body ="""
This is Project Alpha. A GUI program created to consolidate many of the programs that have been created to aid the 3M
billing department. There are also a few neat tools that have been added to streamline opening repeated apps/links and
sign into any open AS400 sessions (top right). Quick apps/links can be found on the left. You can add, remove or launch
apps and links using the 'Edit' buttons. Once setup, you can launch all your links/apps at once using their given buttons.
In the bottom left of the main menu, there are also some settings to change text color and visual mode."""

        versionLabel = ctk.CTkLabel(master=self.desFrame,
                                       text=self.version,
                                       text_color=self.default_text,
                                       font=('Calibri',35,'bold')
                                       )
        versionLabel.place(relx=.5,rely=.25,anchor='center')
        bodyLabel = ctk.CTkLabel(master=self.desFrame,
                                text=body,
                                font=('Calibri',15,'bold'),
                                text_color=self.default_text,
                                width=540
                                )
        bodyLabel.place(relx=.5,rely=.5,anchor='center')
        welcomeText = ctk.CTkLabel(master=self.desFrame,
                               text='Welcome '+self.as400_username.upper(),
                               font=('Calibri',20,'bold'),
                               text_color=self.default_text,
                               width=150,height=45
                               )
        welcomeText.place(relx=.5,rely=.1,anchor='center')
        textButton = ctk.CTkButton(master=self.desFrame,
                                   text='Help',
                                   width=75,
                                   height=30,
                                   text_color=self.default_text,
                                   command=clear
                                   )
        textButton.place(relx=.95,rely=.9,anchor='center')
        if hidden == True:
            bodyLabel.destroy()
            textButton.configure(text='Help')
        else:
            versionLabel.destroy()
            textButton.configure(text='Hide')

#class for the frame that contains user specific programs, these will be revealed/hidden based on user by user basis
class prog_frame:
    def __init__(self,progFrame,as400_username,as400_password,default_text):
        self.progFrame = progFrame
        self.as400_username = as400_username
        self.default_text = default_text
        self.as400_password = as400_password
    #kind of clunky code that goes through each list of valid programs and valid users to be able to use those programs
    def valid_programs(self):
        validAccountingUser = ['bomtved','jlsargen','arburns','tljohnso']
        accountingProg = ['Morning'+'\n'+'Reports','Prepay'+'\n'+'Estimate','AR-Issues']
        validBillingUser = ['bomtved','jlsargen','jmeyer','tljohnso','rwendin']
        billingProg = ['Pops'+'\n'+'Helper','Hold'+'\n'+'Helper','MissQuote'+'\n'+'Helper']
        x,y = 0.1,0.2
        validPrograms = []
        apps = ctk.CTkLabel(master=self.progFrame,
                            text='User Apps',
                            text_color=self.default_text,
                            font=('Calibri',25,'italic')
                            ).place(relx=.0125,rely=.0125)
        if self.as400_username.lower() in validAccountingUser:
            for program in accountingProg:
                if program == 'Morning'+'\n'+'Reports':
                    mrhSetup = accounting.Morning_Report_Helper(self.progFrame,self.as400_username,self.as400_password,self.default_text)
                    run = mrhSetup.mmddyy
                elif program == 'Prepay'+'\n'+'Estimate':
                    proformaSetup = accounting.prepayment_estimate(self.progFrame,self.as400_username,self.as400_password,self.default_text)
                    run = proformaSetup.proforma_window
                elif program == 'AR-Issues':
                    ar_issuesSetup = accounting.ar_issues(self.progFrame,self.as400_username,self.as400_password,self.default_text)
                    run = ar_issuesSetup.ar_issues
                ctk.CTkButton(master=self.progFrame,
                              text=program.replace('_',' '),
                              width=100,height=100,
                              text_color=self.default_text,
                              font=('Calibri',25,'italic'),
                              corner_radius=5,
                              border_color='#808080',
                              border_width=2,
                              command=run
                              ).place(relx=x,rely=y,anchor='center')
                if x + .15 > .6:
                    y = y + .25
                    x = .1
                else:
                    x = x + .15
                    
        if self.as400_username.lower() in validBillingUser:
            for program in billingProg:
                if program == 'Pops'+'\n'+'Helper':
                    popsSetup = billing.popsHelper(self.progFrame,self.default_text)
                    run = popsSetup.popsHelper
                elif program == 'Hold'+'\n'+'Helper':
                    holdSetup = billing.holdHelper(self.progFrame,self.default_text)
                    run = holdSetup.holdHelper
                elif program == 'MissQuote'+'\n'+'Helper':
                    run = lambda: billing.missquoteHelper(self.progFrame,self.default_text)
                ctk.CTkButton(master=self.progFrame,
                              text=program.replace('_',' '),
                              width=100,height=100,
                              text_color=self.default_text,
                              font=('Calibri',25,'italic'),
                              corner_radius=5,
                              border_color='#808080',
                              border_width=2,
                              command=run
                              ).place(relx=x,rely=y,anchor='center')
                if x + .15 > .6:
                    y = y + .25
                    x = .1
                else:
                    x = x + .15
                          
    
    
