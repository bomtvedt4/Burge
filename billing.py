import os
import customtkinter as ctk
from misc_commands import center_root,holdMacro
import pygetwindow as gw
#imports

popsStatus = 'PopsHelper is Inactive'
#class for popshelper which is a basic program that copies recent numbers typed to the clipboard
class popsHelper:
    def __init__(self,progFrame,default_text):
        self.progFrame = progFrame
        self.default_text = default_text
    def popsHelper(self):
        #due to programming limitations, the program is ran outside of burge and opened/closed accordingly that way
        def changePops():
            global popsStatus
            #if its open close it
            if popsStatus == 'PopsHelper is Active':
                os.system("taskkill /f /im  popshelper.exe")
                popsStatus = 'PopsHelper is Inactive'
        def submitEnter():
            #on pressing the button, check whether or not popshelper should be opened or closed
            global popsStatus
            if popsStatus == 'PopsHelper is Inactive':
                popsStatus = 'PopsHelper is Active'
                os.startfile(r'M:\burge/PopsHelper.exe')
                statusButton.configure(text='Activating...')
                statusButton.configure(state='disabled')
                prompt.after(5000,lambda:[statusButton.configure(text='Deactivate'),
                                       statusButton.configure(state='normal')])
            else:
                popsStatus = 'PopsHelper is Inactive'
                os.system("taskkill /f /im  popshelper.exe")
                statusButton.configure(text='Deactivating')
                statusButton.configure(state='disabled')
                prompt.after(2000,lambda:[statusButton.configure(text='Activate'),
                                          statusButton.configure(state='normal')])
        #prompt with a label and a button that starts/stops the program        
        prompt = ctk.CTkToplevel(master=self.progFrame)
        prompt.title('PopsHelper')
        prompt.wm_transient(self.progFrame)
        window = center_root(width=400,
                      height=300,
                      screenX=prompt.winfo_screenwidth(),
                      screenY=prompt.winfo_screenheight())
        prompt.geometry(window.default())
        label = ctk.CTkLabel(master=prompt,
                             text='PopsHelper',
                             text_color=self.default_text,
                             font=('Calibri',20)
                             ).pack(pady=12,padx=10)
        statusButton = ctk.CTkButton(master=prompt,
                                     text='Activate',
                                     text_color=self.default_text,
                                     command=submitEnter
                                     )
        statusButton.pack(pady=12,padx=10)
        #basically a close protocol if popshelper isn't manually closed
        prompt.protocol("WM_DELETE_WINDOW",
                        lambda:[prompt.destroy(),
                                changePops()])

#class for holdhelper which is a basic macro for billers
class holdHelper:
    def __init__(self,progFrame,default_text):
        self.progFrame = progFrame
        self.default_text = default_text

    def holdHelper(self):
        def submitEnter(event):
            #get AS400 session and type of hold before running a macro that places an order on hold
            session = sessionVar.get()
            holdType = holdTypeVar.get()
            holdMacro(session,holdType)
        #prompt with a few lists that give the user the choice of which session to run the macro on and what type of hold
        prompt = ctk.CTkToplevel(master=self.progFrame)
        prompt.title('Hold Helper')
        prompt.wm_transient(self.progFrame)
        window = center_root(width=400,
                      height=300,
                      screenX=prompt.winfo_screenwidth(),
                      screenY=prompt.winfo_screenheight())
        prompt.geometry(window.default())
        label = ctk.CTkLabel(master=prompt,
                             text='Select session and hold type',
                             text_color=self.default_text,
                             font=('Calibri',20)
                             ).place(relx=.5,rely=.1,anchor='center')
        sessions = gw.getAllTitles()
        total = len(sessions)
        step = 0
        while step != total:
            for item in sessions:
                if 'Session' not in item:
                    sessions.remove(item)
                    break
            step = step + 1
        for i in range(0,len(sessions)):
            sessions[i] = sessions[i][0:9]
        sessions.sort()
        sessionVar = ctk.StringVar(prompt,'Session A')
        holdTypeVar = ctk.StringVar(prompt,'General')
        session = ctk.CTkOptionMenu(master=prompt,
                                    variable=sessionVar,
                                    values=sessions,
                                    text_color=self.default_text,
                                    width=90,
                                    height=35,
                                    )
        session.place(relx=.33,rely=.25,anchor='center')
        holdType = ctk.CTkOptionMenu(master=prompt,
                                     variable=holdTypeVar,
                                     values=['General','CC','Ship $'],
                                     text_color=self.default_text,
                                     width=90,
                                     height=35,
                                     )
        holdType.place(relx=.67,rely=.25,anchor='center')
        holdButton = ctk.CTkButton(master=prompt,
                                   text='Hold Macro',
                                   text_color=self.default_text,
                                   command=lambda:submitEnter(event=None)
                                   )
        holdButton.place(relx=.5,rely=.4,anchor='center')
        prompt.bind('<Return>',submitEnter)

def missquoteHelper(progFrame,default_text):
    #designed to do some basic quick math for billers to handle missquotes
    def calculate(event):
        #really basic program just does the math and displays output
        finalB = float(newA.get())/float(originalA.get())*float(originalB.get())
        output.configure(text=finalB)
    #basically the prompt features a few entries and several labels
    #that are self explanatory
    prompt = ctk.CTkToplevel(master=progFrame)
    prompt.title('Missquote Helper')
    prompt.wm_transient(progFrame)
    window = center_root(width=400,
                  height=300,
                  screenX=prompt.winfo_screenwidth(),
                  screenY=prompt.winfo_screenheight())
    prompt.geometry(window.default())
    label = ctk.CTkLabel(master=prompt,
                             text='Enter associated values',
                             text_color=default_text,
                             font=('Calibri',20)
                             ).place(relx=.5,rely=.1,anchor='center')
    originalA = ctk.CTkEntry(master=prompt,
                                      placeholder_text='OG A-Side',
                                      placeholder_text_color=default_text,
                                      text_color=default_text,
                                      width=80
                                      )
    originalA.place(relx=.25,rely=.25,anchor='center')
    newA = ctk.CTkEntry(master=prompt,
                                      placeholder_text='M# A-Side',
                                      placeholder_text_color=default_text,
                                      text_color=default_text,
                                      width=80
                                      )
    newA.place(relx=.50,rely=.25,anchor='center')
    originalB = ctk.CTkEntry(master=prompt,
                                      placeholder_text='OG B-Side',
                                      placeholder_text_color=default_text,
                                      text_color=default_text,
                                      width=80
                                      )
    originalB.place(relx=.75,rely=.25,anchor='center')
    divide = ctk.CTkLabel(master=prompt,
                             text='/',
                             text_color=default_text,
                             font=('Calibri',20)
                             ).place(relx=.375,rely=.25,anchor='center')
    multiply = ctk.CTkLabel(master=prompt,
                             text='x',
                             text_color=default_text,
                             font=('Calibri',20)
                             ).place(relx=.625,rely=.25,anchor='center')
    equal = ctk.CTkLabel(master=prompt,
                             text='=',
                             text_color=default_text,
                             font=('Calibri',20)
                             ).place(relx=.9,rely=.25,anchor='center')
    output = ctk.CTkLabel(master=prompt,
                             text='Press submit to calculate',
                             text_color=default_text,
                             font=('Calibri',20)
                             )
    output.place(relx=.5,rely=.4,anchor='center')
    submit = ctk.CTkButton(master=prompt,
                           text='Submit',
                           text_color=default_text,
                           command=lambda:calculate(event=None)
                           )
    submit.place(relx=.5,rely=.55,anchor='center')
    prompt.bind('<Return>',calculate)
    

    
