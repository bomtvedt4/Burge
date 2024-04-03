""" Burge
The goal of this program is to automate tasks and
improve quality of life for any promotional market
employee who uses it. Burge is intended to be an
ongoing project that improves as suggestions and
changes are made to it.
"""
import customtkinter as ctk
import os,os.path
from itoolkit.transport import DatabaseTransport
import pyodbc
import pyperclip
import psutil
#imports from outside libraries
from misc_commands import center_root,get_options,setup_options,get_appearance,get_text
from menuframes import *
#imports from other parts of Burge
version = 'Burge v1.7'

#one of two primary loops that Burge relies, basic login page that verifies against local users ODBC connection           
def login():
    #function that checks user login credentials, if success proceeds to main menu
    def submitEnter(event):
        global as400_username,as400_password
        try:
            #get username and password entries before attemping to login using those credentials
            as400_username = user_name_entry.get()
            as400_password = user_password_entry.get()
            pyodbc.connect("DSN=QDSN_AS400.PP.COM",uid=as400_username,password=as400_password)
            loginRoot.destroy()
            mainMenu()
        except Exception as e:
            #if unable to login for any reason, error is copied and visual effects are displayed
            pyperclip.copy(str(e))
            failLabel = ctk.CTkLabel(master=frame,
                                     text= 'Invalid User/Pass',
                                     font=('Calibri',18),
                                     text_color='Red'
                                     )
            failLabel.pack(pady=12,padx=10)
            loginRoot.after(250,lambda:failLabel.configure(text_color=default_text))
            loginRoot.after(500,lambda:failLabel.configure(text_color='Red'))
            loginRoot.after(750,lambda:failLabel.configure(text_color=default_text))
            loginRoot.after(1000,lambda:failLabel.destroy())

    #default view for the user        
    ctk.set_appearance_mode('dark')
    ctk.set_default_color_theme('dark-blue')
    default_text = 'white'
    #create login root and center on main screen
    loginRoot = ctk.CTk()
    loginRoot.title(version)
    window = center_root(width=300,
                      height=400,
                      screenX=loginRoot.winfo_screenwidth(),
                      screenY=loginRoot.winfo_screenheight())
    loginRoot.geometry(window.default())
    loginRoot.resizable(0,0)

    #basic looking login page created use the below frame, label, button and entries
    frame = ctk.CTkFrame(master=loginRoot)
    frame.pack(pady=20,padx=40,fill='both',expand=True)

    ctk.CTkLabel(master=frame,
                 text=version+'\nAS400 Login',
                 text_color=default_text).pack(pady=12,padx=10)
    
   
    user_name_entry = ctk.CTkEntry(master=frame,
                                   placeholder_text='Username',
                                   placeholder_text_color=default_text,
                                   text_color=default_text
                                   )
    user_name_entry.pack(pady=12,padx=10)
    user_password_entry = ctk.CTkEntry(master=frame,
                                       placeholder_text='Password',
                                       placeholder_text_color=default_text,
                                       text_color=default_text,
                                       show = '*')
    user_password_entry.pack(pady=12,padx=10)
    
    submit_button = ctk.CTkButton(master=frame,
                                  text = 'Login',
                                  text_color=default_text,
                                  command=lambda: submitEnter(event=None)
                                  ).pack(pady=12,padx=10)
    loginRoot.bind('<Return>',submitEnter)

    #allows someone to destroy any active sessions to make an update or for whatever other reason
    def killSwitch():
        if os.path.isfile('alphasw.txt'):
            loginRoot.destroy()
        else:
            loginRoot.after(1000,killSwitch)
    loginRoot.after(1000,killSwitch)
    loginRoot.after(100,user_name_entry.focus_set)
    loginRoot.mainloop()

#second of two main loops that Burge runs on
def mainMenu():
    #interface settings are loaded fresh or use previous choices
    setup_options(as400_username)
    options = get_options(as400_username)
    default_text = get_text(options)
    appearance = get_appearance(options)
    ctk.set_appearance_mode(appearance)

    #creates our mainloop and centers
    menu_root = ctk.CTk()
    menu_root.title(version)
    window = center_root(width=1280,
                      height=720,
                      screenX=menu_root.winfo_screenwidth(),
                      screenY=menu_root.winfo_screenheight())
    menu_root.geometry(window.default())
    menu_root.resizable(True,True)
    #next portion is a basic setup to allow a user to resize screen but still navigate using arrow keys
    menu_canvas = ctk.CTkCanvas(menu_root)
    scrollFrame = ctk.CTkFrame(menu_root,
                                         width=menu_root.winfo_screenwidth()+50,
                                         height=menu_root.winfo_screenheight()+50,
                                         bg_color='black',
                                         fg_color='black',
                                         border_color='black')
    menu_canvas.configure(scrollregion = (5,5,1280,720))
        
    menu_root.bind('<Up>',lambda event: menu_canvas.yview_scroll(-1, 'units'))
    menu_root.bind('<Down>',lambda event: menu_canvas.yview_scroll(1, 'units'))
    menu_root.bind('<Left>',lambda event: menu_canvas.xview_scroll(-1, 'units'))
    menu_root.bind('<Right>',lambda event: menu_canvas.xview_scroll(1, 'units'))
    
    menu_canvas.pack(fill='both', side='left', expand='true')
    menu_canvas.create_window(0, 0, window=scrollFrame, anchor='nw')
    
                                
    #quickFrame start
    #many basic helpful programs can be found here as well as choosing visual settings
    quickFrame = ctk.CTkFrame(master=scrollFrame,width=310,height=700,corner_radius=15)
    quickFrame.place(x=10,y=10)
    
    quickFrameSetup = quick_frame(quickFrame,default_text,as400_username,as400_password,menu_root,options,login,mainMenu)
    quickFrameSetup.quickApps()
    quickFrameSetup.quickLinks()
    quickFrameSetup.invoice_credit()
    quickFrameSetup.modeChoices()
    
    #quickFrame end and desFrame start
    #really basic frame with the version and an option to display help
    desFrame = ctk.CTkFrame(master=scrollFrame,width=780,height=220,corner_radius=15)
    desFrame.place(x=330,y=10)

    desFrameSetup = des_frame(desFrame,as400_username,default_text,version)
    desFrameSetup.welcome()

    #desFrame end and progFrame start
    #frame that reveals any smaller programs a given user is allowed to use
    progFrame = ctk.CTkFrame(master=scrollFrame,width=940,height=470,corner_radius=15)
    progFrame.place(x=330,y=240)

    progFrameSetup = prog_frame(progFrame,as400_username,as400_password,default_text)
    progFrameSetup.valid_programs()
    
    #progFrame end and userFrame start
    #basic exit/logout options as well as username, a program that quick logs a user to AS400 and an active time
    userFrame = ctk.CTkFrame(master=scrollFrame,width=150,height=220,corner_radius=15)
    userFrame.place(x=1120,y=10)

    userFrameSetup = user_frame(userFrame,as400_username,as400_password,default_text,menu_root,login)
    userFrameSetup.time_label()
    userFrameSetup.user_label()
    userFrameSetup.quick_log()
    userFrameSetup.exit_options()

    #userFrame end
    #allows someone to destroy any active sessions to make an update or for whatever other reason
    def killSwitch():
        if os.path.isfile('alphasw.txt'):
            menu_root.destroy()
        else:
            menu_root.after(1000,killSwitch)
    menu_root.after(1000,killSwitch)
    menu_root.mainloop()
    #due to programming limitations a program has to be run outside of Burge, this cleans it up if it is still open
    try:
        if 'popshelper.exe' in (p.name() for p in psutil.process_iter()):
            os.system("taskkill /f /im  popshelper.exe")
    except:
        os.system("taskkill /f /im  popshelper.exe")

login()
