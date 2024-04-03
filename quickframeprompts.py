import customtkinter as ctk
from misc_commands import center_root,get_options,get_apps,get_links
from quickframecommands import *
import datetime
#imports
def quickAppsPrompt(quickFrame,default_text,as400_username):
    #function that shows if user wants to edit their quick apps
    #clear prompt if it exists
    global prompt        
    try:
        prompt.destroy()
    except:
        pass
    #creates a button and two lists that each function as their name indicates
    #add app is unique in the sense that tkinter is blocked until the user has
    #copied something to their clipboard
    #this was the best way I could find with available libraries that could copy
    #a path directly from an icon such as a shortcut on the desktop
    apps = get_apps(get_options(as400_username))
    prompt = ctk.CTkToplevel(master=quickFrame)
    prompt.title('Quick Apps')
    prompt.wm_transient(quickFrame)
    window = center_root(width=300,
                  height=300,
                  screenX=prompt.winfo_screenwidth(),
                  screenY=prompt.winfo_screenheight())
    prompt.geometry(window.default())
    quickAppsLabel = ctk.CTkLabel(master=prompt,
                 text='Quick Apps',
                 font=('Calibri',20,'bold'),
                 text_color=default_text
                 ).place(relx=.5,rely=.1,anchor='center')
    quickAppsAdd = ctk.CTkButton(master=prompt,
                  text='Add App',
                  width=90,
                  height=35,
                  text_color=default_text
                  )
    quickAppsAdd.place(relx=.5,rely=.25,anchor='center')
    delapp = ctk.StringVar(value='Del App')
    quickAppsRemove = ctk.CTkOptionMenu(master=prompt,
                      variable=delapp,
                      values=apps,
                      width=90,
                      height=35,
                      text_color=default_text
                      )
    quickAppsRemove.place(relx=.33,rely=.40,anchor='center')
    launchapp = ctk.StringVar(value='Launch')
    quickAppsLaunch = ctk.CTkOptionMenu(master=prompt,
                      variable=launchapp,
                      values=apps,
                      width=90,
                      height=35,
                      text_color=default_text
                      )
    quickAppsLaunch.place(relx=.67,rely=.40,anchor='center')
    quickAppsAdd.configure(command=lambda:(quickAppsAdd.configure(text='Ctrl+C on App'),
                                           prompt.after(100,lambda:(quickAppsSetup.add_app(as400_username,quickAppsRemove,quickAppsLaunch,prompt),
                                                                    quickAppsAdd.configure(text='Add App')))
                                           ))
    quickAppsRemove.configure(command=quickAppsSetup.remove_app_wrapper(as400_username,quickAppsRemove,quickAppsLaunch))
    quickAppsLaunch.configure(command=quickAppsSetup.launch_app_wrapper(quickAppsLaunch))

def quickLinksPrompt(quickFrame,default_text,as400_username):
    #function that shows if user wants to edit their quick links
    #clear prompt if it exists
    global prompt        
    try:
        prompt.destroy()
    except:
        pass
    #returns two option lists a button and entry
    links = get_links(get_options(as400_username))
    prompt = ctk.CTkToplevel(master=quickFrame)
    prompt.title('Quick Links')
    prompt.wm_transient(quickFrame)
    window = center_root(width=300,
                  height=300,
                  screenX=prompt.winfo_screenwidth(),
                  screenY=prompt.winfo_screenheight())
    prompt.geometry(window.default())
    quickLinksLabel = ctk.CTkLabel(master=prompt,
                 text='Quick Links',
                 font=('Calibri',20,'bold'),
                 text_color=default_text
                 ).place(relx=.5,rely=.1,anchor='center')
    quickLinksAdd = ctk.CTkButton(master=prompt,
                  text='Add Link',
                  width=90,
                  height=35,
                  text_color=default_text
                  )
    quickLinksAdd.place(relx=.2,rely=.25,anchor='center')
    quickLinksAddEntry = ctk.CTkEntry(master=prompt,
                                      placeholder_text='Browser Link',
                                      placeholder_text_color=default_text,
                                      text_color=default_text,
                                      width=150
                                      )
    quickLinksAddEntry.place(relx=.65,rely=.25,anchor='center')
    dellink = ctk.StringVar(value='Del Link')
    quickLinksRemove = ctk.CTkOptionMenu(master=prompt,
                      variable=dellink,
                      values=links,
                      width=90,
                      height=35,
                      text_color=default_text
                      )
    quickLinksRemove.place(relx=.33,rely=.40,anchor='center')
    launchlink = ctk.StringVar(value='Launch')
    quickLinksLaunch = ctk.CTkOptionMenu(master=prompt,
                      variable=launchlink,
                      values=links,
                      width=90,
                      height=35,
                      text_color=default_text
                      )
    quickLinksLaunch.place(relx=.67,rely=.40,anchor='center')
    quickLinksAdd.configure(command=lambda:quickLinksSetup.add_link(as400_username,quickLinksAddEntry,quickLinksRemove,quickLinksLaunch))
    quickLinksRemove.configure(command=quickLinksSetup.remove_link_wrapper(as400_username,quickLinksRemove,quickLinksLaunch))
    quickLinksLaunch.configure(command=quickLinksSetup.launch_link_wrapper(quickLinksLaunch))

def invoicePrompt(quickFrame,default_text,as400_username,as400_password):
    #function for running opening a send invoice prompt
    def generate(event):
        #generates an invoice, this must be run first before being able
        #to send an email
        try:
            global invoiceNumberGlobal,invoiceYearGlobal
            invoiceNumberGlobal = invoiceEntry.get()
            invoiceYearGlobal = invoiceYearEntry.get()
            invClass = invoiceGen(str(invoiceEntry.get()),str(invoiceYearEntry.get()),as400_username,as400_password)
            invClass.buildInvoice(invClass.header())
            emailButton.configure(text='Send '+str(invoiceEntry.get())+'-'+str(invoiceYearEntry.get()))
            emailButton.configure(state='normal')
            generateButton.configure(text='Copied file!')
            prompt.after(15000,lambda:generateButton.configure(text='Create New & Copy'))
            invoiceEntry.delete(0, 'end')
            invoiceYearEntry.delete(0, 'end')
            invoiceYearEntry.insert(0, str(datetime.date.today().year))
            emailEntry.focus_set()
        except Exception as e:
            generateButton.configure(text='Ctrl+v for error')
            pyperclip.copy(str(e))
            prompt.after(2000,lambda:generateButton.configure(text='Create New & Copy'))
    def email():
        #sends an email with helpful information and the invoice attached
        try:
            emailButton.configure(state='disabled')
            invClass = invoiceGen(invoiceNumberGlobal,invoiceYearGlobal,as400_username,as400_password)
            invClass.sendEmail(emailEntry.get())
            emailButton.configure(text='Success!')
            emailEntry.delete(0, 'end')
            invoiceEntry.focus_set()
            prompt.after(3000,lambda:emailButton.configure(text='Send '+invoiceNumberGlobal+'-'+invoiceYearGlobal))
        except Exception as e:
            emailButton.configure(text='Ctrl+v for error')
            pyperclip.copy(str(e))
            prompt.after(2000,lambda:emailButton.configure(text='Send '+invoiceNumberGlobal+'-'+invoiceYearGlobal))
    #clears the prompt
    global prompt
    try:
        prompt.destroy()
    except:
        pass
    #prompt with two entries and two buttons, user will enter their invoice
    #and invoice year, to generate and inspect the invoice for verification
    #they can then choose to paste into a reply email or send one directly
    prompt = ctk.CTkToplevel(master=quickFrame)
    prompt.title('Invoice Generator')
    prompt.wm_transient(quickFrame)
    window = center_root(width=400,
                  height=300,
                  screenX=prompt.winfo_screenwidth(),
                  screenY=prompt.winfo_screenheight())
    prompt.geometry(window.default())
    label = ctk.CTkLabel(master=prompt,
                             text='Create invoice and paste'+'\n'+'into Outlook or send email',
                             text_color=default_text,
                             font=('Calibri',20),
                             justify='left'
                             )
    label.place(relx=.5,rely=.1,anchor='center')
    invoiceEntry = ctk.CTkEntry(master=prompt,
                              placeholder_text='Invoice No.',
                              placeholder_text_color=default_text,
                              text_color=default_text,
                              width=100
                              )
    invoiceEntry.place(relx=.4,rely=.25,anchor='center')
    invoiceYearEntry = ctk.CTkEntry(master=prompt,
                              placeholder_text='Year',
                              placeholder_text_color=default_text,
                              text_color=default_text,
                              width=50
                              )
    invoiceYearEntry.insert(0, str(datetime.date.today().year)) 
    invoiceYearEntry.place(relx=.6,rely=.25,anchor='center')
    generateButton = ctk.CTkButton(master=prompt,
                                   text='Create New & Copy',
                                   text_color=default_text,
                                   command=lambda:generate(event=None)
                                   )
    generateButton.place(relx=.5,rely=.4,anchor='center')
    prompt.bind('<Return>',generate)
    emailEntry = ctk.CTkEntry(master=prompt,
                              placeholder_text='Requester Email',
                              width=250,
                              placeholder_text_color=default_text,
                              text_color=default_text
                              )
    emailEntry.place(relx=.5,rely=.55,anchor='center')
    emailButton = ctk.CTkButton(master=prompt,
                                text='Send',
                                text_color=default_text,
                                command=email,
                                state='disabled'
                                )
    emailButton.place(relx=.5,rely=.70,anchor='center')

def creditPrompt(quickFrame,default_text,as400_username,as400_password):
    #function for running opening a send invoice prompt
    def generate(event):
        #generates an invoice, this must be run first before being able
        #to send an email
        try:
            global creditNumberGlobal,creditYearGlobal
            creditNumberGlobal = creditEntry.get()
            creditYearGlobal = creditYearEntry.get()
            creditMemoClass = CMClass()
            creditMemoClass.buildCM(creditNumberGlobal,creditYearGlobal,as400_username,as400_password)
            emailButton.configure(text='Send '+creditNumberGlobal+'-'+creditYearGlobal)
            emailButton.configure(state='normal')
            generateButton.configure(text='Copied file!')
            prompt.after(15000,lambda:generateButton.configure(text='Create New & Copy'))
            creditEntry.delete(0, 'end')
            creditYearEntry.delete(0, 'end')
            creditYearEntry.insert(0, str(datetime.date.today().year))
            emailEntry.focus_set()
        except Exception as e:
            generateButton.configure(text='Ctrl+v for error')
            pyperclip.copy(str(e))
            prompt.after(2000,lambda:generateButton.configure(text='Create New & Copy'))
    def email():
        #sends an email with helpful information and the invoice attached
        try:
            emailButton.configure(state='disabled')
            creditMemoClass = CMClass()
            creditMemoClass.sendEmail(emailEntry.get(),creditNumberGlobal,creditYearGlobal,as400_username,as400_password)
            emailButton.configure(text='Success!')
            emailEntry.delete(0, 'end')
            creditEntry.focus_set()
            prompt.after(3000,lambda:emailButton.configure(text='Send '+creditNumberGlobal+'-'+creditYearGlobal))
        except Exception as e:
            emailButton.configure(text='Ctrl+v for error')
            pyperclip.copy(str(e))
            prompt.after(2000,lambda:emailButton.configure(text='Send '+creditNumberGlobal+'-'+creditYearGlobal))
    #clears the prompt
    global prompt
    try:
        prompt.destroy()
    except:
        pass
    #prompt with two entries and two buttons, user will enter their invoice
    #and invoice year, to generate and inspect the invoice for verification
    #they can then choose to paste into a reply email or send one directly
    prompt = ctk.CTkToplevel(master=quickFrame)
    prompt.title('Credit Generator')
    prompt.wm_transient(quickFrame)
    window = center_root(width=400,
                  height=300,
                  screenX=prompt.winfo_screenwidth(),
                  screenY=prompt.winfo_screenheight())
    prompt.geometry(window.default())
    label = ctk.CTkLabel(master=prompt,
                             text='Create credit memo and paste'+'\n'+'into Outlook or send email',
                             text_color=default_text,
                             font=('Calibri',20),
                             justify='left'
                             )
    label.place(relx=.5,rely=.1,anchor='center')
    creditEntry = ctk.CTkEntry(master=prompt,
                              placeholder_text='Credit No.',
                              placeholder_text_color=default_text,
                              text_color=default_text,
                              width=100
                              )
    creditEntry.place(relx=.4,rely=.25,anchor='center')
    creditYearEntry = ctk.CTkEntry(master=prompt,
                              placeholder_text='Year',
                              placeholder_text_color=default_text,
                              text_color=default_text,
                              width=50
                              )
    creditYearEntry.insert(0, str(datetime.date.today().year)) 
    creditYearEntry.place(relx=.6,rely=.25,anchor='center')
    generateButton = ctk.CTkButton(master=prompt,
                                   text='Create New & Copy',
                                   text_color=default_text,
                                   command=lambda:generate(event=None)
                                   )
    generateButton.place(relx=.5,rely=.4,anchor='center')
    prompt.bind('<Return>',generate)
    emailEntry = ctk.CTkEntry(master=prompt,
                              placeholder_text='Requester Email',
                              width=250,
                              placeholder_text_color=default_text,
                              text_color=default_text
                              )
    emailEntry.place(relx=.5,rely=.55,anchor='center')
    emailButton = ctk.CTkButton(master=prompt,
                                text='Send',
                                text_color=default_text,
                                command=email,
                                state='disabled'
                                )
    emailButton.place(relx=.5,rely=.70,anchor='center')
