import customtkinter as ctk
from tkinter import DISABLED,NORMAL,END,LEFT
from misc_commands import center_root,mmm_macro,pp_macro,get_control
import pygetwindow as gw
import MRH
import Proforma
import AR_Issues
import pyperclip
import os
#imports

#our MRH class displays a popup for an AR user if selected
class Morning_Report_Helper:
    def __init__(self,progFrame,as400_username,as400_password,default_text):
        self.progFrame = progFrame
        self.as400_username = as400_username
        self.as400_password = as400_password
        self.default_text = default_text

    def mmddyy(self):
        #step one of this program is to get the previous business day (easiest way probably user input)
        def submitEnter(event):
            #get each of our date fields on submit button
            month = str(monthEntry.get())
            day = str(dayEntry.get())
            year = str(yearEntry.get())
            #basic check process and visual response if failed
            try:
                if 1 <= int(month) <= 12 and len(month) == 2 and 1 <= int(day) <= 31 and len(day) == 2 and len(year) == 4:
                    self.MRH_programs(month,day,year)
                else:
                    prompt.after(250,lambda:failLabel.configure(text_color='Red'))
                    prompt.after(500,lambda:failLabel.configure(text_color=self.default_text))
                    prompt.after(750,lambda:failLabel.configure(text_color='Red'))
                    prompt.after(1000,lambda:failLabel.configure(text_color=self.default_text))
            except Exception as e:
                print(e)
                prompt.after(250,lambda:failLabel.configure(text_color='Red'))
                prompt.after(500,lambda:failLabel.configure(text_color=self.default_text))
                prompt.after(750,lambda:failLabel.configure(text_color='Red'))
                prompt.after(1000,lambda:failLabel.configure(text_color=self.default_text))
        #cleanup non-used prompts
        global prompt
        try:
            prompt.destroy()
        except:
            pass
        #creating our popup for getting previous business day date from user
        prompt = ctk.CTkToplevel(master=self.progFrame)
        prompt.title('Enter Dates for MRH')
        prompt.wm_transient(self.progFrame)
        window = center_root(width=300,
                      height=400,
                      screenX=prompt.winfo_screenwidth(),
                      screenY=prompt.winfo_screenheight())
        prompt.geometry(window.default())
        #basic prompt of three entries, a button and a label
        monthEntry = ctk.CTkEntry(master=prompt,
                                  placeholder_text='Month',
                                  placeholder_text_color=self.default_text,
                                  text_color=self.default_text
                                  )
        monthEntry.pack(pady=12,padx=10)
        dayEntry = ctk.CTkEntry(master=prompt,
                                placeholder_text='Day',
                                placeholder_text_color=self.default_text,
                                text_color=self.default_text,
                                )
        dayEntry.pack(pady=12,padx=10)
        yearEntry = ctk.CTkEntry(master=prompt,
                                 placeholder_text='Year',
                                 placeholder_text_color=self.default_text,
                                 text_color=self.default_text
                                 )
        yearEntry.pack(pady=12,padx=10)
        submit_button = ctk.CTkButton(master=prompt,
                                      text = 'Proceed to Reports',
                                      text_color=self.default_text,
                                      command=lambda:submitEnter(event=None)
                                      ).pack(pady=12,padx=10)        
        prompt.bind('<Return>',submitEnter)
        failLabel = ctk.CTkLabel(master=prompt,
                                 text='Must be MM/DD/YYYY'+'\n'+'Always Backdate',
                                 font=('Calibri',18),
                                 text_color=self.default_text
                                 )
        failLabel.pack(pady=12,padx=10)
    
    def MRH_programs(self,month,day,year):
        #if previous prompt is successful, 4 different programs will be available to run
        def MRH_choice(choice,button):
            #run a single program based on user input, only applies to 3/4
            try:
                choice()
                button.configure(text='Success!')
                button.configure(state=DISABLED)
            except Exception as e:
                print(e)
                button.configure(text='Error Copied! Use CTRL+V')
                pyperclip.copy(str(e))
                prompt.after(5000,lambda:button.configure(text='Try Again'))
        global prompt
        try:
            prompt.destroy()
        except:
            pass
        #create basic prompt that shows each morning report program button to user
        prompt = ctk.CTkToplevel(master=self.progFrame)
        prompt.title('Morning Report Options')
        window = center_root(width=400,
                      height=300,
                      screenX=prompt.winfo_screenwidth(),
                      screenY=prompt.winfo_screenheight())
        prompt.geometry(window.default())
        #basic looking options for user
        label = ctk.CTkLabel(master=prompt,
                             text='MRH Choices',
                             text_color=self.default_text,
                             font=('Calibri',20)
                             ).place(relx=.5,rely=.075,anchor='center')
        oedailyButton = ctk.CTkButton(master=prompt,
                                          text='OEDaily',
                                          width=150,height=35,
                                          text_color=self.default_text,
                                          command=lambda:MRH_choice(lambda:MRH.oedaily(month,
                                                                                day,
                                                                                year,
                                                                                self.as400_username,
                                                                                self.as400_password),
                                                                    oedailyButton)
                                          )
        oedailyButton.place(relx=.5,rely=.2,anchor='center')
        taxButton = ctk.CTkButton(master=prompt,
                                          text='Tax Query',
                                          width=150,height=35,
                                          text_color=self.default_text,
                                          command=lambda:MRH_choice(lambda:MRH.tax(month,
                                                                                day,
                                                                                year,
                                                                                self.as400_username,
                                                                                self.as400_password),
                                                                    taxButton)
                                          )
        taxButton.place(relx=.5,rely=.4,anchor='center')
        freezeButton = ctk.CTkButton(master=prompt,
                                          text='Freeze Query',
                                          width=150,height=35,
                                          text_color=self.default_text,
                                          command=lambda:MRH_choice(lambda:MRH.freeze(month,
                                                                                day,
                                                                                year,
                                                                                self.as400_username,
                                                                                self.as400_password),
                                                                    freezeButton)
                                          )
        freezeButton.place(relx=.5,rely=.6,anchor='center')
        finishedGoodsButton = ctk.CTkButton(master=prompt,
                                          text='Finished Goods',
                                          width=150,height=35,
                                          text_color=self.default_text,
                                          command=lambda:self.MRH_fg(finishedGoodsButton,prompt,month,day,year)
                                          )
        finishedGoodsButton.place(relx=.5,rely=.8,anchor='center')
        
    #if finishedgoods button is selected a more intricate process has to be done by user
    def MRH_fg(self,button,prompt,month,day,year):
        def submitEnter(event):
            #after a basic check, the program creates another prompt that allows the user to perform data entry tasks easily
            beginning = str(beginningEntry.get())
            end = str(endEntry.get())
            try:
                if len(beginning) == 7 and len(end) == 7:
                    numbers = MRH.finishedGoods(month,day,year,self.as400_username,self.as400_password,beginning,end)
                    fgWindow.destroy()
                    button.configure(text='Success!')
                    button.configure(state=DISABLED)
                    #the prompt basically allows user to run two different macros into two different excel sheets
                    copyWindow = ctk.CTkToplevel(master=prompt)
                    copyWindow.wm_transient(prompt)
                    copyWindow.title('Finished Goods EZ Copy')
                    window = center_root(width=400,
                                  height=300,
                                  screenX=copyWindow.winfo_screenwidth(),
                                  screenY=copyWindow.winfo_screenheight())
                    copyWindow.geometry(window.default())
                    label = ctk.CTkLabel(master=copyWindow,
                                         text='Note: Select beginning cells before running macros',
                                         font=('Calibri',18),
                                         width=150,wraplength=250,
                                         text_color=self.default_text
                                         ).pack(pady=12,padx=10)
                    getPP = ctk.CTkButton(master=copyWindow,
                                  text='Run PP Macro',
                                  text_color=self.default_text,
                                  command=lambda:pp_macro(numbers)
                                  ).pack(pady=12,padx=10)
                    getmmm = ctk.CTkButton(master=copyWindow,
                                  text='Run 3M Macro',
                                  text_color=self.default_text,
                                  command=lambda:mmm_macro(numbers)
                                  ).pack(pady=12,padx=10)
                    
                else:
                    fgWindow.after(250,lambda:failLabel.configure(text_color='Red'))
                    fgWindow.after(500,lambda:failLabel.configure(text_color=self.default_text))
                    fgWindow.after(750,lambda:failLabel.configure(text_color='Red'))
                    fgWindow.after(1000,lambda:failLabel.configure(text_color=self.default_text))
            except Exception as e:
                print(e)
                fgWindow.destroy()
                button.configure(text='Error Copied! Use CTRL+V')
                pyperclip.copy(str(e))
                prompt.after(5000,lambda:button.configure(text='Try Again'))
        
        #basic looking prompt with necessary steps to interface a macro with the as400 and run our reports automatically
        fgWindow = ctk.CTkToplevel(master=prompt)
        fgWindow.wm_transient(prompt)
        fgWindow.title('Finished Goods')
        window = center_root(width=400,
                      height=300,
                      screenX=fgWindow.winfo_screenwidth(),
                      screenY=fgWindow.winfo_screenheight())
        fgWindow.geometry(window.default())
        failLabel = ctk.CTkLabel(master=fgWindow,
                             text='Must be 7 digits',
                             font=('Calibri',18),
                             text_color=self.default_text
                             )
        failLabel.pack(pady=12,padx=10)
        beginningEntry = ctk.CTkEntry(master=fgWindow,
                                  placeholder_text='Start Control',
                                  placeholder_text_color=self.default_text,
                                  text_color=self.default_text
                                  )
        beginningEntry.pack(pady=12,padx=10)
        getControl = ctk.CTkButton(master=fgWindow,
                                  text='Get last control(Session A)',
                                  text_color=self.default_text,
                                  command=lambda:get_control(beginningEntry.get())
                                  ).pack(pady=12,padx=10)
        endEntry = ctk.CTkEntry(master=fgWindow,
                                  placeholder_text='End Control',
                                  placeholder_text_color=self.default_text,
                                  text_color=self.default_text
                                  )
        endEntry.pack(pady=12,padx=10)
        submit_button = ctk.CTkButton(master=fgWindow,
                                  text='Run Queries',
                                  text_color=self.default_text,
                                  command=lambda:submitEnter(event=None)
                                  ).pack(pady=12,padx=10)
        fgWindow.bind('<Return>',submitEnter)
        os.startfile('M:/Billing/Job Count Range _'+year+'.xlsx')
        
#class for running a prepayment estimate (basic pdf that shows an estimate along with basic order info for customer)
class prepayment_estimate:
    def __init__(self,progFrame,as400_username,as400_password,default_text):
        self.progFrame = progFrame
        self.as400_username = as400_username
        self.as400_password = as400_password
        self.default_text = default_text
    #window that has two entries for user to input order number and an email if they wish to send one (on generation, the file path is copied)
    def proforma_window(self):
        def generate(event):
            #after a basic check, an estimate is generated, saved and copied to clipboard, otherwise it will display visual effects
            try:
                if len(str(generateEntry.get())) == 7:
                    estimate = Proforma.build(generateEntry.get(),self.as400_username,self.as400_password)
                    emailButton.configure(text='Send C#'+str(generateEntry.get()))
                    emailButton.configure(state=NORMAL)
                    generateButton.configure(text='Copied file!')
                    prompt.after(15000,lambda:generateButton.configure(text='Create New & Copy'))
                    generateEntry.delete(0, END)
                    emailEntry.focus_set()
                else:
                    generateButton.configure(text='Must be 7 digits.')
                    prompt.after(1000,lambda:generateButton.configure(text='Create New & Copy'))
            except Exception as e:
                generateButton.configure(text='Ctrl+v for error')
                pyperclip.copy(str(e))
                prompt.after(2000,lambda:generateButton.configure(text='Create New & Copy'))
        def email():
            #an invoice must be generated first, after a generic email can be sent with attached document
            try:
                emailButton.configure(state=DISABLED)
                Proforma.sendEmail(emailEntry.get())
                emailButton.configure(text='Success!')
                emailEntry.delete(0, END)
                generateEntry.focus_set()
                prompt.after(3000,lambda:emailButton.configure(text='Send C#'+str(generateEntry.get())))
            except Exception as e:
                emailButton.configure(text='Ctrl+v for error')
                pyperclip.copy(str(e))
                prompt.after(2000,lambda:emailButton.configure(text='Send C#'+str(generateEntry.get())))
        global prompt
        try:
            prompt.destroy()
        except:
            pass
        #basic prompt with two entries, two buttons and a label
        prompt = ctk.CTkToplevel(master=self.progFrame)
        prompt.title('Prepayment Estimate')
        prompt.wm_transient(self.progFrame)
        window = center_root(width=400,
                      height=300,
                      screenX=prompt.winfo_screenwidth(),
                      screenY=prompt.winfo_screenheight())
        prompt.geometry(window.default())
        label = ctk.CTkLabel(master=prompt,
                                 text='Create estimate and paste'+'\n'+'into Outlook or send email',
                                 text_color=self.default_text,
                                 font=('Calibri',20),
                                 justify='left'
                                 )
        label.pack(pady=12,padx=10,anchor='center')
        generateEntry = ctk.CTkEntry(master=prompt,
                                  placeholder_text='Control Number',
                                  placeholder_text_color=self.default_text,
                                  text_color=self.default_text
                                  )
        generateEntry.pack(pady=12,padx=10)
        generateButton = ctk.CTkButton(master=prompt,
                                       text='Create New & Copy',
                                       text_color=self.default_text,
                                       command=lambda:generate(event=None)
                                       )
        generateButton.pack(pady=12,padx=10)
        prompt.bind('<Return>',generate)
        emailEntry = ctk.CTkEntry(master=prompt,
                                  placeholder_text='Requester Email',
                                  width=250,
                                  placeholder_text_color=self.default_text,
                                  text_color=self.default_text
                                  )
        emailEntry.pack(pady=12,padx=10)
        emailButton = ctk.CTkButton(master=prompt,
                                    text='Send',
                                    text_color=self.default_text,
                                    command=email,
                                    state=DISABLED
                                    )
        emailButton.pack(pady=12,padx=10)
        
#class for AR issues program
class ar_issues:
    def __init__(self,progFrame,as400_username,as400_password,default_text):
        self.progFrame = progFrame
        self.as400_username = as400_username
        self.as400_password = as400_password
        self.default_text = default_text
    #creates a prompt with options of prepay followup and to edit holidays (for calculating scheduled followups)
    def ar_issues(self):
        def prepays():#prepay followup program
            def send():
                #option to send followups from available batch
                try:
                    AR_Issues.sendEmail(prepayList)
                    prompt.destroy()
                    successPrompt = ctk.CTkToplevel(master=self.progFrame)
                    successPrompt.title('Prepay Follow-up')
                    successPrompt.wm_transient(self.progFrame)
                    visual = 'List that was sent prepayment follow-ups.'+'\n'+'\n'
                    for item in prepayList:
                        if item[7] == True:
                            visual = visual + 'A#'+str(item[0])+' C#'+str(item[1])+ ' '+item[4] + '\n'
                    window = center_root(width=800,
                                  height=600,
                                  screenX=successPrompt.winfo_screenwidth(),
                                  screenY=successPrompt.winfo_screenheight())
                    successPrompt.geometry(window.default())
                    label = ctk.CTkLabel(master=successPrompt,
                                         text=visual,
                                         text_color=self.default_text,
                                         font=('Calibri',20),
                                         justify='left'
                                         )
                    label.pack(pady=12,padx=10,anchor='center')
                except Exception as e:
                    emailBatchButton.configure(text='Ctrl+v for error')
                    pyperclip.copy(str(e))
                    prepayPrompt.after(2000,lambda:emailBatchButton.configure(text='Send Emails'))
            def remove_control():
                #if we want to remove a control from being sent out
                global prepayList
                removed = False
                for item in prepayList:
                    if controlEntry.get() == str(item[1]):
                        item[7] = False
                        removed = True
                        visual = 'Batch that is set to send:'+'\n'+'\n'
                        for item in prepayList:
                            if item[7] == True:
                                visual = visual + 'A#'+str(item[0])+' C#'+str(item[1])+ ' '+item[4] + '\n'
                        label.configure(text=visual)
                        controlEntry.delete(0, END)
                if removed is False:
                    removeButton.configure(text='Invalid Control')
                    prepayPrompt.after(2000,lambda:removeButton.configure(text='Remove Control'))
                    controlEntry.delete(0, END)
            def add_control():
                #if we want to add a control that is "hidden" or has been accidently removed
                global prepayList
                removed = False
                for item in prepayList:
                    if controlEntry.get() == str(item[1]):
                        item[7] = True
                        removed = True
                        visual = 'Batch that is set to send:'+'\n'+'\n'
                        for item in prepayList:
                            if item[7] == True:
                                visual = visual + 'A#'+str(item[0])+' C#'+str(item[1])+ ' '+item[4] + '\n'
                        label.configure(text=visual)
                        controlEntry.delete(0, END)
                if removed is False:
                    addButton.configure(text='Invalid Control')
                    prepayPrompt.after(2000,lambda:addButton.configure(text='Add Control'))
                    controlEntry.delete(0, END)
            #prompt that shows the batch of orders that will be followed up on, also has a few different buttons for adding jobs, removing jobs and sending emails
            prepayPrompt = ctk.CTkToplevel(master=prompt)
            prepayPrompt.title('Prepay Follow-up')
            prepayPrompt.wm_transient(prompt)
            visual = 'Batch that is set to send:'+'\n'+'\n'
            #creates visual list
            for item in prepayList:
                if item[7] == True:
                    visual = visual + 'A#'+str(item[0])+' C#'+str(item[1])+ ' '+item[4] + '\n'
            window = center_root(width=800,
                          height=600,
                          screenX=prepayPrompt.winfo_screenwidth(),
                          screenY=prompt.winfo_screenheight())
            prepayPrompt.geometry(window.default())
            label = ctk.CTkLabel(master=prepayPrompt,
                                 text=visual,
                                 text_color=self.default_text,
                                 font=('Calibri',20),
                                 justify='left'
                                 )
            label.pack(pady=12,padx=10,anchor='center')
            controlEntry = ctk.CTkEntry(master=prepayPrompt,
                                        placeholder_text='Add/Remove Control',
                                        placeholder_text_color=self.default_text,
                                        text_color=self.default_text
                                        )
            controlEntry.pack(pady=12,padx=10,anchor='center')
            removeButton = ctk.CTkButton(master=prepayPrompt,
                                         text='Remove Control',
                                         text_color=self.default_text,
                                         command=remove_control
                                         )
            removeButton.pack(pady=12,padx=10,anchor='center')
            addButton = ctk.CTkButton(master=prepayPrompt,
                                         text='Add Control',
                                         text_color=self.default_text,
                                         command=add_control
                                         )
            addButton.pack(pady=12,padx=10,anchor='center')
            emailBatchButton = ctk.CTkButton(master=prepayPrompt,
                                         text='Send Emails',
                                         text_color=self.default_text,
                                         command=send
                                         )
            emailBatchButton.pack(pady=12,padx=10,anchor='center')
        def holidays():
            #prompt that allows a user to adjust dates for the followup process, add non-business days (weekends are already calculated)
            def submitEnter(event):
                #get each of our values for the total date
                month = str(monthEntry.get())
                day = str(dayEntry.get())
                year = str(yearEntry.get())
                #after a brief check, a date is added to a text document and saved for later use
                try:
                    if 1 <= int(month) <= 12 and len(month) == 2 and 1 <= int(day) <= 31 and len(day) == 2 and len(year) == 4:
                        AR_Issues.holidayEdit(month,day,year)
                        monthEntry.delete(0, END)
                        dayEntry.delete(0, END)
                        yearEntry.delete(0, END)
                        submit_button.configure(text='Success!')
                        monthEntry.focus_set()
                        holidayPrompt.after(2000,lambda:submit_button.configure(text='Add Holiday Date'))
                    else:
                        holidayPrompt.after(250,lambda:failLabel.configure(text_color='Red'))
                        holidayPrompt.after(500,lambda:failLabel.configure(text_color=self.default_text))
                        holidayPrompt.after(750,lambda:failLabel.configure(text_color='Red'))
                        holidayPrompt.after(1000,lambda:failLabel.configure(text_color=self.default_text))
                except Exception as e:
                    print(e)
                    holidayPrompt.after(250,lambda:failLabel.configure(text_color='Red'))
                    holidayPrompt.after(500,lambda:failLabel.configure(text_color=self.default_text))
                    holidayPrompt.after(750,lambda:failLabel.configure(text_color='Red'))
                    holidayPrompt.after(1000,lambda:failLabel.configure(text_color=self.default_text))
            #prompt that shows 3 entries and an add date button
            holidayPrompt = ctk.CTkToplevel(master=prompt)
            holidayPrompt.title('Holidays for AR-Issues')
            holidayPrompt.wm_transient(prompt)
            window = center_root(width=300,
                          height=400,
                          screenX=holidayPrompt.winfo_screenwidth(),
                          screenY=holidayPrompt.winfo_screenheight())
            holidayPrompt.geometry(window.default())
            
            monthEntry = ctk.CTkEntry(master=holidayPrompt,
                                      placeholder_text='Month',
                                      placeholder_text_color=self.default_text,
                                      text_color=self.default_text
                                      )
            monthEntry.pack(pady=12,padx=10)
            dayEntry = ctk.CTkEntry(master=holidayPrompt,
                                    placeholder_text='Day',
                                    placeholder_text_color=self.default_text,
                                    text_color=self.default_text,
                                    )
            dayEntry.pack(pady=12,padx=10)
            yearEntry = ctk.CTkEntry(master=holidayPrompt,
                                     placeholder_text='Year',
                                     placeholder_text_color=self.default_text,
                                     text_color=self.default_text
                                     )
            yearEntry.pack(pady=12,padx=10)
            submit_button = ctk.CTkButton(master=holidayPrompt,
                                          text = 'Add Holiday Date',
                                          text_color=self.default_text,
                                          command=lambda:submitEnter(event=None)
                                          )
            submit_button.pack(pady=12,padx=10)        
            holidayPrompt.bind('<Return>',submitEnter)
            failLabel = ctk.CTkLabel(master=holidayPrompt,
                                     text='Must be MM/DD/YYYY',
                                     font=('Calibri',18),
                                     text_color=self.default_text
                                     )
            failLabel.pack(pady=12,padx=10)
        #cleans out prompts
        global prepayList,prompt
        try:
            prompt.destroy()
        except:
            pass
        #gets prepay list for above use and create prompt with options to add holidays and followup on prepay orders
        prepayList = AR_Issues.main()
        prompt = ctk.CTkToplevel(master=self.progFrame)
        prompt.title('AR-Issues Follow-up')
        prompt.wm_transient(self.progFrame)
        window = center_root(width=400,
                      height=300,
                      screenX=prompt.winfo_screenwidth(),
                      screenY=prompt.winfo_screenheight())
        prompt.geometry(window.default())
        arLabel = ctk.CTkLabel(master=prompt,
                                 text='AR-Issues',
                                 text_color=self.default_text,
                                 font=('Calibri',20)
                               ).pack(pady=12,padx=10)
        prepayButton = ctk.CTkButton(master=prompt,
                                       text='Prepay follow-ups',
                                       text_color=self.default_text,
                                       command=prepays
                                     )
        prepayButton.pack(pady=12,padx=10)
        holidayButton = ctk.CTkButton(master=prompt,
                                       text='Add Holidays',
                                       text_color=self.default_text,
                                       command=holidays
                                     )
        holidayButton.pack(pady=12,padx=10)
                                     
        
