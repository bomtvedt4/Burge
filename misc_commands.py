import customtkinter as ctk
from datetime import date,datetime
from keyboard import press,release,is_pressed
import os
from time import sleep
import pyautogui as py
import pygetwindow as gw
import MRH
#imports
#this class was initally going allow windows to open based on where burge was at the time
#however pretty much only the default function is now used at the time of writing this
class center_root:
    def __init__(self,
                 width=None,
                 height=None,
                 screenX=None,
                 screenY=None,
                 custX=320,
                 custY=180):
        self.width = width
        self.height = height
        self.screenX = screenX
        self.screenY = screenY
        self.custX = custX
        self.custY = custY

    def custom(self):
        output = '%dx%d+%d+%d' % (self.width, self.height, self.custX, self.custY)
        return output

    def default(self):
        x = (self.screenX/2) - (self.width/2)
        y = (self.screenY/2) - (self.height/2)
        output = '%dx%d+%d+%d' % (self.width, self.height, x, y)
        return output

#function that records user choices and writes them to a local txt file
def write_options(as400_username,options):
    with open('UI Choices/'+as400_username+'/Personalizations.txt','w') as file:
        count = 0
        for opt in options:
            if count == 0:
                file.write(opt)
                count = 1
            else:
                file.write('\n'+opt)
#basically if no options have been specified for a user it sets these as a defualt
def setup_options(as400_username):
    uiSelections = 'UI Choices/'+as400_username+'/Personalizations.txt'
    if not os.path.exists(uiSelections):
        os.makedirs('UI Choices', exist_ok=True)
        os.makedirs('UI Choices/'+as400_username, exist_ok=True)
        options = ['default_text=white',
                   'set_appearance_mode=dark',
                   'quick_apps=',
                   'quick_links=']
        write_options(as400_username,options)
#whern we need to retrieve user options for whatever reason
def get_options(as400_username):
    uiSelections = 'UI Choices/'+as400_username+'/Personalizations.txt'
    with open(uiSelections,'r') as file:
        options = file.read()
        options = options.split('\n')
        return options
#return saved apps
def get_apps(options):
    apps = options[2].replace('quick_apps=','').split(',')
    return apps
#return saved links
def get_links(options):
    links = options[3].replace('quick_links=','').split(',')
    return links
#return appearance mode
def get_appearance(options):
    appearance = options[1].replace('set_appearance_mode=','')
    return appearance
#return text color
def get_text(options):
    default_text = options[0].replace('default_text=','')
    return default_text
#function to get time used in user frame
def get_time():
    today = date.today()
    now = datetime.now()
    string2 = today.strftime('%m/%d/%y')
    string1 = now.strftime('%I:%M:%S')
    return string1+' - '+string2
#quick log function used in user frame
def as400_quick_log(as400_username,as400_password):
    windows = gw.getWindowsWithTitle('Session')
    for window in windows:
        window.activate()
        py.write(as400_username)
        press('tab')
        release('tab')
        py.write(as400_password)
        press('enter')
        release('enter')
        press('enter')
        release('enter')
    window = gw.getWindowsWithTitle('Burge')[0]
    window.activate()
#wrapper function that allows a user to log out or exit
def exit_command_wrapper(menu_root,login):
    def exit_command(choice):
        if choice == 'Log Out':
            menu_root.destroy()
            login()
        elif choice == 'Exit':
            menu_root.destroy()
    return exit_command
#macro for one of the accounting programs, automates data entry
def mmm_macro(numbers):
    window = gw.getWindowsWithTitle('finished goods jobs_3m')[0]
    window.activate()
    count = 0
    for num in numbers:
        py.write(str(num[1]))
        py.press('down')
        count = count + 1
        if count == 7:
            count = 0
            break
    for num in numbers:
        if num[0] == 'oeday11' or num[0] == 'oeday5':
            py.write(str(num[1]))
            py.press('down')    
    py.press('down')
    for num in numbers:
        if num[0] == 'oeday6':
            py.write(str(num[1]))
            py.press('down')
            py.write(str(num[2]))
            py.press('down')
            py.write(str(num[3]))
            py.press('down')
    py.press('down')
    for num in numbers:
        if num[0] == 'oeday1':
            py.write(str(num[1]))   
    py.press('down')
#macro for one of the accounting programs, automates data entry
def pp_macro(numbers):
    window = gw.getWindowsWithTitle('finished goods jobs_pp')[0]
    window.activate()
    count = 0   
    for num in numbers:
        py.write(str(num[1]))
        py.press('down')
        count = count + 1        
        if count == 12:
            count = 0
            break       
    py.press('down')   
    for num in numbers:
        if num[0] == 'oeday6':
            py.write(str(num[1]))
            py.press('down')
            py.write(str(num[2]))
            py.press('down')
            py.write(str(num[3]))
            py.press('down')
    py.press('down')
    for num in numbers:
        if num[0] == 'oeday1_pp':
            py.write(str(num[1]))           
    py.press('down')
#macro for one of the accounting programs, automates user input
def get_control(beginning):
    window = gw.getWindowsWithTitle('Session A')[0]
    window.activate()
    py.write('ezview ocnhdrpf')
    py.press('enter')
    py.press('f5')
    py.press('down')
    py.write('v')
    for i in range(1,9):
        py.press('down')
    for i in range(0,3):
        py.write('v')
    py.press('enter')
    py.write(beginning)
    py.press('enter')
#macro for one of the billing programs, automates user input
def holdMacro(session,holdType):
    window = gw.getWindowsWithTitle(session)[0]
    window.activate()
    py.press('f6')
    py.write('h')
    py.press('enter')
    py.press('f3')
    py.press('f4')
    if holdType == 'General':
        py.press('down',presses=7)
    elif holdType == 'CC':
        py.press('pagedown')
        py.press('down',presses=8)
    elif holdType == 'Ship $':
        py.press('down',presses=8)
    py.write('1')
    py.press('enter')
    py.press('f6')
    py.press('down')
    py.write('1')
    py.press('enter')
    py.press('f3')

