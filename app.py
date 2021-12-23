from datetime import datetime
import tkinter as tk
from tkinter import Label, Button
from tkinter.constants import TOP
from tkinter import filedialog

import pyautogui

import pandas as pd

from win32 import win32gui
from win32.lib import win32con

def shorten_path(path):
    print(path)
    try:
        split_path = path.split('/')
        return split_path[0] + '/' + split_path[1] + '/...' + split_path[-1]
    except:
        return path

LOG_FILE = 'logs'

def log(msg):
    with open(LOG_FILE, 'w') as f:
        log_str = f'{datetime.now()} - {msg}\n'
        f.write(log_str + '\n\n')
    

class CorelDrawPrinter(tk.Tk):
    def __init__(self):
        super().__init__()
        
        self.corelDrawWindowHWND = None
        self.excel_path = ''
        
        self.windowName = "CorelDraw printer"
        self.setup()
        self.home_screen()

    def setup(self):
        self.title(self.windowName)

        self.screen_width, self.screen_height = self.winfo_screenwidth(), self.winfo_screenheight()
        self.width, self.height = 904, 500

        self.centerx, self.centery = int(
        self.screen_width/2 - self.width/2), int(self.screen_height/2.5 - self.height/2)

        self.geometry(
            f"{self.width}x{self.height}+{self.centerx}+{self.centery}")
        self.resizable(False, False)
        self.title(self.windowName)
        
        
    
    def home_screen(self):
        instructions_label = Label(self, text="Choose an excel file containing printing instructions", width=100, height=4, pady=1, fg="red")
        

        open_excel_file_button = Button(self, text = "Browse Files", command=self.openExcelFile)

        self.curr_file_path_label = Label(self, text=f"The currently chosen excel file is: {shorten_path(self.excel_path)}", width=100, height=4, fg="blue")
        
        start_btn = Button(self, text = 'Start printing', command = self.initiate_printing)
        
        instructions_label.pack(side=tk.TOP)
        open_excel_file_button.pack(side=tk.TOP)
        self.curr_file_path_label.pack(side=tk.TOP)
        start_btn.pack(side=tk.TOP)
        
        
    def openExcelFile(self):
        filename = filedialog.askopenfilename(initialdir="/", title="Choose excel file with details", filetypes=[("Excel Files", "*.xlsx")])
        self.excel_path = filename
        self.curr_file_path_label.configure(text=f"The currently chosen excel file is: {shorten_path(self.excel_path)}")

    def initiate_printing(self):
        if self.excel_path != '':
            instructions = self.read_excel_file()
            if not instructions:
                self.confirm_dialog(head='Error', text='Couldn\'t process excel file', btnText=None)
                return None
            
            self.maximize_crl_drw()
            #self.click('test_img')

            
        else:
            self.confirm_dialog(head='Error', text='Please choose an excel file to continue', btnText=None)
            return None
        
    def click(self, filename):
        pos = pyautogui.locateCenterOnScreen(f'images/{filename}.png')
    
        if pos is None:
            self.confirm_dialog(head='Error', text='Couldn\'t process corelDraw window', btnText=None)
            log('Image not found on screen')
            return None


        pyautogui.moveTo(pos)
        pyautogui.click()


    def read_excel_file(self):
        try:
            df = pd.read_excel(self.excel_path)
            
        except FileNotFoundError as e:
            log(e)
            self.confirm_dialog(head='Error', text='The excel file you selected does not exist', btnText=None)
            return None 
        
        except Exception as e:
            log(e)
            self.confirm_dialog(head='Error', text=str(e), btnText=None)
            return None
        
        else:  
            return "hello"
    
        

    def confirm_dialog(self, head, text, btnText):
        NewWindow(head=head, text=text, btnText=btnText, root=self)


    def maximize_crl_drw(self):
        self.total_crd_windows = []

        try:
            win32gui.EnumWindows(self.winEnumHandler, None)
                        
            if len(self.total_crd_windows) == 0:
                self.confirm_dialog(head='Error', text='No CorelDraw window found', btnText=None)
                log('No CorelDraw window found')
                return None

            elif len(self.total_crd_windows) == 1:
                self.corelDrawWindowHWND = list(self.total_crd_windows[0].keys())[0]
                print(self.total_crd_windows)
                
            else:
                self.confirm_dialog(head='Error', text='More than one CorelDraw window found.\nPlease keep only the windows you want to work with open.', btnText=None)
                log('More than one CorelDraw window found') 
                return None

            win32gui.ShowWindow(self.corelDrawWindowHWND, win32con.SW_MAXIMIZE)
            win32gui.SetForegroundWindow(self.corelDrawWindowHWND)

        except Exception as e:
            self.confirm_dialog(head='Error', text=f'{e}', btnText=None)
            log(e)
            return None
            
            
    def winEnumHandler(self, hwnd, *_):
        
        try:
            if win32gui.IsWindowVisible(hwnd) and 'coreldraw' in win32gui.GetWindowText(hwnd).lower() and win32gui.GetWindowText(hwnd) != self.windowName:
                self.total_crd_windows.append({hwnd:win32gui.GetWindowText(hwnd)})
        except Exception as e:
            self.confirm_dialog(head='Error', text=str(e), btnText=None)
            log(e)
            return None
            
    

    
class NewWindow(tk.Toplevel):
    
    def __init__(self, head, text, btnText=None, root=None):
        super().__init__(master=root)
        geometry = f"{400}x{200}+{int(root.centerx+root.width/3.5)}+{int(root.centery*1.5)}"

        self.title(head)
        self.geometry(geometry)
        self.resizable(False, False)
        # self.iconbitmap(r'resources/icon.ico')
        self.text = text
        self.btnText = btnText
        
        self.root = root
        
        self.mainScreen()
        
    def mainScreen(self):
        self.disableRoot()
        
        textLabel = Label(self, text=self.text, width=100, height=4, fg="blue")
        textLabel.pack(side=tk.TOP)
        
        if self.btnText:
            confirmButton = Button(self, text=self.btnText)
        else:
            confirmButton = Button(self, text='Ok', command=self.enableRoot)
            
        confirmButton.pack(side=tk.TOP)

    def disableRoot(self):
        self.grab_set()
        
    def enableRoot(self):
        self.grab_release()
        self.destroy()






if __name__ == '__main__':
    CorelDrawPrinter().mainloop()

#take ss of all buttons to be clicked and then click them

# pos = pyautogui.locateCenterOnScreen('images/test_img.png') #If the file is not a png file it will not work

# pyautogui.moveTo(pos)#Moves the mouse to the coordinates of the image
