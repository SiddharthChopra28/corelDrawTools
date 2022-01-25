from datetime import datetime
import tkinter as tk
from tkinter import Label, Button
from tkinter import filedialog

import pyautogui
import math
import time
import pandas as pd

from threading import Thread

from win32 import win32gui
from win32.lib import win32con

def shorten_path(path):
    print(path)
    try:
        split_path = path.split('/')
        return split_path[0] + '/' + split_path[1] + '/...' + split_path[-1]
    except:
        return path

LOG_FILE = 'corelDrawPrinterLogs.txt'

def log(msg):
    with open(LOG_FILE, 'w') as f:
        log_str = f'{datetime.now()} - {msg}\n'
        f.write(log_str + '\n\n')
    

class CorelDrawPrinter(tk.Tk):
    def __init__(self):
        super().__init__()
        
        self.corelDrawWindowHWND = None
        self.excel_path = ''
        self.no_pages_var = tk.StringVar()
        
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
        
        instructions_label1 = Label(self, text="Keep open only one corelDraw window open, containing the file to print", font=('calibre',10, 'bold'), fg="red")
        instructions_label2 = Label(self, text="Arrange the pages in order of the excel file rows", font=('calibre',10, 'bold'), fg="red")
        instructions_label3 = Label(self, text="Choose an excel file containing printing instructions", font=('calibre',10, 'bold'))


        open_excel_file_button = Button(self, text = "Browse Files", command=self.openExcelFile)

        self.curr_file_path_label = Label(self, text=f"The currently chosen excel file is: {shorten_path(self.excel_path)}", fg="blue")

        no_items_per_page_label = tk.Label(self, text = 'Column containing number of sheets, eg. A', font=('calibre',10, 'bold'))

        no_items_per_page_entry = tk.Entry(self,textvariable = self.no_pages_var, font=('calibre',10,'normal'))

        start_btn = Button(self, text = 'Start printing', command = self.initiate_printing)

        instructions_label1.pack(side=tk.TOP, pady=(25,0))
        instructions_label2.pack(side=tk.TOP, pady=(15,0))
        instructions_label3.pack(side=tk.TOP, pady=(15,0))
        open_excel_file_button.pack(side=tk.TOP, pady=(8, 8))
        self.curr_file_path_label.pack(side=tk.TOP, pady=(0, 15))
        no_items_per_page_label.pack(side=tk.TOP)
        no_items_per_page_entry.pack(side=tk.TOP)

        start_btn.pack(side=tk.TOP, pady=(20,0))


                
        
    def openExcelFile(self):
        filename = filedialog.askopenfilename(initialdir="/", title="Choose excel file with details", filetypes=[("Excel Files", "*.xlsx")])
        self.excel_path = filename
        self.curr_file_path_label.configure(text=f"The currently chosen excel file is: {shorten_path(self.excel_path)}")

    def initiate_printing(self):
        if self.excel_path != '':
            self.no_pages_col = self.no_pages_var.get()
            if self.no_pages_col == '':
                log(text='Please enter the column containing number of pages')
                self.confirm_dialog(head='Error', text='Please enter the column containing number of pages', btnText=None)
                return None

            
            self.no_pages_col = self.no_pages_col.upper()

            self.instructions = self.read_excel_file()
            if not self.instructions:
                self.confirm_dialog(head='Error', text='Couldn\'t process excel file', btnText=None)
                return None

            Thread(target=lambda: self.start_printing(self.instructions), daemon=True).start()
            # self.start_printing(self.instructions)

            
        else:
            self.confirm_dialog(head='Error', text='Please choose an excel file to continue', btnText=None)
            return None
        

        

    def read_excel_file(self):
        try:
            df = pd.read_excel(self.excel_path, sheet_name=0)
            
        except FileNotFoundError as e:
            log(e)
            self.confirm_dialog(head='Error', text='The excel file you selected does not exist', btnText=None)
            return None 
        
        except Exception as e:
            log(e)
            self.confirm_dialog(head='Error', text=str(e), btnText=None)
            return None
        
        else:
            col_list = [chr(i+65) for i in range(len(df.columns))]

            df.columns = col_list

            try:
                df_index = col_list.index(self.no_pages_col)
                self.no_pages_list = df[df.columns[df_index]].to_list()
            except ValueError as e:
                log(e)
                self.confirm_dialog(head='Error', text='The quantity column entered is not valid', btnText=None)
                return None 

            except Exception as e:
                log(e)
                self.confirm_dialog(head='Error', text='There was some error in getting the quantity column', btnText=None)
                return None

            try:
                self.no_pages_list = list(map(lambda i: math.ceil(float(i)), self.no_pages_list))

                print(self.no_pages_list)
                return self.no_pages_list
            
            except Exception as e:
                log(e)
                self.confirm_dialog(head='Error', text='There was a non-numeric value in the quantity column', btnText=None)
                return None
 
    
        

    def confirm_dialog(self, head, text, btnText):
        DialogWindow(head=head, text=text, btnText=btnText, root=self)


    def maximize_crl_drw(self):
        self.total_crd_windows = []

        try:
            win32gui.EnumWindows(self.winEnumHandler, None)
            print(self.total_crd_windows)
            if len(self.total_crd_windows) == 0:
                self.confirm_dialog(head='Error', text='No CorelDraw window found', btnText=None)
                log('No CorelDraw window found')
                return None

            elif len(self.total_crd_windows) == 1:
                self.corelDrawWindowHWND = list(self.total_crd_windows[0].keys())[0]

                
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

    def start_printing(self, instructions):
        

        PrintingWindow(self)
            
            
    def winEnumHandler(self, hwnd, *_):
        
        try:
            if win32gui.IsWindowVisible(hwnd) and 'coreldraw x8' in win32gui.GetWindowText(hwnd).lower() and win32gui.GetWindowText(hwnd) != self.windowName:
                self.total_crd_windows.append({hwnd:win32gui.GetWindowText(hwnd)})
        except Exception as e:
            self.confirm_dialog(head='Error', text=str(e), btnText=None)
            log(e)
            return None
            
    

    
class DialogWindow(tk.Toplevel):
    
    def __init__(self, head, text, btnText=None, root=None):
        super().__init__(master=root)
        geometry = f"{400}x{200}+{int(root.centerx+root.width/3.5)}+{int(root.centery*1.5)}"

        self.title(head)
        self.geometry(geometry)
        self.resizable(False, False)

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


class PrintingWindow(tk.Toplevel):
    def __init__(self, root):
        super().__init__(master=root)
        geometry = f"{400}x{200}+{int(root.centerx+root.width/3.5)}+{int(root.centery*1.5)}"

        self.title("Printing in progress")
        self.geometry(geometry)
        self.resizable(False, False)

        self.sleep_time = 5
        
        self.root = root
        
        self.protocol('WM_DELETE_WINDOW', self.close)
        self.manual_print()
        

    def manual_print(self):
        self.root.withdraw()
        
        self.textLabel = Label(self, text="Please print first page manually.\nLeave the first page open on the screen\nPress Done after print command has been dispatched", font=('calibre',10, 'bold'))
        self.doneBtn = Button(self, text="Done", command= self.auto_print)

          
        self.textLabel.pack(side=tk.TOP, pady=(7,0))
        self.doneBtn.pack(side=tk.TOP, pady=(20, 0))
        

    def auto_print(self):
        self.textLabel.destroy()
        self.doneBtn.destroy()
        
        self.textLabel = Label(self, text="Automatic printing in progress\nDo not move the mouse manually untill printing is complete\nTo stop printing drag mouse cursor to any corner of screen", font=('calibre',10, 'bold'))

        self.textLabel.pack(side=tk.TOP, pady=(7,0))

        
        try:
            self.no_pages_list = self.root.no_pages_list
            self.no_pages_list.pop(0)
            for i in self.no_pages_list:
                self._print(i)

        except pyautogui.FailSafeException as e:
            log(e)
            self.close()

        except Exception as e:
            log(e)
            self.close()


    def click(self, filename, retNone=False):
        self.root.maximize_crl_drw()
        
        time.sleep(2)
        
        pos = pyautogui.locateCenterOnScreen(f'C:/Users/Siddharth/Desktop/images/{filename}.png')
    
        if not retNone and pos is None:
            self.destroy()
            self.root.deiconify()
            self.root.confirm_dialog(head='Error', text='Couldn\'t process corelDraw window', btnText=None)
            log('Image not found on screen')
            return None
        
        if retNone and pos is None:
            return "IMG NOT FOUND"



        pyautogui.moveTo(pos)
        pyautogui.click()
        


    def _print(self, page):
        
        next_btn = self.click("next_btn", retNone=True)
        
        if next_btn == "IMG NOT FOUND":
            
            self.destroy()
            self.root.deiconify()
            self.root.confirm_dialog(head='Finsihed Printing', text='Reached end of document, no more pages were found', btnText=None)
            return None

        time.sleep(self.sleep_time)
        
        self.click("file")
        time.sleep(self.sleep_time)
        self.click("print")
        time.sleep(self.sleep_time)
        self.click("no_copies", retNone=True)
        pyautogui.press('backspace', presses=5, interval=0.5)
        page = str(page)
        for i in page:    
            pyautogui.press(i)
            time.sleep(0.5)
        
        time.sleep(self.sleep_time)
        self.click("print_final", retNone=True)
        time.sleep(self.sleep_time+6)

        

    def close(self):
        self.destroy()
        self.root.deiconify()
        self.root.confirm_dialog("Printing Cancelled", "Current printing task has been cancelled", None)


if __name__ == '__main__':
    CorelDrawPrinter().mainloop()

#take ss of all buttons to be clicked and then click them

# pos = pyautogui.locateCenterOnScreen('images/test_img.png') #If the file is not a png file it will not work

# pyautogui.moveTo(pos)#Moves the mouse to the coordinates of the image
