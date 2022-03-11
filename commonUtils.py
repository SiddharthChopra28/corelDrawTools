import os, datetime
import pandas as pd
import sys

if sys.platform == 'win32':
    LOG_FILE = os.path.join('C:\\Users\\', os.environ['USERNAME'], 'Desktop', 'printCodeGenLogs.txt')
else:
    LOG_FILE = os.path.join(os.environ['HOME'], 'Desktop', 'printCodeGenLogs.txt')



def log(msg):
    with open(LOG_FILE, 'a+') as f:
        log_str = f'{datetime.datetime.now()} - {msg}\n'
        f.write(log_str + '\n\n')

def read_excel_file(excel_path, column, sheet_name=1):
    try:
        sheet_name-=1
        # gui sheet name - 1
        df = pd.read_excel(excel_path, sheet_name=sheet_name)

        col_list = [chr(i+65) for i in range(len(df.columns))]

        df.columns = col_list
        #get the required column

        df_index = col_list.index(column)
        val_list = df[df.columns[df_index]].to_list()
        
    except Exception as e:
        log(str(e))
        return e

    return val_list


    