import pyautogui as gui
from PIL import Image
import time
import pandas as pd

def PathFile(path, file):
    location = '{0}/{1}'.format(str(path),str(file))
    return location

SOURCE_PATH = './resources'
SOURCE_FILE = 'input_automation.xlsx'
SOURCE = PathFile(SOURCE_PATH,SOURCE_FILE)
REGION = (450,120,880,800)

df = pd.read_excel(SOURCE)
ColumnNames = df.columns

variable_input_location = {}
variable_input_location['cross'] = Image.open(PathFile(SOURCE_PATH,'cross.PNG'))
variable_input_location['submit'] = Image.open(PathFile(SOURCE_PATH,'submit.PNG'))
variable_input_location['start'] = Image.open(PathFile(SOURCE_PATH,'start.PNG'))
for column in df.columns:
    try:
        variable_input_location[column] = Image.open(PathFile(SOURCE_PATH,str.strip(column))+'.PNG')
    except:
        print('error locating input image files')

def writeNtype(row, column, confidence=0.92):
    column_location = gui.locateCenterOnScreen(variable_input_location[column], region=REGION, confidence=confidence)
    gui.click(column_location[0],column_location[1]+25)
    gui.write(str(row[column]))
    return

if __name__ == "__main__":
    time.sleep(5)
    gui.click(gui.locateCenterOnScreen(variable_input_location['start'], confidence=0.8))
    time.sleep(0.1)
    for n, row in df.iterrows():
        for column in df.columns:    
            CONFIDENCE = 0.9    
            try:    
                if column == 'Phone Number':
                    CONFIDENCE = 0.78
                writeNtype(row, column, CONFIDENCE)
            except:
                print(column,'> Error')
                CONFIDENCE = 0.82    
                writeNtype(row,column,CONFIDENCE)          
            
        submit_location = gui.locateCenterOnScreen(variable_input_location['submit'], region=REGION, confidence=0.8)
        gui.click(submit_location[0],submit_location[1])