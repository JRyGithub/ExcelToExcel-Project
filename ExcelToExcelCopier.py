#Joshua Ryland for Maynard Marks
#Imports
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import PySimpleGUI as sg

#Following function is for obtaining any previous information from ending document
def ending_document():
    EndingDoc = pd.read_excel("C:/Users/joshua.ryland/Desktop/Magic Stuff/AsbestosItems.xlsx")
    return EndingDoc

#Set up GUI
sg.SetOptions(
background_color='#FFCD11',
element_background_color='#FFD948',
text_color='#5C4B0C',
button_color=('white','#98834B'))

layout = [[sg.Text('Filename:')],
        [sg.Input(background_color='#FFD948'), sg.FileBrowse()],
        [sg.Submit(), sg.Cancel()],
        [sg.Output(size=(100,50),background_color='#FFD948')]]

window = sg.Window("Reinspection Wizard",layout)

#Gui functionality
while True:
    event,values = window.Read()
    if event is None or event == "Cancel":
        break
    elif event is event == "Submit":
        pathname = values[0]
        print(pathname +"\n")
        endingDoc = ending_document()
        #checks if pathname contains AMP, as that would mean a different register has been added which is on a different sheet
        if("AMP" in pathname):
            try:
                sheet = pd.read_excel(pathname, 'App C - Asb Reg - Updated')
            except:
                try:
                    sheet = pd.read_excel(pathname)
                except:
                    print("Pathname error.")
                    continue
        else:
            sheet = pd.read_excel(pathname)
        print(sheet.shape)
        print(sheet.columns)
        mask = sheet.iloc[0:,0]
        print(mask)