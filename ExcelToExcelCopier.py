#Joshua Ryland for Maynard Marks
#Imports
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import PySimpleGUI as sg
import numpy as np

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
        
        # gets entire column as a series
        columnGet = sheet.iloc[0:,0]
        surveyID = columnGet.rename("surveyId")
        print(surveyID)
        writer = ExcelWriter('testingdoc.xlsx')
        surveyID.to_excel(writer,'sheet1',index=False, index_label='surveyID')
        writer.save()

