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


#WritingExcelFile and Formatiing
def writeExcel (sampleNumber, asbestosType):
    asbestosTypeSeries = pd.Series(asbestosType).rename("asbestosType")
    # Creates Excel File to be written
    writer = ExcelWriter('testingdoc.xlsx')
    #writes nessecary information
    sampleNumber.to_excel(writer,'sheet1',index=False, index_label='sampleNumber', startcol=11)
    asbestosTypeSeries.to_excel(writer, 'sheet1',index=False, index_label='asbestosType', startcol= 18)
    #saves files
    writer.save()

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
        
        # gets entire columns as a series
        columnUniqueIdentifyer = sheet.iloc[0:,0]
        # Renames a column
        sampleNumber = columnUniqueIdentifyer.rename("sampleNumber")
        
        # Here I am taking the material score, seeing its value and splitting it into 4 columns depending on its values.
        colMatAssessment = sheet.iloc[0:,9]
        # Creating new lists
        productType = []
        condition = []
        surfaceTreatment = []
        asbestosType = []
        # Iterating through Material Scores and adding values to series.
        for num, score in colMatAssessment.iteritems():
            if(score == 1):
                productType.append(0)
                condition.append(0)
                surfaceTreatment.append(0)
                asbestosType.append(1)
            else:
                continue
        writeExcel(sampleNumber, asbestosType)
