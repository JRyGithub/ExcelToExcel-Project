#Joshua Ryland for Maynard Marks
#Imports
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import PySimpleGUI as sg
import numpy as np
import re

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
def writeExcel (sampleNumber, asbestosType, productType, condition, surfaceTreatment):
    asbestosTypeSeries = pd.Series(asbestosType).rename("asbestosType")
    productTypeSeries = pd.Series(productType).rename("productType")
    conditionSeries = pd.Series(condition).rename("condition")
    surfaceTreatmentSeries = pd.Series(surfaceTreatment).rename("surfaceTreatment")
    # Creates Excel File to be written
    writer = ExcelWriter('testingdoc.xlsx')
    #writes nessecary information
    sampleNumber.to_excel(writer,'sheet1',index=False, index_label='sampleNumber', startcol=11)
    productTypeSeries.to_excel(writer, 'sheet1',index=False, index_label='productType', startcol= 15)
    conditionSeries.to_excel(writer, 'sheet1',index=False, index_label='condition', startcol= 16)
    surfaceTreatmentSeries.to_excel(writer, 'sheet1',index=False, index_label='surfaceTreatment', startcol= 17)
    asbestosTypeSeries.to_excel(writer, 'sheet1',index=False, index_label='asbestosType', startcol= 18)
    
    #saves files
    writer.save()

def scores (number):
    if(number == 1):
        productType.append(0)
        condition.append(0)
        surfaceTreatment.append(0)
        asbestosType.append(number)
    elif(number == 2):
        productType.append(0)
        condition.append(0)
        surfaceTreatment.append(0)
        asbestosType.append(number)
    elif(number == 3):
        productType.append(0)
        condition.append(0)
        surfaceTreatment.append(0)
        asbestosType.append(number)
    elif(number == 4):
        productType.append(1)
        condition.append(0)
        surfaceTreatment.append(0)
        asbestosType.append(3)
    elif(number == 5):
        productType.append(1)
        condition.append(1)
        surfaceTreatment.append(0)
        asbestosType.append(3)
    elif(number == 6):
        productType.append(1)
        condition.append(1)
        surfaceTreatment.append(1)
        asbestosType.append(3)
    elif(number == 7):
        productType.append(2)
        condition.append(1)
        surfaceTreatment.append(1)
        asbestosType.append(3)
    elif(number == 8):
        productType.append(2)
        condition.append(2)
        surfaceTreatment.append(1)
        asbestosType.append(3)
    elif(number == 9):
        productType.append(2)
        condition.append(2)
        surfaceTreatment.append(2)
        asbestosType.append(3)
    elif(number == 10):
        productType.append(3)
        condition.append(2)
        surfaceTreatment.append(2)
        asbestosType.append(3)
    elif(number == 11):
        productType.append(3)
        condition.append(3)
        surfaceTreatment.append(2)
        asbestosType.append(3)
    elif(number == 12):
        productType.append(3)
        condition.append(3)
        surfaceTreatment.append(3)
        asbestosType.append(3)
def extentSlice(extent):
    extentStr = str(extent)
    if(extentStr[0] == '>') or (extentStr[0] == '<'):
        unitOM.append(extentStr[-2:])
        extents.append(extentStr[1:-2])
    else:
        unitOM.append(extentStr[-2:])
        extents.append(extentStr[0:-2])

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
            if(type(score) == type(1)):
                scores(score)
            else:
                continue
        # Extent Split into two colums of the format int and string
        extentLock = sheet.iloc[0:,7]
        unitOM = []
        extents = []

        for num, extent in extentLock.iteritems():
            if (extent == 'N/A'):
                continue
            else:
                extentSlice(extent)
        print(unitOM)
        print(extents)
        #calls the write excel file, to begin formatting and writing the file passed all values worked put previously
        # writeExcel(sampleNumber, asbestosType, productType, condition, surfaceTreatment)

        
