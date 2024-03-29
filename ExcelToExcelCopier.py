#Joshua Ryland for Maynard Marks
#Imports
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import PySimpleGUI as sg
import numpy as np
import re
import math

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

def blankNamedSeriesMaker(name):
    nameList = []
    nameList = pd.Series().rename(name)
    return nameList

#WritingExcelFile and Formatiing
def writeExcel (sampleNumber, asbestosType, productType, condition, surfaceTreatment, extents, unitOM, identification, notes, materialDesc, surveyId, datesList, surveyorList, buildingName, floor, locationList,
        locationDescription, items, materialCode, approachList, actionDatesList, normalOccupancyPA, locationPA, accessibilityPA, amountPA, noOfPeoplePA, usePA, averageTimePA,
        maintenanceTypePA,frequencyPA):
    asbestosTypeSeries = pd.Series(asbestosType).rename("asbestosType")
    productTypeSeries = pd.Series(productType).rename("productType")
    conditionSeries = pd.Series(condition).rename("condition")
    surfaceTreatmentSeries = pd.Series(surfaceTreatment).rename("surfaceTreatment")
    extentsSeries = pd.Series(extents).rename("extent")
    unitOMSeries = pd.Series(unitOM).rename("UoM")
    surveyId = pd.Series(surveyId).rename("surveyId")
    datesList = pd.Series(datesList).rename("date")
    surveyorList = pd.Series(surveyorList).rename("surveyor")
    floor = pd.Series(floor).rename("floor")
    location = pd.Series(locationList).rename("location")
    # items = pd.Series(items).rename("item")
    materialDesc = pd.Series(materialDesc).rename("materialDesc")
    materialCode = pd.Series(materialCode).rename("materialCode")
    sampleNotes = []
    sampleNotes = pd.Series().rename("sampleNotes")
    noAccess = blankNamedSeriesMaker("noAccess")
    externalRef = blankNamedSeriesMaker("externalRef")
    recommendedAction = blankNamedSeriesMaker("recommendedAction")
    photofile1 = blankNamedSeriesMaker("photofile1")
    photofile2 =  blankNamedSeriesMaker("photofile2")
    default_pa_id = blankNamedSeriesMaker("default_pa_id")
    actionDatesList = pd.Series(actionDatesList).rename("actionDate")
    normalOccupancyPA = pd.Series(normalOccupancyPA).rename("normalOccupancyPA")
    locationPA = pd.Series(locationPA).rename("locationPA")
    accessibilityPA = pd.Series(accessibilityPA).rename("accessibilityPA")
    amountPA = pd.Series(amountPA).rename("amountPA")
    noOfPeoplePA = pd.Series(noOfPeoplePA).rename("noOfPeoplePA")
    usePA = pd.Series(usePA).rename("usePA")
    averageTimePA = pd.Series(averageTimePA).rename("averageTimePA")
    maintenanceTypePA = pd.Series(maintenanceTypePA).rename("maintenanceTypePA")
    frequencyPA = pd.Series(frequencyPA).rename("frequencyPA")
    approach = pd.Series(approachList).rename("approach")
        
    # Creates Excel File to be written
    writer = ExcelWriter('testingdoc.xlsx')
    #writes nessecary information
    surveyId.to_excel(writer,'sheet1',index=False, index_label='suveyId', startcol=0)
    datesList.to_excel(writer,'sheet1',index=False, index_label='date', startcol=1)
    surveyorList.to_excel(writer,'sheet1',index=False, index_label='surveyor', startcol=2)
    buildingName.to_excel(writer,'sheet1',index=False, index_label='building', startcol=3)
    floor.to_excel(writer,'sheet1',index=False, index_label='floor', startcol=4)
    location.to_excel(writer,'sheet1',index=False, index_label='location', startcol=5)
    locationDescription.to_excel(writer,'sheet1',index=False, index_label='locationDescription', startcol=6)
    items.to_excel(writer,'sheet1',index=False, index_label='item', startcol=7)
    materialCode.to_excel(writer,'sheet1',index=False, index_label='materialCode', startcol =8)
    materialDesc.to_excel(writer,'sheet1',index=False, index_label='materialDesc', startcol=9)
    approach.to_excel(writer,'sheet1',index=False, index_label='approach', startcol=10)
    sampleNumber.to_excel(writer,'sheet1',index=False, index_label='sampleNumber', startcol=11)
    sampleNotes.to_excel(writer,'sheet1',index=False, index_label='sampleNote', startcol=12)
    extentsSeries.to_excel(writer, 'sheet1',index=False, index_label='extent', startcol= 13)
    unitOMSeries.to_excel(writer, 'sheet1',index=False, index_label='UoM', startcol= 14)
    productTypeSeries.to_excel(writer, 'sheet1',index=False, index_label='productType', startcol= 15)
    conditionSeries.to_excel(writer, 'sheet1',index=False, index_label='condition', startcol= 16)
    surfaceTreatmentSeries.to_excel(writer, 'sheet1',index=False, index_label='surfaceTreatment', startcol= 17)
    asbestosTypeSeries.to_excel(writer, 'sheet1',index=False, index_label='asbestosType', startcol= 18)
    identification.to_excel(writer, 'sheet1',index=False, index_label='identification', startcol= 19)
    recommendedAction.to_excel(writer, 'sheet1',index=False, index_label='recommendedAction', startcol= 20)
    noAccess.to_excel(writer, 'sheet1',index=False, index_label='noAccess', startcol= 21)
    externalRef.to_excel(writer, 'sheet1',index=False, index_label='externalRef', startcol= 22)
    notes.to_excel(writer, 'sheet1',index=False, index_label='notes', startcol= 23)
    photofile1.to_excel(writer, 'sheet1',index=False, index_label='photofile1', startcol= 24)
    photofile2.to_excel(writer, 'sheet1',index=False, index_label='photofile2', startcol= 25)
    actionDatesList.to_excel(writer, 'sheet1',index=False, index_label='actionDate', startcol= 26)
    default_pa_id.to_excel(writer, 'sheet1',index=False, index_label='default_pa_id', startcol= 27)
    normalOccupancyPA.to_excel(writer, 'sheet1',index=False, index_label='normalOccupancyPA', startcol= 28)
    locationPA.to_excel(writer, 'sheet1',index=False, index_label='locationPA', startcol= 29)
    accessibilityPA.to_excel(writer, 'sheet1',index=False, index_label='accessibilityPA', startcol= 30)
    amountPA.to_excel(writer, 'sheet1',index=False, index_label='amountPA', startcol= 31)
    noOfPeoplePA.to_excel(writer, 'sheet1',index=False, index_label='noOfPeoplePA', startcol= 32)
    usePA.to_excel(writer, 'sheet1',index=False, index_label='usePA', startcol= 33)
    averageTimePA.to_excel(writer, 'sheet1',index=False, index_label='averageTimePA', startcol= 34)
    maintenanceTypePA.to_excel(writer, 'sheet1',index=False, index_label='maintenanceTypePA', startcol= 35)
    frequencyPA.to_excel(writer, 'sheet1',index=False, index_label='frequencyPA', startcol= 36)
    #Adds the ramining colums which have the same name but with a incremental number
    counter = 2
    listCount = 0
    writeCounter = 37
    coolList = []
    while(counter <= 9):
        tempList = []
        item = "item"+str(counter)
        material = "material"+str(counter)
        item = blankNamedSeriesMaker(item)
        material = blankNamedSeriesMaker(material)
        tempList.append(item)
        tempList.append(material)
        coolList.append(tempList)
        counter = counter + 1
        coolList[listCount][0].to_excel(writer, 'sheet1',index=False, index_label='frequencyPA', startcol= writeCounter )
        writeCounter = writeCounter+1
        coolList[listCount][1].to_excel(writer, 'sheet1',index=False, index_label='frequencyPA', startcol= writeCounter )
        writeCounter = writeCounter+1
        listCount = listCount +1
    #saves files
    writer.save()
    print("Writing Successfull.")
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
def priorScores(number):
    if(number == 1):
        normalOccupancyPA.append(1)
        locationPA.append(0)
        accessibilityPA.append(0)
        amountPA.append(0)
        noOfPeoplePA.append(0)
        usePA.append(0)
        averageTimePA.append(0)
        maintenanceTypePA.append(0)
        frequencyPA.append(0)
    elif(number == 2):
        normalOccupancyPA.append(1)
        locationPA.append(1)
        accessibilityPA.append(0)
        amountPA.append(0)
        noOfPeoplePA.append(0)
        usePA.append(0)
        averageTimePA.append(0)
        maintenanceTypePA.append(0)
        frequencyPA.append(0)
    elif(number == 3):
        normalOccupancyPA.append(1)
        locationPA.append(1)
        accessibilityPA.append(1)
        amountPA.append(0)
        noOfPeoplePA.append(0)
        usePA.append(0)
        averageTimePA.append(0)
        maintenanceTypePA.append(0)
        frequencyPA.append(0)
    elif(number == 4):
        normalOccupancyPA.append(1)
        locationPA.append(1)
        accessibilityPA.append(1)
        amountPA.append(1)
        noOfPeoplePA.append(0)
        usePA.append(0)
        averageTimePA.append(0)
        maintenanceTypePA.append(0)
        frequencyPA.append(0)
    elif(number == 5):
        normalOccupancyPA.append(1)
        locationPA.append(1)
        accessibilityPA.append(1)
        amountPA.append(1)
        noOfPeoplePA.append(1)
        usePA.append(0)
        averageTimePA.append(0)
        maintenanceTypePA.append(0)
        frequencyPA.append(0)
    elif(number == 6):
        normalOccupancyPA.append(1)
        locationPA.append(1)
        accessibilityPA.append(1)
        amountPA.append(1)
        noOfPeoplePA.append(1)
        usePA.append(1)
        averageTimePA.append(0)
        maintenanceTypePA.append(0)
        frequencyPA.append(0)
    elif(number == 7):
        normalOccupancyPA.append(1)
        locationPA.append(1)
        accessibilityPA.append(1)
        amountPA.append(1)
        noOfPeoplePA.append(1)
        usePA.append(1)
        averageTimePA.append(1)
        maintenanceTypePA.append(0)
        frequencyPA.append(0)
    elif(number == 8):
        normalOccupancyPA.append(1)
        locationPA.append(1)
        accessibilityPA.append(1)
        amountPA.append(1)
        noOfPeoplePA.append(1)
        usePA.append(1)
        averageTimePA.append(1)
        maintenanceTypePA.append(1)
        frequencyPA.append(0)
    elif(number == 9):
        normalOccupancyPA.append(1)
        locationPA.append(1)
        accessibilityPA.append(1)
        amountPA.append(1)
        noOfPeoplePA.append(1)
        usePA.append(1)
        averageTimePA.append(1)
        maintenanceTypePA.append(1)
        frequencyPA.append(1)
    elif(number == 10):
        normalOccupancyPA.append(2)
        locationPA.append(1)
        accessibilityPA.append(1)
        amountPA.append(1)
        noOfPeoplePA.append(1)
        usePA.append(1)
        averageTimePA.append(1)
        maintenanceTypePA.append(1)
        frequencyPA.append(1)
    elif(number == 11):
        normalOccupancyPA.append(2)
        locationPA.append(2)
        accessibilityPA.append(1)
        amountPA.append(1)
        noOfPeoplePA.append(1)
        usePA.append(1)
        averageTimePA.append(1)
        maintenanceTypePA.append(1)
        frequencyPA.append(1)
    elif(number == 12):
        normalOccupancyPA.append(2)
        locationPA.append(2)
        accessibilityPA.append(2)
        amountPA.append(1)
        noOfPeoplePA.append(1)
        usePA.append(1)
        averageTimePA.append(1)
        maintenanceTypePA.append(1)
        frequencyPA.append(1)
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
            score = str(score)
            if(score == "nan") or (score == "-") or ("N/A" in score):
                productType.append("")
                condition.append("")
                surfaceTreatment.append("")
                asbestosType.append("")
            else:
                scores(int(float(score)))
        # Extent Split into two colums of the format int and string
        extentLock = sheet.iloc[0:,7]
        unitOM = []
        extents = []

        for num, extent in extentLock.iteritems():
            testString = str(type(extent))
            if(testString == '<class \'float\'>'):
                unitOM.append('')
                extents.append('')
            elif(testString == 'N/A '):
                unitOM.append('')
                extents.append('')
            else:
                extentSlice(extent)
        #Asbestos Presence to identity
        asbestosPresence = sheet.iloc[0:,6]
        identification = asbestosPresence.rename("identification") 

        #Observations and Recommendations to notes
        obsAndRec = sheet.iloc[0:,13]
        notes = obsAndRec.rename("notes")


        #surveyId
        surveyId = []
        rangeId = len(columnUniqueIdentifyer)
        for x in range(0,rangeId):
            surveyId.append('')
        
        #date
        datesList = []
        check = 0
        for num, dates in columnUniqueIdentifyer.iteritems():
            try:
                if(len(dates) > 5):
                    date = dates[0:6]
                    year, month, day = date[:2],date[2:4],date[4:]
                    reformattedDate = str(day)+"/"+str(month)+"/20"+str(year)
                    datesList.append(reformattedDate)
            except:
                check = 1
                break
        if(check == 1):
            datesList = []
            inspectDate = sheet.iloc[0:,16]
        # inspectDate = inspectDate.rename("Date")
            for num, inspects in inspectDate.iteritems():
                testString = str(type(inspects))
                if(testString == '<class \'pandas._libs.tslibs.nattype.NaTType\'>'):
                    datesList.append("")
                else:
                    inspectDates = str(inspects)
                    newDate = inspectDates[8:10]+"/"+inspectDates[5:7]+"/"+inspectDates[:4]
                    datesList.append(newDate)

        #surveyor
        surveyorList = []
        for num, surveyor in columnUniqueIdentifyer.iteritems():
            try:
                surveyorName = ''.join(i for i in surveyor if not i.isdigit())
                if ("SM" in surveyorName) or ("BMI" in surveyorName):
                    surveyorList.append("Samisoni Manu")
                elif "RM" in surveyorName:
                    surveyorList.append("Robert McAllister")
                elif "LE" in surveyorName:
                    surveyorList.append("Luke England")
                elif "DC" in surveyorName:
                    surveyorList.append("Darren Carter")
                elif ("CM" in surveyorName) or ("CP" in surveyorName):
                    surveyorList.append("Chloe Parkins")
                elif ("TD" in surveyorName) or ("DAV" in surveyorName):
                    surveyorList.append("Tony Davison")
                elif ("BAL" in surveyorName) or ("AB" in surveyorList) or ("HP" in surveyorList) or ("SARAB" in surveyorList):
                    surveyorList.append("Andrew Ball")
                elif ("JR" in surveyorName):
                    surveyorList.append("Joshua Ryland")
                elif ("KD" in surveyorName):
                    surveyorList.append("Kurt Downie")
                elif ("SC" in surveyorName):
                    surveyorList.append("Simon Cunliffe")
                elif ("SP" in surveyorName):
                    surveyorList.append("Simon Paykel")
                elif ("RA" in surveyorName):
                    surveyorList.append("Richard Angel")
                else:
                    surveyorList.append(surveyorName)
            except:
                surveyorList.append("Unknown")

        #BuildingName
        propertyName = sheet.iloc[0:,2]
        buildingName = propertyName.rename('building')
        
        #floor
        floor = []
        rangeId = len(columnUniqueIdentifyer)
        for x in range(0,rangeId):
            floor.append('')
        #location
        locationList = []
        rangeId = len(columnUniqueIdentifyer)
        for x in range(0,rangeId):
            locationList.append('')
        
        #location  description
        locationOfSample = sheet.iloc[0:,3]
        locationDescription = locationOfSample.rename("locationDescription")

        
        #actionDate
        reinspectDate = sheet.iloc[0:,15]
        actionDate = reinspectDate.rename("actionDate")
        actionDatesList = []
        for num, actionDates in actionDate.iteritems():
            testString = str(type(actionDates))
            if(testString == '<class \'pandas._libs.tslibs.nattype.NaTType\'>'):
                actionDatesList.append("")
            else:
                actionDates = str(actionDates)
                newDate = actionDates[8:10]+"/"+actionDates[5:7]+"/"+actionDates[:4]
                actionDatesList.append(newDate)

        #priorityScore
        colPriorAssessment = sheet.iloc[:,10]
        normalOccupancyPA = []
        locationPA = []
        accessibilityPA = []
        amountPA = []
        noOfPeoplePA = []
        usePA = []
        averageTimePA = []
        maintenanceTypePA = []
        frequencyPA = []
        for num, score in colPriorAssessment.iteritems():
            score = str(score)
            if(score == "nan") or (score == "-") or ("N/A" in score):
                normalOccupancyPA.append("")
                locationPA.append("")
                accessibilityPA.append("")
                amountPA.append("")
                noOfPeoplePA.append("")
                usePA.append("")
                averageTimePA.append("")
                maintenanceTypePA.append("")
                frequencyPA.append("")
            else:
                priorScores(int(float(score)))
        #approach
        approach = sheet.iloc[0:,1]
        approachList = []
        for num, approaches in approach.iteritems():
            if(approaches == "SS"):
                approachList.append("S")
            else:
                approachList.append("PS")
        
       #accesses material code sheet and reads it in, creating a dictionary series, which needs to be merged into one large dictionary
        materialSheet = pd.read_excel("K:/Resources/Technical Library/20. ASBESTOS/7.AlphaTracker/Materials.xlsx")
        materialSheet['merged'] = materialSheet.apply(lambda row: {row['Material code']:row['Material description']}, axis=1)
        matty = materialSheet.loc[:,'merged']
        
        #Sample Catergory to items, items creates a materialCode which in turn creates material description
        samCat =  sheet.iloc[0:,5]
        materialCode = []
        materialDict = {}
        materialDictTwo = {}
        materialDesc = []
        for number1, mattyDict in matty.iteritems():
            for key, value in mattyDict.items():
                materialDict[value.lower()] = key
                materialDictTwo[key] = value
        for number, samMat in samCat.iteritems():
            try:
                materialCode.append(materialDict[samMat.lower()])
            except:
                materialCode.append('')
        for code in materialCode:
            try:
                materialDesc.append(materialDictTwo[code])
            except:
                materialDesc.append('')

        items = samCat.rename("item")
        
        #calls the write excel file, to begin formatting and writing the file passed all values worked put previously
        writeExcel(sampleNumber, asbestosType, productType, condition, surfaceTreatment, extents, unitOM, identification,
        notes, materialDesc, surveyId, datesList, surveyorList, buildingName, floor, locationList,
        locationDescription, items, materialCode, approachList, actionDatesList, normalOccupancyPA, locationPA, accessibilityPA, amountPA, noOfPeoplePA, usePA, averageTimePA,
        maintenanceTypePA,frequencyPA)

        
