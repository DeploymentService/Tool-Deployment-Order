#! python3
from __future__ import print_function

# System imports
import sys
import os
import os.path
from Autodesk.Revit.DB.Architecture import BuildingPad, Room
currentPath = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, currentPath)
sys.path.insert(0, "{}\libraries".format(currentPath))
sys.path.insert(0, os.getenv("PYTHONPATH"))

# Getting all path variables
pathVariables = os.getenv("path").split(";")
for variable in pathVariables:
    sys.path.insert(1, variable)

# Python imports
import clr
import re
import math
import json
import csv
from string import capwords
from collections import OrderedDict
import System

# The __init__ file with all the initialization procedures is __init__.py in the 'libraries' folder
from libraries import gSheets, gDrive, uploadImagesToDrive, getExportedCSVAsList, formatHeader, getFileInDriveByQuery, getParameterValueByName, consolidateImages, countAndCollapseRepeatedItems, CategoryRowGroup, CBCTemplate, convertUnits, assignUnitType, getTotalCostsPerItem, getTotalQuantitiesPerItem

# Revit API imports
clr.AddReference("System.Drawing")
import System.Drawing
clr.AddReference("RevitAPI")
clr.AddReference('RevitAPIUI')
clr.AddReference("RevitNodes")
import Autodesk
import Autodesk.Revit
from Autodesk.Revit.UI import *
import Autodesk.Revit.UI.Selection
from Autodesk.Revit.DB import * 
from Autodesk.Revit.DB import Transaction, Structure, Architecture

# Third Party Imports
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.http import MediaFileUpload
from googleapiclient.errors import HttpError

# Accessing current Revit Document
doc = __revit__.ActiveUIDocument.Document

def costEstimationApp(sheetName):

    # RPW try---------------------------------------
    # # Setting and running the UI to set the sheet name to be exported
    # UIComponents = [Label("Specify the spreadsheet name"),
    #                 TextBox("spreadSheetName"),
    #                 Button("Confirm")]

    # form = FlexForm("Cost Estimation App", UIComponents)
    # form.show()

    # sheetName = form.values["spreadSheetName"]
    # ----------------------------------------------

    # Calling the methods for spreadsheets
    sheet = gSheets.spreadsheets()

    # --- GOOGLE IDs ---
    RA_SHARED_DRIVE_ID = '0ALf_5x-OEWxnUk9PVA'
    BIM_FOLDER_ID = '1k_CQdKrSGrDnB-nsgs8AgUWb9aX4SZeh'
    
    # Querying for the BIM Software folder within the shared drive
    query = "mimeType='application/vnd.google-apps.folder' and name = 'BIM Software' and parents in '{}' and trashed = false".format(BIM_FOLDER_ID)
    BIM_FOLDER_ID = getFileInDriveByQuery(query, RA_SHARED_DRIVE_ID)[0]['id']

    # Querying for the spreadsheet that contains all the IDs we need
    query = "mimeType='application/vnd.google-apps.spreadsheet' and name = 'Google IDs for scripts' and parents in '{}' and trashed = false".format(BIM_FOLDER_ID)
    idsSpreadsheetId = getFileInDriveByQuery(query, RA_SHARED_DRIVE_ID)[0]['id']

    idFields = sheet.values().get(spreadsheetId=idsSpreadsheetId, 
                                  range="A1:AZ20000").execute()['values']

    # Getting access to the working folders
    MAIN_FOLDER_ID = idFields[0][1]
    REVIT_SCHEDULES_FOLDER = idFields[1][1]
    UPLOADED_IMAGES_FOLDER = idFields[2][1]

    # ---FILE CREATION---
    # Checking if the Revit Schedules spreadsheet exists. If not, create it
    query = "mimeType='application/vnd.google-apps.spreadsheet' and name = '{}' and parents in '{}' and trashed = false".format(sheetName, REVIT_SCHEDULES_FOLDER)

    # Executes the query
    revitSchedule = getFileInDriveByQuery(query, RA_SHARED_DRIVE_ID)

    # If the spreadsheet does not exists, create it
    if not revitSchedule:

        revitSchedule = {
            'properties': {
                'title': '{}'.format(sheetName),
                },

                'sheets' : {
                    'properties' : {
                        'sheetId' : 0,
                        'title' : 'Cost Estimation',
                        'gridProperties' : {
                            'rowCount' : 3000,
                        },
                    }  
                },
        }

        sheet.create(body=revitSchedule,
                     fields='spreadsheetId').execute()
        
        # Checks if the spreadsheet was already created to prepare it to be moved to the correct folder
        query = "mimeType='application/vnd.google-apps.spreadsheet' and name = '{}' and trashed = false".format(sheetName)
        revitSchedule = gDrive.files().list(q=query,
                                            spaces='drive',
                                            fields='nextPageToken, files(id, name)',
                                            pageToken=None
                                            ).execute()['files']

        # Gets the id of the Revit Schedule File
        revitScheduleId = revitSchedule[0]['id']

        # Moving the recently created "Revit Schedule" from Root to the correct folder (Google API has not implemented yet a method to create spreadsheets directly in a folder)
        gDrive.files().update(fileId='{}'.format(revitScheduleId),
                                                 addParents=REVIT_SCHEDULES_FOLDER,
                                                 supportsAllDrives='true'
                                                ).execute()

        # Writing headers in the Schedule
        headerValues = [[x for x in CBCTemplate['headers'].keys()]]

        sheet.values().update(spreadsheetId=revitScheduleId, 
                              range="A1:AZ1", 
                              valueInputOption="USER_ENTERED", 
                              body={"values":headerValues}
                              ).execute()

        # Counts the amount of columns there are going to be
        amountOfFields = len(headerValues[0])

        # General sheet formatting (granular formatting will be applied afterwards)
        formatRequests = {
            "requests" : [
            # Update size of row to 21 pixels
            {
                "updateDimensionProperties" : {
                    "range" : {
                        "dimension": "ROWS",
                        "startIndex": 0,
                        "endIndex": 2000
                    },
                    "properties" : {
                        "pixelSize" : 21
                    },
                    "fields" : "pixelSize"
                }
            },

            # Sets the font to Black Arial Regular 10
            {
                "repeatCell" : {
                    "range" : {
                        "endColumnIndex" : amountOfFields + 1
                            },
                    "cell" : {
                        "userEnteredFormat" : {
                            "textFormat" : {
                                "fontFamily" : "Arial",
                                "fontSize" : 10,
                                "foregroundColor" : {
                                    "red" : 0,
                                    "blue" : 0,
                                    "green" : 0,
                                    "alpha" : 1
                                } 
                            },
                            "backgroundColor" : {
                                "red" : 1,
                                "blue" : 1,
                                "green" : 1,
                                "alpha" : 1
                            },
                            "verticalAlignment" : "MIDDLE",
                            "horizontalAlignment" : "LEFT",
                            "wrapStrategy" : "WRAP"
                        },
                    },
                    "fields" : "userEnteredFormat"       
                },
            }
            ]
        }

        sheet.batchUpdate(
                        spreadsheetId=revitScheduleId,
                        body=formatRequests
                        ).execute()

        # Formatting header
        formattedHeader = formatHeader(sheet, revitScheduleId, amountOfFields)

    # Gets the id of the Revit Schedule File
    revitScheduleId = revitSchedule[0]['id']

    # If the Revit Schedule file already exists or was just created, acquire its headers in the first row
    headers = sheet.values().get(spreadsheetId=revitScheduleId, range="A1:AZ1").execute()["values"][0]

    # ---IMAGE PROCESSING---
    # Executes a query and store all the existing images in a variable
    query = "mimeType='image/jpeg' and parents in '{}' and trashed = false".format(UPLOADED_IMAGES_FOLDER)

    # Uploads images that have not been uploaded it and append them to the imagesInDrive dictionary
    uploadImagesToDrive(UPLOADED_IMAGES_FOLDER)

    # Getting a list of dictionaries (name: id) containing all the images in the working Drive folder
    imagesInDrive = gDrive.files().list(q=query,
                                        spaces='drive',
                                        includeItemsFromAllDrives='true',
                                        supportsAllDrives='true',
                                        fields='nextPageToken, files(id, name)',
                                        pageToken=None
                                        ).execute()['files']

    

    # ---PARSING AND WRITING DATA---
    # Getting all the data-related variables
    
    # Compare acquired headers with valid headers, if there is a match, acquire the "Included" value in a new dictionary
    # (The "Included" value (0 or 1 in the CBCTemplate['headers'] values) determines if the header string is going to be used to evaluate parameters in the Elements)
    fieldsDictionary = {}

    for header in headers:

        if header in CBCTemplate['headers']:
            fieldsDictionary[header] = CBCTemplate["headers"][header]

        else:
            fieldsDictionary[header] = 0
    
    #--- The following process will get the complete headers' texts and their split version in a list of lists
    #--- for later comparison with Revit parameters name for further parameter extraction ---

    # Initialize list that will contain the complete string and its splits
    headerTextItem = []

    # Initialize list that will contain the 'headersText' lists
    headersText = []

    # List of symbols to use to split the strings
    splitSymbols = '; |( |) |# | '
    symbols = '[@_!#$%^&*()<>?/\|}{~:]'

    # Goes through each header field and use its string to look for the adequate parameter in the element to get its value
    # --(EXAMPLE: ITEM URL)
    for key, value in fieldsDictionary.items():
        if value == 0:
            continue
            
        if "UniqueID" in key:
            pass
        else:
            # Turning to lowercase the current key
            # --(EXAMPLE: item url)
            key = key.lower()

            # Capitalizes each word inside the string
            # --(EXAMPLE: Item Url)
            key = capwords(key) 

        # Split the string where it is delimiters
        headerTextItem = re.findall('\w+', key)

        if "UniqueID" in key:
            pass
        else:
            # Turns into Upper Case the words that are acronyms
            # --(EXAMPLE: Item Url
            #             Item --> Item
            #             Url --> URL
            #       Result: Item URL)
            for word in headerTextItem:
                position = headerTextItem.index(word)
                if word.upper() in CBCTemplate['acronyms']:
                    headerTextItem[position] = word.upper()

        # Add the complete string to the 'headerTextItem'
        if len(headerTextItem) > 1:
            sentence = ""

            for word in headerTextItem:
                sentence += "{0} ".format(word)

            # Removing whitespace at the end of sentence
            sentence = sentence.rstrip()

            headerTextItem.append(sentence)

        # Add the position of the column related to this item
        if "UniqueID" in key:
            headerTextItem.insert(0, list(fieldsDictionary.keys()).index(key))
       
        else:
            headerTextItem.insert(0, list(fieldsDictionary.keys()).index(key.upper()))

        # Adds the current header text item to the list of header items
        headersText.append(headerTextItem)
        
    #--- The following process will start looking at the elements in the model to start processing their parameters ---
    # Getting all the schedules in the model
    schedules = FilteredElementCollector(doc).OfClass(ViewSchedule).ToElements()

    # Initializing a list of all the values of a parameter in a model
    categoriesParametersList = []
    
    # List to store all the elements in the schedules
    allElements = []

    # Initializing list that will contain the UniqueIds of collected elements in order to avoid duplicated data
    elementsUniqueIDs = []

    notAllowedTypes = [View,
                       View3D,
                       ViewDrafting,
                       ViewSection,
                       ViewPlan,
                       ViewSheet,
                       FilledRegion,
                       Room,
                       Revision,
                       RevitLinkInstance,
                       Structure.AnalyticalModel,
                       Structure.AnalyticalModelSurface,
                       Structure.AnalyticalModelStick,
                       Area,
                       FootPrintRoof,
                       Mullion,
                       Architecture.Stairs,
                       Architecture.StairsLanding,
                       Architecture.StairsRun,
                       Architecture.Railing,
                       Architecture.BuildingPad
                       ]

    # Traversing each schedule in the model
    for schedule in schedules:
        
        # Getting the current schedule ID in order to get all the Element objects in it
        scheduleId = schedule.Id

        try:
            # All the elements in the current schedule
            elementSet = FilteredElementCollector(doc, scheduleId).ToElements()

            for element in elementSet:
                
                # Filters out elements that are instances of links inside the project
                if type(element) in notAllowedTypes:
                    continue

                # If the element is already collected, omit it
                if element.UniqueId in elementsUniqueIDs:
                    continue

                if type(element) == Panel:
                    if element.Name == "Empty" or element.Name == "M_Empty Panel":
                        continue
                
                elementsUniqueIDs.append(element.UniqueId)
                allElements.append(element)

                # Get type of element to extract Category parameters
                typeId = element.GetTypeId()
                typeElement = doc.GetElement(typeId)

                categoriesParametersList.append(typeElement.GetParameters("Category"))

        except:
            continue

    # Sorting the allElements list
    allElements = sorted(allElements, key=lambda x: x.Name)

    # Iterating through all the categories values and getting a clean list with all of them
    categoriesList = []

    for group in categoriesParametersList:
        for parameter in group:
            if parameter.IsShared:
                parameterString = parameter.AsString()

                if parameterString not in categoriesList and parameterString != "" and parameterString != None:
                    categoriesList.append(parameterString)
       
    categoriesList.sort()

    # --- Creating general table that will be written in the Google Spreadsheet ---
    generalTable = []

    for element in allElements:

        # Row that will be included in the general table
        elementRow = [None] * len(fieldsDictionary)
        cursorPosition = 0
        unitPosition = None

        # Get Type of Element in case that the Instance does not have the parameter intended to be captured
        typeId = element.GetTypeId()
        typeElement = doc.GetElement(typeId)

        # If the element does not have a Category parameter, continue with the next one
        try:
            for parameter in typeElement.GetParameters("Category"):
                try:
                    if parameter.IsShared:
                        category = parameter.AsString()
                
                except:
                    category = None

        except AttributeError:
            continue

        if category not in categoriesList or category == None:
            continue
        

        for group in headersText: # group is [Item URL, Item, URL]
            # Defines in what column the data will be written
            cursorPosition = group[0]
        
            if group[-1] == "Category": # If the current parameter to parse is Category, continue with the next one as Category is already defined in the previous loop
                parameterValue = category
                elementRow[cursorPosition] = parameterValue
                continue
            
            if category.lower().find("fixture") != -1 and cursorPosition > 7 and cursorPosition < 16: # If the current element is a fixture, do not get dimension data
                continue

            for word in group: # Item URL, Item, URL
                
                if isinstance(word, int): # Discards the iteration over the column position value
                    continue
                
                # Get the parameter if any of these special cases is true
                if word.lower() == "name": # If we are parsing the name, just do a simple get property
                    parameterValue = element.Name
                    elementRow[cursorPosition] = parameterValue
                    
                    if parameterValue != None:
                        break
                

                if word.lower() == "uniqueid": # If we are parsing the UniqueId, just do a simple get property
                    parameterValue = element.UniqueId
                    elementRow[cursorPosition] = parameterValue
                    
                    if parameterValue != None:
                        break
                
                if word.lower() == "count": # Temporary measure - If we are parsing the Count, just put 1 in the value as we are parsing through all the elements in the model
                    parameterValue = 1
                    elementRow[cursorPosition] = parameterValue
                    break

                if "image" in word.lower():
                    parameterValue = getParameterValueByName(typeElement, word)
                    elementRow[cursorPosition] = parameterValue

                    if parameterValue != None:
                        break

                # Gets the column position for the UNIT field
                if word.lower() == "unit":
                    if unitPosition == None:
                        unitPosition = group[0]
                
                # If the element being parsed is a wall use this block to return its parameters
                if type(element) == Wall:

                    if word.lower() == "height":
                        parameterValue = getParameterValueByName(element, "Unconnected Height")
                        parameterValue = convertUnits(parameterValue, group)

                    elif word.lower() == "length":
                        parameterValue = getParameterValueByName(element, "Length")
                        parameterValue = convertUnits(parameterValue, group)

                    elif word.lower() == "width":
                        parameterValue = getParameterValueByName(element, "Width")
                        parameterValue = convertUnits(parameterValue, group)

                    elif word.lower() == "area":
                        parameterValue = getParameterValueByName(element, "Area")
                        parameterValue = convertUnits(parameterValue, group)

                    elif "cost" in list(map(lambda x : str(x).lower(), group)):
                        parameterValue = getParameterValueByName(element, "Cost")
                    
                    elementRow[cursorPosition] = parameterValue

                    # Unit type assigning
                    if unitPosition:
                        elementRow[unitPosition] = assignUnitType(category.lower())

                    if parameterValue != None:
                        break

                
                # If reached a normal situation, get the parameter based on the word iteration
                parameterValue = getParameterValueByName(element, word)

                # If the value is numeric and it is a dimension, convert its units so they are displayed correctly
                parameterValue = convertUnits(parameterValue, group)
                
                # Writes the value in the row
                elementRow[cursorPosition] = parameterValue

                # If the instance does not return a parameter, try searching the parameter in the type
                if parameterValue == None: 
                    
                    # Repeats the process above
                    parameterValue = getParameterValueByName(typeElement, word)
                    parameterValue = convertUnits(parameterValue, group)
                    elementRow[cursorPosition] = parameterValue

                    if parameterValue == None: # If there is still not a parameter returned, set the result as None

                        elementRow[cursorPosition] = parameterValue
                        
                    else:
                        break
                
                else:
                    break

        # Unit type assigning
        if elementRow[unitPosition] == None:
            elementRow[unitPosition] = assignUnitType(category.lower())
        else:
            pass
        
        # Adding finished row to generalTable
        generalTable.append(elementRow)

    # Append headers' row to generalTable
    generalTable.insert(0, list(fieldsDictionary.keys()))

    # Collapses and counts any adjacent and repeated elements
    generalTable = countAndCollapseRepeatedItems(generalTable)

    # Getting the total quantities (Count, SQF, FT) per item
    generalTable = getTotalQuantitiesPerItem(generalTable)

    # Defining total cost for each item in the table
    generalTable = getTotalCostsPerItem(generalTable)

    # Consolidating the images in the imported spreadsheet
    generalTable = consolidateImages(generalTable, imagesInDrive)
    

    # Processing the general table and writing the data to Google Sheets
    # Initializes the macro list that will be written to the Google Spreadsheet
    dataToWrite = []

    for category in categoriesList:
        # Initializes/Overrides the list that will contain the data for each group/category
        dataGroupItem = []

        for row in generalTable:
            if row[0] != category:
                continue

            dataGroupItem.append(row)

        dataToWrite.append(dataGroupItem)
        dataGroupItem.insert(0, category)

    # Starts writing the data to the spreadsheet iterating accross each group
    previousGroup = None
    groupCount = 0
    totalProjectCost = 0

    for group in dataToWrite:
        
        # Initializes the group to be written to the spreadsheet
        currentGroup = CategoryRowGroup(group[0], 1, revitScheduleId, sheet) if groupCount == 0 \
                       else CategoryRowGroup(group[0], previousGroup.currentRow, revitScheduleId, sheet)

        # Adds data to the group
        for row in group:
            if type(row) != list:
                continue

            currentGroup.appendRow(row)

        # Creates/Write group in the Google spreadsheet
        currentGroup.create()

        groupCount += 1
        previousGroup = currentGroup

        totalProjectCost += currentGroup.totalCost
        lastRow = currentGroup.currentRow + 1



    # Adding Totals row

    # Write request
    totalValues = [[None] * len(fieldsDictionary)]
    totalValues[0][0] = "PROJECT TOTALS"
    totalValues[0][24] = totalProjectCost
    totalValues[0][25] = "CAD"
    totalValues[0][26] = totalProjectCost * 0.8

    # Format request
    # Variables
    colorWhite = {
        "red" : 1,
        "green" : 1,
        "blue" : 1,
        "alpha" : 1
    }

    colorGray100 = {
        "red" : 0.35,
        "green" : 0.35,
        "blue" : 0.35,
        "alpha" : 1
    }

    colorBlack = {
        "red" : 0,
        "green" : 0,
        "blue" : 0,
        "alpha" : 1
    }

    formatRequests = {
            "requests" : [
            # Sets the font to White Arial Bold 10, and the gray background
            {
                "repeatCell" : {
                    "range" : {
                        "startRowIndex" : lastRow - 1,
                        "endRowIndex" : lastRow,
                        "endColumnIndex" : amountOfFields
                            },
                    "cell" : {
                        "userEnteredFormat" : {
                            "textFormat" : {
                                "bold" : "true",
                                "fontFamily" : "Arial",
                                "fontSize" : 10,
                                "foregroundColor" : colorWhite 
                                },
                            "backgroundColor" : colorGray100,
                            "verticalAlignment" : "MIDDLE",
                            "horizontalAlignment" : "CENTER",
                            "wrapStrategy" : "WRAP"
                            }
                        },
                    "fields" : "userEnteredFormat"
                }
            }
        ]
    }

    try:
        # Writing row
        # Data
        sheet.values().update(spreadsheetId=revitScheduleId, 
                                range="A{}:AZ{}".format(lastRow, lastRow + 1), 
                                valueInputOption="USER_ENTERED", 
                                body={"values":totalValues}
                                ).execute()

        # Style
        sheet.batchUpdate(
                        spreadsheetId=revitScheduleId,
                        body=formatRequests
                        ).execute()

    except HttpError as error:
        if error.resp.status in [429]:
            print("The limits of requests (uploads and downloads to Google Drive was exceeded,\
                   this can happen when you are working with large sets of data. If this error appeared please contact the programmer in charge to solve it")

    return generalTable