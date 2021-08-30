#! python3
from __future__ import print_function

import os

# Getting the local route to the %APPDATA% folder
appDataPath = os.getenv('APPDATA')

# Getting the local route to the %LOCALAPPDATA% folder
localAppDataPath = os.getenv('LOCALAPPDATA')

# Getting the local route to the %TEMP% folder
tempPath = os.getenv('TEMP')

import sys
sys.path.append(appDataPath)
sys.path.append(localAppDataPath)
sys.path.append(tempPath)

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
import csv
from collections import OrderedDict
import System

# The __init__ file with all the initialization procedures is __init__.py in the 'libraries' folder
from libraries import *
from CBCEstimatesTemplate import CBCTemplate

# Revit API imports
clr.AddReference("System.Drawing")
import System.Drawing
clr.AddReference("RevitAPI")
clr.AddReference('RevitAPIUI')
clr.AddReference("RevitNodes")
clr.AddReference("RevitServices")
from Autodesk.Revit.UI import *
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import Transaction

# Third Party Imports
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.http import MediaFileUpload
from googleapiclient.errors import HttpError

# Getting the local base folder path
# Checking if the GoogleAPI base folder exists
# if os.path.exists("{}//Autodesk//Revit//Addins//GoogleAPI".format(appDataPath)):
#     credentialsFolder = "{}//Autodesk//Revit//Addins//GoogleAPI".format(appDataPath)

# else:
#     os.mkdir("{}//Autodesk//Revit//Addins//GoogleAPI".format(appDataPath))
#     credentialsFolder = "{}//Autodesk//Revit//Addins//GoogleAPI".format(appDataPath)

credentialsFolder = "{}//credentials".format(os.path.dirname(currentPath))

# If modifying these scopes, delete the file token.json located at the GoogleAPI folder in your local Dynamo folders.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 
          'https://www.googleapis.com/auth/drive']

creds = None
# The file token.json stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
if os.path.exists('{}//token.json'.format(credentialsFolder)):
    creds = Credentials.from_authorized_user_file('{}//token.json'.format(credentialsFolder), SCOPES)

# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            '{}//credentials.json'.format(credentialsFolder), SCOPES)
        creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open('{}//token.json'.format(credentialsFolder), 'w') as token:
        token.write(creds.to_json())

# Accessing the APIs
gSheets = build('sheets', 'v4', credentials=creds)
gDrive = build('drive', 'v3', credentials=creds)

# Accessing current Revit Document
doc = __revit__.ActiveUIDocument.Document

### ---FUNCTIONS---

def assignUnitType(importedCategory):

    importedCategory = importedCategory.lower()

    # Type of dimensions per category
    dimensionCategories = {
        "casework" : "UNIT",
        "ceiling" : "SQF",
        "curtain panels" : "SQF",
        "doors" : "UNIT",
        "entourage" : "UNIT",
        "facade" : "SQF",
        "fixture" : "UNIT",
        "floor" : "SQF",
        "furniture" : "UNIT",
        "lighting fixture" : "UNIT",
        "railing" : "FT",
        "roof" : "SQF",
        "structural column" : "FT",
        "structural framing" : "FT",
        "column" : "FT",
        "framing" : "FT",
        "wall" : "SQF",
        "window" : "UNIT",
    }

    for key in dimensionCategories.keys():

        if importedCategory.find(key) != -1:
            return dimensionCategories[key]

# DECORATORS
def checkForHttpError429(originalFunction):

    # Handles the error when the quota limit is exceeded
    def wrapperFunction(*args, **kwargs):
        try:
            return originalFunction(*args, **kwargs)
        except HttpError as error:
            if error.resp.status in [429]:
                print("The limits of requests (uploads and downloads to Google Drive was exceeded,\
                      this can happen when you are working with large sets of data. If this error appeared please contact the programmer in charge to solve it")

                pass

    return wrapperFunction

# FUNCTIONS
def consolidateImages(list, imagesInDrive):
    # Transposing the schedule to sobstitute the image fields in an easier way
    transposedList = transpose(list)

    # Starts to iterate across each column
    for column in transposedList:

        # Checking that the script is standing at the image columns
        if 'image' in column[0].lower():

            for i in range(1, len(column)):
                
                # If there is no image in that cell, continue iterating
                if column[i] == "" or column[i] == "<None>":
                    
                    if column[i] == "<None>":
                        column[i] = ""

                    continue 
                
                # Else, start iterating over the uploaded images in Drive to find a match
                else:

                    for item in imagesInDrive:

                        # If there is a match, get the direct link to the image in the 'linkImage' variable, 
                        # and then substitute the data in the Smart Schedule cell with that link
                        if column[i] == item['name']:

                            linkImage = '=IMAGE("https://docs.google.com/uc?export=view&id={}", 1)'.format(item['id'])
                            column[i] = linkImage
                        
                        else:
                            continue

        else:
            continue

    # Transposing the Smart Schedule back to its original state
    list = transpose(transposedList)

    return list

def countAndCollapseRepeatedItems(list):

    previousItem = None
    countIndex = []
    counter = 0

    # Iterates through the original list in order to iterate over the original state
    while counter < len(list):
        
        listLength = len(list)
        item = list[counter]
        coincidences = 0
        
        if previousItem is None:
            previousItem = list[list.index(item)]
            counter += 1
            continue
        
        # Looks for coincidences in just the first 19 fields
        # Span of list to check
        headersRangeToCheck = item[:18]

        for i in range(len(headersRangeToCheck)):

            # Skipping the Count field as they will be different between elements
            if countIndex:
                if i == countIndex[0]:
                    continue
            
            # Converting the data to compare to string in order to avoid errores
            variable1 = str(item[i])
            variable2 = str(previousItem[i])

            if variable1 == variable2:
                coincidences += 1
            else:
                previousItem = list[list.index(item)]
                counter += 1
                count = 1
                break
            
            amountOfFieldsToCheck = len(headersRangeToCheck) - 1
            if coincidences >= amountOfFieldsToCheck:
                count += 1

                # Removes the current item from the list
                del list[counter]


        # Gets the Cost column position from the original list
        if len(countIndex) == 0:
            for i, field in enumerate(map(lambda x : x.lower(), list[0])):

                if "count" in field:

                    countIndex.append(i)

        # Replaces the count on the Count field of the previous item
        previousItem[countIndex[0]] = count

    return list

def convertScheduleNameToNumber(string1):
    result = ""
    numbers = []

    for i in range(0, len(string1), int(len(string1) / 10)):
        numbers.append(ord(string1[i]))

    for number in numbers:

        result = result + str(number)[0]

    return int(result[:9])

def convertUnits(value, listOfUnits):

    # Mapping all the values in the list of units to lower case
    listOfUnits = list(map(lambda x : x.lower() if isinstance(x, str) else x, listOfUnits))

    if value != None:

        if isinstance(value, int) or isinstance(value, float):

            if "ft" in listOfUnits:
                value = UnitUtils.Convert(value, DisplayUnitType.DUT_DECIMAL_FEET, DisplayUnitType.DUT_DECIMAL_FEET)
                return value
            
            elif "mm" in listOfUnits:
                value = UnitUtils.Convert(value, DisplayUnitType.DUT_DECIMAL_FEET, DisplayUnitType.DUT_MILLIMETERS)
                return value

            elif "sqf" in listOfUnits:
                value = UnitUtils.Convert(value, DisplayUnitType.DUT_SQUARE_FEET, DisplayUnitType.DUT_SQUARE_FEET)
                return value

            elif "sqm" in listOfUnits:
                value = UnitUtils.Convert(value, DisplayUnitType.DUT_SQUARE_FEET, DisplayUnitType.DUT_SQUARE_METERS)
                return value

            else:
                return value

    return value

def exportSchedules(schedules):
    exportedFiles = []
    scheduleNames = []
    SSExportOptions = ViewScheduleExportOptions()
    for schedule in schedules:
        
        scheduleName = schedule.LookupParameter("View Name").AsString()
        exportedFilename = "{}.txt".format(scheduleName)

        schedule.Export("{}".format(tempPath), exportedFilename, SSExportOptions)

        exportedFiles.append(exportedFilename) # List of exported files
        scheduleNames.append(scheduleName) # List of the names of the schedules without the .txt extension

    # Reordering and managing the lists
    result = list(sorted(zip(scheduleNames, exportedFiles))) # Sorting by the schedule view name
    result = transpose(result)
    result = list(reversed(result)) # Getting the list with the txt files at index 0

    return result

def formatHeader(gSheetsObject, revitScheduleId, amountOfColumns=21):
    
    # # Adding one number as the endIndex value is exclusive
    # amountOfColumns += 1

    # Users that can edit protected ranges
    allowedUsers = ["ricardo.salas@theciel.co", 
                    "ricardo.salas.cg@gmail.com",
                    "safak@theciel.co",
                    "sam.williams@theciel.co",
                    "enrique.galicia@theciel.co",]
    
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

    # Batch of requests
    requests = {
        "requests" : [
            # Update size of row to 42 pixels
            {
                "updateDimensionProperties" : {
                    "range" : {
                        "dimension": "ROWS",
                        "startIndex": 0,
                        "endIndex": 1
                    },
                    "properties" : {
                        "pixelSize" : 42
                    },
                    "fields" : "pixelSize"
                }
            },

            # Sets the font to White Arial Bold 10, and the gray background
            {
                "repeatCell" : {
                    "range" : {
                        "endRowIndex" : 1,
                        "endColumnIndex" : amountOfColumns
                            },
                    "cell" : {
                        "userEnteredFormat" : {
                            "textFormat" : {
                                "bold" : "true",
                                "fontFamily" : "Arial",
                                "fontSize" : 10,
                                "foregroundColor" : colorWhite 
                                },
                            "borders" : {
                                "bottom" : {
                                    "style" : "SOLID_THICK",
                                    "color" : colorBlack
                                }
                            },
                            "backgroundColor" : colorGray100,
                            "verticalAlignment" : "MIDDLE",
                            "horizontalAlignment" : "CENTER",
                            "wrapStrategy" : "WRAP"
                            }
                        },
                    "fields" : "userEnteredFormat"
                }
            },

            # Freeze the header row
            {
                "updateSheetProperties" : {
                    "properties" : {
                        "gridProperties" : {
                            "frozenRowCount" : 1,
                            "frozenColumnCount" : 2
                        }
                    },
                    "fields" : "gridProperties.frozenRowCount, gridProperties.frozenColumnCount"
                }
            },

            # Protects the header
            {
                "addProtectedRange" : {
                    "protectedRange" : {
                        "protectedRangeId": 0,
                        "range" : {
                            "endRowIndex": 1,
                            "endColumnIndex": amountOfColumns
                        },
                        "description" : "The names of the columns are protected",
                        "warningOnly" : "false",
                        "requestingUserCanEdit" : "true",
                        "editors" : {
                            "users": allowedUsers
                        }
                    },
                },
            }
        ]
    }

    return gSheetsObject.batchUpdate(
                                    spreadsheetId=revitScheduleId,
                                    body=requests
                                    ).execute()

def getExportedCSVAsList():
    tempDummyFilePath = r"C:\Users\Ricardo Salas\Desktop\Temporal\dummySpreadsheet.csv"

    with open(tempDummyFilePath, 'r') as DS:
        dummySpreadsheet = list(csv.reader(DS, delimiter=",", quotechar='"'))

        return dummySpreadsheet

def getExportedSchedule(exportedFiles):

    importedFiles = []

    for scheduleFile in exportedFiles[0]:

        with open("{}\{}".format(tempPath, scheduleFile), 'r', encoding="utf-8") as SS:
            
            importedSchedule = list(csv.reader(SS, delimiter="\t", quotechar='"'))

            # Removes the title row
            importedSchedule.pop(0) 

            while len(importedSchedule[-1][0]) <= 1: # Removes any blank rows at the end of the schedule
                importedSchedule.pop(-1)

            importedFiles.append(importedSchedule)

    return [importedFiles, exportedFiles[1]]

def getLoadedImages():

    filteredImagesNames = []
    filteredImages = []
    listParameters = []

    tt = Transaction(doc, 'getLoadedImages')
    tt.Start()

    images = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_RasterImages)

    for image in images:

        if image.ToString() == "Autodesk.Revit.DB.ImageType":

            filteredImagesNames.append(image.LookupParameter("Type Name").AsString())
            filteredImages.append(image.GetImage())


    # Declaring the dictionary to store the name: bitmap pairs
    imagesDictionary = {}

    # Populating the dictionary
    for name, bitmap in zip(filteredImagesNames, filteredImages):

        imagesDictionary["{}".format(name)] = bitmap

    tt.Commit()

    return imagesDictionary

def getParametersList(target):

    parameters = target.Parameters
    listParameters = []

    for parameter in parameters:

        listParameters.append(parameter.Definition.Name)

    return listParameters

def getTotalCostsPerItem(generalTable):
    
    try:
        for item in generalTable[1:]:

                item[23] = item[7] * item[17]
                
    except TypeError:
        pass

    return generalTable 

def checkAndReturnParameter(parameter, *args):

    storageType = parameter.StorageType

    if storageType == StorageType.Double:
        return parameter.AsDouble()

    elif storageType == StorageType.ElementId:
        if args:
            if "Image" in args[0]:
                result = parameter.AsValueString()

                if result == "<None>" or result == None:
                    result = None
                else:
                    return parameter.AsValueString()
                
            else:
                return parameter.AsElementId().IntegerValue
        else:
            return parameter.AsElementId().IntegerValue
            
    elif storageType == StorageType.Integer:
        return parameter.AsInteger()

    elif storageType == None:
        return parameter.AsValueString()

    elif storageType == StorageType.String:
        return parameter.AsString()
    
    else:
        return ""

def getParameterValue(parameter):
    
    return checkAndReturnParameter(parameter)

def getParameterValueByName(element, string):
    parameter = element.GetParameters(string)

    if isinstance(parameter, list):

        for item in parameter:

            result = checkAndReturnParameter(item, string)

            if result != None or result != "":

                return result
            
            elif parameter.index(item) == len(parameter) - 1: # If we are at the end of the list, return whatever is the result

                return result
    else:
        
        return checkAndReturnParameter(parameter, string)

def getFileInDriveByQuery(query, driveID):
    return gDrive.files().list(q=query,
                               corpora='drive',
                               spaces='drive',
                               driveId='{}'.format(driveID),
                               includeItemsFromAllDrives='true',
                               supportsAllDrives='true',
                               fields='nextPageToken, files(id, name)',
                               pageToken=None
                               ).execute()['files']

def transpose(object1):
    return list(map(list, zip(*object1)))

@checkForHttpError429
def uploadImagesToDrive(UPLOADED_IMAGES_FOLDERID):

    # Getting all the loaded images in the Revit document and putting them in a dictionary "ImageName: Bitmap"
    loadedImages = getLoadedImages()

    query = "mimeType='image/jpeg' and parents in '{}' and trashed = false".format(UPLOADED_IMAGES_FOLDERID)
    
    # Executes a query and store all the existing images in a variable
    imagesInDrive = gDrive.files().list(q=query,
                                  spaces='drive',
                                  includeItemsFromAllDrives='true',
                                  supportsAllDrives='true',
                                  fields='nextPageToken, files(id, name)',
                                  pageToken=None
                                  ).execute()['files']

    # Uploading files to Google Drive
    for key, value in loadedImages.items():
        
        decision = False
        
        # Checks if the file already exists in the Drive Folder by its name
        for item in imagesInDrive:

            if key == item['name']:

                decision = True
                break

            else:

                decision = False

        # If it does not exist in the Drive Folder, it will be first exported to TEMP and then uploaded from there to Drive
        if decision == False:
     
            # Saving the images to TEMP
            tt = Transaction(doc, "saveLoadedImagesToTemp")
            tt.Start()
            value.Save('{}//{}'.format(tempPath, key), System.Drawing.Imaging.ImageFormat.Jpeg)
            tt.Commit()

            # Uploading the temp images to the Google Drive Folder
            file_meta = {
                'name': key,
                'parents': [UPLOADED_IMAGES_FOLDERID]
            }

            media = MediaFileUpload('{}//{}'.format(tempPath, key), mimetype='image/jpeg') # This is from where the file will be uploaded, and the filetype

            uploadedFile = gDrive.files().create(body=file_meta,
                                                 media_body=media,
                                                 fields='id, name',
                                                 supportsAllDrives='true'
                                                 ).execute()

            # Appending the uploaded files to the previous query, this is the updated list of images in the Drive Folder
            imagesInDrive.append(uploadedFile)
    
    return imagesInDrive

### ---CLASSES---
class CategoryRowGroup:

    def __init__(self, categoryName, currentRow, revitScheduleId, sheet):
        self.categoryName = categoryName
        self.currentRow = currentRow
        self.revitScheduleId = revitScheduleId
        self.totalCost = 0
        self.sheet = sheet
        self.data = []

    def appendRow(self, inputRow):
        self.data.append(inputRow)
        self.totalCost += inputRow[17] * inputRow[7]
    
    @checkForHttpError429
    def create(self):
        amountOfColumns = len(CBCTemplate['headers'])

        # Formatting header
        # Adding one number as the endIndex value is exclusive
        amountOfColumns += 1

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

        # Batch of requests
        requests = {
            "requests" : [
                # Sets the font to White Arial Bold 10, and the gray background
                {
                    "repeatCell" : {
                        "range" : {
                            "startRowIndex" : self.currentRow,
                            "endRowIndex" : self.currentRow + 1,
                            "endColumnIndex" : amountOfColumns - 1
                                },
                        "cell" : {
                            "userEnteredFormat" : {
                                "textFormat" : {
                                    "bold" : "true",
                                    "fontFamily" : "Arial",
                                    "fontSize" : 10,
                                    "foregroundColor" : colorWhite 
                                    },
                                "borders" : {
                                    "bottom" : {
                                        "style" : "SOLID_MEDIUM",
                                        "color" : colorBlack
                                    }
                                },
                                "backgroundColor" : colorGray100,
                                "verticalAlignment" : "MIDDLE",
                                "horizontalAlignment" : "CENTER",
                                "wrapStrategy" : "WRAP"
                                }
                            },
                        "fields" : "userEnteredFormat"
                    }
                },
            ]
        }

        # Header of group
        groupHeader = [""] * amountOfColumns
        groupHeader[0] = self.categoryName
        groupHeader[24] = self.totalCost
        groupHeader[25] = "CAD"
        groupHeader[26] = self.totalCost * 0.8

        # Append header to data list
        self.data.insert(0, groupHeader)

        rowNumber = self.currentRow + 1
        range = "A{0}:AZ20000".format(rowNumber)

        self.sheet.values().update(spreadsheetId=self.revitScheduleId, 
                                    range=range, 
                                    valueInputOption="USER_ENTERED", 
                                    body={"values":self.data}
                                    ).execute()

        self.sheet.batchUpdate(
                                spreadsheetId=self.revitScheduleId,
                                body=requests
                                ).execute()


        if len(self.data) > 1:
            # Batch of requests
            requests = {
                "requests" : [
                    # Groups the rows under this header
                    { 
                        "addDimensionGroup": {
                            "range": {
                                "dimension": "ROWS",
                                "sheetId": 0,
                                "startIndex": self.currentRow + 1,
                                "endIndex": self.currentRow + len(self.data)
                            }
                        }
                    },
                ]
            }

            self.sheet.batchUpdate(
                                    spreadsheetId=self.revitScheduleId,
                                    body=requests
                                    ).execute()

        # Determines the row position for the next categoryRowGroup
        self.currentRow = self.currentRow + len(self.data)