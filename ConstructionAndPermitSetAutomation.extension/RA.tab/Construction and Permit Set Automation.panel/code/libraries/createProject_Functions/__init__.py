#!/usr/bin/env python
"""
TESTED REVIT APIs: 2020, 2021
"""
from time import sleep

import os
import sys
currentPath = os.path.abspath(os.path.dirname(os.path.dirname(__file__)))
currentPath = os.path.join(currentPath, "createProject_Functions")

# Revit API imports
import clr
clr.AddReference("System")
clr.AddReference("RevitAPI")
clr.AddReference('RevitAPIUI')
clr.AddReference("RevitNodes")
from Autodesk.Revit.UI import *
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Architecture import *
from Autodesk.Revit.DB.Structure import *
from Autodesk.Revit.DB.Analysis import *
from Autodesk.Revit.DB.Electrical import *
from Autodesk.Revit.DB.Mechanical import *
from Autodesk.Revit.DB.Plumbing import *
from Autodesk.Revit.ApplicationServices import Application

from System import Guid
from System.Collections.Generic import ICollection, List

from materialOptions import MaterialOption

doc = __revit__.ActiveUIDocument.Document
app = doc.Application

# MAIN FUNCTIONS

def importFromExistingProject(projectPath, importOptionsFromUI):

    openedExistingProject = openProjectByFilePath(projectPath)

    # Gettin types from existing document
    existingElements = FilteredElementCollector(openedExistingProject).WhereElementIsElementType().ToElements()
    existingElements = list(map(lambda x : x, existingElements)) # Converting native list Python type
    existingElements.sort()

    # Filtering Types we do not want to include
    notAllowedTypes = [AnalyticalLinkType,
                       BrowserOrganization,
                       CADLinkType,
                       ConceptualConstructionType,
                       DistributionSysType,
                       ElementType,
                       FluidType,
                       GroupType,
                       ImageType,
                       FamilySymbol,
                       MultiReferenceAnnotationType,
                       MechanicalEquipmentSetType,
                       MechanicalSystemType,
                       MEPAnalyticalConnectionType,
                       MEPBuildingConstruction,
                       PointLoadType,
                       RevitLinkType,
                       PipeScheduleType,
                       SiteLocation,
                       StructuralConnectionApprovalType,
                       TemperatureRatingType,
                       TilePattern,
                       WireMaterialType,
                       ]

    existingElements = list(filter(lambda x : type(x) not in notAllowedTypes, existingElements))

    # Getting the IDs of the elements we want to transfer
    existingElementsIDs = List[ElementId](list(map(lambda x : x.Id, existingElements)))

    # Collecting doors, furniture, and windows family types as they were not collected by the previous FilteredElementCollector
    existingFamilies = FilteredElementCollector(openedExistingProject).OfClass(Family)
    existingFamilies = list(map(lambda x : x, existingFamilies))

    allowedFamilyCategories = ["Doors",
                               "Windows"]

    if importOptionsFromUI["Import Products"]:

        allowedFamilyCategories.append("Furniture")
        allowedFamilyCategories.append("Plumbing Fixtures")
        allowedFamilyCategories.append("Electrical Equipment")
        allowedFamilyCategories.append("Specialty Equipment")
                               
    # Controlling the categories we want to transfer                               
    filteredExistingFamiliesSymbolsIDs = []
    existingFamilies.sort(key=lambda x : x.Name)
    for family in existingFamilies:
        
        if family.FamilyCategory.Name in allowedFamilyCategories:
            
            for id in family.GetFamilySymbolIds():

                filteredExistingFamiliesSymbolsIDs.append(id)

    # Getting the IDs of the families we want to transfer
    filteredExistingFamiliesSymbolsIDs = List[ElementId](list(map(lambda x : x, filteredExistingFamiliesSymbolsIDs)))

    # Adding the collected families symbols IDs to the list that contains all the elements to be imported
    for id in filteredExistingFamiliesSymbolsIDs:

        existingElementsIDs.Add(id)

    # Getting the view templates from the existing project
    existingViewTemplates = FilteredElementCollector(openedExistingProject).OfClass(View)
    existingViewTemplates = list(filter(lambda x : x.IsTemplate == True, list(map(lambda x : x, existingViewTemplates))))

    # Adding the collected view templates to the list that contains all the elements to be imported
    for template in existingViewTemplates:
        
        existingElementsIDs.Add(template.Id)

    # Cleaning the list from duplicates in case they are
    cleanedList = []
    for id in existingElementsIDs:

        if id not in cleanedList:
            cleanedList.append(id)

    existingElementsIDs = List[ElementId](cleanedList)

    # Transfering the elements to the new document
    tt = Transaction(doc, "Transfer Elements")
    tt.Start()

    importOptions = CopyPasteOptions()
    importOptions.SetDuplicateTypeNamesHandler(HideAndAcceptDuplicateTypesHandler())
    ElementTransformUtils.CopyElements(openedExistingProject, existingElementsIDs, doc, Transform.Identity, importOptions)

    tt.Commit()

    # Transfering the sheets to the new document
    if importOptionsFromUI["Import Sheets"]:
        transferSheetsFromProject(openedExistingProject)

    # Transferring the schedules as how they are set up to the new project
    transferSchedulesFromProject(openedExistingProject)

    # Close existing project
    openedExistingProject.Close()

def collectConceptualElements():

    collectedWalls = FilteredElementCollector(doc).OfClass(Wall).ToElements()
    collectedWallTypes = FilteredElementCollector(doc).OfClass(WallType).ToElements()
    collectedFloors = FilteredElementCollector(doc).OfClass(Floor).ToElements()
    collectedFloorTypes = FilteredElementCollector(doc).OfClass(FloorType).ToElements()
    collectedDoors = []
    collectedWindows = []
    collectedRoofs = []
    
    collectedElements = {"Collected Walls" : collectedWalls,
                         "Collected Wall Types" : collectedWallTypes,
                         "Collected Floors" : collectedFloors,
                         "Collected Floor Types" : collectedFloorTypes,
                         "Collected Doors" : collectedDoors,
                         "Collected Windows" : collectedWindows,
                         "Collected Roofs" : collectedRoofs
                        }

    return collectedElements

def replaceConceptualElements(projectPath, collectedElements, selectedOption):

    # Opening the project from which we have been importing data again
    openedExistingProject = openProjectByFilePath(projectPath)

    # Collections of exterior and interior walls and floors from the existing project
    existingExteriorWallTypes = list(filter(lambda x : x.Function == WallFunction.Exterior, collectedElements["Collected Wall Types"]))
    existingInteriorWallTypes = list(filter(lambda x : x.Function == WallFunction.Interior, collectedElements["Collected Wall Types"]))

    existingFloorTypes = list(filter(lambda x : x.get_Parameter(BuiltInParameter.FUNCTION_PARAM) != None, collectedElements["Collected Floor Types"]))
    existingExteriorFloorTypes = list(filter(lambda x : x.get_Parameter(BuiltInParameter.FUNCTION_PARAM).AsValueString() == "Exterior", existingFloorTypes))
    existingInteriorFloorTypes = list(filter(lambda x : x.get_Parameter(BuiltInParameter.FUNCTION_PARAM).AsValueString() == "Interior", existingFloorTypes))

    # Checking for what type of wall has the higher amount of area in the existing project in order to determine what is the main (or most used) exterior and interior type of walls
    existingExteriorWallsInstances = FilteredElementCollector(openedExistingProject).OfClass(Wall).ToElements()
    existingInteriorWallsInstances = FilteredElementCollector(openedExistingProject).OfClass(Wall).ToElements()

    existingExteriorFloorsInstances = FilteredElementCollector(openedExistingProject).OfClass(Floor).ToElements()
    existingInteriorFloorsInstances = FilteredElementCollector(openedExistingProject).OfClass(Floor).ToElements()
    
    def getWallTypeNameByMaximumArea(wallTypes, wallInstances):
    
        sortedWallInstances = []

        for wallType in wallTypes:

            wallTypeGroup = []

            for wall in wallInstances:
                
                if wall.WallType.GetParameters("Type Name")[0].AsString() == wallType.GetParameters("Type Name")[0].AsString():

                    wallTypeGroup.append(wall)
            
            sortedWallInstances.append(wallTypeGroup)

        # Cleaning list from empty lists
        sortedWallInstances = [x for x in sortedWallInstances if x != []]

        # Creating a dictionary that will show the wall type as the key and the summed total area of all its instances
        wallsDictionary = {}

        if len(sortedWallInstances) == 1:
            # Getting the name of the wall type related to the current group
            wallTypeName = sortedWallInstances[0][0].WallType.GetParameters("Type Name")[0].AsString()

            # Excluding Curtain Walls
            if sortedWallInstances[0][0].WallType.Kind != WallKind.Curtain:

                # Getting the sum of the areas of all the walls within this group
                wallAreas = list(map(lambda x : x.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).AsDouble(), sortedWallInstances[0]))

                totalWallArea = sum(wallAreas)

                wallsDictionary["{}".format(wallTypeName)] = totalWallArea

        else:

            for group in sortedWallInstances:

                # Getting the name of the wall type related to the current group
                wallTypeName = group[0].WallType.GetParameters("Type Name")[0].AsString()

                # Excluding Curtain Walls and walls that are used as curtain panels
                if group[0].WallType.Kind == WallKind.Curtain or group[0].get_Parameter(BuiltInParameter.ELEM_CATEGORY_PARAM).AsValueString() == "Curtain Panels":
                    continue

                # Getting the sum of the areas of all the walls within this group
                wallAreas = list(map(lambda x : x.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).AsDouble(), group)) # Multiplies the height by the length of the wall in order to get the wall area

                totalWallArea = sum(wallAreas)

                wallsDictionary["{}".format(wallTypeName)] = totalWallArea

        # Getting wall type with more area from the dictionary
        maxWallType = max(wallsDictionary, key=wallsDictionary.get)
                
        return maxWallType

    def getFloorTypeNameByMaximumArea(floorTypes, floorInstances):
    
        sortedFloorInstances = []

        for floorType in floorTypes:

            floorTypeGroup = []

            for floor in floorInstances:
                if floor.FloorType.GetParameters("Type Name")[0].AsString() == floorType.GetParameters("Type Name")[0].AsString():

                    floorTypeGroup.append(floor)
            
            sortedFloorInstances.append(floorTypeGroup)

        # Cleaning list from empty lists
        sortedFloorInstances = [x for x in sortedFloorInstances if x != []]  

        # Creating a dictionary that will show the floor type as the key and the summed total area of all its instances
        floorsDictionary = {}

        if len(sortedFloorInstances) == 1:
            # Getting the name of the floor type related to the current group
            floorTypeName = sortedFloorInstances[0][0].FloorType.GetParameters("Type Name")[0].AsString()

            # Getting the sum of the areas of all the floors within this group
            floorAreas = list(map(lambda x : x.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble(), sortedFloorInstances[0][0]))

            totalFloorArea = sum(floorAreas)

            floorsDictionary["{}".format(floorTypeName)] = totalFloorArea

        else:

            for group in sortedFloorInstances:

                # Getting the name of the floor type related to the current group
                floorTypeName = group[0].FloorType.GetParameters("Type Name")[0].AsString()

                # Getting the sum of the areas of all the floors within this group
                floorAreas = list(map(lambda x : x.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble(), group))

                totalFloorArea = sum(floorAreas)

                floorsDictionary["{}".format(floorTypeName)] = totalFloorArea

        # Getting floor type with more area from the dictionary
        maxFloorType = max(floorsDictionary, key=floorsDictionary.get)
                
        return maxFloorType

    exteriorWallTypeName = getWallTypeNameByMaximumArea(existingExteriorWallTypes, existingExteriorWallsInstances)
    interiorWallTypeName = getWallTypeNameByMaximumArea(existingInteriorWallTypes, existingInteriorWallsInstances)
    exteriorFloorTypeName = getFloorTypeNameByMaximumArea(existingExteriorFloorTypes, existingExteriorFloorsInstances)
    interiorFloorTypeName = getFloorTypeNameByMaximumArea(existingInteriorFloorTypes, existingInteriorFloorsInstances)


    # Creating Material Options
    materialOption1 = MaterialOption.CreateNewOptionSet()
    materialOption1["Walls"]["Exterior"] = list(filter(lambda x : x.GetParameters("Type Name")[0].AsString() == exteriorWallTypeName, collectedElements["Collected Wall Types"]))[0]
    materialOption1["Walls"]["Interior"] = list(filter(lambda x : x.GetParameters("Type Name")[0].AsString() == interiorWallTypeName, collectedElements["Collected Wall Types"]))[0]
    materialOption1["Floors"]["Exterior"] = list(filter(lambda x : x.GetParameters("Type Name")[0].AsString() == exteriorFloorTypeName, collectedElements["Collected Floor Types"]))[0]
    materialOption1["Floors"]["Interior"] = list(filter(lambda x : x.GetParameters("Type Name")[0].AsString() == interiorFloorTypeName, collectedElements["Collected Floor Types"]))[0]

    materialOption2 = MaterialOption.CreateNewOptionSet()
    materialOption2["Walls"]["Exterior"] = list(filter(lambda x : x.GetParameters("Type Name")[0].AsString() == exteriorWallTypeName, collectedElements["Collected Wall Types"]))[0]
    materialOption2["Walls"]["Interior"] = list(filter(lambda x : x.GetParameters("Type Name")[0].AsString() == interiorWallTypeName, collectedElements["Collected Wall Types"]))[0]
    materialOption2["Floors"]["Exterior"] = list(filter(lambda x : x.GetParameters("Type Name")[0].AsString() == exteriorFloorTypeName, collectedElements["Collected Floor Types"]))[0]
    materialOption2["Floors"]["Interior"] = list(filter(lambda x : x.GetParameters("Type Name")[0].AsString() == interiorFloorTypeName, collectedElements["Collected Floor Types"]))[0]

    if selectedOption == 2:
        selectedMaterialOption = materialOption2
    else:
        selectedMaterialOption = materialOption1
    

    # Replacing walls
    replaceWalls(collectedElements, selectedMaterialOption)

    # Replacing Floors
    replaceFloors(collectedElements, selectedMaterialOption)

    # Replacing Doors
    replaceDoors(collectedElements)

    # Close existing project
    openedExistingProject.Close()

def createTopography():

    allowedBoundaryLines = ['Flunkage Line', 'Side Line', 'Front Line', 'Rear Line']
    boundaryLines = FilteredElementCollector(doc).OfClass(CurveElement).ToElements()
    boundaryLinesElements = list(filter(lambda x : x.LineStyle.Name in allowedBoundaryLines, boundaryLines))
    boundaryLines = list(map(lambda x : x.GeometryCurve, boundaryLinesElements))

    pt1 = boundaryLines[0].GetEndPoint(1)
    pt2 = boundaryLines[1].GetEndPoint(1)
    pt3 = boundaryLines[2].GetEndPoint(1)
    pt4 = boundaryLines[3].GetEndPoint(1)
    topoPoints = [pt1, pt2, pt3, pt4]

    # Moving points down by 191mm
    topoPoints = list(map(lambda x : XYZ(x.X, x.Y, x.Z - 0.6266), topoPoints))

    tt = Transaction(doc, "Create Topography")
    tt.Start()
    topography = TopographySurface.Create(doc, topoPoints)

    for line in boundaryLinesElements:
        doc.Delete(line.Id)

    tt.Commit()

# HELPERS

def openProjectByFilePath(projectPath):

    # BIM 360://Show Unit/Ciel V2-DesignOption02_A.rvt
    filePath = ModelPathUtils.ConvertUserVisiblePathToModelPath(projectPath)

    openOptions = OpenOptions()

    # Opening and returning the document we want to import data from
    return app.OpenDocumentFile(filePath, openOptions)

def transferSchedulesFromProject(existingProject):

    # Transferring the schedules as how they are set up to the new project
    existingSchedules = FilteredElementCollector(existingProject).OfClass(ViewSchedule).ToElements()
    existingSchedules = list(filter(lambda x : x.Name.find("Revision Schedule") == -1, existingSchedules))
    existingSchedulesIDs = List[ElementId](list(map(lambda x : x.Id, existingSchedules)))

    tt = Transaction(doc, "Transfer Schedules")
    tt.Start()

    importOptions = CopyPasteOptions()
    importOptions.SetDuplicateTypeNamesHandler(HideAndAcceptDuplicateTypesHandler())
    ElementTransformUtils.CopyElements(existingProject, existingSchedulesIDs, doc, Transform.Identity, importOptions)

    failureOptions = tt.GetFailureHandlingOptions()
    failureOptions.SetFailuresPreprocessor(HidePasteDuplicateTypesPreprocessor())

    tt.Commit(failureOptions)

def transferSheetsFromProject(existingProject):
    
    # Transfering the sheets to the new document
    existingSheets = FilteredElementCollector(existingProject).OfClass(ViewSheet)

    sheetsDataTable = []

    sheetDisciplineSharedParameterGuid = Guid("2fd62d67-c3b3-42aa-b195-6ca33d7e5ad6")
    sheetSetSharedParameterGuid = Guid("e4709840-a4a0-412d-b1e2-5b7f6cfb32dd")
    sheetSeriesSharedParameterGuid = Guid("7dbb64ab-4fe4-4786-a117-b7ff7bbcfd0e")
    drawingTitleSharedParameterGuid = Guid("aaf81681-55ac-473c-8569-75173f4b4ba6")

    for sheet in existingSheets:

        sheetItem = []
        sheetItem.append(sheet.SheetNumber)
        sheetItem.append(sheet.Name)
        sheetItem.append(sheet.get_Parameter(sheetDisciplineSharedParameterGuid).AsString()) # 1-Sheet Discipline
        sheetItem.append(sheet.get_Parameter(sheetSetSharedParameterGuid).AsString()) # 2-Sheet Set
        sheetItem.append(sheet.get_Parameter(sheetSeriesSharedParameterGuid).AsString()) # 3-Sheet Series
        sheetItem.append(sheet.get_Parameter(drawingTitleSharedParameterGuid).AsString()) # 4-Drawing Title

        sheetsDataTable.append(sheetItem)

    sheetsDataTable.sort(key=lambda x : x[0])

    # Accessing the currently available titleblock types in the current document
    titleblockTypes = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TitleBlocks).ToElements()

    # Refactoring the ViewSheet.Create procedure
    def createSheetWithTitleblock(titleblockName):
        
        titleBlockType = list(filter(lambda x : x.GetParameters("Type Name")[0].AsString() == titleblockName, titleblockTypes))[0].Id
        
        return ViewSheet.Create(doc, titleBlockType)

    # Creating the new sheets
    tt = Transaction(doc, "Transfer Sheets")
    tt.Start()

    for sheet in sheetsDataTable:

        try:
            if str(sheet[3]).lower().find("permit") != -1:
                createdSheet = createSheetWithTitleblock("ARCH D1-Building Permit Set")
            
            elif str(sheet[3]).lower().find("construction") != -1:
                createdSheet = createSheetWithTitleblock("ARCH D1-Construction Set")

            else:
                createdSheet = createSheetWithTitleblock("ARCH D1-Building Permit Set")

        except IndexError:
            print("The sheets could not be created. Check that the sheet types are named as the following: [ARCH D1-Building Permit Set, ARCH D1-Construction Set]")
            tt.RollBack()

        # Setting the sheet parameters
        createdSheet.SheetNumber = sheet[0]
        createdSheet.Name = sheet[1]
        createdSheet.get_Parameter(sheetDisciplineSharedParameterGuid).Set(str(sheet[2]))
        createdSheet.get_Parameter(sheetSetSharedParameterGuid).Set(str(sheet[3]))
        createdSheet.get_Parameter(sheetSeriesSharedParameterGuid).Set(str(sheet[4]))
        createdSheet.get_Parameter(drawingTitleSharedParameterGuid).Set(str(sheet[5]))

    tt.Commit()

def replaceDoors(collectedElements):
    #------------------------
    genericModelTypes = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsElementType().ToElements()

    collectedDoorType = list(filter(lambda x : Element.Name.GetValue(x) == "Single Door", genericModelTypes))[0]

    #------------------------

def replaceFloors(collectedElements, materialOptionElements):

    collectedFloorInstances = collectedElements["Collected Floors"]

    # Gets all the floor types within the active document
    floorTypes = collectedElements["Collected Floor Types"]

    collectedInteriorFloorType = list(filter(lambda x : Element.Name.GetValue(x) == 'Generic 150mm', floorTypes))[0]
    selectedRoofFloorType = materialOptionElements["Floors"]["Exterior"]
    selectedInteriorFloorType = materialOptionElements["Floors"]["Interior"]

    # Gets the location of the floors by getting the Max point from their bounding boxes
    collectedFloorsLocation = list(map(lambda x : x.get_BoundingBox(None).Max, collectedFloorInstances))

    # Gets the higher floor element to consider it the roof
    highestPoint = list(map(lambda x : x.Z, collectedFloorsLocation))
    highestPoint = max(highestPoint)

    highestFloor = list(filter(lambda x : x.get_BoundingBox(None).Max.Z == highestPoint, collectedFloorInstances))[0]

    # Removing the highest floor from the floor list as it would be the roof
    collectedFloorInstances.Remove(highestFloor)


    tt = Transaction(doc, "Replace Floors")
    tt.Start()

    # Replacing the roof
    highestFloor.FloorType = selectedRoofFloorType

    # Replacing the interior floors
    for floor in collectedFloorInstances:

        if floor.FloorType.Id == collectedInteriorFloorType.Id:

            floor.FloorType = selectedInteriorFloorType

    tt.Commit()

def replaceRoofs(collectedElements):
    pass

def replaceWalls(collectedElements, materialOptionElements):

    # Replacing the walls based on if they are exterior or interior

    # Collecting wall types from document (Note: from another project in the future)
    wallTypes = collectedElements["Collected Wall Types"]

    collectedExteriorWallType = list(filter(lambda x : Element.Name.GetValue(x) == 'Generic - 200mm', wallTypes))[0]
    collectedInteriorWallType = list(filter(lambda x : Element.Name.GetValue(x) == 'Interior - 135mm Partition (2-hr)', wallTypes))[0]

    selectedExteriorWallType = materialOptionElements["Walls"]["Exterior"]
    selectedInteriorWallType = materialOptionElements["Walls"]["Interior"]


    tt = Transaction(doc, "Replace Walls")
    tt.Start()
    for wall in collectedElements["Collected Walls"]:
        
        if wall.WallType.Id == collectedExteriorWallType.Id:

            # Checking if exterior face of the wall is looking to the outside, and if not, flip it
            # fixWallOrientation(wall)

            # Ensuring that the location line are set to "Core Face: Exterior"
            wall.get_Parameter(BuiltInParameter.WALL_KEY_REF_PARAM).Set(4)

            # Changing to selected type
            wall.WallType = selectedExteriorWallType

        elif wall.WallType.Id == collectedInteriorWallType.Id:
            # Ensuring that the location line are set to "Core Centerline"
            wall.get_Parameter(BuiltInParameter.WALL_KEY_REF_PARAM).Set(1)

            # Changing to selected type
            wall.WallType = selectedInteriorWallType
            
    tt.Commit()

# CLASSES

class HideAndAcceptDuplicateTypesHandler(IDuplicateTypeNamesHandler):

    def OnDuplicateTypeNamesFound(self, args):
        return DuplicateTypeAction.UseDestinationTypes

class HidePasteDuplicateTypesPreprocessor(IFailuresPreprocessor):

    def PreprocessFailures(self, failuresAccessor):

        for failure in failuresAccessor.GetFailureMessages():

            # Delete any "Can't paste duplicate types.  Only non duplicate types will be pasted." warnings
            if failure.GetFailureDefinitionId() == BuiltInFailures.CopyPasterFailures.CannotCopyDuplicates or \
               failure.GetFailureDefinitionId() == BuiltInFailures.CopyPasterFailures.ElementRenamedOnPaste:

                failuresAccessor.DeleteWarning(failure)

        return FailureProcessingResult.ProceedWithCommit
