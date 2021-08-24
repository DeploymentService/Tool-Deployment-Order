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

from System.Collections.Generic import List

from materialOptions import MaterialOption

doc = __revit__.ActiveUIDocument.Document
app = doc.Application

def importFromExistingProject(projectPath):

    # BIM 360://Show Unit/Ciel V2-DesignOption02_A.rvt
    filePath = ModelPathUtils.ConvertUserVisiblePathToModelPath(projectPath)

    openOptions = OpenOptions()
    # openOptions.OpenForeignOption = OpenForeignOption.Prompt

    # Opening the document we want to import data from
    openedExistingProject = app.OpenDocumentFile(filePath, openOptions)

    # Setting shared parameters file
    app.SharedParametersFilename = os.path.join(currentPath, "CIEL_Shared_Parameters.txt")
    # sharedParametersFile = app.OpenSharedParameterFile()

    ### WIP ----------------------------
    # for group in sharedParametersFile.Groups:

    #     sharedParameters = group.Definitions
    #     print("Group: {}".format(group.Name))

    #     for parameter in sharedParameters:

    #         if group.Name == "Identity Data":
                
    #             if parameter.Name == "Category":
    ### ----------------------------------

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
                               "Windows",
                               "Furniture",
                               "Plumbing Fixtures",
                               "Electrical Equipment",
                               "Specialty Equipment"]
                               
    # Controlling the categories we want to transfer                               
    filteredExistingFamiliesSymbolsIDs = []
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


    # Transfering the elements to the new document
    tt = Transaction(doc, "Transfer Elements")
    tt.Start()

    importOptions = CopyPasteOptions()
    ElementTransformUtils.CopyElements(openedExistingProject, existingElementsIDs, doc, Transform.Identity, importOptions)
    
    tt.Commit()

    # Transfering the sheets to the new document
    existingSheets = FilteredElementCollector(openedExistingProject).OfClass(ViewSheet)

    sheetsDataTable = []

    for sheet in existingSheets:

        sheetItem = []

        sheetItem.append(sheet.SheetNumber)
        sheetItem.append(sheet.Name)
        sheetItem.append(sheet.GetParameters("Sheet Set")[0].AsString())
        sheetItem.append(sheet.GetParameters("DisciplineInSheetSet")[0].AsString())
        sheetItem.append(sheet.GetParameters("SheetIsInConstructionSet")[0].AsString())
        sheetItem.append(sheet.GetParameters("SheetIsInPermittingSet")[0].AsString())

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

        if str(sheet[2]).lower().find("permitting") != -1:
            createdSheet = createSheetWithTitleblock("CIEL_PERMIT_A1")
            createdSheet.GetParameters("SheetIsInConstructionSet")[0].Set(False)
            createdSheet.GetParameters("SheetIsInPermittingSet")[0].Set(True)

        elif str(sheet[2]).lower().find("construction") != -1:
            createdSheet = createSheetWithTitleblock("CIL_A1 metric Construction")
            createdSheet.GetParameters("SheetIsInConstructionSet")[0].Set(True)
            createdSheet.GetParameters("SheetIsInPermittingSet")[0].Set(False)

        else:
            createdSheet = createSheetWithTitleblock("CIEL_Working_A1")

        # Setting the sheet parameters
        createdSheet.SheetNumber = sheet[0]
        createdSheet.Name = sheet[1]
        createdSheet.GetParameters("Sheet Set")[0].Set(str(sheet[2]))
        createdSheet.GetParameters("DisciplineInSheetSet")[0].Set(str(sheet[3]))

    tt.Commit()

    # Close existing project
    openedExistingProject.Close()

    #TODO
    # Check on importing window and door types

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

def replaceConceptualElements(collectedElements, selectedOption):

    # Creating Material Options
    materialOption1 = MaterialOption.CreateNewOptionSet()
    materialOption1["Walls"]["Exterior"] = list(filter(lambda x : Element.Name.GetValue(x) == '2x4 Stud Wall @16" O.C.', collectedElements["Collected Wall Types"]))[0]
    materialOption1["Walls"]["Interior"] = list(filter(lambda x : Element.Name.GetValue(x) == 'Bamcore- Interior 6" (door)', collectedElements["Collected Wall Types"]))[0]
    materialOption1["Floors"]["Exterior"] = list(filter(lambda x : Element.Name.GetValue(x) == 'Terrace 6.5"', collectedElements["Collected Floor Types"]))[0]
    materialOption1["Floors"]["Interior"] = list(filter(lambda x : Element.Name.GetValue(x) == 'Generic 191mm - Wood', collectedElements["Collected Floor Types"]))[0]

    materialOption2 = MaterialOption.CreateNewOptionSet()
    materialOption2["Walls"]["Exterior"] = list(filter(lambda x : Element.Name.GetValue(x) == '2x6 SPF Stud Wall @24" O.C.', collectedElements["Collected Wall Types"]))[0]
    materialOption2["Walls"]["Interior"] = list(filter(lambda x : Element.Name.GetValue(x) == 'Bamcore- Interior 6" (door)', collectedElements["Collected Wall Types"]))[0]
    materialOption2["Floors"]["Exterior"] = list(filter(lambda x : Element.Name.GetValue(x) == 'Generic 150mm', collectedElements["Collected Floor Types"]))[0]
    materialOption2["Floors"]["Interior"] = list(filter(lambda x : Element.Name.GetValue(x) == 'Generic 200mm', collectedElements["Collected Floor Types"]))[0]


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

def applyMaterials(setOfMaterials):
    #TODO
    pass

def createTopography():

    allowedBoundaryLines = ['Flunkage Line', 'Side Line', 'Front Line', 'Rear Line']
    boundaryLines = FilteredElementCollector(doc).OfClass(CurveElement).ToElements()
    boundaryLines = list(filter(lambda x : x.LineStyle.Name in allowedBoundaryLines, boundaryLines))
    boundaryLines = list(map(lambda x : x.GeometryCurve, boundaryLines))

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
    tt.Commit()

# --- Main replacement functions ---

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

# --- HELPERS ---

######WIP
def fixWallOrientation(wallElement):

    # Getting the exterior face geometrical data
    exteriorFace = getWallExteriorFaceGeometry(wallElement)
    
    centerPointUV = UV(- wallElement.Location.Curve.Length / 2, wallElement.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).AsDouble() / 2)
    
    # centerPointUV = UV(-3, -1)
    centerPoint = exteriorFace.Evaluate(centerPointUV)
    faceNormal = exteriorFace.ComputeNormal(centerPointUV)

    # Walls Reference Intersector
    view3D = doc.GetElement(ElementId(347549))
    refIntersector = ReferenceIntersector(ElementClassFilter(Wall), FindReferenceTarget.Face, view3D)

    # Looking for walls in front of the exterior wall
    firstIntersectedWall = refIntersector.FindNearest(centerPoint, faceNormal)

    # If there is a wall in the path, get the geometric data
    if firstIntersectedWall:

        # List to store more walls if found
        moreIntersectedWalls = []

        firstIntersectedWallReference = firstIntersectedWall.GetReference()

        # Get the hit point in order to trace two new paths from it
        hitPoint = firstIntersectedWallReference.GlobalPoint

        # First new path
        firstPath = faceNormal.Negate().Normalize()

        rotation = Transform.CreateRotation(XYZ(hitPoint.X, hitPoint.Y, hitPoint.Z + 1), 45)
        firstPath = rotation.OfVector(firstPath)

        referenceFromPath = refIntersector.FindNearest(hitPoint, firstPath)

        if referenceFromPath:
            moreIntersectedWalls.append(referenceFromPath)

        # Second new path
        secondPath = faceNormal.Negate().Normalize()

        rotation = Transform.CreateRotation(XYZ(hitPoint.X, hitPoint.Y, hitPoint.Z + 1), -45)
        secondPath = rotation.OfVector(secondPath)

        referenceFromPath = refIntersector.FindNearest(hitPoint, secondPath)

        if referenceFromPath:
            moreIntersectedWalls.append(referenceFromPath)
        
        # If there are walls to both sides of the first intersected wall, flip the original wall
        if len(moreIntersectedWalls) == 2:
            wallElement.Flip()

    return 0

def getWallExteriorFaceGeometry(wallInstance):

    exteriorFaceReference = HostObjectUtils.GetSideFaces(wallInstance, ShellLayerType.Exterior)[0]
    exteriorFace = doc.GetElement(exteriorFaceReference).GetGeometryObjectFromReference(exteriorFaceReference)

    return exteriorFace

def getWallTypeName(wallType):

    return Element.Name.GetValue(wallType)