import os
import sys
currentPath = os.path.abspath(os.path.dirname(os.path.dirname(__file__)))

# Python imports
import math

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

doc = __revit__.ActiveUIDocument.Document
app = doc.Application

def CreateFloorPlans():

    ViewDataList = getViewDataList()

    buildingBB = getBuildingBoundingBox()

    # Collecting levels within the project
    levels = FilteredElementCollector(doc).OfClass(Level).WhereElementIsNotElementType()

    # Collecting the ViewFamilyTypeIds related to the Floor Plan
    fpViewFamilyTypes = FilteredElementCollector(doc).OfClass(ViewFamilyType)
    fpViewFamilyTypes = list(filter(lambda x : x.GetParameters("Type Name")[0].AsString() == "Floor Plan", fpViewFamilyTypes))
    fpViewFamilyTypesId = list(map(lambda x : x.Id, fpViewFamilyTypes))[0]

    # Creating Floor Plan Views
    tt = Transaction(doc, "Create Floor Plans")
    tt.Start()

    for level in levels:

        levelId = level.Id

        for disciplineName, discipline in ViewDataList["Disciplines"].items():

            for use in ViewDataList["View Uses"]:

                for category in ViewDataList["View Categories"]:

                    floorPlan = ViewPlan.Create(doc, fpViewFamilyTypesId, levelId)
                    floorPlan.Name = "{} - {} - {} - {}".format(disciplineName[:3], use, category, level.Name)
                    floorPlan.Discipline = discipline
                    floorPlan.GetParameters("View Use")[0].Set(use)
                    floorPlan.GetParameters("View Category")[0].Set(category)
                    floorPlan.CropBox = buildingBB
                    floorPlan.CropBoxActive = True

    # Creating Site Plan
    GFLevel = list(filter(lambda x : x.Elevation == 0, levels))[0]

    siteFloorPlan = ViewPlan.Create(doc, fpViewFamilyTypesId, GFLevel.Id)
    sitePlanDisciplineName = "Arc"
    sitePlanDiscipline = ViewDiscipline.Architectural
    sitePlanUse =  "Permit"
    sitePlanCategory =  "Building"
    siteFloorPlan.Name = "{} - {} - {} - {}".format(sitePlanDisciplineName,
                                                sitePlanUse,
                                                sitePlanCategory,
                                                "Site")
    siteFloorPlan.Discipline = sitePlanDiscipline
    siteFloorPlan.GetParameters("View Use")[0].Set(sitePlanUse)
    siteFloorPlan.GetParameters("View Category")[0].Set(sitePlanCategory)
    siteFloorPlan.CropBox = buildingBB
    siteFloorPlan.CropBoxActive = True

    tt.Commit()

def CreateElevations():

    # Collecting levels within the project
    levels = FilteredElementCollector(doc).OfClass(Level).WhereElementIsNotElementType()
    levelZero = list(filter(lambda x : x.Elevation == 0, levels))[0]

    # Collecting the ViewFamilyTypeIds related to the Floor Plan
    ViewFamilyTypes = FilteredElementCollector(doc).OfClass(ViewFamilyType)
    fpViewFamilyTypes = list(filter(lambda x : x.GetParameters("Type Name")[0].AsString() == "Floor Plan", ViewFamilyTypes))
    fpViewFamilyTypesId = list(map(lambda x : x.Id, fpViewFamilyTypes))[0]

    elevationViewFamilyTypes = list(filter(lambda x : x.GetParameters("Type Name")[0].AsString() == "Building Elevation", ViewFamilyTypes))
    elevationViewFamilyTypesId = list(map(lambda x : x.Id, elevationViewFamilyTypes))[0]

    ViewDataList = getViewDataList()

    buildingBB = getBuildingBoundingBox()
    buildingBBLength = buildingBB.Max.X - buildingBB.Min.X
    buildingBBWidth = buildingBB.Max.Y - buildingBB.Min.Y
    buildingBBHeight = buildingBB.Max.Z - buildingBB.Min.Z


    # Creating rectangle in horizontal plane from which we are going to get the middle point of each edge to place our elevation marker

    # Points
    rctPT1 = XYZ(buildingBB.Min.X, buildingBB.Min.Y, 0) # Normalized example 0, 0, 0
    rctPT2 = XYZ(buildingBB.Min.X, buildingBB.Min.Y + buildingBBWidth, 0) # Normalized example 0, 1, 0
    rctPT3 = XYZ(rctPT2.X + buildingBBLength, rctPT2.Y, 0) # Normalized example 1, 1, 0
    rctPT4 = XYZ(rctPT3.X, rctPT2.Y - buildingBBWidth, 0) # Normalized example 1, 0, 0

    # Lines of rectangle
    rctL1 = Line.CreateBound(rctPT1, rctPT2)
    rctL2 = Line.CreateBound(rctPT2, rctPT3)
    rctL3 = Line.CreateBound(rctPT3, rctPT4)
    rctL4 = Line.CreateBound(rctPT4, rctPT1)

    rctLinesList = [rctL1, rctL2, rctL3, rctL4]

    rctCurveArray = CurveArray()
    for line in rctLinesList:
        rctCurveArray.Append(line)

    # Elevation Markers position
    EMPositions = []
    for line in rctLinesList:

        EMPositions.append(line.Evaluate(0.5, True))

    tt = Transaction(doc, "Create Elevations")
    tt.Start()

    # Creating a temporal ViewPlan for the ElevationMarkers creation
    temporalViewPlan = ViewPlan.Create(doc, fpViewFamilyTypesId, levelZero.Id)
    temporalViewPlan.Name = "Temporal ViewPlan"

    for disciplineName, discipline in ViewDataList["Disciplines"].items():

        if disciplineName == "Architectural":

            for use in ViewDataList["View Uses"]:

                for category in ViewDataList["View Categories"]:

                    for position in EMPositions:

                        rotationValueAndPosition = getMarkerRotationValueAndPosition(position, EMPositions)

                        # Creating ElevationMarker
                        elevationMarker = ElevationMarker.CreateElevationMarker(doc, elevationViewFamilyTypesId, position, 50)

                        # Creating Elevation at ElevationMarker
                        elevationView = elevationMarker.CreateElevation(doc, temporalViewPlan.Id, 0)
                        elevationView.Name = getElevationName(rotationValueAndPosition[1], disciplineName, use, category)
                        elevationView.Discipline = discipline
                        elevationView.GetParameters("View Use")[0].Set(use)
                        elevationView.GetParameters("View Category")[0].Set(category)

                        # Rotating elevation marker to look at the building
                        # Getting rotation axis
                        p2 = XYZ(position.X, position.Y, position.Z + 1)

                        rotationAxis = Line.CreateBound(position, p2)

                        ElementTransformUtils.RotateElement(doc, elevationMarker.Id, rotationAxis, rotationValueAndPosition[0] * math.pi / 180)

                        # Setting the far clip offset for each elevation based on its cardinal location (N, S, E, W, etc)
                        setFarClipOffset(elevationView, rotationValueAndPosition[1], buildingBB)

                        # Adjusting the crop region
                        crsm = elevationView.GetCropRegionShapeManager()

                        if rotationValueAndPosition[0] == 0 or rotationValueAndPosition[0] == 180:

                            csrmRectangle = drawElevationCropRegionRectangle(position, buildingBB, "Y")

                            curveLoop = CurveLoop()
                            for line in csrmRectangle:
                                curveLoop.Append(line)

                            crsm.SetCropShape(curveLoop)

                        elif rotationValueAndPosition[0] == 90 or rotationValueAndPosition[0] == 270:

                            csrmRectangle = drawElevationCropRegionRectangle(position, buildingBB, "X")

                            curveLoop = CurveLoop()
                            for line in csrmRectangle:
                                curveLoop.Append(line)

                            crsm.SetCropShape(curveLoop)

    # Deleting temporary view plan
    doc.Delete(temporalViewPlan.Id)

    tt.Commit()

def CreateSections():

    pass

def drawElevationCropRegionRectangle(position, buildingBB, option):

    buildingBB = buildingBB
    buildingBBLength = buildingBB.Max.X - buildingBB.Min.X
    buildingBBWidth = buildingBB.Max.Y - buildingBB.Min.Y
    buildingBBHeight = buildingBB.Max.Z - buildingBB.Min.Z

    if option == "X":
        
        p1 = XYZ(position.X - (buildingBBLength / 2), position.Y, buildingBB.Min.Z)
        p2 = XYZ(p1.X + buildingBBLength, p1.Y, p1.Z)
        p3 = XYZ(p2.X, p2.Y, p2.Z + buildingBBHeight)
        p4 = XYZ(p1.X, p1.Y, p1.Z + buildingBBHeight)

        # Create Rectangle shape
        # Lines of rectangle
        l1 = Line.CreateBound(p1, p2)
        l2 = Line.CreateBound(p2, p3)
        l3 = Line.CreateBound(p3, p4)
        l4 = Line.CreateBound(p4, p1)

        lines = [l1, l2, l3, l4]

        return lines

    if option == "Y":

        p1 = XYZ(position.X, position.Y - (buildingBBWidth / 2), buildingBB.Min.Z)
        p2 = XYZ(p1.X, p1.Y + buildingBBWidth, p1.Z)
        p3 = XYZ(p2.X, p2.Y, p2.Z + buildingBBHeight)
        p4 = XYZ(p1.X, p1.Y, p1.Z + buildingBBHeight)

        # Create Rectangle shape
        # Lines of rectangle
        l1 = Line.CreateBound(p1, p2)
        l2 = Line.CreateBound(p2, p3)
        l3 = Line.CreateBound(p3, p4)
        l4 = Line.CreateBound(p4, p1)

        lines = [l1, l2, l3, l4]

        return lines

def getBuildingBoundingBox():

    floorsInProject = FilteredElementCollector(doc).OfClass(Floor).WhereElementIsNotElementType().ToElements()

    minX = []
    minY = []
    minZ = []
    maxX = []
    maxY = []
    maxZ = []

    for floor in floorsInProject:

        floorBB = floor.get_BoundingBox(None)

        # Collecting all the maximum and minimum points in each direction to then obtain the real minimum and maximum one in each
        minX.append(floorBB.Min.X)
        minY.append(floorBB.Min.Y)
        minZ.append(floorBB.Min.Z)
        maxX.append(floorBB.Max.X)
        maxY.append(floorBB.Max.Y)
        maxZ.append(floorBB.Max.Z)

    # Creating the real minimum and maximum points of the project
    projectMax = XYZ(max(maxX), max(maxY), max(maxZ))
    projectMin = XYZ(min(minX), min(minY), min(minZ))

    # Offset that will be added to increase the building boundingbox size
    generalOffset = XYZ(15, 15, 15)

    projectMax += generalOffset
    projectMin -= generalOffset

    # Creating the boundingbox for the project
    buildingBB = BoundingBoxXYZ()
    buildingBB.Max = projectMax
    buildingBB.Min = projectMin

    return buildingBB

def getElevationName(orientation, disciplineName, use, category):

    if orientation == "N":

        elevationName = "North"

    if orientation == "S":

        elevationName = "South"

    if orientation == "E":

        elevationName = "East"

    if orientation == "W":

        elevationName = "West"

    if orientation == "NW":

        elevationName = "Northwest"

    if orientation == "NE":

        elevationName = "Northeast"

    if orientation == "SW":

        elevationName = "Southwest"

    if orientation == "SE":

        elevationName = "Southeast"

    return "{} - {} - {} - {}".format(disciplineName, use, category, elevationName)

def getMarkerRotationValueAndPosition(point, listOfPoints):

    listOfXs = list(map(lambda x : x.X, listOfPoints))
    listOfYs = list(map(lambda x : x.Y, listOfPoints))

    # If the point is the one most to the South, return 270
    if point.Y == min(listOfYs):
        return [270, "S"]

    # If the point is the one most to the West, return 180
    if point.X == min(listOfXs):
        return [180, "W"]

    # If the point is the one most to the North, return 90
    if point.Y == max(listOfYs):

        return [90, "N"]

    # If the point is the one most to the East, return 0
    if point.X == max(listOfXs):
        return [0, "E"]

    # If the point is the one most to the SW, return 315
    if point.X == min(listOfXs) and point.Y == min(listOfYs):
        return [315, "SW"]

    # If the point is the one most to the NW, return 225
    if point.X == min(listOfXs) and point.Y == max(listOfYs):
        return [225, "NW"]

    # If the point is the one most to the NE, return 135
    if point.X == max(listOfXs) and point.Y == max(listOfYs):
        return [135, "NE"]

    # If the point is the one most to the SE, return 45
    if point.X == max(listOfXs) and point.Y == min(listOfYs):
        return [45, "SE"]

def getViewDataList():

    return {"Disciplines" : {"Architectural" : ViewDiscipline.Architectural, 
                             "Structural" : ViewDiscipline.Structural, 
                             "Mechanical" : ViewDiscipline.Mechanical, 
                             "Electrical" : ViewDiscipline.Electrical, 
                             "Plumbing" : ViewDiscipline.Plumbing, 
                             "Coordination" : ViewDiscipline.Coordination},

            "View Uses" : ["Permit", "Construction", "Presentation", "Working View"],
            "View Categories" : ["Building"]}

def setFarClipOffset(view, elevationPosition, buildingBB):

    buildingBB = buildingBB
    buildingBBLength = buildingBB.Max.X - buildingBB.Min.X
    buildingBBWidth = buildingBB.Max.Y - buildingBB.Min.Y

    # Getting the Far Clip Offset parameter
    fcOffset = view.get_Parameter(BuiltInParameter.VIEWER_BOUND_OFFSET_FAR)

    if elevationPosition == "N":
        return fcOffset.Set(buildingBBWidth)

    if elevationPosition == "S":
        return fcOffset.Set(buildingBBWidth)

    if elevationPosition == "E":
        return fcOffset.Set(buildingBBLength)

    if elevationPosition == "W":
        return fcOffset.Set(buildingBBLength)

#TODO
# Create 3 sections
# Assign view templates to views
# Creation of schedules
# Check on what views should be dependent based on how much will they vary based on their use