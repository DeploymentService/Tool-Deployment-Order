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

doc = __revit__.ActiveUIDocument.Document
app = doc.Application

# COMMON USE VARIABLES

# Retrieving view templates
viewTemplates = FilteredElementCollector(doc).OfClass(View).ToElements()
viewTemplates = list(filter(lambda x : x.IsTemplate, viewTemplates))

# MAIN FUNCTIONS

def CreateFloorPlans():

    ViewDataList = getViewDataList()

    buildingBB = getBuildingBoundingBox()

    # Retrieving the view families for the floor plans
    floorplanViewFamilyTypes = getViewFamilyTypes("Floor Plans")

    # Collecting levels within the project
    levels = FilteredElementCollector(doc).OfClass(Level).WhereElementIsNotElementType()

    # Creating Floor Plan Views
    tt = Transaction(doc, "Create Floor Plans")
    tt.Start()

    for level in levels:

        levelId = level.Id

        for disciplineName, discipline in ViewDataList["Disciplines"].items():

            for use in ViewDataList["View Uses"]:

                for category in ViewDataList["View Categories"]:

                    # Skip this iteration as we do not want MEP plans for roof
                    if "roof" in level.Name.lower():

                        if disciplineName in ["Mechanical", "Electrical", "Plumbing"]:

                            continue

                    # Select view type to be assigned to the view being created
                    viewType = assignViewType(disciplineName, use, floorplanViewFamilyTypes)

                    floorPlan = ViewPlan.Create(doc, viewType.Id, levelId)
                    floorPlan.Name = "{} - {} - {} PLAN".format(disciplineName[:1], use.upper(), level.Name)
                    floorPlan.CropBox = buildingBB
                    floorPlan.CropBoxActive = True
                    floorPlan.CropBoxVisible = False
                    floorPlan.ViewTemplateId = viewType.get_Parameter(BuiltInParameter.DEFAULT_VIEW_TEMPLATE).AsElementId()

    # Creating Site Plan
    GFLevel = list(filter(lambda x : x.Elevation == 0, levels))[0]
    sitePlanViewType = list(filter(lambda x : x.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString().find("PS") != -1, floorplanViewFamilyTypes))[0]

    siteFloorPlan = ViewPlan.Create(doc, sitePlanViewType.Id, GFLevel.Id)
    sitePlanDisciplineName = "A"
    sitePlanUse =  "Permit"
    sitePlanCategory =  "Building"
    siteFloorPlan.Name = "{} - {} - SITE PLAN".format(sitePlanDisciplineName,
                                                sitePlanUse.upper())
    siteFloorPlan.CropBox = buildingBB
    siteFloorPlan.CropBoxActive = True

    tt.Commit()

def CreateElevations():

    # Retrieving the view families for the elevations
    elevationViewFamilyTypes = getViewFamilyTypes("Elevations")

    # Collecting levels within the project
    levels = FilteredElementCollector(doc).OfClass(Level).WhereElementIsNotElementType()
    levelZero = list(filter(lambda x : x.Elevation == 0, levels))[0]

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
    temporalViewTemplate = getViewFamilyTypes("Floor Plans")[0].Id
    temporalViewPlan = ViewPlan.Create(doc, temporalViewTemplate, levelZero.Id)
    temporalViewPlan.Name = "Temporal ViewPlan"

    for disciplineName, discipline in ViewDataList["Disciplines"].items():

        if disciplineName == "Architectural":

            for use in ViewDataList["View Uses"]:

                for category in ViewDataList["View Categories"]:

                    for position in EMPositions:

                        # Select view type to be assigned to the view being created
                        viewType = assignViewType(disciplineName, use, elevationViewFamilyTypes)

                        rotationValueAndPosition = getMarkerRotationValueAndPosition(position, EMPositions)

                        # Creating ElevationMarker
                        elevationMarker = ElevationMarker.CreateElevationMarker(doc, viewType.Id, position, 50)

                        # Creating Elevation at ElevationMarker
                        elevationView = elevationMarker.CreateElevation(doc, temporalViewPlan.Id, 0)
                        elevationView.Name = getElevationName(rotationValueAndPosition[1], disciplineName[:1], use)


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

                        elevationView.CropBoxVisible = False

    # Deleting temporary view plan
    doc.Delete(temporalViewPlan.Id)

    tt.Commit()

def CreateSections():

    # Retrieving the view families for the sections
    sectionViewFamilyTypes = getViewFamilyTypes("Sections")

    # Retrieving building bounding box
    buildingBB = getBuildingBoundingBox()
    buildingBBLength = buildingBB.Max.X - buildingBB.Min.X
    buildingBBWidth = buildingBB.Max.Y - buildingBB.Min.Y

    # Get building bounding box for diagonal for section
    sectionBuildingBB = getBuildingBoundingBox(withOffsetAround=False)

    # Creating a diagonal that connect the Min and Max points of the bounding box
    bbMin = sectionBuildingBB.Min
    bbMax = sectionBuildingBB.Max
    mainDiagonal = Line.CreateBound(bbMin, bbMax)

    # Sections' locations
    sectionLocations = []

    param = 0.3
    sectionALocation = mainDiagonal.Evaluate(param, True)
    sectionALocation = XYZ(sectionALocation.X, sectionALocation.Y, bbMin.Z)
    sectionLocations.append([sectionALocation, param])

    param = 0.4
    sectionBLocation = mainDiagonal.Evaluate(param, True)
    sectionBLocation = XYZ(sectionBLocation.X, sectionBLocation.Y, bbMin.Z)
    sectionLocations.append([sectionBLocation, param])

    param = 0.7
    sectionCLocation = mainDiagonal.Evaluate(param, True)
    sectionCLocation = XYZ(sectionCLocation.X, sectionCLocation.Y, bbMin.Z)
    sectionLocations.append([sectionCLocation, param])

    # Condition to make sure that there is going to be just one section throughout the larger dimension of the building
    if buildingBBLength > buildingBBWidth:
        directionValue = 1

    else:
        directionValue = 0
    
    tt = Transaction(doc, "Create Sections")
    tt.Start()

    centerYCoordinate = mainDiagonal.Evaluate(0.5, True).Y
    centerXCoordinate = mainDiagonal.Evaluate(0.5, True).X
    sectionNumber = 1
    for location in sectionLocations:

        parameterAtPoint = location[1]
        location = location[0]

        # If the directionValue is 0, convert it to 2, as we will decide the direction of the section based on if the current index is even or odd
        if directionValue == 0:
            directionValue = 2

        if directionValue % 2 == 0:
            # Overriding the X coordinate in the location in order to center it to the building
            location = XYZ(centerXCoordinate, location.Y, location.Z)

            transform = Transform.Identity
            transform.Origin = location
            transform.BasisX = XYZ(1, 0, 0)
            transform.BasisY = XYZ(0, 0, 1)
            transform.BasisZ = XYZ(0, -1, 0)

            # Building Depth from section location
            buildingDepth = location.Y - bbMin.Y + 10 # 10 as a extra space beyond the building depth

            # Bounding Box for section
            sectionBB = BoundingBoxXYZ()
            sectionBB.Transform = transform
            sectionBB.Min = XYZ(-buildingBBLength / 2, bbMin.Z, 0)
            sectionBB.Max = XYZ(buildingBBLength / 2, bbMax.Z + 5, buildingDepth)

        else:
            # Overriding the Y coordinate in the location in order to center it to the building
            location = XYZ(location.X, centerYCoordinate, location.Z)

            transform = Transform.Identity
            transform.Origin = location
            transform.BasisX = XYZ(0, -1, 0)
            transform.BasisY = XYZ(0, 0, 1)
            transform.BasisZ = XYZ(-1, 0, 0)

            # Building Depth from section location
            buildingDepth = location.X - bbMin.X + 10 # 10 as a extra space beyond the building depth

            # Bounding Box for section
            sectionBB = BoundingBoxXYZ()
            sectionBB.Transform = transform
            sectionBB.Min = XYZ(-buildingBBWidth / 2, bbMin.Z, 0)
            sectionBB.Max = XYZ(buildingBBWidth / 2, bbMax.Z + 5, buildingDepth)

        ViewDataList = getViewDataList()

        try:
            
            for disciplineName, discipline in ViewDataList["Disciplines"].items():

                if disciplineName == "Architectural":

                    for use in ViewDataList["View Uses"]:

                        for category in ViewDataList["View Categories"]:

                            # Select view type to be assigned to the view being created
                            viewType = assignViewType(disciplineName, use, sectionViewFamilyTypes)

                            viewSection = ViewSection.CreateSection(doc, viewType.Id, sectionBB)
                            viewSection.CropBoxVisible = False
                            viewSection.Name = "{} - {} - SECTION {}".format(disciplineName[:1], use.upper(), sectionNumber)

        except Exception as e:
            print(e)

        sectionNumber += 1
        directionValue += 1

    tt.Commit()

# HELPERS

def assignViewType(disciplineName, use, listOfViewTypes):

    def checkForDiscipline(viewTemplateSymbol):

        if disciplineName == "Structural":

            innerViewType = list(filter(lambda x : x.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString().find(viewTemplateSymbol) != -1 and \
                                                   x.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString().lower().find("struc") != -1, listOfViewTypes))[0]
        
        elif disciplineName == "Mechanical":

            innerViewType = list(filter(lambda x : x.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString().find(viewTemplateSymbol) != -1 and \
                                                   x.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString().lower().find("mech") != -1, listOfViewTypes))[0]

        elif disciplineName == "Electrical":

            innerViewType = list(filter(lambda x : x.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString().find(viewTemplateSymbol) != -1 and \
                                                   x.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString().lower().find("pow") != -1, listOfViewTypes))[0]

        elif disciplineName == "Plumbing":

            innerViewType = list(filter(lambda x : x.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString().find(viewTemplateSymbol) != -1 and \
                                                   x.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString().lower().find("mech") != -1, listOfViewTypes))[0]

        else:

            innerViewType = list(filter(lambda x : x.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString().find(viewTemplateSymbol) != -1 and \
                                                   x.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString().lower().find("struc") == -1 and \
                                                   x.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString().lower().find("mech") == -1 and \
                                                   x.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString().lower().find("interior") == -1 and \
                                                   x.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString().lower().find("basement") == -1 and \
                                                   x.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString().lower().find("pow") == -1, listOfViewTypes))[0]
        return innerViewType
    
    try:

        if use.lower().find("construction") != -1:
            
            viewType = checkForDiscipline("CS")

        elif use.lower().find("permit") != -1:

            viewType = checkForDiscipline("PS")

        elif use.lower().find("working") != -1:

            viewType = checkForDiscipline("WK")
        
        elif use.lower().find("presentation") != -1:

            viewType = checkForDiscipline("PR")

    except Exception as e:
 
        viewType = list(filter(lambda x : x.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString().find("Floor Plan") != -1 or \
                                          x.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString().lower().find("elevation") != -1 or \
                                          x.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString().lower().find("section") != -1 or \
                                          x.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString().lower().find("detail") != -1, listOfViewTypes))[0]

    return viewType

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

def getBuildingBoundingBox(withOffsetAround=True):

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

    if withOffsetAround: 
        # Offset that will be added to increase the building boundingbox size
        generalOffset = XYZ(5, 5, 5)

        projectMax += generalOffset
        projectMin -= generalOffset

    # Creating the boundingbox for the project
    buildingBB = BoundingBoxXYZ()
    buildingBB.Max = projectMax
    buildingBB.Min = projectMin

    return buildingBB

def getElevationName(orientation, disciplineName, use):

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

    return "{} - {} - {} ELEVATION".format(disciplineName, use.upper(), elevationName.upper())

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
                             "Plumbing" : ViewDiscipline.Plumbing},
                            #  "Coordination" : ViewDiscipline.Coordination},

            "View Uses" : ["Permit", "Construction", "Presentation", "Working View"],
            "View Categories" : ["Building"]}

def getViewFamilyTypes(viewTypeCategory):

    viewTypes = FilteredElementCollector(doc).OfClass(ViewFamilyType)

    validSelections = ["Floor Plans",
                       "Elevations",
                       "Sections",
                       "Ceiling Plans",
                       "Details"]

    if viewTypeCategory in validSelections:
        
        if validSelections[0].find(viewTypeCategory) != -1:
            viewTypes = list(filter(lambda x : x.ViewFamily == ViewFamily.FloorPlan, viewTypes))
        
        elif validSelections[1].find(viewTypeCategory) != -1:
            viewTypes = list(filter(lambda x : x.ViewFamily == ViewFamily.Elevation, viewTypes))

        elif validSelections[2].find(viewTypeCategory) != -1:
            viewTypes = list(filter(lambda x : x.ViewFamily == ViewFamily.Section, viewTypes))

        elif validSelections[3].find(viewTypeCategory) != -1:
            viewTypes = list(filter(lambda x : x.ViewFamily == ViewFamily.CeilingPlan, viewTypes))

        elif validSelections[4].find(viewTypeCategory) != -1:
            viewTypes = list(filter(lambda x : x.ViewFamily == ViewFamily.Detail, viewTypes))
    
    else:
        print("No correct family type was written as an argument for the function that collects the ViewFamilyTypes in the code")
        return None

    return viewTypes

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
