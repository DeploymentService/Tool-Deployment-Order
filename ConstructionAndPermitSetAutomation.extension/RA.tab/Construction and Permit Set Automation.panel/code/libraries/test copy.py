import os
import sys
currentPath = os.path.abspath(os.path.dirname(os.path.dirname(__file__)))

# Python imports
import math
from time import sleep, time

# Revit API imports
import clr
clr.AddReference("System")
clr.AddReference("RevitAPI")
clr.AddReference('RevitAPIUI')
clr.AddReference("RevitNodes")
from System import Guid
import Revit
clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)
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

# CONSTANTS
GRID_RESOLUTION = 30

# MAIN FUNCTIONS

def PlaceViewsOnSheets():

    # Collecting current sheets and views in the model
    sheets = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()
    views = FilteredElementCollector(doc).OfClass(View).ToElements()
    views = list(filter(lambda x : x.IsTemplate != True, views))
    views = list(filter(lambda x : type(x) != ViewSheet, views))
    views = list(filter(lambda x : type(x) != ViewSchedule, views))
    views = list(filter(lambda x : type(x) != View, views))
    
    # Declaring variables for future use
    placedViews = []

    tt = Transaction(doc, "Place Views in Sheets")
    tt.Start()
    start_time = time()
    # Main process
    iterator = 0
    limit = 15
    for sheet in sheets:

        # origin = XYZ(0, 0, 0)
        # originEnd = origin + XYZ(-0.5, -0.5, 0)

        # doc.Create.NewDetailCurve(sheet, Line.CreateBound(origin, originEnd))

        # Collecting data from the sheet number of the current sheet to then compare to specific data in each view
        fullSheetNumber = sheet.SheetNumber
        sheetNumberSymbol = fullSheetNumber.split(".")[0]
        sheetNumberNumber = fullSheetNumber.split(".")[1]

        # Getting the category of drawing that should be in the sheet (Floor Plans, Elevations, Sections, Details, etc)
        sheetDrawingTitleParameter = sheet.get_Parameter(Guid("aaf81681-55ac-473c-8569-75173f4b4ba6")).AsString().lower()

        workspaceCells = getWorkspace(sheet, specialCase = None)

        # Collecting list of views that could be added to the current sheet
        viewsForSheet = []
        for view in views:

            if view.UniqueId in placedViews:
                continue

            viewDrawingTypeParameter = view.get_Parameter(Guid("f60b1ea4-247a-4906-a9a6-2e370f6ad50d"))

            if viewDrawingTypeParameter == None or viewDrawingTypeParameter.AsString() == None:
                continue
            
            if viewDrawingTypeParameter.AsString().find(sheetNumberSymbol) != -1:
                viewsForSheet.append(view)

        # Executes filtering of views for particular context
        viewsForSheet = executeContextualFilter(sheet, viewsForSheet)

        # Getting the views and getting a value to determine how similar their name are 
        # to the name of the current sheet in order to decide exactly which view to place
        viewsWithSimilarityScore = []
        for view in viewsForSheet:

            if " - " in view.Name:
                viewName = view.Name.split(" - ")[2]
            
            else:
                viewName = view.Name

            similarityScore = compareSheetAndViewNames(viewName, sheet.Name)

            viewsWithSimilarityScore.append([view, similarityScore])
        
        viewsWithSimilarityScore.sort(key=lambda x : x[1], reverse=True)

        print(sheet.SheetNumber, sheet.Name)
        # for view in viewsWithSimilarityScore:
        #     print(view[0].Name, view[1])

        if viewsWithSimilarityScore:

            # Specifies the amount of views that can be included in a sheet based on its drawings category
            if sheetDrawingTitleParameter == "floor plans":
                limitOfViews = 1
            
            elif sheetDrawingTitleParameter == "elevations":
                limitOfViews = 2
            
            elif sheetDrawingTitleParameter == "sections":
                limitOfViews = 10
            
            elif sheetDrawingTitleParameter == "details":
                limitOfViews = 9999
            
            else:
                limitOfViews = 9999
            
            counter = 0
            for view in viewsWithSimilarityScore:
                
                selectedView = view[0]
                print(selectedView.Name)

                # IMPORTANT --- PROCEDURE OF VIEW PLACEMENT IN SHEET
                viewport = Viewport.Create(doc, sheet.Id, selectedView.Id, XYZ(0, 0, 0))
                viewportSize = getViewportDimensions(viewport, setNewDimensions=True, cell=workspaceCells.cellList[0])
                
                # Determining in which cell this viewport will be located
                locationCell = workspaceCells.checkForAvailableCells(viewportSize)
                
                # If a cell for locating the view was not located due to a lack of cells, continue with the next sheet
                if not locationCell:
                    doc.Delete(viewport.Id)
                    break

                # print("Original cell count: {}".format(len(workspaceCells.cellList)))

                viewport.SetBoxCenter(locationCell.Location)
                viewportOutline = viewport.GetBoxOutline()
                outlineOffset = 0.05
                outlineMax = viewportOutline.MaximumPoint + XYZ(outlineOffset, outlineOffset, 0)
                outlineMin = viewportOutline.MinimumPoint - XYZ(outlineOffset, outlineOffset, 0)

                drawDummyViewportBoundary(viewport, sheet)

                # Filters out the cells that were occupied in this iteration
                workspaceCells.cellList = list(filter(lambda x : x.Location.X > outlineMax.X or \
                                                                 x.Location.Y > outlineMax.Y or \
                                                                 x.Location.X < outlineMin.X or \
                                                                 x.Location.Y < outlineMin.Y, workspaceCells.cellList))
         
                workspaceCells.RemoveOccupiedCells()

                # Stores the placed view in a list to be checked further in order to not repeat view placements
                placedViews.append(selectedView.UniqueId)
                # print("Updated cell count: {}".format(len(workspaceCells.cellList)))

                counter += 1
                if counter == limitOfViews:
                    break
        
        print("--------------------------------")
        # if workspaceCells:
        #     drawDummyLines(workspaceCells, sheet)
            # for cell in workspaceCells.cellList:
            #     print("-")
            #     print(cell.ID, cell.Size, cell.Location)
            #     print(cell.BottomNeighbors, cell.TopNeighbors, cell.LeftNeighbors, cell.RightNeighbors)
            #     print(cell.SpaceAtBottom, cell.SpaceAtTop, cell.SpaceAtLeft, cell.SpaceAtRight)
            
        iterator += 1
        if iterator > limit:
            break

    tt.Commit()
    print("--- {} seconds ---".format(time() - start_time))

# HELPERS

def compareSheetAndViewNames(viewName, sheetName):

    score = 0

    viewName = viewName.lower()
    sheetName = sheetName.lower()

    viewNameWords = viewName.split(" ")

    for word in viewNameWords:

        pointsPerWord = len(word)

        if sheetName.find(word) != -1:

            score += pointsPerWord

    return score

def createGridWithCells(workspaceBoundaryLines, specialCase = None):

    lowResolutionCases = ["floor plans", "elevations", "sections"]
    if not specialCase:
        gridResolution = GRID_RESOLUTION

        cellSize = workspaceBoundaryLines[0].Length / gridResolution

        pointsAtBoundary = []
        for line in workspaceBoundaryLines:

            pointsAtLine = []

            for i in range(0, gridResolution + 1):
                length = cellSize * i

                if length > line.Length:
                    break

                pointAtLength = line.Evaluate(length, False)
                pointsAtLine.append(pointAtLength)

            pointsAtBoundary.append(pointsAtLine)

        gridLines_X = []
        for pairsOfPoints in zip(pointsAtBoundary[1], pointsAtBoundary[3]):

            line = Line.CreateBound(pairsOfPoints[0], pairsOfPoints[1])
            gridLines_X.append(line)

        gridLines_Y = []
        for pairsOfPoints in zip(pointsAtBoundary[0], pointsAtBoundary[2]):

            line = Line.CreateBound(pairsOfPoints[0], pairsOfPoints[1])
            gridLines_Y.append(line)

        gridPoints = []
        for lineY in gridLines_Y:

            for lineX in gridLines_X:
                results = clr.Reference[IntersectionResultArray]()
                comparisonResult = lineY.Intersect(lineX, results)
                intersection = results.Item[0]

                gridPoints.append(intersection.XYZPoint)
        
    else:

        specialCase = specialCase.lower()

        if specialCase in lowResolutionCases:
            gridResolution = 5

            if specialCase == "floor plans":
                pass

            elif specialCase == "elevations":
                pass

            elif specialCase == "sections":
                pass

    # Rectangular cells on each grid point
    cellID = 0
    cells = []
    for point in gridPoints:

        cellID += 1

        cell = Cell()
        cell.ID = cellID
        cell.Location = point
        cell.Size = cellSize
        cell.CountNeighbors(gridPoints)

        cells.append(cell)

    workspace = Workspace()
    workspace.cellList = cells

    return workspace

def drawDummyLines(workspace, sheet):

    for cell in workspace.cellList:
        endpoint = cell.Location + XYZ(0.03, 0.03, 0)
        line = Line.CreateBound(cell.Location, endpoint)

        doc.Create.NewDetailCurve(sheet, line)

def drawDummyViewportBoundary(viewport, sheet):
    # Getting the viewport bounding box
    viewportOutline = viewport.GetBoxOutline()
    
    outlineMax = viewportOutline.MaximumPoint
    outlineMin = viewportOutline.MinimumPoint

    viewportWidth = outlineMax.X - outlineMin.X
    viewportHeight = outlineMax.Y - outlineMin.Y

    p1 = outlineMax
    p2 = outlineMax - XYZ(0, viewportHeight, 0)
    p3 = p2 - XYZ(viewportWidth, 0, 0)
    p4 = p3 + XYZ(0, viewportHeight, 0)

    l1 = Line.CreateBound(p1, p2)
    l2 = Line.CreateBound(p2, p3)
    l3 = Line.CreateBound(p3, p4)
    l4 = Line.CreateBound(p4, p1)

    lines = [l1, l2, l3, l4]

    for line in lines:
        doc.Create.NewDetailCurve(sheet, line)

def getViewportDimensions(viewport, setNewDimensions = False, cell = None):

    # Getting the viewport bounding box
    viewportOutline = viewport.GetBoxOutline()
    
    outlineMax = viewportOutline.MaximumPoint
    outlineMin = viewportOutline.MinimumPoint

    viewportWidth = outlineMax.X - outlineMin.X
    viewportHeight = outlineMax.Y - outlineMin.Y

    if setNewDimensions:

        oldviewportWidth = viewportWidth
        oldviewportHeight = viewportHeight
        viewportWidth = ceilToMultiple(viewportWidth, cell.Size)
        viewportHeight = ceilToMultiple(viewportHeight, cell.Size)

        XTranslation = (viewportWidth - oldviewportWidth) / 2
        YTranslation = (viewportHeight - oldviewportHeight) / 2

        outlineMax = outlineMax + XYZ(XTranslation, YTranslation, 0)
        outlineMin = outlineMin - XYZ(XTranslation, YTranslation, 0)

    return [viewportWidth, viewportHeight]

def executeContextualFilter(sheet, viewsForSheet):
    fullSheetNumber = sheet.SheetNumber
    sheetNumberSymbol = fullSheetNumber.split(".")[0]
    sheetNumberNumber = fullSheetNumber.split(".")[1]

    # Temporary measure to filter Mechanical and Plumbing plans
    if "mechanical" in sheet.Name.lower():
        viewsForSheet = list(filter(lambda x : x.Name.find("M") == 0, viewsForSheet))

    elif "plumbing" in sheet.Name.lower():
        viewsForSheet = list(filter(lambda x : x.Name.find("P") == 0, viewsForSheet))

    # TEMPORARY ---------------------------------------------------------------------------
    # Filters out the building sections if the current sheet is for wall sections
    if sheet.get_Parameter(Guid("7dbb64ab-4fe4-4786-a117-b7ff7bbcfd0e")).AsString().lower().find("wall") != -1:

        viewsForSheet = list(filter(lambda x : x.get_Parameter(Guid("f60b1ea4-247a-4906-a9a6-2e370f6ad50d")).AsString().lower().find("wall section") != -1, viewsForSheet))
    else:
        viewsForSheet = list(filter(lambda x : x.get_Parameter(Guid("f60b1ea4-247a-4906-a9a6-2e370f6ad50d")).AsString().lower().find("wall section") == -1, viewsForSheet))

    # ------------------------------------------------------------------------------------

    if sheet.Name.lower().find("crawl") != -1:
        viewsForSheet = list(filter(lambda x : x.Name.lower().find("crawl space") != -1, viewsForSheet))

    return viewsForSheet

def getWorkspace(sheet, specialCase = None):
    # Collect the titleblock element in the current sheet, and get from it the Workspace extensions
    titleblockElement = FilteredElementCollector(doc, sheet.Id).OfClass(FamilyInstance).ToElements()[0]
    
    # Titleblock element's boundingbox and its Max point
    titleblockBB = titleblockElement.get_BoundingBox(sheet)
    titleBlockBBMax = titleblockBB.Max
    
    # Getting the sheet margin and titleblock width parameter in order to translate the Max point to be the origin of the workspace
    sheetMarginParameterValue = titleblockElement.get_Parameter(Guid("0e291b2e-d37c-4f60-b45d-b62a71fd48cb")).AsDouble()
    sheetTitleBlockWidthParameterValue = titleblockElement.get_Parameter(Guid("22e3d3f1-8cdd-4c07-8508-b1dd48d01c4b")).AsDouble()

    translatedMax = titleBlockBBMax - XYZ(sheetMarginParameterValue, sheetMarginParameterValue, 0)
    translatedMax = translatedMax - XYZ(sheetTitleBlockWidthParameterValue, 0, 0)

    # Workspace creation
    workspaceHeight = titleblockElement.get_Parameter(Guid("1790db12-36f2-4cb9-a049-e75e78c0b27f")).AsDouble()
    workspaceWidth = titleblockElement.get_Parameter(Guid("3f5dfac0-c5b4-4ae5-ad8b-94c8a62fa0a8")).AsDouble()

    # Workspace's corners
    # Specify the inset for the boundary
    boundaryInsetValue = (workspaceWidth / GRID_RESOLUTION) / 2

    # Points
    wsp_P1 = translatedMax - XYZ(boundaryInsetValue, boundaryInsetValue, 0)
    wsp_P2 = wsp_P1 - XYZ(workspaceWidth - (boundaryInsetValue * 2), 0, 0)
    wsp_P3 = wsp_P2 - XYZ(0, workspaceHeight - (boundaryInsetValue * 2), 0)
    wsp_P4 = wsp_P3 + XYZ(workspaceWidth - (boundaryInsetValue * 2), 0, 0)

    # Workspace's boundary lines
    wsp_TopLine = Line.CreateBound(wsp_P1, wsp_P2)
    wsp_LeftLine = Line.CreateBound(wsp_P2, wsp_P3)
    wsp_BottomLine = Line.CreateBound(wsp_P4, wsp_P3)
    wsp_RightLine = Line.CreateBound(wsp_P1, wsp_P4)

    wspLines = [wsp_TopLine, wsp_LeftLine, wsp_BottomLine, wsp_RightLine]

    workspaceCells = createGridWithCells(wspLines, specialCase)

    # for cell in workspaceCells:
        # print("-------------------------")
        # print(cell.ID, cell.Location, cell.Size)
        # print(cell.BottomNeighbors, cell.TopNeighbors, cell.LeftNeighbors, cell.RightNeighbors)
        # print(cell.SpaceAtBottom, cell.SpaceAtTop, cell.SpaceAtLeft, cell.SpaceAtRight)

        # doc.Create.NewDetailCurve(sheet, Line.CreateBound(cell.Location, cell.Location + XYZ(0.01, 0.01, 0)))

    return workspaceCells

def levenshtein_ratio_and_distance(s, t, ratio_calc = False):
    """ levenshtein_ratio_and_distance:
        Calculates levenshtein distance between two strings.
        If ratio_calc = True, the function computes the
        levenshtein distance ratio of similarity between two strings
        For all i and j, distance[i,j] will contain the Levenshtein
        distance between the first i characters of s and the
        first j characters of t

        --- Extracted from Datacamp: https://www.datacamp.com/community/tutorials/fuzzy-string-python ---
    """
    # Helper function
    def implementedNPZeros(dimensions):

        emptyList = []

        for i in range(0, dimensions[0]):

            row = []

            for j in range(0, dimensions[1]):

                row.append(0)

            emptyList.append(row)

        return emptyList
        
    # Initialize matrix of zeros
    rows = len(s)+1
    cols = len(t)+1
    distance = implementedNPZeros((rows,cols))

    # Populate matrix of zeros with the indeces of each character of both strings
    for i in range(1, rows):
        for k in range(1,cols):
            distance[i][0] = i
            distance[0][k] = k

    # Iterate over the matrix to compute the cost of deletions,insertions and/or substitutions    
    for col in range(1, cols):
        for row in range(1, rows):
            if s[row-1] == t[col-1]:
                cost = 0 # If the characters are the same in the two strings in a given position [i,j] then the cost is 0
            else:
                # In order to align the results with those of the Python Levenshtein package, if we choose to calculate the ratio
                # the cost of a substitution is 2. If we calculate just distance, then the cost of a substitution is 1.
                if ratio_calc == True:
                    cost = 2
                else:
                    cost = 1
            distance[row][col] = min(distance[row-1][col] + 1,      # Cost of deletions
                                 distance[row][col-1] + 1,          # Cost of insertions
                                 distance[row-1][col-1] + cost)     # Cost of substitutions
    if ratio_calc == True:
        # Computation of the Levenshtein Distance Ratio
        Ratio = ((len(s)+len(t)) - distance[row][col]) / (len(s)+len(t))
        return Ratio
    else:
        # print(distance) # Uncomment if you want to see the matrix showing how the algorithm computes the cost of deletions,
        # insertions and/or substitutions
        # This is the minimum number of edits needed to convert string a to string b
        return distance[row][col]

def ceilToMultiple(number, multiple):

    result = multiple * math.ceil(number / multiple)

    return result

# CLASSES
class Cell():

    def __init__(self):
        self.ID = None
        self.Location = None
        self.Size = None
    
    # Returns the count of cells in each direction from this cell
    def CountNeighbors(self, gridPoints):

        if not self.Location:
            print("The cell must have a specified location in order to calculate its neighbors")

        self.BottomNeighbors = len(list(filter(lambda x : x.Y < self.Location.Y and \
                                                          x.X == self.Location.X, gridPoints)))
        
        self.TopNeighbors = len(list(filter(lambda x : x.Y > self.Location.Y and \
                                                       x.X == self.Location.X, gridPoints)))

        self.LeftNeighbors = len(list(filter(lambda x : x.X < self.Location.X and \
                                                        x.Y == self.Location.Y, gridPoints)))

        self.RightNeighbors = len(list(filter(lambda x : x.X > self.Location.X and \
                                                         x.Y == self.Location.Y, gridPoints)))

        # Calculating the net space in each direction
        self.SpaceAtBottom = (self.BottomNeighbors * self.Size) + self.Size / 2
        self.SpaceAtTop = (self.TopNeighbors * self.Size) + self.Size / 2
        self.SpaceAtLeft = (self.LeftNeighbors * self.Size) + self.Size / 2
        self.SpaceAtRight = (self.RightNeighbors * self.Size) + self.Size / 2

class Workspace():

    def __init__(self):
        self.cellList = []

    def RemoveOccupiedCells(self):
        gridPoints = list(map(lambda x : x.Location, self.cellList))

        for cell in self.cellList:

            cell.CountNeighbors(gridPoints)

    def checkForAvailableCells(self, requestedArea):

        locationCell = None

        requestedCells = list(map(lambda x : math.ceil(x / self.cellList[0].Size), requestedArea))
        requestedCells = requestedCells[0] * requestedCells[1]

        def drawLine(cell):

            origin = cell.Location
            endpoint = origin + XYZ(0.05, 0.05, 0)

            doc.Create.NewDetailCurve(doc.GetElement(ElementId(16111034)), Line.CreateBound(origin, endpoint))

        # Running first filter. Returns all the possible cells that could be locations
        filteredCellList = list(filter(lambda x : x.SpaceAtBottom >= (requestedArea[1] / 2) and \
                                                  x.SpaceAtTop >= (requestedArea[1] / 2) and \
                                                  x.SpaceAtLeft >= (requestedArea[0] / 2) and \
                                                  x.SpaceAtRight >= (requestedArea[0] / 2), self.cellList))

        # Finding the max and min point of the area around each cell
        for cell in filteredCellList:

            # Get the hypotenuse vector from the spaces at each direction values
            maxVectorValue = XYZ(cell.SpaceAtRight, cell.SpaceAtTop, 0)
            minVectorValue = XYZ((requestedArea[0] / 2), (requestedArea[1] / 2), 0)

            areaMaxPoint = cell.Location + maxVectorValue
            areaMinPoint = cell.Location - minVectorValue

            areaCells = list(filter(lambda x : x.Location.X < areaMaxPoint.X and \
                                               x.Location.Y < areaMaxPoint.Y and \
                                               x.Location.X > areaMinPoint.X and \
                                               x.Location.Y > areaMinPoint.Y, self.cellList))

            amountOfCells = len(areaCells)

            # doc.Create.NewDetailCurve(
            #     doc.GetElement(ElementId(16111034)),
            #     Line.CreateBound(
            #         areaMaxPoint, areaMaxPoint - XYZ(0.1, 0.3, 0)
            #     )
            # )

            # doc.Create.NewDetailCurve(
            #     doc.GetElement(ElementId(16111034)),
            #     Line.CreateBound(
            #         areaMinPoint, areaMinPoint - XYZ(0.1, 0.3, 0)
            #     )
            # )

            # for cell1 in areaCells:

            #     # doc.Create.NewDetailCurve(
            #     #     doc.GetElement(ElementId(16111034)),
            #     #     Line.CreateBound(
            #     #         cell1.Location, cell1.Location - XYZ(0.1, 0.3, 0)
            #     #     )
            #     # )
            print(amountOfCells, requestedCells)
            if amountOfCells >= requestedCells:

                locationCell = cell

            else:

                locationCell = None

            return locationCell


# TODO
### Create list with Sheets-Views pairs establishing what views ideally will be included in each sheet. [[ViewSheet, [Views]], [ViewSheet, [Views]], [ViewSheet, [Views]], ...]
    ### Establish the Views to Sheet relation based on comparing the names of the sheet and the views, and getting the most probable option to place in the view. Levenshtein distance algorithm?
# Iterate through each sheet of this list and execute the main procedure to place the views on it
    # The main procedure involves:
        ### Defining the workspace where we can put the views
        ### Start placing views from the top right to bottom, in a column direction
        # If we filled the sheet and there are still more views left to place
            # Create new sheets to place the remaining views
        
        # If the view was already placed in other sheet, pass
        
# Delete any sheet that does not have views