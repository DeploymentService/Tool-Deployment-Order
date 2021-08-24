#!/usr/bin/env python
import os
import sys
currentPath = os.path.abspath(os.path.dirname(os.path.dirname(__file__)))

# Python imports
from enum import Enum

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

# Enums
class ViewUse(Enum):
    
    PERMIT = 1
    CONSTRUCTION = 2
    PRESENTATION = 3
    WORKING = 4

class ViewCategory(Enum):

    BUILDING = 1
    FURNITURECASEWORK = 2

# Interface with extra properties
class IViewData:

    def __init__(self):
        self.viewName = ""
        self.viewUses = []
        self.viewCategories = []

class ViewFloorPlanRA(IViewData):

    def __init__(self):
        IViewData.__init__(self)
        self.ViewPlan = ViewPlan.Create()

    def Create():
        pass
        

class View3DRA(IViewData):

    def __init__(self):
        IViewData.__init__(self)
        self.View3D = View3D.CreateIsometric()

class ViewDraftingDetailRA(IViewData):

    def __init__(self):
        IViewData.__init__(self)
        self.ViewDrafting = ViewDrafting.Create()

class ViewSectionRA(IViewData):

    def __init__(self):
        IViewData.__init__(self)
        self.ViewSection = ViewSection.CreateSection()

class ViewSheetRA(IViewData):

    def __init__(self):
        IViewData.__init__(self)
        self.ViewSheet = ViewSheet.Create()
