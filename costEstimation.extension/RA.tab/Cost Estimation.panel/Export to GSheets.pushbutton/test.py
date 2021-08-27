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

for path in sys.path:
    print(path)

# Revit API imports
clr.AddReference("System.Drawing")
import System.Drawing
clr.AddReference("RevitAPI")
clr.AddReference('RevitAPIUI')
clr.AddReference("RevitNodes")
clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
import Autodesk
import Autodesk.Revit
from Autodesk.Revit.UI import *
import Autodesk.Revit.UI.Selection
from Autodesk.Revit.DB import * 
from Autodesk.Revit.DB import Transaction, Structure, Architecture

# Accessing current Revit Document
doc = __revit__.ActiveUIDocument.Document

def convertUnits(value, listOfUnits):

    # Mapping all the values in the list of units to lower case
    listOfUnits = list(map(lambda x : x.lower() if isinstance(x, str) else x, listOfUnits))

    if value != None:

        if isinstance(value, int) or isinstance(value, float):

            if "ft" in listOfUnits:
                print(1)
                value = UnitUtils.ConvertFromInternalUnits(value, DisplayUnitType.DUT_DECIMAL_FEET)
                return value
            
            elif "mm" in listOfUnits:
                print(2)
                value = UnitUtils.ConvertFromInternalUnits(value, DisplayUnitType.DUT_MILLIMETERS)
                return value

            elif "sqf" in listOfUnits:
                print(3)
                value = UnitUtils.ConvertFromInternalUnits(value, DisplayUnitType.DUT_SQUARE_FEET)
                return value

            elif "sqm" in listOfUnits:
                print(4)
                value = UnitUtils.ConvertFromInternalUnits(value, DisplayUnitType.DUT_SQUARE_METERS)
                return value

            else:
                return value

    return value

print(convertUnits(20, ["mm"]))