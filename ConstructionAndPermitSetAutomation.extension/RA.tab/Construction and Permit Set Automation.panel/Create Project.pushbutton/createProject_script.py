#!/usr/bin/env python
import os
import sys
extensionFolder = os.path.abspath(os.path.dirname(os.path.dirname(__file__)))
workingCodeFolder = os.path.abspath(os.path.join(extensionFolder, "code"))
sys.path.append(workingCodeFolder)

# Importing local project libraries
from libraries.forms import *
from libraries.ViewCreation import CreateFloorPlans, CreateElevations, CreateSections
from libraries.ViewPlacement import PlaceViewsOnSheets

def main():

    winFormApp.Run(createProjectForm)

    return 0

main()

#TODO
# Organize in a better way the created views
# Creation of schedules
# Check on what views should be dependent based on how much will they vary based on their use
# Create Grids