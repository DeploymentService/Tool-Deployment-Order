#!/usr/bin/env python
import os
import sys
extensionFolder = os.path.abspath(os.path.dirname(os.path.dirname(__file__)))
workingCodeFolder = os.path.abspath(os.path.join(extensionFolder, "code"))
sys.path.append(workingCodeFolder)

# Importing local project libraries
from libraries.forms import *

def main():

    winFormApp.Run(createProjectForm)

    return 0

main()