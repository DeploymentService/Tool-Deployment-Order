#!/usr/bin/env python
import os
import sys
extensionFolder = os.path.abspath(os.path.dirname(os.path.dirname(__file__)))
workingCodeFolder = os.path.abspath(os.path.join(extensionFolder, "code"))
sys.path.append(workingCodeFolder)

# Importing CLR and adding references
import clr
clr.AddReference('System.IO')
clr.AddReference('System.Drawing')
clr.AddReference('System.Reflection')
clr.AddReference('System.Threading')
clr.AddReference('System.Windows.Forms')
import System
import System.Drawing
import System.IO
import System.Windows.Forms

# Importing local project libraries
from libraries.forms import *

def main():

    winFormApp.Run(smartAutoDetailForm)

    return 0

main()