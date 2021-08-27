#! python3
import os
import sys
currentPath = os.path.abspath(os.path.dirname(os.path.dirname(__file__)))
currentPath = os.path.abspath(os.path.join(currentPath, "Export to GSheets.pushbutton"))
workingCodeFolder = os.path.abspath(os.path.join(currentPath, "resources"))
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
from forms import *

def main():

    try:
        costEstimationForm.run()
    except System.ObjectDisposedException as e:
        print("At this moment due to technical limitation the Cost Estimation script can be run once per Revit session, \
              to use it again close Revit and open it again.\n\n \
              Remember to synchronize before closing.")
    
    return 0

main()