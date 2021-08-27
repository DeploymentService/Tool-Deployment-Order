import os
import sys
from collections import OrderedDict
currentPath = os.path.abspath(os.path.dirname(os.path.dirname(__file__)))
currentPath = os.path.abspath(os.path.join(currentPath, "Export to GSheets.pushbutton"))
resourcesPath = os.path.join(currentPath, "resources")

# Importing CLR and adding references
import clr
clr.AddReference('System.IO')
clr.AddReference('System.Collections')
clr.AddReference('System.Data')
clr.AddReference('System.Drawing')
clr.AddReference('System.Reflection')
clr.AddReference('System.Threading')
clr.AddReference('System.Windows.Forms')

# Imports related to the Cost Estimation App
from costEstimationApp import costEstimationApp

import System
from System.Data import DataTable, DataColumn
from System.Drawing import Font, FontStyle, Size, Color, Point, GraphicsUnit, ContentAlignment, Image, Icon
import System.IO
from System.Windows.Forms import *

from time import sleep

# --- Style and Format Variables ---
fontType = "Arial"
titleFont = Font(fontType, System.Single(18), FontStyle.Bold, GraphicsUnit.Point)
groupLabelFont = Font(fontType, System.Single(14), FontStyle.Bold, GraphicsUnit.Point)
h2Font = Font(fontType, System.Single(14), FontStyle.Regular, GraphicsUnit.Point)
generalFont = Font(fontType, System.Single(10), FontStyle.Regular, GraphicsUnit.Point)
buttonFont = Font(fontType, System.Single(14), FontStyle.Bold, GraphicsUnit.Point)
blackColor = Color.FromArgb(22, 22, 22)
grayColor = Color.FromArgb(60, 60, 60)
whiteColor = Color.FromArgb(240, 240, 240)
pureWhiteColor = Color.FromArgb(255, 255, 255)

# Creating our own form class

# Base Form Class
class BaseForm(Form):

    def __init__(self):

        # Default start position
        self.StartPosition = FormStartPosition.CenterScreen
        
        # Define the caption text
        self.Text = "Create Projects"

        # Icon
        self.Icon = Icon("{0}\RAIcon.ico".format(resourcesPath))

        # --- TITLE ---
        self.titleIcon = PictureBox()
        self.titleIcon.Image = Image.FromFile("{0}\RATitleIcon.ico".format(resourcesPath))
        self.titleIcon.Size = Size(64, 64)
        self.titleIcon.Location = Point(0, 0)

        self.titleText = Label()
        self.titleText.Text = "*REPLACE TITLE HERE*"
        self.titleText.Location = Point(80, 10)
        self.titleText.Size = Size(600, 50)
        self.titleText.TextAlign = ContentAlignment.MiddleLeft
        self.titleText.Font = titleFont

        # Define the background color of our form
        self.BackColor = whiteColor

        # Define the size of the form
        self.ClientSize = Size(600, 600)

        # Grab the caption height
        captionHeight = SystemInformation.CaptionHeight

        # Define the minimum size of my form using the MinimumSize Property
        self.MinimumSize = Size(392, (117 + captionHeight))

        # --- BUTTONS ---
        # Confirm Button
        self.AcceptButton = CommitButton()
        self.AcceptButton.Location = Point(0, 536)
        self.AcceptButton.Text = "Confirm"

        # Cancel Button
        self.CancelButton = CommitButton()
        self.CancelButton.Location = Point(300, 536)
        self.CancelButton.Text = "Cancel"    

        # --- CONTROLS ---
        # Define Control Container
        self.componentsContainer = System.ComponentModel.Container()

        # Add controls
        self.Controls.Add(self.titleIcon)
        self.Controls.Add(self.AcceptButton)
        self.Controls.Add(self.CancelButton)

        # # Blocks the manual resizing of the form
        # self.FormBorderStyle = FormBorderStyle.FixedSingle

        # --- CONTROLS ---
        # Define Control Container
        self.componentsContainer = System.ComponentModel.Container()

        # --- BINDING EVENTS ---
        self.Activated += self.onFormResize
        self.ResizeEnd += self.onFormResize
        self.CancelButton.MouseUp += self.onCancelButtonClick

    # --- EVENTS ---
    def onFormResize(self, sender, args):
        # Getting actual size of view
        clientSize = self.ClientSize

        self.AcceptButton.Location = Point(0, clientSize.Height - self.AcceptButton.Height)
        self.AcceptButton.Size = Size(int(clientSize.Width / 2), 64)

        self.CancelButton.Location = Point(int(clientSize.Width / 2), clientSize.Height - self.CancelButton.Height)
        self.CancelButton.Size = Size(int(clientSize.Width / 2), 64)

    def onCancelButtonClick(self, sender, args):
        self.Hide()
        self.Close()
        self.Dispose()

    def run(self):

        """
            Starts our form object
        """

        # Run the Form
        Application.Run(self)

    def dispose(self):
        """
            Dispose of form object and components container
        """

        self.componentsContainer.Dispose()
        Form.Dispose(self)

# User Forms
class CostEstimationForm(BaseForm):

    def __init__(self):
        BaseForm.__init__(self)

        # --- TITLE ---
        self.titleText.Text = "COST ESTIMATION APP"

        # Define the size of the form
        self.ClientSize = Size(600, 260)

        # --- GROUP: EXPORT INFORMATION ---
        self.exportInformationGroup = CommonGroup()
        self.exportInformationGroup.Text = "Name of Spreadsheet"
        self.exportInformationGroup.Location = Point(30, 75)
        self.exportInformationGroup.Font = groupLabelFont
        
        self.exportedSheetName = CommonTextBox()
        self.exportedSheetName.PlaceholderText = "Name"
        self.exportedSheetName.Location = Point(30, 35)

        # --- GROUP: END ---

        # --- CONTROLS ---
        # Add controls
        self.Controls.Add(self.titleText)

        self.Controls.Add(self.exportInformationGroup)
        self.exportInformationGroup.Controls.Add(self.exportedSheetName)

        # --- BINDING EVENTS ---
        self.Activated += self.onCostEstimationFormResize
        self.ResizeEnd += self.onCostEstimationFormResize
        self.AcceptButton.MouseUp += self.onAcceptButtonClick
    
    # --- EVENTS ---
    def onCostEstimationFormResize(self, sender, args):
        # Getting actual size of view
        clientSize = self.ClientSize

        self.exportInformationGroup.Size = Size(int(clientSize.Width - 60), 50)
        self.exportedSheetName.Size = Size(int(clientSize.Width - 120), 110)

    def onAcceptButtonClick(self, sender, args):
        self.Hide()
        costEstimationApp(self.exportedSheetName.Text)
        self.Close()
        self.Dispose()

# --- Custom Control Classes ---
class CommonGroup(GroupBox):

    def __init__(self):
        
        self.AutoSize = True
        self.Size = Size(540, 50)
        self.Font = groupLabelFont

class CommonButton(Button):

    def __init__(self):
        self.Size = Size(300, 64)
        self.FlatStyle = FlatStyle.Flat
        self.FlatAppearance.BorderSize = 1
        self.FlatAppearance.BorderColor = blackColor
        self.BackColor = blackColor
        self.ForeColor = whiteColor
        self.Font = buttonFont

        # --- BINDING EVENTS ---
        self.MouseEnter += self.onMouseEnter
        self.MouseLeave += self.onMouseLeave

    # EVENTS
    def onMouseEnter(self, sender, args):
        self.BackColor = Color.FromArgb(100, 100, 100)

    def onMouseLeave(self, sender, args):
        self.BackColor = blackColor

class CommitButton(CommonButton):
    def __init__(self):
        CommonButton.__init__(self)

        # --- RESPONSIVENESS ---
        self.Anchor = AnchorStyles.Bottom

class CommonTextBox(TextBox):
    def __init__(self):
        
        self.BackColor = pureWhiteColor
        self.ForeColor = blackColor
        self.Font = generalFont
        self.Size = Size(100, 20)
        self.FlatStyle = FlatStyle.Flat

        self.PlaceholderText = "Search"

# --- Forms Creation ---
costEstimationForm = CostEstimationForm()