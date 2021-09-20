#!/usr/bin/env python
import os
import sys
from collections import OrderedDict
currentPath = os.path.abspath(os.path.dirname(os.path.dirname(__file__)))
resourcesPath = os.path.join(currentPath, "resources")
detailLibraryResourcesPath = os.path.join(resourcesPath, "detailCreatorForm\ItemsPreferences\DetailLibrary")

# Importing CLR and adding references
import clr
clr.AddReference('System.IO')
clr.AddReference('System.Collections')
clr.AddReference('System.Data')
clr.AddReference('System.Drawing')
clr.AddReference('System.Reflection')
clr.AddReference('System.Threading')
clr.AddReference('System.Windows.Forms')
clr.AddReference("RevitAPI")
clr.AddReference('RevitAPIUI')
clr.AddReference("RevitNodes")
from Autodesk.Revit.UI import *
from Autodesk.Revit.DB import *

# Custom Imports
from ViewCreation import CreateFloorPlans, CreateElevations, CreateSections
from ViewPlacement import PlaceViewsOnSheets

# Imports related to each pushbutton/form
from createProject_Functions import collectConceptualElements, replaceConceptualElements, createTopography, importFromExistingProject

import System
from System.Data import DataTable, DataColumn
from System.Drawing import Font, FontStyle, Size, Color, Point, GraphicsUnit, ContentAlignment, Image, Icon
import System.IO
from System.Windows.Forms import *

from time import sleep

doc = __revit__.ActiveUIDocument.Document
app = doc.Application

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
        self.AcceptButton.Size = Size(clientSize.Width / 2, 64)

        self.CancelButton.Location = Point((clientSize.Width / 2), clientSize.Height - self.CancelButton.Height)
        self.CancelButton.Size = Size(clientSize.Width / 2, 64)

    def onCancelButtonClick(self, sender, args):
        self.Hide()
        self.Close()

    def Run(self):

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
        System.Windows.Forms.Form.Dispose(self)

# User Forms
class CreateProjectForm(BaseForm):

    def __init__(self):
        BaseForm.__init__(self)

        # --- TITLE ---
        self.titleText.Text = "CREATE PROJECT"

        # --- GROUP: SELECT PROJECT ---
        self.selectProjectGroup = CommonGroup()
        self.selectProjectGroup.Text = "Create From Project"
        self.selectProjectGroup.Location = Point(30, 75)
        self.selectProjectGroup.Font = groupLabelFont

        # COMBOBOX
        # Data for the combobox
        projectSelectionData = ["Default", "V1 - Lazarus", "V2", "V3", "V4", "V5", "V6", "V7", "V8", "V9", "V10"]

        self.projectSelectionCB = CommonComboBox(projectSelectionData)
        self.projectSelectionCB.Location = Point(30, 45)
        self.projectSelectionCB.Size = Size(480, 50)
        
        # Default selection
        self.projectSelectionCB.SelectedIndex = 0
        # --- GROUP: END ---

        # --- GROUP: PROJECT OPTIONS
        self.projectOptionsGroup = CommonGroup()
        self.projectOptionsGroup.Text = "Project Options"
        self.projectOptionsGroup.Location = Point(30, 200)
        self.projectOptionsGroup.Font = groupLabelFont

        # --- CHECKEDLISTBOX ---
        self.projectOptionsChListBox = CommonCheckedListBox()
        self.projectOptionsChListBox.Location = Point(30, 45)
        self.projectOptionsChListBox.Size = Size(480, 110)
        # self.projectOptionsChListBox.Items.Add("Import/Override Model", True)
        self.projectOptionsChListBox.Items.Add("Create Views", True)
        self.projectOptionsChListBox.Items.Add("Import Sheets", True)
        self.projectOptionsChListBox.Items.Add("Place Views on Sheets", True)
        # self.projectOptionsChListBox.Items.Add("Import Annotations", True)
        self.projectOptionsChListBox.Items.Add("Import Products", True)

        # # Hiding the checklistbox by default
        # if self.projectSelectionCB.SelectedText != projectSelectionData[0]:
        #     self.projectOptionsGroup.Hide()
        # --- GROUP: END ---

        # --- CONTROLS ---
        # Add controls
        self.Controls.Add(self.titleText)

        self.Controls.Add(self.selectProjectGroup)
        self.selectProjectGroup.Controls.Add(self.projectSelectionCB)

        self.Controls.Add(self.projectOptionsGroup)
        self.projectOptionsGroup.Controls.Add(self.projectOptionsChListBox)

        # --- BINDING EVENTS ---
        # self.projectSelectionCB.SelectionChangeCommitted += self.onComboBoxChange
        self.ResizeEnd += self.onCreateProjectFormResize
        self.AcceptButton.MouseUp += self.onAcceptButtonClick
    
    # --- EVENTS ---
    # def onComboBoxChange(self, sender, args):
    #     if self.projectSelectionCB.SelectedIndex != 0:
    #         self.projectOptionsGroup.Show()

    #     else:
    #         self.projectOptionsGroup.Hide()

    def onCreateProjectFormResize(self, sender, args):
        # Getting actual size of view
        clientSize = self.ClientSize

        self.selectProjectGroup.Size = Size(clientSize.Width - 60, 50)
        self.projectSelectionCB.Size = Size(clientSize.Width - 120, 110)

        self.projectOptionsGroup.Size = Size(clientSize.Width - 60, 50)
        self.projectOptionsChListBox.Size = Size(clientSize.Width - 120, 110)

    def onAcceptButtonClick(self, sender, args):
        self.Hide()
        # CreateProjectForm1 = CreateProjectForm_ChooseOption()
        # CreateProjectForm1.ShowDialog()

        tg = TransactionGroup(doc, "Create Project")
        tg.Start()
        filePath = r"C:\Users\Ricardo Salas\Desktop\Temporal\2021-ON-001-LEMBITU-ARCH_detached.rvt"

        # Import Options
        importOptionsFromUI = {"Import Products" : self.projectOptionsChListBox.GetItemChecked(3),
                               "Import Sheets" : self.projectOptionsChListBox.GetItemChecked(1)}

        # Base Creation
        importFromExistingProject(filePath, importOptionsFromUI)
        createTopography()
        
        # Replacing elements from output from AutoLayout with standards from existing project
        collectedConceptualElements = collectConceptualElements()
        replaceConceptualElements(filePath, collectedConceptualElements, 1)

        # View Creation
        if self.projectOptionsChListBox.GetItemChecked(0):
            CreateFloorPlans()
            CreateElevations()
            CreateSections()

        # Placing views on sheets
        if self.projectOptionsChListBox.GetItemChecked(2):
            PlaceViewsOnSheets()

        self.Close()
        tg.Commit()

class CreateProjectForm_ChooseOption(BaseForm):

    def __init__(self):
        BaseForm.__init__(self)
        
        self.titleText.Text = "CREATE PROJECT - CHOOSE OPTION"
        self.CancelButton.Text = "Back"

        # Initializing option choice
        self.selectedOption = None

        # --- BUTTONS ---
        # Option 1 Button creation and Properties
        self.option1Button = BuildingOptionButton()
        self.option1Button.Size = Size(self.ClientSize.Width / 2, self.ClientSize.Height - self.titleIcon.Height - self.AcceptButton.Height)
        self.option1Button.Location = Point(0, self.titleIcon.Height)

        self.option1Button.Image = Image.FromFile(r"{0}\buildingOption_ - 3D View - Script Option 1.png".format(resourcesPath))
        # self.option1Button.Image.Height = 128

        # Option 2 Button creation and Properties
        self.option2Button = BuildingOptionButton()
        self.option2Button.Size = Size(self.ClientSize.Width / 2, self.ClientSize.Height - self.titleIcon.Height - self.AcceptButton.Height)
        self.option2Button.Location = Point(self.ClientSize.Width / 2, self.titleIcon.Height)
        self.option2Button.Image = Image.FromFile(r"{0}\buildingOption_ - 3D View - Script Option 2.png".format(resourcesPath))

        # --- CONTROLS ---
        # Add controls
        self.Controls.Add(self.titleText)

        self.Controls.Add(self.option1Button)
        self.Controls.Add(self.option2Button)

        # --- BINDING EVENTS ---
        self.option1Button.MouseUp += self.option1Button_onMouseUp
        self.option2Button.MouseUp += self.option2Button_onMouseUp
        self.CancelButton.MouseUp += self.onBackButtonClick
        self.AcceptButton.MouseUp += self.onAcceptButtonClick

        # Blocks the manual resizing of the form
        self.FormBorderStyle = FormBorderStyle.FixedSingle

    def option1Button_onMouseUp(self, sender, args):
        self.option2Button.MouseUp += self.option2Button.onMouseUp
        self.option2Button.MouseEnter += self.option2Button.onMouseEnter
        self.option2Button.MouseLeave += self.option2Button.onMouseLeave
        self.option2Button.BackColor = grayColor
        self.option2Button.ForeColor = blackColor

        self.selectedOption = 1

    def option2Button_onMouseUp(self, sender, args):
        self.option1Button.MouseUp += self.option1Button.onMouseUp
        self.option1Button.MouseEnter += self.option1Button.onMouseEnter
        self.option1Button.MouseLeave += self.option1Button.onMouseLeave
        self.option1Button.BackColor = grayColor
        self.option1Button.ForeColor = blackColor

        self.selectedOption = 2

    def onBackButtonClick(self, sender, args):
        
        self.Hide()
        mainForm = CreateProjectForm()
        mainForm.ShowDialog()
        self.Close()

    def onAcceptButtonClick(self, sender, args):
        self.Hide()

        if self.selectedOption:

            # filePath = r"C:\Users\Ricardo Salas\Desktop\Temporal\2021-ON-001-LEMBITU-ARCH_detached.rvt"

            # # Base Creation
            # importFromExistingProject(filePath)
            # createTopography()
            
            # # Replacing elements from output from AutoLayout with standards from existing project
            # collectedConceptualElements = collectConceptualElements()
            # replaceConceptualElements(filePath, collectedConceptualElements, self.selectedOption)

            # # View Creation
            # CreateFloorPlans()
            # CreateElevations()
            # CreateSections()

            # # Placing views on sheets
            # PlaceViewsOnSheets()

            self.Close()
        else:
            TaskDialog.Show("Error", "Please run again the script and select a material option.")
            self.Close()

class CreateDocumentsForm(BaseForm):

    def __init__(self):
        BaseForm.__init__(self)

        # --- TITLE ---
        self.titleText.Text = "CREATE DOCUMENTS"

        # Resizing form
        self.ClientSize = Size(self.ClientSize.Width + 400, self.ClientSize.Height + 300)

        # --- GROUP: CREATE/OVERRIDE SHEETS ---
        self.createSheetsGroup = CommonGroup()
        self.createSheetsGroup.Text = "Select Sheets to Create/Override"
        self.createSheetsGroup.Location = Point(30, 75)
        self.createSheetsGroup.Font = groupLabelFont

        # DATATABLE
        self.sheetsDataTable = DataTable("Sheets")
        self.sheetsDataTable.Columns.Add(DataColumn("Select", bool))
        self.sheetsDataTable.Columns.Add(DataColumn("Sheet Name", str))
        self.sheetsDataTable.Columns.Add(DataColumn("Sheet Type", str))
        self.sheetsDataTable.Columns.Add(DataColumn("Status", str))

        # Populating rows
        sheetList = OrderedDict()
        sheetList["A0.0 - COVER SHEET"] = ["RA_Coordination_A1 : 1/50", "TO OVERRIDE"]
        sheetList["A1.0 - LOT GRADING PLAN"] = ["RA_Permit_A1 : 1/50", "TO OVERRIDE"]
        sheetList["A1.3 - BASEMENT PLAN"] = ["RA_Permit_A1 : 1/50", "NEW"]
        sheetList["A1.4 - GROUND FLOOR PLAN"] = ["RA_Permit_A1 : 1/50", "TO OVERRIDE"]
        sheetList["A1.5 - FIRST FLOOR PLAN"] = ["RA_Permit_A1 : 1/50", "TO OVERRIDE"]
        sheetList["A1.6 - ROOFTOP PLAN"] = ["RA_Permit_A1 : 1/50", "NEW"]
        sheetList["A3.1 - SECTIONS"] = ["RA_Permit_A1 : 1/50", "NEW"]
        sheetList["A3.2 - SECTIONS"] = ["RA_Permit_A1 : 1/50", "NEW"]
        sheetList["A3.3 - SECTIONS"] = ["RA_Permit_A1 : 1/50", "NEW"]

        for sheet, status in sheetList.items():
            row = self.sheetsDataTable.NewRow()
            row[0] = False
            row[1] = sheet
            row[2] = status[0]
            row[3] = status[1]

            self.sheetsDataTable.Rows.Add(row)

        # DATAGRIDVIEW
        self.sheetsDataGrid = CommonDataGridView()
        self.sheetsDataGrid.DataSource = self.sheetsDataTable
        self.sheetsDataGrid.Location = Point(30, 45)
        self.sheetsDataGrid.AutoGenerateColumns = True

        # --- GROUP: END ---


        # --- GROUP: CREATE/OVERRIDE VIEWS ---
        self.createViewsGroup = CommonGroup()
        self.createViewsGroup.Text = "Select Views to Create/Override"
        self.createViewsGroup.Location = Point(30, 350)
        self.createViewsGroup.Font = groupLabelFont

        # --- DATAGRIDVIEW FOR VIEWS ---
        self.viewsDataTable = DataTable("Views")
        self.viewsDataTable.Columns.Add(DataColumn("Select", bool))
        self.viewsDataTable.Columns.Add(DataColumn("View Name", str))
        self.viewsDataTable.Columns.Add(DataColumn("View Type", str))
        self.viewsDataTable.Columns.Add(DataColumn("Status", str))

        # Populating rows
        sheetList = OrderedDict()
        sheetList["Permitting-Backyard"] = ["Floor Plan", "TO OVERRIDE"]
        sheetList["Permitting-Basement"] = ["Floor Plan", "NEW"]
        sheetList["Permitting-First Level"] = ["Floor Plan", "TO OVERRIDE"]
        sheetList["Permitting-Ground Level"] = ["Floor Plan", "TO OVERRIDE"]
        sheetList["Permitting-Rooftop"] = ["Floor Plan", "NEW"]
        sheetList["Ground Floor Ceiling Structure"] = ["Reflected Ceiling Plan", "NEW"]
        sheetList["Section 1"] = ["Section", "NEW"]
        sheetList["Section 2"] = ["Section", "NEW"]
        sheetList["Section 3"] = ["Section", "TO OVERRIDE"]
        sheetList["Electrical - Basement Lighting"] = ["Floor Plan - Lighting", "NEW"]
        sheetList["Electrical - Basement Outlets"] = ["Floor Plan - Electrical", "TO OVERRIDE"]
        sheetList["Electrical - First Floor Lighting"] = ["Floor Plan - Lighting", "NEW"]

        for sheet, status in sheetList.items():
            row = self.viewsDataTable.NewRow()
            row[0] = False
            row[1] = sheet
            row[2] = status[0]
            row[3] = status[1]

            self.viewsDataTable.Rows.Add(row)

        self.viewsDataGrid = CommonDataGridView()
        self.viewsDataGrid.DataSource = self.viewsDataTable
        self.viewsDataGrid.Location = Point(30, 45)
        self.viewsDataGrid.AutoGenerateColumns = True
        # --- GROUP: END ---

        # --- SMART ANNOTATE VIEWS ---
        self.smartAnnotateCheckBox = CommonCheckBox()
        self.smartAnnotateCheckBox.Text = "Smart Annotate Views"
        self.smartAnnotateCheckBox.Location = Point(30, 650)

        # --- CONTROLS ---
        # Add controls
        self.Controls.Add(self.titleText)

        self.Controls.Add(self.createSheetsGroup)
        self.createSheetsGroup.Controls.Add(self.sheetsDataGrid)

        self.Controls.Add(self.createViewsGroup)
        self.createViewsGroup.Controls.Add(self.viewsDataGrid)

        self.Controls.Add(self.smartAnnotateCheckBox)

        # --- BINDING EVENTS ---
        self.Activated += self.onCreateDocumentsFormResize
        self.ResizeEnd += self.onCreateDocumentsFormResize
        self.AcceptButton.MouseUp += self.onAcceptButtonClick
    
    # --- EVENTS ---
    def onCreateDocumentsFormResize(self, sender, args):
        # Getting actual size of view
        clientSize = self.ClientSize

        self.createSheetsGroup.Size = Size(clientSize.Width - 60, 50)
        self.sheetsDataGrid.Size = Size(clientSize.Width - 120, 200)

        self.createViewsGroup.Size = Size(clientSize.Width - 60, 50)
        self.viewsDataGrid.Size = Size(clientSize.Width - 120, 200)

    def onAcceptButtonClick(self, sender, args):
        self.Hide()
        self.Close()

class ManageProjectDataForm(BaseForm):

    def __init__(self):
        BaseForm.__init__(self)

        # Define the caption text
        self.Text = "Manage Project Data"

        # --- TITLE ---
        self.titleText.Text = "MANAGE PROJECT DATA"

        # Resizing form
        self.ClientSize = Size(self.ClientSize.Width + 200, self.ClientSize.Height + 25)

        # Initializing categories lists to iterate through
        mainCategories = ['Permit',
                          'Construction',
                          'Shop and Cut Sheets']

        secondaryCategories = ['Architecture',
                               'Kitchen',
                               'Bathroom',
                               'Structure',
                               'MEP']

        tertiaryCategories = ['Floor Plans',
                              'Elevations',
                              'Sections',
                              'Callout Details']

        typesData = ['<View Templates>',
                     '<Annotation Types>']

        shopCutsheetsCategories = ['All wall elevations',
                                   'Casework',
                                   'Fireplace',
                                   'Stair + Rails',
                                   'Hardscaping (pool, jacuzzi, etc)']

        # --- GROUP: CREATE/OVERRIDE SHEET CATEGORIES ---
        self.createSheetCategoriesGroup = CommonGroup()
        self.createSheetCategoriesGroup.Text = "Sheet Categories"
        self.createSheetCategoriesGroup.Location = Point(30, 75)
        self.createSheetCategoriesGroup.Font = groupLabelFont

        # --- TREEVIEW FOR SHEET CATEGORIES ---
        self.sheetCategoriesTreeView = CommonTreeView()
        self.sheetCategoriesTreeView.Location = Point(30, 45)    
        
        for category in mainCategories:

            mainNode = TreeNode(category)
            self.sheetCategoriesTreeView.Nodes.Add(mainNode)

            if category == "Shop and Cut Sheets":
                for subCategory in shopCutsheetsCategories:

                    subNode1 = TreeNode(subCategory)
                    mainNode.Nodes.Add(subNode1)

            else:
                for subCategory in secondaryCategories:

                    subNode1 = TreeNode(subCategory)
                    mainNode.Nodes.Add(subNode1)

                    for subNode12 in tertiaryCategories:

                        subNode2 = TreeNode(subNode12)
                        subNode1.Nodes.Add(subNode2)

                        for subNode22 in typesData:

                            subNode3 = TreeNode(subNode22)
                            subNode2.Nodes.Add(subNode3)

        # --- GROUP: END ---



        # --- GROUP: CREATE/OVERRIDE VIEW CATEGORIES ---
        self.createViewCategoriesGroup = CommonGroup()
        self.createViewCategoriesGroup.Text = "View Categories"
        self.createViewCategoriesGroup.Location = Point(30, (self.createSheetCategoriesGroup.Location.Y + self.sheetCategoriesTreeView.Size.Height + 150))
        self.createViewCategoriesGroup.Font = groupLabelFont

        # --- TREEVIEW FOR VIEW CATEGORIES ---
        self.viewCategoriesTreeView = CommonTreeView()
        self.viewCategoriesTreeView.Location = Point(30, 45)

        for category in mainCategories:

            mainNode = TreeNode(category)
            self.viewCategoriesTreeView.Nodes.Add(mainNode)

            if category == "Shop and Cut Sheets":
                for subCategory in shopCutsheetsCategories:

                    subNode1 = TreeNode(subCategory)
                    mainNode.Nodes.Add(subNode1)

            else:
                for subCategory in secondaryCategories:

                    subNode1 = TreeNode(subCategory)
                    mainNode.Nodes.Add(subNode1)

                    for subNode12 in tertiaryCategories:

                        subNode2 = TreeNode(subNode12)
                        subNode1.Nodes.Add(subNode2)

                        for subNode22 in typesData:

                            subNode3 = TreeNode(subNode22)
                            subNode2.Nodes.Add(subNode3)

        # --- GROUP: END ---



        # --- CONTROLS ---
        # Add controls
        self.Controls.Add(self.titleText)

        self.Controls.Add(self.createSheetCategoriesGroup)
        self.createSheetCategoriesGroup.Controls.Add(self.sheetCategoriesTreeView)

        self.Controls.Add(self.createViewCategoriesGroup)
        self.createViewCategoriesGroup.Controls.Add(self.viewCategoriesTreeView)

        # --- BINDING EVENTS ---
        self.Activated += self.onManageProjectDataFormResize
        self.ResizeEnd += self.onManageProjectDataFormResize
        self.AcceptButton.MouseUp += self.onAcceptButtonClick

    # --- EVENTS ---
    def onManageProjectDataFormResize(self, sender, args):
        # Getting actual size of view
        clientSize = self.ClientSize

        self.createSheetCategoriesGroup.Size = Size(clientSize.Width - 60, 50)
        self.sheetCategoriesTreeView.Size = Size(clientSize.Width - 120, 150)

        self.createViewCategoriesGroup.Size = Size(clientSize.Width - 60, 50)
        self.viewCategoriesTreeView.Size = Size(clientSize.Width - 120, 150)

    def onAcceptButtonClick(self, sender, args):
        self.Hide()
        manageProjectDataForm = ManageProjectDataForm_EditOptions()
        manageProjectDataForm.ShowDialog()
        self.Close()

class ManageProjectDataForm_EditOptions(BaseForm):
    
    def __init__(self):
        BaseForm.__init__(self)

        # Define the caption text
        self.Text = "Manage Project Data"

        # --- TITLE ---
        self.titleText.Text = "View Templates"

        # Resizing Form
        self.ClientSize = Size(self.ClientSize.Width, 400)

        # --- GROUP: PROJECT OPTIONS
        self.viewTemplatesListGroup = CommonGroup()
        self.viewTemplatesListGroup.Text = "Permitting Floor Plans' View Templates"
        self.viewTemplatesListGroup.Location = Point(30, 110)
        self.viewTemplatesListGroup.Font = groupLabelFont

        # --- CHECKEDLISTBOX ---
        self.viewTemplatesChListBox = CommonCheckedListBox()
        self.viewTemplatesChListBox.Location = Point(30, 45)
        self.viewTemplatesChListBox.Size = Size(480, 110)
        self.viewTemplatesChListBox.Items.Add("<Permitting - Floor Plan>", False)
        self.viewTemplatesChListBox.Items.Add("<Permitting - Enlarged Floor Plan>", False)
        self.viewTemplatesChListBox.Items.Add("<Permitting - Cutsheets>", False)
        # --- GROUP: END ---

        # --- CONTROLS ---
        # Add controls
        self.Controls.Add(self.titleText)

        self.Controls.Add(self.viewTemplatesListGroup)
        self.viewTemplatesListGroup.Controls.Add(self.viewTemplatesChListBox)

        # --- BINDING EVENTS ---
        self.ResizeEnd += self.onManageProjectDataForm_EditOption_Resize
        self.AcceptButton.MouseUp += self.onAcceptButtonClick
    
    # --- EVENTS ---
    def onManageProjectDataForm_EditOption_Resize(self, sender, args):
        # Getting actual size of view
        clientSize = self.ClientSize

        self.viewTemplatesListGroup.Size = Size(clientSize.Width - 60, 50)
        self.viewTemplatesChListBox.Size = Size(clientSize.Width - 120, 110)

        self.viewTemplatesListGroup.Size = Size(clientSize.Width - 60, 50)
        self.viewTemplatesChListBox.Size = Size(clientSize.Width - 120, 110)

    def onAcceptButtonClick(self, sender, args):
        self.Hide()
        self.Close()

class DetailCreatorForm(BaseForm):

    def __init__(self):
        BaseForm.__init__(self)

        # --- TITLE ---
        self.titleText.Text = "DETAIL CREATOR"

        # Resizing form
        self.ClientSize = Size(self.ClientSize.Width + 200, self.ClientSize.Height + 200)

        # --- GROUP: ELEMENT TYPE PULLDOWN ---
        self.categoryDropdownGroup = CommonGroup()
        self.categoryDropdownGroup.Text = "Choose Category of Element"
        self.categoryDropdownGroup.Location = Point(30, 75)

        # Data for the combobox
        categorySelectionData = ["Select A Category", "Casework", "Furniture", "Walls", "Floors", "Roofs", "Ceilings"]

        self.categoryComboBox = CommonComboBox(categorySelectionData)
        self.categoryComboBox.Location = Point(30, 45)
        self.categoryComboBox.Size = Size(480, 50)

        # Default selection
        self.categoryComboBox.SelectedIndex = 0
        # --- GROUP: END ---

        # --- GROUP: OBJECT SELECTION ---
        self.objectSelectionGroup = CommonGroup()
        self.objectSelectionGroup.Text = "Object Selection"
        self.objectSelectionGroup.Location = Point(30, self.categoryDropdownGroup.Location.Y + self.categoryDropdownGroup.Height + 60)

        # PROMPT OBJECT SELECTION
        self.selectionButton = CommonButton()
        self.selectionButton.Location = Point(30, 45)
        self.selectionButton.Text = "Select Objects"

        selectedObjects = [
            "Bedroom1_Wardrobe",
            "Kitchen_Island",
            "Kitchen_Core",
            "Bathroom1_Vanity"
        ]

        self.label = Label()
        self.label.Font = generalFont
        self.label.Size = Size(250, self.label.Font.Height)
        self.label.Location = Point(30, 120)
        for text in selectedObjects:

            self.label.Text = self.label.Text + text + "\n"
            self.label.Size = Size(250, self.label.Size.Height + self.label.Font.Height)
        # --- GROUP: END ---

        # --- GROUP: DETAIL TYPE SELECTION AND ANNOTATION PREFERENCES ---
        self.detailTypePreferencesGroup = CommonGroup()
        self.detailTypePreferencesGroup.Text = "Detail Types | Annotation Preferences"
        self.detailTypePreferencesGroup.Location = Point(30, self.objectSelectionGroup.Location.Y + self.objectSelectionGroup.Height + 190)

        # CHECKEDLISTBOX
        self.detailOptionsChListBox = CommonCheckedListBox()
        self.detailOptionsChListBox.Location = Point(30, 45)
        self.detailOptionsChListBox.Size = Size(480, 140)
        self.detailOptionsChListBox.Items.Add("Long Section(s)", False)
        self.detailOptionsChListBox.Items.Add("Short Section(s)", False)
        self.detailOptionsChListBox.Items.Add("Plan Detail", False)
        self.detailOptionsChListBox.Items.Add("Elevation(s)", False)
        self.detailOptionsChListBox.Items.Add("Sub detail(s)", False)

        # COMBOBOX
        # Data for the combobox
        annotationPreferencesData = ["Smart Annotate", "Import Existing Detail From Library"]

        self.annotationPreferencesCombobox = CommonComboBox(annotationPreferencesData)
        self.annotationPreferencesCombobox.Location = Point(30, 45)
        self.annotationPreferencesCombobox.Size = Size(480, 50)

        # Default selection
        self.categoryComboBox.SelectedIndex = 0

        # --- GROUP: END ---

        # --- CONTROLS ---
        # Add controls
        self.Controls.Add(self.titleText)

        self.Controls.Add(self.categoryDropdownGroup)
        self.categoryDropdownGroup.Controls.Add(self.categoryComboBox)

        self.Controls.Add(self.objectSelectionGroup)
        self.objectSelectionGroup.Controls.Add(self.selectionButton)
        

        self.Controls.Add(self.detailTypePreferencesGroup)
        self.detailTypePreferencesGroup.Controls.Add(self.detailOptionsChListBox)
        self.detailTypePreferencesGroup.Controls.Add(self.annotationPreferencesCombobox)

        # --- BINDING EVENTS ---
        self.AcceptButton.MouseUp += self.onAcceptButtonClick
        self.selectionButton.MouseUp += self.onSelectionButtonClick
        self.Activated += self.onDetailCreatorFormResize
        self.ResizeEnd += self.onDetailCreatorFormResize

    # --- EVENTS ---
    def onDetailCreatorFormResize(self, sender, args):
        # Getting actual size of view
        clientSize = self.ClientSize

        self.categoryDropdownGroup.Size = Size(clientSize.Width - 60, 50)
        self.categoryComboBox.Size = Size(clientSize.Width - 120, 110)

        self.objectSelectionGroup.Size = Size(clientSize.Width - 60, 50)
        self.selectionButton.Size = Size(clientSize.Width - 120, 64)

        self.detailTypePreferencesGroup.Size = Size(clientSize.Width - 60, 50)
        self.detailOptionsChListBox.Size = Size(clientSize.Width - 120, 110)
        self.annotationPreferencesCombobox.Size = Size(clientSize.Width - 120, 64)

    def onAcceptButtonClick(self, sender, args):
        self.Hide()
        detailCreatorForm_ItemsPreferences.Activate()
        detailCreatorForm_ItemsPreferences.ShowDialog()
        self.Close()

    def onSelectionButtonClick(self, sender, args):
        self.Hide()
        self.objectSelectionGroup.Controls.Add(self.label)
        sleep(5)
        self.Show()

class DetailCreatorForm_ItemsPreferences(BaseForm):
    def __init__(self):
        BaseForm.__init__(self)

        # --- TITLE ---
        self.titleText.Text = "DETAIL CREATOR - PREFERENCES PER ITEM"

        # Resizing form
        self.ClientSize = Size(self.ClientSize.Width + 200, self.ClientSize.Height - 150)

        # --- GROUP: ELEMENT PREFERENCES PULLDOWN ---
        self.elementPreferencesGroup = CommonGroup()
        self.elementPreferencesGroup.Text = "Annotation and Placement Preferences"
        self.elementPreferencesGroup.Location = Point(30, 75)

        # DATATABLE
        elementsDataTable = DataTable("Annotation Preferences")
        elementsDataTable.Columns.Add(DataColumn("Element", str))
        elementsDataTable.Columns.Add(DataColumn("Annotation Preferences", str))
        elementsDataTable.Columns.Add(DataColumn("Place On Sheet", bool))

        # DATAGRIDVIEW
        self.elementsDataGrid = CommonDataGridView()
        self.elementsDataGrid.DataSource = elementsDataTable
        self.elementsDataGrid.Location = Point(30, 45)
        self.elementsDataGrid.AutoGenerateColumns = False
        
        # Defining columns
        elementColumn = DataGridViewTextBoxColumn()
        elementColumn.HeaderText = "Element"
        elementColumn.DataPropertyName = "Element"

        self.annotationPreferencesDropdownOptions = ["Smart Annotate", "<Select Details from Library>", "*Detail Selected"]

        annotationPreferencesColumn = DataGridViewComboBoxColumn()
        annotationPreferencesColumn.HeaderText = "Annotation Preferences"
        annotationPreferencesColumn.DataSource = self.annotationPreferencesDropdownOptions
        annotationPreferencesColumn.DataPropertyName = "Annotation Preferences"
        annotationPreferencesColumn.FlatStyle = FlatStyle.Flat

        placeOnSheetColumn = DataGridViewCheckBoxColumn()
        placeOnSheetColumn.HeaderText = "Place on Sheet"
        placeOnSheetColumn.DataPropertyName = "Place On Sheet"

        self.elementsDataGrid.Columns.Add(elementColumn)
        self.elementsDataGrid.Columns.Add(annotationPreferencesColumn)
        self.elementsDataGrid.Columns.Add(placeOnSheetColumn)

        # Populating rows
        elementList = [
            "Bedroom1_Wardrobe",
            "Kitchen_Island",
            "Kitchen_Core",
            "Bathroom1_Vanity"
        ]

        for element in elementList:
            row = elementsDataTable.NewRow()
            row[0] = element
            row[1] = self.annotationPreferencesDropdownOptions[0]
            row[2] = True

            elementsDataTable.Rows.Add(row)

        # --- GROUP: END ---

        # --- CONTROLS ---
        # Add controls
        self.Controls.Add(self.titleText)

        self.Controls.Add(self.elementPreferencesGroup)
        self.elementPreferencesGroup.Controls.Add(self.elementsDataGrid)

        # --- BINDING EVENTS ---
        self.AcceptButton.MouseUp += self.onAcceptButtonClick
        self.Activated += self.onDetailCreatorFormResize
        self.ResizeEnd += self.onDetailCreatorFormResize

        self.elementsDataGrid.CurrentCellDirtyStateChanged  += self.onDataGridView_CurrentCellDirtyStateChanged
        self.elementsDataGrid.CellValueChanged  += self.onDataGridView_CellValueChanged

    # --- EVENTS ---
    def onAcceptButtonClick(self, sender, args):
        self.Hide()
        self.Close()

    def onDetailCreatorFormResize(self, sender, args):
        # Getting actual size of view
        clientSize = self.ClientSize

        self.elementPreferencesGroup.Size = Size(clientSize.Width - 60, 50)
        self.elementsDataGrid.Size = Size(clientSize.Width - 120, 200)

    def onDataGridView_CurrentCellDirtyStateChanged(self, sender, args):
        
        self.elementsDataGrid.CommitEdit(DataGridViewDataErrorContexts.Commit)

    def onDataGridView_CellValueChanged(self, sender, args):
        
        if self.elementsDataGrid.CurrentCell.Value == "<Select Details from Library>":
            detailLibraryForm = DetailCreatorForm_ItemsPreferences_DetailLibrary()
            detailLibraryForm.Activate()
            detailLibraryForm.ShowDialog()

class DetailCreatorForm_ItemsPreferences_DetailLibrary(BaseForm):

    def __init__(self):
        BaseForm.__init__(self)

        # --- TITLE ---
        self.titleText.Text = "DETAIL LIBRARY"

        # Resizing form
        self.ClientSize = Size(624, self.ClientSize.Height + 300)

        # FILTERSEARCHBOX
        self.searchTextBox = CommonTextBox()
        self.searchTextBox.Size = Size(200, self.searchTextBox.Size.Width)
        self.searchTextBox.Location = Point(self.ClientSize.Width - 30 - self.searchTextBox.Size.Width, 20)

        # LISTVIEW
        self.listView = CommonListView()
        self.listView.Location = Point(0, self.titleIcon.Height)

        # ImageList
        imageList = ImageList()
        imageList.ImageSize = Size(256, 256)
        # Get list of images in resources folder
        imageFilesPaths = System.IO.Directory.GetFiles(detailLibraryResourcesPath)

        try:
            for imagePath in imageFilesPaths:
                imageList.Images.Add(Image.FromFile(imagePath))
        except:
            print("Images could not be loaded")
            pass

        
        self.listView.LargeImageList = imageList

        self.listView.Items.Add(ListViewItem("Floor Plan", 0))
        self.listView.Items.Add(ListViewItem("Front Elevation", 1))
        self.listView.Items.Add(ListViewItem("Long Section", 2))
        self.listView.Items.Add(ListViewItem("Short Section", 3))
        self.listView.Items.Add(ListViewItem("Short Section", 4))
        self.listView.Items.Add(ListViewItem("Short Section", 5))
        self.listView.Items.Add(ListViewItem("Short Section", 6))

        # --- CONTROLS ---
        # Add controls
        self.Controls.Add(self.titleText)
        self.Controls.Add(self.searchTextBox)
        self.Controls.Add(self.listView)

        # --- BINDING EVENTS ---
        self.AcceptButton.MouseUp += self.onAcceptButtonClick
        self.Activated += self.onDetailLibraryFormResize
        self.ResizeEnd += self.onDetailLibraryFormResize

    # --- EVENTS ---
    def onDetailLibraryFormResize(self, sender, args):
        # Getting actual size of view
        clientSize = self.ClientSize

        self.listView.Size = Size(clientSize.Width, clientSize.Height - self.titleIcon.Height - self.AcceptButton.Size.Height)
        self.searchTextBox.Location = Point(self.ClientSize.Width - 30 - self.searchTextBox.Size.Width, self.searchTextBox.Size.Height)

    def onAcceptButtonClick(self, sender, args):
        self.Hide()
        currentCell = detailCreatorForm_ItemsPreferences.elementsDataGrid.CurrentCell
        currentCell.Value = "*Detail Selected"
        self.Close()

class SmartAutoDetailForm(BaseForm):
    def __init__(self):
        BaseForm.__init__(self)

        # --- TITLE ---
        self.titleText.Text = "SMART AUTO DETAIL"

        # Resizing form
        self.ClientSize = Size(self.ClientSize.Width + 400, self.ClientSize.Height - 150)

        # --- GROUP: DETAIL VIEWS ---
        self.smartAutoDetailGroup = CommonGroup()
        self.smartAutoDetailGroup.Text = "Detail Views"
        self.smartAutoDetailGroup.Location = Point(30, 75)
        self.smartAutoDetailGroup.Font = groupLabelFont

        # DATATABLE
        self.smartAutoDetailDataTable = DataTable("Sheets")
        self.smartAutoDetailDataTable.Columns.Add(DataColumn("Select", bool))
        self.smartAutoDetailDataTable.Columns.Add(DataColumn("Element", str))
        self.smartAutoDetailDataTable.Columns.Add(DataColumn("Type of Drawing", str))
        self.smartAutoDetailDataTable.Columns.Add(DataColumn("Status", str))
        self.smartAutoDetailDataTable.Columns.Add(DataColumn("Place On Sheet", bool))

        # Populating rows
        elementList = OrderedDict()
        elementList["Architectural - Exterior Wall Section 1"] = ["NEW"]
        elementList["Architectural - Exterior Wall Section 2"] = ["NEW"]
        elementList["Architectural - Exterior Wall Section 3"] = ["NEW"]
        elementList["Architectural - Exterior Wall Section 4"] = ["NEW"]
        elementList["Architectural - Exterior Wall Detail 1"] = ["NEW"]
        elementList["Architectural - Exterior Wall Detail 2"] = ["NEW"]
        elementList["Architectural - Exterior Wall Detail 3"] = ["NEW"]
        elementList["Architectural - Basement Wall Detail"] = ["NEW"]
        elementList["Architectural - Roof Detail"] = ["NEW"]
        elementList["Architectural - Stair Long Section"] = ["TO OVERRIDE"]
        elementList["Architectural - Stairs' Stringers Detail"] = ["TO OVERRIDE"]
        elementList["Architectural - Interior Glass Railings Detail"] = ["NEW"]
        elementList["Structural - Wall Foundation Detail"] = ["TO OVERRIDE"]

        for element, status in elementList.items():
            row = self.smartAutoDetailDataTable.NewRow()
            row[0] = False
            row[1] = element
            row[2] = "DETAIL"
            row[3] = status[0]
            row[4] = False

            self.smartAutoDetailDataTable.Rows.Add(row)

        # DATAGRIDVIEW
        self.smartAutoDetailDataGrid = CommonDataGridView()
        self.smartAutoDetailDataGrid.DataSource = self.smartAutoDetailDataTable
        self.smartAutoDetailDataGrid.Location = Point(30, 45)
        self.smartAutoDetailDataGrid.AutoGenerateColumns = True
        # --- GROUP: END ---

        # --- CONTROLS --- 
        # Add Controls
        self.Controls.Add(self.titleText)

        self.Controls.Add(self.smartAutoDetailGroup)
        self.smartAutoDetailGroup.Controls.Add(self.smartAutoDetailDataGrid)

        # --- BINDING EVENTS ---
        self.AcceptButton.MouseUp += self.onAcceptButtonClick
        self.Activated += self.onDetailCreatorFormResize
        self.ResizeEnd += self.onDetailCreatorFormResize

    # --- EVENTS ---
    def onAcceptButtonClick(self, sender, args):
        self.Hide()
        self.Close()

    def onDetailCreatorFormResize(self, sender, args):
        # Getting actual size of view
        clientSize = self.ClientSize

        self.smartAutoDetailGroup.Size = Size(clientSize.Width - 60, 50)
        self.smartAutoDetailDataGrid.Size = Size(clientSize.Width - 120, 200)

# --- Custom Control Classes ---
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

class BuildingOptionButton(Button):
    def __init__(self):
        self.FlatStyle = FlatStyle.Flat
        self.FlatAppearance.BorderSize = 1
        self.FlatAppearance.BorderColor = blackColor
        self.BackColor = grayColor
        self.ForeColor = blackColor
        self.Font = buttonFont

        # --- BINDING EVENTS ---
        self.MouseUp += self.onMouseUp
        self.MouseEnter += self.onMouseEnter
        self.MouseLeave += self.onMouseLeave

    # EVENTS
    def onMouseUp(self, sender, args):
        self.BackColor = blackColor
        self.ForeColor = whiteColor

        self.MouseLeave -= self.onMouseLeave
        self.MouseEnter -= self.onMouseEnter

    def onMouseEnter(self, sender, args):
        self.BackColor = Color.FromArgb(100, 100, 100)

    def onMouseLeave(self, sender, args):
        self.BackColor = grayColor

class CommonComboBox(ComboBox):

    def __init__(self, list):

        # Colors
        self.BackColor = pureWhiteColor
        self.ForeColor = blackColor

        # Heights
        self.DropDownHeight = 100
        self.ItemHeight = 20

        # Font
        self.Font = generalFont

        # Inserting data in the combobox
        for item in list:
            self.Items.Insert(list.index(item), item)

class CommonCheckBox(CheckBox):

    def __init__(self):
        
        self.Size = Size(300, 20)
        self.Font = h2Font

class CommonCheckedListBox(CheckedListBox):

    def __init__(self):

        # Style
        self.BackColor = pureWhiteColor
        self.ForeColor = blackColor
        self.FlatStyle = FlatStyle.Flat

        # Heights
        self.ItemHeight = 15

        # Size
        self.Size = Size(480, 80)

        # Selection Mode
        self.SelectionMode = SelectionMode.One

        # Font
        self.Font = generalFont

class CommonDataGridView(DataGridView):

    def __init__(self):
        self.BackColor = pureWhiteColor
        self.ForeColor = blackColor
        self.BackgroundColor = pureWhiteColor
        self.BorderStyle = BorderStyle.Fixed3D
        self.Size = Size(480, 200)

        # Cells Styles and Format
        # Declaring the styles
        headerCellStyle = DataGridViewCellStyle()
        headerCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        generalCellStyle = DataGridViewCellStyle()
        generalCellStyle.Font = generalFont

        # Assigning the styles
        self.DefaultCellStyle = generalCellStyle
        self.ColumnHeadersDefaultCellStyle = headerCellStyle    

        # Columns AutoSize
        self.AutoSizeColumnsMode = (DataGridViewAutoSizeColumnsMode.Fill)

        # Data properties
        self.AllowUserToDeleteRows = False
        self.AllowUserToAddRows = False

class CommonGroup(GroupBox):

    def __init__(self):
        
        self.AutoSize = True
        self.Size = Size(540, 50)
        self.Font = groupLabelFont

class CommonListView(ListView):
    def __init__(self):
        self.BackColor = blackColor
        self.ForeColor = whiteColor
        self.FlatStyle = FlatStyle.Flat
        self.Font = h2Font

        self.Size = Size(200, 300)

class CommonTreeView(TreeView):

    def __init__(self):
        
        # Style
        self.BackColor = pureWhiteColor
        self.ForeColor = blackColor

        # Heights
        self.ItemHeight = 20

        # Size
        self.Size = Size(480, 80)

        # Font
        self.Font = generalFont

class CommonTextBox(TextBox):
    def __init__(self):
        
        self.BackColor = pureWhiteColor
        self.ForeColor = blackColor
        self.Font = generalFont
        self.Size = Size(100, 20)
        self.FlatStyle = FlatStyle.Flat

        self.PlaceholderText = "Search"

# --- Forms Creation ---
# Define the Forms.Application object
winFormApp = Application

# Initialize all the forms
createProjectForm = CreateProjectForm()
createDocumentsForm = CreateDocumentsForm()
manageProjectDataForm = ManageProjectDataForm()
detailCreatorForm = DetailCreatorForm()
detailCreatorForm_ItemsPreferences = DetailCreatorForm_ItemsPreferences()
smartAutoDetailForm = SmartAutoDetailForm()