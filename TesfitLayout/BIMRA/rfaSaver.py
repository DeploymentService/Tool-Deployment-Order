import clr 
clr.AddReference("RevitServices") 
import RevitServices 
from RevitServices.Persistence import DocumentManager 
from RevitServices.Transactions import TransactionManager 
clr.AddReference("RevitAPI") 
import Autodesk 
clr.AddReference("RevitAPIUI")
from  Autodesk.Revit.UI import *
from Autodesk.Revit.DB import * 
clr.AddReference("RevitNodes")
import Revit
import traceback

doc = DocumentManager.Instance.CurrentDBDocument
uidoc=DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = uidoc.ActiveView

sel1 = uidoc.Selection
obt1 = Selection.ObjectType.Element
msg1 = 'Select Families, hit ESC to stop picking.'
out1 = []

flag = True
TaskDialog.Show("TestFit", msg1)
#Color Elements
TransactionManager.Instance.EnsureInTransaction(doc)
while flag:
	try:
		el1 = doc.GetElement(sel1.PickObject(obt1, msg1).ElementId)
		out1.append(el1)
	except : 
		flag = False
TransactionManager.Instance.TransactionTaskDone()
fami=[]
for el in out1:
	famtype = doc.GetElement(el.GetTypeId())
	fami.append(famtype.Family.Id)
fam2 = list(set(fami))
families=[] 
for a in fam2:
	families.append(doc.GetElement(a))
	
directory = "E:\UPDATER"
filePaths = [directory+"\\"+family.Name+".rfa" for family in families] 

overwrite = True
TransactionManager.Instance.ForceCloseTransaction() 


saveAsOptions = SaveAsOptions() 
saveAsOptions.OverwriteExistingFile = overwrite 
resultSet = [] 

if overwrite:
	successMessage = " Saved out the family documnet.\n Overwrote any rfa with matching name.\n You can recover from the rfa backup \n in the same directory if such exists." 
else: 
	successMessage = " Saved the family document.\n No previously existing rfa at path." 
i=0
famDocs = [doc.EditFamily(family) for family in families] 
for famDoc in famDocs: 
	try: 
		famDoc.SaveAs(filePaths[i], saveAsOptions) 
		resultSet.append(families[i].Name+".rfa:\n"+successMessage) 
	except: 
		resultSet.append(families[i].Name+".rfa:\n File was not saved due to\n the following exception:\n "+traceback.format_exc().split("Exception: ")[-1].strip()) 
		i+=1

OUT = resultSet,filePaths