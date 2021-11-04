import clr
import os
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *
 
clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
doc = DocumentManager.Instance.CurrentDBDocument

class FamilyOption(IFamilyLoadOptions) :
	def OnFamilyFound(self, familyInUse, overwriteParameterValues):
		overwriteParameterValues.Value = True
		return True

	def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
		overwriteParameterValues.Value = True
		return True

directoryPath = "E:/UPDATER/"
files = os.listdir(directoryPath)
pathsFamily = []
for a in range(len(files)):
	if "rfa" in files[a]:
		if ".00" not in files[a]:
			pathsFamily.append(directoryPath+files[a])

TransactionManager.Instance.EnsureInTransaction(doc)
loadedFam = [] 
opts = FamilyOption()
for path in pathsFamily :
	refFam = clr.Reference[Family]()
	doc.LoadFamily(path, opts, refFam)
	loadedFam.append(refFam.Value)

TransactionManager.Instance.TransactionTaskDone()

OUT = refFam

