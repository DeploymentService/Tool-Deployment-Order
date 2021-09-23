import sys,os
from os import path
dumps_dir = path.expandvars(r'C:\Users\e2133\Documents\GitHub\Tool-Deployment-Order\TesfitLayout\BIMRA')
#dumps_dir = path.expandvars(r'%ProgramW6432%\Dynamo\RA\BIMRA')
sys.path.append(dumps_dir)
from CClasses import *

import clr
clr.AddReference("RevitAPIUI")
from  Autodesk.Revit.UI import *
clr.AddReference("RevitAPI")
import Autodesk

from Autodesk.Revit.DB import *
clr.AddReference("RevitNodes")
import Revit
clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)

clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager


doc = DocumentManager.Instance.CurrentDBDocument
uidoc=DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = uidoc.ActiveView

def output1(l1):
	if len(l1) == 1: return l1[0] 
	else: return l1

"""sel1 = uidoc.Selection
obt1 = Selection.ObjectType.Element
msg1 = 'Select Lines on Model, hit ESC to stop picking.'
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
"""
SELLNES = [496364,496363,496365,496362]
out1=[]
for a in SELLNES:
	out1.append(doc.GetElement(ElementId(a)))

VALUES=[]
NAMLIN=["Flunkage Line","Side Line","Front Line","Rear Line"]
SETBACKS = [XYZ(4.1,0,0),XYZ(-4.1,0,0),XYZ(0,19.7,0),XYZ(0,-24.6,0)]
	
for a in NAMLIN:
	for b in output1(out1):
		if b.LineStyle.Name==a: 
			VALUES.append(b)
CLINES=[]
LINES=[]
for a,b in zip(VALUES,SETBACKS):
	LINES.append(Line.CreateBound(a.GeometryCurve.GetEndPoint(0).Add(b),a.GeometryCurve.GetEndPoint(1).Add(b)))
	
PTS1=get4IP(LINES)
LLS=PTStoLINES(PTS1)
	
TransactionManager.Instance.EnsureInTransaction(doc)
for a in LLS:
	crv=doc.Create.NewModelCurve(a,SketchPlane.Create(doc,PlaneFunction(a)))

TransactionManager.Instance.TransactionTaskDone()

XVAL=0
YVAL=0
LENGTH=0
LINES1=[]
for a in LLS:
	LENGTH+=a.Length
	LINES1.append(a.ToProtoType())
	XVAL+=a.GetEndPoint(0).X
	YVAL+=a.GetEndPoint(0).Y
	
CX=XVAL/len(LLS)  
CY=YVAL/len(LLS)
CP=XYZ(CX,CY,0)
VECTORS = [XYZ(0,-1,0),XYZ(1,0,0),XYZ(0,1,0),XYZ(-1,0,0)]

INTP=[]
DVECTORS=[]
for a,b in zip(LLS,VECTORS):
	DVECTORS.append(b.ToVector())
	LN1=Line.CreateBound(CP,CP.Add(b.Multiply(LENGTH)))
	INTP.append(get_intersection(LN1,a))
	
SL=Line.CreateBound(INTP[3],INTP[1])
LL=Line.CreateBound(INTP[0],INTP[2])

XVEC=VECTORS[1]
YVEC=VECTORS[2]

XGRIDS=CMODGRID3(LL,SL,YVEC) 
YGRIDS=CMODGRID3(SL,LL,XVEC)   

FORMACRO=CreateCMODGRID(XGRIDS,YGRIDS)

#WRITE TO C#
SET = len(FORMACRO)/2  
PROG = ARCHPROG(SET)  
 
MACROC=[]   
MODC=[]       
SPOTS=[]     
IOFS=[]   
 
PRIGROU=[] 
for a in FORMACRO:
	if a.SC5==0 and a.SC1==0:  
		PRIGROU.append(a.IOF)
MACRONS=CMACRONSFROMPROG(PROG)   
CMODOLS=CMODOLSFROMMODULES(FORMACRO,MACRONS) 

MACROC=MACRONS 
MODC=CMODOLS
SELMOD=[] 
SELPRO=[]  
SP1=[]
IO1=[]
NUMS=[]

for a in CMODOLS: 
	NUMS.append(a.IOF)

for a in PROG: 
	X=0
	IOSS=[]
	for b in MODC:
		if a in b.NAML:
			X+=1 
			IOSS.append(b.IOF) 
	IO1.append(IOSS)
	SP1.append(X) 
SPOTS=SP1
IOFS=IO1
 
for a in range(len(PROG)):
	ELEMENT=WCFSELECT(PROG,SPOTS,IOFS,SELMOD,SELPRO,PRIGROU,CMODOLS,NUMS)
	MACROC=WCFCOLLAPSE(PROG,MACROC,ELEMENT)	 
	PROPAGATE=WCFPROPAGATE(SELMOD,SELPRO,MODC,CMODOLS,MACRONS,ELEMENT,NUMS) 
	MODC=PROPAGATE[0]	
	PRIGROU=PROPAGATE[1]
	EVALS=WCFEVALUATE(PROG,MODC) 
	SPOTS=EVALS[0]
	IOFS=EVALS[1]

LOCS=[]
NAM=[]
MOD=[]
for a,b in zip(SELMOD,SELPRO):
	for c in FORMACRO:
		if c.IOF==a:
			LOCS.append(c.ORI)
			NAM.append(b[3:])
			MOD.append(c)

#symbName = "MACROS"
collector = FilteredElementCollector(doc)
collector.OfCategory(BuiltInCategory.OST_SpecialityEquipment)
collector.OfClass(FamilySymbol)

famtypeitr = collector.GetElementIdIterator()
famtypeitr.Reset()

TYPES = []
for item in famtypeitr:
	famtypeID = item
	famsymb = doc.GetElement(famtypeID)
	for a in NAM:
		if famsymb.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()==a:
			FAM = famsymb
			TYPES.append(FAM)
#	NAMES.append(famsymb.Family.Name)
#	TYPES.append(famsymb.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString())

TransactionManager.Instance.EnsureInTransaction(doc) 
for a in range(len(LOCS)):
	fam=doc.Create.NewFamilyInstance(LOCS[a],TYPES[a],Structure.StructuralType.NonStructural)
	fam.LookupParameter("WWidth").Set(MOD[a].WID)
	fam.LookupParameter("LLength").Set(MOD[a].LEN)
	fam.LookupParameter("HHeigth").Set(MOD[a].HEI)

TransactionManager.Instance.TransactionTaskDone() 

OUT = LOCS,NAM