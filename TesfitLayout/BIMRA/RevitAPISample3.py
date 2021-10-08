import clr
import math
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

def coltypesofcategory(CAT,TLIST):
	collector = FilteredElementCollector(doc)
	collector.OfCategory(CAT)
	collector.OfClass(FamilySymbol)
	famtypeitr = collector.GetElementIdIterator()
	famtypeitr.Reset()
	PosibleTypes = []
	for item in famtypeitr:
		famtypeID = item
		famsymb = doc.GetElement(famtypeID)
		for a in range(len(TLIST)):
			if famsymb.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()==TLIST[a]:
				FAM = famsymb
				PosibleTypes.append(FAM)
	return PosibleTypes
	
def coltypenameofcategory(CAT,TYPENAME):
	collector = FilteredElementCollector(doc)
	collector.OfCategory(CAT)
	collector.OfClass(FamilySymbol)
	famtypeitr = collector.GetElementIdIterator()
	famtypeitr.Reset()
	PosibleTypes = []
	for item in famtypeitr:
		famtypeID = item
		famsymb = doc.GetElement(famtypeID)
		if famsymb.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()==TYPENAME:
			FAM = famsymb
	return FAM
	
def gettypesbyname(CLTypes,LNames):
	TP=[]
	for a in range(len(LNames)):
		for b in range(len(CLTypes)):
			if CLTypes[b].get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()==LNames[a]:
				FAM2 = CLTypes[b]
				TP.append(FAM2)
	return TP

def getposttype(LIST):
	FLEN=[]
	for a in range(len(LIST)):
		if LIST[a]>10.66:
			FLEN.append("PrimaryPost_DropHead_Universal_Rised_Extension")
		else:
			FLEN.append("PrimaryPost_DropHead_Universal_Rised")
	return FLEN
	
def shortline(l1):
	la=[]
	for a in l1:
		sl=a.GetEndPoint(1).Add(a.Direction.Normalize().Negate().Multiply(1))
		ll=a.GetEndPoint(0).Add(a.Direction.Normalize().Multiply(1))
		la.append(Line.CreateBound(sl,ll))
	return la

def extendedendpointlines(a):
	sl=a.GetEndPoint(0).Add(a.Direction.Normalize().Negate().Multiply(1))
	ll=a.GetEndPoint(0).Add(a.Direction.Normalize().Multiply(1))
	la=Line.CreateBound(sl,ll)
	sl1=a.GetEndPoint(1).Add(a.Direction.Normalize().Negate().Multiply(1))
	ll1=a.GetEndPoint(1).Add(a.Direction.Normalize().Multiply(1))
	lb=Line.CreateBound(sl1,ll1))
	return la,lb
	
doc = DocumentManager.Instance.CurrentDBDocument
uidoc=DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = uidoc.ActiveView

collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()

SELLINES=[]
for item in collector:
	ftype2 = item.Symbol
	ftype = item.Symbol.Family
	if ftype.Name=="Ledger":
		if item.LookupParameter("Used").AsInteger()==0:
			SELLINES.append(item)
			
GEOMETRY=[]
ELEMENT=[]
IDSS=[]
Gopt = Options()
for a in SELLINES:
	geo1=a.get_Geometry(Gopt)
	enum1 = geo1.GetEnumerator(); enum1.MoveNext()
	geo2 = enum1.Current.GetInstanceGeometry()
	SET1=[]
	for g in geo2:
		if g.GetType() != Autodesk.Revit.DB.Solid:
			SET1.append(g)		
	GEOMETRY.append(SET1[-1])		
	IDSS.append(a.Id) 
		
RLINES = GEOMETRY
SURFACE=IN[0][0]

POINS=[]
PYLINES=[]
PLINE=[]
for a in RLINES:
	POINS.append(a.GetEndPoint(0))
	POINS.append(a.GetEndPoint(1))
	setslines=extendedendpointlines(a)
	PYLINES.append(setslines[0])
	PYLINES.append(setslines[1])
	PLINE.append(a)
	PLINE.append(a)

NL=shortline(RLINES)

CLEARPOINTS=[]
CLEARLINES=[]
BOB=[]
BOBD=[]
for a,b,c in zip(POINS,PYLINES,PLINE):
	X=0
	for d in NL:
		if b.Intersect(d).ToString()=="Overlap" or b.Intersect(d).ToString()=="Equal": 
			X=1
	if X!=1:
		CLEARPOINTS.append(a)
		CLEARLINES.append(c)
	else:
		BOB.append(a) 
		if a.Translate(c.Direction,0.23).DoesIntersect(c):
			BOBD.append(c.Direction)
		else:
			BOBD.append(c.Direction.Reverse())
PPOI=[]
PPOI2=[]
P1LEN=[]
P2LEN=[]

for a,b in zip(CLEARPOINTS,CLEARLINES):
	PR=a.Project(SURFACE,Vector.ByCoordinates(0,0,-1))
	PAR=b.ParameterAtPoint(a)
	LLN=b.Project(SURFACE,Vector.ByCoordinates(0,0,-1))	
	if len(PR)>0:
		PPOI.append(a)
		P1LEN.append(a.DistanceTo(PR[0]))
	else:
		if len(LLN)>0:
			PL=LLN[0].SegmentLengthAtParameter(PAR)
			if PL>8: 
				LL1=8
			elif PL>7:
				LL1=7
			elif PL>6:
				LL1=6
			elif PL>5:
				LL1=5
			else:
				LL1=4
			PPOI2.append(b.PointAtSegmentLength(LL1))
			P2LEN.append(b.PointAtSegmentLength(LL1).DistanceTo(LLN[0].PointAtSegmentLength(LL1)))	
			
FPOI=[]
FLEN=[]
PINDEX=[]
for a,b in zip(PPOI,P1LEN):
	if str(str(a.X)+str(a.Y)) not in PINDEX:
		FPOI.append(a.ToXyz())
		FLEN.append(b)
		PINDEX.append(str(str(a.X)+str(a.Y)))

FPOI2=[]
FLEN2=[]
PINDEX2=[]
for a,b in zip(PPOI2,P2LEN):
	if str(str(a.X)+str(a.Y)) not in PINDEX2:
		FPOI2.append(a.Translate(Vector.ByCoordinates(0,0,-0.54)).ToXyz())
		FLEN2.append(round(b)-0.54)
		PINDEX2.append(str(str(a.X)+str(a.Y)))

IPO=[]
IPT=[]
for a in range(len(BOB)):
	IPO.append(BOB[a].Translate(Vector.ByCoordinates(0,0,-0.55)).Translate(BOBD[a],0.23).ToXyz())
	IPT.append(coltypenameofcategory(BuiltInCategory.OST_GenericModel,"PSH8142"))
	 
#NormalPosts
NLEN=getposttype(FLEN) 
NNLEN = list(set(NLEN))
PosibleTypes = coltypesofcategory(BuiltInCategory.OST_GenericModel,NNLEN)
TYPES1=gettypesbyname(PosibleTypes,NLEN)			
#ShorterPosts
NLEN2= getposttype(FLEN2)
NNLEN2 = list(set(NLEN2))
PosibleTypes2 = coltypesofcategory(BuiltInCategory.OST_GenericModel,NNLEN2)
TYPES2=gettypesbyname(PosibleTypes2,NLEN2)

TransactionManager.Instance.EnsureInTransaction(doc) 
for a in range(len(FPOI)):
	fam1=doc.Create.NewFamilyInstance(FPOI[a],TYPES1[a],Structure.StructuralType.NonStructural)
	fam1.LookupParameter("Span").Set(FLEN[a])
for a in range(len(FPOI2)):
	fam2=doc.Create.NewFamilyInstance(FPOI2[a],TYPES2[a],Structure.StructuralType.NonStructural)
	fam2.LookupParameter("Span").Set(FLEN2[a])
for a in range(len(IPO)):
	fam2=doc.Create.NewFamilyInstance(IPO[a],IPT[a],Structure.StructuralType.NonStructural)
TransactionManager.Instance.TransactionTaskDone() 


"""
# Assign your output to the OUT variable.
OUT = CLEARPOINTS,CLEARLINES,PPOI,PPOI2
"""
OUT=BOB,BOBD