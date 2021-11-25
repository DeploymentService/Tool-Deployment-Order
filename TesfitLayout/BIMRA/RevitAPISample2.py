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

def shortline(l1):
	la=[]
	for a in l1:
		sl=a.GetEndPoint(1).Add(a.Direction.Normalize().Negate().Multiply(1))
		ll=a.GetEndPoint(0).Add(a.Direction.Normalize().Multiply(1))
		la.append(Line.CreateBound(sl,ll))
	return la
	
def listtostring(s):
	str1=""
	for ele in s:
		str1+=str(ele)
	return str1
	
def getlnsids(lines,rrids):
	shlines=shortline(lines)
	intr=[]
	for a in range(len(lines)):
		set1=[]
		for b in range(len(shlines)):
			if shlines[a].Intersect(shlines[b]).ToString()=="Overlap" or shlines[a].Intersect(shlines[b]).ToString()=="Equal" or mpline(shlines[a]).DistanceTo(mpline(shlines[b]))<0.5 :
				set1.append(b)
		set1.sort()
		intr.append(set1)
		
	a3=[]	
	for a in intr:
		a3.append(listtostring(a))
	a4=list(set(a3))	
	a5=[]
	for a in a4:
		a5.append(intr[a3.index(a)])	
	lns=[]
	ids=[]
	for a in a5:
		if len(a)>1:
			lens=[]
			for b in a:
				lens.append(lines[b].Length)
			am=max(lens)
			lns.append(lines[a[lens.index(am)]])
			ids.append(rrids[a[lens.index(am)]])
		else:
			lns.append(lines[a[0]])
			ids.append(rrids[a[0]])
	return lns,ids

def mpline(sl):
	tl=sl.GetEndPoint(1).Add(sl.Direction.Negate().Normalize().Multiply(sl.Length/2))
	return tl

def verline(pt):
	vl=Line.CreateBound(pt,pt.Add(XYZ(0,0,1)))
	return vl

doc = DocumentManager.Instance.CurrentDBDocument
uidoc=DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = uidoc.ActiveView

collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Casework).WhereElementIsNotElementType().ToElements()

SELLINES=[]
for item in collector:
	ftype = item.Symbol.Family
	if ftype.Name=="Proshore":
		if item.LookupParameter("Used").AsInteger()==0:
			SELLINES.append(item)
Gopt = Options()
GEOMETRY=[]
ELEMENT=[]
IDSS=[]
for a in SELLINES:
	geo1=a.get_Geometry(Gopt)
	enum1 = geo1.GetEnumerator()
	enum1.MoveNext()
	geo2 = enum1.Current.GetInstanceGeometry()
	SET1=[]
	for g in geo2:
		if g.GetType() != Autodesk.Revit.DB.Solid:
			SET1.append(g)
	ELEMENT.append(a)		
	GEOMETRY.append(SET1)		
	IDSS.append(a.Id) 
#	NAMES.append(famsymb.Family.Name)
#	TYPES.append(famsymb.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString())

l1=[]
i1=[]

for a,b in zip(GEOMETRY,IDSS):
	l1.append(a[2])
	l1.append(a[3])
	i1.append(b)
	i1.append(b)
l2=[]
i2=[]
	
for a,b in zip(GEOMETRY,IDSS):
	for c in range(len(a)):
		if c>3:
			l2.append(a[c])
			i2.append(b)


LNSIDS1 = getlnsids(l1,i1)
LNSIDS2 = getlnsids(l2,i2)
#Arc1=Arc.Create(XYZ.Zero,1,0,1.0* 3.1416,XYZ.BasisX,XYZ.BasisY).ToProtoType()
#L2=Line.CreateBound(XYZ(1,1,0),XYZ(1,1,1))
#L1=RevolvedSurface.Create(XYZ(0,0,0),XYZ(0,0,1),L2)
#L3=RevolvedSurface.Create(XYZ(0.5,0.5,0),XYZ(0.5,0.5,1),L2)

collector2 = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).OfClass(FamilySymbol)
famtypeitr = collector2.GetElementIdIterator()
famtypeitr.Reset()

FAMLED=[]
FAMPRO=[]
for item in famtypeitr:
	famtypeID = item
	famsymb = doc.GetElement(famtypeID)
	if famsymb.Family.Name=="Ledger":
		FAMLED.append(famsymb)
	elif famsymb.Family.Name=="PROSHORE_LVL":
		FAMPRO.append(famsymb)

FAML=[]
FAMP=[]


for a in LNSIDS1[0]:
	for b in FAMLED:
		if b.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()==round(a.Length,0).ToString()+"' LEDGER":
			FAML.append(b)

for a in LNSIDS2[0]:
	for b in FAMPRO:
		if b.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()=="LVL "+round(a.Length,0).ToString()+"'":
			FAMP.append(b)



ANG1=[]
for a in LNSIDS1[0]:
	ANG1.append(XYZ(1,0,0).AngleTo(a.Direction))

ANG2=[]
for a in LNSIDS2[0]:
	ANG2.append(XYZ(1,0,0).AngleTo(a.Direction))

TransactionManager.Instance.EnsureInTransaction(doc) 
for a in range(len(LNSIDS1[0])):
	fam=doc.Create.NewFamilyInstance(mpline(LNSIDS1[0][a]),FAML[a],Structure.StructuralType.NonStructural)
	fam.LookupParameter("Reference").Set(LNSIDS1[1][a].ToString())
	ElementTransformUtils.RotateElement(doc,fam.Id,verline(mpline(LNSIDS1[0][a])),ANG1[a])
	
for a in range(len(LNSIDS2[0])):
	fam=doc.Create.NewFamilyInstance(mpline(LNSIDS2[0][a]),FAMP[a],Structure.StructuralType.NonStructural)
	fam.LookupParameter("Reference").Set(LNSIDS2[1][a].ToString())
	ElementTransformUtils.RotateElement(doc,fam.Id,verline(mpline(LNSIDS2[0][a])),ANG2[a])	
	
	
TransactionManager.Instance.TransactionTaskDone() 	
	

OUT=len(LNSIDS1[0])