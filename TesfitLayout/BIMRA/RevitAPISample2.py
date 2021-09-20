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

clr.AddReference("ProtoGeometry")
from Autodesk.DesignScript.Geometry import *


doc = DocumentManager.Instance.CurrentDBDocument
uidoc=DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = uidoc.ActiveView

collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Casework).WhereElementIsNotElementType().ToElements()

SELLINES=[]
for item in collector:
	ftype2 = item.Symbol
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
	enum1 = geo1.GetEnumerator(); enum1.MoveNext()
	geo2 = enum1.Current.GetInstanceGeometry()
	SET1=[]
	for g in geo2:
		if g.GetType() != Autodesk.Revit.DB.Solid:
			SET1.append(g.Convert())
	ELEMENT.append(a)		
	GEOMETRY.append(SET1)		
	IDSS.append(a.Id) 
#	NAMES.append(famsymb.Family.Name)
#	TYPES.append(famsymb.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString())


def shortline(ln):
	LN=Line.ByStartPointEndPoint(ln.PointAtParameter(0.3),ln.PointAtParameter(0.7))
	return LN
	
def listtostring(s):
	str1=""
	for ele in s:
		str1+=str(ele)
	return str1
	
def getlnsids(lines,rrids):
	CIL=[]
	for a in lines:
		CIL.append(Cylinder.ByPointsRadius(shortline(a).StartPoint,shortline(a).EndPoint,0.3))
	indices=range(len(lines))	
	intr=[]
	for a in indices:
		set1=[]
		for b in indices:
			if CIL[b].DoesIntersect(CIL[a]):
				set1.append(b)
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

OUT=LNSIDS1,LNSIDS2