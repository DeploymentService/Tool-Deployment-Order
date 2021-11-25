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
	
def cylinders(l1):
	la=[]
	for a in l1:
		la.append(CreateCylinder(a.GetEndPoint(0),a.Direction,1,a.Length))
	return la

def extendedendpointlines(a):
	sl=a.GetEndPoint(0).Add(a.Direction.Normalize().Negate().Multiply(1))
	ll=a.GetEndPoint(0).Add(a.Direction.Normalize().Multiply(1))
	la=Line.CreateBound(sl,ll)
	sl1=a.GetEndPoint(1).Add(a.Direction.Normalize().Negate().Multiply(1))
	ll1=a.GetEndPoint(1).Add(a.Direction.Normalize().Multiply(1))
	lb=Line.CreateBound(sl1,ll1)
	return la,lb
	
def TranslateLine(Lnn,Vec,DS):
    SP=Lnn.GetEndPoint(0)
    EP=Lnn.GetEndPoint(1)
    FL=Line.CreateBound(SP.Add(Vec.Multiply(DS)),EP.Add(Vec.Multiply(DS)))
    return FL
    
def SphereByCenterPointRadius(c,r):
	frame = Frame(c,XYZ.BasisX,XYZ.BasisY,XYZ.BasisZ)
	arc = Arc.Create(c-r*XYZ.BasisZ,c+r*XYZ.BasisZ,c+r*XYZ.BasisX)
	line = Line.CreateBound(arc.GetEndPoint(1),arc.GetEndPoint(0))
	lines =[]
	lines.append(arc)
	lines.append(line)
	halfcircle = CurveLoop.Create(lines)
	cloop = [halfcircle]
	Sphere = GeometryCreationUtilities.CreateRevolvedGeometry(frame,cloop,0,2*math.pi)
	return Sphere
	
def GetAxes(nor):
	axis=[]
	limit = 1/64
	if (IsZero(nor.X,limit) and IsZero(nor.Y,limit)):
		cardinal=XYZ.BasisY
	else:
		cardinal=XYZ.BasisZ
	ax=cardinal.CrossProduct(nor).Normalize()
	ay=nor.CrossProduct(ax).Normalize()
	axis.append(ax)
	axis.append(ay)
	return axis

def IsZero(a,tol):
	return tol> abs(a)

def CreateCylinder(ori,vec,rad,hei):
	axes = GetAxes(vec)
	ax = axes[0]
	ay = axes[1]
	az = vec.Normalize()
	px = ori + rad * ax
	pxz = ori + rad * ax + hei * az
	pz = ori + hei * az
	linesa =[]
	linesa.append(Line.CreateBound(ori,px))
	linesa.append(Line.CreateBound(px,pxz))
	linesa.append(Line.CreateBound(pxz,pz))
	linesa.append(Line.CreateBound(pz,ori))
	profile = CurveLoop.Create(linesa)
	cloop = [profile]
	frame = Frame(ori,ax,ay,az)
	cil = GeometryCreationUtilities.CreateRevolvedGeometry(frame,cloop,0,2*math.pi)
	return cil
	


def getPlanefromMesh(surf):
	v1 = surf.Vertices
	plane = Plane.CreateByThreePoints(v1[0].PointGeometry.ToRevitType(),v1[1].PointGeometry.ToRevitType(),v1[2].PointGeometry.ToRevitType())
	return plane
	
def gvpfromline(ln):
	p1 = ln.GetEndPoint(0)
	p2 = ln.GetEndPoint(1)
	p3 = p1.Add(XYZ(0,0,-50))
	plane = Plane.CreateByThreePoints(p1,p2,p3)
	return plane
	
def PointProjectDown(p,pl):
	planepoint = pl.Origin
	planenormal = pl.Normal
	lineDirection  = XYZ(0,0,-1).Normalize()
	lineParameter = (planenormal.DotProduct(planepoint)- planenormal.DotProduct(p))/ planenormal.DotProduct(lineDirection)
	return  p+lineParameter*lineDirection

def getBorderPlanes(FLOOR):
	EDGES = FLOOR.Geometry()[0].Faces[1].Edges
	Curves=[]
	for a in EDGES:
		p1= a.StartVertex.PointGeometry.ToRevitType()
		p2= a.EndVertex.PointGeometry.ToRevitType()
		Curves.append(Line.CreateBound(p1,p2))
	return Curves

def getCurveIntersection(pl,cur):
	projection=DistanceTo(pl,cur.GetEndPoint(0))
	l1 = Line.CreateBound((projection +  cur.Length)*(pl.XVec+pl.YVec),(projection +cur.Length)*(pl.XVec-pl.YVec))
	l2 = Line.CreateBound((projection +  cur.Length)*(-pl.XVec+pl.YVec),(projection +cur.Length)*(-pl.XVec-pl.YVec))
	cl1 = CurveLoop()
	cl1.Append(l1)
	cl2 = CurveLoop()
	cl2.Append(l2)
	lll = [cl1,cl2]
	Gopt = SolidOptions(ElementId.InvalidElementId,ElementId.InvalidElementId)
	so = GeometryCreationUtilities.CreateLoftGeometry(lll,Gopt)
	return so.ToProtoType()
	
def DistanceTo(pl,po):
	tr = Transform.CreateTranslation(pl.Origin)
	tr.BasisX =pl.XVec.Normalize()
	tr.BasisY =pl.YVec.Normalize()
	tr.BasisZ =pl.Normal.Normalize()
	local = tr.Inverse.OfPoint(po)
	projection = XYZ(local.X,local.Y,0)
	projection = tr.OfPoint(projection)
	return local.Z

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


SPH=[]
for a in POINS:
	SPH.append(SphereByCenterPointRadius(a,1))
	
NL=cylinders(shortline(RLINES))

CLEARPOINTS=[]
CLEARLINES=[]
BOB=[]
BOBD=[]
TEST=[]
for a,b,c in zip(SPH,POINS,PLINE):
	X=0
	for d in NL:
		try:
			SOLS = BooleanOperationsUtils.ExecuteBooleanOperation(a,d,BooleanOperationsType.Intersect)
			if SOLS.Volume > 0.00:
				X=1				
		except:
			pass
	if X!=1:
		CLEARPOINTS.append(b)
		CLEARLINES.append(c)		
	else:
		BOB.append(b)
		BOBD.append(c.Direction)
		

FLOOR=IN[0]
BORDPLANES=getBorderPlanes(FLOOR)
SUPCURVES = []
for a in BORDPLANES:
	SUPCURVES.append(gvpfromline(a))


SURFACE=FLOOR.Geometry()[0].Faces[1].SurfaceGeometry()
PLANE = getPlanefromMesh(SURFACE)

PPOI=[]
PPOI2=[]
P1LEN=[]
P2LEN=[]

for a in range(len(CLEARPOINTS)/2):
	PR1 = PointProjectDown(CLEARPOINTS[a*2],PLANE)
	PR2 = PointProjectDown(CLEARPOINTS[a*2+1],PLANE)
	SF1 = SphereByCenterPointRadius(PR1,1)
	SF2 = SphereByCenterPointRadius(PR2,1)
	SOLS = BooleanOperationsUtils.CutWithHalfSpace(SF1,PLANE)
	SOLS2 = BooleanOperationsUtils.CutWithHalfSpace(SF2,PLANE)

	if SOLS.Volume != SF1.Volume or SOLS2.Volume != SF2.Volume:
		PPOI.append(CLEARPOINTS[a*2])
		P1LEN.append(CLEARPOINTS[a*2].DistanceTo(PR1))
		PPOI.append(CLEARPOINTS[a*2+1])
		P1LEN.append(CLEARPOINTS[a*2+1].DistanceTo(PR2)) 
			
""""							
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
"""FPOI=[]
FLEN=[]
PINDEX=[]
for a,b in zip(PPOI,P1LEN):
	if str(str(a.X)+str(a.Y)) not in PINDEX:
		FPOI.append(a)
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

# Assign your output to the OUT variable.
OUT = TEST1,TEST2