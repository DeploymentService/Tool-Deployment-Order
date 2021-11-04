import clr
import os
import json
import math
clr.AddReference("RevitAPIUI")
from  Autodesk.Revit.UI import *
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Architecture import *
 
clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

clr.AddReference("RevitNodes")
import Revit
clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)



def cleanlines(sel):
	lines=[]
	for a in range(len(sel)):
		if a<len(sel)-1:
			lines.append(Line.CreateBound(sel[a],sel[a+1]))	
	lines.append(Line.CreateBound(sel[-1],sel[0]))
	dir=[]
	for a in range(len(lines)):
		dir.append(lines[a].Direction.AngleTo(XYZ(1,0,0)))	
	dir2 = []
	for a in range(len(dir)):
		if a<len(dir)-2:
			dir2.append(dir[a+1])
	dir2.append(dir[-1])
	dir2.append(dir[0])	

	dir3=[]
	for a in range(len(dir2)):
		dir3.append(abs(dir[a]-dir2[a]))

	fp=[]
	index=[]
	for a in range(len(dir3)):
		if dir3[a]>0.055:
			index.append(a)	
	for a in range(len(index)):
		fp.append(sel[a-1])
	return fp


def tolines(lst):
	LINES = []
	for a in range(len(lst)):
		if a<len(lst)-1:
			LINES.append(Line.CreateBound(lst[a],lst[a+1]))
	LINES.append(Line.CreateBound(lst[-1],lst[0]))
	return LINES

def PlaneFunction(LINE):
	p = LINE.GetEndPoint(0)
	q = LINE.GetEndPoint(1)
	norm = XYZ.BasisZ
	plane = Plane.CreateByNormalAndOrigin(norm,p)
	return plane;

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
	
def IsZero(a,tol):
	return tol> abs(a)

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

def TranslatePoint(Pt,Vec,DS):
    FL=Pt.Add(Vec.Multiply(DS))
    return FL
    
def rotate(l, n):
	return l[n:] + l[:n]

doc = DocumentManager.Instance.CurrentDBDocument
directoryPath = "C:/Users/e2133/source/repos/05_RaDABIM360/wwwroot/js/"
files = os.listdir(directoryPath)
pathsFamily=[]
for a in range(len(files)):
	if "json" in files[a]:
		pathsFamily.append(directoryPath+files[a])
f=open(pathsFamily[0])
jsonread = json.load(f)

topo = jsonread["topography"]
lots = jsonread["lotline"]
neib = jsonread["neiborhood"]
tp=[]
for a in topo:
	tp.append(XYZ(a["X"],a["Y"],a["Z"]))
	
sl=[]
for a in lots:
	sl.append(XYZ(a["X"],a["Y"],0))
ll = cleanlines(sl)

nb=[]
for a in neib:
	nblot=[]
	for b in a:
		nblot.append(XYZ(b["X"],b["Y"],0))
	nb.append(cleanlines(nblot))

TransactionManager.Instance.EnsureInTTransaction(doc) 
topo = TopographySurface.Create(doc,tp)
TransactionManager.Instance.TransactionTaskDone() 


doc = DocumentManager.Instance.CurrentDBDocument
uidoc=DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument 
nei=ll
nei2=nb
    
SLIN=tolines(nei)	
ALLCURVES=[]

LINESET=[]
for a in nei2:
	LINESET.append(tolines(a))

CILINDS =[]
for a in SLIN:
	CILINDS.append(CreateCylinder(a.GetEndPoint(0),a.Direction,1,a.Length))

SPHERES =[]
ALLLINES=[]
for a in LINESET:
	for b in a:
		SPHERES.append(SphereByCenterPointRadius(TranslatePoint(b.GetEndPoint(0),b.Direction,b.Length/2),2))
		ALLLINES.append(b)

INTERSECTS=[]
for a in CILINDS:	 
	X=0
	for b in SPHERES:
		SOLS = BooleanOperationsUtils.ExecuteBooleanOperation(a,b,BooleanOperationsType.Intersect)
		if SOLS.Volume>0:
			X+=1;
	INTERSECTS.append(X)
	
CLEANLINES=[]
for a in range(len(SPHERES)):
	X=0
	for b in CILINDS:
		SOLS = BooleanOperationsUtils.ExecuteBooleanOperation(SPHERES[a],b,BooleanOperationsType.Intersect)
		if SOLS.Volume>0:
			X+=1;
	if X==0:
		CLEANLINES.append(ALLLINES[a]) 
	
smallestcurve=0
for a in range(INTERSECTS.Count):
	if INTERSECTS[a] == 1:
		if SLIN[a].Length<len: 
			len=SLIN[a].Length
			smallestcurve=a
	
style = ["Front Line","Side Line","Rear Line", "Flunkage Line"]
nstyle = rotate(style,-smallestcurve)	
styles =FilteredElementCollector(doc).OfClass(GraphicsStyle).WhereElementIsNotElementType().ToElements()
RS=[]
for a in style:
	for b in styles:
		if a==b.Name:
			RS.append(b) 


TransactionManager.Instance.EnsureInTransaction(doc)
for a in range(len(SLIN)):
	crv=doc.Create.NewModelCurve(SLIN[a],SketchPlane.Create(doc,PlaneFunction(SLIN[a])))
	crv.LineStyle=RS[a]
	ALLCURVES.append(crv)
for b in range(len(CLEANLINES)):
	crv2=doc.Create.NewModelCurve(CLEANLINES[b],SketchPlane.Create(doc,PlaneFunction(CLEANLINES[a])))

TransactionManager.Instance.TransactionTaskDone()

OUT=ALLLINES,CLEANLINES