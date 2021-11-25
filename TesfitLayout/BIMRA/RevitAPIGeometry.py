import clr
clr.AddReference("RevitAPIUI")
from  Autodesk.Revit.UI import *
clr.AddReference("RevitAPI")
import Autodesk
import math
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

#01 Point/Vector
a = XYZ(1,4,5)

#02 Line
b = Line.CreateBound(XYZ(0,0,2),XYZ(3,4,2))

#03 Curve
#def NurbsCurveByPoints(lpoints,weig1):
#	HS = HermiteSpline.Create(lpoints,True)
#	OUT1 = NurbSpline.CreateCurve(HS)
#	return OUT1
	
#points = []
#points.append(XYZ(0,0,2))
#points.append(XYZ(0,2,2))
#points.append(XYZ(0,5,2))
#points.append(XYZ(0,10,2))
#lpoints = [points]
#weig = [1,1,1,1]
#lweig = [weig]

#c=NurbsCurveByPoints(lpoints,IN[2])

#04 Superficie Reglada
line1=Line.CreateBound(XYZ(0,0,0),XYZ(5,0,0))
line2=Line.CreateBound(XYZ(4,5,2),XYZ(9,5,2))

d=RuledSurface.Create(line1,line2)

#05 Plano
e = Plane.CreateByNormalAndOrigin(XYZ(0,0,1),XYZ(0,0,0))

#06 Solido
def CuboidByLengths(c,dir,d1,d2,d3):
	profile = []
	profile00 = XYZ(c.X-d1/2,c.Y-d2/2,c.Z-d3/2)
	profile01 = XYZ(c.X-d1/2,c.Y+d2/2,c.Z-d3/2)
	profile11 = XYZ(c.X+d1/2,c.Y+d2/2,c.Z-d3/2)
	profile10 = XYZ(c.X+d1/2,c.Y-d2/2,c.Z-d3/2)
	profile.append( Line.CreateBound( profile00, profile01 ) )
	profile.append( Line.CreateBound( profile01, profile11 ) )
	profile.append( Line.CreateBound( profile11, profile10 ) )
	profile.append( Line.CreateBound( profile10, profile00 ) )
	curveloop = CurveLoop.Create(profile)
	#loop = CurveLoop()
	#for a in curveloop:
	#	loop.Append(a)
	loops = [curveloop]
	Gopt = SolidOptions(ElementId.InvalidElementId,ElementId.InvalidElementId)
	Cuboid = GeometryCreationUtilities.CreateExtrusionGeometry(loops,dir,d3,Gopt)
	return Cuboid
f=CuboidByLengths(XYZ(5,0,0),XYZ(0,0,1), 2, 2, 2 ).ToProtoType()	
	
#07 Frame + Revolve
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

g=SphereByCenterPointRadius(XYZ(0,0,0),2)
g2=SphereByCenterPointRadius(XYZ(0,0,1),2)

#08 Geometry Loft
#def SolidByLoft(sections):
#	curveloop = CurveLoop.Create(sections)
#	Gopt = SolidOptions(ElementId.InvalidElementId,ElementId.InvalidElementId)
#	loops = [curveloop]
#	surface = GeometryCreationUtilities.CreateLoftGeometry(loops,Gopt)
#	return surface
	
#lines=[]
#lines.append(Line.CreateBound(XYZ(0,0,0),XYZ(5,0,0)))
#lines.append(Line.CreateBound(XYZ(4,5,0),XYZ(9,5,0)))
#lines.append(Line.CreateBound(XYZ(0,10,0),XYZ(5,10,0)))

#h=SolidByLoft(lines)
#ElementIntersectsSolidFilter

#ElementIntersectsSolidFilter filterSphere = newElementIntersectsSolidFilter(tolSphere)
#FilteredElementCollector( doc ) .OfClass( typeof( Wall ) ) .WherePasses( filterSphere );

#09 Projection Perpendicular
p1 = XYZ(4,5,10)
def PointProject(p,pl):
	pp = pl.Project(p)
	pr = pl.Origin.Add(pl.XVec.Normalize().Multiply(pp[0].U)).Add(pl.YVec.Normalize().Multiply(pp[0].V))
	return pr	
p5 = PointProject(p1,d)

#10 Projection Vertical
def PointProjectDown(p,pl):
	planepoint = pl.Origin
	planenormal = pl.Normal
	lineDirection  = XYZ(0,0,-1).Normalize()
	lineParameter = (planenormal.DotProduct(planepoint)- planenormal.DotProduct(p))/ planenormal.DotProduct(lineDirection)
	return  p+lineParameter*lineDirection

p6 = PointProjectDown(p1,d)
#11 IntersectionSolidPlane
def IntersectSolids(sol,sol2):
	#SOLS = BooleanOperationsUtils.ExecuteBooleanOperation(sol,sol2,BooleanOperationsType.Union)
	SOLS = BooleanOperationsUtils.ExecuteBooleanOperation(sol,sol2,BooleanOperationsType.Intersect)
	#SOLS = BooleanOperationsUtils.ExecuteBooleanOperation(sol,sol2,BooleanOperationsType.Difference)
	return SOLS

h=IntersectSolids(g2,g)

def IntersectSolidPlane(sol,plane):
	#SOLS = BooleanOperationsUtils.ExecuteBooleanOperation(sol,sol2,BooleanOperationsType.Union)
	SOLS = BooleanOperationsUtils.CutWithHalfSpace(sol,plane)
	#SOLS = BooleanOperationsUtils.ExecuteBooleanOperation(sol,sol2,BooleanOperationsType.Difference)
	#SOLS = BooleanOperationsUtils.ExecuteBooleanOperation(sol,sol2,BooleanOperationsType.Interesect)
	return SOLS

i=IntersectSolidPlane(g,d)

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


OUT=d.ToPlane(),p1.ToPoint(),p5.ToPoint(),p6.ToPoint(),i.ToProtoType()