# Load the Python Standard and DesignScript Libraries
import sys
import clr
import math
clr.AddReference('ProtoGeometry')
from Autodesk.DesignScript.Geometry import *

# The inputs to this node will be stored as a list in the IN variables.
SURF = IN[0]
curves=SURF.PerimeterCurves()
points=[]

def GTL(SUR):
	PCUR=SUR.PerimeterCurves()
	LIN=PCUR[0]
	PY=LIN.PointAtParameter(0.5).Y
	for a in PCUR:
		if a.PointAtParameter(0.5).Y>PY:
			PY=a.PointAtParameter(0.5).Y
			LIN=a
	return(LIN)		

def GBL(SUR):
	PCUR=SUR.PerimeterCurves()
	LIN=PCUR[0]
	PY=LIN.PointAtParameter(0.5).Y
	for a in PCUR:
		if a.PointAtParameter(0.5).Y<PY:
			PY=a.PointAtParameter(0.5).Y
			LIN=a
	return(LIN)	
	

def CurvestoPoints(cur):
	for a in cur:
		points.append(a.StartPoint)
		points.append(curves[-1].EndPoint)
	
	xpo=[]
	ypo=[]
	zpo=[]
	for a in points:
		xpo.append(round(a.X,2))
		ypo.append(round(a.Y,2))
		zpo.append(round(a.Z,2))
	xpo=list(set(xpo))
	xpo.sort()
	
	ypo=list(set(ypo))
	ypo.sort()
	
	zpo=list(set(zpo))
	zpo.sort()
	
	minx=min(xpo)
	maxx=max(xpo)
	miny=min(ypo)
	maxy=max(ypo)
	minz=min(zpo)

	XGRIDS=[]
	YGRIDS=[]
	for a in xpo:
		LNE=Line.ByStartPointEndPoint(Point.ByCoordinates(a,miny,minz),Point.ByCoordinates(a,maxy,minz))		
		XGRIDS.append(LNE)
	for a in ypo:
		LNE=Line.ByStartPointEndPoint(Point.ByCoordinates(minx,a,minz),Point.ByCoordinates(maxx,a,minz))		
		YGRIDS.append(LNE)
	return XGRIDS,YGRIDS

GRIDS=CurvestoPoints(curves)
XGRIDS=GRIDS[0]
YGRIDS=GRIDS[1]
XX=[]
for a in YGRIDS:
	XBLO=[]
	for b in XGRIDS:
		if b.DoesIntersect(a):
			XBLO.append(b.Intersect(a)[0])
	XX.append(XBLO)
	
YY=[]
for a in XGRIDS:
	YBLO=[]
	for b in YGRIDS:
		if b.DoesIntersect(a):
			YBLO.append(b.Intersect(a)[0])
	YY.append(YBLO)
VAR="DATA"

SURFS=[]
for a,b in zip(XX,range(len(XX))):
	if b<len(XX)-1:
		SURB=[]
		for e,f in zip(a,range(len(a))):
			if f<len(a)-1:
				PON=[a[f],a[f+1],XX[b+1][f+1],XX[b+1][f]]
				SURB.append(Surface.ByPerimeterPoints(PON))
		SURFS.append(SURB)

RS=[]
for a in SURFS:
	RS1=[]
	for b in a:
		if Sphere.ByCenterPointRadius(b.PointAtParameter(0.5,0.5),0.1).DoesIntersect(SURF):
			RS1.append(b)
	RS.append(RS1)

NEWSURFS=[]
for a in RS:
	TEMPSURF=[]
	for c in range(len(a)):
		if c<len(a)-1:
			if a[c].DoesIntersect(a[c+1]):
				TEMPSURF.append(a[c])
			else:
				TEMPSURF.append(a[c])
				NEWSURFS.append(TEMPSURF)
				TEMPSURF=[]		
		else:
			TEMPSURF.append(a[c])				
	NEWSURFS.append(TEMPSURF)

ALLSURFS=[]
for a in NEWSURFS:
	SURFACE = Surface.ByUnion(a)
	ALLSURFS.append(SURFACE)	

def extend(list):
	a1=[]	
	for a in list:
		b1=[]
		for b in a:
			for c in (list[b]):
				if c not in b1:
					b1.append(c)
		a1.append(b1)
	return a1
	
def listtostring(s):
	str1=""
	for ele in s:
		str1+=str(ele)
	return str1

a1=[]
for a,b in zip(ALLSURFS,range(len(ALLSURFS))):
	a11=[]
	for c,d in zip(ALLSURFS,range(len(ALLSURFS))):
		if a.DoesIntersect(c):
			if GTL(a).Length==GBL(c).Length:
				a11.append(d)
	a1.append(a11)

a2=extend(extend(extend(extend(extend(a1)))))
a3=[]
for a in a2:
	a3.append(listtostring(a))

a4=list(set(a3))
a5=[]
for a in a4:
	a5.append(a2[a3.index(a)])

SSUR=[]
for a in a5:
	TT=[]
	for b in a:
		TT.append(ALLSURFS[b])
	SSUR.append(TT)

LLL=[]
for a in SSUR:
	SURFACE = Surface.ByUnion(a)
	LLL.append(SURFACE)


# Assign your output to the OUT variable.
OUT = LLL,SSUR