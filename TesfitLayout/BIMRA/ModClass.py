import sys
import clr
import math
import random

clr.AddReference('ProtoGeometry')
from Autodesk.DesignScript.Geometry import *

#Create MOD CLASS
class MOD:
	def _init_(self,IOF,SOL,ORI,CEN,WID,LEN,HEI,USE,SC1,SC2,SC3,SC4,SC5,SC6,CSYS,VEC):
		self.IOF=IOF
		self.SOL=SOL
		self.ORI=ORI
		self.CEN=CEN
		self.WID=WID
		self.LEN=LEN
		self.HEI=HEI
		self.USE=USE
		self.SC1=SC1	
		self.SC2=SC2
		self.SC3=SC3
		self.SC4=SC4
		self.SC5=SC5
		self.SC6=SC6
		self.CSYS=CSYS
		self.VEC=VEC
#Create GRID OF ELEMENTS
def MODGRID(LINE,OLINE,VEC):
	NAL=float(int(LINE.Length/16))
	NALL=float(int((((LINE.Length)-(NAL*16))/8)))
	NALLL=(LINE.Length)-(NAL*16)-(NALL*8)
	TOT=(NAL*16+NALL*8)
	MED=TOT/2
	DIF=NALLL/2
	STARTLINELL=OLINE.Translate(VEC.Reverse(),MED)
	VALUES = []
	for a in range(int(NAL)):
		VALUES.append(16)
	DIS=range(8,int(NAL*16),16)
	if NALL>0:
		DIS.append(int(NAL*16)+4)
		VALUES.append(8)
	GRIDLINES=[]
	for a in DIS:
		NL=STARTLINELL.Translate(VEC,a)
		GRIDLINES.append(NL)	
	return GRIDLINES,VALUES

def MODGRID2(LINE,OLINE,VEC):
	NAL=float(int(LINE.Length/16))-1
	NALL=float(int((((LINE.Length)-(NAL*16))/8)))
	NALLL=(LINE.Length)-(NAL*16)-(NALL*8)
	TOT=(NAL*16+NALL*8)
	MED=TOT/2
	DIF=NALLL/2
	STARTLINELL=OLINE.Translate(VEC.Reverse(),MED)
	VALUES = []
	DIS=[]
	if NALL>2:
		VALUES.append(8)
		X=4
		DIS.append(4)
		X+=12
		for a in range(int(NAL)):
			VALUES.append(16)
			DIS.append(X)
			X+=16			
		VALUES.append(16)		
		DIS.append(X)
	
	elif NALL>1:
		VALUES.append(8)
		X=4
		DIS.append(4)
		X+=12
		for a in range(int(NAL)):
			VALUES.append(16)
			DIS.append(X)
			X+=16			
		VALUES.append(8)
		X-=4
		DIS.append(X)
		
	GRIDLINES=[]
	for a in DIS:
		NL=STARTLINELL.Translate(VEC,a)
		GRIDLINES.append(NL)	
	return GRIDLINES,VALUES


#Create the Modules with their Neiborhoods
def MODNEI(XG,YG,XVEC,YVEC):
	FORMACRO=[]
	X=1
	for e in range(2):
		Y=e*12
		for a,b,f in zip(XG[0],XG[1],range(int(len(XG[0])))):
			for c,d,g in zip(YG[0],YG[1],range(int(len(YG[0])))):
				M1=MOD()
				M1.IOF=X
				M1.ORI=a.Intersect(c)[0].Translate(Vector.ByCoordinates(0,0,Y))
				M1.WID=d
				M1.LEN=b
				M1.HEI=12
				M1.CEN=M1.ORI.Translate(Vector.ByCoordinates(0,0,M1.HEI/2))
				M1.USE=0
				M1.CSYS=CoordinateSystem.ByOriginVectors(M1.CEN,XVEC,YVEC,Vector.ByCoordinates(0,0,1))
				M1.SOL=Cuboid.ByLengths(M1.CSYS,M1.WID,M1.LEN,M1.HEI)	
				M1.VEC=Vector.ByCoordinates(0,1,0)	
				if f==0:
					M1.SC1=0
				else:			
					M1.SC1=((f-1)*(len(YG[0])))+g+1+(e*(len(XG[0]))*(len(YG[0])))
				if f==len(XG[0])-1:
					M1.SC3=0
				else:
					M1.SC3=((f+1)*(len(YG[0])))+g+1+(e*(len(XG[0]))*(len(YG[0])))
				if g==len(YG[0])-1:
					M1.SC2=0
				else:			
					M1.SC2=(f*(len(YG[0])))+g+2+(e*(len(XG[0]))*(len(YG[0])))
				if g==0:
					M1.SC4=0
				else:			
					M1.SC4=(f*(len(YG[0])))+g+(e*(len(XG[0]))*(len(YG[0])))
				if e==0:
					M1.SC5=0
				else:			
					M1.SC5=(f*(len(YG[0])))+g+1
				if e==1:
					M1.SC6=0
				else:			
					M1.SC6=(f*(len(YG[0])))+g+1+(len(XG[0]))*(len(YG[0]))
				FORMACRO.append(M1)
				X=X+1
	return FORMACRO

def MACUBOID(csys,W,L,H):
	CBS=Cuboid.ByLengths(csys,L,W,H)
	return CBS

class MACRO:
	def _init_(self,IOF,MIOF,CSYS,MODS,WID,LEN,HEI,VEC,ORI,CEN,SOL,SIZE,NAM,MM):
		self.IOF=IOF
		self.MIOF=MIOF
		self.CSYS=CSYS
		self.MODS=MODS
		self.WID=WID
		self.LEN=LEN
		self.HEI=HEI
		self.VEC=VEC
		self.ORI=ORI
		self.CEN=CEN
		self.SOL=SOL
		self.SIZE=SIZE
		self.NAM=NAM
		self.MM=MM

def allocinit(mod,mods,hei,size,name,MODALLOC,SELIND):
	SA=mod
	GAR=macalloc(0,mods,SA.WID,SA.LEN,hei,size,name)
	GAR=macnite(GAR,mod)
	MODALLOC.append(GAR)
	SELIND.append(SA.IOF)

def macalloc(IOF,MODS,W,L,H,SIZE,NAM):
	fur1=MACRO()
	fur1.IOF=IOF
	fur1.MODS=MODS
	fur1.WID=W
	fur1.LEN=L
	fur1.HEI=H
	fur1.SIZE=SIZE
	fur1.NAM=NAM
	return fur1

def macnite(MAC,MOD):
	MAC.MIOF=MOD.IOF
	MAC.CSYS=MOD.CSYS
	MAC.VEC=MOD.VEC
	MAC.ORI=MOD.ORI
	MAC.CEN=MOD.CEN
	MAC.SOL=Cuboid.ByLengths(MAC.CSYS,MAC.WID,MAC.LEN,MAC.HEI)
	MAC.MM=MOD
	return MAC

def selfmacro(name,poi,w,l,h,iof):
	MM=MACRO()
	MM.IOF=iof
	MM.MIOF=0
	MM.NAM=name
	MM.ORI=poi
	MM.WID=w
	MM.LEN=l
	MM.HEI=h
	MM.MODS=(w/2)*(l/2)
	MM.VEC=Vector.ByCoordinates(0,1,0)
	CEN=poi.Translate(Vector.ByCoordinates(0,0,h/2))
	MM.CEN=CEN
	MM.CSYS=CoordinateSystem.ByOriginVectors(CEN,Vector.ByCoordinates(1,0,0),Vector.ByCoordinates(0,1,0),Vector.ByCoordinates(0,0,1))
	MM.SIZE=l*w
	MM.MM=0
	MM.SOL=Cuboid.ByLengths(MM.CSYS,MM.WID,MM.LEN,MM.HEI)
	return MM
	

def NewMacrobyMacro(MACRO,CSYS,WID,LEN,SOL):
	NM=macalloc(MACRO.IOF,MACRO.MODS,WID,LEN,MACRO.HEI,WID*LEN,MACRO.NAM)
	NM.MIOF=MACRO.MIOF
	NM.CSYS=CSYS
	NM.VEC=MACRO.VEC
	NM.CSYS=CSYS
	NM.SOL=SOL
	NM.MM=MACRO.MM
	NM.CEN=CSYS.Origin
	NM.ORI=NM.CEN.Translate(Vector.ByCoordinates(0,0,-MACRO.HEI/2))
	return NM

def NewMacrobySurfandName(MACRO,SURF,NAME,XX):
	ORI=SURF.PointAtParameter(0.5,0.5)
	WID=SURF.Intersect(Plane.ByOriginNormal(ORI,Vector.ByCoordinates(0,1,0)))[0].Length
	LEN=SURF.Intersect(Plane.ByOriginNormal(ORI,Vector.ByCoordinates(1,0,0)))[0].Length
	NM=macalloc(XX,SURF.Area/4,WID,LEN,MACRO.HEI,WID*LEN,NAME)
	NM.MIOF=XX	
	NM.VEC=MACRO.VEC
	CEN=ORI.Translate(Vector.ByCoordinates(0,0,MACRO.HEI/2))
	NM.CSYS=CoordinateSystem.ByOriginVectors(CEN,Vector.ByCoordinates(1,0,0),Vector.ByCoordinates(0,1,0),Vector.ByCoordinates(0,0,1))
	NM.CEN=CEN
	NM.SOL=Cuboid.ByLengths(CEN,WID,LEN,MACRO.HEI)
	NM.ORI=ORI
	NM.MM=MACRO.MM
	return NM

def outoflist(li,sl):
	for a in li:
		sl.remove(a)
	
def cleandic(list3):
	list2=[]
	for a in list3:
		list2.append(DSCA.get(str(a)))
	POST=[]
	PRE=[]
	PRE.append(0)
	for a in range(len(list2)):
		if a+1<len(list2):
			if list2[a+1]+1!=(list2[a]):
				POST.append(a)
				PRE.append(a+1)
			else:
				pass
	POST.append(len(list2)-1)
	VAL=[]
	for a,b in zip(PRE,POST):
		X=b-a
		for c in range(b-a+1):
			VAL.append(X)
			X=X-1
	return VAL
	
def listtostring(s):
	str1=""
	for ele in s:
		str1+=str(ele)
	return str1

class MINI:
	def _init_(self,IOF,MIOF,CSYS,VEC,ORI,CEN,SC1,SC2,SC3,SC4,SC5,SC6,LOC):
		self.IOF=IOF
		self.MIOF=MIOF
		self.CSYS=CSYS
		self.VEC=VEC
		self.ORI=ORI
		self.CEN=CEN
		self.SC1=SC1	
		self.SC2=SC2
		self.SC3=SC3
		self.SC4=SC4
		self.SC5=SC5
		self.SC6=SC6
		self.LOC=LOC

def cuboidA(csys):
	CBS=Cuboid.ByLengths(csys,2,2,0.)
	return CBS

def SurfOFLINES(Surf):
	PerCur=Surf.PerimeterCurves()
	XVALUES=[]
	YVALUES=[]
	ZVALUES=[]
	for a in PerCur:
		XVALUES.extend([round(a.StartPoint.X,2),round(a.EndPoint.X,2)])
		YVALUES.extend([round(a.StartPoint.Y,2),round(a.EndPoint.Y,2)])
		ZVALUES.extend([round(a.StartPoint.Z,2),round(a.EndPoint.Z,2)])
	XVALUES=list(set(XVALUES))
	XVALUES.sort()
	YVALUES=list(set(YVALUES))
	YVALUES.sort()
	ZVALUES=list(set(ZVALUES))
	ZVALUES.sort()
	L1=Line.ByStartPointEndPoint(Point.ByCoordinates(XVALUES[0],YVALUES[0],ZVALUES[0]),Point.ByCoordinates(XVALUES[-1],YVALUES[0],ZVALUES[0]))
	L2=Line.ByStartPointEndPoint(Point.ByCoordinates(XVALUES[0],YVALUES[0],ZVALUES[0]),Point.ByCoordinates(XVALUES[0],YVALUES[-1],ZVALUES[0]))
	return L1,L2

def MACTOGRIDS(mod):
	ORIA=mod.ORI.Translate(Vector.ByCoordinates(-mod.WID/2,-mod.LEN/2,0))
	XMA=mod.ORI.Translate(Vector.ByCoordinates(mod.WID/2,-mod.LEN/2,0))
	YMA=mod.ORI.Translate(Vector.ByCoordinates(-mod.WID/2,mod.LEN/2,0))
	XLINE=Line.ByStartPointEndPoint(ORIA,XMA)
	YLINE=Line.ByStartPointEndPoint(ORIA,YMA)
	
	VX=Vector.ByTwoPoints(XLINE.StartPoint,XLINE.EndPoint)
	VY=Vector.ByTwoPoints(YLINE.StartPoint,YLINE.EndPoint)
	LLX=round(XLINE.Length)/2
	LLY=round(YLINE.Length)/2
	LXGRID=[]
	LYGRID=[]
	for a in range(int(LLX)):
		L1=YLINE.Translate(VX,(a*2)+1)
		LXGRID.append(L1)
	for a in range(int(LLY)):
		L1=XLINE.Translate(VY,(a*2)+1)
		LYGRID.append(L1)
	return LXGRID,LYGRID,LLX,LLY,XLINE.Length,YLINE.Length,mod.WID,mod.LEN,mod.ORI

def CREATEMICROS(LXGRID,LYGRID,mod,XVEC,YVEC,lst):
	X=0
	for a,b in zip(LYGRID,range(int(len(LYGRID)))):
		for c,d in zip(LXGRID,range(int(len(LXGRID)))):
			M1=MINI()
			M1.IOF=X
			M1.MIOF=mod.IOF
			M1.ORI=a.Intersect(c)[0]
			M1.CEN=M1.ORI.Translate(Vector.ByCoordinates(0,0,0.1))
			M1.CSYS=CoordinateSystem.ByOriginVectors(M1.CEN,XVEC,YVEC,Vector.ByCoordinates(0,0,1))
			M1.SOL=Cuboid.ByLengths(M1.CSYS,2,2,0.2)	
			M1.VEC=Vector.ByCoordinates(0,1,0)	
			M1.LOC=mod.NAM		
			lst.append(M1)
			if b==0:
				M1.SC1=0
			else:			
				M1.SC1=((b-1)*(len(LYGRID)))+d+1
			if b==len(LXGRID)-1:
				M1.SC3=0
			else:
				M1.SC3=((b+1)*(len(LYGRID)))+d+1
			if d==len(LYGRID)-1:
				M1.SC2=0
			else:			
				M1.SC2=(b*(len(LYGRID)))+d+2
			if d==0:
				M1.SC4=0
			else:			
				M1.SC4=(b*(len(LYGRID)))+d
			M1.SC5=0
			M1.SC6=0
			X=X+1

def allmacrosnamed(lst,nam):
	NAMEDMAC=[]
	for a in lst:
		if a.NAM==nam:
			NAMEDMAC.append(a)
	return NAMEDMAC		

def unifiedvalues(LST):
	NMA=MACRO()
	ALLSOLIDS=[]
	MODS=0
	SIZE=0
	CX=0
	CY=0
	CZ=0
	CZO=0
	if isinstance(LST,list):
		for a,b in zip(LST,range(len(LST))):
			MODS+=a.MODS
			HEI=a.HEI
			SIZE+=a.SIZE
			CX+=a.CEN.X
			CY+=a.CEN.Y
			CZ+=a.CEN.Z
			CZO+=a.ORI.Z
			ALLSOLIDS.append(a.SOL)
		CEN=Point.ByCoordinates(CX/len(LST),CY/len(LST),CZ/len(LST))
		ORI=Point.ByCoordinates(CEN.X,CEN.Y,CZO/len(LST))	
		NAM=LST[0].NAM
		MM=LST[0].MM
	else:
		MODS=LST.MODS
		HEI=LST.HEI
		SIZE=LST.SIZE
		ALLSOLIDS.append(LST.SOL)
		CEN=LST.CEN
		ORI=LST.ORI	
		NAM=LST.NAM
		MM=LST.MM
		
	SOL=Solid.ByUnion(ALLSOLIDS)	
	VEC=Vector.ByCoordinates(1,0,0)
	PLA=Plane.ByOriginNormal(CEN,Vector.ByCoordinates(0,0,1))
	INTSUR=SOL.Intersect(PLA)[0]
	PERCUR=INTSUR.PerimeterCurves()
	POLYC=PolyCurve.ByJoinedCurves(PERCUR)
	SURR=Surface.ByPatch(POLYC)
	CENPOI=SURR.PointAtParameter(0.5,0.5)
	ORI=CENPOI.Translate(Vector.ByCoordinates(0,0,-HEI/2))
	PLANX=Plane.ByOriginNormal(CENPOI,Vector.ByCoordinates(0,1,0))
	LX=SURR.Intersect(PLANX)[0]
	PLANY=Plane.ByOriginNormal(CENPOI,Vector.ByCoordinates(1,0,0))
	LY=SURR.Intersect(PLANY)[0]
	CSYS=CoordinateSystem.ByOriginVectors(CENPOI,Vector.ByCoordinates(1,0,0),Vector.ByCoordinates(0,1,0),Vector.ByCoordinates(0,0,1))
	WID=LX.Length
	LEN=LY.Length
	NMA.IOF=0
	NMA.MODS=MODS
	NMA.WID=round(WID,0)
	NMA.LEN=round(LEN,0)
	NMA.HEI=HEI
	NMA.SIZE=SIZE
	NMA.NAM=NAM
	NMA.MM=MM
	NMA.MIOF=0
	NMA.CSYS=CSYS
	NMA.VEC=VEC
	NMA.ORI=ORI
	NMA.CEN=CENPOI
	NMA.SOL=Cuboid.ByLengths(CSYS,WID,LEN,12)
	return NMA	

def getfarestlnloc(LOCATIONS):
	DISTANCE=0
	LINE=Line.ByStartPointEndPoint(Point.ByCoordinates(0,0,0),Point.ByCoordinates(1,0,0))
	for a,b in zip(LOCATIONS,range(len(LOCATIONS))):
		for c,d in zip(LOCATIONS,range(len(LOCATIONS))):
			if not b==d:
				if a.DistanceTo(c)>DISTANCE:
					DISTANCE=a.DistanceTo(c)
					LINE=Line.ByStartPointEndPoint(a,c)
	return LINE

def ClosestCornerstoLine(RECS,LINE):
	CORNERS=[]
	for a in RECS:
		CORNERS.append(a.Corners())
		
	SELCORNERS=[]	
	for a in CORNERS:
		DIS=50
		POI=a[0]
		for b in a:
			if b.DistanceTo(LINE)<DIS:
				DIS=b.DistanceTo(LINE)
				POI=b
		SELCORNERS.append(POI)
	return SELCORNERS

def ROUTES(VP):
	ROUTES=[]
	for b in range(len(VP)):
		if not (b+1)==len(VP):
			if b==0:
				LINEA=Line.ByStartPointEndPoint(VP[b],VP[b+1])
				LINEB=Line.ByStartPointEndPoint(LINEA.PointAtParameter(0.01),LINEA.EndPoint)
			elif b==len(VP)-1:
				LINEA=Line.ByStartPointEndPoint(VP[b],VP[b+1])
				LINEB=Line.ByStartPointEndPoint(LINEA.StartPoint,LINEA.PointAtParameter(0.99))
			else:
				LINEB=Line.ByStartPointEndPoint(VP[b],VP[b+1])
			ROUTES.append(LINEB)
	return ROUTES

def extendLine(ln,dis):
	ln2=Line.ByStartPointEndPoint(ln.StartPoint,ln.EndPoint)
	PointA=ln2.StartPoint.Translate(ln2.Direction.Reverse(),dis)
	PointB=ln2.EndPoint.Translate(ln2.Direction,dis)
	LLN=Line.ByStartPointEndPoint(PointA,PointB)
	return LLN

def shortline(lst):
	NEWLINES=[]
	for a in lst:
		NEWLINES.append(Line.ByStartPointEndPoint(a.PointAtParameter(0.1),a.PointAtParameter(0.90)))
	return NEWLINES

def MODULESFROMLINES(ALLCURVES,n):
	ML=[]
	MA=[]
	LEN=[]
	for a in ALLCURVES:
		if a.Length<n:
			ML.append(TOL(a))
			MA.append(TOL(a).Direction)
			LEN.append(a.Length)
		else:
			DIVLEN=a.Length/n
			POIS=[]
			for b in range(int(DIVLEN)+1):
				POIS.append(a.PointAtSegmentLength(b*n))
			LINESS=a.SplitByPoints(POIS)
			for c in LINESS:
				if c.Length>1:
					ML.append(TOL(c).PointAtParameter(0.5))
					MA.append(TOL(c).Direction)
					LEN.append(round(c.Length))	
	return ML,MA,LEN	

def TOL(ln):
	NEW=Line.ByStartPointEndPoint(ln.StartPoint,ln.EndPoint)
	return NEW

def curlstTOL(ln):
	LNS=[]
	for a in ln:
		LNS.append(TOL(a))
	return LNS

def solstoCurves(SOLS,off):
	SOL=Solid.ByUnion(SOLS)
	P1=Plane.ByOriginNormal(Point.ByCoordinates(0,0,off),Vector.ByCoordinates(0,0,1))
	SUR=SOL.Intersect(P1)[0]
	PER=SUR.PerimeterCurves()
	return SUR

def solstoCurvesFS(SOLS,off):
	SOL=Solid.ByUnion(SOLS)
	P1=Plane.ByOriginNormal(Point.ByCoordinates(0,0,off),Vector.ByCoordinates(0,0,1))
	SUR=SOL.Intersect(P1)[0]
	TP=SUR.Edges
	PER=[]
	for a in TP:
		PER.append(a.CurveGeometry)
	return PER 

def soltoCurves(SOL,off):
	P1=Plane.ByOriginNormal(Point.ByCoordinates(0,0,off),Vector.ByCoordinates(0,0,1))
	SUR=SOL.Intersect(P1)[0]
	PER=SUR.PerimeterCurves()
	return PER

def getmodulegrids(m):
	XPLANE = Plane.ByOriginNormal(m.CEN,Vector.ByCoordinates(0,1,0))
	YPLANE = Plane.ByOriginNormal(m.CEN,Vector.ByCoordinates(1,0,0))
	ZPLANE = Plane.ByOriginNormal(m.CEN,Vector.ByCoordinates(0,0,1))
	ZINT=m.SOL.Intersect(ZPLANE)
	YINT=ZINT[0].Intersect(YPLANE)
	XINT=ZINT[0].Intersect(XPLANE)
	return XINT[0],YINT[0]

def MACROTOSOL(lst):
	SOLIDS=[]
	for a in lst:
		SOLIDS.append(a.SOL)
	return SOLIDS

def CirculationPath(MACLIST):
	MODULES=[]
	LOCATIONS=[]
	RECS=[]
	LOCATIONS.append(MACROCENTERPOINTS(MACLIST))
	for a in MACLIST:
		LOCATIONS.append(a.ORI)
		RECS.append(Rectangle.ByWidthLength(a.CSYS.Translate(Vector.ByCoordinates(0,0,-a.HEI/2)),a.WID,a.LEN))
		MODULES.append(a)
	#Get the Farest Line Location
	LINE=getfarestlnloc(LOCATIONS)
	SELCORNERS = ClosestCornerstoLine(RECS,LINE)
	CLOP=[]
	for a in SELCORNERS:
		CLOP.append(round(LINE.ParameterAtPoint(LINE.ClosestPointTo(a)),4))
	CLOP2= list(set(CLOP))
	CLOP2.sort()
	POINTSF=[]
	for a in CLOP2:
		POINTSF.append(SELCORNERS[CLOP.index(a)])
	POLY=PolyCurve.ByJoinedCurves(ROUTES(POINTSF))	
	#FCURVE=POLY.Offset(2)
	return POLY,LINE.PointAtParameter(0.5)

def CIRCS(lst):
	MACROS=[]
	for a in lst:
		if a.NAM!="LIVING" and a.NAM!="GARAGE" and a.NAM!="ENTRANCE":
			MACROS.append(a)
	return MACROS

def MACROCENTERPOINTS(lst):
	X=0
	Y=0
	Z=0
	for a in lst:
		X=X+a.ORI.X
		Y=Y+a.ORI.Y
		Z=Z+a.ORI.Z
	TX=X/len(lst)
	TY=Y/len(lst)
	TZ=Z/len(lst)
	PP=Point.ByCoordinates(TX,TY,TZ)
	return PP

def GETLXLYBYNAME(MACROS,NAME):
	for a in MACROS:
		if (a.NAM == NAME) :
			PLANEX=Plane.ByOriginNormal(a.CEN,Vector.ByCoordinates(0,1,0))
			PLANEY=Plane.ByOriginNormal(a.CEN,Vector.ByCoordinates(1,0,0))
			PLANEZ=Plane.ByOriginNormal(a.CEN,Vector.ByCoordinates(0,0,1))
			SURF=a.SOL.Intersect(PLANEZ)[0]
			LX=SURF.Intersect(PLANEX)[0]
			LY=SURF.Intersect(PLANEY)[0]
			LLX=LX.Translate(Vector.ByCoordinates(0,-LY.Length/2,-6))
			LLY=LY.Translate(Vector.ByCoordinates(-LX.Length/2,0,-6))
	return LLX,LLY

def GETINDLXLYBYNAME(a):
	PLANEX=Plane.ByOriginNormal(a.CEN,Vector.ByCoordinates(0,1,0))
	PLANEY=Plane.ByOriginNormal(a.CEN,Vector.ByCoordinates(1,0,0))
	PLANEZ=Plane.ByOriginNormal(a.CEN,Vector.ByCoordinates(0,0,1))
	SURF=a.SOL.Intersect(PLANEZ)[0]
	LX=SURF.Intersect(PLANEX)[0]
	LY=SURF.Intersect(PLANEY)[0]
	LLX=LX.Translate(Vector.ByCoordinates(0,-LY.Length/2,-6))
	LLY=LY.Translate(Vector.ByCoordinates(-LX.Length/2,0,-6))
	if LLX.StartPoint.X<LLX.EndPoint.X:
		LLX2=LLX
	else:
		LLX2=LLX.Reverse()
	if LLY.StartPoint.Y<LLY.EndPoint.Y:
		LLY2=LLY
	else:
		LLY2=LLY.Reverse()

	return LLX2,LLY2


def createmodulesfromcurves(CUR,MOD):
	POL = PolyCurve.ByJoinedCurves(CUR)
	SUR= Surface.ByPatch(POL)
	ISO0=SUR.GetIsoline(0,0.5).Length
	ISO1=SUR.GetIsoline(1,0.5).Length
	L0=1/(ISO0/MOD)
	L1=1/(ISO1/MOD)
	X0=L0/2
	Y0=L1/2
	US=[]
	VS=[]
	
	while X0<1:
		US.append(X0)
		X0=X0+L0
	while Y0<1:
		VS.append(Y0)
		Y0=Y0+L1
		
	REC=[]
	for a in US:
		RET=[]
		for b in VS:
			POINT=SUR.PointAtParameter(a,b)
			COOR=CoordinateSystem.ByOrigin(POINT)
			RE=Rectangle.ByWidthLength(COOR,MOD,MOD)
			SUR1=Surface.ByPatch(RE)
			if Sphere.ByCenterPointRadius(SUR1.PointAtParameter(0.5,0.5),0.1).DoesIntersect(SUR):
				RET.append(RE)
		REC.append(RET)
	return REC

def createmodulesfromsurf(SUR,MOD):
	ISO0=SUR.GetIsoline(0,0.5).Length
	ISO1=SUR.GetIsoline(1,0.5).Length
	L0=1/(ISO0/MOD)
	L1=1/(ISO1/MOD)
	X0=L0/2
	Y0=L1/2
	US=[]
	VS=[]
	
	while X0<1:
		US.append(X0)
		X0=X0+L0
	while Y0<1:
		VS.append(Y0)
		Y0=Y0+L1
		
	REC=[]
	for a in US:
		RET=[]
		for b in VS:
			POINT=SUR.PointAtParameter(a,b)
			CUB=Cuboid.ByLengths(POINT,MOD,MOD,MOD)
			if Sphere.ByCenterPointRadius(POINT,0.2).DoesIntersect(SUR):
				RET.append(CUB)
		REC.append(RET)
	return REC

		

def LinestoWalls(LINES):
	SOLIDS = []
	if isinstance(LINES,list):
		for a in LINES:
			WALL=[]
			LineA=a.Offset(0.25)
			LineB=a.Offset(-0.25)
			WALL.append(LineA)
			WALL.append(LineB)
			SURF = Surface.ByLoft(WALL)	
			SOLIDS.append(SURF.Translate(Vector.ByCoordinates(0,0,6)).Thicken(12))
	else:
		WALL=[]
		LineA=LINES.Offset(0.25)
		LineB=LINES.Offset(-0.25)
		WALL.append(LineA)
		WALL.append(LineB)
		SURF = Surface.ByLoft(WALL)	
		SOLIDS.append(SURF.Translate(Vector.ByCoordinates(0,0,6)).Thicken(12))
	return SOLIDS

def flattenlst3(lst):
	FIN=[]
	if isinstance(lst,list):
		for a in lst:
			if isinstance(a,list):
				for b in a:
					if isinstance(b,list):
						for c in b:
							FIN.append(c)
					else:
						FIN.append(b)
			else:
				FIN.append(a)
	else:
		FIN.append(lst)
	return FIN

def ordlnsasp(CUR):
	POL = PolyCurve.ByJoinedCurves(CUR)
	SUR= Surface.ByPatch(POL)
	return SUR.PerimeterCurves()

class MODB:
	def _init_(self,IOF,SOL,VEC,CEN,ORI,WID,LEN,HEI,USE,CLEA,PRI,SCA,SCB,CSYS):
		self.IOF=IOF
		self.SOL=SOL
		self.VEC=VEC
		self.CEN=CEN
		self.ORI=ORI
		self.WID=WID
		self.LEN=LEN
		self.HEI=HEI
		self.USE=USE
		self.CLEA=CLEA
		self.PRI=PRI
		self.SCA=SCA
		self.SCB=SCB
		self.CSYS=CSYS

def linestocp(lst):
	LST=[]
	for a in lst:
		LST.append(a.PointAtParameter(0.5))
	return LST

def int1(lsa,lsb):
	FP=[]
	for a in lsa:
		for b in lsb:
			if a.DoesIntersect(b):
				FP.append(a)
	return FP
	
def intno(lsa,lsb):
	FP=[]
	for a in lsa:
		Y=0
		for b in lsb:
			if a.DoesIntersect(b):
				Y=1
		if Y==0:
			FP.append(a)
	return FP
	
def cuboid(ori,vect,W,L,H):
	CSys=CoordinateSystem.ByOriginVectors(ori,vect.Rotate(Vector.ByCoordinates(0,0,1),-90),vect,Vector.ByCoordinates(0,0,1))
	CBS=Cuboid.ByLengths(CSys,L,W,H)
	return CBS

def modulate(LLL,VEC,W,L,H,WW,HH,PRI,IOF,CABS):
	X=0
	if len(LLL)>1:
		for a,b,c in zip(LLL,range(len(LLL)),VEC):
			m1=MODB()
			m1.IOF=IOF+X
			m1.VEC=c
			m1.WID=W
			m1.LEN=L
			m1.HEI=H
			m1.ORI=a
			m1.CEN=a.Translate(c,(m1.WID/2)+WW).Translate(Vector.ByCoordinates(0,0,HH))
			m1.USE=0
			m1.SOL=cuboid(m1.CEN,m1.VEC,m1.WID,m1.LEN,m1.HEI)
			m1.SCA=len(LLL)-b-1
			m1.SCB=b
			m1.PRI=PRI
			m1.CSYS=CoordinateSystem.ByOriginVectors(m1.CEN,m1.VEC.Rotate(Vector.ByCoordinates(0,0,1),-90),m1.VEC,Vector.ByCoordinates(0,0,1))
			CABS.append(m1)
			X=X+1

def RECCEN(REC):
	SUR=Surface.ByPatch(REC) 
	CEN=SUR.PointAtParameter(0.5,0.5)
	return CEN

def LINESANDVECS(ORDC,CLRECS):
	ALLLINES=[]
	ALLVECS=[]
	for a in ORDC:
		LINT=[]
		VINT=[]
		for b in CLRECS:
			if b.DoesIntersect(a):
				INTER=b.Intersect(a)[0]
				if isinstance(INTER,Line):	
					PA=	INTER.PointAtParameter(0.5)
					PB = centroido(b)							
					VINT.append(Vector.ByTwoPoints(PA,PB))
					LINT.append(PA)
		ALLVECS.append(VINT)
		ALLLINES.append(LINT)
	return ALLLINES,ALLVECS	
		
def SPV(LINES,POINTS,VECS):
	POINTS2=[]
	VECS2=[]
	for a,b,c in zip(LINES,POINTS,VECS):
		TPAR=[]
		TPOINTS=[]
		TVECS=[]
		for d in b:
			TPAR.append(a.ParameterAtPoint(d))
		TPAR2=list(TPAR)
		TPAR2.sort()
		TPAR2.reverse()
		for e in TPAR2:
			TPOINTS.append(b[TPAR.index(e)])
			TVECS.append(c[TPAR.index(e)])
		POINTS2.append(TPOINTS)
		VECS2.append(TVECS)
	return POINTS2,VECS2

class FUR:
	def _init_(self,IOF,MIOF,CSYS,MODS,WID,LEN,HEI,VEC,ORI,CEN,SOL,CWID,CLEN,CHEI,CLECEN,CLEASOL,SIZE,FAMNAM,TYPE):
		self.IOF=IOF
		self.MIOF=MIOF
		self.CSYS=CSYS
		self.MODS=MODS
		self.WID=WID
		self.LEN=LEN
		self.HEI=HEI
		self.VEC=VEC
		self.ORI=ORI
		self.CEN=CEN
		self.SOL=SOL
		self.CWID=CWID
		self.CLEN=CLEN
		self.CHEI=CHEI
		self.CLECEN=CLECEN
		self.CLEASOL=CLEASOL
		self.SIZE=SIZE		
		self.FAMNAM=FAMNAM
		self.TYPE=TYPE
		
def cuboidT(ori,vect,W,L,H):
	CSys=CoordinateSystem.ByOriginVectors(ori,vect.Rotate(Vector.ByCoordinates(0,0,1),-90),vect,Vector.ByCoordinates(0,0,1))
	CBS=Cuboid.ByLengths(CSys,L,W,H)
	return CBS


def furalloc(IOF,MODS,W,L,H,CW,CL,CH,SIZE,FAMNAM,TYPE):
	fur1=FUR()
	fur1.IOF=IOF
	fur1.MODS=MODS
	fur1.WID=W
	fur1.LEN=L
	fur1.HEI=H
	fur1.CWID=CW
	fur1.CLEN=CL
	fur1.CHEI=CH
	fur1.SIZE=SIZE
	fur1.FAMNAM=FAMNAM
	fur1.TYPE=TYPE
	return fur1
	
def furnite(fur1,M):
	fur1.MIOF=M.IOF
	fur1.CSYS=M.CSYS
	fur1.VEC=M.VEC
	fur1.ORI=M.CEN.Translate(M.VEC.Reverse(),M.WID/2).Translate(Vector.ByCoordinates(0,0,-1),M.HEI/2).Translate(fur1.CSYS.XAxis,((fur1.MODS*2)/2)-1)
	fur1.CEN=fur1.ORI.Translate(M.VEC,fur1.WID/2).Translate(Vector.ByCoordinates(0,0,fur1.HEI/2))
	fur1.SOL=cuboidT(fur1.CEN,fur1.VEC,fur1.WID,fur1.LEN,fur1.HEI)
	fur1.CLEACEN=fur1.ORI.Translate(M.VEC,fur1.CWID/2).Translate(Vector.ByCoordinates(0,0,fur1.CHEI/2))
	fur1.CLEASOL=cuboidT(fur1.CLEACEN,fur1.VEC,fur1.CWID,fur1.CLEN,fur1.CHEI)
	return fur1
	
def cleandic(list3,DSCA):
	list2=[]
	for a in list3:
		list2.append(DSCA.get(str(a)))
	POST=[]
	PRE=[]
	PRE.append(0)
	for a in range(len(list2)):
		if a+1<len(list2):
			if list2[a+1]+1!=(list2[a]):
				POST.append(a)
				PRE.append(a+1)
			else:
				pass
	POST.append(len(list2)-1)
	VAL=[]
	for a,b in zip(PRE,POST):
		X=b-a
		for c in range(b-a+1):
			VAL.append(X)
			X=X-1
	return VAL	
	
def getangleofvector(VEC):
	VX=VEC.X
	VY=VEC.Y
	VEC2=Vector.ByCoordinates(VX,VY,0)
	angle=Vector.ByCoordinates(0,1,0).AngleWithVector(VEC2)
	if VX>0:
		pass
	else:
		angle=angle*-1
	return angle

def tclean(indices,MODULES,INT,DSCA):
	TINDICES=[]
	try:
		for a in indices:
			if MODULES[a].SOL.DoesIntersect(INT):
				TINDICES.append(a)
		
		for a in TINDICES:
			indices.remove(a)
			DSCA[str(a)]=0
					
		DVAL=cleandic(indices,DSCA)			
		for a,b in zip(indices,DVAL):
			DSCA[str(a)]=b
	except:
		pass

def BATHCONFIGS(LLENGTH):
	DIVISIONS=[]
	DIVNAM=[]
	DIVISIONS.append(0)
	if LLENGTH==4:
		DIVISIONS.append(2)
		DIVNAM.append("BATH")
		DIVNAM.append("BATH")
	elif LLENGTH==5 or LLENGTH==6:
		DIVISIONS.append(3)
		DIVNAM.append("BATH")
		DIVNAM.append("BATH")
	elif LLENGTH==7:
		DIVISIONS.append(4)
		DIVNAM.append("BATH")
		DIVNAM.append("BATH")
	elif LLENGTH==8:
		DIVISIONS.append(4)
		DIVNAM.append("BATH")
		DIVNAM.append("BATH")
	elif LLENGTH==9:
		DIVISIONS.append(3)
		DIVISIONS.append(3)
		DIVNAM.append("BATH")
		DIVNAM.append("BATH")
		DIVNAM.append("BATH")
	elif LLENGTH==10 or LLENGTH==11:
		DIVISIONS.append(4)
		DIVISIONS.append(3)
		DIVNAM.append("BATH")
		DIVNAM.append("BATH")
		DIVNAM.append("BATH")
	elif LLENGTH==12:
		DIVISIONS.append(4)
		DIVISIONS.append(4)
		DIVNAM.append("BATH")
		DIVNAM.append("BATH")
		DIVNAM.append("BATH")
	DIVISIONS.append(LLENGTH)	
	return DIVISIONS,DIVNAM

def MASTERCONFIGS(LLENGTH):
	DIVISIONS=[]
	DIVNAM=[]
	DIVISIONS.append(0)
	if LLENGTH>15:		
		DIVISIONS.append(LLENGTH-12)
		DIVISIONS.append(LLENGTH-8)
		DIVNAM.append("MASTERBATH")
		DIVNAM.append("MASTERWARD")
		DIVNAM.append("MASTERROOM")
	elif LLENGTH>11:
		DIVISIONS.append(LLENGTH-8)
		DIVNAM.append("MASTERBATH")
		DIVNAM.append("MASTERROOM")
	else:
		DIVNAM.append("MASTERROOM")
	DIVISIONS.append(LLENGTH)
	return DIVISIONS,DIVNAM

def ENTRANCECONFIGS(LLENGTH):
	DIVISIONS=[]
	DIVNAM=[]	
	DIVISIONS.append(0)
	if LLENGTH>11:
		if (LLENGTH-10)>4:
			DIVISIONS.append(LLENGTH-13)
			DIVISIONS.append(LLENGTH-10)
			DIVISIONS.append(LLENGTH-8)
			DIVNAM.append("WARD")
			DIVNAM.append("TOILET")
			DIVNAM.append("ENTRANCE")
			DIVNAM.append("STUDIO")
		else:
			DIVISIONS.append(LLENGTH-10)
			DIVISIONS.append(LLENGTH-8)
			DIVNAM.append("TOILET")
			DIVNAM.append("ENTRANCE")
			DIVNAM.append("STUDIO")
		
		MIDLINE=LLENGTH-9
	elif LLENGTH>8:
		if (LLENGTH-5)>4:
			DIVISIONS.append(LLENGTH-8)
			DIVISIONS.append(LLENGTH-5)
			DIVISIONS.append(LLENGTH-3)
			DIVNAM.append("WARD")
			DIVNAM.append("TOILET")
			DIVNAM.append("ENTRANCE")
			DIVNAM.append("WARD")
		else:
			DIVISIONS.append(LLENGTH-5)
			DIVISIONS.append(LLENGTH-3)
			DIVNAM.append("WARD")
			DIVNAM.append("ENTRANCE")
			DIVNAM.append("TOILET")
		MIDLINE=LLENGTH-4

	elif LLENGTH>5:
		DIVNAM.append("WARD")
		DIVNAM.append("ENTRANCE")
		DIVNAM.append("TOILET")
		DIVISIONS.append(LLENGTH-5)
		DIVISIONS.append(LLENGTH-3)
		MIDLINE=LLENGTH-4	
	else:
		DIVISIONS.append(2)
		DIVNAM.append("ENTRANCE")
		DIVNAM.append("TOILET")
		MIDLINE=1
	DIVISIONS.append(LLENGTH)
	return DIVISIONS,DIVNAM,MIDLINE

def divlnstosurf(DL):
	FSURF=[]
	for b in range(len(DL)):
		if	b<len(DL)-1:
			LLIST=[]
			LLIST.append(DL[b])
			LLIST.append(DL[b+1])
			S1=Surface.ByLoft(LLIST)
			FSURF.append(S1)
	return FSURF

def createdivlines(DIVS,LINE,VECTOR):
	LOCATIONPOINTS=[]
	for a in DIVS:
		P1=LINE.PointAtSegmentLength(a*2)
		LOCATIONPOINTS.append(P1)
	DL=[]
	for a in LOCATIONPOINTS:
		DL1=Line.ByStartPointEndPoint(a,a.Translate(VECTOR))
		DL.append(DL1)
	return DL

def MACROCONFIG(text,MACRO):
	RESULT=[]
	if text=="ENTRANCE":
		LINES=GETINDLXLYBYNAME(MACRO)
		LLENGTH=round(LINES[0].Length)/2
		CONFIGS=ENTRANCECONFIGS(LLENGTH)
		DIVISIONS=CONFIGS[0]
		NAMES=CONFIGS[1]
		MIDLINE=CONFIGS[2]	
		VECTOR= Vector.ByCoordinates(0,LINES[1].Length,0)
	
		DL=createdivlines(DIVISIONS,LINES[0],VECTOR)
		FSURFS=divlnstosurf(DL)
		
		MDL=Line.ByStartPointEndPoint(LINES[0].PointAtSegmentLength(MIDLINE*2),LINES[0].PointAtSegmentLength(MIDLINE*2).Translate(VECTOR))
		CL=LINES[0].Translate(Vector.ByCoordinates(0,LINES[1].Length+2,0))
		MACROS=[]
		X=20
		for a,b in zip(FSURFS,NAMES):
			MACROS.append(NewMacrobySurfandName(MACRO,a,b,X))
			X=X+1
		
		RESULT.append(MACROS)
		RESULT.append(FSURFS)
		RESULT.append(NAMES)
		RESULT.append(MDL)
		RESULT.append(CL)
		
	elif text=="BATH":
		LINES=GETINDLXLYBYNAME(MACRO)
		LLENGTH=round(LINES[1].Length)/2
		CONFIG=BATHCONFIGS(LLENGTH)
		DIVISIONS=CONFIG[0]
		NAMES=CONFIG[1]
		VECTOR= LINES[0].Direction		
		
		DL=createdivlines(DIVISIONS,LINES[1],VECTOR)
		FSURFS=divlnstosurf(DL)	

		MACROS=[]
		X=30
		for a,b in zip(FSURFS,NAMES):
			MACROS.append(NewMacrobySurfandName(MACRO,a,b,X))
			X=X+1
		
		RESULT.append(MACROS)			
		RESULT.append(FSURFS)
		RESULT.append(NAMES)
		
	elif text=="MASTERS":
		LINES=GETINDLXLYBYNAME(MACRO)
		if LINES[0].Length>LINES[1].Length:
			LLENGTH=round(LINES[0].Length)/2
		else:
			LLENGTH=round(LINES[1].Length)/2
		CONFIG=MASTERCONFIGS(LLENGTH)
		DIVISIONS=CONFIG[0]
		NAMES=CONFIG[1]
		VECTOR= LINES[1].Direction		
	
		DL=createdivlines(DIVISIONS,LINES[0],VECTOR)
		FSURFS=divlnstosurf(DL)	

		MACROS=[]
		X=40
		for a,b in zip(FSURFS,NAMES):
			MACROS.append(NewMacrobySurfandName(MACRO,a,b,X))
			X=X+1
	
		if len(DL)>2:
			MIDLINE=Line.ByStartPointEndPoint(DL[1].PointAtSegmentLength(DL[1].Length-4).Translate(LINES[0].Direction.Reverse(),1),DL[-2].PointAtSegmentLength(DL[1].Length-4).Translate(LINES[0].Direction,1))
			
		RESULT.append(MACROS)	
		RESULT.append(FSURFS)
		RESULT.append(NAMES)
		RESULT.append(MIDLINE)
		RESULT.append(DL)

	return RESULT

def MACCONFIG(text,MACRO,INTS):
	RESULT=[]
	if text=="ENTRANCE":
		LINES=GETINDLXLYBYNAME(MACRO)
		INTER=[]
		for a in INTS:
			OBJ=a.Intersect(MACRO.SOL)[0]
			if str(OBJ.GetType())=="Autodesk.DesignScript.Geometry.Line":
				INTER.append(OBJ)
			else:
				INTER.append(TOL(OBJ.GetIsoline(0,0.5)))
		if len(INTER)>0:
			LENGTH=[]
			for a in INTER:
				LENGTH.append(a.Length)
			MAX=max(LENGTH)
			IND=LENGTH.index(MAX)
			VEC=INTER[IND].Direction
		else:
			VEC=INTER[0].Direction
		if VEC.IsParallel(LINES[0].Direction):
			SELINE=LINES[0]
			ODLINE=LINES[1]
		else:
			SELINE=LINES[1]
			ODLINE=LINES[0]
		LLENGTH=round(SELINE.Length)/2
		CONFIG=ENTRANCECONFIGS(LLENGTH)
		DIVISIONS=CONFIG[0]
		NAMES=CONFIG[1]
		MIDLINE=CONFIG[2]
		VECTOR= ODLINE.Direction				
		DL=createdivlines(DIVISIONS,SELINE,VECTOR)
		FSURFS=divlnstosurf(DL)	
		MDL=Line.ByStartPointEndPoint(ODLINE.PointAtSegmentLength(MIDLINE*2),ODLINE.PointAtSegmentLength(MIDLINE*2).Translate(VECTOR))
		CL=ODLINE.Translate(Vector.ByCoordinates(0,SELINE.Length+2,0))
		MACROS=[]
		X=30
		for a,b in zip(FSURFS,NAMES):
			MACROS.append(NewMacrobySurfandName(MACRO,a,b,X))
			X=X+1
		RESULT.append(MACROS)			
		RESULT.append(FSURFS)
		RESULT.append(NAMES)
		RESULT.append(MDL)
		RESULT.append(CL)	

	elif text=="BATH":
		LINES=GETINDLXLYBYNAME(MACRO)
		INTER=[]
		for a in INTS:
			OBJ=a.Intersect(MACRO.SOL)[0]
			if str(OBJ.GetType())=="Autodesk.DesignScript.Geometry.Line":
				INTER.append(OBJ)
			else:
				INTER.append(TOL(OBJ.GetIsoline(0,0.5)))
		if len(INTER)>0:
			LENGTH=[]
			for a in INTER:
				LENGTH.append(a.Length)
			MAX=max(LENGTH)
			IND=LENGTH.index(MAX)
			VEC=INTER[IND].Direction
		else:
			VEC=INTER[0].Direction
		if VEC.IsParallel(LINES[0].Direction):
			SELINE=LINES[0]
			ODLINE=LINES[1]
		else:
			SELINE=LINES[1]
			ODLINE=LINES[0]
		LLENGTH=round(SELINE.Length)/2
		CONFIG=BATHCONFIGS(LLENGTH)
		DIVISIONS=CONFIG[0]
		NAMES=CONFIG[1]
		VECTOR= ODLINE.Direction				
		DL=createdivlines(DIVISIONS,SELINE,VECTOR)
		FSURFS=divlnstosurf(DL)	
		MACROS=[]
		X=30
		for a,b in zip(FSURFS,NAMES):
			MACROS.append(NewMacrobySurfandName(MACRO,a,b,X))
			X=X+1
		RESULT.append(MACROS)			
		RESULT.append(FSURFS)
		RESULT.append(NAMES)

	elif text=="MASTER":
		LINES=GETINDLXLYBYNAME(MACRO)
		INTER=[]
		for a in INTS:
			OBJ=a.Intersect(MACRO.SOL)[0]
			if str(OBJ.GetType())=="Autodesk.DesignScript.Geometry.Line":
				INTER.append(OBJ)
			else:
				INTER.append(TOL(OBJ.GetIsoline(0,0.5)))
		if len(INTER)>0:
			LENGTH=[]
			for a in INTER:
				LENGTH.append(a.Length)
			MAX=max(LENGTH)
			IND=LENGTH.index(MAX)
			VEC=INTER[IND].Direction
		else:
			VEC=INTER[0].Direction
		if VEC.IsParallel(LINES[0].Direction):
			SELINE=LINES[0]
			ODLINE=LINES[1]
		else:
			SELINE=LINES[1]
			ODLINE=LINES[0]
		RESULT.append(SELINE)
		RESULT.append(ODLINE)
	return RESULT

def notclashingany(TEST,lst):
	KO=0
	if len(lst)>0:
		for a in lst:
			TT=clashlist(TEST,a)
			if TT>0:
				KO=TT
	return KO	

def clashlist(TEST,lst):
	KO=0
	if len(lst)>0:
		for a in lst:
			try:
				if TEST.CLEASOL.DoesIntersect(a.CLEASOL):
					KO=1
			except:
				if TEST.CLEASOL.DoesIntersect(a):
					KO=1
	return KO	
def randomselection2s(lst,per):
	num=len(lst)/2
	used=round(num*per,0)
	result=[]
	for a in range(int(used)):
		A=0
		while A==0:
			V=random.choice(lst)
			if str(int(V)+1) in lst:
				result.append(V)
				lst.remove(V)
				lst.remove(str(int(V)+1))
				A=1 
	return result


def LAYINGOUT(MODULES,SOLWIN,SOLDOR,ADDITIONAL,PLACE):
	indices=[]
	DSCA=dict()
	PDOORS=[]
	PWIND=[]
	ALLPOSIBLES=[]
	#GET INDICES
	for a in MODULES:
		indices.append(a.IOF)
		DSCA.Add(str(a.IOF),a.SCA)
	allSols=[]
	#GET DOOR AND WINDOW LOCATIONS
	UA=0
	for a in MODULES:
		if Sphere.ByCenterPointRadius(a.ORI,0.5).DoesIntersect(Cylinder.ByPointsRadius(ADDITIONAL.StartPoint,ADDITIONAL.EndPoint,1)):
			PDOORS.append(a.IOF)
			UA=1
		else:
			for b in SOLDOR:
				if Sphere.ByCenterPointRadius(a.ORI,0.5).DoesIntersect(b):
					PDOORS.append(a.IOF)

		for c in SOLWIN:
			if Sphere.ByCenterPointRadius(a.ORI,0.5).DoesIntersect(c):
				PWIND.append(a.IOF)
				
	PDOORS=list(set(PDOORS))
	PWIND=list(set(PWIND))

	#POSIBLE LOCATIONS OF DOORS AND WINDOWS
	PLOCDOOR=[]
	PDORS=[]

	for a in indices:
		if DSCA[str(a)]>=1 and a in PDOORS:
			PLOCDOOR.append(str(MODULES[a].IOF))

	if UA==0:
		if len(PLOCDOOR)>1:
			SELDOR=[]		
			SELDOR.append(PLOCDOOR[0])
			for a in SELDOR:
				DOOR=furalloc(0,2,0.25,2.88,10,3.7,3.7,12,4,"Door - Single Panel",'F - W - 31.5" x 120"')
				DOOR1=furnite(DOOR,MODULES[int(a)])
				PDORS.append(DOOR1)
			tclean(indices,MODULES,PDORS[0].CLEASOL,DSCA)

	elif UA>0:
		if len(PLOCDOOR)>1:
			SELDOR=[]		
			SELDOR.append(PLOCDOOR[0])
			if len(PLOCDOOR)>2:
				SELDOR.append(PLOCDOOR[2])
			for a in SELDOR:
				DOOR=furalloc(0,2,0.25,2.88,10,3.7,3.7,12,4,"Door - Single Panel",'F - W - 31.5" x 120"')
				DOOR1=furnite(DOOR,MODULES[int(a)])
				PDORS.append(DOOR1)
			for a in PDORS:
				tclean(indices,MODULES,a.CLEASOL,DSCA)
	
	#Create WINDOW
	PLOCWIN=[]
	PWIN=[]
	for a in indices:
		if DSCA[str(a)]>=1 and a in PWIND:
			PLOCWIN.append(str(MODULES[a].IOF))

	if len(PLOCWIN)>1:
		SELWIN=randomselection2s(PLOCWIN,0.50)
		
		for a in SELWIN:
			WIN=furalloc(1,2,0.25,3.8,10,0.10,3.7,12,4,"Window - CIEL Vertical_SB_Test",'F - G - 45" x 120"')
			WIN1=furnite(WIN,MODULES[int(a)])
			PWIN.append(WIN1)
		for a in PWIN:
			tclean(indices,MODULES,a.CLEASOL,DSCA)
			
	
	if len(PDORS)==0:
		DOOR=furalloc(0,2,0.25,2.88,10,3.7,3.7,12,4,"Door - Single Panel",'F - W - 31.5" x 120"')
		DOOR1=furnite(DOOR,MODULES[int(0)])
		PDORS.append(DOOR1)
		tclean(indices,MODULES,PDORS[0].CLEASOL,DSCA)
	if len(PWIN)==0:
		WIN=furalloc(1,2,0.25,3.8,10,0.10,3.7,12,4,"Window - CIEL Vertical_SB_Test",'F - G - 45" x 120"')
		WIN1=furnite(WIN,MODULES[int(0)])
		PWIN.append(WIN1)
		tclean(indices,MODULES,PWIN[0].CLEASOL,DSCA)
	FURNITE=[]
	if PLACE=="BATH1" or  PLACE=="BATH2":			
		TL=[PDORS,PWIN,FURNITE]
		FURPLA(indices,MODULES,8,1,4,2,5,3.70,10,5.8,3.7,10,6,"Shower","Type A",TL,PDORS[0],FURNITE,DSCA)
		TL=[PDORS,PWIN,FURNITE]
		FURPLA(indices,MODULES,5,2,3,3,1.97,5.25,2.35,3.8,5.8,4,6,"Vanity_3_Drawers","1450 mm",TL,PDORS[0],FURNITE,DSCA)
		TL=[PDORS,PWIN,FURNITE]
		FURPLA(indices,MODULES,2,1,2,2,1.83,1.33,1.48,3.8,3.8,4,4,"Sanitary_Toilets_Roca_ISLAND-Wall-hung-WC1","00 - Glossy white",TL,PDORS[0],FURNITE,DSCA)

	elif PLACE=="MASTERBATH":			
		TL=[PDORS,PWIN,FURNITE]
		FURPLA(indices,MODULES,14,1,5,2,6,3.5,3,6,3.7,3,6,"Simple_Freestanding_Tub2","6' Tub",TL,PDORS[0],FURNITE,DSCA)
		TL=[PDORS,PWIN,FURNITE]
		FURPLA(indices,MODULES,8,1,4,2,5,3.70,10,5.8,3.7,10,6,"Shower","Type A",TL,PDORS[0],FURNITE,DSCA)
		TL=[PDORS,PWIN,FURNITE]
		FURPLA(indices,MODULES,5,2,3,3,1.97,5.25,2.35,3.8,5.8,4,6,"Vanity_3_Drawers","1450 mm",TL,PDORS[0],FURNITE,DSCA)
		TL=[PDORS,PWIN,FURNITE]
		FURPLA(indices,MODULES,2,1,2,2,1.83,1.33,1.48,3.8,3.8,4,4,"Sanitary_Toilets_Roca_ISLAND-Wall-hung-WC1","00 - Glossy white",TL,PDORS[0],FURNITE,DSCA)

	elif PLACE=="ROOMS" or PLACE=="MASTERROOM":			
		TL=[PDORS,PWIN,FURNITE]
		FURPLA(indices,MODULES,4,3,1,4,6.4,3.4,3,8,7.8,3,16,"00_Bedroom","2000x2000 KING",TL,PDORS[0],FURNITE,DSCA)
		TL=[PDORS,PWIN,FURNITE]
		FURPLA(indices,MODULES,4,3,2,4,2.32,7.5,10,3.5,7.8,10,8,"00_Wardrobe","2400x700",TL,PDORS[0],FURNITE,DSCA)
		TL=[PDORS,PWIN,FURNITE]
		FURPLA(indices,MODULES,2,1,3,1,1.97,3.8,2.35,3.8,3.8,4,6,"00_WorkTable","Default",TL,PDORS[0],FURNITE,DSCA)

	elif PLACE=="MASTERWARD":			
		TL=[PDORS,FURNITE]
		FURPLA(indices,MODULES,4,3,2,4,2.32,7.5,10,2.32,7.9,10,8,"00_Wardrobe","2400x700",TL,PDORS[0],FURNITE,DSCA)
		TL=[PDORS,FURNITE]
		FURPLA(indices,MODULES,4,3,2,4,2.32,7.5,10,2.32,7.9,10,8,"00_Wardrobe","2400x700",TL,PDORS[0],FURNITE,DSCA)
		TL=[PDORS,FURNITE]
		FURPLA(indices,MODULES,4,3,2,4,2.32,7.5,10,2.32,7.9,10,8,"00_Wardrobe","2400x700",TL,PDORS[0],FURNITE,DSCA)
		TL=[PDORS,FURNITE]
		FURPLA(indices,MODULES,4,3,2,4,2.32,7.5,10,2.32,7.9,10,8,"00_Wardrobe","2400x700",TL,PDORS[0],FURNITE,DSCA)

	if PLACE=="TOILET" :			
		TL=[PDORS,PWIN,FURNITE]
		FURPLA(indices,MODULES,5,2,3,3,1.97,5.25,2.35,3.8,5.8,4,6,"Vanity_3_Drawers","1450 mm",TL,PDORS[0],FURNITE,DSCA)
		TL=[PDORS,PWIN,FURNITE]
		FURPLA(indices,MODULES,2,1,2,2,1.83,1.33,1.48,3.8,3.8,4,4,"Sanitary_Toilets_Roca_ISLAND-Wall-hung-WC1","00 - Glossy white",TL,PDORS[0],FURNITE,DSCA)

	CLEAR=[]
	# Place your code below this line
	for a in PDORS:
		allSols.append(a.SOL)
		CLEAR.append(a.CLEASOL)

	for a in PWIN:
		allSols.append(a.SOL)
		CLEAR.append(a.CLEASOL)

	FUR2=[]
	for a in FURNITE:
		FUR2.append(a.SOL)

	USED1=[]
	for a in indices:
		USED1.append(str(a)+"_"+str(DSCA.get(str(a))))
	
	return allSols,USED1,CLEAR,FUR2,PDORS,PWIN,FURNITE,PDOORS,PWIND

def LAYINGOUTB(MODULES,SOLWIN,SOLDOR,PLACE):
	indices=[]
	DSCA=dict()
	PDOORS=[]
	PWIND=[]
	ALLPOSIBLES=[]
	#GET INDICES
	for a in MODULES:
		indices.append(a.IOF)
		DSCA.Add(str(a.IOF),a.SCA)
	allSols=[]
	#GET DOOR AND WINDOW LOCATIONS
	for a in MODULES:
		for b in SOLDOR:
			if Sphere.ByCenterPointRadius(a.ORI,0.5).DoesIntersect(b):
				PDOORS.append(a.IOF)

		for c in SOLWIN:
			if Sphere.ByCenterPointRadius(a.ORI,0.5).DoesIntersect(c):
				PWIND.append(a.IOF)
				
	PDOORS=list(set(PDOORS))
	PWIND=list(set(PWIND))

	#POSIBLE LOCATIONS OF DOORS AND WINDOWS
	PLOCDOOR=[]
	PDORS=[]

	for a in indices:
		if DSCA[str(a)]>=1 and a in PDOORS:
			PLOCDOOR.append(str(MODULES[a].IOF))

	if len(PLOCDOOR)>1:
		SELDOR=[]		
		SELDOR.append(PLOCDOOR[0])
		for a in SELDOR:
			DOOR=furalloc(0,2,0.25,2.88,10,3.7,3.7,12,4,"Door - Single Panel",'F - W - 31.5" x 120"')
			DOOR1=furnite(DOOR,MODULES[int(a)])
			PDORS.append(DOOR1)
		tclean(indices,MODULES,PDORS[0].CLEASOL,DSCA)

	#Create WINDOW
	PLOCWIN=[]
	PWIN=[]
	for a in indices:
		if DSCA[str(a)]>=1 and a in PWIND:
			PLOCWIN.append(str(MODULES[a].IOF))

	if len(PLOCWIN)>1:
		SELWIN=randomselection2s(PLOCWIN,0.50)
		
		for a in SELWIN:
			WIN=furalloc(1,2,0.25,3.8,10,0.10,3.7,12,4,"Window - CIEL Vertical_SB_Test",'F - G - 45" x 120"')
			WIN1=furnite(WIN,MODULES[int(a)])
			PWIN.append(WIN1)
		for a in PWIN:
			tclean(indices,MODULES,a.CLEASOL,DSCA)
			
	
	if len(PDORS)==0:
		DOOR=furalloc(0,2,0.25,2.88,10,3.7,3.7,12,4,"Door - Single Panel",'F - W - 31.5" x 120"')
		DOOR1=furnite(DOOR,MODULES[int(0)])
		PDORS.append(DOOR1)
		tclean(indices,MODULES,PDORS[0].CLEASOL,DSCA)
	if len(PWIN)==0:
		WIN=furalloc(1,2,0.25,3.8,10,0.10,3.7,12,4,"Window - CIEL Vertical_SB_Test",'F - G - 45" x 120"')
		WIN1=furnite(WIN,MODULES[int(0)])
		PWIN.append(WIN1)
		tclean(indices,MODULES,PWIN[0].CLEASOL,DSCA)
	FURNITE=[]
	if PLACE=="BATH":			
		TL=[PDORS,PWIN,FURNITE]
		FURPLA(indices,MODULES,8,1,4,2,5,3.70,10,5.8,3.7,10,6,"Shower","Type A",TL,PDORS[0],FURNITE,DSCA)
		TL=[PDORS,PWIN,FURNITE]
		FURPLA(indices,MODULES,5,2,3,3,1.97,5.25,2.35,3.8,5.8,4,6,"Vanity_3_Drawers","1450 mm",TL,PDORS[0],FURNITE,DSCA)
		TL=[PDORS,PWIN,FURNITE]
		FURPLA(indices,MODULES,2,1,2,2,1.83,1.33,1.48,3.8,3.8,4,4,"Sanitary_Toilets_Roca_ISLAND-Wall-hung-WC1","00 - Glossy white",TL,PDORS[0],FURNITE,DSCA)

	elif PLACE=="MASTERBATH":			
		TL=[PDORS,PWIN,FURNITE]
		FURPLA(indices,MODULES,14,1,5,2,6,3.5,3,6,3.7,3,6,"Simple_Freestanding_Tub2","6' Tub",TL,PDORS[0],FURNITE,DSCA)
		TL=[PDORS,PWIN,FURNITE]
		FURPLA(indices,MODULES,8,1,4,2,5,3.70,10,5.8,3.7,10,6,"Shower","Type A",TL,PDORS[0],FURNITE,DSCA)
		TL=[PDORS,PWIN,FURNITE]
		FURPLA(indices,MODULES,5,2,3,3,1.97,5.25,2.35,3.8,5.8,4,6,"Vanity_3_Drawers","1450 mm",TL,PDORS[0],FURNITE,DSCA)
		TL=[PDORS,PWIN,FURNITE]
		FURPLA(indices,MODULES,2,1,2,2,1.83,1.33,1.48,3.8,3.8,4,4,"Sanitary_Toilets_Roca_ISLAND-Wall-hung-WC1","00 - Glossy white",TL,PDORS[0],FURNITE,DSCA)

	elif PLACE=="ROOMS" or PLACE=="MASTERROOM":			
		TL=[PDORS,PWIN,FURNITE]
		FURPLA(indices,MODULES,4,3,1,4,6.4,3.4,3,8,7.8,3,16,"00_Bedroom","2000x2000 KING",TL,PDORS[0],FURNITE,DSCA)
		TL=[PDORS,PWIN,FURNITE]
		FURPLA(indices,MODULES,4,3,2,4,2.32,7.5,10,3.5,7.8,10,8,"00_Wardrobe","2400x700",TL,PDORS[0],FURNITE,DSCA)
		TL=[PDORS,PWIN,FURNITE]
		FURPLA(indices,MODULES,2,1,3,1,1.97,3.8,2.35,3.8,3.8,4,6,"00_WorkTable","Default",TL,PDORS[0],FURNITE,DSCA)

	elif PLACE=="MASTERWARD":			
		TL=[PDORS,FURNITE]
		FURPLA(indices,MODULES,4,3,2,4,2.32,7.5,10,2.32,7.9,10,8,"00_Wardrobe","2400x700",TL,PDORS[0],FURNITE,DSCA)
		TL=[PDORS,FURNITE]
		FURPLA(indices,MODULES,4,3,2,4,2.32,7.5,10,2.32,7.9,10,8,"00_Wardrobe","2400x700",TL,PDORS[0],FURNITE,DSCA)
		TL=[PDORS,FURNITE]
		FURPLA(indices,MODULES,4,3,2,4,2.32,7.5,10,2.32,7.9,10,8,"00_Wardrobe","2400x700",TL,PDORS[0],FURNITE,DSCA)
		TL=[PDORS,FURNITE]
		FURPLA(indices,MODULES,4,3,2,4,2.32,7.5,10,2.32,7.9,10,8,"00_Wardrobe","2400x700",TL,PDORS[0],FURNITE,DSCA)

	if PLACE=="TOILET" :			
		TL=[PDORS,PWIN,FURNITE]
		FURPLA(indices,MODULES,5,2,3,3,1.97,5.25,2.35,3.8,5.8,4,6,"Vanity_3_Drawers","1450 mm",TL,PDORS[0],FURNITE,DSCA)
		TL=[PDORS,PWIN,FURNITE]
		FURPLA(indices,MODULES,2,1,2,2,1.83,1.33,1.48,3.8,3.8,4,4,"Sanitary_Toilets_Roca_ISLAND-Wall-hung-WC1","00 - Glossy white",TL,PDORS[0],FURNITE,DSCA)

	CLEAR=[]
	# Place your code below this line
	for a in PDORS:
		allSols.append(a.SOL)
		CLEAR.append(a.CLEASOL)

	for a in PWIN:
		allSols.append(a.SOL)
		CLEAR.append(a.CLEASOL)

	FUR2=[]
	for a in FURNITE:
		FUR2.append(a.SOL)

	USED1=[]
	for a in indices:
		USED1.append(str(a)+"_"+str(DSCA.get(str(a))))
	
	return allSols,USED1,CLEAR,FUR2,PDORS,PWIN,FURNITE,PDOORS,PWIND

def LAYINGOUTC(MODULES,SOLWIN,SOLDOR,PLACE):
	indices=[]
	DSCA=dict()
	PDOORS=[]
	PWIND=[]
	ALLPOSIBLES=[]
	#GET INDICES
	for a in MODULES:
		indices.append(a.IOF)
		DSCA.Add(str(a.IOF),a.SCA)
	allSols=[]
	#GET DOOR AND WINDOW LOCATIONS
	for a in MODULES:
		for b in SOLDOR:
			if Sphere.ByCenterPointRadius(a.ORI,0.5).DoesIntersect(b):
				PDOORS.append(a.IOF)

		for c in SOLWIN:
			if Sphere.ByCenterPointRadius(a.ORI,0.5).DoesIntersect(c):
				PWIND.append(a.IOF)
				
	PDOORS=list(set(PDOORS))
	PWIND=list(set(PWIND))

	#POSIBLE LOCATIONS OF DOORS AND WINDOWS
	PLOCDOOR=[]
	PDORS=[]

	for a in indices:
		if DSCA[str(a)]>=1 and a in PDOORS:
			PLOCDOOR.append(str(MODULES[a].IOF))

	if len(PLOCDOOR)>1:
		SELDOR=[]		
		SELDOR.append(PLOCDOOR[0])
		for a in SELDOR:
			DOOR=furalloc(0,2,0.25,2.88,10,3.7,3.7,12,4,"Door - Single Panel",'F - W - 31.5" x 120"')
			DOOR1=furnite(DOOR,MODULES[int(a)])
			PDORS.append(DOOR1)
		tclean(indices,MODULES,PDORS[0].CLEASOL,DSCA)

	#Create WINDOW
	PLOCWIN=[]
	PWIN=[]
	for a in indices:
		if DSCA[str(a)]>=1 and a in PWIND:
			PLOCWIN.append(str(MODULES[a].IOF))

	if len(PLOCWIN)>1:
		SELWIN=randomselection2s(PLOCWIN,0.50)
		
		for a in SELWIN:
			WIN=furalloc(1,2,0.25,3.8,10,0.10,3.7,12,4,"Window - CIEL Vertical_SB_Test",'F - G - 45" x 120"')
			WIN1=furnite(WIN,MODULES[int(a)])
			PWIN.append(WIN1)
		for a in PWIN:
			tclean(indices,MODULES,a.CLEASOL,DSCA)
			
	
	if len(PDORS)==0:
		DOOR=furalloc(0,2,0.25,2.88,10,3.7,3.7,12,4,"Door - Single Panel",'F - W - 31.5" x 120"')
		DOOR1=furnite(DOOR,MODULES[int(0)])
		PDORS.append(DOOR1)
		tclean(indices,MODULES,PDORS[0].CLEASOL,DSCA)
	if len(PWIN)==0:
		WIN=furalloc(1,2,0.25,3.8,10,0.10,3.7,12,4,"Window - CIEL Vertical_SB_Test",'F - G - 45" x 120"')
		WIN1=furnite(WIN,MODULES[int(0)])
		PWIN.append(WIN1)
		tclean(indices,MODULES,PWIN[0].CLEASOL,DSCA)

	CLEAR=[]
	# Place your code below this line
	for a in PDORS:
		allSols.append(a.SOL)
		CLEAR.append(a.CLEASOL)

	for a in PWIN:
		allSols.append(a.SOL)
		CLEAR.append(a.CLEASOL)

		USED1=[]
	for a in indices:
		USED1.append(str(a)+"_"+str(DSCA.get(str(a))))
	
	return allSols,USED1,CLEAR,PDORS,PWIN,PDOORS,PWIND

def LAYINGOUTD(MODULES,PLACE,DW):
	indices=[]
	DSCA=dict()
	ODOR=DW[0]
	SDOR=DW[1]
	OWIN=DW[2]
	SWIN=DW[3]
	for a in MODULES:
		indices.append(a.IOF)
		DSCA.Add(str(a.IOF),a.SCA)
	allSols=[]

	if len(SDOR)>0:
		for a in SDOR:
			tclean(indices,MODULES,a,DSCA)
	
	if len(SWIN)>0:
		for a in SWIN:
			tclean(indices,MODULES,a,DSCA)

	PDORS=[]
	if len(SDOR)>0:
		for a in ODOR:
			PDORS.append(Sphere.ByCenterPointRadius(a,3))
		ORISET=ODOR[0]
	elif len(SWIN)>0:
		ORISET=OWIN[0]
	else:
		ORISET=MODULES[0]
	PWIN=[]
	if len(SWIN)>0:
		for a in SWIN:
			PWIN.append(a)

	FURNITE=[]
	if PLACE=="BATH":			
		TL=[PDORS,PWIN,FURNITE]
		FURPLA(indices,MODULES,8,1,4,2,5,3.70,10,5.8,3.7,10,6,"Shower","Type A",TL,ORISET,FURNITE,DSCA)
		TL=[PDORS,PWIN,FURNITE]
		FURPLA(indices,MODULES,5,2,3,3,1.97,5.25,2.35,3.8,5.8,4,6,"Vanity_3_Drawers","1450 mm",TL,ORISET,FURNITE,DSCA)
		TL=[PDORS,PWIN,FURNITE]
		FURPLA(indices,MODULES,2,1,2,2,1.83,1.33,1.48,3.8,3.8,4,4,"Sanitary_Toilets_Roca_ISLAND-Wall-hung-WC1","00 - Glossy white",TL,ORISET,FURNITE,DSCA)

	elif PLACE=="MBATH":			
		TL=[PDORS,PWIN,FURNITE]
		FURPLA(indices,MODULES,14,1,5,2,6,3.5,3,6,3.7,3,6,"Simple_Freestanding_Tub2","6' Tub",TL,ORISET,FURNITE,DSCA)
		TL=[PDORS,PWIN,FURNITE]
		FURPLA(indices,MODULES,8,1,4,2,5,3.70,10,5.8,3.7,10,6,"Shower","Type A",TL,ORISET,FURNITE,DSCA)
		TL=[PDORS,PWIN,FURNITE]
		FURPLA(indices,MODULES,5,2,3,3,1.97,5.25,2.35,3.8,5.8,4,6,"Vanity_3_Drawers","1450 mm",TL,ORISET,FURNITE,DSCA)
		TL=[PDORS,PWIN,FURNITE]
		FURPLA(indices,MODULES,2,1,2,2,1.83,1.33,1.48,3.8,3.8,4,4,"Sanitary_Toilets_Roca_ISLAND-Wall-hung-WC1","00 - Glossy white",TL,ORISET,FURNITE,DSCA)

	elif PLACE=="ROOMS" or PLACE=="MASTER":			
		TL=[PDORS,PWIN,FURNITE]
		FURPLA(indices,MODULES,4,3,1,4,6.4,3.4,3,8,7.8,3,16,"00_Bedroom","2000x2000 KING",TL,ORISET,FURNITE,DSCA)
		TL=[PDORS,PWIN,FURNITE]
		FURPLA(indices,MODULES,4,3,2,4,2.32,7.5,10,3.5,7.8,10,8,"00_Wardrobe","2400x700",TL,ORISET,FURNITE,DSCA)
		TL=[PDORS,PWIN,FURNITE]
		FURPLA(indices,MODULES,2,1,3,1,1.97,3.8,2.35,3.8,3.8,4,6,"00_WorkTable","Default",TL,ORISET,FURNITE,DSCA)

	elif PLACE=="MASTERWARD" or PLACE=="WARD":			
		TL=[PDORS,FURNITE]
		FURPLA(indices,MODULES,4,3,2,4,2.32,7.5,10,2.32,7.9,10,8,"00_Wardrobe","2400x700",TL,ORISET,FURNITE,DSCA)
		TL=[PDORS,FURNITE]
		FURPLA(indices,MODULES,4,3,2,4,2.32,7.5,10,2.32,7.9,10,8,"00_Wardrobe","2400x700",TL,ORISET,FURNITE,DSCA)
		TL=[PDORS,FURNITE]
		FURPLA(indices,MODULES,4,3,2,4,2.32,7.5,10,2.32,7.9,10,8,"00_Wardrobe","2400x700",TL,ORISET,FURNITE,DSCA)
		TL=[PDORS,FURNITE]
		FURPLA(indices,MODULES,4,3,2,4,2.32,7.5,10,2.32,7.9,10,8,"00_Wardrobe","2400x700",TL,ORISET,FURNITE,DSCA)
	
	#elif PLACE=="KITCHEN":
	#	TL=[PDORS,PWIN,FURNITE]
	#	FURPLA(indices,MODULES,2,1,2,2,1.83,1.33,1.48,3.8,3.8,4,4,"Sanitary_Toilets_Roca_ISLAND-Wall-hung-WC1","00 - Glossy white",TL,ORISET,FURNITE,DSCA)			
	#	TL=[PDORS,PWIN,FURNITE]
	#	FURPLA(indices,MODULES,4,3,1,4,6.4,3.4,3,8,7.8,3,16,"00_Bedroom","2000x2000 KING",TL,ORISET,FURNITE,DSCA)
	#	TL=[PDORS,PWIN,FURNITE]
	#	FURPLA(indices,MODULES,4,3,2,4,2.32,7.5,10,3.5,7.8,10,8,"00_Wardrobe","2400x700",TL,ORISET,FURNITE,DSCA)
	#	TL=[PDORS,PWIN,FURNITE]
	#	FURPLA(indices,MODULES,2,1,3,1,1.97,3.8,2.35,3.8,3.8,4,6,"00_WorkTable","Default",TL,ORISET,FURNITE,DSCA)

	elif PLACE=="TOILET" :			
		TL=[PDORS,PWIN,FURNITE]
		FURPLA(indices,MODULES,5,2,3,3,1.97,5.25,2.35,3.8,5.8,4,6,"Vanity_3_Drawers","1450 mm",TL,ORISET,FURNITE,DSCA)
		TL=[PDORS,PWIN,FURNITE]
		FURPLA(indices,MODULES,2,1,2,2,1.83,1.33,1.48,3.8,3.8,4,4,"Sanitary_Toilets_Roca_ISLAND-Wall-hung-WC1","00 - Glossy white",TL,ORISET,FURNITE,DSCA)

	USED1=[]
	for a in indices:
		USED1.append(str(a)+"_"+str(DSCA.get(str(a))))
	
	return FURNITE

def CENTLAYOUT(MODULES,PLACE,DW):
	indices=[]
	DSCA=dict()
	ODOR=DW[0]
	SDOR=DW[1]
	OWIN=DW[2]
	SWIN=DW[3]
	for a in MODULES:
		indices.append(a.IOF)
		DSCA.Add(str(a.IOF),a.SCA)
	allSols=[]

	if len(SDOR)>0:
		for a in SDOR:
			tclean(indices,MODULES,a,DSCA)
	
	if len(SWIN)>0:
		for a in SWIN:
			tclean(indices,MODULES,a,DSCA)
	PDORS=[]
	if len(SDOR)>0:
		for a in ODOR:
			PDORS.append(Sphere.ByCenterPointRadius(a,3))
		ORISET=ODOR[0]
	elif len(SWIN)>0:
		ORISET=OWIN[0]
	else:
		ORISET=MODULE[0]
	PWIN=[]
	if len(SWIN)>0:
		for a in SWIN:
			PWIN.append(a)

	FURNITE=[]
	if PLACE=="BATH":			
		TL=[PDORS,PWIN,FURNITE]
		FURPLA(indices,MODULES,8,1,4,2,5,3.70,10,5.8,3.7,10,6,"Shower","Type A",TL,ORISET,FURNITE,DSCA)
		TL=[PDORS,PWIN,FURNITE]
		FURPLA(indices,MODULES,5,2,3,3,1.97,5.25,2.35,3.8,5.8,4,6,"Vanity_3_Drawers","1450 mm",TL,ORISET,FURNITE,DSCA)
		TL=[PDORS,PWIN,FURNITE]
		FURPLA(indices,MODULES,2,1,2,2,1.83,1.33,1.48,3.8,3.8,4,4,"Sanitary_Toilets_Roca_ISLAND-Wall-hung-WC1","00 - Glossy white",TL,ORISET,FURNITE,DSCA)

	elif PLACE=="MBATH":			
		TL=[PDORS,PWIN,FURNITE]
		FURPLA(indices,MODULES,14,1,5,2,6,3.5,3,6,3.7,3,6,"Simple_Freestanding_Tub2","6' Tub",TL,ORISET,FURNITE,DSCA)
		TL=[PDORS,PWIN,FURNITE]
		FURPLA(indices,MODULES,8,1,4,2,5,3.70,10,5.8,3.7,10,6,"Shower","Type A",TL,ORISET,FURNITE,DSCA)
		TL=[PDORS,PWIN,FURNITE]
		FURPLA(indices,MODULES,5,2,3,3,1.97,5.25,2.35,3.8,5.8,4,6,"Vanity_3_Drawers","1450 mm",TL,ORISET,FURNITE,DSCA)
		TL=[PDORS,PWIN,FURNITE]
		FURPLA(indices,MODULES,2,1,2,2,1.83,1.33,1.48,3.8,3.8,4,4,"Sanitary_Toilets_Roca_ISLAND-Wall-hung-WC1","00 - Glossy white",TL,ORISET,FURNITE,DSCA)

	elif PLACE=="ROOMS" or PLACE=="MASTER":			
		TL=[PDORS,PWIN,FURNITE]
		FURPLA(indices,MODULES,4,3,1,4,6.4,3.4,3,8,7.8,3,16,"00_Bedroom","2000x2000 KING",TL,ORISET,FURNITE,DSCA)
		TL=[PDORS,PWIN,FURNITE]
		FURPLA(indices,MODULES,4,3,2,4,2.32,7.5,10,3.5,7.8,10,8,"00_Wardrobe","2400x700",TL,ORISET,FURNITE,DSCA)
		TL=[PDORS,PWIN,FURNITE]
		FURPLA(indices,MODULES,2,1,3,1,1.97,3.8,2.35,3.8,3.8,4,6,"00_WorkTable","Default",TL,ORISET,FURNITE,DSCA)

	elif PLACE=="MASTERWARD" or PLACE=="WARD":			
		TL=[PDORS,FURNITE]
		FURPLA(indices,MODULES,4,3,2,4,2.32,7.5,10,2.32,7.9,10,8,"00_Wardrobe","2400x700",TL,ORISET,FURNITE,DSCA)
		TL=[PDORS,FURNITE]
		FURPLA(indices,MODULES,4,3,2,4,2.32,7.5,10,2.32,7.9,10,8,"00_Wardrobe","2400x700",TL,ORISET,FURNITE,DSCA)
		TL=[PDORS,FURNITE]
		FURPLA(indices,MODULES,4,3,2,4,2.32,7.5,10,2.32,7.9,10,8,"00_Wardrobe","2400x700",TL,ORISET,FURNITE,DSCA)
		TL=[PDORS,FURNITE]
		FURPLA(indices,MODULES,4,3,2,4,2.32,7.5,10,2.32,7.9,10,8,"00_Wardrobe","2400x700",TL,ORISET,FURNITE,DSCA)
	
	#elif PLACE=="KITCHEN":
	#	TL=[PDORS,PWIN,FURNITE]
	#	FURPLA(indices,MODULES,2,1,2,2,1.83,1.33,1.48,3.8,3.8,4,4,"Sanitary_Toilets_Roca_ISLAND-Wall-hung-WC1","00 - Glossy white",TL,ORISET,FURNITE,DSCA)			
	#	TL=[PDORS,PWIN,FURNITE]
	#	FURPLA(indices,MODULES,4,3,1,4,6.4,3.4,3,8,7.8,3,16,"00_Bedroom","2000x2000 KING",TL,ORISET,FURNITE,DSCA)
	#	TL=[PDORS,PWIN,FURNITE]
	#	FURPLA(indices,MODULES,4,3,2,4,2.32,7.5,10,3.5,7.8,10,8,"00_Wardrobe","2400x700",TL,ORISET,FURNITE,DSCA)
	#	TL=[PDORS,PWIN,FURNITE]
	#	FURPLA(indices,MODULES,2,1,3,1,1.97,3.8,2.35,3.8,3.8,4,6,"00_WorkTable","Default",TL,ORISET,FURNITE,DSCA)

	elif PLACE=="TOILET" :			
		TL=[PDORS,PWIN,FURNITE]
		FURPLA(indices,MODULES,5,2,3,3,1.97,5.25,2.35,3.8,5.8,4,6,"Vanity_3_Drawers","1450 mm",TL,ORISET,FURNITE,DSCA)
		TL=[PDORS,PWIN,FURNITE]
		FURPLA(indices,MODULES,2,1,2,2,1.83,1.33,1.48,3.8,3.8,4,4,"Sanitary_Toilets_Roca_ISLAND-Wall-hung-WC1","00 - Glossy white",TL,ORISET,FURNITE,DSCA)

	USED1=[]
	for a in indices:
		USED1.append(str(a)+"_"+str(DSCA.get(str(a))))
	
	return FURNITE


def farestposibleindex(PLOC,REFLOC,STORAGE):
	D=1
	S=PLOC[0]
	for a in PLOC:
		try:
			if a.ORI.DistanceTo(REFLOC.ORI)>D:
				D=a.ORI.DistanceTo(REFLOC.ORI)
				S=a
		except:
			if a.ORI.DistanceTo(REFLOC)>D:
				D=a.ORI.DistanceTo(REFLOC)
				S=a			
	STORAGE.append(S)

def centroido(sol):
	CENPOINTS=[]
	POINS=sol.Vertices
	POINSG=[]
	for b in POINS:
		POINSG.append(b.PointGeometry)
	X=0
	Y=0
	Z=0
	for b in POINSG:
		X+=b.X 
		Y+=b.Y
		Z+=b.Z
	XR=X/len(POINSG)
	YR=Y/len(POINSG)
	ZR=Z/len(POINSG)
	POINT=Point.ByCoordinates(XR,YR,ZR)
	return POINT

def surfpoints(SURFS):
	curves=SURFS.PerimeterCurves()
	points=[]
	
	for a in curves:
		points.append(a.StartPoint)
		points.append(curves[-1].EndPoint)
	
	xpo=[]
	ypo=[]
	zpo=[]
	for a in points:
		xpo.append(a.X)
		ypo.append(a.Y)
		zpo.append(a.Z)
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
	
	LINEX=Line.ByStartPointEndPoint(Point.ByCoordinates(minx,miny,minz),Point.ByCoordinates(maxx,miny,minz))
	LINEY=Line.ByStartPointEndPoint(Point.ByCoordinates(minx,miny,minz),Point.ByCoordinates(minx,maxy,minz))
	
	XN=math.ceil(LINEX.Length/8)
	YN=math.ceil(LINEX.Length/16)
	
	XGRIDS=[]
	XGRIDS1=[]
	for a in range(int(XN+1)):
		TL=LINEY.Translate(LINEX.Direction,4+a*8)
		INTL=TL.Intersect(SURFS)
		TL1=LINEY.Translate(LINEX.Direction,0+a*8)
		INTL1= TL1.Intersect(SURFS)
		XGRIDS.append(TL)
		XGRIDS1.append(INTL1)
		
	YGRIDS=[]
	YGRIDS1=[]
	for a in range(int(YN+1)):
		TL=LINEX.Translate(LINEY.Direction,8+a*16)
		INTL=TL.Intersect(SURFS)
		TL1=LINEX.Translate(LINEY.Direction,0+a*16)
		INTL1= TL1.Intersect(SURFS)
		YGRIDS.append(TL)
		YGRIDS1.append(INTL1)
		
	INTPOINTS=[]
	for a in XGRIDS:
		for b in YGRIDS:
			if b.DoesIntersect(a):
				if (a.Intersect(b)[0].DoesIntersect(SURFS)):
					INTPOINTS.append(a.Intersect(b)[0])	
				
	
	CLEARLOC=[]
	for a in INTPOINTS:
		REC=Surface.ByPatch(Rectangle.ByWidthLength(CoordinateSystem.ByOrigin(a),8,16))
		INTT=REC.Intersect(SURFS)[0]
		if INTT.Area==(8*16):
			CLEARLOC.append(a)
	return CLEARLOC	

def PANELINGLINES(RESLINES):
	ML=[]
	MA=[]
	CD=[]
	LEN=[]
	for a in RESLINES:
		if a.Length<4:
			ML.append(TOL(a).PointAtParameter(0.5))
			MA.append(TOL(a).Direction)
			LEN.append(a.Length)
			CD.append("11")
		else:
			DIVLEN=a.Length/4
			POIS=[]
			for b in range(int(DIVLEN)+1):
				POIS.append(a.PointAtSegmentLength(b*4))
			LINESS=a.SplitByPoints(POIS)
			for c in range(len(LINESS)):
				if LINESS[c].Length>1:
					ML.append(TOL(LINESS[c]).PointAtParameter(0.5))
					MA.append(TOL(LINESS[c]).Direction)
					LEN.append(round(LINESS[c].Length))	
					if c==len(LINESS)-1:
						CD.append("11")
					else:
						CD.append("10")	
				
	return ML,MA,LEN,CD
"""
def ZONINGCONFIGS(PLACE):
	CONFIG=[]
	if PLACE=="BATH1" or  PLACE=="BATH2"
		DATA=["WIN",0.5]
		DOOR=[0,2,0.25,2.88,10,3.7,3.7,12,4,"Door - Single Panel",'F - W - 31.5" x 120"']
		WINDOWS=[1,2,0.25,3.8,10,0.10,3.7,12,4,"Window - CIEL Vertical_SB_Test",'F - G - 45" x 120"']
		CONFIG.append([11,1,5,2,6,3.5,3,6,3.7,3.6,"Simple_Freestanding_Tub2","6' Tub"])
		CONFIG.append([8,1,4,2,5,3.70,10,5.8,3.7,10,6,"Shower","Type A"])
		CONFIG.append([5,2,3,3,1.97,5.25,2.35,3.8,5.8,4,6,"Vanity_3_Drawers","1450 mm"])
		CONFIG.append([2,1,2,2,1.83,1.33,1.48,3.8,3.8,4,4,"Sanitary_Toilets_Roca_ISLAND-Wall-hung-WC1","00 - Glossy white"])
"""


#def MACREL(NAME):

def FURPLA(index1,MODSS,N,S,I,M,W,L,H,CW,CL,CH,SI,FA,TY,LST,ORIF,FURL,DSCA):
	NUM=len(index1)
	if NUM>N:
		PFUR=[]
		for a in MODSS:
			if DSCA[str(a.IOF)]>=S and a.IOF in index1:
				FUR=furalloc(I,M,W,L,H,CW,CL,CH,SI,FA,TY)
				FUR=furnite(FUR,a)
				TL=LST
				if notclashingany(FUR,TL)<1:
					PFUR.append(FUR)
		if len(PFUR)>0:
			farestposibleindex(PFUR,ORIF,FURL)
			tclean(index1,MODSS,FURL[-1].CLEASOL,DSCA)

def ModNorVec(VALUE,VERS):
	YS=[]
	if VALUE=="X":
		for a in VERS:
			YS.append(a.PointAtParameter(0.5).X)
	else:
		for a in VERS:
			YS.append(a.PointAtParameter(0.5).Y)
	YS2=list(set(YS))
	YS2.sort()
	N=int(len(YS2)/2)
	ROU=round(YS2[N],2)
	VALS=[]
	for a in YS2:
		VALS.append(a-(ROU+(round(round(a-ROU,2)/2,0)*2)))
	indexes=[]
	for a in YS:
		indexes.append(YS2.index(a))
	XValues=[]
	for a in indexes:
		XValues.append(-VALS[a])
	return XValues


def shortlines(ln):
	try:
		NEW=Line.ByStartPointEndPoint(ln.PointAtParameter(0.1),ln.PointAtParameter(0.9))
	except:
		NEW=ln
	return NEW
def TOLs(ln):
	NEW=Line.ByStartPointEndPoint(ln.StartPoint,ln.EndPoint)
	return NEW