import clr
clr.AddReference("RevitAPIUI")
from  Autodesk.Revit.UI import *
clr.AddReference("RevitAPI")
import Autodesk
import math
import random
from Autodesk.Revit.DB import *
clr.AddReference("RevitNodes")
import Revit
clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)

def get_intersection(line1, line2):
	results = clr.Reference[IntersectionResultArray]()
	result = line1.Intersect(line2, results)
	if result != SetComparisonResult.Overlap:
		print('No Intesection')
	intersection = results.Item[0]
	return intersection.XYZPoint

def get4IP(LS):
	PTS=[]
	PTS.append(get_intersection(LS[2],LS[0]))
	PTS.append(get_intersection(LS[2],LS[1]))
	PTS.append(get_intersection(LS[3],LS[1]))
	PTS.append(get_intersection(LS[3],LS[0]))
	return PTS

def PTStoLINES(lst):
	LNS=[]
	for a in range(len(lst)):
		if a<len(lst)-1:
			LNS.append(Line.CreateBound(lst[a],lst[a+1]))
	LNS.append(Line.CreateBound(lst[-1],lst[0]))
	return LNS
	
def PlaneFunction(LINE):
	p = LINE.GetEndPoint(0)
	q = LINE.GetEndPoint(1)
	norm = XYZ.BasisZ
	plane = Plane.CreateByNormalAndOrigin(norm,p)
	return plane;

def TranslateLine(Lnn,Vec,DS):
    SP=Lnn.GetEndPoint(0)
    EP=Lnn.GetEndPoint(1)
    FL=Line.CreateBound(SP.Add(Vec.Multiply(DS)),EP.Add(Vec.Multiply(DS)))
    return FL

def TranslatePoint(Pt,Vec,DS):
    FL=Pt.Add(Vec.Multiply(DS))
    return FL

#Create Grid Lines and Values
def CMODGRID3(LINE,OLINE,VEC):
    NAL=float(int(LINE.Length/16))-1
    NALL=float(int((((LINE.Length)-(NAL*16))/8)))
    NALLL=(LINE.Length)-(NAL*16)-(NALL*8)
    TOT=(NAL*16+NALL*8)
    MED=TOT/2
    DIF=NALLL/2
    STARTLINELL=TranslateLine(OLINE,VEC.Negate(),MED)
    VALUES = []
    DIS=[]
    
    if NALL>2:
		X=12
		for a in range(int(NAL)):
			VALUES.append(16)
			DIS.append(X)
			X+=16			
		VALUES.append(16)		
		DIS.append(X)
	
    elif NALL>1:
	    X=16
	    for a in range(int(NAL)):
	        VALUES.append(16)
	        DIS.append(X)
	        X+=16
    GRIDLINES=[]
    for a in DIS:
	    NL=TranslateLine(STARTLINELL,VEC,a)
	    GRIDLINES.append(NL)	
    return GRIDLINES,VALUES

class CMOD:
    def _init_(self,IOF,SOL,ORI,CEN,WID,LEN,HEI,USE,SC1,SC2,SC3,SC4,SC5,SC6,CSYS,VEC):
	    self.IOF=IOF
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

def CreateCMODGRID(XGRIDS,YGRIDS):
    FORMACRO=[]
    X=1
    for e in range(2):
        Y=e*12
        for a,b,f in zip(XGRIDS[0],XGRIDS[1],range(int(len(XGRIDS[0])))):
            for c,d,g in zip(YGRIDS[0],YGRIDS[1],range(int(len(YGRIDS[0])))):
                M1=CMOD()
                M1.IOF=X
                M1.ORI=TranslatePoint(get_intersection(a,c),XYZ(0,0,Y),1)
                M1.WID=d
                M1.LEN=b
                M1.HEI=12
                M1.CEN=TranslatePoint(M1.ORI,XYZ(0,0,1),M1.HEI/2)
                M1.USE=0
                FORMACRO.append(M1)
                if f==0:
                    M1.SC1=0
                else:			
                    M1.SC1=((f-1)*(len(YGRIDS[0])))+g+1+(e*(len(XGRIDS[0]))*(len(YGRIDS[0])))
                if f==len(XGRIDS[0])-1:
                    M1.SC3=0
                else:
                    M1.SC3=((f+1)*(len(YGRIDS[0])))+g+1+(e*(len(XGRIDS[0]))*(len(YGRIDS[0])))
                if g==len(YGRIDS[0])-1:
                    M1.SC2=0
                else:			
                    M1.SC2=(f*(len(YGRIDS[0])))+g+2+(e*(len(XGRIDS[0]))*(len(YGRIDS[0])))
                if g==0:
                    M1.SC4=0
                else:			
                    M1.SC4=(f*(len(YGRIDS[0])))+g+(e*(len(XGRIDS[0]))*(len(YGRIDS[0])))
                if e==0:
                    M1.SC5=0
                else:			
                    M1.SC5=(f*(len(YGRIDS[0])))+g+1
                if e==1:
                    M1.SC6=0
                else:			
                    M1.SC6=(f*(len(YGRIDS[0])))+g+1+(len(XGRIDS[0]))*(len(YGRIDS[0]))
                X=X+1
    return FORMACRO

def ARCHPROG(MODS):
    PROG=[]
    if MODS>5:
        PROG.append("01_GARAGE")
        PROG.append("02_ENTRANCE")
        PROG.append("03_DINNER")
        PROG.append("04_LIVING")
        PROG.append("05_KITCHEN")
        PROG.append("06_TERRACE")
    if MODS>7:
        PROG.append("07_DINNER")
        PROG.append("08_LIVING")
    if MODS>9:
        PROG.append("09_KITCHEN")
        PROG.append("10_TERRACE")
    if MODS>11:
        PROG.append("11_OFFICE")
        PROG.append("12_TERRACE")
	TOT=len(PROG)	
	NM=MODS-len(PROG)
	for a in range(int(NM)):
		ENEM=TOT+a+1
		PROG.append(str(ENEM)+"_TERRACE")
    UPPROG=UPARCHPROG(PROG)
    for a in UPPROG:
        PROG.append(a)
    return PROG

def UPARCHPROG(DPROG):
    MODS=0
    for a in DPROG:
        if "TERRACE" not in a:
            MODS+=1
    PROG=[]
    if MODS>6:
        PROG.append("01_BATH")
        PROG.append("02_MBATH")
        PROG.append("03_MASTER")
        PROG.append("04_ROOM")
        PROG.append("05_ROOM")
    if MODS>7:
        PROG.append("06_ROOM")
    if MODS>8:
        PROG.append("07_MASTER")
    if MODS>9:
        PROG.append("08_BATH")
    if MODS>10:
        PROG.append("09_MBATH")
    if MODS>11: 
        PROG.append("10_ROOM")
    TOT=len(PROG)	
    NM=MODS-len(PROG)
    for a in range(int(NM)):
        ENEM=TOT+a+1
        if ENEM<10:
            PROG.append("0"+str(ENEM)+"_VOID")
        else:
            PROG.append(str(ENEM)+"_VOID")
    return PROG

def CMACRONSFROMPROG(PROG):
	MACRONS=[]
	for a,b in zip(PROG,range(len(PROG))):
		MCN=MACRON()
		MCN.IOF=b+1
		MCN.NAM=a
		CONECS=POSIBLECONECTORS(a) 
		SOCKS=[]
		NS=0
		for c in CONECS:
			LAA=TPB(c,PROG,a)
			SOCKS.append(LAA)
			if "NONE" not in LAA:  
				NS+=len(LAA) 
		MCN.NS=NS			
		MCN.SOCKETS=SOCKS
		MACRONS.append(MCN)
	return MACRONS

class MACRON:
	def _init_(self,IOF,MDNIOF,NAM,MDNM,SOCKETS,NS):
		self.IOF=IOF
		self.MDNIOF=MDNIOF
		self.NAM=NAM
		self.MDNM=MDNM
		self.SOCKETS=SOCKETS
		self.NS=NS

def POSIBLECONECTORS(NAME):
	CONS=[]
	if "GARAGE" in NAME:
		CONS.append(["00"])
		CONS.append(["ENT","LIV","DIN","KIT","TER","OFF"])
		CONS.append(["LIV","DIN","KIT","TER"])
		CONS.append(["ENT","LIV","DIN","KIT","TER","OFF"])
		CONS.append(["00"])
		CONS.append(["VOI","ROO","BAT","MBA","MAS"])
	elif "ENTRANCE" in NAME:
		CONS.append(["00"])
		CONS.append(["GAR","LIV","DIN","OFF"])
		CONS.append(["LIV","DIN"])
		CONS.append(["GAR","LIV","DIN","OFF"])
		CONS.append(["00"])
		CONS.append(["VOI","ROO","BAT","MBA","MAS"])
	elif "DINNER" in NAME:
		CONS.append(["GAR","LIV","DIN","KIT"])
		CONS.append(["LIV","DIN","KIT","TER"])
		CONS.append(["LIV","DIN","KIT","TER"])
		CONS.append(["LIV","DIN","KIT","TER"])
		CONS.append(["00"])
		CONS.append(["VOI","ROO","BAT","MBA","MAS"])
	elif "LIVING" in NAME:
		CONS.append(["GAR","ENT","LIV","DIN"])
		CONS.append(["LIV","DIN","TER"])
		CONS.append(["LIV","DIN","TER"])
		CONS.append(["LIV","DIN","TER"])
		CONS.append(["00"])
		CONS.append(["VOI","ROO","BAT","MBA","MAS"])
	elif "KITCHEN" in NAME:
		CONS.append(["GAR","DIN","KIT","LIV"])
		CONS.append(["DIN","KIT","TER"])
		CONS.append(["DIN","KIT","TER"])
		CONS.append(["DIN","KIT","TER"])
		CONS.append(["00"])
		CONS.append(["BAT","MBA"])
	elif "TERRACE" in NAME:
		CONS.append(["GAR","LIV","DIN","KIT","TER"])
		CONS.append(["LIV","DIN","KIT","TER"])
		CONS.append(["LIV","DIN","KIT","TER"])
		CONS.append(["LIV","DIN","KIT","TER"])
		CONS.append(["00"])
		CONS.append(["00"])
	elif "OFFICE" in NAME:
		CONS.append(["GAR","ENT","DIN","KIT","TER"])
		CONS.append(["GAR","ENT","DIN","TER"])
		CONS.append(["ENT","DIN","TER"])
		CONS.append(["GAR","ENT","DIN","TER"])
		CONS.append(["00"])
		CONS.append(["VOI","ROO","BAT","MBA","MAS"])
	elif "ROOM" in NAME:
		CONS.append(["ROO","BAT","MAS","MBA"])
		CONS.append(["ROO","BAT","MAS","MBA"])
		CONS.append(["ROO","BAT","MAS","MBA"])
		CONS.append(["ROO","BAT","MAS","MBA"])
		CONS.append(["DIN","KIT","TERR","GAR","ENT","OFF"])
		CONS.append(["00"])
	elif "BATH" in NAME:
		CONS.append(["ROO","BAT","MBA"])
		CONS.append(["ROO","BAT","MBA"])
		CONS.append(["ROO","BAT","MBA"])
		CONS.append(["ROO","BAT","MBA"])
		CONS.append(["DIN","KIT","TERR","GAR","ENT","OFF"])
		CONS.append(["00"])
	elif "MBATH" in NAME:
		CONS.append(["MAS","BAT","MBA"])
		CONS.append(["MAS","BAT","MBA"])
		CONS.append(["MAS","BAT","MBA"])
		CONS.append(["MAS","BAT","MBA"])
		CONS.append(["DIN","KIT","TERR","GAR","ENT","OFF"])
		CONS.append(["00"])
	elif "MASTER" in NAME:
		CONS.append(["BAT","MBA"])
		CONS.append(["BAT","MBA"])
		CONS.append(["BAT","MBA"])
		CONS.append(["BAT","MBA"])
		CONS.append(["DIN","KIT","TERR","GAR","ENT","OFF"])
		CONS.append(["00"])
	elif "VOID" in NAME:
		CONS.append(["00"])
		CONS.append(["00"])
		CONS.append(["00"])
		CONS.append(["00"])
		CONS.append(["00"])
		CONS.append(["00"])
	return CONS	

def TPB(lst,lst2,NAME):
	SELECTED=[]
	for a in lst:
		if a=="00":
			SELECTED.append("00")
		else:
			for b in lst2:
				if a in b:
					if b!=NAME:
						SELECTED.append(b)
	if len(SELECTED)<1:
		SELECTED.append("NONE")
	SEL2=list(set(SELECTED))
	SEL2.sort()
	return SEL2

def CMODOLSFROMMODULES(MODULES,MACRONS):
	CMODOLS=[]
	for a in MODULES: 
		MDLN=MODON()
		MDLN.IOF=a.IOF
		MDLN.MM=a
		NAMEL=[]
		for b in MACRONS:
			LSOC=b.SOCKETS[0]  
			if LSOC[0]=="00":
				if a.SC1==0:
					NAMEL.append(b.NAM)
			else:
				NAMEL.append(b.NAM)
		MDLN.NAML=NAMEL
		SOCKETS=[a.SC1,a.SC2,a.SC3,a.SC4,a.SC5,a.SC6]
		MDLN.SOCKETS=SOCKETS
		CMODOLS.append(MDLN)
	return CMODOLS

class MODON:
	def _init_(self,IOF,NAML,MM,SOCKETS):
		self.IOF=IOF
		self.NAML=NAML
		self.MM=MM
		self.SOCKETS=SOCKETS

def WCFSELECT(PROG,SPOTS,IOFS,SELMOD,SELPRO,PRIGROU,CMODOLS,NUMS):
	#SELECTMACRON
	DUPLI=[]
	if len(SELPRO)>0:
		for a in PROG:
			A=SELPRO[-1]
			TXT=A[3:7]
			if TXT in a:
				DUPLI.append(a)
	if len(DUPLI)==0:
		try:
			MINS=min(SPOTS) 
			SELINDEX=SPOTS.index(MINS)
			ELEMENT=PROG[SELINDEX]
		except:
			SELINDEX=0
			ELEMENT=PROG[0]
		if "TERR" in ELEMENT:
			SELINDEX=0
			ELEMENT=PROG[0]
	else:
		SELINDEX=PROG.index(DUPLI[0])
		ELEMENT=DUPLI[0]

	#SELECTMODULE
	PRISEL=[]
	for a in IOFS[SELINDEX]:
		if a in PRIGROU:
			PRISEL.append(a)
	if len(PRISEL)==0:
		if len(IOFS[SELINDEX])>0:
			for a in IOFS[SELINDEX]:
				PRISEL.append(a)
			CHOICE=random.choice(PRISEL)  
		else:
			POSI=list(set(PRIGROU))
			POSI.sort()
			CHOICE=POSI[0]
	else:
		CHOICE=random.choice(PRISEL) 
	
	SA=SOCKETSAVAILABLE(CMODOLS,SELMOD,NUMS)
	if len(SA)>0:
		if len(DUPLI)>0:
			CHOICE=SA[0]

	SELMOD.append(CHOICE) 
	SELPRO.append(ELEMENT) 
	return ELEMENT

def SOCKETSAVAILABLE(CMODOLS,SELMOD,NUMS):
	SOC=[]
	if len(SELMOD)>0:
		try:
			for a in(CEBIOF(SELMOD[-1],CMODOLS).SOCKETS):
				if a!=0 and a in NUMS:
					if a not in SELMOD:
						SOC.append(a)
		except:
			pass
	return SOC

def WCFCOLLAPSE(PROG,MACROC,ELEMENT):
	#COLLAPSE
	PROG.remove(ELEMENT)
	MACROT=[]
	for a in MACROC:
		if a.NAM!=ELEMENT: 
			NSOCKS=[]
			NP=0
			for b in a.SOCKETS:
				lnew=b
				if ELEMENT in lnew:
					lnew.remove(ELEMENT) 
				if len(lnew)==0:
					lnew.append("NONE")
				if "NONE" not in lnew: 
					NP+=len(lnew)
				NSOCKS.append(lnew)
			a.SOCKETS=NSOCKS
			a.NS=NP
			MACROT.append(a)
	return MACROT

def WCFPROPAGATE(SELMOD,SELPRO,MODC,CMODOLS,MACRONS,ELEMENT,NUMS):
	#PROPAGATE
	PRIGROU2=[]
	for a,b in zip(SELMOD,SELPRO):
		try:
			for c,d in zip(CEBIOF(a,CMODOLS).SOCKETS,CEBNAM(b,MACRONS).SOCKETS):
				if c!=0 and c in NUMS:
					NLNA=[]
					if c not in SELMOD:
						MODI=CEBIOF(c,MODC)
						for e in MODI.NAML:
							if e!="00" and e!="NONE":
								if e in d:
									NLNA.append(e)
							else:
								NLNA.append(e)
						MODI.NAML=NLNA
						if c not in PRIGROU2:
							PRIGROU2.append(c)
		except:
			PRIGROU2.append(GETLOWEREA(MODC))
		
	MODT=[]
	for a in MODC:
		if a.IOF not in SELMOD:
			list=a.NAML
			if ELEMENT in a.NAML:
				list.remove(ELEMENT)
			else:
				pass
			a.NAML=list
			MODT.append(a)
	return MODT,PRIGROU2

def WCFEVALUATE(PROG,MODC):
	#EVALUATE
	SP1=[]
	IO1=[]
	for a in PROG:
		X=0
		IOSS=[]
		for b in MODC:
			if a in b.NAML:
				X+=1
				IOSS.append(b.IOF)
		IO1.append(IOSS)
		SP1.append(X) 
	return SP1,IO1

def GETLOWEREA(MOD):
	IOFS=[]
	for a in MOD:
		IOFS.append(a.IOF)
	IOFS=list(set(IOFS))
	IOFS.sort()
	return IOFS[0]