# Load the Python Standard and DesignScript Libraries
import sys
sys.path.append(r"C:\BIM\CIEL")
from ModClass import *
import clr
clr.AddReference('ProtoGeometry')
from Autodesk.DesignScript.Geometry import *

# The inputs to this node will be stored as a list in the IN variables.
MACROS=IN[0]
RELEASEMACROS=[]
CENT=[]
LIV=0
DIN=0
ENT=0
VOI=0
for a in MACROS:
	if not (a.NAM == "LIVING" or a.NAM== "DINNER" or a.NAM== "ENTRANCE" or a.NAM=="VOID"):
		RELEASEMACROS.append(a)
	else:
		if a.NAM=="LIVING":
			LIV=1
		elif a.NAM=="DINNER":
			DIN=1
		elif a.NAM=="ENTRANCE":
			ENT=1
		elif a.NAM=="VOID":
			VOI=1
	CENT.append(a.CEN)
if LIV==1:		
	LIVING=allmacrosnamed(MACROS,"LIVING")
	RELEASEMACROS.append(unifiedvalues(LIVING))
if DIN==1:
	DINNER=allmacrosnamed(MACROS,"DINNER")
	RELEASEMACROS.append(unifiedvalues(DINNER))
if ENT==1:
	ENTRANCE=allmacrosnamed(MACROS,"ENTRANCE")
	RELEASEMACROS.append(unifiedvalues(ENTRANCE))
if VOI==1:
	VOID=allmacrosnamed(MACROS,"VOID")
	RELEASEMACROS.append(unifiedvalues(VOID))

ALLGEOS=[]
ALLNAMES=[]
ALLORIGINS=[]
WS=[]
HS=[]
LS=[]

for a,b in zip(RELEASEMACROS,range(len(RELEASEMACROS))):
	a.IOF=b+1
	ALLGEOS.append(a.SOL)
	ALLNAMES.append(a.NAM)
	ALLORIGINS.append(a.ORI)
	WS.append(a.WID)
	LS.append(a.LEN)
	HS.append(a.HEI)
	
	
	
# Assign your output to the OUT variable.
OUT = RELEASEMACROS,ALLGEOS,ALLNAMES,ALLORIGINS,WS,LS,HS