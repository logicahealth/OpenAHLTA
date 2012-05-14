VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenuItemInstr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'Public CodeID As Integer
'Public Color As eColor
'Public Hatch As eHatch
'Public Symbol As eSymbol
'Public Operation As eOperation
'Public Location As eLocation
'Public RemoveLocations As Integer
'Public xSize As Integer
'Public ySize As Integer
'Public xOrientation As eXOrientation
'Public yOrientation As eYOrientation
'Public isOutlined As Boolean

'local variable(s) to hold property value(s)
Private mvarMenuInstructionID As Long
Private mvarGroup As eGroup
Private mvarCodeID As Long

Private mvarColor As eColor
Private mvarHatch As eHatch
Private mvarSymbol As eSymbol
Private mvarOperation As eOperation
Private mvarLocation As eLocation
Private mvarRemovedLocations As Integer


Public Enum eColor
  eBlue
  eCyan
  eGreen
  ePurple
  eOrange
  ePink
  eRed
  eYellow
  eMagenta
  ePowderBlue
  ePastelBlue
  eBabyBlue
  eElectricBlue
  eTwilightBlue
  eNavyBlue
  eDeepNavyBlue
  eDesertBlue
  eSkyBlue
  eIceBlue
  eLightBlueGreen
  eOceanGreen
  eMossGreen
  eDarkGreen
  eForestGreen
  eGrassGreen
  eKentuckyGreen
  eLightGreen
  eSpringGreen
  eTurguois
  eSeaGreen
  eFadedGreen
  eGhostGreen
  eBlack
  eWhite
  eGrey
End Enum

Public Enum eHatch
  eHorizontal
  eVertical
  eDiagLeft
  eDiagRight
  eBothVerticalHorizontal
  eBothDiagRightLeft
End Enum

Public Enum eSymbol
  eTriangle
  eLine
  eParallelLines
  eForwardSlash
  eFilledCircle
  eFilledArc
  eOutlineCircle
  eTextS
  eTextO
  eTextW
  eTextE
  eTextC
  eTextANK
  eTextNF
  eHorizontalScroll1
  eHorizontalScroll2
  eHorizontalScroll3
  eSuperNumerary4
  ePolkDots
  eImplant
  eMissing
  eLineBox
  eTigerClaw
  eOutlineCircleDashed
  eFilledCircleMid
  eTextAll
End Enum

Public Enum eLocation
  eSelectedSurfaces
  eCoronalSection
  eRootSection
  eWholeTooth
End Enum

Public Enum eXOrientation
  eLeft
  eCenter
  eRight
End Enum

Public Enum eYOrientation
  eTop
  eCenter
  eBottom
End Enum

Public Enum eOperation
  eFillSurface
  eOutlineSurface
  eHatchSurface
  eSymbolizeSurface
  eDoNothing
End Enum

Friend Property Let CodeID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CodeID = 5
    mvarCodeID = vData
End Property


Public Property Get CodeID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CodeID
    CodeID = mvarCodeID
End Property



Friend Property Let Group(ByVal vData As eGroup)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Group = Form1
    mvarGroup = vData
End Property


Public Property Get Group() As eGroup
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Group
    Group = mvarGroup
End Property



Friend Property Let MenuInstructionID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MenuInstructionID = 5
    mvarMenuInstructionID = vData
End Property


Public Property Get MenuInstructionID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MenuInstructionID
    MenuInstructionID = mvarMenuInstructionID
End Property



Friend Property Let RemovedLocations(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RemovedLocations = 5
    mvarRemovedLocations = vData
End Property


Public Property Get RemovedLocations() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RemovedLocations
    RemovedLocations = mvarRemovedLocations
End Property



Friend Property Let Location(ByVal vData As eLocation)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Location = 5
    mvarLocation = vData
End Property


Public Property Get Location() As eLocation
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Location
    Location = mvarLocation
End Property



Friend Property Let Operation(ByVal vData As eOperation)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Operation = 5
    mvarOperation = vData
End Property


Public Property Get Operation() As eOperation
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Operation
    Operation = mvarOperation
End Property



Friend Property Let Symbol(ByVal vData As eSymbol)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Symbol = 5
    mvarSymbol = vData
End Property


Public Property Get Symbol() As eSymbol
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Symbol
    Symbol = mvarSymbol
End Property



Friend Property Let Hatch(ByVal vData As eHatch)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Hatch = 5
    mvarHatch = vData
End Property


Public Property Get Hatch() As eHatch
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Hatch
    Hatch = mvarHatch
End Property



Friend Property Let Color(ByVal vData As eColor)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Color = 5
    mvarColor = vData
End Property


Public Property Get Color() As eColor
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Color
    Color = mvarColor
End Property



