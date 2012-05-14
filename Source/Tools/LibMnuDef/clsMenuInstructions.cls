'***************************************************************************************
'
'  Copyright (c) 2007-2012 Northrop Grumman Corporation
'
'  Licensed by Tricare Management Activity under license from the Copyright owner.
'
'  This text file must be included in all Derivative Works of the licensed Source Code.
'
'***************************************************************************************

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenuInstructions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public MenuItemID As Integer
Public Color As eColor
Public Hatch As eHatch
Public Symbol As eSymbol
Public Operation As eOperation
Public Location As eLocation
Public RemoveLocations As Integer
Public xSize As Integer
Public ySize As Integer
Public xOrientation As eXOrientation
Public yOrientation As eYOrientation
Public isOutlined As Boolean


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
