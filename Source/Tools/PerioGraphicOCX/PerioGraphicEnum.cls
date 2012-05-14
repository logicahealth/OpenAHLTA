VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PerioGraphicEnum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Rem Constant defined by Perio Graphic OCX
'For internal use
Public Enum ePerioGrphElem
  GM
  CAL
  MGJ
  Bleed
  Plaq
  Supp
  Mob
  Furc
  BL
  PD
End Enum
  

'For external use
Public Enum ePerioToothTypes
  perPrimary = 1
  perSecondary = 2
  perSuper = 3
  perNoDraw = 4
  perImplant = 5
End Enum

'Public Enum ePerioViews
'  perFullMouth = 1
'  perMaxillary = 2
'  perMandibular = 3
'  perUR = 4
'  perUL = 5
'  perLL = 6
'  perLR = 7
'End Enum

Rem Constant from PerioBL "enums.cls" - Perio Business Layer
Public Enum eViewType
    perFullMouth
    perMaxArch
    perManArch
    perURQ
    perULQ
    perLRQ
    perLLQ
    perTextEntry
    perPrintGraphic
End Enum

Public Enum eConstant
    perUnknownValue = -5765
    MagicNumber = -5999
End Enum

Public Enum eFurcation
    perFurcNone
    perperFurcOne
    perFurcTwo
    perFurcThree
    perFurcFour
End Enum

Public Enum eMobility
    perMobNone
    perMobOne
    perMobTwo
    perMobThree
End Enum

Public Enum eSVGPathCommands
    StartPath
    RepeatLastCmd
    Mov_abs
    Mov_rel
    LnTo_abs
    LnTo_rel
    HLn_abs
    HLn_rel
    VLn_abs
    VLn_rel
    CurTo_abs       'cubic Bezier curve, not implemented
    CurTo_rel       'cubic Bezier curve, not implemented
    Smoo_abs        'smooth cubic Bezier curve,  not implemented
    Smoo_rel        'smooth cubic Bezier curve,  not implemented
    QuadBez_abs     'quadratic Bezier curve,  not implemented
    QuadBez_rel     'quadratic Bezier curve,  not implemented
    SmooQuad_abs    'smooth quadratic Bezier curve,  not implemented
    SmooQuad_rel    'smooth quadratic Bezier curve,  not implemented
    Arc_abs         'Arc, not implemented
    Arc_rel         'Arc, not implemented
    Close_abs
    Close_rel
End Enum
