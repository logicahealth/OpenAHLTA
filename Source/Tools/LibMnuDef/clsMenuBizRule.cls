VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenuBizRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public MenuId As Integer
Public MenuRule As eBizRulz
Public MenuItemID As Integer
Public MenuSurface As Integer
Public MenuSurfacePlus As Integer
Public MenuToothSequence As eToothSequence
Public MenuToothPosition As eToothPosition
Public MenuToothType As eToothType
Public MenuCodeID As Integer
Public MenuToothEndo As Integer
Public MenuToothMax As Integer


'Public Enum eEndodonicTherapy
'  eEndo = -1
'  eRetreat = 0
'  eNone = 777
'End Enum

Public Enum eMouth
  eMaxillary = -1
  eMandibular = 0
  eNone = 777
End Enum

'Public Enum eToothSequence
'  ePrimary = -1
'  eSecondary = 0
'  eNone = 777
'End Enum
'
'Public Enum eToothPosition
'  ePosterior = -1
'  eAnterior = 0
'  eNone = 777
'End Enum
'
'Public Enum eToothType
'  ePreMolar = -1
'  eMolar = 0
'  eNone = 777
'End Enum
'
'Public Enum eBizRulz
'  eSurface = 0
'  eSurfaceToothSeq = 1
'  eSurfaceToothSeqPos = 2
'  eExtract = 3
'  eToothSeqPos = 4
'  eToothSeqOnly = 5
'  eMissing = 6
'  eToothPosOnly = 7
'  eToothSeqType = 8
'  eToothTypeOnly = 9
'  eMultiMissing = 10
'  eImplants = 11
'  eMaxMan = 12
'  eNone = 99
'End Enum
