VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "patTooth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public ToothDefinition As CHCSII_Dental_SVG_Lib.clsSVGToothDef
Public colSurfaceDefinitions As Collection
Public bToothVisible As Boolean

Public GCDxs As colGCDxs
'Public colDX As colPatDX

Public GCTxs As colGCTxs
'Public colTX As colPatTX

Public RemoveLocation As Integer
Public IsImplant As Boolean
Public IsExtractOrMissing As Boolean
Public HasPrevEndo As Boolean
Public HasDisplacement As Boolean 'Orientation has been chosen for this tooth
Public ToothID As Integer
Public dDateTimeRootRemoved As Date
Public dDateTimePlannedRootRemoved As Date
Public dDateTimeCrownRemoved As Date
Public dDateTimePonticDenture As Date
Public dDateTimePlannedCrownRemoved As Date
Public dDateTimeWholeToothRemoved As Date
Public dDateTimeWholeToothPlanToRemove As Date
Public dDateTimeImplant As Date
Public dDateTimeSupernumary As Date
Public dDateTimeUnErupted As Date
Public dDateTimeErupted As Date
Public Sub clearAllDates()
 dDateTimeRootRemoved = Empty
 dDateTimeCrownRemoved = Empty
 dDateTimePonticDenture = Empty
 dDateTimeWholeToothRemoved = Empty
 dDateTimeWholeToothPlanToRemove = Empty
 dDateTimeImplant = Empty
 dDateTimeSupernumary = Empty
 dDateTimeUnErupted = Empty
 dDateTimeErupted = Empty
End Sub
Private Sub Class_Initialize()
  Set colSurfaceDefinitions = New Collection
  Set GCDxs = New colGCDxs
  Set GCTxs = New colGCTxs
End Sub

Private Sub Class_Terminate()
  Set colSurfaceDefinitions = Nothing
  Set GCDxs = Nothing
  Set GCTxs = Nothing
End Sub
