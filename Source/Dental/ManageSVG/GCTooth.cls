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
Attribute VB_Name = "clsGCTooth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarToothID As Integer
Private mvarRemoveLocation As Integer
Private mvarIsImplant As Boolean
Private mvarIsExtractOrMissing As Boolean
Private mvarHasPrevEndo As Boolean
Private mvarHasDisplacement As Boolean 'Orientation has been chosen for this tooth
Private mvarDateTimeRootRemoved As Date
Private mvarDateTimePlannedRootRemoved As Date
Private mvarDateTimeCrownRemoved As Date
Private mvarDateTimePonticDenture As Date
Private mvarDateTimePlannedCrownRemoved As Date
Private mvarDateTimeWholeToothRemoved As Date
Private mvarDateTimeWholeToothPlanToRemove As Date
Private mvarDateTimeImplant As Date
Private mvarDateTimeSupernumary As Date
Private mvarDateTimeUnErupted As Date
Private mvarDateTimeErupted As Date
Private mvarToothVisible As Boolean
Private mvarSVGToothDef As CHCSII_Dental_SVG_Lib.clsSVGToothDef
Private mvarGCDxs As colGCDxs
Private mvarGCTxs As colGCTxs
Private mvarSVGSurfaces As CHCSII_Dental_SVG_Lib.colSVGSurfaces

'This is for the "invisible" teeth that have not erupted yet
Private mvarIsToothPlaceHolder As Boolean

Public Sub clearAllDates()
    mvarDateTimeRootRemoved = Empty
    mvarDateTimeCrownRemoved = Empty
    mvarDateTimePonticDenture = Empty
    mvarDateTimeWholeToothRemoved = Empty
    mvarDateTimeWholeToothPlanToRemove = Empty
    mvarDateTimeImplant = Empty
    mvarDateTimeSupernumary = Empty
    mvarDateTimeUnErupted = Empty
    mvarDateTimeErupted = Empty
End Sub
Public Property Set GCDxs(vData As colGCDxs)
    
    Set mvarGCDxs = vData
    
End Property

Public Property Get GCDxs() As colGCDxs
    
    Set GCDxs = mvarGCDxs
    
End Property


Public Property Set GCTxs(vData As colGCTxs)
    
    Set mvarGCTxs = vData
    
End Property

Public Property Get GCTxs() As colGCTxs
    
    Set GCTxs = mvarGCTxs
    
End Property


Public Property Let IsToothPlaceHolder(vData As Boolean)

    mvarIsToothPlaceHolder = vData
    
End Property

Public Property Get IsToothPlaceHolder() As Boolean
    
    IsToothPlaceHolder = mvarIsToothPlaceHolder
    
End Property

Public Property Set SVGSurfaces(vData As CHCSII_Dental_SVG_Lib.colSVGSurfaces)
    
    Set mvarSVGSurfaces = vData
    
End Property

Public Property Get SVGSurfaces() As CHCSII_Dental_SVG_Lib.colSVGSurfaces
    
    Set SVGSurfaces = mvarSVGSurfaces
    
End Property

Public Property Set SVGToothDef(vData As CHCSII_Dental_SVG_Lib.clsSVGToothDef)
    
    Set mvarSVGToothDef = vData
    
End Property

Public Property Get SVGToothDef() As CHCSII_Dental_SVG_Lib.clsSVGToothDef
    
    Set SVGToothDef = mvarSVGToothDef
    
End Property

Public Property Let ToothVisible(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ToothVisible = 5
    mvarToothVisible = vData
End Property


Public Property Get ToothVisible() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ToothVisible
    ToothVisible = mvarToothVisible
End Property



Public Property Let DateTimeErupted(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DateTimeErupted = 5
    mvarDateTimeErupted = vData
End Property


Public Property Get DateTimeErupted() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DateTimeErupted
    DateTimeErupted = mvarDateTimeErupted
End Property



Public Property Let DateTimeUnErupted(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DateTimeUnErupted = 5
    mvarDateTimeUnErupted = vData
End Property


Public Property Get DateTimeUnErupted() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DateTimeUnErupted
    DateTimeUnErupted = mvarDateTimeUnErupted
End Property



Public Property Let DateTimeSupernumary(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DateTimeSupernumary = 5
    mvarDateTimeSupernumary = vData
End Property


Public Property Get DateTimeSupernumary() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DateTimeSupernumary
    DateTimeSupernumary = mvarDateTimeSupernumary
End Property



Public Property Let DateTimeImplant(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DateTimeImplant = 5
    mvarDateTimeImplant = vData
End Property


Public Property Get DateTimeImplant() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DateTimeImplant
    DateTimeImplant = mvarDateTimeImplant
End Property



Public Property Let DateTimeWholeToothPlanToRemove(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DateTimeWholeToothPlanToRemove = 5
    mvarDateTimeWholeToothPlanToRemove = vData
End Property


Public Property Get DateTimeWholeToothPlanToRemove() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DateTimeWholeToothPlanToRemove
    DateTimeWholeToothPlanToRemove = mvarDateTimeWholeToothPlanToRemove
End Property



Public Property Let DateTimeWholeToothRemoved(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DateTimeWholeToothRemoved = 5
    mvarDateTimeWholeToothRemoved = vData
End Property


Public Property Get DateTimeWholeToothRemoved() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DateTimeWholeToothRemoved
    DateTimeWholeToothRemoved = mvarDateTimeWholeToothRemoved
End Property



Public Property Let DateTimePlannedCrownRemoved(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DateTimePlannedCrownRemoved = 5
    mvarDateTimePlannedCrownRemoved = vData
End Property


Public Property Get DateTimePlannedCrownRemoved() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DateTimePlannedCrownRemoved
    DateTimePlannedCrownRemoved = mvarDateTimePlannedCrownRemoved
End Property



Public Property Let DateTimePonticDenture(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DateTimePonticDenture = 5
    mvarDateTimePonticDenture = vData
End Property


Public Property Get DateTimePonticDenture() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DateTimePonticDenture
    DateTimePonticDenture = mvarDateTimePonticDenture
End Property



Public Property Let DateTimeCrownRemoved(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DateTimeCrownRemoved = 5
    mvarDateTimeCrownRemoved = vData
End Property


Public Property Get DateTimeCrownRemoved() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DateTimeCrownRemoved
    DateTimeCrownRemoved = mvarDateTimeCrownRemoved
End Property



Public Property Let DateTimePlannedRootRemoved(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DateTimePlannedRootRemoved = 5
    mvarDateTimePlannedRootRemoved = vData
End Property


Public Property Get DateTimePlannedRootRemoved() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DateTimePlannedRootRemoved
    DateTimePlannedRootRemoved = mvarDateTimePlannedRootRemoved
End Property



Public Property Let DateTimeRootRemoved(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DateTimeRootRemoved = 5
    mvarDateTimeRootRemoved = vData
End Property


Public Property Get DateTimeRootRemoved() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DateTimeRootRemoved
    DateTimeRootRemoved = mvarDateTimeRootRemoved
End Property



Public Property Let ToothID(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ToothID = 5
    mvarToothID = vData
End Property


Public Property Get ToothID() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ToothID
    ToothID = mvarToothID
End Property



Public Property Let HasDisplacement(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HasDisplacement = 5
    mvarHasDisplacement = vData
End Property


Public Property Get HasDisplacement() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HasDisplacement
    HasDisplacement = mvarHasDisplacement
End Property



Public Property Let HasPrevEndo(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HasPrevEndo = 5
    mvarHasPrevEndo = vData
End Property


Public Property Get HasPrevEndo() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HasPrevEndo
    HasPrevEndo = mvarHasPrevEndo
End Property



Public Property Let IsExtractOrMissing(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IsExtractOrMissing = 5
    mvarIsExtractOrMissing = vData
End Property


Public Property Get IsExtractOrMissing() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IsExtractOrMissing
    IsExtractOrMissing = mvarIsExtractOrMissing
End Property



Public Property Let IsImplant(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IsImplant = 5
    mvarIsImplant = vData
End Property


Public Property Get IsImplant() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IsImplant
    IsImplant = mvarIsImplant
End Property



Public Property Let RemoveLocation(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RemoveLocation = 5
    mvarRemoveLocation = vData
End Property


Public Property Get RemoveLocation() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RemoveLocation
    RemoveLocation = mvarRemoveLocation
End Property



Private Sub Class_Initialize()

    Set mvarSVGSurfaces = New CHCSII_Dental_SVG_Lib.colSVGSurfaces
    Set GCDxs = New colGCDxs
    Set GCTxs = New colGCTxs
  
End Sub


Private Sub Class_Terminate()

    Set mvarSVGSurfaces = Nothing
    Set GCDxs = Nothing
    Set GCTxs = Nothing
  
End Sub


