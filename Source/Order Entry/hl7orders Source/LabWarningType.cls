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
Attribute VB_Name = "LabWarningType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private mvarSequenceNumber As String 'local copy
Private mvarWarningType As String 'local copy
Private mvarOrderNumber As String 'local copy
Private mvarReportId As String 'local copy
Private mvarOrderFileId As String 'local copy
Private mvarAllergyId As String 'local copy
Private mvarErrorTypeText As String 'local copy
Private mvarConflictingLabTest As String 'local copy
Private mvarLabCollectMethod As String 'local copy
Private mvarLabSampleType As String 'local copy
Private mvarCapColor As String 'local copy
Private mvarWarningTypeText As String 'local copy

Friend Property Let WarningTypeText(ByVal vData As String)
    mvarWarningTypeText = vData
End Property

Public Property Get WarningTypeText() As String
    WarningTypeText = mvarWarningTypeText
End Property

Friend Property Let CapColorAndTiming(ByVal vData As String)
    mvarCapColor = vData
End Property

Public Property Get CapColorAndTiming() As String
    CapColorAndTiming = mvarCapColor
End Property

Friend Property Let LabSampleType(ByVal vData As String)
    mvarLabSampleType = vData
End Property

Public Property Get LabSampleType() As String
    LabSampleType = mvarLabSampleType
End Property

Friend Property Let LabCollectMethod(ByVal vData As String)
    mvarLabCollectMethod = vData
End Property

Public Property Get LabCollectMethod() As String
    LabCollectMethod = mvarLabCollectMethod
End Property

Friend Property Let ConflictingOrderName(ByVal vData As String)
    mvarConflictingLabTest = vData
End Property

Public Property Get ConflictingOrderName() As String
    ConflictingOrderName = mvarConflictingLabTest
End Property

Friend Property Let WarningDescription(ByVal vData As String)
    mvarErrorTypeText = vData
End Property

Public Property Get WarningDescription() As String
    WarningDescription = mvarErrorTypeText
End Property

Friend Property Let AllergyId(ByVal vData As String)
    mvarAllergyId = vData
End Property

Public Property Get AllergyId() As String
    AllergyId = mvarAllergyId
End Property

Friend Property Let OrderFileId(ByVal vData As String)
    mvarOrderFileId = vData
End Property

Public Property Get OrderFileId() As String
    OrderFileId = mvarOrderFileId
End Property

Friend Property Let ReportId(ByVal vData As String)
    mvarReportId = vData
End Property

Public Property Get ReportId() As String
    ReportId = mvarReportId
End Property

Friend Property Let ConflictData(ByVal vData As String)
    mvarOrderNumber = vData
End Property

Public Property Get ConflictData() As String
    ConflictData = mvarOrderNumber
End Property

Friend Property Let WarningTypeCode(ByVal vData As String)
    mvarWarningType = vData
End Property

Public Property Get WarningTypeCode() As String
    WarningTypeCode = mvarWarningType
End Property

Friend Property Let SequenceNumber(ByVal vData As String)
    mvarSequenceNumber = vData
End Property

Public Property Get SequenceNumber() As String
    SequenceNumber = mvarSequenceNumber
End Property

Public Function CopyWarning() As LabWarningType
    Set CopyWarning = New LabWarningType

    With CopyWarning
        .AllergyId = Me.AllergyId
        .CapColorAndTiming = Me.CapColorAndTiming
        .ConflictData = Me.ConflictData
        .ConflictingOrderName = Me.ConflictingOrderName
        .LabCollectMethod = Me.LabCollectMethod
        .LabSampleType = Me.LabSampleType
        .OrderFileId = Me.OrderFileId
        .ReportId = Me.ReportId
        .SequenceNumber = Me.SequenceNumber
        .WarningDescription = Me.WarningDescription
        .WarningTypeCode = Me.WarningTypeCode
        .WarningTypeText = Me.WarningTypeText
    End With

End Function

