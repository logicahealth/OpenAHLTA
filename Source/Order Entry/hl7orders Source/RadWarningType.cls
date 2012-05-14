VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RadWarningType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarSequenceNumber As String 'local copy
Private mvarTypeCode As String 'local copy
Private mvarWarningType As String 'local copy
Private mvarOrderNumber As String 'local copy
Private mvarReportId As String 'local copy
Private mvarOrderFileId As String 'local copy
Private mvarAllergyId As String 'local copy
Private mvarErrorTypeText As String 'local copy
Private mvarRadProcedure As String 'local copy
Private mvarSectionTitle As String 'local copy
Private mvarWarningTypeCode As String 'local copy
Friend Property Let WarningTypeCode(ByVal vData As String)
    mvarWarningTypeCode = vData
End Property

Public Property Get WarningTypeCode() As String
    WarningTypeCode = mvarWarningTypeCode
End Property

Friend Property Let SectionTitle(ByVal vData As String)
    mvarSectionTitle = vData
End Property

Public Property Get SectionTitle() As String
    SectionTitle = mvarSectionTitle
End Property

Friend Property Let ConflictingOrderName(ByVal vData As String)
    mvarRadProcedure = vData
End Property

Public Property Get ConflictingOrderName() As String
    ConflictingOrderName = mvarRadProcedure
End Property

Friend Property Let ConflictData(ByVal vData As String)
    mvarErrorTypeText = vData
End Property

Public Property Get ConflictData() As String
    ConflictData = mvarErrorTypeText
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

Public Property Let OrderNumber(ByVal vData As String)
    mvarOrderNumber = vData
End Property

Public Property Get OrderNumber() As String
    OrderNumber = mvarOrderNumber
End Property

Friend Property Let WarningTypeText(ByVal vData As String)
    mvarWarningType = vData
End Property

Public Property Get WarningTypeText() As String
    WarningTypeText = mvarWarningType
End Property

Friend Property Let WarningDescription(ByVal vData As String)
    mvarTypeCode = vData
End Property

Public Property Get WarningDescription() As String
    WarningDescription = mvarTypeCode
End Property

Friend Property Let SequenceNumber(ByVal vData As String)
    mvarSequenceNumber = vData
End Property

Public Property Get SequenceNumber() As String
    SequenceNumber = mvarSequenceNumber
End Property

Public Function CopyWarning() As RadWarningType
    Set CopyWarning = New RadWarningType

    With CopyWarning
        .AllergyId = Me.AllergyId
        .ConflictData = Me.ConflictData
        .ConflictingOrderName = Me.ConflictingOrderName
        .SectionTitle = Me.SectionTitle
        .OrderFileId = Me.OrderFileId
        .ReportId = Me.ReportId
        .SequenceNumber = Me.SequenceNumber
        .WarningDescription = Me.WarningDescription
        .WarningTypeCode = Me.WarningTypeCode
        .WarningTypeText = Me.WarningTypeText
    End With

End Function



