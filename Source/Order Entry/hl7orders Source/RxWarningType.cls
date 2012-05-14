VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RxWarningType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private mvarConflictOrderText As String 'local copy
Private mvarSequenceNumber As String 'local copy
Private mvarTypeCode As String 'local copy
Private mvarDescription As String 'local copy
Private mvarWarningType As String 'local copy
Private mvarDrugName As String 'local copy
Private mvarRxType As String 'local copy
Private mvarOrderNumber As String 'local copy
Private mvarOrderSig As String 'local copy
Private mvarReportId As String 'local copy
Private mvarOrderFileId As String 'local copy
Private mvarAllergyId As String 'local copy
Private mvarWarningTypeCode As String 'local copy
Private mvarWarningTypeCodePDTS As hlPDTSWARNING 'local copy
Private mvarConflictsWithOrWhy As String 'local copy
Private mvarPDTSExpandedWarning As String
Public Property Let WarningTypeCodePDTS(ByVal vData As hlPDTSWARNING)
    mvarWarningTypeCodePDTS = vData
End Property

Public Property Get WarningTypeCodePDTS() As hlPDTSWARNING
    WarningTypeCodePDTS = mvarWarningTypeCodePDTS
End Property

Friend Property Let ConflictsWithOrWhy(ByVal vData As String)
    mvarConflictsWithOrWhy = vData
End Property

Public Property Get ConflictsWithOrWhy() As String
    ConflictsWithOrWhy = mvarConflictsWithOrWhy
End Property

Friend Property Let OrderFileId(ByVal vData As String)
    mvarOrderFileId = vData
End Property

Public Property Get OrderFileId() As String
    OrderFileId = mvarOrderFileId
End Property

Friend Property Let WarningTypeCode(ByVal vData As String)
    mvarWarningTypeCode = vData
End Property

Public Property Get WarningTypeCode() As String
    WarningTypeCode = mvarWarningTypeCode
End Property

Friend Property Let AllergyId(ByVal vData As String)
    mvarAllergyId = vData
End Property

Public Property Get AllergyId() As String
    AllergyId = mvarAllergyId
End Property

Friend Property Let ReportId(ByVal vData As String)
    mvarReportId = vData
End Property

Public Property Get ReportId() As String
    ReportId = mvarReportId
End Property

Friend Property Let ConflictSigAndDate(ByVal vData As String)
    mvarOrderSig = vData
End Property

Public Property Get ConflictSigAndDate() As String
    ConflictSigAndDate = mvarOrderSig
End Property

Friend Property Let OrderNumber(ByVal vData As String)
    mvarOrderNumber = vData
End Property

Public Property Get OrderNumber() As String
    OrderNumber = mvarOrderNumber
End Property

Friend Property Let ConflictData(ByVal vData As String)
    mvarRxType = vData
End Property

Public Property Get ConflictData() As String
    ConflictData = mvarRxType
End Property

Friend Property Let ConflictingOrderName(ByVal vData As String)
    mvarDrugName = vData
End Property

Public Property Get ConflictingOrderName() As String
    ConflictingOrderName = mvarDrugName
End Property

Friend Property Let WarningTypeText(ByVal vData As String)
    mvarWarningType = vData
End Property

Public Property Get WarningTypeText() As String
    WarningTypeText = mvarWarningType
End Property

Friend Property Let WarningDescription(ByVal vData As String)
    mvarDescription = vData
End Property

Public Property Get WarningDescription() As String
    WarningDescription = mvarDescription
End Property

Friend Property Let ReasonOrProblem(ByVal vData As String)
    mvarTypeCode = vData
End Property

Public Property Get ReasonOrProblem() As String
    ReasonOrProblem = mvarTypeCode
End Property

Friend Property Let SequenceNumber(ByVal vData As String)
    mvarSequenceNumber = vData
End Property

Public Property Get SequenceNumber() As String
    SequenceNumber = mvarSequenceNumber
End Property

Friend Property Let ConflictingOrderText(ByVal vData As String)
    mvarConflictOrderText = vData
End Property

Public Property Get ConflictingOrderText() As String
    ConflictingOrderText = mvarConflictOrderText
End Property
Friend Property Let PDTSExpandedWarning(ByVal vData As String)
    mvarPDTSExpandedWarning = vData
End Property

Public Property Get PDTSExpandedWarning() As String
    PDTSExpandedWarning = mvarPDTSExpandedWarning
End Property

Public Function CopyWarning() As RxWarningType
    Set CopyWarning = New RxWarningType

    With CopyWarning
        .AllergyId = Me.AllergyId
        .ConflictData = Me.ConflictData
        .ConflictingOrderName = Me.ConflictingOrderName
        .ConflictSigAndDate = Me.ConflictSigAndDate
        .ConflictsWithOrWhy = Me.ConflictsWithOrWhy
        .OrderFileId = Me.OrderFileId
        .ReportId = Me.ReportId
        .SequenceNumber = Me.SequenceNumber
        .WarningDescription = Me.WarningDescription
        .WarningTypeCode = Me.WarningTypeCode
        .WarningTypeText = Me.WarningTypeText
        .ReasonOrProblem = Me.ReasonOrProblem
        .OrderNumber = Me.OrderNumber
    End With

End Function

