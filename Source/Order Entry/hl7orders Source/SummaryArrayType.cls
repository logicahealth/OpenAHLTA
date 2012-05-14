VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SummaryType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private szOrderId As String
Private mhlOrderTypeCode As hlORDERTYPECODE
Private szExpirationDate As String
Private szFullText As String
Private szStatus As String
Private szBriefText As String
Private szOrderTypeText As String
Private lModifiable As hlMODIFIABLE
Private szLastDispensed As String
Private mszLastFillDate As String
Private szRemainingRF As String
'new order signing information
Private mvarszEnteredByIEN As String 'local copy
Private mvarszEnteredByName As String 'local copy
Private mvarszOrderingProviderIEN As String 'local copy
Private mvarszOrderingProviderName As String 'local copy
Private mvarSigningAction As hlORDERSIGNACTION 'local copy
Private mvarszOrderEventDate As String 'local copy
Private mvarszOrderEventDateHL7Format As String 'local copy

Private mvarNurseSignatureNeeded As hlSIGNATURENEEDED 'local copy
Private mvarProviderSignatureNeeded As hlSIGNATURENEEDED 'local copy
Private mvarCounterSignatureNeeded As hlSIGNATURENEEDED 'local copy
Private mvarAncillarySignatureNeeded As hlSIGNATURENEEDED 'local copy
Private mvarProviderSignatureInfo As SignatureInfo
Private mvarNurseSignatureInfo As SignatureInfo
Private mvarCounterSignatureInfo As SignatureInfo
Private mvarAncillarySignatureInfo As SignatureInfo
Private mvarPatientInfo As PatientType 'local copy
Private mvarOrderStartDateTime As String 'local copy
Private mvarUnsignedAddActions As UnsignedAddActions
Private mvarUserHasSigned As Boolean
Private mvarszOrderComment As String 'local copy
Private mvarCompletionStatus As hlSIGNATURECOMPLETION


Public Property Get LastFillDate() As String
  LastFillDate = mszLastFillDate
End Property

Friend Property Let LastFillDate(szData As String)
    On Error GoTo ErrorHandler:

     If IsDate(szData) Then
        mszLastFillDate = Format$(szData, "DD-MMM-YYYY")
     ElseIf Len(szData) = 0 Then
        mszLastFillDate = ""
    Else
        'If we get here (normally we should), then we are assuming that date is in format
        ' as given in ZOD msg i.e. 20000526000000
        If Len(szData) = 14 Then
            Dim szYear As String, szMonth As String, szDay As String
            szYear = Left$(szData, 4)
            szMonth = Mid$(szData, 5, 2)
            szDay = Mid$(szData, 7, 2)
            szData = szMonth & "/" & szDay & "/" & szYear
            szData = Format$(szData, "DD-MMM-YYYY")
                mszLastFillDate = szData
        Else
            mszLastFillDate = ""
        End If
     End If
     Exit Property
ErrorHandler:
     'mszLastFillDate = szData
End Property

Public Property Get RemainingRF() As String
  RemainingRF = szRemainingRF
End Property

Friend Property Let RemainingRF(szData As String)
  szRemainingRF = szData
End Property

Public Property Get OrderId() As String
  OrderId = szOrderId
End Property

Friend Property Let OrderId(szData As String)
  szOrderId = szData
End Property

Public Property Get OrderTypeCode() As hlORDERTYPECODE
  OrderTypeCode = mhlOrderTypeCode
End Property

Friend Property Let OrderTypeCode(szData As hlORDERTYPECODE)
  mhlOrderTypeCode = szData
End Property

Public Property Get ExpirationDate() As String
  ExpirationDate = szExpirationDate
End Property

Friend Property Let ExpirationDate(szData As String)
  szExpirationDate = szData
End Property

Public Property Get FullText() As String
  FullText = szFullText
End Property

Friend Property Let FullText(szData As String)
  szFullText = szData
End Property

Public Property Get Status() As String
  Status = szStatus
End Property
Friend Property Let Status(szData As String)
  szStatus = szData
End Property

Public Property Get BriefText() As String
  BriefText = szBriefText
End Property

Friend Property Let BriefText(szData As String)
  szBriefText = szData
End Property

Public Property Get OrderTypeText() As String
  OrderTypeText = szOrderTypeText
End Property

Friend Property Let OrderTypeText(szData As String)
  szOrderTypeText = szData
End Property

Public Property Get Modifiable() As hlMODIFIABLE
  Modifiable = lModifiable
End Property

Friend Property Let Modifiable(lData As hlMODIFIABLE)
  lModifiable = lData
End Property

Public Property Get LastDispenseDate() As String
     LastDispenseDate = szLastDispensed
End Property

Friend Property Let LastDispenseDate(szData As String)
     szLastDispensed = szData
End Property

Private Sub Class_Initialize()
    'create the mSignatureInfo object when the UnsignedOrder class is created
    Set mvarProviderSignatureInfo = New SignatureInfo
    Set mvarNurseSignatureInfo = New SignatureInfo
    Set mvarCounterSignatureInfo = New SignatureInfo
    Set mvarAncillarySignatureInfo = New SignatureInfo
    Set mvarUnsignedAddActions = New UnsignedAddActions
    mvarUserHasSigned = False

End Sub
Public Property Get ProviderSignatureInfo() As SignatureInfo
    Set ProviderSignatureInfo = mvarProviderSignatureInfo
End Property
Public Property Set ProviderSignatureInfo(vData As SignatureInfo)
    Set mvarProviderSignatureInfo = vData
End Property
Public Property Get NurseSignatureInfo() As SignatureInfo
    Set NurseSignatureInfo = mvarNurseSignatureInfo
End Property
Public Property Set NurseSignatureInfo(vData As SignatureInfo)
    Set mvarNurseSignatureInfo = vData
End Property
Public Property Get CounterSignatureInfo() As SignatureInfo
    Set CounterSignatureInfo = mvarCounterSignatureInfo
End Property
Public Property Set CounterSignatureInfo(vData As SignatureInfo)
    Set mvarCounterSignatureInfo = vData
End Property
Public Property Get AncillarySignatureInfo() As SignatureInfo
    Set AncillarySignatureInfo = mvarAncillarySignatureInfo
End Property
Public Property Set AncillarySignatureInfo(vData As SignatureInfo)
    Set mvarAncillarySignatureInfo = vData
End Property

Private Sub Class_Terminate()
    Set mvarUnsignedAddActions = Nothing
    Set mvarProviderSignatureInfo = Nothing
    Set mvarNurseSignatureInfo = Nothing
    Set mvarCounterSignatureInfo = Nothing
    Set mvarAncillarySignatureInfo = Nothing
    
End Sub
Public Property Let AncillarySignatureNeeded(ByVal vData As hlSIGNATURENEEDED)
    mvarAncillarySignatureNeeded = vData
End Property

Public Property Get AncillarySignatureNeeded() As hlSIGNATURENEEDED
    AncillarySignatureNeeded = mvarAncillarySignatureNeeded
End Property
Public Property Let CounterSignatureNeeded(ByVal vData As hlSIGNATURENEEDED)
    mvarCounterSignatureNeeded = vData
End Property
Public Property Get CounterSignatureNeeded() As hlSIGNATURENEEDED
    CounterSignatureNeeded = mvarCounterSignatureNeeded
End Property
Public Property Let ProviderSignatureNeeded(ByVal vData As hlSIGNATURENEEDED)
    mvarProviderSignatureNeeded = vData
End Property
Public Property Get ProviderSignatureNeeded() As hlSIGNATURENEEDED
    ProviderSignatureNeeded = mvarProviderSignatureNeeded
End Property

Public Property Let NurseSignatureNeeded(ByVal vData As hlSIGNATURENEEDED)
    mvarNurseSignatureNeeded = vData
End Property
Public Property Get NurseSignatureNeeded() As hlSIGNATURENEEDED
    NurseSignatureNeeded = mvarNurseSignatureNeeded
End Property
Public Property Let szOrderEventDate(ByVal vData As String)
    mvarszOrderEventDate = vData
End Property
Public Property Get szOrderEventDate() As String
    szOrderEventDate = mvarszOrderEventDate
End Property

Public Property Let SigningAction(ByVal vData As hlORDERSIGNACTION)
    mvarSigningAction = vData
End Property
Public Property Get SigningAction() As hlORDERSIGNACTION
    SigningAction = mvarSigningAction
End Property
Public Property Let szOrderingProviderName(ByVal vData As String)
    mvarszOrderingProviderName = vData
End Property

Public Property Get szOrderingProviderName() As String
    szOrderingProviderName = mvarszOrderingProviderName
End Property
Public Property Let szOrderingProviderIEN(ByVal vData As String)
    mvarszOrderingProviderIEN = vData
End Property

Public Property Get szOrderingProviderIEN() As String
    szOrderingProviderIEN = mvarszOrderingProviderIEN
End Property
Public Property Let szEnteredByName(ByVal vData As String)
    mvarszEnteredByName = vData
End Property


Public Property Get szEnteredByName() As String
    szEnteredByName = mvarszEnteredByName
End Property

Public Property Let szEnteredByIEN(ByVal vData As String)
    mvarszEnteredByIEN = vData
End Property

Public Property Get szEnteredByIEN() As String
    szEnteredByIEN = mvarszEnteredByIEN
End Property
Public Property Let UserHasSigned(ByVal vData As Boolean)
    mvarUserHasSigned = vData
End Property
Public Property Get UserHasSigned() As Boolean
    UserHasSigned = mvarUserHasSigned
End Property

Public Property Let szOrderEventDateHL7Format(ByVal vData As String)
    mvarszOrderEventDateHL7Format = vData
End Property
Public Property Get szOrderEventDateHL7Format() As String
    szOrderEventDateHL7Format = mvarszOrderEventDateHL7Format
End Property
Public Property Let OrderComment(ByVal vData As String)
    mvarszOrderComment = vData
End Property
Public Property Get OrderComment() As String
    OrderComment = mvarszOrderComment
End Property
Public Property Let CompletionStatus(ByVal vData As hlSIGNATURECOMPLETION)
    mvarCompletionStatus = vData
End Property
Public Property Get CompletionStatus() As hlSIGNATURECOMPLETION
    CompletionStatus = mvarCompletionStatus
End Property
Public Property Let OrderStartDateTime(ByVal vData As String)
    mvarOrderStartDateTime = vData
End Property
Public Property Get OrderStartDateTime() As String
    OrderStartDateTime = mvarOrderStartDateTime
End Property


