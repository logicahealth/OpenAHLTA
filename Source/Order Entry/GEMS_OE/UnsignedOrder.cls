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
Attribute VB_Name = "UnsignedOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"UnsignedAddActions"
Attribute VB_Ext_KEY = "Member1" ,"UnsignedOrders"
Attribute VB_Ext_KEY = "Member2" ,"UnsignedOrders"
Option Explicit

'Implements ICHCSII_UnsgnOrder

Private mvarOrderNumber As String 'local copy
'local variable(s) to hold property value(s)
Private mvarOrderType As HL7Orders.hlORDERTYPECODE 'local copy
'local variable(s) to hold property value(s)
Private mvarModifiable As hlMODIFIABLE 'local copy
Private mvarszDisplayText As String 'local copy
Private mvarszBriefDisplayText As String 'local copy

Private mvarszStatus As String 'local copy
Private mvarszExpirationDate As String 'local copy
Private mvarszDispenseStatus As String 'local copy
Private mvarszRefillsRemaining As String 'local copy
Private mvarszLastFillDate As String 'local copy
Private mvarszOnHoldUntilDate As String 'local copy
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
Private mvarOrderStartDateTime As String 'local copy
'Private mvarAdditionalOrderActions As UnsignedAddActions
'Private mvarUnsignedOrders As UnsignedOrders
Private mvarUnsignedAddActions As UnsignedAddActions
Private mvarUserHasSigned As Boolean
Private mvarszOrderComment As String 'local copy
Private mvarCompletionStatus As hlSIGNATURECOMPLETION
Private mvarWarningCollection As Collection
Private mvarWarningHeader As String
Private mvarWarningReason As String
Private mvarWarningsProcessed As Boolean
Private mvarSigningSuccessful As Boolean
Private mvarDisplayedInListView As Boolean
Private mvarPatientInfo As PatientType


Public Property Get UnsignedAddActions() As UnsignedAddActions
    If mvarUnsignedAddActions Is Nothing Then
        Set mvarUnsignedAddActions = New UnsignedAddActions
    End If

    Set UnsignedAddActions = mvarUnsignedAddActions
End Property


Public Property Set UnsignedAddActions(vData As UnsignedAddActions)
    Set mvarUnsignedAddActions = vData
End Property

Public Property Get WarningCollection() As Collection
    If mvarWarningCollection Is Nothing Then
        Set mvarWarningCollection = New Collection
    End If
    Set WarningCollection = mvarWarningCollection
End Property


Public Property Set WarningCollection(vData As Collection)
    Set mvarWarningCollection = vData
End Property

Private Sub Class_Initialize()
    'create the mSignatureInfo object when the UnsignedOrder class is created
    Set mvarProviderSignatureInfo = New SignatureInfo
    Set mvarNurseSignatureInfo = New SignatureInfo
    Set mvarCounterSignatureInfo = New SignatureInfo
    Set mvarAncillarySignatureInfo = New SignatureInfo
    Set mvarUnsignedAddActions = New UnsignedAddActions
    mvarUserHasSigned = False
    mvarWarningsProcessed = False
    mvarSigningSuccessful = False
    mvarDisplayedInListView = False
    Set mvarPatientInfo = New PatientType
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
    'Set mvarUnsignedOrders = Nothing
    Set mvarProviderSignatureInfo = Nothing
    Set mvarNurseSignatureInfo = Nothing
    Set mvarCounterSignatureInfo = Nothing
    Set mvarAncillarySignatureInfo = Nothing
    Set mvarPatientInfo = Nothing
End Sub



Public Property Let AncillarySignatureNeeded(ByVal vData As hlSIGNATURENEEDED)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.AncillarySignatureNeeded = Form1
    mvarAncillarySignatureNeeded = vData
End Property


Public Property Get AncillarySignatureNeeded() As hlSIGNATURENEEDED
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AncillarySignatureNeeded
    AncillarySignatureNeeded = mvarAncillarySignatureNeeded
End Property



Public Property Let CounterSignatureNeeded(ByVal vData As hlSIGNATURENEEDED)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.CounterSignatureNeeded = Form1
    mvarCounterSignatureNeeded = vData
End Property


Public Property Get CounterSignatureNeeded() As hlSIGNATURENEEDED
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CounterSignatureNeeded
    CounterSignatureNeeded = mvarCounterSignatureNeeded
End Property



Public Property Let ProviderSignatureNeeded(ByVal vData As hlSIGNATURENEEDED)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.ProviderSignatureNeeded = Form1
    mvarProviderSignatureNeeded = vData
End Property


Public Property Get ProviderSignatureNeeded() As hlSIGNATURENEEDED
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ProviderSignatureNeeded
    ProviderSignatureNeeded = mvarProviderSignatureNeeded
End Property



Public Property Let NurseSignatureNeeded(ByVal vData As hlSIGNATURENEEDED)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.NurseSignatureNeeded = Form1
    mvarNurseSignatureNeeded = vData
End Property


Public Property Get NurseSignatureNeeded() As hlSIGNATURENEEDED
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NurseSignatureNeeded
    NurseSignatureNeeded = mvarNurseSignatureNeeded
End Property



Public Property Let szOrderEventDate(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.szOrderEventDate = 5
    mvarszOrderEventDate = vData
End Property


Public Property Get szOrderEventDate() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.szOrderEventDate
    szOrderEventDate = mvarszOrderEventDate
End Property
Public Property Let OrderStartDateTime(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.szOrderEventDate = 5
    mvarOrderStartDateTime = vData
End Property


Public Property Get OrderStartDateTime() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.szOrderEventDate
    OrderStartDateTime = mvarOrderStartDateTime
End Property



Public Property Let SigningAction(ByVal vData As hlORDERSIGNACTION)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.SigningAction = Form1
    mvarSigningAction = vData
End Property


Public Property Get SigningAction() As hlORDERSIGNACTION
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SigningAction
    SigningAction = mvarSigningAction
End Property



Public Property Let szOrderingProviderName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.szOrderingProviderName = 5
    mvarszOrderingProviderName = vData
End Property


Public Property Get szOrderingProviderName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.szOrderingProviderName
    szOrderingProviderName = mvarszOrderingProviderName
End Property



Public Property Let szOrderingProviderIEN(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.szOrderingProviderIEN = 5
    mvarszOrderingProviderIEN = vData
End Property


Public Property Get szOrderingProviderIEN() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.szOrderingProviderIEN
    szOrderingProviderIEN = mvarszOrderingProviderIEN
End Property



Public Property Let szEnteredByName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.szEnteredByName = 5
    mvarszEnteredByName = vData
End Property


Public Property Get szEnteredByName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.szEnteredByName
    szEnteredByName = mvarszEnteredByName
End Property



Public Property Let szEnteredByIEN(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.szEnteredByIEN = 5
    mvarszEnteredByIEN = vData
End Property


Public Property Get szEnteredByIEN() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.szEnteredByIEN
    szEnteredByIEN = mvarszEnteredByIEN
End Property



Public Property Let szOnHoldUntilDate(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.szOnHoldUntilDate = 5
    mvarszOnHoldUntilDate = vData
End Property


Public Property Get szOnHoldUntilDate() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.szOnHoldUntilDate
    szOnHoldUntilDate = mvarszOnHoldUntilDate
End Property



Public Property Let szLastFillDate(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.szLastFillDate = 5
    mvarszLastFillDate = vData
End Property


Public Property Get szLastFillDate() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.szLastFillDate
    szLastFillDate = mvarszLastFillDate
End Property



Public Property Let szRefillsRemaining(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.szRefillsRemaining = 5
    mvarszRefillsRemaining = vData
End Property


Public Property Get szRefillsRemaining() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.szRefillsRemaining
    szRefillsRemaining = mvarszRefillsRemaining
End Property



Public Property Let szDispenseStatus(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.szDispenseStatus = 5
    mvarszDispenseStatus = vData
End Property


Public Property Get szDispenseStatus() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.szDispenseStatus
    szDispenseStatus = mvarszDispenseStatus
End Property



Public Property Let szExpirationDate(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.szExpirationDate = 5
    mvarszExpirationDate = vData
End Property


Public Property Get szExpirationDate() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.szExpirationDate
    szExpirationDate = mvarszExpirationDate
End Property



Public Property Let szStatus(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.szStatus = 5
    mvarszStatus = vData
End Property


Public Property Get szStatus() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.szStatus
    szStatus = mvarszStatus
End Property



Public Property Let szDisplayText(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.szDisplayText = 5
    mvarszDisplayText = vData
End Property


Public Property Get szDisplayText() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.szDisplayText
    szDisplayText = mvarszDisplayText
End Property
Public Property Let szBriefDisplayText(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.szDisplayText = 5
    mvarszBriefDisplayText = vData
End Property


Public Property Get szBriefDisplayText() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.szDisplayText
    szBriefDisplayText = mvarszBriefDisplayText
End Property



Public Property Let Modifiable(ByVal vData As hlMODIFIABLE)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bModifiable = 5
    mvarModifiable = vData
End Property


Public Property Get Modifiable() As hlMODIFIABLE
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bModifiable
    Modifiable = mvarModifiable
End Property



Public Property Let OrderType(ByVal vData As HL7Orders.hlORDERTYPECODE)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.OrderType = Form1
    mvarOrderType = vData
End Property
Public Property Get OrderType() As HL7Orders.hlORDERTYPECODE
    OrderType = mvarOrderType
End Property
Public Property Let OrderNumber(ByVal vData As String)
    mvarOrderNumber = vData
End Property
Public Property Get OrderNumber() As String
    OrderNumber = mvarOrderNumber
End Property

Public Property Let UserHasSigned(ByVal vData As Boolean)
    mvarUserHasSigned = vData
End Property
Public Property Get UserHasSigned() As Boolean
    UserHasSigned = mvarUserHasSigned
End Property

Public Property Let SigningSuccessful(ByVal vData As Boolean)
    mvarSigningSuccessful = vData
End Property
Public Property Get SigningSuccessful() As Boolean
    SigningSuccessful = mvarSigningSuccessful
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

Public Property Let WarningHeader(ByVal vData As String)
    mvarWarningHeader = vData
End Property
Public Property Get WarningHeader() As String
    WarningHeader = mvarWarningHeader
End Property
Public Property Let WarningReason(ByVal vData As String)
    mvarWarningReason = vData
End Property
Public Property Get WarningReason() As String
    WarningReason = mvarWarningReason
End Property
Public Property Let WarningsProcessed(ByVal vData As Boolean)
    mvarWarningsProcessed = vData
End Property
Public Property Get WarningsProcessed() As Boolean
    WarningsProcessed = mvarWarningsProcessed
End Property
Public Property Let DisplayedInListView(ByVal vData As Boolean)
    mvarDisplayedInListView = vData
End Property
Public Property Get DisplayedInListView() As Boolean
    DisplayedInListView = mvarDisplayedInListView
End Property
Public Property Set PatientInfo(ByVal vData As PatientType)
    Set mvarPatientInfo = vData
End Property

Public Property Get PatientInfo() As PatientType
    Set PatientInfo = mvarPatientInfo
End Property

