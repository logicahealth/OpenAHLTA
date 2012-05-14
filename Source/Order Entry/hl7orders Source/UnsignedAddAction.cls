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
Attribute VB_Name = "UnsignedAddAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
Private mvarOrderNumber As String 'local copy
Private mvarszEnteredByIEN As String 'local copy
Private mvarszEnteredByName As String 'local copy
Private mvarszOrderingProviderIEN As String 'local copy
Private mvarszOrderingProviderName As String 'local copy
Private mvarSigningAction As hlORDERSIGNACTION 'local copy
Private mvarszOrderEventDate As String 'local copy
Private mvarszOrderEventDateHL7Format As String 'local copy

Private mvarModToIVRate As String 'local copy
Private mvarModToIVDate As String 'local copy
Private mvarNurseSignatureNeeded As hlSIGNATURENEEDED 'local copy
Private mvarProviderSignatureNeeded As hlSIGNATURENEEDED 'local copy
Private mvarCounterSignatureNeeded As hlSIGNATURENEEDED 'local copy
Private mvarAncillarySignatureNeeded As hlSIGNATURENEEDED 'local copy
Private mvarProviderSignatureInfo As SignatureInfo
Private mvarNurseSignatureInfo As SignatureInfo
Private mvarCounterSignatureInfo As SignatureInfo
Private mvarAncillarySignatureInfo As SignatureInfo
Private mvarszActionComment As String
Private mvarUserHasSigned As Boolean


'Private mvarUnsignedAddOrders As UnsignedAddActions
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
'Public Property Get AncillarySignatureInfo() As SignatureInfo
'    Set AncillarySignatureInfo = mvarAncillarySignatureInfo
'End Property
'Public Property Set AncillarySignatureInfo(vData As SignatureInfo)
'    Set mvarAncillarySignatureInfo = vData
'End Property
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

Public Property Let ModToIVRate(ByVal vData As String)
    mvarModToIVRate = vData
End Property
Public Property Get ModToIVRate() As String
    ModToIVRate = mvarModToIVRate
End Property
Public Property Let ModToIVDate(ByVal vData As String)
    mvarModToIVDate = vData
End Property
Public Property Get ModToIVDate() As String
    ModToIVDate = mvarModToIVDate
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

Public Property Let OrderNumber(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OrderNumber = 5
    mvarOrderNumber = vData
End Property


Public Property Get OrderNumber() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OrderNumber
    OrderNumber = mvarOrderNumber
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

Private Sub Class_Initialize()
    Set mvarProviderSignatureInfo = New SignatureInfo
    Set mvarNurseSignatureInfo = New SignatureInfo
    Set mvarCounterSignatureInfo = New SignatureInfo
    Set mvarAncillarySignatureInfo = New SignatureInfo
'    Set mvarUnsignedAddOrders = New UnsignedAddAction
    mvarUserHasSigned = False
End Sub

Private Sub Class_Terminate()
    Set mvarProviderSignatureInfo = Nothing
    Set mvarNurseSignatureInfo = Nothing
    Set mvarCounterSignatureInfo = Nothing
    Set mvarAncillarySignatureInfo = Nothing
'    Set mvarUnsignedAddOrders = Nothing
End Sub
Public Property Let szOrderEventDateHL7Format(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.szOrderEventDate = 5
    mvarszOrderEventDateHL7Format = vData
End Property


Public Property Get szOrderEventDateHL7Format() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.szOrderEventDate
    szOrderEventDateHL7Format = mvarszOrderEventDateHL7Format
End Property
Public Property Let ActionComment(ByVal vData As String)
    mvarszActionComment = vData
End Property
Public Property Get ActionComment() As String
    ActionComment = mvarszActionComment
End Property
Public Property Let UserHasSigned(ByVal vData As Boolean)
    mvarUserHasSigned = vData
End Property
Public Property Get UserHasSigned() As Boolean
    UserHasSigned = mvarUserHasSigned
End Property
