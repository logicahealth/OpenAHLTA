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
Attribute VB_Name = "SummaryType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private szData_id As Long
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
Private szRefill As String
Private szSigNeed As Long
Private sFDBMedID As String
'76468, 76736
Private msOrderIEN As String
'77760
Private msPatientType As String

Private msParentDataID As String


Public Property Get data_id() As Long
  data_id = szData_id
End Property

Public Property Let data_id(inData_id As Long)
  szData_id = inData_id
End Property
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

Public Property Get Refill() As String
  Refill = szRemainingRF
End Property

Friend Property Let Refill(szData As String)
  szRefill = szData
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
Public Property Get SigNeed() As Long
     SigNeed = szSigNeed
End Property

Friend Property Let SigNeed(szData As Long)
     szSigNeed = szData
End Property

Public Property Get FDBMedID() As String
     FDBMedID = sFDBMedID
End Property
Friend Property Let FDBMedID(sValue As String)
     sFDBMedID = sValue
End Property
Public Property Get OrderIen() As String
     OrderIen = msOrderIEN
End Property
Friend Property Let OrderIen(sValue As String)
     msOrderIEN = sValue
End Property

Public Property Get ParentDataID() As String
     ParentDataID = msParentDataID
End Property
Friend Property Let ParentDataID(sValue As String)
     msParentDataID = sValue
End Property

Public Property Get PatientType() As String
     PatientType = msPatientType
End Property
Friend Property Let PatientType(sValue As String)
     msPatientType = sValue
End Property

