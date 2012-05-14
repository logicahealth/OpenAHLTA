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
Attribute VB_Name = "clsOrderDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private msPatientName As String
Private msPatientFMP As String
Private msPatientSSN As String
Private msPatientAge As String
Private mdPatientDOB As Date
Private msPatientID As String
Private msLocation As String
Private msBedNum As String
Private msOrderProvider As String
Private msOrderNumber As String
Private msMTFPhone As String
Private msMTFInfo As String
Private msAllergies As String
Private msMedicationName As String
Private msQuantity As String
Private msNumRefills As String
Private msSIG As String
Private msDose As String
Private msRoute As String
Private msExpandedSIG As String
Private msDilutent As String
Private msInfusionRate As String
Private mdInfusionDueDate As Date
Private msScheduleFreq As String
Private msOverrideReason As String
Private msComments As String
Private mdOrderDate As Date
Private mdFillDate As Date
Private mbDispenseStatus As Boolean
Private mdExpirationDate As Date
Private msGenericStatement As String
Private mbSign As Boolean
Private msLabelPrinter As String


Public Property Get PatientName() As String
  PatientName = msPatientName
End Property
Public Property Let PatientName(sValue As String)
  msPatientName = sValue
End Property

Public Property Get PatientFMP() As String
  PatientFMP = msPatientFMP
End Property
Public Property Let PatientFMP(sValue As String)
  msPatientFMP = sValue
End Property

Public Property Get PatientSSN() As String
  PatientSSN = msPatientSSN
End Property
Public Property Let PatientSSN(sValue As String)
  msPatientSSN = sValue
End Property

Public Property Get PatientAge() As String
  PatientAge = msPatientAge
End Property
Public Property Let PatientAge(sValue As String)
  msPatientAge = sValue
End Property

Public Property Get PatientDOB() As Date
  PatientDOB = mdPatientDOB
End Property
Public Property Let PatientDOB(dValue As Date)
  mdPatientDOB = dValue
End Property

Public Property Get PatientID() As String
  PatientID = msPatientID
End Property
Public Property Let PatientID(sValue As String)
  msPatientID = sValue
End Property

Public Property Get Location() As String
  Location = msLocation
End Property
Public Property Let Location(sValue As String)
  msLocation = sValue
End Property

Public Property Get BedNum() As String
  BedNum = msBedNum
End Property
Public Property Let BedNum(sValue As String)
  msBedNum = sValue
End Property

Public Property Get OrderProvider() As String
  OrderProvider = msOrderProvider
End Property
Public Property Let OrderProvider(sValue As String)
  msOrderProvider = sValue
End Property

Public Property Get OrderNumber() As String
  OrderNumber = msOrderNumber
End Property
Public Property Let OrderNumber(sValue As String)
  msOrderNumber = sValue
End Property

Public Property Get MTFPhone() As String
  MTFPhone = msMTFPhone
End Property
Public Property Let MTFPhone(sValue As String)
  msMTFPhone = sValue
End Property

Public Property Get MTFInfo() As String
  MTFInfo = msMTFInfo
End Property
Public Property Let MTFInfo(sValue As String)
  msMTFInfo = sValue
End Property

Public Property Get Allergies() As String
  Allergies = msAllergies
End Property
Public Property Let Allergies(sValue As String)
  msAllergies = sValue
End Property

Public Property Get MedicationName() As String
  MedicationName = msMedicationName
End Property
Public Property Let MedicationName(sValue As String)
  msMedicationName = sValue
End Property

Public Property Get Quantity() As String
  Quantity = msQuantity
End Property
Public Property Let Quantity(sValue As String)
  msQuantity = sValue
End Property

Public Property Get NumRefills() As String
  NumRefills = msNumRefills
End Property
Public Property Let NumRefills(sValue As String)
  msNumRefills = sValue
End Property

Public Property Get SIG() As String
  SIG = msSIG
End Property
Public Property Let SIG(sValue As String)
  msSIG = sValue
End Property

Public Property Get Dose() As String
  Dose = msDose
End Property
Public Property Let Dose(sValue As String)
  msDose = sValue
End Property

Public Property Get Route() As String
  Route = msRoute
End Property
Public Property Let Route(sValue As String)
  msRoute = sValue
End Property

Public Property Get ExpandedSIG() As String
  ExpandedSIG = msExpandedSIG
End Property
Public Property Let ExpandedSIG(sValue As String)
  msExpandedSIG = sValue
End Property

Public Property Get Dilutent() As String
  Dilutent = msDilutent
End Property
Public Property Let Dilutent(sValue As String)
  msDilutent = sValue
End Property

Public Property Get InfusionRate() As String
  InfusionRate = msInfusionRate
End Property
Public Property Let InfusionRate(sValue As String)
  msInfusionRate = sValue
End Property

Public Property Get InfusionDueDate() As Date
  InfusionDueDate = mdInfusionDueDate
End Property
Public Property Let InfusionDueDate(dValue As Date)
  mdInfusionDueDate = dValue
End Property

Public Property Get ScheduleFreq() As String
  ScheduleFreq = msScheduleFreq
End Property
Public Property Let ScheduleFreq(sValue As String)
  msScheduleFreq = sValue
End Property

Public Property Get OverrideReason() As String
  OverrideReason = msOverrideReason
End Property
Public Property Let OverrideReason(sValue As String)
  msOverrideReason = sValue
End Property

Public Property Get Comments() As String
  Comments = msComments
End Property
Public Property Let Comments(sValue As String)
  msComments = sValue
End Property

Public Property Get OrderDate() As Date
  OrderDate = mdOrderDate
End Property
Public Property Let OrderDate(dValue As Date)
  mdOrderDate = dValue
End Property

Public Property Get FillDate() As Date
  FillDate = mdFillDate
End Property
Public Property Let FillDate(dValue As Date)
  mdFillDate = dValue
End Property

Public Property Get DispenseStatus() As Boolean
  DispenseStatus = mbDispenseStatus
End Property
Public Property Let DispenseStatus(bValue As Boolean)
  mbDispenseStatus = bValue
End Property

Public Property Get ExpirationDate() As Date
  ExpirationDate = mdExpirationDate
End Property
Public Property Let ExpirationDate(dValue As Date)
  mdExpirationDate = dValue
End Property

Public Property Get GenericStatement() As String
  GenericStatement = msGenericStatement
End Property
Public Property Let GenericStatement(sValue As String)
  msGenericStatement = sValue
End Property

Public Property Get Sign() As Boolean
  Sign = mbSign
End Property
Public Property Let Sign(bValue As Boolean)
  mbSign = bValue
End Property

Public Property Get LabelPrinter() As String
  LabelPrinter = msLabelPrinter
End Property
Public Property Let LabelPrinter(sValue As String)
  msLabelPrinter = sValue
End Property

