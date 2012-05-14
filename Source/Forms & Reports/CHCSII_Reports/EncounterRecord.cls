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
Attribute VB_Name = "EncounterRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private sDataId As String
Private sProviderNCID As String
Private sPatientId As String
Private sEventDate As String
Private sClinicNCID As String
Private sHeaderId As String
Private sAppointmentNCID As String
Private sEandMNCID As String
Private sDispositionNCID As String
Private sCategoryNCID As String
Private sPrimaryDiagnosisNCID As String
Private sMTFNCID As String
Private sEncounterNumber As String

Public Property Get DataId() As String
    DataId = sDataId
End Property

Public Property Let DataId(ByVal sNewValue As String)
    sDataId = sNewValue
End Property

Private Sub Class_Initialize()
    sHeaderId = "0"
End Sub

Public Property Get ProviderNCID() As String
    ProviderNCID = sProviderNCID
End Property

Public Property Let ProviderNCID(ByVal sNewValue As String)
    sProviderNCID = sNewValue
End Property

Public Property Get PatientId() As String
    PatientId = sPatientId
End Property

Public Property Let PatientId(ByVal sNewValue As String)
    sPatientId = sNewValue
End Property

Public Property Get EventDate() As String
    EventDate = sEventDate
End Property

Public Property Let EventDate(ByVal sNewValue As String)
    sEventDate = sNewValue
End Property

Public Property Get ClinicNCID() As String
    ClinicNCID = sClinicNCID
End Property

Public Property Let ClinicNCID(ByVal sNewValue As String)
    sClinicNCID = sNewValue
End Property

Public Property Get HeaderId() As String
    HeaderId = sHeaderId
End Property

Public Property Let HeaderId(ByVal sNewValue As String)
    sHeaderId = sNewValue
End Property

Public Property Get AppointmentNCID() As String
    AppointmentNCID = sAppointmentNCID
End Property

Public Property Let AppointmentNCID(ByVal sNewValue As String)
    sAppointmentNCID = sNewValue
End Property

Public Property Get EandMNCID() As String
    EandMNCID = sEandMNCID
End Property

Public Property Let EandMNCID(ByVal sNewValue As String)
    sEandMNCID = sNewValue
End Property

Public Property Get DispositionNCID() As String
    DispositionNCID = sDispositionNCID
End Property

Public Property Let DispositionNCID(ByVal sNewValue As String)
    sDispositionNCID = sNewValue
End Property

Public Property Get CategoryNCID() As String
    CategoryNCID = sCategoryNCID
End Property

Public Property Let CategoryNCID(ByVal sNewValue As String)
    sCategoryNCID = sNewValue
End Property

Public Property Get PrimaryDiagnosisNCID() As String
    PrimaryDiagnosisNCID = sPrimaryDiagnosisNCID
End Property

Public Property Let PrimaryDiagnosisNCID(ByVal sNewValue As String)
    sPrimaryDiagnosisNCID = sNewValue
End Property

Public Property Get MTF_NCID() As String
    MTF_NCID = sMTFNCID
End Property

Public Property Let MTF_NCID(ByVal sNewValue As String)
    sMTFNCID = sNewValue
End Property
Public Property Get EncounterNumber() As String
    EncounterNumber = sEncounterNumber
End Property

Public Property Let EncounterNumber(ByVal sNewValue As String)
    sEncounterNumber = sNewValue
End Property
