VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Clinic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Name As String           ' Name of Clinic
Private FName As String          ' Name of Facility
Private FacilityNCID As String   ' Facility Ncid
Private CNCID As String            ' Clinic Ncid


Public Property Get clinicName() As String
 clinicName = Name
End Property

Public Property Let clinicName(ByVal vNewValue As String)
 Name = vNewValue
End Property

Public Property Get FacNcid() As String
  FacNcid = FacilityNCID
End Property

Public Property Let FacNcid(ByVal vNewValue As String)
  FacilityNCID = vNewValue
End Property
Public Property Get ClinicNCID() As String
 ClinicNCID = CNCID
End Property

Public Property Let ClinicNCID(ByVal vNewValue As String)
 CNCID = vNewValue
End Property
Public Property Get FacilityName() As String
  FacilityName = FName
End Property

Public Property Let FacilityName(ByVal vNewValue As String)
  FName = vNewValue
End Property
