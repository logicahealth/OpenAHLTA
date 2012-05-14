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

Public clinicName   As String
Public FacilityName As String
Public FacNcid      As String
Public ClinicNCID   As String
Public Addr         As String
Public City         As String
Public State        As String
Public Zip          As String
Public Dmis         As String
Public meprs_code   As String
Public meprs_ien    As String
Public CType        As String
Public SickCall     As Integer
Public WalkIn       As Integer
Public Inactive     As Integer
Public CIen         As String
Public Phone        As String


Public bDirty       As Boolean
Public bNew         As Boolean

Private Sub Class_Initialize()
    bNew = True
    'meprs_code = GEMSmeprs_code
    'meprs_ien = GEMSmeprs_ien
    SickCall = 1
    WalkIn = 1
    Inactive = 0
    FacNcid = cwobjLogonEnviron.FacilityNCID
    FacilityName = cwobjLogonEnviron.FacilityName
End Sub
