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
Attribute VB_Name = "Provider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private msClinicianNCID    As String
Private msProviderName     As String
Private msProviderIEN      As String
Private msUserIEN          As String
Private msLogonName        As String
Private msRoles            As String
Private msSignature_Class  As String
Private mbRolesSet         As Boolean

'''Workload properties
Private mcolWorkloads      As Collection
Friend Property Get fbRolesSet() As Boolean
   fbRolesSet = mbRolesSet
End Property


Public Function Mtd_Clone() As Provider
   Dim objProvider         As New Provider
   
   With objProvider
     .ClinicianNCID = msClinicianNCID
     .ProviderName = msProviderName
     .ProviderIEN = msProviderIEN
     .UserIEN = msUserIEN
     .psLogonName = msLogonName
     .psRoles = msRoles
     .psSignature_Class = msSignature_Class
   End With
   
   Set Mtd_Clone = objProvider
   
End Function
Public Property Get psRoles() As String
   psRoles = msRoles
End Property
Public Property Get pbSystemUser() As Boolean
   pbSystemUser = Not Trim$(msLogonName) = ""
End Property
Public Property Get ProviderIEN() As String
   ProviderIEN = msProviderIEN
End Property
Public Property Get ProviderName() As String
   ProviderName = msProviderName
End Property
Public Property Get ClinicianNCID() As String
   ClinicianNCID = msClinicianNCID
End Property
Friend Function F_Deserialize(objFields As adodb.Fields) As Boolean
'''Warning: System users info from hems_users may not be available (e.g. facility providers)
'''         DO NOT include it in Deserialize
On Error Resume Next
    
    msClinicianNCID = "" & objFields("Clinician_NCID").Value
    msProviderName = "" & objFields("Provider_Name").Value
    msProviderIEN = "" & objFields("Provider_IEN").Value
    msUserIEN = "" & objFields("User_IEN").Value
    msSignature_Class = "" & objFields("SIGNATURE_CLASS").Value
    
    F_Deserialize = True
    
End Function
Public Property Let ClinicianNCID(ByVal sVal As String)
   msClinicianNCID = sVal
End Property
Public Property Get psLogonName() As String
   psLogonName = msLogonName
End Property
Friend Property Let psLogonName(ByVal sVal As String)
   msLogonName = Trim$(sVal)
End Property
Public Property Let ProviderIEN(ByVal sVal As String)
   msProviderIEN = sVal
End Property
Public Property Let ProviderName(ByVal sVal As String)
   msProviderName = sVal
End Property
Friend Property Let psRoles(ByVal sVal As String)
   msRoles = sVal
   mbRolesSet = True
End Property
Public Property Let UserIEN(ByVal sVal As String)
   msUserIEN = sVal
End Property
Public Property Get UserIEN() As String
   UserIEN = msUserIEN
End Property
Public Property Get psSignature_Class() As String
   psSignature_Class = msSignature_Class
End Property
Public Property Let psSignature_Class(ByVal sVal As String)
   msSignature_Class = sVal
End Property

Private Sub Class_Initialize()
   Set mcolWorkloads = New Collection
End Sub

Public Property Set Workloads(ByVal colData As Collection)
    Set mcolWorkloads = colData
End Property

Public Property Get Workloads() As Collection
    Set Workloads = mcolWorkloads
End Property

Private Sub Class_Terminate()
   Set mcolWorkloads = Nothing
End Sub


