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
Attribute VB_Name = "Clinic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public FacilityNCID           As String
Public NCID                   As String
Public Name                   As String
Public IEN                    As String
Public MeprsCode              As String
Public MeprsDescription       As String
Public MeprsIEN               As String
Public LocationType           As String
Public WalkInAllowed          As Boolean
Public SickCallAllowed        As Boolean
Public StreetAddress          As String
Public BuildingNumber         As String
Public City                   As String
Public State                  As String
Public ZipCode                As String
Public DmisIdCode             As String
Public DefPharmacyBDM         As String
'''Workload properties
Public Default_Workload_Type  As String
Public Workloads              As Collection

Public Function Deserialize(robjSQL As ICHCSII_SQL.ISqlOps) As Boolean
   On Error Resume Next
    
   FacilityNCID = "" & robjSQL.Value("FACILITY_NCID")
   NCID = "" & robjSQL.Value("NCID")
   Name = "" & robjSQL.Value("Name")
   IEN = "" & robjSQL.Value("IEN")
   MeprsCode = "" & robjSQL.Value("Meprs_Code")
   MeprsDescription = "" & robjSQL.Value("MeprsDescription")
   MeprsIEN = "" & robjSQL.Value("Meprs_IEN")
   LocationType = "" & robjSQL.Value("LOCATION_TYPE")
   WalkInAllowed = (robjSQL.Value("WALK_IN") = 1)
   SickCallAllowed = (robjSQL.Value("SICK_CALL") = 1)
   StreetAddress = "" & robjSQL.Value("STREET_ADDRESS")
   BuildingNumber = "" & robjSQL.Value("BUILDING_NUMBER")
   City = "" & robjSQL.Value("CITY")
   State = "" & robjSQL.Value("STATE")
   ZipCode = "" & robjSQL.Value("ZIP")
   DmisIdCode = "" & robjSQL.Value("DMIS_ID_CODE")
   DefPharmacyBDM = "" & robjSQL.Value("BDM_DEF_PHARMACY_ID")
   Default_Workload_Type = "" & robjSQL.Value("DEFAULT_WORKLOAD_TYPE")
   If Default_Workload_Type = "" Then
      Default_Workload_Type = "N"
   End If
   Deserialize = True
    
End Function






Friend Sub F_Deserialize(objFields As Fields)
On Error Resume Next
    
   FacilityNCID = "" & objFields("FACILITY_NCID").Value
   NCID = "" & objFields("NCID").Value
   Name = "" & objFields("Name").Value
   IEN = "" & objFields("IEN").Value
   MeprsCode = "" & objFields("Meprs_Code").Value
   MeprsDescription = "" & objFields("MeprsDescription").Value
   MeprsIEN = "" & objFields("Meprs_IEN").Value
   LocationType = "" & objFields("LOCATION_TYPE").Value
   WalkInAllowed = (objFields("WALK_IN").Value = 1)
   SickCallAllowed = (objFields("SICK_CALL").Value = 1)
   StreetAddress = "" & objFields("STREET_ADDRESS").Value
   BuildingNumber = "" & objFields("BUILDING_NUMBER").Value
   City = "" & objFields("CITY").Value
   State = "" & objFields("STATE").Value
   ZipCode = "" & objFields("ZIP").Value
   DmisIdCode = "" & objFields("DMIS_ID_CODE").Value
   DefPharmacyBDM = "" & objFields("BDM_DEF_PHARMACY_ID").Value
   Default_Workload_Type = "" & objFields("DEFAULT_WORKLOAD_TYPE").Value
   If Default_Workload_Type = "" Then
      Default_Workload_Type = "N"
   End If

End Sub

Private Sub Class_Initialize()
   Set Workloads = New Collection
End Sub


Private Sub Class_Terminate()
   Set Workloads = Nothing
End Sub


