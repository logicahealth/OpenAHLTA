VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MapperCDRTraining"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements DL_Interfaces.iMapperClinicClient

Private mobjMapperGeneric As iMapperClinicClient

Private Sub Class_Initialize()
On Error GoTo errHandler
Dim objMapper As MapperGeneric
Dim oSQLFix As iSQLFix

  Set mobjMapperGeneric = New MapperGeneric
  Set oSQLFix = New SQLSQLServer
  
  Set objMapper = mobjMapperGeneric
  Set objMapper.SQLFix = oSQLFix
  objMapper.DataAccessMode = ITT

  Exit Sub
  
errHandler:
  Err.Raise Err.Number, "MapperClinicClient.Initialize", Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Private Sub iMapperClinicClient_Init(Support As DL_Support.iParameters, ByRef Cache As CacheMapper)
  If gobjMiscParam Is Nothing Then
    Set gobjMiscParam = New MiscSearchParam
  End If
  
  Call mobjMapperGeneric.Init(Support, Cache)
End Sub

Private Function iMapperClinicClient_GetClinics(NPParams As DL_Support.iParameters) As DTOs.Clinics
    Set iMapperClinicClient_GetClinics = mobjMapperGeneric.GetClinics(NPParams)
End Function

Private Function iMapperClinicClient_GetProviders(NPParams As iParameters) As DTOs.Providers
    Set iMapperClinicClient_GetProviders = mobjMapperGeneric.GetProviders(NPParams)
End Function

Private Function iMapperClinicClient_GetClinicsWorkload(NPParams As iParameters, ocolClinics As DTOs.Clinics) As Boolean
  iMapperClinicClient_GetClinicsWorkload = mobjMapperGeneric.GetClinicsWorkload(NPParams, ocolClinics)
End Function

Private Function iMapperClinicClient_GetProvidersWorkload(NPParams As iParameters, ocolProviders As DTOs.Providers) As Boolean
  iMapperClinicClient_GetProvidersWorkload = mobjMapperGeneric.GetProvidersWorkload(NPParams, ocolProviders)
End Function

