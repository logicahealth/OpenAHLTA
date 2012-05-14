VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iMapperClinicClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Retrive Clinic information.  SearchNPParams is DM_ClinicClient.ClinicSearchParam
Public Function GetClinics(SearchNPParams As iParameters) As DTOs.Clinics
End Function

'Retrieve Provider information.  SearchNPParams is DM_ClinicClient.ProviderSearchParam
Public Function GetProviders(SearchNPParams As iParameters) As DTOs.Providers
End Function

'Retrieves the Workload information for collection of Clinics
Public Function GetClinicsWorkload(SearchNPParams As iParameters, Clinics As DTOs.Clinics) As Boolean
End Function

'Retrieves the Provider Workload for the collection of Providers
Public Function GetProvidersWorkload(SearchNPParams As iParameters, Providers As DTOs.Providers) As Boolean
End Function

'This should only be called from the DataAccess library
Public Sub Init(ByRef Support As iParameters, ByRef Cache As CacheMapper)
End Sub


