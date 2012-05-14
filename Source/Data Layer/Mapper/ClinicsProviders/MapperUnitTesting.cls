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
Attribute VB_Name = "MapperUnitTesting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Provides all Data Access for ClinicClient module when application is in a test mode.

Implements DL_Interfaces.iMapperClinicClient

Private mCache As CacheMapper        'Core Cache manager

Private Sub iMapperClinicClient_Init(Support As DL_Support.iParameters, ByRef Cache As CacheMapper)
  If gobjMiscParam Is Nothing Then
    Set gobjMiscParam = New MiscSearchParam
  End If
  
  Set mCache = Cache
End Sub

Private Function iMapperClinicClient_GetClinics(NPParams As DL_Support.iParameters) As DTOs.Clinics
Dim objClinics As DTOs.Clinics

    Set objClinics = GetCachedClinics(NPParams)

    SaveCachedClinics objClinics, NPParams
    
    Set iMapperClinicClient_GetClinics = objClinics
End Function

Private Function iMapperClinicClient_GetProviders(NPParams As iParameters) As DTOs.Providers
Dim objProviders As DTOs.Providers

    Set objProviders = GetCachedProviders(NPParams)

    SaveCachedProviders objProviders, NPParams
    
    Set iMapperClinicClient_GetProviders = objProviders
End Function

Private Function iMapperClinicClient_GetClinicsWorkload(NPParams As iParameters, objClinics As DTOs.Clinics) As Boolean
    
    Call GetCachedWorkLoads(NPParams, objClinics)

    SaveCachedClinics objClinics, NPParams
    
    iMapperClinicClient_GetClinicsWorkload = True
End Function

Private Function iMapperClinicClient_GetProvidersWorkload(NPParams As iParameters, objProviders As DTOs.Providers) As Boolean
    
    Call GetCachedWorkLoads(NPParams, objProviders)

    SaveCachedProviders objProviders, NPParams
    
    iMapperClinicClient_GetProvidersWorkload = True
End Function


Private Function GetCachedProviders(ByRef Params As iParameters) As DTOs.Providers
Dim CacheEntry As iCacheEntry
Dim objParams As iParameters
Dim Reader As ObjectPersist.ReaderXML
Dim Persist As iObjectPersist
Dim iReader As iReader

    On Error GoTo ErrHandler
    
    Set objParams = CreateParamsProvider(Params)

    'Get the user's list if it exists
    Set CacheEntry = mCache.Retrieve(CACHE_NAME, objParams)
    If CacheEntry.Valid Then
        Set GetCachedProviders = CacheEntry.Value
    Else
        Set Persist = New Providers
        Set Reader = New ObjectPersist.ReaderXML
        Set iReader = Reader
        iReader.DocumentVersion = Persist.Version
        Reader.Load App.Path & "\ProvidersUnitTest.xml"

        Persist.Restore Reader
        
        Set GetCachedProviders = Persist
    End If
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "MapperUnitTesting.GetCachedProviders", Err.Description, Err.HelpFile, Err.HelpContext
    
End Function

Private Function GetCachedClinics(ByRef Params As iParameters) As DTOs.Clinics
Dim CacheEntry As iCacheEntry
Dim objParams As iParameters
Dim Reader As ObjectPersist.ReaderXML
Dim Persist As iObjectPersist
Dim iReader As iReader

    On Error GoTo ErrHandler
    
    Set objParams = CreateParamsClinic(Params)

    'Get the user's list if it exists
    Set CacheEntry = mCache.Retrieve(CACHE_NAME, objParams)
    If CacheEntry.Valid Then
        Set GetCachedClinics = CacheEntry.Value
    Else
        Set Persist = New DTOs.Clinics
        Set Reader = New ObjectPersist.ReaderXML
        Set iReader = Reader
        iReader.DocumentVersion = Persist.Version
        Reader.Load App.Path & "\ClinicUnitTest.xml"

        Persist.Restore Reader
        
        Set GetCachedClinics = Persist
    End If
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "MapperUnitTesting.GetCachedClinics", Err.Description, Err.HelpFile, Err.HelpContext
    
End Function

Private Function GetCachedWorkLoads(ByRef Params As iParameters, objList As Object) As Boolean
Dim CacheEntry As iCacheEntry
Dim objParams As iParameters
Dim Reader As ObjectPersist.ReaderXML
Dim Persist As iObjectPersist
Dim iReader As iReader
Dim objWorkloads As DTOs.Workloads
Dim objWorkLoad As DTOs.WorkLoad

On Error GoTo ErrHandler
    
    Set objParams = CreateParamsWorkLoads(Params)

    Set Persist = New DTOs.Workloads
    Set Reader = New ObjectPersist.ReaderXML
    Set iReader = Reader
    iReader.DocumentVersion = Persist.Version
    Reader.Load App.Path & "\WorkLoadsUnitTest.xml"

    Persist.Restore Reader
    
    Set objWorkloads = Persist
    If TypeOf objList Is DTOs.Clinics Then  'Clinic Workloads
      If objList.ClinicWorkloads Is Nothing Then
        For Each objWorkLoad In objWorkloads
          objList.ClinicWorkloads.Add objWorkLoad, objWorkLoad.Appt_Code
        Next
      End If
    Else  'Provider Workloads
      If objList.ProviderWorkloads Is Nothing Then
        For Each objWorkLoad In objWorkloads
          objList.ProviderWorkloads.Add objWorkLoad, objWorkLoad.Appt_Code
        Next
      End If
    End If
    
    GetCachedWorkLoads = True
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "MapperUnitTesting.GetCachedWorkloads", Err.Description, Err.HelpFile, Err.HelpContext
    
End Function

Private Sub SaveCachedClinics(ByRef Clinics As DTOs.Clinics, ByRef Params As iParameters)
Dim ClinicParams As ClinicSearchParam
    
    Set ClinicParams = Params
    
    With ClinicParams
      If Len(.ApptCode) <> 0 Or .IncludeWorkload Or .IncludeNonMTF Then
          'Do not Cache when these values are set
          Exit Sub
      End If
    End With
    
    mCache.Save CACHE_NAME, Clinics, Params

End Sub

Private Sub SaveCachedProviders(ByRef Providers As DTOs.Providers, ByRef Params As iParameters)
Dim ProvParams As ProviderSearchParam

    Set ProvParams = Params
    
    With ProvParams
      If .IncludeExtendedData Or .IncludeRoles Or .IncludeWorkload _
        Or Len(.ProviderLastName) <> 0 Or Len(.ProviderRole) <> 0 _
        Or .SignableOnly Then
          'Do not Cache when these values are set
          Exit Sub
      End If
    End With
    
    mCache.Save CACHE_NAME, Providers, Params

End Sub

Private Function CreateParamsProvider(ByRef Params As iParameters) As iParameters
Dim objSearchNPParam As ProviderSearchParam
Dim strProviderNCID As String
Dim objParams As iParameters

    On Error GoTo ErrHandler
    
    'Generate key to determine whose frequently used list to retrieve
    Set objSearchNPParam = Params
    strProviderNCID = objSearchNPParam.ProviderNCID
    
    Set objParams = New ParametersGeneric
    objParams.Add "PROVNCID" & strProviderNCID, strProviderNCID
    objParams.Add "CACHE_KEY", CACHE_NAME
    Set CreateParamsProvider = objParams
    
    Set objSearchNPParam = Nothing
    
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "MapperUnitTesting.CreateParamsProvider", Err.Description, Err.HelpFile, Err.HelpContext
    
End Function

Private Function CreateParamsClinic(ByRef Params As iParameters) As iParameters
Dim objSearchNPParam As ClinicSearchParam
Dim strUserNCID As String
Dim objParams As iParameters

    On Error GoTo ErrHandler
    
    'Generate key to determine whose frequently used list to retrieve
    Set objSearchNPParam = Params
    strUserNCID = objSearchNPParam.UserNCID
    
    Set objParams = New ParametersGeneric
    objParams.Add "USERNCID" & strUserNCID, strUserNCID
    objParams.Add "CACHE_KEY", CACHE_NAME
    Set CreateParamsClinic = objParams
    
    Set objSearchNPParam = Nothing
    
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "MapperUnitTesting.CreateParamsClinic", Err.Description, Err.HelpFile, Err.HelpContext
    
End Function

Private Function CreateParamsWorkLoads(ByRef Params As iParameters) As iParameters
Dim strClinicNCID As String
Dim objParams As iParameters

    On Error GoTo ErrHandler
    
    Set objParams = New ParametersGeneric
    objParams.Add "CLINNCID" & strClinicNCID, strClinicNCID
    objParams.Add "CACHE_KEY", CACHE_NAME
    Set CreateParamsWorkLoads = objParams
    
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "MapperUnitTesting.CreateParamsWorkLoads", Err.Description, Err.HelpFile, Err.HelpContext
    
End Function

