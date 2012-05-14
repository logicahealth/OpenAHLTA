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

Implements DL_Interfaces.iMapperNonProviderList

Private mCache As CacheMapper        'Core Cache manager
Private mDMFactory As iDataMapperFactory   'Reference to DataAccess Factory so we can access other
                                            'mappers

Private Sub iMapperNonProviderList_Init(Support As DL_Support.iParameters, Cache As DL_Support.CacheMapper)

    On Error GoTo ErrHandler

    Init Support, Cache
    Exit Sub
  
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.Init", Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Private Sub iMapperNonProviderList_SaveList(NPParams As DL_Support.iParameters)

On Error GoTo ErrHandler
    SaveCache NPParams
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.SaveList", Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Private Function iMapperNonProviderList_Selected(NPParams As DL_Support.iParameters) As DTOs.Providers
    
On Error GoTo ErrHandler

    Set iMapperNonProviderList_Selected = GetProvider(NPParams)
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.Selected", Err.Description, Err.HelpFile, Err.HelpContext

End Function

Public Sub Init(ByRef Support As DL_Support.iParameters, ByRef Cache As DL_Support.CacheMapper)

On Error GoTo ErrHandler

  'Save objects given to us by DataAccess Factory
  Set mDMFactory = Support.Item("DataMapperFactory").Value
  Set mCache = Cache
  
  Exit Sub
  
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.Init", Err.Description, Err.HelpFile, Err.HelpContext
End Sub


Public Function GetProvider(ByRef NPParams As DL_Support.iParameters) As DTOs.Providers

    'Dim strSQL As String
    Dim objSearchParams As SearchNPParams
    Dim colClinicianNCID As Collection
    Dim objProv As DTOs.Provider
    Dim objProvs As DTOs.Providers
    
    On Error GoTo ErrHandler
    
    Set objSearchParams = NPParams
    Set colClinicianNCID = New Collection
    
    Set objProvs = GetCachedProviders(objSearchParams) 'GetProvidersList

    If LenB(objSearchParams.ClinicianNCID) <> 0 Then
        If Not objProvs.Exists(objSearchParams.ClinicianNCID) Then
            'TODO: Add existing Provider to list
            Set objProv = New Provider
            objProvs.Add objProv, objSearchParams.ClinicianNCID, 1
        Else
            'already exists in list, just move to beginning
            Set objProv = objProvs.Item(objSearchParams.ClinicianNCID)
            objProvs.Remove (objProv.NCID)
            objProvs.Add objProv, objProv.NCID, 1
            objProv.IEN = objSearchParams.ClinicianNCID 'Just give same as NCID
            objProv.Name = "Smith" & objSearchParams.ClinicianNCID & ", John"
            objProv.SignatureClass = 3
        End If
        
        'And make sure we save to cache
        SaveCachedProviders objProvs, objSearchParams

    End If
    
    Set GetProvider = objProvs
    
    Set objSearchParams = Nothing
    Set colClinicianNCID = Nothing
    
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "MapperUnitTesting.GetProvider", Err.Description, Err.HelpFile, Err.HelpContext
End Function

Private Function GetCachedProviders(ByRef Params As iParameters) As DTOs.Providers
Dim CacheEntry As iCacheEntry
Dim objParams As iParameters
Dim Reader As ObjectPersist.ReaderXML
Dim Persist As iObjectPersist
Dim iReader As iReader

    On Error GoTo ErrHandler
    
    Set objParams = CreateCacheParams(Params)

    'Get the user's list if it exists
    Set CacheEntry = mCache.Retrieve(CACHE_NAME, objParams)
    If CacheEntry.Valid Then
        Set GetCachedProviders = CacheEntry.Value
    Else
        Set Persist = New Providers
        Set Reader = New ObjectPersist.ReaderXML
        Set iReader = Reader
        iReader.DocumentVersion = Persist.Version
        Reader.Load App.Path & "\NonProviderListUnitTest.xml"

        Persist.Restore Reader
        
        Set GetCachedProviders = Persist
    End If
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "MapperUnitTesting.GetCachedProviders", Err.Description, Err.HelpFile, Err.HelpContext
    
End Function



Private Sub SaveOSHProviders(ByRef Providers As DTOs.Providers)
    Dim Writer As ObjectPersist.WriterXML
    Dim Persist As iObjectPersist
    Dim iWriter As iWriter
    
    On Error GoTo ErrHandler
    
    Set Writer = New WriterXML
    Set Persist = Providers
    Persist.Persist Writer

    Writer.Save App.Path & "\NonProviderListUnitTest.xml"
    Set iWriter = Writer
    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, "MapperUnitTesting.SaveOSHProviders", Err.Description, Err.HelpFile, Err.HelpContext
    
End Sub

Private Sub SaveCachedProviders(ByRef Providers As DTOs.Providers, ByRef Params As iParameters)
Dim objParams As iParameters

    Set objParams = CreateCacheParams(Params)
    
    mCache.Save CACHE_NAME, Providers, objParams
End Sub


Private Function CreateCacheParams(ByRef Params As iParameters) As iParameters
Dim objSearchNPParam As SearchNPParams
Dim strUserNCID As String
Dim objParams As iParameters

    On Error GoTo ErrHandler
    
    'Generate key to determine whose frequently used list to retrieve
    Set objSearchNPParam = Params
    strUserNCID = objSearchNPParam.UserNCID
    
    Set objParams = New ParametersGeneric
    objParams.Add "USERNCID" & strUserNCID, strUserNCID
    objParams.Add "CACHE_KEY", ListMode2Key(objSearchNPParam.ListMode)
    Set CreateCacheParams = objParams
    
    Set objSearchNPParam = Nothing
    
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "MapperUnitTesting.CreateCacheParams", Err.Description, Err.HelpFile, Err.HelpContext
    
End Function

Private Sub SaveCache(ByRef NPParams As DL_Support.iParameters)
Dim objProvs As DTOs.Providers

    'Get existing providers from Cache
    Set objProvs = GetCachedProviders(NPParams)
    
    'And save permanently
    SaveOSHProviders objProvs
End Sub

Private Function ListMode2Key(ByVal eListMode As ListMode) As String
Dim strRet As String

    Select Case eListMode
        Case Provider
            strRet = CACHE_KEY_NONPROVIDERS
        Case AdditionalProvider
            strRet = CACHE_KEY_ADDPROVIDERS
    End Select
    
    ListMode2Key = strRet
End Function

