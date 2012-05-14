VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MapperGeneric"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const NCID_CW_APPLICATION                   As String = "1801"

Implements DL_Interfaces.iMapperNonProviderList

Private mConn As CHCSII_CONN.Conn    'CHCSII Connection
Private mCache As CacheMapper        'Core Cache manager
Private mDMFactory As iDataMapperFactory   'Reference to DataAccess Factory so we can access other
                                            'mappers

Private mobjLegacyOsh As iLegacyOsh     'Set by DataAccess layer, points to correct method of
                                        'retrieving legacy OSH values

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
'Called by Concrete mappers to pass along necessary references

On Error GoTo ErrHandler

  'Save objects given to us by DataAccess Factory
  Set mConn = Support.Item("DBConn").Value
  Set mDMFactory = Support.Item("DataMapperFactory").Value
  Set mCache = Cache
  
  'Provider the Legacy Osh with a reference to the Conn
  Set mobjLegacyOsh.Conn = mConn
  
  Exit Sub
  
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.Init", Err.Description, Err.HelpFile, Err.HelpContext
End Sub


Private Function GetProvider(ByRef NPParams As DL_Support.iParameters) As DTOs.Providers

    Dim strSQL As String
    Dim objSearchParams As SearchNPParams
    Dim colClinicianNCID As Collection
    Dim rsProviders As Recordset
    Dim objProv As DTOs.Provider
    Dim objProvs As DTOs.Providers
    
    On Error GoTo ErrHandler
    
    Set objSearchParams = NPParams
    Set colClinicianNCID = New Collection
    
    Set objProvs = GetCachedProviders(objSearchParams) 'GetProvidersList
    
    If LenB(objSearchParams.ClinicianNCID) <> 0 Then
        If Not objProvs.Exists(objSearchParams.ClinicianNCID) Then
            colClinicianNCID.Add objSearchParams.ClinicianNCID
            
            strSQL = GetAllProvSQL(colClinicianNCID, objSearchParams.FacilityNCID)
            Set rsProviders = mConn.CHCSII_DAS.OpenRecordset(strSQL)
            
            If rsProviders.RecordCount = 0 Then
                strSQL = GetAllClinicProvSQL(colClinicianNCID, objSearchParams.FacilityNCID)
                Set rsProviders = mConn.CHCSII_DAS.OpenRecordset(strSQL)
            End If
            
            AddProvidersFromRS objProvs, rsProviders
            
        Else
            'already exists in list, just move to beginning
            Set objProv = objProvs.Item(objSearchParams.ClinicianNCID)
            objProvs.Remove (objProv.NCID)
            objProvs.Add objProv, objProv.NCID, 1
        End If
        
        'And make sure we save to cache
        SaveCachedProviders objProvs, objSearchParams

    End If
    
    Set GetProvider = objProvs
    Set objSearchParams = Nothing
    Set colClinicianNCID = Nothing
    
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.GetProvider", Err.Description, Err.HelpFile, Err.HelpContext

End Function

Private Function GetMultiProviders(ByRef strNCIDS() As String, ByVal strFacilityNCID As String) As DTOs.Providers
'strNCIDS is comma delimited string of NCIDS

Dim colNCIDs As Collection
Dim colSecNCIDs As Collection  'NCIDS not found in Provider table
Dim i As Integer
Dim rsProviders As Recordset
Dim objProvs As Providers
Dim objSortedProvs As Providers
Dim objProv As Provider
Dim strSQL As String

    On Error GoTo ErrHandler
    
    'Build collection of NCIDs used to generate SQL
    Set colNCIDs = New Collection
    
    For i = LBound(strNCIDS) To UBound(strNCIDS)
        If Len(strNCIDS(i)) <> 0 Then
            colNCIDs.Add strNCIDS(i)
        End If
    Next i
    
    
    Set objProvs = New Providers
        
    'Query for the NCIDS
    If UBound(strNCIDS) >= 0 Then
        strSQL = GetAllProvSQL(colNCIDs, strFacilityNCID)
        Set rsProviders = mConn.CHCSII_DAS.OpenRecordset(strSQL)
    
        'Populate Providers
        AddProvidersFromRS objProvs, rsProviders
    End If

    'If we're missing info for Clinicians, try Clinic_Provider table
    If Not objProvs.Count = colNCIDs.Count Then
        Set colSecNCIDs = New Collection
        For i = 1 To colNCIDs.Count
            If Len(colNCIDs(i)) <> 0 Then
                If Not objProvs.Exists(CStr(colNCIDs(i))) Then
                    colSecNCIDs.Add colNCIDs(i)
                End If
            End If
        Next i
        
        If colSecNCIDs.Count <> 0 Then
            strSQL = GetAllClinicProvSQL(colSecNCIDs, strFacilityNCID)
            Set rsProviders = mConn.CHCSII_DAS.OpenRecordset(strSQL)
            
            AddProvidersFromRS objProvs, rsProviders
        End If
    End If
    
    Set objSortedProvs = New Providers
    
    For i = 1 To colNCIDs.Count
        If objProvs.Exists(colNCIDs(i)) Then
            Set objProv = objProvs.Item(colNCIDs(i))
            objSortedProvs.Add objProv, objProv.NCID, 1
        End If
    Next i
        
    
    Set GetMultiProviders = objSortedProvs
   
    Set colNCIDs = Nothing
    Set colSecNCIDs = Nothing
    
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.GetMultiProviders", Err.Description, Err.HelpFile, Err.HelpContext
    
End Function

Private Sub AddProvidersFromRS(ByRef Providers As DTOs.Providers, ByRef RS As Recordset)
'Generate new Provider entries from a recordset

Dim strNCID As String
Dim objProv As Provider
    
    On Error GoTo ErrHandler
    
    If RS.RecordCount > 0 Then
        RS.MoveFirst
        
        Do While Not RS.EOF
            strNCID = CStr(RS.Fields("CLINICIAN_NCID"))
            If Not Providers.Exists(strNCID) Then

                Set objProv = New Provider
                Providers.Add objProv, strNCID, 1
                
                objProv.IEN = RS.Fields("PROVIDER_IEN")
                objProv.Name = RS.Fields("PROVIDER_NAME")
                objProv.NCID = strNCID
                objProv.SignatureClass = GetHighestSigClass(RS)
            End If
            
            RS.MoveNext
        Loop
    End If
    
    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.AddProvidersFromRS", Err.Description, Err.HelpFile, Err.HelpContext
    
End Sub

Private Function GetCachedProviders(ByRef Params As iParameters) As DTOs.Providers
Dim CacheEntry As iCacheEntry
Dim objParams As iParameters
Dim objProvs As DTOs.Providers

    On Error GoTo ErrHandler
    
    Set objParams = CreateCacheParams(Params)

    'Get the user's list if it exists
    Set CacheEntry = mCache.Retrieve(CACHE_NAME, objParams)
    If CacheEntry.Valid Then
        Set GetCachedProviders = CacheEntry.Value
    Else
        Set objProvs = GetOSHProviders(Params)
        If objProvs Is Nothing Then
            Set objProvs = New DTOs.Providers
        End If
        Set GetCachedProviders = objProvs
    End If
    
    
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.GetCachedProviders", Err.Description, Err.HelpFile, Err.HelpContext

End Function

Private Function GetOSHProviders(ByRef Params As iParameters) As DTOs.Providers
    Dim objDMCore As iMapperCore
    Dim objBinItem As DTOs.OshBinItem
    Dim objSearchNPParams As SearchNPParams
    Dim objProvBinItem As DTOs.OshBinItem
    Dim strContents As String
    Dim objProvs As Providers
    Dim strNCIDS() As String
    Dim colNPOsh As Collection
    Dim i As Integer
    
    On Error GoTo ErrHandler
    
    Set objSearchNPParams = Params
    
    Set objBinItem = New DTOs.OshBinItem
    objBinItem.ApplicationNCID = NCID_CW_APPLICATION
    objBinItem.OwnerNCID = objSearchNPParams.UserNCID
    objBinItem.Section = CACHE_SECTION
    objBinItem.Key = ListMode2Key(objSearchNPParams.ListMode)
    
    Set objDMCore = mDMFactory.CoreMapper
    Set objProvBinItem = objDMCore.OSHGetValue(objBinItem, CACHE_NAME & "-" & objSearchNPParams.UserNCID)
    
    strContents = objProvBinItem.Value
    
    If Len(strContents) = 0 Then
        'Check "Legacy" OSH in case this is our first time for the current user
        If Not mobjLegacyOsh Is Nothing Then
            Set colNPOsh = mobjLegacyOsh.GetOshItems("Non_Providers_OE", "Providers-" & objSearchNPParams.UserNCID, objSearchNPParams.UserNCID)
        End If
        For i = 1 To colNPOsh.Count
            strContents = strContents & colNPOsh(i) & ","
        Next
        strContents = Left$(strContents, Len(strContents) - 1)
    End If
    
    If Len(strContents) <> 0 Then
        'Now have list of Clinicians in our list, get the data for these
        strNCIDS = Split(strContents, ",")
        Set objProvs = GetMultiProviders(strNCIDS, objSearchNPParams.FacilityNCID)
    Else
        'Hand back an empty set of Providers
        Set objProvs = New DTOs.Providers
    End If
    
    Set GetOSHProviders = objProvs
    
    Set objSearchNPParams = Nothing
    
    Set colNPOsh = Nothing
    
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.GetOSHProviders", Err.Description, Err.HelpFile, Err.HelpContext
    
End Function


Private Sub SaveOSHProviders(ByRef Providers As DTOs.Providers, ByRef Params As iParameters)
    Dim objDMCore As iMapperCore
    Dim objBinItems As OshBinItems
    Dim objBinItem As OshBinItem
    Dim strContents As String
    Dim objProv As Provider
    
    On Error GoTo ErrHandler
    
    For Each objProv In Providers
        strContents = strContents & objProv.NCID & ","
    Next
    
    If Len(strContents) <> 0 Then
        strContents = Left$(strContents, Len(strContents) - 1)
    End If
    
    'Create the BinItem
    Dim objSearchNPParams As SearchNPParams
    
    Set objSearchNPParams = Params
    
    Set objBinItems = New OshBinItems
    Set objBinItem = objBinItems.Add
    
    'Set the BinItem Properties
    objBinItem.ApplicationNCID = NCID_CW_APPLICATION
    objBinItem.OwnerNCID = objSearchNPParams.UserNCID
    objBinItem.Section = CACHE_SECTION 'CACHE_NAME
    objBinItem.Key = ListMode2Key(objSearchNPParams.ListMode)
    objBinItem.Value = strContents
    
    'And save
    Set objDMCore = mDMFactory.CoreMapper
    objDMCore.OSHSaveChanges objBinItems, CACHE_NAME & "-" & objSearchNPParams.UserNCID
    
    Set objSearchNPParams = Nothing
    
    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.SaveOSHProviders", Err.Description, Err.HelpFile, Err.HelpContext
    
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
    Err.Raise Err.Number, "MapperGeneric.CreateCacheParams", Err.Description, Err.HelpFile, Err.HelpContext
    
End Function

Private Sub SaveCache(ByRef NPParams As DL_Support.iParameters)
Dim objProvs As DTOs.Providers

    'Get existing providers from Cache
    Set objProvs = GetCachedProviders(NPParams)
    
    'And save permanently
    SaveOSHProviders objProvs, NPParams
End Sub

Private Function GetAllClinicProvSQL(ByRef colNCIDs As Collection, ByVal strFacilityNCID As String) As String
Dim strFront As String
Dim strTemp As String
Dim strSQL As String

    On Error GoTo ErrHandler
    
    strFront = "SELECT DISTINCT PROVIDER_IEN, PROVIDER_NAME, CLINICIAN_NCID, SIGNATURE_CLASS FROM CLINIC_PROVIDER "
    If colNCIDs.Count <> 0 Then
        strTemp = GetClinicianClause(colNCIDs)
        strTemp = Replace(strTemp, "%CLINICIAN_NCID%", "CLINICIAN_NCID")
        
        strSQL = strFront & strTemp & " AND FACILITY_NCID = " & strFacilityNCID
    End If
    
    GetAllClinicProvSQL = strSQL
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.GetAllClinicProvSQL", Err.Description, Err.HelpFile, Err.HelpContext
    
End Function

Private Function GetAllProvSQL(ByRef colNCIDs As Collection, ByVal strFacilityNCID As String) As String
Dim strTemp As String
Dim strSQL As String

    On Error GoTo ErrHandler
    
    strSQL = "SELECT DISTINCT IEN AS PROVIDER_IEN, Name AS PROVIDER_NAME, "
    strSQL = strSQL & " NCID AS CLINICIAN_NCID, SIGNATURE_CLASS FROM PROVIDER "
    
    strTemp = GetClinicianClause(colNCIDs)
    strTemp = Replace(strTemp, "%CLINICIAN_NCID%", "NCID")
    strSQL = strSQL & strTemp
    
    
    strSQL = strSQL & " AND FACILITY_NCID = " & strFacilityNCID
    
    GetAllProvSQL = strSQL
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.GetAllProvSQL", Err.Description, Err.HelpFile, Err.HelpContext
    
End Function



Private Function GetClinicianClause(ByRef colNCIDs As Collection) As String
Dim strTemp As String
Dim iCount As Integer
    
    On Error GoTo ErrHandler
    
    strTemp = "WHERE %CLINICIAN_NCID% IN ("
    For iCount = 1 To colNCIDs.Count
        If iCount = 1 Then
            strTemp = strTemp & colNCIDs(iCount)
        Else
            strTemp = strTemp & ", " & colNCIDs(iCount)
        End If
    Next iCount
    strTemp = strTemp & ")"

    GetClinicianClause = strTemp
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.GetClinicianClause", Err.Description, Err.HelpFile, Err.HelpContext
    
End Function



Private Function GetHighestSigClass(ByRef rsProv As Recordset) As Integer
Dim intSigClass As Integer
Dim intRSPosition As Integer

    On Error GoTo ErrHandler
    
    'Filter our RS for only the current Clinician
    intRSPosition = rsProv.AbsolutePosition
    rsProv.Filter = "CLINICIAN_NCID = " & rsProv.Fields("CLINICIAN_NCID")

    If rsProv.RecordCount = 1 Then  'If there is more than one then "...Houston we have a problem..."
        If Not IsNull(rsProv.Fields("SIGNATURE_CLASS").Value) And LenB(Trim$(rsProv.Fields("SIGNATURE_CLASS").Value)) <> 0 Then
            intSigClass = rsProv.Fields("SIGNATURE_CLASS").Value
        End If
    Else
        If rsProv.RecordCount > 1 Then 'We have a problem so go and get the greatest value
            rsProv.MoveFirst
            Do Until rsProv.EOF
            'Obviously, if we have more than one, then we need to find the CLASS_SIGNATURE
            'and they exist with more than one SIGNATURE_CLASS in the Dev Environment
            'meaning that it's possible, but I hope not.
                If Not IsNull(rsProv.Fields("SIGNATURE_CLASS").Value) And LenB(Trim$(rsProv.Fields("SIGNATURE_CLASS").Value)) <> 0 Then
                    If rsProv.Fields("SIGNATURE_CLASS").Value > intSigClass Then
                        intSigClass = rsProv.Fields("SIGNATURE_CLASS").Value
                    End If
                End If
                rsProv.MoveNext
            Loop
        Else
            intSigClass = 0
        End If
    End If
    
    'Restore the RS to what it was before we filtered
    rsProv.Filter = vbNullString
    rsProv.Move intRSPosition - 1
    
    
    GetHighestSigClass = intSigClass
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.GetHighestSigClass", Err.Description, Err.HelpFile, Err.HelpContext
    
End Function

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

Public Property Set LegacyOsh(ByRef NewValue As iLegacyOsh)
    Set mobjLegacyOsh = NewValue
End Property

Private Sub Class_Terminate()
    Set mobjLegacyOsh = Nothing
End Sub
