VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MappperCDR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Provides all Data Access for Core modules against the CDR.

Option Explicit

Implements iMapperCore

Private Const NCID_3M_DEFAULT_CONTEXT As String = "2000"

Private mConn As CHCSII_CONN.Conn
Private mCache As CacheMapper

Private Sub iMapperCore_Init(Support As DL_Support.iParameters, Cache As DL_Support.CacheMapper)
  Set mConn = Support.Item("DBConn").Value
  Set mCache = Cache
End Sub

Private Function iMapperCore_IsTrustedUser(ByVal ApplicationNCID As String, ByVal UserID As String, ByVal Password As String) As DTOs.LoggedInUser
Dim Logon As InternalLogon
Dim Success As Boolean
    
  On Error GoTo ErrHandler
         
  Set Logon = New HOSTATLLib.InternalLogon
  
  Success = Logon.bGetConnected(UserID, Password)    'attempt to logon or connect to the 3M LDR
  
  If Not Success Then
    Set Logon = Nothing
    Err.Raise 10001, "MapperCore.iMapperCore_IsTrustedUser", "Unable to make a valid logon connection."
  End If
  
  If Logon.lGetReturnCode = 110 Then
    Set Logon = Nothing
    Err.Raise 10010, "MapperCore.iMapperCore_IsTrustedUser", "This logon name and password have expired, please contact your system administrator."
  End If
  
  Set Logon = Nothing

  
  Set iMapperCore_IsTrustedUser = LoadUserNCIDs(ApplicationNCID)
  
  Exit Function
  
ErrHandler:
  Err.Raise Err.Number, Err.Source, Err.Description
  
End Function

Private Sub iMapperCore_OSHCleanUpTempFiles()
Dim OSH As OSH3m

  Set OSH = New OSH3m
  
  OSH.Init mCache
  OSH.CleanUpTempFiles
  
End Sub

Private Function iMapperCore_OSHDeleteBin(ByVal OwnerNCID As String) As Boolean
Dim OSH As OSH3m

  Set OSH = New OSH3m

  OSH.Init mCache
  iMapperCore_OSHDeleteBin = OSH.DeleteBin(OwnerNCID)
  
End Function

Private Function iMapperCore_OSHGetValue(BinItem As DTOs.OshBinItem, ByVal SessionKey As String) As DTOs.OshBinItem
Dim OSH As OSH3m

  Set OSH = New OSH3m

  OSH.Init mCache
  Set iMapperCore_OSHGetValue = OSH.GetValue(BinItem, SessionKey)
  
End Function

Private Sub iMapperCore_OSHRetrieveConfiguration(Parameters As DL_Support.iParameters)
Dim OSH As OSH3m

  Set OSH = New OSH3m

  OSH.Init mCache
  OSH.RetrieveConfiguration Parameters

End Sub

Private Function iMapperCore_OSHSaveChanges(BinItems As DTOs.OshBinItems, ByVal SessionKey As String) As Boolean
Dim OSH As OSH3m

  Set OSH = New OSH3m

  OSH.Init mCache
  iMapperCore_OSHSaveChanges = OSH.SaveChanges(BinItems, SessionKey)

End Function

Private Sub iMapperCore_ReleaseHostlink()
Dim LocalBlk As LocalBlkCtl
Dim GobalBlk As GlobalBlock
Dim Logon As InternalLogon
    
  On Error GoTo ErrHandler
  
  Set Logon = New InternalLogon
  
  'release reference to "HostLink" and allow hostlink to correctly disconnect from local block
  Logon.Logoff
  
  Set LocalBlk = New LocalBlkCtl
  
  'You need to be sure to release the "internal" reference local block has to itself to ensure that is will "go away".
  LocalBlk.RemoveInternalReference
  LocalBlk.EraseAll
  
  Set GobalBlk = New GlobalBlock
  
  'You need to be sure to release the "internal" reference public block has to itself to ensure that is will "go away".
  GobalBlk.RemoveInternalReference
  GobalBlk.EraseAll
  
  Set Logon = Nothing
  Set LocalBlk = Nothing
  Set GobalBlk = Nothing
  DoEvents   'SCR 7473 ... Previous GlobalBlock not completely "gone" before logging back on creates a new one
  
  Exit Sub
  
ErrHandler:
  Err.Raise Err.Number, "MapperCDR.iMapperCore_ReleaseHostlink", Err.Description

End Sub

'Gets the abbr rank description for a given Rank NCID.
Private Function iMapperCore_RetrieveAbbrRank(ByVal RankNCID As String) As String
Dim Setup As ModuleSetup

  Set Setup = New ModuleSetup
  
  Setup.Init mConn, CDR
  iMapperCore_RetrieveAbbrRank = Setup.RetrieveAbbrRank(RankNCID)
  
End Function

Private Function iMapperCore_RetrieveConceptValues(ConceptValues As DTOs.ConceptValues, Optional ByVal Context As String = "2000") As DTOs.ConceptValues
Dim Concept As ConceptCtrl
Dim Concepts As DTOs.ConceptValues
Dim Value As DTOs.ConceptValue
Dim NCID As String
Dim i As Long

  Set Concepts = New DTOs.ConceptValues
  
  For i = 1 To ConceptValues.Count
    If Len(ConceptValues.Item(i).NCID) > 0 Then
      Set Concept = New ConceptCtrl
      
      NCID = ConceptValues.Item(i).NCID
      Concept.UniqueId = NCID
      
      Set Value = Concepts.Add(NCID)
      
      With Value
        .NCID = NCID
        .Representation = Concept.PrefRep(Context).Representation
      End With
      
      Set Concept = Nothing
    End If
  Next
  
  Set iMapperCore_RetrieveConceptValues = Concepts
  
End Function

'Retrieves Folder Items from Folder_Items table.
Private Function iMapperCore_RetrieveFolderItems() As DTOs.FolderItems
Dim Setup As ModuleSetup

  Set Setup = New ModuleSetup
  
  Setup.Init mConn, CDR
  Set iMapperCore_RetrieveFolderItems = Setup.RetrieveFolderItems

End Function

Private Function iMapperCore_RetrieveLocalBlockValues(Names As DL_Support.iParameters, Optional ByVal ApplicationNCID As String = vbNullString) As DL_Support.iList
Dim LocalBlk As LocalBlkCtl
Dim Concept As ConceptCtrl
Dim List As iList
Dim i As Long
Dim Params As iParameters

  Set LocalBlk = New LocalBlkCtl
  Set List = New ListCollection
  
  If Len(ApplicationNCID) > 0 Then
    Set Concept = New ConceptCtrl
    
    LocalBlk.Value("ApplicationNcid") = ApplicationNCID
    
    Concept.UniqueId = ApplicationNCID
    LocalBlk.Value("ApplicationName") = Concept.PrefRep(NCID_3M_DEFAULT_CONTEXT).Representation
  
    Set Concept = Nothing
  End If
  
  Set Params = New ParametersGeneric
  
  For i = 1 To Names.Count
    Params.Add Names.Item(i).Name, LocalBlk.Value(Names.Item(i).Name)
  Next

  Set LocalBlk = Nothing
  
  With List
    .AddRow Params
    .MoveFirst
  End With
  
  Set iMapperCore_RetrieveLocalBlockValues = List
  
End Function

'Returns the configuration for modules.
'If UserIsPatient is True then modules used by a patient are loaded only.
'If ReturnAll is true then all module configurations are returned even ones that are not selected.
Private Function iMapperCore_RetrieveModuleConfigurations(CommandSwitches As DL_Support.iParameters, Optional ByVal UserIsPatient As Boolean = False, Optional ByVal ReturnAll As Boolean = False) As DTOs.ModuleConfigs
Dim Setup As ModuleSetup

  Set Setup = New ModuleSetup
  
  Setup.Init mConn, CDR
  Set iMapperCore_RetrieveModuleConfigurations = Setup.RetrieveModuleConfigurations(CommandSwitches, UserIsPatient, ReturnAll)
  
End Function

Private Function iMapperCore_RetrieveCHCSIIRegistryValue(ByVal PrimaryKey As String, ByVal SubKey As String, ByVal ValueKey As String) As String
Dim Reg As Registry

  Set Reg = New Registry
  
  iMapperCore_RetrieveCHCSIIRegistryValue = Reg.QueryValue(PrimaryKey, SubKey, ValueKey)
  
End Function

Private Function iMapperCore_SaveCHCSIIRegistryValue(ByVal PrimaryKey As String, ByVal SubKey As String, ByVal ValueKey As String, ByVal Value As Variant) As Boolean
Dim Reg As Registry

  Set Reg = New Registry

  iMapperCore_SaveCHCSIIRegistryValue = Reg.SetValue(PrimaryKey, SubKey, ValueKey, Value)
  
End Function

Private Sub iMapperCore_SaveLocalBlockValues(Values As DL_Support.iParameters)
Dim LocalBlk As LocalBlkCtl
Dim i As Long

  Set LocalBlk = New LocalBlkCtl
  
  For i = 1 To Values.Count
    LocalBlk.Value(Values.Item(i).Name) = Values.Item(i).Value
  Next
  
  Set LocalBlk = Nothing
  
End Sub

Private Sub iMapperCore_SaveUserCache()
Dim Concept As ConceptCtrl
Dim PickList As PickListCtrl
Dim Success As Boolean
    
  On Error GoTo ErrHandler
    
  Set Concept = New ConceptCtrl
  
  'save the user's cache to the OSH tables
  Concept.UniqueId = "1801"
  Success = Concept.PrefRep(NCID_3M_DEFAULT_CONTEXT).SaveCache
  Set Concept = Nothing
  
  DoEvents    'allow the progress bar to update a little
  
  Set PickList = New PickListCtrl
  
  'save the user's pick list cache to the OSH tables
  PickList.SaveCache
  Set PickList = Nothing
  
  DoEvents    'allow the progress bar to update a little

  Exit Sub
  
ErrHandler:
  Err.Raise Err.Number, "MapperCDR.iMapperCore_SaveUserCache", Err.Description

End Sub

Private Function LoadUserNCIDs(ByVal ApplicationNCID As String) As DTOs.LoggedInUser
Dim Params As iParameters
Dim List As iList
Dim User As DTOs.LoggedInUser

Const USER_NCID As String = "UserNcid"
Const USER_ROLE As String = "UserRole"

  Set User = New LoggedInUser
  
  Set Params = New ParametersGeneric
  Params.Add USER_NCID, vbNullString
  Params.Add USER_ROLE, vbNullString
  
  Set List = iMapperCore_RetrieveLocalBlockValues(Params, ApplicationNCID)
  
  
  With List.Fields
    User.UserNCID = .Item(USER_NCID).Value
    User.RoleNCID = .Item(USER_ROLE).Value
  End With
  
  Set LoadUserNCIDs = User
  
End Function

