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
Attribute VB_Name = "MapperPGUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Provides all Data Access for Core modules against a Web Services (PGUI Server).

Option Explicit

Implements iMapperCore

Private mConn As CHCSII_CONN.Conn
Private mCache As CacheMapper

Private Sub iMapperCore_Init(Support As DL_Support.iParameters, Cache As DL_Support.CacheMapper)
  Set mConn = Support.Item("DBConn").Value
  Set mCache = Cache
End Sub

Private Function iMapperCore_IsTrustedUser(ByVal ApplicationNCID As String, ByVal UserID As String, ByVal Password As String) As DTOs.LoggedInUser

End Function

Private Sub iMapperCore_OSHCleanUpTempFiles()
  'Not used
End Sub

Private Function iMapperCore_OSHDeleteBin(ByVal OwnerNCID As String) As Boolean
  iMapperCore_OSHDeleteBin = True
End Function

Private Function iMapperCore_OSHGetValue(BinItem As DTOs.OshBinItem, ByVal SessionKey As String) As DTOs.OshBinItem
Dim OSH As OSHCustom
  
  Set OSH = New OSHCustom
  OSH.Init mConn, mCache, LGS, True
  Set iMapperCore_OSHGetValue = OSH.GetValue(BinItem)
  
End Function

Private Sub iMapperCore_OSHRetrieveConfiguration(Parameters As DL_Support.iParameters)
Dim OSH As OSHCustom

  Set OSH = New OSHCustom
  OSH.Init mConn, mCache, LGS
  OSH.RetrieveConfiguration Parameters

End Sub

Private Function iMapperCore_OSHSaveChanges(BinItems As DTOs.OshBinItems, ByVal SessionKey As String) As Boolean
Dim OSH As OSHCustom

  Set OSH = New OSHCustom
  OSH.Init mConn, mCache, LGS, True
  iMapperCore_OSHSaveChanges = OSH.SaveChanges(BinItems)
  
End Function

Private Sub iMapperCore_ReleaseHostlink()

End Sub

'Gets the abbr rank description for a given Rank NCID.
Private Function iMapperCore_RetrieveAbbrRank(ByVal RankNCID As String) As String
Dim Setup As ModuleSetup

  Set Setup = New ModuleSetup
  
  Setup.Init mConn, LGS
  iMapperCore_RetrieveAbbrRank = Setup.RetrieveAbbrRank(RankNCID)
  
End Function

Private Function iMapperCore_RetrieveConceptValues(ConceptValues As DTOs.ConceptValues, Optional ByVal Context As String = "2000") As DTOs.ConceptValues

End Function

'Retrieves Folder Items from Folder_Items table.
Private Function iMapperCore_RetrieveFolderItems() As DTOs.FolderItems
Dim Setup As ModuleSetup

  Set Setup = New ModuleSetup
  
  Setup.Init mConn, LGS
  Set iMapperCore_RetrieveFolderItems = Setup.RetrieveFolderItems

End Function

Private Function iMapperCore_RetrieveLocalBlockValues(Names As DL_Support.iParameters, Optional ByVal ApplicationNCID As String = vbNullString) As DL_Support.iList

End Function

'Returns the configuration for modules.
'If UserIsPatient is True then modules used by a patient are loaded only.
'If ReturnAll is true then all module configurations are returned even ones that are not selected.
Private Function iMapperCore_RetrieveModuleConfigurations(CommandSwitches As DL_Support.iParameters, Optional ByVal UserIsPatient As Boolean = False, Optional ByVal ReturnAll As Boolean = False) As DTOs.ModuleConfigs
Dim Setup As ModuleSetup

  Set Setup = New ModuleSetup
  
  Setup.Init mConn, LGS
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

End Sub

Private Sub iMapperCore_SaveUserCache()

End Sub
