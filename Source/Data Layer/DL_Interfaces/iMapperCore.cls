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
Attribute VB_Name = "iMapperCore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub Init(ByRef Support As iParameters, ByRef Cache As CacheMapper)
End Sub

'Removes any INI file created by 3M Bin objects.
Public Sub OSHCleanUpTempFiles()
End Sub

'Removes a whole set of configuration information.
Public Function OSHDeleteBin(ByVal OwnerNCID As String) As Boolean
End Function

'Gets a single configuration value.
Public Function OSHGetValue(ByRef BinItem As OshBinItem, ByVal SessionKey As String) As OshBinItem
End Function

'Inits NCIDs used to get and save configuration information.
Public Sub OSHRetrieveConfiguration(ByRef Parameters As iParameters)
End Sub

'Saves any changes to configuration information.
Public Function OSHSaveChanges(ByRef BinItems As OshBinItems, ByVal SessionKey As String) As Boolean
End Function

'Gets the abbr rank description for a given Rank NCID.
Public Function RetrieveAbbrRank(ByVal RankNCID As String) As String
End Function

'Get Folder Items information for a Tree view type display.  Folder Items have a Ordinal number and parent ordinal number
'that can be used to build a tree of OLE Servers.
Public Function RetrieveFolderItems() As FolderItems
End Function

'Returns the configuration for modules.
'If UserIsPatient is True then modules used by a patient are loaded only.
'If ReturnAll is true then all module configurations are returned even ones that are not selected.
Public Function RetrieveModuleConfigurations(ByRef CommandSwitches As iParameters, Optional ByVal UserIsPatient As Boolean = False, _
                                              Optional ByVal ReturnAll As Boolean = False) As ModuleConfigs
End Function

Public Function IsTrustedUser(ByVal ApplicationNCID As String, ByVal UserID As String, ByVal Password As String) As DTOs.LoggedInUser
End Function

Public Function RetrieveConceptValues(ByRef ConceptValues As DTOs.ConceptValues, Optional ByVal Context As String = "2000") As DTOs.ConceptValues
End Function

Public Function RetrieveCHCSIIRegistryValue(ByVal PrimaryKey As String, ByVal SubKey As String, ByVal ValueKey As String) As String
End Function

Public Function SaveCHCSIIRegistryValue(ByVal PrimaryKey As String, ByVal SubKey As String, ByVal ValueKey As String, ByVal Value As Variant) As Boolean
End Function

Public Sub SaveLocalBlockValues(ByRef Values As DL_Support.iParameters)
End Sub

Public Function RetrieveLocalBlockValues(ByRef Names As DL_Support.iParameters, Optional ByVal ApplicationNCID As String = vbNullString) As DL_Support.iList
End Function

Public Sub ReleaseHostlink()
End Sub

Public Sub SaveUserCache()
End Sub
