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
Attribute VB_Name = "CoreUnitTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Provides all Data Access for Core modules when application is in a test mode.

Option Explicit

Implements iMapperCore

Private Sub iMapperCore_Init(Support As DL_Support.iParameters, Cache As DL_Support.CacheMapper)
'
End Sub

Private Function iMapperCore_IsTrustedUser(ByVal ApplicationNCID As String, ByVal UserID As String, ByVal Password As String) As DTOs.LoggedInUser

End Function

Private Sub iMapperCore_OSHCleanUpTempFiles()
'
End Sub

Private Function iMapperCore_OSHDeleteBin(ByVal OwnerNCID As String) As Boolean
'
End Function

Private Function iMapperCore_OSHGetValue(BinItem As DTOs.OshBinItem, ByVal SessionKey As String) As DTOs.OshBinItem
Dim Items As OshBinItems
Dim Item As OshBinItem
Dim List As iList

Const OSHITEMS As String = "\OSH Items.xml"
  
  On Error GoTo ErrHandler
      
  Set List = New ListCollection
  Set Items = New OshBinItems
  
  LoadList List, App.Path & OSHITEMS
            
  Do While Not List.EOF
    Set Item = Items.Add(List.Fields.Item("ApplicationNCID") & "_" & List.Fields.Item("OwnerNCID"))
    
    With Item
      .ApplicationNCID = List.Fields.Item("ApplicationNCID")
      .BinNCID = List.Fields.Item("BinNCID")
      .Key = List.Fields.Item("Key")
      .OwnerNCID = List.Fields.Item("OwnerNCID")
      .Section = List.Fields.Item("Section")
      .Value = List.Fields.Item("Value")
    End With
        
    Set Item = Nothing
    
    List.MoveNext
  Loop
    
  Set iMapperCore_OSHGetValue = Items
  
  Exit Function
  
ErrHandler:
  Err.Raise Err.Number, "CoreUnitTest.iMapperCore_OSHGetValue", Err.Description

End Function

Private Sub iMapperCore_OSHRetrieveConfiguration(Parameters As DL_Support.iParameters)
'
End Sub

Private Function iMapperCore_OSHSaveChanges(BinItems As DTOs.OshBinItems, ByVal SessionKey As String) As Boolean
  iMapperCore_OSHSaveChanges = True
End Function

Private Sub iMapperCore_ReleaseHostlink()

End Sub

Private Function iMapperCore_RetrieveAbbrRank(ByVal RankNCID As String) As String
  iMapperCore_RetrieveAbbrRank = "Test Rank"
End Function

Private Function iMapperCore_RetrieveConceptValues(ConceptValues As DTOs.ConceptValues, Optional ByVal Context As String = "2000") As DTOs.ConceptValues

End Function

Private Function iMapperCore_RetrieveLocalBlockValues(Names As DL_Support.iParameters, Optional ByVal ApplicationNCID As String = vbNullString) As DL_Support.iList

End Function

Private Function iMapperCore_RetrieveCHCSIIRegistryValue(ByVal PrimaryKey As String, ByVal SubKey As String, ByVal ValueKey As String) As String
Dim Reg As Registry

  Set Reg = New Registry
  
  iMapperCore_RetrieveCHCSIIRegistryValue = Reg.QueryValue(PrimaryKey, SubKey, ValueKey)
  
End Function

'Retrieves OLE Server (modules) configurations from a Osh Ini Settings XML file.
Private Function iMapperCore_RetrieveModuleConfigurations(CommandSwitches As DL_Support.iParameters, Optional ByVal UserIsPatient As Boolean = False, Optional ByVal ReturnAll As Boolean = False) As DTOs.ModuleConfigs
Dim Configs As ModuleConfigs
Dim Config As ModuleConfig
Dim ConfigValues As Variant
Dim List As iList
Dim Helper As ModuleConfigHelper
Dim ValidConfig As Boolean

Const CONFIG_SETTINGS As String = "\Osh Ini Settings.xml"
Const OSH_APP_NCID As Integer = 6

  On Error GoTo ErrHandler

  Set Configs = New ModuleConfigs
  Set List = New ListCollection

  LoadList List, App.Path & CONFIG_SETTINGS

  Set Helper = New ModuleConfigHelper

  Do While Not List.EOF                                                 'Add ModuleConfig objects from List object to ModuleConfigs object
    ConfigValues = Split(List.Fields.Item("apvalue"), ";")              'Settings are in a ; separated string

    Set Config = Configs.Add(ConfigValues(OSH_APP_NCID))

    Config.ServerName = List.Fields.Item("apkey")

    'Set values in Config object from settings string
    ValidConfig = Helper.ExtractConfigInfo(Config, ConfigValues, ReturnAll, CommandSwitches)

    If Not ValidConfig Then
      Configs.Remove ConfigValues(OSH_APP_NCID)
    End If

    Set Config = Nothing

    List.MoveNext
  Loop

  Set iMapperCore_RetrieveModuleConfigurations = Configs

  Exit Function

ErrHandler:
    Err.Raise Err.Number, "CoreUnitTest.iMapperCore_RetrieveModuleConfigurations", Err.Description

End Function

Private Function iMapperCore_SaveCHCSIIRegistryValue(ByVal PrimaryKey As String, ByVal SubKey As String, ByVal ValueKey As String, ByVal Value As Variant) As Boolean
Dim Reg As Registry

  Set Reg = New Registry

  iMapperCore_SaveCHCSIIRegistryValue = Reg.SetValue(PrimaryKey, SubKey, ValueKey, Value)
  
End Function

'Retrieves Folder Items from a Folder Item XML file.
Private Function iMapperCore_RetrieveFolderItems() As DTOs.FolderItems
Dim Items As FolderItems
Dim Item As FolderItem
Dim List As iList

Const FOLDERS As String = "\Folder Items.xml"
  
  On Error GoTo ErrHandler
      
  Set List = New ListCollection
  Set Items = New FolderItems
  
  LoadList List, App.Path & FOLDERS
            
  Do While Not List.EOF                                                 'Add FolderItem objects from List object to FolderItems object
    Set Item = Items.Add(List.Fields.Item("appncid"))
    
    With Item
      .ApplicationNCID = List.Fields.Item("appncid")
      .ParentOrdinal = List.Fields.Item("parententry")
      .Key = List.Fields.Item("appname")
      .ImageName = "closedFolder"
      .NodeText = List.Fields.Item("buttontext") & vbNullString
      .Ordinal = List.Fields.Item("ordinal")
      .Section = List.Fields.Item("section")
    End With
        
    Set Item = Nothing
    
    List.MoveNext
  Loop
    
  Set iMapperCore_RetrieveFolderItems = Items
  
  Exit Function
  
ErrHandler:
  Err.Raise Err.Number, "CoreUnitTest.iMapperCore_RetrieveFolderItems", Err.Description
   
End Function

'Load XML file into a List object.
Private Sub LoadList(ByRef List As iList, ByVal XMLFile As String)
Dim XMLReader As ReaderXML
Dim Persist As iObjectPersist

  Set XMLReader = New ReaderXML
  XMLReader.Load XMLFile
  
  Set Persist = List
  Persist.Restore XMLReader

End Sub

Private Sub iMapperCore_SaveLocalBlockValues(Values As DL_Support.iParameters)

End Sub

Private Sub iMapperCore_SaveUserCache()

End Sub
