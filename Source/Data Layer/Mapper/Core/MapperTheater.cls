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
Attribute VB_Name = "MapperTheater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Provides all Data Access for Core modules against the GEMS schema in Oracle.

Option Explicit

Implements iMapperCore

Private mConn As CHCSII_CONN.Conn
Private mCache As CacheMapper

Private Sub iMapperCore_Init(Support As DL_Support.iParameters, Cache As DL_Support.CacheMapper)
  Set mConn = Support.Item("DBConn").Value
  Set mCache = Cache
End Sub

Private Function iMapperCore_IsTrustedUser(ByVal ApplicationNCID As String, ByVal UserID As String, ByVal Password As String) As DTOs.LoggedInUser
Dim User As DTOs.LoggedInUser
Dim rs As Recordset
Dim SQLStmt As String

Const LOGON_USER_SQL As String = "SELECT user_ncid, role_ncid FROM hems_users WHERE logon_name = '%USERID'"
    
  On Error GoTo ErrHandler
  
  Set User = New DTOs.LoggedInUser
  
  SQLStmt = Replace(LOGON_USER_SQL, "%USERID", UserID)
  Set rs = mConn.CHCSII_DAS(Auto).OpenRecordset(SQLStmt)
  
  If Not rs.EOF And Not rs.BOF Then
    User.RoleNCID = vbNullString & rs("role_ncid")
    User.UserNCID = vbNullString & rs("user_ncid")
  End If
  
  rs.Close
  Set rs = Nothing
  
  Set iMapperCore_IsTrustedUser = User
  
  Exit Function
  
ErrHandler:
  Err.Raise Err.Number, "MapperTheater.iMapperCore_IsTrustedUser", Err.Description
    
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
  OSH.Init mConn, mCache
  Set iMapperCore_OSHGetValue = OSH.GetValue(BinItem)
  
End Function

Private Sub iMapperCore_OSHRetrieveConfiguration(Parameters As DL_Support.iParameters)
Dim OSH As OSHCustom

  Set OSH = New OSHCustom
  OSH.Init mConn, mCache
  OSH.RetrieveConfiguration Parameters

End Sub

Private Function iMapperCore_OSHSaveChanges(BinItems As DTOs.OshBinItems, ByVal SessionKey As String) As Boolean
Dim OSH As OSHCustom

  Set OSH = New OSHCustom
  OSH.Init mConn, mCache
  iMapperCore_OSHSaveChanges = OSH.SaveChanges(BinItems)
  
End Function

Private Sub iMapperCore_ReleaseHostlink()
'
End Sub

'Gets the abbr rank description for a given Rank NCID.
Private Function iMapperCore_RetrieveAbbrRank(ByVal RankNCID As String) As String
Dim Setup As ModuleSetup

  Set Setup = New ModuleSetup
  
  Setup.Init mConn
  iMapperCore_RetrieveAbbrRank = Setup.RetrieveAbbrRank(RankNCID)
  
End Function

Private Function iMapperCore_RetrieveConceptValues(ConceptValues As DTOs.ConceptValues, Optional ByVal Context As String = "2000") As DTOs.ConceptValues
Dim rs As Recordset
Dim SQLStmt As String
Dim i As Long
Dim Concepts As DTOs.ConceptValues
Dim Value As DTOs.ConceptValue
Dim NCID As String
Dim oShared As New CWShared

Const NCID_SELECT As String = "SELECT Description, DomainNCID FROM ncids WHERE NCID = '%UniqueID'"
Const ORA_NAME_SELECT As String = "SELECT last_name || ', ' || first_name as FullName FROM mmi_name WHERE unit_number = %UNIT_NUM"
Const MYSQL_NAME_SELECT As String = "SELECT concat(last_name, ', ', first_name) as FullName FROM mmi_name WHERE unit_number = %UNIT_NUM"

  On Error GoTo ErrHandler
  
  Set Concepts = New DTOs.ConceptValues
  
  For i = 1 To ConceptValues.Count
    If Len(ConceptValues.Item(i).NCID) > 0 Then
      NCID = ConceptValues.Item(i).NCID
      Set Value = Concepts.Add(NCID)
      Value.NCID = NCID
      
      SQLStmt = Replace(NCID_SELECT, "%UniqueID", NCID)
      Set rs = mConn.CHCSII_DAS.OpenRecordset(SQLStmt)
            
      If Not rs.EOF And Not rs.BOF Then
        If Not IsNumeric(rs(0)) Then
          Value.Representation = vbNullString & rs.Fields(0).Value
        End If
      End If
                  
      rs.Close
      
      If Len(Value.Representation) = 0 Then
        If oShared.UseMySQLDB Then
            SQLStmt = Replace(MYSQL_NAME_SELECT, "%UNIT_NUM", NCID)
        Else
            SQLStmt = Replace(ORA_NAME_SELECT, "%UNIT_NUM", NCID)
        End If
        
        Set rs = mConn.CHCSII_DAS.OpenRecordset(SQLStmt)
        
        If Not rs.EOF And Not rs.BOF Then
          Value.Representation = vbNullString & rs.Fields(0).Value
        End If
        
        rs.Close
      End If
    End If
  Next
  
  Set rs = Nothing
  
  Set iMapperCore_RetrieveConceptValues = Concepts
  
  Exit Function
  
ErrHandler:
  Err.Raise Err.Number, "MapperTheater.iMapperCore_RetrieveConceptValues", Err.Description

End Function

'Retrieves Folder Items from Folder_Items table.
Private Function iMapperCore_RetrieveFolderItems() As DTOs.FolderItems
Dim Setup As ModuleSetup

  Set Setup = New ModuleSetup
  
  Setup.Init mConn
  Set iMapperCore_RetrieveFolderItems = Setup.RetrieveFolderItems

End Function

Private Function iMapperCore_RetrieveLocalBlockValues(Names As DL_Support.iParameters, Optional ByVal ApplicationNCID As String = vbNullString) As DL_Support.iList
'
End Function

'Returns the configuration for modules.
'If UserIsPatient is True then modules used by a patient are loaded only.
'If ReturnAll is true then all module configurations are returned even ones that are not selected.
Private Function iMapperCore_RetrieveModuleConfigurations(CommandSwitches As DL_Support.iParameters, Optional ByVal UserIsPatient As Boolean = False, Optional ByVal ReturnAll As Boolean = False) As DTOs.ModuleConfigs
Dim Setup As ModuleSetup

  Set Setup = New ModuleSetup
  
  Setup.Init mConn
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
'
End Sub

Private Sub iMapperCore_SaveUserCache()
'
End Sub
