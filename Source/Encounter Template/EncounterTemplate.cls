VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EncounterTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private gobjShared As CWShared
Private msTemplateID  As String
Private msOwner       As String
Private mbIsActive    As Boolean
Private mbIsFavorite    As Boolean
Private meType        As eTemplateType
Private mbIsShared    As Boolean
Private msName        As String
Private moReferences  As ETXRefs
Private moItems       As ETItems
Private meSearchMode  As eSearchMode
Private msCreate_User As String
Private msModify_user As String
Private msFavoriteFolderID As String

'<<
Public gobjDAS As ICHCSII_DAS
Public moConn As CHCSII_CONN.Conn

Public Enum eTemplateType
    tt_null = 0
    tt_enterprise = 1
    tt_facility = 2
    tt_user = 3
End Enum

Public Enum eSearchMode
    sm_null = 0
    sm_exact = 1
    sm_beginswith = 2
End Enum



'Private Sub AddToBaseTable()
'
'    Dim oSQLEx As ICHCSII_SQL.ISqlOpsEx
'    Dim oSQL As ICHCSII_SQL.ISqlOps
'    Dim sSql As String
'
'    Dim oRef As ETXref
'    Dim oItem As ETItem
'    Dim oCHCSInterface_SQL As CHCS_Interface.cCHCSConnection
'
'    If moConn Is Nothing Then Set moConn = New CHCSII_CONN.Conn
'    Set oSQL = moConn.CHCSII_SQLOPS(Auto)
'    If gobjDAS Is Nothing Then Set gobjDAS = moConn.CHCSII_DAS(Auto)
'
'    If gobjShared.IsAppMode(modeCHCSI_GUI) Then
'        Set oCHCSInterface_SQL = gobjShared.CHCSConnection
'        oSQL.Execute "DELETE FROM TEMPLATE_SOURCE where id = " & msTemplateID
'        sSql = "INSERT INTO TEMPLATE_SOURCE (ID, TYPE, NAME, OWNER, ISSHARED, ISACTIVE, DATA_ISREQUIRED, DATA_TYPE) " & vbCrLf & _
'                "VALUES (" & msTemplateID & ", 1, " & SQLQuote(msName) & ", 0, 'Y', 'Y', 0, 2)"
'        oSQL.Execute sSql
'        For Each oItem In moItems
'            oCHCSInterface_SQL.AppendChunk "TEMPLATE_SOURCE", "ID = " & msTemplateID, "DATA", oItem.DataBlob
'        Next
'    End If
'
'    Set oRef = Nothing
'    Set oItem = Nothing
'    Set oSQLEx = Nothing
'    Set oSQL = Nothing
'    Set oCHCSInterface_SQL = Nothing
'
'End Sub


Public Property Get FavoriteFolderID() As String
    FavoriteFolderID = msFavoriteFolderID
End Property
Public Property Let FavoriteFolderID(ByVal RHV As String)
    msFavoriteFolderID = RHV
End Property

Public Property Get TemplateID() As String:               TemplateID = msTemplateID:   End Property
Public Property Let TemplateID(ByVal RHV As String):      msTemplateID = RHV:          End Property

Public Property Get Owner() As String:                    Owner = msOwner:             End Property
Public Property Let Owner(ByVal RHV As String):           msOwner = RHV:               End Property

Public Property Get IsActive() As Boolean:                IsActive = mbIsActive:       End Property
Public Property Let IsActive(ByVal RHV As Boolean):       mbIsActive = RHV:            End Property

Public Property Get eType() As Integer:             eType = meType:              End Property
Public Property Let eType(ByVal RHV As Integer):    meType = RHV:                End Property

Public Property Get IsShared() As Boolean:                IsShared = mbIsShared:       End Property
Public Property Let IsShared(ByVal RHV As Boolean):       mbIsShared = RHV:            End Property

Public Property Get Name() As String:                     Name = msName:               End Property
Public Property Let Name(ByVal RHV As String):            msName = RHV:                End Property

Public Property Get References() As ETXRefs:              Set References = moReferences:   End Property
Public Property Set References(ByVal RHV As ETXRefs):     Set moReferences = RHV:          End Property

Public Property Get Items() As ETItems:                   Set Items = moItems:         End Property
Public Property Set Items(ByVal RHV As ETItems):          Set moItems = RHV:           End Property

Public Property Get SearchMode() As Integer:          SearchMode = meSearchMode:   End Property
Public Property Let SearchMode(ByVal RHV As Integer): meSearchMode = RHV:          End Property

Public Property Get Create_User() As String:              Create_User = msCreate_User: End Property
Public Property Let Create_User(ByVal RHV As String):     msCreate_User = RHV:         End Property

Public Property Get Modify_user() As String:              Modify_user = msModify_user: End Property
Public Property Let Modify_user(ByVal RHV As String):     msModify_user = RHV:         End Property


Public Function Read()
    '-- we need the template id to get the record
    If msTemplateID = "" Then Exit Function
    
    Dim oRef As ETXref
    Dim oItem As ETItem
    Dim oSQLEx As ICHCSII_SQL.ISqlOpsEx
    Dim oSQLOps As ICHCSII_SQL.ISqlOps
    Dim oCHCSInterface_SQL As CHCS_Interface.cCHCSConnection
    
    Dim oRs As Adodb.Recordset
    Set oRs = New Adodb.Recordset
    
    Set oSQLOps = moConn.CHCSII_SQLOPS(Auto)
    Set oSQLEx = moConn.CHCSII_SQLOPS_EX(Auto)
    
    oSQLOps.Execute "select * from template_Info where id = " & msTemplateID
    If oSQLOps.EOF Then
        GoTo cleanup
    End If
    
    msOwner = "" & oSQLOps("Owner")
    msName = "" & oSQLOps("Name")
    If "" & oSQLOps("IsActive") = "Y" Then mbIsActive = True
    meType = "" & oSQLOps("type")
    msCreate_User = "" & oSQLOps("creator_ncid")
    
    If "" & oSQLOps("IsShared") = "Y" Then mbIsShared = True
    msModify_user = "" & oSQLOps("modifior_ncid")

    moReferences.ClearAll
    oSQLOps.Execute "select * from template_xref where id = " & msTemplateID
    Do While Not oSQLOps.EOF
        Set oRef = moReferences.Add
        oRef.eType = "" & oSQLOps!type_ncid
        oRef.Value = "" & oSQLOps!value_ncid
        Set oRef = Nothing
        oSQLOps.MoveNext
    Loop
    
    moItems.ClearAll
    If gobjShared.IsAppMode(modeCHCSI_GUI) Or (gobjShared.IsAppMode(modeCHCSI_GUI) And gobjShared.IsAppMode(modeITT)) Then
        Set oCHCSInterface_SQL = gobjShared.CHCSConnection
        oSQLOps.Execute "select * from template_data where id = " & msTemplateID
        Do While Not oSQLOps.EOF
            Set oItem = moItems.Add
            oItem.eType = "" & oSQLOps!Type
            oItem.DataBlob = oCHCSInterface_SQL.ReadChunk("TEMPLATE_DATA", "id = " & msTemplateID, "DATA")
            Set oItem = Nothing
            oSQLOps.MoveNext
        Loop
    Else
        oSQLEx.Execute "select * from template_data where id = " & msTemplateID
        Do While Not oSQLEx.EOF
            Set oItem = moItems.Add
            oItem.eType = "" & oSQLEx!Type
            'SCR-53016 5/4/2004 RBELL Removed chunking logic as it was not working correctly in theater mode
            oItem.DataBlob = oSQLEx("Data")
            Set oItem = Nothing
            oSQLEx.MoveNext
        Loop
    End If
    
cleanup:
    If oRs.State > 0 Then oRs.Close
    Set oRs = Nothing
End Function
Public Function Delete()
    
    On Error GoTo ErrHandler
    Delete = False
    
    If msTemplateID = "" Then Exit Function
    
    gobjDAS.ExecuteSQL "Delete from template_xref where id = " & msTemplateID
    gobjDAS.ExecuteSQL "Delete from template_data where id = " & msTemplateID
    gobjDAS.ExecuteSQL "Delete from template_favorites where template_id = " & msTemplateID
    gobjDAS.ExecuteSQL "Delete from template_info where id = " & msTemplateID
    Delete = True
    Exit Function
    
ErrHandler:
    'Resume Next
    Err.Raise Err.Number, "Template.Delete", Err.Description
    
End Function
Public Function Save()
    
    Dim oSQLEx As ICHCSII_SQL.ISqlOpsEx
    Dim oSQL As ICHCSII_SQL.ISqlOps
    
    Dim oRef As ETXref
    Dim oItem As ETItem
    Dim sShared As String
    Dim sActive As String
    
    Set moConn = New CHCSII_CONN.Conn
    
    If mbIsShared Then
        sShared = "Y"
    Else
        sShared = "N"
    End If
    
    If mbIsActive Then
        sActive = "Y"
    Else
        sActive = "N"
    End If
    
    If Val(msTemplateID) > 0 Then
    '-- this is an update
        gobjDAS.ExecuteSQL "Update template_info set type = '" & meType & "', name = '" & Replace(msName, "'", "''") & "', owner = '" _
           & msOwner & "', modifior_ncid = " & Val(msModify_user) & ", modifior_time = " & moConn.SQLDate(Now) & vbCrLf _
           & ", isshared = '" & sShared & "', isactive = '" & sActive _
           & "' where id = " & msTemplateID
    Else
        '-- this is an insert
        Dim oRs As Adodb.Recordset
        Set oRs = New Adodb.Recordset
        msTemplateID = gobjDAS.GetNextID(esENC_TEMPLATE_ID)
        
        gobjDAS.ExecuteSQL "insert into template_info (id,type,name,owner,Creator_ncid,create_time,isshared,isActive)" _
        & " values(" & msTemplateID & ",'" & meType & "','" & Replace(msName, "'", "''") & "','" & msOwner & "', '" & CStr(Val(msCreate_User)) _
        & "', " & moConn.SQLDate(Now) & ",'" & sShared & "','" & sActive & "')"
    End If

    gobjDAS.ExecuteSQL "Delete from template_xref where id = " & msTemplateID
    gobjDAS.ExecuteSQL "Delete from template_data where id = " & msTemplateID
    
    If msFavoriteFolderID = "" Then msFavoriteFolderID = "0" 'SCR-46769 12/12/03 RBELL An empty string causes an error on insert statement
    If mbIsFavorite Then
       'If personal template use owner to create favorites, otherwise use creater for a new template or modifier for updated template
       If meType = tt_user Then
            gobjDAS.ExecuteSQL "Delete from template_favorites where template_id = " & msTemplateID & _
                            " AND user_ncid = " & msOwner 'SCR-44782 RBELL 11/6/3 Moved insode IF statement to adjust Owner value
            gobjDAS.ExecuteSQL "INSERT INTO TEMPLATE_FAVORITES(USER_NCID, TEMPLATE_ID, FOLDER_ID) VALUES(" _
                     & msOwner & ", " & msTemplateID & ", 0)"
       ElseIf Val(msModify_user & vbNullString) = 0 Then
            gobjDAS.ExecuteSQL "Delete from template_favorites where template_id = " & msTemplateID & _
                            " AND user_ncid = " & msCreate_User & vbNullString 'SCR-44782 RBELL 11/6/3 Moved insode IF statement to adjust Owner value
            gobjDAS.ExecuteSQL "INSERT INTO TEMPLATE_FAVORITES(USER_NCID, TEMPLATE_ID, FOLDER_ID) VALUES(" _
                         & CStr(Val(msCreate_User & vbNullString)) & ", " & msTemplateID & ", " & msFavoriteFolderID & ")"
       Else
            gobjDAS.ExecuteSQL "Delete from template_favorites where template_id = " & msTemplateID & _
                            " AND user_ncid = " & msModify_user & vbNullString 'SCR-44782 RBELL 11/6/3 Moved insode IF statement to adjust Owner value
            gobjDAS.ExecuteSQL "INSERT INTO TEMPLATE_FAVORITES(USER_NCID, TEMPLATE_ID, FOLDER_ID) VALUES(" _
                     & CStr(Val(msModify_user & vbNullString)) & ", " & msTemplateID & ", " & msFavoriteFolderID & ")"
       End If
    End If
    
    If Not (meType = tt_user And CStr(Val(msFavoriteFolderID)) = "0") Then
        'If template is not a personal template then an additional template_favorites entry needs to be made
        'SCR-47596 RBELL ADDED Criteria to delete to delete correct favorite
        gobjDAS.ExecuteSQL "delete from template_favorites where template_id = " & SQLQuote(msTemplateID) & " AND user_ncid = " & SQLQuote(msOwner)
        gobjDAS.ExecuteSQL "insert into template_favorites(user_ncid, template_id, folder_id) values (" & SQLQuote(msOwner) & ", " & SQLQuote(TemplateID) & ", " & CStr(Val(msFavoriteFolderID)) & ")"
    End If
    
    For Each oRef In moReferences
        gobjDAS.ExecuteSQL "insert into template_xref (id, type_ncid, value_ncid) values (" _
        & msTemplateID & ",'" & oRef.eType & "','" & oRef.Value & "')"
    Next oRef
    
    Set oSQL = moConn.CHCSII_SQLOPS(Auto)
    Set oSQLEx = moConn.CHCSII_SQLOPS_EX(Auto)
    oSQLEx.Execute "Select * from template_data where id = " & msTemplateID
    For Each oItem In moItems
        If oSQLEx.BOF And oSQLEx.EOF Then
          oSQLEx.AddNew
          oSQLEx.Value("id") = msTemplateID
        End If
        oSQLEx.Value("type") = oItem.eType
        oSQLEx.Value("IsRequired") = oItem.IsRequired
        oSQLEx.AppendAsChunk "data", oItem.DataBlob
        oSQLEx.Update
    Next oItem
    Set oSQLEx = Nothing
        
End Function

Public Function IsEmpty() As Long
    IsEmpty = (moItems.Count = 0)
End Function
Public Function ClearAll()
    Set moItems = New ETItems
    Set moReferences = New ETXRefs
End Function
Public Function ClassHandle() As Long

End Function

Private Sub Class_Initialize()
'  Set Cnn = New ADODB.Connection
'  Cnn.ConnectionString = GEMS_Conn
'  Cnn.ConnectionTimeout = 30
'  Cnn.Open
    Set moConn = New CHCSII_CONN.Conn
    Set gobjDAS = moConn.CHCSII_DAS(Auto)
    
    Set moItems = New ETItems
    Set moReferences = New ETXRefs
    Set gobjShared = New CWShared
End Sub
Public Property Get IsFavorite() As Boolean
    IsFavorite = mbIsFavorite
End Property
Public Property Let IsFavorite(ByVal RHV As Boolean)
    mbIsFavorite = RHV
End Property

Private Sub Class_Terminate()
    Set gobjShared = Nothing
End Sub


