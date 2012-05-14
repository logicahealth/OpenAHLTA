VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Folder_Item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mDMIS_ID As String
Private mParent_DMIS_ID As String
Private mFolder_Name As String
Private mFolder_NCID As String
Private mFolder_Type As String
Private mFolder_Root As String 'SCR-44264 RBELL 12/23/03 Define owner value to determine the owner of the folder
Private mSubFolders As Collection

Public Property Get DMIS_ID() As String
    DMIS_ID = mDMIS_ID
End Property

Public Property Let DMIS_ID(ByVal RHV As String)
    mDMIS_ID = RHV
End Property

Public Sub LoadSubFolders()
    
    Dim oConn As CHCSII_CONN.Conn
    Dim oDas As ICHCSII_DAS
    Dim oSQL As ICHCSII_SQL.ISqlOps
    Dim rsFolders As ADODB.Recordset
    Dim oFolder As Folder_Item
    Dim sSQL As String
    
    On Error GoTo ErrHandler
    Set oConn = New CHCSII_CONN.Conn
    Set oDas = oConn.CHCSII_DAS(Auto)
    Set oSQL = oConn.CHCSII_SQLOPS(Auto)
    
    Set mSubFolders = New Collection
    If gcFolderItems Is Nothing Then Set gcFolderItems = New Collection
    If mFolder_Type = "EN" Or mFolder_Type = "MT" Or mFolder_Type = "FV" Then
       sSQL = "SELECT FOLDER_NAME, FOLDER_ID " & vbCrLf
       sSQL = sSQL & "FROM TEMPLATE_FOLDER tf" & vbCrLf
       sSQL = sSQL & "WHERE tf.FOLDER_TYPE = '" & Folder_Type & "' " & vbCrLf
       If mFolder_NCID = mFolder_Root Then
            sSQL = sSQL & "AND tf.USER_NCID = " & mFolder_NCID & vbCrLf
            If mFolder_Type = "EN" Then sSQL = sSQL & "AND tf.FOLDER_PARENT_ID = 0 " & vbCrLf
            If mFolder_Type = "MT" Then sSQL = sSQL & "AND tf.FOLDER_PARENT_ID = tf.USER_NCID " & vbCrLf
            If mFolder_Type = "FV" Then sSQL = sSQL & "AND tf.FOLDER_PARENT_ID = 3 " & vbCrLf
       Else
            sSQL = sSQL & "AND tf.FOLDER_PARENT_ID = " & mFolder_NCID & vbCrLf
       End If
       
       sSQL = sSQL & "ORDER BY tf.FOLDER_NAME "
       Set rsFolders = oDas.OpenRecordset(sSQL)
       With rsFolders
            If Not (.BOF And .EOF) Then
                .MoveFirst
                Do Until .EOF
                    Set oFolder = New Folder_Item
                    oFolder.DMIS_ID = ""
                    oFolder.Folder_Name = rsFolders("FOLDER_NAME")
                    oFolder.Folder_NCID = rsFolders("FOLDER_ID")
                    oFolder.Folder_Type = mFolder_Type
                    oFolder.Parent_DMIS_ID = mFolder_NCID
                    oFolder.Folder_Root = mFolder_Root 'SCR-44264 RBELL 12/23/03 Set owner value to determine the origin of folder
                    mSubFolders.Add oFolder, oFolder.Folder_NCID
                    gcFolderItems.Add oFolder, oFolder.Folder_NCID
                    .MoveNext
                Loop
            End If
       End With
    ElseIf mFolder_Type = "CL" Then
        If Len(mFolder_NCID) = 4 And mFolder_NCID = mParent_DMIS_ID Then  'SCR-44264 RBELL 12/23/03 Refine condition statement
           sSQL = "SELECT DISTINCT c.NCID, c.NAME, c.DMIS_ID_CODE" & vbCrLf
           sSQL = sSQL & "FROM CLINIC c" & vbCrLf
           'SCR-53546 5/17/2004 RBELL Modified query to use subquery to get all related clinic by parent dmis ids
           sSQL = sSQL & "WHERE c.DMIS_ID_CODE IN (SELECT DMIS_ID FROM MTF_DMIS_ID WHERE DMIS_ID_PARENT = " & SQLQuote(mFolder_NCID) & ") AND " & vbCrLf
           sSQL = sSQL & "c.INACTIVE = '0' " & vbCrLf
           sSQL = sSQL & "ORDER BY c.NAME "
           Set rsFolders = oDas.OpenRecordset(sSQL)
           With rsFolders
                If Not (.BOF And .EOF) Then
                    .MoveFirst
                    Do Until .EOF
                        Set oFolder = New Folder_Item
                        oFolder.DMIS_ID = rsFolders("DMIS_ID_CODE")
                        oFolder.Folder_Name = rsFolders("NAME")
                        oFolder.Folder_NCID = rsFolders("NCID")
                        oFolder.Folder_Type = "CL"
                        oFolder.Parent_DMIS_ID = mFolder_NCID
                        oFolder.Folder_Root = rsFolders("NCID") 'SCR-44264 RBELL 12/23/03 Set owner value to determine the origin of folder
                        mSubFolders.Add oFolder, oFolder.Folder_NCID
                        gcFolderItems.Add oFolder, oFolder.Folder_NCID
                        .MoveNext
                    Loop
                End If
           End With
        Else
           sSQL = "SELECT FOLDER_NAME, FOLDER_ID " & vbCrLf
           sSQL = sSQL & "FROM TEMPLATE_FOLDER tf" & vbCrLf
           sSQL = sSQL & "WHERE tf.FOLDER_PARENT_ID  = " & mFolder_NCID & " AND " & vbCrLf 'SCR-44264 RBELL 12/23/03 Refine where clause
           sSQL = sSQL & "tf.FOLDER_TYPE = 'CL' " & vbCrLf
           sSQL = sSQL & "ORDER BY tf.FOLDER_NAME "
           Set rsFolders = oDas.OpenRecordset(sSQL)
           With rsFolders
                If Not (.BOF And .EOF) Then
                    .MoveFirst
                    Do Until .EOF
                        Set oFolder = New Folder_Item
                        oFolder.DMIS_ID = ""
                        oFolder.Folder_Name = rsFolders("FOLDER_NAME")
                        oFolder.Folder_NCID = rsFolders("FOLDER_ID")
                        oFolder.Folder_Type = "CL"
                        oFolder.Parent_DMIS_ID = mFolder_NCID
                        oFolder.Folder_Root = mFolder_Root 'SCR-44264 RBELL 12/23/03 Set owner value to determine the origin of folder
                        mSubFolders.Add oFolder, oFolder.Folder_NCID
                        gcFolderItems.Add oFolder, oFolder.Folder_NCID
                        .MoveNext
                    Loop
                End If
           End With
        End If
    End If
    
    GoTo CleanUpAndExit
    
ErrHandler:
    If Err = 457 Then Resume Next
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Folder_Item.LoadSubFolders", "TemplateClient", vbExclamation)
    'Resume Next
CleanUpAndExit:
    Set oConn = Nothing
    Set oDas = Nothing
    Set rsFolders = Nothing
    Set oFolder = Nothing
    
End Sub

Public Property Get Parent_DMIS_ID() As String
    Parent_DMIS_ID = mParent_DMIS_ID
End Property

Public Property Let Parent_DMIS_ID(ByVal RHV As String)
    mParent_DMIS_ID = RHV
End Property

Public Property Get Folder_Name() As String
    Folder_Name = mFolder_Name
End Property

Public Property Let Folder_Name(ByVal RHV As String)
    mFolder_Name = RHV
End Property

Public Property Get Folder_NCID() As String
    Folder_NCID = mFolder_NCID
End Property

Public Property Let Folder_NCID(ByVal RHV As String)
    mFolder_NCID = RHV
End Property

Public Property Get SubFolders() As Collection
    If mSubFolders Is Nothing Then LoadSubFolders
    Set SubFolders = mSubFolders
End Property

Public Property Let SubFolders(ByVal RHV As Collection)
    Set mSubFolders = RHV
End Property

Public Property Get Folder_Type() As String
    Folder_Type = mFolder_Type
End Property

Public Property Let Folder_Type(ByVal RHV As String)
    mFolder_Type = RHV
End Property

Public Property Get Folder_Root() As String
    'SCR-44264 RBELL 12/23/03 Expose the owner value to determine the origin of folder
    Folder_Root = mFolder_Root
End Property

Public Property Let Folder_Root(ByVal RHV As String)
    'SCR-44264 RBELL 12/23/03 Allow the user to set the owner value to determine the origin of folder
    mFolder_Root = RHV
End Property
