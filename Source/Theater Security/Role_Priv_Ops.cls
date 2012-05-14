VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Role_Priv_Ops"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private oSQL As ICHCSII_SQL.ISqlOps
Private mcolRolePriv As Collection
Public colRoles As Collection
Public colPriv As Collection

'Private mobjSecOps As SecOps
'
'Public Property Set SecurityOps(oNewOps As SecOps)
'    mobjSecOps = oNewOps
'End Property

Private Function GetAllRolePrivileges(oSQL As ICHCSII_SQL.ISqlOps) As Boolean
    Dim oRP As Role_Priv
    Set mcolRolePriv = New Collection
    Dim oPriv As privilege
    Dim sRole As String
    
    sRole = ""
    oSQL.Execute "select role_name, priv_name, priv_string, priv_desc, role_desc, available_priv.priv_id, roles.role_id" _
     & " from roles, available_priv, role_priv" _
     & " where role_priv.role_id = roles.role_id" _
     & " and role_priv.priv_id = available_priv.priv_id" _
     & " order by role_id, priv_id"
    Do While Not oSQL.EOF
        If sRole <> oSQL("role_ID") Then
            Set oRP = New Role_Priv
            oRP.Name = oSQL("role_name")
            oRP.Desc = oSQL("role_desc")
            oRP.ID = oSQL("role_id")
            mcolRolePriv.Add oRP, CStr(oRP.ID)
            sRole = oSQL("role_ID")
        End If
        Set oPriv = New privilege
        oPriv.PrivName = oSQL("priv_name")
        oPriv.PrivID = oSQL("priv_id")
        oPriv.PrivDesc = oSQL("priv_desc")
        oPriv.PrivString = oSQL("priv_string")
        Call oRP.AddPrivilege(oPriv)
        oRP.bDirty = False
        oSQL.MoveNext
    Loop
    If Not oRP Is Nothing Then oRP.bDirty = False
    Set oRP = Nothing
    Set oPriv = Nothing
    
End Function

Private Sub Class_Initialize()
    Set oSQL = New GEMS_SQL.SqlOps
    Call GetAllRolePrivileges(oSQL)
    
End Sub
Public Function display() As Boolean
    Dim sClip As String
    Dim vThing As Variant
    Dim orolepriv As Role_Priv
    Dim orole As TheaterSecurity.Role
    Dim oPriv As TheaterSecurity.privilege
    Dim i As Long
    Dim j As Long
    
    
    frmRolePriv.bSysEvent = True
    With frmRolePriv.Grid
        .Cols = colRoles.count + 2
        .Col = 0
        .Row = 0
        If Not colRoles Is Nothing Then
        For Each orole In colRoles
            .Col = .Col + 1
            .ColData(.Col) = orole.RoleID
            .ColSel = .Col
            .RowSel = .Row + 1
            .Clip = orole.RoleDesc ' & vbCr & orole.RoleName
'            frmRolePriv.cboRole.AddItem orole.RoleName
'            frmRolePriv.cboRole.ItemData(frmRolePriv.cboRole.NewIndex) = orole.RoleID
        Next orole
        End If
        .Rows = colPriv.count + 2
        .Col = 0
        .Row = 0
        For Each oPriv In colPriv
            .Row = .Row + 1
            .RowData(.Row) = oPriv.PrivID
            .ColSel = .Col + 1
            .RowSel = .Row
            .Clip = oPriv.PrivName '& vbTab & Replace(opriv.PrivDesc, "(", Chr(13) & Chr(10) & "(")

'            frmRolePriv.cboPriv.AddItem oPriv.PrivName
'            frmRolePriv.cboPriv.ItemData(frmRolePriv.cboPriv.NewIndex) = oPriv.PrivID
        Next oPriv
        
        For i = 1 To .Cols - 1
            For j = 1 To .Rows - 1
                .TextMatrix(j, i) = GetPrivStr(.RowData(j), .ColData(i))
            Next j
        Next i
        
        DoEvents
        .AutoSize 0, .Cols - 1
        
        .Row = 1
        .Col = 1
'        If frmRolePriv.cboPriv.ListCount > 0 Then
'            frmRolePriv.cboPriv.ListIndex = 0
'        End If
'        If frmRolePriv.cboRole.ListCount > 0 Then
'            frmRolePriv.cboRole.ListIndex = 0
'        End If
    End With
    frmRolePriv.bSysEvent = False
    
    
    frmRolePriv.Show
End Function

Private Function GetPrivStr(ByVal lRoleID As Long, ByVal lPrivID As Long) As String
  Dim oRP As TheaterSecurity.Role_Priv
  Dim oPriv As TheaterSecurity.privilege
  On Error Resume Next
  Set oRP = mcolRolePriv(CStr(lRoleID))
  If Not oRP Is Nothing Then
    Set oPriv = oRP.Privileges(CStr(lPrivID))
    If Err Then
        Err.Clear
        Exit Function
    End If
    If Not oPriv Is Nothing Then
        GetPrivStr = oPriv.PrivString
    End If
  End If
    
  
  'Next oRP
  
End Function
