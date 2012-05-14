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
Attribute VB_Name = "SecOps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum tsPrivAccess
    ts_Read = 1
    ts_Write = 2
    ts_Delete = 3
    ts_BreadGlass = 4
End Enum
    
Private mcolAllUsers      As Collection
Private mcolAllResources  As Collection
Private mcolAllPrivileges As Collection
Private mcolAllRoles      As Collection
Private mcolAllFacilities As Collection
Private mcolAllClinics    As Collection
Private mobjCrypto        As CCrypto
Private mbWrongFileType   As Boolean
Private maPrivData(1 To 300, 1 To 200) As String

Public Property Get Crypto() As CCrypto: Set Crypto = mobjCrypto: End Property

Private Function GetAllUsers(oSQL As SqlOps) As Boolean
    Dim oUser As TheaterSecurity.User
    Set mcolAllUsers = New Collection
    Dim oRes As User_Res
    Dim oFac As User_Fac
    Dim oPriv As privilege
    
        oSQL.Execute "select u.User_Ncid,u.domain,u.timer,ur.user_name " _
        & " from users u, user_resources ur, resources r" _
        & " where u.user_ncid = ur.user_ncid " _
        & " and ur.res_id =  r.res_id" _
        & " and r.resource_name = '" & THEATER_KEYWORD & "'"

    Do While Not oSQL.EOF
        Set oUser = New TheaterSecurity.User
        oUser.NCID = "" & oSQL("User_Ncid")
        oUser.Domain = "" & oSQL("Domain")
        oUser.Timer = "" & oSQL("Timer")
        oUser.UID = "" & oSQL("user_name")
        mcolAllUsers.Add oUser, oUser.NCID
        
        oSQL.MoveNext
    Loop
    Dim osql_ns As SqlOps
    Set osql_ns = GetAppConnection
    
    For Each oUser In mcolAllUsers
        Call GetAddlUserData(oUser, osql_ns)
        Call GetProvider(oUser, osql_ns)
        Call GetUserRes(oUser, oSQL)
        Call GetUserRole(oUser, oSQL)
        Call GetUserPrivilege(oUser, oSQL)
        Call GetUserFacility(oUser, osql_ns)
    Next oUser
End Function

Private Function GetAddlUserData(oUser As User, oSQL As SqlOps) As Boolean
    On Error GoTo ErrHandler
    oSQL.Execute "select h.User_Number , h.Unit_Number, h.Specialty_Ncid, h.logon_name" _
    & " from hems_users h" _
    & " where h.user_ncid = " & oUser.NCID
    If Not oSQL.EOF Then
        oUser.Unit_Number = "" & oSQL("unit_number")
        oUser.User_Number = "" & oSQL("user_number")
        oUser.SpecialyCode = IIf(IsNull(oSQL("Specialty_Ncid")), 0, oSQL("Specialty_Ncid"))
        oUser.UID = "" & oSQL("logon_name")
    Else
        '-- oops. this shouldn't happen, we're out of synch (secadmin.users vs gems.hems_users)
                    
    End If
    If Val(oUser.Unit_Number) > 0 Then
    oSQL.Execute "select last_name, first_name" _
    & " from mmi_name" _
    & " where unit_number = " & oUser.Unit_Number
    If Not oSQL.EOF Then
        oUser.LastName = "" & oSQL("last_name")
        oUser.FirstName = "" & oSQL("first_name")
    End If
    End If
    GetAddlUserData = True
ErrHandler:
End Function

Private Function GetProvider(oUser As User, oSQL As SqlOps) As Boolean
    On Error GoTo ErrHandler

    oSQL.Execute "select pcm_code, ien, Name, Duty_Phone_1, Duty_Phone_2, SSN, specialty_code, EMail" _
    & " from provider" _
    & " where ncid = " & oUser.NCID
    
    'SCR 101231 - LCI - 09/20/2006
    If Not oSQL.EOF Then
        oUser.PCM_Code = "" & oSQL("pcm_code")
        oUser.IEN = "" & oSQL("ien")
        oUser.ProviderName = "" & oSQL("name")
        oUser.SSN = oSQL("SSN")
        'SCR 171931 - Duty_Phone_1 is the workphone when it is inserted
        oUser.HomePhone = "" & oSQL("Duty_Phone_2")
        oUser.WorkPhone = "" & oSQL("Duty_Phone_1")
        oUser.EMail = "" & oSQL("EMail")
        If oUser.SpecialyCode = 0 Then
            oUser.SpecialyCode = IIf(IsNull(oSQL("Specialty_code")), 0, oSQL("Specialty_code"))
        End If
        GetProvider = True
    Else
        oUser.ProviderName = oUser.UID
        GetProvider = False
    End If
    
    Exit Function
ErrHandler:
    MsgBox "The following error ocurred in SecOps.GetProvider:" & vbCrLf & "Error Number: " & Err.number & vbCrLf & "Error Description: " & Err.Description
    
End Function

Private Function GetAllFacilities(oSQL As SqlOps) As Boolean
    On Error GoTo haveerror
    Dim oFac As Faciltiy
    Set mcolAllFacilities = New Collection
    
    oSQL.Execute "select DISTINCT FACILITY_NCID as fncid, CONTEXT_NCID as cncid, Type, NCIDS.DESCRIPTION as Name " & _
                  " from FACILITY_NCID_Correlation, ncids " & _
                  " WHERE ncids.NCID = FACILITY_NCID"
    
    Do While Not oSQL.EOF
    
        Set oFac = New Faciltiy
        oFac.FacName = oSQL("Name")
        oFac.FacNcid = oSQL("fncid")
        oFac.FacContextNcid = oSQL("cncid")
        oFac.Factype = oSQL("Type")
        mcolAllFacilities.Add oFac, CStr(oSQL("fncid"))
        Set oFac = Nothing
        oSQL.MoveNext
    Loop
    
    GetAllFacilities = True
    
haveerror:
        
End Function
Private Function GetAllClinics(oSQL As SqlOps) As Boolean
 
 On Local Error GoTo haveerror
    Dim oclinic As clinic
    Set mcolAllClinics = New Collection
    
    Dim sSQL As String
    sSQL = "Select distinct ncid, Name from Clinic "
    
    If Not gsThisFacility_NCID = "" Then
        sSQL = sSQL & " WHERE facility_ncid = " & gsThisFacility_NCID
    End If
    sSQL = sSQL & " Order by name"
    oSQL.Execute sSQL
    Do While Not oSQL.EOF
        Set oclinic = New clinic
        oclinic.clinicName = oSQL("Name")
        oclinic.FacilityName = gsThisFacility_Name
        oclinic.FacNcid = gsThisFacility_NCID
        oclinic.ClinicNCID = CStr(oSQL("ncid"))
        mcolAllClinics.Add oclinic, oclinic.ClinicNCID
        Set oclinic = Nothing
        oSQL.MoveNext
    Loop
    GetAllClinics = True

haveerror:

End Function

Private Function GetAllRoles(oSQL As SqlOps) As Boolean
  
    Dim orole As TheaterSecurity.role
    Set mcolAllRoles = New Collection
    
    oSQL.Execute "select * from roles order by role_name "
    Do While Not oSQL.EOF
        Set orole = New TheaterSecurity.role
        orole.RoleDesc = oSQL("role_desc")
        orole.roleID = oSQL("role_id")
        orole.RoleName = oSQL("role_name")
        mcolAllRoles.Add orole, CStr(orole.roleID)
        Set orole = Nothing
        oSQL.MoveNext
    Loop

End Function

Public Sub Logoff()

    '89368 SF 3/5/06
    'There is error handling in TMIPAudit and above this sub
    TMIPAudit "SECURITY", "CHCS II-T SECURITY", gsUserID, "", "Successful logout", "", oSQL.moCnn

End Sub

Public Function ReadRolesFromSpreadSheet(ByVal sFileName As String) As Boolean
    Dim Col             As Long
    Dim Row             As Long
    Dim sLine           As String
    Dim iTotalPrivs     As Integer
    Dim iTotalRoles     As Integer
    Dim iPrivRowStart   As Integer
    Dim iPrivNameCol    As Integer
    Dim iPrivDescCol    As Integer
    Dim iRoleNameRow    As Integer
    Dim iRoleDescRow    As Integer
    Dim iRoleColStart   As Integer
    Dim vUR             As Variant
    Dim oPriv           As privilege
    Dim oPriv2          As privilege
    Dim oPrivNew        As privilege
    Dim sPrivName       As String
    Dim sROleName       As String
    Dim PrivID          As Long
    Dim roleID          As Long
    Dim orole           As TheaterSecurity.role
    Dim oRole2          As TheaterSecurity.role
    Dim oRoleNew        As TheaterSecurity.role
    Dim bFoundIt        As Boolean
    Dim colAllPrivileges As Collection
    Dim colAllRoles     As Collection
    Dim colUserRoles    As Collection
    
    Set colAllRoles = New Collection
    Set colAllPrivileges = New Collection
    Set colUserRoles = New Collection
    gbSSProcCancel = False
    On Error GoTo ErrHandler
    
    Dim oSQL As SqlOps
    Set oSQL = GetSecConnection
    
'___ get the current user roles and privileges
    If MsgBox("This action will delete all current Roles, Functions, and Privileges and replace them with information from the spreadsheet: " & sFileName & "." _
    & vbCrLf & vbCrLf & "Continue?", vbExclamation + vbYesNo, "Replace Data") = vbNo Then Exit Function

    Screen.MousePointer = vbHourglass
    
    GetRolePrivData sFileName
    If mbWrongFileType = True Then
        mbWrongFileType = False
        Exit Function
    End If
    
    ResetfrmSSRead
    
    oSQL.Execute "select max(priv_id) from available_priv"
    PrivID = oSQL(0)
    oSQL.Execute "select max(role_id) from roles"
    roleID = oSQL(0)


'___ get all the privileges
'   if this is also in the Old set, use that priv ID and mark as "deleted" in the new set
    iTotalPrivs = GetTotalPrivileges
    iPrivNameCol = 2
    iPrivDescCol = 3
    iPrivRowStart = 4
    frmSSRead.pbPrivs.Min = iPrivRowStart
    frmSSRead.pbPrivs.Max = iTotalPrivs + iPrivRowStart
    For Row = iPrivRowStart To iTotalPrivs + iPrivRowStart
        frmSSRead.pbPrivs.Value = Row
        DoEvents
        If gbSSProcCancel Then GoTo Cleanup
        bFoundIt = False
        Set oPriv = New privilege
        oPriv.PrivName = maPrivData(Row, iPrivNameCol)
        oPriv.PrivDesc = maPrivData(Row, iPrivDescCol)

        '-- does this priv exist?  if so use the old priv id
        For Each oPriv2 In mcolAllPrivileges
            DoEvents
            If gbSSProcCancel Then GoTo Cleanup
            If LCase$(oPriv2.PrivName) = LCase$(oPriv.PrivName) Then
                oPriv.PrivID = oPriv2.PrivID
                oPriv.bRedundant = True   '-- don't need it
                bFoundIt = True
                If Not oPriv2.PrivDesc = oPriv.PrivDesc Then
                    oPriv2.PrivDesc = oPriv.PrivDesc
                    oPriv2.bChanged = True
                End If
                Exit For
            End If
        Next oPriv2

        If Not bFoundIt Then
            PrivID = PrivID + 1
            oPriv.PrivID = PrivID
        End If
        
        If LenB(oPriv.PrivName) > 0 Then   '--- SCR 72007  Sherry Wang   7/11/2005
            colAllPrivileges.Add oPriv, oPriv.PrivName
        End If
        Set oPriv = Nothing

    Next Row
'
'___ get all the roles
    frmSSImport.grdRolePriv.Rows = 1
    iRoleNameRow = 2
    iRoleDescRow = 3
    iRoleColStart = 4
    iTotalRoles = GetTotalRoles
'    if this role is in the Old set, use that role ID and mark as "deleted" in the new set
    frmSSRead.pbRoles.Min = iRoleColStart
    frmSSRead.pbRoles.Max = iTotalRoles + iRoleColStart
    For Col = iRoleColStart To iTotalRoles + iRoleColStart
        frmSSRead.pbRoles.Value = Col
        DoEvents
        If gbSSProcCancel Then GoTo Cleanup
        Set orole = New TheaterSecurity.role
        orole.RoleDesc = maPrivData(iRoleDescRow, Col)
        orole.RoleName = maPrivData(iRoleNameRow, Col)
        
        '--does this role exits?  if so use the old role id
        For Each oRole2 In mcolAllRoles
            bFoundIt = False
            DoEvents
            If gbSSProcCancel Then GoTo Cleanup
            If LCase$(oRole2.RoleName) = LCase$(orole.RoleName) Then
                orole.roleID = oRole2.roleID
                If Not oRole2.RoleDesc = orole.RoleDesc Then
                    oRole2.RoleDesc = orole.RoleDesc
                    oRole2.bChanged = True
                End If
                bFoundIt = True
                orole.bRedundant = True   '-- don't need it
                Exit For
            End If
        Next oRole2

        If Not bFoundIt Then
            roleID = roleID + 1
            orole.roleID = roleID
        End If
        If LenB(orole.RoleName) > 0 Then  '--- SCR 72007  Sherry Wang   7/11/2005
            colAllRoles.Add orole, orole.RoleName
        End If

    '___ get all the privileges for this role
        frmSSRead.pbRolePrivs.Min = iPrivRowStart
        frmSSRead.pbRolePrivs.Max = iTotalPrivs + iPrivRowStart
        For Row = iPrivRowStart To iTotalPrivs + iPrivRowStart
            frmSSRead.pbRolePrivs.Value = Row
            DoEvents
            If gbSSProcCancel Then GoTo Cleanup
            Set oPriv = New privilege
            oPriv.PrivName = maPrivData(Row, iPrivNameCol)
            oPriv.PrivID = colAllPrivileges(oPriv.PrivName).PrivID
            oPriv.PrivString = maPrivData(Row, Col)
             '-- see if this old role has this priv. if so, update string, else add
            On Error Resume Next
'            Set oRole2 = mcolAllRoles(CStr(orole.roleID))
            Set oPriv2 = mcolAllRoles(CStr(orole.roleID)).Privileges(CStr(oPriv.PrivID))

            If Err.number <> 0 Then
                If oPriv.PrivString <> "" Then
                    orole.Privileges.Add oPriv, oPriv.PrivName
                    frmSSImport.grdRolePriv.AddItem "New" & vbTab & orole.RoleName & vbTab & oPriv.PrivName & vbTab & oPriv.PrivString
                End If
                Err.Clear

            Else
                If Not oPriv2.SamePriv(oPriv.PrivString) Then
                    frmSSImport.grdRolePriv.AddItem "Changed" & vbTab & orole.RoleName & vbTab & oPriv.PrivName & vbTab & oPriv.PrivString
                    oPriv2.PrivString = oPriv.PrivString
                    oPriv2.bChanged = True
                End If
            End If
            Set oPriv2 = Nothing
            Set oPriv = Nothing
        Next Row
        Set orole = Nothing

    Next Col

'___ confirm changes
     '__ find deleted roles: go through all mcolallRoles and if not in colallroles, mark for del
     '                                                   and if IN colallroles, del from colcallroles
     Unload frmSSRead

     For Each orole In mcolAllRoles
        orole.bDeleted = True
        For Each oRoleNew In colAllRoles
            If oRoleNew.RoleName = orole.RoleName Then
                orole.bDeleted = False
                oRoleNew.bRedundant = True
            End If
        Next oRoleNew
     Next orole


     With frmSSImport
         '-- show deleted roles from the 'old' pile
         .grdRoleDel.Rows = 1
         For Each orole In mcolAllRoles
            If orole.bDeleted Then
                .grdRoleDel.AddItem True & vbTab & orole.RoleName
                .grdRoleDel.RowData(.grdRoleDel.Rows - 1) = orole.roleID
            End If
         Next orole


         '-- show new roles (ones that are not 'deleted' from the 'new' pile
         .grdRoleAdd.Rows = 1
         For Each orole In colAllRoles
            If Not orole.bRedundant And Not orole.bDeleted Then
                .grdRoleAdd.AddItem True & vbTab & orole.RoleName
                .grdRoleAdd.RowData(.grdRoleAdd.Rows - 1) = orole.roleID
            End If
         Next orole


    '___ now the privileges
         For Each oPriv In mcolAllPrivileges
            oPriv.bDeleted = True
            For Each oPrivNew In colAllPrivileges
                If oPrivNew.PrivName = oPriv.PrivName Then
                    oPriv.bDeleted = False
                    oPrivNew.bRedundant = True
                    'If oPriv.PrivID > 124 Or oPrivNew.PrivID > 124 Then Stop
                End If
            Next oPrivNew
         Next oPriv

         '-- show deleted privileges from the 'old' pile
         .grdFunctDel.Rows = 1
         For Each oPriv In mcolAllPrivileges
            If oPriv.bDeleted Then
                .grdFunctDel.AddItem True & vbTab & oPriv.PrivName
                .grdFunctDel.RowData(.grdFunctDel.Rows - 1) = oPriv.PrivID
            End If
         Next oPriv

         '-- show new privileges (ones that are not 'deleted' from the 'new' pile
         .grdFunctAdd.Rows = 1
         For Each oPriv In colAllPrivileges
            If Not oPriv.bDeleted And Not oPriv.bRedundant Then
                .grdFunctAdd.AddItem True & vbTab & oPriv.PrivName
                .grdFunctAdd.RowData(.grdFunctAdd.Rows - 1) = oPriv.PrivID
            End If
         Next oPriv


     End With
     Unload frmSSRead
     Screen.MousePointer = vbDefault
     frmSSImport.Show vbModal

     If Not frmSSImport.bUserSelect Then Exit Function
     Screen.MousePointer = vbHourglass
'- sync the objects to the grid.
    '--any grid row without a check is a user change.
    Call UpdateCollection(frmSSImport.grdRoleDel, mcolAllRoles, False)
    Call UpdateCollection(frmSSImport.grdRoleAdd, colAllRoles, True)
    Call UpdateCollection(frmSSImport.grdFunctDel, mcolAllPrivileges, False)
    Call UpdateCollection(frmSSImport.grdFunctAdd, colAllPrivileges, True)

    Dim colSQL As Collection
    Set colSQL = New Collection
    Dim sIDList As String

'__ Delete the old ones we should delete
sIDList = ""
    For Each orole In mcolAllRoles
        If orole.bDeleted Then
            sIDList = sIDList & "," & orole.roleID
        End If
    Next orole
    If Len(sIDList) > 0 Then
        sIDList = Mid$(sIDList, 2) '-- drop leading comma
        colSQL.Add "Delete from user_roles where role_id in (" & sIDList & ")"
        colSQL.Add "delete from role_priv where role_id in (" & sIDList & ")"
        colSQL.Add "delete from roles where role_id in (" & sIDList & ")"
    End If
sIDList = ""
    For Each oPriv In mcolAllPrivileges
        If oPriv.bDeleted Then
            sIDList = sIDList & "," & oPriv.PrivID
        End If
    Next oPriv
    If Len(sIDList) > 0 Then
        sIDList = Mid$(sIDList, 2) '-- drop leading comma
        colSQL.Add "delete from role_priv where priv_id in (" & sIDList & ")"
        colSQL.Add "delete from available_priv where priv_id in (" & sIDList & ")"
    End If

'__ add the new ones we should add

    For Each oPriv In colAllPrivileges
        If Not oPriv.bDeleted And Not oPriv.bRedundant Then
            colSQL.Add "Insert into available_priv (Priv_ID, Priv_NAME, Priv_DESC) values ( " _
                & oPriv.PrivID & "," & SQLQuote(oPriv.PrivName) & "," & SQLQuote(oPriv.PrivDesc) & ") "
        End If
    Next oPriv

    For Each orole In colAllRoles
        If Not orole.bDeleted And Not orole.bRedundant Then
            colSQL.Add "Insert into roles (ROLE_ID, ROLE_NAME, ROLE_DESC) values ( " _
                    & orole.roleID & ",'" & orole.RoleName & "'," & SQLQuote(orole.RoleDesc) & ") "
        End If
            For Each oPriv In orole.Privileges
            '__ the role priv is not the same as the real priv which may be deleted
'                If oPriv.PrivID > 125 Then Stop
                Set oPriv2 = colAllPrivileges(oPriv.PrivName)
                If oPriv.PrivString <> "" And Not (oPriv.bDeleted Or oPriv2.bDeleted Or oPriv.bRedundant) Then
                    colSQL.Add "insert into role_priv(role_id, priv_id,Priv_string) values ( " _
                        & orole.roleID & "," & oPriv2.PrivID & ",'" & oPriv.PrivString & "')"
                End If
            Next oPriv

    Next orole
    '__ now batch process the sqls
    oSQL.ExecuteMulti colSQL
    DoEvents
    '__ do a reload of all information
    MsgBox ("Please restart this application to view new data.")
    GoTo Cleanup
    
ErrHandler:
    MsgBox "The following error ocurred:" & vbCrLf & "Error Number: " & Err.number & vbCrLf & "Error Description: " & Err.Description & vbCrLf & "The import has been terminated."
    
Cleanup:
    Unload frmSSRead
    ReadRolesFromSpreadSheet = True
    Screen.MousePointer = vbDefault
End Function


Private Sub UpdateCollection(grd As VSFlexGrid, Col As Collection, bValue As Boolean)
'__ just by coincidence, if 'value' then colkey is 'name' not 'id'
    Dim i As Long
    With grd
        For i = 1 To .Rows - 1
           ' If Col(.TextMatrix(i, 1)).PrivID > 124 Then Stop
            If .TextMatrix(i, 0) <> "True" Then
                '-- user said "don't delete"
                If bValue Then
                    Col(.TextMatrix(i, 1)).bDeleted = bValue
                Else
                    Col(CStr(.RowData(i))).bDeleted = bValue
                End If
            End If
        Next i
    End With

End Sub




Private Function GetAllPrivileges(oSQL As SqlOps) As Boolean
    Dim oPriv As privilege
    Set mcolAllPrivileges = New Collection
    
    oSQL.Execute "select * from available_priv order by priv_name"
    Do While Not oSQL.EOF
        Set oPriv = New privilege
        oPriv.PrivName = oSQL("priv_name")
        oPriv.PrivID = oSQL("priv_id")
        oPriv.PrivDesc = oSQL("priv_desc")
        mcolAllPrivileges.Add oPriv, CStr(oPriv.PrivID)
        Set oPriv = Nothing
        oSQL.MoveNext
    Loop

End Function

Private Function GetAllResources(oSQL As SqlOps) As Boolean
    Dim oRes As resource
    Set mcolAllResources = New Collection
    oSQL.Execute "select * from resources order by resource_name"
    Do While Not oSQL.EOF
        Set oRes = New resource
        oRes.resName = oSQL("resource_name")
        oRes.resId = oSQL("res_id")
        mcolAllResources.Add oRes, CStr(oRes.resId)
        oSQL.MoveNext
    Loop

End Function

Private Function GetAllRolePriv(oSQL As SqlOps) As Boolean
    Dim oRol As role
    Dim oPriv As privilege
    For Each oRol In mcolAllRoles
        Set oRol.Privileges = New Collection
        oSQL.Execute "Select r.priv_id, p.priv_name, p.priv_desc, r.priv_string from role_priv r, available_priv p where r.priv_id = p.priv_id and r.role_id = " & oRol.roleID
        Do While Not oSQL.EOF
            Set oPriv = New privilege
            oPriv.PrivID = oSQL("priv_id")
            oPriv.PrivName = oSQL("priv_name")
            oPriv.PrivString = oSQL("priv_string")
            oPriv.PrivDesc = oSQL("priv_desc")
            oRol.Privileges.Add oPriv, CStr(oPriv.PrivID)
            Set oPriv = Nothing
            oSQL.MoveNext
        Loop
    Next oRol
End Function


Public Sub Initialize()
    

    frmSplash.lblStatus.Caption = "Loading Users..."
    frmSplash.lblStatus.Refresh
        Call GetAllUsers(oSQL)

    frmSplash.lblStatus.Caption = "Loading Roles..."
    frmSplash.lblStatus.Refresh
        Call GetAllRoles(oSQL)

    frmSplash.lblStatus.Caption = "Loading Resources..."
    frmSplash.lblStatus.Refresh
        Call GetAllResources(oSQL)

    frmSplash.lblStatus.Caption = "Loading Functions..."
    frmSplash.lblStatus.Refresh
        Call GetAllPrivileges(oSQL)

    frmSplash.lblStatus.Caption = "Loading Privileges..."
    frmSplash.lblStatus.Refresh
        Call GetAllRolePriv(oSQL)
 
    frmSplash.lblStatus.Caption = "Loading Facilities..."
    frmSplash.lblStatus.Refresh
        Call GetAllFacilities(osql_ns)
    
    frmSplash.lblStatus.Caption = "Loading Clinics..."
    frmSplash.lblStatus.Refresh
        Call GetAllClinics(osql_ns)
        
    'Call GetAllRolePrivileges(oSQL)

    Set frmMain.cUsers = mcolAllUsers
    Set frmMain.cRol = mcolAllRoles
    Set frmMain.cFun = mcolAllPrivileges
    Set frmMain.cRes = mcolAllResources
    
        
    Unload frmSplash

End Sub

Public Function GetCollectionOfClinics() As Collection
    Set GetCollectionOfClinics = mcolAllClinics
End Function
Public Function GetCollectionOfFacilities() As Collection
    Set GetCollectionOfFacilities = mcolAllFacilities
End Function

Public Function GetCollectionOfroles() As Collection
    Set GetCollectionOfroles = mcolAllRoles
End Function

Public Function GetCollectionOfResources() As Collection
    Set GetCollectionOfResources = mcolAllResources
End Function

Public Function GetAUserCollection(ByVal UserNCID As Long) As User
    Set GetAUserCollection = mcolAllUsers.Item(UserNCID)
End Function

Public Function SaveUser(oUser As User) As Boolean
    On Error GoTo ErrHandler

    Dim SQL  As String
    Dim sFac As String
    Dim sDom As String
    Dim LastName As String
    Dim FirstName As String
    
    With oUser
        
        Call UpdateC2TUser_sp(oUser)
        
        sDom = THEATER_KEYWORD
        If .NewUser Then
            On Error Resume Next  ' - may have already added to hems, provider
        
            SQL = "Insert into users (homephone, workphone, email, timer, user_ncid,domain) values (" & _
                  SQLQuote(.HomePhone) & ", " & _
                  SQLQuote(.WorkPhone) & ", " & _
                  SQLQuote(.EMail) & ", " & _
                           .Timer & ", " & _
                           .NCID & "," & _
                           SQLQuote(sDom) & ")"  '2/20/02 - Walker THEATER_KEYWORD & ")"
            oSQL.Execute SQL
            Dim oRes As User_Res
            For Each oRes In .Resources
                With oRes
            SQL = "Insert into user_resources (res_id, user_ncid,user_name, user_password,locked) " _
                    & "values( " & SQLQuote(.ResourceID) & "," & SQLQuote(.User_NCID) & "," _
                    & SQLQuote(.User_Name) & "," & SQLQuote(.Password) & ",0)"
                End With
            
            oSQL.Execute SQL
            Next oRes
            
            If mcolAllUsers Is Nothing Then Set mcolAllUsers = New Collection
            mcolAllUsers.Add oUser, CStr(oUser.NCID)
            
            
        Else  '--- SCR 72585  Sherry Wang   6/20/2005
            SQL = "update users " & vbNewLine
            SQL = SQL & "set homephone = " & SQLQuote(.HomePhone) & ", " & vbNewLine
            SQL = SQL & "workphone = " & SQLQuote(.WorkPhone) & ", " & vbNewLine
            SQL = SQL & "email = " & SQLQuote(.EMail) & ", " & vbNewLine
            SQL = SQL & "timer = " & SQLQuote(.Timer) & vbNewLine
            SQL = SQL & "where user_ncid = " & .NCID & vbNewLine
            
            oSQL.Execute SQL
            
            If .NCID = gsUserNCID Then
                glUserTimer = CLng(.Timer)
                frmMain.ResetTimeoutTimer
            End If
        End If
    
    End With
    
    SaveUser = True
    
    Exit Function
    
ErrHandler:
    MsgBox "Error occured in SecOps.SaveUser: " & Err.Description
End Function


Private Function UpdateC2TUser_sp(oUser As User)
    On Error GoTo ErrHandler
    Dim sSQL As String
    
        'create or replace procedure C2T_Security_User (
            'vNCID         varchar2,
            'vClinicNCID   varchar2,
            'vHomePhone    varchar2,
            'vWorkPhone    varchar2,
            'vEMail        varchar2,
            'vDomain       varchar2,
            'vTimer        varchar2,
            'vProviderName varchar2,
            'vLastName     varchar2,
            'vFirstName    varchar2,
            'vSpecialtyCode varchar2,
            'vSSN          varchar2,
            'vFacility     varchar2,
            'vUnit_Number  varchar2,
            'vUser_Number  varchar2,
            'vPCM_Code     varchar2,
            'vIEN          varchar2,
            'vUID          varchar2)
    '--- SCR SCR 88068, 88070  Sherry Wang 2/15/2006
    If CWShared.UseMySQLDB Then
        sSQL = "CALL C2T_USER_UPDATE (" _
            & Val(oUser.NCID) & ", " _
            & "'" & oUser.ClinicNCID & "', " _
            & "'" & oUser.HomePhone & "', " _
            & "'" & oUser.WorkPhone & "', " _
            & "'" & oUser.EMail & "', " _
            & "'" & oUser.Domain & "', " _
            & "'" & oUser.Timer & "', " _
            & "'" & oUser.ProviderName & "', " _
            & "'" & oUser.LastName & "', " _
            & "'" & oUser.FirstName & "', " _
            & "'" & oUser.SpecialyCode & "', " _
            & "'" & oUser.SSN & "', " _
            & "'" & oUser.Facilities(1).FacilityNCID & "', " _
            & "" & Val(oUser.Unit_Number) & ", " _
            & "" & Val(oUser.User_Number) & ", " _
            & "'" & oUser.PCM_Code & "', " _
            & "" & Val(oUser.IEN) & ", " _
            & "'" & oUser.UID & "')"
            osql_ns.moCnn.Execute sSQL
            
    Else
        Dim Cmd As ADODB.Command
        Set Cmd = New ADODB.Command
        Dim Param As ADODB.Parameter
        Set Param = Cmd.CreateParameter("ncid", adNumeric, adParamInput, 20):           Param.Value = Val(oUser.NCID):           Cmd.Parameters.Append Param
        Set Param = Cmd.CreateParameter("ClinicNCID", adNumeric, adParamInput, 20):     Param.Value = oUser.ClinicNCID:     Cmd.Parameters.Append Param
        Set Param = Cmd.CreateParameter("HomePhone", adVarChar, adParamInput, 16):      Param.Value = oUser.HomePhone:      Cmd.Parameters.Append Param
        Set Param = Cmd.CreateParameter("WorkPhone", adVarChar, adParamInput, 16):      Param.Value = oUser.WorkPhone:      Cmd.Parameters.Append Param
        Set Param = Cmd.CreateParameter("EMail", adVarChar, adParamInput, 100):         Param.Value = oUser.EMail:          Cmd.Parameters.Append Param
        Set Param = Cmd.CreateParameter("Domain", adVarChar, adParamInput, 16):         Param.Value = oUser.Domain:         Cmd.Parameters.Append Param
        Set Param = Cmd.CreateParameter("Timer", adNumeric, adParamInput, 16):          Param.Value = oUser.Timer:          Cmd.Parameters.Append Param
        Set Param = Cmd.CreateParameter("ProviderName", adVarChar, adParamInput, 100):   Param.Value = oUser.ProviderName:   Cmd.Parameters.Append Param
        Set Param = Cmd.CreateParameter("LastName", adVarChar, adParamInput, 26):       Param.Value = oUser.LastName:       Cmd.Parameters.Append Param
        Set Param = Cmd.CreateParameter("firstName", adVarChar, adParamInput, 16):      Param.Value = oUser.FirstName:      Cmd.Parameters.Append Param
        Set Param = Cmd.CreateParameter("SpecialtyCode", adNumeric, adParamInput, 20):  Param.Value = oUser.SpecialyCode:   Cmd.Parameters.Append Param
        Set Param = Cmd.CreateParameter("SSN", adVarChar, adParamInput, 16):            Param.Value = oUser.SSN:            Cmd.Parameters.Append Param
        Set Param = Cmd.CreateParameter("Facility", adNumeric, adParamInput, 20):       Param.Value = oUser.Facilities(1).FacilityNCID:  Cmd.Parameters.Append Param
        Set Param = Cmd.CreateParameter("Unit_Number", adNumeric, adParamInput, 20):    Param.Value = Val(oUser.Unit_Number):    Cmd.Parameters.Append Param
        Set Param = Cmd.CreateParameter("User_Number", adNumeric, adParamInput, 20):    Param.Value = Val(oUser.User_Number):    Cmd.Parameters.Append Param
        Set Param = Cmd.CreateParameter("PCM_Code", adVarChar, adParamInput, 16):       Param.Value = oUser.PCM_Code:       Cmd.Parameters.Append Param
        Set Param = Cmd.CreateParameter("IEN", adNumeric, adParamInput, 16):            Param.Value = Val(oUser.IEN):            Cmd.Parameters.Append Param
        Set Param = Cmd.CreateParameter("UID", adVarChar, adParamInput, 16):            Param.Value = oUser.UID:            Cmd.Parameters.Append Param
        Set Cmd.ActiveConnection = osql_ns.moCnn
        Cmd.CommandType = adCmdStoredProc
        Cmd.CommandText = "C2T_USER_UPDATE"
        Cmd.Execute
    End If
    
    Exit Function
    
ErrHandler:
    MsgBox "Error occured in SecOps.UpdateC2TUser_sp: " & Err.Description
    Exit Function
    Resume
End Function

Public Function SaveRoles(Optional orole As role) As Boolean
    On Error GoTo ErrHandler
    If orole Is Nothing Then
        For Each orole In mcolAllRoles
            SaveRole orole
        Next orole
    Else
        SaveRole orole
    End If
    SaveRoles = True
ErrHandler:

End Function

Private Function SaveRole(orole As role) As Boolean
    
    On Error GoTo ErrHandler
        
    If orole.roleID = 0 Then
        oSQL.Execute "select max(role_id) Max_ID from roles"
        orole.roleID = 1 + ("" & oSQL("max_id"))
        oSQL.Execute "insert into roles (role_id, role_name, role_desc) values (" _
        & orole.roleID & ",'" & orole.RoleName & "','" & orole.RoleDesc & "')"
    Else
        oSQL.Execute "Update roles set role_name = '" & orole.RoleName _
           & "', role_desc = '" & orole.RoleDesc _
           & "' where role_id = " & orole.roleID
    End If
        
    SaveRole = True
    
ErrHandler:

End Function


Public Function SaveResources(Optional oRes As resource) As Boolean
    On Error GoTo ErrHandler
    If oRes Is Nothing Then
        For Each oRes In mcolAllResources
            SaveResource oRes
        Next oRes
    Else
        SaveResource oRes
    End If
    SaveResources = True
ErrHandler:

End Function

Private Function SaveResource(oRes As resource) As Boolean
    
    On Error GoTo ErrHandler
        
    If oRes.resId = 0 Then
        oSQL.Execute "select max(oRes_id) Max_ID from resources"
        oRes.resId = 1 + ("" & oSQL("max_id"))
        oSQL.Execute "insert into resources (res_id, res_name) values (" _
        & oRes.resId & ",'" & oRes.resName & "')"
    Else
        oSQL.Execute "Update resources set res_name = '" & oRes.resName _
           & "' where res_id = " & oRes.resId
    End If
        
    SaveResource = True
    
ErrHandler:
    
End Function


Public Function SaveFunctions(Optional oFun As privilege) As Boolean
    On Error GoTo ErrHandler
    If oFun Is Nothing Then
        For Each oFun In mcolAllPrivileges
            SaveFunc oFun
        Next oFun
    Else
        SaveFunc oFun
    End If
    SaveFunctions = True
ErrHandler:

End Function

Private Function SaveFunc(oFun As privilege) As Boolean
    
    On Error GoTo ErrHandler
        
    If oFun.PrivID = 0 Then
        oSQL.Execute "select max(priv_id) Max_ID from available_priv"
        oFun.PrivID = 1 + ("" & oSQL("max_id"))
        oSQL.Execute "insert into available_priv (priv_id, priv_name, priv_desc) values (" _
        & oFun.PrivID & "," & SQLQuote(oFun.PrivName) & "," & SQLQuote(oFun.PrivDesc) & ")"
    Else
        oSQL.Execute "Update available_priv set priv_name = " & SQLQuote(oFun.PrivName) _
           & ", priv_desc = " & SQLQuote(oFun.PrivDesc) _
           & " where priv_id = " & oFun.PrivID
    End If
        
    SaveFunc = True
    
ErrHandler:
    

End Function

Public Function SaveResourceAttrib(oResAttrib As Res_Atrrib) As Boolean
    Dim oAttr As KeyValuePair
    On Error GoTo ErrHandler
        oSQL.Execute "delete from resource_attributes where res_id = " & oResAttrib.ResourceID
        
        For Each oAttr In oResAttrib.Attributes
          oSQL.Execute "insert into resource_attributes (res_id,key,value) values (" _
            & oResAttrib.ResourceID & "," & SQLQuote(oAttr.Key) & "," & SQLQuote(oAttr.Value) & ")"
        Next oAttr
        
    SaveResourceAttrib = True
    
ErrHandler:

End Function

Public Function DeleteUser(oUser As User) As Boolean
    Dim mCol As New Collection
    Dim SQL As String
    
'11/13/01 - Walker - SCR 18215 - Functionality moved to cmdUserDel for user message display
'11/13/01    If oUser.ProviderName = gsSysAdminName Then Exit Function
    
    SQL = "Delete from user_resources where user_ncid = " & oUser.NCID
    mCol.Add SQL
    SQL = "Delete from user_roles where user_ncid = " & oUser.NCID
    mCol.Add SQL
    SQL = "Delete from user_other_privileges where user_ncid = " & oUser.NCID
    mCol.Add SQL
    oSQL.ExecuteMulti mCol
    
    oSQL.Execute "Delete from users where user_ncid = " & oUser.NCID
    
    '--- SCR 88570  Sherry Wang  2/24/2006
    oSQL.Execute "update gems.hems_users set logon_name = NULL where user_ncid = " & oUser.NCID
        
    mcolAllUsers.Remove CStr(oUser.NCID)

    DeleteUser = True
    
End Function

Public Function DeleteRole(oRol As TheaterSecurity.role) As Boolean
    On Error GoTo ErrHandler
    
    If oRol Is Nothing Then Exit Function
    If oRol.RoleName = gsSysAdminRole Then Exit Function
    
    oSQL.Execute "select count(*) num from user_roles where role_id = " & oRol.roleID
    If Val(oSQL("num")) > 0 Then
        If MsgBox("Role '" & oRol.RoleName & "' is in use by " & oSQL("num") & " user(s).   Delete it anyway?", vbExclamation + vbYesNo, "In Use") = vbNo Then
            Exit Function
        Else
            oSQL.Execute "delete from user_roles where role_id = " & oRol.roleID
        End If
    End If
    oSQL.Execute "delete from role_priv where role_id = " & oRol.roleID
    oSQL.Execute "delete from roles where role_id = " & oRol.roleID
    mcolAllRoles.Remove CStr(oRol.roleID)
    
    DeleteRole = True
ErrHandler:
End Function

Public Function DeleteRes(oRes As Res_Atrrib) As Boolean
    On Error GoTo ErrHandler
    If oRes.ResourceName = THEATER_KEYWORD Then
        MsgBox "This is a Required Resource." & vbCrLf & "Unable to Delete '" & THEATER_KEYWORD & "'", vbExclamation
        Exit Function
    End If
    
    If oRes Is Nothing Then Exit Function
    
    oSQL.Execute "select count(*) num from user_resources where res_id = " & oRes.ResourceID
    If Val(oSQL("num")) > 0 Then
        If MsgBox("Resource '" & oRes.ResourceName & "' is in use by " & oSQL("num") & " user(s).   Delete it anyway?", vbExclamation + vbYesNo, "In Use") = vbNo Then
            Exit Function
        Else
            oSQL.Execute "delete from user_resources where res_id = " & oRes.ResourceID
        End If
        
        Exit Function
    End If
    oSQL.Execute "delete from resource_attributes where res_id = " & oRes.ResourceID
    oSQL.Execute "delete from resources where res_id = " & oRes.ResourceID
    
    DeleteRes = True
ErrHandler:
End Function

Public Function DeleteFun(oFun As privilege) As Boolean
    On Error GoTo ErrHandler
    
    
    If oFun Is Nothing Then Exit Function
    If oFun.PrivName = gsSysAdminFunc Then Exit Function
    
    oSQL.Execute "select count(*) num from role_priv where priv_id = " & oFun.PrivID
    If Val(oSQL("num")) > 0 Then
        If MsgBox("Function '" & oFun.PrivName & "' is in use by " & oSQL("num") & " role(s).   Delete it anyway?", vbExclamation + vbYesNo, "In Use") = vbNo Then
            Exit Function
        Else
            oSQL.Execute "delete from role_priv where priv_id = " & oFun.PrivID
        End If
    End If
    oSQL.Execute "delete from available_priv where priv_id = " & oFun.PrivID
    
    DeleteFun = True
ErrHandler:
End Function

Public Function SaveRolePriv(oRol As role) As Boolean
    
    Dim oPriv As privilege
    If oRol Is Nothing Then Exit Function
    On Error GoTo ErrHandler
    If oRol.RoleName = gsSysAdminRole Then Exit Function
    
    oSQL.Execute "delete from role_priv where role_id = " & oRol.roleID
    
    For Each oPriv In oRol.Privileges
        oSQL.Execute "insert into role_priv (role_id, priv_id, priv_string) values (" _
        & oRol.roleID & "," & oPriv.PrivID & ",'" & oPriv.PrivString & "')"
    Next oPriv
        
    SaveRolePriv = True
    
ErrHandler:
    

End Function

Public Function GetHEMSUser(ByVal sNCID As String) As User
    Dim oUser As User
    
    Dim oRes  As User_Res
    Dim oFac  As User_Fac
    Dim count As Integer
    Dim oPriv As privilege
    
    If Trim$(sNCID) = "" Then Exit Function
    
    osql_ns.Execute "select User_Number , Unit_Number, Specialty_Ncid, logon_name" _
    & " from hems_users" _
    & " where  user_ncid = " & sNCID
    If osql_ns.EOF Then
        Exit Function
    End If
    
    Set oUser = New TheaterSecurity.User
    oUser.NCID = sNCID
    oUser.Timer = 30
    oUser.Unit_Number = "" & osql_ns("unit_number")
    oUser.User_Number = "" & osql_ns("user_number")
    oUser.SpecialyCode = IIf(IsNull(osql_ns("Specialty_Ncid")), 0, osql_ns("Specialty_Ncid"))
    oUser.UID = "" & osql_ns("logon_name")

    Call GetProvider(oUser, osql_ns)
    Call GetUserRes(oUser, oSQL)
    Call GetUserRole(oUser, oSQL)
    Call GetUserPrivilege(oUser, oSQL)
    Call GetUserFacility(oUser, osql_ns)
    

    Set GetHEMSUser = oUser
    Set oUser = Nothing

End Function

Public Function GetResIDByName(ByVal sName As String) As Long
Dim oRes As resource

    For Each oRes In mcolAllResources
        If LCase$(oRes.resName) = LCase$(sName) Then
            GetResIDByName = oRes.resId
            Exit For
        End If
    Next oRes
End Function

Public Function AddOtherPriv() As Boolean
    
'    Call frmOther.Init(frmMain.CurrentUser, Me, mcolAllPrivileges, mcolAllRoles)
'    frmOther.Show vbModal
'    Call frmOtherPriv.Init(frmMain.CurrentUser, Me, mcolAllPrivileges, mcolAllRoles)
'    frmOtherPriv.Show vbModal
    
'    AddOtherPriv = True
End Function

Private Function GetUserRes(oUser As User, oSQL As SqlOps) As Boolean
    On Error GoTo ErrHandler
    Dim oRes As User_Res
    
    oSQL.Execute "select *" _
    & " from user_resources " _
    & " where user_ncid = " & oUser.NCID
    Do While Not oSQL.EOF
        Set oRes = New User_Res
        oRes.ResourceID = oSQL("res_ID")
        oRes.Password = Trim(oSQL("user_password"))
        oRes.User_Name = oSQL("user_name")
        oRes.User_NCID = oUser.NCID
        oRes.Locked = Val("" & oSQL("locked"))
        
        oUser.Resources.Add oRes, CStr(oRes.ResourceID)
        Set oRes = Nothing
        oSQL.MoveNext
    Loop
    GetUserRes = True
ErrHandler:
End Function

Private Function GetUserRole(oUser As User, oSQL As SqlOps) As Boolean
    
    oSQL.Execute "select role_id " _
    & " from user_roles " _
    & " where user_ncid = " & oUser.NCID
    Do While Not oSQL.EOF
        oUser.Roles.Add oSQL("role_id"), CStr(oSQL("role_id"))
        oSQL.MoveNext
    Loop

End Function

Private Function GetUserPrivilege(oUser As User, oSQL As SqlOps) As Boolean

    Dim oPriv As privilege
    oSQL.Execute "select u.priv_id , u.Add_Remove, u.Priv_string, a.priv_name, a.priv_desc" _
    & " from user_other_Privileges u, available_priv a " _
    & " where u.user_ncid = " & oUser.NCID _
    & " and u.priv_id = a.priv_id"
    Do While Not oSQL.EOF
        Set oPriv = New privilege
        oPriv.PrivID = "" & oSQL("priv_ID")
        oPriv.AddRemove = "" & oSQL("Add_Remove")
        oPriv.PrivString = "" & oSQL("Priv_string")
        oPriv.PrivName = "" & oSQL("priv_name")
        oPriv.PrivDesc = "" & oSQL("priv_desc")
        oUser.OtherPrivs.Add oPriv, CStr(oPriv.PrivID) & oPriv.AddRemove
        Set oPriv = Nothing
        oSQL.MoveNext
    Loop

End Function

Private Function GetUserFacility(oUser As User, oSQL As SqlOps) As Boolean
    Dim oFac As User_Fac
    
    '*-- 071202 SC: Added f.domainncid = '119' term since multiple clinic records were
    '*-- being erroneously returned.
'    oSql.Execute "select distinct facility_ncid, clinic_ncid, f.description as facility_name, c.description as clinic_name" _
'    & " from user_clinic u, ncids f, ncids c" _
'    & " where f.ncid = facility_ncid AND f.domainncid = '119'" _
'    & " and c.ncid = clinic_ncid" & " and u.user_ncid = " & oUser.NCID

On Error Resume Next   '__ for duuplicate clinicncids
    oSQL.Execute "select distinct cp.facility_ncid, cp.clinic_ncid, c.name as clinic_name, f.description as facility_name " _
    & " from clinic c, clinic_provider cp, ncids f " _
    & " where f.ncid = cp.facility_ncid and cp.clinician_ncid = " & oUser.NCID _
    & " And cp.facility_ncid = '" & gsThisFacility_NCID & "' and cp.clinic_ncid = c.ncid"

    Do While Not oSQL.EOF
        Set oFac = New User_Fac
        oFac.clinicName = oSQL("clinic_name")
        oFac.FacilityName = oSQL("facility_name")
        oFac.ClinicNCID = oSQL("clinic_ncid")
        oFac.FacilityNCID = oSQL("facility_ncid")
        oUser.Facilities.Add oFac, CStr(oFac.ClinicNCID)
        
        oSQL.MoveNext
    Loop

End Function

Public Function DuplicateUID(ByVal suid As String, ByVal resId As Long, ByVal sUserNCID As String) As Boolean
'- be sure there is not duplicate uid
' return true if one is found
    '--- SCR 83298   Sherry Wang  12/5/2005
    oSQL.Execute "select count(*) from user_resources where res_id = " & resId & " and upper(user_name) = upper(" & SQLQuote(suid) & ") and user_ncid <> " & sUserNCID
    If oSQL(0) > 0 Then
        DuplicateUID = True
        MsgBox "The logon name is already in use. Please select another.", vbCritical, "User Logon Error"
    End If
End Function

Private Sub Class_Initialize()
    Set mobjCrypto = New CCrypto
End Sub

'/// special call for maint use only
Public Sub RefreshAllUsers()
    Dim U As User
    For Each U In mcolAllUsers
        SaveUser U
    Next U
End Sub

Private Sub Class_Terminate()
'11/13/01 - Walker - SCR 18215 - New - Start
    Set mobjCrypto = Nothing
'11/13/01 - Walker - New - End
End Sub


Public Function ExportToSpreadsheet(ByVal sFileName As String) As Boolean
    On Error GoTo ErrHandler

    Dim orole           As role
    Dim oPriv           As privilege
    Dim i               As Integer
    Dim j               As Integer
    Dim aData(1 To 300, 1 To 200) As String
    Dim bFoundIt        As Boolean
    Dim lLastRow        As Long
    Dim lCurCol         As Long
    Dim lCurRow         As Long
    Dim objFS As FileSystemObject
    Dim objTS As TextStream
    
    lCurCol = 3
    lLastRow = 3
    Screen.MousePointer = vbHourglass
    Debug.Print "Start Exporting: " & Time
    For Each orole In mcolAllRoles
    Debug.Print orole.RoleName
        lCurCol = lCurCol + 1  '-- the place where I'm going to write the priv string
        '__ enter the role info
        aData(1, lCurCol) = orole.roleID
        aData(2, lCurCol) = orole.RoleName
        aData(3, lCurCol) = orole.RoleDesc
        '__ look for the priv_id in the first col of each active row
        For Each oPriv In orole.Privileges
            bFoundIt = False
            For i = 4 To lLastRow
                If aData(i, 1) = oPriv.PrivID Then
                    bFoundIt = True
                    Exit For        '__ with i on the proper row...
                End If
            Next i
            If Not bFoundIt Then
                '__ we have to add it
                lLastRow = lLastRow + 1
                aData(lLastRow, 1) = oPriv.PrivID
                aData(lLastRow, 2) = oPriv.PrivName
                aData(lLastRow, 3) = oPriv.PrivDesc
                i = lLastRow
            End If
            aData(i, lCurCol) = oPriv.PrivString
        
        Next oPriv
    Next orole
    
    
    Dim sLine As String
    
    '--- Open File
    Set objFS = New FileSystemObject
    Set objTS = objFS.CreateTextFile(sFileName, True)
    
    '--- Write to File
    For i = 1 To 300
        For j = 1 To mcolAllRoles.count + 3  '--- SCR 72007  Sherry Wang 7/5/2005
            sLine = sLine & aData(i, j) & vbTab
        Next j
        
        If Len(Replace(sLine, vbTab, "")) = 0 Then
            Exit For
        Else
            objTS.WriteLine (sLine)
            sLine = vbNullString
        End If
    Next i

    Debug.Print "End Exporting: " & Time
    Screen.MousePointer = vbDefault
    
    '--- SCR 65473   Sherry Wang   2/14/2005
    objTS.Close
    MsgBox "The data is successfully exported to " & sFileName & "."
       
    Exit Function
ErrHandler:
    Screen.MousePointer = vbDefault
    MsgBox "Error: " & Err.Description
End Function

Public Function RemoveUser(oUser As User) As Boolean
    On Error GoTo ErrHandler
    osql_ns.Execute "delete from mmi_name where unit_number = " & oUser.Unit_Number
    osql_ns.Execute "delete from mmi_site where unit_number = " & oUser.Unit_Number
    osql_ns.Execute "delete from mmi_id where unit_number = " & oUser.Unit_Number
    
    osql_ns.Execute "delete from clinic_provider where clinician_ncid = " & oUser.NCID
    osql_ns.Execute "delete from user_clinic where user_ncid = " & oUser.NCID
    osql_ns.Execute "delete from provider where ncid = " & oUser.NCID
    
    oSQL.Execute "delete from user_resources where user_ncid = " & oUser.NCID
    oSQL.Execute "delete from users where user_ncid = " & oUser.NCID
    
    osql_ns.Execute "delete from hems_users where user_ncid = " & oUser.NCID
    RemoveUser = True
    Exit Function
    
ErrHandler:
    MsgBox "Error occured in SecOps.RemoveUser: " & Err.Description
End Function


Private Sub ResetfrmSSRead()
    On Error GoTo ErrHandler
    
    frmSSRead.Show
    frmSSRead.pbRoles.Min = 0
    frmSSRead.pbRoles.Value = 0
    frmSSRead.pbPrivs.Min = 0
    frmSSRead.pbPrivs.Value = 0
    frmSSRead.pbRolePrivs.Min = 0
    frmSSRead.pbRolePrivs.Value = 0
    frmSSRead.Refresh
    
    Exit Sub
ErrHandler:
    MsgBox "Error: " & Err.Description
End Sub


Private Sub GetRolePrivData(ByVal sFile As String)
    On Error GoTo ErrHandler
    
    Dim sLine As String
    Dim iRow As Integer
    Dim iCol As Integer
    Dim aRowData As Variant
    
    Open sFile For Input As #1   ' Open file.
    iRow = 0
    Do While Not EOF(1)
        Line Input #1, sLine
        iRow = iRow + 1
        aRowData = Split(sLine, vbTab)
        If UBound(aRowData) = 0 Then
            MsgBox "The data format is incorrect!"
            mbWrongFileType = True
            Exit Sub
        End If
        
        For iCol = 0 To UBound(aRowData)
            maPrivData(iRow, iCol + 1) = aRowData(iCol)
        Next iCol
    Loop
    Close #1   ' Close file
    
    Exit Sub
ErrHandler:
    MsgBox "Error: " & Err.Description
End Sub

Private Function GetTotalPrivileges() As Integer
    On Error GoTo ErrHandler
    
    Dim iRow As Integer
    
    For iRow = 4 To 300
        If Len(maPrivData(iRow, 2)) = 0 Then
            GetTotalPrivileges = iRow - 4
            Exit Function
        End If
    Next iRow
    
    Exit Function
ErrHandler:
    MsgBox "Error: " & Err.Description
End Function


Private Function GetTotalRoles() As Integer
    On Error GoTo ErrHandler
    
    Dim iCol As Integer
    
    For iCol = 4 To 200
        If Len(maPrivData(2, iCol)) = 0 Then
            GetTotalRoles = iCol - 4
            Exit Function
        End If
    Next iCol
    
    Exit Function
ErrHandler:
    MsgBox "Error: " & Err.Description
End Function



