VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Security"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements CHCSII_ClientSecurity.ICHCSII_Security
Implements CHCSII_ClientSecurity.ITMIP_AppData

Private Const ACTIVATE_KEY As String = "ASKJHDSFIOUYCXOIVUZXCKBASDFBASOIXZYCVKASDBFMNEQ234KYF8S7D6F98AD7F87ASD9867FQWOEGHFASKDJHGVZXCYTV;"
Private Const CLASS_NAME As String = "C2T_ClientSecurity.Security"
Private m_sDomain        As String
Private m_bActivated     As Boolean
Private m_colUsers       As New Collection        'Collection of User objects
Private m_Roles          As New Roles
Private m_INIPath        As String
Private m_objDefaultUser As ICHCSII_User
Enum SecModeEnum
    c2t = 2
    STUB = 4
    PGUI = 8
    tmip = 16
    CDR = 32
    ITT = 64
End Enum
Private SecMode As SecModeEnum
Private bTMIPAvail As Boolean       '__ is tmip framework on this system?  Set in class_initialize
Private WithEvents goPwdOps As clsPasswordOps
Attribute goPwdOps.VB_VarHelpID = -1


Public Function IsSecMode(ByVal SecModeVal As SecModeEnum) As Boolean
    IsSecMode = CBool(SecMode And SecModeVal)
End Function

Public Property Get TMIPAvailable() As Boolean
    TMIPAvailable = bTMIPAvail
End Property


Public Function UpdateSupportPwdInfo(GemsPwd As String, SecadminPwd As String) As Boolean
    
'SF SCR 93559
    Dim sEPWD As String
    Dim sEFramework_PWD As String
    
    On Error GoTo ErrHandler
    
    If GemsPwd <> vbNullString And SecadminPwd <> vbNullString Then
        sEPWD = EncryptString(GemsPwd)
        sEFramework_PWD = EncryptString(SecadminPwd)
    
        SaveRegistry "", sEPWD, sEFramework_PWD
            
        UpdateSupportPwdInfo = True
    End If
    
    Exit Function
ErrHandler:
    
    UpdateSupportPwdInfo = False
    
End Function

Private Sub Class_Initialize()
    m_bActivated = False
    Dim oShared As CWShared
    Set oShared = New CWShared
    
    If oShared.CmdLineSwitches("STUB") <> vbNullString Then
        SecMode = SecMode + STUB
       '* modified by j.ray 05SEP2001 SCR#16514; ini path was not being set before attempting to read stub file
        m_INIPath = Trim(oShared.INIPath)
        If Right$(m_INIPath, 1) <> "\" Then m_INIPath = m_INIPath & "\"
        ReadSecurityStubInfo
    End If
    
    If oShared.IsAppMode(modeTheater) Then
        SecMode = SecMode + c2t
        If oShared.CmdLineSwitches("CHCSII-T") <> vbNullString Then
            bTMIPAvail = True
            SecMode = SecMode + tmip
        End If
    End If
    If oShared.IsAppMode(modeCDR) Then
        SecMode = SecMode + CDR
    End If
    
    If oShared.IsAppMode(modeCHCSI_GUI) Then
        SecMode = SecMode + PGUI
    End If

    PrivType(1) = "R"    'Read
    PrivType(2) = "W"    'Write
    PrivType(3) = "D"    'Delete
    PrivType(4) = "S"    'Sign
    PrivType(5) = "N"    'Need Cosign
    PrivType(6) = "B"    'Break the Glass
    PrivType(7) = "C"    'Can Cosign
    
    PrivValue(1) = Priv_Read
    PrivValue(2) = Priv_Write
    PrivValue(3) = Priv_Delete
    PrivValue(4) = Priv_Sign
    PrivValue(5) = Priv_Need_Cosign
    PrivValue(6) = Priv_Break_The_Glass
    PrivValue(7) = Priv_Cosign
     
End Sub

Private Sub Class_Terminate()
    Set m_colUsers = Nothing
End Sub


Private Sub goPwdOps_LoginSuccessful(ByVal UserName As String, ByVal Password As String, ByVal SecSQL As SqlOps, User As CHCSII_ClientSecurity.ICHCSII_User)

    Set goSecSQL = SecSQL
    Set User = GetUserEx(UserName, Password)
    
End Sub

Private Function ICHCSII_Security_Activate(ActivateKey As String) As Boolean
On Error GoTo ErrHandler
    If (ActivateKey = ACTIVATE_KEY) Then
        m_bActivated = True
        ICHCSII_Security_Activate = True
    Else
        m_bActivated = False
        ICHCSII_Security_Activate = False
    End If
    Exit Function
ErrHandler:
    MsgBox "Unexpected error in Security Module.Activate: " & vbCrLf & Err.Description
End Function

Private Function ICHCSII_Security_ChangePassword(User As ICHCSII_User) As Boolean

On Error GoTo ErrHandler
    If Not (m_bActivated) Then
        Err.Raise SEC_NOT_ACTIVATED, CLASS_NAME, GetErrMsg(SEC_NOT_ACTIVATED)
        Exit Function
    End If
    If IsSecMode(STUB) Then
        MsgBox "Not currently implemented.", vbOKOnly, "Security - Change Password"
        ICHCSII_Security_ChangePassword = True
    
    ElseIf IsSecMode(tmip) Then
         ICHCSII_Security_ChangePassword = False  '<< scr 36860  goTMIP_Security.ChangePassword(User)
    
    ElseIf IsSecMode(c2t) Then
         ICHCSII_Security_ChangePassword = GetPwdOps.ChangeC2TPassword(User.UserId)
         
    End If
   
    Exit Function
ErrHandler:
    MsgBox "Unexpected error in Security Module.ChangePassword: " & vbCrLf & Err.Description

End Function

Private Property Get ICHCSII_Security_Domain() As String
On Error GoTo ErrHandler
    If Not (m_bActivated) Then
        Err.Raise SEC_NOT_ACTIVATED, CLASS_NAME, GetErrMsg(SEC_NOT_ACTIVATED)
        Exit Function
    End If
    
    If IsSecMode(tmip) Then
        ICHCSII_Security_Domain = THEATER_KEYWORD '<< scr 36860
        
    ElseIf IsSecMode(c2t) Or IsSecMode(STUB) Then
        ICHCSII_Security_Domain = m_sDomain
    End If
    
    Exit Function
ErrHandler:
    MsgBox "Unexpected error in Security Module.Domain: " & vbCrLf & Err.Description
    
End Property

Private Function ICHCSII_Security_GetActiveUser(Optional sFormTitle As String, Optional sWindowTitle As String) As ICHCSII_User
On Error GoTo ErrHandler
    If Not (m_bActivated) Then
        Err.Raise SEC_NOT_ACTIVATED, CLASS_NAME, GetErrMsg(SEC_NOT_ACTIVATED)
    End If
    
    
    If IsSecMode(tmip) Then
        Set ICHCSII_Security_GetActiveUser = BuildTMIPUser
        Exit Function
    End If
    
    Dim sUser As String
    Set ICHCSII_Security_GetActiveUser = GetUserFromLogin(sFormTitle, sWindowTitle)
    
    Exit Function
ErrHandler:
'<< scr 33048
    If Err.Number = -2147220397 Then ' User cancelled login.
        'MsgBox Err.Description, vbInformation
        'Don't display a "User cancelled login" message.
    ElseIf Not Err.Number = -2147220497 Then
        MsgBox "Unexpected error in Security Module.GetActiveUser: " & vbCrLf & Err.Description
    End If
End Function

Private Function ICHCSII_Security_GetPrivilegesByRole(l_User As ICHCSII_User, l_Role As ICHCSII_Role) As Collection
'<< scr 36860
On Error GoTo ErrHandler
    If (m_bActivated) Then
        Dim colPrivileges As New Collection
        Dim userRole As role
        Dim sysRole As role
        Dim i As Integer
        
        On Error Resume Next
        'Check to see if user has role
        Set userRole = l_User.Roles.Item(l_Role.Name)
        If Err Then
            On Error GoTo 0
        End If
        
        'Get system role so we can get privileges
        Set sysRole = m_Roles.Item(l_Role.Name)
        If Err Then
            'Shouldn't happen, but just in case it does...
            On Error GoTo 0
            Err.Raise SEC_NO_SYSTEM_ROLE, CLASS_NAME, GetErrMsg(SEC_NO_SYSTEM_ROLE)
        End If
        
        On Error GoTo 0
        For i = 1 To sysRole.Privileges.Count
            colPrivileges.Add sysRole.Privileges.Item(i)
        Next
        
        i = 1
        Do
            If i > colPrivileges.Count Then Exit Do
            If Not (l_User.HasPrivilege(colPrivileges.Item(i))) Then
                colPrivileges.Remove (i)
            Else
                i = i + 1
            End If
        Loop
        
        Set ICHCSII_Security_GetPrivilegesByRole = colPrivileges
    Else
        Err.Raise SEC_NOT_ACTIVATED, CLASS_NAME, GetErrMsg(SEC_NOT_ACTIVATED)
    End If
    Exit Function
ErrHandler:
    MsgBox "Unexpected error in Security Module.GetPrivilegesByRole: " & vbCrLf & Err.Description
    
End Function

Private Function ICHCSII_Security_Logoff(User As ICHCSII_User) As Boolean
On Error GoTo ErrHandler

    If (m_bActivated) Then
        ICHCSII_Security_Logoff = True
        
        'SF Log it. SCR 83806 2/8/06
        goPwdOps.Logout

    Else
        Err.Raise SEC_NOT_ACTIVATED, CLASS_NAME, GetErrMsg(SEC_NOT_ACTIVATED)
    End If
    Exit Function
ErrHandler:
    MsgBox "Unexpected error in Security Module.LogOff: " & vbCrLf & Err.Description
    
End Function

Private Function ICHCSII_Security_Logon(UserId As String, Password As String) As ICHCSII_User
On Error GoTo ErrHandler
    If bTMIPAvail Then
        Exit Function
    End If
    
    Dim l_User As ICHCSII_User
    
    If (m_bActivated) Then
        On Error Resume Next

        
        If C2TLogon(UserId, Password) Then
             Set ICHCSII_Security_Logon = GetC2TUser(UserId, Password, vbNullString, vbNullString) 'sName As String
        Else
             Err.Raise SEC_USER_NOT_FOUND, CLASS_NAME, GetErrMsg(SEC_USER_NOT_FOUND)
        End If
    Else
        Err.Raise SEC_NOT_ACTIVATED, CLASS_NAME, GetErrMsg(SEC_NOT_ACTIVATED)
    End If
    Exit Function
ErrHandler:
    MsgBox "Unexpected error in Security Module.Logon: " & vbCrLf & Err.Description
    
End Function

Private Function GetUserFromLogin(sFormTitle As String, sWindowTitle As String) As ICHCSII_User

On Error GoTo ErrHandler
    
    Set GetUserFromLogin = GetPwdOps.PromptUserToLogin
    
    Exit Function
ErrHandler:
    If Not Err.Description = "User cancelled login." Then
        MsgBox "Unexpected error in Security Module.GetUserFromLogin: " & vbCrLf & Err.Description
    Else
        Err.Raise SEC_USER_CANCELLED_LOGIN, CLASS_NAME, GetErrMsg(SEC_USER_CANCELLED_LOGIN)
    End If
End Function

Private Function GetC2TUser(ByVal sUserId As String, _
                             ByVal sPassword As String, _
                             ByVal sName As String, _
                             ByVal sUserNCID As String, _
                    Optional ByVal PrivilegesOnly As Boolean = False) As ICHCSII_User

    Dim l_Resources     As New Resources
    Dim objC2TUser      As User
    Dim l_Roles         As New Roles
    Dim sRoles, sRole   As String
    Dim sPhones, sPhone As String
    Dim sAttr           As String
    Dim sKey, sVal      As String
    Dim sAccess         As String
    Dim sVerify         As String
    Dim sUser           As String
    Dim sPass           As String
    Dim l_Attrib        As New Collection  '- of KeyValuePair
    Dim l_ResData       As New Collection  '- of KeyValuePair
    Dim l_Privileges    As New Collection  '- of KeyValuePair
    Dim inKey, inQuote  As Boolean
    Dim objPair         As KeyValuePair
    Dim i               As Integer
'    Dim C               As Integer
    'Dim sUserNCID       As String
    Dim bLoggingIn      As Boolean
    
    On Error Resume Next
'-- if spassword = vbnullstring and sname = vbnullstring then we are just getting a user object to check some privileges, not to logon
    bLoggingIn = CBool(Len(sPassword & sName))
    
    
    If sUserNCID = vbNullString Or sPassword = vbNullString Then
        If Not GetUserInfo(sUserId, sUserNCID, sPassword) Then
            Exit Function
        End If
    End If
    
    
    Set objC2TUser = m_colUsers.Item(sUserNCID)
    
    If Err = 0 Then
        Set GetC2TUser = objC2TUser
        Exit Function
    End If

    Err.Clear
    
    On Error GoTo ErrHandler


    '__ if this is privileges only, don't bother with the rest of the stuff
    If Not PrivilegesOnly Then
        goSecSQL.Execute "select ur.user_password password, ur.user_name name, u.* " _
                  & " FROM users u,user_resources ur, resources r" _
                  & " where ur.user_name = " & SQLQuote(sUserId) _
                  & " AND u.user_ncid = ur.user_ncid " _
                  & " and r.res_id = ur.res_id" _
                  & " AND r.resource_name = '" & THEATER_KEYWORD & "'"
        
        If Not goSecSQL.EOF Then
        '- confirm user has rights to this resource
            If bLoggingIn Then
                If sUserId <> Trim(goSecSQL("name")) Then
                    Err.Raise SEC_USER_NOT_FOUND, CLASS_NAME, GetErrMsg(SEC_USER_NOT_FOUND)
                End If
            End If
            m_sDomain = vbNullString & goSecSQL!domain
            sUser = vbNullString & goSecSQL!Name
            sPass = Trim(vbNullString & goSecSQL!Password)
            
            For i = 0 To goSecSQL.ColumnCount - 1
                Set objPair = New KeyValuePair
                objPair.Key = goSecSQL.ColumnName(i)
    '<< SCR ??? use the value not the name
                objPair.Value = Trim(vbNullString & goSecSQL.Value(i))
                l_Attrib.Add objPair, objPair.Key
            Next i
            
            Set objPair = Nothing
            
        End If
       
        Set l_ResData = New Collection
        
        goSecSQL.Execute "select distinct ra.key, ra.value " _
          & " from user_resources ur, resource_attributes ra" _
          & " where ur.user_ncid = " & SQLQuote(sUserNCID) _
          & " and ur.res_id = ra.res_id"
        Do While Not goSecSQL.EOF
            Set objPair = New KeyValuePair
            objPair.Key = vbNullString & goSecSQL!Key
            objPair.Value = vbNullString & goSecSQL!Value
            l_ResData.Add objPair, objPair.Key
            l_Attrib.Add objPair, objPair.Key
            goSecSQL.MoveNext
        Loop
        
        l_Resources.AddResource THEATER_KEYWORD, sUser, sPass, l_ResData
    
        ' User Role
         goSecSQL.Execute "select role_name from user_roles ur, roles r where ur.role_id = r.role_id and user_ncid = " & SQLQuote(sUserNCID)
        Do While Not goSecSQL.EOF
            l_Roles.AddRole vbNullString & goSecSQL!role_name
            goSecSQL.MoveNext
        Loop
    
    End If  '-- not privileges onlly
    
'-- get privileges
     goSecSQL.Execute "select distinct p.priv_name, rp.priv_string " & vbCrLf _
        & " from user_roles ur, role_priv rp, available_priv p " & vbCrLf _
        & " where ur.role_id = rp.role_id and p.priv_id = rp.priv_id" & vbCrLf _
        & " and ur.user_ncid = " & sUserNCID & vbCrLf _
        & " UNION ALL " & vbCrLf _
        & " select p.priv_name, priv_string from user_other_privileges uop, available_priv p " & vbCrLf _
        & "   where uop.priv_id = p.priv_id and user_ncid = " & sUserNCID & " and add_remove = 'A'"
    
    Do While Not goSecSQL.EOF
        On Error Resume Next
        If vbNullString & goSecSQL!Priv_String <> vbNullString Then
            Set objPair = l_Privileges(vbNullString & goSecSQL!Priv_Name)
            If Err.Number = 0 Then
                objPair.Value = AddPriv(objPair.Value, vbNullString & goSecSQL!Priv_String)
            Else
                Err.Clear
                Set objPair = New KeyValuePair
                objPair.Key = vbNullString & goSecSQL!Priv_Name
                objPair.Value = vbNullString & goSecSQL!Priv_String
                l_Privileges.Add objPair, objPair.Key
            End If
        End If
        On Error GoTo ErrHandler
        goSecSQL.MoveNext
    Loop


'-- now delete the revoked privileges
    Dim SPName      As String
    Dim sPString    As String
    
    goSecSQL.Execute "select p.priv_name, priv_string from user_other_privileges uop, available_priv p " & vbCrLf _
        & "   where uop.priv_id = p.priv_id and user_ncid = " & sUserNCID & " and add_remove = 'R'"
    Do While Not goSecSQL.EOF
        SPName = goSecSQL("priv_name")
        sPString = goSecSQL("priv_string")
        For Each objPair In l_Privileges
            With objPair
            If UCase$(.Key) = UCase$(SPName) Then
                '- take each priv away
                For i = 1 To Len(sPString)
                    .Value = Replace(.Value, Mid$(sPString, i, 1), vbNullString)
                Next i
                Exit For
            End If
            End With
        Next objPair
        goSecSQL.MoveNext
    Loop
    
    Set objC2TUser = New User

    With objC2TUser
    
'<< scr 44784 user ncid was in the userid property
        .UserId = sUser 'Logon Name
        'SCR 102635 - LCI - 08/23/2006 * UserName variable was being set to sUser, which is the same as UserID (logon name).
        'If other parts of the applications require UserName to be the logon name simply change those areas to use UserID
        'instead of having two variables equal the same thing.
        .UserName = GetFullUserName(sUserNCID) 'User's Name
        .Password = sPass
        Set .Roles = l_Roles
        Set .Resources = l_Resources
        Set .Attributes = l_Attrib
        Set .Privileges = l_Privileges
    End With

    m_colUsers.Add objC2TUser, sUserNCID

    Set GetC2TUser = objC2TUser
    If bLoggingIn Then
        objC2TUser.DumpObject
    End If
    GoTo Cleanup
ErrHandler:
    MsgBox "Unexpected error in Security Module.GetC2TUser: " & vbCrLf & Err.Description
    
Cleanup:
    Set objPair = Nothing
   
End Function

Private Function ICHCSII_Security_GetUser(Optional sUserId As String) As ICHCSII_User
On Error GoTo ErrHandler

    Set ICHCSII_Security_GetUser = GetUserEx(sUserId)
    Exit Function

ErrHandler:
    MsgBox "Unexpected error in Security Module.GetUser: " & vbCrLf & Err.Description
    
End Function

Private Function GetUserEx(Optional sUserId As String, Optional sPassword As String, Optional sName As String) As ICHCSII_User
'__ may be called from 'get Active User' > 'get user from login'
On Error GoTo ErrHandler
    Dim l_User As ICHCSII_User
    Dim NewUser As User
    Dim sUserNCID As String
    Dim sUserSearch As String
    Dim objPair As KeyValuePair
    Dim l_Attrib As Collection
    
    If Not (m_bActivated) Then
        Err.Raise SEC_NOT_ACTIVATED, CLASS_NAME, GetErrMsg(SEC_NOT_ACTIVATED)
        Exit Function
    End If
    '-- if the userid is  null, then just get the active user function to get the user.
        If sUserId = vbNullString Then
            Set GetUserEx = ICHCSII_Security_GetActiveUser()
            Exit Function
        End If

''-- this may be an ncid or may be a user logon.
''   we want the ncid
        If Not IsNumeric(sUserId) Then   '- they gave us the user logon name so get the ncid
            Call GetUserInfo(sUserId, sUserNCID, sPassword)
        Else ' - they gave us the ncid in the user id field.  correct this and call for the user logon name
            sUserNCID = sUserId
            sUserId = vbNullString
            Call GetUserInfo(sUserId, sUserNCID, sPassword)
        End If
''-- now we have the userid holding the logon name, and the userncid holding the ncid.
        '--- SCR 79741   Sherry Wang  10/12/2005
        If gbDBLoginInfoChanged = True Then Exit Function
        On Error Resume Next

        Set l_User = m_colUsers.Item(sUserNCID)
        If Err Then
            Err.Clear
            If bTMIPAvail Then
                '__ generate a user object shell for use within tmip only.  where we don't have access to the other users.
                Set objPair = New KeyValuePair
                objPair.Key = "USER_NCID"
                objPair.Value = sUserNCID
                l_Attrib.Add objPair, objPair.Key

                Set NewUser = New User
                With NewUser
                    .UserId = sUserId
                    .UserName = sUserId  ' << SCr 47714 - no get priv response from ngit.  This was not being set so it was not being sent
                    .Password = "PRIVILEGE_CHECK_ONLY"
                    Set .Attributes = l_Attrib
                End With
                Set GetUserEx = NewUser
                Exit Function
             End If
            On Error GoTo 0
                        
            Set l_User = GetC2TUser(sUserId, sPassword, sName, sUserNCID)
                    
        Else
'            Debug.Print "found old user in security obj"
        End If
        
        Set GetUserEx = l_User
    Exit Function
ErrHandler:
    MsgBox "Unexpected error in Security Module.GetUserEx: " & vbCrLf & Err.Description
    
End Function


Public Function SQLQuote(sTxt As String) As String
' doubles up the single quotes and wraps the string in single quotes
    
    'Use the existing SQLQuote function
    SQLQuote = SecGlobal.SQLQuote(sTxt)
    
    'SQLQuote = "'" & Replace(Replace(sTxt, Chr(34), Chr(34) & Chr(34)), "'", "''") & "'"
End Function


Public Function Stuff(sTxt As String, lPosn As Long, sNewTxt As String) As String
'- stick a string (sNewTxt) inside of a string (sTxt) at a position (lposn)
    
    'Use the existing Stuff function
    Stuff = SecGlobal.Stuff(sTxt, lPosn, sNewTxt)
    
    'Stuff = Left$(sTxt, lPosn) & sNewTxt & Mid$(sTxt, lPosn + 1)
End Function

Private Function C2TLogon(suid As String, spwd As String) As Boolean
On Error GoTo ErrHandler
    Dim oCrypto As CCrypto
    Set oCrypto = New CCrypto

    goSecSQL.Execute "select ur.user_password password, ur.user_name name " _
              & " FROM users u,user_resources ur, resources r" _
              & " where ur.user_name = " & SQLQuote(suid) _
              & " AND u.user_ncid = ur.user_ncid " _
              & " and r.res_id = ur.res_id" _
              & " AND r.resource_name = '" & THEATER_KEYWORD & "'"
    
    If Not goSecSQL.EOF Then
        If oCrypto.GetHash(spwd, True) = Left(goSecSQL("password"), 32) Then
            C2TLogon = True
        End If
    End If
    Exit Function
ErrHandler:
    MsgBox "Unexpected error in Security Module.C2TLogon: " & vbCrLf & Err.Description

    
End Function

Private Function AddPriv(ByVal sPriv1 As String, ByVal sPriv2 As String) As String
'__ sPriv1 is the current string representing the user before this method is called
'   sPriv2 is the string from a new role to be added, without repetition, to this user
On Error GoTo ErrHandler
'    R - READ, W - Write, D - Delete, C - Can Cosign, N - Needs Cosign, B - Break Glass, S - Can Sign
       
    Dim sout As String
    
    '__ handle the restrictive privileges
    If InStr(sPriv1, "N") > 0 Or InStr(sPriv1, "B") > 0 _
    Or InStr(sPriv2, "N") > 0 Or InStr(sPriv2, "B") > 0 Then
        Call CheckRestrictions(sPriv1, sPriv2)
    End If
    
    sout = vbNullString
    Dim i As Integer
    '__ for each of the possible letters in the priv string (WRDCNSB) add it to the output string if:
    '    1 - it exists already (priv1), or
    '    2 - it is in the incoming string (priv2)
    For i = 1 To UBound(PrivType)
        If (InStr(sPriv1, PrivType(i)) > 0) Or (InStr(sPriv2, PrivType(i)) > 0) Then
            sout = sout & PrivType(i)
        End If
    Next i
    AddPriv = sout
    Exit Function
ErrHandler:
    MsgBox "Unexpected error in Security Module.AddPriv: " & vbCrLf & Err.Description
    
End Function

Private Function CheckRestrictions(ByRef sPriv1 As String, ByRef sPriv2 As String) As String
    '--- handle the 'negative' privileges (restrictions) of 'N' and 'B'
    '    'N' and 'S' are exclusive, as are 'B' and 'R', when it comes to them appearing in separate roles.
    '    with those, if the incoming priv is a restriction, and they already have the privilege from another role,
    '           then ignore the restriction.
    '           and if the restriction exists from another role, and the pivilege is coming in, the remove the restriction.
    Dim S1 As Long
    Dim N1 As Long
    Dim R1 As Long
    Dim B1 As Long
    
    Dim S2 As Long
    Dim N2 As Long
    Dim R2 As Long
    Dim B2 As Long
    
    
    
    
    
    S1 = InStr(sPriv1, "S")
    N1 = InStr(sPriv1, "N")
    N2 = InStr(sPriv2, "N")
    '-- if we have and s without an N, and we have an incoming N, drop the incoming N
    If (S1 > 0 And N1 = 0) And N2 > 0 Then
        '_remove the N from priv2
        sPriv2 = Replace(sPriv2, "N", "")
    End If
    
    N1 = InStr(sPriv1, "N")
    S2 = InStr(sPriv2, "S")
    N2 = InStr(sPriv2, "N")
'--if we have an N and the priv2 has an S without an N then drop the current N
    If N1 > 0 And (S2 > 0 And N2 = 0) Then
        '_remove the n from priv1
        sPriv1 = Replace(sPriv1, "N", "")
    End If
    
    R1 = InStr(sPriv1, "R")
    B1 = InStr(sPriv1, "B")
    R2 = InStr(sPriv2, "R")
    B2 = InStr(sPriv2, "B")
    
    '-- if we have and s without an N, and we have an incoming N, drop the incoming N
    If R1 > 0 And B2 > 0 Then
        '_remove the b from priv2
        sPriv2 = Replace(sPriv2, "B", "")
    End If
    R2 = InStr(sPriv2, "R")
    B2 = InStr(sPriv2, "B")
    
    If B1 > 0 And (R2 > 0 And B2 = 0) Then
        '_remove the b from priv1
        sPriv1 = Replace(sPriv1, "B", "")
    End If
    
    
End Function


Private Function UpdateNSN(ByVal SID As String) As Boolean
    
End Function


Private Function BuildTMIPUser() As ICHCSII_User
'-- get the decode info from the cmd line and get
On Error GoTo ErrHandler
    Const AppID         As String = "CHCSII-T"
    Dim i               As Integer
    Dim CmdLine()       As String
    Dim RegDecryptKey   As String
    Dim SecInfo         As String
    Dim SecLines()      As String
    Dim KeyString       As String
    Dim ValString       As String
    Dim user_ncid       As String
    Dim HomePhone       As String
    Dim WorkPhone       As String
    Dim sPass        As String
    Dim sUser        As String
    Dim oUser           As User
    
    Dim l_Resources     As New Resources
    Dim l_Roles         As New Roles
    Dim l_Attrib        As New Collection  '- of KeyValuePair
    Dim l_ResData       As New Collection  '- of KeyValuePair
    Dim l_Privileges    As New Collection  '- of KeyValuePair
    Dim inKey, inQuote  As Boolean
    Dim objPair         As CHCSII_Utils.KeyValuePair
    
    Dim oShared As CWShared
    Set oShared = New CWShared
    
  
    RegDecryptKey = Trim(oShared.CmdLineSwitches("CHCSII-T"))
   
    If RegDecryptKey = vbNullString Then Exit Function   '-- we can't go on
'    MsgBox "Calling: [" & AppID & "] [" & RegDecryptKey & "] "
    
    '-- get the key=value key-pairs from the registry
    SecInfo = get_security_info(AppID, RegDecryptKey)
    
    'MsgBox "Returned: " & SecInfo
    If SecInfo = vbNullString Then
        MsgBox "Unexpected internal error: Unable to retreive logon information." & vbCrLf _
            & "Please contact your system administrator.", vbCritical, "Security Error"
        Exit Function
    End If
    
    '-- is in the format of:
'----- Test use only
'    SecInfo = "USER=sysadmin" & vbCrLf _
'    & "PASSWORD=5F4DCC3B5AA765D61D8327DEB882CF99" & vbCrLf _
'    & "ROLE=er_provider" & vbCrLf _
'    & "ROLE=provider_attending" & vbCrLf _
'    & "ROLE=ward_nurse" & vbCrLf _
'    & "Allergies=RWDRWD" & vbCrLf _
'    & "Appointments=WD" & vbCrLf _
'    & "Appointments=R" & vbCrLf _
'    & "Autocite=RWDRWD" & vbCrLf _
'    & "Current_Encounter=RWDSC" & vbCrLf _
'    & "Diagnosis=RWD" & vbCrLf _
'    & "Encounter_Coding=RWDRWD" & vbCrLf _
'    & "Health_History=RWDRWD" & vbCrLf _
'    & "Laboratory_List=RWD"

    SecLines = Split(SecInfo, vbCrLf)
    '__ set some defaults
    m_sDomain = THEATER_KEYWORD
    
'__ add users table elements:
    Set objPair = New KeyValuePair
    objPair.Key = "TIMER"
    objPair.Value = "120"
    l_Attrib.Add objPair, objPair.Key
    
    Set objPair = New KeyValuePair
    objPair.Key = "Domain"
    objPair.Value = THEATER_KEYWORD
    l_Attrib.Add objPair, objPair.Key
    
    
    For i = 0 To UBound(SecLines)
    
        If InStr(SecLines(i), "=") > 0 Then   '-- account for inital or final blank line
            KeyString = UCase$(Trim(Mid$(SecLines(i), 1, InStr(SecLines(i), "=") - 1)))
            ValString = Trim(Mid$(SecLines(i), InStr(SecLines(i), "=") + 1))
            Select Case KeyString
            Case "USER"  '-- user logon name, look it up the hems_user table.
'                If Not GetUserNCID(ValString, User_NCID) Then
'                    '-- couldn't find this guy, return nothing
'                    Exit Function
'                End If
                    Set objPair = New KeyValuePair
                    objPair.Key = "name"
                    objPair.Value = ValString
                    l_Attrib.Add objPair, objPair.Key
               sUser = ValString
                
            Case "PASSWORD"
                    Set objPair = New KeyValuePair
                    objPair.Key = "password"
                    objPair.Value = ValString
                    l_Attrib.Add objPair, objPair.Key
                sPass = ValString
            Case "ROLE"
                l_Roles.AddRole ValString
            Case Else   '- A privilage
                On Error Resume Next
                Set objPair = l_Privileges(KeyString)
                If Err.Number = 0 Then
'<< SCR 38287 changed the second param in add priv
                    objPair.Value = AddPriv(ValString, objPair.Value)
                Else
                    Err.Clear
                    Set objPair = New KeyValuePair
                    objPair.Key = KeyString
                    objPair.Value = ValString
                    l_Privileges.Add objPair, objPair.Key
                End If
                On Error GoTo ErrHandler
            
            End Select
            
        End If
    Next i
'<< Confirm user is in C2T system
    If Not GetUserInfo(sUser, user_ncid, sPass, HomePhone, WorkPhone) Then
        Exit Function
    End If
    Set l_ResData = New Collection
    l_Resources.AddResource THEATER_KEYWORD, sUser, sPass, l_ResData

    Set objPair = New KeyValuePair
    objPair.Key = "HOMEPHONE"
    objPair.Value = HomePhone
    l_Attrib.Add objPair, objPair.Key
    
    Set objPair = New KeyValuePair
    objPair.Key = "WORKPHONE"
    objPair.Value = WorkPhone
    l_Attrib.Add objPair, objPair.Key
    
    Set objPair = New KeyValuePair
    objPair.Key = "USER_NCID"
    objPair.Value = user_ncid
    l_Attrib.Add objPair, objPair.Key
    
    Set oUser = New C2T_Security.User
    With oUser
        .UserId = user_ncid
        .UserName = sUser
        .Password = sPass
        Set .Roles = l_Roles
        Set .Resources = l_Resources
        Set .Attributes = l_Attrib
        Set .Privileges = l_Privileges
    End With

    m_colUsers.Add oUser, user_ncid

    Set BuildTMIPUser = oUser
    Exit Function
ErrHandler:
    MsgBox "Unexpected error in Security Module.BuildTMIPUser: " & vbCrLf & Err.Description
End Function

Private Function GetUserInfo(ByRef User_Logon As String, _
                             ByRef user_ncid As String, _
                    Optional ByVal user_password As String, _
                    Optional ByRef HomePhone As String, _
                    Optional ByRef WorkPhone As String) As Boolean
                             
On Error GoTo ErrHandler

    'If User_Logon = vbnullstring Then Exit Function
    Dim oShared As CWShared
    Set oShared = New CWShared
    Dim App_UID As String
    Dim App_PWD As String
    Dim Framework_PWD As String
    Dim UserNCID As String
    Dim sSQL As String
    
    
    If AppUID = vbNullString Then
        Call GetSupportAppInfo(App_UID, App_PWD, Framework_PWD)
            
        If App_UID = vbNullString Or App_PWD = vbNullString Then Exit Function
        
        If bTMIPAvail Then
            oShared.SetLocalLogon App_UID, Framework_PWD
            AppUID = App_UID
            AppPWD = Framework_PWD
            
        Else
            oShared.SetLocalLogon App_UID, App_PWD
            AppUID = App_UID
            AppPWD = App_PWD
        End If
    End If
    Dim oSQL As SqlOps
    Set oSQL = New SqlOps
    
    '--- SCR 79741   Sherry Wang  10/12/2005
    If oSQL.LoginInfoChanged = True Then
        gbDBLoginInfoChanged = True
        Exit Function
    Else
        gbDBLoginInfoChanged = False
    End If
    '-- we may be going from userid or from ncid or from both.
    '-- if we are going frm userid (logon name) then
    If User_Logon <> vbNullString Then
'<< scr from NGIT.  They may be sending us an upper case user logon (sometimes)
        '--- SCR 83298   Sherry Wang   12/5/2005
        sSQL = "select user_ncid, logon_name from hems_users where upper(logon_name) = " & SQLQuote(UCase$(User_Logon))
    ElseIf user_ncid <> vbNullString Then
        sSQL = "select user_ncid, logon_name from hems_users where user_ncid = " & user_ncid
    Else
    '   -- need one or the other
        GetUserInfo = False
        Exit Function
    End If

'__ in most cases the password will be provided, except when looking for info on a non-logged in user
    If Not user_password = vbNullString And Not bTMIPAvail Then
        sSQL = sSQL & " and logon_passwd = '" & user_password & "'"
    End If
    oSQL.Execute sSQL
    If oSQL.EOF Then
        '__ this user is not a user in our system.  Don't let him in.
        GetUserInfo = False
        Exit Function
    Else
      user_ncid = vbNullString & oSQL("user_ncid")
      User_Logon = vbNullString & oSQL("logon_name")
    End If
    
    oSQL.Execute "select * from provider where ncid = " & user_ncid
    If Not oSQL.EOF Then
       HomePhone = vbNullString & oSQL("Duty_phone_2")
       WorkPhone = vbNullString & oSQL("Duty_phone_1")
    End If
    
    Set oSQL = Nothing
    GetUserInfo = True
    Exit Function
ErrHandler:
    MsgBox "Unexpected error in Security Module.getUserInfo " & vbCrLf & Err.Description, vbCritical, "Security Error"
    
End Function


Public Function BuildUser(ByVal UserLogonID As String, _
                          ByVal UserPWDHash As String, _
                                Roles As Collection, _
                                Privileges As Collection, _
                 Optional ByVal Timer As String = "120", _
                 Optional ByVal domain As String = THEATER_KEYWORD) As ICHCSII_User
                        
'- UserLogonID is the string used to log on the user  (i.e.: "sysadmin")
'  UserPWDHash is the MD5 hash of the user's password
'  Roles is a collection of strings listings the user's roles
'  Privileges is a collection of strings of the format:  function=priv-list  e.g.: Current_Encounter=RWDC

    Const AppID         As String = "CHCSII-T"
    Dim i               As Integer
    Dim CmdLine()       As String
    Dim RegDecryptKey   As String
    Dim SecInfo         As String
    Dim SecLines()      As String
    Dim KeyString       As String
    Dim ValString       As String
    Dim user_ncid       As String
    Dim oUser           As User
    Dim HomePhone       As String
    Dim WorkPhone       As String
    
    Dim l_Resources     As New Resources
    Dim l_Roles         As New Roles
    Dim l_Attrib        As New Collection  '- of KeyValuePair
    Dim l_ResData       As New Collection  '- of KeyValuePair
    Dim l_Privileges    As New Collection  '- of KeyValuePair
    Dim inKey, inQuote  As Boolean
    Dim objPair         As CHCSII_Utils.KeyValuePair
    
    Dim oShared As CWShared
    Set oShared = New CWShared
    '__ set some defaults
    m_sDomain = domain
    
'__ add users table elements:
    Set objPair = New KeyValuePair
    objPair.Key = "TIMER"
    objPair.Value = Timer
    l_Attrib.Add objPair, objPair.Key
    
    Set objPair = New KeyValuePair
    objPair.Key = "Domain"
    objPair.Value = domain
    l_Attrib.Add objPair, objPair.Key
    
    Dim vString As Variant
    For Each vString In Roles
        l_Roles.AddRole vString
    Next vString
    
    Set objPair = New KeyValuePair
    objPair.Key = "name"
    objPair.Value = UserLogonID
    l_Attrib.Add objPair, objPair.Key
    
    Set objPair = New KeyValuePair
    objPair.Key = "password"
    objPair.Value = UserPWDHash
    l_Attrib.Add objPair, objPair.Key

    
    For Each vString In Privileges
        If InStr(vString(i), "=") > 0 Then
            KeyString = UCase$(Trim(Mid$(vString(i), 1, InStr(vString(i), "=") - 1)))
            ValString = Trim(Mid$(vString(i), InStr(vString(i), "=") + 1))
            On Error Resume Next
            Set objPair = l_Privileges(KeyString)
            If Err.Number = 0 Then
                objPair.Value = AddPriv(ValString, KeyString)
            Else
                Err.Clear
                Set objPair = New KeyValuePair
                objPair.Key = KeyString
                objPair.Value = ValString
                l_Privileges.Add objPair, objPair.Key
            End If
            On Error GoTo ErrHandler
        End If
    Next vString
    
    '__ get additional info from Hems_Users table.  User MUST be there.
    Call GetUserInfo(UserLogonID, user_ncid, HomePhone, WorkPhone)
    
    '--- SCR 79741   Sherry Wang  10/12/2005
    If gbDBLoginInfoChanged = True Then Exit Function

    If user_ncid = vbNullString Then
        MsgBox "Unexpected security error:  User is not listed in Theater Database.", vbCritical, "Internal Error"
        Exit Function
    End If
    
    Set l_ResData = New Collection
    
    l_Resources.AddResource THEATER_KEYWORD, UserLogonID, UserPWDHash, l_ResData

    Set objPair = New KeyValuePair
    objPair.Key = "HOMEPHONE"
    objPair.Value = HomePhone
    l_Attrib.Add objPair, objPair.Key
    
    Set objPair = New KeyValuePair
    objPair.Key = "WORKPHONE"
    objPair.Value = WorkPhone
    l_Attrib.Add objPair, objPair.Key
    
    Set objPair = New KeyValuePair
    objPair.Key = "USER_NCID"
    objPair.Value = user_ncid
    l_Attrib.Add objPair, objPair.Key
    
    Set oUser = New C2T_Security.User
    With oUser
        .UserId = user_ncid
        .UserName = UserLogonID
        .Password = UserPWDHash
        Set .Roles = l_Roles
        Set .Resources = l_Resources
        Set .Attributes = l_Attrib
        Set .Privileges = l_Privileges
    End With

    m_colUsers.Add oUser, user_ncid

    Set BuildTMIPUser = oUser
ErrHandler:
    
End Function

Private Function ICHCSII_Security_TMIPAvailable() As Boolean
    ICHCSII_Security_TMIPAvailable = bTMIPAvail
End Function

Private Function ICHCSII_Security_ValidateUser(ByVal UserId As String, ByVal Password As String, ErrNum As Long) As Boolean
'-- return True if this userid and password are legitimate
    Dim oCrypto     As CCrypto
    Dim CWShared    As CWShared
    Dim oSQL        As SqlOps
    
    Dim suid        As String
    Dim spwd        As String
    Dim sDBK        As String
    Dim AppUID      As String
    Dim AppPWD      As String
    Dim Framework_PWD As String
    
On Error GoTo ErrHandler
    Set oCrypto = New CCrypto
    Set CWShared = New CWShared
    
'-- get and hold the current settings
    If Not CWShared.GetLocalLogon(suid, spwd, sDBK) Then
    '  returns false if nothing is returned
    End If
    
    If Not GetSupportAppInfo(AppUID, AppPWD, Framework_PWD) Then
        
    End If
    
    If bTMIPAvail Then
        AppPWD = Framework_PWD
    Else
        '-- we use our own security admin db, but make the password editable in the reg
        SecUID = "secadmin"
        SecPWD = Framework_PWD
        DBKey = vbNullString
    End If
    If SecUID = "" Or SecPWD = "" Then
        Err.Raise SEC_NO_SYSTEM_ACCESS, "C2T_Security.ValidateUser", GetErrMsg(SEC_NO_SYSTEM_ACCESS)
    End If
    
    Call CWShared.SetLocalLogon(SecUID, SecPWD)
    Set oSQL = New SqlOps
    If oSQL Is Nothing Then
        Err.Raise SEC_NO_SYSTEM_ACCESS, "C2T_Security.ValidateUser", GetErrMsg(SEC_NO_SYSTEM_ACCESS)
    End If
    Call CWShared.SetLocalLogon(suid, spwd, sDBK)
    
    oSQL.Execute "select ur.user_password password,ur.locked, ur.user_name name, ur.res_id, u.* " _
          & " FROM users u,user_resources ur, resources r" _
          & " where ur.user_name = " & SQLQuote(UserId) _
          & " AND u.user_ncid = ur.user_ncid " _
          & " and r.res_id = ur.res_id" _
          & " AND r.resource_name = '" & THEATER_KEYWORD & "'"
    
    If oSQL.EOF Then
         Err.Raise SEC_INVALID_USER, "C2T_Security.ValidateUser", GetErrMsg(SEC_INVALID_USER)
    End If
    If oSQL("locked") <> 0 Then
        Err.Raise SEC_ACCOUNT_LOCKED, "C2T_Security.ValidateUser", GetErrMsg(SEC_ACCOUNT_LOCKED)
    End If
        
    If oSQL("Password") <> oCrypto.GetHash(Password) Then
        Err.Raise SEC_INVALID_USER, "C2T_Security.ValidateUser", GetErrMsg(SEC_INVALID_USER)
    End If
        
    ICHCSII_Security_ValidateUser = True
    Exit Function
ErrHandler:
    Select Case Err.Number
    
    Case SEC_SUPPORT_UNAVAILABLE, SEC_NO_SYSTEM_ACCESS, SEC_INVALID_USER, SEC_ACCOUNT_LOCKED
        Err.Raise Err.Number, Err.Source, Err.Description
    End Select
    ICHCSII_Security_ValidateUser = False

End Function

Private Function ITMIP_AppData_GetAppData(ByVal AppName As String, _
                                                UID As String, _
                                                PWD As String, _
                                                SID As String) As Boolean
'<< SCR 31267 add registry-based uid and passwords and
'   use C2T_Security to passthrough to tmip_security if it's there

    Dim oCrypt   As New CCrypto
    Dim ePwd     As String
    Dim eUid     As String
    Dim sExp     As String
    Dim sText    As String
    Dim oShared  As New MMMHISShared.CWShared
    Dim Framework_PWD As String
    Const APP_NAME As String = "CHCSII-T"
    
    If AppUID = vbNullString Or AppPWD = vbNullString Then
        Call GetSupportAppInfo(AppUID, AppPWD, Framework_PWD)
        If bTMIPAvail Then
            AppPWD = Framework_PWD
        End If
    End If
    UID = AppUID
    PWD = AppPWD
    
    ITMIP_AppData_GetAppData = True
    Exit Function
ErrHandler:
    MsgBox "Unexpected error in Security Module.GetAppData: " & vbCrLf & Err.Description, vbCritical, "Security Error"



End Function

Private Sub ReadSecurityStubInfo()

'Reads from the security_stub.txt file to pull in dummy security info
    
    Dim propertyReader As New PropertyParser
    Dim colProperties As Collection
    Dim colUserProperties As Collection
    Dim l_Resources As New Resources
    Dim objKeyValuePair As KeyValuePair
    Dim objUser As User
    Dim sActiveUser As String
    Dim i As Integer
    'security_stub_ADM.txt
    On Error GoTo ErrHandler
    '<< look in reg for a security stub file name. if not there default to "security_stub.txt"
        Dim oShared As New CWShared
        Dim StubfileName As String
        StubfileName = oShared.GetSupportValue("StubFileName")
        Set oShared = Nothing
        If StubfileName = vbNullString Then StubfileName = "security_stub.txt"
    '>>
    Set colProperties = propertyReader.ReadProperties(m_INIPath & StubfileName)
    
    If (colProperties Is Nothing) Then
        On Error GoTo 0
        Err.Raise SEC_STUB_FILE_ERROR, CLASS_NAME, "Error with security file: " & vbCrLf & m_INIPath & StubfileName
    End If
    
    m_sDomain = colProperties.Item("domain")
    m_Roles.Initialize colProperties.Item("roles").Value
    l_Resources.Initialize colProperties.Item("resources").Value
    
    'Get Users
    Set colUserProperties = colProperties.Item("users").Value
    For i = 1 To colUserProperties.Count
        Set objKeyValuePair = colUserProperties.Item(i)
        Set objUser = New User
        objUser.Initialize objKeyValuePair.Key, objKeyValuePair.Value, l_Resources, m_Roles
        m_colUsers.Add objUser, objUser.UserId
    Next
    
    sActiveUser = colProperties.Item("active_user").Value
    If sActiveUser <> vbNullString Then
        Set m_objDefaultUser = m_colUsers.Item(sActiveUser)
    End If
    objUser.DumpObject
    Exit Sub
    
ErrHandler:
    
    Err.Raise SEC_STUB_FILE_ERROR, CLASS_NAME, GetErrMsg(SEC_STUB_FILE_ERROR) + vbCrLf + error
    Exit Sub
    Resume
End Sub

Private Function GetPwdOps() As clsPasswordOps
    
    If goPwdOps Is Nothing Then
        Set goPwdOps = New clsPasswordOps
        goPwdOps.Init bTMIPAvail
    End If
        
    Set GetPwdOps = goPwdOps
    
End Function

'SCR 102635 - LCI - 08/23/2006
Private Function GetFullUserName(sUserNCID As String) As String

Dim sSQL As String
Dim oSQL As SqlOps

Set oSQL = New SqlOps

    On Error GoTo ErrHandler

    sSQL = "select name from provider where ncid = " & sUserNCID
    oSQL.Execute sSQL
    GetFullUserName = vbNullString & oSQL("name")
    
Set oSQL = Nothing

ErrHandler:
End Function
