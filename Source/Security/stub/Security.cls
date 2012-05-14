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
Attribute VB_Name = "Security"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements ICHCSII_Security


Private Const ACTIVATE_KEY As String = "ASKJHDSFIOUYCXOIVUZXCKBASDFBASOIXZYCVKASDBFMNEQ234KYF8S7D6F98AD7F87ASD9867FQWOEGHFASKDJHGVZXCYTV;"
Private Const CLASS_NAME As String = "CHCSII_SecurityStub.Security"
Private m_sDomain        As String
Private m_bActivated     As Boolean
Private m_INIPath        As String
Private m_objDefaultUser As ICHCSII_User

Private m_colUsers       As New Collection        'Collection of User objects
Private m_Roles          As New Roles

Private Sub Class_Initialize()
    '* modified by j.ray 05SEP2001 SCR#16514; ini path was not being set before attempting to read stub file
    Dim oShared As New MMMHISShared.CWShared
    m_INIPath = Trim(oShared.INIPath)
    If Right$(m_INIPath, 1) <> "\" Then m_INIPath = m_INIPath & "\"
    Set oShared = Nothing
    
    m_bActivated = False
    ReadSecurityStubInfo
End Sub

Private Function ICHCSII_Security_Activate(ActivateKey As String) As Boolean
    If (ActivateKey = ACTIVATE_KEY) Then
        m_bActivated = True
        ICHCSII_Security_Activate = True
    Else
        m_bActivated = False
        ICHCSII_Security_Activate = False
    End If
End Function

Private Function ICHCSII_Security_ChangePassword(user As CHCSII_ClientSecurity.ICHCSII_User) As Boolean
    If (m_bActivated) Then
        MsgBox "Not currently implemented.", vbOKOnly, "Security - Change Password"
        ICHCSII_Security_ChangePassword = True
    Else
        Err.Raise SEC_NOT_ACTIVATED, CLASS_NAME, SEC_NOT_ACTIVATED_MSG
    End If
End Function

Private Property Get ICHCSII_Security_Domain() As String
    If (m_bActivated) Then
        ICHCSII_Security_Domain = m_sDomain
    Else
        Err.Raise SEC_NOT_ACTIVATED, CLASS_NAME, SEC_NOT_ACTIVATED_MSG
    End If
End Property

Private Function ICHCSII_Security_GetActiveUser(Optional sFormTitle As String, Optional sWindowTitle As String) As CHCSII_ClientSecurity.ICHCSII_User
    If (m_bActivated) Then
        If (m_objDefaultUser Is Nothing) Then
            Set ICHCSII_Security_GetActiveUser = GetUserFromLogin(sFormTitle, sWindowTitle)
        Else
            Set ICHCSII_Security_GetActiveUser = m_objDefaultUser
        End If
    Else
        Err.Raise SEC_NOT_ACTIVATED, CLASS_NAME, SEC_NOT_ACTIVATED_MSG
    End If
End Function

Private Function ICHCSII_Security_GetPrivilegesByRole(l_User As CHCSII_ClientSecurity.ICHCSII_User, l_Role As CHCSII_ClientSecurity.ICHCSII_Role) As Collection
    If (m_bActivated) Then
        Dim colPrivileges As New Collection
        Dim userRole As Role
        Dim sysRole As Role
        Dim i As Integer
        
        On Error Resume Next
        'Check to see if user has role
        Set userRole = l_User.Roles.Item(l_Role.Name)
        If Err Then
            On Error GoTo 0
            Err.Raise SEC_NO_USER_ROLE, CLASS_NAME, SEC_NO_USER_ROLE_MSG
        End If
        
        'Get system role so we can get privileges
        Set sysRole = m_Roles.Item(l_Role.Name)
        If Err Then
            'Shouldn't happen, but just in case it does...
            On Error GoTo 0
            Err.Raise SEC_NO_SYSTEM_ROLE, CLASS_NAME, SEC_NO_SYSTEM_ROLE_MSG
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
        Err.Raise SEC_NOT_ACTIVATED, CLASS_NAME, SEC_NOT_ACTIVATED_MSG
    End If
End Function

Private Function ICHCSII_Security_GetUser(Optional sUserID As String) As CHCSII_ClientSecurity.ICHCSII_User
    If (m_bActivated) Then
        On Error Resume Next
        'date:01/07/2004 SCR #:47772 developer: jrm Description: Added for ITT mode
        'ITT_USER_NAME is needed to determine ITT, or any other mode
        If m_objDefaultUser.UserId = ITT_USER_NAME Then
            Set ICHCSII_Security_GetUser = m_colUsers.Item(ITT_USER_NAME)
        Else
            Set ICHCSII_Security_GetUser = m_colUsers.Item(sUserID)
        End If
        If Err Then
            On Error GoTo 0
            Err.Raise SEC_USER_NOT_FOUND, CLASS_NAME, SEC_USER_NOT_FOUND_MSG
        End If
    Else
        Err.Raise SEC_NOT_ACTIVATED, CLASS_NAME, SEC_NOT_ACTIVATED_MSG
    End If
End Function

Private Function ICHCSII_Security_Logoff(user As CHCSII_ClientSecurity.ICHCSII_User) As Boolean
    If (m_bActivated) Then
        ICHCSII_Security_Logoff = True
    Else
        Err.Raise SEC_NOT_ACTIVATED, CLASS_NAME, SEC_NOT_ACTIVATED_MSG
    End If
End Function

Private Function ICHCSII_Security_Logon(UserId As String, Password As String) As CHCSII_ClientSecurity.ICHCSII_User
    Dim l_User As ICHCSII_User
    
    If (m_bActivated) Then
        On Error Resume Next
        Set l_User = m_colUsers.Item(UserId)
        If Err Then
            Err.Clear
            On Error GoTo 0
            Err.Raise SEC_USER_NOT_FOUND, CLASS_NAME, SEC_USER_NOT_FOUND_MSG
        Else
            On Error GoTo 0
            If (l_User.ValidatePassword(Password)) Then
                Set ICHCSII_Security_Logon = l_User
            Else
                Err.Raise SEC_INVALID_PASSWORD, CLASS_NAME, SEC_INVALID_PASSWORD_MSG
            End If
        End If
    Else
        Err.Raise SEC_NOT_ACTIVATED, CLASS_NAME, SEC_NOT_ACTIVATED_MSG
    End If
End Function

Private Sub ReadSecurityStubInfo()

'Reads from the security_stub.txt file to pull in dummy security info

Dim propertyReader As New PropertyParser
Dim colProperties As Collection
Dim colUserProperties As Collection
Dim l_Resources As New Resources
Dim objKeyValuePair As KeyValuePair
Dim objUser As user
Dim sActiveUser As String
Dim i As Integer

'date:01/22/2004 SCR #:47772 developer: jrm Description: Added for ITT mode
'jrm:ITT change default clinics and such based on role selected
Dim Connection As CHCSII_CONN.Conn
Dim Conn As ICHCSII_SQL.ISqlOps

'security_stub_ADM.txt
On Error GoTo ErrHandler
'<< look in reg for a security stub file name. if not there default to "security_stub.txt"
    Dim oShared As New CWShared
    Dim StubfileName As String
    'StubfileName = oShared.GetSupportValue("StubFileName")
    Set oShared = Nothing
    
    'date:01/22/2004 SCR #:47772 developer: jrm Description: Added for ITT mode
    'Display roles identity selector for ITT
    If oShared.IsAppMode(modeITT) = True And oShared.CmdLineSwitches.Exists("R2") = True Then
    
        Set Connection = New CHCSII_CONN.Conn
        Set Conn = Connection.CHCSII_SQLOPS
        
        Load frmRoleDialog
        frmRoleDialog.Show (vbModal)
        
        If frmRoleDialog.OptionRole = 0 Then
             StubfileName = "security_stub.txt"
             Conn.Execute ("update osh_ini_settings set apvalue = '1047450' where application_ncid = 10001 and apkey = 'CLINIC_NCID' and apsection = 'LOCAL_ENV'")
             Conn.Execute ("update osh_ini_settings set apvalue = 'CHCSII_APPOINTMENTS.APPOINTMENTS' where apownerncid = 99990004 and apsection = 'startup option'")
        ElseIf frmRoleDialog.OptionRole = 1 Then
             StubfileName = "security_stub.dtn"
             Conn.Execute ("update osh_ini_settings set apvalue = '1050010' where application_ncid = 10001 and apkey = 'CLINIC_NCID' and apsection = 'LOCAL_ENV'")
             Conn.Execute ("update osh_ini_settings set apvalue = 'CHCSII_APPOINTMENTS.APPOINTMENTS' where apownerncid = 99990004 and apsection = 'startup option'")
       Else
             StubfileName = "security_stub.srts"
             Conn.Execute ("update osh_ini_settings set apvalue = '1070010' where application_ncid = 10001 and apkey = 'CLINIC_NCID' and apsection = 'LOCAL_ENV'")
              Conn.Execute ("update osh_ini_settings set apvalue = 'CHCSII_APPOINTMENTS.APPOINTMENTS' where apownerncid = 99990004 and apsection = 'startup option'")
        End If
        
        Set Connection = Nothing
        Set Conn = Nothing
        
    End If
    
    
    
    
    If StubfileName = "" Then StubfileName = "security_stub.txt"
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
    Set objUser = New user
    objUser.Initialize objKeyValuePair.Key, objKeyValuePair.Value, l_Resources, m_Roles
    m_colUsers.Add objUser, objUser.UserId
Next

sActiveUser = colProperties.Item("active_user").Value
If sActiveUser <> "" Then
    Set m_objDefaultUser = m_colUsers.Item(sActiveUser)
End If

Exit Sub

ErrHandler:

Err.Raise SEC_STUB_FILE_ERROR, CLASS_NAME, SEC_STUB_FILE_ERROR_MSG + vbCrLf + Error

End Sub

Private Function GetUserFromLogin(sFormTitle As String, sWindowTitle As String) As ICHCSII_User
    
Dim bSuccess As Boolean
Dim iNumTries As Integer
Dim l_User As ICHCSII_User

Do While bSuccess = False
    Load frmCHCSIILogin
    frmCHCSIILogin.FormTitle = sFormTitle
    frmCHCSIILogin.WindowTitle = sWindowTitle
    Screen.MousePointer = vbDefault
    frmCHCSIILogin.Show vbModal
    
    If frmCHCSIILogin.HasCancelled() Then
        Unload frmCHCSIILogin
        Err.Raise SEC_USER_CANCELLED_LOGIN, CLASS_NAME, SEC_USER_CANCELLED_LOGIN_MSG
    End If
    
    On Error Resume Next
    Set l_User = m_colUsers.Item(frmCHCSIILogin.GetUserId())
    If Err Then
        Err.Clear
        On Error GoTo 0
        If iNumTries <> 3 Then
            iNumTries = iNumTries + 1
            frmCHCSIILogin.ShowError "User Id not found."
        Else
            frmCHCSIILogin.ShowError "User Id not found. Maximum number of attempts has been made."
            Unload frmCHCSIILogin
            Err.Raise SEC_MAX_LOGIN_ATTEMPTS_MET, CLASS_NAME, SEC_MAX_LOGIN_ATTEMPTS_MET_MSG
        End If
    Else
        On Error GoTo 0
        If l_User.ValidatePassword(frmCHCSIILogin.GetPassword()) Then
            bSuccess = True
            Set GetUserFromLogin = l_User
            Unload frmCHCSIILogin
        Else
            If iNumTries <> 3 Then
                iNumTries = iNumTries + 1
                frmCHCSIILogin.ShowError "Password is incorrect."
            Else
                frmCHCSIILogin.ShowError "Password is incorrect. Maximum number of attempts has been made."
                Unload frmCHCSIILogin
                Err.Raise SEC_MAX_LOGIN_ATTEMPTS_MET, CLASS_NAME, SEC_MAX_LOGIN_ATTEMPTS_MET_MSG
            End If
        End If
    End If
Loop
    
End Function

Private Function ICHCSII_Security_TMIPAvailable() As Boolean
    ICHCSII_Security_TMIPAvailable = False
End Function

Private Function ICHCSII_Security_ValidateUser(ByVal UserId As String, ByVal Password As String, ErrNum As Long) As Boolean
    ICHCSII_Security_ValidateUser = False
End Function
