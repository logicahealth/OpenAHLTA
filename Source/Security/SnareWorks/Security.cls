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

'=========================current build with CAC support

Implements ICHCSII_Security
Implements CHCSII_ClientSecurity.ITMIP_AppData

Private m_CardInfo      As CardInfo_

Private SnareHandle     As Long
Private WithEvents CAC_Monitor As cCardMonProxy
Attribute CAC_Monitor.VB_VarHelpID = -1
Private m_CAC_Inserted  As Boolean
Private m_TO_Proxy      As TO_Proxy
Private m_LoginMode     As CHCSII_LoginMode
Private Const CLASS_NAME As String = "CHCSII_SnareWorks.Security"
Private m_sDomain       As String
Private m_bActivated    As Boolean
Private m_colUsers      As New Collection        'Collection of User objects
Private m_Roles         As New Roles
Private objPassword     As New PassWordRules
Private NOCAC           As Boolean
Private frmLogin        As frmCHCSIILogin

Public Property Get Login_Mode() As String
    Select Case m_LoginMode
    Case Mode_CAC:       Login_Mode = "CAC"
    Case Mode_Manual:    Login_Mode = "MANUAL"
    End Select
End Property

Public Property Get CAC_Inserted() As Boolean
    CAC_Inserted = m_CAC_Inserted
End Property

Public Property Get CAC_Card() As TO_Proxy
    Set CAC_Card = m_TO_Proxy
End Property

Private Sub CAC_Monitor_CardInserted(sCardName As String)
    On Error GoTo errhandler
    If m_CardInfo.CardCert = "" Then    '-- no card cert means not logged in
        m_CAC_Inserted = True
        m_LoginMode = Mode_CAC
        m_CardInfo.CardName = sCardName
        
    ElseIf m_CardInfo.CardName = sCardName Then
        '__ I'm logged in and this may be my card coming back into the reader
        'Call CAC_Reauthenticate
        Call m_TO_Proxy.CardInserted
        
    End If
    Exit Sub
errhandler:
    err.Raise err.Number, err.Source, err.Description
    
End Sub

Private Sub CAC_Monitor_CardRemoved(sCardName As String)
    On Error GoTo errhandler
    If sCardName = m_CardInfo.CardName Then     '-- ignore if not for my card
        m_CAC_Inserted = False
        m_LoginMode = Mode_Manual
        If Not frmLogin Is Nothing Then
            If NOCAC Then
            '-- we are in the middle of a login.  the card was pulled. stop everything
                frmLogin.cmdCancel.Value = True
            End If
        End If
        Call m_TO_Proxy.CardRemoved
    End If
    Exit Sub
errhandler:
    err.Raise err.Number, err.Source, err.Description

End Sub

Private Sub Class_Initialize()
    On Error GoTo errhandler
    m_bActivated = False
    ReDim FuncAreaName(0)
    
    Set m_TO_Proxy = New TO_Proxy
    Set m_TO_Proxy.CardMonitor = CAC_Monitor
    
'<< scr 47904
    gOnSnareServer = Not (Val(Environ$("ENCRYPT")) = 1)
'<< CAC
    If Not gOnSnareServer Then
        Set CAC_Monitor = New cCardMonProxy
    End If
    gSnareDialogSelected = SN_DialogNone
    'gSnareDialogSelected = SN_Dialoglogin
    Exit Sub
errhandler:
    err.Raise err.Number, err.Source, err.Description
    
End Sub

Private Sub Class_Terminate()
    Set m_colUsers = Nothing
End Sub

Private Function GetSnareHandle(Optional ByRef ErrNum As Long) As Long
'<< SCR 47904 MM logon problems. special case for "on the security server"
'   Indicator for 'on the server' = environ$("ENCRYPT") is not 1.  this is set in class init
'   if on the server, use the global handle
'   if on a client workstation, use the local handle (allows multiple user sessions)

    If gOnSnareServer Then
        If gSnareHandle = 0 Then
            gSnareHandle = SN_GetSnareHandle(ErrNum)
            SnareHandle = gSnareHandle
        End If
        SnareHandle = gSnareHandle
        GetSnareHandle = gSnareHandle
        
    Else
        If SnareHandle = 0 Then
            SnareHandle = SN_GetSnareHandle(ErrNum)
            gSnareHandle = SnareHandle
        End If
        gSnareHandle = SnareHandle
        GetSnareHandle = SnareHandle
    End If
End Function
Private Function ICHCSII_Security_Activate(ActivateKey As String) As Boolean

    Dim lErrNum As Long
    Dim hHandle As Long
    hHandle = GetSnareHandle(lErrNum)
'<< SCR 9633
    If hHandle = -1 Then  '- user cancelled
        m_bActivated = False
    ElseIf hHandle = 0 Then
        If lErrNum = 1004 Then
            MsgBxARMd "No SnareWorks security server found. Please contact your Administrator"
        ElseIf lErrNum <> 0 Then
            MsgBxARMd "Error connecting to security server." & vbCrLf & GetErrMsg(lErrNum)
        End If
        m_bActivated = False
    Else
        m_bActivated = True
    End If
    ICHCSII_Security_Activate = m_bActivated
    
End Function

Private Function ICHCSII_Security_ChangePassword(User As CHCSII_ClientSecurity.ICHCSII_User) As Boolean

    On Error GoTo errhandler
'-- mode is one of "CHANGE" or "EXPIRED"
    Dim OldUID As String
    Dim OldPWD As String
    Dim NewUID As String
    Dim NewPwd As String
    Dim frm As frmPassword
    
    If Not (m_bActivated) Then
        err.Raise SEC_NOT_ACTIVATED, CLASS_NAME, GetErrMsg(SEC_NOT_ACTIVATED)
    End If
    
'setup the new password form

    Set frm = New frmPassword
    Set frm.currentUser = User
    frm.FrmMode = "CHANGE"
    
    If Not frm.SetPasswordRules(objPassword) Then
        err.Raise SEC_GENERAL_ERROR, "Security.ChangePassword", "Password Rules Not Found"
        Exit Function
    End If
    
    frm.Top = (Screen.Height - frm.Height) / 2
    frm.Left = (Screen.Width - frm.Width) / 2
    Screen.MousePointer = vbDefault
    
    frm.Show vbModal
    
    If frm.bUserCanceled Then
        ICHCSII_Security_ChangePassword = True
        Exit Function
    End If
    
    OldUID = User.UserID
    OldPWD = frm.txtOldPWD.Text
    NewPwd = frm.txtNewPWD1.Text
    NewUID = frm.NewUID
    '--be sure we're loggedon
    If Not GetUser = User.UserID Then

        MsgBxARMd "Unable to validate the current user.  Please contact your system administrator." & vbCrLf _
            & "Password not changed.", vbCritical, "Change Password"
        Exit Function

    End If
'__ we have the right user
    Unload frm
    Dim sErr As String  '- the error that comes back
    
    If NewPwd <> vbNullString And OldPWD <> NewPwd Then
        If Not ChangePass(OldUID, OldPWD, NewPwd, sErr) Then
             MsgBxARMd "Unable to change the password." & vbCrLf _
                & sErr & vbCrLf & "The password was not changed.", vbCritical, "Change Password"
        End If
    End If
    
    If NewUID <> vbNullString And OldUID <> NewUID Then
        If Not ChangeUserID(OldUID, NewUID, sErr) Then
            MsgBxARMd "Unable to change the user Id." & vbCrLf _
            & sErr & vbCrLf & "User Id was not changed.", vbCritical, "Change User ID"
        End If
    End If
    Exit Function
errhandler:
    err.Raise err.Number, err.Source, err.Description
    
End Function

Private Property Get ICHCSII_Security_Domain() As String
    If (m_bActivated) Then
        ICHCSII_Security_Domain = m_sDomain
    Else
        err.Raise SEC_NOT_ACTIVATED, CLASS_NAME, GetErrMsg(SEC_NOT_ACTIVATED)
    End If
End Property

Private Function ICHCSII_Security_GetActiveUser(Optional sFormTitle As String, Optional sWindowTitle As String) As CHCSII_ClientSecurity.ICHCSII_User
    Dim sUser As String
    Dim lErrNum As Long
    Dim ErrMsg As String
    
    On Error GoTo errhandler
    If (m_bActivated) Then
        
        ' See if snare user object exists ?
        sUser = GetUser(lErrNum)
        If lErrNum <> 1007 And lErrNum <> 0 Then   ' 1007 is "there is no user"
            err.Raise lErrNum, "CHCSII Security", GetErrMsg(lErrNum)
            Exit Function
        End If
        
        lErrNum = 0
        If sUser <> vbNullString Then
            Set ICHCSII_Security_GetActiveUser = GetSnareUser(sUser, True, lErrNum, ErrMsg) '__ get the user info
            If lErrNum <> 0 Then
                err.Raise lErrNum, "Get Snare User (XML)", ErrMsg
            End If
        ElseIf Not gOnSnareServer Then
            Set ICHCSII_Security_GetActiveUser = GetUserFromLogin()  '__ get user info from user
        End If
    Else
        err.Raise SEC_NOT_ACTIVATED, CLASS_NAME, GetErrMsg(SEC_NOT_ACTIVATED)
    End If
    
    Exit Function
errhandler:
    If err.Number = -2147220397 Then            '-- user canceled logon
        err.Raise -2147220397, "CHCSII Security", "User Canceled Logon"
    Else
        err.Raise err.Number, err.Source, err.Description
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
        If err Then
            err.Clear
            On Error GoTo 0
            'err.Raise SEC_NO_USER_ROLE, CLASS_NAME, GetErrMsg(SEC_NO_USER_ROLE)
        End If
        
        'Get system role so we can get privileges
        Set sysRole = m_Roles.Item(l_Role.Name)
        If err Then
             err.Clear
            'Shouldn't happen, but just in case it does...
            On Error GoTo 0
            err.Raise SEC_NO_SYSTEM_ROLE, CLASS_NAME, GetErrMsg(SEC_NO_SYSTEM_ROLE)
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
        err.Raise SEC_NOT_ACTIVATED, CLASS_NAME, GetErrMsg(SEC_NOT_ACTIVATED)
    End If
End Function

Private Function ICHCSII_Security_Logoff(User As CHCSII_ClientSecurity.ICHCSII_User) As Boolean
    If (m_bActivated) Then
        Logoff
        ICHCSII_Security_Logoff = True
    Else
        err.Raise SEC_NOT_ACTIVATED, CLASS_NAME, GetErrMsg(SEC_NOT_ACTIVATED)
    End If
End Function

Private Function ICHCSII_Security_Logon(UserID As String, Password As String) As CHCSII_ClientSecurity.ICHCSII_User
    Dim l_User As ICHCSII_User
    Dim nErrCode As Long
    Dim nSecErrCode As Long
    Dim sErrMsg As String
    
    If (m_bActivated) Then
    
'<< scr 45817 better error mesages from snareworks
'   User the secondary error message if available
        ' Do SnareWorks Logon
        If SW_Logon(GetSnareHandle, UserID, Password, nErrCode, nSecErrCode) Then
             Set ICHCSII_Security_Logon = GetSnareUser(UserID, True, nErrCode, sErrMsg)
             If nErrCode <> 0 Then
                err.Raise nErrCode, "Get Snare User (XML)", sErrMsg
             End If
        Else
            If nErrCode = 1000 And nSecErrCode <> 0 Then
                err.Raise nErrCode, CLASS_NAME, GetErrMsg(nSecErrCode)
            Else
                err.Raise nErrCode, CLASS_NAME, GetErrMsg(nErrCode)
            End If
        End If
    Else
        err.Raise SEC_NOT_ACTIVATED, CLASS_NAME, GetErrMsg(SEC_NOT_ACTIVATED)
    End If
End Function

Private Function GetUserFromLogin() As ICHCSII_User
    
    Dim bSuccess    As Boolean
    Dim iNumTries   As Integer
    Dim nErrCode    As Long
    Dim nSecErrCode As Long
    Dim sErrMsg     As String
    Dim oUser       As ICHCSII_User
    Dim frm         As frmSplash
    Set frm = New frmSplash
    
    Dim UID         As String
    Dim PWD         As String
    
    
On Error GoTo errhandler
    If CAC_Monitor.CardPresent Then
        m_LoginMode = Mode_CAC
        m_CAC_Inserted = True
    Else
        m_LoginMode = Mode_Manual
        m_CAC_Inserted = False
    End If
    Screen.MousePointer = vbDefault
    frm.Move (Screen.Width - frm.Width) / 2, (Screen.Height - frm.Height) / 2
        frm.Show vbModal
    If Not frm.Accept Then
        Unload frm
        err.Raise SEC_USER_CANCELLED_LOGIN, CLASS_NAME, GetErrMsg(SEC_USER_CANCELLED_LOGIN)
        Exit Function
    End If
    Unload frm
    
    Do While bSuccess = False
        nErrCode = 0
        Set frmLogin = New frmCHCSIILogin
        frmLogin.NOCAC = NOCAC
        frmLogin.CACAvailable = CAC_Monitor.ReaderPresent
        
        Set frmLogin.oSecurity = Me
        frmLogin.Move (Screen.Width - frmLogin.Width) / 2, (Screen.Height - frmLogin.Height) / 2
        frmLogin.Show vbModal
        
        If frmLogin.HasCancelled() Then
            Unload frmLogin
            err.Raise SEC_USER_CANCELLED_LOGIN, CLASS_NAME, GetErrMsg(SEC_USER_CANCELLED_LOGIN)
        ElseIf frmLogin.FirstTimeuser Then
            Unload frmLogin
            Set GetUserFromLogin = FirstTimeuser()
            Exit Function
        End If
        
        UID = frmLogin.UserID
        PWD = frmLogin.Password

        If iNumTries <= 3 Then
            iNumTries = iNumTries + 1
            nSecErrCode = 0
            nErrCode = 0
            
            m_LoginMode = frmLogin.Login_Mode
            Select Case m_LoginMode
            Case Mode_Manual
            '__ standard login
                If SW_Logon(SnareHandle, UID, PWD, nErrCode, nSecErrCode) Then
                    
                    If NOCAC And m_CardInfo.CardCert <> "" Then
                    '-- we have a cert that needs to be registered with snareworks
                        Call SW_RegisterCert(SnareHandle, m_CardInfo, nErrCode, nSecErrCode)
                    End If
                
                    Set GetUserFromLogin = GetSnareUser(frmLogin.UserID, True, nErrCode, sErrMsg)
                    
                    If nErrCode <> 0 Then
                        err.Raise nErrCode, "Get Snare User (XML)", sErrMsg
                    Else
                        Unload frmLogin
                        bSuccess = True
                    End If
                Else
                    If nErrCode <> 0 Then
                        If nErrCode = 1000 And nSecErrCode <> 0 Then
                            frmLogin.ShowError "Message from Security Server:" & vbCrLf & vbCrLf & GetErrMsg(nSecErrCode)
                        Else
                            frmLogin.ShowError "Message from Security Server:" & vbCrLf & vbCrLf & GetErrMsg(nErrCode)
                            Select Case nErrCode
                            Case 1002           'SN_PASSMUSTCHANGE
                               Set oUser = New User
                               If UserMustChangePassword(UID, PWD, oUser) Then
                                    Unload frmLogin   '-- added via scr 50464
                                    bSuccess = True
                               End If
                            End Select
                        End If
                    Else
                        frmLogin.ShowError "User Id or Password not found."
                    End If
                End If
                
            Case Mode_CAC
            '__ cac login
                Set GetUserFromLogin = CAC_Snare_Logon(frmLogin.PIN, nErrCode, sErrMsg)
                
                If nErrCode = 1030 Then
                    frmLogin.ShowError "The certificate on the CAC is not registered with CHCSII.  Please logon using your CHCSII user ID and password to register this CAC."
                    NOCAC = True
                ElseIf sErrMsg <> "" Then
                    frmLogin.ShowError "Message from CAC Card: " & vbCrLf & vbCrLf & sErrMsg
                    If nErrCode = -30021 Then  '--- card is locked no further attempts allowed (scr #50530)
                        Unload frmLogin
                        Set GetUserFromLogin = Nothing
                        bSuccess = True
                    End If
                Else
                    Unload frmLogin
                    bSuccess = True
                End If
            End Select
            
        Else
            frmLogin.ShowError "User Id or password not found. Maximum number of attempts has been made."
            Unload frmLogin
            Set GetUserFromLogin = Nothing
            bSuccess = True
            err.Raise SEC_MAX_LOGIN_ATTEMPTS_MET, CLASS_NAME, GetErrMsg(SEC_MAX_LOGIN_ATTEMPTS_MET)
        End If
    Loop
    Exit Function
errhandler:
    err.Raise err.Number, err.Source, err.Description
End Function

Private Function UserMustChangePassword(ByVal UID As String, _
                                        ByVal PWD As String, _
                                              User As User) As Boolean
    '-- validate parameters
    If UID = vbNullString Or _
       PWD = vbNullString Or _
       User Is Nothing Then
       Exit Function
    End If
    
    Dim sErrMsg As String
    Dim nErrCode As Long
    Dim sSecErrCode As String
    Dim lRes As Long
  
    'Set User = GetSnareUser(UID, False, nErrCode, sErrMsg)
    Dim oPWDRules As New PassWordRules
    oPWDRules.SetDefaultProperties

    
    Dim frm As frmPassword
    Set frm = New frmPassword
    User.UserID = UID
    User.UserName = UID
    User.Password = PWD
    Set frm.currentUser = User
    Call frm.SetPasswordRules(oPWDRules)
    'frm.txtUID.text = UID
    frm.FrmMode = "EXPIRED"
    frm.Move (Screen.Width - frm.Width) / 2, (Screen.Height - frm.Height) / 2
    frm.Show vbModal
    If frm.bUserCanceled Then
        Exit Function
    End If
    If frm.txtNewPWD1 = "" Then
        Exit Function
    End If
    
    Dim NewPwd As String
    NewPwd = frm.txtNewPWD1.Text
    
    lRes = loginSnareAPI(SnareHandle, UID, PWD, NewPwd, 0, 0, 0)
    
    If lRes = 0 Then
        UserMustChangePassword = True
        Set User = GetSnareUser(UID, True, nErrCode, sErrMsg)
        If Not nErrCode = 0 Then
            err.Raise nErrCode, "User Must Change Password", "Error setting new password: " & sErrMsg
        End If
    Else
        sErrMsg = Space(1000)
        nErrCode = errorSnareAPI(SnareHandle, lRes, sErrMsg)
        UserMustChangePassword = False
        err.Raise lRes, "User Must Change Password", "Error setting new password: " & sErrMsg
    End If
    

End Function

Private Function CAC_Snare_Logon(ByVal PIN As String, _
                       Optional ByRef ErrorCode As Long, _
                       Optional ByRef ErrorMessage As String) As ICHCSII_User

    Dim oCrypto As New CCrypto
    Dim str As String
    Dim Sig As String
    Dim Cert As String
    Dim UserID As String
    Dim oUser As ICHCSII_User
    Dim i       As Long
    Screen.MousePointer = vbHourglass
    ErrorMessage = ""
    '-- need a random string to sign in with
    '   for our puposes, I want a hash of a real string
    
    Dim rndstr As String
    Randomize (CDbl(Now))
    rndstr = ""
    For i = 1 To 128
        rndstr = rndstr & Chr(Int(Rnd * 74) + 48)
    Next i
    
    str = oCrypto.GetHash(rndstr, True)
    str = str & str
    str = str & str
    
    If CAC_Monitor.CHCSII_Login(PIN, str, Sig, Cert, ErrorMessage) Then
        m_CardInfo.CardCert = Cert
        m_CardInfo.CardRndStr = str
        m_CardInfo.CardSig = Sig
    Else
        Screen.MousePointer = vbDefault
        Exit Function
    End If
        
    If Not sw_CertLogon(SnareHandle, m_CardInfo, ErrorCode) Then
        ErrorMessage = GetErrMsg(ErrorCode)
    Else
        Set oUser = ICHCSII_Security_GetActiveUser
        
        If Not oUser Is Nothing Then
            m_CardInfo.UserID = oUser.UserID
            m_CardInfo.UserName = oUser.UserName
            m_CardInfo.UserNCID = oUser.getAttribute("USER_Ncid")
        End If

        Set CAC_Snare_Logon = oUser
    End If
    
    Screen.MousePointer = vbDefault
End Function

Private Function GetSnareUser(sLoginUser As String, _
               Optional ByVal GetResInfo As Boolean = True, _
               Optional ByRef ErrorCode As Long, _
               Optional ByRef ErrorMessage As String, _
               Optional ByVal NCID As String) As ICHCSII_User

'<< if the user is asking about someone else it does so by ncid;
'   and snareworks doesn't give us the ncid as an attribute in that case.
'   So, since the user asked for info by ncid, we already have it and will pass it in and use it here
    Dim l_Resources     As Resources
    Dim objsnareuser    As User
    Dim l_Roles         As New Roles
    Dim l_Attrib        As New Collection
    Dim i               As Long
    Dim j               As Long
'__ placeholders
    Dim sName           As String
    Dim sUser_ID        As String
    Dim sData           As String
    Dim sCHCSPass       As String
    Dim sCHCSUID        As String
    Dim aData()         As String
    Dim sPassword       As String
    Dim sDelimiterLoc   As Long
    Dim oXMLDom         As DOMDocument40
    Dim oNode           As IXMLDOMNode
    Dim oSubNode        As IXMLDOMNode
    Dim oSubNodes       As IXMLDOMNodeList
    Dim oSub2Nodes      As IXMLDOMNodeList
    Dim oSub2Node       As IXMLDOMNode
    Dim UserRes         As SW_Res
    Dim colRes          As Collection       '-- of sw_res objects
    Set colRes = New Collection
    Dim XMLRequest      As String
    Dim LoginTarget     As String           '-- usually of the form appsvr.targetdomain.com
 
    On Error GoTo errhandler
    
'-- LogAccess "GetSnareUser: sloginuser " & sLoginUser
    If Not CBool(InStr(sLoginUser, "@")) And m_sDomain <> vbNullString Then
        sLoginUser = m_sDomain & "@" & sLoginUser
        Debug.Print "modified user id to " & sLoginUser
    End If
    
'-- see if we already have this guy in the collection
    If GetUserFromCollection(objsnareuser, sLoginUser) Then
        Exit Function
    End If
    
    XMLRequest = "<XMLRequest><ShowEntries><Entry>" _
               & "<User><Name>" & sLoginUser & "</Name></User>" _
               & "</Entry></ShowEntries></XMLRequest>" '
    
    If Not SnarexML(SnareHandle, XMLRequest, oXMLDom, ErrorCode, ErrorMessage) Then
        Exit Function
    End If
    
    Set l_Resources = New Resources
    
    Set oNode = oXMLDom.selectSingleNode("//User/*")
    Do While True
        Select Case LCase$(oNode.nodeName)
        '__ these have no action and I want to hit them at the top
       
        Case "name":                        Call NodeToColl("User_ID", oNode.nodeTypedValue, l_Attrib)
                                            sUser_ID = oNode.nodeTypedValue
        
        Case "snareprivatekey":             '__ no op  - don't give this out
        
        Case "accountname":                 Call NodeToColl("Account_Name", oNode.nodeTypedValue, l_Attrib)
                                            sName = oNode.nodeTypedValue
        Case "firstname":                   Call NodeToColl("First_Name", oNode.nodeTypedValue, l_Attrib)
        Case "lastname":                    Call NodeToColl("Last_Name", oNode.nodeTypedValue, l_Attrib)
        Case "email":                       Call NodeToColl("EMail", oNode.nodeTypedValue, l_Attrib)
        Case "mtf-user_ien":                Call NodeToColl("USER_IEN", oNode.nodeTypedValue, l_Attrib)
        Case "mtf-user_ncid":               Call NodeToColl("USER_Ncid", oNode.nodeTypedValue, l_Attrib)
        Case "snarepublickey":              Call NodeToColl("Public_Key", oNode.nodeTypedValue, l_Attrib)
        Case "ssn":                         Call NodeToColl("SSN", oNode.nodeTypedValue, l_Attrib)
                
        Case "fullname":                    sLoginUser = oNode.nodeTypedValue
        Case "organization":                m_sDomain = oNode.nodeTypedValue
                                            gDomain = m_sDomain
        Case "password":                    sPassword = oNode.nodeTypedValue
        Case "snareidletimer":              Call NodeToColl("Timer", oNode.nodeTypedValue, l_Attrib)
        Case "lastloginresult"
                Set oSubNodes = oNode.selectNodes("//AuthenticationInfo")
                For i = 1 To oSubNodes.length
                    Set oSubNode = oSubNodes.nextNode
                    Set oSub2Nodes = oSubNode.childNodes
                     For j = 1 To oSub2Nodes.length
                        Set oSub2Node = oSub2Nodes.nextNode
                        Select Case LCase$(oSub2Node.nodeName)
                        Case "loginsource"
                            Call NodeToColl("LOGINSource", oSub2Node.nodeTypedValue, l_Attrib)
                        Case "logintarget"
                            Call NodeToColl("LOGINTarget", oSub2Node.nodeTypedValue, l_Attrib)
                            LoginTarget = GetLoginTarget(oSub2Node.nodeTypedValue)
                    'debug.print "logintarget " & LoginTarget
                        End Select
                    Next j
                Next i
        
        Case "primarygroup"
            If InStr(oNode.nodeTypedValue, "@") > 0 Then
                l_Roles.AddRole Mid$(oNode.nodeTypedValue, InStr(oNode.nodeTypedValue, "@") + 1)
            Else
                l_Roles.AddRole oNode.nodeTypedValue
            End If
            
        Case "passwordmappings"
            Set oSubNodes = oNode.selectNodes("//PasswordMapping")
            For i = 1 To oSubNodes.length
                Set oSubNode = oSubNodes.nextNode
                Set oSub2Nodes = oSubNode.childNodes
                Set UserRes = New SW_Res
                For j = 1 To oSub2Nodes.length
                    Set oSub2Node = oSub2Nodes.nextNode
                    
                    
                    Select Case LCase$(oSub2Node.nodeName)
                    Case "applicationname"
                        Select Case oSub2Node.nodeTypedValue
                        Case "3m":      UserRes.AppName = "3MCareInnovation"
                        Case "chcs":    UserRes.AppName = "CHCS"
                        Case Else
                           UserRes.AppName = oSub2Node.nodeTypedValue
                        End Select
                    Case "applicationuser"
                        UserRes.UID = oSub2Node.nodeTypedValue
                    Case "applicationpassword"
                        UserRes.PWD = oSub2Node.nodeTypedValue
                    Case "applicationtarget"
                        UserRes.Target = oSub2Node.nodeTypedValue
                    Case "applicationscript"
                    
                    End Select
                Next j
                
                colRes.Add UserRes
            Next i
        
'<< scr 49190 handle the multiple resources for 'any' and specific domain
            For Each UserRes In colRes
                Select Case UserRes.AppName
                Case "CHCS"
                    '__ hold for later
                    sCHCSPass = UserRes.PWD
                    sCHCSUID = UserRes.UID
                Case "oracle"
                    l_Resources.AddResource "oracle", UserRes.UID, UserRes.PWD, vbNullString
                Case "bo"
                    l_Resources.AddResource "Reports", UserRes.UID, UserRes.PWD, vbNullString
                Case Else
                    l_Resources.AddResource UserRes.AppName, UserRes.UID, UserRes.PWD, vbNullString
                End Select
            Next UserRes
        Case "memberofgroup"
            If InStr(oNode.nodeTypedValue, "@") > 0 Then
                l_Roles.AddRole Mid$(oNode.nodeTypedValue, InStr(oNode.nodeTypedValue, "@") + 1)
            Else
                l_Roles.AddRole oNode.nodeTypedValue
            End If
        
        Case "mtf-phone"   '--format:  "{Duty 803-2212} {Home 803-2212} {Pager 803-2212}"
                            '-- or     "{Duty 803-2212} {Home 803-2212} {Pager 803-2212}"
            sData = oNode.nodeTypedValue
            sData = Replace(sData, "{", vbNullString)  '__ drop the leading '{'
            aData = Split(sData, "}")
            
            For i = 0 To UBound(aData) - 1 '__ the last one is blank
                aData(i) = Trim(aData(i))   '__ clean it up a little
                sDelimiterLoc = InStr(aData(i), " ")
                If sDelimiterLoc > 0 And aData(i) <> vbNullString Then
                     Call NodeToColl(Left$(aData(i), sDelimiterLoc - 1), Mid$(aData(i), sDelimiterLoc + 1), l_Attrib)
                End If
            Next i
            
        Case "snaredata"        '-- format: "Military ID=019378463,Email=2@chcsii.com"
            sData = oNode.nodeTypedValue
            aData = Split(sData, ",")
            
            For i = 0 To UBound(aData)
                aData(i) = Trim(aData(i))   '__ clean it up a little
                sDelimiterLoc = InStr(aData(i), "=")
                Call NodeToColl(Left$(aData(i), sDelimiterLoc - 1), Mid$(aData(i), sDelimiterLoc + 1), l_Attrib)
            Next i
        
        Case Else
           '__ just add the element ot the attrib collection
                Call NodeToColl(oNode.nodeName, oNode.nodeTypedValue, l_Attrib)
        
        End Select
        Set oNode = oNode.nextSibling
        If oNode Is Nothing Then
            Exit Do
        End If
    Loop

    
    Call GetUserPolicyInfo(l_Attrib)
    
    Call SetResourceData(colRes, LoginTarget)
    
    Dim chcs_uid As String
    Dim chcs_pwd As String
    For i = 1 To colRes.Count
        If colRes(i).AppName = "CHCS" Then  '-- at this point there should be only one
            chcs_uid = colRes(i).UID
            chcs_pwd = colRes(i).PWD
            Exit For
        End If
    Next i
    If GetResInfo Then
        Dim ResDataMTF As Collection
        Dim ResDataCHCS As Collection
        Set ResDataMTF = New Collection
        Set ResDataCHCS = New Collection
        
        If GetUserResourceInfo(ResDataMTF, ResDataCHCS) Then 'l_Resources, sCHCSUID, sCHCSPass) Then
             l_Resources.AddResource "MTFInfo", vbNullString, vbNullString, ResDataMTF
             If chcs_uid <> "" Then
                l_Resources.AddResource "CHCS", chcs_uid, chcs_pwd, ResDataCHCS
            End If
        End If
    End If
    
    '<< if we came in with an ncid, we have to add it to the attribute collection here
    If NCID <> "" Then
        Call NodeToColl("User_NCID", NCID, l_Attrib)
    End If
    
    Set objsnareuser = New User
    With objsnareuser
        .UserID = sUser_ID
'<< SCR 44711 user-based privileges (not role-based) need the full userid in this field.  move 38894 fix to the applock module.
'<< scr 38894  change username to eliminate the server name (from : sjdev@doc1  to: doc1)
'        If InStr(sLoginUser, "@") > Len(sLoginUser) Then
'            .UserName = Mid$(sLoginUser, InStr(sLoginUser, "@"))
'        Else
            .UserName = sName
'        End If
        .Password = sPassword
        Set .objSec = Me
        Set .Roles = l_Roles
        Set .Resources = l_Resources
        Set .Attributes = l_Attrib
        Set .CAC_Monitor = CAC_Monitor
        .CardID = m_CardInfo.CardName
        .LoginMode = m_LoginMode
        Set .Privileges = SW_GetUserAcl(GetSnareHandle, sUser_ID, Not GetResInfo)  '-- if we're getting the resource data, we don't need to impersonate
    End With
'<< SCR 19749 record this for review
    'Call RecordResource(sLoginUser, vbnullstring, SW_GetPass())
    objsnareuser.DumpObject
    
    If AddUserToCollection(objsnareuser) Then
        Set GetSnareUser = objsnareuser
    End If
    
    Exit Function
errhandler:
    Set GetSnareUser = Nothing
End Function

Private Function SetResourceData(colRes As Collection, ByVal Target As String) As Boolean
'<< scr 49190 handle the 'any' and the possiblity of multiple domain-specific resource entries.
'   if there are more than one res entries for an application name, then use the one with a target of 'any'
'   unless there is one with the proper domain in the target. the 'proper' domain is the one that is in the 'logonTarget' attribute
    On Error Resume Next
    Dim res2        As SW_Res
    Dim vEntry      As SW_Res
    Dim colNew      As New Collection
    Dim AnyEntry    As Long                 '-- set to the entry number
    Dim FoundEntry  As Long                 '-- set to the entry number of a duplicate
    Dim i           As Long
    
    For Each vEntry In colRes
        AnyEntry = 0
        FoundEntry = 0
        For i = 1 To colNew.Count
            If colNew(i).AppName = vEntry.AppName Then
                FoundEntry = i
                If LCase$(colNew(i).Target) = "any" Then       '- this is the one if there is nothing more specific
                    AnyEntry = i
                End If
                
            End If
        Next i
        '-- if the new entry has the target in it and there is an "Any"
        '   then drop the 'any' and go with the new target
        If InStr(vEntry.Target, Target) > 0 Then
            If FoundEntry > 0 Then
                colNew.Remove FoundEntry                '-- drop the old one in favor of the new one
            End If
            colNew.Add vEntry, vEntry.AppName       '   and add the new one
       
        
        ElseIf FoundEntry = 0 Then
            colNew.Add vEntry, vEntry.AppName
        End If
    Next vEntry
    Set colRes = colNew
    SetResourceData = True
    
End Function

Private Function GetLoginTarget(fullTarget As String) As String
    Dim TgtParts() As String
    TgtParts = Split(fullTarget, ".")
    Dim X As Long
    X = UBound(TgtParts)
    If X = 2 Then
        If LCase$(TgtParts(2)) = "com" Then
            GetLoginTarget = TgtParts(1)
        End If
        
    ElseIf UBound(TgtParts) = 0 Then
        GetLoginTarget = TgtParts(0)
    Else
        GetLoginTarget = TgtParts(X)
    End If
    
End Function

Private Function GetUserResourceInfo(l_ResDataMTF As Collection, _
                                     l_ResDataCHCS As Collection) As Boolean

'__ get the resource data for this user/mtf
    Dim XMLResponse         As String
    Dim oXMLDom             As DOMDocument40
    Dim vmsscript           As String
    Dim oNode               As IXMLDOMNode
    Dim sErr                As String                   '-- error msg from snareworks
    Dim lErr                As Long
    Dim XMLRequest          As String
    
    On Error GoTo errhandler
    
    XMLRequest = "<XMLRequest><ShowEntries><Entry>" _
               & "<Organization><Name>" & m_sDomain & "</Name></Organization>" _
               & "</Entry></ShowEntries></XMLRequest>"
       
    If Not SnarexML(SnareHandle, XMLRequest, oXMLDom, lErr, sErr) Then
        Exit Function
    End If
    
    Set oNode = oXMLDom.selectSingleNode("//Organization/*")
    
    vmsscript = vbNullString
    Do While True
        Select Case LCase$(oNode.nodeName)
'__ these are MTF Resource elements
        Case "mtf-ad_hoc_server":           Call NodeToColl("Ad_Hoc_Server", oNode.nodeTypedValue, l_ResDataMTF)
        Case "mtf-chcsii_appsrvr_ipaddr":   Call NodeToColl("CHCSII_AppSrvr_IPAddr", oNode.nodeTypedValue, l_ResDataMTF)
        Case "mtf-chcsii_ipaddr":           Call NodeToColl("CHCSII_IPAddr", oNode.nodeTypedValue, l_ResDataMTF)
        Case "mtf-chcsii_port":             Call NodeToColl("CHCSII_Port", oNode.nodeTypedValue, l_ResDataMTF)
        Case "mtf-chcsii_ie_ipaddr":        Call NodeToColl("CHCSII_IE_IPAddr", oNode.nodeTypedValue, l_ResDataMTF)
        Case "mtf-adtcontextncid":          Call NodeToColl("ADTContextNCID", oNode.nodeTypedValue, l_ResDataMTF)
        Case "mtf-chcsii_ie_dl_port":       Call NodeToColl("CHCSII_IE_DL_Port", oNode.nodeTypedValue, l_ResDataMTF)
        Case "mtf-chcsii_appsrvr_name":     Call NodeToColl("CHCSII_AppSrvr_Name", oNode.nodeTypedValue, l_ResDataMTF)
        Case "mtf-chcsii_ie_sadr_dir":      Call NodeToColl("CHCSII_IE_SADR_Dir", oNode.nodeTypedValue, l_ResDataMTF)
        Case "mtf-chcsii_ie_password":      Call NodeToColl("CHCSII_IE_Password", oNode.nodeTypedValue, l_ResDataMTF)
        Case "mtf-chcsii_ie_user":          Call NodeToColl("CHCSII_IE_User", oNode.nodeTypedValue, l_ResDataMTF)
        Case "mtf-chcsii_enterprise_ncid":  Call NodeToColl("CHCSII_Enterprise_NCID", oNode.nodeTypedValue, l_ResDataMTF)
        Case "mtf-chcsii_facility_ncid":    Call NodeToColl("CHCSII_Facility_NCID", oNode.nodeTypedValue, l_ResDataMTF)
        Case "mtf-cita_data_drive":         Call NodeToColl("CITA_DATA_DRIVE", oNode.nodeTypedValue, l_ResDataMTF)

'__ these are CHCS Resource elements
        Case "mtf-chcs_ipaddr":             Call NodeToColl("CHCS_IPAddr", oNode.nodeTypedValue, l_ResDataCHCS)
        Case "mtf-oe_ipaddr":               Call NodeToColl("OE_IPAddr", oNode.nodeTypedValue, l_ResDataCHCS)
        Case "mtf-oe_port":                 Call NodeToColl("OE_Port", oNode.nodeTypedValue, l_ResDataCHCS)
        Case "mtf-dmis_id":                 Call NodeToColl("DMIS_ID", oNode.nodeTypedValue, l_ResDataCHCS)
        Case "mtf-vmsuser":                 Call NodeToColl("VMSUserName", oNode.nodeTypedValue, l_ResDataCHCS)
        Case "mtf-vmspassword":             Call NodeToColl("VMSPassword", oNode.nodeTypedValue, l_ResDataCHCS)
'<< SCR xxx pipe delimit the vmsscript
        Case "mtf-vmsscript":               vmsscript = vmsscript & oNode.nodeTypedValue & "|"
        
'__add all other entries in the mtf resourse
        Case Else
            If InStr(oNode.nodeName, "-") > 0 Then
                Call NodeToColl(Mid$(oNode.nodeName, InStr(oNode.nodeName, "-") + 1), oNode.nodeTypedValue, l_ResDataMTF)
            Else
                Call NodeToColl(oNode.nodeName, oNode.nodeTypedValue, l_ResDataMTF)
            End If
        End Select
        Set oNode = oNode.nextSibling
        If oNode Is Nothing Then
            Exit Do
        End If
    Loop
    
'<< SCR 45054 drop the trailing pipe
    If Right$(vmsscript, 1) = "|" Then
        vmsscript = Left$(vmsscript, Len(vmsscript) - 1)
    End If
    Call NodeToColl("VMSScript", vmsscript, l_ResDataCHCS)

    
    GetUserResourceInfo = True
    Exit Function
    
errhandler:
    
    GetUserResourceInfo = False
End Function

Private Function GetUserPolicyInfo(ByRef l_Attrib As Collection) As Boolean
'__ get the resource data for this user/mtf
    Dim XMLResponse         As String
    Dim oXMLDom             As DOMDocument40
    Dim oNode               As IXMLDOMNode
    Dim sErr                As String                   '-- error msg from snareworks
    Dim lErr                As Long
    Dim XMLRequest          As String
    
    On Error GoTo errhandler
    
    XMLRequest = "<XMLRequest><ShowEntries><Entry>" _
               & "<Policy><Name>" & m_sDomain & "</Name></Policy>" _
               & "</Entry></ShowEntries></XMLRequest>"
    
    If Not SnarexML(SnareHandle, XMLRequest, oXMLDom, lErr, sErr) Then
        err.Raise lErr, "GetUserPolicyInfo", sErr
    End If
        
    Set oNode = oXMLDom.selectSingleNode("//Policy/*")
    
    Do While Not oNode Is Nothing
        Select Case LCase$(oNode.nodeName)
        Case "name"
            Call NodeToColl("PolicyName", oNode.nodeTypedValue, l_Attrib)
        Case "acctidleactivity"
            Call NodeToColl("Timer", oNode.nodeTypedValue, l_Attrib)
        Case Else
            Select Case Left(LCase$(oNode.nodeName), 6)
            Case "passwd"
                With objPassword
                    Select Case LCase$(oNode.nodeName)
                    Case "passwdlifespan":                  .LifeSpan = oNode.nodeTypedValue
                    Case "passwdexpdate":                   .ExpDate = oNode.nodeTypedValue
                    Case "passwdminlength":                 .MinLength = oNode.nodeTypedValue
                    Case "passwdmaxlength":                 .MaxLength = oNode.nodeTypedValue
                    Case "passwdminnumeric":                .MinNumeric = oNode.nodeTypedValue
                    Case "passwdmixedcase":                 .MixedCase = oNode.nodeTypedValue
                    Case "passwdpercentchanged":            .PercentChanged = oNode.nodeTypedValue
                    Case "passwdpercentduplicate":          .PercentDuplicate = oNode.nodeTypedValue
                    Case "passwddictionarycheck":           .DictionaryCheck = oNode.nodeTypedValue
                    Case "passwdnumbernonalphanumeric":     .NumberNonAlphanumeric = oNode.nodeTypedValue
                    Case "passwdallowchanges":              .AllowChanges = oNode.nodeTypedValue
                    Case "passwdkeephistory":               .KeepHistory = oNode.nodeTypedValue
                    Case "passwdremind":                    .Remind = oNode.nodeTypedValue
                    End Select
                End With
            Case Else
                Call NodeToColl(oNode.nodeName, oNode.nodeTypedValue, l_Attrib)
            End Select
        End Select
        Set oNode = oNode.nextSibling
        
    Loop
    
  
    GetUserPolicyInfo = True
    Exit Function
errhandler:
    GetUserPolicyInfo = False
End Function

Private Function ICHCSII_Security_GetUser(Optional sUserId As String) As CHCSII_ClientSecurity.ICHCSII_User
    Dim l_User          As User
    Dim ErrCode         As Long
    Dim ErrMsg          As String
    Dim sNCID           As String   '-- temp holder for the ncid
    
    On Error GoTo errhandler
    
    If Not (m_bActivated) Then
        err.Raise SEC_NOT_ACTIVATED, CLASS_NAME, GetErrMsg(SEC_NOT_ACTIVATED)
    End If
    '-- LogAccess "getuser  " & sUserId
'__ we send a userid when we just want a user for security reasons, so if the id is blank, this is a login situation
    If sUserId = vbNullString Then
        Set ICHCSII_Security_GetUser = ICHCSII_Security_GetActiveUser()
        Exit Function
    End If

'__ SHAW 15 JUL 03 - new vacman (snareworks). They may send in an ncid, so convert that to login name.
'   we now allow the user to change both his pwd and his uid.  so the snareworks uid may not match what is in the hems users table.
'   snareworks doesn't update the 3m user id (hems_users.logon_name) so they may not match.
'   if we go to 3M.Member object and get logon_name, it may not match the snaareworks name.
'   so the applications should now send in an ncid and we'll handle it.
'   snareworks gets the ncid from 3m during account creation.  It doesn't change
    If IsNumeric(sUserId) Then
        If GetUserByNCID(l_User, sUserId) Then      '-- see if we have him
            Set ICHCSII_Security_GetUser = l_User
            Exit Function
        Else
            sNCID = sUserId                         '-- used this later, and as a state-flag
            sUserId = NCID_To_SnareUserID(sUserId)  '-- change to snare user id if possible
        End If
    End If
    
'__ see if we have him by user id
    '-- LogAccess "getuser looking for user " & sUserId
    If Not GetUserFromCollection(l_User, sUserId) Then
        '-- LogAccess "getuser not found calling getsnareuser"
        Set l_User = GetSnareUser(sUserId, False, ErrCode, ErrMsg, sNCID)
        
        If ErrCode <> 0 Then
            err.Raise ErrCode, "Get Snare User (XML)", ErrMsg
        End If
    Else
        '-- LogAccess "   found it in collection: " & sUserId
        If sNCID <> "" Then     '-- we came in with an ncid and didn't find him using that,
                                '   but we DID find him using the userid we got from vacman.
                                '   So, he's in the collection without an ncid.  Fix that now.
            
            Call l_User.AddAttribute("User_NCID", sNCID)
        End If
    End If

    Set ICHCSII_Security_GetUser = l_User
    Exit Function
    
errhandler:
    err.Raise err.Number, err.Source, err.Description
    
    
End Function

Private Function ICHCSII_Security_TMIPAvailable() As Boolean
    '__ for theater only.  outside of theater this will always be false
    ICHCSII_Security_TMIPAvailable = False
End Function


Private Function NodeToColl(ByVal skey As String, ByVal sValue As String, Col As Collection) As Boolean
    Dim objPair As KeyValuePair
    On Error Resume Next                '-- this key may have come in before, if so, just drive on.
    If sValue = vbNullString Then Exit Function
            
    Set objPair = New KeyValuePair
    objPair.Key = skey
    objPair.Value = sValue
    Col.Add objPair, objPair.Key
    
    NodeToColl = True
End Function

Private Function ICHCSII_Security_ValidateUser(ByVal UserID As String, ByVal Password As String, ErrNum As Long) As Boolean
    On Error GoTo errhandler
    Dim nErr As Long
    Dim sErr As String
    Dim lErr As Long
    
    If Not SW_Logon(GetSnareHandle, UserID, Password, nErr, lErr) Then
        If lErr <> 0 Then
            ErrNum = lErr
        Else
            ErrNum = nErr
        End If
        ICHCSII_Security_ValidateUser = False
    Else
        ErrNum = 0
        ICHCSII_Security_ValidateUser = True
    End If
    Exit Function
errhandler:
    ICHCSII_Security_ValidateUser = False

End Function

Private Function ITMIP_AppData_GetAppData(ByVal AppName As String, _
                                                UID As String, _
                                                PWD As String, _
                                                SID As String) As Boolean
'<< SCR 31267 add registry-based uid and passwords and
'   use C2T_Security to passthrough to tmip_security if it's there

    Dim oCrypto  As New CCrypto
    Dim ePwd     As String
    Dim eUid     As String
    Dim FRMWK_PWD As String
    Dim sText    As String
    Dim oShared  As New MMMHISShared.CWShared
    Dim Framework_PWD As String
    Const sExp As String = "trtkabasnbi"
    
        
    UID = oShared.GetSupportValue("APP_UID")
    If UID <> vbNullString Then UID = oCrypto.Decrypt(UID, oCrypto.GetHash("CHCSII-T" & sExp, True), True, frezBlockEncryption)
    PWD = oShared.GetSupportValue("APP_pwd")
    If PWD <> vbNullString Then PWD = oCrypto.Decrypt(PWD, oCrypto.GetHash("CHCSII-T" & sExp, True), True, frezBlockEncryption)
    FRMWK_PWD = oShared.GetSupportValue("FRMWK_PWD")
    If FRMWK_PWD <> vbNullString Then FRMWK_PWD = oCrypto.Decrypt(FRMWK_PWD, oCrypto.GetHash("CHCSII-T" & sExp, True), True, frezBlockEncryption)
                
    SID = vbNullString
    ITMIP_AppData_GetAppData = True
    Exit Function

errhandler:
    MsgBxARMd "Unexpected error in Security Module.GetAppData: " & vbCrLf & err.Description, vbCritical, "Security Error"

End Function

Private Function Logoff() As Boolean
    Dim hSnare As Long
    hSnare = GetSnareHandle
    If hSnare = 0 Then
        Logoff = False
        Exit Function
    End If

    logoutSnareAPI hSnare
    closesessionSnareAPI hSnare
    SnareHandle = 0
    Logoff = True
End Function

Private Function ChangePass(sUser As String, _
                              sOldPass As String, _
                              sNewPass As String, _
                              sErr As String) As Boolean
    Dim lErr   As Long
    Dim lSecErr As Long
    Dim hSnare As Long
    Dim sErrMsg As String
    
    hSnare = GetSnareHandle
    
    If hSnare = 0 Then
        ChangePass = False
        Exit Function
    End If
    
    If sUser = vbNullString Then
        ChangePass = False
        Exit Function
    End If
    lErr = 0
    If GetUser(lErr) = sUser Then
    
        lErr = changepassSnareAPI(hSnare, sNewPass)
        If lErr <> SN_SUCCESS Then
            If lErr = SN_FAILURE Then
                sErrMsg = Space(2000)
                lErr = errorSnareAPI(SnareHandle, lSecErr, sErrMsg)
                sErr = GetErrMsg(lSecErr)
            Else
                sErr = GetErrMsg(lErr)
            End If
'<< scr 45194 added this to succeed in the call
        Else '__ this is the normal 'success' route
            ChangePass = True
            Exit Function
        End If
        
    ElseIf lErr = SN_SUCCESS Then
        If loginSnareAPI(hSnare, sUser, sOldPass, sNewPass, 0, 0, 0) = SN_SUCCESS Then
             ChangePass = True
        Else
            sErr = "Unable to logon to security server"
             ChangePass = False
        End If
    Else   '__ error on getuser call
        sErr = GetErrMsg(lErr)
    End If
    
    If sErr <> vbNullString Then
        err.Raise lErr, "Change Password", sErr
    End If
End Function
 
Private Function ChangeUserID(ByVal OldUser As String, _
                                ByVal NewUser As String, _
                                ByRef sErr As String) As Boolean
    Dim lErr   As Long
    sErr = vbNullString
    Dim hSnare As Long
    
    hSnare = GetSnareHandle
    
    If hSnare = 0 Then
        ChangeUserID = False
        Exit Function
    End If
'__ be sure we are the current user
    If GetUser(lErr) = OldUser Then
        lErr = renameaccountSnareAPI(hSnare, NewUser)
        sErr = GetErrMsg(lErr)
        If lErr = SN_SUCCESS Then
             sErr = vbNullString
             ChangeUserID = True
        End If
    ElseIf lErr = 0 Then
        ChangeUserID = False
    Else   '__ error on getuser call
        ChangeUserID = False
        sErr = GetErrMsg(lErr)
        
    End If
'ErrHandler:
    If sErr <> vbNullString Then
        err.Raise lErr, "Change Password", sErr
    End If
    

'int errorSnareAPI ( snarehandle h, int *error, char *reason );
'
'h   [in]  the snarehandle obtained from opensessionSnareAPI
'error   [out]  a pointer to an integer to hold the secondary error code
'reason  [out]  a pointer to a caller-provided buffer to hold an error description or a null string.
'
'Return Values:
'Returns the primary error code for last SnareAPI call or
'SN_PARAMFAIL    The snarehandle is NULL or is invalid (opensessionSnareAPI was not called to get the snarehandle or the session has been closed with closesessionSnareAPI)
'___________________

End Function

Private Function GetUser(Optional ByRef nErrCode As Long) As String
'-- returns the userid of the currently logged in user
    Dim sUser   As String * 64
    Dim sPass   As String * 64
    Dim sErrMsg As String * 256
    Dim nLen    As Long
    Dim hSnare As Long
    
    hSnare = GetSnareHandle
    If hSnare = 0 Then
        GetUser = vbNullString
        Exit Function
    End If
    

    If getuserSnareAPI(hSnare, sUser) = SN_SUCCESS Then
        ' Start the timer
        '--- SCR-ARM-16149 MSolano 9/14/01
        '--- FIRST TIME WE HAVE USER
        '--- THIS MUST BE ONLY CALL TO ARM_InitAllTransactionsForSnareworks
        Call ARM_InitAllTransactionsForSnareworks(sUser)
        Call ARM_Start_SW_GetUserTx(gsARMTx_SecurityLogonOk)  '--- SCR-ARM-16149 MSolano 9/14/01
                       
        GetUser = ToVBString(sUser)
        If GetUser <> vbNullString And fnOldFilter = 0 Then
           ' install the exception filter to clean up the cred
           fnOldFilter = SetUnhandledExceptionFilter(AddressOf ExceptFilter)
        End If

    Else
        sErrMsg = Space(2000)
        nErrCode = errorSnareAPI(hSnare, nErrCode, sErrMsg)
        GetUser = vbNullString
    End If
End Function

'-- added for clinic client only.  this and sw_getrolesex should be dropped
Public Function GetUserRoles(ByVal sLogonName As String) As String
    GetUserRoles = SW_GetRoleEx(GetSnareHandle, sLogonName)
End Function


Private Function FirstTimeuser() As ICHCSII_User
'-- we come here if the user clicked on the 'first time user' line on the logon screen
    On Error GoTo errhandler
    Dim AccessCode   As String
    Dim VerifyCode      As String
    Dim bSuccess        As Boolean
    Dim frmInit         As frmInitLogon
    Dim iNumTries       As Integer
    Dim nErrCode        As Long
    Dim nSecErrCode     As Long
    
    
    bSuccess = False
    Do While bSuccess = False
        If iNumTries <= 3 Then
            iNumTries = iNumTries + 1
            Set frmInit = New frmInitLogon
            frmInit.Show vbModal
            If frmInit.UserCancelled Then
                Exit Function
            End If
            AccessCode = frmInit.txtActivate
            VerifyCode = frmInit.txtVerify
            If VerifyCode = vbNullString Or AccessCode = vbNullString Then
                Exit Function
            End If

        '-- forgot to prefix with a carrot
            If SW_Logon(SnareHandle, "^" & AccessCode, VerifyCode, nErrCode, nSecErrCode) Then
                Unload frmInit
                If Not SetNewUserInfo Then
                    Call SW_Logoff(SnareHandle)
                    Set FirstTimeuser = Nothing
                    Exit Function
                Else
                    Set FirstTimeuser = ICHCSII_Security_GetUser
                    bSuccess = True
                End If
             Else
                Unload frmInit
                If nErrCode <> 0 Then
                    If nErrCode = 1000 And nSecErrCode <> 0 Then
                        frmInit.ShowError "Message from Security Server:" & vbCrLf & vbCrLf & GetErrMsg(nSecErrCode)
                    Else
                        frmInit.ShowError "Message from Security Server:" & vbCrLf & vbCrLf & GetErrMsg(nErrCode)
                    End If
                Else
                    frmInit.ShowError "User Id or Password not found."
                End If
            End If
        Else
            frmInit.ShowError "User Id or password not found. Maximum number of attempts has been made."
            Unload frmInit
            Set FirstTimeuser = Nothing
            bSuccess = True
            err.Raise SEC_MAX_LOGIN_ATTEMPTS_MET, CLASS_NAME, GetErrMsg(SEC_MAX_LOGIN_ATTEMPTS_MET)
        End If
    Loop
    
    Exit Function
errhandler:

End Function

Private Function SetNewUserInfo() As Boolean

'-- we have to get a new user id and new password from the user for snareworks.
'       we are currently logged in as a provisional user, now we have to setup a new chcsii/snareworks user
    'Set FirstTimeuser = GetSnareUser(frmLogin.GetUserId)
    Dim oUser As User
    Dim SnareHandle As Long
    Dim ErrNum As Long
    Dim sErr As String  '- the error that comes back
    Dim UID As String
    Dim PWD As String
    Dim OldPWD As String

    Dim frm As frmPassword
    Set frm = New frmPassword
    
'<< scr 50586 need default rules here.
    Dim oPWDRules As New PassWordRules
    oPWDRules.SetDefaultProperties
    Call frm.SetPasswordRules(oPWDRules)
    
    frm.FrmMode = "NEW"
    frm.Move (Screen.Width - frm.Width) / 2, (Screen.Height - frm.Height) / 2
    frm.Show vbModal
    
    If frm.bUserCanceled Then
        SetNewUserInfo = False
        Exit Function
    End If
    
    UID = frm.txtUID
    PWD = frm.txtNewPWD1
    
    Unload frm
    
    SnareHandle = GetSnareHandle(ErrNum)
    If SnareHandle = 0 Then
        SetNewUserInfo = False
        Exit Function
    End If
    
    Set oUser = ICHCSII_Security_GetUser
    If oUser Is Nothing Then
        SetNewUserInfo = False
        MsgBxARMd "Unable to validate the current user.  Please contact your system administrator." & vbCrLf _
            & "Account not created.", vbCritical, "Create Account"
        Exit Function
    End If
    '-we now have at least a suggestion for a userID and password
    'we have to do a change pwd and change uid (in that order!)
    'if the first succeeds, but the second fails, we have to tell the user what we did:
    
    
    With oUser
        If ChangePass(.UserID, .Password, PWD, sErr) Then
            If ChangeUserID(.UserID, UID, sErr) Then
               
            Else
                MsgBxARMd "Unable to set the user Id supplied." & vbCrLf _
                & sErr & vbCrLf & "User Id was not changed.", vbCritical, "Change User ID"
            End If
           
        Else
            MsgBxARMd "Unable to set the password supplied." & vbCrLf _
                & sErr & vbCrLf & "Account not created.", vbCritical, "Create Account"
        End If
        
    End With
    
    SetNewUserInfo = True
    Exit Function
errhandler:
    SetNewUserInfo = False

End Function
Private Function GetUserFromCollection(objsnareuser As ICHCSII_User, ByVal UserID As String) As Boolean
'ensure we are searching using the full name (domain@userID)
    If Not CBool(InStr(UserID, "@")) Then
        UserID = m_sDomain & "@" & UserID
    End If

'-- this is here just to sequester the 'on error resume next' thing
    On Error Resume Next
    Set objsnareuser = m_colUsers.Item(UserID)
    GetUserFromCollection = (err.Number = 0)

End Function

Private Function AddUserToCollection(objsnareuser As ICHCSII_User) As Boolean
'-- this is here just to sequester the 'on error resume next' thing
    On Error Resume Next
    m_colUsers.Add objsnareuser, objsnareuser.UserID    '-- this had been 'sname' but we come in with fully qualified name looking for match
    AddUserToCollection = (err.Number = 0)
End Function

Private Function GetUserByNCID(objsnareuser As ICHCSII_User, ByVal user_Ncid As String) As Boolean
    Dim oUser As ICHCSII_User
    Dim ncid_attrib As String
    
'-- some in the collection don't have the ncid attribute
    On Error GoTo errhandler
    For Each oUser In m_colUsers
        ncid_attrib = oUser.getAttribute("USER_Ncid")    '-- this may error out
        
        If ncid_attrib = user_Ncid Then
            Set objsnareuser = oUser
            GetUserByNCID = True
            Exit Function
        End If
    Next oUser
    GetUserByNCID = False
    Exit Function
    
errhandler:
    GetUserByNCID = False
    Set objsnareuser = Nothing
End Function

Private Function NCID_To_SnareUserID(ByVal NCID As String) As String
    
    '__ see if we have it already
    Dim oUser As ICHCSII_User
    Dim attribs As Collection
    Dim i As Long
    Dim hSnare As Long
    Dim NameLen As Long
    Dim sTmp As String
    
On Error GoTo errhandler
    
    hSnare = GetSnareHandle
    Dim sOutput As String * 512
    NameLen = 512
    If hSnare <> 0 Then
'-- SCR 49216 converting the "commandSnareAPI(hSnare, "Ncid2Logon " & sUserId, 0, sOutput, NameLen)" to an xml call
        If GetSnareUserID(GetSnareHandle, NCID, m_sDomain, sTmp) Then
            NCID_To_SnareUserID = sTmp
        Else
            NCID_To_SnareUserID = NCID
        End If
    End If
    Exit Function
errhandler:
    NCID_To_SnareUserID = NCID
End Function



'Public Function ValidateSignature(ByVal Data As String, _
'                                  ByVal Signature As String, _
'                                  ByVal SigFormate As SigFormate) As Boolean
'
'    Shell "openssl rsautl -in c:\sig.bin -verify -raw -inkey c:\temp.pem -certin -out c:\sigout.txt"
'
'    Str = oCrypto.GetHash("This is the text that I want signed", True)
'    Str = Str & Str
'    Str = Str & Str
'    Call CAC_Monitor.SignText(PIN, Str, Sig)
'
'End Function
