VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPasswordOps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "C2T_Security.clsPasswordOps"

Private moSecSQL As SqlOps 'ICHCSII_SQL.ISqlOps
Private mbTmipAvail As Boolean
Private msUserNCID As String
Private msUserID As String
Private mlUserTimer As Long
Private mbReauthenticateUser As Boolean

Public Event LoginSuccessful(ByVal UserName As String, ByVal Password As String, ByVal SecSQL As SqlOps, ByRef User As ICHCSII_User)

Public Sub Logout()
    
'SCR 83806 2/8/06 Log the event: Logout
    'There is error handling in TMIPAudit and above this sub
    TMIPAudit "SECURITY", "CHCS II-T SECURITY", msUserID, "", "Successful logout", "", moSecSQL.moCnn
    
End Sub


Public Property Get UserNCID() As String
    UserNCID = msUserNCID
End Property
Public Property Let UserNCID(Val As String)
    msUserNCID = Val
End Property

Public Property Get UserTimer() As Long
    UserTimer = mlUserTimer
End Property
Public Property Let UserTimer(Val As Long)
    mlUserTimer = Val
End Property

Public Property Get ReauthenticateUser() As Boolean
    ReauthenticateUser = mbReauthenticateUser
End Property
Public Property Let ReauthenticateUser(Val As Boolean)
    mbReauthenticateUser = Val
End Property

Public Property Get UserId() As String
    UserId = msUserID
End Property
Public Property Let UserId(Val As String)
    msUserID = Val
End Property


Public Sub Init(ByVal TmipAvail As Boolean)
    
    mbTmipAvail = TmipAvail
    
End Sub

Public Function PromptUserToLogin(Optional FromTheaterSecurity As Boolean = False) As ICHCSII_User

On Error GoTo ErrHandler

Dim bSuccess  As Boolean        '-- has the user successfully logged in
Dim iNumTries As Integer        '-- the number of times the user has tried to log in
Dim oUser    As ICHCSII_User   '-- the user trying to log on
Dim sUser     As String         '-- the user id
Dim ors       As ADODB.Recordset
Dim smp       As Integer        '-- temp screen mousepointer holder
Dim frmLogin  As frmCHCSIILogin
Dim Theater_Resource_ID As Long
Dim oCrypto As CCrypto
Dim oComm As MMMHISComm.Comm
Dim dPwdChangeDate As Date
Dim Pwd1 As String
Dim Pwd2 As String
Dim Pwd3 As String
Dim pwd4 As String
Dim ocwShared As CWShared
Dim previousUser As String
Dim invalidPassword As Boolean

    Set oComm = New MMMHISComm.Comm
    Set oCrypto = New CCrypto
    
    'hold the current pointer till we're done
    smp = Screen.MousePointer
    Screen.MousePointer = vbDefault
    
    iNumTries = 1
    
    Do Until bSuccess Or iNumTries > 3
        'SCR 67597
        If frmLogin Is Nothing Then
            Set frmLogin = New frmCHCSIILogin
        End If
        
        frmLogin.ReauthenticateUser = ReauthenticateUser
        frmLogin.UserId = UserId
        
        frmLogin.Show vbModal
        
        If frmLogin.HasCancelled() Then
            Unload frmLogin
            Screen.MousePointer = smp
            Exit Function   '--- SCR 74016  Sherry Wang   7/20/2005
        End If
        
        'SF SCR 89193 3-6-06
        Set ocwShared = New CWShared
        ocwShared.CurrentUserName = frmLogin.txtUserId
        Set ocwShared = Nothing
        
        '__ this is the new c2t security table call:
        If GetUserAccess(frmLogin.txtUserId, frmLogin.txtPassword, mbTmipAvail) Then
            
            StartDas
            
            '--- SCR 79741   Sherry Wang   10/14/2005
            If gbDBLoginInfoChanged = True Then Exit Function
            
            ' SCR 100543, JQ, 7-21-06; make user id compare case-insensitive(added UPPER and UCase)
            moSecSQL.Execute "select ur.user_ncid, ur.user_password password,ur.locked, ur.user_name name," _
                      & " ur.res_id, UR.IS_DEFAULT_PWD, UR.PWD_CHANGE_DATE, u.*, " _
                      & " ur.old_user_pwd1, ur.old_user_pwd2, ur.old_user_pwd3, ur.old_user_pwd4," _
                      & " ur.old_user_pwd5, ur.old_user_pwd6, ur.old_user_pwd7, ur.old_user_pwd8," _
                      & " ur.old_user_pwd9, ur.old_user_pwd10," _
                      & " ur.last_attempt, ur.attempts" _
                      & " FROM users u,user_resources ur, resources r" _
                      & " where UPPER(ur.user_name) = " & SQLQuote(UCase(frmLogin.txtUserId.Text)) _
                      & " AND u.user_ncid = ur.user_ncid " _
                      & " and r.res_id = ur.res_id" _
                      & " AND r.resource_name = '" & THEATER_KEYWORD & "'"
            
            If Not moSecSQL.EOF Then
                Theater_Resource_ID = Val(vbNullString & moSecSQL("res_id"))
                If moSecSQL("PWD_CHANGE_DATE") = Empty Then
                    dPwdChangeDate = Empty
                Else
                    dPwdChangeDate = moSecSQL("PWD_CHANGE_DATE")
                End If
                If moSecSQL("locked") <> 0 Then
                    frmLogin.ShowError "User id is locked. See your system administrator.", 0
                    Set PromptUserToLogin = Nothing
                    'bSuccess = True
                    Screen.MousePointer = smp
                    'SCR 83806
                    TMIPAudit "SECURITY", "CHCS II-T SECURITY", frmLogin.txtUserId.Text, "", "Unsuccessful login. User id is locked.", "", moSecSQL.moCnn
                    Unload frmLogin
                    Exit Function
                ' JQ Note: password expiration changed from 90 to 60 days below, Feb 2010:
                ElseIf CBool(Val(moSecSQL("IS_DEFAULT_PWD") & "")) = True _
                    Or DateAdd("d", 60, oComm.GmtToLocal(dPwdChangeDate)) < Now Then
                    'The user must change their password
                    frmLogin.bChangePwd = True
                    
                End If
                If moSecSQL("Password") = oCrypto.GetHash(frmLogin.txtPassword.Text) Then
                    UserTimer = CLng(moSecSQL("Timer"))
                    UserId = frmLogin.txtUserId
                    
                '-- now the real security check
                    If FromTheaterSecurity Then
                        'Theater security does not support a user object.
                        msUserNCID = moSecSQL("user_ncid") & ""
                    Else
                        'Get the user object from the calling object.
                        RaiseEvent LoginSuccessful(moSecSQL("Name"), moSecSQL("Password"), moSecSQL, oUser)
                    End If
                    
                    'Set oUser = GetUserEx(moSecSQL("Name"), moSecSQL("Password"))
                    If oUser Is Nothing And FromTheaterSecurity = False Then
                        '--- SCR 79741   Sherry Wang  10/12/2005
                        If gbDBLoginInfoChanged = True Then
                            Exit Function
                        Else
                            frmLogin.ShowError "Unable to validate logon.", 0
                            'SCR 83806
                            TMIPAudit "SECURITY", "CHCS II-T SECURITY", frmLogin.txtUserId.Text, "", "Unable to validate logon.", "", moSecSQL.moCnn
                        End If
                    Else
                        Set PromptUserToLogin = oUser
                        bSuccess = True
                        'SCR 83806
                        TMIPAudit "SECURITY", "CHCS II-T SECURITY", frmLogin.txtUserId.Text, "", "Successful login.", "", moSecSQL.moCnn
                    End If
                Else
                    frmLogin.ShowError "Invalid password.", 2
                    'SCR 83806
                    TMIPAudit "SECURITY", "CHCS II-T SECURITY", frmLogin.txtUserId.Text, "", "Unsuccessful login. Invalid password.", "", moSecSQL.moCnn
                    invalidPassword = True
                End If
            Else
                frmLogin.ShowError "Invalid username.", 0
                'SCR 83806
                TMIPAudit "SECURITY", "CHCS II-T SECURITY", frmLogin.txtUserId.Text, "", "Unsuccessful login. Invalid username.", "", moSecSQL.moCnn
            End If  '- If Not moSecSQL.EOF Then
        Else
            'The App_UID and App_PWD are not correct in the registry.
            frmLogin.ShowError "The application could not log on to verify credentials.", 0
            'SCR 83806
            TMIPAudit "SECURITY", "CHCS II-T SECURITY", frmLogin.txtUserId.Text, "", "CHCS II-T could not login to the database.", "", moSecSQL.moCnn
            Exit Do
        End If
        
        'Make sure the counter gets reset every time a different user is used within a single session
        'User authentication is user based and not session based
        If previousUser <> UCase(frmLogin.txtUserId.Text) Then
            previousUser = UCase(frmLogin.txtUserId.Text)
            iNumTries = 1
        End If
        
        If bSuccess = False And invalidPassword Then
            'Only want this to run with invalid passwords
            invalidPassword = False
            
            Dim lastAttempt As Date
            'Attempts gets reset to zero when account gets locked or user successfully logs on...
            If moSecSQL("ATTEMPTS") > 0 Then
                lastAttempt = moSecSQL("LAST_ATTEMPT")
                ' Calculate the difference...
                Dim iHours As Integer, iMins As Integer
                iMins = DateDiff("n", lastAttempt, oComm.LocalToGmt(Now))

                'Only increment iNumTries if attempt made within the hour...still same session
                'Otherwise leave it as is to treat as a new session when time span is > 60 minutes
                If iMins <= 60 Then
                    iNumTries = moSecSQL("ATTEMPTS") + 1
                End If
            End If
            
            Select Case iNumTries
            Case 1, 2
                'Save the current time and number of attempts to calculate 1 hour requirement between attempts
                moSecSQL.Execute "update user_resources set Last_Attempt = " & SQLDate(oComm.LocalToGmt(Now)) & "," _
                    & " Attempts = " & iNumTries _
                    & " where UPPER(user_name) = " & SQLQuote(UCase(frmLogin.txtUserId.Text)) _
                    & " and res_ID = (select res_id from resources where resource_name = '" & THEATER_KEYWORD & "')"
            Case 3
                '- lock the user out by setting the "locked"
                moSecSQL.Execute "update user_resources set locked = 1, Attempts = 0 " _
                    & " where UPPER(user_name) = " & SQLQuote(UCase(frmLogin.txtUserId.Text)) _
                    & " and res_ID = (select res_id from resources where resource_name = '" & THEATER_KEYWORD & "')"
                
                frmLogin.ShowError "Maximum allowable logon attempts exceeded. User id is now locked. See your system administrator.", 0
                'SCR 83806
                TMIPAudit "SECURITY", "CHCS II-T SECURITY", frmLogin.txtUserId.Text, "", "Unsuccessful login. Maximum allowable logon attempts exceeded. User id is now locked.", "", moSecSQL.moCnn
                Exit Do
            End Select
        ElseIf bSuccess = True Then
                'Reset attempts to zero, since user successfully logged on...
                moSecSQL.Execute "update user_resources set Attempts = 0 " _
                    & " where UPPER(user_name) = " & SQLQuote(UCase(frmLogin.txtUserId.Text)) _
                    & " and res_ID = (select res_id from resources where resource_name = '" & THEATER_KEYWORD & "')"
        End If
        
        iNumTries = iNumTries + 1
        
    Loop
    
    Dim NewPWD As String
    If frmLogin.bChangePwd And bSuccess Then
        If ChangeC2TPassword(frmLogin.txtUserId.Text, oUser) = False Then
            'Don't log the user in, they must change their password.
            Set PromptUserToLogin = Nothing
            'Don't allow Theater security to continue if they didn't change their password.
            msUserNCID = ""
        Else
            'SCR 83806
            TMIPAudit "SECURITY", "CHCS II-T SECURITY", frmLogin.txtUserId.Text, "", "The User's password has been changed.", "", moSecSQL.moCnn
        End If
    End If
    Unload frmLogin
    
    Screen.MousePointer = smp
    Exit Function
ErrHandler:
'<< scr 33048
    If Not Err.Description = "User cancelled login." Then
        MsgBox "Unexpected error in Security Module.GetUserFromLogin: " & vbCrLf & Err.Description
    Else
        Err.Raise SEC_USER_CANCELLED_LOGIN, CLASS_NAME, GetErrMsg(SEC_USER_CANCELLED_LOGIN)
    End If
    
    Exit Function
    Resume

End Function

Private Function GetUserAccess(ByVal suid As String, _
                              ByVal spwd As String, _
                     Optional ByVal TmipAvail As Boolean) As Boolean
                      
    Dim oShared         As New CWShared
    
    Dim sxUID           As String           '-- temp placeholder for whatever is in the shared.locallogon area
    Dim sxPWD           As String           '-- temp placeholder for whatever is in the shared.locallogon area
    Dim sAccUID         As String           '-- the registry-based c2t access uid/pwd
    Dim sAccPWD         As String           '-- the registry-based c2t access uid/pwd
    Dim NewPWDIdx       As Long             '-- this is the sequence number of the new/next pwd set in c2t_y
    Dim Framework_PWD   As String           '-- on a framework box, use the tmip-supplied password
    Dim oC2T            As SqlOps           '-- connection to the c2t schema
    
    Dim oCrypto         As CCrypto
    Set oCrypto = New CCrypto
    '__ save the old info in oshared
    Call oShared.GetLocalLogon(sxUID, sxPWD)
    
    '__ get the c2t_x access info from the registry
    If Not GetC2TAccess(sAccUID, sAccPWD, oCrypto) Then

        Call GetSupportAppInfo(AppUID, AppPWD, Framework_PWD)
        If TmipAvail Then
            AppPWD = Framework_PWD
        Else
            '-- we use our own security admin db, but make the password editable in the reg
            SecUID = "secadmin"
            SecPWD = Framework_PWD
            DBKey = vbNullString
        End If
        
        GetUserAccess = True
        'Call oshared.SetLocalLogon(SecUID, SecUID)
        Exit Function
    End If

    
    '__ set these values in shared, then create an sqlops obj then set them back
    Call oShared.SetLocalLogon(sAccUID, sAccPWD)
    
    
    Set oC2T = New SqlOps
    Call oShared.SetLocalLogon(sxUID, sxPWD)
    
    '__ put back the shared info
    
    oC2T.Execute "select * from c2T_X where x0 = '" & oCrypto.GetHash(suid) & "'"
    If oC2T.EOF Then
       GetUserAccess = False
       Exit Function
    End If
        
    '__ confirm this is the right pwd
    If Not oCrypto.GetHash(spwd) = oC2T(db_UsrPWD) Then
        GetUserAccess = False
        Exit Function
    End If
    
    '__ decrypt the rest of the record
    On Error GoTo decrypterr
    AppUID = oCrypto.Decrypt(oC2T(db_APPUID), spwd, True)
    AppPWD = oCrypto.Decrypt(oC2T(db_APPPWD), spwd, True)
    SecUID = oCrypto.Decrypt(oC2T(db_SECUID), spwd, True)
    SecPWD = oCrypto.Decrypt(oC2T(db_SECPWD), spwd, True)
    DBKey = oCrypto.Decrypt(oC2T(db_DBKey), spwd, True)
    
    
    '__ x7 should be null, if not, then we have to change the passwords and it points to the start row in c2t_y
    NewPWDIdx = Val(vbNullString & oC2T(db_NewPWD))
    If Not NewPWDIdx = 0 Then
        If GetNewPWDs(AppUID, AppPWD, SecUID, SecPWD, DBKey, NewPWDIdx) Then
            Call SetUserAccess(suid, db_UsrPWD)
        End If
    End If
    
    GetUserAccess = True
    
    Exit Function
decrypterr:
    MsgBox "Decryption Error: " & Err.Description
    
End Function


Public Function ChangeC2TPassword(ByVal UserName As String, Optional ByVal oUser As C2T_Security.User = Nothing) As Boolean
    
On Error GoTo ErrHandler
    
Dim NewPWD As String
Dim frm As frmPassword
Dim oPwd As PassWordRules
Dim oCrypto As CCrypto
Dim oComm As MMMHISComm.Comm
Dim iHoursLeft As Long
Dim dPasswordChangeDate As Date
        
    Set frm = New frmPassword
    Set oPwd = New PassWordRules
    Set oCrypto = New CCrypto
    Set oComm = New MMMHISComm.Comm
    
    Set oPwd = GetPwdRules(UserName)
    
    ' unless coming from theater security, password change is not allowed within 24 hr of the previous change
    ' should we also give UserName "sysadmin" a break? no--he can use theater security to change the pwd
    If (Not oUser Is Nothing) And (Not oPwd.PwdChangeDate = Empty) Then
        If DateAdd("h", 24, oComm.GmtToLocal(oPwd.PwdChangeDate)) > Now Then
            dPasswordChangeDate = oComm.GmtToLocal(oPwd.PwdChangeDate)
            iHoursLeft = 24 - DateDiff("h", dPasswordChangeDate, Now)
            MsgBox "You cannot change your password twice within a 24-hour period." _
                & " You have approximately" & iHoursLeft _
                & " hours remaining.", vbInformation
            ChangeC2TPassword = False
            Exit Function
        End If
    End If
    
    frm.SetPasswordRules oPwd
    
    frm.lblUID.Caption = UserName
    'Set frm.currentUser = User
    
    '-- SCR 189686 B. Webb
    '-- using the frmCHCSIILogin form to center is unnecessary and prevents application from closing
    'frm.Top = frmCHCSIILogin.Top
    'frm.Left = (Screen.Width - frm.Width) / 2
    
    Do
        frm.txtNewPWD1.Text = vbNullString
        frm.txtNewPWD2.Text = vbNullString
        frm.Show vbModal
        If frm.bUserCanceled Then
            Exit Do
        End If
    Loop Until ValidatePassword(frm.txtOldPWD, frm.txtNewPWD1, oPwd)
    
    If Not frm.bUserCanceled Then
        NewPWD = oCrypto.GetHash(frm.txtNewPWD1.Text)
        
        'SCR 100543 - LCI - 9/11/2006: Change query to use UPPER and UCase
        If NewPWD <> vbNullString Then
            moSecSQL.Execute "Update user_resources set" _
                & " user_password = '" & NewPWD & "'," _
                & " PWD_CHANGE_DATE = " & SQLDate(oComm.LocalToGmt(Now)) & "," _
                & " IS_DEFAULT_PWD = 0," _
                & " OLD_USER_PWD10 = OLD_USER_PWD9, OLD_USER_PWD9 = OLD_USER_PWD8, OLD_USER_PWD8 = OLD_USER_PWD7, " _
                & " OLD_USER_PWD7 = OLD_USER_PWD6, OLD_USER_PWD6 = OLD_USER_PWD5, OLD_USER_PWD5 = OLD_USER_PWD4, " _
                & " OLD_USER_PWD4 = OLD_USER_PWD3, OLD_USER_PWD3 = OLD_USER_PWD2, OLD_USER_PWD2 = OLD_USER_PWD1, " _
                & " OLD_USER_PWD1 = USER_PASSWORD" _
                & " where UPPER(user_name) = " & SQLQuote(UCase(UserName)) _
                & " AND res_id = (select res_id from resources where resource_name = '" & THEATER_KEYWORD & "')"
            
            '<< SCR ??? update the hems users table too
            Dim oSQL As SqlOps
            Set oSQL = New SqlOps
            
            '--- SCR 79741   Sherry Wang  10/12/2005
            If oSQL.LoginInfoChanged = True Then Exit Function
            
            '--- SCR 83298   Sherry Wang 12/5/2005
            'SCR 100543 - LCI - 9/11/2006: Change query to use UPPER and UCase
            oSQL.Execute "Update hems_users set logon_passwd = " & SQLQuote(NewPWD) _
                  & " where UPPER(logon_name) = " & SQLQuote(UCase(UserName))
                  
            Call SetUserAccess(frm.lblUID.Caption, frm.txtNewPWD1.Text)
            
            'Set the new password in the user object
            If Not oUser Is Nothing Then
                oUser.Password = NewPWD
            End If
        Else
            MsgBox "The hash of the new password cannot be calculated.", vbInformation
        
        End If
        
        If Err.Number <> 0 Then
            Err.Clear
            MsgBox "Internal error.  Unable to change password"
        End If
    Else
        Exit Function
        
    End If
    
    Unload frm
    Set oComm = Nothing
    ChangeC2TPassword = True
    Exit Function
    
ErrHandler:
    MsgBox "Error: " & Err.Description & " occurred in ChangeC2TPassword", vbInformation
    
    Exit Function
    Resume
End Function


Private Function GetNewPWDs(ByRef AppUID As String, _
                            ByRef AppPWD As String, _
                            ByRef SecUID As String, _
                            ByRef SecPWD As String, _
                            ByRef DBKey As String, _
                            ByVal Index As Long) As Boolean
'__ resets the supplied passwords and keys with the current versions from the c2t_y table

    On Error GoTo ErrHandler

    Dim sSQL As String
    Dim sxUID           As String           '-- temp placeholder for whatever is in the shared.locallogon area
    Dim sxPWD           As String           '-- temp placeholder for whatever is in the shared.locallogon area
    Dim sAccUID         As String           '-- the registry-based c2t access uid/pwd
    Dim sAccPWD         As String           '-- the registry-based c2t access uid/pwd
    Dim oShared As CWShared
    Set oShared = New CWShared
    Dim oCrypto As CCrypto
    Set oCrypto = New CCrypto
    
    '__get the c2t access information and set the variables and create the connection
    If Not GetC2TAccess(sAccUID, sAccPWD, oCrypto) Then Exit Function
    Call oShared.GetLocalLogon(sxUID, sxPWD)        '-- get and hold current values
     
    Call oShared.SetLocalLogon(sAccUID, sAccPWD)    '-- set to c2t_x values
    Dim oC2T As SqlOps
    Set oC2T = New SqlOps
    Call oShared.SetLocalLogon(sxUID, sxPWD)        '-- replace whatever was there
    
    
    oC2T.Execute "Select * from c2t_y where Y0 >= " & Index & " order by Y0"

    If oC2T.EOF Then
    
    End If
    '-- here we work the magic of cascading down the rows, decrypting all the way to the last (current) one
    Do While Not oC2T.EOF
        AppUID = oCrypto.Decrypt(oC2T(db_NewAppUID), AppUID, True)
        AppPWD = oCrypto.Decrypt(oC2T(db_NewAppPWD), AppPWD, True)
        SecUID = oCrypto.Decrypt(oC2T(db_NewSecUID), SecUID, True)
        SecPWD = oCrypto.Decrypt(oC2T(db_NewSecPWD), SecPWD, True)
        DBKey = oCrypto.Decrypt(oC2T(db_NewDBKey), DBKey, True)
        oC2T.MoveNext
    Loop

    GetNewPWDs = True
    Exit Function
ErrHandler:
    GetNewPWDs = False
End Function


Private Function SetUserAccess(ByVal suid As String, ByVal spwd As String) As Boolean

    '>> mission: connect to the C2T schema
    '>>          find this user's record
    '>>          use the current pwd to decrypt the elements x2-x5
    '>>          use the new pwd to re-encrypt these elements
    '>>          save a hash of the new pwd.
    '>>          reset the 'newpwd index' (X7) to null
    Dim oShared As New CWShared
    Dim sxNSN As String
    Dim sxUID As String
    Dim sxPWD As String
    Dim sNSN As String
    Dim sAccUID As String
    Dim sAccPWD As String
    Dim oCrypto As CCrypto
    Set oCrypto = New CCrypto
    
'__ we need the actual sys info before we can continue
    If AppUID = vbNullString Or AppPWD = vbNullString Or _
       SecUID = vbNullString Or SecPWD = vbNullString Then
        MsgBox "Internal Error: Logon Information is not available."
        Exit Function
    End If
    '__ save the old info in oshared
    Call oShared.GetLocalLogon(sxUID, sxPWD)
    '__ set these values in shared, then create an sqlops obj then set them back
    
    
'<< SCR 33048 change password not working in non-encrypted db
    If Not GetC2TAccess(sAccUID, sAccPWD, oCrypto) Then Exit Function
    
    Call oShared.SetLocalLogon(sAccUID, sAccPWD)
    Dim oC2T As SqlOps
    Set oC2T = New SqlOps
    Call oShared.SetLocalLogon(sxUID, sxPWD)
    
    oC2T.Execute "select * from c2T_X where x0 = '" & oCrypto.GetHash(suid) & "'"
    If oC2T.EOF Then
        '__ this is a new user/entry
        oC2T.AddNew
        oC2T(db_UsrUID) = oCrypto.GetHash(suid)
    End If
    
    oC2T(db_UsrPWD) = oCrypto.GetHash(spwd, True)
    oC2T(db_APPUID) = oCrypto.Encrypt(AppUID, spwd, True)
    oC2T(db_APPPWD) = oCrypto.Encrypt(AppPWD, spwd, True)
    oC2T(db_SECUID) = oCrypto.Encrypt(SecUID, spwd, True)
    oC2T(db_SECPWD) = oCrypto.Encrypt(SecPWD, spwd, True)
    oC2T(db_DBKey) = oCrypto.Encrypt(DBKey, spwd, True)
    oC2T(db_NewPWD) = vbNullString
    oC2T.Update
    Set oC2T = Nothing
    
End Function
 

Private Function GetC2TAccess(ByRef suid As String, ByRef spwd As String, oCrypto As CCrypto) As Boolean
    
    Const sExp As String = "trtkabasnbi"

    Dim oShared As New CWShared
    On Error GoTo ErrHandler
    
    suid = oShared.GetSupportValue("C2T_UID")
    spwd = oShared.GetSupportValue("C2T_PWD")
    If Not spwd & suid = vbNullString Then
        spwd = oCrypto.Decrypt(spwd, oCrypto.GetHash(AppName & sExp, True), True, frezBlockEncryption)
        suid = oCrypto.Decrypt(suid, oCrypto.GetHash(AppName & sExp, True), True, frezBlockEncryption)
    End If
    If spwd = vbNullString Or suid = vbNullString Then
        GetC2TAccess = False
        Exit Function
    End If
    GetC2TAccess = True
    
    Exit Function
'        '___  euid=oCrypto.Encrypt(uid,  oCrypto.GetHash(appname & sExp, true), true)

ErrHandler:
    suid = vbNullString
    spwd = vbNullString
    
End Function


Private Function StartDas() As Boolean
    On Error GoTo ErrHandler
    
    Dim CWShared As CWShared
    Set CWShared = New CWShared
    Call CWShared.SetLocalLogon(SecUID, SecPWD)
    Set moSecSQL = New SqlOps
    
    '--- SCR 79741   Sherry Wang   10/14/2005
    If moSecSQL.LoginInfoChanged = True Then
        gbDBLoginInfoChanged = True
        Exit Function
    End If

    Call CWShared.SetLocalLogon(AppUID, AppPWD, DBKey)
    
    StartDas = True
    Exit Function
ErrHandler:
    MsgBox "Unexpected error in Security Module.StartDAS: " & vbCrLf & Err.Description, vbCritical, "Security Error"
End Function

Public Function ValidatePassword(OldPassword As String, _
                                 NewPassword As String, _
                                 PwdRules As PassWordRules, _
                                 Optional Lite As Boolean = False) As Boolean

'Validate the new password

    Dim i As Integer
    Dim iCharsDifferent As Integer
    Dim ch As Integer
    Dim msg As String
    Dim oCrypto As CCrypto
    Dim sEncryNewPwd As String
    
    '-- ascii values
    Const UPPER_MIN    As Integer = 65
    Const UPPER_MAX    As Integer = 90
    Const LOWER_MIN    As Integer = 97
    Const LOWER_MAX    As Integer = 122
    Const NUMBERS_MIN  As Integer = 48
    Const NUMBERS_MAX  As Integer = 57
    
    Dim AlphaUpper     As Integer
    Dim AlphaLower     As Integer
    Dim Num            As Integer
    Dim Non            As Integer
    
    On Error GoTo ErrHandler
    
    'Did they enter the old password correctly?
    If Lite = False Then
        Set oCrypto = New CCrypto
        If oCrypto.GetHash(OldPassword) <> PwdRules.EncryPwd Then
            msg = "The old password is not correct."
            GoTo InvalidEntry
        End If
    End If
    
    'New Password is of valid form
    With PwdRules
        If Len(NewPassword) < .MinLength Then
            msg = msg & vbCrLf & "The password must be at least " & .MinLength & " characters long."
            GoTo InvalidEntry
        End If
        If Len(NewPassword) > .MaxLength Then
            msg = msg & vbCrLf & "The password must be no more than " & .MaxLength & " characters long."
            GoTo InvalidEntry
        End If
        '-- count the numbers, letters, upper-case letters and special chars
        For i = 1 To Len(NewPassword)
            ch = Asc(Mid$(NewPassword, i, 1))
            If (ch >= UPPER_MIN And ch <= UPPER_MAX) Then
                AlphaUpper = AlphaUpper + 1
            ElseIf (ch >= LOWER_MIN And ch <= LOWER_MAX) Then
                AlphaLower = AlphaLower + 1
            ElseIf (ch >= NUMBERS_MIN And ch <= NUMBERS_MAX) Then
                Num = Num + 1
            Else
                Non = Non + 1
            End If
        Next i
        '-- case
        If .MixedCase And (AlphaUpper = 0 Or AlphaLower = 0) Then
            '--- SCR 88971  Sherry Wang  2/28/2006
            msg = msg & vbCrLf & "The password must contain both upper and lower case letters."
        End If
        '--numbers
        If Num < .MinNumeric Then '.MinNumeric > 0 And Num = 0 Then 'SCR 87891
            msg = msg & vbCrLf & "The password must contain at least " & .MinNumeric & " numeric character(s)."
        End If
        If Non < .NumberNonAlphanumeric Then ' .NumberNonAlphanumeric > 0 And Non = 0 Then 'SCR 87890
            msg = msg & vbCrLf & "The password must contain at least " & .NumberNonAlphanumeric & " non-alphanumeric character(s)."
        End If
        
        'SCR 76583 Don't allow a @ to be a special character.
        'SCR 76583 - LCI - 09/21/2006
        If InStr(NewPassword, "@") > 0 Or InStr(NewPassword, "'") > 0 Or InStr(NewPassword, "%") > 0 Then
            msg = msg & vbCrLf & "The @, ', and % symbols are not allowed in the password."
        End If
        
        If Lite = False Then
            'Compare the new password to the ten previous passwords (jq: used to be four previous)
            Set oCrypto = New CCrypto
            sEncryNewPwd = oCrypto.GetHash(NewPassword)
            If sEncryNewPwd = .EncryPwd1 Or sEncryNewPwd = .EncryPwd2 Or sEncryNewPwd = .EncryPwd3 Or sEncryNewPwd = .EncryPwd4 _
                Or sEncryNewPwd = .EncryPwd5 Or sEncryNewPwd = .EncryPwd6 Or sEncryNewPwd = .EncryPwd7 _
                Or sEncryNewPwd = .EncryPwd8 Or sEncryNewPwd = .EncryPwd9 Or sEncryNewPwd = .EncryPwd10 Then
                msg = msg & vbCrLf & "The password must not be the same as your ten previous passwords."
            End If
            
            'Make sure at least n characters have changed from the old password
            iCharsDifferent = CharsInXNotInY(NewPassword, OldPassword) '+ CharsInXNotInY(OldPassword, NewPassword)
            If iCharsDifferent < .NumberCharsChanged Then
                msg = msg & vbCrLf & "At least " & .NumberCharsChanged & " characters must be different from your old password."
            End If
        End If
        
        If msg <> vbNullString Then
            GoTo InvalidEntry
        End If
        
    End With
    ValidatePassword = True
    
    Exit Function
    
ErrHandler:
    MsgBox "Error: " & Err.Description & " occurred in clsPasswordOps.ValidateNewPassword"
    Exit Function
    
InvalidEntry:
    msg = "The password provided violates the following criteria: " & vbCrLf & msg
    MsgBox msg, vbInformation, "Invalid Password."
    
    ValidatePassword = False
    Exit Function

End Function

Private Function CharsInXNotInY(ByVal x As String, ByVal y As String) As Integer
    
'Count the number of chars in X that are not in Y
    
Dim i As Integer
Dim j As Integer
Dim bFound As Boolean

        For i = 1 To Len(x)
            bFound = False
            For j = 1 To Len(y)
                If Mid(x, i, 1) = Mid(y, j, 1) Then
                    bFound = True
                    Exit For
                End If
            Next j
            If bFound = False Then
                CharsInXNotInY = CharsInXNotInY + 1
            End If
        Next i

End Function

Private Function GetPwdRules(UserName As String, Optional Lite As Boolean = False) As PassWordRules
    
Dim oPwd As PassWordRules
    
    Set oPwd = New PassWordRules
    
    'SF 2/13/06 SCRs 87890, 87891, 87892, 87894
    oPwd.MaxLength = 26
    oPwd.MinLength = 15
    oPwd.MinNumeric = 2
    oPwd.MixedCase = 1
    oPwd.NumberNonAlphanumeric = 2
    oPwd.NumberCharsChanged = 4
    
    If Not Lite Then
        'Get the hashed values of the previous passwords
        'SCR 100543 - LCI - 08/24/2006
        moSecSQL.Execute ("Select USER_PASSWORD, OLD_USER_PWD1, OLD_USER_PWD2, OLD_USER_PWD3, OLD_USER_PWD4, " _
                & " OLD_USER_PWD5, OLD_USER_PWD6, OLD_USER_PWD7, OLD_USER_PWD8, OLD_USER_PWD9, OLD_USER_PWD10, " _
                & " PWD_CHANGE_DATE from USER_RESOURCES" _
                & " where UPPER(user_name) = " & SQLQuote(UCase(UserName)) _
                & " AND res_id = (select res_id from resources where resource_name = '" & THEATER_KEYWORD & "')")
        If Not moSecSQL.EOF Then
            With oPwd
                .EncryPwd = moSecSQL("USER_PASSWORD") & ""
                .EncryPwd1 = moSecSQL("OLD_USER_PWD1") & ""
                .EncryPwd2 = moSecSQL("OLD_USER_PWD2") & ""
                .EncryPwd3 = moSecSQL("OLD_USER_PWD3") & ""
                .EncryPwd4 = moSecSQL("OLD_USER_PWD4") & ""
                .EncryPwd5 = moSecSQL("OLD_USER_PWD5") & ""
                .EncryPwd6 = moSecSQL("OLD_USER_PWD6") & ""
                .EncryPwd7 = moSecSQL("OLD_USER_PWD7") & ""
                .EncryPwd8 = moSecSQL("OLD_USER_PWD8") & ""
                .EncryPwd9 = moSecSQL("OLD_USER_PWD9") & ""
                .EncryPwd10 = moSecSQL("OLD_USER_PWD10") & ""
                If moSecSQL("PWD_CHANGE_DATE") = Empty Then
                  .PwdChangeDate = Empty
                Else
                  .PwdChangeDate = moSecSQL("PWD_CHANGE_DATE")    ' Note: this is in GMT
                End If
            End With
        End If
    End If
    
    Set GetPwdRules = oPwd
    
End Function

Public Function ValidatePasswordLite(NewPassword As String)
    
Dim oPwdRules As PassWordRules
    
    Set oPwdRules = GetPwdRules("", True)
    
    ValidatePasswordLite = ValidatePassword("", NewPassword, oPwdRules, True)
    
End Function
