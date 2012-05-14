VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "User"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements ICHCSII_User

Public objSec           As ICHCSII_Security

Private m_CardID        As String           '-- used for validating the PIN and the Card ID
Private m_sUserName     As String           '-- the userid for snareworks
Private m_sUserId       As String           '-- the fully qualified userid for snareworks, including the domain/organization
Private m_sPassword     As String           '-- the user's password into snareworks
Private m_objRoles      As ICHCSII_Roles    '-- a collection of roles
Private m_objResources  As ICHCSII_Resources
Private m_colPrivileges As New Collection   '-- of keyvalue pairs
Private m_colAttributes As New Collection   '-- of keyvalue pairs
'Private m_nAuthUser As Boolean
Private m_LoginMode     As CHCSII_LoginMode
Private m_CACPresent    As Boolean

Private oCAC_Monitor As cCardMonProxy
Attribute oCAC_Monitor.VB_VarHelpID = -1

Friend Property Let CardID(rhv As String)
    m_CardID = rhv
End Property
Friend Property Set CAC_Monitor(rhv As cCardMonProxy)
    Set oCAC_Monitor = rhv
End Property

Friend Property Let CACPresent(rhv As Boolean)
    m_CACPresent = rhv
End Property
Friend Property Let LoginMode(rhv As CHCSII_LoginMode)
    m_LoginMode = rhv
End Property
Friend Property Get Password() As String
    Password = m_sPassword
End Property

Private Sub Class_Initialize()
 '   m_nAuthUser = False
 Set m_colPrivileges = New Collection
 Set m_colAttributes = New Collection
End Sub

Friend Property Let UserName(ByVal UserName As String)
    m_sUserName = UserName
End Property

Friend Property Let UserId(ByVal UserId As String)
    m_sUserId = UserId
End Property

Friend Property Get UserId() As String
    UserId = m_sUserId
End Property

Friend Property Let Password(ByVal Password As String)
    m_sPassword = Password
    Call RecordPWDValidation(m_sUserId, vbNullString, m_sPassword)
End Property

Friend Property Set Roles(userRoles As ICHCSII_Roles)
    Set m_objRoles = userRoles
End Property

Friend Property Set Privileges(colPrivileges As Collection)
    Set m_colPrivileges = colPrivileges
End Property

Friend Property Set Resources(userResources As ICHCSII_Resources)
    Set m_objResources = userResources
End Property

Friend Property Set Attributes(userAttributes As Collection)
    Set m_colAttributes = userAttributes
End Property

Friend Function AddAttribute(ByVal Key As String, _
                             ByVal Value As String, _
                    Optional ByVal Update As Boolean = False) As Boolean
                    
    '-- add an attribute without repeats
    '   used to add the ncid (mostly)
    Dim kvp As KeyValuePair
    
    For Each kvp In m_colAttributes
        If LCase$(kvp.Key) = LCase$(Key) Then       '-- already there, don't add again
            AddAttribute = False
            Exit Function
        End If
    Next kvp
    
    '-- ok, it's not already here, add it
    Set kvp = New KeyValuePair
    kvp.Key = Key
    kvp.Value = Value
    m_colAttributes.Add kvp, kvp.Key
    
End Function

Private Function ICHCSII_User_CACPresent() As Boolean
    On Error GoTo errhandler
    
    ICHCSII_User_CACPresent = oCAC_Monitor.CardPresent
    
    Exit Function
errhandler:
    err.Raise err.Number, err.Source, err.Description
    
End Function

Private Function ICHCSII_User_CACSignText(ByVal PIN As String, _
                                          ByVal Text As String, _
                                                Signature As String, _
                                       Optional ErrCode As Long, _
                                       Optional ErrMsg As String) As Boolean

    Dim oCrypto As CCrypto
    Dim HText As String
    Dim i As Long
    On Error GoTo errhandler
    
    
'    Shell "openssl rsautl -in c:\sig.bin -verify -raw -inkey c:\temp.pem -certin -out c:\sigout.txt"
'    Str = oCrypto.GetHash("This is the text that I want signed", True)
'    Str = Str & Str
'    Str = Str & Str
'    Call CAC_Monitor.SignText(PIN, Str, Sig)

    
    
    
    '-- must have a 128-byte hash for the text.
    If Len(Text) <> 128 Then
        Set oCrypto = New CCrypto
        HText = oCrypto.GetHash(Text)
        For i = Len(HText) + 1 To 128
            HText = HText & "*"
        Next i
    Else
        HText = Text
    End If
    ICHCSII_User_CACSignText = oCAC_Monitor.SignText(PIN, HText, Signature, ErrCode, ErrMsg)

    Exit Function
errhandler:
    err.Raise ErrCode, "CAC Sign Text", ErrMsg
    
End Function

Private Function ICHCSII_User_CACValidatePIN(ByVal PIN As String) As Boolean
    On Error GoTo errhandler

    ICHCSII_User_CACValidatePIN = oCAC_Monitor.ValidatePIN(PIN)

    Exit Function
errhandler:
    err.Raise err.Number, err.Source, err.Description
    
End Function

Private Function ICHCSII_User_GetAttribute(AttributeName As String) As String
    On Error GoTo errhandler
    ICHCSII_User_GetAttribute = m_colAttributes.Item(AttributeName).Value
    Exit Function
errhandler:
    ICHCSII_User_GetAttribute = vbNullString
    err.Raise SEC_NO_ATTRIBUTE_FOUND, vbNullString, GetErrMsg(SEC_NO_ATTRIBUTE_FOUND)
End Function

Private Function ICHCSII_User_GetAttributeNames() As Collection
    Dim colAttributeNames As New Collection
    Dim i As Integer
    
    For i = 1 To m_colAttributes.Count
        colAttributeNames.Add m_colAttributes(i).Key
    Next
    
    Set ICHCSII_User_GetAttributeNames = colAttributeNames
End Function

Private Function ICHCSII_User_HasPrivilege(sPrivilege As String) As Boolean
'-- just a simple functional area check.  The specific privilege (read/write) isn't looked for

    Dim s As String
    s = m_colPrivileges.Item(sPrivilege)
    ICHCSII_User_HasPrivilege = True
    Exit Function
    
errhandler:
    ICHCSII_User_HasPrivilege = False
End Function


Private Function ICHCSII_User_HasPrivilegeEx(PrivilegeArea As String, PrivilegeType As CHCSII_PrivilegeType, Optional OtherType As Integer) As Boolean
    
    Dim nPerm As Long
    
'<< scr 35657 add cosign and change the meaning of sign.
    nPerm = 0
    Select Case PrivilegeType
        Case Priv_Read:                 nPerm = 2
        Case Priv_Write:                nPerm = 4
        Case Priv_Delete:               nPerm = 8
        Case Priv_Sign:                 nPerm = 64
        Case Priv_Need_Cosign:          nPerm = 32
        Case Priv_Cosign:               nPerm = 128
        Case Priv_Break_The_Glass:      nPerm = 16
        Case Else:                      nPerm = 0
    End Select
    
'-- LogAccess "In user has priv for " & m_sUserId & " - " & PrivilegeArea & " - " & PrivilegeType
    On Error GoTo errhandler
'-- this is a collection of keyvalue pairs.  the value is total of the permission numbers listed above.
    
    ICHCSII_User_HasPrivilegeEx = CBool(m_colPrivileges(PrivilegeArea).Value And nPerm)
    
    Exit Function

errhandler:
'__ cascade this up to the calling function
    If err.Number <> 0 Then
        Select Case err.Number
        Case 5    '-- there was no hit on the priv collection for this priv area
            ICHCSII_User_HasPrivilegeEx = False
    
        Case 1003
            err.Raise 1003, "Security Service", "Unable to communicate with the security service. Privileges can not be confirmed at this time."
        Case Else
            err.Raise 1002, "Security Service", "Unexpected error received from security service. Privileges can not be confirmed at this time." & vbCrLf & err.Number & " - " & err.Description
        End Select
    End If
End Function


Private Property Get ICHCSII_User_LoginMode() As CHCSII_ClientSecurity.CHCSII_LoginMode
    ICHCSII_User_LoginMode = m_LoginMode
End Property

Private Property Get ICHCSII_User_Resources() As CHCSII_ClientSecurity.ICHCSII_Resources
    Set ICHCSII_User_Resources = m_objResources
End Property

Private Property Get ICHCSII_User_Roles() As CHCSII_ClientSecurity.ICHCSII_Roles
    Set ICHCSII_User_Roles = m_objRoles
End Property

Private Function ICHCSII_User_SignText(ByVal sText As String) As String

End Function

Private Property Get ICHCSII_User_UserId() As String
    ICHCSII_User_UserId = m_sUserId
End Property

Private Property Get ICHCSII_User_UserName() As String
    ICHCSII_User_UserName = m_sUserName
End Property

Private Function ICHCSII_User_ValidatePassword(Password As String) As Boolean
    Dim CardID As String        '-- the unique id of a CAC
    If Password = m_sPassword Then
        ICHCSII_User_ValidatePassword = True
    Else
    '-- scr #50532  overload with a call to cac/pin
        If Not oCAC_Monitor Is Nothing Then
            If m_LoginMode = Mode_CAC Then
                If oCAC_Monitor.ValidatePIN(Password, CardID) Then
                    ICHCSII_User_ValidatePassword = (m_CardID = CardID)
                End If
            End If
        End If
    End If
    
'<< SCR 19749
    Call RecordPWDValidation(m_sUserId, Password, m_sPassword)
End Function

'__ for debug only
Friend Function DumpObject() As Boolean
    Dim sout    As String
    Dim MyCol   As Collection
    Dim kvp     As KeyValuePair
    Dim oRole   As ICHCSII_Role
    Dim oRes    As ICHCSII_Resource
    Dim i       As Long
    Dim j       As Long
    Dim sBit    As String
    Dim nval    As Long
    Dim oCrypto As New CCrypto
    

    On Error GoTo errhandler
    Dim oShared As New CWShared
    '<< 10DEC03 - if called by midtier, cw has not set the gobjcmdlineswitches
    If oShared.CmdLineSwitches Is Nothing Then Exit Function
    
    If oShared.CmdLineSwitches("DUMPUSERINFO") = vbNullString Then
        Exit Function
    End If
    
    sout = "user Name:" & vbTab & m_sUserName & vbCrLf
    sout = sout & "ID:       " & m_sUserId & vbCrLf
    sout = sout & "password: " & oCrypto.GetHash(m_sPassword) & vbCrLf
    
    sout = sout & vbCrLf & "Roles:    " & vbCrLf
    For i = 1 To m_objRoles.Count
        Set oRole = m_objRoles.Item(i)
        sout = sout & vbTab & oRole.Name & vbCrLf
    Next i
    
    sout = sout & vbCrLf & "Resources:" & vbCrLf
    For i = 1 To m_objResources.Count
        Set oRes = m_objResources.Item(i)
        sout = sout & vbTab & oRes.Name & vbCrLf
        sout = sout & vbTab & vbTab & "User ID:  " & oRes.UserId & vbCrLf
        sout = sout & vbTab & vbTab & "Password: " & oCrypto.GetHash(oRes.Password) & vbCrLf
        sout = sout & vbTab & vbTab & "Attributes: " & vbCrLf
        For j = 1 To oRes.GetAttributeNames.Count
            If oRes.GetAttributeNames(j) = "VMSPassword" Then
                sout = sout & vbTab & vbTab & vbTab & oRes.GetAttributeNames(j) & " = " & oCrypto.GetHash(oRes.getAttribute(oRes.GetAttributeNames(j))) & vbCrLf
            Else
                sout = sout & vbTab & vbTab & vbTab & oRes.GetAttributeNames(j) & " = " & oRes.getAttribute(oRes.GetAttributeNames(j)) & vbCrLf
            End If
        Next j
    Next i
    
    sout = sout & "Attributes:" & vbCrLf
    For Each kvp In m_colAttributes
        If kvp.Key = "SSN" Then
            sout = sout & vbTab & kvp.Key & " = " & oCrypto.GetHash(kvp.Value) & vbCrLf
        Else
            sout = sout & vbTab & kvp.Key & " = " & kvp.Value & vbCrLf
        End If
    Next kvp
    sout = sout & "Privileges:" & vbCrLf

   
        For Each kvp In m_colPrivileges
            sBit = vbNullString
            nval = kvp.Value

            If nval And 2 Then sBit = sBit & "R"                    ' Read
            If nval And 4 Then sBit = sBit & "W"                    ' Write
            If nval And 8 Then sBit = sBit & "D"                    ' Delete
            If nval And 128 Then sBit = sBit & "C"                  ' CoSign
            If nval And 16 Then sBit = sBit & "B"                   ' BreakGlass
            If nval And 32 Then sBit = sBit & "N"                   ' NeedCoSign
            If nval And 64 Then sBit = sBit & "S"                   ' Sign
            sout = sout & vbTab & vbTab & kvp.Key & vbTab & vbTab & sBit & vbCrLf
        Next

    Dim f_out As Integer
    f_out = FreeFile
    Open "c:\stubout.txt" For Output As f_out
    Print #f_out, sout
    Close #f_out
    Exit Function
errhandler:
    If err.Number = SEC_NO_ATTRIBUTE_FOUND Then
        Resume Next
    End If
End Function




