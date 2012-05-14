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
Attribute VB_Name = "User"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements CHCSII_ClientSecurity.ICHCSII_User

Private Const CLASS_NAME As String = "C2T_Security.User"

Private m_sUserName As String
Private m_sUserId   As String
Private m_sPassword As String
Private m_objRoles  As CHCSII_ClientSecurity.ICHCSII_Roles
Private m_objResources  As CHCSII_ClientSecurity.ICHCSII_Resources
Private m_colPrivileges As New Collection    ' of keyValuePairs - priv name and priv string
Private m_colAttributes As New Collection    ' of keyValuePairs
Private m_nAuthUser As Boolean
Private m_sKeyPair  As Boolean
Private m_LoginMode As CHCSII_LoginMode
Private m_CardPresent   As Boolean

Friend Property Let CardPresent(RHV As Boolean)
    m_CardPresent = RHV
End Property

Private Sub Class_Initialize()
    m_nAuthUser = False
    m_LoginMode = Mode_Manual
End Sub

Friend Property Let UserName(ByVal UserName As String)
    m_sUserName = UserName
End Property

Friend Property Get UserName() As String
    UserName = m_sUserName
End Property

Public Property Let UserId(ByVal UserId As String)
    m_sUserId = UserId
End Property

Public Property Get UserId() As String
    UserId = m_sUserId
End Property

Public Property Get Password() As String
    Password = m_sPassword
End Property
Public Property Let Password(ByVal Password As String)
    m_sPassword = Password
End Property

Friend Property Set Roles(userRoles As ICHCSII_Roles)
    Set m_objRoles = userRoles
End Property

Friend Property Set Resources(userResources As ICHCSII_Resources)
    Set m_objResources = userResources
End Property

Friend Property Set Attributes(userAttributes As Collection)
    Set m_colAttributes = userAttributes
End Property
Friend Property Set Privileges(RHV As Collection)
    Set m_colPrivileges = RHV
End Property

Private Function ICHCSII_User_CACPresent() As Boolean
    ICHCSII_User_CACPresent = False
End Function

Private Function ICHCSII_User_CACSignText(ByVal PIN As String, ByVal Text As String, Signature As String, Optional ErrCode As Long, Optional ErrMsg As String) As Boolean
    ErrCode = SEC_UNSUPPORTED_FEATURE
    ErrMsg = GetErrMsg(SEC_UNSUPPORTED_FEATURE)
    ICHCSII_User_CACSignText = False
End Function

Private Function ICHCSII_User_CACValidatePIN(ByVal PIN As String) As Boolean
    ICHCSII_User_CACValidatePIN = False
End Function

Private Function iCHCSII_User_GetAttribute(AttributeName As String) As String
    On Error Resume Next
    iCHCSII_User_GetAttribute = m_colAttributes.Item(AttributeName).Value
    If Err Then
        On Error GoTo 0
        Err.Raise SEC_NO_ATTRIBUTE_FOUND, CLASS_NAME, GetErrMsg(SEC_NO_ATTRIBUTE_FOUND)
    End If
End Function

Private Function iCHCSII_User_GetAttributeNames() As Collection
    Dim colAttributeNames As New Collection
    Dim i As Integer
    
    For i = 1 To m_colAttributes.Count
        colAttributeNames.Add m_colAttributes(i).Key
    Next
    
    Set iCHCSII_User_GetAttributeNames = colAttributeNames
End Function

Private Function iCHCSII_User_HasPrivilege(sPrivilege As String) As Boolean
    On Error Resume Next
    
    Dim oPair As KeyValuePair
    For Each oPair In m_colPrivileges
        If UCase(oPair.Key) = UCase(sPrivilege) Then
            iCHCSII_User_HasPrivilege = True
            Exit Function
        End If
    Next oPair

    
End Function

Private Function HasRole(ByVal role As String) As Boolean
    Dim objRole As role
    Dim sRole As String
    
    Dim i As Integer
    
    For i = 1 To m_objRoles.Count()
        Set objRole = m_objRoles.Item(i)
        sRole = objRole.Name
        If StrComp(role, sRole) = 0 Then
            HasRole = True
            Exit Function
        End If
    Next i
    HasRole = False
End Function

Private Function ICHCSII_User_HasPrivilegeEx(PrivilegeArea As String, PrivilegeType As CHCSII_ClientSecurity.CHCSII_PrivilegeType, Optional OtherType As Integer) As Boolean
    On Error Resume Next
    Dim oPair As KeyValuePair
    Dim sPrivType As String
    
    '- convert PrivilegeType to string
    
    sPrivType = PrivTypeValueToText(PrivilegeType)
    
    '__ special case for TMIP framework only:
    If m_sPassword = "PRIVILEGE_CHECK_ONLY" Then
        ICHCSII_User_HasPrivilegeEx = TMIP_Get_Priv(PrivilegeArea, sPrivType)
    Else
        For Each oPair In m_colPrivileges
            If UCase(oPair.Key) = UCase(PrivilegeArea) Then
                ICHCSII_User_HasPrivilegeEx = (InStr(oPair.Value, sPrivType) > 0)
                Exit Function
            End If
        Next oPair
    End If
End Function


Private Property Get ICHCSII_User_LoginMode() As CHCSII_ClientSecurity.CHCSII_LoginMode
    ICHCSII_User_LoginMode = m_LoginMode
End Property

Private Property Get ICHCSII_User_Resources() As ICHCSII_Resources
    Set ICHCSII_User_Resources = m_objResources
End Property

Private Property Get iCHCSII_User_Roles() As ICHCSII_Roles
    Set iCHCSII_User_Roles = m_objRoles
End Property


Private Property Get iCHCSII_User_UserId() As String
    iCHCSII_User_UserId = m_sUserId
End Property

Private Property Get iCHCSII_User_UserName() As String
    iCHCSII_User_UserName = m_sUserName
End Property

Private Function iCHCSII_User_ValidatePassword(Password As String) As Boolean
    '<< SCR 38285;   Developer: Brian Mowbray 06/30/2003 09:25 AM
    'This is added just in case
    Dim oCrypto As CCrypto
    Set oCrypto = New CCrypto

    iCHCSII_User_ValidatePassword = (oCrypto.GetHash(Password) = m_sPassword)
End Function


Private Function SignDocument(ByVal DocStr As String, ByVal Password As String, ByRef Signature As String, ByRef ErrorStr As String) As Boolean
' pass in the string to sign, the password to unlock the user's sig and I'll pass back the signature
    
End Function
Private Function TMIP_Get_Priv(ByVal Priv_Name As String, _
                               ByVal Priv_String As String) As Boolean

'-- priv_string is the privilege they want to see.  Convert to string
'__ from NGIT:
'CREATE OR REPLACE PROCEDURE Get_Priv (
'                v_user_id IN varchar2,          -- User Login ID    - Input
'                v_priv_name IN varchar2,        -- Privilege Name   - Input
'                v_priv_str OUT varchar2)        -- Privilege String - Output
'
'        -- Get_Priv returns empty string if the v_user_id or
'        -- v_priv_name is invalid. (it does NOT raise exception)
    Dim oSQL As SqlOps
    Set oSQL = New SqlOps
'__ if the stored proc isn't there, just leave with the default of 'false'
    If Not GetTMIPStatus(oSQL) Then
        Exit Function
    End If
    
    Dim Param As ADODB.Parameter
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    
    With cmd
        .CommandType = adCmdStoredProc
        .CommandText = "GET_PRIV"

        Set Param = .CreateParameter("v_user_id", adVarChar, adParamInput, 128, m_sUserName)
            .Parameters.Append Param
        Set Param = .CreateParameter("v_priv_name", adVarChar, adParamInput, 128, Priv_Name)
            .Parameters.Append Param
        Set Param = .CreateParameter("v_priv_str", adVarChar, adParamOutput, 128, vbNullString)
            .Parameters.Append Param
    End With
    oSQL.ExecuteStoredProc cmd
    
    TMIP_Get_Priv = (InStr(cmd.Parameters("v_priv_str").Value, Priv_String) > 0)
End Function

Private Function GetTMIPStatus(oSQL As SqlOps) As Boolean
    On Error GoTo errhandler
   
    oSQL.Execute "select count(*)  from all_objects where object_name = 'GET_PRIV'"
    GetTMIPStatus = (Val(vbNullString & oSQL(0)) > 0)

errhandler:
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
    
    
    On Error GoTo errhandler
    Dim oShared As New CWShared
    '<< 10DEC03 - if called by midtier, cw has not set the gobjcmdlineswitches
    If oShared.CmdLineSwitches Is Nothing Then Exit Function
    
    If oShared.CmdLineSwitches("DUMPUSERINFO") = vbNullString Then
        Exit Function
    End If
    
    
    
    sout = "user Name:" & vbTab & m_sUserName & vbCrLf
    sout = sout & "ID:" & vbTab & m_sUserId & vbCrLf
    sout = sout & "password:" & vbTab & m_sPassword & vbCrLf
    sout = sout & "Roles:" & vbCrLf
    For i = 1 To m_objRoles.Count
        Set oRole = m_objRoles.Item(i)
        sout = sout & vbTab & oRole.Name & vbCrLf
    Next i
    
    sout = sout & "Resources:" & vbCrLf
    For i = 1 To m_objResources.Count
        Set oRes = m_objResources.Item(i)
        sout = sout & vbTab & oRes.Name & vbCrLf
        sout = sout & vbTab & vbTab & "User ID:  " & vbTab & oRes.UserId & vbCrLf
        sout = sout & vbTab & vbTab & "Password: " & vbTab & oRes.Password & vbCrLf
        sout = sout & vbTab & vbTab & "Attributes: " & vbCrLf
        For j = 1 To oRes.GetAttributeNames.Count
            sout = sout & vbTab & vbTab & vbTab & oRes.GetAttributeNames(j) & " = " & oRes.GetAttribute(oRes.GetAttributeNames(j)) & vbCrLf
        Next j
    Next i
    
    sout = sout & "Attributes:" & vbCrLf
    For Each kvp In m_colAttributes
        sout = sout & vbTab & kvp.Key & " = " & kvp.Value & vbCrLf
    Next kvp
    sout = sout & "Privileges:" & vbCrLf

   
        For Each kvp In m_colPrivileges
            If kvp.Value <> "" Then
                sout = sout & vbTab & vbTab & kvp.Key & vbTab & vbTab & kvp.Value & vbCrLf
            End If

        Next

    Dim f_out As Integer
    f_out = FreeFile
    Open "c:\stubout.txt" For Output As f_out
    Print #f_out, sout
    Close #f_out

errhandler:

End Function

Friend Function Initialize(UserId As String, colUserProperties As Collection, l_Resources As Resources, l_Roles As Roles)

'Accepts a collection of users as defined in the security_stub.txt file

Dim colResources As Collection
Dim colRoles As Collection
Dim colPrivileges As Collection
Dim colAttributes As Collection
Dim colTemp As Collection

Dim objKeyValuePair As KeyValuePair
Dim userRoles As New Roles
Dim userResources As New Resources
Dim l_Role As role
Dim l_Resource As Resource
Dim sResourceUserId As String
Dim sResourcePassword As String

Dim i As Integer
Dim j As Integer

m_sUserId = UserId
m_sUserName = colUserProperties.Item("username")
m_sPassword = colUserProperties.Item("password")

'do resources
Set colResources = colUserProperties.Item("resources").Value

For i = 1 To colResources.Count
    Set objKeyValuePair = colResources.Item(i)
    Set l_Resource = l_Resources.Item(objKeyValuePair.Key)
    Set colTemp = objKeyValuePair.Value
    
    On Error Resume Next
    sResourceUserId = colTemp.Item("userid")
    If Err Then
        Err.Clear
        sResourceUserId = ""
    End If
    
    sResourcePassword = colTemp.Item("password")
    If Err Then
        Err.Clear
        sResourcePassword = ""
    End If
    On Error GoTo 0
    
    userResources.AddResource l_Resource.Name, sResourceUserId, sResourcePassword, l_Resource.ResourceData, l_Resource.Attributes
    
Next
Set m_objResources = userResources

'do roles and privileges
Set colRoles = colUserProperties.Item("roles").Value

For i = 1 To colRoles.Count
    Set objKeyValuePair = colRoles.Item(i)
    Set l_Role = l_Roles.Item(objKeyValuePair.Value)
    userRoles.AddRole l_Role.Name
    
    For j = 1 To l_Role.Privileges.Count
        m_colPrivileges.Add l_Role.Privileges.Item(j), l_Role.Privileges.Item(j)
    Next
Next
Set m_objRoles = userRoles

'do additional/removed privileges
Set colPrivileges = colUserProperties.Item("privileges").Value
AddPrivileges colPrivileges.Item("additional").Value
RemovePrivileges colPrivileges.Item("removed").Value
    
'do attributes
On Error Resume Next        'May not have attributes key
Set m_colAttributes = colUserProperties.Item("attributes").Value

    
End Function

Private Sub AddPrivileges(colPrivileges As Collection)

'__colprivileges is a collection of keyvalue pairs
Dim colTemp As Collection
Dim sPrivilegeName As String
Dim sPrivilegeValue As String

Dim j As Integer
Dim k As Integer

On Error Resume Next

Set colTemp = colPrivileges
For j = 1 To colTemp.Count

    If colTemp.Item(j).Key = "" Then
        m_colPrivileges.Add colTemp.Item(j).Value, colTemp.Item(j).Value
        m_colPrivileges.Add colTemp.Item(j).Value & CStr(PrivValue(1)), colTemp.Item(j).Value & CStr(PrivValue(1))    'Default read privilege
    Else
        'Add specified privileges
        sPrivilegeName = colTemp.Item(j).Key
        sPrivilegeValue = colTemp.Item(j).Value
        
        'kluge multi-privilege per privilege area
        If InStr(1, sPrivilegeValue, "Z") > 0 Then
            sPrivilegeValue = "RWDSNBC"
        End If
        For k = 1 To UBound(PrivType)
            If InStr(1, sPrivilegeValue, PrivType(k)) > 0 Then
                m_colPrivileges.Add sPrivilegeName & CStr(PrivValue(k)), sPrivilegeName & CStr(PrivValue(k))
            End If
        Next
    End If
Next

End Sub

Private Sub RemovePrivileges(colPrivileges As Collection)

Dim colTemp As Collection
Dim sPrivilegeName As String
Dim sPrivilegeValue As String
Dim sPrivilegeTypes(7) As String
Dim iPrivilegeTypeValues(7) As Integer

Dim j As Integer
Dim k As Integer

sPrivilegeTypes(1) = "R"    'Read
sPrivilegeTypes(2) = "W"    'Write
sPrivilegeTypes(3) = "D"    'Delete
sPrivilegeTypes(4) = "S"    'Sign
sPrivilegeTypes(5) = "N"    'Need Cosign
sPrivilegeTypes(6) = "B"    'Break the Glass
sPrivilegeTypes(7) = "C"    'Can Cosign

iPrivilegeTypeValues(1) = Priv_Read
iPrivilegeTypeValues(2) = Priv_Write
iPrivilegeTypeValues(3) = Priv_Delete
iPrivilegeTypeValues(4) = Priv_Sign
iPrivilegeTypeValues(5) = Priv_Need_Cosign
iPrivilegeTypeValues(6) = Priv_Break_The_Glass
iPrivilegeTypeValues(7) = Priv_Cosign

On Error Resume Next

Set colTemp = colPrivileges
For j = 1 To colTemp.Count
    If colTemp.Item(j).Key = "" Then
        m_colPrivileges.Remove colTemp.Item(j).Value
        m_colPrivileges.Remove colTemp.Item(j).Value & CStr(iPrivilegeTypeValues(1))
    Else
        'Remove specified privileges
        sPrivilegeName = colTemp.Item(j).Key
        sPrivilegeValue = colTemp.Item(j).Value
        
        'kluge multi-privilege per privilege area
        If InStr(1, sPrivilegeValue, "Z") > 0 Then
            sPrivilegeValue = "RWDSNBC"
        End If
        For k = 1 To UBound(sPrivilegeTypes)
            If InStr(1, sPrivilegeValue, sPrivilegeTypes(k)) > 0 Then
                m_colPrivileges.Remove sPrivilegeName & CStr(iPrivilegeTypeValues(k))
            End If
        Next
    End If
Next

End Sub

