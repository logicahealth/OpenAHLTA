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
Implements ICHCSII_User
Implements ITMIP_AppData

Private Const CLASS_NAME As String = "CHCSII_SecurityStub.User"

Private m_sUserName As String
Private m_sUserId As String
Private m_sPassword As String
Private m_objRoles As ICHCSII_Roles
Private m_objResources As ICHCSII_Resources
Private m_colPrivileges As New Collection
Private m_colAttributes As New Collection

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
End Property

Friend Property Set Roles(userRoles As ICHCSII_Roles)
    Set m_objRoles = userRoles
End Property

Friend Property Set Resources(userResources As ICHCSII_Resources)
    Set m_objResources = userResources
End Property

Private Function ICHCSII_User_CACPresent() As Boolean
    ICHCSII_User_CACPresent = False
End Function

Private Function ICHCSII_User_CACSignText(ByVal PIN As String, ByVal Text As String, Signature As String, Optional ErrCode As Long, Optional ErrMsg As String) As Boolean
    ICHCSII_User_CACSignText = False
End Function

Private Function ICHCSII_User_CACValidatePIN(ByVal PIN As String) As Boolean
    ICHCSII_User_CACValidatePIN = False
End Function

Private Function ICHCSII_User_GetAttribute(AttributeName As String) As String
    On Error Resume Next
    ICHCSII_User_GetAttribute = m_colAttributes.Item(AttributeName).Value
    If Err Then
        On Error GoTo 0
        Err.Raise SEC_NO_ATTRIBUTE_FOUND, CLASS_NAME, SEC_NO_ATTRIBUTE_FOUND_MSG
    End If
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
    On Error Resume Next
    Dim s As String
    s = m_colPrivileges.Item(sPrivilege)
    If Err Then
        Err.Clear
        ICHCSII_User_HasPrivilege = False
    Else
        ICHCSII_User_HasPrivilege = True
    End If
End Function

Private Function ICHCSII_User_HasPrivilegeEx(PrivilegeArea As String, PrivilegeType As CHCSII_PrivilegeType, Optional OtherType As Integer) As Boolean
    On Error Resume Next
    Dim s As String
    s = m_colPrivileges.Item(PrivilegeArea & PrivilegeType)
    If Err Then
        Err.Clear
        ICHCSII_User_HasPrivilegeEx = False
    Else
        ICHCSII_User_HasPrivilegeEx = True
    End If
End Function

Private Property Get ICHCSII_User_LoginMode() As CHCSII_ClientSecurity.CHCSII_LoginMode
    ICHCSII_User_LoginMode = Mode_Manual
End Property

Private Property Get ICHCSII_User_Resources() As CHCSII_ClientSecurity.ICHCSII_Resources
    Set ICHCSII_User_Resources = m_objResources
End Property

Private Property Get ICHCSII_User_Roles() As CHCSII_ClientSecurity.ICHCSII_Roles
    Set ICHCSII_User_Roles = m_objRoles
End Property

Private Property Get ICHCSII_User_UserId() As String
    ICHCSII_User_UserId = m_sUserId
End Property

Private Property Get ICHCSII_User_UserName() As String
    ICHCSII_User_UserName = m_sUserName
End Property

Private Function ICHCSII_User_ValidatePassword(Password As String) As Boolean
    If (Password = m_sPassword) Then
        ICHCSII_User_ValidatePassword = True
    Else
        ICHCSII_User_ValidatePassword = False
    End If
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
Dim l_Role As Role
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
        m_colPrivileges.Add colTemp.Item(j).Value, colTemp.Item(j).Value
        m_colPrivileges.Add colTemp.Item(j).Value & CStr(iPrivilegeTypeValues(1)), colTemp.Item(j).Value & CStr(iPrivilegeTypeValues(1))    'Default read privilege
    Else
        'Add specified privileges
        sPrivilegeName = colTemp.Item(j).Key
        sPrivilegeValue = colTemp.Item(j).Value
        
        'kluge multi-privilege per privilege area
        If InStr(1, sPrivilegeValue, "Z") > 0 Then
            sPrivilegeValue = "RWDSNBC"
        End If
        For k = 1 To UBound(sPrivilegeTypes)
            If InStr(1, sPrivilegeValue, sPrivilegeTypes(k)) > 0 Then
                m_colPrivileges.Add sPrivilegeName & CStr(iPrivilegeTypeValues(k)), sPrivilegeName & CStr(iPrivilegeTypeValues(k))
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


Private Function ITMIP_AppData_GetAppData(ByVal AppName As String, _
                                                UID As String, _
                                                PWD As String, _
                                                SID As String) As Boolean
'    Select Case AppName
'    Case "CHCSII-T"
'
'        UID = "gems"
'        PWD = "gems"
'        ITMIP_AppData_GetAppData = True
'    Case Else
'    End Select
End Function
