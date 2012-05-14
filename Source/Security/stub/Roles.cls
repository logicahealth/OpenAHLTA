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
Attribute VB_Name = "Roles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements ICHCSII_Roles

Private Const CLASS_NAME As String = "CHCSII_SecurityStub.Roles"
Private m_iCount As Integer
Private m_objRoles As New Collection

Friend Function AddRole(ByVal RoleName As String)
    Dim objRole As New Role
    objRole.Name = RoleName
    m_objRoles.Add objRole, RoleName
End Function

Private Property Get ICHCSII_Roles_Count() As Integer
    ICHCSII_Roles_Count = m_objRoles.Count
End Property

Friend Function Item(Index As Variant) As Role
    Set Item = m_objRoles.Item(Index)
End Function

Private Function ICHCSII_Roles_Item(Index As Variant) As CHCSII_ClientSecurity.ICHCSII_Role
    On Error Resume Next
    Set ICHCSII_Roles_Item = m_objRoles.Item(Index)
    If Err Then
        On Error GoTo 0
        Err.Raise SEC_NO_USER_ROLE, CLASS_NAME, SEC_NO_USER_ROLE_MSG
    End If
End Function

Friend Function Initialize(colRoles As Collection)

'Accepts a collection of roles as defined in the security_stub.txt file
Dim i As Integer
Dim objKeyValuePair As KeyValuePair
Dim objRole As Role

For i = 1 To colRoles.Count
    Set objKeyValuePair = colRoles.Item(i)
    Set objRole = New Role
    objRole.Name = objKeyValuePair.Key
    Set objRole.Privileges = GetPrivileges(objKeyValuePair.Value)
    m_objRoles.Add objRole, objRole.Name
Next

End Function

Private Function GetPrivileges(colPrivileges As Collection) As Collection

Dim colTemp As Collection
Dim colReturnPrivileges As New Collection
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
        colReturnPrivileges.Add colTemp.Item(j).Value, colTemp.Item(j).Value
        colReturnPrivileges.Add colTemp.Item(j).Value & CStr(iPrivilegeTypeValues(1)), colTemp.Item(j).Value & CStr(iPrivilegeTypeValues(1))    'Default read privilege
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
                colReturnPrivileges.Add sPrivilegeName & CStr(iPrivilegeTypeValues(k)), sPrivilegeName & CStr(iPrivilegeTypeValues(k))
            End If
        Next
    End If
Next

Set GetPrivileges = colReturnPrivileges

End Function

