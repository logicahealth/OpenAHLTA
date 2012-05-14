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

Private m_iCount As Integer
Private m_objRoles As New Collection


Friend Function AddRole(ByVal RoleName As String)
    On Error GoTo errhandler
    Dim objRole As New Role
    Dim MyCol As Collection
    Dim j As Long
    
    objRole.Name = RoleName
    m_objRoles.Add objRole, RoleName
    Exit Function
errhandler:
End Function

Private Property Get ICHCSII_Roles_Count() As Integer
    ICHCSII_Roles_Count = m_objRoles.Count
End Property

Friend Function Item(Index As Variant) As Role
    Set Item = m_objRoles.Item(Index)
End Function

Private Function ICHCSII_Roles_Item(Index As Variant) As CHCSII_ClientSecurity.ICHCSII_Role
    On Error GoTo errhandler
    Set ICHCSII_Roles_Item = m_objRoles.Item(Index)
    Exit Function
    
errhandler:
    'err.Raise SEC_NO_USER_ROLE, CLASS_NAME, SEC_NO_USER_ROLE_MSG
End Function

