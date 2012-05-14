VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Role"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ICHCSII_Role

Private Const CLASS_NAME As String = "C2T_Security.Roles"
Private m_sName As String
Private m_colPrivileges As Collection

Friend Property Let Name(ByVal RoleName As String)
    m_sName = RoleName
End Property

Friend Property Get Name() As String
    Name = m_sName
End Property

Private Property Get ICHCSII_Role_Name() As String
    ICHCSII_Role_Name = m_sName
End Property

Friend Property Get Privileges() As Collection
    Set Privileges = m_colPrivileges
End Property

Friend Property Set Privileges(colPrivileges As Collection)
    Set m_colPrivileges = colPrivileges
End Property
