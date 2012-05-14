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
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public RoleName   As String
Public RoleDesc   As String
Public roleID     As Long
Public Privileges As Collection  '- of privilege
Public bChanged   As Boolean
Public bDeleted   As Boolean
Public bRedundant As Boolean   '- during ss import, this role is already there

Private Sub Class_Initialize()
    Set Privileges = New Collection
End Sub

Private Sub Class_Terminate()
    Set Privileges = Nothing
End Sub
