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
Attribute VB_Name = "ICHCSII_Security"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"378B728F02C8"
Option Base 0
Option Explicit

'##ModelId=378B774D007F
Public Function GetActiveUser(Optional sFormTitle As String, Optional sWindowTitle As String) As ICHCSII_User
End Function

'##ModelId=378B776201A1
Public Function ChangePassword(User As ICHCSII_User) As Boolean
End Function

'##ModelId=378B777A020A
Public Function Logoff(User As ICHCSII_User) As Boolean
End Function

'##ModelId=378B77A90384
Public Function Logon(UserId As String, Password As String) As ICHCSII_User
End Function

'##ModelId=378B77FE01D8
Public Function GetPrivilegesByRole(User As ICHCSII_User, Role As ICHCSII_Role) As Collection
End Function

'##ModelId=378B99300368
Public Property Get Domain() As String
End Property

Public Function Activate(ActivateKey As String) As Boolean
End Function

Public Function GetUser(Optional sUserId As String) As ICHCSII_User
End Function

Public Function ValidateUser(ByVal UserId As String, ByVal Password As String, ByRef ErrNum As Long) As Boolean
End Function

'<< SCR 38721
Public Function TMIPAvailable() As Boolean
'__ used for Theater only.  True = a TMIP Framework is present and launched us.
End Function
