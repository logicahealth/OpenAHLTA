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
Attribute VB_Name = "AppLockProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event CloseApp()
Public Event Minimize()
Public Event Restore()
Public Event Validate(ByVal Password As String, ByRef IsValid As Boolean)

Public Sub CloseApp()
  RaiseEvent CloseApp
End Sub

Public Sub Minimize()
  RaiseEvent Minimize
End Sub

Public Sub Restore()
  RaiseEvent Restore
End Sub

Public Function Validate(ByVal Password As String) As Boolean
Dim Valid As Boolean

  RaiseEvent Validate(Password, Valid)
  Validate = Valid
  
End Function

