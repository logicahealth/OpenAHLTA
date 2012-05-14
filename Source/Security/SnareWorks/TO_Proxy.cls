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
Attribute VB_Name = "TO_Proxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event CardIn()
Public Event CardOut()
Private sLoginMode    As String
Private CardMon As cCardMonProxy

Friend Property Get LoginMode() As String
    LoginMode = sLoginMode
End Property

Friend Property Let LoginMode(rhv As String)
    sLoginMode = rhv
End Property

Friend Property Set CardMonitor(rhv As cCardMonProxy)
    Set CardMon = rhv
End Property

Friend Sub CardInserted()
    RaiseEvent CardIn
End Sub

Friend Sub CardRemoved()
    RaiseEvent CardOut
End Sub

Public Sub Terminate()
    CardMon.Terminate
End Sub
