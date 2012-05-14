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
Attribute VB_Name = "OE_GetSessionManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'This class is here to provide a client application with a reference
'to the OE_SessionManager object. If one exists it returns reference
'to that, otherwise it creates one.


Public Function GetSessionManager() As OE_SessionManager
     If goOESession Is Nothing Then
          'create a new OE_SessionManager, and return reference to it
          Set goOESession = New OE_SessionManager
          Set GetSessionManager = goOESession
     Else
          'return reference to already existing OE_SessionManager object
          Set GetSessionManager = goOESession
     End If
     
End Function

Private Sub Class_Initialize()
    If gobjShared Is Nothing Then Set gobjShared = New CWShared
End Sub
