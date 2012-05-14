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
Attribute VB_Name = "SingletonFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Property Get Singleton() As clsSVGComm
' See if it already exists.
    If g_Singleton Is Nothing Then
        Set g_Singleton = New clsSVGComm
    End If
    Set Singleton = g_Singleton
End Property
Public Sub TerminateSF()

'SF 5-17-02
    'Terminate the SVGComm server's instance of SVGManager
    If Not g_Singleton Is Nothing Then
        g_Singleton.TerminateObjects
    End If
    
    'Terminate the instance of the SVGComm server
    Set g_Singleton = Nothing
    
End Sub


