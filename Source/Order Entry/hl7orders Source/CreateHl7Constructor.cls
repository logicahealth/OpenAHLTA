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
Attribute VB_Name = "CreateHl7Constructor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'This Class has only one reason for being: to provide a client application
'with one, and only one, Hl7Constructor object.

Public Function GetHl7Constructor() As Hl7Constructor
     If goHl7Constructor Is Nothing Then
          'create a new hl7constructor, and return reference to it
          Set goHl7Constructor = New Hl7Constructor
          Set GetHl7Constructor = goHl7Constructor
     Else
          'return reference to already existing hl7constructor object
          Set GetHl7Constructor = goHl7Constructor
     End If
End Function

