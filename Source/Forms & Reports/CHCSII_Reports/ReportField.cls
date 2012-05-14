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
Attribute VB_Name = "ReportField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private msFieldName
Private msFieldValue

Public Property Get FieldName() As String
    FieldName = msFieldName
End Property

Public Property Let FieldName(ByVal sNewValue As String)
    msFieldName = sNewValue
End Property

Public Property Get FieldValue() As String
    FieldValue = msFieldValue
End Property

Public Property Let FieldValue(ByVal sNewValue As String)
    msFieldValue = sNewValue
End Property
