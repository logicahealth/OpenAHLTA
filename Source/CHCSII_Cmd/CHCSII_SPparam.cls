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
Attribute VB_Name = "CHCSII_SPparam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Param_Name As String
Private Is_Cursor   As Long
Private Is_Output   As Long
Private Info_Type    As String


Public Property Get Param() As String
    Param = Param_Name
End Property
Public Property Let Param(ByVal vData As String)
    Param_Name = vData
End Property
Public Property Get IsCursor() As Long
    IsCursor = Is_Cursor
End Property
Public Property Let IsCursor(ByVal vData As Long)
    Is_Cursor = vData
End Property

Public Property Get IsOutput() As Long
    IsOutput = Is_Output
End Property
Public Property Let IsOutput(ByVal vData As Long)
    Is_Output = vData
End Property

Public Property Get InfoType() As String
    InfoType = Info_Type
End Property
Public Property Let InfoType(ByVal vData As String)
    Info_Type = vData
End Property

