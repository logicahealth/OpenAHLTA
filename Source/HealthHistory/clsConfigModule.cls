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
Attribute VB_Name = "clsConfigModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Public Module As clsSumOLEModule
Public ChkBox As Control
Public LblBox As Control

Public Property Get LblBoxLeft() As Single
  LblBoxLeft = LblBox.Left
End Property

Public Property Let LblBoxLeft(ByVal NewValue As Single)
  If NewValue >= 0 Then LblBox.Left = NewValue
End Property

Public Property Get LblBoxHeight() As Single
  LblBoxHeight = LblBox.Height
End Property

Public Property Let LblBoxHeight(ByVal NewValue As Single)
  If NewValue >= 0 Then LblBox.Height = NewValue
End Property

Public Property Get LblBoxTop() As Single
  LblBoxTop = LblBox.Top
End Property

Public Property Let LblBoxTop(ByVal NewValue As Single)
  If NewValue >= 0 Then LblBox.Top = NewValue
End Property

Public Property Get LblBoxWidth() As Single
  LblBoxWidth = LblBox.Width
End Property

Public Property Let LblBoxWidth(ByVal NewValue As Single)
  If NewValue >= 0 Then LblBox.Width = NewValue
End Property

Public Property Get ModuleSelected() As CheckBoxConstants
  ModuleSelected = ChkBox.Value
End Property

Public Property Let ModuleSelected(ByVal NewValue As CheckBoxConstants)
  ChkBox.Value = NewValue
End Property

