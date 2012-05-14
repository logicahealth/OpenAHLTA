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
Attribute VB_Name = "ProviderItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private msName As String
Private msNCID As String
Private msUnitNum As String

Public Property Get Name() As String
    Name = msName
End Property

Public Property Let Name(ByVal sNewValue As String)
    msName = sNewValue
End Property

Public Property Get NCID() As String
    NCID = msNCID
End Property

Public Property Let NCID(ByVal sNewValue As String)
    msNCID = sNewValue
End Property

Public Property Get UnitNum() As String
    UnitNum = msUnitNum
End Property

Public Property Let UnitNum(ByVal sNewValue As String)
    msUnitNum = sNewValue
End Property



