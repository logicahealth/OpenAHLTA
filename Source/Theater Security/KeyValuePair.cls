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
Attribute VB_Name = "KeyValuePair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_sKey As String
Private m_sValue As Variant

Public Property Let Key(ByVal sKey As String)
    m_sKey = sKey
End Property

Public Property Let Value(ByVal sValue As Variant)
    If (IsObject(sValue)) Then
        Set m_sValue = sValue
    Else
        m_sValue = sValue
    End If
End Property

Public Property Get Key() As String
    Key = m_sKey
End Property

Public Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0
    If (IsObject(m_sValue)) Then
        Set Value = m_sValue
    Else
        Value = m_sValue
    End If
End Property

