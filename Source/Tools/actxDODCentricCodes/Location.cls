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
Attribute VB_Name = "Location"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Locations"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_lNCID As Long
Private m_sIEN As String
Private m_sName As String
Private m_sAbbrev As String

Public Property Get NCID() As Long
   NCID = m_lNCID
End Property

Public Property Let NCID(ByVal lNewValue As Long)
   m_lNCID = lNewValue
End Property

Public Property Get IEN() As String
   IEN = m_sIEN
End Property

Friend Property Let IEN(ByVal sNewValue As String)
   m_sIEN = sNewValue
End Property

Public Property Get Name() As String
    Name = m_sName
End Property

Friend Property Let Name(ByVal sNewValue As String)
    m_sName = sNewValue
End Property

Public Property Get Abbreviation() As String
   Abbreviation = m_sAbbrev
End Property

Public Property Let Abbreviation(ByVal sNewValue As String)
   m_sAbbrev = sNewValue
End Property
