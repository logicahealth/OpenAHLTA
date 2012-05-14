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
Attribute VB_Name = "Facility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private NCID As Long             ' Facility NCID
Private CNCID As Long            ' Context NCID
Private fType As String          ' Facility Type
Private Name As String           ' Name of facility

'Public Property Get Factype() As String
' Factype = fType
'End Property
'
'Public Property Let Factype(ByVal vNewValue As String)
' fType = vNewValue
'End Property
Public Property Get FacName() As String
 FacName = Name
End Property

Public Property Let FacName(ByVal vNewValue As String)
 Name = vNewValue
End Property


'Public Property Get FacContextNcid() As Long
'  FacContextNcid = CNCID
'End Property
'
'Public Property Let FacContextNcid(ByVal vNewValue As Long)
'  CNCID = vNewValue
'End Property
Public Property Get FacNcid() As Long
 FacNcid = NCID
End Property

Public Property Let FacNcid(ByVal vNewValue As Long)
 NCID = vNewValue
End Property

