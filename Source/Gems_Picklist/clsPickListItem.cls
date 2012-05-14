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
Attribute VB_Name = "clsPickListItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private sNCID As String
Private sRepresentation As String
Public Property Get NCID() As String
    NCID = sNCID
End Property
Public Property Let NCID(ByVal sNCIDValue As String)
    sNCID = sNCIDValue
End Property
Public Property Get Representation() As String
    Representation = sRepresentation
End Property
Public Property Let Representation(ByVal sRepresentationValue As String)
    sRepresentation = sRepresentationValue
End Property
