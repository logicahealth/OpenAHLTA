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
Attribute VB_Name = "cWard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarWardName        As String
Private mvarWardNCID        As String

Public Property Let WardName(ByVal vData As String)
    mvarWardName = vData
End Property
Public Property Get WardName() As String
    WardName = mvarWardName
End Property
Public Property Let WardNCID(ByVal vData As String)
    mvarWardNCID = vData
End Property
Public Property Get WardNCID() As String
    WardNCID = mvarWardNCID
End Property

