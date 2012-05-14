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
Attribute VB_Name = "MedWarnOverride"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mWarnings As colMedWarnings
Private mReason As String
Private mReasonId As EnumWarningReason
Private mComment As String

Public Property Get Warnings() As colMedWarnings
  Set Warnings = mWarnings
End Property
Public Property Set Warnings(oValue As colMedWarnings)
  Set mWarnings = oValue
End Property

Public Property Get Reason() As String
  Reason = mReason
End Property
Public Property Let Reason(sValue As String)
  mReason = sValue
End Property

Public Property Get ReasonId() As EnumWarningReason
  ReasonId = mReasonId
End Property
Public Property Let ReasonId(eValue As EnumWarningReason)
  mReasonId = eValue
End Property

Public Property Get Comment() As String
  Comment = mComment
End Property
Public Property Let Comment(sValue As String)
  mComment = sValue
End Property


Private Sub Class_Initialize()
  Set mWarnings = New colMedWarnings
End Sub

Private Sub Class_Terminate()
  Set mWarnings = Nothing
End Sub
