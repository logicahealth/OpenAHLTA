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
Attribute VB_Name = "Problem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IPersist

Private mAcuity As String
Private mDescription As String
Private mICD9Code As String
Private mStatus As String

Private mValidator As Validation

Private Sub Class_Initialize()
  mStatus = "A"
  mAcuity = "A"
End Sub

''''''''''''''''''''''''''''''''''''''''''''''
' Persistence Block
''''''''''''''''''''''''''''''''''''''''''''''
Private Function IPersist_ChildFactory() As Object

End Function

Private Function IPersist_Persist() As String
    IPersist_Persist = Persist(Me)
End Function

Private Function IPersist_Restore(ByVal Persist As String) As Boolean
    IPersist_Restore = Restore(Me, Persist)
End Function

Public Property Get Acuity() As String
  Acuity = mAcuity
End Property

Public Property Let Acuity(ByVal NewValue As String)
  NewValue = UCase$(NewValue)
  
  mValidator.ValidatePropByList "Problem", "Acuity", NewValue, "A", "C"
  
  mAcuity = NewValue
End Property

Public Property Get Description() As String
  If mDescription = vbNullString Then
      mDescription = mValidator.GetDescription(mICD9Code, elICD9)
  End If
   
  Description = mDescription
End Property

Public Property Let Description(ByVal NewValue As String)
  mDescription = NewValue
End Property

Public Property Get ICD9Code() As String
  ICD9Code = mICD9Code
End Property

Public Property Let ICD9Code(ByVal NewValue As String)
  mICD9Code = NewValue
End Property

Public Property Get Status() As String
  Status = mStatus
End Property

Public Property Let Status(ByVal NewValue As String)
  NewValue = UCase$(NewValue)
  
  mValidator.ValidatePropByList "Problem", "Status", NewValue, "A", "I"

  mStatus = NewValue
End Property

Friend Sub Init(Validatator As Validation)
    Set mValidator = Validatator
End Sub

