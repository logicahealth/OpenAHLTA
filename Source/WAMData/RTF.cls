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
Attribute VB_Name = "RTF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IPersist
            
Private mRTF As String
Private mOnsetDate As String
Private mWrittenBy As String
Private mTitle As String
Private mVerb As String

Private Sub Class_Initialize()
  mVerb = "Written"
End Sub

Public Property Get OnsetDate() As String
  OnsetDate = mOnsetDate
End Property

Public Property Let OnsetDate(ByVal NewValue As String)
  mOnsetDate = NewValue
End Property

Public Property Get RTF() As String
  RTF = mRTF
End Property

Public Property Let RTF(ByVal NewValue As String)
  mRTF = NewValue
End Property

Public Property Get Title() As String
  Title = mTitle
End Property

Public Property Let Title(ByVal NewValue As String)
  mTitle = NewValue
End Property

Public Property Get Verb() As String
  Verb = mVerb
End Property

Public Property Let Verb(ByVal NewValue As String)
  mVerb = NewValue
End Property

Public Property Get WrittenBy() As String
  WrittenBy = mWrittenBy
End Property

Public Property Let WrittenBy(ByVal NewValue As String)
  mWrittenBy = NewValue
End Property

Private Function IPersist_ChildFactory() As Object
'
End Function

Private Function IPersist_Persist() As String
    IPersist_Persist = Persist(Me)
End Function

Private Function IPersist_Restore(ByVal Persist As String) As Boolean
    IPersist_Restore = Restore(Me, Persist)
End Function

