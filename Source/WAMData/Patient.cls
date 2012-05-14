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
Attribute VB_Name = "Patient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Implements IPersist

Public PATIENT_SSN As String
Public SPONSOR_SSN As String
Public PATIENT_IEN As String
Public FMP As String
Public DOB As Date
Public Sex As String
Public PATIENT_NAME As String

Private mvarValidator As Validation

Friend Sub Init(Validatator As Validation)
    Set mvarValidator = Validatator
End Sub




''''''''''''''''''''''''''''''''''''''''''''''
' Initialize/Terminate
''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Set mvarValidator = Nothing
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
