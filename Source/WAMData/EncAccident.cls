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
Attribute VB_Name = "EncAccident"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IPersist

'local variable(s) to hold property value(s)
Private mbAccidentIndicator As Boolean
Private mdAccidentDate As Date
'Private msCauseCode As String
Private mcCauseCodes As CauseCodes
Private msLocationIEN As String          'Number(20) - too big for long
Private msLocationCode As String          'Number(20) - too big for long

'local variable to hold validation object
Private moValidator As Validation

Private Sub Class_Initialize()
   Set mcCauseCodes = New CauseCodes
End Sub

Private Sub Class_Terminate()
   Set mcCauseCodes = Nothing
   Set moValidator = Nothing
End Sub

Friend Sub Init(Validatator As Validation)
    Set moValidator = Validatator
    mcCauseCodes.Init Validatator
End Sub

''''''''''''''''''''''''''''''''''''''''''''''
' Property Values
''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let AccidentIndicator(ByVal bNewValue As Boolean)
   mbAccidentIndicator = bNewValue
End Property

Public Property Get AccidentIndicator() As Boolean
   AccidentIndicator = mbAccidentIndicator
End Property

Public Property Let AccidentDate(ByVal dNewValue As Date)
   mdAccidentDate = dNewValue
End Property

Public Property Get AccidentDate() As Date
   AccidentDate = mdAccidentDate
End Property

Public Property Let LocationIEN(ByVal sNewValue As String)
   msLocationIEN = sNewValue
End Property

Public Property Get LocationIEN() As String
   LocationIEN = msLocationIEN
End Property

Public Property Let LocationCode(ByVal sNewValue As String)
   msLocationCode = sNewValue
End Property

Public Property Get LocationCode() As String
   LocationCode = msLocationCode
End Property

Public Property Get CauseCodes() As CauseCodes
   Set CauseCodes = mcCauseCodes
End Property

Public Property Set CauseCodes(ByVal cNewValue As CauseCodes)
   Set mcCauseCodes = cNewValue
End Property

Public Property Let CauseCode(ByVal sNewValue As String)
   'Left in for compatibility - Remove Property Later
   'msCauseCode = sNewValue
End Property

Public Property Get CauseCode() As String
   'Left in for compatibility - Remove Property Later
   'CauseCode = msCauseCode
End Property
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


