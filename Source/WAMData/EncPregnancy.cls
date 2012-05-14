VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EncPregnancy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IPersist

'local variable(s) to hold property value(s)
Private mbPregnancyIndicator As Boolean
Private mdLastPeriod As Date
Private mdEstDOB As Date

'local variable to hold validation object
Private moValidator As Validation

Private Sub Class_Initialize()
   'Nothing to initialize
End Sub

Private Sub Class_Terminate()
   Set moValidator = Nothing
End Sub

Friend Sub Init(Validatator As Validation)
    Set moValidator = Validatator
End Sub

''''''''''''''''''''''''''''''''''''''''''''''
' Property Values
''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let PregnancyIndicator(ByVal bNewValue As Boolean)
   mbPregnancyIndicator = bNewValue
End Property

Public Property Get PregnancyIndicator() As Boolean
   PregnancyIndicator = mbPregnancyIndicator
End Property

Public Property Let LastMenstrualPeriod(ByVal dNewValue As Date)
   mdLastPeriod = dNewValue
End Property

Public Property Get LastMenstrualPeriod() As Date
   LastMenstrualPeriod = mdLastPeriod
End Property

Public Property Let EstimatedDOB(ByVal dNewValue As Date)
   mdEstDOB = dNewValue
End Property

Public Property Get EstimatedDOB() As Date
   EstimatedDOB = mdEstDOB
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

