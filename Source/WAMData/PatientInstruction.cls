VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PatientInstruction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IPersist

Private mName As String
Private mSnoID As String
Private mICD As String

Private mvarValidator As Validation

Public Property Get ICD() As String
  ICD = mICD
End Property

Public Property Let ICD(ByVal NewValue As String)
  mICD = NewValue
End Property

Public Property Let SnoID(ByVal NewValue As String)
  mSnoID = NewValue
End Property

Public Property Get SnoID() As String
  SnoID = mSnoID
End Property

Public Property Let Name(ByVal NewValue As String)
  mName = NewValue
End Property

Public Property Get Name() As String
  Name = mName
End Property

Friend Sub Init(Validatator As Validation)
    Set mvarValidator = Validatator
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

