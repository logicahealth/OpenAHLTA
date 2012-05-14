VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CauseCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IPersist

'local variable(s) to hold property value(s)
Private msCause_Code As String
Private msCause_Code_Description As String

Private moValidator As Validation

Friend Sub Init(Validatator As Validation)
    Set moValidator = Validatator
End Sub

Public Property Get Cause_Code() As String
   Cause_Code = msCause_Code
End Property

Public Property Let Cause_Code(ByVal sNewValue As String)
   moValidator.ValidatePropByList "CauseCode", "Cause_Code", sNewValue, "AA", "AP", "EM", "OA"
   msCause_Code = sNewValue
End Property

Public Property Let Cause_Code_Description(ByVal sNewValue As String)
    msCause_Code_Description = sNewValue
End Property

Public Property Get Cause_Code_Description() As String
    Cause_Code_Description = msCause_Code_Description
End Property

''''''''''''''''''''''''''''''''''''''''''''''
' Initialize/Terminate
''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
   'Nothing to initialize
End Sub

Private Sub Class_Terminate()
    Set moValidator = Nothing
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
