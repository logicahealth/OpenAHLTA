VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdminCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Implements IPersist

'local variable(s) to hold property value(s)
Private mvarCode As Long 'local copy
Private mvarDescription As String 'local copy
Private mvarValidator As Validation

Friend Sub Init(Validatator As Validation)
    Set mvarValidator = Validatator
End Sub
Public Property Let Description(ByVal vData As String)
End Property


Public Property Get Description() As String
    Description = mvarValidator.GetDescription(Code, elAdminCode)
End Property



Public Property Let Code(ByVal vData As Long)
Attribute Code.VB_Description = "2 Characters. Validated against Modifiers Lookup Table"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Code = 5
    mvarValidator.ValidatePropByIntegerRange "AdminCode", "Code", vData, 1, 5
    mvarCode = vData
End Property


Public Property Get Code() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Code
    Code = mvarCode
End Property
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

