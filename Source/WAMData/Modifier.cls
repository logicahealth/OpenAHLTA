VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Modifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Private mvarCode As String 'local copy
Private mvarDescription As String 'local copy
Implements IPersist
'local variable(s) to hold property value(s)
Private mvarModifierType As EnumModifierType 'local copy
Private mvarValidator As Validation

Friend Sub Init(Validatator As Validation)
    Set mvarValidator = Validatator
End Sub
Public Property Let ModifierType(ByVal vData As EnumModifierType)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.ModifierType = Form1
    mvarModifierType = vData
End Property


Public Property Get ModifierType() As EnumModifierType
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ModifierType
    ModifierType = mvarModifierType
End Property
Public Property Get Description() As String
    Select Case Me.ModifierType
    Case mtCPT4
        Description = mvarValidator.GetDescription(Code, elModifiersCPT4)
    Case mtEandM
        Description = mvarValidator.GetDescription(Code, elModifiersEandM)
    End Select
End Property



Public Property Let Code(ByVal vData As String)
Attribute Code.VB_Description = "2 Characters. Validated against Modifiers Lookup Table"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Code = 5
'Displays 2 char code and description. Can have from zero to 3 modifiers.
'The modifiers further define the E&M code for billing.
'Valid values must be validated against a Modifiers Lookup table.
'Each Modifier must be unique for that E&M/CPT4
    
    Select Case Me.ModifierType
    Case mtCPT4
        mvarValidator.ValidateCode "Modifier", "CODE", vData, elModifiersCPT4
    Case mtEandM
        mvarValidator.ValidateCode "Modifier", "CODE", vData, elModifiersEandM
    End Select
        
    mvarCode = vData
    
End Property


Public Property Get Code() As String
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

Public Property Let Description(ByVal vData As String)
End Property

