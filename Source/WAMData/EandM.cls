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
Attribute VB_Name = "EandM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"Modifiers"
Option Explicit
Implements IPersist

'local variable(s) to hold property value(s)
Private mvarEM_CODE As String 'local copy
Private mvarEM_UNITS_OF_SERVICE As String 'local copy
Private mvarEM_DIAGNOSIS_PRIORITY As String 'local copy
Private mvarDescription As String 'local copy
Private mvarModifiers As Modifiers
Private mvarValidator As Validation

Friend Sub Init(Validatator As Validation)
    Set mvarValidator = Validatator
    Set mvarModifiers = New Modifiers
    mvarModifiers.Init mtEandM, mvarValidator
    EM_UNITS_OF_SERVICE = "1"
End Sub


Public Property Get Modifiers() As Modifiers
Attribute Modifiers.VB_Description = "2 char code and description. Can have from zero to 3 modifiers. The modifiers further define the E&M code for billing. Valid values must be validated against a Modifiers Lookup table. Each Modifier must be unique for that E&M\r\n"

    Set Modifiers = mvarModifiers
End Property
Public Property Set Modifiers(vData As Modifiers)
    Set mvarModifiers = vData
End Property

Public Property Let Description(ByVal vData As String)
End Property
Public Property Get Description() As String
    Description = mvarValidator.GetDescription(EM_CODE, elEandM)
End Property


Public Property Let EM_DIAGNOSIS_PRIORITY(ByVal vData As String)
Attribute EM_DIAGNOSIS_PRIORITY.VB_Description = "4-char numeric. \r\nRequired (at least one diagnosis pointer required per E&M code)  Only Allow chars 1,2,3,4. 1,2,3,4 are only allowed to be used once. Only can use value if Dx_Priority is in Dx collection - ie Dx_Priorities are 1,2,3, then value can be some combination of 1,2,3.\r\n"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EM_DIAGNOSIS_PRIORITY4 = 5
    mvarValidator.ValidateDIAGNOSIS_PRIORITY "EandM", "EM_DIAGNOSIS_PRIORITY", vData
    mvarEM_DIAGNOSIS_PRIORITY = vData
    
'Required (at least one diagnosis pointer required per E&M code)
'Only Allow chars 1,2,3,4. 1,2,3,4 are only allowed to be used once.
'Only can use value if Dx_Priority is in Dx collection - ie Dx_Priorities are 1,2,3,
'then value can be some combination of 1,2,3.
    'Msgbox "Fix validation Logic, complex and invoves Diagnoses"
    
End Property


Public Property Get EM_DIAGNOSIS_PRIORITY() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EM_DIAGNOSIS_PRIORITY4
    EM_DIAGNOSIS_PRIORITY = mvarEM_DIAGNOSIS_PRIORITY
End Property

Public Property Let EM_UNITS_OF_SERVICE(ByVal vData As String)
Attribute EM_UNITS_OF_SERVICE.VB_Description = "1,2,3,4 is Valid\r\nDefaults to 1, can be overridden.\r\n"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EM_UNITS_OF_SERVICE = 5

    mvarValidator.ValidatePropByIntegerRange "EandM", "UNITS_OF_SERVICE", vData, 1, 4
    mvarEM_UNITS_OF_SERVICE = vData
    
End Property
Public Property Get EM_UNITS_OF_SERVICE() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EM_UNITS_OF_SERVICE
    EM_UNITS_OF_SERVICE = mvarEM_UNITS_OF_SERVICE
End Property

Public Property Let EM_CODE(ByVal vData As String)
Attribute EM_CODE.VB_Description = "5-digit numeric\r\nRequired (at least 1, multiples allowed under certain conditions) Multiple E&M codes are allowed under limited circumstances, including but not limited to encounters with excessive face time.\r\n"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EM_CODE = 5
    mvarValidator.ValidateCode "EandM", "EM_CODE", vData, elEandM
    mvarEM_CODE = vData

End Property
Public Property Get EM_CODE() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EM_CODE
    EM_CODE = mvarEM_CODE
End Property

''''''''''''''''''''''''''''''''''''''''''''''
' Initialize/Terminate
''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
        
        
End Sub

Private Sub Class_Terminate()
  Set mvarModifiers = Nothing
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

