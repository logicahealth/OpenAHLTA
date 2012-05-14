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
Attribute VB_Name = "CPT4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Modifiers"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
Implements IPersist

'local variable(s) to hold property value(s)
Private mvarDescription As String 'local copy
Private mvarModifiers As Modifiers 'local copy
Private mvarCPT4_CODE As String 'local copy
Private mvarCPT4_DIAGNOSIS_PRIORITY As String 'local copy
Private mvarCPT4_UNITS_OF_SERVICE As Long 'local copy
Private mvarValidator As Validation
Private mvarSnoID As String
Private mvarComment As String
Private mvarICD As String
Private mvarProviders As String

Friend Sub Init(Validatator As Validation)
    Set mvarValidator = Validatator
    Set mvarModifiers = New Modifiers
    mvarModifiers.Init mtCPT4, mvarValidator
End Sub

Public Property Get ICD() As String
  ICD = mvarICD
End Property

Public Property Let ICD(ByVal NewValue As String)
  mvarICD = NewValue
End Property

Public Property Let Comment(ByRef vData As String)

    mvarValidator.ValidatePropByLength "ICD9", "Note", vData, 4000
    mvarComment = vData

End Property

Public Property Get Comment() As String
    Comment = mvarComment
End Property

Public Property Let CPT4_UNITS_OF_SERVICE(ByVal vData As Long)
Attribute CPT4_UNITS_OF_SERVICE.VB_Description = "3 digit Numeric, 1-999\r\nDefault = 1\r\n"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CPT4_UNITS_OF_SERVICE = 5
    
    mvarValidator.ValidatePropByIntegerRange "CPT4", "CPT4_UNITS_OF_SERVICE", vData, 1, 999
    mvarCPT4_UNITS_OF_SERVICE = vData
    
End Property
Public Property Get CPT4_UNITS_OF_SERVICE() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CPT4_UNITS_OF_SERVICE
    CPT4_UNITS_OF_SERVICE = mvarCPT4_UNITS_OF_SERVICE
End Property

Public Property Let CPT4_DIAGNOSIS_PRIORITY(ByVal vData As String)
Attribute CPT4_DIAGNOSIS_PRIORITY.VB_Description = "4-char numeric. \r\nRequired (at least one diagnosis pointer required per E&M code)  Only Allow chars 1,2,3,4. 1,2,3,4 are only allowed to be used once. Only can use value if Dx_Priority is in Dx collection - ie Dx_Priorities are 1,2,3, then value can be some combination of 1,2,3.\r\n"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CPT4_DIAGNOSIS_PRIORITY = 5
    mvarValidator.ValidateDIAGNOSIS_PRIORITY "CPT4", "CPT4_DIAGNOSIS_PRIORITY", vData
    mvarCPT4_DIAGNOSIS_PRIORITY = vData
    
End Property
Public Property Get CPT4_DIAGNOSIS_PRIORITY() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CPT4_DIAGNOSIS_PRIORITY
    CPT4_DIAGNOSIS_PRIORITY = mvarCPT4_DIAGNOSIS_PRIORITY
End Property

Public Property Let CPT4_CODE(ByVal vData As String)
Attribute CPT4_CODE.VB_Description = "5-digit numeric; Optional (up to 15 allowed)\r\nUse Medicomp lookup. Validate against CHCS tables\r\n"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CPT4_CODE = 5
    mvarValidator.ValidateCode "CPT4", "CPT4_CODE", vData, elCPT4
    mvarCPT4_CODE = vData
    
End Property
Public Property Get CPT4_CODE() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CPT4_CODE
    CPT4_CODE = mvarCPT4_CODE
End Property

Public Property Set Modifiers(ByVal vData As Modifiers)
Attribute Modifiers.VB_Description = "2 char code and description. Can have from zero to 3 modifiers. The modifiers further define the E&M code for billing. Valid values must be validated against a Modifiers Lookup table. Each Modifier must be unique for that E&M\r\n"
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Modifiers = Form1
    Set mvarModifiers = vData
End Property
Public Property Get Modifiers() As Modifiers
    Set Modifiers = mvarModifiers
End Property

Public Property Let Description(ByVal vData As String)
    mvarDescription = vData
End Property
Public Property Get Description() As String

    If mvarDescription = vbNullString Then
        mvarDescription = mvarValidator.GetDescription(CPT4_CODE, elCPT4)
    End If
    
    Description = mvarDescription
    
End Property

Public Property Get SnoID() As String
  SnoID = mvarSnoID
End Property

Public Property Let SnoID(ByVal vData As String)
  mvarSnoID = vData
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

Public Property Get Providers() As String
   'Indicates the number of the provider this CPT is associated with "123"
   Providers = mvarProviders
End Property

Public Property Let Providers(ByVal sData As String)
   mvarProviders = sData
End Property
