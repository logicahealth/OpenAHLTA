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
Attribute VB_Name = "ICD9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
Implements IPersist

'local variable(s) to hold property value(s)
Private mvarICD_CODE As String 'local copy
Private mvarICD_DIAGNOSIS_PRIORITY As String 'local copy
Private mvarACUITY As String 'local copy
Private mvarDATE_ONSET As Date 'local copy
Private mvarSTATUS As String 'local copy
Private mvarComment As String 'local copy
Private mvarDescription As String
Private mvarValidator As Validation
Private mvarSnoID As String
Private mvarPrefix As String
Private mvarDiagnosisType As String

Friend Sub Init(Validatator As Validation)
    Set mvarValidator = Validatator
    ICD_DIAGNOSIS_PRIORITY = "U"
End Sub

Public Property Get DiagnosisType() As String
  DiagnosisType = mvarDiagnosisType
End Property

Public Property Let DiagnosisType(ByVal NewValue As String)
  mvarDiagnosisType = NewValue
End Property

Public Property Let ACUITY(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ACUITY = 5

    'mvarValidator.ValidatePropByList "ICD9", "ACUITY", vData, "Acute", "Chronic"
    mvarACUITY = vData

End Property

Public Property Get ACUITY() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ACUITY
    ACUITY = mvarACUITY
End Property

Public Property Let ICD_DIAGNOSIS_PRIORITY(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ICD_DIAGNOSIS_PRIORITY = 5

    mvarValidator.ValidatePropByList "ICD", "ICD_DIAGNOSIS_PRIORITY", vData, "1", "2", "3", "4", "U"

    mvarICD_DIAGNOSIS_PRIORITY = vData
    
End Property
Public Property Get ICD_DIAGNOSIS_PRIORITY() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ICD_DIAGNOSIS_PRIORITY
    ICD_DIAGNOSIS_PRIORITY = mvarICD_DIAGNOSIS_PRIORITY
End Property


Public Property Let ICD_CODE(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ICD_CODE = 5
    mvarValidator.ValidateCode "ICD9", "ICD_CODE", vData, elICD9
    mvarICD_CODE = vData

End Property
Public Property Get ICD_CODE() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ICD_CODE
    ICD_CODE = mvarICD_CODE
End Property
Public Property Let Description(ByVal vData As String)
    mvarDescription = vData
End Property
Public Property Get Description() As String

    If mvarDescription = vbNullString Then
        mvarDescription = mvarValidator.GetDescription(ICD_CODE, elICD9)
    End If
    
    Description = mvarDescription
    
End Property

Public Property Get SnoID() As String
  SnoID = mvarSnoID
End Property

Public Property Let SnoID(ByVal vData As String)
  mvarSnoID = vData
End Property

Public Property Get Prefix() As String
  Prefix = mvarPrefix
End Property

Public Property Let Prefix(ByVal vData As String)
  mvarPrefix = vData
End Property

Public Property Let Comment(ByRef vData As String)

    mvarValidator.ValidatePropByLength "ICD9", "Note", vData, 4000
    mvarComment = vData

End Property

Public Property Get Comment() As String
    Comment = mvarComment
End Property

'''Public Property Let STATUS(ByVal vData As String)
''''used when assigning a value to the property, on the left side of an assignment.
''''Syntax: X.STATUS = 5
'''
'''    mvarValidator.ValidatePropByList "ICD9", "Status", vData, "active", "inactive"
'''    mvarSTATUS = vData
'''
'''End Property
'''Public Property Get STATUS() As String
''''used when retrieving value of a property, on the right side of an assignment.
''''Syntax: Debug.Print X.STATUS
'''    STATUS = mvarSTATUS
'''End Property
'''
'''Public Property Let DATE_ONSET(ByVal vData As Date)
''''used when assigning a value to the property, on the left side of an assignment.
''''Syntax: X.DATE_ONSET = 5
'''    mvarValidator.ValidateRequiredValue "ICD9", "DATE_ONSET", vData
'''    mvarDATE_ONSET = vData
'''End Property
'''Public Property Get DATE_ONSET() As Date
''''used when retrieving value of a property, on the right side of an assignment.
''''Syntax: Debug.Print X.DATE_ONSET
'''    DATE_ONSET = mvarDATE_ONSET
'''End Property
'''
'''
'''

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

