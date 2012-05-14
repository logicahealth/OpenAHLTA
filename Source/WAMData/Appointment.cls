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
Attribute VB_Name = "Appointment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Contains information about the currently selected appoinment."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Implements IPersist

Public Check_In As Date
Public APPT_Date As Date
Public Clinic_Name As String
Public MEPRS_Code As String
Public APPT_Class As String
Public APPT_ID As Long
Public APPT_IEN As String
Public Encounter_Number As Long
Public APPT_Status As String
Public APPT_Reason As String
Public Provider_Name As String
Public MEPRS_IEN As String
'Public Provider_IEN As String

'local variable(s) to hold property value(s)
Private mvarProvider1_IEN As String
Private mvarProvider1_Role_Code As String
Private mvarProvider2_IEN As String 'local copy
Private mvarProvider2_Role_Code As String 'local copy
Private mvarProvider3_IEN As String 'local copy
Private mvarProvider3_Role_Code As String 'local copy
Private mvarAPPT_Type As String 'local copy

Private mvarValidator As Validation
'local variable(s) to hold property value(s)
Private mvarMEPRS_description As String 'local copy
Public Property Let MEPRS_description(ByVal vData As String)
End Property


Public Property Get MEPRS_description() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MEPRS_description
    MEPRS_description = mvarValidator.GetDescription(MEPRS_Code, elMEPRS)
End Property




Friend Sub Init(Validatator As Validation)
    Set mvarValidator = Validatator
End Sub

Public Property Let Appt_Type(ByVal vData As String) 'Max 5Chars
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.APPT_Type = 5
    Dim sz As String
    If Len(vData) > 5 Then
        sz = Left(vData, 5)
    Else
        sz = vData
    End If
    
    mvarValidator.ValidateCode "Appointment", "APPT_Type", sz, elApptType
    mvarAPPT_Type = sz
    
End Property
Public Property Get Appt_Type() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.APPT_Type
    Appt_Type = mvarAPPT_Type
End Property


Public Property Let Provider2_IEN(ByVal vData As String)
Attribute Provider2_IEN.VB_Description = "Optional, Validated by Provider Lookup table"
    
    If Len(vData) Then
        mvarValidator.ValidateCode "Appointment", "Provider2_IEN", vData, elProviderIEN
    End If
    
    mvarProvider2_IEN = vData
End Property
Public Property Get Provider2_IEN() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Provider2_IEN
    Provider2_IEN = mvarProvider2_IEN
End Property

Public Property Let Provider2_Role_Code(ByVal vData As String)
Attribute Provider2_Role_Code.VB_Description = "Required if Provider2_IEN is set. \r\nValid values:\r\n1-Attending\r\n2-Assisting\r\n3-Supervising\r\n4-Nurse\r\n5-Para-professional\r\n"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Provider2_Role_Code = 5
    If Len(Provider2_IEN) <> 0 Then
        mvarValidator.ValidatePropByList "Appointment", "Provider2_Role_Code", vData, "1", "2", "3", "4", "5", "6"
    End If
    
    mvarProvider2_Role_Code = vData
End Property
Public Property Get Provider2_Role_Code() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Provider2_Role_Code
    Provider2_Role_Code = mvarProvider2_Role_Code
End Property

Public Property Let Provider2_Role_Description(ByVal vData As String)
Attribute Provider2_Role_Description.VB_Description = "Decoded from Provider2_Role_Code: \r\n1-Attending\r\n2-Assisting\r\n3-Supervising\r\n4-Nurse\r\n5-Para-professional"
End Property
Public Property Get Provider2_Role_Description() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Provider2_Role_Description
    Provider2_Role_Description = mvarValidator.GetDescription(Provider2_Role_Code, elProviderRole)
End Property


Public Property Let Provider3_IEN(ByVal vData As String)
Attribute Provider3_IEN.VB_Description = "Optional, Validated by Provider Lookup table"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Provider3_IEN = 5
    If Len(vData) Then
        mvarValidator.ValidateCode "Appointment", "Provider3_IEN", vData, elProviderIEN
    End If
    
    mvarProvider3_IEN = vData
End Property
Public Property Get Provider3_IEN() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Provider3_IEN
    Provider3_IEN = mvarProvider3_IEN
End Property

Public Property Let Provider3_Role_Code(ByVal vData As String)
Attribute Provider3_Role_Code.VB_Description = "Required if Provider3_IEN is set. \r\nValid values:\r\n1-Attending\r\n2-Assisting\r\n3-Supervising\r\n4-Nurse\r\n5-Para-professional"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Provider3_Role_Code = 5
    If Len(Provider3_IEN) <> 0 Then
        mvarValidator.ValidatePropByList "Appointment", "Provider3_Role_Code", vData, "1", "2", "3", "4", "5", "6"
    End If
    mvarProvider3_Role_Code = vData
End Property
Public Property Get Provider3_Role_Code() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Provider3_Role_Code
    Provider3_Role_Code = mvarProvider3_Role_Code
End Property

Public Property Let Provider3_Role_Description(ByVal vData As String)
Attribute Provider3_Role_Description.VB_Description = "Decoded from Provider3_Role_Code: \r\n1-Attending\r\n2-Assisting\r\n3-Supervising\r\n4-Nurse\r\n5-Para-professional"
End Property
Public Property Get Provider3_Role_Description() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Provider3_Role_Description
    Provider3_Role_Description = mvarValidator.GetDescription(Provider3_Role_Code, elProviderRole)
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

Public Property Let Provider1_IEN(ByVal vData As String)
    If Len(vData) Then mvarValidator.ValidateCode "Appointment", "Provider1_IEN", vData, elProviderIEN
    mvarProvider1_IEN = vData
End Property
Public Property Get Provider1_IEN() As String
    Provider1_IEN = mvarProvider1_IEN
End Property

Public Property Let Provider1_Role_Code(ByVal vData As String)
   If Len(Provider1_IEN) <> 0 Then mvarValidator.ValidatePropByList "Appointment", "Provider1_Role_Code", vData, "1", "2", "3", "4", "5", "6"
   mvarProvider1_Role_Code = vData
End Property

Public Property Get Provider1_Role_Code() As String
    Provider1_Role_Code = mvarProvider1_Role_Code
End Property

Public Property Let Provider_IEN(ByVal vData As String)
    Provider1_IEN = vData
    Provider1_Role_Code = "1"    'SCR 50578 - set the default provider role for backwards compatibility
End Property
Public Property Get Provider_IEN() As String
    Provider_IEN = Provider1_IEN
End Property

