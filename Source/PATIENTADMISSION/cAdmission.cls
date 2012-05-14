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
Attribute VB_Name = "cAdmission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarRegisterNumber As String 'local copy
Private mvarPatientIEN As String 'local copy
Private mvarInpatient_Meprs_Code As String 'local copy
Private mvarInpatient_Meprs_IEN As String 'local copy
Private mvarAdmission_Date As Date 'local copy
Private mvarAdmission_Type As String 'local copy
Private mvarAdmission_Code As String 'local copy
Private mvarAdmission_Synonym As String 'local copy
Private mvarDisposition_Type As String 'local copy
Private mvarDisposition_Code As String 'local copy
Private mvarDisposition_Synonym As String 'local copy
Private mvarAbsent As Boolean 'local copy
Private mvarDivision As String 'local copy
Private mvarDMIS_ID As String 'local copy
Private mvarMedicalHold_Disposition_Date As Date 'local copy
Private mvarDisposition_Date As Date 'local copy

Public Property Let Disposition_Date(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Disposition_Date = 5
    mvarDisposition_Date = vData
End Property


Public Property Get Disposition_Date() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Disposition_Date
     Disposition_Date = mvarDisposition_Date
End Property



Public Property Let MedicalHold_Disposition_Date(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MedicalHold_Disposition_Date = 5
    mvarMedicalHold_Disposition_Date = vData
End Property


Public Property Get MedicalHold_Disposition_Date() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MedicalHold_Disposition_Date
     MedicalHold_Disposition_Date = mvarMedicalHold_Disposition_Date
End Property



Public Property Let DMIS_ID(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DMIS_ID = 5
    mvarDMIS_ID = vData
End Property


Public Property Get DMIS_ID() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DMIS_ID
     DMIS_ID = mvarDMIS_ID
End Property



Public Property Let Division(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Division = 5
    mvarDivision = vData
End Property


Public Property Get Division() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Division
     Division = mvarDivision
End Property



Public Property Let Absent(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Absent = 5
    mvarAbsent = vData
End Property


Public Property Get Absent() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Absent
     Absent = mvarAbsent
End Property



Public Property Let Disposition_Synonym(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Disposition_Synonym = 5
    mvarDisposition_Synonym = vData
End Property


Public Property Get Disposition_Synonym() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Disposition_Synonym
     Disposition_Synonym = mvarDisposition_Synonym
End Property



Public Property Let Disposition_Code(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Disposition_Code = 5
    mvarDisposition_Code = vData
End Property


Public Property Get Disposition_Code() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Disposition_Code
     Disposition_Code = mvarDisposition_Code
End Property



Public Property Let Disposition_Type(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Disposition_Type = 5
    mvarDisposition_Type = vData
End Property


Public Property Get Disposition_Type() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Disposition_Type
     Disposition_Type = mvarDisposition_Type
End Property



Public Property Let Admission_Synonym(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Admission_Synonym = 5
    mvarAdmission_Synonym = vData
End Property


Public Property Get Admission_Synonym() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Admission_Synonym
     Admission_Synonym = mvarAdmission_Synonym
End Property



Public Property Let Admission_Code(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Admission_Code = 5
    mvarAdmission_Code = vData
End Property


Public Property Get Admission_Code() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Admission_Code
     Admission_Code = mvarAdmission_Code
End Property



Public Property Let Admission_Type(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Admission_Type = 5
    mvarAdmission_Type = vData
End Property


Public Property Get Admission_Type() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Admission_Type
     Admission_Type = mvarAdmission_Type
End Property



Public Property Let Admission_Date(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Admission_Date = 5
    mvarAdmission_Date = vData
End Property


Public Property Get Admission_Date() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Admission_Date
     Admission_Date = mvarAdmission_Date
End Property



Public Property Let Inpatient_Meprs_IEN(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Inpatient_Meprs_IEN = 5
    mvarInpatient_Meprs_IEN = vData
End Property


Public Property Get Inpatient_Meprs_IEN() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Inpatient_Meprs_IEN
     Inpatient_Meprs_IEN = mvarInpatient_Meprs_IEN
End Property



Public Property Let Inpatient_Meprs_Code(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Inpatient_Meprs_Code = 5
    mvarInpatient_Meprs_Code = vData
End Property


Public Property Get Inpatient_Meprs_Code() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Inpatient_Meprs_Code
     Inpatient_Meprs_Code = mvarInpatient_Meprs_Code
End Property






Public Property Let PatientIEN(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PatientIEN = 5
    mvarPatientIEN = vData
End Property


Public Property Get PatientIEN() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PatientIEN
    PatientIEN = mvarPatientIEN
End Property



Public Property Let RegisterNumber(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RegisterNumber = 5
    mvarRegisterNumber = vData
End Property


Public Property Get RegisterNumber() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RegisterNumber
    RegisterNumber = mvarRegisterNumber
End Property



