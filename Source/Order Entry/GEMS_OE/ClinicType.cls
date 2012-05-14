VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClinicType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"ClinicCollection"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'local variable(s) to hold property value(s)
Private mvarClinic_IEN As String 'local copy
Private mvarMEPRS_Code As String 'local copy
Private mvarMEPRS_IEN As String 'local copy
Private mvarClinic_Name As String 'local copy

Public Property Let Clinic_Name(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Clinic_Name = 5
    mvarClinic_Name = vData
End Property


Public Property Get Clinic_Name() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Clinic_Name
    Clinic_Name = mvarClinic_Name
End Property



Public Property Let MEPRS_IEN(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MEPRS_IEN = 5
    mvarMEPRS_IEN = vData
End Property


Public Property Get MEPRS_IEN() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MEPRS_IEN
    MEPRS_IEN = mvarMEPRS_IEN
End Property



Public Property Let MEPRS_Code(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MEPRS_Code = 5
    mvarMEPRS_Code = vData
End Property


Public Property Get MEPRS_Code() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MEPRS_Code
    MEPRS_Code = mvarMEPRS_Code
End Property



Public Property Let Clinic_IEN(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Clinic_IEN = 5
    mvarClinic_IEN = vData
End Property


Public Property Get Clinic_IEN() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Clinic_IEN
    Clinic_IEN = mvarClinic_IEN
End Property



