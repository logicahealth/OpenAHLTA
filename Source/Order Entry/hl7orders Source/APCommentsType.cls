VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "APCommentsType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarClinicalHistory As String 'local copy
Private mvarPreoperativeDiagnosis As String 'local copy
Private mvarOperativeFindings As String 'local copy
Private mvarPostoperativeDiagnosis As String 'local copy
Private mvarSpecimenCollectiondate As String

Public Property Let SpecimenCollectionDate(ByVal vData As String)
    mvarSpecimenCollectiondate = vData
End Property


Public Property Get SpecimenCollectionDate() As String
    SpecimenCollectionDate = mvarSpecimenCollectiondate
End Property
Public Property Let PostoperativeDiagnosis(ByVal vData As String)
    mvarPostoperativeDiagnosis = vData
End Property


Public Property Get PostoperativeDiagnosis() As String
    PostoperativeDiagnosis = mvarPostoperativeDiagnosis
End Property



Public Property Let OperativeFindings(ByVal vData As String)
    mvarOperativeFindings = vData
End Property


Public Property Get OperativeFindings() As String
   OperativeFindings = mvarOperativeFindings
End Property

Public Property Let PreoperativeDiagnosis(ByVal vData As String)
    mvarPreoperativeDiagnosis = vData
End Property


Public Property Get PreoperativeDiagnosis() As String
    PreoperativeDiagnosis = mvarPreoperativeDiagnosis
End Property

Public Property Let ClinicalHistory(ByVal vData As String)
    mvarClinicalHistory = vData
End Property

Public Property Get ClinicalHistory() As String
    ClinicalHistory = mvarClinicalHistory
End Property



