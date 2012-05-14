VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PatientType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Class1"
Option Explicit

Private mvarIen As String 'local copy
Private mvarName As String 'local copy
Private mvarAddress As String 'local copy
Private mvarSponsor_SSN As String 'local copy
Private mvarFMP As String 'local copy
Private mvarDOB As String 'local copy
Private mvarSex As String 'local copy
Private mvarHL7Name As String 'local copy
Public Function szCalcAge(szDOB As String) As String

Dim age As Variant
Dim szDatePartMonth As String
Dim szDatePartDay As String
Dim szDatePartYear As String
    
      If LenB(szDOB) > 0 Then
        szDatePartMonth = Mid$(szDOB, 5, 2)
        szDatePartDay = Mid$(szDOB, 7, 2)
        szDatePartYear = Left$(szDOB, 4)
        szDOB = szDatePartMonth & "/" & szDatePartDay & "/" & szDatePartYear
        If szDatePartMonth = "00" Or szDatePartDay = "00" Or szDatePartYear = "0000" Then
          age = "   "
        Else
          age = Fix(DateDiff("m", szDOB, Now) / 12)
          If age < 3 Then
            age = DateDiff("m", szDOB, Now) & " months "
          End If
        End If
      Else
        age = "   "
      End If
      
      szCalcAge = age
    
End Function
Friend Property Let HL7Name(ByVal vData As String)
    mvarHL7Name = vData
End Property

Public Property Get HL7Name() As String
    HL7Name = mvarHL7Name
End Property

Friend Property Let Sex(ByVal vData As String)
    mvarSex = vData
End Property

Public Property Get Sex() As String
    Sex = mvarSex
End Property

Friend Property Let DOB(ByVal vData As String)
    mvarDOB = vData
End Property

Public Property Get DOB() As String
    DOB = mvarDOB
End Property

Friend Property Let Fmp(ByVal vData As String)
    mvarFMP = vData
End Property

Public Property Get Fmp() As String
    Fmp = mvarFMP
End Property

Friend Property Let Sponsor_SSN(ByVal vData As String)
    mvarSponsor_SSN = vData
End Property

Public Property Get Sponsor_SSN() As String
    Sponsor_SSN = mvarSponsor_SSN
End Property

Friend Property Let Address(ByVal vData As String)
    mvarAddress = vData
End Property

Public Property Get Address() As String
    Address = mvarAddress
End Property

Friend Property Let Name(ByVal vData As String)
    mvarName = vData
End Property

Public Property Get Name() As String
    Name = mvarName
End Property

Friend Property Let IEN(ByVal vData As String)
    mvarIen = vData
End Property

Public Property Get IEN() As String
Attribute IEN.VB_Description = "Patient IEN"
Attribute IEN.VB_UserMemId = 0
    IEN = mvarIen
End Property

