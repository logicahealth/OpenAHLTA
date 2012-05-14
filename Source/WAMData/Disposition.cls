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
Attribute VB_Name = "Disposition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"AdminCodes"
Option Explicit
Option Compare Text

Implements IPersist

'local variable(s) to hold property value(s)
Private mvarWORK_RELATED As String 'local copy
Private mvarWORK_DATE As Date 'local copy
Private mvarDISP_CODE As String 'local copy
Private mvarDescription As String 'local copy
Private mvarAdminCodes As AdminCodes
Dim mvarCHEIF_COMPLAINT_CODE As String
Public DISP_NCID As String
Private mvarXMLData As String

Private mvarValidator As Validation

Public Property Get XMLData() As String
  XMLData = mvarXMLData
End Property

Public Property Let XMLData(ByVal vData As String)
  mvarXMLData = vData
End Property

Friend Sub Init(Validatator As Validation)
    Set mvarValidator = Validatator
    mvarAdminCodes.Init mvarValidator
End Sub

Public Property Get AdminCodes() As AdminCodes
    Set AdminCodes = mvarAdminCodes
End Property
Public Property Set AdminCodes(vData As AdminCodes)
    Set mvarAdminCodes = vData
End Property

Public Property Let Description(ByVal vData As String)
End Property
Public Property Get Description() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Description
    Description = mvarValidator.GetDescription(DISP_CODE, elDisposition)
End Property

Public Property Let DISP_CODE(ByVal vData As String)
Attribute DISP_CODE.VB_Description = "95 - Admitted\r\n96 - Expired\r\n90 - Immediate Referral\r\n06 - Left Against Medical Advice\r\n97 - Left without being seen\r\n85 - Released w/ work duty limitations\r\n91 - Released w/o limitations\r\n05 - Sick at home/quarters""\r\n"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DISP_CODE = 5
     '{{BOOKMARK Make this table driven later}}
    'mvarValidator.ValidatePropByList "Disposition", "DISP_CODE", vData, "", "95", "96", "90", "06", "97", "85", "91", "05"
    mvarDISP_CODE = vData
    
End Property
Public Property Get DISP_CODE() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DISP_CODE
    DISP_CODE = mvarDISP_CODE
End Property



Public Property Let WORK_DATE(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WORK_DATE = 5
    mvarWORK_DATE = vData
End Property
Public Property Get WORK_DATE() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WORK_DATE
    WORK_DATE = mvarWORK_DATE
End Property

Public Property Let WORK_RELATED(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WORK_FLG = 5
    If vData = "Nil" Then vData = "N"
    mvarValidator.ValidatePropByList "Disposition", "Work Related", vData, "Y", "N"
    mvarWORK_RELATED = vData
End Property
Public Property Get WORK_RELATED() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WORK_FLG
    WORK_RELATED = UCase(mvarWORK_RELATED)
End Property

Public Property Let CHEIF_COMPLAINT_CODE(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WORK_FLG = 5
    If Len(mvarCHEIF_COMPLAINT_CODE) Then
        mvarValidator.ValidateCode "Disposition", "CODE", vData, elICD9
    End If
    
    mvarCHEIF_COMPLAINT_CODE = vData
End Property
Public Property Get CHEIF_COMPLAINT_CODE() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WORK_FLG
    CHEIF_COMPLAINT_CODE = UCase(mvarCHEIF_COMPLAINT_CODE)
End Property

Public Property Let CHEIF_COMPLAINT_Description(ByVal vData As String)
End Property
Public Property Get CHEIF_COMPLAINT_Description() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Description
    If Len(mvarCHEIF_COMPLAINT_CODE) Then
        CHEIF_COMPLAINT_Description = mvarValidator.GetDescription(mvarCHEIF_COMPLAINT_CODE, elICD9)
    End If
End Property

''''''''''''''''''''''''''''''''''''''''''''''
' Initialize/Terminate
''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set mvarAdminCodes = New AdminCodes
End Sub

Private Sub Class_Terminate()
    Set mvarAdminCodes = Nothing
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

