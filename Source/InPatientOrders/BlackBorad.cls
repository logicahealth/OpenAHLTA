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
Attribute VB_Name = "BlackBorad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mComm       As MMMHISComm.Comm
Private mConfig     As MMMHISConfiguration.Config
Private mLogon      As MMMHISLogon.Logon
Private mUser       As ICHCSII_User
Private moShared    As CWShared
Private mPatient    As MMMHISPatient.Patient
Private mAdmissionDate As DTOs.DocAdmissions

Private Sub Class_Terminate()
 
    Set mComm = Nothing
    Set mConfig = Nothing
    Set mLogon = Nothing
    Set mUser = Nothing
    Set moShared = Nothing
    
End Sub

Public Property Get Comm() As MMMHISComm.Comm
  Set Comm = mComm
End Property

Public Property Set Comm(ByRef NewValue As Object)
  Set mComm = NewValue
End Property

Public Property Get Config() As MMMHISConfiguration.Config
  Set Config = mConfig
End Property

Public Property Set Config(ByRef NewValue As Object)
  Set mConfig = NewValue
End Property

Public Property Get Logon() As MMMHISLogon.Logon
  Set Logon = mLogon
End Property

Public Property Set Logon(ByRef NewValue As Object)
  Set mLogon = NewValue
End Property

Public Property Get User() As ICHCSII_User
  Set User = mUser
End Property

Public Property Set User(ByRef NewValue As ICHCSII_User)
  Set mUser = NewValue
End Property

Public Property Get SharedObj() As CWShared
    If moShared Is Nothing Then Set moShared = New CWShared
    Set SharedObj = moShared
End Property

Public Sub ClearBlackBoard()
    Set moShared = Nothing
End Sub

Public Property Get Patient() As MMMHISPatient.Patient
  Set Patient = mPatient
End Property

Public Property Set Patient(ByRef NewValue As MMMHISPatient.Patient)
  Set mPatient = NewValue
End Property
   
Public Property Get AdmissionDate() As DTOs.DocAdmissions

    On Error GoTo ErrHandler

    If mAdmissionDate Is Nothing Then
    
        Set mAdmissionDate = goBB.SharedObj.DataAccessService.DocumentsMapper.RetrieveInpatientAdmissions( _
                    mPatient.UnitNumber, mLogon.FacilityNCID)
    End If
    Set AdmissionDate = mAdmissionDate
    
    Exit Property
    
ErrHandler:
    Call goBB.SharedObj.ShowVBError(Err.Number, Err.Description, ".GetAdmissionDate", "BlackBorad", vbExclamation)
End Property

'SCR 69248 - Need to reset the admission Date
Public Property Set AdmissionDate(ByRef NewValue As DTOs.DocAdmissions)
    Set mAdmissionDate = NewValue
End Property

