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
Attribute VB_Name = "iMapperPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'This should only be called from the DataAccess library
Public Sub Init(ByRef Support As iParameters, ByRef Cache As CacheMapper)
End Sub

'Gets Patient info
Public Function Retrieve(ByRef Params As DL_Support.iParameters) As DTOs.Patients
End Function

'Save the patient
Public Function Save(ByRef Patient As DTOs.Patient) As Boolean
End Function

'Gets value/name pairs for display in dropdown lists.  Currently only used in LocalDB modes.
Public Function RetrieveSelectStrings(ByRef Params As DL_Support.iParameters) As DTOs.PatientSelectStrings
End Function


'Transfer a patient from CHCS
Public Function TransferCHCSPatient(ByRef Params As DL_Support.iParameters) As DTOs.PatientTransferResult
End Function


'Return a Representation from a NCID
Public Function RetrieveRepresentation(ByRef Params As DL_Support.iParameters) As String
End Function


'Return a Rank NCID from a Rank Code
Public Function GetRankNCID(ByVal sRankCode As String, ByVal sBranch As String) As String
End Function

