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
Attribute VB_Name = "DentalPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'''Public Sub Load_Patient_Encounter(ByVal lngPatientID As Long, ByVal lngEncounterId As Long)
'''    sqlStr = "Select SBE, SBE_COMMENTS from DENTAL_ENCOUNTER where DENTAL_ENCOUNTER_ID =" & DentalEncounterID
'''
'''End Sub

Private mobjSBE As SBE

Private Sub Class_Initialize()
    Set mobjSBE = New SBE
End Sub

Private Sub Class_Terminate()
    Set mobjSBE = Nothing
End Sub

Public Property Get SBE() As SBE
    Set SBE = mobjSBE
End Property

