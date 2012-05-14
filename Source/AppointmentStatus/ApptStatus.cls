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
Attribute VB_Name = "ApptStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum ApptClassificationEnum
    ApptOutpatient = 0
    ApptInpatient = 1
    ApptAmbulatoryProcedure = 2
    ApptTelCon = 3
End Enum

Public Enum EncounterStatusEnum
    EncCheckedIn = 1
    EncWaiting = 2
    EncInProgress = 3
    EncCompleted = 4
    EncNeedsCosignature = 5
    EncUpdated = 6
    EncUpdating = 7
    EncCancel = 8
End Enum
'SCR- 26653 S.McAvoy 8/18/02 (added EncCancel)

Public Function SetEncounterStatus(ByVal ApptId As String, ByVal EncStatus As EncounterStatusEnum, Optional ByVal WorkloadCounts As Boolean = True)
    Dim objOps As Object
    Set objOps = CreateObject("CHCSII_AppointmentClient.AppointmentOps")
    SetEncounterStatus = objOps.SetEncounterStatus(ApptId, EncStatus, WorkloadCounts)
    Set objOps = Nothing
End Function

Public Function SetApptClassification(ByVal ApptId As String, ByVal Classification As ApptClassificationEnum) As Boolean
    Dim objOps As Object
    Set objOps = CreateObject("CHCSII_AppointmentClient.AppointmentOps")
    SetApptClassification = objOps.UpdateAppointmentClassification(ApptId, Classification)
    Set objOps = Nothing
End Function

