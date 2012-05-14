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
Attribute VB_Name = "iMapperAppointments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Returns True if Patient currently has any Encounters that are open
Public Function HasOpenEncounters(PatientID As String) As Boolean
End Function

'Returns collection of appointments.  Criteria defined in DM_Appointments.AppointmentSearchParam
Public Function Retrieve(Params As iParameters) As DTOs.Appointments
End Function

'Returns name/value pairs of Appointment Types.  Theater use only
Public Function RetrieveApptTypes(FacilityID As String) As DTOs.ApptTypes
End Function

'Returns Patient/Inpatient Classification Data.
Public Function RetrieveClassification(PatientID As String) As DTOs.ApptClassification
End Function

'Returns Admissions data for the patient
Public Function RetrieveInpatientAdmissions(PatientID As String) As DTOs.InpatientAdmissions
End Function

'Save the appointment.  Op specifies what set of data has been modified.
Public Function SaveAppointment(Appt As DTOs.Appointment, Op As DTOs.ApptSaveOp) As Boolean
End Function

'Locks/unlocks the appointment from futher modification.  Used during checkin proecess
'to lock while encounter is created.  This functionality should eventually go away as the
'encounter creation process is made available through the data layer.
Public Function SetAppointmentLock(AppointmentLock As DTOs.apptlock) As DTOs.apptlock
End Function

