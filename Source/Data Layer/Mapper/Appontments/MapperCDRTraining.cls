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
Attribute VB_Name = "MapperCDRTraining"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements DL_Interfaces.iMapperAppointments
Private Const msMODULE_NAME         As String = "DM_Appointments.MapperCDRTraining."

Private mobjMapperGeneric           As DL_Interfaces.iMapperAppointments


Private Sub Class_Initialize()
On Error GoTo errHandler
Dim objMapper As DM_Appointments.MapperGeneric
Const ProcName As String = "Class_Initialize"


  Set mobjMapperGeneric = New MapperGeneric
  Set objMapper = mobjMapperGeneric
  Exit Sub
  
errHandler:
  Err.Raise Err.Number, msMODULE_NAME + ProcName, Err.Description, Err.HelpFile, Err.HelpContext
End Sub
Private Function iMapperAppointments_HasOpenEncounters(PatientID As String) As Boolean
Const ProcName As String = "iMapperAppointments_HasOpenEncounters"


Exit Function
errHandler:
  Err.Raise Err.Number, msMODULE_NAME + ProcName, Err.Description, Err.HelpFile, Err.HelpContext
End Function

Private Function iMapperAppointments_Retrieve(Params As DL_Support.iParameters) As DTOs.Appointments
Const ProcName As String = "iMapperAppointments_Retrieve"

Exit Function
errHandler:
  Err.Raise Err.Number, msMODULE_NAME + ProcName, Err.Description, Err.HelpFile, Err.HelpContext
End Function

Private Function iMapperAppointments_RetrieveApptTypes(FacilityID As String) As DTOs.ApptTypes
Const ProcName As String = "iMapperAppointments_RetrieveApptTypes"

Exit Function
errHandler:
  Err.Raise Err.Number, msMODULE_NAME + ProcName, Err.Description, Err.HelpFile, Err.HelpContext
End Function

Private Function iMapperAppointments_RetrieveClassification(PatientID As String) As DTOs.ApptClassification
Const ProcName As String = "iMapperAppointments_RetrieveClassification"

Exit Function
errHandler:
  Err.Raise Err.Number, msMODULE_NAME + ProcName, Err.Description, Err.HelpFile, Err.HelpContext
End Function

Private Function iMapperAppointments_RetrieveInpatientAdmissions(PatientID As String) As DTOs.InpatientAdmissions
Const ProcName As String = "iMapperAppointments_RetrieveInpatientAdmissions"

Exit Function
errHandler:
  Err.Raise Err.Number, msMODULE_NAME + ProcName, Err.Description, Err.HelpFile, Err.HelpContext
End Function

Private Function iMapperAppointments_SaveAppointment(Appt As DTOs.Appointment, Op As DTOs.ApptSaveOp) As Boolean
Const ProcName As String = "iMapperAppointments_SaveAppointment"

End Function

Private Function iMapperAppointments_SetAppointmentLock(AppointmentLock As DTOs.ApptLock) As DTOs.ApptLock
Const ProcName As String = "iMapperAppointments_SetAppointmentLock"

End Function

