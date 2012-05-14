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
Attribute VB_Name = "FormAppointments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private moCollRecSet As Collection
Private moCollFormulas As Collection

Public Function BuildAppointmentsReport(oReport As CRAXDRT.Report, _
                                        sTitle As String, _
                                        colAppts As Collection) As Boolean
On Error GoTo ErrHandler
BuildAppointmentsReport = False

Dim oCrystal As Crystal
Dim oRec0 As ADODB.Recordset
Dim oRec As ADODB.Recordset
Dim oRec1 As ADODB.Recordset
Dim collDups As Collection
Dim i As Integer
Dim sStmt As String
Dim adFieldType      As ADODB.DataTypeEnum
Dim lLineProcessing As Long
Dim sLineProcessing As String
Dim vWords As Variant
Dim vWord As Variant

DoEvents

'-- Get Data
'----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection

    '-- Get Appointment Data
    '------------------------------------------------------------------------
   
   Set oRec1 = New ADODB.Recordset
   Call AppendsResultSetFields(oRec1)           'Set up and record set's fields
   Call BuildRSFromCollection(oRec1, colAppts)
   
   moCollRecSet.Add oRec1, "appointment"

'-- Set Formulas
'-----------------------------------------------------------------------------
    Set moCollFormulas = Nothing
    Set moCollFormulas = New Collection
    moCollFormulas.Add "'" & Format(Now, gsDefaultFormat & "  HHMM") & "'", "TodaysDateTime"
    'moCollFormulas.Add "'" & sTitle, "Title"
    
    If Len(sTitle) <= 254 Then
       moCollFormulas.Add "'" & Replace(sTitle, "'", "''") & "'", "Title"
    Else
        vWords = Split(sTitle, " ")
        lLineProcessing = 1
        For Each vWord In vWords
            If Len(sLineProcessing) + Len(vWord) + 1 > 254 Then 'Add one for space
               Select Case lLineProcessing
                    Case 1
                        moCollFormulas.Add "'" & Replace(sLineProcessing, "'", "''") & "'", "Title"
                    Case Else
                        moCollFormulas.Add "'" & Replace(sLineProcessing, "'", "''") & "'", "Title" & CStr(lLineProcessing)
                End Select
                sLineProcessing = vWord
                lLineProcessing = lLineProcessing + 1
            Else
                sLineProcessing = sLineProcessing & " " & vWord
            End If
        Next
        If Len(sLineProcessing) > 0 Then
            'Append remaining text
            Select Case lLineProcessing
                 Case 1
                     moCollFormulas.Add "'" & Replace(sLineProcessing, "'", "''") & "'", "Title"
                 Case Else
                     moCollFormulas.Add "'" & Replace(sLineProcessing, "'", "''") & "'", "Title" & CStr(lLineProcessing)
             End Select
        End If
    End If
    
'-- Create Report Object
'-----------------------------------------------------------------------------
    Set oCrystal = New Crystal
    Set oReport = Nothing
    If gobjshared.IsAppMode(modeCHCSI_GUI) Then
        If Not oCrystal.ReportOpen("AppointmentList_GUI.rpt", oReport) Then GoTo CleanUp
    Else
        If Not oCrystal.ReportOpen("AppointmentList.rpt", oReport) Then GoTo CleanUp
    End If
    If Not oCrystal.AssignRecordSets(moCollRecSet, oReport) Then GoTo CleanUp
    If Not oCrystal.AssignFormulas(moCollFormulas, oReport) Then GoTo CleanUp

BuildAppointmentsReport = True
GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "Appointments.BuildAppointments", "CHCSII_Forms", vbExclamation
    'Resume Next
CleanUp:
    Set oRec = Nothing
    Set oRec1 = Nothing
    Exit Function
End Function

Private Sub BuildRSFromCollection(rsReturn As ADODB.Recordset, Optional pInput As Collection = Nothing)
    Dim lItems As Long
    
    On Error Resume Next
    Set rsReturn.ActiveConnection = Nothing
    rsReturn.CursorLocation = adUseClient
    
    rsReturn.Open
    If pInput Is Nothing And rsReturn.RecordCount = 0 Then
        rsReturn.AddNew
        rsReturn("FACILITY_NCID") = ""
    ElseIf pInput.Count = 0 And rsReturn.RecordCount = 0 Then
        rsReturn.AddNew
        rsReturn("FACILITY_NCID") = ""
    Else
        For lItems = 1 To pInput.Count
            rsReturn.AddNew
            rsReturn("FACILITY_NCID") = pInput(lItems).FacilityNCID
            rsReturn("APPT_ID") = pInput(lItems).ApptId
            rsReturn("APPT_IEN") = pInput(lItems).ApptIEN
            rsReturn("APPT_TIME") = pInput(lItems).ApptDateTime
            rsReturn("CHECKIN_TIME") = pInput(lItems).CheckInTime
            rsReturn("CHECKOUT_TIME") = pInput(lItems).CheckOutTime
            rsReturn("APPT_COMMENT") = pInput(lItems).ApptComment
            rsReturn("APPT_REASON") = pInput(lItems).ApptReason
            rsReturn("APPT_TYPE") = pInput(lItems).ApptType
            rsReturn("APPT_CLASSIFICATION") = pInput(lItems).ApptClassification
            rsReturn("APPT_STATUS") = pInput(lItems).ApptStatus
            rsReturn("APPT_CANCEL_REASON") = pInput(lItems).ApptCancelReason
            rsReturn("ENCOUNTER_NUMBER") = pInput(lItems).EncounterNumber
            rsReturn("ENCOUNTER_STATUS") = "" 'pInput(lItems).EncounterNumber
            rsReturn("ADM_ENCOUNTER") = pInput(lItems).ADM_Encounter
            rsReturn("PATIENT_UNIT_NUMBER") = pInput(lItems).PatientUnitNumber
            rsReturn("ORDER_NUMBER") = pInput(lItems).OrderNumber
            rsReturn("CLINIC_NCID") = pInput(lItems).ClinicNCID
            rsReturn("CLINIC_NAME") = pInput(lItems).ClinicName
            rsReturn("CLINICIAN_NCID") = pInput(lItems).ClinicianNCID
            rsReturn("PROVIDER_NAME") = pInput(lItems).ProviderName
            rsReturn("MEPRS_CODE") = pInput(lItems).MeprsCode
            rsReturn("MEPRS_DESCRIPTION") = pInput(lItems).MeprsDescription
'            rsReturn("DATE_CREATED") = pInput(lItems)
'            rsReturn("DATE_MODIFIED") = pInput(lItems)
            rsReturn("PATIENT_NAME") = pInput(lItems).PatientName
            rsReturn("PATIENT_IEN") = pInput(lItems).PatientIEN
            rsReturn("PATIENT_SSN") = pInput(lItems).PatientSSN
            rsReturn("PATIENT_WORK_PHONE") = pInput(lItems).PatientWorkPhone
            rsReturn("PATIENT_HOME_PHONE") = pInput(lItems).PatientHomePhone
            rsReturn("PATIENT_FMP") = pInput(lItems).PatientFMP
            rsReturn("SPONSOR_SSN") = pInput(lItems).SponsorSSN
            rsReturn("Observation") = pInput(lItems).Observation
            rsReturn("INPATIENT_MEPRS_CODE") = pInput(lItems).Inpatient_MEPRS_Code
            rsReturn("WORKLOAD_TYPE") = pInput(lItems).Workload_Type
        Next
    End If
End Sub

Private Sub AppendsResultSetFields(rsReturn As ADODB.Recordset)
    rsReturn.CursorLocation = adUseClient
    With rsReturn.Fields
            .Append "FACILITY_NCID", adLongVarWChar, 255, 132
            .Append "APPT_ID", adLongVarWChar, 255, 132
            .Append "APPT_IEN", adLongVarWChar, 255, 132
            .Append "APPT_TIME", adLongVarWChar, 255, 132
            .Append "CHECKIN_TIME", adLongVarWChar, 255, 132
            .Append "CHECKOUT_TIME", adLongVarWChar, 255, 132
            .Append "APPT_COMMENT", adLongVarWChar, 255, 132
            .Append "APPT_REASON", adLongVarWChar, 255, 132
            .Append "APPT_TYPE", adLongVarWChar, 255, 132
            .Append "APPT_CLASSIFICATION", adLongVarWChar, 255, 132
            .Append "APPT_STATUS", adLongVarWChar, 255, 132
            .Append "APPT_CANCEL_REASON", adLongVarWChar, 255, 132
            .Append "ENCOUNTER_NUMBER", adLongVarWChar, 255, 132
            .Append "ENCOUNTER_STATUS", adLongVarWChar, 255, 132
            .Append "PATIENT_UNIT_NUMBER", adLongVarWChar, 255, 132
            .Append "ORDER_NUMBER", adLongVarWChar, 255, 132
            .Append "CLINIC_NCID", adLongVarWChar, 255, 132
            .Append "CLINIC_NAME", adLongVarWChar, 255, 132
            .Append "CLINICIAN_NCID", adLongVarWChar, 255, 132
            .Append "PROVIDER_NAME", adLongVarWChar, 255, 132
            .Append "MEPRS_CODE", adLongVarWChar, 255, 132
            .Append "MEPRS_DESCRIPTION", adLongVarWChar, 255, 132
            .Append "DATE_CREATED", adLongVarWChar, 255, 132
            .Append "DATE_MODIFIED", adLongVarWChar, 255, 132
            .Append "PATIENT_NAME", adLongVarWChar, 255, 132
            .Append "PATIENT_IEN", adLongVarWChar, 255, 132
            .Append "PATIENT_SSN", adLongVarWChar, 255, 132
            .Append "PATIENT_WORK_PHONE", adLongVarWChar, 255, 132
            .Append "PATIENT_HOME_PHONE", adLongVarWChar, 255, 132
            .Append "PATIENT_FMP", adLongVarWChar, 255, 132
            .Append "SPONSOR_SSN", adLongVarWChar, 255, 132
            .Append "ADM_ENCOUNTER", adLongVarWChar, 255, 132
            .Append "Observation", adLongVarWChar, 255, 132
            .Append "INPATIENT_MEPRS_CODE", adLongVarWChar, 255, 132
            .Append "WORKLOAD_TYPE", adLongVarWChar, 255, 132
    End With
End Sub

Private Sub Class_Terminate()
   Set moCollRecSet = Nothing
   Set moCollFormulas = Nothing
End Sub
