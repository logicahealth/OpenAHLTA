VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RepStandard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private moCollRecSet As Collection
Private moCollFormulas As Collection

Private Function DataGetReadiness(sReportName As String, _
                                  oReportItem As ReportItem, _
                                  oReport As CRAXDRT.Report, _
                                  oCollFormulas As Collection) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler
DataGetReadiness = False

Dim sStmt As String
Dim i As Integer
Dim oRec As ADODB.Recordset
Dim oField As ADODB.Field

'-- Initialize recordset collection
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection
    
'-- Patient Information
'-----------------------------------------------------------------------
'-- Select
sStmt = "SELECT " & vbNewLine & _
    "mn.unit_number as Patient_Unit_Number," & vbNewLine & _
    "Sponsor.id_value as Sponsor_SSN," & vbNewLine & _
    "fmp.id_value as FMP," & vbNewLine & _
    "mn.first_name," & vbNewLine & _
    "mn.middle_name," & vbNewLine & _
    "mn.last_name" & vbNewLine

'-- From
sStmt = sStmt & "From" & vbNewLine & _
                    "mmi_generic_id Sponsor," & vbNewLine & _
                    "mmi_generic_id fmp," & vbNewLine & _
                    "mmi_name mn" & vbNewLine

'-- Where
sStmt = sStmt & _
    "Where" & vbNewLine & _
        "Sponsor.id_type_ncid=186334 and" & vbNewLine & _
        "Sponsor.unit_number(+)=mn.unit_number and" & vbNewLine & _
        "fmp.id_type_ncid=204708 and" & vbNewLine & _
        "fmp.unit_number(+)=mn.unit_number and" & vbNewLine & _
        "mn.name_type_cid=187 and" & vbNewLine & _
        "mn.unit_number in (" & oReportItem.CriteriaSel1.Item(1).UnitNumber & ")" & vbNewLine
        
'-- Order By
sStmt = sStmt & _
    "Order by last_name, first_name"


Set oRec = goDBDAS.OpenRecordset(sStmt)
DoEvents

'-- Set Data
'----------------------------------------------------------------------
    moCollRecSet.Add oRec, "patient"



'-- Readiness Data
'-----------------------------------------------------------------------

'-- Select
sStmt = "SELECT " & _
    "To_char(LAST_READINESS_DT, 'DD MON YYYY') as LAST_READINESS_DT," & vbNewLine & _
    "Patient_Unit_Number," & _
    "To_char(HIV_TEST_DT, 'DD MON YYYY') as HIV_TEST_DT," & vbNewLine & _
    "HIV_TEST_RSLT," & _
    "To_char(G6PD_TEST_DT, 'DD MON YYYY') as G6PD_TEST_DT," & vbNewLine & _
    "G6PD_TEST_RSLT," & _
    "To_char(SICKLE_CELL_TEST_DT, 'DD MON YYYY') as SICKLE_CELL_TEST_DT," & vbNewLine & _
    "SICKLE_CELL_TEST_RSLT," & _
    "To_char(BLOOD_TYPE_TEST_DT, 'DD MON YYYY') as BLOOD_TYPE_TEST_DT," & vbNewLine & _
    "BLOOD_TYPE_TEST_RSLT," & vbNewLine & _
    "To_char(DNA_TEST_DT, 'DD MON YYYY') as DNA_TEST_DT," & vbNewLine & _
    "To_char(HEAR_EXAM_DT, 'DD MON YYYY') as HEAR_EXAM_DT," & _
    "Decode(HEAR_AID_REQUIRED,'Y','Yes','N','No','') as HEAR_AID_REQUIRED," & vbNewLine & _
    "Decode(HEAR_AID_ISSUED,'Y','Yes','N','No','') as HEAR_AID_ISSUED," & vbNewLine & _
    "To_char(HEAR_AID_ISSUED_DT, 'DD MON YYYY') as HEAR_AID_ISSUED_DT," & vbNewLine & _
    "To_char(LAST_DENTAL_EXAM_DT, 'DD MON YYYY') as LAST_DENTAL_EXAM_DT," & vbNewLine & _
    "DENTAL_READINESS_CD," & vbNewLine & _
    "To_char(DENTAL_READINESS_DT, 'DD MON YYYY') as DENTAL_READINESS_DT," & vbNewLine & _
    "To_char(MEDIC_ASSESSMENT_DT, 'DD MON YYYY') as MEDIC_ASSESSMENT_DT," & vbNewLine & _
    "TEMP_PROFILE," & vbNewLine & _
    "To_char(TEMP_PROFILE_START_DT, 'DD MON YYYY') as TEMP_PROFILE_START_DT," & vbNewLine & _
    "To_char(TEMP_PROFILE_END_DT, 'DD MON YYYY') as TEMP_PROFILE_END_DT," & vbNewLine & _
    "TEMP_PROFILE_DIGNOS_CD, TEMP_PROFILE_DIGNOS_DSC, TEMP_PROFILE_LIMITATION, " & vbNewLine

sStmt = sStmt & _
    "To_char(LAST_VISUAL_ACUITY_DT, 'DD MON YYYY') as LAST_VISUAL_ACUITY_DT," & vbNewLine & _
    "UNCORRECTED_OD," & vbNewLine & _
    "CORRECTED_OD," & vbNewLine & _
    "UNCORRECTED_OS," & vbNewLine & _
    "CORRECTED_OS," & vbNewLine & _
    "UNCORRECTED_OU," & vbNewLine & _
    "CORRECTED_OU," & vbNewLine & _
    "PROTECT_INSERT_REQUIRED," & vbNewLine & _
    "PROTECT_INSERT_ISSUED," & vbNewLine & _
    "To_char(PROTECT_INSERT_EXAM_DT, 'DD MON YYYY') as PROTECT_INSERT_EXAM_DT," & vbNewLine & _
    "Encounter_Number, DNA_ON_FILE  " & vbNewLine
'    "Sponsor.id_value as Sponsor_SSN," & vbNewLine & _
'    "fmp.id_value as FMP," & vbNewLine & _
'    "mn.first_name," & vbNewLine & _
'    "mn.middle_name," & vbNewLine & _
'    "mn.last_name" & vbNewLine

'-- From
sStmt = sStmt & "From Readiness R" & vbNewLine

'                    "mmi_name mn," & vbNewLine & _
'                    "mmi_generic_id Sponsor," & vbNewLine & _
'                    "mmi_generic_id fmp," & vbNewLine

'-- Where
sStmt = sStmt & _
    "Where" & vbNewLine & _
        "LAST_READINESS_DT in" & vbNewLine & _
            "(select" & vbNewLine & _
                "MAX(r2.LAST_READINESS_DT)" & vbNewLine & _
            "from" & vbNewLine & _
                "Readiness r2" & vbNewLine & _
            "where" & vbNewLine & _
                "r2.patient_unit_number=r.patient_unit_number) and" & vbNewLine & _
        "Patient_Unit_Number in (" & oReportItem.CriteriaSel1.Item(1).UnitNumber & ")" & vbNewLine

'        "mn.name_type_cid=187 and" & vbNewLine & _
'        "mn.unit_number=r.patient_unit_number and" & vbNewLine & _
'        "Sponsor.id_type_ncid=186334 and" & vbNewLine & _
'        "Sponsor.unit_number=r.patient_unit_number and" & vbNewLine & _
'        "fmp.id_type_ncid=204708 and" & vbNewLine & _
'        "fmp.unit_number=r.patient_unit_number and" & vbNewLine & _

    '-- Order By
    sStmt = sStmt & _
        "Order by Patient_Unit_Number,R.LAST_READINESS_DT Desc"
    
    
    Set oRec = goDBDAS.OpenRecordset(sStmt)
    If Not (oRec.BOF And oRec.EOF) Then
        'Reset Null fields to an empty string to prevent the dreaded 12/31/9999 from showing up on the report
        oRec.MoveFirst
        For Each oField In oRec.Fields
            If IsDate(oField.Value) Then
                If Format(oField.Value, "mm/dd/yyyy") = "12/31/9999" Then oField.Value = ""
            End If
        Next
    End If
    DoEvents

'-- Set Data
'----------------------------------------------------------------------
    moCollRecSet.Add oRec, "readiness"

'-- Immunization Data
'-----------------------------------------------------------------------
sStmt = "select" & vbNewLine & _
    "r.representation as Type, " & vbNewLine & _
    "r2.representation as SeriesNumber," & vbNewLine & _
    "e.EVENT_START_GMTIME as AdminDate, " & vbNewLine & _
    "to_char(e.EVENT_START_GMTIME, 'DD MON YYYY') as AdminDateFormatted, " & vbNewLine & _
    "e.unit_number as Patient_Unit_Number " & vbNewLine & _
    "From" & vbNewLine & _
        "rsform_context rc2," & vbNewLine & _
        "rsform r2," & vbNewLine & _
        "observation_value ov2," & vbNewLine & _
        "clinical_observation co2," & vbNewLine & _
        "rsform_context rc," & vbNewLine & _
        "rsform r," & vbNewLine & _
        "observation_value ov," & vbNewLine & _
        "clinical_observation co," & vbNewLine & _
        "clinical_event_header ce," & vbNewLine & _
        "Event e" & vbNewLine
        
sStmt = sStmt & _
    "Where" & vbNewLine & _
        "rc2.preferred_score = 0 and" & vbNewLine & _
        "rc2.context_ncid = 2000 and" & vbNewLine & _
        "rc2.rsform_id = r2.rsform_id and" & vbNewLine & _
        "r2.ncid = ov2.value_ncid and" & vbNewLine & _
        "ov2.observation_seq = co2.observation_seq and" & vbNewLine & _
        "ov2.event_id = co2.event_id and" & vbNewLine & _
        "co2.observation_ncid = 154785 and" & vbNewLine & _
        "co2.event_id = e.event_id and " & vbNewLine & _
        "rc.preferred_score = 0 and " & vbNewLine & _
        "rc.context_ncid = 2000 and " & vbNewLine & _
        "rc.rsform_id = r.rsform_id and " & vbNewLine & _
        "r.ncid = ov.value_ncid and " & vbNewLine & _
        "ov.observation_seq = co.observation_seq and " & vbNewLine & _
        "ov.event_id = co.event_id and " & vbNewLine & _
        "co.observation_ncid = 154784 and " & vbNewLine & _
        "co.event_id = e.event_id and " & vbNewLine & _
        "ce.event_id = e.event_id and " & vbNewLine & _
        "e.event_subtype_ncid = 154789 and " & vbNewLine & _
        "e.event_type_ncid = 90782 and " & vbNewLine & _
        "e.Unit_Number in (" & oReportItem.CriteriaSel1.Item(1).UnitNumber & ")" & vbNewLine & _
    "order by 3,1,2"

Set oRec = goDBDAS.OpenRecordset(sStmt)
DoEvents

'-- Set Data
'----------------------------------------------------------------------
    moCollRecSet.Add oRec, "immunization"


'-- Create Report Object
'-----------------------------------------------------------------------------
    If Not CreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo CleanUp


    DoEvents

DataGetReadiness = True
GoTo CleanUp

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.DataGetReadiness", "CHCSII_Reports", vbExclamation)
CleanUp:
    Set oRec = Nothing
    Exit Function
End Function

Public Function BuildReport(sReportName As String, _
                            oReportItem As ReportItem, _
                            oReport As CRAXDRT.Report, _
                            oCollFormulas As Collection) As Boolean

On Error GoTo ErrHandler
 
BuildReport = False

DoEvents

'-- Get Data for the Report
'---------------------------------------------------------------------------------------
Select Case oReportItem.Report
    Case Readiness
        If Not DataGetReadiness(sReportName, oReportItem, oReport, oCollFormulas) Then GoTo CleanUp
    Case Else
End Select

BuildReport = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepRegistryCandidates.BuildReport", "CHCSII_Reports", vbExclamation
CleanUp:
    
End Function

Private Sub Class_Initialize()
    '--- SCR-ARM-16148 MSolano 9/10/01; COMMON TO MULTIPLE ARM-SCRs
    '--- Init routine should be called from first top level creatable coclass
    '--- It can be called multiple times, but will actually run only once.
    '--- Instrument all creatable classes in case dll's multiple entry
    '--- points are actually used
     Call ARM_InitAllTransactionsForCHCSIIReports

End Sub

