VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RepCPGRegistryCandidate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private moCollRecSet As Collection
Private moCollFormulas As Collection

Public Function BuildReport(sReportName As String, _
                            oReportItem As ReportItem, _
                            oReport As CRAXDRT.Report, _
                            oCollFormulas As Collection) As Boolean

On Error GoTo ErrHandler
 
BuildReport = False

Dim oCrystal As Crystal
Dim oRec As ADODB.Recordset
Dim sSponsorUnitNumber As String
Dim sUnitNumber As String
Dim sPatientList As String
DoEvents

'-- Get Data
'----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection
    
    '-- Get Candidates
    '----------------------------------------------------------
    If Not GetCandidates(oReportItem, oRec) Then GoTo CleanUp
    moCollRecSet.Add oRec, "candidates"

    '-- Build Patient List
    '-----------------------------------------------------------
    If Not GetPatientList(oRec, sPatientList) Then GoTo CleanUp

    '-- Get Ambulatory Visits
    '----------------------------------------------------------
    If Not GetAmbulatoryVisits(oReportItem, sPatientList, oRec) Then GoTo CleanUp
    moCollRecSet.Add oRec, "diagnosis"
    
    '-- Get Medications
    '----------------------------------------------------------
    If Not GetMedications(oReportItem, sPatientList, oRec) Then GoTo CleanUp
    moCollRecSet.Add oRec, "medications"

'-- Create Report Object
'-----------------------------------------------------------------------------
    If Not CreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo CleanUp

BuildReport = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepRegistryCandidates.BuildReport", "CHCSII_Reports", vbExclamation
CleanUp:
    Set oCrystal = Nothing
    Set oRec = Nothing
    
End Function

Private Function GetCandidates(oReportItem As ReportItem, _
                               oRec As ADODB.Recordset) As Boolean
    
    On Error GoTo ErrHandler
    
    GetCandidates = False
    
    Dim sStmt As String
    Dim sFromSQL As String
    Dim sWhereSQL As String
    Dim sPatientList As String

'-- Get patients not in registry that have a diagnosis of high blood pressure
'----------------------------------------------------------------------------

    '-- Select
    sStmt = _
            "SELECT" & vbNewLine & _
                "Distinct Patient_Unit_Number" & vbNewLine
    
    '-- From
    sStmt = sStmt & _
            "From" & vbNewLine & _
                "APPOINTMENT a," & vbNewLine & _
                "EVENT E," & vbNewLine & _
                "OBSERVATION_VALUE OV" & vbNewLine
                
    '-- Where
    sStmt = sStmt & _
            "Where" & vbNewLine & _
                "NOT EXISTS" & vbNewLine & _
                    "(SELECT 'x' FROM cpgregistry c WHERE c.unit_number = e.unit_number) and" & vbNewLine & _
                "OV.VALUE_NCID IN" & vbNewLine & _
                    "(SELECT 1002435" & vbNewLine & _
                    "From DUAL" & vbNewLine & _
                    "Union" & vbNewLine & _
                    "SELECT ncid" & vbNewLine & _
                    "FROM rsform_context rc, rsform r" & vbNewLine & _
                    "Where" & vbNewLine & _
                        "r.representation LIKE '250.%' AND" & vbNewLine & _
                        "rc.context_ncid = 1257 AND" & vbNewLine & _
                        "rc.rsform_id = r.rsform_id) AND" & vbNewLine & _
                "E.EVENT_ID=OV.EVENT_ID AND" & vbNewLine & _
                "e.EVENT_SUBTYPE_NCID = 658 AND" & vbNewLine & _
                "e.EVENT_TYPE_NCID = 90803 AND" & vbNewLine & _
                "e.encounter_number = a.encounter_number AND" & vbNewLine & _
                " " & _
                "a.encounter_number IS NOT NULL AND" & vbNewLine & _
                "a.encounter_number <> '0' AND" & vbNewLine & _
                "(a.APPT_DATE_TIME BETWEEN (SYSDATE - 365) AND SYSDATE)"

    Set oRec = goDBDAS.OpenRecordset(sStmt)

'-- Build list of patients found in previous query
'------------------------------------------------------------------

If oRec.RecordCount <> 0 Then

    sPatientList = ""
    oRec.MoveFirst
    Do While Not oRec.EOF
        If sPatientList <> "" Then
            sPatientList = sPatientList & ","
        End If
    
        sPatientList = sPatientList & oRec.Fields(0).Value
    
        oRec.MoveNext
    Loop

    Set oRec = Nothing

'-- Get PCM and other information for registry candidate patients
'------------------------------------------------------------------

'-- Select
    sStmt = _
            "SELECT" & vbNewLine & _
                "/*+ INDEX (id mmi_id_pk) (mn mmi_name_pk) (fs mmi_generic_id_x3) (pm mmi_generic_id_x3) (p provider_pk) */" & vbNewLine & _
                "mn.unit_number," & vbNewLine & _
                "Trim(mn.last_name ||', ' || mn.first_name || ' ' || mn.middle_name) AS Full_Name," & vbNewLine & _
                "fs.id_value AS FMP_SPONSOR_SSN," & vbNewLine & _
                "TO_CHAR(id.birth_date,'DD MON YYYY') AS DOB_FORMATTED," & vbNewLine & _
                "p.NAME AS PCM_Name" & vbNewLine & _
            "From" & vbNewLine & _
                "cdrplus.provider p," & vbNewLine & _
                "mmi.MMI_GENERIC_ID pm," & vbNewLine & _
                "mmi.MMI_GENERIC_ID fs," & vbNewLine & _
                "mmi.MMI_NAME mn," & vbNewLine & _
                "mmi.mmi_id Id" & vbNewLine & _
            "Where" & vbNewLine & _
                "p.facility_ncid = " & oReportItem.CriteriaSel1.Item(1).NCID & " and p.pcm_code = pm.id_value AND" & vbNewLine & _
                "pm.id_type_ncid = 14501724 AND" & vbNewLine & _
                "pm.unit_number = id.unit_number AND" & vbNewLine & _
                "fs.id_type_ncid = 14511309 AND" & vbNewLine & _
                "fs.unit_number = id.unit_number AND" & vbNewLine & _
                "mn.name_type_cid = 187 AND" & vbNewLine & _
                "mn.unit_number = id.unit_number AND" & vbNewLine & _
                "id.unit_number IN" & vbNewLine & _
                "(" & sPatientList & ")"

    Set oRec = goDBDAS.OpenRecordset(sStmt)

End If

    
    GetCandidates = True
    GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepRegistryCandidates.GetCandidates", "CHCSII_Reports", vbExclamation
CleanUp:
    Exit Function

End Function

Private Function GetPatientList(oRec As ADODB.Recordset, _
                                sPatientList As String) As Boolean
On Error GoTo ErrHandler

GetPatientList = False
sPatientList = "("

If oRec.RecordCount = 0 Then
    sPatientList = sPatientList & "NULL"
Else
    oRec.MoveFirst
    Do While Not oRec.EOF
        sPatientList = sPatientList & oRec.Fields(0).Value
        oRec.MoveNext
        If Not oRec.EOF Then
            sPatientList = sPatientList & ","
        End If
    Loop
End If

sPatientList = sPatientList & ")"

GetPatientList = True
GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepRegistryCandidates.GetPatientList", "CHCSII_Reports", vbExclamation
CleanUp:
    Exit Function
End Function

Private Function GetAmbulatoryVisits(oReportItem As ReportItem, _
                                     sPatientList As String, _
                                     oRec As ADODB.Recordset) As Boolean
On Error GoTo ErrHandler

GetAmbulatoryVisits = False

Dim sStmt As String
Dim sFromSQL As String
Dim sWhereSQL As String

'-- Select
    sStmt = "Select  /*+ INDEX (e EVENT_X3) */ distinct" & vbNewLine & _
                "a.patient_unit_number as unit_number, " & vbNewLine & _
                "d.description as Diagnosis, " & vbNewLine & _
                "d.icd9 as ICD9, " & vbNewLine & _
                "TO_CHAR(a.appt_date_time,'DD MON YYYY') as APPT_DATE_TIME " & vbNewLine

'-- From
    sStmt = sStmt & _
            "From " & vbNewLine & _
                "rpt_diagnosis d, " & vbNewLine & _
                "observation_value ov, " & vbNewLine & _
                "clinical_observation co, " & vbNewLine & _
                "event e, " & vbNewLine & _
                "appointment a " & vbNewLine

'-- Where
    '-- Ambulatory Appointment (Classification = 2)
    sStmt = sStmt & _
        "Where " & vbNewLine & _
            "d.ncid = ov.value_ncid and " & vbNewLine & _
            "ov.observation_seq = co.observation_seq AND ov.event_id = e.event_id and " & vbNewLine & _
            "co.event_id = e.event_id and " & vbNewLine & _
            "e.event_subtype_ncid = 658 and " & vbNewLine & _
            "e.event_type_ncid = 90803 and " & vbNewLine & _
            "a.encounter_number IS NOT NULL AND " & vbNewLine & _
            "a.encounter_number <> '0' AND" & vbNewLine & _
            "e.encounter_number = a.encounter_number and " & vbNewLine & _
            "(a.APPT_DATE_TIME between (sysdate - 365) and sysdate) and " & vbNewLine & _
            "a.patient_unit_number in " & sPatientList
 
'            "a.appt_classification = 2 and " & vbNewLine & _

Set oRec = goDBDAS.OpenRecordset(sStmt)

GetAmbulatoryVisits = True
GoTo CleanUp

ErrHandler:
    If Err = -2147220874 Then
        MsgBxARMd "The selected report has timed out, please re-run this report again.", vbOKOnly + vbInformation
        Err.Clear
    Else
        gobjshared.ShowVBError Err.Number, Err.Description, "RepRegistryCandidates.GetAmbulatoryVisits", "CHCSII_Reports", vbExclamation
    End If
CleanUp:
    Exit Function
End Function
Private Function GetMedications(oReportItem As ReportItem, _
                                sPatientList As String, _
                                oRec As ADODB.Recordset) As Boolean
On Error GoTo ErrHandler

GetMedications = False

Dim sStmt As String
Dim sFromSQL As String
Dim sWhereSQL As String

'-- Select
    sStmt = "Select distinct " & vbNewLine & _
                "e.unit_number, " & vbNewLine & _
                "r.representation as Medication, " & vbNewLine & _
                "decode(ro.date_last_dispensed,null,tq.start_gm_time,ro.date_last_dispensed) as date_last_dispensed " & vbNewLine

'-- From
    sStmt = sStmt & _
            "From " & vbNewLine & _
                "timing_qty tq, " & vbNewLine & _
                "rsform_context rc, " & vbNewLine & _
                "rsform r, " & vbNewLine & _
                "drug d, " & vbNewLine & _
                "rx_order ro, " & vbNewLine & _
                "event e " & vbNewLine

'-- Where
    sStmt = sStmt & _
        "Where " & vbNewLine & _
            "tq.event_id = e.event_id and " & vbNewLine & _
            "rc.context_ncid = 2000 and " & vbNewLine & _
            "rc.rsform_id = r.rsform_id and " & vbNewLine & _
            "r.ncid = d.product_ncid and " & vbNewLine & _
            "d.event_id = ro.event_id and " & vbNewLine & _
            "((tq.start_gm_time BETWEEN (SYSDATE - 365) AND SYSDATE) OR (ro.date_last_dispensed between (sysdate - 365) and sysdate)) and " & vbNewLine & _
            "ro.event_id = e.event_id and " & vbNewLine & _
            "e.event_subtype_ncid = 1226 and " & vbNewLine & _
            "e.event_type_ncid = 3006 and " & vbNewLine & _
            "e.unit_number in " & sPatientList

Set oRec = goDBDAS.OpenRecordset(sStmt)

GetMedications = True
GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepRegistryCandidates.GetMedications", "CHCSII_Reports", vbExclamation
CleanUp:
    Exit Function
End Function

Private Sub Class_Initialize()
    '--- SCR-ARM-16148 MSolano 9/10/01; COMMON TO MULTIPLE ARM-SCRs
    '--- Init routine should be called from first top level creatable coclass
    '--- It can be called multiple times, but will actually run only once.
    '--- Instrument all creatable classes in case dll's multiple entry
    '--- points are actually used
     Call ARM_InitAllTransactionsForCHCSIIReports
End Sub


