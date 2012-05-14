VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RepDTDailyProvider"
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
Dim sNCIDS As String
DoEvents

'-- Get Data
'----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection
    
    '-- Get Provider NCIDS
    '----------------------------------------------------------
    If Not GetProviderNCIDs(oReportItem, sNCIDS) Then GoTo CleanUp

    '-- Get Services Due
    '----------------------------------------------------------
    If Not GetAppointments(oReportItem, oRec) Then GoTo CleanUp
    moCollRecSet.Add oRec, "appointments"

'-- Create Report Object
'-----------------------------------------------------------------------------
    If Not CreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo CleanUp


BuildReport = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepAppointment.BuildReport", "CHCSII_Reports", vbExclamation
CleanUp:
    Set oCrystal = Nothing
    Set oRec = Nothing
    
End Function

Private Function GetAppointments(oReportItem As ReportItem, _
                                 oRec As ADODB.Recordset) As Boolean
On Error GoTo ErrHandler

GetAppointments = False

Dim sStmt As String


'-- Select
    sStmt = "Select " & vbNewLine & _
                "a.CLINICIAN_NCID, " & vbNewLine & _
                "rp.representation as Provider_Name," & vbNewLine & _
                "a.PATIENT_UNIT_NUMBER, " & vbNewLine & _
                "Trim(mn.last_name ||', ' || mn.first_name || ' ' || " & vbNewLine & _
                "mn.middle_name) as Full_Name," & vbNewLine & _
                "a.Clinic_NCID, " & vbNewLine & _
                "cl.Name as Clinic_Name, " & vbNewLine & _
                "cl.meprs_code, " & vbNewLine & _
                "cl.dmis_id_code, " & vbNewLine & _
                "a.ENCOUNTER_STATUS,  " & vbNewLine & _
                "Decode(a.ENCOUNTER_STATUS,1,'Checked In',2,'Waiting',3,'In Progress',4,'Complete',5,'Needs CoSignature',6,'Updated',7,'LWOBS','Unknown') as Appt_Status " & vbNewLine

'-- From
    sStmt = sStmt & _
            "From " & vbNewLine & _
            "clinic cl, " & vbNewLine & _
            "rsform_context  rcp, rsform rp, " & vbNewLine & _
            "mmi_name MN, " & vbNewLine

        If oReportItem.Criteria2 = Diagnosis Then
            sStmt = sStmt & _
                "observation_value ov2, " & vbNewLine & _
                "clinical_observation co2, " & vbNewLine & _
                "semantic_link_element sle, " & vbNewLine & _
                "semantic_link sl, " & vbNewLine & _
                "clinical_event_header c, " & vbNewLine & _
                "event e, " & vbNewLine
        End If

    sStmt = sStmt & _
            " appointment a " & vbNewLine

'-- Where
    sStmt = sStmt & _
            "Where " & vbNewLine & _
            "cl.ncid = a.clinic_ncid and " & vbNewLine & _
            "rcp.context_ncid = 2000 and " & vbNewLine & _
            "rcp.rsform_id = rp.rsform_id and " & vbNewLine & _
            "rp.ncid (+) = a.clinician_ncid and " & vbNewLine & _
            "mn.name_type_cid = 187 and " & vbNewLine & _
            "mn.unit_number = a.patient_unit_number and a.encounter_status is not null " & vbNewLine
            
        If oReportItem.Criteria2 = Diagnosis Then
            sStmt = sStmt & _
            "and ov2.value_ncid = " & oReportItem.CriteriaSel2.Item(1).NCID & " " & vbNewLine & _
            "and ov2.observation_seq = co2.observation_seq and " & vbNewLine & _
            "ov2.event_id = co2.event_id and " & vbNewLine & _
            "co2.observation_ncid = 105328 and " & vbNewLine & _
            "co2.event_id = sle.link_event_id and " & vbNewLine & _
            "sle.semantic_link_ord = sl.semantic_link_ord and " & vbNewLine & _
            "sle.event_id = sl.event_id and " & vbNewLine & _
            "sl.relationship_ncid = 205655 and " & vbNewLine & _
            "sl.event_id = e.event_id and " & vbNewLine & _
            "e.event_subtype_ncid = 209683 and " & vbNewLine & _
            "e.event_type_ncid = 209614 and " & vbNewLine & _
            "e.encounter_number = a.encounter_number and " & vbNewLine & _
            "a.encounter_number is not null and " & vbNewLine & _
            "a.encounter_number <> 0 " & vbNewLine
        End If

    If Not BuildWhereClauseforAppointments(oReportItem, sStmt) Then GoTo CleanUp

Set oRec = goDBDAS.OpenRecordset(sStmt)

GetAppointments = True
GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepAppointment.GetAppointments", "CHCSII_Forms", vbExclamation
CleanUp:
    Exit Function
End Function

Private Function GetProviderNCIDs(oReportItem As ReportItem, _
                                  sNCIDS As String) As Boolean

On Error GoTo ErrHandler
GetProviderNCIDs = False

Dim i As Integer


    If oReportItem.Criteria1 = ProviderMulti Then
         
        sNCIDS = "("
    
        For i = 1 To oReportItem.CriteriaSel1.Count
        
            If i <> oReportItem.CriteriaSel1.Count Then
                sNCIDS = sNCIDS & oReportItem.CriteriaSel1.Item(i).NCID & ","
            Else
                sNCIDS = sNCIDS & oReportItem.CriteriaSel1.Item(i).NCID & ") " & vbNewLine
            End If
        Next

    End If

GetProviderNCIDs = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepDTDailyProvider.GetProviderNCIDs", "CHCSII_Reports", vbExclamation
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

