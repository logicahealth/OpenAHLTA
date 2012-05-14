VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RepADSAppointment"
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
DoEvents

'-- Get Data
'----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection

    '-- Get Appointments
    '----------------------------------------------------------
    If Not GetAppointments(oReportItem, oRec) Then GoTo CleanUp
    moCollRecSet.Add oRec, "appointments"

'-- Create Report Object
'-----------------------------------------------------------------------------
    If Not CreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo CleanUp

BuildReport = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "CPSServicesDue.BuildReport", "CHCSII_Reports", vbExclamation
CleanUp:
    Set oCrystal = Nothing
    Set oRec = Nothing
    
End Function

Private Function GetAppointments(oReportItem As ReportItem, _
                                 oRec As ADODB.Recordset) As Boolean
On Error GoTo ErrHandler

GetAppointments = False

Dim sStmt As String
Dim sDateFilter As String
Dim sFilter As String
Dim sFilter2 As String
Dim oRec2 As ADODB.Recordset
Dim sRNDSFilter As String

'-- Get Filter Information for SQL
'----------------------------------------------------------------------------
If Not GetSQLDateRange(oReportItem, sDateFilter) Then GoTo CleanUp
If Not GetSQLCriteria(oReportItem, 1, sFilter) Then GoTo CleanUp

If oReportItem.SecondaryFuncSelected = False Then   'Don't include RNDS
   sRNDSFilter = " And ( a.APPT_TYPE <> " & SQLQuote(RNDSALIAS1) & " ) AND (a.APPT_TYPE <> " & SQLQuote(RNDSALIAS2) & ")"
Else
   sRNDSFilter = NULLStr
End If

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
                "a.APPT_CANCEL_REASON," & vbNewLine & _
                "Decode(a.ENCOUNTER_STATUS,1,'Checked In',2,'Waiting',3,'In Progress',4,'Complete',5,'Needs CoSignature',6,'Updated',7,'LWOBS',Null,'Pending','','Pending',100,'                              ','Unknown') as Appt_Status, " & _
                "a.appt_status as Appointment_Status, a.appt_classification " & vbNewLine

'-- From
    sStmt = sStmt & _
            "From " & vbNewLine & _
            "clinic cl, " & vbNewLine & _
            "rsform_context  rcp, rsform rp, " & vbNewLine & _
            "mmi_name MN, " & vbNewLine & _
            "appointment a " & vbNewLine

'-- Where
    sStmt = sStmt & _
            "Where " & vbNewLine & _
            "cl.ncid = a.clinic_ncid and " & vbNewLine & _
            "rcp.context_ncid = 2000 and " & vbNewLine & _
            "rcp.rsform_id (+) = rp.rsform_id and " & vbNewLine & _
            "rp.ncid (+) = a.clinician_ncid and " & vbNewLine & _
            "mn.name_type_cid = 187 and " & vbNewLine & _
            "mn.unit_number = a.patient_unit_number and" & vbNewLine & _
            sFilter & " and" & vbNewLine & sDateFilter & sRNDSFilter
            
            If oReportItem.Criteria2 = Provider Then
                If Not GetSQLCriteria(oReportItem, 2, sFilter2) Then GoTo CleanUp
                sStmt = sStmt & " and" & vbNewLine & sFilter2
            End If

Set oRec2 = goDBDAS.OpenRecordset(sStmt)

If oRec2.RecordCount = 0 Then
    Set oRec = oRec2
Else

If Not DuplicateDefinition(oRec2, oRec) Then GoTo CleanUp
If Not CopyRecordSet(oRec2, oRec) Then GoTo CleanUp


    oRec.MoveFirst
    Do While Not oRec.EOF
        If Trim(oRec.Fields("APPT_CANCEL_REASON").Value) = "PATIENT CANCELLED" Or _
           Trim(oRec.Fields("APPT_CANCEL_REASON").Value) = "FACILITY CANCELLED" Then
            oRec.Fields("Appt_Status").Value = oRec.Fields("APPT_CANCEL_REASON").Value
            oRec.Update
        End If
        oRec.MoveNext
    Loop
End If

gAdjustApptStatusForDisplay oRec

GetAppointments = True
GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepADSAppointment.GetAppointments", "CHCSII_Forms", vbExclamation
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


