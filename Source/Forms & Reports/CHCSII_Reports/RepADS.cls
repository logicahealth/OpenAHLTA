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
Attribute VB_Name = "RepADS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private moCollRecSet As Collection
Private moCollFormulas As Collection

Private Function DataGetADSDiagnosis(sReportName As String, _
                                     oReportItem As ReportItem, _
                                     oReport As CRAXDRT.Report, _
                                     oCollFormulas As Collection) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler
DataGetADSDiagnosis = False

Dim sStmt As String
Dim i As Integer
Dim oRec As ADODB.Recordset
Dim sDateFilter As String
Dim sFilter As String
Dim sFilter2 As String
Dim sRNDSFilter As String

'-- Get Filter Information for SQL
'----------------------------------------------------------------------------
If Not GetSQLDateRange(oReportItem, sDateFilter) Then GoTo CleanUp
If Not GetSQLCriteria(oReportItem, 1, sFilter) Then GoTo CleanUp
If Not GetSQLCriteria(oReportItem, 2, sFilter2) Then GoTo CleanUp

If oReportItem.SecondaryFuncSelected = False Then   'Don't include RNDS
   sRNDSFilter = " And ( a.APPT_TYPE <> " & SQLQuote(RNDSALIAS1) & " ) AND (a.APPT_TYPE <> " & SQLQuote(RNDSALIAS2) & ")"
Else
   sRNDSFilter = NULLStr
End If

'-- Select
sStmt = "SELECT /*+ RULE */ OV.VALUE_NCID DIAGNOSISNCID, " & vbNewLine & _
                "DG.DESCRIPTION DIAGNOSIS," & vbNewLine & _
                "DG.ICD9" & vbNewLine

'-- Where
sStmt = sStmt & _
        "FROM" & vbNewLine & _
              "APPOINTMENT A," & vbNewLine & _
              "EVENT E," & vbNewLine & _
              "CLINICAL_OBSERVATION CO," & vbNewLine & _
              "OBSERVATION_VALUE OV," & vbNewLine & _
              "RPT_DIAGNOSIS DG" & vbNewLine

'-- Where
sStmt = sStmt & _
        "Where " & vbNewLine & _
            "DG.NCID = OV.VALUE_NCID AND" & vbNewLine & _
            "OV.OBSERVATION_SEQ = CO.OBSERVATION_SEQ AND" & vbNewLine & _
            "OV.EVENT_ID = CO.EVENT_ID AND" & vbNewLine & _
            "CO.EVENT_ID = E.EVENT_ID AND" & vbNewLine & _
            "E.EVENT_SUBTYPE_NCID = 658 AND" & vbNewLine & _
            "E.EVENT_TYPE_NCID = 90803 AND" & vbNewLine & _
            "E.ENCOUNTER_NUMBER=A.ENCOUNTER_NUMBER AND" & vbNewLine & _
            "a.Encounter_Status in (4,5,6) and" & vbNewLine & _
            sDateFilter & " and" & vbNewLine & _
            sFilter2 & " and" & vbNewLine & _
            sFilter & sRNDSFilter

Set oRec = goDBDAS.OpenRecordset(sStmt)
DoEvents

'--------------------------------
'-- Set Data
'----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection

    moCollRecSet.Add oRec, "diagnosis"

'-- Create Report Object
'-----------------------------------------------------------------------------
    If Not CreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo CleanUp


    DoEvents

DataGetADSDiagnosis = True
GoTo CleanUp
ErrHandler:
    If Err = -2147220874 Then
        MsgBxARMd "The selected report has timed out, please re-run this report again.", vbOKOnly + vbInformation
        Err.Clear
    Else
        Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.DataGetADSDiagnosis", "CHCSII_Reports", vbExclamation)
    End If

CleanUp:
    Set oRec = Nothing
    Exit Function
End Function

Private Function DataGetADSOutpatientWorkload(sReportName As String, _
                                              oReportItem As ReportItem, _
                                              oReport As CRAXDRT.Report, _
                                              oCollFormulas As Collection) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler
DataGetADSOutpatientWorkload = False

Dim sStmt As String
Dim i As Integer
Dim rs As ADODB.Recordset
Dim sFilter As String
Dim sDateFilter As String
Dim sClosedFilter As String
Dim oRec As ADODB.Recordset

'-- Get Filter Information for SQL
'----------------------------------------------------------------------------
If Not GetSQLDateRange(oReportItem, sDateFilter) Then GoTo CleanUp
If Not GetSQLCriteria(oReportItem, 1, sFilter) Then GoTo CleanUp
If Not GetClosedEncounters(sClosedFilter) Then GoTo CleanUp

'-- Select
sStmt = "SELECT" & vbNewLine & _
            "SUBSTR(pc.Code,1,1) as Service," & vbNewLine & _
            "GRP_ID" & vbNewLine & _
        "FROM" & vbNewLine & _
            "patient_category pc," & vbNewLine & _
            "mmi_genericdata mg," & vbNewLine & _
            "appointment a" & vbNewLine & _
        "WHERE" & vbNewLine & _
            "pc.code  = Trim(substr(mg.value_text,1,4)) and" & vbNewLine & _
            "mg.value_ncid is not null and mg.type_ncid = 14501721 and" & vbNewLine & _
            "mg.unit_number = a.patient_unit_number and " & vbNewLine & _
            "a.appt_classification in (0,2) and" & vbNewLine & _
            sClosedFilter & " and" & vbNewLine & _
            sDateFilter & " and" & vbNewLine & _
            sFilter

Set oRec = goDBDAS.OpenRecordset(sStmt)
DoEvents

'--------------------------------
'-- Set Data
'----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection

    moCollRecSet.Add oRec, "Workload"

'-- Create Report Object
'-----------------------------------------------------------------------------
    If Not CreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo CleanUp


       
    DoEvents

DataGetADSOutpatientWorkload = True
GoTo CleanUp
ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.DataGetADSOutpatientWorkload", "CHCSII_Reports", vbExclamation)
CleanUp:
    Set rs = Nothing
    Exit Function
End Function

Private Function DataGetADSInpatientWorkload(sReportName As String, _
                                              oReportItem As ReportItem, _
                                              oReport As CRAXDRT.Report, _
                                              oCollFormulas As Collection) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler
DataGetADSInpatientWorkload = False

Dim sStmt As String
Dim i As Integer
Dim rs As ADODB.Recordset
Dim sFilter As String
Dim sDateFilter As String
Dim sClosedFilter As String
Dim oRec As ADODB.Recordset

'-- Get Filter Information for SQL
'----------------------------------------------------------------------------
If Not GetSQLDateRange(oReportItem, sDateFilter) Then GoTo CleanUp
If Not GetSQLCriteria(oReportItem, 1, sFilter) Then GoTo CleanUp
If Not GetClosedEncounters(sClosedFilter) Then GoTo CleanUp

'-- Select
sStmt = "SELECT" & vbNewLine & _
            "SUBSTR(pc.Code,1,1) as Service," & vbNewLine & _
            "GRP_ID" & vbNewLine & _
        "FROM" & vbNewLine & _
            "patient_category pc," & vbNewLine & _
            "mmi_genericdata mg," & vbNewLine & _
            "appointment a" & vbNewLine & _
        "WHERE" & vbNewLine & _
            "pc.code  = Trim(substr(mg.value_text,1,4)) and" & vbNewLine & _
            "mg.value_ncid is not null and mg.type_ncid = 14501721 and" & vbNewLine & _
            "mg.unit_number = a.patient_unit_number and " & vbNewLine & _
            "a.appt_classification in (1) and" & vbNewLine & _
            sClosedFilter & " and" & vbNewLine & _
            sDateFilter & " and" & vbNewLine & _
            sFilter

Set oRec = goDBDAS.OpenRecordset(sStmt)
DoEvents

'--------------------------------
'-- Set Data
'----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection

    moCollRecSet.Add oRec, "Workload"

'-- Create Report Object
'-----------------------------------------------------------------------------
    If Not CreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo CleanUp


       
    DoEvents

DataGetADSInpatientWorkload = True
GoTo CleanUp
ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.DataGetADSInpatientWorkload", "CHCSII_Reports", vbExclamation)
CleanUp:
    Set rs = Nothing
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
    Case ADSPatientCategoriesProvider
        If Not DataGetADSPatientCategoriesProvider(sReportName, oReportItem, oReport, oCollFormulas) Then GoTo CleanUp
    Case ADSDiagnosis
        If Not DataGetADSDiagnosis(sReportName, oReportItem, oReport, oCollFormulas) Then GoTo CleanUp
    Case ADSClinicOutpatientWorkload
        If Not DataGetADSOutpatientWorkload(sReportName, oReportItem, oReport, oCollFormulas) Then GoTo CleanUp
    Case ADSClinicInpatientWorkload
            If Not DataGetADSInpatientWorkload(sReportName, oReportItem, oReport, oCollFormulas) Then GoTo CleanUp
    Case Else
End Select

BuildReport = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepRegistryCandidates.BuildReport", "CHCSII_Reports", vbExclamation
CleanUp:
    
End Function

Private Function DataGetADSPatientCategoriesProvider(sReportName As String, _
                                                     oReportItem As ReportItem, _
                                                     oReport As CRAXDRT.Report, _
                                                     oCollFormulas As Collection) As Boolean


If Not DebugMode Then On Error GoTo ErrHandler
DataGetADSPatientCategoriesProvider = False

Dim oCrystal As Crystal
Dim sStmt As String
Dim sFilter As String
Dim sDateFilter As String
Dim i As Integer
Dim oRec As ADODB.Recordset
Dim moCollRecSet As Collection

'-- Get Filter Information for SQL
'----------------------------------------------------------------------------
If Not GetSQLDateRange(oReportItem, sDateFilter) Then GoTo CleanUp
If Not GetSQLCriteria(oReportItem, 1, sFilter) Then GoTo CleanUp

'-- Select
sStmt = "SELECT" & vbNewLine & _
            "cl.Name as Clinic_Name," & vbNewLine & _
            "cl.meprs_code," & vbNewLine & _
            "cl.dmis_id_code," & vbNewLine & _
            "mg.value_text as PatientCategory" & vbNewLine

'-- From
sStmt = sStmt & _
        "FROM" & vbNewLine & _
            "mmi_genericdata mg," & vbNewLine & _
            "clinic cl," & vbNewLine & _
            "appointment a" & vbNewLine

'-- Where
sStmt = sStmt & _
        "WHERE" & vbNewLine & _
            "mg.type_ncid=14501721 and" & vbNewLine & _
            "mg.unit_number = a.patient_unit_number and" & vbNewLine & _
            "cl.ncid = a.clinic_ncid and" & vbNewLine & _
            sDateFilter & " and" & vbNewLine & _
            sFilter
  
Set oRec = goDBDAS.OpenRecordset(sStmt)
DoEvents

'--------------------------------
'-- Set Data
'----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection

    moCollRecSet.Add oRec, "patientcategories"

'-- Create Report Object
'-----------------------------------------------------------------------------
    If Not CreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo CleanUp


    
DataGetADSPatientCategoriesProvider = True
GoTo CleanUp
ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "RepADS.DataGetADSPatientCategoriesProvider", "CHCSII_Reports", vbExclamation)
CleanUp:
    Set oCrystal = Nothing
    Set oRec = Nothing
    Set moCollRecSet = Nothing
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


