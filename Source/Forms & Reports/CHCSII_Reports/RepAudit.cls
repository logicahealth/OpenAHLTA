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
Attribute VB_Name = "RepAudit"
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

Select Case oReportItem.Report
    Case AUDIT_AllergyVerified
        If Not BuildAllergyVerified(sReportName, oReportItem, oReport, oCollFormulas) Then GoTo CleanUp
End Select

   
BuildReport = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepAudit.BuildReport", "CHCSII_Reports", vbExclamation
CleanUp:
    
End Function

Public Function BuildAllergyVerified(sReportName As String, _
                                     oReportItem As ReportItem, _
                                     oReport As CRAXDRT.Report, _
                                     oCollFormulas As Collection) As Boolean
On Error GoTo ErrHandler
BuildAllergyVerified = False

Dim oRec As ADODB.Recordset
Dim sStmt As String
Dim sFilter As String
Dim sDateFilter As String
Dim sDenominator As String
Dim sNumerator As String


DoEvents


'-- Get Filter Information for SQL
'----------------------------------------------------------------------------
If Not GetSQLDateRange(oReportItem, sDateFilter) Then GoTo CleanUp
If Not GetSQLCriteria(oReportItem, 1, sFilter) Then GoTo CleanUp

'-- Get Data
'----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection

'-- Get Denominator
'----------------------------------------------------------------------------

sStmt = "select " & vbNewLine & _
                "count(*) as denominator" & vbNewLine & _
            "From" & vbNewLine & _
                "appointment a" & vbNewLine & _
            "Where" & vbNewLine & _
                "a.encounter_status in (6,4) and" & vbNewLine & _
                sFilter & " and" & vbNewLine & _
                sDateFilter
    
    goDBSqlOps.Execute sStmt
    sDenominator = goDBSqlOps.Value("Denominator")

    '-- Get Numerator
    '----------------------------------------------------------------------------
    sStmt = "select /*+ ORDERED INDEX (a APPOINTMENT_X3) (e EVENT_X1)  USE_NL(ci) */ count(*) as numerator" & vbCrLf 'USE_NL(ci)
    sStmt = sStmt & "from appointment a, event e, comment_item ci" & vbCrLf
    sStmt = sStmt & "Where " & vbCrLf
    sStmt = sStmt & "substr(ci.item_text,23,1) is not null and" & vbCrLf
    sStmt = sStmt & "substr(ci.item_text,1,22)='ALLERGYVERIFIEDBYNCID=' and" & vbCrLf
    sStmt = sStmt & "ci.event_id = e.event_id and" & vbCrLf
    sStmt = sStmt & "e.event_type_ncid = 209614 and" & vbCrLf
    sStmt = sStmt & "e.event_subtype_ncid = 209683 and" & vbCrLf
    sStmt = sStmt & "e.unit_number = a.patient_unit_number and" & vbCrLf
    sStmt = sStmt & "a.encounter_status in (6,4) and" & vbCrLf
    sStmt = sStmt & sFilter & " and" & vbCrLf
    sStmt = sStmt & sDateFilter

    goDBSqlOps.Execute sStmt
    sNumerator = goDBSqlOps.Value("numerator")

    oCollFormulas.Add CrystalQuote(sNumerator), "Numerator"
    oCollFormulas.Add CrystalQuote(sDenominator), "Denominator"

    '-- Create Report Object -----------------------------------------------------------------------------
    If Not CreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo CleanUp

BuildAllergyVerified = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepAudit.BuildAllergyVerified", "CHCSII_Reports", vbExclamation
CleanUp:
    Set oRec = Nothing
End Function

Private Sub Class_Initialize()
    '--- SCR-ARM-16148 MSolano 9/10/01; COMMON TO MULTIPLE ARM-SCRs
    '--- Init routine should be called from first top level creatable coclass
    '--- It can be called multiple times, but will actually run only once.
    '--- Instrument all creatable classes in case dll's multiple entry
    '--- points are actually used
     Call ARM_InitAllTransactionsForCHCSIIReports
End Sub


