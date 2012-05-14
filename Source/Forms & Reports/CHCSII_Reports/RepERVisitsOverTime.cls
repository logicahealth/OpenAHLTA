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
Attribute VB_Name = "RepERVisitsOverTime"
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
    Dim oPeerRec As ADODB.Recordset
    Dim sSponsorUnitNumber As String
    Dim sUnitNumber As String
    DoEvents
    
    '-- Get Data
    '----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection

    '-- Get Data for the MTF
    '----------------------------------------------------------
    If Not GetData(oReportItem, oRec, oCollFormulas) Then GoTo CleanUp
    moCollRecSet.Add oRec, "ERVsitsOverTimeMTF"

'-- Create Report Object
'-----------------------------------------------------------------------------
    If Not CreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo CleanUp

    BuildReport = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepERVisitsOverTime.BuildReport", "CHCSII_Reports", vbExclamation
CleanUp:
    Set oCrystal = Nothing
    Set oRec = Nothing
    Set oPeerRec = Nothing
    
End Function

Private Function GetData(oReportItem As ReportItem, _
                           oRec As ADODB.Recordset, _
                           ByRef colFormulas As Collection) As Boolean

    Dim sStmt As String
    
    On Error GoTo ErrHandler
    GetData = False
    Dim sDateFilter As String
    Dim sFilter As String
    Dim sQuestionnaireFilter As String
    Dim sStartDate As String
    Dim sEndDate As String
    Dim strDateInfo                    As String
    Dim sHeader As String
    Dim iResponse As Integer
    Dim sHint As String
    
    '-- Get Filter Information for SQL
    '----------------------------------------------------------------------------
    sStartDate = GetOracleStartDate(GetFirstDayOfMonth(oReportItem.DateFrom))
    sEndDate = GetOracleEndDate(GetLastDayOfMonth(oReportItem.DateTo))
    
    sDateFilter = sDateFilter & " (APPT_DATE_TIME >= " & sStartDate & vbCrLf & " AND APPT_DATE_TIME <= " & sEndDate & ") "
    
    If oReportItem.Criteria1 = MTF Then
        sFilter = " WHERE  Facility_NCID = " & oReportItem.CriteriaSel1.Item(1).NCID & " And " & vbCrLf
        sHint = " /*+ RULE */  /*+ INDEX (APPOINTMENT_HISTORY APPOINTMENT_HISTORY_X2) */"
    ElseIf oReportItem.Criteria1 = MTFDMIS Then
        sFilter = " WHERE  CLINIC_NCID IN (" & GetClinicNCIDsFromDMIS(oReportItem.CriteriaSel1) & ") AND " & vbCrLf
        sHint = " /*+ RULE */ /*+ INDEX (APPOINTMENT_HISTORY APPOINTMENT_HISTORY_X3) */"
    End If
    
    sStmt = "SELECT " & sHint & vbCrLf & _
            "TO_CHAR(APPT_DATE_TIME,'MON YYYY') AS APPT_DATE, TO_CHAR(COUNT(*)) AS VISITS " & vbCrLf _
        & "FROM APPOINTMENT_HISTORY" & vbCrLf
    sStmt = sStmt & sFilter & sDateFilter & vbCrLf & _
            "AND (MEPRS_CODE like 'BIA%' or MEPRS_CODE like 'BHI%')" & vbCrLf & _
            "AND (APPT_STATUS = 'KEPT' OR  APPT_STATUS = 'LWOBS' OR  (APPT_STATUS = 'CANCEL' AND APPT_CANCEL_REASON = 'LWOBS'))" & vbCrLf & _
            "GROUP BY TO_CHAR(APPT_DATE_TIME,'MON YYYY')"
    
                
    Set oRec = goDBDAS.OpenRecordset(sStmt)
    GetData = True
    
    On Error Resume Next
    '--- From/To date string in the report formular
    If IsDate(oReportItem.DateFrom) Then strDateInfo = "From  " & Format(oReportItem.DateFrom, "mmm yyyy") & "   "
    If IsDate(oReportItem.DateTo) Then strDateInfo = strDateInfo & "Through  " & Format(oReportItem.DateTo, "mmm yyyy")
    colFormulas.Add CrystalQuote(strDateInfo), "DateInfo"
    If oReportItem.Criteria1 = MTFDMIS Then
        On Error Resume Next
        colFormulas.Remove ("Header2")
        For iResponse = 1 To oReportItem.CriteriaSel1.Count
            If Not sHeader = "" Then sHeader = sHeader & ", "
            sHeader = sHeader & oReportItem.CriteriaSel1(iResponse).MTF_NAME
        Next
        sHeader = "DMIS IDs selected: " & sHeader
        colFormulas.Add CrystalQuote(sHeader), "Header2"
    End If
    
    GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepERVisitsOverTime.GetData", "CHCSII_Reports", vbExclamation
CleanUp:
    Exit Function
End Function

