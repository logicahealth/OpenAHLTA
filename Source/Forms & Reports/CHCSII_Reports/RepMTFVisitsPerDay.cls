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
Attribute VB_Name = "RepMTFVisitsPerDay"
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

    '-- Get Candidates
    '----------------------------------------------------------
    If Not GetData(oReportItem, oRec, oCollFormulas) Then GoTo CleanUp         ' data for provider
    moCollRecSet.Add oRec, "TotalVsitsMTF"

'-- Create Report Object
'-----------------------------------------------------------------------------
    If Not CreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo CleanUp

    BuildReport = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepMTFVisitsPerDay.BuildReport", "CHCSII_Reports", vbExclamation
CleanUp:
    Set oCrystal = Nothing
    Set oRec = Nothing
    Set oPeerRec = Nothing
    
End Function

Private Function GetData(oReportItem As ReportItem, _
                           oRec As ADODB.Recordset, _
                           oCollFormulas As Collection) As Boolean

    Dim sStmt As String
    Dim sStmtNonER As String
    Dim sStmtER As String
    Dim sStmtTotal As String
    Dim sHeader As String
    
    Dim sQuestion As String
    Dim iResponse As Integer
    
    On Error GoTo ErrHandler
    GetData = False
    Dim sDateFilter As String
    Dim sFilter As String
    
    
    '-- Get Filter Information for SQL
    '----------------------------------------------------------------------------
    If DateDiff("d", oReportItem.DateFrom, oReportItem.DateTo) > 31 Then
       sQuestion = "The data graphed may appear unreadable if the number of days in the date range exceed 31, proceed?"
       If MsgBxARMd(sQuestion, vbYesNo + vbQuestion) = vbNo Then Exit Function
    End If
    
    sDateFilter = sDateFilter & " (APPT_DATE_TIME >= " & GetOracleStartDate(oReportItem.DateFrom) & vbCrLf & _
                                " AND APPT_DATE_TIME <= " & GetOracleEndDate(oReportItem.DateTo) & ") "
    
    If oReportItem.Criteria1 = MTF Then
        sFilter = " WHERE  Facility_NCID = " & oReportItem.CriteriaSel1.Item(1).NCID & " And " & vbCrLf
    ElseIf oReportItem.Criteria1 = MTFDMIS Then
        sFilter = " WHERE  CLINIC_NCID IN (" & GetClinicNCIDsFromDMIS(oReportItem.CriteriaSel1) & ") AND " & vbCrLf
    End If
    
    sStmtNonER = "SELECT TO_CHAR(APPT_DATE_TIME,'MM/DD/YYYY') AS APPT_DATE, TO_CHAR(COUNT(*)) AS VISITS" & vbCrLf & _
        "FROM APPOINTMENT_HISTORY" & vbCrLf & _
        sFilter & sDateFilter & vbCrLf & _
            "AND (MEPRS_CODE NOT like 'BIA%' AND MEPRS_CODE NOT like 'BHI%' OR MEPRS_CODE IS NULL)" & vbCrLf & _
            "AND (APPT_STATUS = 'KEPT' OR  APPT_STATUS = 'LWOBS' OR  (APPT_STATUS = 'CANCEL' AND APPT_CANCEL_REASON = 'LWOBS'))" & vbCrLf & _
        "GROUP BY TO_CHAR(APPT_DATE_TIME,'MM/DD/YYYY')"
    
    sStmtER = "SELECT TO_CHAR(APPT_DATE_TIME,'MM/DD/YYYY') AS APPT_DATE, TO_CHAR(COUNT(*)) AS VISITS" & vbCrLf & _
        "FROM APPOINTMENT_HISTORY" & vbCrLf & _
        sFilter & sDateFilter & vbCrLf & _
            "AND (MEPRS_CODE like 'BIA%' or MEPRS_CODE like 'BHI%')" & vbCrLf & _
            "AND (APPT_STATUS = 'KEPT' OR  APPT_STATUS = 'LWOBS' OR  (APPT_STATUS = 'CANCEL' AND APPT_CANCEL_REASON = 'LWOBS'))" & vbCrLf & _
        "GROUP BY TO_CHAR(APPT_DATE_TIME,'MM/DD/YYYY')"
    
    sStmtTotal = "SELECT TO_CHAR(APPT_DATE_TIME,'MM/DD/YYYY') AS APPT_DATE, TO_CHAR(COUNT(*)) AS VISITS" & vbCrLf & _
        "FROM APPOINTMENT_HISTORY" & vbCrLf & _
        sFilter & sDateFilter & vbCrLf & _
            "AND (APPT_STATUS = 'KEPT' OR  APPT_STATUS = 'LWOBS' OR  (APPT_STATUS = 'CANCEL' AND APPT_CANCEL_REASON = 'LWOBS'))" & vbCrLf & _
        "GROUP BY TO_CHAR(APPT_DATE_TIME,'MM/DD/YYYY')"
    
    sStmt = "SELECT C.APPT_DATE, NVL(A.VISITS, '0') AS NONER_VISITS, " & vbCrLf & _
                "NVL(B.VISITS, '0') AS ER_VISITS, C.VISITS AS TOTAL_VISITS" & vbCrLf & _
            "FROM (" & sStmtNonER & ") A," & vbCrLf & _
                 "(" & sStmtER & ") B," & vbCrLf & _
                 "(" & sStmtTotal & ") C" & vbCrLf & _
            "WHERE C.APPT_DATE = B.APPT_DATE(+) AND C.APPT_DATE = A.APPT_DATE(+)" & vbCrLf
    
    Set oRec = goDBDAS.OpenRecordset(sStmt)
    
    If oReportItem.Criteria1 = MTFDMIS Then
        On Error Resume Next
        oCollFormulas.Remove ("Header2")
        For iResponse = 1 To oReportItem.CriteriaSel1.Count
            If Not sHeader = "" Then sHeader = sHeader & ", "
            sHeader = sHeader & oReportItem.CriteriaSel1(iResponse).MTF_NAME
        Next
        sHeader = "DMIS IDs selected: " & sHeader
        oCollFormulas.Add CrystalQuote(sHeader), "Header2"
    End If
    
    GetData = True
    GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepMTFVisitsPerDay.GetData", "CHCSII_Reports", vbExclamation
CleanUp:
    Exit Function
End Function

