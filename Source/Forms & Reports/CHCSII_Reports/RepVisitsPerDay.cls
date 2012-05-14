VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RepVisitsPerDay"
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

    '-- Get Data for the Provider
    '----------------------------------------------------------
    If Not GetData(oReportItem, oRec) Then GoTo CleanUp
    moCollRecSet.Add oRec, "TotalVsitsProvider"
                
'-- Create Report Object
'-----------------------------------------------------------------------------
    If Not CreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo CleanUp

    BuildReport = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepGenericClass.BuildReport", "CHCSII_Reports", vbExclamation
CleanUp:
    Set oCrystal = Nothing
    Set oRec = Nothing
    Set oPeerRec = Nothing
    
End Function

Private Function GetData(oReportItem As ReportItem, _
                           oRec As ADODB.Recordset) As Boolean

    Dim sStmt As String
    Dim sStmtPeer As String
    Dim sStmtPRV As String
    
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
    
    sFilter = " WHERE CLINICIAN_NCID = " & oReportItem.CriteriaSel2.Item(1).NCID & " AND " & vbCrLf _
            & " CLINIC_NCID = " & oReportItem.CriteriaSel1.Item(1).NCID & " AND " & vbCrLf
    
    sStmt = "SELECT TO_CHAR(APPT_DATE_TIME,'MM/DD/YYYY') AS APPT_DATE, TO_CHAR(COUNT(*)) AS VISITS " & vbCrLf _
        & "FROM APPOINTMENT_HISTORY" & vbCrLf
    sStmt = sStmt & sFilter & sDateFilter & vbCrLf & _
            "AND (MEPRS_CODE NOT like 'BIA%' AND MEPRS_CODE NOT like 'BHI%' OR MEPRS_CODE IS NULL)" & vbCrLf & _
            "AND (APPT_STATUS = 'KEPT' OR  APPT_STATUS = 'LWOBS' OR  (APPT_STATUS = 'CANCEL' AND APPT_CANCEL_REASON = 'LWOBS'))" & vbCrLf & _
        " GROUP BY TO_CHAR(APPT_DATE_TIME,'MM/DD/YYYY')"
    
    '--- sub query for provider
    sStmtPRV = sStmt

    '--- sub query for peer
    sFilter = " WHERE Clinician_NCID <> " & oReportItem.CriteriaSel2.Item(1).NCID & " And " & vbCrLf & _
        " CLINIC_NCID = " & oReportItem.CriteriaSel1.Item(1).NCID & " AND " & vbCrLf
    
    sStmtPeer = "SELECT CLINICIAN_NCID, TO_CHAR(APPT_DATE_TIME,'MM/DD/YYYY') AS APPT_DATE, TO_CHAR(COUNT(*)) AS VISITS " & vbCrLf _
        & "FROM APPOINTMENT_HISTORY" & vbCrLf
        
    sStmtPeer = sStmtPeer & sFilter & sDateFilter & vbCrLf & _
            "AND (MEPRS_CODE NOT like 'BIA%' AND MEPRS_CODE NOT like 'BHI%' OR MEPRS_CODE IS NULL)" & vbCrLf & _
            "AND (APPT_STATUS = 'KEPT' OR  APPT_STATUS = 'LWOBS' OR  (APPT_STATUS = 'CANCEL' AND APPT_CANCEL_REASON = 'LWOBS'))" & vbCrLf & _
        " GROUP BY TO_CHAR(APPT_DATE_TIME,'MM/DD/YYYY'), CLINICIAN_NCID"
        
    sStmtPeer = "select appt_date, to_char(count(*)) as num_of_prov, to_char(sum(visits)) as total_visits, " & vbCrLf _
                & "to_char(round(avg(visits), 0)) as mean_visits, " & vbCrLf _
                & "to_char(round(stddev(all visits), 0)) as std" & vbCrLf _
          & "from (" & sStmtPeer & ")" & vbCrLf _
          & "group by appt_date"
       
    '--- Combine two SQL together
    sStmt = "SELECT A.APPT_DATE, A.VISITS, NVL(B.num_of_prov, 0) AS num_of_prov, " & vbNewLine & _
                "NVL(B.total_visits, 0) AS total_visits, " & vbNewLine & _
                "NVL(B.mean_visits, 0) AS mean_visits, " & vbNewLine & _
                "NVL(B.std, 0) AS std" & vbCrLf & _
            "FROM (" & sStmtPRV & ") A, (" & sStmtPeer & ") B" & vbCrLf & _
            "WHERE A.APPT_DATE = B.APPT_DATE(+)" & vbCrLf
    
                
    Set oRec = goDBDAS.OpenRecordset(sStmt)
    GetData = True
    GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepGenericClass.GetData", "CHCSII_Reports", vbExclamation
CleanUp:
    Exit Function
End Function


