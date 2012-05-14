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
Attribute VB_Name = "RepScreenPRV"
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
    moCollRecSet.Add oRec, "ScreenPRV"

'-- Create Report Object
'-----------------------------------------------------------------------------
    If Not CreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo CleanUp

    BuildReport = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepQuestionSatisfaction.BuildReport", "CHCSII_Reports", vbExclamation
CleanUp:
    Set oCrystal = Nothing
    Set oRec = Nothing
    Set oPeerRec = Nothing
    
End Function

Private Function GetData(oReportItem As ReportItem, _
                           oRec As ADODB.Recordset, _
                           oCollFormulas As Collection) As Boolean

    Dim sStmt As String
    Dim sPeerStmt As String
    Dim sCriteria As String
    Dim sName As String
    
    Dim sStartDate As String
    Dim sEndDate As String
    Dim sDateFilter As String
    Dim strDateInfo As String
    Dim sClinicCriteria As String
    Dim lCriteriaCount As Long
    
    On Error GoTo ErrHandler
    GetData = False

    sStartDate = GetOracleStartDate(GetFirstDayOfMonth(oReportItem.DateFrom))
    sEndDate = GetOracleEndDate(GetLastDayOfMonth(oReportItem.DateTo))
    
    If oReportItem.CriteriaSel1.Count = 0 Then
        sClinicCriteria = ""
    Else
'        For lCriteriaCount = 1 To oReportItem.CriteriaSel1.Count
'            If Not sClinicCriteria = "" Then sClinicCriteria = sClinicCriteria & ", "
'            sClinicCriteria = sClinicCriteria & oReportItem.CriteriaSel1.Item(lCriteriaCount).NCID
'        Next
'        If InStr(1, sClinicCriteria, ",", vbTextCompare) = 0 Then
'           sClinicCriteria = "MEPRS_CODE = (SELECT MEPRS_CODE FROM CLINIC WHERE NCID = " & sClinicCriteria & ") AND " & vbCrLf
'        Else
'           sClinicCriteria = "MEPRS_CODE IN (SELECT MEPRS_CODE FROM CLINIC WHERE NCID IN (" & sClinicCriteria & ")) AND " & vbCrLf
'        End If
        sClinicCriteria = "ENTERING_LOCATION_NCID = " & oReportItem.CriteriaSel1.Item(1).NCID & " AND " & vbCrLf
    End If
    
    '-- Get Filter Information for SQL    '----------------------------------------------------------------------------
    sDateFilter = sDateFilter & " (START_DATE >= " & sStartDate & vbCrLf & " AND START_DATE <= " & sEndDate & ") "
    
    sCriteria = GetScreeningServiceItems(oReportItem.CriteriaSel3.Item(1).NCID)
    If sCriteria = "" Then
       MsgBxARMd "The screening service selected has no tests assigned to it. Report cancelled.", vbOKOnly + vbInformation
       Exit Function
    End If
    
    sName = oReportItem.CriteriaSel3.Item(1).Name
    
    sStmt = "SELECT  /*+ INDEX (LAB LAB_RAD_ORDERS_X3) */ TO_CHAR(COUNT(*)) AS TOTAL, TO_CHAR(TO_DATE(START_DATE), 'YYYY/MM') AS PER_DATE, " & vbCrLf
    sStmt = sStmt & "'" & sName & "' AS SERVICE" & vbCrLf
    sStmt = sStmt & "From LAB_RAD_ORDERS" & vbCrLf
    sStmt = sStmt & "WHERE " & sClinicCriteria
    sStmt = sStmt & "ORDERING_PROVIDER_NCID = '" & oReportItem.CriteriaSel2.Item(1).NCID & "' AND" & vbCrLf
    sStmt = sStmt & "UNIVERSAL_SERVICE_NCID IN (" & sCriteria & ") AND " & vbCrLf
    sStmt = sStmt & sDateFilter & vbCrLf
    sStmt = sStmt & "GROUP BY TO_CHAR(TO_DATE(START_DATE), 'YYYY/MM')"
    
    
    sPeerStmt = "SELECT  /*+ INDEX (LAB LAB_RAD_ORDERS_X3) */ PER_DATE, SERVICE, ROUND(AVG(ALL TOTAL), 0) AS MEAN,ROUND(STDDEV(ALL TOTAL), 0) AS STD_DEV" & vbCrLf
    sPeerStmt = sPeerStmt & "FROM (SELECT COUNT(*) AS TOTAL," & vbCrLf
    sPeerStmt = sPeerStmt & "TO_CHAR(TO_DATE(START_DATE), 'YYYY/MM') AS PER_DATE," & vbCrLf
    sPeerStmt = sPeerStmt & "'" & sName & "' AS SERVICE, ORDERING_PROVIDER_NCID" & vbCrLf
    sPeerStmt = sPeerStmt & "From LAB_RAD_ORDERS" & vbCrLf
    sPeerStmt = sPeerStmt & "WHERE " & sClinicCriteria
    sPeerStmt = sPeerStmt & "ORDERING_PROVIDER_NCID <> '" & oReportItem.CriteriaSel2.Item(1).NCID & "' AND" & vbCrLf
    sPeerStmt = sPeerStmt & "UNIVERSAL_SERVICE_NCID IN (" & sCriteria & ") AND " & vbCrLf
    sPeerStmt = sPeerStmt & sDateFilter & vbCrLf
    sPeerStmt = sPeerStmt & "GROUP BY TO_CHAR(TO_DATE(START_DATE), 'YYYY/MM'), ORDERING_PROVIDER_NCID )" & vbCrLf
    sPeerStmt = sPeerStmt & "GROUP BY SERVICE, PER_DATE"
    
    
    sStmt = "SELECT /*+ RULE */ A.TOTAL,  A.PER_DATE, A.SERVICE, TO_CHAR(NVL(B.MEAN, 0)) AS MEAN," & vbCrLf & _
                "TO_CHAR(NVL(B.STD_DEV, 0)) AS STD_DEV" & vbCrLf & _
            "FROM (" & sStmt & ") A, (" & sPeerStmt & ") B" & vbCrLf & _
            "WHERE A.PER_DATE = B.PER_DATE(+)" & vbCrLf
            

    Set oRec = goDBDAS.OpenRecordset(sStmt)
    
    GetData = True
    
    
    'Filter on date
    If IsDate(oReportItem.DateFrom) Then strDateInfo = "From  " & Format(oReportItem.DateFrom, "mmm yyyy") & "   "
    If IsDate(oReportItem.DateTo) Then strDateInfo = strDateInfo & "Through  " & Format(oReportItem.DateTo, "mmm yyyy")
       
    On Error Resume Next
    oCollFormulas.Remove "DateInfo"
    oCollFormulas.Add CrystalQuote(strDateInfo), "DateInfo"
        
    GoTo CleanUp
    
ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepQuestionSatisfaction.GetData", "CHCSII_Reports", vbExclamation
    'Resume Next
CleanUp:
    Exit Function
End Function

