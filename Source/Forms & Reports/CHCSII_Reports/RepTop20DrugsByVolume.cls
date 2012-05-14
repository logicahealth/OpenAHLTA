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
Attribute VB_Name = "RepTop20DrugsByVolume"
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
    moCollRecSet.Add oRec, "Top20DrugsByVolumePRV"

'-- Create Report Object
'-----------------------------------------------------------------------------
    If Not CreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo CleanUp

    BuildReport = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepTop20DrugsByVolume.BuildReport", "CHCSII_Reports", vbExclamation
CleanUp:
    Set oCrystal = Nothing
    Set oRec = Nothing
    Set oPeerRec = Nothing
    
End Function

Private Function GetData(oReportItem As ReportItem, _
                           oRec As ADODB.Recordset, _
                           oCollFormulas As Collection) As Boolean

    Dim sStmt As String
    Dim sPrv As String
    Dim sPeer As String
    Dim sFilter As String
    
    Dim sStartDate As String
    Dim sEndDate As String
    Dim sDateFilter As String
    Dim strDateInfo As String
    Dim sHint As String

    On Error GoTo ErrHandler
    
    GetData = False
    
    sHint = " /*+ ORDERED */ /*+ INDEX  (M MEDICATION_ORDERS_X5)  (F MEDICATION_FILLS_PK)  (D DRUG_COST_DATA_X1)  */"

    '-- Get Filter Information for SQL
    '----------------------------------------------------------------------------
    sStartDate = GetOracleStartDate(GetFirstDayOfMonth(oReportItem.DateFrom))
    sEndDate = GetOracleEndDate(GetLastDayOfMonth(oReportItem.DateTo))
    sDateFilter = sDateFilter & " (START_DATE >= " & sStartDate & vbCrLf & " AND START_DATE <= " & sEndDate & ") "
    
    sPrv = "SELECT TO_CHAR(ROWNUM) AS SEQ, GC3_NAME, TO_CHAR(VOLUME) AS VOLUME" & vbCrLf & _
            "FROM (" & vbCrLf & _
                "SELECT " & sHint & vbCrLf & " D.GC3_NAME, COUNT(*) AS VOLUME" & vbCrLf & _
                "FROM MEDICATION_ORDERS M, MEDICATION_FILLS F, DRUG_COST_DATA D" & vbCrLf & _
                "WHERE M.ORDERING_PROVIDER_NCID = " & oReportItem.CriteriaSel2.Item(1).NCID & " AND " & vbCrLf & _
                    "M.ENTERING_LOCATION_NCID = " & oReportItem.CriteriaSel1.Item(1).NCID & " AND" & vbCrLf & _
                    "M.FACILITY_NCID = F.FACILITY_NCID AND" & vbCrLf & _
                    "M.ORDER_NUMBER = F.ORDER_NUMBER AND" & vbCrLf & _
                    "M.DRUG_ALTERNATE_IDENTIFIER = D.NDC AND" & vbCrLf & _
                    "D.GC3_NAME IS NOT NULL AND" & vbCrLf & sDateFilter & vbCrLf & _
                    "GROUP BY D.GC3_NAME" & vbCrLf & _
                    "ORDER BY COUNT(*) DESC, D.GC3_NAME ASC" & vbCrLf & _
            ")" & vbCrLf & _
            "WHERE ROWNUM < 21" & vbCrLf
            
    sPeer = "SELECT GC3_NAME, AVG(ALL VOLUME) AS PEER_MEAN," & vbCrLf & _
                " ROUND(STDDEV(ALL VOLUME), 0) AS STD" & vbCrLf & _
            "FROM (" & vbCrLf & _
                "SELECT " & sHint & vbCrLf & "  D.GC3_NAME, M.ORDERING_PROVIDER_NCID, COUNT(*) AS VOLUME" & vbCrLf & _
                "FROM MEDICATION_ORDERS M, MEDICATION_FILLS F, DRUG_COST_DATA D" & vbCrLf & _
                "WHERE M.ORDERING_PROVIDER_NCID <> " & oReportItem.CriteriaSel2.Item(1).NCID & " AND" & vbCrLf & _
                    "M.ENTERING_LOCATION_NCID = " & oReportItem.CriteriaSel1.Item(1).NCID & " AND " & vbCrLf & _
                    "M.FACILITY_NCID = F.FACILITY_NCID AND" & vbCrLf & _
                    "M.ORDER_NUMBER = F.ORDER_NUMBER AND" & vbCrLf & _
                    "M.DRUG_ALTERNATE_IDENTIFIER = D.NDC AND" & vbCrLf & _
                    "D.LOWEST_UM_PRICE IS NOT NULL AND D.GC3_NAME IS NOT NULL AND" & vbCrLf & sDateFilter & vbCrLf & _
                "GROUP BY D.GC3_NAME, M.ORDERING_PROVIDER_NCID" & vbCrLf & _
            ")" & vbCrLf & _
            "GROUP BY GC3_NAME" & vbCrLf
            
    sStmt = "SELECT /*+ RULE */ A.SEQ, A.GC3_NAME, A.VOLUME, TO_CHAR(NVL(B.PEER_MEAN, 0)) AS PEER_MEAN, TO_CHAR(NVL(B.STD, 0)) AS STD" & vbCrLf & _
            "FROM (" & sPrv & ") A, (" & sPeer & ") B" & vbCrLf & _
            "WHERE A.GC3_NAME = B.GC3_NAME(+)" & vbCrLf
            
    Set oRec = goDBDAS.OpenRecordset(sStmt)
    GetData = True
    
    If IsDate(oReportItem.DateFrom) Then strDateInfo = "From  " & Format(oReportItem.DateFrom, "mmm yyyy") & "   "
    If IsDate(oReportItem.DateTo) Then strDateInfo = strDateInfo & "Through  " & Format(oReportItem.DateTo, "mmm yyyy")
       
    On Error Resume Next
    oCollFormulas.Remove "DateInfo"
    oCollFormulas.Add CrystalQuote(strDateInfo), "DateInfo"
    
    GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepTop20DrugsByVolume.GetData", "CHCSII_Reports", vbExclamation
CleanUp:
    Exit Function
End Function

