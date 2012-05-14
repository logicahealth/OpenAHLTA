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
Attribute VB_Name = "RepScreenPRVvolumn"
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
    If Not GetData(oReportItem, oRec) Then GoTo CleanUp         ' data for provider
    moCollRecSet.Add oRec, "ScreenPRVvolumn"

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
                           oRec As ADODB.Recordset) As Boolean

    Dim sStmt As String
    Dim sTableA As String
    Dim sTablePeer As String
    Dim sTableB As String
    Dim sClinicCriteria As String
    Dim lCriteriaCount As Long
    
    On Error GoTo ErrHandler
    GetData = False

    Dim sDateFilter As String
    Dim sFilter As String
    
    '-- Get Filter Information for SQL
    '----------------------------------------------------------------------------
    sDateFilter = sDateFilter & " (START_DATE >= " & GetOracleStartDate(oReportItem.DateFrom) & vbCrLf & _
                " AND START_DATE <= " & GetOracleEndDate(oReportItem.DateTo) & ") "
    
    If oReportItem.CriteriaSel1.Count = 0 Then
        sClinicCriteria = ""
    Else
        sClinicCriteria = "ENTERING_LOCATION_NCID = " & oReportItem.CriteriaSel1.Item(1).NCID & " AND " & vbCrLf
    End If
    
    sTableA = "SELECT  /*+ INDEX (LAB LAB_RAD_ORDERS_X3) */  B.SERVICE_NAME, TO_CHAR(COUNT(*)) AS TOTAL, B.FACILITY_NCID " & vbCrLf & _
            "FROM SERVICE_TEMPLATES_DATA A, SERVICE_TEMPLATES_INFO B, LAB_RAD_ORDERS LAB" & vbCrLf & _
            "WHERE  A.SERVICE_ID = B.SERVICE_ID AND" & vbCrLf & _
                "B.FACILITY_NCID = LAB.FACILITY_NCID AND" & vbCrLf & _
                "A.SERVICE_NCID = LAB.UNIVERSAL_SERVICE_ID AND" & vbCrLf & sClinicCriteria & _
                "ORDERING_PROVIDER_NCID = '" & oReportItem.CriteriaSel2.Item(1).NCID & "' AND" & vbCrLf & _
                sDateFilter & vbCrLf & _
                "GROUP BY B.SERVICE_NAME, B.FACILITY_NCID" & vbCrLf
    
    sTablePeer = "SELECT  /*+ INDEX (LAB LAB_RAD_ORDERS_X3) */  B.SERVICE_NAME, TO_CHAR(COUNT(*)) AS TOTAL, ORDERING_PROVIDER_NCID " & vbCrLf & _
            "FROM SERVICE_TEMPLATES_DATA A, SERVICE_TEMPLATES_INFO B, LAB_RAD_ORDERS LAB" & vbCrLf & _
            "WHERE  A.SERVICE_ID = B.SERVICE_ID AND" & vbCrLf & _
                "B.FACILITY_NCID = LAB.FACILITY_NCID AND" & vbCrLf & _
                "A.SERVICE_NCID = LAB.UNIVERSAL_SERVICE_ID AND" & vbCrLf & sClinicCriteria & _
                "ORDERING_PROVIDER_NCID <> '" & oReportItem.CriteriaSel2.Item(1).NCID & "' AND" & vbCrLf & _
                sDateFilter & vbCrLf & _
                "GROUP BY B.SERVICE_NAME, ORDERING_PROVIDER_NCID" & vbCrLf
            
    sTableB = "SELECT SERVICE_NAME, ROUND(AVG(ALL TOTAL), 0) AS MEAN," & vbCrLf & _
                    "ROUND(STDDEV(ALL TOTAL), 0) AS STD_DEV" & vbCrLf & _
               "FROM (" & sTablePeer & ")" & vbCrLf & _
               "GROUP BY SERVICE_NAME" & vbCrLf
                
    sStmt = "SELECT G.SERVICE_NAME, NVL(E.TOTAL, '0') AS TOTAL," & vbCrLf & _
                "NVL(TO_CHAR(F.MEAN), '0') AS MEAN, NVL(TO_CHAR(F.STD_DEV), '0') AS STD_DEV" & vbCrLf & _
            "FROM (" & sTableA & ") E, (" & sTableB & ") F, SERVICE_TEMPLATES_INFO G" & vbCrLf & _
            "WHERE E.SERVICE_NAME(+) = G.SERVICE_NAME AND" & vbCrLf & _
                "F.SERVICE_NAME(+) = G.SERVICE_NAME AND" & vbCrLf & _
                "G.SERVICE_NAME IS NOT NULL AND" & vbCrLf & _
                "E.FACILITY_NCID(+) = G.FACILITY_NCID" & vbCrLf

    Set oRec = goDBDAS.OpenRecordset(sStmt)
    GetData = True
    GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepScreenPRVvolumn.GetData", "CHCSII_Reports", vbExclamation
CleanUp:
    Exit Function
End Function

