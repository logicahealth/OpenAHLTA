VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PrescriptionsByAge"
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
    moCollRecSet.Add oRec, "PrescriptionByAge"

'-- Create Report Object
'-----------------------------------------------------------------------------
    If Not CreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo CleanUp

    BuildReport = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "PrescriptionsByAge.BuildReport", "CHCSII_Reports", vbExclamation
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
    
    On Error GoTo ErrHandler
    
    GetData = False
    
    '-- Get Filter Information for SQL
    '----------------------------------------------------------------------------
    sStartDate = GetOracleStartDate(GetFirstDayOfMonth(oReportItem.DateFrom))
    sEndDate = GetOracleEndDate(GetLastDayOfMonth(oReportItem.DateTo))
    sDateFilter = sDateFilter & " (START_DATE >= " & sStartDate & vbCrLf & " AND START_DATE <= " & sEndDate & ") "
    
    sPrv = "SELECT AGE_GROUP, ROUND(COUNT(ALL UNIT_NUMBER)/COUNT(DISTINCT UNIT_NUMBER), 0) AS AVG_PRESCRIPTIONS" & vbCrLf & _
            "FROM (" & vbCrLf & _
                "SELECT ID.UNIT_NUMBER, " & vbCrLf & _
                    "TRUNC(MONTHS_BETWEEN(SYSDATE, ID.BIRTH_DATE) / 12, 0) AS AGE," & vbCrLf & _
                    "DECODE(LENGTH(TO_CHAR(TRUNC(TRUNC(MONTHS_BETWEEN(SYSDATE, ID.BIRTH_DATE) / 12, 0)/5) * 5)), 1, " & vbCrLf & _
                    "'  ' || TO_CHAR(TRUNC(TRUNC(MONTHS_BETWEEN(SYSDATE, ID.BIRTH_DATE) / 12, 0)/5) * 5) " & vbCrLf & _
                    "|| ' to ' || TO_CHAR(TRUNC(TRUNC(MONTHS_BETWEEN(SYSDATE, ID.BIRTH_DATE) / 12, 0)/5) * 5 + 4), 2," & vbCrLf & _
                    "' ' || TO_CHAR(TRUNC(TRUNC(MONTHS_BETWEEN(SYSDATE, ID.BIRTH_DATE) / 12, 0)/5) * 5) " & vbCrLf & _
                    "|| ' to ' || TO_CHAR(TRUNC(TRUNC(MONTHS_BETWEEN(SYSDATE, ID.BIRTH_DATE) / 12, 0)/5) * 5 + 4), " & vbCrLf & _
                    "TO_CHAR(TRUNC(TRUNC(MONTHS_BETWEEN(SYSDATE, ID.BIRTH_DATE) / 12, 0)/5) * 5) " & vbCrLf & _
                    "|| ' to ' || TO_CHAR(TRUNC(TRUNC(MONTHS_BETWEEN(SYSDATE, ID.BIRTH_DATE) / 12, 0)/5) * 5 + 4)    " & vbCrLf & _
                    ") AS AGE_GROUP " & vbCrLf & _
                "FROM MEDICATION_ORDERS M, MEDICATION_FILLS F, MMI_ID ID" & vbCrLf & _
                "WHERE M.ORDERING_PROVIDER_NCID = " & oReportItem.CriteriaSel2.Item(1).NCID & " AND " & vbCrLf & _
                    "M.ENTERING_LOCATION_NCID = " & oReportItem.CriteriaSel1.Item(1).NCID & " AND" & vbCrLf & _
                    "M.ORDER_NUMBER = F.ORDER_NUMBER AND" & vbCrLf & _
                    "M.UNIT_NUMBER = ID.UNIT_NUMBER AND" & vbCrLf & sDateFilter & vbCrLf & _
            ")" & vbCrLf & _
            "GROUP BY AGE_GROUP" & vbCrLf
            
    sPeer = "SELECT AGE_GROUP, ROUND(AVG(ALL AVG_PRESCRIPTIONS), 1) AS PEER_MEAN, ROUND(STDDEV(ALL AVG_PRESCRIPTIONS), 1) AS STD" & vbCrLf & _
            "FROM (" & vbCrLf & _
                "SELECT AGE_GROUP, ORDERING_PROVIDER_NCID,  ROUND(COUNT(ALL UNIT_NUMBER)/COUNT(DISTINCT UNIT_NUMBER), 0) AS AVG_PRESCRIPTIONS" & vbCrLf & _
                "FROM (" & vbCrLf & _
                    "SELECT M.ORDERING_PROVIDER_NCID, ID.UNIT_NUMBER, " & vbCrLf & _
                        "TRUNC(MONTHS_BETWEEN(SYSDATE, ID.BIRTH_DATE) / 12, 0) AS AGE," & vbCrLf & _
                        "DECODE(LENGTH(TO_CHAR(TRUNC(TRUNC(MONTHS_BETWEEN(SYSDATE, ID.BIRTH_DATE) / 12, 0)/5) * 5)), 1, " & vbCrLf & _
                        "'  ' || TO_CHAR(TRUNC(TRUNC(MONTHS_BETWEEN(SYSDATE, ID.BIRTH_DATE) / 12, 0)/5) * 5) " & vbCrLf & _
                        "|| ' to ' || TO_CHAR(TRUNC(TRUNC(MONTHS_BETWEEN(SYSDATE, ID.BIRTH_DATE) / 12, 0)/5) * 5 + 4), 2," & vbCrLf & _
                        "' ' || TO_CHAR(TRUNC(TRUNC(MONTHS_BETWEEN(SYSDATE, ID.BIRTH_DATE) / 12, 0)/5) * 5) " & vbCrLf & _
                        "|| ' to ' || TO_CHAR(TRUNC(TRUNC(MONTHS_BETWEEN(SYSDATE, ID.BIRTH_DATE) / 12, 0)/5) * 5 + 4), " & vbCrLf & _
                        "TO_CHAR(TRUNC(TRUNC(MONTHS_BETWEEN(SYSDATE, ID.BIRTH_DATE) / 12, 0)/5) * 5) " & vbCrLf & _
                        "|| ' to ' || TO_CHAR(TRUNC(TRUNC(MONTHS_BETWEEN(SYSDATE, ID.BIRTH_DATE) / 12, 0)/5) * 5 + 4)    " & vbCrLf & _
                        ") AS AGE_GROUP " & vbCrLf & _
                    "FROM MEDICATION_ORDERS M, MEDICATION_FILLS F, MMI_ID ID" & vbCrLf & _
                    "WHERE M.ORDERING_PROVIDER_NCID <> " & oReportItem.CriteriaSel2.Item(1).NCID & " AND " & vbCrLf & _
                        "M.ENTERING_LOCATION_NCID = " & oReportItem.CriteriaSel1.Item(1).NCID & " AND" & vbCrLf & _
                        "M.ORDER_NUMBER = F.ORDER_NUMBER AND" & vbCrLf & _
                        "M.UNIT_NUMBER = ID.UNIT_NUMBER AND" & vbCrLf & sDateFilter & vbCrLf & _
                ")" & vbCrLf & _
                "GROUP BY AGE_GROUP, ORDERING_PROVIDER_NCID " & vbCrLf & _
            ")" & vbCrLf & _
            "GROUP BY AGE_GROUP" & vbCrLf
            
    sStmt = "SELECT A.AGE_GROUP, TO_CHAR(A.AVG_PRESCRIPTIONS) AS AVG_PRESCRIP, " & vbCrLf & _
                "TO_CHAR(NVL(B.PEER_MEAN, 0)) AS PEER_MEAN, TO_CHAR(NVL(B.STD, 0)) AS STD" & vbCrLf & _
            "FROM (" & sPrv & ") A, (" & sPeer & ") B" & vbCrLf & _
            "WHERE A.AGE_GROUP = B.AGE_GROUP (+)" & vbCrLf
            
    Set oRec = goDBDAS.OpenRecordset(sStmt)
    GetData = True
    
    If IsDate(oReportItem.DateFrom) Then strDateInfo = "From  " & Format(oReportItem.DateFrom, "mmm yyyy") & "   "
    If IsDate(oReportItem.DateTo) Then strDateInfo = strDateInfo & "Through  " & Format(oReportItem.DateTo, "mmm yyyy")
       
    On Error Resume Next
    oCollFormulas.Remove "DateInfo"
    oCollFormulas.Add CrystalQuote(strDateInfo), "DateInfo"
    
    GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "PrescriptionsByAge.GetData", "CHCSII_Reports", vbExclamation
CleanUp:
    Exit Function
End Function

