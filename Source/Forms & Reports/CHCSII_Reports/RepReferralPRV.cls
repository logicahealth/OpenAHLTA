VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RepReferralPRV"
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
    
    '*************************************************************************************************************************************************
    '*** THIS CLASS IS TO BE USED AS A TEMPLATE FOR NEW CLASSES OR TO PROVIDE A BLANK POPUP REPORT ***
    '*** WHEN THE APPLICATION IS RUN FOR ONE OF THE NEW POPULATION HEALTH REPORTS.                                ***
    '*************************************************************************************************************************************************
    
    Dim oCrystal As Crystal
    Dim oRec As ADODB.Recordset
    Dim oPeerRec As ADODB.Recordset
    Dim sSponsorUnitNumber As String
    Dim sUnitNumber As String
    Dim strDateInfo As String
    DoEvents
    
    '-- Get Data
    '----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection

    '-- Get Data for the provider
    '----------------------------------------------------------
    If Not GetData(oReportItem, oRec, False) Then GoTo CleanUp
    moCollRecSet.Add oRec, "Referral"
    
    '-- Get data for Percentage Ranking
    '---------------------------------------------------------
    If Not GetData(oReportItem, oRec, True) Then GoTo CleanUp
    moCollRecSet.Add oRec, "PercentRank"
    
    '--- From To date info
    If IsDate(oReportItem.DateFrom) Then strDateInfo = "From " & Format(oReportItem.DateFrom, "dd mmm yyyy") & "   "
    If IsDate(oReportItem.DateTo) Then strDateInfo = strDateInfo & "Through " & Format(oReportItem.DateTo, "dd mmm yyyy")
       
    On Error Resume Next
    oCollFormulas.Remove "DateInfo"
    oCollFormulas.Add CrystalQuote(strDateInfo), "DateInfo"

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
                           oRec As ADODB.Recordset, PercentRank As Boolean) As Boolean

    Dim sStmt As String
    
    On Error GoTo ErrHandler
    GetData = False
    Dim sDateFilter As String
    Dim sFilter As String
    Dim sTablePRV As String
    Dim sTablePeer As String
    Dim sTableA As String
    Dim sTableA1 As String
    
    '-- Get Filter Information for SQL
    '----------------------------------------------------------------------------
    sDateFilter = sDateFilter & " (ORDER_DATE >= " & GetOracleStartDate(oReportItem.DateFrom) & vbCrLf & _
                                " AND ORDER_DATE <= " & GetOracleEndDate(oReportItem.DateTo) & ") "
    
    sFilter = " WHERE REF_PROVIDER_NCID = " & oReportItem.CriteriaSel2.Item(1).NCID & " AND " & vbCrLf
    
    If PercentRank Then
        Dim sTableB As String
        Dim sTableC As String
        Dim sTableD As String
        
        '--------------------------------------------------------------------------------
        ' Percent Rank = Sort Position / (Total Provider Count - 1)
        '
        ' Sort Position:  total number of providers whose total referrals are less than
        '                 current provider
        '--------------------------------------------------------------------------------
        
        sTableA = "SELECT REF_PROVIDER_NCID, COUNT(*) AS TOTAL_COUNT " & vbCrLf & _
                   "FROM CONSULT_ORDER" & vbCrLf & _
                   "WHERE ENC_FAC_NCID = (SELECT FACILITY_NCID FROM CLINIC WHERE IEN = '" & _
                        oReportItem.CriteriaSel1.Item(1).IEN & "') AND" & vbCrLf & _
                        sDateFilter & vbCrLf & _
                   "GROUP BY REF_PROVIDER_NCID"
                   
        '--- Total Provider Count - 1
        sTableB = "SELECT '1234' AS PROV, COUNT(*) - 1 AS TOTAL_PROV" & vbCrLf & _
                  "FROM (" & sTableA & ")"
                  
        '--- Total referrals for the current provider
        sTableC = "SELECT TOTAL_COUNT" & vbCrLf & _
                  "FROM (" & sTableA & ")" & vbCrLf & _
                  "WHERE REF_PROVIDER_NCID = " & oReportItem.CriteriaSel2.Item(1).NCID & vbCrLf
                  
        '--- Sort Position
        sTableD = "SELECT '1234' AS PROV, COUNT(*) AS POSITION" & vbCrLf & _
                  "FROM (" & sTableA & ")" & vbCrLf & _
                  "WHERE TOTAL_COUNT < (" & sTableC & ")" & vbCrLf

        sStmt = "SELECT TO_CHAR(ROUND(D.POSITION / B.TOTAL_PROV, 2) * 100) || '%' AS PERCENTRANK" & vbCrLf & _
                "FROM (" & sTableB & ") B, (" & sTableD & ") D" & vbCrLf & _
                "WHERE B.PROV = D.PROV" & vbCrLf & _
                "AND B.TOTAL_PROV <> 0" & vbCrLf

    Else
        sTableA = "SELECT RECV_CLINIC_NCID," & vbCrLf & _
                        "CONSULTING_NETWORK, " & vbCrLf & _
                        "REF_CLINIC_NCID as FROM_CLINIC_NCID " & vbCrLf & _
                   "FROM CONSULT_ORDER" & vbCrLf & _
                    sFilter & sDateFilter '_
                    '"RECV_ENTITY_NCID IS NOT NULL AND" & vbCrLf & sDateFilter
    
        sTableA1 = "SELECT RECV_CLINIC_NCID," & vbCrLf & _
                        "CONSULTING_NETWORK, REF_PROVIDER_NCID " & vbCrLf & _
                   "FROM CONSULT_ORDER" & vbCrLf & _
                    Replace(sFilter, "=", "<>") & _
                    "ENC_FAC_NCID = (SELECT FACILITY_NCID FROM CLINIC WHERE IEN = '" & _
                        oReportItem.CriteriaSel1.Item(1).IEN & "') AND" & vbCrLf & sDateFilter '_
                    '"RECV_ENTITY_NCID IS NOT NULL AND" & vbCrLf & sDateFilter
    
     
        sTablePRV = "SELECT B.DMIS_ID_CODE, B.MEPRS_CODE, B.NAME," & vbCrLf & _
                        "A.CONSULTING_NETWORK, TO_CHAR(COUNT(*)) AS TOT," & vbCrLf & _
                        "C.DMIS_ID_CODE AS FROM_DMIS_ID_CODE, C.MEPRS_CODE AS FROM_MEPRS_CODE, C.NAME AS FROM_NAME" & vbCrLf & _
                    "FROM (" & sTableA & ") A, CLINIC B, CLINIC C" & vbCrLf & _
                    "WHERE A.RECV_CLINIC_NCID = B.NCID AND A.FROM_CLINIC_NCID = C.NCID(+)" & vbCrLf & _
                    "GROUP BY B.DMIS_ID_CODE, B.MEPRS_CODE, B.NAME, A.CONSULTING_NETWORK, " & vbCrLf & _
                        "C.DMIS_ID_CODE, C.MEPRS_CODE, C.NAME" & vbCrLf
                        
                    
        sTablePeer = "SELECT DMIS_ID_CODE, MEPRS_CODE, NAME, CONSULTING_NETWORK," & vbCrLf & _
                        "ROUND(AVG(ALL TOT1), 0) AS PEER_MEAN," & vbCrLf & _
                        "ROUND(STDDEV(ALL TOT1), 1) AS STD_DEV" & vbCrLf & _
                    "FROM (" & vbCrLf & _
                        "SELECT B1.DMIS_ID_CODE, B1.MEPRS_CODE, B1.NAME," & vbCrLf & _
                            "A1.CONSULTING_NETWORK, A1.REF_PROVIDER_NCID," & vbCrLf & _
                            "COUNT(*) AS TOT1" & vbCrLf & _
                        "FROM (" & sTableA1 & ") A1, CLINIC B1" & vbCrLf & _
                        "WHERE A1.RECV_CLINIC_NCID = B1.NCID" & vbCrLf & _
                        "GROUP BY B1.DMIS_ID_CODE, B1.MEPRS_CODE, B1.NAME, A1.CONSULTING_NETWORK, A1.REF_PROVIDER_NCID" & vbCrLf & _
                    ")" & vbCrLf & _
                    "GROUP BY DMIS_ID_CODE, MEPRS_CODE, NAME, CONSULTING_NETWORK" & vbCrLf
                    
                    
        sStmt = "SELECT PVD.DMIS_ID_CODE, PVD.MEPRS_CODE, PVD.NAME," & vbCrLf & _
                    "PVD.CONSULTING_NETWORK, TO_CHAR(PVD.TOT) AS TOT_REF," & vbCrLf & _
                    "TO_CHAR(NVL(PEER.PEER_MEAN, 0)) AS MEAN," & vbCrLf & _
                    "TO_CHAR(NVL(PEER.STD_DEV, 0)) AS STDDEV," & vbCrLf & _
                    "PVD.FROM_DMIS_ID_CODE, PVD.FROM_MEPRS_CODE, PVD.FROM_NAME" & vbCrLf & _
                "FROM (" & sTablePRV & ") PVD, (" & sTablePeer & ") PEER" & vbCrLf & _
                "WHERE PVD.DMIS_ID_CODE = PEER.DMIS_ID_CODE(+) " & vbCrLf & _
                    "AND PVD.MEPRS_CODE = PEER.MEPRS_CODE(+)" & vbCrLf & _
                    "AND PVD.NAME = PEER.NAME(+) " & vbCrLf & _
                    "AND PVD.CONSULTING_NETWORK = 0" & vbCrLf & _
                    "AND PVD.CONSULTING_NETWORK = PEER.CONSULTING_NETWORK(+)" & vbCrLf
                    
    End If
           
    Set oRec = goDBDAS.OpenRecordset(sStmt)
    GetData = True
    GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepGenericClass.GetData", "CHCSII_Reports", vbExclamation
CleanUp:
    Exit Function
End Function

