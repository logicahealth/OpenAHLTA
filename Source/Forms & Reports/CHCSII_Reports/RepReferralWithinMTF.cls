VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RepReferralWithinMTF"
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
    DoEvents
    
    '-- Get Data
    '----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection

    '-- Get Data for the MTF
    '----------------------------------------------------------
    If Not GetData(oReportItem, oRec, oCollFormulas) Then GoTo CleanUp
    moCollRecSet.Add oRec, "ReferralWithinMTF"

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
                           oRec As ADODB.Recordset, _
                           oCollFormulas As Collection) As Boolean

    Dim sStmt As String
    
    On Error GoTo ErrHandler
    GetData = False
    Dim sDateFilter As String
    Dim sFilter As String
    Dim sQuestionnaireFilter As String
    Dim sHeader As String
    Dim iResponse As Integer
    Dim sHint As String
    
    
    
    '-- Get Filter Information for SQL
    '----------------------------------------------------------------------------
    sDateFilter = sDateFilter & " (A.ORDER_DATE >= " & GetOracleStartDate(oReportItem.DateFrom) & vbCrLf & _
                                " AND A.ORDER_DATE <= " & GetOracleEndDate(oReportItem.DateTo) & ") "
    
    If oReportItem.Criteria1 = MTF Then
        sFilter = " A.ENC_FAC_NCID = " & oReportItem.CriteriaSel1.Item(1).NCID & " And " & vbCrLf
        sHint = " /*+ RULE */ /*+ ORDERED */ /*+ INDEX (A CONSULT_ORDER_X7) */ "
    ElseIf oReportItem.Criteria1 = MTFDMIS Then
        sFilter = " A.RECV_CLINIC_NCID IN (" & GetClinicNCIDsFromDMIS(oReportItem.CriteriaSel1) & ") AND " & vbCrLf
        sHint = " /*+ RULE */ /*+ ORDERED */ /*+ INDEX (A CONSULT_ORDER_X8) (B CLINIC_PK) */ "
    End If
    
    '--- Begin:SCR #24115;   Developer: Sherry Wang 05/23/2002 11:20 AM
    'SCR Description: Replaced recv_entity_text & recv_entity_ncid  with recv_clinic_ncid & recv_provider_ncid
    sFilter = " WHERE A.CONSULTING_NETWORK = 0 AND" & vbCrLf & sFilter & _
                    "A.RECV_CLINIC_NCID IS NOT NULL AND " & vbCrLf
    
    sStmt = "SELECT " & sHint & vbCrLf & " TO_CHAR(A.ORDER_DATE, 'MON YYYY') AS REF_DATE, TO_CHAR(COUNT(*)) AS TOT_REFERRAL " & vbCrLf & _
            "FROM CONSULT_ORDER A, CLINIC B" & vbCrLf & _
            sFilter & sDateFilter & vbCrLf & " AND " & _
                "A.RECV_CLINIC_NCID = B.NCID AND B.FACILITY_NCID = A.ENC_FAC_NCID" & vbCrLf & _
            "GROUP BY TO_CHAR(A.ORDER_DATE, 'MON YYYY')"
    '--- End: SCR #24115;
           
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
    gobjshared.ShowVBError Err.Number, Err.Description, "RepGenericClass.GetData", "CHCSII_Reports", vbExclamation
CleanUp:
    Exit Function
End Function

