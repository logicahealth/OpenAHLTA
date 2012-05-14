VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RepReferralMTFtoMCSC"
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
    moCollRecSet.Add oRec, "ReferralMTFtoMCSC"

'-- Create Report Object
'-----------------------------------------------------------------------------
    If Not CreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo CleanUp

    BuildReport = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepRefferalMTFtoMCSC.BuildReport", "CHCSII_Reports", vbExclamation
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
    Dim sStartDate As String
    Dim sEndDate As String
    Dim sHeader As String
    Dim iResponse As Integer
    Dim strDateInfo As String
    
    '-- Get Filter Information for SQL
    '----------------------------------------------------------------------------
    sStartDate = GetOracleStartDate(GetFirstDayOfMonth(oReportItem.DateFrom))
    sEndDate = GetOracleEndDate(GetLastDayOfMonth(oReportItem.DateTo))
    
    sDateFilter = sDateFilter & " (ORDER_DATE >= " & sStartDate & vbCrLf & _
                                " AND ORDER_DATE <= " & sEndDate & ") "
    
    If oReportItem.Criteria1 = MTF Then
        sFilter = " WHERE  ENC_FAC_NCID = " & oReportItem.CriteriaSel1.Item(1).NCID & " And " & vbCrLf
    ElseIf oReportItem.Criteria1 = MTFDMIS Then
        sFilter = " WHERE  RECV_CLINIC_NCID IN (" & GetClinicNCIDsFromDMIS(oReportItem.CriteriaSel1) & ") AND " & vbCrLf
    End If
    sFilter = sFilter & _
             "(CONSULTING_NETWORK = 1 " & vbCrLf & _
                "OR  (ROUTING_STATUS IN (20, 24) " & _
                     "OR (ROUTING_STATUS = 6 AND ORDER_STATUS = 1))" & _
             ") AND" & vbCrLf
    
    sStmt = "SELECT TO_CHAR(ORDER_DATE, 'MON YYYY') AS REF_DATE, TO_CHAR(COUNT(*)) AS TOT_REFERRAL " & vbCrLf _
        & "FROM CONSULT_ORDER" & vbCrLf
    sStmt = sStmt & sFilter & sDateFilter & vbCrLf _
        & "GROUP BY TO_CHAR(ORDER_DATE, 'MON YYYY')"
           
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
    
    '-- SCR 36338, Sherry Wang 5/29/2003
    '-- Reformat date, only  display month and year
    If IsDate(oReportItem.DateFrom) Then strDateInfo = "From   " & Format(oReportItem.DateFrom, "mmm yyyy") & "   "
    If IsDate(oReportItem.DateTo) Then strDateInfo = strDateInfo & "Through   " & Format(oReportItem.DateTo, "mmm yyyy")
       
    On Error Resume Next
    oCollFormulas.Remove "DateInfo"
    oCollFormulas.Add CrystalQuote(strDateInfo), "DateInfo"
    
    GetData = True
    GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepGenericClass.GetData", "CHCSII_Reports", vbExclamation
CleanUp:
    Exit Function
End Function

