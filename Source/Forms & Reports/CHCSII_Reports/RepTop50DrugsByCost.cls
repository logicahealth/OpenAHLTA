VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RepTop50DrugsByCost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private moCollRecSet As Collection
Private moCollFormulas As Collection

Private Type SCREEN_SERVICE_RS
    PER_DATE                As String
    SERVICE                 As String
    UNIVERSAL_SERVICE_NCID  As Long
    TOTAL                   As String
End Type


Public Function BuildReport(sReportName As String, _
                            oReportItem As ReportItem, _
                            oReport As CRAXDRT.Report, _
                            oCollFormulas As Collection) As Boolean
 
    On Error GoTo ErrHandler
     
    BuildReport = False
    
    Dim oCrystal As Crystal
    Dim oRec1 As ADODB.Recordset
    Dim oRec2 As ADODB.Recordset
    DoEvents
    
    '-- Get Data
    '----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection

    '-- Get Candidates
    '----------------------------------------------------------
    If Not GetData(oReportItem, oRec1, oRec2, oCollFormulas) Then GoTo CleanUp         ' data for provider
    moCollRecSet.Add oRec1, "Top50DrugsByCostGraphMTF"
    moCollRecSet.Add oRec2, "Top50DrugsByCostGraphMTF2"

'-- Create Report Object
'-----------------------------------------------------------------------------
    If Not CreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo CleanUp

    BuildReport = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepTop50DrugsByVolume.BuildReport", "CHCSII_Reports", vbExclamation
CleanUp:
    Set oCrystal = Nothing
    Set oRec1 = Nothing
    Set oRec2 = Nothing
    
End Function

Private Function GetData(oReportItem As ReportItem, _
                           oRec1 As ADODB.Recordset, _
                           oRec2 As ADODB.Recordset, _
                           oCollFormulas As Collection) As Boolean

    Dim sStmt As String
    Dim sFilter As String
    Dim oRec As ADODB.Recordset
    Dim sHeader As String
    Dim iResponse As Integer
    
    Dim sStartDate As String
    Dim sEndDate As String
    Dim sDateFilter As String
    Dim strDateInfo As String
    Dim sHint As String
    
    On Error GoTo ErrHandler
    
    GetData = False
    
    '-- Get Filter Information for SQL
    '----------------------------------------------------------------------------
    sStartDate = GetOracleStartDate(GetFirstDayOfMonth(oReportItem.DateFrom))
    sEndDate = GetOracleEndDate(GetLastDayOfMonth(oReportItem.DateTo))
    sDateFilter = sDateFilter & " (START_DATE >= " & sStartDate & vbCrLf & " AND START_DATE <= " & sEndDate & ") "
    
    If oReportItem.Criteria1 = MTF Then
        sFilter = " M.FACILITY_NCID = F.FACILITY_NCID AND" & vbCrLf & _
                    "F.FACILITY_NCID = " & oReportItem.CriteriaSel1.Item(1).NCID & " AND" & vbCrLf
        sHint = " /*+ ORDERED */ /*+ INDEX  (M MEDICATION_ORDERS_X4)  (F MEDICATION_FILLS_PK)  (D DRUG_COST_DATA_X1)  */"
    ElseIf oReportItem.Criteria1 = MTFDMIS Then
        sFilter = " M.ENTERING_LOCATION_NCID IN (" & GetClinicNCIDsFromDMIS(oReportItem.CriteriaSel1) & ") AND " & vbCrLf
        sHint = " /*+ ORDERED */ /*+ INDEX  (M MEDICATION_ORDERS_X5)  (F MEDICATION_FILLS_PK)  (D DRUG_COST_DATA_X1)  */"
    End If
    
    sStmt = "SELECT /*+ RULE */ TO_CHAR(ROWNUM) AS SEQ, GC3_NAME, TO_CHAR(TOTAL_COST) AS TOTAL_COST" & vbCrLf & _
            "FROM (" & vbCrLf & _
                "SELECT " & sHint & vbCrLf & " D.GC3_NAME, ROUND(SUM(D.LOWEST_UM_PRICE * M.REQUESTED_DISPENSE_AMOUNT), 2) AS TOTAL_COST" & vbCrLf & _
                "FROM MEDICATION_ORDERS M, MEDICATION_FILLS F, DRUG_COST_DATA D" & vbCrLf & _
                "WHERE M.ORDER_NUMBER = F.ORDER_NUMBER AND" & vbCrLf & sFilter & _
                    "M.DRUG_ALTERNATE_IDENTIFIER = D.NDC AND" & vbCrLf & _
                    "M.REQUESTED_DISPENSE_AMOUNT IS NOT NULL AND" & vbCrLf & _
                    "D.LOWEST_UM_PRICE IS NOT NULL AND D.GC3_NAME IS NOT NULL AND" & vbCrLf & sDateFilter & vbCrLf & _
                "GROUP BY D.GC3_NAME, D.NDC, D.DRUG_CLASS" & vbCrLf & _
                "ORDER BY SUM(D.LOWEST_UM_PRICE * M.REQUESTED_DISPENSE_AMOUNT) DESC, D.GC3_NAME ASC" & vbCrLf & _
            ")" & vbCrLf & _
            "WHERE ROWNUM < 51" & vbCrLf
            
    Set oRec = goDBDAS.OpenRecordset(sStmt)
    
    If Not SplitRecordSet(oRec, oRec1, oRec2) Then GoTo CleanUp
        
    GetData = True
    
    If IsDate(oReportItem.DateFrom) Then strDateInfo = "From  " & Format(oReportItem.DateFrom, "mmm yyyy") & "   "
    If IsDate(oReportItem.DateTo) Then strDateInfo = strDateInfo & "Through  " & Format(oReportItem.DateTo, "mmm yyyy")
       
    On Error Resume Next
    oCollFormulas.Remove "DateInfo"
    oCollFormulas.Add CrystalQuote(strDateInfo), "DateInfo"
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
    
    GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepTop50DrugsByCost.GetData", "CHCSII_Reports", vbExclamation
CleanUp:
    oRec.Close
    Set oRec = Nothing
    Exit Function
End Function

'--- Split the top 50 drugs recordset to 2 recordsets.
'--- Save top 1 to 25 drugs to the first recordset
'--- and top 26 to 50 drugs to the second recordset
Private Function SplitRecordSet(oRS As ADODB.Recordset, oRec1 As ADODB.Recordset, oRec2 As ADODB.Recordset) As Boolean
    Dim i As Integer
    Dim sFilter As String
    
    SplitRecordSet = False
    
    '--- open oRec1
    Set oRec1 = New ADODB.Recordset
    oRec1.CursorLocation = adUseClient
    oRec1.CursorType = adOpenDynamic
    Set oRec1.ActiveConnection = Nothing
    oRec1.Fields.Append "SEQ", adChar, 15
    oRec1.Fields.Append "GC3_NAME", adChar, 40
    oRec1.Fields.Append "TOTAL_COST", adChar, 40
    oRec1.Open

    '--- open oRec2
    Set oRec2 = New ADODB.Recordset
    oRec2.CursorLocation = adUseClient
    oRec2.CursorType = adOpenDynamic
    Set oRec2.ActiveConnection = Nothing
    oRec2.Fields.Append "SEQ", adChar, 15
    oRec2.Fields.Append "GC3_NAME", adChar, 40
    oRec2.Fields.Append "TOTAL_COST", adChar, 40
    oRec2.Open

    Do Until oRS.EOF
        If CDbl(oRS("SEQ")) < 26 Then
            oRec1.AddNew
                oRec1("SEQ") = oRS("SEQ")
                oRec1("GC3_NAME") = oRS("GC3_NAME")
                oRec1("TOTAL_COST") = oRS("TOTAL_COST")
            oRec1.Update
        Else
            oRec2.AddNew
                oRec2("SEQ") = oRS("SEQ")
                oRec2("GC3_NAME") = oRS("GC3_NAME")
                oRec2("TOTAL_COST") = oRS("TOTAL_COST")
            oRec2.Update
        End If
        oRS.MoveNext
    Loop
    
    SplitRecordSet = True
    
    GoTo CleanUp
ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepTop50DrugsByCost.SplitRecordSet", "CHCSII_Reports", vbExclamation
CleanUp:
    Exit Function
End Function






