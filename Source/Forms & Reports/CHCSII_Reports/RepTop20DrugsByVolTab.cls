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
Attribute VB_Name = "RepTop20DrugsByVolTab"
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
    moCollRecSet.Add oRec, "Top20byVolumeTabular"
    
    If oReportItem.SecondaryFuncSelected = True Then
       If Not GetSubData(oReportItem, oRec, oCollFormulas) Then GoTo CleanUp         ' data for provider
    Else
       Set oRec = New ADODB.Recordset
       With oRec.Fields
            .Append "NDC", adLongVarWChar, -1, 132
            .Append "Seq", adLongVarWChar, -1, 132
            .Append "GC3_NAME", adLongVarWChar, -1, 132
            .Append "DRUG_CLASS", adLongVarWChar, -1, 132
            .Append "NNUM_PATIENT", adLongVarWChar, -1, 132
            .Append "TOTAL_COST", adLongVarWChar, -1, 132
            .Append "AVG_COST", adLongVarWChar, -1, 132
            .Append "Volume", adLongVarWChar, -1, 132
       End With
       oRec.Open
    End If
    moCollRecSet.Add oRec, "Top20byVolTabular"


'-- Create Report Object
'-----------------------------------------------------------------------------
    If Not CreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo CleanUp

    BuildReport = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepTop20DrugsByVolTab.BuildReport", "CHCSII_Reports", vbExclamation
CleanUp:
    Set oCrystal = Nothing
    Set oRec = Nothing
    Set oPeerRec = Nothing
    
End Function

Private Function GetSubData(oReportItem As ReportItem, _
                           oRec As ADODB.Recordset, _
                           oCollFormulas As Collection) As Boolean

    Dim sStmt As String
    Dim sFilter As String
    
    Dim sStartDate As String
    Dim sEndDate As String
    Dim sDateFilter As String
    Dim strDateInfo As String
    Dim oHighlight As MTFTop50RankDrugs
    Dim sHint As String
    Dim oRec2 As ADODB.Recordset
    Dim oRecUnordered As ADODB.Recordset
    On Error GoTo ErrHandler
    
    GetSubData = False
    
    Set oRec = New ADODB.Recordset
    With oRec.Fields
           .Append "NDC", adLongVarWChar, -1, 132
           .Append "Seq", adLongVarWChar, -1, 132
           .Append "GC3_NAME", adLongVarWChar, -1, 132
           .Append "DRUG_CLASS", adLongVarWChar, -1, 132
           .Append "NNUM_PATIENT", adLongVarWChar, -1, 132
           .Append "TOTAL_COST", adLongVarWChar, -1, 132
           .Append "AVG_COST", adLongVarWChar, -1, 132
           .Append "Volume", adLongVarWChar, -1, 132
    End With
    oRec.Open
    
    '-- Get Filter Information for SQL    '----------------------------------------------------------------------------
    sStartDate = GetOracleStartDate(GetFirstDayOfMonth(oReportItem.DateFrom))
    sEndDate = GetOracleEndDate(GetLastDayOfMonth(oReportItem.DateTo))
    sDateFilter = " (START_DATE >= " & sStartDate & vbCrLf & " AND START_DATE <= " & sEndDate & ") "
    
    If oReportItem.Criteria1 = MTF Then
        sFilter = "M.FACILITY_NCID = F.FACILITY_NCID AND" & vbCrLf
        sHint = " /*+ ORDERED */ /*+ INDEX  (M MEDICATION_ORDERS_X4)  (F MEDICATION_FILLS_PK)  (D DRUG_COST_DATA_X1)  */"
    ElseIf oReportItem.Criteria1 = MTFDMIS Then
        sFilter = " M.MEPRS_CODE IN (" & GetMEPRSsFromDMIS(oReportItem.CriteriaSel1) & ") AND " & vbCrLf
        sFilter = " M.ENTERING_LOCATION_NCID IN (" & GetClinicNCIDsFromDMIS(oReportItem.CriteriaSel1) & ") AND " & vbCrLf
        sHint = " /*+ ORDERED */ /*+ INDEX  (M MEDICATION_ORDERS_X5)  (F MEDICATION_FILLS_PK)  (D DRUG_COST_DATA_X1)  */"
    End If
    
    sStmt = "SELECT /*+ RULE */ NDC, TO_CHAR(NUM_PATIENT) AS NNUM_PATIENT," & vbCrLf & _
        "TO_CHAR(TOTAL_DISPENSED) AS TOTAL_DISPENSED, TO_CHAR(VOLUME) AS VOLUME" & vbCrLf & _
        "FROM (SELECT " & sHint & vbCrLf & _
        "COUNT(DISTINCT M.UNIT_NUMBER) AS NUM_PATIENT," & vbCrLf & _
        "ROUND(SUM(M.REQUESTED_DISPENSE_AMOUNT), 2) AS TOTAL_DISPENSED," & vbCrLf & _
        "COUNT(*) AS VOLUME, M.DRUG_ALTERNATE_IDENTIFIER AS NDC" & vbCrLf & _
        "FROM MEDICATION_ORDERS M, MEDICATION_FILLS F" & vbCrLf & _
        "WHERE M.ORDER_NUMBER = F.ORDER_NUMBER AND M.FACILITY_NCID = F.FACILITY_NCID AND" & vbCrLf & _
        "M.ORDERING_PROVIDER_NCID = " & oReportItem.CriteriaSel2.Item(1).NCID & " AND" & vbCrLf
        If oReportItem.Criteria1 = MTF Then
           sStmt = sStmt & "M.ENTERING_LOCATION_NCID = " & oReportItem.CriteriaSel1.Item(1).NCID & " AND" & vbCrLf
        ElseIf oReportItem.Criteria1 = MTFDMIS Then
            sStmt = sStmt & " M.ENTERING_LOCATION_NCID IN (" & GetClinicNCIDsFromDMIS(oReportItem.CriteriaSel1) & ") AND " & vbCrLf
        End If
        sStmt = sStmt & "M.REQUESTED_DISPENSE_AMOUNT IS NOT NULL AND" & vbCrLf & _
         sDateFilter & vbCrLf
    If oReportItem.Criteria1 = MTF Then
        sStmt = sStmt & "AND M.Facility_NCID = " & oReportItem.CriteriaSel1.Item(1).NCID & vbCrLf
    ElseIf oReportItem.Criteria1 = MTFDMIS Then
        sStmt = sStmt & "AND M.Facility_NCID = " & GetFacilityNCIDFromDMIS(oReportItem.CriteriaSel1) & vbCrLf
    End If
    sStmt = sStmt & "GROUP BY M.DRUG_ALTERNATE_IDENTIFIER)"

    Set oRec2 = goDBDAS.OpenRecordset(sStmt)
    
    sStmt = "SELECT  /*+ ORDERED */ /*+ INDEX  (D DRUG_COST_DATA_X1)  */" & vbCrLf & _
                 "D.GC3_NAME, D.DRUG_CLASS, '' AS NUM_PATIENT, D.LOWEST_UM_PRICE AS TOTAL_COST, D.LOWEST_UM_PRICE AS AVG_COST," & vbCrLf & _
                 "'' AS VOLUME, A.DRUG_NDC AS NDC" & vbCrLf & _
                "FROM DRUG_COST_DATA D, DRUG_RANKING_ALTERNATIVE A" & vbCrLf & _
                "WHERE D.NDC = A.ALTERNATIVE_DRUG_NDC AND" & vbCrLf
    
    If oReportItem.Criteria1 = MTF Then
        sStmt = sStmt & "A.Facility_NCID = " & oReportItem.CriteriaSel1.Item(1).NCID & " AND" & vbCrLf
    ElseIf oReportItem.Criteria1 = MTFDMIS Then
        sStmt = sStmt & "A.Facility_NCID = " & GetFacilityNCIDFromDMIS(oReportItem.CriteriaSel1) & " AND " & vbCrLf
    End If
    sStmt = sStmt & "D.GC3_NAME IS NOT NULL" & vbCrLf & _
                "ORDER BY D.GC3_NAME ASC"
    
    Set oRecUnordered = goDBDAS.OpenRecordset(sStmt)
    If Not (oRecUnordered.BOF And oRecUnordered.EOF) Then
        oRecUnordered.MoveFirst
        Do Until oRecUnordered.EOF
            oRec.AddNew
            oRec("GC3_NAME") = oRecUnordered("GC3_NAME")
            oRec("DRUG_CLASS") = oRecUnordered("DRUG_CLASS")
            oRec("TOTAL_COST") = Format(oRecUnordered("TOTAL_COST"), "Fixed")
            oRec("AVG_COST") = Format(oRecUnordered("AVG_COST"), "Fixed")
            oRec("NDC") = oRecUnordered("NDC")
            If Not (oRec2.BOF And oRec2.EOF) Then 'Prevents errors if orec2 has no records
                oRec2.Filter = "NDC = '" & oRecUnordered("NDC") & "'"
                If Not (oRec2.BOF And oRec2.EOF) Then
                    oRec2.MoveFirst
                    oRec("NNUM_PATIENT") = oRec2("NNUM_PATIENT")
                    oRec("VOLUME") = oRec2("VOLUME")
                    oRec("TOTAL_COST") = Format(Round(oRec2("TOTAL_DISPENSED") * oRecUnordered("AVG_COST"), 2), "Fixed")
                    oRec("AVG_COST") = Format(CStr(Round(oRec("TOTAL_COST") / oRec2("VOLUME"), 2)), "Fixed")
                End If
            End If
            oRec2.Filter = ""
            oRecUnordered.MoveNext
        Loop
    End If
    
    '--- Highlight the drugs that are in the top 50 highest cost drugs by MTF
    Set oHighlight = New MTFTop50RankDrugs
    Set oRec = oHighlight.HighlightClinicDrugs(oRec, oReportItem.CriteriaSel1.Item(1).NCID)
    GetSubData = True
    
    GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepTop20DrugsByCostTab.GetSubData", "CHCSII_Reports", vbExclamation
CleanUp:
    Set oHighlight = Nothing
    Set oRec2 = Nothing
    Set oRecUnordered = Nothing
    
End Function


Private Function GetData(oReportItem As ReportItem, _
                           oRec As ADODB.Recordset, _
                           oCollFormulas As Collection) As Boolean

    Dim sStmt As String
    Dim sFilter As String
    
    Dim sStartDate As String
    Dim sEndDate As String
    Dim sDateFilter As String
    Dim strDateInfo As String
    Dim oHighlight As MTFTop50RankDrugs
    Dim sHint As String
    
    On Error GoTo ErrHandler
    
    GetData = False
    
    sHint = " /*+ ORDERED */ /*+ INDEX  (M MEDICATION_ORDERS_X5)  (F MEDICATION_FILLS_PK)  (D DRUG_COST_DATA_X1)  */"
    
    '-- Get Filter Information for SQL
    '----------------------------------------------------------------------------
    sStartDate = GetOracleStartDate(GetFirstDayOfMonth(oReportItem.DateFrom))
    sEndDate = GetOracleEndDate(GetLastDayOfMonth(oReportItem.DateTo))
    sDateFilter = sDateFilter & " (START_DATE >= " & sStartDate & vbCrLf & " AND START_DATE <= " & sEndDate & ") "
    
    sStmt = "SELECT  /*+ RULE */  NDC, TO_CHAR(ROWNUM) AS SEQ, GC3_NAME, DRUG_CLASS, TO_CHAR(NUM_PATIENT) AS NNUM_PATIENT, " & vbCrLf & _
                "TO_CHAR(TOTAL_COST) AS TOTAL_COST, TO_CHAR(AVG_COST) AS AVG_COST, TO_CHAR(VOLUME) AS VOLUME" & vbCrLf & _
            "FROM (" & vbCrLf & _
                "SELECT " & sHint & vbCrLf & " D.GC3_NAME, D.DRUG_CLASS, COUNT(DISTINCT M.UNIT_NUMBER) AS NUM_PATIENT," & vbCrLf & _
                    "ROUND(SUM(D.LOWEST_UM_PRICE * M.REQUESTED_DISPENSE_AMOUNT), 2) AS TOTAL_COST," & vbCrLf & _
                    "ROUND(AVG(D.LOWEST_UM_PRICE * M.REQUESTED_DISPENSE_AMOUNT), 2) AS AVG_COST," & vbCrLf & _
                    "COUNT(*) AS VOLUME, D.NDC" & vbCrLf & _
                "FROM MEDICATION_ORDERS M, MEDICATION_FILLS F, DRUG_COST_DATA D" & vbCrLf & _
                "WHERE M.ORDERING_PROVIDER_NCID = " & oReportItem.CriteriaSel2.Item(1).NCID & " AND " & vbCrLf & _
                    "M.ENTERING_LOCATION_NCID = " & oReportItem.CriteriaSel1.Item(1).NCID & " AND" & vbCrLf & _
                    "M.ORDER_NUMBER = F.ORDER_NUMBER AND" & vbCrLf & _
                    "M.FACILITY_NCID = F.FACILITY_NCID AND" & vbCrLf & _
                    "M.DRUG_ALTERNATE_IDENTIFIER = D.NDC AND" & vbCrLf & _
                    "M.REQUESTED_DISPENSE_AMOUNT IS NOT NULL AND" & vbCrLf & _
                    "D.LOWEST_UM_PRICE IS NOT NULL AND D.GC3_NAME IS NOT NULL AND" & vbCrLf & sDateFilter & vbCrLf & _
                "GROUP BY D.GC3_NAME, D.NDC, D.DRUG_CLASS" & vbCrLf & _
                "ORDER BY COUNT(*) DESC, D.GC3_NAME ASC" & vbCrLf & _
            ")" & vbCrLf & _
            "WHERE ROWNUM < 21" & vbCrLf
            
'"M.MEPRS_CODE = (SELECT MEPRS_CODE FROM CLINIC WHERE NCID = " & oReportItem.CriteriaSel1.Item(1).NCID & ") AND " & vbCrLf
    If oReportItem.SortItemSelected <> "" Then
        sStmt = sStmt & "Order by " & oReportItem.SortItemSelected & " " & oReportItem.SortDirection & vbCrLf
    End If
            
    Set oRec = goDBDAS.OpenRecordset(sStmt)
    '--- Highlight the drugs that are in the top 50 highest cost drugs by MTF
    Set oHighlight = New MTFTop50RankDrugs
    Set oRec = oHighlight.HighlightClinicDrugs(oRec, oReportItem.CriteriaSel1.Item(1).NCID)
    GetData = True
    
    If IsDate(oReportItem.DateFrom) Then strDateInfo = "From  " & Format(oReportItem.DateFrom, "mmm yyyy") & "   "
    If IsDate(oReportItem.DateTo) Then strDateInfo = strDateInfo & "Through  " & Format(oReportItem.DateTo, "mmm yyyy")
       
    On Error Resume Next
    oCollFormulas.Remove "DateInfo"
    oCollFormulas.Add CrystalQuote(strDateInfo), "DateInfo"
    
    GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepTop20DrugsByVolTab.GetData", "CHCSII_Reports", vbExclamation
CleanUp:
    Set oHighlight = Nothing
    Exit Function
End Function

