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
Attribute VB_Name = "RepTop50DrugsByVolTab"
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
    Dim oRecUnordered As ADODB.Recordset
    Dim oRec2 As ADODB.Recordset
    
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
        "M.REQUESTED_DISPENSE_AMOUNT IS NOT NULL AND" & vbCrLf & _
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
            oRec("TOTAL_COST") = oRecUnordered("TOTAL_COST")
            oRec("AVG_COST") = oRecUnordered("AVG_COST")
            oRec("NDC") = oRecUnordered("NDC")
            If Not (oRec2.BOF And oRec2.EOF) Then
                oRec2.Filter = "NDC = '" & oRecUnordered("NDC") & "'"
                If Not (oRec2.BOF And oRec2.EOF) Then
                    oRec2.MoveFirst
                    oRec("NNUM_PATIENT") = oRec2("NNUM_PATIENT")
                    oRec("VOLUME") = oRec2("VOLUME")
                    oRec("TOTAL_COST") = Round(oRec2("TOTAL_DISPENSED") * oRecUnordered("AVG_COST"), 2)
                    oRec("AVG_COST") = CStr(Round(oRec("TOTAL_COST") / oRec2("VOLUME"), 2))
                End If
                oRec2.Filter = ""
            End If
            oRecUnordered.MoveNext
        Loop
    End If
    
    '--- Highlight the drugs that are in the top 50 highest cost drugs by MTF
    Set oHighlight = New MTFTop50RankDrugs
    
    If oReportItem.Criteria1 = MTF Then
        Set oRec = oHighlight.HighlightTop50MTFDrugs(oRec, oReportItem.CriteriaSel1.Item(1).NCID)
    ElseIf oReportItem.Criteria1 = MTFDMIS Then
        Set oRec = oHighlight.HighlightTop50MTFDrugs(oRec, GetFacilityNCIDFromDMIS(oReportItem.CriteriaSel1))
    End If
    GetSubData = True
    
    GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepTop20DrugsByCostTab.GetSubData", "CHCSII_Reports", vbExclamation
    'Resume Next
CleanUp:
    Set oHighlight = Nothing
    Exit Function
End Function

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
    moCollRecSet.Add oRec, "Top50byVolumeTabular"

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
    moCollRecSet.Add oRec, "Top50DrugsVolSubMTF"

'-- Create Report Object
'-----------------------------------------------------------------------------
    If Not CreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo CleanUp

    BuildReport = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepTop50DrugsByVolTab.BuildReport", "CHCSII_Reports", vbExclamation
CleanUp:
    Set oCrystal = Nothing
    Set oRec = Nothing
    Set oPeerRec = Nothing
    
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
    
    Dim sHeader As String
    Dim iResponse As Integer
    
    On Error GoTo ErrHandler
    
    GetData = False
    
    '-- Get Filter Information for SQL
    '----------------------------------------------------------------------------
    sStartDate = GetOracleStartDate(GetFirstDayOfMonth(oReportItem.DateFrom))
    sEndDate = GetOracleEndDate(GetLastDayOfMonth(oReportItem.DateTo))
    sDateFilter = " (START_DATE >= " & sStartDate & vbCrLf & " AND START_DATE <= " & sEndDate & ") "
    
    If oReportItem.Criteria1 = MTF Then
        sFilter = " M.FACILITY_NCID = F.FACILITY_NCID AND" & vbCrLf & _
                    "F.FACILITY_NCID = " & oReportItem.CriteriaSel1.Item(1).NCID & " AND" & vbCrLf
        sHint = " /*+ ORDERED */ /*+ INDEX  (M MEDICATION_ORDERS_X4)  (F MEDICATION_FILLS_PK)  (D DRUG_COST_DATA_X1)  */"
    ElseIf oReportItem.Criteria1 = MTFDMIS Then
        'sFilter = " M.MEPRS_CODE IN (" & GetMEPRSsFromDMIS(oReportItem.CriteriaSel1) & ") AND " & vbCrLf
        sFilter = " M.ENTERING_LOCATION_NCID IN (" & GetClinicNCIDsFromDMIS(oReportItem.CriteriaSel1) & ") AND " & vbCrLf
        sHint = " /*+ ORDERED */ /*+ INDEX  (M MEDICATION_ORDERS_X5)  (F MEDICATION_FILLS_PK)  (D DRUG_COST_DATA_X1)  */"
    End If
    
    sStmt = "SELECT  /*+ RULE */  TO_CHAR(ROWNUM) AS SEQ, GC3_NAME, DRUG_CLASS, TO_CHAR(NUM_PATIENT) AS NNUM_PATIENT, " & vbCrLf & _
                "TO_CHAR(TOTAL_COST) AS TOTAL_COST, TO_CHAR(AVG_COST) AS AVG_COST, TO_CHAR(VOLUME) AS VOLUME, NDC" & vbCrLf & _
            "FROM (" & vbCrLf & _
                "SELECT " & sHint & vbCrLf & " D.GC3_NAME, D.DRUG_CLASS, COUNT(DISTINCT M.UNIT_NUMBER) AS NUM_PATIENT," & vbCrLf & _
                    "ROUND(SUM(D.LOWEST_UM_PRICE * M.REQUESTED_DISPENSE_AMOUNT), 2) AS TOTAL_COST," & vbCrLf & _
                    "ROUND(AVG(D.LOWEST_UM_PRICE * M.REQUESTED_DISPENSE_AMOUNT), 2) AS AVG_COST," & vbCrLf & _
                    "COUNT(*) AS VOLUME, D.NDC" & vbCrLf & _
                "FROM MEDICATION_ORDERS M, MEDICATION_FILLS F, DRUG_COST_DATA D" & vbCrLf & _
                "WHERE " & sFilter & _
                    "M.ORDER_NUMBER = F.ORDER_NUMBER AND" & vbCrLf & _
                    "M.DRUG_ALTERNATE_IDENTIFIER = D.NDC AND" & vbCrLf & _
                    "M.REQUESTED_DISPENSE_AMOUNT IS NOT NULL AND" & vbCrLf & _
                    "D.LOWEST_UM_PRICE IS NOT NULL AND D.GC3_NAME IS NOT NULL AND" & vbCrLf & sDateFilter & vbCrLf & _
                "GROUP BY D.GC3_NAME, D.NDC, D.DRUG_CLASS" & vbCrLf & _
                "ORDER BY COUNT(*) DESC, D.GC3_NAME ASC" & vbCrLf & _
            ")" & vbCrLf & _
            "WHERE ROWNUM < 51" & vbCrLf
            
    Set oRec = goDBDAS.OpenRecordset(sStmt)
    '--- Highlight the drugs that are in the top 50 highest cost drugs by MTF
    Set oHighlight = New MTFTop50RankDrugs
    If oReportItem.Criteria1 = MTF Then Set oRec = oHighlight.HighlightTop50MTFDrugs(oRec, oReportItem.CriteriaSel1.Item(1).NCID)
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
    gobjshared.ShowVBError Err.Number, Err.Description, "RepTop50DrugsByVolTab.GetData", "CHCSII_Reports", vbExclamation
CleanUp:
    Set oHighlight = Nothing
    Exit Function
End Function

