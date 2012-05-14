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
Attribute VB_Name = "RepScreenMTF"
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
    If Not GetData(oReportItem, oRec, oCollFormulas) Then GoTo CleanUp          ' data for provider
    moCollRecSet.Add oRec, "ScreenMTF"

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
    Dim sCriteria As String
    Dim sName As String
    
    Dim sStartDate As String
    Dim sEndDate As String
    Dim strDateInfo As String
    Dim sDateFilter As String
    Dim sFilter As String
    
    Dim sHeader As String
    Dim iResponse As Integer
    
    On Error GoTo ErrHandler
    GetData = False

    '-- Get Filter Information for SQL    '----------------------------------------------------------------------------
    sStartDate = GetOracleStartDate(GetFirstDayOfMonth(oReportItem.DateFrom))
    sEndDate = GetOracleEndDate(GetLastDayOfMonth(oReportItem.DateTo))
    
    sDateFilter = sDateFilter & " (START_DATE >= " & sStartDate & vbCrLf & " AND START_DATE <= " & sEndDate & ") "
    
    sCriteria = GetScreeningServiceItems(oReportItem.CriteriaSel2.Item(1).NCID)
    If sCriteria = "" Then
       MsgBxARMd "The screening service selected has no tests assigned to it. Report cancelled.", vbOKOnly + vbInformation
       Exit Function
    End If
    
    If oReportItem.Criteria1 = MTF Then
        sFilter = " Facility_NCID = " & oReportItem.CriteriaSel1.Item(1).NCID & " And " & vbCrLf
        sFilter = " MEPRS_CODE IN (SELECT DISTINCT MEPRS_CODE FROM CLINIC WHERE FACILITY_NCID = " & oReportItem.CriteriaSel1.Item(1).NCID & ") AND "
    ElseIf oReportItem.Criteria1 = MTFDMIS Then
        'sFilter = " ENTERING_LOCATION_NCID IN (SELECT DISTINCT NCID FROM CLINIC WHERE DMIS_ID_CODE = '" & oReportItem.CriteriaSel1.Item(1).DMIS_ID & "') AND "
        sFilter = " MEPRS_CODE IN (SELECT DISTINCT MEPRS_CODE FROM CLINIC WHERE DMIS_ID_CODE = " & oReportItem.CriteriaSel1.Item(1).DMIS_ID & ") AND "
    End If
    
    sName = oReportItem.CriteriaSel2.Item(1).Name
    
    sStmt = "SELECT UNIVERSAL_SERVICE_NCID, TO_CHAR(COUNT(*)) AS TOTAL," & vbCrLf & _
                "TO_CHAR(TO_DATE(START_DATE), 'YYYY-MM') AS PER_DATE," & vbCrLf & _
                "'" & sName & "' AS SERVICE" & vbCrLf & _
            "FROM LAB_RAD_ORDERS" & vbCrLf & _
            "WHERE " & sFilter & vbCrLf & _
            "UNIVERSAL_SERVICE_NCID IN (" & sCriteria & ") AND " & vbCrLf & sDateFilter & vbCrLf & _
            "GROUP BY UNIVERSAL_SERVICE_NCID, TO_CHAR(TO_DATE(START_DATE), 'YYYY-MM')" & vbCrLf & _
            "ORDER BY TO_CHAR(TO_DATE(START_DATE), 'YYYY-MM') " & vbCrLf


    Set oRec = goDBDAS.OpenRecordset(sStmt)
    Set oRec = CheckData(oRec, CDate(GetFirstDayOfMonth(oReportItem.DateFrom)), CDate(GetFirstDayOfMonth(oReportItem.DateTo)))
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
    gobjshared.ShowVBError Err.Number, Err.Description, "RepScreenMTF.GetData", "CHCSII_Reports", vbExclamation
    'Resume Next
CleanUp:
    Exit Function
End Function

'----------------------------------------------------------------------------------------
'---If a service month has no data, add 0 as the total for the month to the recordset.
'----------------------------------------------------------------------------------------
Private Function CheckData(oRS As ADODB.Recordset, dStartDate As Date, dEndDate As Date) As ADODB.Recordset
    Dim dServiceDate As Date
    Dim sUniversalNCID As String
    Dim sService As String
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim oRsTmp() As SCREEN_SERVICE_RS
    
    If oRS.RecordCount > 0 Then
        oRS.MoveFirst
        dServiceDate = dStartDate
    Else
        Set CheckData = oRS
        GoTo CleanUp
    End If
    
    j = -1
    ReDim oRsTmp(0)
    
    Do While Not oRS.EOF
        '--- if the start date is less than the date in the first row of the recordset,
        '--- add the missing data.
        '--- if the start date is the same as the date in the first row of the recordset,
        '--- check if there are any skipped month.
        If dStartDate = dServiceDate And CDate(oRS("PER_DATE")) > dStartDate Then
            k = 1
        Else
            k = 2
        End If
        
        If DateDiff("m", dServiceDate, CDate(oRS("PER_DATE"))) > k - 1 Then
            sUniversalNCID = oRS("UNIVERSAL_SERVICE_NCID")
            sService = oRS("SERVICE")
            
            For i = k To DateDiff("m", dServiceDate, CDate(oRS("PER_DATE")))
                j = j + 1
                ReDim Preserve oRsTmp(j)
                
                oRsTmp(j).PER_DATE = DateAdd("m", i - 1, dServiceDate)
                oRsTmp(j).SERVICE = sService
                oRsTmp(j).UNIVERSAL_SERVICE_NCID = sUniversalNCID
                oRsTmp(j).TOTAL = "0"
            Next i
        End If
        
        dServiceDate = CDate(oRS("PER_DATE"))
        sUniversalNCID = oRS("UNIVERSAL_SERVICE_NCID")
        sService = oRS("SERVICE")
        oRS.MoveNext
    Loop
    
    '--- check the date gap between the maximum date in the recordset and the end date
    For i = 1 To DateDiff("m", dServiceDate, dEndDate)
        j = j + 1
        ReDim Preserve oRsTmp(j)

        oRsTmp(j).PER_DATE = DateAdd("m", i, dServiceDate)
        oRsTmp(j).SERVICE = sService
        oRsTmp(j).UNIVERSAL_SERVICE_NCID = sUniversalNCID
        oRsTmp(j).TOTAL = "0"
    Next i

    '--- Add the additional data to the recordset
    If Not UBound(oRsTmp) = 0 Then
        For i = 0 To UBound(oRsTmp)
             oRS.AddNew
             oRS("PER_DATE") = oRsTmp(i).PER_DATE
             oRS("SERVICE") = oRsTmp(i).SERVICE
             oRS("UNIVERSAL_SERVICE_NCID") = oRsTmp(i).UNIVERSAL_SERVICE_NCID
             oRS("TOTAL") = oRsTmp(i).TOTAL
             oRS.Update
        Next i
    End If
    
    Set CheckData = oRS
    GoTo CleanUp
    
ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepScreenMTF.CheckData", "CHCSII_Reports", vbExclamation
CleanUp:
    Exit Function
End Function




