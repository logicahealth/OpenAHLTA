VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RepCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private moCollRecSet As Collection
Private moCollFormulas As Collection

Private Function DataGetLab(sReportName As String, _
                            oReportItem As ReportItem, _
                            oReport As CRAXDRT.Report, _
                            oCollFormulas As Collection) As Boolean

    If Not DebugMode Then On Error GoTo ErrHandler
    DataGetLab = False
    
    Dim sStmt As String
    Dim i As Integer
    Dim oRec As ADODB.Recordset
    Dim sDateFilter As String
    Dim sFilter As String
    Dim sFilter2 As String

    '--- SCR 30176
    '--- Rewrote the SQL to get the lab results data from 3M tables.
    
    '-- Get Filter Information for SQL
    '----------------------------------------------------------------------------
    sDateFilter = "(TO_DATE(SUBSTR(tc.text_comment,15,14),'YYYYMMDDHH24MISS') >= " & GetOracleStartDate(oReportItem.DateFrom) & vbNewLine & _
                  "and TO_DATE(SUBSTR(tc.text_comment,15,14),'YYYYMMDDHH24MISS') <= " & GetOracleEndDate(oReportItem.DateTo) & ")"
    
    If Not GetSQLCriteria(oReportItem, 1, sFilter) Then GoTo CleanUp
    
    If oReportItem.Criteria1 = MTF Then
        DataGetLab = DataGetLabMTF(sReportName, oReportItem, oReport, oCollFormulas)
        Exit Function
    ElseIf oReportItem.Criteria1 = PCM Then
        DataGetLab = DataGetLabPCM(sReportName, oReportItem, oReport, oCollFormulas)
        Exit Function
    ElseIf oReportItem.Criteria1 = Provider Then
        DataGetLab = DataGetLabProvider(sReportName, oReportItem, oReport, oCollFormulas)
        Exit Function
    End If
    
    '-- Select
    sStmt = "SELECT " & vbNewLine & _
                "e.event_id," & vbNewLine & _
                "e.unit_number," & vbNewLine & _
                "e.facility_ncid," & vbNewLine & _
                "r.representation as lab_name," & vbNewLine & _
                "TO_DATE(SUBSTR(tc.text_comment,15,14),'YYYYMMDDHH24MISS') as result_date," & vbNewLine & _
                "SUBSTR(TC2.TEXT_COMMENT,INSTR(TC2.TEXT_COMMENT, ' ', 1, 2) + 1) AS PROVIDER_NAME," & vbNewLine & _
                "e.event_subtype_ncid AS Lab_Type_Id" & vbNewLine
    
    '-- From
    sStmt = sStmt & _
            "FROM" & vbNewLine & _
                "event e," & vbNewLine & _
                "lab_test lt," & vbNewLine & _
                "test_comment tc," & vbNewLine & _
                "test_comment tc2," & vbNewLine & _
                "rsform_context rc," & vbNewLine & _
                "rsform r" & vbNewLine
    
    
    '-- Where
    sStmt = sStmt & _
            "Where" & vbNewLine & _
                "rc.preferred_score=0 and" & vbNewLine & _
                "rc.context_ncid=21916 and" & vbNewLine & _
                "rc.rsform_id = r.rsform_id and" & vbNewLine & _
                "r.ncid = e.event_subtype_ncid and" & vbNewLine & _
                "lt.event_id  = e.event_id AND" & vbNewLine & _
                "tc.event_id = e.event_id and" & vbNewLine & _
                "tc.text_comment like 'DATE_RESULTED%' AND" & vbNewLine & _
                "tc2.event_id = e.event_id and" & vbNewLine & _
                "tc2.text_comment like 'PROVIDER%' AND" & vbNewLine & _
                sDateFilter & " and" & vbNewLine & _
                sFilter
    
    
    Set oRec = goDBDAS.OpenRecordset(sStmt)
    DoEvents
    
    '--------------------------------
    '-- Set Data
    '----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection
    
    moCollRecSet.Add oRec, "lab"
    
    '-- Create Report Object
    '-----------------------------------------------------------------------------
    If Not CreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo CleanUp
    
    
    DoEvents
    DataGetLab = True
    GoTo CleanUp

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.DataGetLab", "CHCSII_Reports", vbExclamation)
CleanUp:
    Set oRec = Nothing

End Function
Private Function DataGetLabProvider(sReportName As String, _
                            oReportItem As ReportItem, _
                            oReport As CRAXDRT.Report, _
                            oCollFormulas As Collection) As Boolean

    If Not DebugMode Then On Error GoTo ErrHandler
    DataGetLabProvider = False
    
    Dim sStmt As String
    Dim sMachineName As String
    
    Dim i As Integer
    Dim oRec As ADODB.Recordset
    Dim sDateFilter As String
    Dim sFilter As String
    Dim sFilter2 As String

    sMachineName = GetLocalComputerName
    '-- Get Filter Information for SQL
    '----------------------------------------------------------------------------
    sDateFilter = "(TO_DATE(SUBSTR(tc.text_comment,15,14),'YYYYMMDDHH24MISS') >= " & GetOracleStartDate(oReportItem.DateFrom) & vbNewLine & _
                  "and TO_DATE(SUBSTR(tc.text_comment,15,14),'YYYYMMDDHH24MISS') <= " & GetOracleEndDate(oReportItem.DateTo) & ")"
    
    If Not GetSQLCriteria(oReportItem, 1, sFilter) Then GoTo CleanUp
    sFilter = "SUBSTR(TC2.TEXT_COMMENT,INSTR(TC2.TEXT_COMMENT, ' ', 1, 2) + 1) = '" & Replace(oReportItem.CriteriaSel1.Item(1).Name, ",", ", ") & "'"

    sStmt = "DELETE FROM INT_TEST_RESULTS " & vbCrLf
    sStmt = sStmt & "WHERE COMPUTER_NAME = " & SQLQuote(sMachineName)
    goDBDAS.ExecuteSQL sStmt

    sStmt = "INSERT INTO INT_TEST_RESULTS (EVENT_ID, PROVIDER_NAME, COMPUTER_NAME, PROCESS_STATUS)" & vbCrLf
    sStmt = sStmt & "SELECT tc2.event_id, SUBSTR(TC2.TEXT_COMMENT,INSTR(TC2.TEXT_COMMENT, ' ', 1, 2) + 1) AS PROVIDER_NAME, '" & sMachineName & "', 1" & vbCrLf
    sStmt = sStmt & "FROM test_comment tc2" & vbCrLf
    sStmt = sStmt & "Where tc2.text_comment like 'PROVIDER%' AND" & vbCrLf
    sStmt = sStmt & sFilter
    goDBDAS.ExecuteSQL sStmt
    
    sStmt = "INSERT INTO INT_TEST_RESULTS (RESULT_DATE, EVENT_ID, PROVIDER_NAME, COMPUTER_NAME, PROCESS_STATUS)" & vbCrLf
    sStmt = sStmt & "SELECT TO_DATE(SUBSTR(tc.text_comment,15,14),'YYYYMMDDHH24MISS') as result_date, " & vbCrLf
    sStmt = sStmt & "i.EVENT_ID, i.PROVIDER_NAME, i.COMPUTER_NAME, 2" & vbCrLf
    sStmt = sStmt & "FROM test_comment tc, INT_TEST_RESULTS i" & vbCrLf
    sStmt = sStmt & "Where tc.text_comment like 'DATE_RESULTED%' AND" & vbCrLf
    sStmt = sStmt & sDateFilter
    sStmt = sStmt & "AND tc.event_id =  i.event_id "
    sStmt = sStmt & "AND i.PROCESS_STATUS =  1 "
    sStmt = sStmt & "AND i.COMPUTER_NAME = " & SQLQuote(sMachineName)
    goDBDAS.ExecuteSQL sStmt

    sStmt = "DELETE FROM INT_TEST_RESULTS " & vbCrLf
    sStmt = sStmt & "WHERE PROCESS_STATUS =  1 "
    sStmt = sStmt & "AND COMPUTER_NAME = " & SQLQuote(sMachineName)
    goDBDAS.ExecuteSQL sStmt

    sStmt = "INSERT INTO INT_TEST_RESULTS (UNIT_NUMBER, FACILITY_NCID, TEST_TYPE_ID, RESULT_DATE, EVENT_ID, PROVIDER_NAME, COMPUTER_NAME, PROCESS_STATUS)" & vbCrLf
    sStmt = sStmt & "SELECT e.UNIT_NUMBER, e.FACILITY_NCID, e.EVENT_SUBTYPE_NCID, " & vbCrLf
    sStmt = sStmt & "I.RESULT_DATE, i.EVENT_ID, i.PROVIDER_NAME, i.COMPUTER_NAME, 3" & vbCrLf
    sStmt = sStmt & "FROM event e, INT_TEST_RESULTS i" & vbCrLf
    sStmt = sStmt & "Where i.event_id  = e.event_id " & vbCrLf
    sStmt = sStmt & "AND (e.EVENT_TYPE_NCID = 485 OR e.EVENT_TYPE_NCID = 75412) " & vbCrLf
    sStmt = sStmt & "AND i.PROCESS_STATUS =  2 " & vbCrLf
    sStmt = sStmt & "AND i.COMPUTER_NAME = " & SQLQuote(sMachineName)
    goDBDAS.ExecuteSQL sStmt

    sStmt = "DELETE FROM INT_TEST_RESULTS " & vbCrLf
    sStmt = sStmt & "WHERE PROCESS_STATUS =  2 "
    sStmt = sStmt & "AND COMPUTER_NAME = " & SQLQuote(sMachineName)
    goDBDAS.ExecuteSQL sStmt

    sStmt = "INSERT INTO INT_TEST_RESULTS (TEST_NAME, UNIT_NUMBER, FACILITY_NCID, TEST_TYPE_ID, RESULT_DATE, EVENT_ID, PROVIDER_NAME, COMPUTER_NAME, PROCESS_STATUS)" & vbCrLf
    sStmt = sStmt & "SELECT r.REPRESENTATION, i.UNIT_NUMBER, i.FACILITY_NCID, i.TEST_TYPE_ID, " & vbCrLf
    sStmt = sStmt & "i.RESULT_DATE, i.EVENT_ID, i.PROVIDER_NAME, i.COMPUTER_NAME, 4" & vbCrLf
    sStmt = sStmt & "FROM rsform_context rc, rsform r, INT_TEST_RESULTS i" & vbCrLf
    sStmt = sStmt & "Where rc.PREFERRED_SCORE=0 and" & vbCrLf
    sStmt = sStmt & "rc.context_ncid=21916 and" & vbCrLf
    sStmt = sStmt & "rc.rsform_id = r.rsform_id and" & vbCrLf
    sStmt = sStmt & "r.NCID = i.TEST_TYPE_ID" & vbCrLf
    sStmt = sStmt & "AND i.PROCESS_STATUS =  3 " & vbCrLf
    sStmt = sStmt & "AND i.COMPUTER_NAME = " & SQLQuote(sMachineName)
    goDBDAS.ExecuteSQL sStmt

    sStmt = "DELETE FROM INT_TEST_RESULTS " & vbCrLf
    sStmt = sStmt & "WHERE PROCESS_STATUS =  3 "
    sStmt = sStmt & "AND COMPUTER_NAME = " & SQLQuote(sMachineName)
    goDBDAS.ExecuteSQL sStmt

    sStmt = "SELECT EVENT_ID, UNIT_NUMBER, FACILITY_NCID, TEST_NAME AS LAB_NAME, RESULT_DATE, PROVIDER_NAME, TEST_TYPE_ID AS LAB_TYPE_ID " & vbCrLf
    sStmt = sStmt & "FROM INT_TEST_RESULTS " & vbCrLf
    sStmt = sStmt & "WHERE PROCESS_STATUS =  4 "
    sStmt = sStmt & "AND COMPUTER_NAME = " & SQLQuote(sMachineName)
    Set oRec = goDBDAS.OpenRecordset(sStmt)
    DoEvents

'    sStmt = "DELETE FROM INT_TEST_RESULTS " & vbCrLf
'    sStmt = sStmt & "WHERE COMPUTER_NAME = " & SQLQuote(sMachineName)
'    goDBDAS.ExecuteSQL sStmt

    '--------------------------------
    '-- Set Data
    '----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection

    moCollRecSet.Add oRec, "lab"

    '-- Create Report Object
    '-----------------------------------------------------------------------------
    If Not CreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo CleanUp

    DoEvents
    DataGetLabProvider = True
    GoTo CleanUp
    
ErrHandler:
    If Err = -2147221503 Then Resume Next 'No data found error proceed to next step
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.DataGetLabProvider", "CHCSII_Reports", vbExclamation)
CleanUp:
    Set oRec = Nothing

End Function

Private Function DataGetLabPCM(sReportName As String, _
                            oReportItem As ReportItem, _
                            oReport As CRAXDRT.Report, _
                            oCollFormulas As Collection) As Boolean

    Dim sStmt As String
    Dim sMachineName As String
    
    Dim i As Integer
    Dim oRec As ADODB.Recordset
    Dim sDateFilter As String
    Dim sFilter As String
    Dim sFilter2 As String

    If Not DebugMode Then On Error GoTo ErrHandler
    
    DataGetLabPCM = False
    sMachineName = GetLocalComputerName
    
    '-- Get Filter Information for SQL
    '----------------------------------------------------------------------------
    sDateFilter = "(TO_DATE(SUBSTR(tc.text_comment,15,14),'YYYYMMDDHH24MISS') >= " & GetOracleStartDate(oReportItem.DateFrom) & vbNewLine & _
                  "and TO_DATE(SUBSTR(tc.text_comment,15,14),'YYYYMMDDHH24MISS') <= " & GetOracleEndDate(oReportItem.DateTo) & ")"
    
    If Not GetSQLCriteria(oReportItem, 1, sFilter) Then GoTo CleanUp
    'sFilter = "SUBSTR(TC2.TEXT_COMMENT,INSTR(TC2.TEXT_COMMENT, ' ', 1, 2) + 1) = '" & Replace(oReportItem.CriteriaSel1.Item(1).Name, ",", ", ") & "'"

    sStmt = "DELETE FROM INT_TEST_RESULTS " & vbCrLf
    sStmt = sStmt & "WHERE COMPUTER_NAME = " & SQLQuote(sMachineName)
    goDBDAS.ExecuteSQL sStmt

    sStmt = "INSERT INTO INT_TEST_RESULTS (EVENT_ID, UNIT_NUMBER, FACILITY_NCID, TEST_TYPE_ID, RESULT_DATE, COMPUTER_NAME, PROCESS_STATUS)" & vbCrLf
    sStmt = sStmt & "SELECT e.event_id, e.unit_number, e.facility_ncid, e.event_subtype_ncid AS Lab_Type_Id, "
    sStmt = sStmt & "TO_DATE(SUBSTR(tc.text_comment,15,14),'YYYYMMDDHH24MISS') as result_date, '" & sMachineName & "', 1" & vbCrLf
    sStmt = sStmt & "FROM event e, test_comment tc" & vbCrLf
    sStmt = sStmt & "Where tc.text_comment like 'DATE_RESULTED%' AND" & vbCrLf
    sStmt = sStmt & "tc.event_id = e.event_id and" & vbCrLf
    sStmt = sStmt & sDateFilter & vbCrLf
    sStmt = sStmt & "AND (e.EVENT_TYPE_NCID = 485 OR e.EVENT_TYPE_NCID = 75412) " & vbCrLf
    sStmt = sStmt & "AND e.UNIT_NUMBER IN " & vbCrLf
    sStmt = sStmt & "(SELECT DISTINCT UNIT_NUMBER FROM PATIENT_GENERAL_CODES WHERE PCM_CODE = " & SQLQuote(oReportItem.CriteriaSel1.Item(1).PCMCode) & ")"
    goDBDAS.ExecuteSQL sStmt
    
    sStmt = "INSERT INTO INT_TEST_RESULTS (PROVIDER_NAME, EVENT_ID, UNIT_NUMBER, FACILITY_NCID, TEST_TYPE_ID, RESULT_DATE, COMPUTER_NAME, PROCESS_STATUS)" & vbCrLf
    sStmt = sStmt & "SELECT SUBSTR(TC2.TEXT_COMMENT,INSTR(TC2.TEXT_COMMENT, ' ', 1, 2) + 1) AS PROVIDER_NAME, "
    sStmt = sStmt & "i.EVENT_ID, i.UNIT_NUMBER, i.FACILITY_NCID, i.TEST_TYPE_ID, i.RESULT_DATE, i.COMPUTER_NAME, 2" & vbCrLf
    sStmt = sStmt & "FROM test_comment tc2, INT_TEST_RESULTS i" & vbCrLf
    sStmt = sStmt & "Where tc2.text_comment like 'PROVIDER%' " & vbCrLf
    sStmt = sStmt & "AND tc2.event_id =  i.event_id " & vbCrLf
    sStmt = sStmt & "AND i.PROCESS_STATUS =  1 " & vbCrLf
    sStmt = sStmt & "AND i.COMPUTER_NAME = " & SQLQuote(sMachineName)
    goDBDAS.ExecuteSQL sStmt
    
    sStmt = "DELETE FROM INT_TEST_RESULTS " & vbCrLf
    sStmt = sStmt & "WHERE PROCESS_STATUS =  1 "
    sStmt = sStmt & "AND COMPUTER_NAME = " & SQLQuote(sMachineName)
    goDBDAS.ExecuteSQL sStmt
    
    sStmt = "INSERT INTO INT_TEST_RESULTS (TEST_NAME, UNIT_NUMBER, FACILITY_NCID, TEST_TYPE_ID, RESULT_DATE, EVENT_ID, PROVIDER_NAME, COMPUTER_NAME, PROCESS_STATUS)" & vbCrLf
    sStmt = sStmt & "SELECT r.REPRESENTATION, i.UNIT_NUMBER, i.FACILITY_NCID, i.TEST_TYPE_ID, " & vbCrLf
    sStmt = sStmt & "i.RESULT_DATE, i.EVENT_ID, i.PROVIDER_NAME, i.COMPUTER_NAME, 3" & vbCrLf
    sStmt = sStmt & "FROM rsform_context rc, rsform r, INT_TEST_RESULTS i" & vbCrLf
    sStmt = sStmt & "Where rc.PREFERRED_SCORE=0 and" & vbCrLf
    sStmt = sStmt & "rc.context_ncid=21916 and" & vbCrLf
    sStmt = sStmt & "rc.rsform_id = r.rsform_id and" & vbCrLf
    sStmt = sStmt & "r.NCID = i.TEST_TYPE_ID" & vbCrLf
    sStmt = sStmt & "AND i.PROCESS_STATUS =  2 " & vbCrLf
    sStmt = sStmt & "AND i.COMPUTER_NAME = " & SQLQuote(sMachineName)
    goDBDAS.ExecuteSQL sStmt

    sStmt = "DELETE FROM INT_TEST_RESULTS " & vbCrLf
    sStmt = sStmt & "WHERE PROCESS_STATUS =  2 "
    sStmt = sStmt & "AND COMPUTER_NAME = " & SQLQuote(sMachineName)
    goDBDAS.ExecuteSQL sStmt

    sStmt = "SELECT EVENT_ID, UNIT_NUMBER, FACILITY_NCID, TEST_NAME AS LAB_NAME, RESULT_DATE, PROVIDER_NAME, TEST_TYPE_ID AS LAB_TYPE_ID " & vbCrLf
    sStmt = sStmt & "FROM INT_TEST_RESULTS " & vbCrLf
    sStmt = sStmt & "WHERE PROCESS_STATUS =  3 "
    sStmt = sStmt & "AND COMPUTER_NAME = " & SQLQuote(sMachineName)
    Set oRec = goDBDAS.OpenRecordset(sStmt)
    DoEvents

'    sStmt = "DELETE FROM INT_TEST_RESULTS " & vbCrLf
'    sStmt = sStmt & "WHERE COMPUTER_NAME = " & SQLQuote(sMachineName)
'    goDBDAS.ExecuteSQL sStmt

    '--------------------------------
    '-- Set Data
    '----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection

    moCollRecSet.Add oRec, "lab"

    '-- Create Report Object
    '-----------------------------------------------------------------------------
    If Not CreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo CleanUp

    DoEvents
    DataGetLabPCM = True
    GoTo CleanUp
    
ErrHandler:
    If Err = -2147221503 Then Resume Next 'No data found error proceed to next step
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.DataGetLabPCM", "CHCSII_Reports", vbExclamation)
CleanUp:
    Set oRec = Nothing

End Function

Private Function DataGetLabMTF(sReportName As String, _
                            oReportItem As ReportItem, _
                            oReport As CRAXDRT.Report, _
                            oCollFormulas As Collection) As Boolean

    If Not DebugMode Then On Error GoTo ErrHandler
    DataGetLabMTF = False
    
    Dim sStmt As String
    Dim sMachineName As String
    
    Dim i As Integer
    Dim oRec As ADODB.Recordset
    Dim sDateFilter As String
    Dim sFilter As String
    Dim sFilter2 As String

    sMachineName = GetLocalComputerName
    '-- Get Filter Information for SQL
    '----------------------------------------------------------------------------
    sDateFilter = "(TO_DATE(SUBSTR(tc.text_comment,15,14),'YYYYMMDDHH24MISS') >= " & GetOracleStartDate(oReportItem.DateFrom) & vbNewLine & _
                  "and TO_DATE(SUBSTR(tc.text_comment,15,14),'YYYYMMDDHH24MISS') <= " & GetOracleEndDate(oReportItem.DateTo) & ")"
    
    If Not GetSQLCriteria(oReportItem, 1, sFilter) Then GoTo CleanUp
    sFilter = Replace(sFilter, "a.facility_ncid", "e.facility_ncid")

    sStmt = "DELETE FROM INT_TEST_RESULTS " & vbCrLf
    sStmt = sStmt & "WHERE COMPUTER_NAME = " & SQLQuote(sMachineName)
    goDBDAS.ExecuteSQL sStmt

    sStmt = "INSERT INTO INT_TEST_RESULTS (EVENT_ID, UNIT_NUMBER, FACILITY_NCID, TEST_TYPE_ID, RESULT_DATE, COMPUTER_NAME, PROCESS_STATUS)" & vbCrLf
    sStmt = sStmt & "SELECT e.event_id, e.unit_number, e.facility_ncid, e.event_subtype_ncid AS Lab_Type_Id, "
    sStmt = sStmt & "TO_DATE(SUBSTR(tc.text_comment,15,14),'YYYYMMDDHH24MISS') as result_date, '" & sMachineName & "', 1" & vbCrLf
    sStmt = sStmt & "FROM event e, test_comment tc" & vbCrLf
    sStmt = sStmt & "Where tc.text_comment like 'DATE_RESULTED%' AND" & vbCrLf
    sStmt = sStmt & "tc.event_id = e.event_id and" & vbCrLf
    sStmt = sStmt & sDateFilter & vbCrLf
    sStmt = sStmt & "AND (e.EVENT_TYPE_NCID = 485 OR e.EVENT_TYPE_NCID = 75412) " & vbCrLf
    sStmt = sStmt & "AND " & sFilter
    goDBDAS.ExecuteSQL sStmt
    
    sStmt = "INSERT INTO INT_TEST_RESULTS (PROVIDER_NAME, EVENT_ID, UNIT_NUMBER, FACILITY_NCID, TEST_TYPE_ID, RESULT_DATE, COMPUTER_NAME, PROCESS_STATUS)" & vbCrLf
    sStmt = sStmt & "SELECT SUBSTR(TC2.TEXT_COMMENT,INSTR(TC2.TEXT_COMMENT, ' ', 1, 2) + 1) AS PROVIDER_NAME, "
    sStmt = sStmt & "i.EVENT_ID, i.UNIT_NUMBER, i.FACILITY_NCID, i.TEST_TYPE_ID, i.RESULT_DATE, i.COMPUTER_NAME, 2" & vbCrLf
    sStmt = sStmt & "FROM test_comment tc2, INT_TEST_RESULTS i" & vbCrLf
    sStmt = sStmt & "Where tc2.text_comment like 'PROVIDER%' " & vbCrLf
    sStmt = sStmt & "AND tc2.event_id =  i.event_id " & vbCrLf
    sStmt = sStmt & "AND i.PROCESS_STATUS =  1 " & vbCrLf
    sStmt = sStmt & "AND i.COMPUTER_NAME = " & SQLQuote(sMachineName)
    goDBDAS.ExecuteSQL sStmt
    
    sStmt = "DELETE FROM INT_TEST_RESULTS " & vbCrLf
    sStmt = sStmt & "WHERE PROCESS_STATUS =  1 "
    sStmt = sStmt & "AND COMPUTER_NAME = " & SQLQuote(sMachineName)
    goDBDAS.ExecuteSQL sStmt
    
    sStmt = "INSERT INTO INT_TEST_RESULTS (TEST_NAME, UNIT_NUMBER, FACILITY_NCID, TEST_TYPE_ID, RESULT_DATE, EVENT_ID, PROVIDER_NAME, COMPUTER_NAME, PROCESS_STATUS)" & vbCrLf
    sStmt = sStmt & "SELECT r.REPRESENTATION, i.UNIT_NUMBER, i.FACILITY_NCID, i.TEST_TYPE_ID, " & vbCrLf
    sStmt = sStmt & "i.RESULT_DATE, i.EVENT_ID, i.PROVIDER_NAME, i.COMPUTER_NAME, 3" & vbCrLf
    sStmt = sStmt & "FROM rsform_context rc, rsform r, INT_TEST_RESULTS i" & vbCrLf
    sStmt = sStmt & "Where rc.PREFERRED_SCORE=0 and" & vbCrLf
    sStmt = sStmt & "rc.context_ncid=21916 and" & vbCrLf
    sStmt = sStmt & "rc.rsform_id = r.rsform_id and" & vbCrLf
    sStmt = sStmt & "r.NCID = i.TEST_TYPE_ID" & vbCrLf
    sStmt = sStmt & "AND i.PROCESS_STATUS =  2 " & vbCrLf
    sStmt = sStmt & "AND i.COMPUTER_NAME = " & SQLQuote(sMachineName)
    goDBDAS.ExecuteSQL sStmt

    sStmt = "DELETE FROM INT_TEST_RESULTS " & vbCrLf
    sStmt = sStmt & "WHERE PROCESS_STATUS =  2 "
    sStmt = sStmt & "AND COMPUTER_NAME = " & SQLQuote(sMachineName)
    goDBDAS.ExecuteSQL sStmt

    sStmt = "SELECT EVENT_ID, UNIT_NUMBER, FACILITY_NCID, TEST_NAME AS LAB_NAME, RESULT_DATE, PROVIDER_NAME, TEST_TYPE_ID AS LAB_TYPE_ID " & vbCrLf
    sStmt = sStmt & "FROM INT_TEST_RESULTS " & vbCrLf
    sStmt = sStmt & "WHERE PROCESS_STATUS =  3 "
    sStmt = sStmt & "AND COMPUTER_NAME = " & SQLQuote(sMachineName)
    Set oRec = goDBDAS.OpenRecordset(sStmt)
    DoEvents

'    sStmt = "DELETE FROM INT_TEST_RESULTS " & vbCrLf
'    sStmt = sStmt & "WHERE COMPUTER_NAME = " & SQLQuote(sMachineName)
'    goDBDAS.ExecuteSQL sStmt

    '--------------------------------
    '-- Set Data
    '----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection
    
    moCollRecSet.Add oRec, "lab"
    
    '-- Create Report Object
    '-----------------------------------------------------------------------------
    If Not CreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo CleanUp

    DoEvents
    DataGetLabMTF = True
    GoTo CleanUp

ErrHandler:
    If Err = -2147221503 Then Resume Next 'No data found error proceed to next step
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.DataGetLabMTF", "CHCSII_Reports", vbExclamation)
CleanUp:
    Set oRec = Nothing

End Function


Private Function DataGetDiagnosis(sReportName As String, _
                                  oReportItem As ReportItem, _
                                  oReport As CRAXDRT.Report, _
                                  oCollFormulas As Collection) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler
DataGetDiagnosis = False

Dim sStmt As String
Dim i As Integer
Dim oRec As ADODB.Recordset
Dim sDateFilter As String
Dim sFilter As String
Dim sRNDSFilter As String

'-- Get Filter Information for SQL
'----------------------------------------------------------------------------
If Not GetSQLDateRange(oReportItem, sDateFilter) Then GoTo CleanUp
If Not GetSQLCriteria(oReportItem, 1, sFilter) Then GoTo CleanUp

If oReportItem.SecondaryFuncSelected = False Then   'Don't include RNDS
   sRNDSFilter = " AND ( a.APPT_TYPE <> " & SQLQuote(RNDSALIAS1) & " ) AND (a.APPT_TYPE <> " & SQLQuote(RNDSALIAS2) & ")"
Else
   sRNDSFilter = NULLStr
End If

'-- Select
sStmt = "SELECT /*+ USE_NL(P) USE_NL(PC) INDEX (C CLINIC_PK) (E EVENT_X1)  (A APPOINTMENT_X3) */ OV.VALUE_NCID DIAGNOSISNCID," & vbNewLine & _
            "DG.DESCRIPTION DIAGNOSIS," & vbNewLine & _
            "DG.ICD9," & vbNewLine & _
            "P.REPRESENTATION PROVIDERNAME," & vbNewLine & _
            "C.NAME AS CLINICNAME," & vbNewLine & _
            "C.MEPRS_CODE, C.DMIS_ID_CODE" & vbNewLine

'-- From
sStmt = sStmt & _
            "From" & vbNewLine & _
                "APPOINTMENT A," & vbNewLine & _
                "EVENT E, " & vbNewLine & _
                "CLINICAL_OBSERVATION CO," & vbNewLine & _
                "OBSERVATION_VALUE OV," & vbNewLine & _
                "RPT_DIAGNOSIS DG," & vbNewLine & _
                "CLINIC C," & vbNewLine & _
                "RSFORM_CONTEXT PC, RSFORM P" & vbNewLine
'-- Where
sStmt = sStmt & _
        "Where " & vbNewLine & _
            "DG.NCID = OV.VALUE_NCID AND" & vbNewLine & _
            "OV.OBSERVATION_SEQ = CO.OBSERVATION_SEQ AND" & vbNewLine & _
            "OV.EVENT_ID = CO.EVENT_ID AND" & vbNewLine & _
            "CO.EVENT_ID = E.EVENT_ID AND" & vbNewLine & _
            "E.EVENT_SUBTYPE_NCID = 658 AND " & vbNewLine & _
            "E.EVENT_TYPE_NCID = 90803 AND" & vbNewLine & _
            "E.UNIT_NUMBER=A.PATIENT_UNIT_NUMBER AND" & vbNewLine & _
            "E.ENCOUNTER_NUMBER=A.ENCOUNTER_NUMBER AND" & vbNewLine & _
            "C.NCID = A.CLINIC_NCID AND" & vbNewLine & _
            "PC.PREFERRED_SCORE=0 AND" & vbNewLine & _
            "PC.CONTEXT_NCID=2000 AND" & vbNewLine & _
            "PC.RSFORM_ID = P.RSFORM_ID AND" & vbNewLine & _
            "A.CLINICIAN_NCID = P.NCID AND " & vbNewLine & _
            sDateFilter & " And " & vbNewLine & _
            sFilter & sRNDSFilter & vbNewLine

Set oRec = goDBDAS.OpenRecordset(sStmt)
DoEvents

'--------------------------------
'-- Set Data
'----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection

    moCollRecSet.Add oRec, "diagnosis"

'-- Create Report Object
'-----------------------------------------------------------------------------
    If Not CreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo CleanUp

    DoEvents

DataGetDiagnosis = True
GoTo CleanUp
ErrHandler:
    If Err = -2147220874 Then
        MsgBxARMd "The selected report has timed out, please re-run this report again.", vbOKOnly + vbInformation
        Err.Clear
    Else
        Call gobjshared.ShowVBError(Err.Number, Err.Description, "RepCustom.DataGetDiagnosis", "CHCSII_Reports", vbExclamation)
    End If
CleanUp:
    Set oRec = Nothing
    Exit Function
End Function
Private Function DataGetPrimaryDiagnosis(sReportName As String, _
                                         oReportItem As ReportItem, _
                                         oReport As CRAXDRT.Report, _
                                         oCollFormulas As Collection) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler
DataGetPrimaryDiagnosis = False

Dim sStmt As String
Dim i As Integer
Dim oRec As ADODB.Recordset
Dim sDateFilter As String
Dim sFilter As String
Dim sFilter2 As String
Dim sRNDSFilter As String

'-- Get Filter Information for SQL
'----------------------------------------------------------------------------
If Not GetSQLDateRange(oReportItem, sDateFilter) Then GoTo CleanUp
If Not GetSQLCriteria(oReportItem, 1, sFilter) Then GoTo CleanUp

If oReportItem.SecondaryFuncSelected = False Then   'Don't include RNDS
   sRNDSFilter = " AND ( a.APPT_TYPE <> " & SQLQuote(RNDSALIAS1) & " ) AND (a.APPT_TYPE <> " & SQLQuote(RNDSALIAS2) & ")"
Else
   sRNDSFilter = NULLStr
End If

'-- Select
'SCR-43766 10/13/03 RBELL Removed the RULE optimizer hint from the query per Ed Clearys instruction
'sStmt = "SELECT /*+ RULE */ OV.VALUE_NCID DIAGNOSISNCID," & vbNewLine &
sStmt = "SELECT OV.VALUE_NCID DIAGNOSISNCID," & vbNewLine & _
            "DG.DESCRIPTION DIAGNOSIS," & vbNewLine & _
            "DG.ICD9," & vbNewLine & _
            "P.REPRESENTATION PROVIDERNAME," & vbNewLine & _
            "C.NAME AS CLINICNAME," & vbNewLine & _
            "C.MEPRS_CODE, C.DMIS_ID_CODE" & vbNewLine

'-- From
sStmt = sStmt & _
            "From" & vbNewLine & _
                "APPOINTMENT A," & vbNewLine & _
                "EVENT E, " & vbNewLine & _
                "CLINICAL_OBSERVATION CO," & vbNewLine & _
                "OBSERVATION_VALUE OV," & vbNewLine & _
                "RPT_DIAGNOSIS DG," & vbNewLine & _
                "CLINIC C," & vbNewLine & _
                "RSFORM_CONTEXT PC, RSFORM P" & vbNewLine
'-- Where
sStmt = sStmt & _
        "Where " & vbNewLine & _
            "DG.NCID = OV.VALUE_NCID AND" & vbNewLine & _
            "OV.OBSERVATION_SEQ = CO.OBSERVATION_SEQ AND" & vbNewLine & _
            "OV.EVENT_ID = CO.EVENT_ID AND" & vbNewLine & _
            "CO.EVENT_ID = E.EVENT_ID AND" & vbNewLine & _
            "E.EVENT_SUBTYPE_NCID = 658 AND " & vbNewLine & _
            "E.EVENT_TYPE_NCID = 90803 AND" & vbNewLine & _
            "E.ENCOUNTER_NUMBER=A.ENCOUNTER_NUMBER AND" & vbNewLine & _
            "A.ENCOUNTER_STATUS IN (4,5,6) AND" & vbNewLine & _
            "C.NCID = A.CLINIC_NCID AND" & vbNewLine & _
            "PC.PREFERRED_SCORE=0 AND" & vbNewLine & _
            "PC.CONTEXT_NCID=2000 AND" & vbNewLine & _
            "PC.RSFORM_ID = P.RSFORM_ID AND" & vbNewLine & _
            "A.CLINICIAN_NCID = P.NCID AND " & vbNewLine & _
            sDateFilter & " And " & vbNewLine & _
            sFilter & vbNewLine & _
            sRNDSFilter & vbNewLine             'Add RNDS filter

If oReportItem.Criteria1 = Provider Then
    '-- Select
    sStmt = "SELECT /*+ ORDERED */ OV.VALUE_NCID DIAGNOSISNCID," & vbNewLine & _
                "DG.DESCRIPTION DIAGNOSIS," & vbNewLine & _
                "DG.ICD9," & vbNewLine & _
                "'" & oReportItem.CriteriaSel1(1).Name & "' AS PROVIDERNAME," & vbNewLine & _
                "C.NAME AS CLINICNAME," & vbNewLine & _
                "C.MEPRS_CODE, C.DMIS_ID_CODE" & vbNewLine
    
    '-- From
    sStmt = sStmt & _
                "From" & vbNewLine & _
                    "APPOINTMENT A," & vbNewLine & _
                    "EVENT E, " & vbNewLine & _
                    "CLINICAL_OBSERVATION CO," & vbNewLine & _
                    "OBSERVATION_VALUE OV," & vbNewLine & _
                    "RPT_DIAGNOSIS DG," & vbNewLine & _
                    "CLINIC C" & vbNewLine
    '-- Where
    sStmt = sStmt & _
            "Where " & vbNewLine & _
                "DG.NCID = OV.VALUE_NCID AND" & vbNewLine & _
                "OV.OBSERVATION_SEQ = CO.OBSERVATION_SEQ AND" & vbNewLine & _
                "OV.EVENT_ID = CO.EVENT_ID AND" & vbNewLine & _
                "CO.OBSERVATION_NCID=658 AND" & vbNewLine & _
                "CO.EVENT_ID = E.EVENT_ID AND" & vbNewLine & _
                "E.EVENT_SUBTYPE_NCID = 658 AND " & vbNewLine & _
                "E.EVENT_TYPE_NCID = 90803 AND" & vbNewLine & _
                "E.ENCOUNTER_NUMBER=A.ENCOUNTER_NUMBER AND" & vbNewLine & _
                "C.NCID = A.CLINIC_NCID AND" & vbNewLine & _
                "A.ENCOUNTER_STATUS IN (4,5,6) AND" & vbNewLine & _
                sDateFilter & " And " & vbNewLine & _
                sFilter & vbNewLine & _
                sRNDSFilter & vbNewLine

End If

If oReportItem.Criteria2 = Med Then

    sStmt = sStmt & _
        "AND E.EVENT_ID IN" & vbNewLine & _
            "(SELECT" & vbNewLine & _
                "E.EVENT_ID" & vbNewLine & _
            "FROM" & vbNewLine & _
                "EVENT E," & vbNewLine & _
                "SEMANTIC_LINK_ELEMENT SLE," & vbNewLine & _
                "SEMANTIC_LINK SL," & vbNewLine & _
                "OE_RX OER" & vbNewLine & _
            "WHERE" & vbNewLine & _
                "e.encounter_number is not null and" & vbNewLine & _
                "e.encounter_number <> 0 and" & vbNewLine & _
                "e.event_id=sl.event_id and" & vbNewLine & _
                "sl.relationship_ncid=205071 and" & vbNewLine & _
                "sl.event_id=sle.event_id and" & vbNewLine & _
                "sl.semantic_link_ord=sle.semantic_link_ord and" & vbNewLine & _
                "sle.element_type_ncid=108702 and" & vbNewLine & _
                "sle.link_event_id=oer.OE_RX_DATA_ID and" & vbNewLine & _
                "oer.order_ien =" & oReportItem.CriteriaSel2.Item(1).IEN & ")"
                
ElseIf oReportItem.Criteria2 = Lab Then
    
    sStmt = sStmt & _
        "AND E.EVENT_ID IN" & vbNewLine & _
            "(SELECT" & vbNewLine & _
                "E.EVENT_ID" & vbNewLine & _
            "FROM" & vbNewLine & _
                "EVENT E," & vbNewLine & _
                "SEMANTIC_LINK_ELEMENT SLE," & vbNewLine & _
                "SEMANTIC_LINK SL," & vbNewLine & _
                "OE_LAB OEL" & vbNewLine & _
            "WHERE" & vbNewLine & _
                "e.encounter_number is not null and" & vbNewLine & _
                "e.encounter_number <> 0 and" & vbNewLine & _
                "e.event_id=sl.event_id and" & vbNewLine & _
                "sl.relationship_ncid=205652 and" & vbNewLine & _
                "sl.event_id=sle.event_id and" & vbNewLine & _
                "sl.semantic_link_ord=sle.semantic_link_ord and" & vbNewLine & _
                "sle.element_type_ncid=108702 and" & vbNewLine & _
                "sle.link_event_id=oel.OE_LAB_DATA_ID and" & vbNewLine & _
                "oel.order_ien =" & oReportItem.CriteriaSel2.Item(1).IEN & ")"
        
ElseIf oReportItem.Criteria2 = Rad Then

    sStmt = sStmt & _
        "AND E.EVENT_ID IN" & vbNewLine & _
            "(SELECT" & vbNewLine & _
                "E.EVENT_ID" & vbNewLine & _
            "FROM" & vbNewLine & _
                "EVENT E," & vbNewLine & _
                "SEMANTIC_LINK_ELEMENT SLE," & vbNewLine & _
                "SEMANTIC_LINK SL," & vbNewLine & _
                "OE_RAD OER" & vbNewLine & _
            "WHERE" & vbNewLine & _
                "e.encounter_number is not null and" & vbNewLine & _
                "e.encounter_number <> 0 and" & vbNewLine & _
                "e.event_id=sl.event_id and" & vbNewLine & _
                "sl.relationship_ncid=205653 and" & vbNewLine & _
                "sl.event_id=sle.event_id and" & vbNewLine & _
                "sl.semantic_link_ord=sle.semantic_link_ord and" & vbNewLine & _
                "sle.element_type_ncid=108702 and" & vbNewLine & _
                "sle.link_event_id=oer.OE_RAD_DATA_ID and" & vbNewLine & _
                "oer.order_ien =" & oReportItem.CriteriaSel2.Item(1).IEN & ")"

ElseIf oReportItem.Criteria2 = Procedure Then

    sStmt = sStmt & _
        "AND E.EVENT_ID IN" & vbNewLine & _
            "(SELECT" & vbNewLine & _
                "E.EVENT_ID" & vbNewLine & _
            "FROM" & vbNewLine & _
                "EVENT E," & vbNewLine & _
                "SEMANTIC_LINK_ELEMENT SLE," & vbNewLine & _
                "SEMANTIC_LINK SL," & vbNewLine & _
                "EVENT E2," & vbNewLine & _
                "clinical_observation co," & vbNewLine & _
                "observation_value ov" & vbNewLine

    sStmt = sStmt & _
            "WHERE" & vbNewLine & _
                "e.encounter_number is not null and" & vbNewLine & _
                "e.encounter_number <> 0 and" & vbNewLine & _
                "e.event_id=sl.event_id and" & vbNewLine & _
                "sl.relationship_ncid=205069 and" & vbNewLine & _
                "sl.event_id=sle.event_id and" & vbNewLine & _
                "sl.semantic_link_ord=sle.semantic_link_ord and" & vbNewLine & _
                "sle.element_type_ncid=108702 and" & vbNewLine & _
                "sle.link_event_id=e2.event_id and" & vbNewLine & _
                "e2.encounter_number is not null and" & vbNewLine & _
                "e2.encounter_number <> 0 and" & vbNewLine & _
                "e2.event_id=co.event_id and" & vbNewLine & _
                "co.observation_ncid=41 and" & vbNewLine & _
                "co.observation_seq = ov.observation_seq and" & vbNewLine & _
                "co.event_id = ov.event_id and" & vbNewLine & _
                "ov.value_ncid =" & oReportItem.CriteriaSel2.Item(1).NCID & ")"
End If

Set oRec = goDBDAS.OpenRecordset(sStmt)
DoEvents

'--------------------------------
'-- Set Data
'----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection

    moCollRecSet.Add oRec, "diagnosis"

'-- Create Report Object
'-----------------------------------------------------------------------------
    If Not CreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo CleanUp

    DoEvents

DataGetPrimaryDiagnosis = True
GoTo CleanUp
ErrHandler:
    If Err = -2147220874 Then
        MsgBxARMd "The selected report has timed out, please re-run this report again.", vbOKOnly + vbInformation
        Err.Clear
    Else
        Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.DataGetPrimaryDiagnosis", "CHCSII_Reports", vbExclamation)
    End If
CleanUp:
    Set oRec = Nothing
    Exit Function
End Function

Private Function DataGetProcedure(sReportName As String, _
                                  oReportItem As ReportItem, _
                                  oReport As CRAXDRT.Report, _
                                  oCollFormulas As Collection) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler
DataGetProcedure = False

Dim sStmt As String
Dim i As Integer
Dim oRec As ADODB.Recordset
Dim sDateFilter As String
Dim sFilter As String
Dim sSubQueryA As String
Dim sSubQueryB As String
Dim sSubQueryC As String
Dim sSubQueryD As String
Dim sRNDSFilter As String

'-- Get Filter Information for SQL
'----------------------------------------------------------------------------
If Not GetSQLDateRange(oReportItem, sDateFilter) Then GoTo CleanUp
If Not GetSQLCriteria(oReportItem, 1, sFilter) Then GoTo CleanUp

If oReportItem.SecondaryFuncSelected = False Then   'Don't include RNDS
   sRNDSFilter = " And ( a.APPT_TYPE <> " & SQLQuote(RNDSALIAS1) & " ) AND (a.APPT_TYPE <> " & SQLQuote(RNDSALIAS2) & ")"
Else
   sRNDSFilter = NULLStr
End If

'-- Select
'----------------------------------------------------------------------------
sStmt = "SELECT" & vbNewLine

        '-- ORACLE HINT
        '--------------------------------------------------------------------
sStmt = sStmt & _
        "/*+ ORDERED "

        If oReportItem.Criteria1 = MTF Then
            sStmt = sStmt & " INDEX (A APPOINTMENT_X3) "
        ElseIf oReportItem.Criteria1 = Clinic Then
            sStmt = sStmt & " INDEX (A APPOINTMENT_X1) "
        ElseIf oReportItem.Criteria1 = Provider Then
            sStmt = sStmt & " INDEX (A APPOINTMENT_X2) "
        ElseIf oReportItem.Criteria1 = PCM Then
            sStmt = sStmt & " INDEX (A APPOINTMENT_X4) "
        End If
        
        If oReportItem.Criteria2 = Diagnosis Then
            sStmt = sStmt & " INDEX (E EVENT_PK) "
        Else
            sStmt = sStmt & " INDEX (E EVENT_X3) "
        End If
        
sStmt = sStmt & _
        "USE_NL(OV) USE_NL(CO) USE_NL(P_R) USE_NL(P_RC) USE_NL(PC_R) USE_NL(PC_RC) " & vbNewLine & _
        "USE_NL(ED) USE_NL(COD) USE_NL(OVD) INDEX (ED EVENT_X3) INDEX (COD CLINICAL_OBSERVATION_X2) INDEX (CO CLINICAL_OBSERVATION_X2)*/" & vbNewLine

        '-- Select Clause
        '--------------------------------------------------------------------
sStmt = sStmt & _
        "OV.VALUE_NCID PROCEDURENCID," & vbNewLine & _
        "P_R.REPRESENTATION PROCEDURE," & vbNewLine & _
        "PC_R.REPRESENTATION CPT," & vbNewLine & _
        "P.REPRESENTATION PROVIDERNAME," & vbNewLine & _
        "C.NAME AS CLINICNAME, C.MEPRS_CODE, C.DMIS_ID_CODE" & vbNewLine


'-- From
'----------------------------------------------------------------------------
sStmt = sStmt & _
        "From" & vbNewLine & _
        "APPOINTMENT A," & vbNewLine
        
        If oReportItem.Criteria2 = Diagnosis Then 'If Diagnosis is limiting criteria
            sStmt = sStmt & _
                    "/* Begin IF BASED ON DIAGNOSIS */" & vbNewLine & _
                    "EVENT ED," & vbNewLine & _
                    "CLINICAL_OBSERVATION COD," & vbNewLine & _
                    "OBSERVATION_VALUE OVD," & vbNewLine & _
                    "SEMANTIC_LINK SL," & vbNewLine & _
                    "SEMANTIC_LINK_ELEMENT SLE," & vbNewLine & _
                    "/* End IF BASED ON DIAGNOSIS */" & vbNewLine
        End If
        
sStmt = sStmt & _
        "EVENT E," & vbNewLine & _
        "CLINICAL_OBSERVATION CO," & vbNewLine & _
        "OBSERVATION_VALUE OV," & vbNewLine & _
        "RSFORM P_R, RSFORM_CONTEXT P_RC," & vbNewLine & _
        "RSFORM PC_R, RSFORM_CONTEXT PC_RC," & vbNewLine & _
        "RSFORM P, RSFORM_CONTEXT PC," & vbNewLine & _
        "Clinic C" & vbNewLine

'-- Where
'----------------------------------------------------------------------------
sStmt = sStmt & _
    "Where" & vbNewLine & _
    "C.NCID = A.CLINIC_NCID AND" & vbNewLine & _
    "PC.PREFERRED_SCORE=0 AND" & vbNewLine & _
    "PC.CONTEXT_NCID=2000 AND" & vbNewLine & _
    "PC.RSFORM_ID = P.RSFORM_ID AND" & vbNewLine & _
    "P.NCID = A.CLINICIAN_NCID AND" & vbNewLine & _
    "PC_RC.PREFERRED_SCORE=0 AND" & vbNewLine & _
    "PC_RC.CONTEXT_NCID = 17274 AND" & vbNewLine & _
    "PC_RC.RSFORM_ID = PC_R.RSFORM_ID AND" & vbNewLine & _
    "PC_R.NCID = OV.VALUE_NCID AND" & vbNewLine & _
    "P_RC.PREFERRED_SCORE=0 AND" & vbNewLine & _
    "P_RC.CONTEXT_NCID=2000 AND" & vbNewLine & _
    "P_RC.RSFORM_ID = P_R.RSFORM_ID AND" & vbNewLine & _
    "P_R.NCID = OV.VALUE_NCID AND" & vbNewLine & _
    "OV.OBSERVATION_SEQ = CO.OBSERVATION_SEQ AND" & vbNewLine & _
    "OV.EVENT_ID = CO.EVENT_ID AND" & vbNewLine & _
    "CO.OBSERVATION_NCID=41 AND" & vbNewLine & _
    "CO.EVENT_ID = E.EVENT_ID " & vbNewLine & _
    sRNDSFilter & " AND" & vbNewLine
    
    If oReportItem.Criteria2 = Diagnosis Then 'If Diagnosis is limiting criteria
        sStmt = sStmt & _
                "/* Begin IF BASED ON DIAGNOSIS */" & vbNewLine & _
                "E.EVENT_ID=SLE.LINK_EVENT_ID AND" & vbNewLine & _
                "SLE.ELEMENT_TYPE_NCID=108702 AND" & vbNewLine & _
                "SLE.SEMANTIC_LINK_ORD=SL.SEMANTIC_LINK_ORD AND" & vbNewLine & _
                "SLE.EVENT_ID=SL.EVENT_ID AND" & vbNewLine & _
                "SL.RELATIONSHIP_NCID=205069 AND" & vbNewLine & _
                "SL.EVENT_ID=OVD.EVENT_ID AND" & vbNewLine & _
                "1009771 = OVD.VALUE_NCID AND" & vbNewLine & _
                "OVD.OBSERVATION_SEQ = COD.OBSERVATION_SEQ AND" & vbNewLine & _
                "OVD.EVENT_ID = COD.EVENT_ID AND" & vbNewLine & _
                "COD.OBSERVATION_NCID=658 AND" & vbNewLine & _
                "COD.EVENT_ID = ED.EVENT_ID AND" & vbNewLine & _
                "ED.EVENT_SUBTYPE_NCID=658 AND" & vbNewLine & _
                "ED.EVENT_TYPE_NCID=90803 AND" & vbNewLine & _
                "ED.ENCOUNTER_NUMBER=A.ENCOUNTER_NUMBER AND" & vbNewLine & _
                "/* End IF BASED ON DIAGNOSIS */" & vbNewLine
    Else
        sStmt = sStmt & _
                "E.EVENT_SUBTYPE_NCID=41 AND" & vbNewLine & _
                "E.EVENT_TYPE_NCID=90803 AND" & vbNewLine & _
                "E.ENCOUNTER_NUMBER=A.ENCOUNTER_NUMBER AND" & vbNewLine
    End If

    sStmt = sStmt & _
            "A.ENCOUNTER_NUMBER <> '0' AND" & vbNewLine & _
            "A.ENCOUNTER_STATUS IN (4,5,6) AND" & vbNewLine & _
            sDateFilter & " And" & vbNewLine & _
            sFilter & vbNewLine


Set oRec = goDBDAS.OpenRecordset(sStmt)
DoEvents

'--------------------------------
'-- Set Data
'----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection

    moCollRecSet.Add oRec, "procedure"

'-- Create Report Object
'-----------------------------------------------------------------------------
    If Not CreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo CleanUp

    DoEvents

DataGetProcedure = True
GoTo CleanUp
ErrHandler:
    If Err = -2147220874 Then
        MsgBxARMd "The selected report has timed out, please re-run this report again.", vbOKOnly + vbInformation
        Err.Clear
    Else
        Call gobjshared.ShowVBError(Err.Number, Err.Description, "RepCustom.DataGetProcedure", "CHCSII_Reports", vbExclamation)
    End If
    'resume next
CleanUp:
    Set oRec = Nothing
    Exit Function
End Function

Private Function DataGetEandM(sReportName As String, _
                              oReportItem As ReportItem, _
                              oReport As CRAXDRT.Report, _
                              oCollFormulas As Collection) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler
DataGetEandM = False

Dim sStmt                    As String
Dim i                        As Integer
Dim oRec                     As ADODB.Recordset
Dim sDateFilter              As String
Dim sFilter                  As String
Dim sSubQueryA               As String
Dim sSubQueryB               As String
Dim sSubQueryC               As String
Dim sSubQueryD               As String


'-- Get Filter Information for SQL
'----------------------------------------------------------------------------
If Not GetSQLDateRange(oReportItem, sDateFilter) Then GoTo CleanUp
If Not GetSQLCriteria(oReportItem, 1, sFilter) Then GoTo CleanUp

'--- Sub Query A: get PROCEDURENCID
sSubQueryA = "SELECT OV.VALUE_NCID PROCEDURENCID, A.CLINICIAN_NCID, A.CLINIC_NCID, A.ENCOUNTER_NUMBER" & vbNewLine & _
            "FROM OBSERVATION_VALUE OV, CLINICAL_OBSERVATION CO, EVENT E, APPOINTMENT A" & vbNewLine & _
            "WHERE OV.OBSERVATION_SEQ = CO.OBSERVATION_SEQ AND" & vbNewLine & _
                "OV.EVENT_ID = CO.EVENT_ID AND" & vbNewLine & _
                "CO.OBSERVATION_NCID=41 AND" & vbNewLine & _
                "CO.EVENT_ID = E.EVENT_ID AND" & vbNewLine & _
                "E.EVENT_SUBTYPE_NCID=184917 AND" & vbNewLine & _
                "E.EVENT_TYPE_NCID = 209614 AND" & vbNewLine & _
                "E.ENCOUNTER_NUMBER=A.ENCOUNTER_NUMBER AND" & vbNewLine & _
                "A.ENCOUNTER_NUMBER IS NOT NULL AND" & vbNewLine & _
                "A.ENCOUNTER_NUMBER <> 0 AND" & vbNewLine & _
                "A.ENCOUNTER_STATUS IN (4,5,6) AND" & vbNewLine & _
                sDateFilter & " and" & vbNewLine & _
                sFilter & vbNewLine

'--- Sub Query B: get EMCODE
sSubQueryB = "SELECT RCODE.NCID, TRIM(RCODE.REPRESENTATION) EMCODE" & vbNewLine & _
             "FROM RSFORM_CONTEXT RCCODE, RSFORM RCODE" & vbNewLine & _
             "WHERE RCCODE.CONTEXT_NCID = 17274 AND" & vbNewLine & _
             "RCCODE.RSFORM_ID = RCODE.RSFORM_ID " & vbNewLine

'--- Sub Query C: get DESCRIPTION
sSubQueryC = "SELECT RDESC.NCID, TRIM(RDESC.REPRESENTATION) DESCRIPTION" & vbNewLine & _
             "FROM RSFORM_CONTEXT RCDESC, RSFORM RDESC" & vbNewLine & _
             "WHERE RCDESC.PREFERRED_SCORE=0 AND" & vbNewLine & _
                "RCDESC.CONTEXT_NCID = 17276 AND" & vbNewLine & _
                "RCDESC.RSFORM_ID = RDESC.RSFORM_ID" & vbNewLine

'--- Sub Query D: get PROVIDERNAME
sSubQueryD = "SELECT P.NCID, P.REPRESENTATION AS PROVIDERNAME" & vbNewLine & _
             "FROM RSFORM_CONTEXT PC, RSFORM P" & vbNewLine & _
             "WHERE PC.PREFERRED_SCORE=0 AND" & vbNewLine & _
                "PC.CONTEXT_NCID=2000 AND" & vbNewLine & _
                "PC.RSFORM_ID = P.RSFORM_ID" & vbNewLine
                
sStmt = "SELECT A.PROCEDURENCID, B.EMCODE, C.DESCRIPTION, D.PROVIDERNAME," & vbNewLine & _
            "E.NAME AS CLINICNAME, E.MEPRS_CODE, E.DMIS_ID_CODE" & vbNewLine & _
        "FROM (" & sSubQueryA & ") A, " & vbNewLine & _
            "(" & sSubQueryB & ") B, " & vbNewLine & _
            "(" & sSubQueryC & ") C, " & vbNewLine & _
            "(" & sSubQueryD & ") D, CLINIC E" & vbNewLine & _
        "WHERE A.PROCEDURENCID = B.NCID AND A.PROCEDURENCID = C.NCID" & vbNewLine & _
            "AND A.CLINICIAN_NCID = D.NCID AND E.NCID=A.CLINIC_NCID" & vbNewLine



        If oReportItem.Criteria2 = Diagnosis Then

            Dim sSub_SubQueryB As String
            
            sSub_SubQueryB = "SELECT E2.ENCOUNTER_NUMBER" & vbNewLine & _
                            "FROM EVENT E2, CLINICAL_OBSERVATION CO, OBSERVATION_VALUE OV" & vbNewLine & _
                            "WHERE E2.ENCOUNTER_NUMBER IS NOT NULL AND" & vbNewLine & _
                                "E2.ENCOUNTER_NUMBER <> 0 AND" & vbNewLine & _
                                "E2.EVENT_ID=CO.EVENT_ID AND" & vbNewLine & _
                                "CO.OBSERVATION_NCID=658 AND" & vbNewLine & _
                                "CO.OBSERVATION_SEQ = OV.OBSERVATION_SEQ AND" & vbNewLine & _
                                "CO.EVENT_ID = OV.EVENT_ID AND " & vbNewLine & _
                                "OV.VALUE_NCID =1009771" & vbNewLine
                                
            sSubQueryA = "SELECT SUB_A.PROCEDURENCID, SUB_A.CLINICIAN_NCID, SUB_A.CLINIC_NCID, SUB_A.ENCOUNTER_NUMBER" & vbNewLine & _
                        "FROM (" & sSubQueryA & ") SUB_A, (" & sSub_SubQueryB & ") SUB_B" & vbNewLine & _
                        "WHERE SUB_A.ENCOUNTER_NUMBER = SUB_B.ENCOUNTER_NUMBER" & vbNewLine

            sStmt = "SELECT A.PROCEDURENCID, B.EMCODE, C.DESCRIPTION, D.PROVIDERNAME," & vbNewLine & _
                        "E.NAME AS CLINICNAME, E.MEPRS_CODE, E.DMIS_ID_CODE" & vbNewLine & _
                    "FROM (" & sSubQueryA & ") A, " & vbNewLine & _
                        "(" & sSubQueryB & ") B, " & vbNewLine & _
                        "(" & sSubQueryC & ") C, " & vbNewLine & _
                        "(" & sSubQueryD & ") D, CLINIC E" & vbNewLine & _
                    "WHERE A.PROCEDURENCID = B.NCID AND A.PROCEDURENCID = C.NCID" & vbNewLine & _
                        "AND A.CLINICIAN_NCID = D.NCID AND E.NCID=A.CLINIC_NCID" & vbNewLine

        End If
 

Set oRec = goDBDAS.OpenRecordset(sStmt)
DoEvents

'--------------------------------
'-- Set Data
'----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection

    moCollRecSet.Add oRec, "EandM"

'-- Create Report Object
'-----------------------------------------------------------------------------
    If Not CreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo CleanUp

    DoEvents

DataGetEandM = True
GoTo CleanUp
ErrHandler:
    If Err = -2147220874 Then
        MsgBxARMd "The selected report has timed out, please re-run this report again.", vbOKOnly + vbInformation
        Err.Clear
    Else
        Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.DataGetEandM", "CHCSII_Reports", vbExclamation)
    End If
CleanUp:
    Set oRec = Nothing
    Exit Function
End Function

Private Function DataGetDisposition(sReportName As String, _
                              oReportItem As ReportItem, _
                              oReport As CRAXDRT.Report, _
                              oCollFormulas As Collection) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler
DataGetDisposition = False

Dim sStmt As String
Dim i As Integer
Dim oRec As ADODB.Recordset
Dim sDateFilter As String
Dim sFilter As String

'-- Get Filter Information for SQL
'----------------------------------------------------------------------------
If Not GetSQLDateRange(oReportItem, sDateFilter) Then GoTo CleanUp
If Not GetSQLCriteria(oReportItem, 1, sFilter) Then GoTo CleanUp

'-- Select
sStmt = "SELECT /*+ RULE no_index(e event_x4)  */" & vbNewLine & _
                "ov.value_ncid DispositionNCID, " & vbNewLine & _
                "TRIM(RDIS.REPRESENTATION) Description, " & vbNewLine & _
                "p.representation as ProviderName," & vbNewLine & _
                "c.name as ClinicName," & vbNewLine & _
                "c.meprs_code," & vbNewLine & _
                "c.dmis_id_code" & vbNewLine

'-- From
sStmt = sStmt & _
            "From" & vbNewLine & _
                "rsform_context pc," & vbNewLine & _
                "rsform p," & vbNewLine & _
                "clinic c," & vbNewLine & _
                "RSFORM_CONTEXT RCDIS," & vbNewLine & _
                "RSFORM RDIS," & vbNewLine & _
                "observation_value ov," & vbNewLine & _
                "clinical_observation co," & vbNewLine & _
                "event e," & vbNewLine & _
                "appointment a" & vbNewLine

'-- Where
sStmt = sStmt & _
        "Where " & vbNewLine & _
            "pc.preferred_score=0 and" & vbNewLine & _
            "pc.context_ncid=2000 and" & vbNewLine & _
            "pc.rsform_id = p.rsform_id and" & vbNewLine & _
            "p.ncid = a.clinician_ncid and" & vbNewLine & _
            "c.ncid=a.clinic_ncid and" & vbNewLine & _
            "RCDIS.preferred_score=0 and" & vbNewLine & _
            "RCDIS.CONTEXT_NCID = 2000 AND " & vbNewLine & _
            "RCDIS.RSFORM_ID = RDIS.RSFORM_ID AND " & vbNewLine & _
            "RDIS.NCID = ov.value_ncid AND" & vbNewLine

'549,204847,204848,204849,14511049,14511050,14511051
'204847,14503719,540,204848,546,204849,549 OLD

sStmt = sStmt & _
            "ov.value_ncid in (549,204847,204848,204849,14511049,14511050,14511051) and" & vbNewLine & _
            "ov.observation_Seq = co.observation_seq and" & vbNewLine & _
            "ov.event_id = co.event_id and" & vbNewLine & _
            "co.observation_ncid=373 and" & vbNewLine & _
            "co.event_id = e.event_id and" & vbNewLine & _
            "e.event_subtype_ncid=184917 and" & vbNewLine & _
            "e.event_type_ncid = 209614 and" & vbNewLine & _
            "e.encounter_number=a.encounter_number and" & vbNewLine & _
            "a.encounter_number is not null and" & vbNewLine & _
            "a.encounter_number <> '0' and" & vbNewLine & _
            "a.Encounter_Status in (4,5,6) and" & vbNewLine

            If oReportItem.Criteria1 <> MTF Then
                sStmt = sStmt & _
                    sDateFilter & " and" & vbNewLine & _
                    sFilter
            Else
                sStmt = sStmt & _
                    sFilter & " and" & vbNewLine & _
                    sDateFilter
            End If

'        If oReportItem.Criteria2 <> NoCriteria Then
'            sStmt = sStmt & " and" & vbNewLine & "a.encounter_number = e.encounter_number and"
'        End If

        If oReportItem.Criteria2 = Diagnosis Then

            sStmt = sStmt & vbNewLine & _
                "and" & vbNewLine & "a.encounter_number in" & vbNewLine & _
                    "(select" & vbNewLine & _
                        "e2.encounter_number" & vbNewLine & _
                    "from" & vbNewLine & _
                        "event e2," & vbNewLine & _
                        "clinical_observation co," & vbNewLine & _
                        "observation_value ov" & vbNewLine

            sStmt = sStmt & _
                    "where" & vbNewLine & _
                        "e2.encounter_number is not null and" & vbNewLine & _
                        "e2.encounter_number <> 0 and" & vbNewLine & _
                        "e2.event_id=co.event_id and" & vbNewLine & _
                        "co.observation_ncid=658 and" & vbNewLine & _
                        "co.observation_seq = ov.observation_seq and" & vbNewLine & _
                        "co.event_id = ov.event_id and" & vbNewLine & _
                        "ov.value_ncid =" & oReportItem.CriteriaSel2.Item(1).NCID & ")"

        End If
 

Set oRec = goDBDAS.OpenRecordset(sStmt)
DoEvents

'--------------------------------
'-- Set Data
'----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection

    moCollRecSet.Add oRec, "disposition"

'-- Create Report Object
'-----------------------------------------------------------------------------
    If Not CreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo CleanUp

    DoEvents

DataGetDisposition = True
GoTo CleanUp
ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.DataGetDisposition", "CHCSII_Reports", vbExclamation)
CleanUp:
    Set oRec = Nothing
    Exit Function
End Function


Private Function DataGetRad(sReportName As String, _
                            oReportItem As ReportItem, _
                            oReport As CRAXDRT.Report, _
                            oCollFormulas As Collection) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler
DataGetRad = False

Dim sStmt As String
Dim i As Integer
Dim oRec As ADODB.Recordset
Dim sDateFilter As String
Dim sFilter As String
Dim sFilter2 As String

'-- Get Filter Information for SQL
'----------------------------------------------------------------------------
If Not GetSQLDateRange(oReportItem, sDateFilter) Then GoTo CleanUp
If Not GetSQLCriteria(oReportItem, 1, sFilter) Then GoTo CleanUp

'-- Select
sStmt = "SELECT /*+ RULE no_index(e event_x4) */" & vbNewLine & _
                "p.representation as ProviderName," & vbNewLine & _
                "c.name as ClinicName," & vbNewLine & _
                "c.meprs_code," & vbNewLine & _
                "c.dmis_id_code," & vbNewLine & _
                "oer.order_name, " & vbNewLine & _
                "oer.order_ien, " & vbNewLine & _
                "0 as event_id, " & vbNewLine & _
                "a.patient_unit_number patient_unit_number, " & vbNewLine & _
                "a.encounter_number, " & vbNewLine & _
                "a.facility_ncid mtf_ncid, " & vbNewLine & _
                "a.clinic_ncid clinic_ncid," & vbNewLine & _
                "a.clinician_ncid provider_ncid," & vbNewLine & _
                "a.APPT_DATE_TIME encounter_date" & vbNewLine

'-- From
sStmt = sStmt & _
        "FROM" & vbNewLine & _
            "rsform_context pc," & vbNewLine & _
            "rsform p," & vbNewLine & _
            "clinic c," & vbNewLine & _
            "oe_rad oer," & vbNewLine

        If oReportItem.Criteria2 = Diagnosis Then
            sStmt = sStmt & "semantic_link_element sle," & vbNewLine & _
                            "semantic_link sl," & vbNewLine & _
                            "observation_value ov," & vbNewLine & _
                            "clinical_observation co," & vbNewLine & _
                            "event e," & vbNewLine
        End If

sStmt = sStmt & "appointment a" & vbNewLine


'-- Where
sStmt = sStmt & _
        "Where" & vbNewLine & _
            "pc.preferred_score=0 and" & vbNewLine & _
            "pc.context_ncid=2000 and" & vbNewLine & _
            "pc.rsform_id = p.rsform_id and" & vbNewLine & _
            "p.ncid = a.clinician_ncid and" & vbNewLine & _
            "c.ncid=a.clinic_ncid and" & vbNewLine


        If oReportItem.Criteria2 = Diagnosis Then
            '-- In this case e2 represents a Diagnosis event that is linked to a _
                procedure via semantic links
            sStmt = sStmt & _
                    "oer.OE_RAD_DATA_ID = sle.link_event_id and" & vbNewLine & _
                    "sle.element_type_ncid=108702 and" & vbNewLine & _
                    "sle.semantic_link_ord = sl.semantic_link_ord and" & vbNewLine & _
                    "sle.event_id = sl.event_id and" & vbNewLine & _
                    "sl.relationship_ncid=205653 and" & vbNewLine & _
                    "sl.event_id = ov.event_id and" & vbNewLine & _
                    "ov.value_ncid=" & oReportItem.CriteriaSel2.Item(1).NCID & " and" & vbNewLine & _
                    "ov.observation_seq = co.observation_seq and " & vbNewLine & _
                    "ov.event_id = co.event_id and" & vbNewLine & _
                    "co.observation_ncid=658 and" & vbNewLine & _
                    "co.event_id = e.event_id and" & vbNewLine & _
                    "e.event_type_ncid=90803 and" & vbNewLine & _
                    "e.Encounter_Number = a.Encounter_Number and" & vbNewLine
        Else
            sStmt = sStmt & "oer.Enc_Num = a.Encounter_Number and" & vbNewLine
        End If
        
        sStmt = sStmt & _
            "a.Encounter_Status in (4,5,6) and" & vbNewLine & _
            sDateFilter & " and" & vbNewLine & _
            sFilter


Set oRec = goDBDAS.OpenRecordset(sStmt)
DoEvents

'--------------------------------
'-- Set Data
'----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection

    moCollRecSet.Add oRec, "rad"

'-- Create Report Object
'-----------------------------------------------------------------------------
    If Not CreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo CleanUp


    DoEvents

DataGetRad = True
GoTo CleanUp
ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.DataGetRad", "CHCSII_Reports", vbExclamation)
CleanUp:
    Set oRec = Nothing
    Exit Function
End Function

Private Function DataGetConsult(sReportName As String, _
                                oReportItem As ReportItem, _
                                oReport As CRAXDRT.Report, _
                                oCollFormulas As Collection) As Boolean

    If Not DebugMode Then On Error GoTo ErrHandler
    DataGetConsult = False
    
    Dim sStmt As String
    Dim i As Integer
    Dim oRec As ADODB.Recordset
    Dim sDateFilter As String
    Dim sFilter As String
    Dim sFilter2 As String
    
    Dim sSubQueryA               As String
    Dim sSubQueryB               As String
    Dim sSubQueryC               As String
    
    '-- Get Filter Information for SQL    '----------------------------------------------------------------------------
    If Not GetSQLDateRange(oReportItem, sDateFilter) Then GoTo CleanUp
    If Not GetSQLCriteria(oReportItem, 1, sFilter) Then GoTo CleanUp
    'sFilter = "co.RECV_PROVIDER_NCID = " & oReportItem.CriteriaSel1.Item(1).NCID
    '--- Sub Query A: get CONSULT
    sSubQueryA = "SELECT c.name as ClinicName, c.meprs_code AS MEPRSCODE, c.dmis_id_code, a.patient_unit_number patient_unit_number,  a.encounter_number, " & vbNewLine & _
                "a.facility_ncid mtf_ncid,  a.clinic_ncid clinic_ncid, a.clinician_ncid provider_ncid, a.APPT_DATE_TIME encounter_date, " & vbNewLine & _
                "co.REF_PROVIDER_NCID ref_provider_ncid, co.SEND_REQUEST_TO,  co.recv_entity_ncid, Decode(ap.Name,NULL,'Unknown',ap.Name) Procedure_Name, co.RECV_PROVIDER_NCID " & vbNewLine & _
                "FROM clinic c, Ancillary_Procedure ap, Consult_Order co, appointment a" & vbNewLine & _
                "WHERE c.ncid=a.clinic_ncid and" & vbNewLine
    
    sSubQueryA = sSubQueryA & "ap.facility_ncid=a.facility_ncid and ap.IEN (+) = co.anc_procedure_ien and" & vbNewLine & _
                "co.Enc_Num = a.Encounter_Number and" & vbNewLine & _
                sDateFilter & " and" & vbNewLine & sFilter
'                "a.Encounter_Status in (4,5,6) and" & vbNewLine & _

    '--- Sub Query B: get RECV_ENTITY
    sSubQueryB = "SELECT re_r.NCID, TRIM(re_r.REPRESENTATION) RECV_ENTITY" & vbNewLine & _
                "FROM rsform_context re_rc, rsform re_r " & vbNewLine & _
                "where re_rc.preferred_score=0 and" & vbNewLine & _
                "re_rc.context_ncid=2000 and" & vbNewLine & _
                "re_rc.rsform_id=re_r.rsform_id"
                
    '--- Sub Query C: get CLINIC
    sSubQueryC = "SELECT p.NCID, TRIM(p.REPRESENTATION) PROVIDERNAME " & vbNewLine & _
                 "FROM rsform_context pc, rsform p " & vbNewLine & _
                 "WHERE  pc.preferred_score=0 and" & vbNewLine & _
                 "pc.context_ncid=2000 and" & vbNewLine & _
                 "pc.rsform_id = p.rsform_id"

    sStmt = "SELECT C.PROVIDERNAME, A.CLINICNAME, A.MEPRSCODE, A.DMIS_ID_CODE, A.PATIENT_UNIT_NUMBER, " & vbNewLine & _
                "A.ENCOUNTER_NUMBER, A.MTF_NCID, A.CLINIC_NCID, A.PROVIDER_NCID, A.ENCOUNTER_DATE, A.REF_PROVIDER_NCID, " & vbNewLine & _
                "B.RECV_ENTITY, A.SEND_REQUEST_TO, A.Procedure_Name" & vbNewLine & _
            "FROM (" & sSubQueryA & ") A, " & vbNewLine & _
                "(" & sSubQueryB & ") B, " & vbNewLine & _
                "(" & sSubQueryC & ") C " & vbNewLine & _
            "WHERE  C.NCID = A.RECV_PROVIDER_NCID AND B.NCID=A.CLINIC_NCID"

    If Not oReportItem.Criteria1 = Provider Then
        sStmt = sStmt & vbCrLf & "UNION" & vbCrLf
        sStmt = sStmt & "SELECT '' AS PROVIDERNAME, A.CLINICNAME, A.MEPRSCODE, A.DMIS_ID_CODE, A.PATIENT_UNIT_NUMBER, " & vbNewLine & _
                    "A.ENCOUNTER_NUMBER, A.MTF_NCID, A.CLINIC_NCID, A.PROVIDER_NCID, A.ENCOUNTER_DATE, A.REF_PROVIDER_NCID, " & vbNewLine & _
                    "B.RECV_ENTITY, A.SEND_REQUEST_TO, A.Procedure_Name" & vbNewLine & _
                "FROM (" & sSubQueryA & ") A, " & vbNewLine & _
                    "(" & sSubQueryB & ") B " & _
                "WHERE A.RECV_PROVIDER_NCID IS NULL AND B.NCID=A.CLINIC_NCID"
    End If
    
    Set oRec = goDBDAS.OpenRecordset(sStmt)
    DoEvents

    '--------------------------------
    '-- Set Data
    '----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection

    moCollRecSet.Add oRec, "consult"

    '-- Create Report Object
    '-----------------------------------------------------------------------------
    If Not CreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo CleanUp
    DoEvents

    DataGetConsult = True
    GoTo CleanUp

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.DataGetConsult", "CHCSII_Reports", vbExclamation)
CleanUp:
    Set oRec = Nothing

End Function

Private Function DataGetMed(sReportName As String, _
                            oReportItem As ReportItem, _
                            oReport As CRAXDRT.Report, _
                            oCollFormulas As Collection) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler
DataGetMed = False

Dim sStmt As String
Dim i As Integer
Dim oRec As ADODB.Recordset
Dim sDateFilter As String
Dim sFilter As String
Dim sFilter2 As String

'-- Get Filter Information for SQL
'----------------------------------------------------------------------------
If Not GetSQLDateRange(oReportItem, sDateFilter) Then GoTo CleanUp
If Not GetSQLCriteria(oReportItem, 1, sFilter) Then GoTo CleanUp

'-- Select
sStmt = "SELECT /*+ RULE no_index(e event_x4)  */" & vbNewLine & _
                "p.representation as ProviderName," & vbNewLine & _
                "c.name as ClinicName," & vbNewLine & _
                "c.meprs_code," & vbNewLine & _
                "c.dmis_id_code," & vbNewLine & _
                "oer.order_name, " & vbNewLine & _
                "oer.order_ien, " & vbNewLine & _
                "0 as event_id, " & vbNewLine & _
                "a.patient_unit_number patient_unit_number, " & vbNewLine & _
                "a.encounter_number, " & vbNewLine & _
                "a.facility_ncid mtf_ncid, " & vbNewLine & _
                "a.clinic_ncid clinic_ncid," & vbNewLine & _
                "a.clinician_ncid provider_ncid," & vbNewLine & _
                "a.APPT_DATE_TIME encounter_date" & vbNewLine

'-- From
sStmt = sStmt & _
        "FROM" & vbNewLine & _
            "rsform_context pc," & vbNewLine & _
            "rsform p," & vbNewLine & _
            "clinic c," & vbNewLine & _
            "oe_rx oer," & vbNewLine

        If oReportItem.Criteria2 = Diagnosis Then
            sStmt = sStmt & "semantic_link_element sle," & vbNewLine & _
                            "semantic_link sl," & vbNewLine & _
                            "observation_value ov," & vbNewLine & _
                            "clinical_observation co," & vbNewLine & _
                            "event e," & vbNewLine
        End If

sStmt = sStmt & "appointment a" & vbNewLine


'-- Where
sStmt = sStmt & _
        "Where" & vbNewLine & _
            "pc.preferred_score=0 and" & vbNewLine & _
            "pc.context_ncid=2000 and" & vbNewLine & _
            "pc.rsform_id = p.rsform_id and" & vbNewLine & _
            "p.ncid = a.clinician_ncid and" & vbNewLine & _
            "c.ncid=a.clinic_ncid and" & vbNewLine


        If oReportItem.Criteria2 = Diagnosis Then
            '-- In this case e2 represents a Diagnosis event that is linked to a _
                procedure via semantic links
            sStmt = sStmt & _
                    "oer.OE_RX_DATA_ID = sle.link_event_id and" & vbNewLine & _
                    "sle.element_type_ncid=108702 and" & vbNewLine & _
                    "sle.semantic_link_ord = sl.semantic_link_ord and" & vbNewLine & _
                    "sle.event_id = sl.event_id and" & vbNewLine & _
                    "sl.relationship_ncid=205071 and" & vbNewLine & _
                    "sl.event_id = ov.event_id and" & vbNewLine & _
                    "ov.value_ncid=" & oReportItem.CriteriaSel2.Item(1).NCID & " and" & vbNewLine & _
                    "ov.observation_seq = co.observation_seq and " & vbNewLine & _
                    "ov.event_id = co.event_id and" & vbNewLine & _
                    "co.observation_ncid=658 and" & vbNewLine & _
                    "co.event_id = e.event_id and" & vbNewLine & _
                    "e.event_type_ncid=90803 and" & vbNewLine & _
                    "e.Encounter_Number = a.Encounter_Number and" & vbNewLine
        Else
            sStmt = sStmt & "oer.Enc_Num = a.Encounter_Number and" & vbNewLine
        End If
        
        sStmt = sStmt & _
            "a.Encounter_Status in (4,5,6) and" & vbNewLine & _
            sDateFilter & " and" & vbNewLine & _
            sFilter

 
Set oRec = goDBDAS.OpenRecordset(sStmt)
DoEvents

'--------------------------------
'-- Set Data
'----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection

    moCollRecSet.Add oRec, "med"

'-- Create Report Object
'-----------------------------------------------------------------------------
    If Not CreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo CleanUp


    DoEvents

DataGetMed = True
GoTo CleanUp
ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.DataGetMed", "CHCSII_Reports", vbExclamation)
CleanUp:
    Set oRec = Nothing
    Exit Function
End Function



Public Function BuildReport(sReportName As String, _
                            oReportItem As ReportItem, _
                            oReport As CRAXDRT.Report, _
                            oCollFormulas As Collection) As Boolean

On Error GoTo ErrHandler
 
BuildReport = False

DoEvents

'-- Get Data for the Report
'---------------------------------------------------------------------------------------
Select Case oReportItem.Report
    Case LabRep
        If Not DataGetLab(sReportName, oReportItem, oReport, oCollFormulas) Then GoTo CleanUp
    Case RadRep
        If Not DataGetRad(sReportName, oReportItem, oReport, oCollFormulas) Then GoTo CleanUp
    Case MedRep
        If Not DataGetMed(sReportName, oReportItem, oReport, oCollFormulas) Then GoTo CleanUp
    Case DiagnosisRep
        If Not DataGetDiagnosis(sReportName, oReportItem, oReport, oCollFormulas) Then GoTo CleanUp
    Case Consults
        If Not DataGetConsult(sReportName, oReportItem, oReport, oCollFormulas) Then GoTo CleanUp
    Case PrimaryDiagnosisRep
        If Not DataGetPrimaryDiagnosis(sReportName, oReportItem, oReport, oCollFormulas) Then GoTo CleanUp
    Case ProcedureRep
        If Not DataGetProcedure(sReportName, oReportItem, oReport, oCollFormulas) Then GoTo CleanUp
    Case EandMRep
        If Not DataGetEandM(sReportName, oReportItem, oReport, oCollFormulas) Then GoTo CleanUp
    Case DispositionRep
        If Not DataGetDisposition(sReportName, oReportItem, oReport, oCollFormulas) Then GoTo CleanUp
    Case Else
End Select

BuildReport = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepRegistryCandidates.BuildReport", "CHCSII_Reports", vbExclamation
CleanUp:
    
End Function

Private Sub Class_Initialize()
    '--- SCR-ARM-16148 MSolano 9/10/01; COMMON TO MULTIPLE ARM-SCRs
    '--- Init routine should be called from first top level creatable coclass
    '--- It can be called multiple times, but will actually run only once.
    '--- Instrument all creatable classes in case dll's multiple entry
    '--- points are actually used
     Call ARM_InitAllTransactionsForCHCSIIReports
End Sub


