VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RepeatVisitors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private moCollRecSet As Collection
Private moCollFormulas As Collection


'***********************************************************************************************************
'  Function Name: BuildReport
'
'  Description: This method:
'                 1.) Calls the BuildReportSQL method to get the SQL for the RepeatVistors report
'                 2.) Creates a recordset using that SQL
'                 3.) Assigns the formulas for the report
'                 4.) Sets the recordset as the data source for the report and runs the report
'
'  Parameters:
'    strReportName - Name of the report
'    objReportItem - ReportItem object used for filtering
'    objReport     - Crystal Report object
'    colFormulas   - Collection of formulas that the report will use.  This shoud not be a passed in parameter but is
'                    kept so to keep compatibility with the other classes
'
'  Return Value:  True/False indicating the success or failure of the report generation
'***********************************************************************************************************
Public Function BuildReport(ByVal strReportName As String, _
                            ByVal objReportItem As ReportItem, _
                            ByRef objReport As CRAXDRT.Report, _
                            ByVal colFormulas As Collection) As Boolean
   
   On Error GoTo Proc_Exit
   
   Dim strSQL                 As String
   Dim objRecordset           As ADODB.Recordset
   
   BuildReport = False 'Assume false

   Select Case objReportItem.Report
      Case ProviderRepeatVisitReport
          strSQL = BuildReportSQLReport(objReportItem, colFormulas)  'Get SQL
          
      Case ProviderRepeatVisitList
         strSQL = BuildReportSQLForList(objReportItem, colFormulas)  'Get SQL
   End Select
   
    '-- Get Data    '----------------------------------------------------------------------------
   Set moCollRecSet = Nothing
   Set moCollRecSet = New Collection
   
   Set objRecordset = goDBDAS.OpenRecordset(strSQL)  'Create recordset
    moCollRecSet.Add objRecordset, "ProviderRepeatVisits"

   Set objReport = Nothing   'In case the object has been set prior
   
   If Not CreateReport(strReportName & ".rpt", objReport, colFormulas, moCollRecSet) Then
      BuildReport = False
      GoTo Proc_Exit
   Else
      objReport.Database.SetDataSource objRecordset, 3, 1 'Set data source
      BuildReport = True
   End If
      
Proc_Exit:
   Set objRecordset = Nothing  'Destroy reference to object
   Set colFormulas = Nothing     'Destroy reference to object
   
   If Err.Number <> 0 Then
      gobjshared.ShowVBError Err.Number, Err.Description, "Repeat Visitors.BuildReport", "CHCSII_Reports", vbExclamation
   End If

End Function



Private Function BuildReportSQLForList(ByRef objReportItem As ReportItem, _
                                       ByRef colFormulas As Collection) As String
                                
    On Error GoTo Proc_Exit
    
    Dim strSQL                 As String
    Dim strDateInfo            As String
    Dim strWHERE               As String
    Dim strHeader2             As String
    Dim sDateFilter            As String
    
    Dim sStartDate  As String
    Dim sEndDate As String
    
    BuildReportSQLForList = vbNullString
    
    '-- Get Date Filter Information for SQL
    '----------------------------------------------------------------------------
    sStartDate = GetOracleStartDate(GetFirstDayOfMonth(objReportItem.DateFrom))
    sEndDate = GetOracleEndDate(GetLastDayOfMonth(objReportItem.DateTo))
    
    sDateFilter = sDateFilter & " AND (A.APPT_DATE_TIME >= " & sStartDate & vbCrLf & " AND A.APPT_DATE_TIME <= " & sEndDate & ") "
            
    strSQL = "SELECT (B.LAST_NAME || ', ' || B.First_Name || ' ' || B.MIDDLE_NAME)  AS PatientName," & vbCrLf & _
                "(SUBSTR(C.SSN, 1, 3)|| '-' || SUBSTR(C.SSN, 4, 2)|| '-' || SUBSTR(C.SSN, 6, 4)) AS PatientSSn," & vbCrLf & _
                "(SUBSTR(Trim(D.ID_VALUE), 1, 3)|| '-' || SUBSTR(Trim(D.ID_VALUE), 4, 2)|| '-' || SUBSTR(Trim(D.ID_VALUE), 6, 4)) as SPONSOR_SSN," & vbCrLf & _
                "TO_CHAR (Count(A.Patient_Unit_Number)) AS NUMBEROFVISITS " & vbCrLf & _
            "FROM APPOINTMENT_HISTORY A, MMI_NAME B, MMI_ID C, MMI_GENERIC_ID D" & vbCrLf & _
            "WHERE A.Patient_Unit_Number = B.Unit_Number" & vbCrLf & _
                "AND B.NAME_TYPE_CID = 187" & vbCrLf & _
                "AND A.Patient_Unit_Number = C.Unit_Number(+)" & vbCrLf & _
                "AND A.Patient_Unit_Number = D.Unit_Number(+)" & vbCrLf & _
                "AND D.ID_TYPE_NCID(+) =186334 " & vbCrLf & _
                "AND A.Clinician_NCID= '" & objReportItem.CriteriaSel1.Item(1).NCID & "'" & vbCrLf & _
                sDateFilter & vbCrLf & _
                "AND (APPT_STATUS = 'KEPT' OR  APPT_STATUS = 'LWOBS' OR  (APPT_STATUS = 'CANCEL' AND APPT_CANCEL_REASON = 'LWOBS'))" & vbCrLf & _
            "GROUP BY A.PATIENT_UNIT_NUMBER, B.LAST_NAME, B.FIRST_NAME, B.MIDDLE_NAME, C.SSN, D.ID_VALUE" & vbCrLf & _
            "HAVING COUNT(A.Patient_Unit_Number)  > 1" & vbCrLf & _
            "ORDER BY COUNT(A.Patient_Unit_Number) DESC, B.LAST_NAME || ', ' || B.First_Name || ' ' || B.MIDDLE_NAME"
            
    '--- Display from date if available
    If IsDate(objReportItem.DateFrom) Then
       strDateInfo = "From  " & Format(objReportItem.DateFrom, "mmm yyyy") & "  "
    End If
    
    'Filter on to date if available
    If IsDate(objReportItem.DateTo) Then
       strDateInfo = strDateInfo & "Through  " & Format(objReportItem.DateTo, "mmm yyyy")
    End If
    
    strHeader2 = "Provider: " & objReportItem.CriteriaSel1.Item(1).Name
             
On Error Resume Next
    colFormulas.Add CrystalQuote(strHeader2), "Header2"
    colFormulas.Add CrystalQuote(strDateInfo), "DateInfo"
    Err.Clear
    
    BuildReportSQLForList = strSQL
   
Proc_Exit:
   If Err.Number <> 0 Then
      BuildReportSQLForList = vbNullString
      Err.Raise Err.Number, Err.Source, Err.Description
   End If

End Function


          


Private Function BuildReportSQLReport(ByRef objReportItem As ReportItem, _
                                      ByRef colFormulas As Collection) As String
                                
    On Error GoTo Proc_Exit
    
    Dim strProviderQuery               As String
    Dim strProviderSubQuery            As String
    Dim strDateInfo                    As String
    Dim strPeerQuery                   As String
    Dim strPeerSubQuery                As String
    
    Dim sStartDate As String
    Dim sEndDate As String
    
    BuildReportSQLReport = vbNullString
    
    sStartDate = GetOracleStartDate(GetFirstDayOfMonth(objReportItem.DateFrom))
    sEndDate = GetOracleEndDate(GetLastDayOfMonth(objReportItem.DateTo))
    
    strProviderSubQuery = "SELECT CLINICIAN_NCID, COUNT(DISTINCT PATIENT_UNIT_NUMBER) AS Number_Of_Patients," & vbCrLf & _
                             "COUNT(PATIENT_UNIT_NUMBER) AS Repeat_Visits" & vbCrLf & _
                         "FROM APPOINTMENT_HISTORY " & vbCrLf & _
                         "WHERE Clinic_NCID = '" & objReportItem.CriteriaSel1.Item(1).NCID & "'" & vbCrLf & _
                             "AND APPT_DATE_TIME >= " & sStartDate & vbCrLf & _
                             "AND APPT_DATE_TIME <= " & sEndDate & vbCrLf & _
                             "AND CLINICIAN_NCID = '" & objReportItem.CriteriaSel2.Item(1).NCID & "'" & vbCrLf & _
                            "AND (APPT_STATUS = 'KEPT' OR  APPT_STATUS = 'LWOBS' OR  (APPT_STATUS = 'CANCEL' AND APPT_CANCEL_REASON = 'LWOBS'))" & vbCrLf & _
                         "GROUP BY CLINICIAN_NCID, Patient_Unit_Number " & vbCrLf & _
                         "HAVING COUNT(PATIENT_UNIT_NUMBER) > 1"
                         
    strProviderQuery = "SELECT REPEAT_VISITS, SUM(Number_Of_Patients) AS NUM_OF_PATIENTS" & vbCrLf & _
                       "FROM (" & strProviderSubQuery & ")" & vbCrLf & _
                       "GROUP BY REPEAT_VISITS"
    
    
    strPeerSubQuery = Replace(strProviderSubQuery, "CLINICIAN_NCID =", "CLINICIAN_NCID <>")
    
    strPeerQuery = "SELECT REPEAT_VISITS, ROUND(AVG(NUM_OF_PATIENTS), 0) AS PEER_MEAN" & vbCrLf & _
                   "FROM (" & vbCrLf & _
                                "SELECT CLINICIAN_NCID, REPEAT_VISITS, SUM(Number_Of_Patients) AS NUM_OF_PATIENTS" & vbCrLf & _
                                "FROM (" & strPeerSubQuery & ")" & vbCrLf & _
                                "GROUP BY REPEAT_VISITS, CLINICIAN_NCID" & vbCrLf & _
                         ")" & vbCrLf & _
                    "GROUP BY REPEAT_VISITS"

    BuildReportSQLReport = "SELECT TO_CHAR(PRV.REPEAT_VISITS) AS REPEAT_VISITS," & vbCrLf & _
                               "TO_CHAR(PRV.NUM_OF_PATIENTS) AS NUMBER_OF_PATIENTS, " & vbCrLf & _
                               "TO_CHAR(PEER.PEER_MEAN) AS PEER_MEAN" & vbCrLf & _
                            "FROM (" & strProviderQuery & ") PRV," & vbCrLf & _
                               "(" & strPeerQuery & ") PEER" & vbCrLf & _
                            "WHERE PRV.REPEAT_VISITS = PEER.REPEAT_VISITS(+)"


'Add text to formula collection
On Error Resume Next
    '--- From/To date string in the report formular
    If IsDate(objReportItem.DateFrom) Then
       strDateInfo = "From  " & Format(objReportItem.DateFrom, "mmm yyyy") & "   "
    End If
    
    If IsDate(objReportItem.DateTo) Then
       strDateInfo = strDateInfo & "Through  " & Format(objReportItem.DateTo, "mmm yyyy")
    End If
    
    colFormulas.Add CrystalQuote("Clinic: " & objReportItem.CriteriaSel1.Item(1).Name), "Header2"
    colFormulas.Add CrystalQuote(strDateInfo), "DateInfo"
    colFormulas.Add CrystalQuote("Provider: " & objReportItem.CriteriaSel2.Item(1).Name), "Provider"
    Err.Clear
   
Proc_Exit:
   If Err.Number <> 0 Then
      BuildReportSQLReport = vbNullString
      Err.Raise Err.Number, Err.Source, Err.Description
   End If
   Exit Function
End Function


Private Sub Class_Terminate()
    Set moCollRecSet = Nothing
    Set moCollFormulas = Nothing
End Sub


