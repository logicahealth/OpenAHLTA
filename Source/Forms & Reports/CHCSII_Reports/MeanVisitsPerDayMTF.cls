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
Attribute VB_Name = "MeanVisitsPerDayMTF"
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
'                 1.) Calls the BuildReportSQL method to get the SQL for the RepeatVistsMTF report
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
   
    Dim objRecordset As ADODB.Recordset
    Dim strSQL As String
    
    BuildReport = False 'Assume false
    '-- Get Data    '----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection
    
    strSQL = BuildReportSQL(objReportItem, colFormulas)  'Get SQL
    If strSQL = "" Then Exit Function
    
   Set objRecordset = goDBDAS.OpenRecordset(strSQL)  'Create recordset
    moCollRecSet.Add objRecordset, "TotalVsitsMTF"
   
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
      gobjshared.ShowVBError Err.Number, Err.Description, "Repeat Visitors.BuildReport", _
                             "CHCSII_Reports", vbExclamation
   End If

End Function

'***********************************************************************************************************
'  Function Name: BuildReportSQL
'
'  Description: This method concatenates an SQL string using the members in the ReportItem
'               object that is passed into it
'
'  Parameters:
'    objReportItem - contains members that are used to filter
'    colFormulas - collection of formulas that the report will use
'
'  Return Value:  An SQL statment for that retrieves the list of repeat vistors grouped by provider
'***********************************************************************************************************
Private Function BuildReportSQL(ByRef objReportItem As ReportItem, _
                                ByRef colFormulas As Collection) As String
                                
   On Error GoTo Proc_Exit
   
   Dim strSQL                 As String
   Dim strSubQuery            As String
   Dim strDateInfo            As String
   Dim strFromDate            As String
   Dim strToDate              As String
   Dim strPatientField        As String
   Dim strWHERE               As String
   Dim sQuestion              As String
   Dim sHeader                As String
   Dim iResponse              As Integer
   Dim sFilter                As String
   
   BuildReportSQL = vbNullString

   
   strFromDate = Format(objReportItem.DateFrom, "mm/dd/yyyy")
   strToDate = Format(objReportItem.DateTo, "mm/dd/yyyy")
     
    If DateDiff("d", strFromDate, strToDate) > 31 Then
       sQuestion = "The data graphed may appear unreadable if the number of days in the date range exceed 31, proceed?"
       If MsgBxARMd(sQuestion, vbYesNo + vbQuestion) = vbNo Then Exit Function
    End If
   
    If objReportItem.Criteria1 = MTF Then
        sFilter = " WHERE  Facility_NCID = " & objReportItem.CriteriaSel1.Item(1).NCID & " " & vbCrLf
    ElseIf objReportItem.Criteria1 = MTFDMIS Then
        sFilter = " WHERE  CLINIC_NCID IN (" & GetClinicNCIDsFromDMIS(objReportItem.CriteriaSel1) & ") " & vbCrLf
    End If
   
   strSubQuery = "SELECT Count(Patient_Unit_Number) AS PATIENTNUMBEROFVISITS, " & vbCrLf & "TRUNC(Appt_Date_Time) AS APPT_DATE " & vbCrLf & _
                 "FROM APPOINTMENT " & sFilter & " " & _
                 "GROUP BY TRUNC(APPT_DATE_TIME), Patient_Unit_Number " & vbCrLf & _
                 "HAVING (COUNT(PATIENT_UNIT_NUMBER) > 2) " & vbCrLf & _
                     "AND (TRUNC(APPT_DATE_TIME) >= TO_DATE('" & strFromDate & "','MM/DD/YYYY')  " & vbCrLf & _
                     "AND TRUNC(APPT_DATE_TIME) <= TO_DATE('" & strToDate & "','MM/DD/YYYY'))  " & vbCrLf
                     
   strSQL = "SELECT to_char(COUNT(PATIENTNUMBEROFVISITS)) AS RepeatVisitMean, " & vbCrLf & "TO_CHAR(APPT_DATE, 'YYYY/MM/DD') AS APPT_DATE " & vbCrLf & _
           "FROM (" & strSubQuery & ") GROUP BY APPT_DATE"

   strDateInfo = "From  " & Format(strFromDate, "dd mmm yyyy") & "   " & "Through  " & Format(strToDate, "dd mmm yyyy")
         
   colFormulas.Add CrystalQuote(strDateInfo), "DateInfo"
    If objReportItem.Criteria1 = MTFDMIS Then
        On Error Resume Next
        colFormulas.Remove ("Header2")
        For iResponse = 1 To objReportItem.CriteriaSel1.Count
            If Not sHeader = "" Then sHeader = sHeader & ", "
            sHeader = sHeader & objReportItem.CriteriaSel1(iResponse).MTF_NAME
        Next
        sHeader = "DMIS IDs selected: " & sHeader
        colFormulas.Add CrystalQuote(sHeader), "Header2"
    End If
   
   BuildReportSQL = strSQL
   
Proc_Exit:
   If Err.Number <> 0 Then
      BuildReportSQL = vbNullString
      Err.Raise Err.Number, Err.Source, Err.Description
   End If
   Exit Function
End Function

Private Sub Class_Terminate()
    Set moCollRecSet = Nothing
    Set moCollFormulas = Nothing
End Sub


