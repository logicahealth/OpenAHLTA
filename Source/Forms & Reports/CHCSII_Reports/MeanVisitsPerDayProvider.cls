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
Attribute VB_Name = "MeanVisitsPerDayProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private moCollRecSet As Collection
Private moCollFormulas As Collection

Public Function BuildReport(ByVal strReportName As String, _
                            ByVal objReportItem As ReportItem, _
                            ByRef objReport As CRAXDRT.Report, _
                            ByVal objCollFormulas As Collection) As Boolean
   
   On Error GoTo Proc_Exit
   
   Dim objRecordset           As ADODB.Recordset
   Dim colFormulas            As Collection
   Dim strSQL                 As String
   Dim strSubQuery            As String
   Dim strDateInfo            As String
   Dim strFromDate            As String
   Dim strToDate              As String
   Dim strHeader2 As String
   Dim sQuestion As String
   
   BuildReport = False
   
   strFromDate = Format(objReportItem.DateFrom, "mm/dd/yyyy")
   strToDate = Format(objReportItem.DateTo, "mm/dd/yyyy")
     
    If DateDiff("d", strFromDate, strToDate) > 31 Then
       sQuestion = "The data graphed may appear unreadable if the number of days in the date range exceed 31, proceed?"
       If MsgBxARMd(sQuestion, vbYesNo + vbQuestion) = vbNo Then Exit Function
    End If
   
   strSQL = "SELECT APPT_DATE AS APPTMT_DATE, " & vbCrLf & _
                "TO_CHAR(ROUND(AVG(VISITS), 0)) AS MEAN_VISITS_PER_DAY" & vbCrLf & _
            "FROM (" & vbCrLf & _
                "SELECT TO_CHAR(APPT_DATE_TIME,'MM/DD/YYYY') AS APPT_DATE, CLINICIAN_NCID, COUNT(*) AS VISITS" & vbCrLf & _
                "FROM APPOINTMENT_HISTORY" & vbCrLf & _
                "WHERE (MEPRS_CODE NOT like 'BIA%' AND MEPRS_CODE NOT like 'BHI%' OR MEPRS_CODE IS NULL)" & vbCrLf & _
                    "AND (APPT_STATUS = 'KEPT' OR  APPT_STATUS = 'LWOBS' OR  (APPT_STATUS = 'CANCEL' AND APPT_CANCEL_REASON = 'LWOBS'))" & vbCrLf & _
                    "AND APPT_DATE_TIME >= TO_DATE('" & strFromDate & "' , 'MM/DD/YYYY')" & vbCrLf & _
                    "AND APPT_DATE_TIME <= TO_DATE('" & strToDate & "' , 'MM/DD/YYYY')" & vbCrLf & _
                    "AND Facility_NCID = " & objReportItem.CriteriaSel1.Item(1).NCID & vbCrLf & _
                "GROUP BY TO_CHAR(APPT_DATE_TIME,'MM/DD/YYYY'), CLINICIAN_NCID " & vbCrLf & _
            ") A" & vbCrLf & _
            "GROUP BY APPT_DATE" & vbCrLf

   strDateInfo = "From  " & strFromDate & "  Through  " & strToDate
   
    '-- Get Data    '----------------------------------------------------------------------------
   Set moCollRecSet = Nothing
   Set moCollRecSet = New Collection
   
   Set objRecordset = goDBDAS.OpenRecordset(strSQL)
   moCollRecSet.Add objRecordset, "TotalVsitsMTF"

   On Error Resume Next
   objCollFormulas.Add CrystalQuote(strDateInfo), "DateInfo"
   Err.Clear
   
   Set objReport = Nothing
   
   If Not CreateReport(strReportName & ".rpt", objReport, objCollFormulas, moCollRecSet) Then
      BuildReport = False
      GoTo Proc_Exit
   Else
      objReport.Database.SetDataSource objRecordset, 3, 1 'Set data source
      BuildReport = True
   End If
   
Proc_Exit:
   Set objRecordset = Nothing
   Set colFormulas = Nothing
   
   If Not Err.Number = 0 Then
      gobjshared.ShowVBError Err.Number, Err.Description, "MeanVisitsPerDayMTF.BuildReport", "CHCSII_Reports", vbExclamation
   End If

End Function

Private Sub Class_Terminate()
    Set moCollRecSet = Nothing
    Set moCollFormulas = Nothing
End Sub


