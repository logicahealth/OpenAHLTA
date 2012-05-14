VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Top20DrugsByCost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private moCollRecSet As Collection
Private moCollFormulas As Collection

Private Enum ColumnEnum
   eDrugName
   eProviderName
End Enum

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
'  Return Value:  A SQL statment for that retrieves the list of repeat vistors grouped by provider
'***********************************************************************************************************
Private Function BuildReportSQL(ByRef objReportItem As ReportItem, _
                                ByRef colFormulas As Collection, _
                                ByVal blnGetCurrentMonthOnly As Boolean, _
                                ByVal eGroupByColumn As ColumnEnum, _
                                ByVal lngRowCount As Long) As String
                                
   On Error GoTo Proc_Exit
   
   Dim strSQL                 As String
   Dim strDateInfo            As String
   Dim strWHERE               As String
   Dim strGROUPBY             As String
   Dim strHeader2             As String
   
   
   Select Case eGroupByColumn
      Case eDrugName
           strGROUPBY = "GROUP BY CDRPLUS.DRUG_COST_DATA.GENERIC_NAME, CDRPLUS.PROVIDER.NAME  "
      Case eProviderName
           strGROUPBY = "GROUP BY CDRPLUS.PROVIDER.NAME, CDRPLUS.DRUG_COST_DATA.GENERIC_NAME "
   End Select
   
   BuildReportSQL = vbNullString

   'Filter on Provider or Clinic
    If objReportItem.Criteria1 = Provider Then
      strWHERE = "MEDICATION_ORDERS.ORDERING_PROVIDER_NCID= '" & objReportItem.CriteriaSel1.Item(1).NCID & "' "
      If Not blnGetCurrentMonthOnly Then
         strHeader2 = "Provider: " & objReportItem.CriteriaSel1.Item(1).Name
      End If
   ElseIf objReportItem.Criteria1 = Clinic Then
      strWHERE = "MEDICATION_ORDERS.Entering_Location_NCID = '" & objReportItem.CriteriaSel1(1).NCID & "' "
      If Not blnGetCurrentMonthOnly Then
         strHeader2 = "Clinic: " & objReportItem.CriteriaSel1.Item(1).Name
      End If
   End If
   
   If objReportItem.Criteria2 = Top20Drug Then
      'Filter on the Drug NDC
      strWHERE = strWHERE & " AND CDRPLUS.DRUG_COST_DATA.NDC = '" & objReportItem.CriteriaSel2(1).NDC & "' "
      'Add drug name for the header
      strHeader2 = strHeader2 & "   Drug " & _
                   objReportItem.CriteriaSel2(1).Sequence_Position & " Of The Top 20:" & _
                   objReportItem.CriteriaSel2(1).Drug_Name
    
      'Filter on date
      If IsDate(objReportItem.CriteriaSel2(1).StartDate) Then
         strWHERE = strWHERE & " AND MEDICATION_ORDERS.START_DATE >= " & _
                                 SQLDate(objReportItem.CriteriaSel2(1).StartDate) & " "
         strDateInfo = "From " & Format(objReportItem.DateFrom, "mmm yyyy") & "   "
      End If
   
      If IsDate(objReportItem.CriteriaSel2(1).StartDate) Then
         strWHERE = strWHERE & " AND MEDICATION_ORDERS.START_DATE <= " & _
                                 SQLDate(objReportItem.CriteriaSel2(1).EndDate) & " "
         strDateInfo = strDateInfo & "Through " & Format(objReportItem.DateTo, "mmm yyyy")
      End If
   
      On Error Resume Next
      colFormulas.Remove "DateInfo"
      colFormulas.Add CrystalQuote(strDateInfo), "DateInfo"  'Add date info to formula collection
   End If
   
   On Error Resume Next
   colFormulas.Remove "Header2"
   colFormulas.Add CrystalQuote(strHeader2), "Header2"
   Err.Clear
   On Error GoTo Proc_Exit
  
  
   If blnGetCurrentMonthOnly Then
      strWHERE = strWHERE & "AND MONTHS_BETWEEN(SYSDATE, MEDICATION_ORDERS.START_DATE) = 1 "
   End If
   
   strSQL = "SELECT Drug_Name, TO_CHAR(TotalCost) AS TotalCost, To_Char(AverageCost) AS AverageCost, Provider_Name " & _
            "FROM (" & _
                 "SELECT CDRPLUS.DRUG_COST_DATA.GENERIC_NAME AS Drug_Name, CDRPLUS.PROVIDER.NAME As Provider_Name, " & _
                     "SUM(CDRPLUS.MEDICATION_ORDERS.QUANTITY * CDRPLUS.DRUG_COST_DATA.UNIT_PRICE) AS TotalCost, " & _
                     "Avg(CDRPLUS.MEDICATION_ORDERS.QUANTITY * CDRPLUS.DRUG_COST_DATA.UNIT_PRICE) AS AverageCost " & _
                  "FROM CDRPLUS.DRUG_COST_DATA, CDRPLUS.MEDICATION_ORDERS, CDRPLUS.PROVIDER " & _
                  "WHERE CDRPLUS.DRUG_COST_DATA.NDC = CDRPLUS.MEDICATION_ORDERS.DRUG_ALTERNATE_IDENTIFIER " & _
                     "AND CDRPLUS.MEDICATION_ORDERS.ORDERING_PROVIDER_NCID = CDRPLUS.PROVIDER.NCID " & _
                     "AND " & strWHERE & strGROUPBY & _
                 "ORDER BY SUM(CDRPLUS.MEDICATION_ORDERS.QUANTITY * CDRPLUS.DRUG_COST_DATA.UNIT_PRICE) DESC " & _
                  ") SubQuery1 "
                  
   If lngRowCount > 0 Then
        strSQL = strSQL & "WHERE ROWNUM < " & lngRowCount + 1
   End If
   
    
   
   BuildReportSQL = strSQL
   
   
Proc_Exit:
   If Err.Number <> 0 Then
      BuildReportSQL = vbNullString
      Err.Raise Err.Number, Err.Source, Err.Description
   End If
   Exit Function
End Function




'***********************************************************************************************************
'  Function Name: BuildReport
'
'  Description: This method:
'                 1.) Calls the BuildReportSQL method to get the SQL for the Top20DrugsByCost report
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
   Dim objTotalRecordset      As ADODB.Recordset
   Dim objMonthlyRecordset    As ADODB.Recordset
   Dim objCrystal             As Crystal
   Dim strHeader1             As String
   
   
   BuildReport = False 'Assume false
   
    '-- Get Data    '----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection
   
   Select Case objReportItem.Report
   
      Case Top20DrugsByCostGraph
         strSQL = BuildReportSQL(objReportItem, colFormulas, False, eDrugName, 20) 'Get SQL
         Set objTotalRecordset = goDBDAS.OpenRecordset(strSQL)  'Create recordset for main report
         strSQL = BuildReportSQL(objReportItem, colFormulas, True, eDrugName, 20) 'Get SQL. Pass true to the GetCurrentMonthOnly parameter
         Set objMonthlyRecordset = goDBDAS.OpenRecordset(strSQL)  'Create recordset for sub report
         strHeader1 = "Top 20 Drugs Cost"
         
      Case Top20DrugsByCostAverageGraph
         strSQL = BuildReportSQL(objReportItem, colFormulas, False, eProviderName, 0) 'Get SQL
         Set objTotalRecordset = goDBDAS.OpenRecordset(strSQL)  'Create recordset for main report
      End Select
      moCollRecSet.Add objTotalRecordset, "TotalVsitsMTF"
   
   Set objReport = Nothing   'In case the object has been set prior
   
   If Not CreateReport(strReportName & ".rpt", objReport, colFormulas, moCollRecSet) Then
      BuildReport = False
      GoTo Proc_Exit
   Else
      'Set data source for sub report.  This is the Top20Drugs total
      objReport.Database.SetDataSource objTotalRecordset, 3, 1
      'Set data source for sub report.  This is the Top20Drugs for the Current Month
      If objReportItem.Report = Top20DrugsByCostGraph Then
         objReport.OpenSubreport("MonthTotalCost").Database.SetDataSource objMonthlyRecordset, 3, 1
      End If
      
      BuildReport = True
   End If
   

Proc_Exit:
   Set objTotalRecordset = Nothing  'Destroy reference to object+
   Set objMonthlyRecordset = Nothing  'Destroy reference to object
   Set colFormulas = Nothing     'Destroy reference to object
   
   If Err.Number <> 0 Then
      gobjshared.ShowVBError Err.Number, Err.Description, "Top20DrugsByCost.BuildReport", _
                             "CHCSII_Reports", vbExclamation
      'Resume Next
   End If
   Exit Function
End Function





Private Function HighlightTop50(ByRef objRecordset As Recordset, ByVal objReportItem As ReportItem) As Boolean

   Dim objDrugCostForMTF      As DrugCostForMTF
   Dim strClinicNCID          As String
   
   
   Set objDrugCostForMTF = New DrugCostForMTF
         
    If objReportItem.Criteria1 = Provider Then
      strClinicNCID = objReportItem.CriteriaSel1.Item(1).Clinic_NCID
   ElseIf objReportItem.Criteria1 = Clinic Then
      strClinicNCID = objReportItem.CriteriaSel1(1).NCID
   End If
   
   
  HighlightTop50 = objDrugCostForMTF.HighlightMTFRecords(objRecordset, "Drug_Name", strClinicNCID, _
                                                         objReportItem.DateFrom, objReportItem.DateTo, 50)
   
   
Proc_Exit:
   Set objDrugCostForMTF = Nothing
   
   If Err.Number <> 0 Then
      HighlightTop50 = False
      Err.Raise Err.Number, Err.Source, Err.Description
   End If
   Exit Function
End Function



