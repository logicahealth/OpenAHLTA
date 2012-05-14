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
Attribute VB_Name = "DrugCostForMTF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private moCollRecSet As Collection
Private moCollFormulas As Collection

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
  
   Dim strDateInfo            As String
   Dim strSortColumn          As String
   Dim sttrSortDirection      As String
   Dim strSortInfo            As String
   Dim strHeader2 As String
      
   BuildReportSQL = vbNullString
   strHeader2 = vbNullString

   If IsDate(objReportItem.DateFrom) Then
      strDateInfo = "From " & Format(objReportItem.DateFrom, "mm/dd/yyyy") & "   "
   End If

   If IsDate(objReportItem.DateTo) Then
      strDateInfo = strDateInfo & "Through " & Format(objReportItem.DateTo, "mm/dd/yyyy")
   End If

   If objReportItem.Criteria1 = MTF Then
     strHeader2 = "MTF: " & objReportItem.CriteriaSel1.Item(1).Name
   End If

   If Trim(objReportItem.SortItemSelected) = vbNullString Then
      objReportItem.SortItemSelected = "AverageDispense"
      objReportItem.SortDirection = "Desc"
   End If
   
   If Trim(objReportItem.SortDirection) = "Asc" Or Trim(objReportItem.SortDirection) = vbNullString Then
      sttrSortDirection = "Ascending"
   Else
      sttrSortDirection = "Descending"
   End If
      
   Select Case objReportItem.SortItemSelected
      Case "AverageDispense"
         strSortInfo = sttrSortDirection & " Sort on Average Dispense"
      Case "Drug_Name"
         strSortInfo = sttrSortDirection & " Sort on Drug Name"
      Case "DrugClass"
         strSortInfo = sttrSortDirection & " Sort on Drug Class"
      Case "TotalCost"
         strSortInfo = sttrSortDirection & " Sort on Total Cost"
      Case "TotalNumberOfPrescriptions"
         strSortInfo = sttrSortDirection & " Sort on  Number Of Prescriptions"
      Case "AverageUnitCost"
         strSortInfo = sttrSortDirection & " Sort on  Average Unit Cost"
   End Select

   BuildReportSQL = TopDrugSQL(objReportItem.CriteriaSel1.Item(1).NCID, CDate(objReportItem.DateFrom), _
                               CDate(objReportItem.DateTo), 50, objReportItem.SortItemSelected, (sttrSortDirection = "Ascending"))
   
   On Error Resume Next
   colFormulas.Remove "Header2"
   colFormulas.Add CrystalQuote(strHeader2), "Header2"
   colFormulas.Remove "DateInfo"
   colFormulas.Add CrystalQuote(strDateInfo), "DateInfo"
   colFormulas.Remove "SortInfo"
   colFormulas.Add CrystalQuote(strSortInfo), "SortInfo"
   Err.Clear
   
Proc_Exit:
   If Err.Number <> 0 Then
      BuildReportSQL = vbNullString
      Err.Raise Err.Number, Err.Source, Err.Description
   End If

End Function


Public Function TopDrugSQL(ByVal strFacilityNCID As String, _
                           ByVal dteFromDate As Date, _
                           ByVal dteToDate As Date, _
                           ByVal lngNumberOfRows As Long, _
                           ByVal strSortColumn As String, _
                           ByVal blnAscendingSort As Boolean) As String
                                
   Dim sStartDate As String
   Dim sEndDate As String
   
   On Error GoTo Proc_Exit
    sStartDate = GetOracleStartDate(GetFirstDayOfMonth(dteFromDate))
    sEndDate = GetOracleEndDate(GetLastDayOfMonth(dteToDate))
   
   Const strSELECT   As String = "SELECT /*+ ordered */ dcd.GENERIC_NAME AS Drug_Name, " & vbCrLf & _
                                    "Avg(dcd.UNIT_PRICE) AS AverageUnitCost, " & vbCrLf & _
                                    "Count(mo.Order_Number) AS TotalNumberOfPrescriptions, " & vbCrLf & _
                                    "Sum(mo.QUANTITY * dcd.UNIT_PRICE) AS TotalCost, " & vbCrLf & _
                                    "Avg(Decode(mo.Give_Amount, Null,0,mo.Give_Amount)) AS AverageDispense, " & vbCrLf & _
                                    "dcd.GENSEQNUM AS DrugClass,  " & vbCrLf & _
                                    "COUNT(DISTINCT PATIENT.EVENT.UNIT_NUMBER)  AS TotalNumberOfPatients " & vbCrLf
   
   Const strFROM     As String = "FROM DRUG_COST_DATA dcd,  MEDICATION_ORDERS mo, ORDERS.PATIENT_ORDER, PATIENT.EVENT " & vbCrLf
   Const strWHERE    As String = "WHERE dcd.NDC = mo.DRUG_ALTERNATE_IDENTIFIER " & _
                                    "AND mo.ORDER_NUMBER = ORDERS.PATIENT_ORDER.PLACER_NUMBER " & _
                                    "AND ORDERS.PATIENT_ORDER.EVENT_ID = Patient.EVENT.EVENT_ID " & vbCrLf
   Const strGROUPBY  As String = "GROUP BY  dcd.ndc, dcd.GENERIC_NAME, dcd.GENSEQNUM " & vbCrLf

   Const strORDERBY1 As String = "ORDER BY Sum(mo.QUANTITY)  "
   
   Dim strSubQuery            As String
   Dim strDateInfo            As String
   Dim strFilter              As String
   Dim strORDERBY2            As String
   
   If blnAscendingSort Then
      strORDERBY2 = "ORDER BY " & strSortColumn & " ASC "
   Else
      strORDERBY2 = "ORDER BY " & strSortColumn & " DESC "
   End If
   
   TopDrugSQL = vbNullString

   'Filter on From Date
   strFilter = strFilter & " AND  mo.START_DATE >= " & sStartDate & " "
   
   'Filter on To Date
   strFilter = strFilter & " AND  mo.START_DATE <= " & sEndDate & " "
   
   'Filter on Facility NCID
   strFilter = strFilter & "AND mo.Facility_NCID= '" & strFacilityNCID & "' "
   '********************************************************************************************************
   
   'Concatenate sub query
   strSubQuery = strSELECT & strFROM & strWHERE & strFilter & strGROUPBY & strORDERBY1
   
   TopDrugSQL = "SELECT * FROM (" & strSubQuery & ") SubQuery1 " & _
                         "WHERE RowNum < " & CStr(lngNumberOfRows + 1) & " " & strORDERBY2
Proc_Exit:
   If Err.Number <> 0 Then
      TopDrugSQL = vbNullString
      Err.Raise Err.Number, Err.Source, Err.Description
   End If
   Exit Function
End Function

Public Function HighlightMTFRecords(ByRef objRecordset As Recordset, _
                                    ByVal strDrugField As String, _
                                    ByVal strClinicNCID As String, _
                                    ByVal dteFromDate As Date, _
                                    ByVal dteToDate As Date, _
                                    ByVal lngNumberOfRows As Long) As Boolean
   On Error GoTo Proc_Exit
   
   Dim strMTFNCID As String
   
   goDBSqlOps.Execute "SELECT Facility_NCID FROM CDRPLUS.Clinic WHERE NCID = '" & strClinicNCID & "'"
                           
   If Not (goDBSqlOps.BOF And goDBSqlOps.EOF) Then
        strMTFNCID = goDBSqlOps.Value("Facility_NCID")
   Else
      HighlightMTFRecords = False
      GoTo Proc_Exit
   End If
   
   goDBSqlOps.Execute TopDrugSQL(strMTFNCID, dteFromDate, dteToDate, lngNumberOfRows, "AverageDispense", False)
    
   If Not (objRecordset.BOF And objRecordset.EOF) Then
      objRecordset.MoveFirst
   Else
      HighlightMTFRecords = False
      GoTo Proc_Exit
   End If
   
   Do Until objRecordset.EOF
      
      goDBSqlOps.MoveFirst
         
      Do Until goDBSqlOps.EOF
         If goDBSqlOps.Value("Drug_Name") = objRecordset.Fields(strDrugField) Then
            objRecordset.Fields(strDrugField) = "**" & objRecordset.Fields(strDrugField) & "**"
            Exit Do
         End If
         goDBSqlOps.MoveNext
      Loop
   
      objRecordset.MoveNext
   
   Loop
   
Proc_Exit:
   If Err.Number <> 0 Then
      HighlightMTFRecords = False
      Err.Raise Err.Number, Err.Source, Err.Description
   End If
   Exit Function
   
End Function
'***********************************************************************************************************
'  Function Name: BuildReport
'
'  Description: This method:
'                 1.) Calls the BuildReportSQL method to get the SQL for the DrugCostForMTF report
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
   Dim objReportRecordset     As ADODB.Recordset
   
   BuildReport = False 'Assume false

   strSQL = BuildReportSQL(objReportItem, colFormulas)  'Get SQL
       
   '-- Get Data   '----------------------------------------------------------------------------
   Set moCollRecSet = Nothing
   Set moCollRecSet = New Collection
   
   Set objReportRecordset = goDBDAS.OpenRecordset(strSQL)  'Create recordset
   moCollRecSet.Add objReportRecordset, "TotalVsitsMTF"

   Set objReport = Nothing   'In case the object has been set prior
   
   If Not CreateReport(strReportName & ".rpt", objReport, colFormulas, moCollRecSet) Then
      BuildReport = False
      GoTo Proc_Exit
   Else
      objReport.Database.SetDataSource objReportRecordset, 3, 1 'Set data source
      BuildReport = True
   End If
   
   
Proc_Exit:
   Set objReportRecordset = Nothing  'Destroy reference to object
   Set colFormulas = Nothing     'Destroy reference to object
   
   If Err.Number <> 0 Then
      gobjshared.ShowVBError Err.Number, Err.Description, "Repeat Visitors.BuildReport", _
                             "CHCSII_Reports", vbExclamation
      'Resume Next
   End If
   Exit Function
End Function













