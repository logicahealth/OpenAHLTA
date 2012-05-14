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
Attribute VB_Name = "ProviderDrugComparison"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit



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
   Dim strDateInfo            As String

   
   BuildReportSQL = vbNullString

      
   'SELECT clause
   strSQL = "SELECT DRUG_COST_DATA.GENERIC_NAME, TO_CHAR (Avg(DRUG_COST_DATA.UNIT_PRICE)) AS AverageUnitCost, " & _
               "TO_CHAR (Count(MEDICATION_ORDERS.Order_Number)) AS   TotalNumberOfPrescriptions, " & _
               "TO_CHAR ((MEDICATION_ORDERS.QUANTITY * DRUG_COST_DATA.UNIT_PRICE)) AS TotalCost, " & _
               "TO_CHAR (Sum(DRUG_COST_DATA.UP_Quantity / MEDICATION_ORDERS.QUANTITY)) AS AverageDispense "
           
               
   'FROM clause
   strSQL = strSQL & "FROM DRUG_COST_DATA,  MEDICATION_ORDERS,  CLINIC_PROVIDER "

   'WHERE Clause
   strSQL = strSQL & "WHERE DRUG_COST_DATA.NDC = MEDICATION_ORDERS.DRUG_ALTERNATE_IDENTIFIER " & _
                        "AND MEDICATION_ORDERS.ORDERING_PROVIDER_NCID = CLINIC_PROVIDER.CLINICIAN_NCID "

   If IsDate(objReportItem.DateFrom) Then
      strSQL = strSQL & " AND  MEDICATION_ORDERS.START_DATE >= " & SQLDate(objReportItem.DateFrom) & " "
      strDateInfo = "From " & Format(objReportItem.DateFrom, "mm/dd/yyyy") & "   "
   End If

   If IsDate(objReportItem.DateTo) Then
      strSQL = strSQL & " AND  MEDICATION_ORDERS.START_DATE <= " & SQLDate(objReportItem.DateTo) & " "
      strDateInfo = strDateInfo & "To " & Format(objReportItem.DateTo, "mm/dd/yyyy")
   End If
   
  
   'Filter on Provider or Clinic
    If objReportItem.Criteria1 = Provider Then
     '  strSQL = strSQL & "AND MEDICATION_ORDERS.ORDERING_PROVIDER_NCID= '" & objReportItem.CriteriaSel1.Item(1).NCID & "' "
       colFormulas.Add CrystalQuote(objReportItem.CriteriaSel1.Item(1).Name), "Header2"
   ElseIf objReportItem.Criteria1 = Clinic Then
      If objReportItem.CriteriaSel1.Count > 0 Then  'Check to see if there is clinic has any items in it
          strSQL = strSQL & "AND MEDICATION_ORDERS.Entering_Location_NCID = '" & objReportItem.CriteriaSel2(1).NCID & "' "
          colFormulas.Add CrystalQuote("Clinic: " & objReportItem.CriteriaSel2.Item(1).Name), "Header2"
      Else
         'This filter was futhur narrowed by a clinic's provider
         'strSQL = strSQL & "AND MEDICATION_ORDERS.ORDERING_PROVIDER_NCID = '" & objReportItem.CriteriaSel2(1).NCID & "' "
         colFormulas.Add CrystalQuote("Provider: " & objReportItem.CriteriaSel2.Item(1).Name), "Header2"
      End If
   End If
   

   'Create Group by clause
   strSQL = strSQL & "GROUP BY CLINIC_PROVIDER.PROVIDER_NAME, MEDICATION_ORDERS.ORDERING_PROVIDER_NCID, " & _
                        "DRUG_COST_DATA.GENERIC_NAME, DRUG_COST_DATA.UNIT_PRICE, MEDICATION_ORDERS.QUANTITY, " & _
                        "MEDICATION_ORDERS.START_DATE, MEDICATION_ORDERS.END_DATE, MEDICATION_ORDERS.Facility_NCID "

   colFormulas.Add CrystalQuote(strDateInfo), "DateInfo"
   
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
'                 1.) Calls the BuildReportSQL method to get the SQL for the ProviderDrugComparison report
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
   Dim objCrystal             As Crystal

   
   BuildReport = False 'Assume false

   Set colFormulas = New Collection  'Create new collection to store formula information
   
   strSQL = BuildReportSQL(objReportItem, colFormulas)  'Get SQL
       
   Set objRecordset = goDBDAS.OpenRecordset(strSQL)  'Create recordset

   colFormulas.Add CrystalQuote("Provider Drug Comparison Report"), "Header1"  'Add main header formula
   
   Set objCrystal = New Crystal  'Create new crystal object
   
   Set objReport = Nothing   'In case the object has been set prior
   
   If Not objCrystal.ReportOpen(strReportName, objReport) Then   'Build the report
      BuildReport = False
      GoTo Proc_Exit
   ElseIf Not objCrystal.AssignFormulas(colFormulas, objReport) Then 'Assign formulas
      BuildReport = False
      GoTo Proc_Exit
   Else
      objReport.Database.SetDataSource objRecordset, 3, 1 'Set data source
      BuildReport = True
   End If
   
   
Proc_Exit:
   Set objRecordset = Nothing  'Destroy reference to object
   Set colFormulas = Nothing     'Destroy reference to object
   Set objCrystal = Nothing     'Destroy reference to object
   
   If Err.Number <> 0 Then
      gobjshared.ShowVBError Err.Number, Err.Description, "Repeat Visitors.BuildReport", _
                             "CHCSII_Reports", vbExclamation
   End If
   Exit Function
End Function















