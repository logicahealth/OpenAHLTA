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
Attribute VB_Name = "Top20DrugsByProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private moCollRecSet As Collection
Private moCollFormulas As Collection

Private Enum eTop20ReportType
   eVolume = 1
   eAverageCost = 2
   eTotalCost = 3
End Enum


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
   Dim objReportRecordset     As ADODB.Recordset
   Dim strReportHeader        As String
   Dim strMean                As String
   Dim strMeanInfo            As String
   Dim strSTDV                As String
   
    '-- Get Data    '----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection
   
   'Set colFormulas = New Collection  'Create new collection to store formula information
   strMeanInfo = vbNullString
   
   Select Case objReportItem.Report
      Case Top20DrugsByProviderGraph
         strSQL = GetProviderDrugSQL(objReportItem, colFormulas, eVolume, False, True) 'Get SQL
         GetMeanInfo objReportItem.CriteriaSel1(1).NCID, objReportItem.CriteriaSel2(1).StartDate, objReportItem.CriteriaSel2(1).EndDate, objReportItem.CriteriaSel2(1).NDC, strMean, strSTDV
         strMeanInfo = objReportItem.CriteriaSel2(1).Sequence_Position & " Of The Top 20: " & _
                           objReportItem.CriteriaSel2(1).Drug_Name & "   " & " Clinic Mean = " & _
                           strMean & " with a Standard Deviation of " & strSTDV
      Case ProviderDrugComparison
         strSQL = GetProviderAndPeerDrugSQL(objReportItem, colFormulas, eAverageCost, False)  'Get SQL
      Case Top20DrugsByVolumeTabular
         strSQL = GetProviderDrugSQL(objReportItem, colFormulas, eVolume, True, False) 'Get SQL
      Case Top20DrugsByCostAverageGraphAndTabular
         'TotalVsitsMTF
         strSQL = GetProviderDrugSQL(objReportItem, colFormulas, eAverageCost, True, False) 'Get SQL
   End Select
   
   BuildReport = False 'Assume false
       
   Set objReportRecordset = goDBDAS.OpenRecordset(strSQL)   'Create recordset
    moCollRecSet.Add objReportRecordset, "TotalVsitsMTF"
    
   Select Case objReportItem.Report
      Case Top20DrugsByVolumeTabular, Top20DrugsByCostAverageGraph
         HighlightTop50 objReportRecordset, objReportItem
      Case Else
         'Do Nothing
   End Select

   'Add main header formula
   'colFormulas.Add CrystalQuote(strReportHeader), "Header1"
   On Error Resume Next
   If Not Trim(strMeanInfo) = "" Then colFormulas.Add CrystalQuote(strMeanInfo), "MeanInfo"
   Err.Clear
   On Error GoTo Proc_Exit
    
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
   End If

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


'***********************************************************************************************************
'  Function Name: GetProviderAndPeerDrugSQL
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
Private Function GetProviderAndPeerDrugSQL(ByRef objReportItem As ReportItem, _
                                    ByRef colFormulas As Collection, _
                                    ByVal eReportType As eTop20ReportType, _
                                    ByVal blnTop20Only As Boolean) As String
                                
   On Error GoTo Proc_Exit
                        
   Const strFROM As String = "FROM DRUG_COST_DATA dcd,  MEDICATION_ORDERS mo " & vbCrLf
   Const strWHERE As String = "WHERE dcd.NDC = mo.DRUG_ALTERNATE_IDENTIFIER " & vbCrLf
   Const strGROUPBY  As String = "GROUP BY dcd.GENERIC_NAME "
   

   Dim strSQL                 As String
   Dim strDateInfo            As String
   Dim strProviderNCID        As String
   Dim strClinicNCID          As String
   Dim strDateFilter          As String
   Dim strProviderQuery       As String
   Dim strPeerQuery           As String
   Dim strPeerSubQuery        As String
   Dim strSELECT              As String
   Dim strSummaryField        As String
   Dim strORDERBY             As String
   Dim sStartDate As String
   Dim sEndDate As String
       
   GetProviderAndPeerDrugSQL = vbNullString 'initialize function value
    
    '-- Get Filter Information for SQL    '----------------------------------------------------------------------------
    sStartDate = GetOracleStartDate(GetFirstDayOfMonth(objReportItem.DateFrom))
    sEndDate = GetOracleEndDate(GetLastDayOfMonth(objReportItem.DateTo))
   
   strSELECT = "SELECT dcd.GENERIC_NAME AS Drug_Name "
   
   Select Case eReportType
   
      Case eVolume
         strSummaryField = "Volume"
         strSELECT = strSELECT & ", TO_CHAR(SUM(mo.QUANTITY)) AS " & strSummaryField & " "
         strORDERBY = "ORDER BY Average_Cost DESC " '"
         
      Case eAverageCost
         strSummaryField = "Average_Cost"
         strSELECT = strSELECT & ", TO_CHAR(Avg(mo.QUANTITY * dcd.UM_PRICE)) AS " _
                              & strSummaryField & vbCrLf
         strORDERBY = "ORDER BY Average_Cost DESC "
         
      Case eTotalCost
         strSummaryField = "Total_Cost"
         strSELECT = strSELECT & ", TO_CHAR(Sum(mo.QUANTITY * dcd.Unit_Price)) AS " & _
                              strSummaryField & " "
         strORDERBY = "ORDER BY Sum(mo.QUANTITY * dcd.UM_PRICE) DESC "
         
   End Select

   strProviderNCID = objReportItem.CriteriaSel2.Item(1).NCID  'Capture the NCID for the provider
   strClinicNCID = objReportItem.CriteriaSel1.Item(1).NCID  'Capture the provider's clinic ncid
   
   'Filter on date
   If IsDate(objReportItem.DateFrom) Then
      strDateFilter = "AND mo.START_DATE >= " & sStartDate & vbCrLf
      strDateInfo = "From " & Format(objReportItem.DateFrom, "mm/dd/yyyy") & "   "
   End If

   If IsDate(objReportItem.DateTo) Then
      strDateFilter = strDateFilter & " AND mo.START_DATE <= " & sEndDate & vbCrLf
      strDateInfo = strDateInfo & "Through " & Format(objReportItem.DateTo, "mm/dd/yyyy")
   End If
   
   'Add text to formula collection
   On Error Resume Next
   colFormulas.Remove "Header2"
   colFormulas.Add CrystalQuote("Provider: " & objReportItem.CriteriaSel2.Item(1).Name & _
                                 "   Clinic: " & objReportItem.CriteriaSel1.Item(1).Name), "Header2"
   colFormulas.Remove "DateInfo"
   colFormulas.Add CrystalQuote(strDateInfo), "DateInfo"
   Err.Clear
   On Error GoTo Proc_Exit
   

   'This is the SQL to get the volume of drugs used for the provider
   strProviderQuery = strSELECT & strFROM & strWHERE & _
                        "AND mo.Ordering_Provider_NCID = '" & strProviderNCID & "' " & vbCrLf & _
                           "AND mo.Entering_Location_NCID = '" & strClinicNCID & "' " & vbCrLf & strDateFilter & _
                           strGROUPBY & vbCrLf & strORDERBY
                      
                  
   'This is the SQL to get the volume of drugs used for the  provider's peer in the clinic
    strPeerSubQuery = strSELECT & strFROM & strWHERE & _
                        "AND mo.Ordering_Provider_NCID <> '" & strProviderNCID & "' " & vbCrLf & _
                        "AND mo.Entering_Location_NCID = '" & strClinicNCID & "' " & vbCrLf & strDateFilter & _
                        strGROUPBY & vbCrLf & strORDERBY
                  
                      
                      
   'This is the SQL to get then MEAN, Standard Devision
   strPeerQuery = "SELECT Drug_Name, ROUND(AVG(" & strSummaryField & "), 1) AS MEAN," & _
                        "ROUND(2 * STDDEV(" & strSummaryField & "), 1) AS STD_DEV2 " & _
                   "FROM (" & strPeerSubQuery & ") PeerSubQuery " & vbCrLf & _
                   "GROUP BY Drug_Name "
                   


    'Get the Drug_Name From ProviderQuery, MEAN from PeerQuery, and STD_DEV2 from PeerQuery
    strSQL = "SELECT ProviderQuery.Drug_Name AS Drug_Name, " & _
                  "TO_CHAR(ProviderQuery." & strSummaryField & ") AS " & strSummaryField & ", " & _
                  "TO_CHAR(PeerQuery.MEAN) AS MEAN, TO_CHAR(PeerQuery.STD_DEV2) AS STD_DEV2 " & _
              "FROM (" & strProviderQuery & ") ProviderQuery, (" & strPeerQuery & ") PeerQuery " & vbCrLf & _
              "WHERE PeerQuery.Drug_Name (+) = ProviderQuery.Drug_Name"
              
   'Return only the top 20 rows if parameter is true
   If blnTop20Only Then
      strSQL = strSQL & " AND ROWNUM < 21"
   End If
   GetProviderAndPeerDrugSQL = strSQL
   
   
Proc_Exit:
   If Err.Number <> 0 Then
      GetProviderAndPeerDrugSQL = vbNullString
      Err.Raise Err.Number, Err.Source, Err.Description
   End If
   Exit Function
End Function



'***********************************************************************************************************
'  Function Name: GetProviderAndPeerDrugSQL
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
Private Function GetProviderDrugSQL(ByRef objReportItem As ReportItem, _
                                    ByRef colFormulas As Collection, _
                                    ByVal eReportType As eTop20ReportType, _
                                    ByVal blnTop20Only As Boolean, _
                                    ByVal blnGroupByProvider As Boolean) As String
                                
   On Error GoTo Proc_Exit


   Dim strSQL                 As String
   Dim strDateInfo            As String
   Dim strFilter              As String
   Dim strSELECT              As String
   Dim strFROM                As String
   Dim strWHERE               As String
   Dim strORDERBY             As String
   Dim strGROUPBY             As String
   Dim strHeader2 As String
   
   Dim sStartDate As String
   Dim sEndDate As String
    
    '-- Get Filter Information for SQL
    '----------------------------------------------------------------------------
    sStartDate = GetOracleStartDate(GetFirstDayOfMonth(objReportItem.DateFrom))
    sEndDate = GetOracleEndDate(GetLastDayOfMonth(objReportItem.DateTo))
   
   GetProviderDrugSQL = vbNullString  'initialize function value
   
   strSELECT = "SELECT dcd.GENERIC_NAME AS Drug_Name, " & _
                      "TO_CHAR(SUM(mo.QUANTITY)) AS Number_Of_Units, " & _
                      "dcd.Unit_Price as Cost_Per_Unit, " & _
                      "TO_CHAR(Sum(mo.QUANTITY * dcd.Unit_Price)) AS Total_Cost," & _
                      "TO_CHAR(Avg(mo.QUANTITY * dcd.Unit_Price)) AS Average_Cost," & _
                      "COUNT(mo.Order_Number) AS Volume " & vbCrLf
                                 
   strFROM = "FROM DRUG_COST_DATA dcd,  MEDICATION_ORDERS mo " & vbCrLf
   strWHERE = "WHERE dcd.NDC = mo.DRUG_ALTERNATE_IDENTIFIER " & vbCrLf
   
   
   If blnGroupByProvider Then
      strSELECT = strSELECT & ",  p.NAME As Provider_Name, " & _
                  "COUNT(mo.Order_Number) AS NumberOfPrescriptions "
      strFROM = strFROM & ", PROVIDER p "
      strWHERE = strWHERE & " AND mo.ORDERING_PROVIDER_NCID = + p.NCID (+) " & vbCrLf
      strWHERE = strWHERE & " AND NOT mo.ORDERING_PROVIDER_NCID IS NULL " & vbCrLf
      strGROUPBY = "GROUP BY p.NAME, dcd.GENERIC_NAME, Unit_Price "
   Else
      strGROUPBY = "GROUP BY dcd.GENERIC_NAME, Unit_Price "
   End If
      
   Select Case eReportType
      Case eVolume
         strORDERBY = "ORDER BY Sum(mo.QUANTITY) DESC"
      Case eTotalCost
         strORDERBY = "ORDER BY Sum(mo.QUANTITY * dcd.Unit_Price) DESC"
      Case eAverageCost
         strORDERBY = "ORDER BY Avg(mo.QUANTITY * dcd.Unit_Price) DESC"
   End Select
   
   strHeader2 = vbNullString
   'Filter on Provider or Clinic
    If objReportItem.Criteria1 = Provider Then
      strFilter = "AND mo.ORDERING_PROVIDER_NCID= '" & objReportItem.CriteriaSel1.Item(1).NCID & "' "
      strHeader2 = "Provider: " & objReportItem.CriteriaSel1.Item(1).Name
   ElseIf objReportItem.Criteria1 = Clinic Then
      strFilter = "AND mo.Entering_Location_NCID = '" & objReportItem.CriteriaSel1(1).NCID & "' "
      strHeader2 = "Clinic: " & objReportItem.CriteriaSel1.Item(1).Name
   End If
   
   'Filter on date
   If Not objReportItem.Criteria2 = Top20Drug Then
   
      If IsDate(objReportItem.DateFrom) Then
         strFilter = strFilter & "AND mo.START_DATE >= " & sStartDate & " "
         strDateInfo = "From " & Format(objReportItem.DateFrom, "mmm yyyy") & "   "
      End If
   
      If IsDate(objReportItem.DateTo) Then
         strFilter = strFilter & " AND mo.START_DATE <= " & sEndDate & " "
         strDateInfo = strDateInfo & "Through " & Format(objReportItem.DateTo, "mmm yyyy")
      End If
   
   Else
      'Filter on the Drug NDC
      strFilter = strFilter & " AND dcd.NDC = '" & objReportItem.CriteriaSel2(1).NDC & "' "
    
      'Filter on date
      If IsDate(objReportItem.CriteriaSel2(1).StartDate) Then
         strFilter = strFilter & " AND mo.START_DATE >= " & sStartDate & " "
         strDateInfo = "From " & Format(objReportItem.CriteriaSel2(1).StartDate, "mmm yyyy") & "   "
      End If
   
      If IsDate(objReportItem.CriteriaSel2(1).EndDate) Then
         strFilter = strFilter & " AND mo.START_DATE <= " & sEndDate & " "
         strDateInfo = strDateInfo & "Through " & Format(objReportItem.CriteriaSel2(1).EndDate, "mmm yyyy")
      End If
      
   End If
   
   'Do a select on the sub query to filter the results to 20 rows later if need be
   strSQL = "SELECT Drug_Name, TO_CHAR(Number_Of_Units) AS Number_Of_Units, TO_CHAR(Cost_Per_Unit) AS Cost_Per_Unit, " & _
            "TO_CHAR(Total_Cost) AS Total_Cost, TO_CHAR(Volume) AS Volume, TO_CHAR(Average_Cost) AS  Average_Cost "
            
   If blnGroupByProvider Then
      strSQL = strSQL & ", Provider_Name, TO_CHAR(NumberOfPrescriptions) AS NumberOfPrescriptions "
   End If
            
   strSQL = strSQL & " FROM (" & strSELECT & strFROM & strWHERE & strFilter & strGROUPBY & strORDERBY & ") SubQuery1 "

   
   'Return only the top 20 rows if the "blnTop20Only" parameter is true
   If blnTop20Only Then
      strSQL = strSQL & " WHERE ROWNUM < 21"
   End If
   
   GetProviderDrugSQL = strSQL
   
   On Error Resume Next
   If Not Trim(strHeader2) = "" Then colFormulas.Add CrystalQuote(strHeader2), "Header2"
   colFormulas.Add CrystalQuote(strDateInfo), "DateInfo" 'Add date information to formula collection
   Err.Clear
   
Proc_Exit:
   If Err.Number <> 0 Then
      GetProviderDrugSQL = vbNullString
      Err.Raise Err.Number, Err.Source, Err.Description
   End If
   Exit Function
End Function


Private Function GetMeanInfo(ByVal strClinicNCID As String, _
                            ByVal dteFromDate As Date, _
                            ByVal dteToDate As Date, _
                            ByVal strDrugNDC As String, _
                            ByRef strMean As String, _
                            ByRef strSDV As String) As String
                                
   On Error GoTo Proc_Exit


   Dim strSQL                 As String
   Dim strWHERE               As String
   
   GetMeanInfo = vbNullString
                    
   strSQL = "SELECT AVG(COUNT(mo.Order_Number)) AS MEAN," & _
                        "ROUND(2 * STDDEV(COUNT(mo.Order_Number)), 1) AS STD_DEV2 " & _
                   "FROM MEDICATION_ORDERS mo "
                   
   strWHERE = "WHERE mo.Entering_Location_NCID = '" & strClinicNCID & "' " & _
                  "AND mo.START_DATE >= " & SQLDate(dteFromDate) & " " & _
                  "AND mo.START_DATE <= " & SQLDate(dteToDate) & " " & _
                  "AND mo.DRUG_ALTERNATE_IDENTIFIER = '" & strDrugNDC & "' "
   
   goDBSqlOps.Execute strSQL & strWHERE & "GROUP BY mo.Ordering_Provider_NCID "

   If Not (goDBSqlOps.BOF And goDBSqlOps.EOF) Then
      strMean = CStr(goDBSqlOps.Value("MEAN"))
      strSDV = CStr(goDBSqlOps.Value("STD_DEV2"))
   Else
      strMean = "0"
      strSDV = "0"
   End If
   
   If strMean = vbNullString Then
      strMean = "0"
   End If
   
   If strSDV = vbNullString Then
      strSDV = "0"
   End If
   
Proc_Exit:
   If Err.Number <> 0 Then
      GetMeanInfo = vbNullString
      Err.Raise Err.Number, Err.Source, Err.Description
   End If
   Exit Function
End Function

