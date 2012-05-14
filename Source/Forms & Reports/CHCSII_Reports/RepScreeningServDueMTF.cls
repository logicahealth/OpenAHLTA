VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RepScreeningServDueMTF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private moCollRecSet As Collection
Private moCollFormulas As Collection

Private msMachineName As String

Public Function BuildReport(sReportName As String, _
                            oReportItem As ReportItem, _
                            oReport As CRAXDRT.Report, _
                            oCollFormulas As Collection) As Boolean
 
    On Error GoTo ErrHandler
     
    BuildReport = False
    
    Dim oCrystal As Crystal
    Dim oRec As ADODB.Recordset
    Dim sSponsorUnitNumber As String
    Dim sUnitNumber As String
    DoEvents
    
    '-- Get Data
    '----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection

    '-- Get Candidates
    '----------------------------------------------------------
    If Not GetData(oReportItem, oRec, oCollFormulas) Then GoTo CleanUp
    moCollRecSet.Add oRec, "Reminders"

'-- Create Report Object
'-----------------------------------------------------------------------------
    If oReportItem.Criteria2 = Wellness Then
       oCollFormulas.Add CrystalQuote(oReportItem.CriteriaSel2(1).Name), "ServiceTitle"
       If Not CreateReport(sReportName & "2.rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo CleanUp
    Else
       If Not CreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo CleanUp
    End If
    BuildReport = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepScreeningServDue.BuildReport", "CHCSII_Reports", vbExclamation
CleanUp:
    Set oCrystal = Nothing
    Set oRec = Nothing
    
End Function

Private Function GetData(oReportItem As ReportItem, _
                           oRec As ADODB.Recordset, _
                           oCollFormulas As Collection) As Boolean

    Dim sSql As String
    Dim sTemp As String
    Dim sServiceFilter As String
    Dim sMTFFilter As String
    Dim sSortFilter As String
    Dim dCompareDate As Date
    
    Dim sHeader As String
    Dim iResponse As Integer
    
    Dim oIntermediate As New ADODB.Recordset
    Dim oReturnSet As New ADODB.Recordset
    
    On Error GoTo ErrHandler
    GetData = False
        
    If Not GetSQLCriteria(oReportItem, 1, sMTFFilter) Then GoTo CleanUp
    sMTFFilter = Replace(sMTFFilter, "a.facility_ncid", "FACILITY_CID")
    
    Set oReturnSet.ActiveConnection = Nothing
    oReturnSet.CursorLocation = adUseClient
    oReturnSet.Fields.Append "PCM_Code", adChar, 30
    oReturnSet.Fields.Append "PCM_Name", adChar, 30
    oReturnSet.Fields.Append "event_subtype_ncid", adChar, 20
    oReturnSet.Fields.Append "Overdue_Count", adInteger
    oReturnSet.Fields.Append "Comingdue_Count", adInteger
        
    oReturnSet.Open
    msMachineName = GetLocalComputerName
    
    sSql = "DELETE FROM REMINDER_REPORT_DATA WHERE COMPUTER_NAME = '" & msMachineName & "'"
    goDBSqlOps.Execute sSql
    
    If Not Populate_Table(oReportItem) Then GoTo CleanUp
    
'    cwobjCWComm.setStatusText "Load Patient PCM Matrix..."
'    DoEvents
'    Load_MTF_PCM_Data oReportItem
'
'    cwobjCWComm.setStatusText "Finding Reminder(s) For Patients..."
'    DoEvents
'    Load_MTF_Event_Data
'
'    cwobjCWComm.setStatusText "Checking Reminder(s) Status..."
'    DoEvents
'    Load_MTF_Event_Status
'
'    cwobjCWComm.setStatusText "Checking Reminder(s) Dates..."
'    DoEvents
'    Load_MTF_Event_Dates
    
    cwobjCWComm.setStatusText "Totalling Reminder(s) Data..."
    DoEvents
    Load_MTF_Totals oReturnSet
        
    GetData = True
    GoTo CleanUp

ErrHandler:
    If Err = -2147220874 Then
       MsgBxARMd "The selected report has timed out, please re-run this report again.", vbOKOnly + vbInformation
       Err.Clear
    Else
       gobjshared.ShowVBError Err.Number, Err.Description, "RepScreeningServDue.GetData", "CHCSII_Reports", vbExclamation
    End If
    'Resume Next
CleanUp:
    Set oRec = oReturnSet
    Set oIntermediate = Nothing
    Set oReturnSet = Nothing
End Function

Function GetPCMs(oRec As ADODB.Recordset, oReportItem As ReportItem) As Boolean
    GetPCMs = False
    
    Dim sFilter As String
    Dim sSql As String
    Dim sTemp As String
    
    If oReportItem.Criteria1 = MTF Then
       sFilter = "p.FACILITY_NCID = " & cwobjLogonEnviron.FacilityNCID
    ElseIf oReportItem.Criteria1 = MTFDMIS Then
       sTemp = GetClinicNCIDsFromDMIS(oReportItem.CriteriaSel1)
       If sTemp = "" Then
          If oReportItem.CriteriaSel1.Count = 0 Then
             MsgBxARMd "No DMIS Ids were selected to report on. Please select a valid DMIS then re-run the report.", vbOKOnly + vbInformation, "Screening Services Due"
             Exit Function
          Else
             MsgBxARMd "Either no clinics are assigned to the DMIS id(s) selected or the assigned clinics did not have a MEPS code indicating an outpatient status. Please select a valid DMIS then re-run the report.", vbOKOnly + vbInformation, "Screening Services Due"
             Exit Function
          End If
       End If
       sFilter = " p.CLINIC_NCID IN (" & sTemp & ") "
    End If
    
    sSql = "SELECT /*+ RULE */ /*+ ORDERED */ p.PCM_CODE, p.NAME" & vbCrLf & _
           "FROM PROVIDER p" & vbCrLf & _
           "Where Exists (SELECT 'X'" & vbCrLf & _
               "FROM MMI_GENERIC_ID mg" & vbCrLf & _
               "Where mg.id_value = p.pcm_code AND" & vbCrLf & _
                    "mg.id_type_ncid = 14501724) AND" & vbCrLf & sFilter & vbCrLf & "ORDER BY NAME"
    
    Set oRec = goDBDAS.OpenRecordset(sSql)

    GetPCMs = True
    GoTo CleanUp

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "RepScreeningServDueMTF", "CHCSII_Reports", vbExclamation)
CleanUp:
    
End Function
Private Function Populate_Table(oReportItem As ReportItem) As Boolean

    On Error GoTo ErrHandler
    Populate_Table = False
    
    Dim oPCMsRec As ADODB.Recordset
    Dim oRec As ADODB.Record
    Dim sSql As String
    Dim sServiceList As String
    Dim i As Integer
    
    cwobjCWComm.setStatusText "Pulling PCM information ..."
    DoEvents
    
    If Not GetPCMs(oPCMsRec, oReportItem) Then GoTo CleanUp
    
    sServiceList = CStr(oReportItem.CriteriaSel2(1).NCID)

    If oPCMsRec.RecordCount > 0 Then
        oPCMsRec.MoveFirst
        
        Do
            
            DoEvents
            DoEvents
            cwobjCWComm.setStatusText "Load Patient PCM Matrix for " & oPCMsRec("NAME").Value & " ..."
            DoEvents
    
           ' sSQL = "INSERT INTO REMINDER_REPORT_DATA" & vbCrLf & _
                        "(COMPUTER_NAME, PROCESS_STATUS, UNIT_NUMBER, PCM_CODE," & vbCrLf & _
                        "PCM_NAME, ITEM_COUNT, DUE_DATE)" & vbCrLf & _
                   "SELECT '" & msMachineName & "',4, e.unit_number," & vbCrLf & _
                          "'" & oPCMsRec("PCM_CODE") & "'," & SQLQuote(oPCMsRec("NAME")) & "," & vbCrLf & _
                          "e.event_subtype_ncid, ov.VALUE_DATETIME" & vbCrLf & _
                   "FROM" & vbCrLf & _
                        "EVENT E, CLINICAL_EVENT_HEADER C, CLINICAL_OBSERVATION CO, OBSERVATION_VALUE OV" & vbCrLf & _
                   "Where" & vbCrLf & _
                        "OV.OBSERVATION_SEQ = CO.OBSERVATION_SEQ AND" & vbCrLf & _
                        "OV.EVENT_ID = CO.EVENT_ID AND" & vbCrLf & _
                        "CO.EVENT_ID = E.EVENT_ID AND" & vbCrLf & _
                        "C.PRODUCTION_STATUS_NCID = 101984 AND" & vbCrLf & _
                        "C.STATUS_NCID = 1024 AND" & vbCrLf & _
                        "C.EVENT_ID = E.EVENT_ID AND" & vbCrLf & _
                        "E.EVENT_SUBTYPE_NCID IN (" & sServiceList & ") AND" & vbCrLf & _
                        "E.EVENT_TYPE_NCID = 162156 AND" & vbCrLf & _
                        "E.UNIT_NUMBER IN" & vbCrLf & _
                            "(SELECT mg.UNIT_NUMBER" & vbCrLf & _
                            "From" & vbCrLf & _
                                "MMI_GENERIC_ID mg" & vbCrLf & _
                            "Where" & vbCrLf & _
                                "mg.id_value = '" & oPCMsRec("PCM_CODE") & "' AND " & vbCrLf & _
                                "mg.id_type_ncid = 14501724)"
                                
            DoEvents
            cwobjCWComm.setStatusText "Load Patient PCM Matrix..."
            DoEvents
            Load_MTF_PCM_Data oPCMsRec("PCM_CODE"), sServiceList
        
            cwobjCWComm.setStatusText "Finding Reminder(s) For Patients..."
            DoEvents
            Load_MTF_Event_Data
        
            cwobjCWComm.setStatusText "Checking Reminder(s) Status..."
            DoEvents
            Load_MTF_Event_Status
        
            cwobjCWComm.setStatusText "Checking Reminder(s) Dates..."
            DoEvents
            Load_MTF_Event_Dates
            'goDBSqlOps.Execute sSQL
            DoEvents
            oPCMsRec.MoveNext
            If oPCMsRec.EOF Then
                Exit Do
            End If
        Loop

End If

Populate_Table = True
GoTo CleanUp

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "Populate_Table", "CHCSII_Reports", vbExclamation)
CleanUp:
    
End Function
Private Sub Load_MTF_Totals(oMTFTotals As ADODB.Recordset)

    Dim sSql As String
    Dim sStartDate As String
    Dim sEndDate As String
    Dim dReminderDate As Date
    
    On Error GoTo ErrHandler
    
    sSql = "SELECT  count(*) as PCM_Count, due_date, pcm_code, pcm_name" & vbCrLf
    sSql = sSql & "FROM reminder_report_data r" & vbCrLf
    sSql = sSql & "Where "
    sSql = sSql & "r.COMPUTER_NAME = " & SQLQuote(msMachineName) & " AND" & vbCrLf
    sSql = sSql & "r.PROCESS_STATUS = 4" & vbCrLf
    sSql = sSql & "GROUP BY PCM_CODE, PCM_NAME, DUE_DATE" & vbCrLf
    sSql = sSql & "ORDER BY PCM_NAME"
    
    goDBSqlOps.Execute sSql
    Do Until goDBSqlOps.EOF
         oMTFTotals.Filter = "PCM_Code = " & SQLQuote(goDBSqlOps("PCM_CODE"))
         If oMTFTotals.BOF And oMTFTotals.EOF Then
            oMTFTotals.AddNew
            oMTFTotals("PCM_CODE") = goDBSqlOps("PCM_CODE")
            oMTFTotals("PCM_NAME") = goDBSqlOps("PCM_NAME")
         End If
        
        dReminderDate = goDBSqlOps("DUE_DATE")
        If dReminderDate >= Format(Now, "mm/dd/yyyy") Then
            oMTFTotals.Fields("Comingdue_Count").Value = oMTFTotals.Fields("Comingdue_Count").Value + goDBSqlOps("PCM_Count")
        Else
            oMTFTotals.Fields("Overdue_Count").Value = oMTFTotals.Fields("Overdue_Count").Value + goDBSqlOps("PCM_Count")
        End If
        oMTFTotals.Filter = ""
        goDBSqlOps.MoveNext
    Loop
    
    GoTo CleanUpAndExit
    oMTFTotals.Sort = "PCM_NAME ASC"
ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ScreeningServices.Load_MTFData", "Load_MTFData", vbExclamation)
    'Resume Next
CleanUpAndExit:
    Screen.MousePointer = vbDefault
    'Resume Next
    'Err.Clear
End Sub



Private Function Load_MTF_Event_Dates() As Boolean
        
    Dim sSql As String
        
    sSql = "INSERT INTO REMINDER_REPORT_DATA" & vbCrLf
    sSql = sSql & "(COMPUTER_NAME, PROCESS_STATUS, UNIT_NUMBER, PCM_CODE, PCM_NAME, ITEM_COUNT, EVENT_ID, DUE_DATE)" & vbCrLf
    sSql = sSql & "SELECT /*+ RULE */ /*+ ORDERED */ /*+ INDEX (r REMINDER_REPORT_DATA_X1) (co CLINICAL_OBSERVATION_PK) (ov OBSERVATION_VALUE_X3)  */ r.COMPUTER_NAME, 4, r.UNIT_NUMBER, r.PCM_CODE, r.PCM_NAME, r.ITEM_COUNT, r.EVENT_ID, ov.VALUE_DATETIME " & vbCrLf
    sSql = sSql & "From reminder_report_data r, clinical_observation co, observation_value ov" & vbCrLf
    sSql = sSql & "WHERE ov.OBSERVATION_SEQ = co.OBSERVATION_SEQ and" & vbCrLf & _
                "ov.event_id = co.event_id and" & vbCrLf & _
                "co.observation_ncid = 162168 and" & vbCrLf & _
                "co.event_id = r.event_id and" & vbCrLf & _
                "r.PROCESS_STATUS = 3 AND " & _
                "r.COMPUTER_NAME = " & SQLQuote(msMachineName)

'    sSQL = "INSERT INTO REMINDER_REPORT_DATA" & vbCrLf
'    sSQL = sSQL & "(COMPUTER_NAME, PROCESS_STATUS, UNIT_NUMBER, PCM_CODE, PCM_NAME, ITEM_COUNT, EVENT_ID, DUE_DATE)" & vbCrLf
'    sSQL = sSQL & "SELECT /*+ RULE */ /*+ ORDERED */ /*+ INDEX (r REMINDER_REPORT_DATA_X1) (co CLINICAL_OBSERVATION_PK) (ov OBSERVATION_VALUE_X3)  */ r.COMPUTER_NAME, 4, r.UNIT_NUMBER, r.PCM_CODE, r.PCM_NAME, r.ITEM_COUNT, r.EVENT_ID, ov.VALUE_DATETIME " & vbCrLf
'    sSQL = sSQL & "From reminder_report_data r, clinical_observation co, observation_value ov" & vbCrLf
'    sSQL = sSQL & "WHERE ov.OBSERVATION_SEQ (+) = co.OBSERVATION_SEQ and" & vbCrLf & _
'                "ov.event_id (+) = co.event_id and" & vbCrLf & _
'                "co.observation_ncid = 162168 and" & vbCrLf & _
'                "co.event_id (+) = r.event_id and" & vbCrLf & _
'                "r.PROCESS_STATUS = 3 AND " & _
'                "r.COMPUTER_NAME = " & SQLQuote(msMachineName)
                
    goDBSqlOps.Execute sSql
    
    sSql = "DELETE FROM REMINDER_REPORT_DATA " & vbCrLf
    sSql = sSql & "WHERE COMPUTER_NAME = " & SQLQuote(msMachineName) & vbCrLf
    sSql = sSql & "AND PROCESS_STATUS = 3"
    goDBSqlOps.Execute sSql
    
End Function


Private Function Load_MTF_Event_Status() As Boolean
        
    Dim sSql As String
    
    sSql = "UPDATE /*+ RULE */ /*+ ORDERED */ /*+ INDEX (R REMINDER_REPORT_DATA_X1) (c event_pk) */ REMINDER_REPORT_DATA r" & vbCrLf & _
           "SET PROCESS_STATUS = 3" & vbCrLf & _
           "WHERE EXISTS (" & _
               "SELECT 'X' " & vbCrLf & _
               "FROM clinical_event_header c " & vbCrLf & _
               "WHERE " & vbCrLf & _
               "c.PRODUCTION_STATUS_NCID = 101984 and" & vbCrLf & _
               "c.status_ncid = 1024 and" & vbCrLf & _
               "c.event_id = r.event_id) and " & vbCrLf & _
               "r.PROCESS_STATUS = 2 AND " & vbCrLf & _
               "r.COMPUTER_NAME = " & SQLQuote(msMachineName)

    goDBSqlOps.Execute sSql
    
    sSql = "DELETE FROM REMINDER_REPORT_DATA " & vbCrLf
    sSql = sSql & "WHERE COMPUTER_NAME = " & SQLQuote(msMachineName) & vbCrLf
    sSql = sSql & "AND PROCESS_STATUS = 2"
    goDBSqlOps.Execute sSql
    
End Function


Private Function Load_MTF_Event_Data() As Boolean
        
    Dim sSql As String
    
    sSql = "INSERT /*+ RULE */ /*+ ORDERED */ /*+ INDEX (r REMINDER_REPORT_DATA_X1) (e EVENT_X1) */ INTO REMINDER_REPORT_DATA" & vbCrLf
    sSql = sSql & "(COMPUTER_NAME, PROCESS_STATUS, UNIT_NUMBER, PCM_CODE, PCM_NAME, ITEM_COUNT, EVENT_ID)" & vbCrLf
    sSql = sSql & "SELECT r.COMPUTER_NAME, 2, r.UNIT_NUMBER, r.PCM_CODE, r.PCM_NAME, r.ITEM_COUNT, e.event_id " & vbCrLf
    sSql = sSql & "FROM EVENT e, REMINDER_REPORT_DATA r " & vbCrLf
    sSql = sSql & "WHERE e.unit_number = r.unit_number and " & vbCrLf & _
                "e.EVENT_SUBTYPE_NCID = r.ITEM_COUNT and (e.event_type_ncid = 162156 OR e.event_type_ncid = 162158) AND " & vbCrLf & _
                "r.PROCESS_STATUS = 1 AND " & _
                "r.COMPUTER_NAME = " & SQLQuote(msMachineName)

    goDBSqlOps.Execute sSql
    
    sSql = "DELETE FROM REMINDER_REPORT_DATA " & vbCrLf
    sSql = sSql & "WHERE COMPUTER_NAME = " & SQLQuote(msMachineName) & vbCrLf
    sSql = sSql & "AND PROCESS_STATUS = 1"
    goDBSqlOps.Execute sSql

    
End Function


Private Function Load_MTF_PCM_Data(pPCMCode As String, pItem As String) As Boolean
        
    Dim sSql As String
    Dim sAdditionalReminder  As String
    
    sSql = "INSERT /*+ RULE */ /*+ ORDERED */ /*+ INDEX (P PROVIDER_X2) (mg MMI_GENERIC_ID_X1) */ INTO REMINDER_REPORT_DATA" & vbCrLf
    sSql = sSql & "(COMPUTER_NAME, PROCESS_STATUS, UNIT_NUMBER, PCM_CODE, PCM_NAME, ITEM_COUNT)" & vbCrLf
    sSql = sSql & "SELECT  " & SQLQuote(msMachineName) & ", 1, mg.UNIT_NUMBER, p.PCM_CODE, p.NAME, " & pItem & vbCrLf
    sSql = sSql & "FROM PROVIDER p , MMI_GENERIC_ID mg" & vbCrLf
    sSql = sSql & "WHERE mg.id_value = p.pcm_code and" & vbCrLf
    sSql = sSql & "mg.id_type_ncid= 14501724 and" & vbCrLf
    sSql = sSql & "p.pcm_code = " & SQLQuote(pPCMCode)
    goDBSqlOps.Execute sSql
    
    sAdditionalReminder = CheckForAltReadinessNCID(pItem)
    If Not sAdditionalReminder = "" Then
           sSql = "INSERT INTO REMINDER_REPORT_DATA " & vbCrLf
           sSql = sSql & "(COMPUTER_NAME, PROCESS_STATUS, UNIT_NUMBER, PCM_CODE, PCM_NAME, ITEM_COUNT)" & vbCrLf
           sSql = sSql & "SELECT /*+ ORDERED index(p provider_pk) */  " & SQLQuote(msMachineName) & ", 1, mg.UNIT_NUMBER, p.PCM_CODE, p.NAME, " & sAdditionalReminder & vbCrLf
           sSql = sSql & "FROM PROVIDER p , MMI_GENERIC_ID mg" & vbCrLf
           sSql = sSql & "WHERE mg.id_value = p.pcm_code and" & vbCrLf
           sSql = sSql & "mg.id_type_ncid= 14501724 and" & vbCrLf
           sSql = sSql & "p.pcm_code = " & SQLQuote(pPCMCode)
           goDBSqlOps.Execute sSql
    End If
    
End Function

Private Function CheckForAltReadinessNCID(pReminderNCID As String) As String
    
    Select Case pReminderNCID
        Case "14592298"
            CheckForAltReadinessNCID = "14592299"
        Case "14592296"
            CheckForAltReadinessNCID = "14592297"
        Case "14592294"
            CheckForAltReadinessNCID = "14592295"
        Case "14592292"
            CheckForAltReadinessNCID = "14592292"
        Case "14592290"
            CheckForAltReadinessNCID = "14592291"
        Case "14592288"
            CheckForAltReadinessNCID = "14592289"
        Case "14592286"
            CheckForAltReadinessNCID = "14592287"
        Case "14592284"
            CheckForAltReadinessNCID = "14592285"
        Case "14592282"
            CheckForAltReadinessNCID = "14592283"
        Case "14592280"
            CheckForAltReadinessNCID = "14592281"
        Case "14592278"
            CheckForAltReadinessNCID = "14592279"
        Case "14592276"
            CheckForAltReadinessNCID = "14592277"
        Case "14592274"
            CheckForAltReadinessNCID = "14592275"
        Case "14592272"
            CheckForAltReadinessNCID = "14592273"
        Case "14592270"
            CheckForAltReadinessNCID = "14592271"
        Case "14592268"
            CheckForAltReadinessNCID = "14592269"
        Case "14592266"
            CheckForAltReadinessNCID = "14592267"
        Case "14592264"
            CheckForAltReadinessNCID = "14592265"
        Case "14592262"
            CheckForAltReadinessNCID = "14592263"
        Case "14592260"
            CheckForAltReadinessNCID = "14592261"
        Case "14592258"
            CheckForAltReadinessNCID = "14592259"
        Case "14592256"
            CheckForAltReadinessNCID = "14592257"
        Case "14592254"
            CheckForAltReadinessNCID = "14592255"
        Case "14592252"
            CheckForAltReadinessNCID = "14592253"
        Case "14592250"
            CheckForAltReadinessNCID = "14592251"
        Case "14592300"
            CheckForAltReadinessNCID = "14510330"
        Case Else
            CheckForAltReadinessNCID = ""
    End Select

End Function
Private Function CheckForStdReadinessNCID(pReminderNCID As String) As String
    
    Select Case pReminderNCID
        Case "14592299"
            CheckForStdReadinessNCID = "14592298"
        Case "14592297"
            CheckForStdReadinessNCID = "14592296"
        Case "14592295"
            CheckForStdReadinessNCID = "14592294"
        Case "14592293"
            CheckForStdReadinessNCID = "14592292"
        Case "14592291"
            CheckForStdReadinessNCID = "14592290"
        Case "14592289"
            CheckForStdReadinessNCID = "14592288"
        Case "14592287"
            CheckForStdReadinessNCID = "14592286"
        Case "14592285"
            CheckForStdReadinessNCID = "14592284"
        Case "14592283"
            CheckForStdReadinessNCID = "14592282"
        Case "14592281"
            CheckForStdReadinessNCID = "14592280"
        Case "14592279"
            CheckForStdReadinessNCID = "14592278"
        Case "14592277"
            CheckForStdReadinessNCID = "14592276"
        Case "14592275"
            CheckForStdReadinessNCID = "14592274"
        Case "14592273"
            CheckForStdReadinessNCID = "14592272"
        Case "14592271"
            CheckForStdReadinessNCID = "14592270"
        Case "14592269"
            CheckForStdReadinessNCID = "14592268"
        Case "14592267"
            CheckForStdReadinessNCID = "14592266"
        Case "14592265"
            CheckForStdReadinessNCID = "14592264"
        Case "14592263"
            CheckForStdReadinessNCID = "14592262"
        Case "14592261"
            CheckForStdReadinessNCID = "14592260"
        Case "14592259"
            CheckForStdReadinessNCID = "14592258"
        Case "14592257"
            CheckForStdReadinessNCID = "14592256"
        Case "14592255"
            CheckForStdReadinessNCID = "14592254"
        Case "14592253"
            CheckForStdReadinessNCID = "14592252"
        Case "14592251"
            CheckForStdReadinessNCID = "14592250"
        Case "14510330"
            CheckForStdReadinessNCID = "14592300"
        Case Else
            CheckForStdReadinessNCID = pReminderNCID
    End Select

End Function

