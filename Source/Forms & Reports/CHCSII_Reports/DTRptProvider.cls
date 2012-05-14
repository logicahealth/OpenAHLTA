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
Attribute VB_Name = "DTRptProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Wilson24
'********************************************************************************************
'DTRptProvider.cls is for both Daily and Monthly
'History:
'Date:      Developer   Description
'---------  ----------- -------------------------------------------------------------------
'11/19/01   Wilson Sun  Added constant miNOTES_WORKLOAD for Notes template workload report
'05/02/01   Wilson Sun  Modified SqlProcess case 1,
'                       chage trunc(D1.ENCOUNTER_DATE) to trunc(A.APPT_DATE_TIME)
'04/10/01   Wilson Sun  Modified SqlProcess case 1, adding checking signed encounter code
'01/22/01   Wilson Sun  Created
'********************************************************************************************
Option Explicit

Private Const msDATE_FORMAT_VB   As String = "mm/dd/yyyy hh:mm:ss AM/PM"
Private Const msDATE_FORMAT_ORA  As String = "MM/DD/YYYY HH:MI:SS AM"
Private Const msNCID_PATIENT_CATEGORY As String = "14501721"
Private Const msNCID_PATIENT_FMP As String = "204708"
Private Const msNCID_PATIENT_CURRENT_NAME_USED As String = "187"
Private Const miNOTES_WORKLOAD As Integer = -4

Private Const UPDATE_SERVICE_DISP_STEP = 4 'KDUNNE 7/1/2002 - SCR 8069
Private Const UPDATE_FAMILY_MEMBER_CHILD_STEP = 5 'KDUNNE 7/1/2002 - SCR 8069
Private Const UPDATE_FAMILY_MEMBER_ADULT_STEP = 6 'KDUNNE 7/1/2002 - SCR 8069

Private mrsMain As ADODB.Recordset
Private mbLogSql As Boolean

'*********************************************************************************************
'Name:          GetProviderNCIDs
'Description:   Get ProviderNCIDs from search screen
'Parameters:    oReportItem -
'               sNCIDS      - ProviderNCIDs
'Return Value:  Ture - if sucessful, False - if faild
'Called in:     BuildReport()
'History
'Date       Developer       Description
'========   =============== ===============================================================
'07/09/01   Wilson Sun      Changed sNCIDs output value if no provider selected. SCR 14241
'02/12/01   Wilson Sun      Created
'********************************************************************************************
Private Function GetProviderNCIDs(ByVal oReportItem As ReportItem, _
                                  ByRef sNCIDS As String) As Boolean
    Dim i As Integer
    Dim icount As Integer
    
    On Error GoTo ErrHandler
    GetProviderNCIDs = False

    icount = oReportItem.CriteriaSel1.Count
    If icount = 0 Then
        sNCIDS = "(0)"
    Else
        If oReportItem.Criteria1 = ProviderMulti Then
            sNCIDS = "("
            For i = 1 To icount
                If i <> icount Then
                    sNCIDS = sNCIDS & oReportItem.CriteriaSel1.Item(i).NCID & ","
                Else
                    sNCIDS = sNCIDS & oReportItem.CriteriaSel1.Item(i).NCID & ") "
                End If
            Next
        End If
    End If
    GetProviderNCIDs = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "DTRptProvider.GetProviderNCIDs", "CHCSII_Reports", vbExclamation
CleanUp:
    Exit Function
End Function
'***********************************************************************
'Function name: BuildReport
'Description:   This is the main interface for calling to print SF603A
'Parameters:
'   pRepItem    - Report Info
'   prRs        - for final report data
'Return Value:  Ture, sucessful; False, otherwise
'History
'Date       Developer   Description
'--------   ----------  -----------------------------------------------------------------------
'01/22/02   Wilson Sun  Comment Step 1, since all workload go to TEMPL_WORKLOAD table.
'11/19/01   Wilson Sun  Added Step -4 (miNOTES_WORKLOAD) for NotePlus Template workload
'09/06/01   Wilson Sun  Added Step -3 for Exam/Evals General Tab, Oral Exam to the workload
'09/06/01   Wilson Sun  Added Step -2 for Exam/Evals General Tab, Tobacco Use to the workload
'09/05/01   Wilson Sun  Added Step 0 for Radiograph Code to the workload (SCR: 8517, 13105)
'04/18/01   Wilson Sun  Changed vDateFrom value to first day of month if report type is Monthly
'02/12/01   Wilson Sun  Created
'***********************************************************************
Public Function BuildReport(ByVal pRepItem As ReportItem, _
                            ByRef prRs As ADODB.Recordset) As Boolean
        
    Dim lRowCount As Long
    Dim lPeportItem As Long
    Dim sFacilityNCID As String
    Dim sModifiedBy As String
    Dim sInProviderNCIDs As String
    Dim sProcessTime As String
    
    Dim sSql As String
    Dim sLogSqlDesc As String
    Dim vDateFrom As Variant
    Dim vDateTo As Variant
    
    
    On Error GoTo ErrHandler

    BuildReport = False
    
    '--------------------------
    '-- Get Selection Criteria
    '--------------------------
    '--Provider NCIDS
    If Not GetProviderNCIDs(pRepItem, sInProviderNCIDs) Then GoTo CleanUp

    '--Login FacilityNCID
    sFacilityNCID = cwobjLogonEnviron.FacilityNCID
    
    
    '--Login UserName
    sModifiedBy = cwobjLogonEnviron.UserName
    
    '--TimeStamp *Important for application concurrency.
    '--convert parameter date value to oracle date format
    sProcessTime = goDBConn.FormatDateForDBInsert(Now) ' OracleDate(Now) 'this variable will also be used for delete condition

    '-- Report item
    lPeportItem = pRepItem.Report
                '*************************************
                'Test Only, Comment out when delivery
                '*************************************
'                If Not SkipProcess4Test(lPeportItem, prRs) Then GoTo CleanUp
'                BuildReport = True
'                Exit Function

    
    '**********************
    '--Start Process
    '**********************
    'KDunne 8/16/2002
    'SCR 26347, 8907 - We've removed the NotesPlus Switch, NotesPlus has been part of dental
    '                  for several months now.
    '--------------------------------------------------------
    'Step -4. Retrieve Provider Workload Data
    '--------------------------------------------------------
    If lPeportItem = DailyProvider Then
        vDateFrom = goDBConn.FormatDateForDBInsert(Format(pRepItem.DateFrom, "mm/dd/yyyy"))
        If Not SqlProcess(miNOTES_WORKLOAD, sModifiedBy, sProcessTime, sFacilityNCID, sInProviderNCIDs, _
                            vDateFrom) Then GoTo CleanUp
    Else
        vDateFrom = goDBConn.FormatDateForDBInsert(Format(pRepItem.DateFrom, "mm/01/yyyy"))
        vDateTo = goDBConn.FormatDateForDBInsert(Format(pRepItem.DateTo, "mm/dd/yyyy"))
        
        If Not SqlProcess(miNOTES_WORKLOAD, sModifiedBy, sProcessTime, sFacilityNCID, sInProviderNCIDs, _
                            vDateFrom, vDateTo) Then GoTo CleanUp
    End If
    'End SCRs 26347, 8907
    
    '--------------------------------------------------------
    'Step 2. Make sure data selected, if not stop process.
    '--------------------------------------------------------
    If Not SqlProcess(2, sModifiedBy, sProcessTime, prRs:=prRs) Then GoTo CleanUp
    If prRs.Fields(0).Value = 0 Then
        MsgBxARMd "No data selected for the report."
        GoTo CleanUp
    End If
    
    '----------------------------------------------------------
    'Step 3. Get Patient info."
    '----------------------------------------------------------
    If Not SqlProcess(3, sModifiedBy, sProcessTime) Then GoTo CleanUp
    
    'KDunne 7/1/2002
    'SCR 8906 - The steps to update the service_dispaly have been condensed
    '           into one step. We're now looking up the info from a lookup table.
    If Not SqlProcess(UPDATE_SERVICE_DISP_STEP, sModifiedBy, sProcessTime) Then GoTo CleanUp

    If Not SqlProcess(UPDATE_FAMILY_MEMBER_CHILD_STEP, sModifiedBy, sProcessTime) Then GoTo CleanUp

    If Not SqlProcess(UPDATE_FAMILY_MEMBER_ADULT_STEP, sModifiedBy, sProcessTime) Then GoTo CleanUp
'''    '--------------------------------------------------------
'''    'Step 4. Update Active Duty for display"
'''    '--------------------------------------------------------
'''    If Not SqlProcess(4, sModifiedBy, sProcessTime) Then GoTo CleanUp
'''
'''    '----------------------------------------------------------
'''    'Step 5. Update Other military for display"
'''    '----------------------------------------------------------
'''    If Not SqlProcess(5, sModifiedBy, sProcessTime) Then GoTo CleanUp
'''
'''    '----------------------------------------------------------
'''    'Step 6. Update Family Member Adult for display
'''    '----------------------------------------------------------
'''    If Not SqlProcess(6, sModifiedBy, sProcessTime) Then GoTo CleanUp
'''
'''    '----------------------------------------------------------
'''    'Step 7. Update Family Member Child for display
'''    '----------------------------------------------------------
'''    If Not SqlProcess(7, sModifiedBy, sProcessTime) Then GoTo CleanUp
'''
'''    '----------------------------------------------------------
'''    'Step 8. Update Retired for display
'''    '----------------------------------------------------------
'''    If Not SqlProcess(8, sModifiedBy, sProcessTime) Then GoTo CleanUp
'''
'''    '----------------------------------------------------------
'''    'Step 9. Update Other for display
'''    '----------------------------------------------------------
'''    If Not SqlProcess(9, sModifiedBy, sProcessTime) Then GoTo CleanUp
    
    '**********************************
    '--Ready to print Daily Report now
    '**********************************
    If lPeportItem = DailyProvider Then
        '----------------------------------------------------------
        'Step 10. <<<Final DAILY REPORT data>>>
        '----------------------------------------------------------
        If Not SqlProcess(10, sModifiedBy, sProcessTime, prRs:=prRs) Then GoTo CleanUp
        If prRs.EOF And prRs.BOF Then
            MsgBxARMd "No data selected for the daily workload report."
            GoTo CleanUp
        End If
    
    Else
        '************************************
        'Continue process for Monthly Report
        '************************************
        '----------------------------------------------------------
        'Step 11. Summation by Service
        '----------------------------------------------------------
        If Not SqlProcess(11, sModifiedBy, sProcessTime) Then GoTo CleanUp
        
        '----------------------------------------------------------
        'Step 12. <<<Final MONTHLY REPORT data>>>
        '----------------------------------------------------------
        If Not SqlProcess(12, sModifiedBy, sProcessTime, prRs:=prRs) Then GoTo CleanUp
        If prRs.EOF And prRs.BOF Then
            MsgBxARMd "No data selected for the monthly workload report."
            GoTo CleanUp
        End If
    
        '----------------------------------------------------------
        'Step 23. delete Z_MONTHLY_RPT table
        '----------------------------------------------------------
        If Not SqlProcess(23, sModifiedBy, sProcessTime) Then GoTo CleanUp
    End If
    
'    '*******************
'    'Delete Work tables
'    '*******************
'    '----------------------------------------------------------
'    'Step 21. delete Z_WORKLOAD table
'    '----------------------------------------------------------
'    If Not SqlProcess(21, sModifiedBy, sProcessTime) Then GoTo CleanUp
'
'    '----------------------------------------------------------
'    'Step 22. delete Z_DAILY_RPT table
'    '----------------------------------------------------------
'    If Not SqlProcess(22, sModifiedBy, sProcessTime) Then GoTo CleanUp

'    '******************************
'    '-- Process Completed
'    '******************************
'    '--LogSql
'    If mbLogSql Then
'        LogSql "Process Complete On:", Format(Now)
'    End If
    
    BuildReport = True

    GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "DTRptProvider.BuildReport", "CHCSII_Reports", vbExclamation
CleanUp:
    Set mrsMain = Nothing
    '*******************
    'Delete Work tables
    '*******************
    '----------------------------------------------------------
    'Step 21. delete Z_WORKLOAD table
    '----------------------------------------------------------
    Call SqlProcess(21, sModifiedBy, sProcessTime)
        
    '----------------------------------------------------------
    'Step 22. delete Z_DAILY_RPT table
    '----------------------------------------------------------
    Call SqlProcess(22, sModifiedBy, sProcessTime)
    
    '******************************
    '-- Process Completed
    '******************************
    '--LogSql
    If mbLogSql Then
        LogSql "Process Complete On:", Format(Now)
    End If

End Function
Private Sub Class_Initialize()
    Set mrsMain = New ADODB.Recordset
    If cwobjCWComm.CmdLineSwitch("LOGSQL") = "" Then
        mbLogSql = False
    Else
        mbLogSql = True
    End If
End Sub
Private Sub Class_Terminate()
    Set mrsMain = Nothing
    mbLogSql = False
End Sub

'**********************************************************************************
'Name:          SqlProcess
'Description:   This is a main container of Sql statments for the whole process
'Parameters:
'   piProcessStep   - Process Step
'   psModifiedBy    - Login User Name
'   psProcessTime   - When you click Report button time
'   plFacilityNCID  - Login FacilityNCID
'   psInProviders   - Value for Sql In statement
'   pvDateFrom      - Screen Date from value
'   pvDateTo        - Screen Date to value
'   prRs            - Record set for the report
'Return Value:
'   True - sucessful,
    'False - failed
'Called by:     BuildReport()
'History
'   Developer   Date        Description
'   =========   ==========  ==========================================================
'   Wilson Sun  01/22/02    Change count(*) to sum(MULTIPLIER) in Step -4.
'   Wilson Sun  11/19/01    Added Step -4 (miNOTES_WORKLOAD) for NotesPlus template
'   Wilson Sun  09/06/01    Added Step -3 for retrieveing Exam/Evals, General, Oral Exam Data
'   Wilson Sun  09/06/01    Added Step -2 for retrieveing Exam/Evals, General, Tobacco Use Data
'   Wilson Sun  09/05/01    Added Step 0 for retrieveing Radiograph code data.
'   Wilson Sun  08/08/01    1. Took  DENTAL and CDRPULS schema away from SQL statment.
'                           2. Modified DAS object declaretion to work in theater.
'   Wilson Sun  06/12/01    1. Modified Step 1 from D1.ENCOUNTER_DATE to A.APPT_DATE_TIME
'                           2. Add Step 1 for checking D2.STATUS_ID = 102 (Completed)
'   Wilson Sun  02/14/01    Created
'**********************************************************************************
Private Function SqlProcess(ByVal piProcessStep As Integer, _
                            ByVal psModifiedBy As String, _
                            ByVal psProcessTime As String, _
                            Optional ByVal plFacilityNCID As String, _
                            Optional ByVal psInProviders As String, _
                            Optional ByVal pvDateFrom As Variant, _
                            Optional ByVal pvDateTo As Variant, _
                            Optional ByRef prRs As ADODB.Recordset) As Boolean
    '--Boolean
    Dim bCreatLogFile As Boolean
    
    '--String
    Dim sAppendWhereClause As String
    Dim sFacilityName As String
    Dim sLogSqlDesc As String
    Dim sSql As String
        
    On Error GoTo ErrHandler

    SqlProcess = False
    
    bCreatLogFile = False
    
   '--Login FacilityName
    sFacilityName = cwobjLogonEnviron.FacilityName

    sAppendWhereClause = "Z.MODIFIED_BY = '" & RQ(psModifiedBy) & "'" & vbCrLf & _
       "and    Z.TIME_STAMP = " & psProcessTime
    
    Select Case piProcessStep
        Case miNOTES_WORKLOAD
            '--------------------------------------------------------
            'Step -4. Retrieve NotesPlus Template Data from
            '--CDRPLUS.APPOINTMENT,
            '--DENTAL.DENTAL_ENCOUNTER,
            '--DENTAL.TEMPL_WORKLOAD,
            '--DENTAL.TX_ADA_CODE_DEFINITION
            '** Clone from Step 1
            '** Since TEMPL_WORKLOAD did not contains PROVIDER_ID, I use A.CLINICIAN_NCID for now results slow retrieve data
            '** Scott Foster will create PROVIDER_ID in the futer, when change A.CLINICIAN_NCID to D2.PROVIDER_ID by then.
            '--------------------------------------------------------
            bCreatLogFile = True
            sLogSqlDesc = "Step -4. Retrieve Workload Data"

            '--Insert
            sSql = "insert into Z_WORKLOAD(FACILITY_NCID, CLINIC_NCID, PROVIDER_NCID, UNIT_NUMBER, " & vbCrLf & _
                    vbTab & "DEFINITION, DRC, ENCOUNTER_DATE, COUNT_NO, MODIFIED_BY, TIME_STAMP) "
            '--select
            sSql = sSql & vbCrLf & _
                    "select  A.FACILITY_NCID, A.CLINIC_NCID, A.CLINICIAN_NCID, D1.PATIENT_UNIT_ID, " & vbCrLf & _
                    vbTab & "D3.ADA_CODE ||' '|| D3.ADA_SHORT, D1.DRC, A.APPT_DATE_TIME, " & vbCrLf & _
                    vbTab & "sum(MULTIPLIER), '" & RQ(psModifiedBy) & "'," & psProcessTime
            '--from
            sSql = sSql & vbCrLf & _
                    "from APPOINTMENT A, " & vbCrLf & _
                    vbTab & " DENTAL_ENCOUNTER D1, " & vbCrLf & _
                    vbTab & " TEMPL_WORKLOAD D2, " & vbCrLf & _
                    vbTab & " TX_ADA_CODE_DEFINITION D3 "
            '--where
            'SCR 44425 RWalker 11/6/2003 Added Amended Status Updated (6) to Completed (4)
            '
            'PART I - SCRs 8905/8906 RWalker 12/4/2003
            'Only show workload credits for the current provider, screen others out
            sSql = sSql & vbCrLf & _
                    "where A.PATIENT_UNIT_NUMBER = D1.PATIENT_UNIT_ID " & vbCrLf & _
                    "and   A.FACILITY_NCID = D1.ENCOUNTER_FACILITY_NCID " & vbCrLf & _
                    "and   A.ENCOUNTER_NUMBER = D1.ENCOUNTER_NUMBER " & vbCrLf & _
                    "and   D1.DENTAL_ENCOUNTER_ID = D2.DENTAL_ENCOUNTER_ID " & vbCrLf & _
                    "and   D2.TX_ADA_CODE_ID = D3.TX_ADA_CODE_ID " & vbCrLf & _
                    "and   D1.ENCOUNTER_FACILITY_NCID = " & plFacilityNCID & vbCrLf & _
                    "and   A.CLINICIAN_NCID in " & psInProviders & vbCrLf & _
                    "and   D2.DELETED = 0 " & vbCrLf & _
                    "and   (D2.Provider_ID = A.Clinician_NCID or D2.Provider_ID = 0 or D2.Provider_ID IS NULL) " & vbCrLf & _
                    "and   A.ENCOUNTER_STATUS in (4,6) "
                    
            If IsMissing(pvDateTo) Then 'For Daily
                sSql = sSql & vbCrLf & _
                    "and " & goDBConn.FormatDateForDBSelect("A.Appt_Date_Time", DD_MON_YYYY, True) & " = " & pvDateFrom
            Else    'For Monthly
                sSql = sSql & vbCrLf & _
                    "and " & goDBConn.FormatDateForDBSelect("A.APPT_DATE_TIME", DD_MON_YYYY, True) & vbCrLf & _
                    vbTab & "between " & pvDateFrom & vbCrLf & _
                    vbTab & "and " & pvDateTo
            End If
            '--group by
            sSql = sSql & vbCrLf & _
                    "group by A.FACILITY_NCID, A.CLINIC_NCID, A.CLINICIAN_NCID, D1.PATIENT_UNIT_ID, " & vbCrLf & _
                    vbTab & "D3.ADA_CODE ||' '|| D3.ADA_SHORT, " & vbCrLf & _
                    vbTab & "D1.DRC, A.APPT_DATE_TIME "
        Case -3
            '--------------------------------------------------------
            'Step -3. Retrieve Exam/Evals, General Tab, Oral Exam Data
            '--CDRPLUS.APPOINTMENT,
            '--DENTAL.DENTAL_ENCOUNTER,
            '--DENTAL.NOTE_PERIODIC,
            '--DENTAL.TX_ADA_CODE_DEFINITION
            '--------------------------------------------------------
            bCreatLogFile = True
            sLogSqlDesc = "Step -3. Retrieve Exam/Evals, General Tab, Oral Exam Data"

            '--Insert
            sSql = "insert into Z_WORKLOAD(FACILITY_NCID, CLINIC_NCID, PROVIDER_NCID, UNIT_NUMBER, " & vbCrLf & _
                    vbTab & "DEFINITION, DRC, ENCOUNTER_DATE, COUNT_NO, MODIFIED_BY, TIME_STAMP) "
            '--select
            sSql = sSql & vbCrLf & _
                    "select  A.FACILITY_NCID, A.CLINIC_NCID, D2.PROVIDER_ID, D1.PATIENT_UNIT_ID, " & vbCrLf & _
                    vbTab & "D3.ADA_CODE ||' '|| D3.ADA_SHORT, D1.DRC, A.APPT_DATE_TIME, " & vbCrLf & _
                    vbTab & "count(*), '" & RQ(psModifiedBy) & "'," & psProcessTime
            '--from
            sSql = sSql & vbCrLf & _
                    "from APPOINTMENT A, " & vbCrLf & _
                    vbTab & " DENTAL_ENCOUNTER D1, " & vbCrLf & _
                    vbTab & " NOTE_PERIODIC D2, " & vbCrLf & _
                    vbTab & " TX_ADA_CODE_DEFINITION D3 "
            '--where
            sSql = sSql & vbCrLf & _
                    "where A.PATIENT_UNIT_NUMBER = D1.PATIENT_UNIT_ID " & vbCrLf & _
                    "and   A.FACILITY_NCID = D1.ENCOUNTER_FACILITY_NCID " & vbCrLf & _
                    "and   A.ENCOUNTER_NUMBER = D1.ENCOUNTER_NUMBER " & vbCrLf & _
                    "and   D1.DENTAL_ENCOUNTER_ID = D2.NOTE_PERIODIC_ID " & vbCrLf & _
                    "and   D2.TX_ADA_CODE_ID = D3.TX_ADA_CODE_ID " & vbCrLf & _
                    "and   D1.ENCOUNTER_FACILITY_NCID = " & plFacilityNCID & vbCrLf & _
                    "and   D2.PROVIDER_ID in " & psInProviders  ' & vbCrLf & _
                '   "and   A.ENCOUNTER_STATUS = 4 "

            If IsMissing(pvDateTo) Then 'For Daily
                sSql = sSql & vbCrLf & _
                    "and " & goDBConn.FormatDateForDBSelect("A.APPT_DATE_TIME", DD_MON_YYYY, True) = " & pvDateFrom"
'                sSql = sSql & vbCrLf & _
'                    "and   trunc(A.APPT_DATE_TIME) = " & pvDateFrom
            Else    'For Monthly
                sSql = sSql & vbCrLf & _
                    "and " & goDBConn.FormatDateForDBSelect("A.APPT_DATE_TIME", DD_MON_YYYY, True) & vbCrLf & _
                    vbTab & "between " & pvDateFrom & vbCrLf & _
                    vbTab & "and " & pvDateTo
'                sSql = sSql & vbCrLf & _
'                    "and   trunc(A.APPT_DATE_TIME) " & vbCrLf & _
'                    vbTab & "between " & pvDateFrom & vbCrLf & _
'                    vbTab & "and " & pvDateTo
            End If
            '--group by
            sSql = sSql & vbCrLf & _
                    "group by A.FACILITY_NCID, A.CLINIC_NCID, D2.PROVIDER_ID, D1.PATIENT_UNIT_ID, " & vbCrLf & _
                    vbTab & "D3.ADA_CODE ||' '|| D3.ADA_SHORT, " & vbCrLf & _
                    vbTab & "D1.DRC, A.APPT_DATE_TIME "
        
        Case -2
            '--------------------------------------------------------
            'Step -2. Retrieve Exam/Evals, General Tab, Tobacco Use Data
            '--CDRPLUS.APPOINTMENT,
            '--DENTAL.DENTAL_ENCOUNTER,
            '--DENTAL.NOTE_PERIODIC,
            '--DENTAL.TX_ADA_CODE_DEFINITION
            '--DENTAL.PREVENTIVE_COUNSELING
            '--------------------------------------------------------
            bCreatLogFile = False
            sLogSqlDesc = "Step -2. Retrieve Exam/Evals, General Tab, Tobacco Use Data"

            '--Insert
            sSql = "insert into Z_WORKLOAD(FACILITY_NCID, CLINIC_NCID, PROVIDER_NCID, UNIT_NUMBER, " & vbCrLf & _
                    vbTab & "DEFINITION, DRC, ENCOUNTER_DATE, COUNT_NO, MODIFIED_BY, TIME_STAMP) "
            '--select
            sSql = sSql & vbCrLf & _
                    "select  A.FACILITY_NCID, A.CLINIC_NCID, D2.PROVIDER_ID, D1.PATIENT_UNIT_ID, " & vbCrLf & _
                    vbTab & "D3.ADA_CODE ||' '|| D3.ADA_SHORT, D1.DRC, A.APPT_DATE_TIME, " & vbCrLf & _
                    vbTab & "count(*), '" & RQ(psModifiedBy) & "'," & psProcessTime
            '--from
            sSql = sSql & vbCrLf & _
                    "from APPOINTMENT A, " & vbCrLf & _
                    vbTab & " DENTAL_ENCOUNTER D1, " & vbCrLf & _
                    vbTab & " NOTE_PERIODIC D2, " & vbCrLf & _
                    vbTab & " TX_ADA_CODE_DEFINITION D3, " & vbCrLf & _
                    vbTab & " PREVENTIVE_COUNSELING D4 "
            '--where
            sSql = sSql & vbCrLf & _
                    "where A.PATIENT_UNIT_NUMBER = D1.PATIENT_UNIT_ID " & vbCrLf & _
                    "and   A.FACILITY_NCID = D1.ENCOUNTER_FACILITY_NCID " & vbCrLf & _
                    "and   A.ENCOUNTER_NUMBER = D1.ENCOUNTER_NUMBER " & vbCrLf & _
                    "and   D1.DENTAL_ENCOUNTER_ID = D2.NOTE_PERIODIC_ID " & vbCrLf & _
                    "and   D3.TX_ADA_CODE_ID = D4.TX_ADA_CODE_ID " & vbCrLf & _
                    "and   D2.NOTE_PERIODIC_ID = D4.NOTE_PREVENTIVE_ID " & vbCrLf & _
                    "and   D2.TOBACCO_COUNSELING = -1 " & vbCrLf & _
                    "and   D1.ENCOUNTER_FACILITY_NCID = " & plFacilityNCID & vbCrLf & _
                    "and   D2.PROVIDER_ID in " & psInProviders  ' & vbCrLf & _
                '   "and   A.ENCOUNTER_STATUS = 4 "

            If IsMissing(pvDateTo) Then 'For Daily
                sSql = sSql & vbCrLf & _
                    "and " & goDBConn.FormatDateForDBSelect("A.APPT_DATE_TIME", DD_MON_YYYY, True) & " = " & pvDateFrom
'                sSql = sSql & vbCrLf & _
'                    "and   trunc(A.APPT_DATE_TIME) = " & pvDateFrom
            Else    'For Monthly
                sSql = sSql & vbCrLf & _
                    "and " & goDBConn.FormatDateForDBSelect("A.APPT_DATE_TIME", DD_MON_YYYY, True) & vbCrLf & _
                    vbTab & "between " & pvDateFrom & vbCrLf & _
                    vbTab & "and " & pvDateTo
'                sSql = sSql & vbCrLf & _
'                    "and   trunc(A.APPT_DATE_TIME) " & vbCrLf & _
'                    vbTab & "between " & pvDateFrom & vbCrLf & _
'                    vbTab & "and " & pvDateTo
            End If
            '--group by
            sSql = sSql & vbCrLf & _
                    "group by A.FACILITY_NCID, A.CLINIC_NCID, D2.PROVIDER_ID, D1.PATIENT_UNIT_ID, " & vbCrLf & _
                    vbTab & "D3.ADA_CODE ||' '|| D3.ADA_SHORT, " & vbCrLf & _
                    vbTab & "D1.DRC, A.APPT_DATE_TIME "
        
        Case 0
            '--------------------------------------------------------
            'Step 0. Retrieve Dental Notes Radiograph Code Data
            '--CDRPLUS.APPOINTMENT,
            '--DENTAL.DENTAL_ENCOUNTER,
            '--DENTAL.NOTE_RADIOGRAPH,
            '--DENTAL.DENTAL_RADIOGRAPHS,
            '--DENTAL.TX_ADA_CODE_DEFINITION
            '--------------------------------------------------------
            bCreatLogFile = False
            sLogSqlDesc = "Step 0. Retrieve Dental Notes Radiograph Code Data"

            '--Insert
            sSql = "insert into Z_WORKLOAD(FACILITY_NCID, CLINIC_NCID, PROVIDER_NCID, UNIT_NUMBER, " & vbCrLf & _
                    vbTab & "DEFINITION, DRC, ENCOUNTER_DATE, COUNT_NO, MODIFIED_BY, TIME_STAMP) "
            '--select
            sSql = sSql & vbCrLf & _
                    "select  A.FACILITY_NCID, A.CLINIC_NCID, D2.PROVIDER_ID, D1.PATIENT_UNIT_ID, " & vbCrLf & _
                    vbTab & "D3.ADA_CODE ||' '|| D3.ADA_SHORT, D1.DRC, A.APPT_DATE_TIME, " & vbCrLf & _
                    vbTab & "count(*), '" & RQ(psModifiedBy) & "'," & psProcessTime
            '--from
            sSql = sSql & vbCrLf & _
                    "from APPOINTMENT A, " & vbCrLf & _
                    vbTab & " DENTAL_ENCOUNTER D1, " & vbCrLf & _
                    vbTab & " NOTE_RADIOGRAPH D2, " & vbCrLf & _
                    vbTab & " DENTAL_RADIOGRAPHS D4, " & vbCrLf & _
                    vbTab & " TX_ADA_CODE_DEFINITION D3 "
            '--where
            sSql = sSql & vbCrLf & _
                    "where A.PATIENT_UNIT_NUMBER = D1.PATIENT_UNIT_ID " & vbCrLf & _
                    "and   A.FACILITY_NCID = D1.ENCOUNTER_FACILITY_NCID " & vbCrLf & _
                    "and   A.ENCOUNTER_NUMBER = D1.ENCOUNTER_NUMBER " & vbCrLf & _
                    "and   D1.DENTAL_ENCOUNTER_ID = D2.DENTAL_ENCOUNTER_ID " & vbCrLf & _
                    "and   D2.DENTAL_ENCOUNTER_ID = D4.NOTE_RADIOGRAPH_ID " & vbCrLf & _
                    "and   D3.TX_ADA_CODE_ID = D4.TX_ADA_CODE_ID " & vbCrLf & _
                    "and   D1.ENCOUNTER_FACILITY_NCID = " & plFacilityNCID & vbCrLf & _
                    "and   D2.PROVIDER_ID in " & psInProviders & vbCrLf & _
                    "and   A.ENCOUNTER_STATUS = 4 "

            If IsMissing(pvDateTo) Then 'For Daily
                sSql = sSql & vbCrLf & _
                    "and " & goDBConn.FormatDateForDBSelect("A.APPT_DATE_TIME", DD_MON_YYYY, True) & " = " & pvDateFrom
'                    "and   trunc(A.APPT_DATE_TIME) = " & pvDateFrom
            Else    'For Monthly
                sSql = sSql & vbCrLf & _
                    "and " & goDBConn.FormatDateForDBSelect("A.APPT_DATE_TIME", DD_MON_YYYY, True) & vbCrLf & _
                    vbTab & "between " & pvDateFrom & vbCrLf & _
                    vbTab & "and " & pvDateTo
'                    "and   trunc(A.APPT_DATE_TIME) " & vbCrLf & _
'                    vbTab & "between " & pvDateFrom & vbCrLf & _
'                    vbTab & "and " & pvDateTo
            End If
            '--group by
            sSql = sSql & vbCrLf & _
                    "group by A.FACILITY_NCID, A.CLINIC_NCID, D2.PROVIDER_ID, D1.PATIENT_UNIT_ID, " & vbCrLf & _
                    vbTab & "D3.ADA_CODE ||' '|| D3.ADA_SHORT, " & vbCrLf & _
                    vbTab & "D1.DRC, A.APPT_DATE_TIME "
        
        Case 1
            '--------------------------------------------------------
            'Step 1. Retrieve Dental Data from
            '--CDRPLUS.APPOINTMENT,
            '--DENTAL.DENTAL_ENCOUNTER,
            '--DENTAL.DENTAL_ENCOUNTER_TX,
            '--DENTAL.TX_ADA_CODE_DEFINITION
            '--------------------------------------------------------
            bCreatLogFile = False
            sLogSqlDesc = "Step 1. Retrieve Dental Data"

            '--Insert
            sSql = "insert into Z_WORKLOAD(FACILITY_NCID, CLINIC_NCID, PROVIDER_NCID, UNIT_NUMBER, " & vbCrLf & _
                    vbTab & "DEFINITION, DRC, ENCOUNTER_DATE, COUNT_NO, MODIFIED_BY, TIME_STAMP) "
            '--select
            sSql = sSql & vbCrLf & _
                    "select  A.FACILITY_NCID, A.CLINIC_NCID, D2.PROVIDER_ID, D1.PATIENT_UNIT_ID, " & vbCrLf & _
                    vbTab & "D3.ADA_CODE ||' '|| D3.ADA_SHORT, D1.DRC, A.APPT_DATE_TIME, " & vbCrLf & _
                    vbTab & "count(*), '" & RQ(psModifiedBy) & "'," & psProcessTime
            '--from
            sSql = sSql & vbCrLf & _
                    "from APPOINTMENT A, " & vbCrLf & _
                    vbTab & " DENTAL_ENCOUNTER D1, " & vbCrLf & _
                    vbTab & " DENTAL_ENCOUNTER_TX D2, " & vbCrLf & _
                    vbTab & " TX_ADA_CODE_DEFINITION D3 "
            '--where
            sSql = sSql & vbCrLf & _
                    "where A.PATIENT_UNIT_NUMBER = D1.PATIENT_UNIT_ID " & vbCrLf & _
                    "and   A.FACILITY_NCID = D1.ENCOUNTER_FACILITY_NCID " & vbCrLf & _
                    "and   A.ENCOUNTER_NUMBER = D1.ENCOUNTER_NUMBER " & vbCrLf & _
                    "and   D1.DENTAL_ENCOUNTER_ID = D2.DENTAL_ENCOUNTER_ID " & vbCrLf & _
                    "and   D2.TX_ADA_CODE_ID = D3.TX_ADA_CODE_ID " & vbCrLf & _
                    "and   D1.ENCOUNTER_FACILITY_NCID = " & plFacilityNCID & vbCrLf & _
                    "and   D2.PROVIDER_ID in " & psInProviders & vbCrLf & _
                    "and   D2.STATUS_ID = 102 " & vbCrLf & _
                    "and   A.ENCOUNTER_STATUS = 4 "
                    
            If IsMissing(pvDateTo) Then 'For Daily
                sSql = sSql & vbCrLf & _
                    "and " & goDBConn.FormatDateForDBSelect("A.APPT_DATE_TIME", DD_MON_YYYY, True) & " = " & pvDateFrom
'                    "and   trunc(A.APPT_DATE_TIME) = " & pvDateFrom
            Else    'For Monthly
                sSql = sSql & vbCrLf & _
                    "and " & goDBConn.FormatDateForDBSelect("A.APPT_DATE_TIME", DD_MON_YYYY, True) & vbCrLf & _
                    vbTab & "between " & pvDateFrom & vbCrLf & _
                    vbTab & "and " & pvDateTo
'                    "and   trunc(A.APPT_DATE_TIME) " & vbCrLf & _
'                    vbTab & "between " & pvDateFrom & vbCrLf & _
'                    vbTab & "and " & pvDateTo
            End If
            '--group by
            sSql = sSql & vbCrLf & _
                    "group by A.FACILITY_NCID, A.CLINIC_NCID, D2.PROVIDER_ID, D1.PATIENT_UNIT_ID, " & vbCrLf & _
                    vbTab & "D3.ADA_CODE ||' '|| D3.ADA_SHORT, " & vbCrLf & _
                    vbTab & "D1.DRC, A.APPT_DATE_TIME "
        Case 2
            '---------------------------------------
            'Step 2. Make sure data selected
            '---------------------------------------
            sLogSqlDesc = "Step 2. Make sure data selected"
            
            sSql = "select count(*) " & vbCrLf & _
            "from Z_WORKLOAD Z" & vbCrLf & _
            "where " & sAppendWhereClause
        Case 3
            '----------------------------------------
            'Step 3. Get Patient info."
            '-- CDRPLUS.MMI_NAME
            '-- CDRPLUS.MMI_ID
            '-- CDRPLUS.MMI_GENERICDATA
            '-- CDRPLUS.MMI_GENERIC_ID
            '-- DENTALZ._WORKLOAD
            '----------------------------------------
            sLogSqlDesc = "Step 3. Get Patient info."
            '--insert
            sSql = "insert into Z_DAILY_RPT(FACILITY_NCID, CLINIC_NCID, PROVIDER_NCID, UNIT_NUMBER, " & vbCrLf & _
                vbTab & "DEFINITION , DRC, ENCOUNTER_DATE, " & vbCrLf & _
                vbTab & "PATIENT_NAME, " & vbCrLf & _
                vbTab & "SSN, " & vbCrLf & _
                vbTab & "SERVICE, SERVICE_DISPLAY, FMP, COUNT_NO, MODIFIED_BY, TIME_STAMP) "
            '--select
            sSql = sSql & vbCrLf & _
                "select Z.FACILITY_NCID, Z.CLINIC_NCID, Z.PROVIDER_NCID, Z.UNIT_NUMBER, " & vbCrLf & _
                vbTab & "Z.DEFINITION, Z.DRC, Z.ENCOUNTER_DATE, " & vbCrLf & _
                vbTab & "trim(substr(MN.LAST_NAME || ', ' || MN.FIRST_NAME || ' ' || substr(MN.MIDDLE_NAME, 1, 1), 1, 40)), " & vbCrLf & _
                vbTab & "substr(MI.SSN, 1, 3) || '-' ||substr(MI.SSN, 4, 2) || '-' || substr(MI.SSN, 6, 4), " & vbCrLf & _
                vbTab & "substr(MG.VALUE_TEXT,1,3), 'TBD', substr(MGI.ID_VALUE, 1, 2), Z.COUNT_NO, Z.MODIFIED_BY, Z.TIME_STAMP "
            '--from
            sSql = sSql & vbCrLf & _
                "from   MMI_NAME MN, MMI_ID MI, " & vbCrLf & _
                vbTab & "MMI_GENERICDATA MG,  " & vbCrLf & _
                vbTab & "MMI_GENERIC_ID MGI,  " & vbCrLf & _
                vbTab & "Z_WORKLOAD Z "
            '--where
            sSql = sSql & vbCrLf & _
                "where MN.UNIT_NUMBER = MI.UNIT_NUMBER " & vbCrLf & _
                "and   MN.UNIT_NUMBER = MG.UNIT_NUMBER " & vbCrLf & _
                "and   MN.UNIT_NUMBER = MGI.UNIT_NUMBER " & vbCrLf & _
                "and   MN.UNIT_NUMBER = Z.UNIT_NUMBER " & vbCrLf & _
                "and   MN.NAME_TYPE_CID = " & msNCID_PATIENT_CURRENT_NAME_USED & vbCrLf & _
                "and   MG.TYPE_NCID = " & msNCID_PATIENT_CATEGORY & vbCrLf & _
                "and   MGI.ID_TYPE_NCID = " & msNCID_PATIENT_FMP & vbCrLf & _
                "and   " & sAppendWhereClause
        Case UPDATE_SERVICE_DISP_STEP
            sLogSqlDesc = "Step " & UPDATE_SERVICE_DISP_STEP & ". Update Service Display"
            If goDBConn.UseSQLServer Then
                sSql = "UPDATE Z_DAILY_RPT " & _
                        "SET service_display = Trim(" & _
                            "(SELECT distinct(service_code) FROM dental_patcat_lkup p " & _
                            "WHERE p.patcatcode = Z_DAILY_RPT.service)) " & _
                            "WHERE " & Replace(sAppendWhereClause, "Z.", "")
            Else
                sSql = "UPDATE Z_DAILY_RPT Z " & _
                        "SET z.service_display = Trim(" & _
                            "(SELECT distinct(service_code) FROM dental_patcat_lkup p " & _
                            "WHERE p.patcatcode = z.service)) " & _
                            "WHERE " & sAppendWhereClause
            End If
            
        Case UPDATE_FAMILY_MEMBER_CHILD_STEP
            sLogSqlDesc = "Step " & UPDATE_SERVICE_DISP_STEP & ". Update Family Member Child Display"
            If goDBConn.UseSQLServer Then
                sSql = "UPDATE Z_DAILY_RPT set service_display = 'FMC' Where SERVICE_DISPLAY = 'FM' " _
                & " and fmp between '00' and '19' and " & Replace(sAppendWhereClause, "Z.", "")
            Else
                sSql = "UPDATE Z_DAILY_RPT Z set z.service_display = 'FMC' Where Z.SERVICE_DISPLAY = 'FM' " _
                & " and z.fmp between '00' and '19' and " & sAppendWhereClause
            End If
            
        Case UPDATE_FAMILY_MEMBER_ADULT_STEP
            sLogSqlDesc = "Step " & UPDATE_SERVICE_DISP_STEP & ". Update Family Member Adult Display"
            If goDBConn.UseSQLServer Then
                sSql = "UPDATE Z_DAILY_RPT set service_display = 'FMA' Where SERVICE_DISPLAY = 'FM' and fmp between '30' and '69' and " & Replace(sAppendWhereClause, "Z.", "")
            Else
                sSql = "UPDATE Z_DAILY_RPT Z set z.service_display = 'FMA' Where Z.SERVICE_DISPLAY = 'FM' and z.fmp between '30' and '69' and " & sAppendWhereClause
            End If
            
''        Case 4
''            '----------------------------------------------------------
''            'Step 4. Update Active Duty for display"
''            '----------------------------------------------------------
''            sLogSqlDesc = "Step 4. Update Active Duty for display"
''
''            sSQL = "update Z_DAILY_RPT Z set " & vbCrLf & _
''                    vbTab & "SERVICE_DISPLAY = decode(substr(SERVICE,1,1), 'A', 'Army', 'F', 'AF', 'M', 'MC', 'N', 'Navy', SERVICE)" & vbCrLf & _
''                    "where  substr(SERVICE, 1, 3) in " & _
''                    "(select pc.patcatcode from dental_patcat_lkup pc where pc.service_code in ('A', 'F', 'AF', 'M', 'N') " & vbCrLf & _
''                    "and    FMP = '20' " & vbCrLf & _
''                    "and    " & sAppendWhereClause
''        Case 5
''            '----------------------------------------------------------
''            'Step 5. Update Other military for display"
''            '----------------------------------------------------------
''            sLogSqlDesc = "Step 5. Update Other military for display"
''
''            sSQL = "update Z_DAILY_RPT Z set SERVICE_DISPLAY = 'OM' " & vbCrLf & _
''                    "where  substr(SERVICE, 1, 3) in ('A12','A15', 'F12', 'F15', 'M12','M15', 'N12','N15')" & vbCrLf & _
''                    "and    FMP = '20' " & vbCrLf & _
''                    "and    " & sAppendWhereClause
''        Case 6
''            '----------------------------------------------------------
''            'Step 6. Update Family Member Adult for display
''            '----------------------------------------------------------
''            sLogSqlDesc = "Step 6. Update Family Member Adult for display"
''
''            sSQL = "update Z_DAILY_RPT Z set SERVICE_DISPLAY = 'FMA' " & vbCrLf & _
''                    "where  substr(SERVICE, 1, 2) in ('A4','F4', 'M4','N4') " & vbCrLf & _
''                    "and    FMP between '30' and '69' " & vbCrLf & _
''                    "and    " & sAppendWhereClause
''        Case 7
''            '----------------------------------------------------------
''            'Step 7. Update Family Member Child for display
''            '----------------------------------------------------------
''            sLogSqlDesc = "Step 7. Update Family Member Child for display"
''
''            sSQL = "update Z_DAILY_RPT Z set SERVICE_DISPLAY = 'FMC' " & vbCrLf & _
''                    "where  substr(SERVICE, 1, 2) in ('A4','F4', 'M4','N4') " & vbCrLf & _
''                    "and    FMP between '00' and '19' " & vbCrLf & _
''                    "and    " & sAppendWhereClause
''        Case 8
''            '----------------------------------------------------------
''            'Step 8. Update Retired for display
''            '----------------------------------------------------------
''            sLogSqlDesc = "Step 8. Update Retired for display"
''
''            sSQL = "update Z_DAILY_RPT Z set SERVICE_DISPLAY = 'RET' " & vbCrLf & _
''                    "where  substr(SERVICE, 1, 2) in ('A3','F3', 'M3','N3') " & vbCrLf & _
''                    "and    FMP ='20'  " & vbCrLf & _
''                    "and    " & sAppendWhereClause
''        Case 9
''            '----------------------------------------------------------
''            'Step 9. Update Other for display"
''            '----------------------------------------------------------
''            sLogSqlDesc = "Step 9. Update Other for display"
''
''            sSQL = "update Z_DAILY_RPT Z set SERVICE_DISPLAY = 'OTH' " & vbCrLf & _
''                    "where  (substr(SERVICE, 1, 2) in ('A2','F2', 'M2','N2') " & vbCrLf & _
''                    "or     substr(SERVICE, 1, 1) not in ('A','F', 'M','N')) " & vbCrLf & _
''                    "and    " & sAppendWhereClause
        Case 10
            '----------------------------------------------------------
            'Step 10. <<<Final DAILY REPORT data>>>
            '----------------------------------------------------------
            sLogSqlDesc = "Step 10. <<<Final DAILY REPORT data>>> "
            
            If goDBConn.UseSQLServer Then
                sSql = "select '" & sFacilityName & "' as FACILITY_NAME, C.NAME as CLINIC_NAME, " & vbCrLf & _
                vbTab & "CP.PROVIDER_NAME, Z.ENCOUNTER_DATE, Z.DEFINITION, " & vbCrLf & _
                vbTab & "convert(varchar,Z.DRC) + ' ' + convert(varchar,Z.PATIENT_NAME) + ' ' + convert(varchar,Z.SSN) + ' ' + convert(varchar,Z.SERVICE_DISPLAY) as PATIENT_INFO, " & vbCrLf & _
                vbTab & "Z.COUNT_NO, Z.MODIFIED_BY " & vbCrLf & _
                "from  CLINIC C, CLINIC_PROVIDER CP, Z_DAILY_RPT Z " & vbCrLf & _
                "where C.NCID = Z.CLINIC_NCID " & vbCrLf & _
                "and   CP.CLINICIAN_NCID = Z.PROVIDER_NCID " & vbCrLf & _
                "and   CP.CLINIC_NCID = C.NCID " & vbCrLf & _
                "and    " & sAppendWhereClause & vbCrLf & _
                "order by 2, 3, 5, 6"
            Else
                sSql = "select '" & sFacilityName & "' as FACILITY_NAME, C.NAME as CLINIC_NAME, " & vbCrLf & _
                vbTab & "CP.PROVIDER_NAME, Z.ENCOUNTER_DATE, Z.DEFINITION, " & vbCrLf & _
                vbTab & "Z.DRC || ' ' || Z.PATIENT_NAME ||' ' || Z.SSN || ' ' || Z.SERVICE_DISPLAY as PATIENT_INFO, " & vbCrLf & _
                vbTab & "Z.COUNT_NO, Z.MODIFIED_BY " & vbCrLf & _
                "from  CLINIC C, CLINIC_PROVIDER CP, Z_DAILY_RPT Z " & vbCrLf & _
                "where C.NCID = Z.CLINIC_NCID " & vbCrLf & _
                "and   CP.CLINICIAN_NCID = Z.PROVIDER_NCID " & vbCrLf & _
                "and   CP.CLINIC_NCID = C.NCID " & vbCrLf & _
                "and    " & sAppendWhereClause & vbCrLf & _
                "order by 2, 3, 5, 6"
            End If
        Case 11
            '----------------------------------------------------------
            'Step 11. Summation by Service
            '----------------------------------------------------------
            sLogSqlDesc = "Step 11. Summation by Service"
                    '--insert
            sSql = "Insert into Z_MONTHLY_RPT( " & vbCrLf & _
                vbTab & "FACILITY_NCID, CLINIC_NCID, PROVIDER_NCID, " & vbCrLf & _
                vbTab & "DEFINITION, MONTHLY, MODIFIED_BY, TIME_STAMP, " & vbCrLf & _
                vbTab & "AF_CNT, AM_CNT, MC_CNT, NV_CNT, OM_CNT, MI_TOT, " & vbCrLf & _
                vbTab & "FMA_CT, FMC_CT, RET_CT, OTH_CT, TOTAL) "
            '--select
            sSql = sSql & vbCrLf & _
                "select FACILITY_NCID, CLINIC_NCID, PROVIDER_NCID, " & vbCrLf
            
            If goDBConn.UseSQLServer Then
                sSql = sSql & vbTab & "DEFINITION, " _
                    & "'1/' + CONVERT(varchar, DATEPART(m, ENCOUNTER_DATE)) + '/' + CONVERT(varchar, DATEPART(yy, ENCOUNTER_DATE)) AS ENCOUNTER_DATE" _
                    & ", MODIFIED_BY, TIME_STAMP, " & vbCrLf
            Else
                sSql = sSql & vbTab & "DEFINITION, Trunc(ENCOUNTER_DATE, 'Month'), MODIFIED_BY, TIME_STAMP, " & vbCrLf
            End If
            
            If goDBConn.UseSQLServer Then
                sSql = sSql & vbTab & "case SERVICE_DISPLAY when 'F' then sum(COUNT_NO) else 0 end, " & vbCrLf & _
                   vbTab & "case SERVICE_DISPLAY when 'A' then sum(COUNT_NO) else 0 end, " & vbCrLf & _
                   vbTab & "case SERVICE_DISPLAY when 'M' then sum(COUNT_NO) else 0 end, " & vbCrLf & _
                   vbTab & "case SERVICE_DISPLAY when 'N' then sum(COUNT_NO) else 0 end, " & vbCrLf & _
                   vbTab & "case SERVICE_DISPLAY when 'OM' then sum(COUNT_NO) else 0 end, " & vbCrLf & _
                   vbTab & "case SERVICE_DISPLAY when " & vbCrLf & _
                            vbTab & vbTab & "'F' then sum(COUNT_NO) " & vbCrLf & _
                            vbTab & vbTab & " when 'A' then sum(COUNT_NO) " & vbCrLf & _
                            vbTab & vbTab & " when 'M' then sum(COUNT_NO) " & vbCrLf & _
                            vbTab & vbTab & " when 'N' then sum(COUNT_NO) " & vbCrLf & _
                            vbTab & vbTab & " when 'OM' then sum(COUNT_NO) else 0 end, " & vbCrLf & _
                   vbTab & "case SERVICE_DISPLAY when 'FMA' then sum(COUNT_NO) else 0 end, " & vbCrLf & _
                   vbTab & "case SERVICE_DISPLAY when 'FMC' then sum(COUNT_NO) else 0 end, " & vbCrLf & _
                   vbTab & "case SERVICE_DISPLAY when 'R' then sum(COUNT_NO) else 0 end, " & vbCrLf & _
                   vbTab & "case SERVICE_DISPLAY when 'O' then sum(COUNT_NO) else 0 end, " & vbCrLf & _
                   vbTab & "Sum (COUNT_NO) "
            Else
                sSql = sSql & vbTab & "decode(SERVICE_DISPLAY, 'F', sum(COUNT_NO), 0), " & vbCrLf & _
                   vbTab & "decode(SERVICE_DISPLAY, 'A', sum(COUNT_NO), 0), " & vbCrLf & _
                   vbTab & "decode(SERVICE_DISPLAY, 'M', sum(COUNT_NO), 0), " & vbCrLf & _
                   vbTab & "decode(SERVICE_DISPLAY, 'N', sum(COUNT_NO), 0), " & vbCrLf & _
                   vbTab & "decode(SERVICE_DISPLAY, 'OM', sum(COUNT_NO), 0), " & vbCrLf & _
                   vbTab & "decode(SERVICE_DISPLAY, " & vbCrLf & _
                            vbTab & vbTab & "'F', sum(COUNT_NO), " & vbCrLf & _
                            vbTab & vbTab & "'A', sum(COUNT_NO), " & vbCrLf & _
                            vbTab & vbTab & "'M', sum(COUNT_NO), " & vbCrLf & _
                            vbTab & vbTab & "'N', sum(COUNT_NO), " & vbCrLf & _
                            vbTab & vbTab & "'OM', sum(COUNT_NO), 0), " & vbCrLf & _
                   vbTab & "decode(SERVICE_DISPLAY, 'FMA', sum(COUNT_NO), 0), " & vbCrLf & _
                   vbTab & "decode(SERVICE_DISPLAY, 'FMC', sum(COUNT_NO), 0), " & vbCrLf & _
                   vbTab & "decode(SERVICE_DISPLAY, 'R', sum(COUNT_NO), 0), " & vbCrLf & _
                   vbTab & "decode(SERVICE_DISPLAY, 'O', sum(COUNT_NO), 0), " & vbCrLf & _
                   vbTab & "Sum (COUNT_NO) "
            End If
            
            '--from/where
            sSql = sSql & vbCrLf & _
                "from Z_DAILY_RPT Z" & vbCrLf & _
                "where  " & sAppendWhereClause
            '--group by
            If goDBConn.UseSQLServer Then
                sSql = sSql & vbCrLf & _
                    "group by FACILITY_NCID, CLINIC_NCID, PROVIDER_NCID, " & vbCrLf & _
                    vbTab & "DEFINITION, " & "'1/' + CONVERT(varchar, DATEPART(m, ENCOUNTER_DATE)) + '/' + CONVERT(varchar, DATEPART(yy, ENCOUNTER_DATE)), MODIFIED_BY, TIME_STAMP, SERVICE_DISPLAY"
            Else
                sSql = sSql & vbCrLf & _
                    "group by FACILITY_NCID, CLINIC_NCID, PROVIDER_NCID, " & vbCrLf & _
                    vbTab & "DEFINITION, trunc(ENCOUNTER_DATE, 'MONTH'), MODIFIED_BY, TIME_STAMP, SERVICE_DISPLAY"
            End If
            
        Case 12
            '----------------------------------------------------------
            'Step 12. <<<Final MONTHLY REPORT data>>>
            '----------------------------------------------------------
            sLogSqlDesc = "Step 12. <<<Final MONTHLY REPORT data>>> "
            sSql = "select '" & sFacilityName & "' as FACILITY_NAME, C.NAME as CLINIC_NAME, " & vbCrLf & _
                "CP.PROVIDER_NAME, Z.DEFINITION, Z.MONTHLY, Z.MODIFIED_BY, " & vbCrLf & _
                "sum(Z.AF_CNT),sum(Z.AM_CNT), sum(Z.MC_CNT), sum(Z.NV_CNT), sum(Z.OM_CNT), sum(Z.MI_TOT), " & vbCrLf & _
                "Sum (Z.FMA_CT), Sum(Z.FMC_CT), Sum(Z.RET_CT), Sum(Z.OTH_CT), Sum(Z.TOTAL) " & vbCrLf & _
                "from Z_MONTHLY_RPT Z, CLINIC C, CLINIC_PROVIDER CP " & vbCrLf & _
                "where C.NCID = Z.CLINIC_NCID " & vbCrLf & _
                "and   CP.CLINICIAN_NCID = Z.PROVIDER_NCID " & vbCrLf & _
                "and   CP.CLINIC_NCID = C.NCID " & vbCrLf & _
                "and   " & sAppendWhereClause & vbCrLf & _
                "group by C.NAME, CP.PROVIDER_NAME, Z.DEFINITION, Z.MONTHLY, Z.MODIFIED_BY"
       
        Case 21
            '----------------------------------------------------------
            'Step 21. delete Z_WORKLOAD table
            '----------------------------------------------------------
            sLogSqlDesc = "Step 21. delete Z_WORKLOAD table"
            If goDBConn.UseSQLServer Then
                sSql = "delete from Z_WORKLOAD " & vbCrLf & "where " & Replace(sAppendWhereClause, "Z.", "")
            Else
                sSql = "delete from Z_WORKLOAD Z " & vbCrLf & "where " & sAppendWhereClause
            End If
            
        Case 22
            '----------------------------------------------------------
            'Step 22. delete Z_DAILY_RPT table
            '----------------------------------------------------------
            sLogSqlDesc = "Step 22. delete Z_DAILY_RPT table"
            If goDBConn.UseSQLServer Then
                sSql = "delete from Z_DAILY_RPT " & vbCrLf & "where " & Replace(sAppendWhereClause, "Z.", "")
            Else
                sSql = "delete from Z_DAILY_RPT Z " & vbCrLf & "where " & sAppendWhereClause
            End If
            
        Case 23
            '----------------------------------------------------------
            'Step 23. delete Z_MONTHLY_RPT table
            '----------------------------------------------------------
            sLogSqlDesc = "Step 23. delete Z_MONTHLY_RPT table"
            If goDBConn.UseSQLServer Then
                sSql = "delete from Z_MONTHLY_RPT " & vbCrLf & "where " & Replace(sAppendWhereClause, "Z.", "")
            Else
                sSql = "delete from Z_MONTHLY_RPT Z " & vbCrLf & "where " & sAppendWhereClause
            End If
            
        Case 100
            '----------------------------------------------------------
            'Step 100. Test Data <<<Final DAILY REPORT data>>>
            '----------------------------------------------------------
            bCreatLogFile = True
            sLogSqlDesc = "Step 100. Test Data <<<Final DAILY REPORT data>>> "
            
            sSql = "select '" & sFacilityName & "' as FACILITY_NAME, C.NAME as CLINIC_NAME, " & vbCrLf & _
            vbTab & "CP.PROVIDER_NAME, Z.ENCOUNTER_DATE, Z.DEFINITION, " & vbCrLf & _
            vbTab & "Z.DRC || ' ' || Z.PATIENT_NAME ||' ' || Z.SSN || ' ' || Z.SERVICE_DISPLAY as PATIENT_INFO, " & vbCrLf & _
            vbTab & "Z.COUNT_NO, Z.MODIFIED_BY " & vbCrLf & _
            "from  CLINIC C, CLINIC_PROVIDER CP, Z_DAILY_RPT Z " & vbCrLf & _
            "where C.NCID = Z.CLINIC_NCID " & vbCrLf & _
            "and   CP.CLINICIAN_NCID = Z.PROVIDER_NCID " & vbCrLf & _
            "and   CP.CLINIC_NCID = C.NCID "
        Case 200
            '----------------------------------------------------------
            'Step 200. Test Data <<<Final MONTHLY REPORT data>>>
            '----------------------------------------------------------
            bCreatLogFile = True
            sLogSqlDesc = "Step 200. Test Data <<<Final MONTHLY REPORT data>>> "
            
            sSql = "select '" & sFacilityName & "' as FACILITY_NAME, C.NAME as CLINIC_NAME, " & vbCrLf & _
                "CP.PROVIDER_NAME, Z.DEFINITION, Z.MONTHLY, Z.MODIFIED_BY, " & vbCrLf & _
                "sum(Z.AF_CNT),sum(Z.AM_CNT), sum(Z.MC_CNT), sum(Z.NV_CNT), sum(Z.OM_CNT), sum(Z.MI_TOT), " & vbCrLf & _
                "Sum (Z.FMA_CT), Sum(Z.FMC_CT), Sum(Z.RET_CT), Sum(Z.OTH_CT), Sum(Z.TOTAL) " & vbCrLf & _
                "from Z_MONTHLY_RPT Z, CLINIC C, CLINIC_PROVIDER CP " & vbCrLf & _
                "where C.NCID = Z.CLINIC_NCID " & vbCrLf & _
                "and   CP.CLINICIAN_NCID = Z.PROVIDER_NCID " & vbCrLf & _
                "and   CP.CLINIC_NCID = C.NCID " & vbCrLf & _
                "group by C.NAME, CP.PROVIDER_NAME, Z.DEFINITION, Z.MONTHLY, Z.MODIFIED_BY"

    End Select
                       
    '--log sql
    If mbLogSql Then
        LogSql sLogSqlDesc, sSql, bCreatLogFile
    End If
     
    '--execute mothed
    Set prRs = Nothing
    
    If goDBConn.UseSQLServer Then
        'SQL Server format
        sSql = Replace(sSql, "||", "+")
        sSql = Replace(LCase(sSql), "substr", "SUBSTRING")
        sSql = Replace(LCase(sSql), "trim", "ltrim")
    End If
    
    Select Case piProcessStep
        Case 2, 10, 12, 100, 200 'for select sql only
            Set prRs = goDBDAS.OpenRecordset(sSql)
        'Case 4, 5, 6, 7, 8, 9
        '    goDBDAS.ExecuteSQL sSQL
        Case Else
            goDBSqlOps.Execute sSql
    End Select
            '
    'PART II - SCRs 8905/8906 RWalker 12/4/2003
    'Retrieve workload credits for the provider from other provider's encounters
    If piProcessStep = miNOTES_WORKLOAD Then

        '--Insert
        sSql = "insert into Z_WORKLOAD(FACILITY_NCID, CLINIC_NCID, PROVIDER_NCID, UNIT_NUMBER, " & vbCrLf & _
                vbTab & "DEFINITION, DRC, ENCOUNTER_DATE, COUNT_NO, MODIFIED_BY, TIME_STAMP) "
        
        '--select
        sSql = sSql & vbCrLf & _
                "select  A.FACILITY_NCID, A.CLINIC_NCID, D2.Provider_ID, D1.PATIENT_UNIT_ID, " & vbCrLf & _
                vbTab & "D3.ADA_CODE ||' '|| D3.ADA_SHORT, D1.DRC, A.APPT_DATE_TIME, " & vbCrLf & _
                vbTab & "sum(MULTIPLIER), '" & RQ(psModifiedBy) & "'," & psProcessTime
        
        '--from
        sSql = sSql & vbCrLf & _
                "from APPOINTMENT A, " & vbCrLf & _
                vbTab & " DENTAL_ENCOUNTER D1, " & vbCrLf & _
                vbTab & " TEMPL_WORKLOAD D2, " & vbCrLf & _
                vbTab & " TX_ADA_CODE_DEFINITION D3 "
        
        sSql = sSql & vbCrLf & _
                "where A.PATIENT_UNIT_NUMBER = D1.PATIENT_UNIT_ID " & vbCrLf & _
                "and   A.FACILITY_NCID = D1.ENCOUNTER_FACILITY_NCID " & vbCrLf & _
                "and   A.ENCOUNTER_NUMBER = D1.ENCOUNTER_NUMBER " & vbCrLf & _
                "and   D1.DENTAL_ENCOUNTER_ID = D2.DENTAL_ENCOUNTER_ID " & vbCrLf & _
                "and   D2.TX_ADA_CODE_ID = D3.TX_ADA_CODE_ID " & vbCrLf & _
                "and   D1.ENCOUNTER_FACILITY_NCID = " & plFacilityNCID & vbCrLf & _
                "and   A.CLINICIAN_NCID NOT in " & psInProviders & vbCrLf & _
                "and   D2.Provider_ID in " & psInProviders & vbCrLf & _
                "and   D2.DELETED = 0 " & vbCrLf & _
                "and   A.ENCOUNTER_STATUS in (4,6) "

        If IsMissing(pvDateTo) Then 'For Daily
            sSql = sSql & vbCrLf & _
                "and " & goDBConn.FormatDateForDBSelect("A.Appt_Date_Time", DD_MON_YYYY, True) & " = " & pvDateFrom
        Else    'For Monthly
            sSql = sSql & vbCrLf & _
                "and " & goDBConn.FormatDateForDBSelect("A.APPT_DATE_TIME", DD_MON_YYYY, True) & vbCrLf & _
                vbTab & "between " & pvDateFrom & vbCrLf & _
                vbTab & "and " & pvDateTo
            End If
            '--group by
            sSql = sSql & vbCrLf & _
                    "group by A.FACILITY_NCID, A.CLINIC_NCID, " & vbCrLf & _
                    "D2.Provider_ID, D1.PATIENT_UNIT_ID, " & vbCrLf & _
                    vbTab & "D3.ADA_CODE ||' '|| D3.ADA_SHORT, " & vbCrLf & _
                    vbTab & "D1.DRC, A.APPT_DATE_TIME "
        
        If goDBConn.UseSQLServer Then
            'SQL Server format
            sSql = Replace(sSql, "||", "+")
            sSql = Replace(LCase(sSql), "substr", "SUBSTRING")
            sSql = Replace(LCase(sSql), "trim", "ltrim")
        End If
    
        goDBSqlOps.Execute sSql
    
    End If
    
    SqlProcess = True
    GoTo CleanUp
    
ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "DTRptProvider.SqlProcess of step " & piProcessStep, "CHCSII_Reports", vbExclamation
    Exit Function
    
    Err.Clear
    Resume
    
CleanUp:
End Function

Private Function SkipProcess4Test(ByVal lPeportItem As Long, prRs As ADODB.Recordset) As Boolean
    If lPeportItem = DailyProvider Then
        '----------------------------------------------------------
        'Step 100. <<<Final DAILY REPORT data>>>
        '----------------------------------------------------------
        If Not SqlProcess(100, "", "", prRs:=prRs) Then GoTo CleanUp
    Else
        '----------------------------------------------------------
        'Step 200. <<<Final MONTHLY REPORT data>>>
        '----------------------------------------------------------
        If Not SqlProcess(200, "", "", prRs:=prRs) Then GoTo CleanUp
    End If
    SkipProcess4Test = True
    Exit Function
CleanUp:
    SkipProcess4Test = False
End Function

'*******************************************************************
'Function: RQ
'Author: Scott Foster
'Date: 11/28/01
'Returns: String
'Purpose: To replace single quotes with two single quotes for sql statements
'*******************************************************************
Private Function RQ(ByVal Value As String) As String
    
    'Replace regular single quotes
    RQ = Replace(Value, "'", "''")
    
    'Replace MS Word single quotes
    RQ = Replace(RQ, Chr(146), "''")
    
    'Replace MS Word open double quotes
    RQ = Replace(RQ, Chr(147), """")
   
    'Replace MS Word close double quotes
    RQ = Replace(RQ, Chr(148), """")
   
End Function
