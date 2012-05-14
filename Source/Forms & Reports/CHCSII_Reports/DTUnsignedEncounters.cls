VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DTUnsignedEncounters"
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
'01/22/01   Wilson S.   Created
'04/10/01   Wilson S.   Modified SqlProcess case 1, adding checking signed encounter code
'05/02/01   Wilson S.   Modified SqlProcess case 1,
'                       chage trunc(D1.ENCOUNTER_DATE) to trunc(A.APPT_DATE_TIME)
'********************************************************************************************
Option Explicit

Private Const msDATE_FORMAT_VB   As String = "mm/dd/yyyy hh:mm:ss AM/PM"
Private Const msDATE_FORMAT_ORA  As String = "MM/DD/YYYY HH:MI:SS AM"
Private Const msNCID_PATIENT_CATEGORY As String = "14501721"
Private Const msNCID_PATIENT_FMP As String = "204708"
Private Const msNCID_PATIENT_CURRENT_NAME_USED As String = "187"

Private mrsMain As ADODB.Recordset
Private mbLogSql As Boolean

'***********************************************************************
'Function name: BuildReport
'Description:   This is the main interface for calling to print report
'Parameters:
'   pRepItem    - Report Info
'   prRs        - for final report data
'Return Value:  Ture, sucessful; False, otherwise
'History
'Date       Developer   Description
'--------   ----------  -----------------------------------------------------------------------
'10/22/01   Wilson Sun  Created. Clone from DTRptProvider.cls
'***********************************************************************
Public Function BuildReport(ByVal pRepItem As ReportItem, _
                            ByRef prRs As ADODB.Recordset) As Boolean
        
    Dim sFacilityNCID As String
    Dim sLogSqlDesc As String
    Dim sProcessTime As String
    Dim sSql As String
    
    On Error GoTo ErrHandler
    BuildReport = False
    
    '--Login FacilityNCID
    sFacilityNCID = cwobjLogonEnviron.FacilityNCID
    
    '--TimeStamp *Important for application concurrency.
    '--convert parameter date value to oracle date format
    sProcessTime = SQLDate(Now) 'this variable will also be used for delete condition

    
    '**********************
    '--Start Process
    '**********************
    If Not SqlProcess(pRepItem, sFacilityNCID, sProcessTime, prRs:=prRs) Then GoTo CleanUp
    
    If prRs.RecordCount = 0 Then
        MsgBxARMd "No data selected for the report."
        GoTo CleanUp
    End If
    
    BuildReport = True

    GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "DTUnsignedEncounters.BuildReport", "CHCSII_Reports", vbExclamation
CleanUp:
    Set mrsMain = Nothing

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
'   Wilson Sun  10/22/01    Created. Clone from DTRptProvider.SqlProcess()
'**********************************************************************************
Private Function SqlProcess(ByVal pRepItem As ReportItem, _
                            ByVal plFacilityNCID As String, _
                            ByVal psProcessTime As String, _
                            Optional ByRef prRs As ADODB.Recordset) As Boolean
    '--Boolean
    Dim bCreatLogFile As Boolean
    
    '--String
    Dim sFilter As String
    Dim sLogSqlDesc As String
    Dim sSql As String
        
    On Error GoTo ErrHandler

    SqlProcess = False
    
    bCreatLogFile = False
    '--------------------------
    '-- Get Selection Criteria
    '--------------------------
    '--Provider NCIDS
    If Not GetSQLCriteria(pRepItem, 1, sFilter) Then GoTo CleanUp
    
    '--------------------------------------------------------
    'Step 1.Retrieving Data from following tables
    '   APPOINTMENT A,
    '   CLINIC C,
    '   CLINIC_PROVIDER CP,
    '   DENTAL_ENCOUNTER D,
    '   MMI_ID I,
    '   MMI_NAME N
    '--------------------------------------------------------
    bCreatLogFile = True
    Select Case pRepItem.Criteria1
       Case Clinic
            sLogSqlDesc = "Retrieving Data By Selected Clinic"
            
            If pRepItem.CriteriaSel1.Count > 1 Then
                MsgBxARMd "More than one Clinic/Lab selection is not allowed for this report.", , "Clinic/Lab selection"
                GoTo CleanUp
            End If

            '-- select
            sSql = "select 'Clinic/Lab' as caption_1, " & vbCrLf & _
                "'Provider' as caption_2, " & vbCrLf & _
                "'(SSN) Patient' as caption_3, " & vbCrLf & _
                "C.NAME ||' (MEPRS: ' || MEPRS_CODE || ', DMIS: ' || DMIS_ID_CODE ||')' as COLUMN_1," & vbCrLf & _
                "CP.PROVIDER_NAME as COLUMN_2, " & vbCrLf & _
                "'(' ||substr(I.SSN, 1, 3) || '-' ||substr(I.SSN, 4, 2) || '-' || substr(I.SSN, 6, 4) || ') ' || " & vbCrLf & _
                "trim(substr(N.LAST_NAME || ', ' || N.FIRST_NAME || ' ' || substr(N.MIDDLE_NAME, 1, 1), 1, 40)) as COLUMN_3, "
        
        Case Provider
            sLogSqlDesc = "Retrieving Data By Selected Provider"

            '-- select
            sSql = "select 'Provider' as caption_1, " & vbCrLf & _
                "'Clinic/Lab' as caption_2, " & vbCrLf & _
                "'(SSN) Patient' as caption_3, " & vbCrLf & _
                "CP.PROVIDER_NAME as COLUMN_1, " & vbCrLf & _
                "C.NAME ||' (MEPRS: ' || MEPRS_CODE || ', DMIS: ' || DMIS_ID_CODE ||')' as COLUMN_2," & vbCrLf & _
                "'(' ||substr(I.SSN, 1, 3) || '-' ||substr(I.SSN, 4, 2) || '-' || substr(I.SSN, 6, 4) || ') ' || " & vbCrLf & _
                "trim(substr(N.LAST_NAME || ', ' || N.FIRST_NAME || ' ' || substr(N.MIDDLE_NAME, 1, 1), 1, 40)) as COLUMN_3, "
        
        Case Patient
            sLogSqlDesc = "Retrieving Data By Selected Patient"

            sSql = "select '(SSN) Patient' as caption_1, " & vbCrLf & _
                "'Clinic/Lab' as caption_2, " & vbCrLf & _
                "'Provider' as caption_3, " & vbCrLf & _
                "'(' ||substr(I.SSN, 1, 3) || '-' ||substr(I.SSN, 4, 2) || '-' || substr(I.SSN, 6, 4) || ') ' || " & vbCrLf & _
                "trim(substr(N.LAST_NAME || ', ' || N.FIRST_NAME || ' ' || substr(N.MIDDLE_NAME, 1, 1), 1, 40)) as COLUMN_1, " & vbCrLf & _
                "C.NAME ||' (MEPRS: ' || MEPRS_CODE || ', DMIS: ' || DMIS_ID_CODE ||')' as COLUMN_2," & vbCrLf & _
                "CP.PROVIDER_NAME as COLUMN_3, "
    End Select
    '--select
    sSql = sSql & vbCrLf & _
        "A.APPT_DATE_TIME, " & vbCrLf & _
        "A.ENCOUNTER_NUMBER, "
    
    If goDBConn.UseSQLServer Then
        'SQL Server format
        sSql = sSql & " CASE A.ENCOUNTER_STATUS WHEN 1 THEN 'CheckedIn' WHEN 2 THEN 'Waiting' WHEN 3 THEN 'InProgress' WHEN" & vbCrLf & _
        " 4 THEN 'Completed' WHEN 5 THEN 'NeedsCosignature' WHEN 6 THEN 'Update' ELSE 'Unknown' END as Encounter_status," & vbCrLf & _
        "D.DENTAL_ENCOUNTER_ID , A.Clinic_NCID, A.Clinician_NCID, A.Patient_Unit_Number "
    Else
        'Oracle format
        sSql = sSql & " decode(A.ENCOUNTER_STATUS, 1, 'CheckedIn', 2, 'Waiting', 3,'InProgress'," & vbCrLf & _
        "4,'Completed', 5, 'NeedsCosignature', 6, 'Update', 'Unknow') as Encounter_status," & vbCrLf & _
        "D.DENTAL_ENCOUNTER_ID , A.Clinic_NCID, A.Clinician_NCID, A.Patient_Unit_Number "
    End If
    
    '-- from
    sSql = sSql & vbCrLf & _
        "from APPOINTMENT A, " & vbCrLf & _
        vbTab & " CLINIC C, " & vbCrLf & _
        vbTab & " CLINIC_PROVIDER CP, " & vbCrLf & _
        vbTab & " DENTAL_ENCOUNTER D, " & vbCrLf & _
        vbTab & " MMI_ID I, " & vbCrLf & _
        vbTab & " MMI_NAME N "
    '--where
    sSql = sSql & vbCrLf & _
        "where A.FACILITY_NCID = D.ENCOUNTER_FACILITY_NCID " & vbCrLf & _
        "and   A.PATIENT_UNIT_NUMBER = D.PATIENT_UNIT_ID " & vbCrLf & _
        "and   A.ENCOUNTER_NUMBER = D.ENCOUNTER_NUMBER " & vbCrLf & _
        "and   A.CLINIC_NCID = C.NCID " & vbCrLf & _
        "and   C.NCID = CP.CLINIC_NCID " & vbCrLf & _
        "and   A.CLINICIAN_NCID = CP.CLINICIAN_NCID " & vbCrLf & _
        "and   A.PATIENT_UNIT_NUMBER = I.UNIT_NUMBER " & vbCrLf & _
        "and   A.PATIENT_UNIT_NUMBER = N.UNIT_NUMBER " & vbCrLf & _
        "and   A.FACILITY_NCID = " & plFacilityNCID & vbCrLf & _
        "and   A.ENCOUNTER_STATUS <> 4 " & vbCrLf & _
        "and   " & sFilter
    '--order by
    sSql = sSql & vbCrLf & _
        "order by C.NAME, CP.PROVIDER_NAME, LAST_NAME, N.FIRST_NAME, N.MIDDLE_NAME, A.APPT_DATE_TIME desc"
    
    If goDBConn.UseSQLServer Then
        sSql = Replace(sSql, "||", "+")
        sSql = Replace(UCase(sSql), "SUBSTR", "SUBSTRING")
        sSql = Replace(UCase(sSql), "TRIM", "LTRIM")
    End If
    
    '--log sql
    If mbLogSql Then
        LogSql sLogSqlDesc, sSql, bCreatLogFile
    End If
     
    '--execute mothed
    Set prRs = goDBDAS.OpenRecordset(sSql)
    SqlProcess = True
    GoTo CleanUp
ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "DTUnsignedEncounters.SqlProcess.", "CHCSII_Reports", vbExclamation
CleanUp:
End Function


