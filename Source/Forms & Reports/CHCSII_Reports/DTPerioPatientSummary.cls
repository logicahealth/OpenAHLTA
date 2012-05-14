VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DTPerioPatientSummary"
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
'02/23/02   Wilson Sun  Modified based on SCR 17228 Design Note
'12/13/01   Wilson Sun  Created. Clone from DTUnsignedEncounters.cls
'***********************************************************************
Public Function BuildReport(ByVal pRepItem As ReportItem, _
                            ByRef prRs As ADODB.Recordset) As Boolean
        
    Dim sFacilityNCID As String
    Dim sModifiedBy As String
    Dim sProcessTime As String
    Dim sSql As String
    
    On Error GoTo ErrHandler
    BuildReport = False
    
    '--Login FacilityNCID
    sFacilityNCID = cwobjLogonEnviron.FacilityNCID
    
    '--Login UserName
    sModifiedBy = cwobjLogonEnviron.UserName

    '--TimeStamp *Important for application concurrency.
    '--convert parameter date value to oracle date format
    sProcessTime = goDBConn.FormatDateForDBInsert(Now) ' OracleDate(Now) 'this variable will also be used for delete condition

    
    '**********************
    '--Start Process
    '**********************
    If Not SqlProcess(1, sModifiedBy, sProcessTime, , pRepItem, pbCreatLogFile:=True) Then GoTo CleanUp
    If Not SqlProcess(2, sModifiedBy, sProcessTime) Then GoTo CleanUp
    If Not SqlProcess(3, sModifiedBy, sProcessTime) Then GoTo CleanUp
    If Not SqlProcess(4, sModifiedBy, sProcessTime) Then GoTo CleanUp
    If Not SqlProcess(5, sModifiedBy, sProcessTime) Then GoTo CleanUp
    If Not SqlProcess(6, sModifiedBy, sProcessTime, , pRepItem, prRs) Then GoTo CleanUp
    
    If prRs.RecordCount = 0 Then
        MsgBxARMd "No data selected for the report."
        GoTo CleanUp
    End If
    
    BuildReport = True

    GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "DTPerioPatientSummary.BuildReport", "CHCSII_Reports", vbExclamation
CleanUp:
    
    '*******************
    'Delete Work tables
            'Z_PERIO_PHASE_DRC
            'Z_PERIO_PHASE_DRC_DATE
            'Z_PERIO_PRIMARY_DX
            'Z_PERIO_DX_TX
    '*******************
    '----------------------------------------------------------
    'Step 7. delete Z_PERIO_PHASE_DRC table
    '----------------------------------------------------------
    Call SqlProcess(7, sModifiedBy, sProcessTime, "Z_PERIO_PHASE_DRC")

    '----------------------------------------------------------
    'Step 8. delete Z_PERIO_PHASE_DRC_DATE table
    '----------------------------------------------------------
    Call SqlProcess(8, sModifiedBy, sProcessTime, "Z_PERIO_PHASE_DRC_DATE")

    '----------------------------------------------------------
    'Step 9. delete Z_PERIO_PRIMARY_DX table
    '----------------------------------------------------------
    Call SqlProcess(9, sModifiedBy, sProcessTime, "Z_PERIO_PRIMARY_DX")

    '----------------------------------------------------------
    'Step 10. delete Z_PERIO_DX_TX table
    '----------------------------------------------------------
    Call SqlProcess(10, sModifiedBy, sProcessTime, "Z_PERIO_DX_TX")

    '******************************
    '-- Process Completed
    '******************************
    '--LogSql
    If mbLogSql Then
        LogSql "Process Complete On:", Format(Now)
    End If
End Function
Private Sub Class_Initialize()
    If cwobjCWComm.CmdLineSwitch("LOGSQL") = "" Then
        mbLogSql = False
    Else
        mbLogSql = True
    End If
End Sub
Private Sub Class_Terminate()
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
'   pRepItem        - Report Item
'   prRs            - Record set for the report
'Return Value:
'   True - sucessful,
    'False - failed
'Called by:     BuildReport()
'History
'   Developer   Date        Description
'   =========   ==========  ==========================================================
'   Wilson Sun  02/23/02    Modified based on SCR 17228 Design Note
'   Wilson Sun  12/13/01    Created. Clone from DTUnsignedEncounters.SqlProcess()
'''Private Function SqlProcess(ByVal pRepItem As ReportItem,
'                            ByVal plFacilityNCID As String,
'                            ByVal psProcessTime As String,
'                            Optional ByRef prRs As ADODB.Recordset) As Boolean
'**********************************************************************************
Private Function SqlProcess(ByVal piProcessStep As Integer, _
                            ByVal psModifiedBy As String, _
                            ByVal psProcessTime As String, _
                            Optional ByVal plFacilityNCID As String, _
                            Optional ByVal pRepItem As ReportItem = Nothing, _
                            Optional ByRef prRs As ADODB.Recordset, _
                            Optional ByVal pbCreatLogFile As Boolean = False) As Boolean
       
    '--String
    Dim sAppendWhereClause As String
    Dim sCol1Caption As String
    Dim sCol2Caption As String
    Dim sCol1Value As String
    Dim sCol2Value As String
    Dim sClinicName As String
    Dim sClinicNCID As String
    Dim sFilter As String
    Dim sLogSqlDesc As String
    Dim sProviderName As String
    Dim sProviderNCID As String
    Dim sSql As String
    Dim sZTable As String
        
    On Error GoTo ErrHandler

    SqlProcess = False
    
    If Not pRepItem Is Nothing Then
        '--------------------------
        '-- Get Selection Criteria
        '--------------------------
        '--Provider and Clinic name
        sClinicName = pRepItem.CriteriaSel1.Item(1).Clinic_Name
        sClinicNCID = pRepItem.CriteriaSel1.Item(1).Clinic_NCID
        sProviderName = pRepItem.CriteriaSel1.Item(1).Name
        sProviderNCID = pRepItem.CriteriaSel1.Item(1).NCID
    End If
     
    sAppendWhereClause = "Z.MODIFIED_BY = '" & RQ(psModifiedBy) & "'" & vbCrLf & _
       "AND   Z.TIME_STAMP = " & psProcessTime

    Select Case piProcessStep
        Case 1
        'Step 7. delete Z_PERIO_PHASE_DRC table
           sLogSqlDesc = "Step 1a. Perio Tx Phase, DRC"
            '-- Insert
            sSql = "INSERT INTO Z_PERIO_PHASE_DRC "
            '--select
            sSql = sSql & vbCrLf & _
                "SELECT DISTINCT V.PATIENT_UNIT_ID, V.DENTAL_ENCOUNTER_ID, V.TREATPHASE,SUBSTR(M.VALUE_TEXT, 1, 4) AS DRC," & vbCrLf & _
                "'" & RQ(psModifiedBy) & "'," & psProcessTime
            If goDBConn.UseSQLServer Then
                sSql = sSql & " FROM APPOINTMENT A INNER JOIN MMI_GENERICDATA M ON A.PATIENT_UNIT_NUMBER = M.UNIT_NUMBER RIGHT OUTER JOIN" _
                    & " V_PERIO_LAST_PHASE V ON A.PATIENT_UNIT_NUMBER = V.PATIENT_UNIT_ID" _
                    & " Where (A.Clinician_NCID = " & sProviderNCID & ") And (M.TYPE_NCID = 14501725)"
            Else
                '-- from
                sSql = sSql & vbCrLf & _
                    "FROM APPOINTMENT A," & vbCrLf & _
                    vbTab & "V_PERIO_LAST_PHASE V," & vbCrLf & _
                    vbTab & "MMI_GENERICDATA M"
                '-- where
                sSql = sSql & vbCrLf & _
                    "WHERE A.PATIENT_UNIT_NUMBER (+) = V.PATIENT_UNIT_ID" & vbCrLf & _
                    "AND   M.UNIT_NUMBER = A.PATIENT_UNIT_NUMBER" & vbCrLf & _
                    "AND   A.CLINICIAN_NCID = " & sProviderNCID & _
                    "AND   M.TYPE_NCID = 14501725"
            End If
            
            '--order by
            sSql = sSql & vbCrLf & _
                "ORDER BY V.PATIENT_UNIT_ID"
        Case 2
        'Step 8. delete Z_PERIO_PHASE_DRC_DATE table
            sLogSqlDesc = "Step 2. Perio Tx Phase, DRC, First/Last Visit Date"
            sSql = "INSERT INTO Z_PERIO_PHASE_DRC_DATE " & vbCrLf & _
                "SELECT Z.PATIENT_UNIT_ID, Z.DENTAL_ENCOUNTER_ID, Z.TREATPHASE, Z.DRC, V.FIRST_VISIT, V.LAST_VISIT," & vbCrLf & _
                "Z.MODIFIED_BY, Z.TIME_STAMP " & vbCrLf
                
                If goDBConn.UseSQLServer Then
                    sSql = sSql & "FROM Z_PERIO_PHASE_DRC Z LEFT OUTER JOIN" _
                        & " V_PERIO_DATE V ON Z.PATIENT_UNIT_ID = V.PATIENT_UNIT_ID" _
                        & " WHERE " & sAppendWhereClause
                Else
                    sSql = sSql & "FROM Z_PERIO_PHASE_DRC Z, V_PERIO_DATE V " & vbCrLf & _
                    "WHERE  Z.PATIENT_UNIT_ID = V.PATIENT_UNIT_ID (+)" & vbCrLf & _
                    "AND   " & sAppendWhereClause
                End If
        Case 3
        'Step 9. delete Z_PERIO_PRIMARY_DX table
            sLogSqlDesc = "Step 3. Primary DX"
            If goDBConn.UseSQLServer Then
                sSql = "INSERT INTO Z_PERIO_PRIMARY_DX" _
                    & " SELECT Z.PATIENT_UNIT_ID, Z.TREATPHASE, Z.FIRST_VISIT, Z.LAST_VISIT, Z.DRC, V.ICD_DEFINITION, Z.MODIFIED_BY, Z.TIME_STAMP" _
                    & " FROM Z_PERIO_PHASE_DRC_DATE Z LEFT OUTER JOIN" _
                    & " V_PERIO_PRIMARY_DX V ON Z.PATIENT_UNIT_ID = V.PATIENT_UNIT_ID" _
                    & " WHERE " & sAppendWhereClause
            Else
                sSql = "INSERT INTO Z_PERIO_PRIMARY_DX" & vbCrLf & _
                    "SELECT Z.PATIENT_UNIT_ID, Z.TREATPHASE, Z.FIRST_VISIT, Z.LAST_VISIT, Z.DRC, V.ICD_DEFINITION," & vbCrLf & _
                    "Z.MODIFIED_BY , Z.TIME_STAMP" & vbCrLf & _
                    "FROM Z_PERIO_PHASE_DRC_DATE Z, V_PERIO_PRIMARY_DX V" & vbCrLf & _
                    "WHERE Z.PATIENT_UNIT_ID = V.PATIENT_UNIT_ID (+)" & vbCrLf & _
                    "AND   " & sAppendWhereClause
            End If
        Case 4
        'Step 10. delete Z_PERIO_DX_TX table
            sLogSqlDesc = "Step 4. Planned and Inprogress TX"
            If goDBConn.UseSQLServer Then
                sSql = "INSERT INTO Z_PERIO_DX_TX SELECT DISTINCT" _
                    & " Z.PATIENT_UNIT_ID, Z.TREATPHASE, Z.DRC, Z.FIRST_VISIT, Z.LAST_VISIT, 'No  ' AS MEMBRANE, Z.ICD_DEFINITION, V.PROVIDER_ID," _
                    & " V.DENTAL_ENCOUNTER_ID , V.STATUS_ID, V.ADA_CODE, V.ADA_SHORT, Z.MODIFIED_BY, Z.TIME_STAMP" _
                    & " FROM Z_PERIO_PRIMARY_DX Z LEFT OUTER JOIN" _
                    & " V_PERIO_TX V ON Z.PATIENT_UNIT_ID = V.PATIENT_UNIT_ID" _
                    & " WHERE " & sAppendWhereClause
            Else
                sSql = "INSERT INTO Z_PERIO_DX_TX" & vbCrLf & _
                    " SELECT DISTINCT Z.PATIENT_UNIT_ID, Z.TREATPHASE, Z.DRC, Z.FIRST_VISIT, Z.LAST_VISIT, 'No  ' AS MEMBRANE, Z.ICD_DEFINITION," & vbCrLf & _
                    " V.PROVIDER_ID , V.DENTAL_ENCOUNTER_ID, V.STATUS_ID, V.ADA_CODE, V.ADA_SHORT, Z.MODIFIED_BY, Z.TIME_STAMP" & vbCrLf & _
                    " FROM Z_PERIO_PRIMARY_DX Z, V_PERIO_TX V" & vbCrLf & _
                    " WHERE Z.PATIENT_UNIT_ID = V.PATIENT_UNIT_ID (+)" & vbCrLf & _
                    " AND   " & sAppendWhereClause
            End If
        Case 5
            sLogSqlDesc = "Step 5. Update Membrane"
            If goDBConn.UseSQLServer Then
                sSql = "Update Z_PERIO_DX_TX SET MEMBRANE = 'Yes'" _
                    & " WHERE (PATIENT_UNIT_ID IN" _
                    & " (SELECT PATIENT_UNIT_ID From Z_PERIO_DX_TX" _
                    & " WHERE STATUS_ID = 101)) AND " & Replace(sAppendWhereClause, "Z.", "")
            Else
                sSql = "UPDATE Z_PERIO_DX_TX Z SET Z.MEMBRANE = 'Yes' " & vbCrLf & _
                    "WHERE Z.PATIENT_UNIT_ID IN " & vbCrLf & _
                        "(SELECT PATIENT_UNIT_ID FROM Z_PERIO_DX_TX WHERE STATUS_ID = 101)" & vbCrLf & _
                    "AND   " & sAppendWhereClause
            End If
        Case 6
            '--------------------------------------------------------
            'Step 6. Records for Report
            'From Z_PERIO_PRIMARY_DX,Z_PERIO_TX,MMI_ID,MMI_NAME tables
            '--------------------------------------------------------
            sLogSqlDesc = "Step 6. Records for Report"
        
            'SCR 34168 RWalker 9/10/2003 - Start
            'The order by drop-down options do not mention the SSN, only by the
            'Patient Name or Perio Tx Phase.  Thus, the SSN will remain in column 1
            'with the caption modified to reflect the selection and appended with "/SSN".
            If pRepItem.SortItemSelected = "PATIENT_NAME" Then
                sCol1Caption = "'Patient Name/SSN' as COL1_CAPTION,"
                sCol1Value = "MN.LAST_NAME || ', ' || MN.FIRST_NAME || ' ' || MN.MIDDLE_NAME " & vbCrLf & _
                    "|| CHR(10) || SUBSTR(TRIM(MI.SSN), 1, 3) || '-' || SUBSTR(TRIM(MI.SSN), 4, 2) || '-' || SUBSTR(TRIM(MI.SSN), 6, 4) as COL1_VALUE,"
                sCol2Caption = "'Perio Tx Phase' as COL2_CAPTION,"
                sCol2Value = "Z.TREATPHASE as COL2_VALUE,"
            Else
                sCol1Caption = "'Perio Tx Phase' as COL1_CAPTION,"
                sCol1Value = "Z.TREATPHASE as COL1_VALUE,"
                sCol2Caption = "'Patient Name/SSN' as COL2_CAPTION,"
                sCol2Value = "MN.LAST_NAME|| ', '||MN.FIRST_NAME||' '||MN.MIDDLE_NAME" & vbCrLf & _
                    "||CHR(10)||SUBSTR(TRIM(MI.SSN),1,3)||'-'||SUBSTR(TRIM(MI.SSN),4,2)||'-'||SUBSTR(TRIM(MI.SSN),6,4) as COL2_VALUE,"
            End If
            '-- select
            sSql = "SELECT " & vbCrLf & _
                sCol1Caption & sCol2Caption & "'" & RQ(sProviderName) & "' as PROVIDER_NAME," & vbCrLf & _
                "'" & sClinicName & "' as CLINIC_NAME," & vbCrLf & _
                sCol1Value & sCol2Value & vbCrLf & _
                " MN.LAST_NAME || ', ' || MN.FIRST_NAME || ' ' || MN.MIDDLE_NAME as PATIENT_NAME," & vbCrLf & _
                "Z.DRC, Z.FIRST_VISIT, Z.LAST_VISIT, Z.MEMBRANE," & vbCrLf & _
                "SUBSTR(Z.ICD_DEFINITION,1, 145) as ICD_DEFINITION, Z.ADA_SHORT"
            '-- from
            sSql = sSql & vbCrLf & _
                "FROM Z_PERIO_DX_TX Z, " & vbCrLf & _
                vbTab & " MMI_NAME MN, " & vbCrLf & _
                vbTab & " MMI_ID MI"
           '--where
            sSql = sSql & vbCrLf & _
            "WHERE Z.PATIENT_UNIT_ID = MN.UNIT_NUMBER " & vbCrLf & _
                "AND   MN.UNIT_NUMBER = MI.UNIT_NUMBER " & vbCrLf & _
                "AND   MN.NAME_TYPE_CID = 187 " & vbCrLf & _
                "AND   (Z.STATUS_ID = 100 OR Z.STATUS_ID IS NULL)" & vbCrLf & _
                "AND   " & sAppendWhereClause
            '--order by
            sSql = sSql & vbCrLf & _
                "ORDER BY COL1_VALUE " & pRepItem.SortDirection & ", COL2_VALUE"
         Case Else
            '----------------------------------------------------------
            'Step 7-12. delete Z tables
            'Z_PERIO_DATE
            'Z_PERIO_LAST_PHASE
            'ok Z_PERIO_PHASE_DRC
            'ok Z_PERIO_PHASE_DRC_DATE
            'ok Z_PERIO_PRIMARY_DX
            'ok Z_PERIO_DX_TX
            '----------------------------------------------------------
            sZTable = plFacilityNCID
            sLogSqlDesc = "Step " & CStr(piProcessStep) & ". Delete Worktable " & sZTable
            If goDBConn.UseSQLServer Then
                sSql = "DELETE FROM " & sZTable & " where " & Replace(sAppendWhereClause, "Z.", "")
            Else
                sSql = "DELETE FROM " & sZTable & " Z where " & sAppendWhereClause
            End If
    End Select

    '--log sql
    If mbLogSql Then
        LogSql sLogSqlDesc, sSql, pbCreatLogFile
    End If
    
    If goDBConn.UseSQLServer Then
        sSql = Replace(sSql, "||", "+")
        sSql = Replace(sSql, "SUBSTR", "SUBSTRING")
        sSql = Replace(sSql, "CHR(", "CHAR(")
        sSql = Replace(sSql, "TRIM(", "LTRIM(")
    End If
    
    '--execute mothed
    If piProcessStep = 6 Then
        Set prRs = goDBDAS.OpenRecordset(sSql)
    Else
        'goDBSqlOps.Execute sSql
        goDBDAS.ExecuteSQL sSql
    End If

'SCR 34169 RWalker 9/12/03 - START
'Insert recs based upon the Periodontal check box
'being checked, even if a Phase was never assigned
    If piProcessStep = 1 Then
        sLogSqlDesc = "Step 1b. Perio Exam Eval Checked but no Phase Assigned"
        '-- Insert
        sSql = "INSERT INTO Z_PERIO_PHASE_DRC "
        '--select
        sSql = sSql & vbCrLf & _
            "SELECT DISTINCT V.PATIENT_UNIT_ID, V.DENTAL_ENCOUNTER_ID, '',SUBSTR(M.VALUE_TEXT, 1, 4) AS DRC," & vbCrLf & _
           " '" & RQ(psModifiedBy) & "'," & psProcessTime
                
        If goDBConn.UseSQLServer Then
            sSql = sSql & " FROM APPOINTMENT A INNER JOIN" _
                & " MMI_GENERICDATA M ON A.PATIENT_UNIT_NUMBER = M.UNIT_NUMBER RIGHT OUTER JOIN" _
                & " V_PERIO_CHECKED V ON A.PATIENT_UNIT_NUMBER = V.PATIENT_UNIT_ID" _
                & " Where (A.Clinician_NCID = " & sProviderNCID & ") And (M.TYPE_NCID = 14501725)" _
                & ""
        Else
            '-- from
            sSql = sSql & vbCrLf & _
                "FROM APPOINTMENT A," & vbCrLf & _
                vbTab & "V_PERIO_CHECKED V," & vbCrLf & _
                vbTab & "MMI_GENERICDATA M"
            '-- where
            sSql = sSql & vbCrLf & _
                "WHERE A.PATIENT_UNIT_NUMBER (+) = V.PATIENT_UNIT_ID" & vbCrLf & _
                " AND   A.CLINICIAN_NCID = " & sProviderNCID & _
                " AND   M.UNIT_NUMBER = A.PATIENT_UNIT_NUMBER" & vbCrLf & _
                " AND   M.TYPE_NCID = 14501725"
        End If
                '--order by
                sSql = sSql & vbCrLf & _
                    " ORDER BY V.PATIENT_UNIT_ID"
        
        '--log sql
        If mbLogSql Then
            LogSql sLogSqlDesc, sSql, pbCreatLogFile
        End If
        
        If goDBConn.UseSQLServer Then
            sSql = Replace(sSql, "||", "+")
            sSql = Replace(sSql, "SUBSTR", "SUBSTRING")
        End If
        
        goDBSqlOps.Execute sSql
    End If
'SCR 34169 RWalker 9/12/03 - END
    
    SqlProcess = True
    GoTo CleanUp
    
ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "DTPerioPatientSummary.SqlProcess.", "CHCSII_Reports", vbExclamation
CleanUp:
    Exit Function
    
    Err.Clear
    Resume
    
End Function


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
