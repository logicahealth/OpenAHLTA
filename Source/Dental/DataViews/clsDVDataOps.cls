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
Attribute VB_Name = "clsDVDataOps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
    Private Const glENCOUNTER_TYPE As Long = 100
    Private Const glDXTX As Long = 200
    Private Const glLAB_CASE As Long = 300
    Private Const glWORKLOAD As Long = 400
    Private Const glDXSVG As Long = 500
    Private Const glTXSVG As Long = 600

    Private Const gsMODULE_NAME As String = "basLoad."
    Private Const gsDATE_FORMAT_VB    As String = "mm/dd/yyyy hh:mm:ss AM/PM"
    Private Const gsDATE_FORMAT_ORA  As String = "MM/DD/YYYY HH:MI:SS AM"
    Private Const cLogFileName As String = "C:\Temp\log.sql"
    
    Private mstrModifiedBy As String
    Private mstrAudit As String
    Private mblnLogSQL As Boolean
    
    Private mobjCHCSIIConn As CHCSII_CONN.Conn
    Private mobjDas As ICHCSII_DataAccess.ICHCSII_DAS
    Private mobjSQLOps As ICHCSII_SQL.ISqlOps
    Private mobjShared  As CWShared
    Private mobjCWComm  As Object
    Private mobjPatient As Object
    

Public Property Get SQLOps() As ICHCSII_SQL.ISqlOps
    Set SQLOps = mobjSQLOps
End Property

Private Function GetDxStatusSQL(FieldName As String) As String
    
'Returns the SQL to lookup the status code text

    If mobjCHCSIIConn.UseSQLServer Then
        'SQL Server
        GetDxStatusSQL = "CASE " & FieldName _
            & " when 10 then 'New'" _
            & " when 11 then 'Ongoing'" _
            & " when 12 then 'Provisional'" _
            & " when 13 then 'Finding'" _
            & " when 14 then 'Observe'" _
            & " when 20 then 'Resolved'" _
            & " when 21 then 'Deleted'" _
            & " when 22 then 'Replaced'" _
            & " when 23 then 'Confirmed Negative'" _
            & " when 24 then 'Unable to Confirm'" _
            & " else 'Unknown' END"
    Else
        'Oracle
        GetDxStatusSQL = "DECODE(" & FieldName & ",10,'New',11,'Ongoing',12,'Provisonal',13,'Finding',14,'Observe'," & vbCrLf & _
            vbTab & "20,'Resolved',21,'Deleted',22,'Replaced', 23, 'Confirmed Negative', 24, 'Unable to Confirm', 'Unknown')"
    End If
    
End Function

Private Function GetEncounterStatusSQL(FieldName As String) As String

    If mobjCHCSIIConn.UseSQLServer Then
        'SQL Server format
        GetEncounterStatusSQL = "Case " & FieldName _
            & " when 1 then 'CheckedIn'" _
            & " when 2 then 'Waiting'" _
            & " when 3 then 'InProgress'" _
            & " when 4 then 'Completed'" _
            & " when 5 then 'NeedsCosignature'" _
            & " when 6 then 'Update'" _
            & " else 'Unknown' END"
    Else
        'Oracle format
        GetEncounterStatusSQL = "DECODE(" & FieldName & ", 1, 'CheckedIn', 2, 'Waiting', 3,'InProgress', 4,'Completed', " & vbCrLf & _
        vbTab & "5, 'NeedsCosignature', 6, 'Update', 'Unknown')"
    End If
    
End Function

Private Function GetExamEncounterTypeSQL(FieldName As String) As String
    
    If mobjCHCSIIConn.UseSQLServer Then
        'SQL Server
        GetExamEncounterTypeSQL = "Case " & FieldName _
            & " when 1 then ' Periodic.'" _
            & " when 2 then ' Initial Charting.'" _
            & " when 3 then ' Comprehensive.' END"
    Else
        'Oracle
        GetExamEncounterTypeSQL = "DECODE(" & FieldName & ",1,' Periodic.',2,' Initial Charting.',3,' Comprehensive.')"
    End If
    
End Function

Private Function GetTxStatusSQL(FieldName As String) As String

    If mobjCHCSIIConn.UseSQLServer Then
        'SQL Server format
        GetTxStatusSQL = "Case " & FieldName _
            & " when 100 then 'Planned'" _
            & " when 101 then 'InProgress'" _
            & " when 102 then 'Completed'" _
            & " when 103 then 'PreExisting'" _
            & " when 104 then 'Deleted'" _
            & " when 105 then 'Revised'" _
            & " else 'Unknown' END"
    Else
        'Oracle format
        GetTxStatusSQL = "DECODE(" & FieldName & ",100,'Planned',101,'InProgress',102,'Completed',103, " & vbCrLf & _
        vbTab & "'PreExisting',104,'Deleted', 105, 'Revised', 'Unknown')"
    End If
    
End Function
        
Private Sub Class_Initialize()
        
    '11/24/01. application not work if codes are here
    '--Determin need log sql or not
    
    If mobjCHCSIIConn Is Nothing Then
        Set mobjCHCSIIConn = New CHCSII_CONN.Conn
    End If
    
    Set mobjDas = mobjCHCSIIConn.CHCSII_DAS

End Sub
Public Sub Initialize(ByVal strModifiedBy As String, objCWShared As CWShared, _
                       objCWComm As Object, objDas As ICHCSII_DAS, _
                       objSQLOps As ICHCSII_SQL.ISqlOps, ByVal blnLogSQL As Boolean, _
                       objPatient As Object, Optional ByVal blnAudit As Boolean = False)
                       
    Set mobjShared = objCWShared
    Set mobjCWComm = objCWComm

    Set mobjDas = objDas
    Set mobjSQLOps = objSQLOps

    Set mobjPatient = objPatient
    mstrModifiedBy = strModifiedBy

    '--Using AUDIT table or not
    If blnAudit = True Then
        mstrAudit = "_AUDIT"
    Else
        mstrAudit = ""
    End If

End Sub

'********************************************************************************
' Name:         LoadEncounterTypeSql
' Description:  Retrieving Encounter type data through DB tables
'               (TEMPL_EXAMSOAP, TEMPL_EXAMGENERAL, TEMPL_EXAMSOAP, DENTAL_ENCOUNTER_TX )
' Parameter:    plUnitNumber - Patient UnitNumber **Do not use it
' Return:       True - Sucess, False - Fail
' Called By:    DataViews.Mtd_bDentalTabChange()
' Histary:
'   Developer:      Date:       Description
'   ----------      ---------   -----------------------------------------------
'   Wilson sun      01/07/2002  Created. Clone from LoadLabCaseSql
'********************************************************************************
Public Function LoadEncounterTypeSql() As ADODB.Recordset

    Const sPROC_NAME As String = "LoadEncounterTypeSql"
    Dim dtStart As Date
    
    Dim sLogSqlDesc As String
    Dim sProcessTime As String
    Dim sSQL As String
    Dim objRSNotes As ADODB.Recordset
    Dim vDateFrom As Variant
    Dim vDateTo As Variant
    Dim lngMousePointer As VBRUN.MousePointerConstants
    
    On Error GoTo ErrHandler

'    LoadEncounterTypeSql = False
          
    '--Login UserName
    dtStart = Now
    '--TimeStamp *Important for application concurrency.
    '--convert parameter date value to oracle date format
    sProcessTime = mobjCHCSIIConn.FormatDateForDBInsert(dtStart) 'this variable will also be used for delete condition
    'sProcessTime = OracleDate(dtStart) 'this variable will also be used for delete condition
   
    Screen.MousePointer = vbHourglass
    mobjCWComm.setStatusText "Loading Encounter Type Data ..."

    '**********************
    '--Start Process
    '**********************

    '----------------------------------------------------------
    '(I). Step glENCOUNTER_TYPE (100). Retrieve Encounter Type Data
    '----------------------------------------------------------
    If Not SqlProcess(glENCOUNTER_TYPE, "", mstrModifiedBy, sProcessTime, , True, prRs:=objRSNotes) Then GoTo CleanUp

    '
    '-- Process Completed
    '
    '--LogSql
    If mblnLogSQL Then
        LogSql "Encounter Type Process Complete On: " & Format(Now), "--> Total Process time --> " & CStr(DateDiff("s", dtStart, Now)) & " Sec."
    End If
    

    If Not objRSNotes Is Nothing Then
        FilterNotesRS objRSNotes
        Set LoadEncounterTypeSql = objRSNotes
    End If
    
    GoTo CleanUp

ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description, gsMODULE_NAME & sPROC_NAME, App.Title, vbCritical)
CleanUp:
    mobjCWComm.setStatusText ""

    Screen.MousePointer = lngMousePointer
End Function


'********************************************************************************
' Name:         LoadWorkloadSql
' Description:  Retrieving work load data through DB tables (DENTAL_ENCOUNTER_TX, TEMPL_WORKLOAD)
' Parameter:    plUnitNumber - Patient UnitNumber **Do not use it
' Return:       True - Sucess, False - Fail
' Called By:    DataViews.Mtd_bDentalTabChange()
' Histary:
'   Developer:      Date:       Description
'   ----------      ---------   -----------------------------------------------
'   Wilson sun      11/21/2001  Created. Clone from LoadLabCaseSql
'********************************************************************************
Public Function LoadWorkloadSql(ByVal plUnitNumber As Long) As ADODB.Recordset

    Const sPROC_NAME As String = "LoadWorkloadSql"
    Dim dtStart As Date
    
    Dim sLogSqlDesc As String
    Dim sProcessTime As String
    Dim sSQL As String
    Dim objRSWorkLoad As ADODB.Recordset
    
    Dim vDateFrom As Variant
    Dim vDateTo As Variant
    Dim lngMousePointer As VBRUN.MousePointerConstants
          
    On Error GoTo ErrHandler

          
    '--Login UserName
    dtStart = Now
    '--TimeStamp *Important for application concurrency.
    '--convert parameter date value to oracle date format
    sProcessTime = mobjCHCSIIConn.FormatDateForDBInsert(dtStart) 'this variable will also be used for delete condition
    'sProcessTime = OracleDate(dtStart) 'this variable will also be used for delete condition
   
    Screen.MousePointer = vbHourglass
    mobjCWComm.setStatusText "Loading Workload Data ..."

    '**********************
    '--Start Process
    '**********************

    '----------------------------------------------------------
    '(I). Step glWORKLOAD (400). Retrieve Lab Case Data
    '----------------------------------------------------------
    If Not SqlProcess(glWORKLOAD, "", mstrModifiedBy, sProcessTime, plUnitNumber, False, prRs:=objRSWorkLoad) Then GoTo CleanUp
    '
    '-- Process Completed
    '
    '--LogSql
    If mblnLogSQL Then
        LogSql "Workload Process Complete On: " & Format(Now), "--> Total Process time --> " & CStr(DateDiff("s", dtStart, Now)) & " Sec."
    End If
    
    Set LoadWorkloadSql = objRSWorkLoad

    GoTo CleanUp

ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description, gsMODULE_NAME & sPROC_NAME, App.Title, vbCritical)
CleanUp:
    mobjCWComm.setStatusText ""

    Screen.MousePointer = lngMousePointer
End Function


'********************************************************************************
' Name:         LoadLabCaseSql
' Description:  Retrieving DxTx data through DB tables
' Parameter:    plUnitNumber - Patient UnitNumber
' Return:       True - Sucess, False - Fail
' Called By:    DataViews.Mtd_bDentalTabChange()
' Histary:
'   Developer:      Date:       Description
'   ----------      ---------   -----------------------------------------------
'   Wilson sun      10/31/2001  Created. Clone from LoadDxTxDataSql
'********************************************************************************
Public Function LoadLabCaseSql(ByVal plUnitNumber As Long) As ADODB.Recordset

    Const sPROC_NAME As String = "LoadLabCaseSql"
    Dim dtStart As Date
    
    Dim sLogSqlDesc As String
    Dim sProcessTime As String
    Dim sSQL As String
    Dim objRSLab As ADODB.Recordset
    Dim vDateFrom As Variant
    Dim vDateTo As Variant
    Dim lngMousePointer As VBRUN.MousePointerConstants
        
    On Error GoTo ErrHandler
          
    '--Login UserName
    dtStart = Now
    '--TimeStamp *Important for application concurrency.
    '--convert parameter date value to oracle date format
    sProcessTime = mobjCHCSIIConn.FormatDateForDBInsert(dtStart) 'this variable will also be used for delete condition
    'sProcessTime = OracleDate(dtStart) 'this variable will also be used for delete condition
   
    Screen.MousePointer = vbHourglass
    mobjCWComm.setStatusText "Loading Labs Data ..."

    '----------------------------------------------------------
    '(I). Step glLAB_CASE (300). Retrieve Lab Case Data
    '----------------------------------------------------------
    If Not SqlProcess(glLAB_CASE, "", mstrModifiedBy, sProcessTime, plUnitNumber, False, prRs:=objRSLab) Then GoTo CleanUp
    '
    '-- Process Completed
    '
    '--LogSql
    If mblnLogSQL Then
        LogSql "Lab Case Process Complete On: " & Format(Now), "--> Total Process time --> " & CStr(DateDiff("s", dtStart, Now)) & " Sec."
    End If
    
    Set LoadLabCaseSql = objRSLab

    GoTo CleanUp

ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description, gsMODULE_NAME & sPROC_NAME, App.Title, vbCritical)
CleanUp:
    mobjCWComm.setStatusText ""
    Set mobjSQLOps = Nothing
'    Set mobjDas = Nothing

    Screen.MousePointer = lngMousePointer
End Function

'********************************************************************************
' Name:         LoadDxTxDataSql
' Description:  Retrieving DxTx data through DB tables
' Parameter:    plUnitNumber - Patient UnitNumber
' Return:       True - Sucess, False - Fail
' Called By:    DataViews.Mtd_bDentalTabChange()
' Histary:
'   Developer:      Date:       Description
'   ----------      ---------   -----------------------------------------------
'   Wilson Sun      12/18/01    Changed Called By this function from vsIndexTab1_Click()
'   Wilson Sun      11/24/01    Move mobjDas , goSqlOps, mblnLogSQL, gbAudit setting to Classs_initailize
'   Wilson sun      05/30/2001  Created
'********************************************************************************
Public Function LoadDxTxDataSql(ByVal plUnitNumber As Long) As ADODB.Recordset

    Const sPROC_NAME As String = "LoadDxTxDataSql"
    Dim sProcessTime As String
    Dim objRSViews As ADODB.Recordset
    Dim sSQL As String
    Dim sLogSqlDesc As String
    Dim vDateFrom As Variant
    Dim vDateTo As Variant
    Dim lngMousePointer As VBRUN.MousePointerConstants
    
    'Wilson 5/30/01
    Dim dtStart As Date
    
    On Error GoTo ErrHandler
         
    '--Login UserName
    dtStart = Now
    '--TimeStamp *Important for application concurrency.
    '--convert parameter date value to oracle date format
    sProcessTime = mobjCHCSIIConn.FormatDateForDBInsert(dtStart) 'this variable will also be used for delete condition
    'sProcessTime = OracleDate(dtStart) 'this variable will also be used for delete condition
   
    Screen.MousePointer = vbHourglass
    mobjCWComm.setStatusText "Loading Dx/Tx Data ..."
    
    '----------------------------------------------------------
    '(I). Step 1 Retrieve Dx Data
    '----------------------------------------------------------
    If Not SqlProcess(glDXTX + 1, mstrAudit, mstrModifiedBy, sProcessTime, plUnitNumber) Then GoTo CleanUp
    
    '----------------------------------------------------------
    '(I). Step 2 Retrieve Dx Appointment and Provider Info
    '----------------------------------------------------------
    If Not SqlProcess(glDXTX + 2, mstrAudit, mstrModifiedBy, sProcessTime) Then GoTo CleanUp
    
    '----------------------------------------------------------
    '(I). Step 3 Retrieve Dx Chartable Data
    '----------------------------------------------------------
    If Not SqlProcess(glDXTX + 3, mstrAudit, mstrModifiedBy, sProcessTime) Then GoTo CleanUp
   
    '----------------------------------------------------------
    '(I). Step 4 Retrieve Dx No Chartable Data
    '----------------------------------------------------------
    If Not SqlProcess(glDXTX + 4, mstrAudit, mstrModifiedBy, sProcessTime) Then GoTo CleanUp
   
    '==========================================================
    '(II). Step 5 Retrieve Tx Data
    '----------------------------------------------------------
    If Not SqlProcess(glDXTX + 5, mstrAudit, mstrModifiedBy, sProcessTime, plUnitNumber) Then GoTo CleanUp

    '----------------------------------------------------------
    '(II). Step 6 Retrieve Tx Appointment and Provider Info
    '----------------------------------------------------------
    If Not SqlProcess(glDXTX + 6, mstrAudit, mstrModifiedBy, sProcessTime) Then GoTo CleanUp
   
    '----------------------------------------------------------
    '(II). Step 7 Retrieve Tx Chartable Data
    '----------------------------------------------------------
    If Not SqlProcess(glDXTX + 7, mstrAudit, mstrModifiedBy, sProcessTime) Then GoTo CleanUp
   
    '----------------------------------------------------------
    '(II). Step 8 Retrieve Tx No Chartable Data
    '----------------------------------------------------------
    If Not SqlProcess(glDXTX + 8, mstrAudit, mstrModifiedBy, sProcessTime) Then GoTo CleanUp
   
    '==========================================================
    '(III). Step 9 Retrieve DxTx Data
    'Wilson commented on 12/21/01 because using GROUP_ID as join
    '----------------------------------------------------------
'    If Not SqlProcess(glDXTX + 9, mstrAudit, mstrModifiedBy, sProcessTime) Then GoTo CleanUp
   
    '----------------------------------------------------------
    '(III). Step 10 Record set for displaying in Grid
    '----------------------------------------------------------
    If Not SqlProcess(glDXTX + 10, mstrAudit, mstrModifiedBy, sProcessTime, prRs:=objRSViews) Then GoTo CleanUp
    
    '==========================================================
    'Delete Work tables
    '
    '(IV) Step 11. delete Z_DATAVIEW_DX_1 table
    '----------------------------------------------------------
    If Not SqlProcess(glDXTX + 11, mstrAudit, mstrModifiedBy, sProcessTime) Then GoTo CleanUp
    
    '----------------------------------------------------------
    '(IV) Step 12. delete Z_DATAVIEW_DX_2 table
    '----------------------------------------------------------
    If Not SqlProcess(glDXTX + 12, mstrAudit, mstrModifiedBy, sProcessTime) Then GoTo CleanUp
    
    '----------------------------------------------------------
    '(IV) Step 13. delete Z_DATAVIEW_DX_3 table
    '----------------------------------------------------------
    If Not SqlProcess(glDXTX + 13, mstrAudit, mstrModifiedBy, sProcessTime) Then GoTo CleanUp
    
    '----------------------------------------------------------
    '(IV) Step 14. delete Z_DATAVIEW_TX_1 table
    '----------------------------------------------------------
    If Not SqlProcess(glDXTX + 14, mstrAudit, mstrModifiedBy, sProcessTime) Then GoTo CleanUp

    '----------------------------------------------------------
    '(IV) Step 15. delete Z_DATAVIEW_TX_2 table
    '----------------------------------------------------------
    If Not SqlProcess(glDXTX + 15, mstrAudit, mstrModifiedBy, sProcessTime) Then GoTo CleanUp
    
    '----------------------------------------------------------
    '(IV) Step 16. delete Z_DATAVIEW_TX_3 table
    '----------------------------------------------------------
    If Not SqlProcess(glDXTX + 16, mstrAudit, mstrModifiedBy, sProcessTime) Then GoTo CleanUp
    
    '----------------------------------------------------------
    '(IV) Step 17. delete Z_DATAVIEW_DXTX table
    'Comment it out since commented out Step 9
    '----------------------------------------------------------
'    If Not SqlProcess(glDXTX + 17, mstrAudit, mstrModifiedBy, sProcessTime) Then GoTo CleanUp
   
    '
    '-- Process Completed
    '
    '--LogSql
    If mblnLogSQL Then
        LogSql "DxTx Process Complete On: " & Format(Now), "--> Total Process time --> " & CStr(DateDiff("s", dtStart, Now)) & " Sec."
    End If
    
    Set LoadDxTxDataSql = objRSViews

    GoTo CleanUp

ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description, gsMODULE_NAME & sPROC_NAME, App.Title, vbCritical)
CleanUp:
    mobjCWComm.setStatusText ""

    Screen.MousePointer = lngMousePointer
End Function
'**********************************************************************************
'Name:          LogSql
'Description:   Log Sql script to log.sql file
'Parameters:    1. psDesc   - Description of the sql
'               2. psSql    - Sql statement
'               3. pbCreate - True: Create log file, False: Appending to file
'               4. plProcessType - Encounter, DxTx, Labcase, Workload
'Called by:     Any access Sql function, procedure
'History:
'Date       Developer       Description
'========== ============    =============================================
'01/08/2002 Wilson Sun      Added Head of start process type
'02/12/2001 Wilson Sun      Created
'**********************************************************************************
Private Sub LogSql(ByVal psDesc As String, _
            ByVal psSql As String, _
            Optional ByVal pbCreate As Boolean = False, _
            Optional ByVal plProcessType As Long)

    Dim fs As New FileSystemObject
    Dim f As TextStream
    Dim sHead As String
    Dim sStar As String
    
    sStar = String(80, "*") 'For head
    psDesc = String(80, "-") & vbCrLf & _
            "--" & psDesc & vbCrLf & _
            String(80, "-")
    Select Case plProcessType
        Case glENCOUNTER_TYPE
            sHead = vbCrLf & sStar & vbCrLf & _
                "--Log Encounter Type Start On: " & Format(Now) & vbCrLf & _
                sStar
        Case glDXTX + 1
            sHead = vbCrLf & sStar & vbCrLf & _
                "--Log DXTX Start On: " & Format(Now) & vbCrLf & _
                sStar
        Case glLAB_CASE
            sHead = vbCrLf & sStar & vbCrLf & _
                "--Log Lab Case Start On: " & Format(Now) & vbCrLf & _
                sStar
        Case glWORKLOAD
            sHead = vbCrLf & sStar & vbCrLf & _
                "--Log Workload Start On: " & Format(Now) & vbCrLf & _
                sStar
        Case glDXSVG, glTXSVG
            sHead = vbCrLf & sStar & vbCrLf & _
                "--Log DX/TX SVG Start On: " & Format(Now) & vbCrLf & _
                sStar
'         Case glTXSVG
'            sHead = vbCrLf & sStar & vbCrLf & _
'                "--Log TX SVG Start On: " & Format(Now) & vbCrLf & _
'                sStar
        Case Else
            sHead = ""
    End Select
    
    If pbCreate Then
        Set f = fs.OpenTextFile(cLogFileName, ForWriting, True)
    Else
        Set f = fs.OpenTextFile(cLogFileName, ForAppending, False)
    End If
       
    f.Write sHead & vbCrLf
    f.Write psDesc & vbCrLf
    f.Write psSql & vbCrLf & ";"
    f.Close
    Set f = Nothing
    Set fs = Nothing
End Sub


'*********************************************************************************************************
'Name:          SqlProcess
'Description:   This is a main container of Sql statments for the whole process
'Parameters:
'   piProcessStep   - Process Step
'   psModifiedBy    - Login User Name
'   psProcessTime   - When you click Report button time
'   prRs            - Record set for the report
'Return Value:
'   True - sucessful,
    'False - failed
'Called by:     LoadDxTxDataSql, LoadLabCaseSql, LoadWorkloadSql, LoadEncounterTypeSql
'History:
' Date      Developer   Description
' --------  ----------- -----------------------------------------------------------------------------
'06/27/2002 KDunne      Providers with apostrophe's in their name (Erin O'Donohue) were generating
'                       VB errors. Added RQ function to all SQL queries.
' 01/07/02  Wilson Sun  Added Step glENCOUNTERTYPE (100) for load Encounter Type (SCR 8452)
' 12/24/01  Wilson Sun  CREATE VIEW V_DX_SUB_LOCATION as
'                       SELECT T1.SUB_LOCATION_ID, T1.SUB_LOCATION_NAME, T2.DX_SUBLOC_ID, T2.DX_LOCATION_ID, T2.DELETED
'                       FROM NOTE_SUB_LOCATION T1, NOTE_DX_SUB_LOCATION T2
'                       Where T1.SUB_LOCATION_ID = T2.SUB_LOCATION_ID
'
'                       CREATE VIEW V_TX_SUB_LOCATION as
'                       SELECT T1.SUB_LOCATION_ID, T1.SUB_LOCATION_NAME, T2.TX_SUBLOC_ID, T2.TX_LOCATION_ID, T2.DELETED
'                       FROM NOTE_SUB_LOCATION T1, NOTE_TX_SUB_LOCATION T2
'                       Where T1.SUB_LOCATION_ID = T2.SUB_LOCATION_ID
' 12/24/01              Changed Step 4 and 8 for retrieving no chartable data from V_DX_SUB_LOCATION and V_TX_SUB_LOCATION
' 11/21/01  Wilson Sun  Added Step glWORKLOAD (400) for load workload data (SCR 17422)
' 10/31/01  Wilson Sun  Added Step glLAB_CASE (300) for load lab case data
' 09/05/01  Wilson Sun  Modified Step 10. (SCR: 8474)
'                       Setup default column order (Move TX_CODE after Treatment)
' 08/15/01  Wilson Sun  Modified Step 9.
'                       Remove DENTAL_ENCOUNTER_TXS_DXS table and add DENTAL_ENCOUNTER_GROUPGS table
' 08/15/01  Wilson Sun  Modified Step 1.
'                       Select Clause from E.DENTAL_ENCOUNTER_ID to EDX.MODIFIED_IN_DENTAL_ENCOUNTER
'                       Where Clause from EDX.DENTAL_ENCOUNTER_ID to EDX.MODIFIED_IN_DENTAL_ENCOUNTER
' 08/15/01  Wilson Sun  Modified Step 5.
'                       Select Clause from E.DENTAL_ENCOUNTER_ID to ETX.MODIFIED_IN_DENTAL_ENCOUNTER
'                       Where Clause from ETX.DENTAL_ENCOUNTER_ID to ETX.MODIFIED_IN_DENTAL_ENCOUNTER
' 08/08/01  Wilson Sun  Modified goDas and goSqlOps declaration for theater project
' 08/06/01  Wilson Sun  Took DENTAL and CDRPLUS schma away from SQL statements
' 08/03/01  Wilson Sun  Changed Step 5 from TXD.ADA_CATEGORY_CODE to TXD.ADA_CODE in Select statement
' 05/31/01  Wilson Sun  Created
'*********************************************************************************************************
Private Function SqlProcess(ByVal piProcessStep As Long, _
                            ByVal pgsAudit As String, _
                            ByVal psModifiedBy As String, _
                            ByVal psProcessTime As String, _
                            Optional ByVal plPatientUnitNumber As Long, _
                            Optional ByVal pbCreateLogFile As Boolean = False, _
                            Optional ByRef prRs As ADODB.Recordset) As Boolean
   
    '--String
    Dim sAppendWhereClause As String
    Dim sAppendWhereClause2 As String

    Dim sLogSqlDesc As String
    Dim sSQL As String
       
    On Error GoTo ErrHandler
    
    SqlProcess = False
        
    sAppendWhereClause = "Z.MODIFIED_BY = '" & RQ(psModifiedBy) & "'" & vbCrLf & _
       "and   Z.TIME_STAMP = " & psProcessTime
    
    sAppendWhereClause2 = "Z2.MODIFIED_BY = '" & RQ(psModifiedBy) & "'" & vbCrLf & _
       "and   Z2.TIME_STAMP = " & psProcessTime
    
    Select Case piProcessStep
        '===================================================================================================
        '(Problems/Planned Tx and Current Oral State Tab)
        'Step glENCOUNTER_TYPE (300). Retrieve Encounter Type Data
        '===================================================================================================
        Case glENCOUNTER_TYPE
            sLogSqlDesc = ">Encounter Type Process<-- Step " & CStr(glENCOUNTER_TYPE) & ". Retrieve Encounter Type Data"
            sSQL = ProcessEncounterType
        
        '===================================================================================================
        ' DxTx Data
        '===================================================================================================
        '--------------------------------------------------------------------------------
        '(I) Step 1. Retrieve Dx Data
        'Target Table:  Z_DATAVIEW_DX_1
        'Source Table:  DENTAL_ENCOUNTER, DENTAL_ENCOUNTER_DX, DX_ICD_CODE_DEFINITION
        'Note:          Exclode UNSIGNED and DELETED status records
        '--------------------------------------------------------------------------------
        Case glDXTX + 1
            'pbCreateLogFile = True
            sLogSqlDesc = ">DxTx Process<-- (I) Step " & CStr(glDXTX + 1) & ". Retrieve Dx Data"

            '-- Insert
            sSQL = "insert into Z_DATAVIEW_DX_1 (" & vbCrLf & _
                vbTab & "GROUP_ID, ENCOUNTER_NUMBER,DENTAL_ENCOUNTER_ID,DX_CODE_ID,STATUS_ID,STATUS," & vbCrLf & _
                vbTab & "CHARTABLE,DRC,SIGNED,CREATE_DATE_TIME,ICD_CODE,ICD_DEFINITION," & vbCrLf & _
                vbTab & "PROVIDER_ID,ENCOUNTER_FACILITY_NCID,PATIENT_UNIT_ID,MODIFIED_BY,TIME_STAMP)"
            '-- select
            sSQL = sSQL & vbCrLf & _
                "select G.GROUP_ID, E.ENCOUNTER_NUMBER, EDX.MODIFIED_IN_DENTAL_ENCOUNTER, EDX.DX_CODE_ID, EDX.STATUS_ID, " & vbCrLf & _
                vbTab & GetDxStatusSQL("EDX.STATUS_ID") & "," & _
                vbTab & "EDX.CHARTABLE, EDX.DRC, EDX.SIGNED, EDX.CREATE_DATE_TIME, DXD.ICD_CODE, DXD.ICD_DEFINITION," & vbCrLf & _
                vbTab & "EDX.PROVIDER_ID, E.ENCOUNTER_FACILITY_NCID, E.PATIENT_UNIT_ID, " & vbCrLf & _
                vbTab & "'" & RQ(psModifiedBy) & "', " & psProcessTime
            '-- from
            sSQL = sSQL & vbCrLf & _
                "from DENTAL_ENCOUNTER" & pgsAudit & " E," & vbCrLf & _
                vbTab & "DENTAL_ENCOUNTER_DX" & pgsAudit & " EDX," & vbCrLf & _
                vbTab & "DENTAL_ENCOUNTER_GROUPS " & pgsAudit & " G, " & vbCrLf & _
                vbTab & "DX_ICD_CODE_DEFINITION DXD"
            '-- where
            sSQL = sSQL & vbCrLf & _
                "where E.PATIENT_UNIT_ID = " & plPatientUnitNumber & vbCrLf & _
                "and   E.DENTAL_ENCOUNTER_ID = EDX.MODIFIED_IN_DENTAL_ENCOUNTER " & vbCrLf & _
                "and   EDX.DX_CODE_ID = G.DX_CODE_ID " & vbCrLf & _
                "and   EDX.DX_ICD_CODE_ID = DXD.DX_ICD_CODE_ID " & vbCrLf
'                "and  NOT (EDX.STATUS_ID = 21 and EDX.SIGNED = 0) "
                'KDunne 6/24/2002
                'SCR 17451 - SA note says that now we're only displaying DX's that are New, OnGoing or provisional

        '--------------------------------------------------------------------------------
        '(I) Step 2. Retrieve Dx Appointment and Provider Info
        'Target Table: Z_DATAVIEW_DX_2
        'Source Table: Z_DATAVIEW_DX_1, APPOINTMENT, CLINIC_PROVIDER
        '--------------------------------------------------------------------------------
        Case glDXTX + 2
            sLogSqlDesc = "(I) Step " & CStr(glDXTX + 2) & ". Retrieve Dx Appointment and Provider Info"
            '-- Insert
            sSQL = "insert into Z_DATAVIEW_DX_2 (" & vbCrLf & _
                vbTab & "GROUP_ID, ENCOUNTER_NUMBER,DENTAL_ENCOUNTER_ID,ENCOUNTER_STATUS_ID,ENCOUNTER_STATUS," & vbCrLf & _
                vbTab & "APPT_DATE_TIME,DX_CODE_ID,STATUS_ID,STATUS,CHARTABLE,DRC,SIGNED," & vbCrLf & _
                vbTab & "CREATE_DATE_TIME,ICD_CODE,ICD_DEFINITION,PROVIDER_NAME," & vbCrLf & _
                vbTab & "PATIENT_UNIT_ID,MODIFIED_BY,TIME_STAMP)"
            '-- select
            sSQL = sSQL & vbCrLf & _
                "select Z.GROUP_ID, Z.ENCOUNTER_NUMBER, Z.DENTAL_ENCOUNTER_ID, A.ENCOUNTER_STATUS, " & vbCrLf & _
                vbTab & GetEncounterStatusSQL("A.ENCOUNTER_STATUS") & ", A.APPT_DATE_TIME, " & vbCrLf & _
                vbTab & "Z.DX_CODE_ID,  Z.STATUS_ID, Z.STATUS, Z.CHARTABLE, Z.DRC, Z.SIGNED, " & vbCrLf & _
                vbTab & "Z.CREATE_DATE_TIME , Z.ICD_CODE, Z.ICD_DEFINITION, CP.PROVIDER_NAME, Z.PATIENT_UNIT_ID, Z.MODIFIED_BY, Z.TIME_STAMP"
            '-- from
            sSQL = sSQL & vbCrLf & _
                "from Z_DATAVIEW_DX_1 Z, " & vbCrLf & _
                vbTab & "APPOINTMENT A, " & vbCrLf & _
                vbTab & "CLINIC_PROVIDER CP "
            '-- where
            sSQL = sSQL & vbCrLf & _
                "where Z.ENCOUNTER_FACILITY_NCID = A.FACILITY_NCID " & vbCrLf & _
                "and   Z.PATIENT_UNIT_ID = A.PATIENT_UNIT_NUMBER " & vbCrLf & _
                "and   Z.ENCOUNTER_NUMBER = A.ENCOUNTER_NUMBER " & vbCrLf & _
                "and   A.CLINIC_NCID = CP.CLINIC_NCID " & vbCrLf & _
                "and   Z.PROVIDER_ID = CP.CLINICIAN_NCID " & vbCrLf & _
                "and   " & sAppendWhereClause
        
        '--------------------------------------------------------------------------------
        '(I) Step 3. Retrieve Dx Chartable Data
        'Target Table Z_DATAVIEW_DX_3
        'Source Table: Z_DATAVIEW_DX_2, DENTAL_ENCOUNTER_DXS_SS, DENTAL_ENCOUNTER_TEETH
        '--------------------------------------------------------------------------------
        Case glDXTX + 3
            sLogSqlDesc = "(I) Step " & CStr(glDXTX + 3) & ". Retrieve Dx Chartable Data"
            '-- Insert
            sSQL = "insert into Z_DATAVIEW_DX_3 (" & vbCrLf & _
                vbTab & "GROUP_ID, ENCOUNTER_NUMBER,DENTAL_ENCOUNTER_ID,ENCOUNTER_STATUS_ID,ENCOUNTER_STATUS," & vbCrLf & _
                vbTab & "APPT_DATE_TIME,DX_CODE_ID,STATUS_ID,STATUS,CHARTABLE,DRC,SIGNED,TOOTH_NAME," & vbCrLf & _
                vbTab & "SURFACEABBR,CREATE_DATE_TIME,ICD_CODE,ICD_DEFINITION,PROVIDER_NAME," & vbCrLf & _
                vbTab & "PATIENT_UNIT_ID,MODIFIED_BY,TIME_STAMP)"
            '-- select
            sSQL = sSQL & vbCrLf & _
                "select Z.GROUP_ID, Z.ENCOUNTER_NUMBER, Z.DENTAL_ENCOUNTER_ID, Z.ENCOUNTER_STATUS_ID, " & vbCrLf & _
                vbTab & "Z.ENCOUNTER_STATUS, Z.APPT_DATE_TIME, Z.DX_CODE_ID, Z.STATUS_ID, Z.STATUS, " & vbCrLf & _
                vbTab & "Z.CHARTABLE, Z.DRC,Z.SIGNED, ET.TOOTH_NAME, ESS.SURFACEABBR, Z.CREATE_DATE_TIME, " & vbCrLf & _
                vbTab & "Z.ICD_CODE , Z.ICD_DEFINITION, Z.PROVIDER_NAME, Z.PATIENT_UNIT_ID, Z.MODIFIED_BY, Z.TIME_STAMP"
            '-- from
            sSQL = sSQL & vbCrLf & _
                "from Z_DATAVIEW_DX_2 Z, " & vbCrLf & _
                vbTab & "DENTAL_ENCOUNTER_DXS_SS" & pgsAudit & " EDXS, " & vbCrLf & _
                vbTab & "DENTAL_ENC_SURFACE_SET" & pgsAudit & " ESS, " & vbCrLf & _
                vbTab & "DENTAL_ENCOUNTER_TEETH" & pgsAudit & " ET "
            '-- where
            sSQL = sSQL & vbCrLf & _
                "where Z.DX_CODE_ID = EDXS.DX_CODE_ID " & vbCrLf & _
                "and   ESS.SURFACE_SET_ID = EDXS.SURFACE_SET_ID " & vbCrLf & _
                "and   ET.DENTAL_ENCOUNTER_TEETH_ID = ESS.DENTAL_ENCOUNTER_TEETH_ID " & vbCrLf & _
                "and   Z.CHARTABLE = -1 " & vbCrLf & _
                "and   " & sAppendWhereClause
                
        '--------------------------------------------------------------------------------
        '(I) Step 4. Retrieve Dx No Chartable Data
        'Target Table:  Z_DATAVIEW_DX_3
        'Source Table:  Z_DATAVIEW_DX_2, NOTE_DXS_LOCATION, NOTE_LOCATION,
        '               NOTE_DX_SUB_LOCATION, NOTE_SUB_LOCATION
        '--------------------------------------------------------------------------------
        Case glDXTX + 4
            sLogSqlDesc = "(I) Step " & CStr(glDXTX + 4) & ". Retrieve Dx No Chartable Data"
            '-- Insert
            sSQL = "insert into Z_DATAVIEW_DX_3 (" & vbCrLf & _
                vbTab & "GROUP_ID, ENCOUNTER_NUMBER,DENTAL_ENCOUNTER_ID,ENCOUNTER_STATUS_ID,ENCOUNTER_STATUS," & vbCrLf & _
                vbTab & "APPT_DATE_TIME,DX_CODE_ID,STATUS_ID,STATUS,CHARTABLE,DRC,SIGNED,TOOTH_NAME," & vbCrLf & _
                vbTab & "SURFACEABBR,CREATE_DATE_TIME,ICD_CODE,ICD_DEFINITION,PROVIDER_NAME," & vbCrLf & _
                vbTab & "PATIENT_UNIT_ID,MODIFIED_BY,TIME_STAMP)"
            '-- select
            sSQL = sSQL & vbCrLf & _
                "select Z.GROUP_ID, Z.ENCOUNTER_NUMBER, Z.DENTAL_ENCOUNTER_ID, Z.ENCOUNTER_STATUS_ID, " & vbCrLf & _
                vbTab & "Z.ENCOUNTER_STATUS, Z.APPT_DATE_TIME, Z.DX_CODE_ID, Z.STATUS_ID, " & vbCrLf & _
                vbTab & "Z.STATUS, Z.CHARTABLE, Z.DRC, Z.SIGNED, NL.LOCATION_NAME, " & vbCrLf & _
                vbTab & "V.SUB_LOCATION_NAME, Z.CREATE_DATE_TIME, Z.ICD_CODE, " & vbCrLf & _
                vbTab & "Z.ICD_DEFINITION , Z.PROVIDER_NAME, Z.PATIENT_UNIT_ID, Z.MODIFIED_BY, Z.TIME_STAMP"
            '-- from
            
            If mobjCHCSIIConn.UseSQLServer Then
                'SQL Server format
                '-- from and where
                sSQL = sSQL & " FROM Z_DATAVIEW_DX_2 Z INNER JOIN NOTE_DXS_LOCATION NDXSL ON Z.DX_CODE_ID = NDXSL.DX_CODE_ID INNER JOIN" _
                    & " NOTE_LOCATION NL ON NDXSL.LOCATION_ID = NL.LOCATION_ID LEFT OUTER JOIN" _
                    & " V_DX_SUB_LOCATION V ON NDXSL.DX_LOCATION_ID = V.DX_LOCATION_ID" _
                    & " WHERE (Z.CHARTABLE = 0) and " & sAppendWhereClause
            Else
                'Oracle format
                '-- from
                sSQL = sSQL & vbCrLf & _
                    "from Z_DATAVIEW_DX_2 Z, " & vbCrLf & _
                    vbTab & "NOTE_DXS_LOCATION" & pgsAudit & " NDXSL, " & vbCrLf & _
                    vbTab & "NOTE_LOCATION NL, " & vbCrLf & _
                    vbTab & "V_DX_SUB_LOCATION V "
                
                '-- where
                sSQL = sSQL & vbCrLf & _
                    "where Z.DX_CODE_ID = NDXSL.DX_CODE_ID " & vbCrLf & _
                    "and   NL.LOCATION_ID = NDXSL.LOCATION_ID " & vbCrLf & _
                    "and   NDXSL.DX_LOCATION_ID = V.DX_LOCATION_ID (+) " & vbCrLf & _
                    "and   Z.CHARTABLE = 0 " & vbCrLf & _
                    "and   " & sAppendWhereClause
            End If
            
        '--------------------------------------------------------------------------------
        '(II) Step 5. Retrieve Tx Data
        'Target Table:  Z_DATAVIEW_TX_1
        'Source Table:  DENTAL_ENCOUNTER, DENTAL_ENCOUNTER_TX, TX_ADA_CODE_DEFINITION
        'Note:          Exclode UNSIGNED and DELETED status records
        '--------------------------------------------------------------------------------
        Case glDXTX + 5
            sLogSqlDesc = "(II) Step " & CStr(glDXTX + 5) & ". Retrieve Tx Data"
            '-- Insert
            sSQL = "insert into Z_DATAVIEW_TX_1 (" & vbCrLf & _
                 vbTab & "GROUP_ID, ENCOUNTER_NUMBER,DENTAL_ENCOUNTER_ID,TX_CODE_ID,STATUS_ID,STATUS,CHARTABLE," & vbCrLf & _
                 vbTab & "SIGNED,CREATE_DATE_TIME,ADA_CATEGORY_CODE,ADA_DEFINITION," & vbCrLf & _
                 vbTab & "ENCOUNTER_FACILITY_NCID,PROVIDER_ID,PATIENT_UNIT_ID,MODIFIED_BY,TIME_STAMP)"
            '-- select
            'KDunne 4/23/2002
            'SCR 17455 - Added TX status code of Revised and corrected misspelling of Unknown.
            '
            sSQL = sSQL & vbCrLf & _
                "select G.GROUP_ID, E.ENCOUNTER_NUMBER, ETX.MODIFIED_IN_DENTAL_ENCOUNTER, ETX.TX_CODE_ID, ETX.STATUS_ID, " & vbCrLf & _
                vbTab & GetTxStatusSQL("ETX.STATUS_ID") & ", " & vbCrLf & _
                vbTab & "ETX.CHARTABLE, ETX.SIGNED, ETX.CREATE_DATE_TIME, TXD.ADA_CODE, " & vbCrLf & _
                vbTab & "TXD.ADA_DEFINITION, E.ENCOUNTER_FACILITY_NCID , ETX.PROVIDER_ID, E.PATIENT_UNIT_ID, " & vbCrLf & _
                vbTab & "'" & RQ(psModifiedBy) & "', " & psProcessTime
            '-- from
            sSQL = sSQL & vbCrLf & _
                "from DENTAL_ENCOUNTER" & pgsAudit & " E, " & vbCrLf & _
                vbTab & "DENTAL_ENCOUNTER_TX" & pgsAudit & " ETX, " & vbCrLf & _
                vbTab & "DENTAL_ENCOUNTER_GROUPS " & pgsAudit & " G, " & vbCrLf & _
                vbTab & "TX_ADA_CODE_DEFINITION TXD"
            '-- where
            sSQL = sSQL & vbCrLf & _
                "where E.PATIENT_UNIT_ID = " & plPatientUnitNumber & vbCrLf & _
                "and   E.DENTAL_ENCOUNTER_ID = ETX.MODIFIED_IN_DENTAL_ENCOUNTER " & vbCrLf & _
                "and   ETX.TX_CODE_ID = G.TX_CODE_ID " & vbCrLf & _
                "and   ETX.TX_ADA_CODE_ID = TXD.TX_ADA_CODE_ID " '& vbCrLf & _
                '"and   not (ETX.STATUS_ID = 104 and ETX.SIGNED = 0) "
                'WilsonNote: uncomment above if needed
        '--------------------------------------------------------------------------------
        '(II) Step 6. Retrieve Tx Appointment and Provider Info
        'Target Table Z_DATAVIEW_TX_2
        'Source Table: Z_DATAVIEW_TX_1, APPOINTMENT, CLINIC_PROVIDER
        '--------------------------------------------------------------------------------
        Case glDXTX + 6
            sLogSqlDesc = "(II) Step " & CStr(glDXTX + 6) & ". Retrieve Tx Appointment and Provider Info"
            '-- Insert
            sSQL = "insert into Z_DATAVIEW_TX_2 (" & vbCrLf & _
                "GROUP_ID, ENCOUNTER_NUMBER,DENTAL_ENCOUNTER_ID,ENCOUNTER_STATUS_ID,ENCOUNTER_STATUS," & vbCrLf & _
                "APPT_DATE_TIME,TX_CODE_ID,STATUS_ID,STATUS,CHARTABLE,SIGNED," & vbCrLf & _
                "CREATE_DATE_TIME,ADA_CATEGORY_CODE,ADA_DEFINITION,PROVIDER_NAME," & vbCrLf & _
                "PATIENT_UNIT_ID,MODIFIED_BY,TIME_STAMP)"
            '-- select
            sSQL = sSQL & vbCrLf & _
                "select Z.GROUP_ID, Z.ENCOUNTER_NUMBER, Z.DENTAL_ENCOUNTER_ID, A.ENCOUNTER_STATUS, " & vbCrLf & _
                vbTab & GetEncounterStatusSQL("A.ENCOUNTER_STATUS") & ", " & vbCrLf & _
                vbTab & "A.APPT_DATE_TIME, Z.TX_CODE_ID, Z.STATUS_ID, Z.STATUS, " & vbCrLf & _
                vbTab & "Z.CHARTABLE, Z.SIGNED, Z.CREATE_DATE_TIME, Z.ADA_CATEGORY_CODE, " & vbCrLf & _
                vbTab & "Z.ADA_DEFINITION, CP.PROVIDER_NAME, A.PATIENT_UNIT_NUMBER, Z.MODIFIED_BY, Z.TIME_STAMP"
            '-- from
            sSQL = sSQL & vbCrLf & _
                "from Z_DATAVIEW_TX_1 Z, " & vbCrLf & _
                vbTab & "APPOINTMENT A, " & vbCrLf & _
                vbTab & "CLINIC_PROVIDER CP "
            '-- where
            sSQL = sSQL & vbCrLf & _
                "where Z.ENCOUNTER_FACILITY_NCID = A.FACILITY_NCID " & vbCrLf & _
                "and   Z.PATIENT_UNIT_ID = A.PATIENT_UNIT_NUMBER " & vbCrLf & _
                "and   Z.ENCOUNTER_NUMBER = A.ENCOUNTER_NUMBER " & vbCrLf & _
                "and   A.CLINIC_NCID = CP.CLINIC_NCID " & vbCrLf & _
                "and   Z.PROVIDER_ID = CP.CLINICIAN_NCID " & vbCrLf & _
                "and   " & sAppendWhereClause
        
        '--------------------------------------------------------------------------------
        '(II) Step 7. Retrieve Tx Chartable Data
        'Target Table: Z_DATAVIEW_TX_2
        'Source Table: Z_DATAVIEW_TX_1, APPOINTMENT, CLINIC_PROVIDER
        '--------------------------------------------------------------------------------
        Case glDXTX + 7
            sLogSqlDesc = "(II) Step " & CStr(glDXTX + 7) & ". Retrieve Tx Chartable Data"
            '-- insert
            sSQL = "insert into Z_DATAVIEW_TX_3 (" & vbCrLf & _
                vbTab & "GROUP_ID, ENCOUNTER_NUMBER,DENTAL_ENCOUNTER_ID,ENCOUNTER_STATUS_ID,ENCOUNTER_STATUS," & vbCrLf & _
                vbTab & "APPT_DATE_TIME,TX_CODE_ID,STATUS_ID,STATUS,CHARTABLE,SIGNED,TOOTH_NAME," & vbCrLf & _
                vbTab & "SURFACEABBR,CREATE_DATE_TIME,ADA_CATEGORY_CODE,ADA_DEFINITION," & vbCrLf & _
                vbTab & "PROVIDER_NAME,PATIENT_UNIT_ID,MODIFIED_BY,TIME_STAMP)"
            '-- select
            sSQL = sSQL & vbCrLf & _
                "select Z.GROUP_ID, Z.ENCOUNTER_NUMBER, Z.DENTAL_ENCOUNTER_ID, Z.ENCOUNTER_STATUS_ID, " & vbCrLf & _
                vbTab & "Z.ENCOUNTER_STATUS, Z.APPT_DATE_TIME, Z.TX_CODE_ID, Z.STATUS_ID, " & vbCrLf & _
                vbTab & "Z.STATUS, Z.CHARTABLE,Z.SIGNED, ET.TOOTH_NAME, ESS.SURFACEABBR, " & vbCrLf & _
                vbTab & "Z.CREATE_DATE_TIME, Z.ADA_CATEGORY_CODE, Z.ADA_DEFINITION, " & vbCrLf & _
                vbTab & "Z.PROVIDER_NAME, Z.PATIENT_UNIT_ID, Z.MODIFIED_BY, Z.TIME_STAMP"
            '-- from
            sSQL = sSQL & vbCrLf & _
                "from Z_DATAVIEW_TX_2 Z, " & vbCrLf & _
                vbTab & "DENTAL_ENCOUNTER_TXS_SS" & pgsAudit & " ETXS, " & vbCrLf & _
                vbTab & "DENTAL_ENCOUNTER_TEETH" & pgsAudit & " ET, " & vbCrLf & _
                vbTab & "DENTAL_ENC_SURFACE_SET" & pgsAudit & " ESS"
            '-- where
            sSQL = sSQL & vbCrLf & _
                "where Z.TX_CODE_ID = ETXS.TX_CODE_ID " & vbCrLf & _
                "and   ESS.SURFACE_SET_ID = ETXS.SURFACE_SET_ID " & vbCrLf & _
                "and   ET.DENTAL_ENCOUNTER_TEETH_ID = ESS.DENTAL_ENCOUNTER_TEETH_ID " & vbCrLf & _
                "and   Z.CHARTABLE = -1 " & vbCrLf & _
                "and   " & sAppendWhereClause
        
        '--------------------------------------------------------------------------------
        '(II) Step 8. Retrieve Tx No Chartable Data
        'Target Table:  Z_DATAVIEW_TX_3
        'Source Table:  Z_DATAVIEW_TX_2, NOTE_TXS_LOCATION, NOTE_LOCATION,
        '               NOTE_TX_SUB_LOCATION, NOTE_SUB_LOCATION
        '--------------------------------------------------------------------------------
        Case glDXTX + 8
            sLogSqlDesc = "(II) Step " & CStr(glDXTX + 8) & ". Retrieve Tx No Chartable Data"
            '-- Insert
            sSQL = "insert into Z_DATAVIEW_TX_3 (" & vbCrLf & _
                vbTab & "GROUP_ID, ENCOUNTER_NUMBER,DENTAL_ENCOUNTER_ID,ENCOUNTER_STATUS_ID,ENCOUNTER_STATUS," & vbCrLf & _
                vbTab & "APPT_DATE_TIME,TX_CODE_ID,STATUS_ID,STATUS,CHARTABLE,SIGNED,TOOTH_NAME," & vbCrLf & _
                vbTab & "SURFACEABBR,CREATE_DATE_TIME,ADA_CATEGORY_CODE,ADA_DEFINITION," & vbCrLf & _
                vbTab & "PROVIDER_NAME,PATIENT_UNIT_ID,MODIFIED_BY,TIME_STAMP)"
            '-- select
            sSQL = sSQL & vbCrLf & _
                "select Z.GROUP_ID, Z.ENCOUNTER_NUMBER, Z.DENTAL_ENCOUNTER_ID, Z.ENCOUNTER_STATUS_ID, " & vbCrLf & _
                vbTab & "Z.ENCOUNTER_STATUS, Z.APPT_DATE_TIME, Z.TX_CODE_ID, Z.STATUS_ID, " & vbCrLf & _
                vbTab & "Z.STATUS, Z.CHARTABLE, Z.SIGNED, NL.LOCATION_NAME, " & vbCrLf & _
                vbTab & "V.SUB_LOCATION_NAME, Z.CREATE_DATE_TIME, Z.ADA_CATEGORY_CODE, " & vbCrLf & _
                vbTab & "Z.ADA_DEFINITION, Z.PROVIDER_NAME, Z.PATIENT_UNIT_ID, Z.MODIFIED_BY, Z.TIME_STAMP"
                
            If mobjCHCSIIConn.UseSQLServer Then
                'SQL Server format
                sSQL = sSQL & " FROM Z_DATAVIEW_TX_2 Z INNER JOIN" _
                    & " NOTE_TXS_LOCATION NTXSL ON Z.TX_CODE_ID = NTXSL.TX_CODE_ID INNER JOIN" _
                    & " NOTE_LOCATION NL ON NTXSL.LOCATION_ID = NL.LOCATION_ID LEFT OUTER JOIN" _
                    & " V_TX_SUB_LOCATION V ON NTXSL.TX_LOCATION_ID = V.TX_LOCATION_ID" _
                    & " WHERE (Z.CHARTABLE = 0) AND " & sAppendWhereClause
            Else
                'Oracle format
                 '-- from
                sSQL = sSQL & vbCrLf & _
                    "from Z_DATAVIEW_TX_2 Z," & vbCrLf & _
                    vbTab & "NOTE_TXS_LOCATION" & pgsAudit & " NTXSL," & vbCrLf & _
                    vbTab & "NOTE_LOCATION NL," & vbCrLf & _
                    vbTab & "V_TX_SUB_LOCATION V"
                '-- where
                sSQL = sSQL & vbCrLf & _
                    "where Z.TX_CODE_ID = NTXSL.TX_CODE_ID " & vbCrLf & _
                    "and   NL.LOCATION_ID = NTXSL.LOCATION_ID " & vbCrLf & _
                    "and   NTXSL.TX_LOCATION_ID = V.TX_LOCATION_ID (+) " & vbCrLf & _
                    "and   Z.CHARTABLE = 0 " & vbCrLf & _
                    "and   " & sAppendWhereClause
            End If
            
               
        '--------------------------------------------------------------------------------
        '(III) Step 9. Retrieve DxTx No Data
        'Target Table:  Z_DATAVIEW_DXTX
        'Source Table:  Z_DATAVIEW_DX_3, DENTAL_ENCOUNTER_TXS_DXS, Z_DATAVIEW_TX_3
        '--------------------------------------------------------------------------------
        Case glDXTX + 9
            sLogSqlDesc = "(III) Step " & CStr(glDXTX + 9) & ". Retrieve DxTx Data"
            '-- Insert
            sSQL = "insert into Z_DATAVIEW_DXTX (" & vbCrLf & _
                vbTab & "ENCOUNTER_NUMBER,DENTAL_ENCOUNTER_ID,ENCOUNTER_STATUS_ID,ENCOUNTER_STATUS," & vbCrLf & _
                vbTab & "APPT_DATE_TIME,DX_CODE_ID,DX_STATUS_ID,DX_STATUS,DX_CHARTABLE,DRC," & vbCrLf & _
                vbTab & "DX_SIGNED,DX_LOCATION,DX_SURFACE,DX_DATE,DX_PROVIDER_NAME,DX_CODE," & vbCrLf & _
                vbTab & "DIAGNOSIS,TX_CODE_ID,TX_STATUS_ID,TX_STATUS,TX_CHARTABLE,TX_SIGNED," & vbCrLf & _
                vbTab & "TX_LOCATION,TX_SURFACE,TX_DATE,TX_PROVIDER_NAME,ADA_CATEGORY_CODE," & vbCrLf & _
                vbTab & "TREATMENT,PATIENT_UNIT_ID,MODIFIED_BY,TIME_STAMP)"
            '-- select
            sSQL = sSQL & vbCrLf & _
                "select Z.ENCOUNTER_NUMBER, Z.DENTAL_ENCOUNTER_ID, Z.ENCOUNTER_STATUS_ID, " & vbCrLf & _
                vbTab & "Z.ENCOUNTER_STATUS, Z.APPT_DATE_TIME, Z.DX_CODE_ID, " & vbCrLf & _
                vbTab & "Z.STATUS_ID as DX_STATUS_ID, Z.STATUS as DX_STATUS, " & vbCrLf & _
                vbTab & "Z.CHARTABLE as DX_CHARTABLE, Z.DRC, Z.SIGNED as DX_SIGNED, " & vbCrLf & _
                vbTab & "Z.TOOTH_NAME as DX_LOCATION, Z.SURFACEABBR as DX_SURFACE, " & vbCrLf & _
                vbTab & "Z.CREATE_DATE_TIME as DX_DATE, Z.PROVIDER_NAME as DX_PROVIDER_NAME, " & vbCrLf & _
                vbTab & "Z.ICD_CODE as DX_CODE, Z.ICD_DEFINITION as DIAGNOSIS, " & vbCrLf & _
                vbTab & "Z2.TX_CODE_ID, Z2.STATUS_ID as TX_STATUS_ID, Z2.STATUS as TX_STATUS, " & vbCrLf & _
                vbTab & "Z2.CHARTABLE as TX_CHARTABLE, Z2.SIGNED as TX_SIGNED, " & vbCrLf & _
                vbTab & "Z2.TOOTH_NAME as TX_LOCATION, Z2.SURFACEABBR as TX_SURFACE, " & vbCrLf & _
                vbTab & "Z2.APPT_DATE_TIME as TX_DATE, Z2.PROVIDER_NAME as TX_PROVIDER_NAME, " & vbCrLf & _
                vbTab & "Z2.ADA_CATEGORY_CODE, Z2.ADA_DEFINITION as Treatment, " & vbCrLf & _
                vbTab & "Z.PATIENT_UNIT_ID, Z.MODIFIED_BY, Z.TIME_STAMP"
            '-- from
            sSQL = sSQL & vbCrLf & _
                "from Z_DATAVIEW_DX_3 Z," & vbCrLf & _
                vbTab & "DENTAL_ENCOUNTER_GROUPS" & pgsAudit & " G1," & vbCrLf & _
                vbTab & "DENTAL_ENCOUNTER_GROUPS" & pgsAudit & " G2," & vbCrLf & _
                vbTab & "Z_DATAVIEW_TX_3 Z2"
            '-- where
            sSQL = sSQL & vbCrLf & _
                "where Z.DX_CODE_ID = G1.DX_CODE_ID " & vbCrLf & _
                "and   Z2.TX_CODE_ID = G2.TX_CODE_ID " & vbCrLf & _
                "and   G1.GROUP_ID = G2.GROUP_ID " & vbCrLf & _
                "and   Z.MODIFIED_BY = Z2.MODIFIED_BY " & vbCrLf & _
                "and   Z.TIME_STAMP = Z2.TIME_STAMP " & vbCrLf & _
                "and   " & sAppendWhereClause
                    
        '----------------------------------------------------------
        '(III) Step 10. <<<Final DxTx data>>>
        'Target Table:  Final Record set
        'Source Table:  Z_DATAVIEW_DX_3, Z_DATAVIEW_DXTX
        '----------------------------------------------------------
        Case glDXTX + 10
            sLogSqlDesc = "(III) Step " & CStr(glDXTX + 10) & ". <<<Final DxTx data>>> "
            '-- select
            sSQL = "select Z.GROUP_ID, Z.ICD_DEFINITION as DIAGOSIS, " & vbCrLf & _
                vbTab & "Z.ICD_CODE as DX_CODE, " & vbCrLf & _
                vbTab & "Z.TOOTH_NAME as DX_LOCATION, " & vbCrLf & _
                vbTab & "Z.SURFACEABBR as DX_SURFACE, " & vbCrLf & _
                vbTab & "Z.APPT_DATE_TIME as DX_DATE, " & vbCrLf & _
                vbTab & "Z.PROVIDER_NAME as DX_PROVIDER_NAME, " & vbCrLf & _
                vbTab & "Z.STATUS as DX_STATUS, " & vbCrLf & _
                vbTab & "Z.DRC, " & vbCrLf & _
                vbTab & "Z2.ADA_DEFINITION as TREATMENT, " & vbCrLf & _
                vbTab & "Z2.TOOTH_NAME as TX_LOCATION, " & vbCrLf & _
                vbTab & "Z2.SURFACEABBR as TX_SURFACE, " & vbCrLf & _
                vbTab & "Z2.STATUS as TX_STATUS, " & vbCrLf & _
                vbTab & "Z2.ADA_CATEGORY_CODE as TX_CODE, " & vbCrLf & _
                vbTab & "Z2.APPT_DATE_TIME as TX_DATE, " & vbCrLf & _
                vbTab & "Z2.PROVIDER_NAME as TX_PROVIDER_NAME, " & vbCrLf & _
                vbTab & "Z2.ENCOUNTER_NUMBER, " & vbCrLf & _
                vbTab & "Z2.ENCOUNTER_STATUS_ID, " & vbCrLf & _
                vbTab & "Z2.ENCOUNTER_STATUS, " & vbCrLf & _
                vbTab & "Z.STATUS_ID as DX_STATUS_ID, " & vbCrLf & _
                vbTab & "Z.SIGNED as DX_SIGNED, " & vbCrLf & _
                vbTab & "Z.CHARTABLE as DX_CHARTABLE, " & vbCrLf & _
                vbTab & "Z2.STATUS_ID as TX_STATUS_ID, " & vbCrLf & _
                vbTab & "Z2.SIGNED AS TX_SIGNED, " & vbCrLf & _
                vbTab & "Z2.CHARTABLE as TX_CHARTABLE "
            '-- from
            If mobjCHCSIIConn.UseSQLServer Then
                'SQL Server format
                '-- from and where
                sSQL = sSQL & "FROM Z_DATAVIEW_DX_3 Z LEFT OUTER JOIN" _
                    & " Z_DATAVIEW_TX_3 Z2 ON Z.GROUP_ID = Z2.GROUP_ID AND Z.MODIFIED_BY = Z2.MODIFIED_BY AND Z.TIME_STAMP = Z2.TIME_STAMP" _
                    & " WHERE " & sAppendWhereClause
            Else
                'Oracle format
                '-- from
                sSQL = sSQL & vbCrLf & _
                    "from Z_DATAVIEW_DX_3 Z, " & vbCrLf & _
                    vbTab & "Z_DATAVIEW_TX_3 Z2 "
                '-- where
                sSQL = sSQL & vbCrLf & _
                    "where Z.GROUP_ID = Z2.GROUP_ID (+) " & vbCrLf & _
                    "and   Z.MODIFIED_BY = Z2.MODIFIED_BY (+) " & vbCrLf & _
                    "and   Z.TIME_STAMP = Z2.TIME_STAMP (+) " & vbCrLf & _
                    "and   " & sAppendWhereClause
            End If
            '-- union
            sSQL = sSQL & vbCrLf & _
                "Union "
            '-- select
            sSQL = sSQL & vbCrLf & _
                "select Z2.GROUP_ID, Z.ICD_DEFINITION as DIAGOSIS, " & vbCrLf & _
                vbTab & "Z.ICD_CODE as DX_CODE, " & vbCrLf & _
                vbTab & "Z.TOOTH_NAME as DX_LOCATION, " & vbCrLf & _
                vbTab & "Z.SURFACEABBR as DX_SURFACE, " & vbCrLf & _
                vbTab & "Z.APPT_DATE_TIME as DX_DATE, " & vbCrLf & _
                vbTab & "Z.PROVIDER_NAME as DX_PROVIDER_NAME, " & vbCrLf & _
                vbTab & "Z.STATUS as DX_STATUS, " & vbCrLf & _
                vbTab & "Z.DRC, " & vbCrLf & _
                vbTab & "Z2.ADA_DEFINITION as TREATMENT, " & vbCrLf & _
                vbTab & "Z2.TOOTH_NAME as TX_LOCATION, " & vbCrLf & _
                vbTab & "Z2.SURFACEABBR as TX_SURFACE, " & vbCrLf & _
                vbTab & "Z2.STATUS as TX_STATUS, " & vbCrLf & _
                vbTab & "Z2.ADA_CATEGORY_CODE as TX_CODE, " & vbCrLf & _
                vbTab & "Z2.APPT_DATE_TIME as TX_DATE, " & vbCrLf & _
                vbTab & "Z2.PROVIDER_NAME as TX_PROVIDER_NAME, " & vbCrLf & _
                vbTab & "Z2.ENCOUNTER_NUMBER, " & vbCrLf & _
                vbTab & "Z2.ENCOUNTER_STATUS_ID, " & vbCrLf & _
                vbTab & "Z2.ENCOUNTER_STATUS, " & vbCrLf & _
                vbTab & "Z.STATUS_ID as DX_STATUS_ID, " & vbCrLf & _
                vbTab & "Z.SIGNED as DX_SIGNED, " & vbCrLf & _
                vbTab & "Z.CHARTABLE as DX_CHARTABLE, " & vbCrLf & _
                vbTab & "Z2.STATUS_ID as TX_STATUS_ID, " & vbCrLf & _
                vbTab & "Z2.SIGNED AS TX_SIGNED, " & vbCrLf & _
                vbTab & "Z2.CHARTABLE as TX_CHARTABLE "
            '-- from
            If mobjCHCSIIConn.UseSQLServer Then
                'SQL Server format
                '-- from and where and order
                sSQL = sSQL & " FROM Z_DATAVIEW_DX_3 Z RIGHT OUTER JOIN" _
                    & " Z_DATAVIEW_TX_3 Z2 ON Z.GROUP_ID = Z2.GROUP_ID AND Z.MODIFIED_BY = Z2.MODIFIED_BY AND Z.TIME_STAMP = Z2.TIME_STAMP" _
                    & " WHERE " & sAppendWhereClause2 & " ORDER BY 1"
            Else
                'Oracle format
                '-- from
                sSQL = sSQL & vbCrLf & _
                    "from Z_DATAVIEW_DX_3 Z, " & vbCrLf & _
                    vbTab & "Z_DATAVIEW_TX_3 Z2 "
                    
                '-- where
                sSQL = sSQL & vbCrLf & _
                    "where Z.GROUP_ID (+) = Z2.GROUP_ID " & vbCrLf & _
                    "and   Z.MODIFIED_BY (+) = Z2.MODIFIED_BY " & vbCrLf & _
                    "and   Z.TIME_STAMP (+) = Z2.TIME_STAMP " & vbCrLf & _
                    "and   " & sAppendWhereClause2
                    
                '-- order by
                sSQL = sSQL & vbCrLf & _
                    "order by 1"
            End If
            
        '----------------------------------------------------------
        '(IV) Step 11. delete Z_DATAVIEW_DX_1 table
        '----------------------------------------------------------
        Case glDXTX + 11
            sLogSqlDesc = "(IV) Step " & CStr(glDXTX + 11) & ". delete Z_DATAVIEW_DX_1 table"
            If mobjCHCSIIConn.UseSQLServer Then
                'SQL Server format
                sSQL = "delete from Z_DATAVIEW_DX_1 " & vbCrLf & _
                "where " & Replace(sAppendWhereClause, "Z.", "")
            Else
                'Oracle format
                sSQL = "delete from Z_DATAVIEW_DX_1 Z " & vbCrLf & _
                "where " & sAppendWhereClause
            End If
            
        '----------------------------------------------------------
        '(IV) Step 12. delete Z_DATAVIEW_DX_2 table
        '----------------------------------------------------------
        Case glDXTX + 12
            sLogSqlDesc = "(IV) Step " & CStr(glDXTX + 12) & ". delete Z_DATAVIEW_DX_2 table"
            If mobjCHCSIIConn.UseSQLServer Then
                'SQL Server format
                sSQL = "delete from Z_DATAVIEW_DX_2 " & vbCrLf & _
                "where " & Replace(sAppendWhereClause, "Z.", "")
            Else
                'Oracle
                sSQL = "delete from Z_DATAVIEW_DX_2 Z " & vbCrLf & _
                "where " & sAppendWhereClause
            End If

        '----------------------------------------------------------
        '(IV) Step 13. delete Z_DATAVIEW_DX_3 table
        '----------------------------------------------------------
        Case glDXTX + 13
            sLogSqlDesc = "(IV) Step " & CStr(glDXTX + 13) & ". delete Z_DATAVIEW_DX_3 table"
            If mobjCHCSIIConn.UseSQLServer Then
                'SQL Server format
                sSQL = "delete from Z_DATAVIEW_DX_3 " & vbCrLf & _
                "where " & Replace(sAppendWhereClause, "Z.", "")
            Else
                'Oracle
                sSQL = "delete from Z_DATAVIEW_DX_3 Z " & vbCrLf & _
                "where " & sAppendWhereClause
            End If
            
        '----------------------------------------------------------
        '(IV) Step 14. delete Z_DATAVIEW_TX_1 table
        '----------------------------------------------------------
        Case glDXTX + 14
            sLogSqlDesc = "(IV) Step " & CStr(glDXTX + 14) & ". delete Z_DATAVIEW_TX_1 table"
            If mobjCHCSIIConn.UseSQLServer Then
                'SQL Server format
                sSQL = "delete from Z_DATAVIEW_TX_1 " & vbCrLf & _
                "where " & Replace(sAppendWhereClause, "Z.", "")
            Else
                'Oracle
                sSQL = "delete from Z_DATAVIEW_TX_1 Z " & vbCrLf & _
                "where " & sAppendWhereClause
            End If
        '----------------------------------------------------------
        '(IV) Step 15. delete Z_DATAVIEW_TX_2 table
        '----------------------------------------------------------
        Case glDXTX + 15
            sLogSqlDesc = "(IV) Step " & CStr(glDXTX + 15) & ". delete Z_DATAVIEW_TX_2 table"
            If mobjCHCSIIConn.UseSQLServer Then
                'SQL Server format
                sSQL = "delete from Z_DATAVIEW_TX_2 " & vbCrLf & _
                "where " & Replace(sAppendWhereClause, "Z.", "")
            Else
                'Oracle
                sSQL = "delete from Z_DATAVIEW_TX_2 Z " & vbCrLf & _
                "where " & sAppendWhereClause
            End If

        '----------------------------------------------------------
        '(IV) Step 16. delete Z_DATAVIEW_TX_3 table
        '----------------------------------------------------------
        Case glDXTX + 16
            sLogSqlDesc = "(IV) Step " & CStr(glDXTX + 16) & ". delete Z_DATAVIEW_TX_3 table"
            If mobjCHCSIIConn.UseSQLServer Then
                'SQL Server format
                sSQL = "delete from Z_DATAVIEW_TX_3 " & vbCrLf & _
                "where " & Replace(sAppendWhereClause, "Z.", "")
            Else
                'Oracle
                sSQL = "delete from Z_DATAVIEW_TX_3 Z " & vbCrLf & _
                "where " & sAppendWhereClause
            End If
        '----------------------------------------------------------
        '(IV) Step 17. delete Z_DATAVIEW_DXTX table
        '----------------------------------------------------------
        Case glDXTX + 17
            sLogSqlDesc = "(IV) Step " & CStr(glDXTX + 17) & ". delete Z_DATAVIEW_DXTX table"
            If mobjCHCSIIConn.UseSQLServer Then
                'SQL Server format
                sSQL = "delete from Z_DATAVIEW_DXTX " & vbCrLf & _
                "where " & Replace(sAppendWhereClause, "Z.", "")
            Else
                'Oracle
                sSQL = "delete from Z_DATAVIEW_DXTX Z " & vbCrLf & _
                "where " & sAppendWhereClause
            End If
        '***************************************************************************************************
        '(Lab Tab) Step glLAB_CASE (300). Retrieve Lab Case Data
        '***************************************************************************************************
        Case glLAB_CASE
            sLogSqlDesc = ">Lab Case Process<-- Step " & CStr(glLAB_CASE) & ". Retrieve Lab Case Data"
            sSQL = ProcessLabCase()

        '***************************************************************************************************
        '(Workload Tab) Step glWORKLOAD (400). Retrieve Workload Data
        '***************************************************************************************************
        Case glWORKLOAD
            sLogSqlDesc = ">Workload Process<-- Step " & CStr(glWORKLOAD) & ". Retrieve Workload Data"
            sSQL = ProcessWorkload()

    End Select
                       
    '--log sql
    If mblnLogSQL Then
        LogSql sLogSqlDesc, sSQL, pbCreateLogFile, piProcessStep
    End If
    
    'Convert the || to + for SQL Server
    If mobjCHCSIIConn.UseSQLServer Then
        sSQL = Replace(sSQL, "||", "+")
    End If
    
    '--execute mothed
    Set prRs = Nothing
    Select Case piProcessStep
        Case glDXTX + 10, glLAB_CASE, glWORKLOAD, glENCOUNTER_TYPE 'for select SQL only
            Set prRs = mobjDas.OpenRecordset(sSQL)
        Case Else
            mobjSQLOps.Execute sSQL
    End Select
    
    SqlProcess = True
    GoTo CleanUp
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "basLoad.SqlProcess of step " & piProcessStep, App.Title, vbCritical
CleanUp:
'    Set goSqlOps = Nothing
'    Set goDas = Nothing
    Exit Function
    Resume
End Function

'***************************************************************************************************
' Name:         ProcessEncounterType
' Description:  Build Sql Statement for process Encounter Type
' Parameter:    spEncounterType - Encounter Type
' Return:       String
' Called By:    SqlProcess
' Histary:
'   Developer:      Date:       Description
'   ----------      ---------   -----------------------------------------------
'   Wilson Sun      01/30/2002  Display detail Info for every Encounter Type.
'   Wilson Sun      01/08/2002  Created
'**************************************************************************************************
Private Function ProcessEncounterType() As String
    '==================================================================
    '(Problems/Planned Tx and Current Oral State Tab)
    'Step glENCOUNTER_TYPE (300). Retrieve Encounter Type Data
    '==================================================================
    Dim sEncounterType As String
    Dim sAppendFrom As String
    Dim sAppendWhere As String
    Dim sSqlEmergency As String
    Dim sSqlExam As String
    Dim sSqlEval As String
    Dim sSqlPreventive As String
    Dim sSqlTxHygiene As String
    Dim sSqlTxChartable As String
    Dim sSqlTxUnChartable As String
    Dim sSqlNull As String
    Dim sSQLUnknown As String
        
    '--------------------------------------------------------------------------------------------
    '1. Emergency Type
    'Source Table: APPOINTMENT A, DENTAL_ENCOUNTER D, TEMPL_EXAMSOAP T
    'Note: Exam/Evals (SOAP) Tab. Check Sick Call , After Hour Care Chief Complaint
    '--------------------------------------------------------------------------------------------
    '-- Select
'    sEncounterType = "'Emergency (Chief Complaint, Sick Call, After Hours)'"
    sEncounterType = "'Emergency -'"
        
    If mobjCHCSIIConn.UseSQLServer Then
        'SQL Server format
        
        sEncounterType = "CASE CASE T .SICK_CALL WHEN - 1 THEN ' Sick Call.' ELSE '' END + CASE T .AFTER_HOURSCARE WHEN - 1 THEN ' After Hour Care.' ELSE '' END + CASE" _
            & " T .CHIEF_COMPLAINT WHEN '' THEN '' ELSE ' Chief Complaint (' + T .CHIEF_COMPLAINT + ').' END WHEN '' THEN 'Unknown Encounter Type' ELSE 'Emergency -'" _
        & " + CASE T .SICK_CALL WHEN - 1 THEN ' Sick Call.' ELSE '' END + CASE T .AFTER_HOURSCARE WHEN - 1 THEN ' After Hour Care.' ELSE '' END + CASE T.CHIEF_COMPLAINT" _
        & " WHEN '' THEN '' ELSE ' Chief Complaint (' + T .CHIEF_COMPLAINT + ').' END END"
                       
'        sEncounterType = "CASE Case T.SICK_CALL when -1 then ' Sick Call.' ELSE '' END " & vbCrLf & _
'           "+ Case T.AFTER_HOURSCARE when -1 then ' After Hour Care.' ELSE '' END " & vbCrLf & _
'           "+ Case T.CHIEF_COMPLAINT when '' then '' ELSE ' Chief Complaint (' + T.CHIEF_COMPLAINT + ').' END " _
'           "WHEN '' then 'Unknown Encounter Type' ELSE " &  sEncounterType & " & " END "
    Else
        'Oracle format
        sEncounterType = sEncounterType & " || DECODE(T.SICK_CALL,-1,' Sick Call.') " & vbCrLf & _
           "|| DECODE(T.AFTER_HOURSCARE,-1,' After Hour Care.')" & vbCrLf & _
           "|| DECODE(T.CHIEF_COMPLAINT,null,'', ' Chief Complaint (' || T.CHIEF_COMPLAINT || ').')"
    End If
    
    '-- From
    sAppendFrom = ", TEMPL_EXAMGENERAL T "
    '-- Where
    sAppendWhere = "AND   D.DENTAL_ENCOUNTER_ID = T.DENTAL_ENCOUNTER_ID " & vbCrLf & _
            "AND   (T.SICK_CALL = - 1" & vbCrLf & _
            "   OR T.AFTER_HOURSCARE  = - 1 " & vbCrLf & _
            "   OR T.CHIEF_COMPLAINT is not null) "
    
    '-- Sql for Emergency
    sSqlEmergency = BuildSql(sEncounterType, sAppendFrom, sAppendWhere)
    
    '--------------------------------------------------------------------------------------------
    '2. Exam Type
    'Source Table: APPOINTMENT A, DENTAL_ENCOUNTER D, TEMPL_EXAMGENERAL T
    'Note: Exam/Evals (General) Tab. Check Periodic Initial Charting and Comprehensive
    '--------------------------------------------------------------------------------------------
    '-- Select
    '    sEncounterType = "'Exam (Periodic, Initial Charging, Comprehensive)'"
    sEncounterType = "'Exam -'"
    sEncounterType = sEncounterType & vbCrLf & _
       "|| " & GetExamEncounterTypeSQL("T.ORALEXAM")

    '-- From
    sAppendFrom = ", TEMPL_EXAMGENERAL T "
    '-- Where
    sAppendWhere = "AND   D.DENTAL_ENCOUNTER_ID = T.DENTAL_ENCOUNTER_ID " & vbCrLf & _
            "AND   T.ORALEXAM IN (1, 2, 3) "
    '-- Sql for Exam
    sSqlExam = BuildSql(sEncounterType, sAppendFrom, sAppendWhere)

    '--------------------------------------------------------------------------------------------
    '3. Eval Type
    'Source Table: APPOINTMENT A, DENTAL_ENCOUNTER D, TEMPL_EXAMSOAP T
    'Note: Exam/Evals (General) Tab. Check Evaluation group
    '--------------------------------------------------------------------------------------------
    '-- Select
'    sEncounterType = "'Eval (SOAP - Evaluation) Specialty: '"
    sEncounterType = "'Eval -'"
    
    If mobjCHCSIIConn.UseSQLServer Then
        'SQL Server format
        sEncounterType = sEncounterType & " + Case T.SICK_CALL when -1 then ' Sick Call.' ELSE '' END " & vbCrLf & _
           "+ Case T.ORAL_SURGERY when -1 then ' Oral Surgery.' ELSE '' END " & vbCrLf & _
           "+ Case T.PERIODONTAL when -1 then ' Periodontal.' ELSE '' END " & vbCrLf & _
           "+ Case T.PROSTHODONTICS when -1 then ' Prosthodontics.' ELSE '' END " & vbCrLf & _
           "+ Case T.ENDODONTICS when -1 then ' Endodontics.' ELSE '' END " & vbCrLf & _
           "+ Case T.OralExam when 9 then ' Postmortem.' ELSE '' END " & vbCrLf & _
           "+ Case T.AFTER_HOURSCARE when -1 then ' After Hour Care.' ELSE '' END " & vbCrLf & _
           "+ Case T.ORTHODONTICS when -1 then ' Orthodontics.' ELSE '' END " & vbCrLf & _
           "+ Case T.Orofacial_Pain when -1 then ' Orofacial pain .' ELSE '' END " & vbCrLf & _
           "+ Case T.CONSULTATION_RESPONSE when -1 then ' Consultation Response.' ELSE '' END "
    Else
        'Oracle format
        sEncounterType = sEncounterType & " || DECODE(T.SICK_CALL,-1,' Sick Call.') " & vbCrLf & _
           "|| DECODE(T.ORAL_SURGERY,-1,' Oral Surgery.')" & vbCrLf & _
           "|| DECODE(T.PERIODONTAL,-1,' Periodontal.')" & vbCrLf & _
           "|| DECODE(T.PROSTHODONTICS,-1,' Prosthodontics.')" & vbCrLf & _
           "|| DECODE(T.ENDODONTICS,-1,' Endodontics.')" & vbCrLf & _
           "|| DECODE(T.OralExam,9,' Postmortem.')" & vbCrLf & _
           "|| DECODE(T.AFTER_HOURSCARE,-1,' After Hour Care.')" & vbCrLf & _
           "|| DECODE(T.ORTHODONTICS,-1,' Orthodontics.')" & vbCrLf & _
           "|| DECODE(T.Orofacial_Pain,-1,' Orofacial pain .')" & vbCrLf & _
           "|| DECODE(T.CONSULTATION_RESPONSE,-1,' Consultation Response.')"
    End If
    
    '-- From
    sAppendFrom = ", TEMPL_EXAMGENERAL T "
    '-- Where
    sAppendWhere = "AND   D.DENTAL_ENCOUNTER_ID = T.DENTAL_ENCOUNTER_ID " & vbCrLf & _
            "AND    (T.SICK_CALL = - 1 " & vbCrLf & _
            "   OR T.ORAL_SURGERY = - 1 " & vbCrLf & _
            "   OR T.PERIODONTAL = - 1 " & vbCrLf & _
            "   OR T.PROSTHODONTICS = - 1 " & vbCrLf & _
            "   OR T.ENDODONTICS = - 1 " & vbCrLf & _
            "   OR T.OralExam = 9 " & vbCrLf & _
            "   OR AFTER_HOURSCARE = - 1 " & vbCrLf & _
            "   OR T.ORTHODONTICS = - 1 " & vbCrLf & _
            "   OR T.Orofacial_Pain = - 1 " & vbCrLf & _
            "   OR T.CONSULTATION_RESPONSE = - 1 )"

   '-- Sql for Eval
    sSqlEval = BuildSql(sEncounterType, sAppendFrom, sAppendWhere)
          
    '--------------------------------------------------------------------------------------------
    '4. Preventive Type
    'Source Table: APPOINTMENT A, DENTAL_ENCOUNTER D, TEMPL_PREVENTIVE T
    'Note:  Preventive (Prophylaxis/Fluoride) Tab.
    '       Check Dental Prophylaxis
    '--------------------------------------------------------------------------------------------
    '-- Select
'    sEncounterType = "'Preventive (Prophy, Mouthguard, Fluoride, Polishing, Scaling)'"
    sEncounterType = "'Preventive -'"
    
    If mobjCHCSIIConn.UseSQLServer Then
        'SQL Server format
        sEncounterType = "CASE " & sEncounterType & vbCrLf & _
            "+ Case T.PRH_PROPH_COMPLETE when -1 then ' Complete prophylaxis.' ELSE '' END " & vbCrLf & _
            "+ Case T.PRH_PROPH_INCOMPLETE when -1 then ' Incomplete prophylaxis.' ELSE '' END " & vbCrLf & _
            "+ Case T.PRH_PROPH_FLUORIDE_SNF2 when -1 then ' SnF2.' ELSE '' END " & vbCrLf & _
            "+ Case T.PRH_PROPH_FLUORIDE_APF when -1 then ' APF.' ELSE '' END " & vbCrLf & _
            "+ Case T.PRH_PROPH_FLUORIDE_NAF when -1 then ' NaF.' ELSE '' END " & vbCrLf & _
            "+ Case T.PRH_PROPH_FLUORIDE_VARNISH when -1 then ' Fluoride varnish.' ELSE '' END " & vbCrLf & _
            "+ Case T.PRH_SUP_FULLMOUTH when -1 then ' Full mouth supragingival scaling.' ELSE '' END " & vbCrLf & _
            "+ Case T.PRH_SUP_MAXRIGHT when -1 then ' Scaling Max right sextant.' ELSE '' END " & vbCrLf & _
            "+ Case T.PRH_SUP_MAXANT when -1 then ' Scaling Max ant sextant.' ELSE '' END " & vbCrLf & _
            "+ Case T.PRH_SUP_MAXLEFT when -1 then ' Scaling Man right sextant.' ELSE '' END " & vbCrLf & _
            "+ Case T.PRH_SUP_MANRIGHT when -1 then ' Scaling Man right sextant.' ELSE '' END " & vbCrLf & _
            "+ Case T.PRH_SUP_MANANT when -1 then ' Scaling Man ant sextant.' ELSE '' END " & vbCrLf & _
            "+ Case T.PRH_SUP_MANLEFT when -1 then ' Scaling Man left sextant.' ELSE '' END " & vbCrLf & _
            "+ Case T.PRH_POL_FULLMOUTH when -1 then ' Full mouth polished.' ELSE '' END " & vbCrLf & _
            "+ Case T.PRH_POL_MAXRIGHT when -1 then ' Polish Max right sextant.' ELSE '' END " & vbCrLf & _
            "+ Case T.PRH_POL_MAXANT when -1 then ' Polish Max ant sextant.' ELSE '' END " & vbCrLf & _
            "+ Case T.PRH_POL_MAXLEFT when -1 then ' Polish Max left sextant.' ELSE '' END " & vbCrLf & _
            "+ Case T.PRH_POL_MANRIGHT when -1 then ' Polish Man right sextant.' ELSE '' END " & vbCrLf & _
            "+ Case T.PRH_POL_MANANT when -1 then ' Polish Man ant sextant.' ELSE '' END " & vbCrLf & _
            "+ Case T.PRH_POL_MANLEFT when -1 then ' Polish Man left sextant.' ELSE '' END "
        sEncounterType = sEncounterType & vbCrLf & _
            "+ Case T.PRH_PROPHYLAXIS_COMMENTS when '' then '' else ' Prophylaxis Comments (' + T.PRH_PROPHYLAXIS_COMMENTS + ').' END " & vbCrLf & _
            "+ Case T.PRV_IM_MAXMOUTHGUARD when -1 then ' Max mouthguard.' ELSE '' END " & vbCrLf & _
            "+ Case T.PRV_IM_MANMOUTHGUARD when -1 then ' Man mouthguard.' ELSE '' END " & vbCrLf & _
            "+ Case T.PRV_DL_MAXATHLETIC when -1 then ' Max Athletic mouthguard.' ELSE '' END " & vbCrLf & _
            "+ Case T.PRV_DL_MANATHLETIC when -1 then ' Man Athletic mouthguard.' ELSE '' END WHEN 'Preventive -' THEN 'Unknown Encounter Type' ELSE "
            
        sEncounterType = sEncounterType & vbCrLf _
            & " 'Preventive -' + CASE T .PRH_PROPH_COMPLETE WHEN - 1 THEN ' Complete prophylaxis.' ELSE '' END + CASE T .PRH_PROPH_INCOMPLETE WHEN" _
            & " - 1 THEN ' Incomplete prophylaxis.' ELSE '' END + CASE T .PRH_PROPH_FLUORIDE_SNF2 WHEN - 1 THEN ' SnF2.' ELSE '' END + CASE T .PRH_PROPH_FLUORIDE_APF" _
            & " WHEN - 1 THEN ' APF.' ELSE '' END + CASE T .PRH_PROPH_FLUORIDE_NAF WHEN - 1 THEN ' NaF.' ELSE '' END + CASE T .PRH_PROPH_FLUORIDE_VARNISH" _
            & " WHEN - 1 THEN ' Fluoride varnish.' ELSE '' END + CASE T .PRH_SUP_FULLMOUTH WHEN - 1 THEN ' Full mouth supragingival scaling.' ELSE '' END +" _
            & " CASE T .PRH_SUP_MAXRIGHT WHEN - 1 THEN ' Scaling Max right sextant.' ELSE '' END + CASE T .PRH_SUP_MAXANT WHEN - 1 THEN ' Scaling Max ant sextant.'" _
            & " Else '' END + CASE T .PRH_SUP_MAXLEFT WHEN - 1 THEN ' Scaling Man right sextant.' ELSE '' END + CASE T .PRH_SUP_MANRIGHT WHEN - 1 THEN" _
            & " ' Scaling Man right sextant.' ELSE '' END + CASE T .PRH_SUP_MANANT WHEN - 1 THEN ' Scaling Man ant sextant.' ELSE '' END + CASE T .PRH_SUP_MANLEFT" _
            & " WHEN - 1 THEN ' Scaling Man left sextant.' ELSE '' END + CASE T .PRH_POL_FULLMOUTH WHEN - 1 THEN ' Full mouth polished.' ELSE '' END + CASE" _
            & " T .PRH_POL_MAXRIGHT WHEN - 1 THEN ' Polish Max right sextant.' ELSE '' END + CASE T .PRH_POL_MAXANT WHEN - 1 THEN ' Polish Max ant sextant.'" _
            & " Else '' END + CASE T .PRH_POL_MAXLEFT WHEN - 1 THEN ' Polish Max left sextant.' ELSE '' END + CASE T .PRH_POL_MANRIGHT WHEN - 1 THEN ' Polish Man right sextant.'" _
            & " Else '' END + CASE T .PRH_POL_MANANT WHEN - 1 THEN ' Polish Man ant sextant.' ELSE '' END + CASE T .PRH_POL_MANLEFT WHEN - 1 THEN ' Polish Man left sextant.'" _
            & " Else '' END + CASE T .PRH_PROPHYLAXIS_COMMENTS WHEN '' THEN '' ELSE ' Prophylaxis Comments (' + T .PRH_PROPHYLAXIS_COMMENTS + ').'" _
            & " END + CASE T .PRV_IM_MAXMOUTHGUARD WHEN - 1 THEN ' Max mouthguard.' ELSE '' END + CASE T .PRV_IM_MANMOUTHGUARD WHEN - 1 THEN ' Man mouthguard.'" _
            & " Else '' END + CASE T .PRV_DL_MAXATHLETIC WHEN - 1 THEN ' Max Athletic mouthguard.' ELSE '' END + CASE T .PRV_DL_MANATHLETIC WHEN - 1 THEN" _
            & " ' Man Athletic mouthguard.' ELSE '' END END"
    Else
        'Oracle format
        sEncounterType = sEncounterType & vbCrLf & _
            "|| DECODE(T.PRH_PROPH_COMPLETE,-1,' Complete prophylaxis.')" & vbCrLf & _
            "|| DECODE(T.PRH_PROPH_INCOMPLETE,-1,' Incomplete prophylaxis.')" & vbCrLf & _
            "|| DECODE(T.PRH_PROPH_FLUORIDE_SNF2,-1,' SnF2.')" & vbCrLf & _
            "|| DECODE(T.PRH_PROPH_FLUORIDE_APF,-1,' APF.')" & vbCrLf & _
            "|| DECODE(T.PRH_PROPH_FLUORIDE_NAF,-1,' NaF.')" & vbCrLf & _
            "|| DECODE(T.PRH_PROPH_FLUORIDE_VARNISH,-1,' Fluoride varnish.')" & vbCrLf & _
            "|| DECODE(T.PRH_SUP_FULLMOUTH,-1,' Full mouth supragingival scaling.')" & vbCrLf & _
            "|| DECODE(T.PRH_SUP_MAXRIGHT,-1,' Scaling Max right sextant.')" & vbCrLf & _
            "|| DECODE(T.PRH_SUP_MAXANT,-1,' Scaling Max ant sextant.')" & vbCrLf & _
            "|| DECODE(T.PRH_SUP_MAXLEFT,-1,' Scaling Man right sextant.')" & vbCrLf & _
            "|| DECODE(T.PRH_SUP_MANRIGHT,-1,' Scaling Man right sextant.')" & vbCrLf & _
            "|| DECODE(T.PRH_SUP_MANANT,-1,' Scaling Man ant sextant.')" & vbCrLf & _
            "|| DECODE(T.PRH_SUP_MANLEFT,-1,' Scaling Man left sextant.')" & vbCrLf & _
            "|| DECODE(T.PRH_POL_FULLMOUTH,-1,' Full mouth polished.')" & vbCrLf & _
            "|| DECODE(T.PRH_POL_MAXRIGHT,-1,' Polish Max right sextant.')" & vbCrLf & _
            "|| DECODE(T.PRH_POL_MAXANT,-1,' Polish Max ant sextant.')" & vbCrLf & _
            "|| DECODE(T.PRH_POL_MAXLEFT,-1,' Polish Max left sextant.')" & vbCrLf & _
            "|| DECODE(T.PRH_POL_MANRIGHT,-1,' Polish Man right sextant.')" & vbCrLf & _
            "|| DECODE(T.PRH_POL_MANANT,-1,' Polish Man ant sextant.')" & vbCrLf & _
            "|| DECODE(T.PRH_POL_MANLEFT,-1,' Polish Man left sextant.')"
        sEncounterType = sEncounterType & vbCrLf & _
            "|| DECODE(T.PRH_PROPHYLAXIS_COMMENTS,null,'', ' Prophylaxis Comments (' || T.PRH_PROPHYLAXIS_COMMENTS || ').')" & vbCrLf & _
            "|| DECODE(T.PRV_IM_MAXMOUTHGUARD,-1,' Max mouthguard.')" & vbCrLf & _
            "|| DECODE(T.PRV_IM_MANMOUTHGUARD,-1,' Man mouthguard.')" & vbCrLf & _
            "|| DECODE(T.PRV_DL_MAXATHLETIC,-1,' Max Athletic mouthguard.')" & vbCrLf & _
            "|| DECODE(T.PRV_DL_MANATHLETIC,-1,' Man Athletic mouthguard.')"
    End If
    '-- From
    sAppendFrom = ", TEMPL_PREVENTIVE T "
    '-- Where
    sAppendWhere = "AND   D.DENTAL_ENCOUNTER_ID = T.DENTAL_ENCOUNTER_ID AND (T.PRH_PROPH_COMPLETE = - 1 " & vbCrLf & _
            "   OR T.PRH_PROPH_INCOMPLETE = - 1 " & vbCrLf & _
            "   OR T.PRH_PROPH_FLUORIDE_SNF2 = - 1 " & vbCrLf & _
            "   OR T.PRH_PROPH_FLUORIDE_APF = - 1 " & vbCrLf & _
            "   OR T.PRH_PROPH_FLUORIDE_NAF = - 1 " & vbCrLf & _
            "   OR T.PRH_PROPH_FLUORIDE_VARNISH = - 1 " & vbCrLf & _
            "   OR T.PRH_SUP_FULLMOUTH = - 1 " & vbCrLf & _
            "   OR T.PRH_SUP_MAXRIGHT = - 1 " & vbCrLf & _
            "   OR T.PRH_SUP_MAXANT = - 1 " & vbCrLf & _
            "   OR T.PRH_SUP_MAXLEFT = - 1 " & vbCrLf & _
            "   OR T.PRH_SUP_MANRIGHT = - 1 " & vbCrLf & _
            "   OR T.PRH_SUP_MANANT = - 1 " & vbCrLf & _
            "   OR T.PRH_SUP_MANLEFT = - 1 " & vbCrLf & _
            "   OR T.PRH_POL_FULLMOUTH = - 1 " & vbCrLf & _
            "   OR T.PRH_POL_MAXRIGHT = - 1 " & vbCrLf & _
            "   OR T.PRH_POL_MAXANT = - 1 " & vbCrLf & _
            "   OR T.PRH_POL_MAXLEFT = - 1 " & vbCrLf & _
            "   OR T.PRH_POL_MANRIGHT = - 1 " & vbCrLf & _
            "   OR T.PRH_POL_MANANT = - 1 " & vbCrLf & _
            "   OR T.PRH_POL_MANLEFT = - 1 " & vbCrLf & _
            "   OR T.PRH_PROPHYLAXIS_COMMENTS is not null " & vbCrLf & _
            "   OR T.PRV_IM_MAXMOUTHGUARD = - 1 " & vbCrLf & _
            "   OR T.PRV_IM_MANMOUTHGUARD = - 1 " & vbCrLf & _
            "   OR T.PRV_DL_MAXATHLETIC = - 1 " & vbCrLf & _
            "   OR T.PRV_DL_MANATHLETIC = - 1) "
    '-- Sql for Preventive
    sSqlPreventive = BuildSql(sEncounterType, sAppendFrom, sAppendWhere)
           
    '--------------------------------------------------------------------------------------------
    '5.1 Treatment Type (Debridement, Maintenance, Root Planing)
    'Source Table: APPOINTMENT A, DENTAL_ENCOUNTER D, TEMPL_PREVENTIVE T
    'Note:  Preventive (Hygiene) Tab.
    '       Check Debridement, Maintenance, Root Planing
    '--------------------------------------------------------------------------------------------
    '-- Select
'    sEncounterType = "'Treatment (Debridement, Maintenance, Root Planing)'"
    sEncounterType = "'Treatment -'"
    
    If mobjCHCSIIConn.UseSQLServer Then
        'SQL Server format
        sEncounterType = sEncounterType & " + Case T.HG_SCALING_MAXRIGHT_1To3 when -1 then ' Max right qudrant 1-3 teeth.' ELSE '' END " & vbCrLf & _
            "+ Case T.HG_SCALING_MAXRIGHT_4Plus when -1 then ' Max right qudrant 4+ teeth.' ELSE '' END " & vbCrLf & _
            "+ Case T.HG_SCALING_MAXLEFT_1To3 when -1 then ' Max left qudrant 1-3 teeth.' ELSE '' END " & vbCrLf & _
            "+ Case T.HG_SCALING_MAXLEFT_4Plus when -1 then ' Max left qudrant 4+ teeth.' ELSE '' END " & vbCrLf & _
            "+ Case T.HG_SCALING_MANRIGHT_1To3 when -1 then ' Man right qudrant 1-3 teeth.' ELSE '' END " & vbCrLf & _
            "+ Case T.HG_SCALING_MANRIGHT_4Plus when -1 then ' Man right qudrant 4+ teeth.' ELSE '' END " & vbCrLf & _
            "+ Case T.HG_SCALING_MANLEFT_1To3 when -1 then ' Man left qudrant 1-3 teeth.' ELSE '' END " & vbCrLf & _
            "+ Case T.HG_SCALING_MANLEFT_4Plus when -1 then ' Man left qudrant 4+ teeth.' ELSE '' END " & vbCrLf & _
            "+ Case T.HG_DM_FULLMOUTH when -1 then ' Full mouth debridement.' ELSE '' END " & vbCrLf & _
            "+ Case T.HG_DM_PERIO when -1 then ' Periodontal maintenance.' ELSE '' END " & vbCrLf & _
            "+ Case T.HG_DM_IMPLANT when -1 then ' Implant maintenance.' ELSE '' END "
    Else
        'Oracle format
        sEncounterType = sEncounterType & " || DECODE(T.HG_SCALING_MAXRIGHT_1To3,-1,' Max right qudrant 1-3 teeth.') " & vbCrLf & _
            "|| DECODE(T.HG_SCALING_MAXRIGHT_4Plus,-1,' Max right qudrant 4+ teeth.') " & vbCrLf & _
            "|| DECODE(T.HG_SCALING_MAXLEFT_1To3,-1,' Max left qudrant 1-3 teeth.') " & vbCrLf & _
            "|| DECODE(T.HG_SCALING_MAXLEFT_4Plus,-1,' Max left qudrant 4+ teeth.') " & vbCrLf & _
            "|| DECODE(T.HG_SCALING_MANRIGHT_1To3,-1,' Man right qudrant 1-3 teeth.') " & vbCrLf & _
            "|| DECODE(T.HG_SCALING_MANRIGHT_4Plus,-1,' Man right qudrant 4+ teeth.') " & vbCrLf & _
            "|| DECODE(T.HG_SCALING_MANLEFT_1To3,-1,' Man left qudrant 1-3 teeth.') " & vbCrLf & _
            "|| DECODE(T.HG_SCALING_MANLEFT_4Plus,-1,' Man left qudrant 4+ teeth.') " & vbCrLf & _
            "|| DECODE(T.HG_DM_FULLMOUTH,-1,' Full mouth debridement.')" & vbCrLf & _
            "|| DECODE(T.HG_DM_PERIO,-1,' Periodontal maintenance.')" & vbCrLf & _
            "|| DECODE(T.HG_DM_IMPLANT,-1,' Implant maintenance.')"
    End If
    '-- From
    sAppendFrom = ", TEMPL_PREVENTIVE T "
    '-- Where
    sAppendWhere = "AND   D.DENTAL_ENCOUNTER_ID = T.DENTAL_ENCOUNTER_ID " & vbCrLf & _
            "AND   (" _
            & " T.HG_SCALING_MAXRIGHT_1To3 = - 1 " & vbCrLf & _
            "   OR T.HG_SCALING_MAXLEFT_1To3 = - 1 " & vbCrLf & _
            "   OR T.HG_SCALING_MANRIGHT_1To3 = - 1 " & vbCrLf & _
            "   OR T.HG_SCALING_MANLEFT_1To3 = - 1 " & vbCrLf & _
            "   OR T.HG_SCALING_MAXRIGHT_4PLUS = - 1 " & vbCrLf & _
            "   OR T.HG_SCALING_MAXLEFT_4PLUS = - 1 " & vbCrLf & _
            "   OR T.HG_SCALING_MANRIGHT_4PLUS = - 1 " & vbCrLf & _
            "   OR T.HG_SCALING_MANLEFT_4PLUS = - 1 " & vbCrLf & _
            "   OR T.HG_DM_FULLMOUTH = - 1 " & vbCrLf & _
            "   OR T.HG_DM_PERIO = - 1 " & vbCrLf & _
            "   OR T.HG_DM_IMPLANT = - 1)"
    '-- Sql for Treatment Hygiene
    sSqlTxHygiene = BuildSql(sEncounterType, sAppendFrom, sAppendWhere)
    
    '--------------------------------------------------------------------------------------------
    '5.2 Treatment Type (Completed Chartable)
    'Source Table: APPOINTMENT A, DENTAL_ENCOUNTER D, DENTAL_ENCOUNTER_TX T,
    '   DENTAL_ENCOUNTER_TXS_SS ETXS, DENTAL_ENCOUNTER_TEETH ET, DENTAL_ENC_SURFACE_SET ESS
    'Note:  Any completed chartable status in DENTAL_ENCOUNTER_TX table
    '--------------------------------------------------------------------------------------------
    '-- Select
'    sEncounterType = "'Treatment (Completed Location -  ' || ET.TOOTH_NAME || ')'"
    sEncounterType = "ADA.ADA_Definition || ': Treatment Location -  ' || ET.TOOTH_NAME "
    '-- From
    sAppendFrom = ", DENTAL_ENCOUNTER_TX T, " & vbCrLf & _
        "   DENTAL_ENCOUNTER_TXS_SS ETXS, " & vbCrLf & _
        "   DENTAL_ENCOUNTER_TEETH ET, " & vbCrLf & _
        "   DENTAL_ENC_SURFACE_SET ESS,  " & vbCrLf & _
        "   TX_ADA_CODE_DEFINITION ADA " & vbCrLf
    '-- Where
    sAppendWhere = "AND   T.STATUS_ID = 102" & vbCrLf & _
                   "AND   D.DENTAL_ENCOUNTER_ID = T.MODIFIED_IN_DENTAL_ENCOUNTER " & vbCrLf & _
                   "AND   T.TX_CODE_ID = ETXS.TX_CODE_ID " & vbCrLf & _
                   "AND   ESS.SURFACE_SET_ID = ETXS.SURFACE_SET_ID " & vbCrLf & _
                   "AND   ADA.TX_ADA_CODE_ID = T.TX_ADA_CODE_ID " & vbCrLf & _
                   "AND   ET.DENTAL_ENCOUNTER_TEETH_ID = ESS.DENTAL_ENCOUNTER_TEETH_ID " & vbCrLf & _
                   "AND   T.CHARTABLE = -1 "
    '-- Sql for Treatment Chartable
    sSqlTxChartable = BuildSql(sEncounterType, sAppendFrom, sAppendWhere)
   
    '--------------------------------------------------------------------------------------------
    '5.3 Treatment Type (Completed UnChartable)
    'Source Table: APPOINTMENT A, DENTAL_ENCOUNTER D, DENTAL_ENCOUNTER_TX T,
    '   NOTE_TXS_LOCATION NTXSL, NOTE_LOCATION NL, V_TX_SUB_LOCATION V
    'Note:  Any completed Unchartable status in DENTAL_ENCOUNTER_TX table
    '--------------------------------------------------------------------------------------------
    '-- Select
'    sEncounterType = "'Treatment (Completed Location -  ' || NL.LOCATION_NAME || ')'"
    sEncounterType = "'Treatment Location -  ' || NL.LOCATION_NAME "
    '-- From
'    sAppendFrom = ", DENTAL_ENCOUNTER_TX T, " & vbCrLf & _
'        "   NOTE_TXS_LOCATION NTXSL, " & vbCrLf & _
'        "   NOTE_LOCATION NL, " & vbCrLf & _
'        "   V_TX_SUB_LOCATION V "
    sAppendFrom = ", DENTAL_ENCOUNTER_TX T, " & vbCrLf & _
        "   NOTE_TXS_LOCATION NTXSL, " & vbCrLf & _
        "   NOTE_LOCATION NL "
    '-- Where

     sAppendWhere = "AND   T.STATUS_ID = 102" & vbCrLf & _
                   "AND   D.DENTAL_ENCOUNTER_ID = T.MODIFIED_IN_DENTAL_ENCOUNTER " & vbCrLf & _
                   "AND   T.TX_CODE_ID = NTXSL.TX_CODE_ID " & vbCrLf & _
                   "AND   NL.LOCATION_ID = NTXSL.LOCATION_ID " & vbCrLf & _
                   "AND   T.CHARTABLE = 0"
    '-- Sql for Treatment Chartable
    sSqlTxUnChartable = BuildSql(sEncounterType, sAppendFrom, sAppendWhere)
    
    '------------------------------------------------------------------
    '6 Unknown Encounter Type
    'Note: Encounters not touched by the above filters.
    '------------------------------------------------------------------
    sEncounterType = "'Unknown Encounter Type'"
    sAppendFrom = ""
    sAppendWhere = ""

    sSQLUnknown = BuildSql(sEncounterType, sAppendFrom, sAppendWhere)
    '--------------------------------------------------------------------------------------------
    '6 NULL Type
    'Source Table: APPOINTMENT A, DENTAL_ENCOUNTER D
    'Checking Table: TEMPL_EXAMSOAP, TEMPL_EXAMGENERAL, TEMPL_PREVENTIVE, DENTAL_ENCOUNTER_TX
    'Note:  Encounters exist in Source tables but Checking tables)
    '--------------------------------------------------------------------------------------------
    '-- Select
    sEncounterType = "''"
    '-- From
    sAppendFrom = ""
    '-- Where
    sAppendWhere = "AND   (D.DENTAL_ENCOUNTER_ID NOT IN " & vbCrLf & _
                   "       (SELECT DENTAL_ENCOUNTER_ID FROM TEMPL_EXAMSOAP) " & vbCrLf & _
                   "    OR D.DENTAL_ENCOUNTER_ID NOT IN " & vbCrLf & _
                   "       (SELECT DENTAL_ENCOUNTER_ID FROM TEMPL_EXAMGENERAL) " & vbCrLf & _
                   "    OR D.DENTAL_ENCOUNTER_ID NOT IN " & vbCrLf & _
                   "       (SELECT DENTAL_ENCOUNTER_ID FROM TEMPL_PREVENTIVE) " & vbCrLf & _
                   "    OR D.DENTAL_ENCOUNTER_ID NOT IN " & vbCrLf & _
                   "       (SELECT MODIFIED_IN_DENTAL_ENCOUNTER FROM DENTAL_ENCOUNTER_TX) " & vbCrLf & _
                   "      )"
    '-- Sql for Treatment Chartable
    sSqlNull = BuildSql(sEncounterType, sAppendFrom, sAppendWhere)
                  
    '--------------------------------------------------------------------------------------------
    '-- Final Sql
    '--------------------------------------------------------------------------------------------
    ProcessEncounterType = sSQLUnknown & " UNION " & vbCrLf & sSqlEmergency & vbCrLf & _
            "UNION " & vbCrLf & _
            sSqlExam & vbCrLf & _
            "UNION " & vbCrLf & _
            sSqlEval & vbCrLf & _
            "UNION " & vbCrLf & _
            sSqlPreventive & vbCrLf & _
            "UNION " & vbCrLf & _
            sSqlTxHygiene & vbCrLf & _
            "UNION " & vbCrLf & _
            sSqlTxChartable & vbCrLf & _
            "UNION " & vbCrLf & _
            sSqlTxUnChartable & vbCrLf & "ORDER BY 1 DESC "
'            "UNION " & vbCrLf & _
            sSqlNull & vbCrLf & "ORDER BY 1 DESC "

End Function

'***************************************************************************************************
' Name:         ProcessWorkload
' Description:  Build Sql Statement for process Workload
' Target Table:  N/A
' Source Table:  APPOINTMENT A, DENTAL_ENCOUNTER D1,DENTAL_ENCOUNTER_TX D2,TX_ADA_CODE_DEFINITION D3
'           and TEMPL_WORKLOAD D2
' Parameter:    spEncounterType - Encounter Type
' Return:       String
' Called By:    SqlProcess
' Histary:
'   Developer:      Date:       Description
'   ----------      ---------   -----------------------------------------------
'   Wilson Sun      01/22/2002  Removed TEMPL_WORKLOAD.WORKLOAD_CODE and
'                               added   TEMPL_WORKLOAD.TX_ADA_CODE_ID in Database
'                               Code changed accordingly with DB change.
'                               Commented data from DENTAL_ENCOUNTER_TX sql statement.
'   Wilson sun      01/08/2002  Created.
'**************************************************************************************************
Private Function ProcessWorkload() As String
    Dim sSQL As String
    
    Dim sPatientNameSSN As String
    Dim sSSN As String
    Dim sFormatedSSN As String
    
    sSSN = Trim(mobjPatient.SSN)
    sFormatedSSN = Left(sSSN, 3) & "-" & Mid(sSSN, 4, 2) & "-" & Right(sSSN, 4)
    sPatientNameSSN = "'" & mobjPatient.Name & "' as PATIENT_NAME,'" & sFormatedSSN & "' as SSN"
    
    '-- select
     sSQL = _
        "select " & sPatientNameSSN & ", " & mobjCHCSIIConn.FormatDateForDBSelect("A.APPT_DATE_TIME", DD_MON_YYYY) & " as APP_DATE," & vbCrLf & _
        "       '" & RQ(guCurEncInfo.ProviderFullName) & "' as PROVIDER_NAME," & vbCrLf & _
        "       D3.ADA_CODE ||' '|| D3.ADA_DEFINITION as DESCRIPTION," & vbCrLf & _
        "       sum(MULTIPLIER) as TOTAL, DWV, Provider.Name as Given_BY "
    If mobjCHCSIIConn.UseSQLServer Then
        '-- from
        sSQL = sSQL & " FROM APPOINTMENT A INNER JOIN" _
            & " DENTAL_ENCOUNTER D1 ON A.PATIENT_UNIT_NUMBER = D1.PATIENT_UNIT_ID AND A.FACILITY_NCID = D1.ENCOUNTER_FACILITY_NCID AND" _
            & " A.ENCOUNTER_NUMBER = D1.ENCOUNTER_NUMBER INNER JOIN" _
            & " TEMPL_WORKLOAD D2 ON D1.DENTAL_ENCOUNTER_ID = D2.DENTAL_ENCOUNTER_ID INNER JOIN" _
            & " TX_ADA_CODE_DEFINITION D3 ON D2.TX_ADA_CODE_ID = D3.TX_ADA_CODE_ID LEFT OUTER JOIN" _
            & " PROVIDER ON D2.PROVIDER_ID = PROVIDER.NCID" _
            & " Where (D2.DELETED = 0) And (A.ENCOUNTER_STATUS <> 4) And (A.ENCOUNTER_NUMBER = " & guCurEncInfo.EncounterID & ")" _
            & " GROUP BY A.APPT_DATE_TIME, D3.ADA_CODE + ' ' + D3.ADA_DEFINITION, D3.DWV, PROVIDER.NAME"
    Else
        '-- from
        sSQL = sSQL & vbCrLf & _
            "from APPOINTMENT A," & vbCrLf & _
            "     DENTAL_ENCOUNTER D1," & vbCrLf & _
            "     TEMPL_WORKLOAD D2," & vbCrLf & _
            "     TX_ADA_CODE_DEFINITION D3, " & _
            "     PROVIDER "
        '-- where
        sSQL = sSQL & vbCrLf & _
            "where A.PATIENT_UNIT_NUMBER = D1.PATIENT_UNIT_ID " & vbCrLf & _
            "and   A.FACILITY_NCID = D1.ENCOUNTER_FACILITY_NCID" & vbCrLf & _
            "and   A.ENCOUNTER_NUMBER = D1.ENCOUNTER_NUMBER" & vbCrLf & _
            "and   D1.DENTAL_ENCOUNTER_ID = D2.DENTAL_ENCOUNTER_ID" & vbCrLf & _
            "and   D2.TX_ADA_CODE_ID = D3.TX_ADA_CODE_ID" & vbCrLf & _
            "and   D2.DELETED = 0" & vbCrLf & _
            "and   A.ENCOUNTER_STATUS <> 4" & vbCrLf & _
            "and   A.ENCOUNTER_NUMBER = " & guCurEncInfo.EncounterID & vbCrLf & _
            "and   D2.Provider_ID = Provider.NCID(+) " & _
            "group by A.APPT_DATE_TIME, D3.ADA_CODE ||' '|| D3.ADA_DEFINITION, DWV, Provider.Name "
    End If
    
    ProcessWorkload = sSQL
End Function
'***************************************************************************************************
' Name:         ProcessLabCase
' Description:  Build Sql Statement for process Lab Case
' Target Table:  N/A
' Source Table:  APPOINTMENT ap, DENTAL_ENCOUNTER a, LAB_CASE b, LAB_TX_SPECIFICATIONS c,
'               DENTAL_ENCOUNTER_TX d, DENTAL_ENCOUNTER_TXS_SS e, DENTAL_ENCOUNTER_TEETH f,
'               DENTAL_ENC_SURFACE_SET g, TX_ADA_CODE_DEFINITION h, CLINIC cl, CLINIC_PROVIDER cp
'           and NOTE_TXS_LOCATION e, NOTE_LOCATION f
' Parameter:    spEncounterType - Encounter Type
' Return:       String
' Called By:    SqlProcess
' Histary:
'   Developer:      Date:       Description
'   ----------      ---------   -----------------------------------------------
'   Wilson sun      01/08/2002  Created
'**************************************************************************************************
Private Function ProcessLabCase() As String
    Dim sSQL As String
    
    '-- select
    sSQL = _
        "SELECT cl.Name, p.Name as Provider_Name, ap.appt_date_time," & vbCrLf & _
        "       f.Tooth_Name as Tooth_Location, h.ADA_Definition, " & vbCrLf & _
        "       ap.Encounter_Number, b.Lab_Case_Number, h.ADA_Code," & vbCrLf & _
        "       ap.Facility_Ncid, ap.Clinic_Ncid, ap.Clinician_Ncid "
    '-- from
    sSQL = sSQL & vbCrLf & _
        "FROM APPOINTMENT ap," & vbCrLf & _
        "     DENTAL_ENCOUNTER a," & vbCrLf & _
        "     LAB_CASE b," & vbCrLf & _
        "     LAB_TX_SPECIFICATIONS c," & vbCrLf & _
        "     DENTAL_ENCOUNTER_TX d," & vbCrLf & _
        "     DENTAL_ENCOUNTER_TXS_SS e," & vbCrLf & _
        "     DENTAL_ENCOUNTER_TEETH f," & vbCrLf & _
        "     DENTAL_ENC_SURFACE_SET g," & vbCrLf & _
        "     TX_ADA_CODE_DEFINITION h," & vbCrLf & _
        "     CLINIC cl," & vbCrLf & _
        "      provider p "
    '-- where
    sSQL = sSQL & vbCrLf & _
        "WHERE ap.FACILITY_NCID = a.ENCOUNTER_FACILITY_NCID" & vbCrLf & _
        "AND   ap.CLINIC_NCID = cl.NCID" & vbCrLf & _
        "AND   p.NCID = b.PROVIDERID" & vbCrLf & _
        "AND   ap.ENCOUNTER_NUMBER = a.ENCOUNTER_NUMBER" & vbCrLf & _
        "AND   ap.PATIENT_UNIT_NUMBER = a.PATIENT_UNIT_ID" & vbCrLf & _
        "AND   a.Patient_Unit_ID = " & mobjPatient.UnitNumber & vbCrLf & _
        "AND   a.Dental_Encounter_ID = b.Dental_Encounter_ID" & vbCrLf & _
        "AND   b.Lab_Case_ID = c.Lab_Case_ID" & vbCrLf & _
        "AND   c.Tx_Code_ID = d.Tx_Code_ID" & vbCrLf & _
        "AND   d.Tx_Code_ID = e.Tx_Code_ID" & vbCrLf & _
        "AND   h.Tx_ADA_Code_ID = d.Tx_ADA_Code_ID" & vbCrLf & _
        "AND   h.LabFlag = -1" & vbCrLf & _
        "AND   f.Dental_Encounter_Teeth_ID = g.Dental_Encounter_Teeth_ID" & vbCrLf & _
        "AND   e.Surface_Set_ID = g.Surface_Set_ID" & vbCrLf  '& _
        "AND  (d.Status_ID = 101 OR d.Status_ID = 102) "
    '--Union select
     sSQL = sSQL & vbCrLf & "UNION ALL " & vbCrLf & _
        "SELECT cl.Name, p.Name as Provider_Name, ap.appt_date_time," & vbCrLf & _
        "       f.Location_Name as Tooth_Location, h.ADA_Definition," & vbCrLf & _
        "       ap.Encounter_Number, b.Lab_Case_Number, h.ADA_Code," & vbCrLf & _
        "       ap.Facility_Ncid, ap.Clinic_Ncid, ap.Clinician_Ncid "
    '-- from
    sSQL = sSQL & vbCrLf & _
        "FROM APPOINTMENT ap, " & vbCrLf & _
        "     DENTAL_ENCOUNTER a, " & vbCrLf & _
        "     LAB_CASE b, " & vbCrLf & _
        "     LAB_TX_SPECIFICATIONS c, " & vbCrLf & _
        "     DENTAL_ENCOUNTER_TX d, " & vbCrLf & _
        "     NOTE_TXS_LOCATION e, " & vbCrLf & _
        "     NOTE_LOCATION f, " & vbCrLf & _
        "     TX_ADA_CODE_DEFINITION h, " & vbCrLf & _
        "     CLINIC cl, " & vbCrLf & _
        "      provider p "
    '-- where
    sSQL = sSQL & vbCrLf & _
    "WHERE ap.FACILITY_NCID = a.ENCOUNTER_FACILITY_NCID " & vbCrLf & _
    "AND   ap.CLINIC_NCID = cl.NCID " & vbCrLf & _
    "AND   p.NCID = b.PROVIDERID " & vbCrLf & _
    "AND   ap.ENCOUNTER_NUMBER = a.ENCOUNTER_NUMBER " & vbCrLf & _
    "AND   a.Patient_Unit_ID = " & mobjPatient.UnitNumber & vbCrLf & _
    "AND   a.Dental_Encounter_ID = b.Dental_Encounter_ID " & vbCrLf & _
    "AND   b.Lab_Case_ID = c.Lab_Case_ID " & vbCrLf & _
    "AND   c.Tx_Code_ID = d.Tx_Code_ID " & vbCrLf & _
    "AND   d.Tx_Code_ID = e.Tx_Code_ID " & vbCrLf & _
    "AND   e.Location_ID = f.Location_ID " & vbCrLf & _
    "AND   h.Tx_ADA_Code_ID = d.Tx_ADA_Code_ID " & vbCrLf & _
    "AND   h.LabFlag = -1 " '& vbCrLf & _
    "AND  (d.Status_ID = 101 OR d.Status_ID = 102) " & vbCrLf & _
    "ORDER BY 3 DESC"
    sSQL = sSQL & "ORDER BY 3 DESC"
    
    ProcessLabCase = sSQL
End Function

'**********************************************************************************
'Name:          OracleDate
'Description:   Return TO_DATE oracle function and format
'Parameters:    vDate - to be converted value
'Called by:     LoadDxTxDataSql
'Developer:     Wilson Sun
'Create Date:   02/12/2001
'**********************************************************************************
'Private Function OracleDate(ByVal vDate As Variant) As String
'    OracleDate = "TO_DATE('" & Format$(vDate, gsDATE_FORMAT_VB) & "', '" & gsDATE_FORMAT_ORA & "')"
'End Function

'********************************************************************************
'Sub:FilterNotesRS
'Author:KDunne
'Date:3/19/2002
'Purpose: Filters grsNotes to remove entries of "Unknown Encounter Type" for
'         Encounters already in the recordset.
'Notes: Change was made to query that populates grsNotes to to expand it to
'       bring back all encounters, not just encounters with a known encounter
'       type. Some of the unknown encounter types belong to encounters that have
'       at least one known encounter type. We want to remove entries of Unknown
'       Encounter Type for known encounters. The SQL Query to do this contained
'       too many outer joins, filtering the recordset was chosen for performance
'       reasons.
'*********************************************************************************
Private Sub FilterNotesRS(objRS As ADODB.Recordset)
    Dim strFilterIds As String
    Dim strFilter As String
    Dim arrEncNums() As String
    Dim intCounter As Integer
    
    
    strFilterIds = Empty
    strFilter = Empty


    'First, loop through existing recordset and get all of the dental encounter
    'ids that have at least one 'Unknown ENcounter Type. Store DENTAL_ENCOUNTER_ID
    'in a string variable.
    
    With objRS
        .Filter = " ENCOUNTER_TYPE = 'Unknown Encounter Type' "
        Do While Not .EOF
            strFilterIds = strFilterIds & ", " & .Fields("DENTAL_ENCOUNTER_ID")
            .MoveNext
        Loop
    End With 'objRSNotes
    

    'Remove the first ", " from the string of DENTAL_ENCOUNTER_ID's. Next,
    'convert values into an array.
    strFilterIds = Replace(strFilterIds, ",", "", , 1)
    arrEncNums = Split(strFilterIds, ",")
        
        
    'Loop through array of DENTAL_ENCOUNTER_ID's. For each element in the array,
    'Filter the recordset on that DENTAL_ENCOUNTER_ID. Add DENTAL_ENCOUNTER_ID to
    'our list of DENTAL_ENCOUNTER_ID's to be included in the final filter. If the
    'record count is greater than 1, then we want to exclude the record which
    'contains the DENTAL_ENCOUNTER_ID and 'Unknown Encounter Type
    '
    For intCounter = 0 To UBound(arrEncNums)
        objRS.Filter = "DENTAL_ENCOUNTER_ID = " & arrEncNums(intCounter)
    
        strFilter = strFilter & " OR " & " (DENTAL_ENCOUNTER_ID = " & arrEncNums(intCounter)
        If objRS.RecordCount > 1 Then
            strFilter = strFilter & " AND ENCOUNTER_TYPE <> 'Unknown Encounter Type' )"
        Else
            strFilter = strFilter & ")"
        End If
    Next
    
    'Build and set the final Filter String
    strFilter = Replace(strFilter, " OR ", "", , 1)
    objRS.Filter = strFilter

End Sub
'***************************************************************************************************
' Name:         BuildSql
' Description:  Build complete SQL statement. This function for retrieving Encounter Type data only
' Parameter:    spEncounterType - Encounter Type
'               spAppendFrom    - Append to main From clause if any
'               spAppendWhere   - Append to main Where clause if any
' Return:       String
' Called By:    ProcessEncounterType
' Histary:
'   Developer:      Date:       Description
'   ----------      ---------   -----------------------------------------------
'   Wilson sun      01/07/2002  Created
'**************************************************************************************************
Private Function BuildSql(ByVal spEncounterType As String, _
                        ByVal spAppendFrom As String, _
                        ByVal spAppendWhere As String) As String
    Dim sMainFrom As String
    Dim sMainWhere As String
    '-- Comm From
    sMainFrom = "FROM APPOINTMENT A, DENTAL_ENCOUNTER D"
    '-- Comm Where
    sMainWhere = "WHERE D.ENCOUNTER_FACILITY_NCID = A.FACILITY_NCID " & vbCrLf & _
                 "AND   D.PATIENT_UNIT_ID = A.PATIENT_UNIT_NUMBER " & vbCrLf & _
                 "AND   D.ENCOUNTER_NUMBER = A.ENCOUNTER_NUMBER " & vbCrLf & _
                 "AND   A.PATIENT_UNIT_NUMBER = " & mobjPatient.UnitNumber
    '-- Sql complete
    BuildSql = MainSelect(spEncounterType) & vbCrLf & _
            sMainFrom & spAppendFrom & vbCrLf & _
            sMainWhere & vbCrLf & spAppendWhere
End Function
'***************************************************************************************************
' Name:         MainSelect
' Description:  Build Common Select Clause
' Parameter:    spEncounterType - Encounter Type
' Return:       String
' Called By:    BuildSql
' Histary:
'   Developer:      Date:       Description
'   ----------      ---------   -----------------------------------------------
'   Wilson sun      01/07/2002  Created
'**************************************************************************************************
Private Function MainSelect(ByVal spEncounterType As String) As String
    MainSelect = "SELECT A.APPT_DATE_TIME, D.DRC," & spEncounterType & " as ENCOUNTER_TYPE," & vbCrLf & _
    "D.DENTAL_ENCOUNTER_ID, D.ENCOUNTER_NUMBER, A.FACILITY_NCID "
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

Public Function LoadLabCaseSqlSP(ByVal lngPatientUnitId As Long) As ADODB.Recordset

    Dim objCmd As CHCSII_Command.CHCSII_Cmd
    Dim lngMousePointer As VBRUN.MousePointerConstants
    
    Screen.MousePointer = vbHourglass
    On Error GoTo ErrHandler
    
    Set objCmd = New CHCSII_Command.CHCSII_Cmd
    objCmd.SPName = "DENT_DV_LABS"
    
    'We must add paramaters in the same order as the appear in the Stored Proc Signature
    Call objCmd.AddParam(False, False, "textString", lngPatientUnitId)
    'This Parameter is defined as a cursor, used for output
    Call objCmd.AddParam(True, True, "", "")
    
    Set LoadLabCaseSqlSP = mobjDas.OpenRecordsetSP(objCmd)


    ProcessRecordSet LoadLabCaseSqlSP
FuncExit:
    Set objCmd = Nothing
    Screen.MousePointer = lngMousePointer
    
    Exit Function
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "clsDVDataOps.LoadLabCaseSqlSP", App.Title, vbCritical
    GoTo FuncExit
End Function

Public Function LoadDxTxDataSqlSP(ByVal lngPatientUnitId As Long, ByVal strUserName As String) As ADODB.Recordset
    Dim objCmd As CHCSII_Command.CHCSII_Cmd
    Dim lngMousePointer As VBRUN.MousePointerConstants

    Screen.MousePointer = vbHourglass
    On Error GoTo ErrHandler
    
    'Make sure we remove any apostrophe's from the user name
    strUserName = RQ(strUserName)
    
    Set objCmd = New CHCSII_Command.CHCSII_Cmd
    objCmd.SPName = "DENT_DV_LOAD_DXTX_DATA"
    
    'We must add paramaters in the same order as the appear in the Stored Proc Signature
    Call objCmd.AddParam(False, False, "textString", lngPatientUnitId)
    Call objCmd.AddParam(False, False, "textString", strUserName)
    'This Parameter is defined as a cursor, used for output
    Call objCmd.AddParam(True, True, "", "")
    
     Set LoadDxTxDataSqlSP = mobjDas.OpenRecordsetSP(objCmd)
     
     ProcessRecordSet LoadDxTxDataSqlSP
FuncExit:
    Screen.MousePointer = lngMousePointer
    Set objCmd = Nothing
    Exit Function
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "clsDVDataOps.LoadDxTxDataSqlSP", App.Title, vbCritical
    GoTo FuncExit
End Function

Public Function LoadEncounterTypeSqlSP(ByVal lngPatientUnitId As Long) As ADODB.Recordset
    Dim objCmd As CHCSII_Command.CHCSII_Cmd
    Dim lngMousePointer As VBRUN.MousePointerConstants

    Screen.MousePointer = vbHourglass
    On Error GoTo ErrHandler
    
    Set objCmd = New CHCSII_Command.CHCSII_Cmd
    objCmd.SPName = "DENT_DV_LOAD_ENC"
    
    'We must add paramaters in the same order as the appear in the Stored Proc Signature
    Call objCmd.AddParam(False, False, "textString", lngPatientUnitId)
    'This Parameter is defined as a cursor, used for output
    Call objCmd.AddParam(True, True, "", "")
    
    Set LoadEncounterTypeSqlSP = mobjDas.OpenRecordsetSP(objCmd)

    ProcessRecordSet LoadEncounterTypeSqlSP
    
FuncExit:
    Set objCmd = Nothing
    Screen.MousePointer = lngMousePointer
    
    Exit Function
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "clsDVDataOps.LoadEncounterTypeSqlSP", App.Title, vbCritical
    GoTo FuncExit

End Function

Public Function LoadWorkloadSqlSP(ByVal strPatName As String, ByVal strPatSSN As String, _
            ByVal strUserName As String, ByVal lngEncNum As Long) As ADODB.Recordset
    Dim objCmd As CHCSII_Command.CHCSII_Cmd
    Dim strFormattedSSN As String
    Dim lngMousePointer As VBRUN.MousePointerConstants

    Screen.MousePointer = vbHourglass
    On Error GoTo ErrHandler
    
    'Make sure we remove any apostrophe's from the user name
    strUserName = RQ(strUserName)
    strPatName = RQ(strPatName)
    strFormattedSSN = Left(strPatSSN, 3) & "-" & Mid(strPatSSN, 4, 2) & "-" & Right(strPatSSN, 4)

    Set objCmd = New CHCSII_Command.CHCSII_Cmd
    objCmd.SPName = "DENT_DV_WORKLOAD"
    
    'We must add paramaters in the same order as the appear in the Stored Proc Signature
    Call objCmd.AddParam(False, False, "textString", strPatName)
    Call objCmd.AddParam(False, False, "textString", strFormattedSSN)
    Call objCmd.AddParam(False, False, "textString", strUserName)
    Call objCmd.AddParam(False, False, "textString", lngEncNum)
    'This Parameter is defined as a cursor, used for output
    Call objCmd.AddParam(True, True, "", "")
    
     Set LoadWorkloadSqlSP = mobjDas.OpenRecordsetSP(objCmd)
     
     ProcessRecordSet LoadWorkloadSqlSP
FuncExit:
    Screen.MousePointer = lngMousePointer
    Set objCmd = Nothing
    Exit Function
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "clsDVDataOps.LoadWorkloadSqlSP", App.Title, vbCritical
    GoTo FuncExit
End Function

Private Sub Class_Terminate()
    Set mobjDas = Nothing
    Set mobjShared = Nothing
    Set mobjCHCSIIConn = Nothing
    Set mobjCWComm = Nothing
    Set mobjSQLOps = Nothing
    
End Sub
'********************************************************************************
'Sub:ProcessRecordSet
'Author:KDunne
'Date:12/19/2002
'Purpose: Processes the recordset
'*********************************************************************************
Private Sub ProcessRecordSet(ByRef objRS As ADODB.Recordset)
    Dim x As Integer
    
    With objRS
        If .BOF And .EOF Then
            Exit Sub
        End If
        .MoveFirst
        
        Do While Not .EOF
            
            For x = 0 To (.Fields.Count - 1)
                If .Fields(x).Type = adDate Then
                    .Fields(x).Value = ConvertSPDate(.Fields(x).Value)
                End If
            Next
            .MoveNext
        Loop
    
    End With
End Sub

'*********************************************************************************
'Function:ConvertSPDate
'Author:KDunne
'Date:12/19/2002
'Purpose: Currently, CHCSII_DASsp is returning all dates in GMT time. We must convert
'         Them back to Local.
'*********************************************************************************
Private Function ConvertSPDate(ByVal dteDate As Date) As Date
    Dim dteReturn As Date
    
    dteReturn = mobjCWComm.GMTTOLocal(dteDate)
    
    ConvertSPDate = dteReturn
End Function

