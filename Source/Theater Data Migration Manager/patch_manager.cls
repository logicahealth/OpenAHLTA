VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "patch_manager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'-------------------------------------------------------------------------------
'--
'-- * Routine    : Patch Manager
'-- * Purpose    : API that manages the entire patch application process
'-- * Written by : B. Webb
'-- * Date       : 17/03/2006
'--
'-- Modifications
'-- Date       Programmer Summary of Changes
'-- ---------- ---------- ----------------------------------------------------
'-- DD/MM/YYYY
'-------------------------------------------------------------------------------

Option Explicit


Private mOraSession As Object
Public mOraDatabase As Object 'OracleInProcServer.OraDatabase
Private mobjPatches_NotApplied As patch_collection
Private mobjPatches_Applied As patch_collection
Private mobjPatches_Available As patch_collection
Private mobjDosCommandManager As dos_command_manager
Private mobjImpExpManager As imp_exp_manager
Private mobjOracleManager As oracle_manager
Private msRunPath As String
Private msCurrentStatement As String
Private msLatestError As String
Private miErrorCount As Long
Private msCurrentErrorNumber As Long
Private msCurrentPatch As String
Private mlSessionErrors As Long
Private miPatchesProcessed As Long
Private mlStatementsProcessed As Long

Private mstrExportFile As String
Private mstrImportFile As String
Private mbStopOnError As Boolean

Private mobjPatchLog As Log_Manager
Private mobjErrorLog As Log_Manager

Public Event NewStatement()
Public Event StatementProcessed()
Public Event PatchStarted(objPatch As patch)
Public Event PatchCompleted(objPatch As patch, iErrors As Long)
Public Event StatementError()
'Public Event ApplicationError()

Private mobjParseDefinitions As parser_def_collection

Private mbWaitForEventConfirmation As Boolean
Private mbEventConfirmed As Boolean
Private mbLogin As Boolean
Private mbQuitOnError As Boolean
Private mbHaltProcessing As Boolean

Private msHost As String
Private msUID As String
Private msPassword As String
Private msConnectionString As String

Private mbSessionError As Boolean
Private mbExecute_Statements As Boolean
Private msPatchPath As String

'SF added Ignore Oracle Errors
Private mOracleErrorsToIgnore As Collection
Private Sub Get_Import_Errors_To_Ignore(Errors As Collection)
On Error GoTo ErrHandler

    Dim sTemp As String
    Dim TempArray
    Dim objIniFile As New INI_File_Manager
    Dim i As Integer

    Set Errors = Nothing
    Set Errors = New Collection
    
    objIniFile.File_Name = App.Path & "\patch.ini"
    objIniFile.Application_Name = "Import Errors Ignore"
    i = 1
    Do
        sTemp = objIniFile.Get_Value(CStr(i))
        If sTemp <> "The_End" And sTemp <> "" Then
            Errors.Add sTemp
        End If
        i = i + 1
    Loop While sTemp <> "The_End"


GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "Patch_Manager", "Get_Import_Errors_To_Ignore", Err.Number, Err.Description
CleanUp:
    Set objIniFile = Nothing
End Sub
'-- SCR 188803 Webb - Check for Correct Service Pack (Version and patch id) prior to data migration
Public Function Correct_Version_Patch() As Boolean
On Error GoTo ErrHandler

    Dim sTemp As String
    Dim TempArray
    Dim objIniFile As New INI_File_Manager
    Dim objOraDynaset As Object 'OraDynaset
    Dim sql As String
    Dim version_id As String
    Dim patch_id As String

    Correct_Version_Patch = False
    
    '-- Get required Version and Patch Id from patch.ini file
    objIniFile.File_Name = JoinPathAndFileName(App.Path, "\patch.ini")
    objIniFile.Application_Name = "settings"
    sTemp = Trim(objIniFile.Get_Value("required_patch"))

    If sTemp = "" Then
        Correct_Version_Patch = True  '-- If there is no value in the ini file then continue
    ElseIf mbLogin Then
        TempArray = Split(sTemp, ",") '-- Split the Version and patch id from the string
        
        If UBound(TempArray) < 1 Then
            '-- Value in ini file formatted incorrectly
            MsgBox "The 'required_patch' section under the 'settings' section in the patch.ini file is formatted incorrectly.", vbCritical, "Patch.ini configuration error"
        Else
            '-- Check if version and patch id are in the version_patch table
            version_id = Trim(TempArray(0))
            patch_id = TempArray(1)
            sql = "select count(*) total from gems.version_patch where version_id='" & version_id & "' and patch_id= " & patch_id
            Set objOraDynaset = mOraDatabase.DbCreateDynaset(sql, 0&)
            
            If objOraDynaset.Fields("total") > 0 Then
                Correct_Version_Patch = True
            Else
                MsgBox "Cannot perform data migration because the database is not up-to-date." & vbNewLine & vbNewLine & "Version: '" & version_id & "' Patch Id: '" & patch_id & "' and all preceeding patches must be applied prior to database migration.", vbCritical, "Latest Patches Not Applied"
            End If
        End If
    End If

GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "Patch_Manager", "Correct_Version_Patch", Err.Number, Err.Description
CleanUp:
    Set objIniFile = Nothing
    Set objOraDynaset = Nothing
End Function
Private Sub Get_Oracle_Errors_To_Ignore(Errors As Collection)
On Error GoTo ErrHandler

    Dim sTemp As String
    Dim TempArray
    Dim objIniFile As New INI_File_Manager
    Dim i As Integer

    Set Errors = Nothing
    Set Errors = New Collection
    
    objIniFile.File_Name = App.Path & "\patch.ini"
    objIniFile.Application_Name = "Oracle Errors Ignore"
    i = 1
    Do
        sTemp = objIniFile.Get_Value(CStr(i))
        If sTemp <> "The_End" And sTemp <> "" Then
            Errors.Add sTemp, sTemp
        End If
        i = i + 1
    Loop While sTemp <> "The_End"


GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "Patch_Manager", "Get_Import_Errors_To_Ignore", Err.Number, Err.Description
CleanUp:
    Set objIniFile = Nothing
End Sub
Private Sub Get_Export_Errors_To_Ignore(Errors As Collection)
On Error GoTo ErrHandler

    Dim sTemp As String
    Dim TempArray
    Dim objIniFile As New INI_File_Manager
    Dim i As Integer

    Set Errors = Nothing
    Set Errors = New Collection
    
    objIniFile.File_Name = App.Path & "\patch.ini"
    objIniFile.Application_Name = "Export Errors Ignore"
    i = 1
    Do
        sTemp = objIniFile.Get_Value(CStr(i))
        If sTemp <> "The_End" And sTemp <> "" Then
            Errors.Add sTemp
        End If
        i = i + 1
    Loop While sTemp <> "The_End"


GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "Patch_Manager", "Get_Export_Errors_To_Ignore", Err.Number, Err.Description
CleanUp:
    Set objIniFile = Nothing
End Sub
Private Function GetResult(Comment As String, Error_Code As Long) As String
On Error GoTo ErrHandler

    If Comment = "" And Error_Code = 0 Then
        GetResult = "Success"
    ElseIf Error_Code <> 0 Then
        GetResult = "Failed"
    ElseIf Comment <> "" Then
        GetResult = "NA"
    End If

GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "Patch_Manager", "GetResult", Err.Number, Err.Description
CleanUp:
End Function

Public Property Get Quit_On_Error() As Boolean
    Quit_On_Error = mbQuitOnError
End Property

Public Property Let Quit_On_Error(vNewValue As Boolean)
    mbQuitOnError = vNewValue
End Property

Public Property Get Patches_Applied(Optional MigratePatchOnly As Boolean = False) As Variant
On Error GoTo ErrHandler

    Dim objPatch As patch
    Dim objOraDynaset As Object 'OraDynaset
    Dim sql As String

    Set mobjPatches_Applied = Nothing
    Set mobjPatches_Applied = New patch_collection
    Set Patches_Applied = mobjPatches_Applied

    If mbLogin Then
        sql = "select version_id, patch_id, apply_date, vp_comment from gems.version_patch order by 1 desc, 2 desc"
        Set objOraDynaset = mOraDatabase.DbCreateDynaset(sql, 0&)
        Do While (objOraDynaset.EOF = False)
            If MigratePatchOnly And UCase(objOraDynaset.Fields("version_id")) = "MIGRATE" Then
                'don't add it.
            Else
                mobjPatches_Applied.Add objOraDynaset.Fields("vp_comment"), "", "", objOraDynaset.Fields("patch_id"), objOraDynaset.Fields("version_id"), objOraDynaset.Fields("apply_date"), 0, objOraDynaset.Fields("version_id") & "|" & objOraDynaset.Fields("patch_id")
            End If
           objOraDynaset.MoveNext
        Loop
    End If

    Set Patches_Applied = mobjPatches_Applied

GoTo CleanUp
ErrHandler:
    If Err.Number <> 440 Then
        gobjErrManager.Err_Log "Patch_Manager", "Patches_Applied", Err.Number, Err.Description
    End If
CleanUp:
    Set objOraDynaset = Nothing
End Property
Public Property Get Patches_Available() As Variant
On Error GoTo ErrHandler

    Dim sTemp As String
    Dim TempArray
    Dim objIniFile As New INI_File_Manager
    Dim i As Integer

    Set mobjPatches_Available = Nothing
    Set mobjPatches_Available = New patch_collection
    
    objIniFile.File_Name = App.Path & "\patch.ini"
    objIniFile.Application_Name = "patches"
    i = 1
    Do
        sTemp = objIniFile.Get_Value(CStr(i))
        If sTemp <> "The_End" And sTemp <> "" Then
            TempArray = Split(sTemp, "~")
            mobjPatches_Available.Add TempArray(4), TempArray(3), TempArray(2), TempArray(1), TempArray(0), "", Trim(TempArray(0)) & "|" & Trim(TempArray(1))
        End If
        i = i + 1
    Loop While sTemp <> "The_End"

    Set Patches_Available = mobjPatches_Available

GoTo CleanUp
ErrHandler:
        gobjErrManager.Err_Log "Patch_Manager", "Patches_Available", Err.Number, Err.Description
CleanUp:
'    Set objFSO = Nothing
'    Set objTS = Nothing
End Property
Public Property Get xPatches_Available() As Variant
On Error GoTo ErrHandler

    Dim objFSO As Object 'FileSystemObject
    Dim objTS As Object 'TextStream
    Dim sTemp As String
    Dim TempArray

    Set mobjPatches_Available = Nothing
    Set mobjPatches_Available = New patch_collection
    Set objFSO = CreateObject("Scripting.FileSystemObject") 'New FileSystemObject
    Set objTS = objFSO.OpenTextFile(App.Path & "\patch.ini", 1, False) 'objFSO.OpenTextFile(App.Path & "\patch.ini", ForReading, False)
    
    Do While Not (objTS.AtEndOfStream)
        sTemp = Trim(objTS.ReadLine)
        If sTemp = "" Then
            Exit Do
        ElseIf Mid(sTemp, 1, 2) <> "--" Then
            TempArray = Split(sTemp, "~")
            mobjPatches_Available.Add TempArray(4), TempArray(3), TempArray(2), TempArray(1), TempArray(0), "", Trim(TempArray(0)) & "|" & Trim(TempArray(1))
        End If
    Loop
    objTS.Close
    Set Patches_Available = mobjPatches_Available

GoTo CleanUp
ErrHandler:
        gobjErrManager.Err_Log "Patch_Manager", "Patches_Available", Err.Number, Err.Description
CleanUp:
    Set objFSO = Nothing
    Set objTS = Nothing
End Property
Public Property Get Patches_Not_Applied(Optional MigratePatchOnly As Boolean = False) As Variant
On Error GoTo ErrHandler

    Dim sTemp As String
    Dim TempArray
    Dim objApplied As patch_collection
    Dim objPatch As patch
    Dim objIniFile As New INI_File_Manager
    Dim i As Integer
    
    

    Set objApplied = Me.Patches_Applied
    
    Set mobjPatches_NotApplied = Nothing
    Set mobjPatches_NotApplied = New patch_collection

    objIniFile.File_Name = App.Path & "\patch.ini"
    objIniFile.Application_Name = "patches"
    i = 1
    Do
        sTemp = objIniFile.Get_Value(CStr(i))
        If sTemp <> "The_End" And sTemp <> "" Then
            TempArray = Split(sTemp, "~")
            'Only return the MIGRATE always
            If MigratePatchOnly Then
                If UCase(Trim(TempArray(0))) = "MIGRATE" Then
                    mobjPatches_NotApplied.Add TempArray(4), TempArray(3), TempArray(2), TempArray(1), TempArray(0), "", CLng(TempArray(5)), Trim(TempArray(0)) & "|" & Trim(TempArray(1))
                    Exit Do
                End If
            Else
                On Error Resume Next
                Set objPatch = objApplied(Trim(TempArray(0)) & "|" & Trim(TempArray(1)))
                If Err.Number <> 0 Then 'Or Trim(TempArray(0)) = "MIGRATE" Then 'Always perform migration
                    On Error GoTo ErrHandler
                    mobjPatches_NotApplied.Add TempArray(4), TempArray(3), TempArray(2), TempArray(1), TempArray(0), "", CLng(TempArray(5)), Trim(TempArray(0)) & "|" & Trim(TempArray(1))
                End If
            End If
            On Error GoTo ErrHandler
        End If
        i = i + 1
        DoEvents
        If i > 200 Then
            DoEvents
        End If
    Loop While sTemp <> "The_End"
    Set Patches_Not_Applied = mobjPatches_NotApplied

GoTo CleanUp
ErrHandler:
        gobjErrManager.Err_Log "Patch_Manager", "Patches_Not_Applied", Err.Number, Err.Description
CleanUp:
End Property
Public Property Get CurrentStatement() As Variant
    CurrentStatement = msCurrentStatement
End Property
Public Property Get ErrorCount() As Long
    ErrorCount = mlSessionErrors
End Property
Public Property Get LatestError() As Variant
    LatestError = msLatestError
End Property

Private Sub Recompile_Schemas()
On Error GoTo ErrHandler
'    Dim objFileManager As File_Load_Manager
    Dim objColl As Collection
    Dim strSchema As Variant
    Dim strComment As String
    Dim strErrorMessage As String
    Dim lErrorCode As Long
    Dim strStatement As String
    Dim i As Integer
    Dim sLine As String
    Dim objIniFile As New INI_File_Manager
    
    objIniFile.File_Name = App.Path & "\patch.ini"
    objIniFile.Application_Name = "compile schemas"

    Set objColl = New Collection
    
    i = 1
    Do
        sLine = Trim(objIniFile.Get_Value(CStr(i)))
        If sLine <> "" And sLine <> "The_End" Then
            objColl.Add sLine
        End If
        i = i + 1
    Loop While sLine <> "The_End"
    
'    Set objFileManager = New File_Load_Manager
'    objFileManager.File_Open JoinPathAndFileName(GetShortFolderName(App.Path), "compile_schemas.ini"), True, True
'    If Not objFileManager.EOF Then
'
'        Do While Not objFileManager.EOF
'            sLine = Trim(objFileManager.Get_Line)
'            If sLine <> "" Then
'                objColl.Add sLine
'            End If
'        Loop
'    End If
'    objFileManager.File_Close

    If mobjOracleManager Is Nothing Then
        Set mobjOracleManager = New oracle_manager
    End If
    mobjOracleManager.Connect msHost, msUID, msPassword, strErrorMessage, lErrorCode
    
    If lErrorCode = 0 Then
        For i = 1 To 3
            For Each strSchema In objColl
                strStatement = "declare begin dbms_utility.compile_schema('" & strSchema & "', FALSE); end;"
                mobjOracleManager.Run_SQL strStatement, strComment, strErrorMessage, lErrorCode
                If lErrorCode = 0 Then
                    Write_To_Log mobjPatchLog, "NA", 0, strStatement, "Compile invalid objects", "Success", strErrorMessage, lErrorCode
                Else
                    Write_To_Log mobjPatchLog, "NA", 0, strStatement, "Compile invalid objects", "Failed", strErrorMessage, lErrorCode
                    Write_To_Log mobjErrorLog, "NA", 0, strStatement, "Compile invalid objects", "Failed", strErrorMessage, lErrorCode
                End If
            Next
        Next
    Else
        Write_To_Log mobjPatchLog, "NA", 0, "Oracle Connection for compile objects", "Failed to connect to Oracle as user '" & msUID & "'", "Failed", strErrorMessage, lErrorCode
        Write_To_Log mobjErrorLog, "NA", 0, "Oracle Connection for compile objects", "Failed to connect to Oracle as user '" & msUID & "'", "Failed", strErrorMessage, lErrorCode
    End If

GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "Patch_Manager", "Recompile_Schemas", Err.Number, Err.Description, True
CleanUp:
End Sub

Private Sub Run_Script(objParserSqlplus As Parser_Sqlplus_Def, Optional Parent_Path As String = "")
On Error GoTo ErrHandler

    Dim sStatement As String
    Dim strErrorMessage As String
    Dim lErrorCode As Long
    Dim strComment As String
    Dim objParserSqlplusDefCall As Parser_Sqlplus_Def
    Dim objOracleParser As parser_sql_script
    Dim iLineNumber As Long
    Dim strFile As String
    Dim bConnected As Boolean
    Dim strError As String
    
    bConnected = False
    strFile = objParserSqlplus.Filepath & "\" & objParserSqlplus.Filename
    
    '-- Connect to Oracle
    If mobjOracleManager Is Nothing Then
        Set mobjOracleManager = New oracle_manager
        mobjOracleManager.Current_Folder = objParserSqlplus.Filepath
    
        If LCase(objParserSqlplus.UID) <> "system" Then
            mobjOracleManager.Connect objParserSqlplus.Host, objParserSqlplus.UID, objParserSqlplus.Password, strErrorMessage, lErrorCode
        Else
            mobjOracleManager.Connect msHost, msUID, msPassword, strErrorMessage, lErrorCode
        End If
        
        Write_To_Log mobjPatchLog, strFile, 0, "Connect to Oracle", "", GetResult("", lErrorCode), strErrorMessage, lErrorCode
        If lErrorCode > 0 Then
            mlSessionErrors = mlSessionErrors + 1
            Write_To_Log mobjErrorLog, strFile, 0, "Connect to Oracle", "", GetResult("", lErrorCode), strErrorMessage, lErrorCode
        Else
            bConnected = True
        End If

    Else
        bConnected = True
    End If
    
    '-- Execute Script Commands
    If bConnected Then
        mobjOracleManager.Variables = objParserSqlplus.Variables
        
        Set objOracleParser = New parser_sql_script
        objOracleParser.parse_definitions = mobjParseDefinitions
        objOracleParser.Parse_Start objParserSqlplus.Filepath & "\" & objParserSqlplus.Filename

        Do While Not objOracleParser.EOF And Not (mbStopOnError = True And mbSessionError = True)
            lErrorCode = 0
            If objOracleParser.Next_Statement(sStatement, iLineNumber) Then
                If Mid$(sStatement, 1, 1) = "@" Then '-- Run New Script Command
                    If Parent_Path = "" Then '-- Script located in current script folder
                        msCurrentStatement = sStatement
                        RaiseEvent NewStatement
                        WaitForConfirmation
                        Set objParserSqlplusDefCall = mobjOracleManager.Create_Sqlplus_Definition(sStatement, objParserSqlplus, objParserSqlplus.Filepath)
                        Run_Script objParserSqlplusDefCall, objParserSqlplus.Filepath
                    Else '-- Script located in different folder
                        msCurrentStatement = sStatement
                        RaiseEvent NewStatement
                        WaitForConfirmation
                        Set objParserSqlplusDefCall = mobjOracleManager.Create_Sqlplus_Definition(sStatement, objParserSqlplus, Parent_Path)
                        Run_Script objParserSqlplusDefCall, Parent_Path
                    End If
                    
                    mobjOracleManager.Variables = objParserSqlplus.Variables '-- Set variables back to current script variables
                Else '-- Execute SQL Command
                    msCurrentStatement = sStatement
                    RaiseEvent NewStatement
                    WaitForConfirmation
            
                    If mbExecute_Statements Then
                        mobjOracleManager.Run_SQL sStatement, strComment, strErrorMessage, lErrorCode
                    End If
                    Write_To_Log mobjPatchLog, strFile, CInt(iLineNumber), sStatement, strComment, GetResult(strComment, lErrorCode), strErrorMessage, lErrorCode
                    
                    If lErrorCode > 0 Then
                        'SF added Ignore Oracle Errors
                        If Mid(strErrorMessage, 1, 4) = "ORA-" Then
                            strError = Mid(strErrorMessage, 1, 9)
                        Else
                            strError = ""
                        End If
                        If Not InList(mOracleErrorsToIgnore, strError) Then
                            Write_To_Log mobjErrorLog, strFile, CInt(iLineNumber), sStatement, strComment, GetResult(strComment, lErrorCode), strErrorMessage, lErrorCode
                            mlSessionErrors = mlSessionErrors + 1
                            mbSessionError = True
                            RaiseEvent StatementError
                            WaitForConfirmation
                        End If
                    End If
                End If
            End If
            
            mlStatementsProcessed = mlStatementsProcessed + 1
            RaiseEvent StatementProcessed
            WaitForConfirmation
            DoEvents
            
        Loop
        objOracleParser.Parse_End
    
        If Parent_Path = "" Then
            mobjOracleManager.Commit_Transaction
            mobjOracleManager.Disconnect
        End If
    
    End If

GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "Patch_Manager", "Run_Script", Err.Number, Err.Description
CleanUp:
    If Parent_Path = "" Then
        Set mobjOracleManager = Nothing
    End If
    Set objOracleParser = Nothing
    Set objParserSqlplusDefCall = Nothing
End Sub

Private Function InList(objColl As Collection, ErrMess As String) As Boolean
On Error GoTo ErrHandler
    
    Dim Temp As Variant
    
    On Error Resume Next
    Temp = objColl(ErrMess)
    If Err.Number = 0 Then
        InList = True
    Else
        InList = False
    End If
    On Error GoTo ErrHandler

GoTo CleanUp
ErrHandler:
        gobjErrManager.Err_Log "Imp_Exp_Manager", "InList", Err.Number, Err.Description
CleanUp:
End Function

Public Sub Run_Batch(Filename As String, Errors As Long)
On Error GoTo ErrHandler

    Dim objParserBatch As New parser_batch
    Dim objLine As parser_batch_line
    Dim sPathName As String
    Dim sFileName As String
    Dim sTempPath As String
    Dim sImportFileCommand As String
    Dim sErrorMessage As String
    Dim lErrorNumber As Long
    Dim lOraErrors As Long
    Dim lImpErrors As Long
    Dim lExpErrors As Long
    Dim strError As String
    Dim strWarning As String
    Dim objOraErrorCount As Collection
    Dim objImpErrorCount As Collection
    Dim objExpErrorCount As Collection
    Dim strExportFileCommand As String
    Dim iExpCount As Long
    Dim iImpCount As Long
    Dim strRegFileCommand As String
    
    Dim iImportErrors As Long
    Dim iExportErrors As Long
    Dim iDeleteErrors As Long
    Dim iScriptErrors As Long


    sTempPath = Environ("TEMP")
    SplitPathAndFileName Filename, sPathName, sFileName
    objParserBatch.Parse_Start Filename
    Do While Not objParserBatch.EOF And Not (mbStopOnError = True And mbSessionError = True)
        lOraErrors = 0: lExpErrors = 0: lImpErrors = 0: lErrorNumber = 0
        objParserBatch.Next_Statement objLine
        Select Case objLine.Batch_Command
            Case sqlplus
                Write_To_Log mobjPatchLog, Filename, objLine.Line_Number, objLine.Statement, "Start"
                Run_Script objLine.CommandObject
                Write_To_Log mobjPatchLog, Filename, objLine.Line_Number, objLine.Statement, "End"
            Case del
                msCurrentStatement = "Del " & objLine.CommandObject
                RaiseEvent NewStatement
                WaitForConfirmation
                If mbExecute_Statements = True Then
                    mobjDosCommandManager.Delete objLine.CommandObject, sErrorMessage, lErrorNumber
                    Write_To_Log mobjPatchLog, Filename, objLine.Line_Number, objLine.Statement, "", GetResult("", lErrorNumber), sErrorMessage, lErrorNumber

                    If lErrorNumber > 0 Then
                        Write_To_Log mobjErrorLog, Filename, objLine.Line_Number, objLine.Statement, "", GetResult("", lErrorNumber), sErrorMessage, lErrorNumber
                    End If
                End If
            Case regedit
                'iRegCount = iRegCount + 1
                mobjDosCommandManager.Prepare_Regedit_Command objLine.CommandObject, gstrLogPath & msCurrentPatch & "\", "reglogfile.log", GetShortFolderName(msPatchPath), strRegFileCommand
                msCurrentStatement = strRegFileCommand
                RaiseEvent NewStatement
                WaitForConfirmation
                If mbExecute_Statements = True Then
                    mobjDosCommandManager.ShellAndWait strRegFileCommand
                    'mobjImpExpManager.Get_Import_Results gstrLogPath & msCurrentPatch & "\" & "implogfile.log", lOraErrors, lImpErrors, objOraErrorCount, objImpErrorCount
                    Write_To_Log mobjPatchLog, Filename, objLine.Line_Number, objLine.Statement, "", "NA", "", 0
'                    If lOraErrors > 0 Or lImpErrors > 0 Then
'                        Write_To_Log_Imp_Exp mobjErrorLog, "Import", Filename, objLine.Line_Number, objLine.Statement, lOraErrors, lImpErrors, objOraErrorCount, objImpErrorCount
'                    End If
                End If
            Case import
                iImpCount = iImpCount + 1
                mobjImpExpManager.Prepare_Import_Command objLine.CommandObject, sPathName, gstrLogPath & msCurrentPatch & "\", "implogfile" & iImpCount & ".log", gstrImportFile, msConnectionString, GetShortFolderName(msPatchPath), sImportFileCommand
                msCurrentStatement = sImportFileCommand
                RaiseEvent NewStatement
                WaitForConfirmation
                If mbExecute_Statements = True Then
                    mobjDosCommandManager.ShellAndWait sImportFileCommand
                    mobjImpExpManager.Get_Import_Results gstrLogPath & msCurrentPatch & "\" & "implogfile" & iImpCount & ".log", lOraErrors, lImpErrors, objOraErrorCount, objImpErrorCount
                    Write_To_Log_Imp_Exp mobjPatchLog, "Import", Filename, objLine.Line_Number, objLine.Statement, lOraErrors, lImpErrors, objOraErrorCount, objImpErrorCount
                    If lOraErrors > 0 Or lImpErrors > 0 Then
                        Write_To_Log_Imp_Exp mobjErrorLog, "Import", Filename, objLine.Line_Number, objLine.Statement, lOraErrors, lImpErrors, objOraErrorCount, objImpErrorCount
                    End If
                End If
            Case export
                iExpCount = iExpCount + 1
                mobjImpExpManager.Prepare_Export_Command objLine.CommandObject, sPathName, gstrLogPath & msCurrentPatch & "\", "explogfile" & Trim(CStr(iExpCount)) & ".log", gstrExportFile, msConnectionString, GetShortFolderName(msPatchPath), strExportFileCommand
                msCurrentStatement = strExportFileCommand
                RaiseEvent NewStatement
                WaitForConfirmation
                If mbExecute_Statements = True Then
                    mobjDosCommandManager.ShellAndWait strExportFileCommand
                    mobjImpExpManager.Get_Export_Results gstrLogPath & msCurrentPatch & "\" & "explogfile" & Trim(CStr(iExpCount)) & ".log", lOraErrors, lExpErrors, objOraErrorCount, objExpErrorCount
                    Write_To_Log_Imp_Exp mobjPatchLog, "Export", Filename, objLine.Line_Number, objLine.Statement, lOraErrors, lExpErrors, objOraErrorCount, objExpErrorCount
                    If lOraErrors > 0 Or lExpErrors > 0 Then
                        Write_To_Log_Imp_Exp mobjErrorLog, "Export", Filename, objLine.Line_Number, objLine.Statement, lOraErrors, lExpErrors, objOraErrorCount, objExpErrorCount
                    End If
                End If
        End Select
        
        mlStatementsProcessed = mlStatementsProcessed + 1
        RaiseEvent StatementProcessed
        WaitForConfirmation
        If (lOraErrors > 0 Or lExpErrors > 0 Or lImpErrors <> 0 Or lErrorNumber <> 0) Then
            mlSessionErrors = mlSessionErrors + 1
            mbSessionError = True
            RaiseEvent StatementError
            WaitForConfirmation
        End If
    Loop
    


GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "Patch_Manager", "Run_Batch", Err.Number, Err.Description, True
CleanUp:
    Set objParserBatch = Nothing
End Sub

Public Function Login(Host As String, UID As String, Password As String, Message As String) As Boolean
On Error GoTo ErrHandler

    Login = False
    
    Set mOraSession = CreateObject("OracleInProcServer.XOraSession")
    Set mOraDatabase = mOraSession.DbOpenDatabase(Host, UID + "/" + Password, 0&)
    
    msHost = Host
    msUID = UID
    msPassword = Password
    
    If msHost = "" Then
        msConnectionString = UID & "/" & Password
    Else
        msConnectionString = UID & "/" & Password & "@" & Host
    End If
    
    Login = True

GoTo CleanUp
ErrHandler:
    'gobjErrManager.Err_Log "Excute_Batch", "Execute_Batch", Err.Number, Err.Description
    Message = Err.Description
CleanUp:
mbLogin = Login
End Function

Private Sub Create_Parse_Definitions()
On Error GoTo ErrHandler

    Dim objparser_def As parser_definition
    Dim strlinefeed As String
    
    Set mobjParseDefinitions = New parser_def_collection
    
    Set objparser_def = New parser_definition
    objparser_def.start_def = "create or replace view"
    objparser_def.start_precede.Add " "
    objparser_def.start_precede.Add gstr_strlinefeed
    objparser_def.start_follow.Add " "
    objparser_def.start_follow.Add gstr_strlinefeed
    objparser_def.end_precede.Add gstr_anycharacter
    objparser_def.end_def.Add ";"
    objparser_def.end_def.Add gstr_strlinefeed & "/"
    objparser_def.end_follow.Add gstr_anycharacter
    objparser_def.type_def = Statement
    objparser_def.embedded_sql = False
    mobjParseDefinitions.Add objparser_def
    
    '--- SCR 133350  Sherry Wang  6/28/2007
    '--- Add new parser definition for "create or replace public synonym"
    Set objparser_def = New parser_definition
    objparser_def.start_def = "create or replace public synonym"
    objparser_def.start_precede.Add " "
    objparser_def.start_precede.Add gstr_strlinefeed
    objparser_def.start_follow.Add " "
    objparser_def.start_follow.Add gstr_strlinefeed
    objparser_def.end_precede.Add gstr_anycharacter
    objparser_def.end_def.Add ";"
    objparser_def.end_follow.Add gstr_strlinefeed
    objparser_def.type_def = Statement
    objparser_def.embedded_sql = True
    mobjParseDefinitions.Add objparser_def
    
    Set objparser_def = New parser_definition
    objparser_def.start_def = "create or replace"
    objparser_def.start_precede.Add " "
    objparser_def.start_precede.Add gstr_strlinefeed
    objparser_def.start_follow.Add " "
    objparser_def.start_follow.Add gstr_strlinefeed
    objparser_def.end_precede.Add gstr_anycharacter
    objparser_def.end_def.Add gstr_strlinefeed & "/"
    objparser_def.end_follow.Add gstr_strlinefeed
    objparser_def.type_def = Statement
    objparser_def.embedded_sql = False
    mobjParseDefinitions.Add objparser_def
    
    Set objparser_def = New parser_definition
    objparser_def.start_def = "declare"
    objparser_def.start_precede.Add " "
    objparser_def.start_precede.Add gstr_strlinefeed
    objparser_def.start_follow.Add " "
    objparser_def.start_follow.Add gstr_strlinefeed
    objparser_def.end_precede.Add gstr_anycharacter
    objparser_def.end_def.Add gstr_strlinefeed & "/"
    objparser_def.end_follow.Add gstr_anycharacter
    objparser_def.type_def = Statement
    objparser_def.embedded_sql = False
    mobjParseDefinitions.Add objparser_def
    
    Set objparser_def = New parser_definition
    objparser_def.start_def = "begin"
    objparser_def.start_precede.Add " "
    objparser_def.start_precede.Add gstr_strlinefeed
    objparser_def.start_follow.Add " "
    objparser_def.start_follow.Add gstr_strlinefeed
    objparser_def.end_precede.Add gstr_anycharacter
    objparser_def.end_def.Add gstr_strlinefeed & "/"
    objparser_def.end_follow.Add gstr_anycharacter
    objparser_def.type_def = Statement
    objparser_def.embedded_sql = False
    mobjParseDefinitions.Add objparser_def
    
    Set objparser_def = New parser_definition
    objparser_def.start_def = "truncate"
    objparser_def.start_precede.Add " "
    objparser_def.start_precede.Add gstr_strlinefeed
    objparser_def.start_follow.Add " "
    objparser_def.start_follow.Add gstr_strlinefeed
    objparser_def.end_precede.Add gstr_anycharacter
    objparser_def.end_def.Add ";"
    objparser_def.end_follow.Add gstr_strlinefeed
    objparser_def.type_def = Statement
    objparser_def.embedded_sql = False
    mobjParseDefinitions.Add objparser_def

    Set objparser_def = New parser_definition
    objparser_def.start_def = "alter"
    objparser_def.start_precede.Add " "
    objparser_def.start_precede.Add gstr_strlinefeed
    objparser_def.start_follow.Add " "
    objparser_def.start_follow.Add gstr_strlinefeed
    objparser_def.end_precede.Add gstr_anycharacter
    objparser_def.end_def.Add ";"
    objparser_def.end_follow.Add gstr_anycharacter
    objparser_def.type_def = Statement
    objparser_def.embedded_sql = True
    mobjParseDefinitions.Add objparser_def

    Set objparser_def = New parser_definition
    objparser_def.start_def = "select"
    objparser_def.start_precede.Add " "
    objparser_def.start_precede.Add gstr_strlinefeed
    objparser_def.start_follow.Add " "
    objparser_def.start_follow.Add gstr_strlinefeed
    objparser_def.end_precede.Add gstr_anycharacter
    objparser_def.end_def.Add ";"
    objparser_def.end_follow.Add gstr_strlinefeed
    objparser_def.type_def = Statement
    objparser_def.embedded_sql = True
    mobjParseDefinitions.Add objparser_def
    
    Set objparser_def = New parser_definition
    objparser_def.start_def = "insert"
    objparser_def.start_precede.Add " "
    objparser_def.start_precede.Add gstr_strlinefeed
    objparser_def.start_follow.Add " "
    objparser_def.start_follow.Add gstr_strlinefeed
    objparser_def.end_precede.Add gstr_anycharacter
    objparser_def.end_def.Add ";"
    objparser_def.end_follow.Add gstr_strlinefeed
    objparser_def.type_def = Statement
    objparser_def.embedded_sql = True
    mobjParseDefinitions.Add objparser_def
    
    Set objparser_def = New parser_definition
    objparser_def.start_def = "update"
    objparser_def.start_precede.Add " "
    objparser_def.start_precede.Add gstr_strlinefeed
    objparser_def.start_follow.Add " "
    objparser_def.start_follow.Add gstr_strlinefeed
    objparser_def.end_precede.Add gstr_anycharacter
    objparser_def.end_def.Add ";"
    objparser_def.end_follow.Add gstr_strlinefeed
    objparser_def.type_def = Statement
    objparser_def.embedded_sql = True
    mobjParseDefinitions.Add objparser_def
    
    Set objparser_def = New parser_definition
    objparser_def.start_def = "delete"
    objparser_def.start_precede.Add " "
    objparser_def.start_precede.Add gstr_strlinefeed
    objparser_def.start_follow.Add " "
    objparser_def.start_follow.Add gstr_strlinefeed
    objparser_def.end_precede.Add gstr_anycharacter
    objparser_def.end_def.Add ";"
    objparser_def.end_follow.Add gstr_strlinefeed
    objparser_def.type_def = Statement
    objparser_def.embedded_sql = True
    mobjParseDefinitions.Add objparser_def
    
    Set objparser_def = New parser_definition
    objparser_def.start_def = "set"
    objparser_def.start_precede.Add " "
    objparser_def.start_precede.Add gstr_strlinefeed
    objparser_def.start_follow.Add " "
    objparser_def.start_follow.Add gstr_strlinefeed
    objparser_def.end_precede.Add gstr_anycharacter
    objparser_def.end_def.Add gstr_strlinefeed
    objparser_def.end_follow.Add gstr_notapplicable
    objparser_def.type_def = Statement
    objparser_def.embedded_sql = False
    mobjParseDefinitions.Add objparser_def

    Set objparser_def = New parser_definition
    objparser_def.start_def = "drop"
    objparser_def.start_precede.Add " "
    objparser_def.start_precede.Add gstr_strlinefeed
    objparser_def.start_follow.Add " "
    objparser_def.start_follow.Add gstr_strlinefeed
    objparser_def.end_precede.Add gstr_anycharacter
    objparser_def.end_def.Add ";"
    objparser_def.end_follow.Add gstr_anycharacter
    objparser_def.type_def = Statement
    objparser_def.embedded_sql = False
    mobjParseDefinitions.Add objparser_def

    Set objparser_def = New parser_definition
    objparser_def.start_def = "create"
    objparser_def.start_precede.Add " "
    objparser_def.start_precede.Add gstr_strlinefeed
    objparser_def.start_follow.Add " "
    objparser_def.start_follow.Add gstr_strlinefeed
    objparser_def.end_precede.Add gstr_anycharacter
    objparser_def.end_def.Add ";"
    objparser_def.end_follow.Add gstr_anycharacter
    objparser_def.type_def = Statement
    objparser_def.embedded_sql = False
    mobjParseDefinitions.Add objparser_def

    Set objparser_def = New parser_definition
    objparser_def.start_def = "grant"
    objparser_def.start_precede.Add " "
    objparser_def.start_precede.Add gstr_strlinefeed
    objparser_def.start_follow.Add " "
    objparser_def.start_follow.Add gstr_strlinefeed
    objparser_def.end_precede.Add gstr_anycharacter
    objparser_def.end_def.Add ";"
    objparser_def.end_follow.Add gstr_anycharacter
    objparser_def.type_def = Statement
    objparser_def.embedded_sql = False
    mobjParseDefinitions.Add objparser_def
    
    Set objparser_def = New parser_definition
    objparser_def.start_def = "commit"
    objparser_def.start_precede.Add " "
    objparser_def.start_precede.Add gstr_strlinefeed
    objparser_def.start_follow.Add gstr_notapplicable
    objparser_def.end_precede.Add gstr_notapplicable
    objparser_def.end_def.Add ";"
    objparser_def.end_def.Add "/"
    objparser_def.end_follow.Add gstr_notapplicable
    objparser_def.type_def = Statement
    objparser_def.embedded_sql = False
    mobjParseDefinitions.Add objparser_def
    
    Set objparser_def = New parser_definition
    objparser_def.start_def = "@"
    objparser_def.start_precede.Add " "
    objparser_def.start_precede.Add gstr_strlinefeed
    objparser_def.start_follow.Add gstr_notapplicable
    objparser_def.end_precede.Add gstr_anycharacter
    objparser_def.end_def.Add gstr_strlinefeed
    objparser_def.end_follow.Add gstr_notapplicable
    objparser_def.type_def = Statement
    objparser_def.embedded_sql = False
    mobjParseDefinitions.Add objparser_def
    
    Set objparser_def = New parser_definition
    objparser_def.start_def = "--"
    objparser_def.start_precede.Add gstr_anycharacter
    objparser_def.start_follow.Add gstr_notapplicable
    objparser_def.end_precede.Add gstr_notapplicable
    objparser_def.end_def.Add gstr_strlinefeed
    objparser_def.end_follow.Add gstr_notapplicable
    objparser_def.type_def = ignore_contents
    objparser_def.embedded_sql = False
    mobjParseDefinitions.Add objparser_def

    Set objparser_def = New parser_definition
    objparser_def.start_def = "rem"
    objparser_def.start_precede.Add gstr_anycharacter
    objparser_def.start_follow.Add gstr_notapplicable
    objparser_def.end_precede.Add gstr_notapplicable
    objparser_def.end_def.Add gstr_strlinefeed
    objparser_def.end_follow.Add gstr_notapplicable
    objparser_def.type_def = ignore_contents
    objparser_def.embedded_sql = False
    mobjParseDefinitions.Add objparser_def
    
    Set objparser_def = New parser_definition
    objparser_def.start_def = "/*"
    objparser_def.start_precede.Add gstr_anycharacter
    objparser_def.start_follow.Add gstr_notapplicable
    objparser_def.end_precede.Add gstr_notapplicable
    objparser_def.end_def.Add "*/"
    objparser_def.end_follow.Add gstr_notapplicable
    objparser_def.type_def = ignore_contents
    objparser_def.embedded_sql = False
    mobjParseDefinitions.Add objparser_def
    
    Set objparser_def = New parser_definition
    objparser_def.start_def = "spool"
    objparser_def.start_precede.Add " "
    objparser_def.start_precede.Add gstr_strlinefeed
    objparser_def.start_follow.Add " "
    objparser_def.end_precede.Add gstr_anycharacter
    objparser_def.end_def.Add gstr_strlinefeed
    objparser_def.end_follow.Add gstr_notapplicable
    objparser_def.type_def = Statement
    objparser_def.embedded_sql = False
    mobjParseDefinitions.Add objparser_def

GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "Patch_Manager", "Create_Parse_Definitions", Err.Number, Err.Description
CleanUp:
End Sub
Public Function Total_Time(StartPatchKey As String, EndPatchKey As String) As Long
On Error GoTo ErrHandler
    
    Dim objPatch As Variant
    Dim bStarted As Boolean
    
    bStarted = False
    Total_Time = 0
    
    For Each objPatch In mobjPatches_NotApplied
        If objPatch.key = StartPatchKey Or bStarted Then
            bStarted = True
            Total_Time = Total_Time + objPatch.Run_Time
            If objPatch.key = EndPatchKey Then
                Exit For
            End If
        End If
    Next

GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "Patch_Manager", "Total_Time", Err.Number, Err.Description
CleanUp:
End Function

Private Sub WaitForConfirmation()
On Error GoTo ErrHandler

    Do While mbWaitForEventConfirmation And mbEventConfirmed
        DoEvents
    Loop
    mbEventConfirmed = False
    
GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "Patch_Manager", "WaitForConfirmation", Err.Number, Err.Description
CleanUp:
End Sub

Private Sub Write_To_Log(Log As Log_Manager, Filename As String, Line_Number As Long, _
                         Statement As String, Optional Comment As String, Optional Result As String = "", _
                         Optional Error_Message As String = "", Optional Error_Code As Long = 0)
On Error GoTo ErrHandler
    
    Log.WriteField "File: " & Filename
    
    If Line_Number <> 0 Then
        Log.WriteField "Line: " & CStr(Line_Number)
    End If
    
    If Statement <> "" Then
        Log.WriteField "Command: " & Statement
    End If
    Log.WriteField "Time: " & Now
    
    If Comment <> "" Then
        Log.WriteField "Comment: " & Comment
    End If
    
    If Result <> "" Then
        Log.WriteField "Result: " & Result
    End If
    
    If Error_Message <> "" Then
        Log.WriteField "Error: " & Error_Message
    End If

    If Error_Code <> 0 Then
        Log.WriteField "Error Code: " & CStr(Error_Code)
    End If
    
    Log.WriteRecord "---------------------------------------------------------------------------------------"
        
GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "Patch_Manager", "Write_To_Patch_Log", Err.Number, Err.Description
CleanUp:
End Sub

Private Sub Write_To_Log_Imp_Exp(objLog As Log_Manager, Log_File_Type As String, Filename As String, Line_Number As Long, _
                                Statement As String, _
                                lOraErrors As Long, lImpErrors As Long, _
                                objOraErrorCount As Collection, objImpErrorCount As Collection)
On Error GoTo ErrHandler

    Dim objErrorMsg As Error_Message_Count
    
    objLog.WriteField "File: " & Filename
    
    If Line_Number <> 0 Then
        objLog.WriteField "Line: " & CStr(Line_Number)
    End If
    
    If Statement <> "" Then
        objLog.WriteField "Command: " & Statement
    End If
    objLog.WriteField "Time: " & Now
    
    If lOraErrors = 0 And lImpErrors = 0 Then
        objLog.WriteField "Result: Success"
    Else
        objLog.WriteField "Result: Failed"
        If lImpErrors <> 0 Then
            For Each objErrorMsg In objImpErrorCount
                objLog.WriteField Log_File_Type & " Error: " & objErrorMsg.Message & " - " & CStr(objErrorMsg.Count) & " Time(s)"
            Next
        End If
        
        If lOraErrors <> 0 Then
            For Each objErrorMsg In objOraErrorCount
                objLog.WriteField "Oracle Error: " & objErrorMsg.Message & " - " & CStr(objErrorMsg.Count) & " Time(s)"
            Next
        End If
        
    End If

    objLog.WriteRecord "---------------------------------------------------------------------------------------"
        
GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "Patch_Manager", "Write_To_Patch_Log", Err.Number, Err.Description
CleanUp:
End Sub

Private Sub Class_Initialize()
On Error GoTo ErrHandler

    Dim objColl As Collection

    Set mobjImpExpManager = New imp_exp_manager
    Set mobjDosCommandManager = New dos_command_manager
    'Set mobjOracleManager = New oracle_manager
    
    Set mobjPatches_NotApplied = New patch_collection
    Set mobjPatches_Applied = New patch_collection
    Set mobjPatches_Available = New patch_collection
    
    
    msRunPath = Environ("TEMP")
    Create_Parse_Definitions
    mbWaitForEventConfirmation = True
    mbEventConfirmed = False
    mbExecute_Statements = True
    
    Get_Import_Errors_To_Ignore objColl
    mobjImpExpManager.Import_Ignore_Errors = objColl
    
    Get_Export_Errors_To_Ignore objColl
    mobjImpExpManager.Export_Ignore_Errors = objColl
    
    Get_Oracle_Errors_To_Ignore mOracleErrorsToIgnore
    
GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "Patch_Manager", "Class_Initialize", Err.Number, Err.Description
CleanUp:
End Sub

Public Sub EventConfirmed()
    mbWaitForEventConfirmation = True
End Sub

Public Sub Run_Patches(StartPatchKey As String, EndPatchKey As String, _
                       Export_File As String, _
                       Import_File As String, _
                       Stop_On_Error As Boolean, _
                       Optional Compile_Schemas As Boolean = False, _
                       Optional MigratePatchOnly As Boolean = False)
                       
On Error GoTo ErrHandler

    Dim sFileName As String
    Dim sCurrentPatch As String
    Dim i As Long
    Dim iTotalErrors As Long
    Dim objPatchesNotApplied As patch_collection
    Dim objPatch As patch
    Dim iErrors As Long
    Dim lTimer
        
        
    mstrExportFile = Export_File
    mstrImportFile = Import_File
    mbStopOnError = Stop_On_Error

    mbSessionError = False
    miPatchesProcessed = 0
    mlStatementsProcessed = 0

    Set objPatchesNotApplied = Me.Patches_Not_Applied(MigratePatchOnly)

    If objPatchesNotApplied(1).key = StartPatchKey Then

        For i = 1 To objPatchesNotApplied.Count

            Set objPatch = objPatchesNotApplied.Item(i)

            sFileName = App.Path & "\" & objPatch.folder & "\" & objPatch.Filename
            msPatchPath = App.Path & "\" & objPatch.folder
            msCurrentPatch = objPatch.version_id & "-" & objPatch.patch_id
            
            lTimer = Timer
        
            Set mobjPatchLog = New Log_Manager
            mobjPatchLog.OpenLog gstrLogPath & msCurrentPatch & "\" & "patch.log", True
            
            RaiseEvent PatchStarted(objPatch)
            Me.Run_Batch sFileName, iErrors
            RaiseEvent PatchCompleted(objPatch, iErrors)
            miPatchesProcessed = miPatchesProcessed + 1
            
            
            If Compile_Schemas = True And i = objPatchesNotApplied.Count Then
                Recompile_Schemas
            End If

            lTimer = Timer - lTimer
            Write_To_Log mobjPatchLog, sFileName, 0, "", "Running Time = " & CStr(Round(lTimer)) & " Seconds"
            mobjPatchLog.CloseLog

            If mbSessionError = True And mbStopOnError = True Then
                Exit For
            End If
             

            If objPatch.key = EndPatchKey Then
                Exit For
            End If

            If mbStopOnError And mlSessionErrors > 0 Then
                Exit For
            End If
        
        Next
    
    End If


GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "Patch_Manager", "Run_Patches", Err.Number, Err.Description
CleanUp:
    Set mobjPatchLog = Nothing
    Set objPatchesNotApplied = Nothing
    Set objPatch = Nothing
End Sub

Public Property Get Execute_Statement() As Boolean
    Execute_Statement = mbExecute_Statements
End Property

Public Property Let Execute_Statement(ByVal vNewValue As Boolean)
    mbExecute_Statements = vNewValue
End Property

Public Property Get Wait_For_Confirmation() As Boolean
    Wait_For_Confirmation = mbWaitForEventConfirmation
End Property

Public Property Let Wait_For_Confirmation(ByVal vNewValue As Boolean)
    mbWaitForEventConfirmation = vNewValue
End Property







Public Property Get Patches_Processed() As Long
    Patches_Processed = miPatchesProcessed
End Property



Public Property Get Statements_Processed() As Variant
    Statements_Processed = mlStatementsProcessed
End Property



Public Property Get Error_Log() As Variant
    Set Error_Log = mobjErrorLog
End Property

Public Property Let Error_Log(ByVal vNewValue As Variant)
    Set mobjErrorLog = vNewValue
End Property

Private Sub Class_Terminate()
    Set mobjImpExpManager = Nothing
    Set mobjDosCommandManager = Nothing
    Set mobjOracleManager = Nothing
    
    Set mobjPatches_NotApplied = Nothing
    Set mobjPatches_Applied = Nothing
    Set mobjPatches_Available = Nothing
End Sub


