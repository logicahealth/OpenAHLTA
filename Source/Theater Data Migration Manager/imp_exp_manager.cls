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
Attribute VB_Name = "imp_exp_manager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mobjImpErrorsIgnore As Collection
Private mobjExpErrorsIgnore As Collection
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
Private Sub Load_Ignore_Errors()
On Error GoTo ErrHandler




GoTo CleanUp
ErrHandler:
        gobjErrManager.Err_Log "Imp_Exp_Manager", "Load_Ignore_Errors", Err.Number, Err.Description
CleanUp:
End Sub

Public Sub Prepare_Import_Command(Command_Line As String, Import_File_Path As String, _
                               Log_File_Path As String, Log_File_Name As String, _
                               Import_File As String, Connection_String, _
                               Batch_Path As String, _
                               New_Command_Line As String)
On Error GoTo ErrHandler

    Dim objFSO As Object 'New Scripting.FileSystemObject
    Dim objFolder As Object 'Scripting.folder
    Dim TempString As Variant
    Dim TempArray
    Dim objColl As Collection
    Dim i As Long
    Dim iPosition As Long
    Dim sLogFile As String
    Dim sNewValue As String
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    'Find "file=" and include path along with filename
    '---------------------------------------------------------------------
    Set objColl = Split_Statement(Command_Line)
    i = 0
    iPosition = 0
    For Each TempString In objColl
        i = i + 1
        If Mid(LCase(TempString), 1, 5) = "file=" Then
            If Mid(LCase(TempString), 6) = "<importfile>" Then
                objColl.Remove (i)
                objColl.Add "file=" & GetShortFileName(Import_File)
            ElseIf Not (InStr(6, TempString, "\") > 0) Then
                objColl.Remove (i)
                objColl.Add "file=" & Batch_Path & "\" & Mid(TempString, 6)
            End If
            'iPosition = i
            Exit For
        End If
    Next
    
'    If iPosition > 0 Then
'        sNewValue = "file=" & JoinPathAndFileName(GetShortFolderName(Import_File_Path), Mid(objColl.Item(iPosition), 6))
'        objColl.Remove (iPosition)
'        objColl.Add sNewValue
'    End If
    
    'Find "log="
    '---------------------------------------------------------------------
    'Set objColl = Split_Statement(Command_Line)
    i = 0
    iPosition = 0
    For Each TempString In objColl
        i = i + 1
        If Mid(LCase(TempString), 1, 4) = "log=" Then
            iPosition = i
            Exit For
        End If
    Next
      
    If iPosition > 0 Then
        'if log= does exist replace log with %TEMP%\implogfile.log
        sLogFile = JoinPathAndFileName(GetShortFolderName(Log_File_Path), Log_File_Name)
        objColl.Remove (iPosition)
        objColl.Add "log=" & sLogFile
    Else
        'if log= doesn't exist then add it and include name of log file
        sLogFile = JoinPathAndFileName(GetShortFolderName(Log_File_Path), Log_File_Name)
        objColl.Add "log=" & sLogFile
    End If
    
    '-- Connection String
    '---------------------------------------------------------------------
    iPosition = 0
    i = 0
    For Each TempString In objColl
        i = i + 1
        If Mid(LCase(TempString), 1, 7) = "userid=" Then
            iPosition = i
            Exit For
        End If
    Next
    
    If iPosition = 0 Then
        'if log= doesn't exist then add it and include name of log file
        objColl.Add "userid=" & Connection_String
    Else
        objColl.Remove (iPosition)
        objColl.Add "userid=" & Connection_String
    End If
    

    '-- Parfile
    '---------------------------------------------------------------------
    iPosition = 0
    i = 0
    For Each TempString In objColl
        i = i + 1
        If Mid(LCase(TempString), 1, 8) = "parfile=" Then
            iPosition = i
            Exit For
        End If
    Next

    If iPosition <> 0 Then
        objColl.Remove (iPosition)
        objColl.Add "parfile=" & Batch_Path & "\" & Mid(TempString, 9)
    End If


    'Prepare statement
    '---------------------------------------------------------------------
    New_Command_Line = ""
    For Each TempString In objColl
        New_Command_Line = New_Command_Line & " " & TempString
    Next
    New_Command_Line = Trim(New_Command_Line)

GoTo CleanUp
ErrHandler:
        gobjErrManager.Err_Log "Imp_Exp_Manager", "Prep_Import_Command", Err.Number, Err.Description
CleanUp:
End Sub
Public Sub Prepare_Export_Command(Command_Line As String, Export_File_Path As String, _
                               Log_File_Path As String, Log_File_Name As String, _
                               Export_File As String, Connection_String, _
                               Batch_Path As String, _
                               New_Command_Line As String)
On Error GoTo ErrHandler

'exp userid=system/password@lcd tables=gems.tasking grants=n indexes=n triggers=n file=tasking.dmp

    Dim objFSO As Object 'New Scripting.FileSystemObject
    Dim objFolder As Object 'Scripting.folder
    Dim TempString As Variant
    Dim TempArray
    Dim objColl As Collection
    Dim i As Long
    Dim iPosition As Long
    Dim sLogFile As String
    Dim sNewValue As String

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    'Find "file=" and include path along with filename
    '---------------------------------------------------------------------
    Set objColl = Split_Statement(Command_Line)
    i = 0
    iPosition = 0
    For Each TempString In objColl
        i = i + 1
        If Mid(LCase(TempString), 1, 5) = "file=" Then
            If Mid(LCase(TempString), 6) = "<exportfile>" Then
                objColl.Remove (i)
                objColl.Add "file=" & GetShortFileName(Export_File)
            End If
            Exit For
        End If
    Next
    
'    If iPosition > 0 Then
'        sNewValue = "file=" & JoinPathAndFileName(GetShortFolderName(Import_File_Path), Mid(objColl.Item(iPosition), 6))
'        objColl.Remove (iPosition)
'        objColl.Add sNewValue
'    End If
    
    'Find "log="
    '---------------------------------------------------------------------
    'Set objColl = Split_Statement(Command_Line)
    i = 0
    iPosition = 0
    For Each TempString In objColl
        i = i + 1
        If Mid(LCase(TempString), 1, 4) = "log=" Then
            iPosition = i
            Exit For
        End If
    Next
      
    If iPosition > 0 Then
        'if log= does exist replace log with %TEMP%\implogfile.log
        sLogFile = JoinPathAndFileName(GetShortFolderName(Log_File_Path), Log_File_Name)
        objColl.Remove (iPosition)
        objColl.Add "log=" & sLogFile
    Else
        'if log= doesn't exist then add it and include name of log file
        sLogFile = JoinPathAndFileName(GetShortFolderName(Log_File_Path), Log_File_Name)
        objColl.Add "log=" & sLogFile
    End If
    
    '-- Parfile
    '---------------------------------------------------------------------
    iPosition = 0
    i = 0
    For Each TempString In objColl
        i = i + 1
        If Mid(LCase(TempString), 1, 8) = "parfile=" Then
            iPosition = i
            Exit For
        End If
    Next

    If iPosition <> 0 Then
        objColl.Remove (iPosition)
        objColl.Add "parfile=" & Batch_Path & "\" & Mid(TempString, 9)
    End If

    
    '-- Connection String
    '---------------------------------------------------------------------
    iPosition = 0
    i = 0
    For Each TempString In objColl
        i = i + 1
        If Mid(LCase(TempString), 1, 7) = "userid=" Then
            iPosition = i
            Exit For
        End If
    Next
    
    If iPosition = 0 Then
        'if log= doesn't exist then add it and include name of log file
        objColl.Add "userid=" & Connection_String
    Else
        objColl.Remove (iPosition)
        objColl.Add "userid=" & Connection_String
    End If
    
    
    'Prepare statement
    '---------------------------------------------------------------------
    New_Command_Line = ""
    For Each TempString In objColl
        New_Command_Line = New_Command_Line & " " & TempString
    Next
    New_Command_Line = Trim(New_Command_Line)

GoTo CleanUp
ErrHandler:
        gobjErrManager.Err_Log "Imp_Exp_Manager", "Prepare_Export_Command", Err.Number, Err.Description
        Exit Sub
        Resume
CleanUp:
End Sub




Public Sub Get_Import_Results(LogFile As String, _
                              ORA_Count As Long, _
                              IMP_Count As Long, _
                              ORA_Error_Mess As Collection, _
                              IMP_Error_Mess As Collection)
On Error GoTo ErrHandler

    Dim objFileManager As New File_Load_Manager
    Dim strFileText As String
    Dim lFilePosition As Long
    Dim strLine As String
    Dim strError As String
    Dim iWarnTemp As Long
    Dim objErrorMessCount As Error_Message_Count
    
    Set ORA_Error_Mess = Nothing
    Set ORA_Error_Mess = New Collection
    Set IMP_Error_Mess = Nothing
    Set IMP_Error_Mess = New Collection
    
    ORA_Count = 0
    IMP_Count = 0
    
    objFileManager.File_Open LogFile, True
    
    lFilePosition = 0
    
    If Not objFileManager.Instr_Bite("Import terminated successfully without warnings.") > 0 Then
        Do While Not objFileManager.EOF
            strLine = objFileManager.Get_Line
            If Mid(strLine, 1, 4) = "IMP-" Then
                strError = Mid(strLine, 1, 9)
                If Not InList(mobjImpErrorsIgnore, strError) Then
                    IMP_Count = IMP_Count + 1
                    On Error Resume Next
                    Set objErrorMessCount = IMP_Error_Mess.Item(strError)
                    If Err.Number = 0 Then
                        On Error GoTo ErrHandler
                        IMP_Error_Mess.Item(strError).Count = IMP_Error_Mess.Item(strError).Count + 1
                    Else
                        Set objErrorMessCount = New Error_Message_Count
                        On Error GoTo ErrHandler
                        objErrorMessCount.Count = 1
                        objErrorMessCount.Message = strError
                        IMP_Error_Mess.Add objErrorMessCount, strError
                    End If
                End If
                
            ElseIf Mid(strLine, 1, 4) = "ORA-" Then
                strError = Mid(strLine, 1, 9)
                If Not InList(mobjImpErrorsIgnore, strError) Then
                    ORA_Count = ORA_Count + 1
                    On Error Resume Next
                    Set objErrorMessCount = ORA_Error_Mess.Item(strError)
                    'ORA_Error_Mess.Add objErrorMessCount, strError
                    If Err.Number = 0 Then
                        On Error GoTo ErrHandler
                        ORA_Error_Mess.Item(strError).Count = ORA_Error_Mess.Item(strError).Count + 1
                    Else
                        On Error GoTo ErrHandler
                        Set objErrorMessCount = New Error_Message_Count
                        objErrorMessCount.Count = 1
                        objErrorMessCount.Message = strError
                        ORA_Error_Mess.Add objErrorMessCount, strError
                    End If
                End If
            End If
            DoEvents
        Loop
        If ORA_Count = 0 And IMP_Count = 0 And Not objFileManager.Instr_Bite("Import terminated") > 0 Then
            IMP_Count = 1
            Set objErrorMessCount = New Error_Message_Count
            On Error GoTo ErrHandler
            objErrorMessCount.Count = 1
            objErrorMessCount.Message = "Import process was stopped prematurely."
            IMP_Error_Mess.Add objErrorMessCount, "Import process was stopped prematurely."
        End If
    End If

    objFileManager.File_Close
    Set objFileManager = Nothing
    
GoTo CleanUp
ErrHandler:
        gobjErrManager.Err_Log "Imp_Exp_Manager", "Get_Import_Results", Err.Number, Err.Description
        Exit Sub
        Resume
        
CleanUp:
End Sub
Public Sub Get_Export_Results(LogFile As String, _
                              ORA_Count As Long, _
                              EXP_Count As Long, _
                              ORA_Error_Mess As Collection, _
                              EXP_Error_Mess As Collection)
On Error GoTo ErrHandler

    Dim objFileManager As New File_Load_Manager
    Dim strFileText As String
    Dim lFilePosition As Long
    Dim strLine As String
    Dim strError As String
    Dim iWarnTemp As Long
    Dim objErrorMessCount As Error_Message_Count
    
    Set ORA_Error_Mess = Nothing
    Set ORA_Error_Mess = New Collection
    Set EXP_Error_Mess = Nothing
    Set EXP_Error_Mess = New Collection
    
    ORA_Count = 0
    EXP_Count = 0
    
    objFileManager.File_Open LogFile, True
    
    lFilePosition = 0
    
    If Not objFileManager.Instr_Bite("Export terminated successfully without warnings.") > 0 Then
        Do While Not objFileManager.EOF
            strLine = objFileManager.Get_Line

            If Mid(strLine, 1, 4) = "EXP-" Then
                strError = Mid(strLine, 1, 9)
                If Not InList(mobjExpErrorsIgnore, strError) Then
                    EXP_Count = EXP_Count + 1
                    On Error Resume Next
                    Set objErrorMessCount = EXP_Error_Mess.Item(strError)
                    If Err.Number = 0 Then
                        On Error GoTo ErrHandler
                        EXP_Error_Mess.Item(strError).Count = EXP_Error_Mess.Item(strError).Count + 1
                    Else
                        Set objErrorMessCount = New Error_Message_Count
                        On Error GoTo ErrHandler
                        objErrorMessCount.Count = 1
                        objErrorMessCount.Message = strError
                        EXP_Error_Mess.Add objErrorMessCount, strError
                    End If
                End If
            ElseIf Mid(strLine, 1, 4) = "ORA-" Then
                strError = Mid(strLine, 1, 9)
                If Not InList(mobjExpErrorsIgnore, strError) Then
                    ORA_Count = ORA_Count + 1
                    On Error Resume Next
                    Set objErrorMessCount = ORA_Error_Mess.Item(strError)
                    'ORA_Error_Mess.Add objErrorMessCount, strError
                    If Err.Number = 0 Then
                        On Error GoTo ErrHandler
                        ORA_Error_Mess.Item(strError).Count = ORA_Error_Mess.Item(strError).Count + 1
                    Else
                        On Error GoTo ErrHandler
                        Set objErrorMessCount = New Error_Message_Count
                        objErrorMessCount.Count = 1
                        objErrorMessCount.Message = strError
                        ORA_Error_Mess.Add objErrorMessCount, strError
                    End If
                End If
                
            End If
            DoEvents
        Loop
        If ORA_Count = 0 And EXP_Count = 0 And Not objFileManager.Instr_Bite("Export terminated") > 0 Then
            EXP_Count = 1
            Set objErrorMessCount = New Error_Message_Count
            On Error GoTo ErrHandler
            objErrorMessCount.Count = 1
            objErrorMessCount.Message = "Export process was stopped prematurely."
            EXP_Error_Mess.Add objErrorMessCount, "Export process was stopped prematurely."
        End If
    End If
         
    objFileManager.File_Close
    Set objFileManager = Nothing
    
GoTo CleanUp
ErrHandler:
        gobjErrManager.Err_Log "EXP_Exp_Manager", "Get_Import_Results", Err.Number, Err.Description
CleanUp:
End Sub



Public Property Get Import_Ignore_Errors() As Collection
    Set Import_Ignore_Errors = mobjImpErrorsIgnore
End Property

Public Property Let Import_Ignore_Errors(ByVal vNewValue As Collection)
On Error GoTo ErrHandler

    Dim strMess As Variant

    Set mobjImpErrorsIgnore = Nothing
    Set mobjImpErrorsIgnore = New Collection
    For Each strMess In vNewValue
        On Error Resume Next
        mobjImpErrorsIgnore.Add strMess, strMess
        On Error GoTo ErrHandler
    Next

GoTo CleanUp
ErrHandler:
        gobjErrManager.Err_Log "Patch_Manager", "Import_Ignore_Errors", Err.Number, Err.Description
CleanUp:
End Property

Public Property Get Export_Ignore_Errors() As Collection
    Set Export_Ignore_Errors = mobjExpErrorsIgnore
End Property

Public Property Let Export_Ignore_Errors(ByVal vNewValue As Collection)
On Error GoTo ErrHandler

    Dim strMess As Variant

    Set mobjExpErrorsIgnore = Nothing
    Set mobjExpErrorsIgnore = New Collection
    For Each strMess In vNewValue
        On Error Resume Next
        mobjExpErrorsIgnore.Add strMess, strMess
        On Error GoTo ErrHandler
    Next

GoTo CleanUp
ErrHandler:
        gobjErrManager.Err_Log "Patch_Manager", "Export_Ignore_Errors", Err.Number, Err.Description
CleanUp:
End Property

Private Sub Class_Initialize()
    Set mobjImpErrorsIgnore = New Collection
    Set mobjExpErrorsIgnore = New Collection
End Sub

Private Sub Class_Terminate()
    Set mobjImpErrorsIgnore = Nothing
    Set mobjExpErrorsIgnore = Nothing
End Sub
