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
Attribute VB_Name = "dos_command_manager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Delete(Filename As String, Error_Message As String, Error_Number As Long)
On Error GoTo ErrHandler

    Error_Message = ""
    Error_Number = 0

    On Error Resume Next
    Kill Filename
    If Err.Number <> 0 Then
        Error_Message = Err.Description
        Error_Number = Err.Number
    End If

GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "dos_command_manager", "Delete", Err.Number, Err.Description
CleanUp:
End Sub

Public Sub Prepare_Regedit_Command(Command_Line As String, _
                                Log_File_Path As String, Log_File_Name As String, _
                                Batch_Path As String, _
                                New_Command_Line As String)
On Error GoTo ErrHandler

'regedit /s db_install.reg

    Dim TempString As Variant
    Dim TempArray
    Dim objColl As Collection
    Dim i As Long
    Dim iPosition As Long
    Dim sLogFile As String
    Dim sNewValue As String


    Set objColl = Split_Statement(Command_Line)
    
    'Find ".reg" and include path along with filename
    '---------------------------------------------------------------------

    i = 0
    For Each TempString In objColl
        i = i + 1
        If InStr(1, LCase(TempString), ".reg") > 0 Then
            objColl.Remove (i)
            objColl.Add JoinPathAndFileName(GetShortFolderName(Batch_Path), CStr(TempString))
            Exit For
        End If
    Next
    
    
    'Prepare statement
    '---------------------------------------------------------------------
    New_Command_Line = ""
    For Each TempString In objColl
        New_Command_Line = New_Command_Line & " " & TempString
    Next
    New_Command_Line = Trim(New_Command_Line)
    'New_Command_Line = New_Command_Line & " > " & JoinPathAndFileName(GetShortFolderName(Log_File_Path), Log_File_Name)

GoTo CleanUp
ErrHandler:
        gobjErrManager.Err_Log "Imp_Exp_Manager", "Prepare_Export_Command", Err.Number, Err.Description
CleanUp:
End Sub

Public Sub ShellAndWait(CommandLine As String, Optional Default_Folder As String = "")
On Error GoTo ErrHandler

    Dim ShellId As Long
    Dim ShellHandle As Long
    ShellId = Shell(CommandLine, vbNormalFocus)
    ShellHandle = OpenProcess(SYNCHRONIZE, 0, ShellId)
    If ShellHandle <> 0 Then
        WaitForSingleObject ShellHandle, INFINITE
        CloseHandle ShellHandle
    End If
    
GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "dos_command_manager", "ShellAndWait", Err.Number, Err.Description
CleanUp:
End Sub

Public Sub ConsolePrint(MsgOut As String)
On Error GoTo ErrHandler

    WriteConsole _
    hConsoleOut, _
    MsgOut, _
    Len(MsgOut), _
    vbNull, _
    vbNull
    
GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "dos_command_manager", "ConsolePrint", Err.Number, Err.Description
CleanUp:
End Sub

Public Function ConsoleRead() As String
On Error GoTo ErrHandler

    Dim MsgIn As String * 256
    Call ReadConsole( _
    hConsoleIn, _
    MsgIn, _
    Len(MsgIn), _
    vbNull, _
    vbNull)
    'Trim the NULL charactors and CRLF.
    ConsoleRead = _
        Left$(MsgIn, InStr(MsgIn, Chr$(0)) - 3)
        
GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "dos_command_manager", "ConsoleRead", Err.Number, Err.Description
CleanUp:
End Function

Public Function ReadCName() As String
On Error GoTo ErrHandler

    Dim sz As Long
    Dim s As String
    sz = MAX_COMPUTERNAME_LENGTH + 1
    s = String$(sz, 0)
    Dim dl As Long
    dl = GetComputerName(s, sz)
    If dl <> 0 Then
        ReadCName = Mid(s, 1, sz)
    Else
        ReadCName = "Invalid Call"
    End If
    
GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "dos_command_manager", "ReadCName", Err.Number, Err.Description
CleanUp:
End Function

