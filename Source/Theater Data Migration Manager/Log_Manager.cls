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
Attribute VB_Name = "Log_Manager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mobjFSO As Object 'New FileSystemObject
Private mobjTS As Object 'TextStream

Private Function FileExists(Filename As String) As Boolean
On Error GoTo ErrHandler

   If Dir(Filename) <> "" Then
      FileExists = True
   Else
      FileExists = False
   End If


GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "Log_Manager", "FileExists", Err.Number, Err.Description
CleanUp:
End Function

Private Function FolderExists(Foldername As String) As Boolean
On Error GoTo ErrHandler

    FolderExists = False
    If Dir(Foldername, vbDirectory) <> "" Then
        FolderExists = True
    End If

GoTo CleanUp
ErrHandler:
    If Err.Number <> 52 Then
        gobjErrManager.Err_Log "Log_Manager", "FolderExists", Err.Number, Err.Description
    End If
CleanUp:
End Function


Public Sub OpenLog(Filename As String, Optional OverWrite As Boolean = False)
On Error GoTo ErrHandler

'    If OverWrite = True Then
'        On Error Resume Next
'        Kill FileName
'        On Error GoTo ErrHandler
'    End If

    If OverWrite = True Then
        Remove_Log Filename
    End If

    CreateFolderStructure StrReverse(Mid(StrReverse(Filename), (InStr(StrReverse(Filename), "\"))))
    
    If FileExists(Filename) And Not OverWrite Then
        Set mobjTS = mobjFSO.OpenTextFile(Filename, 8) 'mobjFSO.OpenTextFile(Filename, ForAppending)
    Else
        Set mobjTS = mobjFSO.CreateTextFile(Filename, True)
    End If
    
GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "Log_Manager", "OpenLog", Err.Number, Err.Description
CleanUp:
End Sub

Public Sub CloseLog()
On Error GoTo ErrHandler

    mobjTS.Close
    
GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "Log_Manager", "CloseLog", Err.Number, Err.Description
CleanUp:
End Sub

Private Sub Remove_Log(Filename As String)
On Error GoTo ErrHandler

    If FileExists(Filename) Then
        On Error Resume Next
        Kill Filename
    End If

GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "Log_Manager", "Remove_Log", Err.Number, Err.Description
CleanUp:
End Sub

Public Sub WriteField(FieldValue As String)
On Error GoTo ErrHandler

    mobjTS.WriteLine (FieldValue)
    
GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "Log_Manager", "WriteField", Err.Number, Err.Description
CleanUp:
End Sub

Public Sub WriteRecord(Line As String)
On Error GoTo ErrHandler

    mobjTS.WriteLine Line & Chr(0)

GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "Log_Manager", "WriteRecord", Err.Number, Err.Description
CleanUp:
End Sub

Public Sub WriteText(Text As String)
On Error GoTo ErrHandler

    mobjTS.Write (Text)

GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "Log_Manager", "WriteText", Err.Number, Err.Description
CleanUp:
End Sub

Public Sub WriteLine(Line As String)
On Error GoTo ErrHandler

    mobjTS.WriteLine (Line)
    
GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "Log_Manager", "WriteLine", Err.Number, Err.Description
CleanUp:
End Sub

Public Sub CreateFolderStructure(Filepath As String)
On Error GoTo ErrHandler

    Dim strTempFilePath As String
    Dim TempArray
    Dim i As Long

    If Not FolderExists(Filepath) Then
        TempArray = Split(Filepath, "\")
        strTempFilePath = ""
        For i = LBound(TempArray) To UBound(TempArray)
            If TempArray(i) <> "" Then
                strTempFilePath = strTempFilePath & TempArray(i) & "\"
                If Not FolderExists(strTempFilePath) Then
                    MkDir strTempFilePath
                End If
            End If
        Next
    End If
    
GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "Log_Manager", "CreateFolderStructure", Err.Number, Err.Description
CleanUp:
End Sub

Private Sub Class_Initialize()
    Set mobjFSO = CreateObject("Scripting.FileSystemObject")
End Sub


