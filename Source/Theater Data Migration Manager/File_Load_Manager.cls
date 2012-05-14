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
Attribute VB_Name = "File_Load_Manager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private msFileText As String
Private mbEOF As Boolean
Private mlStartPosition As Long
Private mlLineOffset As Long

Private mobjFSO As Object 'FileSystemObject
Private mobjTS As Object 'TextStream


Public Sub File_Close()
    mobjTS.Close
    Set mobjTS = Nothing
    Set mobjFSO = Nothing
End Sub

Public Sub File_Open(file_name As String, _
                     Optional Load_All As Boolean = False, _
                     Optional Ignore_Error As Boolean = False)
On Error GoTo ErrHandler

    If Ignore_Error = True Then
        On Error Resume Next
        Set mobjFSO = CreateObject("Scripting.FileSystemObject")
        Set mobjTS = mobjFSO.OpenTextFile(file_name, 1, False)
        If Err.Number > 0 Then
            mbEOF = True
        Else
            mbEOF = False
        End If
    Else
        Set mobjFSO = CreateObject("Scripting.FileSystemObject") 'New FileSystemObject
        Set mobjTS = mobjFSO.OpenTextFile(file_name, 1, False)
        mbEOF = False
    End If
    
    msFileText = ""

    mlStartPosition = 0
    If Load_All And Not mbEOF Then
        Get_File_Bite mlStartPosition, Load_All
        mlStartPosition = 0
    End If
    
GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "File_Load_Manager", "File_Open", Err.Number, Err.Description
CleanUp:
End Sub

Public Function Get_File_Bite(piFile_Position As Long, Load_All As Boolean) As String
On Error GoTo ErrHandler

    Dim strTemp As String
       
    If piFile_Position <> 0 Then
        mlLineOffset = mlLineOffset + Get_Line_Number(9999)
        msFileText = Mid(msFileText, 10000)
    End If

    Do While Not (mobjTS.AtEndOfStream)
        strTemp = mobjTS.ReadLine
        msFileText = msFileText & strTemp & Chr(10)
        DoEvents
        If (mobjTS.AtEndOfStream) Then
            msFileText = msFileText & Chr(10)
        End If
        If Len(msFileText) >= 30000 And Not Load_All Then
            Exit Do
        End If
    Loop
    
    Get_File_Bite = msFileText

GoTo CleanUp
ErrHandler:
        gobjErrManager.Err_Log "File_Load_Manager", "Get_File_Bite", Err.Number, Err.Description
CleanUp:
End Function

Public Function Get_Line() As String
On Error GoTo ErrHandler

    Dim lPosition As Long
    Dim lLineEnd As Long
    
    If mlStartPosition = 0 Then
        lPosition = 1
    Else
        lPosition = mlStartPosition
    End If

    lLineEnd = InStr(lPosition, msFileText, Chr(10))
    If lLineEnd > 0 Then
        lLineEnd = lLineEnd - 1
    End If
    
    Get_Line = Mid(msFileText, lPosition, (lLineEnd - lPosition) + 1)
    
    mlStartPosition = lPosition + (lLineEnd - lPosition) + 2
    If mobjTS.AtEndOfStream And mlStartPosition > Len(msFileText) Then
        mbEOF = True
    End If
    
GoTo CleanUp
ErrHandler:
        gobjErrManager.Err_Log "File_Load_Manager", "Get_Line", Err.Number, Err.Description
CleanUp:
End Function

Private Function Get_Line_Number(plPosition As Long) As Long
On Error GoTo ErrHandler

    Dim lCount As Long
    Dim lPosition As Long
    
    lCount = 1
    lPosition = 1
    Do
        lPosition = InStr(lPosition, msFileText, gstr_strlinefeed)
        If (lPosition <= plPosition) And (lPosition <> 0) Then
            lCount = lCount + 1
        Else
            Exit Do
        End If
        lPosition = lPosition + 1
    Loop

    Get_Line_Number = lCount

GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "File_Load_Manager", "Get_Line_Number", Err.Number, Err.Description
CleanUp:
End Function


Public Property Get EOF() As Boolean
    EOF = mbEOF
End Property


Public Function Instr_Bite(Search_String As String, Optional lStart As Long = 1) As Long
On Error GoTo ErrHandler
    
    Instr_Bite = InStr(lStart, msFileText, Search_String)
    
GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "File_Load_Manager", "Instr_Bite", Err.Number, Err.Description
CleanUp:
End Function


