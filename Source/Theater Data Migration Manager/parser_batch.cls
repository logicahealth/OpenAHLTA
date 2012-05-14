VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "parser_batch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mobjFSO As Object 'FileSystemObject
Private mobjTS As Object 'TextStream
Private mobjFileLines As Collection
Private miCurrentLine As Long

Private objParseDefinitions As parser_def_collection
Private msFileText As String
Private msLine As String
Private miLastPos As Long
Private msLastStatement As String
Private miCurrentPos As Long
Private mbFileLoaded As Boolean

Private mbEOF As Boolean
Private miStartPosition As Long
Private miFilePosition As Long
Private miStatementStart As Long
Private miStatementEnd As Long
Private mobjStrip As Collection

Public Function EOF() As Boolean
    EOF = mbEOF
End Function

Public Sub Next_Statement(Statement As parser_batch_line)
On Error GoTo ErrHandler

    Dim iStatementStart As Long
    Dim iStatementEnd As Long
    Dim typedef As type_def_enum

    If Not mbEOF Then
        Set Statement = mobjFileLines.Item(miCurrentLine)
        miCurrentLine = miCurrentLine + 1
        If miCurrentLine > mobjFileLines.Count Then
            mbEOF = True
        End If
    End If

GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "Parser_Batch", "Next_Statement", Err.Number, Err.Description
CleanUp:
End Sub

Public Sub Parse_Start(file_name As String)
On Error GoTo ErrHandler

    Set mobjFSO = CreateObject("Scripting.FileSystemObject") 'New FileSystemObject
    Set mobjTS = mobjFSO.OpenTextFile(file_name, 1, False) 'mobjFSO.OpenTextFile(file_name, ForReading, False)
    
    Dim sTemp As String
    Dim objColl As Collection
    Dim TempArray As Variant
    Dim i As Long
    Dim iLineCount As Long
    Dim objBatchLine As parser_batch_line
    Dim objCommand As Object
    Dim sFilePath As String
    Dim iSlashCount As Long
    Dim iLine As Long

    Set mobjFileLines = Nothing
    Set mobjFileLines = New Collection
    iLineCount = 0
    iLine = 0
    miCurrentLine = 0

    sFilePath = StrReverse(file_name)
    sFilePath = StrReverse(Mid(sFilePath, InStr(sFilePath, "\") + 1))
    'iSlashCount = StringCount(file_name, "\")
    
    Do While Not (mobjTS.AtEndOfStream)
        sTemp = Trim(mobjTS.ReadLine)
        iLine = iLine + 1
        If sTemp <> "" Then
            Set objColl = Split_Statement(sTemp)
            iLineCount = iLineCount + 1
            'TempArray = Split(sTemp, " ")
            
'            Set objColl = New Collection
'            For i = LBound(TempArray) To UBound(TempArray)
'                If i = 0 Then
'                    objColl.Add LCase(TempArray(i)), CStr(i)
'                ElseIf Len(Trim(TempArray(i))) > 0 Then
'                    objColl.Add TempArray(i), CStr(i)
'                End If
'            Next
            Set objBatchLine = New parser_batch_line
            objBatchLine.Statement = sTemp
            objBatchLine.Line_Number = iLine
            Select Case LCase(objColl.Item(1))
                Case "sqlplus"
                    Parser_Sqlplus objColl, sFilePath, objCommand
                    Set objBatchLine.CommandObject = objCommand
                    objBatchLine.Batch_Command = sqlplus
                    mobjFileLines.Add objBatchLine, CStr(iLineCount)
                Case "del"
                    objBatchLine.Batch_Command = del
                    objBatchLine.CommandObject = Get_Filename(objColl.Item(2))
                    mobjFileLines.Add objBatchLine, CStr(iLineCount)
                Case "regedit"
                    objBatchLine.Batch_Command = regedit
                    objBatchLine.CommandObject = sTemp
                    mobjFileLines.Add objBatchLine, CStr(iLineCount)
                Case "imp"
                    objBatchLine.Batch_Command = import
                    objBatchLine.CommandObject = sTemp
                    mobjFileLines.Add objBatchLine, CStr(iLineCount)
                Case "exp"
                    objBatchLine.Batch_Command = export
                    objBatchLine.CommandObject = sTemp
                    mobjFileLines.Add objBatchLine, CStr(iLineCount)
                    
            End Select
           
        End If
        DoEvents
    Loop
    
    If mobjFileLines.Count > 0 Then
        mbEOF = False
        miCurrentLine = 1
    Else
        mbEOF = True
    End If
    mobjTS.Close

GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "Parser_Batch", "Parse_Start", Err.Number, Err.Description
    If Not mobjTS Is Nothing Then
        mobjTS.Close
    End If
CleanUp:
    Set mobjFSO = Nothing
    Set mobjTS = Nothing
End Sub

Public Sub Parse_End()
On Error GoTo ErrHandler

    Dim obj As Object
    For Each obj In mobjFileLines
        Set obj = Nothing
    Next
    Set mobjFSO = Nothing
    Set mobjTS = Nothing

GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "Parser_Batch", "Parse_End", Err.Number, Err.Description
    If Not mobjTS Is Nothing Then
        mobjTS.Close
    End If
CleanUp:
End Sub

Public Sub Parser_Sqlplus(objColl As Collection, Filepath As String, objSqlplus_Def)
On Error GoTo ErrHandler

Dim TempString
Dim TempArray

Set objSqlplus_Def = Nothing
Set objSqlplus_Def = New Parser_Sqlplus_Def

For Each TempString In objColl
    If InStr(TempString, "/") > 0 And InStr(TempString, "@") Then
        TempArray = Split(TempString, "@")
        objSqlplus_Def.Host = TempArray(1)
        TempArray = Split(TempArray(0), "/")
        objSqlplus_Def.UID = TempArray(0)
        objSqlplus_Def.Password = TempArray(1)
    ElseIf Mid(TempString, 1, 1) = "@" Then
        objSqlplus_Def.Filename = Mid(TempString, 2)
        objSqlplus_Def.Filepath = Filepath
    ElseIf Mid(TempString, 1, 1) = "%" Then
        TempString = Replace(TempString, "%", "")
        objSqlplus_Def.Variables.Add Environ(TempString)
    ElseIf LCase(TempString) <> "sqlplus" Then
        objSqlplus_Def.Variables.Add TempString
    End If
Next


GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "Parser_Batch", "Parser_Sqlplus", Err.Number, Err.Description
CleanUp:
End Sub

Private Function Get_Filename(Filename As String) As String
On Error GoTo ErrHandler

    Dim TempArray
    Dim TempString As String
    Dim i As Long
    
    TempArray = Split(Filename, "%")
    TempString = ""
    
    For i = LBound(TempArray) To UBound(TempArray)
        If (i Mod 2) = 0 Then
            TempString = TempString & TempArray(i)
        Else
            TempString = TempString & Environ(TempArray(i))
        End If
    Next i
    
    Get_Filename = TempString
    'TempString = Replace(TempString, "%", "")

GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "Parser_Batch", "Get_Filename", Err.Number, Err.Description
CleanUp:
End Function

