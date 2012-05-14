VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "parser_sql_script"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mobjFSO As Object 'FileSystemObject
Private mobjTS As Object 'TextStream
Private mobjParseDefinitions As parser_def_collection
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
Private miLineOffset As Long

Public Function EOF() As Boolean
    EOF = mbEOF
End Function

Private Function Get_Line_Number(piPosition) As Long
On Error GoTo ErrHandler

    Dim iCount As Long
    Dim iPosition As Long
    
    iCount = 1
    iPosition = 1
    Do
        iPosition = InStr(iPosition, msFileText, gstr_strlinefeed)
        If (iPosition <= piPosition) And (iPosition <> 0) Then
            iCount = iCount + 1
        Else
            Exit Do
        End If
        iPosition = iPosition + 1
    Loop

    Get_Line_Number = iCount

GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "Parser_SQL_Script", "Get_Line_Number", Err.Number, Err.Description
CleanUp:
End Function

Private Function Get_Next_Statement(piFilePosition As Long, _
                                    piStatementStart As Long, _
                                    piStatementEnd As Long, _
                                    pbTopStatement As Boolean, _
                                    ptypedef As type_def_enum, _
                                    piLineNumber As Long) As Boolean
On Error GoTo ErrHandler
Get_Next_Statement = False

    Dim iFilePosition As Long
    Dim iParseObject As Long
    Dim iStatementStart As Long
    Dim iStatementEnd As Long
    Dim iStartDefinition As Long
    Dim typedef As type_def_enum
    Dim iLineNumber As Long

    iStatementEnd = -1
    iFilePosition = piFilePosition
    iStatementStart = piStatementStart
    Get_Next_Statement = False
    iStartDefinition = 0
    piStatementStart = -1

    Do While (iFilePosition <= Len(msFileText))

        For iParseObject = 1 To (mobjParseDefinitions.Count)
            If piStatementStart = -1 Then '-- If Statement start not found yet (not sub-statement)
                If Check_For_Statement_Start(mobjParseDefinitions.Item(iParseObject), iFilePosition, iStatementStart) Then
                    '-- Statement Start Found
                    piStatementStart = iStatementStart
                    iStartDefinition = iParseObject
                End If
            End If
            If piStatementStart > -1 And (iParseObject = iStartDefinition) Then
                If Check_For_Statement_End(mobjParseDefinitions.Item(iParseObject), iFilePosition, piStatementEnd) Then
                    piFilePosition = iFilePosition
                    Get_Next_Statement = True
                    ptypedef = mobjParseDefinitions.Item(iParseObject).type_def
                    Exit For
                End If
            End If
            DoEvents
        Next

        If Get_Next_Statement = True Then
            piLineNumber = Get_Line_Number(piStatementStart)
            Exit Do
        End If

        iFilePosition = iFilePosition + 1
        If iFilePosition > 250000 Then
            Get_File_Bite iFilePosition
            iFilePosition = iFilePosition - 100000
            If piStatementStart >= 100000 Then
                piStatementStart = piStatementStart - 100000
            End If
            If iStatementStart >= 100000 Then
                iStatementStart = iStatementStart - 100000
            End If
            'MsgBox "piStatementStart:" & CStr(iStatementStart) & "     iStatementStart:" & CStr(iStatementStart)
        End If
    Loop
    
    piFilePosition = iFilePosition
    
    If Get_Next_Statement = False And piStatementStart <> -1 Then
        If Len(Mid(msFileText, piStatementStart)) > 300 Then
            frmIncompleteStatement.txtStatement = Mid(msFileText, piStatementStart, 300)
        Else
            frmIncompleteStatement.txtStatement = Mid(msFileText, piStatementStart)
        End If
        frmIncompleteStatement.Show vbModal
    End If

GoTo CleanUp
ErrHandler:
        gobjErrManager.Err_Log "Parser_SQL_Script", "Get_Next_Statement", Err.Number, Err.Description
CleanUp:
End Function

Private Sub Get_Next_Word(Text As String, Start As Long, Start_Pos As Long, Length As Long, Word As String)
On Error GoTo ErrHandler

    Dim lStart As Long
    Dim lEnd As Long
    
    lStart = Start
    Do While (Mid(Text, lStart, 1) = " " Or Mid(Text, lStart, 1) = Chr(10))
        lStart = lStart + 1
        DoEvents
    Loop
    
    lEnd = lStart + 1
    
    Do While (Mid(Text, lEnd, 1) <> " " And Mid(Text, lEnd, 1) <> Chr(10))
        lEnd = lEnd + 1
        DoEvents
    Loop
    
    Start_Pos = lStart
    Length = lEnd - Start_Pos
    Word = Mid(Text, Start_Pos, Length)

GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "Parser_SQL_Script", "Get_Next_Word", Err.Number, Err.Description
CleanUp:
End Sub

Private Function Check_For_Word_Set(Word_String As String, Start As Long) As Boolean
On Error GoTo ErrHandler

    Check_For_Word_Set = True

    Dim objWords As Collection
    Dim i As Integer
    Dim strWord As String
    Dim lStartPos As Long
    Dim lLength As Long
    
    
    Set objWords = Split_Statement(Word_String)
    
    For i = 1 To objWords.Count
        Get_Next_Word msFileText, Start, lStartPos, lLength, strWord
        If LCase(strWord) <> LCase(objWords(i)) Then
            Check_For_Word_Set = False
            Exit For
        End If
    Next
    
GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "Parser_SQL_Script", "Check_For_Word_Set", Err.Number, Err.Description
CleanUp:
End Function
Public Function Next_Statement(Statement As String, Line_Number As Long, Optional iFilePositionStart As Long = 0) As Boolean
On Error GoTo ErrHandler

Next_Statement = False

    Dim iStatementStart As Long
    Dim iStatementEnd As Long
    Dim typedef As type_def_enum

'    If (1 / 0) = 1 Then
'    End If
    If Not mbEOF Then
        Do
            iStatementStart = -1
            iStatementEnd = -1
    
            Set mobjStrip = Nothing
            Set mobjStrip = New Collection
            Next_Statement = Get_Next_Statement(miFilePosition, iStatementStart, iStatementEnd, True, typedef, Line_Number)
    
            If miFilePosition >= Len(msFileText) And mobjTS.AtEndOfStream Then
                mbEOF = True
            End If
            
            If typedef = ignore_contents Then
                Statement = ""
            Else
                If Next_Statement Then
                    Statement = Mid$(msFileText, iStatementStart, (iStatementEnd - iStatementStart) + 1)
                    Statement = Trim(Statement)
                    If Mid(Statement, Len(Statement), 1) = ";" Then
                        Statement = Mid(Statement, 1, Len(Statement) - 1)
                    ElseIf Mid(Statement, Len(Statement), 1) = "/" Then
                        Statement = Mid(Statement, 1, Len(Statement) - 1)
                    ElseIf Mid(Statement, Len(Statement) - 2, 3) = gstr_strlinefeed & "/" & gstr_strlinefeed Then
                        Statement = Mid(Statement, 1, Len(Statement) - 3)
                    ElseIf Mid(Statement, Len(Statement), 1) = Chr(10) Then
                        Statement = Mid(Statement, 1, Len(Statement) - 1)
                    End If
                End If
            End If
        Loop While (Not mbEOF And Statement = "")
    End If

GoTo CleanUp
ErrHandler:
        gobjErrManager.Err_Log "Parser_SQL_Script", "Next_Statement", Err.Number, Err.Description
CleanUp:
End Function

Public Sub Parse_Start(File_Name As String)
On Error GoTo ErrHandler

    Set mobjFSO = CreateObject("Scripting.FileSystemObject") 'New FileSystemObject
    Set mobjTS = mobjFSO.OpenTextFile(File_Name, 1, False) 'mobjFSO.OpenTextFile(file_name, ForReading, False)
    msFileText = ""
    
'    Dim sTemp As String
'    Do While Not (mobjTS.AtEndOfStream)
'        sTemp = Trim(mobjTS.ReadLine)
'        If sTemp <> "" Then
'            msFileText = msFileText & sTemp & Chr(13)
'        End If
'        DoEvents
'    Loop
'    msFileText = msFileText & Chr(13)
    mbEOF = False
    miStartPosition = 0
    Get_File_Bite miStartPosition
    miStartPosition = 0
    
GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "Parser_SQL_Script", "Parse_Start", Err.Number, Err.Description
CleanUp:
End Sub

Public Sub Parse_End()
On Error GoTo ErrHandler

    mobjTS.Close
    Set mobjFSO = Nothing
    Set mobjTS = Nothing
    
GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "Parser_SQL_Script", "Parse_End", Err.Number, Err.Description
CleanUp:
End Sub

Public Property Get parse_definitions() As Variant
    Set parse_definitions = mobjParseDefinitions
End Property

Public Property Let parse_definitions(ByVal vNewValue As Variant)
    Set mobjParseDefinitions = vNewValue
End Property

Private Function Check_For_Statement_Start(pobjParser As parser_definition, piFilePosition As Long, piStatementStartPosition As Long) As Boolean
On Error GoTo ErrHandler

    Check_For_Statement_Start = False

    Dim strCompare As Variant
    Dim bMatch As Boolean
    Dim iStatementStartPosition As Long
    Dim iFilePosition As Long
    
    'iStatementStartPosition = piStatementStartPosition
    iFilePosition = piFilePosition

    For Each strCompare In pobjParser.start_precede
        If (iFilePosition = 0) Then
            If ((strCompare = gstr_anycharacter) Or (strCompare = gstr_strlinefeed)) Then
                bMatch = True
                Exit For
            End If
        ElseIf (piFilePosition > 0) Then
            If (strCompare = gstr_anycharacter) Then
                bMatch = True
                Exit For
            ElseIf (UCase(strCompare) = UCase(Mid(msFileText, piFilePosition, 1))) Then
                bMatch = True
                Exit For
            End If
        End If
    Next
    iFilePosition = iFilePosition + 1

    '--- Check for Start Character(s)
    If bMatch Then
        bMatch = False
        strCompare = pobjParser.start_def

        If (UCase(strCompare) = UCase(Mid(msFileText, iFilePosition, Len(strCompare)))) Then
            bMatch = True
            iStatementStartPosition = iFilePosition
            iFilePosition = iFilePosition + Len(strCompare)
        End If
    End If

    '--- Check for Start Preceding Character
    If bMatch Then
        bMatch = False
        For Each strCompare In pobjParser.start_follow
            If (strCompare = gstr_notapplicable) Then
                piStatementStartPosition = iStatementStartPosition
                piFilePosition = iFilePosition
                bMatch = True
                Exit For
            ElseIf (strCompare = gstr_anycharacter) Or (UCase(strCompare) = UCase(Mid(msFileText, iFilePosition, Len(strCompare)))) Then
                piStatementStartPosition = iStatementStartPosition
                piFilePosition = iFilePosition + 1
                bMatch = True
                Exit For
            End If
        Next
    End If

    Check_For_Statement_Start = bMatch
    
GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "Parser_SQL_Script", "Check_For_Statement_Start", Err.Number, Err.Description
CleanUp:
End Function



Private Function Check_For_Statement_End(pobjParser As parser_definition, piFilePosition As Long, piStatementEndPosition As Long) As Boolean
On Error GoTo ErrHandler

    Check_For_Statement_End = False

    Dim strCompare As Variant
    Dim bMatch As Boolean
    Dim iStatementEndPosition As Long
    Dim iFilePosition As Long

    iFilePosition = piFilePosition

    For Each strCompare In pobjParser.end_precede
        If (iFilePosition = 0) Then
            If (strCompare = gstr_strlinefeed) Then
                iFilePosition = iFilePosition + 1
                bMatch = True
                Exit For
            ElseIf ((strCompare = gstr_anycharacter) Or (strCompare = gstr_strlinefeed)) Then
                iFilePosition = iFilePosition + 1
                bMatch = True
                Exit For
            End If
        ElseIf (iFilePosition > 0) Then
            If (strCompare = gstr_notapplicable) Then
                bMatch = True
                Exit For
            ElseIf (strCompare = gstr_anycharacter) Then
                iFilePosition = iFilePosition + 1
                bMatch = True
                Exit For
            ElseIf (UCase(strCompare) = UCase(Mid(msFileText, iFilePosition, 1))) Then
                iFilePosition = iFilePosition + 1
                bMatch = True
                Exit For
            End If
        End If
    Next

    'iFilePosition = iFilePosition + 1
    
    '--- Check for End Character(s)
    If bMatch Then
        bMatch = False
        For Each strCompare In pobjParser.end_def
            If (UCase(strCompare) = UCase(Mid(msFileText, iFilePosition, Len(strCompare)))) Then
                bMatch = True
                iStatementEndPosition = iFilePosition + Len(strCompare) - 1
                iFilePosition = iFilePosition + Len(strCompare)
                Exit For
            End If
        Next
    End If
    
    
    '--- Check for Start Preceding Character
    If bMatch Then
        If (strCompare = gstr_strlinefeed) Then
            piFilePosition = iStatementEndPosition
            'MsgBox Asc(Mid(msFileText, piFilePosition, 1))
            piStatementEndPosition = iStatementEndPosition
        Else
            bMatch = False
        
            For Each strCompare In pobjParser.end_follow
                If (strCompare = gstr_notapplicable) Then
                    piFilePosition = iStatementEndPosition + 1
                    piStatementEndPosition = iStatementEndPosition
                    'MsgBox Mid(msFileText, piStatementEndPosition, 1)
                    bMatch = True
                    Exit For
                ElseIf (strCompare = gstr_anycharacter) Then
                    piFilePosition = iStatementEndPosition + 1
                    piStatementEndPosition = iStatementEndPosition
                    'MsgBox Mid(msFileText, piStatementEndPosition, 1)
                    bMatch = True
                    Exit For
                ElseIf (UCase(strCompare) = UCase(Mid(msFileText, iFilePosition, Len(strCompare)))) Then
                    piFilePosition = iStatementEndPosition + 1
                    piStatementEndPosition = iStatementEndPosition
                    'MsgBox Mid(msFileText, piStatementEndPosition, 1)
                    bMatch = True
                    Exit For
                End If
            Next
        End If
    End If

    Check_For_Statement_End = bMatch
    
GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "Parser_SQL_Script", "Check_For_Statement_End", Err.Number, Err.Description
CleanUp:
End Function

Private Sub Class_Initialize()
    mbFileLoaded = False
    miLineOffset = 0
End Sub

Private Sub Get_File_Bite(piFile_Position As Long)
On Error GoTo ErrHandler

    Dim sTemp As String
       
    If piFile_Position <> 0 Then
        miLineOffset = miLineOffset + Get_Line_Number(99999)
        msFileText = Mid(msFileText, 100000)
    End If

    Do While Not (mobjTS.AtEndOfStream)
        sTemp = Trim(mobjTS.ReadLine)
        'If sTemp <> "" Then
            msFileText = msFileText & sTemp & Chr(10)
        'End If
        DoEvents
        If (mobjTS.AtEndOfStream) Then
            msFileText = msFileText & Chr(10)
        End If
        If Len(msFileText) >= 300000 Then
            Exit Do
        End If
    Loop
    
   ' piFile_Position = 1

GoTo CleanUp
ErrHandler:
        gobjErrManager.Err_Log "Parser_SQL_Script", "Get_File_Bite", Err.Number, Err.Description
CleanUp:
End Sub

