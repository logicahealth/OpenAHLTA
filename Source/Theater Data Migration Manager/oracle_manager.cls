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
Attribute VB_Name = "oracle_manager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mobj_SpoolManager As New Spool_Manager
Private msCurrent_Folder As String
Private mbSpooling As Boolean
Private mobjParameters As New Collection
Private mobjOraSession As Object
Private mobjOraDatabase As Object 'OracleInProcServer.OraDatabase
Private msPatch_Get_Lines As String
Private msGet_Lines_Newcode As String
Private msOutput_Enable As String

Public Sub Commit_Transaction()
On Error GoTo ErrHandler

    mobjOraDatabase.CommitTrans
    
GoTo CleanUp
ErrHandler:
        gobjErrManager.Err_Log "oracle_manager", "Commit_Transaction", Err.Number, Err.Description
CleanUp:
End Sub

Public Sub Rollback_Transaction()
On Error GoTo ErrHandler

    mobjOraDatabase.Rollback
    
GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "oracle_manager", "Rollback_Transaction", Err.Number, Err.Description
CleanUp:
End Sub
Public Sub Begin_Transaction()
On Error GoTo ErrHandler

    mobjOraDatabase.BeginTrans

GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "oracle_manager", "Begin_Transaction", Err.Number, Err.Description
CleanUp:
End Sub
Public Sub Connect(Host As String, UserId As String, Password As String, ErrorMessage As String, ErrorCode As Long)
On Error GoTo ErrHandler

'    Dim Message As String
'    Dim ErrorCode As Long
    Dim bConnected As Boolean
    bConnected = False

    Set mobjOraSession = CreateObject("OracleInProcServer.XOraSession")
    
'    Dim x As OracleInProcServer.OraSessionClass
'    Dim y As OraDatabase
'    y.Parameters
'    y.Session
    
    On Error Resume Next
    Set mobjOraDatabase = mobjOraSession.DbOpenDatabase(Host, UserId & "/" & Password, 0&)
    If Err.Number <> 0 Then
        ErrorMessage = Err.Description
        ErrorCode = Err.Number
    Else
        bConnected = True
    End If
    On Error GoTo ErrHandler
    
    If bConnected Then
        Execute_SQL msPatch_Get_Lines, ErrorMessage, ErrorCode
        If Not mobjOraDatabase Is Nothing Then
            mobjOraDatabase.BeginTrans
        End If
    End If

GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "Oracle_Manager", "Connect", Err.Number, Err.Description
CleanUp:
End Sub

Public Sub Disconnect()
On Error GoTo ErrHandler

    
    ' Close the Oracle SQL/NET connection
    mobjOraDatabase.Parameters.Remove ("dbPatchLines")
    mobjOraDatabase.Close
    Set mobjOraDatabase = Nothing
    Set mobjOraSession = Nothing


GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "oracle_manager", "Disconnect", Err.Number, Err.Description
CleanUp:
End Sub

Public Sub Run_SQL(Statement As String, Comment As String, Error_Message As String, Error_Code As Long)
On Error GoTo ErrHandler

    Dim sTempStatement As String
    
    Comment = ""
    Error_Message = ""
    Error_Code = 0

    If LCase(Mid(Statement, 1, 3)) = "set" Then
        Comment = "Statement Ignored"
    ElseIf LCase(Mid(Statement, 1, 5)) = "spool" Then
        Spool Statement
    ElseIf LCase(Mid(Statement, 1, 7)) = "declare" Or LCase(Mid(Statement, 1, 3)) = "begin" Then
        Execute_Inline_PLSQL Statement, Error_Message, Error_Code
    ElseIf LCase(Mid(Statement, 1, 4)) = "drop" Or LCase(Mid(Statement, 1, 6)) = "create" Or _
           LCase(Mid(Statement, 1, 5)) = "grant" Or LCase(Mid(Statement, 1, 5)) = "alter" Or _
           LCase(Mid(Statement, 1, 6)) = "delete" Or LCase(Mid(Statement, 1, 6)) = "insert" Or _
           LCase(Mid(Statement, 1, 6)) = "update" Or LCase(Mid(Statement, 1, 6)) = "select" Or _
           LCase(Mid(Statement, 1, 8)) = "truncate" Then
        sTempStatement = Replace_Variables(Statement)
        Execute_SQL sTempStatement, Error_Message, Error_Code
    ElseIf LCase(Mid(Statement, 1, 6)) = "commit" Then
        mobjOraDatabase.CommitTrans
        mobjOraDatabase.BeginTrans
    ElseIf LCase(Mid(Statement, 1, 1)) = "@" Then
        MsgBox Statement
    End If

GoTo CleanUp
ErrHandler:
        gobjErrManager.Err_Log "oracle_manager", "Run_SQL", Err.Number, Err.Description
CleanUp:
End Sub

Private Sub Spool(Statement As String)
On Error GoTo ErrHandler
    
    Dim objColl As Collection
    Dim strFile As String

    Set objColl = Split_Statement(Statement)
    If Mid(LCase(objColl.Item(2)), 1, 3) = "off" Then
        mobj_SpoolManager.Spool_Off
        mbSpooling = False
    Else
        strFile = objColl.Item(2)
        strFile = Replace_Variables(strFile)
        If InStr(strFile, "\") > 0 Then
            mobj_SpoolManager.Spool_On strFile
        Else
            mobj_SpoolManager.Spool_On msCurrent_Folder & "\" & strFile
        End If
        mbSpooling = True
    End If

GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "oracle_manager", "Spool", Err.Number, Err.Description
CleanUp:
End Sub

Private Sub Execute_SQL(sql As String, Message As String, ErrorCode As Long)
On Error GoTo ErrHandler

    Const ORAPARM_INPUT = 1: Const ORAPARM_OUTPUT = 2
    Dim MyQuery As Object 'OraSqlStmt
    Dim objOraDynaset As Object 'OraDynaset

    mobjOraDatabase.Parameters.Remove ("dbPatchLines")
    mobjOraDatabase.Parameters.Add "dbPatchLines", "nada", 2 'ORAPARM_OUTPUT
    mobjOraDatabase.Parameters.Item("dbPatchLines").serverType = 1 'ORATYPE_VARCHAR2
    mobjOraDatabase.Parameters.Item("dbPatchLines").MinimumSize = 4000
    mobjOraDatabase.Parameters.Item("dbPatchLines").AutoBindEnable
    
    'Set MyQuery = mobjOraDatabase.CreateSql(SQL, 0)
    'MyQuery.Refresh
    
    Set objOraDynaset = mobjOraDatabase.CreateSql(sql, CInt(0))

    Message = mobjOraDatabase.LastServerErrText
    ErrorCode = mobjOraDatabase.LastServerErr
    mobjOraDatabase.LastServerErrReset
    
    If mbSpooling And mobjOraDatabase.Parameters("dbPatchLines").Value <> "nada" Then
        If Not mobj_SpoolManager.Put_Line(mobjOraDatabase.Parameters("dbPatchLines").Value) Then GoTo CleanUp
    End If

GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "oracle_manager", "Execute_SQL", Err.Number, Err.Description
CleanUp:
    objOraDynaset.Close
    Set objOraDynaset = Nothing
End Sub

Public Function Split_Statement(Statement As String) As Collection
On Error GoTo ErrHandler

    Dim TempArray
    Dim i As Long
    Dim TempString As String
    Dim objColl As Collection
    
    Set objColl = Nothing
    Set objColl = New Collection
    
    TempArray = Split(Statement, " ")
    For i = LBound(TempArray) To UBound(TempArray)
        TempString = Trim(TempArray(i))
        If TempString <> "" Then
            objColl.Add TempString
        End If
    Next

    Set Split_Statement = objColl

GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "oracle_manager", "Split_Statement", Err.Number, Err.Description
CleanUp:
    Set objColl = Nothing
End Function

Public Property Get Current_Folder() As String
    Current_Folder = msCurrent_Folder
End Property

Public Property Let Current_Folder(ByVal vNewValue As String)
    msCurrent_Folder = vNewValue
End Property

Private Sub Class_Initialize()
On Error GoTo ErrHandler

    mbSpooling = False
'    msPatch_Get_Lines = _
'        "CREATE OR REPLACE FUNCTION Patch_Get_Lines RETURN VARCHAR2 IS" & Chr(10) & _
'        "v_status NUMBER := 4000;" & Chr(10) & _
'        "v_lines  VARCHAR2(4000) := '';" & Chr(10) & _
'        "v_line_array dbms_output.CHARARR;" & Chr(10) & _
'        "BEGIN" & Chr(10) & _
'        "dbms_output.get_lines(v_line_array, v_status);" & Chr(10) & _
'        "FOR linenum IN 1 .. (v_status) LOOP" & Chr(10) & _
'        "v_lines := v_lines || NVL(v_line_array(linenum),'') || CHR(13) || CHR(10);" & Chr(10) & _
'        "END LOOP;" & Chr(10) & _
'        "dbms_output.put('');" & Chr(10) & _
'        "RETURN v_lines;" & Chr(10) & _
'        "END;"
'
    msPatch_Get_Lines = _
        "CREATE OR REPLACE PROCEDURE Patch_Get_Lines (p_lines OUT VARCHAR2) IS" & Chr(10) & _
        "v_status NUMBER := 4000;" & Chr(10) & _
        "v_lines  VARCHAR2(4000) := '';" & Chr(10) & _
        "v_line_array dbms_output.CHARARR;" & Chr(10) & _
        "BEGIN" & Chr(10) & _
        "dbms_output.get_lines(v_line_array, v_status);" & Chr(10) & _
        "FOR linenum IN 1 .. (v_status) LOOP" & Chr(10) & _
        "v_lines := v_lines || NVL(v_line_array(linenum),'') || CHR(13) || CHR(10);" & Chr(10) & _
        "END LOOP;" & Chr(10) & _
        "p_lines := v_lines;" & Chr(10) & _
        "END;"
        
        
     'msGet_Lines_Newcode = Chr(10) & ":dbPatchLines := patch_get_lines();" & Chr(10)
     msGet_Lines_Newcode = Chr(10) & "patch_get_lines(:dbPatchLines);" & Chr(10)
     
     msOutput_Enable = Chr(10) & "DBMS_OUTPUT.ENABLE;" & Chr(10)
     
GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "oracle_manager", "Class_Initialize", Err.Number, Err.Description
CleanUp:
End Sub

Public Property Get Variables() As Variant
On Error GoTo ErrHandler

    Set Variables = mobjParameters
    
GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "oracle_manager", "Variables(Get)", Err.Number, Err.Description
CleanUp:
End Property

Public Property Let Variables(ByVal vNewValue As Variant)
On Error GoTo ErrHandler

    Set mobjParameters = vNewValue
    
GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "oracle_manager", "Variables(Let)", Err.Number, Err.Description
CleanUp:
End Property

Private Sub Class_Terminate()
On Error GoTo ErrHandler

    Set mobjParameters = Nothing
    Set mobj_SpoolManager = Nothing
    Set mobjOraDatabase = Nothing
    Set mobjOraSession = Nothing

GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "oracle_manager", "Class_Terminate", Err.Number, Err.Description
CleanUp:
End Sub

Private Function Replace_Variables(Statement As String) As String
On Error GoTo ErrHandler
    
    Dim variable As Variant
    Dim NewStatement As String
    Dim iCount As Long
    
    NewStatement = Statement
    iCount = 0
    
    For Each variable In mobjParameters
        iCount = iCount + 1
        NewStatement = Replace(NewStatement, "&&" & CStr(iCount), variable)
        NewStatement = Replace(NewStatement, "&" & CStr(iCount), variable)
    Next
    Replace_Variables = NewStatement
    
GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "oracle_manager", "Replace_Variables", Err.Number, Err.Description
CleanUp:
End Function

Private Sub Execute_Inline_PLSQL(Statement As String, Message As String, ErrorCode As Long)
On Error GoTo ErrHandler

    Dim TempStatement As String
    Dim TempStatement2 As String
    Dim bCheckForOutput As Boolean
    Dim i As Long
    Dim iLen As Long
    Dim iException_Pos As Long
    Dim iEnd_Pos As Long
    Dim iBegin_Pos As Long
    
    If mbSpooling Then
        TempStatement2 = LCase(Statement)
        If InStr(TempStatement2, "dbms_output.put") > 0 Then
            bCheckForOutput = True
                       
            TempStatement = Statement
            TempStatement2 = StrReverse(TempStatement2)
            iLen = Len(TempStatement)
            iEnd_Pos = iLen - (InStr(1, TempStatement2, ";") + 4)
            
            TempStatement = Mid(TempStatement, 1, iEnd_Pos) & _
                            msGet_Lines_Newcode & _
                            Mid(TempStatement, iEnd_Pos + 1)

            'iLen = Len(TempStatement)
            'TempStatement2 = StrReverse(TempStatement)
            iException_Pos = InStr(1, TempStatement2, "noitpecxe")
            If iException_Pos > 0 Then
                iException_Pos = iLen - (iException_Pos + 9)
                TempStatement = Mid(TempStatement, 1, iException_Pos) & _
                                msGet_Lines_Newcode & _
                                Mid(TempStatement, iException_Pos + 1)
            End If
            
'            iBegin_Pos = InStr(1, TempStatement2, Chr(10) & "nigeb")
'            If iBegin_Pos > 0 Then
'                iBegin_Pos = iLen - (iBegin_Pos)
'                TempStatement = Mid(TempStatement, 1, iBegin_Pos) & _
'                                msOutput_Enable & _
'                                Mid(TempStatement, iBegin_Pos + 1)
'            End If

            TempStatement2 = StrReverse(TempStatement2)
            iBegin_Pos = InStr(1, TempStatement2, "begin" & Chr(10))
            If iBegin_Pos > 0 Then
                iBegin_Pos = iBegin_Pos + 5
                TempStatement = Mid(TempStatement, 1, iBegin_Pos) & _
                                msOutput_Enable & _
                                Mid(TempStatement, iBegin_Pos + 1)
            End If
            
            
        End If
    Else
        TempStatement = Statement
    End If
    
    Execute_SQL TempStatement, Message, ErrorCode

GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "oracle_manager", "Execute_Inline_PLSQL", Err.Number, Err.Description
CleanUp:
End Sub

Public Function Create_Sqlplus_Definition(sStatement As String, Parent_Sqlplus_Def As Parser_Sqlplus_Def, Parent_Path As String) As Parser_Sqlplus_Def
On Error GoTo ErrHandler

    Dim objParserSqlplusDef As New Parser_Sqlplus_Def
    Dim objColl As Collection
    Dim sTempStatement As String
    
    sTempStatement = Replace_Variables(sStatement)
    Set objColl = Split_Statement(sTempStatement)

    Parser_Sqlplus objColl, Parent_Path, Create_Sqlplus_Definition
    
'    ObjParserSqlplusDef.Filename =
'
'    If InStr(objColl.Item(1), "\") > 0 Then
'        Loop
'        ObjParserSql
'    Else
'        ObjParserSqlplusDef.FilePath = Parent_Sqlplus_Def.FilePath
'        ObjParserSqlplusDef.Filename = Mid(objColl.Item(1), 2)
'    End If
    
    
GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "oracle_manager", "Create_Sqlplus_Definition", Err.Number, Err.Description
CleanUp:
End Function

Public Sub Parser_Sqlplus(objColl As Collection, Filepath As String, objSqlplus_Def As Object)
On Error GoTo ErrHandler

Dim TempString
Dim TempArray
Dim sFilePath As String
Dim sFileName As String

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
        TempString = Mid(TempString, 2)
        
        If InStr(TempString, "\") > 0 Then
            sFilePath = StrReverse(TempString)
            objSqlplus_Def.Filepath = StrReverse(Mid(sFilePath, InStr(sFilePath, "\") + 1))
            objSqlplus_Def.Filename = StrReverse(Mid(sFilePath, 1, InStr(sFilePath, "\") - 1))
        Else
            objSqlplus_Def.Filename = TempString
            objSqlplus_Def.Filepath = Filepath
        End If
        If Mid(objSqlplus_Def.Filename, Len(objSqlplus_Def.Filename) - 3) <> ".sql" Then
            objSqlplus_Def.Filename = objSqlplus_Def.Filename & ".sql"
        End If
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
