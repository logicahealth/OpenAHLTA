VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRetirementProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private indentlevel As Integer
Public prgBarDataRetirement As ProgressBar
Private RSColRefData As ADODB.Recordset
Private iPurgeOrder As Integer
Public bLogging As Boolean
Public bDetailed As Boolean
Public sBaseTablename As String

Public Sub Log(Data As String)
    Dim I As Integer
    Dim Temp As String
    
    Temp = Data
    For I = 0 To indentlevel
        Temp = "    " & Temp
    Next
    Open "c:\logfile.txt" For Append As #1
    Print #1, Temp
    Close #1
End Sub

Private Function GetConstraintName(TableName As String) As String

    Dim sSQL As String
    Dim RS As ADODB.Recordset
    
    GetConstraintName = ""
    
    sSQL = "select constraint_name from all_constraints " & _
           " where owner = 'GEMS' " & _
           " and table_name = '" & TableName & "' and CONSTRAINT_TYPE='P'"
    
    If oDAS Is Nothing Then
       Set oDAS = oDatabase.CHCSII_DAS(Auto)  'GEMS_DAS
    End If
    Set RS = oDAS.OpenRecordset(sSQL, adOpenDynamic, adLockOptimistic)
   
    If RS.BOF = True And RS.EOF = True Then
        'no Primary Key
        RS.Close
        Exit Function
    End If
    GetConstraintName = RS.Fields(0).Value
    RS.Close
        
End Function

Private Function GetTableByConstraintName(ConstraintName As String) As String

    Dim sSQL As String
    Dim RS As ADODB.Recordset
    
    GetTableByConstraintName = ""
    
    sSQL = "select table_name from all_constraints " & _
           " where owner = 'GEMS' " & _
           " and constraint_name = '" & ConstraintName & "' and CONSTRAINT_TYPE='P'"
    
    Set RS = oDAS.OpenRecordset(sSQL, adOpenDynamic, adLockOptimistic)
   
    If RS.BOF = True And RS.EOF = True Then
        RS.Close
        Exit Function
    End If
    GetTableByConstraintName = RS.Fields(0).Value
    RS.Close
    
End Function

Private Function GetRefByTables(ByRef RSTableRefData As ADODB.Recordset, ConstraintName As String) As Boolean

    Dim sSQL As String
    Dim RS As ADODB.Recordset
    
    GetRefByTables = False
    
    sSQL = "select constraint_name, table_name from all_constraints " & _
           " where owner = 'GEMS' " & _
           " and r_constraint_name = '" & ConstraintName & "' and CONSTRAINT_TYPE='R'"
    Set RS = oDAS.OpenRecordset(sSQL, adOpenDynamic, adLockOptimistic)
    If RS.BOF = True And RS.EOF = True Then
        'No tables that rely on data from this table
        RS.Close
        Exit Function
    End If
    While RS.EOF = False
        With RSTableRefData
            .AddNew
            .Fields!CONSTRAINT_NAME = RS.Fields(0).Value
            .Fields!TABLE_NAME = RS.Fields(1).Value
            .Fields!R_CONSTRAINT_NAME = ConstraintName
            .Fields!CONSTRAINT_TYPE = "R"
        End With
        RS.MoveNext
    Wend
    
    RS.Close
    GetRefByTables = True
    
End Function

Private Function GetRefFromTables(ByRef RSTableRefData As ADODB.Recordset, TableName As String) As Boolean

    Dim sSQL As String
    Dim RS As ADODB.Recordset
    
    GetRefFromTables = False
    
    sSQL = "select constraint_name, r_constraint_name from all_constraints " & _
           " where owner = 'GEMS' " & _
           " and table_name = '" & TableName & "' and CONSTRAINT_TYPE='R'"
    Set RS = oDAS.OpenRecordset(sSQL, adOpenDynamic, adLockOptimistic)
    
    If RS.BOF = True And RS.EOF = True Then
        'No tables that rely on data from this table
        RS.Close
        Exit Function
    End If
    
    While RS.EOF = False
        With RSTableRefData
            .AddNew
            .Fields!CONSTRAINT_NAME = RS.Fields(0).Value
            .Fields!TABLE_NAME = TableName
            .Fields!R_CONSTRAINT_NAME = RS.Fields(1).Value
            .Fields!CONSTRAINT_TYPE = "R"
        End With
        RS.MoveNext
    Wend
    
    RS.Close
    GetRefFromTables = True
    
End Function

Private Function BuildSQL(Table1Name As String, Table1Cols As String, _
                          Table2Name As String, Table2Cols As String) As String

    Dim Temp As String
    Dim aTable1Cols() As String
    Dim aTable2Cols() As String
    Dim iCounter As Integer
    
    BuildSQL = ""
    Temp = "where "
    
    aTable1Cols = Split(Table1Cols, ",")
    aTable2Cols = Split(Table2Cols, ",")
    
    For iCounter = 0 To UBound(aTable1Cols)
        If Temp = "where " Then
        Else
            Temp = Temp & " and "
        End If
        Temp = Temp & Table2Name & "." & aTable2Cols(iCounter) & " = "
        Temp = Temp & Table1Name & "." & aTable1Cols(iCounter)
    Next
    
    BuildSQL = Temp
        
End Function

Public Function CopyData(TableName As String, WhereClause As String) As Boolean

On Error GoTo ErrHandler
    
    Dim sSQL As String
    Dim sSQL2 As String
    Dim sConstraintName As String
    Dim RS As ADODB.Recordset
    Dim RS2 As ADODB.Recordset
    Dim Temp As String
    Dim oConn As New CHCSII_CONN.Conn
    
    Dim RSTableRefData As ADODB.Recordset
    
    Dim vEntry As Field
    Dim I As Integer
    Dim I2 As Integer
    Dim sFieldName As String
    
    CopyData = False
    indentlevel = indentlevel + 1
    
    'Make sure tablename is uppercase
    TableName = UCase(TableName)
        
    'get the data
    sSQL = "Select * from " & TableName & " " & WhereClause
    Set RS = oDAS.OpenRecordset(sSQL)
    If RS.EOF = True Then
        'if there is no data to archive in this table there will be no data in the tables in refernces
        CopyData = True
        If bLogging And bDetailed Then
            Log "No data found for: " & sSQL
        End If
        GoTo EndFunct
    End If
    sSQL2 = "Select * from GEMSA." & TableName & " " & WhereClause
    Set RS2 = oDAS.OpenRecordset(sSQL2, adOpenDynamic, adLockOptimistic)
    If RS2.EOF = True Then
        'Add the new data to GEMSA
        If bLogging Then
            Log sSQL
            'Log sSQL2
        End If
        While RS.EOF = False
            RS2.AddNew
            'SCR 89413 log insert statement into GEMSA
            sSQL = "insert into GEMSA." & TableName & " ("
            For Each vEntry In RS.Fields
                For I = 0 To RS2.Fields.Count - 1
                    If vEntry.Name = RS2.Fields(I).Name Then
                        sSQL = sSQL & RS2.Fields(I).Name & ","
                    End If
                Next
            Next vEntry
            sSQL = Left(sSQL, Len(sSQL) - 1) & ") Values ("
            For Each vEntry In RS.Fields
                For I = 0 To RS2.Fields.Count - 1
                    If vEntry.Name = RS2.Fields(I).Name Then
                        RS2.Fields(I).Value = vEntry.Value
                        
                        If IsNull(vEntry.Value) Then
                            sSQL = sSQL & "NULL,"
                        ElseIf vEntry.Type = adLongVarBinary Then
                            'cannot copy blob data through SQL so just change to NULL for logging purposes
                            sSQL = sSQL & "NULL,"
                        ElseIf vEntry.Type = adNumeric Or vEntry.Type = adVarNumeric Then
                            sSQL = sSQL & vEntry.Value & ","
                        ElseIf vEntry.Type = adDate _
                          Or vEntry.Type = adDBDate _
                          Or vEntry.Type = adDBTime _
                          Or vEntry.Type = adDBTimeStamp Then
                            ' does this date have a timestamp also?
                            ' possible formats: '1/1/2006'
                            ' or: '10/20/2005 2:37:06 AM'
                            If InStr(1, vEntry.Value, ":", 1) = 0 Then
                                ' no time substring
                                sSQL = sSQL & oConn.SQLDate(vEntry.Value, DD_MON_YYYY) & ","
                            Else
                                ' includes time
                                sSQL = sSQL & oConn.SQLDate(vEntry.Value, DD_MON_YYYY_HH_MM_SS) & ","
                            End If
                        Else
                            sSQL = sSQL & "'" & vEntry.Value & "',"
                        End If
                        
                        I = RS2.Fields.Count
                    End If
                Next
            Next vEntry
            RS2.Update
            sSQL = Left(sSQL, Len(sSQL) - 1) & ")"
            If bLogging Then
                Log sSQL
            End If
            oDatabase.AuditSQL "GEMSA." & TableName, sSQL, oDAS.ActiveConnection
            RS.MoveNext
        Wend
    Else
        'if the data is already archived in this table, we do not need to rearchive it
        If bLogging And bDetailed Then
            Log "Bypassed as already archived: " & sSQL2
        End If
        CopyData = True
        GoTo EndFunct
    End If
    
    sSQL = "Select columns, r_table_name, r_columns from RET_REFRULES where table_name = '" & TableName & "'"
    Set RSTableRefData = oDAS.OpenRecordset(sSQL)
    If RSTableRefData.EOF = True Then
        If bLogging And bDetailed Then
            Log "No references set for " & TableName
        End If
        CopyData = True
        GoTo EndFunct
    End If
    RS.MoveFirst
    While RS.EOF = False
        'Get the tables that need to be copied
        RSTableRefData.MoveFirst
        While RSTableRefData.EOF = False
            Dim brunCopyData As Boolean
            Dim bFound As String
            brunCopyData = True
            Temp = BuildSQL(TableName, RSTableRefData.Fields!Columns, RSTableRefData.Fields!R_TABLE_NAME, RSTableRefData.Fields!R_COLUMNS)
            I = InStr(Temp, TableName & ".")
            While I > 0
                I2 = InStr(I, Temp, " ")
                If I2 = 0 Then
                    sFieldName = Mid(Temp, I)
                Else
                    sFieldName = Mid(Temp, I, I2 - I)
                End If
                sFieldName = Replace(sFieldName, TableName & ".", "")
                bFound = False
                For Each vEntry In RS.Fields
                    If vEntry.Name = sFieldName Then
                        If IsNull(vEntry.Value) Then
                            Temp = Replace(Temp, TableName & "." & sFieldName, "NULL")
                            brunCopyData = False
                            bFound = True
                        Else
                            ' SCR 85864 (JQ) add single quotes around non-numeric fields:
                            If vEntry.Type = adNumeric Or vEntry.Type = adVarNumeric Then
                                Temp = Replace(Temp, TableName & "." & sFieldName, vEntry.Value)
                            ElseIf vEntry.Type = adDate _
                              Or vEntry.Type = adDBDate _
                              Or vEntry.Type = adDBTime _
                              Or vEntry.Type = adDBTimeStamp Then
                                ' does this date have a timestamp also?
                                ' possible formats: '1/1/2006'
                                ' or: '10/20/2005 2:37:06 AM'
                                If InStr(1, vEntry.Value, ":", 1) = 0 Then
                                    ' no time substring
                                    Temp = Replace(Temp, TableName & "." & sFieldName, oConn.SQLDate(vEntry.Value, DD_MON_YYYY))
                                Else
                                    ' includes time
                                    Temp = Replace(Temp, TableName & "." & sFieldName, oConn.SQLDate(vEntry.Value, DD_MON_YYYY_HH_MM_SS))
                                End If
                                
                            Else
                                Temp = Replace(Temp, TableName & "." & sFieldName, "'" & vEntry.Value & "'")
                            End If
                            bFound = True
                        End If
                        Exit For
                    End If
                Next vEntry
                I = InStr(Temp, TableName & ".")
                If bFound = False Then
                    If bLogging Then
                        Log "DATA ERROR: Field Not found in main table"
                        Log "---------------------------------------------------------------------"
                    End If
                    I = 0
                    brunCopyData = False
                End If
            Wend
            If brunCopyData = True Then
                If CopyData(RSTableRefData.Fields!R_TABLE_NAME, Temp) = False Then
                    CopyData = False
                    GoTo EndFunct
                End If
            End If
            RSTableRefData.MoveNext
        Wend
        
        RS.MoveNext
        
        ' show progress on progress bar in terms of base table records (e.g., patients)
        If TableName = sBaseTablename And prgBarDataRetirement.Value < prgBarDataRetirement.Max Then
            prgBarDataRetirement.Value = prgBarDataRetirement.Value + 1
        End If
    
        DoEvents
        
    Wend
    
    CopyData = True

EndFunct:
    
    On Error Resume Next  ' if we're here, we're done with this pass
    Set RSTableRefData = Nothing
    Set RS = Nothing
    Set RS2 = Nothing
    DoEvents
    
    indentlevel = indentlevel - 1
    
    ' show progress on progress bar in terms of base table records (e.g., patients)
    'If TableName = sBaseTablename And prgBarDataRetirement.Value < prgBarDataRetirement.Max Then
    '    prgBarDataRetirement.Value = prgBarDataRetirement.Value + 1
    'End If

    DoEvents
    Exit Function

ErrHandler:
    'This is where the error handler should go
    Log "---------------------------------------------------------------------"
    Log "Table = " & TableName
    Log "Where = " & WhereClause
    Log "-----------------------"
    Log Err.Description
    Log "---------------------------------------------------------------------"
'    MsgBox "VB Error: " & Err.Number & " in CopyData. " & Err.Description, vbExclamation, "DataPump"
    
    DoEvents

End Function

Public Function RecordCount(TableName As String, WhereClause As String) As Integer

On Error GoTo ErrHandler
    
    Dim sSQL As String
    Dim RS As ADODB.Recordset
    
    RecordCount = 0
    
    'Make sure tablename is uppercase
    TableName = UCase(TableName)
    
    'get the data
    If oDAS Is Nothing Then
       Set oDAS = oDatabase.CHCSII_DAS(Auto)  'GEMS_DAS
    End If
    sSQL = "Select count(*) from " & TableName & " " & WhereClause
    Set RS = oDAS.OpenRecordset(sSQL)
    If RS.EOF = True Then
        'if there is no data to archive in this table there should be now in the tables that it uses
        Exit Function
    End If
    RecordCount = RS.Fields(0).Value
    
    Exit Function

ErrHandler:
    RecordCount = -1

End Function

Public Function GEMSARecordCount(TableName As String) As Integer

On Error GoTo ErrHandler
    
    Dim sSQL As String
    Dim RS As ADODB.Recordset
    
    GEMSARecordCount = 0
    
    'Make sure tablename is uppercase
    TableName = UCase(TableName)
    
    'get the data
    If oDAS Is Nothing Then
       Set oDAS = oDatabase.CHCSII_DAS(Auto)  'GEMS_DAS
    End If
    sSQL = "Select count(*) from GEMSA." & TableName
    Set RS = oDAS.OpenRecordset(sSQL)
    If RS.EOF = True Then
        'if there is no data to archive in this table there should be now in the tables that it uses
        Exit Function
    End If
    GEMSARecordCount = RS.Fields(0).Value
    
    Exit Function

ErrHandler:
    GEMSARecordCount = -1

End Function

Private Sub PrivPurgeData(TableName As String, Optional sWhere As String = "")

On Error GoTo ErrHandler
    
    Dim sSQL As String
    
    If oDAS Is Nothing Then
       Set oDAS = oDatabase.CHCSII_DAS(Auto)  'GEMS_DAS
    End If
    
    sSQL = "delete from " & TableName & " " & sWhere
    If bLogging Then
        Log sSQL
    End If
    oDAS.ExecuteSQL sSQL
    
    Exit Sub

ErrHandler:
    If bLogging Then
        Log "SQL error.  Could not delete. " & Err.Description
        Log "---------------------------------------------------------------------"
    End If
    
End Sub

Private Sub SetRetiredFlag(sWhere As String)
' SCR 87549; adding retired flag to encounters table (JQ)
On Error GoTo ErrHandler
    
    Dim sSQL As String
    
    If oDAS Is Nothing Then
       Set oDAS = oDatabase.CHCSII_DAS(Auto)  'GEMS_DAS
    End If
    
    sSQL = "update encounters set retired=1 " & sWhere
    If bLogging Then
        Log sSQL
    End If
    oDAS.ExecuteSQL sSQL
    
    Exit Sub

ErrHandler:
    If bLogging Then
        Log "SQL error.  Could not update retired flag. " & Err.Description
        Log "---------------------------------------------------------------------"
    End If
    
End Sub

Public Sub PurgeGEMSAData()

On Error GoTo ErrHandler
    
    Dim sSQL As String
    Dim RS As ADODB.Recordset
    
    If oDAS Is Nothing Then
       Set oDAS = oDatabase.CHCSII_DAS(Auto)  'GEMS_DAS
    End If
    
    If bLogging Then
        Log "---Purging GEMSA Data---"
    End If
    
    'get from database
'    sSQL = "select table_name from all_tables where owner = 'GEMSA'"
    sSQL = "select TABLE_NAME from RET_PURGELIST order by PURGE_ORDER "
    Set RS = oDAS.OpenRecordset(sSQL)
    If RS.EOF = True Then
        If bLogging Then
            Log "Purge list is empty."
        End If
        Exit Sub
    End If
    While RS.EOF = False
        PrivPurgeData "GEMSA." & RS.Fields(0).Value
        If prgBarDataRetirement.Value < prgBarDataRetirement.Max Then
            prgBarDataRetirement.Value = prgBarDataRetirement.Value + 1
        End If
        DoEvents
        RS.MoveNext
    Wend
    
    Exit Sub
    
ErrHandler:
    '
End Sub

Public Sub LogTableSummary()
' for Testing feedback (JQ)

On Error GoTo ErrHandler
    
    Dim sSQL As String
    Dim RS As ADODB.Recordset
    Dim iGEMScount As Integer
    Dim iGEMSAcount As Integer
    
    ' only log table summary if we're currently logging
    If Not bLogging Then
        Exit Sub
    End If
    
    If oDAS Is Nothing Then
       Set oDAS = oDatabase.CHCSII_DAS(Auto)  'GEMS_DAS
    End If
    
    If bLogging Then
        Log "---Table Summary:---"
    End If
    
    sSQL = "select TABLE_NAME from RET_PURGELIST order by TABLE_NAME "
    Set RS = oDAS.OpenRecordset(sSQL)
    If RS.EOF = True Then
        If bLogging Then
            Log "Oops. Purge list is empty. I can't build a table list."
        End If
        Exit Sub
    End If
    
    Log "Table" & Space(30) & vbTab & "GEMS Rcds" & vbTab & "GEMSA Rcds"
    Log "-----" & Space(30) & vbTab & "---------" & vbTab & "----------"
    While RS.EOF = False
        iGEMScount = RecordCount(RS.Fields(0).Value, "")
        iGEMSAcount = GEMSARecordCount(RS.Fields(0).Value)
        
        Log RS.Fields(0).Value & Space(35 - Len(RS.Fields(0).Value)) & vbTab _
            & iGEMScount & vbTab & vbTab & iGEMSAcount
        
        DoEvents
        RS.MoveNext
    Wend
    
    Log "--------------------"
    
    Exit Sub
    
ErrHandler:
    '
End Sub
Public Function PurgeGEMSData() As Boolean

On Error GoTo ErrHandler
    
    Dim sSQL As String
    Dim RS As ADODB.Recordset
    Dim RS2 As ADODB.Recordset
    Dim sTemp As String
    Dim sFields As String
    Dim sWhere As String
    Dim iCounter As Integer
    
    Dim oCmd As ADODB.command
    Dim oParam As ADODB.Parameter
    Dim oGEMSSQL As ICHCSII_SQL.ISqlOpsEx
    Dim oDbLDDB As New CHCSII_CONN.Conn
    Dim bGotsome As Boolean
    
    PurgeGEMSData = False
    
    If oDAS Is Nothing Then
       Set oDAS = oDatabase.CHCSII_DAS(Auto)  'GEMS_DAS
    End If
    
    If bLogging Then
        Log "---Purging GEMS Data---"
    End If
    
    'get list of tables from database
    ' SCR 87549: added ENCOUNTERS even if not being purged, s.t. we can toggle the retirement flag (JQ)
    sSQL = "select TABLE_NAME, PRIMARY_KEYS, PURGE_ORDER from RET_PURGELIST where PURGE_ORDER > -1 " _
        & " or table_name = 'ENCOUNTERS' order by PURGE_ORDER "
    Set RS = oDAS.OpenRecordset(sSQL)
    If RS.EOF = True Then
        'if there is no data to archive in this table there should be now in the tables that it uses
        If bLogging Then
            Log "Purge list is empty."
        End If
        Exit Function
    End If
    While RS.EOF = False
        sTemp = RS.Fields!TABLE_NAME
        sFields = ""
        bGotsome = False
        If IsNull(RS.Fields!PRIMARY_KEYS) Then
            sSQL = "select column_name from ALL_CONS_COLUMNS " & _
               " where owner = 'GEMS' " & _
               " and CONSTRAINT_NAME = '" & GetConstraintName(sTemp) & "' order by position"
            Set RS2 = oDAS.OpenRecordset(sSQL)

            While RS2.EOF = False
                If sFields = "" Then
                    sFields = RS2.Fields!COLUMN_NAME
                Else
                    sFields = sFields & ", " & RS2.Fields!COLUMN_NAME
                End If
                RS2.MoveNext
            Wend
            RS2.Close
        Else
            sFields = RS.Fields!PRIMARY_KEYS
        End If
        sSQL = "select " & sFields & " from GEMSA." & sTemp
        Set RS2 = oDAS.OpenRecordset(sSQL)

        If RS2.EOF = False Then  ' do we have data to purge? if so, we'll reindex later
            bGotsome = True
        End If
        
        While RS2.EOF = False
            For iCounter = 0 To RS2.Fields.Count - 1
                If iCounter = 0 Then
                    sWhere = "where " & RS2.Fields(iCounter).Name & " = "
                Else
                    sWhere = sWhere & " and " & RS2.Fields(iCounter).Name & " = "
                End If
                
                'If RS2.Fields(iCounter).Type = 131 Then
                '    sWhere = sWhere & RS2.Fields(iCounter).Value
                'Else
                '    sWhere = sWhere & "'" & RS2.Fields(iCounter).Value & "'"
                'End If
                ' SCR 85863 (JQ) adding provision for dates also
                
                If RS2.Fields(iCounter).Type = adNumeric _
                  Or RS2.Fields(iCounter).Type = adVarNumeric Then
                    sWhere = sWhere & RS2.Fields(iCounter).Value
                ElseIf RS2.Fields(iCounter).Type = adDate _
                  Or RS2.Fields(iCounter).Type = adDBDate _
                  Or RS2.Fields(iCounter).Type = adDBTime _
                  Or RS2.Fields(iCounter).Type = adDBTimeStamp Then
                    ' does this date have a timestamp also?
                    ' possible formats: '1/1/2006'
                    ' or: '10/20/2005 2:37:06 AM'
                    If InStr(1, RS2.Fields(iCounter).Value, ":", 1) = 0 Then
                        ' no time substring
                        sWhere = sWhere & oDbLDDB.SQLDate(RS2.Fields(iCounter).Value, DD_MON_YYYY)
                    Else
                        ' includes time
                        sWhere = sWhere & oDbLDDB.SQLDate(RS2.Fields(iCounter).Value, DD_MON_YYYY_HH_MM_SS)
                    End If
                Else
                    sWhere = sWhere & "'" & RS2.Fields(iCounter).Value & "'"
                End If
                
            Next
            
            If RS.Fields!PURGE_ORDER < 0 And RS.Fields!TABLE_NAME = "ENCOUNTERS" Then
                ' SCR 87549, JQ
                ' toggle retired flag in GEMS
                SetRetiredFlag sWhere
            Else
                ' remove record from GEMS
                PrivPurgeData sTemp, sWhere
            End If
            
            DoEvents
            RS2.MoveNext
        Wend
        If prgBarDataRetirement.Value < prgBarDataRetirement.Max Then
            prgBarDataRetirement.Value = prgBarDataRetirement.Value + 1
        End If
        DoEvents
        
        ' if data was deleted, refresh indices for this table, SCR 84888, JQ
        If bGotsome Then
            sSQL = "execute UPDATE_STATS(" & sTemp & ")" ' for logging only
            Set oCmd = New ADODB.command
            Set oGEMSSQL = oDbLDDB.CHCSII_SQLOPS(Auto)
            With oCmd
                .CommandText = "UPDATE_STATS"
                .CommandType = adCmdStoredProc
                .Parameters.Append .CreateParameter("TableName", adVarChar, adParamInput, 128, sTemp)
            End With

            Call oGEMSSQL.ExecuteCmd(oCmd)
        
            If bLogging Then
                Log sSQL
            End If
        End If
        
        RS.MoveNext
    Wend
    
    PurgeGEMSData = True
    
    Set oGEMSSQL = Nothing
    
    Exit Function
    
ErrHandler:
    MsgBox "VB Error: " & Err.Number & " in CopyData. " & Err.Description, vbExclamation, "DataPump"

End Function

Public Function UpdatePurgeOrder(TableName As String) As Boolean

On Error GoTo ErrHandler
    
    Dim sSQL As String
    Dim sSQL2 As String
    Dim sConstraintName As String
    Dim RS As ADODB.Recordset
    Dim RS2 As ADODB.Recordset
    Dim Temp As String
    Dim Temp2 As String
    
    Dim RSTableRefData As ADODB.Recordset
    
    indentlevel = indentlevel + 1
    
    'Make sure tablename is uppercase
    TableName = UCase(TableName)
    
    'Get the constraint name
    sConstraintName = GetConstraintName(TableName)
    
    If sConstraintName = "" Then
        GoTo EndFunct
    End If
    
    If RSTableRefData Is Nothing Then
        Set RSTableRefData = New ADODB.Recordset
        RSTableRefData.CursorLocation = adUseClient
        RSTableRefData.Fields.Append "CONSTRAINT_NAME", adBSTR
        RSTableRefData.Fields.Append "TABLE_NAME", adBSTR
        RSTableRefData.Fields.Append "R_CONSTRAINT_NAME", adBSTR
        RSTableRefData.Fields.Append "CONSTRAINT_TYPE", adBSTR
        RSTableRefData.Open
    End If
    
    GetRefByTables RSTableRefData, sConstraintName
    
    'get the data
    RSTableRefData.Filter = "R_CONSTRAINT_NAME = '" & sConstraintName & "'"
            
    While RSTableRefData.EOF = False
        Temp = RSTableRefData.Fields!TABLE_NAME
        If Temp = "ENCOUNTERS" Then
        Else
            UpdatePurgeOrder Temp
        End If
        RSTableRefData.MoveNext
    Wend
    sConstraintName = GetConstraintName(TableName)
    If sConstraintName = "" Then
        Log "insert into ret_purgelist (table_name, purge_order, primary_keys) " & _
            "VALUES ('" & TableName & "'," & iPurgeOrder & ",'NEED A PRIMARY KEY');"
    Else
        Log "insert into ret_purgelist (table_name, purge_order, primary_keys) " & _
            "VALUES ('" & TableName & "'," & iPurgeOrder & ", NULL);"
    End If
    iPurgeOrder = iPurgeOrder + 1

EndFunct:
    DoEvents
    Exit Function

ErrHandler:
    Log Err.Description
    DoEvents
End Function
