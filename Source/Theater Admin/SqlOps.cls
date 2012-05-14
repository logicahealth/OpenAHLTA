VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SqlOps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements ICHCSII_SQL.ISqlOps
'Implements ISec_SQL.ISecSqlOps
Implements ICHCSII_SQL.ISqlOpsEx
'Implements ICHCSII_SQL.ISQLDirect

Private Const SQLQueryNCID As String = "1450582"
Private Const OraErrorNoDataFound As String = "ORA-00100"
Private Const OraErrorFetchOutOfSequence As String = "ORA-01002"


'Private moCnn As ADODB.Connection
Private mvarRecordSet As ADODB.Recordset
Public UserNCID As String
Private mvarSQLStatement As String
Private mNoAudit As Collection
#If DEBUGON Then
  Public DL                             As cDebugLog
  Private Const Exe                      As String = "Gems_sql"
  Private Const Module                  As String = "Gems_SQL Cls"
#End If



Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Type TIME_ZONE_INFORMATION
        Bias As Long
        StandardName(32) As Integer
        StandardDate As SYSTEMTIME
        StandardBias As Long
        DaylightName(32) As Integer
        DaylightDate As SYSTEMTIME
        DaylightBias As Long
End Type

Private lTimeBias As Long
Private TZInfo As TIME_ZONE_INFORMATION

Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

Private Function ISqlOps_ExecuteSP(Parms As CHCSII_Command.CHCSII_Cmd) As Variant

End Function

Private Function ISqlOps_ValiadeFacilityAgainstCDR(FacilityNCID As String, ByVal EnterpriseNCID As String, NCIDIsRepresentation As Boolean, ByVal CreateIfNotFound As Boolean) As Boolean

End Function

Private Function ISqlOps_ValidateClinicAgainstCDR(ClinicNCID As String, ByVal FacilityNCID As String, FacilityIsValid As Boolean, NCIDIsRepresentation As Boolean, ByVal EnterpriseNCID As String, ByVal CreateIfNotFound As Boolean) As Boolean

End Function

'Private Function ISqlOps_ValidateProviderAgainstCDR(ProviderNCID As String, ByVal ProviderSSN As String, Optional ByVal ProviderNAME As String) As ICHCSII_SQL.eMatch
'
'End Function

Private Function ISqlOpsEx_AppendAsChunk(ByVal vColumn As Variant, vText As Variant) As Boolean
'- for the column indicated, append the byte array as chunks
    Dim lSize       As Long
    Dim loffset     As Long
    Dim varChunk    As Variant
    Dim ChunkSize   As Long
    Dim index       As Integer
'    Dim oZLib       As ZLib
    Dim lResult     As Long
    Dim sTxt As String
    
    If IsNumeric(vColumn) Then
        index = vColumn
    Else
        index = ISqlOps_ColumnIndex(CStr(vColumn))
    End If
    
    
    lSize = LenB(vText)
'- should we compress it?
'    If lsize > 4000 Then
'    '-- yes
'        Set oZLib = New ZLib
'        sTxt = vText
'
'        lResult = oZLib.CompressString(sTxt)
'        Set oZLib = Nothing
'    '-- label it as such
'        vText = Format$(lsize, "ZLIB000000000") & sTxt
'    End If

    '-- pick a chunk size
    ChunkSize = 2000
    If LenB(vText) > 10000 Then
        ChunkSize = LenB(vText) \ 10
        '- Make it EVEN (2-bytes per char, remember?)
        If ChunkSize \ 2 <> ChunkSize / 2 Then ChunkSize = ChunkSize + 1
        If ChunkSize > 20000 Then ChunkSize = 20000
    End If

    
    lSize = LenB(vText)
    loffset = 0 ' Reset offset
    
    Do While loffset < lSize
       varChunk = LeftB(RightB(vText, lSize - loffset), ChunkSize)
        mvarRecordSet.Fields(index).AppendChunk varChunk
       loffset = loffset + ChunkSize
    Loop
    
    
    ISqlOpsEx_AppendAsChunk = True
    Exit Function
errhandler:

End Function

Private Sub ISqlOpsEx_BeginTrans():    moCnn.BeginTrans:    End Sub
Private Sub ISqlOpsEx_CommitTrans():   moCnn.CommitTrans:   End Sub

Private Function ISqlOpsEx_GetAsChunk(ByVal vColumn As Variant, vText As Variant) As Boolean
'- for the column indicated, get the data as chunks and return in the byte array
'- for the column indicated, append the byte array as chunks
    Dim lSize       As Long
    Dim loffset     As Long
    Dim varChunk    As Variant
    Dim index       As Integer
    Dim ChunkSize   As Long
'    Dim oZLib       As ZLib
    Dim lResult     As Long
    Dim sTxt        As String
    
    If IsNumeric(vColumn) Then
        index = vColumn
    Else
        index = ISqlOps_ColumnIndex(CStr(vColumn))
    End If
    
       
   
    lSize = mvarRecordSet.Fields(index).ActualSize
    loffset = 0 ' Reset offset
    vText = ""
    
    '-- set chunksize
    ChunkSize = 2000
    If lSize > 10000 Then
        ChunkSize = lSize \ 10
        '- Make it EVEN (2-bytes per char, remember?)
        If ChunkSize \ 2 <> ChunkSize / 2 Then ChunkSize = ChunkSize + 1
        If ChunkSize > 20000 Then ChunkSize = 20000
    End If
    
    Do While loffset < lSize
       varChunk = mvarRecordSet.Fields(index).GetChunk(ChunkSize)
       vText = vText & varChunk
       loffset = loffset + ChunkSize
    Loop
        
    '-- see if we compressed it
'    If Left(vText, 4) = "ZLIB" Then
'        lsize = Val(Mid(vText, 5, 9))
'        vText = Mid(vText, 14)
'        sTxt = vText
'        Set oZLib = New ZLib
'        lResult = oZLib.DecompressString(sTxt, lsize)
'        Set oZLib = Nothing
'    End If
            

End Function

Private Sub ISqlOpsEx_RollBackTrans(): moCnn.RollbackTrans: End Sub


Private Function ISecSqlOps_Connect() As Boolean
'    On Error Resume Next
'    moCnn.Close
'    'Set moCnn = DESec.Cnn
'    moCnn.ConnectionString = GEMS_Sec_Conn
'    moCnn.Open
'    If moCnn.State = 1 Then
'        Call SetDBDateFormat(moCnn)
'    End If
'
'  Set mvarRecordSet = New ADODB.Recordset

End Function


Private Function ISqlOps_ExecuteCmd(oAdoCmd As ADODB.Command) As Variant
    Dim lRecs As Long     '- number of records affected
    
    oAdoCmd.ActiveConnection = moCnn
    
    Select Case oAdoCmd.CommandType
    Case adCmdStoredProc
        Set ISqlOps_ExecuteCmd = oAdoCmd.Execute
        
    Case adCmdText, adCmdTable
        Select Case UCase$(Left$(oAdoCmd.CommandText, 1))
        Case "S"   '- it's a select stmt
            Set ISqlOps_ExecuteCmd = oAdoCmd.Execute
        Case Else  '- it's an update, insert, or delete stmt
            Call oAdoCmd.Execute(lRecs)
            ISqlOps_ExecuteCmd = lRecs
        End Select
        
    Case Else
        Set ISqlOps_ExecuteCmd = oAdoCmd.Execute
        
    End Select
End Function


Private Function ISqlOps_GetNextID(Optional ByVal eSeqName As ICHCSII_SQL.eSeq = esMAINSEQ) As Variant
'-- pulls the next value from the MainSeq sequence
    On Error GoTo errhandler
    Dim oRs As ADODB.Recordset
    Set oRs = New ADODB.Recordset
    Dim sSeqName As String
    sSeqName = SeqEnumToString(eSeqName)
    oRs.Open "Select " & sSeqName & ".nextval from dual", moCnn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not oRs.EOF Then
        ISqlOps_GetNextID = oRs(0)
    End If
    
errhandler:
    If oRs.State > 0 Then oRs.Close
    Set oRs = Nothing
    
End Function

Private Function SeqEnumToString(ByVal eiSeq As eSeq) As String
    Select Case eiSeq
    Case esMAINSEQ:         SeqEnumToString = "MainSeq"
    Case esAPPT_ID:         SeqEnumToString = "APPT_ID"
    Case esAPPT_TRANS_ID:   SeqEnumToString = "APPT_TRANS_ID"
    Case esENC_TEMPLATE_ID: SeqEnumToString = "ENC_TEMPLATE_ID"
    Case esENCOUNTER_ID:    SeqEnumToString = "ENCOUNTER_ID"
    Case esIEN:             SeqEnumToString = "IEN"
    Case esPCM_CODE:        SeqEnumToString = "PCM_CODE"
    Case esPI_DATA_ID:      SeqEnumToString = "PI_DATA_ID"
    Case esUNIT_NUMBER:     SeqEnumToString = "UNIT_NUMBER"
    Case esUSER_NUMBER:     SeqEnumToString = "USER_NUMBER"
    Case esJDOE_NUMBER:     SeqEnumToString = "JDOE_NUMBER"
    Case esTEMP_SSN:        SeqEnumToString = "TEMP_SSN"
    End Select

End Function

Private Function ISqlOps_GetNextNcid() As Variant
'-- pulls the next value from the NCID sequence
    Dim oRs As ADODB.Recordset
    Set oRs = New ADODB.Recordset
    
    oRs.Open "Select NCID.nextval from dual", moCnn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not oRs.EOF Then
        ISqlOps_GetNextNcid = oRs(0)
    End If
    oRs.Close
    Set oRs = Nothing

End Function


Private Function ISqlOps_ValidateFacility(ByVal FacilityNCID As Long, ByVal CreateIfNotFound As Boolean, Optional ByVal ClinicNCID As Long = 0&) As Boolean
    '-- see if the facility is in the facility table
    Dim bNeedFacility As Boolean
    Dim bNeedClinic As Boolean
    Dim bSuccess As Boolean
    
    On Error GoTo errhandler
'- see what's out there

    bNeedFacility = Not ValidateFacility(FacilityNCID)
    bNeedClinic = Not ValidateClinic(FacilityNCID, ClinicNCID)

    If Not bNeedFacility And Not bNeedClinic Then
        ISqlOps_ValidateFacility = True
        Exit Function
    End If
    
    If Not CreateIfNotFound Then
        ISqlOps_ValidateFacility = Not bNeedFacility And Not bNeedClinic
        Exit Function
    End If
    
'- so we should create what we don't have

    If bNeedFacility Then
        bSuccess = BackfillFacility(FacilityNCID)
    End If
    If bNeedClinic Then
        bSuccess = bSuccess And BackfillClinic(FacilityNCID, ClinicNCID)
    End If
    ISqlOps_ValidateFacility = bSuccess

errhandler:

End Function

Private Function ValidateFacility(ByVal FacilityNCID As Long) As Boolean
    Dim oRs As ADODB.Recordset
    Set oRs = New ADODB.Recordset
    
    oRs.Open "Select facility_ncid from facility_ncid_correlation where facility_ncid = " & FacilityNCID, moCnn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not oRs.EOF Then
        ValidateFacility = True
    Else
        ValidateFacility = False
    End If
    oRs.Close
    Set oRs = Nothing

End Function
Private Function ValidateClinic(ByVal FacilityNCID As Long, ByVal ClinicNCID As Long) As Boolean
    Dim oRs As ADODB.Recordset
    Set oRs = New ADODB.Recordset
    
    If ClinicNCID = 0 Then
        ValidateClinic = True
        Exit Function
    End If
    
    oRs.Open "Select ncid from clinic where facility_ncid = " & FacilityNCID _
           & " and ncid = " & ClinicNCID, moCnn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not oRs.EOF Then
        ValidateClinic = True
    Else
        ValidateClinic = False
    End If
    
    oRs.Close
    Set oRs = Nothing

End Function
Private Function BackfillFacility(ByVal FacilityNCID As Long) As Boolean
    If FacilityNCID = 0 Then Exit Function
    On Error GoTo errhandler
    
    moCnn.Execute "insert into facility_ncid_correlation (facility_ncid, context_ncid, type) values (" _
                & FacilityNCID & ",119,'ADT')"
    moCnn.Execute "insert into facility_ncid_correlation (facility_ncid, context_ncid, type) values (" _
                & FacilityNCID & ",33962,'ADT')"
    
    moCnn.Execute "insert into ncids (ncid, domain, description) values (" _
                & FacilityNCID & ",119,'Added by CHCSII-T to support foreign data')"
    moCnn.Execute "insert into ncids (ncid, domain, description) values (" _
                & FacilityNCID & ",33962,'Added by CHCSII-T to support foreign data')"
    BackfillFacility = True
errhandler:
End Function
Private Function BackfillClinic(ByVal FacilityNCID As Long, ByVal ClinicNCID As Long) As Boolean
    If ClinicNCID = 0 Then Exit Function
    
    Dim meprs_code As String
    Dim meprs_ien As String
    Const sName As String = "CHCSII-T AutoAdd"
    On Error GoTo errhandler
    meprs_code = ISqlOps_GetNextID()
    meprs_ien = ClinicNCID
    
    moCnn.Execute "insert into clinic (ncid, facility_ncid, location_type,meprs_code, meprs_ien, name) values (" _
                & ClinicNCID & "," & FacilityNCID & ",'CLINIC'," & meprs_code & "," & meprs_ien & "," & SQLQuote(sName) & ")"
    
    moCnn.Execute "insert into ncids (ncid, domain, description) values (" _
                & ClinicNCID & ",1003," & SQLQuote(sName) & ")"

    moCnn.Execute "insert into meprs_code(facility_ncid, ien, code, description) values (" _
                & FacilityNCID & "," & meprs_ien & "," & meprs_code & "," & SQLQuote(sName) & ")"
    BackfillClinic = True
errhandler:
End Function
Public Property Get ISqlOps_Value(Column As Variant) As Variant
Attribute ISqlOps_Value.VB_UserMemId = 0
Attribute ISqlOps_Value.VB_MemberFlags = "200"

    Dim index As Integer
    On Error GoTo errhandler
    If ISqlOps_BOF Or ISqlOps_EOF Then
        GoTo errhandler
    End If
    
    If IsNumeric(Column) Then
        index = Column
    Else
        index = ISqlOps_ColumnIndex(CStr(Column))
    End If
    
    ISqlOps_Value = mvarRecordSet.Fields(index).value
'    If mvarRecordSet.Fields(index).Type = adDate Or mvarRecordSet.Fields(index).Type = adDBDate Or mvarRecordSet.Fields(index).Type = adDBTime Or mvarRecordSet.Fields(index).Type = adDBTimeStamp Then
'      If ISqlOps_Value <> #12/31/9999# Then
'          ' adjust to local from gmt
'          ISqlOps_Value = DateAdd("n", -lTimeBias, ISqlOps_Value)
'      End If
'    End If
    If IsNull(ISqlOps_Value) Then
        Select Case mvarRecordSet.Fields(index).Type
            Case adBoolean
                ISqlOps_Value = False
            Case adNumeric, adDecimal, adDouble, adSmallInt, adTinyInt, adCurrency, adSingle, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                ISqlOps_Value = 0
            Case Else
                ISqlOps_Value = ""
        End Select
    End If  ' IsNull(ISqlOps_Value)
errhandler:
End Property

Public Property Get ISqlOps_InfoType(Column As Variant) As String

    Dim index As Integer
    If (ISqlOps_BOF Or ISqlOps_EOF) Then
        GoTo errhandler
    End If
    
    If IsNumeric(Column) Then
        index = Column
    Else
        index = ISqlOps_ColumnIndex(CStr(Column))
    End If
    
    ISqlOps_InfoType = mvarRecordSet.Fields(index).Type  ' mvarDSQLConnect.Table(mvarAbsolutePosition, index).Value.InfoType

errhandler:

End Property

Public Property Get ISqlOps_ColumnIndex(Name As String) As Integer

    Dim i As Integer
    Dim sText As String
    For i = 0 To ISqlOps_ColumnCount - 1
        If UCase(ISqlOps_ColumnName(i)) = UCase(Name) Then
            ISqlOps_ColumnIndex = i
            Exit Property
        End If
    Next i
    
    sText = "There is no column named '" & Name & "'.  Please select from:"
    For i = 0 To ISqlOps_ColumnCount - 1
        sText = sText & Chr(10) & mvarRecordSet.Fields(i).Name
    Next
    Err.Raise vbObjectError + 1, "CHCSII_SQL:SqlOps.ColumnIndex", sText
End Property

Public Property Get ISqlOps_ColumnName(index As Integer) As String

  
  ISqlOps_ColumnName = ""
    If Not (ISqlOps_BOF Or ISqlOps_EOF) Then
        If index > ISqlOps_ColumnCount Then
            Err.Raise vbObjectError + 1, "CHCSII_SQL:SqlOps.ColumnName", "Index (" & index & ") out of range.  There are " & ISqlOps_ColumnCount & " colummns."
            Exit Property
        End If
        ISqlOps_ColumnName = mvarRecordSet.Fields(index).Name
    End If
End Property

Public Property Get ISqlOps_ColumnCount() As Integer

'    ColumnCount = mvarDSQLConnect.Table.Rows(mvarAbsolutePosition).Columns.Count
  ISqlOps_ColumnCount = mvarRecordSet.Fields.count
End Property

Public Function ISqlOps_Execute(SqlStatement As String) As Variant
    ISqlOps_Execute = Execute(SqlStatement)
End Function
Public Function Execute(SqlStatement As String, Optional ByVal direct As Boolean = False) As Variant
        
    Dim lRecCnt     As Long
    Dim FirstInsPt  As Long
    Dim SecondInsPt As Long
    Dim ValuesStrPt As Long                 '*-- Coleman 040401
    Dim words()     As String
    Dim WhereClause As String
    
#If DEBUGON Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Execute", "Begin", DR, "SQL", SqlStatement
#End If
    
    '- ensure there are always spaces between parens and words
    SqlStatement = Replace(SqlStatement, "(", " (")
    SqlStatement = Replace(SqlStatement, ")", ") ")
    SqlStatement = Replace(SqlStatement, "  ", " ")  '- get rid of double spaces that resulted from above stmts
    words = Split(SqlStatement, " ")
    mvarSQLStatement = SqlStatement

  If mvarRecordSet Is Nothing Then
    Set mvarRecordSet = New ADODB.Recordset
  Else
    If mvarRecordSet.State = adStateExecuting Then
      mvarRecordSet.Cancel
    End If
    
    If mvarRecordSet.State = adStateOpen Then
        On Error Resume Next
'        Debug.Print mvarRecordSet.EditMode
        mvarRecordSet.CancelUpdate
        mvarRecordSet.Close
        On Error GoTo 0
    End If
    
  End If
  
  If moCnn Is Nothing Then
     Call OpenDB
  End If
  
  If moCnn.State = 0 Then
    moCnn.Open
  End If
  
  If UBound(words) > 1 Then
    Select Case UCase$(Left$(SqlStatement, 1))
    
    Case "D"
        '- is as "Delete from tablename where field...
        If Not direct And Not DoNoAuditing(words(2)) Then
            If InStr(LCase$(SqlStatement), " where ") Then
                WhereClause = Mid$(SqlStatement, InStr(LCase$(SqlStatement), " where "))
            Else
                WhereClause = ""
            End If
            moCnn.Execute "update " & words(2) & " set UpdatedBy = " & SQLQuote(UserNCID) _
                        & ", UpdatedOn = " & OracleDate(Now()) & WhereClause, lRecCnt
        End If
        moCnn.Execute SqlStatement, lRecCnt
        Execute = lRecCnt
        
    Case "I"    '- may be "Insert into Tablename ( field1....) values (value1...)"
                '- or     "Insert into TableName (field1...) select ....."
                
        If InStr(words(2), "(") > 1 Then words(2) = Left$(words(2), InStr(words(2), "(") - 1)
        
        If Not direct And Not DoNoAuditing(words(2)) Then
            FirstInsPt = InStr(SqlStatement, "(")
            '*-- Coleman 040401: Handles case where no space before parentheses - "values("
            '*-- If InStr(LCase(SqlStatement), " values ") > 0 Then
            '*--    SecondInsPt = InStr(FirstInsPt + 1, SqlStatement, "(")
            ValuesStrPt = InStr(LCase(SqlStatement), " values")         '*-- Coleman 040401
            If ValuesStrPt > 0 Then                                     '*-- Coleman 040401
                SecondInsPt = InStr(ValuesStrPt, SqlStatement, "(")     '*-- Coleman 040401
            Else
                SecondInsPt = InStr(LCase(SqlStatement), " select ") + 7
            End If
       
            SqlStatement = Stuff(SqlStatement, SecondInsPt, SQLQuote(UserNCID) & ", " & OracleDate(Now()) & ", ")
            SqlStatement = Stuff(SqlStatement, FirstInsPt, "CreatedBy, CreatedOn, ")
        End If
        
        moCnn.Execute SqlStatement, lRecCnt
        Execute = lRecCnt
        
    Case "U"    '- must be "Update TableName Set Field1 = ..."
    '- get the table name
       
        If Not direct And Not DoNoAuditing(words(1)) Then
            FirstInsPt = InStr(UCase$(SqlStatement), "SET")
            SqlStatement = Stuff(SqlStatement, FirstInsPt + 3, " UpdatedBy = " & SQLQuote(UserNCID) & ", UpdatedOn = " & OracleDate(Now()) & ", ")
        End If
        moCnn.Execute SqlStatement, lRecCnt
        Execute = lRecCnt
        
    Case "S"
       mvarRecordSet.Open SqlStatement, moCnn, adOpenDynamic, adLockPessimistic, adCmdText
       'Set ISqlOps_Execute = mvarRecordSet
        
    End Select
  Else  '- ubound(words) > 1
     mvarRecordSet.Open SqlStatement, moCnn, adOpenKeyset, adLockOptimistic, adCmdTable
  End If
  
#If DEBUGON Then
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Execute", "End", DR
#End If
    
  
End Function
Public Sub ISqlOps_ExecuteMulti(SqlStatements As Collection)
    Call ExecuteMulti(SqlStatements)
End Sub

Public Sub ExecuteMulti(SqlStatements As Collection, Optional ByVal direct As Boolean = False)
  Dim vStmt As Variant
  On Error GoTo errhandler
  
  moCnn.BeginTrans
  
  For Each vStmt In SqlStatements
    mvarSQLStatement = vStmt
    Execute CStr(vStmt), direct
  Next vStmt
      
  moCnn.CommitTrans
  Exit Sub
  
errhandler:
  moCnn.RollbackTrans
  
End Sub


Public Sub ISqlOps_MovePrevious()
    If mvarRecordSet.State = 0 Or mvarRecordSet.BOF Then Exit Sub
    mvarRecordSet.MovePrevious
End Sub

Public Sub ISqlOps_MoveNext()
    If mvarRecordSet.State = 0 Or mvarRecordSet.EOF Then Exit Sub
    mvarRecordSet.MoveNext
End Sub

Public Sub ISqlOps_MoveLast()
    If mvarRecordSet.State = 0 Then Exit Sub
    mvarRecordSet.MoveLast
    
End Sub

Public Sub ISqlOps_MoveFirst()
    If mvarRecordSet.State = 0 Then Exit Sub
    mvarRecordSet.MoveFirst

End Sub

Public Property Get ISqlOps_RecordCount() As Long
    On Error Resume Next
    ISqlOps_RecordCount = mvarRecordSet.RecordCount
End Property

Public Property Get ISqlOps_Source() As String:  ISqlOps_Source = mvarSQLStatement: End Property
Public Property Get ISqlOps_EOF() As Boolean:    ISqlOps_EOF = mvarRecordSet.EOF:   End Property
Public Property Get ISqlOps_BOF() As Boolean:    ISqlOps_BOF = mvarRecordSet.BOF:   End Property

Private Sub Class_Initialize()
#If DEBUGON Then
  Set DL = GetObject("", "dlog.cdebuglog")
  If Not DL Is Nothing Then DL.debugStart
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Class_Initialize", "", DR
#End If
    
    GetTimeZoneInformation TZInfo
    lTimeBias = TZInfo.Bias

    Dim oShared As New CWShared
    UserNCID = oShared.CurrentUserNCID

    mlCnnCnt = mlCnnCnt + 1
    If moCnn Is Nothing Then
        Set moCnn = New ADODB.Connection
        moCnn.ConnectionString = GEMS_Conn
        moCnn.ConnectionTimeout = 30
        moCnn.CursorLocation = adUseClient
        If moCnn.State = 0 Then moCnn.Open
        If moCnn.State = 1 Then
            Call SetDBDateFormat(moCnn)
        End If
    End If

  Set mvarRecordSet = New ADODB.Recordset

End Sub

Private Sub Class_Terminate()
    On Error Resume Next
'- handle the recordset
    If mvarRecordSet.State > 0 Then
        If Not (mvarRecordSet.EOF Or mvarRecordSet.BOF) Then
            mvarRecordSet.CancelUpdate      '- this will error if not in an update, but we're closing...
        End If
        mvarRecordSet.Close
    End If
    Set mvarRecordSet = Nothing
    
'- handle the connection
    mlCnnCnt = -1
    If mlCnnCnt <= 0 Then
        If moCnn.State > 0 Then moCnn.Close
        Set moCnn = Nothing
    End If
#If DEBUGON Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Class_Terminate", "", DR
#End If
    
End Sub

Private Function OpenDB()
  'Set moCnn = DE.Cnn

  Set moCnn = New ADODB.Connection
  moCnn.ConnectionString = GEMS_Conn
'  moCnn.ConnectionTimeout = 30
  moCnn.Open
  
End Function

Private Function DoNoAuditing(sTable As String) As Boolean
    Dim sMember As Variant
    
    If mNoAudit Is Nothing Then LoadNoAudit
    
    For Each sMember In mNoAudit
        If LCase(sMember) = LCase(sTable) Then
            DoNoAuditing = True
            Exit For
        End If
    Next sMember
End Function

Private Sub LoadNoAudit()
    Set mNoAudit = New Collection
    
    Dim oRs As ADODB.Recordset
    Dim oSQL As ICHCSII_SQL.ISqlOps
    Set oSQL = New GEMS_SQL.SqlOps
    
    oSQL.Execute "DoNotAudit"
    Do While Not oSQL.EOF
        mNoAudit.Add "" & oSQL(0)
        oSQL.MoveNext
    Loop
    
    Set oSQL = Nothing

End Sub


'<< These are pass-through functions , same as isqlopsex_
Private Property Get ISQLDirect_BOF() As Boolean: ISQLDirect_BOF = ISqlOps_BOF: End Property
Private Property Get ISQLDirect_ColumnCount() As Integer: ISQLDirect_ColumnCount = ISqlOps_ColumnCount: End Property
Private Property Get ISQLDirect_ColumnIndex(Name As String) As Integer: ISQLDirect_ColumnIndex = ISqlOps_ColumnIndex(Name): End Property
Private Property Get ISQLDirect_ColumnName(index As Integer) As String: ISQLDirect_ColumnName = ISqlOps_ColumnName(index): End Property
Private Property Get ISQLDirect_EOF() As Boolean: ISQLDirect_EOF = ISqlOps_EOF: End Property
Private Function ISQLDirect_GetNextID(Optional ByVal eSeqName As ICHCSII_SQL.edSeq = edMAINSEQ) As Variant: ISQLDirect_GetNextID = ISqlOps_GetNextID(eSeqName): End Function
Private Function ISQLDirect_GetNextNcid() As Variant: ISQLDirect_GetNextNcid = ISqlOps_GetNextNcid: End Function
Private Property Get ISQLDirect_InfoType(Column As Variant) As String: ISQLDirect_InfoType = ISqlOps_InfoType(Column): End Property
Private Sub ISQLDirect_MoveFirst(): ISqlOps_MoveFirst: End Sub
Private Sub ISQLDirect_MoveLast(): ISqlOps_MoveLast: End Sub
Private Sub ISQLDirect_MoveNext(): ISqlOps_MoveNext: End Sub
Private Sub ISQLDirect_MovePrevious(): ISqlOps_MovePrevious: End Sub
Private Property Get ISQLDirect_RecordCount() As Long: ISQLDirect_RecordCount = ISqlOps_RecordCount: End Property
Private Property Get ISQLDirect_Source() As String: ISQLDirect_Source = ISqlOps_Source: End Property
Private Function ISQLDirect_ExecuteCmd(oAdoCmd As ADODB.Command) As Variant: ISQLDirect_ExecuteCmd = ISqlOps_ExecuteCmd(oAdoCmd): End Function
Private Property Get ISQLDirect_Value(Column As Variant) As Variant: ISQLDirect_Value = ISqlOps_Value(Column): End Property
Private Function ISQLDirect_ValidateFacility(ByVal FacilityNCID As Long, ByVal CreateIfNotFound As Boolean, Optional ByVal ClinicNCID As Long = 0&) As Boolean
    ISQLDirect_ValidateFacility = ISqlOps_ValidateFacility(FacilityNCID, CreateIfNotFound, ClinicNCID)
End Function
'>> end of pass-through
Private Function ISQLDirect_Execute(SqlStatement As String) As Variant
    ISQLDirect_Execute = Execute(SqlStatement, True)
End Function
Private Sub ISQLDirect_ExecuteMulti(SqlStatements As Collection)
    ExecuteMulti SqlStatements, True
End Sub



'<< these are all 'pass-throughs' to the isqlops_
Private Property Get ISqlOpsEx_BOF() As Boolean: ISqlOpsEx_BOF = ISqlOps_BOF: End Property
Private Property Get ISqlOpsEx_ColumnCount() As Integer: ISqlOpsEx_ColumnCount = ISqlOps_ColumnCount: End Property
Private Property Get ISqlOpsEx_ColumnIndex(Name As String) As Integer: ISqlOpsEx_ColumnIndex = ISqlOps_ColumnIndex(Name): End Property
Private Property Get ISqlOpsEx_ColumnName(index As Integer) As String: ISqlOpsEx_ColumnName = ISqlOps_ColumnName(index): End Property
Private Property Get ISqlOpsEx_EOF() As Boolean: ISqlOpsEx_EOF = ISqlOps_EOF: End Property
Private Function ISqlOpsEx_Execute(SqlStatement As String) As Variant: ISqlOpsEx_Execute = ISqlOps_Execute(SqlStatement): End Function
Private Sub ISqlOpsEx_ExecuteMulti(SqlStatements As Collection): ExecuteMulti SqlStatements: End Sub
Private Function ISqlOpsEx_GetNextID(Optional ByVal eSeqName As ICHCSII_SQL.exSeq = exMAINSEQ) As Variant: ISqlOpsEx_GetNextID = ISqlOps_GetNextID(eSeqName): End Function
Private Function ISqlOpsEx_GetNextNcid() As Variant: ISqlOpsEx_GetNextNcid = ISqlOps_GetNextNcid: End Function
Private Property Get ISqlOpsEx_InfoType(Column As Variant) As String: ISqlOpsEx_InfoType = ISqlOps_InfoType(Column): End Property
Private Sub ISqlOpsEx_MoveFirst(): ISqlOps_MoveFirst: End Sub
Private Sub ISqlOpsEx_MoveLast(): ISqlOps_MoveLast: End Sub
Private Sub ISqlOpsEx_MoveNext(): ISqlOps_MoveNext: End Sub
Private Sub ISqlOpsEx_MovePrevious(): ISqlOps_MovePrevious: End Sub
Private Property Get ISqlOpsEx_RecordCount() As Long: ISqlOpsEx_RecordCount = ISqlOps_RecordCount: End Property
Private Property Get ISqlOpsEx_Source() As String: ISqlOpsEx_Source = ISqlOps_Source: End Property
Private Function ISqlOpsEx_ExecuteCmd(oAdoCmd As ADODB.Command) As Variant: ISqlOpsEx_ExecuteCmd = ISqlOps_ExecuteCmd(oAdoCmd): End Function

Private Function ISqlOpsEx_SQLDate(ByVal dDate As Date) As String

End Function

Private Function ISqlOpsEx_SQLQuote(ByVal sText As String) As String

End Function

Private Function ISqlOpsEx_ValidateFacility2(FacilityNCID As Variant, FacilityName As Variant, ByVal CreateIfNotFound As Boolean, Optional ClinicNCID As Variant, Optional clinicName As Variant) As Boolean

End Function

Private Function ISqlOpsEx_ValidateProvider(ByVal UserNCID As Long, ByVal CreateIfNotFound As Boolean, Optional ByVal FacilityNCID As Long = 0&, Optional ByVal ClinicNCID As Long = 0&) As Boolean

End Function

Private Function ISqlOpsEx_ValidateProvider2(ProviderNCID As String, ByVal ProviderSSN As String, ByVal ProviderFirstName As String, ByVal ProviderMiddleName As String, ByVal ProviderLastName As String, ByVal CreateIfNotFound As Boolean) As Boolean

End Function

Private Property Get ISqlOpsEx_Value(Column As Variant) As Variant: ISqlOpsEx_Value = ISqlOps_Value(Column): End Property
Private Function ISqlOpsEx_ValidateFacility(ByVal FacilityNCID As Long, ByVal CreateIfNotFound As Boolean, Optional ByVal ClinicNCID As Long = 0&) As Boolean
    ISqlOpsEx_ValidateFacility = ISqlOps_ValidateFacility(FacilityNCID, CreateIfNotFound, ClinicNCID)
End Function
'>> end of pass-throughs


Private Function ISqlOpsex_AddNew() As Boolean
   If mvarRecordSet Is Nothing Then Exit Function
   On Error GoTo errhandler
   
   mvarRecordSet.AddNew
   ISqlOpsex_AddNew = True

errhandler:
   
End Function

Private Function ISqlOpsEx_Update() As Boolean
    On Error GoTo errhandler
    mvarRecordSet.Update
    ISqlOpsEx_Update = True
errhandler:
End Function

Private Function ISqlOpsEx_Cancel() As Boolean
    On Error GoTo errhandler
    mvarRecordSet.Cancel
    ISqlOpsEx_Cancel = True
errhandler:
End Function

Private Property Let ISqlOpsEx_Value(Column As Variant, ByVal RHS As Variant)
    Dim index As Integer
    On Error GoTo errhandler
    
    If ISqlOps_BOF Or ISqlOps_EOF Then
        GoTo errhandler
    End If
        
'- get the column
    If IsNumeric(Column) Then
        index = Column
    Else
        index = ISqlOps_ColumnIndex(CStr(Column))
    End If
    
    If mvarRecordSet.Fields(index).Type = adDate Or mvarRecordSet.Fields(index).Type = adDBDate Or mvarRecordSet.Fields(index).Type = adDBTime Or mvarRecordSet.Fields(index).Type = adDBTimeStamp Then
'      If RHS <> #12/31/9999# Then
'          '- adjust from local to Zulu time by subtracting the time bias
'          RHS = AdjustDate(RHS)
'      End If
    End If

    mvarRecordSet.Fields(index).value = RHS

errhandler:

End Property

Private Function AdjustDate(ByVal dDate As Date) As Date
    Static lYear As Integer
    Static dStandard As Date
    Static dDaylight As Date
    Dim lDay As Integer
    Dim lTimeBias As Long
    If Year(dDate) <> lYear Then
        lYear = Year(dDate)
        dStandard = GetTZInfoDate(TZInfo.StandardDate, lYear)
        dDaylight = GetTZInfoDate(TZInfo.DaylightDate, lYear)
    End If
    If dStandard > dDaylight Then
        If dDate >= dDaylight And dDate < dStandard Then
            lTimeBias = TZInfo.Bias + TZInfo.DaylightBias
        Else
            lTimeBias = TZInfo.Bias + TZInfo.StandardBias
        End If
    Else
        If dDate >= dStandard And dDate < dDaylight Then
            lTimeBias = TZInfo.Bias + TZInfo.StandardBias
        Else
            lTimeBias = TZInfo.Bias + TZInfo.DaylightBias
        End If
    End If
    AdjustDate = DateAdd("n", lTimeBias, dDate)
End Function

Private Function GetTZInfoDate(TZDate As SYSTEMTIME, TZYear As Integer) As Date
    With TZDate
        If .wDay = 5 Then
            ' means last .wDayOfWeek in month
            GetTZInfoDate = DateSerial(TZYear, .wMonth + 1, 0) + TimeSerial(.wHour, .wMinute, .wSecond)
            Do While Weekday(GetTZInfoDate) <> .wDayOfWeek + 1
                GetTZInfoDate = DateAdd("d", -1, GetTZInfoDate)
            Loop
        Else
            ' means .wday'TH .wdayofweek in month
            GetTZInfoDate = DateSerial(TZYear, .wMonth, 1) + TimeSerial(.wHour, .wMinute, .wSecond)
            Do While Weekday(GetTZInfoDate) <> .wDayOfWeek + 1
                GetTZInfoDate = DateAdd("d", 1, GetTZInfoDate)
            Loop
            GetTZInfoDate = DateAdd("d", (.wDay - 1) * 7, GetTZInfoDate)
        End If
    End With
End Function
