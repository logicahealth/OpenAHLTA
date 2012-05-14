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
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private sConnString As String
Public moCnn                   As ADODB.Connection

Private Const SQLQueryNCID As String = "1450582"
Private Const OraErrorNoDataFound As String = "ORA-00100"
Private Const OraErrorFetchOutOfSequence As String = "ORA-01002"
Private mobjShared As CWShared

'Private moCnn As ADODB.Connection    ' let's keep it common/shared connection.
                                        'but, give a way to reset it with a new Connection string
Private mvarRecordSet As ADODB.Recordset
Public UserNCID As String
Private mvarSQLStatement As String
Private mNoAudit As Collection

Private mFacilityNCID As String  '- for backfill of providers
Private mClinicNCID   As String
Private msDBKey       As String

'--- SCR 79741  Sherry Wang 10/12/2005
Private mbLoginInfoChanged As Boolean


Private Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Private Type TIME_ZONE_INFORMATION
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

Public Function GetNextID(Optional ByVal eSeqName As ICHCSII_SQL.eSeq = 0&) As Variant
'-- pulls the next value from the MainSeq sequence
    Dim ors As ADODB.Recordset
    Set ors = New ADODB.Recordset
    Dim sSeqName As String
    Dim lSeqNum As Long
    
    sSeqName = SeqEnumToString(eSeqName)
    If CWShared.UseMySQLDB Then
        moCnn.Execute "UPDATE " & sSeqName & " SET id=LAST_INSERT_ID(id+1)"
        ors.Open "SELECT LAST_INSERT_ID()", moCnn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Else
        ors.Open "Select " & sSeqName & ".nextval from dual", moCnn, adOpenForwardOnly, adLockOptimistic, adCmdText
    End If

    If Not ors.EOF Then
        GetNextID = ors(0)
    End If
    
    If ors.State > 0 Then ors.Close
    Set ors = Nothing

End Function

Private Function AppendAsChunk(ByVal vColumn As Variant, vText As Variant) As Boolean
'- for the column indicated, append the byte array as chunks
    Dim lsize       As Long
    Dim loffset     As Long
    Dim varChunk    As Variant
    Dim ChunkSize   As Long
    Dim Index       As Integer
'    Dim oZLib       As ZLib
    Dim lResult     As Long
    Dim sTxt As String
    
    If IsNumeric(vColumn) Then
        Index = vColumn
    Else
        Index = ColumnIndex(CStr(vColumn))
    End If
    
    
    lsize = LenB(vText)
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

    
    lsize = LenB(vText)
    loffset = 0 ' Reset offset
    
    Do While loffset < lsize
       varChunk = LeftB(RightB(vText, lsize - loffset), ChunkSize)
        mvarRecordSet.Fields(Index).AppendChunk varChunk
       loffset = loffset + ChunkSize
    Loop
    
    
    AppendAsChunk = True
    Exit Function
ErrHandler:

End Function


Private Function GetAsChunk(ByVal vColumn As Variant, vText As Variant) As Boolean
'- for the column indicated, get the data as chunks and return in the byte array
'- for the column indicated, append the byte array as chunks
    Dim lsize       As Long
    Dim loffset     As Long
    Dim varChunk    As Variant
    Dim Index       As Integer
    Dim ChunkSize   As Long
'    Dim oZLib       As ZLib
    Dim lResult     As Long
    Dim sTxt        As String
    
    If IsNumeric(vColumn) Then
        Index = vColumn
    Else
        Index = ColumnIndex(CStr(vColumn))
    End If
    
       
   
    lsize = mvarRecordSet.Fields(Index).ActualSize
    loffset = 0 ' Reset offset
    vText = ""
    
    '-- set chunksize
    ChunkSize = 2000
    If lsize > 10000 Then
        ChunkSize = lsize \ 10
        '- Make it EVEN (2-bytes per char, remember?)
        If ChunkSize \ 2 <> ChunkSize / 2 Then ChunkSize = ChunkSize + 1
        If ChunkSize > 20000 Then ChunkSize = 20000
    End If
    
    Do While loffset < lsize
       varChunk = mvarRecordSet.Fields(Index).GetChunk(ChunkSize)
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


Private Function SeqEnumToString(ByVal eiSeq As eSeq) As String
    If CWShared.UseMySQLDB Then
        Select Case eiSeq
            Case esMAINSEQ:         SeqEnumToString = "SEQ_MAIN"
            Case esAPPT_ID:         SeqEnumToString = "SEQ_APPT_ID"
            Case esAPPT_TRANS_ID:   SeqEnumToString = "SEQ_APPT_TRANS_ID"
            Case esENC_TEMPLATE_ID: SeqEnumToString = "SEQ_ENC_TEMPLATE_ID"
            Case esENCOUNTER_ID:    SeqEnumToString = "SEQ_ENCOUNTER_ID"
            Case esIEN:             SeqEnumToString = "SEQ_IEN"
            Case esPCM_CODE:        SeqEnumToString = "SEQ_PCM_CODE"
            Case esPI_DATA_ID:      SeqEnumToString = "SEQ_PI_DATA_ID"
            Case esUNIT_NUMBER:     SeqEnumToString = "SEQ_UNIT_NUMBER"
            Case esUSER_NUMBER:     SeqEnumToString = "SEQ_USER_NUMBER"
            Case esJDOE_NUMBER:     SeqEnumToString = "SEQ_JDOE_NUMBER"
            Case esTEMP_SSN:        SeqEnumToString = "SEQ_TEMP_SSN"
        End Select
    Else
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
    End If

End Function

Public Function GetNextNcid() As Variant
'-- pulls the next value from the NCID sequence
    Dim ors As ADODB.Recordset
    Set ors = New ADODB.Recordset
    Dim lSeqNum As Long
    
'<< scr 29784 mixed bag of mainseq and ncid sequencers were colliding.  So change this to MainSeq.
    If CWShared.UseMySQLDB Then
        moCnn.Execute "UPDATE seq_main SET id=LAST_INSERT_ID(id+1)"
        ors.Open "SELECT LAST_INSERT_ID()", moCnn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Else
        ors.Open "Select mainseq.nextval from dual", moCnn, adOpenForwardOnly, adLockOptimistic, adCmdText
    End If
    
    If Not ors.EOF Then
        GetNextNcid = ors(0)
    End If
    ors.Close
    Set ors = Nothing
End Function

Public Property Get Value(Column As Variant) As Variant
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "200"

    Dim Index As Integer
    On Error GoTo ErrHandler
    If BOF Or EOF Then
        GoTo ErrHandler
    End If
    
    If IsNumeric(Column) Then
        Index = Column
    Else
        Index = ColumnIndex(CStr(Column))
    End If
    
    Value = mvarRecordSet.Fields(Index).Value
'    If mvarRecordSet.Fields(index).Type = adDate Or mvarRecordSet.Fields(index).Type = adDBDate Or mvarRecordSet.Fields(index).Type = adDBTime Or mvarRecordSet.Fields(index).Type = adDBTimeStamp Then
'      If Value <> #12/31/9999# Then
'          ' adjust to local from gmt
'          Value = DateAdd("n", -lTimeBias, Value)
'      End If
'    End If
    If IsNull(Value) Then
        Select Case mvarRecordSet.Fields(Index).Type
            Case adBoolean
                Value = False
            Case adNumeric, adDecimal, adDouble, adSmallInt, adTinyInt, adCurrency, adSingle, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                Value = 0
            Case Else
                Value = ""
        End Select
    End If  ' IsNull(Value)
ErrHandler:
End Property
Public Property Let Value(Column As Variant, ByVal RHS As Variant)
    Dim Index As Integer
    On Error GoTo ErrHandler
    
    If BOF Or EOF Then
        GoTo ErrHandler
    End If
        
'- get the column
    If IsNumeric(Column) Then
        Index = Column
    Else
        Index = ColumnIndex(CStr(Column))
    End If
    
    If mvarRecordSet.Fields(Index).Type = adDate Or mvarRecordSet.Fields(Index).Type = adDBDate Or mvarRecordSet.Fields(Index).Type = adDBTime Or mvarRecordSet.Fields(Index).Type = adDBTimeStamp Then
'      If RHS <> #12/31/9999# Then
'          '- adjust from local to Zulu time by subtracting the time bias
'          RHS = AdjustDate(RHS)
'      End If
    End If

    mvarRecordSet.Fields(Index).Value = RHS

ErrHandler:

End Property

Public Property Get InfoType(Column As Variant) As String

    Dim Index As Integer
    If (BOF Or EOF) Then
        GoTo ErrHandler
    End If
    
    If IsNumeric(Column) Then
        Index = Column
    Else
        Index = ColumnIndex(CStr(Column))
    End If
    
    InfoType = mvarRecordSet.Fields(Index).Type  ' mvarDSQLConnect.Table(mvarAbsolutePosition, index).Value.InfoType

ErrHandler:

End Property

Public Property Get ColumnIndex(name As String) As Integer

    Dim i As Integer
    Dim sText As String
    For i = 0 To ColumnCount - 1
        If UCase(ColumnName(i)) = UCase(name) Then
            ColumnIndex = i
            Exit Property
        End If
    Next i
    
    sText = "There is no column named '" & name & "'.  Please select from:"
    For i = 0 To ColumnCount - 1
        sText = sText & Chr(10) & mvarRecordSet.Fields(i).name
    Next
    Err.Raise vbObjectError + 1, "SqlOps.ColumnIndex", sText
End Property

Public Property Get ColumnName(Index As Integer) As String

  
  ColumnName = ""
    If Not (BOF Or EOF) Then
        If Index > ColumnCount Then
            Err.Raise vbObjectError + 1, "SqlOps.ColumnName", "Index (" & Index & ") out of range.  There are " & ColumnCount & " colummns."
            Exit Property
        End If
        ColumnName = mvarRecordSet.Fields(Index).name
    End If
End Property

Public Property Get ColumnCount() As Integer

'    ColumnCount = mvarDSQLConnect.Table.Rows(mvarAbsolutePosition).Columns.Count
  ColumnCount = mvarRecordSet.Fields.count
End Property

Public Function Execute(SqlStatement As String, Optional ByVal direct As Boolean = False) As Variant
        
    Dim lRecCnt     As Long
    Dim FirstInsPt  As Long
    Dim SecondInsPt As Long
    Dim ValuesStrPt As Long                 '*-- Coleman 040401
    Dim words()     As String
    Dim WhereClause As String
    
    ' SCR #24346;   Developer:  06/05/2002 11:05 AM
    ' Note: Function Changed:  Execute    Impact of Change:  added call to FixParens()
    
    '- ensure there are always spaces between parens and words, exclude quoted string
    SqlStatement = FixParans(SqlStatement)
    Call RedirectEncrypted(SqlStatement)
    
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
    Select Case UCase$(Left$(SqlStatement, 6))
    
    Case "DELETE", "INSERT", "UPDATE"
        '- is as "Delete from tablename where field...
        moCnn.Execute SqlStatement, lRecCnt
        Execute = lRecCnt
                
    Case "SELECT"
        mvarRecordSet.Open SqlStatement, moCnn, adOpenDynamic, adLockPessimistic, adCmdText
       
    Case "EXECUTE"
        moCnn.Execute SqlStatement, lRecCnt, adCmdText
        
    Case Else
        moCnn.Execute SqlStatement, lRecCnt, adCmdText
        
    End Select
  Else  '- ubound(words) > 1
     mvarRecordSet.Open SqlStatement, moCnn, adOpenKeyset, adLockOptimistic, adCmdTable
  End If
  Exit Function
  Resume
  
End Function

Public Sub ExecuteMulti(SqlStatements As Collection, Optional ByVal direct As Boolean = False)
  Dim vStmt As Variant
  On Error GoTo ErrHandler
  
  moCnn.BeginTrans
  
  For Each vStmt In SqlStatements
    mvarSQLStatement = vStmt
    Execute CStr(vStmt), direct
  Next vStmt
      
  moCnn.CommitTrans
  Exit Sub
  
ErrHandler:
    MsgBox "error on: " & vStmt
  moCnn.RollBackTrans
  
End Sub


Public Sub MovePrevious()
    If mvarRecordSet.State = 0 Or mvarRecordSet.BOF Then Exit Sub
    mvarRecordSet.MovePrevious
End Sub

Public Sub MoveNext()
    If mvarRecordSet.State = 0 Or mvarRecordSet.EOF Then Exit Sub
    mvarRecordSet.MoveNext
End Sub

Public Sub MoveLast()
    If mvarRecordSet.State = 0 Then Exit Sub
    mvarRecordSet.MoveLast
    
End Sub

Public Sub MoveFirst()
    If mvarRecordSet.State = 0 Then Exit Sub
    mvarRecordSet.MoveFirst

End Sub

Public Property Get RecordCount() As Long
    On Error Resume Next
    RecordCount = mvarRecordSet.RecordCount
End Property

Public Property Get Source() As String:  Source = mvarSQLStatement: End Property
Public Property Get EOF() As Boolean:    EOF = mvarRecordSet.EOF:   End Property
Public Property Get BOF() As Boolean:    BOF = mvarRecordSet.BOF:   End Property

Private Sub Class_Initialize()
    On Error GoTo ErrHandler
    
    'SCR 128275
    Dim bLoggingIn As Boolean
    
    GetTimeZoneInformation TZInfo
    lTimeBias = TZInfo.Bias

    Set mobjShared = New CWShared
    UserNCID = mobjShared.CurrentUserNCID

    If mobjShared.UseSQLServer Then
        geDBUsed = edb_sqlserver
    ElseIf mobjShared.UseMySQLDB Then
        geDBUsed = edb_MySQL
    Else
        geDBUsed = edb_Oracle
    End If

    Dim suid As String
    Dim spwd As String
    Dim sdbkey As String
    
    If moCnn Is Nothing Then
        Set moCnn = New ADODB.Connection
        moCnn.ConnectionString = GetConnectionString(geDBUsed)
        moCnn.ConnectionTimeout = 30
        moCnn.CursorLocation = adUseClient
        
        bLoggingIn = True
        If moCnn.State = 0 Then moCnn.Open
        bLoggingIn = False
        
        If moCnn.State = 1 Then
            Call SetDBDateFormat(moCnn)
            Call mobjShared.GetLocalLogon(suid, spwd, sdbkey)
'            If suid = AppUID And sdbkey <> "" Then
'                moCnn.Execute "Insert into mmi_x (x0) values ('" & sdbkey & "')"
'            End If
        End If
    End If

  Set mvarRecordSet = New ADODB.Recordset

  Exit Sub
  
'--- SCR 79741   Sherry Wang   10/12/2005
ErrHandler:
   '--- Handle ORA-01017: invalid username/password; logon denied error
    If Err.number = -2147217843 Then
        Dim frmAppLogin  As frmC2TAppLogin
        
        Err.Clear
        mbLoginInfoChanged = True
        
        If frmAppLogin Is Nothing Then Set frmAppLogin = New frmC2TAppLogin
        frmAppLogin.Show vbModal
        
        If frmAppLogin.HasCancelled() Then
            Unload frmAppLogin
            Screen.MousePointer = vbDefault
        Else
            mobjShared.SetLocalLogon frmAppLogin.txtAppUserId, frmAppLogin.txtAppPassword
            SaveRegistry EncryptString(frmAppLogin.txtAppUserId), EncryptString(frmAppLogin.txtAppPassword), EncryptString(frmAppLogin.txtSecurityPassword)
        End If
    
    'SF SCR 128275
    ElseIf bLoggingIn = True And (InStr(Err.Description, "ORA-28001") Or InStr(Err.Description, "ORA-28000")) Then 'ORA-28001: the password has expired 'ORA-28000: the account is locked
        MsgBox "The application database password cannot login and may be either locked or expired." & vbCrLf & "Please contact your System Administrator to have the password reset." & vbCrLf & vbCrLf & "The application will now close.  Press OK to continue.", vbInformation
        mbLoginInfoChanged = True
    
    'SF SCR 128275
    ElseIf bLoggingIn = True And InStr(Err.Description, "ORA-28002") Then 'ORA-28002: The password will expire within %s days
        If gbShownDBErrorMessage = False Then
            MsgBox "The application database " & Mid(Err.Description, (InStr(Err.Description, "password"))) & "." & vbCrLf & "Please contact your System Administrator to update the application database passwords." & vbCrLf & vbCrLf & "Press OK to continue.", "OpenAHLTA Logon"
            gbShownDBErrorMessage = True
        End If
        Resume Next
        'Allow the user to continue and log in.
        
    'SF SCR 12875
    ElseIf bLoggingIn = True And InStr(Err.Description, "ORA-28011") Then 'ORA-28011: The account will expire soon; change your password now
        If gbShownDBErrorMessage = False Then
            MsgBox "The application database password will expire soon." & vbCrLf & "Please contact your System Administrator to update the application database passwords." & vbCrLf & vbCrLf & "Press OK to continue.", "OpenAHLTA Logon"
            gbShownDBErrorMessage = True
        End If
        Resume Next
        'Allow the user to continue and log in.
        
    'SF SCR 128275
    ElseIf bLoggingIn = True And InStr(Err.Description, "ORA-12541") Then 'ORA-12541: TNS:no listener
        MsgBox "The application cannot connect to the Oracle Listener." & vbCrLf & "The Oracle service may not be started." & vbCrLf & vbCrLf & "Please contact your System Administrator.", vbInformation
        mbLoginInfoChanged = True
    
    Else
         MsgBox "Error in SQLOps.Initialize: " & Err.Description
    End If
    Exit Sub
    Resume
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
    
'- handle the connection.  don't close it untill all users have teminated the use of this module
    If moCnn.State > 0 Then moCnn.Close
    Set moCnn = Nothing

    
End Sub

Public Function Reconnect() As Boolean
'-- we assume the connection string info in cwshared has changed.
    
    Set moCnn = New ADODB.Connection
    moCnn.ConnectionString = GetConnectionString(geDBUsed)
    moCnn.ConnectionTimeout = 30
    moCnn.CursorLocation = adUseClient
    If moCnn.State = 0 Then moCnn.Open
    If moCnn.State = 1 Then
        Call SetDBDateFormat(moCnn)
    End If

End Function
Private Function OpenDB()
  'Set moCnn = DE.Cnn

  Set moCnn = New ADODB.Connection

  moCnn.ConnectionString = GetConnectionString(geDBUsed)
  moCnn.Open
    If moCnn.State = 1 Then
        Call SetDBDateFormat(moCnn)
    End If
  
End Function


Public Function AddNew() As Boolean
   If mvarRecordSet Is Nothing Then Exit Function
   On Error GoTo ErrHandler
   
   mvarRecordSet.AddNew
   AddNew = True

ErrHandler:
   
End Function

Public Function Update() As Boolean
    On Error GoTo ErrHandler
    mvarRecordSet.Update
    Update = True
ErrHandler:
End Function

Public Function Cancel() As Boolean
    On Error GoTo ErrHandler
    mvarRecordSet.Cancel
    Cancel = True
ErrHandler:
End Function

Private Property Let ISqlOpsEx_Value(Column As Variant, ByVal RHS As Variant)
    Dim Index As Integer
    On Error GoTo ErrHandler
    
    If BOF Or EOF Then
        GoTo ErrHandler
    End If
        
'- get the column
    If IsNumeric(Column) Then
        Index = Column
    Else
        Index = ColumnIndex(CStr(Column))
    End If
    
    If mvarRecordSet.Fields(Index).Type = adDate Or mvarRecordSet.Fields(Index).Type = adDBDate Or mvarRecordSet.Fields(Index).Type = adDBTime Or mvarRecordSet.Fields(Index).Type = adDBTimeStamp Then
'      If RHS <> #12/31/9999# Then
'          '- adjust from local to Zulu time by subtracting the time bias
'          RHS = AdjustDate(RHS)
'      End If
    End If

    mvarRecordSet.Fields(Index).Value = RHS

ErrHandler:

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

Private Sub RedirectEncrypted(ByRef sSQL As String)
    Dim words() As String
    Dim TablePosn As Integer
    
    words = Split(sSQL, " ")
    
    If UBound(words) > 1 Then
      Select Case UCase$(Left$(sSQL, 6))
      Case "DELETE", "INSERT"
          TablePosn = 2
      Case "UPDATE"
          TablePosn = 1
      Case Else
          Exit Sub
      End Select
      
    Else  '- this is just a tablename
      TablePosn = 0
    End If
    
    If EncryptedTable(words(TablePosn)) Then
       words(TablePosn) = words(TablePosn) & "_X"
        '__ now rebuild it, but only if we messed with it
       sSQL = Join(words, " ")
    End If
    
End Sub


Public Function EncryptedTable(ByVal TableName As String) As Boolean
'__ check with the 'EncryptedTable' table to see if this is a table that is encrypted
    Dim sMember As Variant
   
    If gcolEncrypted Is Nothing Then
        If Not LoadEncrypted Then
            EncryptedTable = False
            Exit Function
        End If
    End If
    
    On Error Resume Next
    sMember = gcolEncrypted(TableName)
    '__ this will error-out if the tablename doesn't exist in the col
     EncryptedTable = (Err.number = 0)
    
End Function

Private Function LoadEncrypted() As Boolean
    If bCheckingEncrypted Then Exit Function
    
    bCheckingEncrypted = True '-- prevent recursion
    
    Dim ors As ADODB.Recordset
    Dim oSQL As SqlOps
    Set oSQL = GetAppConnection
    If oSQL Is Nothing Then
        bCheckingEncrypted = False
        LoadEncrypted = False
        Exit Function
    End If
    oSQL.Execute "Encrypted"
    Set gcolEncrypted = New Collection
    Do While Not oSQL.EOF
        gcolEncrypted.Add "" & oSQL(0), "" & oSQL(0)
        oSQL.MoveNext
    Loop
    Set oSQL = Nothing
    bCheckingEncrypted = False
    LoadEncrypted = True
End Function

'--- SCR 79741   Sherry Wang   10/12/2005
'Save the Application Username and Password to Registry.
Private Sub SaveRegistry(ByVal sApp_UID As String, ByVal sApp_Pwd As String, ByVal sFramework_Pwd As String)
    
    Dim objReg As RegObj.Registry
    Dim objKey   As RegObj.RegKey
    Dim objValue As RegObj.RegValue
    Dim oShared  As New CWShared
    
    
    On Error GoTo ErrHandler
    
    Set objReg = New RegObj.Registry
    Set objKey = objReg.RegKeyFromString(oShared.RegistryKey & "\Support")
             
    If Not objKey Is Nothing Then
        For Each objValue In objKey.Values
          
            If UCase(Trim$(objValue.name)) = "APP_UID" Then
              objValue.Value = sApp_UID
            ElseIf UCase(Trim$(objValue.name)) = "APP_PWD" Then
              objValue.Value = sApp_Pwd
            ElseIf UCase(Trim$(objValue.name)) = "FRMWK_PWD" Then
              objValue.Value = sFramework_Pwd
            End If
        
        Next objValue
    End If

GoTo Cleanup

ErrHandler:
    MsgBox "Error: " & Err.Description & " occurred in SqlOps.SaveRegistryAppUser"
    Exit Sub

Cleanup:
    Set objReg = Nothing
    Set objKey = Nothing
    Set objValue = Nothing
End Sub


'--- SCR 79741   Sherry Wang   10/12/2005
Private Function EncryptString(ByVal sStr As String) As String
    Const sExp As String = "trtkabasnbi"
    Dim oCrypto  As CCrypto
    
    On Error GoTo ErrHandler
    
    Set oCrypto = New CCrypto
    EncryptString = oCrypto.Encrypt(sStr, oCrypto.GetHash(AppName & sExp, True), True)
    
    Exit Function
ErrHandler:
    Set oCrypto = Nothing
    Resume Next
End Function

'--- SCR 79741  Sherry Wang 10/12/2005
Public Property Get LoginInfoChanged() As Boolean
    LoginInfoChanged = mbLoginInfoChanged
End Property


