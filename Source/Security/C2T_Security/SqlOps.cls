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

Private msConnString As String
Public moCnn                   As ADODB.Connection

Private Const SQLQueryNCID As String = "1450582"
Private Const OraErrorNoDataFound As String = "ORA-00100"
Private Const OraErrorFetchOutOfSequence As String = "ORA-01002"
Private mobjShared As CWShared

'--- SCR 79741   Sherry Wang  10/12/2005
Private mbLoginInfoChanged As Boolean

'Private moCnn As ADODB.Connection    ' let's keep it common/shared connection.
                                        'but, give a way to reset it with a new Connection string
Private mvarRecordSet As ADODB.Recordset
Public UserNCID As String
Private mvarSQLStatement As String

Private mFacilityNCID As String  '- for backfill of providers
Private mClinicNCID   As String

Enum eSeq
    esMAINSEQ = 0
    esAPPT_ID
    esAPPT_TRANS_ID
    esENC_TEMPLATE_ID
    esENCOUNTER_ID
    esIEN
    esPCM_CODE
    esPI_DATA_ID
    esUNIT_NUMBER
    esUSER_NUMBER
    esJDOE_NUMBER
    esTEMP_SSN
End Enum


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


Private Function ISqlOpsEx_AppendAsChunk(ByVal vColumn As Variant, vText As Variant) As Boolean
'- for the column indicated, append the byte array as chunks
    Dim lSize       As Long
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
        mvarRecordSet.Fields(Index).AppendChunk varChunk
       loffset = loffset + ChunkSize
    Loop
    
    
    ISqlOpsEx_AppendAsChunk = True
    Exit Function
ErrHandler:

End Function


Private Function ISqlOpsEx_GetAsChunk(ByVal vColumn As Variant, vText As Variant) As Boolean
'- for the column indicated, get the data as chunks and return in the byte array
'- for the column indicated, append the byte array as chunks
    Dim lSize       As Long
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
    
       
   
    lSize = mvarRecordSet.Fields(Index).ActualSize
    loffset = 0 ' Reset offset
    vText = vbNullString
    
    '-- set chunksize
    ChunkSize = 2000
    If lSize > 10000 Then
        ChunkSize = lSize \ 10
        '- Make it EVEN (2-bytes per char, remember?)
        If ChunkSize \ 2 <> ChunkSize / 2 Then ChunkSize = ChunkSize + 1
        If ChunkSize > 20000 Then ChunkSize = 20000
    End If
    
    Do While loffset < lSize
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

Private Function GetNextNcid() As Variant
'-- pulls the next value from the NCID sequence
    Dim ors As ADODB.Recordset
    Set ors = New ADODB.Recordset
    Dim lSeqNum As Long
    
    If geDBUsed = edb_sqlserver Then
        ors.Open "Select Value from Sequences where seq_type = 'PROVIDER_NCID'", moCnn, adOpenForwardOnly, adLockOptimistic, adCmdText
        If Not ors.EOF Then
            lSeqNum = Val(ors(0))
        End If
        moCnn.Execute "update Sequences set value = " & lSeqNum + 1 & " where seq_type = 'PROVIDER_NCID'"
        GetNextNcid = lSeqNum
    Else
'<< scr 29784 mixed bag of mainseq and ncid sequencers were colliding.  So change this to MainSeq.
        ors.Open "Select mainseq.nextval from dual", moCnn, adOpenForwardOnly, adLockOptimistic, adCmdText
        If Not ors.EOF Then
            GetNextNcid = ors(0)
        End If
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
                Value = vbNullString
        End Select
    End If  ' IsNull(Value)
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

Public Property Get ColumnIndex(Name As String) As Integer

    Dim i As Integer
    Dim sText As String
    For i = 0 To ColumnCount - 1
        If UCase(ColumnName(i)) = UCase(Name) Then
            ColumnIndex = i
            Exit Property
        End If
    Next i
    
    sText = "There is no column named '" & Name & "'.  Please select from:"
    For i = 0 To ColumnCount - 1
        sText = sText & Chr(10) & mvarRecordSet.Fields(i).Name
    Next
    Err.Raise vbObjectError + 1, "CHCSII_SQL:SqlOps.ColumnIndex", sText
End Property

Public Property Get ColumnName(Index As Integer) As String

  
  ColumnName = vbNullString
    If Not (BOF Or EOF) Then
        If Index > ColumnCount Then
            Err.Raise vbObjectError + 1, "CHCSII_SQL:SqlOps.ColumnName", "Index (" & Index & ") out of range.  There are " & ColumnCount & " colummns."
            Exit Property
        End If
        ColumnName = mvarRecordSet.Fields(Index).Name
    End If
End Property

Public Property Get ColumnCount() As Integer

'    ColumnCount = mvarDSQLConnect.Table.Rows(mvarAbsolutePosition).Columns.Count
  ColumnCount = mvarRecordSet.Fields.Count
End Property

Public Function Execute(SqlStatement As String, Optional ByVal direct As Boolean = False) As Variant
        
    Dim lRecCnt     As Long
    Dim FirstInsPt  As Long
    Dim SecondInsPt As Long
    Dim ValuesStrPt As Long                 '*-- Coleman 040401
    Dim words()     As String
    Dim WhereClause As String
    Dim sTableName As String

    'SF SCR 89066 3/2/06
    SqlStatement = Replace(SqlStatement, vbCrLf, " ")

    ' SCR #24346;   Developer:  06/05/2002 11:05 AM
    ' Note: Function Changed:  Execute    Impact of Change:  added call to FixParens()
    
    '- ensure there are always spaces between parens and words, exclude quoted string
    SqlStatement = FixParans(SqlStatement)
    
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
'        mvarRecordSet.CancelUpdate
        mvarRecordSet.Close
        On Error GoTo 0
    End If
    
  End If
  
  If moCnn Is Nothing Then
    OpenDB
  End If
  
  If moCnn.State = 0 Then
    moCnn.Open
  End If
  
  If UBound(words) > 1 Then
    Select Case UCase$(Left$(SqlStatement, 6))
    
    Case "DELETE", "INSERT", "UPDATE"
        'SCR 83806
        If UCase$(Left$(SqlStatement, 6)) = "UPDATE" Then
            sTableName = words(1)
        Else
            sTableName = words(2)
        End If
        
        moCnn.Execute SqlStatement, lRecCnt
        Execute = lRecCnt
        
    Case "SELECT"
        'SCR 83806
        Dim iFrom As Integer
        Dim iTable As Integer
        Dim iEndTable As Integer
    
        iFrom = InStr(UCase(SqlStatement), "FROM")
        If iFrom > 0 Then
            iEndTable = InStr(iFrom + Len("FROM "), UCase(SqlStatement), " ")
            sTableName = Mid(UCase(SqlStatement), iFrom + Len("FROM "), iEndTable - (iFrom + Len("FROM ")))
        End If
        
        mvarRecordSet.Open SqlStatement, moCnn, adOpenDynamic, adLockPessimistic, adCmdText
       
    Case "EXECUT"
        'SF SCR 89066 3/2/06
        sTableName = SqlStatement
        moCnn.Execute SqlStatement, lRecCnt, adCmdText
        
    End Select
  Else  '- ubound(words) > 1
     mvarRecordSet.Open SqlStatement, moCnn, adOpenKeyset, adLockOptimistic, adCmdTable
  End If
  
  'SCR 83806
  Call TMIPAudit("C2T_DAS", "GEMS_DAS.ExecuteSQL", CurrentUserName, sTableName, SqlStatement, Err.Description, moCnn)

  
End Function

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
    
    If moCnn Is Nothing Then
        Set moCnn = New ADODB.Connection
        moCnn.ConnectionString = GetConnectionString(geDBUsed)
        moCnn.ConnectionTimeout = 30
        moCnn.CursorLocation = adUseClient
        bLoggingIn = True
        If moCnn.State = 0 Then moCnn.Open
        bLoggingIn = False

        'Both gems and secadmin accounts must log in successfully to restore the db error message.
        If InStr(moCnn.ConnectionString, "gems") > 0 And moCnn.Errors.Count = 0 Then
            gbGEMSAccountOK = True
        End If
        If InStr(moCnn.ConnectionString, "secadmin") > 0 And moCnn.Errors.Count = 0 Then
            gbSCADBAccountOK = True
        End If
        If gbSCADBAccountOK And gbGEMSAccountOK Then
            'Both accounts secadmin and gems have successfully logged in, so clear the switch that hides the error message for ORA-28002 and ORA-28011 below
            'It is stored in HKey_Current_User\Software\VB and VBA Program Settings
            SaveSetting App.EXEName, "Logon", "DBErrorMessage", "1"
        End If
        
        If moCnn.State = 1 Then
            Call SetDBDateFormat(moCnn)
        End If
    End If

  Set mvarRecordSet = New ADODB.Recordset
  
  Exit Sub
  
'--- SCR 79741   Sherry Wang   10/12/2005
ErrHandler:
   '--- Handle ORA-01017: invalid username/password; logon denied error
   If Err.Number = -2147217843 Then
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
        MsgBxARMd "The application database password cannot login and may be either locked or expired." & vbCrLf & "Please contact your System Administrator to have the password reset." & vbCrLf & vbCrLf & "The application will now close.  Press OK to continue.", vbInformation
        mbLoginInfoChanged = True
    
    'SF SCR 128275
    ElseIf bLoggingIn = True And InStr(Err.Description, "ORA-28002") Then 'ORA-28002: The password will expire within %s days
        frmMessage.DisplayMessage "The application database " & Mid(Err.Description, (InStr(Err.Description, "password"))) & "." & vbCrLf & "Please contact your System Administrator to update the application database passwords." & vbCrLf & vbCrLf & "Press OK to continue.", "OpenAHLTA Logon"
        Resume Next
        'Allow the user to continue and log in.
        
    'SF SCR 12875
    ElseIf bLoggingIn = True And InStr(Err.Description, "ORA-28011") Then 'ORA-28011: The account will expire soon; change your password now
        frmMessage.DisplayMessage "The application database password will expire soon." & vbCrLf & "Please contact your System Administrator to update the application database passwords." & vbCrLf & vbCrLf & "Press OK to continue.", "OpenAHLTA Logon"
        Resume Next
        'Allow the user to continue and log in.
        
    'SF SCR 128275
    ElseIf bLoggingIn = True And InStr(Err.Description, "ORA-12541") Then 'ORA-12541: TNS:no listener
        MsgBxARMd "The application cannot connect to the Oracle Listener." & vbCrLf & "The Oracle service may not be started." & vbCrLf & vbCrLf & "Please contact your System Administrator.", vbInformation
        mbLoginInfoChanged = True
'   ElseIf bLoggingIn = True Then
'        'SCR 128275 The gems or secadmin oracle account is locked
'        Dim sLogin As String
'        Dim spwd As String
'        mobjShared.GetLocalLogon sLogin, spwd
'        MsgBxARMd "The application cannot login with the '" & sLogin & "' oracle account and cannot continue.  The account may be locked or the password expired." & vbCrLf & "The application will close after the 'OK' button is pressed.", vbOKOnly + vbInformation
'        mbLoginInfoChanged = True
   Else
        MsgBxARMd "Error in SQLOps.Initialize: " & Err.Description
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
Dim DBType As eDBType

    If mobjShared.UseSQLServer Then
        DBType = edb_sqlserver
    ElseIf mobjShared.UseMySQLDB Then
        DBType = edb_MySQL
    Else
        DBType = edb_Oracle
    End If

  moCnn.ConnectionString = GetConnectionString(DBType)
'  moCnn.ConnectionTimeout = 30
  moCnn.Open
    If moCnn.State = 1 And moCnn.Properties("DBMS Name") = "Microsoft SQL Server" Then    '-- ITT
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

Private Function GetConnectionString(ByVal DBType As eDBType) As String

    Dim sNSN As String
    Dim iDS  As Integer
    Dim iSC  As Integer
    Dim sDSN As String
    Dim sInitCat As String
    Dim suid As String
    Dim spwd As String
    Dim SearchStr As String
    
    Select Case DBType
    
        Case edb_Oracle
            msConnString = GEMS_Conn
            GetConnectionString = msConnString
            sNSN = mobjShared.NetServiceName
            If sNSN <> vbNullString Then
                GetConnectionString = ReplaceParameter(GetConnectionString, "DATA SOURCE=", sNSN)
            End If
            If mobjShared.GetLocalLogon(suid, spwd) Then
                 GetConnectionString = ReplaceParameter(GetConnectionString, "USER ID=", suid)
                 GetConnectionString = ReplaceParameter(GetConnectionString, "PASSWORD=", spwd)
            End If
            
        Case edb_MySQL
            If mobjShared.GetLocalLogon(suid, spwd) Then
                If suid = "secadmin" Then
                    msConnString = MYSQL_SEC_Conn
                Else
                    msConnString = MYSQL_GEMS_Conn
                End If
                GetConnectionString = msConnString
                GetConnectionString = ReplaceParameter(GetConnectionString, "SERVER=", mobjShared.MySQLDataSource)
                GetConnectionString = ReplaceParameter(GetConnectionString, "UID=", suid)
                GetConnectionString = ReplaceParameter(GetConnectionString, "PWD=", spwd)
            End If
        
        Case edb_sqlserver
    
    End Select

End Function

Public Function ExecuteStoredProc(oAdoCmd As ADODB.Command) As Variant
    oAdoCmd.ActiveConnection = moCnn
    Set ExecuteStoredProc = oAdoCmd.Execute
End Function



'--- SCR 79741   Sherry Wang   10/12/2005
Public Property Get LoginInfoChanged() As Boolean
    LoginInfoChanged = mbLoginInfoChanged
End Property


