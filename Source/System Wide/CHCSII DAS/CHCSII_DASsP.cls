VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHCSII_DASsP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Implements ICHCSII_DAS


Private oSqlConnect As SQLSERVATLLib.SqlServConnect

Private oStatement As SQLSERVATLLib.SqlServStatement

Private m_saRelevantTables() As String

Private Const SQLQueryNCID As String = "1450582"
Private Const OraErrorNoDataFound As String = "ORA-00100"
Private Const OraErrorFetchOutOfSequence As String = "ORA-01002"
Private Const cnst_MAX_NUM_OF_RECORDS = 5000 '* Add by j.ray 31AUG2001 to fix SCR #16462


Private TZInfo As TIME_ZONE_INFORMATION

Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private bSkipFacilityWorkAround As Boolean     '- so mid-tier can function...
Public Property Let ICHCSII_DAS_FacilityWorkaround(ByVal Enabled As Boolean)
    bSkipFacilityWorkAround = Not Enabled
End Property
Public Function OpenRecordsetSP(Params As CHCSII_Command.CHCSII_Cmd) As ADODB.Recordset

    Call ExecuteSP(Params)
    
    Set OpenRecordsetSP = ProcessRecordSet(oSqlConnect.Transaction(1))

End Function

Public Sub ExecuteSP(Params As CHCSII_Command.CHCSII_Cmd)

    Dim mSPCols As Collection
    Set mSPCols = New Collection
    
    mSPCols.Add Params
    Call ExecutMultiSP(mSPCols)
        
End Sub
Private Sub ExecutMultiSP(mParams As Collection)
   
    Dim vParam  As CHCSII_Command.CHCSII_SPparam
    Dim sParam  As CHCSII_Command.CHCSII_Cmd
    Dim oBindVar As SQLSERVATLLib.SqlServBindVar
    Dim i As Integer
    
    oSqlConnect.ClearAll
    
    For Each sParam In mParams
        Set oStatement = oSqlConnect.Transaction.Add
        
        With oStatement
            .ProcedureName = sParam.SPName
            For i = 1 To sParam.Count
                Set vParam = sParam.Item(i)
                Set oBindVar = .BindVariables.Add
                If Not vParam.InfoType = "" Then
                    oBindVar.Value.InfoType = vParam.InfoType
                    If vParam.InfoType = "date" Then
                        oBindVar.Value.DateNoTZConversion = vParam.Param
                    ElseIf vParam.InfoType = "textString" Then
                        oBindVar.Value.TextString = vParam.Param
                    End If
                End If
                oBindVar.IsCursor = vParam.IsCursor
                oBindVar.IsOutput = vParam.IsOutput
            Next
        End With
    Next
    
    oSqlConnect.ExecuteTransaction

End Sub


Public Function ProcessRecordSet(oStatement As SQLSERVATLLib.SqlServStatement) As ADODB.Recordset

Dim objRecordset As New Recordset
Dim iNumRows As Integer
Dim iNumCols As Integer
Dim i As Integer
Dim j As Integer

Dim sNewSQL As String
Dim sOraError As String
Dim sFieldName As String
Dim sInfoType As String
Dim eDataType As DataTypeEnum
Dim sStringValue As String


    With oStatement
        
        iNumRows = .Rows.Count

        If iNumRows > 0 Then
            iNumCols = .ColumnNames.Count

            If iNumRows = 1 Then

                For i = 1 To iNumCols
                    sFieldName = .ColumnNames(i) ' scr 27474
                    sInfoType = .Rows(1)(i).InfoType
                    Select Case sInfoType
                        Case "textString"
                            eDataType = adBSTR
                        Case "binaryString"
                            eDataType = adBinary
                        Case "date"
                            eDataType = adDate
                    End Select
                    objRecordset.Fields.Append sFieldName, eDataType
                Next
            Else
                For i = 1 To iNumCols
                    sFieldName = .ColumnNames(i)
                    
                    sInfoType = .Rows(1)(i).InfoType
                    Select Case sInfoType
                        Case "textString"
                            eDataType = adBSTR
                        Case "binaryString"
                            eDataType = adBinary
                        Case "date"
                            eDataType = adDate
                    End Select
                    objRecordset.Fields.Append sFieldName, eDataType
                Next
            End If

            'Get rows next
            objRecordset.Open
            For i = 1 To iNumRows
                objRecordset.AddNew
                For j = 1 To iNumCols
                    objRecordset.Fields.Item(j - 1).Value = getValue(.Rows(i)(j))
                Next
            Next
            objRecordset.MoveFirst
'******** The following block was commented as it was causing a runtime error opening the recordset with no fields appended
'        ElseIf oSqlConnect.Transaction(1).BindVariables.Count > 0 Then ' get BindVariables
'            Dim k As Integer
'            Dim sVars As String
'
'            With oSqlConnect.Transaction(1).BindVariables
'                objRecordset.Open
'                For k = 1 To .Count
'                    If (.Item(k).IsCursor) Then
'                        sVars = sVars & "Cursor Variable"
'                    Else
'                        If (.Item(k).IsOutput) Then
'                            'sVars = sVars & "output"
'                            If (.Item(k).Value.InfoType = "textString") Then
'                                sVars = .Item(k).Value.TextString
'                            Else
'                                sVars = .Item(k).Value.DateNoTZConversion
'                            End If
'                            objRecordset.AddNew
'                            objRecordset.Fields.Item(0).Value = sVars
'                        End If
'                    End If
'                Next
'                objRecordset.MoveFirst
'            End With
'******** The previous block was commented as it was causing a runtime error opening the recordset with no fields appended
        Else
            'Create dummy field and open recordset
            objRecordset.Fields.Append "No Records Found", adBSTR
            objRecordset.Open
        End If
    End With


    Set ProcessRecordSet = objRecordset

End Function


Private Sub Class_Initialize()

    bSkipFacilityWorkAround = False
    
    LoadRelevantWorkaroundTables
    OpenDatabase
    GetTimeZoneInformation TZInfo

End Sub

Private Sub Class_Terminate()

    CloseDatabase

End Sub

Private Sub OpenDatabase()

    Set oSqlConnect = New SQLSERVATLLib.SqlServConnect

End Sub

Private Sub CloseDatabase()

    Set oSqlConnect = Nothing

End Sub

Private Function DoFacilityNCIDWorkaround(sSQL As String) As String

Dim iFromIndex As Integer
Dim iWhereIndex As Integer
Dim iOrderByIndex As Integer
Dim iUnionIndex As Integer
Dim sTableNames As String
Dim saTableNames() As String
Dim bIsTableRelevant As Boolean
Dim bDoLabTestFix As Boolean
Dim sTableToCheck As String
Dim i As Integer
Dim j As Integer
Dim sAdditionalWhereClause As String
Dim sBeginClause As String
Dim sEndClause As String
Dim sNewSQL As String
Dim sFromClause As String
Dim sFacilityNCID As String

If Len(sSQL) < 1 Or bSkipFacilityWorkAround Then
    DoFacilityNCIDWorkaround = sSQL
    Exit Function
End If

    Dim oShared As MMMHISShared.CWShared
    Set oShared = New MMMHISShared.CWShared
    sFacilityNCID = oShared.CurrentFacilityNCID
    Set oShared = Nothing

'Check for UNION
iUnionIndex = InStr(1, sSQL, "UNION ", vbTextCompare)
If (iUnionIndex > 0) Then
    DoFacilityNCIDWorkaround = HandleUnion(sSQL)
    Exit Function
End If

'Parse SQL statement looking for FROM, WHERE, ORDER BY
iFromIndex = InStr(1, sSQL, "FROM ", vbTextCompare)
If (iFromIndex < 1) Then
    DoFacilityNCIDWorkaround = sSQL
    Exit Function
End If

iWhereIndex = InStr(iFromIndex, sSQL, "WHERE ", vbTextCompare)
If (iWhereIndex > 0) Then
    iOrderByIndex = InStr(iWhereIndex, sSQL, "ORDER BY ", vbTextCompare)
    sTableNames = Mid(sSQL, iFromIndex + 5, iWhereIndex - (iFromIndex + 5))
Else
    iOrderByIndex = InStr(iFromIndex, sSQL, "ORDER BY ", vbTextCompare)
    If (iOrderByIndex > 0) Then
        sTableNames = Mid(sSQL, iFromIndex + 5, iOrderByIndex - (iFromIndex + 5))
    Else
        sTableNames = Right(sSQL, Len(sSQL) - (iFromIndex + 4))
    End If
End If

saTableNames = Split(sTableNames, ",")
For i = 0 To UBound(saTableNames)
    saTableNames(i) = Trim(saTableNames(i))
Next

'Figure out if we care about these tables
bIsTableRelevant = False
bDoLabTestFix = False
For i = 0 To UBound(saTableNames)
    sTableToCheck = UCase(saTableNames(i))
    
    If (sTableToCheck = "LAB_TEST") Then
        bDoLabTestFix = True
    End If
    
    For j = 1 To UBound(m_saRelevantTables)
        If (sTableToCheck = m_saRelevantTables(j)) Then
            bIsTableRelevant = True
            Exit For
        End If
    Next
    
    If (bIsTableRelevant) Then
        Exit For
    End If
Next

If (bIsTableRelevant) Then
    'Now that we care about the table(s), make the Facility_NCID fix
    For i = 0 To UBound(saTableNames)
        If (i > 0) Then
            sAdditionalWhereClause = sAdditionalWhereClause & " And " & saTableNames(i) & ".facility_ncid = " & sFacilityNCID
        Else
            sAdditionalWhereClause = saTableNames(i) & ".facility_ncid = " & sFacilityNCID
        End If
    Next
    
    If (iWhereIndex > 0) Then
        'Have WHERE clause
        sBeginClause = Left(sSQL, iWhereIndex + 5)
        sEndClause = Right(sSQL, Len(sSQL) - (iWhereIndex + 5))
        sNewSQL = sBeginClause & sAdditionalWhereClause & " AND " & sEndClause
    ElseIf (iWhereIndex < 1 And iOrderByIndex > 0) Then
        'Only have ORDER BY clause
        sBeginClause = Left(sSQL, iOrderByIndex - 1)
        sEndClause = Right(sSQL, Len(sSQL) - (iOrderByIndex - 1))
        sNewSQL = sBeginClause & " WHERE " & sAdditionalWhereClause & " " & sEndClause
    Else
        'Only FROM clause
        sNewSQL = sSQL & " WHERE " & sAdditionalWhereClause
    End If
    
    'Do special fix for lab_test table
    If (bDoLabTestFix) Then
        sBeginClause = Left(sNewSQL, iFromIndex + 4)
        
        iWhereIndex = InStr(iFromIndex, sNewSQL, "WHERE ", vbTextCompare)
        sEndClause = Right(sNewSQL, Len(sNewSQL) - (iWhereIndex - 1))
        
        For i = 0 To UBound(saTableNames)
            If (UCase(saTableNames(i)) = "LAB_TEST") Then
                If (i > 0) Then
                    sFromClause = sFromClause & ", lab_test_oe lab_test"
                Else
                    sFromClause = "lab_test_oe lab_test"
                End If
            Else
                If (i > 0) Then
                    sFromClause = sFromClause & ", " & saTableNames(i)
                Else
                    sFromClause = saTableNames(i)
                End If
            End If
        Next
        
        sNewSQL = sBeginClause & sFromClause & " " & sEndClause
        
    End If
    
    If Right(sNewSQL, 1) = ";" Then
        sNewSQL = Left(sNewSQL, Len(sNewSQL) - 1)
    End If

    DoFacilityNCIDWorkaround = sNewSQL
Else
    DoFacilityNCIDWorkaround = sSQL
End If

End Function

Private Function HandleUnion(sSQL As String) As String

Dim sSQLSegments() As String
Dim iUnionIndex As Integer
Dim sNewSQL As String
Dim i As Integer

iUnionIndex = InStr(1, sSQL, " UNION ", vbTextCompare)
If (iUnionIndex > 0) Then
    sSQLSegments = Split(sSQL, " UNION ")
    For i = 0 To UBound(sSQLSegments)
        If (i > 0) Then
            sNewSQL = sNewSQL & " UNION " & DoFacilityNCIDWorkaround(sSQLSegments(i))
        Else
            sNewSQL = DoFacilityNCIDWorkaround(sSQLSegments(i))
        End If
    Next
        
    HandleUnion = sNewSQL
    Exit Function
Else
    HandleUnion = sSQL
End If

End Function

Private Sub LoadRelevantWorkaroundTables()

    ReDim m_saRelevantTables(1 To 15)
    
    m_saRelevantTables(1) = "LAB_PANEL"
    m_saRelevantTables(2) = "LAB_SPECIMEN"
    m_saRelevantTables(3) = "LAB_TEST"
    m_saRelevantTables(4) = "LAB_TEST_SYNONYM"
    m_saRelevantTables(5) = "LAB_TEST_TYPE"
    m_saRelevantTables(6) = "LAB_TOPO"
    m_saRelevantTables(7) = "PHARMACY"
    m_saRelevantTables(8) = "RAD_TEST"
    m_saRelevantTables(9) = "RAD_TEST_SYNONYM"
    m_saRelevantTables(10) = "RAD_TEST_TYPE"
    m_saRelevantTables(11) = "RX_LIST"
    m_saRelevantTables(12) = "RX_SYNONYM"
    m_saRelevantTables(13) = "LAB_TEST_WTYPE"
    m_saRelevantTables(14) = "MEPRS"
    '<< SCR 19158  Added:
    m_saRelevantTables(15) = "LAB_TEST_OE"
    'm_saRelevantTables(16) = "LAB_TEST_OE"

End Sub

Private Function AdjustDate(ByVal dDate As Date) As Date
    Dim lYear As Integer
    Dim dStandard As Date
    Dim dDaylight As Date
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

Function CreateRecordsetFromFile(sFileName As String, sDelimiter As String) As ADODB.Recordset

Dim iFileNum As Integer
Dim i As Integer
Dim lRow As Long
Dim sData As String
Dim sHeaderRow As String
Dim sValues() As String
Dim sColumnNames() As String
Dim bIsOpen As Boolean
Dim objRecordset As New ADODB.Recordset

iFileNum = FreeFile

Open sFileName For Input As #iFileNum

Do While Not EOF(iFileNum)
    Line Input #iFileNum, sData
    If Left(sData, 3) <> "***" Then
        'Assume first good row is column headers
        If lRow = 0 Then
            sHeaderRow = sData
            lRow = lRow + 1
        Else
            sValues = Split(sData, sDelimiter)
            If lRow = 1 Then
                'Do headers
                sColumnNames = Split(sHeaderRow, sDelimiter)
                
                For i = LBound(sValues) To UBound(sValues)
                    If sColumnNames(i) <> "" Then
                        If IsDate(sValues(i)) = True And InStr(1, sValues(i), "-") < 1 Then
                            objRecordset.Fields.Append sColumnNames(i), adDate
                        Else
                            objRecordset.Fields.Append sColumnNames(i), adBSTR
                        End If
                    End If
                Next
                
            End If
            
            'Do Data
            If Not (bIsOpen) Then
                objRecordset.Open
                bIsOpen = True
            End If
            
            objRecordset.AddNew
            For i = 0 To objRecordset.Fields.Count - 1
                If objRecordset.Fields.Item(i).Type = adDate Then
                    objRecordset.Fields.Item(i).Value = CDate(sValues(i))
                Else
                    If sValues(i) <> "" Then
                        objRecordset.Fields.Item(i).Value = CStr(sValues(i))
                    End If
                End If
            Next
            
            lRow = lRow + 1
        End If
    End If
Loop

If bIsOpen Then
    objRecordset.MoveFirst
End If

Close #iFileNum

Set CreateRecordsetFromFile = objRecordset

End Function


Private Property Get ICHCSII_DAS_ActiveConnection() As ADODB.Connection
    Set ICHCSII_DAS_ActiveConnection = Nothing
End Property

'- GEMS
Private Function ICHCSII_DAS_CreateRecordsetFromFile(sFileName As String, sDelimiter As String) As ADODB.Recordset
    Set ICHCSII_DAS_CreateRecordsetFromFile = CreateRecordsetFromFile(sFileName, sDelimiter)
End Function
'- GEMS
Private Function ICHCSII_DAS_ExecuteSQL(SQL As String) As Variant

End Function

Private Function ICHCSII_DAS_GetNextID(Optional ByVal eSeqName As ICHCSII_DataAccess.eSeq = 0&) As Variant

End Function

Private Function ICHCSII_DAS_GetNextNcid() As Variant

End Function

'- GEMS
Private Function ICHCSII_DAS_OpenRecordset(SQL As String, _
                            Optional ByVal CursorType As ADODB.CursorTypeEnum = 0&, _
                            Optional ByVal LockType As ADODB.LockTypeEnum = 1&) As ADODB.Recordset
    
'    Set ICHCSII_DAS_OpenRecordset = CHCSII_DAS.OpenRecordset(SQL)
    
End Function
'- GEMS
Private Property Let ICHCSII_DAS_UserNCID(RHS As String): End Property
Public Function ICHCSII_DAS_BeginTrans() As Boolean:      Err.Raise 1001, "CHCSII_DAS", "Transaction control is not supported in CHCSII Data Access Service.": End Function
Public Function ICHCSII_DAS_CommitTrans() As Boolean:     Err.Raise 1001, "CHCSII_DAS", "Transaction control is not supported in CHCSII Data Access Service.": End Function
Public Function ICHCSII_DAS_RollbackTrans() As Boolean:   Err.Raise 1001, "CHCSII_DAS", "Transaction control is not supported in CHCSII Data Access Service.": End Function

Private Function ICHCSII_DAS_OpenRecordsetSP(Parms As CHCSII_Command.CHCSII_Cmd) As ADODB.Recordset
    Set ICHCSII_DAS_OpenRecordsetSP = OpenRecordsetSP(Parms)
End Function
Private Function ICHCSII_DAS_OpenRecordsetMSP(Parms As Collection) As Collection
    Set ICHCSII_DAS_OpenRecordsetMSP = OpenRecordsetMSP(Parms)
End Function
Private Function ICHCSII_DAS_ExecuteSP(Parms As CHCSII_Command.CHCSII_Cmd) As Variant
    Call ExecuteSP(Parms)
End Function

Private Function getColumnNames(oColumnNames As SQLSERVATLLib.SqlServColumnNames) As String

    Dim iLoop As Integer
    Dim sColumnNames As String
    
    sColumnNames = ""
    With oColumnNames
    For iLoop = 1 To .Count
        If (iLoop = 1) Then
            sColumnNames = .Item(iLoop)
        Else
            sColumnNames = sColumnNames & " | " & .Item(iLoop)
        End If
    Next
    End With
    
    getColumnNames = sColumnNames
    
End Function

Private Function getValue(sqlValue As SQLSERVATLLib.SqlServValue) As Variant

Dim sInfoType As String
Dim sStringValue As String

    sInfoType = sqlValue.InfoType
    Select Case sInfoType
    
        Case "textString"
            sStringValue = Trim(sqlValue.TextString)
            If Not (sStringValue = "") Then
                getValue = CStr(sStringValue)
            End If
        Case "binaryString"
            getValue = sqlValue.BinaryString
        Case "date"
' Begin:SCR #28638;   Developer: Jane Shen 10/29/2002 06:01 PM
'            If sqlValue.Date <> #12/31/9999# Then
                getValue = AdjustDate(sqlValue.DateNoTZConversion)
'            Else
'                getValue = sqlValue.Date
'            End If
' End: SCR #28638;
            
    End Select

End Function


Public Function OpenRecordsetMSP(mSPCols As Collection) As Collection

    Dim Params As CHCSII_Command.CHCSII_Cmd
    Dim mRS As Collection
    Dim rs As ADODB.Recordset
    Dim i As Integer
    
    Call ExecutMultiSP(mSPCols)
    
    Set mRS = New Collection
 
    For i = 1 To oSqlConnect.Transaction.Count
        Set rs = ProcessRecordSet(oSqlConnect.Transaction(i))
        mRS.Add rs
    Next
    
End Function
