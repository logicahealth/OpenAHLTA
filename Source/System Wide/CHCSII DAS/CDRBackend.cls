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
Attribute VB_Name = "CDRBackend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This class is used to access data from/ to the CDR via the DSql control
'provided by 3M.

Option Explicit

Implements iBackend

Private m_objDSQL As DSQLATLLib.DSqlConnect

Private Const SQLQueryNCID As String = "1450582"
Private Const OraErrorNoDataFound As String = "ORA-00100"
Private Const OraErrorFetchOutOfSequence As String = "ORA-01002"
Private Const cnst_MAX_NUM_OF_RECORDS = 5000 '* Add by j.ray 31AUG2001 to fix SCR #16462

Private TZInfo As TIME_ZONE_INFORMATION

Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

Private bSkipFacilityWorkAround As Boolean     '- so mid-tier can function...

Private Sub Class_Initialize()
  Set m_objDSQL = New DSQLATLLib.DSqlConnect
  
  LoadRelevantWorkaroundTables
  GetTimeZoneInformation TZInfo
End Sub

Private Sub Class_Terminate()
  Set m_objDSQL = Nothing
End Sub

Private Function iBackend_CreateRecordsetFromFile(ByVal sFileName As String, ByVal sDelimiter As String) As ADODB.Recordset
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

Set iBackend_CreateRecordsetFromFile = objRecordset

End Function

Private Sub iBackend_ExecuteSQL(ByVal SQL As String)
Dim sOraError As String

On Error GoTo CheckErr

With m_objDSQL
    .ClearAll
    .Table.sqlId.ValueConcept.UniqueId = SQLQueryNCID
    .Table.arguments.Add
    .Table.arguments(1).Text = SQL
    .Query

CheckErr:
    If Len(.Table.Error.ErrorText) <> 0 Then
        On Error Resume Next
        sOraError = Left(.Table.Error.ErrorText, 9)
        Select Case sOraError
            Case OraErrorFetchOutOfSequence
                ' currently, the DSQL stuff always attempts a fetch after executing an SQL statement
                ' for non-select statements (insert, update, delete), this produces a 'fetch out of sequence' error
                ' so, instead of propagating this error, we do a commit and keep going
            Case Else
                Err.Raise vbObjectError + 1, "Error with Statement: (" & SQL & ")", "Error Text = " & .Table.Error.ErrorText
                Exit Sub
        End Select
    End If
End With

End Sub

Private Function iBackend_OpenRecordset(ByVal SQL As String) As ADODB.Recordset
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
Dim bReverseLast As Boolean

sNewSQL = DoFacilityNCIDWorkaround(SQL, bSkipFacilityWorkAround)

With m_objDSQL
    .ClearAll
    .Table.sqlId.ValueConcept.UniqueId = SQLQueryNCID
    .Table.arguments.Add
    .Table.arguments(1).Text = sNewSQL
    .Query
    If Len(.Table.Error.ErrorText) <> 0 Then
        sOraError = Left(.Table.Error.ErrorText, 9)
        Select Case sOraError
            Case OraErrorNoDataFound
                Set iBackend_OpenRecordset = objRecordset
                Exit Function
            Case Else
                Err.Raise vbObjectError + 1, "Error with Statement: (" & SQL & ")", "Error Text = " & .Table.Error.ErrorText
                Exit Function
        End Select
    End If
    
    'Construct an ADO Recordset with it
    iNumRows = .Table.Rows.Count
    
    bReverseLast = False
    'If (iNumRows > 1 And iNumRows < 1000) Then bReverseLast = True
    'If (iNumRows < 1000) Then bReverseLast = True
    bReverseLast = True
    
    '* Add by j.ray 31AUG2001 to fix SCR #16462
    If (iNumRows >= cnst_MAX_NUM_OF_RECORDS) Then bReverseLast = False
    
    
    If iNumRows > 0 Then
        'Get columns first
        iNumCols = .Table.Rows(1).Columns.Count
        If iNumRows = 1 Then
            'Need to reverse columns
            For i = iNumCols To 1 Step -1
                sFieldName = .Table.Rows(1).Columns(i).Name
                sInfoType = .Table.Rows(1).Columns(i).Value.InfoType
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
                sFieldName = .Table.Rows(1).Columns(i).Name
                sInfoType = .Table.Rows(1).Columns(i).Value.InfoType
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
            If (bReverseLast And i = iNumRows) Then
                For j = 1 To iNumCols
                    sInfoType = .Table.Rows(i).Columns(iNumCols - (j - 1)).Value.InfoType
                    Select Case sInfoType
                        Case "textString"
                            sStringValue = .Table.Rows(i).Columns(iNumCols - (j - 1)).Value.TextString
                            If Not (sStringValue = "") Then
                                objRecordset.Fields.Item(j - 1).Value = CStr(sStringValue)
                            End If
                        Case "binaryString"
                            objRecordset.Fields.Item(j - 1).Value = .Table.Rows(i).Columns(iNumCols - (j - 1)).Value.BinaryString
                        Case "date"
                            If .Table.Rows(i).Columns(iNumCols - (j - 1)).Value.Date <> #12/31/9999# Then
                                ' The DSQL atl component is spuriously adjusting for local time,
                                ' so we "undo" that adjustment here
                                objRecordset.Fields.Item(j - 1).Value = AdjustDate(.Table.Rows(i).Columns(iNumCols - (j - 1)).Value.Date)
                            Else
                                objRecordset.Fields.Item(j - 1).Value = .Table.Rows(i).Columns(iNumCols - (j - 1)).Value.Date
                            End If
                    End Select
                Next
            Else
                For j = 1 To iNumCols
                    sInfoType = .Table.Rows(i).Columns(j).Value.InfoType
                    Select Case sInfoType
                        Case "textString"
                            sStringValue = .Table.Rows(i).Columns(j).Value.TextString
                            If Not (sStringValue = "") Then
                                objRecordset.Fields.Item(j - 1).Value = CStr(sStringValue)
                            End If
                        Case "binaryString"
                            objRecordset.Fields.Item(j - 1).Value = .Table.Rows(i).Columns(j).Value.BinaryString
                        Case "date"
                            If .Table.Rows(i).Columns(j).Value.Date <> #12/31/9999# Then
                                ' The DSQL atl component is spuriously adjusting for local time,
                                ' so we "undo" that adjustment here
                                objRecordset.Fields.Item(j - 1).Value = AdjustDate(.Table.Rows(i).Columns(j).Value.Date)
                            Else
                                objRecordset.Fields.Item(j - 1).Value = .Table.Rows(i).Columns(j).Value.Date
                            End If
                    End Select
                Next
            End If
        Next
        objRecordset.MoveFirst
    Else
        'Create dummy field and open recordset
        objRecordset.Fields.Append "No Records Found", adBSTR
        objRecordset.Open
    End If
End With

Set iBackend_OpenRecordset = objRecordset

End Function

Private Property Let iBackend_SkipFacilityWorkAround(ByVal RHS As Boolean)
  bSkipFacilityWorkAround = RHS
End Property

Private Property Get iBackend_SkipFacilityWorkAround() As Boolean
  iBackend_SkipFacilityWorkAround = bSkipFacilityWorkAround
End Property

Private Function AdjustDate(ByVal dDate As Date) As Date
Dim lYear As Integer
Dim dStandard As Date
Dim dDaylight As Date
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

