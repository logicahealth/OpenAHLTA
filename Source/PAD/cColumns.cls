VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"cColumn"
Attribute VB_Ext_KEY = "Member0" ,"cColumn"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mcol        As Collection
Public clsInpatient As Inpatient

Public Sorted       As Boolean   ' is this thing sorted?
Public SortCol      As Long      ' from the 'col' in grid.aftersort method
Public SortOrder    As Integer   ' from the 'order' in grid.aftersort method

Public Property Get Columns() As Collection
    Set Columns = mcol
End Property
Public Property Get ColumnOrderString() As String
    
    Dim oCol As cColumn
    Dim i As Long
    ColumnOrderString = Sorted & "|" & SortCol & "|" & SortOrder & "^"
    For i = 0 To mcol.Count - 1
        For Each oCol In mcol
            If oCol.columnOrdinal = i Then
                ColumnOrderString = ColumnOrderString & oCol.colName & "|"
            End If
        Next oCol
    Next i
'__ drop trailing pipe
    If Right$(ColumnOrderString, 1) = "|" Then
        ColumnOrderString = Left$(ColumnOrderString, Len(ColumnOrderString) - 1)
    End If
End Property

Public Function AddColumn(oCol As cColumn) As Boolean
On Error GoTo errHandler
    mcol.Add oCol, oCol.ColID
    AddColumn = True
    Exit Function
errHandler:
    AddColumn = False
End Function

Public Function Add(ByVal colName As String, _
                    ByVal colHeader As String, _
                    ByVal colNumber As Long, _
                    ByVal colOrdinal As Long, _
                    ByVal ObjectName As String, _
                    ByVal RSName As String) As cColumn
                    
On Error GoTo errHandler
    Dim oCol As New cColumn
    With oCol
        .ColID = mcol.Count + 1
        .colName = colName
        .ColumnHeader = colHeader
        .ColumnNumber = colNumber
        .columnOrdinal = colOrdinal
        .ObjectName = ObjectName
        .RSName = RSName
    End With
    mcol.Add oCol, oCol.colName
    Set Add = oCol
    Exit Function
errHandler:

End Function
Public Property Get Item(vntIndexKey As Variant) As cColumn
Attribute Item.VB_UserMemId = 0
  Set Item = mcol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mcol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mcol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mcol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    Set mcol = New Collection
End Sub


Private Sub Class_Terminate()
    Set mcol = Nothing
End Sub

Friend Function LoadcolumnInfo() As Boolean
    Dim rsColumnInfo As String
    Dim rsColumnOrder As String
    Dim columnOrdinals() As String
    Dim columnSort() As String
    Dim columnInfo() As String
    
    Dim i As Long
    Dim oCol As cColumn
    
'__ create the collection of ocol objects
    Call ResetColumns
'-- get the conlumn information from the osh and load the collection of column objects

'--  col info is in the form of:   sorted (T/F) | SortCol | SortOrder ^ col name | col name | ...

    If Not gobjconfig.GetValue(NCID_INPATIENT_APPLICATION, "default settings", "columns", rsColumnInfo) Then
        '-- no op
    End If
    If rsColumnInfo = vbNullString Then Exit Function
    
    columnInfo = Split(rsColumnInfo, "^")
    If UBound(columnInfo) = 1 Then
        columnSort = Split(columnInfo(0), "|")
        columnOrdinals = Split(columnInfo(1), "|")
    Else
        columnOrdinals = Split(columnInfo(0), "|")
    End If
    '__ if the col count is zero then they have not saved anything and we go with the default setup above
    'validate the ordinals from the osh
    '__ do we have enough?
    If UBound(columnOrdinals) <> mcol.Count - 1 Then
        '-- use the loaded defaults
        Exit Function
    End If
    
    
    '__ mark the column ordinals with a flag
    For Each oCol In mcol
        oCol.columnOrdinal = 99
    Next oCol
    On Error Resume Next   '-- if we changed the columns names since this guy logged in then we go with defaults
    For i = 0 To UBound(columnOrdinals)
        Set oCol = clsInpatient.Columns.Item(columnOrdinals(i))
        If Err.Number <> 0 Then
            Err.Clear
            Call ResetColumns
            Exit Function
        End If
        oCol.columnOrdinal = i
    Next i
    
    '__ now see that all the columns were reset
    For Each oCol In mcol
        If oCol.columnOrdinal = 99 Then
            Call ResetColumns
            Exit Function
        End If
    Next oCol
    
    
'__ set the sort order
    If UBound(columnInfo) = 1 Then
        Sorted = (columnSort(0) = "True")
        SortCol = columnSort(1)
        SortOrder = columnSort(2)
    End If
    Exit Function

End Function

Friend Function ResetColumns() As Boolean
    Set mcol = New Collection
    Call Add("RegNum", "Register Number", 0, 0, "ID", "Inpatient_ID")
    Call Add("Specialty", "Specialty", 1, 1, "Specialty", "Specialty")
    Call Add("Capacity", "% Capacity", 2, 2, "Capacity", "Capacity")
    Call Add("Bed", "Bed", 3, 3, "Bed", "Bed")
    Call Add("Rank", "Rank", 4, 4, "Rank", "Rank")
    Call Add("Patient", "Patient", 5, 5, "PatientName", "Patient")
    Call Add("FMPSSN", "FMP/Sponsor SSN", 6, 6, "FMPSSN", "FMPSSN")
    Call Add("Unit", "Unit", 7, 7, "Unit", "Unit")
    Call Add("PatService", "Inpatient Service", 8, 8, "Specialty", "PRIMARY_SERVICE_NCID")
    Call Add("AdmittingDiagnosis", "Admitting Diagnosis", 9, 9, "AdmittingDiagnosis", "AdmittingDiagnosis")
    Call Add("AdmissionDate", "Admission Date", 10, 10, "AdmissionDate", "AdmissionDate")
    Call Add("AdmissionComments", "Admission Comments", 11, 11, "AdmissionComments", "AdmissionComments")
    Call Add("Day", "Day #", 12, 12, "Day", "Day")
    Call Add("Location", "Location", 13, 13, "Location", "Location")
    Call Add("Status", "Status", 14, 14, "Status", "Status")
    Call Add("DischargeDate", "Discharge Date", 15, 15, "DischargeDate", "DischargeDate")
    Call Add("Disposition", "Disposition", 16, 16, "Disposition", "Disposition")
    Call Add("DischargeDiagnosis", "Discharge Diagnosis", 17, 17, "DischargeDiagnosis", "DischargeDiagnosis")
    Call Add("AdmittingProvider", "Admitting Provider", 18, 18, "AdmittingProviderName", "AdmittingProvider")
    Call Add("AttendingProvider", "Provider", 19, 19, "AttendingProvider", "AttendingProvider")
    Call Add("TransferDate", "Transfer Date", 20, 20, "LastTransferDate", "TransferDate")
    Call Add("Service", "Service", 21, 21, "Service", "Service")

End Function


Public Sub CaptureColumnOrder(ByVal OldPosn As Long, ByVal NewPosn As Long)

' called by grid action with the old and new positions
'  we have to change the columnordinal from the old to the new
'  and change everything inbetween up one

'  0 1 2 3 4 5 6 7 8 9
'  with old =7 and new =3 then
'  0 1 2 7 3 4 5 6 8 9
    '__ so, get the lowest of the two
    Dim oCol As cColumn
    For Each oCol In mcol
        With oCol
            If .columnOrdinal = OldPosn Then
               .columnOrdinal = NewPosn
            ElseIf NewPosn < OldPosn Then
                If .columnOrdinal >= NewPosn And .columnOrdinal < OldPosn Then
                    .columnOrdinal = .columnOrdinal + 1
                End If
            ElseIf NewPosn > OldPosn Then
                If .columnOrdinal <= NewPosn And .columnOrdinal > OldPosn Then
                    .columnOrdinal = .columnOrdinal - 1
                End If
               
            End If
        End With
    Next oCol
    
End Sub
