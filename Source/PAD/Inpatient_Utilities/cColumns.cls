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
Attribute VB_Name = "cColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mcol        As Collection
Private mAutoResize        As Boolean
Private mSorted            As Boolean   ' is this thing sorted?
Private mSortCol           As Long      ' from the 'col' in grid.aftersort method
Private mSortOrder         As Integer   ' from the 'order' in grid.aftersort method



Public Property Get Columns() As Collection
    Set Columns = mcol
End Property

Public Function AddColumn(oCol As cColumn) As Boolean
On Error GoTo ErrHandler
    mcol.Add oCol, oCol.ColID
    AddColumn = True
    Exit Function
ErrHandler:
    AddColumn = False
End Function

Public Function Add(ByVal ColName As String, _
                    ByVal colHeader As String, _
                    ByVal colNumber As Long, _
                    ByVal colOrdinal As Long, _
                    ByVal ObjectName As String, _
                    ByVal RSName As String, _
                    ByVal colWidth As Long) As Boolean
                    
On Error GoTo ErrHandler
    Add = False
    
    Dim oCol As New cColumn
    With oCol
        .ColID = mcol.Count + 1
        .ColName = ColName
        .ColumnHeader = colHeader
        .ColumnNumber = colNumber
        .ColumnOrdinal = colOrdinal
        .columnWidth = colWidth
        .ObjectName = ObjectName
        .RSName = RSName
    End With
    mcol.Add oCol, oCol.ColName
    Add = True
    Exit Function
ErrHandler:

End Function
Public Property Get Item(vntIndexKey As Variant) As cColumn
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



Public Sub CaptureColumnOrder(ByVal OldPosn As Long, ByVal NewPosn As Long)

' called by grid action with the old and new positions
'  we have to change the columnordinal from the old to the new
'  and change everything inbetween up one

'  0 1 2 3 4 5 6 7 8 9
'  with old =7 and new =3 then
'  0 1 2 7 3 4 5 6 8 9
    '__ so, get the lowest of the two
    Dim oCol As cColumn
    
    On Error GoTo ErrHandler
    
    For Each oCol In mcol
        With oCol
            If .ColumnOrdinal = OldPosn Then
               .ColumnOrdinal = NewPosn
            ElseIf NewPosn < OldPosn Then
                If .ColumnOrdinal >= NewPosn And .ColumnOrdinal < OldPosn Then
                    .ColumnOrdinal = .ColumnOrdinal + 1
                End If
            ElseIf NewPosn > OldPosn Then
                If .ColumnOrdinal <= NewPosn And .ColumnOrdinal > OldPosn Then
                    .ColumnOrdinal = .ColumnOrdinal - 1
                End If
               
            End If
        End With
    Next oCol
    
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "cColumns.CaptureColumnOrder", Err.Description, Err.HelpFile, Err.HelpContext
End Sub


Public Sub CaptureColumnWidth(ByVal lCol As Long, ByVal lWidth As Long)
    Dim oCol As cColumn
    
    On Error GoTo ErrHandler
    
    For Each oCol In mcol
        If oCol.ColumnOrdinal = lCol Then
            oCol.columnWidth = lWidth
            Exit For
        End If
    Next oCol
    
    Set oCol = Nothing
    
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "cColumns.CaptureColumnOrder", Err.Description, Err.HelpFile, Err.HelpContext
End Sub


Public Property Get Sorted() As Boolean
    Sorted = mSorted
End Property

Public Property Let Sorted(ByVal NewValue As Boolean)
    mSorted = NewValue
End Property

Public Property Get SortCol() As Long
    SortCol = mSortCol
End Property

Public Property Let SortCol(ByVal NewValue As Long)
    mSortCol = NewValue
End Property

Public Property Get SortOrder() As Integer
    SortOrder = mSortOrder
End Property

Public Property Let SortOrder(ByVal NewValue As Integer)
    mSortOrder = NewValue
End Property

Private Function ColumnOrderString(ByVal bDefaultSort As Boolean, _
                                   ByVal lDefaultColumnOrdinal As Long, _
                                   ByVal iDefaultSortOrder As Integer) As String
    
    Dim oCol As Inpatient_Utilities.cColumn
    Dim i As Long
    
    On Error GoTo ErrHandler
    
    '--- The default sort order is by Specialty.
    Sorted = bDefaultSort
    SortOrder = iDefaultSortOrder
    ColumnOrderString = Sorted & "|" & lDefaultColumnOrdinal & "|" & SortOrder & "^"
    
    
    For i = 0 To mcol.Count - 1
        For Each oCol In mcol
            If oCol.ColumnOrdinal = i Then
                ColumnOrderString = ColumnOrderString & oCol.ColName & "|"
            End If
        Next oCol
    Next i
'__ drop trailing pipe
    If Right$(ColumnOrderString, 1) = "|" Then
        ColumnOrderString = Left$(ColumnOrderString, Len(ColumnOrderString) - 1) & "^"
    End If
    
'---Column Width
    For i = 0 To mcol.Count - 1
        For Each oCol In mcol
            If oCol.ColumnOrdinal = i Then
                ColumnOrderString = ColumnOrderString & oCol.ColName & "=" & CStr(oCol.columnWidth) & "|"
            End If
        Next oCol
    Next i
    
'__ drop trailing pipe
    If Right$(ColumnOrderString, 1) = "|" Then
        ColumnOrderString = Left$(ColumnOrderString, Len(ColumnOrderString) - 1)
    End If
    
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "cColumns.ColumnOrderString", Err.Description, Err.HelpFile, Err.HelpContext

End Function


Public Function Save(ByVal sAppNcid As String, ByVal sUserNCID As String, _
                    ByVal sCofigName, ByVal bDefaultSort As Boolean, _
                    ByVal iDefaultSortOrder As Integer, _
                    ByVal lDefaultColumnOrdinal As Long) As Boolean

    Dim oConfig As MMMHISConfiguration.Config
    Dim sOshString As String
    
    On Error GoTo ErrHandler
    
    Save = False
    
    Set oConfig = New MMMHISConfiguration.Config
    
    sOshString = ColumnOrderString(bDefaultSort, lDefaultColumnOrdinal, iDefaultSortOrder)
    oConfig.SetValue sAppNcid, sUserNCID, sCofigName, "columns", sOshString
    oConfig.SaveChanges sAppNcid, sUserNCID
   
    Set oConfig = Nothing
    
    Save = True
    
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "cColumns.Save", Err.Description, Err.HelpFile, Err.HelpContext
End Function


Public Function RetrieveDefaultInfo(ByVal sAppNcid As String, _
                                    ByVal sCofigName) As Boolean

    Dim oConfig As MMMHISConfiguration.Config
    Dim rsColumnInfo As String
    Dim columnOrdinals() As String
    Dim columnSort() As String
    Dim columnInfo() As String
    Dim columnWidth() As String
    Dim i As Long
    Dim oCol As cColumn
    
    On Error GoTo ErrHandler
    
    RetrieveDefaultInfo = False
    Set oConfig = New MMMHISConfiguration.Config
    
    If Not oConfig.GetValue(sAppNcid, sCofigName, "columns", rsColumnInfo) Then
        '-- no op
    End If
    If rsColumnInfo = vbNullString Then Exit Function
    
    columnInfo = Split(rsColumnInfo, "^")
    If UBound(columnInfo) = 2 Then
        columnSort = Split(columnInfo(0), "|")
        columnOrdinals = Split(columnInfo(1), "|")
        columnWidth = Split(columnInfo(2), "|")
        AutoResize = False
    ElseIf UBound(columnInfo) = 1 Then
        columnSort = Split(columnInfo(0), "|")
        columnOrdinals = Split(columnInfo(1), "|")
        AutoResize = True
    Else
        columnOrdinals = Split(columnInfo(0), "|")
        AutoResize = True
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
        oCol.ColumnOrdinal = 99
    Next oCol
    On Error Resume Next   '-- if we changed the columns names since this guy logged in then we go with defaults
    For i = 0 To UBound(columnOrdinals)
        Set oCol = mcol.Item(columnOrdinals(i))
        If Err.Number <> 0 Then
            Err.Clear
            RetrieveDefaultInfo = False
            Exit Function
        End If
        oCol.ColumnOrdinal = i
    Next i
    
    '__ now see that all the columns were reset
    For Each oCol In mcol
        If oCol.ColumnOrdinal = 99 Then
            RetrieveDefaultInfo = False
            Exit Function
        End If
    Next oCol
    
'--- set the column width
    For i = 0 To UBound(columnWidth)
        Set oCol = mcol.Item(Split(columnWidth(i), "=")(0))
        oCol.columnWidth = Split(columnWidth(i), "=")(1)
    Next i
    
    
'__ set the sort order
    If UBound(columnSort) > 0 Then
        Sorted = (columnSort(0) = "True")
        SortCol = columnSort(1)
        SortOrder = columnSort(2)
    End If
   
    Set oConfig = Nothing
    RetrieveDefaultInfo = True
    
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "cColumns.RetrieveDefaultInfo", Err.Description, Err.HelpFile, Err.HelpContext
End Function



Public Property Get AutoResize() As Boolean
    AutoResize = mAutoResize
End Property

Public Property Let AutoResize(ByVal NewValue As Boolean)
    mAutoResize = NewValue
End Property
