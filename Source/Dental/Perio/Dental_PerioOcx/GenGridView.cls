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
Attribute VB_Name = "GenGridView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit



Private WithEvents mctlGrid As VSFlexGrid
Attribute mctlGrid.VB_VarHelpID = -1
Private menumViewType As eViewType
Private mlngToothNameRow As Long
Private mlngMissingRow As Long
Private mlngCARow As Long
Private mlngGMRow As Long
Private mlngProbDepthRow As Long
Private mlngBleedingRow As Long
Private mlngPrognosisRow As Long
Private mlngMobilityRow As Long
Private mlngFurcationRow As Long
Private mlngBoneLossRow As Long
Private mlngSuppurationRow As Long
Private mlngPlaqueRow As Long
Private mlngMgjRow As Long
Private menumGridType As eGridType

Public Event AdvanceCellsFail(Row As Long, Col As Long, ToRow As Long, ToCol As Long)
Public Event ValueSet(Measurement As eMeasurementType, Tooth As Integer, Location As eMeasurementLocation, Value As Variant)
Public Event TogglePrevValue(Measurement As eMeasurementType, Tooth As Integer, Location As eMeasurementLocation, Value As Variant)
Private Const DISABLED_COLOR = &HC0C0C0
Private Const ENABLED_COLOR = &H80000005

Private mblnIsLoading As Boolean

Public Property Get IsLoading() As Boolean
   IsLoading = mblnIsLoading
End Property
Public Property Let IsLoading(blnLoading As Boolean)
   mblnIsLoading = blnLoading
End Property


Public Function CellIsValid(Row As Long, Col As Long) As Boolean
Dim lngDisabledColor As Long
    With mctlGrid
        If .Cell(flexcpBackColor, Row, Col) <> DISABLED_COLOR Then
            CellIsValid = True
        Else
            CellIsValid = False
        End If
    End With
End Function
Public Function CurrentCellIsValid() As Boolean
    CurrentCellIsValid = CellIsValid(mctlGrid.Row, mctlGrid.Col)
End Function


Public Property Get DisabledColor() As Long
    DisabledColor = DISABLED_COLOR
End Property
Public Property Get EnabledColor() As Long
    EnabledColor = ENABLED_COLOR
End Property

Public Property Get GridType() As eGridType
    GridType = menumGridType
End Property
Public Property Let GridType(enumGridType As eGridType)
    menumGridType = enumGridType
End Property

Friend Property Set Grid(ctlGrid As VSFlexGrid)
    Set mctlGrid = ctlGrid
End Property
Public Property Let ViewType(enumView As eViewType)
    menumViewType = enumView
End Property
Public Property Get ToothNameRow() As Long
    ToothNameRow = mlngToothNameRow
End Property
Public Property Get MissingRow() As Long
    MissingRow = mlngMissingRow
End Property
Public Property Get CARow() As Long
    CARow = mlngCARow
End Property
Public Property Get GMRow() As Long
    GMRow = mlngGMRow
End Property
Public Property Get ProbDepthRow() As Long
    ProbDepthRow = mlngProbDepthRow
End Property
Public Property Get BleedingRow() As Long
    BleedingRow = mlngBleedingRow
End Property
Public Property Get PrognosisRow() As Long
    PrognosisRow = mlngPrognosisRow
End Property
Public Property Get MobilityRow() As Long
    MobilityRow = mlngMobilityRow
End Property
Public Property Get FurcationRow() As Long
    FurcationRow = mlngFurcationRow
End Property
Public Property Get BoneLossRow() As Long
    BoneLossRow = mlngBoneLossRow
End Property
Public Property Get SuppurationRow() As Long
    SuppurationRow = mlngSuppurationRow
End Property
Public Property Get PlaqueRow() As Long
    PlaqueRow = mlngPlaqueRow
End Property
Public Property Get MgjRow() As Long
    MgjRow = mlngMgjRow
End Property

Public Function AddToothNameRow() As Long
    AddToothNameRow = Add1ValuePerToothRow(mctlGrid, "Tooth", "0")
    mlngToothNameRow = AddToothNameRow
    

    With mctlGrid
        'Center the labels for this row
        .Cell(flexcpAlignment, mlngToothNameRow, 1, mlngToothNameRow, .COLS - 1) = flexAlignCenterCenter
    End With
End Function
Public Function AddMissingRow() As Long
    AddMissingRow = Add1ValuePerToothRow(mctlGrid, "IsMissing", "0")
    mlngMissingRow = AddMissingRow
    mctlGrid.RowHidden(AddMissingRow) = True
End Function
Public Function AddCARow() As Long
    AddCARow = Add3ValuePerToothRow(mctlGrid, "CA", "0")
    mlngCARow = AddCARow
End Function
Public Function AddGMRow()
    AddGMRow = Add3ValuePerToothRow(mctlGrid, "GM", "0")
    mlngGMRow = AddGMRow
End Function
Public Function AddProbDepthRow()
    AddProbDepthRow = Add3ValuePerToothRow(mctlGrid, "Probing Depth")
    mlngProbDepthRow = AddProbDepthRow
End Function
Public Function AddPlaqueRow()
    AddPlaqueRow = Add3ValuePerToothRow(mctlGrid, "Plaque")
        mlngPlaqueRow = AddPlaqueRow
End Function
Public Function AddSuppurationRow()
    AddSuppurationRow = Add3ValuePerToothRow(mctlGrid, "Suppuration")
    mlngSuppurationRow = AddSuppurationRow
End Function
Public Function AddBleedingRow()
    AddBleedingRow = Add3ValuePerToothRow(mctlGrid, "Bleeding")
    mlngBleedingRow = AddBleedingRow
End Function
Public Function AddMobilityRow()
    AddMobilityRow = Add1ValuePerToothRow(mctlGrid, "Mobility")
    mlngMobilityRow = AddMobilityRow
End Function
Public Function AddPrognosisRow()
    AddPrognosisRow = Add1ValuePerToothRow(mctlGrid, "Prognosis")
    mlngPrognosisRow = AddPrognosisRow
End Function
Public Function AddBoneLossRow() As Long
    AddBoneLossRow = Add2ValuePerToothRow(mctlGrid, "Bone Loss", "M", "D")
    mlngBoneLossRow = AddBoneLossRow
End Function
Public Function AddMgjRow()
Dim intCurrTooth As Integer
Dim i As Integer
Dim lngNewRow As Long

    Select Case menumGridType
        Case perBuccal 'all teeth have a buccal value
            lngNewRow = Add1ValuePerToothRow(mctlGrid, "MGJ")
        Case perLingual 'mandibular teeth have lingual value
            lngNewRow = Add1ValuePerToothRow(mctlGrid, "MGJ")
    
        With mctlGrid
            For i = 1 To .COLS Step 4
                intCurrTooth = CInt(.TextMatrix(0, i))
                Select Case intCurrTooth
                    Case Is <= 16 'maxillary tooth
                        '.Cell(flexcpBackColor, lngNewRow, i) = DISABLED_COLOR
                        '.Cell(flexcpBackColor, lngNewRow, i + 1) = DISABLED_COLOR
                        '.Cell(flexcpBackColor, lngNewRow, i + 2) = DISABLED_COLOR
                        .Cell(flexcpBackColor, lngNewRow, i, lngNewRow, i + 2) = DISABLED_COLOR
                    Case Is >= 17 'mandibular tooth
                End Select
            Next i
        End With
    End Select
    AddMgjRow = lngNewRow
    mlngMgjRow = lngNewRow
End Function

Private Sub FormatFurcation(intStartCol As Integer)
Dim strToothName As String
Dim r1 As Long
Dim r2 As Long
Dim c1 As Long
Dim c2 As Long


    With mctlGrid
        'Try to use the tooth name to set furcation values, but if it
        'isn't set, use the tooth position
        strToothName = Trim(.TextMatrix(mlngToothNameRow, intStartCol))
        If strToothName = "0" Or Len(strToothName) = 0 Then
            strToothName = Trim(.TextMatrix(0, intStartCol))
        End If
        
        If Len(Trim(.TextMatrix(0, intStartCol))) <> 0 Then
            Select Case strToothName
                Case "1", "2", "3", "14", "15", "16", "A", "B", "I", "J"  '3 roots, 3 columns
                    '.TextMatrix(mlngFurcationRow, intStartCol) = " "
                    '.TextMatrix(mlngFurcationRow + 1, intStartCol) = " "
                    '.TextMatrix(mlngFurcationRow + 2, intStartCol) = " "
                    .Cell(flexcpText, mlngFurcationRow, intStartCol, mlngFurcationRow + 2, intStartCol) = " "
                    .Cell(flexcpBackColor, mlngFurcationRow, intStartCol, mlngFurcationRow + 2, intStartCol) = ENABLED_COLOR
                    
                Case "5", "12"  '2 roots, 2 columns
        
                    '.TextMatrix(mlngFurcationRow, intStartCol) = " "
                    '.TextMatrix(mlngFurcationRow + 1, intStartCol) = " "
                    '.TextMatrix(mlngFurcationRow + 2, intStartCol) = " "
                    .Cell(flexcpText, mlngFurcationRow, intStartCol, mlngFurcationRow + 2, intStartCol) = " "
                    If intStartCol Mod 4 = 2 Then
                        .Cell(flexcpBackColor, mlngFurcationRow, intStartCol) = DISABLED_COLOR
                    End If
                
                Case "17", "18", "19", "30", "31", "32", "K", "L", "S", "T" '2 roots, 2 rows
                    .TextMatrix(mlngFurcationRow, intStartCol) = "B"
                    .TextMatrix(mlngFurcationRow + 1, intStartCol) = ""
                    .TextMatrix(mlngFurcationRow + 2, intStartCol) = "L"
                    .Cell(flexcpBackColor, mlngFurcationRow, intStartCol, mlngFurcationRow + 2, intStartCol) = ENABLED_COLOR
                    
                Case "4", "6", "7", "8", "9", "10", "11", "13", "20", "21", _
                        "22", "23", "24", "25", "26", "27", "28", "29", _
                        "C", "D", "E", "F", "G", "H", "M", "N", "O", "P", "Q", "R" '1 root
                    '.TextMatrix(mlngFurcationRow, intStartCol) = " "
                    '.TextMatrix(mlngFurcationRow + 1, intStartCol) = " "
                    '.TextMatrix(mlngFurcationRow + 2, intStartCol) = " "
                    .Cell(flexcpText, mlngFurcationRow, intStartCol, mlngFurcationRow + 2, intStartCol) = " "
                    .Cell(flexcpBackColor, mlngFurcationRow, intStartCol, mlngFurcationRow + 2, intStartCol) = DISABLED_COLOR
                Case Else
                    Debug.Assert False
            End Select
            
            If intStartCol Mod 4 = 3 Then
             .GetSelection r1, c1, r2, c2
             .Select mlngFurcationRow, intStartCol, mlngFurcationRow + 2, intStartCol
             .CellBorder RGB(0, 0, 0), -1, -1, 1, -1, -1, -1
             .Select r1, c1, r2, c2
            End If
            
            
        End If
    End With
End Sub


Public Function AddFurcationRow()
Dim lngNewRow As Long
Dim i As Long
Dim eRoots As eRootType

    With mctlGrid
        lngNewRow = .ROWS
        mlngFurcationRow = lngNewRow
        
        .ROWS = .ROWS + 4
        
        '.TextMatrix(lngNewRow, 0) = "Furcation"
        '.TextMatrix(lngNewRow + 1, 0) = "Furcation"
        '.TextMatrix(lngNewRow + 2, 0) = "Furcation"
        .Cell(flexcpText, lngNewRow, 0, lngNewRow + 2, 0) = "Furcation"
        
        .MergeRow(lngNewRow) = True
        .MergeRow(lngNewRow + 2) = True
        
        .RowHidden(lngNewRow + 1) = True
        .RowHidden(lngNewRow + 3) = True
        
        For i = 1 To .COLS - 1
            FormatFurcation CInt(i)
        Next i
        
    End With

    AddFurcationRow = lngNewRow
End Function

Private Function Add1ValuePerToothRow(ctlGrid As VSFlexGrid, strRowName As String, Optional strDefaultValue As String = " ") As Long
Dim i As Long
Dim lngNewRow As Long
Dim strDefault As String
Dim r1 As Long
Dim r2 As Long
Dim c1 As Long
Dim c2 As Long

    
    With ctlGrid
        lngNewRow = .ROWS
        .ROWS = .ROWS + 2
        .TextMatrix(lngNewRow, 0) = strRowName
        .MergeRow(lngNewRow) = True
        .RowHidden(lngNewRow + 1) = True
    
        For i = 1 To .COLS Step 4
            '.TextMatrix(lngNewRow, i) = strDefaultValue
            '.TextMatrix(lngNewRow, i + 1) = strDefaultValue
            '.TextMatrix(lngNewRow, i + 2) = strDefaultValue
            .Cell(flexcpText, lngNewRow, i, lngNewRow, i + 2) = strDefaultValue
            
            .GetSelection r1, c1, r2, c2
            .Select lngNewRow, i, lngNewRow, i + 2
            .CellBorder RGB(0, 0, 0), -1, -1, 1, -1, -1, 1
            .Select r1, c1, r2, c2
        Next i
        
    End With
    Add1ValuePerToothRow = lngNewRow

End Function

Private Function Add3ValuePerToothRow(ctlGrid As VSFlexGrid, strRowName As String, Optional strDefaultValue As String = " ") As Long
Dim i As Long
Dim lngNewRow As Long
Dim r1 As Long
Dim r2 As Long
Dim c1 As Long
Dim c2 As Long

    With ctlGrid
        lngNewRow = .ROWS
        .ROWS = .ROWS + 2
        .TextMatrix(lngNewRow, 0) = strRowName
        .RowHidden(lngNewRow + 1) = True
        
        For i = 1 To .COLS Step 4
            '.TextMatrix(lngNewRow, i) = strDefaultValue
            '.TextMatrix(lngNewRow, i + 1) = strDefaultValue
            '.TextMatrix(lngNewRow, i + 2) = strDefaultValue
            .Cell(flexcpText, lngNewRow, i, lngNewRow, i + 2) = strDefaultValue
        Next i
      
      For i = 4 To .COLS - 1 Step 4
         .GetSelection r1, c1, r2, c2
         .Select lngNewRow, i - 1, lngNewRow, i - 1
         .CellBorder RGB(0, 0, 0), -1, -1, 1, -1, -1, -1
         .Select r1, c1, r2, c2
      Next i
        
    End With
    Add3ValuePerToothRow = lngNewRow
End Function

Private Function Add2ValuePerToothRow(ctlGrid As VSFlexGrid, strRowName As String, Optional strDefaultValue1 As String = " ", Optional strdefaultvalue2 = " ") As Long

Dim lngNewRow As Long
Dim i As Long
Dim eRoots As eRootType
Dim r1 As Long
Dim r2 As Long
Dim c1 As Long
Dim c2 As Long

    With mctlGrid
        lngNewRow = .ROWS
        .ROWS = .ROWS + 4
        
        '.TextMatrix(lngNewRow, 0) = strRowName
        '.TextMatrix(lngNewRow + 1, 0) = strRowName
        '.TextMatrix(lngNewRow + 2, 0) = strRowName
        .Cell(flexcpText, lngNewRow, 0, lngNewRow + 2, 0) = strRowName
        
        .MergeRow(lngNewRow) = True
        .MergeRow(lngNewRow + 2) = True
        
        .RowHidden(lngNewRow + 1) = True
        .RowHidden(lngNewRow + 3) = True
        
        For i = 1 To .COLS - 1
            If Len(Trim(.TextMatrix(0, i))) <> 0 Then  'check not a seperator column
                .TextMatrix(lngNewRow, i) = strDefaultValue1
                .TextMatrix(lngNewRow + 1, i) = ""
                .TextMatrix(lngNewRow + 2, i) = strdefaultvalue2
            End If
            
            .GetSelection r1, c1, r2, c2
            .Select lngNewRow, i, lngNewRow + 2, i
            .CellBorder RGB(0, 0, 0), -1, -1, 1, -1, -1, -1
            .Select r1, c1, r2, c2
            
        Next i
    End With
    Add2ValuePerToothRow = lngNewRow
End Function

Public Function InitCols() As Boolean

Dim i As Integer
Dim lngStartTooth As Long
Dim lngMidEndTooth As Long  'used for full mouth where the tooth order is not sequential
Dim lngMidStartTooth As Long
Dim lngEndTooth As Long
Dim lngCols As Long
Dim lngMidCols As Long
Dim lngValue As Long
Dim lngStep As Long
Dim lngMidStep As Long
Dim lngLimit As Long

    Select Case menumViewType
        Case perFullMouth, perTextEntry, perPrintGraphic
            lngCols = 16
            lngStartTooth = 1
            lngMidEndTooth = 16
            lngStep = 1
            lngMidCols = 16
            lngMidStartTooth = 32
            lngEndTooth = 17
            lngMidStep = -1
        Case perLLQ
            lngCols = 8
            lngStartTooth = 24
            lngEndTooth = 17
            lngStep = -1
        Case perLRQ
            lngCols = 8
            lngStartTooth = 32
            lngEndTooth = 25
            lngStep = -1
        Case perULQ
            lngCols = 8
            lngStartTooth = 9
            lngEndTooth = 16
            lngStep = 1
        Case perURQ
            lngCols = 8
            lngStartTooth = 1
            lngEndTooth = 8
            lngStep = 1
        Case perManArch
            lngCols = 16
            lngStartTooth = 32
            lngEndTooth = 17
            lngStep = -1
        Case perMaxArch
            lngCols = 16
            lngStartTooth = 1
            lngEndTooth = 16
            lngStep = 1
    End Select
    
    lngCols = lngCols * 4 '3 cols per tooth, plus a column seperators,
                            'plus 1 col to hold row descriptions
    lngMidCols = lngMidCols * 4
                            
    With mctlGrid
        '.Cols = lngCols
        .Redraw = False
        .COLS = lngCols + lngMidCols
                                    
        'Hide the column seperators
        lngLimit = lngCols + lngMidCols - 1
        For i = 4 To lngLimit Step 4
            .ColHidden(i) = True
        Next i
        
        If lngMidStartTooth = 0 Then
        
            'Set the column Headers
            lngValue = lngStartTooth
            For i = 1 To lngCols Step 4
                .Cell(flexcpText, 0, i, 0, i + 2) = lngValue
                lngValue = lngValue + lngStep
            Next i
        Else  'have to switch the tooth number in mid-stream
            lngValue = lngStartTooth
            For i = 1 To lngCols Step 4
                .Cell(flexcpText, 0, i, 0, i + 2) = lngValue
                lngValue = lngValue + lngStep
            Next i
            
            lngValue = lngMidStartTooth
            lngLimit = lngCols + lngMidCols
            For i = lngCols + 1 To lngLimit Step 4
                .Cell(flexcpText, 0, i, 0, i + 2) = lngValue
                lngValue = lngValue + lngMidStep
            Next i
        End If
    
        lngLimit = .COLS - 1
        
        '-1 index is supposed to apply to all columns instead of having to loop through
        .MergeCol(-1) = True
        .ColWidth(-1) = 300
        
        .ColWidth(0) = 1110 'Make big enough to display "Probing Depth"
        '.Redraw = True
    End With
    
    
    InitCols = True
End Function

Public Function AdvanceCells(lngCells As Long) As Boolean
Dim blnAdvance As Boolean
    
    With mctlGrid
        If (.COLS >= .Col + lngCells) And (.Col + lngCells > 0) Then
            'Make sure not crossing "boundary" in FullMouth/Text views
            
            
            Select Case menumViewType
                Case perFullMouth, perTextEntry
                    If .TextMatrix(0, .Col) = 16 And .TextMatrix(0, .Col + lngCells) = 32 Then 'check 16 to 32 crossing
                        blnAdvance = False
                    ElseIf .TextMatrix(0, .Col) = 32 And .TextMatrix(0, .Col + lngCells) = 16 Then
                        blnAdvance = False
                    ElseIf .Cell(flexcpBackColor, .Row, .Col + lngCells) = DISABLED_COLOR Then
                        blnAdvance = False
                    Else
                        blnAdvance = True
                    End If
                Case Else
                    If .Cell(flexcpBackColor, .Row, .Col + lngCells) = DISABLED_COLOR Then
                        blnAdvance = False
                    Else
                        blnAdvance = True
                    End If
                    
            End Select
            
            
            
        Else 'not enough cols left to move that far
            blnAdvance = False
        End If
        
        If blnAdvance Then
            .Col = .Col + lngCells
            AdvanceCells = True
        Else
            AdvanceCells = False
            RaiseEvent AdvanceCellsFail(.Row, .Col, .Row, .Col + lngCells)
        End If
    End With
End Function

Public Sub UpdateGM(intCol As Integer)
Dim lngCA As Long
Dim lngGM As Long
Dim lngProb As Long
Dim strTemp As String
Dim intCurrTooth As Integer
Dim eLoc As eMeasurementLocation

    With mctlGrid
        strTemp = .TextMatrix(mlngCARow, intCol)
        If Len(Trim(strTemp)) = 0 Then
            lngCA = 0
        Else
            If IsNumeric(strTemp) Then
                lngCA = CLng(strTemp)
            End If
        End If
        
        strTemp = .TextMatrix(mlngProbDepthRow, intCol)
        If Len(Trim(strTemp)) = 0 Then
            lngProb = 0
        Else
            If IsNumeric(strTemp) Then
                lngProb = CLng(strTemp)
            End If
        End If

        '.TextMatrix(mlngGMRow, intCol) = lngCA - lngProb
        intCurrTooth = .TextMatrix(0, intCol)
        eLoc = intCol Mod 4 'the location for all values with 6 values per tooth
        SetValue lngCA - lngProb, intCurrTooth, perGM, eLoc
        'RaiseEvent ValueSet(perGM, intCurrTooth, eLoc, CStr(lngCA - lngProb))
        
        'Indicate the GM can/can't be manually updated
'        If lngCA = 0 Or lngProb = 0 Then
'            .Cell(flexcpBackColor, mlngGMRow, intCol) = ENABLED_COLOR
'        Else
'            .Cell(flexcpBackColor, mlngGMRow, intCol) = DISABLED_COLOR
'        End If
        
    End With
End Sub

Public Sub UpdateCA(intCol As Integer)
Dim lngCA As Long
Dim lngGM As Long
Dim lngProb As Long
Dim strTemp As String
Dim intCurrTooth As Integer
Dim eLoc As eMeasurementLocation
    
    With mctlGrid
        strTemp = .TextMatrix(mlngProbDepthRow, intCol)
        If Len(Trim(strTemp)) = 0 Then
            lngProb = 0
        Else
            If IsNumeric(strTemp) Then
                lngProb = CLng(strTemp)
            End If
        End If
        
        strTemp = .TextMatrix(mlngGMRow, intCol)
        If Len(Trim(strTemp)) = 0 Then
            lngGM = 0
        Else
            If IsNumeric(strTemp) Then
                lngGM = CLng(strTemp)
            End If
        End If
        
        '.TextMatrix(mlngCARow, intCol) = lngProb + lngGM
        intCurrTooth = .TextMatrix(0, intCol)
        eLoc = intCol Mod 4 'the location for all values with 6 values per tooth
        SetValue lngProb + lngGM, intCurrTooth, perCA, eLoc
        'RaiseEvent ValueSet(perCA, intCurrTooth, eLoc, CStr(lngProb + lngGM))
        
        'Indicate the CA can/can't be manually updated
'        If lngGM = 0 Or lngProb = 0 Then
'            .Cell(flexcpBackColor, mlngCARow, intCol) = ENABLED_COLOR
'        Else
'            .Cell(flexcpBackColor, mlngCARow, intCol) = DISABLED_COLOR
'        End If
    End With
End Sub

Public Sub MakeCurrentCellVisible()
'    With mctlGrid
'        If .Col + 4 >= 0 And .Col + 4 <= .Cols Then
'            If Not .ColIsVisible(.Col) Then
'                .LeftCol = .Col
'            End If
'        End If
'        If .Col - 4 >= 0 Then
'            If Not .ColIsVisible(.Col - 4) Then
'                .LeftCol = .Col - 4
'            End If
'        End If
'    End With

    Dim lngCol

    With mctlGrid
        lngCol = .Col - (.Col Mod 4)
        If .Col + 4 >= 0 And .Col + 4 <= .COLS Then
            'If Not .ColIsVisible(.Col) Then
            If Not .ColIsVisible(lngCol + 3) Then
                '.LeftCol = .Col
                .LeftCol = lngCol + 1
            End If
        End If
        If .Col - 4 >= 0 Then
            'If Not .ColIsVisible(.Col - 4) Then
            If Not .ColIsVisible(lngCol + 1) Then
                '.LeftCol = .Col - 4
                .LeftCol = lngCol + 1
            End If
        End If
    End With
    

End Sub

Public Sub SetProbingDepthColor(ProbDepth As String, Row As Integer, Col As Integer)
    With mctlGrid
        If IsNumeric(ProbDepth) Then
            If CInt(ProbDepth) >= 4 Then
                .Cell(flexcpForeColor, Row, Col) = vbRed
            Else
                .Cell(flexcpForeColor, Row, Col) = vbBlue
            End If
        End If
    End With
End Sub

Private Sub SetToggleColor(enumMeasurement As eMeasurementType, lngCol As Long)
Dim lngTempRow As Long
Dim lngTempCol As Long
Dim lngTrueColor As Long
Dim lngRGB As Long
    
    With mctlGrid
        lngTempRow = .Row
        lngTempCol = .Col
        
        If lngCol = 0 Then
            lngCol = .Col
        End If
    
        Select Case enumMeasurement
            Case perBleeding
                lngTrueColor = vbRed
                .Row = mlngBleedingRow
            Case perSuppuration
                lngTrueColor = RGB(202, 202, 0)  'darker than vbYellow which is hard to see against white
                .Row = mlngSuppurationRow
            Case perPlaque
                lngTrueColor = vbBlue
                .Row = mlngPlaqueRow
        End Select
    
        .Col = lngCol
        
        If Len(Trim(.Text)) <> 0 Then
            .CellForeColor = lngTrueColor
        Else
            .CellForeColor = vbBlack
        End If
        
        .Row = lngTempRow
        .Col = lngTempCol
        
    End With
End Sub
Public Sub SetSuppurationColor(Optional lngCol As Long = 0)
    SetToggleColor perSuppuration, lngCol
End Sub
Public Sub SetBleedingColor(Optional lngCol As Long)
    SetToggleColor perBleeding, lngCol
'Dim lngTempRow As Long
'Dim lngTempCol As Long
'
'    With mctlGrid
'        lngTempRow = .Row
'        lngTempCol = .Col
'
'        If lngcol = 0 Then
'            lngcol = .Col
'        End If
'
'        .Row = mlngBleedingRow
'        .Col = lngcol
'
'        If Len(Trim(.Text)) <> 0 Then
'            .CellForeColor = vbRed
'        Else
'            .CellForeColor = vbBlack
'        End If
'
'        .Row = lngTempRow
'        .Col = lngTempCol
'
'    End With
End Sub
Public Sub SetPlaqueColor(Optional lngCol As Long)
    SetToggleColor perPlaque, lngCol
'Dim lngTempRow As Long
'Dim lngTempCol As Long
'
'    With mctlGrid
'        lngTempRow = .Row
'        lngTempCol = .Col
'
'        If lngcol = 0 Then
'            lngcol = .Col
'        End If
'
'        .Row = mlngPlaqueRow
'        .Col = lngcol
'
'        If Len(Trim(.Text)) <> 0 Then
'            .CellForeColor = vbBlue
'        Else
'            .CellForeColor = vbBlack
'        End If
'
'        .Row = lngTempRow
'        .Col = lngTempCol
'
'    End With
End Sub
    

Public Sub MoveLeft()
Dim lngCurrTooth As Long
Dim lngToothCol As Long
Dim lngCellsToAdvance As Long

    lngCellsToAdvance = 0
    
    With mctlGrid
        lngCurrTooth = .TextMatrix(0, .Col)
        lngToothCol = .Col

        Select Case .Row

            Case mlngCARow, mlngGMRow, mlngProbDepthRow, mlngBleedingRow, mlngSuppurationRow, mlngPlaqueRow
                Select Case .Col Mod 4
                    Case 1 '1st entry per tooth
                        lngCellsToAdvance = -2
                    Case 2 '2nd entry per tooth
                        lngCellsToAdvance = -1
                    Case 3 '3rd entry per tooth
                        lngCellsToAdvance = -1
                End Select
            Case mlngPrognosisRow, mlngMobilityRow, mlngBoneLossRow, mlngBoneLossRow + 2, mlngMgjRow
                lngCellsToAdvance = -4
            
            Case mlngFurcationRow, mlngFurcationRow + 2
                Dim strToothName As String
                Dim strNextToothName As String
                Dim strGrid1 As String
                Dim strGrid2 As String
                
                strToothName = Trim(.TextMatrix(mlngToothNameRow, .Col))
                If Len(strToothName) = 0 Then
                    strToothName = lngCurrTooth
                End If
                
                
                'Select Case lngCurrTooth
                Select Case strToothName
                    Case "3", "4", "5", "A", "B", "I", "J", "12", "13", _
                         "29", "28", "S", "T", "K", "L", "20", "21"
                        
                        strNextToothName = Trim(.TextMatrix(mlngToothNameRow, .Col - 4))
                        If Len(strNextToothName) = 0 Then
                            strNextToothName = Trim(.TextMatrix(0, .Col - 4))
                        End If
                        
                        strGrid1 = FurcationGridType(strToothName)
                        strGrid2 = FurcationGridType(strNextToothName)
                         
                        If strGrid1 = "3Cols" And strGrid2 = "3Cols" Then
                            Select Case .Col Mod 4
                                Case 1
                                    lngCellsToAdvance = -2
                                Case 2
                                    lngCellsToAdvance = -1
                                Case 3
                                    lngCellsToAdvance = -1
                            End Select
                        ElseIf strGrid1 = "3Cols" And strGrid2 = "2Rows" Then
                            Select Case lngToothCol Mod 4
                                Case 1
                                    lngCellsToAdvance = -4
                                Case 2
                                    lngCellsToAdvance = -1
                                Case 3
                                    lngCellsToAdvance = -1
                            End Select
                        ElseIf strGrid1 = "2Rows" And strGrid2 = "3Cols" Then
                            Select Case lngToothCol Mod 4
                                Case 1
                                    lngCellsToAdvance = -2
                                Case 2
                                    lngCellsToAdvance = -3
                                Case 3
                                    lngCellsToAdvance = -4
                            End Select
                        ElseIf strGrid1 = "2Rows" And strGrid2 = "2Rows" Then
                            lngCellsToAdvance = -4
                        Else
                        End If
                         
                         
                    Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", _
                        "13", "14", "15", "16", "21", "22", "23", "24", _
                        "25", "26", "27", "28", "29"               'currently 3 cols, moving to 3 cols
                        Select Case .Col Mod 4
                            Case 1
                                lngCellsToAdvance = -2
                            Case 2
                                lngCellsToAdvance = -1
                            Case 3
                                lngCellsToAdvance = -1
                        End Select
                    Case "20"  '3cols to 2 rows
                        Select Case lngToothCol Mod 4
                            Case 1
                                lngCellsToAdvance = -4
                            Case 2
                                lngCellsToAdvance = -1
                            Case 3
                                lngCellsToAdvance = -1
                        End Select
                    Case "18", "19", "31", "32" '2 rows to 2 rows
                        lngCellsToAdvance = -4
                    Case "17", "30" '2 rows to 3 cols
                        Select Case lngToothCol Mod 4
                            Case 1
                                lngCellsToAdvance = -2
                            Case 2
                                lngCellsToAdvance = -3
                            Case 3
                                lngCellsToAdvance = -4
                        End Select
                End Select
        End Select
    
    End With
    AdvanceCells lngCellsToAdvance
End Sub

Public Sub MoveRight()
Dim lngCurrTooth As Long
Dim lngToothCol As Long
Dim lngCellsToAdvance As Long

    lngCellsToAdvance = 0
    
    With mctlGrid
        lngCurrTooth = .TextMatrix(0, .Col)
        lngToothCol = .Col

        Select Case .Row
            Case mlngCARow, mlngGMRow, mlngProbDepthRow, mlngBleedingRow, mlngSuppurationRow, mlngPlaqueRow
                Select Case .Col Mod 4
                    Case 1 '1st entry per tooth
                        lngCellsToAdvance = 1
                    Case 2 '2nd entry per tooth
                        lngCellsToAdvance = 1
                    Case 3 '3rd entry per tooth
                        lngCellsToAdvance = 2
                End Select
            Case mlngPrognosisRow, mlngMobilityRow, mlngBoneLossRow, mlngBoneLossRow + 2, mlngMgjRow
                lngCellsToAdvance = 4
            Case mlngFurcationRow, mlngFurcationRow + 2
                Dim strToothName As String
                Dim strNextToothName As String
                Dim strGrid1 As String
                Dim strGrid2 As String
                
                strToothName = Trim(.TextMatrix(mlngToothNameRow, .Col))
                If Len(strToothName) = 0 Then
                    strToothName = lngCurrTooth
                End If
                
                
                'Select Case lngCurrTooth
                
                'Most of these, we know ahead of time what to do, but for position on and around
                'teeth that can have 2 or 3 roots, need to figure out what to do based
                'on primary/secondary status
                
                Select Case strToothName
                    Case "3", "4", "5", "A", "B", "I", "J", "12", "13", _
                         "29", "28", "S", "T", "K", "L", "20", "21"
                         
                        strNextToothName = Trim(.TextMatrix(mlngToothNameRow, .Col + 4))
                        If Len(strNextToothName) = 0 Then
                            strNextToothName = Trim(.TextMatrix(0, .Col + 4))
                        End If
                        
                        strGrid1 = FurcationGridType(strToothName)
                        strGrid2 = FurcationGridType(strNextToothName)
                    
                        If strGrid1 = "3Cols" And strGrid2 = "3Cols" Then
                            Select Case .Col Mod 4
                                Case 1
                                    lngCellsToAdvance = 1
                                Case 2
                                    lngCellsToAdvance = 1
                                Case 3
                                    lngCellsToAdvance = 2
                            End Select
                        ElseIf strGrid1 = "3Cols" And strGrid2 = "2Rows" Then
                            Select Case lngToothCol Mod 4
                                Case 1
                                    lngCellsToAdvance = 1
                                Case 2
                                    lngCellsToAdvance = 1
                                Case 3
                                    lngCellsToAdvance = 2
                            End Select
                        ElseIf strGrid1 = "2Rows" And strGrid2 = "3Cols" Then
                            Select Case lngToothCol Mod 4
                                Case 1
                                    lngCellsToAdvance = 4
                                Case 2
                                    lngCellsToAdvance = 3
                                Case 3
                                    lngCellsToAdvance = 2
                            End Select
                        ElseIf strGrid1 = "2Rows" And strGrid2 = "2Rows" Then
                            lngCellsToAdvance = 4
                        Else
                        End If
                    
                    Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", _
                        "13", "14", "15", "20", "21", "22", "23", "24", "25", "26", "27", _
                        "28"                  'currently 3 cols, moving to 3 cols
                        Select Case .Col Mod 4
                            Case 1
                                lngCellsToAdvance = 1
                            Case 2
                                lngCellsToAdvance = 1
                            Case 3
                                lngCellsToAdvance = 2
                        End Select
                    Case "16", "29" '3cols to 2 rows
                        Select Case lngToothCol Mod 4
                            Case 1
                                lngCellsToAdvance = 1
                            Case 2
                                lngCellsToAdvance = 1
                            Case 3
                                lngCellsToAdvance = 2
                        End Select
                    Case "17", "18", "30", "31", "32" '2 rows to 2 rows
                        lngCellsToAdvance = 4
                    Case "19" '2 rows to 3 cols
                        Select Case lngToothCol Mod 4
                            Case 1
                                lngCellsToAdvance = 4
                            Case 2
                                lngCellsToAdvance = 3
                            Case 3
                                lngCellsToAdvance = 2
                        End Select
                End Select
        End Select
       
    End With
    
    AdvanceCells lngCellsToAdvance
    
End Sub

Private Function FurcationGridType(strToothName As String) As String
    
    Select Case strToothName
        Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", _
                        "13", "14", "15", "16", "20", "21", "22", "23", "24", "25", "26", "27", _
                        "28", "29", "A", "B", "I", "J"
            FurcationGridType = "3Cols"
        Case "17", "18", "19", "30", "31", "32", "K", "L", "S", "T"
            FurcationGridType = "2Rows"
    End Select
End Function



Public Sub SetCurrValue(strText As String)
Dim intCurrTooth As Integer
Dim intToothCol As Integer
Dim strTemp As String
Dim blnSetValue As Boolean
Dim eLoc As eMeasurementLocation

    With mctlGrid
        intCurrTooth = .TextMatrix(0, .Col)
        intToothCol = Int(.Col / 4) * 4 + 1
        eLoc = .Col Mod 4 'the location for all values with 6 values per tooth
        
        Select Case .Row
            Case mlngCARow
                If .Cell(flexcpBackColor, mlngCARow, .Col) <> DISABLED_COLOR Then
                    '.Text = strText
                    If IsNumeric(strText) Or strText = " " Then
                        SetValue strText, intCurrTooth, perCA, eLoc
                        UpdateGM .Col
                    Else
                        .Text = strText
                    End If
                End If
                
            Case mlngGMRow
                '.Text = strText
                If .Cell(flexcpBackColor, mlngGMRow, .Col) <> DISABLED_COLOR Then
                    If IsNumeric(strText) Then
                        If CInt(strText) >= -20 And CInt(strText) <= 10 Then
                            SetValue strText, intCurrTooth, perGM, eLoc
                            UpdateCA .Col
                        Else
                            SetValue 0, intCurrTooth, perGM, eLoc
                            Beep
                        End If
                    ElseIf strText = " " Then
                        SetValue strText, intCurrTooth, perGM, eLoc
                        UpdateCA .Col
                    Else
                    End If
                End If
            
            Case mlngProbDepthRow
                If .Cell(flexcpBackColor, mlngProbDepthRow, .Col) <> DISABLED_COLOR Then
                    If IsNumeric(strText) Or strText = " " Then
                        SetValue strText, intCurrTooth, perProbDepth, eLoc
                        UpdateCA .Col
                        UpdateGM .Col
                    Else
                        .Text = strText
                    End If
                End If

            Case mlngBoneLossRow, mlngBoneLossRow + 2
                If .Cell(flexcpBackColor, mlngBoneLossRow, .Col) <> DISABLED_COLOR Then
                    If .Row = mlngBoneLossRow Then
                        strTemp = "M "
                        eLoc = perLocOne
                    Else
                        strTemp = "D "
                        eLoc = perLocTwo
                    End If
                    SetValue strTemp & strText, intCurrTooth, perBoneLoss, eLoc
                End If
            
            Case mlngPrognosisRow
                If .Cell(flexcpBackColor, mlngPrognosisRow, .Col) <> DISABLED_COLOR Then
                    SetValue strText, intCurrTooth, perPrognosis, perLocWholeTooth
                End If
            Case mlngMobilityRow
                If .Cell(flexcpBackColor, mlngMobilityRow, .Col) <> DISABLED_COLOR Then
                    If IsNumeric(strText) Then
                        If CInt(strText) >= 1 And CInt(strText) <= 3 Then
                            SetValue strText, intCurrTooth, perMobility, perLocWholeTooth
                        End If
                    ElseIf Len(Trim(strText)) = 0 Then
                        SetValue strText, intCurrTooth, perMobility, perLocWholeTooth
                    End If
                End If
            Case mlngMgjRow
                If .CellBackColor <> DISABLED_COLOR Then
                    SetValue strText, intCurrTooth, perMGJ, perLocOne
                End If
            Case mlngFurcationRow, mlngFurcationRow + 2
                Dim strToothName
                
                strToothName = Trim(.TextMatrix(mlngToothNameRow, .Col))
                If Len(strToothName) = 0 Then
                    strToothName = intCurrTooth
                End If
                
                If .CellBackColor = DISABLED_COLOR Then
                    blnSetValue = False
                ElseIf IsNumeric(strText) Then
                    If CInt(strText) >= 1 And CInt(strText) <= 4 Then
                        blnSetValue = True
                    End If
                ElseIf strText = " " Then
                    blnSetValue = True
                Else
                    blnSetValue = False
                End If
                
                If blnSetValue Then
                    'Select Case intCurrTooth
                    Select Case strToothName
                        'not allowed
                        Case "4", "6", "7", "8", "9", "10", "11", "13", "20", "21", _
                        "22", "23", "24", "25", "26", "27", "28", "29"  '1 root
                        'two values per tooth
                        Case "5", "12" '2 roots, 2 cols
                            If .Col Mod 4 = 1 Then
                                eLoc = perLocOne
                            ElseIf .Col Mod 4 = 3 Then
                                eLoc = perLocTwo
                            End If
                            SetValue strText, intCurrTooth, perFurcation, eLoc
                        Case "17", "18", "19", "30", "31", "32", "K", "L", "S", "T" '2 roots, 2 rows
                            If .Row = mlngFurcationRow Then
                                strTemp = "B "
                                eLoc = perLocOne
                            Else
                                strTemp = "L "
                                eLoc = perLocTwo
                            End If
                            SetValue strTemp & strText, intCurrTooth, perFurcation, eLoc
                        'three values per tooth
                        Case "1", "2", "3", "14", "15", "16", "A", "B", "I", "J"
                            SetValue strText, intCurrTooth, perFurcation, eLoc
                    End Select
                End If
        End Select
    
    End With

End Sub

Public Function SetValueFontColor(lngColor As Long, Tooth As Integer, Measurement As eMeasurementType, Location As eMeasurementLocation) As Boolean
'code copied from SetValue

    Dim intOrigRow As Integer
    Dim intOrigCol As Integer
    Dim intCol As Integer
    Dim intRow As Integer
    Dim i As Integer

    intOrigRow = mctlGrid.Row
    intOrigCol = mctlGrid.Col
        
    'find the first col for the tooth
    intCol = -1
    For i = 1 To mctlGrid.COLS - 1 Step 4
        If mctlGrid.TextMatrix(0, i) = Tooth Then
            intCol = i
            Exit For
        End If
    Next i
    If intCol = -1 Then 'couldn't find the col
        SetValueFontColor = False
        Exit Function
    End If
    
    intRow = GetRow(Measurement)
        
    With mctlGrid
        Select Case Measurement
            Case perFurcation
                Select Case Tooth
                    'not allowed
                    Case 4, 6, 7, 8, 9, 10, 11, 13, 20, 21, 22, 23, 24, _
                                                    25, 26, 27, 28, 29  '1 root
                    'two values per tooth
                    Case 5, 12 '2 roots, 2 cols
                        Select Case Location
                            Case perLocOne
                            Case perLocTwo
                                intCol = intCol + 2
                        End Select
                        '.Cell(flexcpForeColor, intRow, intCol) = lngColor
                        '.Cell(flexcpForeColor, intRow + 1, intCol) = lngColor
                        '.Cell(flexcpForeColor, intRow + 2, intCol) = lngColor
                        .Cell(flexcpForeColor, intRow, intCol, intRow + 2, intCol) = lngColor
                    
                    Case 17, 18, 19, 30, 31, 32 '2 roots, 2 rows
                        Select Case Location
                            Case perLocOne
                            Case perLocTwo
                                intRow = intRow + 2
                        End Select
                        '.Cell(flexcpForeColor, intRow, intCol) = lngColor
                        '.Cell(flexcpForeColor, intRow, intCol + 1) = lngColor
                        '.Cell(flexcpForeColor, intRow, intCol + 2) = lngColor
                        .Cell(flexcpForeColor, intRow, intCol, intRow, intCol + 2) = lngColor
                    
                    'three values per tooth
                    Case 1, 2, 3, 14, 15, 16
                        Select Case Location
                            Case perLocOne
                            Case perLocTwo
                                intCol = intCol + 1
                            Case perLocThree
                                intCol = intCol + 2
                        End Select
                        '.Cell(flexcpForeColor, intRow, intCol) = lngColor
                        '.Cell(flexcpForeColor, intRow + 1, intCol) = lngColor
                        '.Cell(flexcpForeColor, intRow + 2, intCol) = lngColor
                        .Cell(flexcpForeColor, intRow, intCol, intRow + 2, intCol) = lngColor
                End Select
            
            Case perBoneLoss
                Select Case Location
                    Case perLocOne
                    Case perLocTwo
                        intRow = intRow + 2
                End Select
                '.Cell(flexcpForeColor, intRow, intCol) = lngColor
                '.Cell(flexcpForeColor, intRow, intCol + 1) = lngColor
                '.Cell(flexcpForeColor, intRow, intCol + 2) = lngColor
                .Cell(flexcpForeColor, intRow, intCol, intRow, intCol + 2) = lngColor
                
            Case perMGJ, perMobility, perPrognosis
                '.Cell(flexcpForeColor, intRow, intCol) = lngColor
                '.Cell(flexcpForeColor, intRow, intCol + 1) = lngColor
                '.Cell(flexcpForeColor, intRow, intCol + 2) = lngColor
                .Cell(flexcpForeColor, intRow, intCol, intRow, intCol + 2) = lngColor
            
            Case Else
                Select Case Location
                    Case perLocOne
                    Case perLocTwo
                        intCol = intCol + 1
                    Case perLocThree
                        intCol = intCol + 2
                    Case perLocWholeTooth
                End Select
                .Cell(flexcpForeColor, intRow, intCol) = lngColor
        End Select
    End With
End Function

Public Function SetToothName(strToothName As String, Tooth As Integer)
    SetValue strToothName, Tooth, perToothName, perLocWholeTooth
End Function
Public Function SetMissing(blnMissing As Boolean, Tooth As Integer)
    SetValue CStr(blnMissing), Tooth, perMissing, perLocWholeTooth
End Function

Public Function SetValue(strText As String, Tooth As Integer, Measurement As eMeasurementType, Location As eMeasurementLocation) As Boolean
    Dim intOrigRow As Integer
    Dim intOrigCol As Integer
    Dim intCol As Integer
    Dim intRow As Integer
    Dim i As Integer
    Dim lngBackColor As Long
    Dim strToothName As String

    intOrigRow = mctlGrid.Row
    intOrigCol = mctlGrid.Col
        
    'find the first col for the tooth
    intCol = -1
    For i = 1 To mctlGrid.COLS - 1 Step 4
        If mctlGrid.TextMatrix(0, i) = Tooth Then
            intCol = i
            Exit For
        End If
    Next i
    If intCol = -1 Then 'couldn't find the col
        SetValue = False
        Exit Function
    End If
    
    intRow = GetRow(Measurement)
    
    'Read the toothname, if it's not set, use the tooth position
    '(need this to set furcation)
    strToothName = Trim(mctlGrid.TextMatrix(mlngToothNameRow, intCol))
    If Len(strToothName) = 0 Then
        strToothName = Tooth
    End If
    
    With mctlGrid
        Select Case Measurement
        
            Case perToothName
                'Have to add trailing space for mergecells to work correctly, don't know why
                .Cell(flexcpText, intRow, intCol, intRow, intCol + 2) = strText & " "
                If menumGridType = perLingual Then
                    FormatFurcation (intCol)
                    FormatFurcation (intCol + 1)
                    FormatFurcation (intCol + 2)
                End If
            Case perMissing, perPontic, perUnerupted
                'if missing need to disable all measurements for that tooth
                If CBool(strText) Then
                    lngBackColor = DISABLED_COLOR
                    
                    If menumGridType = perBuccal Then
                        .Cell(flexcpBackColor, mlngProbDepthRow, intCol, mlngPlaqueRow, intCol + 2) = lngBackColor
                    ElseIf menumGridType = perLingual Then
                        .Cell(flexcpBackColor, mlngProbDepthRow, intCol, mlngPrognosisRow, intCol + 2) = lngBackColor
                    End If
                Else
                    lngBackColor = ENABLED_COLOR
                End If
                                
            Case perFurcation
                Select Case strToothName
                    'not allowed
                    Case "4", "6", "7", "8", "9", "10", "11", "13", "20", "21", _
                            "22", "23", "24", "25", "26", "27", "28", "29" '1 root
                    'two values per tooth
                    Case "5", "12" '2 roots, 2 cols
                        Select Case Location
                            Case perLocOne
                            Case perLocTwo
                                intCol = intCol + 2
                        End Select
                        '.TextMatrix(intRow, intCol) = strText
                        '.TextMatrix(intRow + 1, intCol) = strText
                        '.TextMatrix(intRow + 2, intCol) = strText
                        .Cell(flexcpText, intRow, intCol, intRow + 2, intCol) = strText
                    
                    Case "17", "18", "19", "30", "31", "32", "K", "L", "S", "T" '2 roots, 2 rows
                        Select Case Location
                            Case perLocOne
                            Case perLocTwo
                                intRow = intRow + 2
                        End Select
                        '.TextMatrix(intRow, intCol) = strText
                        '.TextMatrix(intRow, intCol + 1) = strText
                        '.TextMatrix(intRow, intCol + 2) = strText
                        .Cell(flexcpText, intRow, intCol, intRow, intCol + 2) = strText
                    
                    'three values per tooth
                    Case "1", "2", "3", "14", "15", "16", "A", "B", "I", "J"
                        Select Case Location
                            Case perLocOne
                            Case perLocTwo
                                intCol = intCol + 1
                            Case perLocThree
                                intCol = intCol + 2
                        End Select
                        '.TextMatrix(intRow, intCol) = strText
                        '.TextMatrix(intRow + 1, intCol) = strText
                        '.TextMatrix(intRow + 2, intCol) = strText
                        .Cell(flexcpText, intRow, intCol, intRow + 2, intCol) = strText
                End Select
            
            Case perBoneLoss
                Select Case Location
                    Case perLocOne
                    Case perLocTwo
                        intRow = intRow + 2
                End Select
                '.TextMatrix(intRow, intCol) = strText
                '.TextMatrix(intRow, intCol + 1) = strText
                '.TextMatrix(intRow, intCol + 2) = strText
                .Cell(flexcpText, intRow, intCol, intRow, intCol + 2) = strText
            Case perMGJ, perMobility, perPrognosis
                '.TextMatrix(intRow, intCol) = strText
                '.TextMatrix(intRow, intCol + 1) = strText
                '.TextMatrix(intRow, intCol + 2) = strText
                .Cell(flexcpText, intRow, intCol, intRow, intCol + 2) = strText
            Case Else
                Select Case Location
                    Case perLocOne
                    Case perLocTwo
                        intCol = intCol + 1
                    Case perLocThree
                        intCol = intCol + 2
                    Case perLocWholeTooth
                End Select
                .TextMatrix(intRow, intCol) = strText
                
                Select Case Measurement
                    Case perProbDepth
                        SetProbingDepthColor strText, intRow, intCol
                    Case perPlaque
                        SetPlaqueColor CLng(intCol)
                    Case perBleeding
                        SetBleedingColor CLng(intCol)
                    Case perSuppuration
                        SetSuppurationColor CLng(intCol)
                End Select
                
        End Select
    End With
    
    If Not mblnIsLoading Then
      RaiseEvent ValueSet(Measurement, Tooth, Location, strText)
    End If
    
End Function

Friend Function InitGrid() As Boolean
   
   With mctlGrid
      .Redraw = False
    'Hide the first row of Tooth Position
    .RowHidden(0) = True
    
    .Editable = flexEDNone
    'mctlGrid.BackColorAlternate = &HC0FFFF
    .BackColorAlternate = mctlGrid.BackColor
    
    'mctlGrid.ExtendLastCol = True
    .ExtendLastCol = False
    .MergeCells = flexMergeFree
    
    'No need for selections
    .AllowSelection = False
    .AllowBigSelection = False
    
    InitCols
        
    AddToothNameRow
    AddMissingRow
    AddProbDepthRow
    AddBleedingRow
    AddSuppurationRow
    AddGMRow
    AddCARow
    If menumGridType = perLingual Then
        AddMobilityRow
        AddFurcationRow
    End If
    AddMgjRow
    
    AddPlaqueRow
    If menumGridType = perLingual Then
        AddBoneLossRow
        AddPrognosisRow
    End If
    
    'Change FixedRows to 2 so that ToothName shows up correctly
    .FixedRows = 2
    
    .GridLines = flexGridFlat
   End With
    InitGrid = True
End Function


Friend Sub TogglePlaque()

Dim strPlaque As String
Dim lngPrevCol As Long
Dim intCurrTooth As Integer
Dim strValue As String
Dim blnApply As Boolean
    
    Select Case mctlGrid.Row
        
        Case mlngGMRow, mlngCARow, mlngProbDepthRow  'apply Plaque to previous tooth location
            'get the column for the previous tooth location
            Select Case menumGridType
                Case perBuccal
                    'If at far left side, can't apply to previous
                    If mctlGrid.Col > 1 Then
                        lngPrevCol = mctlGrid.Col - 1
                        If lngPrevCol Mod 4 = 0 Then
                            lngPrevCol = lngPrevCol - 1
                        End If
                        blnApply = True
                    Else
                        blnApply = False
                    End If
                Case perLingual
                    'if at far right side, can't apply to previous
                    If mctlGrid.Col < mctlGrid.COLS - 1 Then
                        lngPrevCol = mctlGrid.Col + 1
                        If lngPrevCol Mod 4 = 0 Then
                            lngPrevCol = lngPrevCol + 1
                        End If
                        blnApply = True
                    Else
                        blnApply = False
                    End If
            End Select
            
            If blnApply Then
                If mctlGrid.Cell(flexcpBackColor, mlngPlaqueRow, lngPrevCol) <> DISABLED_COLOR Then
                    strPlaque = mctlGrid.TextMatrix(mlngPlaqueRow, lngPrevCol)
                    If UCase(strPlaque) = "P" Then
                        strValue = ""
                    Else
                        strValue = "P"
                    End If
                    mctlGrid.TextMatrix(mlngPlaqueRow, lngPrevCol) = strValue
                    
                    SetPlaqueColor (lngPrevCol)
                    intCurrTooth = mctlGrid.TextMatrix(0, lngPrevCol)
                    RaiseEvent ValueSet(perPlaque, intCurrTooth, lngPrevCol Mod 4, strValue)
                End If
            Else 'Can't apply it to the grid, but need to have it changed at data level
                intCurrTooth = mctlGrid.TextMatrix(0, mctlGrid.Col)
                'Input values will get overwritten
                RaiseEvent TogglePrevValue(perPlaque, intCurrTooth, mctlGrid.Col Mod 4, strValue)
            End If
            
        Case mlngPlaqueRow 'apply bleeding to current tooth location
            If mctlGrid.Cell(flexcpBackColor, mlngPlaqueRow, mctlGrid.Col) <> DISABLED_COLOR Then
                If UCase(mctlGrid.Text) = "P" Then
                    strValue = ""
                Else
                    strValue = "P"
                End If
                mctlGrid.Text = strValue
                
                SetPlaqueColor
                intCurrTooth = mctlGrid.TextMatrix(0, mctlGrid.Col)
                RaiseEvent ValueSet(perPlaque, intCurrTooth, mctlGrid.Col Mod 4, strValue)
            End If
    End Select

End Sub

Private Sub Class_Initialize()

'Set the row values to a value that's obviously bad
 mlngCARow = -50
 mlngGMRow = -50
 mlngProbDepthRow = -50
 mlngBleedingRow = -50
 mlngPrognosisRow = -50
 mlngMobilityRow = -50
 mlngFurcationRow = -50
 mlngBoneLossRow = -50
 mlngSuppurationRow = -50
 mlngPlaqueRow = -50
 mlngMgjRow = -50

End Sub

Public Sub ToggleBleeding()

Dim strBleeding As String
Dim lngPrevCol As Long
Dim strValue As String
Dim intCurrTooth As Integer
Dim blnApply As Boolean
        'intCurrTooth = .TextMatrix(0, .Col)
    
    Select Case mctlGrid.Row
        
        Case mlngGMRow, mlngCARow, mlngProbDepthRow  'apply bleeding to previous tooth location
            'get the column for the previous tooth location
                    
            If menumGridType = perBuccal Then
                'If at far left side, can't apply to previous
                If mctlGrid.Col > 1 Then
                    lngPrevCol = mctlGrid.Col - 1
                    If lngPrevCol Mod 4 = 0 Then
                        lngPrevCol = lngPrevCol - 1
                    End If
                    blnApply = True
                Else
                    blnApply = False
                End If
            ElseIf menumGridType = perLingual Then
                'if at far right side, can't apply to previous
                If mctlGrid.Col < mctlGrid.COLS - 1 Then
                    lngPrevCol = mctlGrid.Col + 1
                    If lngPrevCol Mod 4 = 0 Then
                        lngPrevCol = lngPrevCol + 1
                    End If
                    blnApply = True
                Else
                    blnApply = False
                End If
            End If
                
            If blnApply Then
                If mctlGrid.Cell(flexcpBackColor, mlngBleedingRow, lngPrevCol) <> DISABLED_COLOR Then
                    strBleeding = mctlGrid.TextMatrix(mlngBleedingRow, lngPrevCol)
                    If UCase(strBleeding) = "B" Then
                        strValue = ""
                    Else
                        strValue = "B"
                    End If
                    
                    mctlGrid.TextMatrix(mlngBleedingRow, lngPrevCol) = strValue
                    SetBleedingColor (lngPrevCol)
                    intCurrTooth = mctlGrid.TextMatrix(0, lngPrevCol)
                    RaiseEvent ValueSet(perBleeding, intCurrTooth, lngPrevCol Mod 4, strValue)
                End If
            Else 'Can't apply it to the grid, but need to have it changed at data level
                intCurrTooth = mctlGrid.TextMatrix(0, mctlGrid.Col)
                'Input values will get overwritten
                RaiseEvent TogglePrevValue(perBleeding, intCurrTooth, mctlGrid.Col Mod 4, strValue)
            End If
        Case mlngBleedingRow 'apply bleeding to current tooth location
            If mctlGrid.Cell(flexcpBackColor, mlngBleedingRow, mctlGrid.Col) <> DISABLED_COLOR Then
                If UCase(mctlGrid.Text) = "B" Then
                    strValue = ""
                Else
                    strValue = "B"
                End If
                
                mctlGrid.Text = strValue
                SetBleedingColor
                intCurrTooth = mctlGrid.TextMatrix(0, mctlGrid.Col)
                RaiseEvent ValueSet(perBleeding, intCurrTooth, mctlGrid.Col Mod 4, strValue)
            End If
            
    End Select

End Sub

Public Sub ToggleSuppuration()

Dim strSupp As String
Dim lngPrevCol As Long
Dim strValue As String
Dim intCurrTooth As Integer
Dim blnApply As Boolean
    
    Select Case mctlGrid.Row
        
        Case mlngGMRow, mlngCARow, mlngProbDepthRow  'apply to previous tooth location
            'get the column for the previous tooth location
            If menumGridType = perBuccal Then
                'If at far left side, can't apply to previous
                If mctlGrid.Col > 1 Then
                    lngPrevCol = mctlGrid.Col - 1
                    If lngPrevCol Mod 4 = 0 Then
                        lngPrevCol = lngPrevCol - 1
                    End If
                    blnApply = True
                Else
                    blnApply = False
                End If
            ElseIf menumGridType = perLingual Then
                'if at far right side, can't apply to previous
                If mctlGrid.Col < mctlGrid.COLS - 1 Then
                    lngPrevCol = mctlGrid.Col + 1
                    If lngPrevCol Mod 4 = 0 Then
                        lngPrevCol = lngPrevCol + 1
                    End If
                    blnApply = True
                Else
                    blnApply = False
                End If
            End If
                            
            If blnApply Then
                If mctlGrid.Cell(flexcpBackColor, mlngSuppurationRow, lngPrevCol) <> DISABLED_COLOR Then
                    strSupp = mctlGrid.TextMatrix(mlngSuppurationRow, lngPrevCol)
                    If UCase(strSupp) = "S" Then
                        strValue = ""
                    Else
                        strValue = "S"
                    End If
                    mctlGrid.TextMatrix(mlngSuppurationRow, lngPrevCol) = strValue
                    SetSuppurationColor (lngPrevCol)
                    intCurrTooth = mctlGrid.TextMatrix(0, lngPrevCol)
                    RaiseEvent ValueSet(perSuppuration, intCurrTooth, lngPrevCol Mod 4, strValue)
                End If
            Else 'Can't apply it to the grid, but need to have it changed at data level
                intCurrTooth = mctlGrid.TextMatrix(0, mctlGrid.Col)
                'Input values will get overwritten
                RaiseEvent TogglePrevValue(perSuppuration, intCurrTooth, mctlGrid.Col Mod 4, strValue)
            End If
        Case mlngSuppurationRow 'apply bleeding to current tooth location
            If mctlGrid.Cell(flexcpBackColor, mlngSuppurationRow, mctlGrid.Col) <> DISABLED_COLOR Then
                If UCase(mctlGrid.Text) = "S" Then
                    strValue = ""
                Else
                    strValue = "S"
                End If
                mctlGrid.Text = strValue
                SetSuppurationColor
                intCurrTooth = mctlGrid.TextMatrix(0, mctlGrid.Col)
                RaiseEvent ValueSet(perSuppuration, intCurrTooth, mctlGrid.Col Mod 4, strValue)
            End If
    End Select

End Sub


Public Function CurrentRowType() As eMeasurementType
Dim enumRow As eMeasurementType

    Select Case mctlGrid.Row
        Case mlngCARow
            enumRow = perCA
        Case mlngGMRow
            enumRow = perGM
        Case mlngProbDepthRow
            enumRow = perProbDepth
        Case mlngBleedingRow
            enumRow = perBleeding
        Case mlngFurcationRow
            enumRow = perFurcation
        Case mlngMobilityRow
            enumRow = perMobility
        Case mlngPrognosisRow
            enumRow = perPrognosis
        Case mlngBoneLossRow
            enumRow = perBoneLoss
        Case mlngSuppurationRow
            enumRow = perSuppuration
        Case mlngPlaqueRow
            enumRow = perPlaque
        Case mlngMgjRow
            enumRow = perMGJ
        Case Else
            If menumGridType = perLingual Then
                If mctlGrid.Row = mlngBoneLossRow + 2 Then
                    enumRow = perBoneLoss
                ElseIf mctlGrid.Row = mlngFurcationRow + 2 Then
                    enumRow = perFurcation
                End If
            End If
    End Select
    CurrentRowType = enumRow
End Function

Public Function GetRow(enumRow As eMeasurementType) As Integer
Dim eRow As eMeasurementType
Dim intRow As Integer
    eRow = enumRow
    Select Case eRow
        Case eMeasurementType.perCA
            intRow = mlngCARow
        Case eMeasurementType.perGM
            intRow = mlngGMRow
        Case eMeasurementType.perProbDepth
            intRow = mlngProbDepthRow
        Case eMeasurementType.perBleeding
            intRow = mlngBleedingRow
        Case eMeasurementType.perFurcation
            intRow = mlngFurcationRow
        Case eMeasurementType.perMobility
            intRow = mlngMobilityRow
        Case eMeasurementType.perPrognosis
            intRow = mlngPrognosisRow
        Case eMeasurementType.perBoneLoss
            intRow = mlngBoneLossRow
        Case eMeasurementType.perSuppuration
            intRow = mlngSuppurationRow
        Case eMeasurementType.perPlaque
            intRow = mlngPlaqueRow
        Case eMeasurementType.perMGJ
            intRow = mlngMgjRow
        Case eMeasurementType.perToothName
            intRow = mlngToothNameRow
    End Select
    GetRow = intRow
End Function

Public Sub SetRow(enumRow As eMeasurementType)
Dim eRow As eMeasurementType
Dim lngNewRow As Long
    eRow = enumRow
    Select Case eRow
        Case eMeasurementType.perCA
            lngNewRow = mlngCARow
        Case eMeasurementType.perGM
            lngNewRow = mlngGMRow
        Case eMeasurementType.perProbDepth
            lngNewRow = mlngProbDepthRow
        Case eMeasurementType.perBleeding
            lngNewRow = mlngBleedingRow
        Case eMeasurementType.perFurcation
            lngNewRow = mlngFurcationRow
        Case eMeasurementType.perMobility
            lngNewRow = mlngMobilityRow
        Case eMeasurementType.perPrognosis
            lngNewRow = mlngPrognosisRow
        Case eMeasurementType.perBoneLoss
            lngNewRow = mlngBoneLossRow
        Case eMeasurementType.perSuppuration
            lngNewRow = mlngSuppurationRow
        Case eMeasurementType.perPlaque
            lngNewRow = mlngPlaqueRow
        Case eMeasurementType.perMGJ
            lngNewRow = mlngMgjRow
    End Select
    mctlGrid.Row = lngNewRow
End Sub

Public Function GetValue() As String
Dim strValue As String
    strValue = Trim(mctlGrid.Text)
    If CurrentRowType = perBoneLoss Then
        'Strip off the "M/D"
        strValue = Right(strValue, Len(strValue) - 1)
    End If
    GetValue = Trim(strValue)
End Function


Private Sub mctlGrid_BeforeSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long, Cancel As Boolean)
    'Keep selection colors the same as the cell itself, so selection isn't "seen"
   With mctlGrid
    .BackColorSel = .Cell(flexcpBackColor, NewRowSel, NewColSel)
    If .BackColorSel = 0 Then 'getting back black, which shouldn't be
        .BackColorSel = vbWhite
    End If
    .ForeColorSel = .Cell(flexcpForeColor, NewRowSel, NewColSel)
   End With
End Sub

Friend Sub Dispose()
    Set mctlGrid = Nothing
End Sub
