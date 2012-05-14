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
Attribute VB_Name = "ToothGridViews"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents mobjBuccalGrid As BuccalGridView
Attribute mobjBuccalGrid.VB_VarHelpID = -1
Private WithEvents mobjLingualGrid As LingualGridView
Attribute mobjLingualGrid.VB_VarHelpID = -1

'Private mobjMouth As MouthBL

Public Event EndOfGrid(CurrentGrid As eGridType, CurrentTooth As Integer) 'ala EOF
Public Event BeginningOfGrid(CurrentGrid As eGridType, CurrentTooth As Integer) 'ala BOF
Public Event ValueSet(Measurement As eMeasurementType, Tooth As Integer, GridType As eGridType, Location As eMeasurementLocation, Value As Variant)
Public Event AdvanceToDisabled(CurrentGrid As eGridType, FromRow As Long, FromCol As Long, ToRow As Long, ToCol As Long, CurrentTooth As Integer)
Public Event TogglePrevValue(Measurement As eMeasurementType, Tooth As Integer, GridType As eGridType, Location As eMeasurementLocation, Value As Variant)

Private mblnEnableEdit As Boolean


'Public Property Get Mouth() As MouthBL
'    Set Mouth = mobjMouth
'End Property
'Public Property Let Mouth(objMouth As MouthBL)
'    Set mobjMouth = objMouth
'    Set BuccalGrid.Mouth = objMouth
'    Set LingualGrid.Mouth = objMouth
'End Property

Private Sub Class_Initialize()
    Set mobjBuccalGrid = New BuccalGridView
    Set mobjLingualGrid = New LingualGridView
End Sub

Private Sub Class_Terminate()
    Set mobjBuccalGrid = Nothing
    Set mobjLingualGrid = Nothing
End Sub

Friend Property Get BuccalGrid() As BuccalGridView
    Set BuccalGrid = mobjBuccalGrid
End Property
Friend Property Get LingualGrid() As LingualGridView
    Set LingualGrid = mobjLingualGrid
End Property

Public Function Init(View As eViewType) As Boolean
    mobjBuccalGrid.Grid.Redraw = False
    mobjLingualGrid.Grid.Redraw = False
    
    'Perform generic initializing
    InitGrid mobjBuccalGrid.Grid, View
    InitGrid mobjLingualGrid.Grid, View
    
    'Perform grid specific init
    mobjBuccalGrid.InitGrid View
    mobjLingualGrid.InitGrid View
    
    mobjBuccalGrid.Grid.Redraw = True
    mobjLingualGrid.Grid.Redraw = True
    
    
'    'testing cell merge
'    With mobjLingualGrid.Grid
'        .MergeCells = flexMergeRestrictAll
'        .MergeCol(0) = True: .MergeCol(1) = True: .MergeCol(2) = True
'        .MergeCol(3) = False
'        .TextMatrix(1, 1) = "Row1Col1"
'        .TextMatrix(2, 1) = "Row1Col1"
'        .RowHidden(3) = True
'        .TextMatrix(1, 2) = "Row1Col1"
'        .TextMatrix(2, 2) = "Row1Col1"
'    End With


    
'    'test for writing multiple lines in a cell
'    mobjBuccalGrid.Grid.RowHeight(1) = mobjBuccalGrid.Grid.RowHeight(1) * 2
'    mobjBuccalGrid.Grid.TextMatrix(1, 1) = "Text1" & vbCrLf & "text2"
'
'    'test for writing values vertically
'    With mobjBuccalGrid.Grid
'        .TextMatrix(1, 2) = "1  2  3" & vbCrLf & "4  5  6"
'        .TextMatrix(1, 3) = "14  25  36"
'    End With
    
    Init = True
End Function

Friend Function InitGrid(ctlGrid As VSFlexGrid, View As eViewType) As Boolean
    
    With ctlGrid
         .Redraw = False
        .COLS = 1
        .ROWS = 1
        .Editable = True 'flexEDKbd
        .BackColorAlternate = &HC0FFFF
        .ExtendLastCol = True
        .MergeCells = flexMergeFree
        .MergeRow(0) = True
        .MergeCol(0) = True
        '.Redraw = True
    End With
    
    InitGrid = True
End Function

Private Sub mobjBuccalGrid_AdvanceCellsFail(Row As Long, Col As Long, ToRow As Long, ToCol As Long)
Dim intTooth As Integer
Dim blnDone As Boolean

    'If at end of row, switch over to Lingual grid
    With mobjBuccalGrid.Grid
        intTooth = .TextMatrix(0, .Col)
        
        'if tooth is disabled, find next valid tooth
        If .COLS >= ToCol And ToCol >= 0 Then
            If .Cell(flexcpBackColor, ToRow, ToCol) = mobjBuccalGrid.DisabledColor Then
                RaiseEvent AdvanceToDisabled(perBuccal, Row, Col, ToRow, ToCol, intTooth)
                blnDone = True
            End If
        End If
        
        If Not blnDone Then
            Select Case mobjBuccalGrid.CurrentRowType
                Case perCA, perGM, perProbDepth
                    If Col = .COLS - 1 Then
                        RaiseEvent EndOfGrid(perBuccal, intTooth)
                    ElseIf Col = 1 Then
                        RaiseEvent BeginningOfGrid(perBuccal, intTooth)
                    ElseIf .TextMatrix(0, Col) = 16 Then
                        RaiseEvent EndOfGrid(perBuccal, intTooth)
                    ElseIf .TextMatrix(0, Col) = 32 Then
                        RaiseEvent BeginningOfGrid(perBuccal, intTooth)
                    End If
                Case perMGJ
                    If Col = .COLS - 3 Then
                        RaiseEvent EndOfGrid(perBuccal, intTooth)
                    ElseIf Col = 1 Then
                        RaiseEvent BeginningOfGrid(perBuccal, intTooth)
                    ElseIf .TextMatrix(0, Col) = 16 Then
                    ElseIf .TextMatrix(0, Col) = 32 Then
                    
                    End If
            End Select
        End If
    End With
End Sub

Private Sub mobjBuccalGrid_TogglePrevValue(Measurement As PerioBL.eMeasurementType, Tooth As Integer, Location As PerioBL.eMeasurementLocation, Value As Variant)
    Dim intOrigTooth As Integer
    Dim eOrigGrid As eGridType
    Dim eGrid As eGridType
    Dim strValue As String
    
    eGrid = perBuccal
    
    'Usually the prev value is not visible, but Tooth 16 and 17 are, so make sure we update the grid values
    'in the other grid
    intOrigTooth = Tooth
    eOrigGrid = eGrid
    RaiseEvent TogglePrevValue(Measurement, Tooth, eGrid, Location, Value)
    
    'Tooth now indicates which tooth was updated, if it's the same tooth, then it's grid value is visible
    'so update
    If Tooth = intOrigTooth Then
        If Value = True Then
            Select Case Measurement
                Case perBleeding
                    strValue = "B"
                Case perPlaque
                    strValue = "P"
                Case perSuppuration
                    strValue = "S"
            End Select
        Else
            strValue = ""
        End If
        mobjBuccalGrid.SetValue strValue, Tooth, Measurement, Location
    End If
End Sub

Private Sub mobjBuccalGrid_ValueSet(Measurement As PerioBL.eMeasurementType, Tooth As Integer, Location As PerioBL.eMeasurementLocation, Value As Variant)
    RaiseEvent ValueSet(Measurement, Tooth, perBuccal, Location, Value)
End Sub


Private Sub mobjLingualGrid_AdvanceCellsFail(Row As Long, Col As Long, ToRow As Long, ToCol As Long)
Dim intTooth As Integer
Dim blnDone As Boolean

    With mobjLingualGrid.Grid
        intTooth = .TextMatrix(0, .Col)
        
        'if tooth is disabled, find next valid tooth
        If .COLS >= ToCol And ToCol >= 0 Then
            If .Cell(flexcpBackColor, ToRow, ToCol) = mobjLingualGrid.DisabledColor Then
                RaiseEvent AdvanceToDisabled(perLingual, Row, Col, ToRow, ToCol, intTooth)
                blnDone = True
            End If
        End If
        
        If Not blnDone Then
            Select Case mobjLingualGrid.CurrentRowType
                Case perCA, perGM, perProbDepth ', perBoneLoss, perMGJ
                    'If at end of row, switch over to Buccal grid
                    If Col = .COLS - 1 Then
                        RaiseEvent BeginningOfGrid(perLingual, intTooth)
                    ElseIf Col = 1 Then
                        RaiseEvent EndOfGrid(perLingual, intTooth)
                    ElseIf .TextMatrix(0, Col) = 16 Then
                        RaiseEvent BeginningOfGrid(perLingual, intTooth)
                    ElseIf .TextMatrix(0, Col) = 32 Then
                        RaiseEvent EndOfGrid(perLingual, intTooth)
                    End If
            End Select
        End If
    End With
End Sub

Private Sub mobjLingualGrid_TogglePrevValue(Measurement As PerioBL.eMeasurementType, Tooth As Integer, Location As PerioBL.eMeasurementLocation, Value As Variant)
    Dim intOrigTooth As Integer
    Dim eOrigGrid As eGridType
    Dim eGrid As eGridType
    Dim strValue As String
    
    eGrid = perLingual
    
    'Usually the prev value is not visible, but Tooth 16 and 17 are, so make sure we update the grid values
    'in the other grid
    intOrigTooth = Tooth
    eOrigGrid = eGrid
    RaiseEvent TogglePrevValue(Measurement, Tooth, eGrid, Location, Value)
    
    'Tooth now indicates which tooth was updated, if it's the same tooth, then it's grid value is visible
    'so update
    If Tooth = intOrigTooth Then
        If Value = True Then
            Select Case Measurement
                Case perBleeding
                    strValue = "B"
                Case perPlaque
                    strValue = "P"
                Case perSuppuration
                    strValue = "S"
            End Select
        Else
            strValue = ""
        End If
        mobjBuccalGrid.SetValue strValue, Tooth, Measurement, Location
    End If
    
End Sub

Private Sub mobjLingualGrid_ValueSet(Measurement As PerioBL.eMeasurementType, Tooth As Integer, Location As PerioBL.eMeasurementLocation, Value As Variant)
    RaiseEvent ValueSet(Measurement, Tooth, perLingual, Location, Value)
End Sub

Friend Sub Dispose()
    
    mobjBuccalGrid.Dispose
    Set mobjBuccalGrid = Nothing
    
    mobjLingualGrid.Dispose
    Set mobjLingualGrid = Nothing

End Sub

Public Property Get EnableEdit() As Boolean
   EnableEdit = mblnEnableEdit
End Property
Public Property Let EnableEdit(blnEnable As Boolean)
   mblnEnableEdit = blnEnable
   
   mobjLingualGrid.EnableEdit = blnEnable
   mobjBuccalGrid.EnableEdit = blnEnable
End Property


