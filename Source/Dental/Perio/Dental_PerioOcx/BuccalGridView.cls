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
Attribute VB_Name = "BuccalGridView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mlngToothNameRow As Long
Private mlngMissingRow As Long
Private mlngProbDepthRow As Long
Private mlngGMRow As Long
Private mlngCARow As Long
Private mlngSuppurationRow As Long
Private WithEvents mctlGrid As VSFlexGrid
Attribute mctlGrid.VB_VarHelpID = -1
Private menumViewType As eViewType
Private mblnPlusKeyPending 'Indicates a "+" key was pressed and is awaiting the next value
Private mblnMinusKeyPending 'Indicates a "-" key was pressed to make a value negative
Private WithEvents mobjGenGrid As GenGridView
Attribute mobjGenGrid.VB_VarHelpID = -1
Private mblnGMRowVisible As Boolean
Private mblnMGJRowVisible As Boolean


'Private mobjMouth As MouthBL

Public Event AdvanceCellsFail(Row As Long, Col As Long, ToRow As Long, ToCol As Long)
Public Event ValueSet(Measurement As eMeasurementType, Tooth As Integer, Location As eMeasurementLocation, Value As Variant)
Public Event TogglePrevValue(Measurement As eMeasurementType, Tooth As Integer, Location As eMeasurementLocation, Value As Variant)

Private mblnEnableEdit As Boolean

Private mblnIsLoading As Boolean

Public Property Get IsLoading() As Boolean
   IsLoading = mblnIsLoading
End Property
Public Property Let IsLoading(blnLoading As Boolean)
   mblnIsLoading = blnLoading
   mobjGenGrid.IsLoading = mblnIsLoading
End Property

Public Property Get DisabledColor() As Long
    DisabledColor = mobjGenGrid.DisabledColor
End Property
Public Property Get EnabledColor() As Long
    EnabledColor = mobjGenGrid.EnabledColor
End Property

Public Function CellIsValid(Row As Long, Col As Long) As Boolean
    CellIsValid = mobjGenGrid.CellIsValid(Row, Col)
End Function
Public Function CurrentCellIsValid() As Boolean
    CurrentCellIsValid = mobjGenGrid.CurrentCellIsValid
End Function


'Public Property Get Mouth() As MouthBL
'    Set Mouth = mobjMouth
'End Property
'Public Property Let Mouth(objMouth As MouthBL)
'    Set mobjMouth = objMouth
'    Set mobjGenGrid.Mouth = objMouth
'End Property

Public Property Get ViewType() As eViewType
    ViewType = menumViewType
End Property
Public Property Let ViewType(View As eViewType)
    menumViewType = View
    mobjGenGrid.ViewType = View
End Property

Friend Property Get Grid() As VSFlexGrid
    Set Grid = mctlGrid
End Property
Friend Property Set Grid(ctlGrid As VSFlexGrid)
    Set mctlGrid = ctlGrid
    Set mobjGenGrid.Grid = ctlGrid
    InitGrid menumViewType
End Property

Friend Function InitGrid(View As eViewType) As Boolean
    
    ViewType = View
   
   With mobjGenGrid
      InitGrid = .InitGrid
       
      mlngToothNameRow = .ToothNameRow
      mlngMissingRow = .MissingRow
      mlngCARow = .CARow
      mlngGMRow = .GMRow
      mlngProbDepthRow = .ProbDepthRow
      mlngSuppurationRow = .SuppurationRow
   End With
    
End Function


Public Sub KeyPress(KeyAscii As Integer)
    If mblnPlusKeyPending Or mblnMinusKeyPending Then  'currently doing a 10+ entry or negative number
    'If mblnPlusKeyPending Then  'currently doing a 10+ entry
        Select Case KeyAscii
            'If setting the value, immediately set mblnPlusKeyPending to false
            'so that leaving the cell does not cause the value to revert to 10
            Case vbKey1, vbKey2, vbKey3, vbKey4, vbKey5, vbKey6, vbKey7, vbKey7, vbKey8, vbKey9, vbKey0
                If mblnPlusKeyPending And mblnMinusKeyPending Then
                    SetCurrValue (-1 * (CInt(Chr(KeyAscii)) + 10))
                    AutoAdvance
                    mblnPlusKeyPending = False
                    mblnMinusKeyPending = False
                ElseIf mblnPlusKeyPending Then
                    SetCurrValue (CInt(Chr(KeyAscii)) + 10)
                    AutoAdvance
                    mblnPlusKeyPending = False
                Else 'minus key pending
                    SetCurrValue (-1 * CInt(Chr(KeyAscii)))
                    AutoAdvance
                    mblnMinusKeyPending = False
                End If
            Case Asc("+")
                If mblnPlusKeyPending And mblnMinusKeyPending Then
                    SetCurrValue (-20)
                    AutoAdvance
                    mblnPlusKeyPending = False
                    mblnMinusKeyPending = False
                ElseIf mblnPlusKeyPending Then
                    SetCurrValue (20)
                    AutoAdvance
                    mblnPlusKeyPending = False
                Else 'minus key pending
                    SetCurrValue ("-+")
                    mblnPlusKeyPending = True
                End If
            Case Else
                Beep
        End Select
        'mblnPlusKeyPending = False
    Else
        Select Case KeyAscii
            Case vbKey1, vbKey2, vbKey3, vbKey4, vbKey5, vbKey6, vbKey7, vbKey7, vbKey8, vbKey9, vbKey0
                SetCurrValue (Chr(KeyAscii))
                AutoAdvance
            Case Asc("P"), Asc("p")
                TogglePlaque
            Case Asc("B"), Asc("b"), Asc("/")
                ToggleBleeding
            Case Asc("S"), Asc("s"), Asc("*")
                ToggleSuppuration
            Case Asc("+")
                SetCurrValue ("+")
                mblnPlusKeyPending = True
            Case vbKeyRight
                MoveRight
            Case vbKeyLeft
                MoveLeft
            Case Asc(" ")
                SetCurrValue (" ")
                AutoAdvance
            Case Asc("-")
                If CurrentRowType = perGM Then 'only GM takes negative values
                    SetCurrValue ("-")
                    mblnMinusKeyPending = True
                Else
                    Beep
                End If
            Case vbKeyDelete
                SetCurrValue (" ")
            Case vbKeyBack
                SetCurrValue (" ")
                MoveLeft
            Case Else
                Beep
        End Select
    End If
End Sub

Private Sub SetCurrValue(strText As String)
    mobjGenGrid.SetCurrValue (strText)
End Sub

Private Function GetValue() As String
'    GetValue = mctlGrid.Text
    GetValue = mobjGenGrid.GetValue
End Function

Private Sub AutoAdvance()
    Select Case Me.CurrentRowType
        Case perProbDepth, perGM, perCA, perMGJ, perBoneLoss
            mobjGenGrid.MoveRight
        Case Else 'do not auto-advance
    End Select
End Sub
Private Sub MoveNext()
    mobjGenGrid.MoveRight
End Sub
Private Sub MovePrev()
    mobjGenGrid.MoveLeft
End Sub
Private Sub MoveLeft()
    mobjGenGrid.MoveLeft
End Sub
Private Sub MoveRight()
    mobjGenGrid.MoveRight
End Sub

Private Function AdvanceCells(lngCells As Long) As Boolean
    AdvanceCells = mobjGenGrid.AdvanceCells(lngCells)
End Function

Private Sub Class_Initialize()
    Set mobjGenGrid = New GenGridView
    mobjGenGrid.GridType = perBuccal
End Sub

Private Sub mctlGrid_EnterCell()
    'Make sure the cell is visible
    mobjGenGrid.MakeCurrentCellVisible
End Sub

Private Sub mctlGrid_LeaveCell()
    'Handle plus key entries
    If mblnPlusKeyPending Then
        If GetValue = "+" Then
            SetCurrValue (10)
        End If
    End If
    mblnPlusKeyPending = False
    
    'make probing values >= 4 red
    'mobjGenGrid.SetCurrentProbingDepthColor

End Sub

'Private Sub UpdateCA()
'    mobjGenGrid.UpdateCA
'End Sub
'Private Sub UpdateGM()
'    mobjGenGrid.UpdateGM
'End Sub

Private Function AddToothNameRow() As Long
    AddToothNameRow = mobjGenGrid.AddToothNameRow
End Function
Private Function AddMissingRow() As Long
    AddMissingRow = mobjGenGrid.AddMissingRow
End Function
Private Function AddCARow() As Long
    AddCARow = mobjGenGrid.AddCARow
End Function
Private Function AddGMRow() As Long
    AddGMRow = mobjGenGrid.AddGMRow
End Function
Private Function AddProbDepthRow()
    AddProbDepthRow = mobjGenGrid.AddProbDepthRow
End Function
Private Function InitCols() As Boolean
    InitCols = mobjGenGrid.InitCols
End Function

Public Sub SetCol(lngCol As Long)
    mctlGrid.Col = lngCol
End Sub
Public Sub SetRow(enumRow As eMeasurementType)
    mobjGenGrid.SetRow (enumRow)
End Sub

Private Sub TogglePlaque()
    mobjGenGrid.TogglePlaque
End Sub

Private Sub ToggleBleeding()
    mobjGenGrid.ToggleBleeding
End Sub

Private Sub ToggleSuppuration()
    mobjGenGrid.ToggleSuppuration
End Sub
Public Sub MakeCurrentCellVisible()
    mobjGenGrid.MakeCurrentCellVisible
End Sub

Private Sub mobjGenGrid_AdvanceCellsFail(Row As Long, Col As Long, ToRow As Long, ToCol As Long)
    RaiseEvent AdvanceCellsFail(Row, Col, ToRow, ToCol)
End Sub



Public Function CurrentRowType() As eMeasurementType
    CurrentRowType = mobjGenGrid.CurrentRowType
End Function

Public Function SetUnerupted(blnUnerupted As Boolean, Tooth As Integer)
    'Treat Unerupted same as missing
    If blnUnerupted Then
        mobjGenGrid.SetMissing blnUnerupted, Tooth
    End If
End Function
Public Function SetPontic(blnPontic As Boolean, Tooth As Integer)
    'Treat Pontic same as missing
    If blnPontic Then
        mobjGenGrid.SetMissing blnPontic, Tooth
    End If
End Function
Public Function SetMissing(blnMissing As Boolean, Tooth As Integer)
    If blnMissing Then
        mobjGenGrid.SetMissing blnMissing, Tooth
    End If
End Function

Public Function SetToothName(strName As String, Tooth As Integer)
    mobjGenGrid.SetToothName strName, Tooth
End Function

Public Function SetValue(strText As String, Tooth As Integer, Measurement As eMeasurementType, Location As eMeasurementLocation) As Boolean
    mobjGenGrid.SetValue strText, Tooth, Measurement, Location
End Function
Public Function SetValueFontColor(lngColor As Long, Tooth As Integer, Measurement As eMeasurementType, Location As eMeasurementLocation) As Boolean
    mobjGenGrid.SetValueFontColor lngColor, Tooth, Measurement, Location
End Function

Private Sub mobjGenGrid_TogglePrevValue(Measurement As PerioBL.eMeasurementType, Tooth As Integer, Location As PerioBL.eMeasurementLocation, Value As Variant)
    RaiseEvent TogglePrevValue(Measurement, Tooth, Location, Value)
End Sub

Private Sub mobjGenGrid_ValueSet(Measurement As PerioBL.eMeasurementType, Tooth As Integer, Location As PerioBL.eMeasurementLocation, Value As Variant)
    RaiseEvent ValueSet(Measurement, Tooth, Location, Value)
End Sub

Public Property Get GMRowVisible() As Boolean
    GMRowVisible = mblnGMRowVisible
End Property
Public Property Let GMRowVisible(blnVisible As Boolean)
    mblnGMRowVisible = blnVisible
    mctlGrid.RowHidden(mlngGMRow) = Not blnVisible
End Property
Public Property Get MGJRowVisible() As Boolean
    MGJRowVisible = mblnMGJRowVisible
End Property
Public Property Let MGJRowVisible(blnVisible As Boolean)
    mblnMGJRowVisible = blnVisible
    mctlGrid.RowHidden(mobjGenGrid.MgjRow) = Not blnVisible
End Property

Friend Sub Dispose()
    Set mctlGrid = Nothing
    
    mobjGenGrid.Dispose
    Set mobjGenGrid = Nothing
    
End Sub

Public Property Get EnableEdit() As Boolean
   EnableEdit = mblnEnableEdit
End Property
Public Property Let EnableEdit(blnEnable As Boolean)
   mblnEnableEdit = blnEnable
End Property

