VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LingualGridView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mlngProbDepthRow As Long
Private mlngGMRow As Long
Private mlngCARow As Long
Private mlngBleedingRow As Long
Private mlngMobilityRow As Long
Private mlngFurcationRow As Long
Private mlngPrognosisRow As Long
Private mlngSuppurationRow As Long
Private mlngPlaqueRow As Long
Private WithEvents mctlGrid As VSFlexGrid
Attribute mctlGrid.VB_VarHelpID = -1
Private menumViewType As eViewType
Private mblnPlusKeyPending As Boolean
Private mblnMinusKeyPending 'Indicates a "-" key was pressed to make a value negative
Private mblnGMRowVisible As Boolean
Private mblnMGJRowVisible As Boolean
Private mlngDisabledColor As Long


Private WithEvents mobjGenGrid As GenGridView
Attribute mobjGenGrid.VB_VarHelpID = -1

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


Friend Property Get Grid() As VSFlexGrid
    Set Grid = mctlGrid
End Property
Friend Property Set Grid(ctlGrid As VSFlexGrid)
    Set mctlGrid = ctlGrid
    Set mobjGenGrid.Grid = ctlGrid
    
    InitGrid menumViewType
End Property
Public Property Get ViewType() As eViewType
    ViewType = menumViewType
End Property
Public Property Let ViewType(enumViewType As eViewType)
    menumViewType = enumViewType
    mobjGenGrid.ViewType = enumViewType
End Property

Friend Function InitGrid(View As eViewType) As Boolean
    ViewType = View
    
    InitGrid = mobjGenGrid.InitGrid
    
    With mobjGenGrid
        mlngCARow = .CARow
        mlngGMRow = .GMRow
        mlngProbDepthRow = .ProbDepthRow
        mlngBleedingRow = .BleedingRow
        mlngMobilityRow = .MobilityRow
        mlngFurcationRow = .FurcationRow
        mlngPrognosisRow = .PrognosisRow
        mlngSuppurationRow = .SuppurationRow
        mlngPlaqueRow = .PlaqueRow
    End With
    
End Function

Public Sub KeyPress(KeyAscii As Integer)
    
    If mblnPlusKeyPending Or mblnMinusKeyPending Then  'currently doing a 10+ entry or negative number
    'If mblnPlusKeyPending Then  'currently doing a 10+ entry
        Select Case KeyAscii
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
            Case Asc("-")
                If CurrentRowType = perGM Then 'only GM takes negative values
                    SetCurrValue ("-")
                    mblnMinusKeyPending = True
                Else
                    Beep
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
            Case Asc("B"), Asc("b"), Asc("/")
                ToggleBleeding
            Case Asc("S"), Asc("s"), Asc("*")
                ToggleSuppuration
            Case Asc("P"), Asc("p")
                TogglePlaque
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
                MoveRight

            'Prognosis Keys
            Case Asc("E"), Asc("e"), Asc("G"), Asc("g"), Asc("F"), Asc("f"), Asc("Q"), Asc("q"), Asc("P"), Asc("p"), Asc("H"), Asc("h")
                If mctlGrid.Row = mlngPrognosisRow Then
                    SetCurrValue UCase(Chr(KeyAscii))
                    'SetPrognosis KeyAscii
                    'Having problems setting autoadvance on this row.  Works first time, but
                    'isn't in the proper mode to select a value the 2nd time
                    'AutoAdvance
                End If
            Case Else
                Beep
        End Select
    End If
    
End Sub

Private Sub SetPrognosis(KeyAscii As Integer)
Dim lngCol As Long
Dim strProgValue As String
Dim strTemp As String

    strProgValue = ""
    
    Select Case mctlGrid.Row
        Case mlngPrognosisRow
            
            'lngCol = mctlGrid.Col
            lngCol = Int(mctlGrid.Col / 4) * 4 + 1
            
            strTemp = UCase(Chr(KeyAscii))
            Select Case strTemp
                Case "E", "G", "F", "Q", "P", "H"
                    strProgValue = strTemp
            End Select
                
            'mctlGrid.TextMatrix(mlngPrognosisRow, lngCol) = strProgValue
            'mctlGrid.TextMatrix(mlngPrognosisRow, lngCol + 1) = strProgValue
            'mctlGrid.TextMatrix(mlngPrognosisRow, lngCol + 2) = strProgValue
            mctlGrid.Cell(flexcpText, mlngPrognosisRow, lngCol, mlngPrognosisRow, lngCol + 2) = strProgValue
                
        Case Else
    End Select
End Sub

Private Sub ToggleBleeding()
    mobjGenGrid.ToggleBleeding
End Sub

Private Sub ToggleSuppuration()
    mobjGenGrid.ToggleSuppuration
'Dim strSuppuration As String
'Dim lngPrevCol As Long
'
'    Select Case mctlGrid.Row
'
'        Case mlngGMRow, mlngCARow, mlngProbDepthRow  'apply suppuration to previous tooth location
'            'get the column for the previous tooth location
'            lngPrevCol = mctlGrid.Col + 1
'            If lngPrevCol Mod 4 = 0 Then
'                lngPrevCol = lngPrevCol + 1
'            End If
'
'            strSuppuration = mctlGrid.TextMatrix(mlngSuppurationRow, lngPrevCol)
'            If UCase(strSuppuration) = "S" Then
'                mctlGrid.TextMatrix(mlngSuppurationRow, lngPrevCol) = ""
'            Else
'                mctlGrid.TextMatrix(mlngSuppurationRow, lngPrevCol) = "S"
'            End If
'
'        Case mlngSuppurationRow 'apply bleeding to current tooth location
'            If UCase(mctlGrid.Text) = "S" Then
'                mctlGrid.Text = ""
'            Else
'                mctlGrid.Text = "S"
'            End If
'    End Select
End Sub

Private Sub SetCurrValue(strText As String)
    mobjGenGrid.SetCurrValue (strText)
End Sub


Private Function GetValue() As String
    GetValue = mobjGenGrid.GetValue
'Dim strValue As String
'    strValue = Trim(mctlGrid.Text)
'    If CurrentRowType = perBoneLoss Then
'        'Strip off the "M/D"
'        strValue = Right(strValue, Len(strValue) - 1)
'    End If
'    GetValue = strValue
End Function

Private Sub AutoAdvance()
    Select Case Me.CurrentRowType
        Case perProbDepth, perGM, perCA, perMGJ
            mobjGenGrid.MoveLeft
        Case Else 'do not auto-advance
    End Select
End Sub
Private Sub MoveNext()
    mobjGenGrid.MoveLeft
End Sub

Private Sub MovePrev()
    mobjGenGrid.MoveRight
End Sub
Private Sub MoveRight()
    mobjGenGrid.MoveRight
End Sub
Private Sub MoveLeft()
    mobjGenGrid.MoveLeft
End Sub

Private Function AdvanceCells(lngCells As Long) As Boolean
    AdvanceCells = mobjGenGrid.AdvanceCells(lngCells)
End Function

Private Sub Class_Initialize()
    Set mobjGenGrid = New GenGridView
        mobjGenGrid.GridType = perLingual
        mlngDisabledColor = mobjGenGrid.DisabledColor
End Sub

Private Sub mctlGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Row
        Case mlngPrognosisRow
            'Set the Prognosis Row Combo drop down
            mctlGrid.ComboList = "Excellent|Good|Fair|Questionable|Poor|Hopeless"
        Case Else
    End Select
End Sub

Private Sub mctlGrid_ChangeEdit()

    'SF
    'Only invoke the validate edit method if the column is a dropdown list box
    With mctlGrid
        If (.ComboList <> "" And .Text <> "") Or (.ColComboList(.Col) <> "" And .Text <> "") Then
            .Select 0, 0
        End If
    End With

End Sub


Private Sub mctlGrid_EnterCell()
    'Make sure the cell is visible
    mobjGenGrid.MakeCurrentCellVisible

   With mctlGrid
    'Only make the grid editable if selecting a combo-list row
    Select Case .Row
        Case mlngPrognosisRow
            If EnableEdit And (.Cell(flexcpBackColor, mlngPrognosisRow, .Col) <> mlngDisabledColor) Then
                .Editable = flexEDKbdMouse
            Else
                .Editable = flexEDNone
            End If
        Case Else
            .Editable = flexEDNone
    End Select
   End With
End Sub

Private Sub mctlGrid_LeaveCell()
Dim strText As String
Dim lngCol As Long
Dim lngRow As Long
Dim intCurrTooth As Integer

    'Handle plus key entries
    If mblnPlusKeyPending Then
        If GetValue = "+" Then
            SetCurrValue (10)
        End If
    End If
    mblnPlusKeyPending = False
    
    'make probing values >= 4 red
'    mobjGenGrid.SetCurrentProbingDepthColor

        
    With mctlGrid
        lngRow = .Row
        lngCol = .Col
        Select Case lngRow
            'Only display the first character of the Prognosis
            Case mlngPrognosisRow
                strText = .Text
                If strText = "" Then 'if user doesn't select an item in the list, it comes back empty
                    strText = " "
                End If
                If Len(strText) <> 0 Then
                    strText = Left(strText, 1)
                    lngCol = Int(lngCol / 4) * 4 + 1

                    '.TextMatrix(lngRow, lngCol) = strText
                    '.TextMatrix(lngRow, lngCol + 1) = strText
                    '.TextMatrix(lngRow, lngCol + 2) = strText
                        
                    intCurrTooth = .TextMatrix(0, .Col)
                    'Really should be using the SetValue call which would in turn raise the
                    'event.  Then the .textMatrix statements above would not be needed
                    SetValue strText, intCurrTooth, perPrognosis, perLocWholeTooth
                    'RaiseEvent ValueSet(perPrognosis, intCurrTooth, perLocWholeTooth, strText)
                End If
        End Select
    End With

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
Private Function AddCARow() As Long
    AddCARow = mobjGenGrid.AddCARow
End Function
Private Function AddGMRow() As Long
    AddGMRow = mobjGenGrid.AddGMRow
End Function
Private Function AddProbDepthRow()
    AddProbDepthRow = mobjGenGrid.AddProbDepthRow
End Function
Private Function AddBleedingRow()
    AddBleedingRow = mobjGenGrid.AddBleedingRow
End Function
Private Function AddPrognosisRow()
    AddPrognosisRow = mobjGenGrid.AddPrognosisRow
End Function
Private Function AddMobilityRow()
    AddMobilityRow = mobjGenGrid.AddMobilityRow
End Function
Private Function AddFurcationRow()
    AddFurcationRow = mobjGenGrid.AddFurcationRow
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

Private Sub mobjGenGrid_AdvanceCellsFail(Row As Long, Col As Long, ToRow As Long, ToCol As Long)
    RaiseEvent AdvanceCellsFail(Row, Col, ToRow, ToCol)
End Sub

Public Function CurrentRowType() As eMeasurementType
    CurrentRowType = mobjGenGrid.CurrentRowType
End Function
Public Sub MakeCurrentCellVisible()
    mobjGenGrid.MakeCurrentCellVisible
End Sub

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
Public Function SetToothName(strName As String, Tooth As Integer)
    mobjGenGrid.SetToothName strName, Tooth
End Function
Public Function SetMissing(blnMissing As Boolean, Tooth As Integer)
    If blnMissing Then
        mobjGenGrid.SetMissing blnMissing, Tooth
    End If
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

