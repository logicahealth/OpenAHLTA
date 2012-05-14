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
Attribute VB_Name = "SummaryReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Const ROWS As Integer = 21
Const COLS As Integer = 17

Const ROW_PD_B = 2
Const ROW_BLEED_B = 3
Const ROW_SUPP_B = 4
Const ROW_GM_B = 5
Const ROW_CA_B = 6
Const ROW_MGJ_B = 7
Const ROW_PLAQ_B = 8

Const ROW_PD_L = 10
Const ROW_BLEED_L = 11
Const ROW_SUPP_L = 12
Const ROW_GM_L = 13
Const ROW_CA_L = 14
Const ROW_MOB_L = 15
Const ROW_FURC_L = 16
Const ROW_MGJ_L = 17
Const ROW_PLAQ_L = 18
Const ROW_BL_L = 19
Const ROW_PROG_L = 20


Private WithEvents mctlPrinter As VSPrinter
Attribute mctlPrinter.VB_VarHelpID = -1
Private mobjExam As ExamBL
Private mstrHeader As String
Private mobjContainer As Object
Private mobjImage As IPictureDisp

Public Property Get Container() As Object
   Set Container = mobjContainer
End Property
Public Property Set Container(objContainer As Object)
   Set mobjContainer = objContainer
End Property

Public Property Get Header() As String
   Header = mstrHeader
End Property
Public Property Let Header(strHeader As String)
   mstrHeader = strHeader
End Property


Public Property Get Exam() As ExamBL
   Set Exam = mobjExam
End Property
Public Property Set Exam(objExam As ExamBL)
   Set mobjExam = objExam
End Property

Public Property Get SummaryPrinter() As VSPrinter
   Set SummaryPrinter = mctlPrinter
End Property
Public Property Set SummaryPrinter(ctlPrinter As VSPrinter)
   Set mctlPrinter = ctlPrinter
End Property

Public Property Get Image() As IPictureDisp
   Set Image = mobjImage
End Property
Public Property Set Image(objImage As IPictureDisp)
   Set mobjImage = objImage
End Property


Public Sub ShowSummary()
   InitPrinterControl
   DrawSummary
End Sub


Private Sub DrawSummary()
Dim eCurrentView As eViewType
'Dim blnHighlightCells As Boolean
Dim dtStart As Date
Dim lngNonMarginWidth As Long
Dim lngY As Long
Dim lngX As Long
Dim lngStartImageY As Long
Dim lngStartTableY As Long
Dim lngColor As OLE_COLOR


    Screen.MousePointer = vbHourglass

    With mctlPrinter
        .Preview = True
        
      .TablePen = 1
      .TablePenLR = 1
      .TablePenTB = 1
      .TableBorder = tbAll
        
        
        .StartDoc

        lngNonMarginWidth = .PageWidth - .MarginRight - .MarginLeft
        '.Header = PatientReportHeader() & "| |" & Format(mobjCurrExam.ExamDate, gsDefaultDateTime)
        .Header = mstrHeader
        .Footer = "||Page %d"

        .PenColor = vbBlack
        .PenStyle = psSolid
        .PenWidth = 50

         lngStartTableY = .CurrentY
         DrawTableData perMaxArch


'        'Draw the graphic
'        .NewPage
'        .Text = "Perio Graphic" & vbCrLf
'        ChangeViewType perFullMouth
'        DoEvents

         
        '.CurrentY = .CurrentY - 2000
        lngStartImageY = .CurrentY 'save x value of where we start drawing image
        .DrawPicture Me.Image, .MarginLeft, .CurrentY, , , , False
        '.CurrentY = .CurrentY + Me.Image.Height

        'Draw labels for the graphic
        'First save X/Y coords
        lngX = .CurrentX
        lngY = .CurrentY
        
        'Set Font
        .FontSize = 14
        
        'Apply offsets and write the labels
        .CurrentX = 500
        .CurrentY = .CurrentY + 500
        .Text = "F"

        .CurrentX = 500
        .CurrentY = .CurrentY + 1200
        .Text = "P"
        
        .CurrentX = 500
        .CurrentY = .CurrentY + 1800
        .Text = "L"

        .CurrentX = 500
        .CurrentY = .CurrentY + 1200
        .Text = "F"
        
        'Finally restory coords
        .CurrentX = lngX
        .CurrentY = lngY


        'Draw 2nd table
        .CurrentY = .CurrentY + 6000 'jump to the bottom of actual graphic
        lngY = .CurrentY
        .BrushColor = vbWhite
        .BrushStyle = bsSolid
        lngColor = .PenColor
        .PenColor = vbWhite
        'Draw rectangle to obscure any remainder of the graphic
        .DrawRectangle .MarginLeft, lngY, .PageWidth - .MarginRight, .CurrentY + (lngStartImageY - lngStartTableY)
         .PenColor = lngColor
         
        'reset current Y
        .CurrentY = lngY
        
        DrawTableData perManArch
        
        .EndDoc
        
        'Zoom to 100%
        '.ZoomMode = zmPercentage
        '.Zoom = 100

    End With

    Screen.MousePointer = vbNormal
End Sub




Private Sub InitPrinterControl()

      With mctlPrinter
         Set .Container = mobjContainer
         .Height = 8000
         .Left = 360
         .TabIndex = 45
         .Top = 360
         .Width = 6975
         .Appearance = 1
         .BorderStyle = 1
         .Enabled = True
         .MousePointer = 0
         .BackColor = -2147483643
         With .Font
            .Name = "Arial"
            .Size = 11.25
            .Charset = 0
            .Weight = 400
            .Underline = False
            .Italic = False
            .Strikethrough = False
         End With

         With .HdrFont
            .Name = "Courier New"
            .Size = 14.25
            .Charset = 0
            .Weight = 400
            .Underline = False
            .Italic = False
            .Strikethrough = False
         End With

         .AutoRTF = True
         .Preview = True
         .DefaultDevice = False
         .PhysicalPage = True
         .AbortWindow = True
         .AbortWindowPos = 0
         .AbortCaption = "Printing..."
         .AbortTextButton = "Cancel"
         .AbortTextDevice = "on the %s on %s"
         .AbortTextPage = "Now printing Page %d of"
         .FileName = ""
         .MarginLeft = 360 '1440
         .MarginTop = 1440
         .MarginRight = 360 '1440
         .MarginBottom = 1440
         .MarginHeader = 0
         .MarginFooter = 0
         .IndentLeft = 0
         .IndentRight = 0
         .IndentFirst = 0
         .IndentTab = 720
         .SpaceBefore = 0
         .SpaceAfter = 0
         .LineSpacing = 100
         .Columns = 1
         .ColumnSpacing = 180
         .ShowGuides = 2
         .LargeChangeHorz = 300
         .LargeChangeVert = 300
         .SmallChangeHorz = 30
         .SmallChangeVert = 30
         .Track = False
         .ProportionalBars = True
         .Zoom = 45.3598484848485
         .ZoomMode = 3
         .ZoomMax = 400
         .ZoomMin = 10
         .ZoomStep = 25
         .EmptyColor = -2147483636
         .TextColor = 0
         .HdrColor = 0
         .BrushColor = 0
         .BrushStyle = 0
         .PenColor = 0
         .PenStyle = 0
         .PenWidth = 0
         .PageBorder = 0
         .Header = ""
         .Footer = ""
         .TableSep = "|;"
         '.TableBorder = 7
         '.TablePen = 0
         '.TablePenLR = 0
         '.TablePenTB = 0
         .NavBar = 3
         .NavBarColor = -2147483633
         .ExportFormat = 0
         .URL = ""
         .Navigation = 3
         .NavBarMenuText = "Whole &Page|Page &Width|&Two Pages|Thumb&nail"

         .Visible = True
      End With
End Sub

Private Sub DrawTableData(enumArch As eViewType)
   Dim i As Integer
   Dim j As Integer
   Dim objMouth As MouthBL
   Dim objTooth As ToothBL
   
   Dim intOrigPenWidth As Integer
   Dim strValue As String
   
   On Error GoTo ErrHandler
   
   Set objMouth = Me.Exam.Mouth
   
   With mctlPrinter
      'start table definition
      
      .TableBorder = tbAll
      
      intOrigPenWidth = .PenWidth
      .PenWidth = 1
      .TablePenLR = 30
      .TablePenTB = 30
      .TablePen = 1
      
      
      .StartTable
      
      
      'Create table with 17 columns and X Rows
      .TableCell(tcCols) = COLS
      .TableCell(tcRows) = ROWS
      
      'Set the text for tooth labels
      For i = 1 To COLS
         If i = 1 Then
            
         Else
            If enumArch = perMaxArch Then
               .TableCell(tcText, 1, i) = objMouth.ToothById(i - 1).ToothName
               .TableCell(tcText, ROWS, i) = objMouth.ToothById(i - 1).ToothName
            Else
               .TableCell(tcText, 1, i) = objMouth.ToothById(33 - i + 1).ToothName
               .TableCell(tcText, ROWS, i) = objMouth.ToothById(33 - i + 1).ToothName
            End If
         End If
      Next i
      
      .TableCell(tcText, ROW_PD_B, 1) = "PD"
      .TableCell(tcText, ROW_PD_L, 1) = "PD"
      .TableCell(tcText, ROW_BLEED_B, 1) = "Bleed"
      .TableCell(tcText, ROW_BLEED_L, 1) = "Bleed"
      .TableCell(tcText, ROW_SUPP_B, 1) = "Supp"
      .TableCell(tcText, ROW_SUPP_L, 1) = "Supp"
      .TableCell(tcText, ROW_GM_B, 1) = "GM"
      .TableCell(tcText, ROW_GM_L, 1) = "GM"
      .TableCell(tcText, ROW_CA_B, 1) = "CA"
      .TableCell(tcText, ROW_CA_L, 1) = "CA"
      .TableCell(tcText, ROW_MGJ_B, 1) = "MGJ"
      .TableCell(tcText, ROW_MGJ_L, 1) = "MGJ"
      .TableCell(tcText, ROW_PLAQ_B, 1) = "Plaque"
      .TableCell(tcText, ROW_PLAQ_L, 1) = "Plaque"
      .TableCell(tcText, ROW_MOB_L, 1) = "Mobility"
      .TableCell(tcText, ROW_FURC_L, 1) = "Furc"
      .TableCell(tcText, ROW_BL_L, 1) = "BL"
      .TableCell(tcText, ROW_PROG_L, 1) = "Prog"
      
      For i = 2 To COLS
         If enumArch = perMaxArch Then
            Set objTooth = objMouth.ToothById(i - 1)
         Else
            Set objTooth = objMouth.ToothById(33 - i + 1)
         End If
         
         'PD
         .TableCell(tcText, ROW_PD_B, i) = GenericValue(objTooth, perProbDepth, perBuccal)
         
         .TableCell(tcText, ROW_PD_L, i) = GenericValue(objTooth, perProbDepth, perLingual)
         
         .TableCell(tcText, ROW_BLEED_B, i) = GenericValue(objTooth, perBleeding, perBuccal)
         .TableCell(tcText, ROW_BLEED_L, i) = GenericValue(objTooth, perBleeding, perLingual)
         
         .TableCell(tcText, ROW_SUPP_B, i) = GenericValue(objTooth, perSuppuration, perBuccal)
         .TableCell(tcText, ROW_SUPP_L, i) = GenericValue(objTooth, perSuppuration, perLingual)
         
         If objTooth.IsChartable Then
            .TableCell(tcText, ROW_GM_B, i) = GenericValue(objTooth, perGM, perBuccal)
            .TableCell(tcText, ROW_GM_L, i) = GenericValue(objTooth, perGM, perLingual)
         End If
         .TableCell(tcText, ROW_CA_B, i) = GenericValue(objTooth, perCA, perBuccal)
         .TableCell(tcText, ROW_CA_L, i) = GenericValue(objTooth, perCA, perLingual)
         
         'MGJ
         .TableCell(tcText, ROW_MGJ_B, i) = MGJValue(objTooth, perBuccal)
         .TableCell(tcText, ROW_MGJ_L, i) = MGJValue(objTooth, perLingual)
         
         .TableCell(tcText, ROW_PLAQ_B, i) = GenericValue(objTooth, perPlaque, perBuccal)
         .TableCell(tcText, ROW_PLAQ_L, i) = GenericValue(objTooth, perPlaque, perLingual)
         
         
         'Mobility
         .TableCell(tcText, ROW_MOB_L, i) = MobilityValue(objTooth)
         
         'Furcation
         .TableCell(tcText, ROW_FURC_L, i) = FurcationValue(objTooth)
         
         'BoneLoss
         .TableCell(tcText, ROW_BL_L, i) = BoneLossValue(objTooth)
         
         'Prognosis
         .TableCell(tcText, ROW_PROG_L, i) = PrognosisValue(objTooth)
         
      Next i
      
      'Format the cells
      .TableCell(tcFontSize, 1, 1, 1, COLS) = 10
      .TableCell(tcFontSize, ROWS, 1, ROWS, COLS) = 10
      
      .TableCell(tcFontSize, 2, 1, ROWS - 1, COLS) = 6
      
      .TableCell(tcAlign, 1, 1, ROWS, 1) = taLeftMiddle
      .TableCell(tcAlign, 1, 2, ROWS, COLS) = taCenterMiddle
      
      .TableCell(tcColWidth, 1, 1, ROWS, 1) = "675"
      .TableCell(tcColWidth, 1, 2, ROWS, COLS) = "675"
      
      .TableCell(tcFontBold, 1, 1, ROWS, 1) = True
      
      For i = 1 To ROWS
         Select Case i
            Case 1, 3, 5, 7, 11, 13, 15, 17, 21
               .TableCell(tcBackColor, i, 1, i, COLS) = vbCyan
         End Select
      Next i
      '.TableCell(tcBackColor, 3, 1, 3, COLS) = vbCyan
      
      .EndTable
      
      .PenWidth = intOrigPenWidth
   End With
    
   Exit Sub
   
ErrHandler:
    Err.Raise Err.Number, "SummaryReport.DrawTableData", Err.Description, Err.HelpFile, Err.HelpContext

End Sub



Private Function GenericValue(objTooth As ToothBL, eMeas As eMeasurementType, eGrid As eGridType) As String

   Dim strValue As String
   strValue = objTooth.Value(eMeas, eGrid, perLocOne) & " "
   strValue = strValue & objTooth.Value(eMeas, eGrid, perLocTwo) & " "
   strValue = strValue & objTooth.Value(eMeas, eGrid, perLocThree)


   If eMeas = perBleeding Then
      strValue = Replace(strValue, "False", " ")
      strValue = Replace(strValue, "True", "B")
   ElseIf eMeas = perSuppuration Then
      strValue = Replace(strValue, "False", " ")
      strValue = Replace(strValue, "True", "S")
   ElseIf eMeas = perPlaque Then
      strValue = Replace(strValue, "False", " ")
      strValue = Replace(strValue, "True", "P")
   End If
   
   GenericValue = strValue
End Function

Private Function MGJValue(objTooth As ToothBL, eGrid As eGridType) As String
   Dim strValue As String
   
   strValue = objTooth.MGJ(eGrid, perLocOne)
   
   MGJValue = strValue
End Function

Private Function MobilityValue(objTooth As ToothBL) As String
   Dim strValue As String
   
   strValue = objTooth.Mobility
   
   MobilityValue = strValue
End Function

Private Function FurcationValue(objTooth As ToothBL) As String

   Dim strValue As String
   Dim strToothName As String
   
   strToothName = objTooth.ToothName
   
    Select Case strToothName
        Case "1", "2", "3", "14", "15", "16", "A", "B", "I", "J"  '3 roots, 3 columns
            strValue = objTooth.Furcation(perLingual, perLocOne) & " "
            strValue = strValue & objTooth.Furcation(perLingual, perLocTwo) & " "
            strValue = strValue & objTooth.Furcation(perLingual, perLocThree)
        Case "5", "12"  '2 roots, 2 columns
             strValue = objTooth.Furcation(perLingual, perLocOne) & "  "
             strValue = strValue & objTooth.Furcation(perLingual, perLocTwo)
        
        Case "17", "18", "19", "30", "31", "32", "K", "L", "S", "T" '2 roots, 2 rows
             strValue = "B " & objTooth.Furcation(perLingual, perLocOne) & "  "
             strValue = strValue & "L " & objTooth.Furcation(perLingual, perLocTwo)
        Case "4", "6", "7", "8", "9", "10", "11", "13", "20", "21", _
                "22", "23", "24", "25", "26", "27", "28", "29", _
                "C", "D", "E", "F", "G", "H", "M", "N", "O", "P", "Q", "R" '1 root
             strValue = " "
        Case Else
            Debug.Assert False
    End Select
    
    FurcationValue = strValue
End Function

Private Function BoneLossValue(objTooth As ToothBL)
   Dim strValue As String
   Dim strBL As String
   
   strBL = objTooth.BoneLoss(perLingual, perLocOne)
   If Len(Trim(strBL)) = 0 Then
      strValue = Space(3)
   Else
      strValue = "M " & strBL & "  "
   End If
   
   strBL = objTooth.BoneLoss(perLingual, perLocTwo)
   If Len(Trim(strBL)) = 0 Then
      strValue = strValue & Space(3)
   Else
      strValue = strValue & "D " & strBL
   End If
   
   BoneLossValue = strValue
End Function

Private Function PrognosisValue(objTooth As ToothBL)
   Dim strValue As String
   Dim eProg As ePrognosis
   
   eProg = objTooth.Prognosis
   
   Select Case eProg
      Case ePrognosis.perExcellent
         strValue = "E"
      Case ePrognosis.perFair
         strValue = "F"
      Case ePrognosis.perGood
         strValue = "G"
      Case ePrognosis.perHopeless
         strValue = "H"
      Case ePrognosis.perPoor
         strValue = "P"
      Case ePrognosis.perQuestionable
         strValue = "Q"
      Case ePrognosis.perUnknown
         strValue = " "
   End Select
   
   PrognosisValue = strValue
End Function

Private Sub mctlPrinter_NewPage()
    Dim lngX As Long
    Dim lngY As Long
    Dim strFontSize As String
    Dim strWarning As String
    
    'Show warning for ITT Mode
    If gobjShared.IsAppMode(modeITT) Then
        With mctlPrinter
            lngX = .CurrentX
            lngY = .CurrentY
            strFontSize = .Font.Size
            
            .Font.Size = 20
            
            strWarning = "Training Purposes Only"
            
            'Center the text
            .CurrentX = (.PageWidth - .TextWidth(strWarning)) / 2
            .CurrentY = 100
            
            .Text = strWarning
            
            .CurrentX = lngX
            .CurrentY = lngY
            .Font.Size = strFontSize
            
        End With
    End If
    
    'SF Add the disclaimer to the report.
    AddDisclaimerToReport mctlPrinter

End Sub
