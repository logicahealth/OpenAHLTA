VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSVGDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mObjSVGLib As CHCSII_Dental_SVG_Lib.SVG_Lib

'This has been replaced by the respective Dx or Tx
'TODO: don't see why this is public yet
'Public mobjdataelement As clsDataElem
Public mObjDataProc As clsDataProc

Private mObjMenuLib As CHCSII_Dntl_Mnu_Lib.Mnu_Lib
Private mOSF As New CHCSII_Dntl_Mnu_Lib.SF

Public mColTeeth As Collection

Public Event MissingTooth(ToothID As Integer)

Public Function patternSurface(ToothID As Integer, surfaceID As Integer, eRGB As eColor, symbol As eSymbol, Optional findText As String, Optional replaceText As String) As String
Dim objSurface As CHCSII_Dental_SVG_Lib.clsSVGSurface
Dim objDxTxTooth As CHCSII_Dental_SVG_Lib.clsSVGDxTx
Dim strFind As String
Dim strReplace As String
Dim strFindImplant As String
Dim strReplaceImplant As String
Dim strReplaceText As String
Dim strColor As String
Dim strSymbol As String
Dim strFindAnk As String

On Error GoTo ErrorHandler

    
'SCR 36412
If symbol = eSuperNumerary4 Then
    strFind = "OutlineRedSm"
Else
    strFind = "NoFill"
    strFindImplant = "#implant"
End If
  
  Select Case eRGB
     Case eColor.eRed
         strReplace = "FillRed"
         strReplaceImplant = "#implantRed"
         strColor = "Red"
         'SCR 36412
         If symbol = eSuperNumerary4 Then
            strReplace = "OutlineRedSm"
         End If
     Case eColor.eGreen
         strReplace = "FillGreen"
         strReplaceImplant = "#implantGreen"
         strColor = "Green"
         
         'SCR 36412
         If symbol = eSuperNumerary4 Then
            strReplace = "OutlineGreenSm"
         End If
     Case eColor.eBlue
         strReplace = "FillBlue"
         strReplaceImplant = "#implantBlue"
         strColor = "Blue"
     
     Case eColor.eWhite
        'SCR 36412
         If symbol = eSuperNumerary4 Then
            strReplace = "NoFill"
        End If
End Select
    

      strSymbol = CStr(symbol)
      
      If symbol = 15 Then
          replaceText = ">2<"
          findText = ">1<"
          strSymbol = "14"
      ElseIf symbol = 16 Then
          findText = ">1<"
          replaceText = ">3<"
          strSymbol = "14"
      End If
      
      If (symbol > 8 And symbol < 12) Or symbol = 3 Then
         strSymbol = "25"
      End If
      
      'Enables Sealant to be used :Perf:  Only do once !
      If symbol = eTextS Or symbol = eLineBox Then
         surfaceID = 777
      End If
'      ' Because of a DB error a temporary fix has been placed for tooth 32 surfaces 24,26
'      If toothID = 32 Then
'         If surfaceID = 26 Then
'            surfaceID = 24
'         ElseIf surfaceID = 24 Then
'            surfaceID = 26
'         End If
'      End If
      
      If Not surfaceID = 777 Then
        Set objSurface = mObjSVGLib.SVGSurfaces(CStr(ToothID) & "|" & CStr(surfaceID))
      End If
      
      Set objDxTxTooth = mObjSVGLib.SVGDxsTxs(CStr(ToothID) & "|" & CStr(surfaceID) & "|" & strSymbol)
 
    If objDxTxTooth Is Nothing Then
        Exit Function 'Nothing to do
    End If
      
      patternSurface = objDxTxTooth.DxTxSVG
      patternSurface = Replace(patternSurface, strFind, strReplace)
      
      'Because the resource file is wrong for Ank
      If strSymbol = "12" And Not ToothID = 1 Then
      
        patternSurface = Left(patternSurface, InStr(patternSurface, "><g")) & "</g>"
      
      End If
      If Len(replaceText) > 0 And Len(findText) > 0 Then
         patternSurface = Replace(patternSurface, findText, replaceText)
      End If

   Select Case symbol
''     Case eSymbol.ePolkDots
''       If eRGB = eRed Then
''          hbrush = CreatePatternBrush(imgBrushes.ListImages.Item("PolkaRed").Picture)
''       ElseIf eRGB = eGreen Then
''          hbrush = CreatePatternBrush(imgBrushes.ListImages.Item("PolkaGreen").Picture)
''       ElseIf eRGB = eBlue Then
''          hbrush = CreatePatternBrush(imgBrushes.ListImages.Item("PolkaBlue").Picture)
''       End If
''     Case eSymbol.eTextS
''       textOnOcclusal surfaceID, eRGB, "S"
''     Case eSymbol.eTextO
''       textOnSurface surfaceID, eRGB, "O"
''     Case eSymbol.eTextNF
''       textOnSurface surfaceID, eRGB, "NF"
     Case eSymbol.eTextW
        strFind = "SymbolText"
         strReplaceText = "Wear" & strColor
         patternSurface = Replace(patternSurface, strFind, strReplaceText)
     Case eSymbol.eTextC
         strFind = "SymbolText"
         strReplaceText = "Contact" & strColor
         patternSurface = Replace(patternSurface, strFind, strReplaceText)
     Case eSymbol.eTextE
         strFind = "SymbolText"
         strReplaceText = "Esthetic" & strColor
         patternSurface = Replace(patternSurface, strFind, strReplaceText)
''     Case eSymbol.eTextANK
'     Case eSymbol.eSuperNumerary4
'         patternSurface = objDxTxTooth.DxTxSVG
''     Case eSymbol.eLine
''       'fillCircleTooth surfaceID, eRGB, 20, 20
''       rectOnSurface surfaceID, eRGB, xSize, ySize, xPosition, yPosition
''     Case eSymbol.eOutlineCircle
''       circleTooth eRGB, CInt(Height), CInt(Width)
''     Case eSymbol.eHorizontalScroll1
''
     Case eSymbol.eImplant
       patternSurface = Replace(patternSurface, strFindImplant, strReplaceImplant)

''       implantTooth eRGB
     Case eSymbol.eForwardSlash
         strFind = "SymbolText"
         strReplaceText = "Fracture" & strColor
         patternSurface = Replace(patternSurface, strFind, strReplaceText)
'     Case eSymbol.eParallelLines
'       patternSurface = objDxTxTooth.DxTxSVG
''       parallelLines eRGB
''     Case eSymbol.eFilledArc
''       filledArcTooth surfaceID, eRGB, 10, 10
''     Case eSymbol.eTriangle
''       filledTriTooth surfaceID, eRGB, 10, 10
''     Case eSymbol.eLineBox
''       lineBox surfaceID, eRGB
''     Case Else
''          hbrush = CreatePatternBrush(imgBrushes.ListImages.Item("PolkaRed").Picture)
   End Select
Exit Function
    
ErrorHandler:
    If Err.Number <> 5 Then
        MsgBxARMd Err.Number & "," & Err.Description & ",clsSVGDraw.patternSurface, SVGManager"
    End If
    Resume Next
    
End Function
Public Function hatchSurface(ToothID As Integer, surfaceID As Integer, eRGB As eColor, eCHatch As eHatch) As String
On Error GoTo errHandler
 
  Dim strFill As String
  Dim objSurface As CHCSII_Dental_SVG_Lib.clsSVGSurface
  Dim GCTooth As CHCSII_Dental_SVG_Lib.clsSVGToothDef
  Dim strSurfaceFind As String
  Dim strSurfaceReplace As String
  Dim strFillType As String
  
  strSurfaceFind = "SS-S-" & CStr(surfaceID) '& LCase(LoadResString(200 + surfaceID))
  strSurfaceReplace = "SS-S-" & CStr(surfaceID) & "-dx"
      
      If Not surfaceID = 111 Then
         Set objSurface = mObjSVGLib.SVGSurfaces(CStr(ToothID) & "|" & CStr(surfaceID))
         strFill = objSurface.Surface
         strFillType = "NoFill"
      Else
         Set GCTooth = mObjSVGLib.SVGTeethDef(ToothID)
         strFill = GCTooth.getToothCrownBody
         strSurfaceFind = "CR-T-" & ToothID
         strSurfaceReplace = "SS-S-" & CStr(surfaceID) & "-dx-" & "T-" & ToothID
         strFillType = "Outline"
         
      End If
      Select Case eRGB
      
        Case eColor.eRed
        
           Select Case eCHatch
                Case eHatch.eBothDiagRightLeft
                     hatchSurface = LoadResString(509)
                     hatchSurface = hatchSurface & Replace(strFill, strFillType, "FillSurfaceDiamondRed")
                     hatchSurface = Replace(hatchSurface, strSurfaceFind, strSurfaceReplace)
               
                Case eHatch.eBothVerticalHorizontal
                     hatchSurface = LoadResString(506)
                     hatchSurface = hatchSurface & Replace(strFill, strFillType, "FillSurfaceCrossRed")
                     hatchSurface = Replace(hatchSurface, strSurfaceFind, strSurfaceReplace)

                Case eHatch.eDiagLeft
                Case eHatch.eDiagRight
                     hatchSurface = LoadResString(503)
                     hatchSurface = hatchSurface & Replace(strFill, strFillType, "FillSurfaceDiagRightRed")
                     hatchSurface = Replace(hatchSurface, strSurfaceFind, strSurfaceReplace)
                Case eHatch.eHorizontal
                     hatchSurface = LoadResString(500)
                     hatchSurface = hatchSurface & Replace(strFill, strFillType, "FillSurfaceHorizRed")
                     hatchSurface = Replace(hatchSurface, strSurfaceFind, strSurfaceReplace)
                
                Case eHatch.eVertical
                   
           End Select
           
        Case eColor.eBlue
           Select Case eCHatch
                Case eHatch.eBothDiagRightLeft
                     hatchSurface = LoadResString(511)
                     hatchSurface = hatchSurface & Replace(strFill, strFillType, "FillSurfaceDiamondBlue")
                     hatchSurface = Replace(hatchSurface, strSurfaceFind, strSurfaceReplace)
                
                Case eHatch.eBothVerticalHorizontal
                     hatchSurface = LoadResString(507)
                     hatchSurface = hatchSurface & Replace(strFill, strFillType, "FillSurfaceCrossBlue")
                     hatchSurface = Replace(hatchSurface, strSurfaceFind, strSurfaceReplace)

                Case eHatch.eDiagLeft
                Case eHatch.eDiagRight
                     hatchSurface = LoadResString(504)
                     hatchSurface = hatchSurface & Replace(strFill, strFillType, "FillSurfaceDiagRightBlue")
                     hatchSurface = Replace(hatchSurface, strSurfaceFind, strSurfaceReplace)
                 Case eHatch.eHorizontal
                     hatchSurface = LoadResString(501)
                     hatchSurface = hatchSurface & Replace(strFill, strFillType, "FillSurfaceHorizBlue")
                     hatchSurface = Replace(hatchSurface, strSurfaceFind, strSurfaceReplace)
                Case eHatch.eVertical
                   
           End Select
        Case eColor.eGreen
           Select Case eCHatch
                Case eHatch.eBothDiagRightLeft
                     hatchSurface = LoadResString(510)
                     hatchSurface = hatchSurface & Replace(strFill, strFillType, "FillSurfaceDiamondGreen")
                     hatchSurface = Replace(hatchSurface, strSurfaceFind, strSurfaceReplace)

                Case eHatch.eBothVerticalHorizontal
                     hatchSurface = LoadResString(508)
                     hatchSurface = hatchSurface & Replace(strFill, strFillType, "FillSurfaceCrossGreen")
                     hatchSurface = Replace(hatchSurface, strSurfaceFind, strSurfaceReplace)

                Case eHatch.eDiagLeft
                Case eHatch.eDiagRight
                     hatchSurface = LoadResString(505)
                     hatchSurface = hatchSurface & Replace(strFill, strFillType, "FillSurfaceDiagRightGreen")
                     hatchSurface = Replace(hatchSurface, strSurfaceFind, strSurfaceReplace)
 
                Case eHatch.eHorizontal
                     hatchSurface = LoadResString(502)
                     hatchSurface = hatchSurface & Replace(strFill, strFillType, "FillSurfaceHorizGreen")
                     hatchSurface = Replace(hatchSurface, strSurfaceFind, strSurfaceReplace)
                Case eHatch.eVertical
                   
           End Select
      End Select
       
CleanUp:
  
  Exit Function
errHandler:
  Err.Clear
  Screen.MousePointer = vbDefault
End Function
Public Function outlineSurface(ToothID As Integer, surfaceID As Integer, eRGB As eColor, Optional clearSurface As Boolean = True) As String
 Dim strFill As String
 Dim strFillType As String
  Dim objSurface As CHCSII_Dental_SVG_Lib.clsSVGSurface
  Dim GCTooth As CHCSII_Dental_SVG_Lib.clsSVGToothDef
  Dim strSurfaceFind As String
  Dim strSurfaceReplace As String
  
  strSurfaceFind = "SS-S-" & CStr(surfaceID) '& LCase(LoadResString(200 + surfaceID))
  strSurfaceReplace = "SS-S-" & CStr(surfaceID) & "-dx"

'  strSurfaceFind = "surface-" & LCase(LoadResString(200 + surfaceID))
'  strSurfaceReplace = "surface-" & LCase(LoadResString(200 + surfaceID)) & "-dx"
'
       If Not surfaceID = 111 Then
         Set objSurface = mObjSVGLib.SVGSurfaces(CStr(ToothID) & "|" & CStr(surfaceID))
         strFill = objSurface.Surface
         strFillType = "NoFill"
      Else
         Set GCTooth = mObjSVGLib.SVGTeethDef(ToothID)
         strFill = GCTooth.getToothCrownBody
         strSurfaceFind = "CR-T-" & ToothID
         strSurfaceReplace = "SS-S-" & CStr(surfaceID) & "-dx-" & "T-" & ToothID
         strFillType = "Outline"
         
      End If
      
      Select Case eRGB
      
        Case eColor.eRed
           outlineSurface = Replace(strFill, strFillType, "OutlineRedMed")
           outlineSurface = Replace(outlineSurface, strSurfaceFind, strSurfaceReplace)
        Case eColor.eBlue
           outlineSurface = Replace(strFill, strFillType, "OutlineBlueMed")
          outlineSurface = Replace(outlineSurface, strSurfaceFind, strSurfaceReplace)
        Case eColor.eGreen
           outlineSurface = Replace(strFill, strFillType, "OutlineGreenMed")
           outlineSurface = Replace(outlineSurface, strSurfaceFind, strSurfaceReplace)
      End Select
      
End Function
Public Function fillSurface(ToothID As Integer, surfaceID As Integer, eRGB As eColor) As String
  Dim strFill As String
  Dim objSurface As CHCSII_Dental_SVG_Lib.clsSVGSurface
  Dim GCTooth As CHCSII_Dental_SVG_Lib.clsSVGToothDef
  Dim strSurfaceFind As String
  Dim strSurfaceReplace As String
  Dim strFillType As String

  strSurfaceFind = "SS-S-" & CStr(surfaceID) '& LCase(LoadResString(200 + surfaceID))
  strSurfaceReplace = "SS-S-" & CStr(surfaceID) & "-dx"

      If Not surfaceID = 111 Then
         Set objSurface = mObjSVGLib.SVGSurfaces(CStr(ToothID) & "|" & CStr(surfaceID))
         strFill = objSurface.Surface
         strFillType = "NoFill"
      Else
         Set GCTooth = mObjSVGLib.SVGTeethDef(ToothID)
         strFill = GCTooth.getToothCrownBody
         strSurfaceFind = "CR-T-" & ToothID
         strSurfaceReplace = "SS-S-" & CStr(surfaceID) & "-dx-" & "T-" & ToothID
         strFillType = "Outline"
         
      End If

      Select Case eRGB
      
        Case eColor.eRed
           fillSurface = Replace(strFill, strFillType, "FillRed")
           fillSurface = Replace(fillSurface, strSurfaceFind, strSurfaceReplace)
        Case eColor.eBlue
           fillSurface = Replace(strFill, strFillType, "FillBlue")
           fillSurface = Replace(fillSurface, strSurfaceFind, strSurfaceReplace)
        Case eColor.eGreen
           fillSurface = Replace(strFill, strFillType, "FillGreen")
           fillSurface = Replace(fillSurface, strSurfaceFind, strSurfaceReplace)
        Case eColor.eBlack
           fillSurface = Replace(strFill, strFillType, "FillBlack")
           fillSurface = Replace(fillSurface, strSurfaceFind, strSurfaceReplace)

      End Select
      
End Function
Public Sub DrawSVG(DxCollection As Collection, TxCollection As Collection, GCTooth As clsGCTooth, Optional objDLDx As CHCSII_Dental_DXTX.clsDx, Optional objDLTx As CHCSII_Dental_DXTX.clsTx)

Dim iToothID As Integer
Dim oToothDLDxs As colToothDxs
Dim oToothDLTxs As colToothTxs
Dim oDx As CHCSII_Dental_DXTX.clsDx
Dim oTx As CHCSII_Dental_DXTX.clsTx

On Error GoTo ErrorHandler
   
    If mObjDataProc.TeethDxData.Count > 0 Then
        'drawSVGByTooth DxCollection, TxCollection, mObjDataProc.TeethDxData.Item(GCTooth.ToothID), mObjDataProc.TeethTxData.Item(GCTooth.ToothID), GCTooth, objDLDx, objDLTx
        
        'Get all the DLDxs for the tooth
        Set oToothDLDxs = mObjDataProc.TeethDxData.Item(GCTooth.ToothID)
        
        'Get all the DLTxs for the tooth
        Set oToothDLTxs = mObjDataProc.TeethTxData.Item(GCTooth.ToothID)
        
        'Get the tooth id of the tooth
        iToothID = GCTooth.ToothID
        
        For Each oDx In oToothDLDxs
            If oDx.GCIsDirty = True Then
                'The Dx is dirty so draw it
                If oDx.DxICDCodeID <> 766 Then 'Don't draw Orientation
                    DrawDxSVGByTooth oDx, DxCollection, GCTooth
                Else
                    'The tooth has an orientation dx so set the HasDisplacement flag
                    GCTooth.HasDisplacement = True
                End If
                'We have drawn the tooth so set the GCIsDirty flag to false
                oDx.GCIsDirty = False
            End If
        Next
        
        For Each oTx In oToothDLTxs
            If oTx.GCIsDirty = True Then
                'The Tx is dirty so draw it
                DrawTxSVGByTooth oTx, TxCollection, GCTooth
                'We have drawn the tooth so set the GCIsDirty flag to false
                oTx.GCIsDirty = False
            End If
        Next
        
    '    For Each mobjdataelement In colTeethDxData
    '
    '        If mobjdataelement.IsDirty = True Then
    '            If Not mobjdataelement.Dx Is Nothing Then
    '                If Not mobjdataelement.Dx.DxICDCodeID = 766 Then 'Orientation
    '                    drawDxSVGByTooth mobjdataelement, DxCollection, GCTooth
    '                Else
    '                    GCTooth.HasDisplacement = True
    '                End If
    '            ElseIf Not mobjdataelement.Tx Is Nothing Then
    '                drawTxSVGByTooth mobjdataelement, TxCollection, GCTooth
    '            End If
    '
    '            mobjdataelement.IsDirty = False
    '        End If
    '
    '    Next
   
   End If

Exit Sub
    
ErrorHandler:
    MsgBxARMd Err.Number & "," & Err.Description & ",clsSVGDraw.drawSVG, SVGManager"
    Resume Next

End Sub

Private Sub DrawDxSVGByTooth(DLDx As CHCSII_Dental_DXTX.clsDx, DxCollection As Collection, GCTooth As clsGCTooth)

Dim j, k As Integer
Dim iToothID As Integer
Dim sMenuID As String
'Dim objPatDx As patDx 'Not used
'Dim objDlDx As CHCSII_Dental_DXTX.clsDx
Dim strArrSurfaceSet() As String
Dim strArrTeethSet() As String
Dim objMenuInstruction As clsMenuItemInstr
Dim oSurfaceGroup As clsSurfaceGroup
'Dim objMacro As clsMacro
Dim DxStatus As DxStatusEnum

    'Get the ToothID
    iToothID = GCTooth.ToothID
    
    DxStatus = DLDx.StatusID
    
    'Get the menu id
    sMenuID = CStr(DLDx.SurfaceSets(1).SymbolID)
    Set objMenuInstruction = getMenuInstruction(sMenuID, DLDx.StatusID)
    
    'jrm 04/09/2002  set TreatRetreat code based on menu id
    setTreatRetreatToToothObject sMenuID, GCTooth, 0, DLDx.StatusID
    
    'Get the surfaces
    If DLDx.SurfaceSets(1).Teeth(1).ToothName = "Max" Or DLDx.SurfaceSets(1).Teeth(1).ToothName = "Mand" Then
       strArrSurfaceSet = Split("111", ",")
    Else
       strArrSurfaceSet = Split(DLDx.SurfaceSets(1).SurfaceSet, "|")
    End If
    
    Select Case objMenuInstruction.Operation
        
        Case eOperation.eFillSurface
            For j = 0 To UBound(strArrSurfaceSet)
                DxCollection.Add ConvertDLDxToGCDx(sMenuID, mObjMenuLib.MenuItems(sMenuID).MnuButtonID, iToothID, strArrSurfaceSet(j), fillSurface(CInt(iToothID), CInt(strArrSurfaceSet(j)), objMenuInstruction.Color), DLDx.StatusID, DLDx.ModifiedDateTime, objMenuInstruction, DLDx.DRC)
            Next j
            
        Case eOperation.eHatchSurface
            For j = 0 To UBound(strArrSurfaceSet)
                DxCollection.Add ConvertDLDxToGCDx(sMenuID, mObjMenuLib.MenuItems(sMenuID).MnuButtonID, iToothID, strArrSurfaceSet(j), hatchSurface(CInt(iToothID), CInt(strArrSurfaceSet(j)), objMenuInstruction.Color, objMenuInstruction.Hatch), DLDx.StatusID, DLDx.ModifiedDateTime, objMenuInstruction, DLDx.DRC)
            Next j
            
        Case eOperation.eSymbolizeSurface
            'Crown only surfaces reset to do all surfaces.
            If strArrSurfaceSet(0) = 111 Then
                 Set oSurfaceGroup = mObjSVGLib.SurfaceGroups(GCTooth.ToothID)
                 strArrSurfaceSet = Split(oSurfaceGroup.CrownSurfaces, ",")
            End If
            
            For j = 0 To UBound(strArrSurfaceSet)
                If objMenuInstruction.symbol = eSymbol.eImplant Then
                   GCTooth.IsImplant = True
                   GCTooth.DateTimeImplant = DLDx.ModifiedDateTime
                End If
                
                If objMenuInstruction.symbol = eParallelLines Then
                    GCTooth.DateTimeWholeToothPlanToRemove = DLDx.ModifiedDateTime
                End If
                
                'Mark type of whole tooth
                If objMenuInstruction.symbol = eSuperNumerary4 Then
                   GCTooth.DateTimeSupernumary = DLDx.ModifiedDateTime
                ElseIf objMenuInstruction.symbol = eOutlineCircleDashed Then
                   GCTooth.DateTimeErupted = DLDx.ModifiedDateTime
                ElseIf objMenuInstruction.symbol = eOutlineCircle Then
                   GCTooth.DateTimeUnErupted = DLDx.ModifiedDateTime
                End If
                
                If objMenuInstruction.symbol = eSymbol.eMissing Then
                      GCTooth.IsExtractOrMissing = True
                      'RaiseEvent MissingTooth(iToothID)
                End If
                    DxCollection.Add ConvertDLDxToGCDx(sMenuID, mObjMenuLib.MenuItems(sMenuID).MnuButtonID, iToothID, strArrSurfaceSet(j), patternSurface(CInt(iToothID), CInt(strArrSurfaceSet(j)), objMenuInstruction.Color, objMenuInstruction.symbol), DLDx.StatusID, DLDx.ModifiedDateTime, objMenuInstruction, DLDx.DRC)
            Next j
        
        Case eOperation.eOutlineSurface
            'Crown only surfaces reset to do all surfaces.
            If strArrSurfaceSet(0) = 111 Then
                Set oSurfaceGroup = mObjSVGLib.SurfaceGroups(GCTooth.ToothID)
                strArrSurfaceSet = Split(oSurfaceGroup.CrownSurfaces, ",")
            End If
            
            For j = 0 To UBound(strArrSurfaceSet)
                DxCollection.Add ConvertDLDxToGCDx(sMenuID, mObjMenuLib.MenuItems(sMenuID).MnuButtonID, iToothID, strArrSurfaceSet(j), outlineSurface(CInt(iToothID), CInt(strArrSurfaceSet(j)), objMenuInstruction.Color), DLDx.StatusID, DLDx.ModifiedDateTime, objMenuInstruction, DLDx.DRC)
            Next j
    
    End Select
    
    'Set tooth removeLocation for all occasions
    GCTooth.RemoveLocation = objMenuInstruction.RemovedLocations
    
    'Set Date/Time when location was removed
    If GCTooth.RemoveLocation = eLocation.eRootSection Then
        If DxStatus = dxActiveNew Or DxStatus = dxActiveOngoing Or DxStatus = dxActiveProvisional Then
                GCTooth.DateTimeRootRemoved = DLDx.ModifiedDateTime
        End If
           
    ElseIf GCTooth.RemoveLocation = eLocation.eCoronalSection Then
        'PCR 9981 'Only set the removed location property on the tooth if the Dx has an active status
        If DxStatus = dxActiveNew Or DxStatus = dxActiveOngoing Or DxStatus = dxActiveProvisional Then
            GCTooth.DateTimeCrownRemoved = DLDx.ModifiedDateTime
        End If
    
    ElseIf DLDx.SurfaceSets(1).SymbolID = 151 Then 'Missing Dx Menu Code
        GCTooth.DateTimeWholeToothRemoved = DLDx.ModifiedDateTime
    
    End If
 
End Sub
Private Sub DrawTxSVGByTooth(Tx As CHCSII_Dental_DXTX.clsTx, TxCollection As Collection, GCTooth As clsGCTooth)

Dim j As Integer
Dim iToothID As Integer
Dim sMenuID As String
Dim oGCTx As clsGCTx
Dim strArrSurfaceSet() As String
Dim objMenuInstruction As clsMenuItemInstr
Dim oSurfaceGroup As clsSurfaceGroup

    iToothID = GCTooth.ToothID
    sMenuID = CStr(Tx.SurfaceSets(1).SymbolID)
    
    If Tx.SurfaceSets(1).Teeth(1).ToothName = "Max" Or Tx.SurfaceSets(1).Teeth(1).ToothName = "Mand" Then
       strArrSurfaceSet = Split("111", ",")
    Else
       strArrSurfaceSet = Split(Tx.SurfaceSets(1).SurfaceSet, "|")
    End If
    
    Set objMenuInstruction = getMenuInstruction(sMenuID, Tx.StatusID)
    
    'jrm 04/09/2002  set TreatRetreat code based on menu id
    setTreatRetreatToToothObject sMenuID, GCTooth, Tx.StatusID
    
    Select Case objMenuInstruction.Operation
       
        Case eOperation.eFillSurface
            For j = 0 To UBound(strArrSurfaceSet)
              TxCollection.Add ConvertDLTxToGCTx(sMenuID, mObjMenuLib.MenuItems(sMenuID).MnuButtonID, iToothID, strArrSurfaceSet(j), fillSurface(CInt(iToothID), CInt(strArrSurfaceSet(j)), objMenuInstruction.Color), Tx.StatusID, Tx.ModifiedDateTime, objMenuInstruction)
            Next j
            
        Case eOperation.eHatchSurface
             For j = 0 To UBound(strArrSurfaceSet)
              TxCollection.Add ConvertDLTxToGCTx(sMenuID, mObjMenuLib.MenuItems(sMenuID).MnuButtonID, iToothID, strArrSurfaceSet(j), hatchSurface(CInt(iToothID), CInt(strArrSurfaceSet(j)), objMenuInstruction.Color, objMenuInstruction.Hatch), Tx.StatusID, Tx.ModifiedDateTime, objMenuInstruction)
            Next j
            
        Case eOperation.eSymbolizeSurface
            'Crown only surfaces reset to do all surfaces.
            If strArrSurfaceSet(0) = 111 Then
                 Set oSurfaceGroup = mObjSVGLib.SurfaceGroups(GCTooth.ToothID)
                 strArrSurfaceSet = Split(oSurfaceGroup.CrownSurfaces, ",")
            End If
            
            For j = 0 To UBound(strArrSurfaceSet)
             
                 'TODO: Why is this here???
                 Set oGCTx = New clsGCTx
                 
                 If objMenuInstruction.symbol = eSymbol.eImplant Then
                    GCTooth.IsImplant = True
                    GCTooth.DateTimeImplant = Tx.ModifiedDateTime
                 End If
                 
                 If objMenuInstruction.symbol = eParallelLines Then
                     GCTooth.DateTimeWholeToothPlanToRemove = Tx.ModifiedDateTime
                 End If
                 
                 If objMenuInstruction.symbol = eSymbol.eMissing Then
                    GCTooth.IsExtractOrMissing = True
        '                           RaiseEvent MissingTooth(iToothID)
                 End If
                 
                 TxCollection.Add ConvertDLTxToGCTx(sMenuID, mObjMenuLib.MenuItems(sMenuID).MnuButtonID, iToothID, strArrSurfaceSet(j), patternSurface(CInt(iToothID), CInt(strArrSurfaceSet(j)), objMenuInstruction.Color, objMenuInstruction.symbol), Tx.StatusID, Tx.ModifiedDateTime, objMenuInstruction)
             
             Next j
        Case eOperation.eOutlineSurface
            'Crown only surfaces reset to do all surfaces.
            If strArrSurfaceSet(0) = 111 Then
                 Set oSurfaceGroup = mObjSVGLib.SurfaceGroups(GCTooth.ToothID)
                 strArrSurfaceSet = Split(oSurfaceGroup.CrownSurfaces, ",")
            End If
            
            For j = 0 To UBound(strArrSurfaceSet)
                TxCollection.Add ConvertDLTxToGCTx(sMenuID, mObjMenuLib.MenuItems(sMenuID).MnuButtonID, iToothID, strArrSurfaceSet(j), outlineSurface(CInt(iToothID), CInt(strArrSurfaceSet(j)), objMenuInstruction.Color), Tx.StatusID, Tx.ModifiedDateTime, objMenuInstruction)
            Next j
    End Select
    
    'Set tooth removeLocation for all occasions
    GCTooth.RemoveLocation = objMenuInstruction.RemovedLocations
    
    'Set Date/Time when location was removed
    If GCTooth.RemoveLocation = eLocation.eRootSection Then
         If Tx.StatusID = txPlanned Or Tx.StatusID = txInProgress Then
            GCTooth.DateTimePlannedRootRemoved = Tx.ModifiedDateTime
         Else
            GCTooth.DateTimeRootRemoved = Tx.ModifiedDateTime
         End If
    
    ElseIf GCTooth.RemoveLocation = eLocation.eCoronalSection Then
         If Tx.StatusID = txPlanned Then
            GCTooth.DateTimePlannedCrownRemoved = Tx.ModifiedDateTime
         Else
            GCTooth.DateTimeCrownRemoved = Tx.ModifiedDateTime
         End If
         
    ElseIf GCTooth.RemoveLocation = eLocation.eWholeTooth Then
         GCTooth.DateTimeWholeToothRemoved = Tx.ModifiedDateTime
    
    End If
            
End Sub
Private Sub setTreatRetreatToToothObject(MenuID As String, GCTooth As clsGCTooth, Optional eTxStatus As CHCSII_Dental_DXTX.TxStatusEnum, Optional eDxStatus As CHCSII_Dental_DXTX.DxStatusEnum)
'>> SCR #8323; 8324; 8325; 8417; 8418; 11047; 11048; 11076;   Developer: Jon Mack 04/11/2002 09:42 AM
'SCR Description: based upon menuID we need to set the HasPrevEndo on tooth

Dim objMenuItem As clsMenuItem
Dim tmpStrMenuID As Integer


            Set objMenuItem = mObjMenuLib.MenuItems(MenuID)
            tmpStrMenuID = CInt(objMenuItem.MnuButtonID)


    If (tmpStrMenuID = 47 And (eDxStatus = dxActiveNew Or eDxStatus = dxActiveOngoing Or eDxStatus = dxActiveProvisional)) _
    Or (tmpStrMenuID = 48 And eTxStatus = txPreExisting) Then  '47 defective 48 preexisting
           GCTooth.HasPrevEndo = True
    End If
    
End Sub
Private Function getMenuInstruction(MenuID As String, iDxTxStatus As Integer) As clsMenuItemInstr

Dim objMenuItem As clsMenuItem

            Set objMenuItem = mObjMenuLib.MenuItems(MenuID)
            Set getMenuInstruction = objMenuItem.getMenuInstruction(CLng(iDxTxStatus))

             
End Function
Public Function ConvertDLDxToGCDx(sMenuID As String, iRootMenuID As Integer, iToothID As Integer, sSurface As String, sSVGDX As String, eDxStatus As CHCSII_Dental_DXTX.DxStatusEnum, dModifiedDateTime As Date, objMenuInstruction As clsMenuItemInstr, iDRC As Integer) As clsGCDx

Dim oGCDx As clsGCDx
    
    Set oGCDx = New clsGCDx
    With oGCDx
        .MenuID = CInt(sMenuID)
        .RootMenuID = iRootMenuID
        .ToothID = iToothID
        .Surface = sSurface
        .SVGDx = sSVGDX
        .DxStatus = eDxStatus
        .ModifiedDateTime = dModifiedDateTime
        .DRC = iDRC
        .ToothLocation = objMenuInstruction.Location
    End With
    
    Set ConvertDLDxToGCDx = oGCDx
    
    Set oGCDx = Nothing
    
End Function
Public Function ConvertDLTxToGCTx(sMenuID As String, iRootMenuID As Integer, iToothID As Integer, sSurface As String, sSVGTX As String, eTxStatus As CHCSII_Dental_DXTX.TxStatusEnum, dModifiedDateTime As Date, objMenuInstruction As clsMenuItemInstr) As clsGCTx

Dim oGCTx As clsGCTx
    
    Set oGCTx = New clsGCTx
    With oGCTx
        .MenuID = CInt(sMenuID)
        .RootMenuID = iRootMenuID
        .ToothID = iToothID
        .Surface = sSurface
        .SVGTx = sSVGTX
        .TxStatus = eTxStatus
        If sMenuID = 817 Or sMenuID = 819 Then  'Root Amputation and HemiSection
            .TxStatusInBothViews = True
        End If
        
        .ModifiedDateTime = dModifiedDateTime
        .ToothLocation = objMenuInstruction.Location
    End With
    
    Set ConvertDLTxToGCTx = oGCTx
    
    Set oGCTx = Nothing
    
End Function

Private Sub Class_Initialize()

  Set mObjMenuLib = mOSF.Singleton
  
  Dim mOSVGLibSF As New CHCSII_Dental_SVG_Lib.clsSVGLibSF
  
  Set mObjSVGLib = mOSVGLibSF.SVGLib
  
  'Set mobjdataelement = New clsDataElem
  Set mObjDataProc = New clsDataProc
End Sub

Private Sub Class_Terminate()

    Set mObjMenuLib = Nothing
    Set mObjSVGLib = Nothing
    'Set mobjdataelement = Nothing
    Set mObjDataProc = Nothing
    Set mOSF = Nothing
    Set mColTeeth = Nothing
    
End Sub
