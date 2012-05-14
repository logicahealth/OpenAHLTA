VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTooth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const TOTAL_SURFACES = 32

Private Const WHITE = "NoFill" ' str0"
Private Const YELLOW = "FillYellow" ' str0"
Private Const BLUE = "fil2" ' str0"
Private Const GREY = "fil3" ' str0"

Private Const WHITE_HIDE = "NoFill"
Private Const YELLOW_HIDE = "YellowYellow"
Private Const BLUE_HIDE = "BlueBlue"
Private Const GREY_HIDE = "GreyGrey"

Private m_ToothDoc As Object
Public m_ToothWin As Object
Private m_ToothStr As String
Private m_TeethObject As clsTeeth
Private m_DataProcObject As clsDataProc
Private mObjSVGMngr As CHCSII_Dental_SVG_Manager.clsManage

Private m_LastViewTooth As Integer
Private miCurrentViewTooth As Integer
Public mbMouseDown As Boolean
Private miCurrentSurface As Integer
Private strFillClass As String
Private colSelectedSurfaceIds As Collection

'Private objMacro As clsMacro
'Used in frmgraphicalcharting
Private mvarSurfaceGroups As colSurfaceGroups

Private colToothDef As Collection
Private GCToothDef As clsToothDefinition
Public objSurfaceDef As clsSurface_Definitions
Private objSurfDefAttr As clsSurface_Def_Attr

Public selectedSurfaces As String
Public mbIsAutoSelected As Boolean

'37476
'Private mBrowserViewType As eViewType
'Public mbIsDxView As Boolean

Public mbIsProcessingMenuSelection As Boolean

Private mbWhereAmI As eToothMode

'37476
Private meDxTxViewMode As eViewMode
'Private mbWhichViewAmI As eViewMode

Public Enum eToothId
  e1_2_3_14_15_16_33_34_41_42
  e4_5_12_13
  e4_13s24
  e5_12s24
  e5_12s25
  e6_7_8_9_10_11_22_23_24_25_26_27_35_36_37_38_39_40_45_46_47_48_49_50
  e17_19_30_32_43_52
  e18_31_44_51
  e20_21_28_29
End Enum

Public Enum eMouthArch
    eNone = 0
    eMax = 1
    eMan = 2
End Enum

Public Event MouseMove(surfaceName As String)
Public Event MouseOut()
Public Event SurfacesSelected(iSurfaceCount As Integer)
'Menu Oriented
Public Event ToothPosition(eToothPos As eToothPosition)
Public Event ToothType(eToothTyp As eToothType)
Public Event IsMolar(bIsMolar As Boolean)
Public Event ToothSequence(eToothSeq As eToothSequence)
Public Event IsMax(bIsMax As Boolean)
Public Event HasPrevEndo(bHasPrevEndo As Boolean)
Public Event IsMissing(bIsMissing As Boolean)
Public Event IsImplant(bIsImplant As Boolean)
Public Event IsMultiMissing(bIsMultiMissing As Boolean)
Public Event IsMultiExtract(bIsMultiExtract As Boolean)

'37476
'Can consolidate this with DxTxViewModeChanged?
'Public Event BrowserViewType(ViewType As eViewMode)
'Public Event IsDxView(bIsDxView As Boolean)

Public Event DoMenuMagicNow(bToothChanged As Boolean)
Public Event ModeIs(toothMode As eToothMode)

'37476
Public Event DxTxViewModeChanged(ViewMode As eViewMode)
'Public Event ViewModeIs(ViewMode As eViewMode)

Public Event SelectedArch(eArch As eMouthArch)
Public Event PlannedExtract(bIsPlannedExtract As Boolean)
Public Event TeethInArch(bTeethInArch As Boolean)
Public Event TeethInArchTxP(bTeethInArchTxP As Boolean)
Public Event RootSurfaces(iRootSurfaceCount As Integer)
Public Event RootSurfacesSelected(iRootSurfaceCount As Integer)
Public Event RemoveItemsByDate(eReasonForDeletion As DeleteReasonEnum)
Public Property Get SurfaceGroups() As colSurfaceGroups
    Set SurfaceGroups = mvarSurfaceGroups
End Property
Public Sub RemoveItemsByDate(eReasonForDeletion As DeleteReasonEnum)
  RaiseEvent RemoveItemsByDate(eReasonForDeletion)
End Sub
Public Property Let DxTxViewMode(vData As eViewMode)

      meDxTxViewMode = vData
      RaiseEvent DxTxViewModeChanged(meDxTxViewMode)

'Debug.Print "DxTxMode: DoMenuMagic with No surface selected"
      RaiseEvent DoMenuMagicNow(True)

End Property
Public Property Get DxTxViewMode() As eViewMode
      DxTxViewMode = meDxTxViewMode
End Property
Public Function GCToothDxTx(iToothID As Integer) As clsGCTooth
     Set GCToothDxTx = mObjSVGMngr.mColTeeth(iToothID)
End Function
Public Property Let BTFMMode(BTFMMode As eToothMode)
      mbWhereAmI = BTFMMode
      RaiseEvent ModeIs(BTFMMode)
'Debug.Print "BTFMMode: DoMenuMagic with No surface selected"
      RaiseEvent DoMenuMagicNow(True)

End Property
Public Property Get BTFMMode() As eToothMode
      BTFMMode = mbWhereAmI
End Property
Public Property Get CurrentViewTooth() As Integer
  CurrentViewTooth = miCurrentViewTooth
End Property
Public Property Let CurrentViewTooth(iToothID As Integer)
  ToothChanged iToothID
End Property
Public Sub ToothChanged(iToothID As Integer, Optional bMenuUpdate As Boolean = True, Optional bForceFull As Boolean)
   
   If iToothID = miCurrentViewTooth Then
        If Not bForceFull Then Exit Sub
   End If
   
   miCurrentViewTooth = iToothID
   If bMenuUpdate Then
      If Not mbIsProcessingMenuSelection Then
        MenuUpdateBasedOnTooth
      End If
   End If
End Sub
Public Sub MenuUpdateBasedOnTooth()
Dim i As Integer
Dim strAnterior As String
Dim strPosterior As String
Dim strMolar As String
Dim strPreMolar As String
Dim strArrPosterior() As String
Dim strArrAnterior() As String
Dim strArrMolar() As String
Dim strArrPreMolar() As String
Dim bToothPosition As Boolean
Dim bToothType As Boolean
Dim bFoundMolar As Boolean
Dim GCTooth As CHCSII_Dental_SVG_Manager.clsGCTooth
Dim oSurfaceGroup As clsSurfaceGroup
'Dim objMacro As clsMacro
Dim strArrSelected() As String


  If miCurrentViewTooth = 0 Then Exit Sub
  
  strPosterior = "1,2,3,4,5,12,13,14,15,16,17,18,19,20,21,28,29,30,31,32,33,34,41,42,43,44,51,52"
  strAnterior = "6,7,8,9,10,11,22,23,24,25,26,27,35,36,37,38,39,40,45,46,47,48,49,50"
  
  strMolar = "1,2,3,33,34,41,42,14,15,16,17,18,19,30,31,32,43,44,51,52"
  strPreMolar = "4,5,12,13,20,21,28,29"
  
  strArrPosterior = Split(strPosterior, ",")
  strArrAnterior = Split(strAnterior, ",")
  
  strArrMolar = Split(strMolar, ",")
  strArrPreMolar = Split(strPreMolar, ",")
  
  'Set GCTooth to current patients tooth
  Set GCTooth = mObjSVGMngr.mColTeeth(CStr(miCurrentViewTooth))
  
  For i = 0 To UBound(strArrPosterior)
  
     If strArrPosterior(i) = CStr(miCurrentViewTooth) Then bToothPosition = True
    
  Next i
  
  For i = 0 To UBound(strArrPreMolar)
  
     If strArrPreMolar(i) = CStr(miCurrentViewTooth) Then
        bToothType = True
        bFoundMolar = True
     End If

  Next i
  
  For i = 0 To UBound(strArrMolar)
  
     If strArrMolar(i) = CStr(miCurrentViewTooth) Then
        bToothType = False
        bFoundMolar = True
     End If

  Next i
  
  'jrm Set to toothposition (True if Posterior, Default False)
  If bToothPosition Then
     RaiseEvent ToothPosition(ePosterior)
  Else
     RaiseEvent ToothPosition(eAnterior)
  End If
  
  'jrm set toothtype (True if PreMolar, False if Molar, default False)
  If bToothType Then
    RaiseEvent ToothType(ePreMolar)
  ElseIf bFoundMolar Then
    RaiseEvent ToothType(eMolar)
  Else
    RaiseEvent ToothType(777)
  End If
 
  
  'jrm set IsMolar (True if Molar or PreMolar)
  'Deprecated 04/9/2002
  'RaiseEvent IsMolar(bFoundMolar)
  
  
  'jrm Set ToothSequence based on ToothId
  If miCurrentViewTooth < 33 Then
     RaiseEvent ToothSequence(eToothSequence.eSecondary)
  Else
     RaiseEvent ToothSequence(eToothSequence.ePrimary)
  End If
  
  'jrm set IsMax  based on ToothPosition (mandibular/maxillary)
  If miCurrentViewTooth < 17 Then
     RaiseEvent SelectedArch(eMax)
     RaiseEvent TeethInArch(AreTeethInArch(eMax))
     RaiseEvent TeethInArchTxP(AreTeethInArchTxP(eMax))
  ElseIf miCurrentViewTooth > 16 And miCurrentViewTooth < 33 Then
     RaiseEvent SelectedArch(eMan)
     RaiseEvent TeethInArch(AreTeethInArch(eMan))
     RaiseEvent TeethInArchTxP(AreTeethInArchTxP(eMan))
  ElseIf miCurrentViewTooth > 32 And miCurrentViewTooth < 43 Then
     RaiseEvent SelectedArch(eMax)
     RaiseEvent TeethInArch(AreTeethInArch(eMax))
     RaiseEvent TeethInArchTxP(AreTeethInArchTxP(eMax))
  ElseIf miCurrentViewTooth > 42 Then
     RaiseEvent SelectedArch(eMan)
     RaiseEvent TeethInArch(AreTeethInArch(eMan))
     RaiseEvent TeethInArchTxP(AreTeethInArchTxP(eMan))
  Else
     RaiseEvent SelectedArch(eNone)
     RaiseEvent TeethInArch(False)
     RaiseEvent TeethInArchTxP(False)
  End If
    
    'jrm 04/09/2002 set HasPrevEndo False
    RaiseEvent HasPrevEndo(GCTooth.HasPrevEndo)
    
    'jrm Set Missing False
    If Not GCTooth.DateTimeWholeToothRemoved = "12:00:00 AM" Then
            
            RaiseEvent IsMissing(True)
    Else
    
            RaiseEvent IsMissing(False)
    
    End If

    If Not GCTooth.DateTimeWholeToothPlanToRemove = "12:00:00 AM" Then
        RaiseEvent PlannedExtract(True)
     Else
        RaiseEvent PlannedExtract(False)
    End If
        
    'jrm Set Implant False
    If Not GCTooth.DateTimeImplant = "12:00:00 AM" Then
        RaiseEvent IsImplant(True)
    Else
        RaiseEvent IsImplant(False)
    End If
    
    RaiseEvent IsMultiMissing(False)
    
    RaiseEvent IsMultiExtract(False)
    
    '05/03/02 JRM 23103 If we are changing teeth then set surfaces to zero
     RaiseEvent SurfacesSelected(0)
     RaiseEvent RootSurfacesSelected(0)
    'jrm Set IsDxView
    '37476
    RaiseEvent DxTxViewModeChanged(meDxTxViewMode)
'    RaiseEvent BrowserViewType(meDxTxViewMode)
    'RaiseEvent IsDxView(mbIsDxView)

'Debug.Print "MenuUpdateBasedOnTooth: DoMenuMagic with No surface selected"
    RaiseEvent DoMenuMagicNow(True)



'Find # of Root Surfaces
    Set oSurfaceGroup = mvarSurfaceGroups(GCTooth.toothID)
    strArrSelected = Split(oSurfaceGroup.RootSurfaces, ",")
    RaiseEvent RootSurfaces(UBound(strArrSelected) + 1)
    

'  '<bjr092501 [SCR8309]
'  'DxDefMenuX.doMenuMagic
'  If bMultiToothMode Then
'    dxWMenuX(1).doMenuMagic   'May need to add dxWMenuX(1) later depending on view.
'  Else
'    dxWMenuX(0).doMenuMagic
'  End If
End Sub
Private Function AreTeethInArch(eArch As eMouthArch) As Boolean
Dim GCTooth As CHCSII_Dental_SVG_Manager.clsGCTooth
Dim bFound As Boolean

  'Set GCTooth to current patients tooth
  For Each GCTooth In mObjSVGMngr.mColTeeth
       
      If eArch = eMax Then
          
          If GCTooth.toothID < 17 And GCTooth.IsToothPlaceHolder = False Then  'Max
             bFound = True
          ElseIf GCTooth.toothID > 32 And GCTooth.toothID < 43 And GCTooth.IsToothPlaceHolder = False Then  'Max
             bFound = True
          End If
          
          If bFound Then
             If GCTooth.DateTimeWholeToothRemoved = "12:00:00 AM" Then
                AreTeethInArch = True
                Exit For
             Else
                AreTeethInArch = False
             End If
             
             bFound = False
          End If
          
      ElseIf eArch = eMan Then
          
          If GCTooth.toothID > 16 And GCTooth.toothID < 33 And GCTooth.IsToothPlaceHolder = False Then 'Mand
             bFound = True
          ElseIf GCTooth.toothID > 42 And GCTooth.IsToothPlaceHolder = False Then 'Mand
             bFound = True
          End If
          
          If bFound Then
             If GCTooth.DateTimeWholeToothRemoved = "12:00:00 AM" Then
                AreTeethInArch = True
                Exit For
             Else
                AreTeethInArch = False
                
             End If
             
             bFound = False
          End If
          
      End If
  
  Next
  
End Function
Private Function AreTeethInArchTxP(eArch As eMouthArch) As Boolean
Dim GCTooth As CHCSII_Dental_SVG_Manager.clsGCTooth
Dim bFound As Boolean

  'Set GCTooth to current patients tooth
  For Each GCTooth In mObjSVGMngr.mColTeeth
       
      If eArch = eMax Then
          
          If GCTooth.toothID < 17 And GCTooth.IsToothPlaceHolder = False Then 'Max
             bFound = True
          ElseIf GCTooth.toothID > 32 And GCTooth.toothID < 43 And GCTooth.IsToothPlaceHolder = False Then  'Max
             bFound = True
          End If
          
          If bFound Then
             If GCTooth.DateTimeWholeToothRemoved = "12:00:00 AM" And GCTooth.DateTimeWholeToothPlanToRemove = "12:00:00 AM" Then
                AreTeethInArchTxP = True
                Exit For
             Else
                AreTeethInArchTxP = False
             End If
             
             bFound = False
          End If
          
      ElseIf eArch = eMan Then
          
          If GCTooth.toothID > 16 And GCTooth.toothID < 33 And GCTooth.IsToothPlaceHolder = False Then 'Mand
             bFound = True
          ElseIf GCTooth.toothID > 42 And GCTooth.IsToothPlaceHolder = False Then 'Mand
             bFound = True
          End If
          
          If bFound Then
             If GCTooth.DateTimeWholeToothRemoved = "12:00:00 AM" And GCTooth.DateTimeWholeToothPlanToRemove = "12:00:00 AM" Then
                AreTeethInArchTxP = True
                Exit For
             Else
                AreTeethInArchTxP = False
             End If
             
             bFound = False
          End If
          
      End If
  
  Next
End Function

Public Property Set DataProcObject(oDataProc As clsDataProc)
    Set m_DataProcObject = oDataProc
End Property

Public Property Get DataProcObject() As clsDataProc
    Set DataProcObject = m_DataProcObject
End Property
'Public Property Set TeethObject(oTeeth As clsTeeth)
'    Set m_TeethObject = oTeeth
'End Property
'
'Public Property Get TeethObject() As clsTeeth
'    Set TeethObject = m_TeethObject
'End Property

Public Property Set ToothDoc(oToothDoc As Object)
    Set m_ToothDoc = oToothDoc
End Property

Public Property Get ToothDoc() As Object
    Set ToothDoc = m_ToothDoc
End Property

Public Property Let ToothStr(sToothStr As String)
    m_ToothStr = sToothStr
End Property

Public Property Get ToothStr() As String
    ToothStr = m_ToothStr
End Property

Public Property Let LastViewTooth(nTooth As Integer)
    m_LastViewTooth = nTooth
    mbMouseDown = False
End Property

Public Property Get LastViewTooth() As Integer
    LastViewTooth = m_LastViewTooth
End Property

Public Function init(myWin As Object)
  
    Set m_ToothDoc = myWin.myDoc
    Set m_ToothWin = myWin

    ' default in the tooth.svg is 1
    m_LastViewTooth = 1
    
End Function

Public Function SvgMouseClick(key As String, evt As Object)

End Function
Public Function SvgMouseDown(key As String, evt As Object)
Dim oSSvg As Object
Dim strarr() As String
Dim strarrSK() As String
Dim strarrGK() As String
Dim i, iTmpSurface As Integer
Dim iQPos As Integer
Dim iAPos As Integer
Dim strKey As String
Dim strArrChars(2) As String


mbMouseDown = True
strArrChars(0) = "f"
strArrChars(1) = "o"
strArrChars(2) = "l"

    strarr = Split(key, "-")
    iTmpSurface = CInt(strarr(2))
    miCurrentSurface = iTmpSurface
    
    
    
    iAPos = InStr(key, "AN")
    
    If iAPos > 0 Then
      key = Replace(key, "AN", "SS")
    End If
    
    iQPos = InStr(key, "Q")
       
   If iQPos > 0 Then
       strKey = Left(key, iQPos + 1) & "*" & Right(key, Len(key) - iQPos - 2)
       For i = 0 To UBound(strArrChars)
           Set oSSvg = m_ToothDoc.getElementById(CStr(Replace(strKey, "*", strArrChars(i))))
            If oSSvg Is Nothing Then GoTo ContinueSurfaceSel
    
            If oSSvg.getAttribute("class") = YELLOW Then
               Call oSSvg.setAttribute("class", WHITE)
               colSelectedSurfaceIds(strarr(2)).isSelected = False
            Else
               Call oSSvg.setAttribute("class", YELLOW)
               colSelectedSurfaceIds(strarr(2)).isSelected = True
            End If
            
ContinueSurfaceSel:

      Next
   Else
            Set oSSvg = m_ToothDoc.getElementById(CStr(key))
            If oSSvg Is Nothing Then GoTo ContinueSurfaceSelected
    
            If oSSvg.getAttribute("class") = YELLOW Then
               Call oSSvg.setAttribute("class", WHITE)
               colSelectedSurfaceIds(strarr(2)).isSelected = False
            Else
               Call oSSvg.setAttribute("class", YELLOW)
               colSelectedSurfaceIds(strarr(2)).isSelected = True
            End If
ContinueSurfaceSelected:
   End If
End Function
Public Function SvgMouseOver(key As String, evt As Object)
Dim oSSvg As Object
Dim strarr() As String
Dim strarrSK() As String
Dim i, iTmpSurface As Integer
Dim iQPos As Integer
Dim iAPos As Integer
Dim strKey As String
Dim strArrChars(2) As String

strarr = Split(key, "-")

If Not IsNumeric(strarr(2)) Then Exit Function

iTmpSurface = CInt(strarr(2))
RaiseEvent MouseMove(getSurfaceName(CInt(strarr(2))))

If Not mbMouseDown Then Exit Function

 
    If Not miCurrentSurface = iTmpSurface Then
       miCurrentSurface = iTmpSurface
    Else
       Exit Function
    End If
    
strArrChars(0) = "f"
strArrChars(1) = "o"
strArrChars(2) = "l"

     iAPos = InStr(key, "AN")
    
    If iAPos > 0 Then
      key = Replace(key, "AN", "SS")
    End If
    
    iQPos = InStr(key, "Q")
       
   If iQPos > 0 Then
       strKey = Left(key, iQPos + 1) & "*" & Right(key, Len(key) - iQPos - 2)
       For i = 0 To UBound(strArrChars)
           Set oSSvg = m_ToothDoc.getElementById(CStr(Replace(strKey, "*", strArrChars(i))))
            If oSSvg Is Nothing Then GoTo ContinueSurfaceSel
    
            If oSSvg.getAttribute("class") = YELLOW Then
               Call oSSvg.setAttribute("class", WHITE)
               colSelectedSurfaceIds(strarr(2)).isSelected = False
            Else
               Call oSSvg.setAttribute("class", YELLOW)
               colSelectedSurfaceIds(strarr(2)).isSelected = True
            End If
            
ContinueSurfaceSel:

      Next
   Else
            Set oSSvg = m_ToothDoc.getElementById(CStr(key))
            If oSSvg Is Nothing Then GoTo ContinueSurfaceSel
    
            If oSSvg.getAttribute("class") = YELLOW Then
               Call oSSvg.setAttribute("class", WHITE)
               colSelectedSurfaceIds(strarr(2)).isSelected = False
            Else
               Call oSSvg.setAttribute("class", YELLOW)
               colSelectedSurfaceIds(strarr(2)).isSelected = True
            End If
   End If
End Function
Public Sub UpdateTooth(iToothID As Integer, Optional DxTxViewMode As eViewMode = eDxView)

Dim GCTooth As CHCSII_Dental_SVG_Manager.clsGCTooth

    '37476
    If DxTxViewMode = eDxView Then
         m_ToothStr = mObjSVGMngr.getsvgDxViewbyTooth(iToothID)
    ElseIf DxTxViewMode = eTxView Then
         m_ToothStr = mObjSVGMngr.getsvgTxViewbyTooth(iToothID)
    ElseIf DxTxViewMode = eDx603aView Then
        m_ToothStr = mObjSVGMngr.getsvgDx603aViewbyTooth(iToothID)
    End If

'jrm 04/09/2002  must update menu after menu is selected only for now.
'Once selection is done this can be deleted because after DXTX then no tooth should
'be selected
'Set GCTooth to current patients tooth
    Set GCTooth = mObjSVGMngr.mColTeeth(CStr(iToothID))
          'jrm 04/09/2002 set HasPrevEndo False
    RaiseEvent HasPrevEndo(GCTooth.HasPrevEndo)
    
    resetSelectedSurfaces

End Sub

 
Public Function SvgMouseOut(key As String, evt As Object)
Dim strarr() As String

strarr = Split(key, "!")
RaiseEvent MouseOut

End Function
Public Function SvgMouseUp(key As String, evt As Object)
Dim objselectedsurface As clsSelectedSurface
Dim iSurfaceCounter As Integer
Dim iRootSurfaceCounter As Integer

mbMouseDown = False
selectedSurfaces = ""
For Each objselectedsurface In colSelectedSurfaceIds
   If objselectedsurface.isSelected Then
      If selectedSurfaces = "" Then
        selectedSurfaces = CStr(objselectedsurface.surfaceID)
      Else
        selectedSurfaces = selectedSurfaces & "|" & CStr(objselectedsurface.surfaceID)
      End If
   End If
Next

'Debug.Print selectedSurfaces
'Debug.Print "SvgMouseUp: raiseevent SurfaceSelected"

If selectedSurfaces = "" Or mbWhereAmI = eFullMouth Then
    RaiseEvent SurfacesSelected(0)
    RaiseEvent RootSurfacesSelected(0)
Else
    selectedSurfacesCount iSurfaceCounter, iRootSurfaceCounter
    RaiseEvent SurfacesSelected(iSurfaceCounter)
    RaiseEvent RootSurfacesSelected(iRootSurfaceCounter)
End If
'SF 3-20-02 - Update the menu here instead of in the surface change property to increase the performance in full mouth
'Debug.Print "SvgMouseUp: DoMenuMagicNow"
RaiseEvent DoMenuMagicNow(False)

End Function
Private Sub selectedSurfacesCount(iSurfaceCounter As Integer, iRootCounter As Integer)
Dim colSurfaces As Collection
Dim objselectedsurface As clsSelectedSurface
Dim oSurfaceGroup As clsSurfaceGroup
'Dim objMacro As clsMacro
Dim strArrSelected() As String
Dim i As Integer
Dim bFoundRoot As Boolean
Dim bErr As Boolean

On Error GoTo errHandler

iSurfaceCounter = 0
iRootCounter = 0

Set colSurfaces = New Collection

Set oSurfaceGroup = mvarSurfaceGroups(miCurrentViewTooth)
strArrSelected = Split(oSurfaceGroup.RootSurfaces, ",")

For Each objselectedsurface In colSelectedSurfaceIds
   If objselectedsurface.isSelected Then
     For i = 0 To UBound(strArrSelected)
        If objselectedsurface.surfaceID = CInt(strArrSelected(i)) Then
          bFoundRoot = True
          Exit For
        End If
        
     Next i
     
     If Not bFoundRoot Then
        
       colSurfaces.Add getSurfaceAbbreviation(objselectedsurface.surfaceID), getSurfaceAbbreviation(objselectedsurface.surfaceID)

        
     Else
     
        iRootCounter = iRootCounter + 1
        
     End If
     bFoundRoot = False
   End If
Next

iSurfaceCounter = colSurfaces.Count
Set colSurfaces = Nothing
Exit Sub

errHandler:
   
   If Err.Number = 457 Then
        bErr = True
        Err.Clear
        Resume Next
   Else
        MsgBxARMd Err.Number & "," & Err.Description & ",clsManage.ViewableSurfaceDxView, SVGManager"

   End If
End Sub
Public Sub resetSelectedSurfaces()

Dim i As Integer
Dim objselectedsurface As clsSelectedSurface

    For i = 1 To TOTAL_SURFACES  'Max Number of Surfaces
          Set objselectedsurface = colSelectedSurfaceIds(i)
          objselectedsurface.isSelected = False
    Next i
    selectedSurfaces = ""

'Debug.Print "resetSelectedSurfaces: raiseevent SurfaceSelected"
'    If Not mbIsProcessingMenuSelection Then
        RaiseEvent SurfacesSelected(0)
        RaiseEvent RootSurfacesSelected(0)
    'SF 3-20-02 - Update the menu here instead of in the surface change property to increase the performance in full mouth
'Debug.Print "resetSelectedSurfaces: DoMenuMagicNow"
        RaiseEvent DoMenuMagicNow(False)
'    End If
End Sub
Public Sub macroCrown()
Dim oSurfaceGroup As clsSurfaceGroup
'Dim objMacro As clsMacro

   Set oSurfaceGroup = mvarSurfaceGroups(miCurrentViewTooth)
   selectedSurfaces = Replace(oSurfaceGroup.CrownSurfaces, ",", "|")
   selectSurfaces selectedSurfaces
   selectedSurfaces = "111"
   
''Debug.Print "macroCrown: raiseevent SurfaceSelected"
'    If selectedSurfaces = "" Then
'         RaiseEvent SurfacesSelected(0)
'         RaiseEvent RootSurfacesSelected(0)
'    Else
'         RaiseEvent SurfacesSelected(5)
'    End If
'
'    'SF 3-20-02 - Update the menu here instead of in the surface change property to increase the performance in full mouth
''Debug.Print "macroCrown: DoMenuMagicNow"
'    RaiseEvent DoMenuMagicNow(False)

End Sub
Public Sub macroRoots()
Dim oSurfaceGroup As clsSurfaceGroup

    Set oSurfaceGroup = mvarSurfaceGroups(miCurrentViewTooth)
    selectedSurfaces = Replace(oSurfaceGroup.RootSurfaces, ",", "|")
    selectSurfaces selectedSurfaces
    
''Debug.Print "macroRoots: raiseevent SurfaceSelected"
'    If selectedSurfaces = "" Then
'         RaiseEvent SurfacesSelected(0)
'         RaiseEvent RootSurfacesSelected(0)
'    Else
'         RaiseEvent SurfacesSelected(Len(getSurfaceAbbrevs))
'    End If
'
'    'SF 3-20-02 - Update the menu here instead of in the surface change property to increase the performance in full mouth
''Debug.Print "macroRoots: DoMenuMagicNow"
'    RaiseEvent DoMenuMagicNow(False)

End Sub
Public Sub selectSurfaces(sSelected As String)
Dim i As Integer
Dim strArrSelected() As String
Dim strSurface As String

strArrSelected = Split(sSelected, "|")

For i = 0 To UBound(strArrSelected)
   
   strSurface = "SS-S-" & strArrSelected(i) & "-Q-f-T-" & miCurrentViewTooth
   SvgMouseDown strSurface, Nothing
   strSurface = "SS-S-" & strArrSelected(i) & "-Q-o-T-" & miCurrentViewTooth
   SvgMouseDown strSurface, Nothing
   strSurface = "SS-S-" & strArrSelected(i) & "-Q-l-T-" & miCurrentViewTooth
   SvgMouseDown strSurface, Nothing
   strSurface = "SS-S-" & strArrSelected(i) & "-T-" & miCurrentViewTooth
   SvgMouseDown strSurface, Nothing
   
Next i

    SvgMouseUp "", Nothing
    

End Sub

Public Function createTooth(nToothID As Integer)
       
End Function
Public Sub initializeToothDefinitions()
Dim i As Integer
Dim sizeOfString As String
Dim strCompressed As String
Dim strUniSurface As String
Dim strAnsiSurface As String
Dim oSurfaceGroup As clsSurfaceGroup
Dim rows() As String
Dim cols() As String
Dim j As Integer



Set colToothDef = New Collection
For i = 1 To 32
  Set GCToothDef = New clsToothDefinition
  Select Case i
     Case 4, 5, 6, 7, 8, 9, 10, 11, 12, 13
        GCToothDef.toothID = i
        GCToothDef.toothPrimaryName = LoadResString(i + 10029)
        GCToothDef.toothSecondaryName = LoadResString(i + 10000)
        colToothDef.Add GCToothDef, CStr(i)
     Case 20, 21, 22, 23, 24, 25, 26, 27, 28, 29
        GCToothDef.toothID = i
        GCToothDef.toothPrimaryName = LoadResString(i + 10023)
        GCToothDef.toothSecondaryName = LoadResString(i + 10000)
        colToothDef.Add GCToothDef, CStr(i)
     Case Else
        GCToothDef.toothID = i
        GCToothDef.toothPrimaryName = LoadResString(i + 10000)
        GCToothDef.toothSecondaryName = LoadResString(i + 10000)
        colToothDef.Add GCToothDef, CStr(i)
  End Select
  Set GCToothDef = Nothing
Next i
   
    Set objSurfaceDef = New clsSurface_Definitions
    
    Set mvarSurfaceGroups = mObjSVGMngr.mObjSVGLib.SurfaceGroups
   
End Sub
Public Function getToothName() As String

     If miCurrentViewTooth = 0 Then Exit Function
     
     If miCurrentViewTooth < 33 Then
        getToothName = colToothDef(CStr(miCurrentViewTooth)).toothSecondaryName
      ElseIf miCurrentViewTooth < 43 Then
         getToothName = colToothDef(CStr(miCurrentViewTooth - 29)).toothPrimaryName
      Else
         getToothName = colToothDef(CStr(miCurrentViewTooth - 23)).toothPrimaryName
      End If
   
End Function
Private Function getSurfaceName(surfaceID As Integer) As String
  Select Case miCurrentViewTooth
    Case 1, 2, 3, 14, 15, 16, 33, 34, 41, 42
      Set objSurfDefAttr = objSurfaceDef.surfDefAttr(e1_2_3_14_15_16_33_34_41_42, surfaceID)
      getSurfaceName = objSurfDefAttr.surfaceName
    Case 17, 19, 30, 32, 43, 52
      Set objSurfDefAttr = objSurfaceDef.surfDefAttr(e17_19_30_32_43_52, surfaceID)
      getSurfaceName = objSurfDefAttr.surfaceName
    Case 18, 31, 44, 51
      Set objSurfDefAttr = objSurfaceDef.surfDefAttr(e18_31_44_51, surfaceID)
      getSurfaceName = objSurfDefAttr.surfaceName
    Case 4, 5, 12, 13
      If (miCurrentViewTooth = 4 Or miCurrentViewTooth = 13) And surfaceID = 24 Then
        Set objSurfDefAttr = objSurfaceDef.surfDefAttr(e4_13s24, surfaceID)
        getSurfaceName = objSurfDefAttr.surfaceName
      ElseIf (miCurrentViewTooth = 5 Or miCurrentViewTooth = 12) And surfaceID = 24 Then
        Set objSurfDefAttr = objSurfaceDef.surfDefAttr(e5_12s24, surfaceID)
        getSurfaceName = objSurfDefAttr.surfaceName
      ElseIf (miCurrentViewTooth = 5 Or miCurrentViewTooth = 12) And surfaceID = 25 Then
        Set objSurfDefAttr = objSurfaceDef.surfDefAttr(e5_12s25, surfaceID)
        getSurfaceName = objSurfDefAttr.surfaceName
      Else
        Set objSurfDefAttr = objSurfaceDef.surfDefAttr(e4_5_12_13, surfaceID)
        getSurfaceName = objSurfDefAttr.surfaceName
     End If
    Case 6, 7, 8, 9, 10, 11, 22, 23, 24, 25, 26, 27, 35, 36, 37, 38, 39, 40, 45, 46, 47, 48, 49, 50
      Set objSurfDefAttr = objSurfaceDef.surfDefAttr(e6_7_8_9_10_11_22_23_24_25_26_27_35_36_37_38_39_40_45_46_47_48_49_50, surfaceID)
      getSurfaceName = objSurfDefAttr.surfaceName
    Case 20, 21, 28, 29
      Set objSurfDefAttr = objSurfaceDef.surfDefAttr(e20_21_28_29, surfaceID)
      getSurfaceName = objSurfDefAttr.surfaceName
    Case Else
      Set objSurfDefAttr = objSurfaceDef.surfDefAttr(e4_5_12_13, surfaceID)
      getSurfaceName = objSurfDefAttr.surfaceName
    
  End Select
End Function

Private Function getSurfaceAbbreviation(surfaceID As Integer) As String

  Select Case miCurrentViewTooth
    Case 1, 2, 3, 14, 15, 16, 33, 34, 41, 42
      Set objSurfDefAttr = objSurfaceDef.surfDefAttr(e1_2_3_14_15_16_33_34_41_42, surfaceID)
      getSurfaceAbbreviation = objSurfDefAttr.surfaceAbbr
    Case 17, 19, 30, 32, 43, 52
      Set objSurfDefAttr = objSurfaceDef.surfDefAttr(e17_19_30_32_43_52, surfaceID)
      getSurfaceAbbreviation = objSurfDefAttr.surfaceAbbr
    Case 18, 31, 44, 51
      Set objSurfDefAttr = objSurfaceDef.surfDefAttr(e18_31_44_51, surfaceID)
      getSurfaceAbbreviation = objSurfDefAttr.surfaceAbbr
    Case 4, 5, 12, 13
      If (miCurrentViewTooth = 4 Or miCurrentViewTooth = 13) And surfaceID = 24 Then
        Set objSurfDefAttr = objSurfaceDef.surfDefAttr(e4_13s24, surfaceID)
        getSurfaceAbbreviation = objSurfDefAttr.surfaceAbbr
      ElseIf (miCurrentViewTooth = 5 Or miCurrentViewTooth = 12) And surfaceID = 24 Then
        Set objSurfDefAttr = objSurfaceDef.surfDefAttr(e5_12s24, surfaceID)
        getSurfaceAbbreviation = objSurfDefAttr.surfaceAbbr
      ElseIf (miCurrentViewTooth = 5 Or miCurrentViewTooth = 12) And surfaceID = 25 Then
        Set objSurfDefAttr = objSurfaceDef.surfDefAttr(e5_12s25, surfaceID)
        getSurfaceAbbreviation = objSurfDefAttr.surfaceAbbr
      Else
        Set objSurfDefAttr = objSurfaceDef.surfDefAttr(e4_5_12_13, surfaceID)

        getSurfaceAbbreviation = objSurfDefAttr.surfaceAbbr
     End If
    Case 6, 7, 8, 9, 10, 11, 22, 23, 24, 25, 26, 27, 35, 36, 37, 38, 39, 40, 45, 46, 47, 48, 49, 50
      Set objSurfDefAttr = objSurfaceDef.surfDefAttr(e6_7_8_9_10_11_22_23_24_25_26_27_35_36_37_38_39_40_45_46_47_48_49_50, surfaceID)
      getSurfaceAbbreviation = objSurfDefAttr.surfaceAbbr
    Case 20, 21, 28, 29
      Set objSurfDefAttr = objSurfaceDef.surfDefAttr(e20_21_28_29, surfaceID)
      getSurfaceAbbreviation = objSurfDefAttr.surfaceAbbr

    Case Else
      Set objSurfDefAttr = objSurfaceDef.surfDefAttr(e4_5_12_13, surfaceID)
      getSurfaceAbbreviation = objSurfDefAttr.surfaceAbbr
    
  End Select
End Function

Private Function ConvertAbbrevs(colAbbreviations As Collection) As String

Dim strTeethAbbrev As String
Dim strAnteriorTeethAbbrevIndex As String

Dim arrTeethAbbrev() As String
Dim arrTeethAbbrevIndex() As Integer

Dim strAbbrev As String
Dim i As Integer
Dim j As Integer
Dim bFound As Boolean
Dim selChar As String
Dim searchCharacter As String
Dim iAscCodePointer As Integer
Dim iCount As Integer

Dim MacroSurfaces() As String

Dim colPossibleAbbrev As Collection
Dim colFilteredAbbrev As Collection

On Error GoTo errHandler

'default return value
ConvertAbbrevs = ""

   bFound = False
   iAscCodePointer = 0
   strAbbrev = ""
   Set colPossibleAbbrev = New Collection
   Set colFilteredAbbrev = New Collection
   
   For i = 1 To colAbbreviations.Count
      If InStr(1, UCase(colAbbreviations.Item(i)), "R") = 0 And InStr(1, UCase(colAbbreviations.Item(i)), "ROOT") = 0 Then
         strAbbrev = strAbbrev & colAbbreviations.Item(i)
      End If
   Next i
   
   i = InStr(1, UCase(strAbbrev), UCase("Outline"))
   If i > 0 Then
      strAbbrev = Mid(strAbbrev, 1, i - 1) & Mid(strAbbrev, i + Len(UCase("Outline")))
   End If
   
   'create new collection of single characters (SCR 12414)
   For i = 1 To Len(strAbbrev)
      selChar = Mid(strAbbrev, i, 1)
      colPossibleAbbrev.Add selChar
   Next i
   
   'remove duplicated characters
   For i = 1 To colPossibleAbbrev.Count
      selChar = colPossibleAbbrev.Item(i)
      If InStr(1, UCase(colPossibleAbbrev.Item(i)), "R") = 0 And InStr(1, UCase(colPossibleAbbrev.Item(i)), "ROOT") = 0 Then
         For j = 1 To colFilteredAbbrev.Count
            If selChar = colFilteredAbbrev.Item(j) Then
               bFound = True
               Exit For
            Else
               bFound = False
            End If
         Next j
         If Not bFound Then
            colFilteredAbbrev.Add selChar
            bFound = False
         End If
      End If
   Next i
     
   strAbbrev = ""
   For i = 1 To colFilteredAbbrev.Count
   'get string from collection of items
      strAbbrev = strAbbrev & colFilteredAbbrev.Item(i)
     'create index for each abbreviation
     'based on sum of combination of ASCII code
      For j = 1 To Len(colFilteredAbbrev.Item(i))
         selChar = Mid(colFilteredAbbrev.Item(i), j, 1)
         iAscCodePointer = iAscCodePointer + Asc(UCase(selChar))
     Next j
     
   Next i
   
   'returned string as default value
   'strAbbrev = ""
   bFound = False
   ConvertAbbrevs = strAbbrev
   
   'currently Root abbrev is not being processed
   'If InStr(1, UCase(strAbbrev), "ROOT") > 0 Then GoTo CleanUp
         
   'to specify Anterior or Posterior based on "I" character presence in abbrev.
   If InStr(1, UCase(strAbbrev), "I") > 0 Then
      'possible combinations of surface abbreviations for anterior teeth
      '("...\CM Documents - Do not modify\BK_Tooth mapping version 3_1.doc")
      strTeethAbbrev = "DF,DI,IF,IL,MF,MI,DFL,DIF,DIL,MFD,MFL,MLD,MID,MIF,MIL,DIFL,MFDL,MIDF,MIDL,MIFL,MIDFL"
      arrTeethAbbrev = Split(strTeethAbbrev, ",")
   Else
      'possible combinations of surface abbreviations for posterior teeth
      '("...\CM Documents - Do not modify\BK_Tooth mapping version 3_1.doc")
      strTeethAbbrev = "DF,DL,DO,MF,ML,MO,OF,OL,DFL,DOF,DOL,FOL,MFD,MFL,MLD,MOD,MOF,MOL,DOFL,MFDL,MODF,MODL,MOFL,MODFL"
      arrTeethAbbrev = Split(strTeethAbbrev, ",")
   End If
     
   ReDim arrTeethAbbrevIndex(UBound(arrTeethAbbrev))
   For i = LBound(arrTeethAbbrev) To UBound(arrTeethAbbrev)
      arrTeethAbbrevIndex(i) = 0
   Next i
            
   'initial index settings
   For i = LBound(arrTeethAbbrev) To UBound(arrTeethAbbrev)
      'parse abbrev to define index
      For j = 1 To Len(arrTeethAbbrev(i))
         selChar = Mid(arrTeethAbbrev(i), j, 1)
         arrTeethAbbrevIndex(i) = arrTeethAbbrevIndex(i) + Asc(UCase(selChar))
      Next j
   Next i
            
   strAbbrev = ""
   iCount = -1
   
   'to determine right index
   For i = LBound(arrTeethAbbrev) To UBound(arrTeethAbbrev)
      If arrTeethAbbrevIndex(i) = iAscCodePointer Then   'index found
         searchCharacter = Mid(arrTeethAbbrev(i), 1, 1)
          For j = 1 To colFilteredAbbrev.Count
            'since the index value is not unique
            'the occurrence of character should be verified
            If searchCharacter = colFilteredAbbrev.Item(j) Then
               iCount = i
               bFound = True
               Exit For
            End If
         Next j
      End If
      If bFound Then Exit For
   Next
   
If bFound Then
    'add roots from original collection
   For i = 1 To colAbbreviations.Count
      strAbbrev = colAbbreviations.Item(i)
      If InStr(1, UCase(strAbbrev), "ROOT") > 0 Or InStr(1, UCase(strAbbrev), "R") > 0 Then
         arrTeethAbbrev(iCount) = arrTeethAbbrev(iCount) & "," & strAbbrev
      End If
   Next i
   ConvertAbbrevs = arrTeethAbbrev(iCount)
Else
   '==========================================================================
   'add delimiter to string with single characters
   Dim sStr As String
   'the separation of single characaters is currently commented out
   ''sStr = ""
   ''For I = 1 To Len(ConvertAbbrevs)
   ''   sStr = sStr & Mid(ConvertAbbrevs, I, 1) & ","
   ''Next I
   ''sStr = ConvertAbbrevs
   ''If Len(ConvertAbbrevs) <> 0 Then ConvertAbbrevs = Mid(sStr, 1, Len(sStr) - 1) '-1 because of comma at the end of line
   '============================================================================
   'add roots from original collection
      For i = 1 To colAbbreviations.Count
         strAbbrev = colAbbreviations.Item(i)
         If InStr(1, UCase(strAbbrev), "ROOT") > 0 Or InStr(1, UCase(strAbbrev), "R") > 0 Then
            If Len(ConvertAbbrevs) <> 0 Then
               ConvertAbbrevs = ConvertAbbrevs & "," & strAbbrev
            Else
               ConvertAbbrevs = ConvertAbbrevs & strAbbrev & ","
            End If
         End If
      Next i
      sStr = ""
      If Len(ConvertAbbrevs) = 0 Then GoTo CleanUp
      If InStrRev(ConvertAbbrevs, ",", -1) = Len(ConvertAbbrevs) Then
         ConvertAbbrevs = Mid(ConvertAbbrevs, 1, Len(ConvertAbbrevs) - 1)
      End If
      Dim iInstring As Integer
      iInstring = InStr(1, ConvertAbbrevs, ",,")
      If iInstring > 0 Then
         ConvertAbbrevs = Mid(ConvertAbbrevs, 1, iInstring - 1) & Mid(ConvertAbbrevs, iInstring + 1)
      End If
End If
GoTo CleanUp
               
errHandler:
   '-- -show the VB error---
   Call gobjShared.ShowVBError(Err.Number, Err.Description, "frmGraphicalCharting.ConvertAbbrevs", "CHCSII_Dental_BigTooth", vbCritical)

CleanUp:
   If Not colPossibleAbbrev Is Nothing Then Set colPossibleAbbrev = Nothing
   If Not colFilteredAbbrev Is Nothing Then Set colFilteredAbbrev = Nothing
   Exit Function
   
End Function
Public Function getSurfaceAbbrevs() As String

Dim i, j As Integer
Dim colAbbrev As Collection
Dim bFound As Boolean
Dim strAbbrev As String
Dim strArrAbbrev() As String
Dim strTmpAbbrev As String
Dim strMacroCrownAbbrev As String
Dim strMacroRootAbbrev As String
Dim oSurfaceGroup As clsSurfaceGroup

       Set oSurfaceGroup = mvarSurfaceGroups(miCurrentViewTooth)
       strMacroCrownAbbrev = Replace(oSurfaceGroup.CrownSurfaces, ",", "|")
       strMacroRootAbbrev = Replace(oSurfaceGroup.RootSurfaces, ",", "|")

Set colAbbrev = New Collection

      strAbbrev = selectedSurfaces
      If mbIsAutoSelected = True Then
            If selectedSurfaces = "777" Then
               getSurfaceAbbrevs = " "
               Exit Function
            ElseIf strMacroCrownAbbrev = strAbbrev Then
               '>> SCR #22690;   Developer: Jon Mack 04/03/2002 07:38 PM
               'SCR Description: Len Problem from macro selection
               'JRM len of this produces 5 surfaces
               getSurfaceAbbrevs = "     "
               Exit Function
            ElseIf strMacroRootAbbrev = strAbbrev Then
               '>> SCR #22690;   Developer: Jon Mack 04/03/2002 07:39 PM
               'SCR Description: Len Problem from macro selection
               'JRM len of this produces 3 surfaces
               getSurfaceAbbrevs = "   "
               Exit Function
            ElseIf Len(selectedSurfaces) < 1 Then
               getSurfaceAbbrevs = " "
               Exit Function
            End If
      Else
            If strAbbrev = "111" Then
                strAbbrev = strMacroCrownAbbrev
            End If
      End If
      strArrAbbrev = Split(strAbbrev, "|")
    For i = 0 To UBound(strArrAbbrev)
       strTmpAbbrev = getSurfaceAbbreviation(CInt(strArrAbbrev(i)))
       For j = 1 To colAbbrev.Count
         
         If strTmpAbbrev = colAbbrev.Item(j) Then
            bFound = True
            Exit For
         Else
            bFound = False
         End If
       Next j
       If Not bFound Then
          colAbbrev.Add strTmpAbbrev
          bFound = False
       End If
    
    Next i
       strAbbrev = ""
     
       
         getSurfaceAbbrevs = ConvertAbbrevs(colAbbrev)
        
Set colAbbrev = Nothing

End Function
Private Sub Class_Initialize()
Dim i As Integer
Dim objselectedsurface As clsSelectedSurface

    Dim oSingleton As New CHCSII_Dental_SVG_Manager.SF
    Set mObjSVGMngr = oSingleton.Singleton  'Set SVG Manager to one instance
    Set colSelectedSurfaceIds = New Collection
    For i = 1 To TOTAL_SURFACES  'Max Number of Surfaces
        Set objselectedsurface = New clsSelectedSurface
        objselectedsurface.surfaceID = i
        objselectedsurface.isSelected = False
        colSelectedSurfaceIds.Add objselectedsurface, CStr(i)
    Next i
    initializeToothDefinitions
End Sub

Private Sub Class_Terminate()

    Set m_ToothDoc = Nothing
    Set m_ToothWin = Nothing
    Set m_TeethObject = Nothing
    Set m_DataProcObject = Nothing
    Set mObjSVGMngr = Nothing
    Set colSelectedSurfaceIds = Nothing
    Set mvarSurfaceGroups = Nothing
    Set colToothDef = Nothing
    Set objSurfaceDef = Nothing
    Set objSurfDefAttr = Nothing
    
    
End Sub


