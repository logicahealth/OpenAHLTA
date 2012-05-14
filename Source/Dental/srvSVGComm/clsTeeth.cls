VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTeeth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private moPositions As colPositions

'These all contain an array of teeth positions
'Private aTeethStatus(TOTAL_TEETH) As Integer
'Private aTeethSeen(TOTAL_TEETH) As Integer
'Private aTeethStatusPre(TOTAL_TEETH) As Integer
'Private aTeethSeenPre(TOTAL_TEETH) As Integer
'Private aTeethSelected(TOTAL_TEETH) As Integer

Private Const DB_CLICK = 2
        
Private Const WHITE = "WhiteOnly"
Private Const YELLOW = "YellowOnly"
Private Const YELLOW2 = "FillYellow"
Private Const BLUE = "BlueOnly"
Private Const GREY = "GreyOnly"

Private Const WHITE_HIDE = "NoFill"
Private Const YELLOW_HIDE = "YellowYellow"
Private Const BLUE_HIDE = "BlueBlue"
Private Const GREY_HIDE = "GreyGrey"

Private m_gTeethDoc As Object
Private m_gTeethWin As Object

Private m_TeethDocDx As Object
Private m_TeethWinDx As Object

Private m_TeethDocTx As Object
Private m_TeethWinTx As Object

'37476
Private m_TeethDocDx603a As Object
Private m_TeethWinDx603a As Object

Private m_ToothObject As clsTooth

Private m_ToothMode As eToothMode
Private WithEvents mObjSVGMngr As CHCSII_Dental_SVG_Manager.clsManage
Attribute mObjSVGMngr.VB_VarHelpID = -1
Private WithEvents mObjSVGDraw As CHCSII_Dental_SVG_Manager.clsSVGDraw
Attribute mObjSVGDraw.VB_VarHelpID = -1

'This has been replaced by moPositions
'Private m_sAgePath As String

Private m_objShared As CWShared

Private mbNewTeethPath As Boolean

Private miCurrentTooth As Integer
Private bMolar As Boolean
Private bDblClick As Boolean
Private bShift As Boolean
Private bCtrl As Boolean

Private meDxTxViewMode As eViewMode
'Private mbIsDxView As Boolean

Private nNumOfSelected As Integer

Private mColSelectedTeeth As colSelectedTeeth

Public Event ToothChanged(iToothIDNew As Integer, iToothIDOld As Integer, RefreshGrid As Boolean)
Public Event SelectionCountChange(bSingleSelected As Boolean)
Public Event dblClickMultiTooth(RHV As Integer)
Public Event singleClickMultiTooth(RHV As Integer)
Public Event priSecondClick(currentToothID As Integer)
Public Event DrawingFinished()

Public Property Get CurrentTooth() As Integer

    CurrentTooth = miCurrentTooth
    
End Property

Public Property Let CurrentTooth(vData As Integer)

    miCurrentTooth = vData

End Property
Public Property Get Positions() As colPositions

   Set Positions = moPositions
    
End Property


Public Property Let toothMode(bTooth As eToothMode)
    
    If bTooth = -1 Then
        Debug.Assert False
    End If
    
    
    m_ToothMode = bTooth
    
    If m_ToothMode = eBigTooth Then
       If Not m_ToothObject Is Nothing Then
            m_ToothObject.BTFMMode = eBigTooth
       End If
    ElseIf m_ToothMode = eFullMouth Then
        If Not m_ToothObject Is Nothing Then
            m_ToothObject.BTFMMode = eFullMouth
        End If
    End If
    
End Property

Public Property Get toothMode() As eToothMode

    If m_ToothMode = -1 Then
        Debug.Assert False
    End If
    
    toothMode = m_ToothMode
    
End Property

Public Property Set ToothObject(oTooth As clsTooth)
    Set m_ToothObject = oTooth
End Property

Public Property Get ToothObject() As clsTooth
    Set ToothObject = m_ToothObject
End Property
Public Property Get MaxillaryTeeth() As String

'Get a string of maxillary teeth numbers based on the positions

Dim iCounter As Integer

    For iCounter = 1 To 16
        If moPositions(iCounter).ToothErupted Then
            MaxillaryTeeth = MaxillaryTeeth & "," & moPositions(iCounter).ToothNum
        End If
    Next
    
    'Remove the first ","
    MaxillaryTeeth = Mid(MaxillaryTeeth, 2)

'Dim i As Integer
'Dim strArrTeeth() As String
'
'strArrTeeth = Split(m_sAgePath, ",")
'
'For i = 0 To UBound(strArrTeeth)
'               If strArrTeeth(i) < 17 Then
'                   MaxillaryTeeth = MaxillaryTeeth & "," & strArrTeeth(i)
'                ElseIf strArrTeeth(i) > 32 And strArrTeeth(i) < 43 Then
'                   MaxillaryTeeth = MaxillaryTeeth & "," & strArrTeeth(i)
'                End If
'Next i
'
'MaxillaryTeeth = Right(MaxillaryTeeth, Len(MaxillaryTeeth) - 1)

End Property
Public Property Get AllTeeth() As String

Debug.Assert False
'Use Positions off of this object instead

'AllTeeth = m_sAgePath

End Property
Public Property Get MandibularTeeth() As String

Dim iCounter As Integer

    For iCounter = 17 To 32
        If moPositions(iCounter).ToothErupted Then
            MandibularTeeth = MandibularTeeth & "," & moPositions(iCounter).ToothNum
        End If
    Next
    
    'Remove the first ","
    MandibularTeeth = Mid(MandibularTeeth, 2)
    
'Dim i As Integer
'Dim strArrTeeth() As String
'
'strArrTeeth = Split(m_sAgePath, ",")
'
'For i = 0 To UBound(strArrTeeth)
'                If strArrTeeth(i) > 16 And strArrTeeth(i) < 33 Then
'                   MandibularTeeth = MandibularTeeth & "," & strArrTeeth(i)
'                ElseIf strArrTeeth(i) > 42 Then
'                   MandibularTeeth = MandibularTeeth & "," & strArrTeeth(i)
'                End If
'Next i
'
' MandibularTeeth = Right(MandibularTeeth, Len(MandibularTeeth) - 1)
            
End Property
Public Property Get AgePath() As String
    
'Use Positions instead of this
Debug.Assert False

    'AgePath = m_sAgePath
End Property

Public Property Set TeethDocDx(oTeethDocDx As Object)
    Set m_TeethDocDx = oTeethDocDx
End Property

Public Property Set TeethDocDx603a(oTeethDocDx603a As Object)
    Set m_TeethDocDx603a = oTeethDocDx603a
End Property

Public Property Get TeethDocDx() As Object
    Set TeethDocDx = m_TeethDocDx
End Property

Public Property Set TeethDocTx(oTeethDocTx As Object)
    Set m_TeethDocTx = oTeethDocTx
End Property

Public Property Get TeethDocTx() As Object
    Set TeethDocTx = m_TeethDocTx
End Property
Public Property Get TeethDocDx603a() As Object
    Set TeethDocDx603a = m_TeethDocDx603a
End Property

Public Property Set TeethWinDx(oTeethWinDx As Object)
    Set m_TeethWinDx = oTeethWinDx
End Property

Public Property Get TeethWinDx() As Object
    Set TeethWinDx = m_TeethWinDx
End Property
Public Property Set TeethWinTx(oTeethWinTx As Object)
    Set m_TeethWinTx = oTeethWinTx
End Property

Public Property Set TeethWinDx603a(oTeethWinDx603a As Object)
    Set m_TeethWinDx603a = oTeethWinDx603a
End Property

Public Property Get TeethWinTx() As Object
    Set TeethWinTx = m_TeethWinTx
End Property

Public Property Get TeethWinDx603a() As Object
    Set TeethWinDx603a = m_TeethWinDx603a
End Property

Public Function initDx(myWin As Object)
  
    Set m_TeethDocDx = myWin.myDoc
    Set m_TeethWinDx = myWin
    
End Function

Public Function initDx603a(myWin As Object)
  
    Set m_TeethDocDx603a = myWin.myDoc
    Set m_TeethWinDx603a = myWin
    
End Function

Public Function initTx(myWin As Object)
  
    Set m_TeethDocTx = myWin.myDoc
    Set m_TeethWinTx = myWin
    
End Function

'Public Sub SetDXTXMode(bIsDx As Boolean)
'    If bIsDx Then
'        Set m_gTeethDoc = m_TeethDocDx
'        Set m_gTeethWin = m_TeethWinDx
'    Else
'        Set m_gTeethDoc = m_TeethDocTx
'        Set m_gTeethWin = m_TeethWinTx
'    End If
'
'    mbIsDxView = bIsDx
'    If Not m_ToothObject Is Nothing Then
'        m_ToothObject.mbIsDxView = bIsDx
'        If bIsDx Then
'            m_ToothObject.DxTxMode = eDxView
'        Else
'            m_ToothObject.DxTxMode = eTxView
'        End If
'    End If
'    setTeethView
'
'End Sub

Public Function SvgClickTeeth(nTooth As Integer, evt As Object)
        
    If Not evt Is Nothing Then
        If evt.detail = DB_CLICK And bShift = False And bCtrl = False Then
        'If miCurrentTooth = nTooth And evt.detail = DB_CLICK And bShift = False Then
            Call dbClickTeeth(nTooth)
            'Call selectedToothOnly(convNumToPos(nTooth), SEL)
            Exit Function
        End If
    End If
    
    'Commented this out because the current tooth can be set although the tooth is not highlighted yellow
'    If miCurrentTooth = nTooth And bShift = False And bCtrl = False Then
'        Exit Function
'    End If
    
    'The user clicked on a tooth, so change to that tooth, but don't refresh the grid because they may have
    'already selected a Dx or Tx
    Call ClickTeeth(nTooth, False)
       
End Function
Public Function ClickTeeth(ByVal nTooth As Integer, RefreshGrid As Boolean)

'Dim i As Integer
'Dim strArrTeethPath() As String
'Dim iCurrentToothIndex As Integer
Dim iStartPos As Integer
Dim iEndPos As Integer
Dim iTempPos As Integer
Dim iCounter As Integer
Dim TempToothNum As Integer

On Error GoTo ErrHnd

    Set mColSelectedTeeth.m_gTeethDoc = m_gTeethDoc
    Set mColSelectedTeeth.m_gTeethWin = m_gTeethWin
    
    'strarrteethpath = Split(m_sAgePath, ",")
    
    If Not bCtrl And Not bShift Then
        'User selected one tooth
        
        mColSelectedTeeth.removeAll
        
        On Error GoTo ToothNotFoundErr
        mColSelectedTeeth.Add mObjSVGMngr.mColTeeth(CStr(nTooth))
        'If the line above gets an error mColSelectedTeeth.count will be 0
        On Error GoTo ErrHnd
        If mColSelectedTeeth.Count > 0 Then
            moPositions.GetPositionByToothNum(nTooth).Seen = True
        End If


'        If nTooth < 33 Then
'            aTeethSeen(nTooth) = Seen
'        Else
'            aTeethSeen(ConvNumToPos(nTooth)) = Seen
'        End If
        
    ElseIf bCtrl Then
        'User is adding another tooth to the current selection
        
        'Don't allow a user to add a non-erupted tooth to erupted teeth
        If moPositions.GetPositionByToothNum(nTooth).ToothErupted = False Then
            MsgBox "You cannot select un-erupted teeth with other teeth.", vbInformation, "Select tooth"
            bCtrl = False
            Exit Function
        End If
        
        'If the user selects a Tx with a max or man arch and one of the teeth have switched from primary
        'to secondary or vice versa, the list of teeth will be incorrect. so counter-act for it.
        'Ex the teeth to select are 1,2,3,4,5,6,7,8,9,10.
        'But the teeth in the mouth are 1,2,3,A,5,6,7,8,9,10.
        If moPositions.GetPositionByToothNum(nTooth).ToothNum = nTooth Then
            mColSelectedTeeth.Add mObjSVGMngr.mColTeeth(CStr(nTooth))
            moPositions.GetPositionByToothNum(nTooth).Seen = True
        Else
            'Do nothing because the tooth is no longer part of the max or man selection
'            TempToothNum = moPositions.GetPositionByToothNum(nTooth).ToothNum
'            mColSelectedTeeth.Add mObjSVGMngr.mColTeeth(CStr(TempToothNum))
'            moPositions.GetPositionByToothNum(TempToothNum).Seen = True
        End If
        
'        If nTooth < 33 Then
'            aTeethSeen(nTooth) = Seen
'        Else
'            aTeethSeen(ConvNumToPos(nTooth)) = Seen
'        End If
       
    ElseIf bShift And Not miCurrentTooth = 0 Then
        'User wants to select a range of teeth
        
        mColSelectedTeeth.removeAll
        
        'Get the starting position from miCurrentTooth
        iStartPos = moPositions.GetPositionByToothNum(miCurrentTooth).Position
        
'        'Begin nTooth to the current tooth position in string path
'        For i = 0 To UBound(strArrTeethPath)
'
'            If miCurrentTooth = strArrTeethPath(i) Then
'               iCurrentToothIndex = i
'               Exit For
'            End If
'        Next i
        
        'Get the ending position from the selected tooth passed to this function
        iEndPos = moPositions.GetPositionByToothNum(nTooth).Position
        
        'Set the start position as the least number and the end position is the greater number
        If iEndPos < iStartPos Then
            iTempPos = iStartPos
            iStartPos = iEndPos
            iEndPos = iTempPos
        End If
        
        'Add the teeth between start pos and end pos
        For iCounter = iStartPos To iEndPos
            If moPositions(iCounter).ToothErupted Then
                mColSelectedTeeth.Add mObjSVGMngr.mColTeeth(CStr(moPositions(iCounter).ToothNum))
                '39108
                moPositions(iCounter).Seen = True
            End If
        Next
        
        
'        For i = 0 To UBound(strArrTeethPath)
'            mColSelectedTeeth.Add mObjSVGMngr.mColTeeth(strArrTeethPath(iCurrentToothIndex))
'
'            moPositions.GetPositionByToothNum(iCurrentToothIndex).Seen = True
''            If CInt(strArrTeethPath(iCurrentToothIndex)) < 33 Then
''                aTeethSeen(CInt(strArrTeethPath(iCurrentToothIndex))) = Seen
''            Else
''                aTeethSeen(ConvNumToPos(CInt(strArrTeethPath(iCurrentToothIndex)))) = Seen
''            End If
'
'            If strArrTeethPath(iCurrentToothIndex) = nTooth Then Exit For
'
'            iCurrentToothIndex = iCurrentToothIndex + 1
'
'           If iCurrentToothIndex = UBound(strArrTeethPath) + 1 Then  'At the end of the array lets start over cases T:31-T:2
'              iCurrentToothIndex = 0
'           End If
'
'        Next i
        
        
        bShift = False
    
    End If
    
    If mColSelectedTeeth.Count = 1 Then
        RaiseEvent ToothChanged(nTooth, miCurrentTooth, RefreshGrid)
        m_ToothObject.ToothChanged nTooth, , True
        RaiseEvent SelectionCountChange(True)
    Else
        RaiseEvent SelectionCountChange(False)
    End If
    
    'Set last known tooth to one just selected
    miCurrentTooth = nTooth

Exit Function

ErrHnd:
    MsgBox Err.Description & ": Cannot find tooth:" & nTooth & " in the positions collection", , "ClickTeeth - Error"
    Resume Next
    Resume
Exit Function

ToothNotFoundErr:
'Ignore if the tooth cannot be found in mObjSVGMngr.mColTeeth(CStr(nTooth))
    Resume Next

End Function
Public Function ShiftKey(bKeyDown As Boolean)
    bShift = bKeyDown
End Function

Public Function CtrlKey(bKeyDown As Boolean)
    bCtrl = bKeyDown
End Function

Public Function getSingleTooth(ByVal nTooth As Integer) As String
    ' called by Dx or Tx on the mode of full mouth
    Dim oSvg As Object
       
    On Error GoTo ErrorHandler
    
    If m_gTeethDoc Is Nothing Then
        Set m_gTeethDoc = m_TeethDocDx
        Set m_gTeethWin = m_TeethWinDx
    End If
    
    ' get tooth info from teeth doc
    Set oSvg = m_gTeethDoc.getElementById("svg" & CStr(nTooth))
    getSingleTooth = m_gTeethWin.printNode(oSvg)
    Exit Function
    
ErrorHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "clsTeeth.getSingleTooth", "srvSVGComm", vbCritical)
    
End Function

Public Function setSingleTooth(ByVal nTooth As Integer, ByVal sTooth As String)
    ' called by Dx or Tx on the mode of full mouth
    Dim oSvg, oSSvg As Object
       
    On Error GoTo ErrorHandler
    
    If m_gTeethDoc Is Nothing Then
        Set m_gTeethDoc = m_TeethDocDx
        Set m_gTeethWin = m_TeethWinDx
    End If
    
    Set oSvg = m_gTeethDoc.getElementById("svg0")
    
    ' get tooth info from teeth doc
    Set oSSvg = m_gTeethDoc.getElementById("svg" & CStr(nTooth))
    If Not oSSvg Is Nothing Then
        oSvg.removeChild m_gTeethDoc.getElementById("svg" & CStr(nTooth))
    End If
    
    oSvg.appendChild m_gTeethWin.parseXML(sTooth, m_gTeethDoc)
    Exit Function
    
ErrorHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "clsTeeth.getSingleTooth", "srvSVGComm", vbCritical)
    
End Function

Public Function DisplaySingleTooth(ByVal nTooth As Integer, Optional IsToothVisible As Boolean = True)

    Dim oSvg, oSSvg, oRect, oToothDoc As Object
    Dim x, y, nLastViewTooth As Integer
    Dim s As String
    Dim iToothNumberToDisplay As Integer
    
    On Error GoTo ErrorHandler
    
    If moPositions.GetPositionByToothNum(nTooth).ToothNum <> nTooth Then
        'The tooth is not in the patients mouth, so don't draw it.
        Exit Function
    End If
    
    ' set tooth doc
    Set oToothDoc = m_ToothObject.ToothDoc
    nLastViewTooth = m_ToothObject.LastViewTooth
    
    m_ToothObject.resetSelectedSurfaces
    
    ' set the background to white
    Set oRect = m_gTeethDoc.getElementById(CStr(nTooth))
'    Set oRect = m_gTeethDoc.getElementById("svg" & CStr(nTooth))
    If Not oRect Is Nothing Then
        Call oRect.setAttribute("class", WHITE_HIDE)
    Else
        'MsgBox "Could not find the tooth: svg" & nTooth & " to set it to white (hide it) in the teeth document."
    End If
    
    ' get the original tooth document
    Set oSvg = oToothDoc.getElementById("svg0")
      
    
     ' remove the old one from big tooth
    Set oSSvg = oToothDoc.getElementById("svg" & CStr(nLastViewTooth))
    If Not oSSvg Is Nothing Then
        oSvg.removeChild oToothDoc.getElementById("svg" & CStr(nLastViewTooth))
        'Added this to test 12-27
        'oSSvg.removeChild oToothDoc.getElementById("svg" & CStr(nLastViewTooth))
    Else
        'MsgBox "Could not find the tooth: svg" & nLastViewTooth & " to remove it from the tooth document."
    End If
   
    ''''''''''''''''''''''''''''''''''
    's = m_gTeethWin.printNode(oSvg)
    'Open "c:\teeth4_out.svg" For Output As #1
    'Print #1, s
    'Close #1
    ''''''''''''''''''''''''''''''''''
    
    ' insert the tooth
    ' get tooth from teeth
    
    '*********************
    'SF replaced nTooth with nLastViewTooth
    'SF replaced nLastViewTooth with iToothNumberToDisplay
    
    If IsToothVisible = True Then
        iToothNumberToDisplay = nTooth
    Else
        iToothNumberToDisplay = nLastViewTooth
    End If
    
    'Get the tooth from the TEETH document
    Set oSSvg = m_gTeethDoc.getElementById("svg" & CStr(nLastViewTooth))
    
'Not sure why the if was here, it causes problems after switching from primary to secondary and vice versa
'    If Not oSSvg Is Nothing Then
    
        'todo: this may be a problem because DxView does not look like it is correctly set in BigTooth
        '37476
        If meDxTxViewMode = eDxView Then
        'If mbIsDxView Then
            s = mObjSVGMngr.getsvgDxViewbyTooth(iToothNumberToDisplay, eBigTooth) 'm_gTeethWin.printNode(oSSvg)
        ElseIf meDxTxViewMode = eTxView Then
            s = mObjSVGMngr.getsvgTxViewbyTooth(iToothNumberToDisplay, eBigTooth)
        ElseIf meDxTxViewMode = eDx603aView Then
            s = mObjSVGMngr.getsvgDx603aViewbyTooth(iToothNumberToDisplay, eBigTooth)
        End If
        
        'Add the new tooth to the tooth document
        oSvg.appendChild m_gTeethWin.parseXML(s, oToothDoc)
           
        'Add the new tooth to the teeth document
'        oSvg.appendChild m_gTeethWin.parseXML(s, m_gTeethDoc)
                
        '****************** Added to test
        'If oRect Is Nothing Then
            Set oRect = m_gTeethDoc.getElementById(CStr(iToothNumberToDisplay))
            'Set oRect = m_gTeethDoc.getElementById("svg" & CStr(iToothNumberToDisplay))
        'End If
        
        'If oRect is still nothing, then don't do anything.
        If Not oRect Is Nothing Then
            Dim x_svg, y_svg, w_svg, h_svg As Integer
            x_svg = oRect.getBBox.x
            y_svg = oRect.getBBox.y

            Set oSSvg = oToothDoc.getElementById("svg" & CStr(iToothNumberToDisplay))
            If iToothNumberToDisplay > 16 Then 'This may only work for secondary teeth
                oSSvg.setAttribute "viewBox", CStr(x_svg - 10) & " " & CStr(y_svg + 45) & " 185% 180%"
            Else
                oSSvg.setAttribute "viewBox", CStr(x_svg - 10) & " " & CStr(y_svg + 50) & " 185% 180%"
            End If
        Else
            '*** MsgBox Fix - Code_Validator auto change made 3/21/2003
            MsgBox "Cannot draw the tooth because the tooth is not in the patient's mouth"
        End If
          
    m_ToothMode = eBigTooth
    
    m_ToothObject.BTFMMode = eBigTooth
    miCurrentTooth = iToothNumberToDisplay
    m_ToothObject.ToothChanged (miCurrentTooth)
    m_ToothObject.LastViewTooth = iToothNumberToDisplay
    
    '***********************
    
    Exit Function
    
ErrorHandler:
    If Err.Number = 438 Then 'GC has not loaded, so don't update any graphics
        Exit Function
    Else
        Call gobjShared.ShowVBError(Err.Number, Err.Description, "clsTeeth.DisplaySingleTooth", "srvSVGComm", vbCritical)
    End If
    
    Exit Function
    Resume 'Debug
End Function


Public Function dbClickTeeth(nTooth As Integer)
    
    Call DisplaySingleTooth(nTooth)
    RaiseEvent dblClickMultiTooth(nTooth)
    
    mColSelectedTeeth.removeAll
    mColSelectedTeeth.Add mObjSVGMngr.mColTeeth(CStr(nTooth))
    'SF 10-31-03 Added this to enable the pri/sec button when the user first enters bigtooth
    RaiseEvent SelectionCountChange(True)

End Function

Public Function ClickPanel(nToothPosOrNum As Integer, Optional bForceful As Boolean = False)
    
    Dim iToothPos As Integer
    
    iToothPos = moPositions.ConvertNumToPos(nToothPosOrNum)

    'If aTeethStatus(iToothPos) = NONE Then
    'If moPositions(iToothPos).Status = ePositionStatus.eNone Then
    
'    If moPositions(iToothPos).ToothErupted = False Then
'        'Need to modify to work with the invisible tooth
'        Debug.Assert False
'       ' don't think this needs to be here.
'        Exit Function
'    End If
    
    Dim nTooth As Integer
    nTooth = iToothPos
  
    'If aTeethStatus(iToothPos) = PRIMARY Then
    If moPositions(iToothPos).Status = ePositionStatus.ePrimary Then
         nTooth = moPositions.ConvertPosToPrimaryNum(iToothPos)
    End If
    
    Dim nOldTooth As Integer
    nOldTooth = miCurrentTooth
    
    If miCurrentTooth <> nTooth Or bForceful = True Then
        'Call updateTeeth(miCurrentTooth)
        'Call refreshTeeth(nTooth)
        Call DisplaySingleTooth(nTooth)
        
        'aTeethSeen(iToothPos) = Seen
        moPositions(iToothPos).Seen = True
        
        'Call selectedToothOnly(iToothPos, SEL)
        
        'The user wants to change the tooth in big tooth mode, but don't refresh the grid because
        'they may have already selected a Dx or Tx to link to.
        RaiseEvent ToothChanged(nTooth, nOldTooth, False)
    End If
    
    
    If Not bDblClick Or m_ToothObject.BTFMMode = eBigTooth Then
       mColSelectedTeeth.removeAll
       mColSelectedTeeth.Add mObjSVGMngr.mColTeeth(CStr(nTooth))
       RaiseEvent SelectionCountChange(True)
    Else
       bDblClick = False
    End If
    
End Function

Public Function dbClickPanel(nToothPos As Integer)
    
    Dim iTooth As Integer
    
    bDblClick = True
    
    iTooth = nToothPos
  
'    Debug.Assert False
    
    'If aTeethStatus(nToothPos) = PRIMARY Then
    
    iTooth = moPositions(nToothPos).ToothNum
    
'    If moPositions(nToothPos).Status = ePositionStatus.ePrimary Then
'        iTooth = ConvPosToPrimaryNum(nToothPos)
'    End If
        
    ' that is called by double clicking the tooth label
    Dim oRect, oSvg, oSSvg As Object
    Dim i As Integer
    Dim s, ss As String
     
    ' set current tooth
    miCurrentTooth = iTooth
    
    'jrm 5/06/02 Ensure eFullMouth happens first
    m_ToothObject.BTFMMode = eFullMouth
    m_ToothObject.ToothChanged miCurrentTooth, , True
    
    'Refresh to zero state
    RaiseEvent singleClickMultiTooth(iTooth)
        
    mColSelectedTeeth.removeAll
    RaiseEvent SelectionCountChange(False)
    'Call updateTeeth(iTooth)
'    Call refreshTeeth(iTooth)

    m_ToothMode = eFullMouth
          
End Function

Public Function UpdateTeeth(ByVal nTooth As Integer, Optional bRefreshSelected As Boolean = True)
    
    Dim oRect, oSvg, oSSvg As Object
    Dim i As Integer
    Dim s, ss, sToothStr As String
    
    On Error GoTo ErrorHandler
    
    If m_gTeethDoc Is Nothing Then
        Set m_gTeethDoc = m_TeethDocDx
        Set m_gTeethWin = m_TeethWinDx
    End If
    
    ' set tooth doc
    '37476
    If meDxTxViewMode = eDxView Then
    'If mbIsDxView Then
        sToothStr = mObjSVGMngr.getsvgDxViewbyTooth(nTooth, eFullMouth) 'm_ToothObject.ToothStr
    ElseIf meDxTxViewMode = eTxView Then
        sToothStr = mObjSVGMngr.getsvgTxViewbyTooth(nTooth, eFullMouth)
    ElseIf meDxTxViewMode = eDx603aView Then
        sToothStr = mObjSVGMngr.getsvgDx603aViewbyTooth(nTooth, eFullMouth)
    End If
    
    Set oSvg = m_gTeethDoc.getElementById("svg0")
    
    ' remove old one
    Set oSSvg = m_gTeethDoc.getElementById("svg" & CStr(nTooth))
    If Not oSSvg Is Nothing Then
        oSvg.removeChild m_gTeethDoc.getElementById("svg" & CStr(nTooth))
    End If
        
    If sToothStr <> "" Then
           
        ' remove old one
        'Set oSSvg = m_gTeethDoc.getElementById("svg" & CStr(nTooth))
        'If Not oSSvg Is Nothing Then
        '    oSvg.removeChild m_gTeethDoc.getElementById("svg" & CStr(nTooth))
        'End If
                
        ' insert new one
        'sToothStr = Replace(sToothStr, YELLOW, WHITE) ' get rid of yellow color
        sToothStr = Replace(sToothStr, YELLOW2, WHITE_HIDE) ' get rid of yellow color
        
        m_ToothObject.resetSelectedSurfaces
        'Debug.Print sToothStr
          
        oSvg.appendChild m_gTeethWin.parseXML(sToothStr, m_gTeethDoc)
       
        ' set view box
        Set oSSvg = m_gTeethDoc.getElementById("svg" & CStr(nTooth))
        oSSvg.setAttribute "viewBox", ""
        
        m_ToothObject.ToothStr = ""
        
        If bRefreshSelected Then
                Call refreshTeeth(nTooth)
        Else
                Set oRect = m_gTeethDoc.getElementById(CStr(nTooth))
                
                'If Not oRect Is Nothing And aTeethSeen(nTooth) = NOT_SEEN Then
                If Not oRect Is Nothing And moPositions.GetPositionByToothNum(nTooth).Seen = False Then
                    Call oRect.setAttribute("class", GREY_HIDE)
                'ElseIf Not oRect Is Nothing And aTeethSeen(nTooth) = Seen Then
                ElseIf Not oRect Is Nothing And moPositions.GetPositionByToothNum(nTooth).Seen = True Then
                    Call oRect.setAttribute("class", WHITE_HIDE)
                End If
        End If
                
    End If
    
    'm_ToothObject.ToothStr = ""
    'Call refreshTeeth(nTooth)

        
    Exit Function
    
ErrorHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "clsTeeth.updateTeeth", "srvSVGComm", vbCritical)
    
End Function
Public Sub UpdateTeethSVG()
    
    Dim oRect, oSvg, oSSvg As Object
    Dim i As Integer
    Dim s, ss, sToothStr As String
    Dim max() As String
    Dim iTooth As Integer
    Dim tmpSVG As String
    
    On Error GoTo ErrorHandler
    
    If m_gTeethDoc Is Nothing Then
        Set m_gTeethDoc = m_TeethDocDx
        Set m_gTeethWin = m_TeethWinDx
    End If
    
    ' set tooth doc
    max = Split(MaxillaryTeeth, ",")
    '37476
    If meDxTxViewMode = eDxView Then
    'If mbIsDxView Then
        sToothStr = mObjSVGMngr.getsvgDxViewTeeth(max)
    Else
'This is strange, why was this commented out?
'        sToothStr = mObjSVGMngr.getTxViewTooth(nTooth, False)
    End If
    
    Set oSvg = m_gTeethDoc.getElementById("svg0")
    
    ' remove old one
    For iTooth = 0 To UBound(max)
        Set oSSvg = m_gTeethDoc.getElementById("svg" & CStr(max(iTooth)))
        If Not oSSvg Is Nothing Then
            oSvg.removeChild m_gTeethDoc.getElementById("svg" & CStr(max(iTooth)))
        End If
    Next iTooth
    If sToothStr <> "" Then
           
        ' remove old one
        'Set oSSvg = m_gTeethDoc.getElementById("svg" & CStr(nTooth))
        'If Not oSSvg Is Nothing Then
        '    oSvg.removeChild m_gTeethDoc.getElementById("svg" & CStr(nTooth))
        'End If
                
        ' insert new one
        'sToothStr = Replace(sToothStr, YELLOW, WHITE) ' get rid of yellow color
'        sToothStr = Replace(sToothStr, YELLOW2, WHITE_HIDE) ' get rid of yellow color
        
'        m_ToothObject.resetSelectedSurfaces
        'Debug.Print sToothStr
        '<svg id="svg1"
          tmpSVG = "<svg id=" & Chr(147) & "svgTeeth" & Chr(148) & ">" & sToothStr & " </svg>"
        oSvg.appendChild m_gTeethWin.parseXML(tmpSVG, m_gTeethDoc)
       
        ' set view box
'        Set oSSvg = m_gTeethDoc.getElementById("svg" & CStr(nTooth))
'        oSSvg.setAttribute "viewBox", ""
        
'        m_ToothObject.ToothStr = ""
'
'        If bRefreshSelected Then
'                Call refreshTeeth(nTooth)
'        Else
'                Set oRect = m_gTeethDoc.getElementById(CStr(nTooth))
'
'                If nTooth > 32 Then
'                    nTooth = convNumToPos(nTooth)
'                End If
'
'                If Not oRect Is Nothing And aTeethSeen(nTooth) = NOT_SEEN Then
'                    Call oRect.setAttribute("class", GREY_HIDE)
'                ElseIf Not oRect Is Nothing And aTeethSeen(nTooth) = SEEN Then
'                    Call oRect.setAttribute("class", WHITE_HIDE)
'                End If
'        End If
'
    End If
    
    'm_ToothObject.ToothStr = ""
    'Call refreshTeeth(nTooth)

        
    Exit Sub
    
ErrorHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "clsTeeth.updateTeeth", "srvSVGComm", vbCritical)
    
End Sub
Private Sub refreshTeeth(nTooth As Integer)

    Dim oRect As Object
    Dim i, nToothNum As Integer
    Dim oPos As clsPosition
    
    ' set current tooth to yellow
    Set oRect = m_gTeethDoc.getElementById(CStr(nTooth))
    If Not oRect Is Nothing Then
        Call oRect.setAttribute("class", YELLOW_HIDE)
    End If
    
'    ' set other yellow ones to white
'    For i = 1 To TOTAL_TEETH
'        nToothNum = i
'        'If aTeethStatus(i) = PRIMARY Then
'        If moPositions(i).Status = ePositionStatus.ePrimary Then
'            nToothNum = ConvPosToPrimaryNum(i)
'        End If
'        Set oRect = m_gTeethDoc.getElementById(CStr(nToothNum))
'        If Not oRect Is Nothing Then
'            If nToothNum <> nTooth Then
'                If oRect.getAttribute("class") = YELLOW_HIDE Then
'                    Call oRect.setAttribute("class", WHITE_HIDE)
'                End If
'            End If
'        End If
'    Next

    ' set other yellow ones to white
    For i = 1 To moPositions.Count
        nToothNum = moPositions(i).ToothNum
        
        Set oRect = m_gTeethDoc.getElementById(CStr(nToothNum))
        If Not oRect Is Nothing Then
            If nToothNum <> nTooth Then
                If oRect.getAttribute("class") = YELLOW_HIDE Then
                    Call oRect.setAttribute("class", WHITE_HIDE)
                End If
            End If
        End If
    Next

End Sub
       
Public Function ClickPrimary(ViewMode As eViewMode) As Boolean

Dim GCTooth As clsGCTooth
Dim iNextTooth As Integer
Dim oPos As clsPosition
Dim iPrevTooth As Integer

    'This function is used to switch from primary to secondary and vice versa
    
    'If a tooth isn't selected, then exit.
    If miCurrentTooth = 0 Then
        ClickPrimary = False
        Exit Function
    End If
    
    On Error GoTo ErrorHandler
    
    'Debug.Assert False
    
    If m_gTeethDoc Is Nothing Then
        Set m_gTeethDoc = m_TeethDocDx
        Set m_gTeethWin = m_TeethWinDx
    End If
    
    'Reset the selected teeth
    Set mColSelectedTeeth.m_gTeethDoc = m_gTeethDoc
    Set mColSelectedTeeth.m_gTeethWin = m_gTeethWin
    mColSelectedTeeth.removeAll
    
    Set oPos = moPositions.GetPositionByToothNum(miCurrentTooth)
    'If the tooth has not erupted, then erupt the tooth
    If oPos.ToothErupted = False Then
        oPos.ToothErupted = True
    Else
        'Switch the status of the position
        If oPos.Status = ePositionStatus.ePrimary Then
            oPos.Status = ePositionStatus.eSecondary
        Else
            oPos.Status = ePositionStatus.ePrimary
            'Should not be able to click on the pri/sec button if the tooth has erupted and the position
            'can't have a primary tooth
        End If
    End If
    
    iNextTooth = oPos.ToothNum

'    If miCurrentTooth > 32 Then
'        'The current tooth is primary, so give me the secondary equivalent.
'        iNextTooth = ConvNumToPos(miCurrentTooth)
'    Else
'        'The current tooth is secondary, so give me the primary equivalent.
'        iNextTooth = ConvPosToPrimaryNum(miCurrentTooth)
'    End If
    
'    'If there isn't a secondary or primary equivalent tooth, then exit.
'    If miCurrentTooth = iNextTooth Then
'        ClickPrimary = False
'        Exit Function
'    End If
    
    'JRM What is visible
    'Hide the current tooth.
    'Set GCTooth = mObjSVGMngr.mColTeeth(CStr(miCurrentTooth))
    'GCTooth.bToothVisible = False
    'Hide the primary or secondary tooth
    mObjSVGMngr.mColTeeth(CStr(miCurrentTooth)).ToothVisible = False
    
    'Remove the current tooth from the svg document.
    Dim oRect As Object
    Set oRect = m_gTeethDoc.getElementById("svg0")
    
    If Not oRect Is Nothing Then
        oRect.removeChild m_gTeethDoc.getElementById("svg" & CStr(miCurrentTooth))
    Else
        MsgBox "Could not find tooth: " & miCurrentTooth & " in the svg document."
        ClickPrimary = False
        Exit Function
    End If
    
    ''''''''''''''''''''''
    Dim s As String
    
    'Add the next tooth to the collection of teeth (if it doesn't exist in the mouth, it will add it)
    mObjSVGMngr.drawSVG (iNextTooth)
    
    'Make sure the new tooth is visible
    Set GCTooth = mObjSVGMngr.mColTeeth(CStr(iNextTooth))
    GCTooth.ToothVisible = True
    GCTooth.IsToothPlaceHolder = Not oPos.ToothErupted
    'Get the svg string for the Dx view or the Tx view.
    If ViewMode = eDxView Then
        s = mObjSVGMngr.getsvgDxViewbyTooth(iNextTooth, eFullMouth) ''WARNING, THESE MUST BE FULLMOUTH ALL THE TIME, NOT m_ToothMode - or the tooth number doesn't print at the botton of the tooth
    ElseIf ViewMode = eTxView Then
        s = mObjSVGMngr.getsvgTxViewbyTooth(iNextTooth, eFullMouth)
    ElseIf ViewMode = eDx603aView Then
        s = mObjSVGMngr.getsvgDx603aViewbyTooth(iNextTooth, eFullMouth)
    End If
    
    'Add the new tooth to the svg document.
    ' replace the tooth
    oRect.appendChild m_gTeethWin.parseXML(s, m_gTeethDoc)
    
    ' set the background
'    Set oRect = m_gTeethDoc.getElementById(CStr(iNextTooth))
    If Not oRect Is Nothing Then
'        Call oRect.setAttribute("class", YELLOW)
    Else
        ClickPrimary = False
        Exit Function
    End If
    
    
    'The current tooth becomes the previous tooth
    iPrevTooth = miCurrentTooth
    
    'The next tooth is now the current tooth.
    miCurrentTooth = iNextTooth
    
    'Update the menus based on the new tooth
    m_ToothObject.ToothChanged miCurrentTooth, , True
    
    'Determine the tooth position given the tooth number.
'    Dim nToothPos As Integer
'    nToothPos = ConvNumToPos(miCurrentTooth)
'
'    aTeethSeen(nToothPos) = Seen
    moPositions.GetPositionByToothNum(miCurrentTooth).Seen = True
    
'    If aTeethStatus(nToothPos) = PRIMARY Then
'        aTeethStatus(nToothPos) = SECONDARY
'    ElseIf aTeethStatus(nToothPos) = SECONDARY Then
'        aTeethStatus(nToothPos) = PRIMARY
'    End If

'    If moPositions.GetPositionByToothNum(miCurrentTooth).Status = ePositionStatus.ePrimary Then
'        moPositions.GetPositionByToothNum(miCurrentTooth).Status = ePositionStatus.eSecondary
'    ElseIf moPositions.GetPositionByToothNum(miCurrentTooth).Status = ePositionStatus.eSecondary Then
'        moPositions.GetPositionByToothNum(miCurrentTooth).Status = ePositionStatus.ePrimary
'    End If
    
    ' update age path
'    m_sAgePath = getPathTeeth()
'    m_sAgePath = moPositions.TeethPath
    
    'Update the other views
'    If m_ToothMode = False Then
    If ViewMode = eDxView Then
        Call UpdateOtherViews(iPrevTooth, eTxView)
        Call UpdateOtherViews(iPrevTooth, eDx603aView)
    ElseIf ViewMode = eTxView Then
        Call UpdateOtherViews(iPrevTooth, eDxView)
        Call UpdateOtherViews(iPrevTooth, eDx603aView)
    ElseIf ViewMode = eDx603aView Then
        Call UpdateOtherViews(iPrevTooth, eDxView)
        Call UpdateOtherViews(iPrevTooth, eTxView)
    End If
'    End If
    
    ClickPrimary = True
    
    'This doesn't do anything.
    RaiseEvent priSecondClick(miCurrentTooth)
    
    'Tell bigtooth.cls that the selection changed, so it can enable/disable the buttons on frmgraphicalcharting.frm
    RaiseEvent SelectionCountChange(True)
    
    Exit Function
    
ErrorHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "clsTeeth.ClickPrimary", "srvSVGComm", vbCritical)
   Resume Next
End Function

Private Function UpdateOtherViews(PrevTooth As Integer, ViewMode As eViewMode)
    Dim s As String
    Dim oTeethDoc, oTeethWin, oRect As Object
    
    'Only used when changing a tooth from primary to secondary
    
    'Update the Tx document if the user changed the current tooth to primary or secondary in the Dx document.
    'and update the Dx document if the user changed the current tooth to primary or secondary in the Tx document.

On Error GoTo ErrHnd

    If ViewMode = eDxView Then
        ' do Tx here
        s = mObjSVGMngr.getsvgDxViewbyTooth(miCurrentTooth, eFullMouth) 'WARNING, THESE MUST BE FULLMOUTH ALL THE TIME, NOT m_ToothMode - or the tooth number doesn't print at the botton of the tooth
        Set oTeethDoc = m_TeethDocDx
        Set oTeethWin = m_TeethWinDx
    ElseIf ViewMode = eTxView Then
        ' do Dx here
        s = mObjSVGMngr.getsvgTxViewbyTooth(miCurrentTooth, eFullMouth)
        Set oTeethDoc = m_TeethDocTx
        Set oTeethWin = m_TeethWinTx
    ElseIf ViewMode = eDx603aView Then
        ' do Dx here
        s = mObjSVGMngr.getsvgDx603aViewbyTooth(miCurrentTooth, eFullMouth)
        Set oTeethDoc = m_TeethDocDx603a
        Set oTeethWin = m_TeethWinDx603a
    End If
    
    
'    If m_gTeethDoc Is m_TeethDocDx Then
'        ' do Tx here
'        s = mObjSVGMngr.getsvgTxViewbyTooth(miCurrentTooth, False)
'        Set oTeethDoc = m_TeethDocTx
'        Set oTeethWin = m_TeethWinTx
'    Else
'        ' do Dx here
'        s = mObjSVGMngr.getsvgDxViewbyTooth(miCurrentTooth, False)
'        Set oTeethDoc = m_TeethDocDx
'        Set oTeethWin = m_TeethWinDx
'    End If
   
    If Not oTeethDoc Is Nothing Then
        Set oRect = oTeethDoc.getElementById("svg0")
        
'        Debug.Assert False
'        'This needs to be removed and replace PrevTooth with the var from the calling function
'
'        Dim nPreviousTooth As Integer
'        If miCurrentTooth > 32 Then
'            'Given the primary tooth, get the secondary tooth
'            nPreviousTooth = moPositions.ConvertNumToPos(miCurrentTooth)
'        Else
'            'Given the secondary tooth, get the primary tooth
'            nPreviousTooth = moPositions.ConvertPosToPrimaryNum(miCurrentTooth)
'        End If
'
        'Remove the old tooth
        If Not oRect Is Nothing Then
            oRect.removeChild oTeethDoc.getElementById("svg" & CStr(PrevTooth))
        End If
        
        'Add the new tooth
        oRect.appendChild oTeethWin.parseXML(s, oTeethDoc)
        
    End If
     
Exit Function

ErrHnd:
    If Err.Number = -2147352567 Then
        'Error removing the node from the document (could not find it).
        Resume Next
    Else
        Call gobjShared.ShowVBError(Err.Number, Err.Description, "clsTeeth.ClickPrimaryDxTx", "srvSVGComm", vbCritical)
    End If
    
End Function
    
Public Function ClickMolar()
'    Dim oRect As Object
'    Dim i As Integer
'
    On Error GoTo ErrorHandler
    
    Set mColSelectedTeeth.m_gTeethDoc = m_gTeethDoc
    Set mColSelectedTeeth.m_gTeethWin = m_gTeethWin
        
    mColSelectedTeeth.removeAll
    mColSelectedTeeth.Add mObjSVGMngr.mColTeeth(CStr(1))
    mColSelectedTeeth.Add mObjSVGMngr.mColTeeth(CStr(16))
    mColSelectedTeeth.Add mObjSVGMngr.mColTeeth(CStr(17))
    mColSelectedTeeth.Add mObjSVGMngr.mColTeeth(CStr(32))
        
'    aTeethSeen(1) = Seen
'    aTeethSeen(16) = Seen
'    aTeethSeen(17) = Seen
'    aTeethSeen(32) = Seen
    moPositions(1).Seen = True
    moPositions(16).Seen = True
    moPositions(17).Seen = True
    moPositions(32).Seen = True
        
    RaiseEvent SelectionCountChange(False)
    
'    For i = 1 To TOTAL_TEETH
'        Set oRect = m_gTeethDoc.getElementById(CStr(i))
'        If Not oRect Is Nothing Then
'            If Not bMolar Then
'                If i = 1 Or i = 16 Or i = 17 Or i = 32 Then ' should be 1, 16, 17, 32?
'                    Call oRect.setAttribute("class", YELLOW_HIDE)
'                Else
'                    If oRect.getAttribute("class") = YELLOW_HIDE Or oRect.getAttribute("class") = WHITE_HIDE Then
'                        Call oRect.setAttribute("class", WHITE_HIDE)
'                    Else
'                        Call oRect.setAttribute("class", GREY_HIDE)
'                    End If
'                End If
'            Else
'                If i = 1 Or i = 16 Or i = 17 Or i = 32 Then ' should be 1, 16, 17, 32?
'                    Call oRect.setAttribute("class", WHITE_HIDE)
'                End If
'            End If
'        End If
'    Next
'
'    If bMolar Then
'        bMolar = False
'        Call selectedToothOnly(1, NOT_SEL)
'        'raise selectionCountChange(True)
'    Else
'        bMolar = True
'        Call selectedToothOnly(1, SEL)
'        Call selectedTooth(16, SEL)
'        Call selectedTooth(17, SEL)
'        Call selectedTooth(32, SEL)
'        'raise selectionCountChange(False)
'    End If
'
'    miCurrentTooth = 0
    Exit Function
    
ErrorHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "clsTeeth.ClickMolar", "srvSVGComm", vbCritical)
    
End Function

Public Function CreateTeethDx(lPatientID As Long, nAge As Integer)
       
    'Debug.Assert False
    'Should load positions be here? Yes
    LoadPositions lPatientID, nAge
    'Call getTeethPath(lPatientID, nAge)
    
    Call LoadTeeth(eViewMode.eDxView)
    
    Set mObjSVGDraw = mObjSVGMngr.mObjSVGDraw
       
End Function

Public Function CreateTeethTx()

    Set m_gTeethDoc = m_TeethDocTx
    Set m_gTeethWin = m_TeethWinTx
    
    DoEvents
    
    Call LoadTeeth(eViewMode.eTxView)
    
    DoEvents
    
End Function

Public Function CreateTeethDx603a()

    Set m_gTeethDoc = m_TeethDocDx603a
    Set m_gTeethWin = m_TeethWinDx603a
    
    DoEvents
    
    Call LoadTeeth(eViewMode.eDx603aView)
    
    DoEvents
    
End Function

Public Sub LoadPositions(PatientID As Long, Age As Integer)
    
    'Loads the positions from moPositions
    
    moPositions.LoadPositions PatientID, mObjSVGMngr.GetTeethPathFromRc(Age)
    
'    Dim sAgePath As String
'
'    sAgePath = ReadTeethPathFromDB(lPatientID) '"3,8,9,14,19,23,24,25,26,30,33,34,35,36,39,40,41,42,43,44,45,50,51,52"
'    '''''''''''''''''''''''
'
'    ' 1. read from db table dental_enc_view_teeth
'    If sAgePath = "" Then
'    ' 2. if it is new or nothing in the columns then read from the resource
''        Debug.Assert False
'        'SF 4-09-04 Not needed anymore
'        'sAgePath = readTeethPathFromRC(nAge)
'
'        sAgePath = mObjSVGMngr.GetTeethPathFromRc(nAge)
'
'        moPositions.LoadFromTeethList sAgePath
'    End If

End Sub

Public Function setTeethPath(lPatientID As Long)
    
    moPositions.Save
    
    ' save aTeethStatus and aTeethSeen to the table dental_enc_view_teeth
    'Call SaveTeethPathToDB(lPatientID)

End Function

'Private Function SaveTeethPathToDB(lPatientID As Long)
'
'    Dim objDas As ICHCSII_DAS
'    Dim objRS As ADODB.Recordset
'    Dim sSQL As String
'    Dim iStatus As Integer
'    Dim bSeen As Boolean
'
'    On Error GoTo ErrorHandler
'
'    If m_objShared Is Nothing Then
'        Set m_objShared = New CWShared
'    End If
'
'    moPositions.Save lPatientID
'
'    Set goCHCSIIConn = New CHCSII_CONN.Conn
'    Set objDas = goCHCSIIConn.CHCSII_DAS
'
''    If m_objShared.UseLocalDB Then
''        Set objDas = New GEMS_DAS
''    Else
''        Set objDas = New CHCSII_DAS
''    End If
'
'    Dim i As Integer
'    If mbNewTeethPath Then
'        ' insert into dental_enc_view_teeth
'        For i = 1 To moPositions.Count
'            With moPositions(i)
'                'Need to map database status to object status
'                If .ToothErupted = False Then
'                    iStatus = 0
'                Else
'                    iStatus = .Status
'                End If
'
'                sSQL = "INSERT INTO DENTAL_ENC_VIEW_TEETH ("
'                sSQL = sSQL & "TOOTH_STATUS, TOOTH_SEEN, TOOTH_ID, PATIENT_UNIT_ID) "
'                sSQL = sSQL & "VALUES (" & iStatus & ", " & Abs(CInt(.Seen)) & ", " & CStr(i) & ", " & lPatientID & ")"
'                objDas.ExecuteSQL (sSQL)
'
'                .StatusPre = .Status
'                .SeenPre = .Seen
'            End With
'        Next
'
'    Else
'        ' update dental_enc_view_teeth
'        For i = 1 To moPositions.Count
'            With moPositions(i)
'
'                'Need to map database status to object status
'                If .ToothErupted = False Then
'                    iStatus = 0
'                Else
'                    iStatus = .Status
'                End If
'
'                If .StatusPre <> .Status Or .SeenPre <> .Seen Then
'                    sSQL = "UPDATE DENTAL_ENC_VIEW_TEETH SET"
'                    sSQL = sSQL & " TOOTH_STATUS = " & iStatus & ", TOOTH_SEEN=" & Abs(CInt(.Seen))
'                    sSQL = sSQL & " Where TOOTH_ID = " & i & " And PATIENT_UNIT_ID = " & lPatientID
'                    objDas.ExecuteSQL (sSQL)
'
'                    .StatusPre = .Status
'                    .SeenPre = .Seen
'                End If
'            End With
'        Next
'    End If
'
'    mbNewTeethPath = False
'    Exit Function
'
'ErrorHandler:
'    Call gobjShared.ShowVBError(Err.Number, Err.Description, "clsTeeth.saveTeethPath", "srvSVGComm", vbCritical)
'
'End Function
'
'Private Function ReadTeethPathFromDB(lPatientID As Long) As String
'    Dim objDas As ICHCSII_DAS
'    Dim objRS As ADODB.Recordset
'    Dim sSQL As String
'
'    On Error GoTo ErrorHandler
'
'    If m_objShared Is Nothing Then
'        Set m_objShared = New CWShared
'    End If
'
'    Set goCHCSIIConn = New CHCSII_CONN.Conn
'    Set objDas = goCHCSIIConn.CHCSII_DAS
'
'    ' get SelectedTeeth info for this patient id from db
'    sSQL = "SELECT TOOTH_ID, TOOTH_STATUS, TOOTH_SEEN FROM DENTAL_ENC_VIEW_TEETH"
'    sSQL = sSQL & " WHERE PATIENT_UNIT_ID = " & lPatientID
'
'    Set objRS = objDas.OpenRecordset(sSQL)
'    If objRS.RecordCount <> moPositions.Count And objRS.RecordCount <> 0 Then
'        'Somehow we don't have all 32 teeth, so delete it and start over 'NO NO NO
''        sSQL = "DELETE FROM DENTAL_ENC_VIEW_TEETH"
''        sSQL = sSQL & " WHERE PATIENT_UNIT_ID = " & lPatientID
''        objDas.ExecuteSQL (sSQL)
''
''        ReadTeethPathFromDB = ""
''        mbNewTeethPath = True
''        Exit Function
'        MsgBox Abs(CLng(moPositions.Count) - CLng(objRS.RecordCount)) & " tooth (teeth) position(s) were not found in the database. Please reset the tooth state information (primary/secondary) on the grey tooth (teeth) in full mouth.", vbOKOnly + vbInformation
'    ElseIf objRS.RecordCount = 0 Then
'
'        ReadTeethPathFromDB = ""
'        mbNewTeethPath = True
'        Exit Function
'    End If
'
'    Do While Not objRS.EOF
'        With moPositions(CStr(objRS.Fields("TOOTH_ID"))) 'This is really the position number, NOT the tooth id
'            'Set seen before Status when loading from the db becuase the db doesn't reflect the real state of the data
'            .Seen = CBool(objRS.Fields("TOOTH_SEEN"))
'            .SeenPre = .Seen
'            .Status = CInt(objRS.Fields("TOOTH_STATUS"))
'            .StatusPre = .Status
'            'If the tooth status is not 0, then it has erupted
'            If CInt(objRS.Fields("TOOTH_STATUS")) <> 0 Then
'                .ToothErupted = True
'            End If
'        End With
'
'        objRS.MoveNext
'    Loop
'
'    ReadTeethPathFromDB = moPositions.TeethPath
'
'    mbNewTeethPath = False
'    Exit Function
'
'ErrorHandler:
'    Call gobjShared.ShowVBError(Err.Number, Err.Description, "clsTeeth.readTeethPathFromDB", "srvSVGComm", vbCritical)
'    Exit Function
'    Resume
'End Function
'
'Private Function readTeethPathFromRC(nAge As Integer) As String
'
'    'SF 4-09-04 This function is no longer needed, see LoadPositions
'
'    Dim i As Integer
'
'    On Error GoTo ErrorHandler
'
'    'SF 4-09-04 commented out because i set it in class_initialize
'    'Set mObjSVGMngr.Positions = Nothing
'
'    'SF This was commented out a long time ago
'    'mObjSVGMngr.teethAgePath = ""
'
'    'SF 4-09-04 commented out because this is now set when the user sets the patientid in svgcomm
'    'mObjSVGMngr.Age = nAge
'
'    'readTeethPathFromRC = mObjSVGMngr.GetTeethPathFromRc(nAge)
'
'    Exit Function
'
'ErrorHandler:
'    Call gobjShared.ShowVBError(Err.Number, Err.Description, "clsTeeth.readTeethPathFromRC", "srvSVGComm", vbCritical)
'
'End Function


Public Function LoadTeeth(ViewMode As eViewMode)

    Call mObjSVGMngr.LoadTeeth(moPositions, ViewMode)
    Exit Function
    
ErrorHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "clsTeeth.loadTeethFromRC", "srvSVGComm", vbCritical)
        
End Function

'Private Function selectedToothOnly(ByVal nTooth As Integer, nSelect As Integer)
'    Dim i As Integer
'    For i = 1 To TOTAL_TEETH
'        If i = nTooth Then
'            aTeethSelected(i) = nSelect
'        Else
'            aTeethSelected(i) = NOT_SEL
'        End If
'    Next
'    If nSelect = SEL Then
'        RaiseEvent SelectionCountChange(True)
'    End If
'End Function
'
'Private Function selectedTooth(ByVal nTooth As Integer, nSelect As Integer)
'   aTeethSelected(nTooth) = nSelect
'   If selectedCount > 1 Then
'        RaiseEvent SelectionCountChange(False)
'   Else
'        RaiseEvent SelectionCountChange(True)
'   End If
'End Function

Public Function selectedCount() As Integer
    Dim i, nCount As Integer
    
    'Is this used?
    Debug.Assert False
    
    For i = 1 To moPositions.Count
        If moPositions(i).Selected = True Then
            nCount = nCount + 1
        End If
    Next
    
    selectedCount = nCount
    
End Function

Public Sub InitializeManager()
   mObjSVGMngr.Initialize
End Sub

Public Property Let DxTxViewMode(ViewMode As eViewMode)

'This replaced SetDxTxMode
    meDxTxViewMode = ViewMode
    
    If ViewMode = eDxView Then
        Set m_gTeethDoc = m_TeethDocDx
        Set m_gTeethWin = m_TeethWinDx
    ElseIf ViewMode = eTxView Then
        Set m_gTeethDoc = m_TeethDocTx
        Set m_gTeethWin = m_TeethWinTx
    ElseIf ViewMode = eDx603aView Then
        Set m_gTeethDoc = m_TeethDocDx603a
        Set m_gTeethWin = m_TeethWinDx603a
    End If
    
    If Not m_ToothObject Is Nothing Then
        m_ToothObject.DxTxViewMode = ViewMode
        
        'This is the conversion (link) between ViewType and ViewMode
        'm_ToothObject.DxTxMode = ViewType
        
'        If bIsDx Then
'            m_ToothObject.DxTxMode = eDxView
'        Else
'            m_ToothObject.DxTxMode = eTxView
'        End If
    Else
        'This should only happend once when GC loads
        'Debug.Assert False
    End If
    
    setTeethView

End Property

Private Sub Class_Initialize()

    'Set m_SVGLIB = New CHCSII_Dental_SVG_Lib.SVG_Lib
    Dim oSingleton As New CHCSII_Dental_SVG_Manager.SF
    Set mColSelectedTeeth = New colSelectedTeeth
    Set mObjSVGMngr = oSingleton.Singleton  'Set SVG Manager to one instance
    Set mObjSVGDraw = mObjSVGMngr.mObjSVGDraw
    
    'SF 4-9-04
    Set moPositions = New colPositions
    Set mObjSVGMngr.Positions = moPositions
    
'    mObjSVGMngr.Initialize

End Sub


Sub Class_Terminate()

    Set mColSelectedTeeth = Nothing
    
    Set m_gTeethDoc = Nothing
    Set m_gTeethWin = Nothing
    
    Set m_TeethDocDx = Nothing
    Set m_TeethWinDx = Nothing
    
    Set m_TeethDocTx = Nothing
    Set m_TeethWinTx = Nothing
    
    Set mObjSVGMngr = Nothing
    Set mObjSVGDraw = Nothing

    Set m_ToothObject = Nothing

End Sub

Private Sub mObjSVGDraw_MissingTooth(toothID As Integer)
'    Dim n As Integer
'    n = ToothID
'    If ToothID > 32 Then
'        n = convNumToPos(ToothID)
'    End If
'
'    aTeethStatus(n) = NONE
'    m_sAgePath = getPathTeeth()
End Sub
Public Function getSelectedTeeth() As String
    Dim i, nNum As Integer
    Dim s As String
    Dim GCTooth As clsGCTooth
    
    For Each GCTooth In mColSelectedTeeth
        
                s = s + "," + CStr(GCTooth.toothID)
                
    Next
    
    If Not s = "" Then
        getSelectedTeeth = Right(s, Len(s) - 1)
    End If
End Function


Private Sub setTeethView()

    'This sub can use work
    'Performance issue
    
    Dim i, nTooth As Integer
    Dim oRect As Object
    
    If m_gTeethDoc Is Nothing Then
        'There is nothing to do
        Exit Sub
    End If
    
    'Set all teeth to white_hide
    For i = 1 To moPositions.Count
        nTooth = moPositions(i).ToothNum
        
        If Not m_gTeethDoc Is Nothing Then
            Set oRect = m_gTeethDoc.getElementById(CStr(nTooth))
            If Not oRect Is Nothing Then
                If oRect.getAttribute("class") = YELLOW Or _
                    oRect.getAttribute("class") = YELLOW_HIDE Then
                    Call oRect.setAttribute("class", WHITE_HIDE)
                    'SF SCR 47877 Added doevents
                    DoEvents
                End If
            End If
        End If
    Next
    
    'Set the teeth as selected (yellow) or white
    For i = 1 To moPositions.Count
        nTooth = moPositions(i).ToothNum
        
        If moPositions(i).Selected = True Then
            If Not m_gTeethDoc Is Nothing Then
                Set oRect = m_gTeethDoc.getElementById(CStr(nTooth))
                If Not oRect Is Nothing Then
                    Call oRect.setAttribute("class", YELLOW)
                End If
            End If
        End If
        
        If moPositions(i).Seen = True Then
            If Not m_gTeethDoc Is Nothing Then
                Set oRect = m_gTeethDoc.getElementById(CStr(nTooth))
                If Not oRect Is Nothing Then
                    If oRect.getAttribute("class") = GREY_HIDE Then
                        Call oRect.setAttribute("class", WHITE_HIDE)
                        'SF SCR 47877 Added doevents
                        DoEvents
                    End If
                End If
            End If
        End If
    Next
   
End Sub
Private Sub mObjSVGMngr_ToothLoaded(iToothID As Integer, strToothSVG As String)
            
    Dim oSvg, oSSvg, oRect As Object
    Dim strArrAgePath() As String
    Dim oPos As clsPosition
    
   ' DoEvents
    'Exit Sub
    'This should only happen the first time GC loads
    If m_gTeethDoc Is Nothing Then
        Set m_gTeethDoc = m_TeethDocDx
        Set m_gTeethWin = m_TeethWinDx
    End If
    
    Set oSvg = m_gTeethDoc.getElementById("svg0")
    Set oSSvg = m_gTeethDoc.getElementById("svg" & CStr(iToothID))
    If Not oSSvg Is Nothing Then
        oSvg.removeChild m_gTeethDoc.getElementById("svg" & CStr(iToothID))
    End If
    
    oSvg.appendChild m_gTeethWin.parseXML(strToothSVG, m_gTeethDoc)
    
    Set oRect = m_gTeethDoc.getElementById(CStr(iToothID))
    
    Set oPos = moPositions.GetPositionByToothNum(iToothID)
    
    If oPos.Seen = False Then
        Call oRect.setAttribute("class", GREY_HIDE)
    ElseIf oPos.Seen = True Then
        Call oRect.setAttribute("class", WHITE_HIDE)
    End If

    If oPos.Selected = True Then
        Call oRect.setAttribute("class", YELLOW)
    End If
              
    'If we have the last tooth then raise the event that we are finished
    If moPositions(moPositions.Count).ToothNum = iToothID Then
       'Debug.Assert False
       RaiseEvent DrawingFinished
    End If

'    strArrAgePath = Split(m_sAgePath, ",")
'    If CInt(strArrAgePath(UBound(strArrAgePath))) = iToothID Then
'       RaiseEvent DrawingFinished
'    End If
    
    'DoEvents
    
End Sub

Public Function ResetSelection()

    Dim oSvg, oSSvg, oRect, oToothDoc As Object
    Dim x, y, nLastViewTooth As Integer
    Dim s As String
    
    On Error GoTo ErrorHandler
    
    If m_ToothObject Is Nothing Or m_gTeethDoc Is Nothing Then
        Exit Function
    End If
    
    m_ToothObject.resetSelectedSurfaces
    ' set the background to white
    Set oRect = m_gTeethDoc.getElementById(CStr(miCurrentTooth))
    If Not oRect Is Nothing Then
        Call oRect.setAttribute("class", WHITE_HIDE)
    End If

    Exit Function
    
ErrorHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "clsTeeth.ResetSelection", "srvSVGComm", vbCritical)

End Function

Public Function clearSelectedTeeth()

    Set mColSelectedTeeth.m_gTeethDoc = m_gTeethDoc
    Set mColSelectedTeeth.m_gTeethWin = m_gTeethWin
 
    mColSelectedTeeth.removeAll
    RaiseEvent SelectionCountChange(False)
'    m_ToothObject.CurrentViewTooth = 0
    
End Function
