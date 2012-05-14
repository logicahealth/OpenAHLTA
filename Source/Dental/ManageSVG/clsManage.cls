VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public mObjMenuLib As CHCSII_Dntl_Mnu_Lib.Mnu_Lib

Public mObjSVGDraw As clsSVGDraw
Public WithEvents mObjSVGQueue As clsQueue
Attribute mObjSVGQueue.VB_VarHelpID = -1
Public mObjSVGLib As CHCSII_Dental_SVG_Lib.SVG_Lib
Public mColTeeth As Collection

'DxView tooth
Public mColDxViewToothDxs As colGCDxs
Public mColDxViewToothTxs As colGCTxs

'TxView tooth
Public mColTxViewToothDxs As colGCDxs
Public mColTxViewToothTxs As colGCTxs

'Dx603aView tooth
Public mColDx603aViewToothDxs As colGCDxs
Public mColDx603aViewToothTxs As colGCTxs

'DxView teeth
Public mColDxViewTeethDxs As Collection
Public mColDxViewTeethTxs As Collection

'TxView teeth
Public mColTxViewTeethDxs As Collection
Public mColTxViewTeethTxs As Collection

'Dx603aView teeth
Public mColDx603aViewTeethDxs As Collection
Public mColDx603aViewTeethTxs As Collection


Public Event ToothLoaded(iToothID As Integer, strToothSVG As String)

'SCR 28296 RWalker 10/10/02
Public goARM As ARM.cARMWrapper

Private miAge As Integer
Private mlPatientID As Long
Private mbPatientChanged As Boolean

Private moPositions As colPositions
Private mDxViewMouth As String
Private mTxViewMouth As String
Private mDxViewTooth As String
Private mTxViewTooth As String
Private mOSF As New CHCSII_Dntl_Mnu_Lib.SF



'------------------------------------------------------------------------------------------------------------------
'--- PROCEDURE:ARM_InitAllTransactionsForDental_ManageSVG
'---
'--- PURPOSE:   Called from Class_Initialize. Replacing all
'---            occurrences of 'MsgBox' with 'MsgBxARMd'
'---
'--- HISTORY:   SCR 28296 RWalker 10/10/02 COMMON TO ALL ARM-SCRs
'------------------------------------------------------------------------------------------------------------------
Public Function ARM_InitAllTransactionsForDental_ManageSVG() As Boolean
    Static bDoneThis    As Boolean
    Dim objARMFactory   As New ARM.cARMFactory
    
    On Error GoTo EH
    
    '--- Init Tx classes only once
    '--- Init routine should be called from first top level creatable coclass
    '--- It can be called multiple times, but will actually run only once.
    '--- Instrument all creatable classes in case dll's multiple entry
    '--- points are actually used
    If bDoneThis Then
        Exit Function
    Else
        bDoneThis = True
    End If

    Set goARM = objARMFactory.CreateARMApp("")  'BMStop
    
    On Error Resume Next
    
    '-------------------------------------------------------------------------------------------------------------------------
    '--- MESSAGE MAP REGISTERING TRANSACTIONS BY NAME W/ ARM SERVER
    '-------------------------------------------------------------------------------------------------------------------------
    
    ARM_InitAllTransactionsForDental_ManageSVG = True
    Exit Function

EH:
    ARM_InitAllTransactionsForDental_ManageSVG = False
End Function

Public Property Get ofrmMenu() As Object
    
    Set ofrmMenu = frmMenu
    
End Property
Public Property Get PatientID() As Long
    PatientID = mlPatientID
End Property
Public Property Let PatientID(lPatientID As Long)
    If lPatientID <> mlPatientID Then
        mlPatientID = lPatientID
        mbPatientChanged = True
    Else
        mbPatientChanged = False
    End If
End Property

Public Property Get Positions() As colPositions
    Set Positions = moPositions
End Property

Public Property Set Positions(vData As colPositions)
        
    'Debug.Assert False
    Set moPositions = vData
    
End Property

Public Function GetTeethPathFromRc(age As Integer) As String

Dim strArrTeethPath() As String
Dim strTmpTeethPath As String
Dim i As Integer
Dim j As Integer

  If age < 18 Then
         GetTeethPathFromRc = LoadResString(age + 101)
  Else
         GetTeethPathFromRc = LoadResString(118)
  End If

  strArrTeethPath = Split(GetTeethPathFromRc, ",")

  'Put the teeth into the correct order based on the tooth positions
  For i = 1 To 32
    For j = 0 To UBound(strArrTeethPath)
       If CInt(strArrTeethPath(j)) < 33 Then
            'Secondary tooth
            If CInt(strArrTeethPath(j)) = i Then
                strTmpTeethPath = strTmpTeethPath & "," & strArrTeethPath(j)
            End If
       ElseIf CInt(strArrTeethPath(j)) < 43 Then
            'Primary tooth in upper quad
            If CInt(strArrTeethPath(j)) - 29 = i Then
                strTmpTeethPath = strTmpTeethPath & "," & strArrTeethPath(j)
            End If
       Else
            If CInt(strArrTeethPath(j)) - 23 = i Then
                strTmpTeethPath = strTmpTeethPath & "," & strArrTeethPath(j)
            End If
       End If
    Next j
  Next i

  'Remove the first ","
  GetTeethPathFromRc = Right(strTmpTeethPath, Len(strTmpTeethPath) - 1)

End Function
Public Property Let age(RHV As Integer)
    If RHV = miAge Then
      Exit Property
    Else
      miAge = RHV
    End If
End Property
Public Sub Initialize()
    'sf don't do this on init
'    TerminateObjects
    InitializeObjects
    initializeTeeth
    
    '--- SCR 28296 RWalker 10/10/02 Instantiate ARM / Init transactions
    Call ARM_InitAllTransactionsForDental_ManageSVG
End Sub
Private Sub initializeTeeth()

Dim oPos As clsPosition

If moPositions Is Nothing Then
    'Debug.Assert False
    'moPositions should be set by now? don't think so.
Else
    'Init all the teeth
    Debug.Assert False
    For Each oPos In moPositions
        mColTeeth.Add InitializeTooth(oPos), Str(oPos.ToothNum)
    Next
End If


'Dim strArrToothPath() As String
'Dim i As Integer
'
'    If miTeethAgePath = "" Then miTeethAgePath = getTeethPath(miAge)  'If no teeth path grab default
'
'    strArrToothPath = Split(miTeethAgePath, ",")
'    For i = 0 To UBound(strArrToothPath)
'       mColTeeth.Add InitializeTooth(CInt(strArrToothPath(i))), strArrToothPath(i)
'    Next
'
'    'initializeTooth

End Sub
Private Function InitializeTooth(Position As clsPosition) As clsGCTooth

'This creates a new tooth and sets up the svg tooth and surface definitions for it

Dim i As Integer
Dim oGCTooth As clsGCTooth
Dim objSurface As CHCSII_Dental_SVG_Lib.clsSVGSurface
Dim iToothNum As Integer

On Error GoTo ErrorHandler

    iToothNum = Position.ToothNum
    
    Set oGCTooth = New clsGCTooth
    
    oGCTooth.ToothID = iToothNum
    
    'SF The magic used to select a tooth that is not in the patients mouth
    oGCTooth.IsToothPlaceHolder = Not Position.ToothErupted
    
    'SF We are loading positions and all positions are available all the time.
    oGCTooth.ToothVisible = True
    
    Set oGCTooth.SVGToothDef = mObjSVGLib.SVGTeethDef(CStr(iToothNum))
    
    For i = 1 To 32
        'Get each svgsurface for the tooth
        Set objSurface = mObjSVGLib.SVGSurfaces(CStr(iToothNum) & "|" & CStr(i))
        If Not objSurface Is Nothing Then
            'Add each svgsurface to the GCTooth object
            oGCTooth.SVGSurfaces.Add objSurface, CStr(objSurface.ToothID) & "|" & CStr(objSurface.surfaceID)
        End If
        Set objSurface = Nothing
    Next
    
    Set InitializeTooth = oGCTooth
    
    Set oGCTooth = Nothing
    
    Exit Function
    
ErrorHandler:
    'Ignore Invalid procedure call or argument
    If Err.Number <> 5 Then
        MsgBxARMd Err.Number & "," & Err.Description & ",clsManage.initializeTooth, SVGManager"
    End If
    Resume Next

End Function
Public Property Get age() As Integer
   age = miAge
End Property
Private Function GetSVGDxViewBySurface(iSurface As Integer, GCTooth As clsGCTooth) As String

'Get the svg string for the DX View the the given tooth and the given surface

Dim oGCDxs As colGCDxs
Dim oGCTxs As colGCTxs
Dim oGCDx As clsGCDx
Dim oGCTx As clsGCTx
    
    'Get the svg string for the txs on the given tooth and the given surface (for DxView)
    Set oGCTxs = mColDxViewTeethTxs(CStr(GCTooth.ToothID))
    
    For Each oGCTx In oGCTxs
        If oGCTx.DrawThisInDxViewTx = True Then
            If oGCTx.Surface <> "" Then
              If oGCTx.Surface = CStr(iSurface) Then
                     GetSVGDxViewBySurface = GetSVGDxViewBySurface & " " & oGCTx.SVGTx
              End If
            End If
        End If
    Next
    
    'Get the svg string for the dxs on the given tooth and the given surface (for DxView)
    Set oGCDxs = mColDxViewTeethDxs(CStr(GCTooth.ToothID))
    
    For Each oGCDx In oGCDxs
        If oGCDx.DrawThisInDxViewDx Then
            If Not oGCDx.Surface = "" Then
                 If oGCDx.Surface = CStr(iSurface) Then
                       GetSVGDxViewBySurface = GetSVGDxViewBySurface & " " & oGCDx.SVGDx
                 End If
            End If
        End If
    Next


End Function
Private Function GetSVGTxViewBySurface(iSurface As Integer, GCTooth As clsGCTooth) As String

'Get the svg string for the Tx View for the given tooth and the given surface

Dim oGCDxs As colGCDxs
Dim oGCTxs As colGCTxs
Dim oGCDx As clsGCDx
Dim oGCTx As clsGCTx

    'Get the svg string for the txs on the given tooth and the given surface (Tx View)
    Set oGCDxs = mColTxViewTeethDxs(CStr(GCTooth.ToothID))
    
    For Each oGCDx In oGCDxs
        If oGCDx.DrawThisInTxViewDx = True Then
            If oGCDx.Surface <> "" Then
                    If oGCDx.Surface = CStr(iSurface) Then
                         GetSVGTxViewBySurface = GetSVGTxViewBySurface & " " & oGCDx.SVGDx
                   End If
            End If
        End If
    Next
    
    'Get the svg string for the dxs on the given tooth and the given surface (Tx View)
    Set oGCTxs = mColTxViewTeethTxs(CStr(GCTooth.ToothID))
    
    For Each oGCTx In oGCTxs
        If oGCTx.DrawThisInTxViewTx = True Then
            If oGCTx.Surface <> "" Then
              If oGCTx.Surface = CStr(iSurface) Then
                     GetSVGTxViewBySurface = GetSVGTxViewBySurface & " " & oGCTx.SVGTx
              End If
            End If
        End If
    Next
    
End Function

Private Function GetSVGDx603aViewBySurface(iSurface As Integer, GCTooth As clsGCTooth) As String

'Get the svg string for the Tx View for the given tooth and the given surface

Dim oGCDxs As colGCDxs
Dim oGCTxs As colGCTxs
Dim oGCDx As clsGCDx
Dim oGCTx As clsGCTx

    'Get the svg string for the txs on the given tooth and the given surface (Tx View)
    Set oGCDxs = mColDx603aViewTeethDxs(CStr(GCTooth.ToothID))
    
    For Each oGCDx In oGCDxs
        If oGCDx.DrawThisInDx603aViewDx = True Then
            If oGCDx.Surface <> "" Then
                    If oGCDx.Surface = CStr(iSurface) Then
                         GetSVGDx603aViewBySurface = GetSVGDx603aViewBySurface & " " & oGCDx.SVGDx
                   End If
            End If
        End If
    Next
    
    'Get the svg string for the dxs on the given tooth and the given surface (Tx View)
    Set oGCTxs = mColDx603aViewTeethTxs(CStr(GCTooth.ToothID))
    
    For Each oGCTx In oGCTxs
        If oGCTx.DrawThisInDx603aViewTx = True Then
            If oGCTx.Surface <> "" Then
              If oGCTx.Surface = CStr(iSurface) Then
                     GetSVGDx603aViewBySurface = GetSVGDx603aViewBySurface & " " & oGCTx.SVGTx
              End If
            End If
        End If
    Next
    
End Function

Private Function getDxTooth(GCTooth As clsGCTooth) As String

'Get the svg string for Dx View for the given tooth for whole tooth dxs and txs (not surface specific)

Dim oGCDxs As colGCDxs
Dim oGCTxs As colGCTxs
Dim oGCDx As clsGCDx
Dim oGCTx As clsGCTx

    'getDxSurface = mObjSVGDraw.hatchSurface(GCTooth.SVGToothDef.iToothID, iSurface, eRed, eHorizontal)
        
    'Get the svg string for the txs for the given tooth for whole tooth txs
    Set oGCTxs = mColDxViewTeethTxs(CStr(GCTooth.ToothID))
    For Each oGCTx In oGCTxs
        If oGCTx.DrawThisInDxViewTx Then
            If oGCTx.Surface <> "" Then
                If oGCTx.Surface = "777" Then
                    getDxTooth = getDxTooth & " " & oGCTx.SVGTx
                End If
            End If
        End If
    Next
    
    'Get the svg string for the txs for the given tooth for whole tooth txs
    Set oGCDxs = mColDxViewTeethDxs(CStr(GCTooth.ToothID))
    For Each oGCDx In oGCDxs
        If oGCDx.DrawThisInDxViewDx Then
            If oGCDx.Surface <> "" Then
                If oGCDx.Surface = "777" Then
                    getDxTooth = getDxTooth & " " & oGCDx.SVGDx
                End If
            End If
        End If
    Next

End Function
Private Function GetTxTooth(GCTooth As clsGCTooth) As String

'Get the svg string for the Tx View for the given tooth for whole tooth dxs and txs

Dim oGCDxs As colGCDxs
Dim oGCTxs As colGCTxs
Dim oGCDx As clsGCDx
Dim oGCTx As clsGCTx

    'Get the svg string for the Tx View for the given tooth for whole tooth Dxs
    Set oGCDxs = mColTxViewTeethDxs(CStr(GCTooth.ToothID))
    For Each oGCDx In oGCDxs
        If oGCDx.DrawThisInTxViewDx Then
            If oGCDx.Surface <> "" Then
                If oGCDx.Surface = "777" Then
                    GetTxTooth = GetTxTooth & " " & oGCDx.SVGDx
                End If
            End If
        End If
    Next
    
    'Get the svg string for the Tx View for the given tooth for whole tooth Txs
    Set oGCTxs = mColTxViewTeethTxs(CStr(GCTooth.ToothID))
    For Each oGCTx In oGCTxs
        If oGCTx.DrawThisInTxViewTx Then
            If oGCTx.Surface <> "" Then
                If oGCTx.Surface = "777" Then
                    GetTxTooth = GetTxTooth & " " & oGCTx.SVGTx
                End If
            End If
        End If
    Next

End Function
Private Function GetDx603aTooth(GCTooth As clsGCTooth) As String

'Get the svg string for the Tx View for the given tooth for whole tooth dxs and txs

Dim oGCDxs As colGCDxs
Dim oGCTxs As colGCTxs
Dim oGCDx As clsGCDx
Dim oGCTx As clsGCTx

    'Get the svg string for the Tx View for the given tooth for whole tooth Dxs
    Set oGCDxs = mColDx603aViewTeethDxs(CStr(GCTooth.ToothID))
    For Each oGCDx In oGCDxs
        If oGCDx.DrawThisInDx603aViewDx Then
            If oGCDx.Surface <> "" Then
                If oGCDx.Surface = "777" Then
                    GetDx603aTooth = GetDx603aTooth & " " & oGCDx.SVGDx
                End If
            End If
        End If
    Next
    
    'Get the svg string for the Tx View for the given tooth for whole tooth Txs
    Set oGCTxs = mColDx603aViewTeethTxs(CStr(GCTooth.ToothID))
    For Each oGCTx In oGCTxs
        If oGCTx.DrawThisInDx603aViewTx Then
            If oGCTx.Surface <> "" Then
                If oGCTx.Surface = "777" Then
                    GetDx603aTooth = GetDx603aTooth & " " & oGCTx.SVGTx
                End If
            End If
        End If
    Next

End Function

Private Function getDefSurfaceHeader(iSurface As Integer, GCTooth As clsGCTooth) As String

Dim objSurface As CHCSII_Dental_SVG_Lib.clsSVGSurface
    
    Set objSurface = GCTooth.SVGSurfaces(CStr(GCTooth.SVGToothDef.ToothID) & "|" & CStr(iSurface))
    
    getDefSurfaceHeader = getDefSurfaceHeader & " " & objSurface.getSurfaceDefHeader
   
End Function
Private Function getDefSurfaceBody(iSurface As Integer, GCTooth As clsGCTooth) As String

Dim objSurface As CHCSII_Dental_SVG_Lib.clsSVGSurface
    
    Set objSurface = GCTooth.SVGSurfaces(CStr(GCTooth.SVGToothDef.ToothID) & "|" & CStr(iSurface))
    
    getDefSurfaceBody = getDefSurfaceBody & " " & objSurface.getSurfaceDefBody
   
End Function
Private Function GetDefASurfaceBody(iSurface As Integer, GCTooth As clsGCTooth) As String

Dim objSurface As CHCSII_Dental_SVG_Lib.clsSVGSurface
    
    Set objSurface = GCTooth.SVGSurfaces(CStr(GCTooth.SVGToothDef.ToothID) & "|" & CStr(iSurface))
    
    GetDefASurfaceBody = GetDefASurfaceBody & " " & objSurface.getASurfaceDefBody
   
End Function
Private Function getDefSurfaceFooter(iSurface As Integer, GCTooth As clsGCTooth) As String

Dim objSurface As CHCSII_Dental_SVG_Lib.clsSVGSurface
    
    Set objSurface = GCTooth.SVGSurfaces(CStr(GCTooth.SVGToothDef.ToothID) & "|" & CStr(iSurface))
    
    getDefSurfaceFooter = getDefSurfaceFooter & " " & objSurface.getSurfaceDefFooter
   
End Function
Private Function GetSVGDxSurface(GCTooth As clsGCTooth, Optional RemovedLocations As eLocation, Optional ToothMode As eToothMode = eBigTooth) As String

Dim objSurface As CHCSII_Dental_SVG_Lib.clsSVGSurface
Dim RemovedSurfaces As String
Dim bRemoveSurface As Boolean
Dim i As Integer
Dim strArrRemovedSurfaces() As String
Dim oSurfaceGroup As clsSurfaceGroup
'Dim objMacro As clsMacro

    Set oSurfaceGroup = mObjSVGLib.SurfaceGroups(GCTooth.ToothID)
    
    If RemovedLocations = eCoronalSection Then
        RemovedSurfaces = oSurfaceGroup.CrownSurfaces
    ElseIf RemovedLocations = eRootSection Then
        RemovedSurfaces = oSurfaceGroup.RootSurfaces
    End If
    
    strArrRemovedSurfaces = Split(RemovedSurfaces, ",")
    
    'lets do this only for specialized modes such as crowns (111) if more need to be added we will do this for each case
    GetSVGDxSurface = GetSVGDxSurface & " " & GetSVGDxViewBySurface(111, GCTooth)

    'Iterate through each surface assigned to the tooth
    For Each objSurface In GCTooth.SVGSurfaces
        'Iterate through all the surfaces removed from the tooth
        For i = 0 To UBound(strArrRemovedSurfaces)
            If objSurface.surfaceID = CInt(strArrRemovedSurfaces(i)) Then
            'if I have a match, then don't get the svg for that surface, and skip to the next surface on the tooth
                bRemoveSurface = True
                Exit For
            End If
        Next i
        
        If bRemoveSurface = False Then
            'I need to display the surface, so get the svg for it.
            If ToothMode = eBigTooth Then
                'ToDo: room for improvement
                GetSVGDxSurface = GetSVGDxSurface & " " & getDefSurfaceHeader(objSurface.surfaceID, GCTooth) & _
                             " " & GetSVGDxViewBySurface(objSurface.surfaceID, GCTooth) & _
                             " " & getDefSurfaceBody(objSurface.surfaceID, GCTooth) & _
                             " " & GetDefASurfaceBody(objSurface.surfaceID, GCTooth) & _
                             " " & getDefSurfaceFooter(objSurface.surfaceID, GCTooth)
            ElseIf ToothMode = eFullMouth Then
                'ToDo: room for improvement
                GetSVGDxSurface = GetSVGDxSurface & " " & getDefSurfaceHeader(objSurface.surfaceID, GCTooth) & _
                             " " & GetSVGDxViewBySurface(objSurface.surfaceID, GCTooth) & _
                             " " & GetDefASurfaceBody(objSurface.surfaceID, GCTooth) & _
                             " " & getDefSurfaceFooter(objSurface.surfaceID, GCTooth)
            End If
        End If
        
        bRemoveSurface = False
    Next
    
 
End Function
Private Function GetSVGTxSurface(GCTooth As clsGCTooth, Optional RemovedLocations As eLocation, Optional ToothMode As eToothMode = eBigTooth) As String

Dim objSurface As CHCSII_Dental_SVG_Lib.clsSVGSurface
Dim RemovedSurfaces As String
Dim bRemoveSurface As Boolean
Dim i As Integer
Dim strArrRemovedSurfaces() As String
Dim oSurfaceGroup As clsSurfaceGroup
'Dim objMacro As clsMacro

    If ToothMode = -1 Then
        MsgBox "Toothmode cannot = -1"
        Exit Function
    End If

    Set oSurfaceGroup = mObjSVGLib.SurfaceGroups(GCTooth.ToothID)
    
    If RemovedLocations = eCoronalSection Then
        RemovedSurfaces = oSurfaceGroup.CrownSurfaces
    ElseIf RemovedLocations = eRootSection Then
        RemovedSurfaces = oSurfaceGroup.RootSurfaces
    End If

    
    strArrRemovedSurfaces = Split(RemovedSurfaces, ",")
    'lets do this only for specialized modes such as crowns (111) if more need to be added we will do this for each case
    GetSVGTxSurface = GetSVGTxSurface & " " & GetSVGTxViewBySurface(111, GCTooth)

    'Iterate through each surface assigned to the tooth
    For Each objSurface In GCTooth.SVGSurfaces
        'Iterate through all the surfaces removed from the tooth
        For i = 0 To UBound(strArrRemovedSurfaces)
            If objSurface.surfaceID = CInt(strArrRemovedSurfaces(i)) Then
            'if I have a match, then don't get the svg for that surface, and skip to the next surface on the tooth
                bRemoveSurface = True
                Exit For
            End If
        Next i
      
        If bRemoveSurface = False Then
            'I need to display the surface, so get the svg for it.
            If ToothMode = eBigTooth Then
                'ToDo: room for improvement
                GetSVGTxSurface = GetSVGTxSurface & " " & getDefSurfaceHeader(objSurface.surfaceID, GCTooth) & _
                              " " & GetSVGTxViewBySurface(objSurface.surfaceID, GCTooth) & _
                              " " & getDefSurfaceBody(objSurface.surfaceID, GCTooth) & _
                              " " & GetDefASurfaceBody(objSurface.surfaceID, GCTooth) & _
                              " " & getDefSurfaceFooter(objSurface.surfaceID, GCTooth)
            ElseIf ToothMode = eFullMouth Then
                GetSVGTxSurface = GetSVGTxSurface & " " & getDefSurfaceHeader(objSurface.surfaceID, GCTooth) & _
                              " " & GetSVGTxViewBySurface(objSurface.surfaceID, GCTooth) & _
                              " " & GetDefASurfaceBody(objSurface.surfaceID, GCTooth) & _
                              " " & getDefSurfaceFooter(objSurface.surfaceID, GCTooth)
            End If
        End If
        'SF 03-06-03 SCR 33154 -- Reset the var
        bRemoveSurface = False
    Next
End Function

Private Function GetSVGDx603aSurface(GCTooth As clsGCTooth, Optional RemovedLocations As eLocation, Optional ToothMode As eToothMode = eBigTooth) As String

Dim objSurface As CHCSII_Dental_SVG_Lib.clsSVGSurface
Dim RemovedSurfaces As String
Dim bRemoveSurface As Boolean
Dim i As Integer
Dim strArrRemovedSurfaces() As String
Dim oSurfaceGroup As clsSurfaceGroup
'Dim objMacro As clsMacro

    If ToothMode = -1 Then
        MsgBox "Toothmode cannot = -1"
        Exit Function
    End If

    Set oSurfaceGroup = mObjSVGLib.SurfaceGroups(GCTooth.ToothID)
    
    If RemovedLocations = eCoronalSection Then
        RemovedSurfaces = oSurfaceGroup.CrownSurfaces
    ElseIf RemovedLocations = eRootSection Then
        RemovedSurfaces = oSurfaceGroup.RootSurfaces
    End If

    
    strArrRemovedSurfaces = Split(RemovedSurfaces, ",")
    'lets do this only for specialized modes such as crowns (111) if more need to be added we will do this for each case
    GetSVGDx603aSurface = GetSVGDx603aSurface & " " & GetSVGDx603aViewBySurface(111, GCTooth)

    'Iterate through each surface assigned to the tooth
    For Each objSurface In GCTooth.SVGSurfaces
        'Iterate through all the surfaces removed from the tooth
        For i = 0 To UBound(strArrRemovedSurfaces)
            If objSurface.surfaceID = CInt(strArrRemovedSurfaces(i)) Then
            'if I have a match, then don't get the svg for that surface, and skip to the next surface on the tooth
                bRemoveSurface = True
                Exit For
            End If
        Next i
      
        If bRemoveSurface = False Then
            'I need to display the surface, so get the svg for it.
            If ToothMode = eBigTooth Then
                'ToDo: room for improvement
                GetSVGDx603aSurface = GetSVGDx603aSurface & " " & getDefSurfaceHeader(objSurface.surfaceID, GCTooth) & _
                              " " & GetSVGDx603aViewBySurface(objSurface.surfaceID, GCTooth) & _
                              " " & getDefSurfaceBody(objSurface.surfaceID, GCTooth) & _
                              " " & GetDefASurfaceBody(objSurface.surfaceID, GCTooth) & _
                              " " & getDefSurfaceFooter(objSurface.surfaceID, GCTooth)
            ElseIf ToothMode = eFullMouth Then
                GetSVGDx603aSurface = GetSVGDx603aSurface & " " & getDefSurfaceHeader(objSurface.surfaceID, GCTooth) & _
                              " " & GetSVGDx603aViewBySurface(objSurface.surfaceID, GCTooth) & _
                              " " & GetDefASurfaceBody(objSurface.surfaceID, GCTooth) & _
                              " " & getDefSurfaceFooter(objSurface.surfaceID, GCTooth)
            End If
        End If
        'SF 03-06-03 SCR 33154 -- Reset the var
        bRemoveSurface = False
    Next
End Function

Public Function GetSVGDxViewTeeth(strArrTeeth() As String) As String

'This function iterates through each tooth and returns the svg string for all the teeth
Dim iTooth As Integer

        For iTooth = 0 To UBound(strArrTeeth)
            GetSVGDxViewTeeth = GetSVGDxViewTeeth & " " & GetSVGDxViewByTooth(CInt(strArrTeeth(iTooth)), False)
            
        Next iTooth
 
End Function
Public Function GetSVGDxViewByTooth(ToothID As Integer, Optional ToothMode As eToothMode = eBigTooth) As String

'This function returns the svg string for a tooth given a tooth id

Dim GCTooth As clsGCTooth
Dim strSVGDxView As String

On Error GoTo ErrorHandler

    If ToothMode = -1 Then
        MsgBox "Toothmode cannot = -1"
        Exit Function
    End If
    
'Debug.Print toothID & " Now: " & Now
    Set GCTooth = mColTeeth(CStr(ToothID))
    If GCTooth Is Nothing Then
        Set GCTooth = InitializeTooth(moPositions.GetPositionByToothNum(ToothID))
        GCTooth.ToothVisible = True
        mColTeeth.Add GCTooth, CStr(GCTooth.ToothID)
    End If
   
    If GCTooth.ToothVisible = False Then
        Exit Function
    End If
   
    If Not GCTooth.DateTimeRootRemoved = "12:00:00 AM" Then
        GCTooth.RemoveLocation = eRootSection
    End If
   
    If Not GCTooth.DateTimeCrownRemoved = "12:00:00 AM" Then
        GCTooth.RemoveLocation = eCoronalSection
    End If
   
    If Not GCTooth.DateTimeWholeToothRemoved = "12:00:00 AM" Then
        If Not GCTooth.DateTimeCrownRemoved = Empty Then
            GCTooth.RemoveLocation = eCoronalSection
        
        ElseIf Not GCTooth.DateTimeRootRemoved = Empty Then
            GCTooth.RemoveLocation = eRootSection
        
        Else
            GCTooth.RemoveLocation = eWholeTooth
        
        End If
    End If
   
    If Not GCTooth.DateTimeWholeToothPlanToRemove = "12:00:00 AM" Then
        If Not GCTooth.DateTimeImplant = "12:00:00 AM" Then
            GCTooth.RemoveLocation = eWholeTooth
        
        ElseIf Not GCTooth.DateTimeCrownRemoved = "12:00:00 AM" Then
            GCTooth.RemoveLocation = eCoronalSection
        
        ElseIf Not GCTooth.DateTimeRootRemoved = Empty Then
            GCTooth.RemoveLocation = eRootSection
        
        Else
            GCTooth.RemoveLocation = eSelectedSurfaces
        
        End If
    End If
   
    'SF Invisible tooth
    If GCTooth.IsToothPlaceHolder Then
    
        strSVGDxView = GCTooth.SVGToothDef.getToothDefHeader(ToothMode) & _
                        " " & getDxTooth(GCTooth) & _
                        " " & GCTooth.SVGToothDef.getToothDefBody & _
                        " " & GCTooth.SVGToothDef.getToothDefFooter
                        
        'Hide the tooth lines
        strSVGDxView = Replace(strSVGDxView, "class=""Outline""", "class=""FillNone""")
        
        'Make the tooth number grey
        strSVGDxView = Replace(strSVGDxView, "class=""FillBlack", "class=""GreyOnly")
        
    ElseIf GCTooth.RemoveLocation = eWholeTooth Then
        strSVGDxView = strSVGDxView & " " & GCTooth.SVGToothDef.getToothDefHeader(ToothMode) & _
                                      " " & getDxTooth(GCTooth) & _
                                      " " & GCTooth.SVGToothDef.getToothDefFooter
       'strSVGDxView = ""
    ElseIf GCTooth.RemoveLocation = eRootSection Then
        strSVGDxView = strSVGDxView & " " & GCTooth.SVGToothDef.getToothDefHeader(ToothMode) & _
                        " " & getDxTooth(GCTooth) & _
                        " " & GetSVGDxSurface(GCTooth, eRootSection, ToothMode) & _
                        " " & GCTooth.SVGToothDef.getToothCrownBody & _
                        " " & GCTooth.SVGToothDef.getToothDefFooter
    
    ElseIf GCTooth.RemoveLocation = eCoronalSection Then
        strSVGDxView = strSVGDxView & " " & GCTooth.SVGToothDef.getToothDefHeader(ToothMode) & _
                        " " & getDxTooth(GCTooth) & _
                        " " & GetSVGDxSurface(GCTooth, eCoronalSection, ToothMode) & _
                        " " & GCTooth.SVGToothDef.getToothRootBody & _
                        " " & GCTooth.SVGToothDef.getToothDefFooter
    
    Else
        strSVGDxView = strSVGDxView & " " & GCTooth.SVGToothDef.getToothDefHeader(ToothMode) & _
                        " " & getDxTooth(GCTooth) & _
                        " " & GetSVGDxSurface(GCTooth, , ToothMode) & _
                        " " & GCTooth.SVGToothDef.getToothDefBody & _
                        " " & GCTooth.SVGToothDef.getToothDefFooter
    End If
   
    GetSVGDxViewByTooth = strSVGDxView
    
Exit Function
    
ErrorHandler:
    MsgBxARMd Err.Number & "," & Err.Description & ",clsManage.GetSVGDxViewByTooth, SVGManager"
    Resume Next
    
End Function
Public Function GetSVGDx603aViewbyTooth(ToothID As Integer, Optional ToothMode As eToothMode = eBigTooth) As String

'TODO: Add 603a Logic here
'This function returns the svg string for a tooth given a tooth id
'Taken from GetSVGTxViewByTooth

Dim GCTooth As clsGCTooth
Dim strSVGTxView As String

On Error GoTo ErrorHandler

    If ToothMode = -1 Then
        MsgBox "Toothmode cannot = -1"
        Exit Function
    End If

'Debug.Print toothID & " Now: " & Now
'Debug.Assert False

    If mColTeeth Is Nothing Then
        Exit Function
    End If
    
    Set GCTooth = mColTeeth(CStr(ToothID))
    If GCTooth Is Nothing Then
        Set GCTooth = InitializeTooth(moPositions.GetPositionByToothNum(ToothID))
        'GCTooth.ToothVisible = True
        mColTeeth.Add GCTooth, CStr(GCTooth.ToothID)
    End If
    
    If GCTooth.ToothVisible = False Then
        Exit Function
    End If
    
    If Not GCTooth.DateTimeRootRemoved = "12:00:00 AM" Then
        GCTooth.RemoveLocation = eRootSection
    End If
    
    If Not GCTooth.DateTimePlannedRootRemoved = "12:00:00 AM" Then
        GCTooth.RemoveLocation = eRootSection
    End If
    
    If Not GCTooth.DateTimeCrownRemoved = "12:00:00 AM" Then
        GCTooth.RemoveLocation = eCoronalSection
    End If
   
    If Not GCTooth.DateTimeWholeToothRemoved = "12:00:00 AM" Then
        If Not GCTooth.DateTimeCrownRemoved = Empty Then
            GCTooth.RemoveLocation = eCoronalSection
        
        ElseIf Not GCTooth.DateTimeRootRemoved = Empty Then
            GCTooth.RemoveLocation = eRootSection
            
        ElseIf Not GCTooth.DateTimePlannedRootRemoved = Empty Then
            GCTooth.RemoveLocation = eRootSection
            
        Else
            GCTooth.RemoveLocation = eWholeTooth
        
        End If
    End If
   
    If Not GCTooth.DateTimeWholeToothPlanToRemove = "12:00:00 AM" Then
        If Not GCTooth.DateTimeImplant = "12:00:00 AM" Then
            GCTooth.RemoveLocation = eWholeTooth
        
        ElseIf Not GCTooth.DateTimeCrownRemoved = "12:00:00 AM" Then
            GCTooth.RemoveLocation = eCoronalSection
        
        Else
            GCTooth.RemoveLocation = eSelectedSurfaces
            
        End If
    End If
   
   'SF Invisible tooth
   If GCTooth.IsToothPlaceHolder Then

        strSVGTxView = GCTooth.SVGToothDef.getToothDefHeader(ToothMode) & _
                        " " & GetDx603aTooth(GCTooth) & _
                        " " & GCTooth.SVGToothDef.getToothDefBody & _
                        " " & GCTooth.SVGToothDef.getToothDefFooter
                        
        'Hide the tooth lines
        strSVGTxView = Replace(strSVGTxView, "class=""Outline""", "class=""FillNone""")
        
        'Make the tooth number grey
        strSVGTxView = Replace(strSVGTxView, "class=""FillBlack", "class=""GreyOnly")
        
   ElseIf GCTooth.RemoveLocation = eWholeTooth Then
        strSVGTxView = strSVGTxView & " " & GCTooth.SVGToothDef.getToothDefHeader(ToothMode) & _
                                      " " & GetDx603aTooth(GCTooth) & _
                                      " " & GCTooth.SVGToothDef.getToothDefFooter
       'strSVGDxView = ""
    ElseIf GCTooth.RemoveLocation = eRootSection Then
        strSVGTxView = strSVGTxView & " " & GCTooth.SVGToothDef.getToothDefHeader(ToothMode) & _
                        " " & GetDx603aTooth(GCTooth) & _
                        " " & GetSVGDx603aSurface(GCTooth, eRootSection, ToothMode) & _
                        " " & GCTooth.SVGToothDef.getToothCrownBody & _
                        " " & GCTooth.SVGToothDef.getToothDefFooter
    
    ElseIf GCTooth.RemoveLocation = eCoronalSection Then
        strSVGTxView = strSVGTxView & " " & GCTooth.SVGToothDef.getToothDefHeader(ToothMode) & _
                        " " & GetDx603aTooth(GCTooth) & _
                        " " & GetSVGDx603aSurface(GCTooth, eCoronalSection, ToothMode) & _
                        " " & GCTooth.SVGToothDef.getToothRootBody & _
                        " " & GCTooth.SVGToothDef.getToothDefFooter
    
    Else
        strSVGTxView = strSVGTxView & " " & GCTooth.SVGToothDef.getToothDefHeader(ToothMode) & _
                        " " & GetDx603aTooth(GCTooth) & _
                        " " & GetSVGDx603aSurface(GCTooth, , ToothMode) & _
                        " " & GCTooth.SVGToothDef.getToothDefBody & _
                        " " & GCTooth.SVGToothDef.getToothDefFooter
            
    End If

    GetSVGDx603aViewbyTooth = strSVGTxView
   
Exit Function
    
ErrorHandler:
    MsgBxARMd Err.Number & "," & Err.Description & ",clsManage.GetSVGDx603aViewbyTooth, SVGManager"
    Resume Next
    
End Function
Public Function GetSVGTxViewByTooth(ToothID As Integer, Optional ToothMode As eToothMode = eBigTooth) As String

'This function returns the svg string for a tooth given a tooth id

Dim GCTooth As clsGCTooth
Dim strSVGTxView As String

On Error GoTo ErrorHandler

    If ToothMode = -1 Then
        MsgBox "Toothmode cannot = -1"
        Exit Function
    End If

'Debug.Print toothID & " Now: " & Now
'Debug.Assert False

    If mColTeeth Is Nothing Then
        Exit Function
    End If
    
    Set GCTooth = mColTeeth(CStr(ToothID))
    If GCTooth Is Nothing Then
        Set GCTooth = InitializeTooth(moPositions.GetPositionByToothNum(ToothID))
        'GCTooth.ToothVisible = True
        mColTeeth.Add GCTooth, CStr(GCTooth.ToothID)
    End If
    
    If GCTooth.ToothVisible = False Then
        Exit Function
    End If
    
    If Not GCTooth.DateTimeRootRemoved = "12:00:00 AM" Then
        GCTooth.RemoveLocation = eRootSection
    End If
    
    If Not GCTooth.DateTimePlannedRootRemoved = "12:00:00 AM" Then
        GCTooth.RemoveLocation = eRootSection
    End If
    
    If Not GCTooth.DateTimeCrownRemoved = "12:00:00 AM" Then
        GCTooth.RemoveLocation = eCoronalSection
    End If
   
    If Not GCTooth.DateTimeWholeToothRemoved = "12:00:00 AM" Then
        If Not GCTooth.DateTimeCrownRemoved = Empty Then
            GCTooth.RemoveLocation = eCoronalSection
        
        ElseIf Not GCTooth.DateTimeRootRemoved = Empty Then
            GCTooth.RemoveLocation = eRootSection
            
        ElseIf Not GCTooth.DateTimePlannedRootRemoved = Empty Then
            GCTooth.RemoveLocation = eRootSection
            
        Else
            GCTooth.RemoveLocation = eWholeTooth
        
        End If
    End If
   
    If Not GCTooth.DateTimeWholeToothPlanToRemove = "12:00:00 AM" Then
        If Not GCTooth.DateTimeImplant = "12:00:00 AM" Then
            GCTooth.RemoveLocation = eWholeTooth
        
        ElseIf Not GCTooth.DateTimeCrownRemoved = "12:00:00 AM" Then
            GCTooth.RemoveLocation = eCoronalSection
        
        Else
            GCTooth.RemoveLocation = eSelectedSurfaces
            
        End If
    End If
   
   'SF Invisible tooth
   If GCTooth.IsToothPlaceHolder Then

        strSVGTxView = GCTooth.SVGToothDef.getToothDefHeader(ToothMode) & _
                        " " & GetTxTooth(GCTooth) & _
                        " " & GCTooth.SVGToothDef.getToothDefBody & _
                        " " & GCTooth.SVGToothDef.getToothDefFooter
                        
        'Hide the tooth lines
        strSVGTxView = Replace(strSVGTxView, "class=""Outline""", "class=""FillNone""")
        
        'Make the tooth number grey
        strSVGTxView = Replace(strSVGTxView, "class=""FillBlack", "class=""GreyOnly")
        
   ElseIf GCTooth.RemoveLocation = eWholeTooth Then
        strSVGTxView = strSVGTxView & " " & GCTooth.SVGToothDef.getToothDefHeader(ToothMode) & _
                                      " " & GetTxTooth(GCTooth) & _
                                      " " & GCTooth.SVGToothDef.getToothDefFooter
       'strSVGDxView = ""
    ElseIf GCTooth.RemoveLocation = eRootSection Then
        strSVGTxView = strSVGTxView & " " & GCTooth.SVGToothDef.getToothDefHeader(ToothMode) & _
                        " " & GetTxTooth(GCTooth) & _
                        " " & GetSVGTxSurface(GCTooth, eRootSection, ToothMode) & _
                        " " & GCTooth.SVGToothDef.getToothCrownBody & _
                        " " & GCTooth.SVGToothDef.getToothDefFooter
    
    ElseIf GCTooth.RemoveLocation = eCoronalSection Then
        strSVGTxView = strSVGTxView & " " & GCTooth.SVGToothDef.getToothDefHeader(ToothMode) & _
                        " " & GetTxTooth(GCTooth) & _
                        " " & GetSVGTxSurface(GCTooth, eCoronalSection, ToothMode) & _
                        " " & GCTooth.SVGToothDef.getToothRootBody & _
                        " " & GCTooth.SVGToothDef.getToothDefFooter
    
    Else
        strSVGTxView = strSVGTxView & " " & GCTooth.SVGToothDef.getToothDefHeader(ToothMode) & _
                        " " & GetTxTooth(GCTooth) & _
                        " " & GetSVGTxSurface(GCTooth, , ToothMode) & _
                        " " & GCTooth.SVGToothDef.getToothDefBody & _
                        " " & GCTooth.SVGToothDef.getToothDefFooter
            
    End If

    GetSVGTxViewByTooth = strSVGTxView
   
Exit Function
    
ErrorHandler:
    MsgBxARMd Err.Number & "," & Err.Description & ",clsManage.initializeTooth, SVGManager"
    Resume Next
    
End Function
Public Function IsMissing(iTooth As Integer) As Boolean

'This function determines if there is a missing dx associated with this tooth

Dim oGCDx As clsGCDx
Dim GCTooth As clsGCTooth
 
    Set GCTooth = mColTeeth(CStr(iTooth))

    For Each oGCDx In GCTooth.GCDxs
        If oGCDx.MenuID = 151 Then  'Found missing for this tooth
            IsMissing = True
            Exit Function
        End If
    Next

End Function
Public Function IsRootMissing(iTooth As Integer) As Boolean

'Determine if the tooth has a dx of missing and a RemovedLocation of RootSection

Dim oGCDx As clsGCDx
Dim GCTooth As clsGCTooth
Dim objRemovedLoc As clsRemovedLocations
 
    Set GCTooth = mColTeeth(CStr(iTooth))
    
    For Each oGCDx In GCTooth.GCDxs
        If Not oGCDx.MenuID = 151 Then  'Found missing for this tooth
            
            Set objRemovedLoc = mObjMenuLib.RemovedLocations(CStr(oGCDx.MenuID))
            If objRemovedLoc.eRemovedLocation = eRootSection Then  'Found missing for this tooth
                IsRootMissing = True
                Exit Function
            End If
        
        End If
    Next

End Function
Private Function getMenuInstruction(MenuID As String, iDxTxStatus As Integer) As clsMenuItemInstr

Dim objMenuItem As clsMenuItem

            Set objMenuItem = mObjMenuLib.MenuItems(MenuID)
            Set getMenuInstruction = objMenuItem.getMenuInstruction(CLng(iDxTxStatus))

             
End Function
Public Function IsImplant(iTooth As Integer) As Boolean

'This function determines if there is a dx that has a menu instruction symbol of Implant

Dim oGCDx As clsGCDx
Dim GCTooth As clsGCTooth
Dim objMenuInstruction As clsMenuItemInstr
 
    Set GCTooth = mColTeeth(CStr(iTooth))
    
    For Each oGCDx In GCTooth.GCDxs
        Set objMenuInstruction = getMenuInstruction(oGCDx.MenuID, DxStatusEnum.dxActiveNew)
                   
        If objMenuInstruction.symbol = eSymbol.eImplant Then 'Found missing for this tooth
                 IsImplant = True
                 Exit Function
        End If
    Next

End Function
Public Sub updateSVG(oDx As CHCSII_Dental_DXTX.clsDx, oTx As CHCSII_Dental_DXTX.clsTx)

Dim GCTooth As clsGCTooth
Dim sMenuID As String
Dim sTmpMenuID As String
Dim iToothID As Integer
Dim i As Integer
Dim strArrTeeth() As String
'Dim bMaxManSelect As Boolean
Dim j As Integer
Dim oTempDx As CHCSII_Dental_DXTX.clsDx
Dim oTempTx As CHCSII_Dental_DXTX.clsTx

Screen.MousePointer = vbHourglass

    If Not oDx Is Nothing Then
   
        'Get an array of teeth from this Dx
        If oDx.SurfaceSets(1).Teeth(1).ToothName = "Max" Or oDx.SurfaceSets(1).Teeth(1).ToothName = "Mand" Then
            strArrTeeth = Split(oDx.SurfaceSets(1).SurfaceSet, ",")
'            bMaxManSelect = True
        Else
            strArrTeeth = Split(oDx.SurfaceSets(1).Teeth(1).ToothID, ",")
'            bMaxManSelect = False
        End If
        
    Else
        
        'Get an array of teeth from this Tx
        If oTx.SurfaceSets(1).Teeth(1).ToothName = "Max" Or oTx.SurfaceSets(1).Teeth(1).ToothName = "Mand" Then
            strArrTeeth = Split(oTx.SurfaceSets(1).SurfaceSet, ",")
        Else
            strArrTeeth = Split(oTx.SurfaceSets(1).Teeth(1).ToothID, ",")
        End If
    
    End If
    
    
'     iToothID = CInt(strArrTeeth(j))
'     Set GCTooth = mColTeeth(CStr(iToothID))
    
    'Interate through the array of teeth ids
    For j = 0 To UBound(strArrTeeth)
        
        iToothID = CInt(strArrTeeth(j))
        
        'If the user switched the tooth to primary or secondary, then don't include it in the array of teeth to be updated
        If moPositions.GetPositionByToothNum(iToothID).ToothNum = iToothID Then
            Set GCTooth = mColTeeth(CStr(iToothID))
           
            Set GCTooth.GCDxs = Nothing
            Set GCTooth.GCDxs = New colGCDxs
            Set GCTooth.GCTxs = Nothing
            Set GCTooth.GCTxs = New colGCTxs
            
            'Set the Dxs for the tooth
            For Each oTempDx In mObjSVGDraw.mObjDataProc.TeethDxData(iToothID)
                If Not oTempDx Is Nothing Then
                    If oTempDx.StatusID = dxNonActiveDeleted Then
                        setTreatRetreatToToothObject oTempDx.SurfaceSets(1).SymbolID, GCTooth
                        Set oTempDx = Nothing
                        
                        If IsImplant(GCTooth.ToothID) = False Then
                            GCTooth.IsImplant = False
                            GCTooth.DateTimeImplant = Empty
                        End If
                        
                        If IsMissing(GCTooth.ToothID) = False Then
                             GCTooth.RemoveLocation = 0
                             GCTooth.IsExtractOrMissing = False
                             GCTooth.DateTimeWholeToothRemoved = Empty
                        End If
                        
                        If IsRootMissing(GCTooth.ToothID) = False Then
                             GCTooth.DateTimeRootRemoved = Empty
                        End If
                        
                        GCTooth.DateTimeCrownRemoved = Empty
                        
                        GCTooth.DateTimePlannedRootRemoved = Empty
                        GCTooth.DateTimeErupted = Empty
                        GCTooth.DateTimeUnErupted = Empty
                        GCTooth.DateTimeSupernumary = Empty
                        GCTooth.DateTimeWholeToothPlanToRemove = Empty
                        GCTooth.DateTimePonticDenture = Empty
                    Else
                        oTempDx.GCIsDirty = True
                    End If
                End If
            Next
                  
            'Set the Txs for the tooth
            For Each oTempTx In mObjSVGDraw.mObjDataProc.TeethTxData(iToothID)
                If Not oTempTx Is Nothing Then
                    If oTempTx.StatusID = TxStatusEnum.txRevised Then
                        'The status is revised, so don't draw it
                    Else
                        If oTempTx.StatusID = TxStatusEnum.txdeleted Then
                            setTreatRetreatToToothObject oTempTx.SurfaceSets(1).SymbolID, GCTooth
                            Set oTempTx = Nothing
                            GCTooth.IsExtractOrMissing = False
                            GCTooth.RemoveLocation = 0
                            GCTooth.DateTimeCrownRemoved = Empty
    
                            If GCTooth.DateTimeRootRemoved = "12:00:00 AM" And GCTooth.IsImplant Then
                            'When pontic deleted this resets which we do not want
                                If IsMissing(GCTooth.ToothID) = False Then
                                    GCTooth.DateTimeWholeToothRemoved = Empty
                                End If
                            End If
                           
                            If IsImplant(GCTooth.ToothID) = False Then
                                GCTooth.IsImplant = False
                                GCTooth.DateTimeImplant = Empty
                            End If
                           
                            If IsRootMissing(GCTooth.ToothID) = False Then
                                GCTooth.DateTimeRootRemoved = Empty
                            End If
                            
                            'GCTooth.dDateTimeRootRemoved = Empty
                            GCTooth.DateTimePlannedRootRemoved = Empty
                            GCTooth.DateTimeErupted = Empty
                            GCTooth.DateTimeUnErupted = Empty
                            GCTooth.DateTimeSupernumary = Empty
                            GCTooth.DateTimeWholeToothPlanToRemove = Empty
                            GCTooth.DateTimePonticDenture = Empty
                        Else
                            oTempTx.GCIsDirty = True
                        End If
                        
                        If Not oTx Is Nothing Then
                            If oTempTx Is oTx Then
                                oTempTx.GCIsDirty = True
                                sTmpMenuID = oTempTx.SurfaceSets(1).SymbolID
                                '                      oTempTx.SurfaceSets(1).SymbolID = getMenuID(oTx.SurfaceSets(1).SymbolID, oTx.StatusID)
                                GCTooth.DateTimeWholeToothPlanToRemove = Empty
                                GCTooth.DateTimePonticDenture = Empty
                                If IsRootMissing(GCTooth.ToothID) = False Then
                                     GCTooth.DateTimeRootRemoved = Empty
                                End If
                                GCTooth.DateTimePlannedRootRemoved = Empty
                            End If
                        End If
                    End If
                End If
            Next
                
            DrawSVG iToothID
        End If
    Next j
              
    Screen.MousePointer = vbDefault
    
End Sub
'Private Sub resetVisibleTooth(iToothID As Integer, bIsDx As Boolean)
'
''This sets the DrawThisIn..View.. property of all GCDxs and GCTxs to true for the given tooth
''so all the dxs and txs associated with this tooth will be drawn.
'
'Dim oGCDx As clsGCDx
'Dim oGCTx As clsGCTx
'Dim objPatTooth As patTooth
'
'Set objPatTooth = mColTeeth(CStr(iToothID))
'
'    'Reset the GCDxs
'    For Each oGCDx In objPatTooth.GCDxs
'        oGCDx.DrawThisInDxViewDx = True
'        oGCDx.DrawThisInTxViewDx = True
'    Next
'
'    'Reset the GCTxs
'    For Each oGCTx In objPatTooth.GCTxs
'        oGCTx.DrawThisInDxViewTx = True
'        oGCTx.DrawThisInTxViewTx = True
'    Next
'
'End Sub
Private Sub setTreatRetreatToToothObject(MenuID As String, GCTooth As clsGCTooth)

'Called when a Dx or Tx is deleted.
'MenuID is from the symbolid of the dx or tx
'This sub sets the tooth as having a previous endo for retreat codes

Dim objMenuItem As clsMenuItem
Dim tmpStrMenuID As Integer

    Set objMenuItem = mObjMenuLib.MenuItems(MenuID)
    tmpStrMenuID = CInt(objMenuItem.MnuButtonID)
    
    If (tmpStrMenuID = 47) Or (tmpStrMenuID = 48) Then '47 defective 48 preexisting
        GCTooth.HasPrevEndo = False
    End If
    
End Sub

Public Sub DrawSVG(ToothID As Integer, Optional SurfaceSet As String, Optional MenuID As String, Optional objDLDx As CHCSII_Dental_DXTX.clsDx, Optional objDLTx As CHCSII_Dental_DXTX.clsTx)

Dim colDX As Collection
Dim colTX As Collection
Dim GCTooth As clsGCTooth
Dim i As Integer
Dim strArrTeeth() As String
Dim j As Integer
Dim oDx As CHCSII_Dental_DXTX.clsDx
Dim oTx As CHCSII_Dental_DXTX.clsTx

On Error GoTo ErrorHandler

    Set GCTooth = mColTeeth(CStr(ToothID))
    Set mObjSVGDraw.mColTeeth = mColTeeth
   
    If GCTooth Is Nothing Then
        Set GCTooth = InitializeTooth(moPositions.GetPositionByToothNum(ToothID))
        'GCTooth.ToothVisible = True
        mColTeeth.Add GCTooth, CStr(GCTooth.ToothID)
    End If
    
    If Len(SurfaceSet) > 0 And Len(MenuID) > 0 Then
        If Not objDLDx Is Nothing Then
            objDLDx.GCIsDirty = True
            mObjSVGDraw.mObjDataProc.TeethDxData(ToothID).Add objDLDx
        End If
        
        If Not objDLTx Is Nothing Then
            objDLTx.GCIsDirty = True
            mObjSVGDraw.mObjDataProc.TeethTxData(ToothID).Add objDLTx
        End If
        
    End If

    Set colDX = New Collection
    Set colTX = New Collection
    
    mObjSVGDraw.DrawSVG colDX, colTX, GCTooth, objDLDx, objDLTx

    'Only need to do this when a user changes the status (or something) of a DLDx
    'Reorder collection based on modified date
    For i = 1 To colDX.Count
        GCTooth.GCDxs.Add colDX(i)
    Next i
    For i = 1 To colTX.Count
        GCTooth.GCTxs.Add colTX(i)
    Next i

    'Precedence for date and time modified, Dxs and Txs
    SetPrecedenceForDateTimeDxTx GCTooth

    'Precedence for coronal areas
    SetPrecedenceForSealantAndSurfaces GCTooth
    
    'Disable any pontics if planned extract
    SetPrecedenceDisableSurfacesPlannedExtract GCTooth
    
    'Viewable Surfaces for each view; DxView, TxView
    SetPrecedenceForDxViewSurfaces GCTooth
    SetPrecedenceForTxViewSurfaces GCTooth
    SetPrecedenceForDx603aViewSurfaces GCTooth

Exit Sub
    
ErrorHandler:
    If Err.Number = 5 Then
        'do nothing
    Else
        MsgBxARMd Err.Number & "," & Err.Description & ",clsManage.drawSVG, SVGManager"
    End If
    
    Resume Next

End Sub
Private Sub SetPrecedenceForDateTimeDxTx(GCTooth As clsGCTooth)

'This sub contains display rules for dxs and txs for removed roots, removed crown, and removed whole tooth for the given tooth

Dim oGCDx As clsGCDx
Dim oGCTx As clsGCTx

    'Dx rules for removed roots, removed crown, and removed whole tooth
    For Each oGCDx In GCTooth.GCDxs
            
        'Reset the Draw vars based on the status
        oGCDx.DxStatus = oGCDx.DxStatus
        
        If oGCDx.ModifiedDateTime <= GCTooth.DateTimeRootRemoved And GCTooth.DateTimeRootRemoved <> "12:00:00 AM" Then
            'If the root was removed after the dx was modified and the dx applies to the root, then don't draw the dx
            If oGCDx.ToothLocation = eRootSection Then
                oGCDx.DrawThisInDxViewDx = False
                oGCDx.DrawThisInTxViewDx = False
                oGCDx.DrawThisInDx603aViewDx = False
            End If
        End If
         
        If oGCDx.ModifiedDateTime <= GCTooth.DateTimeCrownRemoved And GCTooth.DateTimeCrownRemoved <> "12:00:00 AM" Then
            'if the crown was removed after the dx was modified and the dx applies to the coronal section, then don't draw the dx
            If oGCDx.ToothLocation = eCoronalSection Then
                oGCDx.DrawThisInDxViewDx = False
                oGCDx.DrawThisInTxViewDx = False
                oGCDx.DrawThisInDx603aViewDx = False
            End If
        End If
        
        If oGCDx.ModifiedDateTime <= GCTooth.DateTimeWholeToothRemoved And GCTooth.DateTimeWholeToothRemoved <> "12:00:00 AM" Then
            'If the whole tooth was removed after the dx was modified and
            'the dx applies to the whole tooth or the selected surfaces
            'and the tooth is not an implant, then don't draw the dx
            If (oGCDx.ToothLocation = eWholeTooth Or oGCDx.ToothLocation = eSelectedSurfaces) And GCTooth.IsImplant = False Then
                oGCDx.DrawThisInDxViewDx = False
                oGCDx.DrawThisInTxViewDx = False
                oGCDx.DrawThisInDx603aViewDx = False
            End If
        End If
        
    Next
    
    'Tx rules for removed roots, removed crown, and removed whole tooth
    For Each oGCTx In GCTooth.GCTxs
        
        'Reset the Draw vars based on the status
        oGCTx.TxStatus = oGCTx.TxStatus
    
        If oGCTx.ModifiedDateTime <= GCTooth.DateTimeRootRemoved And GCTooth.DateTimeRootRemoved <> "12:00:00 AM" Then
            'If the root was removed after the tx was modified and the tx applies to the root section, then don't draw the tx
            If oGCTx.ToothLocation = eRootSection Then
                oGCTx.DrawThisInDxViewTx = False
                oGCTx.DrawThisInTxViewTx = False
                oGCTx.DrawThisInDx603aViewTx = False
            End If
        End If
      
        If oGCTx.ModifiedDateTime <= GCTooth.DateTimeCrownRemoved And GCTooth.DateTimeCrownRemoved <> "12:00:00 AM" Then
            'if the crown was removed after the tx was modified and the tx applies to the crown, then don't draw the tx
            If oGCTx.ToothLocation = eCoronalSection Then
                oGCTx.DrawThisInDxViewTx = False
                oGCTx.DrawThisInTxViewTx = False
                oGCTx.DrawThisInDx603aViewTx = False
            End If
        End If
      
        If oGCTx.ModifiedDateTime <= GCTooth.DateTimeWholeToothRemoved And GCTooth.DateTimeWholeToothRemoved <> "12:00:00 AM" Then
            'if the whole tooth was removed after the tx was modified and
            'the tx applies to the whole tooth or the selected surfaces
            'and the tooth is not an implant, then don't draw the tx
            If (oGCTx.ToothLocation = eWholeTooth Or oGCTx.ToothLocation = eSelectedSurfaces) And GCTooth.IsImplant = False Then
                oGCTx.DrawThisInDxViewTx = False
                oGCTx.DrawThisInTxViewTx = False
                oGCTx.DrawThisInDx603aViewTx = False
            End If
        End If
      
    Next

End Sub
Private Function CloneGCDx(oGCDx As clsGCDx) As clsGCDx

On Error GoTo errHandler
    
    Set CloneGCDx = New clsGCDx
    
    With CloneGCDx
        .ModifiedDateTime = oGCDx.ModifiedDateTime
        .DRC = oGCDx.DRC
        .DxStatus = oGCDx.DxStatus
        .ToothLocation = oGCDx.ToothLocation
        .MenuID = oGCDx.MenuID
        .ToothID = oGCDx.ToothID
        .Surface = oGCDx.Surface
        .SVGDx = oGCDx.SVGDx
        .DrawThisInDxViewDx = oGCDx.DrawThisInDxViewDx
        .DrawThisInTxViewDx = oGCDx.DrawThisInTxViewDx
    End With
    
Exit Function

errHandler:
       MsgBxARMd Err.Number & "," & Err.Description & ",clsManage.clonePatDx, SVGManager"
   
End Function
Private Function CloneGCTx(oGCTx As clsGCTx) As clsGCTx

On Error GoTo errHandler
    
    Set CloneGCTx = New clsGCTx
    
    With CloneGCTx
        .ModifiedDateTime = oGCTx.ModifiedDateTime
        .TxStatus = oGCTx.TxStatus
        .ToothLocation = oGCTx.ToothLocation
        .MenuID = oGCTx.MenuID
        .ToothID = oGCTx.ToothID
        .Surface = oGCTx.Surface
        .SVGTx = oGCTx.SVGTx
        .DrawThisInDxViewTx = oGCTx.DrawThisInDxViewTx
        .DrawThisInTxViewTx = oGCTx.DrawThisInTxViewTx
    End With
    
Exit Function

errHandler:
       MsgBxARMd Err.Number & "," & Err.Description & ",clsManage.clonePatTx, SVGManager"
     
End Function
Private Sub SetPrecedenceForDxViewSurfaces(GCTooth As clsGCTooth)

'This sub compares the surfaces of dxs and txs for Dx View

'Note: this is called first and then ViewableSurfaceTxView is called

'Reset the collection of GCDxs and GCTxs for this tooth in Dx view
'Iterate through the dxs and txs for the dx view backwards and don't draw the second dx or tx
'that has the same surface as a previous dx or tx

'Question: do all surfaces appear in the same order (delimited) no matter how the surfaces were selected?
'Question: what about dxs that have surfaces which overlap but are not equal?

Dim oGCDx As clsGCDx
Dim oGCTx As clsGCTx
Dim i As Integer
Dim j As Integer
Dim colSurfaces As Collection
Dim oSurfaceGroup As clsSurfaceGroup
Dim strArrSurfaceSet() As String

Dim bErr As Boolean

On Error GoTo errHandler
    
    'Reset the collection of GCDxs and GCTxs for this tooth in Dx view
    Set mColDxViewToothDxs = New colGCDxs
    Set mColDxViewToothTxs = New colGCTxs
    
    'Take all the GCDxs associated with this tooth and make a copy of them and assign the copy to
    'the Dx View for this tooth
    For Each oGCDx In GCTooth.GCDxs
        mColDxViewToothDxs.Add CloneGCDx(oGCDx)
    Next
    
    'Take all the GCTxs associated with this tooth and make a copy of them and assign the copy to
    'the Dx View for this tooth
    For Each oGCTx In GCTooth.GCTxs
        mColDxViewToothTxs.Add CloneGCTx(oGCTx)
    Next
    
    Set colSurfaces = New Collection
    
    'Loop through the Dxs for this tooth (for the Dx view) backwards and set the DrawThisInDxViewDx property
    'to false when I encounter a second dx with the same surface
    For i = mColDxViewToothDxs.Count To 1 Step -1
        Set oGCDx = mColDxViewToothDxs.Item(i)
        If oGCDx.Surface <> "777" And oGCDx.DrawThisInDxViewDx = True And IsRootSurface(GCTooth, oGCDx.Surface) = False Then
            'If the dx is not a whole tooth dx and it should be drawn in the Dx view and it is not on a root surface, then
            
            If oGCDx.Surface = "111" Then 'Any special code needs surface extraction
                'Add all the individual crown surfaces
                Set oSurfaceGroup = mObjSVGLib.SurfaceGroups(GCTooth.ToothID)
                strArrSurfaceSet = Split(oSurfaceGroup.CrownSurfaces, ",")
                For j = 0 To UBound(strArrSurfaceSet)
                     colSurfaces.Add strArrSurfaceSet(j), CStr(strArrSurfaceSet(j))
                Next j
            Else
                'Add the surface
                colSurfaces.Add oGCDx.Surface, CStr(oGCDx.Surface)
            End If
            
        End If
        If bErr Then
           'Only for complete denture & RPD keep everything
    '       If GCTooth.dDateTimePonticDenture = "12:00:00 AM" Then
                oGCDx.DrawThisInDxViewDx = False
    '       End If
           bErr = False
        End If
    Next i
    
    'Do the same as above for Txs
    For i = mColDxViewToothTxs.Count To 1 Step -1
        Set oGCTx = mColDxViewToothTxs.Item(i)
        
        If Not oGCTx.Surface = "777" And (oGCTx.DrawThisInDxViewTx) And Not IsRootSurface(GCTooth, oGCTx.Surface) Then
            If oGCTx.Surface = "111" Then 'Any special code needs surface extraction
                
                Set oSurfaceGroup = mObjSVGLib.SurfaceGroups(GCTooth.ToothID)
                strArrSurfaceSet = Split(oSurfaceGroup.CrownSurfaces, ",")
                For j = 0 To UBound(strArrSurfaceSet)
                    colSurfaces.Add strArrSurfaceSet(j), CStr(strArrSurfaceSet(j))
                Next j
            
            Else
                colSurfaces.Add oGCTx.Surface, CStr(oGCTx.Surface)
            End If
        End If
        
        If bErr Then
            'Don't draw the second tx with a matching surface in either view
            oGCTx.DrawThisInDxViewTx = False
            oGCTx.DrawThisInTxViewTx = False
            bErr = False
        End If
    Next i
    
    Set colSurfaces = Nothing
    
    'Add Dxs from tooth to teeth view collection
    mColDxViewTeethDxs.Add mColDxViewToothDxs, CStr(GCTooth.ToothID)
    
    'If new collection has been added already then overwrite the old collection
    If bErr Then
        mColDxViewTeethDxs.Remove (CStr(GCTooth.ToothID))
        mColDxViewTeethDxs.Add mColDxViewToothDxs, CStr(GCTooth.ToothID)
        bErr = False
    End If
    
    'Add Txs from the tooth to the teeth view collection
    mColDxViewTeethTxs.Add mColDxViewToothTxs, CStr(GCTooth.ToothID)
    
    'If new collection has been added already then overwrite the old collection
    If bErr Then
        mColDxViewTeethTxs.Remove (CStr(GCTooth.ToothID))
        mColDxViewTeethTxs.Add mColDxViewToothTxs, CStr(GCTooth.ToothID)
        bErr = False
    End If

Exit Sub

errHandler:
   
   If Err.Number = 457 Then
        bErr = True
        Err.Clear
        Resume Next
   Else
        MsgBxARMd Err.Number & "," & Err.Description & ",clsManage.SetPrecedenceForDxViewSurfaces, SVGManager"
   End If

End Sub
Private Sub SetPrecedenceForTxViewSurfaces(GCTooth As clsGCTooth)

'This sub compares the surfaces of dxs and txs for Tx View

'Reset the collection of GCDxs and GCTxs for this tooth in Tx view
'Iterate through the dxs and txs for the dx view backwards and don't draw the second dx or tx
'that has the same surface as a previous dx or tx

'Question: do all surfaces appear in the same order (delimited) no matter how the surfaces were selected?
'Question: what about dxs that have surfaces which overlap but are not equal?

Dim oGCDx As clsGCDx
Dim oGCTx As clsGCTx
Dim i As Integer
Dim j As Integer
Dim colSurfaces As Collection
Dim oSurfaceGroup As clsSurfaceGroup
Dim strArrSurfaceSet() As String

Dim bErr As Boolean

On Error GoTo errHandler

    'Reset the collection of GCDxs and GCTxs for this tooth in Tx view
    Set mColTxViewToothDxs = New colGCDxs
    Set mColTxViewToothTxs = New colGCTxs
    
    'Take all the GCDxs associated with this tooth and assign them to the Tx View for this tooth
    For Each oGCDx In GCTooth.GCDxs
        If oGCDx.DrawThisInDxViewDx = True Then
            'If I draw this Dx in Dx view then draw it in Tx view
            oGCDx.DrawThisInTxViewDx = True
        End If
        
        mColTxViewToothDxs.Add oGCDx
    Next
    
    
    'Take all the GCTxs associated with this tooth and assign them to the Tx View for this tooth
    For Each oGCTx In GCTooth.GCTxs
        mColTxViewToothTxs.Add oGCTx
    Next
    
    Set colSurfaces = New Collection
    
    'Loop through the Txs for this tooth (for the Tx view) backwards and set the DrawThisInTxViewTx and DrawThisInDxViewTx property
    'to false when I encounter a second tx with the same surface
    For i = mColTxViewToothTxs.Count To 1 Step -1
        Set oGCTx = mColTxViewToothTxs.Item(i)
        
        If oGCTx.Surface <> "777" And oGCTx.DrawThisInTxViewTx = True And IsRootSurface(GCTooth, oGCTx.Surface) = False Then
            'If the tx is not a whole tooth tx and it should be drawn in the tx view and it is not on a root surface, then
            
            If oGCTx.Surface = "111" Then 'Any special code needs surface extraction
                'Add all the individual crown surfaces
                Set oSurfaceGroup = mObjSVGLib.SurfaceGroups(GCTooth.ToothID)
                strArrSurfaceSet = Split(oSurfaceGroup.CrownSurfaces, ",")
                For j = 0 To UBound(strArrSurfaceSet)
                     colSurfaces.Add strArrSurfaceSet(j), CStr(strArrSurfaceSet(j))
                Next j
            Else
                'Add the surface
                colSurfaces.Add oGCTx.Surface, CStr(oGCTx.Surface)
            End If
        
        End If
        
        If bErr Then
            'This tx has the same surface as a previous tx in the collection so don't draw it.
            oGCTx.DrawThisInDxViewTx = False
            oGCTx.DrawThisInTxViewTx = False
            oGCTx.DrawThisInDx603aViewTx = False
            bErr = False
        End If
        
    Next i
    
    'Do the same as above with Dxs
    For i = mColTxViewToothDxs.Count To 1 Step -1
        Set oGCDx = mColTxViewToothDxs.Item(i)
        If oGCDx.Surface <> "777" And oGCDx.DrawThisInTxViewDx = True And IsRootSurface(GCTooth, oGCDx.Surface) = False Then
            If oGCDx.Surface = "111" Then 'Any special code needs surface extraction
                'Add all the individual crown surfaces
                Set oSurfaceGroup = mObjSVGLib.SurfaceGroups(GCTooth.ToothID)
                strArrSurfaceSet = Split(oSurfaceGroup.CrownSurfaces, ",")
                For j = 0 To UBound(strArrSurfaceSet)
                     colSurfaces.Add strArrSurfaceSet(j), CStr(strArrSurfaceSet(j))
                Next j
            Else
                'Add the surface
                colSurfaces.Add oGCDx.Surface, CStr(oGCDx.Surface)
            End If
        End If
        
        If bErr Then
           'Only for complete denture & RPD keep everything
    '       If GCTooth.dDateTimePonticDenture = "12:00:00 AM" Then
                oGCDx.DrawThisInTxViewDx = False
    '       End If
           bErr = False
        End If
    Next i

    Set colSurfaces = Nothing
    
    'Add Dxs from the tooth to the teeth view collection
    mColTxViewTeethDxs.Add mColTxViewToothDxs, CStr(GCTooth.ToothID)
    
    'If new collection has been added already then overwrite the old collection
    If bErr Then
       mColTxViewTeethDxs.Remove (CStr(GCTooth.ToothID))
       mColTxViewTeethDxs.Add mColTxViewToothDxs, CStr(GCTooth.ToothID)
       bErr = False
    End If
    
    'Add Txs from tooth to teeth view collection
    mColTxViewTeethTxs.Add mColTxViewToothTxs, CStr(GCTooth.ToothID)
    
    'If new collection has been added already then overwrite the old collection
    If bErr Then
       mColTxViewTeethTxs.Remove (CStr(GCTooth.ToothID))
       mColTxViewTeethTxs.Add mColTxViewToothTxs, CStr(GCTooth.ToothID)
       bErr = False
    End If
    
Exit Sub

errHandler:
   
    If Err.Number = 457 Then
        bErr = True
        Err.Clear
        Resume Next
    Else
        MsgBxARMd Err.Number & "," & Err.Description & ",clsManage.SetPrecedenceForTxViewSurfaces, SVGManager"
    End If

End Sub
Private Sub SetPrecedenceForDx603aViewSurfaces(GCTooth As clsGCTooth)

'This sub compares the surfaces of dxs and txs for Tx View

'Reset the collection of GCDxs and GCTxs for this tooth in Tx view
'Iterate through the dxs and txs for the dx view backwards and don't draw the second dx or tx
'that has the same surface as a previous dx or tx

'Question: do all surfaces appear in the same order (delimited) no matter how the surfaces were selected?
'Question: what about dxs that have surfaces which overlap but are not equal?

Dim oGCDx As clsGCDx
Dim oGCTx As clsGCTx
Dim i As Integer
Dim j As Integer
Dim colSurfaces As Collection
Dim oSurfaceGroup As clsSurfaceGroup
Dim strArrSurfaceSet() As String

Dim bErr As Boolean

On Error GoTo errHandler

    'Reset the collection of GCDxs and GCTxs for this tooth in Tx view
    Set mColDx603aViewToothDxs = New colGCDxs
    Set mColDx603aViewToothTxs = New colGCTxs
    
    'Take all the GCDxs associated with this tooth and assign them to the Tx View for this tooth
    For Each oGCDx In GCTooth.GCDxs
'        If oGCDx.DrawThisInDx603aViewDx = True Then
'            'If I draw this Dx in Dx view then draw it in Tx view
'            oGCDx.DrawThisInTxViewDx = True
'        End If
        
        mColDx603aViewToothDxs.Add oGCDx
    Next
    
    
    'Take all the GCTxs associated with this tooth and assign them to the Tx View for this tooth
    For Each oGCTx In GCTooth.GCTxs
        mColDx603aViewToothTxs.Add oGCTx
    Next
    
    Set colSurfaces = New Collection
    
    'Loop through the Txs for this tooth (for the Tx view) backwards and set the DrawThisInTxViewTx and DrawThisInDxViewTx property
    'to false when I encounter a second tx with the same surface
    For i = mColDx603aViewToothTxs.Count To 1 Step -1
        Set oGCTx = mColDx603aViewToothTxs.Item(i)
        
        If oGCTx.Surface <> "777" And oGCTx.DrawThisInDx603aViewTx = True And IsRootSurface(GCTooth, oGCTx.Surface) = False Then
            'If the tx is not a whole tooth tx and it should be drawn in the tx view and it is not on a root surface, then
            
            If oGCTx.Surface = "111" Then 'Any special code needs surface extraction
                'Add all the individual crown surfaces
                Set oSurfaceGroup = mObjSVGLib.SurfaceGroups(GCTooth.ToothID)
                strArrSurfaceSet = Split(oSurfaceGroup.CrownSurfaces, ",")
                For j = 0 To UBound(strArrSurfaceSet)
                     colSurfaces.Add strArrSurfaceSet(j), CStr(strArrSurfaceSet(j))
                Next j
            Else
                'Add the surface
                colSurfaces.Add oGCTx.Surface, CStr(oGCTx.Surface)
            End If
        
        End If
        
        If bErr Then
            'This tx has the same surface as a previous tx in the collection so don't draw it.
            oGCTx.DrawThisInDx603aViewTx = False
            bErr = False
        End If
        
    Next i
    
    'Do the same as above with Dxs
    For i = mColDx603aViewToothDxs.Count To 1 Step -1
        Set oGCDx = mColDx603aViewToothDxs.Item(i)
        If oGCDx.Surface <> "777" And oGCDx.DrawThisInDx603aViewDx = True And IsRootSurface(GCTooth, oGCDx.Surface) = False Then
            If oGCDx.Surface = "111" Then 'Any special code needs surface extraction
                'Add all the individual crown surfaces
                Set oSurfaceGroup = mObjSVGLib.SurfaceGroups(GCTooth.ToothID)
                strArrSurfaceSet = Split(oSurfaceGroup.CrownSurfaces, ",")
                For j = 0 To UBound(strArrSurfaceSet)
                     colSurfaces.Add strArrSurfaceSet(j), CStr(strArrSurfaceSet(j))
                Next j
            Else
                'Add the surface
                colSurfaces.Add oGCDx.Surface, CStr(oGCDx.Surface)
            End If
        End If
        
        If bErr Then
           'Only for complete denture & RPD keep everything
    '       If GCTooth.dDateTimePonticDenture = "12:00:00 AM" Then
                oGCDx.DrawThisInDx603aViewDx = False
    '       End If
           bErr = False
        End If
    Next i

    Set colSurfaces = Nothing
    
    'Add Dxs from the tooth to the teeth view collection
    mColDx603aViewTeethDxs.Add mColDx603aViewToothDxs, CStr(GCTooth.ToothID)
    
    'If new collection has been added already then overwrite the old collection
    If bErr Then
       mColDx603aViewTeethDxs.Remove (CStr(GCTooth.ToothID))
       mColDx603aViewTeethDxs.Add mColDx603aViewToothDxs, CStr(GCTooth.ToothID)
       bErr = False
    End If
    
    'Add Txs from tooth to teeth view collection
    mColDx603aViewTeethTxs.Add mColDx603aViewToothTxs, CStr(GCTooth.ToothID)
    
    'If new collection has been added already then overwrite the old collection
    If bErr Then
       mColDx603aViewTeethTxs.Remove (CStr(GCTooth.ToothID))
       mColDx603aViewTeethTxs.Add mColDx603aViewToothTxs, CStr(GCTooth.ToothID)
       bErr = False
    End If
    
Exit Sub

errHandler:
   
    If Err.Number = 457 Then
        bErr = True
        Err.Clear
        Resume Next
        Resume
    Else
        MsgBxARMd Err.Number & "," & Err.Description & ",clsManage.SetPrecedenceForDxView603aSurfaces, SVGManager"
    End If

End Sub

Private Sub SetPrecedenceDisableSurfacesPlannedExtract(GCTooth As clsGCTooth)

'Dim oGCDx As clsGCDx
Dim oGCTx As clsGCTx

    For Each oGCTx In GCTooth.GCTxs
        'If the tooth is "planned to be removed"
        'and it will show in TxView
        'and it is not a whole tooth Tx
        'and the Tx is on the coronal section
        'then don't display the Tx in TxView
        If GCTooth.DateTimeWholeToothPlanToRemove <> "12:00:00 AM" _
            And oGCTx.DrawThisInTxViewTx = True _
            And oGCTx.Surface <> "777" _
            And oGCTx.ToothLocation = eCoronalSection Then
            
                oGCTx.DrawThisInTxViewTx = False
                
        End If
    Next

End Sub
Private Sub SetPrecedenceForSealantAndSurfaces(GCTooth As clsGCTooth)

'Determines which dxs and txs should display based on Sealant

Dim oGCDx As clsGCDx
Dim oGCTx As clsGCTx
Dim colSealantDxs As Collection
Dim colSealantTxs As Collection
Dim colSurfaces As Collection
Dim oSurfaceGroup As clsSurfaceGroup

Dim i As Integer
Dim j As Integer
Dim bErr As Boolean
Dim strArrSurfaceSet() As String
Dim bFoundDxOnOcclusalSurfaceInDxView As Boolean
Dim bFoundDxOnOcclusalSurfaceInTxView As Boolean
Dim bFoundTxOnOcclusalSurfaceInDxView As Boolean
Dim bFoundTxOnOcclusalSurfaceInTxView As Boolean

'37476
Dim bFoundTxOnOcclusalSurfaceInDx603aView As Boolean

Dim bFoundTxPOnOcclusalSurfaceInDxView As Boolean
Dim bFoundTxPOnOcclusalSurfaceInTxView As Boolean

On Error GoTo ErrHnd

    Set colSealantDxs = New Collection
    Set colSealantTxs = New Collection
    
    Set oSurfaceGroup = mObjSVGLib.SurfaceGroups(GCTooth.ToothID)
    strArrSurfaceSet = Split(oSurfaceGroup.CrownSurfaces, ",")
    
    'Add all Dx sealants to the colSealantDxs collection
    For Each oGCDx In GCTooth.GCDxs
        If oGCDx.DrawThisInDxViewDx = True Or oGCDx.DrawThisInTxViewDx = True Then
            If oGCDx.RootMenuID = 28 Then   'Root Menu for Sealant
                colSealantDxs.Add oGCDx
            End If
        End If
    Next
    
    'Add all Tx sealants to the colSelantTxs collection
    For Each oGCTx In GCTooth.GCTxs
        If oGCTx.DrawThisInDxViewTx Or oGCTx.DrawThisInTxViewTx Then
            If oGCTx.RootMenuID = 29 Or oGCTx.RootMenuID = 31 Then   'Root Menu for Sealant
                colSealantTxs.Add oGCTx
            End If
        End If
    Next
    
    'Find any non-sealant coronal dx's if there are sealant dx's or tx's
    If colSealantDxs.Count > 0 Or colSealantTxs.Count > 0 Then
        For Each oGCDx In GCTooth.GCDxs
           If oGCDx.DrawThisInDxViewDx Or oGCDx.DrawThisInTxViewDx Then
                    For j = 0 To UBound(strArrSurfaceSet)
                        If (strArrSurfaceSet(j) = oGCDx.Surface Or (oGCDx.Surface = 111 And UBound(strArrSurfaceSet) > 0)) And oGCDx.RootMenuID <> 28 Then  'Test for same surface or crown surface macro for occlussal type teeth
                            If oGCDx.DrawThisInDxViewDx = True Then
                                bFoundDxOnOcclusalSurfaceInDxView = True
                            End If
                            If oGCDx.DrawThisInTxViewDx = True Then
                                bFoundDxOnOcclusalSurfaceInTxView = True
                            End If
                            
'                            bFoundDxOcclusalSurface = True
                            Exit For
                        End If
                    Next j
            End If
        Next
    End If
    
    Dim test As Boolean
    Dim bSetSealantTxsInDxView As Boolean
    Dim bSetSealantTxsInTxView As Boolean
    
    'If a dx sealant is drawn in dx view, then don't draw any tx sealants in dx view
    If colSealantDxs.Count > 0 Then
        For Each oGCDx In colSealantDxs
            If oGCDx.DrawThisInDxViewDx = True Then
                'Hide all Tx sealants in Dx view
                For Each oGCTx In colSealantTxs
                    oGCTx.DrawThisInDxViewTx = False
                Next
                Exit For
            End If
        Next
    End If
    'If a tx sealant is drawn in tx view, then don't draw any dx sealants in tx view
    If colSealantTxs.Count > 0 Then
        For Each oGCTx In colSealantTxs
            If oGCTx.DrawThisInTxViewTx = True Then
                For Each oGCDx In colSealantDxs
                    oGCDx.DrawThisInTxViewDx = False
                Next
                Exit For
            End If
        Next
    End If
    
'    If colSealantDxs.Count > 0 Then
'        For Each oGCTx In colSealantTxs
'            oGCTx.DrawThisInDxViewTx = False
'            oGCTx.DrawThisInTxViewTx = False
'        Next
'    End If
    
    'Find any non-sealant coronal tx's if there are sealant dx's or tx's
    If colSealantTxs.Count > 0 Or colSealantDxs.Count > 0 Then
        For Each oGCTx In GCTooth.GCTxs
           If oGCTx.DrawThisInDxViewTx Or oGCTx.DrawThisInTxViewTx Then
                    For j = 0 To UBound(strArrSurfaceSet)
                        If (strArrSurfaceSet(j) = oGCTx.Surface Or (oGCTx.Surface = 111 And UBound(strArrSurfaceSet) > 0)) And Not (oGCTx.RootMenuID = 29 Or oGCTx.RootMenuID = 31) Then  'Test for same surface or crown surface macro for occlussal type teeth
                            
                            If oGCTx.TxStatus = txPlanned Or oGCTx.TxStatus = txInProgress Then
                                    'Txp
                                    bFoundTxPOnOcclusalSurfaceInDxView = oGCTx.DrawThisInDxViewTx
                                    bFoundTxPOnOcclusalSurfaceInTxView = oGCTx.DrawThisInTxViewTx
                            Else
                                    bFoundTxOnOcclusalSurfaceInDxView = oGCTx.DrawThisInDxViewTx
                                    bFoundTxOnOcclusalSurfaceInTxView = oGCTx.DrawThisInTxViewTx
                                    bFoundTxOnOcclusalSurfaceInDx603aView = oGCTx.DrawThisInDx603aViewTx
                            End If
    
                        End If
                    Next j
            End If
        Next
    End If
    
    'If there is a tx on an coronal surface in dx view, then don't display the dx of sealant in the dx view
    'If there is a tx on an coronal surface in tx view, then don't display the dx of sealant in the tx view
    For Each oGCTx In colSealantTxs
        If bFoundDxOnOcclusalSurfaceInDxView Or bFoundTxOnOcclusalSurfaceInDxView Or bFoundTxPOnOcclusalSurfaceInDxView Then
            oGCTx.DrawThisInDxViewTx = False
        End If
        If bFoundDxOnOcclusalSurfaceInTxView Or bFoundTxOnOcclusalSurfaceInTxView Or bFoundTxPOnOcclusalSurfaceInTxView Then
            oGCTx.DrawThisInTxViewTx = False
        End If
        '37476
        If bFoundTxOnOcclusalSurfaceInDx603aView = True Then
            oGCTx.DrawThisInDx603aViewTx = False
        End If
    Next
    
    'If there is a dx on an coronal surface in dx view, then don't display the dx of sealant in the dx view
    'If there is a dx on an coronal surface in tx view, then don't display the dx of sealant in the tx view
    For Each oGCDx In colSealantDxs
        If bFoundDxOnOcclusalSurfaceInDxView Or bFoundTxOnOcclusalSurfaceInDxView Or bFoundTxPOnOcclusalSurfaceInDxView Then
            oGCDx.DrawThisInDxViewDx = False
        End If
        If bFoundDxOnOcclusalSurfaceInTxView Or bFoundTxOnOcclusalSurfaceInTxView Or bFoundTxPOnOcclusalSurfaceInTxView Then
            oGCDx.DrawThisInTxViewDx = False
        End If
    Next
        
'****************************************

'    'Dx's on Dx's (the same surface)
'    Set colSurfaces = New Collection
'    For i = GCTooth.GCDxs.Count To 1 Step -1
'        Set oGCDx = GCTooth.GCDxs().Item(i)
'        If oGCDx.Surface <> "777" And (oGCDx.DrawThisInDxViewDx Or oGCDx.DrawThisInTxViewDx) And Not IsRootSurface(GCTooth, oGCDx.Surface) Then
'            If Not bFoundDxOcclusalSurface And Not bFoundTxOcclusalSurface Then
'                'Add the oGCDx to a collection with the surface as the key, if i get an error, then that surface already exists
'                'and don't display it.
'                colSurfaces.Add oGCDx.Surface, CStr(oGCDx.Surface)
'            End If
'        End If
'        If bErr Then
'           'Only for complete denture & RPD keep everything
'    '       If GCTooth.dDateTimePonticDenture = "12:00:00 AM" Then
'                    oGCDx.DrawThisInDxViewDx = False
'                    oGCDx.DrawThisInTxViewDx = False
'    '       End If
'           bErr = False
'        End If
'    Next i
'
'    Set colSurfaces = New Collection
'
'    'Look for other surfaces on top of each other Tx's
'    For i = GCTooth.GCTxs.Count To 1 Step -1
'        Set oGCTx = GCTooth.GCTxs().Item(i)
'        If Not oGCTx.Surface = "777" And (oGCTx.DrawThisInDxViewTx Or oGCTx.DrawThisInTxViewTx) And Not IsRootSurface(GCTooth, oGCTx.Surface) Then
'            If bFoundDxOcclusalSurface = False And bFoundTxOcclusalSurface = False Then
'                colSurfaces.Add oGCTx.Surface, CStr(oGCTx.Surface)
'            End If
'        End If
'        If bErr Then
'            If bFoundTxPOcclusalSurface = True Then
'                oGCTx.DrawThisInDxViewTx = False
'                oGCTx.DrawThisInTxViewTx = False
'            Else
'                If oGCTx.TxStatus = txPlanned Or oGCTx.TxStatus = txInProgress Then
'                    oGCTx.DrawThisInDxViewTx = False
'                    oGCTx.DrawThisInTxViewTx = True
'
'                Else
'                    oGCTx.DrawThisInDxViewTx = True
'                    oGCTx.DrawThisInTxViewTx = False
'                End If
'            End If
'
'            bErr = False
'        End If
'
'    Next i

    Set colSealantDxs = Nothing
    Set colSealantTxs = Nothing
    Set colSurfaces = Nothing
   
Exit Sub

ErrHnd:
   If Err.Number = 457 Then
        bErr = True
        Err.Clear
        Resume Next
   Else
        MsgBxARMd Err.Number & "," & Err.Description & ",clsManage.SetPrecedenceForSealantAndSurface, SVGManager"
   End If

End Sub
Public Function IsRootSurface(GCTooth As clsGCTooth, sSurface As String) As Boolean

'Determine if the passed in surface is a root surface for the passed in tooth

'Dim objMacro As clsMacro
Dim oSurfaceGroup As clsSurfaceGroup
Dim strArrSurfaceGroup() As String
Dim i As Integer

    Set oSurfaceGroup = mObjSVGLib.SurfaceGroups(GCTooth.ToothID)
    strArrSurfaceGroup = Split(oSurfaceGroup.RootSurfaces, ",")
    
    For i = 0 To UBound(strArrSurfaceGroup)
        If strArrSurfaceGroup(i) = sSurface Then
            IsRootSurface = True
            Exit For
        End If
    Next i

End Function
Private Sub Class_Initialize()
   miAge = 18 'the parent should call initialize
End Sub
Private Sub InitializeObjects()

   Set mObjSVGDraw = New clsSVGDraw
   
   Dim mOSVGLibSF As New CHCSII_Dental_SVG_Lib.clsSVGLibSF
   Set mObjSVGLib = mOSVGLibSF.SVGLib
   
   Set mObjSVGQueue = New clsQueue
   Set mObjSVGQueue.mObjManage = Me
   
   Set mObjMenuLib = mOSF.Singleton
   
   Set mColTeeth = New Collection
   
   Set mColDxViewTeethDxs = New Collection
   Set mColDxViewTeethTxs = New Collection
   Set mColTxViewTeethDxs = New Collection
   Set mColTxViewTeethTxs = New Collection
    Set mColDx603aViewTeethDxs = New Collection
    Set mColDx603aViewTeethTxs = New Collection

   Set mOSF = New CHCSII_Dntl_Mnu_Lib.SF
   
   Load frmMenu
   Set frmMenu.Parent = Me
   
End Sub
Friend Sub TerminateObjects()
    
    frmMenu.tmrMultithread.Enabled = False
    Unload frmMenu
    
    Set mObjMenuLib = Nothing
    Set mObjSVGDraw = Nothing
    Set mObjSVGQueue = Nothing
    Set mObjSVGLib = Nothing
    Set mColTeeth = Nothing
    Set mColDxViewToothDxs = Nothing
    Set mColDxViewToothTxs = Nothing
    Set mColTxViewToothDxs = Nothing
    Set mColTxViewToothTxs = Nothing
    Set mColTeeth = Nothing
    
    Set mColDxViewTeethDxs = Nothing
    Set mColDxViewTeethTxs = Nothing
    Set mColTxViewTeethDxs = Nothing
    Set mColTxViewTeethTxs = Nothing
    
    Set mOSF = Nothing
    
End Sub
Public Function loadTeeth(Positions As colPositions, ViewMode As eViewMode)

  Dim iCounter As Integer
  Dim oPos As clsPosition
  'Dim strArrTeethPath() As String
  Dim GCTooth As clsGCTooth
  
  On Error GoTo ErrorHandler
  
  'Debug.Assert False
  
  'Set moPositions?
  Set moPositions = Positions
  
  Set frmMenu.Positions = moPositions
  frmMenu.ViewMode = ViewMode
  
'  frmMenu.sAgePath = sAgePath
'  frmMenu.bIsDx = bIsDx

'TODO: This should only happend the first time? when GC loads
    For Each oPos In Positions
    
        Set GCTooth = mColTeeth(CStr(oPos.ToothNum))
        
        If GCTooth Is Nothing Then
            Set GCTooth = InitializeTooth(oPos)
            'GCTooth.ToothVisible = True
            mColTeeth.Add GCTooth, CStr(GCTooth.ToothID)
        End If
        
        'GCTooth.ToothVisible = True
        
        'Init it the first time
        If ViewMode = eDxView Then
            GCTooth.IsImplant = False
            GCTooth.IsExtractOrMissing = False
            GCTooth.RemoveLocation = 0
        
            Set GCTooth.GCDxs = Nothing
            Set GCTooth.GCDxs = New colGCDxs
            Set GCTooth.GCTxs = Nothing
            Set GCTooth.GCTxs = New colGCTxs
        End If
    Next

'  strArrTeethPath = Split(sAgePath, ",")
'  For i = 0 To UBound(strArrTeethPath)
'
'        Set GCTooth = mColTeeth(strArrTeethPath(i))
'
'        If GCTooth Is Nothing Then
'            Set GCTooth = InitializeTooth(CInt(strArrTeethPath(i)))
'            GCTooth.ToothVisible = True
'            mColTeeth.Add GCTooth, CStr(GCTooth.ToothID)
'        End If
'
'        GCTooth.ToothVisible = True
'
'        If bIsDx Then 'Being initialized 2x for dx/tx view.
'            GCTooth.IsImplant = False
'            GCTooth.IsExtractOrMissing = False
'            GCTooth.RemoveLocation = 0
'
'            Set GCTooth.GCDxs = Nothing
'            Set GCTooth.GCDxs = New colGCDxs
'            Set GCTooth.GCTxs = Nothing
'            Set GCTooth.GCTxs = New colGCTxs
'        End If
'  Next i
  
  'Set this so we don't duplicate our collections later on
  
  If mbPatientChanged Then  ' We really don't need (added Insurance) a boolean If/Else, but we must reset the property
        mObjSVGQueue.mbPatientChanged = mbPatientChanged
        mbPatientChanged = False
  Else
        mObjSVGQueue.mbPatientChanged = mbPatientChanged
  End If
  
  Debug.Print "frmMenu.tmrMultithread.Enabled = True"
  frmMenu.tmrMultithread.Enabled = True
  DoEvents
  
Exit Function
    
ErrorHandler:
    'Ignore Invalid procedure call or argument
    If Err.Number <> 5 Then
        MsgBxARMd Err.Number & "," & Err.Description & ",clsManage.loadTeeth, SVGManager"
    Else
        Set GCTooth = Nothing
    End If
    Resume Next
    
End Function

Private Sub Class_Terminate()

'singleton's don't terminate correctly
'    TerminateObjects

End Sub

Private Sub mObjSVGQueue_ToothLoaded(iToothID As Integer, strToothSVG As String)
  RaiseEvent ToothLoaded(iToothID, strToothSVG)
End Sub
