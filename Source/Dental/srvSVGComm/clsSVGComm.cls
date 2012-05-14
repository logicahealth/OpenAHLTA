VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSVGComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_TeethObject As clsTeeth
Private m_ToothObject As clsTooth
Private m_DataObject As CHCSII_Dental_DXTX.clsNotes
Private m_DataProcObject As New CHCSII_Dental_SVG_Manager.clsDataProc
Private moSVGMngrSingleton As New CHCSII_Dental_SVG_Manager.SF
Public mObjSVGMngr As CHCSII_Dental_SVG_Manager.clsManage
Public mObjPSR As clsPSR

Private m_PatientID As Long
Private m_Age As Integer
Private m_DentalEncounterID As Long  ' 12/21
Private m_ProviderId As Long
Private m_bNewPatient As Boolean
Private m_strLastKey As String

Private Const TEETH_HTML_FILE_DX = "teeth_dx.html"
Private Const TEETH_HTML_FILE_TX = "teeth_tx.html"
'37476
Private Const TEETH_HTML_FILE_DX_603a = "teeth_603a.html"
Private Const TOOTH_HTML_FILE = "tooth2.html"
Private Const TEETH_SVG_FILE_DX = "teethx_dx.svg"
Private Const TEETH_SVG_FILE_TX = "teethx_tx.svg"
'37476
Private Const TEETH_SVG_FILE_603a = "teethx_603a.svg"
Private Const TOOTH_SVG_FILE = "toothx.svg"
Private Const TEETH_SVG_FILE_REP = "teethrpt.svg"
Public mbReadOnly As Boolean

'48054 SF 1-23-04 - Use comm to relay the message from the Grid to GC
Public Event TxChangedStatus(Tx As CHCSII_Dental_DXTX.clsTx)

Public Property Set CWComm(vData As MMMHISComm.Comm)
    
    Set cwobjCWComm = vData
    
End Property


Public Property Set CWConfig(vData As MMMHISConfiguration.Config)

    Set cwobjConfig = vData
    
End Property

Public Property Set CWLogonEnviron(vData As MMMHISLogon.Logon)
    
    Set cwobjLogonEnviron = vData
    
End Property

Public Property Set CWPatient(vData As MMMHISPatient.Patient)

    Set cwobjPatient = vData
    
End Property


Public Property Let PatientID(lPatientID As Long)
    If lPatientID <> m_PatientID Then
        m_bNewPatient = True
        m_PatientID = lPatientID
        
        'Tell manager about patient change
        If Not mObjSVGMngr Is Nothing Then mObjSVGMngr.PatientID = lPatientID
        
        createDataLayer
    End If
    
    
    'Call initializeUser
    
End Property

Public Sub RefreshPatient()

        m_bNewPatient = True
        createDataLayer

End Sub

Public Property Get PatientID() As Long
    PatientID = m_PatientID
End Property

Public Function initializeUser() As Boolean
    m_TeethObject.InitializeManager
    createDataLayer
End Function
Public Property Let Age(nAge As Integer)
    m_Age = nAge
    
    'SF 4-09-04 Tell SVGManage what the patient's age is.
    mObjSVGMngr.Age = nAge
    
End Property

Public Property Get Age() As Integer
    Age = m_Age
End Property
' 12/21
Public Property Let DentalEncounterID(lDentalEncounterID As Long)
    m_DentalEncounterID = lDentalEncounterID
    If Not m_DataObject Is Nothing And Not lDentalEncounterID = 999 Then
        m_DataObject.DentalEncounterID = lDentalEncounterID
    End If
End Property
' 12/21
Public Property Get DentalEncounterID() As Long
    DentalEncounterID = m_DentalEncounterID
End Property

Public Property Get DataObject() As CHCSII_Dental_DXTX.clsNotes
    If m_DataObject Is Nothing Then
        Call createDataLayer
    End If
    Set DataObject = m_DataObject
End Property

Public Property Get DataProcObject() As CHCSII_Dental_SVG_Manager.clsDataProc
    Set DataProcObject = m_DataProcObject
End Property

Public Property Set TeethObject(oTeeth As clsTeeth)
    Set m_TeethObject = oTeeth
End Property

Public Property Get TeethObject() As clsTeeth
    Set TeethObject = m_TeethObject
End Property
Public Property Set ToothObject(oTooth As clsTooth)
    Set m_ToothObject = oTooth
End Property

Public Property Get ToothObject() As clsTooth
    Set ToothObject = m_ToothObject
End Property

Public Property Let ProviderID(lProviderID As Long)
    m_ProviderId = lProviderID
    If Not m_DataObject Is Nothing And Not lProviderID = 999 Then
    
        m_DataObject.ProviderID = lProviderID
        
    End If
End Property

Public Function initTeethLoadDx(myWin As Object)
    ' first to be called by svg document
    Call m_TeethObject.initDx(myWin)
    
    If m_PatientID > 0 And m_Age >= 0 Then
        Call m_TeethObject.CreateTeethDx(m_PatientID, m_Age)
    End If
           
End Function

Public Function initTeethLoadTx(myWin As Object)
    ' first to be called by svg document
    Call m_TeethObject.initTx(myWin)
    
    'If m_PatientID > 0 And m_Age > 0 Then
    Call m_TeethObject.CreateTeethTx
    'End If
           
End Function

Public Function initTeethLoadDx603a(myWin As Object)
    ' first to be called by svg document
    Call m_TeethObject.initDx603a(myWin)
    
    'If m_PatientID > 0 And m_Age > 0 Then
    Call m_TeethObject.CreateTeethDx603a
    'End If
           
End Function

Public Function initToothLoad(myWin As Object)
    ' first to be called by svg document
     Call m_ToothObject.init(myWin)
    
End Function

Public Function getTeethHTMLDx() As String
    ' called by graphic charting
    ' to generate svg file for the patient
    ' return the svg file name
      
    m_TeethObject.toothMode = eFullMouth
    If Not m_ToothObject Is Nothing Then
        Set m_TeethObject.ToothObject = m_ToothObject
    End If
       
    'SF 9-16-03 SCR 40464
    'getTeethHTMLDx = gobjShared.ImagesPath & "\" & TEETH_HTML_FILE_DX
    getTeethHTMLDx = Replace(gobjShared.ImagesPath & "\" & TEETH_HTML_FILE_DX, "\\", "\")
 
End Function

Public Function getTeethHTMLDx603a() As String
    ' called by graphic charting
    ' to generate svg file for the patient
    ' return the svg file name
      
    m_TeethObject.toothMode = eFullMouth
    If Not m_ToothObject Is Nothing Then
        Set m_TeethObject.ToothObject = m_ToothObject
    End If
       
    getTeethHTMLDx603a = Replace(gobjShared.ImagesPath & "\" & TEETH_HTML_FILE_DX_603a, "\\", "\")
 
End Function

Public Function getTeethHTMLTx() As String
    ' called by graphic charting
    ' to generate svg file for the patient
    ' return the svg file name
      
    'm_TeethObject.ToothMode = False
    'If Not m_ToothObject Is Nothing Then
    '    Set m_TeethObject.ToothObject = m_ToothObject
    'End If
       
    'SF 9-16-03 SCR 40464
    'getTeethHTMLTx = gobjShared.ImagesPath & "\" & TEETH_HTML_FILE_TX
    getTeethHTMLTx = Replace(gobjShared.ImagesPath & "\" & TEETH_HTML_FILE_TX, "\\", "\")
 
End Function

Public Function getToothHTML() As String
    ' called by graphic charting
    ' to generate svg file for the patient
    ' return the svg file name

    'SF 9-16-03 SCR 40464
    'getToothHTML = gobjShared.ImagesPath & "\" & TOOTH_HTML_FILE
    getToothHTML = Replace(gobjShared.ImagesPath & "\" & TOOTH_HTML_FILE, "\\", "\")

End Function
Private Function createDataLayer()

''    '''''''''''''''''''''''''
'#If BACKDOOR Then
'    Dim oLogon As InternalLogon
'    Dim bConnected As Boolean
'
'    Set oLogon = New InternalLogon
'    bConnected = oLogon.bGetConnected("mmm", "mmm")
'    Set oLogon = Nothing
'    If (Not bConnected) Then
'        MsgBxARMd "Logon Failed!"
'    End If
'#End If
''    '''''''''''''''''''''''''
    
    On Error GoTo ErrorHandler
     
    If m_DataObject Is Nothing Then
                
        Set m_DataObject = New CHCSII_Dental_DXTX.clsNotes
    End If
        
    If m_bNewPatient Then 'And m_DentalEncounterID > 0 Then
    
        m_bNewPatient = False
        ''''''''''''''''''''''
        Dim colGroups As CHCSII_Dental_DXTX.colGroups
        
        'Debug.Print "** START LoadGroups - " & Now
        
        Call m_DataObject.LoadGroups(cwobjPatient, cwobjLogonEnviron, cwobjConfig, cwobjCWComm, m_DentalEncounterID, m_ProviderId)
        'Call m_DataObject.LoadGroups(m_PatientID, m_DentalEncounterID, m_ProviderId)
        
        'Debug.Print "** END LoadGroups - " & Now
        
        Set colGroups = m_DataObject.Groups
            
        Set m_DataProcObject = mObjSVGMngr.mObjSVGDraw.mObjDataProc
        'Set m_DataProcObject = New clsDataProc
           
        Call m_DataProcObject.LoadDataLayer(colGroups)
        
        'Debug.Print "** END loadDataLayer - " & Now
        ''''''''''''''''''''''''
    End If


    Exit Function
    
ErrorHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "clsSVGComm.createDataLayer", "srvSVGComm", vbCritical)
'Resume
End Function

Public Function saveData()
    ' save age path
    Call m_TeethObject.setTeethPath(m_PatientID)
    
    'Remove this
    MsgBxARMd "This is NOT an error: (clsSVGComm.saveData) clsSVGComm is saving the datalayer."
    
    ' sava data layer
    m_DataObject.SaveGroups
    
End Function

Public Function LoadData(ViewMode As eViewMode)

    Call m_TeethObject.LoadTeeth(ViewMode)
    
End Function

Public Function SvgClick(nTooth As Integer, evt As Object)

    If mbReadOnly Then Exit Function
     
    If m_TeethObject.toothMode = eBigTooth Then
        Call m_ToothObject.SvgMouseClick(CStr(nTooth), evt)
    Else
        Call m_TeethObject.SvgClickTeeth(nTooth, evt)
    End If
    
End Function
Public Function SvgMouseClick(key As String, evt As Object)

  If mbReadOnly Then Exit Function
   
  If Not IsNumeric(key) And Len(key) > 0 And InStr(key, "RO") < 1 And InStr(key, "CR") < 1 Then
        Call m_ToothObject.SvgMouseClick(key, evt)
  End If
End Function
Public Function SvgMouseOver(key As String, evt As Object)

     If mbReadOnly Then Exit Function
      
    If Not IsNumeric(key) And Len(key) > 0 And InStr(key, "RO") < 1 And InStr(key, "CR") < 1 Then
        Call m_ToothObject.SvgMouseOver(key, evt)
     End If
        'Debug.Print "mouseover " & key
End Function
Public Function SvgMouseOut(key As String, evt As Object)

    If mbReadOnly Then Exit Function
       
    If Not IsNumeric(key) And Len(key) > 0 And InStr(key, "RO") < 1 And InStr(key, "CR") < 1 Then
        Call m_ToothObject.SvgMouseOut(key, evt)
        m_strLastKey = key
    ElseIf IsNumeric(key) And m_ToothObject.mbMouseDown Then
        m_ToothObject.mbMouseDown = False
        Call m_ToothObject.SvgMouseUp(m_strLastKey, evt)
    End If
    'Debug.Print "mouseout " & key
End Function
Public Function SvgMouseUp(key As String, evt As Object)
Dim oSvg As Object
    
   If mbReadOnly Then Exit Function
    
    If Not IsNumeric(key) And Len(key) > 0 And InStr(key, "RO") < 1 And InStr(key, "CR") < 1 Then
        Call m_ToothObject.SvgMouseUp(key, evt)
         'Debug.Print "mouseup " & key
    End If
        
                 
End Function
Public Function SvgMouseDown(key As String, evt As Object)

     If mbReadOnly Then Exit Function
     
    If Not IsNumeric(key) And Len(key) > 0 And InStr(key, "RO") < 1 And InStr(key, "CR") < 1 Then
        Call m_ToothObject.SvgMouseDown(key, evt)
    End If
        'Debug.Print "mousedown " & key
End Function
Public Function ShiftKey(bKeyDown As Boolean)
    'bShift = bKeyDown
    
     If mbReadOnly Then Exit Function
     
    Call m_TeethObject.ShiftKey(bKeyDown)
    
End Function

Public Function CtrlKey(bKeyDown As Boolean)
    'bCtrl = bKeyDown
    
     If mbReadOnly Then Exit Function
     
    Call m_TeethObject.CtrlKey(bKeyDown)
    
End Function
Public Function createDentalSVG(bIsDx As Boolean) As String
    Dim s As String, sTeethSVG As String, sHeader As String, sFooter As String
    'dim sAgePath As String
    Dim oPositions As colPositions
    Dim oPos As clsPosition
    Dim iTooth As Integer
    
    Dim nFilePos As Long
    
'    Dim aAgePath() As String
'    Dim varTooth As Variant
    
    On Error GoTo ErrorHandler
    
    Debug.Assert False
    
    ' if core has not started the server
    'Call createDataLayer
    
    ' read template file
    'SF 9-16-03 SCR 40464
    'sTeethSVG = gobjShared.ImagesPath & "\" & TEETH_SVG_FILE_REP
    sTeethSVG = Replace(gobjShared.ImagesPath & "\" & TEETH_SVG_FILE_REP, "\\", "\")
     
    Open sTeethSVG For Input As #1
    s = Input$(LOF(1), #1)
    sFooter = "</svg>"
    nFilePos = InStr(s, sFooter)
    sHeader = Left(s, nFilePos - 1)
    Close #1
    
    'Debug.Print sHeader
    'Debug.Print sFooter
       
    If m_TeethObject.Positions.Count = 0 Then
        'sAgePath = m_TeethObject.LoadPositions(m_PatientID, m_Age)
        m_TeethObject.LoadPositions m_PatientID, m_Age
    End If
    Set oPositions = m_TeethObject.Positions
            
    mObjSVGMngr.Age = m_Age
    'mObjSVGMngr.teethAgePath = sAgePath
    Set mObjSVGMngr.Positions = oPositions
      
    s = sHeader
    
    For Each oPos In oPositions
        iTooth = oPos.ToothNum
                
        mObjSVGMngr.drawSVG (iTooth)
        
        If bIsDx Then
            s = s & mObjSVGMngr.getsvgDxViewbyTooth(iTooth)
        Else
            s = s & mObjSVGMngr.getsvgTxViewbyTooth(iTooth)
        End If
        
    Next
    
'    aAgePath = Split(sAgePath, ",")
'    For Each varTooth In aAgePath
'        iTooth = CInt(varTooth)
'
'        mObjSVGMngr.drawSVG (iTooth)
'        If bIsDx Then
'            s = s & mObjSVGMngr.getsvgDxViewbyTooth(iTooth)
'        Else
'            s = s & mObjSVGMngr.getsvgTxViewbyTooth(iTooth)
'        End If
'
'    Next
    s = s & sFooter
    
    createDentalSVG = s
    'Debug.Print s
    
'    Open "c:\temp\test.svg" For Output As #1
'    Print #1, s
'    Close #1
    
    Exit Function
    
ErrorHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "clsSVGComm.createDentalSVG", "srvSVGComm", vbCritical)
    
End Function
Friend Sub TerminateObjects()

'SF 5-17-02
    'Terminate the instance of the SVGManager
    moSVGMngrSingleton.TerminateSF
    
    Set m_TeethObject = Nothing
    Set m_ToothObject = Nothing
    Set m_DataObject = Nothing
    Set m_DataProcObject = Nothing
    Set mObjSVGMngr = Nothing
    Set mObjPSR = Nothing
    
End Sub
Public Sub InitializeObjects()
'SF 5-17-02
   Set moSVGMngrSingleton = New CHCSII_Dental_SVG_Manager.SF
   'MsgBxARMd "SVGComm is creating an instance of manager"
    Set mObjSVGMngr = moSVGMngrSingleton.Singleton  'Set SVG Manager to one instance
End Sub
Public Function TxChangedStatus(Tx As CHCSII_Dental_DXTX.clsTx)

    'Tell GC that the status of a tx changed.
    '48054 SF 1-23-04 - Use comm to relay the message to GC
    RaiseEvent TxChangedStatus(Tx)

End Function

Private Sub Class_Initialize()
    Set gobjShared = New CWShared
    Set mObjSVGMngr = moSVGMngrSingleton.Singleton  'Set SVG Manager to one instance

    Set m_TeethObject = New clsTeeth
    Set m_ToothObject = New clsTooth
    Set mObjPSR = New clsPSR
    m_DentalEncounterID = 0
    '--- SCR 28304 RWalker 10/10/02 Instantiate ARM / Init transactions
    Call ARM_InitAllTransactionsForDental_SrvSVGComm
End Sub

Public Sub drawMenuItem(toothID As Integer, SurfaceSet As String, MenuID As String, Optional objDLDx As CHCSII_Dental_DXTX.clsDx, Optional objDLTx As CHCSII_Dental_DXTX.clsTx)
    mObjSVGMngr.drawSVG toothID, SurfaceSet, MenuID, objDLDx, objDLTx
End Sub
Public Function updateSVG(objDLDx As CHCSII_Dental_DXTX.clsDx, objDLTx As CHCSII_Dental_DXTX.clsTx, Optional bDoUpdateDraw As Boolean = True, Optional IsCompleteDenture As Boolean = False) As Boolean
Dim iToothID As Integer
Dim iCurrentToothId As Integer
Dim i As Integer
Dim strArrTeeth() As String
Dim GCTooth As clsGCTooth

    m_ToothObject.mbIsProcessingMenuSelection = True
   
    

    If Not objDLDx Is Nothing Then
        'Get the tooth associated with the Dx
        Set GCTooth = mObjSVGMngr.mColTeeth(CStr(objDLDx.SurfaceSets(1).Teeth(1).toothID))
'        If GCTooth.bToothVisible = False Then
'            MsgBox ("Cannot delete dx/tx because the current tooth is not active.")
'            updateSVG = False
'            Exit Function
'        End If
        
        
        If objDLDx.StatusID = dxNonActiveDeleted Then
            If objDLDx.SurfaceSets(1).SymbolID = 151 And IsCompleteDenture = False Then 'missing logic
                'TeethObject.ClickTeeth iToothID
                m_ToothObject.RemoveItemsByDate (drChartedIncorrectly)  'Yank the items after tooth comes back from a missing state
                
                '37476
                'Update the tooth for the current view
                UpdateTooth objDLDx.SurfaceSets(1).Teeth(1).toothID, m_ToothObject.DxTxViewMode
'                If m_ToothObject.DxTxViewMode = eDxView Then
'                    UpdateTooth objDLDx.SurfaceSets(1).Teeth(1).toothID, True
'                Else
'                    UpdateTooth objDLDx.SurfaceSets(1).Teeth(1).toothID, False
'                End If
            End If
        End If
    Else
        'Get the tooth associated with the Tx
        Set GCTooth = mObjSVGMngr.mColTeeth(CStr(objDLTx.SurfaceSets(1).Teeth(1).toothID))
'        If GCTooth.bToothVisible = False Then
'           MsgBox ("Cannot delete dx/tx because the current tooth is not active.")
'           updateSVG = False
'           Exit Function
'        End If
    End If
    
   mObjSVGMngr.updateSVG objDLDx, objDLTx
    
    If bDoUpdateDraw Then
        If Not objDLDx Is Nothing Then
        
            If objDLDx.SurfaceSets(1).SymbolID = 151 And objDLDx.StatusID = dxNonActiveDeleted Then 'Missing state being deleted
                MsgBxARMd "The previous Dx's and Tx's of this missing tooth have been removed and must be re-charted.", vbInformation, "Graphic Charting"
            End If
            If objDLDx.SurfaceSets(1).Teeth(1).ToothName = "Max" Or objDLDx.SurfaceSets(1).Teeth(1).ToothName = "Mand" Then
                strArrTeeth = Split(objDLDx.SurfaceSets(1).SurfaceSet, ",")
            Else
                strArrTeeth = Split(objDLDx.SurfaceSets(1).Teeth(1).toothID, ",")
            End If
        
        Else
        
            If objDLTx.SurfaceSets(1).Teeth(1).ToothName = "Max" Or objDLTx.SurfaceSets(1).Teeth(1).ToothName = "Mand" Then
                strArrTeeth = Split(objDLTx.SurfaceSets(1).SurfaceSet, ",")
            Else
                strArrTeeth = Split(objDLTx.SurfaceSets(1).Teeth(1).toothID, ",")
            End If
        
        End If
            
        For i = 0 To UBound(strArrTeeth)
            
            m_ToothObject.CurrentViewTooth = CInt(strArrTeeth(i))
            iCurrentToothId = m_ToothObject.CurrentViewTooth
            iToothID = iCurrentToothId
            
            '37476
            'Update the current view last
            If m_ToothObject.DxTxViewMode = eDxView Then
                 'If Not objDLDx Is Nothing Then
                    UpdateTooth iToothID, eDx603aView
                    UpdateTooth iToothID, eTxView
                    UpdateTooth iToothID, eDxView
                 'Else
                 '   UpdateTooth iToothID, False
                 '   UpdateTooth iToothID, True
                 'End If
            ElseIf m_ToothObject.DxTxViewMode = eTxView Then
                 'If Not objDLDx Is Nothing Then
                    UpdateTooth iToothID, eDx603aView
                    UpdateTooth iToothID, eDxView
                    UpdateTooth iToothID, eTxView
                 'Else
                 '   UpdateTooth iToothID, True
                 '   UpdateTooth iToothID, False
                 'End If
            ElseIf m_ToothObject.DxTxViewMode = eDx603aView Then
                UpdateTooth iToothID, eDxView
                UpdateTooth iToothID, eTxView
                UpdateTooth iToothID, eDx603aView
            End If
            
'            If m_ToothObject.DxTxMode = eDxView Then
'                 If Not objDLDx Is Nothing Then
'                    UpdateTooth iToothID, False
'                    UpdateTooth iToothID, True
'                 Else
'                    UpdateTooth iToothID, False
'                    UpdateTooth iToothID, True
'
'                 End If
'            Else
'                 If Not objDLDx Is Nothing Then
'                    UpdateTooth iToothID, True
'                    UpdateTooth iToothID, False
'                 Else
'                    UpdateTooth iToothID, True
'                    UpdateTooth iToothID, False
'                 End If
'            End If
            
            '37476
            'Update the current tooth
            UpdateTooth iCurrentToothId, m_ToothObject.DxTxViewMode
'            If m_ToothObject.DxTxMode = eDxView Then
'               UpdateTooth iCurrentToothId, True
'            Else
'               UpdateTooth iCurrentToothId, False
'            End If
            
            m_ToothObject.mbIsProcessingMenuSelection = False
        Next i
        
        If m_ToothObject.BTFMMode = eFullMouth Then
            m_TeethObject.clearSelectedTeeth
        End If
    
    End If
    
    updateSVG = True

End Function
Public Sub UpdateTooth(iToothID As Integer, ViewMode As eViewMode, Optional bClearCurrentTooth As Boolean = False)


    If Not bClearCurrentTooth Then
        '37476
        m_TeethObject.DxTxViewMode = ViewMode
        'm_TeethObject.SetDXTXMode bIsDxView
        
        m_ToothObject.UpdateTooth (iToothID), ViewMode
        m_TeethObject.UpdateTeeth (iToothID)
    End If

    If m_TeethObject.toothMode = eBigTooth Then
        m_TeethObject.DisplaySingleTooth iToothID, mObjSVGMngr.mColTeeth(CStr(iToothID)).ToothVisible
    Else
        m_TeethObject.DisplaySingleTooth iToothID, mObjSVGMngr.mColTeeth(CStr(iToothID)).ToothVisible
        m_TeethObject.toothMode = eFullMouth
    End If

End Sub
Public Sub UpdateTeeth(ViewMode As eViewMode)

Dim oPos As clsPosition
    
    '37476
    m_TeethObject.DxTxViewMode = ViewMode
    
    For Each oPos In m_TeethObject.Positions
        Call m_TeethObject.UpdateTeeth(CInt(oPos.ToothNum), False)
    Next

'Dim i As Integer
'Dim strArrTeethPath() As String
'
'    strArrTeethPath = Split(m_TeethObject.AgePath, ",")
'    m_TeethObject.SetDXTXMode bIsDxView
'
'    For i = 0 To UBound(strArrTeethPath)
'       Call m_TeethObject.UpdateTeeth(CInt(strArrTeethPath(i)), False)
'    Next i

End Sub
Public Sub refreshTooth(iToothID As Integer, ViewMode As eViewMode)
    
    '37476
    m_TeethObject.DxTxViewMode = ViewMode
    'm_TeethObject.SetDXTXMode bIsDxView
    
    If m_TeethObject.toothMode Then
        m_TeethObject.DisplaySingleTooth (iToothID)
    Else
        m_TeethObject.DisplaySingleTooth (iToothID)
        m_TeethObject.toothMode = False
    End If
    
End Sub

Private Sub Class_Terminate()

'singletons don't terminate correctly
'    TerminateObjects

End Sub
