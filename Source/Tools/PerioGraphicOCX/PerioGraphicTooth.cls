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
Attribute VB_Name = "PerioGraphicTooth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const bHWCCSupportSupernumerary As Boolean = True

Public mi_Offset As Integer                 'Default to 0
Public mstr_SecondaryToothName As String    'Set when this tooth is created
Public mstr_PrimaryToothName As String      'Set when this tooth is created
Public mstr_ToothName As String             'Default to secondary tooth
Public mo_SVGDOMDoc As Object               'SVG DOM Doc Object
Public mo_SVGDOMWin As Object
Public mi_ToothType As Long                 'Default to perSecondary
Public mb_IsPontic As Boolean               'Default to False
Public mb_IsPrimTooth As Boolean            'Flag if there is PrimaryTooth, Default to true
Public mb_IsImplant As Boolean
Public mb_IsDirty As Boolean                'Dirty Flag
Public mi_Position As Integer               'Position of the tooth
Public m_PMeas_Bleed As PerioMeasBleed      'Bleeding
Public m_PMeas_Plaq As PerioMeasPlaq        'Plaque
Public m_PMeas_Supp As PerioMeasSupp        'Supp
Public m_PMeas_Mob As PerioMeasMob          'Mobility
Public m_PMeas_Furc As PerioMeasFurc        'Furcation
Public m_PMeas_BL As PerioMeasBL            'Bone Level
Public m_PMeas_PD As PerioMeasPD            'Probing Depth
Public m_PMeas_CAL As PerioMeasCAL          'CAL
Public m_PMeas_MGJ As PerioMeasMGJ          'MGJ
Public m_PMeas_GM As PerioMeasGM            'GM
Private mcol_GrphElem As Collection         'Collection to hold Graphic Element
Private mo_PMGrphElem As PerioMeasGrphElem  'PerioMesurment Graphical Element
Private mPCTL As PerioGraphicCtl            'Reference bk to the OCX
Private mstrSVGToothString As String
Private mdictSVG As Dictionary
Private mblnPrimaryGraphicLoaded As Boolean
Private mblnVisible As Boolean

Friend Property Get Visible() As Boolean
   Visible = mblnVisible
End Property
Friend Property Let Visible(blnVisible As Boolean)
   mblnVisible = blnVisible
End Property

Friend Property Get PrimaryGraphicLoaded() As Boolean
    PrimaryGraphicLoaded = mblnPrimaryGraphicLoaded
End Property

Friend Sub LoadPrimaryGraphic()
    Dim strElementId As String

    strElementId = mstr_PrimaryToothName
    SetElementClass strElementId, ""
    mblnPrimaryGraphicLoaded = True
End Sub

Public Property Get SVGSTring() As String
    SVGSTring = mstrSVGToothString
End Property
Public Property Let SVGSTring(strSVG As String)
    mstrSVGToothString = strSVG
End Property

Public Property Get MeasSVGString(strId As String) As String
    MeasSVGString = mdictSVG(strId)
End Property
Public Property Let MeasSVGString(strId As String, strSVG As String)
    mdictSVG(strId) = strSVG
End Property
'

Public Sub Draw()
  'Set the Show flag in each MeasGrphElem
  ResetDataValue
  ClearTooth
  SetTooth
  ClearGrphElem
  SetShowFlag
  If (mb_IsPontic = False) And (mi_ToothType <> perNoDraw) Then
    DrawGrphElem
  End If
End Sub

Public Sub ReDraw()
  If mb_IsDirty = False Then
    Exit Sub: End If
  ClearTooth
  SetTooth
  ClearGrphElem
  SetShowFlag
  DrawGrphElem
  mb_IsDirty = False
End Sub

Public Sub ResetDataValue()
  'Reset all parameters
  mi_Offset = 0
  mi_ToothType = perSecondary
  mb_IsPontic = False
  mb_IsImplant = False
  mb_IsDirty = True
  'Reset all PerioMeas object's data
  Dim PMGrphElem As PerioMeasGrphElem
  For Each PMGrphElem In mcol_GrphElem
    PMGrphElem.ResetDataValue
    DoEvents
  Next PMGrphElem
End Sub

Public Sub Init(SecondaryTooth As String, PrimaryTooth As String, TheOCX As PerioGraphicCtl, iPos As Integer)
  mstr_ToothName = SecondaryTooth      'Default the tooth name to Secondary Tooth
  mstr_SecondaryToothName = SecondaryTooth
  mstr_PrimaryToothName = PrimaryTooth
  If Not (PrimaryTooth = "") Then
    mb_IsPrimTooth = True               'There is no primary tooth pass in
  End If
  mi_Position = iPos                    'The Position of the tooth
  mb_IsDirty = True                     'Initial Dirty bit setting
  mblnPrimaryGraphicLoaded = False      'Primary graphics are not loaded
  Set mPCTL = TheOCX
  'Loop through the mcol_GrphElem
  Dim PMGrphElem As PerioMeasGrphElem
  For Each PMGrphElem In mcol_GrphElem
    PMGrphElem.Init Me
    DoEvents
  Next PMGrphElem
End Sub

Public Sub SetSVGDoc(SVGDoc As Object)
    ' This is called after SVG Doc is loaded
    Set mo_SVGDOMDoc = Nothing 'Release reference if any
    Set mo_SVGDOMDoc = SVGDoc  'Reset SVGDOMDoc
End Sub

Private Sub SetShowFlag()
  m_PMeas_Bleed.bShow = mPCTL.ShowBleeding
  m_PMeas_Plaq.bShow = mPCTL.ShowPlaque
  m_PMeas_Supp.bShow = mPCTL.ShowSuppuration
  m_PMeas_Mob.bShow = mPCTL.ShowMobility
  m_PMeas_Furc.bShow = mPCTL.ShowFurcation
  m_PMeas_BL.bShow = mPCTL.ShowBoneLevel
  m_PMeas_PD.bShow = mPCTL.ShowProbDepth  'PD
  m_PMeas_CAL.bShow = mPCTL.ShowCAL       'CAL
  m_PMeas_MGJ.bShow = mPCTL.ShowMGJ       'MGJ
  m_PMeas_GM.bShow = mPCTL.ShowGM         'GM

End Sub

Private Sub DrawGrphElem() 'Draw the graphic Element alone
  'Draw all the Graphic Elements
  'Loop through the mcol_GrphElem
  Dim PMGrphElem As PerioMeasGrphElem
  For Each PMGrphElem In mcol_GrphElem
    PMGrphElem.Draw
    DoEvents
  Next PMGrphElem
End Sub

Private Sub SetTooth()  'Set/Display the correct toof for Perio Charting
Static strSecondaryLabel As String

  Dim toofObj As Object
  Dim SVGElementId As String
  
  If mb_IsPontic Then    'Handle the pontic situation
    If mb_IsPrimTooth And _
      (mi_ToothType = perPrimary) Then 'Tooth type is primary tooth
      
      If Not mblnPrimaryGraphicLoaded Then
          Call LoadPrimaryGraphic
      End If
      
      'Turn on the Primary tooth Crown
      SVGElementId = mstr_PrimaryToothName & "-" & GE_Crown
      SetElementClass SVGElementId, "stOutline"
      'Turn off the root of the Primary tooth
      SVGElementId = mstr_PrimaryToothName & "-" & GE_Roots
      ClearElement SVGElementId
        'Set Label
        SVGElementId = mstr_PrimaryToothName & "-" & GE_Location
        DrawElement SVGElementId
    Else
      'Turn the tooth back on
      SVGElementId = mstr_SecondaryToothName & "-" & GE_Crown
      SetElementClass SVGElementId, "stOutline"
      'Turn off the root of the Secondary tooth
      SVGElementId = mstr_SecondaryToothName & "-" & GE_Roots
      ClearElement SVGElementId
        'Set Label
        SVGElementId = mstr_SecondaryToothName & "-" & GE_Location & "-" & GE_TSpan
        If Len(strSecondaryLabel) = 0 Then
            SVGElementId = mstr_SecondaryToothName & "-" & GE_Location
            DrawElement SVGElementId
        Else
            SVGElementId = mstr_SecondaryToothName & "-" & GE_Location & "-" & GE_TSpan
            SetText SVGElementId, strSecondaryLabel
        End If
    End If
    Exit Sub  'No need to do anymore, just display the Pontics
  End If
  
'Check tooth type, for perNoDraw tooth type, we just do nothing
  If Visible = False Then
      If mi_ToothType = perSecondary Then
          mstr_ToothName = mstr_SecondaryToothName
          'Set Label
          SVGElementId = mstr_SecondaryToothName & "-" & GE_Location & "-" & GE_TSpan
          If Len(strSecondaryLabel) = 0 Then
              SVGElementId = mstr_SecondaryToothName & "-" & GE_Location
              DrawElement SVGElementId
          Else
              SVGElementId = mstr_SecondaryToothName & "-" & GE_Location & "-" & GE_TSpan
              SetText SVGElementId, strSecondaryLabel
          End If
      
          If Not (mstr_PrimaryToothName = "") Then
            SVGElementId = mstr_PrimaryToothName & "-" & GE_Location
            ClearElement SVGElementId
          End If
   
      ElseIf mi_ToothType = perPrimary Then
         If Not mblnPrimaryGraphicLoaded Then
           Call LoadPrimaryGraphic
         End If
   
         mstr_ToothName = mstr_PrimaryToothName
         
         SVGElementId = mstr_PrimaryToothName & "-" & GE_Crown
         ClearElement SVGElementId
         SVGElementId = mstr_PrimaryToothName & "-" & GE_Roots
         ClearElement SVGElementId
         SVGElementId = mstr_PrimaryToothName & "-" & GE_Location
         DrawElement SVGElementId
         
      End If
  
  ElseIf mi_ToothType = perSecondary Then
    'Turn the tooth's root and crown back on
    SVGElementId = mstr_SecondaryToothName & "-" & GE_Crown
    SetElementClass SVGElementId, "stOutline"
    'Make sure the roots get turn back on
    SVGElementId = mstr_SecondaryToothName & "-" & GE_Roots
    SetElementClass SVGElementId, "stOutline"
    'Set Label
    SVGElementId = mstr_SecondaryToothName & "-" & GE_Location & "-" & GE_TSpan
    If Len(strSecondaryLabel) = 0 Then
        SVGElementId = mstr_SecondaryToothName & "-" & GE_Location
        DrawElement SVGElementId
    Else
        SVGElementId = mstr_SecondaryToothName & "-" & GE_Location & "-" & GE_TSpan
        SetText SVGElementId, strSecondaryLabel
    End If

    If Not (mstr_PrimaryToothName = "") Then
      SVGElementId = mstr_PrimaryToothName & "-" & GE_Location
      ClearElement SVGElementId
    End If
    'Set the toothName
    mstr_ToothName = mstr_SecondaryToothName
  ElseIf mb_IsPrimTooth And mi_ToothType = perPrimary Then
  
    If Not mblnPrimaryGraphicLoaded Then
        Call LoadPrimaryGraphic
    End If
    'Turn the tooth back on
    SVGElementId = mstr_PrimaryToothName & "-" & GE_Crown
    SetElementClass SVGElementId, "stOutline"
    'Make sure the roots get turn back on
    SVGElementId = mstr_PrimaryToothName & "-" & GE_Roots
    SetElementClass SVGElementId, "stOutline"
    
    'Label
    SVGElementId = mstr_SecondaryToothName & "-" & GE_Location
    ClearElement SVGElementId
    SVGElementId = mstr_PrimaryToothName & "-" & GE_Location
    DrawElement SVGElementId
    'Set the toothName
    mstr_ToothName = mstr_PrimaryToothName
  ElseIf mi_ToothType = perSuper Then
    'Turn the Secondary tooth back on
    SVGElementId = mstr_SecondaryToothName
    DrawElement SVGElementId
    'Set the toothName
    mstr_ToothName = mstr_SecondaryToothName
    'Turn off the Crown and Root
    SVGElementId = mstr_SecondaryToothName & "-" & GE_Roots
    ClearElement SVGElementId
    SVGElementId = mstr_SecondaryToothName & "-" & GE_Crown
    ClearElement SVGElementId
    'turn on the Supernumerary tooth
    SVGElementId = mstr_SecondaryToothName & "-" & GE_SNum
    DrawElement SVGElementId
    'change the tooth label to "S"
    SVGElementId = mstr_SecondaryToothName & "-" & GE_Location & "-" & GE_TSpan
    If Len(strSecondaryLabel) = 0 Then
        strSecondaryLabel = GetText(SVGElementId)
    End If
    SetText SVGElementId, "S"
  ElseIf mi_ToothType = perImplant Then
    mstr_ToothName = mstr_SecondaryToothName
    SVGElementId = mstr_SecondaryToothName & "-" & GE_Implant
    DrawElement SVGElementId
    'change the tooth label to "S"
    SVGElementId = mstr_SecondaryToothName & "-" & GE_Location & "-" & GE_TSpan
    
    'Label
    SVGElementId = mstr_SecondaryToothName & "-" & GE_Location
    DrawElement SVGElementId
  'ElseIf mi_ToothType = perNoDraw Then
  '  mstr_ToothName = mstr_SecondaryToothName
  '  SVGElementId = mstr_SecondaryToothName & "-" & GE_Location
  '  DrawElement SVGElementId
  End If
End Sub

Private Sub ClearTooth()
  'This will fullfill the toothtype perNoDraw
  Dim toofObj As Object
  Dim SVGElementId As String
  'Clear the Secondary tooth root & crown
  SVGElementId = mstr_SecondaryToothName & "-" & GE_Roots
  ClearElement SVGElementId
  
  Set toofObj = Nothing
  SVGElementId = mstr_SecondaryToothName & "-" & GE_Crown
  ClearElement SVGElementId
  
  Set toofObj = Nothing
  SVGElementId = mstr_SecondaryToothName & "-" & GE_SNum
  ClearElement SVGElementId
  
  Set toofObj = Nothing
  SVGElementId = mstr_SecondaryToothName & "-" & GE_Implant
  ClearElement SVGElementId
  
  
  Set toofObj = Nothing
  
    'Clear the label
    SVGElementId = mstr_SecondaryToothName & "-" & GE_Location
    ClearElement SVGElementId
    Set toofObj = Nothing
    
  'Clear the Primary tooth if there is one
  If mb_IsPrimTooth Then
    SVGElementId = mstr_PrimaryToothName & "-" & GE_Roots
    ClearElement SVGElementId
    
    Set toofObj = Nothing
    SVGElementId = mstr_PrimaryToothName & "-" & GE_Crown
    ClearElement SVGElementId
    
    Set toofObj = Nothing
    SVGElementId = mstr_PrimaryToothName & "-" & GE_SNum
    ClearElement SVGElementId
    Set toofObj = Nothing
    
    Set toofObj = Nothing
    SVGElementId = mstr_SecondaryToothName & "-" & GE_Implant
    ClearElement SVGElementId
    
    SVGElementId = mstr_PrimaryToothName & "-" & GE_Location
    ClearElement SVGElementId
    
    Set toofObj = Nothing
    
  End If
End Sub

Private Sub ClearGrphElem()
  'Loop through the mcol_GrphElem
  Dim PMGrphElem As PerioMeasGrphElem
  For Each PMGrphElem In mcol_GrphElem
    PMGrphElem.Clear
    DoEvents
  Next PMGrphElem
End Sub

Private Sub Class_Initialize()
  mi_Offset = 0
  Visible = True
  mi_ToothType = perSecondary
  mb_IsPontic = False
  mb_IsImplant = False
  mb_IsPrimTooth = False
  Set mo_SVGDOMDoc = Nothing
  Set mo_SVGDOMWin = Nothing
  
  'Create a brunch of Measurement objects here and then add to collection
  Set mcol_GrphElem = New Collection     'Create a new Collection
  Set m_PMeas_Bleed = New PerioMeasBleed
  mcol_GrphElem.Add m_PMeas_Bleed, PMGE_Bleed
  Set m_PMeas_Plaq = New PerioMeasPlaq
  mcol_GrphElem.Add m_PMeas_Plaq, PMGE_Plaq, after:=mcol_GrphElem.Count
  Set m_PMeas_Supp = New PerioMeasSupp
  mcol_GrphElem.Add m_PMeas_Supp, PMGE_Supp, after:=mcol_GrphElem.Count
  Set m_PMeas_Mob = New PerioMeasMob
  mcol_GrphElem.Add m_PMeas_Mob, PMGE_Mob, after:=mcol_GrphElem.Count
  Set m_PMeas_Furc = New PerioMeasFurc
  mcol_GrphElem.Add m_PMeas_Furc, PMGE_Furc, after:=mcol_GrphElem.Count
  Set m_PMeas_BL = New PerioMeasBL
  mcol_GrphElem.Add m_PMeas_BL, PMGE_BL, after:=mcol_GrphElem.Count
  Set m_PMeas_CAL = New PerioMeasCAL
  mcol_GrphElem.Add m_PMeas_CAL, PMGE_CAL, after:=mcol_GrphElem.Count
  Set m_PMeas_MGJ = New PerioMeasMGJ
  mcol_GrphElem.Add m_PMeas_MGJ, PMGE_MGJ, after:=mcol_GrphElem.Count
  Set m_PMeas_GM = New PerioMeasGM
  mcol_GrphElem.Add m_PMeas_GM, PMGE_GM, after:=mcol_GrphElem.Count
  Set m_PMeas_PD = New PerioMeasPD
  mcol_GrphElem.Add m_PMeas_PD, PMGE_PD, after:=mcol_GrphElem.Count
  
  Set mdictSVG = New Dictionary

End Sub

Private Sub Class_Terminate()
  'Release all references
  'Loop through the mcol_GrphElem
  
  'For Each mo_PMGrphElem In mcol_GrphElem
  '  Set mo_PMGrphElem = Nothing
  'Next mo_PMGrphElem
  
  If Not mcol_GrphElem Is Nothing Then
    mcol_GrphElem.Remove PMGE_Bleed
    mcol_GrphElem.Remove PMGE_Plaq
    mcol_GrphElem.Remove PMGE_Supp
    mcol_GrphElem.Remove PMGE_Mob
    mcol_GrphElem.Remove PMGE_Furc
    mcol_GrphElem.Remove PMGE_BL
    mcol_GrphElem.Remove PMGE_PD
    mcol_GrphElem.Remove PMGE_MGJ
    mcol_GrphElem.Remove PMGE_CAL
    mcol_GrphElem.Remove PMGE_GM
  End If
  
  Set mPCTL = Nothing
  Set mcol_GrphElem = Nothing
  Set mo_SVGDOMDoc = Nothing
  Set mo_SVGDOMWin = Nothing
End Sub

Private Function GetText(strElementId As String) As String

Dim objSVGObj As Object
Dim objElement As Object
Dim strSVG As String
Dim objSVGTooth As Object
Dim objRet As Object
Dim objText As Object

    GetSVGElementByID mo_SVGDOMDoc, strElementId, objSVGObj
    
    If objSVGObj Is Nothing Then 'need to load the element

        strSVG = MeasSVGString(strElementId)

        Set objElement = mo_SVGDOMWin.parseXML(CStr(strSVG), mo_SVGDOMDoc)


        Set objSVGTooth = mo_SVGDOMDoc.getElementById(mstr_ToothName)
        If Not objSVGTooth Is Nothing Then
            Set objSVGObj = objSVGTooth.appendChild(objElement.firstChild)
        End If
    End If
    
    Set objText = objSVGObj.getFirstChild()
    GetText = objText.getNodeValue

End Function


Private Sub SetText(strElementId As String, strText As String)

Dim objSVGObj As Object
Dim objElement As Object
Dim strSVG As String
Dim objSVGTooth As Object
Dim objRet As Object
Dim objText As Object

    GetSVGElementByID mo_SVGDOMDoc, strElementId, objSVGObj
    
    If objSVGObj Is Nothing Then 'need to load the element

        strSVG = MeasSVGString(strElementId)

        Set objElement = mo_SVGDOMWin.parseXML(CStr(strSVG), mo_SVGDOMDoc)


        Set objSVGTooth = mo_SVGDOMDoc.getElementById(mstr_ToothName)
        If Not objSVGTooth Is Nothing Then
            Set objSVGObj = objSVGTooth.appendChild(objElement.firstChild)
        End If
    End If
    
    'Set the tspan properties
    Call objSVGObj.setAttribute("fill-opacity", "1")
    Call objSVGObj.setAttribute("stroke-width", "none")
    
    'set the tspan value
    Set objText = objSVGObj.getFirstChild()
    Call objText.setNodeValue(CStr(strText))

End Sub

Private Sub SetElementClass(strElementId As String, strClass As String)
Dim objSVGObj As Object
Dim objElement As Object
Dim strSVG As String
Dim objSVGTooth As Object
Dim objRet As Object

    GetSVGElementByID mo_SVGDOMDoc, strElementId, objSVGObj
    
    If objSVGObj Is Nothing Then 'need to load the element

        strSVG = MeasSVGString(strElementId)

        Set objElement = mo_SVGDOMWin.parseXML(CStr(strSVG), mo_SVGDOMDoc)


        Set objSVGTooth = mo_SVGDOMDoc.getElementById(CStr(mstr_ToothName))
        If Not objSVGTooth Is Nothing Then
            Set objSVGObj = objSVGTooth.appendChild(objElement.firstChild)
        End If
    End If
    
    'Set drawing properties
    Call objSVGObj.setAttribute("class", CStr(strClass))
    
End Sub

Private Sub DrawElement(strElementId As String)
Dim objSVGObj As Object
Dim objElement As Object
Dim strSVG As String
Dim objSVGTooth As Object
Dim objRet As Object

    GetSVGElementByID mo_SVGDOMDoc, strElementId, objSVGObj
    
    If objSVGObj Is Nothing Then 'need to load the element

        strSVG = MeasSVGString(strElementId)

        Set objElement = mo_SVGDOMWin.parseXML(CStr(strSVG), mo_SVGDOMDoc)


        Set objSVGTooth = mo_SVGDOMDoc.getElementById(CStr(mstr_ToothName))
        If Not objSVGTooth Is Nothing Then
            Set objSVGObj = objSVGTooth.appendChild(objElement.firstChild)
        End If
    End If
    
    'Set drawing properties
    Call objSVGObj.setAttribute("class", "")
    Call objSVGObj.setAttribute("fill-opacity", "1")
    Call objSVGObj.setAttribute("stroke-width", "none")
    
End Sub

Private Sub ClearElement(strElementId As String)
  Dim objSVGObj As Object
  
  GetSVGElementByID mo_SVGDOMDoc, strElementId, objSVGObj
  
  If objSVGObj Is Nothing Then 'nothing to do
  Else
    Call objSVGObj.setAttribute("class", "stOpac")
  End If
  
End Sub

'Perform any cleanup needed
Public Sub Dispose()

Dim objElem As PerioMeasGrphElem
Dim i As Integer

    Set mo_SVGDOMDoc = Nothing
    Set mo_SVGDOMWin = Nothing
    
    For Each objElem In mcol_GrphElem
        objElem.Dispose
    Next objElem
    
    For i = mcol_GrphElem.Count To 1 Step -1
        mcol_GrphElem.Remove (i)
    Next i
    Set mcol_GrphElem = Nothing
    
    Set m_PMeas_Bleed = Nothing
    Set m_PMeas_Plaq = Nothing
    Set m_PMeas_Supp = Nothing
    Set m_PMeas_Mob = Nothing
    Set m_PMeas_Furc = Nothing
    Set m_PMeas_BL = Nothing
    Set m_PMeas_PD = Nothing
    Set m_PMeas_CAL = Nothing
    Set m_PMeas_MGJ = Nothing
    Set m_PMeas_GM = Nothing
    
    Set mPCTL = Nothing
    Set mdictSVG = Nothing
End Sub
