VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PerioMeasPD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements PerioMeasGrphElem

Public iMB As Integer     'Mesial Facial
Public iB As Integer      'Center Facial
Public iDB As Integer     'Distal Facial
Public iDL As Integer     'Distal Lingual
Public iL As Integer      'Center Lingual
Public iML As Integer     'Mesial Lingual

Public bShow As Boolean     'User set False (Default) = no show, True = show
Public MyTooth As PerioGraphicTooth  'Save a reference
Public b_IsDirty As Boolean 'Dirty Flag

Private PMesialGuideX As Integer
Private PCenterGuideX As Integer
Private PDistalGuideX As Integer
Private SMesialGuideX As Integer
Private SCenterGuideX As Integer
Private SDistalGuideX As Integer

Private bGuidesSet As Boolean
Private strMBId As String
Private strBId As String
Private strDBId As String
Private strDLId As String
Private strLId As String
Private strMLId As String
Private mstrSVGString As String

Private Sub PerioMeasGrphElem_Dispose()
    Set MyTooth = Nothing
End Sub

'
Private Property Let PerioMeasGrphElem_SVGString(RHS As String)
    mstrSVGString = RHS
End Property

Private Property Get PerioMeasGrphElem_SVGString() As String
    PerioMeasGrphElem_SVGString = mstrSVGString
End Property

Private Sub ScaleTop(iTo As Integer, iFrom As Integer)
  If iFrom = perUnknownValue Then
    iTo = 0
  Else
    iTo = 0 - ((iFrom * iTopScale) / DataRange)
  End If
End Sub

Private Sub ScaleBot(iTo As Integer, iFrom As Integer)
  If iFrom = perUnknownValue Then
    iTo = 0
  Else
    iTo = ((iFrom * iBotScale) / DataRange)
  End If
End Sub

Private Sub SetCSS(CSS As String, iPD As Integer)
  If iPD = 0 Then
    CSS = CSS_Opac
  ElseIf iPD >= ProbDepthThreshold Then
    CSS = CSS_PDRed
  Else
    CSS = CSS_PDBlue
  End If
End Sub

Private Sub Class_Initialize()
  PerioMeasGrphElem_ResetDataValue  'init all data values
  strMBId = ""
  strBId = ""
  strDBId = ""
  strDLId = ""
  strLId = ""
  strMLId = ""
  PMesialGuideX = 0
  PCenterGuideX = 0
  PDistalGuideX = 0
  SMesialGuideX = 0
  SCenterGuideX = 0
  SDistalGuideX = 0
 iMB = perUnknownValue   'Data values default to unknown
 iB = perUnknownValue
 iDB = perUnknownValue
 iDL = perUnknownValue
 iL = perUnknownValue
 iML = perUnknownValue
  bGuidesSet = False
End Sub

Private Sub Class_Terminate()
  Set MyTooth = Nothing  'Release reference
End Sub

Private Sub PerioMeasGrphElem_Clear()
   '1st time Clear get called, there is a SVGDoc availiable, we can
   'create a rect object
  If b_IsDirty = False Then
    Exit Sub: End If

  If Not bGuidesSet Then       'Set the Mesial Center and Distal Guides X coordinates
    If SMesialGuideX = 0 Or SCenterGuideX = 0 Or SDistalGuideX = 0 Then
      Dim GuideId As String
      GuideId = MyTooth.mstr_SecondaryToothName & "-" & GE_Mesi_Guide
      GetSVGLnElementX MyTooth.mo_SVGDOMDoc, GuideId, SMesialGuideX
      GuideId = MyTooth.mstr_SecondaryToothName & "-" & GE_Cntr_Guide
      GetSVGLnElementX MyTooth.mo_SVGDOMDoc, GuideId, SCenterGuideX
      GuideId = MyTooth.mstr_SecondaryToothName & "-" & GE_Dist_Guide
      GetSVGLnElementX MyTooth.mo_SVGDOMDoc, GuideId, SDistalGuideX
    End If
    
    If PMesialGuideX = 0 Or PCenterGuideX = 0 Or PDistalGuideX = 0 Then
        If MyTooth.mi_ToothType = perPrimary Then
            If (MyTooth.mi_Position >= 4 And MyTooth.mi_Position <= 13) Or _
              (MyTooth.mi_Position >= 20 And MyTooth.mi_Position <= 29) Then

                If Not MyTooth.PrimaryGraphicLoaded Then
                    Debug.Assert False
                    MyTooth.LoadPrimaryGraphic
                End If
                GuideId = MyTooth.mstr_PrimaryToothName & "-" & GE_Mesi_Guide
                GetSVGLnElementX MyTooth.mo_SVGDOMDoc, GuideId, PMesialGuideX
                GuideId = MyTooth.mstr_PrimaryToothName & "-" & GE_Cntr_Guide
                GetSVGLnElementX MyTooth.mo_SVGDOMDoc, GuideId, PCenterGuideX
                GuideId = MyTooth.mstr_PrimaryToothName & "-" & GE_Dist_Guide
                GetSVGLnElementX MyTooth.mo_SVGDOMDoc, GuideId, PDistalGuideX
            End If
        End If
    End If
    
    If PMesialGuideX = 0 Or PCenterGuideX = 0 Or PDistalGuideX = 0 Then
        Debug.Assert (MyTooth.mi_ToothType <> perPrimary)
        PMesialGuideX = SMesialGuideX
        PCenterGuideX = SCenterGuideX
        PDistalGuideX = SDistalGuideX
    End If
    
    bGuidesSet = True
   End If
   
   Dim objDynElem As Object
   If strMBId = "" Then
     strMBId = "Pos" & MyTooth.mi_Position & "-MB-PDepth"
     CreateSVGRectElement MyTooth.mo_SVGDOMDoc, strMBId, CSS_Opac, 0, 0, 0, 0
   Else
     GetSVGElementByID MyTooth.mo_SVGDOMDoc, strMBId, objDynElem
     Call objDynElem.setAttribute("class", CStr(CSS_Opac))
     Set objDynElem = Nothing
   End If
   If strBId = "" Then
     strBId = "Pos" & MyTooth.mi_Position & "-B-PDepth"
     CreateSVGRectElement MyTooth.mo_SVGDOMDoc, strBId, CSS_Opac, 0, 0, 0, 0
   Else
     GetSVGElementByID MyTooth.mo_SVGDOMDoc, strBId, objDynElem
     Call objDynElem.setAttribute("class", CStr(CSS_Opac))
     Set objDynElem = Nothing
   End If
   If strDBId = "" Then
     strDBId = "Pos" & MyTooth.mi_Position & "-DB-PDepth"
     CreateSVGRectElement MyTooth.mo_SVGDOMDoc, strDBId, CSS_Opac, 0, 0, 0, 0
   Else
     GetSVGElementByID MyTooth.mo_SVGDOMDoc, strDBId, objDynElem
     Call objDynElem.setAttribute("class", CStr(CSS_Opac))
     Set objDynElem = Nothing
   End If
   If strDLId = "" Then
     strDLId = "Pos" & MyTooth.mi_Position & "-DL-PDepth"
     CreateSVGRectElement MyTooth.mo_SVGDOMDoc, strDLId, CSS_Opac, 0, 0, 0, 0
   Else
     GetSVGElementByID MyTooth.mo_SVGDOMDoc, strDLId, objDynElem
     Call objDynElem.setAttribute("class", CStr(CSS_Opac))
     Set objDynElem = Nothing
   End If
   If strLId = "" Then
     strLId = "Pos" & MyTooth.mi_Position & "-L-PDepth"
     CreateSVGRectElement MyTooth.mo_SVGDOMDoc, strLId, CSS_Opac, 0, 0, 0, 0
   Else
     GetSVGElementByID MyTooth.mo_SVGDOMDoc, strLId, objDynElem
     Call objDynElem.setAttribute("class", CStr(CSS_Opac))
     Set objDynElem = Nothing
   End If
   If strMLId = "" Then
     strMLId = "Pos" & MyTooth.mi_Position & "-ML-PDepth"
     CreateSVGRectElement MyTooth.mo_SVGDOMDoc, strMLId, CSS_Opac, 0, 0, 0, 0
   Else
     GetSVGElementByID MyTooth.mo_SVGDOMDoc, strMLId, objDynElem
     Call objDynElem.setAttribute("class", CStr(CSS_Opac))
     Set objDynElem = Nothing
   End If
End Sub

Private Sub PerioMeasGrphElem_Draw()
  If bShow = False Then
   Exit Sub: End If
  
  If b_IsDirty = False Then
    Exit Sub: End If
   b_IsDirty = False       'Not dirty anymore
   
   Dim iFaceMesi As Integer, iFaceDist As Integer, iFaceCntr As Integer
   Dim iLingMesi As Integer, iLingDist As Integer, iLingCntr As Integer
   Dim MesiX As Integer, CntrX As Integer, DistX As Integer
   Dim iGM_MB As Integer, iGM_DB As Integer, iGM_B As Integer
   Dim iGM_ML As Integer, iGM_DL As Integer, iGM_L As Integer
   Dim strCSS As String
   
  If MyTooth.mi_ToothType = perSecondary Then
    MesiX = SMesialGuideX
    CntrX = SCenterGuideX
    DistX = SDistalGuideX
  Else
    MesiX = PMesialGuideX
    CntrX = PCenterGuideX
    DistX = PDistalGuideX
  End If
   
  Dim objDynElem As Object
  Dim X As Integer, Y As Integer, Width As Integer, Height As Integer
  If MyTooth.mi_Position <= 16 Then  'The top teeth, max arch
'
'  iTopScale = (iYT3 - iYT2)
'  iBotScale = (iYT2 - iYT3)
'
    ScaleTop iFaceMesi, iMB
    ScaleTop iFaceDist, iDB
    ScaleTop iFaceCntr, iB
    ScaleTop iLingMesi, iML
    ScaleTop iLingDist, iDL
    ScaleTop iLingCntr, iL
    
    ScaleTop iGM_MB, MyTooth.m_PMeas_GM.iMB
    ScaleTop iGM_B, MyTooth.m_PMeas_GM.iB
    ScaleTop iGM_DB, MyTooth.m_PMeas_GM.iDB
    ScaleTop iGM_ML, MyTooth.m_PMeas_GM.iML
    ScaleTop iGM_L, MyTooth.m_PMeas_GM.iL
    ScaleTop iGM_DL, MyTooth.m_PMeas_GM.iDL
    
     'Do the Facial Y first
     X = MesiX - BarChartWidth
     Y = iYT2
     Width = BarChartWidth * 2
     
     SetCSS strCSS, iMB
     SetSVGRectElement MyTooth.mo_SVGDOMDoc, strMBId, strCSS, X, Y - iFaceMesi - iGM_MB, Width, iFaceMesi
     
     X = CntrX - BarChartWidth
     SetCSS strCSS, iB
     SetSVGRectElement MyTooth.mo_SVGDOMDoc, strBId, strCSS, X, Y - iFaceCntr - iGM_B, Width, iFaceCntr
     
     X = DistX - BarChartWidth
     SetCSS strCSS, iDB
     SetSVGRectElement MyTooth.mo_SVGDOMDoc, strDBId, strCSS, X, Y - iFaceDist - iGM_DB, Width, iFaceDist
     
     'Now do the lingual
     Y = iYT1
     X = MesiX - BarChartWidth
     SetCSS strCSS, iML
     SetSVGRectElement MyTooth.mo_SVGDOMDoc, strMLId, strCSS, X, Y - iLingMesi - iGM_ML, Width, iLingMesi
     
     X = CntrX - BarChartWidth
     SetCSS strCSS, iL
     SetSVGRectElement MyTooth.mo_SVGDOMDoc, strLId, strCSS, X, Y - iLingCntr - iGM_L, Width, iLingCntr
     
     X = DistX - BarChartWidth
     SetCSS strCSS, iDL
     SetSVGRectElement MyTooth.mo_SVGDOMDoc, strDLId, strCSS, X, Y - iLingDist - iGM_DL, Width, iLingDist
     
  Else 'The bottom teeth
  
    ScaleBot iFaceMesi, iMB
    ScaleBot iFaceDist, iDB
    ScaleBot iFaceCntr, iB
    ScaleBot iLingMesi, iML
    ScaleBot iLingDist, iDL
    ScaleBot iLingCntr, iL
 
    ScaleBot iGM_MB, MyTooth.m_PMeas_GM.iMB
    ScaleBot iGM_B, MyTooth.m_PMeas_GM.iB
    ScaleBot iGM_DB, MyTooth.m_PMeas_GM.iDB
    ScaleBot iGM_ML, MyTooth.m_PMeas_GM.iML
    ScaleBot iGM_L, MyTooth.m_PMeas_GM.iL
    ScaleBot iGM_DL, MyTooth.m_PMeas_GM.iDL
      
      'Do the Facial Y first
     X = MesiX - BarChartWidth
     Y = iYB2
     Width = BarChartWidth * 2
     
     SetCSS strCSS, iMB
     SetSVGRectElement MyTooth.mo_SVGDOMDoc, strMBId, strCSS, X, Y + iGM_MB, Width, iFaceMesi
     
     X = CntrX - BarChartWidth
     SetCSS strCSS, iB
     SetSVGRectElement MyTooth.mo_SVGDOMDoc, strBId, strCSS, X, Y + iGM_B, Width, iFaceCntr
     
     X = DistX - BarChartWidth
     SetCSS strCSS, iDB
     SetSVGRectElement MyTooth.mo_SVGDOMDoc, strDBId, strCSS, X, Y + iGM_DB, Width, iFaceDist
     
     'Now do the lingual
     Y = iYB1
     X = MesiX - BarChartWidth
     SetCSS strCSS, iML
     SetSVGRectElement MyTooth.mo_SVGDOMDoc, strMLId, strCSS, X, Y + iGM_ML, Width, iLingMesi
     
     X = CntrX - BarChartWidth
     SetCSS strCSS, iL
     SetSVGRectElement MyTooth.mo_SVGDOMDoc, strLId, strCSS, X, Y + iGM_L, Width, iLingCntr
     
     X = DistX - BarChartWidth
     SetCSS strCSS, iDL
     SetSVGRectElement MyTooth.mo_SVGDOMDoc, strDLId, strCSS, X, Y + iGM_DL, Width, iLingDist
  End If
End Sub

Private Sub PerioMeasGrphElem_Init(ToothObj As PerioGraphicTooth)
   Set MyTooth = ToothObj
   b_IsDirty = True       'After reset always need redraw
End Sub

Public Sub SetShowFlag(Flag As Boolean)
  If bShow <> Flag Then
    If Not bShow Then
      b_IsDirty = True: End If
    bShow = Flag: End If
End Sub

Private Sub PerioMeasGrphElem_ResetDataValue()
 bShow = False
  strMBId = ""
  strBId = ""
  strDBId = ""
  strDLId = ""
  strLId = ""
  strMLId = ""
  PMesialGuideX = 0
  PCenterGuideX = 0
  PDistalGuideX = 0
  SMesialGuideX = 0
  SCenterGuideX = 0
  SDistalGuideX = 0
  bGuidesSet = False
 iMB = perUnknownValue   'Data values default to unknown
 iB = perUnknownValue
 iDB = perUnknownValue
 iDL = perUnknownValue
 iL = perUnknownValue
 iML = perUnknownValue
 b_IsDirty = True       'After reset always need redraw
End Sub
