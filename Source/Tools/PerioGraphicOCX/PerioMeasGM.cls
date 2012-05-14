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
Attribute VB_Name = "PerioMeasGM"
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
'

Private Sub ScaleTop(iTo As Integer, iFrom As Integer)
  If iFrom = perUnknownValue Then
    iTo = perUnknownValue
  Else
    iTo = 0 - ((iFrom * iTopScale) / DataRange)
  End If
End Sub

Private Sub ScaleBot(iTo As Integer, iFrom As Integer)
  If iFrom = perUnknownValue Then
    iTo = perUnknownValue
  Else
    iTo = ((iFrom * iBotScale) / DataRange)
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
End Sub

Private Sub PerioMeasGrphElem_Dispose()
    Set MyTooth = Nothing
End Sub

Private Sub PerioMeasGrphElem_Draw()

  If bShow = False Then
   Exit Sub
  End If
  
  If b_IsDirty = False Then
    Exit Sub
  End If
  
   b_IsDirty = False       'Not dirty anymore
   
   Dim iFaceMesi As Integer, iFaceDist As Integer, iFaceCntr As Integer
   Dim iLingMesi As Integer, iLingDist As Integer, iLingCntr As Integer
   Dim MesiX As Integer, CntrX As Integer, DistX As Integer
   Dim StartIdx As Integer
   
  If MyTooth.mi_ToothType = perSecondary Then
    MesiX = SMesialGuideX
    CntrX = SCenterGuideX
    DistX = SDistalGuideX
  Else
    MesiX = PMesialGuideX
    CntrX = PCenterGuideX
    DistX = PDistalGuideX
  End If
  
  If MyTooth.mi_Position > 16 Then
    StartIdx = ((33 - MyTooth.mi_Position) - 1) * 3  'Man arch
  Else
    StartIdx = (MyTooth.mi_Position - 1) * 3        'Max arch
  End If
  
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
    
    If MyTooth.mi_Position >= 9 And MyTooth.mi_Position <= 24 Then
      MaxFacGM(StartIdx + 1) = iFaceMesi   'Mesi is left
      MaxFacGM(StartIdx + 3) = iFaceDist
      MaxLingGM(StartIdx + 1) = iLingMesi
      MaxLingGM(StartIdx + 3) = iLingDist
      maxhcoors_48pts(StartIdx + 1) = MesiX     'Mesial is on the left
      maxhcoors_48pts(StartIdx + 3) = DistX
    Else
      MaxFacGM(StartIdx + 1) = iFaceDist   'Dist is left
      MaxFacGM(StartIdx + 3) = iFaceMesi
      MaxLingGM(StartIdx + 1) = iLingDist
      MaxLingGM(StartIdx + 3) = iLingMesi
      maxhcoors_48pts(StartIdx + 1) = DistX     'Distal is on the left
      maxhcoors_48pts(StartIdx + 3) = MesiX
    End If
    MaxFacGM(StartIdx + 2) = iFaceCntr
    MaxLingGM(StartIdx + 2) = iLingCntr
    maxhcoors_48pts(StartIdx + 2) = CntrX
 
  Else 'The bottom teeth
  
    ScaleBot iFaceMesi, iMB
    ScaleBot iFaceDist, iDB
    ScaleBot iFaceCntr, iB
    ScaleBot iLingMesi, iML
    ScaleBot iLingDist, iDL
    ScaleBot iLingCntr, iL
 
    If MyTooth.mi_Position >= 9 And MyTooth.mi_Position <= 24 Then
      ManFacGM(StartIdx + 1) = iFaceMesi   'Mesi is left
      ManFacGM(StartIdx + 3) = iFaceDist
      ManLingGM(StartIdx + 1) = iLingMesi
      ManLingGM(StartIdx + 3) = iLingDist
      manhcoors_48pts(StartIdx + 1) = MesiX     'Mesial is on the left
      manhcoors_48pts(StartIdx + 3) = DistX
    Else                                    'Tooth 32 to 25
      ManFacGM(StartIdx + 1) = iFaceDist   'Dist is left
      ManFacGM(StartIdx + 3) = iFaceMesi
      ManLingGM(StartIdx + 1) = iLingDist
      ManLingGM(StartIdx + 3) = iLingMesi
      manhcoors_48pts(StartIdx + 1) = DistX     'Distal is on the left
      manhcoors_48pts(StartIdx + 3) = MesiX
    End If
    ManFacGM(StartIdx + 2) = iFaceCntr
    ManLingGM(StartIdx + 2) = iLingCntr
    manhcoors_48pts(StartIdx + 2) = CntrX
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

Private Property Let PerioMeasGrphElem_SVGString(RHS As String)
    mstrSVGString = RHS
End Property

Private Property Get PerioMeasGrphElem_SVGString() As String
    PerioMeasGrphElem_SVGString = mstrSVGString
End Property
