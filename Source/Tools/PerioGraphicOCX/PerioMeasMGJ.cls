VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PerioMeasMGJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements PerioMeasGrphElem

Public iB As Integer      'Center Facial
Public iL As Integer      'Center Lingual

Public bShow As Boolean     'User set False (Default) = no show, True = show
Public MyTooth As PerioGraphicTooth  'Save a reference
Public b_IsDirty As Boolean 'Dirty Flag

Private PCenterGuideX As Integer
Private SCenterGuideX As Integer

Private bGuidesSet As Boolean
Private strBId As String
Private strLId As String
'
Private mstrSVGString As String

Private Sub PerioMeasGrphElem_Dispose()
    Set MyTooth = Nothing
End Sub

Private Property Let PerioMeasGrphElem_SVGString(RHS As String)
    mstrSVGString = RHS
End Property

Private Property Get PerioMeasGrphElem_SVGString() As String
    PerioMeasGrphElem_SVGString = mstrSVGString
End Property

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
  strBId = ""
  strLId = ""
  PCenterGuideX = 0
  SCenterGuideX = 0
 iB = perUnknownValue
 iL = perUnknownValue
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
    If SCenterGuideX = 0 Then
      Dim GuideId As String
      GuideId = MyTooth.mstr_SecondaryToothName & "-" & GE_Cntr_Guide
      GetSVGLnElementX MyTooth.mo_SVGDOMDoc, GuideId, SCenterGuideX
    End If
    
    If PCenterGuideX = 0 Then
        If MyTooth.mi_ToothType = perPrimary Then
            If (MyTooth.mi_Position >= 4 And MyTooth.mi_Position <= 13) Or _
              (MyTooth.mi_Position >= 20 And MyTooth.mi_Position <= 29) Then
                If Not MyTooth.PrimaryGraphicLoaded Then
                    Debug.Assert False
                    MyTooth.LoadPrimaryGraphic
                End If
                GuideId = MyTooth.mstr_PrimaryToothName & "-" & GE_Cntr_Guide
                GetSVGLnElementX MyTooth.mo_SVGDOMDoc, GuideId, PCenterGuideX
            End If
        End If
    End If
    
    If PCenterGuideX = 0 Then
        Debug.Assert (MyTooth.mi_ToothType <> perPrimary)
        PCenterGuideX = SCenterGuideX
    End If
    
    bGuidesSet = True
   End If
End Sub

Private Sub PerioMeasGrphElem_Draw()

  If bShow = False Then
   Exit Sub: End If
  
  If b_IsDirty = False Then
    Exit Sub: End If
   b_IsDirty = False       'Not dirty anymore
   
   Dim iFaceCntr As Integer
   Dim iLingCntr As Integer
   Dim CntrX As Integer
   Dim StartIdx As Integer
   
  If MyTooth.mi_ToothType = perSecondary Then
    CntrX = SCenterGuideX
  Else
    CntrX = PCenterGuideX
  End If
  
  If MyTooth.mi_Position > 16 Then
    StartIdx = 33 - MyTooth.mi_Position   'Man arch
  Else
    StartIdx = MyTooth.mi_Position        'Max arch
  End If
  
  Dim X As Integer, Y As Integer, Width As Integer, Height As Integer
  If MyTooth.mi_Position <= 16 Then  'The top teeth, max arch
'
'  iTopScale = (iYT3 - iYT2)
'  iBotScale = (iYT2 - iYT3)
'
    ScaleTop iFaceCntr, iB
    
    MaxFacMGJ(StartIdx) = iFaceCntr
    maxhcoors_16pts(StartIdx) = CntrX
 
  Else 'The bottom teeth
    ScaleBot iFaceCntr, iB
    ScaleBot iLingCntr, iL
 
    ManFacMGJ(StartIdx) = iFaceCntr
    ManLingMGJ(StartIdx) = iLingCntr
    manhcoors_16pts(StartIdx) = CntrX
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
  strBId = ""
  strLId = ""
  PCenterGuideX = 0
  SCenterGuideX = 0
  bGuidesSet = False
 iB = perUnknownValue
 iL = perUnknownValue
 b_IsDirty = True       'After reset always need redraw
End Sub
