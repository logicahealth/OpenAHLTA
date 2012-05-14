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
Attribute VB_Name = "PerioMeasBL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements PerioMeasGrphElem

'Public Tooth As String
Public iMBL As Integer             'Mesial Bone Loss
Public iDBL As Integer             'Distal Bone Loss

Public bShow As Boolean     'User set False (Default) = no show, True = show
Public MyTooth As PerioGraphicTooth  'Save a reference
Public b_IsDirty As Boolean 'Dirty Flag
Private FacBLId As String
Private LingBLId As String
Private objBLRect As Object
'
Private mstrSVGString As String
'
Public Property Get SVGSTring() As String
    SVGSTring = PerioMeasGrphElem_SVGString
End Property
Public Property Let SVGSTring(strSVG As String)
    PerioMeasGrphElem_SVGString = strSVG
End Property

Private Sub PerioMeasGrphElem_Dispose()
    Set MyTooth = Nothing
End Sub

Private Property Let PerioMeasGrphElem_SVGString(RHS As String)
    mstrSVGString = RHS
End Property

Private Property Get PerioMeasGrphElem_SVGString() As String
    PerioMeasGrphElem_SVGString = mstrSVGString
End Property

Private Sub Class_Initialize()
 PerioMeasGrphElem_ResetDataValue  'init all data values
 FacBLId = ""
 LingBLId = ""
End Sub

Private Sub Class_Terminate()
  Set MyTooth = Nothing  'Release reference
End Sub

Private Sub PerioMeasGrphElem_Clear()
   '1st time Clear get called, there is a SVGDoc availiable, we can
   'create a rect object
  If b_IsDirty = False Then
    Exit Sub: End If

   Dim objDynElem As Object
   If FacBLId = "" Then
     FacBLId = "Pos" & MyTooth.mi_Position & "-Facial-BL"
     'CreateSVGRectElement MyTooth.mo_SVGDOMDoc, FacBLId, CSS_Opac, 0, 0, 0, 0
     CreateSVGPathElement MyTooth.mo_SVGDOMDoc, FacBLId, CSS_Opac, "M0,0 L0,10 5,5 Z"
   Else
     GetSVGElementByID MyTooth.mo_SVGDOMDoc, FacBLId, objDynElem
     Call objDynElem.setAttribute("class", CStr(CSS_Opac))
     Set objDynElem = Nothing
   End If
   If LingBLId = "" Then
     LingBLId = "Pos" & MyTooth.mi_Position & "-Lingual-BL"
     'CreateSVGRectElement MyTooth.mo_SVGDOMDoc, LingBLId, CSS_Opac, 0, 0, 0, 0
     CreateSVGPathElement MyTooth.mo_SVGDOMDoc, LingBLId, CSS_Opac, "M0,0 L0,10 5,5 Z"
   Else
     GetSVGElementByID MyTooth.mo_SVGDOMDoc, LingBLId, objDynElem
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
   
   Dim iLeft As Integer, iRight As Integer
   
  If MyTooth.mi_Position >= 9 And MyTooth.mi_Position <= 24 Then
    iLeft = iMBL: iRight = iDBL
  Else
    iLeft = iDBL: iRight = iMBL
  End If
  
  If iLeft = perUnknownValue Then
    iLeft = 1: End If
    
  If iRight = perUnknownValue Then
    iRight = 1: End If
  
  Dim objDynElem As Object
  Dim dAttrib As String
  Dim X As Integer, Y As Integer, Width As Integer
  If MyTooth.mi_Position <= 16 Then  'The top teeth, max arch
     iRight = iYT2 - iYT3 + ((iRight * iTopScale) / DataRange)
     iLeft = iYT2 - iYT3 + ((iLeft * iTopScale) / DataRange)
     
     X = iX(MyTooth.mi_Position - 1)
     Width = iX(MyTooth.mi_Position) - X
     'Do the Facial Y first
     Y = iYT3
  
     dAttrib = "M " & X & "," & Y & " "  'set X,Y
     dAttrib = dAttrib & " h" & Width    'draw horizontally
     dAttrib = dAttrib & " v" & iRight   'draw downward
     dAttrib = dAttrib & " L " & X & "," & (iLeft + iYT3) & " Z" 'line to and then close
     
     GetSVGElementByID MyTooth.mo_SVGDOMDoc, FacBLId, objDynElem
     Call objDynElem.setAttribute("class", CStr(CSS_Bone))
     Call objDynElem.setAttribute("d", CStr(dAttrib))
     Set objDynElem = Nothing
     'Now do the lingual, Height is the same
     Y = iYT1 + iTopScale
     
     dAttrib = "M " & X & "," & Y & " "  'set X,Y
     dAttrib = dAttrib & " h" & Width    'draw horizontally
     dAttrib = dAttrib & " v" & iRight   'draw downward
     dAttrib = dAttrib & " L " & X & "," & (iLeft + Y) & " Z" 'line to and then close
     
     GetSVGElementByID MyTooth.mo_SVGDOMDoc, LingBLId, objDynElem
     Call objDynElem.setAttribute("class", CStr(CSS_Bone))
     Call objDynElem.setAttribute("d", CStr(dAttrib))
     Set objDynElem = Nothing
  Else 'The bottom teeth
     Dim eqToothPos As Integer  'Equivalent Tooth position
     eqToothPos = 33 - MyTooth.mi_Position
     X = iX(eqToothPos - 1)
     Width = iX(eqToothPos) - X
     Y = iYB2 + iBotScale
     'Do the Facial Y first
     iRight = ((iRight * iBotScale) / DataRange)
     iLeft = ((iLeft * iBotScale) / DataRange)
     
     dAttrib = "M " & X & "," & Y & " "  'set X,Y
     dAttrib = dAttrib & " h" & Width    'draw horizontally 'h
     dAttrib = dAttrib & " V" & (iYB2 + iRight)  'draw upward
     dAttrib = dAttrib & " L " & X & "," & (iYB2 + iLeft) & " Z" 'line to and then close
     
     GetSVGElementByID MyTooth.mo_SVGDOMDoc, FacBLId, objDynElem
     Call objDynElem.setAttribute("class", CStr(CSS_Bone))
     Call objDynElem.setAttribute("d", CStr(dAttrib))
     Set objDynElem = Nothing
     'Now do the lingual, Height is the same
     Y = iYB1 + iBotScale
     
     dAttrib = "M " & X & "," & Y & " "  'set X,Y
     dAttrib = dAttrib & " h" & Width    'draw horizontally
     dAttrib = dAttrib & " V" & (iYB1 + iRight)  'draw upward
     dAttrib = dAttrib & " L " & X & "," & (iYB1 + iLeft) & " Z" 'line to and then close
     
     GetSVGElementByID MyTooth.mo_SVGDOMDoc, LingBLId, objDynElem
     Call objDynElem.setAttribute("class", CStr(CSS_Bone))
     Call objDynElem.setAttribute("d", CStr(dAttrib))
     Set objDynElem = Nothing
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
 FacBLId = ""
 LingBLId = ""
 iMBL = 1                   'Bone Level default value
 iDBL = 1
 b_IsDirty = True       'After reset always need redraw
End Sub

