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
Attribute VB_Name = "clsPerioSVG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'This class is for taking care of the low level Perio drawings
'to the SVG file
Option Explicit

'Event
Event SVGDocLoaded()

'This is from Perio Business Layer (PerioBL) "enums.cls"
'Public Enum eConstant
'    perUnknownValue = -5765
'End Enum

Rem PerioGraphicErrorCode
Const ErrCd_clsPerioSVG_SVGDoc_Notloaded As Integer = 20001

Rem PerioGraphicErrorMsg
Const ErrMsg_clsPerioSVG_SVGDoc_Notloaded As String = "No SVG document loaded"

'Store the 7 allowable SVGDocs in Perio Charting
Private m_CurrentSVGWin As Object
Private m_CurrentSVGDoc As Object 'Could be one of the above doc
Private m_LastErrCd As Long
Private m_LastErrMsg As String
'

Public Function GetErrorCode() As Integer
  GetErrorCode = m_LastErrCd
End Function

Public Function GetErrorMsg() As String
  GetErrorMsg = m_LastErrMsg
End Function

Public Function GetSVGDoc(SVGDoc As Object) As Boolean
  
  If m_CurrentSVGDoc Is Nothing Then
    m_LastErrCd = ErrCd_clsPerioSVG_SVGDoc_Notloaded
    m_LastErrMsg = ErrMsg_clsPerioSVG_SVGDoc_Notloaded
    GetSVGDoc = False
    Exit Function
  End If

  Set SVGDoc = m_CurrentSVGDoc
  GetSVGDoc = True
End Function

Public Function GetSVGWin(SVGWin As Object) As Boolean
  If m_CurrentSVGWin Is Nothing Then
    m_LastErrCd = ErrCd_clsPerioSVG_SVGDoc_Notloaded
    m_LastErrMsg = ErrMsg_clsPerioSVG_SVGDoc_Notloaded
    GetSVGWin = False
    Exit Function
  End If
  
  Set SVGWin = m_CurrentSVGWin
  GetSVGWin = True
End Function

Public Function initPerioSVGDoc(SVGWin As Object)
   Set m_CurrentSVGWin = SVGWin.SVGWindow
   Set m_CurrentSVGDoc = SVGWin.SVGDoc
   If m_CurrentSVGWin Is Nothing Then
     MyDebugMsg ("Nothing for m_CurrentSVGWin")
   End If
   If m_CurrentSVGDoc Is Nothing Then
     MyDebugMsg ("Nothing for m_CurrentSVGDoc")
   End If
   MyDebugMsg ("Loading Perio SVGDoc")
   RaiseEvent SVGDocLoaded
End Function

Public Function readyToExit()
  MyDebugMsg ("readyToExit")
  InitVariables   're-init variables
  ClearDOMObj
End Function

Private Sub Class_Initialize()
  MyDebugMsg ("Class_Initialize")
  'InitVariables
  'ClearDOMObj
End Sub

Private Sub Class_Terminate()
  MyDebugMsg ("Class_Terminate")
  ClearDOMObj
End Sub

Private Sub ClearDOMObj()
     Set m_CurrentSVGWin = Nothing
     Set m_CurrentSVGDoc = Nothing
MyDebugMsg ("Clearing SVGWin and SVGDoc")
End Sub

Private Sub InitVariables()
  m_LastErrCd = 0
  m_LastErrMsg = ""
End Sub
