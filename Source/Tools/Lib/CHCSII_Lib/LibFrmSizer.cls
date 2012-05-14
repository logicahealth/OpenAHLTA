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
Attribute VB_Name = "mclsLibFrmSizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''************************************************************************
'''Module:           mclsLibFrmSizer/LibFrmSizer.cls
'''
'''Description:      Object resizing the form and its controls
'''
'''Note:             If client form has tab you me need to call P_Resize
'''                  tab change event because some tab controls may change
'''                  positioning properties of the background controls
'''
'''Modules Required:
'''
'''
'''Author         :  Alexander Alberstein
'''Modified       :
'''               :
'''************************************************************************
Option Explicit
''''''''''''''' Constanst Section '''''''''''''''''''''''''''''''''''''''''

Private Const msMODULE_NAME   As String = "mclsLibFrmSizer."
Private Const mlRESIZE_INIT   As Long = -999999

''''''''''''''' Variables Section '''''''''''''''''''''''''''''''''''''''''

Private mbInitialized         As Boolean

Private mlScaleHeight         As Long
Private mlScaleWidth          As Long
Private mlMaxFontSize         As Long
Private mlMinFontSize         As Long
Private mbFreeze              As Boolean

Public pbButtonsMoveOnly      As Boolean  '''Indicates whether dimensions of the command buttons will be changed.  Defaults to true
''''''''''''''' Records Section '''''''''''''''''''''''''''''''''''''''''''

Private Type mudtRZP_ResizeParameters
   lLeft                As Long
   lTop                 As Long
   lWidth               As Long
   lHeight              As Long
   lFontSize            As Long
End Type

Private muRZP()         As mudtRZP_ResizeParameters

''''''''''''''' Objects Section '''''''''''''''''''''''''''''''''''''''''

Private mfrmSizer       As Form


Private Sub M_Resize(Optional ByVal bInitialSize As Boolean = False)
'''************************************************************************
'''Called By:        P_Resize
'''                  P_Move
'''                  Let plMaxFontSize
'''                  Let plMinFontSize
'''                  P_InitialSize
'''
'''Arguments
'''
'''Description:      Resize all controls proportionally to the current form
'''                  size
'''Note:
'''Warning:
'''Updates Globals:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME              As String = "M_Resize"
   
   Dim ctlDTN                    As Control
   Dim lIdx                      As Long
   Dim dRatioWidth               As Double
   Dim dRatioHeight              As Double
   Dim lFontSize                 As Long
   
On Error GoTo ErrHnd

   Call gobjLibApp.LibApp_sInitErr(True, mbInitialized)
   
On Error GoTo 0
On Error Resume Next
   
   If bInitialSize Then
      dRatioHeight = 1
      dRatioWidth = 1
   Else
      dRatioHeight = mfrmSizer.ScaleHeight / mlScaleHeight
      dRatioWidth = mfrmSizer.ScaleWidth / mlScaleWidth
   End If
   
   lIdx = 0
   For Each ctlDTN In mfrmSizer.Controls
      With ctlDTN
         lIdx = lIdx + 1
         '''Controls on background tabs may have negative left property (~ -60000 and less)
         If .Left >= -100 Then
            '''Control size and Position
            '''If initial value is not saved - first resize of the control
            If muRZP(lIdx).lLeft = mlRESIZE_INIT Then
               muRZP(lIdx).lLeft = .Left
               muRZP(lIdx).lTop = .Top
               muRZP(lIdx).lWidth = .Width
               muRZP(lIdx).lHeight = .Height
               muRZP(lIdx).lFontSize = .FontSize
            End If
            Call Err.Clear
            .Left = muRZP(lIdx).lLeft * dRatioWidth
            .Top = muRZP(lIdx).lTop * dRatioHeight

            If (Not pbButtonsMoveOnly) Or (Not TypeOf ctlDTN Is CommandButton) Then
               .Width = muRZP(lIdx).lWidth * dRatioWidth
               .Height = muRZP(lIdx).lHeight * dRatioHeight

               If Err.Number > 0 Then
                  Call Err.Clear
                  Call .Move(muRZP(lIdx).lLeft * dRatioWidth, _
                             muRZP(lIdx).lTop * dRatioHeight, _
                             muRZP(lIdx).lWidth * dRatioWidth, _
                             muRZP(lIdx).lHeight * dRatioHeight _
                            )
               End If
               '''Font
               If dRatioWidth < dRatioHeight Then
                  lFontSize = Fix(muRZP(lIdx).lFontSize * dRatioWidth)
               Else
                  lFontSize = Fix(muRZP(lIdx).lFontSize * dRatioHeight)
               End If
               If mlMaxFontSize > 0 Then
                  If lFontSize > mlMaxFontSize Then
                     lFontSize = mlMaxFontSize
                  End If
               End If
               If mlMinFontSize > 0 Then
                  If lFontSize < mlMinFontSize Then
                     lFontSize = mlMinFontSize
                  End If
               End If
               .FontSize = lFontSize
            End If
            DoEvents
         End If
      End With
   Next ctlDTN
   
   DoEvents
   Call Err.Clear
   
Exit Sub
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", True)

End Sub


Public Sub P_InitialSize()
   Call M_Resize(True)
End Sub


Public Sub P_Move(ByVal lLeft As Long, ByVal lTop As Long, ByVal lWidth As Long, ByVal lHeight As Long)
'''************************************************************************
'''Called By:        any
'''
'''Arguments
'''
'''Description:      Move the form and Resize all controls proportionally to the
'''                  new form size
'''Note:
'''Warning:
'''Updates Globals:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME              As String = "P_Move"
   
On Error GoTo ErrHnd

   Call gobjLibApp.LibApp_sInitErr(True, mbInitialized)
      
   Call mfrmSizer.Move(lLeft, lTop, lWidth, lHeight)
   DoEvents
   '''resize if not frozen
   Call P_Resize

Exit Sub
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", True)

End Sub


Public Property Get pbFreeze() As Boolean
   pbFreeze = mbFreeze
End Property

Public Property Let pbFreeze(ByVal bVal As Boolean)
   mbFreeze = bVal
   If mbInitialized And Not mbFreeze Then
      '''If unfreeze - resize
      Call P_Resize
   End If
   
End Property

Public Property Set pfrmForm(frmForm As Object)
'''************************************************************************
'''Called By:           P_Resize
'''Arguments
'''Description:         Set reference to the client form
'''Note:
'''Warning:
'''Updates Globals:
'''
'''Programmer:          Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Set pfrmForm"
   
On Error GoTo ErrHnd

   Set mfrmSizer = frmForm
   
   Call M_InitResize
      
Exit Property
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", True)
   
End Property


Public Sub P_Resize()
'''************************************************************************
'''Called By:        P_Move or Any
'''Arguments
'''
'''Description:      If resize function is not locked
'''                     Resize all controls proportionally to the current form size
'''Note:
'''Warning:
'''Updates Globals:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME              As String = "P_Resize"
   
On Error GoTo ErrHnd

   If Not mbFreeze Then
      Call M_Resize
   End If
   
Exit Sub
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")

End Sub


Private Sub M_InitResize()
'''************************************************************************
'''Called By:           set pfrmForm
'''Arguments
'''Description:         Evaluate form and prepare local properties
'''                     Based on its number of controls and current hight
'''                     and width
'''Note:
'''Warning:
'''Updates Globals:
'''
'''Programmer:          Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "M_InitResize"
   
   Dim lIdx             As Long
   Dim ctlDTN           As Control
   Dim lLine            As Long
   
On Error GoTo ErrHnd

   mbInitialized = False
   If Not mfrmSizer Is Nothing Then
                                                                                      lLine = 10
      mlScaleHeight = mfrmSizer.ScaleHeight
      mlScaleWidth = mfrmSizer.ScaleWidth
                                                                                      lLine = 20
      ReDim muRZP(1 To mfrmSizer.Controls.Count) As mudtRZP_ResizeParameters
                                                                                      lLine = 30
      For lIdx = 1 To UBound(muRZP)
         muRZP(lIdx).lLeft = mlRESIZE_INIT
      Next lIdx
   
      Call Err.Clear
      mbInitialized = True
   Else
                                                                                      lLine = 100
      Erase muRZP
      mlScaleHeight = 0
      mlScaleWidth = 0
   End If
   
Exit Sub
ErrHnd:
   Call gobjLibApp.LibApp_Error _
   (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE & lLine)

End Sub


Public Property Get plMaxFontSize() As Long
   plMaxFontSize = mlMaxFontSize
End Property

Public Property Let plMaxFontSize(ByVal lVal As Long)
   mlMaxFontSize = lVal
   Call M_Resize
End Property

Public Property Get plMinFontSize() As Long
   plMinFontSize = mlMinFontSize
End Property

Public Property Let plMinFontSize(ByVal lVal As Long)
   mlMinFontSize = lVal
   Call M_Resize
End Property

Private Sub Class_Initialize()
   pbButtonsMoveOnly = True
End Sub

Private Sub Class_Terminate()
On Error Resume Next
   Set mfrmSizer = Nothing
End Sub


