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
Attribute VB_Name = "clsLibMsPnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''************************************************************************
'''Module:        LibMsPnt.cls - clsLibMsPnt Mouse Pointer Class
'''
'''Description:   Allowes to change mouse pointer and automaticaly restore
'''               in terminate event
'''
'''Suggested Use:
'''
'''Note:
'''
'''Modules Required:
'''
'''Author         : Alexander Alberstein
'''Modified       :
'''               :
'''************************************************************************
Option Explicit
''''''''''''''' Constanst Section '''''''''''''''''''''''''''''''''''''''''

Const msMODULE_NAME        As String = "LibMsPnt."

Private mlOriginalPointer  As MousePointerConstants
Private mlLastSetPointer   As MousePointerConstants
Public Sub Mtd_SetMousePointer(Optional nPointer As MousePointerConstants = vbHourglass)
   mlLastSetPointer = nPointer
   If Not (Screen.MousePointer = nPointer) Then
      Screen.MousePointer = nPointer
   End If
End Sub


Public Sub Mtd_Toggle(ByVal bOriginal As Boolean)
   
   If bOriginal Then
      Screen.MousePointer = mlOriginalPointer
   Else
      Screen.MousePointer = mlLastSetPointer
   End If
   
End Sub

Private Sub Class_Initialize()
   mlOriginalPointer = Screen.MousePointer
   
End Sub


Private Sub Class_Terminate()
   Screen.MousePointer = mlOriginalPointer
   
End Sub


