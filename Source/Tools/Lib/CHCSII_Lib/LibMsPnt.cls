VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mclsLibMsPnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
Private mobjScreen         As Screen

Public Sub Mtd_SetMousePointer(ByRef objScreen As Object, Optional nPointer As MousePointerConstants = vbHourglass)
   Static bFirstPointerIsSaved_static As Boolean
   
On Error Resume Next

   Set mobjScreen = objScreen
   
   If Not bFirstPointerIsSaved_static Then
      mlOriginalPointer = mobjScreen.MousePointer
      bFirstPointerIsSaved_static = True
   End If
   
   mlLastSetPointer = nPointer
   
   If Not (mobjScreen.MousePointer = nPointer) Then
      mobjScreen.MousePointer = nPointer
   End If
   
End Sub


Public Sub Mtd_Toggle(ByVal bOriginal As Boolean)
   
   If bOriginal Then
      mobjScreen.MousePointer = mlOriginalPointer
   Else
      mobjScreen.MousePointer = mlLastSetPointer
   End If
   
End Sub

Private Sub Class_Initialize()
   mlOriginalPointer = Screen.MousePointer
   
End Sub


Private Sub Class_Terminate()
   mobjScreen.MousePointer = mlOriginalPointer
   
End Sub


