VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Property Get Singleton() As SVGTeeth.clsTeeth
' See if it already exists.
    If g_Singleton Is Nothing Then
        Set g_Singleton = New SVGTeeth.clsTeeth
    End If
    Set Singleton = g_Singleton
End Property

Public Sub TerminateSF()
    
    If Not g_Singleton Is Nothing Then
        g_Singleton.TerminateObjects
        Set g_Singleton = Nothing
    End If
    
    'Terminate clsTeeth
'    Set g_Singleton = Nothing
    
End Sub


