VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDebugLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'This record should be declared in each procedure and tracks the time elapsed
'while in a procedure
Public Type DebugRec
    TimerTime As Single
End Type
Public Sub DebugLog(Exe As String, Module As String, Proc As String, Oper As String, tDR As DebugRec, ParamArray aVarList() As Variant)
    Dim vntList() As Variant    'Array of parameters and variables passed into this procedure
    If giDefaultDebugMode = 3 Then Exit Sub
    vntList = aVarList
    DL.DebugLog Exe, Module, Proc, Oper, tDR, vntList()
End Sub


Public Sub DebugLogPing(Exe As String, Module As String, Proc As String, Oper As String, DR As DebugRec, PingName As String, ParamArray VarList() As Variant)
    Dim vntList() As Variant    'Array of parameters and variables passed into this procedure
    If giDefaultDebugMode = 3 Then Exit Sub
    vntList = VarList
    DL.DebugLogPing Exe, Module, Proc, Oper, DR, PingName, vntList()
End Sub

Public Sub debugStart()
    DL.debugStart
End Sub


Public Sub debugStop()
    If giDefaultDebugMode = 3 Then Exit Sub
    DL.debugStop
End Sub
'<<< SCR 22541  MSolano 3/29/2002
Public Function ARMOverride() As Boolean

  If giARMOverride = gcARMOverrideOn Then
    ARMOverride = True
  Else
    ARMOverride = False
  End If

End Function
'>>> SCR 22541  MSolano 3/29/2002


