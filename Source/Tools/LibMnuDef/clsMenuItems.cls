VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenuItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public MnuID As Integer
Public MnuItemId As Integer
Public MnuItemName As String
Public MnuItemCode As String
Public MnuItemCodeID As Integer
Public MnuItemDRC As String
Public MnuItemVisible As Boolean
Public MnuItemInstructions As Collection

Private Sub Class_Initialize()
Set MnuItemInstructions = New Collection
End Sub

Private Sub Class_Terminate()
Set MnuItemInstructions = Nothing
End Sub
