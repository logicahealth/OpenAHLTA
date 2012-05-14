VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public MnuButtonID As Integer 'Unique ID for the menu button
Public MnuButtonName As String
Public MnuSection As Boolean 'Single Tooth (F) or Multi (T)
Public MnuGroup As eGroup
Public MnuItems As colMenuItems
Private Sub Class_Initialize()

Set MnuItems = New colMenuItems

End Sub

Private Sub Class_Terminate()
Set MnuItems = Nothing
End Sub
