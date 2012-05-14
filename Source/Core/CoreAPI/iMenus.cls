VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iMenus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum enumMenuType
  mtAction = 1
  mtEdit = 2
  mtPrint = 3
  mtPrintPreview = 4
  mtView = 5
End Enum

Public Property Get Menu(ByVal MenuType As enumMenuType) As iMenu

End Property

Public Property Get ProgramID() As Long

End Property

Public Property Let ProgramID(ByVal OLEServerID As Long)

End Property

Public Sub ClearMenus()

End Sub

Public Sub Refresh(ByVal MenuType As enumMenuType)

End Sub

Public Sub RefreshAll()

End Sub
