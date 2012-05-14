VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iGUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get ContainerControl() As Object
End Property

Public Property Get HoldGUIUpdates() As Boolean
End Property

Public Property Let HoldGUIUpdates(ByVal NewValue As Boolean)
End Property

Public Sub DisableGUI(ByVal CallingMethod As String)
End Sub

Public Sub EnableGUI(ByVal CallingMethod As String, Optional ByVal ForceEnable As Boolean = False)
End Sub

Public Function GetVisualIndicators() As Scripting.Dictionary
End Function

Public Function Message(ByVal MessageType As String, ByRef Value As Variant) As Boolean
End Function

Public Sub PatientChanged(Optional ByVal Cleared As Boolean = False)
End Sub

Public Sub RefershGUI(Optional ByVal DelayRefresh As Boolean = False)
End Sub

Public Sub UpdateVisualIndicators()
End Sub

Public Function UserChanged() As Boolean
End Function

