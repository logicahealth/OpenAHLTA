VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IParent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum MenuNames
  mnAction = 0
  mnView = 1
End Enum

Public Property Get Hwnd() As Long
End Property

Public Property Get CurrentEncounter() As Encounter
End Property

Public Property Get SectionObject() As Object
End Property

Public Function AddMenuItem(ByVal iMenu As MenuNames, _
                            ByVal sMenuText As String, _
                            ByVal sMenuBMPFileName As String, _
                            ByVal sTip As String) As Long
End Function

Public Function RemoveMenuItem(ByVal lMenuID As Long) As Boolean

End Function

Public Function LockEventID(ByVal nEventId As Long) As Boolean
End Function

Public Function UnlockEventID(ByVal nEventId As Long) As Boolean
End Function

Public Function StartSave(Optional ByVal nEventId As Long) As Boolean
End Function

Public Sub SetRTF(ByRef RTF As String)
End Sub

Public Sub SetStatus(ByVal eStatus As EncStatus)
End Sub

Public Function EndSave(Optional ByVal nEventId As Long) As Boolean
End Function

Public Sub ClosedView()
End Sub


