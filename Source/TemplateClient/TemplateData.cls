VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TemplateData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub DoMoveDown()

End Sub

Public Sub DoMoveUp()

End Sub


Public Property Get Items() As Scripting.Dictionary
End Property

Public Property Get CanEdit() As Boolean
End Property

Public Property Get isDirty() As Boolean
End Property

Public Property Let isDirty(bVal As Boolean)
End Property

Public Property Get Label() As String
End Property

Public Property Get ListBox() As Object
End Property

Public Property Set ListBox(rListBox As Object)
End Property

Public Sub DoAdd()
End Sub
Public Sub DoEdit()
End Sub

Public Sub DoRemove()
End Sub

Public Function Serialize(rvData As Variant) As Boolean
End Function

Public Function Deserialize(rvData As Variant) As Boolean
End Function

Public Function GetFromEncounter(ByRef objEncounter As CHCSIIEncounterOps.Encounter) As Boolean
End Function

Public Function GetQueryCriteria() As String
End Function

Public Function Import(rsString As String) As Boolean
End Function

Public Function Export() As String
End Function

