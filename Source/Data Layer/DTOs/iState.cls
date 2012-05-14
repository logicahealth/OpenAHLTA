VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'This interface is used by DTOs that must support state.
'
'RULES OF USE:
'State = Saved when a existing object is loaded.
'State = New when a blank object is created.
'State = Saved after object is successfully saved to the data source.
'State = Removed after object is successfully deleted from the data source.

Option Explicit

Public Enum enumObjectState
  osUnknown = 0
  osNew = 1
  osDelete = 2
  osModified = 3
  osQueued = 4
  osSaved = 5
  osRemoved = 6
End Enum

Public State As enumObjectState
Public ID As String

