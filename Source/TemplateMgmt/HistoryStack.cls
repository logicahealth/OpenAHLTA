VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HistoryStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCol As Collection

Public Sub Push(ByRef NewEntry As HistoryEntry)
    mCol.Add NewEntry
End Sub
Public Function Pop() As HistoryEntry
    If mCol.Count = 0 Then Exit Function
    Set Pop = mCol.Item(mCol.Count)
    mCol.Remove mCol.Count
End Function

Public Function Count() As Integer
    Count = mCol.Count
End Function

Public Function Item(Index As Integer) As HistoryEntry
Attribute Item.VB_UserMemId = 0
    Set Item = mCol.Item(Index)
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Function

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

