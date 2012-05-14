VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ETItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"ETItem"
Attribute VB_Ext_KEY = "Member0" ,"ETItem"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection

Public Function Add() As ETItem
    Dim objItem As ETItem
    Set objItem = New ETItem
    mCol.Add objItem
    Set Add = objItem
    Set objItem = Nothing
End Function

Public Property Get Item(lIndexKey As Long) As ETItem
Attribute Item.VB_UserMemId = 0
  Set Item = mCol(lIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

Public Function IsEmpty() As Long
    IsEmpty = (mCol.Count = 0)
End Function
Public Function ClearAll()
    Set mCol = New Collection
End Function
Public Function ClassHandle() As Long

End Function

