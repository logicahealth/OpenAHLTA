VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ETXRefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"ETXref"
Attribute VB_Ext_KEY = "Member0" ,"ETXref"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private mcolRefs As Collection

Public Function Add() As ETXref
    Dim objRef As ETXref
    Set objRef = New ETXref
    mcolRefs.Add objRef
    Set Add = objRef
    Set objRef = Nothing
End Function

Public Property Get Item(lIndexKey As Long) As ETXref
Attribute Item.VB_UserMemId = 0
  Set Item = mcolRefs(lIndexKey)
End Property


Public Property Get Count() As Long
    Count = mcolRefs.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    mcolRefs.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mcolRefs.[_NewEnum]
End Property


Private Sub Class_Initialize()
    Set mcolRefs = New Collection
End Sub


Private Sub Class_Terminate()
    Set mcolRefs = Nothing
End Sub

Public Function ClassHandle() As Long

End Function

Public Function ClearAll()
    Set mcolRefs = New Collection
End Function

Public Function IsEmpty() As Long
    IsEmpty = (mcolRefs.Count = 0)
End Function

