VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "I_IDMappers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'<SCR 24303 CC -  new class>

Option Explicit

Public Function Add(ApplicationNCID As Long, name As String, Before As String, After As String) As I_IDMapper
End Function

Public Property Get Item(vntIndexKey As Variant) As I_IDMapper
End Property

Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
End Property

Public Sub Clear()
End Sub

Public Property Get NewEnum() As IUnknown
    'this property allows you to enumerate
    'this collection with the For...Each syntax
End Property

Public Function Find(ByVal ApplicationNCID As Long, ByVal name As String, _
                     Optional ByVal Before As String, Optional ByVal After As String) As String
End Function

Public Function Exists(ByVal ApplicationNCID As Long, ByVal name As String, _
                        Optional ByVal Before As String, Optional ByVal After As String) As Boolean
End Function
