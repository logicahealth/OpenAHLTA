VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ModuleConfigs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'NOTE: This class should NEVER be modified.

Option Explicit

Implements iObjectPersist

Private Const UNIQUE_IDENTIFIER As String = "ModuleConfigs"
Private Const OBJECT_VERSION As String = "1.0"

Private mCol As Collection

Private Sub Class_Initialize()
  Set mCol = New Collection
End Sub

Private Function iObjectPersist_Copy() As Object
Dim NewMe As ModuleConfigs
Dim Persist As iObjectPersist
Dim Writer As iWriter
Dim Reader As iReader

  Set NewMe = New ModuleConfigs
  
  Set Writer = New WriterXML
  iObjectPersist_Persist Writer
  
  Set Reader = New ReaderXML
  Set Persist = NewMe
  
  Reader.DocumentVersion = Persist.Version
  Reader.Contents = Writer.Contents
  
  Persist.Restore Reader
  
  Set iObjectPersist_Copy = NewMe
  
End Function

Private Sub iObjectPersist_Persist(Writer As ObjectPersist.iWriter)
Dim i As Long

  With Writer
    .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION, mCol.Count
    
    For i = 1 To mCol.Count
      .WriteObject "Item_" & i, mCol.Item(i)
    Next
  End With

End Sub

Private Function iObjectPersist_Restore(Reader As ObjectPersist.iReader) As Boolean
Dim Obj As ModuleConfig
Dim i As Long
Dim NumObjs As Long

  On Error GoTo ErrHandler
  
  With Reader
    NumObjs = .CollectionCount
    
    For i = 1 To NumObjs
      Set Obj = New ModuleConfig
      
      .ReadObject "Item_" & i, Obj
      mCol.Add Obj, Obj.ApplicationNCID               'Note: If Collection has a Key, developer must add it here.
      
      Set Obj = Nothing
    Next
  End With

  iObjectPersist_Restore = True
  
  Exit Function
  
ErrHandler:

End Function

Private Property Get iObjectPersist_UniqueIdentifier() As String
  iObjectPersist_UniqueIdentifier = UNIQUE_IDENTIFIER
End Property

Private Property Get iObjectPersist_Version() As String
  iObjectPersist_Version = OBJECT_VERSION
End Property

Public Function Add(Optional ByVal Key As String = vbNullString) As ModuleConfig
Dim Item As ModuleConfig

  Set Item = New ModuleConfig
  
  If Len(Key) > 0 Then
    mCol.Add Item, Key
  Else
    mCol.Add Item
  End If
  
  Set Add = Item
  
End Function

Public Function Count() As Long
  Count = mCol.Count
End Function

Public Function Item(ByVal Key As Variant) As ModuleConfig
Attribute Item.VB_UserMemId = 0
  Set Item = mCol.Item(Key)
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
  Set NewEnum = mCol.[_NewEnum]
End Function

Public Sub Remove(ByVal Key As Variant)
  On Error Resume Next
  mCol.Remove Key
End Sub

