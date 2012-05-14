VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OEOtherOrderTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--- begin DTO Builder; DTO=OEOtherOrderTypes
'--- 4/1/2004 3:33:57 PM
'---
'--- NOTE: DO NOT MODIFY THIS CLASS
'
'
Implements iObjectPersist

Private Const OBJECT_VERSION = "1.0"
Private Const UNIQUE_IDENTIFIER = "OEOtherOrderTypes"


Private mCol As Collection
Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub


Public Function Add(Optional ByVal o As OEOtherOrderType = Nothing, Optional ByVal Key As String = "") As OEOtherOrderType
    If o Is Nothing Then Set o = New OEOtherOrderType

    If LenB(Key) = 0 Then
        mCol.Add o
    Else
        mCol.Add o, Key
    End If
End Function
Private Function iObjectPersist_Copy() As Object

On Error GoTo ErrHandler

    Dim NewMe As OEOtherOrderTypes
    Dim Persist As iObjectPersist
    Dim Writer As iWriter
    Dim Reader As iReader

    Set NewMe = New OEOtherOrderTypes
  
    Set Writer = New WriterXML
    iObjectPersist_Persist Writer
  
    Set Reader = New ReaderXML
    Set Persist = NewMe
  
    Reader.DocumentVersion = Persist.Version
    Reader.Contents = Writer.Contents
  
    Persist.Restore Reader
  
    Set iObjectPersist_Copy = NewMe

Exit Function
ErrHandler:
    Err.Raise Err.Number, "DTOsOEOtherOrderTypes.iObjectPersist_Copy", Err.Description, Err.HelpFile, Err.HelpContext
  
End Function
Private Sub iObjectPersist_Persist(Writer As iWriter)
    
On Error GoTo ErrHandler
    
    Dim i As Long

    With Writer
        .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION, mCol.Count
    
        For i = 1 To mCol.Count
            .WriteObject "Item_" & i, mCol.Item(i)
        Next
    End With

Exit Sub
ErrHandler:
    Err.Raise Err.Number, "DTOsOEOtherOrderTypes.iObjectPersist_Persist", Err.Description, Err.HelpFile, Err.HelpContext

End Sub
Private Function iObjectPersist_Restore(Reader As iReader) As Boolean
    Dim Obj As OEOtherOrderType
    Dim i As Long
    Dim NumObjs As Long

    On Error GoTo ErrHandler
  
    With Reader
        NumObjs = .CollectionCount
    
        For i = 1 To NumObjs
            Set Obj = New OEOtherOrderType
            .ReadObject "Item_" & i, Obj
            mCol.Add Obj
      
            Set Obj = Nothing
        Next
    End With

    iObjectPersist_Restore = True
    
Exit Function
ErrHandler:
    Err.Raise Err.Number, "DTOsOEOtherOrderTypes.iObjectPersist_Restore", Err.Description, Err.HelpFile, Err.HelpContext
End Function
Private Property Get iObjectPersist_UniqueIdentifier() As String
    iObjectPersist_UniqueIdentifier = UNIQUE_IDENTIFIER
End Property

Private Property Get iObjectPersist_Version() As String
    iObjectPersist_Version = OBJECT_VERSION
End Property
Public Function Count() As Long
    Count = mCol.Count
End Function

Public Function Item(ByVal Key As Variant) As OEOtherOrderType
    Set Item = mCol.Item(Key)
End Function

Public Function NewEnum() As IUnknown
    Set NewEnum = mCol.[_NewEnum]
End Function

Public Sub Remove(ByVal Key As Variant)
    On Error Resume Next
    mCol.Remove Key
End Sub
'--- end DTO Builder


