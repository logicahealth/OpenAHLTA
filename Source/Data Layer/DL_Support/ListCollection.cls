'***************************************************************************************
'
'  Copyright (c) 2007-2012 Northrop Grumman Corporation
'
'  Licensed by Tricare Management Activity under license from the Copyright owner.
'
'  This text file must be included in all Derivative Works of the licensed Source Code.
'
'***************************************************************************************

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ListCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Wrapper around a collection of ParameterGeneric objects.  Each ParametersGeneric object is treated like a row in a
'recordset.  Provides the iList interface for Navigation.

Option Explicit

Implements iList
Implements iObjectPersist

Private Const UNIQUE_IDENTIFIER As String = "ListCollection"
Private Const OBJECT_VERSION As String = "1.0"

Private mList As Collection                         'Collection of ParametersGeneric objects
Private mCurrentIndex As Long                       'Pointer to currently selected ParametersGeneric object in collection

Private Sub Class_Initialize()
  Set mList = New Collection
End Sub

Private Sub iList_AddRow(Parameters As iParameters)
  mList.Add Parameters
  mCurrentIndex = mList.Count
End Sub

Private Property Get iList_BOF() As Boolean
  iList_BOF = (mCurrentIndex = 0 Or mList.Count = 0)
End Property

Private Property Get iList_EOF() As Boolean
  iList_EOF = (mCurrentIndex > mList.Count Or mList.Count = 0)
End Property

Private Property Get iList_Fields() As iParameters
  If mCurrentIndex >= 1 And mCurrentIndex <= mList.Count Then
    Set iList_Fields = mList.Item(mCurrentIndex)
  Else
    Err.Raise 1000, "ListColection.iList_Fields", "Item not found in the collection."
  End If
End Property

Private Sub iList_MoveFirst()
  If mList.Count > 0 Then
    mCurrentIndex = 1
  Else
    mCurrentIndex = 0
  End If
End Sub

Private Sub iList_MoveLast()
  If mList.Count > 0 Then
    mCurrentIndex = mList.Count
  Else
    mCurrentIndex = 0
  End If
End Sub

Private Sub iList_MoveNext()
  If mList.Count >= mCurrentIndex Then
    mCurrentIndex = mCurrentIndex + 1
  End If
End Sub

Private Sub iList_MovePrevious()
  If mCurrentIndex > 0 Then
    mCurrentIndex = mCurrentIndex - 1
  End If
End Sub

Private Property Get iList_RecordCount() As Long
  iList_RecordCount = mList.Count
End Property

Private Sub iList_RemoveCurrentRow()
  If Not (mCurrentIndex >= 1 And mCurrentIndex <= mList.Count) Then Exit Sub
  
  mList.Remove mCurrentIndex
  
  If mList.Count = 0 Then
    mCurrentIndex = 0
  ElseIf mCurrentIndex >= mList.Count Then
    mCurrentIndex = mList.Count
  ElseIf mCurrentIndex <= 0 Then
    iList_MoveNext
  End If
End Sub

Private Property Get iList_Value(ByVal FieldName As String) As Variant
Dim Params As iParameters

  If Not (mCurrentIndex >= 1 And mCurrentIndex <= mList.Count) Then Exit Sub

  Set Params = mList.Item(mCurrentIndex)
  If Params.Exists(FieldName) Then iList_Value = Params.Item(FieldName).Value
  
End Property

Private Function iObjectPersist_Copy() As Object
Dim NewMe As ListCollection
Dim Persist As iObjectPersist
Dim Writer As iWriter
Dim Reader As iReader

  Set NewMe = New ListCollection
  
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
    .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION, mList.Count
    
    For i = 1 To mList.Count
      .WriteObject "Item_" & i, mList.Item(i)
    Next
  End With

End Sub

Private Function iObjectPersist_Restore(Reader As ObjectPersist.iReader) As Boolean
Dim Obj As ParametersGeneric
Dim i As Long
Dim NumObjs As Long

  On Error GoTo ErrHandler
  
  Set mList = Nothing
  Set mList = New Collection
  
  With Reader
    NumObjs = .CollectionCount
    
    For i = 1 To NumObjs
      Set Obj = New ParametersGeneric
      .ReadObject "Item_" & i, Obj
      mList.Add Obj
      
      Set Obj = Nothing
    Next
  End With

  If mList.Count > 0 Then mCurrentIndex = 1 Else mCurrentIndex = 0
  
  iObjectPersist_Restore = True
  
  Exit Function
  
ErrHandler:
  mCurrentIndex = 0
  
End Function

Private Property Get iObjectPersist_UniqueIdentifier() As String
  iObjectPersist_UniqueIdentifier = UNIQUE_IDENTIFIER
End Property

Private Property Get iObjectPersist_Version() As String
  iObjectPersist_Version = OBJECT_VERSION
End Property

Public Property Get Count() As Long
  Count = iList_RecordCount
End Property

Public Function Item(ByVal Key As Long) As ParametersGeneric
  Set Item = mList.Item(Key)
End Function
