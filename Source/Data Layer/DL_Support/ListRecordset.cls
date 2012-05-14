VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ListRecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Wrapper around a recordset.  Provides the iList interface for navigation.

Option Explicit

Implements iList
Implements iObjectPersist

Private Const UNIQUE_IDENTIFIER As String = "ListRecordset"
Private Const OBJECT_VERSION As String = "1.0"

Private mRecordSource As Recordset

'Creates a new row in the recordset and adds the values from a iParameters object.
Private Sub iList_AddRow(Parameters As iParameters)
Dim i As Long
  
  On Error GoTo ErrHandler
  
  With mRecordSource
    .AddNew
          
    For i = 1 To Parameters.Count
      .Fields(Parameters.Item(i).Name) = Parameters.Item(i).Value
    Next
    
    .Update
  End With
  
  Exit Sub
  
ErrHandler:
  mRecordSource.CancelUpdate
  
End Sub

Private Property Get iList_BOF() As Boolean
  iList_BOF = mRecordSource.BOF Or mRecordSource.RecordCount = 0
End Property

Private Property Get iList_EOF() As Boolean
  iList_EOF = mRecordSource.EOF Or mRecordSource.RecordCount = 0
End Property

'Creates and returns a collection of parameters (ParametersGeneric object) from a row in the recordset.
'NOTE:  You cannot update the recordset by updating the the parameters collection.
Private Property Get iList_Fields() As iParameters
Dim Parameters As iParameters
Dim flds As Fields
Dim fld As Field

  Set Parameters = New ParametersGeneric

  If Not (mRecordSource.BOF Or mRecordSource.EOF) Then
    Set flds = mRecordSource.Fields
    
    For Each fld In flds
      Parameters.Add fld.Name, fld.Value
    Next
  End If
  
  Set iList_Fields = Parameters
  
End Property

Private Sub iList_MoveFirst()
  If mRecordSource.RecordCount > 0 Then mRecordSource.MoveFirst
End Sub

Private Sub iList_MoveLast()
  If mRecordSource.RecordCount > 0 Then mRecordSource.MoveLast
End Sub

Private Sub iList_MoveNext()
  If mRecordSource.RecordCount > 0 Then mRecordSource.MoveNext
End Sub

Private Sub iList_MovePrevious()
  If mRecordSource.RecordCount > 0 Then mRecordSource.MovePrevious
End Sub

Private Property Get iList_RecordCount() As Long
  iList_RecordCount = mRecordSource.RecordCount
End Property

Private Sub iList_RemoveCurrentRow()
  If mRecordSource.RecordCount > 0 Then mRecordSource.Delete
  
  If mRecordSource.RecordCount > 0 Then
    mRecordSource.MoveNext
    
    If mRecordSource.EOF Then
      mRecordSource.MovePrevious
    End If
  End If

End Sub

Private Property Get iList_Value(ByVal FieldName As String) As Variant
  iList_Value = mRecordSource.Fields(FieldName).Value
End Property

Private Function iObjectPersist_Copy() As Object
Dim NewMe As ListRecordset

  Set NewMe = New ListRecordset
  Set NewMe.Source = mRecordSource.Clone
    
  Set iObjectPersist_Copy = NewMe

End Function

Private Sub iObjectPersist_Persist(Writer As ObjectPersist.iWriter)
  Writer.WriteObject vbNullString, mRecordSource
End Sub

Private Function iObjectPersist_Restore(Reader As ObjectPersist.iReader) As Boolean
  Set mRecordSource = New Recordset
  Reader.ReadObject vbNullString, mRecordSource
End Function

Private Property Get iObjectPersist_UniqueIdentifier() As String
  iObjectPersist_UniqueIdentifier = UNIQUE_IDENTIFIER
End Property

Private Property Get iObjectPersist_Version() As String
  iObjectPersist_Version = OBJECT_VERSION
End Property

Public Property Set Source(ByRef rs As Recordset)
  Set mRecordSource = rs
  
  With mRecordSource
    Set .ActiveConnection = Nothing
    
    If .State <> adStateOpen Then
      .CursorLocation = adUseClient
      .CursorType = adOpenDynamic
      .LockType = adLockBatchOptimistic
      .Open
    End If
  End With
End Property

Public Property Get Source() As Recordset
  Set Source = mRecordSource
End Property

