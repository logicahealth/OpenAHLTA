VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Parameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'A single key value pair object.

Option Explicit

Implements iObjectPersist

Private Const UNIQUE_IDENTIFIER As String = "Parameter"
Private Const OBJECT_VERSION As String = "1.0"

Private mName As String             'Key
Private mValue As Variant           'Value

Private Function iObjectPersist_Copy() As Object
Dim NewMe As Parameter

  Set NewMe = New Parameter
  
  NewMe.Name = mName
  NewMe.Value = mValue
  
  Set iObjectPersist_Copy = NewMe
  
End Function

Private Sub iObjectPersist_Persist(Writer As ObjectPersist.iWriter)
  
  With Writer
    .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION
    
    .WriteProperty "Name", mName
    .WriteProperty "Value", mValue
  End With

End Sub

Private Function iObjectPersist_Restore(Reader As ObjectPersist.iReader) As Boolean
Dim Failed As Boolean

  On Error GoTo ErrHandler
  
  With Reader
    mName = .ReadProperty("Name")
    mValue = .ReadProperty("Value")
  End With

  iObjectPersist_Restore = Not Failed
  
  Exit Function
  
ErrHandler:
  Failed = True
  Resume Next

End Function

Private Property Get iObjectPersist_UniqueIdentifier() As String
  iObjectPersist_UniqueIdentifier = UNIQUE_IDENTIFIER
End Property

Private Property Get iObjectPersist_Version() As String
  iObjectPersist_Version = OBJECT_VERSION
End Property

Public Property Get Name() As String
  Name = mName
End Property

Public Property Let Name(ByVal NewValue As String)
  mName = NewValue
End Property

Public Property Get Value() As Variant
  If IsObject(mValue) Then
    Set Value = mValue
  Else
    Value = mValue
  End If
End Property

Public Property Let Value(ByVal NewValue As Variant)
  On Error Resume Next
  
  If IsObject(NewValue) Then
    Set mValue = NewValue
  Else
    mValue = NewValue
  End If
End Property

