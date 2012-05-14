VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iObjectPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Abstract interface that must be implemented by a class to persist instances of the class.

Option Explicit

'A unique id for the class.
Public Property Get UniqueIdentifier() As String
End Property

'Current version of persistable class.  The version number should be updated anytime the persistable object is changed.
Public Property Get Version() As String
End Property

'Returns an exact copy of the object in its current state
Public Function Copy() As Object
End Function

'Saves the names and values of objects properties using the passed in Writer object.  The writer object defines the format
'used to persist the object.
Public Sub Persist(ByRef Writer As iWriter)
End Sub

'Restores values of an object from a Reader object.  The Reader should be compatiable with the writer used to persist the object.
'The properties MUST be restored in the same order as they where written.
Public Function Restore(ByRef Reader As iReader) As Boolean
End Function

