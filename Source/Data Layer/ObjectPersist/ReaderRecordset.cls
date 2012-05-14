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
Attribute VB_Name = "ReaderRecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Converts XML to a recordset.

Option Explicit

Implements iReader

Private mStream As Stream

Private Property Get iReader_CollectionCount() As Long
'
End Property

Private Property Let iReader_Contents(RHS As String)
  Set mStream = Nothing
  Set mStream = New Stream
  
  With mStream
    .Open
    .WriteText RHS
  End With
End Property

Private Property Let iReader_DocumentVersion(ByVal RHS As String)
'
End Property

Private Property Get iReader_DocumentVersion() As String
'
End Property

Private Function iReader_ReadArray(ByVal Name As String) As Variant
'
End Function

Private Function iReader_ReadCollection(ByVal Name As String) As Collection
'
End Function

'Object must be a recordset
Private Sub iReader_ReadObject(ByVal Name As String, Obj As Object)
    
  mStream.Position = 0
  Obj.Open mStream

End Sub

Private Function iReader_ReadProperty(ByVal Name As String) As String

End Function
