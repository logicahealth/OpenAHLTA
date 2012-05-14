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
Attribute VB_Name = "WriterRecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Converts a recordset to XML.

Option Explicit

Implements iWriter

Private mStream As Stream

Private Property Get iWriter_Contents() As String
   iWriter_Contents = mStream.ReadText
End Property

Private Sub iWriter_InitHeader(ByVal Name As String, Optional ByVal Version As String = vbNullString, Optional ByVal Count As Long = -1&)
'
End Sub

Private Sub iWriter_WriteArray(ByVal Name As String, Ary As Variant)
'
End Sub

Private Sub iWriter_WriteCollection(ByVal Name As String, col As Collection)
'
End Sub

'Object must be a recordset
Private Sub iWriter_WriteObject(ByVal Name As String, Obj As Object)
    
  On Error Resume Next
  
  Set mStream = Nothing
  Set mStream = New Stream
  
  Obj.Save mStream, adPersistXML

End Sub

Private Sub iWriter_WriteProperty(ByVal Name As String, ByVal Value As Variant)
'
End Sub
