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
Attribute VB_Name = "PersistHelper"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'This global multiuse class provides functions to simplify the persistance and restoration of objects.

Option Explicit

'Copy data from one instance of a class to a new instance of the same class.
Public Function CopyObject(ByRef CopyMe As Object) As Object
Dim Persist As iObjectPersist

  Set Persist = CopyMe
  Set CopyObject = Persist.Copy
  
End Function

'Persists a recordset or an object with the iObjectPersist interface and a recordset as its interal data source.
Public Function Recordset2XML(ByRef Obj As Object) As String
Dim Writer As iWriter
Dim Persist As iObjectPersist

  Set Writer = New WriterRecordset
    
  If TypeOf Obj Is Recordset Then
    Writer.WriteObject vbNullString, Obj
  ElseIf TypeOf Obj Is iObjectPersist Then
    Set Persist = Obj
    Persist.Persist Writer
  End If
  
  Recordset2XML = Writer.Contents

End Function

'Persists an object with the iObjectPersist interface to XML.
Public Function Object2XML(ByRef Obj As Object) As String
Dim Writer As iWriter
Dim Persist As iObjectPersist

  Set Writer = New WriterXML
    
  Set Persist = Obj
  Persist.Persist Writer
  
  Object2XML = Writer.Contents
  
End Function

'Determins if the values of 2 are exactly the same.  Both objects must support the iObjectPersist interface.
Public Function ObjectsEqual(ByRef CompareMe As Object, ByRef WithMe As Object) As Boolean
Dim CompareXML As String
Dim WithXML As String

  CompareXML = Object2XML(CompareMe)
  WithXML = Object2XML(WithMe)

  ObjectsEqual = (CompareXML = WithXML)
  
End Function

'Determins if the values of 2 are exactly the same.  Both objects must support the iObjectPersist interface.
Public Function RecordsetsEqual(ByRef CompareMe As Object, ByRef WithMe As Object) As Boolean
Dim CompareXML As String
Dim WithXML As String

  CompareXML = Recordset2XML(CompareMe)
  WithXML = Recordset2XML(WithMe)

  RecordsetsEqual = (CompareXML = WithXML)
  
End Function

'Updates data from one instance of a class to the values of another instance of the same class.
'Both objects must support the iObjectPersist interface.
Public Sub UpdateObject(ByRef UpdateMe As Object, ByRef WithMe As Object)
Dim XML As String

  XML = Object2XML(WithMe)
  
  XML2Object XML, UpdateMe
  
End Sub

'Updates data from one instance of a recordset to the values of another instance of the same recordset.
'Both objects must support the iObjectPersist interface.
Public Sub UpdateRecordset(ByRef UpdateMe As Object, ByRef WithMe As Object)
Dim XML As String

  XML = Recordset2XML(WithMe)
  
  XML2Recordset XML, UpdateMe
  
End Sub

'Restores a recordset or an object's recordset from a XML string.
Public Function XML2Recordset(ByVal XML As String, ByRef Obj As Object) As Boolean
Dim Reader As iReader
Dim Persist As iObjectPersist
Dim Success As Boolean

  Set Reader = New ReaderRecordset
  Reader.Contents = XML
  
  If TypeOf Obj Is Recordset Then
    Reader.ReadObject vbNullString, Obj
  ElseIf TypeOf Obj Is iObjectPersist Then
    Set Persist = Obj
    Success = Persist.Restore(Reader)
  End If
  
  XML2Recordset = Success
  
End Function

'Restores an object's values from a XML string.
Public Function XML2Object(ByVal XML As String, ByRef Obj As Object) As Boolean
Dim Reader As iReader
Dim Persist As iObjectPersist

  Set Persist = Obj

  Set Reader = New ReaderXML
  Reader.DocumentVersion = Persist.Version
  Reader.Contents = XML
  
  XML2Object = Persist.Restore(Reader)
  
End Function

