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
Attribute VB_Name = "iWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Abstract interface that must be implemented by a class to support object persistance.  The classes that implement this
'interface will be responsible for saving values of an object to a string.

Option Explicit

'The persisted valued of an object.
Public Property Get Contents() As String
End Property

'Used to create save document name and version in the main document or a sub document.  Each object saved is a document so
'one document may contain several sub documents.  The count property is used to saved the number of object persisted in a
'collection class object.
Public Sub InitHeader(ByVal Name As String, Optional ByVal Version As String = vbNullString, Optional ByVal Count As Long = -1)
End Sub

'Saves an array of simple data types only.
Public Sub WriteArray(ByVal Name As String, ByRef Ary As Variant)
End Sub

'Saves a collection of simple data types only.  If you want to persist a collection of object, you must use a collection class.
Public Sub WriteCollection(ByVal Name As String, ByRef col As Collection)
End Sub

'Saves values of an object.
Public Sub WriteObject(ByVal Name As String, ByRef Obj As Object)
End Sub

'Saves simple data types.
Public Sub WriteProperty(ByVal Name As String, ByVal Value As Variant)
End Sub
