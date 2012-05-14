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
Attribute VB_Name = "iReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Abstract interface that must be implemented by a class to support object persistance.  The classes that implement this
'interface will be responsible for restoring values to an object from a string.

Option Explicit

'Count of like objects persisted in a document.
Public Property Get CollectionCount() As Long
End Property

'Sets the persitent value of an object.
Public Property Let Contents(ByRef NewValue As String)
End Property

'Document version
Public Property Get DocumentVersion() As String
End Property

'Document version
Public Property Let DocumentVersion(ByVal Version As String)
End Property

'Returns an array of strings.
Public Function ReadArray(ByVal Name As String) As Variant
End Function

'Returns a collection of strings.
Public Function ReadCollection(ByVal Name As String) As Collection
End Function

'Restores the values of an object.
Public Sub ReadObject(ByVal Name As String, ByRef Obj As Object)
End Sub

'Returns all simple data types as a string.  VB allows us to use this cheat by doing the type conversion for us.
Public Function ReadProperty(ByVal Name As String) As String
End Function


