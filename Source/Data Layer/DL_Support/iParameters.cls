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
Attribute VB_Name = "iParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Abstract interface that classes should implement to support a collection of parameters.

Option Explicit

'Adds a new parameter key value pair object.
Public Function Add(ByVal Name As String, ByVal Value As Variant) As Parameter
End Function

'Number of parameters in collection of parameters.
Public Function Count() As Long
End Function

'True if a parameter exists in the collection.
Public Function Exists(ByVal Name As String) As Boolean
End Function

'One Parameter object from the collection of Parameter objects.
Public Function Item(ByVal Key As Variant) As Parameter
End Function

'Delete a Parameter object from the collection of Parameter objects.
Public Sub Remove(ByVal Key As Variant)
End Sub
