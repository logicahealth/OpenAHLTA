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
Attribute VB_Name = "iList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Abstract interface that can be implemented by classes to support navigation of set of data with exposing how the data is
'saved interally.

Option Explicit

'True if pointer to the current data is before the begining of file (BOF).
Public Property Get BOF() As Boolean
End Property

'True if pointer to the current data is after the end of file (EOF).
Public Property Get EOF() As Boolean
End Property

'Returns a row of data.
Public Property Get Fields() As iParameters
End Property

'Returns the number of rows of data.
Public Property Get RecordCount() As Long
End Property

'Retuns a value from the current row of data.
Public Property Get Value(ByVal FieldName As String) As Variant
End Property

'Adds a new row of data.
Public Sub AddRow(ByRef Parameters As iParameters)
End Sub

'Moves pointer to the first row of data.
Public Sub MoveFirst()
End Sub

'Moves pointer to the last row of data.
Public Sub MoveLast()
End Sub

'Moves pointer to the next row of data.
Public Sub MoveNext()
End Sub

'Moves pointer to the previous row of data.
Public Sub MovePrevious()
End Sub

'Delete the current row of data.
Public Sub RemoveCurrentRow()
End Sub
