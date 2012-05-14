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
Attribute VB_Name = "iReadStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Abstract interface used to abstract the read algorithm used.

Option Explicit

'Read algorithms
Public Enum enumReadMode
  rmFast = 0
  rmFind = 1
End Enum

'Returns the algorithms type of this object.
Public Property Get Mode() As enumReadMode
End Property

'Passes in reference to 3rd party XML parser.
Public Sub Init(ByRef XMLDoc As ChilkatXml)
End Sub

'Reads a specific node value from the XML document.
Public Function ReadProperty(ByVal Name As String) As String
End Function
