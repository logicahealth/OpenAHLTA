VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReadXMLFindStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Implements the Find Node algorithm for a property value in the XML document.
'This algorithm is used when the document versions are NOT the same.

Option Explicit

Implements iReadStrategy

Private mXMLDoc As ChilkatXml       'XML parser

Private Sub iReadStrategy_Init(XMLDoc As CHILKATXMLLib.IChilkatXml)
  Set mXMLDoc = XMLDoc
End Sub

Private Property Get iReadStrategy_Mode() As enumReadMode
  iReadStrategy_Mode = rmFind
End Property

Private Function iReadStrategy_ReadProperty(ByVal Name As String) As String
  iReadStrategy_ReadProperty = mXMLDoc.FindChild(Name).Content
End Function
