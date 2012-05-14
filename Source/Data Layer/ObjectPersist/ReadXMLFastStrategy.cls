VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReadXMLFastSrtategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Implements the Read Node algorithm for a property value in the XML document.
'This algorithm is used when the document versions are the same.  It assumes
'that the property value are in the correct order.  This is super fast.

Option Explicit

Implements iReadStrategy

Private mXMLDoc As ChilkatXml           'XML parser

Private Sub iReadStrategy_Init(XMLDoc As CHILKATXMLLib.IChilkatXml)
  Set mXMLDoc = XMLDoc
End Sub

Private Property Get iReadStrategy_Mode() As enumReadMode
  iReadStrategy_Mode = rmFast
End Property

Private Function iReadStrategy_ReadProperty(ByVal Name As String) As String
  With mXMLDoc
    iReadStrategy_ReadProperty = .Content
    
    .NextSibling2
  End With
End Function
