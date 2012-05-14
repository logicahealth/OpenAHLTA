VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WriterXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Creates an XML document.  It uses Chilkat XML parse to create the XML string.  This parse was used because it is much faster
'than Microsoft's XML parser.

Option Explicit

Implements iWriter

Private mXMLDoc As ChilkatXml       'XML parser

Private Sub Class_Initialize()
  Set mXMLDoc = New ChilkatXml
End Sub

'Loads XML into parser.
Private Property Let iWriter_Contents(RHS As String)
  mXMLDoc.LoadXml RHS
End Property

'Gets XML from parser.
Private Property Get iWriter_Contents() As String
  iWriter_Contents = mXMLDoc.GetXml
End Property

'Used to create save document name and version in the main document or a sub document.  Each object saved is a document so
'one document may contain several sub documents.  The count property is used to saved the number of object persisted in a
'collection class object.
Private Sub iWriter_InitHeader(ByVal Name As String, Optional ByVal Version As String = vbNullString, Optional ByVal Count As Long = -1&)

  On Error GoTo ErrHandler
  
  With mXMLDoc
    .Tag = Name
    
    If Len(Version) > 0 Then .AddAttribute "Version", Version
    
    If Count > -1 Then .AddAttribute "CollectionCount", Count
  End With
  
  Exit Sub
  
ErrHandler:
  
End Sub

'Saves an array of simple data types as a sub document (node with childen).
Private Sub iWriter_WriteArray(ByVal Name As String, Ary As Variant)
Dim Writer As iWriter
Dim xmlwriter As WriterXML
Dim i As Long

  On Error Resume Next
    
  Set Writer = New WriterXML
    
  With Writer
    .InitHeader Name
    
    If IsArray(Ary) Then
      For i = 0 To UBound(Ary)
        .WriteProperty "Item", Ary(i)
      Next
    End If
  End With
  
  Set xmlwriter = Writer
  
  mXMLDoc.AddChildTree xmlwriter.XMLDoc

End Sub

'Saves a collection of simple data types as a sub document (node with childen).
Private Sub iWriter_WriteCollection(ByVal Name As String, col As Collection)
Dim Writer As iWriter
Dim xmlwriter As WriterXML
Dim i As Long

  On Error Resume Next
  
  Set Writer = New WriterXML
    
  With Writer
    .InitHeader Name
    
    For i = 1 To col.Count
      .WriteProperty "Item", col.Item(i)
    Next
  End With
  
  Set xmlwriter = Writer
  
  mXMLDoc.AddChildTree xmlwriter.XMLDoc

End Sub

'Has the object persist itself to a XML string then add the XML string to the main document
'as a sub doucment (node with children).
Private Sub iWriter_WriteObject(ByVal Name As String, Obj As Object)
Dim Writer As iWriter
Dim xmlwriter As WriterXML
Dim Persist As iObjectPersist
Dim XMLDoc As ChilkatXml

  On Error Resume Next
  
  Set XMLDoc = mXMLDoc.NewChild(Name, vbNullString)
  
  If Obj Is Nothing Then Exit Sub
    
  Set Writer = New WriterXML
  Writer.InitHeader Name
    
  Set Persist = Obj
  Persist.Persist Writer
  
  Set xmlwriter = Writer
  
  'mXMLDoc.AddChildTree xmlwriter.XMLDoc
  XMLDoc.AddChildTree xmlwriter.XMLDoc
  
End Sub

'Saves a new node to the XML string (no children).
Private Sub iWriter_WriteProperty(ByVal Name As String, ByVal Value As Variant)
  
  On Error GoTo ErrHandler
  
  If Len(Value) > 5 Then
    mXMLDoc.Cdata = 1
  End If
  
  mXMLDoc.NewChild2 Name, Value
  
  Exit Sub
  
ErrHandler:
  If Len(Value) = 0 Then
    mXMLDoc.NewChild Name, vbNullString       'Add blank values
  End If

End Sub

Friend Property Get XMLDoc() As ChilkatXml
  Set XMLDoc = mXMLDoc
End Property

Public Sub Save(ByVal FileName As String, Optional ByVal Encrypt As Boolean = True)
  With mXMLDoc
    .Encoding = "ASCII"
    .Standalone = "yes"
    
    If Encrypt Then
      .GetRoot2
      .CompressSubtree
      .EncryptNode "JiGEzqbG9zB+ILJb5TOgBlHWjSK/72+5xylfTqWI7eFguGOemWbInBO7ZBTnK6xX6Y4uKR6v"
    End If
    
    .SaveXml FileName
  End With
End Sub

