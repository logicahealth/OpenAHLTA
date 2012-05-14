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
Attribute VB_Name = "ReaderXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Reads an XML document.  It uses Chilkat XML parse to read the XML string.  This parse was used because it is much faster
'than Microsoft's XML parser.

Option Explicit

Implements iReader

Private mXMLDoc As ChilkatXml             'XML parser

Private mCount As Long                    'Number of like sub documents
Private mDocumentVersion As String        'Version of document

Private mReadStrategy As iReadStrategy    'Strategy used to read property values of the XML document

Private Property Get iReader_CollectionCount() As Long
  iReader_CollectionCount = mCount
End Property

'Init parser and this object.
Private Property Let iReader_Contents(RHS As String)
  If mXMLDoc Is Nothing Then Set mXMLDoc = New ChilkatXml
  
  mXMLDoc.LoadXml RHS
  Init
End Property

Private Property Get iReader_Contents() As String
  On Error Resume Next
  iReader_Contents = mXMLDoc.GetXml
End Property

Private Property Let iReader_DocumentVersion(ByVal RHS As String)
  mDocumentVersion = RHS
End Property

Private Property Get iReader_DocumentVersion() As String
  iReader_DocumentVersion = mDocumentVersion
End Property

'Creates an array from a sub document (node with children).
'Uses a fast read method.
Private Function iReader_ReadArray(ByVal Name As String) As Variant
Dim Success As Long
Dim Ary() As String
Dim i As Long
Dim Failed As Boolean
Dim ErrDesc As String

  On Error GoTo ErrHandler
  
  Success = InitDocPointer(Name)            'Sets parser to 1st child of document
  
  Do While Success = 1
    ReDim Ary(i)
    
    Ary(UBound(Ary)) = mXMLDoc.Content
    
    Success = mXMLDoc.NextSibling2
    
    If Success Then i = i + 1
  Loop

  AdvanceDocPointer Name                    'Sets parser to next sibling of the sub document (array)

  iReader_ReadArray = Ary
  
  If Failed Then Err.Raise 5000, "ReaderXML.iReader_ReadArray", "Error restoring array.  " & ErrDesc
  
  Exit Function
  
ErrHandler:
  Failed = True
  ErrDesc = Err.Description
  Resume Next
  
End Function

'Creates a collection from a sub document (node with children).
'Uses a fast read method.
Private Function iReader_ReadCollection(ByVal Name As String) As Collection
Dim Success As Long
Dim col As Collection
Dim Failed As Boolean
Dim ErrDesc As String

  On Error GoTo ErrHandler
  
  If mXMLDoc.NumChildren = 0 Then
    Exit Function
  End If
  
  Set col = New Collection
  
  Success = InitDocPointer(Name)            'Sets parser to 1st child of document
  
  Do While Success = 1
    col.Add mXMLDoc.Content
    
    Success = mXMLDoc.NextSibling2
  Loop

  AdvanceDocPointer Name                    'Sets parser to next sibling of the sub document (collection)
  
  Set iReader_ReadCollection = col
  
  If Failed Then Err.Raise 5001, "ReaderXML.iReader_ReadCollection", "Error restoring collection.  " & ErrDesc
  
  Exit Function
  
ErrHandler:
  Failed = True
  ErrDesc = Err.Description
  Resume Next
  
End Function


'Restores an object from the XML document.  Each object is its own sub document.
'The read strategy (Mode) is determined based on the version object and the mode of this writer.
'If the XML doucment and the object have the same version and the parent node is using
'the fast read strategy then this object is read using the fast strategy else it is read by a new reader object.
Private Sub iReader_ReadObject(ByVal Name As String, Obj As Object)
Dim Mode As enumReadMode

  On Error GoTo ErrHandler
  
  If mReadStrategy.Mode = rmFast Then InitDocPointer Name                         'Sets parser to 1st child of document
  
  If mXMLDoc.NumChildren = 0 Then
    AdvanceDocPointer Name                    'Sets parser to next sibling of the sub document (collection)
    Exit Sub
  End If
    
  Mode = DetermineDocumentMode(Obj)
      
  If Mode = rmFast Then
    RestoreObjectLocal Obj
  Else
    RestoreObjectNewReader Name, Obj
  End If

  If mReadStrategy.Mode = rmFast Then AdvanceDocPointer Name                    'Sets parser to next sibling of the sub document (collection)

  Exit Sub
  
ErrHandler:
  Err.Raise Err.Number, "ReaderXML.iReader_ReadObject", Err.Description
  
End Sub

Private Function iReader_ReadProperty(ByVal Name As String) As String
  iReader_ReadProperty = mReadStrategy.ReadProperty(Name)
End Function

Friend Property Get XMLDoc() As ChilkatXml
  Set XMLDoc = mXMLDoc
End Property

'Init parser and this object.  This method is used by other ReaderXML objects to init a new ReaderXML object.  It is
'much faster to do it this way then via the contents property.
Friend Property Set XMLDoc(ByRef Doc As ChilkatXml)
  Set mXMLDoc = Doc
  Init
End Property

'Used to advance the parsers pointer to the next sibling node.
Private Sub AdvanceDocPointer(ByVal Name As String)

  If mXMLDoc.Tag <> Name Then mXMLDoc.GetParent2      'Move back up the tree if needed
  
  If mReadStrategy.Mode = rmFast Then
    mXMLDoc.NextSibling2                              'Fast mode order is assumed to be correct so move next
  Else
    If mXMLDoc.Tag = Name Then mXMLDoc.GetParent2     'Move back up the tree if needed
  End If

End Sub

'If the current writer object is using a fast read strategy then the object is checked to determine if the version of the
'object matches what is in the XML document.  If the current writer is using the find read strategy then the find read
'strategy is always returned.
Private Function DetermineDocumentMode(ByRef Persist As iObjectPersist) As enumReadMode

  On Error GoTo ErrHandler
  
  If mReadStrategy.Mode = rmFast Then
    DetermineDocumentMode = DetermineMode(Persist.Version, mXMLDoc)
  Else
    DetermineDocumentMode = rmFind
  End If

  Exit Function
  
ErrHandler:
  DetermineDocumentMode = rmFind
  
End Function

'If the document version number matches the object version number then return fast mode.
'If the document version number does not match the object version number then return find mode.
Private Function DetermineMode(ByVal CurrentVersion As String, ByRef XMLDoc As ChilkatXml) As enumReadMode
  
  On Error GoTo ErrHandler
  
  If XMLDoc.GetAttributeValue(0) = CurrentVersion Then
    DetermineMode = rmFast
  Else
    DetermineMode = rmFind
  End If
  
  Exit Function
  
ErrHandler:
  DetermineMode = rmFind
  
End Function

'Inits the Reader object
Private Sub Init()
Dim Mode As enumReadMode

  On Error Resume Next
  
  'Init count
  If mXMLDoc.NumAttributes >= 2 Then
    mCount = mXMLDoc.GetAttributeValue(1)
  Else
    mCount = 1
  End If
  
  'Init read strategy
  Mode = DetermineMode(iReader_DocumentVersion, mXMLDoc)
  
  If Mode = rmFast Then
    Set mReadStrategy = New ReadXMLFastSrtategy
  Else
    Set mReadStrategy = New ReadXMLFindStrategy
  End If
      
  mReadStrategy.Init mXMLDoc
  
  'Init parser if needed
  If mReadStrategy.Mode = rmFast Then
    mXMLDoc.FirstChild2                   'In fast mode order is assumed
  End If
  
End Sub

'Based on read strategy move the parser pointer to the first child node of the document.
'Parser returns a 1 on a successful call and a 0 on failures
Private Function InitDocPointer(ByVal Name As String) As Long
 
  If mReadStrategy.Mode = rmFast Then
    InitDocPointer = mXMLDoc.FirstChild2                'Fast mode order is assumed to be correct
  Else
    InitDocPointer = mXMLDoc.FindChild2(Name)
    If InitDocPointer Then InitDocPointer = mXMLDoc.FirstChild2
  End If

End Function

Public Sub Load(ByVal FileName As String)
  If mXMLDoc Is Nothing Then Set mXMLDoc = New ChilkatXml
  
  With mXMLDoc
    .LoadXmlFile FileName
    
    .DecryptNode "JiGEzqbG9zB+ILJb5TOgBlHWjSK/72+5xylfTqWI7eFguGOemWbInBO7ZBTnK6xX6Y4uKR6v"
    .DecompressSubtree
  End With
  
  Init
End Sub

'Restores object and move to the next sibling node.
Private Sub RestoreObjectLocal(ByRef Persist As iObjectPersist)
Dim Success As Long
Dim Reader As iReader
Dim XMLReader As ReaderXML

  On Error Resume Next
  
  Set Reader = New ReaderXML
  Reader.DocumentVersion = Persist.Version

  Set XMLReader = Reader
  Set XMLReader.XMLDoc = mXMLDoc
  
  Persist.Restore Reader
    
  With mXMLDoc
    .GetParent2
    .NextSibling2
  End With
  
End Sub

'Create a new Reader and use the new new Reader object to restore the object.
Private Sub RestoreObjectNewReader(ByVal Name As String, ByRef Persist As iObjectPersist)
Dim Reader As iReader
Dim XMLReader As ReaderXML
    
  On Error GoTo ErrHandler
  
  Set Reader = New ReaderXML
  Reader.DocumentVersion = Persist.Version

  Set XMLReader = Reader

  'Init Reader object
  If mReadStrategy.Mode = rmFast Then
    Set XMLReader.XMLDoc = mXMLDoc
  Else
    Set XMLReader.XMLDoc = mXMLDoc.FindChild(Name)
    
    If XMLReader.XMLDoc.NumChildren > 0 Then
      Set XMLReader.XMLDoc = XMLReader.XMLDoc.FindChild(Persist.UniqueIdentifier)
    Else
      Exit Sub
    End If
  End If

  On Error Resume Next
  
  Persist.Restore Reader

  If mReadStrategy.Mode = rmFast Then
    mXMLDoc.NextSibling2                    'In fast mode order is assumed
  End If

  Exit Sub
  
ErrHandler:
  Err.Raise Err.Number, "ReaderXML.RestoreObjectNewReader", Err.Description
  
End Sub
