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
Attribute VB_Name = "ConceptValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Respresentation of an NCID.
'NOTE: Any time you modify this class, increment the version number

Option Explicit

Implements iObjectPersist

Private Const UNIQUE_IDENTIFIER As String = "ConceptValue"
Private Const OBJECT_VERSION As String = "1.0"

Private mNCID As String                     'Unique Identifier
Private mRepresentation As String           'Text respresentation of the NCID

Private Function iObjectPersist_Copy() As Object
Dim Item As ConceptValue

  Set Item = New ConceptValue
  
  With Item
    .NCID = mNCID
    .Representation = mRepresentation
  End With
  
  Set iObjectPersist_Copy = Item
  
End Function

Private Sub iObjectPersist_Persist(Writer As ObjectPersist.iWriter)
  With Writer
    .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION
    
    .WriteProperty "NCID", mNCID
    .WriteProperty "Representation", mRepresentation
  End With
End Sub

Private Function iObjectPersist_Restore(Reader As ObjectPersist.iReader) As Boolean
Dim Failed As Boolean

  On Error GoTo ErrHandler
  
  With Reader
    mNCID = .ReadProperty("NCID")
    mRepresentation = .ReadProperty("Representation")
  End With

  iObjectPersist_Restore = Not Failed
  
  Exit Function
  
ErrHandler:
  Failed = True
  Resume Next

End Function

Private Property Get iObjectPersist_UniqueIdentifier() As String
  iObjectPersist_UniqueIdentifier = UNIQUE_IDENTIFIER
End Property

Private Property Get iObjectPersist_Version() As String
  iObjectPersist_Version = OBJECT_VERSION
End Property

Public Property Get NCID() As String
  NCID = mNCID
End Property

Public Property Let NCID(ByVal NewValue As String)
  mNCID = NewValue
End Property

Public Property Get Representation() As String
  Representation = mRepresentation
End Property

Public Property Let Representation(ByVal NewValue As String)
  mRepresentation = NewValue
End Property
