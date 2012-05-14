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
Attribute VB_Name = "FolderItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Application reference item in tree structure.
'NOTE: Any time you modify this class, increment the version number

Option Explicit

Implements iObjectPersist

Private Const UNIQUE_IDENTIFIER As String = "FolderItem"
Private Const OBJECT_VERSION As String = "1.0"

Private mApplicationNCID As String          'NCID of Application to launch
Private mImageName As String                'Image File Name
Private mIndex As Integer                   'Index in Tree
Private mKey As String                      'Key in Tree
Private mLaunchSequence As String           'Sequence of Application NCIDs.  These applications must be all be launched to get to the application this item represents.
Private mNodeText As String                 'Text to display in tree
Private mOrdinal As Integer                 'Position in tree relative to parent node
Private mParentOrdinal As Variant           'Parent Node's ordinal  (usually an interger but may be a string for external servers)
Private mProgramID As Long                  'Program ID of OLEServer.  Used to find the tree.node.
Private mSection As String                  'Parent Node
Private mTag As String                      'Stores the Menu ID of OLE Server in the GO menu

Private Function iObjectPersist_Copy() As Object
Dim Item As FolderItem

  Set Item = New FolderItem
  
  With Item
    .ApplicationNCID = mApplicationNCID
    .ImageName = mImageName
    .Index = mIndex
    .Key = mKey
    .LaunchSequence = mLaunchSequence
    .NodeText = mNodeText
    .Ordinal = mOrdinal
    .ParentOrdinal = mParentOrdinal
    .ProgramID = mProgramID
    .Section = mSection
    .Tag = mTag
  End With
  
  Set iObjectPersist_Copy = Item
  
End Function

Private Sub iObjectPersist_Persist(Writer As ObjectPersist.iWriter)
  With Writer
    .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION
    
    .WriteProperty "ApplicationNCID", mApplicationNCID
    .WriteProperty "ImageName", mImageName
    .WriteProperty "Index", mIndex
    .WriteProperty "Key", mKey
    .WriteProperty "LaunchSequence", mLaunchSequence
    .WriteProperty "NodeText", mNodeText
    .WriteProperty "Ordinal", mOrdinal
    .WriteProperty "ParentOrdinal", mParentOrdinal
    .WriteProperty "ProgramID", mProgramID
    .WriteProperty "Section", mSection
    .WriteProperty "Tag", mTag
  End With
End Sub

Private Function iObjectPersist_Restore(Reader As ObjectPersist.iReader) As Boolean
Dim Failed As Boolean

  On Error GoTo ErrHandler
  
  With Reader
    mApplicationNCID = .ReadProperty("ApplicationNCID")
    mImageName = .ReadProperty("ImageName")
    mIndex = .ReadProperty("Index")
    mKey = .ReadProperty("Key")
    mLaunchSequence = .ReadProperty("LaunchSequence")
    mNodeText = .ReadProperty("NodeText")
    mOrdinal = .ReadProperty("Ordinal")
    mParentOrdinal = .ReadProperty("ParentOrdinal")
    mProgramID = .ReadProperty("ProgramID")
    mSection = .ReadProperty("Section")
    mTag = .ReadProperty("Tag")
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

Public Property Get ApplicationNCID() As String
  ApplicationNCID = mApplicationNCID
End Property

Public Property Let ApplicationNCID(ByVal NewValue As String)
  mApplicationNCID = NewValue
End Property

Public Property Get ImageName() As String
  ImageName = mImageName
End Property

Public Property Let ImageName(ByVal NewValue As String)
  mImageName = NewValue
End Property

Public Property Get Index() As Integer
  Index = mIndex
End Property

Public Property Let Index(ByVal NewValue As Integer)
  mIndex = NewValue
End Property

Public Property Get Key() As String
  Key = mKey
End Property

Public Property Let Key(ByVal NewValue As String)
  mKey = NewValue
End Property

Public Property Get LaunchSequence() As String
  LaunchSequence = mLaunchSequence
End Property

Public Property Let LaunchSequence(ByVal NewValue As String)
  mLaunchSequence = NewValue
End Property

Public Property Get NodeText() As String
  NodeText = mNodeText
End Property

Public Property Let NodeText(ByVal NewValue As String)
  mNodeText = NewValue
End Property

Public Property Get Ordinal() As Integer
  Ordinal = mOrdinal
End Property

Public Property Let Ordinal(ByVal NewValue As Integer)
  mOrdinal = NewValue
End Property

Public Property Get ParentOrdinal() As Variant
  ParentOrdinal = mParentOrdinal
End Property

Public Property Let ParentOrdinal(ByVal NewValue As Variant)
  mParentOrdinal = NewValue
End Property

Public Property Get ProgramID() As Long
  ProgramID = mProgramID
End Property

Public Property Let ProgramID(ByVal NewValue As Long)
  mProgramID = NewValue
End Property

Public Property Get Section() As String
  Section = mSection
End Property

Public Property Let Section(ByVal NewValue As String)
  mSection = NewValue
End Property

Public Property Get Tag() As String
  Tag = mTag
End Property

Public Property Let Tag(ByVal NewValue As String)
  mTag = NewValue
End Property

