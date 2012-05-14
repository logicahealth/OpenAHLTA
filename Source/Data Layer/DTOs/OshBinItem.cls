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
Attribute VB_Name = "OshBinItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Configuration value.
'NOTE: Any time you modify this class, increment the version number

Option Explicit

Implements iObjectPersist

Private Const UNIQUE_IDENTIFIER As String = "OshBinItem"
Private Const OBJECT_VERSION As String = "1.0"

Private mApplicationNCID As Variant                       'NCID of application saving config value
Private mBinNCID As String                                'NCID of Bin value if found in (Enterprise, Facility, Role, User)
Private mDelete As Boolean                                'True if value should be deleted
Private mOwnerNCID As String                              'Whom the config value is for
Private mSection As String                                'Applications/Owner pairs can save values in multiple sections
Private mKey As String                                    'Key for value
Private mValue As Variant                                 'Value

Private Function iObjectPersist_Copy() As Object
Dim Item As OshBinItem

  Set Item = New OshBinItem
  
  With Item
    .ApplicationNCID = mApplicationNCID
    .BinNCID = mBinNCID
    
    .Key = mKey
    .OwnerNCID = mOwnerNCID
    .Section = mSection
    .Value = mValue
  End With
  
  Set iObjectPersist_Copy = Item
  
End Function

Private Sub iObjectPersist_Persist(Writer As ObjectPersist.iWriter)
  With Writer
    .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION
    
    .WriteProperty "ApplicationNCID", mApplicationNCID
    .WriteProperty "BinNCID", mBinNCID
    .WriteProperty "Delete", mDelete
    .WriteProperty "Key", mKey
    .WriteProperty "OwnerNCID", mOwnerNCID
    .WriteProperty "Section", mSection
    .WriteProperty "Value", mValue
  End With
End Sub

Private Function iObjectPersist_Restore(Reader As ObjectPersist.iReader) As Boolean
Dim Failed As Boolean

  On Error GoTo ErrHandler
  
  With Reader
    mApplicationNCID = .ReadProperty("ApplicationNCID")
    mBinNCID = .ReadProperty("BinNCID")
    mDelete = .ReadProperty("Delete")
    mKey = .ReadProperty("Key")
    mOwnerNCID = .ReadProperty("OwnerNCID")
    mSection = .ReadProperty("Section")
    mValue = .ReadProperty("Value")
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

Public Property Get ApplicationNCID() As Variant
  ApplicationNCID = mApplicationNCID
End Property

Public Property Let BinNCID(ByVal NewValue As String)
  mBinNCID = NewValue
End Property

Public Property Get BinNCID() As String
  BinNCID = mBinNCID
End Property

Public Property Let Delete(ByVal NewValue As Boolean)
  mDelete = NewValue
End Property

Public Property Get Delete() As Boolean
  Delete = mDelete
End Property

Public Property Let ApplicationNCID(ByVal NewValue As Variant)
  mApplicationNCID = NewValue
End Property

Public Property Get OwnerNCID() As String
  OwnerNCID = mOwnerNCID
End Property

Public Property Let OwnerNCID(ByVal NewValue As String)
  mOwnerNCID = NewValue
End Property

Public Property Get Section() As String
  Section = mSection
End Property

Public Property Let Section(ByVal NewValue As String)
  mSection = NewValue
End Property

Public Property Get Key() As String
  Key = mKey
End Property

Public Property Let Key(ByVal NewValue As String)
  mKey = NewValue
End Property

Public Property Get Value() As Variant
  Value = mValue
End Property

Public Property Let Value(ByVal NewValue As Variant)
  mValue = NewValue
End Property

