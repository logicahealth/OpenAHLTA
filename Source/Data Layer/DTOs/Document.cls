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
Attribute VB_Name = "Document"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--- begin DTO Builder; DTO=Document
'--- 3/16/04 7:59:58 PM
'---
'--- NOTE: Any time you modify this class, increment the version number
'
'

Option Explicit

Implements iObjectPersist
Implements iState

Private Const OBJECT_VERSION = "1.0"
Private Const UNIQUE_IDENTIFIER = "Document"

Private mState As enumObjectState
Private mLastModifiedDate As Date   '
Private mOwnerName As String    '
Private mOwnerNCID As String
Private mHeader As String   '
Private mSection As String  '
Private mBody As String '
Private mTitle As String    '
Private mCategory As String '



Public Property Get LastModifiedDate() As Date
    LastModifiedDate = mLastModifiedDate
End Property
Public Property Let LastModifiedDate(value As Date)
    mLastModifiedDate = value
End Property
Public Property Get OwnerNCID() As String
    OwnerNCID = mOwnerNCID
End Property
Public Property Let OwnerNCID(value As String)
    mOwnerNCID = value
End Property
Public Property Get OwnerName() As String
    OwnerName = mOwnerName
End Property
Public Property Let OwnerName(value As String)
    mOwnerName = value
End Property
Public Property Get Header() As String
    Header = mHeader
End Property
Public Property Let Header(value As String)
    mHeader = value
End Property
Public Property Get Section() As String
    Section = mSection
End Property
Public Property Let Section(value As String)
    mSection = value
End Property
Public Property Get Body() As String
    Body = mBody
End Property
Public Property Let Body(value As String)
    mBody = value
End Property
Public Property Get Title() As String
    Title = mTitle
End Property
Public Property Let Title(value As String)
    mTitle = value
End Property
Public Property Get Category() As String
    Category = mCategory
End Property
Public Property Let Category(value As String)
    mCategory = value
End Property




Private Function iObjectPersist_Copy() As Object
    Dim NewCopy As Document
    Dim State As iState
    
    Set NewCopy = New Document
    
    With NewCopy
        .LastModifiedDate = mLastModifiedDate
        .OwnerName = mOwnerName
        .Header = mHeader
        .Section = mSection
        .Body = mBody
        .Title = mTitle
        .Category = mCategory
        
        'iState interface
        Set State = NewCopy
        State.State = mState
    End With

    Set iObjectPersist_Copy = NewCopy

End Function




Private Sub iObjectPersist_Persist(Writer As iWriter)
    With Writer
        .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION
        .WriteProperty "LastModifiedDate", mLastModifiedDate
        .WriteProperty "OwnerName", mOwnerName
        .WriteProperty "Header", mHeader
        .WriteProperty "Section", mSection
        .WriteProperty "Body", mBody
        .WriteProperty "Title", mTitle
        .WriteProperty "Category", mCategory
        
        'iState interface
        .WriteProperty "iState.State", mState
    End With
End Sub




Private Function iObjectPersist_Restore(Reader As iReader) As Boolean
Dim Failed As Boolean

    On Error GoTo ErrHandler
  
    With Reader
        mLastModifiedDate = .ReadProperty("LastModifiedDate")
        mOwnerName = .ReadProperty("OwnerName")
        mHeader = .ReadProperty("Header")
        mSection = .ReadProperty("Section")
        mBody = .ReadProperty("Body")
        mTitle = .ReadProperty("Title")
        mCategory = .ReadProperty("Category")
        
        'iState interface
        mState = .ReadProperty("iState.State")
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
'--- end DTO Builder

Private Property Let iState_State(ByVal NewValue As enumObjectState)
    mState = NewValue
End Property

Private Property Get iState_State() As enumObjectState
    iState_State = mState
End Property

Private Property Let iState_ID(ByVal RHS As String)

End Property

Private Property Get iState_ID() As String
    iState_ID = UNIQUE_IDENTIFIER
End Property
