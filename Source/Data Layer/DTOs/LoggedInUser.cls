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
Attribute VB_Name = "LoggedInUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Describes logged in user.
'NOTE Any time you modify this class, increment the version number

Option Explicit

Implements iObjectPersist

Private Const UNIQUE_IDENTIFIER As String = "LoggedInUser"
Private Const OBJECT_VERSION As String = "1.0"

Private mRoleNCID As String
Private mRoleName As String
Private mUserNCID As String
Private mUserName As String

Private Function iObjectPersist_Copy() As Object
Dim NewMe As LoggedInUser
Dim Persist As iObjectPersist
Dim Writer As iWriter
Dim Reader As iReader

  Set NewMe = New LoggedInUser
  
  Set Writer = New WriterXML
  iObjectPersist_Persist Writer
  
  Set Reader = New ReaderXML
  Set Persist = NewMe
  
  Reader.DocumentVersion = Persist.Version
  Reader.Contents = Writer.Contents
  
  Persist.Restore Reader
  
  Set iObjectPersist_Copy = NewMe
  
End Function

Private Sub iObjectPersist_Persist(Writer As ObjectPersist.iWriter)
  With Writer
    .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION
    
    .WriteProperty "RoleNCID", mRoleNCID
    .WriteProperty "RoleName", mRoleName
    .WriteProperty "UserNCID", mUserNCID
    .WriteProperty "UserName", mUserName
  End With
End Sub

Private Function iObjectPersist_Restore(Reader As ObjectPersist.iReader) As Boolean
Dim Failed As Boolean

  On Error GoTo ErrHandler
  
  With Reader
    mRoleNCID = .ReadProperty("RoleNCID")
    mRoleName = .ReadProperty("RoleName")
    mUserNCID = .ReadProperty("UserNCID")
    mUserName = .ReadProperty("UserName")
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

Public Property Get RoleNCID() As String
  RoleNCID = mRoleNCID
End Property

Public Property Let RoleNCID(ByVal NewValue As String)
  mRoleNCID = NewValue
End Property

Public Property Get RoleName() As String
  RoleName = mRoleName
End Property

Public Property Let RoleName(ByVal NewValue As String)
  mRoleName = NewValue
End Property

Public Property Get UserNCID() As String
  UserNCID = mUserNCID
End Property

Public Property Let UserNCID(ByVal NewValue As String)
  mUserNCID = NewValue
End Property

Public Property Get UserName() As String
  UserName = mUserName
End Property

Public Property Let UserName(ByVal NewValue As String)
  mUserName = NewValue
End Property

