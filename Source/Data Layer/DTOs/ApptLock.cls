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
Attribute VB_Name = "ApptLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--- begin DTO Builder; DTO=ApptLock
'--- 3/18/2004 11:45:23 AM
'---
'--- NOTE: Any time you modify this class, increment the version number
'
'

Option Explicit

Implements iObjectPersist

Private Const OBJECT_VERSION = "1.0"
Private Const UNIQUE_IDENTIFIER = "ApptLock"


Private mAppointmentId As String    '
Private mDateTime As Date   '
Private mUserNCID As String '
Private mMachineName As String  '
Private mTypeLock As String '
Private mStatus As String   '



Public Property Get AppointmentId() As String
    AppointmentId = mAppointmentId
End Property
Public Property Let AppointmentId(ByVal value As String)
    mAppointmentId = value
End Property
Public Property Get DateTime() As Date
    DateTime = mDateTime
End Property
Public Property Let DateTime(ByVal value As Date)
    mDateTime = value
End Property
Public Property Get UserNCID() As String
    UserNCID = mUserNCID
End Property
Public Property Let UserNCID(ByVal value As String)
    mUserNCID = value
End Property
Public Property Get MachineName() As String
    MachineName = mMachineName
End Property
Public Property Let MachineName(ByVal value As String)
    mMachineName = value
End Property
Public Property Get TypeLock() As String
    TypeLock = mTypeLock
End Property
Public Property Let TypeLock(ByVal value As String)
    mTypeLock = value
End Property
Public Property Get Status() As String
    Status = mStatus
End Property
Public Property Let Status(ByVal value As String)
    mStatus = value
End Property




Private Function iObjectPersist_Copy() As Object
    Dim NewCopy As ApptLock
    Set NewCopy = New ApptLock
    
    With NewCopy
        .AppointmentId = mAppointmentId
        .DateTime = mDateTime
        .UserNCID = mUserNCID
        .MachineName = mMachineName
        .TypeLock = mTypeLock
        .Status = mStatus
    End With

    Set iObjectPersist_Copy = NewCopy

End Function




Private Sub iObjectPersist_Persist(Writer As iWriter)
    With Writer
        .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION
        .WriteProperty "AppointmentId", mAppointmentId
        .WriteProperty "DateTime", mDateTime
        .WriteProperty "UserNCID", mUserNCID
        .WriteProperty "MachineName", mMachineName
        .WriteProperty "TypeLock", mTypeLock
        .WriteProperty "Status", mStatus
    End With
End Sub




Private Function iObjectPersist_Restore(Reader As iReader) As Boolean
Dim Failed As Boolean

    On Error GoTo ErrHandler
  
    With Reader
        mAppointmentId = .ReadProperty("AppointmentId")
        mDateTime = .ReadProperty("DateTime")
        mUserNCID = .ReadProperty("UserNCID")
        mMachineName = .ReadProperty("MachineName")
        mTypeLock = .ReadProperty("TypeLock")
        mStatus = .ReadProperty("Status")
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

