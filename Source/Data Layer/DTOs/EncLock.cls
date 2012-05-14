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
Attribute VB_Name = "EncLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--- begin DTO Builder; DTO=EncLock
'--- 3/18/2004 4:38:44 PM
'---
'--- NOTE: Any time you modify this class, increment the version number
'
'  Provides info on who has encounter section locked.

Option Explicit

Implements iObjectPersist

Private Const OBJECT_VERSION = "1.0"
Private Const UNIQUE_IDENTIFIER = "EncLock"

Public Enum LockReasonEnum
    SectionLock = 1
    SignLock = 2
    CosignLock = 3
    AmendLock = 4
    AppendLock = 5
End Enum

Private mOwnerNCID As String    '
Private mSection As EncounterSectionEnum    '
Private mWorkstation As String  '
Private mLockDate As Date   '
Private mLockStatus As LockReasonEnum
Private mEventID As Long
Private mUnbreakable As Boolean


Public Property Get Unbreakable() As Boolean
    Unbreakable = mUnbreakable
End Property
Public Property Let Unbreakable(ByVal value As Boolean)
    mUnbreakable = value
End Property

Public Property Get EventID() As Long
    EventID = mEventID
End Property
Public Property Let EventID(ByVal value As Long)
    mEventID = value
End Property

Public Property Get LockStatus() As LockReasonEnum
    LockStatus = mLockStatus
End Property
Public Property Let LockStatus(ByVal value As LockReasonEnum)
    mLockStatus = value
End Property

Public Property Get OwnerNCID() As String
    OwnerNCID = mOwnerNCID
End Property
Public Property Let OwnerNCID(ByVal value As String)
    mOwnerNCID = value
End Property
Public Property Get Section() As EncounterSectionEnum
    Section = mSection
End Property
Public Property Let Section(ByRef value As EncounterSectionEnum)
    mSection = value
End Property
Public Property Get Workstation() As String
    Workstation = mWorkstation
End Property
Public Property Let Workstation(ByVal value As String)
    mWorkstation = value
End Property
Public Property Get LockDate() As Date
    LockDate = mLockDate
End Property
Public Property Let LockDate(ByVal value As Date)
    mLockDate = value
End Property




Private Function iObjectPersist_Copy() As Object
    Dim NewCopy As EncLock
    Set NewCopy = New EncLock
    
    With NewCopy
        .OwnerNCID = mOwnerNCID
        .Section = mSection
        .Workstation = mWorkstation
        .LockDate = mLockDate
        .LockStatus = mLockStatus
        .EventID = mEventID
        .Unbreakable = mUnbreakable
    End With

    Set iObjectPersist_Copy = NewCopy

End Function




Private Sub iObjectPersist_Persist(Writer As iWriter)
    With Writer
        .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION
        .WriteProperty "OwnerNCID", mOwnerNCID
        .WriteProperty "Section", mSection
        .WriteProperty "Workstation", mWorkstation
        .WriteProperty "LockDate", mLockDate
        .WriteProperty "LockStatus", mLockStatus
        .WriteProperty "EventID", mEventID
        .WriteArray "Unbreakable", mUnbreakable
    End With
End Sub




Private Function iObjectPersist_Restore(Reader As iReader) As Boolean
Dim Failed As Boolean

    On Error GoTo ErrHandler
  
    With Reader
        mOwnerNCID = .ReadProperty("OwnerNCID")
        mSection = .ReadProperty("Section")
        mWorkstation = .ReadProperty("Workstation")
        mLockDate = .ReadProperty("LockDate")
        mLockStatus = .ReadProperty("LockStatus")
        mEventID = .ReadProperty("EventID")
        mUnbreakable = .ReadProperty("Unbreakable")
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

