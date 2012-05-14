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
Attribute VB_Name = "EncProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--- begin DTO Builder; DTO=EncProvider
'--- 3/16/04 7:51:12 PM
'---
'--- NOTE: Any time you modify this class, increment the version number
'
'

Option Explicit

Implements iObjectPersist
Implements iState

Private Const OBJECT_VERSION = "1.0"
Private Const UNIQUE_IDENTIFIER = "EncProvider"

Private mState As enumObjectState
Private mIEN As String  '
Private mNCID As Long   '
Private mFullName As String '
Private mFacilityNCID As String '
Private mFacility As String '
Private mRoleNCID As Long   '
Private mRole As String '
Private mUnitNumber As String   '
Private mSigLine2 As String '
Private mSigLine3 As String '
Private mProcedures As EncProvProcedures    '
Private mSigned As Date '
Private mPrevNCID As String

Public Property Get PrevNCID() As Long
    PrevNCID = mPrevNCID
End Property
Public Property Let PrevNCID(value As Long)
    mPrevNCID = value
End Property

Public Property Get IEN() As String
    IEN = mIEN
End Property
Public Property Let IEN(value As String)
    mIEN = value
End Property
Public Property Get NCID() As Long
    NCID = mNCID
End Property
Public Property Let NCID(value As Long)
    mNCID = value
End Property
Public Property Get FullName() As String
    FullName = mFullName
End Property
Public Property Let FullName(value As String)
    mFullName = value
End Property
Public Property Get FacilityNCID() As String
    FacilityNCID = mFacilityNCID
End Property
Public Property Let FacilityNCID(value As String)
    mFacilityNCID = value
End Property
Public Property Get Facility() As String
    Facility = mFacility
End Property
Public Property Let Facility(value As String)
    mFacility = value
End Property
Public Property Get RoleNCID() As Long
    RoleNCID = mRoleNCID
End Property
Public Property Let RoleNCID(value As Long)
    mRoleNCID = value
End Property
Public Property Get Role() As String
    Role = mRole
End Property
Public Property Let Role(value As String)
    mRole = value
End Property
Public Property Get UnitNumber() As String
    UnitNumber = mUnitNumber
End Property
Public Property Let UnitNumber(value As String)
    mUnitNumber = value
End Property
Public Property Get SigLine2() As String
    SigLine2 = mSigLine2
End Property
Public Property Let SigLine2(value As String)
    mSigLine2 = value
End Property
Public Property Get SigLine3() As String
    SigLine3 = mSigLine3
End Property
Public Property Let SigLine3(value As String)
    mSigLine3 = value
End Property
Public Property Get Procedures() As EncProvProcedures
    If mProcedures Is Nothing Then Set mProcedures = New EncProvProcedures
    Set Procedures = mProcedures
End Property
Public Property Set Procedures(value As EncProvProcedures)
    Set mProcedures = value
End Property
Public Property Get Signed() As Date
    Signed = mSigned
End Property
Public Property Let Signed(value As Date)
    mSigned = value
End Property




Private Function iObjectPersist_Copy() As Object
    Dim NewCopy As EncProvider
    Dim State As iState
    
    Set NewCopy = New EncProvider
    
    With NewCopy
        .IEN = mIEN
        .NCID = mNCID
        .FullName = mFullName
        .FacilityNCID = mFacilityNCID
        .Facility = mFacility
        .RoleNCID = mRoleNCID
        .Role = mRole
        .UnitNumber = mUnitNumber
        .SigLine2 = mSigLine2
        .SigLine3 = mSigLine3
        Set .Procedures = mProcedures
        .Signed = mSigned
        .PrevNCID = mPrevNCID
        
        'iState interface
        Set State = NewCopy
        State.State = mState
    End With

    Set iObjectPersist_Copy = NewCopy

End Function




Private Sub iObjectPersist_Persist(Writer As iWriter)
    With Writer
        .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION
        .WriteProperty "IEN", mIEN
        .WriteProperty "NCID", mNCID
        .WriteProperty "FullName", mFullName
        .WriteProperty "FacilityNCID", mFacilityNCID
        .WriteProperty "Facility", mFacility
        .WriteProperty "RoleNCID", mRoleNCID
        .WriteProperty "Role", mRole
        .WriteProperty "UnitNumber", mUnitNumber
        .WriteProperty "SigLine2", mSigLine2
        .WriteProperty "SigLine3", mSigLine3
        .WriteObject "Procedures", mProcedures
        .WriteProperty "Signed", mSigned
        .WriteProperty "PrevNCID", mPrevNCID
        
        'iState interface
        .WriteProperty "iState.State", mState
    End With
End Sub




Private Function iObjectPersist_Restore(Reader As iReader) As Boolean
Dim Failed As Boolean

    On Error GoTo ErrHandler
  
    With Reader
        mIEN = .ReadProperty("IEN")
        mNCID = .ReadProperty("NCID")
        mFullName = .ReadProperty("FullName")
        mFacilityNCID = .ReadProperty("FacilityNCID")
        mFacility = .ReadProperty("Facility")
        mRoleNCID = .ReadProperty("RoleNCID")
        mRole = .ReadProperty("Role")
        mUnitNumber = .ReadProperty("UnitNumber")
        mSigLine2 = .ReadProperty("SigLine2")
        mSigLine3 = .ReadProperty("SigLine3")
        .ReadObject "Procedures", Procedures
        mSigned = .ReadProperty("Signed")
        mPrevNCID = .ReadProperty("PrevNCID")
        
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
