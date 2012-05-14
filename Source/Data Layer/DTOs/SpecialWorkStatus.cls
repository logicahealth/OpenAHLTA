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
Attribute VB_Name = "SpecialWorkStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Implements iObjectPersist
Implements iState

Private Const OBJECT_VERSION = "1.0"
Private Const UNIQUE_IDENTIFIER = "SpecialWorkStatus"


Private mJumpingStatus As String
Private mDivingStatus As String
Private mFlyingStatus As String
Private mMilitaryPolice As String
Private mSubmarine As String
Private mOnMobility As String
Private mState As enumObjectState

Public Property Get JumpingStatus() As String
    JumpingStatus = mJumpingStatus
End Property
Public Property Let JumpingStatus(ByVal value As String)
    mJumpingStatus = value
End Property
Public Property Get DivingStatus() As String
    DivingStatus = mDivingStatus
End Property
Public Property Let DivingStatus(ByVal value As String)
    mDivingStatus = value
End Property
Public Property Get FlyingStatus() As String
    FlyingStatus = mFlyingStatus
End Property
Public Property Let FlyingStatus(ByVal value As String)
    mFlyingStatus = value
End Property
Public Property Get MilitaryPolice() As String
    MilitaryPolice = mMilitaryPolice
End Property
Public Property Let MilitaryPolice(ByVal value As String)
    mMilitaryPolice = value
End Property
Public Property Get Submarine() As String
    Submarine = mSubmarine
End Property
Public Property Let Submarine(ByVal value As String)
    mSubmarine = value
End Property
Public Property Get OnMobility() As String
    OnMobility = mOnMobility
End Property
Public Property Let OnMobility(ByVal value As String)
    mOnMobility = value
End Property




Private Function iObjectPersist_Copy() As Object
    Dim NewCopy As SpecialWorkStatus
    Dim State As iState
    
    On Error GoTo ErrHandler
    
    Set NewCopy = New SpecialWorkStatus
    
    With NewCopy
        .JumpingStatus = mJumpingStatus
        .DivingStatus = mDivingStatus
        .FlyingStatus = mFlyingStatus
        .MilitaryPolice = mMilitaryPolice
        .Submarine = mSubmarine
        .OnMobility = mOnMobility
        
        'iState interface
        Set State = NewCopy
        State.State = mState
    End With

    Set iObjectPersist_Copy = NewCopy

    Set NewCopy = Nothing
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "DTOsSpecialWorkStatus.iObjectPersist_Persist", Err.Description, Err.HelpFile, Err.HelpContext

End Function




Private Sub iObjectPersist_Persist(ByRef Writer As iWriter)
    On Error GoTo ErrHandler
    
    With Writer
        .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION
        .WriteProperty "JumpingStatus", mJumpingStatus
        .WriteProperty "DivingStatus", mDivingStatus
        .WriteProperty "FlyingStatus", mFlyingStatus
        .WriteProperty "MilitaryPolice", mMilitaryPolice
        .WriteProperty "Submarine", mSubmarine
        .WriteProperty "OnMobility", mOnMobility
        
        'iState interface
        .WriteProperty "iState.State", mState
    End With
    
    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, "DTOsDeersEligibility.iObjectPersist_Persist", Err.Description, Err.HelpFile, Err.HelpContext
End Sub




Private Function iObjectPersist_Restore(ByRef Reader As iReader) As Boolean
Dim Failed As Boolean

    On Error GoTo ErrHandler
  
    With Reader
        mJumpingStatus = .ReadProperty("JumpingStatus")
        mDivingStatus = .ReadProperty("DivingStatus")
        mFlyingStatus = .ReadProperty("FlyingStatus")
        mMilitaryPolice = .ReadProperty("MilitaryPolice")
        mSubmarine = .ReadProperty("Submarine")
        mOnMobility = .ReadProperty("OnMobility")
        
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
