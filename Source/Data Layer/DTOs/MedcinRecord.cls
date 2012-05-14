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
Attribute VB_Name = "MedcinRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--- begin DTO Builder; DTO=MedcinRecord
'--- 3/16/04 8:23:58 PM
'---
'--- NOTE: Any time you modify this class, increment the version number
'
'

Option Explicit

Implements iObjectPersist
Implements iState

Private Const OBJECT_VERSION = "1.0"
Private Const UNIQUE_IDENTIFIER = "MedcinRecord"

Private mState As enumObjectState
Private mSnoID As String '
Private mPrefix As String   '
Private mModifier As String '
Private mResult As String   '
Private mStatus As String   '
Private mOnset As String    '
Private mDuration As String '
Private mValue As String    '
Private mUnit As String '
Private mRangeScale As Integer  '
Private mRangeNormalLow As Double   '
Private mRangeNormalHigh As Double  '
Private mReferid As Long    '
Private mFlag As Integer    '
Private mNote As String '
Private mEpisode As String  '
Private mQualifierLink As String    '
Private mQualifierLinkSequence As String    '
Private mDataID As Long '
Private mPCKText As String  '



Public Property Get SnoID() As String
    SnoID = mSnoID
End Property
Public Property Let SnoID(value As String)
    mSnoID = value
End Property
Public Property Get Prefix() As String
    Prefix = mPrefix
End Property
Public Property Let Prefix(value As String)
    mPrefix = value
End Property
Public Property Get Modifier() As String
    Modifier = mModifier
End Property
Public Property Let Modifier(value As String)
    mModifier = value
End Property
Public Property Get Result() As String
    Result = mResult
End Property
Public Property Let Result(value As String)
    mResult = value
End Property
Public Property Get Status() As String
    Status = mStatus
End Property
Public Property Let Status(value As String)
    mStatus = value
End Property
Public Property Get Onset() As String
    Onset = mOnset
End Property
Public Property Let Onset(value As String)
    mOnset = value
End Property
Public Property Get Duration() As String
    Duration = mDuration
End Property
Public Property Let Duration(value As String)
    mDuration = value
End Property
Public Property Get value() As String
    value = mValue
End Property
Public Property Let value(value As String)
    mValue = value
End Property
Public Property Get Unit() As String
    Unit = mUnit
End Property
Public Property Let Unit(value As String)
    mUnit = value
End Property
Public Property Get RangeScale() As Integer
    RangeScale = mRangeScale
End Property
Public Property Let RangeScale(value As Integer)
    mRangeScale = value
End Property
Public Property Get RangeNormalLow() As Double
    RangeNormalLow = mRangeNormalLow
End Property
Public Property Let RangeNormalLow(value As Double)
    mRangeNormalLow = value
End Property
Public Property Get RangeNormalHigh() As Double
    RangeNormalHigh = mRangeNormalHigh
End Property
Public Property Let RangeNormalHigh(value As Double)
    mRangeNormalHigh = value
End Property
Public Property Get Referid() As Long
    Referid = mReferid
End Property
Public Property Let Referid(value As Long)
    mReferid = value
End Property
Public Property Get Flag() As Integer
    Flag = mFlag
End Property
Public Property Let Flag(value As Integer)
    mFlag = value
End Property
Public Property Get Note() As String
    Note = mNote
End Property
Public Property Let Note(value As String)
    mNote = value
End Property
Public Property Get Episode() As String
    Episode = mEpisode
End Property
Public Property Let Episode(value As String)
    mEpisode = value
End Property
Public Property Get QualifierLink() As String
    QualifierLink = mQualifierLink
End Property
Public Property Let QualifierLink(value As String)
    mQualifierLink = value
End Property
Public Property Get QualifierLinkSequence() As String
    QualifierLinkSequence = mQualifierLinkSequence
End Property
Public Property Let QualifierLinkSequence(value As String)
    mQualifierLinkSequence = value
End Property
Public Property Get DataID() As Long
    DataID = mDataID
End Property
Public Property Let DataID(value As Long)
    mDataID = value
End Property
Public Property Get PCKText() As String
    PCKText = mPCKText
End Property
Public Property Let PCKText(value As String)
    mPCKText = value
End Property




Private Function iObjectPersist_Copy() As Object
    Dim NewCopy As MedcinRecord
    Dim State As iState
    
    Set NewCopy = New MedcinRecord
    
    With NewCopy
        .SnoID = mSnoID
        .Prefix = mPrefix
        .Modifier = mModifier
        .Result = mResult
        .Status = mStatus
        .Onset = mOnset
        .Duration = mDuration
        .value = mValue
        .Unit = mUnit
        .RangeScale = mRangeScale
        .RangeNormalLow = mRangeNormalLow
        .RangeNormalHigh = mRangeNormalHigh
        .Referid = mReferid
        .Flag = mFlag
        .Note = mNote
        .Episode = mEpisode
        .QualifierLink = mQualifierLink
        .QualifierLinkSequence = mQualifierLinkSequence
        .DataID = mDataID
        .PCKText = mPCKText
        
        'iState interface
        Set State = NewCopy
        State.State = mState
    
    End With

    Set iObjectPersist_Copy = NewCopy

End Function




Private Sub iObjectPersist_Persist(Writer As iWriter)
    With Writer
        .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION
        .WriteProperty "SnoID", mSnoID
        .WriteProperty "Prefix", mPrefix
        .WriteProperty "Modifier", mModifier
        .WriteProperty "Result", mResult
        .WriteProperty "Status", mStatus
        .WriteProperty "Onset", mOnset
        .WriteProperty "Duration", mDuration
        .WriteProperty "Value", mValue
        .WriteProperty "Unit", mUnit
        .WriteProperty "RangeScale", mRangeScale
        .WriteProperty "RangeNormalLow", mRangeNormalLow
        .WriteProperty "RangeNormalHigh", mRangeNormalHigh
        .WriteProperty "Referid", mReferid
        .WriteProperty "Flag", mFlag
        .WriteProperty "Note", mNote
        .WriteProperty "Episode", mEpisode
        .WriteProperty "QualifierLink", mQualifierLink
        .WriteProperty "QualifierLinkSequence", mQualifierLinkSequence
        .WriteProperty "DataID", mDataID
        .WriteProperty "PCKText", mPCKText
        
        'iState interface
        .WriteProperty "iState.State", mState
    End With
End Sub




Private Function iObjectPersist_Restore(Reader As iReader) As Boolean
Dim Failed As Boolean

    On Error GoTo ErrHandler
  
    With Reader
        mSnoID = .ReadProperty("SnoID")
        mPrefix = .ReadProperty("Prefix")
        mModifier = .ReadProperty("Modifier")
        mResult = .ReadProperty("Result")
        mStatus = .ReadProperty("Status")
        mOnset = .ReadProperty("Onset")
        mDuration = .ReadProperty("Duration")
        mValue = .ReadProperty("Value")
        mUnit = .ReadProperty("Unit")
        mRangeScale = .ReadProperty("RangeScale")
        mRangeNormalLow = .ReadProperty("RangeNormalLow")
        mRangeNormalHigh = .ReadProperty("RangeNormalHigh")
        mReferid = .ReadProperty("Referid")
        mFlag = .ReadProperty("Flag")
        mNote = .ReadProperty("Note")
        mEpisode = .ReadProperty("Episode")
        mQualifierLink = .ReadProperty("QualifierLink")
        mQualifierLinkSequence = .ReadProperty("QualifierLinkSequence")
        mDataID = .ReadProperty("DataID")
        mPCKText = .ReadProperty("PCKText")
        
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
