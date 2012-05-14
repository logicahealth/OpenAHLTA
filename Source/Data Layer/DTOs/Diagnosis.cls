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
Attribute VB_Name = "Diagnosis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--- begin DTO Builder; DTO=Diagnosis
'--- 3/16/04 8:33:22 PM
'---
'--- NOTE: Any time you modify this class, increment the version number
'
'

Option Explicit

Implements iObjectPersist
Implements iState

Private Const OBJECT_VERSION = "1.0"
Private Const UNIQUE_IDENTIFIER = "Diagnosis"

Private mState As enumObjectState
Private mChronicity As String   '
Private mClinicNCID As String   '
Private mComment As String  '
Private mDataID As String   '
Private mEncID As String    '
Private mExtenderID As String   '
Private mExtenderCD As String   '
Private mExtenderNM As String   '
Private mFacilityNCID As String '
Private mICD9Code As String '
Private mKey As String  '
Private mLongNM As String   '
Private mSnoID As String '
Private mName As String '
Private mNCID As String '
Private mOnsetDate As String    '
Private mPrefix As String   '
Private mPriority As String '
Private mShortNM As String  '
Private mDxType As String   '
Private mFromProbList As Boolean    '



Public Property Get Chronicity() As String
    Chronicity = mChronicity
End Property
Public Property Let Chronicity(value As String)
    mChronicity = value
End Property
Public Property Get ClinicNCID() As String
    ClinicNCID = mClinicNCID
End Property
Public Property Let ClinicNCID(value As String)
    mClinicNCID = value
End Property
Public Property Get Comment() As String
    Comment = mComment
End Property
Public Property Let Comment(value As String)
    mComment = value
End Property
Public Property Get DataID() As String
    DataID = mDataID
End Property
Public Property Let DataID(value As String)
    mDataID = value
End Property
Public Property Get EncID() As String
    EncID = mEncID
End Property
Public Property Let EncID(value As String)
    mEncID = value
End Property
Public Property Get ExtenderID() As String
    ExtenderID = mExtenderID
End Property
Public Property Let ExtenderID(value As String)
    mExtenderID = value
End Property
Public Property Get ExtenderCD() As String
    ExtenderCD = mExtenderCD
End Property
Public Property Let ExtenderCD(value As String)
    mExtenderCD = value
End Property
Public Property Get ExtenderNM() As String
    ExtenderNM = mExtenderNM
End Property
Public Property Let ExtenderNM(value As String)
    mExtenderNM = value
End Property
Public Property Get FacilityNCID() As String
    FacilityNCID = mFacilityNCID
End Property
Public Property Let FacilityNCID(value As String)
    mFacilityNCID = value
End Property
Public Property Get ICD9Code() As String
    ICD9Code = mICD9Code
End Property
Public Property Let ICD9Code(value As String)
    mICD9Code = value
End Property
Public Property Get Key() As String
    Key = mKey
End Property
Public Property Let Key(value As String)
    mKey = value
End Property
Public Property Get LongNM() As String
    LongNM = mLongNM
End Property
Public Property Let LongNM(value As String)
    mLongNM = value
End Property
Public Property Get SnoID() As String
    SnoID = mSnoID
End Property
Public Property Let SnoID(value As String)
    mSnoID = value
End Property
Public Property Get Name() As String
    Name = mName
End Property
Public Property Let Name(value As String)
    mName = value
End Property
Public Property Get NCID() As String
    NCID = mNCID
End Property
Public Property Let NCID(value As String)
    mNCID = value
End Property
Public Property Get OnsetDate() As String
    OnsetDate = mOnsetDate
End Property
Public Property Let OnsetDate(value As String)
    mOnsetDate = value
End Property
Public Property Get Prefix() As String
    Prefix = mPrefix
End Property
Public Property Let Prefix(value As String)
    mPrefix = value
End Property
Public Property Get Priority() As String
    Priority = mPriority
End Property
Public Property Let Priority(value As String)
    mPriority = value
End Property
Public Property Get ShortNM() As String
    ShortNM = mShortNM
End Property
Public Property Let ShortNM(value As String)
    mShortNM = value
End Property
Public Property Get DxType() As String
    DxType = mDxType
End Property
Public Property Let DxType(value As String)
    mDxType = value
End Property
Public Property Get FromProbList() As Boolean
    FromProbList = mFromProbList
End Property
Public Property Let FromProbList(value As Boolean)
    mFromProbList = value
End Property




Private Function iObjectPersist_Copy() As Object
    Dim NewCopy As Diagnosis
    Dim State As iState
    
    Set NewCopy = New Diagnosis
    
    With NewCopy
        .Chronicity = mChronicity
        .ClinicNCID = mClinicNCID
        .Comment = mComment
        .DataID = mDataID
        .EncID = mEncID
        .ExtenderID = mExtenderID
        .ExtenderCD = mExtenderCD
        .ExtenderNM = mExtenderNM
        .FacilityNCID = mFacilityNCID
        .ICD9Code = mICD9Code
        .Key = mKey
        .LongNM = mLongNM
        .SnoID = mSnoID
        .Name = mName
        .NCID = mNCID
        .OnsetDate = mOnsetDate
        .Prefix = mPrefix
        .Priority = mPriority
        .ShortNM = mShortNM
        .DxType = mDxType
        .FromProbList = mFromProbList
        
        'iState interface
        Set State = NewCopy
        State.State = mState
    End With

    Set iObjectPersist_Copy = NewCopy

End Function




Private Sub iObjectPersist_Persist(Writer As iWriter)
    With Writer
        .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION
        .WriteProperty "Chronicity", mChronicity
        .WriteProperty "ClinicNCID", mClinicNCID
        .WriteProperty "Comment", mComment
        .WriteProperty "DataID", mDataID
        .WriteProperty "EncID", mEncID
        .WriteProperty "ExtenderID", mExtenderID
        .WriteProperty "ExtenderCD", mExtenderCD
        .WriteProperty "ExtenderNM", mExtenderNM
        .WriteProperty "FacilityNCID", mFacilityNCID
        .WriteProperty "ICD9Code", mICD9Code
        .WriteProperty "Key", mKey
        .WriteProperty "LongNM", mLongNM
        .WriteProperty "SnoID", mSnoID
        .WriteProperty "Name", mName
        .WriteProperty "NCID", mNCID
        .WriteProperty "OnsetDate", mOnsetDate
        .WriteProperty "Prefix", mPrefix
        .WriteProperty "Priority", mPriority
        .WriteProperty "ShortNM", mShortNM
        .WriteProperty "DxType", mDxType
        .WriteProperty "FromProbList", mFromProbList
        
        'iState interface
        .WriteProperty "iState.State", mState
    End With
End Sub




Private Function iObjectPersist_Restore(Reader As iReader) As Boolean
Dim Failed As Boolean

    On Error GoTo ErrHandler
  
    With Reader
        mChronicity = .ReadProperty("Chronicity")
        mClinicNCID = .ReadProperty("ClinicNCID")
        mComment = .ReadProperty("Comment")
        mDataID = .ReadProperty("DataID")
        mEncID = .ReadProperty("EncID")
        mExtenderID = .ReadProperty("ExtenderID")
        mExtenderCD = .ReadProperty("ExtenderCD")
        mExtenderNM = .ReadProperty("ExtenderNM")
        mFacilityNCID = .ReadProperty("FacilityNCID")
        mICD9Code = .ReadProperty("ICD9Code")
        mKey = .ReadProperty("Key")
        mLongNM = .ReadProperty("LongNM")
        mSnoID = .ReadProperty("SnoID")
        mName = .ReadProperty("Name")
        mNCID = .ReadProperty("NCID")
        mOnsetDate = .ReadProperty("OnsetDate")
        mPrefix = .ReadProperty("Prefix")
        mPriority = .ReadProperty("Priority")
        mShortNM = .ReadProperty("ShortNM")
        mDxType = .ReadProperty("DxType")
        mFromProbList = .ReadProperty("FromProbList")
        
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
