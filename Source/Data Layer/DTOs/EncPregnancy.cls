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
Attribute VB_Name = "EncPregnancy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--- begin DTO Builder; DTO=EncPregnancy
'--- 3/16/04 7:55:13 PM
'---
'--- NOTE: Any time you modify this class, increment the version number
'
'

Option Explicit

Implements iObjectPersist
Implements iState

Private Const OBJECT_VERSION = "1.0"
Private Const UNIQUE_IDENTIFIER = "EncPregnancy"

Private mState As enumObjectState
Private mPregnancyIndicator As Boolean  '
Private mLastPeriod As Date '
Private mEstDOB As Date '
Private mMenopauseIndicator As Boolean  '
Private mHysterectomyIndicator As Boolean   '
Private mTotalAbortus As Boolean    '
Private mSpontaneousAbortions As Long   '
Private mElectiveAbortions As Long  '
Private mEctopicAbortions As Long   '
Private mGravida As Long    '
Private mPara As Long   '
Private mLivingChildren As Long '



Public Property Get PregnancyIndicator() As Boolean
    PregnancyIndicator = mPregnancyIndicator
End Property
Public Property Let PregnancyIndicator(value As Boolean)
    mPregnancyIndicator = value
End Property
Public Property Get LastPeriod() As Date
    LastPeriod = mLastPeriod
End Property
Public Property Let LastPeriod(value As Date)
    mLastPeriod = value
End Property
Public Property Get EstDOB() As Date
    EstDOB = mEstDOB
End Property
Public Property Let EstDOB(value As Date)
    mEstDOB = value
End Property
Public Property Get MenopauseIndicator() As Boolean
    MenopauseIndicator = mMenopauseIndicator
End Property
Public Property Let MenopauseIndicator(value As Boolean)
    mMenopauseIndicator = value
End Property
Public Property Get HysterectomyIndicator() As Boolean
    HysterectomyIndicator = mHysterectomyIndicator
End Property
Public Property Let HysterectomyIndicator(value As Boolean)
    mHysterectomyIndicator = value
End Property
Public Property Get TotalAbortus() As Boolean
    TotalAbortus = mTotalAbortus
End Property
Public Property Let TotalAbortus(value As Boolean)
    mTotalAbortus = value
End Property
Public Property Get SpontaneousAbortions() As Long
    SpontaneousAbortions = mSpontaneousAbortions
End Property
Public Property Let SpontaneousAbortions(value As Long)
    mSpontaneousAbortions = value
End Property
Public Property Get ElectiveAbortions() As Long
    ElectiveAbortions = mElectiveAbortions
End Property
Public Property Let ElectiveAbortions(value As Long)
    mElectiveAbortions = value
End Property
Public Property Get EctopicAbortions() As Long
    EctopicAbortions = mEctopicAbortions
End Property
Public Property Let EctopicAbortions(value As Long)
    mEctopicAbortions = value
End Property
Public Property Get Gravida() As Long
    Gravida = mGravida
End Property
Public Property Let Gravida(value As Long)
    mGravida = value
End Property
Public Property Get Para() As Long
    Para = mPara
End Property
Public Property Let Para(value As Long)
    mPara = value
End Property
Public Property Get LivingChildren() As Long
    LivingChildren = mLivingChildren
End Property
Public Property Let LivingChildren(value As Long)
    mLivingChildren = value
End Property




Private Function iObjectPersist_Copy() As Object
    Dim NewCopy As EncPregnancy
    Dim State As iState
    
    Set NewCopy = New EncPregnancy
    
    With NewCopy
        .PregnancyIndicator = mPregnancyIndicator
        .LastPeriod = mLastPeriod
        .EstDOB = mEstDOB
        .MenopauseIndicator = mMenopauseIndicator
        .HysterectomyIndicator = mHysterectomyIndicator
        .TotalAbortus = mTotalAbortus
        .SpontaneousAbortions = mSpontaneousAbortions
        .ElectiveAbortions = mElectiveAbortions
        .EctopicAbortions = mEctopicAbortions
        .Gravida = mGravida
        .Para = mPara
        .LivingChildren = mLivingChildren
        
        'iState interface
        Set State = NewCopy
        State.State = mState
    End With

    Set iObjectPersist_Copy = NewCopy

End Function




Private Sub iObjectPersist_Persist(Writer As iWriter)
    With Writer
        .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION
        .WriteProperty "PregnancyIndicator", mPregnancyIndicator
        .WriteProperty "LastPeriod", mLastPeriod
        .WriteProperty "EstDOB", mEstDOB
        .WriteProperty "MenopauseIndicator", mMenopauseIndicator
        .WriteProperty "HysterectomyIndicator", mHysterectomyIndicator
        .WriteProperty "TotalAbortus", mTotalAbortus
        .WriteProperty "SpontaneousAbortions", mSpontaneousAbortions
        .WriteProperty "ElectiveAbortions", mElectiveAbortions
        .WriteProperty "EctopicAbortions", mEctopicAbortions
        .WriteProperty "Gravida", mGravida
        .WriteProperty "Para", mPara
        .WriteProperty "LivingChildren", mLivingChildren
        
        'iState interface
        .WriteProperty "iState.State", mState
    End With
End Sub




Private Function iObjectPersist_Restore(Reader As iReader) As Boolean
Dim Failed As Boolean

    On Error GoTo ErrHandler
  
    With Reader
        mPregnancyIndicator = .ReadProperty("PregnancyIndicator")
        mLastPeriod = .ReadProperty("LastPeriod")
        mEstDOB = .ReadProperty("EstDOB")
        mMenopauseIndicator = .ReadProperty("MenopauseIndicator")
        mHysterectomyIndicator = .ReadProperty("HysterectomyIndicator")
        mTotalAbortus = .ReadProperty("TotalAbortus")
        mSpontaneousAbortions = .ReadProperty("SpontaneousAbortions")
        mElectiveAbortions = .ReadProperty("ElectiveAbortions")
        mEctopicAbortions = .ReadProperty("EctopicAbortions")
        mGravida = .ReadProperty("Gravida")
        mPara = .ReadProperty("Para")
        mLivingChildren = .ReadProperty("LivingChildren")
        
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
