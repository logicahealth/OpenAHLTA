VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Pregnancy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True


Option Explicit

Implements iObjectPersist
Implements iState

Private Const OBJECT_VERSION = "1.0"
Private Const UNIQUE_IDENTIFIER = "Pregnancy"


Private mLastMenstrualPeriod As Date
Private mPregnancyIndicator As Boolean
Private mEstimatedDOB As String


'--- Details
Private mMenopauseIndicator As Boolean
Private mHysterectomyIndicator As Boolean
Private mTotalAbortus As Long
Private mSpontaneousAbortions As Long
Private mElectiveAbortions As Long
Private mEctopicAbortions As Long
Private mGravida As Long
Private mPara As Long
Private mLivingChildren As Long
Private mBirthControlMethods As Collection

Private mState As enumObjectState


Public Property Get LastMenstrualPeriod() As Date
    LastMenstrualPeriod = mLastMenstrualPeriod
End Property
Public Property Let LastMenstrualPeriod(ByVal value As Date)
    mLastMenstrualPeriod = value
End Property
Public Property Get PregnancyIndicator() As Boolean
    PregnancyIndicator = mPregnancyIndicator
End Property
Public Property Let PregnancyIndicator(ByVal value As Boolean)
    mPregnancyIndicator = value
End Property
Public Property Get EstimatedDOB() As Date
    If LenB(mEstimatedDOB) > 0 Then
        EstimatedDOB = CDate(mEstimatedDOB)
    End If
End Property
Public Property Let EstimatedDOB(ByVal value As Date)
    mEstimatedDOB = value
End Property

Public Property Let MenopauseIndicator(ByVal bNewValue As Boolean)
    mMenopauseIndicator = bNewValue
End Property

Public Property Get MenopauseIndicator() As Boolean
   MenopauseIndicator = mMenopauseIndicator
End Property


Public Property Let HysterectomyIndicator(ByVal bNewValue As Boolean)
    mHysterectomyIndicator = bNewValue
End Property

Public Property Get HysterectomyIndicator() As Boolean
   HysterectomyIndicator = mHysterectomyIndicator
End Property


Public Property Let TotalAbortus(ByVal lNewValue As Long)
    mTotalAbortus = lNewValue
End Property

Public Property Get TotalAbortus() As Long
   TotalAbortus = mTotalAbortus
End Property

Public Property Let SpontaneousAbortions(ByVal lNewValue As Long)
    mSpontaneousAbortions = lNewValue
End Property

Public Property Get SpontaneousAbortions() As Long
   SpontaneousAbortions = mSpontaneousAbortions
End Property


Public Property Let ElectiveAbortions(ByVal lNewValue As Long)
    mElectiveAbortions = lNewValue
End Property

Public Property Get ElectiveAbortions() As Long
   ElectiveAbortions = mElectiveAbortions
End Property


Public Property Let EctopicAbortions(ByVal lNewValue As Long)
      mEctopicAbortions = lNewValue
End Property

Public Property Get EctopicAbortions() As Long
   EctopicAbortions = mEctopicAbortions
End Property

Public Property Let Gravida(ByVal lNewValue As Long)
    mGravida = lNewValue
End Property

Public Property Get Gravida() As Long
   Gravida = mGravida
End Property

Public Property Let Para(ByVal lNewValue As Long)
    mPara = lNewValue
End Property

Public Property Get Para() As Long
   Para = mPara
End Property


Public Property Let LivingChildren(ByVal lNewValue As Long)
    mLivingChildren = lNewValue
End Property

Public Property Get LivingChildren() As Long
   LivingChildren = mLivingChildren
End Property

Public Property Get BirthControlMethods() As Collection
   If mBirthControlMethods Is Nothing Then
      Set mBirthControlMethods = New Collection
   End If
   Set BirthControlMethods = mBirthControlMethods
End Property

Public Property Set BirthControlMethods(ByRef oNewValue As Collection)
   Set mBirthControlMethods = oNewValue
End Property



Private Function iObjectPersist_Copy() As Object
    Dim NewCopy As Pregnancy
    Dim State As iState
    Dim i As Long
    
    On Error GoTo ErrHandler
    
    Set NewCopy = New Pregnancy
    
    With NewCopy
        .LastMenstrualPeriod = mLastMenstrualPeriod
        .PregnancyIndicator = mPregnancyIndicator
        .EstimatedDOB = mEstimatedDOB
        .MenopauseIndicator = mMenopauseIndicator
        .HysterectomyIndicator = mHysterectomyIndicator
        .TotalAbortus = mTotalAbortus
        .SpontaneousAbortions = mSpontaneousAbortions
        .ElectiveAbortions = mElectiveAbortions
        .EctopicAbortions = mEctopicAbortions
        .Gravida = mGravida
        .Para = mPara
        .LivingChildren = mLivingChildren
        
        For i = 1 To mBirthControlMethods.Count
          .BirthControlMethods.Add mBirthControlMethods.Item(i)
        Next
        
        'iState interface
        Set State = NewCopy
        State.State = mState
    End With

    Set iObjectPersist_Copy = NewCopy

    Set NewCopy = Nothing
    Set State = Nothing
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "DTOsPregnancy.iObjectPersist_Persist", Err.Description, Err.HelpFile, Err.HelpContext

End Function




Private Sub iObjectPersist_Persist(ByRef Writer As iWriter)
    
    On Error GoTo ErrHandler
    
    With Writer
        .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION
        .WriteProperty "LastMenstrualPeriod", mLastMenstrualPeriod
        .WriteProperty "PregnancyIndicator", mPregnancyIndicator
        .WriteProperty "EstimatedDOB", mEstimatedDOB
        .WriteProperty "MenopauseIndicator", mMenopauseIndicator
        .WriteProperty "HysterectomyIndicator", mHysterectomyIndicator
        .WriteProperty "TotalAbortus", mTotalAbortus
        .WriteProperty "SpontaneousAbortions", mSpontaneousAbortions
        .WriteProperty "ElectiveAbortions", mElectiveAbortions
        .WriteProperty "EctopicAbortions", mEctopicAbortions
        .WriteProperty "Gravida", mGravida
        .WriteProperty "Para", mPara
        .WriteProperty "LivingChildren", mLivingChildren
        .WriteObject "BirthControlMethods", mBirthControlMethods
        
        'iState interface
        .WriteProperty "iState.State", mState
    End With
    
    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, "DTOsPregnancy.iObjectPersist_Persist", Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Private Function iObjectPersist_Restore(ByRef Reader As iReader) As Boolean
Dim Failed As Boolean

    On Error GoTo ErrHandler
  
    With Reader
        mLastMenstrualPeriod = .ReadProperty("LastMenstrualPeriod")
        mPregnancyIndicator = .ReadProperty("PregnancyIndicator")
        mEstimatedDOB = .ReadProperty("EstimatedDOB")
        mMenopauseIndicator = .ReadProperty("MenopauseIndicator")
        mHysterectomyIndicator = .ReadProperty("HysterectomyIndicator")
        mTotalAbortus = .ReadProperty("TotalAbortus")
        mSpontaneousAbortions = .ReadProperty("SpontaneousAbortions")
        mElectiveAbortions = .ReadProperty("ElectiveAbortions")
        mEctopicAbortions = .ReadProperty("EctopicAbortions")
        mGravida = .ReadProperty("Gravida")
        mPara = .ReadProperty("Para")
        mLivingChildren = .ReadProperty("LivingChildren")
        Set mBirthControlMethods = .ReadCollection("BirthControlMethods")
        
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
