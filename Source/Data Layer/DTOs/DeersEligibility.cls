VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DeersEligibility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Implements iObjectPersist
Implements iState

Private Const OBJECT_VERSION = "1.0"
Private Const UNIQUE_IDENTIFIER = "DeersEligibility"


Private mDescription As String
Private mStartDate As String
Private mEndDate As String

Private mState As enumObjectState


Public Property Get Description() As String
    Description = mDescription
End Property
Public Property Let Description(ByVal value As String)
    mDescription = value
End Property
Public Property Get StartDate() As String
    StartDate = mStartDate
End Property
Public Property Let StartDate(ByVal value As String)
    mStartDate = value
End Property
Public Property Get EndDate() As String
    EndDate = mEndDate
End Property
Public Property Let EndDate(ByVal value As String)
    mEndDate = value
End Property




Private Function iObjectPersist_Copy() As Object
    Dim NewCopy As DeersEligibility
    Dim State As iState
    
    On Error GoTo ErrHandler
    
    Set NewCopy = New DeersEligibility
    
    With NewCopy
        .Description = mDescription
        .StartDate = mStartDate
        .EndDate = mEndDate
        
        'iState interface
        Set State = NewCopy
        State.State = mState
    End With

    Set iObjectPersist_Copy = NewCopy
    
    Set NewCopy = Nothing
    Set State = Nothing

    Exit Function
ErrHandler:
    Err.Raise Err.Number, "DTOsDeersEligibility.iObjectPersist_Persist", Err.Description, Err.HelpFile, Err.HelpContext

End Function




Private Sub iObjectPersist_Persist(ByRef Writer As iWriter)

    On Error GoTo ErrHandler
    
    With Writer
        .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION
        .WriteProperty "Description", mDescription
        .WriteProperty "StartDate", mStartDate
        .WriteProperty "EndDate", mEndDate
        
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
        mDescription = .ReadProperty("Description")
        mStartDate = .ReadProperty("StartDate")
        mEndDate = .ReadProperty("EndDate")
        
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
