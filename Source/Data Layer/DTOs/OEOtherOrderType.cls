VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OEOtherOrderType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

'--- begin XML DTO Builder; DTO=OEOtherOrderType
'--- [Insert DateTime Stamp here]
'---
'--- NOTE: Any time you modify this class, increment the version number
'

Implements iObjectPersist
Implements iState

Private Const OBJECT_VERSION = "1.0"
Private Const UNIQUE_IDENTIFIER = "OEOtherOrderType"

Private mID As String 'UniqueID
Private mCategory As String 'The type of other order
Private mDescription As String '
Private mComment As String '
Private mStatus As Long 'Active, Inactive, Deleted
Private mFacilityNCID As String '
Private mState As enumObjectState

Public Property Get ID() As String
    ID = mID
End Property
Public Property Let ID(ByVal value As String)
    mID = value
    If value <> mID And mState = osSaved Then
        mState = osModified
    End If
End Property
                        
Public Property Get Category() As String
    Category = mCategory
End Property
Public Property Let Category(ByVal value As String)
    mCategory = value
    If value <> mCategory And mState = osSaved Then
        mState = osModified
    End If
End Property
                        
Public Property Get Description() As String
    Description = mDescription
End Property
Public Property Let Description(ByVal value As String)
    mDescription = value
    If value <> mDescription And mState = osSaved Then
        mState = osModified
    End If
End Property
                        
Public Property Get Comment() As String
    Comment = mComment
End Property
Public Property Let Comment(ByVal value As String)
    mComment = value
    If value <> mComment And mState = osSaved Then
        mComment = osModified
    End If
End Property
                        
Public Property Get Status() As Long
    Status = mStatus
End Property
Public Property Let Status(ByVal value As Long)
    mStatus = value
    If value <> mStatus And mState = osSaved Then
        mState = osModified
    End If
End Property
                        
Public Property Get FacilityNCID() As String
    FacilityNCID = mFacilityNCID
End Property
Public Property Let FacilityNCID(ByVal value As String)
    mFacilityNCID = value
    If value <> mFacilityNCID And mState = osSaved Then
        mState = osModified
    End If

End Property
                        
Private Sub Class_Initialize()
    mState = osNew
End Sub

Private Function iObjectPersist_Copy() As Object
    Dim NewCopy As OEOtherOrderType
    Dim PersistHelper As ObjectPersist.PersistHelper
    Dim State As iState
    Dim i As Long 'Iterator if we need to copy a collection
    
    Set NewCopy = New OEOtherOrderType
    Set PersistHelper = New ObjectPersist.PersistHelper

    With NewCopy
                .ID = mID
                .Category = mCategory
                .Description = mDescription
                .Comment = mComment
                .Status = mStatus
                .FacilityNCID = mFacilityNCID

        'iState interface
        Set State = NewCopy
        State.State = mState
    End With

    Set iObjectPersist_Copy = NewCopy
        Set NewCopy = Nothing
        Set State = Nothing
End Function

Private Sub iObjectPersist_Persist(Writer As iWriter)
    With Writer
        .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION
        
        .WriteProperty "ID", mID
        .WriteProperty "Category", mCategory
        .WriteProperty "Description", mDescription
        .WriteProperty "Comment", mComment
        .WriteProperty "Status", mStatus
        .WriteProperty "FacilityNCID", mFacilityNCID

        'iState interface
        .WriteProperty "iState.State", mState

    End With
End Sub


Private Function iObjectPersist_Restore(Reader As iReader) As Boolean
Dim Failed As Boolean

    On Error GoTo ErrHandler

    With Reader
        
        mID = .ReadProperty("ID")
        mCategory = .ReadProperty("Category")
        mDescription = .ReadProperty("Description")
        mComment = .ReadProperty("Comment")
        mStatus = .ReadProperty("Status")
        mFacilityNCID = .ReadProperty("FacilityNCID")

        'iState interface
        mState = .ReadProperty("iState.State")

    End With

    iObjectPersist_Restore = Not Failed

    Exit Function

ErrHandler:
    Failed = True
    Resume Next

End Function


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



Private Property Get iObjectPersist_UniqueIdentifier() As String
    iObjectPersist_UniqueIdentifier = UNIQUE_IDENTIFIER
End Property

Private Property Get iObjectPersist_Version() As String
    iObjectPersist_Version = OBJECT_VERSION
End Property
'--- end DTO Builder
