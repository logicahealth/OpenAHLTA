VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Disposition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--- begin DTO Builder; DTO=Disposition
'--- 3/16/04 8:20:34 PM
'---
'--- NOTE: Any time you modify this class, increment the version number
'
'

Option Explicit

Implements iObjectPersist
Implements iSection
Implements iState

Private Const OBJECT_VERSION = "1.0"
Private Const UNIQUE_IDENTIFIER = "Disposition"

Private mState As enumObjectState
Private mDentalDispStatus As String '
Private mDentalDispStatusDesc As String '
Private mDentalDispComments As String   '
Private mDentalDispOptions As String    '
Private mStatus As EncStatus    '
Private mText As String '
Private mNCID As String '
Private mAdminOptions As String '
Private mEandMNCID As String    '
Private mEandMCalc As String    '
Private mEMCalcOverrideNote As String   '
Private mWSDisp As String   '
Private mFollowupTimeFrame As String    '
Private mFollowupComments As String '
Private mDiscussedComments As String    '
Private mItemsDiscussed As String   '
Private mDBNICategory As String '
Private mDBNICause As String    '
Private mDispMetaData As String '

'Section
Private mDataID As String
Private mID As ENCOUNTER_INTERFACES.EncounterSectionEnum
Private mOwnerNCID As String
Private mDocument As Document
Private mUpdated As Boolean


Public Property Get Updated() As Boolean
    Updated = mUpdated
End Property
Public Property Let Updated(value As Boolean)
    mUpdated = value
End Property



Public Property Get DentalDispStatus() As String
    DentalDispStatus = mDentalDispStatus
End Property
Public Property Let DentalDispStatus(value As String)
    mDentalDispStatus = value
End Property
Public Property Get DentalDispStatusDesc() As String
    DentalDispStatusDesc = mDentalDispStatusDesc
End Property
Public Property Let DentalDispStatusDesc(value As String)
    mDentalDispStatusDesc = value
End Property
Public Property Get DentalDispComments() As String
    DentalDispComments = mDentalDispComments
End Property
Public Property Let DentalDispComments(value As String)
    mDentalDispComments = value
End Property
Public Property Get DentalDispOptions() As String
    DentalDispOptions = mDentalDispOptions
End Property
Public Property Let DentalDispOptions(value As String)
    mDentalDispOptions = value
End Property
Public Property Get Status() As EncStatus
    Status = mStatus
End Property
Public Property Let Status(value As EncStatus)
    mStatus = value
End Property
Public Property Get Text() As String
    Text = mText
End Property
Public Property Let Text(value As String)
    mText = value
End Property
Public Property Get NCID() As String
    NCID = mNCID
End Property
Public Property Let NCID(value As String)
    mNCID = value
End Property
Public Property Get AdminOptions() As String
    AdminOptions = mAdminOptions
End Property
Public Property Let AdminOptions(value As String)
    mAdminOptions = value
End Property
Public Property Get EandMNCID() As String
    EandMNCID = mEandMNCID
End Property
Public Property Let EandMNCID(value As String)
    mEandMNCID = value
End Property
Public Property Get EandMCalc() As String
    EandMCalc = mEandMCalc
End Property
Public Property Let EandMCalc(value As String)
    mEandMCalc = value
End Property
Public Property Get EMCalcOverrideNote() As String
    EMCalcOverrideNote = mEMCalcOverrideNote
End Property
Public Property Let EMCalcOverrideNote(value As String)
    mEMCalcOverrideNote = value
End Property
Public Property Get WSDisp() As String
    WSDisp = mWSDisp
End Property
Public Property Let WSDisp(value As String)
    mWSDisp = value
End Property
Public Property Get FollowupTimeFrame() As String
    FollowupTimeFrame = mFollowupTimeFrame
End Property
Public Property Let FollowupTimeFrame(value As String)
    mFollowupTimeFrame = value
End Property
Public Property Get FollowupComments() As String
    FollowupComments = mFollowupComments
End Property
Public Property Let FollowupComments(value As String)
    mFollowupComments = value
End Property
Public Property Get DiscussedComments() As String
    DiscussedComments = mDiscussedComments
End Property
Public Property Let DiscussedComments(value As String)
    mDiscussedComments = value
End Property
Public Property Get ItemsDiscussed() As String
    ItemsDiscussed = mItemsDiscussed
End Property
Public Property Let ItemsDiscussed(value As String)
    mItemsDiscussed = value
End Property
Public Property Get DBNICategory() As String
    DBNICategory = mDBNICategory
End Property
Public Property Let DBNICategory(value As String)
    mDBNICategory = value
End Property
Public Property Get DBNICause() As String
    DBNICause = mDBNICause
End Property
Public Property Let DBNICause(value As String)
    mDBNICause = value
End Property
Public Property Get DispMetaData() As String
    DispMetaData = mDispMetaData
End Property
Public Property Let DispMetaData(value As String)
    mDispMetaData = value
End Property




Private Function iObjectPersist_Copy() As Object
    Dim NewCopy As Disposition
    Dim State As iState
    
    Set NewCopy = New Disposition
    
    With NewCopy
        .DentalDispStatus = mDentalDispStatus
        .DentalDispStatusDesc = mDentalDispStatusDesc
        .DentalDispComments = mDentalDispComments
        .DentalDispOptions = mDentalDispOptions
        .Status = mStatus
        .Text = mText
        .NCID = mNCID
        .AdminOptions = mAdminOptions
        .EandMNCID = mEandMNCID
        .EandMCalc = mEandMCalc
        .EMCalcOverrideNote = mEMCalcOverrideNote
        .WSDisp = mWSDisp
        .FollowupTimeFrame = mFollowupTimeFrame
        .FollowupComments = mFollowupComments
        .DiscussedComments = mDiscussedComments
        .ItemsDiscussed = mItemsDiscussed
        .DBNICategory = mDBNICategory
        .DBNICause = mDBNICause
        .DispMetaData = mDispMetaData
        .Updated = Updated
        
        'iState interface
        Set State = NewCopy
        State.State = mState
    End With

    Set iObjectPersist_Copy = NewCopy

End Function




Private Sub iObjectPersist_Persist(Writer As iWriter)
    With Writer
        .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION
        .WriteProperty "DentalDispStatus", mDentalDispStatus
        .WriteProperty "DentalDispStatusDesc", mDentalDispStatusDesc
        .WriteProperty "DentalDispComments", mDentalDispComments
        .WriteProperty "DentalDispOptions", mDentalDispOptions
        .WriteProperty "Status", mStatus
        .WriteProperty "Text", mText
        .WriteProperty "NCID", mNCID
        .WriteProperty "AdminOptions", mAdminOptions
        .WriteProperty "EandMNCID", mEandMNCID
        .WriteProperty "EandMCalc", mEandMCalc
        .WriteProperty "EMCalcOverrideNote", mEMCalcOverrideNote
        .WriteProperty "WSDisp", mWSDisp
        .WriteProperty "FollowupTimeFrame", mFollowupTimeFrame
        .WriteProperty "FollowupComments", mFollowupComments
        .WriteProperty "DiscussedComments", mDiscussedComments
        .WriteProperty "ItemsDiscussed", mItemsDiscussed
        .WriteProperty "DBNICategory", mDBNICategory
        .WriteProperty "DBNICause", mDBNICause
        .WriteProperty "DispMetaData", mDispMetaData
        .WriteProperty "Updated", mUpdated
        
        'iState interface
        .WriteProperty "iState.State", mState
    End With
End Sub




Private Function iObjectPersist_Restore(Reader As iReader) As Boolean
Dim Failed As Boolean

    On Error GoTo ErrHandler
  
    With Reader
        mDentalDispStatus = .ReadProperty("DentalDispStatus")
        mDentalDispStatusDesc = .ReadProperty("DentalDispStatusDesc")
        mDentalDispComments = .ReadProperty("DentalDispComments")
        mDentalDispOptions = .ReadProperty("DentalDispOptions")
        mStatus = .ReadProperty("Status")
        mText = .ReadProperty("Text")
        mNCID = .ReadProperty("NCID")
        mAdminOptions = .ReadProperty("AdminOptions")
        mEandMNCID = .ReadProperty("EandMNCID")
        mEandMCalc = .ReadProperty("EandMCalc")
        mEMCalcOverrideNote = .ReadProperty("EMCalcOverrideNote")
        mWSDisp = .ReadProperty("WSDisp")
        mFollowupTimeFrame = .ReadProperty("FollowupTimeFrame")
        mFollowupComments = .ReadProperty("FollowupComments")
        mDiscussedComments = .ReadProperty("DiscussedComments")
        mItemsDiscussed = .ReadProperty("ItemsDiscussed")
        mDBNICategory = .ReadProperty("DBNICategory")
        mDBNICause = .ReadProperty("DBNICause")
        mDispMetaData = .ReadProperty("DispMetaData")
        mUpdated = .ReadProperty("Updated")
        
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

Private Property Let iSection_DataID(ByVal RHS As Long)
    mDataID = RHS
End Property

Private Property Get iSection_DataID() As Long
    iSection_DataID = mDataID
End Property

Private Property Set iSection_Document(ByVal RHS As Document)
    Set mDocument = RHS
End Property

Private Property Get iSection_Document() As Document
    If mDocument Is Nothing Then Set mDocument = New Document
    Set iSection_Document = mDocument
End Property

Private Property Let iSection_ID(ByVal RHS As ENCOUNTER_INTERFACES.EncounterSectionEnum)
    mID = RHS
End Property

Private Property Get iSection_ID() As ENCOUNTER_INTERFACES.EncounterSectionEnum
    iSection_ID = mID
End Property

Private Property Let iSection_OwnerNCID(ByVal RHS As String)
    mOwnerNCID = RHS
End Property

Private Property Get iSection_OwnerNCID() As String
    iSection_OwnerNCID = mOwnerNCID
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
