VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SONote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--- begin DTO Builder; DTO=SONote
'--- 3/16/04 8:13:20 PM
'---
'--- NOTE: Any time you modify this class, increment the version number
'
'

Option Explicit

Implements iObjectPersist
Implements iSection
Implements iState

Private Const OBJECT_VERSION = "1.0"
Private Const UNIQUE_IDENTIFIER = "SONote"

Private mState As enumObjectState

'Section
Private mDataID As String
Private mID As ENCOUNTER_INTERFACES.EncounterSectionEnum
Private mOwnerNCID As String
Private mDocument As Document

Private mReadOnly As Boolean    '
Private mNewCouplerData As Boolean  '
Private mCouplerData As Boolean '
Private mTextNote As Note   '
Private mMedcinRecords As MedcinRecords '



Public Property Get ReadOnly() As Boolean
    ReadOnly = mReadOnly
End Property
Public Property Let ReadOnly(value As Boolean)
    mReadOnly = value
End Property
Public Property Get NewCouplerData() As Boolean
    NewCouplerData = mNewCouplerData
End Property
Public Property Let NewCouplerData(value As Boolean)
    mNewCouplerData = value
End Property
Public Property Get CouplerData() As Boolean
    CouplerData = mCouplerData
End Property
Public Property Let CouplerData(value As Boolean)
    mCouplerData = value
End Property
Public Property Get TextNote() As Note
    If mTextNote Is Nothing Then Set mTextNote = New Note
    Set TextNote = mTextNote
End Property
Public Property Set TextNote(value As Note)
    Set mTextNote = value
End Property
Public Property Get MedcinRecords() As MedcinRecords
    If mMedcinRecords Is Nothing Then Set mMedcinRecords = New MedcinRecords
    Set MedcinRecords = mMedcinRecords
End Property
Public Property Set MedcinRecords(value As MedcinRecords)
    Set mMedcinRecords = value
End Property




Private Function iObjectPersist_Copy() As Object
    Dim NewCopy As SONote
    Dim State As iState
    
    Set NewCopy = New SONote
    
    With NewCopy
        .ReadOnly = mReadOnly
        .NewCouplerData = mNewCouplerData
        .CouplerData = mCouplerData
        .TextNote = mTextNote
        Set .MedcinRecords = mMedcinRecords
        
        'iState interface
        Set State = NewCopy
        State.State = mState
    End With

    Set iObjectPersist_Copy = NewCopy

End Function




Private Sub iObjectPersist_Persist(Writer As iWriter)
    With Writer
        .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION
        .WriteProperty "ReadOnly", mReadOnly
        .WriteProperty "NewCouplerData", mNewCouplerData
        .WriteProperty "CouplerData", mCouplerData
        .WriteProperty "TextNote", mTextNote
        .WriteObject "MedcinRecords", mMedcinRecords
        
        'iState interface
        .WriteProperty "iState.State", mState
    End With
End Sub




Private Function iObjectPersist_Restore(Reader As iReader) As Boolean
Dim Failed As Boolean

    On Error GoTo ErrHandler
  
    With Reader
        mReadOnly = .ReadProperty("ReadOnly")
        mNewCouplerData = .ReadProperty("NewCouplerData")
        mCouplerData = .ReadProperty("CouplerData")
        mTextNote = .ReadProperty("TextNote")
        .ReadObject "MedcinRecords", MedcinRecords
        
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
