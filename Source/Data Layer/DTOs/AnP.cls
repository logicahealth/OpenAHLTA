VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AnP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--- begin DTO Builder; DTO=AnP
'--- 3/16/04 8:11:29 PM
'---
'--- NOTE: Any time you modify this class, increment the version number
'
'

Option Explicit

Implements iObjectPersist
Implements iSection
Implements iState

Private Const OBJECT_VERSION = "1.0"
Private Const UNIQUE_IDENTIFIER = "AnP"

Private mState As enumObjectState
Private mStatus As EncStatus    '
Private mDiagnoses As Diagnoses '
Private mProcedures As Procedures   '

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

Public Property Get Status() As EncStatus
    Status = mStatus
End Property
Public Property Let Status(value As EncStatus)
    mStatus = value
End Property
Public Property Get Diagnoses() As Diagnoses
    If mDiagnoses Is Nothing Then Set mDiagnoses = New Diagnoses
    Set Diagnoses = mDiagnoses
End Property
Public Property Set Diagnoses(value As Diagnoses)
    Set mDiagnoses = value
End Property
Public Property Get Procedures() As Procedures
    If mProcedures Is Nothing Then Set mProcedures = New Procedures
    Set Procedures = mProcedures
End Property
Public Property Set Procedures(value As Procedures)
    Set mProcedures = value
End Property


Private Function iObjectPersist_Copy() As Object
    Dim NewCopy As AnP
    Dim State As iState
    
    Set NewCopy = New AnP
    
    With NewCopy
        .Status = mStatus
        Set .Diagnoses = mDiagnoses
        Set .Procedures = mProcedures
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
        .WriteProperty "Status", mStatus
        .WriteObject "Diagnoses", mDiagnoses
        .WriteObject "Procedures", mProcedures
        .WriteProperty "Updated", mUpdated
        
        'iState interface
        .WriteProperty "iState.State", mState
    End With
End Sub




Private Function iObjectPersist_Restore(Reader As iReader) As Boolean
Dim Failed As Boolean

    On Error GoTo ErrHandler
  
    With Reader
        mStatus = .ReadProperty("Status")
        .ReadObject "Diagnoses", Diagnoses
        .ReadObject "Procedures", Procedures
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
