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
Attribute VB_Name = "RFV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--- begin DTO Builder; DTO=RFV
'--- 3/16/04 8:05:12 PM
'---
'--- NOTE: Any time you modify this class, increment the version number
'
'

Option Explicit

Implements iObjectPersist
Implements iSection
Implements iState

Private Const OBJECT_VERSION = "1.0"
Private Const UNIQUE_IDENTIFIER = "RFV"

Private mState As enumObjectState
Private mWSRFV As String    '  RFV Work Status
Private mAllergyVerifiedByNCID As String    '
Private mAllergyVerifiedByDate As Date  '

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



Public Property Get WSRFV() As String
    WSRFV = mWSRFV
End Property
Public Property Let WSRFV(value As String)
    mWSRFV = value
End Property
Public Property Get AllergyVerifiedByNCID() As String
    AllergyVerifiedByNCID = mAllergyVerifiedByNCID
End Property
Public Property Let AllergyVerifiedByNCID(value As String)
    mAllergyVerifiedByNCID = value
End Property
Public Property Get AllergyVerifiedByDate() As Date
    AllergyVerifiedByDate = mAllergyVerifiedByDate
End Property
Public Property Let AllergyVerifiedByDate(value As Date)
    mAllergyVerifiedByDate = value
End Property




Private Function iObjectPersist_Copy() As Object
    Dim NewCopy As RFV
    Dim State As iState
    
    Set NewCopy = New RFV
    
    With NewCopy
        .WSRFV = mWSRFV
        .AllergyVerifiedByNCID = mAllergyVerifiedByNCID
        .AllergyVerifiedByDate = mAllergyVerifiedByDate
        .Updated = mUpdated
        
        'iState interface
        Set State = NewCopy
        State.State = mState
    End With

    Set iObjectPersist_Copy = NewCopy

End Function




Private Sub iObjectPersist_Persist(Writer As iWriter)
    With Writer
        .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION
        .WriteProperty "WSRFV", mWSRFV
        .WriteProperty "AllergyVerifiedByNCID", mAllergyVerifiedByNCID
        .WriteProperty "AllergyVerifiedByDate", mAllergyVerifiedByDate
        .WriteProperty "Updated", Updated
        
        'iState interface
        .WriteProperty "iState.State", mState
    End With
End Sub




Private Function iObjectPersist_Restore(Reader As iReader) As Boolean
Dim Failed As Boolean

    On Error GoTo ErrHandler
  
    With Reader
        mWSRFV = .ReadProperty("WSRFV")
        mAllergyVerifiedByNCID = .ReadProperty("AllergyVerifiedByNCID")
        mAllergyVerifiedByDate = .ReadProperty("AllergyVerifiedByDate")
        Updated = .ReadProperty("Updated")
        
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
