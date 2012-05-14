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
Attribute VB_Name = "EncProvProcedure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--- begin DTO Builder; DTO=EncProvProcedure
'--- 3/16/04 7:52:35 PM
'---
'--- NOTE: Any time you modify this class, increment the version number
'
'

Option Explicit

Implements iObjectPersist
Implements iState

Private Const OBJECT_VERSION = "1.0"
Private Const UNIQUE_IDENTIFIER = "EncProvProcedure"

Private mState As enumObjectState
Private mDataID As String   '
Private mSnoID As String '
Private mDescription As String  '



Public Property Get DataID() As String
    DataID = mDataID
End Property
Public Property Let DataID(value As String)
    mDataID = value
End Property
Public Property Get SnoID() As String
    SnoID = mSnoID
End Property
Public Property Let SnoID(value As String)
    mSnoID = value
End Property
Public Property Get Description() As String
    Description = mDescription
End Property
Public Property Let Description(value As String)
    mDescription = value
End Property




Private Function iObjectPersist_Copy() As Object
    Dim NewCopy As EncProvProcedure
    Dim State As iState
    
    Set NewCopy = New EncProvProcedure
    
    With NewCopy
        .DataID = mDataID
        .SnoID = mSnoID
        .Description = mDescription
        
        'iState interface
        Set State = NewCopy
        State.State = mState
    End With

    Set iObjectPersist_Copy = NewCopy

End Function




Private Sub iObjectPersist_Persist(Writer As iWriter)
    With Writer
        .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION
        .WriteProperty "DataID", mDataID
        .WriteProperty "SnoID", mSnoID
        .WriteProperty "Description", mDescription
        
        'iState interface
        .WriteProperty "iState.State", mState
    End With
End Sub




Private Function iObjectPersist_Restore(Reader As iReader) As Boolean
Dim Failed As Boolean

    On Error GoTo ErrHandler
  
    With Reader
        mDataID = .ReadProperty("DataID")
        mSnoID = .ReadProperty("SnoID")
        mDescription = .ReadProperty("Description")
        
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
