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
Attribute VB_Name = "Phone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--- begin DTO Builder; DTO=Phone
'--- 2/10/2004 10:30:46 AM
'---
'--- NOTE: Any time you modify this class, increment the version number
'
'

Option Explicit

Implements iObjectPersist
Implements iState

Private Const OBJECT_VERSION = "1.0"
Private Const UNIQUE_IDENTIFIER = "Phone"


Private mCountryCode As String
Private mAreaCode As String
Private mPhone As String
Private mExtension As String
Private mPhoneType As String
Private mState As enumObjectState


Public Property Get CountryCode() As String
    CountryCode = mCountryCode
End Property
Public Property Let CountryCode(ByVal value As String)
    mCountryCode = value
End Property
Public Property Get AreaCode() As String
    AreaCode = mAreaCode
End Property
Public Property Let AreaCode(ByVal value As String)
    mAreaCode = value
End Property
Public Property Get Phone() As String
    Phone = mPhone
End Property
Public Property Let Phone(ByVal value As String)
    mPhone = value
End Property
Public Property Get Extension() As String
    Extension = mExtension
End Property
Public Property Let Extension(ByVal value As String)
    mExtension = value
End Property
Public Property Get PhoneType() As String
    PhoneType = mPhoneType
End Property
Public Property Let PhoneType(ByVal value As String)
    mPhoneType = value
End Property




Private Function iObjectPersist_Copy() As Object
    Dim NewCopy As Phone
    Dim State As iState
    On Error GoTo ErrHandler
    
    Set NewCopy = New Phone
    
    With NewCopy
        .CountryCode = mCountryCode
        .AreaCode = mAreaCode
        .Phone = mPhone
        .Extension = mExtension
        .PhoneType = mPhoneType
        
        'iState interface
        Set State = NewCopy
        State.State = mState
    End With

    Set iObjectPersist_Copy = NewCopy

    Set NewCopy = Nothing
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "DTOsPhone.iObjectPersist_Persist", Err.Description, Err.HelpFile, Err.HelpContext

End Function




Private Sub iObjectPersist_Persist(Writer As iWriter)

    On Error GoTo ErrHandler
    
    With Writer
        .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION
        .WriteProperty "CountryCode", mCountryCode
        .WriteProperty "AreaCode", mAreaCode
        .WriteProperty "Phone", mPhone
        .WriteProperty "Extension", mExtension
        .WriteProperty "PhoneType", mPhoneType
        
        'iState interface
        .WriteProperty "iState.State", mState
    End With
    
    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, "DTOsPhone.iObjectPersist_Persist", Err.Description, Err.HelpFile, Err.HelpContext
End Sub




Private Function iObjectPersist_Restore(Reader As iReader) As Boolean
Dim Failed As Boolean

    On Error GoTo ErrHandler
  
    With Reader
        mCountryCode = .ReadProperty("CountryCode")
        mAreaCode = .ReadProperty("AreaCode")
        mPhone = .ReadProperty("Phone")
        mExtension = .ReadProperty("Extension")
        mPhoneType = .ReadProperty("PhoneType")
        
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
