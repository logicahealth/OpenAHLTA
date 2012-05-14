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
Attribute VB_Name = "Address"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements iObjectPersist
Implements iState

Private Const OBJECT_VERSION = "1.0"
Private Const UNIQUE_IDENTIFIER = "Address"


Private mAddrType As String
Private mCity As String
Private mCountry As String
Private mCounty As String
Private mState As String
Private mStreet1 As String
Private mStreet2 As String
Private mZip As String

Private mObjectState As enumObjectState


Public Property Get AddrType() As String
    AddrType = mAddrType
End Property
Public Property Let AddrType(ByVal value As String)
    mAddrType = value
End Property
Public Property Get City() As String
    City = mCity
End Property
Public Property Let City(ByVal value As String)
    mCity = value
End Property
Public Property Get Country() As String
    Country = mCountry
End Property
Public Property Let Country(ByVal value As String)
    mCountry = value
End Property
Public Property Get County() As String
    County = mCounty
End Property
Public Property Let County(ByVal value As String)
    mCounty = value
End Property
Public Property Get State() As String
    State = mState
End Property
Public Property Let State(ByVal value As String)
    mState = value
End Property
Public Property Get Street1() As String
    Street1 = mStreet1
End Property
Public Property Let Street1(ByVal value As String)
    mStreet1 = value
End Property
Public Property Get Street2() As String
    Street2 = mStreet2
End Property
Public Property Let Street2(ByVal value As String)
    mStreet2 = value
End Property
Public Property Get Zip() As String
    Zip = mZip
End Property
Public Property Let Zip(ByVal value As String)
    mZip = value
End Property




Private Function iObjectPersist_Copy() As Object
    Dim NewCopy As Address
    Dim State As iState
    
    On Error GoTo ErrHandler
    
    Set NewCopy = New Address
    
    With NewCopy
        .AddrType = mAddrType
        .City = mCity
        .Country = mCountry
        .County = mCounty
        .State = mState
        .Street1 = mStreet1
        .Street2 = mStreet2
        .Zip = mZip
        
        'iState Interface
        Set State = NewCopy
        State.State = mObjectState
    End With

    Set iObjectPersist_Copy = NewCopy
    
    Set NewCopy = Nothing
    Set State = Nothing
    
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "DTOsAddress.iObjectPersist_Persist", Err.Description, Err.HelpFile, Err.HelpContext

End Function




Private Sub iObjectPersist_Persist(ByRef Writer As iWriter)

    On Error GoTo ErrHandler
    
    With Writer
        .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION
        .WriteProperty "AddrType", mAddrType
        .WriteProperty "City", mCity
        .WriteProperty "Country", mCountry
        .WriteProperty "County", mCounty
        .WriteProperty "State", mState
        .WriteProperty "Street1", mStreet1
        .WriteProperty "Street2", mStreet2
        .WriteProperty "Zip", mZip
        
        'iState interface
        .WriteProperty "iState.State", mObjectState
    End With
    
    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, "DTOsClinic.iObjectPersist_Persist", Err.Description, Err.HelpFile, Err.HelpContext
End Sub




Private Function iObjectPersist_Restore(ByRef Reader As iReader) As Boolean
Dim Failed As Boolean

    On Error GoTo ErrHandler
  
    With Reader
        mAddrType = .ReadProperty("AddrType")
        mCity = .ReadProperty("City")
        mCountry = .ReadProperty("Country")
        mCounty = .ReadProperty("County")
        mState = .ReadProperty("State")
        mStreet1 = .ReadProperty("Street1")
        mStreet2 = .ReadProperty("Street2")
        mZip = .ReadProperty("Zip")
        
        'iState interface
        mObjectState = .ReadProperty("iState.State")
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
    mObjectState = NewValue
End Property

Private Property Get iState_State() As enumObjectState
    iState_State = mObjectState
End Property

Private Property Let iState_ID(ByVal RHS As String)

End Property

Private Property Get iState_ID() As String
    iState_ID = UNIQUE_IDENTIFIER
End Property
