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
Attribute VB_Name = "MedFill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Implements iObjectPersist
Implements iState
Private Const OBJECT_VERSION = "1.0"
Private Const UNIQUE_IDENTIFIER = "MedFill"

Private mOrderId    As String
Private mName       As String
Private mQuantity   As String
Private mSig        As String
Private mFillDate   As Date
Private mAction     As String
Private mComment    As String
Private mUnitNumber As String   '--- SCR 79823  Sherry Wang  10/11/2005
Private mFacilityNCID As String
Private mActionOrderId As String

'** 5/11/05 - Frank
'** Added OrderIEN (new column added for Inpatient Med Order)
Private mOrderIEN As String

'** 7/19/05 - Frank
'** Addedd Refills and OrderName
Private mRefills As String
Private mOrderName As String

Private mState As enumObjectState

Public Property Get OrderId() As String
    OrderId = mOrderId
End Property
Public Property Let OrderId(ByVal value As String)
    mOrderId = value
End Property

Public Property Get OrderIEN() As String
    OrderIEN = mOrderIEN
End Property
Public Property Let OrderIEN(ByVal value As String)
    mOrderIEN = value
End Property

Public Property Get Name() As String
    Name = mName
End Property
Public Property Let Name(value As String)
    mName = value
End Property
Public Property Get Quantity() As String
    Quantity = mQuantity
End Property
Public Property Let Quantity(value As String)
    mQuantity = value
End Property
Public Property Get Action() As String
    Action = mAction
End Property
Public Property Let Action(value As String)
    mAction = value
End Property

Public Property Get Comment() As String
    Comment = mComment
End Property
Public Property Let Comment(value As String)
    mComment = value
End Property

Public Property Get Refills() As String
    Refills = mRefills
End Property
Public Property Let Refills(value As String)
    mRefills = value
End Property

Public Property Get OrderName() As String
    OrderName = mOrderName
End Property
Public Property Let OrderName(value As String)
    mOrderName = value
End Property

Public Property Get Sig() As String
    Sig = mSig
End Property
Public Property Let Sig(value As String)
    mSig = value
End Property
Public Property Get FillDate() As Date
    FillDate = mFillDate
End Property
Public Property Let FillDate(value As Date)
    mFillDate = value
End Property

Public Property Get UnitNumber() As String
    UnitNumber = mUnitNumber
End Property
Public Property Let UnitNumber(value As String)
    mUnitNumber = value
End Property

Public Property Get FacilityNCID() As String
    FacilityNCID = mFacilityNCID
End Property
Public Property Let FacilityNCID(value As String)
    mFacilityNCID = value
End Property

Public Property Get ActionOrderId() As String
    ActionOrderId = mActionOrderId
End Property
Public Property Let ActionOrderId(ByVal value As String)
    mActionOrderId = value
End Property

Private Function iObjectPersist_Copy() As Object
    Dim NewCopy As MedFill
    Dim State As iState
    
    Set NewCopy = New MedFill
    
    With NewCopy
        .OrderId = mOrderId
        .Name = mName
        .Quantity = mQuantity
        .Action = mAction
        .Sig = mSig
        .FillDate = mFillDate
        .Comment = mComment
        .OrderName = mOrderName
        .Refills = mRefills
        .UnitNumber = mUnitNumber
        .FacilityNCID = mFacilityNCID
        .ActionOrderId = mActionOrderId
        'iState interface
        Set State = NewCopy
        State.State = mState
    End With

    Set iObjectPersist_Copy = NewCopy

End Function

Private Sub iObjectPersist_Persist(Writer As iWriter)
    With Writer
        .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION
        .WriteProperty "OrderId", mOrderId
        .WriteProperty "Name", mName
        .WriteProperty "Quantity", mQuantity
        .WriteProperty "Action", mAction
        .WriteProperty "Sig", mSig
        .WriteProperty "FillDate", mFillDate
        .WriteProperty "Comment", mComment
        .WriteProperty "Refills", mRefills
        .WriteProperty "OrderName", mOrderName
        .WriteProperty "UnitNumber", mUnitNumber
        .WriteProperty "FacilityNCID", mFacilityNCID
        .WriteProperty "ActionOrderId", mActionOrderId
        'iState interface
        .WriteProperty "iState.State", mState
    End With
End Sub

Private Function iObjectPersist_Restore(Reader As iReader) As Boolean
Dim Failed As Boolean

    On Error GoTo ErrHandler
  
    With Reader
        mOrderId = .ReadProperty("OrderId")
        mName = .ReadProperty("Name")
        mQuantity = .ReadProperty("Quantity")
        mAction = .ReadProperty("Action")
        mSig = .ReadProperty("Sig")
        mFillDate = .ReadProperty("FillDate")
        mComment = .ReadProperty("Comment")
        mRefills = .ReadProperty("Refills")
        mOrderName = .ReadProperty("OrderName")
        mUnitNumber = .ReadProperty("UnitNumber")
        mFacilityNCID = .ReadProperty("FacilityNCID")
        mActionOrderId = .ReadProperty("ActionOrderId")
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

