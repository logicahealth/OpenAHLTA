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
Attribute VB_Name = "InPatientOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--- begin DTO Builder; DTO=InPatientOrder
'--- 10/2/2004 5:30:54 PM
'---
'--- NOTE: Any time you modify this class, increment the version number
'
'

Option Explicit

Implements iObjectPersist
Implements iState

Private Const OBJECT_VERSION = "1.0"
Private Const UNIQUE_IDENTIFIER = "InPatientOrder"

Private mOrderId As String  '
Private mOrderName As String    '
Private mComment As String  '
Private mOrderType As String   '
Private mOrderStatus As Long '
Private mPreviousOrderStatus As Long '
Private mNurseAck As Boolean    '
Private mOrderDate As Date  '
Private mProvider As ConceptValue    '
Private mOrderExpr As Date  '
Private mSigned As Boolean  '
Private mState As enumObjectState
Private mUnitNumber As String  '
Private mEncounterNumber As String  '
Private mFacilityNCID As String  '
Private mNurseAckName As String    '
Private mNurseAckDate As Date    '
Private mPatientType As String    '
Private mDetails As String    '
Private mOrderIEN As String
Private mChildMedIEN As String
Private mIsMultiMedOrder As Boolean
Private mDataID As String
Private mAncillaryComment As String 'SCR 85951 - 1/18/06 - Frank

Private mIsPendingRefill As Boolean
Private mIsContinuousMedOrder As Boolean

Private mParentDataID As String

Public Property Get Details() As String
    Details = mDetails
End Property
Public Property Let Details(ByVal value As String)
    mDetails = value
End Property

Public Property Get PatientType() As String
    PatientType = mPatientType
End Property
Public Property Let PatientType(ByVal value As String)
    mPatientType = value
End Property

Public Property Get NurseAckDate() As Date
    NurseAckDate = mNurseAckDate
End Property
Public Property Let NurseAckDate(ByVal value As Date)
    mNurseAckDate = value
End Property


Public Property Get NurseAckName() As String
    NurseAckName = mNurseAckName
End Property
Public Property Let NurseAckName(ByVal value As String)
    mNurseAckName = value
End Property

Public Property Get FacilityNCID() As String
    FacilityNCID = mFacilityNCID
End Property
Public Property Let FacilityNCID(ByVal value As String)
    mFacilityNCID = value
End Property

Public Property Get EncounterNumber() As String
    EncounterNumber = mEncounterNumber
End Property
Public Property Let EncounterNumber(ByVal value As String)
    mEncounterNumber = value
End Property

Public Property Get UnitNumber() As String
    UnitNumber = mUnitNumber
End Property
Public Property Let UnitNumber(ByVal value As String)
    mUnitNumber = value
End Property

Public Property Get OrderId() As String
    OrderId = mOrderId
End Property
Public Property Let OrderId(ByVal value As String)
    mOrderId = value
End Property
Public Property Get OrderName() As String
    OrderName = mOrderName
End Property
Public Property Let OrderName(ByVal value As String)
    mOrderName = value
End Property
Public Property Get Comment() As String
    Comment = mComment
End Property
Public Property Let Comment(ByVal value As String)
    mComment = value
End Property
Public Property Get OrderType() As String
    OrderType = mOrderType
End Property
Public Property Let OrderType(ByVal value As String)
    mOrderType = value
End Property
Public Property Get OrderStatus() As Long
    OrderStatus = mOrderStatus
End Property
Public Property Let OrderStatus(ByVal value As Long)
    mOrderStatus = value
End Property
Public Property Get PreviousOrderStatus() As Long
    PreviousOrderStatus = mPreviousOrderStatus
End Property
Public Property Let PreviousOrderStatus(ByVal value As Long)
    mPreviousOrderStatus = value
End Property
Public Property Get NurseAck() As Boolean
    NurseAck = mNurseAck
End Property
Public Property Let NurseAck(ByVal value As Boolean)
    mNurseAck = value
End Property
Public Property Get OrderDate() As Date
    OrderDate = mOrderDate
End Property
Public Property Let OrderDate(ByVal value As Date)
    mOrderDate = value
End Property
Public Property Get Provider() As ConceptValue
    Set Provider = mProvider
End Property
Public Property Set Provider(ByVal value As ConceptValue)
    Set mProvider = value
End Property
Public Property Get OrderExpr() As Date
    OrderExpr = mOrderExpr
End Property
Public Property Let OrderExpr(ByVal value As Date)
    mOrderExpr = value
End Property
Public Property Get Signed() As Boolean
    Signed = mSigned
End Property
Public Property Let Signed(ByVal value As Boolean)
    mSigned = value
End Property

Private Sub Class_Initialize()
    Set mProvider = New ConceptValue
    mIsMultiMedOrder = False
End Sub

Private Function iObjectPersist_Copy() As Object
    Dim NewCopy As InPatientOrder
    Dim Persist As iObjectPersist
    Dim State As iState
   
    Set NewCopy = New InPatientOrder
    
    With NewCopy
        .OrderId = mOrderId
        .OrderName = mOrderName
        .Comment = mComment
        .OrderType = mOrderType
        .OrderStatus = mOrderStatus
        .NurseAck = mNurseAck
        .OrderDate = mOrderDate
        Set Persist = mProvider
        Set .Provider = Persist.Copy
        .OrderExpr = mOrderExpr
        .Signed = mSigned
        Set State = NewCopy
        State.State = mState
        .UnitNumber = mUnitNumber
        .EncounterNumber = mEncounterNumber
        .FacilityNCID = mFacilityNCID
        .NurseAckName = mNurseAckName
        .NurseAckDate = mNurseAckDate
        .PreviousOrderStatus = mPreviousOrderStatus
        .PatientType = mPatientType
        .Details = mDetails
    End With

    Set iObjectPersist_Copy = NewCopy

End Function

Private Sub iObjectPersist_Persist(Writer As iWriter)
    With Writer
        .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION
        .WriteProperty "OrderId", mOrderId
        .WriteProperty "OrderName", mOrderName
        .WriteProperty "Comment", mComment
        .WriteProperty "OrderType", mOrderType
        .WriteProperty "OrderStatus", mOrderStatus
        .WriteProperty "NurseAck", mNurseAck
        .WriteProperty "OrderDate", mOrderDate
        .WriteProperty "OrderExpr", mOrderExpr
        .WriteProperty "Signed", mSigned
        .WriteObject "Provider", mProvider
        .WriteProperty "iState.State", mState
        .WriteProperty "UnitNumber", mUnitNumber
        .WriteProperty "EncounterNumber", mEncounterNumber
        .WriteProperty "FacilityNCID", mFacilityNCID
        .WriteProperty "NurseAckName", mNurseAckName
        .WriteProperty "NurseAckDate", mNurseAckDate
        .WriteProperty "PreviousOrderStatus", mPreviousOrderStatus
        .WriteProperty "PatientType", mPatientType
        .WriteProperty "Details", mDetails
    End With
End Sub

Private Function iObjectPersist_Restore(Reader As iReader) As Boolean
Dim Failed As Boolean

    On Error GoTo ErrHandler
  
    With Reader
        mOrderId = .ReadProperty("OrderId")
        mOrderName = .ReadProperty("OrderName")
        mComment = .ReadProperty("Comment")
        mOrderType = .ReadProperty("OrderType")
        mOrderStatus = .ReadProperty("OrderStatus")
        mNurseAck = .ReadProperty("NurseAck")
        mOrderDate = .ReadProperty("OrderDate")
        mProvider = .ReadProperty("Provider")
        mOrderExpr = .ReadProperty("OrderExpr")
        mSigned = .ReadProperty("Signed")
        .ReadObject "Provider", mProvider
        mState = .ReadProperty("iState.state")
        mUnitNumber = .ReadProperty("UnitNumber")
        mEncounterNumber = .ReadProperty("EncounterNumber")
        mFacilityNCID = .ReadProperty("FacilityNCID")
        mNurseAckName = .ReadProperty("NurseAckName")
        mNurseAckDate = .ReadProperty("NurseAckDate")
        mPreviousOrderStatus = .ReadProperty("PreviousOrderStatus")
        mPatientType = .ReadProperty("PatientType")
        mDetails = .ReadProperty("Details")
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

Public Property Get OrderIEN() As String
    OrderIEN = mOrderIEN
End Property

Public Property Let OrderIEN(ByVal value As String)
    mOrderIEN = value
End Property

Public Property Get ChildMedIEN() As String
    ChildMedIEN = mChildMedIEN
End Property

Public Property Let ChildMedIEN(ByVal value As String)
    mChildMedIEN = value
End Property

Public Property Get IsMultiMedOrder() As Boolean
    IsMultiMedOrder = mIsMultiMedOrder
End Property

Public Property Let IsMultiMedOrder(ByVal value As Boolean)
    mIsMultiMedOrder = value
End Property


Public Property Get DataID() As String
    DataID = mDataID
End Property

Public Property Let DataID(ByVal value As String)
    mDataID = value
End Property

Public Property Get AncillaryComment() As String
    AncillaryComment = mAncillaryComment
End Property

Public Property Let AncillaryComment(ByVal value As String)
    mAncillaryComment = value
End Property


Public Property Get IsPendingRefill() As Boolean
    IsPendingRefill = mIsPendingRefill
End Property

Public Property Let IsPendingRefill(ByVal value As Boolean)
    mIsPendingRefill = value
End Property

Public Property Get IsContinuousMedOrder() As Boolean
    IsContinuousMedOrder = mIsContinuousMedOrder
End Property

Public Property Let IsContinuousMedOrder(ByVal value As Boolean)
    mIsContinuousMedOrder = value
End Property

Public Property Get ParentDataID() As String
    ParentDataID = mParentDataID
End Property

Public Property Let ParentDataID(ByVal value As String)
    mParentDataID = value
End Property



