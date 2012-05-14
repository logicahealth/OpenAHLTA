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
Attribute VB_Name = "Insurance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Implements iObjectPersist

Private Const OBJECT_VERSION = "1.0"
Private Const UNIQUE_IDENTIFIER = "Insurance"


Private mInsuranceName As String
Private mAddress As String
Private mPhone As String
Private mGroupNumber As String
Private mGroupName As String
Private mEmployerName As String
Private mPolicyNumber As String
Private mRelationToSponsor As String
Private mPlanEffectiveTime As String
Private mPlanExpirationTime As String



Public Property Get InsuranceName() As String
    InsuranceName = mInsuranceName
End Property
Public Property Let InsuranceName(ByVal value As String)
    mInsuranceName = value
End Property
Public Property Get Address() As String
    Address = mAddress
End Property
Public Property Let Address(ByVal value As String)
    mAddress = value
End Property
Public Property Get Phone() As String
    Phone = mPhone
End Property
Public Property Let Phone(ByVal value As String)
    mPhone = value
End Property
Public Property Get GroupNumber() As String
    GroupNumber = mGroupNumber
End Property
Public Property Let GroupNumber(ByVal value As String)
    mGroupNumber = value
End Property
Public Property Get GroupName() As String
    GroupName = mGroupName
End Property
Public Property Let GroupName(ByVal value As String)
    mGroupName = value
End Property
Public Property Get EmployerName() As String
    EmployerName = mEmployerName
End Property
Public Property Let EmployerName(ByVal value As String)
    mEmployerName = value
End Property
Public Property Get PolicyNumber() As String
    PolicyNumber = mPolicyNumber
End Property
Public Property Let PolicyNumber(ByVal value As String)
    mPolicyNumber = value
End Property
Public Property Get RelationToSponsor() As String
    RelationToSponsor = mRelationToSponsor
End Property
Public Property Let RelationToSponsor(ByVal value As String)
    mRelationToSponsor = value
End Property
Public Property Get PlanEffectiveTime() As String
    PlanEffectiveTime = mPlanEffectiveTime
End Property
Public Property Let PlanEffectiveTime(ByVal value As String)
    mPlanEffectiveTime = value
End Property
Public Property Get PlanExpirationTime() As String
    PlanExpirationTime = mPlanExpirationTime
End Property
Public Property Let PlanExpirationTime(ByVal value As String)
    mPlanExpirationTime = value
End Property




Private Function iObjectPersist_Copy() As Object
    Dim NewCopy As Insurance
    
    On Error GoTo ErrHandler
    
    Set NewCopy = New Insurance
    
    With NewCopy
        .InsuranceName = mInsuranceName
        .Address = mAddress
        .Phone = mPhone
        .GroupNumber = mGroupNumber
        .GroupName = mGroupName
        .EmployerName = mEmployerName
        .PolicyNumber = mPolicyNumber
        .RelationToSponsor = mRelationToSponsor
        .PlanEffectiveTime = mPlanEffectiveTime
        .PlanExpirationTime = mPlanExpirationTime
    End With

    Set iObjectPersist_Copy = NewCopy

    Set NewCopy = Nothing
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "DTOsInsurance.iObjectPersist_Persist", Err.Description, Err.HelpFile, Err.HelpContext

End Function




Private Sub iObjectPersist_Persist(ByRef Writer As iWriter)
    On Error GoTo ErrHandler
    
    With Writer
        .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION
        .WriteProperty "InsuranceName", mInsuranceName
        .WriteProperty "Address", mAddress
        .WriteProperty "Phone", mPhone
        .WriteProperty "GroupNumber", mGroupNumber
        .WriteProperty "GroupName", mGroupName
        .WriteProperty "EmployerName", mEmployerName
        .WriteProperty "PolicyNumber", mPolicyNumber
        .WriteProperty "RelationToSponsor", mRelationToSponsor
        .WriteProperty "PlanEffectiveTime", mPlanEffectiveTime
        .WriteProperty "PlanExpirationTime", mPlanExpirationTime
    End With
    
    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, "DTOsInsurance.iObjectPersist_Persist", Err.Description, Err.HelpFile, Err.HelpContext
End Sub




Private Function iObjectPersist_Restore(ByRef Reader As iReader) As Boolean
Dim Failed As Boolean

    On Error GoTo ErrHandler
  
    With Reader
        mInsuranceName = .ReadProperty("InsuranceName")
        mAddress = .ReadProperty("Address")
        mPhone = .ReadProperty("Phone")
        mGroupNumber = .ReadProperty("GroupNumber")
        mGroupName = .ReadProperty("GroupName")
        mEmployerName = .ReadProperty("EmployerName")
        mPolicyNumber = .ReadProperty("PolicyNumber")
        mRelationToSponsor = .ReadProperty("RelationToSponsor")
        mPlanEffectiveTime = .ReadProperty("PlanEffectiveTime")
        mPlanExpirationTime = .ReadProperty("PlanExpirationTime")
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


