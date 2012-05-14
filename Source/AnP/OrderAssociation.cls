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
Attribute VB_Name = "OrderAssociation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This class is used to link a diangosis to an order.

Option Explicit

Implements iGridRow

Public Enum enumAssociationState
  asUnknown
  asAssociate
  asAssociated
  asUnassociate
  asUnassociated
  asQueued
End Enum

Private mOrderKey As String
Private mOrderDataID As String
Private mState As enumAssociationState
Private mOrderType As enumOrderType

Private Function iGridRow_IsDiagnosis() As Boolean
  iGridRow_IsDiagnosis = False
End Function

Private Function iGridRow_IsOrderAssociation() As Boolean
  iGridRow_IsOrderAssociation = True
End Function

Public Property Get OrderDataID() As String
  OrderDataID = mOrderDataID
End Property

Public Property Let OrderDataID(ByVal NewValue As String)
  If Val(mOrderDataID) = 0 Then
    mOrderDataID = NewValue
  End If
End Property

Public Property Get OrderKey() As String
  OrderKey = mOrderKey
End Property

Public Property Let OrderKey(ByVal NewValue As String)
  If Len(mOrderKey) = 0 Then
    mOrderKey = NewValue
  End If
End Property

Public Property Get OrderName() As String
  
  On Error GoTo ErrHandler

  If Len(mOrderKey) > 0 Then
    OrderName = gBB.Order(mOrderType, mOrderKey).Name
  End If

  Exit Property
  
ErrHandler:
   OrderName = "Unidentifiable Association for DataID = " & mOrderKey
   
End Property

Public Property Let OrderType(ByVal NewValue As enumOrderType)
  mOrderType = NewValue
End Property

Public Property Get OrderType() As enumOrderType
  OrderType = mOrderType
End Property

Public Property Get RTF(ByRef Diagnosis As AnPDiagnosis) As String

  On Error GoTo ErrHandler
  
  If Not gBB.Order(mOrderType, mOrderKey) Is Nothing Then
    If Len(mOrderKey) > 0 Then
      RTF = gBB.Order(mOrderType, mOrderKey).RTF(Diagnosis)
    End If
  End If
  
  Exit Property
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "OrderAssociation.RTF", "A/P", vbExclamation
  
End Property

Public Property Get State() As enumAssociationState
  State = mState
End Property

Public Property Let State(ByVal NewValue As enumAssociationState)
Dim ValidNewState As Boolean

  Select Case mState
    Case asUnknown
      ValidNewState = True
      
    Case asAssociate
      ValidNewState = (NewValue = asAssociated Or NewValue = asUnassociate)
      
    Case asAssociated
      ValidNewState = (NewValue = asUnassociate)
      
    Case asUnassociate, asQueued
      ValidNewState = (NewValue = asAssociate Or NewValue = asUnassociate)
    
    Case asUnassociated
      ValidNewState = (NewValue = asAssociate)
  End Select

  If ValidNewState Then mState = NewValue
  
End Property

Private Function ConvertOrderTypeToModuleID(ByVal OrderType As enumOrderType) As enumAnPModules
  
  Select Case OrderType
    Case typProc
      ConvertOrderTypeToModuleID = modProcedure
      
    Case typConsult
      ConvertOrderTypeToModuleID = modConsults
      
    Case typLab
      ConvertOrderTypeToModuleID = modLab
      
    Case typRad
      ConvertOrderTypeToModuleID = modRad
      
    Case typMed
      ConvertOrderTypeToModuleID = modMed
      
    Case typPI
      ConvertOrderTypeToModuleID = modPI
  End Select
  
End Function

Public Sub Init(ByVal OrderKey As String, ByVal OrderType As enumOrderType)
  mOrderKey = OrderKey
  mOrderDataID = mOrderKey
  
  mOrderType = OrderType
  
  mState = asAssociated
End Sub

Public Function Save(ByVal DiagnosisDataID As String) As Boolean
Dim Valid As Boolean
Dim Proxy As ProxyProblems
Dim Success As Boolean

  On Error GoTo ErrHandler

  Valid = (mState = asAssociate Or mState = asUnassociate) And (Len(mOrderKey) > 0)
  
  If Not Valid Then
    Save = True
    Exit Function
  End If
  
  Set Proxy = gBB.ProblemInterface
  
  If Len(DiagnosisDataID) > 0 And Len(mOrderDataID) > 0 Then
    Select Case mState
      Case asAssociate
        Success = Proxy.AddNewAssociation(DiagnosisDataID, mOrderDataID, mOrderType)

      Case asUnassociate
        Success = Proxy.DeleteAssociation(DiagnosisDataID, mOrderDataID, mOrderType)
    End Select
  Else
    Success = True
  End If
  
  If Success Then UpdateState
  
  Save = Success
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "OrderAssociation.Save", "A/P", vbExclamation
  
End Function

Private Sub UpdateState()

  Select Case mState
    Case asAssociate
      mState = asAssociated
      
    Case asUnassociate
      mState = asUnassociated
  End Select

End Sub

Public Function UserHasPriv(ByVal PrivType As CHCSII_PrivilegeType) As Boolean
Dim ModuleID As enumAnPModules

  On Error GoTo ErrHandler

  ModuleID = ConvertOrderTypeToModuleID(mOrderType)
  UserHasPriv = gBB.AnPChild(ModuleID).UserHasPriv(PrivType)
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "OrderAssociation.UserHasPriv", "A/P", vbExclamation
  
End Function

