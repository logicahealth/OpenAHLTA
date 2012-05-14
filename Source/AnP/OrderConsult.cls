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
Attribute VB_Name = "OrderConsult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This is sub-set of the Consult order from the consult module.  This class represents
'a consult specific to one Encounter (Consult module has all consults for patient).
'It has 2 main functions.  One, it contains the data needed to create the RTF for the
'Encounter Doc.  Two, its state can be changed without saving immediately (cache state
'changes).
'Important Note:  The Consult can be delete in either A/P or in the Consult module.

Option Explicit

Implements iOrder

Private mOrder As iOrder      'Handles DataID, Key, and State code common to all Orders

Private mClinic As String
Private mOrderBy As String
Private mOrderDate As Date
Private mPriority As String
Private mProvider As String
Private mProvisionalDiagnosis As String
Private mReferTo As String
Private mOrderId As String

Private Sub Class_Initialize()
  Set mOrder = New iOrder
End Sub

Private Function iOrder_CreateAssociation() As OrderAssociation
Dim Assoc As OrderAssociation
  
  On Error GoTo ErrHandler

  Set Assoc = New OrderAssociation
  
  With Assoc
    .OrderKey = mOrder.Key
    .OrderType = typConsult
    .OrderDataID = mOrder.DataID
    .State = asAssociate
  End With
  
  Set iOrder_CreateAssociation = Assoc
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "OrderConsult.iOrder_CreateAssociation", "A/P", vbExclamation
  
End Function

Private Property Get iOrder_DataID() As String
  iOrder_DataID = mOrder.DataID
End Property

Private Property Let iOrder_DataID(ByVal RHS As String)
  mOrder.DataID = RHS
End Property

Private Sub iOrder_Init(ByVal DataID As String, ByVal Name As String)
  mOrder.Init DataID, Name
End Sub

Private Function iOrder_IsAssociated(Optional ByVal SavedOnly As Boolean = False) As Boolean
  On Error Resume Next
  
  iOrder_IsAssociated = gBB.AllDiagnoses.IsOrderAssociated(typConsult, mOrder.Key, SavedOnly)
End Function

Private Property Let iOrder_Key(ByVal RHS As String)
  mOrder.Key = RHS
End Property

Private Property Get iOrder_Key() As String
  iOrder_Key = mOrder.Key
End Property

Private Property Let iOrder_Name(ByVal RHS As String)
  mOrder.Name = RHS
End Property

Private Property Get iOrder_Name() As String
  iOrder_Name = mOrder.Name
End Property

Private Property Get iOrder_RTF(Optional Diagnosis As AnPDiagnosis) As String
Dim UserName As String
Dim AppointmentDate As Date
Dim Enc As CHCSIIEncounterOps.Encounter

  On Error GoTo ErrHandler

  Set Enc = gBB.CurrentEncounter
  UserName = Enc.PrimaryProvider.FullName
  AppointmentDate = Enc.ApptDTS
  Set Enc = Nothing
  
  'JT SCR 44366 - encode rtf characters in user entered data
  iOrder_RTF = REFER_TO_TAG & RTFCodeFilter(mReferTo) & " (" & mPriority & ")" & _
              CLINIC_TAG & mClinic
              
  'SCR 32982  4/11/03  Phil Crowder
  'JT SCR 38146 - 8/18/03 - Always display the consult diagnosis information in the rtf
  'JT SCR 44366 - encode rtf characters in user entered data
  iOrder_RTF = iOrder_RTF & PRIM_DIAGNOSIS_TAG & RTFCodeFilter(mProvisionalDiagnosis)
  
  If DateDiff("d", AppointmentDate, mOrderDate) <> 0 Then
    iOrder_RTF = iOrder_RTF & ORDER_DT_TAG & Format$(mOrderDate, "mm/dd/yyyy hh:mm")
  End If
  
  If Trim$(Replace(mProvider, Space(1), vbNullString)) <> Trim$(Replace(mOrderBy, Space(1), vbNullString)) Then
    iOrder_RTF = iOrder_RTF & ORDERBY_TAG & mOrderBy
  End If
  
  If Replace(UserName, Space(1), vbNullString) <> Replace(mProvider, Space(1), vbNullString) Then
    iOrder_RTF = iOrder_RTF & ORDER_PROVIDER_TAG & mProvider
  End If

  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "OrderConsult.iOrder_RTF", "A/P", vbExclamation

End Property

Private Property Let iOrder_State(ByVal RHS As enumOrderState)
  mOrder.State = RHS
End Property

Private Property Get iOrder_State() As enumOrderState
  iOrder_State = mOrder.State
End Property

Private Sub iOrder_RollbackState()
'Do nothing here
End Sub

'Can only Delete Consults.  All actions are handled by the Consult module.
Private Function iOrder_Save() As Boolean
Dim ConsultProxy As ProxyConsult
Dim Success As Boolean
Dim sDataId As String

  On Error GoTo ErrHandler

  Select Case mOrder.State
    Case osDelete
      
      If Len(ConsultKey) <> 0 Then
        Set ConsultProxy = gBB.ConsultInterface
        Success = ConsultProxy.DeleteConsult(ConsultKey)
      Else
        Success = True
      End If
      
    Case Else
      Success = True
  End Select
      
  If Success Then mOrder.Save Else mOrder.RollbackState
  
  If mOrder.State = osRemove Then iOrder_Unassociate
  
  iOrder_Save = Success
      
  '>>> SCR 32490 RWalker 3/21/03 Converted tx's to one 'generic' complex
  '>>> where the type will be inserted into metric 7 by the stop call
  If Success Then glARMStatus = enmARMTxStatusOk Else glARMStatus = enmARMTxStatusFailed  'SCR 32490 RWalker 3/21/2003
    
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "OrderConsult.iOrder_Save", "A/P", vbExclamation
  
End Function

'Set the State of All Associations of me to Unassociate
Private Sub iOrder_Unassociate()
Dim Diagnoses As AnPDiagnoses

  On Error GoTo ErrHandler
  
  Set Diagnoses = gBB.AllDiagnoses
  Diagnoses.UnassociateFromAll typConsult, mOrder.Key
  
  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "OrderConsult.iOrder_Unassociate", "A/P", vbExclamation
  
End Sub

Private Function iOrder_UserHasPriv(ByVal PrivType As CHCSII_ClientSecurity.CHCSII_PrivilegeType) As Boolean

  On Error GoTo ErrHandler

  iOrder_UserHasPriv = gBB.AnPChild(modConsults).UserHasPriv(PrivType)

  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "OrderConsult.iOrder_UserHasPriv", "A/P", vbExclamation

End Function

Public Property Let Clinic(ByVal NewValue As String)
  mClinic = NewValue
End Property

Public Property Get Clinic() As String
  Clinic = mClinic
End Property

Public Property Let LongName(ByVal NewValue As String)
  mOrder.Name = NewValue
End Property

Public Property Get LongName() As String
  LongName = mOrder.Name
End Property

Public Property Let OrderBy(ByVal NewValue As String)
  mOrderBy = NewValue
End Property

Public Property Get OrderBy() As String
  OrderBy = mOrderBy
End Property

Public Property Let OrderDate(ByVal NewValue As Variant)
  If IsDate(NewValue) Then
    mOrderDate = NewValue
  End If
End Property

Public Property Get OrderDate() As Variant
  OrderDate = mOrderDate
End Property

Public Property Let Priority(ByVal NewValue As String)
  mPriority = NewValue
End Property

Public Property Get Priority() As String
  Priority = mPriority
End Property

Public Property Let Provider(ByVal NewValue As String)
  mProvider = NewValue
End Property

Public Property Get ProvisionalDiagnosis() As String
  ProvisionalDiagnosis = mProvisionalDiagnosis
End Property

Public Property Let ProvisionalDiagnosis(ByVal NewValue As String)
  mProvisionalDiagnosis = NewValue
End Property

Public Property Get Provider() As String
  Provider = mProvider
End Property

Public Property Let ReferTo(ByVal NewValue As String)
  mReferTo = NewValue
End Property

Public Property Get ReferTo() As String
  ReferTo = mReferTo
End Property

Public Property Get OrderID() As String
  OrderID = mOrderId
End Property

Public Property Let OrderID(ByVal NewValue As String)
  mOrderId = NewValue
End Property

Public Property Get ConsultKey() As String
'Returns the key used to identify this order to the consults module

   'Determine which key value to use based upon the operational mode
   If gShared.IsAppMode(modeCHCSI_GUI) And Not gShared.IsAppMode(modeITT) Then
      ConsultKey = OrderID
   Else
      ConsultKey = mOrder.DataID
   End If
   
End Property
