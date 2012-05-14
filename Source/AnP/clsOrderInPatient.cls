VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderInPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This class represents a In Patient General Orders which is stored in the
'CDRPLUS.
'It has 2 main functions.  One, it contains the data needed to create the RTF for the
'Encounter Doc.  Two, its state can be changed without saving immediately (cache state
'changes).

Option Explicit

Implements iOrder
Private mColMethod As String
Private mComment As String
Private mFacilityNCID As String
Private mNurseACK As Boolean
Private mNurseAckDate As Date
Private mNurseAckName As String
Private mOrderDate As Date
Private mOrderedBy As String
Private mOrderStatus As Long
Private mOrderType As String
Private mSigned As Boolean
Private mUnitNumber As String
Private mOrderId As Long
Private mOrderName As String
Private mOrder As iOrder              'Handles DataID, Key, and State code common to all Orders

Private Sub Class_Initialize()
    Set mOrder = New iOrder
End Sub
Private Function iOrder_CreateAssociation() As OrderAssociation
Dim Assoc As OrderAssociation
  
  On Error GoTo ErrHandler

  Set Assoc = New OrderAssociation
  
  With Assoc
    .OrderKey = mOrder.Key
    .OrderType = typInPat
    .OrderDataID = mOrder.DataID
  
    If mOrder.State <> osQueued Then
      .State = asAssociate
    Else
      .State = asQueued
    End If
  End With
  
  Set iOrder_CreateAssociation = Assoc
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "OrderInPatient.iOrder_CreateAssociation", "A/P", vbExclamation
  
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
  
  iOrder_IsAssociated = gBB.AllDiagnoses.IsOrderAssociated(typInPat, mOrder.Key, SavedOnly)
End Function

Private Property Let iOrder_Key(ByVal RHS As String)
  mOrder.Key = RHS
End Property

Private Property Get iOrder_Key() As String
  iOrder_Key = mOrder.Key
End Property

Private Property Let iOrder_Name(ByVal RHS As String)
  mOrderName = RHS
End Property

Private Property Get iOrder_Name() As String
  iOrder_Name = mOrderName
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
      
  If Len(gShared.CmdLineSwitches("GREEN")) = 0 Then
    iOrder_RTF = mOrderName
  Else
    iOrder_RTF = mOrderName
    
    If Len(mComment) > 0 Then
      iOrder_RTF = iOrder_RTF & DIAGNOSIS_COMMENT_HEADER & RTFCodeFilter(mComment)
    End If
  End If
  
  iOrder_RTF = iOrder_RTF & ORDER_DT_TAG & Format$(OrderDate, "mm/dd/yyyy")
    
  If LenB(Trim$(mOrderedBy)) > 0 Then
    iOrder_RTF = iOrder_RTF & ORDERBY_TAG & mOrderedBy
    iOrder_RTF = iOrder_RTF & ORDER_PROVIDER_TAG & mOrderedBy
  End If
  
  If Len(mComment) > 0 And Len(gShared.CmdLineSwitches("GREEN")) = 0 Then
    iOrder_RTF = iOrder_RTF & END_LINE & " " & ORDER_COMMENT_TAG & RTFCodeFilter(mComment)
  End If

  Exit Property
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "OrderInPatient.iOrder_RTF", "A/P", vbExclamation

End Property

Private Property Let iOrder_State(ByVal RHS As enumOrderState)

  If (RHS = osNew Or RHS = osSaved) And mOrder.State = osQueued Then
    gBB.AllDiagnoses.AssociateAllQueued typInPat, mOrder.Key
  End If
  
  mOrder.State = RHS
End Property

Private Property Get iOrder_State() As enumOrderState
  iOrder_State = mOrder.State
End Property

Private Sub iOrder_RollbackState()
'Do nothing here
End Sub

Private Function iOrder_Save() As Boolean
Dim Required As Boolean
Dim CDRProxy As ProxyCDRPlus
Dim DataID As String
Dim Success As Boolean

  On Error GoTo ErrHandler

  Required = mOrder.State = osNew Or mOrder.State = osDelete
  
  If Not Required Then
    iOrder_Save = True
    Exit Function
  End If
  
  If mOrder.State = osNew Then
    Set CDRProxy = gBB.CDRPlusInterface
    DataID = CDRProxy.AddNewInPat
    
    If DataID = "Fake Save" Then
        DataID = mOrder.DataID
    End If
    
    If Val(DataID) > 0 Then
      mOrder.DataID = DataID
      
      'Set the Order Data ID in all Associations of me
      gBB.AllDiagnoses.UpdateOrderDataID typInPat, mOrder.Key, DataID
      
      Success = True
    Else
      Success = False
    End If
    
  ElseIf mOrder.State = osDelete And Len(mOrder.DataID) > 0 Then
    Set CDRProxy = gBB.CDRPlusInterface
    Success = CDRProxy.DeletePI(mOrder.DataID)
  Else
    Success = True      'Deleted but not saved to database yet
  End If

  If Success Then mOrder.Save Else mOrder.RollbackState

  If mOrder.State = osRemove Then iOrder_Unassociate

  iOrder_Save = Success

  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "OrderInPatient.iOrder_Save", "A/P", vbExclamation

  
End Function

'Set the State of All Associations of me to Unassociate
Private Sub iOrder_Unassociate()
Dim Diagnoses As AnPDiagnoses

  On Error GoTo ErrHandler

  Set Diagnoses = gBB.AllDiagnoses
  Diagnoses.UnassociateFromAll typInPat, mOrder.Key

  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "OrderInPatient.iOrder_Unassociate", "A/P", vbExclamation

End Sub

Private Function iOrder_UserHasPriv(ByVal PrivType As CHCSII_ClientSecurity.CHCSII_PrivilegeType) As Boolean

  On Error GoTo ErrHandler

  iOrder_UserHasPriv = gBB.AnPChild(modOrderSummary).UserHasPriv(PrivType)

  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "OrderInPatient.iOrder_UserHasPriv", "A/P", vbExclamation

End Function

Public Property Let CollectionMethod(ByVal NewValue As String)
  mColMethod = NewValue
End Property

Public Property Get CollectionMethod() As String
  CollectionMethod = mColMethod
End Property

Public Property Let OrderType(ByVal NewValue As String)
  mOrderType = NewValue
End Property

Public Property Get OrderType() As String
  OrderType = mOrderType
End Property

Public Property Let OrderName(ByVal NewValue As String)
  mOrderName = NewValue
End Property

Public Property Get OrderName() As String
  OrderName = mOrderName
End Property

Public Property Let OrderStatus(ByVal NewValue As Long)
  mOrderStatus = NewValue
End Property

Public Property Get OrderStatus() As Long
  OrderStatus = mOrderStatus
End Property

Public Property Let OrderID(ByVal NewValue As Long)
  mOrderId = NewValue
End Property

Public Property Get OrderID() As Long
  OrderID = mOrderId
End Property

Public Property Let OrderedBy(ByVal NewValue As String)
  mOrderedBy = NewValue
End Property

Public Property Get OrderedBy() As String
  OrderedBy = mOrderedBy
End Property

Public Property Let OrderDate(ByVal NewValue As Date)
  mOrderDate = NewValue
End Property

Public Property Get OrderDate() As Date
  OrderDate = mOrderDate
End Property

Public Property Let Comment(ByVal NewValue As String)
  mComment = NewValue
End Property

Public Property Get Comment() As String
  Comment = mComment
End Property

Public Property Let Signed(ByVal NewValue As Boolean)
  mSigned = NewValue
End Property

Public Property Get Signed() As Boolean
  Signed = mSigned
End Property

Public Property Let NurseACK(ByVal NewValue As Boolean)
  mNurseACK = NewValue
End Property

Public Property Get NurseACK() As Boolean
  NurseACK = mNurseACK
End Property

Public Property Let NurseAckDate(ByVal NewValue As Date)
  mNurseAckDate = NewValue
End Property

Public Property Get NurseAckDate() As Date
  NurseAckDate = mNurseAckDate
End Property

Public Property Let NurseAckName(ByVal NewValue As String)
  mNurseAckName = NewValue
End Property

Public Property Get NurseAckName() As String
  NurseAckName = mNurseAckName
End Property

Public Property Let UnitNumber(ByVal NewValue As String)
  mUnitNumber = NewValue
End Property

Public Property Get UnitNumber() As String
  UnitNumber = mUnitNumber
End Property

Public Property Let FacilityNCID(ByVal NewValue As String)
  mFacilityNCID = NewValue
End Property

Public Property Get FacilityNCID() As String
  FacilityNCID = mFacilityNCID
End Property

