VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderPatientInstruc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This class represents a Patient Instruction (Other Therapies) which is stored in the
'CDRPLUS.  The PIs are generated from Medcin.
'It has 2 main functions.  One, it contains the data needed to create the RTF for the
'Encounter Doc.  Two, its state can be changed without saving immediately (cache state
'changes).

Option Explicit

Implements iOrder

Private mOrder As iOrder        'Handles DataID, Key, and State code common to all Orders

Private mSnoID As String

Private Sub Class_Initialize()
  Set mOrder = New iOrder
End Sub

Private Function iOrder_CreateAssociation() As OrderAssociation
Dim Assoc As OrderAssociation
  
  On Error GoTo ErrHandler

  Set Assoc = New OrderAssociation
  
  With Assoc
    .OrderKey = mOrder.Key
    .OrderType = typPI
    .OrderDataID = mOrder.DataID
    .State = asAssociate
  End With
  
  Set iOrder_CreateAssociation = Assoc
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "OrderPatientInstruc.iOrder_CreateAssociation", "A/P", vbExclamation
  
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
  
  iOrder_IsAssociated = gBB.AllDiagnoses.IsOrderAssociated(typPI, mOrder.Key, SavedOnly)
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
  iOrder_RTF = mOrder.Name
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
    DataID = CDRProxy.AddNewPI(SnoID)
    
    If Len(DataID) > 0 Then
      mOrder.DataID = DataID
      
      'Set the Order Data ID in all Associations of me
      gBB.AllDiagnoses.UpdateOrderDataID typPI, mOrder.Key, DataID
      
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
  gShared.ShowVBError Err.Number, Err.Description, "OrderPatientInstruc.iOrder_Save", "A/P", vbExclamation

End Function

'Set the State of All Associations of me to Unassociate
Private Sub iOrder_Unassociate()
Dim Diagnoses As AnPDiagnoses

  On Error GoTo ErrHandler

  Set Diagnoses = gBB.AllDiagnoses
  Diagnoses.UnassociateFromAll typPI, mOrder.Key

  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "OrderPatientInstruc.iOrder_Unassociate", "A/P", vbExclamation

End Sub

Private Function iOrder_UserHasPriv(ByVal PrivType As CHCSII_ClientSecurity.CHCSII_PrivilegeType) As Boolean

  On Error GoTo ErrHandler

  iOrder_UserHasPriv = gBB.AnPChild(modPI).UserHasPriv(PrivType)

  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "OrderPatientInstruc.iOrder_UserHasPriv", "A/P", vbExclamation

End Function

Public Property Let SnoID(ByVal NewValue As String)
  mSnoID = NewValue
End Property

Public Property Get SnoID() As String
  SnoID = mSnoID
End Property

Public Property Let SynonymName(ByVal NewValue As String)
  mOrder.Name = NewValue
End Property

Public Property Get SynonymName() As String
  SynonymName = mOrder.Name
End Property
