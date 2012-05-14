VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderRad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This is sub-set of the Rad order from the OE or GEM_OE module.  This class represents
'a Rad specific to one Encounter (OE/ GEMS_OE module has all Rads for patient).
'It has 2 main functions.  One, it contains the data needed to create the RTF for the
'Encounter Doc.  Two, its state can be changed without saving immediately (cache state
'changes).OE/ GEM_OE do not provide a way to retrieve orders for an encounter so
'the data needed for the encounter document is saved in the CDR.
'Important Note:  The state of a Rad can be changed in either A/P (deleted) or in
'OE/ GEMS_OE (deleted, modified, changed from Queued to New).

Option Explicit

Implements iOrder

Private mOrder As iOrder            'Handles DataID, Key, and State code common to all Orders

Private mClinImpress As String
Private mComment As String
Private mOrderedBy As String
Private mOrderId As String
Private mOrderIEN As Long
Private mProcessPriority As Long
Private mProvider As String
Private mResults As String          'Theater Only
Private mStatus As String           'Theater Only
Private mTimingAdmTm As String
Private mTimingDur As Long
Private mTimingFreq As String
Private mTimingOneCont As Long
Private mTimingStartDt As String
Private mTimeOrdered As Date
Private mTransportMd As String
Private mWetReadWanted As Long

Private mDeletedInOE As Boolean

Private Sub Class_Initialize()
  Set mOrder = New iOrder
End Sub

Private Function iOrder_CreateAssociation() As OrderAssociation
Dim Assoc As OrderAssociation
  
  On Error GoTo ErrHandler

  Set Assoc = New OrderAssociation
  
  With Assoc
    .OrderKey = mOrder.Key
    .OrderType = typRad
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
  gShared.ShowVBError Err.Number, Err.Description, "OrderRad.iOrder_CreateAssociation", "A/P", vbExclamation

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
  
  iOrder_IsAssociated = gBB.AllDiagnoses.IsOrderAssociated(typRad, mOrder.Key, SavedOnly)
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
    
  iOrder_RTF = mOrder.Name & " (" & ConvertPriorityNumber(mProcessPriority) & ")"
                
  'scr 27286  10/18/02  Phil Crowder
  If DateDiff("d", AppointmentDate, StartDate) <> 0 Then
    iOrder_RTF = iOrder_RTF & START_DT_TAG & Format$(StartDate, "mm/dd/yyyy")
  End If
                
  If DateDiff("d", AppointmentDate, mTimeOrdered) <> 0 Then
    iOrder_RTF = iOrder_RTF & ORDER_DT_TAG & Format$(mTimeOrdered, "mm/dd/yyyy hh:mm")
  End If
    
  'scr 21524  7/17/02 Phil Crowder
  If gShared.IsAppMode(modeTheater) Then
    iOrder_RTF = iOrder_RTF & ORDER_RESULT & mResults & ORDER_RESULT_STATUS & mStatus
  End If
        
  If Trim$(Replace(mProvider, Space(1), vbNullString)) <> Trim$(Replace(mOrderedBy, Space(1), vbNullString)) Then
    iOrder_RTF = iOrder_RTF & ORDERBY_TAG & mOrderedBy
  End If
  
  If Replace(UserName, Space(1), vbNullString) <> Trim$(Replace(mProvider, Space(1), vbNullString)) Then
    iOrder_RTF = iOrder_RTF & ORDER_PROVIDER_TAG & mOrderedBy
  End If
                
  If Len(mClinImpress) > 0 Then
    iOrder_RTF = iOrder_RTF & CLN_IMPRESS_TAG & RTFCodeFilter(mClinImpress)
  End If

  If Len(mComment) > 0 Then
    iOrder_RTF = iOrder_RTF & END_LINE & " " & ORDER_COMMENT_TAG & RTFCodeFilter(mComment)
  End If

  Exit Property
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "OrderRad.iOrder_RTF", "A/P", vbExclamation

End Property

Private Property Let iOrder_State(ByVal RHS As enumOrderState)
  
  If (RHS = osNew Or RHS = osSaved) And mOrder.State = osQueued Then
    gBB.AllDiagnoses.AssociateAllQueued typRad, mOrder.Key
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
Dim Success As Boolean

  On Error GoTo ErrHandler

  Required = mOrder.State = osNew Or mOrder.State = osDelete Or mOrder.State = osModified
  
  If Not Required Then
    iOrder_Save = True
    Exit Function
  End If
  
  Select Case mOrder.State
    Case osNew
      Success = AddOEOrder
      
    Case osDelete
      Success = DeleteOEOrder
      
      '>>> SCR 32490 RWalker 3/21/03 Converted tx's to one 'generic' complex
      '>>> where the type will be inserted into metric 7 by the stop call
      If Success Then glARMStatus = enmARMTxStatusOk Else glARMStatus = enmARMTxStatusFailed  'SCR 32490 RWalker 3/21/2003
      
    Case osModified
      Success = UpdateOEOrder
  End Select

  If Success Then mOrder.Save Else mOrder.RollbackState

  If mOrder.State = osRemove Then iOrder_Unassociate

  iOrder_Save = Success
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "OrderRad.iOrder_Save", "A/P", vbExclamation

End Function

'Set the State of All Associations of me to Unassociate
Private Sub iOrder_Unassociate()
Dim Diagnoses As AnPDiagnoses

  On Error GoTo ErrHandler

  Set Diagnoses = gBB.AllDiagnoses
  Diagnoses.UnassociateFromAll typRad, mOrder.Key

  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "OrderRad.iOrder_Unassociate", "A/P", vbExclamation

End Sub

Private Function iOrder_UserHasPriv(ByVal PrivType As CHCSII_ClientSecurity.CHCSII_PrivilegeType) As Boolean

  On Error GoTo ErrHandler

  iOrder_UserHasPriv = gBB.AnPChild(modRad).UserHasPriv(PrivType)

  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "OrderRad.iOrder_UserHasPriv", "A/P", vbExclamation

End Function

Public Property Let ClinicalImpression(ByVal NewValue As String)
  mClinImpress = NewValue
End Property
      
Public Property Get ClinicalImpression() As String
  ClinicalImpression = mClinImpress
End Property

Public Property Let Comment(ByVal NewValue As String)
  mComment = NewValue
End Property

Public Property Get Comment() As String
  Comment = mComment
End Property

Public Property Let DeletedInOE(ByVal NewValue As Boolean)
  mDeletedInOE = NewValue
End Property

Public Property Let OrderedBy(ByVal NewValue As String)
  mOrderedBy = NewValue
End Property

Public Property Get OrderedBy() As String
  OrderedBy = mOrderedBy
End Property

Public Property Let OrderID(ByVal NewValue As String)
  mOrderId = NewValue
End Property

Public Property Get OrderID() As String
  OrderID = mOrderId
End Property

Public Property Let OrderIEN(ByVal NewValue As Long)
  mOrderIEN = NewValue
End Property

Public Property Get OrderIEN() As Long
  OrderIEN = mOrderIEN
End Property

Public Property Let OrderName(ByVal NewValue As String)
  mOrder.Name = NewValue
End Property

Public Property Get OrderName() As String
  OrderName = mOrder.Name
End Property

Public Property Let ProcessPriority(ByVal NewValue As Long)
  mProcessPriority = NewValue
End Property

Public Property Get ProcessPriority() As Long
  ProcessPriority = mProcessPriority
End Property

Public Property Let Provider(ByVal NewValue As String)
  mProvider = NewValue
End Property

Public Property Get Provider() As String
  Provider = mProvider
End Property

Public Property Let Results(ByVal NewValue As String)
  mResults = NewValue
End Property

Public Property Get Results() As String
  Results = mResults
End Property

'OE returns the Start Date is yyyymmdd format and it is saved to the CDR that way.
'However, GEMS_OE passes in the Start Date in mm/dd/yyyy format.
Private Property Get StartDate() As Date
Dim StartDT As String

  On Error GoTo ErrHandler
  
  If IsDate(mTimingStartDt) Then
    StartDate = mTimingStartDt
    Exit Property
  End If

  If Len(mTimingStartDt) = 8 Then
    StartDT = Mid$(mTimingStartDt, 5, 2) & "/" & Right$(mTimingStartDt, 2) & "/" & Left$(mTimingStartDt, 4)
    StartDate = CDate(StartDT)
  Else
    StartDate = Now
  End If
  
  Exit Property
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "OrderRad.StartDate", "A/P", vbExclamation
  
End Property

Public Property Let Status(ByVal NewValue As String)
  mStatus = NewValue
End Property

Public Property Get Status() As String
  Status = mStatus
End Property

Public Property Let TimingAdminTime(ByVal NewValue As String)
  mTimingAdmTm = NewValue
End Property

Public Property Get TimingAdminTime() As String
  TimingAdminTime = mTimingAdmTm
End Property

Public Property Let TimingDuration(ByVal NewValue As Long)
  mTimingDur = NewValue
End Property

Public Property Get TimingDuration() As Long
  TimingDuration = mTimingDur
End Property

Public Property Let TimingFrequence(ByVal NewValue As String)
  mTimingFreq = NewValue
End Property

Public Property Get TimingFrequence() As String
  TimingFrequence = mTimingFreq
End Property

Public Property Let TimingOneOrContinued(ByVal NewValue As Long)
  mTimingOneCont = NewValue
End Property

Public Property Get TimingOneOrContinued() As Long
  TimingOneOrContinued = mTimingOneCont
End Property

Public Property Let TimeOrdered(ByVal NewValue As Date)
  mTimeOrdered = NewValue
End Property

Public Property Get TimeOrdered() As Date
  TimeOrdered = mTimeOrdered
End Property

Public Property Let TimingStartDate(ByVal NewValue As String)
  mTimingStartDt = NewValue
End Property

Public Property Get TimingStartDate() As String
  TimingStartDate = mTimingStartDt
End Property

Public Property Let TransportionMode(ByVal NewValue As String)
  mTransportMd = NewValue
End Property
      
Public Property Get TransportionMode() As String
  TransportionMode = mTransportMd
End Property

Public Property Let WetReadingWanted(ByVal NewValue As Long)
  mWetReadWanted = NewValue
End Property
      
Public Property Get WetReadingWanted() As Long
  WetReadingWanted = mWetReadWanted
End Property

Private Function AddOEOrder() As Boolean
Dim ProxyCDR As ProxyCDRPlus
Dim DataID As String
      
  On Error GoTo ErrHandler

  Set ProxyCDR = gBB.CDRPlusInterface
  DataID = ProxyCDR.AddNewRad(Me)         'Save Order to CDR
      
  If Len(DataID) > 0 Then
    mOrder.DataID = DataID
    
    'Set the Order Data ID in all Associations of me
    gBB.AllDiagnoses.UpdateOrderDataID typRad, mOrder.Key, DataID
    
    AddOEOrder = True
  End If

  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "OrderRad.AddOEOrder", "A/P", vbExclamation

End Function

Private Function DeleteOEOrder() As Boolean
Dim ProxyOE As ProxyOE
Dim ProxyCDR As ProxyCDRPlus
Dim Success As Boolean

  On Error GoTo ErrHandler

  If Len(mOrderId) > 0 Then
    If Not mDeletedInOE Then                          'Delete Started in A/P
      Set ProxyOE = gBB.OEInterface
      Success = ProxyOE.DeleteOrder(typRad, mOrderId)
      If Success Then gsDelFromOE = "TRUE" Else gsDelFromOE = "FALSE" 'SCR 32490 RWalker 3/21/2003
    End If
  Else
    Success = True                                    'Queued Order is Deleted
  End If
  
  If (Success Or mDeletedInOE) Then                   'Deleted in OE
    If Len(mOrder.DataID) > 0 Then                    'Order is in CDR
      Set ProxyCDR = gBB.CDRPlusInterface
      Success = ProxyCDR.DeleteRad(mOrder.DataID)
      
      If Success Then gsDelFromCDR = "TRUE" Else gsDelFromCDR = "FALSE" 'SCR 32490 RWalker 3/21/2003
    Else
      Success = True                                  'Order has not been saved to CDR
    End If
  End If

  DeleteOEOrder = Success
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "OrderRad.DeleteOEOrder", "A/P", vbExclamation
  
End Function

Private Function UpdateOEOrder() As Boolean
Dim ProxyCDR As ProxyCDRPlus

  On Error GoTo ErrHandler

  If Len(mOrder.DataID) > 0 Then
    Set ProxyCDR = gBB.CDRPlusInterface
    UpdateOEOrder = ProxyCDR.UpdateRad(Me)
  Else
    UpdateOEOrder = True
  End If

  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "OrderRad.UpdateOEOrder", "A/P", vbExclamation

End Function
