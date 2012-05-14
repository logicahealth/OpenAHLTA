VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderMed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This is sub-set of the Med order from the OE or GEM_OE module.  This class represents
'a Med specific to one Encounter (OE/ GEMS_OE module has all Meds for patient).
'It has 2 main functions.  One, it contains the data needed to create the RTF for the
'Encounter Doc.  Two, its state can be changed without saving immediately (cache state
'changes).OE/ GEM_OE do not provide a way to retrieve orders for an encounter so
'the data needed for the encounter document is saved in the CDR.
'Important Note:  The state of a Med can be changed in either A/P (deleted) or in
'OE/ GEMS_OE (deleted, modified, changed from Queued to New).

Option Explicit

Implements iOrder

Private mOrder As iOrder              'Handles DataID, Key, and State code common to all Orders

Private mChildCap As Long
Private mComment As String
Private mOrderedBy As String
Private mOrderId As String
Private mOrderIEN As Long
Private mPharmacySent As String
Private mPharmacySentIEN As String
Private mProvider As String
Private mQuantity As Long
Private mRefills As Long
Private mSigCd As String
Private mStartDate As Date
Private mTimeOrdered As Date

Private mOverrideComment As String
Private mOverrideReason As String
Private mWarningInfo As String

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
    .OrderType = typMed
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
  gShared.ShowVBError Err.Number, Err.Description, "OrderMed.iOrder_CreateAssociation", "A/P", vbExclamation

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
  
  iOrder_IsAssociated = gBB.AllDiagnoses.IsOrderAssociated(typMed, mOrder.Key, SavedOnly)
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

Dim sRawSIG As String  ' SCR 102777, JQ, 8/10/06
Dim sSIGflag As String

  On Error GoTo ErrHandler

  Set Enc = gBB.CurrentEncounter
  UserName = Enc.PrimaryProvider.FullName
  AppointmentDate = Enc.ApptDTS
  Set Enc = Nothing
    
  ' SCR 102777, JQ, 8/10/06
    sRawSIG = mSigCd
    If sRawSIG <> "" Then
        sSIGflag = Right$(sRawSIG, 3)
    End If
    If sSIGflag = "||0" Then
        sRawSIG = Left$(sRawSIG, Len(sRawSIG) - 3)
    End If
    
  iOrder_RTF = mOrder.Name & SIG_TAG & sRawSIG & QT_TAG & mQuantity & _
                REFILL_TAG & mRefills
                
  If DateDiff("d", AppointmentDate, mStartDate) <> 0 Then
    iOrder_RTF = iOrder_RTF & START_DT_TAG & Format$(mStartDate, "mm/dd/yyyy")
  End If
    
  If DateDiff("d", AppointmentDate, mTimeOrdered) <> 0 Then
    iOrder_RTF = iOrder_RTF & ORDER_DT_TAG & Format$(mTimeOrdered, "mm/dd/yyyy hh:mm")
  End If
    
  If Trim$(Replace(mProvider, Space(1), vbNullString)) <> Trim$(Replace(mOrderedBy, Space(1), vbNullString)) Then
    iOrder_RTF = iOrder_RTF & ORDERBY_TAG & mOrderedBy
  End If
    
  If Replace(UserName, Space(1), vbNullString) <> Trim$(Replace(mProvider, Space(1), vbNullString)) Then
    iOrder_RTF = iOrder_RTF & ORDER_PROVIDER_TAG & mOrderedBy
  End If
                
  If Len(mComment) > 0 Then
    iOrder_RTF = iOrder_RTF & END_LINE & " " & ORDER_COMMENT_TAG & RTFCodeFilter(mComment)
  End If


 If LenB(gShared.CmdLineSwitches("INPATIENT2")) > 0 Then
      
      Call GetOverrideInfo
  
      If Len(mOverrideComment) > 0 Then
        iOrder_RTF = iOrder_RTF & END_LINE & " " & RTF_COMMENT_BEGIN_TAG & MED_OVERRIDE_COMMENT_TAG & RTFCodeFilter(mOverrideComment) & RTF_COMMENT_END_TAG
      End If
    
      If Len(mOverrideReason) > 0 Then
        iOrder_RTF = iOrder_RTF & END_LINE & " " & RTF_COMMENT_BEGIN_TAG & MED_OVERRIDE_REASON_TAG & RTFCodeFilter(mOverrideReason) & RTF_COMMENT_END_TAG
      End If
      
      If Len(mWarningInfo) > 0 Then
        iOrder_RTF = iOrder_RTF & END_LINE & " " & RTF_COMMENT_BEGIN_TAG & MED_WARNING_INFO_TAG & RTFCodeFilter(mWarningInfo) & RTF_COMMENT_END_TAG
      End If
      
  End If
  
  Exit Property
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "OrderMed.iOrder_RTF", "A/P", vbExclamation

End Property

Private Sub InitConn()

  On Error GoTo ErrHandler
  
  
  
  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyCDRPlus.InitConn", "A/P", vbExclamation

End Sub

Private Sub GetOverrideInfo()
 
 Dim SQLStmt As String
 Dim rs As ADODB.Recordset
 Dim mDAS As ICHCSII_DAS
 
 Dim OverrideReasonID As String
 Dim Connection As CHCSII_CONN.Conn

 If LenB(gShared.CmdLineSwitches("INPATIENT2")) > 0 Then
    
  If Not (mDAS Is Nothing) Then Exit Sub
  
    Set Connection = New CHCSII_CONN.Conn
    Set mDAS = Connection.CHCSII_DAS
  
    SQLStmt = "Select Override_comment, Override_Reason, Warnings from Medications where FDB_MEDID  = " & mOrderIEN & " and ordernumber = '" & mOrderId & "'"
    Set rs = mDAS.OpenRecordset(SQLStmt)
    If Not (rs.BOF And rs.EOF) Then
      mOverrideComment = rs.Fields("Override_comment").Value & vbNullString
      mWarningInfo = rs.Fields("Warnings").Value & vbNullString
      OverrideReasonID = "0" & rs.Fields("Override_Reason").Value
    End If
    
    If CLng("0" & OverrideReasonID) > 0 Then
        SQLStmt = "Select Reason_Description from Override_Reason where ID = " & OverrideReasonID
        Set rs = mDAS.OpenRecordset(SQLStmt)
        If Not (rs.BOF And rs.EOF) Then
            mOverrideReason = rs.Fields("Reason_Description").Value & vbNullString
        End If
    End If
    
    Set rs = Nothing
    Set mDAS = Nothing
  End If
  
End Sub

Private Property Let iOrder_State(ByVal RHS As enumOrderState)
  
  If (RHS = osNew Or RHS = osSaved) And mOrder.State = osQueued Then
    gBB.AllDiagnoses.AssociateAllQueued typMed, mOrder.Key
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
  gShared.ShowVBError Err.Number, Err.Description, "OrderMed.iOrder_Save", "A/P", vbExclamation
  
End Function

'Set the State of All Associations of me to Unassociate
Private Sub iOrder_Unassociate()
Dim Diagnoses As AnPDiagnoses

  On Error GoTo ErrHandler

  Set Diagnoses = gBB.AllDiagnoses
  Diagnoses.UnassociateFromAll typMed, mOrder.Key

  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "OrderMed.iOrder_Unassociate", "A/P", vbExclamation

End Sub

Private Function iOrder_UserHasPriv(ByVal PrivType As CHCSII_ClientSecurity.CHCSII_PrivilegeType) As Boolean

  On Error GoTo ErrHandler

  iOrder_UserHasPriv = gBB.AnPChild(modMed).UserHasPriv(PrivType)

  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "OrderMed.iOrder_UserHasPriv", "A/P", vbExclamation

End Function

Public Property Let ChildCap(ByVal NewValue As Long)
  mChildCap = NewValue
End Property

Public Property Get ChildCap() As Long
  ChildCap = mChildCap
End Property

Public Property Let Comment(ByVal NewValue As String)
  mComment = NewValue
End Property

Public Property Get Comment() As String
  Comment = mComment
End Property

Public Property Get OverrideComment() As String
  OverrideComment = mOverrideComment
End Property
Public Property Let OverrideComment(ByVal NewValue As String)
  mOverrideComment = NewValue
End Property

Public Property Get OverrideReason() As String
  OverrideReason = mOverrideReason
End Property
Public Property Let OverrideReason(ByVal NewValue As String)
  mOverrideReason = NewValue
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

Public Property Let PharmacySent(ByVal NewValue As String)
  mPharmacySent = NewValue
End Property

Public Property Get PharmacySent() As String
  PharmacySent = mPharmacySent
End Property

Public Property Let PharmacySentIEN(ByVal NewValue As String)
  mPharmacySentIEN = NewValue
End Property

Public Property Get PharmacySentIEN() As String
  PharmacySentIEN = mPharmacySentIEN
End Property

Public Property Let Provider(ByVal NewValue As String)
  mProvider = NewValue
End Property

Public Property Get Provider() As String
  Provider = mProvider
End Property

Public Property Let Quantity(ByVal NewValue As Long)
  mQuantity = NewValue
End Property

Public Property Get Quantity() As Long
  Quantity = mQuantity
End Property

Public Property Let Refills(ByVal NewValue As Long)
  mRefills = NewValue
End Property

Public Property Get Refills() As Long
  Refills = mRefills
End Property

Public Property Let SigCode(ByVal NewValue As String)
  mSigCd = NewValue
End Property

Public Property Get SigCode() As String
  SigCode = mSigCd
End Property

Public Property Let StartDate(ByVal NewValue As Date)
  mStartDate = NewValue
End Property

Public Property Get StartDate() As Date
  StartDate = mStartDate
End Property

Public Property Let TimeOrdered(ByVal NewValue As Date)
  mTimeOrdered = NewValue
End Property

Public Property Get TimeOrdered() As Date
  TimeOrdered = mTimeOrdered
End Property

Private Function AddOEOrder() As Boolean
Dim ProxyCDR As ProxyCDRPlus
Dim DataID As String
      
  On Error GoTo ErrHandler
      
  Set ProxyCDR = gBB.CDRPlusInterface
  DataID = ProxyCDR.AddNewMed(Me)           'Save Order to CDR
      
  If Len(DataID) > 0 Then
    mOrder.DataID = DataID
    
    'Set the Order Data ID in all Associations of me
    gBB.AllDiagnoses.UpdateOrderDataID typMed, mOrder.Key, DataID
    
    AddOEOrder = True
  End If

  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "OrderMed.AddOEOrder", "A/P", vbExclamation

End Function

Private Function DeleteOEOrder() As Boolean
Dim ProxyOE As ProxyOE
Dim ProxyCDR As ProxyCDRPlus
Dim Success As Boolean

  On Error GoTo ErrHandler

  Set ProxyOE = gBB.OEInterface

  If Len(mOrderId) > 0 Then                               'Delete Started in A/P
    If Not mDeletedInOE Then
      Success = ProxyOE.DeleteOrder(typMed, mOrderId, mOrder.Key)
      If Success Then gsDelFromOE = "TRUE" Else gsDelFromOE = "FALSE" 'SCR 32490 RWalker 3/21/2003
    Else
      'Order deleted in OE but Note must be added
      ProxyOE.AddDeletedMedNote
    End If
  Else
    Success = True                                        'Queued Order is Deleted
  End If
  
  If (Success Or mDeletedInOE) Then                       'Deleted in OE
    If Len(mOrder.DataID) > 0 Then                        'Order is in CDR
      Set ProxyCDR = gBB.CDRPlusInterface
      Success = ProxyCDR.DeleteMed(mOrder.DataID)
      
      If Success Then gsDelFromCDR = "TRUE" Else gsDelFromCDR = "FALSE" 'SCR 32490 RWalker 3/21/2003
    Else
      Success = True                                      'Order has not been saved to CDR
    End If
  End If

  DeleteOEOrder = Success
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "OrderMed.DeleteOEOrder", "A/P", vbExclamation
  
End Function

Private Function UpdateOEOrder() As Boolean
Dim ProxyCDR As ProxyCDRPlus

  On Error GoTo ErrHandler

  If Len(mOrder.DataID) > 0 Then
    Set ProxyCDR = gBB.CDRPlusInterface
    UpdateOEOrder = ProxyCDR.UpdateMed(Me)
  Else
    UpdateOEOrder = True
  End If

  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "OrderMed.UpdateOEOrder", "A/P", vbExclamation

End Function
