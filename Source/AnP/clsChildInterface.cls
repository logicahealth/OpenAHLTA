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
Attribute VB_Name = "ChildInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'This class is the interface to A/P used by all A/P child modules.
'It is stateless by design.

Option Explicit

'Below are list of constants for Diagnosis, Order and Procedrue property names.
'The constants are used to reduce busgs from property name type-Os.
Private Const FLD_DATAID = "DataId"
Private Const FLD_LONGNAME = "LongName"
Private Const FLD_COMMENT = "Comment"
Private Const FLD_SnoID = "SnoID"
Private Const FLD_CPTID = "CPTID"
Private Const FLD_SYNONYM = "SynonymName"
Private Const FLD_ORDER_DT = "OrderDate"
Private Const FLD_SIG = "SigCode"
Private Const FLD_QUANTITY = "Quantity"
Private Const FLD_REFILLS = "Refills"
Private Const FLD_CHILDCAP = "ChildCap"
Private Const FLD_ORDERID = "OrderId"
Private Const FLD_ORDERIEN = "OrderIen"
Private Const FLD_ORDERNAME = "OrderName"
Private Const FLD_COLLECT = "CollectMethod"
Private Const FLD_P_PRIORITY = "ProcessPriority"
Private Const FLD_LAB_PRIORITY = "LabCollectPriority"
Private Const FLD_TIME_ONE_CONT = "TimingOnexOrContinued"
Private Const FLD_TIME_START = "TimingStartDate"
Private Const FLD_CLIN_IMPRESS = "ClinicalImpressionText"
Private Const FLD_CPT = "CPTCode"
Private Const FLD_ORDEREDBY = "OrderedBy"
Private Const FLD_PROVIDER = "OrderProvider"
Private Const FLD_START_DT = "StartDate"
Private Const FLD_REFER_TO = "ReferTo"
Private Const FLD_CLINIC = "Clinic"
Private Const FLD_PRIORITY = "Priority"
Private Const FLD_RESULTS = "Results"
Private Const FLD_RESULTS_STATUS = "Status"
Private Const FLD_PRIM_DIAG = "PrimaryDiagnosis"
' BEGIN - SCR 56524 KMonroe 11/03/04
Private Const FLD_FACILITY_NCID = "FacilityNCID"
Private Const FLD_NURSE_ACK = "NurseACK"
Private Const FLD_NURSEACK_DATE = "NurseAckDate"
Private Const FLD_NURSEACK_NAME = "NurseAckName"
Private Const FLD_ORDER_STATUS = "OrderStatus"
Private Const FLD_ORDER_TYPE = "OrderType"
Private Const FLD_SIGNED = "Signed"
Private Const FLD_UNIT_NUMBER = "UnitNumber"
'END - SCR 56524
Private mlHasTemplate As Boolean            'True if the encounter has a template that may need to be used by A/P children

Private mLock As AnPLock                   'Manages locking of A/P module and the main view

Private mAnPMainView As AnPMainView                  'Main A/P form

Private Sub Class_Terminate()
  DestroyAllReferences
End Sub

Public Property Get EncounterApptType() As String

Dim EncCurrent As CHCSIIEncounterOps.Encounter

  If Not (gBB.EncParent Is Nothing) Then
    Set EncCurrent = gBB.CurrentEncounter
    EncounterApptType = EncCurrent.ApptType
  End If

End Property
Public Property Get EncounterApptIEN() As String
Dim EncCurrent As CHCSIIEncounterOps.Encounter

  If Not (gBB.EncParent Is Nothing) Then
    Set EncCurrent = gBB.CurrentEncounter
    
    EncounterApptIEN = EncCurrent.ApptIen
  End If
End Property

'SCR 34072  3/26/03  Phil Crowder
Public Property Get EncounterClinicNCID() As String
Dim EncCurrent As CHCSIIEncounterOps.Encounter

  If Not (gBB.EncParent Is Nothing) Then
    Set EncCurrent = gBB.CurrentEncounter
    
    EncounterClinicNCID = EncCurrent.ClinicNCID
  End If
End Property

Public Property Get EncounterNumber() As String
  If Not (gBB.EncParent Is Nothing) Then
    EncounterNumber = gBB.CurrentEncounter.EncounterID
  End If
End Property

Public Property Get HasTemplate() As Boolean
  HasTemplate = mlHasTemplate
End Property

Friend Property Let HasTemplate(ByVal NewValue As Boolean)
  mlHasTemplate = NewValue
End Property

Friend Property Set MainForm(ByRef Frm As AnPMainView)
  Set mAnPMainView = Frm
End Property

'Used by OE.
Public Property Get primaryProviderIEN() As String

   On Error GoTo ErrHandler

Dim oClinicOpsEx As ClinicOpsEx
Dim oEncCurrent As CHCSIIEncounterOps.Encounter
  
   If Not (gBB.EncParent Is Nothing) Then
      Set oClinicOpsEx = New ClinicOpsEx
      Set oEncCurrent = gBB.CurrentEncounter
      
      primaryProviderIEN = oClinicOpsEx.GetProviderIEN(oEncCurrent.FacilityNCID, oEncCurrent.Providers.Item(1).NCID)
   End If
  
ErrHandler:
   'Clean Up Objects
   Set oEncCurrent = Nothing
   Set oClinicOpsEx = Nothing
   
   If CBool(Err) Then gShared.ShowVBError Err.Number, Err.Description, "ChildInterface.primaryProviderIEN", "A/P", vbExclamation
  
End Property

'SCR 33473 3/11/03  Phil Crowder  Needed by Reminders tab
Public Property Get SelectedDiagnosisComment() As String
Dim Diagnosis As AnPDiagnosis

  Set Diagnosis = mAnPMainView.SelectedDiagnosis
  
  If Not (Diagnosis Is Nothing) Then
    SelectedDiagnosisComment = Diagnosis.Comment
  End If
  
End Property

'SCR 33473 3/11/03  Phil Crowder  Needed by Reminders tab
Public Property Let SelectedDiagnosisComment(ByVal NewValue As String)
Dim Diagnosis As AnPDiagnosis

  Set Diagnosis = mAnPMainView.SelectedDiagnosis
  
  If Not (Diagnosis Is Nothing) Then
    Diagnosis.Comment = NewValue
    mAnPMainView.RefreshDiagnosis Diagnosis
  End If

End Property

Private Sub SetProviderDefault(ByVal sDataId As String, ByRef oProcedure As OrderProcedure)

On Error GoTo ErrHandler

Dim oEnc As CHCSIIEncounterOps.Encounter
Dim oProvider As EncProvider
Dim lProvIdx As Long
   
   If IsHIPAAEnabled Then
      
      'Setup the default procedure linkage if only one provider is assigned to the encounter
      Set oEnc = gBB.CurrentEncounter
   
      If oEnc.Providers.Count = 1 Then
         Set oProvider = oEnc.PrimaryProvider()
         oProvider.AssociateProcedure sDataId, oProcedure.CPTID
         oProcedure.Providers = "1"
      Else
         'SCR 49267 - If more than one provider still don't set a default but display the modifiers display
         mAnPMainView.DisplayModifiersDialog
      End If
      
   End If
      
ErrHandler:
   Set oEnc = Nothing
   Set oProvider = Nothing
   
   If CBool(Err) Then gShared.ShowVBError Err.Number, Err.Description, "ChildInterface.SetProviderDefault", "A/P", vbExclamation

End Sub

'SCR 20879  2/5/02  Phil Crowder
Public Property Get SelectedDiagnosisSnoID() As String
Dim Diagnosis As AnPDiagnosis

  If Not (mAnPMainView Is Nothing) Then
    Set Diagnosis = mAnPMainView.SelectedDiagnosis
    
    If (Not Diagnosis Is Nothing) Then
      SelectedDiagnosisSnoID = Diagnosis.SnoID
    End If
  End If
  
End Property

'Used by OE to determine if it should show lab history.
Public Property Get ShowAllLabs() As Boolean
Dim EncCurrent As CHCSIIEncounterOps.Encounter
  
  On Error GoTo ErrHandler
  
  If Not (gBB.EncParent Is Nothing) Then
    Set EncCurrent = gBB.CurrentEncounter
    
    ShowAllLabs = (EncCurrent.EncProperties.ShowActiveLabs = 1)
  End If

  Exit Property
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ChildInterface.ShowAllLabs", "A/P", vbExclamation

End Property

'Used by OE to determine if it should show med history.
Public Property Get ShowAllMeds() As Boolean
Dim EncCurrent As CHCSIIEncounterOps.Encounter
  
  On Error GoTo ErrHandler
  
  If Not (gBB.EncParent Is Nothing) Then
    Set EncCurrent = gBB.CurrentEncounter
    
    ShowAllMeds = (EncCurrent.EncProperties.ShowActiveMeds = 1)
  End If

  Exit Property
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ChildInterface.ShowAllMeds", "A/P", vbExclamation

End Property

'Used by OE to determine if it should show rad history.
Public Property Get ShowAllRads() As Boolean
Dim EncCurrent As CHCSIIEncounterOps.Encounter
  
  On Error GoTo ErrHandler
  
  If Not (gBB.EncParent Is Nothing) Then
    Set EncCurrent = gBB.CurrentEncounter
    
    ShowAllRads = (EncCurrent.EncProperties.ShowActiveRads = 1)
  End If

  Exit Property
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ChildInterface.ShowAllRads", "A/P", vbExclamation

End Property

'Used by diagnosis module to get the current encounters StartDTS value.
Public Property Get StartDTS() As Date
  
  On Error GoTo ErrHandler
  
  If Not (gBB.EncParent Is Nothing) Then
    StartDTS = gBB.CurrentEncounter.StartDTS
  End If

  Exit Property
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ChildInterface.StartDTS", "A/P", vbExclamation

End Property

'BR (SCR 29851) E-Code can not be primary diagnosis.  1/17/03  Phil Crowder
'JT SCR 40955 - added new property to indicate if diagnosis selected from problem list
Public Sub AddDiagnosis(ByVal Name As String, ByVal Prefix As String, _
                        ByVal ICD9Code As String, ByVal NCID As String, _
                        ByVal SnoID As String, ByVal Chronicity As String, _
                        ByVal DiagnosisType As String, ByVal ExtenderCode As String, _
                        ByVal ExtenderName As String, ByVal ExtenderNCID As String, _
                        Optional ByVal FromProblemList As Boolean = False)
                        
Dim Diagnosis As AnPDiagnosis
Dim ICDValid As Boolean

  On Error GoTo ErrHandler
  
  ICD9Code = Trim$(ICD9Code)
  SnoID = Trim$(SnoID)
  ExtenderCode = Trim$(ExtenderCode)
  
  If mLock Is Nothing Then
    Err.Raise 60000, , "The A/P Child Interface has not been initialized.  Can not Add New Diagnosis."
  End If
  
  If DiagnosisExists(SnoID, , Prefix) Then
    MsgBxARMd "The selected diagnosis already exists for this patient.", , "A/P"
    Exit Sub
  End If
  
  If (Len(gShared.CmdLineSwitches("WAM")) > 0 Or gShared.IsAppMode(modeCHCSI_GUI)) And gBB.AllDiagnoses.Filter(fltActive).Count = 0 And UCase$(Left$(ICD9Code, 1)) = "E" Then
    MsgBxARMd "This diagnosis is a 'cause of injury code' or 'E-code' and cannot be " & _
              "used as the primary diagnosis.  Please choose another diagnosis " & _
              "(typically describes the injury, not the cause) first, then you may " & _
              "add the injury code as a secondary diagnosis.", , "A/P"
    Exit Sub
  End If
      
    CheckForRollup ICD9Code
  
  mLock.StartLock "AddDiagnosis"
    
  Set Diagnosis = gBB.AllDiagnoses.AddDiagnosis
  
  With Diagnosis
    .State = dsNew
    
    .Name = Name
    .Prefix = Prefix
    .ICD9Code = ICD9Code
    .NCID = NCID
    .SnoID = SnoID
    .Chronicity = Chronicity
    .DiagnosisType = DiagnosisType
    .ExtenderCode = ExtenderCode
    .ExtenderName = ExtenderName
    .ExtenderNCID = ExtenderNCID
    .OnsetDate = Now
    .FacilityNCID = gBB.CurrentEncounter.FacilityNCID
    .EncounterID = gBB.CurrentEncounter.EncounterID
    .ClinicNCID = gBB.User.GetAttribute("USER_Ncid")
    .FromProblemList = FromProblemList          'JT SCR 40955
  End With
    
  If Not (mAnPMainView Is Nothing) Then mAnPMainView.AddDiagnosis
  
  'SCRs 44504, 44505, 43256, and 43257 - Determine if this is an accident related code
  If IsHIPAAEnabled Then
      If Not (mAnPMainView Is Nothing) Then mAnPMainView.CheckForAccident cstDiagnosis, ICD9Code
  End If
  
  mLock.StopLock "AddDiagnosis"
  
  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ChildInterface.AddDiagnosis", "A/P", vbExclamation
  Exit Sub
      Debug.Assert False
      Err.Clear
      Resume

End Sub

Private Function AddProcedure(ByRef Order As clsItem) As iOrder
Dim AnPOrder As OrderProcedure
Dim CPTValid As Boolean

  On Error GoTo ErrHandler
  
  Set AnPOrder = gBB.AllOrders.AddOrder(typProc)
  
  With AnPOrder
    .CPTID = Trim$(Order.PropValue(FLD_CPTID))
    .SynonymName = Order.PropValue(FLD_SYNONYM)
    .CPTCode = Trim$(Order.PropValue(FLD_CPT))
  End With
  
  Set AddProcedure = AnPOrder
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ChildInterface.AddProcedure", "A/P", vbExclamation
  
End Function

Private Function AddPI(ByRef Order As clsItem) As iOrder
Dim AnPOrder As OrderPatientInstruc

  On Error GoTo ErrHandler

  Set AnPOrder = gBB.AllOrders.AddOrder(typPI)
  
  With AnPOrder
    .SnoID = Order.PropValue(FLD_SnoID)
    .SynonymName = Order.PropValue(FLD_SYNONYM)
  End With
  
  Set AddPI = AnPOrder

  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ChildInterface.AddPI", "A/P", vbExclamation

End Function

Private Function AddConsult(ByRef Order As clsItem) As iOrder 'Delete me
'Dim AnPOrder As iOrder
''Dim Consult As OrderConsult
'
'  On Error GoTo ErrHandler
'
'  Set AnPOrder = gBB.AllOrders.AddOrder(typConsult)
'  AnPOrder.DataID = Order.PropValue(FLD_DATAID)
'
''  Set Consult = AnPOrder
'
'  With Consult
'    .LongName = Order.PropValue(FLD_LONGNAME)
'    .ReferTo = Order.PropValue(FLD_REFER_TO)
'    .Priority = Order.PropValue(FLD_PRIORITY)
'    .Clinic = Order.PropValue(FLD_CLINIC)
'    .OrderBy = Order.PropValue(FLD_ORDEREDBY)
'    .OrderDate = Order.PropValue(FLD_ORDER_DT)
'    .Provider = Order.PropValue(FLD_PROVIDER)
'    .ProvisionalDiagnosis = Order.PropValue(FLD_PRIM_DIAG)
'    .OrderID = Order.PropValue(FLD_ORDERID)
'  End With
'
'  Set AddConsult = AnPOrder
'
'  Exit Function
'
'ErrHandler:
'  gShared.ShowVBError Err.Number, Err.Description, "ChildInterface.AddConsult", "A/P", vbExclamation
'
End Function

Private Function AddMed(ByRef Order As clsItem) As iOrder
Dim AnPOrder As iOrder
Dim Med As OrderMed

  On Error GoTo ErrHandler
  
  If Order.ItemState = isQueued Then
    Set AnPOrder = gBB.AllOrders.AddOrder(typMed)
  Else
    Set AnPOrder = FindQueuedOrder(Order.PropValue(FLD_ORDERNAME), typMed)
  
    If AnPOrder Is Nothing Then
      Set AnPOrder = gBB.AllOrders.AddOrder(typMed)
    End If
  End If
  
  AnPOrder.DataID = Order.PropValue(FLD_DATAID)
  
  Set Med = AnPOrder
  
  With Med
    .StartDate = Order.PropValue(FLD_START_DT)
    .Comment = Order.PropValue(FLD_COMMENT)
    .SigCode = Order.PropValue(FLD_SIG)
    .Quantity = Order.PropValue(FLD_QUANTITY)
    .Refills = Order.PropValue(FLD_REFILLS)
    .OrderIEN = Order.PropValue(FLD_ORDERIEN)
    .OrderName = Order.PropValue(FLD_ORDERNAME)
    .OrderedBy = Order.PropValue(FLD_ORDEREDBY)
    .Provider = Order.PropValue(FLD_PROVIDER)
    .TimeOrdered = Order.PropValue(FLD_ORDER_DT)
    .ChildCap = Order.PropValue(FLD_CHILDCAP)
    .OrderID = Order.PropValue(FLD_ORDERID)
  End With
  
  Set AddMed = AnPOrder

  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ChildInterface.AddMed", "A/P", vbExclamation

End Function

Private Function AddLab(ByRef Order As clsItem) As iOrder
Dim AnPOrder As iOrder
Dim Lab As OrderLab

  On Error GoTo ErrHandler
  
  If Order.ItemState = isQueued Then
    Set AnPOrder = gBB.AllOrders.AddOrder(typLab)
  Else
    Set AnPOrder = FindQueuedOrder(Order.PropValue(FLD_ORDERNAME), typLab)
  
    If AnPOrder Is Nothing Then
      Set AnPOrder = gBB.AllOrders.AddOrder(typLab)
    End If
  End If
  
  AnPOrder.DataID = Order.PropValue(FLD_DATAID)
  
  Set Lab = AnPOrder
  
  With Lab
    .OrderIEN = Order.PropValue(FLD_ORDERIEN)
    .OrderName = Order.PropValue(FLD_ORDERNAME)
    .Comment = Order.PropValue(FLD_COMMENT)
    .CollectionMethod = Order.PropValue(FLD_COLLECT)
    .ProcessPriority = Order.PropValue(FLD_P_PRIORITY)
    .LabCollectionPriority = Order.PropValue(FLD_LAB_PRIORITY)
    '.SpecimenName = Order.PropValue(FLD_SPEC_NAME)
    '.SpecimenIEN = Order.PropValue(FLD_SPEC_IEN)
    '.SpecimenTogography = Order.PropValue(FLD_SPEC_TOP)
    '.SpecimenTogographyIEN = Order.PropValue(FLD_SPEC_TOP_IEN)
    '.SpecimenAnswerCode = Order.PropValue(FLD_SPEC_ANS_CD)
    '.SpecimenAnswerText = Order.PropValue(FLD_SPEC_ANS_TX)
    '.SpecimenContainer = Order.PropValue(FLD_SPEC_CAP)
    '.TimingOneOrContinued = Order.PropValue(FLD_TIME_ONE_CONT)
    '.TimingAdminTime = Order.PropValue(FLD_TIME_ADMIN)
    '.TimingDuration = Order.PropValue(FLD_TIME_DUR)
    '.TimingFrequence = Order.PropValue(FLD_TIME_FREQ)
    .TimingStartDate = Order.PropValue(FLD_TIME_START)
    .TimeOrdered = Order.PropValue(FLD_ORDER_DT)
    '.UnexpandedTime = Order.PropValue(FLD_UNEXP_TIME)
    .OrderedBy = Order.PropValue(FLD_ORDEREDBY)
    .Provider = Order.PropValue(FLD_PROVIDER)
    .OrderID = Order.PropValue(FLD_ORDERID)
    .Results = Order.PropValue(FLD_RESULTS)
    .Status = Order.PropValue(FLD_RESULTS_STATUS)
  End With
  
  Set AddLab = AnPOrder

  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ChildInterface.AddLab", "A/P", vbExclamation

End Function


Private Function AddRad(ByRef Order As clsItem) As iOrder
Dim AnPOrder As iOrder
Dim Rad As OrderRad

  On Error GoTo ErrHandler
  
  If Order.ItemState = isQueued Then
    Set AnPOrder = gBB.AllOrders.AddOrder(typRad)
  Else
    Set AnPOrder = FindQueuedOrder(Order.PropValue(FLD_ORDERNAME), typRad)
  
    If AnPOrder Is Nothing Then
      Set AnPOrder = gBB.AllOrders.AddOrder(typRad)
    End If
  End If
  
  AnPOrder.DataID = Order.PropValue(FLD_DATAID)
  
  Set Rad = AnPOrder
  
  With Rad
    .OrderIEN = Order.PropValue(FLD_ORDERIEN)
    .Comment = Order.PropValue(FLD_COMMENT)
    .ClinicalImpression = Order.PropValue(FLD_CLIN_IMPRESS)
    '.TimingOneOrContinued = Order.PropValue(FLD_TIME_ONE_CONT)
    '.TransportionMode = Order.PropValue(FLD_TRANS_MD)
    '.TimingAdminTime = Order.PropValue(FLD_TIME_ADMIN)
    '.TimingDuration = Order.PropValue(FLD_TIME_DUR)
    '.TimingFrequence = Order.PropValue(FLD_TIME_FREQ)
    .TimingStartDate = Order.PropValue(FLD_TIME_START)
    .OrderName = Order.PropValue(FLD_ORDERNAME)
    .OrderedBy = Order.PropValue(FLD_ORDEREDBY)
    .Provider = Order.PropValue(FLD_PROVIDER)
    .TimeOrdered = Order.PropValue(FLD_ORDER_DT)
    .ProcessPriority = Order.PropValue(FLD_P_PRIORITY)
    .OrderID = Order.PropValue(FLD_ORDERID)
    .Results = Order.PropValue(FLD_RESULTS)
    .Status = Order.PropValue(FLD_RESULTS_STATUS)
  End With
  
  Set AddRad = AnPOrder
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ChildInterface.AddRad", "A/P", vbExclamation
  
End Function
' BEGIN - SCR 56524 KMonroe 11/03/04
Private Function AddInPat(ByRef Order As clsItem) As iOrder
Dim AnPOrder As iOrder
Dim OrderInPat As OrderInPatient

  On Error GoTo ErrHandler
  
  If Order.ItemState = isQueued Then
    Set AnPOrder = gBB.AllOrders.AddOrder(typInPat)
  Else
    Set AnPOrder = FindQueuedOrder(Order.PropValue(FLD_ORDERNAME), typInPat)
  
    If AnPOrder Is Nothing Then
      Set AnPOrder = gBB.AllOrders.AddOrder(typInPat)
    End If
  End If
  
  AnPOrder.DataID = Order.PropValue(FLD_DATAID)
  
  Set OrderInPat = AnPOrder
  
  With OrderInPat
    .OrderID = Order.PropValue(FLD_DATAID)
    .OrderName = Order.PropValue(FLD_ORDERNAME)
    .FacilityNCID = Order.PropValue(FLD_FACILITY_NCID)
    .Comment = Order.PropValue(FLD_COMMENT)
    .NurseACK = Order.PropValue(FLD_NURSE_ACK)
    .NurseAckDate = Order.PropValue(FLD_NURSEACK_DATE)
    .NurseAckName = Order.PropValue(FLD_NURSEACK_NAME)
    .OrderedBy = Order.PropValue(FLD_ORDEREDBY)
    .OrderDate = Order.PropValue(FLD_ORDER_DT)
    .OrderStatus = Order.PropValue(FLD_ORDER_STATUS)
    .OrderType = Order.PropValue(FLD_ORDER_TYPE)
    .Signed = Order.PropValue(FLD_SIGNED)
    .UnitNumber = Order.PropValue(FLD_UNIT_NUMBER)
  End With
  
  Set AddInPat = AnPOrder
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ChildInterface.AddInPat", "A/P", vbExclamation
  
End Function
'END SCR 56524

'Used by procedure and patient instruction modules to add orders to the current
'collection of orders for the current encounter and assocaite that order with the
'currently selected diagnosis (the main view is used to determine the currently
'selected diagnosis).
'Parameters:
' Order    The A/P child module creates the object.  Then passes the object to this module.
Public Sub AddOrder(ByRef Order As clsItem, Optional ByRef Tablet As Object)
Dim AnPOrder As iOrder

  On Error GoTo ErrHandler
  
  If mLock Is Nothing Or gBB.AllOrders Is Nothing Then
    Err.Raise 60001, , "The A/P Child Interface has not been initialized.  Can not Add New Order."
  End If

  mLock.StartLock "AddOrder"
  
  Select Case Order.ItemType
    Case itConsult
      Set AnPOrder = AddConsult(Order)
      
    Case itProcedure
      Set AnPOrder = AddProcedure(Order)
    
    Case itPatientInstr
      Set AnPOrder = AddPI(Order)
    
    Case itLab
      Set AnPOrder = AddLab(Order)
    
    Case itMed
      Set AnPOrder = AddMed(Order)
    
    Case itRad
      Set AnPOrder = AddRad(Order)
      
    Case itInPatient
      Set AnPOrder = AddInPat(Order)

  End Select
  
  If Not (AnPOrder Is Nothing) Then
    SetOrderState ConvertItemStateToOrderState(Order.ItemState), AnPOrder, Tablet
      
    AssociateOrderToDiagnosis AnPOrder
    
    'SCR 49267 - Modifiers dialog expects the grid row to exist so this call needed to be moved here
    If Order.ItemType = itProcedure And IsHIPAAEnabled Then SetProviderDefault Order.DataID, AnPOrder
    
  End If
  
  mLock.StopLock "AddOrder"
  
  Select Case Order.ItemType
    Case itLab, itMed, itRad
      gBB.Comm.Message cwiOE_NEW, "RECEIEVED ORDER", NCID_BROADCAST_MSG, NCID_OE_MED_APPLICATION
  End Select
  
  Exit Sub
  
ErrHandler:
  mLock.StopLock "AddOrder"
  gShared.ShowVBError Err.Number, Err.Description, "ChildInterface.AddOrder", "A/P", vbExclamation
  
End Sub

'BR (28153)  Display warning for procedures  1/20/02 Phil Crowder
Private Sub AssociateOrderToDiagnosis(ByRef Order As iOrder)
Dim Diagnosis As AnPDiagnosis
Dim Association As OrderAssociation
  
  If Not (mAnPMainView Is Nothing) Then
    Set Diagnosis = mAnPMainView.SelectedDiagnosis
    
    If Not (Diagnosis Is Nothing) Then
      Set Association = Order.CreateAssociation
            
      Diagnosis.AssociateOrder Association
      
      mAnPMainView.RefreshDiagnosis Diagnosis
    End If
    
    mAnPMainView.AddOrder Order
  End If

End Sub

Private Sub CheckForRollup(ByVal ICDCode As String)
Dim Diagnoses As Collection
Dim Diagnosis As AnPDiagnosis

  Set Diagnoses = gBB.AllDiagnoses.Filter(fltActive)
  
  For Each Diagnosis In Diagnoses
    If Diagnosis.FullICDCode = ICDCode Then
      WarningDialog.Display "This diagnosis has the same ICD as another selected diagnosis.  When the ADM record is Filed, all duplicate ICDs will be rolled up to the Highest priority ICD.", True
      Exit For
    End If
  Next
  
End Sub

Private Function ConvertItemStateToOrderState(ByRef ItemState As enumItemState) As enumOrderState
  Select Case ItemState
    Case isUnknown
      ConvertItemStateToOrderState = osUnknown
    Case isNew
      ConvertItemStateToOrderState = osNew
    Case isQueued
      ConvertItemStateToOrderState = osQueued
    Case isSaved
      ConvertItemStateToOrderState = ossaved
    Case isDelete
      ConvertItemStateToOrderState = osDelete
    Case isModified
      ConvertItemStateToOrderState = osModified
    Case isRemove
      ConvertItemStateToOrderState = osRemove
    Case Else
      ConvertItemStateToOrderState = osUnknown
  End Select
End Function

Private Sub SetOrderState(ByVal NewState As enumOrderState, ByRef Order As iOrder, ByRef Tablet As Object)
Dim OrderSets As Object
Dim ErrorMessage As String

  Select Case NewState
    Case osQueued
      Set OrderSets = gBB.AnPChild(modOrderSets)
      
      'Queue up Order in Order Sets module
      If OrderSets.Server Is Nothing Then
        OrderSets.InitView
        If cwbMainViewRunning Then gBB.AnPChild(modOrderSets).OpenView mAnPMainView.vseContainer.hWnd
      End If
      OrderSets.Server.Mtd_bPushOrder Tablet, "QID_" & Order.Key, ErrorMessage
      
      Order.State = osQueued
      
    Case ossaved
      Order.State = ossaved
      
    Case Else
      Order.State = osNew
  End Select

End Sub

'Used by Diagnosis module if the current encounter is a telcon encounter.
'Returns:
  '2D array of diagnosis names, SnoIDs and NCIDs from the collection
  'of active diagnosis
Public Function AllDiagnosisNames() As Variant
Dim Diagnoses As Collection
Dim Diagnosis As AnPDiagnosis
Dim ListOfNames() As String
Dim i As Integer
  
  On Error GoTo ErrHandler
  
  Set Diagnoses = gBB.AllDiagnoses.Filter(fltActive)
  
  ReDim ListOfNames(2, Diagnoses.Count)
  
  For i = 1 To Diagnoses.Count
    Set Diagnosis = Diagnoses.Item(CStr(i))     'SCR 33139  3/3/03  Phil Crowder  Load array by Priority order
    
    Select Case Diagnosis.State
      Case dsNew, dsModified, dsSaved
        ListOfNames(0, i - 1) = Diagnosis.Name
        ListOfNames(1, i - 1) = Diagnosis.SnoID
        ListOfNames(2, i - 1) = Diagnosis.NCID
    End Select
  Next
  
  AllDiagnosisNames = ListOfNames
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ChildInterface.AllDiagnosisNames", "A/P", vbExclamation
  
End Function

'Used by Orders Sets module to force a load of a module if it is not loaded.
'Parameter:
' TabIndex    Corresponds to the tab index for tabAnPChildren on AnPMainView.frm (main view).
Public Sub ChangeTab(ByVal TabIndex As Integer)
  If Not (mAnPMainView Is Nothing) Then mAnPMainView.tabAnPChildren.CurrTab = TabIndex
End Sub

'>>> SCR 38383 MSolano 7/8/2003
'Used by children of AnP to avoid refreshing Template_Nav.ocx
Public Function GetCurrentTab() As Integer
  If Not (mAnPMainView Is Nothing) Then
    GetCurrentTab = mAnPMainView.tabAnPChildren.CurrTab
  End If
End Function
'<<< SCR 38383

Friend Sub DestroyAllReferences()
  On Error Resume Next
  
  Set mLock = Nothing
  Set mAnPMainView = Nothing
  
End Sub


'Used by the diagnosis module to get the total number of active diagnoses for the
'current encounter.
'Returns:
' Number of active diagnosis
Public Function DiagnosisCount() As Long
Dim Diagnoses As Collection

  On Error GoTo ErrHandler
    
  Set Diagnoses = gBB.AllDiagnoses.Filter(fltActive)
  
  DiagnosisCount = Diagnoses.Count
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnP.DiagnosisCount", "A/P", vbExclamation
  
End Function

'--- SCR 36315 MSolano 5/21/03 method signature SnoID As String
Public Function DiagnosisExists(Optional ByVal SnoID As String = vbNullString, _
                                Optional ByVal SynonymName As String = vbNullString, _
                                Optional ByVal Prefix As String = vbNullString, _
                                Optional ByVal SelectIt As Boolean = False) As Boolean
Dim Diagnoses As Collection
Dim Diagnosis As AnPDiagnosis

  On Error GoTo ErrHandler
    
  Set Diagnoses = gBB.AllDiagnoses.Filter(fltActive)
  
  For Each Diagnosis In Diagnoses
    If Len(SnoID) Then       'SCR 36315 MSolano 5/21/03
      'SCR 20414 1/31/02  Phil Crowder  Need to check for prefix, combo is unique
      If Diagnosis.SnoID = CStr(SnoID) And Diagnosis.Prefix = Trim$(Prefix) Then
        DiagnosisExists = True
        Exit For
      End If
      
    Else
      'SCR 19767  12/17/01  Phil Crowder   Older templates saved all diagnosis in upper case
      If UCase$(Diagnosis.Name) = UCase$(SynonymName) Then
        DiagnosisExists = True
        Exit For
      End If
    End If
  Next
  
  If DiagnosisExists And SelectIt Then
    mAnPMainView.RefreshDiagnosis Diagnosis
  End If
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnP.DiagnosisExists", "A/P", vbExclamation
  
End Function

Private Function FindQueuedOrder(ByVal OrderName As String, ByVal OrderType As enumOrderType) As iOrder
Dim Orders As Collection
Dim Order As iOrder

  On Error GoTo ErrHandler
  
  Set Orders = gBB.AllOrders.Filter(OrderType, fltQueued)
  
  For Each Order In Orders
    If UCase$(Order.Name) = UCase$(OrderName) Then
      Set FindQueuedOrder = Order
      Exit For
    End If
  Next
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ChildInterface.FindQueuedOrder", "A/P", vbExclamation
  
End Function

Friend Sub Init(ByRef LockObj As AnPLock)

  On Error GoTo ErrHandler
  
  Set mLock = LockObj
  
  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ChildInterface.Init", "A/P", vbExclamation
                
End Sub

'Used by other modules to lock the A/P module during some lengthy processing.
'Parameters:
' ModuleName    Used to identify who has locked A/P.  Should be in the format of
'               Module.Procedure
Public Sub LockAnP(ByVal ModuleName As String)
  If Not (mLock Is Nothing) Then mLock.StartLock ModuleName
End Sub

'This method is called by the Procedure and Patient Instructions modules to determine
'the number of active orders of a specified type in the collection of orders for the
'current encounter.
'Paramters:
' ItemType    Optional parameter to filter search
'Returns:
' Number of active orders
Public Function OrderCount(Optional ByVal ItemType As enumItemType = itUndefined) As Long
Dim OrderType As enumOrderType

  On Error GoTo ErrHandler
    
  OrderType = CovertItemTypeToOrderType(ItemType)
  
  OrderCount = gBB.AllOrders.Filter(OrderType, fltActive).Count
  
  Exit Function
   
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ChildInterface.OrderCount", "A/P", vbExclamation
  
End Function

'SCR 59845 12/03/04 Frank Lourng
'This method is called by the Med Order tab of GEM_OE modules to determine
'patient pregnancy status.
'Paramters:
'Returns:
' True if Pregnant
Public Function IsPregnant() As Boolean

  On Error GoTo ErrHandler
    
  IsPregnant = GetPregnancyStatus
   
  Exit Function
   
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ChildInterface.IsPregnant", "A/P", vbExclamation
  
End Function

'Used to determine if an order is already part of the collection of active orders
'for the current encounter.
'Parameters:
' IdField       Field used to locate an order
' IdValue       Value for the field to find
'Returns:
' Key if found and null string if not found
Public Function OrderExists(ByVal IdField As String, ByVal IdValue As Variant) As String
Dim Key As String
Dim AllOrders As Collection
Dim Order As iOrder

  On Error GoTo ErrHandler
  
  '@@@@ Searching on name only.
  Set AllOrders = gBB.AllOrders.Filter(, fltActive)
  
  For Each Order In AllOrders
    If UCase$(Order.Name) = UCase$(IdValue) Then
      Key = Order.Key
      Exit For
    End If
  Next
  
  OrderExists = Key
  
  Exit Function
   
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ChildInterface.OrderExists", "A/P", vbExclamation
  
End Function

'This is used by the diagnosis module for a telcon encounter.  This method removes
'a diagnosis from the active diagnosis collection for the current encounter.
'Paramters:
' SnoID    Used to find the diagnosis
' Changes:
'     SCR 36315 MSolano 5/21/03 method signature SnoID As String
Public Sub RemoveDiagnosis(ByVal SnoID As String)

  Dim Diagnoses As Collection
  Dim Diagnosis As AnPDiagnosis

  On Error GoTo ErrHandler
  
  If mLock Is Nothing Then
    Err.Raise 60005, , "The A/P Child Interface has not been initialized.  Can not Remove Diagnosis."
  End If

  mLock.StartLock "RemoveDiagnosis"
  
  Set Diagnoses = gBB.AllDiagnoses.Filter(fltActive)
  
  For Each Diagnosis In Diagnoses
    If Diagnosis.SnoID = SnoID Then
      Diagnosis.State = dsDelete
      gBB.AllDiagnoses.UpdatePriorities Diagnosis.Priority    'SCR 35345  5/1/03 Phil Crowder
      
      If Not (mAnPMainView Is Nothing) Then
        mAnPMainView.RefreshData
      End If
      
      Exit For
    End If
  Next
  
  mLock.StopLock "RemoveDiagnosis"
  
  Exit Sub
   
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ChildInterface.RemoveDiagnosis", "A/P", vbExclamation
  
End Sub

'This method is used by the consults module.  This method removes an order from the active
'orders and unassociates it from all diagnoses.
'Paramters:
' OrderId    Used to find the order
Public Sub RemoveOrder(ByVal OrderID As String, ByVal ItemType As enumItemType)
Dim i As Long
Dim Orders As Collection
Dim Rad As OrderRad
Dim Lab As OrderLab
Dim Med As OrderMed
'Dim Consult As OrderConsult
Dim Order As iOrder
Dim InPat As OrderInPatient ' SCR 56524 KMonroe 11/03/04

  On Error GoTo ErrHandler
  
  If mLock Is Nothing Or gBB.AllOrders Is Nothing Then
    Err.Raise 60006, , "The A/P Child Interface has not been initialized.  Can not Remove Order."
  End If

  mLock.StartLock "RemoveOrder"
  
  Select Case ItemType
    Case itConsult
'      Set Orders = gBB.AllOrders.Filter(typConsult, fltActive)
'
'      For i = 1 To Orders.Count
'
'        Set Consult = Orders.Item(i)
'        If Consult.ConsultKey = OrderID Then
'          Set Order = Orders.Item(i)
'          Exit For
'        End If
'
'      Next
    
    Case itRad    'SCR 25928  Phil Crowder 7/23/02  For GEMS_OE
      Set Orders = gBB.AllOrders.Filter(typRad, fltActive)
      
      For i = 1 To Orders.Count
        Set Rad = Orders.Item(i)
        
        If Rad.OrderID = OrderID Then
          Set Order = Orders.Item(i)
          Exit For
        End If
      Next
    
    Case itLab    'SCR 25928  Phil Crowder 7/23/02  For GEMS_OE
      Set Orders = gBB.AllOrders.Filter(typLab, fltActive)
      
      For i = 1 To Orders.Count
        Set Lab = Orders.Item(i)
        
        If Lab.OrderID = OrderID Then
          Set Order = Orders.Item(i)
          Exit For
        End If
      Next
    
    Case itMed    'SCR 25928  Phil Crowder 7/23/02  For GEMS_OE
      Set Orders = gBB.AllOrders.Filter(typMed, fltActive)
      
      For i = 1 To Orders.Count
        Set Med = Orders.Item(i)
        
        If Med.OrderID = OrderID Then
          Set Order = Med
          
          gBB.EncounterInterface.AddEncounterNote "Discontinued " & _
                                          Order.Name & " - " & Med.SigCode
          Exit For
        End If
      Next
      
     Case itInPatient ' BEGIN - SCR 56524 KMonroe 11/03/04
       Set Orders = gBB.AllOrders.Filter(typInPat, fltActive)
      
      For i = 1 To Orders.Count
        Set InPat = Orders.Item(i)
        
        If InPat.OrderID = OrderID Then
          Set Order = Orders.Item(i)
          Exit For
        End If
      Next
'         gBB.AllOrders.RemoveOrder OrderID
  End Select
    
  If Not (Order Is Nothing) Then
    Order.State = osRemove
    Order.Unassociate
    
    'Need to save here so the user does not see the order as unassoicated when they try to close.
    gBB.EncounterInterface.Save
                          
    If Not (mAnPMainView Is Nothing) Then mAnPMainView.RefreshData
  End If

  mLock.StopLock "RemoveOrder"
  
  Exit Sub
   
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ChildInterface.RemoveOrder", "A/P", vbExclamation
  
End Sub

'This method is used by diagnosis module to get a reference to list of Acute problems for the current
'patient.
'Returns:
'   Instance of ProblemOps class with Acute problems loaded
Public Function RetrieveAcuteProblemList() As ProblemOps

  On Error GoTo ErrHandler
        
  Set RetrieveAcuteProblemList = gBB.ReadModule(NCID_PROBLEMS_APPLICATION).ProblemOps
  
  Exit Function
   
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ChildInterface.RetrieveAcuteProblemList", "A/P", vbExclamation

End Function

'This method is used by diagnosis module to get a reference to list of Chronic problems
'for the current patient.
'Returns:
'   Instance of ProblemOps class with Chronic problems loaded
Public Function RetrieveChronicProblemList() As ProblemOps

  On Error GoTo ErrHandler
        
  Set RetrieveChronicProblemList = gBB.ReadModule(NCID_PROBLEMS_APPLICATION).ProblemOps
  
  Exit Function
   
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ChildInterface.RetrieveChronicProblemList", "A/P", vbExclamation
  
End Function

'Used by other modules to unlock the A/P module if they locked A/P.
'Parameters:
' ModuleName    Used to identify who has locked A/P.  Should be in the format of
'               Module.Procedure
Public Sub UnlockAnP(ByVal ModuleName As String)
  If Not (mLock Is Nothing) Then mLock.StopLock ModuleName
End Sub

'Used by Theater OE to update the data for a modified order.
'Parameter:
' Order   Order object that is going to replace the current order object in A/P
Public Sub UpdateOrder(ByRef Order As clsItem)
Dim i As Long
Dim Orders As Collection
Dim Rad As OrderRad
Dim Lab As OrderLab
Dim Med As OrderMed
Dim InPat As OrderInPatient ' SCR 56524 KMonroe 11/03/04

Dim APOrder As iOrder

  On Error GoTo ErrHandler
  
  If gBB.AllOrders Is Nothing Then
    Err.Raise 60007, , "The A/P Child Interface has not been initialized.  Can not Update Order."
  End If
  
  Select Case Order.ItemType
    Case itLab
      Set Orders = gBB.AllOrders.Filter(typLab, fltActive)
      
      For i = 1 To Orders.Count
        Set Lab = Orders.Item(i)
        
        If Lab.OrderID = Order.PropValue(FLD_ORDERID) Then
          Set APOrder = Lab
          Exit For
        End If
      Next
      
      If Not (Lab Is Nothing) Then
        With Lab
          .OrderIEN = Order.PropValue(FLD_ORDERIEN)
          .OrderName = Order.PropValue(FLD_ORDERNAME)
          .Comment = Order.PropValue(FLD_COMMENT)
          .CollectionMethod = Order.PropValue(FLD_COLLECT)
          .ProcessPriority = Order.PropValue(FLD_P_PRIORITY)
          .LabCollectionPriority = Order.PropValue(FLD_LAB_PRIORITY)
          .TimingStartDate = Order.PropValue(FLD_TIME_START)
          .TimeOrdered = Order.PropValue(FLD_ORDER_DT)
          .OrderedBy = Order.PropValue(FLD_ORDEREDBY)
          .Provider = Order.PropValue(FLD_PROVIDER)
          .OrderID = Order.PropValue(FLD_ORDERID)
          .Results = Order.PropValue(FLD_RESULTS)
          .Status = Order.PropValue(FLD_RESULTS_STATUS)
        End With
      End If
      
    Case itMed
      Set Orders = gBB.AllOrders.Filter(typMed, fltActive)
      
      For i = 1 To Orders.Count
        Set Med = Orders.Item(i)
        
        If Med.OrderID = Order.PropValue(FLD_ORDERID) Then
          Set APOrder = Med
          Exit For
        End If
      Next
      
      If Not (Med Is Nothing) Then
        With Med
          .StartDate = Order.PropValue(FLD_START_DT)
          .Comment = Order.PropValue(FLD_COMMENT)
          .SigCode = Order.PropValue(FLD_SIG)
          .Quantity = Order.PropValue(FLD_QUANTITY)
          .Refills = Order.PropValue(FLD_REFILLS)
          .OrderIEN = Order.PropValue(FLD_ORDERIEN)
          .OrderName = Order.PropValue(FLD_ORDERNAME)
          .OrderedBy = Order.PropValue(FLD_ORDEREDBY)
          .Provider = Order.PropValue(FLD_PROVIDER)
          .TimeOrdered = Order.PropValue(FLD_ORDER_DT)
          .ChildCap = Order.PropValue(FLD_CHILDCAP)
          .OrderID = Order.PropValue(FLD_ORDERID)
        End With
      End If
      
    Case itRad
      Set Orders = gBB.AllOrders.Filter(typRad, fltActive)
      
      For i = 1 To Orders.Count
        Set Rad = Orders.Item(i)
        
        If Rad.OrderID = Order.PropValue(FLD_ORDERID) Then
          Set APOrder = Rad
          Exit For
        End If
      Next
       
      If Not (Rad Is Nothing) Then
        With Rad
          .OrderIEN = Order.PropValue(FLD_ORDERIEN)
          .Comment = Order.PropValue(FLD_COMMENT)
          .ClinicalImpression = Order.PropValue(FLD_CLIN_IMPRESS)
          .TimingOneOrContinued = Order.PropValue(FLD_TIME_ONE_CONT)
          .TimingStartDate = Order.PropValue(FLD_TIME_START)
          .OrderName = Order.PropValue(FLD_ORDERNAME)
          .OrderedBy = Order.PropValue(FLD_ORDEREDBY)
          .Provider = Order.PropValue(FLD_PROVIDER)
          .TimeOrdered = Order.PropValue(FLD_ORDER_DT)
          .ProcessPriority = Order.PropValue(FLD_P_PRIORITY)
          .OrderID = Order.PropValue(FLD_ORDERID)
          .Results = Order.PropValue(FLD_RESULTS)
          .Status = Order.PropValue(FLD_RESULTS_STATUS)
        End With
      End If
   ' BEGIN - SCR 56524 KMonroe 11/03/04
    Case itInPatient
      Set Orders = gBB.AllOrders.Filter(typInPat, fltActive)
      
      For i = 1 To Orders.Count
        Set InPat = Orders.Item(i)
        
        If InPat.OrderID = Order.PropValue(FLD_ORDERID) Then
          Set APOrder = InPat
          Exit For
        End If
      Next
       
      If Not (InPat Is Nothing) Then
        With InPat
          .OrderName = Order.PropValue(FLD_ORDERNAME)
          .FacilityNCID = Order.PropValue(FLD_FACILITY_NCID)
          .Comment = Order.PropValue(FLD_COMMENT)
          .NurseACK = Order.PropValue(FLD_NURSE_ACK)
          .NurseAckDate = Order.PropValue(FLD_NURSEACK_DATE)
          .NurseAckName = Order.PropValue(FLD_NURSEACK_NAME)
          .OrderedBy = Order.PropValue(FLD_ORDEREDBY)
          .OrderDate = Order.PropValue(FLD_ORDER_DT)
          .OrderStatus = Order.PropValue(FLD_ORDER_STATUS)
          .OrderType = Order.PropValue(FLD_ORDER_TYPE)
          .Signed = Order.PropValue(FLD_SIGNED)
          .UnitNumber = Order.PropValue(FLD_UNIT_NUMBER)
        End With
      End If
      ' BEGIN - SCR 56524 KMonroe 11/03/04
  End Select
 
  'SCR 53391 - Update order state to force save to be process to update RTF when changes have been made
  If Not APOrder Is Nothing Then
    If APOrder.State <> osUnknown And Not APOrder.State = osNew Then
      APOrder.State = osModified
    End If
  End If
  
  gBB.Comm.Message cwiOE_NEW, "RECEIEVED ORDER", NCID_BROADCAST_MSG, NCID_OE_MED_APPLICATION
  
  Exit Sub
   
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ChildInterface.UpdateOrder", "A/P", vbExclamation
  
End Sub

Private Function CovertItemTypeToOrderType(ByVal ItemType As enumItemType) As enumOrderType
On Error GoTo ErrHandler

  Select Case ItemType
    Case itProcedure
      CovertItemTypeToOrderType = typProc
      
    Case itConsult
      CovertItemTypeToOrderType = typConsult
      
    Case itRad
      CovertItemTypeToOrderType = typRad
      
    Case itLab
      CovertItemTypeToOrderType = typLab
      
    Case itMed
      CovertItemTypeToOrderType = typMed
      
    Case itPatientInstr
      CovertItemTypeToOrderType = typPI
      
     ' SCR 56524 KMonroe 11/03/04
    Case itInPatient
      CovertItemTypeToOrderType = typInPat

  End Select
  
 Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ChildInterface.CovertItemTypeToOrderType", "A/P", vbExclamation
End Function

'SCR 59845 12/03/04 Frank Lourng
'Get Pregnancy Status from Vitals
Private Function GetPregnancyStatus() As Boolean
    
    Dim Vitals As Object
    Dim VitalPanels As Collection
    Dim i As Integer
    Dim PregnancyValue As Boolean
    
    Const VitalTypePregnant = 108397
    
  On Error GoTo ErrHandler

  PregnancyValue = False
    
  'Init server and get collection of VitalPanel objects
  Set Vitals = gBB.Comm.InitializeOLEServer(NCID_VITAL_SIGNS_APPLICATION)

  'SCR 50994 - Only attempt to get vitals if server is initialized
  If Not Vitals Is Nothing Then Set VitalPanels = Vitals.Vitals
  
  'SCR 50994 - Only look for Pregnancy data if there are vital entries
  If Not VitalPanels Is Nothing Then

      For i = 1 To VitalPanels.Count
        
        If VitalPanels(i).HasValueType(VitalTypePregnant) Then
          PregnancyValue = VitalPanels(i).Values(VitalTypePregnant).Value
          
          If PregnancyValue > 0 Then Exit For
        End If
      Next
  
  End If
  
  GetPregnancyStatus = PregnancyValue
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnPMainView.GetPregnancyStatus", "A/P", vbCritical

End Function
