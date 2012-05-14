VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProxyProblems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Handles all interaction with the problem module.

Option Explicit

Private mProblem As MMMHISProblems.Problem

Private mCachedDiagnosesWrites As Collection    'Keys for Diagnoses that have related
                                                'item changes but have not been saved.
                                                
Private mAllModifiers As Recordset
                                                
Private Sub Class_Initialize()
  Set mCachedDiagnosesWrites = New Collection
End Sub

'Save a Diagnosis' related items to a Problem.
Private Sub AddAllRelatedItemsToProblem(ByVal EncDataId As Long, _
                                        ByRef Problem As OneProblem, _
                                        ByRef Diagnosis As OneDiagnosis)
  
  On Error GoTo ErrHandler
    
  Problem.AddRelatedItem EncDataId, riNCID_RELATED_ENCOUNTERS, True
          
  AddRelatedItemsToProblem Problem, Diagnosis.RelatedConsultTracking, riNCID_RELATED_CONSULT_TRACKING
  AddRelatedItemsToProblem Problem, Diagnosis.RelatedProcedures, riNCID_RELATED_PROCEDURES
  AddRelatedItemsToProblem Problem, Diagnosis.RelatedRads, riNCID_RELATED_ORDERS_RADS
  AddRelatedItemsToProblem Problem, Diagnosis.RelatedLabs, riNCID_RELATED_ORDERS_LABS
  AddRelatedItemsToProblem Problem, Diagnosis.RelatedMedications, riNCID_RELATED_ORDERS_MEDICATIONS
  AddRelatedItemsToProblem Problem, Diagnosis.RelatedPatientInstructions, riNCID_RELATED_PATIENT_INSTRUCTIONS
  AddRelatedItemsToProblem Problem, Diagnosis.RelatedInpatOrders, riNCID_RELATED_ORDERS_OTHER

  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyProblems.AddRelatedItemsToProblem", "A/P", vbExclamation
  
End Sub

Private Sub AddCachedWrite(ByVal DiagnosisDataID As String)
  
  On Error Resume Next
  mCachedDiagnosesWrites.Add DiagnosisDataID, DiagnosisDataID
  
End Sub

Public Function AddNewAssociation(ByVal DiagnosisDataID As String, ByVal OrderDataID As String, _
                                  ByVal OrderType As enumOrderType, _
                                  Optional ByVal DiagnosisSave As Boolean = True) As Boolean
Dim Success As Boolean
Dim ProblemDiagnosis As OneDiagnosis
Dim RelatedItemType As EnumRelatedItem
  
  On Error GoTo ErrHandler
  
  If Val(DiagnosisDataID) = 0 Or Val(OrderDataID) = 0 Then Exit Function
  
  InitProblemModule
      
  Set ProblemDiagnosis = mProblem.ProblemOps.Diagnosis.Item(DiagnosisDataID)
  
  If Not (ProblemDiagnosis Is Nothing) Then
    RelatedItemType = ConvertOrderType(OrderType)
    If Val(OrderDataID) > 0 Then
        Success = ProblemDiagnosis.AddRelatedItem(OrderDataID, RelatedItemType, DiagnosisSave)
    End If
    If Success And DiagnosisSave Then AddCachedWrite DiagnosisDataID
  End If
  
  AddNewAssociation = Success
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyProblems.AddNewAssociation", "A/P", vbExclamation
  
End Function

Public Function AddNewDiagnosis(ByRef Diagnosis As AnPDiagnosis) As Boolean
Dim ProblemDiagnosis As OneDiagnosis
Dim Chronicity As Long
Dim FacilityNCID As String
Dim lDxNCID As Long

  On Error GoTo ErrHandler
  
  InitProblemModule
  
  If Diagnosis.Chronicity = "Acute" Then
    Chronicity = CLng(NCID_ChronicityAcute)
  Else
    Chronicity = CLng(NCID_ChronicityChronic)
  End If
    
  FacilityNCID = gBB.Logon.FacilityNCID
    
  With Diagnosis
      'JT BEGIN SCR 40955 - Convert the medcin id for PGUI mode
      
      'JT SCR 44133 - In PGUI mode if the diagnosis is from the problem list the NCID will be the ICD9 code
      'problems doesn't use the NCID in this case so if it's not numeric just pass in a -1 so it will ignore it
'OA
'      If IsNumeric(.NCID) Then
'         lDxNCID = CLng(.NCID)
'      Else
'         lDxNCID = -1
'      End If
      If IsNumeric(.SnoID) Then
         lDxNCID = CLng(.SnoID)
      Else
         lDxNCID = -1
      End If
      'JT SCR 44133
      
      Set ProblemDiagnosis = mProblem.ProblemOps.Diagnosis.AddNew(lDxNCID, _
                            .SnoID, .Comment, .Prefix, Chronicity, _
                            CLng(NCID_ACTIVE), CLng(NCID_SourceEncounter), _
                            .DiagnosisType, Now, FacilityNCID, , .EncounterID, _
                            FacilityNCID, .FacilityNCID, , , _
                            .Priority, .ExtenderCode)
                        
    If Not (ProblemDiagnosis Is Nothing) Then
      .DataID = ProblemDiagnosis.DataID
    End If
  End With
    
  AddNewDiagnosis = Len(Diagnosis.DataID) > 0
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyProblems.AddNewDiagnosis", "A/P", vbExclamation
  Exit Function
  Resume
End Function

Public Function AddNewProcedure(ByVal CPT As String) As String
Dim Ops As ProblemOps
Dim ProblemProc As OneProcedure
Dim FacilityID As String
Dim FacilityNCID As String
Dim EncounterID As String

  On Error GoTo ErrHandler
  
  InitProblemModule
  
  Set Ops = mProblem.ProblemOps

  FacilityNCID = gBB.Logon.FacilityNCID
  FacilityID = gBB.CurrentEncounter.FacilityNCID
  EncounterID = gBB.CurrentEncounter.EncounterID
  
  Set ProblemProc = Ops.Procedures.AddNew(CPT, vbNullString, _
                                          NCID_ACTIVE, , , FacilityID, _
                                          , , EncounterID, FacilityNCID)
        
  If Not (ProblemProc Is Nothing) Then
    AddNewProcedure = ProblemProc.DataID
  End If
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyProblems.AddNewProcedure", "A/P", vbExclamation
  
End Function

Private Sub AddRelatedItemsToProblem(ByRef Problem As OneProblem, _
                                      ByRef RelatedItems As Collection, _
                                      ByVal ItemType As EnumRelatedItem)
Dim i As Long

  On Error GoTo ErrHandler

  For i = 1 To RelatedItems.Count
    Problem.AddRelatedItem RelatedItems.Item(i), ItemType, True
  Next
  
  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyProblems.AddRelatedItemsToProblem", "A/P", vbExclamation
  
End Sub

Public Sub ConvertDiagnosesToProblems()
Dim Enc As CHCSIIEncounterOps.Encounter
Dim EncounterID As String
Dim EncounterDataID As Long
Dim Diagnoses As IDiagnosis
Dim Diagnosis As OneDiagnosis

  On Error GoTo ErrHandler
   
  InitProblemModule
  
  EncounterID = gBB.CurrentEncounter.EncounterID
  
  Set Enc = gBB.CurrentEncounter
  EncounterDataID = Enc.DataID

  Set Diagnoses = mProblem.ProblemOps.Diagnosis.Filter(EncounterID, vbNullString)
  
  For Each Diagnosis In Diagnoses
    SaveDiagnosisAsProblem EncounterDataID, Diagnosis
  Next
  
  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyProblems.ConvertDiagnosesToProblems", "A/P", vbExclamation

End Sub

Private Function ConvertOrderType(ByVal OrderType As enumOrderType) As EnumRelatedItem
  On Error GoTo ErrHandler
   
  Select Case OrderType
    Case typProc
      ConvertOrderType = riNCID_RELATED_PROCEDURES
      
    Case typConsult
      ConvertOrderType = riNCID_RELATED_CONSULT_TRACKING
      
    Case typRad
      ConvertOrderType = riNCID_RELATED_ORDERS_RADS
      
    Case typLab
      ConvertOrderType = riNCID_RELATED_ORDERS_LABS
      
    Case typMed
      ConvertOrderType = riNCID_RELATED_ORDERS_MEDICATIONS
      
    Case typPI
      ConvertOrderType = riNCID_RELATED_PATIENT_INSTRUCTIONS
      
    Case typInPat
      ConvertOrderType = riNCID_RELATED_ORDERS_OTHER
  End Select

  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyProblems.ConvertOrderType", "A/P", vbExclamation
End Function

Public Function DeleteAssociation(ByVal DiagnosisDataID As String, ByVal OrderDataID As String, _
                                ByVal OrderType As enumOrderType, _
                                Optional ByVal DiagnosisSave As Boolean = True) As Boolean
Dim Success As Boolean
Dim ProblemDiagnosis As OneDiagnosis
Dim RelatedItemType As EnumRelatedItem
  
  On Error GoTo ErrHandler
  
  If Len(DiagnosisDataID) = 0 Or Len(OrderDataID) = 0 Then Exit Function
  
  InitProblemModule
      
  Set ProblemDiagnosis = mProblem.ProblemOps.Diagnosis.Item(DiagnosisDataID)
  
  If Not (ProblemDiagnosis Is Nothing) Then
    RelatedItemType = ConvertOrderType(OrderType)
    Success = ProblemDiagnosis.DeleteRelatedItem(OrderDataID, RelatedItemType, DiagnosisSave)
  
    If Success And DiagnosisSave Then AddCachedWrite DiagnosisDataID
  End If
  
  DeleteAssociation = Success
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyProblems.DeleteAssociation", "A/P", vbExclamation

End Function

Public Function DeleteDiagnosis(ByVal DataID As String) As Boolean
Dim ProblemDiagnosis As OneDiagnosis

  On Error GoTo ErrHandler
  
  InitProblemModule
      
  Set ProblemDiagnosis = mProblem.ProblemOps.Diagnosis.Item(DataID)
      
  If ProblemDiagnosis Is Nothing Then Exit Function
  
  With ProblemDiagnosis
    .StatusNCID = NCID_DELETED_ERROR
    DeleteDiagnosis = .Update
  End With
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyProblems.DeleteDiagnosis", "A/P", vbExclamation

End Function

Public Function DeleteProcedure(ByVal DataID As String) As Boolean
Dim Ops As ProblemOps
Dim ProblemProc As OneProcedure
Dim Success As Boolean

  On Error GoTo ErrHandler

  InitProblemModule
  
  Set Ops = mProblem.ProblemOps
  Set ProblemProc = Ops.Procedures.Item(DataID)
      
  If Not (ProblemProc Is Nothing) Then
    With ProblemProc
      .StatusNCID = NCID_DELETED_ERROR
      Success = .Update
    End With
  Else
    Success = False
  End If

  DeleteProcedure = Success

  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyProblems.DeleteProcedure", "A/P", vbExclamation

End Function

Private Sub InitProblemModule()
  
  On Error GoTo ErrHandler
  
  If mProblem Is Nothing Then
    Set mProblem = gBB.ReadModule(NCID_PROBLEMS_APPLICATION)
  End If

  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyProblems.InitProblemModule", "A/P", vbExclamation

End Sub

Private Function IsValidToAddProblem(ByRef Problem As OneProblem, ByRef Diagnosis As OneDiagnosis) As Boolean

  On Error GoTo ErrHandler

  'Do not add Deleted Diagnosis if they are not pre-existing, hopefully we do not get here, but...
  If Problem Is Nothing Then
    If Diagnosis.StatusNCID = NCID_DELETED_ERROR Then
      IsValidToAddProblem = False
    Else
      IsValidToAddProblem = True
    End If
  Else
    'If we found a pre-exisiting item created by the GUI then
    'If the Diagnosis is deleted, do nothing
    If Diagnosis.StatusNCID = NCID_DELETED_ERROR Then
      If Problem.Encounters.Count = 0 Then              'Nest check for preformance (most likely it won't get this far)
        IsValidToAddProblem = False
      Else
        IsValidToAddProblem = True
      End If
    Else
      IsValidToAddProblem = True
    End If
  End If
    
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyProblems.IsValidToAddProblem", "A/P", vbExclamation
    
End Function

Private Sub LoadAssociations(ByRef Diagnosis As AnPDiagnosis, _
                              ByRef RelatedItems As Collection, ByVal OrderType As enumOrderType)
Dim Association As OrderAssociation
Dim i As Long
Dim OrderDataID As String

  On Error GoTo ErrHandler
   
  If RelatedItems Is Nothing Then Exit Sub
  
  If RelatedItems.Count = 0 Then Exit Sub
  
  For i = 1 To RelatedItems.Count
    OrderDataID = CStr(RelatedItems.Item(i))
    
    If Len(OrderDataID) > 0 Then
      Set Association = New OrderAssociation
      
      Association.Init OrderDataID, OrderType
          
      Diagnosis.AssociateOrder Association
      
      Set Association = Nothing
    End If
  Next
  
  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyProblems.LoadAssociations", "A/P", vbExclamation

End Sub

Private Sub LoadAllAssociations(ByRef Diagnosis As AnPDiagnosis, ByRef ProblemDiagnosis As OneDiagnosis)
  On Error GoTo ErrHandler
  
  LoadAssociations Diagnosis, ProblemDiagnosis.RelatedProcedures, typProc
  LoadAssociations Diagnosis, ProblemDiagnosis.RelatedMedications, typMed
  LoadAssociations Diagnosis, ProblemDiagnosis.RelatedLabs, typLab
  LoadAssociations Diagnosis, ProblemDiagnosis.RelatedRads, typRad
  LoadAssociations Diagnosis, ProblemDiagnosis.RelatedPatientInstructions, typPI
  LoadAssociations Diagnosis, ProblemDiagnosis.RelatedConsultTracking, typConsult
  LoadAssociations Diagnosis, ProblemDiagnosis.RelatedInpatOrders, typInPat

  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyProblems.LoadAllAssociations", "A/P", vbExclamation
End Sub

Private Function LoadDiagnosis(ByRef ProblemDiagnosis As OneDiagnosis, ByVal lNextPriority As Long) As AnPDiagnosis
Dim Diagnosis As AnPDiagnosis
 
  On Error GoTo ErrHandler
   
  Set Diagnosis = New AnPDiagnosis

  With Diagnosis
    .Init ProblemDiagnosis.DataID, ProblemDiagnosis.Diagnosis
    
    'JT BEGIN SCR 41004 - convert priority if in PGUI mode and change state to force save
    If IsNumeric(ProblemDiagnosis.Priority) Then
      .Priority = ProblemDiagnosis.Priority
    Else
      .Priority = lNextPriority
      .State = dsModified
    End If
    'JT END SCR 41004
    
    .Chronicity = ProblemDiagnosis.Chronicity
    .DiagnosisType = ProblemDiagnosis.NewFollowUp
    .ShortName = ProblemDiagnosis.ICD9ShortName
    .LongName = ProblemDiagnosis.ICD9LongName
    .ICD9Code = ProblemDiagnosis.ICD9Code
    .NCID = ProblemDiagnosis.ICD9NCID
    .Comment = ProblemDiagnosis.Comment
    If IsDate(ProblemDiagnosis.OnsetDate) Then .OnsetDate = ProblemDiagnosis.OnsetDate
    .ExtenderCode = ProblemDiagnosis.ExtendedICD9
    .Prefix = ProblemDiagnosis.Prefix           'SCR 11825  Phil Crowder 11/28/01
    .SnoID = ProblemDiagnosis.SnoID
    .EncounterID = ProblemDiagnosis.EncounterNumber
    .FacilityNCID = ProblemDiagnosis.FacilityNCID
    .ClinicNCID = ProblemDiagnosis.ClinicianNCID
    .FromProblemList = ProblemDiagnosis.FromProblemList     'JT SCR 40944 - indicate if from problem list for PGUI
    SetExtenderValues Diagnosis
  End With
  
  Set LoadDiagnosis = Diagnosis
  
  Exit Function
  
ErrHandler:
  Set LoadDiagnosis = Diagnosis
  gShared.ShowVBError Err.Number, Err.Description, "ProxyProblems.LoadDiagnosis", "A/P", vbExclamation

End Function

Public Function LoadDiagnoses(ByVal EncounterID As String) As Collection
Dim Diagnoses As Collection
Dim ProblemDiagnoses As IDiagnosis
Dim ProblemDiagnosis As OneDiagnosis
Dim Diagnosis As AnPDiagnosis
Dim bAddDiagnosis As Boolean

  On Error GoTo ErrHandler
   
  Set Diagnoses = New Collection
  
  InitProblemModule

  Set ProblemDiagnoses = mProblem.ProblemOps.Diagnosis.Filter(EncounterID, NCID_ACTIVE)
  
   For Each ProblemDiagnosis In ProblemDiagnoses
          
      'JT BEGIN SCR 41004 - Determine if it is ok to add the diagnosis
      If IsNumeric(ProblemDiagnosis.Priority) Then
         bAddDiagnosis = True
      ElseIf gShared.IsAppMode(modeCHCSI_GUI) Then
         'Its ok to add but the priority will be reset when loaded
         bAddDiagnosis = True
      Else
         bAddDiagnosis = False
      End If
      
      If bAddDiagnosis Then
         
        Set Diagnosis = LoadDiagnosis(ProblemDiagnosis, Diagnoses.Count + 1)
        
        LoadAllAssociations Diagnosis, ProblemDiagnosis
        
        Diagnoses.Add Diagnosis, Diagnosis.Key
      
      End If
      'JT END SCR 41004
      
  Next
    
  Set LoadDiagnoses = Diagnoses
  
  Exit Function
  
ErrHandler:
  Set LoadDiagnoses = Diagnoses
  gShared.ShowVBError Err.Number, Err.Description, "ProxyProblems.LoadDiagnoses", "A/P", vbExclamation

End Function

Private Sub LoadModifiers(ByRef Procedure As OrderProcedure, ByRef Modifiers As WAMData.Modifiers)
Dim i As Integer
Dim Description As String

  On Error GoTo ErrHandler

  If Modifiers.Count = 0 Then Exit Sub
  
  If mAllModifiers Is Nothing Then
    Set mAllModifiers = gBB.CDRPlusInterface.ListOfAllModifiers
  End If
  
  For i = 1 To Modifiers.Count
    mAllModifiers.MoveFirst
    mAllModifiers.Find "Modifier_Code = '" & Modifiers.Item(i).Code & "'"
    
    If Not mAllModifiers.EOF Then
      Description = mAllModifiers!Short_Description
    Else
      Description = "Unknown Description"
    End If
    
    Procedure.Modifiers.Add Modifiers.Item(i).Code, Description, True
  Next

  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyProblems.LoadModifiers", "A/P", vbExclamation

End Sub

Public Function LoadProcedures(ByVal EncounterID As String) As Collection
Dim ProblemProcs As IProcedures
Dim ProblemProc As OneProcedure
Dim AnPOrders As Collection
Dim AnPOrder As iOrder

  On Error GoTo ErrHandler

  InitProblemModule
  
  Set AnPOrders = New Collection
  
  Set ProblemProcs = mProblem.ProblemOps.Procedures.Filter(EncounterID, NCID_ACTIVE)
    
  For Each ProblemProc In ProblemProcs
    Set AnPOrder = LoadProcedure(ProblemProc)
    AnPOrders.Add AnPOrder, AnPOrder.Key
  
    Set AnPOrder = Nothing
  Next
  
  Set LoadProcedures = AnPOrders
  
  Exit Function
  
ErrHandler:
  Set LoadProcedures = AnPOrders
  gShared.ShowVBError Err.Number, Err.Description, "ProxyProblems.LoadProcedures", "A/P", vbExclamation
  
End Function

Private Function LoadProcedure(ByRef ProblemProc As OneProcedure) As iOrder
Dim AnPOrder As iOrder
Dim AnPProc As OrderProcedure

  On Error GoTo ErrHandler
    
  Set AnPOrder = New OrderProcedure
  
  AnPOrder.Init ProblemProc.DataID, ProblemProc.Procedure
  
  Set AnPProc = AnPOrder
  
  With AnPProc
    .CPTID = ProblemProc.CPTID
    .NCID = ProblemProc.ProcedureNCID
    .CPTCode = ProblemProc.CPT4Code
    .Comment = ProblemProc.Comment
    If Val(ProblemProc.UnitsOfService) > 0 Then .Modifiers.UnitsOfService = ProblemProc.UnitsOfService
  End With

  LoadModifiers AnPProc, ProblemProc.Modifiers

  Set LoadProcedure = AnPOrder

  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyProblems.LoadProcedure", "A/P", vbExclamation
  Resume Next
  
End Function

'Save all diagnoses that have updated related items but have not been saved.
'Problem module can cache related item changes to a diagnsis and save them at one time.
'The related item changes are cached in AddNewAssociation and DeleteAssociation methods.
Public Function SaveCachedWrites() As Boolean
Dim i As Long
Dim ProblemDiagnosis As OneDiagnosis

  On Error GoTo ErrHandler

  If mCachedDiagnosesWrites.Count = 0 Then
    SaveCachedWrites = True
    Exit Function
  End If
  
  For i = mCachedDiagnosesWrites.Count To 1 Step -1
    Set ProblemDiagnosis = mProblem.ProblemOps.Diagnosis.Item(mCachedDiagnosesWrites.Item(i))
    
    If Not (ProblemDiagnosis Is Nothing) Then
      SaveCachedWrites = ProblemDiagnosis.Update
      
      If Not SaveCachedWrites Then Exit For
    End If
    
    If SaveCachedWrites Then
      mCachedDiagnosesWrites.Remove i
    End If
  Next
  
  Exit Function
  
ErrHandler:
  
End Function

Private Sub SaveDiagnosisAsProblem(ByVal EncounterDataID As Long, _
                                    ByRef Diagnosis As OneDiagnosis)
    
Dim Problem As OneProblem
Dim OnsetDate As String
Dim NCIDStatus As Long
Dim NCIDChronicity As Long
Dim UpdatedProblem As Boolean
Dim Success As Boolean
Dim ValidSave As Boolean

  On Error GoTo ErrHandler

  Set Problem = mProblem.ProblemOps.Problems.FindSnoID(Diagnosis.SnoID)
    
  ValidSave = IsValidToAddProblem(Problem, Diagnosis)
  
  If Not ValidSave Then Exit Sub
  
  'Do not add Deleted Diagnosis if they are not pre-existing, hopefully we do not get here, but...
  If Not Problem Is Nothing Then
    'If we found a pre-exisiting item and we are adding, then cache some of it's data
    OnsetDate = Problem.OnsetDate
    NCIDStatus = Problem.StatusNCID
    NCIDChronicity = Problem.ChronicityNCID
  End If
    
  'Do the Add, pre-exisiting Problems are automatically updated, new items are not created.
  'bolUpdate is set to True, if Problem already Exists (Update Occurs)
  With Diagnosis
   'JT BEGIN SCR 40955 - convert the medcin id for PGUI mode
    Set Problem = mProblem.ProblemOps.Problems.AddNew(.DiagnosisNCID, .SnoID, .Prefix, _
      .Comment, .StatusNCID, .ChronicityNCID, .SourceNCID, .OnsetDate, _
      .FacilityNCID, .NotedDate, .PointOfCareFacilityNCID, UpdatedProblem, True)
  End With
    
  If Problem Is Nothing Then
    Err.Raise vbObjectError, Err.Source, "Error creating problem from diagnosis: " & Diagnosis.Diagnosis
  End If
    
  'Additional Fields not covered in ADDNEW method
  Problem.ClinicianNCID = Diagnosis.ClinicianNCID
  Problem.EncounterFacilityNCID = Diagnosis.EncounterFacilityNCID
        
  'Add or Remove Related Items
  If Diagnosis.StatusNCID = NCID_ACTIVE Then
    AddAllRelatedItemsToProblem EncounterDataID, Problem, Diagnosis  'Add Related Items including the Encounter
  Else
    Problem.DeleteRelatedEncounter EncounterDataID, True                'Remove the Encounter
  End If
    
  'COMPLEX RULES WHEN UPDATING
  'Status Rules, only worry if prexisiting Problem...
  If UpdatedProblem Then
    '1) If UPDATE, set 'Onset Date' to Original 'Onset Date'
    Problem.OnsetDate = OnsetDate
    
    '2) If Diagnosis is Deleted and Problem IS Related to other encounter(s), reset Status to original
    If Diagnosis.StatusNCID = NCID_DELETED_ERROR Then
      If Problem.Encounters.Count > 0 Then              'Put here for preformance reasons
        Problem.StatusNCID = NCIDStatus
      End If
    End If
        
    '3)If Problem was Pre-exisiting and Deleted, preserve chronicity
    'because Chronicy could have been set by the GUI, not only during Signing
    If Diagnosis.StatusNCID = NCID_DELETED_ERROR Then
      Problem.ChronicityNCID = NCIDChronicity
    End If
  End If
    
  'Save Any/All The Changes
  Success = Problem.Update    'MUST CALL in order to perform any saves, all other writes using "No Save" option
    
  If Not Success Then
    Err.Raise vbObjectError, Err.Source, "Update error during creation of problem from diagnosis: " & Diagnosis.Diagnosis
  End If
    
  Exit Sub
    
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyProblems.SaveDiagnosisAsProblem", "A/P", vbExclamation
  
End Sub

Private Sub SetExtenderValues(ByRef Diagnosis As AnPDiagnosis)
Dim MedcinForm As MedcinDialog

  On Error Resume Next
  
  If Len(Diagnosis.ExtenderCode) < 3 Then Exit Sub
  
  Set MedcinForm = New MedcinDialog
  
  MedcinForm.SetExtenderValues Diagnosis
  
  Unload MedcinForm
  Set MedcinForm = Nothing
    
End Sub

Public Function UpdateDiagnosis(ByRef Diagnosis As AnPDiagnosis) As Boolean
Dim ProblemDiagnosis As OneDiagnosis
    
  On Error GoTo ErrHandler
  
  InitProblemModule
    
  Set ProblemDiagnosis = mProblem.ProblemOps.Diagnosis.Item(Diagnosis.DataID)
    
  If ProblemDiagnosis Is Nothing Then Exit Function
  
  With ProblemDiagnosis
    .Priority = Diagnosis.Priority
    .NewFollowUp = Diagnosis.DiagnosisType
    .Comment = Diagnosis.Comment
    
    If Diagnosis.Chronicity = "Acute" Then
      .ChronicityNCID = NCID_ChronicityAcute
    Else
      .ChronicityNCID = NCID_ChronicityChronic
    End If
           
    UpdateDiagnosis = .Update
  End With

  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyProblems.UpdateDiagnosis", "A/P", vbExclamation

End Function

Public Function UpdateProcedure(ByRef Procedure As OrderProcedure) As Boolean
Dim ProblemProcedure As OneProcedure
Dim Order As iOrder
Dim Codes As Collection
Dim i As Long

  On Error GoTo ErrHandler
  
  InitProblemModule
    
  Set Order = Procedure
  Set ProblemProcedure = mProblem.ProblemOps.Procedures.Item(Order.DataID)
    
  If ProblemProcedure Is Nothing Then Exit Function

  With ProblemProcedure
    .Comment = Procedure.Comment
    .UnitsOfService = Procedure.Modifiers.UnitsOfService
    
    .Modifiers.Clear
    
    Set Codes = Procedure.Modifiers.Codes
    For i = 1 To Codes.Count
      .Modifiers.Add Codes.Item(i)
    Next
    
    UpdateProcedure = .Update
  End With
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyProblems.UpdateProcedure", "A/P", vbExclamation

End Function

Public Sub RefreshProblems()
 
   On Error GoTo ErrHandler
  
   InitProblemModule
   mProblem.ProblemOps.Refresh
   
   Exit Sub
   
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyProblems.RefreshProblems", "A/P", vbExclamation
 
End Sub
