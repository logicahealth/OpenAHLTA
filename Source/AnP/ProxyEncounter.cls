VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProxyEncounter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const NOTE_TITLE As String = "Additional A/P Information:"
Private Const WARNING_TITLE As String = "                                                                        WARNING"
Private Const WARNING_MSG As String = "A user has taken over ownership of the A/P section of this encounter.  " & _
                                      "Any orders that you have placed have been submitted.  Your A/P documentation " & _
                                      "will be added as free text  to the end of the encounter as an addendum.  " & _
                                      "Please review the A/P section in the current encounter for correctness.  " & _
                                      "You may need to reenter your diagnoses and procedures."
Private Const WARNING_MSG_HIPAA As String = "A user has taken over ownership of the A/P section or the Header section of this encounter.  " & _
                                      "Any orders that you have placed have been submitted.  Your A/P documentation" & _
                                      ", excluding any additional provider and/or procedure associations, " & _
                                      "will be added as free text to the end of the encounter as an addendum.  " & _
                                      "Please review the A/P section in the current encounter for correctness.  " & _
                                      "You may need to reenter your diagnoses, procedures, additional providers, and procedure associations."
              
Private mCurrentEncounter As CHCSIIEncounterOps.Encounter

Private mNoteAdded As Boolean              'SCR 22949, 22899 Phil Crowder 4/11/02  Used to make sure only one note is added

Public Sub AddEncounterNote(ByVal Note As String)
  
  On Error GoTo ErrHandler
  
  InitEncounter
  
  mCurrentEncounter.AddTextNote NOTE_TITLE, Note
  
  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyEncounter.AddEncounterNote", "A/P", vbExclamation
  
End Sub

'Begin SCR 49601 - added helper function to add list tool record to the SnoIDs collection
Private Sub AddSnoID(ByRef cSnoIDs As Collection, ByVal sSnoID As String, ByVal sPrefix As String)

Dim ListToolRec As ListToolRecord
   
   Set ListToolRec = New ListToolRecord
   ListToolRec.SnoID = sSnoID
   If sPrefix <> vbNullString Then ListToolRec.Prefix = sPrefix
   cSnoIDs.Add ListToolRec
   
End Sub
'End SCR 49601

'Begin SCR 49601 - Updated to pass a collection of listtoolrecords so that encounters will have both SnoID and prefix
Private Sub AddSnoIDs(ByRef Diagnoses As Collection)

Dim SnoIDs As Collection
Dim Diagnosis As AnPDiagnosis
Dim PROC As OrderProcedure
Dim PI As OrderPatientInstruc
Dim Orders As Collection
Dim i As Long

  On Error GoTo ErrHandler

  Set SnoIDs = New Collection
  
  For Each Diagnosis In Diagnoses
    AddSnoID SnoIDs, Diagnosis.SnoID, Diagnosis.Prefix
  Next
  
  Set Orders = gBB.AllOrders.Filter(typProc, fltSaved)
  For i = 1 To Orders.Count
    Set PROC = Orders.Item(i)
    AddSnoID SnoIDs, PROC.CPTID, vbNullString
  Next
  
  Set Orders = gBB.AllOrders.Filter(typPI, fltSaved)
  For i = 1 To Orders.Count
    Set PI = Orders.Item(i)
    AddSnoID SnoIDs, PI.SnoID, vbNullString
  Next
  
  Set gBB.EncParent.SectionObject.SnoIDs = SnoIDs

  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyEncounter.AddSnoIDs", "A/P", vbExclamation

End Sub
'End SCR 49601

Private Sub ClearPrimaryDiagnosis()

  On Error GoTo ErrHandler
  
  With mCurrentEncounter
    .PrimaryDiagnosisNCID = ""
    .PrimaryDxSnoID = ""
    .PrimaryDxMedcinPrefix = vbNullString
  End With

  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyEncounter.ClearPrimaryDiagnosis", "A/P", vbExclamation

End Sub

Public Function CreateRTF(Optional ByVal IsPreview As Boolean = False) As String

  On Error GoTo ErrHandler
    
  CreateRTF = RTF_HEADER & gBB.AllDiagnoses.RTF(Not IsPreview) & _
              gBB.AllOrders.UnassociatedOrdersRTF(Not IsPreview) & RTF_FOOTER
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyEncounter.CreateRTF", "A/P", vbExclamation
  
End Function

Private Function EndEncounterSave() As Boolean
Dim RTF As String
Dim Diagnoses As Collection

  On Error GoTo ErrHandler
  
  RTF = CreateRTF(False)
  If Len(RTF) > 0 Then gBB.EncParent.SetRTF RTF
                
  Set Diagnoses = gBB.AllDiagnoses.Filter(fltActive)
                
  If Diagnoses.Count > 0 Then
    Set mCurrentEncounter.RelatedDiagnoses = Diagnoses
    
    AddSnoIDs Diagnoses
  
    SetPrimaryDiagnosis Diagnoses
  Else
    ClearPrimaryDiagnosis
  End If
  

  'Code Review Note Soano 3/18/04: to reduce complexity consider if this shoulb be broken out to a function,
  ' using "short circuited logic"; see ProAnalyzer: Consider short-circuited logic
  'Force a save of encounter data if accident or provider info has changed
  If IsHIPAAEnabled And gShared.IsAppMode(modeCDR) And Not gShared.IsAppMode(modeITT) And (mCurrentEncounter.AssociatedProviders.IsDirty Or mCurrentEncounter.AccidentData.IsDirty) Then
      mCurrentEncounter.Save
  End If
  
  gBB.EncParent.EndSave

  EndEncounterSave = True
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyEncounter.EndEcounterSave", "A/P", vbExclamation
  EndEncounterSave = False
  'Try to end the encounter save
  gBB.EncParent.EndSave
  
End Function

Private Sub InitEncounter()

  If mCurrentEncounter Is Nothing Then
    Set mCurrentEncounter = gBB.CurrentEncounter
  End If
  
End Sub

'SCR 32576  2/27/03  Phil Crowder
Private Function IsAllOrdersAssociated() As Boolean

  IsAllOrdersAssociated = IsAllOrderTypeAssociated(typConsult)
  
  If IsAllOrdersAssociated Then IsAllOrdersAssociated = IsAllOrderTypeAssociated(typLab)

  If IsAllOrdersAssociated Then IsAllOrdersAssociated = IsAllOrderTypeAssociated(typMed)

  If IsAllOrdersAssociated Then IsAllOrdersAssociated = IsAllOrderTypeAssociated(typRad)

End Function

Private Function IsAllProceduresAssociated() As Boolean

   IsAllProceduresAssociated = IsAllOrderTypeAssociated(typProc)
  
End Function

'Make sure all orders are associated
Private Function IsAllOrderTypeAssociated(ByVal OrderType As enumOrderType) As Boolean
Dim Associated As Boolean
Dim Orders As Collection
Dim Order As iOrder
Dim PROC As OrderProcedure

  On Error GoTo ErrHandler
  
  Set Orders = gBB.AllOrders.Filter(OrderType, fltSaved)
  
  If Orders.Count = 0 Then Associated = True
  
  For Each Order In Orders
    Associated = Order.IsAssociated(True)     'Function accounts for associations that failed to saved
      
    If OrderType = typProc And IsHIPAAEnabled Then
      'Check for provider associations
      Set PROC = Order
      Associated = Associated And Len(PROC.Providers) > 0
    End If
      
    If Not Associated Then Exit For
  Next

  IsAllOrderTypeAssociated = Associated
    
  Exit Function
  
ErrHandler:
  IsAllOrderTypeAssociated = False
  gShared.ShowVBError Err.Number, Err.Description, "ProxyEncounter.IsAllOrderTypeAssociated", "A/P", vbExclamation

End Function

'SCR 34746 4/16/03  Phil Crowder
Private Function IsPrimaryDiagnosisValid() As Boolean
Dim Diagnoses As Collection
  
  On Error GoTo ErrHandler
  
  If Len(gShared.CmdLineSwitches("WAM")) = 0 And Not gShared.IsAppMode(modeCHCSI_GUI) Then
    IsPrimaryDiagnosisValid = True
    Exit Function
  End If
  
  Set Diagnoses = gBB.AllDiagnoses.Filter(fltActive)
  
  If Diagnoses.Count = 0 Then
    IsPrimaryDiagnosisValid = True
  Else
    IsPrimaryDiagnosisValid = UCase$(Left$(Diagnoses.Item("1").FullICDCode, 1)) <> "E"
  End If

  Exit Function
  
ErrHandler:
  IsPrimaryDiagnosisValid = False
  gShared.ShowVBError Err.Number, Err.Description, "ProxyEncounter.IsPrimaryDiagnosisValid", "A/P", vbExclamation

End Function

Private Function IsProviderDirty() As Boolean

On Error GoTo ErrHandler

Dim oEnc As CHCSIIEncounterOps.Encounter
Dim oProv As EncProvider

   Set oEnc = gBB.CurrentEncounter
   For Each oProv In oEnc.Providers
      If oProv.IsDirty() Then
         IsProviderDirty = True
         Exit For
      End If
   Next
   
ErrHandler:
   Set oEnc = Nothing
   Set oProv = Nothing
   
   If CBool(Err) Then gShared.ShowVBError Err.Number, Err.Description, "ProxyEncounter.IsProviderDirty", "A/P", vbExclamation
  
End Function

Private Function IsSaveRequired() As Boolean
Dim DiagnosesDirty As Boolean
Dim OrdersDirty As Boolean
Dim ProvidersDirty As Boolean
Dim AccidentDirty As Boolean

  On Error GoTo ErrHandler

  'Check to see if save needed.
  DiagnosesDirty = (gBB.AllDiagnoses.Filter(fltDirty).Count > 0)
  OrdersDirty = (gBB.AllOrders.Filter(, fltDirty).Count > 0)
  ProvidersDirty = IsProviderDirty
  If Len(gShared.CmdLineSwitches("HIPAA837")) > 0 Then AccidentDirty = mCurrentEncounter.AccidentData.IsDirty
  
  'JT SCR 38256 - Don't need to force the save in PGUI mode now that layer two is complete and we save data in a DB
  IsSaveRequired = DiagnosesDirty Or OrdersDirty Or ProvidersDirty Or AccidentDirty
  
  Exit Function
  
ErrHandler:
  IsSaveRequired = True
  gShared.ShowVBError Err.Number, Err.Description, "ProxyEncounter.IsSaveRequired", "A/P", vbExclamation
  
End Function

Public Function Save() As Boolean
Dim Success As Boolean
Dim StatusDialog As StatusDialog
Dim SaveNeeded As Boolean
Dim StateChanged As Boolean
Dim bRefreshAutoCites As Boolean

  On Error GoTo ErrHandler
        
  InitEncounter
              
  SaveNeeded = IsSaveRequired
  
  If Not SaveNeeded Then
    StateChanged = SetEncounterState
  End If
  
  If Not SaveNeeded And Not StateChanged Then
    Save = True
    Exit Function
  End If
  
  Set StatusDialog = New StatusDialog
  StatusDialog.Display "Saving", "Please Wait While Saving", True, 10
  
  Success = StartEncounterSave
      
  If Success Then
    StatusDialog.UpdateProgress 20
      
    SyncProviderProcedures
    
    Success = gBB.AllOrders.Save        'Save Orders
        
    StatusDialog.UpdateProgress 50
        
    Success = gBB.AllDiagnoses.Save               'Save diagnoses and related items

    StatusDialog.UpdateProgress 80
    
    If gShared.IsAppMode(modeCHCSI_GUI) And (mCurrentEncounter.AccidentData.IsDirty Or mCurrentEncounter.AssociatedProviders.IsDirty) Then
      'SCR 51723 - In PGUI the refresh of the autocites rtf below is causing a refresh before the provider data is saved in encounter - save all HIPAA data here
       bRefreshAutoCites = mCurrentEncounter.AccidentData.IsDirty       'Need to refresh autocite rtf when accident data has changed
       mCurrentEncounter.SaveHIPAAData False
    End If
    
    If Not StateChanged Then StateChanged = SetEncounterState
    
    Success = EndEncounterSave
    
    If bRefreshAutoCites Then mCurrentEncounter.RefreshAutocites
    
  End If
  
  Unload StatusDialog
  Set StatusDialog = Nothing
  DoEvents
  
  Save = True
  
  Exit Function
   
ErrHandler:
  'SCR 26383 8/7/02  Phil Crowder
  'Ignore modal display error caused by trying to show status form when a modal
  'form is showing such as New Appointment Form.
  If Err.Number <> 401 Then
    Save = False
    gShared.ShowVBError Err.Number, Err.Description, "ProxyEncounter.Save", "A/P", vbExclamation
  Else
    Resume Next
  End If
  
End Function

Private Function SetEncounterState() As Boolean
Dim Signable As Boolean
Dim AnPComplete As Boolean
Dim Associated As Boolean
Dim StatusComment As String
Dim OrdersAssociated As Boolean
Dim ValidPrimaryDiagnosis As Boolean

  On Error GoTo ErrHandler

  StatusComment = vbNullString
  
  'See if encounter complete (ie. All consults are signable and all procedures are associated)
  If gBB.AllDiagnoses.Filter(fltSaved).Count > 0 Then
    'Signable = gBB.ConsultInterface.ValidateAllConsultsSignable
    'If Not Signable Then StatusComment = "DRAFT_CONSULTS;"
    
    Associated = IsAllProceduresAssociated
    If Not Associated Then StatusComment = StatusComment & "UNASSOCIATED_PROCEDURES;"
    
    ValidPrimaryDiagnosis = IsPrimaryDiagnosisValid
    If Not ValidPrimaryDiagnosis Then StatusComment = StatusComment & "INVALID_PRIMARY_DIAGNOSIS;"
    
    If gShared.IsAppMode(modeCHCSI_GUI) Then    'SCR 32576  2/27/03  Phil Crowder
      OrdersAssociated = IsAllOrdersAssociated
      If Not OrdersAssociated Then StatusComment = StatusComment & "UNASSOCIATED_ORDERS;"
    End If
  Else
    StatusComment = "NO_DIAGNOSES;"
  End If
  
  'OA
  AnPComplete = Associated And ValidPrimaryDiagnosis 'Signable And Associated And ValidPrimaryDiagnosis
              
  If AnPComplete And mCurrentEncounter.APStatus <> complete Then
    mCurrentEncounter.APStatus = complete
    
    SetEncounterState = True
  ElseIf Not AnPComplete And mCurrentEncounter.APStatus <> Incomplete Then
    mCurrentEncounter.APStatus = Incomplete
    
    SetEncounterState = True
  End If

  If StatusComment <> mCurrentEncounter.APStatusComment Then
    mCurrentEncounter.APStatusComment = StatusComment
  End If

  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyEncounter.SetEncounterState", "A/P", vbExclamation
    Exit Function
    Resume
End Function

Private Sub SetPrimaryDiagnosis(ByRef Diagnoses As Collection)
Dim Diagnosis As AnPDiagnosis

  On Error GoTo ErrHandler
  
  For Each Diagnosis In Diagnoses
    If Diagnosis.Priority = 1 Then
      With mCurrentEncounter
         'JT SCR 40944 - convert code for PGUI
        .PrimaryDiagnosisNCID = Diagnosis.NCID
        'scr 19371  Phil Crowder  12/5/01
        .PrimaryDxSnoID = Diagnosis.SnoID
        .PrimaryDxMedcinPrefix = Diagnosis.Prefix
      End With
      
      Exit For
    End If
  Next
  
  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyEncounter.SetPrimaryDiagnosis", "A/P", vbExclamation
  Exit Sub
  Resume
End Sub

Private Function StartEncounterSave() As Boolean
Dim Success As Boolean
Dim WarningMessage As String
Dim RTF As String

  On Error GoTo ErrHandler
  
  Success = gBB.EncParent.StartSave()
  
  If Not Success And Not mNoteAdded Then
    'Get Encounter class instance for new lock message for SCR 16521, Phil Crowder 10/5/01
    mCurrentEncounter.Refresh
    gBB.ProblemInterface.RefreshProblems        'SCR 47301 - Lock was broken so refresh data
    
    'SCR 22949, 22899 Phil Crowder 4/11/02  Changed message
    If IsHIPAAEnabled Then
      'SCR 51459 - For HIPAA use an expanded message that includes references to associated provider/procedures data
      WarningMessage = WARNING_TITLE & vbCrLf & WARNING_MSG_HIPAA
    Else
      WarningMessage = WARNING_TITLE & vbCrLf & WARNING_MSG
    End If
    
    If Not (mCurrentEncounter.WhoSigned Is Nothing) Then
      WarningMessage = WarningMessage & "  This encounter has been completed and signed by " & _
                        mCurrentEncounter.WhoSigned.FullName & " on " & _
                        Format$(mCurrentEncounter.WhenSigned, gsDefaultDateTime) & ".  "
    End If
    
    MsgBxARMd WarningMessage, vbExclamation + vbOKOnly, "Encounter"
                
    'SCR 22949, 22899 Phil Crowder 4/11/02  Set the encounter rtf for note
    RTF = CreateRTF(True)
    
    mCurrentEncounter.AddTextNote NOTE_TITLE, RTF
    mNoteAdded = True
    
    'SCR 21227  7/28/02  Phil Crowder  Added per Chris Atencio
    If Not (mCurrentEncounter.WhoSigned Is Nothing) Then
      mCurrentEncounter.ReplaceEncRTF
    End If
  End If
  
  StartEncounterSave = Success
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyEncounter.StartEncounterSave", "A/P", vbExclamation
  
End Function

'JT SCRs 38958, 38959, 38960 - Check encounter option to determine if codes should be displayed in RTF
Public Property Get IsIncludeRTFCodesSelected() As Boolean
  
On Error GoTo ErrHandler
  
   InitEncounter
   
   IsIncludeRTFCodesSelected = CBool(mCurrentEncounter.EncProperties.IncludeCodes = 1)
       
ErrHandler:
   If CBool(Err) Then gShared.ShowVBError Err.Number, Err.Description, "ProxyEncounter.IsIncludeRTFCodesSelected", "A/P", vbExclamation
  
End Property

Friend Sub SyncProviderProcedures()
'Make sure the procedures reflect the proper provider linkage information
On Error GoTo ErrHandler

Dim oOrder As iOrder
Dim oOrders As Collection
Dim oProc As OrderProcedure
Dim oEnc As CHCSIIEncounterOps.Encounter
Dim lIdx As Long

Dim oEncProv As EncProvider

   On Error GoTo ErrHandler

   'Set the Order Data ID in all provider associations of me
   If IsHIPAAEnabled Then
      Set oEnc = gBB.CurrentEncounter
      Set oOrders = gBB.AllOrders.Filter(typProc, fltActive)
      For Each oOrder In oOrders
         Set oProc = oOrder
         oProc.Providers = vbNullString
         For lIdx = 1 To oEnc.Providers.Count
            Set oEncProv = oEnc.Providers(lIdx)
            If oEncProv.HasProcedure(oOrder.DataID, oProc.CPTID) Then
               oProc.Providers = oProc.Providers & CStr(lIdx)
            End If
         Next
      Next
   End If

ErrHandler:
   Set oOrder = Nothing
   Set oOrders = Nothing
   Set oProc = Nothing
   Set oEnc = Nothing

   If CBool(Err) Then gShared.ShowVBError Err.Number, Err.Description, "ProxyEncounter.SyncProviderProcedures", "A/P", vbExclamation

End Sub
