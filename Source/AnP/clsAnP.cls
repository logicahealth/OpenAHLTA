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
Attribute VB_Name = "AnP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'This class is used to provide an interface to A/P for other modules such
'as core, diagnosis, procedure, and patient instructions.  This class
'creates the main view (AnPMainView.frm) on an OpenView/ ResizeView.  All
'public methods are intended to be called by other modules.

Option Explicit

Implements ICHCSII_CoreAPI2
Implements ICHCSII_CoreAPI3
Implements IChild

Private mLock As AnPLock                   'Manages locking of A/P module and the main view

Private mEncOpen As Boolean                'True if an Encounter is open

Private mDataSaved As Boolean              'True if the patient changed or close view message is recieved from Encounter and A/P says it is ok.
                                           'If true then a save has already happened so don't do a save during closing.
                                              
Private WithEvents mAnPMainView As AnPMainView       'Main A/P form
Attribute mAnPMainView.VB_VarHelpID = -1

Private mSwitchingSection As Boolean          'True if user is switching to a different section of the encounter from A/P.
                                              'This var is needed because of 2 reasons.  One is that the only way to tell
                                              'if the app lock screen is showning is to check to see if a modal form is loaded.
                                              'Two is that we use the same message to lock core during long processing.  During
                                              'a close A/P shows a modal dialog if there are unassociated orders when the module.
                                              'This is not shown if the app lock screen is up (cause error).  However, A/P checks
                                              'to see if a modal form is showing and during a switch, encounter locks core with
                                              'the modal message.  The var allows the modal dialog to be shown on a close.
                                              
Private mProcessingOrderSet As Boolean        'SCR 23627  Phil Crowder 8/13/02  Don't allow a Patient Change or Close of Module if Processing an Order Set

Private mMenuModifierEnabled As Boolean

'Menu Ids
Private mMenuTemplates As Long
Private mMenuPreviewSummary As Long
Private mMenuSave As Long
Private mMenuDelete As Long
Private mMenuNext As Long
Private mMenuGoalSummary As Long
Private mMenuBack As Long
Private mMenuSign As Long
Private mMenuModifiers As Long
Private mMenuSubmitAll As Long

'SCR 136632, JQ, 7/26/07; note NCID changes
Private Const NCID_Discharge_Note As String = "15149155" ' "681"

Private Sub Class_Initialize()
  
  ARM_InitAllTransactionsForAnP   '--- MSolano 8/21/01 SCR-ARM-16141
  
  Set mLock = New AnPLock
  Set gShared = New CWShared
  Set gBB = New AnPBlackBoard
  
  App.HelpFile = gShared.CWHelpFile    'Set the help file for this application
  
End Sub

Private Sub Class_Terminate()
  On Error Resume Next
  
  Set gShared = Nothing
  Set gBB = Nothing
  
End Sub

'Interface used by core to Initialize the A/P module.  See Core API for details.
Private Sub ICHCSII_CoreAPI2_Initialize2(ByVal vlMyProgId As Long, ByRef robjComm As Object, ByRef robjLogonEnvironment As Object, ByRef robjConfig As Object, ByRef robjPatient As Object, ByRef robjUser As Object)
  On Error GoTo ErrHandler
    
  cwlMyProgId = vlMyProgId  'SCR 84204 -- 12/27/05 - Frank
  
  Set gBB.User = robjUser
  
  
  Initialize ByVal vlMyProgId, robjComm, robjLogonEnvironment, robjConfig, robjPatient
  
  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnP.ICHCSII_CoreAPI2_Initialize2", "A/P", vbExclamation

End Sub

Private Function ICHCSII_CoreAPI3_CloseView2(ByVal viViewType As Integer) As Boolean
Dim Success As Boolean

  On Error GoTo ErrHandler
  
  'If A/P or A/P child module is in the middle of processing then do not close module.
  If mLock.Locked Then
    ICHCSII_CoreAPI3_CloseView2 = False
    goARM.StopFarTx NCID_ASSESMENT_AND_PLAN_APPLICATION & gsARM_ACT_CLOSE, enmARMTxStatusFailed  'ARM SCR 19540 MSolano 1/16/2002
    Exit Function
  End If

  mLock.StartLock "ICHCSII_CoreAPI3_CloseView2"

  If cwbMainViewRunning Then

    'SCR 20059 Phil Crowder 1/2/02  Added mSwitchingSection
    If Not mDataSaved Or mSwitchingSection Then
      Success = ShowUnassociatedItems     'Returns true if the user wants to exit A/P
      If Success Then Success = ShowInvalidPrimaryDiagnosis
      
      mSwitchingSection = False
      
      'scr 26517  Phil Crowder 8/12/02
      DoEvents
    Else
      Success = True
    End If

    If Success Then
      If Not mDataSaved Then Success = Save
      If Success Then CloseAnP

      ICHCSII_CoreAPI3_CloseView2 = Success
    Else
      ICHCSII_CoreAPI3_CloseView2 = False
    End If
  Else
      ICHCSII_CoreAPI3_CloseView2 = True
  End If

  mLock.StopLock "ICHCSII_CoreAPI3_CloseView2"
  
  '>>> ARM SCR 19540 MSolano 1/16/2002
  If ICHCSII_CoreAPI3_CloseView2 = False Then
    goARM.StopFarTx NCID_ASSESMENT_AND_PLAN_APPLICATION & gsARM_ACT_CLOSE, enmARMTxStatusFailed
  End If
  '<<< ARM SCR 19540
  
  Exit Function
  
ErrHandler:
  mLock.StopLock "ICHCSII_CoreAPI3_CloseView2"
  ICHCSII_CoreAPI3_CloseView2 = False
  gShared.ShowVBError Err.Number, Err.Description, "AnP.ICHCSII_CoreAPI3_CloseView2", "A/P", vbExclamation
  
End Function

'Not Used by A/P module
Private Function IChild_GetAutoCiteNarrative(TestType As String, NumPeriods As Integer, Period As AC_Period) As String
'
End Function

'Passes encounter module interface object to the A/P module.
Private Function IChild_SetParent(Parent As IParent) As Boolean
  Set gBB.EncParent = Parent
End Function

Private Sub mAnPMainView_ProcedureSelected(ByVal IsSelected As Boolean)
  If mMenuModifiers = 0 Then Exit Sub
  
  If mMenuModifierEnabled <> IsSelected Then
    mMenuModifierEnabled = IsSelected
    
    If mMenuModifierEnabled Then
      gBB.Comm.UpdateMenuItem mMenuModifiers, "ENABLE"
    Else
      gBB.Comm.UpdateMenuItem mMenuModifiers, "DISABLE"
    End If
  End If
  
End Sub

'<<< SCR 38048 MSolano 6/24/2003
Private Sub mAnPMainView_ProcedureModified()
  If mMenuModifiers = 0 Then
    Exit Sub
  ElseIf cwbMainViewRunning Then
    mAnPMainView.DisplayModifiersDialog
  End If
End Sub
'>>>SCR 38048 MSolano 6/24/2003

'SCR 23329  Phil Crowder 4/24/02
'Needed for Telcons because encounters passes a reference to Diagnosis when a telcon
'is edited in the Quick Telcon screen.
Public Property Get AnPChildInterface() As ChildInterface
  Set AnPChildInterface = gBB.AnPChildInterface
End Property

Public Property Let HasTemplate(ByVal NewValue As Boolean)
  gBB.AnPChildInterface.HasTemplate = NewValue
End Property

Private Sub AddMenuItems()
Dim EncounterParent As IParent
Dim InterModuleComm As Comm
Dim LoggedInUser As ICHCSII_User

Const Priv_Patient_Disposition As String = "Patient_Disposition"
Const Priv_Current_Encounter As String = "Current_Encounter"
Const Priv_S_O As String = "S_O"
Const Priv_Order_Sets_Orders As String = "Orders/Order_Sets_Orders"
Const Priv_Patient_Plan_Of_Care As String = "Patient_Plan_Of_Care"

  On Error GoTo ErrHandler
  
  Set EncounterParent = gBB.EncParent
  Set LoggedInUser = gBB.User
  Set InterModuleComm = gBB.Comm
  
  With EncounterParent
    'scr 25777 7/17/02  Phil Crowder
    If LoggedInUser.HasPrivilegeEx(Priv_Current_Encounter, Priv_Read) Then
      mMenuPreviewSummary = .AddMenuItem(0, "&Preview", "summary.bmp", "Preview Assessment & Plan")
    End If
    
    mMenuSave = .AddMenuItem(0, "&Save", "save.bmp", "Save Assessment & Plan")
    
    '** 12/27/05 - Frank
    If gShared.IsAppMode(modeTheater) Then
        'SCR 84204
        glMenuPrint = .AddMenuItem(0, "Print Orders", "print.bmp", "Print Orders")
        
        'Default tab is Diagnosis, Enable only for Orders Summary tab
        gBB.Comm.UpdateMenuItem glMenuPrint, "DISABLE"
    End If
    
    mMenuDelete = .AddMenuItem(0, gsARMTx_Is_Order_Caption_Delete, "delete.bmp", "Delete from Assessment & Plan")
    
    gMenuDelete = mMenuDelete
    goARM.InitCmplxTxCls gsARM_Tx_DelOrder, _
                         NCID_ENCOUNTER_APPLICATION & gsARMTx_Is_Order_Caption_Delete & CStr(gMenuDelete), _
                         "Order Type", "CHCS", eARM_String8, "CDR", eARM_String8
    
    mMenuTemplates = .AddMenuItem(0, "&Templates", "select_encounter.bmp", "Select an Encounter Template")
    
    If LoggedInUser.HasPrivilegeEx(Priv_Patient_Disposition, Priv_Read) Or LoggedInUser.HasPrivilegeEx(Priv_Patient_Disposition, Priv_Write) Then
      mMenuNext = .AddMenuItem(0, "D&isposition", "Next.bmp", "Go to the Disposition screen")
      
      If (gBB.CurrentEncounter.EncounterType = InpatientNote _
        And Val(gBB.CurrentEncounter.InptNoteTypeNCID) <> NCID_Discharge_Note) Then
        gBB.Comm.UpdateMenuItem mMenuNext, "DISABLE"
      End If
    End If
    
    mMenuSign = .AddMenuItem(0, "Si&gn", "dispsign.bmp", "Sign the Completed Encounter")
       
    If LoggedInUser.HasPrivilegeEx(Priv_Patient_Plan_Of_Care, Priv_Read) And Len(gShared.CmdLineSwitches("CPG")) > 0 Then
      mMenuGoalSummary = .AddMenuItem(0, "&Goal Summary", "Health_History.bmp", "Review patient goal summary")
    End If
                
    If (Len(gShared.CmdLineSwitches("WAM")) > 0 And gShared.IsAppMode(modeCDR)) Or gShared.IsAppMode(modeCHCSI_GUI) Then
      mMenuModifiers = .AddMenuItem(0, "&Modifiers", "Details.bmp", "Procedure Modifiers")
      
      InterModuleComm.UpdateMenuItem mMenuModifiers, "DISABLE"
      mMenuModifierEnabled = False
      
      If Not (mAnPMainView Is Nothing) Then
        mAnPMainView.ProcedureSelected
      End If
    End If
    
    If LoggedInUser.HasPrivilegeEx(Priv_Order_Sets_Orders, Priv_Write) Then
      mMenuSubmitAll = .AddMenuItem(0, "Submit &All", "export_template.bmp", "Submit All Queued Orders")
      
      'JT SCR 34705 - Enable submit all menu item based upon whether there are queued orders for submission
      UpdateMenuState mMenuSubmitAll
            
    End If
    
  End With
    
  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnP.AddMenuItems", "A/P", vbExclamation

End Sub

'create listners for OE
Private Sub AddMessages()
Dim InterModuleComm As Comm

    On Error GoTo ErrHandler
      
    Set InterModuleComm = gBB.Comm
    
    With InterModuleComm
      .AddListenerForMessage Me, NCID_OE_LAB_APPLICATION, hlACTION_NEW, NCID_ASSESMENT_AND_PLAN_APPLICATION
      .AddListenerForMessage Me, NCID_OE_LAB_APPLICATION, hlACTION_RENEW, NCID_ASSESMENT_AND_PLAN_APPLICATION
      .AddListenerForMessage Me, NCID_OE_LAB_APPLICATION, hlACTION_DELETE, NCID_ASSESMENT_AND_PLAN_APPLICATION
      .AddListenerForMessage Me, NCID_OE_LAB_APPLICATION, hlACTION_MODIFY, NCID_ASSESMENT_AND_PLAN_APPLICATION
      .AddListenerForMessage Me, NCID_OE_LAB_APPLICATION, hlACTION_RENEW_MODIFY, NCID_ASSESMENT_AND_PLAN_APPLICATION
      .AddListenerForMessage Me, NCID_OE_LAB_APPLICATION, hlACTION_ADD_TO_QUEUE, NCID_ASSESMENT_AND_PLAN_APPLICATION
      .AddListenerForMessage Me, NCID_OE_LAB_APPLICATION, hlACTION_RELEASE, NCID_ASSESMENT_AND_PLAN_APPLICATION

      .AddListenerForMessage Me, NCID_OE_RAD_APPLICATION, hlACTION_NEW, NCID_ASSESMENT_AND_PLAN_APPLICATION
      .AddListenerForMessage Me, NCID_OE_RAD_APPLICATION, hlACTION_RENEW, NCID_ASSESMENT_AND_PLAN_APPLICATION
      .AddListenerForMessage Me, NCID_OE_RAD_APPLICATION, hlACTION_DELETE, NCID_ASSESMENT_AND_PLAN_APPLICATION
      .AddListenerForMessage Me, NCID_OE_RAD_APPLICATION, hlACTION_MODIFY, NCID_ASSESMENT_AND_PLAN_APPLICATION
      .AddListenerForMessage Me, NCID_OE_RAD_APPLICATION, hlACTION_RENEW_MODIFY, NCID_ASSESMENT_AND_PLAN_APPLICATION
      .AddListenerForMessage Me, NCID_OE_RAD_APPLICATION, hlACTION_ADD_TO_QUEUE, NCID_ASSESMENT_AND_PLAN_APPLICATION
      .AddListenerForMessage Me, NCID_OE_RAD_APPLICATION, hlACTION_RELEASE, NCID_ASSESMENT_AND_PLAN_APPLICATION

      .AddListenerForMessage Me, NCID_OE_MED_APPLICATION, hlACTION_NEW, NCID_ASSESMENT_AND_PLAN_APPLICATION
      .AddListenerForMessage Me, NCID_OE_MED_APPLICATION, hlACTION_RENEW, NCID_ASSESMENT_AND_PLAN_APPLICATION
      .AddListenerForMessage Me, NCID_OE_MED_APPLICATION, hlACTION_DELETE, NCID_ASSESMENT_AND_PLAN_APPLICATION
      .AddListenerForMessage Me, NCID_OE_MED_APPLICATION, hlACTION_MODIFY, NCID_ASSESMENT_AND_PLAN_APPLICATION
      .AddListenerForMessage Me, NCID_OE_MED_APPLICATION, hlACTION_RENEW_MODIFY, NCID_ASSESMENT_AND_PLAN_APPLICATION
      .AddListenerForMessage Me, NCID_OE_MED_APPLICATION, hlACTION_ADD_TO_QUEUE, NCID_ASSESMENT_AND_PLAN_APPLICATION
      .AddListenerForMessage Me, NCID_OE_MED_APPLICATION, hlACTION_RELEASE, NCID_ASSESMENT_AND_PLAN_APPLICATION
  
  
  End With

  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnP.AddMessages", "A/P", vbExclamation

End Sub

'Used to unload all loaded forms, delete menu items and remove message listeners when
'the A/P module is closed.
Private Sub CloseAnP()
Dim StatusDialog As StatusDialog
Dim EncounterParent As IParent
Dim InterModuleComm As Comm

  On Error GoTo ErrHandler
  
  Set StatusDialog = New StatusDialog
  
  StatusDialog.Display "Closing down modules", vbNullString, True
  
  'Tell core we are not using Meds any more
  Set InterModuleComm = gBB.Comm
  If InterModuleComm.OE_MEDS_STATE = NCID_ASSESMENT_AND_PLAN_APPLICATION Then
    InterModuleComm.OE_MEDS_STATE = 0
  End If
  Set InterModuleComm = Nothing
  
  StatusDialog.UpdateProgress 10
  
  gBB.AnPChildInterface.DestroyAllReferences
  
  gBB.AnPChildren.CloseOLEServers

  StatusDialog.UpdateProgress 30

  'gBB.AllDiagnoses.Reset @@@@ remove
  
  StatusDialog.UpdateWorking "Closing module A/P"
  StatusDialog.UpdateProgress 50
    
  'Remove AnP form
  Unload mAnPMainView
  Set mAnPMainView = Nothing
  
  cwbMainViewRunning = False
  RemoveMenuItems
  
  StatusDialog.UpdateProgress 75
    
  RemoveMessages

  StatusDialog.UpdateProgress 90
       
  Unload StatusDialog
  DoEvents
  Set StatusDialog = Nothing
       
  Set EncounterParent = gBB.EncParent
  If Not EncounterParent Is Nothing Then EncounterParent.ClosedView
  Set EncounterParent = Nothing
      
  gBB.ClearBlackBoard
          
  Exit Sub
  
ErrHandler:
  'SCR 26383 8/7/02  Phil Crowder
  'Ignore modal display error caused by trying to show status form when a modal
  'form is showing such as New Appointment Form.
  If Err.Number <> 401 Then
    gShared.ShowVBError Err.Number, Err.Description, "AnP.CloseAnP", "A/P", vbExclamation
  End If
  
  Resume Next
  
End Sub

Public Sub CloseView(ByVal viViewType As Integer)
  ICHCSII_CoreAPI3_CloseView2 viViewType
End Sub

'Used by Encounter Ops module to change all diagnoses to problems when the encounter
'is signed.
Public Sub completedEncounter()
Dim EncounterParent As IParent
Dim Proxy As ProxyProblems

  On Error GoTo ErrHandler
    
  mLock.StartLock "completeEncounter"
  
  If Not mEncOpen Then InitializeAnP True      'Just in case encounters does not call Initialize AnP
  
  Set EncounterParent = gBB.EncParent
  
  'SCR 25875  7/19/02  Phil Crowder  Orginally was using Updated for 11 and completed  for 10.
  'After the new encounter interface was added, there is a conflict with the ConsultOrders
  'OrderStatus enum (Updated = 5).
  If EncounterParent.CurrentEncounter.Status = 11 Or EncounterParent.CurrentEncounter.Status = 10 Then
    Set Proxy = gBB.ProblemInterface
    
    Proxy.ConvertDiagnosesToProblems
  End If
      
  mLock.StopLock "completeEncounter"
  
  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnP.completedEncounter", "A/P", vbExclamation
  mLock.StopLock "completeEncounter"
  
End Sub

'Creates a string of all order not associated to a diagnosis or stilled queued.
'Returns:
' String of all unassociated and queued orders if any and a null string on error
' or if orders are associated
Private Function FindUnassociatedOrders() As String
Dim Orders As Collection
Dim Order As iOrder
Dim UnassociatedOrders As String
Dim oProc As OrderProcedure
Dim i As Long

  On Error GoTo ErrHandler
  
  Set Orders = gBB.AllOrders.Filter(, fltActive)
  
  gBB.EncounterInterface.SyncProviderProcedures
  
  For Each Order In Orders
    If Not Order.IsAssociated And Order.State <> osQueued Then
      i = i + 1
      UnassociatedOrders = UnassociatedOrders & i & ".  " & Order.Name & " - Unassociated *" & vbCrLf
    ElseIf Order.State = osQueued Then
      i = i + 1
      UnassociatedOrders = UnassociatedOrders & i & ".  " & Order.Name & " - Still Queued *" & vbCrLf
    ElseIf TypeName(Order) = "OrderProcedure" And IsHIPAAEnabled Then
      Set oProc = Order
      If oProc.Providers = vbNullString Then
         i = i + 1
         UnassociatedOrders = UnassociatedOrders & i & ".  " & Order.Name & " - No Provider *" & vbCrLf
      End If
    End If
  Next
  
  FindUnassociatedOrders = UnassociatedOrders
  
  Exit Function

ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnP.FindUnassociatedOrders", "A/P", vbExclamation

End Function

Public Sub Initialize(ByVal vlMyProgId As Long, ByRef robjComm As Object, _
                      ByRef robjLogonEnviron As Object, ByRef robjConfig As Object, _
                      ByRef robjPatient As Object)

  On Error GoTo ErrHandler
  
  mLock.StartLock "Initialize"
        
  With gBB
    Set .Comm = robjComm
    Set .Logon = robjLogonEnviron
    Set .Config = robjConfig
    Set .Patient = robjPatient
  End With
  
  mLock.StopLock "Initialize"
    
  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnP.Initialize", "A/P", vbExclamation
  mLock.StopLock "Initialize"

End Sub

'This method is used by the encounter module to load A/P data.
Public Sub InitializeAnP(Optional ByVal isComplete As Boolean = False, Optional ByVal HasTemplate As Boolean = False)

  On Error GoTo ErrHandler
  
  mLock.StartLock "InitializeAnP"
  
  gBB.Init
  
  gBB.AnPChildInterface.HasTemplate = HasTemplate
  
  mDataSaved = False
  
  mEncOpen = True
  
  'Don't need to load data to save diagnoses as problems.
  If Not isComplete Then
    LoadAnPData
    SyncWarning
  End If
  
  gBB.AnPChildInterface.Init mLock
  
  mLock.StopLock "InitializeAnP"
    
  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnP.InitializeAnP", "A/P", vbExclamation
  mLock.StopLock "InitializeAnP"

End Sub

'Loads all orders and diagnoses for the current encounter.
Private Sub LoadAnPData()
Dim EncounterParent As IParent
Dim InterModuleComm As Comm
Dim Problems As MMMHISProblems.Problem

  On Error GoTo ErrHandler
  
  Set EncounterParent = gBB.EncParent
  Set InterModuleComm = gBB.Comm
  
  If EncounterParent.CurrentEncounter.EncounterID > 0 And _
     Len(EncounterParent.CurrentEncounter.FacilityNCID) > 0 Then
         
    ResetAnP    'modeCHCSI_GUI may require a reset (does not hurt)
    
    '9/1/03  Phil Crowder   It is possible that data was added to the LCD on the LGS
    'when an appointment that was not originated in PGUI version 2 is amended.
    'Problems caches data locally and would have this data.
    If gShared.IsAppMode(modeCHCSI_GUI) And EncounterParent.CurrentEncounter.Status = Updating Then
      Set Problems = gBB.ReadModule(NCID_PROBLEMS_APPLICATION)
      Problems.ProblemOps.Refresh
    End If
    
    With InterModuleComm
      .SetStatusText "Loading A/P Data"
      
      'Load orders
      gBB.AllOrders.Load
      
      'Load Diagnoses and related items
      gBB.AllDiagnoses.Load
   
      .SetStatusText vbNullString
    End With
  End If
  
  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnP.LoadAnPData", "A/P", vbExclamation
  
End Sub

'Used by core to pass which menu or toolbar item a user clicked in core.
Public Sub MenuItemSelected(ByVal vlMenuId As Long)

  On Error GoTo ErrHandler
    
  mLock.StartLock "MenuItemSelected"

  'Perform the actions associated with the menu item that was selected
  Select Case vlMenuId
    Case mMenuPreviewSummary
      ShowSummary
      
    Case mMenuSave
      Save
      If cwbMainViewRunning Then mAnPMainView.RefreshData
         
    Case glMenuPrint        'SCR 84204 -- 12/27/05 - Frank
        '** 12/27/05 - Frank
        If gShared.IsAppMode(modeTheater) Then
            gBB.AnPChildren.ForwardMessage cwiGENERIC, "print orders summary", NCID_ANP_DIAGNOSIS_APPLICATION
        End If
        
    Case glMenuPrintOrder 'SCR 84207 -- 1/10/06 - Sherry Wang
        If gShared.IsAppMode(modeTheater) Then
            gBB.AnPChildren.ForwardMessage cwiGENERIC, "print orders summary - file menu", NCID_ANP_DIAGNOSIS_APPLICATION
        End If
        
    Case glMenuPrintPreviewOrder 'SCR 84207 -- 1/10/06 - Sherry Wang
        If gShared.IsAppMode(modeTheater) Then
            gBB.AnPChildren.ForwardMessage cwiGENERIC, "print orders summary - file menu", NCID_ANP_DIAGNOSIS_APPLICATION
        End If
        
    Case mMenuDelete
      If cwbMainViewRunning Then
        mAnPMainView.Delete     'Let form decide which grid (diagnosis or orders) user wants to delete from (see AnPMainView)
        
        '>>> SCR 32490 RWalker 3/21/03 Converted tx's to one 'generic' complex
        '>>> where the type will be inserted into metric 7 by the stop call
        goARM.StopCmplxTx NCID_ENCOUNTER_APPLICATION & gsARMTx_Is_Order_Caption_Delete & CStr(gMenuDelete), _
                          glARMStatus, enmARMStop_MeOnly, _
                          gsDelOrderType, _
                          gsDelFromOE, gsDelFromCDR
        gsDelFromOE = ""
        gsDelFromCDR = ""
        
        'JT SCR 34705 - Update menu state enabled/disabled in case this was a queued order that was deleted
        UpdateMenuState mMenuSubmitAll
        
      End If
      
    Case mMenuTemplates
      OpenTemplates
      
    Case mMenuNext
      OpenDisposition
      
    Case mMenuBack
      OpenSO
      
    Case mMenuGoalSummary
      gBB.Comm.Message cwiGENERIC, "show form", NCID_GOAL_SUMMARY, NCID_ANP_DIAGNOSIS_APPLICATION
            
    Case mMenuSign
      SignEncounter
      
    Case mMenuModifiers
      If cwbMainViewRunning Then mAnPMainView.DisplayModifiersDialog
      
    Case mMenuSubmitAll
      SubmitQueuedOrders
         
  End Select
  
  mLock.StopLock "MenuItemSelected"
  
  Exit Sub

ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnP.MenuItemSelected", "A/P", vbExclamation
  mLock.StopLock "MenuItemSelected"

End Sub

'Used by core to pass message from other modules to the A/P module.  See core API for
'descriptions of parameters.
Public Sub Message(ByVal viMessageType As Integer, ByVal vsMessageText As String, ByVal vsSourceServerNCID As String)

  On Error GoTo ErrHandler
   
  If Left$(UCase$(vsMessageText), 6) = "HL7ERR" Then Exit Sub
  
  If vsMessageText = "RECEIEVED ORDER" Then
      'SCR 46201 - Catch this message and enable the submit all button for ITT mode since this mode doesn't send a new order message like other modes
      UpdateMenuState mMenuSubmitAll
      Exit Sub
  End If
  
  If Not (viMessageType = cwiDONE_SHOWING_MODAL_FORM Or viMessageType = cwiSHOWING_MODAL_FORM) Then
    mLock.StartLock viMessageType
  End If
  
  Select Case viMessageType
    Case cwiOE_NEW                      'hlACTION_NEW from OE, New Med, Lab or Rad
      ProcessNewOrderFromOE
      
    Case cwiOE_QUEUE                  'hlACTION_SEND_TO_QUEUE from OE, Queue a Med, Lab or Rad
      ProcessQueuedOrderFromOE
      
    Case cwiDONE_SHOWING_MODAL_FORM
      gBB.Comm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ASSESMENT_AND_PLAN_APPLICATION
      
    Case cwiSHOWING_MODAL_FORM
      gBB.Comm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ASSESMENT_AND_PLAN_APPLICATION
      
    Case cwiOE_RENEW, cwiOE_RENEW_MODIFY, cwiOE_MODIFY, cwiOE_RELEASE   'hlACTION_RENEW, hlACTION_RENEW_MODIFY, hlACTION_MODIFY, hlACTION_RELEASE from OE
      ProcessRenewModifyOrderFromOE vsMessageText
      
    Case cwiOE_DELETE             'hlACTION_DELETE from OE,  delete an existing order from A/P
      'When a users deletes an order from A/P or core (using menu or toolbar) then
      'a message is sent to OE to delete the order form CHCS I and OE.  When the order
      'is deleted by OE it sends this message to A/P.  We do not want to process that
      'message so we make sure A/P is not locked before starting a delete.
      If mLock.Method = CStr(cwiOE_DELETE) Then
        ProcessDeletedOrderFromOE
      End If
      
    Case cwiREFRESH_DATA
      If vsMessageText = "PatientSelected" Then
        ProcessRefreshData
      ' SCR 91389, 4/19/06 Quan
      ElseIf vsMessageText = "DispenseFromLabels" Then
        ' SCR 103466, JQ, 8/18/06: if a/p module isn't loaded, don't send it a message
        If Not gBB.AnPChildren Is Nothing Then
            gBB.AnPChildren.Item(modMed).Message cwiREFRESH_DATA, "", NCID_OE_MED_APPLICATION
        End If
      End If
      
    Case cwiGENERIC
      'Forward to generic message handler
      MessageGeneric vsMessageText

    Case cwiOK_TO_CHANGE_PATIENT
      ProcessOkToChangePatient vsSourceServerNCID
      
    Case cwiOK_TO_CLOSE_VIEW
      ProcessOKToCloseView vsSourceServerNCID
      
    Case cwiORDER_SETS_PROCESSING
      'SCR 23627  Phil Crowder 8/13/02
      'Don't allow a Patient Change or Close of Module if Processing an Order Set
      mProcessingOrderSet = (vsMessageText = "Y")
      
    Case Else
      'A/P doesn't handle the message so forward it to all open A/P child modules.
      If cwbMainViewRunning Then
        gBB.AnPChildren.ForwardMessage viMessageType, vsMessageText, vsSourceServerNCID
      End If
  End Select
    
  mLock.StopLock viMessageType
  
  Exit Sub

ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnP.Message", "A/P", vbExclamation
  mLock.StopLock viMessageType
  
End Sub

'This method handles the generic launch and refresh template message that are passed to
'A/P by other modules (see Message method).
Private Sub MessageGeneric(ByVal MessageText As String)
Dim MessageItems As Variant

  On Error GoTo ErrHandler
   
  If Not cwbMainViewRunning Then Exit Sub
  
  MessageItems = Split(MessageText, "|")
    
  Select Case UCase$(MessageItems(0))
    Case "LAUNCH"
      'Switch tabs on main view to load and display an A/P child module.
      If Len(MessageItems(1)) > 0 Then
        
        'SCR 45227 RWalker 11/19/2003 Suppress local start when from CORE
        gbFromCore = True
        
        Select Case MessageItems(1)
          Case NCID_ANP_DIAGNOSIS_APPLICATION
            SwitchTabs modDiagnosis
            
          Case NCID_ORDERS_SUMMARY_APPLICATION
            SwitchTabs modOrderSummary
          
         Case NCID_ORDER_SETS_APPLICATION
            SwitchTabs modOrderSets
            
          Case NCID_ANP_PROCEDURES_APPLICATION
            SwitchTabs modProcedure
            
          Case NCID_REMINDERS_APPLICATION
            SwitchTabs modReminders
            
          Case NCID_CONSULT_TRACKING_APPLICATION
            SwitchTabs modConsults
            
          Case NCID_OE_LAB_APPLICATION, NCID_GEMS_OE_LAB_APPLICATION
            SwitchTabs modLab
            
          Case NCID_OE_RAD_APPLICATION, NCID_GEMS_OE_RAD_APPLICATION
            SwitchTabs modRad
            
          Case NCID_OE_MED_APPLICATION, NCID_GEMS_OE_MED_APPLICATION
            SwitchTabs modMed
            
          Case NCID_PATIENTINSTRUCTION_APPLICATION
            SwitchTabs modPI
            
          Case NCID_SRTS_OE_APPLICATION  'SRTS
            SwitchTabs modSpec
        End Select
        
        'SCR 45227 RWalker 11/19/2003 Suppress local start when from CORE
        gbFromCore = False
        
      End If
      
    Case "REFRESH TEMPLATES"
      gBB.AnPChildInterface.HasTemplate = True
      
      'Forward generic message to all open modules
      gBB.AnPChildren.ForwardMessage cwiGENERIC, MessageText, NCID_ASSESMENT_AND_PLAN_APPLICATION
      
    Case "CHANGE HEIGHT"
      'SCR 35637  5/13/03 Phil Crowder
      If Not (mAnPMainView Is Nothing) Then
        gBB.AnPChildren.Item(mAnPMainView.tabAnPChildren.CurrTab).Height = Val(MessageItems(1))
        'SCR 82703 Call tabAnPChildren_Click to Refresh the tab
        mAnPMainView.tabAnPChildren_Click
'        mAnPMainView.ContainerHeight
'        gBB.AnPChildren.DisplayServer mAnPMainView.tabAnPChildren.CurrTab, mAnPMainView.vseContainer.hWnd
        mAnPMainView.ZOrder 0           '54267 - bring window back on top after changing height
      End If
          
    Case Else
      MessageText = "Unknown message type"
  End Select

  Exit Sub

ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnP.MessageGeneric", "A/P", vbExclamation

End Sub

Private Sub OpenDisposition()
Dim EncounterParent As Object
Dim EncounterCurrent As CHCSIIEncounterOps.Encounter
      
  On Error GoTo ErrHandler
  
  If gBB.EncParent Is Nothing Then Exit Sub
  
  '--- SCR 25485 MSolano 8/22/02 "workflow" Tx flags
  '--- mere presence of Tx in Tivoli log, not timed
  goARM.StartFarTx gsARMTx_NavigateBtnAnP2Disp

  mLock.StopLock "MenuItemSelected"        'End the lock to allow this module to be closed
  
  'SCR 20059 Phil Crowder 1/2/02  Added mSwitchingSection
  mSwitchingSection = True
  
  'SCR 19825  Phil Crowder 1/18/02
  Set EncounterParent = gBB.EncParent
  
  'SCR 25225 WR
  Set EncounterCurrent = EncounterParent.CurrentEncounter
  
  EncounterCurrent.OpenSection eDisp
  Set EncounterParent = Nothing
  '--- SCR 25485 MSolano 8/22/02
  '--- place stops below where something can go wrong; if so it involved
  '--- the MsgBxARMd; which stops current Tx's w/ fail; these stops are ignored
  goARM.StopFarTx gsARMTx_NavigateBtnAnP2Disp

  Exit Sub

ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnP.OpenDisposition", "A/P", vbExclamation

End Sub

Private Sub OpenSO()
Dim EncounterParent As Object
Dim EncounterCurrent As CHCSIIEncounterOps.Encounter
      
  On Error GoTo ErrHandler
  
  If gBB.EncParent Is Nothing Then Exit Sub
  
  '--- SCR 25485 MSolano 8/22/02 "workflow" Tx flags
  '--- mere presence of Tx in Tivoli log, not timed
  goARM.StartFarTx gsARMTx_NavigateBtnAnP2SO

  mLock.StopLock "MenuItemSelected"        'End the lock to allow this module to be closed
  
  'SCR 20059 Phil Crowder 1/2/02  Added mSwitchingSection
  mSwitchingSection = True
  
  'SCR 19332, 19337  Phil Crowder 1/18/02
  Set EncounterParent = gBB.EncParent
  'SCR 25225 WR
   Set EncounterCurrent = EncounterParent.CurrentEncounter
  EncounterCurrent.OpenSection eSO
  Set EncounterParent = Nothing
  '--- SCR 25485 MSolano 8/22/02
  '--- place stops below where something can go wrong; if so it involved
  '--- the MsgBxARMd; which stops current Tx's w/ fail; these stops are ignored
  goARM.StopFarTx gsARMTx_NavigateBtnAnP2SO

  Exit Sub

ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnP.OpenSO", "A/P", vbExclamation

End Sub

Private Sub OpenTemplates()
Dim Templates As Object
      
  On Error GoTo ErrHandler
  
  Set Templates = gBB.ReadModule(NCID_TEMPLATE_MANAGEMENT)
  
  If Not Templates Is Nothing Then
    Templates.SelectEncounterTemplate
    Set Templates = Nothing
  End If

  Exit Sub

ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnP.OpenTemplates", "A/P", vbExclamation

End Sub

'This method is used by core to open an A/P view.  See the core API for paramter
'descriptions.  The view is loaded but not displayed.  ResizeView method is used to
'display forms.
Public Function OpenView(ByVal viViewType As Integer, ByVal vlParenthWnd As Long, ByRef rhMainForm As Long) As Boolean
Dim EncounterRTF As String

  On Error GoTo ErrHandler
    
  mLock.StartLock "OpenView"
  
  'perform actions based on the "view type"
  If viViewType = cwiREGULAR_VIEW Then
    cwbFirstRegularResize = True
    cwbMainViewRunning = True
    
    Set mAnPMainView = New AnPMainView
    Set gBB.AnPChildInterface.MainForm = mAnPMainView
    
    SetParent mAnPMainView.hWnd, vlParenthWnd
    
    rhMainForm = mAnPMainView.hWnd
    OpenView = True
    
  ElseIf viViewType = cwiPATIENT_SUMMARY_VIEW Then
    OpenView = True
    
  ElseIf viViewType = cwiENCOUNTER_SUMMARY_VIEW Then
    SetParent SummaryDialog.hWnd, vlParenthWnd
        
    EncounterRTF = gBB.EncounterInterface.CreateRTF(True)
    
    SummaryDialog.InitRTF EncounterRTF
    rhMainForm = SummaryDialog.hWnd
    OpenView = True
  
  Else
    Screen.MousePointer = vbDefault
    Beep
    MsgBxARMd "Unsupported view type: " & CStr(viViewType) & " in AnP.OpenView.", vbCritical, "A/P"
    
    OpenView = False
  End If
   
  mLock.StopLock "OpenView"
  
  Exit Function
   
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnP.OpenView", "A/P", vbExclamation
  OpenView = False
  mLock.StopLock "OpenView"
   
End Function

'This method deletes a med, lab or rad from the collection of all orders and makes sure it
'is unassocaited from all diagnoses when the order is deleted in the OE module.
Private Sub ProcessDeletedOrderFromOE()
Dim Proxy As ProxyOE
Dim Order As iOrder

  On Error GoTo ErrHandler
    
  Set Proxy = gBB.OEInterface
  Set Order = Proxy.FindMatchingOrder(True)
  
  If Not (Order Is Nothing) Then
    Order.State = osDelete
    
    Order.Unassociate
  
    Save
    
    If cwbMainViewRunning Then mAnPMainView.RefreshData
  Else
    Proxy.AddDeletedMedNote
  End If

  Exit Sub

ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnP.ProcessDeletedOrderFromOE", "A/P", vbExclamation

End Sub

'This method creates a new lab, med or rad from information created in OE.  If the order
'is queued then it does not create a new one but changes the state of the queued order
'to new so that it will be saved by A/P.
Private Sub ProcessNewOrderFromOE()
Dim Proxy As ProxyOE
Dim Order As iOrder
Dim Diagnosis As AnPDiagnosis
Dim Association As OrderAssociation
Dim sForegroundAppNCID As String

  On Error GoTo ErrHandler
   
  'SCR 48699 - Only process message when encounter module is the foreground application before
  gBB.Comm.Message cwiFOREGROUND_APPLICATION, sForegroundAppNCID, NCID_CW_APPLICATION, NCID_ASSESMENT_AND_PLAN_APPLICATION
  If sForegroundAppNCID <> CStr(NCID_ENCOUNTER_APPLICATION) Then Exit Sub
  
  Set Proxy = gBB.OEInterface
  Set Order = Proxy.FindMatchingOrder
        
  If Not (Order Is Nothing) Then
    Proxy.UpdateOrder Order
  Else
    Set Order = Proxy.CreateOrder
    Set Association = Order.CreateAssociation
  End If
    
  Order.State = osNew
  
  If cwbMainViewRunning Then
    Set Diagnosis = mAnPMainView.SelectedDiagnosis          'Determine curently selected diagnosis if any
    
    If Not (Diagnosis Is Nothing) And Not (Association Is Nothing) Then
      Diagnosis.AssociateOrder Association
    End If
    
    If Not (Diagnosis Is Nothing) Then
      mAnPMainView.RefreshDiagnosis Diagnosis
    End If
  
    mAnPMainView.AddOrder Order
    
    'JT SCR 34705 - Enable submit all menu item based upon whether there are queued orders for submission
    UpdateMenuState mMenuSubmitAll
      
  End If
    
  gBB.Comm.Message cwiOE_NEW, "RECEIEVED ORDER", NCID_BROADCAST_MSG, NCID_OE_MED_APPLICATION
  
  Exit Sub

ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnP.ProcessNewOrderFromOE", "A/P", vbExclamation

End Sub

'This method is used to determine if it is ok to change patients.
Private Sub ProcessOkToChangePatient(ByRef SourceServerNCID As String)
Dim Success As Boolean

  On Error GoTo ErrHandler
   
  If mLock.Method = CStr(cwiOK_TO_CHANGE_PATIENT) And Not mProcessingOrderSet Then
    'Make sure to save the A/P data before allowing changing of patients.
    If Not (gBB.EncParent Is Nothing) Then
      
      If Not gBB.CurrentEncounter.Status = complete Then
        Success = ShowUnassociatedItems
        If Success Then Success = ShowInvalidPrimaryDiagnosis
        
        If Success Then Save
        If cwbMainViewRunning Then mAnPMainView.RefreshData
      End If
    
    Else
      Success = True
    End If
  End If
  
  'SCR 23627  Phil Crowder 8/13/02
  'Don't allow a Patient Change or Close of Module if Processing an Order Set
  If Not Success Or mProcessingOrderSet Then
    gBB.Comm.Message cwiOK_TO_CHANGE_PATIENT, "N", SourceServerNCID, NCID_ASSESMENT_AND_PLAN_APPLICATION
    mDataSaved = False
  Else
    gBB.Comm.Message cwiOK_TO_CHANGE_PATIENT, "Y", SourceServerNCID, NCID_ASSESMENT_AND_PLAN_APPLICATION
    mDataSaved = True
  End If

  Exit Sub

ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnP.ProcessOkToChangePatient", "A/P", vbExclamation

End Sub

Private Sub ProcessOKToCloseView(ByRef SourceServerNCID As String)
Dim Success As Boolean

  On Error GoTo ErrHandler
   
  If mDataSaved Then Exit Sub
  
  If mLock.Method = CStr(cwiOK_TO_CLOSE_VIEW) And Not mProcessingOrderSet Then
    'Make sure to save the A/P data before allowing changing of patients.
    If Not (gBB.EncParent Is Nothing) Then
    
      If Not gBB.CurrentEncounter.Status = complete Then
        Success = ShowUnassociatedItems
        If Success Then Success = ShowInvalidPrimaryDiagnosis
        
        If Success Then Save
        If cwbMainViewRunning Then mAnPMainView.RefreshData
      End If
    
    Else
      Success = True
    End If
  End If
      
  'SCR 23627  Phil Crowder 8/13/02
  'Don't allow a Close of Module if Processing an Order Set
  If Success And Not mProcessingOrderSet Then
    gBB.Comm.Message cwiOK_TO_CLOSE_VIEW, "Y", SourceServerNCID, NCID_ASSESMENT_AND_PLAN_APPLICATION
    mDataSaved = True
  Else
    gBB.Comm.Message cwiOK_TO_CLOSE_VIEW, "N", SourceServerNCID, NCID_ASSESMENT_AND_PLAN_APPLICATION
    mDataSaved = False
  End If

  Exit Sub

ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnP.ProcessOKToCloseView", "A/P", vbExclamation
      
End Sub

'This method handle queuing an order from the OE module.
'Orders are queued by OE to improve the processes of adding orders.  The orders
'and assoicated items for diagnosis have a state of queued.  The order is also
'added to OrderSets.  The user can then Submit all orders in the Order Set module.
'The Order Sets module sends messages to OE to create new orders.  OE then sends a
'message to A/P that a new order is created.  We check the orders collection for
'the order and change its state from queued to new (see cwiOE_NEW case).
'The associated items' states are changed from queued to associated for the order.
Private Sub ProcessQueuedOrderFromOE()
Dim Success As Boolean
Dim ErrorMessage As String
Dim OrderSets As AnPOLEServer
Dim Proxy As ProxyOE
Dim Order As iOrder
Dim Diagnosis As AnPDiagnosis
Dim Association As OrderAssociation

  On Error GoTo ErrHandler
   
  Set OrderSets = gBB.AnPChild(modOrderSets)
  
  If OrderSets.Server Is Nothing Then
    OrderSets.InitView
    If cwbMainViewRunning Then gBB.AnPChild(modOrderSets).OpenView mAnPMainView.vseContainer.hWnd
  End If
  
  If Not (OrderSets.Server Is Nothing) Then
    Set Proxy = gBB.OEInterface
    
    Set Order = Proxy.CreateOrder
    Order.State = osQueued
    
    If cwbMainViewRunning Then
      Set Diagnosis = mAnPMainView.SelectedDiagnosis          'Determine curently selected diagnosis if any
      
      If Not (Diagnosis Is Nothing) Then
        Set Association = Order.CreateAssociation
        
        Diagnosis.AssociateOrder Association
      End If
    End If
    
    Success = OrderSets.Server.Mtd_bPushOrder(Proxy.Tablet, "QID_" & Order.Key, ErrorMessage)
    
    'Remove order if it failed during addition to order sets
    If Not Success Then
      Order.State = osRemove
      Order.Unassociate
    End If
  
    If cwbMainViewRunning Then
      If Not (Diagnosis Is Nothing) Then mAnPMainView.RefreshDiagnosis Diagnosis
      If Not (Order Is Nothing) Then mAnPMainView.AddOrder Order
      
      'JT SCR 34705 - Enable submit all menu item based upon whether there are queued orders for submission
      UpdateMenuState mMenuSubmitAll
      
    End If
  End If
    
  gBB.Comm.Message cwiOE_NEW, "RECEIEVED ORDER", NCID_BROADCAST_MSG, NCID_OE_MED_APPLICATION

  Exit Sub

ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnP.ProcessQueuedOrderFromOE", "A/P", vbExclamation

End Sub

'This method is used to either save the current data or refresh the display of the A/P data.
Private Sub ProcessRefreshData()
  On Error GoTo ErrHandler
   
  If cwbMainViewRunning Then
    Save
    If cwbMainViewRunning Then mAnPMainView.RefreshData
  Else
    If Not (gBB.EncParent Is Nothing) And cwbMainViewRunning Then
      InitializeAnP
      If cwbMainViewRunning Then mAnPMainView.RefreshData
    End If
  End If

  Exit Sub

ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnP.ProcessRefreshData", "A/P", vbExclamation

End Sub


'This method updates the data of a med, lab or rad with information from the OE module
'if the order is in A/P.  If the order is not part of the collection of orders then
'the order is added as a new order.
Private Sub ProcessRenewModifyOrderFromOE(ByVal MessageText As String)
Dim varMessage As Variant
Dim Orders As Collection
Dim Lab As OrderLab
Dim Rad As OrderRad
Dim Med As OrderMed
Dim Order As iOrder
Dim i As Long
Dim Proxy As ProxyOE
Dim Diagnosis As AnPDiagnosis
Dim Association As OrderAssociation

  On Error GoTo ErrHandler
   
  'Modify the data for an existing order.
  varMessage = Split(MessageText, "|")
  
  Set Orders = gBB.AllOrders.Filter(typLab, fltActive)
  For i = 1 To Orders.Count
    Set Lab = Orders.Item(i)
    
    If Lab.OrderID = varMessage(1) Then
      Set Order = Lab
      Exit For
    End If
  Next
  
  If Order Is Nothing Then
    Set Orders = gBB.AllOrders.Filter(typMed, fltActive)
    For i = 1 To Orders.Count
      Set Med = Orders.Item(i)
      
      If Med.OrderID = varMessage(1) Then
        Set Order = Med
        Exit For
      End If
    Next
  End If
  
  If Order Is Nothing Then
    Set Orders = gBB.AllOrders.Filter(typRad, fltActive)
    For i = 1 To Orders.Count
      Set Rad = Orders.Item(i)
      
      If Rad.OrderID = varMessage(1) Then
        Set Order = Rad
        Exit For
      End If
    Next
  End If
  
  Set Proxy = gBB.OEInterface
    
  If Not (Order Is Nothing) Then
    Proxy.UpdateOrder Order
  Else
    Set Order = Proxy.CreateOrder
    Set Association = Order.CreateAssociation
  End If
    
  If Order.State <> osUnknown Then
    Order.State = osModified
  Else
    Order.State = osNew
  End If
  
  If cwbMainViewRunning Then
    Set Diagnosis = mAnPMainView.SelectedDiagnosis          'Determine curently selected diagnosis if any
    
    If Not (Diagnosis Is Nothing) And Not (Association Is Nothing) Then
      Diagnosis.AssociateOrder Association
    End If
    
    If Not (Diagnosis Is Nothing) Then
      mAnPMainView.RefreshDiagnosis Diagnosis
    End If
  
    mAnPMainView.AddOrder Order
  End If
  
  gBB.Comm.Message cwiOE_NEW, "RECEIEVED ORDER", NCID_BROADCAST_MSG, NCID_OE_MED_APPLICATION

  Exit Sub

ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnP.ProcessRenewModifyOrderFromOE", "A/P", vbExclamation

End Sub

'Required for Core Complatability.
Public Function QueryView(ByVal viViewType As Integer) As Boolean

   Select Case viViewType
      Case cwiREGULAR_VIEW
         QueryView = True              'this view is NOT optional
      Case cwiPATIENT_SUMMARY_VIEW
         QueryView = False
      Case cwiENCOUNTER_SUMMARY_VIEW
         QueryView = False
      Case cwiPROPERTY_VIEW
         QueryView = True
      Case Else                        'not supported
         QueryView = False
   End Select

End Function

'Remove menu items
Private Sub RemoveMenuItems()
Dim InterModuleComm As Comm

  On Error Resume Next
  
  Set InterModuleComm = gBB.Comm
  
  If mMenuPreviewSummary > 0 Then
    InterModuleComm.DeleteMenuItem mMenuPreviewSummary
    mMenuPreviewSummary = 0
  End If

  If mMenuSave > 0 Then
    InterModuleComm.DeleteMenuItem mMenuSave
    mMenuSave = 0
  End If

  If mMenuDelete > 0 Then
    InterModuleComm.DeleteMenuItem mMenuDelete
    mMenuDelete = 0
  End If

  If mMenuNext > 0 Then
    InterModuleComm.DeleteMenuItem mMenuNext
    mMenuNext = 0
  End If

  If mMenuGoalSummary > 0 Then
    InterModuleComm.DeleteMenuItem mMenuGoalSummary
    mMenuGoalSummary = 0
  End If

  If mMenuBack > 0 Then
    InterModuleComm.DeleteMenuItem mMenuBack
    mMenuBack = 0
  End If

  If mMenuTemplates > 0 Then
    InterModuleComm.DeleteMenuItem mMenuTemplates
    mMenuTemplates = 0
  End If
   
  'SCR 16617 11/12/01  Phil Crowder
  If mMenuSign > 0 Then
    InterModuleComm.DeleteMenuItem mMenuSign
    mMenuSign = 0
  End If

  If mMenuModifiers > 0 Then
    InterModuleComm.DeleteMenuItem mMenuModifiers
    mMenuModifiers = 0
  End If

  If mMenuSubmitAll > 0 Then
    InterModuleComm.DeleteMenuItem mMenuSubmitAll
    mMenuSubmitAll = 0
  End If

    'SCR 84204
    If glMenuPrint > 0 Then
        InterModuleComm.DeleteMenuItem glMenuPrint
        glMenuPrint = 0
    End If
        
  InterModuleComm.Message cwiGENERIC, "Refresh_Menu", NCID_CW_APPLICATION, NCID_ASSESMENT_AND_PLAN_APPLICATION
   
End Sub

'Remove message listeners
Private Sub RemoveMessages()
Dim InterModuleComm As Comm

  On Error Resume Next
  
  Set InterModuleComm = gBB.Comm
  
  With InterModuleComm
    .RemoveListenerForMessage NCID_OE_LAB_APPLICATION, hlACTION_NEW, NCID_ASSESMENT_AND_PLAN_APPLICATION
    .RemoveListenerForMessage NCID_OE_LAB_APPLICATION, hlACTION_RENEW, NCID_ASSESMENT_AND_PLAN_APPLICATION
    .RemoveListenerForMessage NCID_OE_LAB_APPLICATION, hlACTION_DELETE, NCID_ASSESMENT_AND_PLAN_APPLICATION
    .RemoveListenerForMessage NCID_OE_LAB_APPLICATION, hlACTION_MODIFY, NCID_ASSESMENT_AND_PLAN_APPLICATION
    .RemoveListenerForMessage NCID_OE_LAB_APPLICATION, hlACTION_RENEW_MODIFY, NCID_ASSESMENT_AND_PLAN_APPLICATION
    .RemoveListenerForMessage NCID_OE_LAB_APPLICATION, hlACTION_ADD_TO_QUEUE, NCID_ASSESMENT_AND_PLAN_APPLICATION
    .RemoveListenerForMessage NCID_OE_LAB_APPLICATION, hlACTION_RELEASE, NCID_ASSESMENT_AND_PLAN_APPLICATION
    
    .RemoveListenerForMessage NCID_OE_RAD_APPLICATION, hlACTION_NEW, NCID_ASSESMENT_AND_PLAN_APPLICATION
    .RemoveListenerForMessage NCID_OE_RAD_APPLICATION, hlACTION_RENEW, NCID_ASSESMENT_AND_PLAN_APPLICATION
    .RemoveListenerForMessage NCID_OE_RAD_APPLICATION, hlACTION_DELETE, NCID_ASSESMENT_AND_PLAN_APPLICATION
    .RemoveListenerForMessage NCID_OE_RAD_APPLICATION, hlACTION_MODIFY, NCID_ASSESMENT_AND_PLAN_APPLICATION
    .RemoveListenerForMessage NCID_OE_RAD_APPLICATION, hlACTION_RENEW_MODIFY, NCID_ASSESMENT_AND_PLAN_APPLICATION
    .RemoveListenerForMessage NCID_OE_RAD_APPLICATION, hlACTION_ADD_TO_QUEUE, NCID_ASSESMENT_AND_PLAN_APPLICATION
    .RemoveListenerForMessage NCID_OE_RAD_APPLICATION, hlACTION_RELEASE, NCID_ASSESMENT_AND_PLAN_APPLICATION
    
    .RemoveListenerForMessage NCID_OE_MED_APPLICATION, hlACTION_NEW, NCID_ASSESMENT_AND_PLAN_APPLICATION
    .RemoveListenerForMessage NCID_OE_MED_APPLICATION, hlACTION_RENEW, NCID_ASSESMENT_AND_PLAN_APPLICATION
    .RemoveListenerForMessage NCID_OE_MED_APPLICATION, hlACTION_DELETE, NCID_ASSESMENT_AND_PLAN_APPLICATION
    .RemoveListenerForMessage NCID_OE_MED_APPLICATION, hlACTION_MODIFY, NCID_ASSESMENT_AND_PLAN_APPLICATION
    .RemoveListenerForMessage NCID_OE_MED_APPLICATION, hlACTION_RENEW_MODIFY, NCID_ASSESMENT_AND_PLAN_APPLICATION
    .RemoveListenerForMessage NCID_OE_MED_APPLICATION, hlACTION_ADD_TO_QUEUE, NCID_ASSESMENT_AND_PLAN_APPLICATION
    .RemoveListenerForMessage NCID_OE_MED_APPLICATION, hlACTION_RELEASE, NCID_ASSESMENT_AND_PLAN_APPLICATION
  End With
  
End Sub

'Resets the diagnosis and orders collections.  Resets main view.
Private Sub ResetAnP()

  On Error GoTo ErrHandler
  
  gBB.AllDiagnoses.Reset
  gBB.AllOrders.Reset
  
  If Not (mAnPMainView Is Nothing) Then mAnPMainView.Reset
    
  Exit Sub
   
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnP.ResetAnP", "A/P", vbExclamation
  Resume Next
  
End Sub

'This method is used by core to display a view.  See core API document for paramter
'descriptions.
Public Sub ResizeView(ByVal viViewType As Integer, ByVal vlLeft As Long, ByVal vlTop As Long, ByVal vlWidth As Long, ByVal vlHeight As Long)

  On Error GoTo ErrHandler
      
  mLock.StartLock "ResizeView"
    
  'reset global "client area" parameters
  cwlLeft = vlLeft
  cwlTop = vlTop
  cwlWidth = vlWidth
  cwlHeight = vlHeight
   
  'perform different actions based upon the "view type"
  Select Case viViewType
    Case cwiREGULAR_VIEW
      If Not cwbFirstRegularResize Then
        'set flag so that the form "resize" event will be ignored
        cwbFirstRegularResize = True
        If mAnPMainView.WindowState <> vbNormal Then mAnPMainView.WindowState = vbNormal
           
        mAnPMainView.Visible = False
                    
        'set flag so that the form "resize" event will be processed
        cwbFirstRegularResize = False
        mAnPMainView.Move cwlLeft, cwlTop, cwlWidth, cwlHeight
              
        mAnPMainView.Visible = True
  
      Else
        'first "regular" resize, need to perform "initailization" procedures
        'reset flag so this code will not be executed again
        cwbFirstRegularResize = False
        
        If gBB.EncParent Is Nothing Then
          MsgBxARMd "Please Open an Encounter first", vbCritical
          Exit Sub
        End If
      
        If Not mEncOpen Then
          InitializeAnP
        Else
          mEncOpen = False
        End If
           
        'display the "main" form
        mAnPMainView.Initialize mLock
        If Not mAnPMainView.Display Then
          mLock.ForceUnlock
          ICHCSII_CoreAPI3_CloseView2 cwiREGULAR_VIEW
        End If
        
        AddMenuItems
        AddMessages
        
      End If

    Case Else
      Screen.MousePointer = vbDefault
      Beep
      MsgBxARMd "Unsupported resize view type: " & CStr(viViewType) & " in AnPMainView.ResizeView.", vbExclamation, "Problems"
  End Select
      
  mLock.StopLock "ResizeView"
   
  Exit Sub
   
ErrHandler:
  mLock.StopLock "ResizeView"
  Resume Next                       'SCR 17667  Phil Crowder 10/91/01
                                    'Core calls this method on shutdown and tries to
                                    'display AnPMainView when the modal app lock screen is up
                                    'which cause an error.
  
End Sub

'Saves orders, diagnosis and related items.  Updates encounter with saved information.
'Returns:
' True if all save actions are successful and false if any save operation fails
Private Function Save() As Boolean

  Save = gBB.EncounterInterface.Save

  If gShared.IsAppMode(modeCHCSI_GUI) Then
    gBB.ADMInterface.Save
  End If
      
End Function

Private Sub ShowSummary()
Dim EncounterRTF As String
  
  EncounterRTF = gBB.EncounterInterface.CreateRTF(True)
  
  Screen.MousePointer = vbDefault
  SummaryDialog.InitRTF EncounterRTF
  
  gBB.Comm.Message cwiSHOWING_MODAL_FORM, vbNullString, NCID_CW_APPLICATION, NCID_ASSESMENT_AND_PLAN_APPLICATION
  SummaryDialog.Show vbModal
  gBB.Comm.Message cwiDONE_SHOWING_MODAL_FORM, vbNullString, NCID_CW_APPLICATION, NCID_ASSESMENT_AND_PLAN_APPLICATION

End Sub

'SCR 34746 4/16/03  Phil Crowder
Private Function ShowInvalidPrimaryDiagnosis() As Boolean
Dim CloseAnP As Integer
Dim Warning As String
Dim Diagnoses As Collection

  On Error GoTo ErrHandler

  If Len(gShared.CmdLineSwitches("WAM")) = 0 And Not gShared.IsAppMode(modeCHCSI_GUI) Then
    ShowInvalidPrimaryDiagnosis = True
    Exit Function
  End If
  
  If cwbMainViewRunning And Not mAnPMainView.Visible Then
    ShowInvalidPrimaryDiagnosis = True
    Exit Function
  End If

  'Find primary diagnosis
  Set Diagnoses = gBB.AllDiagnoses.Filter(fltActive)
  
  If Diagnoses.Count = 0 Then
    ShowInvalidPrimaryDiagnosis = True
    Exit Function
  End If
  
  If UCase$(Left$(Diagnoses.Item("1").FullICDCode, 1)) = "E" Then
    Screen.MousePointer = vbDefault
    
    'JT BEGIN SCR 38897 - Don't give the users the option to close AnP with this error - notify them with appropriate message and cancel close operation
    Warning = "This diagnosis is a 'cause of injury code' or 'E-code' and cannot be " & _
              "used as the primary diagnosis.  You will need to either remove it or " & _
              "select another primary diagnosis before you close."

    
    CloseAnP = Not WarningDialog.Display(Warning, True)
    'JT END SCR 38897
    
    DoEvents
    Set WarningDialog = Nothing
    
    ShowInvalidPrimaryDiagnosis = CloseAnP

    Screen.MousePointer = vbHourglass
  Else
    ShowInvalidPrimaryDiagnosis = True
  End If
    
  Exit Function

ErrHandler:
  If Err.Number = 401 Then
    'SCR 26090  8/2/02 Phil Crowder
    'Another modal form is showing so we have to skip this and allow the save
    ShowInvalidPrimaryDiagnosis = True
  Else
    ShowInvalidPrimaryDiagnosis = False
    gShared.ShowVBError Err.Number, Err.Description, "AnP.ShowInvalidPrimaryDiagnosis", "A/P", vbExclamation
  End If

End Function

'This method determines if there are any unassociated orders for the current encounter.
'If there are any other orders that are not associated, the user is asked if they wish
'to leave A/P without associating the orders.
Private Function ShowUnassociatedItems() As Boolean
Dim CloseAnP As Integer
Dim Warning As String

  On Error GoTo ErrHandler

  If cwbMainViewRunning And Not mAnPMainView.Visible Then
    ShowUnassociatedItems = True
    Exit Function
  End If

  If InpatientEncounter = True Then
    Warning = vbNullString
  Else
    'Find any unassociated orders
    Warning = FindUnassociatedOrders
  End If
  
  Screen.MousePointer = vbDefault
  
  If Len(Warning) > 0 Then
    'Unassociated order(s) found
    Warning = "The following items have been found incomplete or still queued within A/P:" & vbNewLine & _
              Warning & vbNewLine & vbNewLine & vbTab & vbTab & "Do you still wish to close?"

    CloseAnP = WarningDialog.Display(Warning, False)
    DoEvents
    Set WarningDialog = Nothing
    
    ShowUnassociatedItems = CloseAnP
  Else
    ShowUnassociatedItems = True
  End If

  Screen.MousePointer = vbHourglass
  
  Exit Function

ErrHandler:
  If Err.Number = 401 Then
    'SCR 26090  8/2/02 Phil Crowder
    'Another modal form is showing so we have to skip this and allow the save
    ShowUnassociatedItems = True
  Else
    ShowUnassociatedItems = False
    gShared.ShowVBError Err.Number, Err.Description, "AnP.ShowUnassociatedItems", "A/P", vbExclamation
  End If
  
End Function

Private Sub SignEncounter()
Dim EncounterParent As Object
Dim CurrentEncounter As CHCSIIEncounterOps.Encounter
      
  On Error GoTo ErrHandler
  
  '--- SCR 25485 MSolano 8/22/02 "workflow" Tx flags
  '--- mere presence of Tx in Tivoli log, not timed
  goARM.StartFarTx gsARMTx_NavigateBtnAnP2Sign
  
  'SCR 16617 11/12/01  Phil Crowder
  Set EncounterParent = gBB.EncParent
  mLock.ForceUnlock
  CloseView 0
  
  'SCR 25225 WR
  Set CurrentEncounter = EncounterParent.CurrentEncounter

  If Not cwbMainViewRunning Then
    CurrentEncounter.LockingObject.UnlockSection eap, 0
    CurrentEncounter.SignEncounter
  End If
  
  Set EncounterParent = Nothing
  
  '--- SCR 25485 MSolano 8/22/02
  '--- place stops below where something can go wrong; if so it involved
  '--- the MsgBxARMd; which stops current Tx's w/ fail; these stops are ignored
  goARM.StopFarTx gsARMTx_NavigateBtnAnP2Sign

  Exit Sub

ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnP.SignEncounter", "A/P", vbExclamation

End Sub

'Call a method is the Order Sets module to begin processing the order set.
'When a queued item is created it is added to the current encounter's order set.  When
'the orders are submitted, the A/P modules recieves a message for each order and changes
'the state of the order from queued to new.  Queued orders are displayed underlined.
Private Sub SubmitQueuedOrders()
Dim ErrorMessage As String
Dim OrderSets As AnPOLEServer

  On Error GoTo ErrHandler
    
  If mAnPMainView Is Nothing Then Exit Sub
  
  mLock.StartLock "AnP.SubmitQueuedOders"
  
  mAnPMainView.tabAnPChildren.CurrTab = modOrderSets        'Switch to order sets module
  
  Set OrderSets = gBB.AnPChild(modOrderSets)
  
  If Not (OrderSets.Server Is Nothing) Then
    'Call Order Sets module to submit orders
    OrderSets.Server.Mtd_bSubmitAll ErrorMessage
  Else
    MsgBxARMd "Unable to load Order Sets Automatically.  Please switch to the Order Sets tab.  Click the Select All button.  Then click the Submit button.", vbOKOnly, "A/P"
  End If
  
  mLock.StopLock "AnP.SubmitQueuedOders"
  
  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnP.SubmitQueuedOders", "A/P", vbCritical
  mLock.StopLock "AnP.SubmitQueuedOders"
    
End Sub

'Checks to see if tab is enabled and if it is then it makes it the current tab.
Private Function SwitchTabs(ByVal ToTab As enumAnPModules) As Boolean
Dim Enabled As Boolean

  If mAnPMainView Is Nothing Then Exit Function
  
  With mAnPMainView.tabAnPChildren
    Enabled = .TabEnabled(ToTab)
    If Enabled Then .CurrTab = ToTab
  End With

  SwitchTabs = Enabled
  
End Function

Private Sub SyncWarning()
Dim WarningRequired As Boolean
Dim SyncTest As SyncAnP
Dim Problems As MMMHISProblems.Problem
    
  On Error GoTo ErrHandler
  
  If InpatientEncounter = True Then
    Exit Sub
  End If
      
  'Init Problems
  Set Problems = gBB.ReadModule(NCID_PROBLEMS_APPLICATION)
    
  'SCR 16771 Phil Crowder 10/24/01
  'Make that the Encounter RTF and A/P are in sync.  If not display warning message.
  Set SyncTest = New SyncAnP
  WarningRequired = SyncTest.SyncWarningRequired(gBB.CurrentEncounter.SectionRTF(eap))
  
  'SCR 22949, 22899 Phil Crowder 4/11/02 Refresh problems cache to make sure that is not causing the sync issue.
  If WarningRequired Then
    Problems.ProblemOps.Refresh
    ResetAnP
    LoadAnPData

    WarningRequired = SyncTest.SyncWarningRequired(gBB.CurrentEncounter.SectionRTF(eap))
    'SCR 107019
    If WarningRequired Then
        WarningRequired = SyncTest.SyncWarningRequired(gBB.EncounterInterface.CreateRTF)
        If Not WarningRequired Then
            gBB.EncParent.SetRTF gBB.EncounterInterface.CreateRTF
        End If
    End If
  End If
  
  If WarningRequired Then
    WarningDialog.Display "The Encounter you have selected may contain corrupt data - verify that all data is correct before signing.", True
    Set WarningDialog = Nothing
  End If

  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnP.SyncWarning", "A/P", vbExclamation

End Sub

'Called by the encounter module to save A/P data for a telecon encounter.
Public Function teleconEnd() As Boolean

  On Error GoTo ErrHandler
  
  mLock.StartLock "teleconEnd"
  
  teleconEnd = gBB.EncounterInterface.Save
  
  If teleconEnd Then gBB.ClearBlackBoard
  
  mLock.StopLock "teleconEnd"
  
  Exit Function
  
ErrHandler:
  teleconEnd = False
  gShared.ShowVBError Err.Number, Err.Description, "AnP.teleconEnd", "A/P", vbExclamation

End Function

'JT BEGIN SCR 34705 - Update menu state enabled/disabled
Private Sub UpdateMenuState(ByVal lMenuItem As Long)

On Error GoTo ErrHandler

   If lMenuItem = 0 Then Exit Sub
   
   Select Case lMenuItem
   
   Case mMenuSubmitAll
      
      'Enable submit all menu item based upon whether there are queued orders for submission
      If gBB.AllOrders.Filter(typAll, fltQueued).Count > 0 Then
         gBB.Comm.UpdateMenuItem mMenuSubmitAll, "ENABLE"
      Else
         gBB.Comm.UpdateMenuItem mMenuSubmitAll, "DISABLE"
      End If
   
   End Select
   
ErrHandler:
   If CBool(Err) Then gShared.ShowVBError Err.Number, Err.Description, "AnP.UpdateMenuState", "A/P", vbExclamation

End Sub
'JT END SCR 34705



