VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Appointments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"37E2657A03B8"
Option Explicit
Option Base 0

Implements ICHCSII_CoreAPI2

Public Sub Initialize(ByVal vlMyProgId As Long, robjComm As Object, robjLogonEnvironment As Object, robjConfig As Object, robjPatient As Object)
    MsgBxARMd "Internal Error: CHCSII_Appointments module requires initialization via ICHCSII_CoreAPI2."
End Sub

Public Sub ICHCSII_CoreAPI2_Initialize2(ByVal vlMyProgId As Long, ByRef robjComm As Object, ByRef robjLogonEnvironment As Object, ByRef robjConfig As Object, ByRef robjPatient As Object, ByRef robjUser As Object)
    
    On Error GoTo ErrHandler
    
    '---save the "ProgId" assigned to me by CW---
    cwlMyProgId = vlMyProgId
    
    '---create pointers to objects---
    Set cwobjCWComm = robjComm
    Set cwobjLogonEnviron = robjLogonEnvironment
    Set cwobjConfig = robjConfig
    Set cwobjPatient = robjPatient
    Set gobjUser = robjUser

    '---set global flags---
    cwbMainViewRunning = False
    
    '---create OLE objects---
    If gobjAppointmentOps Is Nothing Then
        Set gobjAppointmentOps = New CHCSII_AppointmentClient.AppointmentOps
        gobjAppointmentOps.Initialize vlMyProgId, robjComm, robjLogonEnvironment, robjConfig, robjPatient, robjUser
    End If
    
    'We only set gobjTelconOps to new here.
    'It is only initialized (from NewVisit.frm) when creating a new Telcon
    'for the first time to reduce startup overhead.
    'Most telcon appts are created from telcon module; can check ProgID prop
    'to see if initialized.
    
    If gobjClinicOps Is Nothing Then Set gobjClinicOps = New CHCSII_ClinicClient.ClinicOpsEx
    'SCR-42772
    If gobjShared.IsAppMode(modeTheater) Then
        Dim objListManagment As Object
        Set objListManagment = cwobjCWComm.InitializeOLEServer(NCID_LIST_MGMT_APPLICATION)
        Set gobjListMgmtOps = objListManagment.ListManagementOps
        Set objListManagment = Nothing
    End If
    
    gbGetDefaultConfigValues
    
    Exit Sub

ErrHandler:
    gobjShared.ShowVBError Err.Number, Err.Description, "CHCSII_Appointments.Initialize", "CHCSII_Appointments", vbExclamation

End Sub

Private Sub AddMenus()
    On Error GoTo ErrHandler
    If gobjShared.IsAppMode(modeTheater) Then
        AddMenus_GEMS
        Exit Sub
    End If
    
    '--- SCR-18843
    If goCoreMenuID Is Nothing Then
      Set goCoreMenuID = New cCoreMenuID
    End If
    
    '<<< SCR-18843 MSolano 11/16/01
    '--- MENU ACTIONS OTHER THAN SETTING APPT SELECTION FILTER CRITERIA
    With cwobjCWComm
      goCoreMenuID.Add .AddActionMenuItem("&Refresh", "refresh.bmp", cwlMyProgId, "Refresh from database"), apptMnuActRefreshMenuId
      goCoreMenuID.Add .AddActionMenuItem("&Open Appt", "new.bmp", cwlMyProgId, "Open appointment"), apptMnuActOpenMenuId
      
      If gobjShared.IsAppMode(modeCHCSI_GUI) Then
        goCoreMenuID.Add .AddActionMenuItem("C&ancel", "cancel.bmp", cwlMyProgId, "Cancel selected appointment"), apptMnuActCancelMenuId
        goCoreMenuID.Add .AddActionMenuItem("&Transfer", "transfer.bmp", cwlMyProgId, "Transfer appointment to another Provider"), apptMnuActTransferMenuId
      End If
      
        'SCR-31309
        goCoreMenuID.Add .AddActionMenuItem("&New Appt", "new.bmp", cwlMyProgId, "Create new appointment"), apptMnuActNewMenuId
      If Not gobjShared.IsAppMode(modeCHCSI_GUI) Then
        goCoreMenuID.Add .AddActionMenuItem("Ch&eck-In", "checkin.bmp", cwlMyProgId, "Check-In selected patient(s)"), apptMnuActCheckinMenuId
        goCoreMenuID.Add .AddActionMenuItem("C&heck-Out", "checkout.bmp", cwlMyProgId, "Check-Out selected patient"), apptMnuActCheckoutMenuId
        goCoreMenuID.Add .AddActionMenuItem("C&ancel", "cancel.bmp", cwlMyProgId, "Cancel selected appointment"), apptMnuActCancelMenuId
        goCoreMenuID.Add .AddActionMenuItem("&Transfer", "transfer.bmp", cwlMyProgId, "Transfer appointment to another Provider"), apptMnuActTransferMenuId
        goCoreMenuID.Add .AddActionMenuItem("Add &Providers", "link.bmp", cwlMyProgId, "Add additional providers"), apptMnuActAddProvider
      End If
      
      goCoreMenuID.Add .AddActionMenuItem("Pri&nt Appts", "print.bmp", cwlMyProgId, "Print appointment list"), apptMnuActPrint
      goCoreMenuID.Add .AddActionMenuItem("&View Comments", "note09.ico", cwlMyProgId, "View appointment comments"), apptMnuActComments
      
      If Not gobjShared.IsAppMode(modeCHCSI_GUI) Then
        goCoreMenuID.Add .AddActionMenuItem("&Ins. Form", "reports.bmp", cwlMyProgId, "Print Insurance Form"), apptMnuActPrintInsuranceForm
        goCoreMenuID.Add .AddActionMenuItem("&Undo Cancel", "undo.bmp", cwlMyProgId, "Restore a cancelled appointment"), apptMnuActUndoCancel
      Else
        .UpdateMenuItem apptMnuActPrint, "DISABLE"
        GoTo SkipForCHCS_GUI
      
      End If
                     
      'KDunne 6/14/2002
      'SCR 8900. For Dental providers, they can print SF600 if its a medical encounter,
      '          or SF603A for dental providers. All others can still only print SF600.
      '
      
      
      If gobjUser.HasPrivilegeEx(Priv_Dental_Access, Priv_Read) Then
         gPrintSF600 = .AddPrintMenuItem("&SF603A/SF600", cwlMyProgId, "Print SF600")
         gPrintPreviewSF600 = .AddPrintPreviewMenuItem("SF603A/SF600", cwlMyProgId, "Print SF600")
      Else
         gPrintSF600 = .AddPrintMenuItem("S&F600", cwlMyProgId, "Print SF600")
         gPrintPreviewSF600 = .AddPrintPreviewMenuItem("SF600", cwlMyProgId, "Print SF600")
         
      End If
      
      'SCR-28478 S.McAvoy 11/12/02
        .UpdateMenuItem gPrintSF600, "DISABLE"
        .UpdateMenuItem gPrintPreviewSF600, "DISABLE"
      
      '--- MENU ITEMS SETTING SELECTION FILTER CRITERIA
      goCoreMenuID.Add .AddViewMenuItem("Default", cwlMyProgId, "Default Filter"), apptDefault
      goCoreMenuID.Add .AddViewMenuItem("Today", cwlMyProgId, "Today"), apptToday
      goCoreMenuID.Add .AddViewMenuItem("Today+Incomplete", cwlMyProgId, "Today+Incomplete"), apptTodayPlusIncomplete
      goCoreMenuID.Add .AddViewMenuItem("Any Status", cwlMyProgId, "Any Status"), apptAll
      goCoreMenuID.Add .AddViewMenuItem("Pending", cwlMyProgId, "Pending Only"), apptPending
      goCoreMenuID.Add .AddViewMenuItem("Checked-In", cwlMyProgId, "Checked-In"), apptCheckedIn
      goCoreMenuID.Add .AddViewMenuItem("Waiting", cwlMyProgId, "Default Filter"), apptWaiting
      goCoreMenuID.Add .AddViewMenuItem("InProgress", cwlMyProgId, "InProgress Only"), apptInProgress
      goCoreMenuID.Add .AddViewMenuItem("Needs Co-Signature", cwlMyProgId, "Needs Co-Signature"), apptNeedsCoSignature
      goCoreMenuID.Add .AddViewMenuItem("Updated", cwlMyProgId, "Updated"), apptUpdated
      goCoreMenuID.Add .AddViewMenuItem("Completed", cwlMyProgId, "Completed Only"), apptCompleted
      goCoreMenuID.Add .AddViewMenuItem("Incomplete", cwlMyProgId, "Incomplete"), apptIncomplete
      goCoreMenuID.Add .AddViewMenuItem("Checked In or Pending", cwlMyProgId, "Checked In or Pending"), apptCheckedInPending       'SCR 7454 - Phil Crowder 7/9/01
      
      
      '>>> SCR-18843
      
      '<<< SCR 17687 - 11/26/01 - Walker - Check-In Is Not Applicable to Theater mode
      '<<< SCR ITT 816 S.McAvoy 9/28/02 Walker incorrectly checking for appmode.
      If gobjShared.IsAppMode(modeTheater) Then
          goCoreMenuID.Remove .DeleteMenuItem(goCoreMenuID(apptMnuActCheckinMenuId))
          goCoreMenuID.Remove .DeleteMenuItem(goCoreMenuID(apptMnuActPrintInsuranceForm))
      End If
SkipForCHCS_GUI:
    End With
    
    '>>> SCR 17687

    Exit Sub
    
ErrHandler:
    gobjShared.ShowVBError Err.Number, Err.Description, "CHCSII_Appointments.AddMenus", "CHCSII_Appointments", vbExclamation
    
End Sub
Private Sub AddMenus_GEMS() 'SCR-42772
    On Error GoTo ErrHandler
    
    '--- SCR-18843
    If goCoreMenuID Is Nothing Then
      Set goCoreMenuID = New cCoreMenuID
    End If
    
    '--- MENU ACTIONS OTHER THAN SETTING APPT SELECTION FILTER CRITERIA
    With cwobjCWComm
      goCoreMenuID.Add .AddActionMenuItem("Refresh", "refresh.bmp", cwlMyProgId, "Refresh from database"), apptMnuActRefreshMenuId
    If LenB(gobjShared.CmdLineSwitches.Item("GEMS_ARMY")) > 0 Then
      goCoreMenuID.Add .AddActionMenuItem("New Visit", "new.bmp", cwlMyProgId, "Create new visit"), apptMnuActNewMenuId
      goCoreMenuID.Add .AddActionMenuItem("Print Visits", "print.bmp", cwlMyProgId, "Print Patient Tracking list"), apptMnuActPrint
    Else
      goCoreMenuID.Add .AddActionMenuItem("New Appointment", "new.bmp", cwlMyProgId, "Create new appointment"), apptMnuActNewMenuId
      goCoreMenuID.Add .AddActionMenuItem("Print Appointments", "print.bmp", cwlMyProgId, "Print Appointment list"), apptMnuActPrint
    End If
      goCoreMenuID.Add .AddActionMenuItem("Check-Out", "checkout.bmp", cwlMyProgId, "Check-Out selected patient"), apptMnuActCheckoutMenuId
      goCoreMenuID.Add .AddActionMenuItem("Cancel", "cancel.bmp", cwlMyProgId, "Cancel selected visit"), apptMnuActCancelMenuId
      goCoreMenuID.Add .AddActionMenuItem("Transfer", "transfer.bmp", cwlMyProgId, "Transfer visit to another Provider"), apptMnuActTransferMenuId
      goCoreMenuID.Add .AddActionMenuItem("Add Providers", "link.bmp", cwlMyProgId, "Add additional providers"), apptMnuActAddProvider
      goCoreMenuID.Add .AddActionMenuItem("View Comments", "note09.ico", cwlMyProgId, "View Patient Tracking comments"), apptMnuActComments
      goCoreMenuID.Add .AddActionMenuItem("Undo Cancel", "undo.bmp", cwlMyProgId, "Restore a cancelled visit"), apptMnuActUndoCancel
                     
      'SCR 8900. For Dental providers, they can print SF600 if its a medical encounter,
      '          or SF603A for dental providers. All others can still only print SF600.
      '
      If gobjUser.HasPrivilegeEx(Priv_Dental_Access, Priv_Read) Then
         gPrintSF600 = .AddPrintMenuItem("SF603A/SF600", cwlMyProgId, "Print SF600")
         gPrintPreviewSF600 = .AddPrintPreviewMenuItem("SF603A/SF600", cwlMyProgId, "Print SF600")
      Else
         gPrintSF600 = .AddPrintMenuItem("SF600", cwlMyProgId, "Print SF600")
         gPrintPreviewSF600 = .AddPrintPreviewMenuItem("SF600", cwlMyProgId, "Print SF600")
      End If
      
      '--- MENU ITEMS SETTING SELECTION FILTER CRITERIA
      goCoreMenuID.Add .AddViewMenuItem("Default", cwlMyProgId, "Default Filter"), apptDefault
      goCoreMenuID.Add .AddViewMenuItem("Today", cwlMyProgId, "Today"), apptToday
      goCoreMenuID.Add .AddViewMenuItem("Today+Incomplete", cwlMyProgId, "Today+Incomplete"), apptTodayPlusIncomplete
      goCoreMenuID.Add .AddViewMenuItem("Any Status", cwlMyProgId, "Any Status"), apptAll
      goCoreMenuID.Add .AddViewMenuItem("Pending Only", cwlMyProgId, "Pending Only"), apptPending
      goCoreMenuID.Add .AddViewMenuItem("Checked-In", cwlMyProgId, "Checked-In"), apptCheckedIn
      goCoreMenuID.Add .AddViewMenuItem("Waiting", cwlMyProgId, "Default Filter"), apptWaiting
      goCoreMenuID.Add .AddViewMenuItem("InProgress Only", cwlMyProgId, "InProgress Only"), apptInProgress
      goCoreMenuID.Add .AddViewMenuItem("Needs Co-Signature", cwlMyProgId, "Needs Co-Signature"), apptNeedsCoSignature
      goCoreMenuID.Add .AddViewMenuItem("Updated", cwlMyProgId, "Updated"), apptUpdated
      goCoreMenuID.Add .AddViewMenuItem("Completed Only", cwlMyProgId, "Completed Only"), apptCompleted
      goCoreMenuID.Add .AddViewMenuItem("Incomplete", cwlMyProgId, "Incomplete"), apptIncomplete
      goCoreMenuID.Add .AddViewMenuItem("Checked In or Pending", cwlMyProgId, "Checked In or Pending"), apptCheckedInPending       'SCR 7454 - Phil Crowder 7/9/01
    End With
    
    '>>> SCR 17687

    Exit Sub
    
ErrHandler:
    gobjShared.ShowVBError Err.Number, Err.Description, "CHCSII_Appointments.AddMenus_GEMS", "CHCSII_Appointments", vbExclamation
    
End Sub

Public Function QueryView(ByVal viViewType As Integer) As Boolean
    Select Case viViewType
        Case cwiREGULAR_VIEW
            QueryView = True    '--this view is NOT optional--
        Case cwiPATIENT_SUMMARY_VIEW
            QueryView = True
        Case cwiENCOUNTER_SUMMARY_VIEW
            QueryView = False
        Case cwiPROPERTY_VIEW
            QueryView = True
        Case Else   '--unknown "view type", not supported--
            QueryView = False
    End Select
End Function

'----------------------------------------------------------------------
'--- PROCEDURE: MenuItemSelected
'---
'--- CHANGE HISTORY:
'--- * MSolano 2/4/2002 SCR-16109
'---   Instrumented err handler w/ Tx stop / aborted status
'----------------------------------------------------------------------
Public Sub MenuItemSelected(ByVal vlMenuId As Long)
Dim strOk As String

    On Error GoTo ErrHandler
    
    '--- SCR-18843 MSolano 11/16/01
    '--- Reconciliation of several variables tracking appt selection criteria:
    '---  frmAppointmentView.mstrSelectionStatus
    '---  frmAppointmentView.menumSelection
    '---  giLastViewMenuId
    Screen.MousePointer = vbHourglass
    DoEvents
    Select Case vlMenuId
        
        Case goCoreMenuID(apptMnuActRefreshMenuId)
            frmAppointmentView.mCHCS_GUI_RefreshQuery = True
            frmAppointmentView.RefreshData
            
        Case goCoreMenuID(apptDefault)
            frmAppointmentView.SetDefaultSelections
            frmAppointmentView.RefreshData , False
        
        Case goCoreMenuID(apptToday)
            frmAppointmentView.SetApptSelectionCriteria apptToday     'SCR 18843  Phil Crowder 1/3/02
            frmAppointmentView.RefreshData , False
        
        Case goCoreMenuID(apptTodayPlusIncomplete)
            frmAppointmentView.SetApptSelectionCriteria apptTodayPlusIncomplete     'SCR 18843  Phil Crowder 1/3/02
            frmAppointmentView.RefreshData , False
        
        Case goCoreMenuID(apptAll)
            frmAppointmentView.SetApptSelectionCriteria apptAll                     'SCR 18843  Phil Crowder 1/3/02
            frmAppointmentView.RefreshData , False
        
        Case goCoreMenuID(apptPending)
            frmAppointmentView.SetApptSelectionCriteria apptPending                 'SCR 18843  Phil Crowder 1/3/02
            frmAppointmentView.RefreshData , False
        
        Case goCoreMenuID(apptCompleted)
            frmAppointmentView.SetApptSelectionCriteria apptCompleted               'SCR 18843  Phil Crowder 1/3/02
            frmAppointmentView.RefreshData , False
        
        Case goCoreMenuID(apptIncomplete)
            frmAppointmentView.SetApptSelectionCriteria apptIncomplete              'SCR 18843  Phil Crowder 1/3/02
            frmAppointmentView.RefreshData , False
        
        Case goCoreMenuID(apptInProgress)
            frmAppointmentView.SetApptSelectionCriteria apptInProgress              'SCR 18843  Phil Crowder 1/3/02
            frmAppointmentView.RefreshData , False
        
        Case goCoreMenuID(apptCheckedIn)
            frmAppointmentView.SetApptSelectionCriteria apptCheckedIn               'SCR 18843  Phil Crowder 1/3/02
            frmAppointmentView.RefreshData , False
        
        Case goCoreMenuID(apptWaiting)
            frmAppointmentView.SetApptSelectionCriteria apptWaiting                 'SCR 18843  Phil Crowder 1/3/02
            frmAppointmentView.RefreshData , False
        
        Case goCoreMenuID(apptNeedsCoSignature)
            frmAppointmentView.SetApptSelectionCriteria apptNeedsCoSignature        'SCR 18843  Phil Crowder 1/3/02
            frmAppointmentView.RefreshData , False
        
        Case goCoreMenuID(apptUpdated)
            frmAppointmentView.SetApptSelectionCriteria apptUpdated                 'SCR 18843  Phil Crowder 1/3/02
            frmAppointmentView.RefreshData , False
        
        Case goCoreMenuID(apptCheckedInPending)     'SCR 7454 - Phil Crowder 7/9/01
            frmAppointmentView.SetApptSelectionCriteria apptCheckedInPending        'SCR 18843  Phil Crowder 1/3/02
            frmAppointmentView.RefreshData , False
            
        Case goCoreMenuID(apptMnuActCheckinMenuId)
            frmAppointmentView.CheckInAppointment
        
        Case goCoreMenuID(apptMnuActNewMenuId)
            '--- SCR-ARM-16109 MSolano 8/23/01
            '--- user opened menu item from Core
            '--- this is the first time we know both
            '--- * ID of OLEServer;
            '--- * appt menu clicked
            '--- far call reconciled in Patient Search :: frmPatientSearch.Form_Activate
            goARM.StartFarTx gsARMTx_OpenPatSrchFromAppts
            
            'SCR 21533 Phil Crowder 2/26/02
            cwobjCWComm.Message cwiOK_TO_CHANGE_PATIENT, strOk, NCID_CW_APPLICATION, NCID_CW_APPLICATION, False
            If strOk <> "N" Then frmAppointmentView.NewAppointment
        
        Case goCoreMenuID(apptMnuActOpenMenuId) 'SCR-31309
            frmAppointmentView.OpenAppointment
        
        Case goCoreMenuID(apptMnuActCancelMenuId)
            frmAppointmentView.CancelAppointment
        
        Case goCoreMenuID(apptMnuActCheckoutMenuId)
            frmAppointmentView.CheckOutAppointment
        
        Case goCoreMenuID(apptMnuActTransferMenuId)
            frmAppointmentView.TransferAppointment
        
        Case goCoreMenuID(apptMnuActAddProvider)
            frmAppointmentView.AddProviders
        
        Case goCoreMenuID(apptMnuActPrint)
            'frmAppointmentView.PrintAppointments
            PrintAppointments
            
        Case goCoreMenuID(apptMnuActComments)
            frmAppointmentView.ViewComments
        
        Case goCoreMenuID(apptMnuActPrintInsuranceForm)
            frmAppointmentView.PrintInsuranceForm
        
        Case goCoreMenuID(apptMnuActUndoCancel)
            frmAppointmentView.UndoCancel
        
        Case gPrintSF600
            frmAppointmentView.PrintEncounters False
        
        Case gPrintPreviewSF600
            frmAppointmentView.PrintEncounters True
    End Select
    Screen.MousePointer = vbNormal



    Exit Sub

ErrHandler:
    gobjShared.ShowVBError Err.Number, Err.Description & " MenuItemSelected: " & vlMenuId, "CHCSII_Appointments.MenuItemSelected", "CHCSII_Appointments", vbExclamation

End Sub


Public Sub PatientSelected()
    
    On Error GoTo ErrHandler
    
    If cwbMainViewRunning Then
        frmAppointmentView.VerifySelectedPatient
    End If
    
    Exit Sub

ErrHandler:
    gobjShared.ShowVBError Err.Number, Err.Description, "CHCSII_Appointments.PatientSelected", "CHCSII_Appointments", vbExclamation
    
End Sub

Private Sub Class_Initialize()

#If coiDebug Then
    gCOI_DL.debugStart
    gDR.TimerTime = 0
    gCOI_DL.DebugLog "COI31", "APPOINTMENT", "", "Begin", gDR
#End If

    ARM_InitAllTransactionsForAppts
    Set gobjShared = New CWShared
    App.HelpFile = gobjShared.CWHelpFile
    Set goCoreMenuID = New cCoreMenuID
    
    '------------------------------------
    '  This statement is needed to debug
    '    this OLE Server.
    '------------------------------------
    #If ciDEBUG = 1 Then
        If cwbDebugLogon = False Then
            Exit Sub
        End If
    #End If
    If LenB(gobjShared.CmdLineSwitches.Item("RNDS")) > 0 And _
      Not gobjShared.IsAppMode(modeTheater) Then 'SCR-RNDS
        gbRNDS_On_and_NonTheater_mode = True
    End If
End Sub

Public Sub Properties(robjLogonEnviron As Object, robjConfig As Object, ByVal vsOwnerNCID As String, ByVal vbCalledByDesigner As Boolean)
    
    On Error GoTo ErrHandler
    
    If vbCalledByDesigner Then
        Set cwobjLogonEnviron = robjLogonEnviron
        Set robjConfig = cwobjConfig
        vsOwnerNCID = cwobjLogonEnviron.UserNCID
    End If
    
    Exit Sub
    
ErrHandler:
    gobjShared.ShowVBError Err.Number, Err.Description, "CHCSII_Appointments.Properties", "CHCSII_Appointments", vbExclamation
    
    If vbCalledByDesigner Then
        Set cwobjLogonEnviron = Nothing
        Set cwobjConfig = Nothing
    End If
    
End Sub
'-------------------------------------------------------------------------------------------------
'  Sub Name: ResizeView
'
'  Description: This method is called when a change has occurred in the "client area"
'               used by this OLE server and it needs to resize a particular "view".
'
'  Parameters:
'    viViewType - Contains the "type" of view being requested to resize by CW.
'    vlLeft - Contains the "left" coordinate for the OLE server.
'    vlTop - Contains the "top" coordinate for the OLE server.
'    vlWidth - Contains the "width" for the OLE server.
'    vlHeight - Contains the "height" for the OLE server.
'-------------------------------------------------------------------------------------------------
Public Sub ResizeView(ByVal viViewType As Integer, ByVal vlLeft As Long, ByVal vlTop As Long, ByVal vlWidth As Long, ByVal vlHeight As Long)
    On Error GoTo ErrHandler

    cwlLeft = vlLeft
    cwlTop = vlTop
    cwlWidth = vlWidth
    cwlHeight = vlHeight
    
    '---perform different actions based upon the "view type"---
    If viViewType = cwiREGULAR_VIEW Then
        If cwobjCWComm.Message(cwiMODAL_FORM_SHOWING, "", NCID_CW_APPLICATION, NCID_CLINIC_SCHEDULE_APPLICATION) Then
           Exit Sub
        End If
        If Not cwbFirstRegularResize Then
            With frmAppointmentView
              .WindowState = vbNormal
              .Visible = False
              .Move cwlLeft, cwlTop, cwlWidth, cwlHeight
              .Visible = True
            End With
        Else    '--first "regular" resize, need to perform "initailization" procedures--
            cwbFirstRegularResize = False
        
            If Not frmAppointmentView.Display() Then
                DeleteMenus
                Unload frmAppointmentView
                cwbMainViewRunning = False
                Exit Sub
            End If
        End If
    
    ElseIf viViewType = cwiPATIENT_SUMMARY_VIEW Then
        
        If Not cwbFirstPatientSummaryResize Then
            '---resize form to fill entire "client area", NOTE: This generates a "resize" event---
            
            If cwlHeight > 200 Then
                gfrmHealthHistAppts.grdAppt.Height = cwlHeight - 250
            End If
            If cwlWidth > 125 Then
                gfrmHealthHistAppts.grdAppt.Width = cwlWidth - 125
            End If
            gfrmHealthHistAppts.Move cwlLeft, cwlTop, cwlWidth, cwlHeight
            'gfrmHealthHistAppts.grdAppt.Move 0, 300, cwlWidth - 125, cwlHeight - 250
            
        Else    '--first "patient summary" resize, need to perform "initailization" procedures--
            '---reset flag so this code will not be executed again---
            cwbFirstPatientSummaryResize = False
        
            '---display the "summary" form---
            gfrmHealthHistAppts.Display
        End If
    Else
        Screen.MousePointer = vbDefault
        Beep
        MsgBxARMd "Unsupported resize view type: " & CStr(viViewType) & " in Appointments.ResizeView.", vbExclamation, "Vital Signs"
    End If
    
    Exit Sub
    
ErrHandler:
    If Err.Number = 401 Then
      'SCR 18742  11/28/01  Phil Crowder
      'Core calls this method on shutdown and tries to refresh the display
      'of frmAppointmentView when the modal app lock screen is up which cause an error.
      Resume Next
    Else
      gobjShared.ShowVBError Err.Number, Err.Description, "CHCSII_Appointments.ResizeView", "CHCSII_Appointments", vbExclamation
    End If

End Sub

Public Function OpenView(ByVal viViewType As Integer, ByVal vlParenthWnd As Long, rhMainForm As Long) As Boolean
    
    On Error GoTo ErrHandler
    
    
    If gobjShared.IsAppMode(modeCHCSI_GUI) Then 'SCR-39944
        gbTelcons_Enabled_CHCS_GUI = gobjUser.GetAttribute("TelconsEnabled")
    End If
    If viViewType = cwiREGULAR_VIEW Then
        cwbFirstRegularResize = True
        cwbMainViewRunning = True
        
        AddMenus
        
        SetParent frmAppointmentView.hwnd, vlParenthWnd
        
        rhMainForm = frmAppointmentView.hwnd
        
    ElseIf viViewType = cwiPATIENT_SUMMARY_VIEW Then
        cwbFirstPatientSummaryResize = True
        cwbPatientSummaryViewRunning = True
        '---create a patient summary form---
        Set gfrmHealthHistAppts = New frmHealthHistAppts
        'gfrmHealthHistAppts.SummaryViewType = cwiPATIENT_SUMMARY_VIEW
        
        '----------------------------------------------------------------------
        '  Reset the "parent window" for this OLE server's "main" summary form.
        '
        '  NOTE: This invokes the form's "Load" event, but the form is NOT
        '        shown.  Make sure that you DO NOT perform a "Me.Show" in the
        '        form's "Load" event.
        '----------------------------------------------------------------------
        
        SetParent gfrmHealthHistAppts.hwnd, vlParenthWnd
    Else
        Screen.MousePointer = vbDefault
        Beep
        MsgBxARMd "Unsupported view type: " & CStr(viViewType) & " in Appointments.OpenView..", vbCritical, "Vital Signs"
        
        '---set return value---
        OpenView = False
        Screen.MousePointer = vbDefault
        Exit Function
        
    End If
    OpenView = True
    Screen.MousePointer = vbDefault
    Exit Function
    
ErrHandler:
    gobjShared.ShowVBError Err.Number, Err.Description, "CHCSII_Appointments.OpenView", "CHCSII_Appointments", vbCritical
    
End Function

Public Sub Message(ByVal viMessageType As Integer, ByRef vsMessageText As String, ByVal vsSourceServerNCID As String)
              
    On Error GoTo ErrHandler
        
    Select Case viMessageType
    
        Case cwiSELECT_PATIENT
            PatientSelected
        
        Case cwiFOREGROUND_APPLICATION
            ' SCR 103302, jq, 8/11/06: force a refresh s.t. extraneous items are cleaned up
            'SCR 103820 Refresh was causing appointments to be unselected when menu items selected
            'added login in frmAppointmentView.FillOneRow_GEMS to fix problems with SCR 103302
            'RefreshData
            'SCR 106488
            'SCR 107152
            On Error Resume Next
'            frmAppointmentView.Visible = False
            frmAppointmentView.Visible = True
        Case cwiREFRESH
            If cwbMainViewRunning Then
                '---NOTE: Changing the form's "visible" property ensures ALL child controls get repainted---
                frmAppointmentView.Visible = False
                frmAppointmentView.Visible = True
            End If

        Case cwiPRINT_REGULAR_VIEW
            PrintAppointments
            
        Case cwiPROPERTIES
            frmAppointmentView.EditSelections
            
        Case cwiABOUT
            DisplayAbout
            
        Case cwiREFRESH_DATA 'SCR-42772
            'Re-Enable the refresh menu button to force a refresh SCR 64202
            cwobjCWComm.UpdateMenuItem gPrintSF600, "ENABLE"
            
            If vsMessageText <> vbNullString Then
                If Mid(vsMessageText, 1, 6) = "APPTID" Then
                    UpdateAppointment vsMessageText
                    Exit Sub
                End If
            End If
            If gobjShared.IsAppMode(modeTheater) Then
                If vsSourceServerNCID = NCID_CLINIC_SCHEDULE_APPLICATION Or _
                    vsSourceServerNCID = NCID_CW_APPLICATION Or _
                    vsSourceServerNCID = NCID_ENCOUNTER_APPLICATION Then
                    If frmAppointmentView.Visible = True Then
                    RefreshData
                    End If
                End If
            Else
                If vsSourceServerNCID = NCID_PROBLEMS_APPLICATION Then Exit Sub 'SCR-34028
                RefreshData
            End If
            
        '<< SCR 0768
        Case cwiOK_TO_CHANGE_PATIENT
            cwobjCWComm.Message cwiOK_TO_CHANGE_PATIENT, "Y", vsSourceServerNCID, NCID_CLINIC_SCHEDULE_APPLICATION
            
        Case cwiOK_TO_CLOSE_VIEW
            cwobjCWComm.Message cwiOK_TO_CLOSE_VIEW, "Y", vsSourceServerNCID, NCID_CLINIC_SCHEDULE_APPLICATION
        
        Case cwiGENERIC
            If InStr(1, vsMessageText, "UPDATE_MENU_ITEMS", vbTextCompare) Then
                'As of 12/30/03 this is only called by and used for theater menus.
                UpdateMenuItems (vsMessageText)
            Else
                'SCR 31497
                Select Case UCase(vsMessageText)
                    Case "CHCS_GUI_APPT_GRIDINFO", "APPT_GRIDINFO"
                        vsMessageText = SendBackGridInfo()
                    Case "UPDATECLINICS"
                        UpdateMyClinics (vsSourceServerNCID)
                    ' 78339 Quan
                    Case "CREATE_APPT"
                        frmAppointmentView.NewAppointment
                    Case Else
                        UpdateAppointment vsMessageText
                End Select
            End If
        Case cwiUPDATE_ADM_STATUS
            'vsMessagetext in the format of "ApptIEN|Complete" "or ApptIEN|"
            Update_Encounter_Status_CHCS_GUI vsMessageText
        
        Case cwiCHECKIN_APPOINTMENT
            CheckInAppointment_CHCS_GUI vsMessageText
                
        End Select

    Exit Sub
    
ErrHandler:
    gobjShared.ShowVBError Err.Number, Err.Description + ", viMessageType = " + CStr(viMessageType) + ", vsMessageText = " + vsMessageText, "CHCSII_Appointments.Message", "CHCSII_Appointments", vbExclamation
    
End Sub
Private Sub UpdateMenuItems(vsMessageText As Variant)
Dim sMsg As Variant

    sMsg = Split(vsMessageText, "|")
    If gobjShared.IsAppMode(modeTheater) Then 'SCR-47103
        'vMessageText will come in the format of: "UPDATE_MENU_ITEMS|ApptType|boolean"
        'ApptType will be either INPT or not, the boolean will represent the value of
        'GEMS_InpatConsultsMenusDisabled which is private on frmAppointmentView.
        If UBound(sMsg) <> 2 Then Exit Sub
    
        SetInpatientConsultMenuItems_GEMS CStr(sMsg(1)), CBool(sMsg(2))
    End If
    'There is another SCR to work later in which we need to update menu items in other modes.

End Sub

Private Sub SetInpatientConsultMenuItems_GEMS(sApptType As String, GEMS_InpatConsultsMenusDisabled As Boolean)  'SCR-47103
    If Not gobjShared.IsAppMode(modeTheater) Then Exit Sub
    If sApptType = vbNullString Then Exit Sub

    If sApptType = "INPT" Or sApptType = "INPATIENT CONSULTS" Then
        If GEMS_InpatConsultsMenusDisabled = False Then 'Selected MenuItems are Enabled
            DisableInpatientMenuItems_GEMS
        End If
    Else
        If GEMS_InpatConsultsMenusDisabled = True Then  'Selected MenuItems are Disabled
            EnableInpatientMenuItems_GEMS
        End If
    End If
End Sub
Private Sub DisableInpatientMenuItems_GEMS() 'SCR-47103
    With cwobjCWComm
        .UpdateMenuItem goCoreMenuID(apptMnuActCancelMenuId), "DISABLE"
        .UpdateMenuItem goCoreMenuID(apptMnuActTransferMenuId), "DISABLE"
        .UpdateMenuItem goCoreMenuID(apptMnuActCheckinMenuId), "DISABLE"
        .UpdateMenuItem goCoreMenuID(apptMnuActCheckoutMenuId), "DISABLE"
        .UpdateMenuItem goCoreMenuID(apptMnuActAddProvider), "DISABLE"
        .UpdateMenuItem goCoreMenuID(apptMnuActComments), "DISABLE"
        .UpdateMenuItem goCoreMenuID(apptMnuActUndoCancel), "DISABLE"
        'GEMS_InpatConsultsMenusDisabled = True
    End With
End Sub
Private Sub EnableInpatientMenuItems_GEMS() 'SCR-47103
    With cwobjCWComm
        .UpdateMenuItem goCoreMenuID(apptMnuActCancelMenuId), "ENABLE"
        .UpdateMenuItem goCoreMenuID(apptMnuActTransferMenuId), "ENABLE"
        .UpdateMenuItem goCoreMenuID(apptMnuActCheckinMenuId), "ENABLE"
        .UpdateMenuItem goCoreMenuID(apptMnuActCheckoutMenuId), "ENABLE"
        .UpdateMenuItem goCoreMenuID(apptMnuActAddProvider), "ENABLE"
        .UpdateMenuItem goCoreMenuID(apptMnuActComments), "ENABLE"
        .UpdateMenuItem goCoreMenuID(apptMnuActUndoCancel), "ENABLE"
        'GEMS_InpatConsultsMenusDisabled = False
    End With

End Sub



Private Sub UpdateCHCS_GUI_ADMStatus(vsMessageText As Variant)
Dim sApptID As String
Dim sStatus As String
Dim nRow As Integer
    
Dim sMsg As Variant

    sMsg = Split(vsMessageText, "|")
    If Not gobjShared.IsAppMode(modeCHCSI_GUI) Then GoTo ErrorMessage
    If UBound(sMsg) < 1 Then GoTo ErrorMessage
    sApptID = sMsg(0)
    sStatus = sMsg(1)
    'Update the appt collection  & 'Update the specific grid item.
    Update_Encounter_Status_CHCS_GUI (vsMessageText)
        
Exit Sub
ErrorMessage:
    MsgBxARMd "The Appointment Encounter status could not be updated: " + CStr(vsMessageText)

End Sub
Private Function TranslateEncStatus(rnStatus As CHCSII_AppointmentStatus.EncounterStatusEnum) As String
    Select Case rnStatus
    Case EncCheckedIn
        TranslateEncStatus = "CheckedIn"
    Case EncWaiting
        TranslateEncStatus = "Waiting"
    Case EncInProgress
        TranslateEncStatus = "InProgress"
    Case EncUpdated
        TranslateEncStatus = "Updated"
    Case EncNeedsCosignature
        TranslateEncStatus = "NeedsCoSignature"
    Case EncCompleted
        TranslateEncStatus = "Complete"
    Case EncUpdating
        TranslateEncStatus = "Updating"
    Case Else
        TranslateEncStatus = "Unknown"
    End Select
End Function
Private Sub Update_Encounter_Status_CHCS_GUI(MessageText As Variant)
Dim sApptID As String
Dim sStatus As String
Dim nRow As Integer
Dim sMsg As Variant
Dim eNewEncStatus As CHCSII_AppointmentStatus.EncounterStatusEnum
On Error GoTo ErrHandler  'SCR-40435

'MessageText consists of apptID | EncounterStatus
'Encounter status can be numeric or text.

    sMsg = Split(MessageText, "|")
    If UBound(sMsg) < 1 Then Exit Sub
    sApptID = sMsg(0)
    If IsNumeric(sApptID) = False Then Exit Sub
    sStatus = sMsg(1)
    'If IsNumeric(sStatus) = False Then Exit Sub
    
    'Update the specific grid item.
    If frmAppointmentView Is Nothing Then Exit Sub
    If frmAppointmentView.Visible = False Then Exit Sub
    With frmAppointmentView
        'Search the appt grid to see if in fact the row is listed.
        For nRow = 1 To .grdAppt.Rows - 1
            If .grdAppt.TextMatrix(nRow, gnAPPT_IEN_HEADER) = sMsg(0) Then
                If IsNumeric(sStatus) = True Then
                    eNewEncStatus = CInt(sStatus)
                    sStatus = TranslateEncStatus(eNewEncStatus)
                End If
                If UCase(sStatus) = "COMPLETE" Then
                    .grdAppt.TextMatrix(nRow, gnENCOUNTER_HEADER) = "Complete"
                Else
                    .grdAppt.TextMatrix(nRow, gnENCOUNTER_HEADER) = ""
                End If
                'Now Update the selected appt ADMstatus in our collection of appts.
                gobjAppointmentOps.UpdateADMStatus_CHCS_GUI sApptID, sStatus
                Exit For
            End If
        Next
    End With
    
Exit Sub
ErrHandler:
    gobjShared.ShowVBError Err.Number, Err.Description, "CHCSII_Appointments.Update_Encounter_Status_CHCS_GUI", "CHCSII_Appointments", vbExclamation
    

End Sub

Private Sub CheckInAppointment_CHCS_GUI(MessageText As Variant)
Dim sApptID As String
Dim nRow As Integer


On Error GoTo ErrHandler
'SCR-46153
'MessageText should consist of apptIEN only
    
    If Not gobjShared.IsAppMode(modeCHCSI_GUI) Then Exit Sub
    If MessageText = vbNull Then Exit Sub
    sApptID = MessageText
    If IsNumeric(sApptID) = False Then Exit Sub
    
    If frmAppointmentView Is Nothing Then Exit Sub
    If frmAppointmentView.Visible = False Then Exit Sub
    With frmAppointmentView
        'Search the appt grid to see if in fact the row is listed.
        For nRow = 1 To .grdAppt.Rows - 1
            If .grdAppt.TextMatrix(nRow, gnAPPT_IEN_HEADER) = sApptID Then
                frmAppointmentView.CheckInPGUIAppointment (nRow)
                Exit For
            End If
        Next
    End With
    
Exit Sub
ErrHandler:
    gobjShared.ShowVBError Err.Number, Err.Description, "CHCSII_Appointments.CheckInAppointment_CHCS_GUI", "CHCSII_Appointments", vbExclamation
    

End Sub


Private Function SendBackGridInfo() As String
'Function called by encounters via disposition to retrieve grid info.
    Dim sTemp As String
    Dim nRow  As Long
    Dim sClinicNCID As String
    Dim sInpatient_MEPRS_Code As String
    Dim enumEncounterType As CHCSIIEncounterOps.EncType
    Dim objAppointment As CHCSII_AppointmentClient.Appointment
    Dim sClinicianNCID As String
    
    'SCR 31497
    On Error GoTo ErrHandler
    SendBackGridInfo = ""

        If frmAppointmentView Is Nothing Then GoTo Cleanup

        If frmAppointmentView.grdAppt.SelectedRows = 1 And frmAppointmentView.grdAppt.IsSelected(0) = False Then
            nRow = frmAppointmentView.grdAppt.RowSel
            
            'SCR 34581  4/23/03  Phil Crowder
            'SCR-46153 commented this out and created specific message to checkin appt.
'            If gobjShared.IsAppMode(modeCHCSI_GUI) Then
'              bSuccess = frmAppointmentView.CheckInPGUIAppointment(nRow)
'              'If Not bSuccess Then GoTo Cleanup  'SCR-40989
'            End If
           
            sTemp = frmAppointmentView.grdAppt.TextMatrix(nRow, gnAPPT_IEN_HEADER) + "|" + _
                frmAppointmentView.grdAppt.TextMatrix(nRow, gnTYPE_HEADER) + "|" + _
                frmAppointmentView.grdAppt.TextMatrix(nRow, gnTIME_HEADER) + "|" + _
                frmAppointmentView.grdAppt.TextMatrix(nRow, gnCLINIC_HEADER)
                
                sClinicNCID = "0"
                If frmAppointmentView.mcolAppointments Is Nothing Then
                     sClinicNCID = "0"
                Else
                    If frmAppointmentView.mcolAppointments.Count = 0 Then
                        sClinicNCID = "0"
                    Else
                        On Error Resume Next
                        
                        Set objAppointment = frmAppointmentView.mcolAppointments.Item(frmAppointmentView.grdAppt.TextMatrix(nRow, gnAPPT_ID_HEADER))
                        If objAppointment Is Nothing Then
                            MsgBxARMd "Could not retrieve the appointment data." + vbCrLf + "Appointments.cls - SendBackGridInfo", , "Application Error"
                            GoTo Cleanup
                        End If
                        If gobjShared.IsAppMode(modeCHCSI_GUI) Then 'SCR-43291
                            RetrieveNCIDS_CHCS_GUI objAppointment, sClinicNCID, sClinicianNCID
                        Else
                            sClinicNCID = objAppointment.ClinicNCID
                        End If
                        
                        If gobjShared.IsAppMode(modeCHCSI_GUI) Then
                            sInpatient_MEPRS_Code = objAppointment.MeprsCode
                        Else
                            sInpatient_MEPRS_Code = objAppointment.Inpatient_MEPRS_Code
                        End If
                        
                    End If
                End If
                sTemp = sTemp + "|" + sClinicNCID
                sTemp = sTemp + "|" + sInpatient_MEPRS_Code
                If objAppointment.ApptClassEnum = ApptOutpatient Then
                    enumEncounterType = Outpatient
                ElseIf objAppointment.ApptClassEnum = ApptInpatient Then
                    enumEncounterType = Inpatient
                Else
                    enumEncounterType = Outpatient
                End If
                sTemp = sTemp + "|" + CStr(enumEncounterType)
                sTemp = sTemp + "|" + objAppointment.Workload_Type
                sTemp = sTemp + "|" + sClinicianNCID  'SCR-43291
                
                SendBackGridInfo = sTemp
        End If
Cleanup:
    Set objAppointment = Nothing
Exit Function
ErrHandler:
    gobjShared.ShowVBError Err.Number, Err.Description, "CHCSII_Appointments.SendBackGridInfo", "CHCSII_Appointments", vbExclamation
    Set objAppointment = Nothing
End Function

Private Sub RefreshData()

  On Error GoTo ErrHandler
  
  If cwbMainViewRunning Then
      frmAppointmentView.VerifySelectedPatient
      frmAppointmentView.RefreshData
  End If
  
  Exit Sub
    
ErrHandler:
  gobjShared.ShowVBError Err.Number, Err.Description, "CHCSII_Appointments.RefreshData", "CHCSII_Appointments", vbExclamation

End Sub

Private Sub DisplayAbout()
Dim objAbout As New CWAbout
            
  cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, cwlMyProgId
  
  objAbout.Display App, frmAppointmentView
  Set objAbout = Nothing
  
  cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, cwlMyProgId

End Sub

Private Sub UpdateAppointment(ByVal MessageText As String)
Dim oAppt As CHCSII_AppointmentClient.Appointment
Dim sMsg As Variant
Dim intAppClass As AppointmentClassifications
Dim strOk As String
On Error GoTo ErrHandler
  '<< SCR 10812 update only the single row in appt screen
  sMsg = Split(MessageText, "|")
  
  If UBound(sMsg) < 1 Then Exit Sub
  
  Select Case UCase$(sMsg(0))
    Case "UPDATE_APPT_CLASSIFICATION"
      If UBound(sMsg) < 2 Then Exit Sub
      Select Case UCase$(sMsg(2))
        Case "IN-PATIENT"
          intAppClass = ApptInpatient
        Case "OUT-PATIENT"
          intAppClass = ApptOutpatient
        Case "AMBULATORY"
          intAppClass = ApptAmbulatoryProcedure
        Case "TELECONS"
          intAppClass = ApptTelCon
        Case Else
          intAppClass = ApptInpatient
      End Select
      
      gobjAppointmentOps.UpdateAppointmentClassification CStr(sMsg(1)), intAppClass
      
    Case "UPDATE_APPT_RFV"
      'SCR 24097  Phil Crowder 6/4/02
      If UBound(sMsg) < 2 Then Exit Sub
      gobjAppointmentOps.UpdateAppointmentRFV CStr(sMsg(1)), CStr(sMsg(2))
      
    Case "UPDATE_APPT_WORKLOADTYPE"
        If UBound(sMsg) < 2 Then Exit Sub
       gobjAppointmentOps.UpdateAppointmentWorkload CStr(sMsg(1)), CStr(sMsg(2))
       
    Case "UPDATE_ENCOUNTER_STATUS"
        If gobjShared.IsAppMode(modeCHCSI_GUI) Then
            GoTo Cleanup  'SCR-40435
        End If
        If gobjShared.IsAppMode(modeCDR) Then
            If UBound(sMsg) < 2 Then Exit Sub
            If CStr(sMsg(2)) = "8" Then GoTo Cleanup 'SCR-42716
        End If
        
    Case "NEW_APPOINTMENT" 'SCR-42772
        If Not gobjShared.IsAppMode(modeTheater) Then Exit Sub
        DoEvents
        If Not gdlgGEMS_NewApptArmy Is Nothing Then
            If gdlgGEMS_NewApptArmy.Visible = True Then Exit Sub            'NewVisit is already displayed.
        Else
            Set gdlgGEMS_NewApptArmy = New dlgGEMS_NewApptArmy
        End If
'        'SCR-52771 Commented out... Check done by CAC Options.
'        Dim sApptDateFound As String
'        If PatientHasOpenAppts(sApptDateFound) Then
'            MsgBxARMd "The selected patient currently has an open appointment on " + sApptDateFound + "."
'            GoTo SetBackToDefaults
'        End If
        
        'SCR-52771
        If CheckForPrivileges_GEMS() Then
            goARM.StartFarTx gsARMTx_OpenPatSrchFromAppts
            cwobjCWComm.Message cwiOK_TO_CHANGE_PATIENT, strOk, NCID_CW_APPLICATION, NCID_CW_APPLICATION, False
            If strOk <> "N" Then
                gbNotCalledFromAppointmentView = True
                gbShowPatientSearchForm = False
                'Call cwobjCWComm.Message(cwiSTART_OLE_SERVER, NCID_CLINIC_SCHEDULE_APPLICATION, NCID_CW_APPLICATION, NCID_CLINIC_SCHEDULE_APPLICATION)
                'SCR-50896
                cwbFirstRegularResize = False
                frmAppointmentView.Display
                frmAppointmentView.NewAppointment_GEMS  'do NOT auto open the Patient Search form
            End If


        End If
SetBackToDefaults:        'Set back to defaults
        gbNotCalledFromAppointmentView = False
        gbShowPatientSearchForm = True
        gbNeedPatientSelect = True
        GoTo Cleanup
        
        
  End Select
 
  Set oAppt = gobjAppointmentOps.GetAppointment(CStr(sMsg(1)), False)
  frmAppointmentView.FillOneRow oAppt, frmAppointmentView.FindApptRow(CStr(sMsg(1)))
Cleanup:
  Set oAppt = Nothing

  Exit Sub
    
ErrHandler:
  gobjShared.ShowVBError Err.Number, Err.Description, "CHCSII_Appointments.UpdateAppointment", "CHCSII_Appointments", vbExclamation

End Sub
Private Function CheckForPrivileges_GEMS() As Boolean 'SCR-52771
On Error GoTo ErrHandler
    Dim CanSign As CHCSII_PrivilegeType
    Dim NeedCosign As CHCSII_PrivilegeType

        If gobjUser.HasPrivilegeEx(Priv_Appointments, Priv_Read) And _
            gobjUser.HasPrivilegeEx(Priv_Appointments, Priv_Write) And _
            gobjUser.HasPrivilegeEx(Priv_Appointments, Priv_Delete) Then
                
            CheckForPrivileges_GEMS = True
            
            '--- SCR 57613   Sherry Wang   2/16/2005
            If cwobjPatient.CACpatient = False Then
                Exit Function
            Else
                cwobjPatient.CACpatient = False
            End If
            
            If LenB(gobjShared.CmdLineSwitches.Item("CAC")) > 0 Then
                If gobjUser.HasPrivilegeEx(Priv_Current_Encounter, Priv_Sign) Or _
                     gobjUser.HasPrivilegeEx(Priv_Current_Encounter, Priv_Need_Cosign) Then
                             
                    CheckForPrivileges_GEMS = True
                Else
                    CheckForPrivileges_GEMS = False
                    MsgBxARMd "The user " + gobjUser.UserName + " does not have the security " + _
                    "privilege to access appointments. Current Encounter Sign or Need Cosign required."
                    Exit Function
                End If
            End If
        Else
            MsgBxARMd "The user " + gobjUser.UserName + " does not have the security " + _
                "privilege to access appointments.  Appointments Read/Write/Deleted required."
        
        End If

Exit Function
ErrHandler:
  gobjShared.ShowVBError Err.Number, Err.Description, "CHCSII_Appointments.CheckForPrivileges_GEMS", "CHCSII_Appointments", vbExclamation

End Function
'Private Function PatientHasOpenAppts(sApptDateFound As String) As Boolean  'SCR-52771
''Returns True if patient has open appointments.
''Set sApptDateFound to the first open appt found.
'Dim cTemp As Collection
'On Error GoTo ErrHandler
'
'    Set cTemp = gobjAppointmentOps.GetAllAppointmentsForPatient(123, "INCOMPLETE")
'    If Not cTemp Is Nothing Then
'        If cTemp.Count > 0 Then
'            sApptDateFound = CStr(cTemp.Item(1).ApptDateTime)
'            PatientHasOpenAppts = True
'        End If
'    End If
'    Set cTemp = Nothing
'Exit Function
'ErrHandler:
'  gobjShared.ShowVBError Err.Number, Err.Description, "CHCSII_Appointments.PatientHasOpenAppts", "CHCSII_Appointments", vbExclamation
'  Set cTemp = Nothing
'End Function
Public Sub CloseView(ByVal viViewType As Integer)
Dim intCount As Integer

    On Error Resume Next

    Select Case viViewType
        Case cwiALL_VIEWS
            cwbMainViewRunning = False
            cwbCloseViewMyself = False
            cwbPatientSummaryViewRunning = False
            
            For intCount = 0 To Forms.Count
              Unload Forms(0)
            Next
            
            DeleteMenus
            
            Set goCoreMenuID = Nothing
            Set gfrmHealthHistAppts = Nothing
        
        Case cwiREGULAR_VIEW
            cwbMainViewRunning = False
            cwbCloseViewMyself = False
        
            For intCount = 0 To Forms.Count
              If Forms(0).Name <> gfrmHealthHistAppts.Name Then Unload Forms(0)
            Next
            
            DeleteMenus
            Set goCoreMenuID = Nothing
        
        Case cwiPATIENT_SUMMARY_VIEW
            cwbPatientSummaryViewRunning = False
            
            Unload gfrmHealthHistAppts
            Set gfrmHealthHistAppts = Nothing
    End Select
    
End Sub

Private Sub DeleteMenus()
  On Error Resume Next
  
  '--- SCR-18843 MSolano 11/16/01
  '--- Reconciliation of several variables tracking appt selection criteria:
  '---  goCoreMenuID maps value, key pair
  '---  between Appointments menu actions & CORE menus
  
  With cwobjCWComm
    '--- DELETE CORE ACTION MENUS OTHER THAN THOSE SETTING APPT
    '--- SELECTION CRITERIA (type AppointmentOps.AppointmentMenuActions)
    .DeleteMenuItem goCoreMenuID(apptMnuActRefreshMenuId)
    .DeleteMenuItem goCoreMenuID(apptMnuActNewMenuId)
    .DeleteMenuItem goCoreMenuID(apptMnuActOpenMenuId) 'SCR-31309
    .DeleteMenuItem goCoreMenuID(apptMnuActCheckinMenuId)
    .DeleteMenuItem goCoreMenuID(apptMnuActCancelMenuId)
    .DeleteMenuItem goCoreMenuID(apptMnuActCheckoutMenuId)
    .DeleteMenuItem goCoreMenuID(apptMnuActTransferMenuId)
    .DeleteMenuItem goCoreMenuID(apptMnuActAddProvider)
    .DeleteMenuItem goCoreMenuID(apptMnuActPrint)
    .DeleteMenuItem goCoreMenuID(apptMnuActComments)
    .DeleteMenuItem goCoreMenuID(apptMnuActPrintInsuranceForm)
    .DeleteMenuItem goCoreMenuID(apptMnuActUndoCancel)
    .DeleteMenuItem goCoreMenuID(apptDefault)
    
    '--- DELETE CORE ACTION MENUS THAT SET APPT SELECTION
    '--- CRITERIA (type AppointmentOps.AppointmentFilters)
    .DeleteMenuItem goCoreMenuID(apptToday)
    .DeleteMenuItem goCoreMenuID(apptTodayPlusIncomplete)
    .DeleteMenuItem goCoreMenuID(apptAll)
    .DeleteMenuItem goCoreMenuID(apptPending)
    .DeleteMenuItem goCoreMenuID(apptCheckedIn)
    .DeleteMenuItem goCoreMenuID(apptWaiting)
    .DeleteMenuItem goCoreMenuID(apptInProgress)
    .DeleteMenuItem goCoreMenuID(apptNeedsCoSignature)
    .DeleteMenuItem goCoreMenuID(apptUpdated)
    .DeleteMenuItem goCoreMenuID(apptCompleted)
    .DeleteMenuItem goCoreMenuID(apptIncomplete)
    .DeleteMenuItem goCoreMenuID(apptCheckedInPending)      'SCR 7454 - Phil Crowder 7/9/01
  End With
    'giLastViewMenuID = 0 '--- SCR-18843 MSolano 11/16/01
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    
    Set cwobjCWComm = Nothing
    Set cwobjLogonEnviron = Nothing
    Set cwobjConfig = Nothing
    Set cwobjPatient = Nothing
    Set gobjShared = Nothing
    Set gobjAppointmentOps = Nothing
    Set gobjClinicOps = Nothing
    Set goCoreMenuID = Nothing   '--- SCR-18843 MSolano 11/16/01
    
    Set mDTOsMyClinics = Nothing
    Set mcolClinicProviders = Nothing
End Sub

Public Sub Navigate()
    Dim Index As Integer
            
    On Error GoTo ErrHandler
    
    With gobjUser
        For Index = 1 To .Resources.Count
            If .Roles.Item(Index).Name = "physician" Then
                cwobjCWComm.Message cwiSTART_PATIENT_SUMMARY, "", NCID_PATIENT_SUMMARY_APPLICATION, cwlMyProgId
                Exit For
            ElseIf .Roles.Item(Index).Name = "nurse" Then
                cwobjCWComm.Message cwiSTART_OLE_SERVER, "", NCID_REASON_FOR_VISIT, cwlMyProgId
                Exit For
            End If
        Next
    End With
    
    Exit Sub
    
ErrHandler:
    gobjShared.ShowVBError Err.Number, Err.Description, "CHCSII_Appointments.Navigate", "CHCSII_Appointments", vbExclamation
    
End Sub

'DEM 01/03/2000
Public Function DisplayCreateNewAppt(frmParent As Object, sAppModuleNCID As String, sNoteAuthor As String, sEncounterNote As String) As Boolean
    
  On Error GoTo ErrHandler
  
    cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, sAppModuleNCID
    gbNotCalledFromAppointmentView = True
    
    LoadClinicsANDProviders
    
    'SCR-52771
    If Not gobjShared.IsAppMode(modeTheater) Then
        If frmAppointmentView.mbWorkload_Data_Loaded = False Then
            If Not frmAppointmentView.GetWorkLoadData Then
                cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, sAppModuleNCID
                Exit Function
            End If
        End If
    End If
    If Not (frmAppointmentView.Display()) Then
        gbNotCalledFromAppointmentView = False
        cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, sAppModuleNCID
        Exit Function
    End If
    

    
    If gobjShared.IsAppMode(modeCDR) And LenB(gobjShared.CmdLineSwitches.Item("WAM")) > 0 Then
        Load dlgNewVisit
        If sNoteAuthor <> "" Then
            dlgNewVisit.NoteAuthor = sNoteAuthor
        End If
        If sEncounterNote <> "" Then
            dlgNewVisit.EncounterNote = sEncounterNote
        End If
        dlgNewVisit.cmdChangePatient.Enabled = False
        dlgNewVisit.NeedPatientSelect = False
        dlgNewVisit.Show vbModal, frmParent
        'SCR-39578
        'gbNotCalledFromAppointmentView = False
        cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, sAppModuleNCID
        gbNotCalledFromAppointmentView = False
        'SCR-39578
        'If dlgNewVisit Is Nothing Then Exit Function
        'Unload dlgNewVisit
        DisplayCreateNewAppt = True
        
    Else
        Load dlgNewAppt
        If sNoteAuthor <> "" Then
            dlgNewAppt.NoteAuthor = sNoteAuthor
        End If
        If sEncounterNote <> "" Then
            dlgNewAppt.EncounterNote = sEncounterNote
        End If
        dlgNewAppt.cmdChangePatient.Enabled = False
        dlgNewAppt.NeedPatientSelect = False
        dlgNewAppt.Show vbModal, frmParent
        gbNotCalledFromAppointmentView = False
        cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, sAppModuleNCID
        gbNotCalledFromAppointmentView = False
        DisplayCreateNewAppt = True
'        frmAppointmentView.RefreshData
        
    End If
  
Exit Function
ErrHandler:
    gbNotCalledFromAppointmentView = False
    gobjShared.ShowVBError Err.Number, Err.Description, "CHCSII_Appointments.DisplayCreateNewAppt", "CHCSII_Appointments", vbExclamation
  
  
  
        
End Function
Public Function AppointmentOps() As CHCSII_AppointmentClient.AppointmentOps
    
    Set AppointmentOps = gobjAppointmentOps
End Function
Public Function DisplayCheckOutAppt(frmParent As Object, sAppModuleNCID As String, objAppointment As CHCSII_AppointmentClient.Appointment) As Boolean
'SCR-23202
  On Error GoTo ErrHandler

    If objAppointment.CheckOutTime = 0 Or objAppointment.CheckOutTime > #12/31/2100# Or objAppointment.CheckOutTime < #12/31/1990# Then
        gbNotCalledFromAppointmentView = True
        cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, sAppModuleNCID
        Load dlgCheckout
        dlgCheckout.Init objAppointment
        dlgCheckout.Show vbModal, frmParent
        cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, sAppModuleNCID
        gbNotCalledFromAppointmentView = False
    
    Else
        MsgBxARMd "This appointment was already checked out on " + Format(objAppointment.CheckOutTime, "dd MMM yyyy HHnn") + ". " + _
                "You can not check out the appointment again."
    End If
    
    DisplayCheckOutAppt = True
    Exit Function
ErrHandler:
    gobjShared.ShowVBError Err.Number, Err.Description, "CHCSII_Appointments.DisplayCheckOutAppt", "CHCSII_Appointments", vbExclamation
    DisplayCheckOutAppt = False
    
End Function
Public Function DisplayCancelAppt(frmParent As Object, sAppModuleNCID As String, objAppointment As CHCSII_AppointmentClient.Appointment) As Boolean
Dim bAllOK As Boolean
  On Error GoTo ErrHandler
    
    If objAppointment Is Nothing Then
        MsgBxARMd "Invalid Appointment reference set to Appointment.DisplayCancelAppt"
        GoTo Cleanup
    End If
    'SCR-52771
    If Not gobjShared.IsAppMode(modeCDR) And Not gobjShared.IsAppMode(modeTheater) Then
        GoTo Cleanup
    End If
    If LenB(gobjShared.CmdLineSwitches.Item("CAC")) = 0 Then
        MsgBxARMd "Cancelling appointments from outside of the appointments module is only " + _
            "available with the CAC interface."
        GoTo Cleanup
    End If
    
    cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, sAppModuleNCID
    gbNotCalledFromAppointmentView = True
    Load dlgCancel
    
    'CAC Cancel Appointment Requirement
    
    Set dlgCancel.CAC_ObjAppointment = objAppointment
    If Not dlgCancel.ValidateAppointmentCancelfromID(objAppointment) Then
        bAllOK = False
        GoTo Cleanup
    End If

    If Not dlgCancel.SetCancelOptions Then
        bAllOK = False
        GoTo Cleanup
    End If
    
    dlgCancel.Show vbModal, frmParent
    DisplayCancelAppt = True
Cleanup:
    
    gbNotCalledFromAppointmentView = False
    cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, sAppModuleNCID
    Exit Function
ErrHandler:
    gobjShared.ShowVBError Err.Number, Err.Description, "CHCSII_Appointments.DisplayCancelAppt", "CHCSII_Appointments", vbExclamation
    DisplayCancelAppt = False
    GoTo Cleanup
End Function
Private Sub UpdateMyClinics(ByVal vsSourceServerNCID As String)
Dim objDTOsClinic As DTOs.Clinic
On Error GoTo ErrHandler
    
    If frmAppointmentView Is Nothing Then
        Exit Sub
    End If
    Set mDTOsMyClinics = gobjClinicOps.GetUserClinics(cwobjLogonEnviron.FacilityNCID, cwobjLogonEnviron.UserNCID)
    
    If mDTOsMyClinics.Count < 1 Then
        cwobjCWComm.Message cwiOK_TO_CLOSE_VIEW, "Y", vsSourceServerNCID, NCID_CLINIC_SCHEDULE_APPLICATION
        Exit Sub
    End If
    Set mcolClinicProviders = New Collection
    For Each objDTOsClinic In mDTOsMyClinics
        mcolClinicProviders.Add gobjClinicOps.GetClinicProviders(objDTOsClinic.NCID), objDTOsClinic.NCID
    Next objDTOsClinic
    Set objDTOsClinic = Nothing
Exit Sub
ErrHandler:
    gobjShared.ShowVBError Err.Number, Err.Description, "CHCSII_Appointments.UpdateMyClinics", "CHCSII_Appointments", vbExclamation
    Set objDTOsClinic = Nothing

End Sub
