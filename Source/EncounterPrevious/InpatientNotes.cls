VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InPatientNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements ICHCSII_CoreAPI2             '-- implements the security obj

Private mobjUser As ICHCSII_User        '-- The Security User Object
Private mcolEncChild As Collection      '-- encounter child modules
Private mobjPatient As Patient          '-- patient object from core
Private mobjConfig As Config            '-- config obj from core
Private mobjLogon As Logon              '-- logon env obj from core
Private mobjComm As Comm                '-- comm obj from core
Private mlProgID As Long                '-- my prog id from core
Private mbMainViewRunning As Boolean
Private mbFirstRegularResize As Boolean
Private mbFirstSummaryResize As Boolean
Private mobjShared As CWShared
Private mlMenuAddNote As Long
Private mlMenuAmend As Long
Private mlMenuSign As Long
Private mfrmRegularView As frmEncPrevious
Private mfrmSummaryView As frmEncPrevious
Private mbButtonClicked As Boolean

'-- implements the dlog.dll timing module
#If debugon Then
  Private DR As DebugRec
  Private DL As cDebugLog
  Private Const Exe = "PrevEnc"
  Private Const Module = "PrevEnc.cls"
#End If

Private mobjEncounters As CHCSIIEncounterPrevious.Encounters '-- previous encounters

'=========== End of standard property exposure code ===========================

'*******************************************************************************************
'  Description: This method is called when a client wants to know if this OLE Server
'               supports a specific "view type".
'
'  Parameters:
'    viViewType - Contains the "type" of view being queried for.
'
'  Return Value:
'    Returns TRUE if this OLE Server supports the passed in "view type".
'    Returns FALSE if this OLE Server does not support the passed in "view type".
'*******************************************************************************************
Public Function QueryView(ByVal viViewType As Integer) As Boolean
    Select Case viViewType
    Case cwiREGULAR_VIEW
        QueryView = True    '--this view is NOT optional--
    Case cwiPATIENT_SUMMARY_VIEW
        QueryView = True
    Case cwiCLINICIAN_SUMMARY_VIEW
        QueryView = False
    Case cwiENCOUNTER_SUMMARY_VIEW
        QueryView = False
    Case cwiPROPERTY_VIEW
        QueryView = False
    Case Else
        QueryView = False
    End Select
End Function

'**************************************************************************************************
'  Description: This method is called when the CW Core attempts to open a "view" provided by this
'               OLE server.
'
'  Parameters:
'    viViewType - Contains the "type" of view being requested by the CW Core to open.
'    vlParenthWnd - Contains the hWnd of the client area window to be used by this OLE Server.
'    rhMainForm - If "viViewType" is a "regular view", upon return this will contain
'                 the window "handle" of the main "regular" form for this OLE Server.
'
'  Return Value:
'    Returns TRUE if the "view" was opened successfully.
'    Returns FALSE if an error occurred openning the "view".
'
'**************************************************************************************************
Public Function OpenView(ByVal viViewType As Integer, _
                         ByVal vlParenthWnd As Long, _
                               rhMainForm As Long) As Boolean
    Dim i As Integer
    Dim msg As String
    On Error GoTo ErrHandler
        
    
    Select Case viViewType
    Case cwiREGULAR_VIEW
        mbFirstRegularResize = True
        mbMainViewRunning = True
        Set mfrmRegularView = New frmEncPrevious
        
        mfrmRegularView.Inpatient = True
        mfrmRegularView.lblEnc(1).Caption = "Inpatient Notes"
        mfrmRegularView.lblEnc(0).Caption = Space(24) & "Inpatient Notes:"
        Set mfrmRegularView.ParentObject = Me
        
        SetParent mfrmRegularView.hwnd, vlParenthWnd
        
        Call AddMenus
        
        rhMainForm = mfrmRegularView.hwnd
        OpenView = True
        
        Call mfrmRegularView.RefreshDisplay(GetInpatientNotes(True, False))
        
    Case cwiPATIENT_SUMMARY_VIEW
        mbFirstSummaryResize = True
        Set mfrmSummaryView = New frmEncPrevious
        mfrmSummaryView.SummaryView = True
        mfrmSummaryView.Inpatient = True
        mfrmSummaryView.InitListView
        ' SCR #46974;   Developer:  12/22/2003
        mfrmSummaryView.lblEnc(1).Caption = "Inpatient Notes"
        mfrmSummaryView.lblEnc(0).Caption = Space(12) & "Signed Inpatient Notes:"
        
        SetParent mfrmSummaryView.hwnd, vlParenthWnd
        rhMainForm = mfrmSummaryView.hwnd
        OpenView = True
        
        Call mfrmSummaryView.RefreshDisplay(GetInpatientNotes(True, True))
    Case Else    '--unsupported view "type"--
        OpenView = False
        goARM.StopFarTx NCID_INPATIENT_NOTES_APPLICATION & gsARM_ACT_OPEN, enmARMTxStatusFailed  '--- SCR-ARM-19545 MSolano 12/21/01
    End Select
    Exit Function

ErrHandler:
    goARM.StopFarTx NCID_INPATIENT_NOTES_APPLICATION & gsARM_ACT_OPEN, enmARMTxStatusAborted '--- SCR-ARM-19545 MSolano 12/21/01
    Call mobjShared.ShowVBError(Err.Number, Err.Description, "InPatientNotes.OpenView", "Enc", vbCritical)
    OpenView = False
End Function

Private Function GetInpatientNotes(ByVal bRefresh As Boolean, ByVal Lite As Boolean) As CHCSIIEncounterPrevious.Encounters
    
    If bRefresh Or (mobjEncounters Is Nothing) Then
        Set mobjEncounters = New CHCSIIEncounterPrevious.Encounters
        mobjEncounters.GetInpatientNotes mobjPatient.UnitNumber, Lite
    End If
    
    Set GetInpatientNotes = mobjEncounters
    
End Function
'*******************************************************************************************
'  Description: This method is called when the CW Core closes a "view" provided by this OLE
'               Server.
'
'  Parameters:
'    viViewType - Contains the "type" of view that the CW Core is closing.
'*******************************************************************************************
Public Sub CloseView(ByVal viViewType As Integer)
    On Error Resume Next
    
    Select Case viViewType
    Case cwiREGULAR_VIEW
        If Not mfrmRegularView Is Nothing Then Unload mfrmRegularView
        Set mfrmRegularView = Nothing
        Call RemoveMenus
        
    Case cwiPATIENT_SUMMARY_VIEW
        If Not mfrmSummaryView Is Nothing Then Unload mfrmSummaryView
        Set mfrmSummaryView = Nothing
    
    Case cwiALL_VIEWS
        Call CloseView(cwiREGULAR_VIEW)
        Call CloseView(cwiPATIENT_SUMMARY_VIEW)
    Case Else
    
    End Select
End Sub

'*******************************************************************************************
'  Description: This method is called when the user has selected a menu item (from the CW
'               Core's main menu), which is "owned" by this OLE Server.
'
'               NOTE: This is an OPTIONAL method.
'
'  Parameters:
'    vlMenuId - Contains the menu id of the menu item the user selected.
'*******************************************************************************************
Public Sub MenuItemSelected(ByVal vlMenuId As Long)
    Dim objTemplates As Object
    Dim szKey As String
    Dim mobjCurrEnc As Encounter
    Dim objEncounterOps As EncounterOps
    Dim objApptOps As CHCSII_AppointmentClient.AppointmentOps
    Dim objAppt As CHCSII_AppointmentClient.Appointment
    Dim objAppts As Appointments

    On Error GoTo ErrHandler
    
    ' Begin:SCR #45240;   Developer:  11/20/2003
    If mbButtonClicked Then Exit Sub
    mbButtonClicked = True
    ' End: SCR #45240;
    
    
    'MAG: Tracker: 40845
    Screen.MousePointer = vbHourglass
    
    szKey = mfrmRegularView.CurrentEncounterKey()
    If szKey = vbNullString Then GoTo Cleanup
    Set mobjCurrEnc = mobjEncounters(szKey)
    
    Select Case vlMenuId
    Case mlMenuAddNote
        ' Begin:SCR #42435;   Developer: MAG 09/15/2003 04:23 PM
        If mobjShared.IsAppMode(modeCHCSI_GUI) Then
          Set objAppts = mobjComm.InitializeOLEServer(NCID_CLINIC_SCHEDULE_APPLICATION)
          Set objApptOps = objAppts.AppointmentOps
          Set objAppt = objApptOps.GetAppointment_CHCS_GUI(mobjCurrEnc.AppointmentID, False)
    
          If Not objAppt Is Nothing Then
            Select Case UCase(objAppt.ApptStatus)
            Case "CANCEL"
                MsgBxARMd "You cannot open an appointment that has already been cancelled."
                GoTo Cleanup
            End Select
          End If
        End If
        ' End: SCR #42435;
        
        ' Begin:SCR #42275;   Developer: MAG 09/11/2003 11:20 AM
        If mobjShared.IsAppMode(modeCHCSI_GUI) Then
            mobjLogon.NursingDivisionName = mobjCurrEnc.Clinic
            mobjLogon.NursingDivisionNCID = mobjCurrEnc.ClinicNCID
        End If
        ' End: SCR #42275;
        
        ' Begin:SCR #39936;#46162   Developer: MAG 08/25/2003 10:11 AM
        If mobjShared.IsAppMode(modeCHCSI_GUI) And mobjCurrEnc.EncounterID = 11111111 Then
            Set objEncounterOps = GetEncounterOps()
            Set mobjCurrEnc = objEncounterOps.CreateCHCSEncounter(mobjCurrEnc.AppointmentID, mobjCurrEnc.PrimaryProvider.NCID)
        Else
            Call mobjCurrEnc.Refresh
        End If
        ' End: SCR #39936;
        
        If mobjCurrEnc.OpenEncounter(Me) Then '{{BOOKMARK Circular Ref ????}}
            'Tracker #50506
            mobjComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_PREVIOUS_ENCOUNTERS_APPLICATION
            Call mobjCurrEnc.AppendNote
            mobjComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_PREVIOUS_ENCOUNTERS_APPLICATION
            Call mobjCurrEnc.CloseEncounter
            Call mfrmRegularView.RefreshDisplay(GetInpatientNotes(True, False))
        End If
    Case mlMenuAmend
        ' Begin:SCR #42435;   Developer: MAG 09/15/2003 04:23 PM
        If mobjShared.IsAppMode(modeCHCSI_GUI) Then
          Set objAppts = mobjComm.InitializeOLEServer(NCID_CLINIC_SCHEDULE_APPLICATION)
          Set objApptOps = objAppts.AppointmentOps
          Set objAppt = objApptOps.GetAppointment_CHCS_GUI(mobjCurrEnc.AppointmentID, False)
    
          If Not objAppt Is Nothing Then
            Select Case UCase(objAppt.ApptStatus)
            Case "CANCEL"
                MsgBxARMd "You cannot open an appointment that has already been cancelled."
                GoTo Cleanup
            End Select
          End If
        End If
        ' End: SCR #42435;
        
        ' Begin:SCR #42275;   Developer: MAG 09/11/2003 11:20 AM
        If mobjShared.IsAppMode(modeCHCSI_GUI) Then
            mobjLogon.NursingDivisionName = mobjCurrEnc.Clinic
            mobjLogon.NursingDivisionNCID = mobjCurrEnc.ClinicNCID
        End If
        ' End: SCR #42275;
        
        ' Begin:SCR #39936;#46162  Developer: MAG 08/25/2003 10:11 AM
        If mobjShared.IsAppMode(modeCHCSI_GUI) And mobjCurrEnc.EncounterID = 11111111 Then
            Set objEncounterOps = GetEncounterOps()
            Set mobjCurrEnc = objEncounterOps.CreateCHCSEncounter(mobjCurrEnc.AppointmentID, mobjCurrEnc.PrimaryProvider.NCID)
        Else
            Call mobjCurrEnc.Refresh
        End If
        ' End: SCR #39936;
        
        If mobjCurrEnc.OpenEncounter(Me) Then
            Call mobjCurrEnc.Amend
            '''''''''''''''''''''''''''''''
            ' SCR 11537                   '
            ' Previous Encounters Retains '
            ' focus when Ammend Encounter '
            ' is selected.                '
            '''''''''''''''''''''''''''''''
            'Call mfrmRegularView.RefreshDisplay(GetInpatientNotes(True))
        End If
    Case mlMenuSign
        Call mobjCurrEnc.Refresh
        If mobjCurrEnc.OpenEncounter(Me) Then
            Call mobjCurrEnc.SignEncounter
            Call mobjCurrEnc.CloseEncounter
            Call mfrmRegularView.RefreshDisplay(GetInpatientNotes(True, False))
        End If
    End Select
    GoTo Cleanup
ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description, "EncPrevious.MenuItemSelected", "EncPrevious", vbExclamation)
Cleanup:
    Set objTemplates = Nothing
    Set mobjCurrEnc = Nothing
    Set objApptOps = Nothing
    Set objAppt = Nothing
    Screen.MousePointer = vbDefault
    ' Begin:SCR #45240;   Developer:  11/20/2003
    mbButtonClicked = False
    ' End: SCR #45240;
End Sub

'**************************************************************************************************************
'  Description: This method gets called when a new patient has been selected from the CW Core.
'
'               NOTE: This is an OPTIONAL method.
'
'  Parameters:
'    none
'**************************************************************************************************************
Public Sub PatientSelected()
    If Not mfrmRegularView Is Nothing Then
        mfrmRegularView.RefreshDisplay GetInpatientNotes(True, False)
    End If
   
End Sub

'*******************************************************************************************
'  Description: This method is called when a change has occurred in the "client area"
'               used by this OLE Server, resulting in a need to resize a particular "view".
'
'  Parameters:
'    viViewType - Contains the "type" of view being requested to resize by the CW Core.
'    vlLeft - Contains the "left" coordinate for the OLE server.
'    vlTop - Contains the "top" coordinate for the OLE server.
'    vlWidth - Contains the "width" for the OLE server.
'    vlHeight - Contains the "height" for the OLE server.
'
' Change HISTORY:
' * MSolano  1/14/2002 SCR-19545
'    The ARM transaction corresponding to the launch of an
'    OLE child server is instrumented in Core:
'        CW.frmCW.ServiceOLEServerMenuItem
'    Can detect here if transaction fails (business rules) or aborts (run time error).
'    Resolve ARM start call w/ ARM stop call using appropriate flag.
'*******************************************************************************************
Public Sub ResizeView(ByVal viViewType As Integer, ByVal vlLeft As Long, ByVal vlTop As Long, ByVal vlWidth As Long, ByVal vlHeight As Long)
    On Error GoTo ErrHandler
    
    Select Case viViewType
    Case cwiREGULAR_VIEW
        If Not mbFirstRegularResize Then
            '---------------------------------------------------------------------------------
            '  Reset the "WindowState" to "Normal", so the form can be resized.  This is
            '    needed because a window CANNOT be moved or resized if it is "minimized" or
            '    "maximized".
            '
            '  NOTE: This generates a form "resize" event.
            '---------------------------------------------------------------------------------
            mfrmRegularView.WindowState = vbNormal
            
            '---hide the form to so that annyoing "screen blinking" will not be seen while controls are moved and resized---
            mfrmRegularView.Visible = False
            
            '---resize form to fill entire "client area", NOTE: This generates a "resize" event---
            mfrmRegularView.ResizeView vlLeft, vlTop, vlWidth, vlHeight
            mfrmRegularView.Visible = True
        Else    '--first "regular" resize, need to perform "initailization" procedures--
            '---reset flag so this code will not be executed again---
            mbFirstRegularResize = False
            mfrmRegularView.ResizeView vlLeft, vlTop, vlWidth, vlHeight
        End If
        
    Case cwiPATIENT_SUMMARY_VIEW
      ' SCR #48958, 49463;   Developer: MAG 02/11/2004
      If mbFirstSummaryResize Then
        mfrmSummaryView.WindowState = vbNormal
        
        '---hide the form to so that annyoing "screen blinking" will not be seen while controls are moved and resized---
        mfrmSummaryView.Visible = False
        mfrmSummaryView.ResizeView vlLeft, vlTop, vlWidth, vlHeight
        mfrmSummaryView.Visible = True
        mbFirstSummaryResize = False
      Else
        mfrmSummaryView.ListView1.Visible = False
        mfrmSummaryView.ListView1.Visible = True
      End If
    End Select
        
    Exit Sub
ErrHandler:
    goARM.StopFarTx NCID_PREVIOUS_ENCOUNTERS_APPLICATION & gsARM_ACT_OPEN, enmARMTxStatusAborted  '--- SCR-ARM-19545 MSolano 12/21/01
    Call mobjShared.ShowVBError(Err.Number, Err.Description, "EncPrevious.ResizeView", "EncPrevious", vbExclamation)
End Sub

'**************************************************************************************************************
'  Description: This method is used by the CW Core to send messages to this OLE Server.  The
'               messages can be from the CW Core or from another OLE Server.
'
'  Parameters:
'    viMessageType - Contains the "type" of message to process.
'    vsMessageText - Contains any "data" assoicated with this message.
'    vsSourceServerNCID - Contains the NCID of the OLE Server that "sent" the message.
'**************************************************************************************************************
Public Sub Message(ByVal viMessageType As Integer, _
                   ByVal vsMessageText As String, _
                   ByVal vsSourceServerNCID As String)
    On Error Resume Next
    
    Dim oEncounter As Encounter
    Dim iCount As Integer
    
    Select Case viMessageType
        Case cwiREFRESH_DATA '{{BOOKMARK Only refresh when appropriate app sends cwiREFRESH_DATA.}}
            Select Case vsSourceServerNCID
            Case NCID_ENCOUNTER_APPLICATION, NCID_CW_APPLICATION
                If Not mfrmRegularView Is Nothing Then
                    mfrmRegularView.RefreshDisplay GetInpatientNotes(True, False)
                End If
            
            Case Else
            
            End Select
        Case cwiFOREGROUND_APPLICATION
            ' SCR #48025;   Developer: MAG 01/30/2004
            Select Case vsSourceServerNCID
            Case NCID_ENCOUNTER_APPLICATION, NCID_CW_APPLICATION
                If Not mfrmRegularView Is Nothing Then
                    mfrmRegularView.RefreshDisplay GetInpatientNotes(True, False)
                End If
            
            Case Else
            
            End Select
        Case cwiPRINT_PREVIEW
        Case cwiPRINT_REGULAR_VIEW
        Case cwiABOUT
            '---inform CW that I am showing a modal form---
            mobjComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_PREVIOUS_ENCOUNTERS_APPLICATION
            Dim objAbout As New CWAbout
            Call objAbout.Display(App, mfrmRegularView)
            Set objAbout = Nothing
            '---inform CW that I am done showing a modal form---
            mobjComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_PREVIOUS_ENCOUNTERS_APPLICATION
        Case cwiCLOSE_OLE_SERVER
            CloseView cwiREGULAR_VIEW
            mobjComm.CloseView cwiREGULAR_VIEW, mlProgID
    End Select
    
    
End Sub

'****************************************************************************************************
'  Description: This method is called when the CW Core creates an "instance" of this OLE Server.  It
'               is used to create and initialize global variables that are shared across all
'               "views" supported by this OLE Server, i.e. variables that only need to be
'               initialized ONCE, upon object creation.
'
'  Parameters:
'    vlMyProgId - Contains the "ProgId" that the CW Core has assigned to this OLE Server.
'    robjComm - Contains a communications object that is used to handle all communication
'               between this OLE Server and the CW Core.
'    robjLogon - Contains the logon environment object.  Contains information such
'                           as who logged on, where they logged on, etc.
'    robjConfig - Contains the configuration object that is used to get and set configuration
'                 information for a specific application.
'    robjPatient - Contains the patient object containing the currently selected patient.
'    robjUser - Security User information object
'****************************************************************************************************
Public Sub ICHCSII_CoreAPI2_Initialize2(ByVal vlMyProgId As Long, _
                                              robjComm As Object, _
                                              robjLogon As Object, _
                                              robjConfig As Object, _
                                              robjPatient As Object, _
                                              robjUser As Object)
    mlProgID = vlMyProgId
    Set cwobjCWComm = robjComm
    
    Set mobjComm = robjComm
    Set mobjLogon = robjLogon
    Set mobjConfig = robjConfig
    Set mobjPatient = robjPatient
    Set mobjUser = robjUser

End Sub

Private Sub AddMenus()
    On Error Resume Next
    If mlMenuAddNote = 0 Then
        ' SCR #43000;   Developer:  01/05/2004
        mlMenuAddNote = mobjComm.AddActionMenuItem("Append &Note", "new.bmp", mlProgID, "Append a note to the selected note.")
        mfrmRegularView.lMenuAddNote = mlMenuAddNote
    End If
    If mlMenuAmend = 0 Then
        ' SCR #42999;   Developer:  01/05/2004
        mlMenuAmend = mobjComm.AddActionMenuItem("&Amend Note", "edit.bmp", mlProgID, "Amend the selected note.")
        mfrmRegularView.lMenuAmend = mlMenuAmend
    End If
    If mlMenuSign = 0 Then
        mlMenuSign = mobjComm.AddActionMenuItem("&Sign Note", "dispsign.bmp", mlProgID, "Sign the selected note.")
        mfrmRegularView.lMenuSign = mlMenuSign
    End If
    
End Sub

Friend Sub RemoveMenus()
    If mlMenuAddNote <> 0 Then
        Call mobjComm.DeleteMenuItem(mlMenuAddNote)
        mlMenuAddNote = 0
    End If
    If mlMenuAmend <> 0 Then
        Call mobjComm.DeleteMenuItem(mlMenuAmend)
        mlMenuAmend = 0
    End If
    If mlMenuSign <> 0 Then
        Call mobjComm.DeleteMenuItem(mlMenuSign)
        mlMenuSign = 0
    End If
End Sub


Public Sub SetEncounterKey(ByRef StorageKey As String)
    Call mfrmRegularView.SelectEncounter(StorageKey)
End Sub
Private Sub Class_Initialize()
    
    '---create shared object used to show VB errors, etc.---
    Set mobjShared = New CWShared
    '---set the help file for this application---
    App.HelpFile = mobjShared.CWHelpFile
    
    '--- SCR-ARM-19545 MSolano 12/21/01
    Call ARM_InitAllTransactionsForEncounterPrev
    
End Sub

Private Sub Class_Terminate()
    
    '---release reference to objects---
    Set mobjShared = Nothing
    Set mobjComm = Nothing
    Set mobjLogon = Nothing
    Set mobjConfig = Nothing
    Set mobjPatient = Nothing
    Set mobjUser = Nothing
    Set mobjEncounters = Nothing
    Set mcolEncChild = Nothing

End Sub

' Begin:SCR #34000;   Developer: MAG 05/13/2003 09:54 AM
Public Function FormatBasedOnStatus(oCurEnc As Encounter) As Long
   mobjComm.UpdateMenuItem mlMenuAddNote, "DISABLE"
   mobjComm.UpdateMenuItem mlMenuAmend, "DISABLE"
   mobjComm.UpdateMenuItem mlMenuSign, "DISABLE"
    
  
  With oCurEnc
    If Not (.Status = PatientLeftWithoutBeingSeen _
         Or .Status = CancelByPatient _
         Or .Status = CancelByFacility _
         Or .Status = NoShow) _
    Then
      ' Begin:SCR #42998;   Developer: MAG 01/14/2004
      'If .Status = NeedsCoSignature _
      'And mobjUser.HasPrivilegeEx(Priv_Current_Encounter, Priv_Sign) Then
      If mobjUser.HasPrivilegeEx(Priv_Current_Encounter, Priv_Sign) Then
        If (.Status = Updating) Or (.Status = InProgress) Then
          Call mobjComm.UpdateMenuItem(mlMenuSign, "ENABLE")
        End If
      End If
      ' End: SCR #42998;
        
      ' Begin:SCR #39946;   Developer: MAG 08/21/2003 10:17 AM
      If mobjShared.IsAppMode(modeCHCSI_GUI) And .UserIsPrimaryProvider Then
        If Not .UpdateLock Then
          Call mobjComm.UpdateMenuItem(mlMenuAmend, "ENABLE")
        End If
      ElseIf (Not mobjShared.IsAppMode(modeCHCSI_GUI)) And (Not .WhoSigned Is Nothing) Then
        If (.WhoSigned.NCID = mobjShared.CurrentUserNCID) Then
        'Developer: MAG 02/02/2004 Tracker #42999, 48955 Only the owner of an Inpatient Note allowed to Ammend Note
          If Not .UpdateLock Then
            Call mobjComm.UpdateMenuItem(mlMenuAmend, "ENABLE")
          End If
        End If
      End If
      ' End: SCR #39946;
      
      If mobjUser.HasPrivilegeEx(Priv_Current_Encounter, Priv_Write) Then
        If Not .UpdateLock And Not .Status = Updating Then
          Call mobjComm.UpdateMenuItem(mlMenuAddNote, "ENABLE")
        End If
      End If
    End If
  End With

End Function
' End: SCR #34000;

