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
Attribute VB_Name = "InpatientDocs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements ICHCSII_CoreAPI2             '-- implements the security obj
Implements ICHCSII_CoreAPI3

Private mcolEncChild As Collection      '-- encounter child modules
Private mlProgID As Long                '-- my prog id from core
Private mbMainViewRunning As Boolean
Private mbFirstRegularResize As Boolean
Private mbFirstSummaryResize As Boolean
Private mbUserCanWrite As Boolean 'SCR 72586 Checks the "Inpatient" - "Write" priv.

'MenuIDs
Private mlMenuRefresh As Long
Private mlMenuAddNote As Long 'For adding inpatient notes
Private mlMenuAppendEnc As Long 'For adding (appending) outpatient encounters
Private mlMenuAmend As Long
Private mlMenuSign As Long
Private mlMenuCreateTemplate As Long
Private mlMenuPrintSF600  As Long
Private mlMenuPrintPreviewSF600  As Long
Private mlMenuPrintSF600Button As Long
Private mlMenuPrintPreviewSF600Button As Long
Private mlMenuNewOrder As Long
Private mlMenuPrintList As Long

Private moDocAdmissions As DocAdmissions

'Private mfrmRegularView As frmEncPrevious
'Private mfrmSummaryView As frmEncPrevious
Private WithEvents mfrmDocuments As frmDocuments
Attribute mfrmDocuments.VB_VarHelpID = -1
Private mbButtonClicked As Boolean

'-- implements the dlog.dll timing module
#If debugon Then
  Private DR As DebugRec
  Private DL As cDebugLog
  Private Const Exe = "Documents"
  Private Const Module = "InpatientDocs.cls"
#End If

Private mobjEncounters As CHCSIIEncounterPrevious.Encounters '-- previous encounters

Public Sub SetEncounterKey(ByRef StorageKey As String)
    Call mfrmDocuments.SelectEncounter(StorageKey)
End Sub

Private Function DisplayDocuments(SetDefaultFilter As Boolean)

'Load the Admissions and display the documents form
Dim sInpatientAdmissionID As String
Dim ocwInPatient As CHCSII_PAD.Inpatient

Static LastInitTime As Double
        
    If DateDiff("s", LastInitTime, Now) <= 1 Then
        Exit Function
    End If
    
    Call AddMenus
    
    Load mfrmDocuments
    
    mfrmDocuments.Initialize
    
    'Get the InpatientAdmissionID from the Inpatient List Module
    Set ocwInPatient = goBB.Comm.InitializeOLEServer(NCID_INPATIENT_APPLICATION)
    If Not ocwInPatient.SelectedPatients Is Nothing Then
        sInpatientAdmissionID = ocwInPatient.SelectedPatients.Inpatient.ID

        If SetDefaultFilter = True Then
            If sInpatientAdmissionID = "" Then
                'RQT 62107 If the user hasn't selected an admission from Inpatient List, then default the
                'view to Both
                mfrmDocuments.optFilter(0).Value = True
            Else
                mfrmDocuments.optFilter(1).Value = True
            End If
        End If
    End If
    
    'Refresh based on the selected option button
    mfrmDocuments.RefreshGridBasedOnFilter
    
    mfrmDocuments.Show vbModeless
    
    LastInitTime = Now
        
End Function

Private Sub EnableMenu(MenuID As Long, Enable As Boolean)
    
    'Enable or disable the menu
    
    If Enable = True Then
        'Enable the menu
        Call goBB.Comm.UpdateMenuItem(MenuID, "ENABLE")
    Else
        'Disable the menu
        Call goBB.Comm.UpdateMenuItem(MenuID, "DISABLE")
    End If
    
End Sub


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
        QueryView = False
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
        Set mfrmDocuments = New frmDocuments
        
        SetParent mfrmDocuments.hwnd, vlParenthWnd
        
        rhMainForm = mfrmDocuments.hwnd
        OpenView = True
        
        DisplayDocuments True
        
        'Call mfrmRegularView.RefreshDisplay(GetInpatientNotes(True, False))
        
'    Case cwiPATIENT_SUMMARY_VIEW
'        mbFirstSummaryResize = True
'        Set mfrmSummaryView = New frmEncPrevious
'        mfrmSummaryView.SummaryView = True
'        mfrmSummaryView.Inpatient = True
'        mfrmSummaryView.InitListView
'        ' SCR #46974;   Developer:  12/22/2003
'        mfrmSummaryView.lblEnc(1).Caption = "Inpatient Notes"
'        mfrmSummaryView.lblEnc(0).Caption = Space(12) & "Signed Inpatient Notes:"
'
'        SetParent mfrmSummaryView.hwnd, vlParenthWnd
'        rhMainForm = mfrmSummaryView.hwnd
'        OpenView = True
'
'        Call mfrmSummaryView.RefreshDisplay(GetInpatientNotes(True, True))
    Case Else    '--unsupported view "type"--
        OpenView = False
        goARM.StopFarTx NCID_INPATIENT_DOCUMENTS_APPLICATION & gsARM_ACT_OPEN, enmARMTxStatusFailed  '--- SCR-ARM-19545 MSolano 12/21/01
    End Select
    Exit Function

ErrHandler:
    goARM.StopFarTx NCID_INPATIENT_DOCUMENTS_APPLICATION & gsARM_ACT_OPEN, enmARMTxStatusAborted '--- SCR-ARM-19545 MSolano 12/21/01
    Call goBB.SharedObj.ShowVBError(Err.Number, Err.Description, "InPatientNotes.OpenView", "Enc", vbCritical)
    OpenView = False
    Exit Function
    
    Resume
End Function

'Private Function GetInpatientNotes(ByVal bRefresh As Boolean, ByVal Lite As Boolean) As CHCSIIEncounterPrevious.Encounters
'
'    If bRefresh Or (mobjEncounters Is Nothing) Then
'        Set mobjEncounters = New CHCSIIEncounterPrevious.Encounters
'        mobjEncounters.GetInpatientNotes goBB.Patient.UnitNumber, Lite
'    End If
'
'    Set GetInpatientNotes = mobjEncounters
'
'End Function
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
        If Not mfrmDocuments Is Nothing Then Unload mfrmDocuments
        Set mfrmDocuments = Nothing
        Call RemoveMenus
        
'    Case cwiPATIENT_SUMMARY_VIEW
'        If Not mfrmSummaryView Is Nothing Then Unload mfrmSummaryView
'        Set mfrmSummaryView = Nothing
'
    Case cwiALL_VIEWS
        Call CloseView(cwiREGULAR_VIEW)
        'Call CloseView(cwiPATIENT_SUMMARY_VIEW)
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
    
    'Dim szKey As String
    'Dim oEncParent As CHCSIIEncounterCurrent.EncounterParent
    Dim oEncounter As CHCSIIEncounterOps.Encounter
    'Dim oEncounterOps As CHCSIIEncounterOps.EncounterOps
    'Dim objApptOps As CHCSII_AppointmentClient.AppointmentOps
    'Dim objAppt As CHCSII_AppointmentClient.Appointment
    'Dim objAppts As Appointments
    Dim mobjCurrEnc As Encounter
    
    On Error GoTo ErrHandler
    
    ' Begin:SCR #45240;   Developer:  11/20/2003
    If mbButtonClicked Then Exit Sub
    mbButtonClicked = True
    ' End: SCR #45240;
    
    
    'MAG: Tracker: 40845
    Screen.MousePointer = vbHourglass
    
'SF TODO:
'    szKey = mfrmRegularView.CurrentEncounterKey()
'    If szKey = vbNullString Then GoTo Cleanup
'    Set oEncounter = mobjEncounters(szKey)
    
    Select Case vlMenuId
    
        Case mlMenuRefresh
            'Refresh from the DB
            'Don't call this because by clicking on the menu button, the Foreground message is sent which refreshes the grid.
            DisplayDocuments False
            
        Case mlMenuAddNote
            'Call the add note method of the inpatient list module
            mfrmDocuments.AddInpatientNote
    

        Case mlMenuAppendEnc
            
            mfrmDocuments.AppendOutpatientEncounter
            
            'SCR 67248 Refresh after adding a note
            DisplayDocuments False
            
'            Call oEncounter.Refresh
            
'            If oEncounter.OpenEncounter(Me) Then '{{BOOKMARK Circular Ref ????}}
'                'Tracker #50506
'                goBB.Comm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_PREVIOUS_ENCOUNTERS_APPLICATION
'                Call oEncounter.AppendNote
'                goBB.Comm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_PREVIOUS_ENCOUNTERS_APPLICATION
'                Call oEncounter.CloseEncounter
'                'SF TODO:
'                'Call mfrmRegularView.RefreshDisplay(GetInpatientNotes(True, False))
'            End If
        
        Case mlMenuAmend
'
'                Call oEncounter.Refresh
            
            mfrmDocuments.AmendEncounter
            
'            If oEncounter.OpenEncounter(Me) Then
'                Call oEncounter.Amend
'                '''''''''''''''''''''''''''''''
'                ' SCR 11537                   '
'                ' Previous Encounters Retains '
'                ' focus when Ammend Encounter '
'                ' is selected.                '
'                '''''''''''''''''''''''''''''''
'                'Call mfrmRegularView.RefreshDisplay(GetInpatientNotes(True))
'            End If
            
        Case mlMenuSign
            mfrmDocuments.SignEncounter
'            Call oEncounter.Refresh
'            If oEncounter.OpenEncounter(Me) Then
'                Call oEncounter.SignEncounter
'                Call oEncounter.CloseEncounter
'                'SF TODO:
'                'Call mfrmRegularView.RefreshDisplay(GetInpatientNotes(True, False))
'            End If

        Case mlMenuCreateTemplate
            mfrmDocuments.CreateTemplate
                        
        Case mlMenuPrintSF600, mlMenuPrintSF600Button
            ' Begin:SCR #42777;   Developer:  10/17/2003
            EnableMenu mlMenuPrintSF600, False
            EnableMenu mlMenuPrintPreviewSF600, False
            mfrmDocuments.PrintEncounters False
            EnableMenu mlMenuPrintPreviewSF600, True
            EnableMenu mlMenuPrintSF600, True
            ' End: SCR #42777;
            
        Case mlMenuPrintPreviewSF600, mlMenuPrintPreviewSF600Button
            ' Begin:SCR #42777;   Developer:  10/17/2003
            EnableMenu mlMenuPrintSF600, False
            EnableMenu mlMenuPrintPreviewSF600, False
            mfrmDocuments.PrintEncounters True
            EnableMenu mlMenuPrintPreviewSF600, True
            EnableMenu mlMenuPrintSF600, True
            ' End: SCR #42777;

        Case mlMenuNewOrder
            Call mfrmDocuments.AddInpatientNote(NCID_Orders_Note)
            
        Case mlMenuPrintList
            Call mfrmDocuments.PrintList
            
    End Select
    GoTo Cleanup
ErrHandler:
    Call goBB.SharedObj.ShowVBError(Err.Number, Err.Description, "EncPrevious.MenuItemSelected", "EncPrevious", vbExclamation)
Cleanup:
    'Set objTemplates = Nothing
    Set oEncounter = Nothing
    'Set oEncounterOps = Nothing
    'Set objApptOps = Nothing
    'Set objAppt = Nothing
    Screen.MousePointer = vbDefault
    ' Begin:SCR #45240;   Developer:  11/20/2003
    mbButtonClicked = False
    ' End: SCR #45240;
    
    Exit Sub
    Resume
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
    If Not mfrmDocuments Is Nothing Then
        'TODO: Refresh based on new patient.
        'mfrmRegularView.RefreshDisplay GetInpatientNotes(True, False)
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
            If Not mfrmDocuments Is Nothing Then
                mfrmDocuments.WindowState = vbNormal
                
                '---hide the form to so that annyoing "screen blinking" will not be seen while controls are moved and resized---
                mfrmDocuments.Visible = False
                
                '---resize form to fill entire "client area", NOTE: This generates a "resize" event---
                'mfrmDocuments.ResizeView vlLeft, vlTop, vlWidth, vlHeight
                mfrmDocuments.Move vlLeft, vlTop, vlWidth, vlHeight
                mfrmDocuments.Visible = True
                mfrmDocuments.Refresh    '--- SCR 88279  Sherry Wang  2/17/2006
            Else
                MsgBox "Can't display documents."
            End If
        Else    '--first "regular" resize, need to perform "initailization" procedures--
            '---reset flag so this code will not be executed again---
            mbFirstRegularResize = False
            'mfrmDocuments.ResizeView vlLeft, vlTop, vlWidth, vlHeight
            mfrmDocuments.Move vlLeft, vlTop, vlWidth, vlHeight
        End If
        
'    Case cwiPATIENT_SUMMARY_VIEW
'      ' SCR #48958, 49463;   Developer: MAG 02/11/2004
'      If mbFirstSummaryResize Then
'        mfrmSummaryView.WindowState = vbNormal
'
'        '---hide the form to so that annyoing "screen blinking" will not be seen while controls are moved and resized---
'        mfrmSummaryView.Visible = False
'        mfrmSummaryView.ResizeView vlLeft, vlTop, vlWidth, vlHeight
'        mfrmSummaryView.Visible = True
'        mbFirstSummaryResize = False
'      Else
'        mfrmSummaryView.ListView1.Visible = False
'        mfrmSummaryView.ListView1.Visible = True
'      End If
    End Select
        
    Exit Sub
ErrHandler:
    goARM.StopFarTx NCID_INPATIENT_DOCUMENTS_APPLICATION & gsARM_ACT_OPEN, enmARMTxStatusAborted  '--- SCR-ARM-19545 MSolano 12/21/01
    Call goBB.SharedObj.ShowVBError(Err.Number, Err.Description, "EncPrevious.ResizeView", "EncPrevious", vbExclamation)
    Exit Sub
    Resume
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
        'SCR 67541
        Case cwiOK_TO_CHANGE_PATIENT
             goBB.Comm.Message cwiOK_TO_CHANGE_PATIENT, "Y", vsSourceServerNCID, NCID_INPATIENT_DOCUMENTS_APPLICATION
           'The Close_OLE_Server messages causes documents to display itself.
            'goBB.Comm.Message cwiCLOSE_OLE_SERVER, NCID_INPATIENT_DOCUMENTS_APPLICATION, NCID_INPATIENT_DOCUMENTS_APPLICATION, NCID_INPATIENT_DOCUMENTS_APPLICATION
'            If Not mfrmDocuments Is Nothing Then
'                Unload mfrmDocuments
'                Set mfrmDocuments = Nothing
'                Call RemoveMenus
'            End If
            CloseView cwiREGULAR_VIEW
            goBB.Comm.CloseView cwiREGULAR_VIEW, mlProgID
            
        Case cwiREFRESH_DATA '{{BOOKMARK Only refresh when appropriate app sends cwiREFRESH_DATA.}}
            Select Case vsSourceServerNCID
            Case NCID_INPATIENT_DOCUMENTS_APPLICATION, NCID_CW_APPLICATION
                If Not mfrmDocuments Is Nothing Then
                    'TODO: Does this need to be here?
                    DisplayDocuments False
                End If
            
            Case Else
            
            End Select
        Case cwiFOREGROUND_APPLICATION
            'This causes so many refresh problems because core calls it when it shouldn't

            Select Case vsSourceServerNCID
            Case NCID_ENCOUNTER_APPLICATION, NCID_CW_APPLICATION, NCID_INPATIENT_DOCUMENTS_APPLICATION
                If Not mfrmDocuments Is Nothing Then
                    'TODO: Does this need to be here?
                    DisplayDocuments False
                End If

            Case Else

            End Select
        Case cwiPRINT_PREVIEW
        Case cwiPRINT_REGULAR_VIEW
            
        Case cwiABOUT
            '---inform CW that I am showing a modal form---
            goBB.Comm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_PREVIOUS_ENCOUNTERS_APPLICATION
            Dim objAbout As New CWAbout
            'Call objAbout.Display(App, mfrmDocuments)
            Set objAbout = Nothing
            '---inform CW that I am done showing a modal form---
            goBB.Comm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_PREVIOUS_ENCOUNTERS_APPLICATION
        Case cwiCLOSE_OLE_SERVER
            
            CloseView cwiREGULAR_VIEW
            goBB.Comm.CloseView cwiREGULAR_VIEW, mlProgID
            
        Case cwiCOPY
            Clipboard.Clear
            Clipboard.SetText (mfrmDocuments.TxText.SelText)

    End Select
End Sub

Public Sub SetInpatientKey(ByRef InpatientID As String)
    
    Call mfrmDocuments.SelectAdmission(InpatientID)
    
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
    
    Set goBB = New Blackboard
    Set goBB.Comm = robjComm
    Set goBB.Logon = robjLogon
    Set goBB.Config = robjConfig
    Set goBB.Patient = robjPatient
    Set goBB.User = robjUser
    
    mbUserCanWrite = goBB.User.HasPrivilegeEx("Inpatient", Priv_Write)

End Sub

Private Sub AddMenus()
    
    On Error GoTo ErrHnd
    
    If mlMenuRefresh = 0 Then
        mlMenuRefresh = goBB.Comm.AddActionMenuItem("&Refresh", "refresh.bmp", mlProgID, "Refresh the display")
    End If
    
    If mlMenuAddNote = 0 Then
        mlMenuAddNote = goBB.Comm.AddActionMenuItem("Add Inpt &Note", "new.bmp", mlProgID, "Add a note to the selected admission.")
        'mfrmRegularView.lMenuAddNote = mlMenuAddNote
    End If
    
    If mlMenuNewOrder = 0 Then
        mlMenuNewOrder = goBB.Comm.AddActionMenuItem("&New Order", "new.bmp", mlProgID, "Create a new order.")
        'mfrmRegularView.lMenuSign = mlMenuSign
    End If
    
    If mlMenuAppendEnc = 0 Then
         '--- SCR 85391 Sherry Wang   1/24/2006
        mlMenuAppendEnc = goBB.Comm.AddActionMenuItem("Append N&ote", "new.bmp", mlProgID, "Append a note to the selected encounter.")
        'mfrmRegularView.lMenuAddNote = mlMenuAddNote
    End If
    
    If mlMenuAmend = 0 Then
        ' SCR #42999;   Developer:  01/05/2004
        mlMenuAmend = goBB.Comm.AddActionMenuItem("&Amend Note", "edit.bmp", mlProgID, "Amend the selected note.")
        'mfrmRegularView.lMenuAmend = mlMenuAmend
    End If
    
    If mlMenuSign = 0 Then
        mlMenuSign = goBB.Comm.AddActionMenuItem("&Sign Note", "dispsign.bmp", mlProgID, "Sign the selected note.")
        'mfrmRegularView.lMenuSign = mlMenuSign
    End If
   
    If Not goBB.SharedObj.IsAppMode(modeCHCSI_GUI) Then
      If mlMenuCreateTemplate = 0 Then
          mlMenuCreateTemplate = goBB.Comm.AddActionMenuItem("&NewTemplate", "create_enc_template.bmp", mlProgID, "Create new Encounter Template from this Encounter")
      End If
    End If
   
    'SFoster 2/24/2003
    'SCR 32913. Dental providers can print SF600 or SF603a.  All others can only print SF600.
'    If goBB.User.HasPrivilegeEx(Priv_Dental_Access, Priv_Read) Then
'        If mlMenuPrintSF600 = 0 Then
'            mlMenuPrintSF600 = goBB.Comm.AddPrintMenuItem("&SF603A/SF600", mlProgID, "Print SF600")
'        End If
'        If mlMenuPrintPreviewSF600 = 0 Then
'            mlMenuPrintPreviewSF600 = goBB.Comm.AddPrintPreviewMenuItem("S&F603A/SF600", mlProgID, "Print SF600")
'        End If
'    Else
        If mlMenuPrintSF600 = 0 Then
            mlMenuPrintSF600 = goBB.Comm.AddPrintMenuItem("SF600", mlProgID, "Print SF600")
        End If
        If mlMenuPrintPreviewSF600 = 0 Then
            mlMenuPrintPreviewSF600 = goBB.Comm.AddPrintPreviewMenuItem("SF600", mlProgID, "Print SF600")
        End If
               
    If mlMenuPrintSF600Button = 0 Then
        mlMenuPrintSF600Button = goBB.Comm.AddActionMenuItem("&Print", "print.bmp", mlProgID, "Print the selected note.")
        'mfrmRegularView.lMenuSign = mlMenuSign
    End If
               
    If mlMenuPrintPreviewSF600Button = 0 Then
        mlMenuPrintPreviewSF600Button = goBB.Comm.AddActionMenuItem("&Print Preview", "print.bmp", mlProgID, "Print preview the selected note.")
        'mfrmRegularView.lMenuSign = mlMenuSign
    End If
    
    'SCR 56560
    If mlMenuPrintList = 0 Then
        mlMenuPrintList = goBB.Comm.AddActionMenuItem("&Print List", "print.bmp", mlProgID, "Print the list.")
        'mfrmRegularView.lMenuSign = mlMenuSign
    End If
    
'    End If
   
    'Set the defaults
    EnableMenu mlMenuAddNote, False
    EnableMenu mlMenuAppendEnc, False
    EnableMenu mlMenuNewOrder, False
    EnableMenu mlMenuAmend, False
    EnableMenu mlMenuSign, False
    EnableMenu mlMenuPrintSF600Button, False
    EnableMenu mlMenuPrintPreviewSF600Button, False
    
Exit Sub

ErrHnd:
    MsgBxARMd "Error: " & Err.Description & ". In: " & Err.Source
    
End Sub

Friend Sub RemoveMenus()
    
    If mlMenuRefresh <> 0 Then
        Call goBB.Comm.DeleteMenuItem(mlMenuRefresh)
        mlMenuRefresh = 0
    End If
    If mlMenuAddNote <> 0 Then
        Call goBB.Comm.DeleteMenuItem(mlMenuAddNote)
        mlMenuAddNote = 0
    End If
    If mlMenuAppendEnc <> 0 Then
        Call goBB.Comm.DeleteMenuItem(mlMenuAppendEnc)
        mlMenuAppendEnc = 0
    End If
    If mlMenuAmend <> 0 Then
        Call goBB.Comm.DeleteMenuItem(mlMenuAmend)
        mlMenuAmend = 0
    End If
    If mlMenuSign <> 0 Then
        Call goBB.Comm.DeleteMenuItem(mlMenuSign)
        mlMenuSign = 0
    End If
    If mlMenuCreateTemplate <> 0 Then
        Call goBB.Comm.DeleteMenuItem(mlMenuCreateTemplate)
        mlMenuCreateTemplate = 0
    End If
    If mlMenuPrintSF600 <> 0 Then
        Call goBB.Comm.DeleteMenuItem(mlMenuPrintSF600)
        mlMenuPrintSF600 = 0
    End If
    If mlMenuPrintPreviewSF600 <> 0 Then
        Call goBB.Comm.DeleteMenuItem(mlMenuPrintPreviewSF600)
        mlMenuPrintPreviewSF600 = 0
    End If
    If mlMenuPrintSF600Button <> 0 Then
        Call goBB.Comm.DeleteMenuItem(mlMenuPrintSF600Button)
        mlMenuPrintSF600Button = 0
    End If
    If mlMenuPrintPreviewSF600Button <> 0 Then
        Call goBB.Comm.DeleteMenuItem(mlMenuPrintPreviewSF600Button)
        mlMenuPrintPreviewSF600Button = 0
    End If
    If mlMenuNewOrder <> 0 Then
        Call goBB.Comm.DeleteMenuItem(mlMenuNewOrder)
        mlMenuNewOrder = 0
    End If
    If mlMenuPrintList <> 0 Then
        Call goBB.Comm.DeleteMenuItem(mlMenuPrintList)
        mlMenuPrintList = 0
    End If
    
End Sub


Private Sub Class_Initialize()

    '---create shared object used to show VB errors, etc.---
    Set goBB = New Blackboard
    '---set the help file for this application---
    App.HelpFile = goBB.SharedObj.CWHelpFile
    
    '--- SCR-ARM-19545 MSolano 12/21/01
    Call ARM_InitAllTransactionsForEncounterPrev

End Sub

Private Sub Class_Terminate()
    
    '---release reference to objects---
    Set mobjEncounters = Nothing
    Set mcolEncChild = Nothing
    Set moDocAdmissions = Nothing
    Set mfrmDocuments = Nothing

End Sub

Private Function ICHCSII_CoreAPI3_CloseView2(ByVal viViewType As Integer) As Boolean
    
    CloseView viViewType
    
    ICHCSII_CoreAPI3_CloseView2 = True
    
End Function


Private Sub mfrmDocuments_FilterChange(UpdateMenus As Boolean)

    'Default the menus
    If UpdateMenus Then
        AddMenus
    End If

End Sub

Private Sub mfrmDocuments_GridRowChange(Admission As DTOs.DocAdmission, Encounter As DTOs.DocEncounter)
    
    'SCR 87550
    Dim mobjConn  As CHCSII_CONN.Conn
    Dim oSQLEx As ICHCSII_SQL.ISqlOps
    Dim sSQL As String
        
    'Refresh the menus based on the grid selection.
    'Logic based on RQT 42999, 42998
    If Not Admission Is Nothing Then
        'Admission object was selected
        
        'Disable the print feature for admissions
        EnableMenu mlMenuPrintSF600Button, False
        EnableMenu mlMenuPrintPreviewSF600Button, False
            
        'Disable create template
        EnableMenu mlMenuCreateTemplate, False
        
        'Enable the new order button
        EnableMenu mlMenuNewOrder, mbUserCanWrite
    
        If Admission.IsInpatient = True Then
            'Menu settings for inpatient
            EnableMenu mlMenuAddNote, mbUserCanWrite
            EnableMenu mlMenuAmend, False
            EnableMenu mlMenuAppendEnc, False
            EnableMenu mlMenuSign, False
            
        Else
            'Menu settings for outpatient - there is no admission for outpatient
            EnableMenu mlMenuAddNote, False
            EnableMenu mlMenuAppendEnc, False
            EnableMenu mlMenuAmend, False
            EnableMenu mlMenuSign, False
            
        End If
        
    ElseIf Not Encounter Is Nothing Then
        'Encounter object was selected
        
        'Enable the print feature for admissions
        EnableMenu mlMenuPrintSF600Button, True
        EnableMenu mlMenuPrintPreviewSF600Button, True
        
        'Enable create template
        EnableMenu mlMenuCreateTemplate, True
        
        'Enable the new order button
        EnableMenu mlMenuNewOrder, False

        If Encounter.Admission.IsInpatient = True Then
            
            'Menu settings for inpatient
            EnableMenu mlMenuAddNote, mbUserCanWrite
                                                                                        'RQT 31682                                              'RQT 27495
            If (Encounter.StatusEnum = complete Or Encounter.StatusEnum = Updated) And goBB.Logon.FacilityNCID = Encounter.FacilityNCID Then  'And Encounter.UpdateLock = False Then
                EnableMenu mlMenuAmend, True
                EnableMenu mlMenuAppendEnc, mbUserCanWrite
                EnableMenu mlMenuSign, False
                                                                                                                         'RQT 31682
            ElseIf (Encounter.StatusEnum = Waiting Or Encounter.StatusEnum = InProgress Or Encounter.StatusEnum = Updating) And goBB.Logon.FacilityNCID = Encounter.FacilityNCID Then
                EnableMenu mlMenuSign, True
                EnableMenu mlMenuAmend, False
                EnableMenu mlMenuAppendEnc, False
                
            Else
                'Disable for all other statuses
                EnableMenu mlMenuAmend, False
                EnableMenu mlMenuSign, False
            End If
            
        Else
        
            'Menu settings for outpatient
            EnableMenu mlMenuAddNote, False
                                                                                    'RQT 31682                                                  'RQT 27495
            '--- SCR 133860  Sherry Wang   7/6/2007
            '--- Do not enable append and amend buttons if updatelock = true.
            If (Encounter.StatusEnum = complete Or Encounter.StatusEnum = Updated) And goBB.Logon.FacilityNCID = Encounter.FacilityNCID And Encounter.UpdateLock = False Then
                
                EnableMenu mlMenuAppendEnc, mbUserCanWrite
                EnableMenu mlMenuAmend, True
                EnableMenu mlMenuSign, False
                                                                                                            'RQT 31682
            ElseIf (Encounter.StatusEnum = Waiting Or Encounter.StatusEnum = InProgress Or Encounter.StatusEnum = Updating) And goBB.Logon.FacilityNCID = Encounter.FacilityNCID Then
                EnableMenu mlMenuSign, True
            Else
                'Disable for all other statuses
                EnableMenu mlMenuAppendEnc, False
                EnableMenu mlMenuAmend, False
                EnableMenu mlMenuSign, False
            End If
            
        End If
    End If
    
    'SCR 87550
    If Encounter Is Nothing Then
        Exit Sub
    End If
    
    If mobjConn Is Nothing Then
        Set mobjConn = New CHCSII_CONN.Conn
    End If
    
    Set oSQLEx = mobjConn.CHCSII_SQLOPS(Auto)
    
    sSQL = "Select RETIRED, BMIST_ENC "
    sSQL = sSQL & " from ENCOUNTERS "
    sSQL = sSQL & " Where ENCOUNTERNUMBER = " & Encounter.EncounterNumber
    sSQL = sSQL & " And FACILITYNCID = " & Encounter.FacilityNCID
      
    oSQLEx.Execute (sSQL)

    If oSQLEx("RETIRED") = 1 Then
        EnableMenu mlMenuAppendEnc, False
        EnableMenu mlMenuAmend, False
    ElseIf oSQLEx("BMIST_ENC") = 1 Then 'SCR 128313
        EnableMenu mlMenuAmend, False
    End If
    
    Set oSQLEx = Nothing
    
End Sub

