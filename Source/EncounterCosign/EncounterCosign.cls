VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EncCosign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ICHCSII_CoreAPI2             '-- implements the security obj

Private Const SECT_NOTES As Integer = 8

Private mobjEncounterOps As CHCSIIEncounterOps.EncounterOps
Private mobjUser As ICHCSII_User        '-- The Security User Object
Private mcolEncounters As Collection    '-- Cosign (and current) encounters
Private mcolEncChild As Collection      '-- encounter child modules
Private mobjCurrEnc As Encounter        '-- current active encounter
Private mobjPatient As Patient          '-- patient object from core
Private mobjConfig As Config            '-- config obj from core
Private mobjLogon As Logon              '-- logon env obj from core
Private mobjComm As Comm                '-- comm obj from core
Private mlProgID As Long                '-- my prog id from core
Private mbMainViewRunning As Boolean
Private mbCloseViewMyself As Boolean
Private mbFirstRegularResize As Boolean
Private mobjShared As CWShared
Private mlMenuAddNote As Long
Private mlMenuAmend As Long
Private mlMenuSign As Long
Private mlMenuPrintSF600 As Long
Private mlMenuPrintPreviewSF600 As Long

Public Event OpenEncounter(objencounter As CHCSIIEncounterOps.Encounter)
Public Event CloseEncounter(objencounter As CHCSIIEncounterOps.Encounter)

'-- implements the dlog.dll timing module
#If DEBUGON Then
  Private DR As DebugRec
  Private DL As cDebugLog
  Private Const Exe = "PrevEnc"
  Private Const Module = "PrevEnc.cls"
#End If

'SCR 106386 This is called from EncounterOps.Encounter.SignEncounter
Public Sub RefreshDisplay(Optional ByVal ReloadData As Boolean = False)
    Call frmEncCosign.RefreshDisplay '(ReloadData) Do not pass this on because it does not mean reload data in the cosigns module.
End Sub

Friend Property Get EncounterOps() As EncounterOps:     Set EncounterOps = mobjEncounterOps:           End Property
Friend Property Get CWComm() As Comm:                   Set CWComm = mobjComm:                         End Property
Friend Property Get Config() As Config:                 Set Config = mobjConfig:                       End Property
Friend Property Get Logon() As Logon:                   Set Logon = mobjLogon:                         End Property
Friend Property Get Patient() As Patient:               Set Patient = mobjPatient:                     End Property
Friend Property Get ProgID() As Long:                   ProgID = mlProgID:                             End Property
Friend Property Get User() As ICHCSII_User:             Set User = mobjUser:                           End Property
Friend Property Get CWShared() As CWShared:             Set CWShared = mobjShared:                     End Property
Friend Property Get CloseViewMyself() As Boolean:       CloseViewMyself = mbCloseViewMyself:           End Property
Friend Property Let CloseViewMyself(ByVal RHV As Boolean): mbCloseViewMyself = CloseViewMyself:        End Property

Public Property Get CurrentEncounter() As CHCSIIEncounterOps.Encounter:    Set CurrentEncounter = mobjCurrEnc:            End Property

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
    Case Else   '--unknown "view type", not supported--
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
        Set frmEncCosign.ParentObject = Me
        SetParent frmEncCosign.hwnd, vlParenthWnd
        Call AddMenus
        rhMainForm = frmEncCosign.hwnd
        OpenView = True
    Case Else    '--unsupported view "type"--
        OpenView = False
    End Select
    Exit Function
ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description, "EncCosign.OpenView", "Enc", vbCritical)
    OpenView = False
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
    If viViewType = cwiALL_VIEWS Or viViewType = cwiREGULAR_VIEW Then
        If mbMainViewRunning Then
            mbMainViewRunning = False
            '---this indicates that the CW Core is closing the "view", NOT this OLE Server---
            mbCloseViewMyself = False
            Unload frmEncCosign
        End If
    End If
    Call RemoveMenus
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
    On Error GoTo ErrHandler
    '<< SCR 9944 mobjcurrenc is not set if no prev encs are listed
    If mobjCurrEnc Is Nothing Then Exit Sub
    '>>
    Select Case vlMenuId

    Case mlMenuAddNote
    
        'Select the Cosign patient
        Dim lngpatientnumber As Long
        lngpatientnumber = Me.CurrentEncounter.PatientID
        If lngpatientnumber <> 0 Then
            mobjEncounterOps.Comm.Message cwiSELECT_PATIENT, lngpatientnumber & ";", NCID_CW_APPLICATION, mlProgID
        End If
        
        Call mobjCurrEnc.Refresh
        If mobjCurrEnc.OpenEncounter(Me) Then
            Call mobjCurrEnc.AppendNote
            Call mobjCurrEnc.CloseEncounter
            Call frmEncCosign.RefreshDisplay
        End If
    Case mlMenuAmend
        Call mobjCurrEnc.Amend
        Call frmEncCosign.RefreshDisplay(True)
    Case mlMenuSign
        Call mobjCurrEnc.Refresh
        If mobjCurrEnc.OpenEncounter(Me) Then
            Call mobjCurrEnc.SignEncounter
            '--- SCR 62482   Sherry Wang 1/11/2005
            If Not mobjCurrEnc Is Nothing Then
                Call mobjCurrEnc.CloseEncounter
            End If
            Call frmEncCosign.RefreshDisplay(True)
             'SCR-15393  Return user to the Co-Sign screen
            mobjComm.Message cwiBRING_SERVER_TO_FOREGROUND, NCID_COSIGNATURE_APPLICATION, NCID_CW_APPLICATION, NCID_COSIGNATURE_APPLICATION
            'Let core refresh app menu items
            mobjComm.Message cwiGENERIC, "REFRESH", NCID_CW_APPLICATION, NCID_COSIGNATURE_APPLICATION
        End If
    Case mlMenuPrintSF600
        Call mobjCurrEnc.Refresh
        mobjCurrEnc.PrintEncounter False
    Case mlMenuPrintPreviewSF600
        Call mobjCurrEnc.Refresh
        mobjCurrEnc.PrintEncounter True
    End Select
    Exit Sub
ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description, "EncCosign.MenuItemSelected", "EncCosign", vbExclamation)
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
            frmEncCosign.WindowState = vbNormal
            
            '---hide the form to so that annyoing "screen blinking" will not be seen while controls are moved and resized---
            frmEncCosign.Visible = False
            
            '---resize form to fill entire "client area", NOTE: This generates a "resize" event---
            frmEncCosign.Move vlLeft, vlTop, vlWidth, vlHeight
            DoEvents
            frmEncCosign.Visible = True
        Else    '--first "regular" resize, need to perform "initailization" procedures--
            '---reset flag so this code will not be executed again---
            mbFirstRegularResize = False
            frmEncCosign.Display vlLeft, vlTop, vlWidth, vlHeight
        End If
    End Select
    Exit Sub
ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description, "EncCosign.ResizeView", "EncCosign", vbExclamation)
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
    Select Case viMessageType
    Case cwiREFRESH_DATA
        If mbMainViewRunning Then
            frmEncCosign.RefreshDisplay
        End If
    Case cwiPRINT_PREVIEW
        If Not mobjCurrEnc Is Nothing Then
            Call mobjCurrEnc.PrintEncounter(True)
        End If
    Case cwiPRINT_REGULAR_VIEW
        If Not mobjCurrEnc Is Nothing Then
            Call mobjCurrEnc.PrintEncounter(False)
        End If
    Case cwiABOUT 'SCR-20954
        mobjComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_COSIGNATURE_APPLICATION
        Dim objAbout As New CWAbout
        Call objAbout.Display(App, frmEncCosign)
        Set objAbout = Nothing
        '---inform CW that I am done showing a modal form---
        mobjComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_COSIGNATURE_APPLICATION
        
        
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
    Set mobjComm = robjComm
    Set mobjLogon = robjLogon
    Set mobjConfig = robjConfig
    Set mobjPatient = robjPatient
    Set mobjUser = robjUser
    Set mobjEncounterOps = New CHCSIIEncounterOps.EncounterOps
    Call mobjEncounterOps.Initialize(robjComm, robjLogon, robjConfig, robjPatient, robjUser)
End Sub

Private Sub AddMenus()
    On Error Resume Next
    If mlMenuAddNote = 0 Then
        mlMenuAddNote = mobjComm.AddActionMenuItem("Append &Narrative", "new.bmp", mlProgID, "Append a note to the selected encounter.")
    End If
    If mlMenuAmend = 0 Then
        mlMenuAmend = mobjComm.AddActionMenuItem("&Amend Encounter", "edit.bmp", mlProgID, "Amend the selected encounter.")
    End If
    If mlMenuSign = 0 Then
        mlMenuSign = mobjComm.AddActionMenuItem("&Sign Encounter", "dispsign.bmp", mlProgID, "Sign the selected encounter.")
    End If
    
    'Add Print Option SCR-42647
    'Dental providers can print SF600 or SF603a.  All others can only print SF600.
    If mobjUser.HasPrivilegeEx(Priv_Dental_Access, Priv_Read) Then
        If mlMenuPrintSF600 = 0 Then
            mlMenuPrintSF600 = mobjComm.AddPrintMenuItem("&SF603A/SF600", mlProgID, "Print SF600")
        End If
        If mlMenuPrintPreviewSF600 = 0 Then
            mlMenuPrintPreviewSF600 = mobjComm.AddPrintPreviewMenuItem("S&F603A/SF600", mlProgID, "Print SF600")
        End If
    Else
        If mlMenuPrintSF600 = 0 Then
            mlMenuPrintSF600 = mobjComm.AddPrintMenuItem("SF600", mlProgID, "Print SF600")
        End If
        If mlMenuPrintPreviewSF600 = 0 Then
            mlMenuPrintPreviewSF600 = mobjComm.AddPrintPreviewMenuItem("SF600", mlProgID, "Print SF600")
        End If
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
    If Not mlMenuPrintSF600 = 0 Then
        mobjComm.DeleteMenuItem mlMenuPrintSF600
        mlMenuPrintSF600 = 0
    End If
    If Not mlMenuPrintPreviewSF600 = 0 Then
        mobjComm.DeleteMenuItem mlMenuPrintPreviewSF600
        mlMenuPrintPreviewSF600 = 0
    End If
End Sub

Public Sub SetEncounterKey(ByRef StorageKey As String)
    Call frmEncCosign.RefreshDisplay
    Call frmEncCosign.SetEncounterKey(StorageKey)
End Sub

Friend Sub SetCurrentEncounter(RHV As Encounter)
    On Error Resume Next
    If RHV Is Nothing Then
        If mobjCurrEnc Is Nothing Then
            Exit Sub
        End If
    End If
    If Not mobjCurrEnc Is Nothing Then
        RaiseEvent CloseEncounter(mobjCurrEnc)
    End If
    Set mobjCurrEnc = RHV
    Call mobjComm.UpdateMenuItem(mlMenuAddNote, "DISABLE")
    Call mobjComm.UpdateMenuItem(mlMenuAmend, "DISABLE")
    Call mobjComm.UpdateMenuItem(mlMenuSign, "DISABLE")
    If Not mobjCurrEnc Is Nothing Then
        With mobjCurrEnc
            If .Status <> PatientLeftWithoutBeingSeen Then
                If .Status = NeedsCoSignature _
                And mobjUser.HasPrivilegeEx(Priv_Current_Encounter, Priv_Sign) Then
                    Call mobjComm.UpdateMenuItem(mlMenuSign, "ENABLE")
                End If
                If .UserIsPrimaryProvider Or .UserIsCosigner Then
                    Call mobjComm.UpdateMenuItem(mlMenuAmend, "ENABLE")
                End If
                If mobjUser.HasPrivilegeEx(SectionEnumToSecurityName(eNotes), Priv_Write) _
                And mobjUser.HasPrivilegeEx(Priv_Current_Encounter, Priv_Write) Then
                    Call mobjComm.UpdateMenuItem(mlMenuAddNote, "ENABLE")
                End If
            End If
        End With
        RaiseEvent OpenEncounter(mobjCurrEnc)
    End If
End Sub

Private Sub Class_Initialize()
#If DEBUGON Then
    Set DL = GetObject("", "dlog.cdebuglog")
    DL.debugStart
    Dim DR As DebugRec
    If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Class_Initialize", "Begin", DR
#End If
    '---create shared object used to show VB errors, etc.---
    Set mobjShared = New CWShared
    '---set the help file for this application---
    App.HelpFile = mobjShared.CWHelpFile
    Call InitializeGlobalStrings
#If DEBUGON Then
    If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Class_Initialize", "End", DR
#End If
End Sub

Private Sub Class_Terminate()
#If DEBUGON Then
    Dim DR As DebugRec
    If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Class_Terminate", "Begin", DR
#End If
    '---release reference to objects---
    Set mobjShared = Nothing
    Set mobjComm = Nothing
    Set mobjLogon = Nothing
    Set mobjConfig = Nothing
    Set mobjPatient = Nothing
    Set mobjUser = Nothing
    Set mobjCurrEnc = Nothing
#If DEBUGON Then
    If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Class_Terminate", "End", DR
    Set DL = Nothing
#End If
End Sub
