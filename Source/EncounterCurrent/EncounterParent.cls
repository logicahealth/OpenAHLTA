VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EncounterParent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "The main controing class for Encounter Summary."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"EncProperties"
Attribute VB_Ext_KEY = "Member1" ,"EncRTF"
Attribute VB_Ext_KEY = "Member2" ,"EncProperties"


Option Explicit

Implements ICHCSII_CoreAPI2                  '-- implements the security obj
Implements I_RSTransfer

'Inpatient functionality (used in AdmitPatient)
Private Const NCID_Admitted As String = "204849" '"208849"  'SCR#43518 'Came from MMMHISDisposition.modDisposition
Private Const NCID_Admission_HP_Note As String = "15639660" ' "15119311" 'SCR 134416, JQ

Public Event EnctrErr(ErrNum As Integer, ErrStr As String)
Public Event CurrEncStatusChanged(eStatus As EncStatus)
Public Event CreatedEncounter(oEnc As Encounter)
Public Event OpenEncounter(oEnc As Encounter)
Public Event CloseEncounter(oEnc As Encounter)
Private mobjEncounterOps            As CHCSIIEncounterOps.EncounterOps
Private WithEvents mobjCurrEnc      As CHCSIIEncounterOps.Encounter     '-- current active encounter
Attribute mobjCurrEnc.VB_VarHelpID = -1
Private mcolMenus                   As Collection   '-- menus I've put on the screen
Private mobjUser                    As ICHCSII_User '-- The Security User Object
Private mobjPatient                 As Object       '-- patient object from core
Private mobjLogonEnvironment        As Logon        '-- logon env obj from core
Public mobjComm                     As MMMHISComm.Comm '-- comm obj from core
Private mobjConfig                  As MMMHISConfiguration.Config
Private mobjShared                  As MMMHISShared.CWShared
Private mobjOpenChildModule         As Object
Private mlProgID                    As Long         '-- my prog id from core
Private mcolChildMenus              As Collection
Private mlMenuRefresh               As Long
Private mlMenuNote                  As Long
Private mlMenuProviders             As Long
Private mlMenuTemplates             As Long
Private mlMenuGoalSummary           As Long
Private mlMenuSign                  As Long
Private mlMenuCreateTemplate        As Long
Private mlMenuPrintSF600            As Long
Private mlMenuPrintPreviewSF600     As Long
Private mlMenuRunCoupler            As Long
'Private mlTestApp                   As Long
Private bFirstPatientSummaryResize  As Boolean
Private mbMainViewRunning           As Boolean
Private bPatientSummaryViewRunning  As Boolean
Private mbQETelconRunning           As Boolean
Private m_bInitializing As Boolean
Private m_bClosing As Boolean
Private m_iDebugCount As Integer
Private m_lPatientUnitNumber As Long
Private m_lFacilityNCID As Long
Private m_lEncounterNumber As Long
Private m_bEncSelectCanceled As Boolean
Private m_oEncParentForm As frmEncParent
Private m_oEncSelectForm As frmEncSelect
Private mobjAudit As Audit
Private mbSuppressEvents            As Boolean
Private mbLaunchSucceeded           As Boolean

'-- implements the dlog.dll timing module
Private Const Exe = "EnctrCurrent"
Private Const Module = "EncounterParent.cls"
Private Const ExeModule = Exe & "." & Module

#If DEBUGON Then
    Private DR As DebugRec
    Private DL As cDebugLog
#End If

Public Sub CopyForwardEncounter(PatientUnitNumberFrom As String, FacilityNCIDFrom As String, EncounterNumberFrom As String)
    
    'This copies forward the specified encounter to the current encounter.
    'Should this be in PAD?
    On Error GoTo ErrHnd

    Dim oEncParent As CHCSIIEncounterCurrent.EncounterParent  'Object 'EncounterParent  '-- reference to parent of calling class
    Dim oEncCurrent As CHCSIIEncounterOps.Encounter 'Encounter
    Dim objTemplates As Object
    Dim oTemplate As Object
    Dim oNewTemplate As Object
    Dim oSection As SectionParent
    Dim oEncounterFrom As Encounter
    Dim lSelectedrow As Long
    Dim lScreenMousepointer As Long

    Dim sPatientID As String
    Dim sEncounterID As String
    Dim sFacilityNCID As String

    lScreenMousepointer = Screen.MousePointer
    Screen.MousePointer = vbHourglass

    Set oEncParent = mobjComm.InitializeOLEServer(NCID_ENCOUNTER_APPLICATION)

    If oEncParent Is Nothing Then
        MsgBxARMd "Cannot initialize the encounter module to copy forward the selected encounter.", vbOKOnly + vbInformation, "Previous Encounters"
        GoTo CleanUpandExit
    ElseIf oEncParent.CurrentEncounter Is Nothing Then
        mobjComm.Message cwiSTART_OLE_SERVER, NCID_ENCOUNTER_APPLICATION, NCID_CW_APPLICATION, NCID_PREVIOUS_ENCOUNTERS_APPLICATION
        If oEncParent.CurrentEncounter Is Nothing Then
            MsgBxARMd "No encounter open to copy forward the selected encounter to", vbOKOnly + vbInformation, "Previous Encounters"
            GoTo CleanUpandExit
        End If
    End If

    Set oEncCurrent = oEncParent.CurrentEncounter
    Set objTemplates = mobjComm.InitializeOLEServer(NCID_TEMPLATE_MANAGEMENT)

    'Set oEncounterFrom = mobjEncounters(oListItem.Key)
'    sPatientID = CStr(oEncounterFrom.PatientID)
'    sEncounterID = CStr(oEncounterFrom.EncounterID)
'    sFacilityNCID = oEncounterFrom.FacilityNCID

    Set oEncounterFrom = oEncParent.GetEncounter(PatientUnitNumberFrom, FacilityNCIDFrom, EncounterNumberFrom)
    Set oTemplate = objTemplates.GetEncounterAsTemplate(oEncounterFrom, True)
    If oNewTemplate Is Nothing Then
       Set oNewTemplate = oTemplate
    Else
        oNewTemplate.Import oTemplate.Export
    End If

    If Not oNewTemplate Is Nothing Then
        oEncCurrent.TemplateData = oNewTemplate.Export
        oNewTemplate.EnableCopyForward False
    End If

    'Call goBB.Comm.Message(cwiCLOSE_OLE_SERVER, NCID_CW_APPLICATION, NCID_PREVIOUS_ENCOUNTERS_APPLICATION, NCID_PREVIOUS_ENCOUNTERS_APPLICATION)

CleanUpandExit:
    Set oEncParent = Nothing
    Set oEncCurrent = Nothing
    Set objTemplates = Nothing
    Set oTemplate = Nothing
    Set oSection = Nothing
    Set oEncounterFrom = Nothing

    Screen.MousePointer = lScreenMousepointer

Exit Sub

ErrHnd:
    mobjShared.ShowVBError Err.Number, Err.Description, "InpatientDocs.CopyForwardEncounter", "CHCSIIEncounterPrevious", vbInformation
    GoTo CleanUpandExit
    
End Sub






Public Property Get MainViewRunning() As Boolean:           MainViewRunning = mbMainViewRunning:          End Property
Public Property Get CWComm() As Object:                     Set CWComm = mobjComm:                       End Property
Public Property Get Config() As Object:                     Set Config = mobjConfig:                     End Property
Public Property Get CWShared() As Object:                   Set CWShared = mobjShared:                   End Property
Public Property Get EncounterOps() As EncounterOps:         Set EncounterOps = mobjEncounterOps:         End Property
Public Property Get hwnd() As Long:                         hwnd = frmEncParent.hwnd:                    End Property
Public Property Get LogonEnvironment() As Object:           Set LogonEnvironment = mobjLogonEnvironment: End Property
Public Property Get Menus() As Collection:                  Set Menus = mcolMenus:                       End Property
Public Property Get Patient() As Object:                    Set Patient = mobjPatient:                   End Property
Public Property Get ProgID() As Long:                       ProgID = mlProgID:                           End Property
Public Property Get User() As ICHCSII_User:                 Set User = mobjUser:                         End Property
Public Property Get CurrentEncounter() As CHCSIIEncounterOps.Encounter:        Set CurrentEncounter = mobjCurrEnc:          End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' The modal EncSelect dialog causes recursion            '
' that affects the refreshing of the encounter           '
' and the encounter parent form.  The attempted          '
' solution to this issue is to rework the                '
' EncSelect form so that it is simply a container.       '
' Processing will occur in a functionality class.        '
' The following properties are required for this         '
' solution:                                              '
Friend Property Get PatientUnitNumber() As Long          '
    PatientUnitNumber = m_lPatientUnitNumber             '''''''''''''''''''''
End Property                                                                 '
Friend Property Let PatientUnitNumber(ByVal p_lPatientUnitNumber As Long)    '
    m_lPatientUnitNumber = p_lPatientUnitNumber                              '
End Property                                                                 '
Friend Property Get FacilityNCID() As Long                                   '
    FacilityNCID = m_lFacilityNCID                                           '
End Property                                                                 '
Friend Property Let FacilityNCID(ByVal p_lFacilityNCID As Long)              '
    m_lFacilityNCID = p_lFacilityNCID                                        '
End Property                                                                 '
Friend Property Get EncounterNumber() As Long                                '
    EncounterNumber = m_lEncounterNumber                                     '
End Property                                                                 '
Friend Property Let EncounterNumber(ByVal p_lEncounterNumber As Long)        '
    m_lEncounterNumber = p_lEncounterNumber                                  '
End Property                                                                 '
Friend Property Get EncSelectCanceled() As Boolean                           '
    EncSelectCanceled = m_bEncSelectCanceled                                 '
End Property                                                                 '
Friend Property Let EncSelectCanceled(ByVal p_bEncSelectCanceled As Boolean) '
    m_bEncSelectCanceled = p_bEncSelectCanceled                              '
End Property                                                                 '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''
' SCR 6685 Out of body.  We need to know     '
' if we are initializing or not to avoid     '
' a redundant refresh of the encounter       '
' form itself.                               '
'''''''''''''''''''''''''''''''''''''''''''''''''                                                    '
Public Property Get Initializing() As Boolean   '''''''''''''''''''''''
    Initializing = m_bInitializing                                    '
End Property                                                          '
Public Property Let Initializing(ByVal p_bInitializing As Boolean)    '
    m_bInitializing = p_bInitializing                                 '
End Property                                                          '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Property Let MainViewRunning(ByVal RHV As Boolean)
    mbMainViewRunning = False
End Property

'Friend Property Set CurrentEncounter(ByVal RHV As Encounter)
'    ' this is for use only by frmEncSelect
'    Set mobjCurrEnc = RHV
'End Property
Private Property Get EncSelectForm() As frmEncSelect
    Set EncSelectForm = m_oEncSelectForm
End Property
Private Property Set EncSelectForm(ByRef p_oEncSelectForm As frmEncSelect)
    Set m_oEncSelectForm = p_oEncSelectForm
End Property

Private Property Get EncParentForm() As frmEncParent
    Set EncParentForm = m_oEncParentForm
End Property
Private Property Set EncParentForm(ByRef p_oEncParentForm As frmEncParent)
    Set m_oEncParentForm = p_oEncParentForm
End Property

'--- SCR 81705  Sherry Wang   11/9/2005
Public Property Get LaunchSucceeded() As Boolean
    LaunchSucceeded = mbLaunchSucceeded
End Property


'=========== End of standard property exposure code ===========================

Public Function GetEncounter(ByVal PatientUnitNumber As String, _
                             ByVal FacilityNCID As String, _
                             ByVal EncounterNumber As String) As Encounter
Attribute GetEncounter.VB_Description = "Returns a single (existing) encounter based on patient and encounter numbers. "
    Set GetEncounter = mobjEncounterOps.GetEncounter(PatientUnitNumber, FacilityNCID, EncounterNumber)
End Function

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
Attribute QueryView.VB_Description = "As mandated by CW"
    #If DEBUGON Then
        Dim DR As DebugRec
        If Not DL Is Nothing Then DL.DebugLog Exe, Module, "QueryView", "Begin", DR
    #End If
    If Not DebugMode Then On Error GoTo ErrHandler
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
        QueryView = True
        ' Do I support a properties view?
    Case Else   '--unknown "view type", not supported--
        QueryView = False
    End Select
    #If DEBUGON Then
        If Not DL Is Nothing Then DL.DebugLog Exe, Module, "QueryView", "End", DR
    #End If
    Exit Function
ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description, "QueryView", ExeModule, vbExclamation)
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
' Comment: Returning False will cause core to not continue loading the module.
'**************************************************************************************************
Public Function OpenView(ByVal viViewType As Integer, _
                         ByVal vlParenthWnd As Long, _
                               rhMainForm As Long) As Boolean
Attribute OpenView.VB_Description = "As mandated by CW"
    Dim msg As String    '- message string holder
    Dim s   As String    '- generic string
    Dim bOpenView As Boolean
    
    If Not DebugMode Then On Error GoTo ErrHandler
    If mobjUser Is Nothing Then msg = msg & Chr(10) & "Security User Object"
    If mobjPatient Is Nothing Then msg = msg & Chr(10) & "Patient Object"
    If mobjConfig Is Nothing Then msg = msg & Chr(10) & "Configuration Object"
    If mobjLogonEnvironment Is Nothing Then msg = msg & Chr(10) & "Logon Object"
    
    
    
    If mobjComm Is Nothing Then msg = msg & Chr(10) & "Core Communication Object"
    If msg <> "" Then
        msg = Mid$(msg, 2)   'drop the leading 'lf'
        MsgBxARMd "The following object(s) were not handed-off to the Encounter Object:" & Chr(10) & Chr(10) _
           & msg & Chr(10) & Chr(10) & "Encounter cannot continue to load and is shutting down.", vbCritical, "Open Encounter Error"
        Exit Function
    End If
    '-- are you allowed in here?
    ' SCR #11001 This check doesn't work because core ignores OpenView = False.
'    If Not mobjUser.HasPrivilegeEx(Priv_Current_Encounter, Priv_Read) Then
'        MsgBxARMd "You do not have permission to view this information." & Chr$(10) _
'          & "Please contact your security administrator if you believe you have received this message in error.", vbExclamation
'        Exit Function
'    End If

    On Error Resume Next
    Err.Clear
    s = mobjUser.GetAttribute("USER_Ncid")
    If Err Then
        MsgBxARMd "Your Identification was not found in the security file. Please notify your security administrator.", vbCritical, "Security Error"
        Exit Function
        Err.Clear
    End If
    If Not DebugMode Then On Error GoTo ErrHandler
    '---perform actions based on the "view type"---
    Select Case viViewType
    Case cwiREGULAR_VIEW
        If Not mbMainViewRunning Then
'            OpenView = refeshview()
            Initializing = True
            OpenView = OpenRegularView()
            If OpenView Then 'And Not mbMainViewRunning Then
                mbMainViewRunning = True
                'mbFirstRegularResize = True
                frmEncParent.ParentObject = Me
                rhMainForm = frmEncParent.hwnd
                'Here we initialize the form need refresh should be set here or in the form....
                SetParent frmEncParent.hwnd, vlParenthWnd
            Else
                'MsgBox "Error: Could not open the encounter window. EncounterParent.OpenView"
            End If
        End If
    Case cwiPATIENT_SUMMARY_VIEW
        SetParent frmEncSOC.hwnd, vlParenthWnd
        frmEncSOC.ParentObject = Me
        Set frmEncSOC.Encounters = mobjEncounterOps.GetPreviousEncounters(mobjPatient.UnitNumber)
        rhMainForm = frmEncSOC.hwnd
        bFirstPatientSummaryResize = True
        bPatientSummaryViewRunning = True
        OpenView = True
    Case Else    '--unsupported view "type"--
        OpenView = False
    End Select
    Exit Function
ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description, "EncounterParent.OpenView", "Enc", vbCritical)
    OpenView = False
End Function

Private Function OpenRegularView() As Boolean
    Dim objEncounterPrev As Object
    Dim sMsgText As String
    Dim buf() As String
    Dim EncType As EncType
    Dim objAppts As Object
    Dim objCHCSAppt    As Object
    Dim bOpenPrevEncounters As Boolean
    
    If Not DebugMode Then On Error GoTo ErrHandler
    If mobjCurrEnc Is Nothing Then
        If mobjPatient Is Nothing Then
            Exit Function
        End If
       If Not mobjShared.IsAppMode(modeCHCSI_GUI) Then  '<SCR SCR 36493
            If mobjPatient.UnitNumber = 0 Or mobjPatient.UnitNumber = "" Then
                Exit Function
            End If
            If Not OpenViewGetEncounter Then
                Exit Function
            End If
        Else
            'Send Generic Msg to Appts to obtain Appt IEN, Appt Type and Appt DTS SCR-31496
            sMsgText = "CHCS_GUI_APPT_GRIDINFO"
            mobjComm.Message cwiGENERIC, sMsgText, NCID_CLINIC_SCHEDULE_APPLICATION, NCID_ENCOUNTER_APPLICATION
            '    0       1          2         3          4                5                     6               7           8
            'ApptType|ApptIEN|ApptDateTime|ClinicName|ClinicIEN|Inpatient_MEPRS_Code|CStr(enumEncounterType)|WorkLoad|Clinician NCID
            If Len(sMsgText) Then
                buf = Split(sMsgText, "|")
                
                Dim dApptDate As Date
                
                dApptDate = buf(2)
                
                EncType = Outpatient
                'Set correct EncType
                Select Case buf(1)
                    Case "T-CON*", "T-CON", "TCON", "TEL-CON", "OCC-SVR"
                        EncType = TeleConsult
                    Case Else
                        EncType = CLng(buf(6))
                End Select
    
                'Need to check if appointment is completed (or cancelled) 'SCR#45270
                Set objAppts = mobjComm.InitializeOLEServer(NCID_CLINIC_SCHEDULE_APPLICATION)
    
                'Set UseCache parameter = False
                Set objCHCSAppt = objAppts.AppointmentOps.GetAppointment_CHCS_GUI(buf(0), False)
                
                If StrComp(UCase(objCHCSAppt.ADM_ENCOUNTER), "COMPLETE", vbTextCompare) <> 0 Then
                    'Appointment Not Complete
                    'Check in Appt SCR#46153
                    mobjComm.Message cwiCHECKIN_APPOINTMENT, buf(0), NCID_CLINIC_SCHEDULE_APPLICATION, NCID_ENCOUNTER_APPLICATION
    
                    'Get Clinic NCID from appoitments  SCR-31237
                    Set mobjCurrEnc = CreateEncounter(buf(0), buf(1), EncType, mobjLogonEnvironment.FacilityNCID, buf(4), buf(8), , , dApptDate)
                Else
                    'Appointment completed but need to check Enc Status (may be Updating) 'SCR#45270
                    Set mobjCurrEnc = objAppts.AppointmentOps.GetEncounter_CHCS_GUI(objCHCSAppt.ApptID, objCHCSAppt.FacilityNCID, objCHCSAppt.PatientIEN)
                    If mobjCurrEnc Is Nothing Then
                        bOpenPrevEncounters = True
                    Else
                        If objCHCSAppt.ApptStatus = "LWOBS" Or objCHCSAppt.ApptStatus = "CANCEL" Or objCHCSAppt.ApptStatus = "NO-SHOW" Then
                            bOpenPrevEncounters = True
                        End If
                    End If
                End If
            
                'Set IEN Data
                If Not (mobjCurrEnc Is Nothing) Then
                  mobjCurrEnc.ApptIEN = buf(0)
                  mobjCurrEnc.Patient_IEN = mobjPatient.IEN_number
                End If
            Else
                'No Appt selected then do not create encounter SCR-31600
                Exit Function
            End If
        End If
    Else
        'In theater mode RFV RTF is added after encounter is created so we need to retrieve the RTF before displaying
        'If using encounter from cache
        If mobjShared.IsAppMode(modeTheater) Then
            mobjCurrEnc.Refresh 'SCR#43829
        End If

    End If
    If Not bOpenPrevEncounters Then
        If mobjCurrEnc Is Nothing Then
            Set mobjCurrEnc = mobjEncounterOps.GetEncounter(mobjCurrEnc.PatientID, mobjCurrEnc.FacilityNCID, mobjCurrEnc.EncounterID)
        End If
        If mobjCurrEnc Is Nothing Then
            Exit Function
        End If
        If Not mobjCurrEnc.OpenEncounter(Me) Then
            Exit Function
        End If
        Select Case mobjCurrEnc.Status
        Case complete, Updated, PatientLeftWithoutBeingSeen, _
            CancelByPatient, CancelByFacility, NoShow ' , NeedsCoSignature (Removed) - SCR-14496
        '-- since I don't show these in the 'enc select' window, I must be being asked programmatically to display them.
        '   Send the user to the prev enc module.
            bOpenPrevEncounters = True
        End Select
    End If
    
    If bOpenPrevEncounters Then
        Set objEncounterPrev = mobjComm.InitializeOLEServer(NCID_PREVIOUS_ENCOUNTERS_APPLICATION)
        If Not objEncounterPrev Is Nothing Then
            mobjComm.Message cwiSTART_OLE_SERVER, NCID_PREVIOUS_ENCOUNTERS_APPLICATION, NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION
            If Not mobjCurrEnc Is Nothing Then
                Call objEncounterPrev.SetEncounterKey(mobjCurrEnc.StorageKey)
            End If
        End If
        'need to reset in case user tries to open a different encounter via folders
        Set mobjCurrEnc = Nothing
        OpenRegularView = False
        Exit Function
    End If
    
    Call AddMenus
    OpenRegularView = True
    RaiseEvent OpenEncounter(mobjCurrEnc)
    
    'Send a message to Patient Demographics to Refresh Data SCR-13621
    If mobjComm.Message(cwiOLE_SERVER_ACTIVE, _
                                        NCID_DEMOGRAPHICS_APPLICATION, _
                                        NCID_CW_APPLICATION, _
                                        NCID_ENCOUNTER_APPLICATION) Then

       mobjComm.Message cwiREFRESH_DATA, _
                                        "", _
                                        NCID_DEMOGRAPHICS_APPLICATION, _
                                        NCID_ENCOUNTER_APPLICATION, False

    End If
    

    Exit Function
ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description, "EncounterParent.OpenRegularView", mbMainViewRunning, vbCritical)
End Function

'----------------------------------------------------------------------------
'- Purpose: get the encounter to be displayed into the mobjCurrEnc object and prepare to display it
'
' Input: ShowForm - Should we show the "encounter select" form to the user, or is there enough info
'                   to know which one enc to display?
' Output - True on success
'----------------------------------------------------------------------------
Private Function OpenViewGetEncounter() As Boolean
    
    If Not DebugMode Then On Error GoTo ErrHandler
    
    'SCR-ARM-28779 RWalker 3/07/03 (formerly SCR 31247)
    If Not goARM Is Nothing Then goARM.StartCmplxTx gsARMTx_OpenExistingEncounter, , , , , , , , eAnyFamily
    
    Set EncSelectForm = frmEncSelect
    
    EncSelectForm.ParentObject = Me
    EncSelectForm.Encounters = mobjEncounterOps.GetOpenEncounters(mobjPatient.UnitNumber)
    EncSelectForm.Top = (Screen.Height - frmEncSelect.Height) / 3
    EncSelectForm.Left = (Screen.Width - frmEncSelect.Width) / 2
    
    mobjComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION
    EncSelectForm.Show vbModal
   ' Testing, just for now.
    mobjComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION
    
    If Not EncSelectCanceled Then
        SetCurrentEncounter PatientUnitNumber, FacilityNCID, EncounterNumber
    End If
    
    'SCR-ARM-28779 RWalker 2/11/2003 (formerly SCR 31247)
    If Not goARM Is Nothing Then goARM.StopCmplxTx gsARMTx_OpenExistingEncounter, enmARMTxStatusOk, enmARMStop_MeOnly
    
    If mobjCurrEnc Is Nothing Then
        Set EncSelectForm = Nothing
        OpenViewGetEncounter = False
        Exit Function
    End If
    
    OpenViewGetEncounter = True
    Exit Function
    
ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description, "EncounterParent.OpenViewGetEncounter", "Enc", vbCritical)
End Function


'*******************************************************************************************
'  Description: This method is called by the CW Core as a close command
'
'  Parameters:
'    viViewType - Contains the "type" of view that the CW Core is closing.
'*******************************************************************************************
Public Sub CloseView(ByVal viViewType As Integer)
Attribute CloseView.VB_Description = "CW's Mandated method to direct the closing of the Enctr module"
    If Not DebugMode Then On Error GoTo ErrHandler
    If viViewType = cwiALL_VIEWS Or viViewType = cwiREGULAR_VIEW Then
        If mbMainViewRunning Then
            If Not frmEncParent.Enabled Then
                Exit Sub
            End If
            
            'Call SignAndSave Here
            If mobjShared.IsAppMode(modeCHCSI_GUI) Then  '<SCR SCR 36493
            
            End If
            If CloseCurrentEncounter Then
                Set frmEncParent = Nothing
                ' Testing....
                mbMainViewRunning = False
            End If
            'NeedsRefresh = True
        End If
    End If
    If viViewType = cwiALL_VIEWS Or viViewType = cwiPATIENT_SUMMARY_VIEW Then
        If bPatientSummaryViewRunning Then
            bPatientSummaryViewRunning = False
            frmEncSOC.mbCloseMyself = False
            Unload frmEncSOC
        End If
    End If
    Exit Sub
ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description, "EncounterParent.CloseView", "EncounterParent", vbExclamation)
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
Attribute MenuItemSelected.VB_Description = "As mandated by CW"
    Dim objTemplates As Object
    Dim sMsg As String
    
    On Error GoTo ErrHandler
    
    If mbSuppressEvents Then Exit Sub

    mbSuppressEvents = True
    
    Select Case vlMenuId
    Case mlMenuRefresh
        Call frmEncParent.RefreshDisplay(True)
    Case mlMenuNote
        mobjCurrEnc.OpenSection eNotes
    Case mlMenuProviders
        mobjCurrEnc.AddProvider
    Case mlMenuTemplates
        Set objTemplates = mobjComm.InitializeOLEServer(NCID_TEMPLATE_MANAGEMENT)
        If Not objTemplates Is Nothing Then
            Call objTemplates.SelectEncounterTemplate
            Set objTemplates = Nothing
        End If
    Case mlMenuGoalSummary
        Call mobjComm.Message(cwiGENERIC, "show form", NCID_GOAL_SUMMARY, NCID_ENCOUNTER_APPLICATION)
    Case mlMenuSign
        'Disable core and encounter SCR-25974
        sMsg = "Call Show Modal"
        mobjComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION
        
        '--- SCR 22247 flag ARM to time sign from Enc only (not children)  MSolano 4/24/02
        mobjCurrEnc.SignByEncParent = True
        
        sMsg = "Call Sign Encounter"
        If mobjCurrEnc.SignEncounter = True Then
'           Can't put this here because the encounter can be signed from disposition.
'            If Len(mobjComm.CmdLineSwitch("INPATIENT2")) > 0 And mobjShared.IsAppMode(modeTheater) = True Then 'SF RQT 56543, 56541, 56542, 61021
'                If mobjCurrEnc.EncounterType <> InpatientNote Then
'                    'Determine if the disposition is "Admitted"
'                    If mobjCurrEnc.DispositionNCID = NCID_Admitted Then
'                        AdmitPatient
'                    End If
'                End If
'            End If
        End If
        
        sMsg = "Call Done Show Modal"
        mobjComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION
        
        '--- SCR-ARM-32970/42483 RWalker 10/29/03 Result of transaction is dependent on a successful
        'update, which is indicated by the current encounter object being set to nothing
        If mobjCurrEnc Is Nothing Then
          goARM.StopCmplxTx NCID_ENCOUNTER_APPLICATION & gsARMTx_Is_Encounter_Caption_Sign & CStr(vlMenuId), _
                            enmARMTxStatusOk, enmARMStop_MeOnly, App.EXEName
        Else
          goARM.StopCmplxTx NCID_ENCOUNTER_APPLICATION & gsARMTx_Is_Encounter_Caption_Sign & CStr(vlMenuId), _
                            enmARMTxStatusFailed, enmARMStop_MeOnly, App.EXEName
        End If

    Case mlMenuCreateTemplate
        Set objTemplates = mobjComm.InitializeOLEServer(NCID_TEMPLATE_MANAGEMENT)
        If Not objTemplates Is Nothing Then
            Call objTemplates.CreateNewFromEncounter(mobjCurrEnc)
            Set objTemplates = Nothing
        End If
    Case mlMenuPrintSF600
        'KDunne 6/18/2002
        'SCR 24688 - The Encounter object contains the logic  necessary for printing Encounters.
        'Disable Print options when printing 'SCR-42777
        mobjComm.UpdateMenuItem mlMenuPrintSF600, "DISABLE"
        mobjComm.UpdateMenuItem mlMenuPrintPreviewSF600, "DISABLE"
        Call mobjCurrEnc.PrintEncounter(False)
        mobjComm.UpdateMenuItem mlMenuPrintPreviewSF600, "ENABLE"
        mobjComm.UpdateMenuItem mlMenuPrintSF600, "ENABLE"
    Case mlMenuPrintPreviewSF600
        'Disable Print options when printing 'SCR-42777
        mobjComm.UpdateMenuItem mlMenuPrintPreviewSF600, "DISABLE"
        mobjComm.UpdateMenuItem mlMenuPrintSF600, "DISABLE"
        Call mobjCurrEnc.PrintEncounter(True) 'scr 24688
        mobjComm.UpdateMenuItem mlMenuPrintPreviewSF600, "ENABLE"
        mobjComm.UpdateMenuItem mlMenuPrintSF600, "ENABLE"
    Case mlMenuRunCoupler
        Dim objCoupler As Object
        'Set objCoupler = mobjComm.InitializeOLEServer(NCID_HEAR_APPLICATION)
        'Call objCoupler.RunEncounterCoupler(mobjCurrEnc.GetCouplerData.mcolListRecords) 'SCR-30929
        mobjComm.Message cwiSTART_OLE_SERVER, NCID_HEAR_APPLICATION, NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION 'SCR-38194
'    Case mlTestApp
'        Dim objDisp As Object
'        Set objDisp = mobjComm.InitializeOLEServer(NCID_DISPOSITION_APPLICATION)
'        If Not objDisp Is Nothing Then
'            objDisp.InitDisplayMode 2, 999999, Now, Now
'            mobjComm.Message cwiSTART_OLE_SERVER, NCID_DISPOSITION_APPLICATION, NCID_CW_APPLICATION, NCID_INPATIENT_APPLICATION
'        End If
'        Set objDisp = Nothing

    Case Else
        Call mobjCurrEnc.MenuItemSelected(vlMenuId)
    End Select
    
    mbSuppressEvents = False
    
    Exit Sub
    
ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description, "EncounterParent.MenuItemSelected", "EncounterParent", vbExclamation)
    mbSuppressEvents = False
    
    Exit Sub
    Resume
End Sub
'Private Function AdmitPatient() As Boolean

'    Dim oPAD As Object
'    Dim oInpatient As Object
'    Dim oDocuments As Object
''    Dim objEnc As CHCSIIEncounterOps.Encounter
'    Dim sErrMsg As String
'    Dim sEncounterNumberFrom As String
'    Dim lMousePointer As Long
'
'    On Error GoTo ErrHandler
'
'    lMousePointer = Screen.MousePointer
'    Screen.MousePointer = vbHourglass
'
'    sEncounterNumberFrom = mobjCurrEnc.EncounterID
'
''    LogAction "AdmitPatient - Start"
'
''    Set objEnc = gobjEncParent.CurrentEncounter
'    mobjComm.Message cwiSTART_OLE_SERVER, NCID_INPATIENT_APPLICATION, NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION
'    Set oPAD = mobjComm.InitializeOLEServer(NCID_INPATIENT_APPLICATION)
'    If Not (oPAD Is Nothing) Then 'SCR#48990
'        If oPAD.AdmitPatient(mobjPatient, mobjLogonEnvironment.UserNcid, mobjCurrEnc.PrimaryProvider.NCID, Now, mobjCurrEnc.EncounterID, sErrMsg) = False Then
'            If Len(sErrMsg) Then
'                'Display all other err msgs 'SCR#48605
'                'If StrComp(sErrMsg, "Patient is already Admitted", vbTextCompare) <> 0 Then
'                    MsgBxARMd sErrMsg, vbInformation, "Encounter" '48248
'                'End If
'            End If
'            CloseCurrentEncounter
'        Else
'            'prompt for completing admission...
'            CloseCurrentEncounter
'            If MsgBxARMd("Do you want to complete an Admission History and Physical now?", vbYesNo) = vbYes Then
'                'Create the note
'                Set oInpatient = oPAD.SelectedPatients.getInPatientByUnitNumber(mobjPatient.UnitNumber)
'                'Select that patient
'                Set oPAD.SelectedPatients.Inpatient = oInpatient
'                'Add the note
'                oPAD.AddNote oInpatient, NCID_Admission_HP_Note, sEncounterNumberFrom
'                DoEvents
'
'                'Debug.Assert False
'                'copy forward the encounter
'                'Should this be in PAD?
'    '            Set oDocuments = mobjComm.InitializeOLEServer(NCID_INPATIENT_DOCUMENTS_APPLICATION)
'    '            oDocuments.CopyForwardEncounter mobjPatient.UnitNumber, mobjLogonEnvironment.FacilityNCID, sEncounterNumberFrom
'            End If
'        End If
'    End If
'
'    'LogAction "AdmitPatient - End"
'
'    Set oPAD = Nothing
''    Set objEnc = Nothing
'
'    Screen.MousePointer = lMousePointer
'
'    Exit Function
'
'ErrHandler:
'    MsgBox Err.Description & " in EncounterParent.AdmitPatient"
    
    
'End Function



'KDunne 6/18/2002
'SCR 24688 - PrintEncounter sub is now obsolete.
'
'Private Sub PrintEncounter(oCurrEnc As Object, pPreview As Boolean)
'
'    Dim objForms As Object
'
'    On Error Resume Next
'
'    'SCR 8902 KDunne 2/26/2002
'    'For Dental Encounters only encounter form is SF603A. For all others,
'    'encounter form is SF600.
'    If mobjCurrEnc.EncounterType = Dental Then
'        Set objForms = CreateObject("CHCSII_Forms.DentalForms")
'        Call objForms.PrintSF603A2(mobjPatient, oCurrEnc, Nothing, Nothing, Nothing, pPreview)
'    Else
'        If objForms Is Nothing Then Set objForms = CreateObject("CHCSII_Forms.Forms")
'        If Not objForms Is Nothing Then
'           objForms.printSF600 oCurrEnc, oCurrEnc.CurrentRTF, pPreview
'        End If
'    End If
'
'    GoTo CleanUpAndExit
'
'ErrHandler:
'    Call mobjShared.ShowVBError(Err.Number, Err.Description, "EncounterParent.PrintEncounters", "Patient summary, Encounter", vbExclamation)
'CleanUpAndExit:
'    Set objForms = Nothing
'End Sub

'**************************************************************************************************************
'  Description: This method gets called when a new patient has been selected from the CW Core.
'
'               NOTE: This is an OPTIONAL method.
'
'  Parameters:
'    01.1.00 WAS: added optioanl ForceRefresh for Dental to allow them to change the encounter and gui
'**************************************************************************************************************
Public Sub PatientSelected()
Attribute PatientSelected.VB_Description = "As mandated by CW"
    #If DEBUGON Then
        Dim DR As DebugRec
        If Not DL Is Nothing Then DL.DebugLog Exe, Module, "PatientSelected", "Begin", DR
    #End If
    
    If Not mobjCurrEnc Is Nothing Then
        If mobjCurrEnc.PatientID <> mobjPatient.UnitNumber Then
            If mbMainViewRunning Then
                If Not mobjShared.NextAntemortem Then
                    Call CloseCurrentEncounter
                End If
            End If
            Set mobjCurrEnc = Nothing
        End If
    End If
    #If DEBUGON Then
        If Not DL Is Nothing Then DL.DebugLog Exe, Module, "PatientSelected", "after CloseCurrentEncounter", DR
    #End If
    
    If bPatientSummaryViewRunning Then
        Set frmEncSOC.Encounters = Nothing
        Set frmEncSOC.Encounters = mobjEncounterOps.GetPreviousEncounters(mobjPatient.UnitNumber)
        frmEncSOC.Display
    End If
    #If DEBUGON Then
        If Not DL Is Nothing Then DL.DebugLog Exe, Module, "PatientSelected", "End", DR
    #End If

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
Public Sub ResizeView(ByVal viViewType As Integer, _
                      ByVal vlLeft As Long, _
                      ByVal vlTop As Long, _
                      ByVal vlWidth As Long, _
                      ByVal vlHeight As Long)
    
    
    ' SCR 8973 and 10712.''''''''''''''''''''''''
    ' For whatever reason the vlTop param       '
    ' is coming in at 0.  This causes part      '
    ' of the rich text to be hidden by another  '
    ' control on the form.                      '
'    '''''''''''''''''''''''''''''''''''''''''''''
    Dim lSummaryTop As Long
    Dim sMsg As String
    
    sMsg = "Set top parameters"
    lSummaryTop = (frmEncParent.vseHeader.Height)

    lSummaryTop = lSummaryTop + 1
    
    frmEncParent.VP1.Top = lSummaryTop
    
    If Not DebugMode Then On Error GoTo ErrHandler
    
    Select Case viViewType
    Case cwiREGULAR_VIEW
        If Initializing Then 'mbFirstRegularResize Then
            '--first "regular" resize, need to perform "initailization" procedures--
            '---reset flag so this code will not be executed again---
            sMsg = "Regular View - Init"
            frmEncParent.Display vlLeft, vlTop, vlWidth, vlHeight
            Initializing = False
        Else
            sMsg = "Regular View - resize"
            ''''''''''''''''''''''''''''''''''''''''''''''
            ' SCR 6685 Out of body.  We need to know     '
            ' if we are orphaned or not to avoid         '
            ' a redundant refresh of the encounter       '
            ' form itself.                               '
            ''''''''''''''''''''''''''''''''''''''''''''''
            If frmEncParent.m_bFormOrphaned Then
                Exit Sub
            End If
            
            ''''''''''''''''''''''''''''''''''''''''''''
            ' SCR 6685                                 '
            ' Testing based on Width and Height.       '
            ' The following are never                  '
            ' in sync so remove from test.             '
            ' frmEncParent.Left = vlLeft _             '
            ' And frmEncParent.Top = vlTop _           '
            ''''''''''''''''''''''''''''''''''''''''''''
            If frmEncParent.Width = vlWidth _
            And frmEncParent.Height = vlHeight Then
                ' an optimization - don't resize if not necessary
                Exit Sub
            End If
            
            On Error Resume Next 'SCR#48616
            
            frmEncParent.Visible = False
            frmEncParent.WindowState = vbNormal
            frmEncParent.Move vlLeft, vlTop, vlWidth, vlHeight
            If Not mobjOpenChildModule Is Nothing Then
                sMsg = "Regular View - resize child"
                Call mobjOpenChildModule.ResizeView(viViewType, vlLeft, vlTop, vlWidth, vlHeight)
            End If
            frmEncParent.Visible = True
            
            On Error GoTo ErrHandler
        End If
    Case cwiPATIENT_SUMMARY_VIEW
        If bFirstPatientSummaryResize Then
            sMsg = "Summary View - Init"
            bFirstPatientSummaryResize = False
            frmEncSOC.Display
            frmEncSOC.WindowState = vbNormal
            frmEncSOC.Visible = False
            frmEncSOC.Move vlLeft, vlTop, vlWidth, vlHeight
            frmEncSOC.Visible = True
        Else
            sMsg = "Summary View - Resize"
            frmEncSOC.WindowState = vbNormal
            frmEncSOC.Visible = False
            frmEncSOC.Move vlLeft, vlTop, vlWidth, vlHeight
            frmEncSOC.Visible = True
        End If
    End Select
    Exit Sub
ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description & " - " & sMsg, "EncounterParent.ResizeView", "EncounterParent", vbExclamation)
End Sub

Private Function ShouldUseQETelcon() As Boolean
    
    'Do not display QET for Medcin Lite SCR-32512
    If mobjShared.IsAppMode(modeCHCSI_GUI) Then  '<SCR SCR 36493
       Exit Function
    End If
    If mobjCurrEnc.UserIsPrimaryProvider _
    And mobjCurrEnc.EncounterType = TeleConsult _
    And Len(Trim(mobjCurrEnc.SectionRTF(eap, True))) < 1 _
    And Len(Trim(mobjCurrEnc.SectionRTF(eDisp, True))) < 1 _
    And Len(Trim(mobjCurrEnc.SORTFOnly)) < 1 Then
    'SO Section RTF may contain Questionnaire RTF ergo only return SO RTF SCR-28020
    'And Len(Trim(mobjCurrEnc.SectionRTF(eSO))) < 1 Then
        ShouldUseQETelcon = True
    End If
End Function

Friend Sub QETelconDone()
    If mbQETelconRunning Then
        mbQETelconRunning = False
        If Not mobjCurrEnc Is Nothing Then
            Call mobjCurrEnc.RefreshDisplay
            Call AddMenus
            frmEncParent.picSummary.Enabled = True
        End If
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
        
    '--- ENTRY POINT TO ARM
    Call ARM_InitAllTransactionsForCurrEncounters   '--- SCR-ARM-16120 MSolano 8/17/01; COMMON TO ALL ARM-SCRs
    
    'Testing
    ' Get a reference to the parent form.
    'Set EncParentForm = frmEncParent
    
    '---set the help file for this application---
    App.HelpFile = mobjShared.CWHelpFile
    Call InitializeGlobalStrings
    #If DEBUGON Then
      If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Class_Initialize", "End", DR
    #End If

End Sub

Private Sub Class_Terminate()
    Set mobjEncounterOps = Nothing
    Set mobjCurrEnc = Nothing
    Set mobjPatient = Nothing
    Set mobjLogonEnvironment = Nothing
    Set mobjComm = Nothing
    Set mobjShared = Nothing
    Set mobjComm = Nothing
    Set mobjConfig = Nothing
    Set mobjUser = Nothing
    Set mobjAudit = Nothing
    
    #If DEBUGON Then
        Dim DR As DebugRec
        If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Class_Terminate", "End", DR
        Set DL = Nothing
    #End If
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


Public Sub Message(ByVal viMessageType As Integer, ByRef vsMessageText As String, ByVal vsSourceServerNCID As String)
Attribute Message.VB_Description = "As mandated by CW"
    Dim buff() As String
    Dim i As Integer
    Dim oSec As SectionParent
    Dim objTemp As Object
    Dim sTemp As String
    Dim sMsg As String
    Dim EncounterInfo As Encounter_Key_Information
    
    
    sMsg = "Type: " & CStr(viMessageType) & " - Text: " & vsMessageText & " - Source: " & vsSourceServerNCID
    
    'Return App NCID to Core to display correct module in title bar SCR-15566
    Select Case viMessageType
        Case cwiCURRENT_APPLICATION_NCID
            If Not mobjCurrEnc Is Nothing Then
                If Not mobjCurrEnc.CurrSection Is Nothing Then
                    If mobjCurrEnc.CurrSection.meSection <> eAutoCite Then   'SCR-35133
                       vsMessageText = mobjCurrEnc.CurrSection.msAppNCID
                    End If
                End If
            End If
    End Select
    
'   Possible cause of TXOLE error SCR-22331
'    If Not frmEncParent.Enabled Then
'        Exit Sub
'    End If
    If Not DebugMode Then On Error GoTo ErrHandler
    Select Case viMessageType
    Case cwiREFRESH
    Case cwiREFRESH_DATA
   '    Added for Theater Release Patch 11/27/02
        If Len(mobjComm.CmdLineSwitch("GEMS_ARMY")) Then
            If vsSourceServerNCID = NCID_BODYIMAGE_APPLICATION Then
                'Verify valid NCID
                If SectionNCIDToEnum(vsMessageText) Then
                    If Not mobjCurrEnc Is Nothing Then 'SCR#47049
                        If mbMainViewRunning Then 'SCR#49640
                            sMsg = "Reload section data"
                            mobjCurrEnc.GEMS_LoadSection (SectionNCIDToEnum(vsMessageText))
                            sMsg = "Refresh w/o loading any data"
                            RefreshDisplay False
                        End If
                    End If
                End If
            End If
        End If

    Case cwiPROPERTIES  '-- show the properties page
        'Vitals has its own properties form SCR-18500
        If Not mobjOpenChildModule Is Nothing Then
            If TypeName(mobjOpenChildModule) = "Vitals" Then
               mobjOpenChildModule.Message cwiPROPERTIES, "", NCID_ENCOUNTER_APPLICATION
               Exit Sub
            End If
        End If
        If Not frmEncParent.mobjParent.CurrentEncounter Is Nothing Then
           sTemp = frmEncParent.mobjParent.CurrentEncounter.TemplateData
           EncounterInfo.ID = CStr(frmEncParent.mobjParent.CurrentEncounter.EncounterID)
           EncounterInfo.Facility = frmEncParent.mobjParent.CurrentEncounter.FacilityNCID
           EncounterInfo.Unit_NUmber = frmEncParent.mobjParent.CurrentEncounter.PatientID
           'Needed to determine if RNDS appt type 'SCR#44005
           EncounterInfo.Category = frmEncParent.mobjParent.CurrentEncounter.Category
           'Add Appt Type to determine if RNDS appt
        End If
        If mobjEncounterOps.EncProperties.Display(EncounterInfo, sTemp) Then
            'Check to see if child is open SCR-19131
            If mobjOpenChildModule Is Nothing Then
                If mobjCurrEnc.SectionIsEditable(eAutoCite) Then
                    Call frmEncParent.cmds_Click(0)
                End If
            Else
                'Tell child to refresh
                mobjOpenChildModule.Message cwiREFRESH, "", NCID_ENCOUNTER_APPLICATION
            End If
        End If
    Case cwiGENERIC       '-- parse the message to see what needs to be done
        Call MessageGeneric(vsMessageText)
    Case cwiPRINT_REGULAR_VIEW
        If Not mobjCurrEnc Is Nothing Then
            mobjCurrEnc.PrintEncounter False
        End If
    Case cwiPRINT_PREVIEW
        If Not mobjCurrEnc Is Nothing Then
            mobjCurrEnc.PrintEncounter True
        End If
    Case cwiABOUT
        '---inform CW that I am showing a modal form---
        mobjComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION
        Dim objAbout As New CWAbout
        Call objAbout.Display(App, frmEncParent)
        Set objAbout = Nothing
        '---inform CW that I am done showing a modal form---
        mobjComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION

'<< SCR 0768
    Case cwiOK_TO_CHANGE_PATIENT
            '-- check to see if it's ok with the active kid
        If mobjCurrEnc Is Nothing Then Exit Sub
        If mobjCurrEnc.CurrSection Is Nothing Then
            'Display message is no section is open
            Exit Sub 'Exit if currsection is nothing SCR-32919
        End If
        
        On Error Resume Next
        Set oSec = mobjCurrEnc.CurrSection
        
        If vsSourceServerNCID = NCID_CW_APPLICATION Then
        
            oSec.msOKToChangePatient = ""
            'First check if section is active
            If mobjComm.Message(cwiOLE_SERVER_ACTIVE, oSec.msAppNCID, NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION) Then
                'SCR 21703 2/25/02  Phil Crowder  Set StartOLEServer flag to false
                mobjComm.Message cwiOK_TO_CHANGE_PATIENT, "", oSec.msAppNCID, NCID_ENCOUNTER_APPLICATION, False
            Else
                oSec.msOKToChangePatient = "Y"
            End If
            
            'Only display message if section returns "Y" SCR-32854, '<SCR SCR 36493
            
            'SCR 21703 2/25/02  Phil Crowder  Set StartOLEServer flag to false
            mobjComm.Message cwiOK_TO_CHANGE_PATIENT, oSec.msOKToChangePatient, NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION, False
        Else '- it's from a child
            oSec.msOKToChangePatient = UCase$(vsMessageText)
        End If
        Err.Clear
        
    Case cwiOK_TO_CLOSE_VIEW

        '-- check to see if it's ok
        'Testing out of body.
        If mobjCurrEnc Is Nothing Then Exit Sub
        'If mobjCurrEnc Is Nothing Then Exit Sub
        
        On Error Resume Next
        Set oSec = mobjCurrEnc.CurrSection
        
        If vsSourceServerNCID = NCID_CW_APPLICATION Then
                    
            oSec.msOKToCloseView = ""
            'Check if section is active and ask if OK_TO_CLOSE_VIEW SCR-26703
            If mobjComm.Message(cwiOLE_SERVER_ACTIVE, oSec.msAppNCID, NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION) Then
                 mobjComm.Message cwiOK_TO_CLOSE_VIEW, "", oSec.msAppNCID, NCID_ENCOUNTER_APPLICATION, False
            Else
                oSec.msOKToCloseView = "Y"
            End If
            mobjComm.Message cwiOK_TO_CLOSE_VIEW, oSec.msOKToCloseView, NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION, False
        Else '- it's from a child
            oSec.msOKToCloseView = UCase$(vsMessageText)
        End If
        Err.Clear
    Case cwiPASTE
        If mobjCurrEnc Is Nothing Then Exit Sub
        If mobjCurrEnc.CurrSection Is Nothing Then Exit Sub
        
        Set oSec = mobjCurrEnc.CurrSection
        'Send paste to currently open section where paste will take place
        mobjComm.Message cwiPASTE, "", oSec.msAppNCID, NCID_ENCOUNTER_APPLICATION, False
        
    Case cwiCUT
        If mobjCurrEnc Is Nothing Then Exit Sub
        If mobjCurrEnc.CurrSection Is Nothing Then Exit Sub
        
        Set oSec = mobjCurrEnc.CurrSection
        'Send paste to currently open section where paste will take place
        mobjComm.Message cwiCUT, "", oSec.msAppNCID, NCID_ENCOUNTER_APPLICATION, False
    
    Case cwiCOPY
        If mobjCurrEnc Is Nothing Then Exit Sub
        If mobjCurrEnc.CurrSection Is Nothing Then Exit Sub
        
        Set oSec = mobjCurrEnc.CurrSection
        'Send paste to currently open section where paste will take place
        mobjComm.Message cwiCOPY, "", oSec.msAppNCID, NCID_ENCOUNTER_APPLICATION, False

    Case cwiSPELL_CHECK
        If mobjCurrEnc Is Nothing Then Exit Sub
        If mobjCurrEnc.CurrSection Is Nothing Then Exit Sub
        
        Set oSec = mobjCurrEnc.CurrSection
        'Send paste to currently open section where paste will take place
        mobjComm.Message cwiSPELL_CHECK, "", oSec.msAppNCID, NCID_ENCOUNTER_APPLICATION

    Case cwiFOREGROUND_APPLICATION
        'Pass message to Dental SCR-29438
        If Not mobjOpenChildModule Is Nothing Then
            If TypeName(mobjOpenChildModule) = "Dental" Then
               mobjOpenChildModule.Message cwiFOREGROUND_APPLICATION, "", NCID_ENCOUNTER_APPLICATION
               Exit Sub
            End If
        End If
    
        
    End Select
    Err.Clear
    Exit Sub
ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description & " - " & sMsg, "EncounterParent.Message", "EncounterParent", vbExclamation)
End Sub

Private Sub MessageGeneric(vsMessageText As String)
    Dim buff() As String
    Dim lSection As Long
    
    On Error Resume Next
    buff = Split(vsMessageText, "|")
    Select Case UCase$(buff(0))
    Case "AMENDENC"
        Call AmendEnc
    Case "OPENCHILD"
        '  part 2 should be the ncid of the child to open
        If mbMainViewRunning _
        And buff(1) <> "" _
        And Not mobjCurrEnc Is Nothing Then
            'Call mobjCurrEnc.OpenSection(SectionNCIDToEnum(buff(1)))
            mobjCurrEnc.OpenSection (SectionNCIDToEnum(buff(1)))
        End If
    Case "LAUNCH"
        If mbMainViewRunning _
        And buff(1) <> "" _
        And Not mobjCurrEnc Is Nothing Then
            If buff(1) = NCID_BROWSER_APPLICATION Then
                'Call mobjCurrEnc.EditNewFormToolNote
                'chart notes is no longer supported
            Else
                '--- SCR 81705   Sherry Wang   11/9/2005
                mbLaunchSucceeded = mobjCurrEnc.OpenSection(SectionNCIDToEnum(buff(1)))
                If mbLaunchSucceeded = True Then
                    If UBound(buff) > 1 Then
                        mobjComm.Message cwiGENERIC, "LAUNCH" & "|" & buff(2), buff(1), NCID_ENCOUNTER_APPLICATION
                    End If
                Else
                    '--- SCR 78012    Sherry Wang  1/10/2006
                    If mobjCurrEnc.EncounterType <> InpatientNote And SectionNCIDToEnum(buff(1)) <> eRFV Then
                        CloseView cwiREGULAR_VIEW
                    End If
                End If
            End If
        End If
    Case "CLOSEMODULE"   '- the little x in the corner of the core frame was hit; close child
            
            '''''''''''''''''''''''''''''''
            ' SCR 11660 object reference  '
            ' error occurs on close.      '
            '''''''''''''''''''''''''''''''
            m_bClosing = True            '
            '''''''''''''''''''''''''''''''
            If mobjCurrEnc.CurrSection Is Nothing Then  'frmEncParent.VP1.Visible And mobjCurrEnc.CurrSection Is Nothing Then
                Call CloseCurrentEncounter
            Else
                Call mobjCurrEnc.CloseSection
            End If
            
            m_bClosing = False
            
    Case "EXIT_APP"  'SCR-25199
    
        m_bClosing = True
        
        Call CloseCurrentEncounter
            
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
'    robjLogonEnvironment - Contains the logon environment object.  Contains information such
'                           as who logged on, where they logged on, etc.
'    robjConfig - Contains the configuration object that is used to get and set configuration
'                 information for a specific application.
'    robjPatient - Contains the patient object containing the currently selected patient.
'****************************************************************************************************

Public Sub Initialize(ByVal vlMyProgId As Long, _
                            robjComm As Object, _
                            robjLogonEnvironment As Object, _
                            robjConfig As Object, _
                            robjPatient As Object)
    #If DEBUGON Then
        Dim DR As DebugRec
        If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Initialize", "Begin", DR, "Initialize", robjPatient.Name
    #End If
    If Not DebugMode Then On Error GoTo ErrHandler
    App.HelpFile = mobjShared.CWHelpFile
    mlProgID = vlMyProgId
    If Not mobjLogonEnvironment Is Nothing Then
        MsgBxARMd "Internal Error: EncounterParent module Initialize() routine was called a second time.", vbExclamation, "Encounter"
        Exit Sub
    End If
    Set mobjComm = robjComm
    Set mobjLogonEnvironment = robjLogonEnvironment
    Set mobjConfig = robjConfig
    Set mobjPatient = robjPatient
    Set mobjShared = New CWShared
    mbMainViewRunning = False
    'NeedsRefresh = True
    bPatientSummaryViewRunning = False
    #If DEBUGON Then
        If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Initialize", "End", DR, "vlMyProgId", vlMyProgId
    #End If
    Call mobjComm.SetStatusText("")
    Exit Sub
ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description, "EncounterParent.Initialize", "EncounterParent", vbExclamation)
    Call mobjComm.SetStatusText("")
End Sub

'**************************************************************************************************************
'  Description: This method is called by a client when they want to modify the "properties", or the
'               configuration of this OLE Server.
'
'               NOTE: This is an OPTIONAL method.
'
'  Parameters:
'    robjLogonEnvironment - Contains the logon environment object.
'    robjConfig - Contains the configuration object that is used to get and set configuration
'                 information for a specific application.
'    vsOwnerNCID - Contains the NCID of the "owner" to use in storing modifications to the configuration.
'    vbCalledByDesigner - TRUE if this method is being called by the CW Designer, FALSE otherwise.
'**************************************************************************************************************
Public Sub Properties(robjLogonEnvironment As Object, _
                      robjConfig As Object, _
                ByVal vsOwnerNCID As String, _
                ByVal vbCalledByDesigner As Boolean)
End Sub

Private Function I_RSTransfer_Transfer(oDMBE As IDMBE.I_DMBE) As Variant
Dim objXfer As EncTransfer
Dim lngRet As Long
Dim cCol As Collection
Dim bEICTransfer As Boolean

On Error GoTo ErrHandler


    If oDMBE.Source = eDM_CDR And oDMBE.Destination = eDM_LDDB Then
        'Don't Do anything
        lngRet = 1 'set non-zero value
    ElseIf oDMBE.Source = eDM_LDDB And oDMBE.Destination = eDM_CDR Then
        lngRet = 1 'transfer failed, set non-zero value
        
    ElseIf (oDMBE.Source = eDM_PIC Or oDMBE.Source = eDM_AGGPIC Or oDMBE.Source = eDM_EIC) And oDMBE.Destination = eDM_LDDB Then
        lngRet = mobjEncounterOps.TransferPIC2LDDB(oDMBE)
    Else
        Set objXfer = New EncTransfer
        Set objXfer.Comm = mobjComm
        objXfer.Init
        lngRet = objXfer.Transfer(oDMBE)
    End If


I_RSTransfer_Transfer = lngRet

Exit Function

ErrHandler:
 lngRet = Err.Number
 
 MsgBxARMd "Error Number:  " & CStr(Err.Number) & vbCrLf & _
        "Description:   " & Err.Description, vbExclamation, "EncounterParent"
        
 Debug.Print Err.Number, Err.Description
 I_RSTransfer_Transfer = lngRet

End Function

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
'    robjLogonEnvironment - Contains the logon environment object.  Contains information such
'                           as who logged on, where they logged on, etc.
'    robjConfig - Contains the configuration object that is used to get and set configuration
'                 information for a specific application.
'    robjPatient - Contains the patient object containing the currently selected patient.
'    robjUser - Security User information object
'****************************************************************************************************
Public Sub ICHCSII_CoreAPI2_Initialize2(ByVal vlMyProgId As Long, _
                                              robjComm As Object, _
                                              robjLogonEnvironment As Object, _
                                              robjConfig As Object, _
                                              robjPatient As Object, _
                                              robjUser As Object)
  
    #If DEBUGON Then
        Dim DR As DebugRec
        If Not DL Is Nothing Then DL.DebugLog Exe, Module, "CoreAPI2_Initialize2", "Begin", DR, "vlMyProgId", vlMyProgId
    #End If
    If Not DebugMode Then On Error GoTo ErrHandler
    Set mobjUser = robjUser
    Set mobjEncounterOps = New CHCSIIEncounterOps.EncounterOps
    Call mobjEncounterOps.Initialize(robjComm, robjLogonEnvironment, robjConfig, robjPatient, robjUser)
    Call Initialize(vlMyProgId, robjComm, robjLogonEnvironment, robjConfig, robjPatient)
    #If DEBUGON Then
        If Not DL Is Nothing Then DL.DebugLog Exe, Module, "CoreAPI2_Initialize2", "End", DR, "vlMyProgId", vlMyProgId
    #End If
    Exit Sub
ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description, "EncounterParent.Initialize2", "EncounterParent", vbExclamation)
    Exit Sub
End Sub

Public Sub ChildClosed()
    Call HideChild
End Sub

Friend Sub AddMenus()
    Dim sSignStatus As String
    Dim bAddMenu As Boolean
    On Error Resume Next
    If mlMenuRefresh = 0 Then
        mlMenuRefresh = mobjComm.AddActionMenuItem("&Refresh", "refresh.bmp", mlProgID, "Refresh Screen")
    End If
    
    If Not mobjShared.IsAppMode(modeCHCSI_GUI) Then  '<SCR SCR 36493
        If mlMenuNote = 0 Then
            mlMenuNote = mobjComm.AddActionMenuItem("Add &Note", "new.bmp", mlProgID, "Add a Note to the Encounter")
        End If
        
        If mlMenuProviders = 0 Then
            mlMenuProviders = mobjComm.AddActionMenuItem("Add. &Providers", "link.bmp", mlProgID, "Add Additional Providers")
        End If
        If mlMenuTemplates = 0 Then
    
            mlMenuTemplates = mobjComm.AddActionMenuItem("&Templates", "select_encounter.bmp", mlProgID, "Select an Encounter Template")
    
           
        End If
        
        If Not mobjComm.CmdLineSwitch("CPG") = "" Then
            If mobjUser.HasPrivilegeEx(Priv_Patient_Plan_Of_Care, Priv_Read) Then 'SCR-32190
                If mlMenuGoalSummary = 0 Then
                   mlMenuGoalSummary = mobjComm.AddActionMenuItem("&Goal Summary", "Health_History.bmp", mlProgID, "Review patient goal summary")
                End If
            End If
        End If
        
        If mlMenuSign = 0 Then
            mlMenuSign = mobjComm.AddActionMenuItem(gsARMTx_Is_Encounter_Caption_Sign, "dispsign.bmp", mlProgID, gsARMTx_Is_Encounter_Sign_Prompt)
            '--- SCR-ARM-25468/27364/32970 RWalker 5/20/03
            'Unit testing in parallel with existing gsARM_Tx_SignEncounter
            goARM.InitCmplxTxCls gsARM_Tx_CorrUT_SignFamily, _
                                 NCID_ENCOUNTER_APPLICATION & gsARMTx_Is_Encounter_Caption_Sign & CStr(mlMenuSign), _
                                 gsARM_ACT_FROM
        End If
        
        If mlMenuCreateTemplate = 0 Then
            mlMenuCreateTemplate = mobjComm.AddActionMenuItem("Save &As Template", "saveas.bmp", mlProgID, "Create new Encounter Template from this Encounter")
            
        End If
        
        If mlMenuPrintSF600 = 0 Then
            'SCR 8902 KDunne 2/26/2002
            'For Dental Encounters only, encounter form is SF603A. For all others
            'encounter form is SF600.
            If mobjCurrEnc.EncounterType = Dental Then
                mlMenuPrintSF600 = mobjComm.AddPrintMenuItem("&SF603A", mlProgID, "Print SF603A")
            Else
                mlMenuPrintSF600 = mobjComm.AddPrintMenuItem("&SF600", mlProgID, "Print SF600")
            End If
        End If
        If mlMenuPrintPreviewSF600 = 0 Then
            'SCR 8902 KDunne 2/26/2002
            If mobjCurrEnc.EncounterType = Dental Then
                mlMenuPrintPreviewSF600 = mobjComm.AddPrintPreviewMenuItem("&SF603A", mlProgID, "Preview SF603A")
            Else
                mlMenuPrintPreviewSF600 = mobjComm.AddPrintPreviewMenuItem("&SF600", mlProgID, "Preview SF600")
            End If
        End If
                
        If Len(mobjComm.CmdLineSwitch("PKC")) Then
            If mlMenuRunCoupler = 0 Then
                mlMenuRunCoupler = mobjComm.AddActionMenuItem("Run PKC Coupler", "Couplers.ico", mlProgID, "Run Coupler")
            End If
        End If
        

        
    Else
        If mlMenuPrintSF600 = 0 Then
            'SCR 8902 KDunne 2/26/2002
            'For Dental Encounters only, encounter form is SF603A. For all others
            'encounter form is SF600.
            If mobjCurrEnc.EncounterType = Dental Then
                mlMenuPrintSF600 = mobjComm.AddPrintMenuItem("&SF603A", mlProgID, "Print SF603A")
            Else
                mlMenuPrintSF600 = mobjComm.AddPrintMenuItem("&SF600", mlProgID, "Print SF600")
            End If
        End If
        If mlMenuPrintPreviewSF600 = 0 Then
            'SCR 8902 KDunne 2/26/2002
            If mobjCurrEnc.EncounterType = Dental Then
                mlMenuPrintPreviewSF600 = mobjComm.AddPrintPreviewMenuItem("&SF603A", mlProgID, "Preview SF603A")
            Else
                mlMenuPrintPreviewSF600 = mobjComm.AddPrintPreviewMenuItem("&SF600", mlProgID, "Preview SF600")
            End If
        End If
        
        If mlMenuSign = 0 Then
            mlMenuSign = mobjComm.AddActionMenuItem(gsARMTx_Is_Encounter_Caption_Sign, "dispsign.bmp", mlProgID, gsARMTx_Is_Encounter_Sign_Prompt)
            '--- SCR-ARM-25468/27364/32970 RWalker 5/20/03
            'Supports new correlated transaction.  The PGUI transaction will share common
            'modal dialog transactions with the CDR mode 'Sign Encounter' transaction.
            'The CDR and PGUI versions of Sign Encounter now share the same constant literal.
            goARM.InitCmplxTxCls gsARM_Tx_CorrUT_SignFamily, _
                                 NCID_ENCOUNTER_APPLICATION & gsARMTx_Is_Encounter_Caption_Sign & CStr(mlMenuSign), _
                                 gsARM_ACT_FROM
        End If
        
'        If mlTestApp = 0 Then
'            mlTestApp = mobjComm.AddActionMenuItem("Test App", "Couplers.ico", mlProgID, "Run Coupler")
'        End If
        
    End If
End Sub

Friend Sub RemoveMenus()
    On Error Resume Next
    If mlMenuRefresh <> 0 Then
        mobjComm.DeleteMenuItem mlMenuRefresh
        mlMenuRefresh = 0
    End If
    If mlMenuNote <> 0 Then
        mobjComm.DeleteMenuItem mlMenuNote
        mlMenuNote = 0
    End If
    If mlMenuProviders <> 0 Then
        mobjComm.DeleteMenuItem mlMenuProviders
        mlMenuProviders = 0
    End If
    If mlMenuTemplates <> 0 Then
        mobjComm.DeleteMenuItem mlMenuTemplates
        mlMenuTemplates = 0
    End If
    If mlMenuGoalSummary <> 0 Then
        mobjComm.DeleteMenuItem mlMenuGoalSummary
        mlMenuGoalSummary = 0
    End If
    If mlMenuSign <> 0 Then
        mobjComm.DeleteMenuItem mlMenuSign
        mlMenuSign = 0
    End If
    If mlMenuCreateTemplate <> 0 Then
        mobjComm.DeleteMenuItem mlMenuCreateTemplate
        mlMenuCreateTemplate = 0
    End If
    If mlMenuPrintPreviewSF600 <> 0 Then
        mobjComm.DeleteMenuItem mlMenuPrintPreviewSF600
        mlMenuPrintPreviewSF600 = 0
    End If
    If mlMenuPrintSF600 <> 0 Then
        mobjComm.DeleteMenuItem mlMenuPrintSF600
        mlMenuPrintSF600 = 0
    End If
    If mlMenuRunCoupler <> 0 Then
        mobjComm.DeleteMenuItem mlMenuRunCoupler
        mlMenuRunCoupler = 0
    End If
    
    
End Sub

Public Function AddToEncounterData(ByRef AppNCID As String, ByRef RTF As String) As Boolean
    If mobjCurrEnc Is Nothing Then
        MsgBxARMd "There is no encounter open at this time."
    Else
        If mobjCurrEnc.AddToSection(AppNCID, RTF) Then
            AddToEncounterData = True
            If mbMainViewRunning Then
                Call frmEncParent.RefreshDisplay
            End If
        End If
    End If
End Function

Friend Function CloseEncounterPrivate() As Boolean
    #If DEBUGON Then
        Dim DR As DebugRec
        If Not DL Is Nothing Then DL.DebugLog Exe, Module, "CloseEncounterPrivate", "Begin", DR
    #End If
    On Error Resume Next
    If Not mobjCurrEnc Is Nothing Then
        Set mobjOpenChildModule = Nothing
        If mbQETelconRunning Then
            mbQETelconRunning = False
            Unload frmQETelecon
        End If
        
        If mobjCurrEnc.EncounterType = TeleConsult Then
            'When opening TelCons, Encounter is locked SCR-29725
            Call mobjCurrEnc.LockingObject.UnlockSection(eEntireEncounter, 0)
        End If
'<< SCR  10768  save first then close the encounter?  if children are open, close them first
              
        If Not mobjCurrEnc.CloseEncounter Then
            CloseEncounterPrivate = False
            Exit Function
        End If
        Call mobjCurrEnc.Save
'>>
        RaiseEvent CloseEncounter(mobjCurrEnc)
        Set frmEncParent.mobjCurrEnc = Nothing
        Call mobjEncounterOps.ResetCurrLoadedEncounter 'SCR-31509
        Set mobjCurrEnc = Nothing
        CloseEncounterPrivate = True
    End If
    #If DEBUGON Then
        If Not DL Is Nothing Then DL.DebugLog Exe, Module, "CloseEncounterPrivate", "End", DR
    #End If
    
End Function

Public Function CloseCurrentEncounter() As Boolean
    #If DEBUGON Then
        Dim DR As DebugRec
        If Not DL Is Nothing Then DL.DebugLog Exe, Module, "CloseCurrentEncounter", "Begin", DR
    #End If
    On Error Resume Next
    
    If CloseEncounterPrivate Then
        CloseCurrentEncounter = True
        frmEncParent.mbCloseMyself = True
        Unload frmEncParent
    End If
    
    #If DEBUGON Then
        If Not DL Is Nothing Then DL.DebugLog Exe, Module, "CloseCurrentEncounter", "End", DR
    #End If

End Function

Public Function SetCurrentEncounter(ByVal PatientUnitNumber As String, _
                                    ByVal FacilityNCID As String, _
                                    ByVal EncounterNumber As String, _
                                    Optional p_bOLEServerCalling As Boolean, _
                                    Optional CallBackPhone As String = "", _
                                    Optional bAmendEncounter As Boolean = False, _
                                    Optional sApptIEN As String = "", _
                                    Optional enumApptEnc_Status As CHCSII_AppointmentStatus.EncounterStatusEnum = 0) As Boolean
                                    

'Public Enum EncounterStatusEnum
'    EncCheckedIn = 1
'    EncWaiting = 2
'    EncInProgress = 3
'    EncCompleted = 4
'    EncNeedsCosignature = 5
'    EncUpdated = 6
'    EncUpdating = 7
'    EncCancel = 8
'End Enum

    Dim objEncounterPrev As Object
    Dim sReturnMsg As String
    
    On Error GoTo ErrHandler
    
    Screen.MousePointer = vbHourglass
    
    If Not mobjCurrEnc Is Nothing Then
      'Phil Crowder 8/20/03  Performance Improvements
      'Should check to make sure checking encounter number is also valid for other modes
      If (Not mobjShared.IsAppMode(modeCHCSI_GUI)) Or (mobjShared.IsAppMode(modeCHCSI_GUI) And EncounterNumber <> mobjCurrEnc.EncounterID) Then
          If CloseCurrentEncounter Then
            Set mobjCurrEnc = Nothing
          Else
            'If unable to close current encounter then exti function SCR-38897
            'Do not display msg scr#44074
            'MsgBxARMd "Unable to close current encounter.", "Encounters"
            Exit Function
          End If
      End If
    End If
    
    If mobjCurrEnc Is Nothing Then
        Set mobjCurrEnc = mobjEncounterOps.GetEncounter(PatientUnitNumber, FacilityNCID, EncounterNumber)
    End If
    
    If mobjCurrEnc Is Nothing Then
        If mbMainViewRunning Then
            frmEncParent.mbCloseMyself = True
            Unload frmEncParent
        End If
        MsgBxARMd "Unable to retrieve encounter. (" & PatientUnitNumber & " - " & FacilityNCID & " - " & EncounterNumber & ") in EncouterParent.SetCurrentEncounter", vbCritical + vbOKOnly, "EncounterParent"
        Exit Function
    End If

    If enumApptEnc_Status <> 0 Then
        'Check if Appt and Enc Status are in sync SCR-42360
        If Not Validate_Appt_Enc_Status(mobjCurrEnc.AppointmentID, mobjCurrEnc.EncounterID, mobjCurrEnc.Status, enumApptEnc_Status, sReturnMsg) Then
            MsgBxARMd sReturnMsg
        Else
            'If returns true with msg then update appt encounter status SCR#51097, 52460
            If Len(sReturnMsg) Then
                MsgBxARMd sReturnMsg
                mobjCurrEnc.UpdateApptEncounterStatus
            End If
        End If
    End If
    
    If Not mobjEncounterOps.ViewSensitiveEncounter(mobjCurrEnc) Then
        Set mobjCurrEnc = Nothing
        Exit Function
    End If
    
    If mobjShared.IsAppMode(modeCHCSI_GUI) Then  '<SCR SCR 36493
        mobjCurrEnc.ApptIEN = sApptIEN
        mobjCurrEnc.Patient_IEN = Patient.IEN_number
    End If
    
    'Is user is amending encounter then do not bring to Prev Enc SCR-833 ITT
    If Not bAmendEncounter Then
        Select Case mobjCurrEnc.Status
            Case complete, Updated, PatientLeftWithoutBeingSeen, _
                CancelByPatient, CancelByFacility, NoShow, NeedsCoSignature 'SCR-27718
                    '-- since I don't show these in the 'enc select' window, I must be being asked programmatically to display them.
                    '   Send the user to the prev enc module.
                    OpenPreviousEncounters
                    SetCurrentEncounter = True
                    Exit Function
            
        End Select
    End If
    
    Select Case True
    
    Case Is = p_bOLEServerCalling
        If mbMainViewRunning Then
            Set frmEncParent.mobjCurrEnc = mobjCurrEnc
            Call frmEncParent.RefreshDisplay
            ' unwanted recursion.
            ' this is associated with the modal select
            ' dialog causing recursion.
            ' tell core to put us on top, if we are not already there
            mobjComm.Message cwiSTART_OLE_SERVER, NCID_ENCOUNTER_APPLICATION, NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION
            AddMenus
            SetCurrentEncounter = True
        Else
            ' unwanted recursion.
            ' this is associated with the modal select
            ' dialog causing recursion.
            If mobjComm.Message(cwiSTART_OLE_SERVER, NCID_ENCOUNTER_APPLICATION, NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION) Then
                If Not frmEncParent.mobjCurrEnc Is Nothing Then 'The user doesn't have permission to open the current encounter.
                    If frmEncParent.mobjCurrEnc.EncounterID = mobjCurrEnc.EncounterID Then
                        SetCurrentEncounter = True
                    End If
                End If
            End If
        End If
        
    Case Else
    
         AddMenus
         Set frmEncParent.mobjCurrEnc = mobjCurrEnc
         SetCurrentEncounter = True
         
    End Select
     
    '''''''''''''''''''''''''''''''
    ' SCR 11660 object reference  '
    ' error occurs on close.      '
    '''''''''''''''''''''''''''''''
    If SetCurrentEncounter And Not m_bClosing Then
        If ShouldUseQETelcon() Then
            OpenQETelcon CallBackPhone
        End If
        RaiseEvent OpenEncounter(mobjCurrEnc)
    End If
    'Reset status text when a new encounter is opened SCR-21218
    mobjComm.SetStatusText ""
    Screen.MousePointer = vbDefault
    ''''''''''''''''''''''''''
     m_bClosing = False      '
    ''''''''''''''''''''''''''
    
    Exit Function
    
ErrHandler:

    mobjShared.ShowVBError Err.Number, Err.Description, "EncounterParent.SetCurrentEncounter", "EncounterParent", vbExclamation

Exit Function
Resume
End Function

Private Sub OpenPreviousEncounters()
    
    Dim objEncounterPrev As Object
    
    On Error GoTo ErrHandler
    
    If Len(mobjComm.CmdLineSwitch("INPATIENT2")) > 0 Then
        'Open up the documents module.
        Call mobjComm.Message(cwiSTART_OLE_SERVER, NCID_INPATIENT_DOCUMENTS_APPLICATION, NCID_CW_APPLICATION, NCID_PROBLEMS_APPLICATION) 'SCR-48615 RBELL 1/23/04 Added
        Set objEncounterPrev = mobjComm.InitializeOLEServer(NCID_INPATIENT_DOCUMENTS_APPLICATION) 'SCR-48615 RBELL 1/23/04 Added
        Call objEncounterPrev.SetEncounterKey(mobjCurrEnc.StorageKey)
        Set mobjCurrEnc = Nothing
    Else
        Set objEncounterPrev = mobjComm.InitializeOLEServer(NCID_PREVIOUS_ENCOUNTERS_APPLICATION)
        If Not objEncounterPrev Is Nothing Then
            mobjComm.Message cwiSTART_OLE_SERVER, NCID_PREVIOUS_ENCOUNTERS_APPLICATION, NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION
            Call objEncounterPrev.SetEncounterKey(mobjCurrEnc.StorageKey)
            Set mobjCurrEnc = Nothing
        End If
    End If
    
    Set objEncounterPrev = Nothing
            
    Exit Sub
    
ErrHandler:

    mobjShared.ShowVBError Err.Number, Err.Description, "EncounterParent.OpenPreviousEncounters", "EncounterParent", vbExclamation

End Sub

Private Sub OpenQETelcon(ByVal CallBackPhone As String)

    On Error GoTo ErrHandler
    
    Screen.MousePointer = vbHourglass
    If mobjCurrEnc.LockingObject.LockSection(eEntireEncounter, 0, SectionLock, "Telcon Quick Entry") Then
        If ShouldUseQETelcon() Then
            mbQETelconRunning = True
            Call RemoveMenus
            frmQETelecon.CallBackPhone = CallBackPhone
            Call frmQETelecon.Display(Me)
            frmEncParent.picSummary.Enabled = False
        Else
            MsgBxARMd "Another user has changed this encounter so that it is no longer available for Telcon Quick Entry."
            Call mobjCurrEnc.LockingObject.UnlockSection(eEntireEncounter, 0)
        End If
    End If
            
    Exit Sub
    
ErrHandler:

    mobjShared.ShowVBError Err.Number, Err.Description, "EncounterParent.OpenQETelecon", "EncounterParent", vbExclamation

End Sub

Public Function CreateEncounter(ApptID As String, _
                                 ApptType As String, _
                                 Enc_Type As EncType, _
                                 FacilityNCID As String, _
                                 ClinicNCID As String, _
                                 PrimaryProviderNCID As String, _
                        Optional ClerkNote As String = "", _
                        Optional bArrivedByAmbulance As Boolean = False, _
                        Optional dApptDate As Date, _
                        Optional bCheckPriv As Boolean = True, _
                        Optional bResultedConsultOrder As Boolean = False, _
                        Optional sApptReason As String = "", _
                        Optional sApptComment As String = "", _
                        Optional lUnitNumber As Long = 0, _
                        Optional sWorkLoadOptions As String = "", _
                        Optional sWorkLoad As String = "", _
                        Optional bObservation As Boolean = False, _
                        Optional sAMEPRSCode As String = "", _
                        Optional sBMEPRSCode As String = "", _
                        Optional sApptIEN As String = "", _
                        Optional sInpatientService As String = "", _
                        Optional oEncAccidentData As EncAccidentData = Nothing, _
                        Optional bTransferToLDDB As Boolean = False, _
                        Optional InptNoteTypeNCID As String = "", _
                        Optional InptHospDay As Long = 0) As Encounter
                        
    Dim oThisEncounter As Encounter
    
    On Error GoTo ErrHnd
    
    Set oThisEncounter = mobjEncounterOps.CreateEncounter(ApptID, _
                                        ApptType, _
                                        Enc_Type, _
                                        FacilityNCID, _
                                        ClinicNCID, _
                                        PrimaryProviderNCID, _
                                        ClerkNote, _
                                        bArrivedByAmbulance, _
                                        dApptDate, _
                                        bCheckPriv, _
                                        bResultedConsultOrder, _
                                        sApptReason, _
                                        sApptComment, _
                                        lUnitNumber, _
                                        sWorkLoadOptions, _
                                        sWorkLoad, _
                                        bObservation, _
                                        sAMEPRSCode, _
                                        sBMEPRSCode, _
                                        sApptIEN, _
                                        sInpatientService, _
                                        oEncAccidentData, _
                                        bTransferToLDDB, _
                                        InptNoteTypeNCID, _
                                        InptHospDay)
    
    If Not oThisEncounter Is Nothing Then
        RaiseEvent CreatedEncounter(oThisEncounter)
   End If
    
    'Phil Crowder 8/20/03  Performance improvement
    'Not sure this shouldn't be done on all modes
    If mobjShared.IsAppMode(modeCHCSI_GUI) Then Set mobjCurrEnc = oThisEncounter
    
    Set CreateEncounter = oThisEncounter
    Set oThisEncounter = Nothing

Exit Function

ErrHnd:
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

'-----------------------------------------------------
'Purpose:  Prev Enc comes here to let user amend an enc
'
'Parameters: none
'
'Comments:
'----------------------------------------------------
Private Sub AmendEnc()
    MsgBxARMd "jf todo: Implement AmendEnc() in EncounterParent."
End Sub

Public Function LeftWithoutBeingSeen(ByVal sEncounterID As String, ByVal sEncounterFacility As String) As Boolean
    Dim objEncounter As Encounter
    On Error Resume Next
    If Not mobjCurrEnc Is Nothing Then
        If mobjCurrEnc.EncounterID = sEncounterID And mobjCurrEnc.FacilityNCID = sEncounterFacility Then
            Set objEncounter = mobjCurrEnc
        End If
    End If
    If objEncounter Is Nothing Then
        Set objEncounter = mobjEncounterOps.GetEncounter(mobjPatient.UnitNumber, sEncounterFacility, sEncounterID)
        If objEncounter Is Nothing Then Exit Function
    End If
    If objEncounter.LeftWithoutBeingSeen() Then
        If objEncounter = mobjCurrEnc And mbMainViewRunning Then
            Call CloseCurrentEncounter
        End If
        LeftWithoutBeingSeen = True
    End If
    Set objEncounter = Nothing
End Function

Public Function OpenChildView(ByRef Module As Object) As Boolean
    On Error Resume Next
    If Module.OpenView(cwiREGULAR_VIEW, frmEncParent.vseChildClient.hwnd, hwnd) Then
        Call ShowChild
        Call Module.ResizeView(cwiREGULAR_VIEW, 0, 0, frmEncParent.vseChildClient.Width, frmEncParent.vseChildClient.Height)
        Set mobjOpenChildModule = Module
        OpenChildView = True
        frmEncParent.VP1.Visible = False
    End If
End Function


Public Sub AllowUI(ByVal bAllowed As Boolean)
    If bAllowed Then
        frmEncParent.Enabled = True
        mobjComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION
    Else
        frmEncParent.Enabled = False
        mobjComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION
    End If
End Sub

Public Sub ShowChild()
    Call RemoveMenus
    frmEncParent.vseChildClient.Visible = True
    frmEncParent.vseChildClient.ZOrder 0
    frmEncParent.VP1.Visible = False
End Sub

Public Sub HideChild()
    Call AddMenus
    frmEncParent.VP1.Visible = True
    frmEncParent.VP1.ZOrder 0
    frmEncParent.vseChildClient.Visible = False
    'Don't need to reload data just refresh Encounter document RTF
    Call RefreshDisplay(False) 'Set to False for performance enhancements SCR-16062,18972
    Set mobjOpenChildModule = Nothing
End Sub

Public Sub RefreshAutoCites(ByVal AppNCID As Long)
    Dim objAutoCite As AutoCiteProp
    
    If mobjCurrEnc Is Nothing Then
        Exit Sub
    End If
    If Not mbMainViewRunning Then
        Exit Sub
    End If
    If mobjCurrEnc.SectionIsEditable(eAutoCite) Then
                        
        For Each objAutoCite In mobjCurrEnc.EncProperties.AutoCites

            If objAutoCite.AppNCID = AppNCID Then
                'Refresh autocites w/o closing any modules SCR-37596
                mobjCurrEnc.RefreshAutoCites
                'Call frmEncParent.cmds_Click(0)
                Set objAutoCite = Nothing
                Exit Sub
            End If
        Next objAutoCite
    End If
End Sub

Public Sub RefreshDisplay(Optional ByVal ReloadData As Boolean = False)
    Call frmEncParent.RefreshDisplay(ReloadData)
End Sub

Private Sub mobjCurrEnc_EncounterClosed()
    RaiseEvent CloseEncounter(mobjCurrEnc)
End Sub

Private Sub mobjCurrEnc_EncStatusChanged(mvarStatus As EncStatus)
'Private Sub mobjCurrEnc_EncStatusChanged(mvarStatus As ENCOUNTER_INTERFACES.EncStatus)

   RaiseEvent CurrEncStatusChanged(mvarStatus)
End Sub

Public Function GetEandMCoding(mCol As Collection) As Boolean

'Assumed this collection is made up of EncEandMCoding objects
Dim objEncounter As CHCSIIEncounterOps.Encounter
Dim objEncEandMCoding As EncEandMCoding

On Error GoTo ErrHandler

GetEandMCoding = True

For Each objEncEandMCoding In mCol
    With objEncEandMCoding
        'Appointments completed in CHCS will not have encounter numbers 'SCR-24305
        If .CriteriaEncounterNumber > 0 Then
            Set objEncounter = mobjEncounterOps.GetEncounter(.CriteriaPatientUnitNumber, .CriteriaEncounterFacilityNCID, .CriteriaEncounterNumber)
            If Not objEncounter Is Nothing Then
                .CodeReview = objEncounter.EandMCodeReview
                .CurrentRTF = objEncounter.CurrentRTF
                .EandMNCID = objEncounter.EandMNCID
                If objEncounter.EandMReviewed Then
                    .EandMReviewed = "Yes"
                Else
                    .EandMReviewed = "No"
                End If
                .EandMReviewedBy = objEncounter.EandMReviewedBy
                .EandMReviewedDate = objEncounter.EandMReviewDate
            Else
                GetEandMCoding = False
            End If
            Set objEncounter = Nothing
        End If
    End With
    Set objEncEandMCoding = Nothing
Next

Exit Function

ErrHandler:

GetEandMCoding = False
Call mobjShared.ShowVBError(Err.Number, Err.Description, "EncounterParent.GetEandMCoding", "EncounterCurrent", vbCritical)

End Function

Private Function CheckTemplateWritePriv() As Boolean

    CheckTemplateWritePriv = mobjUser.HasPrivilegeEx(Priv_Templates_Clinic_Templates, Priv_Write) Or _
                             mobjUser.HasPrivilegeEx(Priv_Templates_Enterprise_Template_Guidelines, Priv_Write) Or _
                             mobjUser.HasPrivilegeEx(Priv_Templates_MTF_Template_Guidelines, Priv_Write) Or _
                             mobjUser.HasPrivilegeEx(Priv_Templates_Personal_Template, Priv_Write)

End Function
'Called by SetCurrentEncounter
Private Function Validate_Appt_Enc_Status(sApptID As String, sEncNum As String, _
        enumEnc_Status As ENCOUNTER_INTERFACES.EncStatus, enumApptEnc_Status As CHCSII_AppointmentStatus.EncounterStatusEnum, sReturnMsg As String) As Boolean
'SCR-36381, 42360
Dim bOutofSync As Boolean
On Error GoTo ErrHandler
    sReturnMsg = ""
    If sApptID = "" Or sEncNum = "" Or CInt(enumEnc_Status) < 1 Or CInt(enumApptEnc_Status) < 1 Then
        sReturnMsg = "Unable to validate appointment status to encounter status. Invalid arguments"
        Exit Function
    End If
    
    Select Case enumEnc_Status

        Case CancelByFacility, PatientLeftWithoutBeingSeen, CancelByPatient, NoShow
            If enumApptEnc_Status <> EncCancel Then bOutofSync = True

        Case CheckedIn
            If enumApptEnc_Status <> EncCheckedIn Then bOutofSync = True
        
        Case complete
            If enumApptEnc_Status <> EncCompleted Then bOutofSync = True
        
        Case InProgress
            If enumApptEnc_Status <> EncInProgress Then bOutofSync = True
            
        Case NeedsCoSignature
            If enumApptEnc_Status <> EncNeedsCosignature Then bOutofSync = True
        
        Case Updated
            If enumApptEnc_Status <> EncUpdated Then bOutofSync = True
        
        Case Updating
            If enumApptEnc_Status <> EncUpdating Then bOutofSync = True
        
        Case Waiting
            If enumApptEnc_Status <> EncWaiting Then bOutofSync = True
            
        Case NeedsSignature, NoSave, Pending, Incomplete
            'No equivilent appt status
            bOutofSync = True
            
        Case Else
            bOutofSync = True
            sReturnMsg = "Unknown Encounter Status: " + CStr(enumEnc_Status) + ". "
    
    End Select
    If bOutofSync = True Then
        Validate_Appt_Enc_Status = True
        sReturnMsg = sReturnMsg + "The appointment status and the encounter status are out of sync for ApptID/EncID " + sApptID + "/" + sEncNum + "." + _
                    vbCrLf + vbCrLf + "Appointment Status will be changed from  " + TranslateApptStatus(enumApptEnc_Status) + " to " + TranslateEncounterStatus(enumEnc_Status) + "."
                    
                
    Else
        Validate_Appt_Enc_Status = True
    End If
 
 Exit Function
ErrHandler:
 mobjShared.ShowVBError Err.Number, Err.Description, "AppointmentsOps.cls - Validate_Appt_Enc_Status", App.Title, vbCritical

End Function

'*******************************************************************************************************************
Private Function TranslateApptStatus(rnStatus As CHCSII_AppointmentStatus.EncounterStatusEnum) As String
    Select Case rnStatus
    Case EncCheckedIn
        TranslateApptStatus = "CheckedIn"
    Case EncWaiting
        TranslateApptStatus = "Waiting"
    Case EncInProgress
        TranslateApptStatus = "InProgress"
    Case EncUpdated
        TranslateApptStatus = "Updated"
    Case EncNeedsCosignature
        TranslateApptStatus = "NeedsCoSignature"
    Case EncCompleted
        TranslateApptStatus = "Complete"
    Case EncUpdating
        TranslateApptStatus = "Updating"
    Case Else
        TranslateApptStatus = "Unknown" + CStr(rnStatus)
    End Select
End Function

'*******************************************************************************************************************
Private Function TranslateEncounterStatus(enumEnc_Status As ENCOUNTER_INTERFACES.EncStatus) As String


    Select Case enumEnc_Status
        Case CancelByFacility
            TranslateEncounterStatus = "CancelByFacility"
        Case CancelByPatient
            TranslateEncounterStatus = "CancelByPatient"
        Case CheckedIn
            TranslateEncounterStatus = "CheckedIn"
        Case complete
            TranslateEncounterStatus = "complete"
        Case Incomplete
            TranslateEncounterStatus = "Incomplete"
        Case InProgress
            TranslateEncounterStatus = "InProgress"
        Case NeedsCoSignature
            TranslateEncounterStatus = "NeedsCoSignature"
        Case NeedsSignature
            TranslateEncounterStatus = "NeedsSignature"
        Case NoSave
            TranslateEncounterStatus = "NoSave"
        Case NoShow
            TranslateEncounterStatus = "NoShow"
        Case PatientLeftWithoutBeingSeen
            TranslateEncounterStatus = "PatientLeftWithoutBeingSeen"
        Case Pending
            TranslateEncounterStatus = "Pending"
        Case Updated
            TranslateEncounterStatus = "Updated"
        Case Updating
            TranslateEncounterStatus = "Updating"
        Case Waiting
            TranslateEncounterStatus = "Waiting"
        Case Else
            TranslateEncounterStatus = "Unknown: " + CStr(enumEnc_Status)

    End Select

End Function
