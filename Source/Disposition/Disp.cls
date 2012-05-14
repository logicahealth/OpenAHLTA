VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Disp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Disposition"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements ICHCSII_CoreAPI2
Implements ICHCSII_CoreAPI3
Implements IChild

Private Const msCLASS_NAME$ = "Disp"
Private Const msMODULE_NAME$ = msCLASS_NAME$

Private mobjDisposition         As Disposition   '-- The current active disposition

Private mlngProgID              As Long          '-- my prog id from core
'Private mbEntryValidation       As Boolean
'Private mbDentalEntryValidation As Boolean
'Private mbEntryValidated        As Boolean
Private mobjDataRepository      As DataRepository
Private msInpatientID           As String
Private mdAdmissionDate         As Date
Private mdLastTransferDate      As Date
Private mbSuppressEvents        As Boolean

Public Enum enumDispDisplayMode
    Default = 1
    InpatientDischarge = 2
End Enum

'Internal Flags
Private m_lGoToAnp              As Boolean
Private m_lGoToSO               As Boolean
Private m_lGoToEncSign          As Boolean
Private m_bCancel               As Boolean
Private miOverrideNags          As Integer

Public Function DispositionOps() As DataRepository

    If mobjDataRepository Is Nothing Then
        Set mobjDataRepository = New DataRepository
        mobjDataRepository.LoadFromEncounter
    End If
    
    Set DispositionOps = mobjDataRepository
    
End Function
Public Sub CloseView(ByVal pintViewType As Integer)



End Sub 'CloseView

Public Sub Initialize(ByVal plngMyProgID As Long, _
                      ByRef pobjComm As Object, _
                      ByRef pobjLogonEnvironment As Object, _
                      ByRef pobjConfig As Object, _
                      ByRef pobjPatient As Object)
                      
    '====================================================================================
    ' Purpose:  This method is called when the CW Core creates an "instance" of this OLE
    '           Server.  It is used to create and initialize global variables that are
    '           shared across all "views" supported by this OLE Server, i.e. variables
    '           that only need to be initialized ONCE, upon object creation.
    '
    ' WARNING:  This subroutine is not to be confused with Class_Initialize( )
    '           which happens automatically when an instance of this class is created.
    '           This Public subroutine (Initialize) must explicitly be called from
    '           outside this class AFTER Class_Initialize( ) happens.
    '
    ' INPUTS:
    '       plngMyProgId - Contains the "ProgId" that the CW Core has assigned to this OLE
    '                      Server.
    '       pobjComm - Contains a communications object that is used to handle all
    '                  communication between this OLE Server and the CW Core.
    '       pobjLogonEnvironment - Contains the logon environment object.  Contains
    '                              information such as who logged on, where they logged on,
    '                              etc.
    '       pobjConfig - Contains the configuration object that is used to get and set
    '                    configuration information for a specific application.
    '       pobjPatient - Contains the patient object containing the currently selected
    '                     patient.
    '
    ' RETURNS:  None
    '
    ' NOTES:
    '
    'TO DO: Change parameters to "As <explicit object type>"
    '       instead of "As Object" to allow early binding.
    '       WARNING: Must make same change in all other apps & DLLs
    '       that reference this class.
    '----------------------------------------------------------------------------
    '   Modifications:
    '02/28/00 RBELL SCR-3968 Set user ncid for ICHCSII_EncSum2 interface
    '03/23/00 RBELL SCR-4269 Initialzed MedcinSO object
    '07/26/20 H STINSON SCR-6519: Added error trapping , with breakpoint in IDE-mode,
    '         in order to find odd bug.
    '====================================================================================
    
    Dim lngReturn       As Long

    On Error GoTo errXIT
    
    Screen.MousePointer = vbHourglass
    
    '---create references to passed in objects---
    Set cwobjCWComm = pobjComm
    Set cwobjLogonEnviron = pobjLogonEnvironment
    Set cwobjConfig = pobjConfig
    Set cwobjPatient = pobjPatient
    
    'casted vairables
    Set gobjComm = pobjComm
    Set gobjLogonEnviron = pobjLogonEnvironment
    Set gobjConfig = pobjConfig
    Set gobjPatient = pobjPatient
    
    '---Instantiate Help-file object
    App.HelpFile = gobjShared.CWHelpFile 'SCR-HELP
    
    '---save the "ProgId" assigned to me by the CW Core---
    mlngProgID = plngMyProgID
    cwlMyProgId = plngMyProgID
    
    Screen.MousePointer = vbDefault
    
    Exit Sub

errXIT:
    Screen.MousePointer = vbDefault
    
    'CALL THE GENERIC ERROR HANDLER:
    Call gobjShared.ShowVBError(Err.Number, Err.DESCRIPTION, msCLASS_NAME$ + ".Initialize", App.Title, vbExclamation)
    
    Exit Sub

End Sub 'Initialize

Public Sub MenuItemSelected(ByVal MenuId As Long)

    Dim oEncDoc As CHCSIIEncounterOps.Encounter
    
    On Error GoTo errXIT
    
    If mbSuppressEvents Then Exit Sub
    
    mbSuppressEvents = True
    
    m_lGoToAnp = False
    m_lGoToSO = False
    m_lGoToEncSign = False
    m_bCancel = False
    
    '---perform the actions associated with the menu item that was selected---
    Select Case MenuId
        
        Case g_lMenuSave
            miOverrideNags = 0 'reset
            cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_DISPOSITION_APPLICATION
            Call SaveDispositionData              'Save the disposition data to the CDR.
            cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_DISPOSITION_APPLICATION
        
        Case g_lMenuNext, g_lMenuSO, g_lMenuAP
            m_lGoToEncSign = (MenuId = g_lMenuNext)   'Related to SCR-6214?
            m_lGoToSO = (MenuId = g_lMenuSO)          'Related to SCR-6214?
            m_lGoToAnp = (MenuId = g_lMenuAP)         'Related to SCR-6214?
            
            'SCR 35925/28656 RWalker 5/12/03
            Call ARM_StartWorkflowTransactions(m_lGoToEncSign, m_lGoToSO, m_lGoToAnp)
            'Close current disposition without saving.
            ICHCSII_CoreAPI3_CloseView2 cwiALL_VIEWS
            Call ARM_StopWorkflowTransactions(m_lGoToEncSign, m_lGoToSO, m_lGoToAnp)
              
        Case g_lMenuAddProvider
            'WR IParent.CurrentEncounter() now returns IEncDoc interface SCR 25225
            Set oEncDoc = gobjEncParent.CurrentEncounter
            Call oEncDoc.AddProvider
            
        Case g_lMenuCancel
            'Close current disposition without saving.
            m_bCancel = True
            ICHCSII_CoreAPI3_CloseView2 cwiALL_VIEWS
        
        Case g_lMenuGoalSummary
            Call cwobjCWComm.Message(cwiGENERIC, "show form", NCID_GOAL_SUMMARY, NCID_DISPOSITION_APPLICATION)
            
        Case g_lMenuClearAll 'SCR-35155
            ClearDispositionData
            m_bCancel = True
            ICHCSII_CoreAPI3_CloseView2 cwiALL_VIEWS
            
        Case g_lMenuDischarge
            DischargePatient
        Case Else
    
    End Select 'MenuId
    
    m_lGoToAnp = False
    m_lGoToSO = False
    m_lGoToEncSign = False
    m_bCancel = False

    mbSuppressEvents = False

    Exit Sub

errXIT:

    ReportError "Disp", Err, "MenuItemSelected", vbExclamation
    mbSuppressEvents = False
End Sub 'MenuItemSelected

Public Sub Message(ByVal pintMessageType As Integer, _
                   ByVal pstrMessageText As String, _
                   ByVal pstrSourceServerNCID As String)
                   
    Const sSUBR_NAME$ = "Message"
    '===============================================================================
    '   Purpose:  This method is used by the CW Core to send messages to this OLE Server.
    '             The messages can be from the CW Core or from another OLE Server.
    '
    ' Inputs:
    '   pintMessageType      - Contains the "type" of message to process.
    '                           (NOTE: "As Integer" -- should be "As Long",
    '                           but this change has to be made at the same
    '                           time in all apps or ActiveX dlls in this system.)
    '   pstrMessageText      - Contains any "data" being sent with this message.
    '   pstrSourceServerNCID - Contains the NCID of the OLE Server that "sent"
    '                          the message.
    '                           (NOTE: Not used in this subroutine,
    '                            but we cannot remove it, because this format
    '                            for this subroutine must present the same
    '                            interface as Message( ) in other classes
    '                            in this system.
    '   Returns:  None
    '
    '   Notes:
    'TO DO: Add standard data type prefixes (Hungarian notation)
    ' to the parameters.  NOTE: Cannot do this until all apps are done same
    ' thing at once, because this subroutine's interface must match exactly
    ' that defined in ICHILD.  (This module "Implements" ICHILD.)
    'TO DO: Change "As Integer" to "As Long".  Integers are always converted
    '   to Longs by VB before passing anyway.
    ' NOTE: Cannot do this until all apps are done same
    ' thing at once, because this subroutine's interface must match exactly
    ' that defined in ICHILD.  (This module "Implements" ICHILD.)
    '-------------------------------------------------------------------------
    'MODIFICATIONS:
    '07/26/20 H STINSON SCR-6519: Added error trapping , with breakpoint in IDE-mode,
    '         in order to find odd bug.
    '02/04/01 C Atencio Added new message types OK_TO_CHANGE_PATIENT & OK_TO_CLOSE_VIEW
    '
    '===============================================================================

    
    '---Declare local variables
    
    '-------------------------------------------------------------------------------
    '  Take corresponding actions based upon the message "type".
    '
    '  NOTE: This method is required, but currently there are NO REQUIRED messages
    '        to process.
    '-------------------------------------------------------------------------------
    Select Case CLng(pintMessageType)
        Case cwiSTART_OLE_SERVER
            'Thanks for the info
        Case cwiCOPY
             SendKeys "^c", True
        Case cwiPASTE
             SendKeys "^v", True
                     
        Case cwiREFRESH 'SCR-19131
            'frmDisposition.DisplayDefaultOptions
            
        ' ADD CODE
        Case cwiOK_TO_CHANGE_PATIENT
            If gbDataChanged Then
                'frmDisposition.bSaveDispositionChangesAndOrQuit
                If SaveDispositionData(True) Then 'SCR-31564
                    cwobjCWComm.Message cwiOK_TO_CHANGE_PATIENT, "Y", pstrSourceServerNCID, NCID_DISPOSITION_APPLICATION
                Else
                    cwobjCWComm.Message cwiOK_TO_CHANGE_PATIENT, "N", pstrSourceServerNCID, NCID_DISPOSITION_APPLICATION
                End If
            End If
            'cwobjCWComm.Message cwiOK_TO_CHANGE_PATIENT, "N", pstrSourceServerNCID, NCID_DISPOSITION_APPLICATION
            
        Case cwiOK_TO_CLOSE_VIEW
            If gbDataChanged Then
                'If frmDisposition.bSaveDispositionChangesAndOrQuit Then
                If SaveDispositionData(True) Then 'SCR-31564
                    cwobjCWComm.Message cwiOK_TO_CLOSE_VIEW, "Y", pstrSourceServerNCID, NCID_DISPOSITION_APPLICATION
                Else
                    cwobjCWComm.Message cwiOK_TO_CLOSE_VIEW, "N", pstrSourceServerNCID, NCID_DISPOSITION_APPLICATION
                End If
            End If
            'cwobjCWComm.Message cwiOK_TO_CLOSE_VIEW, "Y", pstrSourceServerNCID, NCID_DISPOSITION_APPLICATION
            
        ' END ADD CODE
        Case Else   '--do nothing--
            'DO NOT display error message if the message from core
            'is not supported to prevent pop up messages
            'when core broadcasts a message:
            '------------------------------------------------------------
            'MsgBxARMd "This application does not support this action", _
            '       vbInformation, "Disposition"
    End Select

   Exit Sub
    
errXIT:

    'CALL THE GENERIC ERROR HANDLER:
    With Err
        Call gobjShared.ShowVBError(.Number, _
                                    .DESCRIPTION, _
                                    msCLASS_NAME$ + "." + sSUBR_NAME$, _
                                    App.Title, _
                                    vbExclamation)
    End With
    
End Sub 'Message

'****************************************************************************************
'  Function Name: OpenView
'
'   Purpose:  This method is called when the CW Core attempts to open a "view"
'             provided by this OLE server.
'
'   Inputs:
'       pintiViewType - Contains the "type" of view being requested by the CW Core
'                       to open.
'       plngParenthWnd - Contains the hWnd of the client area window to be used by
'                        this OLE Server.
'       pblnMainForm - If "viViewType" is a "regular view", upon return this will
'                      contain the window "handle" of the main "regular" form for
'                      this OLE Server.
'   Returns:
'       Returns TRUE if the "view" was opened successfully.
'       Returns FALSE if an error occurred openning the "view".
'
'   Notes:
'03/01/00 RBELL SCR-4005 Move Addmenu call to ResizeView Subroutine
'07/26/20 H STINSON SCR-6519: Added error trapping , with breakpoint in IDE-mode,
'         in order to find odd bug.
'****************************************************************************************
Public Function OpenView(ByVal ViewType As Integer, _
                         ByVal ParenthWnd As Long, _
                         ByRef MainForm As Long) As Boolean
                         
    Dim bRet As Boolean
    
    On Error GoTo errXIT
    
    If gDisplayMode = 0 Then gDisplayMode = Default
    
    '---perform actions based on the "view type"---
    If gobjEncParent Is Nothing And gDisplayMode = Default Then
       'MsgBxARMd "Please select an encounter before opening the disposition module.", vbInformation, "Disposition"
       Screen.MousePointer = vbDefault
       Exit Function
    End If
    
    If ViewType = cwiREGULAR_VIEW Then
        
        With frmDisp
            SetParent .hwnd, ParenthWnd
            MainForm = .hwnd
            
            DispositionOps.SetClinicCapabilities
            bRet = .Display(Me.DispositionOps, mdAdmissionDate, mdLastTransferDate)
            If bRet Then
                Call AddMenus
                Call EnableMenus(True)
                gbDataChanged = True 'SCR-31564
                OpenView = True
            Else
                Unload frmDisp
                OpenView = False
            End If
        End With
    

    End If
        

        
XIT:
    Screen.MousePointer = vbDefault
    Exit Function
    
errXIT:

    Call gobjShared.ShowVBError(Err.Number, Err.DESCRIPTION, msCLASS_NAME$ & ".Openview", App.Title, vbCritical)
    OpenView = False
    Resume XIT
    
    Exit Function
    Resume
    
End Function 'OpenView

Public Sub PatientSelected()
End Sub 'PatientSelected

Public Sub Properties(ByRef pobjLogonEnvironment As Object, _
                      ByRef pobjConfig As Object, _
                      ByVal pstrOwnerNCID As String, _
                      ByVal pblnCalledByDesigner As Boolean)
                      
End Sub 'Properties

'****************************************************************************************
'  Function Name: QueryView
'
'   Purpose: This method is called when a client wants to know if this OLE Server
'            supports a specific "view type".    '
'
'   Inputs:
'       pintViewType - Contains the "type" of view being queried for.
'                       Note: "As Integer" should be changed to "As Long".
'
'  Returns:
'       Returns TRUE if this OLE Server supports the passed in "view type".
'       Returns FALSE if this OLE Server does not support the passed in "view type".
'
'   Notes:
'****************************************************************************************
Public Function QueryView(ByVal pintViewType As Integer) As Boolean

    Const sSUBR_NAME$ = "QueryView"

    'TO DO: Add standard data type prefixes (Hungarian notation)
    ' to the parameters.  NOTE: Cannot do this until all apps are done same
    ' thing at once, because this subroutine's interface must match exactly
    ' that defined in ICHILD.  (This module "Implements" ICHILD.)
    'TO DO: Change "As Integer" to "As Long".  Integers are always converted
    '   to Longs by VB before passing anyway.
    ' NOTE: Cannot do this until all apps are done same
    ' thing at once, because this subroutine's interface must match exactly
    ' that defined in ICHILD.  (This module "Implements" ICHILD.)
    '-------------------------------------------------------------------------
    'MODIFICATIONS:
    '=========================================================================

    Select Case pintViewType
        
        Case cwiREGULAR_VIEW
            QueryView = True    '--this view is NOT optional--
        
        Case cwiPATIENT_SUMMARY_VIEW
            'QueryView = False  'Don't need to set -- Is default value.
        
        Case cwiCLINICIAN_SUMMARY_VIEW
            'QueryView = False  'Don't need to set -- Is default value.
        
        Case cwiENCOUNTER_SUMMARY_VIEW
            'QueryView = False  'Don't need to set -- Is default value.
        
        Case cwiPROPERTY_VIEW
            QueryView = True
            
        Case Else   '--unknown "view type", not supported--
            'QueryView = False  'Don't need to set -- Is default value.
    
    End Select
    
End Function 'QueryView

'****************************************************************************************
'  Sub Name: ResizeView
'
'   Purpose: This method is called when a change has occurred in the "client area"
'            used by this OLE Server, resulting in a need to resize a particular "view".
'
'   Inputs:
'       pintViewType - Contains the "type" of view being requested to resize by the
'                      CW Core.
'       plngLeft - Contains the "left" coordinate for the OLE server.
'       plngTop - Contains the "top" coordinate for the OLE server.
'       plngWidth - Contains the "width" for the OLE server.
'       plngHeight - Contains the "height" for the OLE server.
'
'  Returns:  None
'
'   Notes:
'
'----------------------------------------------------------------------------
'   Modifications:
'03/01/00 RBELL SCR-4005 Added call to add menus
'****************************************************************************************
Public Sub ResizeView(ByVal viViewType As Integer, _
                      ByVal vlLeft As Long, _
                      ByVal vlTop As Long, _
                      ByVal vlWidth As Long, _
                      ByVal vlHeight As Long)
    
    On Error GoTo errXIT
    Dim lMouse As MousePointerConstants
    
    lMouse = Screen.MousePointer
    Screen.MousePointer = vbHourglass

    '---perform different actions based upon the "view type"---
    Select Case viViewType
        Case cwiREGULAR_VIEW
            '---------------------------------------------------------------------------------
            '  Reset the "WindowState" to "Normal", so the form can be resized.  This is
            '    needed because a window CANNOT be moved or resized if it is "minimized" or
            '    "maximized".
            '
            '  NOTE: This generates a "resize" event that is ignored.
            '---------------------------------------------------------------------------------
'            frmDisp.WindowState = vbNormal
'
'            '---hide the form to so that annyoing "screen blinking" will not be seen while controls are moved and resized---
'            frmDisp.Visible = False
'
''            '---resize form to fill entire "client area", NOTE: This generates a "resize" event---
''            If Not goBB.EncParent Is Nothing Then
''               SetParent goBB.VitalsForm.hwnd, goBB.EncParent.hwnd
''            End If
'            frmDisp.Move cwlLeft, cwlTop, cwlWidth, cwlHeight
'
'            frmDisp.Visible = True
            
            'Resize and Show form if required
            frmDisp.ResizeView vlLeft, vlTop, vlWidth, vlHeight
            
        Case Else
            Beep
            '*** MsgBox Fix - Code_Validator auto change made 3/21/2003
            MsgBxARMd "Unsupported resize view type: " & CStr(viViewType) & "ListMangement.ResizeView.", vbExclamation, App.Title
            
        End Select
    
        Screen.MousePointer = lMouse
        
    Exit Sub

errXIT:
    
    Call gobjShared.ShowVBError(Err.Number, Err.DESCRIPTION, msCLASS_NAME$ & ".ResizeView", App.Title, vbCritical)

End Sub 'ResizeView

Private Sub Class_Initialize()

    Const sSUBR_NAME$ = "Class_Initialize"

    
    On Error GoTo errXIT

   '---create shared object used to show VB errors, etc.---
    Set gobjShared = New CWShared

    'SCR 35925/28656 RWalker 5/12/03 "workflow"
    Call ARM_InitAllTransactionsForDisposition
    
    Exit Sub

errXIT:

    Call gobjShared.ShowVBError(Err.Number, Err.DESCRIPTION, msCLASS_NAME$ + "." + sSUBR_NAME$, App.Title, vbExclamation)

End Sub 'Class_Initialize


Private Sub Class_Terminate()

    Const sSUBR_NAME$ = "Class_Terminate"

    Set cwobjCWComm = Nothing
    Set cwobjLogonEnviron = Nothing
    Set cwobjConfig = Nothing
    Set cwobjPatient = Nothing
    Set mobjDisposition = Nothing
    'Set mobjParent = Nothing
    '********************************************************************
    '  TO DO:
    '    Add code here to release reference to all other global objects.
    '********************************************************************
    Set gobjShared = Nothing
    Set gobjSO = Nothing 'ADDED SCR-4269
    
End Sub 'Class_Terminate

Public Sub ICHCSII_CoreAPI2_Initialize2(ByVal plngMyProgID As Long, _
                                        ByRef pobjComm As Object, _
                                        ByRef pobjLogonEnvironment As Object, _
                                        ByRef pobjConfig As Object, _
                                        ByRef pobjPatient As Object, _
                                        ByRef pobjUser As Object)
    
    Const sSUBR_NAME$ = "ICHCSII_CoreAPI2_Initialize2"
    '=================================================================================
    '
    ' Purpose:  This method is called when the CW Core creates an "instance" of this
    '           OLE Server.  It is used to create and initialize global variables
    '           that are shared across all "views" supported by this OLE Server,
    '           i.e. variables that only need to be initialized ONCE, upon object
    '           creation.
    '
    ' Inputs:
    '       plngMyProgId - Contains the "ProgId" that the CW Core has assigned to this
    '                      OLE Server.
    '       pobjComm - Contains a communications object that is used to handle all
    '                  communication between this OLE Server and the CW Core.
    '       pobjLogonEnvironment - Contains the logon environment object.  Contains
    '                              information such as who logged on, where they logged
    '                              on, etc.
    '       pobjConfig - Contains the configuration object that is used to get and set
    '                    configuration information for a specific application.
    '       pobjPatient - Contains the patient object containing the currently selected
    '                     patient.
    '       pobjUser - Security User information object
    '
    '  Returns:  None
    '
    '   Notes:
    '
    'TO DO: Change parameters to "As <explicit object type>"
    '       instead of "As Object" to allow early binding.
    '       WARNING: Must make same change in all other apps & DLLs
    '       that reference this class.
    '=================================================================================

    On Error GoTo errXIT
    
    Set gobjUser = pobjUser
    Call Initialize(plngMyProgID, pobjComm, pobjLogonEnvironment, pobjConfig, pobjPatient)
    
    Exit Sub

errXIT:

    Call gobjShared.ShowVBError(Err.Number, Err.DESCRIPTION, msCLASS_NAME$ + "." + sSUBR_NAME$, App.Title, vbExclamation)

End Sub 'ICHCSII_CoreAPI2_Initialize2
Private Function ICHCSII_CoreAPI3_CloseView2(ByVal viViewType As Integer) As Boolean

    Dim F               As Form
    Dim bFound          As Boolean
    Dim tmpEncParent    As IParent
    Dim oEncDoc         As CHCSIIEncounterOps.Encounter
    Dim bRet            As Boolean
    
    On Error GoTo errXIT
    
    Screen.MousePointer = vbHourglass  'SCR#45993
    
    Select Case viViewType
    Case cwiALL_VIEWS, cwiREGULAR_VIEW
       
       For Each F In Forms
            If F.Name = "frmDisp" Then
                bFound = True
                Exit For
            End If
       Next
       
       Dim obj As IPersist
       
        If bFound Then
            
            'Save data to The Buffer
            If Not m_bCancel Then
                
                'Save Buffer To Wherever It Goes
                If gbDataChanged Then 'SCR-31564
                    bRet = SaveDispositionData(True)
                    If bRet = False Then
                        ICHCSII_CoreAPI3_CloseView2 = False
                        Exit Function
                    End If
                End If
            End If
            
            Set obj = mobjDataRepository.MetaData
            Debug.Print obj.Persist
                        
            Unload frmDisp
            
            Call RemoveMenus
            
            'Do The Encounter Stuff
            Set oEncDoc = gobjEncParent.CurrentEncounter
            
            '--- SCR 84519  Sherry Wang   1/31/2006
            '--- Commented out the bypass SO code
'            oEncDoc.BypassSODialog = True
            
            Set tmpEncParent = gobjEncParent 'Used To Maintain A Handle To Encounter Object In Case OpenSection(eSO) Fails
            
            With tmpEncParent
             
                 .ClosedView
                
                 If m_lGoToEncSign Then '
                     'if oEncDoc.PrimaryProvider.NCID <> gobjLogonEnviron.UserNCID Then
                     'user will get a "You are not authorized to Sign" Message from Encounter
                     'PROMPT THE USER TO SIGN THIS ENCOUNTER:
                     '----------------------------------------
                     With oEncDoc
                         Call .LockingObject.UnlockSection(eDisp, 0)
                         Call .SignEncounter
                     End With
                     
                 ElseIf m_lGoToAnp Then
                     'SCR 16732 - WRogers 9/21/2001
                     Call oEncDoc.LockingObject.UnlockSection(eDisp, 0)
                     Call oEncDoc.OpenSection(eap)
                     
                 ElseIf m_lGoToSO Then
                     Call oEncDoc.LockingObject.UnlockSection(eDisp, 0) 'Need to manually clear section lock when navigating
                     If Not oEncDoc.OpenSection(eSO) Then
                        If Not tmpEncParent Is Nothing Then tmpEncParent.ClosedView 'Causes encounter document to redisplay
                     End If
                     
                 Else
                     cwobjCWComm.Message cwiSTART_OLE_SERVER, gsEMPTY$, NCID_ENCOUNTER_APPLICATION, cwlMyProgId
                     
                 End If
                             
             End With 'gobjEncParent
        
        End If
        
    End Select

XIT:
    ICHCSII_CoreAPI3_CloseView2 = True
    Exit Function
    
errXIT:
    ICHCSII_CoreAPI3_CloseView2 = False
End Function

Private Function IChild_GetAutoCiteNarrative(TestType As String, _
                                             NumPeriods As Integer, _
                                             Period As AC_Period) _
                                        As String
    'No Autocite for Disposition

End Function

Private Function IChild_SetParent(Parent As IParent) As Boolean
    
    Set mobjDataRepository = Nothing
    Set gobjEncParent = Nothing
    Set gobjEncParent = Parent

End Function

'Menu Handlers

Private Sub AddMenus()
    
    Dim objEnc As CHCSIIEncounterOps.Encounter
    
    Set objEnc = gobjEncParent.CurrentEncounter
    
    With gobjEncParent
    
        g_lMenuSave = .AddMenuItem(0, "&Save", "Save.bmp", "Save Disposition")
        
        
        g_lMenuAddProvider = .AddMenuItem(0, "Add &Providers", "link.bmp", "Add Additional Providers")
        
        If gobjUser.HasPrivilegeEx(Priv_S_O, Priv_Write) Then
            g_lMenuAP = .AddMenuItem(0, "&A/P", "Prev.bmp", "Save Disposition And Go To The A/P Screen")
        End If
        
        g_lMenuNext = .AddMenuItem(0, "Si&gn", "DispSign.bmp", "Sign Encounter")
        
        g_lMenuCancel = .AddMenuItem(0, "Ca&ncel", "cancel.bmp", "Close current disposition without saving.")

        If Not objEnc.PrimaryProvider Is Nothing Then
            If (gobjLogonEnviron.UserNCID = CStr(objEnc.PrimaryProvider.NCID)) Then
               Call cwobjCWComm.UpdateMenuItem(g_lMenuNext, "ENABLE")
            Else
               Call cwobjCWComm.UpdateMenuItem(g_lMenuNext, "DISABLE")
            End If
        End If
        
        If Not cwobjCWComm.CmdLineSwitch("CPG") = "" Then
            If gobjUser.HasPrivilegeEx(Priv_Patient_Plan_Of_Care, Priv_Read) Then 'SCR-32190
                g_lMenuGoalSummary = .AddMenuItem(0, "Goal Summary", "Health_History.bmp", "Review patient goal summary")
            End If
        End If
        
        If Len(cwobjCWComm.CmdLineSwitch("WAM")) Then
            g_lMenuClearAll = .AddMenuItem(0, "Clear All", "cancel.bmp", "Clear disposition data.")
        End If
    End With
        
End Sub 'AddMenus

Private Sub RemoveMenus()

    With cwobjCWComm

        If g_lMenuSO > 0 Then
            Call .DeleteMenuItem(g_lMenuSO)
            g_lMenuSO = 0
        End If

        If g_lMenuAP > 0 Then
            Call .DeleteMenuItem(g_lMenuAP)
            g_lMenuAP = 0
        End If

        If g_lMenuSave > 0 Then
            Call .DeleteMenuItem(g_lMenuSave)
            g_lMenuSave = 0
        End If

        If g_lMenuNext > 0 Then
            Call .DeleteMenuItem(g_lMenuNext)
            g_lMenuNext = 0
        End If

        If g_lMenuAddProvider > 0 Then
            Call .DeleteMenuItem(g_lMenuAddProvider)
            g_lMenuAddProvider = 0
        End If

        If g_lMenuCancel > 0 Then
            Call .DeleteMenuItem(g_lMenuCancel)
            g_lMenuCancel = 0
        End If

        If g_lMenuGoalSummary > 0 Then
            Call .DeleteMenuItem(g_lMenuGoalSummary)
            g_lMenuGoalSummary = 0
        End If

    End With

End Sub 'RemoveMenus

Private Function SaveDispositionData(Optional Closing As Boolean = False) As Boolean


    Dim objEnc                      As CHCSIIEncounterOps.Encounter
    Dim sDispositionRTF             As String
    Dim lDispositionValidityStatus  As Long     'SCR-8191 10/05/2000 HS Added.
    Dim sMsg                        As String  'SCR 22873  4/8/02 Phil Crowder
    Dim bValid                      As Boolean
    Dim bRet                        As Boolean
    
    On Error GoTo errXIT
    
    SaveDispositionData = False
    
    'Move all Data From Form To DataRepository
    If frmDisp.Save(Closing) = False Then
        SaveDispositionData = False
        Exit Function
    End If
        
    Set objEnc = gobjEncParent.CurrentEncounter

    With gobjEncParent
        If Not .StartSave Then
'           SCR 16521 - RC 9/25/01 - Refreshed and changed message to include details of who broke lock for this encounter and when
'           MsgBxARMd "Your lock on the disposition section of the current encounter has been taken away by another user."
            objEnc.Refresh


            'SCR 22873  4/8/02 Phil Crowder
            sMsg = "Your lock on the disposition section of the current encounter has been taken away by another user." & vbCrLf

            If Not (objEnc.WhoSigned Is Nothing) Then
              sMsg = sMsg & "This encounter has been completed and signed by " & objEnc.WhoSigned.FullName & _
                " on " & Format$(objEnc.WhenSigned, gsDefaultDateTime) & ".  "
            End If
            'objEnc.se
            MsgBxARMd sMsg & "This Disposition will not be saved to the encounter.", vbExclamation + vbOKOnly, "Disposition"

            SaveDispositionData = True

            GoTo XIT
        End If
        
    End With
    
    bRet = mobjDataRepository.SaveToEncounter
    
    With gobjEncParent
        Call .SetRTF(mobjDataRepository.RTFs.RTFNote, mobjDataRepository.Updated) 'SCR#39592
        Call .SetStatus(mobjDataRepository.DispositionStatus)
        Call .EndSave
        
    End With
    
    UpdateInpatientAdmission
    
    SaveDispositionData = True
    
XIT:
    'Moved here SCR#44687
    'So OK_TO_CLOSE wont process SaveDispositionData SCR-31564
    gbDataChanged = Not Closing
    
    Screen.MousePointer = vbDefault
    Exit Function

errXIT:

    ReportError "Disp", Err, "SaveDispositionData", vbExclamation
    Resume XIT
    
End Function 'SaveDispositionData

Private Function ClearDispositionData() As Boolean

    Dim objEnc                      As CHCSIIEncounterOps.Encounter
    Dim sMsg As String
    Dim bRet As Boolean
    
    On Error GoTo errXIT
    
    ClearDispositionData = False
        
    mobjDataRepository.MetaData.Clear
    
    Set objEnc = gobjEncParent.CurrentEncounter

    With gobjEncParent
        If Not .StartSave Then
            objEnc.Refresh

            sMsg = "Your lock on the disposition section of the current encounter has been taken away by another user." & vbCrLf

            If Not (objEnc.WhoSigned Is Nothing) Then
              sMsg = sMsg & "This encounter has been completed and signed by " & objEnc.WhoSigned.FullName & _
                " on " & Format$(objEnc.WhenSigned, gsDefaultDateTime) & ".  "
            End If
            
            MsgBxARMd sMsg & "This Disposition cannot be removed from the encounter.", vbExclamation + vbOKOnly, "Disposition"

            ClearDispositionData = True

            GoTo XIT
        End If
        
    End With
    
    With objEnc
        .DispositionNCID = vbNullString  'SCR-35278
        .EandMNCID = vbNullString
        .DispMetadata = vbNullString
    End With
    
    With gobjEncParent
        Call .SetRTF(vbNullString)
        Call .SetStatus(Incomplete)
        Call .EndSave
        
    End With

    ClearDispositionData = True
    
XIT:
    Screen.MousePointer = vbDefault
    Exit Function

errXIT:

    ReportError "Disp", Err, "ClearDispositionData", vbExclamation
    Resume XIT
    
End Function 'SaveDispositionData

Private Sub EnableMenus(ByVal bFlag As Boolean)
    Dim sz As String
    
    If bFlag Then
        sz = "ENABLE"
    Else
        sz = "DISABLE"
    End If
    
    gobjComm.UpdateMenuItem g_lMenuSave, sz
    gobjComm.UpdateMenuItem g_lMenuCancel, sz
    
    gobjComm.UpdateMenuItem g_lMenuSO, sz
    gobjComm.UpdateMenuItem g_lMenuAP, sz
    gobjComm.UpdateMenuItem g_lMenuNext, sz
    gobjComm.UpdateMenuItem g_lMenuGoalSummary, sz
    
    'gobjComm.UpdateMenuItem g_lMenuAddProvider, sz
    
    'gobjComm.UpdateMenuItem g_lMenuDischarge, sz
End Sub
'Called by QETelecon
Public Function BuildRTF() As String

    Dim objEnc As CHCSIIEncounterOps.Encounter
    Dim rs As ADODB.Recordset
    Dim sz As String
    
    On Error GoTo errXIT:
    
    'Return blank space so signature line will be displayed SCR-34979
    sz = " "
    
    ' Do no longer include E&M Code in Diposition RTF SCR-35183 - OBE
    'Include E&M Code if Option Selected SCR-38961
    Set objEnc = gobjEncParent.CurrentEncounter
    
    If objEnc.EncProperties.IncludeCodes Or gobjShared.IsAppMode(modeTheater) Then
        Set rs = GetEMTable()
        rs.Find "NCID='" & objEnc.EandMNCID & "'", , adSearchForward, adBookmarkFirst
        If Not rs.EOF Then
            sz = RTF_BOLD_HEADER & "E&M Code: " & RTF_FONT2_16PTS
            sz = sz & rs.Fields("EM_CODE").Value & " (" & rs.Fields("CODE_DESCRIPTION") & ")"
        End If
    End If
    
    BuildRTF = sz
    
    Set objEnc = Nothing
    
    Exit Function
    
errXIT:
    gobjShared.ShowVBError Err.Number, Err.DESCRIPTION, "Disp.BuildTRTF", App.Title, vbExclamation
    BuildRTF = vbNullString
    
End Function 'BuildRTF

Public Sub InitDisplayMode(ByVal DisplayMode As enumDispDisplayMode, Optional ByVal InpatientID As String = "", Optional ByVal AdmissionDate As Date, Optional ByVal LastTransferDate As Date)

    gDisplayMode = DisplayMode

    If gDisplayMode = InpatientDischarge Then
        msInpatientID = InpatientID
        mdAdmissionDate = AdmissionDate 'SCR#47910
        mdLastTransferDate = LastTransferDate
    Else
        'SCR 133684 clear inpatient id if not in inpatient discharge mode
        msInpatientID = ""
    End If
    
'    'TESTING
'    Dim p As Profile
'
'    Set p = New Profile
'
'    p.Comment = "Test edit"
'    p.DESCRIPTION = "Test Dx"
'    p.EndDate = Now
'    p.StartDate = Now
'    p.SnoID = "10"
'    p.Save
    
End Sub

Private Sub UpdateInpatientAdmission()

On Error GoTo ErrHnd

    Dim oPAD As Object
    
    'Only do this for inpatient notes
    
    If msInpatientID <> Empty Then
        Set oPAD = cwobjCWComm.InitializeOLEServer(NCID_INPATIENT_APPLICATION)
    
        oPAD.UpdateAdmissionDischargeDisposition msInpatientID, _
                                frmDisp.DischargeCtl1.DischargeDate, _
                                frmDisp.DischargeCtl1.DischargeComments, _
                                frmDisp.DispositionCtl1.DispositionNCID, _
                                frmDisp.DispositionCtl1.Comment, _
                                frmDisp.Discussed1.GetItemsDiscussed, _
                                frmDisp.Discussed1.Comment, _
                                frmDisp.BillingAdmin1.SnoID, _
                                frmDisp.BillingAdmin1.CheifComplaint, _
                                frmDisp.FollowUp1.TextNote
    End If
    
Exit Sub

ErrHnd:
    gobjShared.ShowVBError Err.Number, Err.DESCRIPTION, "Disp.UpdateInpatientAdmission", App.Title, vbExclamation
    Exit Sub
    Resume
End Sub
Private Sub DischargePatient()
    'Save Discharge info
'    If SaveDischargeInfo Then
'        m_bCancel = True
'        gDisplayMode = 0
'        ICHCSII_CoreAPI3_CloseView2 cwiALL_VIEWS
'        cwobjCWComm.CloseView cwiREGULAR_VIEW, mlngProgID
'    End If
MsgBox "Is this discharging the patient?"

End Sub
