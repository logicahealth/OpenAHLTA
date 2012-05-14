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
Attribute VB_Name = "ReasonForVisit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Implements IChild
Implements ICHCSII_CoreAPI2
Implements ICHCSII_CoreAPI3

Public Event Refresh()
Public Event Display(Message As String)
Public Event ShutDown()
Public Event ToggleTemperatureSettings()
Public Event UpdatePatientType(PatientType As Long)
Public Event PatientChanged()
Public Event Save()
Public Event ShowProperties(OwnerNCID As String)
Public Event Cut()
Public Event Copy()
Public Event Paste()
Private mobjconcept As Object

Public Function getComplaintsClass() As Complaints
    Set getComplaintsClass = New Complaints
End Function


'****************************************************************************************
'   Sub Name: AddMenus                                                                  *
'                                                                                       *
'   Purpose:  Adds this module's menus to the core's Action menu and toolbar.           *
'                                                                                       *
'   Inputs:   None                                                                      *
'                                                                                       *
'   Returns:  None                                                                      *
'                                                                                       *
'   Notes:                                                                              *
'   Modifications:
'   03/01/00   RBELL   SCR-4003 Added Subroutine
'   03/02/00   RBELL   SCR-3974 Added functionality for Save, and Close menu items
'   03/06/00   RBELL   SCR-3969 Changed icon
'****************************************************************************************
Private Sub AddMenus()

End Sub
'****************************************************************************************
'  Sub Name: RemoveMenus                                                                *
'                                                                                       *
'   Purpose: This method gets called to remove the menu items (from the CW Core's main  *
'            menu), which is "owned" by this OLE Server.                                *
'                                                                                       *
'   Inputs:  None                                                                       *
'                                                                                       *
'   Returns: None                                                                       *
'                                                                                       *
'   Notes:   This is an OPTIONAL method.                                                *
'   Modifications:
'   03/01/00 RBELL Added subroutine SCR-4003
'   03/02/00 RBELL SCR-3974 Added functionality for Close, and Save menu items
'****************************************************************************************
Public Sub RemoveMenus()
    
    '---Declare local variables
    If glngAddProvider > 0 Then
      Call goBB.Comm.DeleteMenuItem(glngAddProvider)
      glngAddProvider = 0
    End If

    If glngSave > 0 Then
      Call goBB.Comm.DeleteMenuItem(glngSave)
      glngSave = 0
    End If

    If glngVerifyAllergies > 0 Then
      Call goBB.Comm.DeleteMenuItem(glngVerifyAllergies)
      glngVerifyAllergies = 0
    End If

    If glngNext > 0 Then
      Call goBB.Comm.DeleteMenuItem(glngNext)
      glngNext = 0
    End If

    If glngQuestionnaires > 0 Then
      Call goBB.Comm.DeleteMenuItem(glngQuestionnaires)
      glngQuestionnaires = 0
    End If

    If glngGoalSummary > 0 Then
      Call goBB.Comm.DeleteMenuItem(glngGoalSummary)
      glngGoalSummary = 0
    End If

End Sub
'*******************************************************************************************
'  Function Name: QueryView
'
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
'  Function Name: OpenView
'
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
'   Modifications:
'03/01/00   RBELL   SCR-4003 Added subroutine call to add menus and establish a
'                       connection to the main forms ParentClass property
'**************************************************************************************************
Public Function OpenView(ByVal viViewType As Integer, ByVal vlParenthWnd As Long, rhMainForm As Long) As Boolean

On Error GoTo ErrHandler

    '---perform actions based on the "view type"---
    If goBB.EncParent Is Nothing Then
       MsgBxARMd "Please select an encounter before opening the reason for visit module.", vbInformation, "Reason For Visit"
       Screen.MousePointer = vbDefault
       Exit Function
    ElseIf goBB.EncParent.CurrentEncounter.EncounterType = 3 Then
       'Encounters is displaying its own message SCR-9984
       'MsgBxARMd "The reason for visit module is not available for telephone consults.", vbInformation, "Reason For Visit"
       Screen.MousePointer = vbDefault
       Exit Function
    End If
    
    Select Case viViewType
    Case cwiREGULAR_VIEW
        '---set global flags---
        cwbFirstRegularResize = True
        cwbMainViewRunning = True
        '---------------------------------------------------------------------------------------
        '  If this OLE Server needs it's own menu items, here is a good place to add these
        '    menu items onto the CW Core's main menu.
        '
        '  NOTE: This is OPTIONAL.  You do NOT have to create menu items. Thats why they are
        '        commented out.
        '---------------------------------------------------------------------------------------
        '---add a menu item "Template" on CW's main menu---
        'ghTemplate = gobb.comm.AddPopupMenuToCW("&Template", cwlMyProgId)
        '---add a menu item "Ole Server Menu Item" to the menu "Template"---
        'glOleMenuItem = gobb.comm.AddMenuItem("&Ole Server Menu Item", ghTemplate, cwlMyProgId, "Template menu item")
        
        '-----------------------------------------------------------------------------------------
        '  Reset the "parent window" for this OLE server's "main" regular form.  This is the main
        '    "trick" that makes the OLE server look like it is part of the CW Core.
        '
        '  NOTE: This action invokes the main form's "Load" event, but the form is NOT shown.
        '        Make sure that you DO NOT perform a "Me.Show" in the form's "Load" event, as this
        '        will mess up the "SetParent" API call.  Make sure that you DO NOT "Show" the
        '        "main" regular form in this method, that will happen in the "ResizeView" method.
        '-----------------------------------------------------------------------------------------
        
        
        SetParent goBB.RFVForm.hwnd, vlParenthWnd
        '-- add a reference to the form so the form can access the RFV methods
        goBB.RFVForm.ParentClass = Me  'SCR-4003 ADDED
        '---set parameter---
        rhMainForm = goBB.RFVForm.hwnd
        
        AlertAllergies ("OPEN") 'SCR-12146
               
    Case Else    '--unsupported view "type"--
        '---set return value---
        OpenView = False
        Screen.MousePointer = vbDefault
        Exit Function
    End Select
        
    '---set return value---
    OpenView = True
    Screen.MousePointer = vbDefault
    Exit Function
ErrHandler:
    Call goBB.SharedObj.ShowVBError(Err.number, Err.Description, "ReasonForVisit.OpenView", "ReasonForVisit", vbCritical)
    OpenView = False
    Screen.MousePointer = vbDefault
End Function

'*******************************************************************************************
'  Sub Name: MenuItemSelected
'
'  Description: This method is called when the user has selected a menu item (from the CW
'               Core's main menu), which is "owned" by this OLE Server.
'
'               NOTE: This is an OPTIONAL method.
'
'  Parameters:
'    vlMenuId - Contains the menu id of the menu item the user selected.
'   Modifications:
'   03/01/00 RBELL Enabled menu selection and added the add provider menu
'   03/02/00 RBELL SCR-3974 Added functionality for Save, and close menu items
'*******************************************************************************************
Public Sub MenuItemSelected(ByVal vlMenuId As Long)
On Error GoTo ErrHandler

    '----------------------------------------------------------------------------------
    '  NOTE: You don't have to create menu items, these are optional, that is why
    '        the code is commented out.
    '----------------------------------------------------------------------------------
    '---perform the actions associated with the menu item that was selected---
    Select Case vlMenuId 'UNCOMMENTED SCR-4003
         Case glngVerifyAllergies
            ' added 2/14/2002 WR scr 21188
            Call goBB.Comm.Message(cwiSTART_OLE_SERVER, NCID_ALLERGIES_APPLICATION, NCID_CW_APPLICATION, NCID_REASON_FOR_VISIT)
            Call goBB.Comm.Message(cwiENCOUNTER_SUMMARY_DRILL_DOWN, "", NCID_ALLERGIES_APPLICATION, NCID_REASON_FOR_VISIT)
            gbUpdateEncounterDocument = True
         Case glngAddProvider 'SCR-4003 ADDED
            Call goBB.CurrentEncounter.AddProvider 'SCR-4003 ADDED
         Case glngSave 'SCR-3974 ADDED
            Call goBB.RFVForm.SaveRFV 'SCR-3974 ADDED
         Case glngQuestionnaires
            Call goBB.Comm.Message(cwiGENERIC, "show form", NCID_PATIENT_QUESTIONNAIRES, NCID_REASON_FOR_VISIT)
            gbUpdateEncounterDocument = True
         Case glngGoalSummary
            Call goBB.Comm.Message(cwiGENERIC, "show form", NCID_GOAL_SUMMARY, NCID_REASON_FOR_VISIT)
         Case glngNext
            
            '--- SCR 35925/25485 RWalker 5/12/03 "workflow" Tx flags
            '--- mere presence of Tx in Tivoli log, not timed
            goARM.StartFarTx gsARMTx_NavigateBtnRFV2Vitals
            goARM.StopFarTx gsARMTx_NavigateBtnRFV2Vitals, enmARMTxStatusOk, enmARMStop_MeOnly
            
            Call goBB.RFVForm.SaveRFV
            goBB.RFVForm.mGoToNextScreen = True
            Unload goBB.RFVForm
            goBB.RFVClass.RemoveMenus
            
        Case Else   '--do nothing-- 'UNCOMMENTED SCR-4003
    End Select 'UNCOMMENTED SCR-4003
    
    Exit Sub
ErrHandler:
    Call goBB.SharedObj.ShowVBError(Err.number, Err.Description, "ReasonForVisit.MenuItemSelected", "ReasonForVisit", vbExclamation)
End Sub

'**************************************************************************************************************
'  Sub Name: PatientSelected
'
'  Description: This method gets called when a new patient has been selected from the CW Core.
'
'               NOTE: This is an OPTIONAL method.
'
'  Parameters:
'    none
'**************************************************************************************************************
Public Sub PatientSelected()
    Dim oCollectionItem As Object
    
    If goBB.Patient.UnitNumber <> goBB.Patient.UnitNumber Then
        Screen.MousePointer = vbHourglass
        
        If Not gcolProblems Is Nothing Then
           For Each oCollectionItem In gcolProblems
                gcolProblems.Remove 1
           Next
        End If
        If Not gcolClinicFavorites Is Nothing Then
           For Each oCollectionItem In gcolClinicFavorites
                gcolClinicFavorites.Remove 1
           Next
        End If
        
        If cwbMainViewRunning Then
            Call goBB.RFVForm.RefreshNewPatient
            '---add code for changing encounter number
        End If
      
        Screen.MousePointer = vbDefault
    End If

End Sub


'*******************************************************************************************
'  Sub Name: ResizeView
'
'  Description: This method is called when a change has occurred in the "client area"
'               used by this OLE Server, resulting in a need to resize a particular "view".
'
'  Parameters:
'    viViewType - Contains the "type" of view being requested to resize by the CW Core.
'    vlLeft - Contains the "left" coordinate for the OLE server.
'    vlTop - Contains the "top" coordinate for the OLE server.
'    vlWidth - Contains the "width" for the OLE server.
'    vlHeight - Contains the "height" for the OLE server.
'   Modifications:
'03/01/00 RBELL SCR-4003 Added menus
'*******************************************************************************************
Public Sub ResizeView(ByVal viViewType As Integer, ByVal vlLeft As Long, ByVal vlTop As Long, ByVal vlWidth As Long, ByVal vlHeight As Long)
On Error GoTo ErrHandler
    '---reset global "client area" parameters---
    cwlLeft = vlLeft
    cwlTop = vlTop
    cwlWidth = vlWidth
    cwlHeight = vlHeight
    
    'If the app is not running then ignore this call from core..
    If Not cwbMainViewRunning Then Exit Sub          'SCR 1196
    
    '---perform different actions based upon the "view type"---
    If viViewType = cwiREGULAR_VIEW Then
        If Not cwbFirstRegularResize Then
            '---------------------------------------------------------------------------------
            '  Reset the "WindowState" to "Normal", so the form can be resized.  This is
            '    needed because a window CANNOT be moved or resized if it is "minimized" or
            '    "maximized".
            '
            '  NOTE: This generates a form "resize" event.
            '---------------------------------------------------------------------------------
            goBB.RFVForm.WindowState = vbNormal
            
            '---hide the form to so that annyoing "screen blinking" will not be seen while controls are moved and resized---
            goBB.RFVForm.Visible = False
            
            '---resize form to fill entire "client area", NOTE: This generates a "resize" event---
            If Not goBB.EncParent Is Nothing Then
               SetParent goBB.RFVForm.hwnd, goBB.EncParent.hwnd
            End If
            goBB.RFVForm.Move cwlLeft, cwlTop, cwlWidth, cwlHeight
                
            goBB.RFVForm.Visible = True
        Else    '--first "regular" resize, need to perform "initailization" procedures--
            '---reset flag so this code will not be executed again---
            cwbFirstRegularResize = False
            '----------------------------------------------------------------------------
            '  NOTE: Here is where you "Show" the "Regular View" form for the FIRST time.
            '----------------------------------------------------------------------------
            goBB.RFVForm.Display
        End If
        
    End If
    
    Exit Sub
ErrHandler:
    Call goBB.SharedObj.ShowVBError(Err.number, Err.Description, "ReasonForVisit.ResizeView", "ReasonForVisit", vbExclamation)
End Sub
'**************************************************************************************************************
'  Sub Name: Properties
'
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
Public Sub Properties(robjLogonEnvironment As Object, robjConfig As Object, ByVal vsOwnerNCID As String, ByVal vbCalledByDesigner As Boolean)
On Error GoTo ErrHandler

    
    Exit Sub
ErrHandler:
    Call goBB.SharedObj.ShowVBError(Err.number, Err.Description, "ReasonForVisit.Properties", "ReasonForVisit", vbExclamation)
End Sub

Private Sub Class_Initialize()
    Call ARM_InitAllTransactionsForReasonForVisit
    
    gsDateFormat = "DD MMM YYYY"
    
    If goBB Is Nothing Then Set goBB = New Blackboard
    If goBB.RFVClass Is Nothing Then Set goBB.RFVClass = Me
    
End Sub

Private Sub Class_Terminate()

    '---release reference to objects---
    Set goBB = Nothing

End Sub



'**************************************************************************************************************
'  Sub Name: Message
'
'  Description: This method is used by the CW Core to send messages to this OLE Server.  The
'               messages can be from the CW Core or from another OLE Server.
'
'  Parameters:
'    viMessageType - Contains the "type" of message to process.
'    vsMessageText - Contains any "data" assoicated with this message.
'    vsSourceServerNCID - Contains the NCID of the OLE Server that "sent" the message.
'**************************************************************************************************************
Public Sub Message(ByVal viMessageType As Integer, ByVal vsMessageText As String, ByVal vsSourceServerNCID As String)
On Error GoTo ErrHandler

    '-------------------------------------------------------------------------------
    '  Take corresponding actions based upon the message "type".
    '  NOTE: This method is required, but currently there are NO REQUIRED messages to process.
    '-------------------------------------------------------------------------------
    Select Case viMessageType
    
        Case cwiSTART_OLE_SERVER
        'Thanks for the notification
        Case cwiCOPY
             SendKeys "^c", True
        Case cwiPASTE
             SendKeys "^v", True
        Case cwiGENERIC
             If vsMessageText = "SAVE_DATA" Then
                If Not goBB.EncParent Is Nothing Then
                   goBB.RFVForm.SaveRFV
                End If
             'Added WR per SCR 23323, 20280
             ElseIf vsMessageText = "REFRESH" Then
                'to capture changes to Special Work Status
                If frmSWS.Visible Then
                    'frmSWS.Display
                    'frmSWS.LoadSWS
                    ' SCR #50456;   Developer: MAG 02/26/2004
                    goBB.Comm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_REASON_FOR_VISIT
                    frmSWS.Show vbModal, Me
                    goBB.Comm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_REASON_FOR_VISIT
                End If
             End If
        Case Else   '--do nothing--
            'MsgBxARMd "Reason for Visit does not support this action", vbInformation, "Reason for Visit"
    End Select

    Exit Sub
ErrHandler:
    Call goBB.SharedObj.ShowVBError(Err.number, Err.Description, "ReasonForVisit.Message", "ReasonForVisit", vbExclamation)
End Sub

'****************************************************************************************************
'  Sub Name: Initialize
'
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
Public Sub Initialize(ByVal vlMyProgId As Long, robjComm As Object, robjLogonEnvironment As Object, robjConfig As Object, robjPatient As Object)
On Error GoTo ErrHandler
    
    '---save the "ProgId" assigned to me by the CW Core---
    cwlMyProgId = vlMyProgId
    
    '---create references to passed in objects---
    Set goBB.Comm = robjComm
    Set goBB.Logon = robjLogonEnvironment
    Set goBB.Config = robjConfig
    Set goBB.Patient = robjPatient
    
    App.HelpFile = goBB.SharedObj.CWHelpFile 'SCR-HELP

    '---set global flags---
    cwbMainViewRunning = False
    
    '---variable to ignore click event when loading data in lists
    gbIgnoreClickEvent = True
       
    Exit Sub
ErrHandler:
    Call goBB.SharedObj.ShowVBError(Err.number, Err.Description, "ReasonForVisit.Initialize", "ReasonForVisit", vbExclamation)
End Sub

Public Sub ICHCSII_CoreAPI2_Initialize2(ByVal vlMyProgId As Long, robjComm As Object, robjLogonEnvironment As Object, robjConfig As Object, robjPatient As Object, robjUser As Object)

    Call Initialize(vlMyProgId, robjComm, robjLogonEnvironment, robjConfig, robjPatient)
    Set goBB.User = robjUser
    
End Sub

'*******************************************************************************************
'  Sub Name: ICHCSII_CoreAPI3_CloseView2
'
'  Description: This method is called when the CW Core closes a "view" provided by this OLE
'               Server.
'
'  Parameters:
'    viViewType - Contains the "type" of view that the CW Core is closing.
'
' CHANGE HISTORY:
' * MSolano  1/`/2002 SCR-195`
'    The closing of an OLE child server starts in Core:
'       CW.frmCW.cmdModuleClose_Click()
'    --- ARM FAR CALLS:
'    The ARM start transaction is instrumented as a far call [ARM.StartFarTx()].
'    This gives the OLE server the opportunity to detect in class's CloseView()
'    if transaction fails (business rules) or aborts (run time error).
'    Resolve ARM start w/ ARM.StopFarTx() either here (status ok)
'    or OLE Server class' CloseView() (w/ appropriate status).
'    (ARM documentation [VSS/Source ARM folder] describes far calls & status flags.)
'*******************************************************************************************
Public Function ICHCSII_CoreAPI3_CloseView2(ByVal viViewType As Integer) As Boolean
    Dim lARMStatus  As enmARMTxStatus   'SCR-ARM-19529 MSolano 12/20/01
    
    lARMStatus = enmARMTxStatusFailed   'SCR-ARM-19529 MSolano 12/20/01
    
    'Innocent till guilty
    ICHCSII_CoreAPI3_CloseView2 = True
    
    If Not gSavingRFV Then
        Select Case viViewType
        Case cwiALL_VIEWS, cwiREGULAR_VIEW
            '---close the "main view"---
            If cwbMainViewRunning Then
            
                AlertAllergies ("CLOSE") 'SCR-12146
                '---set global flags---
                cwbMainViewRunning = False
                
                '---this indicates that the CW Core is closing the "view", NOT this OLE Server---
                cwbCloseViewMyself = False
                
                '---Save Data
                If goBB.RFVForm.SaveRFV() Then
                    '---close ALL forms that are associated with all "views"---
                    Unload goBB.RFVForm
                    '--- SCR-ARM-19529 MSolano 12/20/01 Success at last
                    lARMStatus = enmARMTxStatusOk
                Else
                    ICHCSII_CoreAPI3_CloseView2 = False
                End If
                
            End If
            
        End Select
    End If

    '<<< SCR-ARM-19529 MSolano 1/17/2002
    If lARMStatus = enmARMTxStatusFailed Then
        '--- SCR 22578 Close Reason For Visit via Ok btn (vs. Actions menu) MSolano 4/25/2002
        '--- can close via 2 paths: core/Actions; RFV/OkButton; ARM Stop's are ignored for path not used
        '--- SCR-ARM-48428 RWalker 1/23/2004 Converted to Complex for Multiple Entry Point from Core and Object
        goARM.StopCmplxTx NCID_REASON_FOR_VISIT & gsARM_ACT_CLOSE, enmARMTxStatusFailed, enmARMStop_MeOnly, gsARM_ACT_MENU
    End If

End Function

Private Function IChild_GetAutoCiteNarrative(TestType As String, NumPeriods As Integer, Period As ENCOUNTER_INTERFACES.AC_Period) As String
    'No Autocites for RFV
End Function

Private Function IChild_SetParent(Parent As ENCOUNTER_INTERFACES.IParent) As Boolean
    Set goBB.EncParent = Parent
End Function

' Checks to see if Allergies is running
' if running and screening is opening then allow user to check verifies allergies
' if running and screening is closing then disable check box
' SCR-12146
Private Function AlertAllergies(Action As String)
   
    If goBB.Comm.Message(cwiOLE_SERVER_RUNNING, NCID_ALLERGIES_APPLICATION, NCID_CW_APPLICATION, NCID_REASON_FOR_VISIT_APPLICATION) Then
        Select Case Action
             Case "CLOSE"
                goBB.AllergyObj.Message cwiGENERIC, "SCREENING CLOSE", NCID_REASON_FOR_VISIT_APPLICATION
            Case "OPEN"
                goBB.AllergyObj.Message cwiGENERIC, "SCREENING OPEN", NCID_REASON_FOR_VISIT_APPLICATION
        End Select
     End If
        
End Function

