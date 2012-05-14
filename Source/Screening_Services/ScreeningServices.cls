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
Attribute VB_Name = "Service_Maintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements ICHCSII_CoreAPI2

'*************
'CONSTANTS
'*************
#If DebugOn Then
      Private Const Module = "ScreeningServices"
#End If

Public Enum SearchTypes
        All = 0
        Cholesterol
        Fecal_Blood
        Mammogram
        PAP
        PSA
End Enum

Public Enum NotificationMethods
    eMail = 0
    LetterOnly
    eMailOrLetter
    eMailAndLetter
    eByTelephone
    eByLetterNoPrint
    eByEmailNoSend
End Enum

Public Enum LabelOptions
    Envelope = 0
    Label
End Enum

Public Type tPlanOfCare
    POCDEF_ID As String
    Name As String
End Type

Public Type PatientServicesRecord
    Unit_Number As String
    Notified As Boolean
    eMail_Address As String
    Full_Name As String
    SSN As String
    Area_Code As String
    Home_Phone As String
    Address As String
    City As String
    State As String
    Post_Code As String
    Cholesterol_Date As Date
    Fecal_Blood_Date As Date
    Mammogram_Date As Date
    PAP_Date As Date
    PSA_Date As Date
    Last_Notification_Date As Date
    Last_Notification_Method As Integer
    Processed_This_Run As Boolean
    PCM_Name As String
End Type
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

    Dim bSelect As Boolean
    #If DebugOn Then
        Dim DR As DebugRec
        DL.DebugLog Exe, Module, "AddMenus", "Begin", DR, "ProgId", cwlMyProgId
    #End If

    If cwobjCWComm Is Nothing Then Exit Sub
    RemoveMenus
    '---Declare local variables
    'glngOptions = cwobjCWComm.AddActionMenuItem("Search", "configure.bmp", cwlMyProgId, "Set Search Reviewed")
    
    If Not gMTFConfig Then
        glngSelect = cwobjCWComm.AddActionMenuItem("Select All", "reorder.bmp", cwlMyProgId, "Select All Results")
        glngDeselect = cwobjCWComm.AddActionMenuItem("Deselect All", "resolve.bmp", cwlMyProgId, "Deselect All Results")
    End If
    glngResults = cwobjCWComm.AddActionMenuItem("Refresh", "refresh.bmp", cwlMyProgId, "Search for services due")
    'If Not gMTFConfig Then
    glngNotify = cwobjCWComm.AddActionMenuItem("Notify", "Reports.bmp", cwlMyProgId, "Notify selected patients of screening services")
    
    giPrintMenuId = cwobjCWComm.AddPrintMenuItem("Reminders", cwlMyProgId, "Print the selected reminders")
    giPrintPreviewMenuId = cwobjCWComm.AddPrintPreviewMenuItem("Reminders", cwlMyProgId, "Preview the selected reminders")
    cwobjCWComm.UpdateMenuItem giPrintMenuId, "DISABLE"
    cwobjCWComm.UpdateMenuItem giPrintPreviewMenuId, "DISABLE"

    #If DebugOn Then
        DL.DebugLog Exe, Module, "AddMenus", "End", DR, "ProgId", cwlMyProgId
    #End If

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
'****************************************************************************************
Public Sub RemoveMenus()

    #If DebugOn Then
        Dim DR As DebugRec
        DL.DebugLog Exe, Module, "RemoveMenus", "Begin", DR, "ProgId", cwlMyProgId
    #End If
    
    '---Declare local variables
    If glngOptions > 0 Then
      Call cwobjCWComm.DeleteMenuItem(glngOptions)
      glngOptions = 0
    End If

    If glngSelect > 0 Then
      Call cwobjCWComm.DeleteMenuItem(glngSelect)
      glngSelect = 0
    End If

    If glngDeselect > 0 Then
      Call cwobjCWComm.DeleteMenuItem(glngDeselect)
      glngDeselect = 0
    End If

    If glngResults > 0 Then
      Call cwobjCWComm.DeleteMenuItem(glngResults)
      glngResults = 0
    End If

    If glngNotify > 0 Then
      Call cwobjCWComm.DeleteMenuItem(glngNotify)
      glngNotify = 0
    End If
    
    cwobjCWComm.DeleteMenuItem giPrintMenuId
    giPrintMenuId = 0
    
    cwobjCWComm.DeleteMenuItem giPrintPreviewMenuId
    giPrintPreviewMenuId = 0
    
    #If DebugOn Then
        DL.DebugLog Exe, Module, "RemoveMenus", "End", DR, "ProgId", cwlMyProgId
    #End If

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

    #If DebugOn Then
        Dim DR As DebugRec
    #End If

    #If DebugOn Then
        DL.DebugLog Exe, Module, "QueryView", "P-Before QueryView", DR
    #End If

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
        Case Else   '--unknown "view type", not supported--
            QueryView = False
    End Select

    #If DebugOn Then
        DL.DebugLog Exe, Module, "QueryView", "P-After QueryView", DR
    #End If

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
'**************************************************************************************************
Public Function OpenView(ByVal viViewType As Integer, ByVal vlParenthWnd As Long, rhMainForm As Long) As Boolean

    On Error GoTo ErrHandler
    OpenView = False
    
    #If DebugOn Then
        Dim DR As DebugRec
        Set DL = GetObject("", "DLog.cDebugLog")
        DL.debugStart
    #End If

    #If DebugOn Then
        DL.DebugLog Exe, Module, "OpenView", "P-Before OpenView", DR
    #End If
    
    '---perform actions based on the "view type"---
    Select Case viViewType
    Case cwiREGULAR_VIEW
        '---set global flags---
        cwbFirstRegularResize = True
        cwbMainViewRunning = True
       
        '-----------------------------------------------------------------------------------------
        '  Reset the "parent window" for this OLE server's "main" regular form.  This is the main
        '    "trick" that makes the OLE server look like it is part of the CW Core.
        '
        '  NOTE: This action invokes the main form's "Load" event, but the form is NOT shown.
        '        Make sure that you DO NOT perform a "Me.Show" in the form's "Load" event, as this
        '        will mess up the "SetParent" API call.  Make sure that you DO NOT "Show" the
        '        "main" regular form in this method, that will happen in the "ResizeView" method.
        '-----------------------------------------------------------------------------------------
        SetParent frmScreeningServices.hwnd, vlParenthWnd
        frmScreeningServices.ParentClass = Me
        '---set parameter---
        rhMainForm = frmScreeningServices.hwnd
    Case Else    '--unsupported view "type"--
        GoTo CleanUpAndExit
    End Select
        
    '---set return value---
    OpenView = True
    #If DebugOn Then
        DL.DebugLog Exe, Module, "OpenView", "P-After OpenView", DR
    #End If
    GoTo CleanUpAndExit
    
ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "ScreeningServices.OpenView", "ScreeningServices", vbCritical)
    
CleanUpAndExit:
    '---set return value---
    Screen.MousePointer = vbDefault

End Function

'*******************************************************************************************
'  Sub Name: CloseView
'
'  Description: This method is called when the CW Core closes a "view" provided by this OLE
'               Server.
'
'  Parameters:
'    viViewType - Contains the "type" of view that the CW Core is closing.
'*******************************************************************************************
Public Sub CloseView(ByVal viViewType As Integer)
    #If DebugOn Then
        Dim DR As DebugRec
    #End If

    #If DebugOn Then
        DL.DebugLog Exe, Module, "CloseView", "P-Before CloseView", DR
    #End If
    
    If viViewType = cwiALL_VIEWS Then
        '---close the "main view"---
        If cwbMainViewRunning Then
           cwbMainViewRunning = False  '---set global flags---
           cwbCloseViewMyself = False '---this indicates that the CW Core is closing the "view", NOT this OLE Server---
           Unload frmScreeningServices '---close ALL forms that are associated with all "views"---
        End If
    ElseIf viViewType = cwiREGULAR_VIEW Then
        '---close the "main view"---
        If cwbMainViewRunning Then
           cwbMainViewRunning = False '---set global flags---
           cwbCloseViewMyself = False '---this indicates that the CW Core is closing the "view", NOT this OLE Server---
           Unload frmScreeningServices '---close ALL forms that are associated with all "views"---
        End If
    End If
    
    #If DebugOn Then
        DL.DebugLog Exe, Module, "CloseView", "P-After CloseView", DR
    #End If
    
End Sub

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
'*******************************************************************************************
Public Sub MenuItemSelected(ByVal vlMenuId As Long)
    On Error GoTo ErrHandler
    #If DebugOn Then
        Dim DR As DebugRec
    #End If

    #If DebugOn Then
        DL.DebugLog Exe, Module, "MenuItemSelected", "P-Before MenuItemSelected", DR
    #End If
    
    '---perform the actions associated with the menu item that was selected---
    Select Case vlMenuId
        Case glngOptions
        Case glngResults
            frmScreeningServices.ShowResults
        Case glngSelect
            frmScreeningServices.SelectItems True
        Case glngDeselect
            frmScreeningServices.SelectItems False
        Case glngNotify
            frmScreeningServices.Notify_Patients
        Case giPrintMenuId
            frmScreeningServices.PrintReminders False
        Case giPrintPreviewMenuId
            frmScreeningServices.PrintReminders True
        Case Else
    End Select
    
    #If DebugOn Then
        DL.DebugLog Exe, Module, "MenuItemSelected", "P-After MenuItemSelected", DR
    #End If
    Exit Sub

ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "ScreeningServices.MenuItemSelected", "ScreeningServices", vbExclamation)

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
    
    #If DebugOn Then
        Dim DR As DebugRec
    #End If

    #If DebugOn Then
        DL.DebugLog Exe, Module, "PatientSelected", "P-Before PatientSelected", DR
    #End If
    Dim oCollectionItem As Object
    
    If Not glCurPat = cwobjPatient.UnitNumber Then
       Screen.MousePointer = vbHourglass
       glCurPat = cwobjPatient.UnitNumber
       Screen.MousePointer = vbDefault
    End If
    
    #If DebugOn Then
        DL.DebugLog Exe, Module, "PatientSelected", "P-After PatientSelected", DR
    #End If

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
'*******************************************************************************************
Public Sub ResizeView(ByVal viViewType As Integer, ByVal vlLeft As Long, ByVal vlTop As Long, ByVal vlWidth As Long, ByVal vlHeight As Long)
    
    On Error GoTo ErrHandler
    #If DebugOn Then
        Dim DR As DebugRec
    #End If

    #If DebugOn Then
        DL.DebugLog Exe, Module, "ResizeView", "P-Before ResizeView", DR
    #End If
    '---reset global "client area" parameters---
    cwlLeft = vlLeft
    cwlTop = vlTop
    cwlWidth = vlWidth
    cwlHeight = vlHeight
    
    'If the app is not running then ignore this call from core..
    If Not cwbMainViewRunning Then Exit Sub
    
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
            If Not cwobjCWComm Is Nothing Then If cwobjCWComm.Message(cwiMODAL_FORM_SHOWING, "", NCID_CW_APPLICATION, NCID_SCREENING_SERVICES_APPLICATION) Then Exit Sub
            frmScreeningServices.WindowState = vbNormal
            '---hide the form to so that annyoing "screen blinking" will not be seen while controls are moved and resized---
            frmScreeningServices.Visible = False
            '---resize form to fill entire "client area", NOTE: This generates a "resize" event---
            frmScreeningServices.Move cwlLeft, cwlTop, cwlWidth, cwlHeight
            frmScreeningServices.Visible = True
        Else    '--first "regular" resize, need to perform "initailization" procedures--
            '---reset flag so this code will not be executed again---
            cwbFirstRegularResize = False
            '----------------------------------------------------------------------------
            '  NOTE: Here is where you "Show" the "Regular View" form for the FIRST time.
            '----------------------------------------------------------------------------
            frmScreeningServices.Display
            Call AddMenus
        End If
        
    End If
    #If DebugOn Then
        DL.DebugLog Exe, Module, "ResizeView", "P-After ResizeView", DR
    #End If
    Exit Sub
    
ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "ScreeningServices.ResizeView", "ScreeningServices", vbExclamation)
    
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
    #If DebugOn Then
        Dim DR As DebugRec
    #End If

    #If DebugOn Then
        DL.DebugLog Exe, Module, "Properties", "P-Before Properties", DR
    #End If
    '*****************************************************************
    '  TO DO:
    '    Add code here to invoke this OLE Server's properties page.
    '*****************************************************************
    
    #If DebugOn Then
        DL.DebugLog Exe, Module, "Properties", "P-After Properties", DR
    #End If

    Exit Sub
    
ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "ScreeningServices.Properties", "ScreeningServices", vbExclamation)
    
    Exit Sub
End Sub

Private Sub Class_Initialize()
    #If DebugOn Then
        Dim DR As DebugRec
        Set DL = GetObject("", "DLog.cDebugLog")
        DL.debugStart
    #End If


    #If DebugOn Then
        DL.DebugLog Exe, Module, "Class_Initialize", "P-Before Class_Initialize", DR
    #End If
    '---create shared object used to show VB errors, etc.---
    Set gobjShared = New CWShared
    Set gSearchOptions = New ScreeningOptions
    '---set the help file for this application---
    App.HelpFile = gobjShared.CWHelpFile
   
   '---Create OLE objects
    
    '------------------------------------
    '  This statement is needed to debug
    '    this OLE Server.
    '------------------------------------
    #If ciDEBUG = 1 Then
        If cwbDebugLogon = False Then
            Exit Sub
        End If
    #End If

    #If DebugOn Then
        DL.DebugLog Exe, Module, "Class_Initialize", "P-After Class_Initialize", DR
    #End If
    
End Sub

Private Sub Class_Terminate()

    #If DebugOn Then
        Dim DR As DebugRec
    #End If

    #If DebugOn Then
        DL.DebugLog Exe, Module, "Class_Terminate", "P-Before Class_Terminate", DR
    #End If
    
    '---release reference to objects---
    Set gobjSQL = Nothing
    Set gobjShared = Nothing
    Set cwobjCWComm = Nothing
    Set cwobjLogonEnviron = Nothing
    Set gcwobjUser = Nothing
    Set cwobjConfig = Nothing
    Set cwobjPatient = Nothing
    Set gcwobjUser = Nothing
    '********************************************************************
    '  TO DO:
    '    Add code here to release reference to all other global objects.
    '********************************************************************
    '----------------------------------------
    '  If debugging, make sure to "log off".
    '----------------------------------------
    #If ciDEBUG = 1 Then
        cwbDebugLogOff
    #End If

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

    #If DebugOn Then
        Dim DR As DebugRec
    #End If

    #If DebugOn Then
        DL.DebugLog Exe, Module, "Message", "P-Before Message", DR
    #End If
    
    '-------------------------------------------------------------------------------
    '  Take corresponding actions based upon the message "type".
    '
    '  NOTE: This method is required, but currently there are NO REQUIRED messages
    '        to process.
    '-------------------------------------------------------------------------------
    Select Case viMessageType
        Case cwiSTART_OLE_SERVER
        'Thanks for the notification
        Case cwiCOPY
             SendKeys "^c", True
        Case cwiPASTE
             SendKeys "^v", True
        Case cwiGENERIC
             If vsMessageText = "" Then
             End If
        Case cwiPROPERTIES
            frmScreeningServices.Properties cwobjLogonEnviron.UserNCID, False
            DoEvents
            frmScreeningServices.grdResults.Refresh
        Case cwiREFRESH_DATA
            'frmScreeningServices.ShowResults
        Case Else   '--do nothing--
            'MsgBox "Reason for Visit does not support this action", vbInformation, "Reason for Visit"
    End Select


    #If DebugOn Then
        DL.DebugLog Exe, Module, "Message", "P-After Message", DR
    #End If
    Exit Sub
    
ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "ScreeningServices.Message", "ScreeningServices", vbExclamation)

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
    #If DebugOn Then
        Dim DR As DebugRec
    #End If

    #If DebugOn Then
        DL.DebugLog Exe, Module, "Initialize", "P-Before Initialize", DR
    #End If
    '---save the "ProgId" assigned to me by the CW Core---
    cwlMyProgId = vlMyProgId
    
    '---create references to passed in objects---
    Set cwobjCWComm = robjComm
    Set cwobjLogonEnviron = robjLogonEnvironment
    Set cwobjConfig = robjConfig
    Set cwobjPatient = robjPatient

    If gobjSQL Is Nothing Then Call SetSQLOpsObject(gobjSQL)
    
    '---get user's default configuration values---
    Call gGetDefaultConfigValues
    
    '---Instantiate Help-file object
    App.HelpFile = gobjShared.CWHelpFile 'SCR-HELP
    
    '---set global flags---
    cwbMainViewRunning = False
    
    '****************************************************************************
    '  TO DO:
    '    Add code here to create objects, initialize global variables, etc.
    '****************************************************************************
    '---variable to ignore click event when loading data in lists
    #If DebugOn Then
        DL.DebugLog Exe, Module, "Initialize", "P-After Initialize", DR
    #End If
       
    Exit Sub

ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "ScreeningServices.Initialize", "ScreeningServices", vbExclamation)

End Sub


Public Sub ICHCSII_CoreAPI2_Initialize2(ByVal vlMyProgId As Long, robjComm As Object, robjLogonEnvironment As Object, robjConfig As Object, robjPatient As Object, robjUser As Object)
'// handle this guy , robjSecurity As Object)
    
    #If DebugOn Then
        Dim DR As DebugRec
    #End If

    #If DebugOn Then
        DL.DebugLog Exe, Module, "ICHCSII_CoreAPI2_Initialize2", "P-Before ICHCSII_CoreAPI2_Initialize2", DR
    #End If
    
    Call Initialize(vlMyProgId, robjComm, robjLogonEnvironment, robjConfig, robjPatient)
    Set gcwobjUser = robjUser
    
    Set gSearch = New CHCSII_Search.search
    gSearch.Initialize cwobjCWComm, cwobjLogonEnviron, cwobjConfig, cwobjPatient, cwobjUser
    
    gMTFSecurity = True
    gPCMSecurity = True
    
    #If DebugOn Then
        DL.DebugLog Exe, Module, "ICHCSII_CoreAPI2_Initialize2", "P-After ICHCSII_CoreAPI2_Initialize2", DR
    #End If
End Sub

