VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Reports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements ICHCSII_CoreAPI2

Rem Public gobjshared As CWShared
Dim rsValue As String

 
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
        Case cwiPROPERTY_VIEW
            If Not gPopulationSecurity Then
               QueryView = False
            Else
               QueryView = True
            End If
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
'**************************************************************************************************
Public Function OpenView(ByVal viViewType As Integer, ByVal vlParenthWnd As Long, rhMainForm As Long) As Boolean
   
   On Error GoTo ErrHandler
   Screen.MousePointer = vbHourglass
    OpenView = False

    '---perform actions based on the "view type"---
    If viViewType = cwiREGULAR_VIEW Then
        '---set global flags---
        cwbFirstRegularResize = True
        cwbMainViewRunning = True
        
        'Load frmReports
        bDontTerminate = True
        frmReports.SetMyParent = Me
        bDontTerminate = False
        SetParent frmReports.hwnd, vlParenthWnd
        rhMainForm = frmReports.hwnd '---set parameter---
    Else    '--unsupported view "type"--
        OpenView = False '---set return value---
        GoTo CleanUpAndExit
    End If
        
    '---set return value---
    OpenView = True
    GoTo CleanUpAndExit
    
ErrHandler:
    '---show the VB error---
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "Reports.OpenView", "CHCSII_Reports", vbCritical)
    
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
    If viViewType = cwiALL_VIEWS Then
        '---close the "main view"---
        If cwbMainViewRunning Then
            '---set global flags---
            cwbMainViewRunning = False
            
            '---this indicates that the CW Core is closing the "view", NOT this OLE Server---
            cwbCloseViewMyself = False
            
            '---close ALL forms that are associated with all "views"---
            Unload frmReports
        End If
    ElseIf viViewType = cwiREGULAR_VIEW Then
        '---close the "main view"---
        If cwbMainViewRunning Then
            '---set global flags---
            cwbMainViewRunning = False
            
            '---this indicates that the CW Core is closing the "view", NOT this OLE Server---
            cwbCloseViewMyself = False
            
            '---close ALL forms that are associated with all "views"---
            Unload frmReports
        End If
    End If
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
'*******************************************************************************************
Public Sub MenuItemSelected(ByVal vlMenuId As Long)
    
    Static intCount As Integer
    On Error GoTo ErrHandler
    
    '----------------------------------------------------------------------------------
    '  NOTE: You don't have to create menu items, these are optional, that is why
    '        the code is commented out.
    '----------------------------------------------------------------------------------
'    '---perform the actions associated with the menu item that was selected---
'    Select Case vlMenuId
'        Case glOleMenuItem
'            '******************************************************************************
'            '  TO DO:
'            '    Take actions needed to process this menu item.
'            '******************************************************************************
        
        'frmBrowser.timTimer.Enabled = True
         
'        Select Case vlMenuId 'Button.Key
'            Case glOleMenuItem0
                'frmBrowser.brwWebBrowser.GoBack
'            End Select

    Exit Sub

ErrHandler:
    '---show the VB error---
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "Template.MenuItemSelected", "Template", vbExclamation)
    
    Exit Sub
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
    '---update the all running "views" with the new patient---
    If cwbMainViewRunning Then
        '****************************************************************************
        '  TO DO:
        '    Add code here to update the data in all "running" views with this new
        '    patient.
        '****************************************************************************
        
        
    
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
'*******************************************************************************************
Public Sub ResizeView(ByVal viViewType As Integer, ByVal vlLeft As Long, ByVal vlTop As Long, ByVal vlWidth As Long, ByVal vlHeight As Long)
    
    On Error GoTo ErrHandler
    Screen.MousePointer = vbHourglass
    
    '---reset global "client area" parameters---
    cwlLeft = vlLeft
    cwlTop = vlTop
    cwlWidth = vlWidth
    cwlHeight = vlHeight
    
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
            frmReports.WindowState = vbNormal
            
            '---hide the form to so that annyoing "screen blinking" will not be seen while controls are moved and resized---
            frmReports.Visible = False
            
            '---resize form to fill entire "client area", NOTE: This generates a "resize" event---
            frmReports.Move cwlLeft, cwlTop, cwlWidth, cwlHeight
            Rem WBW frmReports.Resize cwlLeft, cwlTop, cwlWidth, cwlHeight
                
            frmReports.Visible = True
        Else    '--first "regular" resize, need to perform "initailization" procedures--
            '---reset flag so this code will not be executed again---
            cwbFirstRegularResize = False
            
            frmReports.Move cwlLeft, cwlTop, cwlWidth, cwlHeight
            'frmReports.WindowState = 2
            '----------------------------------------------------------------------------
            '  NOTE: Here is where you "Show" the "Regular View" form for the FIRST time.
            '----------------------------------------------------------------------------
            frmReports.Show
            frmReports.Display
        End If
    End If
    
    GoTo Cleanup
    
ErrHandler:
    '---show the VB error---
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "Template.ResizeView", "Template", vbExclamation)
Cleanup:
    Screen.MousePointer = vbDefault
End Sub

Private Sub CreateDBAccessObjs()
On Error GoTo CreateDBAccessObjsError
    'Clr everything
    Call DestroyDBAccessObjs
    
    Set goDBConn = New CHCSII_CONN.Conn
    Set goDBSqlOps = goDBConn.CHCSII_SQLOPS(Auto)
    Set goDBDAS = goDBConn.CHCSII_DAS(Auto)
    If Not gobjshared.IsAppMode(modeCDR) Then
      Set goDBSqlOpsEx = goDBConn.CHCSII_SQLOPS_EX(Auto)
    End If
    
    GoTo ExitCreateDBAccessObjs
    
CreateDBAccessObjsError:
    gobjshared.ShowVBError Err.Number, Err.Description, "Forms.CreateDBAccessObjs", "CHCSII_Forms", vbExclamation
    'Clean up if there is an error
    Call DestroyDBAccessObjs
ExitCreateDBAccessObjs:
End Sub

Private Sub DestroyDBAccessObjs()
    'Clean house
    Set goDBSqlOps = Nothing
    Set goDBDAS = Nothing
    Set goDBConn = Nothing
    Set goDBSqlOpsEx = Nothing
End Sub

Private Sub Class_Initialize()

    On Error GoTo ErrHandler
    
    '--- SCR-ARM-16148 MSolano 9/10/01; COMMON TO MULTIPLE ARM-SCRs
    '--- Init routine should be called from first top level creatable coclass
    '--- It can be called multiple times, but will actually run only once.
    '--- Instrument all creatable classes in case dll's multiple entry
    '--- points are actually used
     Call ARM_InitAllTransactionsForCHCSIIReports
    
    '---create shared object used to show VB errors, etc.---
    Set gobjshared = New CWShared
    
    '---set the help file for this application---
    App.HelpFile = gobjshared.CWHelpFile
    
    Call CreateDBAccessObjs   'hwcc 12/11/2003
    
    '------------------------------------
    '  This statement is needed to debug
    '    this OLE Server.
    '------------------------------------
    
    GoTo Cleanup

ErrHandler:
    MsgBxARMd "Error@Reports.Class_Initialize Err#" & CStr(Err.Number) & " Error:" & Err.Description, vbCritical
Cleanup:
    #If ciDEBUG = 1 Then
        If cwbDebugLogon = False Then
            Exit Sub
        End If
    #End If
    Exit Sub
End Sub

Private Sub Class_Terminate()
    '---release reference to objects---
    If bDontTerminate Then Exit Sub
    Set gSearch = Nothing
    Set gobjshared = Nothing
    Set cwobjCWComm = Nothing
    Set cwobjLogonEnviron = Nothing
    Set cwobjConfig = Nothing
    Set cwobjPatient = Nothing
    Set cwobjUser = Nothing

    '********************************************************************
    '  TO DO:
    '    Add code here to release reference to all other global objects.
    '********************************************************************
    
    Set poMTF = Nothing
    Set poMTFSearch = Nothing
    Rem Public poClinic As CHCSII_ClinicClient.Clinic '-- Clinic Object
    Set poClinicColl = Nothing
    Set poClinicSearch = Nothing
    Set poDiagnosis = Nothing
    Set poDSearch = Nothing
    Set poProcedure = Nothing
    Set poPSearch = Nothing
    Set poLabTest = Nothing
    Set poLSearch = Nothing
    Set poMeds = Nothing
    Set poMedSearch = Nothing
    Set poProvItem = Nothing
    Set poProvSearch = Nothing
    Set poRadSearch = Nothing
    Set poRads = Nothing
    Set poPatSearch = Nothing
    Set poPatient = Nothing
    
    Call DestroyDBAccessObjs  'hwcc 12/11/2003
    
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
    
    '-------------------------------------------------------------------------------
    '  Take corresponding actions based upon the message "type".
    '
    '  NOTE: This method is required, but currently there are NO REQUIRED messages
    '        to process.
    '-------------------------------------------------------------------------------
    Select Case viMessageType
        Case cwiPRINT_REGULAR_VIEW
            On Error Resume Next
            'If frmReports.picView.  CRViewer1.ViewCount <> 0 Then
            'If frmReports.CRViewer1.ViewCount <> 0 Then
            '    frmReports.CRViewer1.PrintReport
            'End If
        Case cwiPRINT_PREVIEW
            On Error Resume Next
            Call frmReports.cmdRunReport_Click(frmReports.tabReports.CurrTab)
        Case cwiABOUT
            '---inform CW that I am showing a modal form---
            cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_GENERAL_REPORTS_APPLICATION
            Dim objAbout As New CWAbout
            Call objAbout.Display(App, frmReports)
            Set objAbout = Nothing
            '---inform CW that I am done showing a modal form---
            cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_GENERAL_REPORTS_APPLICATION
        Case cwiOK_TO_CHANGE_PATIENT
            cwobjCWComm.Message cwiOK_TO_CHANGE_PATIENT, "Y", vsSourceServerNCID, NCID_GENERAL_REPORTS_APPLICATION
        Case cwiOK_TO_CLOSE_VIEW
            cwobjCWComm.Message cwiOK_TO_CLOSE_VIEW, "Y", vsSourceServerNCID, NCID_GENERAL_REPORTS_APPLICATION
        Case cwiPROPERTIES
            frmProperties.Display
        Case Else   '--do nothing--
'            MsgBxARMd "This application does not support this action", vbInformation, "Template"
    End Select

    Exit Sub
    
ErrHandler:
    '---show the VB error---
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "CHCSII_Report.Message", "Reports", vbExclamation)
    
    Exit Sub
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
    Set cwobjCWComm = robjComm
    Set cwobjLogonEnviron = robjLogonEnvironment
    Set cwobjConfig = robjConfig
    Set cwobjPatient = robjPatient
    
    '---set global flags---
    cwbMainViewRunning = False
    
    '****************************************************************************
    '  TO DO:
    '    Add code here to create objects, initialize global variables, etc.
    '****************************************************************************
    
    Set gSearch = New CHCSII_Search.Search
    gSearch.Initialize cwobjCWComm, cwobjLogonEnviron, cwobjConfig, cwobjPatient, cwobjUser
    
    App.HelpFile = gobjshared.CWHelpFile '---Instantiate Help-file object

    Exit Sub

ErrHandler:
    '---show the VB error---
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "Reports.Initialize", "CHCSII_Reports", vbExclamation)
    
    Exit Sub
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
    
    '*****************************************************************
    '  TO DO:
    '    Add code here to invoke this OLE Server's properties page.
    '*****************************************************************
    


    Exit Sub
    
ErrHandler:
    '---show the VB error---
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "Template.Properties", "Template", vbExclamation)
    
    Exit Sub
End Sub

Public Sub ICHCSII_CoreAPI2_Initialize2(ByVal vlMyProgId As Long, robjComm As Object, robjLogonEnvironment As Object, robjConfig As Object, robjPatient As Object, robjUser As Object)
    
    Const Priv_Reports_Pop As String = "reports/Pop_Health"
    Const Priv_Reports_SRTS As String = "reports/SRTS"
    Const Priv_Reports_CPG As String = "reports/CPG"
    Const Priv_Reports_Dental As String = "reports/Dental"
    Const Priv_Reports_ESF As String = "reports/ESF"
    
    Set cwobjUser = robjUser
    Call Initialize(vlMyProgId, robjComm, robjLogonEnvironment, robjConfig, robjPatient)
    DoEvents
    
    gStandardSecurity = cwobjUser.HasPrivilegeEx(Priv_Reports_Standard, Priv_Read)
    gPreventitiveSecurity = cwobjUser.HasPrivilegeEx(Priv_Reports_Preventive, Priv_Read)
    gCustomizedSecurity = cwobjUser.HasPrivilegeEx(Priv_Reports_Customized, Priv_Read)
    gPopulationSecurity = cwobjUser.HasPrivilegeEx(Priv_Reports_Pop, Priv_Read)
    gSRTSSecurity = cwobjUser.HasPrivilegeEx(Priv_Reports_SRTS, Priv_Read)
    gCPGSecurity = cwobjUser.HasPrivilegeEx(Priv_Reports_CPG, Priv_Read)
    gDentalSecurity = cwobjUser.HasPrivilegeEx(Priv_Reports_Dental, Priv_Read)
    gESFSecurity = cwobjUser.HasPrivilegeEx(Priv_Reports_ESF, Priv_Read)
    
    gMTFSecurity = cwobjUser.HasPrivilegeEx(Priv_Reports_MTF, Priv_Read)
    gClinicSecurity = cwobjUser.HasPrivilegeEx(Priv_Reports_Clinic, Priv_Read)
    
    'Set poPatSearch.Logon = cwobjLogonEnviron
    'Set poPatSearch.User = cwobjUser

End Sub









