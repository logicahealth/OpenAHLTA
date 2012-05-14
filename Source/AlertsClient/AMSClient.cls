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
Attribute VB_Name = "AlertsClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Configuration"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements ICHCSII_CoreAPI2
Implements MessageTransceiver
'local variable(s) to hold property value(s)
Private mvarConfiguration As Object 'local copy
'local variable(s) to hold property value(s)
Private mvarLogonEnviron As Object 'local copy
'--- Menu actions
Private Const MENU_ENABLE As String = "ENABLE"
Private Const MENU_DISABLE As String = "DISABLE"
Private Const MENU_CHECK As String = "CHECK"
'local variable(s) to hold property value(s)
Private mvarEncounterCosignsNeeded As Boolean 'local copy
'local variable(s) to hold property value(s)
Private mvarUnresolvedAlerts As Boolean 'local copy

Private mvarAlertsToolTip As String

'**************************************
'SCR 14692
'Local Variable that indicates whether or not
'  there are any unresolved "Orders-To-Sign" Alerts
Private mvarOrdersToSignNeeded As Boolean
'**************************************

Public Property Let UnresolvedAlerts(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UnresolvedAlerts = 5
    mvarUnresolvedAlerts = vData
End Property


Public Property Get UnresolvedAlerts() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UnresolvedAlerts
    UnresolvedAlerts = mvarUnresolvedAlerts
End Property



Public Property Let EncounterCosignsNeeded(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EncounterCosignsNeeded = 5
    mvarEncounterCosignsNeeded = vData
End Property


Public Property Get EncounterCosignsNeeded() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EncounterCosignsNeeded
    EncounterCosignsNeeded = mvarEncounterCosignsNeeded
End Property

'*******************************************
'SCR 14692
Public Property Let OrdersToSignNeeded(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
    mvarOrdersToSignNeeded = vData
End Property
Public Property Get OrdersToSignNeeded() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
    OrdersToSignNeeded = mvarOrdersToSignNeeded
End Property
'*******************************************

'***********************************************************
'29936
Public Property Let AlertsToolTip(ByVal vData As String)
    mvarAlertsToolTip = vData
End Property
Public Property Get AlertsToolTip() As String
    AlertsToolTip = mvarAlertsToolTip
End Property
'***********************************************************

Public Property Set LogonEnviron(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.LogonEnviron = Form1
    Set mvarLogonEnviron = vData
    Set cwobjLogonEnviron = vData
End Property


Public Property Get LogonEnviron() As Object
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LogonEnviron
    Set mvarLogonEnviron = cwobjLogonEnviron
    Set LogonEnviron = mvarLogonEnviron
End Property

Public Property Set Configuration(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Configuration = Form1
    Set mvarConfiguration = vData
    Set gConfiguration = vData
End Property


Public Property Get Configuration() As Object
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Configuration
    If Not gConfiguration Is Nothing Then
        Set mvarConfiguration = gConfiguration
    End If
    Set Configuration = mvarConfiguration
End Property

Public Function EnableDetailsMenu() As Boolean
    EnableDetailsMenu = cwobjCWComm.UpdateMenuItem(glDetailsMenuId, MENU_ENABLE)
End Function

Public Function DisableDetailsMenu() As Boolean
  'Here it is!
    DisableDetailsMenu = cwobjCWComm.UpdateMenuItem(glDetailsMenuId, MENU_DISABLE)
End Function

Public Function EnableResolveMenu() As Boolean
    EnableResolveMenu = cwobjCWComm.UpdateMenuItem(glResolveMenuId, MENU_ENABLE)
End Function

Public Function DisableResolveMenu() As Boolean
    DisableResolveMenu = cwobjCWComm.UpdateMenuItem(glResolveMenuId, MENU_DISABLE)
End Function
Public Function EnableRemoveMenu() As Boolean
    EnableRemoveMenu = cwobjCWComm.UpdateMenuItem(glRemoveMenuId, MENU_ENABLE)
End Function
Public Function DisableRemoveMenu() As Boolean
    DisableRemoveMenu = cwobjCWComm.UpdateMenuItem(glRemoveMenuId, MENU_DISABLE)
End Function
Public Function EnableRefreshMenu() As Boolean
    EnableRefreshMenu = cwobjCWComm.UpdateMenuItem(glRefreshMenuId, MENU_ENABLE)
End Function
Public Function DisableRefreshMenu() As Boolean
    DisableRefreshMenu = cwobjCWComm.UpdateMenuItem(glRefreshMenuId, MENU_DISABLE)
End Function





'Private cwobjSecurity As Object

'*******************************************************************************-
'-
'************
'  Function Name: QueryView
'
'  Description: This method is called when a client wants to know if this OLE
'Server
'               supports a specific "view type".
'
'  Parameters:
'    viViewType - Contains the "type" of view being queried for.
'
'  Return Value:
'    Returns TRUE if this OLE Server supports the passed in "view type".
'    Returns FALSE if this OLE Server does not support the passed in "view
'type".
'******************************************************************************-
'-
'*************
Public Function QueryView(ByVal viViewType As Integer) As Boolean
    Select Case viViewType
      Case cwiREGULAR_VIEW
        QueryView = True    '--this view is NOT optional--
      Case cwiPATIENT_SUMMARY_VIEW
        QueryView = False
      Case cwiENCOUNTER_SUMMARY_VIEW
        QueryView = False
      Case cwiPROPERTY_VIEW
        '*************************
        '10-Jan-2002
        'SCR 19958
        'Donot allow the user to enter the Options Screen.
        'QueryView = True
        QueryView = False
        '*************************
      Case Else   '--unknown "view type", not supported--
        QueryView = False
    End Select
    
End Function

'*******************************************************************************-
'-
'************
'  Sub Name: MenuItemSelected
'
'  Description: This method is called when the user has selected a menu item
'(from the CW
'               Core's main menu), which is "owned" by this OLE Server.
'
'               NOTE: This is an OPTIONAL method.
'
'  Parameters:
'    vlMenuId - Contains the menu id of the menu item the user selected.
'******************************************************************************-
'-
'*************
Public Sub MenuItemSelected(ByVal vlMenuId As Long)
    On Error GoTo ErrHandler
    
    Dim i As Integer
    Dim sMessageGuid As String
    Dim oAlert As Object
    
    '----------------------------------------------------------------------------------
    '  NOTE: You don't have to create menu items, these are optional, that is why
    '        the code is commented out.
    '----------------------------------------------------------------------------------


    Select Case vlMenuId
        
        Case glResolveMenuId
          
          If FrmAlerts.AlertsGrid.Rows >= 2 Then
          
            For i = 1 To FrmAlerts.AlertsGrid.Rows - 1
            
              If FrmAlerts.AlertsGrid.Cell(flexcpData, i, alertCOL_MESSAGE_GUID) = gsTextMessageGuid Then
               
                sMessageGuid = gsTextMessageGuid
                Call FrmAlerts.AddressAlert(sMessageGuid)
                Exit For
              
              End If
              
            Next i
            
          End If
              
        Case glRemoveMenuId
        
          Screen.MousePointer = vbHourglass
        
          If FrmAlerts.AlertsGrid.Rows >= 2 Then
          
          For i = 1 To FrmAlerts.AlertsGrid.Rows - 1
          
            If FrmAlerts.AlertsGrid.Cell(flexcpData, i, alertCOL_MESSAGE_GUID) = gsTextMessageGuid Then
             
              sMessageGuid = gsTextMessageGuid
              
              Set oAlert = gMessages.Item(sMessageGuid)
              If DeleteAlertGlobal(oAlert) Then
                Exit For
              End If
              
'              Set oAlert = gMessages.Item(sMessageGuid)
'
'              If cwobjLogonEnviron.UserNCID = oAlert.ProviderNcid Then
'                  oAlert.Resolved = True
'                  If (oAlert.MessageTypeGuid = GUID_NEW_RESULT Or oAlert.MessageTypeGuid = GUID_PRIORITY_RESULT) Then
'                      oAlert.Chcs_Clear = True
'                      If Not ResolveAlert(oAlert.ProviderNcid, oAlert) Then
'                          Err.Description = "Error in ResolveAlert ProviderNCID = " & oAlert.ProviderNcid & "   MessageGUID = " & sMessageGuid
'                          GoTo ErrHandler
'                      End If
'                  Else
'                      If Not DeleteAlert(oAlert.ProviderNcid, sMessageGuid) Then
'                          Err.Description = "Error in DeleteAlert ProviderNCID = " & oAlert.ProviderNcid & "   MessageGUID = " & sMessageGuid
'                          GoTo ErrHandler
'                      End If
'                  End If
'                  gMessages.Remove (sMessageGuid)
'              Else
'                  MsgBox "This message belongs to " & GetByNCID(oAlert.ProviderNcid).Name & ". You may not delete it.", vbInformation, "ALERTS EXCEPTION MESSAGE"
'              End If
            
            End If
          
          Next i
                   
           FrmAlerts.InitialData

         End If
         
         Screen.MousePointer = vbDefault
        
        Case glDetailsMenuId
        
          If FrmAlerts.AlertsGrid.Rows >= 2 Then
            
            For i = 1 To FrmAlerts.AlertsGrid.Rows - 1
              
              If FrmAlerts.AlertsGrid.Cell(flexcpData, i, alertCOL_MESSAGE_GUID) = gsTextMessageGuid Then
                sMessageGuid = gsTextMessageGuid
                Set oAlert = gMessages(sMessageGuid)
                Details oAlert
                'SCR #5584
                Call DlgDetails.Display
                Exit For
              End If
              
            Next i
            
          End If
        
        Case glRefreshMenuId
        
          '******************************
          'SCR 21410
          Screen.MousePointer = vbHourglass
          FrmAlerts.InitialData
          DoEvents
          Screen.MousePointer = vbDefault
          DoEvents
          '******************************
          
        'Case giConfigMenuId
        ' Go to configuration if I am administrator
        '    DlgConfiguration.Show vbModeless
        'Case giCloseMenuId
        ' Close this view
        '    CloseView cwiREGULAR_VIEW
        '    gobjCoreProxy.CloseView cwiREGULAR_VIEW, cwlMyProgId
    End Select

    Screen.MousePointer = vbDefault
    Exit Sub

ErrHandler:
    '---show the VB error---
    Screen.MousePointer = vbDefault
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "AlertsClient.MenuItemSelected", "AMSClient", vbExclamation)
    
    Exit Sub
End Sub

'*******************************************************************************-
'-
'*******************************
'  Sub Name: PatientSelected
'
'  Description: This method gets called when a new patient has been selected
'from the CW Core.
'
'               NOTE: This is an OPTIONAL method.
'
'  Parameters:
'    none
'******************************************************************************-
'-
'********************************
Public Sub PatientSelected()
    '---update the all running "views" with the new patient---
    If cwbMainViewRunning Then
      Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub Class_Initialize()

#If coiDebug Then
     gCOI_DL.debugStart
     DR.TimerTime = 0
     gbDebugStarted = True
'     gCOI_DL.DebugLog "COI31", "ALERT", "", "Begin", DR
#End If

    '---create shared object used to show VB errors, etc.---
    Set gobjShared = New CWShared
    'Set cwobjNavigation = New Navigation
    

    '---set the help file for this application---
    App.HelpFile = gobjShared.CWHelpFile
    
    '*********************************************
    '10-Jan-2002
    'SCR 19958
    'This will Delete ALL records from the Surrogates Table.
    If DeleteSurrogates = True Then
    End If
    '*********************************************
        
    '------------------------------------
    '  This statement is needed to debug
    '    this OLE Server.
    '------------------------------------
    #If ciDEBUG = 1 Then
        If cwbDebugLogon = False Then
            Exit Sub
        End If
    #End If
End Sub

'*******************************************************************************-
'-
'************
'  Sub Name: Properties
'
'  Description: This method allows a user the ability to modify the "properties" of
'this OLE
'               Server (i.e. configuration values).
'
'  Parameters:
'    robjLogonEnviron - Contains the logon environment object.
'    robjConfig - Contains the configuration object.
'    vsOwnerNCID - Contains the "owner" for which to store the new/updated OLE
'Server
'                  property information.
'    vbCalledByDesigner - TRUE if this method is being called by the CW
'Designer Tool,
'                         FALSE otherwise.
'******************************************************************************-
'-
'*************
Public Sub Properties(robjLogonEnviron As Object, robjConfig As Object, ByVal vsOwnerNCID As String, ByVal vbCalledByDesigner As Boolean)
    Dim sBuf As String
    
    On Error GoTo ErrHandler
    
    '--------------------------------------------------------------------------------------
    '  If this method is being called by the CW Designer, you need to perform actions that
    '    the OLE Server would preform in the "Initialize" method and other methods (i.e.
    '    create objects, initialize global variables, etc.).
    '--------------------------------------------------------------------------------------
    If vbCalledByDesigner Then
        '---initialize variables---
        ReDim gatFacilities(0)
        
        '---create pointers to objects---
        Set cwobjLogonEnviron = robjLogonEnviron
        Set cwobjConfig = robjConfig
        
        '---create OLE object---
        'Set gobjConcept = New ConceptCtrl    5/14/03  Phil Crowder  not used
        
        If vbCalledByDesigner Then
           '---clear out values---
           ReDim gatFacilities(0)
           
           '---release reference to objects---
           Set cwobjLogonEnviron = Nothing
           Set cwobjConfig = Nothing
           'Set gobjConcept = Nothing
        End If
    End If

    Exit Sub
    
ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "CHCSII_AlertsList.Properties", "CHCSII_AlertsList", _
                                vbExclamation)
    
    '---if you were called by the CW Designer, clean up the objects that were used---
    If vbCalledByDesigner Then
        '---clear out values---
        
        '---release reference to objects---
        Set cwobjLogonEnviron = Nothing
        Set cwobjConfig = Nothing
        'Set gobjConcept = Nothing
    End If
    
    Exit Sub
End Sub

'*******************************************************************************-
'-
'************
'  Sub Name: ResizeView
'
'  Description: This method is called when a change has occurred in the
'"client area"
'               used by this OLE server and it needs to resize a particular
'"view".
'
'  Parameters:
'    viViewType - Contains the "type" of view being requested to resize by CW.
'    vlLeft - Contains the "left" coordinate for the OLE server.
'    vlTop - Contains the "top" coordinate for the OLE server.
'    vlWidth - Contains the "width" for the OLE server.
'    vlHeight - Contains the "height" for the OLE server.
'******************************************************************************-
'-
'*************
Public Sub ResizeView(ByVal viViewType As Integer, ByVal vlLeft As Long, ByVal vlTop As Long, ByVal vlWidth As Long, ByVal vlHeight As Long)
    On Error GoTo ErrHandler

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
            '  NOTE: This generates a "resize" event that is ignored.
            '---------------------------------------------------------------------------------
            FrmAlerts.WindowState = vbNormal
            
            '---hide the form to so that annyoing "screen blinking" will not be seen while controls are moved and resized---
            FrmAlerts.Visible = False
            
            '---resize form to fill entire "client area", NOTE: This generates a "resize" event---
            FrmAlerts.Move cwlLeft, cwlTop, cwlWidth, cwlHeight
                
            FrmAlerts.Visible = True
        Else    '--first "regular" resize, need to perform "initailization" procedures--
            '---reset flag so this code will not be executed again---
            cwbFirstRegularResize = False
        
            '---display the "main" form---
            FrmAlerts.Display
        End If

        #If coiDebug Then
            If gbIsCOI31 = True Then
                gCOI_DL.DebugLog "COI31", "ALERT", "", "End", DR
            End If
            gbDebugStarted = True
            gbIsCOI31 = False
        #End If

    Else
        Screen.MousePointer = vbDefault
        Beep
        MsgBox "Unsupported resize view type: " & CStr(viViewType) & " in CHCSII_AlertsList.ResizeView.", _
               vbExclamation, "CHCSII_AlertsList"
    End If
    
    Exit Sub
    
ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "CHCSII_AlertsList.ResizeView", "CHCSII_AlertsList", _
                                vbExclamation)
    
    Exit Sub
End Sub

'*******************************************************************************-
'-
'************
'  Function Name: OpenView
'
'  Description: This method is called when CW attempts to open a "view"
'provided by this
'               OLE server.
'
'  Parameters:
'    viViewType - Contains the "type" of view being requested by CW to open.
'    vlParenthWnd - Contains the hWnd of the client area window to be used by
'this OLE
'                   server.
'    rhMainForm - If "viViewType" is a "regular view", upon return this will
'contain
'                 the window "handle" of the main "regular" form for this OLE
'Server.
'
'  Return Value:
'    Returns TRUE if openning the "view" was successfully.
'    Returns FALSE if an error occurred openning the "view".
'******************************************************************************-
'-
'*************
'##ModelId=37A5FB9301FE
Public Function OpenView(ByVal viViewType As Integer, ByVal vlParenthWnd As Long, rhMainForm As Long) As Boolean
    On Error GoTo ErrHandler
    
    '---perform actions based on the "view type"---
    If viViewType = cwiREGULAR_VIEW Then
    

#If coiDebug Then
    If gbDebugStarted Then
        DR.TimerTime = 0
        gCOI_DL.DebugLog "COI31", "ALERT", "", "Begin", DR
    End If
    gbIsCOI31 = True
#End If
        

        '---set global flags---
        cwbFirstRegularResize = True
        cwbMainViewRunning = True
    
        glDetailsMenuId = gobjCoreProxy.AddActionMenuItem("Details", "details.bmp", cwlMyProgId, "Show Details")
        glResolveMenuId = gobjCoreProxy.AddActionMenuItem("Address Alert...", "resolve.bmp", cwlMyProgId, "Address Alert...")
        glRemoveMenuId = gobjCoreProxy.AddActionMenuItem("Delete", "delete.bmp", cwlMyProgId, "Delete Alert from List")
        glRefreshMenuId = gobjCoreProxy.AddActionMenuItem("Refresh", "refresh.bmp", cwlMyProgId, "Refresh Alerts")
'        giSeperator1MenuId = gobjCoreProxy.AddActionMenuItem("", "", cwlMyProgId, "")
'        giConfigMenuId = gobjCoreProxy.AddActionMenuItem("Configuration...", "configure.bmp", cwlMyProgId, "Configuration...")
'        giSeperator2MenuId = gobjCoreProxy.AddActionMenuItem("", "", cwlMyProgId, "")
'        giCloseMenuId = gobjCoreProxy.AddActionMenuItem("Close Alerts", "close.bmp", cwlMyProgId, "Close Alerts")
        
        '----------------------------------------------------------------------
        '  Reset the "parent window" for this OLE server's "main" regular form.
        '
        '  NOTE: This invokes the form's "Load" event, but the form is NOT
        '        shown.  Make sure that you DO NOT perform a "Me.Show" in the
        '        form's "Load" event.
        '----------------------------------------------------------------------
'< fix for failure to terminate:
        Set FrmAlerts.objParent = Me
        SetParent FrmAlerts.hwnd, vlParenthWnd
        
        '---set parameter---
        rhMainForm = FrmAlerts.hwnd
        glfrmCWHandle = vlParenthWnd
        '----------------------------------------------------------------------
        '  Reset the "parent window" for this OLE server's "main" summary form.
        '
        '  NOTE: This invokes the form's "Load" event, but the form is NOT
        '        shown.  Make sure that you DO NOT perform a "Me.Show" in the
        '        form's "Load" event.
        '----------------------------------------------------------------------
    Else
        Screen.MousePointer = vbDefault
        Beep
        MsgBox "Unsupported view type: " & CStr(viViewType) & " in CHCSII_AlertsList.OpenView.", _
               vbCritical, "CHCSII_AlertsList"
        
        '---set return value---
        OpenView = False
        
        Exit Function
    End If
        
    '---set return value---
    OpenView = True
    
    Exit Function
    
ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "CHCSII_AlertsList.OpenView", "CHCSII_AlertsList", _
                                vbCritical)
    
    '---set return value---
    OpenView = False
    
    Exit Function
End Function


'*******************************************************************************-
'-
'******************
'  Sub Name: Message
'
'  Description: This method is used by CW to send miscellaneous messages to
'this OLE Server.
'
'  Parameters:
'    viMessageType - Contains the "type" of message to process.
'    vsMessageText - Contains any "data" assoicated with this message.
'    vsSourceServerNCID - Contains the NCID of the OLE Server that "sent" the
'message.
'******************************************************************************-
'-
'*******************
'##ModelId=37A5FB930202
Public Sub Message(ByVal viMessageType As Integer, ByVal vsMessageText As String, ByVal vsSourceServerNCID As String)
    'Dim objAbout As CWAbout
    Dim lRetVal As Long
    Dim lUnitNumber As Long
    Dim sMedRecNumber As Long
    
    On Error GoTo ErrHandler
    
    Select Case viMessageType
    
        Case cwiOK_TO_CLOSE_VIEW
        
'          bAlertsFormIsOpen = False
                        
        Case cwiFOREGROUND_APPLICATION
        
'          If bAlertsFormIsOpen = False And cwbMainViewRunning = True Then
'            FrmAlerts.InitialData
'          End If
'          bAlertsFormIsOpen = True
          

        Case cwiSELECT_PATIENT
            '---need to parse out the "Unit#" and "MedRec#" from the message---
            
            lRetVal = InStr(1, vsMessageText, ";")  '--look for the Unit# and MedRec# separator--
            
            If lRetVal <> 0 Then
                lUnitNumber = CLng(Mid$(vsMessageText, 1, lRetVal - 1))
                sMedRecNumber = Mid$(vsMessageText, lRetVal + 1, Len(vsMessageText) - lRetVal)
                
                '---perform all actions necessary when selecting a patient---
'                frmCW.SelectPatient lUnitNumber, sMedRecNumber, True
            Else
                Beep
                MsgBox "Invalid message format for selecting a patient: " & vsMessageText & ". It needs to be: 'UnitNumber;MedRecNumber'.", vbExclamation, App.Title
                
                Exit Sub
            End If
    
        Case cwiREFRESH
            '---"refresh" or repaint all running "views"---
            If cwbMainViewRunning Then
                '---NOTE: Changing the form's "visible" property ensures ALL child controls get repainted---
                FrmAlerts.Visible = False
                FrmAlerts.Visible = True
            End If
            
        Case cwiPRINT_REGULAR_VIEW
            '---print the clinic CHCSII_AlertsList---
            'FrmAlerts.PrintAppointments
        Case cwiPRINT_PREVIEW
            '---set flag---
            'cwbPrintPreview = True
            
            '---print the clinic CHCSII_AlertsList---
            'FrmAlerts.PrintAppointments
            
            '---set flag---
            'cwbPrintPreview = False
        Case cwiITEM_SPECIFIC_HELP
            '---invoke the "what's this mode" for the CHCSII_AlertsList form---
            FrmAlerts.WhatsThisMode
        Case cwiPROPERTIES
            '---allow the user to change the "properties" of this application--
            'Call Properties(cwobjLogonEnviron, cwobjConfig, cwobjLogonEnviron.UserNCID, False)
            
            '*********************************************
            '10-Jan-2002
            'SCR 19958
            'Don't Allow User to Configure anything.
            'This will NOT happen because of QueryView.
            DlgConfiguration.Show vbModeless
            '*********************************************
        Case cwiABOUT
            Dim objAbout As New CWAbout
            
            '---inform CW that I am showing a modal form---
            cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ALERT_APPLICATION
            
            objAbout.Display App, FrmAlerts
            
            '---inform CW that I am done showing a modal form---
            cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ALERT_APPLICATION
            
            Set objAbout = Nothing
            
            FrmAlerts.Visible = False
            FrmAlerts.Visible = True
            
        Case cwiREFRESH_DATA
            If cwbMainViewRunning Then
            
                '---refresh the current clinic CHCSII_AlertsList---
                FrmAlerts.InitialData
            End If
            
        Case cwiGENERIC
        
           Select Case UCase(Trim(vsMessageText))
           Case "KILL_KEA|FALSE"
                gsKEAStatus = "FALSE"
           
           Case "KILL_KEA|TRUE"
                gsKEAStatus = "TRUE"
           
           Case "EXIT_APP"
                Set mvarLogonEnviron = Nothing
                Set cwobjLogonEnviron = Nothing
                Set mvarConfiguration = Nothing
                Set gConfiguration = Nothing
                Set moClinician = Nothing
                Set moItem = Nothing
                Set moLogon = Nothing
                Set gobjShared = Nothing
                Set gobjCoreProxy = Nothing
                Set gobjUser = Nothing
                Set gMessages = Nothing
                Set gConfiguration = Nothing
                Set gColumnNames = Nothing
                Set gColumnShow = Nothing
                Set moConn = Nothing '<SCR 36493
                'Set gobjConcept = Nothing
           End Select

        Case Else
            'Beep
            'MsgBox "CHCSII_AlertsList does not support this action.", vbInformation, "CHCSII_AlertsList"
    End Select

    Exit Sub
    
ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "CHCSII_AlertsList.Message", "CHCSII_AlertsList", _
                                vbExclamation)
    
    Exit Sub
End Sub

'*******************************************************************************-
'-
'************
'  Sub Name: Initialize
'
'  Description: This method is called when CW "creates" an instance of this
'OLE Server.  It
'               is used to create and initialize global variables that are
'shared across all
'               "views" supported by this OLE Server, i.e. variables that only
'need to be
'               initialized ONCE, upon object creation.
'
'  Parameters:
'    vlMyProgId - Contains the "ProgId" that CW has assigned to this OLE
'Server.
'    robjComm - Contains a communications object that is used to handle all
'communication
'               between this OLE Server and the CW core.
'    robjLogonEnviron - Contains the logon environment object.
'    robjConfig - Contains the configuration object used for getting and
'setting configuration
'                 information for a specific application.
'    robjPatient - Contains the patient object of the currently selected
'patient.
'******************************************************************************-
'-
'*************
'##ModelId=37A5FB930230
Public Sub Initialize(ByVal vlMyProgId As Long, robjComm As Object, robjLogonEnvironment As Object, robjConfig As Object, robjPatient As Object)
    On Error GoTo ErrHandler
    
    Dim bSuccess As Boolean
       
    '---save the "ProgId" assigned to me by CW---
    cwlMyProgId = vlMyProgId
    
    '---create pointers to objects---
    Set cwobjCWComm = robjComm
    Set gobjCoreProxy = cwobjCWComm
    Set cwobjLogonEnviron = robjLogonEnvironment
    
    Set cwobjConfig = robjConfig
    Set cwobjPatient = robjPatient
    Set gMessages = New Collection
    '---set global flags---
    cwbMainViewRunning = False
    
    '37879
    gbDoAlertsReceive = True
    
        '*********************************************
    'SCR 21082
    alertCOL_ID = 0
    alertCOL_MESSAGE_GUID = 1
    alertCOL_TYPE = 2
    alertCOL_MESSAGE = 3
    alertCOL_DISPLAYMESSAGE = 4
    alertCOL_TIME = 5
    alertCOL_TIME_SORT = 6
    alertCOL_RESOLVED = 7
    alertCOL_PRIORITY = 8
    alertCOL_PROVIDER = 9
    alertCOL_SURROGATE = 10
    alertCOL_IEN = 11
    alertCOL_PATIENT = 12
    alertCOL_NOTE = 13
    alertCOL_DOCUMENT = 14
    '*********************************************
    
    Exit Sub

ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "CHCSII_AlertsList.Initialize", "CHCSII_AlertsList", _
                                vbExclamation)
    
    Exit Sub
End Sub

'*******************************************************************************-
'-
'************
'  Sub Name: CloseView
'
'  Description: This method is called when CW closes a "view" provided by this
'OLE server.
'
'  Parameters:
'    viViewType - Contains the "type" of view that CW is closing.
'******************************************************************************-
'-
'*************
'##ModelId=37A5FB930236
Public Sub CloseView(ByVal viViewType As Integer)


    If Not (gobjCoreProxy Is Nothing) Then
      gobjCoreProxy.DeleteMenuItem glResolveMenuId
      gobjCoreProxy.DeleteMenuItem glRemoveMenuId
      gobjCoreProxy.DeleteMenuItem glDetailsMenuId
      gobjCoreProxy.DeleteMenuItem glRefreshMenuId
    End If
    
    If viViewType = cwiALL_VIEWS Then
        '---set global flags---
        cwbMainViewRunning = False
        
        '---this indicates that CW is closing the "view", NOT the OLE Server---
        cwbCloseViewMyself = False
        
        '---unload all forms---
        Unload FrmAlerts

    ElseIf viViewType = cwiREGULAR_VIEW Then
        '---set global flag---
        cwbMainViewRunning = False
        
        '---this indicates that CW is closing the "view", NOT the OLE Server---
        cwbCloseViewMyself = False
        
        Unload FrmAlerts
    End If
        
            
End Sub

Private Sub Class_Terminate()
'    Set gobjAlerts = Nothing
'    If Not gMessages Is Nothing Then
'        Set gMessages = Nothing
'    End If

'Debug.Print "Terminating alerts client " & mvarInitTime & " at "; Now

#If coiDebug Then
     gCOI_DL.debugStop
     Set gCOI_DL = Nothing
#End If

End Sub

Public Sub ICHCSII_CoreAPI2_Initialize2(ByVal vlMyProgId As Long, robjComm As Object, robjLogonEnvironment As Object, robjConfig As Object, robjPatient As Object, robjUser As Object)
    
    On Error GoTo ErrHandler
             
    Call Initialize(vlMyProgId, robjComm, robjLogonEnvironment, robjConfig, robjPatient)
    
    Set gobjUser = robjUser
'    Set gobjConcept = New ConceptCtrl   5/14/03  Phil Crowder  not used
'< dropped for fix of failure to terminate:    Set gobjAlerts = Me

      Configure gConfiguration

Retry:
    
    Exit Sub
    
ErrHandler:
        
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "AlertsClient.CoreAPI2_Initialize2", "CHCSII_AlertsList", _
                                vbExclamation)
    Exit Sub
End Sub

Public Function Navigate(aAlert As Object) As Boolean

On Error GoTo NavigateError

    Dim sMessageTypeGUID As String
    Dim oAlert As Object
    Dim oType As Object
    Dim sInteraction As String
    Dim iResult As Integer
    
Navigate = False

    Screen.MousePointer = vbHourglass
    
    Set oAlert = aAlert
       
    sMessageTypeGUID = oAlert.MessageTypeGuid
    Set oType = gConfiguration.AMSTypes.Item(sMessageTypeGUID)
    sInteraction = oType.Interaction
        
'  If sMessageTypeGuid = GUID_NEW_RESULT Or _
'      sMessageTypeGuid = GUID_PRIORITY_RESULT Then
'
'    gsKEAStatus = "NOTHING"
'    cwobjCWComm.Message cwiGENERIC, "KILL_KEA", NCID_CW_APPLICATION, NCID_ALERT_APPLICATION
'
'    If gsKEAStatus = "TRUE" Then
'      If KEA_RunLegacySession(sMessageTypeGuid) = True Then
'        Navigate = True
'      End If
'    End If
'
'  Else


    Navigate = False
    
  '***************************************
  'SCR 35188
  'GUID_ENCOUNTER_WRITEBACK_INCOMPLETE
  If sMessageTypeGUID = GUID_ENCOUNTER_WRITEBACK_INCOMPLETE Then
    MsgBox "Please go to the previous encounters for this patient and update the encounter in question.", vbApplicationModal, "Alerts Navigation"
    Navigate = True
  Else
  '***************************************
 
    If sInteraction = Empty Then
        MsgBox "This type of alert does not have a related OLE server", vbApplicationModal, "Alerts Navigation"
    ElseIf sInteraction = "KEA" Then
        MsgBox "Go to tools menu to bring up legacy system.", vbApplicationModal, "Alerts Navigation"
        Navigate = True
    Else

      If aAlert.PatientNcid <> "0" Then
        gobjCoreProxy.Message cwiSELECT_PATIENT, aAlert.PatientNcid & ";", NCID_CW_APPLICATION, NCID_ALERT_APPLICATION
      End If
      
      
      '***********************************************
      'SCR 14668
      'Make Sure the OLE Server is Running.
      gobjCoreProxy.Message cwiSTART_OLE_SERVER, sInteraction, NCID_CW_APPLICATION, NCID_ALERT_APPLICATION
      If cwobjCWComm.Message(cwiOLE_SERVER_RUNNING, sInteraction, NCID_CW_APPLICATION, NCID_ALERT_APPLICATION) Then
        Navigate = True
      Else
        Navigate = False
        MsgBox "The neccssary CHCSII OLE Server is not available." & vbCrLf & "The Alert will not be resolved.", vbInformation, "CHCSII Alerts Client"
      End If
      '***********************************************
      
      '***********************************************************
      'SCR 15406
      'When Navigating to Previous Encounters, a specific Encounter should be selected.
      If sMessageTypeGUID = GUID_ENCOUNTER_DOC_MODIFIED And Navigate = True Then
        Dim objPreviousEncounter As Object
        Set objPreviousEncounter = cwobjCWComm.InitializeOLEServer(NCID_PREVIOUS_ENCOUNTERS_APPLICATION)
        Call objPreviousEncounter.SetEncounterKey(aAlert.FacilityNcid & "|" & aAlert.DocumentNcid)
        Set objPreviousEncounter = Nothing
      End If
      '***********************************************************

    End If
  
  End If
    
  Set oAlert = Nothing
        
  Screen.MousePointer = vbDefault
    
  Exit Function
    
NavigateError:
    
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "AlertsClient.Navigate", "AMSClient", _
                                vbExclamation)
    
    Screen.MousePointer = vbDefault
    Navigate = False
    Exit Function

End Function

Public Sub Details(aAlert As Object)
    
    Dim oAlert As Object
    
    Set oAlert = aAlert
    DlgDetails.Message = oAlert
    
End Sub

Public Function CheckForSurrogate(aAlert As Object) As Integer
    
    Dim iResult As Integer
    Dim oAlert As Object
    
    Set oAlert = aAlert
    
    iResult = HasSurrogate(oAlert)
    
    If iResult = True Then
        If cwobjLogonEnviron.UserNCID = oAlert.ProviderNcid Then
            If (MsgBox("This alert has been assigned a surrogate." & vbNewLine & " Would you like to reclaim this alert?", vbYesNo, "Surrogate Binding Alert") = vbYes) Then
                If Not Reclaim(oAlert) Then
                    MsgBox "Could not reclaim alert. See Administrator", vbApplicationModal, "Surrogate Binding Alert"
                End If
            Else
                CheckForSurrogate = True ' At this point we can kill navigation
                Exit Function
            End If
        End If
    ElseIf iResult = False Then
        If cwobjLogonEnviron.UserNCID = oAlert.SurrogateNcid Then
            ' warn user
            MsgBox "This alert has been reclaimed by the primary provider.", vbApplicationModal, "Surrogate Binding Alert"
            ' remove this alert
            gMessages.Remove (oAlert.MessageTypeGuid)
            FrmAlerts.AlertsGrid.RemoveItem (FrmAlerts.AlertsGrid.RowSel)
        End If
    Else
        MsgBox "An error occurred while checking for surrogates", vbExclamation, "Internal Alerts Error"
    End If

    CheckForSurrogate = iResult
End Function


'*************************************************************************
'       O L D   M E D I A T O R   F U N C T I O N S
'*************************************************************************


'***********************************
'   Arguments for Send
'***********************************
'============================================================
' MODIFICATION HISTORY
' SCR-9794 R. Plouffe  1/22/01 Patient Unit was added to Alert Message
'    from Coding Review so Alerts knows which patient the Coding Change
'    is associated with.  Send now updates the database with the patient ncid
'    if the aMessageTypeGuid is a CODING_CHANGE.
'
' For Encounter_Doc_Modified, aDocumentGuid is the Encounter Number

Public Function Send(ByVal aProviderNcid As String, _
                    ByVal aMessageTypeGuid As String, _
                    ByVal aMessage As String, _
                    Optional ByVal aMessageIen As Long, _
                    Optional ByVal aPriority As Long, _
                    Optional ByVal aPatientNcid As String = "", _
                    Optional ByVal aSurrogateNcid As String = "", _
                    Optional ByVal aDocumentGuid As String = "") As Boolean
                    
    
    '*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#
    '(PGUI Layer 2) SCR 44507
    'Bill Fote 18-Nov-2003
    If gobjShared.IsAppMode(modeCHCSI_GUI) Then
      Send = True
      Exit Function
    End If
    '*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#
                    
    Dim moSQL As ADODB.Recordset
    Dim Connection As CHCSII_Conn.Conn
    Dim moGemsDAS As ICHCSII_DAS
    Dim oSurrogate As Surrogate
    Dim sMessageGuid As String
    Dim strDateTime As String
    Dim sSql As String
    Dim oAlertType As AMSType
    Dim sLogonFacilityNCID As String
    On Error GoTo SendError
    
'    Debug.Print "Send Start"
'    Debug.Print "aProviderNcid : " & aProviderNcid
'    Debug.Print "aMessageTypeGuid : " & aMessageTypeGuid
'    Debug.Print "aMessage : " & aMessage
'    Debug.Print "aMessageIen : " & aMessageIen
'    Debug.Print "aPriority : " & aPriority
'    Debug.Print "aPatientNcid : " & aPatientNcid
'    Debug.Print "aSurrogateNcid : " & aSurrogateNcid
'    Debug.Print "aDocumentGuid : " & aDocumentGuid
'    Debug.Print "Send End"
    
    '*******************
    ' Do database open here
    '*******************

    Set Connection = New CHCSII_Conn.Conn
    Set moGemsDAS = Connection.CHCSII_DAS
    
    Set oSurrogate = New Surrogate
    oSurrogate.ProviderNcid = aProviderNcid
    GetSurrogate oSurrogate
    
    'Create New Guid
    sMessageGuid = GetGUIDString
    
    'Find message type
    'scr 25245 - 01JUL2002 - j.ray - added support for additional messages
    If InStr(1, aMessageTypeGuid, "{") Then
        'do nothing
    ElseIf InStr(1, aMessageTypeGuid, "Test") Then
        'do nothing
    ElseIf InStr(1, aMessageTypeGuid, "XM-NEW MAIL") Then
        aMessageTypeGuid = GUID_NEW_MAIL
    ElseIf InStr(1, aMessageTypeGuid, "OR-NEW-RESULT") Then
        aMessageTypeGuid = GUID_NEW_RESULT
    ElseIf InStr(1, aMessageTypeGuid, "OR-SURR-NEW-RESULT") Then
        aMessageTypeGuid = GUID_NEW_RESULT
    ElseIf InStr(1, aMessageTypeGuid, "OR-PRIORITY-RESULT") Then
        aMessageTypeGuid = GUID_PRIORITY_RESULT
    ElseIf InStr(1, aMessageTypeGuid, "OR-PRIORITY-RESULT") Then
        aMessageTypeGuid = GUID_PRIORITY_RESULT
    ' Begin:SCR #25245;   Developer: Eric A. Chang 07/22/2002
    ElseIf InStr(1, aMessageTypeGuid, "OR-SURR-PRIORITY-RESULT") Then
        aMessageTypeGuid = GUID_PRIORITY_RESULT
    ' End: SCR #25245;
    ElseIf InStr(1, aMessageTypeGuid, "CODING_CHANGE") Then
        aMessageTypeGuid = GUID_CODING_CHANGED
    ElseIf InStr(1, aMessageTypeGuid, "RAD-REPORT-APPROVAL") Then
        aMessageTypeGuid = GUID_RADIOLOGY_RESULTS
    ElseIf InStr(1, aMessageTypeGuid, "OR-ORDERS TO SIGN") Then
        aMessageTypeGuid = GUID_ORDERS_TO_SIGN
    
    'Pharmacy 523
    ElseIf InStr(1, aMessageTypeGuid, "OR-ORDERS TO SIGN COTS RX") Then
        aMessageTypeGuid = GUID_ORDERS_TO_SIGN_COTS_RX
        
    
    ElseIf InStr(1, aMessageTypeGuid, "OR-TC NEW TELEPHONE") Then
        aMessageTypeGuid = GUID_NEW_TELEPHONE_CONSULT
    ElseIf InStr(1, aMessageTypeGuid, "OR-TC INCOMPLETE TELE CON") Then 'Chang 7/22/02 -- SCR #25245; "CON" was missing
        aMessageTypeGuid = GUID_INCOMPLETE_TELEPHONE_CONSULT
    ElseIf InStr(1, aMessageTypeGuid, "OR-TC INC TELE CON") Then 'Chang 7/22/02 -- SCR #25245; "CON" was missing
        aMessageTypeGuid = GUID_INCOMPLETE_TELEPHONE_CONSULT
    ElseIf InStr(1, aMessageTypeGuid, "GP-CONSULT-RESULT-INCOMPLETE") Then
        aMessageTypeGuid = GUID_CONSULT_RESULT_INCOMPLETE
'    ElseIf InStr(1, aMessageTypeGuid, "513 VER") Then
'        aMessageTypeGuid = newGuid
'    ElseIf InStr(1, aMessageTypeGuid, "513 CTR") Then
'        aMessageTypeGuid = newGuid
    Else
        GoTo SendError
    End If
        
    Set oAlertType = gConfiguration.AMSTypes.Item(aMessageTypeGuid)
    
    strDateTime = Format(Now, "DDMMMYYYY HH:MM:SS")
    If Trim(aProviderNcid) = Empty Then
        aProviderNcid = "0"
    End If
    If Trim(aPatientNcid) = Empty Then
        aPatientNcid = "0"
    End If
    If Trim(aSurrogateNcid) = Empty Then
        If oSurrogate.SurrogateNcid <> Empty And oSurrogate.Active = True Then
            aSurrogateNcid = oSurrogate.SurrogateNcid
        Else
            aSurrogateNcid = "0"
        End If
    End If
    If Trim(aMessageIen) = Empty Then
        aMessageIen = "0"
    End If
    sLogonFacilityNCID = cwobjLogonEnviron.FacilityNcid
    If Trim(sLogonFacilityNCID) = Empty Then
        sLogonFacilityNCID = "0"
    End If
    If Not oAlertType.Multiplicity Then
    
        sSql = "SELECT message_type_guid " & _
                          "FROM alerts " & _
                          "WHERE message_type_guid = '" & aMessageTypeGuid & "' " & _
                          "AND message = '" & Trim(Left(aMessage, 255)) & "' " & _
                          "AND provider_ncid = '" & aProviderNcid & "'"
    
        Err.Description = sSql & " did not execute correctly."
        Set moSQL = moGemsDAS.OpenRecordset(sSql)
                
        If (moSQL.BOF And moSQL.EOF) Then
            sSql = "INSERT INTO alerts (message_guid, message_type_guid, message, message_ien, time_sent, resolved, provider_ncid, patient_ncid, surrogate_ncid, note, document_guid, facility_ncid, chcs_clear) VALUES('" & sMessageGuid & "', '" & aMessageTypeGuid & "', '" & Trim(Left(aMessage, 255)) & "', " & aMessageIen & ", '" & strDateTime & "', '0', '" & aProviderNcid & "','" & aPatientNcid & "', '" & aSurrogateNcid & "', ' ', '" & aDocumentGuid & " ', " & sLogonFacilityNCID & ", '0')"
            Err.Description = sSql & " did not execute correctly."
            moGemsDAS.ExecuteSQL (sSql)
        Else
            'Record already exists just update the date, resolved, and chcs_clear flags
            If aMessageTypeGuid = GUID_CODING_CHANGED Then
                'SCR 9794
                'If this is an updated record from Coding review, update the patient ncid as well.
                sSql = "UPDATE alerts set time_sent = '" & strDateTime & "', resolved = '0', surrogate_ncid = " & aSurrogateNcid & ", patient_ncid = " & aPatientNcid & ", chcs_clear = '0'" & " WHERE message_type_guid = '" & aMessageTypeGuid & "' AND provider_ncid = '" & aProviderNcid & "'"
                Err.Description = sSql & " did not execute correctly."
                moGemsDAS.ExecuteSQL (sSql)
            Else
                sSql = "UPDATE alerts set time_sent = '" & strDateTime & "', resolved = '0', surrogate_ncid = " & aSurrogateNcid & ", chcs_clear = '0', message = '" & Left(aMessage, 255) & "' WHERE message_type_guid = '" & aMessageTypeGuid & "' AND provider_ncid = '" & aProviderNcid & "' and message = '" & Trim(Left(aMessage, 255)) & "'"
                Err.Description = sSql & " did not execute correctly."
                moGemsDAS.ExecuteSQL (sSql)
            End If
        End If
    Else
        sSql = "SELECT message_type_guid " & _
                          "FROM alerts " & _
                          "WHERE message_type_guid = '" & aMessageTypeGuid & "' " & _
                          "AND message = '" & Trim(Left(aMessage, 255)) & "' " & _
                          "AND provider_ncid = '" & aProviderNcid & "'"
                          
        Err.Description = sSql & " did not execute correctly."
        Set moSQL = moGemsDAS.OpenRecordset(sSql)
        
        If (moSQL.BOF And moSQL.EOF) Then
            sSql = "INSERT INTO alerts (message_guid, message_type_guid, message, message_ien, time_sent, resolved, provider_ncid, patient_ncid, surrogate_ncid, note, document_guid, facility_ncid, chcs_clear) VALUES('" & sMessageGuid & "', '" & aMessageTypeGuid & "', '" & Trim(Left(aMessage, 255)) & "', " & aMessageIen & ", '" & strDateTime & "', '0', '" & aProviderNcid & "','" & aPatientNcid & "', '" & aSurrogateNcid & "', ' ', '" & aDocumentGuid & "', " & sLogonFacilityNCID & ", '0')"
            Err.Description = sSql & " did not execute correctly."
        Else
            'Record already exists just update the date, resolved, and chcs_clear flags
            If aMessageTypeGuid = GUID_CODING_CHANGED Then
                'SCR 9794
                'If this is an updated record from Coding review, update the patient ncid as well.
                sSql = "UPDATE alerts set time_sent = '" & strDateTime & "', resolved = '0', surrogate_ncid = " & aSurrogateNcid & ", patient_ncid = " & aPatientNcid & ", chcs_clear = '0', message = '" & Left(aMessage, 255) & "' WHERE message_guid = '" & moSQL("message_guid") & "'"
                Err.Description = sSql & " did not execute correctly."
                moGemsDAS.ExecuteSQL (sSql)
            Else
                sSql = "UPDATE alerts set time_sent = '" & strDateTime & "', resolved = '0', surrogate_ncid = " & aSurrogateNcid & ", chcs_clear = '0', message = '" & Trim(Left(aMessage, 255)) & "' WHERE message_type_guid = '" & aMessageTypeGuid & "' AND provider_ncid = '" & aProviderNcid & "' and message = '" & Trim(Left(aMessage, 255)) & "'"
                Err.Description = sSql & " did not execute correctly."
                moGemsDAS.ExecuteSQL (sSql)
            End If
        End If
    End If
    Err.Description = Empty
    Set moSQL = Nothing
    Set moGemsDAS = Nothing
    Set Connection = Nothing
    Set oSurrogate = Nothing
    Send = True
    
'    If cwbMainViewRunning Then
'      FrmAlerts.InitialData
'    End If
  
    Exit Function
    
SendError:
    
    Set moSQL = Nothing
    Set moGemsDAS = Nothing
    Set oAlertType = Nothing
    Set Connection = Nothing
    Set oSurrogate = Nothing
    Send = False
    
End Function

'Private Function GEMS_Send(ByVal aProviderNcid As String, _
'                    ByVal aMessageTypeGuid As String, _
'                    ByVal aMessage As String, _
'                    Optional ByVal aMessageIen As Long, _
'                    Optional ByVal aPriority As Long, _
'                    Optional ByVal aPatientNcid As String = "", _
'                    Optional ByVal aSurrogateNcid As String = "", _
'                    Optional ByVal aDocumentGuid As String = "") As Boolean
'
'    Dim Connection As CHCSII_CONN.Conn
'    Dim moSQL As ICHCSII_SQL.ImoSQL
'    Dim oSurrogate As Surrogate
'    Dim sMessageGuid As String
'    Dim strDateTime As String
'    Dim sSql As String
'    Dim oAlertType As AMSType
'    Dim sLogonFacilityNCID As String
'    On Error GoTo SendError
'
'    '*******************
'    ' Do database open here
'    '*******************
'
'    Set Connection = New CHCSII_CONN.Conn
'    Set moSQL = Connection.CHCSII_moSQL
'
'    Set oSurrogate = New Surrogate
'    oSurrogate.ProviderNcid = aProviderNcid
'    GetSurrogate oSurrogate
'
'    'Create New Guid
'    sMessageGuid = GetGUIDString
'
'    'Find message type
'    'scr 25245 - 01JUL2002 - j.ray - added support for additional messages
'    If InStr(1, aMessageTypeGuid, "{") Then
'        'do nothing
'    ElseIf InStr(1, aMessageTypeGuid, "Test") Then
'        'do nothing
'    ElseIf InStr(1, aMessageTypeGuid, "XM-NEW MAIL") Then
'        aMessageTypeGuid = GUID_NEW_MAIL
'    ElseIf InStr(1, aMessageTypeGuid, "OR-NEW-RESULT") Then
'        aMessageTypeGuid = GUID_NEW_RESULT
'    ElseIf InStr(1, aMessageTypeGuid, "OR-SURR-NEW-RESULT") Then
'        aMessageTypeGuid = GUID_NEW_RESULT
'    ElseIf InStr(1, aMessageTypeGuid, "OR-PRIORITY-RESULT") Then
'        aMessageTypeGuid = GUID_PRIORITY_RESULT
'    ElseIf InStr(1, aMessageTypeGuid, "OR-PRIORITY-RESULT") Then
'        aMessageTypeGuid = GUID_PRIORITY_RESULT
'    ' Begin:SCR #25245;   Developer: Eric A. Chang 07/22/2002
'    ElseIf InStr(1, aMessageTypeGuid, "OR-SURR-PRIORITY-RESULT") Then
'        aMessageTypeGuid = GUID_PRIORITY_RESULT
'    ' End: SCR #25245;
'    ElseIf InStr(1, aMessageTypeGuid, "CODING_CHANGE") Then
'        aMessageTypeGuid = GUID_CODING_CHANGED
'    ElseIf InStr(1, aMessageTypeGuid, "RAD-REPORT-APPROVAL") Then
'        aMessageTypeGuid = GUID_RADIOLOGY_RESULTS
'    ElseIf InStr(1, aMessageTypeGuid, "OR-ORDERS TO SIGN") Then
'        aMessageTypeGuid = GUID_ORDERS_TO_SIGN
'    ElseIf InStr(1, aMessageTypeGuid, "OR-TC NEW TELEPHONE") Then
'        aMessageTypeGuid = GUID_NEW_TELEPHONE_CONSULT
'    ElseIf InStr(1, aMessageTypeGuid, "OR-TC INCOMPLETE TELE CON") Then 'Chang 7/22/02 -- SCR #25245; "CON" was missing
'        aMessageTypeGuid = GUID_INCOMPLETE_TELEPHONE_CONSULT
'    ElseIf InStr(1, aMessageTypeGuid, "OR-TC INC TELE CON") Then 'Chang 7/22/02 -- SCR #25245; "CON" was missing
'        aMessageTypeGuid = GUID_INCOMPLETE_TELEPHONE_CONSULT
'    ElseIf InStr(1, aMessageTypeGuid, "GP-CONSULT-RESULT-INCOMPLETE") Then
'        aMessageTypeGuid = GUID_CONSULT_RESULT_INCOMPLETE
''    ElseIf InStr(1, aMessageTypeGuid, "513 VER") Then
''        aMessageTypeGuid = newGuid
''    ElseIf InStr(1, aMessageTypeGuid, "513 CTR") Then
''        aMessageTypeGuid = newGuid
'    Else
'        GoTo SendError
'    End If
'
'    Set oAlertType = gConfiguration.AMSTypes.Item(aMessageTypeGuid)
'
'    strDateTime = Format(Now, "DDMMMYYYY HH:MM:SS")
'    If Trim(aProviderNcid) = Empty Then
'        aProviderNcid = "0"
'    End If
'    If Trim(aPatientNcid) = Empty Then
'        aPatientNcid = "0"
'    End If
'    If Trim(aSurrogateNcid) = Empty Then
'        If oSurrogate.SurrogateNcid <> Empty And oSurrogate.Active = True Then
'            aSurrogateNcid = oSurrogate.SurrogateNcid
'        Else
'            aSurrogateNcid = "0"
'        End If
'    End If
'    If Trim(aMessageIen) = Empty Then
'        aMessageIen = "0"
'    End If
'    sLogonFacilityNCID = cwobjLogonEnviron.FacilityNcid
'    If Trim(sLogonFacilityNCID) = Empty Then
'        sLogonFacilityNCID = "0"
'    End If
'    If Not oAlertType.Multiplicity Then
'
'        sSql = "SELECT message_type_guid " & _
'                          "FROM alerts " & _
'                          "WHERE message_type_guid = '" & aMessageTypeGuid & "' " & _
'                          "AND message = '" & Trim(Left(aMessage, 255)) & "' " & _
'                          "AND provider_ncid = '" & aProviderNcid & "'"
'
'        Err.Description = sSql & " did not execute correctly."
'        moSQL.Execute sSql
'
'        If (moSQL.BOF And moSQL.EOF) Then
'            sSql = "INSERT INTO alerts (message_guid, message_type_guid, message, message_ien, time_sent, resolved, provider_ncid, patient_ncid, surrogate_ncid, note, document_guid, facility_ncid, chcs_clear) VALUES('" & sMessageGuid & "', '" & aMessageTypeGuid & "', '" & Trim(Left(aMessage, 255)) & "', " & aMessageIen & ", '" & strDateTime & "', '0', '" & aProviderNcid & "','" & aPatientNcid & "', '" & aSurrogateNcid & "', ' ', '" & aDocumentGuid & " ', " & sLogonFacilityNCID & ", '0')"
'            Err.Description = sSql & " did not execute correctly."
'            moSQL.Execute sSql
'        Else
'            'Record already exists just update the date, resolved, and chcs_clear flags
'            If aMessageTypeGuid = GUID_CODING_CHANGED Then
'                'SCR 9794
'                'If this is an updated record from Coding review, update the patient ncid as well.
'                sSql = "UPDATE alerts set time_sent = '" & strDateTime & "', resolved = '0', surrogate_ncid = " & aSurrogateNcid & ", patient_ncid = " & aPatientNcid & ", chcs_clear = '0'" & " WHERE message_type_guid = '" & aMessageTypeGuid & "' AND provider_ncid = '" & aProviderNcid & "'"
'                Err.Description = sSql & " did not execute correctly."
'                moSQL.Execute sSql
'            Else
'                sSql = "UPDATE alerts set time_sent = '" & strDateTime & "', resolved = '0', surrogate_ncid = " & aSurrogateNcid & ", chcs_clear = '0', message = '" & Left(aMessage, 255) & "' WHERE message_type_guid = '" & aMessageTypeGuid & "' AND provider_ncid = '" & aProviderNcid & "' and message = '" & Trim(Left(aMessage, 255)) & "'"
'                Err.Description = sSql & " did not execute correctly."
'                moSQL.Execute sSql
'            End If
'        End If
'    Else
'        sSql = "SELECT message_type_guid " & _
'                          "FROM alerts " & _
'                          "WHERE message_type_guid = '" & aMessageTypeGuid & "' " & _
'                          "AND message = '" & Trim(Left(aMessage, 255)) & "' " & _
'                          "AND provider_ncid = '" & aProviderNcid & "'"
'
'        Err.Description = sSql & " did not execute correctly."
'        moSQL.Execute sSql
'
'        If (moSQL.BOF And moSQL.EOF) Then
'            sSql = "INSERT INTO alerts (message_guid, message_type_guid, message, message_ien, time_sent, resolved, provider_ncid, patient_ncid, surrogate_ncid, note, document_guid, facility_ncid, chcs_clear) VALUES('" & sMessageGuid & "', '" & aMessageTypeGuid & "', '" & Trim(Left(aMessage, 255)) & "', " & aMessageIen & ", '" & strDateTime & "', '0', '" & aProviderNcid & "','" & aPatientNcid & "', '" & aSurrogateNcid & "', ' ', '" & aDocumentGuid & "', " & sLogonFacilityNCID & ", '0')"
'            Err.Description = sSql & " did not execute correctly."
'        Else
'            'Record already exists just update the date, resolved, and chcs_clear flags
'            If aMessageTypeGuid = GUID_CODING_CHANGED Then
'                'SCR 9794
'                'If this is an updated record from Coding review, update the patient ncid as well.
'                sSql = "UPDATE alerts set time_sent = '" & strDateTime & "', resolved = '0', surrogate_ncid = " & aSurrogateNcid & ", patient_ncid = " & aPatientNcid & ", chcs_clear = '0', message = '" & Left(aMessage, 255) & "' WHERE message_guid = '" & moSQL("message_guid") & "'"
'                Err.Description = sSql & " did not execute correctly."
'                moSQL.Execute sSql
'            Else
'                sSql = "UPDATE alerts set time_sent = '" & strDateTime & "', resolved = '0', surrogate_ncid = " & aSurrogateNcid & ", chcs_clear = '0', message = '" & Trim(Left(aMessage, 255)) & "' WHERE message_type_guid = '" & aMessageTypeGuid & "' AND provider_ncid = '" & aProviderNcid & "' and message = '" & Trim(Left(aMessage, 255)) & "'"
'                Err.Description = sSql & " did not execute correctly."
'                moSQL.Execute sSql
'            End If
'        End If
'    End If
'    Err.Description = Empty
'    Set moSQL = Nothing
'    Set oSurrogate = Nothing
'    GEMS_Send = True
'
''    If cwbMainViewRunning Then
''      FrmAlerts.InitialData
''    End If
'
'    Exit Function
'
'SendError:
'
'    Set moSQL = Nothing
'    Set oAlertType = Nothing
'    Set oSurrogate = Nothing
'    GEMS_Send = False
'
'End Function

Public Function DeleteAlert(ByVal aProviderNcid As String, ByVal aMessageGuid As String) As Boolean

    Dim Connection As CHCSII_Conn.Conn
    Dim moGemsDAS As ICHCSII_DAS
    
    Dim sSql As String
    On Error GoTo SendError
    
    '*******************
    ' Do database open here
    '*******************

    Set Connection = New CHCSII_Conn.Conn
    Set moGemsDAS = Connection.CHCSII_DAS
    
        
    'Delete the record from the alerts table
    sSql = "DELETE from alerts WHERE message_guid = '" & aMessageGuid & "' AND provider_ncid = '" & aProviderNcid & "'"
    Err.Description = sSql & " did not execute correctly."
    moGemsDAS.ExecuteSQL (sSql)
    Err.Description = Empty
    Set moGemsDAS = Nothing
    Set Connection = Nothing
    DeleteAlert = True
    Exit Function
    
SendError:
    
    Set moGemsDAS = Nothing
    Set Connection = Nothing
    DeleteAlert = False
    
End Function

'Private Function GEMS_DeleteAlert(ByVal aProviderNcid As String, ByVal aMessageGuid As String) As Boolean
'
'    Dim GemsSQL As ICHCSII_SQL.ISqlOps
'    Dim sSql As String
'    On Error GoTo SendError
'
'    '*******************
'    ' Do database open here
'    '*******************
'
'    Set GemsSQL = New GEMS_SQL.SQLOps
'
'    'Delete the record from the alerts table
'    sSql = "DELETE from alerts WHERE message_guid = '" & aMessageGuid & "' AND provider_ncid = '" & aProviderNcid & "'"
'    Err.Description = sSql & " did not execute correctly."
'    GemsSQL.Execute sSql
'    Err.Description = Empty
'    Set GemsSQL = Nothing
'    GEMS_DeleteAlert = True
'    Exit Function
'
'SendError:
'
'    Set GemsSQL = Nothing
'    GEMS_DeleteAlert = False
'
'End Function

Public Function DeleteHL7Alert(ByVal aProviderNcid As String, ByVal aMessageTypeGuid As String) As Boolean

    Dim Connection As CHCSII_Conn.Conn
    Dim moGemsDAS As ICHCSII_DAS

    Dim sSql As String
    On Error GoTo SendError
    
    '*******************
    ' Do database open here
    '*******************

    Set Connection = New CHCSII_Conn.Conn
    Set moGemsDAS = Connection.CHCSII_DAS
    'Find message type
    'scr 25245 - 01JUL2002 - j.ray - added support for additional messages
    If InStr(1, aMessageTypeGuid, "{") Then
        'do nothing
    ElseIf InStr(1, aMessageTypeGuid, "Test") Then
        'do nothing
    ElseIf InStr(1, aMessageTypeGuid, "XM-NEW MAIL") Then
        aMessageTypeGuid = GUID_NEW_MAIL
    ElseIf InStr(1, aMessageTypeGuid, "OR-NEW-RESULT") Then
        aMessageTypeGuid = GUID_NEW_RESULT
    ElseIf InStr(1, aMessageTypeGuid, "OR-SURR-NEW-RESULT") Then
        aMessageTypeGuid = GUID_NEW_RESULT
    ElseIf InStr(1, aMessageTypeGuid, "OR-PRIORITY-RESULT") Then
        aMessageTypeGuid = GUID_PRIORITY_RESULT
    ElseIf InStr(1, aMessageTypeGuid, "OR-SURR-PRIORITY-RESULT") Then
        aMessageTypeGuid = GUID_PRIORITY_RESULT
    ElseIf InStr(1, aMessageTypeGuid, "RAD-REPORT-APPROVAL") Then
        aMessageTypeGuid = GUID_RADIOLOGY_RESULTS
    ElseIf InStr(1, aMessageTypeGuid, "OR-ORDERS TO SIGN") Then
        aMessageTypeGuid = GUID_ORDERS_TO_SIGN
    'Pharmacy 523
    ElseIf InStr(1, aMessageTypeGuid, "OR-ORDERS TO SIGN COTS RX") Then
        aMessageTypeGuid = GUID_ORDERS_TO_SIGN_COTS_RX
    ElseIf InStr(1, aMessageTypeGuid, "OR-TC NEW TELEPHONE") Then
        aMessageTypeGuid = GUID_NEW_TELEPHONE_CONSULT
    ElseIf InStr(1, aMessageTypeGuid, "OR-TC-INCOMPLETE TELE") Then
        aMessageTypeGuid = GUID_INCOMPLETE_TELEPHONE_CONSULT
    ElseIf InStr(1, aMessageTypeGuid, "OR-TC INCOMPLETE TELE") Then
        'scr 26606 - 14AUG2002 - j.ray - unhandled type: OR-TC INCOMPLETE TELE CON
        aMessageTypeGuid = GUID_INCOMPLETE_TELEPHONE_CONSULT
    ElseIf InStr(1, aMessageTypeGuid, "OR-TC-INC TELE") Then
        aMessageTypeGuid = GUID_INCOMPLETE_TELEPHONE_CONSULT
    ElseIf InStr(1, aMessageTypeGuid, "OR-TC INC TELE") Then
        aMessageTypeGuid = GUID_INCOMPLETE_TELEPHONE_CONSULT
    ElseIf InStr(1, aMessageTypeGuid, "GP-CONSULT-RESULT-INCOMPLETE") Then
        aMessageTypeGuid = GUID_CONSULT_RESULT_INCOMPLETE
'    ElseIf InStr(1, aMessageTypeGuid, "513 VER") Then
'        aMessageTypeGuid = newGuid
'    ElseIf InStr(1, aMessageTypeGuid, "513 CTR") Then
'        aMessageTypeGuid = newGuid
    Else
        GoTo SendError
    End If

        
    'Delete the record from the alerts table
    sSql = "DELETE from alerts WHERE message_type_guid = '" & aMessageTypeGuid & "' AND provider_ncid = '" & aProviderNcid & "'"
    Err.Description = sSql & " did not execute correctly."
    moGemsDAS.ExecuteSQL (sSql)
    Err.Description = Empty
    Set moGemsDAS = Nothing
    DeleteHL7Alert = True
    Exit Function
    
SendError:
    
    Set moGemsDAS = Nothing
    DeleteHL7Alert = False
    
End Function

'Public Function GEMS_DeleteHL7Alert(ByVal aProviderNcid As String, ByVal aMessageTypeGuid As String) As Boolean
'
'    Dim GemsSQL As ICHCSII_SQL.ISqlOps
'    Dim sSql As String
'    On Error GoTo SendError
'
'    '*******************
'    ' Do database open here
'    '*******************
'
'    Set GemsSQL = New GEMS_SQL.SQLOps
'
'    'Find message type
'    'scr 25245 - 01JUL2002 - j.ray - added support for additional messages
'    If InStr(1, aMessageTypeGuid, "{") Then
'        'do nothing
'    ElseIf InStr(1, aMessageTypeGuid, "Test") Then
'        'do nothing
'    ElseIf InStr(1, aMessageTypeGuid, "XM-NEW MAIL") Then
'        aMessageTypeGuid = GUID_NEW_MAIL
'    ElseIf InStr(1, aMessageTypeGuid, "OR-NEW-RESULT") Then
'        aMessageTypeGuid = GUID_NEW_RESULT
'    ElseIf InStr(1, aMessageTypeGuid, "OR-SURR-NEW-RESULT") Then
'        aMessageTypeGuid = GUID_NEW_RESULT
'    ElseIf InStr(1, aMessageTypeGuid, "OR-PRIORITY-RESULT") Then
'        aMessageTypeGuid = GUID_PRIORITY_RESULT
'    ElseIf InStr(1, aMessageTypeGuid, "OR-SURR-PRIORITY-RESULT") Then
'        aMessageTypeGuid = GUID_PRIORITY_RESULT
'    ElseIf InStr(1, aMessageTypeGuid, "RAD-REPORT-APPROVAL") Then
'        aMessageTypeGuid = GUID_RADIOLOGY_RESULTS
'    ElseIf InStr(1, aMessageTypeGuid, "OR-ORDERS TO SIGN") Then
'        aMessageTypeGuid = GUID_ORDERS_TO_SIGN
'    ElseIf InStr(1, aMessageTypeGuid, "OR-TC NEW TELEPHONE") Then
'        aMessageTypeGuid = GUID_NEW_TELEPHONE_CONSULT
'    ElseIf InStr(1, aMessageTypeGuid, "OR-TC-INCOMPLETE TELE") Then
'        aMessageTypeGuid = GUID_INCOMPLETE_TELEPHONE_CONSULT
'    ElseIf InStr(1, aMessageTypeGuid, "OR-TC-INC TELE") Then
'        aMessageTypeGuid = GUID_INCOMPLETE_TELEPHONE_CONSULT
'    ElseIf InStr(1, aMessageTypeGuid, "GP-CONSULT-RESULT-INCOMPLETE") Then
'        aMessageTypeGuid = GUID_CONSULT_RESULT_INCOMPLETE
''    ElseIf InStr(1, aMessageTypeGuid, "513 VER") Then
''        aMessageTypeGuid = newGuid
''    ElseIf InStr(1, aMessageTypeGuid, "513 CTR") Then
''        aMessageTypeGuid = newGuid
'    Else
'        GoTo SendError
'    End If
'
'    'Delete the record from the alerts table
'    sSql = "DELETE from alerts WHERE message_type_guid = '" & aMessageTypeGuid & "' AND provider_ncid = '" & aProviderNcid & "'"
'    Err.Description = sSql & " did not execute correctly."
'    GemsSQL.Execute sSql
'    Err.Description = Empty
'    Set GemsSQL = Nothing
'    GEMS_DeleteHL7Alert = True
'    Exit Function
'
'SendError:
'
'    Set GemsSQL = Nothing
'    GEMS_DeleteHL7Alert = False
'
'End Function

'This function should only be called from the "Delete Alert" menu button.
Public Function ResolveAlert(ByVal aProviderNcid As String, ByVal aAlert As Alert) As Boolean
    
    Dim Connection As CHCSII_Conn.Conn
    Dim moGemsDAS As ICHCSII_DAS
    
    Dim sSql As String
    On Error GoTo SendError
    
    '*******************
    ' Do database open here
    '*******************

    Set Connection = New CHCSII_Conn.Conn
    Set moGemsDAS = Connection.CHCSII_DAS
    
    If Not aAlert.SurrogateNcid = cwobjLogonEnviron.UserNCID Then
        moGemsDAS.ExecuteSQL "UPDATE alerts SET resolved = 1, chcs_clear = '1', note = '(D) " & aAlert.Note & "' WHERE message_guid = '" & aAlert.MessageGuid & "' AND message = '" & aAlert.Message & "' AND provider_ncid = '" & cwobjLogonEnviron.UserNCID & "'"
    Else
        moGemsDAS.ExecuteSQL "UPDATE alerts SET resolved = 1, surrogate_ncid = ' ', note = '(D) " & aAlert.Note & "' WHERE message_guid = '" & aAlert.MessageGuid & "' AND message = '" & aAlert.Message & "' AND provider_ncid = '" & cwobjLogonEnviron.UserNCID & "'"
    End If
    
    If aAlert.MessageTypeGuid = GUID_ENCOUNTER_DOC_NEEDS_COSIGN Then
        mvarEncounterCosignsNeeded = False
    End If
    
    Err.Description = sSql & " did not execute correctly."
    moGemsDAS.ExecuteSQL (sSql)
    Err.Description = Empty
    Set moGemsDAS = Nothing
    
    ResolveAlert = True
    Exit Function
    
SendError:
    
    Set moGemsDAS = Nothing
    ResolveAlert = False
    
End Function

'  Private Function GEMS_ResolveAlert(ByVal aProviderNcid As String, ByVal aAlert As Alert) As Boolean
'
'      Dim GemsSQL As ICHCSII_SQL.ISqlOps
'      Dim sSql As String
'      On Error GoTo SendError
'
'      '*******************
'      ' Do database open here
'      '*******************
'
'      Set GemsSQL = New GEMS_SQL.SQLOps
'
'      If Not aAlert.SurrogateNcid = cwobjLogonEnviron.UserNCID Then
'          GemsSQL.Execute "UPDATE alerts SET resolved = 1, chcs_clear = '1', note = '(D) " & aAlert.Note & "' WHERE message_guid = '" & aAlert.MessageGuid & "' AND message = '" & aAlert.Message & "' AND provider_ncid = '" & cwobjLogonEnviron.UserNCID & "'"
'      Else
'          GemsSQL.Execute "UPDATE alerts SET resolved = 1, surrogate_ncid = ' ', note = '(D) " & aAlert.Note & "' WHERE message_guid = '" & aAlert.MessageGuid & "' AND message = '" & aAlert.Message & "' AND provider_ncid = '" & cwobjLogonEnviron.UserNCID & "'"
'      End If
'
'      If aAlert.MessageTypeGuid = GUID_ENCOUNTER_DOC_NEEDS_COSIGN Then
'          mvarEncounterCosignsNeeded = False
'      End If
'
'      Err.Description = sSql & " did not execute correctly."
'      GemsSQL.Execute sSql
'      Err.Description = Empty
'      Set GemsSQL = Nothing
'
'      GEMS_ResolveAlert = True
'      Exit Function
'
'SendError:
'
'      Set GemsSQL = Nothing
'      GEMS_ResolveAlert = False
'
'  End Function


'********************
'SCR 26551
'25-Sep-2002
'Change Embedded SQL to Stored Procedure
'********************
Public Function Receive(aMessages As Collection, Optional bLocal As Boolean = False, Optional ByVal aPatientNcid As String, Optional ByVal aMessageTypeGuid As String) As Boolean
          
    '********************************************
    '29936
    If gobjShared.IsAppMode(modeCHCSI_GUI) Then
      Receive = CHCS_Receive(aMessages)
      Exit Function
    End If
    '********************************************
    
    If gobjShared.IsAppMode(modeTheater) Or gobjShared.IsAppMode(modeITT) Then
      Receive = GEMS_Receive(aMessages, aPatientNcid, aMessageTypeGuid)
      Exit Function
    End If
        
    Dim moGemsDAS As ICHCSII_SQL.ISqlOps
    Dim Var As CHCSII_Command.CHCSII_Cmd

    Dim oAlert As Alert
    Dim colMessages As Collection
    Dim sSql As String
    
    Dim sMessageTypeGUID As String

    On Error GoTo ReceiveError
    
    '37879
    If gbDoAlertsReceive = False Then
      GoTo CleanUp
    End If
    
    '************************************
    'SCR 20862
    'Alerts "Visual Indicator" is still being displyed when it shouldn't
    mvarEncounterCosignsNeeded = False
    mvarUnresolvedAlerts = False
    mvarOrdersToSignNeeded = False
    '************************************

    '*******************
    ' Do database open here
    '*******************
    
    Set Var = New CHCSII_Command.CHCSII_Cmd
    
    '*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#
    'SCR 37879
    'There are two stored procedures.
    'ALERTS_RECEIVE will be used when the Alerts module is populating the GUI with Data.
    'ALERTS_RECEIVE_CORE will be used when Core is refreshing it's Visual Indicators.
    '*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#
    
    If bLocal = True Then
      Var.SPName = "ALERTS_RECEIVE"
    Else
      Var.SPName = "ALERTS_RECEIVE_CORE"
    End If
    
    Call Var.AddParam(False, False, "textString", CStr(cwobjLogonEnviron.UserNCID))
    Call Var.AddParam(True, False, "", "")
    
    
    Set moGemsDAS = moConn.CHCSII_SQLOPS_SP(Auto) '<SCR 36493
    'Call moGemsDas.OpenRecordsetSP(Var)
    Call moGemsDAS.ExecuteSP(Var)
    
    If moGemsDAS.BOF And moGemsDAS.EOF Then
      '************************************
      'SCR 20320
      '11-Jan-2001
      'These two Properties need to be set to False if there are NO Alerts.
      mvarEncounterCosignsNeeded = False
      mvarUnresolvedAlerts = False
      mvarOrdersToSignNeeded = False
      '*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#
      'SCR 46155
      Receive = False
      '*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#
      '************************************
      GoTo CleanUp
    Else
      Receive = True
    End If

    mvarEncounterCosignsNeeded = False
    mvarUnresolvedAlerts = False
    
    
    moGemsDAS.MoveFirst
    
    If bLocal = True Then
          
      Set colMessages = New Collection

      While Not moGemsDAS.EOF
          Set oAlert = New Alert
          oAlert.MessageGuid = moGemsDAS("message_guid")
  
          oAlert.MessageTypeGuid = moGemsDAS("message_type_guid")
          oAlert.Message = moGemsDAS("message")
          
          Select Case Trim(UCase(oAlert.Message))
            
            Case "YOU HAVE NEW RESULTS ON CHCS."
              oAlert.DisplayMessage = "You have New Results."
  
            Case "YOU HAVE PRIORITY RESULTS ON CHCS."
              oAlert.DisplayMessage = "You have Priority Results."
          
            Case Else
              
              oAlert.DisplayMessage = oAlert.Message
            
          End Select
          
          oAlert.MessageIen = moGemsDAS("message_ien")
          oAlert.TimeSent = moGemsDAS("time_sent")
          oAlert.Priority = moGemsDAS("priority")
          oAlert.Resolved = moGemsDAS("resolved")
          oAlert.ProviderNcid = moGemsDAS("provider_ncid")
          oAlert.PatientNcid = moGemsDAS("patient_ncid")
          oAlert.SurrogateNcid = moGemsDAS("surrogate_ncid")
          oAlert.Note = moGemsDAS("note")
          oAlert.DocumentNcid = moGemsDAS("document_guid")
          oAlert.FacilityNcid = moGemsDAS("facility_ncid")
          oAlert.Chcs_Clear = moGemsDAS("chcs_clear")
          colMessages.Add oAlert, oAlert.MessageGuid
  
  
          moGemsDAS.MoveNext
      Wend
    
    Else
    
      mvarUnresolvedAlerts = True
    
      While Not moGemsDAS.EOF
      
        sMessageTypeGUID = moGemsDAS("message_type_guid")
        
        Select Case sMessageTypeGUID
        
          Case GUID_ENCOUNTER_DOC_NEEDS_COSIGN
            mvarEncounterCosignsNeeded = True
          Case GUID_ORDERS_TO_SIGN, GUID_ORDERS_TO_SIGN_COTS_RX
            mvarOrdersToSignNeeded = True
        
        End Select
        
        moGemsDAS.MoveNext
      Wend
        
      '***********************************************
      'SCR 34946
      'If there are any Unresolved "Undigned Orders" Alerts, then send a Refresh_Data Message to the "Sign Orders Module".
      If mvarOrdersToSignNeeded = True Then
        Call cwobjCWComm.Message(cwiREFRESH_DATA, "", NCID_SIGN_ORDERS_APPLICATION, NCID_ALERT_APPLICATION, False)
      End If
      '***********************************************
    
    End If
    

CleanUp:
    Set aMessages = colMessages

    Set Var = Nothing
    Set moGemsDAS = Nothing
    Set oAlert = Nothing
    
    Exit Function

ReceiveError:

    Call gobjShared.ShowVBError(Err.Number, Err.Description, "AlertsClient.Receive", "AMSClient", vbExclamation)
    Set Var = Nothing
    Set moGemsDAS = Nothing
    Set oAlert = Nothing

    Receive = False

    Exit Function

End Function

Private Function GEMS_Receive(aMessages As Collection, Optional ByVal aPatientNcid As String, Optional ByVal aMessageTypeGuid As String) As Boolean

    Dim moGemsDAS As ICHCSII_DAS
    Dim rsAlert As ADODB.Recordset

    Dim oAlert As Alert
    Dim colMessages As Collection
    Dim sSql As String
        
    On Error GoTo GEMS_ReceiveError
            
    '************************************
    'SCR 20862
    'Alerts "Visual Indicator" is still being displyed when it shouldn't
    mvarEncounterCosignsNeeded = False
    mvarUnresolvedAlerts = False
    mvarOrdersToSignNeeded = False
    '************************************

    '*******************
    ' Do database open here
    '*******************
    
    Set moGemsDAS = moConn.CHCSII_DAS(Auto) '<SCR 36493
        
    If Not aPatientNcid = "" Then
        sSql = "SELECT a.message_guid, " & _
                  "a.message_type_guid, " & _
                  "a.message, " & _
                  "a.message_ien, " & _
                  "a.time_sent, " & _
                  "a.resolved, " & _
                  "a.provider_ncid, " & _
                  "a.patient_ncid, " & _
                  "a.surrogate_ncid, " & _
                  "a.note, " & _
                  "a.document_guid, " & _
                  "a.facility_ncid, " & _
                  "a.chcs_clear, " & _
                  "at.priority " & _
                "FROM alerts a, alert_types at " & _
                "WHERE a.message_type_guid = at.message_type_guid " & _
                  "AND patient_ncid = '" & aPatientNcid & "'"
                  '"AND a.note not like '(D)%'"
        
        Set rsAlert = moGemsDAS.OpenRecordset(sSql)

    Else
        sSql = "SELECT a.message_guid, " & _
                  "a.message_type_guid, " & _
                  "a.message, " & _
                  "a.message_ien, " & _
                  "a.time_sent, " & _
                  "a.resolved, " & _
                  "a.provider_ncid, " & _
                  "a.patient_ncid, " & _
                  "a.surrogate_ncid, " & _
                  "a.note, " & _
                  "a.document_guid, " & _
                  "a.facility_ncid, " & _
                  "a.chcs_clear, " & _
                  "at.priority " & _
                "FROM alerts a, alert_types at " & _
                "WHERE a.message_type_guid = at.message_type_guid " & _
                  "AND (provider_ncid = '" & cwobjLogonEnviron.UserNCID & "' OR (surrogate_ncid = '" & cwobjLogonEnviron.UserNCID & "' AND resolved = '0')) "
                  '"AND a.note not like '(D)%'"
      
      Set rsAlert = moGemsDAS.OpenRecordset(sSql)
    
    End If
    
    Set colMessages = New Collection
    
    If rsAlert.BOF And rsAlert.EOF Then
      '************************************
      'SCR 20320
      '11-Jan-2001
      'These two Properties need to be set to False if there are NO Alerts.
      mvarEncounterCosignsNeeded = False
      mvarUnresolvedAlerts = False
      mvarOrdersToSignNeeded = False
      '************************************
      GoTo CleanUp
    End If

    mvarEncounterCosignsNeeded = False
    mvarUnresolvedAlerts = False
    While Not rsAlert.EOF
        Set oAlert = New Alert
        oAlert.MessageGuid = IIf(IsNull(rsAlert.Fields("message_guid")), "", rsAlert.Fields("message_guid"))
                
        oAlert.MessageTypeGuid = IIf(IsNull(rsAlert.Fields("message_type_guid")), "", rsAlert.Fields("message_Type_guid"))
        'MsgBox oAlert.MessageTypeGuid
        oAlert.Message = IIf(IsNull(rsAlert.Fields("message")), "", rsAlert.Fields("message"))
        
'        MsgBox oAlert.Message
        
        oAlert.MessageIen = IIf(IsNull(rsAlert.Fields("message_ien")), 0, rsAlert.Fields("message_ien"))
        oAlert.TimeSent = IIf(IsNull(rsAlert.Fields("time_sent")), CDate(0), rsAlert.Fields("time_sent"))
        oAlert.Priority = IIf(IsNull(rsAlert.Fields("priority")), 0, rsAlert.Fields("priority"))
        oAlert.Resolved = IIf(IsNull(rsAlert.Fields("resolved")), "0", rsAlert.Fields("resolved"))
        oAlert.ProviderNcid = IIf(IsNull(rsAlert.Fields("provider_ncid")), 0, rsAlert.Fields("provider_ncid"))
        oAlert.PatientNcid = IIf(IsNull(rsAlert.Fields("patient_ncid")), "", rsAlert.Fields("patient_ncid"))
        oAlert.SurrogateNcid = IIf(IsNull(rsAlert.Fields("surrogate_ncid")), "", rsAlert.Fields("surrogate_ncid"))
        oAlert.Note = IIf(IsNull(rsAlert.Fields("note")), "", rsAlert.Fields("note"))
        oAlert.DocumentNcid = IIf(IsNull(rsAlert.Fields("document_guid")), 0, rsAlert.Fields("document_guid"))
        oAlert.FacilityNcid = IIf(IsNull(rsAlert.Fields("facility_ncid")), 0, rsAlert.Fields("facility_ncid"))
        oAlert.Chcs_Clear = IIf(IsNull(rsAlert.Fields("chcs_clear")), 0, rsAlert.Fields("chcs_clear"))
        colMessages.Add oAlert, oAlert.MessageGuid
        
        If oAlert.Resolved = False Then
          mvarUnresolvedAlerts = True
          Select Case oAlert.MessageTypeGuid
            Case GUID_ENCOUNTER_DOC_NEEDS_COSIGN
              mvarEncounterCosignsNeeded = True
            '*****************************
            'SCR 14692
            'Pharmacy 523
            Case GUID_ORDERS_TO_SIGN, GUID_ORDERS_TO_SIGN_COTS_RX
              mvarOrdersToSignNeeded = True
            '*****************************
          End Select
        End If
        
        rsAlert.MoveNext
    Wend
        
    '***********************************************
    'SCR 34946
    'If there are any Unresolved "Undigned Orders" Alerts, then send a Refresh_Data Message to the "Sign Orders Module".
    If mvarOrdersToSignNeeded = True Then
      cwobjCWComm.Message cwiREFRESH_DATA, "", NCID_SIGN_ORDERS_APPLICATION, NCID_ALERT_APPLICATION
    End If
    '***********************************************

CleanUp:
    Set aMessages = colMessages
    
    Set rsAlert = Nothing
    Set oAlert = Nothing
    
    Set moGemsDAS = Nothing
        
    GEMS_Receive = True
            
    Exit Function
    
GEMS_ReceiveError:

    Call gobjShared.ShowVBError(Err.Number, Err.Description, "AlertsClient.GEMS_Receive", "AMSClient", vbExclamation)
    
    GEMS_Receive = False
    
    Exit Function

End Function


Public Function Reclaim(aAlert As Alert) As Boolean

    Dim Connection As CHCSII_Conn.Conn
    Dim moGemsDAS As ICHCSII_DAS
    Dim oAlert As Alert
        
    On Error GoTo ReclaimError
    
    '*******************
    ' Do database open here
    '*******************
    
    Set oAlert = aAlert
    
    'If this message is already resolved
    'we cannot reclaim it, so exit now.
    If oAlert.Resolved Then
        Reclaim = False
        Exit Function
    End If
    
    Set Connection = New CHCSII_Conn.Conn
    Set moGemsDAS = Connection.CHCSII_DAS

    moGemsDAS.ExecuteSQL "UPDATE alerts SET surrogate_ncid = ' ' WHERE message_guid = '" & oAlert.MessageGuid & "'"
    oAlert.SurrogateNcid = " "
            
    Set moGemsDAS = Nothing
    
    Reclaim = True
    
    Exit Function
    
ReclaimError:
    Set moGemsDAS = Nothing
    
    Reclaim = False

End Function

'Private Function GEMS_Reclaim(aAlert As Alert) As Boolean
'
'    Dim GemsSQL As ICHCSII_SQL.ISqlOps
'    Dim oAlert As Alert
'
'    On Error GoTo ReclaimError
'
'    '*******************
'    ' Do database open here
'    '*******************
'
'    Set oAlert = aAlert
'
'    'If this message is already resolved
'    'we cannot reclaim it, so exit now.
'    If oAlert.Resolved Then
'        GEMS_Reclaim = False
'        Exit Function
'    End If
'
'    Set GemsSQL = New GEMS_SQL.SQLOps
'
'    GemsSQL.Execute "UPDATE alerts SET surrogate_ncid = ' ' WHERE message_guid = '" & oAlert.MessageGuid & "'"
'    oAlert.SurrogateNcid = " "
'
'    Set GemsSQL = Nothing
'
'    GEMS_Reclaim = True
'
'    Exit Function
'
'ReclaimError:
'    Set GemsSQL = Nothing
'
'    GEMS_Reclaim = False
'
'End Function

Public Function Resolve(Optional ByVal aDocumentNcid As String, Optional aAlert As Alert, Optional ByVal aNote As String = "") As Boolean
    
    
    Dim Connection As CHCSII_Conn.Conn
    Dim moGemsDAS As ICHCSII_DAS
        
    Dim oAlert As Alert
    Dim sSql As String
        
    On Error GoTo ResolveError
    
    '*******************
    ' Do database open here
    '*******************
    
    Set oAlert = aAlert
    
    'If this message is already resolved
    'we cannot resolve it again so exit now.
    If oAlert.Resolved Then
        Resolve = False
        Exit Function
    End If
    
    Set Connection = New CHCSII_Conn.Conn
    Set moGemsDAS = Connection.CHCSII_DAS

'    If oAlert.MessageTypeGuid = GUID_NEW_RESULT Or oAlert.MessageTypeGuid = GUID_PRIORITY_RESULT Then
'
'      'Comment this out while testing
'      If oAlert.ProviderNcid = cwobjLogonEnviron.UserNCID Then
'          sSql = "UPDATE alerts " & _
'                            "SET Surrogate_NCID = NULL, resolved = 1, chcs_clear = '0', note = '" & aNote & "' " & _
'                            "WHERE MESSAGE_TYPE_GUID in ('" & GUID_NEW_RESULT & "', '" & GUID_PRIORITY_RESULT & "') " & _
'                                "AND provider_ncid = '" & cwobjLogonEnviron.UserNCID & "'"
'      Else
'        If oAlert.SurrogateNcid = cwobjLogonEnviron.UserNCID Then
'          sSql = "UPDATE alerts " & _
'                            "SET resolved = 0, chcs_clear = '0', surrogate_ncid = NULL, note = '" & aNote & "' " & _
'                            "WHERE MESSAGE_TYPE_GUID in ('" & GUID_NEW_RESULT & "', '" & GUID_PRIORITY_RESULT & "') " & _
'                                "AND surrogate_ncid = '" & cwobjLogonEnviron.UserNCID & "'"
'        End If
'      End If
'
'      'MsgBox sSql
'      moGemsDas.ExecuteSQL(ssql)
'
''      'NOTE: We will began using the MessageTranseiver Interface when
''      'modules began implementing it.
''      sSql = "UPDATE alerts " & _
''              "SET resolved = '1' " & _
''              "WHERE MESSAGE_TYPE_GUID in ('" & GUID_NEW_RESULT & "', '" & GUID_PRIORITY_RESULT & "') " & _
''                  "AND provider_ncid = '" & oAlert.ProviderNcid & "'"
''      'sSql = "UPDATE alerts set resolved = '1' WHERE message_type_guid = '" & oAlert.MessageTypeGuid & "' AND provider_ncid = '" & oAlert.ProviderNcid & "'"
''      Err.Description = sSql & " did not execute correctly."
''      moGemsDas.ExecuteSQL(ssql)

    'Pharmacy 523
    If oAlert.MessageTypeGuid = GUID_ORDERS_TO_SIGN Or oAlert.MessageTypeGuid = GUID_ORDERS_TO_SIGN_COTS_RX Then

      '***************
      'SCR 14668
      'Orders-To-Sign is a new Alert Type
      '***************

    '******************************************************************

      'For Orders-To-Sign Alerts, donot Resolve them.
      '  Let them get deleted by Middleman when the Order is Signed in Legacy.

      If oAlert.SurrogateNcid = cwobjLogonEnviron.UserNCID Then
        'Surrogate is Logged in.
          sSql = "UPDATE alerts SET resolved = 0, surrogate_ncid = NULL, chcs_clear = '0', " & _
                    "note = '" & aNote & "' WHERE MESSAGE_TYPE_GUID = '" & oAlert.MessageTypeGuid & "' " & _
                        "AND message = '" & oAlert.Message & "' " & _
                        "AND surrogate_ncid = '" & cwobjLogonEnviron.UserNCID & "'"
      Else
        'Primary Provider is Looged in
          sSql = "UPDATE alerts SET resolved = 0, surrogate_ncid = NULL, chcs_clear = '0', " & _
                      "note = '" & aNote & "' " & _
                "WHERE MESSAGE_TYPE_GUID = '" & oAlert.MessageTypeGuid & "' " & _
                  "AND message = '" & oAlert.Message & "' " & _
                  "AND provider_ncid = '" & cwobjLogonEnviron.UserNCID & "'"
      End If

      Err.Description = sSql & " did not execute correctly."
      moGemsDAS.ExecuteSQL (sSql)
    '******************************************************************

    Else

      '******************
      'SCR 15426
      'Bill Fote 09/18/2001
      'Don't Set the Surrogate NCID to NULL then the Alert is Addressed.
      '******************
      'Comment this out while testing
      If Not oAlert.SurrogateNcid = cwobjLogonEnviron.UserNCID Then
        'Primary Provider is Looged in
          sSql = "UPDATE alerts SET resolved = 1, chcs_clear = '0', " & _
                      "note = '" & aNote & "' " & _
                "WHERE message_guid = '" & aAlert.MessageGuid & "' " & _
                  "AND message = '" & oAlert.Message & "' " & _
                  "AND provider_ncid = '" & cwobjLogonEnviron.UserNCID & "'"
      Else
        'Surrogate is Logged in.
          sSql = "UPDATE alerts SET resolved = 1, " & _
                    "note = '" & aNote & "' WHERE message_guid = '" & aAlert.MessageGuid & "' " & _
                        "AND message = '" & oAlert.Message & "' " & _
                        "AND surrogate_ncid = '" & cwobjLogonEnviron.UserNCID & "'"
      End If
      Err.Description = sSql & " did not execute correctly."
      moGemsDAS.ExecuteSQL (sSql)

'      'NOTE: We will began using the MessageTranseiver Interface when
'      'modules began implementing it.
'      sSql = "UPDATE alerts set resolved = '1' WHERE message_guid = '" & oAlert.MessageGuid & "' AND provider_ncid = '" & oAlert.ProviderNcid & "'"
'      'sSql = "UPDATE alerts set resolved = '1' WHERE message_type_guid = '" & oAlert.MessageTypeGuid & "' AND provider_ncid = '" & oAlert.ProviderNcid & "'"
'      Err.Description = sSql & " did not execute correctly."
'      moGemsDas.ExecuteSQL(ssql)

    End If

    If oAlert.MessageTypeGuid = GUID_ENCOUNTER_DOC_NEEDS_COSIGN Then
        mvarEncounterCosignsNeeded = False
    End If

    oAlert.Resolved = True
    oAlert.Note = aNote

    Set moGemsDAS = Nothing
    
    Resolve = True
    
    Exit Function
    
ResolveError:

    Call gobjShared.ShowVBError(Err.Number, Err.Description, "AlertsClient.Resolve", "AMSClient", vbExclamation)
            
    Resolve = False
    
    Exit Function

End Function

'Private Function GEMS_Resolve(Optional ByVal aDocumentNcid As String, Optional aAlert As Alert, Optional ByVal aNote As String = "") As Boolean
'
'    Dim GemsSQL As ICHCSII_SQL.ISqlOps
'
'    Dim oAlert As Alert
'    Dim sSql As String
'
'    On Error GoTo ResolveError
'
'    '*******************
'    ' Do database open here
'    '*******************
'
'    Set oAlert = aAlert
'
'    'If this message is already resolved
'    'we cannot resolve it again so exit now.
'    If oAlert.Resolved Then
'        GEMS_Resolve = False
'        Exit Function
'    End If
'
'    Set GemsSQL = New GEMS_SQL.SQLOps
'
''    If oAlert.MessageTypeGuid = GUID_NEW_RESULT Or oAlert.MessageTypeGuid = GUID_PRIORITY_RESULT Then
''
''      'Comment this out while testing
''      If oAlert.ProviderNcid = cwobjLogonEnviron.UserNCID Then
''          sSql = "UPDATE alerts " & _
''                            "SET Surrogate_NCID = NULL, resolved = 1, chcs_clear = '0', note = '" & aNote & "' " & _
''                            "WHERE MESSAGE_TYPE_GUID in ('" & GUID_NEW_RESULT & "', '" & GUID_PRIORITY_RESULT & "') " & _
''                                "AND provider_ncid = '" & cwobjLogonEnviron.UserNCID & "'"
''      Else
''        If oAlert.SurrogateNcid = cwobjLogonEnviron.UserNCID Then
''          sSql = "UPDATE alerts " & _
''                            "SET resolved = 0, chcs_clear = '0', surrogate_ncid = NULL, note = '" & aNote & "' " & _
''                            "WHERE MESSAGE_TYPE_GUID in ('" & GUID_NEW_RESULT & "', '" & GUID_PRIORITY_RESULT & "') " & _
''                                "AND surrogate_ncid = '" & cwobjLogonEnviron.UserNCID & "'"
''        End If
''      End If
''
''      'MsgBox sSql
''      GemsSQL.Execute sSql
'
'''      'NOTE: We will began using the MessageTranseiver Interface when
'''      'modules began implementing it.
'''      sSql = "UPDATE alerts " & _
'''              "SET resolved = '1' " & _
'''              "WHERE MESSAGE_TYPE_GUID in ('" & GUID_NEW_RESULT & "', '" & GUID_PRIORITY_RESULT & "') " & _
'''                  "AND provider_ncid = '" & oAlert.ProviderNcid & "'"
'''      'sSql = "UPDATE alerts set resolved = '1' WHERE message_type_guid = '" & oAlert.MessageTypeGuid & "' AND provider_ncid = '" & oAlert.ProviderNcid & "'"
'''      Err.Description = sSql & " did not execute correctly."
'''      GemsSQL.Execute sSql
'
'    If oAlert.MessageTypeGuid = GUID_ORDERS_TO_SIGN Then
'
'      '***************
'      'SCR 14668
'      'Orders-To-Sign is a new Alert Type
'      '***************
'
'    '******************************************************************
'
'      'For Orders-To-Sign Alerts, donot Resolve them.
'      '  Let them get deleted by Middleman when the Order is Signed in Legacy.
'
'      If oAlert.SurrogateNcid = cwobjLogonEnviron.UserNCID Then
'        'Surrogate is Logged in.
'          sSql = "UPDATE alerts SET resolved = 0, surrogate_ncid = NULL, chcs_clear = '0', " & _
'                    "note = '" & aNote & "' WHERE MESSAGE_TYPE_GUID = '" & GUID_ORDERS_TO_SIGN & "' " & _
'                        "AND message = '" & oAlert.Message & "' " & _
'                        "AND surrogate_ncid = '" & cwobjLogonEnviron.UserNCID & "'"
'      Else
'        'Primary Provider is Looged in
'          sSql = "UPDATE alerts SET resolved = 0, surrogate_ncid = NULL, chcs_clear = '0', " & _
'                      "note = '" & aNote & "' " & _
'                "WHERE MESSAGE_TYPE_GUID = '" & GUID_ORDERS_TO_SIGN & "' " & _
'                  "AND message = '" & oAlert.Message & "' " & _
'                  "AND provider_ncid = '" & cwobjLogonEnviron.UserNCID & "'"
'      End If
'
'      Err.Description = sSql & " did not execute correctly."
'      GemsSQL.Execute sSql
'    '******************************************************************
'
'    Else
'
'      '******************
'      'SCR 15426
'      'Bill Fote 09/18/2001
'      'Don't Set the Surrogate NCID to NULL then the Alert is Addressed.
'      '******************
'      'Comment this out while testing
'      If Not oAlert.SurrogateNcid = cwobjLogonEnviron.UserNCID Then
'        'Primary Provider is Looged in
'          sSql = "UPDATE alerts SET resolved = 1, chcs_clear = '0', " & _
'                      "note = '" & aNote & "' " & _
'                "WHERE message_guid = '" & aAlert.MessageGuid & "' " & _
'                  "AND message = '" & oAlert.Message & "' " & _
'                  "AND provider_ncid = '" & cwobjLogonEnviron.UserNCID & "'"
'      Else
'        'Surrogate is Logged in.
'          sSql = "UPDATE alerts SET resolved = 1, " & _
'                    "note = '" & aNote & "' WHERE message_guid = '" & aAlert.MessageGuid & "' " & _
'                        "AND message = '" & oAlert.Message & "' " & _
'                        "AND surrogate_ncid = '" & cwobjLogonEnviron.UserNCID & "'"
'      End If
'      Err.Description = sSql & " did not execute correctly."
'      GemsSQL.Execute sSql
'
''      'NOTE: We will began using the MessageTranseiver Interface when
''      'modules began implementing it.
''      sSql = "UPDATE alerts set resolved = '1' WHERE message_guid = '" & oAlert.MessageGuid & "' AND provider_ncid = '" & oAlert.ProviderNcid & "'"
''      'sSql = "UPDATE alerts set resolved = '1' WHERE message_type_guid = '" & oAlert.MessageTypeGuid & "' AND provider_ncid = '" & oAlert.ProviderNcid & "'"
''      Err.Description = sSql & " did not execute correctly."
''      GemsSQL.Execute sSql
'
'    End If
'
'    If oAlert.MessageTypeGuid = GUID_ENCOUNTER_DOC_NEEDS_COSIGN Then
'        mvarEncounterCosignsNeeded = False
'    End If
'
'    oAlert.Resolved = True
'    oAlert.Note = aNote
'
'    Set GemsSQL = Nothing
'
'    GEMS_Resolve = True
'
'    Exit Function
'
'ResolveError:
'
'    Call gobjShared.ShowVBError(Err.Number, Err.Description, "AlertsClient.GEMS_Resolve", "AMSClient", vbExclamation)
'
'    GEMS_Resolve = False
'
'    Exit Function
'
'End Function

Public Function Clear(aAlert As Alert) As Boolean

  Clear = True

'    Dim moGemsDAS As ICHCSII_DAS
'    Dim oAlert As Alert
'
'    Dim objUser As ICHCSII_User
'
'    On Error GoTo ClearError
'
'    '*******************
'    ' Do database open here
'    '*******************
'
'    Set oAlert = aAlert
'
'    Set SQLOps = New CHCSII_SQL.SqlOps
'    SQLOps.Execute "DELETE FROM alerts WHERE message_guid = '" & oAlert.MessageGuid & "'"
'
'    Set moGemsDAS = nothing
'
'    Clear = True
'
'    Exit Function
'
'ClearError:
'    Set moGemsDAS = nothing
'
'    Clear = False

End Function

Public Sub Configure(aConfiguration As Configuration)

    Dim cAMSTypes As New Collection
    Dim oAMSType As AMSType
    Dim cSurrogateIds As Collection
    Dim oConfiguration As Configuration
    
    Dim Connection As CHCSII_Conn.Conn
    Dim moGemsDAS As ICHCSII_DAS
    Dim SQLOps As ADODB.Recordset
    
    On Error Resume Next
        
    Set Connection = New CHCSII_Conn.Conn
    Set moGemsDAS = Connection.CHCSII_DAS
    Set oConfiguration = New Configuration
    
    Set SQLOps = moGemsDAS.OpenRecordset("SELECT * FROM alert_types")
    
    If SQLOps.BOF And SQLOps.EOF Then
        GoTo Surrogates
    End If

    While Not SQLOps.EOF
        
        Set oAMSType = New AMSType
        
        oAMSType.MessageType = SQLOps("message_type_guid")
        oAMSType.Label = SQLOps("label")
        oAMSType.Priority = SQLOps("priority")
      
        '****************************************************
        'SCR 18417
        'The Interaction should be NCID_CODING_REVIEW for CODING_CHANGED Alerts Types
        'The "Interaction" field in the Alert_Types table is being modified some how.
        'That is why it is being set in this manner.
        '***************
        'SCR 14668
        'Orders-To-Sign is a new Alert Type
        '***************
        Select Case Trim(UCase(oAMSType.Label))
          Case "CODING_CHANGED"
            oAMSType.Interaction = NCID_CODING_REVIEW
          Case "ORDERS_TO_SIGN"
            oAMSType.Interaction = NCID_SIGN_ORDERS_APPLICATION
          Case "NEW_RESULT", "PRIORITY_RESULT"
            oAMSType.Interaction = NCID_NEW_RESULTS_APPLICATION
          Case Else
            oAMSType.Interaction = SQLOps("interaction")
        End Select
        '******************************************************
        
        If SQLOps("multiplicity") = "Y" Then
            oAMSType.Multiplicity = True
        Else
            oAMSType.Multiplicity = False
        End If
        cAMSTypes.Add oAMSType, oAMSType.MessageType
        SQLOps.MoveNext
        
    Wend
    
Surrogates:
                   
    If (ReceiveSurrogates(cSurrogateIds) > 0) Then
        oConfiguration.SurrogateIds = cSurrogateIds
    Else
        oConfiguration.SurrogateIds = New Collection
    End If
    
    oConfiguration.AMSTypes = cAMSTypes
    Set aConfiguration = oConfiguration
    
    Set SQLOps = Nothing
    Set moGemsDAS = Nothing
    Set Connection = Nothing
    
    Exit Sub
    
End Sub



Public Function ActivateSurrogate(ByVal aSurrogateNcid As String) As Integer


  ActivateSurrogate = -1
    
End Function


Public Function AddSurrogate(ByVal aSurrogateNcid As String, ByVal aPrecedence As Integer, Optional ByVal aActive As Boolean = False) As Integer
        
        AddSurrogate = -3245
            
End Function


Public Function RemoveSurrogate(ByVal aSurrogateNcid As String) As Integer


  RemoveSurrogate = -3245


End Function



Public Function ReceiveSurrogates(aSurrogates As Collection, Optional ByVal Unused As Boolean = False) As Integer
    
  ReceiveSurrogates = -3245
    


End Function

Public Function HasSurrogate(aAlert As Alert) As Integer

    Dim Connection As CHCSII_Conn.Conn
    Dim moGemsDAS As ICHCSII_DAS
    Dim moSQL As ADODB.Recordset
        
    On Error GoTo HasSurrogateError
    
    '*******************
    ' Do database open here
    '*******************
    Set Connection = New CHCSII_Conn.Conn
    Set moGemsDAS = Connection.CHCSII_DAS
    
    'moSQL.Execute "SELECT * FROM alerts WHERE message_guid = '" & aAlert.MessageGuid & "'"
    
    Set moSQL = moGemsDAS.OpenRecordset("SELECT * FROM alerts WHERE message_guid = '" & aAlert.MessageGuid & "'")
    
    If moSQL("surrogate_ncid") = "0" Or moSQL("surrogate_ncid") = Empty Then
     
        HasSurrogate = False
    Else
    
        HasSurrogate = True
        
    End If
    
    Set moSQL = Nothing
    Set moGemsDAS = Nothing
    Set Connection = Nothing
        
    Exit Function
    
    
HasSurrogateError:

    Set moSQL = Nothing
    Set moGemsDAS = Nothing
    Set Connection = Nothing
    
    MsgBox "moSQL error in HasSurrogate function"
    
    HasSurrogate = False

End Function


Public Sub AppendNote(ByVal message_guid As String, ByVal Note As String)
    
    Dim Connection As CHCSII_Conn.Conn
    Dim moGemsDAS As ICHCSII_DAS
    Dim moSQL As ADODB.Recordset
    
    Dim sNote As String
    Dim sTag As String
    
    On Error Resume Next
        
    Set Connection = New CHCSII_Conn.Conn
    Set moGemsDAS = Connection.CHCSII_DAS
    
    Set moSQL = moGemsDAS.OpenRecordset("SELECT note from alerts WHERE message_guid = '" & message_guid & "'")
        
    If moSQL.BOF And moSQL.EOF Then
        GoTo CleanUp
    End If
    
    sNote = moSQL("note")
    moGemsDAS.ExecuteSQL "UPDATE alerts SET note = '" & Note & "' WHERE message_guid = '" & message_guid & "'"
    
CleanUp:
        
    Set moGemsDAS = Nothing
    Set Connection = Nothing
    
End Sub

Private Sub MessageTransceiver_Ack(o_rAlert As Alert)

End Sub

Private Sub MessageTransceiver_Receive(o_rAlert As Alert)

End Sub

Private Sub MessageTransceiver_Transmit(o_rAlert As Alert)

End Sub

'*********************************************
'10-Jan-2002
'SCR 19958
'This will Delete ALL records from the Surrogates Table.
'*********************************************
Private Function DeleteSurrogates() As Boolean

  DeleteSurrogates = True
    
End Function


Public Function DeleteAlertType(ByVal aProviderNcid As String, ByVal aMessageTypeGuid As String) As Boolean

    Dim Connection As CHCSII_Conn.Conn
    Dim moGemsDAS As ICHCSII_DAS
    
    Dim sSql As String
    On Error GoTo SendError
    
    '*******************
    ' Do database open here
    '*******************

    Set Connection = New CHCSII_Conn.Conn
    Set moGemsDAS = Connection.CHCSII_DAS
    
        
    'Delete the record from the alerts table
    sSql = "DELETE from alerts WHERE message_type_guid = '" & aMessageTypeGuid & "' AND provider_ncid = '" & aProviderNcid & "'"
    Err.Description = sSql & " did not execute correctly."
    moGemsDAS.ExecuteSQL (sSql)
    Err.Description = Empty
    Set moGemsDAS = Nothing
    DeleteAlertType = True
    cwobjCWComm.Message cwiRESET_VISUAL_INDICATORS, "", NCID_CW_APPLICATION, NCID_ALERT_APPLICATION
    Exit Function
    
SendError:
    
    Set moGemsDAS = Nothing
    DeleteAlertType = False
    
End Function


Private Function CHCS_Receive(aMessages As Collection) As Boolean

On Error GoTo ErrHandler
    
Dim nRecordCount As Integer
    
Dim oCHCSInterface_Alerts As ICHCSAlerts

Dim oAlert As Alert
Dim colMessages As Collection

    
  CHCS_Receive = False
  mvarEncounterCosignsNeeded = False
  mvarUnresolvedAlerts = False
  mvarOrdersToSignNeeded = False
  mvarAlertsToolTip = ""
    
  Set oCHCSInterface_Alerts = gobjShared.CHCSConnection
  Set CHCS_XML_RecordSet = oCHCSInterface_Alerts.RetrieveAlertsCurrentUser(False) 'modified on build box by j.ray 03jun2003 to not force refresh
    
  Set colMessages = New Collection
  
  If CHCS_XML_RecordSet.EOF And CHCS_XML_RecordSet.BOF Then
  
    CHCS_Receive = True
    
  Else
  
    CHCS_Receive = True
    
    mvarUnresolvedAlerts = True
    CHCS_XML_RecordSet.MoveFirst
    
    Do Until CHCS_XML_RecordSet.EOF
    
      If mvarAlertsToolTip = "" Then
        mvarAlertsToolTip = CHCS_XML_RecordSet("MESSAGE_TEXT") & vbNullString
      Else
        mvarAlertsToolTip = mvarAlertsToolTip & ", " & CHCS_XML_RecordSet("MESSAGE_TEXT") & vbNullString
      End If
    
      Set oAlert = New Alert
      
        '37612 - MessageGuid needs to be Unique.
        oAlert.MessageGuid = GetGUIDString
        '37612 - Assign Message_IEN to MessageTypeGuid
        oAlert.MessageTypeGuid = CHCS_XML_RecordSet("MESSAGE_IEN") & vbNullString
        oAlert.Message = CHCS_XML_RecordSet("MESSAGE_TEXT") & vbNullString
        oAlert.DisplayMessage = CHCS_XML_RecordSet("MESSAGE_TEXT") & vbNullString
        oAlert.MessageIen = 1
        
        If IsDate(CHCS_XML_RecordSet("MESSAGE_DATE") & vbNullString) Then
          oAlert.TimeSent = CDate(CHCS_XML_RecordSet("MESSAGE_DATE") & vbNullString)
        Else
          oAlert.TimeSent = CDate(0)
        End If
        
        oAlert.Priority = 0
        oAlert.Resolved = False
        oAlert.ProviderNcid = CHCS_XML_RecordSet("USER_IEN") & vbNullString
        oAlert.PatientNcid = 0
        oAlert.SurrogateNcid = 0
        oAlert.Note = ""
        oAlert.DocumentNcid = ""
        oAlert.FacilityNcid = 0
        oAlert.Chcs_Clear = False
        colMessages.Add oAlert, oAlert.MessageGuid
        CHCS_XML_RecordSet.MoveNext
        Set oAlert = Nothing
    
    Loop
  
  End If
  
Set aMessages = colMessages
Set oAlert = Nothing

Exit Function

ErrHandler:

    Call gobjShared.ShowVBError(Err.Number, Err.Description, "AlertsClient.CHCS_Receive", "AMSClient", vbExclamation)

    CHCS_Receive = False
    
End Function
