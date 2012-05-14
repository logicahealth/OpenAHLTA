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
Attribute VB_Name = "CITA_Admin_Init"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Implements ICHCSII_CoreAPI2
Implements ICHCSII_CoreAPI3



'Private glOleMenuItem0, glOleMenuItem1, glOleMenuItem2, glOleMenuItem3, glOleMenuItem4, glOleMenuItem5
'Private ghTemplate0, ghTemplate1, ghTemplate2, ghTemplate3




'**************************************************************
'  Description: This method is called when a client wants to know if this OLE Server
'               supports a specific "view type".
'  Parameters:
'    viViewType - Contains the "type" of view being queried for.
'  Return Value:
'    Returns TRUE if this OLE Server supports the passed in "view type".
'    Returns FALSE if this OLE Server does not support the passed in "view type".
'************************************************************
Public Function QueryView(ByVal viViewType As Integer) As Boolean
    Select Case viViewType
        Case cwiREGULAR_VIEW_A
            QueryView = True    '--this view is NOT optional--
        Case Else   '--unknown "view type", not supported--
            QueryView = False
    End Select
    glCallingNode = 1
End Function
'*******************************************************************
'  Description: This method is called when the CW Core attempts to open a "view" provided by this
'               OLE server.
'  Parameters:
'    viViewType - Contains the "type" of view being requested by the CW Core to open.
'    vlParenthWnd - Contains the hWnd of the client area window to be used by this OLE Server.
'    rhMainForm - If "viViewType" is a "regular view", upon return this will contain
'                 the window "handle" of the main "regular" form for this OLE Server.
'  Return Value:
'    Returns TRUE if the "view" was opened successfully.
'    Returns FALSE if an error occurred openning the "view".
'************************************************************************
Public Function OpenView(ByVal viViewType As Integer, _
                         ByVal vlParenthWnd As Long, _
                               rhMainForm As Long) As Boolean
    Dim sMsg As String
    On Error GoTo ErrHandler
    '---perform actions based on the "view type"---
    If viViewType = cwiREGULAR_VIEW_A Then
        '---set global flags---
        cwbFirstRegularResize_A = True
        cwbMainViewRunning_A = True
        '---------------------------------------------------------------------------------------
        '  If this OLE Server needs it's own menu items, here is a good place to add these
        '    menu items onto the CW Core's main menu.
        '  NOTE: This is OPTIONAL.  You do NOT have to create menu items. Thats why they are
        '        commented out.
        '---------------------------------------------------------------------------------------
        '---add a menu item "Template" on CW's main menu---
        'ghTemplate = gobjComm.AddPopupMenuToCW("&Template", cwlMyProgId)
        '---add a menu item "Ole Server Menu Item" to the menu "Template"---
        'glOleMenuItem = gobjComm.AddMenuItem("&Ole Server Menu Item", ghTemplate, cwlMyProgId, "Template menu item")
        
        glMenuRefresh_A = gobjComm.AddActionMenuItem("Refresh", "Refresh.bmp", cwlMyProgId_A, "Refresh Administrative Screen")

        
        
        '-----------------------------------------------------------------------------------------
        '  Reset the "parent window" for this OLE server's "main" regular form.  This is the main
        '    "trick" that makes the OLE server look like it is part of the CW Core.
        '  NOTE: This action invokes the main form's "Load" event, but the form is NOT shown.
        '        Make sure that you DO NOT perform a "Me.Show" in the form's "Load" event, as this
        '        will mess up the "SetParent" API call.  Make sure that you DO NOT "Show" the
        '        "main" regular form in this method, that will happen in the "ResizeView" method.
        '-----------------------------------------------------------------------------------------
        '---set parameter---
        '...This code required because CITA_Init could also
        '...Instantiate Imm_Tables and we only want 1 copy
        '...Similar code exists in the CloseDataObjects
        
        'SCR #20068
        'Begin:SCR #20521 - Immunizations - VB error 91 when log in as Imm_nurse - can't access module  Developer: Brian Mowbray 01/16/2002 03:51 PM
        gsMTF_ID = SayMTF_ID()
        If Len(gsMTF_ID) > 0 Then
          
          ' Begin:SCR #20852;   Developer: Brian Mowbray 05/28/2002 12:27 PM
          SetParent frmAdmin_Main.hwnd, vlParenthWnd
          
          rhMainForm = frmAdmin_Main.hwnd
          ' End: SCR #20852;
          
          '<< SCR #42805; 42809;   Developer: Brian Mowbray 10/02/2003 03:22 PM
          If Not gbITT Then
            If Imm_TablesCounter <= 0 And IMM_Tables Is Nothing Then
              Set IMM_Tables = New IMM_Tables
            End If
            Imm_TablesCounter = Imm_TablesCounter + 1
          End If
        Else
          sMsg = "Please log off the application to complete the initialization process for the immunization module. You are currently setup as Facility Name/NCID: " & gobjLogon.FacilityName & "/" & gobjLogon.FacilityNCID & " and Clinic Name/NCID: " & gobjLogon.NursingDivisionName & "/" & gobjLogon.NursingDivisionNCID & ", you will not be able to use the immunizations module until you have restarted the application." & vbCrLf & "Immunization Module Initialization completed!"
          MsgBxARMd sMsg, vbInformation, "Immunization Module Initialization"
          utLogError "basCITA_Lite:SayMTF_ID", sMsg, 999
          
          ' Begin:SCR #20852;   Developer: Brian Mowbray 05/28/2002 12:27 PM
          '---set return value---
          OpenView = False
          Screen.MousePointer = vbDefault
          Exit Function
          ' End: SCR #20852;
        End If
        'End: SCR #20521 - Immunizations - VB error 91 when log in as Imm_nurse - can't access module
        glCallingNode = 1
    Else    '--unsupported view "type"--
        '---set return value---
        OpenView = False
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    '---set return value---
    OpenView = True
    Screen.MousePointer = vbDefault
    Exit Function
ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Immunization.OpenView", "Immunization", vbCritical)
    '---set return value---
    OpenView = False
    Exit Function
    Resume
End Function
'****************************************************************
'  Description: This method is called when the CW Core closes a "view" provided by this OLE
'               Server.
'  Parameters:
'    viViewType - Contains the "type" of view that the CW Core is closing.
'************************************************************************
Public Sub CloseView(ByVal viViewType As Integer)
  '...T.Noren  08/17/01
  Dim CloseThemBoys As Boolean
    If viViewType = cwiALL_VIEWS_A Then
        If cwbMainViewRunning_A Then CloseThemBoys = True
    ElseIf viViewType = cwiREGULAR_VIEW_A Then
        If cwbMainViewRunning_A Then CloseThemBoys = True
    End If
    If CloseThemBoys Then
        '---close the "main view"---
        '---set global flags---
        cwbMainViewRunning_A = False
        '---this indicates that the CW Core is closing the "view", NOT this OLE Server---
        cwbCloseViewMyself_A = False
        '---close ALL forms that are associated with all "views"---
        Unload frmAdmin_Main
        ' T.Noren  11/01/01 make sure we set objects to nothing that are not being used anymore
        ' Following instruction was moved to Terminate code
        CloseDataObjects

    End If
End Sub
'*******************************************************************************************
'  Description: This method is called when the user has selected a menu item (from the CW
'               Core's main menu), which is "owned" by this OLE Server.
'               NOTE: This is an OPTIONAL method.
'  Parameters:
'    vlMenuId - Contains the menu id of the menu item the user selected.
'*******************************************************************************************
Public Sub MenuItemSelected(ByVal vlMenuId As Long)
    On Error GoTo ErrHandler
    '----------------------------------------------------------------------------------
    '  NOTE: You don't have to create menu items, these are optional, that is why
    '        the code is commented out.
    '----------------------------------------------------------------------------------
      Select Case vlMenuId 'Button.Key
        Case glMenuRefresh_A
          ' SCR 15934 - Brian Mowbray 9/12/2001
          SayStatus "Refreshing Admin Screen"
          frmAdmin_Main.AdminRefresh
    End Select
      
    Exit Sub
ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "CITA_Admin_Init.MenuItemSelected", "CHCSII_Immun", vbExclamation)
    Exit Sub
End Sub

'**********************************************************************
'  Description: This method gets called when a new patient has been selected from the CW Core.
'               NOTE: This is an OPTIONAL method.
'  Parameters:
'    none
'*************************************************************************
Public Sub PatientSelected()
  '<< Begin:SCR #24709; 25036;   Developer: Brian Mowbray 06/24/2002 03:17 PM
  If cwbMainViewRunning_A Then
    gbChangePatient = True
  Else
    gbChangePatient = False
  End If
'>> End: SCR #24709; 25036;
End Sub
'***********************************************************
'  Description: This method is called when a change has occurred in the "client area"
'               used by this OLE Server, resulting in a need to resize a particular "view".
'  Parameters:
'    viViewType - Contains the "type" of view being requested to resize by the CW Core.
'    vlLeft - Contains the "left" coordinate for the OLE server.
'    vlTop - Contains the "top" coordinate for the OLE server.
'    vlWidth - Contains the "width" for the OLE server.
'    vlHeight - Contains the "height" for the OLE server.
'**************************************************************
Public Sub ResizeView(ByVal viViewType As Integer, ByVal vlLeft As Long, ByVal vlTop As Long, ByVal vlWidth As Long, ByVal vlHeight As Long)
    Dim sMsg As String
    Dim bResult As Boolean
    Dim iCnt As Integer
    
    On Error GoTo ErrHandler
    '---reset global "client area" parameters---
    cwlLeft = vlLeft
    cwlTop = vlTop
    cwlWidth = vlWidth
    cwlHeight = vlHeight
    
    '---perform different actions based upon the "view type"---
    If viViewType = cwiREGULAR_VIEW_A Then
        If Not cwbFirstRegularResize_A Then
            '---------------------------------------------------------------------------------
            '  Reset the "WindowState" to "Normal", so the form can be resized.  This is
            '    needed because a window CANNOT be moved or resized if it is "minimized" or
            '    "maximized".
            '  NOTE: This generates a form "resize" event.
            '---------------------------------------------------------------------------------
            
              ' Begin:SCR #23107;   Developer: Brian Mowbray 04/17/2002 04:26 PM
              utCloseImmForms
              ' End: SCR #23107;

              frmAdmin_Main.WindowState = vbNormal
              
              '---hide the form to so that annyoing "screen blinking"
              '  will not be seen while controls are moved and resized---
              frmAdmin_Main.Visible = False
              
              '---resize form to fill entire "client area", NOTE: This generates a "resize" event---
              frmAdmin_Main.Move cwlLeft, cwlTop, cwlWidth, cwlHeight
              ' SCR 16337  T.Noren  09/10/01
              frmAdmin_Main.BringFrameIntoView
              
              ' Begin:SCR #23107;   Developer: Brian Mowbray 04/17/2002 04:26 PM
'              ' T.Noren 09/26/01
'              If Not gbInModalState Then
               glCallingNode = 1
               frmAdmin_Main.Visible = True
               'SCR 19761 12/17/2001 B.Mowbray
'               'frmAdmin_Main.Controls(frmAdmin_Main.sPreviousNode).ZOrder 0
'              End If
              ' End: SCR #23107;
              
        Else    '--first "regular" resize, need to perform "initailization" procedures--
            '---reset flag so this code will not be executed again---
          'Begin:SCR #20521 - Immunizations - VB error 91 when log in as Imm_nurse - can't access module  Developer: Brian Mowbray 01/16/2002 03:56 PM
          If Len(gsMTF_ID) = 0 Then
            CloseView cwiREGULAR_VIEW_A
          Else
            cwbFirstRegularResize_A = False
            '----------------------------------------------------------------------------
            '  NOTE: Here is where you "Show" the "Regular View" form for the FIRST time.
            '----------------------------------------------------------------------------
            ' SCR 16337  T.Noren  09/10/01
            frmAdmin_Main.BringFrameIntoView
            glCallingNode = 1
            frmAdmin_Main.Show
          End If
          'End: SCR #20521 - Immunizations - VB error 91 when log in as Imm_nurse - can't access module
       End If
    
    End If
    Exit Sub
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "CITA_Admin_Init.ResizeView", "CHCSII_Immun", vbExclamation)
    Exit Sub
    Resume
End Sub

Private Sub Class_Initialize()

    '--- COMMON TO MULTIPLE SCR-ARMs: MSolano 9/13/01 SCR-ARM-16110, 16130
    '--- Init routine should be called from first top level creatable coclass
    '--- It can be called multiple times, but will actually run only once.
    '--- Instrument all creatable classes in case dll's multiple entry
    '--- points are actually used
    Call ARM_InitAllTransactionsForImmunizations

    '---create shared object used to show VB errors, etc.---
    Set gobjShared = New CWShared
    '---set the help file for this application---
    App.HelpFile = gobjShared.CWHelpFile
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
  'These only get set to nothing when the class terminates - B Mowbray 11/12/2001
    Set gobjConfig = Nothing
    Set gobjComm = Nothing
    Set gobjLogon = Nothing
    Set gobjUser = Nothing
    Set gobjShared = Nothing
    Set gobjSQL = Nothing
    Set gobjDas = Nothing
    Set gobjPatient = Nothing
    Set PatientIMM = Nothing
    Set IMM_Tables = Nothing
'    CloseDataObjects
End Sub

'*******************************************************
'  Description: This method is used by the CW Core to send messages to this OLE Server.  The
'               messages can be from the CW Core or from another OLE Server.
'  Parameters:
'    viMessageType - Contains the "type" of message to process.
'    vsMessageText - Contains any "data" assoicated with this message.
'    vsSourceServerNCID - Contains the NCID of the OLE Server that "sent" the message.
'****************************************************************************
Public Sub message(ByVal viMessageType As Integer, _
                   ByVal vsMessageText As String, _
                   ByVal vsSourceServerNCID As String)
    On Error GoTo ErrHandler
    '-------------------------------------------------------------------------------
    '  Take corresponding actions based upon the message "type".
    '  NOTE: This method is required, but currently there are NO REQUIRED messages
    '        to process.
    '-------------------------------------------------------------------------------
    Select Case viMessageType
    
'        Case cwiOK_TO_CHANGE_PATIENT
'            gobjComm.message cwiOK_TO_CHANGE_PATIENT, "Y", vsSourceServerNCID, NCID_CITA_ADMIN_APPLICATION
'
        Case cwiOK_TO_CLOSE_VIEW
            gobjComm.message cwiOK_TO_CLOSE_VIEW, "Y", vsSourceServerNCID, NCID_CITA_ADMIN_APPLICATION
            
        Case cwiBarCode
            If utIsLoaded("frmBarCode") Then
              frmBarCode.SetBC_Data = vsMessageText
            End If
        Case Else   '--do nothing--
        
    End Select
    
    Exit Sub
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "CITA_Admin_Init.Message", "CHCSII_Immun", vbExclamation)
    Exit Sub
End Sub

'***********************************************************************
'  Description: This method is called when the CW Core creates an "instance" of this OLE Server.  It
'               is used to create and initialize global variables that are shared across all
'               "views" supported by this OLE Server, i.e. variables that only need to be
'               initialized ONCE, upon object creation.
'  Parameters:
'    vlMyProgId - Contains the "ProgId" that the CW Core has assigned to this OLE Server.
'    robjComm - Contains a communications object that is used to handle all communication
'               between this OLE Server and the CW Core.
'    robjLogonEnvironment - Contains the logon environment object.  Contains information such
'                           as who logged on, where they logged on, etc.
'    robjConfig - Contains the configuration object that is used to get and set configuration
'                 information for a specific application.
'    robjPatient - Contains the patient object containing the currently selected patient.
'***************************************************************************
Public Sub Initialize(ByVal vlMyProgId As Long, _
                            robjComm As Object, _
                            robjLogonEnvironment As Object, _
                            robjConfig As Object, _
                            robjPatient As Object, robjUser As ICHCSII_User)
    On Error GoTo ErrHandler
    'Saves which node called immunizations, 1 = CITA Admin Module, 2 = CITA Immunization Module
    
    '---save the "ProgId" assigned to me by the CW Core---
    cwlMyProgId_A = vlMyProgId
    '---create references to passed in objects---
    Set gobjComm = robjComm
    Set gobjLogon = robjLogonEnvironment
    Set gobjUser = robjUser
    Set gobjPatient = robjPatient
    Dim cx As Collection
    Set cx = gobjUser.GetAttributeNames
    Dim xRoles As ICHCSII_Roles
    Dim xrole As Integer
    Dim xResources As ICHCSII_Resources
    Dim xResource As Integer
    Dim sMsg As String
    Set gobjConfig = robjConfig
    Dim X As ICHCSII_Role
    Set xRoles = gobjUser.Roles
    Set xResources = gobjUser.Resources
    '---set global flags---
    For xrole = 1 To xRoles.Count
     Debug.Print xRoles.Item(xrole).Name
    Next xrole
    For xResource = 1 To xResources.Count
     Debug.Print xResources.Item(xResource).Name
    Next xResource
    cwbMainViewRunning_A = False

    Set gobjSQL = SelectCHCSII_DAS()
    
    SetADODBProvider
    
     
    'Set the DMIS ID
'  SCR 20852 1/25/2002 Brian Mowbray
    gsMTF_ID = SayMTF_ID()
    

    

    Exit Sub
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "CITA_Admin_Init.Initialize", "CHCSII_Immun", vbExclamation)
    Exit Sub
    Resume
End Sub
'**************************************************************
'  Description: This method is called by a client when they want to modify the "properties", or the
'               configuration of this OLE Server.
'               NOTE: This is an OPTIONAL method.
'  Parameters:
'    robjLogonEnvironment - Contains the logon environment object.
'    robjConfig - Contains the configuration object that is used to get and set configuration
'                 information for a specific application.
'    vsOwnerNCID - Contains the NCID of the "owner" to use in storing modifications to the configuration.
'    vbCalledByDesigner - TRUE if this method is being called by the CW Designer, FALSE otherwise.
'****************************************************************
Public Sub Properties(robjLogonEnvironment As Object, _
                      robjConfig As Object, _
                      ByVal vsOwnerNCID As String, _
                      ByVal vbCalledByDesigner As Boolean)
    On Error GoTo ErrHandler
    '*****************************************************************
    '  TO DO:
    '    Add code here to invoke this OLE Server's properties page.
    '*****************************************************************
  Exit Sub
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "CITA_Admin_Init.Properties", "CHCSII_Immun", vbExclamation)
    Exit Sub
End Sub

'***********************************************************************
'  Description: This method is called after the CW Core creates an "instance" of this OLE Server.  It
'               is used to create and initialize global variables that are shared across all
'               "views" supported by this OLE Server, i.e. variables that only need to be
'               initialized ONCE, upon object creation.
'  Parameters:
'    vlMyProgId - Contains the "ProgId" that the CW Core has assigned to this OLE Server.
'    robjComm - Contains a communications object that is used to handle all communication
'               between this OLE Server and the CW Core.
'    robjLogonEnvironment - Contains the logon environment object.  Contains information such
'                           as who logged on, where they logged on, etc.
'    robjConfig - Contains the configuration object that is used to get and set configuration
'                 information for a specific application.
'    robjPatient - Contains the patient object containing the currently selected patient.
'    robjUser - Security User object
'***************************************************************************
Private Sub ICHCSII_CoreAPI2_Initialize2(ByVal vlMyProgId As Long, _
                                               robjComm As Object, _
                                               robjLogonEnvironment As Object, _
                                               robjConfig As Object, _
                                               robjPatient As Object, _
                                               robjUser As Object)
    Set gobjUser = robjUser
    ' SCR 16643 T.Noren 11/01/01
'    Call Initialize(vlMyProgId, robjComm, robjLogonEnvironment, robjConfig, robjPatient)
    Call Initialize(vlMyProgId, robjComm, robjLogonEnvironment, robjConfig, robjPatient, robjUser)
End Sub

'****************************************************************
'  Description: This method is called when the CW Core closes a "view" provided by this OLE
'               Server.  It differs from the CloseView method in that the child module is
'               is permitted to say 'No' to Core.
'  Parameters:
'    viViewType - Contains the "type" of view that the CW Core is closing.
'************************************************************************
Private Function ICHCSII_CoreAPI3_CloseView2(ByVal viViewType As Integer) As Boolean
  Call CloseView(viViewType)
  ICHCSII_CoreAPI3_CloseView2 = True
End Function



