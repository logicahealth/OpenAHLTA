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
Attribute VB_Name = "HealthHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Implements ICHCSII_CoreAPI2
Implements ICHCSII_CoreAPI3

Public mobjSQL           As ICHCSII_SQL.ISqlOps
Public mobjproperties    As clsProperties
Private objSumOLEModules As clsSumOLEModules
Private mConnection As CHCSII_CONN.Conn


'**********************************************************
'  V A R I A B L E S     S E C T I O N
'**********************************************************
'---variables used by OLE servers for "moving, sizing" and communicating with CW---
Private hhlHeight                         As Long
Private hhlWidth                          As Long
Private hhlLeft                           As Long
Private hhlTop                            As Long
Public hhlMyProgId                        As Long
Public hhbCloseViewMyself                 As Boolean
Private hhbFirstRegularResize             As Boolean
Private hhbFirstPatientSummaryResize      As Boolean
Private hhbFirstClinicianSummaryResize    As Boolean
Private hhbFirstEncounterSummaryResize    As Boolean
Private hhobjCWComm                       As MMMHISComm.Comm

Private hhobjLogonEnviron                 As Object
Private hhobjConfig                       As Object
Private hhobjPatient                      As Object
Private hhObjUser                         As Object
Private hhbMainViewRunning                As Boolean
Private hhbPatientSummaryViewRunning      As Boolean
Private hhbPrintPreview                   As Boolean

Public Property Get CWComm() As Object:       Set CWComm = hhobjCWComm:             End Property
Public Property Get LogonEnviron() As Object: Set LogonEnviron = hhobjLogonEnviron: End Property
Public Property Get Config() As Object:       Set Config = hhobjConfig:             End Property
Public Property Get Patient() As Object:      Set Patient = hhobjPatient:           End Property


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
            QueryView = True
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
Public Function OpenView(ByVal viViewType As Integer, ByVal vlParenthWnd As Long, rhMainForm As Long) As Boolean
    On Error GoTo errhandler
    
    If viViewType <> cwiREGULAR_VIEW Then
        GoTo ExitFail
    End If
    
    '---set public flags---
    hhbFirstRegularResize = True
    hhbMainViewRunning = True

    '---add a menu item "Template" on CW's main menu---
    'ghTemplate = cwobjCWComm.AddPopupMenuToCW("&Template", cwlMyProgId)
    '---add a menu item "Ole Server Menu Item" to the menu "Template"---
    'glOleMenuItem = cwobjCWComm.AddMenuItem("&Ole Server Menu Item", ghTemplate, cwlMyProgId, "Template menu item")
    
    SetParent frmHH.hWnd, vlParenthWnd
    rhMainForm = frmHH.hWnd

    OpenView = True
    Exit Function
    
errhandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "HealthHistory.OpenView", "HealthHistory", vbCritical)
    
ExitFail:
    OpenView = False
    Exit Function
End Function

'*******************************************************************************************
'  Description: This method is called when the CW Core closes a "view" provided by this OLE
'               Server.
'
'  Parameters:
'    viViewType - Contains the "type" of view that the CW Core is closing.
'*******************************************************************************************
Public Sub CloseView(ByVal viViewType As Integer)
        
    If hhbMainViewRunning Then
        Call objSumOLEModules.ShutDownSummaryOLEServers
    
        hhbMainViewRunning = False
        hhbCloseViewMyself = False
        Unload frmHH
    End If
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
    On Error GoTo errhandler
    
    Exit Sub

errhandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "HealthHistory.MenuItemSelected", "HealthHistory", vbExclamation)
    
    Exit Sub
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
    '---update the all running "views" with the new patient---
    If hhbMainViewRunning Then
        '****************************************************************************
        '  TO DO:
        '    Add code here to update the data in all "running" views with this new
        '    patient.
        '****************************************************************************
    
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
'*******************************************************************************************
Public Sub ResizeView(ByVal viViewType As Integer, _
                      ByVal vlLeft As Long, _
                      ByVal vlTop As Long, _
                      ByVal vlWidth As Long, _
                      ByVal vlHeight As Long)
    On Error GoTo errhandler

    '---reset public "client area" parameters---
    hhlLeft = vlLeft
    hhlTop = vlTop
    hhlWidth = vlWidth
    hhlHeight = vlHeight
    
    '---perform different actions based upon the "view type"---
    If viViewType = cwiREGULAR_VIEW Then
        If Not hhbFirstRegularResize Then
            '---------------------------------------------------------------------------------
            '  Reset the "WindowState" to "Normal", so the form can be resized.  This is
            '    needed because a window CANNOT be moved or resized if it is "minimized" or
            '    "maximized".
            '
            '  NOTE: This generates a form "resize" event.
            '---------------------------------------------------------------------------------
            frmHH.WindowState = vbNormal
            frmHH.Visible = False
            frmHH.Move hhlLeft, hhlTop, hhlWidth, hhlHeight
            frmHH.Visible = True
        Else    '--first "regular" resize, need to perform "initailization" procedures--
            '---reset flag so this code will not be executed again---
            hhbFirstRegularResize = False
            objSumOLEModules.RetrievOLEServerInfo
            Set frmHH.SumOLEModules = objSumOLEModules
            Call frmHH.Display(hhlLeft, hhlTop, hhlWidth, hhlHeight)

        End If
    End If
    
    Exit Sub
    
errhandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "HealthHistory.ResizeView", "HealthHistory", vbExclamation)
    
    Exit Sub
End Sub

Private Sub Class_Initialize()
    '---create shared object used to show VB errors, etc.---
    Set gobjShared = New CWShared
    Set mobjproperties = New clsProperties
    Set mobjproperties.Parent = Me
    Set objSumOLEModules = New clsSumOLEModules
    Set objSumOLEModules.Parent = Me
    
    Set mConnection = New CHCSII_CONN.Conn
    Set mobjSQL = mConnection.CHCSII_SQLOPS
    
    '---set the help file for this application---
    App.HelpFile = gobjShared.CWHelpFile
    
End Sub

Private Sub Class_Terminate()
    '---release reference to objects---
    Set gobjShared = Nothing
    Set hhobjCWComm = Nothing
    Set hhobjLogonEnviron = Nothing
    Set hhobjConfig = Nothing
    Set hhobjPatient = Nothing
    
    Set mobjSQL = Nothing
    Set mConnection = Nothing

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
    On Error GoTo errhandler
    
    Dim bPrintPreview As Boolean
    Dim objOLEServer As clsSumOLEModule
    Dim objAbout As CWAbout
    
    Select Case viMessageType
        Case cwiPROPERTIES
            Call ConfigSummaryOLEServers(frmHH.hWnd, frmHH.ScaleWidth, frmHH.ScaleHeight)
            
        Case cwiPRINT_REGULAR_VIEW, cwiPRINT_PREVIEW
            bPrintPreview = (viMessageType = cwiPRINT_PREVIEW)
            
            Call objSumOLEModules.PrintSummaryOLEServers(bPrintPreview)
        Case cwiREFRESH_DATA, cwiFOREGROUND_APPLICATION '141904
            
            Screen.MousePointer = vbHourglass
            '------------------------------------------------------------------
            '  Broadcast a "refresh data" message to all running OLE Servers.
            '------------------------------------------------------------------
            For Each objOLEServer In objSumOLEModules
'<< 6.4.1 WAS Look for non-loaded status
                If Not objOLEServer.Module Is Nothing Then
                    objOLEServer.Module.Message cwiREFRESH_DATA, "", NCID_CW_APPLICATION
                End If
            Next objOLEServer
        
            Screen.MousePointer = vbDefault
    
        Case cwiOK_TO_CHANGE_PATIENT
            If vsSourceServerNCID <> NCID_CW_APPLICATION Then
                On Error Resume Next
                    objSumOLEModules.Item(vsSourceServerNCID).sOkTochangePatient = vsMessageText
                On Error GoTo errhandler
            Else
                hhobjCWComm.Message cwiOK_TO_CHANGE_PATIENT, "Y", vsSourceServerNCID, NCID_PATIENT_SUMMARY_APPLICATION
            End If
        
        Case cwiOK_TO_CLOSE_VIEW
            If vsSourceServerNCID <> NCID_CW_APPLICATION Then
                On Error Resume Next
                    objSumOLEModules.Item(vsSourceServerNCID).sOkTocloseView = vsMessageText
                On Error GoTo errhandler
            Else
                hhobjCWComm.Message cwiOK_TO_CLOSE_VIEW, "Y", vsSourceServerNCID, NCID_PATIENT_SUMMARY_APPLICATION
            End If
        '<SCR 20954 CC>
        Case cwiABOUT

            Set objAbout = New CWAbout
            
            '---inform CW that I am showing a modal form---
            hhobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_PATIENT_SUMMARY_APPLICATION
            
            objAbout.Display App, frmHH
            
            '---inform CW that I am done showing a modal form---
            hhobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_PATIENT_SUMMARY_APPLICATION
            
            Set objAbout = Nothing
        '</SCR 20954 CC>
        Case Else   '--do nothing--
    End Select

    Exit Sub
    
errhandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "HealthHistory.Message", "HealthHistory", vbExclamation)
    
    Exit Sub
End Sub



'****************************************************************************************************
'  Description: This method is called when the CW Core creates an "instance" of this OLE Server.  It
'               is used to create and initialize public variables that are shared across all
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
    On Error GoTo errhandler
    
    '---save the "ProgId" assigned to me by the CW Core---
    hhlMyProgId = vlMyProgId
    
    '---create references to passed in objects---
    Set hhobjCWComm = robjComm
    Set hhobjLogonEnviron = robjLogonEnvironment
    Set hhobjConfig = robjConfig
    Set hhobjPatient = robjPatient
    
    '---set public flags---
    hhbMainViewRunning = False
       
    Exit Sub

errhandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "HealthHistory.Initialize", "HealthHistory", vbExclamation)
    
    Exit Sub
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
    On Error GoTo errhandler
    
    '*****************************************************************
    '  TO DO:
    '    Add code here to invoke this OLE Server's properties page.
    '*****************************************************************
    Call ConfigSummaryOLEServers(frmHH.hWnd, frmHH.ScaleWidth, frmHH.ScaleHeight)

    
    Exit Sub
    
errhandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "HealthHistory.Properties", "HealthHistory", vbExclamation)
    
    Exit Sub
End Sub

Private Sub ICHCSII_CoreAPI2_Initialize2(ByVal vlMyProgId As Long, _
                                               robjComm As Object, _
                                               robjLogonEnvironment As Object, _
                                               robjConfig As Object, _
                                               robjPatient As Object, _
                                               robjUser As Object)
    Set hhObjUser = robjUser
    Call Initialize(vlMyProgId, robjComm, robjLogonEnvironment, robjConfig, robjPatient)
End Sub

Private Function ICHCSII_CoreAPI3_CloseView2(ByVal viViewType As Integer) As Boolean
'- I only have one view type

  ICHCSII_CoreAPI3_CloseView2 = True

  If hhbMainViewRunning Then
    If objSumOLEModules.ShutDownSummaryOLEServers Then
        hhbMainViewRunning = False
        hhbCloseViewMyself = False
        Unload frmHH
    Else
        ICHCSII_CoreAPI3_CloseView2 = False
    End If
  End If

End Function


'*******************************************************************************************
'  Description: This subroutine allows the user to "configure" the OLE servers for the
'               specified summary view "type".
'
'  Parameters:
'    viSummaryViewType - Contains the "type" of summary view.
'    vParenthWnd - Contains the hWnd of the parent window.
'    vlScaleWidth - Contains the "scale width" of the parent window.
'    vlScaleHeight - Contains the "scale height" of the parent window.
'*******************************************************************************************
Public Sub ConfigSummaryOLEServers(ByVal vParenthWnd As Long, _
                                   ByVal vlScaleWidth As Long, _
                                   ByVal vlScaleHeight As Long)
    Dim sSummarySection  As String
    Dim sSupportSection  As String
    Dim sSummaryTitle    As String
    Dim oSumMod          As clsSumOLEModule
    Dim bOrigOnSummary   As Boolean
    Dim mcol             As Collection
    Dim oNewSumOLEModuls As clsSumOLEModules
    
    Set oNewSumOLEModuls = New clsSumOLEModules
    Dim lLeft As Long
    Dim lTop As Long
    Dim lWidth As Long
    Dim lHeight As Long
    Dim objoleserver3 As ICHCSII_CoreAPI3

    On Error GoTo errhandler

    '<SCR 23148 CC - Diable the "X" button on core>
    hhobjCWComm.Message cwiSHOWING_MODAL_FORM, "Y", NCID_CW_APPLICATION, NCID_PATIENT_SUMMARY_APPLICATION
    
    If mobjproperties Is Nothing Then
        Set mobjproperties = New clsProperties
        Set mobjproperties.Parent = Me
    End If
    Set mobjproperties.SumOLEModules = objSumOLEModules
    
'- save the current state of the modules
    Set mcol = New Collection
    For Each oSumMod In objSumOLEModules
        mcol.Add oSumMod.bOnSummary, oSumMod.sApplicationNCID
    Next oSumMod

'- display 'properties' form
    If mobjproperties.DisplaySummaryConfig = vbCancel Then
        GoTo Cleanup
    End If
    
'-  Else, repaint the main display to show the new config
    '---create a summary configuration collection

'    ---if the user changed the summary, reconfigure it---
        'Need to disable the "close" botton
        Screen.MousePointer = vbHourglass

        '-------------------------------------------------------------------------------------
        '  Loop through all OLE servers in public "gatOLEServerInfo()" and retrieve the
        '    modified "configuration" information the user justed changed.  Also, you need
        '    to "shut down" all OLE servers the user DOES NOT want on the specified summary
        '    view anymore.
        '-------------------------------------------------------------------------------------
    
    For Each oSumMod In objSumOLEModules
        On Error GoTo errhandler    '--renable original error handler--
        With oSumMod
            '-----------------------------------------------------------------------------
            '  Check to see if this OLE server was on the specified summary view, and
            '    now the user took it off.  In this case, you need to "shut down" the
            '    OLE server.
            '-----------------------------------------------------------------------------
            '- find this one in the local collection
            bOrigOnSummary = mcol(.sApplicationNCID)
            '-- if was displayed and not now, close the view
                If (bOrigOnSummary And Not .bOnSummary) Then
                    .CloseModule
                End If
            '-- if was displayed and also now, resize/reposition
                If (bOrigOnSummary And .bOnSummary) Then
                    .Resize vlScaleWidth, vlScaleHeight
                End If
            '-- if was not displayed and is now, launch
                If (Not bOrigOnSummary And .bOnSummary) Then
                    Call objSumOLEModules.OpenSummaryOLEServer(oSumMod, vParenthWnd, vlScaleWidth, vlScaleHeight)
                End If
            '-- if was not displayed and is not now, no action
                'noop
            
        End With
    Next oSumMod

    hhobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "Y", NCID_CW_APPLICATION, NCID_PATIENT_SUMMARY_APPLICATION  ''<SCR 23148 CC>
    
    GoTo Cleanup
    
errhandler:
    '---show the VB error---
    
    gobjShared.ShowVBError Err.Number, Err.Description, "gConfigSummaryOLEServers", App.Title, vbExclamation
    
Cleanup:
    ''<SCR 23148 CC> Make sure to return to non-modal status
    If Not hhobjCWComm Is Nothing Then
        hhobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "Y", NCID_CW_APPLICATION, NCID_PATIENT_SUMMARY_APPLICATION
    End If
    
    Set mcol = Nothing
    Screen.MousePointer = vbDefault
    hhobjCWComm.SetStatusText ""
    Exit Sub
End Sub



