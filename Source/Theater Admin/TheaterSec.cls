VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TheaterSec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements ICHCSII_CoreAPI2

Public gobjLogon As Logon


'-------------------------------------------------------------------------------------------------
'  Description: This method is called when CW "creates" an instance of this OLE Server.  It
'               is used to create and initialize global variables that are shared across all
'               "views" supported by this OLE Server, i.e. variables that only need to be
'               initialized ONCE, upon object creation.
'
'  Parameters:
'    vlMyProgId - Contains the "ProgId" that CW has assigned to this OLE Server.
'    robjComm - Contains a communications object that is used to handle all communication
'               between this OLE Server and the CW core.
'    robjLogonEnviron - Contains the logon environment object.
'    robjConfig - Contains the configuration object used for getting and setting configuration
'                 information for a specific application.
'    robjPatient - Contains the patient object of the currently selected patient.
'-------------------------------------------------------------------------------------------------
Public Sub Initialize(ByVal vlMyProgId As Long, robjComm As Object, robjLogonEnviron As Logon, robjConfig As Object, robjPatient As Object)
    Dim sBuf As String
    
    On Error GoTo errhandler
    
    '---save the "ProgId" assigned to me by CW---
    cwlMyProgId = vlMyProgId
    
    '---create pointers to objects---
    Set cwobjCWComm = robjComm
    Set cwobjLogonEnviron = robjLogonEnviron
    Set cwobjConfig = robjConfig
    Set cwobjPatient = robjPatient
    Set gobjLogon = robjLogonEnviron
   
    glPrevPat = cwobjPatient.UnitNumber
    
    gsUser = cwobjLogonEnviron.UserNCID
    gsFacility = cwobjLogonEnviron.FacilityNcid
        
    cwbMainViewRunning = False

    If gobjShared.UseLocalDB Then
        gbUseLocalDB = True
        Set gobjGemsSQL = New GEMS_SQL.SqlOps
        Set gobjGemsSQL = gobjGemsSQL
        Set gobjGemsConcept = New GEMS_ConceptCtrl
        Set gobjGemsDas = New GEMS_DAS
    End If
        
    Exit Sub

errhandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "TheaterSecurity.Initialize", "Vital Signs", vbExclamation)
    Exit Sub
End Sub


'-------------------------------------------------------------------------------------------------
'  Function Name: OpenView
'
'  Description: This method is called when CW attempts to open a "view" provided by this
'               OLE server.
'
'  Parameters:
'    viViewType - Contains the "type" of view being requested by CW to open.
'    vlParenthWnd - Contains the hWnd of the client area window to be used by this OLE
'                   server.
'    rhMainForm - If "viViewType" is a "regular view", upon return this will contain
'                 the window "handle" of the main "regular" form for this OLE Server.
'
'  Return Value:
'    Returns TRUE if openning the "view" was successfully.
'    Returns FALSE if an error occurred openning the "view".
'-------------------------------------------------------------------------------------------------
Public Function OpenView(ByVal viViewType As Integer, ByVal vlParenthWnd As Long, rhMainForm As Long) As Boolean
    On Error GoTo errhandler
    '---perform actions based on the "view type"---
        
    gViewType = viViewType
    If viViewType = cwiREGULAR_VIEW Then
        
        cwbFirstRegularResize = True
        cwbMainViewRunning = True
            
        If Not gbReadAccess Then
           MsgBox "Insufficient privileges to access module.", vbInformation, "Security"
           Screen.MousePointer = vbDefault
           Exit Function
        End If
        
        '----------------------------------------------------------------------
        '  Reset the "parent window" for this OLE server's "main" regular form.
        '
        '  NOTE: This invokes the form's "Load" event, but the form is NOT
        '        shown.  Make sure that you DO NOT perform a "Me.Show" in the
        '        form's "Load" event.
        '----------------------------------------------------------------------
        SetParent frmMain.hWnd, vlParenthWnd
        rhMainForm = frmMain.hWnd
    End If
        
    OpenView = True
    Screen.MousePointer = vbDefault
    Exit Function
    
errhandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "TheaterSecurity.OpenView", "Vital Signs", vbCritical)
    OpenView = False
    Screen.MousePointer = vbDefault

End Function

'-------------------------------------------------------------------------------------------------
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
'-------------------------------------------------------------------------------------------------
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
        Case Else   '--unknown "view type", not supported--
            QueryView = False
    End Select
End Function

Private Sub Class_Initialize()
    Set gobjShared = New CWShared
    Set oSQL = New GEMS_SQL.SqlOps
    If gobjShared.UseLocalDB Then
        gbUseLocalDB = True
        Set gobjGemsSQL = New GEMS_SQL.SqlOps
        Set gobjGemsConcept = New GEMS_ConceptCtrl
        Set gobjGemsDas = New GEMS_DAS
    End If
    
End Sub

Private Sub Class_Terminate()
    Set gobjShared = Nothing
    Set gobjGemsSQL = Nothing
    Set gobjGemsConcept = Nothing
    Set gobjGemsDas = Nothing
    Set gobjLogon = Nothing
    Set cwobjCWComm = Nothing
    Set cwobjLogonEnviron = Nothing
    Set cwobjConfig = Nothing
    Set cwobjPatient = Nothing
End Sub

Public Sub ICHCSII_CoreAPI2_Initialize2(ByVal vlMyProgId As Long, robjComm As Object, robjLogonEnvironment As Object, robjConfig As Object, robjPatient As Object, robjUser As Object)
    
    Initialize vlMyProgId, robjComm, robjLogonEnvironment, robjConfig, robjPatient
    Set gcwobjUser = robjUser
    
    'BUG read access check must be performed twice to get correct attribute
    gbReadAccess = gcwobjUser.HasPrivilegeEx(Priv_Vital_Signs, Priv_Read)
    gbReadAccess = gcwobjUser.HasPrivilegeEx(Priv_Vital_Signs, Priv_Read)
    
    gbDeleteAccess = gcwobjUser.HasPrivilegeEx(Priv_Vital_Signs, Priv_Delete)
    gbWriteAccess = gcwobjUser.HasPrivilegeEx(Priv_Vital_Signs, Priv_Write)
    
End Sub

'-------------------------------------------------------------------------------------------------
'  Sub Name: CloseView
'
'  Description: This method is called when CW closes a "view" provided by this OLE server.
'
'  Parameters:
'    viViewType - Contains the "type" of view that CW is closing.
'-------------------------------------------------------------------------------------------------
Public Sub CloseView(ByVal viViewType As Integer)
    On Error Resume Next

    gViewType = viViewType
    If viViewType = cwiALL_VIEWS Then
        '---this indicates that CW is closing the "view", NOT the OLE Server---
        cwbCloseViewMyself = False
        
        '---unload all forms---
        If cwbMainViewRunning Then
            Unload frmMain
        End If
        
        Set frmMain = Nothing
        '---set global flags---
        cwbMainViewRunning = False
        cwbPatientSummaryViewRunning = False
        cwbEncounterSummaryViewRunning = False
    ElseIf viViewType = cwiREGULAR_VIEW Then
        '---this indicates that CW is closing the "view", NOT the OLE Server---
        cwbCloseViewMyself = False
        
        If cwbMainViewRunning Then
           Unload frmMain
        End If
        
        Set frmMain = Nothing
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
    On Error GoTo errhandler

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
            frmMain.WindowState = vbNormal
            
            '---hide the form to so that annyoing "screen blinking" will not be seen while controls are moved and resized---
            frmMain.Visible = False
            
            '---resize form to fill entire "client area", NOTE: This generates a "resize" event---
            frmMain.Move cwlLeft, cwlTop, cwlWidth, cwlHeight
                
            frmMain.Visible = True
        Else    '--first "regular" resize, need to perform "initailization" procedures--
            '---reset flag so this code will not be executed again---
            cwbFirstRegularResize = False
        
            '----------------------------------------------------------------------------
            '  NOTE: Here is where you "Show" the "Regular View" form for the FIRST time.
            '----------------------------------------------------------------------------
            Set frmMain.mobjParent = Me
            frmMain.Display
        End If
    End If
    
    Exit Sub
    
errhandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "TheaterSecurity.ResizeView", "Template", vbExclamation)
    Exit Sub
End Sub

'*******************************************************************************************
'  Description: This method is called when the user has selected a menu item (from the CW
'               Core's main menu), which is "owned" by this OLE Server.
'               NOTE: This is an OPTIONAL method.
'  Parameters:
'    vlMenuId - Contains the menu id of the menu item the user selected.
'*******************************************************************************************
Public Sub MenuItemSelected(ByVal vlMenuId As Long)
    On Error GoTo errhandler
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
'            '*****************************************************************
'        Case Else   '--do nothing--
'    End Select
    Exit Sub
errhandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "TheaterSecurity.MenuItemSelected", "Template", vbExclamation)
    Exit Sub
End Sub

'*******************************************************
'  Description: This method is used by the CW Core to send messages to this OLE Server.  The
'               messages can be from the CW Core or from another OLE Server.
'  Parameters:
'    viMessageType - Contains the "type" of message to process.
'    vsMessageText - Contains any "data" assoicated with this message.
'    vsSourceServerNCID - Contains the NCID of the OLE Server that "sent" the message.
'****************************************************************************
Public Sub Message(ByVal viMessageType As Integer, _
                   ByVal vsMessageText As String, _
                   ByVal vsSourceServerNCID As String)
    On Error GoTo errhandler
    '-------------------------------------------------------------------------------
    '  Take corresponding actions based upon the message "type".
    '  NOTE: This method is required, but currently there are NO REQUIRED messages
    '        to process.
    '-------------------------------------------------------------------------------
    Select Case viMessageType
        Case cwiOK_TO_CHANGE_PATIENT
            cwobjCWComm.Message cwiOK_TO_CHANGE_PATIENT, "Y", vsSourceServerNCID, NCID_THEATER_SECURITY_APPLICATION
        Case cwiOK_TO_CLOSE_VIEW
            cwobjCWComm.Message cwiOK_TO_CLOSE_VIEW, "Y", vsSourceServerNCID, NCID_THEATER_SECURITY_APPLICATION
        Case Else   '--do nothing--
    End Select
    Exit Sub
errhandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "TheaterSecurity.Message", "Template", vbExclamation)
    Exit Sub
End Sub



