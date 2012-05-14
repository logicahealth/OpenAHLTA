VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataManagerOps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ICHCSII_CoreAPI2
Implements ICHCSII_CoreAPI3

'DBMCHANGE
Public nCurrentPatientCnt As Long
Public bResendTMIPMessages As Boolean

Private mlMenuRegen As Long
Private mlMenuDataRet As Long
Private mlMenuImport As Long
Private mlMenuUnverified As Long
Private mlMenuPrint As Long
Private mlMenuPurgeAuditLog As Long

' Begin:SCR #30622;   Developer: Mag 12/18/2002 04:10 PM
Public Sub displayPatientArray(oArrayPatient As MMMHISPatient.Patient)
    Dim idx As Long
    Dim objPatient As MMMHISPatient.Patient 'Temporary
    Dim sTemp As String
    Dim nRow As Long
    
    On Error GoTo ErrHandler
    
    If IsEmpty(oArrayPatient.ResultsArray) Then
        Exit Sub
    End If
    
    '--- SCR 71089   Sherry Wang    5/20/2005
    gfrmRegularView.grdSearch.ColHidden(TXT_UIC) = True
    gfrmRegularView.grdSearch.ColHidden(TXT_SOURCE) = True
    
    For idx = 1 To UBound(oArrayPatient.ResultsArray) + 1
        'Check to see if Patient already loaded
        If Not checkDuplications(CLng(oArrayPatient.ResultsArray(idx - 1, TXT_UNIT_NUMBER))) Then
            'If Patient does not exist in grid then
            With gfrmRegularView.grdSearch

                .AddItem ""
                nRow = .Rows - 1
                .TextMatrix(nRow, TXT_NAME) = oArrayPatient.ResultsArray(idx - 1, TXT_NAME)
                .TextMatrix(nRow, TXT_SSN) = Format(oArrayPatient.ResultsArray(idx - 1, TXT_SSN), "000-00-0000")

                ' Begin:SCR #22589;   Developer: MAG 02/25/2003 03:33 PM
                sTemp = oArrayPatient.ResultsArray(idx - 1, TXT_FMP_SSN)
                sTemp = Replace(sTemp, "-", "")
                sTemp = Replace(sTemp, "/", "")

                If Len(sTemp) > 2 Then
                    .TextMatrix(nRow, TXT_FMP_SSN) = Left(sTemp, 2) & "/" & Format(Right(sTemp, Len(sTemp) - 2), "000-00-0000")
                End If
                ' End: SCR #22589;

                If oArrayPatient.ResultsArray(idx - 1, TXT_DOB) <> cwdNULL_DATE Then
                    .TextMatrix(nRow, TXT_DOB) = Format$(oArrayPatient.ResultsArray(idx - 1, TXT_DOB), "mm/dd/yyyy")
                End If
                .TextMatrix(nRow, TXT_SEX) = oArrayPatient.ResultsArray(idx - 1, TXT_SEX)
                .TextMatrix(nRow, TXT_HOMEPHONE) = oArrayPatient.ResultsArray(idx - 1, TXT_HOMEPHONE)
                .TextMatrix(nRow, TXT_WORKPHONE) = oArrayPatient.ResultsArray(idx - 1, TXT_WORKPHONE)
                .TextMatrix(nRow, TXT_ADDRESS) = oArrayPatient.ResultsArray(idx - 1, TXT_ADDRESS)
                .TextMatrix(nRow, TXT_UNIT_NUMBER) = oArrayPatient.ResultsArray(idx - 1, TXT_UNIT_NUMBER)
                
                '--- SCR 71089   Sherry Wang    5/20/2005
                If LenB(oArrayPatient.ResultsArray(idx - 1, TXT_UIC)) > 0 Then
                    .ColHidden(TXT_UIC) = False
                    .TextMatrix(nRow, TXT_UIC) = oArrayPatient.ResultsArray(idx - 1, TXT_UIC)
                End If

                .IsSelected(nRow) = True


            End With
        End If
    Next idx
    
    '****************************************
    'Inpatient Phase II
    'SCR 55740
    'B. Fote (16-Nov-2004)
    If gfrmRegularView.grdSearch.Rows > 1 Then
      gfrmRegularView.cmdExecute.Enabled = True
      gfrmRegularView.cmdRemove.Enabled = True
    End If
    '****************************************

    Exit Sub
    
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "DataManagerOps.displayPatientArray", "DataManager", vbCritical)
End Sub
' End: SCR #30622;
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
        Case cw32Global.cwiREGULAR_VIEW
            QueryView = True    '--this view is NOT optional--
        Case cw32Global.cwiPATIENT_SUMMARY_VIEW
            QueryView = False
        Case cw32Global.cwiCLINICIAN_SUMMARY_VIEW
            QueryView = False
        Case cw32Global.cwiENCOUNTER_SUMMARY_VIEW
            QueryView = False
        Case cw32Global.cwiPROPERTY_VIEW
            If gobjUser.HasPrivilegeEx(Priv_system_Administration, Priv_Write) Then
                QueryView = True
            Else
                QueryView = False
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
    
    '---perform actions based on the "view type"---
    If viViewType = cw32Global.cwiREGULAR_VIEW Then
        '---set global flags---
        cw32Global.cwbFirstRegularResize = True
        cw32Global.cwbMainViewRunning = True
            
        Call AddMenus
    
        '---------------------------------------------------------------------------------------
        '  If this OLE Server needs it's own menu items, here is a good place to add these
        '    menu items onto the CW Core's main menu.
        '
        '  NOTE: This is OPTIONAL.  You do NOT have to create menu items. Thats why they are
        '        commented out.
        '---------------------------------------------------------------------------------------
        '---add a menu item "DataManager" on CW's main menu---
        'ghDataManager = cwobjCWComm.AddPopupMenuToCW("&DataManager", cwlMyProgId)
        '---add a menu item "Ole Server Menu Item" to the menu "DataManager"---
        'glOleMenuItem = cwobjCWComm.AddMenuItem("&Ole Server Menu Item", ghDataManager, cwlMyProgId, "DataManager menu item")
        
        '-----------------------------------------------------------------------------------------
        '  Reset the "parent window" for this OLE server's "main" regular form.  This is the main
        '    "trick" that makes the OLE server look like it is part of the CW Core.
        '
        '  NOTE: This action invokes the main form's "Load" event, but the form is NOT shown.
        '        Make sure that you DO NOT perform a "Me.Show" in the form's "Load" event, as this
        '        will mess up the "SetParent" API call.  Make sure that you DO NOT "Show" the
        '        "main" regular form in this method, that will happen in the "ResizeView" method.
        '-----------------------------------------------------------------------------------------
        Set gfrmRegularView = New frmRegularView
        Set gclsRetirementProcess = New clsRetirementProcess
        Set gfrmRetireData = New frmRetireData
        SetParent gfrmRegularView.hwnd, vlParenthWnd
        
        '---set parameter---
        rhMainForm = gfrmRegularView.hwnd
    Else    '--unsupported view "type"--
        '---set return value---
        OpenView = False
        
        Exit Function
    End If
        
    '---set return value---
    OpenView = True
    
    Exit Function
    
ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "DataManager.OpenView", "DataManager", vbCritical)
    
    '---set return value---
    OpenView = False
    
    Exit Function
    Resume
    
End Function

Private Sub AddMenus()
    
    If gobjUser.HasPrivilegeEx(Priv_system_Administration, Priv_Write) Then
        If mlMenuRegen = 0 Then
            mlMenuRegen = gobjComm.AddActionMenuItem("&Regenerate Messages", "new.bmp", cwlMyProgId, "Regenerate Encounter Messages")
        End If
        
        If mlMenuPurgeAuditLog = 0 Then
            mlMenuPurgeAuditLog = gobjComm.AddActionMenuItem("&Purge Audit Log", "export_template.bmp", cwlMyProgId, "Purge Audit Log")
        End If
        
        
        'SCR 85617 - LCI
        If mlMenuDataRet = 0 Then
            ' SCR 113155 1/12/07 Quan
            mlMenuDataRet = gobjComm.AddActionMenuItem("&Data Extraction", "DataMove.bmp", cwlMyProgId, "Extract data from database")
        End If
    
        If mlMenuImport = 0 Then
            mlMenuImport = gobjComm.AddActionMenuItem("&Import ADS", "import_template.bmp", cwlMyProgId, "Import ADS Patients")
        End If
    End If
    
    If gobjUser.HasPrivilegeEx(Priv_Patient_Verification, Priv_Write) Then
        If mlMenuUnverified = 0 Then
            mlMenuUnverified = gobjComm.AddActionMenuItem("&Unverified Patient List", "new.bmp", cwlMyProgId, "Unverified Patient List")
        End If
                
        If mlMenuPrint = 0 Then
            mlMenuPrint = gobjComm.AddActionMenuItem("&Print", "print.bmp", cwlMyProgId, "Print List")
        End If
    End If
    
End Sub

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
            Unload gfrmRegularView
        End If
    ElseIf viViewType = cwiREGULAR_VIEW Then
        '---close the "main view"---
        If cwbMainViewRunning Then
            '---set global flags---
            cwbMainViewRunning = False
            
            '---this indicates that the CW Core is closing the "view", NOT this OLE Server---
            cwbCloseViewMyself = False
            
            '---close ALL forms that are associated with all "views"---
            Unload gfrmRegularView
        End If
    End If
    
    RemoveMenus
    Set gfrmRegularView = Nothing
    
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
    On Error GoTo ErrHandler
    
    '----------------------------------------------------------------------------------
    '  NOTE: You don't have to create menu items, these are optional, that is why
    '        the code is commented out.
    '----------------------------------------------------------------------------------
'    '---perform the actions associated with the menu item that was selected---
    Select Case vlMenuId
        Case mlMenuRegen
            gfrmRegularView.ResendMessages
            
        Case mlMenuPurgeAuditLog
            cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_DATA_MANAGER
            frmPurgeAuditLog.Show vbModal
            cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_DATA_MANAGER
        Case mlMenuDataRet
            gfrmRetireData.Show vbModal
        
        Case mlMenuImport
            frmImportADS.Show vbModal
            
        Case mlMenuUnverified
            gfrmRegularView.GetUnverifiedPatientList
            
        Case mlMenuPrint
            gfrmRegularView.PrintList
        
        Case Else   '--do nothing--
    End Select

    Exit Sub

ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "DataManager.MenuItemSelected", "DataManager", vbExclamation)
    
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
      displayCurrentPatient
    End If
End Sub

Private Sub RemoveMenus()

    If mlMenuRegen <> 0 Then
        Call gobjComm.DeleteMenuItem(mlMenuRegen)
        mlMenuRegen = 0
    End If
    
    If mlMenuPurgeAuditLog <> 0 Then
        Call gobjComm.DeleteMenuItem(mlMenuPurgeAuditLog)
        mlMenuPurgeAuditLog = 0
    End If
    
    If mlMenuDataRet <> 0 Then
        Call gobjComm.DeleteMenuItem(mlMenuDataRet)
        mlMenuDataRet = 0
    End If

    If mlMenuImport <> 0 Then
        Call gobjComm.DeleteMenuItem(mlMenuImport)
        mlMenuImport = 0
    End If
    
    If mlMenuUnverified <> 0 Then
        Call gobjComm.DeleteMenuItem(mlMenuUnverified)
        mlMenuUnverified = 0
    End If
        
    If mlMenuPrint <> 0 Then
        Call gobjComm.DeleteMenuItem(mlMenuPrint)
        mlMenuPrint = 0
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
    
    Dim blnResultIsRunning As Boolean
    Dim intI As Integer
    
    
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
            gfrmRegularView.WindowState = vbNormal
            
            ' SCR #23374;   Developer:  05/09/2002 11:46 AM
            'SCR Description: Add check to see whether the only modal form in DMFE is running.
            
            blnResultIsRunning = False
            For intI = 0 To Forms.Count - 1
                If Forms(intI).Name = "frmTransferResult" Then
                    blnResultIsRunning = True
                End If
            Next intI
            
            If blnResultIsRunning = False Then
                '---hide the form to so that annyoing "screen blinking" will not be seen while controls are moved and resized---
                gfrmRegularView.Visible = False
            
                '---resize form to fill entire "client area", NOTE: This generates a "resize" event---
                gfrmRegularView.Move cwlLeft, cwlTop, cwlWidth, cwlHeight
                
                gfrmRegularView.Visible = True
            Else ' Just resize the main form
                gfrmRegularView.Move cwlLeft, cwlTop, cwlWidth, cwlHeight
            End If
                
        Else    '--first "regular" resize, need to perform "initailization" procedures--
            '---reset flag so this code will not be executed again---
            
            cwbFirstRegularResize = False
        
            '----------------------------------------------------------------------------
            '  NOTE: Here is where you "Show" the "Regular View" form for the FIRST time.
            '----------------------------------------------------------------------------
            Set gfrmRegularView.mobjParent = Me
            gfrmRegularView.Display
        End If
    End If
    
    Exit Sub
    
ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "DataManager.ResizeView", "DataManager", vbExclamation)
    Exit Sub
End Sub

Private Sub Class_Initialize()
    '---create shared object used to show VB errors, etc.---
    Set gobjShared = New CWShared
    
    '---set the help file for this application---
    App.HelpFile = gobjShared.CWHelpFile
End Sub

Private Sub Class_Terminate()
    '---release reference to objects---
    Set gobjShared = Nothing
    Set cwobjCWComm = Nothing
    Set cwobjLogonEnviron = Nothing
    Set cwobjConfig = Nothing
    Set cwobjPatient = Nothing

    '********************************************************************
    '  TO DO:
    '    Add code here to release reference to all other global objects.
    '********************************************************************
    
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
Public Sub Message(ByVal viMessageType As Integer, _
                   ByVal vsMessageText As String, _
                   ByVal vsSourceServerNCID As String)
    On Error GoTo ErrHandler
    
    '-------------------------------------------------------------------------------
    '  Take corresponding actions based upon the message "type".
    '
    '  NOTE: This method is required, but currently there are NO REQUIRED messages
    '        to process.
    '-------------------------------------------------------------------------------
    Select Case viMessageType
    
        Case cwiOK_TO_CHANGE_PATIENT
            cwobjCWComm.Message cwiOK_TO_CHANGE_PATIENT, "Y", vsSourceServerNCID, NCID_DATA_MANAGER
        Case cwiPROPERTIES
            cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_CLINIC_SCHEDULE_APPLICATION
            OptionsDialog.Show vbModal
            cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_CLINIC_SCHEDULE_APPLICATION
        Case cwiOK_TO_CLOSE_VIEW
            cwobjCWComm.Message cwiOK_TO_CLOSE_VIEW, "Y", vsSourceServerNCID, NCID_DATA_MANAGER
        Case cwiBarCode
            '-- only react if I'm on top.
            Dim sMsg As String
            '-- ask core who's up
            Call cwobjCWComm.Message(cwiFOREGROUND_APPLICATION, sMsg, NCID_CW_APPLICATION, NCID_DATA_MANAGER)
            If sMsg = NCID_DATA_MANAGER Then
                Call ProcessBarcode(vsMessageText)
            End If
    
        Case Else   '--do nothing--
    End Select

    Exit Sub
    
ErrHandler:
    
        '---show the VB error---
        Call gobjShared.ShowVBError(Err.Number, Err.Description, "DataManager.Message", "DataManager", vbExclamation)
    
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
Public Sub Initialize(ByVal vlMyProgId As Long, _
                            robjComm As Object, _
                            robjLogonEnvironment As Object, _
                            robjConfig As Object, _
                            robjPatient As Object)
    On Error GoTo ErrHandler
    
    '---save the "ProgId" assigned to me by the CW Core---
    cwlMyProgId = vlMyProgId
    
    '---create references to passed in objects---
    Set cwobjCWComm = robjComm
    Set gobjComm = robjComm
    Set cwobjLogonEnviron = robjLogonEnvironment
    Set cwobjConfig = robjConfig
    Set cwobjPatient = robjPatient
    '---set global flags---
    cwbMainViewRunning = False
    '-- subscribe to the messages from the barcode scanner in the eventserver
    Call gobjComm.SubscribeToMessage(NCID_EVENT_SERVER_APPLICATION, cwiBarCode, NCID_DATA_MANAGER)
    
    Exit Sub

ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "DataManager.Initialize", "DataManager", vbExclamation)
    
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
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "DataManager.Properties", "DataManager", vbExclamation)
    
    Exit Sub
End Sub

Private Sub ICHCSII_CoreAPI2_Initialize2(ByVal vlMyProgId As Long, _
                                               robjComm As Object, _
                                               robjLogonEnvironment As Object, _
                                               robjConfig As Object, _
                                               robjPatient As Object, _
                                               robjUser As Object)
    Set gobjUser = robjUser
    Call Initialize(vlMyProgId, robjComm, robjLogonEnvironment, robjConfig, robjPatient)
    
End Sub

Private Function ICHCSII_CoreAPI3_CloseView2(ByVal viViewType As Integer) As Boolean
  Call CloseView(viViewType)
  ICHCSII_CoreAPI3_CloseView2 = True
End Function

'<SCR 23902 CC>

' Begin:SCR #27451;   Developer: MAG 09/20/2002 12:12 PM
Public Sub displayCurrentPatient(Optional objTmpPatient As MMMHISPatient.Patient = Nothing)
  
    Dim strStreet1 As String
    Dim strStreet2 As String
    Dim nRow As Long
    Dim objPatient As MMMHISPatient.Patient
    Dim strAddress As String

    Dim oPhone As GEMS_MMIObj.GEMS_Phone
    Dim sTemp As String
    
On Error GoTo ErrHandler

    ' Begin:SCR #27451;   Developer: MAG 09/20/2002 12:11 PM
    If objTmpPatient Is Nothing Then
        Set objPatient = cwobjPatient
        'Set objPatient = gobjPatient
    Else
        Set objPatient = objTmpPatient
    End If
    ' End: SCR #27451;
    

    If checkDuplications(objPatient.UnitNumber) = True Then
        Exit Sub
    End If
    
    If gfrmRegularView.cboSource.text = "PIC" _
      Or gfrmRegularView.cboSource.text = "DISK" _
      Or gfrmRegularView.cboSource.text = "PVF" _
      Or gfrmRegularView.cboSource.text = "AHLTA-MOBILE" Then
    'Only load PIC/DISK patients from the PIC File, otherwise UnitNumbers will be invalid
    'If gfrmRegularView.grdSearch.Rows > 1 Then
        Exit Sub
    'End If
    ElseIf gfrmRegularView.cboSource.text = "EIC" Then
        If gfrmRegularView.grdSearch.Rows > 1 Then 'SCR 25149
            For nRow = 2 To gfrmRegularView.grdSearch.Rows
                gfrmRegularView.grdSearch.RemoveItem 1
                'remove the patient from EIC first
            Next nRow
        End If
    End If
    
    'Tracker #53005 Developer: Mag 04/28/04
    If Trim$(objPatient.SSN) <> "" Then
      With gfrmRegularView.grdSearch
        If bResendTMIPMessages = True Then
            'DBMCHANGE
            If .Rows = 1 Then
              .AddItem ""
            End If
            
            nRow = 1
            If gfrmRegularView.cmdTMIP.Tag = "PVF" And nCurrentPatientCnt >= PatientCnt Then
              nCurrentPatientCnt = 0
            End If
            
            nCurrentPatientCnt = nCurrentPatientCnt + 1
            
            .TextMatrix(nRow, TXT_NAME) = "Processing " & gfrmRegularView.cmdTMIP.Tag & " " & nCurrentPatientCnt & " of " & PatientCnt
            
            .TextMatrix(nRow, TXT_UNIT_NUMBER) = objPatient.UnitNumber
            
            .IsSelected(nRow) = True
            Exit Sub
            'DBMCHANGE
        End If
        
          .AddItem ""
          nRow = .Rows - 1
          
          .TextMatrix(nRow, TXT_NAME) = objPatient.Name
          If Len(objPatient.SSN) = 9 Then
            .TextMatrix(nRow, TXT_SSN) = Format(objPatient.SSN, "000-00-0000")
          Else
            .TextMatrix(nRow, TXT_SSN) = objPatient.SSN
          End If
          
          ' Begin:SCR #22589;   Developer: MAG 02/25/2003 03:33 PM
          If Len(objPatient.SponsorSSN) = 9 Then
            sTemp = objPatient.FMP & "/" & Format(objPatient.SponsorSSN, "000-00-0000")
          Else
            sTemp = objPatient.FMP & "/" & objPatient.SponsorSSN
          End If
          
          .TextMatrix(nRow, TXT_FMP_SSN) = sTemp
          ' End: SCR #22589;
          If objPatient.DateOfBirth <> cwdNULL_DATE And objPatient.DateOfBirth <> 0 Then
              .TextMatrix(nRow, TXT_DOB) = Format$(objPatient.DateOfBirth, "mm/dd/yyyy")
          End If
          .TextMatrix(nRow, TXT_SEX) = objPatient.Sex
          
            For Each oPhone In objPatient.CurrentPatient.phones
                If oPhone.PhoneType = NCID_HOME_PHONE Then
                    .TextMatrix(nRow, TXT_HOMEPHONE) = oPhone.AreaCode & "-" & oPhone.Phone
                    Exit For
                End If
            Next
            
            For Each oPhone In objPatient.CurrentPatient.phones
                If oPhone.PhoneType = NCID_WORK_PHONE Then
                     .TextMatrix(nRow, TXT_WORKPHONE) = oPhone.AreaCode & "-" & oPhone.Phone
                    Exit For
                End If
            Next
          
          strStreet1 = Trim$(objPatient.Street1)
          strStreet2 = Trim$(objPatient.Street2)
          strAddress = objPatient.City & ", " & objPatient.State & " " & objPatient.ZipCode
          '--- address2 exists ----
          strAddress = IIf(Len(strStreet2), strStreet2 & " " & strAddress, strAddress)
          '--- address1 exits ---
          strAddress = IIf(Len(strStreet1), strStreet1 & " " & strAddress, strAddress)
          '--- either address1 or address2 exists ---
          strAddress = IIf(Len(strStreet1) = 0 And Len(strStreet2) = 0, " ", strAddress)
          
          .TextMatrix(nRow, TXT_ADDRESS) = strAddress
          
          .TextMatrix(nRow, TXT_UNIT_NUMBER) = objPatient.UnitNumber
          
          .IsSelected(nRow) = True
  
      End With
    End If
Exit Sub

ErrHandler:
    gobjShared.ShowVBError Err.Number, Err.Description, "DataManagerOps.displayCurentPatient", "DataManager", vbCritical
End Sub

Public Function checkDuplications(UnitNumber As Long) As Boolean
  Dim match As Boolean
  Dim idx   As Integer
  
  match = False
  For idx = 1 To (gfrmRegularView.grdSearch.Rows - 1) And (Not match)
                                                'SCR  128654
    If gfrmRegularView.grdSearch.TextMatrix(idx, TXT_UNIT_NUMBER) = CStr(UnitNumber) Then
      match = True
    End If
  Next idx

  checkDuplications = match

End Function


Private Sub ProcessBarcode(sEntry As String)
   Dim msg As String
   Dim vEntry As Variant
   Dim PatientSSN As String
   Dim SponsorSSN As String
   Dim DOB As String
   Dim DDS As String
   Dim FirstName As String
   Dim LastName As String
   Dim Sex As String
   Dim grdData As String
   Dim ErrMsg As String
   Dim Unit_Number As String
   Dim oPatient As Patient
   Set oPatient = New Patient
   
      ' initialize storage variables
      msg = sEntry
      PatientSSN = ""
      SponsorSSN = ""
      DOB = ""
      DDS = ""
      FirstName = ""
      LastName = ""
      Sex = ""
      grdData = ""
      ' get data from XML string
      Call GetXMLInfo(msg, PatientSSN, FirstName, LastName, DOB, DDS, SponsorSSN)
      
      If ValidateSSN(PatientSSN) = True Then
         grdData = LastName & ", " & FirstName & vbTab & PatientSSN & vbTab
      End If
      If DDS <> "" And SponsorSSN <> "" Then
         If ValidateSSN(SponsorSSN) = True Then
            grdData = grdData & DDS & "/" & SponsorSSN & vbTab
         End If
      Else
         If grdData <> "" Then
            grdData = grdData & vbTab
         End If
      End If
      If grdData <> "" Then
         If CDbl(PatientSSN) = 0 Then
            ErrMsg = "Barcode contains missing or illegal values" & vbCrLf & vbCrLf
            ErrMsg = ErrMsg & "Patient Name: " & vbTab & LastName & ", " & FirstName & vbCrLf
            If DOB <> "" Then
               ErrMsg = ErrMsg & "Date of Birth: " & vbTab & Format(DOB, "MM/DD/YYYY") & vbCrLf
            End If
            ErrMsg = ErrMsg & "Patient SSN: " & vbTab & PatientSSN & vbCrLf
            ErrMsg = ErrMsg & "FMP/Sponsor SSN: " & DDS & "/" & SponsorSSN & vbCrLf
            ErrMsg = ErrMsg & vbCrLf & "Patient will NOT be added to patient list" & vbCrLf
            Call MsgBox(ErrMsg, vbCritical, "Add BarCode Patient")
         Else
            Unit_Number = oPatient.sGetUnitNumberFromSSN(PatientSSN)  'GetUnitNumberFromSSN(PatientSSN, Unit_Number)
            If Unit_Number = "" Then
               ErrMsg = "Unable to locate Patient in Database" & vbCrLf & vbCrLf
               ErrMsg = ErrMsg & "Patient Name: " & vbTab & LastName & ", " & FirstName & vbCrLf
               If DOB <> "" Then
                  ErrMsg = ErrMsg & "Date of Birth: " & vbTab & Format(DOB, "MM/DD/YYYY") & vbCrLf
               End If
               ErrMsg = ErrMsg & "Patient SSN: " & vbTab & PatientSSN & vbCrLf
               ErrMsg = ErrMsg & "FMP/Sponsor SSN: " & DDS & "/" & SponsorSSN & vbCrLf
               ErrMsg = ErrMsg & vbCrLf & "Patient will NOT be added to patient list" & vbCrLf
               Call MsgBox(ErrMsg, vbCritical, "Add BarCode Patient")
               
               grdData = ""
            Else
               Call oPatient.LoadPatient(CLng(Unit_Number), "")
               grdData = oPatient.Name & vbTab & Format(oPatient.SSN, "000-00-0000") & vbTab & oPatient.FMP & "/" & Format(oPatient.SponsorSSN, "000-00-0000")
               grdData = grdData & vbTab & CDate(oPatient.DateOfBirth)
               grdData = grdData & vbTab & oPatient.Sex '& vbTab & oPatient.HomePhone & vbTab & oPatient.WorkPhone
               Dim oPhone As GEMS_MMIObj.GEMS_Phone
               For Each oPhone In oPatient.CurrentPatient.phones
                    If oPhone.PhoneType = NCID_HOME_PHONE Then
                        grdData = grdData & vbTab & oPhone.AreaCode & oPhone.Phone
                        Exit For
                    End If
               Next
               
               For Each oPhone In oPatient.CurrentPatient.phones
                    If oPhone.PhoneType = NCID_WORK_PHONE Then
                        grdData = grdData & vbTab & oPhone.AreaCode & oPhone.Phone
                        Exit For
                    End If
               Next
               
               grdData = grdData & vbTab & oPatient.Street1 & ", " & oPatient.Street2 & ", " & oPatient.City & ", " & oPatient.State & ", " & oPatient.ZipCode & "  " & oPatient.Country & vbTab & oPatient.UnitNumber
               
               Call gfrmRegularView.GrdSearchAddItem(grdData)
            End If
         End If
      Else
         ErrMsg = "Barcode contains missing or illegal values" & vbCrLf & vbCrLf
         ErrMsg = ErrMsg & "Patient Name: " & vbTab & LastName & ", " & FirstName & vbCrLf
         If DOB <> "" Then
            ErrMsg = ErrMsg & "Date of Birth: " & vbTab & Format(DOB, "MM/DD/YYYY") & vbCrLf
         End If
         ErrMsg = ErrMsg & "Patient SSN: " & vbTab & PatientSSN & vbCrLf
         ErrMsg = ErrMsg & "FMP/Sponsor SSN: " & DDS & "/" & SponsorSSN & vbCrLf
         ErrMsg = ErrMsg & vbCrLf & "Patient will NOT be added to patient list" & vbCrLf
         Call MsgBox(ErrMsg, vbCritical, "Add BarCode Patient")
      End If
      Set oPatient = Nothing
      
'      Call MsgBox(msg, vbInformation, "cmdBarCodes_Click")
   
End Sub

Private Sub GetXMLInfo(XMLObj As String, PatientSSN As String, FirstName As String, LastName As String, DOB As String, DDS As String, SponsorSSN As String)
   Dim objXML As MSXML.DOMDocument
'   Dim objNodeList As MSXML.IXMLDOMNodeList
   Dim bRet As Boolean
   
   On Error GoTo GetXMLInfoError
   
   Set objXML = New MSXML.DOMDocument
   objXML.validateOnParse = False
   bRet = objXML.loadXML(XMLObj)
   If bRet = True Then
      PatientSSN = GetValueOfTag(objXML, "PatientSSN")
      SponsorSSN = GetValueOfTag(objXML, "SponsorSSN")
      FirstName = GetValueOfTag(objXML, "FirstName")
      LastName = GetValueOfTag(objXML, "LastName")
      DOB = GetValueOfTag(objXML, "DOB")
      DDS = GetValueOfTag(objXML, "DDS")
   Else
      Call MsgBox("XML is not well formed: " & XMLObj, vbCritical, "GetXMLInfo Function")
   End If
   Exit Sub
GetXMLInfoError:
   Dim ErrMsg As String
   ErrMsg = Err.Number & vbCrLf
   ErrMsg = ErrMsg & Err.Description
   Call MsgBox(ErrMsg, vbCritical, "GetXMLInfo Function")
End Sub

Private Function ValidateSSN(sSSN As String) As Boolean

  'Valid formats: [a]=123456789 and [b]=123-45-6789
  ValidateSSN = False

  Dim sTokens() As String

  If Len(sSSN) = 9 Then
    If IsNumeric(sSSN) Then
        ValidateSSN = True
    End If
  ElseIf Len(sSSN) = 11 Then
    sTokens = Split(sSSN, "-")
    If UBound(sTokens) = 2 Then
        If Len(sTokens(0)) = 3 And Len(sTokens(1)) = 2 And Len(sTokens(2)) = 4 And IsNumeric(Join(sTokens, "")) Then
            ValidateSSN = True
        End If
    End If
  End If

End Function

Private Function GetValueOfTag(objXMLDoc As MSXML.DOMDocument, sTag As String) As String
    '*** GetValueOfTag(objXMLDoc, "BARCODE")
    
    Dim objNodeList As IXMLDOMNodeList

    GetValueOfTag = ""
    
    If (objXMLDoc Is Nothing) Then
        MsgBox "objXMLDoc is not set", vbInformation, "GetValueOfTag()"
    Else
        Set objNodeList = objXMLDoc.getElementsByTagName(sTag)
        If objNodeList.length > 0 Then
            GetValueOfTag = objNodeList.item(0).text
        End If
    End If
    
End Function

' Begin:SCR #39268;   Developer: Mag 08/05/2003 10:33 AM
' Begin:SCR #37880;   Developer: Mag 06/19/2003 06:11 PM
'Get Called from Encounter when an Encounter is Signed
Public Sub Send2TMIP(eStatus As ENCOUNTER_INTERFACES.EncStatus)
'This code was moved from oCurrentEncounter_CurrEncStatusChanged procedure to here
'No need to catch the CurrEncStatusChanged Event anymore
    'If the status has been changed to "complete" then automatically send to TMIP
    
    '<< Begin:SCR #561;   Developer: Brian Mowbray 07/15/2002 05:49 PM
    '<< Note: Function Changed:  CurrEncStatusChanged Event was sending the signed encounter automatically when the ecounter
    '<< was signed for all connection modes (ITT, CDR, Theater). Changed method to only be called when in Theater mode to keep
    '<< in line with current logic of only allowing tranfer of data to LDDB when in CDR mode. In addition, ITT will not be
    '<< incorporating any data transfer functionality.
    
    If gobjShared.AppMode = modeTheater Then
    
'''      '<< Begin:SCR #29068;   Developer: Brian Mowbray 11/11/2002 06:19 PM
'''      If TMIPFrameworkFlag = 0 Then
'''        Exit Sub
'''      End If
'''      '>> End: SCR #29068;
    
      Dim objDMFData As DMFTransactionData
      Dim objDMFTransferManager As DMFTransferManager
      Dim lDataSets() As Long
      Dim myPatient(0) As DMFPatientInfo  'There will only be one patient in the transfer
      
      On Error GoTo ErrHandler
      
      If eStatus = ENCOUNTER_INTERFACES.EncStatus.complete Or ENCOUNTER_INTERFACES.EncStatus.Updated Then
      
      'TMIP File Transfer
      '======================
        Set objDMFData = New DMFTransactionData
        
        With objDMFData
        
                        
            
            myPatient(0).sUnitNumber = cwobjPatient.UnitNumber
            .PatientInfo = myPatient
                    
            .Source = eDM_LDDB
            .Destination = eDM_TMIP
            
            ReDim lDataSets(1)
            lDataSets(0) = NCID_DEMOGRAPHICS_APPLICATION
            lDataSets(1) = NCID_ENCOUNTER_APPLICATION
            .Dataset = lDataSets
            
            '********************************
            .Compression = eDM_NoCompression
            .Encryption = eDM_Encrypted
            '********************************
            
            .Format = eDM_XML
            .IsEICDestination = False
            .DestinationPath = ""
            .SourcePath = ""
            
            '<< Begin:SCR #30953;   Developer: Brian Mowbray 12/18/2002 03:36 PM
            .EncounterSource = enEncounterSigned
            '>> End: SCR #30953;
            
        End With
        
        Set objDMFTransferManager = New DMFTransferManager
        
        If Not gobjComm Is Nothing Then
            Call gobjComm.SetStatusText("Sending Encounter to TMIP ...")
        End If
    
        Call objDMFTransferManager.ExecuteTransfer(objDMFData)

      End If
    End If
    '>> End: SCR #561;

Exit Sub

ErrHandler:
    Screen.MousePointer = vbDefault
    gobjShared.ShowVBError Err.Number, Err.Description, "DataManagerOps.Send2TMIP", "DMF TransferManager", vbCritical
    Err.Clear
End Sub

'This is called when a patient is added or edited.
'SCR 120505
Public Sub SendDemographics2TMIP(UnitNumber As String)
    
    If gobjShared.AppMode = modeTheater Then
      
      Dim objDMFData As DMFTransactionData
      Dim objDMFTransferManager As DMFTransferManager
      Dim lDataSets() As Long
      Dim myPatient(0) As DMFPatientInfo  'There will only be one patient in the transfer
      
      On Error GoTo ErrHandler
      
      'TMIP File Transfer
      '======================
        Set objDMFData = New DMFTransactionData
        
        With objDMFData
            myPatient(0).sUnitNumber = UnitNumber
            .PatientInfo = myPatient
                    
            .Source = eDM_LDDB
            .Destination = eDM_TMIP
            
            ReDim lDataSets(1)
            lDataSets(0) = NCID_DEMOGRAPHICS_APPLICATION
            .Dataset = lDataSets
            
            '********************************
            .Compression = eDM_NoCompression
            .Encryption = eDM_Encrypted
            '********************************
            
            .Format = eDM_XML
            .IsEICDestination = False
            .DestinationPath = ""
            .SourcePath = ""
            
            '<< Begin:SCR #30953;   Developer: Brian Mowbray 12/18/2002 03:36 PM
            .EncounterSource = enEncounterSigned
            '>> End: SCR #30953;
            
        End With
        
        Set objDMFTransferManager = New DMFTransferManager
        
        If Not gobjComm Is Nothing Then
            Call gobjComm.SetStatusText("Sending Encounter to TMIP ...")
        End If
    
        Call objDMFTransferManager.ExecuteTransfer(objDMFData)

    End If
    '>> End: SCR #561;

Exit Sub

ErrHandler:
    Screen.MousePointer = vbDefault
    gobjShared.ShowVBError Err.Number, Err.Description, "DataManagerOps.Send2TMIP", "DMF TransferManager", vbCritical
    Err.Clear
End Sub

Public Sub SendInpatientStatus(InpatientID As Long)

'This is modeled after Sent2TMIP
'This is called by the pad project to send Inpatient Status messages to JMEWS.

    If gobjShared.AppMode = modeTheater Then
    
      Dim objDMFData As DMFTransactionData
      Dim objDMFTransferManager As DMFTransferManager
      Dim lDataSets() As Long
      Dim myPatient(0) As DMFPatientInfo  'There will only be one patient in the transfer
      
      On Error GoTo ErrHandler
     
      'TMIP File Transfer
      '======================
        Set objDMFData = New DMFTransactionData
        
        With objDMFData
            
            myPatient(0).sUnitNumber = cwobjPatient.UnitNumber
            .PatientInfo = myPatient
                    
            .Source = eDM_LDDB
            .Destination = 8  ' eDM_JMEWS
            
            'SF
            ReDim lDataSets(4)
            lDataSets(0) = NCID_DEMOGRAPHICS_APPLICATION
            lDataSets(1) = NCID_ALLERGIES_APPLICATION
            lDataSets(2) = NCID_PROBLEMS_APPLICATION
            lDataSets(3) = NCID_MEDICATIONS_APPLICATION
            lDataSets(4) = NCID_INPATIENT_APPLICATION
            .Dataset = lDataSets
            
            '**************************************
            .Compression = eDM_NoCompression
            .Encryption = eDM_Encrypted
            '**************************************
            
            .Format = eDM_XML
            .IsEICDestination = False
            .DestinationPath = ""
            .SourcePath = ""
            
            'We don't have a better place for this so put it here until we can break the interfaces
            .EncounterSource = InpatientID
            
        End With
        
        Set objDMFTransferManager = New DMFTransferManager
        
        If Not gobjComm Is Nothing Then
            Call gobjComm.SetStatusText("Sending Inpatient Status message to JMeWS ...")
        End If
    
        Call objDMFTransferManager.ExecuteTransfer(objDMFData)
    
    End If

    '>> End: SCR #561;

Exit Sub

ErrHandler:
    Screen.MousePointer = vbDefault
    gobjShared.ShowVBError Err.Number, Err.Description, "DataManagerOps.SendInpatientStatus", "DMF TransferManager", vbCritical
    Err.Clear
End Sub
' End: SCR #39268;

