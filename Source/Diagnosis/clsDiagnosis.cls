VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Diagnosis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True


Option Explicit   '-- $Reformatted by FormatVB, Ver. 1.6b, on 02-23-00 at 10:50:08
Dim fieldValues(8)  As Variant
Dim colWidths(3)    As Integer
Dim colNames(3)     As String
Dim objListManagementOps As ListManagementOps
Private objForm     As Object
Const chronicComboNames = "|Chronic" & "|Acute "
Dim mDiagForm As frmDiagnosis

Implements ICHCSII_CoreAPI2
Implements ICHCSII_User
Implements ICHCSII_AnP
#If DEBUGON Then
Private Const Module = "Diagnosis.cls"
#End If
Public Property Set objCurrentForm(value As Object)
   Set objForm = value
End Property
Public Sub setPrimaryFocus()
  If mDiagForm.Visible Then mDiagForm.cboSearchDiagnosis.SetFocus
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
         QueryView = True              '--this view is NOT optional--
      Case cwiPATIENT_SUMMARY_VIEW
         QueryView = False
      Case cwiCLINICIAN_SUMMARY_VIEW
         QueryView = False
      Case cwiENCOUNTER_SUMMARY_VIEW
         QueryView = False
      Case cwiPROPERTY_VIEW
         QueryView = True
      Case Else                        '--unknown "view type", not supported--
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
Public Function OpenView(ByVal viViewType As Integer, _
                          ByVal vlParenthWnd As Long, _
                          rhMainForm As Long) _
                          As Boolean

   On Error GoTo ErrHandler

   '-- -perform actions based on the "view type"---

   If viViewType = cwiREGULAR_VIEW Then
      '-- -set global flags---
      cwbFirstRegularResize = True
      cwbMainViewRunning = True
      
      '-- -------------------------------------------------------------------------------------
      '--  If this OLE Server needs it's own menu items, here is a good place to add these
      '--    menu items onto the CW Core's main menu.
      '--
      '--  NOTE: This is OPTIONAL.  You do NOT have to create menu items. Thats why they are
      '--        commented out.
      '-- -------------------------------------------------------------------------------------
      '-- -add a menu item "Template" on CW's main menu---
      '-- ghTemplate = cwobjCWComm.AddPopupMenuToCW("&Template", cwlMyProgId)
      '-- -add a menu item "Ole Server Menu Item" to the menu "Template"---
      '-- glolemenuitem = gobjAnPParent.AddMenuItem(0, "&Ole Server Menu Item", "summary.bmp", "Template menu item")
      
      '-- ---------------------------------------------------------------------------------------
      '--  Reset the "parent window" for this OLE server's "main" regular form.  This is the main
      '--    "trick" that makes the OLE server look like it is part of the CW Core.
      '--
      '--  NOTE: This action invokes the main form's "Load" event, but the form is NOT shown.
      '--        Make sure that you DO NOT perform a "Me.Show" in the form's "Load" event, as this
      '--        will mess up the "SetParent" API call.  Make sure that you DO NOT "Show" the
      '--        "main" regular form in this method, that will happen in the "ResizeView" method.
      '-- ---------------------------------------------------------------------------------------
      Set mDiagForm = New frmDiagnosis
      Load mDiagForm
      mDiagForm.ParentObject = Me
      SetParent mDiagForm.hwnd, vlParenthWnd
      
      '-- -set parameter---
      rhMainForm = mDiagForm.hwnd
   ElseIf viViewType = cwiENCOUNTER_SUMMARY_VIEW Then
      '-- -set global flags---
      cwbFirstRegularResize = True

      frmTeleconView.ParentObject = Me
      SetParent frmTeleconView.hwnd, vlParenthWnd
      
      '-- -set parameter---
      rhMainForm = frmTeleconView.hwnd
   Else                                '--unsupported view "type"--
      '-- -set return value---
      OpenView = False
      
      Exit Function

   End If
   
   '-- -set return value---
   OpenView = True
   
   Exit Function
   
ErrHandler:
   '-- -show the VB error---
   Call gobjShared.ShowVBError(Err.Number, Err.Description, "AnP.OpenView", "AnP", vbCritical)
   
   '-- -set return value---
   OpenView = False
   
   Exit Function

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
      '-- -close the "main view"---

      If cwbMainViewRunning Then
         '-- -set global flags---
         cwbMainViewRunning = False
         
         '-- -this indicates that the CW Core is closing the "view", NOT this OLE Server---
         cwbCloseViewMyself = False
         
         '-- -close ALL forms that are associated with all "views"---
         'Unload frmDiagnosis
         Unload mDiagForm
         Set mDiagForm = Nothing
         Set gobjAnPParent = Nothing   ' - the 'parent node'  Use it to do things like shut down this child
         'Set gobjDxCollection = Nothing        'Phil Crowder 10/02/01  Not needed with new AnP interface
         Set gobjProblemListAcute = Nothing
         Set gobjProblemListChronic = Nothing
      End If

   ElseIf viViewType = cwiREGULAR_VIEW Then
      '-- -close the "main view"---

      If cwbMainViewRunning Then
         '-- -set global flags---
         cwbMainViewRunning = False
         
         '-- -this indicates that the CW Core is closing the "view", NOT this OLE Server---
         cwbCloseViewMyself = False
         
         '-- -close ALL forms that are associated with all "views"---
         'Unload frmDiagnosis
         Unload mDiagForm
         Set mDiagForm = Nothing
         Set gobjAnPParent = Nothing   ' - the 'parent node'  Use it to do things like shut down this child
         'Set gobjDxCollection = Nothing        'Phil Crowder 10/02/01  Not needed with new AnP interface
         Set gobjProblemListAcute = Nothing
         Set gobjProblemListChronic = Nothing
      End If
   ElseIf viViewType = cwiENCOUNTER_SUMMARY_VIEW Then
         '-- -close ALL forms that are associated with all "views"---
         Unload frmTeleconView
         Set gobjAnPParent = Nothing   ' - the 'parent node'  Use it to do things like shut down this child
         'Set gobjDxCollection = Nothing        'Phil Crowder 10/02/01  Not needed with new AnP interface
         Set gobjProblemListAcute = Nothing
         Set gobjProblemListChronic = Nothing
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

   On Error GoTo ErrHandler
   
   '-- --------------------------------------------------------------------------------
   '--  NOTE: You don't have to create menu items, these are optional, that is why
   '--        the code is commented out.
   '-- --------------------------------------------------------------------------------
'    '---perform the actions associated with the menu item that was selected---
'    Select Case vlMenuId
'        Case glOleMenuItem
'            '******************************************************************************
'            '  TO DO:
'            '    Take actions needed to process this menu item.
'            '******************************************************************************
'
'
'
'        Case Else   '--do nothing--
'    End Select
   
   Exit Sub
   
ErrHandler:
   '-- -show the VB error---
   Call gobjShared.ShowVBError(Err.Number, Err.Description, "Diagnosis.MenuItemSelected", "Template", vbExclamation)
   
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

   '-- -update the all running "views" with the new patient---

   If cwbMainViewRunning Then
      '-- ****************************************************************************
      '--  TO DO:
      '--    Add code here to update the data in all "running" views with this new
      '--    patient.
      '-- ****************************************************************************
      
      
      
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
Public Sub ResizeView(ByVal viViewType As Integer, _
                       ByVal vlLeft As Long, _
                       ByVal vlTop As Long, _
                       ByVal vlWidth As Long, _
                       ByVal vlHeight As Long)

   On Error GoTo ErrHandler
   
   '-- -reset global "client area" parameters---
   cwlLeft = vlLeft
   cwlTop = vlTop
   cwlWidth = vlWidth
   cwlHeight = vlHeight
   
   '-- -perform different actions based upon the "view type"---

   If viViewType = cwiREGULAR_VIEW Then

      If Not cwbFirstRegularResize Then
         '-- -------------------------------------------------------------------------------
         '--  Reset the "WindowState" to "Normal", so the form can be resized.  This is
         '--    needed because a window CANNOT be moved or resized if it is "minimized" or
         '--    "maximized".
         '--
         '--  NOTE: This generates a form "resize" event.
         '-- -------------------------------------------------------------------------------
         mDiagForm.WindowState = vbNormal
         
         '-- -hide the form to so that annyoing "screen blinking" will not be seen while controls are moved and resized---
         mDiagForm.AutoRedraw = False
         
         '-- -resize form to fill entire "client area", NOTE: This generates a "resize" event---
         mDiagForm.Move cwlLeft, cwlTop, cwlWidth, cwlHeight
         
         mDiagForm.AutoRedraw = True
         mDiagForm.ZOrder 0
      Else                             '--first "regular" resize, need to perform "initailization" procedures--
         '-- -reset flag so this code will not be executed again---
         cwbFirstRegularResize = False
         
         '-- --------------------------------------------------------------------------
         '--  NOTE: Here is where you "Show" the "Regular View" form for the FIRST time.
         '-- --------------------------------------------------------------------------
         mDiagForm.Display
      End If
   ElseIf viViewType = cwiENCOUNTER_SUMMARY_VIEW Then

      If Not cwbFirstRegularResize Then
         '-- -------------------------------------------------------------------------------
         '--  Reset the "WindowState" to "Normal", so the form can be resized.  This is
         '--    needed because a window CANNOT be moved or resized if it is "minimized" or
         '--    "maximized".
         '--
         '--  NOTE: This generates a form "resize" event.
         '-- -------------------------------------------------------------------------------
         frmTeleconView.WindowState = vbNormal
         
         '-- -hide the form to so that annyoing "screen blinking" will not be seen while controls are moved and resized---
         frmTeleconView.Visible = False
         
         '-- -resize form to fill entire "client area", NOTE: This generates a "resize" event---
         frmTeleconView.Move cwlLeft, cwlTop, cwlWidth, cwlHeight
         
         frmTeleconView.Visible = True
      Else                             '--first "regular" resize, need to perform "initailization" procedures--
         '-- -reset flag so this code will not be executed again---
         cwbFirstRegularResize = False
         
         '-- --------------------------------------------------------------------------
         '--  NOTE: Here is where you "Show" the "Regular View" form for the FIRST time.
         '-- --------------------------------------------------------------------------
         frmTeleconView.Display
      End If
   End If
   
   Exit Sub
   
ErrHandler:
   '-- -show the VB error---
   Call gobjShared.ShowVBError(Err.Number, Err.Description, "Template.ResizeView", "Template", vbExclamation)
   
   Exit Sub

End Sub

Private Sub Class_Initialize()
'ARM-SCR 19539 PJS 12.18.2001
modARMGlobalsDiagnosis.ARM_InitAllTransactions

'-- -create shared object used to show VB errors, etc.---
#If DEBUGON Then
   Set DL = GetObject("", "dlog.cdebuglog")
   DL.debugStart
   Dim DR  As DebugRec

   DL.DebugLog Exe, Module, "Class_Initialize", "Begin", DR
#End If
   Set gobjShared = New CWShared
   
   '-- -set the help file for this application---
   App.HelpFile = gobjShared.CWHelpFile
   
   '-- ----------------------------------
   '--  This statement is needed to debug
   '--    this OLE Server.
   '-- ----------------------------------
#If ciDEBUG = 1 Then

   If cwbDebugLogon = False Then
      Exit Sub

   End If

#End If
#If DEBUGON Then
   DL.DebugLog Exe, Module, "Class_Initialize", "End", DR
#End If

End Sub

Private Sub Class_Terminate()

   '-- -release reference to objects---
#If DEBUGON Then
   Dim DR  As DebugRec

   DL.DebugLog Exe, Module, "Class_Terminate", "Begin", DR
#End If
   Set gobjShared = Nothing
   Set objListManagementOps = Nothing
   Set cwobjCWComm = Nothing
   Set cwobjLogonEnviron = Nothing
   Set cwobjConfig = Nothing
   Set cwobjPatient = Nothing
   
   '-- ********************************************************************
   '--  TO DO:
   '--    Add code here to release reference to all other global objects.
   '-- ********************************************************************
   
   
   
#If ciDEBUG = 1 Then
   cwbDebugLogOff
#End If
#If DEBUGON Then
   DL.DebugLog Exe, Module, "Class_Terminate", "End", DR
   Set DL = Nothing
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
Public Sub Message(ByVal viMessageType As Integer, _
                    ByVal vsMessageText As String, _
                    ByVal vsSourceServerNCID As String)

   On Error GoTo ErrHandler
   
   '-- -----------------------------------------------------------------------------
   '--  Take corresponding actions based upon the message "type".
   '--
   '--  NOTE: This method is required, but currently there are NO REQUIRED messages
   '--        to process.
   '-- -----------------------------------------------------------------------------
   Select Case viMessageType
        Case cwiBACKGROUND_APPLICATION
        
            If gbGreenList = True Then
                objForm.ucTemplateNavigation1.PersistHistory
            End If
          
        Case cwiFOREGROUND_APPLICATION
            If gbGreenList = True Then
                objForm.ucTemplateNavigation1.RetrieveHistory
            End If
        Case cwiGENERIC
            Select Case vsMessageText
            
            'Refresh template data if new template is chosen SCR-11674
            Case "REFRESH TEMPLATES"
                If gbGreenList = False Then

                    'gobjAnPParent.HasTemplate = True  scr 23228  Phil Crowder 4/22/02 Should not be set by Diagnosis Module
                    
                    'scr 31342 Phil Crowder 1/13/02
                    If Not (objForm Is Nothing) Then
                      If Not objForm.DiagListOpt(0).value Then
                        objForm.SelectList 0
                      Else
                        load_template_data
                        load_finalize_datagrid
                      End If
                    End If
                    
                Else
                        load_template_favs
                        load_finalize_datagrid

                End If
            End Select
            
        '102879 Refresh the templates list in A&P when coming from Templates if it was selected.
        Case 33
            If Not objForm Is Nothing Then
                    If Not (objForm Is Nothing) Then
                      If objForm.DiagListOpt(0).value Then
                        objForm.SelectList 0
                      Else
'                        load_template_data
'                        load_finalize_datagrid
                      End If
                    End If
            End If
            
   End Select
   
   Exit Sub
   
ErrHandler:
   '-- -show the VB error---
   Call gobjShared.ShowVBError(Err.Number, Err.Description, "Diagnosis.Message", "Diagnsosis", vbExclamation)
    'Resume Next
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
     '---Instantiate Help-file object
   App.HelpFile = gobjShared.CWHelpFile
   '-- -save the "ProgId" assigned to me by the CW Core---
   cwlMyProgId = vlMyProgId
   
   '-- -create references to passed in objects---
   Set cwobjCWComm = robjComm
   Set cwobjLogonEnviron = robjLogonEnvironment
   Set cwobjConfig = robjConfig
   Set cwobjPatient = robjPatient
   
    'Create List mangement Object
    Dim objListManagement As ListManagement
    
    Set objListManagement = cwobjCWComm.InitializeOLEServer(NCID_LIST_MGMT_APPLICATION)
    Set objListManagementOps = objListManagement.ListManagementOps
    Set objListManagement = Nothing
    
     If Len(gobjShared.CmdLineSwitches("GREEN")) > 0 Then
        gbGreenList = True
     Else
        gbGreenList = False
     End If
   '-- -set global flags---
   cwbMainViewRunning = False
   
   '-- ****************************************************************************
   '--  TO DO:
   '--    Add code here to create objects, initialize global variables, etc.
   '-- ****************************************************************************
   
   
   
   
   Exit Sub
   
ErrHandler:
   '-- -show the VB error---
   Call gobjShared.ShowVBError(Err.Number, Err.Description, "Template.Initialize", "Template", vbExclamation)
   
   Exit Sub

End Sub

 Private Sub ICHCSII_AnP_gridDxClick(row As Long, Col As Long)
   
End Sub

Private Function ICHCSII_AnP_SetAnPParent(objAnPParent As Object) As Boolean

   Set gobjAnPParent = objAnPParent
   'Set gobjDxCollection = gobjAnPParent.DiagnosisCollection  'Phil Crowder 10/02/01  Not needed with new AnP interface

End Function

'******************************************************************************************
'  Sub Name: Initialize
'
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
'    robjPatient - Contains the patient used to do patient searches and selections
'*******************************************************************************************
Public Sub ICHCSII_CoreAPI2_Initialize2(ByVal vlMyProgId As Long, _
                                        robjComm As Object, _
                                        robjLogonEnvironment As Object, _
                                        robjConfig As Object, _
                                        robjPatient As Object, _
                                        robjUser As Object)

   On Error GoTo ErrHandler
   
   Set gobjUser = robjUser
   
   Dim objRoles      As ICHCSII_Roles
   Dim objResources  As ICHCSII_Resources
   
   Call Initialize(ByVal vlMyProgId, robjComm, robjLogonEnvironment, robjConfig, robjPatient)
   
   Exit Sub
   
ErrHandler:
   Call gobjShared.ShowVBError(Err.Number, Err.Description, "in frmDiagnosis.Initialize", "Diagnosis !", vbExclamation)
   
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
   
   '-- *****************************************************************
   '--  TO DO:
   '--    Add code here to invoke this OLE Server's properties page.
   '-- *****************************************************************
   
   
   
   Exit Sub
   
ErrHandler:
   '-- -show the VB error---
   Call gobjShared.ShowVBError(Err.Number, Err.Description, "Template.Properties", "Template", vbExclamation)
   
   Exit Sub

End Sub

Private Function ICHCSII_User_CACPresent() As Boolean

End Function

Private Function ICHCSII_User_CACSignText(ByVal PIN As String, ByVal Text As String, Signature As String, Optional ErrCode As Long, Optional ErrMsg As String) As Boolean

End Function

Private Function ICHCSII_User_CACValidatePIN(ByVal PIN As String) As Boolean

End Function

Private Function ICHCSII_User_GetAttribute(AttributeName As String) As String
   
End Function

Private Function ICHCSII_User_GetAttributeNames() As Collection
   
End Function

Private Function ICHCSII_User_HasPrivilege(NCID As String) As Boolean
   
End Function

Private Function ICHCSII_User_HasPrivilegeEx(PrivilegeArea As String, _
                                              PrivilegeType As CHCSII_ClientSecurity.CHCSII_PrivilegeType, _
                                              Optional OtherType As Integer) _
                                              As Boolean
   
End Function

Private Property Get ICHCSII_User_LoginMode() As CHCSII_ClientSecurity.CHCSII_LoginMode

End Property

Private Property Get ICHCSII_User_Resources() As CHCSII_ClientSecurity.ICHCSII_Resources
   
End Property

Private Property Get ICHCSII_User_Roles() As CHCSII_ClientSecurity.ICHCSII_Roles
   
End Property

Private Property Get ICHCSII_User_UserId() As String
   
End Property

Private Property Get ICHCSII_User_UserName() As String
   
End Property

Private Function ICHCSII_User_ValidatePassword(Password As String) As Boolean
   
End Function

Public Sub load_initialize_data()

   With objForm.recSet1

      .addField "ICD", adChar, 25
      .addField "DiagnosisSynonym", adChar, 256
      .addField "SnoID", adChar, 16
      .addField "BlankCol", adChar, 15
      .addField "Prefix", adChar, 2         'scr 21770 3/29/02  Phil Crowder
      .addField "isNode", adBoolean
      .addField "outlineLevel", adInteger
      .addField "typeFont", adChar, 5
      .addField "setIcon", adChar, 8
      .openRecSet adOpenKeyset, adLockOptimistic

   End With
   
   

   colWidths(0) = 800
   colWidths(1) = 3000
   colWidths(2) = 10
   colNames(0) = "ICD"
   colNames(1) = "Diagnosis"
   colNames(2) = "  "

End Sub

Public Sub lockControls(lockType As Boolean, Optional isAnPCall As Boolean = False)
Dim objcontrol As Control

On Error GoTo ErrHnd

    If Not lockType Then
      Screen.MousePointer = vbHourglass
      objForm.ocxDiagnosisGrid.objGeneralGrid.ForeColor = &HC0C0C0
      'objForm.ocxSearchMedcin.vsFlexGridx.ForeColor = &HC0C0C0
    End If
    
    'Phil Crowder 10/02/01    Changed to support new AnP interface
    If Not isAnPCall Then
      If Not lockType Then
        gobjAnPParent.LockAnP "Diagnosis.lockControls"
      Else
        gobjAnPParent.UnlockAnP "Diagnosis.lockControls"
      End If
    End If
    'If Not isAnPCall Then gobjAnPParent.lockControls lockType
    
    objForm.Enabled = lockType
    If lockType Then
      Screen.MousePointer = vbDefault
      objForm.ocxDiagnosisGrid.objGeneralGrid.ForeColor = &H80000008
      'objForm.ocxSearchMedcin.vsFlexGridx.ForeColor = &H80000008
    End If
    
Exit Sub

ErrHnd:
    MsgBox "Error in CHCSII_Diagnosis.frmDiagnosis.LockControls: " & Err.Description
    
End Sub
Public Sub load_template_favs()


   Dim objTemplateOps  As CHCSII_TemplateClient.TemplateOps
   
   Set objTemplateOps = New CHCSII_TemplateClient.TemplateOps
   objTemplateOps.Initialize cwlMyProgId, cwobjCWComm, cwobjLogonEnviron, cwobjConfig, cwobjPatient, gobjUser
   
   Set objForm.ucTemplateNavigation1.TemplateOps = objTemplateOps
   Set objForm.ucTemplateNavigation1.TemplateHistoryOps = objTemplateOps.TemplateHxOps
   
  'jrm:add only for encounters
    objForm.ucTemplateNavigation1.RunningMode = eEncounterTemplatesOnly
    
  'jrm:Add favorites to templatehx
  objForm.ucTemplateNavigation1.TemplateHistoryOps.Add_Favorites objTemplateOps.GetFavoriteTemplates(False, True)
   
End Sub
Public Sub load_template_data()

   Dim i               As Integer
   Dim counter         As Integer
   Dim objTemplateOps  As CHCSII_TemplateClient.TemplateOps
   Dim objTemplate     As CHCSII_TemplateClient.Template
   
   Dim searchSnoIDs As Variant
   
   objForm.recSet1.RemoveAll
   '-- removeAllFromCollection gobjOrderItems
   Set objTemplateOps = New CHCSII_TemplateClient.TemplateOps
   objTemplateOps.Initialize cwlMyProgId, cwobjCWComm, cwobjLogonEnviron, cwobjConfig, cwobjPatient, gobjUser

   Set objTemplate = objTemplateOps.GetCurrentEncounterTemplate(DiagnosisListTemplateSection)
   If objTemplate.TemplateName = "No Encounter Template Loaded" Then ' Is Nothing Then
      'SCR 48941 - Attempt to retreive the default template
      Set objTemplate = objTemplateOps.GetDefaultTemplate()
   End If
   
   If Not objTemplate Is Nothing Then
      With objTemplate.DiagnosisList.Items
         'SCR 9224.. (ldd)
         For i = 0 To .Count - 1
            If convertToSno(CStr(.Keys(i)), True) <> "" Then
               fieldValues(0) = convertToSno(CStr(.Keys(i)), True)
               fieldValues(1) = convertToSno(CStr(.Keys(i)), False)
               fieldValues(2) = CStr(.Keys(i))
               fieldValues(3) = " "
               fieldValues(4) = ""           'scr 21770 3/29/02  Phil Crowder
               fieldValues(5) = False
               fieldValues(6) = 0
               fieldValues(7) = ""
               fieldValues(8) = " "
               objForm.recSet1.fillRowValues fieldValues()
               objForm.recSet1.AddNewRow True
           End If
             
         Next
      End With
   End If
   
   Set objTemplate = Nothing
   Set objTemplateOps = Nothing

End Sub
Public Sub load_template_data_item(templateID As String)

   Dim i               As Integer
   Dim counter         As Integer
   Dim objTemplateOps  As CHCSII_TemplateClient.TemplateOps
   Dim objTemplate     As CHCSII_TemplateClient.Template
   
   Dim searchSnoIDs As Variant
   
   objForm.recSet1.RemoveAll
   '-- removeAllFromCollection gobjOrderItems
   Set objTemplateOps = New CHCSII_TemplateClient.TemplateOps
   objTemplateOps.Initialize cwlMyProgId, cwobjCWComm, cwobjLogonEnviron, cwobjConfig, cwobjPatient, gobjUser
   
   'JT SCR 36572 - Updated code to check IsAppmode
   Set objTemplate = objTemplateOps.GetTemplate(templateID)
   If objTemplate Is Nothing Then
      Exit Sub

   End If

   With objTemplate.DiagnosisList.Items
    'SCR 9224.. (ldd)
      For i = 0 To .Count - 1
         If convertToSno(CStr(.Keys(i)), True) <> "" Then
            fieldValues(0) = convertToSno(CStr(.Keys(i)), True)
            fieldValues(1) = convertToSno(CStr(.Keys(i)), False)
            fieldValues(2) = CStr(.Keys(i))
            fieldValues(3) = " "
            fieldValues(4) = ""           'scr 21770 3/29/02  Phil Crowder
            fieldValues(5) = False
            fieldValues(6) = 0
            fieldValues(7) = ""
            fieldValues(8) = " "
            objForm.recSet1.fillRowValues fieldValues()
            objForm.recSet1.AddNewRow True
        End If
          
      Next
   End With
   
   Set objTemplate = Nothing
   Set objTemplateOps = Nothing

End Sub
Public Function convertToSno(SnoID As String, ICD9 As Boolean) As String

On Error GoTo ErrorHandler

Dim objIC As IConvert
Dim vDXItem As Variant
Dim sLookupId As String
Dim sCodeType As String
Dim objWAM As WAMData.WAM

#If DEBUGON Then
   Dim DR  As DebugRec
   DL.DebugLog Exe, Module, "convertToSno", "Begin", DR
#End If
   
   Set objIC = New IConvert
   
   'Parse out the string and see if we have info on the type of code value we are using
   vDXItem = Split(SnoID, ",")
   
   If UBound(vDXItem) = 0 Then
      'Item does not have a code designator so assume its a Sno id
      sLookupId = SnoID
      sCodeType = "M"
   ElseIf UBound(vDXItem) > 0 Then
      sLookupId = vDXItem(0)
      sCodeType = vDXItem(1)
   End If
    
   'Now perform the translation
   Select Case True
   Case sCodeType = "M" And ICD9
      'get the icd9 code
      convertToSno = objIC.SnoICD9(sLookupId)
   Case sCodeType = "M" And Not ICD9
      'get the Sno description
      convertToSno = objIC.SnoDescription(sLookupId)
   Case sCodeType = "I" And ICD9
      'convert to ICD9 which we already have so just return it
      convertToSno = sLookupId
   Case sCodeType = "I" And Not ICD9
      'Need to lookup description from WAM
      Set objWAM = gobjShared.CHCSConnection.WAMFactory(True)
      With objWAM.Validator.Lookups.Item(elICD9).Recordset
         .MoveFirst
         .Find "CODE ='" & sLookupId & "'"
         If Not .EOF Then convertToSno = .Fields("Description").value
      End With 'With
   Case Else
      'Unrecognized conversion
      convertToSno = ""
   End Select
     

#If DEBUGON Then
   DL.DebugLog Exe, Module, "convertToSno", "End", DR
#End If

ErrorHandler:
   'Clean Up Objects
   Set objWAM = Nothing
   Set objIC = Nothing
   
   'show any errors
   If CBool(Err) Then Call gobjShared.ShowVBError(Err.Number, Err.Description, "clsDiagnosis.convertToSno", "Diagnosis", vbCritical)
   
End Function

'JT BEGIN SCR 34406 - Handle PGUI Mode and setup field array using a shared routine
Private Sub SetProbFldVals(ByRef objProblem As Object)
'Setup the field value array for problem records so they the data can be added to the recordset

   With objProblem
      fieldValues(0) = .ICD9Code
      fieldValues(1) = .Problem
      'JT SCR 36572 - Updated code to check IsAppmode
      fieldValues(2) = IIf(gobjShared.IsAppMode(modeCHCSI_GUI), .ICD9Code, .SnoID)
      fieldValues(3) = "  "
      fieldValues(4) = .Prefix     'scr 21770 3/29/02  Phil Crowder
      fieldValues(5) = False
      fieldValues(6) = 0
      fieldValues(7) = " "
      fieldValues(8) = " "
   End With
   
End Sub
'JT END SCR 34406

Public Sub load_problemlistacute_data()

   Dim i        As Integer
   Dim counter  As Integer
      
   counter = 0
   'Phil Crowder 10/02/01   Changed to support new AnP interface
   Set gobjProblemListAcute = gobjAnPParent.RetrieveAcuteProblemList.Problems.Filter
   
   objForm.recSet1.RemoveAll
   
   
   With objForm.recSet1

      For i = 1 To gobjProblemListAcute.Count
            If gobjProblemListAcute.Item(i).chronicityNCID = NCID_ChronicityAcute Then   'SCR-8116 Don't display deleted problems
                SetProbFldVals gobjProblemListAcute.Item(i)    'JT SCR 34406 - Handle PGUI Mode and setup field array using a shared routine
                .fillRowValues fieldValues()
                .AddNewRow True
            End If
      Next 'i

   End With
   
   
End Sub
Public Sub load_problemlist_data()

   Dim i        As Integer
   Dim counter  As Integer
      
   counter = 0
   'Phil Crowder 10/02/01   Changed to support new AnP interface
   Set gobjProblemListAcute = gobjAnPParent.RetrieveAcuteProblemList.Problems.Filter
   
   objForm.recSet1.RemoveAll
      
   With objForm.recSet1

      For i = 1 To gobjProblemListAcute.Count
         'jrm:Removed If chronicity statement because we want all the problems.
         SetProbFldVals gobjProblemListAcute.Item(i)     'JT SCR 34406 - Handle PGUI Mode and setup field array using a shared routine
         .fillRowValues fieldValues()
         .AddNewRow True
      Next 'i

   End With
      
End Sub
Public Sub load_listmanagement_data()

    Dim i       As Integer
    Dim bool    As Boolean
    Dim bUser   As Boolean
    Dim bClinic As Boolean
    Dim iClinic As Integer
   
    Dim objDx As Entry
    
  
   objForm.recSet1.RemoveAll    'SCR 9763
   
   With objForm.recSet1
        fieldValues(0) = LBL_USER_FAV
        fieldValues(1) = " "
        fieldValues(2) = " "
        fieldValues(3) = " "
        fieldValues(4) = ""
        fieldValues(5) = True
        fieldValues(6) = 0
        fieldValues(7) = " "
        fieldValues(8) = " "
        '-- fieldValues(8) = "leafIcon"
        .fillRowValues fieldValues()
        .AddNewRow True
        For Each objDx In objListManagementOps.User.Diagnosis
           fieldValues(0) = objDx.Code
           fieldValues(1) = objDx.Description
           fieldValues(2) = objDx.SnoID
           fieldValues(3) = LBL_USER_FAV
           fieldValues(4) = ""
           fieldValues(5) = False
           fieldValues(6) = 1
           fieldValues(7) = " "
           fieldValues(8) = CBool(objDx.NotFromMedcin)   'JT SCR 42399 - track if we have a valid medcin id
           '-- fieldValues(8) = "leafIcon"
           .fillRowValues fieldValues()
           .AddNewRow True
        Next
        fieldValues(0) = LBL_CLINIC_FAV
        fieldValues(1) = " "
        fieldValues(2) = " "
        fieldValues(3) = " "
        fieldValues(4) = ""
        fieldValues(5) = True
        fieldValues(6) = 0
        fieldValues(7) = " "
        fieldValues(8) = " "
        '-- fieldValues(8) = "leafIcon"
        .fillRowValues fieldValues()
        .AddNewRow True
        For Each objDx In objListManagementOps.Clinic.Diagnosis
           fieldValues(0) = objDx.Code
           fieldValues(1) = objDx.Description
           fieldValues(2) = objDx.SnoID
           fieldValues(3) = LBL_CLINIC_FAV
           fieldValues(4) = ""
           fieldValues(5) = False
           fieldValues(6) = 0
           fieldValues(7) = " "
           fieldValues(8) = CBool(objDx.NotFromMedcin)      'JT SCR 42399 - track if we have a valid medcin id
           '-- fieldValues(8) = "leafIcon"
           .fillRowValues fieldValues()
           .AddNewRow True
        Next
   End With
  
theEnd:
  Set objDx = Nothing
  
   
End Sub

Public Sub load_problemlistchronic_data()

   Dim i        As Integer
   Dim counter  As Integer
   Const NCID_ACTIVE As String = "1024"
   
   counter = 0
   'Phil Crowder  10/02/01  Changed to support new AnP interface
   Set gobjProblemListChronic = gobjAnPParent.RetrieveChronicProblemList.Problems.Filter
'   gobjAnPParent.retrieveChronicProblemList
'   Set gobjProblemListChronic = gobjAnPParent.objProbListChronic
   objForm.recSet1.RemoveAll
   
   
   With objForm.recSet1
'Public Const NCID_ChronicityAcute As String = "11312"
'Public Const NCID_ChronicityChronic As String = "30191"
      For i = 1 To gobjProblemListChronic.Count
        If gobjProblemListChronic.Item(i).chronicityNCID = NCID_ChronicityChronic Then   'SCR-8116 Don't display deleted problems
            SetProbFldVals gobjProblemListChronic.Item(i)   'JT SCR 34406 - Handle PGUI Mode and setup field array using a shared routine
            .fillRowValues fieldValues()
            .AddNewRow True
        End If
      Next 'i

   End With
   
   
End Sub

Public Sub load_finalize_datagrid()
      
   With objForm.ocxDiagnosisGrid

      .adoPrimaryRecordSet = objForm.recSet1.adoPrimaryRecordSet
      .setGridColWidths colWidths()
      .setGridColNames colNames()
      .gridDataRefresh
      '-- ocxDiagnosisGrid.setReadOnlyCols (1)
      '-- ocxDiagnosisGrid.setReadOnlyCols (2)
      '-- ocxDiagnosisGrid.setReadOnlyCols (4)
      .gridAttributes
      .setHiddenCol (2)
      .setHiddenCol (3)
      .setHiddenCol (4)
      .setHiddenCol (5)
      'SCR 12537 - User/Clinic Favorites (ldd)
      .setGridBarType flexExSort

   End With

End Sub

 Sub load_searchedlist_data()

 End Sub
