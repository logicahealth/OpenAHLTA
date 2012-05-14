VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Labs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit   '-- $Reformatted by FormatVB, Ver. 1.6b, on 02-23-00 at 10:50:53

'-- Private DR As DebugRec
'-- Public DL As cDebugLog
'-- Public Const Exe = "OrderSets"
Private Const Module = "Dental.cls"

Private Enum mnumDCT_DentalComponentTab
   lDCT_BigTooth = 0
   lDCT_Notes = 1
   lDCT_Labs = 2
   lDCT_DataViews = 3
End Enum

Implements ICHCSII_CoreAPI2
Implements ICHCSII_CoreAPI3
Implements ICHCSII_User
Implements ICHCSII_Dental

'------------------------------------------------------------
'--- PROCEDURE: ARM_InitAllTransactionsForDental_Labs
'---
'--- PURPOSE:   Called from Class_Initialize. Replacing all
'---            occurrences of 'MsgBox' with 'MsgBxARMd'
'---
'--- HISTORY:   RWalker  SCR 28295 10/8/02  COMMON TO ALL ARM-SCRs
'------------------------------------------------------------
Public Function ARM_InitAllTransactionsForDental_Labs() As Boolean
    Static bDoneThis    As Boolean
    Dim objARMFactory   As New ARM.cARMFactory
    
    On Error GoTo EH
    
    '--- Init Tx classes only once
    '--- Init routine should be called from first top level creatable coclass
    '--- It can be called multiple times, but will actually run only once.
    '--- Instrument all creatable classes in case dll's multiple entry
    '--- points are actually used
    If bDoneThis Then
        Exit Function
    Else
        bDoneThis = True
    End If

    Set goARM = objARMFactory.CreateARMApp("")  'BMStop
    
    On Error Resume Next
    
    '--------------------------------------------------------------
    '--- MESSAGE MAP REGISTERING TRANSACTIONS BY NAME W/ ARM SERVER
    '--------------------------------------------------------------
    
    ARM_InitAllTransactionsForDental_Labs = True
    Exit Function

EH:
    ARM_InitAllTransactionsForDental_Labs = False
End Function

Private Function IChild_SetParent(Parent As IParent) As Boolean
   Set gobjEnc = Parent
   IChild_SetParent = True
End Function
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
         QueryView = False
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

Dim sMessage As String
'Dim sBlankString As String

'Dental Labs form visibility status
'- False, if Dental Labs form is invisible
'- True, otherwise
Dim bLabFormVisible As Boolean

'vsElasticBasic visibility status
Dim bElasticFormVisible As Boolean

   On Error GoTo ErrHandler
          
   '-- -perform actions based on the "view type"---
   If viViewType = cwiREGULAR_VIEW Then
      '-- -set global flags---
      cwbFirstRegularResize = True
      cwbMainViewRunning = True
                 
      Call InitLabCase
      
      'OpenLabCase is getting called on two occassions:
      '- any patient is selected when initial CHCSII application starts up (previous PatientUnitID = 0)
      '- current patient is changed
      
      Call OpenLabCase(bLabFormVisible, bElasticFormVisible, sMessage)
                
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
      '-- glOleMenuItem = cwobjCWComm.AddMenuItem("&Ole Server Menu Item", ghTemplate, cwlMyProgId, "Template menu item")
      
      '-- ---------------------------------------------------------------------------------------
      '--  Reset the "parent window" for this OLE server's "main" regular form.  This is the main
      '--    "trick" that makes the OLE server look like it is part of the CW Core.
      '--
      '--  NOTE: This action invokes the main form's "Load" event, but the form is NOT shown.
      '--        Make sure that you DO NOT perform a "Me.Show" in the form's "Load" event, as this
      '--        will mess up the "SetParent" API call.  Make sure that you DO NOT "Show" the
      '--        "main" regular form in this method, that will happen in the "ResizeView" method.
      '-- ---------------------------------------------------------------------------------------
            
      SetParent frmLabs.hWnd, vlParenthWnd
      Set frmLabs.ParentObject = Me
      
      '<<< the following action invokes the main form's "Load" event,
      '<<< but the form is NOT shown
      '-- -set parameter---
      
      '--- The following statement is commented out
      '--- (frmMessage is "Loading data... Please wait")
      'Unload frmMessage
      
      rhMainForm = frmLabs.hWnd
      
      
      If Not bLabFormVisible Then      'bLabFormVisible = False
         'LoadLabCase failed
         '(LabCase was not found or LabCase can not be created)
         'then vsDentalTab should be invisible and vsElastic container is visible
         'with message on the front
         frmLabs.vsDentalLabTab.Visible = bLabFormVisible
         frmLabs.vsElasticBasic.Visible = bElasticFormVisible
         
         frmLabs.vsDentalLabTab.ZOrder 1
         frmLabs.vsElasticBasic.ZOrder 0
         
         '==============================
'         Dim sBlankString As String
'         sBlankString = "                                        "
'         If giLabCaseStatus = gcLabCaseLoadFailed Or giLabCaseStatus = gcNewLabCaseCreateFailed Or giLabCaseStatus = gcLabTreatmentsUnavailable Then
'
'            If Len(sMessage) = 0 Then
'            sMessage = sBlankString & _
'                        "      DENTAL LABS DATA ARE NOT AVAILABLE FOR CURRENT ENCOUNTER. " & _
'                        vbCrLf & vbCrLf & sBlankString & _
'                        "                     DATA ACCESS FAILED."
'            End If
'         End If
         Dim sBlankString As String
         sBlankString = "                                        "
         If giLabCaseStatus = gcLabCaseLoadFailed Or giLabCaseStatus = gcNewLabCaseCreateFailed Or giLabCaseStatus = gcLabTreatmentsUnavailable Then
            
            sMessage = sBlankString & _
                        "      This Dental Lab Order module is not available in an encounter until there is a prosthetics Treatment in-progress."
         End If

         '==============================
         
         frmLabs.vsElasticBasic.Caption = sMessage
                  
      End If
           
      'If giLabCaseStatus = gcLabCaseLoadFailed Or giLabCaseStatus = gcNewLabCaseCreateFailed Or giLabCaseStatus = gcLabTreatmentsUnavailable Then
         'MsgBxARMd " NO LAB DATA AVAILABLE FOR THIS ENCOUNTER ", vbOKOnly, "Dental Labs"
         'OpenView = False
                  
         'Exit Function
      'End If
      
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
   Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental Labs.OpenView", "Dental Labs", vbCritical)
   
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

On Error GoTo ErrHandler

   If viViewType = cwiALL_VIEWS Then
      '-- -close the "main view"---

      If cwbMainViewRunning Then
         '-- -set global flags---
         cwbMainViewRunning = False
         
         '-- -this indicates that the CW Core is closing the "view", NOT this OLE Server---
         cwbCloseViewMyself = False
         
         If CloseLabCase Then
            Unload frmLabs
         Else
            MsgBxARMd ("Unable to close CHCSII Dental Labs. Process failed.")
         End If
         
         '-- -close ALL forms that are associated with all "views"---
'        If gobjEnc Is Nothing Then

'        Else
'        If Not gobjEnc.StartSave() Then
'          MsgBxARMd "Your lock on the Dental section of the current encounter has been taken away by another user."
'          Exit Sub
'        End If
'          Call gobjEnc.SetRTF(frmDental.txtEncSummary.TextRTF)
'        gobjEnc.EndSave
'         Unload frmDental
'         If Not gobjEnc Is Nothing Then
'                gobjEnc.ClosedView
'          End If
'        End If
      End If

   ElseIf viViewType = cwiREGULAR_VIEW Then
      '-- -close the "main view"---

      If cwbMainViewRunning Then
         '-- -set global flags---
         cwbMainViewRunning = False
         
         '-- -this indicates that the CW Core is closing the "view", NOT this OLE Server---
         cwbCloseViewMyself = False
         
         If CloseLabCase Then
            Unload frmLabs
            Set frmLabs = Nothing
         Else
            MsgBxARMd ("Unable to close CHCSII Dental Labs. Process failed.")
         End If
         
         '-- -close ALL forms that are associated with all "views"---
'        If gobjEnc Is Nothing Then
'        Else
'        If Not gobjEnc.StartSave() Then
'          MsgBxARMd "Your lock on the Dental section of the current encounter has been taken away by another user."
'          Exit Sub
'        End If
'          Call gobjEnc.SetRTF(frmDental.txtEncSummary.TextRTF)
'        gobjEnc.EndSave
'         Unload frmDental
'         If Not gobjEnc Is Nothing Then
'                gobjEnc.ClosedView
'          End If
'        End If
         
      End If

   End If
   
Exit Sub

ErrHandler:
   '-- -show the VB error---
   Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental Labs.CloseView", "Dental Labs", vbCritical)
      
   Exit Sub


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
    Select Case vlMenuId
   
        '******************************************************************************
        '  TO DO:
        '    Take actions needed to process this menu item.
        '******************************************************************************
        'WilsonBegin on 04/11/01, 10/31/01
        Case glMenuPrintID
            
            'KDunne 4/15/2002
            'SCR 13107
            'Previously, the value in txtGenProcCaseNo was being passed in for the lab case #.
            'Problem was with new lab cases, an (N) was being appeneded to the lab case # in the
            'text box, when we did a search on {labcase#}(N), we incorrectly returned no records.
            'Solution was to pass in the the LabCaseNo from the global LabCase object, and not
            'from the text box.
            '
            With guCurEncInfo
            'SCR 40565 RWalker 11/07/2003 Use the lab case provider info not the encounter provider
                'Call PrintDD2322(gobjLabCase.LabCaseNo, .PatientUnitID, .EncounterFacilityNCID, _
                    .ClinicNCID, .ClinicName, .ProviderID, .ProviderFullName)
                Call PrintDD2322(gobjLabCase.LabCaseNo, .PatientUnitID, .EncounterFacilityNCID, _
                    .ClinicNCID, .ClinicName, gobjLabCase.ProviderID, gobjLabCase.ProviderName)
            End With
        'WilsonEnd
        Case Else   '--do nothing--
    End Select
   
   Exit Sub
   
ErrHandler:
   '-- -show the VB error---
   Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental_Labs.MenuItemSelected", "Template", vbExclamation)
Err.Raise Err.Number, Err.Source, Err.Description
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
            
            If CloseLabCase Then
               Unload frmLabs
            Else
               MsgBxARMd ("Unable to close CHCSII Dental Labs. Process failed.")
            End If
      
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
         frmLabs.WindowState = vbNormal
         
         '-- -hide the form to so that annyoing "screen blinking" will not be seen while controls are moved and resized---
         frmLabs.Visible = False
         
         '-- -resize form to fill entire "client area", NOTE: This generates a "resize" event---
         frmLabs.Move cwlLeft, cwlTop, cwlWidth, cwlHeight
         
         frmLabs.Visible = True
         'frmLabs.vsDentalLabTab.Visible = True
         'frmLabs.vsDentalLabTab.ZOrder 0
         'frmLabs.vsElasticBasic.Visible = True
         'Call Mtd_bDentalTabChange(0, 2)
      Else                             '--first "regular" resize, need to perform "initailization" OrderSetss--
         '-- -reset flag so this code will not be executed again---
         cwbFirstRegularResize = False
         
         '-- --------------------------------------------------------------------------
         '--  NOTE: Here is where you "Show" the "Regular View" form for the FIRST time.
         '-- --------------------------------------------------------------------------
          'If Not gobjEnc Is Nothing Then
          '  glMenuSave = gobjEnc.AddMenuItem(0, "&Save", "save.bmp", "Save Dental")
          '  glMenuNext = gobjEnc.AddMenuItem(0, "D&isposition", "Next.bmp", "Go to the Disposition screen")
            frmLabs.Display
          'Else
          '  CloseView cwiREGULAR_VIEW
          '  Exit Sub
          'End If
      End If

   End If
   
   Exit Sub
   
ErrHandler:
   '-- -show the VB error---
   Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental_Labs.ResizeView", "Dental", vbExclamation)
Err.Raise Err.Number, Err.Source, Err.Description
   Exit Sub

End Sub

Private Sub Class_Initialize()

#If DEBUGON Then
    Set DL = GetObject("", "dlog.cdebuglog")
    DL.debugStart
    Dim DR  As DebugRec

    DL.DebugLog Exe, Module, "Class_Initialize", "Begin", DR
#End If
    '-- -create shared object used to show VB errors, etc.---
    Set gobjShared = New CWShared
   
  
    '-- -set the help file for this application---
    App.HelpFile = gobjShared.CWHelpFile
   
    '--- SCR 28295 RWalker 10/8/02 Instantiate ARM / Init transactions
    Call ARM_InitAllTransactionsForDental_Labs
   
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

    If Not gobjShared.UseLocalDB And gobjShared.CmdLineSwitches("NOSTOREDPROC") <> "NOSTOREDPROC" Then
        gblnUseStoredProc = True
    End If
    
End Sub

Private Sub Class_Terminate()

#If DEBUGON Then
   Dim DR  As DebugRec

   DL.DebugLog Exe, Module, "Class_Terminate", "Begin", DR
#End If
   '-- -release reference to objects---
   '-- Set gobjAnPCollection = Nothing
   Set gobjShared = Nothing
   Set cwobjCWComm = Nothing
   Set cwobjLogonEnviron = Nothing
   Set cwobjConfig = Nothing
   Set cwobjPatient = Nothing
   Set gobjUser = Nothing
   Set gobjEnc = Nothing
   Set gobjDental = Nothing
   
   '-- ********************************************************************
   '--  TO DO:
   '--    Add code here to release reference to all other global objects.
   '-- ********************************************************************
      
   '-- --------------------------------------
   '--  If debugging, make sure to "log off".
   '-- --------------------------------------
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
        
   End Select
   
   Exit Sub
   
ErrHandler:
   '-- -show the VB error---
   Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental_Labs.Message", "Labs", vbExclamation)
   
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
   '-- -set global flags---
   cwbMainViewRunning = False
   
   '-- ****************************************************************************
   '--  TO DO:
   '--    Add code here to create objects, initialize global variables, etc.
   '-- ****************************************************************************
          
    '---initialize Current Patient variables ---
    guCurEncInfo.PatientUnitID = robjPatient.UnitNumber
          
    guCurEncInfo.ProviderID = cwobjLogonEnviron.UserNCID
    
   Exit Sub
   
ErrHandler:
   '-- -show the VB error---
   Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental_Labs.Initialize", "Labs", vbExclamation)
   
   Exit Sub

End Sub


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
   Call gobjShared.ShowVBError(Err.Number, Err.Description, "in frmLabs.Initialize", "Labs !", vbExclamation)
   
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
   Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental_Labs.Properties", "Labs", vbExclamation)
   
   Exit Sub

End Sub



Private Function ICHCSII_CoreAPI3_CloseView2(ByVal viViewType As Integer) As Boolean
   ICHCSII_CoreAPI3_CloseView2 = True
End Function

Private Property Get ICHCSII_Dental_EncounterDental() As Variant

End Property

Private Property Set ICHCSII_Dental_EncounterDental(RHS As Collection)

End Property

Private Property Get ICHCSII_Dental_EncounterDentalLabs() As Variant

End Property

Private Property Set ICHCSII_Dental_EncounterDentalLabs(RHS As Collection)

End Property

Private Property Get ICHCSII_Dental_EncounterDentalNotes() As Variant

End Property

Private Property Set ICHCSII_Dental_EncounterDentalNotes(RHS As Collection)

End Property

Private Property Get ICHCSII_Dental_EncounterDx() As Variant

End Property

Private Property Set ICHCSII_Dental_EncounterDx(RHS As Collection)

End Property

Private Property Get ICHCSII_Dental_EncounterOps() As Variant

End Property

Private Property Get ICHCSII_Dental_EncounterTeeth() As Variant

End Property

Private Property Set ICHCSII_Dental_EncounterTeeth(RHS As Collection)

End Property

Private Property Get ICHCSII_Dental_EncounterTx() As Variant

End Property

Private Property Set ICHCSII_Dental_EncounterTx(RHS As Collection)

End Property

Private Function ICHCSII_Dental_SetDentalParent(objDentalParent As Object) As Boolean
    gsDentalEncounterID = objDentalParent.psDentalEncounterID
    Set gobjDental = objDentalParent
End Function

Private Function ICHCSII_Dental_SetEncounterDentalBT(objBigToothOps As Object) As Boolean

End Function

Private Function ICHCSII_Dental_SetEncounterDentalDataViews(objDataViewsOps As Object) As Boolean

End Function

Private Function ICHCSII_Dental_SetEncounterDentalFM(objFullMouthOps As Object) As Boolean

End Function

Private Function ICHCSII_Dental_SetEncounterDentalLabs(objLabOps As Object) As Boolean

End Function

Private Function ICHCSII_Dental_SetEncounterDentalNotes(objNoteOps As Object) As Boolean

End Function

Private Function ICHCSII_Dental_SetEncounterOps(objEncounterOps As Object) As Boolean
  Dim objEncounter As CHCSIIEncounterOps.Encounter
  'WilsonBegin on 4/11/01
  Set gobjEnc = objEncounterOps
  'WilsonEnd

  If Not gobjEnc Is Nothing Then
    Set objEncounter = gobjEnc.CurrentEncounter
  
    With objEncounter
        guCurEncInfo.ProviderFullName = .PrimaryProvider.FullName
        guCurEncInfo.EncounterFacilityNCID = .FacilityNCID
        guCurEncInfo.EncounterNumber = .EncounterID
        
        'Wilson added for DD2322
        guCurEncInfo.ClinicNCID = .ClinicNCID
        guCurEncInfo.ClinicName = .Clinic
        
    End With 'objEncounter
  End If

End Function

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

''Private Function IChild_GetAutoCiteNarrative(TestType As String, NumPeriods As Integer, Period As CHCSIIEncounterOps.AC_Period) As String

''End Function

''Private Function IChild_SetParent(Parent As CHCSIIEncounterOps.IParent) As Boolean
''   Set gobjEnc = Parent
''   IChild_SetParent = True
''End Function

Public Function Mtd_bDentalTabChange(ByVal lFrom As Long, ByVal lTo As Long) As Boolean
'''************************************************************************
'''Called By:     Dental Parent
'''
'''Arguments:
'''lFrom          Identifier of the dental the currently active dental module
'''lTo            Identifier of the dental the about to became active
'''
'''
'''Description:   if lFrom equal to your Tab
'''                  Focus is changing from your tab to lTo Module
'''                  if you ready to release focus: save current data...
'''                     Return true
'''                  else
'''                     return false
'''               Elseif lTo equal to your Tab
'''                  Focus is changing from lTo  tab to your Module
'''                  if you ready to recive focus: Objects are initialized...
'''                     Return true
'''                  else
'''                     return false
'''               Else
'''                  Note: May not be called
'''                  Return true
'''
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If successful  - true
'''               Else           - false
'''
'''Programmer:
'''************************************************************************
Dim bResult          As Boolean

Dim sMessage As String

'vsDental Labs form visibility status
Dim bLabFormVisible As Boolean

'vsElasticBasic visibility status
Dim bElasticFormVisible As Boolean

bResult = False
On Error GoTo ErrHnd
   
   'release focus - save dental labs data
  If lFrom = lDCT_Labs Then
      If giLabCaseStatus <> gcLabCaseInit Then
         Me.SaveView
         If CloseLabCase Then
            bResult = True
         Else
            bResult = False
         End If
      
      End If
                  
      frmLabs.ResetGUI
   ElseIf lTo = lDCT_Labs Then
      
      cwobjCWComm.setStatusText gcMsgStatusBar1
   
      bResult = True
      
      '=============================================
      'If giLabCaseStatus = gcLabCaseInit Then
      If giLabCaseStatus <> gcLabCaseUndefined Then
      '=============================================
         'If LabCase is not loaded yet (giLabCaseStatus initial value is 0)
         'LabCase initializing will be performed in OpenView method
         'and the following statements in DentalTabChange will be skipped
         
         Call InitLabCase
         
         Call OpenLabCase(bLabFormVisible, bElasticFormVisible, sMessage)
         'if LoadLabCase was not successful (LabCase was not found or LabCase can not be created)
         'then vsDentalTab should be invisible and vsElastic container is visible
         'with message on the front
         
         frmLabs.vsDentalLabTab.Visible = bLabFormVisible
         
         If bLabFormVisible Then
           frmLabs.vsDentalLabTab.ZOrder 0
           frmLabs.vsElasticBasic.ZOrder 1
         Else
           frmLabs.vsDentalLabTab.ZOrder 1
           frmLabs.vsElasticBasic.ZOrder 0
           
           '==========06/26/01==============
           'Labs Form should be invisible
           'it means that tab change is not ready
         ' MsgBxARMd " NO LAB DATA AVAILABLE FOR THIS ENCOUNTER ", vbOKOnly, "Dental Labs"
         '  bResult = False
         '  cwobjCWComm.setStatusText " "
           
         '   Exit Function
           '================================
         End If
         
         'frmLabs.vsElasticBasic.Visible = Not bElasticFormVisible
         'frmLabs.vsElasticBasic.ZOrder 1
         
         '==============================
'         Dim sBlankString As String
'         sBlankString = "                                        "
'         If giLabCaseStatus = gcLabCaseLoadFailed Or giLabCaseStatus = gcNewLabCaseCreateFailed Or giLabCaseStatus = gcLabTreatmentsUnavailable Then
'
'            If Len(sMessage) = 0 Then
'               sMessage = sBlankString & _
'                           "      DENTAL LABS DATA ARE NOT AVAILABLE. " & _
'                           vbCrLf & vbCrLf & sBlankString & _
'                           "                    DATA ACCESS FAILED."
'            End If
'
'         End If
         '==============================
         
         
         frmLabs.vsElasticBasic.Caption = sMessage
      End If

      cwobjCWComm.setStatusText " "
      
   Else
      bResult = True
   End If
   
   '------------------------
   'WilsonBegin on 4/11/01
   '------------------------
   'KDunne 8/12/2002
   'SCR 26478 - After deleting a menu item, we must refresh the menus.
   cwobjCWComm.DeleteMenuItem (glMenuPrintID)
   glMenuPrintID = 0
   cwobjCWComm.Message cwiGENERIC, "Refresh", NCID_CW_APPLICATION, NCID_DENTAL_LABS_APPLICATION
   'End 26478
   If lTo = lDCT_Labs Then
      glMenuPrintID = gobjEnc.AddMenuItem(0, "&DD2322", "Print.bmp", "Preview DD2322")
   End If
   'WilsonEnd
   
   'bResult = True
   Mtd_bDentalTabChange = bResult

   
Exit Function
ErrHnd:
      '-- -show the VB error---
   Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental Labs.mtd_bDentalTabChange", "Dental Labs", vbCritical)

   cwobjCWComm.setStatusText " "
   Mtd_bDentalTabChange = False

Exit Function
End Function


'********************************************************************************
' Name:         PrintDD2322
' Description:  Calling Form's PrintDD2322 Methord
' Parameter:    N/A
' Called By:    MenuItemSelected()
' Histary:
'   Developer:  Date:       Description
'   ----------  ---------   -----------------------------------------------------
'   KDunne      01/10/03    Query was only returning labs on tx's created by the
'                           current provider. This is different than the query that
'                           creates the labs display. Spoke with SA who said this was
'                           incorrect. Modified query.
'   Wilson Sun  11/14/01    Added define oDAS object
'   Wilson Sun  10/29/01    Added Parameters and using SQL instead of Lab grid. SCR 17423
'   Wilson Sun  10/25/01    Display ADA_CODE in block 26. SCR 13107
'   Wilson Sun  07/16/01    Generate recodset from vsFlexGridTx
'                           for DD2322 Subreport data (SCR: 13107)
'   Wilson Sun  04/10/01    Created
'********************************************************************************
Public Function PrintDD2322(ByVal psLabCaseNumber As String, _
                       ByVal psPatientUnitNumber As String, _
                       ByVal psFacilityNCID As String, _
                       ByVal psClinicNCID As String, _
                       ByVal psClinicName As String, _
                       ByVal psProviderNCID As String, _
                       ByVal psProviderName As String) As Boolean

    Dim bPreview As Boolean
    Dim i As Long
    Dim sSql As String
    Dim sSqlCommon As String
    Dim objForms As Object
    Dim rsSubReport As ADODB.Recordset
    Dim oDAS As ICHCSII_DAS
    
    On Error GoTo ErrHandler
    
    PrintDD2322 = False
    Set objForms = CreateObject("CHCSII_Forms.DentalForms")
   
    Set objForms.Comm = cwobjCWComm
    
    If objForms Is Nothing Then
       MsgBxARMd "Could not create CHCSII_Forms object.", vbExclamation, "Print DD2322"
       Exit Function
    End If
   
    bPreview = True
   
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    'Create DD2322 Subreport record set based on vsFlexGridTx values
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    Set rsSubReport = New ADODB.Recordset
'    rsSubReport.CursorLocation = adUseClient
'    rsSubReport.Fields.Append "TOOTH_LOCATION", adBSTR
'    rsSubReport.Fields.Append "DESCRIPTION", adBSTR
'    rsSubReport.Open
'
'    With frmLabs.vsFlexGridTx
'        For i = 1 To gobjLabCase.LabTXs.Count
'            rsSubReport.AddNew
'            rsSubReport.Fields!TOOTH_LOCATION = .TextMatrix(i, gcClmnTxToothLocation)
'            rsSubReport.Fields!Description = .TextMatrix(i, gcClmnTxADACode) & "   " & .TextMatrix(i, gcClmnTxDescription) & "  "
'        Next i
'    End With
    
    'KDunne 9/11/2002
    'SCR 27346 - Save Labs first
    SaveLabCaseView
    
    Set goChcsiiConn = New CHCSII_CONN.Conn
    Set oDAS = goChcsiiConn.CHCSII_DAS
'    If gobjShared.UseLocalDB Then
'        Set oDAS = New GEMS_DAS
'    Else
'        Set oDAS = New CHCSII_DAS
'    End If

    'SCR 40565 RWalker 11/24/2003
    'Added for retrieving provider's signature block line 2 info
    Dim SigBlockLine2   As String
    Call cwobjConfig.GetValue(NCID_ENCOUNTER_APPLICATION, "SigBlock", "Line2", SigBlockLine2)
    If SigBlockLine2 <> "" Then
        psProviderName = psProviderName & " - " & SigBlockLine2
    End If
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Create DD2322 Subreport record set based on SQL
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'KDunne 11/5/2003
    'SCR 34328 - Removed DX Status from the where clause of the sql statement.
    '            We were not bringing back dx's from previous lab cases.
    ' "AND (d.Status_ID = 101 OR d.Status_ID = 102) " & vbCrLf & _
    '--select
    sSql = "SELECT b.Lab_Case_Number, " & vbCrLf & _
        goChcsiiConn.FormatDateForDBSelect("c.Tx_Date_Ordered", DD_MON_YYYY) & " as Date_Ordered, " & vbCrLf & _
        goChcsiiConn.FormatDateForDBSelect("c.Tx_Date_Required", DD_MON_YYYY) & " as Date_Required, " & vbCrLf & _
        "c.Quantity, f.Tooth_Name, h.ADA_Code, " & vbCrLf & _
        "h.ADA_Definition, b.Comments "
    '--from
    sSql = sSql & vbCrLf & _
        "FROM DENTAL_ENCOUNTER a, " & vbCrLf & _
        "     LAB_CASE b, " & vbCrLf & _
        "     LAB_TX_SPECIFICATIONS c, " & vbCrLf & _
        "     DENTAL_ENCOUNTER_TX d, " & vbCrLf & _
        "     DENTAL_ENCOUNTER_TXS_SS e, " & vbCrLf & _
        "     DENTAL_ENCOUNTER_TEETH f, " & vbCrLf & _
        "     DENTAL_ENC_SURFACE_SET g, " & vbCrLf & _
        "     TX_ADA_CODE_DEFINITION h "
    '--where
    sSql = sSql & vbCrLf & _
        "WHERE a.Patient_Unit_ID = " & psPatientUnitNumber & vbCrLf & _
        "AND   a.Encounter_Facility_NCID = " & psFacilityNCID & vbCrLf & _
        "AND   a.Dental_Encounter_ID = b.Dental_Encounter_ID " & vbCrLf & _
        "AND b.Lab_Case_ID = c.Lab_Case_ID " & vbCrLf & _
        "AND c.Tx_Code_ID = d.Tx_Code_ID " & vbCrLf & _
        "AND d.Tx_Code_ID = e.Tx_Code_ID " & vbCrLf & _
        "AND h.Tx_ADA_Code_ID = d.Tx_ADA_Code_ID " & vbCrLf & _
        "AND h.LabFlag = -1 " & vbCrLf & _
        "AND f.Dental_Encounter_Teeth_ID = g.Dental_Encounter_Teeth_ID " & vbCrLf & _
        "AND e.Surface_Set_ID = g.Surface_Set_ID " & vbCrLf & _
        "AND b.Lab_Case_Number = '" & psLabCaseNumber & "'"
    '--Union select
    sSql = sSql & vbCrLf & _
        "UNION ALL " & vbCrLf & _
        "SELECT b.Lab_Case_Number, " & vbCrLf & _
        goChcsiiConn.FormatDateForDBSelect("c.Tx_Date_Ordered", DD_MON_YYYY) & " as Date_Ordered, " & vbCrLf & _
        goChcsiiConn.FormatDateForDBSelect("c.Tx_Date_Required", DD_MON_YYYY) & " as Date_Required, " & vbCrLf & _
        "c.Quantity, f.Location_Name, h.ADA_Code, " & vbCrLf & _
        "h.ADA_Definition, b.Comments "
    '--from
    sSql = sSql & vbCrLf & _
        "FROM DENTAL_ENCOUNTER a, " & vbCrLf & _
        "     LAB_CASE b, " & vbCrLf & _
        "     LAB_TX_SPECIFICATIONS c, " & vbCrLf & _
        "     DENTAL_ENCOUNTER_TX d, " & vbCrLf & _
        "     NOTE_TXS_LOCATION e, " & vbCrLf & _
        "     NOTE_LOCATION f, " & vbCrLf & _
        "     TX_ADA_CODE_DEFINITION h "
    '--where
    sSql = sSql & vbCrLf & _
        "WHERE a.Patient_Unit_ID = " & psPatientUnitNumber & vbCrLf & _
        "AND   a.Encounter_Facility_NCID = " & psFacilityNCID & vbCrLf & _
        "AND   a.Dental_Encounter_ID = b.Dental_Encounter_ID " & vbCrLf & _
        "AND   b.Lab_Case_ID = c.Lab_Case_ID " & vbCrLf & _
        "AND   c.Tx_Code_ID = d.Tx_Code_ID " & vbCrLf & _
        "AND   d.Tx_Code_ID = e.Tx_Code_ID " & vbCrLf & _
        "AND   e.Location_ID = f.Location_ID " & vbCrLf & _
        "AND   h.Tx_ADA_Code_ID = d.Tx_ADA_Code_ID " & vbCrLf & _
        "AND   h.LabFlag = -1 " & vbCrLf & _
        "AND   b.Lab_Case_Number = '" & psLabCaseNumber & "'" & vbCrLf & _
        "ORDER BY 2 "
        
    Set rsSubReport = oDAS.OpenRecordset(sSql)
    
    'psProviderName = psProviderName & "-" & objEncounter.EncProperties.SigBlockLine2

    If Not objForms.PrintDD2322(cwobjPatient, _
                              psLabCaseNumber, _
                              psClinicName, _
                              psClinicNCID, _
                              psProviderName, _
                              rsSubReport, _
                              bPreview) Then GoTo ErrHandler
    PrintDD2322 = True
    GoTo CleanUp
    Exit Function

ErrHandler:
    PrintDD2322 = False
    gobjShared.ShowVBError Err.Number, Err.Description, "Labs.PrintDD2322", "CHCSII_Dental_Labs", vbExclamation
CleanUp:
    Set objForms = Nothing
    Set oDAS = Nothing
    Exit Function

    Err.Clear
    Resume
    

End Function


'**************************************************************************************************
'  Function Name: SaveView
'
'  Description: This method is called when the CW Core attempts to save a "view" provided by this
'               OLE server.
'
'  Parameters:
'
'  Return Value:
'    Returns TRUE if the "view" was saved successfully.
'    Returns FALSE if an error occurred saving the "view".
'**************************************************************************************************
Public Function SaveView() As Boolean

On Error GoTo ErrHandler

SaveLabCaseView

'-- -set return value---
SaveView = True

Exit Function
   
ErrHandler:
   '-- -set return value---
   SaveView = False

   '-- -show the VB error---
   Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental Labs.SaveView", "Dental Labs", vbCritical)
   
End Function

