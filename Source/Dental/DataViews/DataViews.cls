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
Attribute VB_Name = "DataViews"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'************************************************************************
'Module:        DataViews/DataViews.cls
'
'Description:   Dental DataView CHCSII interface class
''
'Modules Required:
'
'Author         : Wilson Sun
'Modified       :
'               :
'************************************************************************
Option Explicit   '-- $Reformatted by FormatVB, Ver. 1.6b, on 02-23-00 at 10:50:53
'Public currentDEGCOps As GCOps

Private mbUnderHealthHistory As Boolean  'Launched under Health History folder
Private mlngProgId As Long 'scr 27573
Private mlngParenthWnd As Long 'scr 27573
Private mobjMainForm As frmDataViews
Private mobjEncParent       As ENCOUNTER_INTERFACES.IParent
Private mblnMainViewRunning As Boolean
Private mblnFirstRegularResize As Boolean
Private mobjDental      As Object
Private mobjUser        As CHCSII_ClientSecurity.ICHCSII_User
Private mobjShared As CWShared
Private mobjCWComm As Object
Private mobjLogonEnviron As Object
Private mobjConfig As Object
Private mobjPatient As Object

''''''''''''''' Constanst Section '''''''''''''''''''''''''''''''''''''''''

Const msMODULE_NAME As String = "DataView."
#If DEBUGON Then
Private DR As DebugRec
Public DL As cDebugLog
Private Const Exe As String = "OrderSets"
Private Const Module = "DataViews.cls"

#End If

Private mblnUseStoredProc   As Boolean

Dim oSF                     As New SVGTeeth.SF
Public WithEvents oTeethSVG As SVGTeeth.clsTeeth
Attribute oTeethSVG.VB_VarHelpID = -1

Private mobjCHCSIIConn As CHCSII_CONN.Conn
Private mobjDas As ICHCSII_DataAccess.ICHCSII_DAS

Implements ICHCSII_CoreAPI2
Implements ICHCSII_CoreAPI3
Implements ICHCSII_User
Implements ICHCSII_Dental
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
    Dim oEncDoc As Encounter
    'SCR 25225

    Select Case viViewType
        Case cwiPROPERTY_VIEW
            mbUnderHealthHistory = True
            QueryView = False
        Case cwiREGULAR_VIEW
            QueryView = True              '--this view is NOT optional--
        Case cwiPATIENT_SUMMARY_VIEW
            QueryView = False
        Case cwiCLINICIAN_SUMMARY_VIEW
            QueryView = False
        Case cwiENCOUNTER_SUMMARY_VIEW
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
    Dim objCoreAPI3 As ICHCSII_CoreAPI3
    
    Const sPROC_NAME As String = "OpenView"
    On Error GoTo ErrHandler
    
    If mblnMainViewRunning Then
        OpenView = True
        Exit Function
    End If
    
    'KDunne 10/23/2002
    'SCR 27573
    Call CheckMyProgID(mlngProgId)
    mlngParenthWnd = vlParenthWnd
    
    '-- -perform actions based on the "view type"---
   If viViewType = cwiREGULAR_VIEW Then
      '-- -set global flags---
      mblnFirstRegularResize = True
      mblnMainViewRunning = True
      
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
      SetParent mobjMainForm.hwnd, vlParenthWnd
      '-- -set parameter---
      rhMainForm = mobjMainForm.hwnd

   Else                                '--unsupported view "type"--
      '-- -set return value---
      OpenView = False
      
      Exit Function

   End If

   '-- -set return value---
    '-------------------------------------------------------------------
    ' Create Action Menu Item and Enable/Disable Tabs
    '-------------------------------------------------------------------
    mobjCWComm.DeleteMenuItem (glMenuPrintID)
    mobjCWComm.Message cwiGENERIC, "DELETE_SAVE", NCID_DDSA_APPLICATION, NCID_DENTAL_DATAVIEWS_APPLICATION
    mobjMainForm.vsIndexTab1.TabEnabled(TabEnum.eWorkload) = True   '-- Disable Workload Tab.
    mobjMainForm.vsIndexTab1.TabEnabled(TabEnum.ePerioStatistics) = True
    mobjMainForm.vsIndexTab1.TabEnabled(TabEnum.ePerioComparison) = True

    mobjCWComm.Message cwiGENERIC, "Refresh", NCID_CW_APPLICATION, NCID_DENTAL_LABS_APPLICATION

   'KDunne 8/14/2002
   'SCR 26609
   mobjMainForm.UnderHealthHistory = mbUnderHealthHistory
    
   OpenView = True
   
   Exit Function
   
ErrHandler:
   '-- -show the VB error---
    Call mobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME & sPROC_NAME, App.Title, vbCritical)
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

      If mblnMainViewRunning Then
         '-- -set global flags---
         mblnMainViewRunning = False
         
         '-- -this indicates that the CW Core is closing the "view", NOT this OLE Server---
         cwbCloseViewMyself = False
         
         'Unload Teeth SVG image
         If oTeethSVG Is Nothing Then
            Set oTeethSVG = oSF.Singleton
         End If
         
         'Call SetParent(oTeethSVG.getChildHwnd, oTeethSVG.getParentHwnd)
         oTeethSVG.ReleaseTeethControl
         
         mobjCWComm.DeleteMenuItem (glMenuPrintID)
         '-- -close ALL forms that are associated with all "views"---
           Unload mobjMainForm

      End If

   ElseIf viViewType = cwiREGULAR_VIEW Then
      '-- -close the "main view"---

      If mblnMainViewRunning Then
         '-- -set global flags---
         mblnMainViewRunning = False
                  
         'Unload Teeth SVG image
         'Call SetParent(oTeethSVG.getChildHwnd, oTeethSVG.getParentHwnd)
         oTeethSVG.ReleaseTeethControl
         
         If Not mobjCWComm Is Nothing Then
            mobjCWComm.DeleteMenuItem (glMenuPrintID)
         End If

         '-- -this indicates that the CW Core is closing the "view", NOT this OLE Server---
         cwbCloseViewMyself = False
         
         '-- -close ALL forms that are associated with all "views"---
         If mobjMainForm Is Nothing = False Then
           Unload mobjMainForm
         End If

         
      End If

   End If
   
   If mobjMainForm Is Nothing = False Then
        If mobjMainForm.IsLoaded Then
            Unload mobjMainForm
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
    Const sPROC_NAME As String = "MenuItemSelected"
        
    On Error GoTo ErrHandler
   
    Select Case vlMenuId
        Case glMenuPrintID
            '--------------------------------------------------
            ' Build RecordSet for SF603A Report
            ' Notes: grsNotes value get from LoadEncounterTypeSql()
            '--------------------------------------------------
'            If Not grsNotes Is Nothing Then
            
            PrintSF603A False
            GoTo CleanUp
        Case Else   '--do nothing--
    End Select
   
   Exit Sub
   
ErrHandler:
   '-- -show the VB error---
    Call mobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME & sPROC_NAME, App.Title, vbCritical)
CleanUp:
    mobjCWComm.setStatusText ""
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

   If mblnMainViewRunning Then
        '-- ****************************************************************************
        '--  TO DO:
        '--    Add code here to update the data in all "running" views with this new
        '--    patient.
        '-- ****************************************************************************
        '-- Reset global variabes for refresh (reload) data
        Call mobjMainForm.ResetObjectForRefresh
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
    Debug.Print "ResizeView " & Now
    Const sPROC_NAME As String = "ResizeView"
    On Error GoTo ErrHandler
   
   '-- -reset global "client area" parameters---
   cwlLeft = vlLeft
   cwlTop = vlTop
   cwlWidth = vlWidth
   cwlHeight = vlHeight
   
    If CheckMyProgID(mlngProgId) Then
      SetParent mobjMainForm.hwnd, mlngParenthWnd
    End If
   
   '-- -perform different actions based upon the "view type"---

   If viViewType = cwiREGULAR_VIEW Then

      If Not mblnFirstRegularResize Then
         '-- -------------------------------------------------------------------------------
         '--  Reset the "WindowState" to "Normal", so the form can be resized.  This is
         '--    needed because a window CANNOT be moved or resized if it is "minimized" or
         '--    "maximized".
         '--
         '--  NOTE: This generates a form "resize" event.
         '-- -------------------------------------------------------------------------------
         mobjMainForm.WindowState = vbNormal
         
         '-- -hide the form to so that annyoing "screen blinking" will not be seen while controls are moved and resized---
         mobjMainForm.Visible = False
         
         '-- -resize form to fill entire "client area", NOTE: This generates a "resize" event---
         mobjMainForm.Move cwlLeft, cwlTop, cwlWidth, cwlHeight
         
         mobjMainForm.Visible = True
      Else
         '--first "regular" resize, need to perform "initailization" OrderSetss--
         '-- -reset flag so this code will not be executed again---
         mblnFirstRegularResize = False
         
            
        Set oTeethSVG = oSF.Singleton
        
         '-- --------------------------------------------------------------------------
         '--  NOTE: Here is where you "Show" the "Regular View" form for the FIRST time.
         '-- --------------------------------------------------------------------------

      End If

   End If
   
   
   mobjMainForm.Initialize cwlMyProgId, mobjCWComm, mobjLogonEnviron, mobjConfig, _
                            mobjPatient
                            
   Set mobjMainForm.TeethSVG = oTeethSVG
   Set mobjMainForm.EncounterParent = mobjEncParent
   mobjMainForm.NoMoreTimers
   mobjMainForm.Display
   Exit Sub
   
ErrHandler:
   '-- -show the VB error---
    Call mobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME & sPROC_NAME, App.Title, vbCritical)
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
    Set mobjShared = New CWShared

    '-- -set the help file for this application---
    App.HelpFile = mobjShared.CWHelpFile
    
    Set mobjMainForm = New frmDataViews
    
    '--- SCR 28292 RWalker 10/8/02Instantiate ARM / Init transactions
    Call ARM_InitAllTransactionsForDental_DataViews
   
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
#If DEBUGON Then
   Dim DR As DebugRec
   
   DL.DebugLog Exe, Module, "Class_Terminate", "Begin", DR
#End If
   '---release reference to objects---
   '-- Set gobjAnPCollection = Nothing
   Set mobjShared = Nothing
   Set mobjCWComm = Nothing
   Set mobjLogonEnviron = Nothing
   Set mobjConfig = Nothing
   Set mobjPatient = Nothing
   Set mobjUser = Nothing
   
   '-- ********************************************************************
   '--  TO DO:
   '--    Add code here to release reference to all other global objects.
   '-- ********************************************************************
   
    CleanObjects
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
   Const sPROC_NAME As String = "Message"
   On Error GoTo ErrHandler
   
   '-- -----------------------------------------------------------------------------
   '--  Take corresponding actions based upon the message "type".
   '--
   '--  NOTE: This method is required, but currently there are NO REQUIRED messages
   '--        to process.
   '-- -----------------------------------------------------------------------------
   Select Case viMessageType
        Case cwiOK_TO_CLOSE_VIEW
            mobjCWComm.Message cwiOK_TO_CLOSE_VIEW, "Y", vsSourceServerNCID, NCID_DENTAL_DATAVIEWS_APPLICATION
        
        Case cwiFOREGROUND_APPLICATION
            'KDunne 1/13/2002
            'SCR 31404
            'Refresh the tab
            'frmGraphicalCharting.tabGraphicCharting_Switch 0, frmGraphicalCharting.tabGraphicCharting.CurrTab, 0
                
            'Get the teeth graphic.
            If Not oTeethSVG Is Nothing Then
                oTeethSVG.SetTeethControl mobjMainForm.elTeeth.hwnd
            End If
   End Select
   Exit Sub
   
ErrHandler:
   '-- -show the VB error---
    Call mobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME & sPROC_NAME, App.Title, vbCritical)
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
    Debug.Print "Initialize " & Now
    Const sPROC_NAME As String = "Initialize"
    On Error GoTo ErrHandler
    
    mbUnderHealthHistory = False
    
    '---Instantiate Help-file object
    App.HelpFile = mobjShared.CWHelpFile
   
#If STOREDPROC Then
    If Not mobjShared.UseLocalDB And mobjShared.CmdLineSwitches("NOSTOREDPROC") <> "NOSTOREDPROC" Then
        mblnUseStoredProc = True
    End If
#End If

    '-- -create references to passed in objects---
    Set mobjCWComm = robjComm
    Set mobjLogonEnviron = robjLogonEnvironment
    Set mobjConfig = robjConfig
    Set mobjPatient = robjPatient
    
    'KDunne 10/23/2002
    'SCR 27573
    Call CheckMyProgID(vlMyProgId)
          
    '---save the "ProgId" assigned to me by the CW Core---
    cwlMyProgId = vlMyProgId
    mlngProgId = vlMyProgId 'SCR 27573
   
    '---set global flags---
    mblnMainViewRunning = False
   
    '-- ****************************************************************************
    '--  TO DO:
    '--    Add code here to create objects, initialize global variables, etc.
    '-- ****************************************************************************
    
    '**********************
    '--Start Process
    '**********************
    Set mobjCHCSIIConn = New CHCSII_CONN.Conn
    Set mobjDas = mobjCHCSIIConn.CHCSII_DAS
'    If Not mobjShared.UseLocalDB Then
'       Set goDas = New CHCSII_DataAccess.CHCSII_DAS
'       Set goSqlOps = New CHCSII_SQL.SqlOps
'    Else
'       Set goDas = New GEMS_DataAccess.GEMS_DAS
'       Set goSqlOps = New GEMS_SQL.SqlOps
'    End If

'    Set gobjDentalNotes = New CHCSII_Dental_Notes.clsDTNotesOps
    Exit Sub
   
ErrHandler:
   '-- -show the VB error---
    Call mobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME & sPROC_NAME, App.Title, vbCritical)
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
    Const sPROC_NAME As String = "Initialize"
   On Error GoTo ErrHandler
   
   Set mobjUser = robjUser
   
   Dim objRoles      As ICHCSII_Roles
   Dim objResources  As ICHCSII_Resources
   
   Call Initialize(ByVal vlMyProgId, robjComm, robjLogonEnvironment, robjConfig, robjPatient)
   
   Exit Sub
   
ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME & sPROC_NAME, App.Title, vbCritical)
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
    Const sPROC_NAME As String = "Properties"
   On Error GoTo ErrHandler
   
   '-- *****************************************************************
   '--  TO DO:
   '--    Add code here to invoke this OLE Server's properties page.
   '-- *****************************************************************
 
   Exit Sub
   
ErrHandler:
   '-- -show the VB error---
    Call mobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME & sPROC_NAME, App.Title, vbCritical)
   Exit Sub

End Sub
Private Function ICHCSII_CoreAPI3_CloseView2(ByVal viViewType As Integer) As Boolean
    ICHCSII_CoreAPI3_CloseView2 = True
    '--------------------------------------------------------------------------------
    'Wilson 12/27/01
    'If application was lunched under Health History folder. CloseView method was not called.
    'So copy CloseView code here. Or just directory call CloseView
    CloseView (viViewType)
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
    Set mobjDental = objDentalParent
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
    'If Not oDentalNotes.Mtd_bGetNotes(pnPatientUnitId, grsNotes, False, pnEncounterFacilityNCID) Then GoTo ErrorHandler

    ICHCSII_Dental_SetEncounterDentalNotes = True
End Function
Private Function ICHCSII_Dental_SetEncounterOps(objEncounterOps As Object) As Boolean
    Dim objEncounter As CHCSIIEncounterOps.Encounter
    
    Set mobjEncParent = objEncounterOps
    
    'KDunne 7/10/2002
    'SCR 25496 - Encounter Interface has changed.
    If Not mobjEncParent Is Nothing Then
        Set objEncounter = mobjEncParent.CurrentEncounter
        guCurEncInfo.ProviderFullName = objEncounter.PrimaryProvider.FullName
        guCurEncInfo.EncounterID = objEncounter.EncounterID
        Set mobjMainForm.EncounterParent = mobjEncParent
    End If

    Set objEncounter = Nothing
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

Private Function IChild_GetAutoCiteNarrative(TestType As String, NumPeriods As Integer, Period As AC_Period) As String

End Function

Private Function IChild_SetParent(Parent As IParent) As Boolean
   Set mobjEncParent = Parent
   IChild_SetParent = True
End Function
'********************************************************************************
' Name:         CleanObjects
' Description:  Initialize all objects
' Parameter:    N/A
' Return:       N/A
' Called By:    Class_Terminate
' Histary:
'   Developer:      Date:       Description
'   ----------      ---------   -----------------------------------------------
'   Wilson Sun      11/24/01    Set more objects to nothing
'   Wilson sun      3/23/2001   Created
'********************************************************************************
Private Sub CleanObjects()
    Set mobjDas = Nothing
    'Set = Nothing
    
    Set mobjDental = Nothing
    Set mobjEncParent = Nothing

End Sub

Public Function Mtd_bDentalTabChange(ByVal lFrom As Long, ByVal lTo As Long) As Boolean
'************************************************************************
'Called By:     Dental Parent
'
'Arguments:
'lFrom        Identifier of the dental the currently active dental module
'lTo            Identifier of the dental the about to became active
'
'
'Description:
'               if lFrom equal to your Tab
'                  Focus is changing from your tab to lTo Module
'                  if you ready to release focus: save current data...
'                     Return true
'                  else
'                     return false
'               Elseif lTo equal to your Tab
'                  Focus is changing from lTo  tab to your Module
'                  if you ready to recive focus: Objects are initialized...
'                     Return true
'                  else
'                     return false
'               Else
'                  Note: May not be called
'                  Return true
'
'Note:
'
'Warning:
'
'Updates Globals:
'
'Returns:       If successful  - true
'               Else           - false
'
'Programmer:
'''************************************************************************
    Const sPROC_NAME As String = "Mtd_bDentalTabChange"
    Dim bResult As Boolean
    Dim sErrorMsg  As String
    Dim sMsg As String
    On Error GoTo ErrHandler
    
   
    bResult = True
       
    
    If lTo = childTab.DataViews Then
       'Regardless whomever coming from into my DataViews, do it
        Screen.MousePointer = vbHourglass
        
        bResult = True
        mobjCWComm.Message cwiGENERIC, "DELETE_SAVE", NCID_DDSA_APPLICATION, NCID_DENTAL_DATAVIEWS_APPLICATION

    Else
        'Clean Object for next refresh
        mobjMainForm.ResetObjectForRefresh
        mobjCWComm.DeleteMenuItem (glMenuPrintID)
        
        mobjCWComm.Message cwiGENERIC, "ADD_SAVE", NCID_DDSA_APPLICATION, NCID_DENTAL_DATAVIEWS_APPLICATION
        'Call SetParent(oTeethSVG.getChildHwnd, oTeethSVG.getParentHwnd)
        oTeethSVG.ReleaseTeethControl
        
    End If
    
    'Make sure core updates menu items
    mobjCWComm.Message cwiGENERIC, "Refresh", NCID_CW_APPLICATION, NCID_DENTAL_DATAVIEWS_APPLICATION
    
    
    Mtd_bDentalTabChange = bResult

    GoTo CleanUp

ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME & sPROC_NAME, App.Title, vbCritical)
    Mtd_bDentalTabChange = False
CleanUp:
    Screen.MousePointer = vbDefault
End Function
Public Function SaveView() As Boolean
    SaveView = True
End Function
'********************************************************************************
' Name:         SaveEncounterType
' Description:  The interface for other module to call this method
' Parameter:    lpDentalEncounterID
' Return:       Ture or False
' Note:         Return Ture for now.
' Called By:    Other module under SaveView() function
' Histary:
'   Developer:  Date:       Description
'   ----------  ----------  -----------------------------------------------
'   Wilson sun  01/07/02    Created
'********************************************************************************
Public Function SaveEncounterType(ByVal lpDentalEncounterID As Long) As Boolean
    SaveEncounterType = True
End Function
'********************************************************************************
' Name:         DisplayUnsignedWorkload
' Description:  Retrieve Unsigned workload data
' Parameter:
' Return:       N/A
' Called By:    Dental Module
' Histary:
'   Developer:  Date:       Description
'   ----------  ----------  -----------------------------------------------
'   Wilson Sun  02/05/02    Created for SCR: 17424, 18721, 21195
'********************************************************************************
Public Sub DisplayUnsignedWorkload(ByVal psEncounterID As String, ByVal psProviderName As String)
    Dim objDataOps As clsDVDataOps
    Dim objRSWorkLoad As ADODB.Recordset
    
    On Error GoTo ErrHandler
    
    Set objDataOps = New clsDVDataOps
    objDataOps.Initialize guCurEncInfo.ProviderFullName, mobjShared, mobjCWComm, _
            mobjDas, mobjCHCSIIConn.CHCSII_SQLOPS, False, mobjPatient, False
    
    guCurEncInfo.ProviderFullName = psProviderName
    guCurEncInfo.EncounterID = psEncounterID
    
    If mblnUseStoredProc Then
        Set objRSWorkLoad = objDataOps.LoadWorkloadSqlSP(mobjPatient.Name, mobjPatient.SSN, _
            guCurEncInfo.ProviderFullName, guCurEncInfo.EncounterID)
    Else
        Set objRSWorkLoad = objDataOps.LoadWorkloadSql(mobjPatient.UnitNumber)
    End If
    If objRSWorkLoad Is Nothing = False Then
        Set frmUnSignedWorkload.WorkLoadRS = objRSWorkLoad
        '--- SCR-ARM-25468/27364 RWalker 4/10/03
        goARM.StartCmplxTx gsARM_Tx_CorrUT_UnsignedWorkload, , , , , , , , eAnyFamily
        frmUnSignedWorkload.Show vbModal
        '--- SCR-ARM-25468/27364 RWalker 4/10/03
        goARM.StopCmplxTx gsARM_Tx_CorrUT_UnsignedWorkload, enmARMTxStatusOk, enmARMStop_MeOnly
    End If
    
FuncExit:
    Set objDataOps = Nothing
    Set objRSWorkLoad = Nothing
    
    Exit Sub
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, msMODULE_NAME & ".DisplayUnsignedWorkload", App.Title, vbCritical
    GoTo FuncExit
    
End Sub

'********************************************************************************
'Sub:PrintSF603A
'Author:KDunne
'Date:7/15/2002
'Purpose: Sets SVG Imgage back to parent, Prints SF603A an resets SVG image back
'         to server.
'Notes: SCR #25095
'********************************************************************************
Public Sub PrintSF603A(Optional ByVal blnPreview As Boolean = False)
    Const sPROC_NAME As String = "MenuItemSelected"
    Dim rsReport As ADODB.Recordset
    Dim objRSNotes As ADODB.Recordset
    Set objRSNotes = mobjMainForm.Notes_RS
    
    'KDunne 6/27/2002
    'SCR 25234 - Make sure our object isn't nothing
    If objRSNotes Is Nothing Then
        MsgBxARMd "There is no encounter note to be printed", , "Print SF603A"
        GoTo FuncExit
    End If
    'End scr 25234
    
    If objRSNotes.RecordCount > 0 Then
        'Release Image
        'Call SetParent(oTeethSVG.getChildHwnd, oTeethSVG.getParentHwnd)
        'This was causing a problem when printing the sf603a SCR 31149
        'oTeethSVG.ReleaseTeethControl
        
        Set rsReport = LoadReportRTF(objRSNotes, mobjPatient.UnitNumber, mobjShared, mobjCWComm)
    Else
        MsgBxARMd "There is no encounter note to be printed", , "Print SF603A"
        GoTo FuncExit
    End If
    '--------------------------------------------------
    ' Calling PrintSF603A2 method for report
    '--------------------------------------------------
    Dim objForms As New CHCSII_Forms.DentalForms
    Set objForms.Comm = mobjCWComm 'If need to show Status Text
    
    Screen.MousePointer = vbHourglass
    mobjCWComm.setStatusText "Loading SF603A ..."
    
    
    DoEvents
    Call objForms.PrintSF603A2(mobjPatient, Nothing, rsReport, Nothing, Nothing)
    
    'frmDataViews.tmrTeeth.Enabled = True
    mobjMainForm.NoMoreTimers

FuncExit:
    mobjCWComm.setStatusText ""
    Screen.MousePointer = vbDefault

    Set objRSNotes = Nothing
    Set objForms = Nothing
    Exit Sub
    
ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME & sPROC_NAME, App.Title, vbCritical)
    GoTo FuncExit
End Sub

'**********************************************************************************
'Function:CheckMyProgID
'Author:KDunne
'Date: 10/23/2002
'Purpose:  Since Dataviews has two core interfaces, Dataviews.cls and DataviewsHH.cls,
'          when instantiating one interface, we must make sure that the other one
'          was properly closed.
'***********************************************************************************
Private Function CheckMyProgID(ByVal lngProgId As Long) As Boolean
    Dim objCoreAPI3 As ICHCSII_CoreAPI3
    
    'If 0, then Server is initializing for the first time. Don't do anything,
    'let the server initialize on its own.
    If cwlMyProgId = 0 Then Exit Function
    If cwlMyProgId <> lngProgId Then
         'DataviewsHH was previously opened. Make sure that server was shut down,
         'and the frmDataviews and oTeethSVG parents were reset.
         Set objCoreAPI3 = mobjCWComm.InitializeOLEServer(NCID_DENTAL_DATAVIEWSHH_APPLICATION)
         objCoreAPI3.CloseView2 (cwiALL_VIEWS)
         Set objCoreAPI3 = Nothing
         
         'Set the global cwlMyProgId equal to the progid of this server.
         cwlMyProgId = lngProgId
         
         CheckMyProgID = True
    End If
End Function

