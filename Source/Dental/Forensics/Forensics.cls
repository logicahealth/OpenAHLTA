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
Attribute VB_Name = "Forensics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit   '-- $Reformatted by FormatVB, Ver. 1.6b, on 02-23-00 at 10:50:53
Public patientAge As Integer
Public patientUnitID As Long
Public dentalEncounterID As String
Public encounterID As Long
Public dentalProviderID As Long
Public anteMortemID As Long
Public MatchType As String
#If DEBUGON Then
'-- Private DR As DebugRec
'-- Public DL As cDebugLog
'-- Public Const Exe = "OrderSets"
Private Const Module = "Dental.cls"
#End If
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

   Select Case viViewType
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

   On Error GoTo errhandler
      
   
   
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
      SetParent frmForensics.hWnd, vlParenthWnd
      Set frmForensics.ParentObject = Me
      '-- -set parameter---
      rhMainForm = frmForensics.hWnd
   Else                                '--unsupported view "type"--
      '-- -set return value---
      OpenView = False
      
      Exit Function

   End If
   
   '-- -set return value---
   OpenView = True
   
   Exit Function
   
errhandler:
   '-- -show the VB error---
   Call gobjShared.ShowVBError(Err.Number, Err.Description, "Forensics.OpenView", "Forensics", vbCritical)
   
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
    
   'Screen.MousePointer = vbHourglass
    
   If viViewType = cwiALL_VIEWS Then
      '-- -close the "main view"---

      If cwbMainViewRunning Then
         '-- -set global flags---
         cwbMainViewRunning = False
         
         '-- -this indicates that the CW Core is closing the "view", NOT this OLE Server---
         cwbCloseViewMyself = False
         
         '-- -close ALL forms that are associated with all "views"---
'        If gobjEnc Is Nothing Then
           Unload frmForensics
'        Else
'        If Not gobjEnc.StartSave() Then
'          MsgBox "Your lock on the Dental section of the current encounter has been taken away by another user."
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
         
         '-- -close ALL forms that are associated with all "views"---
'        If gobjEnc Is Nothing Then
           Unload frmForensics
'        Else
'        If Not gobjEnc.StartSave() Then
'          MsgBox "Your lock on the Dental section of the current encounter has been taken away by another user."
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

   On Error GoTo errhandler
   
   '-- --------------------------------------------------------------------------------
   '--  NOTE: You don't have to create menu items, these are optional, that is why
   '--        the code is commented out.
   '-- --------------------------------------------------------------------------------
'    '---perform the actions associated with the menu item that was selected---
    Select Case vlMenuId
   
'            '******************************************************************************
'            '  TO DO:
'            '    Take actions needed to process this menu item.
'            '******************************************************************************
'
'
'
'        Case Else   '--do nothing--
    End Select
   
   Exit Sub
   
errhandler:
   '-- -show the VB error---
   Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental_Forensics.MenuItemSelected", "Forensics", vbExclamation)
   
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
      
      patientUnitID = cwobjPatient.UnitNumber
      'frmForensics.doInitialize
      
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

   On Error GoTo errhandler
   
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
         frmForensics.WindowState = vbNormal
         
         '-- -hide the form to so that annyoing "screen blinking" will not be seen while controls are moved and resized---
         frmForensics.Visible = False
         
         '-- -resize form to fill entire "client area", NOTE: This generates a "resize" event---
         frmForensics.Move cwlLeft, cwlTop, cwlWidth, cwlHeight
         
         '<bjr090401 - moved to after updatePSR, etc.
         'frmForensics.Visible = True
         
         'frmForensics.UpdatePSR
         
         ' update DxTxGrid
         'Call frmForensics.fgDxTx.Charts.LoadDxs(patientUnitID, CInt(DentalEncounterID))
         'frmForensics.fgDxTx.RefreshGrid True
         'frmForensics.fgDxTx.HideColumn frmForensics.fgDxTx.FlexGrid.cols - 1, True
         'frmForensics.fgDxTx.SetYellowCell = True
         
         ' refresh
         '<bjr090401 [SCR15676] - removed - we don't need this during a resize
         'frmForensics.cmdView(0).Value = True
         
         '<bjr090401 - this should be last - moved from above
         frmForensics.Visible = True
         
      Else       '--first "regular" resize, need to perform "initialization" OrderSetss--
         '-- -reset flag so this code will not be executed again---
         cwbFirstRegularResize = False
         
         '-- --------------------------------------------------------------------------
         '--  NOTE: Here is where you "Show" the "Regular View" form for the FIRST time.
         '-- --------------------------------------------------------------------------
          'If Not gobjEnc Is Nothing Then
          '  glMenuSave = gobjEnc.AddMenuItem(0, "&Save", "save.bmp", "Save Dental")
          '  glMenuNext = gobjEnc.AddMenuItem(0, "D&isposition", "Next.bmp", "Go to the Disposition screen")
            frmForensics.Display
         
          'Else
          '  CloseView cwiREGULAR_VIEW
          '  Exit Sub
          'End If
      End If

   End If
   
   Exit Sub
   
errhandler:
   '-- -show the VB error---
   Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental_Forensics.ResizeView", "Forensics", vbExclamation)
   
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
   Dim DR  As DebugRec

   DL.DebugLog Exe, Module, "Class_Terminate", "Begin", DR
#End If
   '-- -release reference to objects---
   
'        Set gobjShared = Nothing
'        Set cwobjCWComm = Nothing
'        Set cwobjLogonEnviron = Nothing
'        Set cwobjConfig = Nothing
'        Set cwobjPatient = Nothing
'        Set gobjUser = Nothing
 
   
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

   On Error GoTo errhandler
   
   '-- -----------------------------------------------------------------------------
   '--  Take corresponding actions based upon the message "type".
   '--
   '--  NOTE: This method is required, but currently there are NO REQUIRED messages
   '--        to process.
   '-- -----------------------------------------------------------------------------
   Select Case viMessageType
        
   End Select
   
   Exit Sub
   
errhandler:
   '-- -show the VB error---
   Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental_Forensics.Message", "Forensics", vbExclamation)
   
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

   On Error GoTo errhandler
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
   
   
   Exit Sub
   
errhandler:
   '-- -show the VB error---
   Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental_Forensics.Initialize", "Forensics", vbExclamation)
   
   Exit Sub
End Sub



Private Function GetDentalEncounterID(objPatient As Object, objEncounter As Object) As Long
    ' objEncounter - CHCSIIEncounterOps.IParent
    ' objPatient - patient object
    
    '<bjr082001 theater
    'Dim objDAS          As CHCSII_DAS
    Dim objDAS          As ICHCSII_DAS
    
    Dim objRS           As Recordset
    Dim sqlString       As String
    
    On Error GoTo errhandler
    
    GetDentalEncounterID = 0
    sqlString = "SELECT DENTAL_ENCOUNTER_ID FROM DENTAL_ENCOUNTER WHERE " _
               & "Encounter_Number=" & objEncounter.encounterID & " AND " _
               & "Encounter_Facility_NCID=" & objEncounter.FacilityNCID & " AND " _
               & "Patient_Unit_ID=" & objPatient.UnitNumber
                  
    'sqlString = "SELECT DENTAL_ENCOUNTER_ID FROM DENTAL_ENCOUNTER WHERE " _
    '           & "Encounter_Number=21033 AND " _
    '           & "Encounter_Facility_NCID=1049621 AND " _
    '           & "Patient_Unit_ID=39119"
               
    '<bjr082001 theater
    'Set objDAS = New CHCSII_DAS
    Set objDAS = GetDasObj()
    
    Set objRS = objDAS.OpenRecordset(sqlString)
    '''If found
    If Not objRS.RecordCount = 0 Then
       ' found an existing encounter of patient
       GetDentalEncounterID = objRS.Fields("DENTAL_ENCOUNTER_ID")
    End If
    Exit Function
    
errhandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "in Forensics.GetDentalEncounterID()", "Forensics", vbCritical)
    
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

   On Error GoTo errhandler
   
   Set gobjUser = robjUser
   
   Dim objRoles      As ICHCSII_Roles
   Dim objResources  As ICHCSII_Resources
   
   Call Initialize(ByVal vlMyProgId, robjComm, robjLogonEnvironment, robjConfig, robjPatient)
   
   Exit Sub
   
errhandler:
   Call gobjShared.ShowVBError(Err.Number, Err.Description, "in frmForensics.Initialize", "Forensics !", vbExclamation)
   
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

   On Error GoTo errhandler
   
   '-- *****************************************************************
   '--  TO DO:
   '--    Add code here to invoke this OLE Server's properties page.
   '-- *****************************************************************
   
   
   
   Exit Sub
   
errhandler:
   '-- -show the VB error---
   Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental_Forensics.Properties", "Forensics", vbExclamation)
   
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
  If Not objDentalParent Is Nothing Then
    Set gobjDental = objDentalParent
    dentalEncounterID = gobjDental.psDentalEncounterID
  End If
End Function

Private Function ICHCSII_Dental_SetEncounterDentalBT(objForensicsOps As Object) As Boolean

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
 If Not objEncounterOps Is Nothing Then
     dentalProviderID = objEncounterOps.CurrentEncounter.Providers(1).NCID
     patientUnitID = cwobjPatient.UnitNumber
     patientAge = 21
     encounterID = objEncounterOps.CurrentEncounter.encounterID
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

Private Function IChild_SetParent(Parent As IParent) As Boolean
   
End Function

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

On Error GoTo ErrHnd
   
   If lFrom = 4 Then
      Call SaveView
   End If
   'If frmForensics.Visible Then frmForensics.doInitialize
   Mtd_bDentalTabChange = True
   
Exit Function
ErrHnd:
   Mtd_bDentalTabChange = False

Exit Function
End Function

Public Function SaveView() As Boolean
Dim sSQL As String
Dim objMatch As clsMatchSet
Dim bFound As Boolean

'If Not frmForensics.mcolMatches Is Nothing Then
'  For Each objMatch In frmForensics.mcolMatches
'     If patientUnitID = objMatch.PMP_ID And anteMortemID = objMatch.AMP_ID Then
'        bFound = True
'        Exit For
'     End If
'  Next
'End If
'If Not bFound Then
'    If Not frmForensics.mcolNotMatches Is Nothing Then
'      For Each objMatch In frmForensics.mcolNotMatches
'         If patientUnitID = objMatch.PMP_ID And anteMortemID = objMatch.AMP_ID Then
'            bFound = True
'            Exit For
'         End If
'      Next
'    End If
'End If
'If Not bFound Then
'    If Not frmForensics.mcolNotKnown Is Nothing Then
'      For Each objMatch In frmForensics.mcolNotKnown
'         If patientUnitID = objMatch.PMP_ID And anteMortemID = objMatch.AMP_ID Then
'            bFound = True
'            Exit For
'         End If
'      Next
'    End If
'End If
'  If bFound Then
'    sSQL = "Update PATIENT_MATCH_STATUS " & _
'           "SET MATCH_TYPE = '" & MatchType & "' Where PMP_ID = " & patientUnitID & "and AMP_ID = " & anteMortemID
'  Else
'    sSQL = "INSERT INTO PATIENT_MATCH_STATUS " & _
'           "(AMP_ID, PMP_ID, MATCH_TYPE) " & _
'           " VALUES (" & anteMortemID & "," & patientUnitID & ",'" & MatchType & "')"
'  End If
'  If Not anteMortemID = 0 And Not MatchType = "" And Not anteMortemID = 999 Then
'    mobjDAS.ExecuteSQL sSQL
'  End If
'
'  SaveView = True
End Function
Public Function retrieveData(RHV As Long) As Recordset
Dim sSQL As String

  sSQL = "SELECT AMP_ID, PMP_ID, MATCH_TYPE " & _
         "From PATIENT_MATCH_STATUS where AMP_ID = " & RHV
  Set retrieveData = mobjDAS.OpenRecordset(sSQL)
  
End Function
Public Sub terminate_objects()
 
End Sub

