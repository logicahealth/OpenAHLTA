VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SVG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public patientUnitID As Long
Public dentalProviderID As Long

Public oSingleton As New srvSVGComm.SingletonFactory
Public oSVGComm As srvSVGComm.clsSVGComm
Private mImgFilterForm As clsImgFilterForm

Public WithEvents oTeeth As SVGTeeth.clsTeeth
Attribute oTeeth.VB_VarHelpID = -1
Private isDrawingFinished As Boolean

#If DEBUGON Then
'-- Private DR As DebugRec
'-- Public DL As cDebugLog
'-- Public Const Exe = "OrderSets"
Private Const Module = "Dental.cls"
#End If

Implements ICHCSII_CoreAPI2
Implements ICHCSII_CoreAPI3
Implements ICHCSII_User

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
    
    Set mImgFilterForm = Nothing
    
   'Screen.MousePointer = vbHourglass
'
'   If viViewType = cwiALL_VIEWS Then
'      '-- -close the "main view"---
'
'      If cwbMainViewRunning Then
'         '-- -set global flags---
'         cwbMainViewRunning = False
'
'         '-- -this indicates that the CW Core is closing the "view", NOT this OLE Server---
'         cwbCloseViewMyself = False
'
'         '-- -close ALL forms that are associated with all "views"---
'
'           'Unload frmGraphicalCharting
'
'
'   ElseIf viViewType = cwiREGULAR_VIEW Then
'      '-- -close the "main view"---
'
'      If cwbMainViewRunning Then
'         '-- -set global flags---
'         cwbMainViewRunning = False
'
'         '-- -this indicates that the CW Core is closing the "view", NOT this OLE Server---
'         cwbCloseViewMyself = False
'
'         '-- -close ALL forms that are associated with all "views"---
'           ''Unload frmGraphicalCharting
'
'      End If
'
'   End If

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
   
ErrHandler:
   '-- -show the VB error---
   Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental_BigTooth.MenuItemSelected", "Template", vbExclamation)
   
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
Dim patientAge As Integer
Dim sCalcAge As String
Dim Age As Integer
Dim lStart As Long

   '-- -update the all running "views" with the new patient---

      '-- ****************************************************************************
      '--  TO DO:
      '--    Add code here to update the data in all "running" views with this new
      '--    patient.
      '-- ****************************************************************************
'If Not patientUnitID = cwobjPatient.UnitNumber Then

lStart = Timer
    
    Screen.MousePointer = vbDefault
    'Screen.MousePointer = vbHourglass

    If Not oSVGComm.mObjSVGMngr Is Nothing Then
        'Stop all the timers
        oSVGComm.mObjSVGMngr.ofrmMenu.tmrMultithread.Enabled = False
        oSVGComm.mObjSVGMngr.mObjSVGQueue.Cancel = True
    End If
    
    If Not oTeeth Is Nothing Then
        If Not oTeeth.frmTeethX Is Nothing Then
            oTeeth.frmTeethX.TxTimer.Enabled = False
            oTeeth.frmTeethX.tmrFinished.Enabled = False
        End If
    End If
        
    patientUnitID = cwobjPatient.UnitNumber

'>> SCR #8908; 8909;   Developer: Jon Mack 04/16/2002 10:10 AM
'SCR Description: Calculate Patient Age from cwshared method which is what the patient ID bar is using so Dental will match Patient ID bar all the time.

     If (cwobjPatient.DateOfBirth = DateValue("31-Dec-9999")) Or cwobjPatient.DateOfBirth = 0 Then
        patientAge = 20
     Else
        ' Get the patient demographic data available
        ' Age in days = difference between birthdate and encounter time
        ' Need age in minutes
        sCalcAge = gobjShared.CalcAge(cwobjPatient.DateOfBirth)
        
        If InStr(sCalcAge, "mo") > 0 Then
            If CInt(Left(sCalcAge, Len(sCalcAge) - 2)) < 12 Then
                patientAge = 0
            Else
                patientAge = CInt(Left(sCalcAge, Len(sCalcAge) - 2) / 12)
                
            End If
        Else
        
            patientAge = CInt(Left(sCalcAge, Len(sCalcAge) - 2))
        End If
    End If
    
    If Not oSVGComm Is Nothing Then
        'unload svgcomm
        oSingleton.TerminateSF
        
        'unload oteeth
        Dim oSF As New SVGTeeth.SF
        oSF.TerminateSF
        
        
        'Load svgcomm
        'why is this terminating svgcomm after it inits?
        Set oSVGComm = oSingleton.Singleton
        
'        Set oSVGComm.TeethObject.TeethDocDx = Nothing
'        Set oSVGComm.TeethObject.TeethDocTx = Nothing
        
        With oSVGComm
            Set .CWComm = cwobjCWComm
            Set .CWConfig = cwobjConfig
            Set .CWLogonEnviron = cwobjLogonEnviron
            Set .CWPatient = cwobjPatient
            
            .DentalEncounterID = 999
            .ProviderID = dentalProviderID
            .PatientID = patientUnitID '87214
            .Age = patientAge
        End With
        
        If gobjShared.CmdLineSwitches("IMAGING") = "IMAGING" Then
            'Reset the image filter form
            Set mImgFilterForm = Nothing
            Set mImgFilterForm = New clsImgFilterForm
            Set mImgFilterForm = mImgFilterForm.SingletonFilterForm
            mImgFilterForm.init oSVGComm.DataObject, gobjShared, oSVGComm.mObjSVGMngr
        End If

        'load oteeth
        Set oTeeth = oSF.Singleton
        oTeeth.Initialize
        'loadTeeth
    End If
    
    'Screen.MousePointer = vbDefault
    Debug.Print "svg patient selected: " & Timer - lStart
    
'End If
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

   

End Sub

Private Sub Terminate()

Dim oSF As SVGTeeth.SF

    'Shutdown everything just as if we were changing patients.
    If Not oSVGComm.mObjSVGMngr Is Nothing Then
        'Stop all the timers
        oSVGComm.mObjSVGMngr.ofrmMenu.tmrMultithread.Enabled = False
        oSVGComm.mObjSVGMngr.mObjSVGQueue.Cancel = True
    End If
    If Not oTeeth Is Nothing Then
        If Not oTeeth.frmTeethX Is Nothing Then
            oTeeth.frmTeethX.TxTimer.Enabled = False
            oTeeth.frmTeethX.tmrFinished.Enabled = False
        End If
    End If
    
    If Not oSVGComm Is Nothing Then
        'unload svgcomm
        oSingleton.TerminateSF
        
        'unload oteeth
        Set oSF = New SVGTeeth.SF
        oSF.TerminateSF
    End If
    
    If Not mImgFilterForm Is Nothing Then
        'mImgFilterForm.Terminate
        Set mImgFilterForm = Nothing
    End If
    
    Set oSingleton = Nothing
    Set oTeeth = Nothing
    Set oSVGComm = Nothing
    Set oSF = Nothing

End Sub
Private Sub Class_Initialize()

'#If DEBUGON Then
'   Set DL = GetObject("", "dlog.cdebuglog")
'   DL.debugStart
'   Dim DR  As DebugRec
'
'   DL.DebugLog Exe, Module, "Class_Initialize", "Begin", DR
'#End If
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
   
'#If DEBUGON Then
'   DL.DebugLog Exe, Module, "Class_Initialize", "End", DR
'#End If

End Sub

Private Sub Class_Terminate()
    
    Terminate

    Set oSVGComm = Nothing
    Set mImgFilterForm = Nothing
    Set oSingleton = Nothing
    Set oTeeth = Nothing
'#If DEBUGON Then
'   Dim DR  As DebugRec
'
'   DL.DebugLog Exe, Module, "Class_Terminate", "Begin", DR
'#End If
   '-- -release reference to objects---
   '-- Set gobjAnPCollection = Nothing

        Set gobjShared = Nothing
        Set cwobjCWComm = Nothing
        Set cwobjLogonEnviron = Nothing
        Set cwobjConfig = Nothing
        Set cwobjPatient = Nothing
        Set gobjUser = Nothing
   
   
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
'#If DEBUGON Then
'   DL.DebugLog Exe, Module, "Class_Terminate", "End", DR
'   Set DL = Nothing
'#End If

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

Dim arBuff() As String
Dim oSF As SVGTeeth.SF

   On Error GoTo ErrHandler
   
   '-- -----------------------------------------------------------------------------
   '--  Take corresponding actions based upon the message "type".
   '--
   '--  NOTE: This method is required, but currently there are NO REQUIRED messages
   '--        to process.
   '-- -----------------------------------------------------------------------------
   Select Case viMessageType
   
        Case cwiGENERIC
            arBuff = Split(vsMessageText, "|")
            Select Case UCase$(arBuff(0))
                Case "ISSVGDRAWN"
                    cwobjCWComm.Message cwiGENERIC, "SVGDRAWN|", NCID_DDSA_APPLICATION, NCID_DENTAL_SVG
                Case "EXIT_APP"
                    'Need a core change so this gets called
                    Terminate

                Case Else
                    'MsgBox "Unknown generic message type."
                    
            End Select
        
        Case cwiOK_TO_CLOSE_VIEW
        
        Case cwiREFRESH_DATA
            If vsMessageText = "REFRESHPATIENT" Then
                'This called from Dental.CompletedEncounter so that GC can reload the patient data.
                oSVGComm.RefreshPatient
            Else
                'gobjShared.ShowVBError 0, "I recieved cwiREFRESH_DATA message, but I don't know what to do with it.", "Dental_SVG.Message", "Dental_SVG", vbExclamation
            End If
        
        Case cwiCLEAR_PATIENT
            'Clear the objects
            If Not oSVGComm Is Nothing Then
                'unload svgcomm
                oSingleton.TerminateSF
                
                'unload oteeth
                Set oSF = New SVGTeeth.SF
                oSF.TerminateSF
                
            End If
            
   End Select
   
   Exit Sub
   
ErrHandler:
   '-- -show the VB error---
   Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental_SVG.Message", "Dental_SVG.Message", vbExclamation)
   
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



    patientUnitID = cwobjPatient.UnitNumber


    Set oSVGComm = oSingleton.Singleton

   '-- -set global flags---
   cwbMainViewRunning = False
   
   '-- ****************************************************************************
   '--  TO DO:
   '--    Add code here to create objects, initialize global variables, etc.
   '-- ****************************************************************************
   
   
   Exit Sub
   
ErrHandler:
   '-- -show the VB error---
   Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental_BigTooth.Initialize", "BigTooth", vbExclamation)
   
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
   
   dentalProviderID = 999 'gobjUser.GetAttribute("NCID")
   
   Dim objRoles      As ICHCSII_Roles
   Dim objResources  As ICHCSII_Resources
   
   'Screen.MousePointer = vbHourglass
   
   Call Initialize(ByVal vlMyProgId, robjComm, robjLogonEnvironment, robjConfig, robjPatient)
   
   'Screen.MousePointer = vbDefault
   
   Exit Sub
   
ErrHandler:
   Call gobjShared.ShowVBError(Err.Number, Err.Description, "in frmGraphicalCharting.Initialize", "BigTooth !", vbExclamation)
   
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
   Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental_BigTooth.Properties", "BigTooth", vbExclamation)
   
   Exit Sub

End Sub

Private Function ICHCSII_CoreAPI3_CloseView2(ByVal viViewType As Integer) As Boolean
  ICHCSII_CoreAPI3_CloseView2 = True
End Function
'Private Function ICHCSII_Dental_SetEncounterOps(objEncounterOps As Object) As Boolean
'  Dim Age As Single
'  If Not objEncounterOps Is Nothing Then
'    Set gobjEnc = objEncounterOps
'    dentalProviderID = gobjEnc.CurrentEncounter.Providers(1).NCID
'  If (cwobjPatient.DateOfBirth = DateValue("31-Dec-9999")) Then
'       patientAge = 20
'     Else
'       ' Get the patient demographic data available
'       ' Age in days = difference between birthdate and encounter time
'       ' Need age in minutes
'       Age = DateDiff("d", Format(cwobjPatient.DateOfBirth, "mm/dd/yyyy"), Format(gobjEnc.CurrentEncounter.StartDTS, "mm/dd/yyyy"))
'       patientAge = Int(Age / 364.5)
'     End If
'     patientUnitID = cwobjPatient.UnitNumber
'     encounterID = gobjEnc.CurrentEncounter.encounterID
'        If Not oSVGComm Is Nothing Then
'           oSVGComm.DentalEncounterID = DentalEncounterID
'           oSVGComm.ProviderID = dentalProviderID
'           oSVGComm.PatientID = patientUnitID
'           Set oSVGComm.TeethObject.TeethDocTx = Nothing
'           'oSVGComm.initializeUser
'        End If
'  End If
'End Function

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

Public Function loadTeeth()

    '*** MsgBox Fix - Code_Validator auto change made 3/21/2003
    MsgBxARMd "svg.loadteeth, should not be called"
    Exit Function
    
Dim oSF As New SVGTeeth.SF

    oSF.TerminateSF
    Set oTeeth = oSF.Singleton
    oTeeth.Initialize
             
End Function

Private Sub oTeeth_AllFinished()
    
    'Send a message to Dental that the drawing is finished
    cwobjCWComm.Message cwiGENERIC, "SVGDRAWN|", NCID_DDSA_APPLICATION, NCID_DENTAL_SVG

End Sub

Private Sub oTeeth_DrawingFinished()
  isDrawingFinished = True
End Sub
