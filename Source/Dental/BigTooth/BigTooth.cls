VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BigTooth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public patientAge As Integer
Public patientUnitID As Long
Public DentalEncounterID As String
Public encounterID As Long
Public dentalProviderID As Long
Public oSingleton As New srvSVGComm.SingletonFactory
Public oSVGComm As srvSVGComm.clsSVGComm
Private mstrSupProcs As String
Private mobjExamGeneral As CHCSII_Dental_NotesPlus.clsExamGeneral
'Private moSingletonFilterForm As clsImgFilterForm

Dim oSF As New SVGTeeth.SF
Public WithEvents oTeethSVG As SVGTeeth.clsTeeth
Attribute oTeethSVG.VB_VarHelpID = -1
Private WithEvents oToothSVG As srvSVGComm.clsTooth
Attribute oToothSVG.VB_VarHelpID = -1


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
Enum eStatus
  eNew
  eUpdate
  eDelete
  eNone
End Enum
Enum eView
  eDxView
  eTxView
End Enum
Private m_PatientID As Long
Private m_Age As Integer


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
      SetParent frmGraphicalCharting.hwnd, vlParenthWnd
      Set frmGraphicalCharting.ParentObject = Me
      '-- -set parameter---
      rhMainForm = frmGraphicalCharting.hwnd
      
   ''''''''''''''''''''''''''''''''''''''''''''' new code start  '''''''''''''''''''''''''''''''''''''''''''''''''
   
    Set oSVGComm = oSingleton.Singleton
    Set frmGraphicalCharting.FMPSRX1.cwobjCWComm = cwobjCWComm
    
    'oSVGComm.DentalEncounterID = DentalEncounterID
    'oSVGComm.ProviderID = dentalProviderID
    'oSVGComm.PatientID = patientUnitID
    'oSVGComm.Age = patientAge
    
    frmGraphicalCharting.fgDxTx.Notes = oSVGComm.DataObject
    frmGraphicalCharting.fgDxTx.GridLocation = GridLocationEnum.eFullMouthDxV

'    ' load teeth and it is visible
'    Call loadTeeth(patientUnitID, patientAge)
'
'    Call frmGraphicalCharting.showTeeth
'
'    ' load tooth but it is not visible
'    loadTooth

   ''''''''''''''''''''''''''''''''''''''''''''' new code end '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      
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
   Call gobjShared.ShowVBError(Err.Number, Err.Description, "Graphical_Charting.OpenView", "Graphical_Charting", vbCritical)
   
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
           'oTeethSVG.putTeethBack
           'SetParent frmGraphicalCharting.elTeeth.hwnd, oTeethSVG.getHwnd
           'oTeethSVG.putTeethBack frmGraphicalCharting.elTeethX.hwnd

           ChangeToFullMouth
    
           'Call SetParent(oTeethSVG.getChildHwnd, oTeethSVG.getParentHwnd)
           oTeethSVG.ReleaseTeethControl
           
           Unload frmGraphicalCharting
           terminate_objects
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
         
         '-- -close ALL forms that are associated with all "views"---
'        If gobjEnc Is Nothing Then
        '   oTeethSVG.putTeethBack frmGraphicalCharting.elTeeth.hwnd
           ChangeToFullMouth
           'Call SetParent(oTeethSVG.getChildHwnd, oTeethSVG.getParentHwnd)
           oTeethSVG.ReleaseTeethControl
           
           Unload frmGraphicalCharting
           terminate_objects
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

Dim tFilter As tRadFilter
Dim oMyRads As colReadOnlyRads

   On Error GoTo ErrHandler
   
   '-- --------------------------------------------------------------------------------
   '--  NOTE: You don't have to create menu items, these are optional, that is why
   '--        the code is commented out.
   '-- --------------------------------------------------------------------------------
'    '---perform the actions associated with the menu item that was selected---
    Select Case vlMenuId
   
        Case glMenuRefresh
        
                Screen.MousePointer = vbHourglass
                'Call SetParent(oTeethSVG.getChildHwnd, oTeethSVG.getParentHwnd)
                oTeethSVG.ReleaseTeethControl
                
                DoEvents
                
                If frmGraphicalCharting.bMultiToothMode = True Then
                        'Full mouth
                        '37476
                        oSVGComm.UpdateTeeth frmGraphicalCharting.ViewMode
                        
                Else
                        'Big tooth
                        '37476
                        oSVGComm.ToothObject.UpdateTooth oSVGComm.ToothObject.CurrentViewTooth, frmGraphicalCharting.ViewMode
                         'oSVGComm.ToothObject.UpdateTooth oSVGComm.ToothObject.CurrentViewTooth, frmGraphicalCharting.IsDxView
                End If
                
                frmGraphicalCharting.tmrTeeth.Enabled = True
                Screen.MousePointer = vbDefault
                
        Case glSupProc
    
            Dim oList As New CStandardList
            Dim oItem As CMultiListItem
            Dim sTemp As String
    
            Screen.MousePointer = vbHourglass
    
            'Load the default values
            If mstrSupProcs = "" Then
                mstrSupProcs = LoadSupportingProcedures()
            End If
    
            oList.Load_String mstrSupProcs
    
            Screen.MousePointer = vbDefault
            oList.Form_Caption = "Supporting Procedures"
    
            If oList.Get_User_Input Then
                'Update the object
                Screen.MousePointer = vbHourglass
                mstrSupProcs = oList.ValueString
            End If
        
        Case glMenuRadPano
            'Display the latest pano
            tFilter.RadType = eRadType.Pano
            Set oMyRads = oSVGComm.DataObject.Rads.Find(tFilter)
            If oMyRads.Count > 0 Then
'                moSingletonFilterForm.ImageViewer.LoadRad oMyRads(1)
            Else
                MsgBox "There are not any panos for this patient.", vbOKOnly + vbInformation, "No panos"
            End If
        
        Case glMenuRadFMSeries
            'Display the latest FM Series
            tFilter.RadType = eRadType.FMSeries
            Set oMyRads = oSVGComm.DataObject.Rads.Find(tFilter)
            If oMyRads.Count > 0 Then
'                moSingletonFilterForm.ImageViewer.LoadRad oMyRads(1)
            Else
                MsgBox "There are not any FM series for this patient.", vbOKOnly + vbInformation, "No full mouth series"
            End If
        
        Case glMenuRadBW
            'Display the latest bitewing
            If oSVGComm.ToothObject.CurrentViewTooth = 0 Then
                MsgBox "Please select a tooth."
                Exit Sub
            End If
            
            tFilter.RadType = eRadType.BiteWing
            tFilter.Location = oSVGComm.ToothObject.CurrentViewTooth
            Set oMyRads = oSVGComm.DataObject.Rads.Find(tFilter)
            If oMyRads.Count > 0 Then
'                moSingletonFilterForm.ImageViewer.LoadRad oMyRads(1)
            Else
                MsgBox "There are not any bitewings that include tooth #" & oSVGComm.ToothObject.CurrentViewTooth & " for this patient.", vbOKOnly + vbInformation, "No bitewings"
            End If
        
        Case glMenuRadPA
            'Display the latest PA
            If oSVGComm.ToothObject.CurrentViewTooth = 0 Then
                MsgBox "Please select a tooth."
                Exit Sub
            End If
            
            tFilter.RadType = eRadType.PA
            tFilter.Location = oSVGComm.ToothObject.CurrentViewTooth
            Set oMyRads = oSVGComm.DataObject.Rads.Find(tFilter)
            If oMyRads.Count > 0 Then
'                moSingletonFilterForm.ImageViewer.LoadRad oMyRads(1)
            Else
                MsgBox "There are not any PAs that include tooth #" & oSVGComm.ToothObject.CurrentViewTooth & " for this patient.", vbOKOnly + vbInformation, "No PAs"
            End If
        
        Case glMenuRadImages
            'Display the Image Filter
'             moSingletonFilterForm.ShowFilterForm
    
    End Select

   Set oList = Nothing
   Set oItem = Nothing
   
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
   
    '43818
    If vlWidth < 0 Then
        vlWidth = 0
    End If
    If vlHeight < 0 Then
        vlHeight = 0
    End If
    
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
            frmGraphicalCharting.WindowState = vbNormal
    
            '-- -hide the form to so that annyoing "screen blinking" will not be seen while controls are moved and resized---
            frmGraphicalCharting.Visible = False
    
            '-- -resize form to fill entire "client area", NOTE: This generates a "resize" event---
            frmGraphicalCharting.Move cwlLeft, cwlTop, cwlWidth, cwlHeight
    
            '<bjr090401 - moved to after updatePSR, etc.
            'frmGraphicalCharting.Visible = True
    
            frmGraphicalCharting.UpdatePSR
    
            ' update DxTxGrid
            'Call frmGraphicalCharting.fgDxTx.Charts.LoadDxs(patientUnitID, CInt(DentalEncounterID))
            'frmGraphicalCharting.fgDxTx.RefreshGrid True
            'frmGraphicalCharting.fgDxTx.HideColumn frmGraphicalCharting.fgDxTx.FlexGrid.cols - 1, True
            'frmGraphicalCharting.fgDxTx.SetYellowCell = True
    
            ' refresh
            '<bjr090401 [SCR15676] - removed - we don't need this during a resize
            'frmGraphicalCharting.cmdView(0).Value = True
    
            '<bjr090401 - this should be last - moved from above
            'cwiFOREGROUND_APPLICATION

            frmGraphicalCharting.Visible = True
            oTeethSVG.SetTeethControl frmGraphicalCharting.elTeeth.hwnd
         
        Else       '--first "regular" resize, need to perform "initialization" OrderSetss--
            '-- -reset flag so this code will not be executed again---
            cwbFirstRegularResize = False
            
            '-- --------------------------------------------------------------------------
            '--  NOTE: Here is where you "Show" the "Regular View" form for the FIRST time.
            '-- --------------------------------------------------------------------------
             'If Not gobjEnc Is Nothing Then
             '  glMenuSave = gobjEnc.AddMenuItem(0, "&Save", "save.bmp", "Save Dental")
                'Call frmGraphicalCharting.showTeeth
            
            Set oTeethSVG = oSF.Singleton
            Set oToothSVG = oSVGComm.ToothObject
         
         
            If Not gobjDental Is Nothing Then
               glMenuRefresh = gobjDental.AddMenuItem(0, "&Refresh", "Refresh.bmp", "Refresh Graphical Charting")
               'Kdunne 7/23/2002
               'SCR 19977
               glSupProc = gobjDental.AddMenuItem(0, "S&up. Procs.", "MedcinLists.bmp", "Supporting Procedures")
        
               AddRadMenus frmGraphicalCharting.bMultiToothMode, oSVGComm
            End If
          
            'SCR 45179,45989,45991 RWalker 12/12/2003
            gbFromCore = True
              
            frmGraphicalCharting.Display
            ' load tooth but it is not visible, it has to be first
            Call loadTooth
                    
            ' load teeth and it is visible
            'Call loadTeeth
            
            Call frmGraphicalCharting.fgDxTx.RefreshGrid(False)
            
            
            frmGraphicalCharting.tmrTeeth.Enabled = True
        
            'Else
            '  CloseView cwiREGULAR_VIEW
            '  Exit Sub
            'End If
            
            'SCR 45179,45989,45991 RWalker 12/12/2003
            gbFromCore = False
        
        End If

    End If
   
    Exit Sub
   
ErrHandler:
   '-- -show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental_BigTooth.ResizeView", "Dental", vbExclamation)
   
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
   
   '--- SCR 28290 RWalker 10/10/02 Instantiate ARM / Init transactions
   Call ARM_InitAllTransactionsForDental_BigTooth
   
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
 
   terminate_objects
'        Set moSingletonFilterForm = Nothing
        Set gobjShared = Nothing
        Set cwobjCWComm = Nothing
        Set cwobjLogonEnviron = Nothing
        Set cwobjConfig = Nothing
        Set cwobjPatient = Nothing
        Set gobjUser = Nothing
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
        Case cwiFOREGROUND_APPLICATION
            'Refresh the tab
            'frmGraphicalCharting.tabGraphicCharting_Switch 0, frmGraphicalCharting.tabGraphicCharting.CurrTab, 0
            
            If oTeethSVG Is Nothing Then
                'Don't try to refresh if we are not in graphic charting.
                Exit Sub
            End If
            
            oSVGComm.mbReadOnly = False
            
            'Set the mode
            '37476
            oTeethSVG.BrowserViewMode = frmGraphicalCharting.ViewMode
            'oTeethSVG.setDxTxMode frmGraphicalCharting.IsDxView
            
            'Get the teeth graphic.
            oTeethSVG.SetTeethControl frmGraphicalCharting.elTeeth.hwnd
            
            'SCR 33870
            'ChangeToFullMouth
   End Select
   
   Exit Sub
   
ErrHandler:
   '-- -show the VB error---
   Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental_BigTooth.Message", "BigTooth", vbExclamation)
   
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
   
   
   Exit Sub
   
ErrHandler:
   '-- -show the VB error---
   Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental_BigTooth.Initialize", "BigTooth", vbExclamation)
   
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
    
    On Error GoTo ErrHandler
    
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
    
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "in BigTooth.GetDentalEncounterID()", "BigTooth", vbCritical)
    
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
    DentalEncounterID = gobjDental.psDentalEncounterID
  Else
    MsgBxARMd ("Dental Parent should never be nothing but just incase:")
  End If
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
  Dim Age As Single
  Dim oEncDoc As Encounter
  
  On Error GoTo ErrHandler
   
  If Not objEncounterOps Is Nothing Then
    Set gobjEnc = objEncounterOps
 'SCR 25225
  Set oEncDoc = gobjEnc.CurrentEncounter
    dentalProviderID = oEncDoc.Providers(1).NCID
  If (cwobjPatient.DateOfBirth = DateValue("31-Dec-9999")) Then
       patientAge = 20
     Else
       ' Get the patient demographic data available
       ' Age in days = difference between birthdate and encounter time
       ' Need age in minutes
       Age = DateDiff("d", Format(cwobjPatient.DateOfBirth, "mm/dd/yyyy"), Format(oEncDoc.StartDTS, "mm/dd/yyyy"))
       patientAge = Int(Age / 364.5)
     End If
     patientUnitID = cwobjPatient.UnitNumber
     encounterID = oEncDoc.encounterID
  End If

Exit Function
ErrHandler:
 '-- -show the VB error---
   Call gobjShared.ShowVBError(Err.Number, Err.Description, "Dental_BigTooth.ICHCSII_Dental_SetEncounterOps", "Dental Big Tooth", vbExclamation)
   
   Exit Function
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
   Set gobjEnc = Parent
   IChild_SetParent = True
End Function

Public Function Mtd_bDentalTabChange(ByVal lFrom As Long, ByVal lTo As Long, Optional ByVal bFromCore As Boolean = False) As Boolean

'SCR 45991 RWalker 12/12/2003 Since the original Open BigTooth has been replaced with either the FullMouthGC or
'BigToothGC transaction we need to fire these when the BigTooth tab is clicked, but only when clicked on the Dental form
'NOT in the Core Folder List.  Hence the addition of the optional parameter to denote when Core initiated this event

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
   
    If Not cwbMainViewRunning Then
        Mtd_bDentalTabChange = True
        Exit Function
    End If
   
    'SCR 45991 RWalker 12/12/2003 Since the original Open BigTooth
    'has been replaced with either the FullMouthGC or BigToothGC
    'transaction we need to fire these when the BigTooth tab is clicked,
    'but only when clicked on the Dental form NOT in the Core Folder List
    If (Not bFromCore) And (lTo = 0) Then
        Call frmGraphicalCharting.ARM_Tab_Switch(frmGraphicalCharting.tabGraphicCharting.CurrTab, "Start")
    End If
   
    If lFrom = 0 Then
        frmGraphicalCharting.elTeeth.Caption = "Saving Graphics Please Wait ..."
        DoEvents
        Call SaveView
        ChangeToFullMouth
        'Call SetParent(oTeethSVG.getChildHwnd, oTeethSVG.getParentHwnd)
        oTeethSVG.ReleaseTeethControl
        
        oSVGComm.TeethObject.ResetSelection
        oSVGComm.mbReadOnly = True
        cwobjCWComm.DeleteMenuItem (glMenuRefresh)
        cwobjCWComm.DeleteMenuItem (glSupProc) 'scr 19977
        DeleteRadMenus
        cwobjCWComm.Message cwiGENERIC, "Refresh", NCID_CW_APPLICATION, NCID_DENTAL_BIGTOOTH_APPLICATION 'SCR 26478
    End If
    
    If lTo = 0 Then
        frmGraphicalCharting.elTeeth.Caption = "Loading Graphics Please Wait ..."
        DoEvents
        Call frmGraphicalCharting.fgDxTx.RefreshGrid(False)
        oSVGComm.mbReadOnly = False
        frmGraphicalCharting.doMenuMagic
        frmGraphicalCharting.tmrTeeth.Enabled = True
        glMenuRefresh = gobjDental.AddMenuItem(0, "&Refresh", "Refresh.bmp", "Refresh Graphical Charting")
        glSupProc = gobjDental.AddMenuItem(0, "S&up. Procs.", "MedcinLists.bmp", "Supporting Procedures")
        AddRadMenus frmGraphicalCharting.bMultiToothMode, oSVGComm
    End If
   
    'SCR 45991 RWalker 12/12/2003 Since the original Open BigTooth
    'has been replaced with either the FullMouthGC or BigToothGC
    'transaction we need to fire these when the BigTooth tab is clicked,
    'but only when clicked on the Dental form NOT in the Core Folder List
    If (Not bFromCore) And (lTo = 0) Then
        Call frmGraphicalCharting.ARM_Tab_Switch(frmGraphicalCharting.tabGraphicCharting.CurrTab, "Stop")
    End If
       
   Mtd_bDentalTabChange = True
   
Exit Function
ErrHnd:
   Mtd_bDentalTabChange = False

Exit Function
End Function

Public Function SaveView() As Boolean
  Dim strSQL As String
  

'If the app times out, then GC is not visible, but it should still save.
'So, don't check if it is visible.

'  If frmGraphicalCharting.Visible Then
    '>> SCR #8510;   Developer: Jon Mack 04/16/2002 10:03 AM
    'SCR Description: Save PSR so all can see the changes


    frmGraphicalCharting.fgDxTx.Notes.SaveGroups
    
    If frmGraphicalCharting.DirtyDxTx Then
        gobjDental.UpdateDentalEncDRC
        frmGraphicalCharting.CleanDxTx
    End If
    
    'SF 8-14-03 39609
    If Not oSVGComm Is Nothing Then
        oSVGComm.TeethObject.setTeethPath patientUnitID
    End If
    
    frmGraphicalCharting.FMPSRX1.saveData
    
    '37476
    If frmGraphicalCharting.ViewMode = eDxView Then
    'If frmGraphicalCharting.IsDxView Then
        'svg
        'frmGraphicalCharting.cmdView(0).Value = True
        'frmGraphicalCharting.cmdView(4).Value = True
    ElseIf frmGraphicalCharting.ViewMode = eTxView Then
        'svg
        'frmGraphicalCharting.cmdView(1).Value = True
        'frmGraphicalCharting.cmdView(5).Value = True
    ElseIf frmGraphicalCharting.ViewMode = eDx603aView Then
    
    End If
'    SaveView = True
'  Else
'    SaveView = False
'  End If
  
  SaveSupportingProcedures mstrSupProcs
  
  SaveView = True
End Function

Public Sub terminate_objects()
  Set oSVGComm = Nothing
  Set oTeethSVG = Nothing
  Set oSingleton = Nothing
  Set oSF = Nothing
  Set gobjUser = Nothing
  Set gobjDental = Nothing
End Sub


'Private Function loadTooth()
'    ' this is to simulate graphic charting
'    ' call svg server
'    Dim oSingleton As New srvSVGComm.SingletonFactory
'    Dim oSVGComm As srvSVGComm.clsSVGComm
'    Dim S As String
'
'    Set oSVGComm = oSingleton.Singleton
'
'    S = oSVGComm.getToothHTML()
'    'Call frmGraphicalCharting.toothSVGX1.loadHTML(S)
'    Call frmGraphicalCharting.toothSVGX1.loadHTML(S)
'
'
'End Function
'
'Private Function loadTeeth(ByVal nPatientID As Long, ByVal nAge As Integer)
'    ' this is to simulate graphic charting
'    ' call svg server
'    Dim oSingleton As New srvSVGComm.SingletonFactory
'    Dim oSVGComm As srvSVGComm.clsSVGComm
'    Dim S As String
'
'    Set oSVGComm = oSingleton.Singleton
'
'    oSVGComm.PatientID = nPatientID
'    oSVGComm.Age = nAge
'
'    S = oSVGComm.getTeethHTML()
'    Call frmGraphicalCharting.teethSVGX1.loadHTML(S)
'
'End Function
Private Function loadTooth()
    ' this is to simulate graphic charting
    ' call svg server
    Dim oSingleton As New srvSVGComm.SingletonFactory
    Dim oSVGComm As srvSVGComm.clsSVGComm
    Dim S As String
   
    Set oSVGComm = oSingleton.Singleton
     
    S = oSVGComm.getToothHTML()
    Call frmGraphicalCharting.toothSVGX1.loadHTML(S)
    
    
    If gobjShared.CmdLineSwitches("IMAGING") = "IMAGING" Then
        'Get the filter form instance
'        Set moSingletonFilterForm = New clsImgFilterForm
'        Set moSingletonFilterForm = moSingletonFilterForm.SingletonFilterForm
    End If
    
End Function
 


Public Function startServer()
Dim oSF As New SVGTeeth.SF

  
     
  ''''''''''''''''''''''''''''''''''''''''''''' new code start  '''''''''''''''''''''''''''''''''''''''''''''''''
    Dim oSingleton As New srvSVGComm.SingletonFactory
    Dim oSVGComm As srvSVGComm.clsSVGComm

    Set oSVGComm = oSingleton.Singleton

'    oSVGComm.PatientID = 999 '87214
'    oSVGComm.DentalEncounterID = 999
    patientUnitID = 999
    oSVGComm.Age = 18
    dentalProviderID = 2100

    ' load tooth but it is not visible, it has to be first
    Set oTeethSVG = oSF.Singleton
    
    
   ' load the form
    Load frmGraphicalCharting
    Set frmGraphicalCharting.ParentObject = Me
    frmGraphicalCharting.Show
    frmGraphicalCharting.tabGraphicCharting.CurrTab = 0
    
    Call loadTooth
            
    ' load teeth and it is visible
    'Call frmGraphicalCharting.showTeeth
    'Call loadTeeth
    frmGraphicalCharting.tmrTeeth.Enabled = True
    
    
    frmGraphicalCharting.fgDxTx.Notes = oSVGComm.DataObject
    
    'Call frmGraphicalCharting.fgDxTx.RefreshGrid(False)
    
    
    'oSVGComm.createDentalSVG True
'    ' load teeth and it is visible
'    Call loadTeeth(patientUnitID, patientAge)
'
'    Call frmGraphicalCharting.showTeeth
'
'    ' load tooth but it is not visible
'    loadTooth

   ''''''''''''
    
End Function

Private Sub oTeethSVG_DblClickMultiTooth(iToothID As Integer)
    
    'SCR 45991 RWalker 12/12/2003 Merge 3 BigTooth Simple transactions
    'and 1 Complex transaction from Core into 1 Complex transaction
    'with the data type value in  metric 7
    '--- SCR-ARM-24401 RWalker 12/12/2003 - CHCSII_Dental_BigTooth
    If frmGraphicalCharting.ViewMode = eDxView Then
        goARM.StartCmplxTx gsARMTx_OpenDentalBigToothGC, "Dx"
    Else
        goARM.StartCmplxTx gsARMTx_OpenDentalBigToothGC, "Tx"
    End If
    
    Call frmGraphicalCharting.showTooth2(iToothID)
    frmGraphicalCharting.refreshSingleTooth
    oSVGComm.ToothObject.MenuUpdateBasedOnTooth
    
    '--- SCR-ARM-24401 RWalker 12/12/2003 - CHCSII_Dental_BigTooth
    goARM.StopCmplxTx gsARMTx_OpenDentalBigToothGC, enmARMTxStatusOk, enmARMStop_MeOnly

End Sub
Public Sub ChangeToFullMouth()
           If Not frmGraphicalCharting.bMultiToothMode Then
                oTeethSVG.showTeeth
           End If
End Sub

Private Sub oTeethSVG_SelectionCountChange(bSingleSelected As Boolean)
Dim SetOfTeeth As String
Dim ID As Integer

        SetOfTeeth = oSVGComm.TeethObject.getSelectedTeeth
        
        If SetOfTeeth <> "" And InStr(SetOfTeeth, ",") <= 0 Then
            'One tooth selected
            ID = CInt(SetOfTeeth)
            
            '37476
            If frmGraphicalCharting.ViewMode = eDxView Then
            'If frmGraphicalCharting.IsDxView Then
                frmGraphicalCharting.cmdOrientX.Enabled = True
            ElseIf frmGraphicalCharting.ViewMode = eTxView Then
                frmGraphicalCharting.cmdOrientX.Enabled = False
            ElseIf frmGraphicalCharting.ViewMode = eDx603aView Then
                'Do nothing
            End If
            
            If (ID >= 4 And ID <= 13) Or _
               (ID >= 33 And ID <= 42) Or _
               (ID >= 20 And ID <= 29) Or _
               (ID >= 43 And ID <= 52) Or oSVGComm.TeethObject.Positions.GetPositionByToothNum(ID).ToothErupted = False Then
               
                frmGraphicalCharting.cmdPrimarySecondary.Enabled = True
            Else
                frmGraphicalCharting.cmdPrimarySecondary.Enabled = False
            End If
        Else
            'Multiple teeth selected or no teeth selected
            frmGraphicalCharting.cmdOrientX.Enabled = False
            frmGraphicalCharting.cmdPrimarySecondary.Caption = "2°    to    1°"
            frmGraphicalCharting.cmdPrimarySecondary.Enabled = False

        End If
    
End Sub

Private Sub oTeethSVG_SingleClickMultiTooth(iToothID As Integer)
    
    'SCR 45989 RWalker 12/12/2003 Merge 3 FullMouth Simple transactions
    'and 1 Complex transaction from Core into 1 Complex transaction
    'with the data type value in  metric 7
    '--- SCR-ARM-24407 RWalker 12/12/2003 - CHCSII_Dental_BigTooth
    If frmGraphicalCharting.ViewMode = eDxView Then
        goARM.StartCmplxTx gsARMTx_OpenDentalFullMouthGC, "Dx"
    Else
        goARM.StartCmplxTx gsARMTx_OpenDentalFullMouthGC, "Tx"
    End If
    
    Call frmGraphicalCharting.showTeeth2

    frmGraphicalCharting.refreshMultiTooth
    oSVGComm.ToothObject.MenuUpdateBasedOnTooth
    
    '--- SCR-ARM-24407 RWalker 12/12/2003 - CHCSII_Dental_BigTooth
    goARM.StopCmplxTx gsARMTx_OpenDentalFullMouthGC, enmARMTxStatusOk, enmARMStop_MeOnly
    
End Sub

'Private Sub teethSVGX1_dblClickMultiTooth(RHV As Integer)
'    'MsgBxARMd "dblclick tooth from full mouth - " & RHV
'    Call showTooth2(RHV)
'    refreshSingleTooth
'End Sub
'
'Private Sub teethSVGX1_singleClickMultiTooth(RHV As Integer)
'    'MsgBxARMd "dblclick tooth from a single tooth - " & RHV
'    Call showTeeth2
'
'    refreshMultiTooth
'End Sub
'
'Private Sub teethSVGX1_SelectionCountChange(bSingleSelected As Boolean)
'    'MsgBxARMd "SelectionCountChange - " & bSingleSelected
'End Sub
'
''Private Sub teethSVGX1_ToothChanged(iToothIDNew As Integer, iToothIDOld As Integer)
''    'MsgBxARMd "ToothChanged! New - " & CStr(iToothIDNew) & " Old - " & CStr(iToothIDOld)
''    If toothSVGX1.Visible Then
''        teethSVGX1.showPanel True, iToothIDNew
''        teethSVGX1.updateTeethPanel (iToothIDNew)
''    End If
''End Sub
'Private Sub teethSVGX1_ToothChanged(iToothIDNew As Integer, iToothIDOld As Integer)
'    'MsgBxARMd "ToothChanged! New - " & CStr(iToothIDNew) & " Old - " & CStr(iToothIDOld)
'    If toothSVGX1.Visible Then
'        teethSVGX1.showPanel True, iToothIDNew
'        teethSVGX1.updateTeethPanel (iToothIDNew)
'    End If
'
'    'Alek - 01/25/02
'    Dim SetOfTeeth As String
'    SetOfTeeth = oSVGComm.TeethObject.getSelectedTeeth
'    If SetOfTeeth <> "" And InStr(SetOfTeeth, ",") <= 0 Then
'
'        If mbIsDxView And Not frmMenuTool.OrientXEntry1.Visible Then
'            frmMenuTool.OrientXEntry1.Visible = True
'        End If
'
'        frmMenuTool.OrientXEntry1.ToothID = oSVGComm.ToothObject.CurrentViewTooth  'iToothIDNew
'        frmMenuTool.OrientXEntry1.PatientID = objparent.patientUnitID
'    End If
'    '
'
'End Sub
'Private Sub teethSVGX1_priSecondClick(currentToothID As Integer)
'    'MsgBxARMd "priSecondClick - " & CStr(currentToothID)
'End Sub


Private Sub oTeethSVG_ToothChanged(iToothIDNew As Integer, iToothIDOld As Integer, RefreshGrid As Boolean)
   
   If oSVGComm.ToothObject.CurrentViewTooth = 0 Then
      oSVGComm.ToothObject.ToothChanged iToothIDNew
   End If

   frmGraphicalCharting.lblToothName = oSVGComm.ToothObject.getToothName

   If iToothIDNew >= 1 And iToothIDNew <= 8 Then
      frmGraphicalCharting.lblEast = "D"
      frmGraphicalCharting.lblWest = "M"
   ElseIf iToothIDNew >= 9 And iToothIDNew <= 24 Then
      frmGraphicalCharting.lblEast = "M"
      frmGraphicalCharting.lblWest = "D"
   ElseIf iToothIDNew >= 25 And iToothIDNew <= 32 Then
      frmGraphicalCharting.lblEast = "D"
      frmGraphicalCharting.lblWest = "M"
   End If

  If iToothIDNew >= 33 And iToothIDNew <= 37 Then
      frmGraphicalCharting.lblEast = "D"
      frmGraphicalCharting.lblWest = "M"
   ElseIf iToothIDNew >= 38 And iToothIDNew <= 47 Then
      frmGraphicalCharting.lblEast = "M"
      frmGraphicalCharting.lblWest = "D"
   ElseIf iToothIDNew >= 48 And iToothIDNew <= 52 Then
      frmGraphicalCharting.lblEast = "D"
      frmGraphicalCharting.lblWest = "M"
   End If
   
   If iToothIDNew >= 1 And iToothIDNew <= 16 Then
      frmGraphicalCharting.lblNorth = "F"
      frmGraphicalCharting.lblSouth = "L"
   ElseIf iToothIDNew >= 17 And iToothIDNew <= 32 Then
      frmGraphicalCharting.lblNorth = "L"
      frmGraphicalCharting.lblSouth = "F"
   End If

   If iToothIDNew >= 33 And iToothIDNew <= 42 Then
      frmGraphicalCharting.lblNorth = "F"
      frmGraphicalCharting.lblSouth = "L"
   ElseIf iToothIDNew >= 43 And iToothIDNew <= 52 Then
      frmGraphicalCharting.lblNorth = "L"
      frmGraphicalCharting.lblSouth = "F"
   End If

   If iToothIDNew <= 32 Then
        frmGraphicalCharting.cmdPrimarySecondary.Caption = "2°    to    1°"
   Else
        frmGraphicalCharting.cmdPrimarySecondary.Caption = "1°    to    2°"
   End If

   
   If frmGraphicalCharting.fgDxTx.GridLocation = GridLocationEnum.eBigToothDxV _
   Or frmGraphicalCharting.fgDxTx.GridLocation = GridLocationEnum.eBigToothTxV _
   Or frmGraphicalCharting.fgDxTx.GridLocation = GridLocationEnum.eFullMouthDxV _
   Or frmGraphicalCharting.fgDxTx.GridLocation = GridLocationEnum.eFullMouthTxV Then
        If frmGraphicalCharting.fgDxTx.GridFilter = eThisTooth Then
            frmGraphicalCharting.fgDxTx.ToothIDFilter = iToothIDNew
            frmGraphicalCharting.fgDxTx.RefreshGrid False, True
        ElseIf RefreshGrid = True Then
            frmGraphicalCharting.fgDxTx.ToothIDFilter = iToothIDNew
            frmGraphicalCharting.fgDxTx.RefreshGrid False, False
        Else
            frmGraphicalCharting.fgDxTx.ToothIDFilter = iToothIDNew
        End If
    
    
    End If

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
    Dim objEnc As Encounter
    
    Set objEnc = gobjEnc.CurrentEncounter
    'Call SetParent(oTeethSVG.getChildHwnd, oTeethSVG.getParentHwnd)
    oTeethSVG.ReleaseTeethControl
    
    objEnc.PrintEncounter blnPreview
    frmGraphicalCharting.tmrTeeth.Enabled = True
    
    Set objEnc = Nothing
End Sub

Private Function LoadSupportingProcedures() As String
    Dim strReturn As String
    Dim objDAS As ICHCSII_DataAccess.ICHCSII_DAS
    Dim objRS As ADODB.Recordset
    Dim strSQL As String
    
    Set objDAS = GetDasObj
    
    strSQL = "select GC_SUPP_PROCS from TEMPL_EXAMGENERAL where Dental_Encounter_ID = " & _
        DentalEncounterID
    
    Set objRS = objDAS.OpenRecordset(strSQL)
    
    If Not objRS Is Nothing Then
        If objRS.EOF = False Then
            strReturn = objRS.Fields("GC_SUPP_PROCS") & ""
        End If
    End If
        
FuncExit:
    If Trim(strReturn) = Empty Then
        strReturn = gcSupportingProcs
    End If
    
    LoadSupportingProcedures = strReturn

    Set objDAS = Nothing
    Set objRS = Nothing
    
    Exit Function
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "BigTooth.LoadSupportingProcedures", "BigTooth", vbCritical)

    GoTo FuncExit
End Function

Private Sub SaveSupportingProcedures(ByVal strSuppProcs As String)
    Dim objMod As Object
    Dim objSection As CHCSII_Dental_NotesPlus.clsISection
    Dim oMultiList As NotesPlusUtilities.CStandardList
    
    'Don't save something that doesn't need to be saved.
    If strSuppProcs <> "" Then
        'Make sure that notes is loaded.
        Set objMod = cwobjCWComm.InitializeOLEServer(NCID_DENTAL_NOTESPLUS_APPLICATION)
        
        If mobjExamGeneral Is Nothing Then
            Set mobjExamGeneral = New CHCSII_Dental_NotesPlus.clsExamGeneral
            Set objSection = mobjExamGeneral
            
            Call objSection.LoadSection(patientUnitID, DentalEncounterID)
        End If
        
        mobjExamGeneral.GC_Supporting_Procedures = strSuppProcs
        Set objSection = mobjExamGeneral
        
        objSection.SaveSection
    End If
    
    Set oMultiList = New NotesPlusUtilities.CStandardList

    If oMultiList.Load_String(strSuppProcs) Then
        If oMultiList.ValueStringDisplay <> "" Then
            gobjDental.Misc_RTF = "\viewkind4\uc1\pard\ul\b\f0\fs20 " & "Graphic Charting Supporting Procedures:" & _
                    "\ulnone\b0\par " & oMultiList.ValueStringDisplay & "\par "
        Else
            gobjDental.Misc_RTF = Empty
        End If
    End If
FuncExit:
    
    Set objMod = Nothing
    Set objSection = Nothing
    Set oMultiList = Nothing
    Exit Sub
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "BigTooth.SaveSupportingProcedures", "BigTooth", vbCritical)
    GoTo FuncExit
End Sub

Private Sub oToothSVG_RemoveItemsByDate(eReasonForDeletion As CHCSII_Dental_DXTX.DeleteReasonEnum)
  
  If frmGraphicalCharting.Visible = True Then
    frmGraphicalCharting.RemoveItemsByDate eReasonForDeletion
  End If
End Sub
