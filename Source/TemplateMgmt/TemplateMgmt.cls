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
Attribute VB_Name = "TemplateMgmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements ICHCSII_CoreAPI2

Public Sub CreateNewFromEncounter(ByRef robjEncounter As encounter)
    On Error GoTo ErrHandler
    
    cwobjCWComm.Message cwiSTART_OLE_SERVER, NCID_TEMPLATE_MANAGEMENT, NCID_CW_APPLICATION, NCID_TEMPLATE_MANAGEMENT
    Screen.MousePointer = vbHourglass
        With frmMain
            If .mobjTemplateOps Is Nothing Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            Set .mobjTemplate = .mobjTemplateOps.newTemplate(EncounterTemplateType)
            .mobjTemplate.LoadFromEncounter robjEncounter
            .DisplayTemplateDetail
        End With
    Screen.MousePointer = vbDefault
        
    Exit Sub
ErrHandler:
    Screen.MousePointer = vbDefault
    Call gobjShared.ShowVBError(Err.Number, Err.description, "TemplateMgmt.CreateNewFromEncounter", "TemplateMgmt", vbCritical)
    'Resume Next
End Sub

Public Function GetEncounterAsTemplate(ByRef robjEncounter As encounter, Optional bCopyForward As Boolean = False) As Template
    
    Dim oTemplate As Template
    Dim oTemplateOps As CHCSII_TemplateClient.TemplateOps
    On Error GoTo ErrHandler
    
    Set oTemplateOps = New CHCSII_TemplateClient.TemplateOps
    oTemplateOps.Initialize cwlMyProgId, cwobjCWComm, cwobjLogonEnviron, cwobjConfig, cwobjPatient, gobjUser
    Set oTemplate = oTemplateOps.newTemplate(EncounterTemplateType)
    
    oTemplate.EnableCopyForward bCopyForward
    oTemplate.LoadFromEncounter robjEncounter
    
    Set GetEncounterAsTemplate = oTemplate
    
    GoTo CleanUpAndExit
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.description, "TemplateMgmt.GetEncounterAsTemplate", "TemplateMgmt", vbCritical)
CleanUpAndExit:
    Screen.MousePointer = vbDefault
    Set oTemplate = Nothing
    Set oTemplateOps = Nothing

End Function

Public Function DefaultEncounterTemplateData(Optional pProviderNCID As String) As String
        
    Dim oTemplateOps As CHCSII_TemplateClient.TemplateOps
    Dim oTemplate As Template
    Dim sBuf As String
    Dim oDefaultProvider As String
    
    On Error GoTo ErrHandler
    sBuf = ""
    DefaultEncounterTemplateData = ""
    
    If gobjShared.UseSQLServer Then
        'Enable Default Encounters in PG SCR-35673
        If Not gobjShared.IsAppMode(modeCHCSI_GUI) Then
            Exit Function
        End If
    End If
    
    If pProviderNCID = "" Then
    ElseIf cwobjLogonEnviron.UserNCID <> pProviderNCID Then
       oDefaultProvider = pProviderNCID
       cwobjConfig.RetrieveConfiguration cwobjLogonEnviron.EnterpriseNCID, cwobjLogonEnviron.FacilityNCID, "322", pProviderNCID 'cwobjLogonEnviron.UserNcid
    End If
    cwobjConfig.getValue NCID_TEMPLATE_MANAGEMENT, "Default", "DefaultTemplate", sBuf
    
    If Not oDefaultProvider = "" Then cwobjConfig.RetrieveConfiguration cwobjLogonEnviron.EnterpriseNCID, cwobjLogonEnviron.FacilityNCID, "322", cwobjLogonEnviron.UserNCID
    
    gDefaultEncounterTemplate = sBuf
    If gDefaultEncounterTemplate = "" Then Exit Function
        
    Set oTemplateOps = New CHCSII_TemplateClient.TemplateOps
    oTemplateOps.Initialize cwlMyProgId, cwobjCWComm, cwobjLogonEnviron, cwobjConfig, cwobjPatient, gobjUser
    Set oTemplate = oTemplateOps.GetTemplate(gDefaultEncounterTemplate)
    If Not oTemplate Is Nothing Then DefaultEncounterTemplateData = oTemplate.Export
    
    Set oTemplateOps = Nothing
    Set oTemplate = Nothing
    
    Exit Function
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.description, "TemplateMgmt.DefaultEncounterTemplateData", "Template Management", vbExclamation)
    
End Function

Public Sub SelectEncounterTemplate()
    
    On Error GoTo ErrHandler

    cwobjCWComm.Message cwiSTART_OLE_SERVER, NCID_TEMPLATE_MANAGEMENT, NCID_CW_APPLICATION, NCID_TEMPLATE_MANAGEMENT
    Screen.MousePointer = vbHourglass
    frmMain.tabMain.Tab = gnSuggestionsTab
    If Not frmMain.tabMain.Enabled Then Exit Sub
    If frmMain.tabMain.Visible Then frmMain.tabMain.SetFocus
    Screen.MousePointer = vbDefault

    Exit Sub
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.description, "TemplateMgmt.SelectEncounterTemplate", "Template Management", vbExclamation)

End Sub

Public Sub Initialize(ByVal vlMyProgId As Long, robjComm As Object, robjLogonEnvironment As Object, robjConfig As Object, robjPatient As Object)
    MsgBxARMd "Internal Error: The template management module requires initialization via ICHCSII_CoreAPI2."
    Exit Sub
End Sub

Public Sub ICHCSII_CoreAPI2_Initialize2(ByVal vlMyProgId As Long, ByRef robjComm As Object, ByRef robjLogonEnvironment As Object, ByRef robjConfig As Object, ByRef robjPatient As Object, ByRef robjUser As Object)
    
    On Error GoTo ErrHandler
    
    '---save the "ProgId" assigned to me by CW---
    cwlMyProgId = vlMyProgId
    
    '---create pointers to objects---
    Set cwobjCWComm = robjComm
    Set cwobjLogonEnviron = robjLogonEnvironment
    Set cwobjConfig = robjConfig
    Set cwobjPatient = robjPatient
    Set gobjUser = robjUser

    '---set global flags---
    cwbMainViewRunning = False
    gbSOWriteAccess = robjUser.HasPrivilegeEx(Priv_S_O, Priv_Write)
    gbAnPWriteAccess = robjUser.HasPrivilegeEx(Priv_Encounter_Coding, Priv_Write)
    
    gbEnterpriseWriteAccess = robjUser.HasPrivilegeEx(Priv_Templates_Enterprise_Template_Guidelines, Priv_Write) Or robjUser.HasPrivilegeEx(Priv_Order_Sets_Enterprise_Order_Set, Priv_Write) 'SCR-44268 RBELL 1/9/4 Set security variable
    gbMTFWriteAccess = robjUser.HasPrivilegeEx(Priv_Templates_MTF_Template_Guidelines, Priv_Write) Or robjUser.HasPrivilegeEx(Priv_Order_Sets_MTF_Order_Set, Priv_Write) 'SCR-44268 RBELL 1/9/4 Set security variable
    gbClinicWriteAccess = robjUser.HasPrivilegeEx(Priv_Templates_Clinic_Templates, Priv_Write) Or robjUser.HasPrivilegeEx(Priv_Order_Sets_Clinic_Order_Set, Priv_Write) 'SCR-44268 RBELL 1/9/4 Set security variable
    gbPersonalWriteAccess = robjUser.HasPrivilegeEx(Priv_Templates_Personal_Template, Priv_Write) Or robjUser.HasPrivilegeEx(Priv_Order_Sets_Personal_Order_Sets, Priv_Write) 'SCR-44268 RBELL 1/9/4 Set security variable
    
    gbEnterpriseReadAccess = robjUser.HasPrivilegeEx(Priv_Templates_Enterprise_Template_Guidelines, Priv_Read) Or robjUser.HasPrivilegeEx(Priv_Order_Sets_Enterprise_Order_Set, Priv_Read)
    gbMTFReadAccess = robjUser.HasPrivilegeEx(Priv_Templates_MTF_Template_Guidelines, Priv_Read) Or robjUser.HasPrivilegeEx(Priv_Order_Sets_MTF_Order_Set, Priv_Read)
    gbClinicReadAccess = robjUser.HasPrivilegeEx(Priv_Templates_Clinic_Templates, Priv_Read) Or robjUser.HasPrivilegeEx(Priv_Order_Sets_Clinic_Order_Set, Priv_Read)
    gbPersonalReadAccess = robjUser.HasPrivilegeEx(Priv_Templates_Personal_Template, Priv_Read) Or robjUser.HasPrivilegeEx(Priv_Order_Sets_Personal_Order_Sets, Priv_Read)
    
    'SCR-48122 RBELL 1/15/04 Changed to Priv_Write only for security problem
    'SCR-48122 RBELL 1/21/04 Added Priv_Read as a seperate security key
    gbAllClinicReadAccess = robjUser.HasPrivilegeEx(gsClinicAccess, Priv_Read)
    gbAllClinicWriteAccess = robjUser.HasPrivilegeEx(gsClinicAccess, Priv_Write)
    gbAllClinicDeleteAccess = robjUser.HasPrivilegeEx(gsClinicAccess, Priv_Delete)
    
    gbAllMTFReadAccess = robjUser.HasPrivilegeEx(gsMTFAccess, Priv_Read)
    gbAllMTFWriteAccess = robjUser.HasPrivilegeEx(gsMTFAccess, Priv_Write)
    gbAllMTFDeleteAccess = robjUser.HasPrivilegeEx(gsMTFAccess, Priv_Delete)
    
    Set gobjContext = New ContextCtrl
    
    '---get user's default configuration values---
    Call gGetDefaultConfigValues
    
    'Call CallStack("ICHCSII_CoreAPI2_Initialize2", False)
    Exit Sub

ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.description, "TemplateMgmt.Initialize", "Template Management", vbExclamation)
    'Call CallStack("ICHCSII_CoreAPI2_Initialize2", False)
End Sub

Public Sub MenuItemSelected(ByVal vlMenuId As Long)
    
    Dim oForm As Form
    
       Set oForm = frmMain
    
    On Error GoTo ErrHandler
    Select Case vlMenuId
        Case giFindMenuId
            oForm.DoSearch
        Case giDetailsMenuId
            oForm.ViewDetails
        Case giNewMenuId
            oForm.menuPopupNew_Click
        Case giMergeMenuId
            oForm.menuPopupMerge_Click
        Case giSaveAsMenuId
            oForm.SaveAs
        Case giAddToFavoritesMenuId
            oForm.menuPopupAddFavorite_Click
        Case giRemoveFromFavoritesMenuId
            oForm.RemoveSelectedFavorite
        Case giImportMenuId
            oForm.menuPopupImport_Click
        Case giExportMenuId
            oForm.menuPopupExport_Click
        Case giDeleteId
            oForm.menuPopupDelete_Click
        Case giSaveMenuId
            oForm.SaveTemplate
        Case giCancelMenuId
            oForm.cmdCancel_Click
        Case giNextId
            goARM.StartFarTx gsARMTx_NavigateBtnTmplateMgmt2SO  'SCR 27123 MSolano 8/27/02
            oForm.AcceptSelections True, False
            goARM.StopFarTx gsARMTx_NavigateBtnTmplateMgmt2SO  'SCR 27123 MSolano 8/27/02
        Case giAnpId
            goARM.StartFarTx gsARMTx_NavigateBtnTmplateMgmt2AnP 'SCR 27123 MSolano 8/27/02
            oForm.AcceptSelections False, True
            goARM.StopFarTx gsARMTx_NavigateBtnTmplateMgmt2AnP  'SCR 27123 MSolano 8/27/02
        Case giRefreshId
            oForm.CustomFolders = New TemplateFolders
            oForm.RefreshSuggestions
        Case giColumnOrderId
            oForm.SetColumnDefaults
        Case giResetColumnOrderId
            oForm.ReSetColumnDefaults
        Case glMenuMoveToFolder
            oForm.MoveToFolder
        Case glMenuCopy
            oForm.CopyToFolder 'SCR-44247 5/6/2004 RBELL Support new menu item
        Case giCancel
            oForm.ViewTemplateList
        Case glMenuDxPrompt
            oForm.Menu_DXPrompt
        Case glMenuSearch
            oForm.Menu_FindTerm
        Case glMenuBrowse
            oForm.Menu_BrowseFromHere
        Case glMenuAdd
            oForm.Menu_AddInTemplate
        Case giLoad
            oForm.AcceptSelections
        Case Else
            ' do nothing for now
    End Select
    Exit Sub
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.description, "TemplageMgmt.MenuItemSelected", "TemplateManagement", vbExclamation)
End Sub

'*******************************************************************************************
'  Sub Name: Properties
'
'  Description: This method allows a user the ability to modify the "properties" of this OLE
'               Server (i.e. configuration values).
'
'  Parameters:
'    robjLogonEnviron - Contains the logon environment object.
'    robjConfig - Contains the configuration object.
'    vsOwnerNCID - Contains the "owner" for which to store the new/updated OLE Server
'                  property information.
'    vbCalledByDesigner - TRUE if this method is being called by the CW Designer Tool,
'                         FALSE otherwise.
'*******************************************************************************************
Public Sub Properties(robjLogonEnviron As Object, robjConfig As Object, ByVal vsOwnerNCID As String, ByVal vbCalledByDesigner As Boolean)
    On Error GoTo ErrHandler
    
    '--------------------------------------------------------------------------------------
    '  If this method is being called by the CW Designer, you need to perform actions that
    '    the OLE Server would preform in the "Initialize" method and other methods (i.e.
    '    create objects, initialize global variables, etc.).
    '--------------------------------------------------------------------------------------
    If vbCalledByDesigner Then
        '---create pointers to objects---
        Set cwobjLogonEnviron = robjLogonEnviron
        Set cwobjConfig = robjConfig
        
        '---get user's default configuration values---
        Call gGetDefaultConfigValues
        
    End If  '--if called by the CW Designer--
    
    '---allow the user to modify the "properties" of this application---
'    frmmain.Properties vsOwnerNCID, vbCalledByDesigner
    
    '---if you were called by the CW Designer, clean up the objects that were used---
    If vbCalledByDesigner Then
        '---release reference to objects---
        Set cwobjLogonEnviron = Nothing
        Set cwobjConfig = Nothing
    End If

    Exit Sub
    
ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.description, "TemplateMgmt.Properties", "Template Management", vbExclamation)
    
    Exit Sub
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

    On Error GoTo ErrHandler
    Select Case viViewType
        Case cwiREGULAR_VIEW
            QueryView = True    '--this view is NOT optional--
        Case Else   '--unknown "view type", not supported--
            QueryView = False
    End Select

    Exit Function
    
ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.description, "TemplateMgmt.QueryView", "Template Management", vbExclamation)
    
End Function

'**************************************************************************************************************
'  Sub Name: PatientSelected
'
'  Description: This method gets called when a new patient has been selected.
'
'  Parameters:
'    none
'**************************************************************************************************************
Public Sub PatientSelected()
    ' We don't care
End Sub

'**************************************************************************************************************
'  Sub Name: Message
'
'  Description: This method is used by CW to send miscellaneous messages to this OLE Server.
'
'  Parameters:
'    viMessageType - Contains the "type" of message to process.
'    vsMessageText - Contains any "data" assoicated with this message.
'    vsSourceServerNCID - Contains the NCID of the OLE Server that "sent" the message.
'**************************************************************************************************************
Public Sub Message(ByVal viMessageType As Integer, ByVal vsMessageText As String, ByVal vsSourceServerNCID As String)
    
    Dim objAbout As CWAbout
    Dim objClinicOps As ClinicOpsEx
    On Error GoTo ErrHandler
    
    If (viMessageType = cwiREFRESH_DATA Or viMessageType = cwiREFRESH) Then
        If vsSourceServerNCID = NCID_PROBLEMS_APPLICATION Then Exit Sub
    End If
    
    Select Case viMessageType
        Case cwiREFRESH
            '---"refresh" or repaint all running "views"---
            If cwbMainViewRunning Then
                frmMain.Visible = False
                frmMain.Visible = True
            End If
        Case cwiPRINT_REGULAR_VIEW
            '---print the selected Template data---
'            frmmain.PrintTemplateMgmtigns
        Case cwiPRINT_PREVIEW
            '---set flag---
            cwbPrintPreview = True
            
            '---print the selected Template Management data---
'            frmmain.PrintTemplateMgmtigns
            
            '---set flag---
            cwbPrintPreview = False
        Case cwiITEM_SPECIFIC_HELP
            '---invoke the "what's this mode" for the Template Management form---
'            frmmain.WhatsThisMode
        Case cwiPROPERTIES
            '---allow the user to change the "properties" of this application--
'            frmmain.Properties cwobjLogonEnviron.UserNCID, False
        Case cwiABOUT
            Set objAbout = New CWAbout
            
            '---inform CW that I am showing a modal form---
            cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_TEMPLATE_MANAGEMENT
            
            objAbout.Display App, frmMain
            
            '---inform CW that I am done showing a modal form---
            cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_TEMPLATE_MANAGEMENT
            
            Set objAbout = Nothing
        Case cwiREFRESH_DATA
            '---"refresh" or repaint all running "views"---
            If cwbMainViewRunning Then
                    frmMain.RefreshCurrentTab
            End If
        Case cwiGENERIC
                'SCR-49320 RBELL 2/19/2004 Refresh clinic then form based on updateclinic message
                Select Case UCase(vsMessageText)
                    Case "UPDATECLINICS"
                    If cwbMainViewRunning Then
                       Set objClinicOps = New ClinicOpsEx
                            Set frmMain.mcolUserClinics = objClinicOps.GetUserCLinics(cwobjLogonEnviron.FacilityNCID, cwobjLogonEnviron.UserNCID)
                            frmMain.RefreshCurrentTab
                    End If
                End Select
                Set objClinicOps = Nothing
        Case Else
'            Beep
'            MsgBxARMd "Template Management does not support this action.", vbInformation, "Template Management"
    End Select

    Exit Sub
    
ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.description, "TemplateMgmt.Message", "Template Management", vbExclamation)

End Sub

'*******************************************************************************************
'  Sub Name: CloseView
'
'  Description: This method is called when CW closes a "view" provided by this OLE server.
'
'  Parameters:
'    viViewType - Contains the "type" of view that CW is closing.
'*******************************************************************************************
Public Sub CloseView(ByVal viViewType As Integer)
    On Error GoTo ErrHandler
    
    If viViewType = cwiALL_VIEWS Then
        '---this indicates that CW is closing the "view", NOT the OLE Server---
        cwbCloseViewMyself = False
        
        '---unload all forms---
        If cwbMainViewRunning Then
                Unload frmMain
        End If
        
        '---set global flags---
        cwbMainViewRunning = False
    ElseIf viViewType = cwiREGULAR_VIEW Then
        '---this indicates that CW is closing the "view", NOT the OLE Server---
        cwbCloseViewMyself = False
    
        If cwbMainViewRunning Then
                Unload frmMain
        End If
    
        '---set global flag---
        cwbMainViewRunning = False
    End If

    Exit Sub
    
ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.description, "TemplateMgmt.CloseView", "Template Management", vbExclamation)

End Sub

'*******************************************************************************************
'  Sub Name: ResizeView
'
'  Description: This method is called when a change has occurred in the "client area"
'               used by this OLE server and it needs to resize a particular "view".
'
'  Parameters:
'    viViewType - Contains the "type" of view being requested to resize by CW.
'    vlLeft - Contains the "left" coordinate for the OLE server.
'    vlTop - Contains the "top" coordinate for the OLE server.
'    vlWidth - Contains the "width" for the OLE server.
'    vlHeight - Contains the "height" for the OLE server.
'*******************************************************************************************
Public Sub ResizeView(ByVal viViewType As Integer, ByVal vlLeft As Long, ByVal vlTop As Long, ByVal vlWidth As Long, ByVal vlHeight As Long)
    On Error GoTo ErrHandler

    '---reset global "client area" parameters---
    cwlLeft = vlLeft
    cwlTop = vlTop
    cwlWidth = vlWidth
    cwlHeight = vlHeight
    
    '--- SCR 26340 MSolano 8/22/02
    'Call CallStack("ResizeView", True)
    Call ARM_TemplateMgmtEntryTx
    
    '---perform different actions based upon the "view type"---
    If viViewType = cwiREGULAR_VIEW Then
        If Not cwbFirstRegularResize Then
            '---------------------------------------------------------------------------------
            '  Reset the "WindowState" to "Normal", so the form can be resized.  This is
            '    needed because a window CANNOT be moved or resized if it is "minimized" or
            '    "maximized".
            '
            '  NOTE: This generates a "resize" event that is ignored.
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
        
            '---display the "main" form---
                frmMain.Display
        End If
    Else
        Screen.MousePointer = vbDefault
        Beep
        MsgBxARMd "Unsupported resize view type: " & CStr(viViewType) & " in TemplateMgmt.ResizeView.", vbExclamation, "Template Management"
    End If
    
    'Call CallStack("ResizeView", False)
    Exit Sub
    
ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.description, "TemplateMgmt.ResizeView", "Template Management", vbExclamation)
    
    'Call CallStack("ResizeView", False)
    Exit Sub
End Sub



'*******************************************************************************************
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
'*******************************************************************************************
Public Function OpenView(ByVal viViewType As Integer, ByVal vlParenthWnd As Long, rhMainForm As Long) As Boolean
    On Error GoTo ErrHandler
    
    'Call CallStack("OpenView", True)
    
    '---perform actions based on the "view type"---
    If viViewType = cwiREGULAR_VIEW Then
        '---set global flags---
        cwbFirstRegularResize = True
        cwbMainViewRunning = True
    
        '----------------------------------------------------------------------
        '  Reset the "parent window" for this OLE server's "main" regular form.
        '
        '  NOTE: This invokes the form's "Load" event, but the form is NOT
        '        shown.  Make sure that you DO NOT perform a "Me.Show" in the
        '        form's "Load" event.
        '----------------------------------------------------------------------
            SetParent frmMain.hwnd, vlParenthWnd
            rhMainForm = frmMain.hwnd
    Else
        Screen.MousePointer = vbDefault
        Beep
        MsgBxARMd "Unsupported view type: " & CStr(viViewType) & " in TemplateMgmt.OpenView..", vbCritical, "Template Management"
        
        '---set return value---
        OpenView = False
        Screen.MousePointer = vbDefault
        Exit Function
    End If
        
    '---set return value---
    OpenView = True
    Exit Function

ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.description, "TemplateMgmt.OpenView", "Template Management", vbCritical)
    
    '---set return value---
    OpenView = False
    Screen.MousePointer = vbDefault

End Function


Private Sub Class_Initialize()
    
    On Error GoTo ErrHandler
    Set gobjShared = New CWShared
    
    '--- SCR 26340 Msolano 8/15/2002
    Call ARM_InitAllTransactionsForTemplateMgmt

    '------------------------------------
    '  This statement is needed to debug
    '    this OLE Server.
    '------------------------------------
    #If ciDEBUG = 1 Then
        If cwbDebugLogon = False Then
            'Call CallStack("Class_Initialize", False)
            Exit Sub
        End If
    #End If
    
    '---set the help file for this application---
    App.HelpFile = gobjShared.CWHelpFile
    
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.description
    
End Sub

Private Sub Class_Terminate()
    
    On Error Resume Next
    '---release reference to objects---
    Set gobjShared = Nothing
    Set gobjContext = Nothing
    Set cwobjCWComm = Nothing
    Set cwobjLogonEnviron = Nothing
    Set cwobjConfig = Nothing
    Set cwobjPatient = Nothing
    
    
    '--- SCR 26340 Msolano 8/15/2002
    'Call ARM_TemplateMgmtEntryTx(False)
    Set goARM = Nothing  'SCR 26340 Msolano 8/15/2002
    

    '----------------------------------------
    '  If debugging, make sure to "log off".
    '----------------------------------------
    #If ciDEBUG = 1 Then
        cwbDebugLogOff
    #End If

End Sub



