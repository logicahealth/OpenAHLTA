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
Attribute VB_Name = "Setup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Implements ICHCSII_CoreAPI2
Implements ICHCSII_CoreAPI3

Private objUser      As CHCSII_ClientSecurity.ICHCSII_User
Private mfrmSetup    As frmSetup
Private mlMenuAdd    As Long
Private mlMenuEdit   As Long
Private mlMenuSave   As Long
Private mlMenuCancel As Long
Private mlMenuRefresh As Long
Private mlMenuActiveOnly As Long
Private mlMenuAllSpecialties As Long
Private mlMenuPrint  As Long

Public UserCanWrite  As Boolean   '-- true if user can write (set in openview)


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
            QueryView = True    '--this view is NOT optional--
        Case cwiPATIENT_SUMMARY_VIEW
            QueryView = False
        Case cwiPRINT                   '- do you want the print menu available
            QueryView = False
        Case cwiPRINTPREVIEW            '- do you want the print preview menu available
            QueryView = False
        Case cwiPROPERTY_VIEW           '- do you want the properties/options button available
            QueryView = False
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
    If viViewType = cwiREGULAR_VIEW Then
        '---set global flags---
        cwbFirstRegularResize = True
        gSetupViewRunning = True
        
        mlMenuAdd = gobjComm.AddActionMenuItem("&Add", "add.bmp", gSetupProgID, "Add specialty")
        mlMenuEdit = gobjComm.AddActionMenuItem("&Edit", "edit.bmp", gSetupProgID, "Edit currently selected specialty-location")
        'mlMenuSave = gobjComm.AddActionMenuItem("&Save", "save.bmp", gSetupProgID, "Save Changes")   ' SCR 80452
        mlMenuRefresh = gobjComm.AddActionMenuItem("&Refresh", "refresh.bmp", gSetupProgID, "Refresh Display")
        mlMenuPrint = gobjComm.AddActionMenuItem("&Print", "print.bmp", gSetupProgID, "Print Hospital Setup")
        
        mlMenuActiveOnly = gobjComm.AddViewMenuItem("Active Specialties Only", gSetupProgID, "view only active specialty/locations")
        mlMenuAllSpecialties = gobjComm.AddViewMenuItem("Include Inactive Specialties", gSetupProgID, "view both active and Inactive specialty/locations")
        
        'SF SCR 68132
        Call gobjComm.UpdateMenuItem(mlMenuAllSpecialties, "CHECK")
        '--- SCR 88642   Sherry Wang  2/28/2006
        gbActiveonly = False
        
        UserCanWrite = objUser.HasPrivilegeEx("Hospital_Setup", Priv_Write)
        If Not UserCanWrite Then
            gobjComm.UpdateMenuItem mlMenuAdd, "disable"
            gobjComm.UpdateMenuItem mlMenuEdit, "disable"
            gobjComm.UpdateMenuItem mlMenuSave, "disable"
        End If
        Set mfrmSetup = New frmSetup
        SetParent mfrmSetup.hwnd, vlParenthWnd
        
                
        rhMainForm = mfrmSetup.hwnd
    Else
        
        OpenView = False
        Exit Function
    End If
        
    
    OpenView = True
    Exit Function
    
ErrHandler:
    
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "PAD_Setup.OpenView", "PAD_Setup", vbCritical)
    
    
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
        '---close the "main view"---
        If gSetupViewRunning Then
            '---set global flags---
            gSetupViewRunning = False
            
            '---this indicates that the CW Core is closing the "view", NOT this OLE Server---
            cwbCloseViewMyself = False
            
            '---close ALL forms that are associated with all "views"---
            Unload mfrmSetup
        End If
    ElseIf viViewType = cwiREGULAR_VIEW Then
        '---close the "main view"---
        If gSetupViewRunning Then
            '---set global flags---
            gSetupViewRunning = False
            
            '---this indicates that the CW Core is closing the "view", NOT this OLE Server---
            cwbCloseViewMyself = False
            
            '---close ALL forms that are associated with all "views"---
            Unload mfrmSetup
            Set mfrmSetup = Nothing
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
    On Error GoTo ErrHandler
    
    '----------------------------------------------------------------------------------
    '  NOTE: You don't have to create menu items, these are optional, that is why
    '        the code is commented out.
    '----------------------------------------------------------------------------------
    '---perform the actions associated with the menu item that was selected---
    
    Select Case vlMenuId
        Case mlMenuAdd
            Call AddNewWard
        Case mlMenuEdit
            Call EditWard
        Case mlMenuRefresh
           Screen.MousePointer = vbHourglass
           InitializeCollections True
           mfrmSetup.Display
           Screen.MousePointer = vbDefault
       Case mlMenuCancel
        
        Case mlMenuSave
            Call SaveWardSetup
        Case mlMenuActiveOnly
            '--- SCR 88642  Sherry Wang  2/27/2006
            gbActiveonly = True
            Call mfrmSetup.DisplayAllWards(gbActiveonly)
            Call gobjComm.UpdateMenuItem(mlMenuAllSpecialties, "UNCHECK")
            Call gobjComm.UpdateMenuItem(mlMenuActiveOnly, "CHECK")
        Case mlMenuAllSpecialties
            '--- SCR 88642  Sherry Wang  2/27/2006
            gbActiveonly = False
            Call mfrmSetup.DisplayAllWards(gbActiveonly)
            Call gobjComm.UpdateMenuItem(mlMenuActiveOnly, "UNCHECK")
            Call gobjComm.UpdateMenuItem(mlMenuAllSpecialties, "CHECK")
        Case mlMenuPrint
            Call PrintGrid
        Case Else   '--do nothing--
    End Select
    

    Exit Sub

ErrHandler:
    
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "PAD_Setup.MenuItemSelected", "PAD_Setup", vbExclamation)
    
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
    If gSetupViewRunning Then
        '****************************************************************************
        '  TO DO:
        '    Add code here to update the data in all "running" views with this new
        '    patient.
        '****************************************************************************
    
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

    '---reset global "client area" parameters---
    SetupLeft = vlLeft
    SetupTop = vlTop
    SetupWidth = vlWidth
    SetupHeight = vlHeight
    
    '---perform different actions based upon the "view type"---
    If viViewType = cwiREGULAR_VIEW Then
        With mfrmSetup
            If Not cwbFirstRegularResize Then
                .WindowState = vbNormal
                .Visible = False
                .Move SetupLeft, SetupTop, SetupWidth, SetupHeight
                .Visible = True
            Else    '--first "regular" resize, need to perform "initailization" procedures--
                '---reset flag so this code will not be executed again---
                cwbFirstRegularResize = False
                .Visible = False
                .Move SetupLeft, SetupTop, SetupWidth, SetupHeight
                .Visible = True
                Set .objSetup = Me
                .Display
            End If
        End With
    End If
    
    Exit Sub
    
ErrHandler:
    
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "PAD_Setup.ResizeView", "PAD_Setup", vbExclamation)
    
    Exit Sub
End Sub

Private Sub Class_Initialize()
    '---create shared object used to show VB errors, etc.---
    Set gobjShared = New CWShared
    
    '---set the help file for this application---
    App.HelpFile = gobjShared.CWHelpFile
    If gColProviders Is Nothing Then
        Set gColProviders = New Collection
    End If
End Sub

Private Sub Class_Terminate()
    '---release reference to objects---
    
'SCR 107376, need a better solution, but have no time now.
'    Set gobjShared = Nothing
'    Set gobjComm = Nothing
'    Set gobjLogon = Nothing
'    Set gobjconfig = Nothing
'    Set gobjPatient = Nothing

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
        Case cwiFOREGROUND_APPLICATION
            'Refresh 64119
           Screen.MousePointer = vbHourglass
           InitializeCollections False
           mfrmSetup.Display
           Screen.MousePointer = vbDefault
            
        Case cwiOK_TO_CHANGE_PATIENT
            gobjComm.Message cwiOK_TO_CHANGE_PATIENT, "Y", vsSourceServerNCID, NCID_HOSPITAL_SETUP
        
        Case cwiOK_TO_CLOSE_VIEW
            gobjComm.Message cwiOK_TO_CLOSE_VIEW, "Y", vsSourceServerNCID, NCID_HOSPITAL_SETUP
        
        Case Else   '--do nothing--
    End Select

    Exit Sub
    
ErrHandler:
    
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "PAD_Setup.Message", "PAD_Setup", vbExclamation)
    
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
    gSetupProgID = vlMyProgId
    
    '---create references to passed in objects---
    Set gobjComm = robjComm
    Set gobjLogon = robjLogonEnvironment
    Set gobjconfig = robjConfig
    Set gobjPatient = robjPatient
    
    Call InitializeCollections
    
    '---set global flags---
    gSetupViewRunning = False
    Exit Sub

ErrHandler:
    
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "PAD_Setup.Initialize", "PAD_Setup", vbExclamation)
    
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
    
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "PAD_Setup.Properties", "PAD_Setup", vbExclamation)
    
    Exit Sub
End Sub

Private Sub ICHCSII_CoreAPI2_Initialize2(ByVal vlMyProgId As Long, _
                                               robjComm As Object, _
                                               robjLogonEnvironment As Object, _
                                               robjConfig As Object, _
                                               robjPatient As Object, _
                                               robjUser As Object)
    Set objUser = robjUser
    Call Initialize(vlMyProgId, robjComm, robjLogonEnvironment, robjConfig, robjPatient)
End Sub

Private Function ICHCSII_CoreAPI3_CloseView2(ByVal viViewType As Integer) As Boolean
  Call CloseView(viViewType)
  ICHCSII_CoreAPI3_CloseView2 = True
End Function


Friend Function SaveWardSetup() As Boolean
    Dim oRS As ADODB.Recordset
    Dim oDal As cDAL
    Set oDal = New cDAL
    On Error GoTo ErrHandler
    Dim oWard As cWard
    Dim oBed As cBed
    '<< scr 42898 control 'write' access
    If Not UserCanWrite Then Exit Function
    
    For Each oWard In gWards
        If oWard.objectstate And osNew Then
            '__ insert into ward

            Set oRS = oDal.GetPADData(edtp_Ward_facility, "0")
            oRS.AddNew
            Call oWard.SerializeToRecordset(oRS)
            oRS.Update
            oWard.objectstate = osClean

        ElseIf oWard.objectstate And osDirty Then

            Set oRS = oDal.GetPADData(edtp_Ward_One, oWard.WardID)
            If Not oRS.EOF Then
                Call oWard.SerializeToRecordset(oRS)
                oRS.Update
                oWard.objectstate = osClean
            End If

        End If
        '_ kill all the old beds and replace with the current ones
        Set oRS = oDal.SetPADData(edtp_Ward_Bed, oWard.WardID)
        For Each oBed In oWard.Beds
            oRS.AddNew
            Call oBed.SerializeToRecordset(oRS)
            oRS.Update
        Next oBed
    
    Next oWard
    SaveWardSetup = True
    Exit Function
ErrHandler:
    SaveWardSetup = False
End Function


Friend Function EditWard() As Boolean
    On Error GoTo ErrHandler
    '<< scr 42898 control 'write' access
    If Not UserCanWrite Then Exit Function
    
    If mfrmSetup.fGrid.row < 1 Then Exit Function
    If Not UserCanWrite Then Exit Function
    Dim frm As frmWardDetails
    Set frm = New frmWardDetails
    Set frm.CurrWard = gWards(CStr(mfrmSetup.fGrid.RowData(mfrmSetup.fGrid.row)))
    
'-- we have to hold the original version of the ward to see if a change was made
    Dim OrigWard As cWard
    Set OrigWard = CloneWard(frm.CurrWard)
    
    frm.CurrWard.objectstate = osDirty
    
    frm.Display         '--- 'showing modal is handled in this function
    
    If frm.UserCancelled Then Exit Function
    Screen.MousePointer = vbHourglass
    gobjComm.SetStatusText "Updating Specialty/Location Information..."
    gsDMIS_ID = frm.txtDMIS.Text
    Call gWards.SaveWard(frm.CurrWard)
    Call SetDMIS_ID
    
    Call SendTMIPBedStatusChangesOnly(OrigWard, frm.CurrWard)
    
    Call mfrmSetup.DisplayWard(frm.CurrWard, mfrmSetup.fGrid.row)
    Call SaveWardSetup
    Screen.MousePointer = vbDefault
    gobjComm.SetStatusText ""
    EditWard = True
    Exit Function
ErrHandler:
    
    EditWard = False
End Function

Friend Function AddNewWard() As Boolean
    On Error GoTo ErrHandler
    Dim frm As frmWardDetails
    
    '<< scr 42898 control 'write' access
    If Not UserCanWrite Then Exit Function
    
    Set frm = New frmWardDetails
    Set frm.CurrWard = New cWard
    frm.CurrWard.objectstate = osNew
    
    frm.Display
    
    ' Problems were reported on a number of SCR's that the grid was not refreshing correctly in a variety of situations. So always refresh the grid.
    MenuItemSelected mlMenuRefresh
    
    If frm.UserCancelled Then
        Exit Function
    End If
    
    If Not frm.CurrWard Is Nothing Then
        gWards.SaveWard frm.CurrWard
        gWards.AddWard frm.CurrWard
        If frm.DMIS_ID_Changed Then
                gsDMIS_ID = frm.txtDMIS.Text
                Call SetDMIS_ID
        End If
    End If
    Screen.MousePointer = vbHourglass
    Dim Spec As String
    Dim oKVP As KeyValuePair
    If frm.Locations_Changed Then
        Spec = vbNullString
        For Each oKVP In gcolLocations
            Spec = Spec & "#" & oKVP.Key & ";" & oKVP.Value & "|"
        Next oKVP
'<< SCR 48768 this was erroring out because it was looking at the wrong form and the wrong grid and had the wrong col identified
        mfrmSetup.fGrid.ColComboList(GRID_LOCATION) = Spec
    End If

    Call SendTMIPBedStatusWard(frm.CurrWard)
    Call SaveWardSetup
    '--- SCR 88393   Sherry Wang  2/21/2006
    '--- SCR 88642  Sherry Wang  2/27/2006
    Call mfrmSetup.DisplayAllWards(gbActiveonly)
    Screen.MousePointer = vbDefault
    Set frm = Nothing
    AddNewWard = True
    Exit Function
ErrHandler:
    

End Function

Public Function GetBedStatus(ByRef XMLStr As String, _
                    Optional ByRef ErrStr As String) As Boolean

On Error GoTo ErrHandler
    XMLStr = vbNullString
    '--build a collection of all beds
    Dim colBeds As New Collection
    Dim oWard As cWard
    Dim oBed As cBed
    
    If gWards Is Nothing Then
        Err.Raise 1201, "PAD.Hospital_setup", "No wards found."
    End If
    If gWards.Count = 0 Then
        Err.Raise 1201, "PAD.Hospital_setup", "No wards found."
    End If
   
    For Each oWard In gWards
        If oWard.Status <> "INACTIVE" Then
            If Not oWard.Beds Is Nothing Then
                For Each oBed In oWard.Beds
                    colBeds.Add oBed
                Next oBed
            End If
        End If
    Next oWard
    If colBeds.Count > 0 Then
        XMLStr = BuildBedStatusXML(colBeds)
    End If
    ErrStr = vbNullString
    GetBedStatus = True
    Exit Function
ErrHandler:
    GetBedStatus = False
    ErrStr = Err.Description
    XMLStr = vbNullString
    Exit Function
End Function

Private Sub PrintGrid()
    If mfrmSetup Is Nothing Then Exit Sub
    
    Call mfrmSetup.PrintGrid
End Sub

