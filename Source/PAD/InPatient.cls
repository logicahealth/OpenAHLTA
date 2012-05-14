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
Attribute VB_Name = "InPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Implements ICHCSII_CoreAPI2
Implements ICHCSII_CoreAPI3
Implements I_RSTransfer

Private objUser         As CHCSII_ClientSecurity.ICHCSII_User
Private mfrmInpatient   As frmInpatient
Private msDischargeComments As String
Private moSelectedPatients As cInPatients        '-- the collection of inpatient objects for the frmInpatient display
                                                '   based on the curroptions object
Private mlMenuAdd       As Long
Private mlMenuEdit      As Long
Private mlMenuDischarge As Long
Private mlMenuTransfer  As Long
Private mlMenuRefresh   As Long
Private mlMenuPrint     As Long
Private mlMenuAddNote   As Long
Private mlMenuAdminClose     As Long
'--- SCR 88340  Sherry Wang  2/21/2006
Private mbFirstTimeLoading   As Boolean

Public CurrOptions      As cOptions

Public UserCanWrite     As Boolean          '-- true if user has write privielges
Public Enum ObjectState_
    osClean = 0
    osNew = 1
    osDirty = 2
End Enum
Public I_change_Patient_Now As Boolean

Private moNoteParentForm As frmNewAdmission '--- SCR 68731  Sherry Wang  1/27/2006


Public Sub DisplayDocuments()
    'SCR 66901
    
    Dim oDocuments As Object
    
    On Error GoTo ErrHnd
    
        Set oDocuments = gobjComm.InitializeOLEServer(NCID_INPATIENT_DOCUMENTS_APPLICATION)
        If Not oDocuments Is Nothing Then
            gobjComm.Message cwiSTART_OLE_SERVER, NCID_INPATIENT_DOCUMENTS_APPLICATION, NCID_CW_APPLICATION, NCID_INPATIENT_APPLICATION
            If Not oDocuments Is Nothing Then
                Call oDocuments.SetInpatientKey(SelectedPatients.Inpatient.ID)
            End If
        End If
        'need to reset in case user tries to open a different encounter via folders
        Set oDocuments = Nothing
    
Exit Sub

ErrHnd:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "PAD_Inpatient.MenuItemSelected", "PAD_Inpatient", vbExclamation)

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
            QueryView = True    '--this view is NOT optional--
        Case cwiPATIENT_SUMMARY_VIEW
            QueryView = False
        Case cwiPRINT                   '- do you want the print menu available
            QueryView = False
        Case cwiPRINTPREVIEW            '- do you want the print preview menu available
            QueryView = False
        Case cwiPROPERTY_VIEW           '- do you want the properties/options button available
            QueryView = True
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
        gInPatientViewRunning = True
        
        
        Call PatientMenus(Val(gobjPatient.UnitNumber) <> 0)
        
        Set mfrmInpatient = New frmInpatient
        SetParent mfrmInpatient.hwnd, vlParenthWnd
        
        Set mfrmInpatient.mobjParent = Me
        
        If SelectedPatients Is Nothing Then
            Set SelectedPatients = New cInPatients
            SelectedPatients.LoadInPatients CurrOptions
        Else
            'SCR 105460
            SelectedPatients.LoadInPatients CurrOptions
        End If
    
        UserCanWrite = objUser.HasPrivilegeEx("Inpatient", Priv_Write)
        rhMainForm = mfrmInpatient.hwnd
    Else
        
        OpenView = False
        
        Exit Function
    End If
    OpenView = True
    Exit Function
    
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "PAD_Inpatient.OpenView", "PAD_Inpatient", vbCritical)
    OpenView = False
    
    Exit Function
    Resume
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
        If gInPatientViewRunning Then
            '---set global flags---
            gInPatientViewRunning = False
            
            '---this indicates that the CW Core is closing the "view", NOT this OLE Server---
            cwbCloseViewMyself = False
            
            '---close ALL forms that are associated with all "views"---
            Unload mfrmInpatient
        End If
    ElseIf viViewType = cwiREGULAR_VIEW Then
        '---close the "main view"---
        If gInPatientViewRunning Then
            '---set global flags---
            gInPatientViewRunning = False
            
            '---this indicates that the CW Core is closing the "view", NOT this OLE Server---
            cwbCloseViewMyself = False
            
            '---close ALL forms that are associated with all "views"---
            Unload mfrmInpatient
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
            Call AddInPatientRecord
            
        Case mlMenuEdit
            Call EditPatientRecord
            
        Case mlMenuTransfer
            Call TransferPatient
        
        Case mlMenuDischarge
            Call PatientDischarge

        Case mlMenuRefresh
            RefreshFromDB
            
        Case mlMenuAddNote
            Call AddNote(SelectedPatients.Inpatient)
            
        Case mlMenuAdminClose
            Call AdminDischarge
            
        Case mlMenuPrint
            Call PrintInpatientList
    End Select

    Exit Sub

ErrHandler:
    
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "PAD_Inpatient.MenuItemSelected", "PAD_Inpatient", vbExclamation)
    
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
    If gInPatientViewRunning Then
        '****************************************************************************
        '  TO DO:
        '    Add code here to update the data in all "running" views with this new
        '    patient.
        '****************************************************************************
    
    End If
End Sub

Private Sub RefreshFromDB()

'SCR 76739, 76740
'Refresh from the database.

Dim lMouse As Long
Static lLastTimeRefreshed As Long
    
    'If the screen was just refreshed, then don't refresh it again because it's costly.
    Debug.Print Timer - lLastTimeRefreshed
    If Timer - lLastTimeRefreshed < 1 Then
        Exit Sub
    End If
    
    lMouse = Screen.MousePointer
    
    Screen.MousePointer = vbHourglass
    
    InitializeCollections True
    SelectedPatients.LoadInPatients CurrOptions
    mfrmInpatient.Display
    
    lLastTimeRefreshed = Timer
    
    Screen.MousePointer = lMouse
    
Exit Sub
    
ErrHnd:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "PAD_Inpatient.RefreshFromDB", "PAD_Inpatient", vbInformation)
    
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
    IPatLeft = vlLeft
    IPatTop = vlTop
    IPatWidth = vlWidth
    IPatHeight = vlHeight
    
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
            mfrmInpatient.WindowState = vbNormal
            
            mfrmInpatient.Visible = False
            mfrmInpatient.Move IPatLeft, IPatTop, IPatWidth, IPatHeight
            mfrmInpatient.Visible = True
        Else    '--first "regular" resize, need to perform "initailization" procedures--
            '---reset flag so this code will not be executed again---
            cwbFirstRegularResize = False
            '--- SCR 88340  Sherry Wang  2/21/2006
            mbFirstTimeLoading = True
        
            '----------------------------------------------------------------------------
            '  NOTE: Here is where you "Show" the "Regular View" form for the FIRST time.
            '----------------------------------------------------------------------------
            mfrmInpatient.Visible = False
            mfrmInpatient.Move IPatLeft, IPatTop, IPatWidth, IPatHeight
            mfrmInpatient.Visible = True

            Set mfrmInpatient.mobjParent = Me
            Call AddMenus
            
        End If
        
        '--- SCR 56597  Automatically refresh every time the module regains focus
        mfrmInpatient.sEventSource = vbNullString
        mfrmInpatient.Display
    End If
    
    Exit Sub
    
ErrHandler:
    
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "PAD_Inpatient.ResizeView", "PAD_Inpatient", vbExclamation)
    
    Exit Sub
End Sub

Public Property Get SelectedPatients() As CHCSII_PAD.cInPatients
    
    Set SelectedPatients = moSelectedPatients
    
End Property

Public Property Set SelectedPatients(ByVal vData As CHCSII_PAD.cInPatients)
    
    Set moSelectedPatients = vData
    
End Property
Private Sub Class_Initialize()
    '---create shared object used to show VB errors, etc.---
    Set gobjShared = New CWShared
    
    If gobjShared.UseMySQLDB Then
        SetDBType edb_MySQL
    ElseIf gobjShared.UseSQLServer Then
       Call SetDBType(edb_sqlserver)
    Else
        Call SetDBType(edb_Oracle)
    End If
    '---set the help file for this application---
    App.HelpFile = gobjShared.CWHelpFile
    
    '__ get a collection of prividers
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
    
        Case cwiOK_TO_CHANGE_PATIENT
            If Not mfrmInpatient Is Nothing Then
                If Not I_change_Patient_Now Then '__(someone else is doing it)
                    '__ clear the current patient selected line
                    Set SelectedPatients.Inpatient = Nothing
                    mfrmInpatient.fGrid.row = -1
                End If
            End If
            gobjComm.Message cwiOK_TO_CHANGE_PATIENT, "Y", vsSourceServerNCID, NCID_INPATIENT_APPLICATION
        
        Case cwiOK_TO_CLOSE_VIEW
            gobjComm.Message cwiOK_TO_CLOSE_VIEW, "Y", vsSourceServerNCID, NCID_INPATIENT_APPLICATION
        Case cwiPROPERTIES
            If gInPatientViewRunning Then
                mfrmInpatient.cmdSelection.Value = True
            End If
        Case cwiCLEAR_PATIENT
            If gInPatientViewRunning Then
                Set SelectedPatients.Inpatient = New cInpatient
            End If
        Case cwiFOREGROUND_APPLICATION
            '-- put the column order back the way it should be
             If vsMessageText = NCID_INPATIENT_APPLICATION And gInPatientViewRunning Then        '-- we've just come back to the foreground
                'SCR 76740 Refresh the display when gaining focus.
                '--- SCR 88340  Sherry Wang  2/21/2006
                If mbFirstTimeLoading = False Then
                    RefreshFromDB
                Else
                    mbFirstTimeLoading = False
                End If
             End If
        
        Case Else   '--do nothing--
    End Select

    Exit Sub
    
ErrHandler:
    
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "PAD_Inpatient.Message", "PAD_Inpatient", vbExclamation)
    
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
    
    gInpatientProgID = vlMyProgId
    
    Set gobjComm = robjComm
    Set gobjLogon = robjLogonEnvironment
    Set gobjconfig = robjConfig
    Set gobjPatient = robjPatient
    
    gInPatientViewRunning = False
    Set CurrOptions = New cOptions
    Set CurrOptions.Config = gobjconfig
    Call CurrOptions.GetUserOptions
    
    Call InitializeCollections
    
    If SelectedPatients Is Nothing Then
        Set SelectedPatients = New cInPatients
        SelectedPatients.LoadInPatients CurrOptions
    End If

    UserCanWrite = objUser.HasPrivilegeEx("Inpatient", Priv_Write)
    
    Exit Sub

ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "PAD_Inpatient.Initialize", "PAD_Inpatient", vbExclamation)
    
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
    
    '<< we don't support property pages
    
    Exit Sub
    
ErrHandler:
    
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "PAD_Inpatient.Properties", "PAD_Inpatient", vbExclamation)
    
    Exit Sub
End Sub

Private Function I_RSTransfer_Transfer(oDMBE As IDMBE.I_DMBE) As Variant

    I_RSTransfer_Transfer = modTransfer.I_RSTransfer_Transfer(oDMBE)

End Function

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


Friend Function AddInPatientRecord() As Boolean
    On Error GoTo ErrHandler
    Dim oInPat As cInpatient
    Dim fNewAdmit As frmNewAdmission
    '<< scr 42898 control 'write' access
    If Not UserCanWrite Then Exit Function
    If gWards.Count = 0 Then
        MsgBxARMd "you must create at least one specialty/location before admitting a patient."
        Exit Function
    End If
'<< SCR 42923 improperly clearing patient object
    'gobjPatient.ClearPatient
    
'    'SF SCR 63619
'    If SelectedPatients.Inpatient.DischargeDate = Empty Then
'        MsgBxARMd "The selected patient is already admitted."
'        Exit Function
'    End If
    
    Set fNewAdmit = New frmNewAdmission
    
        gobjComm.Message cwiSHOWING_MODAL_FORM, vbNullString, NCID_CW_APPLICATION, NCID_INPATIENT_APPLICATION, False
    
    Set fNewAdmit.Parent = Me
    Call fNewAdmit.Display(NewAdmission)
    
        gobjComm.Message cwiDONE_SHOWING_MODAL_FORM, vbNullString, NCID_CW_APPLICATION, NCID_INPATIENT_APPLICATION, False

    If fNewAdmit.Cancelled Then Exit Function
    If gobjPatient Is Nothing Then Exit Function
    
        
'__ the current patient is now the patient to be added
'    copy patient to cinpatient

    Set oInPat = fNewAdmit.Inpatient
    If oInPat.objectstate = osNew Then
         gInPatients.AddPatient oInPat
    End If
        
    
    '-- add appt id
    Call CreateAppointment(oInPat)
    
    Call gInPatients.Save(oInPat)
    
    '-- ADD To selected patients and the display if it fits the current options
    If Not mfrmInpatient Is Nothing Then
        SelectedPatients.LoadInPatients CurrOptions
        mfrmInpatient.Display
    End If
    
    If fNewAdmit.UserClickedAddNote Then
        Call AddNote(oInPat)
    End If
    Set fNewAdmit = Nothing
    
    'Send the Inpatient Status message
    'SF 62830
    'This will eventually call I_RSTransfer_Transfer in this module
    Dim dm As Object
    Set dm = gobjComm.InitializeOLEServer("1999988")  'NCID_DATA_MANAGER
    Call dm.SendInpatientStatus(oInPat.ID)
    Set dm = Nothing
    
    AddInPatientRecord = True
    Exit Function
ErrHandler:
    MsgBxARMd "Application Error: " & Err.Description & " Occurred in CHCSII_OPS.Inpatient.AddInPatientRecord."
    AddInPatientRecord = False
End Function


Public Function EditPatientRecord(Optional ByVal bFromDocuments As Boolean = False) As Boolean
    '__ for now If mfrmInpatient.fGrid.row < 1 Then Exit Function
    '__ get the patient info into the current patient object
On Error GoTo ErrHandler
    Dim gfrmNewAdmission As New frmNewAdmission
    '<< scr 42898 control 'write' access
    If Not UserCanWrite Then Exit Function
    
    If SelectedPatients.Inpatient.AdmissionDate = Empty Then
        MsgBxARMd "Please select an inpatient record to edit.", vbInformation
        Exit Function
    End If
    
    With gfrmNewAdmission
        'SCR 68888 Hide the change patient button and make sure the patient is switched.
        .cmdChangePatient.Visible = False
        
        '--- SCR 68731  Sherry Wang  1/27/2006
        .EditPatientRecord = True
        
        Set .Parent = Me
        Set .Inpatient = SelectedPatients.Inpatient
        .lblPatIDLine = gobjComm.GetPatientIDLine
        .Move (Screen.Width - .Width) / 2, (Screen.Height - .Height) / 3
        
        If CDbl(.Inpatient.DischargeDate) = 0 Then
            .Display Admitted
        Else
            .Display Discharged
        End If
        
        If .Inpatient.objectstate = osDirty Then
            
            If mfrmInpatient.Visible = True Then
                'Only refresh if the form is visible RQT: 61008
                Call mfrmInpatient.RefreshPatientDisplay(SelectedPatients.Inpatient)
            End If
            
            Call SelectedPatients.Save(SelectedPatients.Inpatient)
            .Inpatient.objectstate = osClean
        End If
        
        
    End With
    
    '--- SCR 85288  Sherry Wang  2/13/2006
    If bFromDocuments = False Then
        'SCR 76739 SF Refresh after showing the admission details window.
        RefreshFromDB
    End If

    EditPatientRecord = True
    
    Exit Function
ErrHandler:
    
    
End Function

Friend Function TransferPatient() As Boolean
   '__ for now  If mfrmInpatient.fGrid.row < 1 Then Exit Function
    '__ get the patient info into the current patient object
On Error GoTo ErrHandler
    Dim frm As New frmpatientTrans
    Dim oBed As cBed
    Dim colBeds As New Collection
    
    '<< scr 42898 control 'write' access
    If Not UserCanWrite Then Exit Function
    
    If SelectedPatients.Inpatient Is Nothing Then Exit Function
    If SelectedPatients.Inpatient.Unit_Number = vbNullString Then Exit Function
    
    Set oBed = SelectedPatients.Inpatient.Bed
    With frm
        Set .oInpatient = SelectedPatients.Inpatient
        .lblPatIDLine.Caption = gobjComm.GetPatientIDLine
        .dtDate = Now
        .dtTime = Now
        .Move (Screen.Width - .Width) / 2, (Screen.Height - .Height) / 3
        
        .ObjectToGUI
        gobjComm.Message cwiSHOWING_MODAL_FORM, vbNullString, NCID_CW_APPLICATION, NCID_INPATIENT_APPLICATION, False
        .show vbModal
        gobjComm.Message cwiDONE_SHOWING_MODAL_FORM, vbNullString, NCID_CW_APPLICATION, NCID_INPATIENT_APPLICATION, False
        
'__ if changes were made then refresh db and screen
        If .userCanceled Then
            TransferPatient = True
            Exit Function
        End If
        Screen.MousePointer = vbHourglass
        If oBed.BedNumber <> .oInpatient.Bed.BedNumber Or _
               oBed.Ward.LocationNCID <> .oInpatient.Bed.Ward.LocationNCID Or _
               oBed.Ward.Specialty <> .oInpatient.Bed.Ward.Specialty Then
               '-- this is a reportable action to TMIP
               
               colBeds.Add oBed
               colBeds.Add .oInpatient.Bed
               Call SendTMIPBedStatus(colBeds)
        End If
        
        If .oInpatient.objectstate = osDirty Then
            '__if he moved, recompute all capacities and redisplay everything
            gWards.RecomputeCapacity
            SelectedPatients.Save .oInpatient
            Call mfrmInpatient.Display
            SelectedPatients.Inpatient.objectstate = osClean
            '-- need to reload this guy in ginpatients
            If SelectedPatients.Inpatient.Status = "INPATIENT" Then
                Call gInPatients.ReloadPatient(.oInpatient)
            End If
        End If
        
        'Send the Inpatient Status message
        'SF 62830
        'This will eventually call I_RSTransfer_Transfer in this module
        Dim dm As Object
        Set dm = gobjComm.InitializeOLEServer("1999988")  'NCID_DATA_MANAGER
        Call dm.SendInpatientStatus(.oInpatient.ID)
        Set dm = Nothing

    End With
    Screen.MousePointer = vbDefault
    TransferPatient = True
    Exit Function
ErrHandler:
    MsgBox "Application error: " & Err.Description & " Occurred in PAD.TransferPatient"
    Exit Function
    Resume
    
End Function

Private Function AddMenus() As Boolean
'__ add the menus to core's list
On Error GoTo ErrHandler

    mlMenuAdd = gobjComm.AddActionMenuItem("&New Admission", "add.bmp", gInpatientProgID, "Add Patient")
    mlMenuEdit = gobjComm.AddActionMenuItem("&Edit", "edit.bmp", gInpatientProgID, "Edit currently selected inpatient record")
    mlMenuAddNote = gobjComm.AddActionMenuItem("&Add Note", "clinnote.bmp", gInpatientProgID, "Add a clinical note to the currently selected inpatient")
    mlMenuTransfer = gobjComm.AddActionMenuItem("&Patient Transfer", "transfer.bmp", gInpatientProgID, "Transfer currently selected inpatient")
    mlMenuDischarge = gobjComm.AddActionMenuItem("&Discharge", "summary.bmp", gInpatientProgID, "Discharge currently selected inpatient")
    mlMenuRefresh = gobjComm.AddActionMenuItem("&Refresh", "refresh.bmp", gInpatientProgID, "Refresh the display")
    mlMenuPrint = gobjComm.AddActionMenuItem("&Print", "print.bmp", gInpatientProgID, "Print current inpatient list")
    '--The system shall include an "Admin Close" Action item on the Inpatient List module.
    mlMenuAdminClose = gobjComm.AddActionMenuItem("&Admin Close", "EncDocNot.bmp", gInpatientProgID, "Administratively discharge the current patient")
    
    '__if the user doesn't have write privileges, turn off add, edit, transfer,discharge, addnote
    If Not UserCanWrite Then
        gobjComm.UpdateMenuItem mlMenuAdd, "disable"
        gobjComm.UpdateMenuItem mlMenuEdit, "disable"
        gobjComm.UpdateMenuItem mlMenuTransfer, "disable"
        gobjComm.UpdateMenuItem mlMenuDischarge, "disable"
        gobjComm.UpdateMenuItem mlMenuAddNote, "disable"
        gobjComm.UpdateMenuItem mlMenuAdminClose, "disable"
    End If
    If gWards.Count = 0 Then
        gobjComm.UpdateMenuItem mlMenuAdd, "disable"
    End If
    AddMenus = True
    Exit Function
ErrHandler:
    
End Function
Friend Function PatientMenus(enabled As Boolean) As Boolean
'__ turn on or off the several patient-specific menu items
'   edit, transfer, discharge, addnote
'<< SCR 47619

    
    gobjComm.UpdateMenuItem mlMenuAdd, "enable"
On Error GoTo ErrHandler
    Dim sAction As String
    If enabled Then
        sAction = "enable"
    Else
        sAction = "disable"
    End If
    
    gobjComm.UpdateMenuItem mlMenuEdit, sAction
    gobjComm.UpdateMenuItem mlMenuTransfer, sAction
    gobjComm.UpdateMenuItem mlMenuDischarge, sAction
    gobjComm.UpdateMenuItem mlMenuAddNote, sAction
    gobjComm.UpdateMenuItem mlMenuAdminClose, sAction
    
    '-- you have to be able to write to the record to write to the record
    If Not UserCanWrite Then
        gobjComm.UpdateMenuItem mlMenuAdd, "disable"
        gobjComm.UpdateMenuItem mlMenuEdit, "disable"
        gobjComm.UpdateMenuItem mlMenuTransfer, "disable"
        gobjComm.UpdateMenuItem mlMenuDischarge, "disable"
        gobjComm.UpdateMenuItem mlMenuAddNote, "disable"
        gobjComm.UpdateMenuItem mlMenuAdminClose, "disable"
    End If
    If Not SelectedPatients Is Nothing Then
        If SelectedPatients.Inpatient.Status = "DISCHARGED" Then
            gobjComm.UpdateMenuItem mlMenuEdit, "disable"
            gobjComm.UpdateMenuItem mlMenuTransfer, "disable"
            gobjComm.UpdateMenuItem mlMenuDischarge, "disable"
            gobjComm.UpdateMenuItem mlMenuAdminClose, "disable"
    '<< scr 42958 only providers can add a note post-discharge
            If Not CanSign Then
                gobjComm.UpdateMenuItem mlMenuAddNote, "disable"
            End If
        End If
    End If
    
    If gWards Is Nothing Then
        gobjComm.UpdateMenuItem mlMenuAdd, "disable"
    ElseIf gWards.Count = 0 Then
        gobjComm.UpdateMenuItem mlMenuAdd, "disable"
    End If
' SCR 71439 Let anyone discharge.
'<< scr 42975  Action item will only be enabled if the user has sign privileges for encounters
'    If Not CanSign Then
'        gobjComm.UpdateMenuItem mlMenuDischarge, "disable"
'        gobjComm.UpdateMenuItem mlMenuAdminClose, "disable"
'    End If
    '-- can't admin disch if assigned a bed
    If Not SelectedPatients Is Nothing Then
        If Val(SelectedPatients.Inpatient.AssignedBedNumber) <> 0 Then
            gobjComm.UpdateMenuItem mlMenuAdminClose, "disable"
        End If
    End If
    PatientMenus = True
    Exit Function
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "PAD_Inpatient.PatientMenus", "PAD_Inpatient", vbExclamation)
End Function

Friend Function TransferHistory() As Boolean
On Error GoTo ErrHandler
    If SelectedPatients.Inpatient Is Nothing Then Exit Function
    If SelectedPatients.Inpatient.ID = vbNullString Then Exit Function
    Dim frm As frmTransferHistory
    Dim mcol As New Collection
    Dim oTrans As cTransfer
    Set frm = New frmTransferHistory
    
    
    Dim oDal As cDAL
    Set oDal = New cDAL
    Dim oRS As ADODB.Recordset

    Set oRS = oDal.GetInpatientData(edti_Transfer, SelectedPatients.Inpatient.ID)
    Do While Not oRS.EOF
        Set oTrans = New cTransfer
        Call oTrans.DeserializeFromRecordset(oRS)
        mcol.Add oTrans
        oRS.MoveNext
    Loop
    
    '--- do a faux transfer record for the current state (for looks on the grid)
    Set oTrans = New cTransfer
    With SelectedPatients.Inpatient
        If Not .Bed Is Nothing Then
            oTrans.BedNumber = .Bed.BedNumber
            If Not .Bed.Ward Is Nothing Then
                oTrans.WardID = .Bed.Ward.WardID
            End If
        End If
        If Val(oTrans.BedNumber) = 0 Then
            oTrans.WardID = .Ward_ID
        End If
        oTrans.ProviderNCID = .AttendingProviderNCID
        oTrans.ServiceNCID = .PrimaryServiceNCID
        mcol.Add oTrans
    End With

    Call frm.Display(mcol)
    frm.Move (Screen.Width - frm.Width) / 2, (Screen.Height - frm.Height) / 3
    
    gobjComm.Message cwiSHOWING_MODAL_FORM, vbNullString, NCID_CW_APPLICATION, NCID_INPATIENT_APPLICATION, False
    frm.show vbModal
    gobjComm.Message cwiDONE_SHOWING_MODAL_FORM, vbNullString, NCID_CW_APPLICATION, NCID_INPATIENT_APPLICATION, False
    
    
    TransferHistory = True
    Exit Function
ErrHandler:
    TransferHistory = False
End Function


Public Function AdmitPatient(Patient As MMMHISPatient.Patient, _
              Optional ByVal AdmittingProviderNCID As String, _
              Optional ByVal AttendingProviderNCID As String, _
              Optional ByVal AdmissionDate As Date, _
              Optional ByVal EncounterNumber As String = vbNullString, _
              Optional ByRef ErrString As String, _
              Optional ByRef InpatientID As Long = 0, _
              Optional ByVal AdmittedFromEncNum As String = vbNullString) As Boolean
              '__ facility ncid is assumed to be the logon facility.
    
On Error GoTo ErrHandler
Dim frm As New frmNewAdmission

    If gInPatients Is Nothing Then
        Set gInPatients = New cInPatients
        Call LoadAllInpatients
    End If
    
    With frm
        '__ensure that this patient is not currently admitted
        Set .Inpatient = gInPatients.getInPatientByUnitNumber(Patient.UnitNumber)
        If Not .Inpatient Is Nothing Then
            ErrString = "Patient is already Admitted"
            AdmitPatient = False
            Exit Function
        End If
            
        Set .Inpatient = gInPatients.CreateInpatient(Patient, _
                                                   .searchSnoMed21.IConvert, _
                                                   AdmittingProviderNCID, _
                                                   AttendingProviderNCID, _
                                                   AdmissionDate, _
                                                   EncounterNumber, _
                                                   AdmittedFromEncNum)
                                                  
        If .Inpatient Is Nothing Then
            ErrString = "Error Creating Inpatient Record"
            AdmitPatient = False
            Exit Function
        End If
        
        If .Inpatient.Unit_Number = gobjPatient.UnitNumber Then
            .lblPatIDLine = gobjComm.GetPatientIDLine
        Else
            .lblPatIDLine = .Inpatient.PatientName
        End If
        
        .Move (Screen.Width - .Width) / 2, (Screen.Height - .Height) / 3
        Screen.MousePointer = vbDefault
        
        gobjComm.Message cwiSHOWING_MODAL_FORM, vbNullString, NCID_CW_APPLICATION, NCID_INPATIENT_APPLICATION, False
        .Display DirectAdmission
        gobjComm.Message cwiDONE_SHOWING_MODAL_FORM, vbNullString, NCID_CW_APPLICATION, NCID_INPATIENT_APPLICATION, False
        
        If .Cancelled Then
            ErrString = "User Cancelled Admission"
            AdmitPatient = False
            Exit Function
        End If
       
       
        Call CreateAppointment(.Inpatient)
        If .Inpatient.objectstate = osNew Then
            gInPatients.AddPatient .Inpatient
        End If
        
        Screen.MousePointer = vbHourglass
        
        Call gInPatients.Save(.Inpatient)
        
        InpatientID = .Inpatient.ID
        
    '<< SCR 49460  INPT: Add Note button does not function when selected from Outpatient Admission.
        '-- becuase we were not checking for it before
        If frm.UserClickedAddNote Then
            Call AddNote(.Inpatient)
        End If
    
    End With
    Set frm = Nothing
   
    
'    display on form  if it fits the criteria
'<< scr 48953 if entering from disposition, mfrm is not up
    If Not mfrmInpatient Is Nothing Then
        SelectedPatients.LoadInPatients CurrOptions
        mfrmInpatient.Display
    End If
    
    'Send the Inpatient Status message
    'SF 62830
    'This will eventually call I_RSTransfer_Transfer in this module
    Dim dm As Object
    Set dm = gobjComm.InitializeOLEServer("1999988")  'NCID_DATA_MANAGER
    Call dm.SendInpatientStatus(InpatientID)
    Set dm = Nothing
    
    'SF 12-10-04
    AdmitPatient = True
    
    Screen.MousePointer = vbDefault

Exit Function
ErrHandler:
    ErrString = "Unexpected error: " & Err.Description
    AdmitPatient = False
End Function

Private Sub PrintInpatientList()
Dim objForms As CHCSII_forms.Forms

  On Error GoTo ErrHandler
        
  Set objForms = New CHCSII_forms.Forms
  
  If mfrmInpatient Is Nothing Then Exit Sub
  If Not objForms.printInpatient(mfrmInpatient.fGrid.hwnd, "", "") Then
  End If
  
  Set objForms = Nothing

  Exit Sub
    
ErrHandler:
  gobjShared.ShowVBError Err.Number, Err.Description, "CHCSII_PAD.PrintInpatientList", "CHCSII_Inpatient", vbExclamation

End Sub

Friend Function PatientDischarge() As Boolean

'Prompt the provider for thier password and discharge the patient.

On Error GoTo ErrHnd
    
    Dim dPotentialDischargeDate As Date
    Dim oDal As cDAL
    Dim oRS As Recordset
    
    'SCR 102116 Check that the discharge summary has been done
    Set oDal = New cDAL
    'Added the updated status: SCR 103875
    Set oRS = oDal.GetInpatientData(edti_HasInptNoteType, SelectedPatients.Inpatient.AppointmentID, NCID_Discharge_Note, "205657, 14510309")
    'Set oRS = oDal.GetInpatientData(edti_Diagnoses, SelectedPatients.Inpatient.ID, "DISCHARGE")
    If (oRS("count") = 0) Then
        'We don't have a requirement to check for a discharge order.
        'SCR 103021
        'If Not frmInpatient.HasDischargeOrder(SelectedPatients.Inpatient.Unit_Number) Then
            MsgBxARMd "Please complete a signed Discharge Summary note for this inpatient visit.", vbInformation
            Exit Function
        'End If
    End If
    
    'Confirm Discharge
    With SelectedPatients.Inpatient
        Set frmConfirmDischarge.User = objUser
        'This date is for display purposes only, the date is pre-determined by disposition or
        'if it is blank, it will be set to Now in cDal.SetInpatientData
        If .PotentialDischargeDate = Empty Then
            dPotentialDischargeDate = Now
        Else
            dPotentialDischargeDate = gobjComm.GmtToLocal(.PotentialDischargeDate)
        End If
        'SCR 106287
        If dPotentialDischargeDate < gobjComm.GmtToLocal(.AdmissionDate) Then
            dPotentialDischargeDate = gobjComm.GmtToLocal(.AdmissionDate)
        End If
        If dPotentialDischargeDate < gobjComm.GmtToLocal(.LastTransferDate) Then
            dPotentialDischargeDate = gobjComm.GmtToLocal(.LastTransferDate)
        End If
        frmConfirmDischarge.Display gobjPatient.Name & "@" & Format(dPotentialDischargeDate, gsDefaultDateTime), gobjComm.GmtToLocal(.AdmissionDate), gobjComm.GmtToLocal(.LastTransferDate)
    End With
    
    If frmConfirmDischarge.bDischargePatient Then
        Dim oDiag As cDiagnosis
        Dim oRSD As ADODB.Recordset
        
        Set oRSD = oDal.GetInpatientData(edti_Diagnoses, SelectedPatients.Inpatient.ID, "DISCHARGE")
        If oRSD.EOF = False Then
            Set oDiag = New cDiagnosis
            Call oDiag.DeserializeFromRecordset(oRSD)
            'SCR 131342 Added True so that the discharge data is not removed
            DischargePatient SelectedPatients.Inpatient.ID, dPotentialDischargeDate, DischargeComments, vbNullString, _
                        vbNullString, vbNullString, vbNullString, oDiag.SnoID, oDiag.Description, True
        Else
            'SCR 131342 Added True so that the discharge data is not removed
            DischargePatient SelectedPatients.Inpatient.ID, dPotentialDischargeDate, DischargeComments, vbNullString, _
                        vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, True
        End If
        oRSD.Close
    End If
            
Exit Function

ErrHnd:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "PAD_Inpatient.PatientDischarge", "PAD_Inpatient", vbExclamation)
    
    Exit Function
    Resume
End Function


Public Function DischargePatient(ByVal Inpatient_ID As String, _
                                 ByVal Discharge_Date As Date, _
                                 ByVal Discharge_Comment As String, _
                                 ByVal Disposition_NCID As String, _
                                 ByVal Disposition_Comment As String, _
                                 ByVal Discussed_Items As String, _
                                 ByVal Discussed_Comment As String, _
                                 ByVal Discharge_Diagnosis_SnoID As String, _
                                 ByVal Discharge_Diagnosis_Text As String, _
                                 Optional ByVal DischargeOnly As Boolean = False) As Boolean
 


'-- discussed_Items is a string of form: "0/1/2/3/NCID"
'   where the first 4 elements are the diagnosis, Medications, SideEffects, Alternatives, respectively
'   and the ncid is the "discussed with whom" ncid

On Error GoTo ErrHandler
    Dim dischNote As String
    Dim oPat As cInpatient
    Dim oDal As cDAL
    Set oDal = New cDAL
    Dim Col As New Collection
    
    
    Dim gobjAppointmentOps As CHCSII_AppointmentClient.AppointmentOps
    Screen.MousePointer = vbHourglass
    
'<< scr 47929 decode the dispostion ncid and add to disposition comment
    Dim sDisp As String
'<< scr 42981 add "user @ time"
'<< scr 49575 add a crlf to the end of that line
    If LenB(Disposition_NCID) > 0 Then
        sDisp = Disposition_NCID_to_Text(Disposition_NCID)
        dischNote = gobjLogon.username & " @" & Now & vbCrLf _
                  & "Discharge Disposition: " & sDisp & vbCrLf _
                  & "Discharge Diagnosis: " & Discharge_Diagnosis_Text & vbCrLf _
                  & "Discharge Comment: " & Discharge_Comment
    Else
        dischNote = gobjLogon.username & " @" & Now & vbCrLf _
                  & "Discharge Comment: " & Discharge_Comment
    End If
    
    'Only update the discharge info if DischargeOnly is false
    If DischargeOnly = False Then
        '-- this updates the db, then we have to reload the object in ginpatients
        Call oDal.SetInpatientData(edti_UpdateDischargeInfo, _
                                    Inpatient_ID, _
                                    gobjComm.LocalToGmt(Discharge_Date), _
                                    dischNote, _
                                    Disposition_NCID, _
                                    Disposition_Comment, _
                                    Discussed_Items, _
                                    Discussed_Comment, _
                                    Discharge_Diagnosis_SnoID, _
                                    Discharge_Diagnosis_Text)
    End If
    
    'Discharge the patient
    Call oDal.SetInpatientData(edti_Discharge, Inpatient_ID)
    
    Set oPat = gInPatients.getInPatientByID(Inpatient_ID)
    
    '56672
    'Discontinue all inpatient orders
    oDal.DiscontinueAllInpatientOrders gobjComm.GmtToLocal(oPat.AdmissionDate), Now, oPat.Unit_Number
                                
    Screen.MousePointer = vbHourglass
    '--clear that bed
    
'<< scr  47109 change the status to 'completed'
    Set gobjAppointmentOps = New CHCSII_AppointmentClient.AppointmentOps
    gobjAppointmentOps.Initialize 1, gobjComm, gobjLogon, gobjconfig, gobjPatient, objUser
    Call gobjAppointmentOps.SetEncounterStatus(oPat.AppointmentID, 4, True) '-  4 is completed
    Set gobjAppointmentOps = Nothing
    
    If Not oPat Is Nothing Then
        Call gInPatients.ReloadPatient(oPat)
'<< scr 43003 send status on discharge
        Col.Add oPat.Bed
        SendTMIPBedStatus Col
        
        'SF SCR 72841 Clear the patient from the bed and save it
        Set oPat.Bed.Patient = Nothing
        If Val(oPat.Ward_ID) <> 0 Then
            Call gWards.SaveWard(gWards(oPat.Ward_ID))
        End If


        Call oPat.Bed.Ward.ComputeCapacity
'        Set oPat = gInPatients.getInPatientByID(Inpatient_ID)
        If SelectedPatients.Inpatient.ID = oPat.ID Then
            Set SelectedPatients.Inpatient = oPat
        End If
        If Not mfrmInpatient Is Nothing Then
            SelectedPatients.LoadInPatients CurrOptions
            'may just need a dislay here
            Call MenuItemSelected(mlMenuRefresh)
        End If
    Else
        MsgBxARMd "An unexpected error occurred in the discharge patient process", vbExclamation, "Patient Discharge"
        DischargePatient = False
        Exit Function
    End If
    
    'Send the PVF file and the JMEWS file. The JMEWS file will be sent with the PVF file.
    'Send the Inpatient Status message
    'SF 62830
    'This will eventually call I_RSTransfer_Transfer in this module
    Dim dm As Object
    Set dm = gobjComm.InitializeOLEServer("1999988")  'NCID_DATA_MANAGER
    Call dm.SendInpatientStatus(CLng(Inpatient_ID))
    'SF 76357 8-25-05
    Call dm.Send2TMIP(11) 'from encstatus.Complete enum

    Set dm = Nothing

'    display on form  if it fits the criteria
    Screen.MousePointer = vbDefault
    DischargePatient = True
    Exit Function
    
ErrHandler:
    MsgBxARMd "Application error: " & Err.Description & " Occurred in PAD.DischargePatient"

    DischargePatient = False
    Exit Function
    Resume
End Function
Public Function UpdateAdmissionDischargeDisposition(ByVal Inpatient_ID As String, _
                                 ByVal Discharge_Date As Date, _
                                 ByVal Discharge_Comment As String, _
                                 ByVal Disposition_NCID As String, _
                                 ByVal Disposition_Comment As String, _
                                 ByVal Discussed_Items As String, _
                                 ByVal Discussed_Comment As String, _
                                 ByVal Discharge_Diagnosis_SnoID As String, _
                                 ByVal Discharge_Diagnosis_Text As String, _
                                 ByVal Discharge_FollowUp_Text As String) As Boolean

'-- discussed_Items is a string of form: "0/1/2/3/NCID"
'   where the first 4 elements are the diagnosis, Medications, SideEffects, Alternatives, respectively
'   and the ncid is the "discussed with whom" ncid

On Error GoTo ErrHandler
    Dim dischNote As String
    Dim oPat As cInpatient
    Dim oDal As cDAL
    Set oDal = New cDAL
    Dim Col As New Collection
    
    
    Screen.MousePointer = vbHourglass
    
'<< scr 47929 decode the dispostion ncid and add to disposition comment
    Dim sDisp As String
'<< scr 42981 add "user @ time"
'<< scr 49575 add a crlf to the end of that line
    If LenB(Disposition_NCID) > 0 Then
        sDisp = Disposition_NCID_to_Text(Disposition_NCID)
        dischNote = gobjLogon.username & " @" & Now & vbCrLf _
                  & "Discharge Disposition: " & sDisp & vbCrLf _
                  & "Discharge Diagnosis: " & Discharge_Diagnosis_Text & vbCrLf _
                  & "Discharge Comment: " & Discharge_Comment
    Else
        dischNote = gobjLogon.username & " @" & Now & vbCrLf _
                  & "Discharge Comment: " & Discharge_Comment
    End If
    
    dischNote = dischNote & vbCrLf & Discharge_FollowUp_Text
    
    '-- this updates the db, then we have to reload the object in ginpatients
    Call oDal.SetInpatientData(edti_UpdateDischargeInfo, _
                                Inpatient_ID, _
                                gobjComm.LocalToGmt(Discharge_Date), _
                                dischNote, _
                                Disposition_NCID, _
                                Disposition_Comment, _
                                Discussed_Items, _
                                Discussed_Comment, _
                                Discharge_Diagnosis_SnoID, _
                                Discharge_Diagnosis_Text)

    '--- SCR 74320   Sherry Wang  1/18/2005
'    Set oPat = gInPatients.getInPatientByID(Inpatient_ID)
      
'    Screen.MousePointer = vbHourglass
'    '--clear that bed
'
''TODO: Do we really need this?
'    If Not oPat Is Nothing Then
'        Call gInPatients.ReloadPatient(oPat)
'
'        If SelectedPatients.Inpatient.ID = oPat.ID Then
'            Set SelectedPatients.Inpatient = oPat
'        End If
'        If Not mfrmInpatient Is Nothing Then
'            SelectedPatients.LoadInPatients CurrOptions
'            'may just need a display here
'            Call MenuItemSelected(mlMenuRefresh)
'        End If
'    Else
'        MsgBxARMd "An unexpected error occurred in the discharge patient process", vbExclamation, "Patient Discharge"
'        UpdateAdmissionDischargeDisposition = False
'        Exit Function
'    End If
    
'    display on form  if it fits the criteria
    Screen.MousePointer = vbDefault
    UpdateAdmissionDischargeDisposition = True
    Exit Function
    
ErrHandler:
    MsgBxARMd "Application error: " & Err.Description & " Occurred in PAD.DischargePatient"

    UpdateAdmissionDischargeDisposition = False
    Exit Function
    Resume
End Function


Public Function AddNote(oPat As cInpatient, Optional SpecifiedNoteTypeNCID As String = "", Optional sHelpContextId As String = "") As Boolean
    
    Dim objEnc As CHCSIIEncounterCurrent.EncounterParent
    Dim oEncounter As CHCSIIEncounterOps.Encounter
    Dim Spec As cSpecialty
    Dim SpecName As String
    Dim frmMyNoteTypes As frmNoteTypes
    Dim sInpatientNoteTypeNCID As String
    Dim oEncOps As CHCSIIEncounterOps.EncounterOps
    Dim lHospDay As Long
    Dim oDal As cDAL
    Dim oRS As Recordset
    
On Error GoTo ErrHandler

    If oPat Is Nothing Then Exit Function
    'if selectedpatients.Inpatient.
    '<< scr 42898 control 'write' access
    If Not UserCanWrite Then Exit Function
    
    '----------------------
    'Prompt the user to select which type of inpatient note to create.
    Set frmMyNoteTypes = New frmNoteTypes
    
    ' Quan 79941
    If sHelpContextId <> "" Then
         frmMyNoteTypes.HelpContextID = sHelpContextId
    End If
    
    'Determine if the user has sign privileges.
    'RQT 56556
    Set objEnc = gobjComm.InitializeOLEServer(NCID_ENCOUNTER_APPLICATION)
    Set oEncOps = objEnc.EncounterOps
    frmMyNoteTypes.HasSignPriv = oEncOps.UserCanBePrimaryProvider(gobjLogon.UserNCID)
    
    'RQT 56563
    If SpecifiedNoteTypeNCID = "" Then
        If frmMyNoteTypes.GetUserInput = True Then
            sInpatientNoteTypeNCID = frmMyNoteTypes.SelectedNoteTypeNCID
        Else
            Exit Function
        End If
    Else
        sInpatientNoteTypeNCID = SpecifiedNoteTypeNCID
    End If
    '----------------------
            
'    '__ if the patient is discharged, only add a note to the post-discharge note area.
'    If oPat.Status = "DISCHARGED" Then
'        AddNote = AddPostDischargeNote(frmMyNoteTypes.SelectedNoteTypeString)
'        Exit Function
'    End If
    
    '76436 If the patient already has a discharge summary note for this appointment, don't let them
    'create another discharge summary note.
    If sInpatientNoteTypeNCID = NCID_Discharge_Note Then
        Set oDal = New cDAL
        Set oRS = oDal.GetInpatientData(edti_HasInptNoteType, oPat.AppointmentID, sInpatientNoteTypeNCID)
        If oRS("count") > 0 Then
            MsgBxARMd "A discharge summary already exists for this inpatient visit.", vbInformation
            Exit Function
        End If
    End If
    
    
    Screen.MousePointer = vbHourglass
    
    With oPat
        Set Spec = gSpecialties.GetSpecialty(.PrimaryServiceNCID)
        If Not Spec Is Nothing Then
            SpecName = Spec.SpecName
        End If
    '-- be sure we have an apt id or create enc fails
        If Val(.AppointmentID) = 0 Then
            Call CreateAppointment(SelectedPatients.Inpatient)
        End If
        
        Set objEnc = gobjComm.InitializeOLEServer(NCID_ENCOUNTER_APPLICATION)
        
        'Calculate the hospital day
        If .AdmissionDate <> 0 Then
            lHospDay = DateDiff("d", gobjComm.GmtToLocal(.AdmissionDate), Now) + 1
        End If
        
'        Set oEncounter = objEnc.CreateEncounter(.AppointmentID, "INPT", 5, _
'               gobjLogon.FacilityNCID, gobjLogon.NursingDivisionNCID, _
'               .AttendingProviderNCID, , , DateAdd("n", -.TimeBias, .AdmissionDate), False, _
'               , , , .Unit_Number, , , , , , , SpecName, , , sInpatientNoteTypeNCID, lHospDay)
        
        '64769 SF 2-16-05 Changed .AddendingProviderNCID to gobjLogon.UserNCID
         Set oEncounter = objEnc.CreateEncounter(.AppointmentID, "INPT", 5, _
               gobjLogon.FacilityNCID, gobjLogon.NursingDivisionNCID, _
               gobjLogon.UserNCID, , , gobjComm.GmtToLocal(.AdmissionDate), False, _
               , , , .Unit_Number, , , , , , , SpecName, , , sInpatientNoteTypeNCID, lHospDay)
       
        Screen.MousePointer = vbDefault
        
        If oEncounter Is Nothing Then
            Exit Function 'The encounter was not loaded because the user cancelled the loading. SCR 108581
        End If
        Call objEnc.SetCurrentEncounter(.Unit_Number, gobjLogon.FacilityNCID, oEncounter.EncounterID)
        
        'This is in encounter ops encounter.cls - admit patient
'        'Copy forward the encounter
'        If EncNumToCopyFrom <> "" Then
'            objEnc.CopyForwardEncounter .Unit_Number, gobjLogon.FacilityNCID, EncNumToCopyFrom
'        End If
        
    End With
    
    If sInpatientNoteTypeNCID = NCID_Orders_Note Then
        'RQT 56577
        'If the provider created a "orders note", then open up Orders Summary.
        gobjComm.Message cwiGENERIC, "LAUNCH|" & NCID_ASSESMENT_AND_PLAN_APPLICATION & "|" & NCID_ORDERS_SUMMARY_APPLICATION, NCID_ENCOUNTER_APPLICATION, NCID_INPATIENT_APPLICATION
    Else
        gobjComm.Message cwiSTART_OLE_SERVER, NCID_ENCOUNTER_APPLICATION, NCID_CW_APPLICATION, NCID_INPATIENT_APPLICATION
        
        If sInpatientNoteTypeNCID = NCID_Admission_HP_Note Then
            'RQT 56543, SCR 64209
            If Val(oPat.AdmittedFromEncNum) <> 0 Then 'SCR 74118 - When the patient is first created, the value is ""
                '--- SCR 85371   Sherry Wang  1/25/2006
                If MsgBxARMd("Do you want to copy forward the information from the outpatient encounter?", vbYesNo + vbQuestion) = vbYes Then
                    Set objEnc = gobjComm.InitializeOLEServer(NCID_ENCOUNTER_APPLICATION)
                    objEnc.CopyForwardEncounter gobjPatient.UnitNumber, gobjLogon.FacilityNCID, oPat.AdmittedFromEncNum
                End If
            End If
        End If
    End If
    
    AddNote = True
    Exit Function
ErrHandler:
    gobjShared.ShowVBError Err.Number, Err.Description, "Inpatient.AddNote", "Error creating encounter", vbExclamation
    AddNote = False
    
    Exit Function
    Resume
End Function

Friend Function AdminDischarge(Optional Unit_Number As String = "") As Boolean
    On Error GoTo ErrHandler
    Dim oPat As cInpatient
    Dim bAdminDischarge As Boolean
    Dim iResponse As Integer
    Dim fAdminClose As frmAdminClose
    
'-- get the patient in question
    If Unit_Number <> vbNullString Then
        Set oPat = gInPatients.getInPatientByUnitNumber(Unit_Number)
    Else
        Set oPat = SelectedPatients.Inpatient
    End If
    
    bAdminDischarge = True
    '-- rules:
    'The system shall only allow a user to administratively close an inpatient item that meets the following criteria:
    '1) has a patient assigned,
    If oPat Is Nothing Then
        bAdminDischarge = False
    End If
    
    '2) no bed assignment, and
    If Not oPat.Bed Is Nothing Then
        If Not Val(oPat.Bed.BedNumber) = 0 Then
            bAdminDischarge = False
        End If
    End If
    
    '3) status is not "Discharged".
    If Not oPat.Status = "INPATIENT" Then
        bAdminDischarge = False
    End If

    '-- ok we're good to go, now discharge him
    If bAdminDischarge = True Then
        iResponse = MsgBxARMd("Are you sure you want to close this inpatient item?", vbYesNo + vbQuestion, "Admin Close")
           If iResponse = vbNo Then
               Exit Function
           Else
                gobjComm.Message cwiSHOWING_MODAL_FORM, vbNullString, NCID_CW_APPLICATION, NCID_INPATIENT_APPLICATION, False
                Set fAdminClose = New frmAdminClose
                Set fAdminClose.Parent = Me
                fAdminClose.show vbModal
                gobjComm.Message cwiDONE_SHOWING_MODAL_FORM, vbNullString, NCID_CW_APPLICATION, NCID_INPATIENT_APPLICATION, False
                
                If oPat.Diagnoses.Count > 0 Then
                    AdminDischarge = DischargePatient(oPat.ID, Now, DischargeComments, vbNullString, _
                                vbNullString, vbNullString, vbNullString, _
                                oPat.Diagnoses.Item(1).SnoID, oPat.Diagnoses.Item(1).Description)
                Else
                    AdminDischarge = DischargePatient(oPat.ID, Now, DischargeComments, vbNullString, _
                                vbNullString, vbNullString, vbNullString, _
                                vbNullString, vbNullString, vbNullString)
                End If
                
           End If
    End If
    Exit Function
ErrHandler:
    AdminDischarge = False
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "PAD_Inpatient.AdminDischarge", "PAD_Inpatient", vbExclamation)
End Function

Private Function CreateAppointment(oPat As cInpatient) As Long
'__ returns the appointment ID for this patient

On Error GoTo ErrHnd

    Dim objClinicOps    As New CHCSII_ClinicClient.ClinicOps
    Dim mcolMyClinics   As New Collection
    
    
    Dim gobjAppointmentOps As CHCSII_AppointmentClient.AppointmentOps
    Dim objAppointment  As CHCSII_AppointmentClient.Appointment
    

        Set gobjAppointmentOps = New CHCSII_AppointmentClient.AppointmentOps
        gobjAppointmentOps.Initialize 1, gobjComm, gobjLogon, gobjconfig, gobjPatient, objUser
       
        Screen.MousePointer = vbHourglass
    
        Set objAppointment = gobjAppointmentOps.GetNewAppointment
          
        If objAppointment Is Nothing Then
            MsgBxARMd "Unable to create an appointment.", vbCritical, "Unexpected Error"
            Exit Function
        End If
        With objAppointment
            
            .ApptComment = "Inpatient Record"
            '<< scr #47822 appoihtment date is admission date minus offset
            .ApptDateTime = gobjComm.GmtToLocal(oPat.AdmissionDate)   '  move back to local time here
            '.ApptDateTime = DateAdd("n", -oPat.TimeBias, oPat.AdmissionDate)  '  move back to local time here
            .ApptReason = vbNullString
            .ApptType = "INPT"
            .ApptClassification = "Inpatient"  '-- inpatient  : scr 47180
            .ApptStatus = "KEPT"   '<< scr 47109  this should be 'kept' and appts will make it 'in progress' on checkin
            
            .ClinicianNCID = oPat.AttendingProviderNCID
'<< SCr #42942 if clinician ncid is blank, we can't create an encounter. so use unassigned there.
'<< scr 45073 attend/admittin provider ncids are now comming back as -1, so change the =0 to < 1
            If Val(.ClinicianNCID) < 1 Then
                .ClinicianNCID = GetUnassignedNCID
            End If
            .ClinicNCID = gobjLogon.NursingDivisionNCID
            .PatientUnitNumber = oPat.Unit_Number
            .FacilityNCID = gobjLogon.FacilityNCID
            .EncounterStatus = 3        'superceded: << scr 47109 enc status is 4
                                        '<< scr 47109 status is now 'in progress'
            
        End With
        
    '__ Save the new appt and get appt id
        CreateAppointment = gobjAppointmentOps.SaveNewAppointment(objAppointment)

        'Call gobjAppointmentOps.SetEncounterStatus(CreateAppointment, EncInProgress, True)
        oPat.AppointmentID = CreateAppointment
        
        
Cleanup:
    Screen.MousePointer = vbDefault
    Set objAppointment = Nothing
    Set gobjAppointmentOps = Nothing
    
    Exit Function

ErrHnd:
    MsgBxARMd "Application Error: " & Err.Description & " Occured in CHCSII_PAD.Inpatient.CreateAppointment"
    GoTo Cleanup
    
End Function

Friend Function AddPostDischargeNote() As Boolean

On Error GoTo ErrHandler

    Dim frm As New frmTextNote
    Dim Category As String
    Dim Title As String
    Dim comment As String
    Dim oDal As cDAL
    Dim NewNote As String
    Dim msg As String
    Screen.MousePointer = vbHourglass
    
    '--- SCR 68731  Sherry Wang  1/27/2006
    Set frm.ParentForm = Me.NoteParentForm
    
'<< scr 42920    " any user who has “Can Cosign” privileges for the “Current_Encounter” function
    msg = "HasPrivilege|" & objUser.userid & "|Current_Encounter|" & Priv_Cosign
    Call gobjComm.Message(cwiGENERIC, msg, NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION)
    If msg <> "TRUE" Then
        Screen.MousePointer = vbDefault
        MsgBxARMd "You do not have sufficient privileges to add a post-discharge note"
        Exit Function
    End If
    '>>
    
    Screen.MousePointer = vbDefault
    gobjComm.Message cwiSHOWING_MODAL_FORM, vbNullString, NCID_CW_APPLICATION, NCID_INPATIENT_APPLICATION, False
    frm.Display gobjLogon.username, Category, Title, comment
    gobjComm.Message cwiDONE_SHOWING_MODAL_FORM, vbNullString, NCID_CW_APPLICATION, NCID_INPATIENT_APPLICATION, False
    
    If Not frm.userCanceled Then
        Screen.MousePointer = vbHourglass
        If Not Trim$(comment) = vbNullString Then
            NewNote = "By: " & gobjLogon.username & " @ " & Now & vbCrLf
            NewNote = NewNote & "Category: " & Category & vbCrLf
            NewNote = NewNote & "Title: " & Title & vbCrLf
            NewNote = NewNote & comment & vbCrLf & vbCrLf
            Set oDal = New cDAL
            Call oDal.SetInpatientData(edti_PostDischargeNote, SelectedPatients.Inpatient.ID, NewNote)
            SelectedPatients.Inpatient.PostDischargeNote = SelectedPatients.Inpatient.PostDischargeNote & NewNote
            '--- SCR 68731  Sherry Wang  1/27/2006
            Me.NoteParentForm.txtPostDischComments = SelectedPatients.Inpatient.PostDischargeNote
        End If
    End If
    Screen.MousePointer = vbDefault
    AddPostDischargeNote = True
    Exit Function
ErrHandler:
    Screen.MousePointer = vbDefault
    AddPostDischargeNote = False
    Exit Function
End Function


Public Property Get DischargeComments() As String
    DischargeComments = msDischargeComments
End Property

Public Property Let DischargeComments(ByVal vNewValue As String)
    msDischargeComments = vNewValue
End Property


Public Function GetSpecialtyLoaction() As Collection
    On Error GoTo ErrHandler
    
    Dim oWard As cWard
    
    Set GetSpecialtyLoaction = New Collection
    
    For Each oWard In gWards
        '170082 Don't include INACTIVE wards.
        If oWard.Status <> "INACTIVE" Then
            GetSpecialtyLoaction.Add oWard.SpecialtyName & "/" & oWard.LocationName
        End If
    Next oWard
    
    Exit Function
ErrHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "Inpatient.GetSpecialtyLoaction", "PAD_Inpatient", vbExclamation)
End Function

Public Sub DeSelectGrid()
    
    mfrmInpatient.fGrid.row = 0
    PatientMenus False
    
End Sub

'--- SCR 68731  Sherry Wang  1/27/2006
Friend Property Get NoteParentForm() As frmNewAdmission
    Set NoteParentForm = moNoteParentForm
End Property

Friend Property Set NoteParentForm(ByVal vNewValue As frmNewAdmission)
    Set moNoteParentForm = vNewValue
End Property
