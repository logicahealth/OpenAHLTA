VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ListManagement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"ListManagementOps"
Attribute VB_Ext_KEY = "Member1" ,"ListManagementOps2"
Option Explicit
#If DebugOn Then
    Private DL As cDebugLog
    Private Const Exe = "ListManagement"
    Private Const Module = "ListManagement.cls"
#End If
Implements ICHCSII_CoreAPI2
'--- Menu actions
'local variable(s) to hold property value(s)
Private mvarOps As ListManagementOps 'local copy

Friend Property Set ListManagementOps(RHV As ListManagementOps)
    'SCR-47892 RBELL 1/21/04 Added freind property to reset ovject from form
    Set mvarOps = RHV
End Property

Public Property Get ListManagementOps() As ListManagementOps
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ProblemOps

    'New Default ProblemOps Object
    If mvarOps Is Nothing Then
        Set mvarOps = New ListManagementOps
    End If
    
    'Return Default Object
    Set ListManagementOps = mvarOps
    
        

End Property

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
Public Sub Initialize(ByVal vlMyProgId As Long, robjComm As Object, robjLogonEnviron As Object, robjConfig As Object, robjPatient As Object)
MsgBox "Tester: This method 'Problems.Initialize' should not have been called." _
    & vbCrLf & "Please create a SCR. Do not assign it to problems." & vbCrLf _
    & vbCrLf & "The calling application is not following design rules.", vbCritical, App.Title
    
    On Error GoTo errhandler
    #If DebugOn Then
        Dim DR As DebugRec
        If Not DL Is Nothing Then
            DL.DebugLog Exe, Module, "Initialize", "Begin", DR
        End If
    #End If

    '---set global variables---
    cwlMyProgId = vlMyProgId

    '---create a pointer to objects---
    Set cwobjCWComm = robjComm
    Set cwobjLogonEnviron = robjLogonEnviron
    Set cwobjConfig = robjConfig
    Set cwobjPatient = robjPatient

    '---set global flags---
    #If DebugOn Then
        If Not DL Is Nothing Then
            DL.DebugLog Exe, Module, "Initialize", "End", DR
        End If
    #End If

    Exit Sub

errhandler:
    #If DebugOn Then
        If Not DL Is Nothing Then
            DL.DebugLog Exe, Module, "Initialize", "ErrHandler", DR
        End If
    #End If
    Call gobjShared.ShowVBError(Err.Number, Err.DESCRIPTION, "ListManagement.Initialize", App.Title, vbExclamation)

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
    Select Case viViewType
        Case cwiREGULAR_VIEW
            QueryView = True    '--this view is NOT optional--
        Case cwiPATIENT_SUMMARY_VIEW
            QueryView = False
        Case cwiPROPERTY_VIEW
            QueryView = False
        Case cwiPRINT_REGULAR_VIEW
            QueryView = False
        Case cwiPRINT_SUMMARY_VIEW
            QueryView = False
        
        Case Else   '--not supported--
            QueryView = False
    End Select
End Function

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
Public Function OpenView(ByVal viViewType As Integer, ByVal vlParenthWnd As Long, ByRef rhMainForm As Long) As Boolean

    On Error GoTo errhandler
    #If DebugOn Then
        Dim DR As DebugRec
        If Not DL Is Nothing Then
            DL.DebugLog Exe, Module, "OpenView", "Begin", DR
        End If
    #End If

    '---perform actions based on the "view type"---
    'Innocent till Proven Guilty
    OpenView = True
    
    Dim lMouse As MousePointerConstants
    
    lMouse = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    
    Select Case viViewType
    Case cwiREGULAR_VIEW
        '---------------------------------------------------------------------------------------
        '  If this OLE Server needs it's own menu items, here is a good place to add these
        '    menu items onto the CW Core's main menu.
        '
        '  NOTE: This is OPTIONAL.  You do NOT have to create menu items. Thats why they are
        '        commented out.
        '---------------------------------------------------------------------------------------
        '-----------------------------------------------------------------------------------------
        '  Reset the "parent window" for this OLE server's "main" regular form.  This is the main
        '    "trick" that makes the OLE server look like it is part of the CW Core.
        '
        '  NOTE: This action invokes the main form's "Load" event, but the form is NOT shown.
        '        Make sure that you DO NOT perform a "Me.Show" in the form's "Load" event, as this
        '        will mess up the "SetParent" API call.  Make sure that you DO NOT "Show" the
        '        "main" regular form in this method, that will happen in the "ResizeView" method.
        '-----------------------------------------------------------------------------------------
        frmListManagement.SetListManagement Me
        SetParent frmListManagement.hWnd, vlParenthWnd
        rhMainForm = frmListManagement.hWnd
        
        frmListManagement.RefreshDisplay           'Only Fill Form, do not display it

        OpenView = True
        
    Case Else
        Screen.MousePointer = vbDefault
        Beep
        MsgBox "Unsupported view type: " & CStr(viViewType) & "ListMangement.OpenView.", vbCritical, App.Title

        '---set return value---
        OpenView = False

    End Select

    
    Screen.MousePointer = lMouse
    
    #If DebugOn Then
        If Not DL Is Nothing Then
            DL.DebugLog Exe, Module, "OpenView", "End", DR
        End If
    #End If

    Exit Function

errhandler:
    #If DebugOn Then
        If Not DL Is Nothing Then
            DL.DebugLog Exe, Module, "OpenView", "ErrHandler", DR
        End If
    #End If
    Call gobjShared.ShowVBError(Err.Number, Err.DESCRIPTION, "ListMangement.OpenView", App.Title, vbExclamation)
    '---set return value---
    OpenView = False

    Exit Function
End Function
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
    On Error GoTo errhandler
    #If DebugOn Then
        Dim DR As DebugRec
        If Not DL Is Nothing Then
            DL.DebugLog Exe, Module, "ResizeView", "Begin", DR
        End If
    #End If

    Dim lMouse As MousePointerConstants
    
    lMouse = Screen.MousePointer
    Screen.MousePointer = vbHourglass

    '---perform different actions based upon the "view type"---
    Select Case viViewType
        Case cwiREGULAR_VIEW
            'Resize and Show form if required
            frmListManagement.ResizeView vlLeft, vlTop, vlWidth, vlHeight
            
        Case Else
            Beep
            MsgBox "Unsupported resize view type: " & CStr(viViewType) & "ListMangement.ResizeView.", vbExclamation, App.Title
            
        End Select
    
        Screen.MousePointer = lMouse
        
    #If DebugOn Then
        If Not DL Is Nothing Then
            DL.DebugLog Exe, Module, "ResizeView", "End", DR
        End If
    #End If

    Exit Sub

errhandler:
    #If DebugOn Then
        If Not DL Is Nothing Then
            DL.DebugLog Exe, Module, "ResizeView", "ErrHandler", DR
        End If
    #End If
    
    Call gobjShared.ShowVBError(Err.Number, Err.DESCRIPTION, "ListMangement.ResizeView", App.Title, vbExclamation)
    Screen.MousePointer = lMouse
    
    Exit Sub
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
    #If DebugOn Then
        Dim DR As DebugRec
        If Not DL Is Nothing Then
            DL.DebugLog Exe, Module, "CloseView", "Begin", DR
        End If
    #End If
     
    
    Dim lMouse As MousePointerConstants
    
    lMouse = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    
    Select Case viViewType
        Case cwiALL_VIEWS
        '---this indicates that CW is closing the "view", NOT the OLE Server---
        'Close All Possible Views
            Unload frmListManagement
            
        Case cwiREGULAR_VIEW
        '---this indicates that CW is closing the "view", NOT the OLE Server---
            Unload frmListManagement
        
    End Select
    
    If Not mvarOps Is Nothing Then mvarOps.Save
    
    Screen.MousePointer = lMouse
    
    #If DebugOn Then
        If Not DL Is Nothing Then
            DL.DebugLog Exe, Module, "CloseView", "End", DR
        End If
    #End If
End Sub

'*******************************************************************************************
'  Sub Name: MenuItemSelected
'
'  Description: This method is called when the user has selected a menu item, "owned" by
'               this OLE server.
'
'  Parameters:
'    vlMenuId - Contains the menu id of the menu item the user selected.
'*******************************************************************************************
Public Sub MenuItemSelected(ByVal vlMenuId As Long)
    #If DebugOn Then
        Dim DR As DebugRec
        If Not DL Is Nothing Then
            DL.DebugLog Exe, Module, "MenuItemSelected", "Begin", DR
        End If
    #End If
    
    If vlMenuId = 121 Then Set mvarOps = Nothing
    frmListManagement.DoMenu vlMenuId
    
    #If DebugOn Then
        If Not DL Is Nothing Then
            DL.DebugLog Exe, Module, "MenuItemSelected", "End", DR
        End If
    #End If
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
    
    Dim f As Form
    Dim bLoaded As Boolean
    
    On Error GoTo errhandler
    #If DebugOn Then
        Dim DR As DebugRec
        If Not DL Is Nothing Then
            DL.DebugLog Exe, Module, "Message", "Begin", DR
        End If
    #End If

    Dim lMouse As MousePointerConstants
    
    lMouse = Screen.MousePointer
    Screen.MousePointer = vbHourglass

    Select Case viMessageType
        Case cwiPATIENT_SUMMARY_DRILL_DOWN
            
        Case cwiPRINT_PREVIEW
            
        Case cwiPRINT_REGULAR_VIEW
            
        Case cwiPRINT_PATIENT_SUMMARY_VIEW
        
        Case cwiREFRESH_DATA '{{BOOKMARK Only refresh when appropriate app sends cwiREFRESH_DATA.}}
            'SCR-50350 Added refresh message processing so other module who update the favorites list can get their data displayed
            bLoaded = False
            For Each f In Forms
                If UCase(f.Name) = UCase("frmListManagement") Then
                  bLoaded = True
                  Exit For
                End If
            Next
            If bLoaded Then frmListManagement.RefreshDisplay
        Case cwiABOUT
            '---inform CW that I am showing a modal form---
            Dim objAbout As New CWAbout
            
            cwobjCWComm.Message cwiSHOWING_MODAL_FORM, vbNullString, NCID_CW_APPLICATION, NCID_LIST_MGMT_APPLICATION
            
            Call objAbout.Display(App, frmListManagement)
            Set objAbout = Nothing
            '---inform CW that I am done showing a modal form---
            cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, vbNullString, NCID_CW_APPLICATION, NCID_LIST_MGMT_APPLICATION
            
        Case cwiOK_TO_CHANGE_PATIENT
    
    End Select
    
    Screen.MousePointer = lMouse

    #If DebugOn Then
        If Not DL Is Nothing Then
            DL.DebugLog Exe, Module, "Message", "End", DR
        End If
    #End If

    Exit Sub

errhandler:
    #If DebugOn Then
        If Not DL Is Nothing Then
            DL.DebugLog Exe, Module, "Message", "ErrHandler", DR
        End If
    #End If
    Call gobjShared.ShowVBError(Err.Number, Err.DESCRIPTION, "ListMangement.Message", App.Title, vbExclamation)
    Exit Sub
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

    On Error GoTo errhandler
    #If DebugOn Then
        Dim DR As DebugRec
        If Not DL Is Nothing Then
            DL.DebugLog Exe, Module, "Properties", "Begin", DR
        End If
    #End If

    '--------------------------------------------------------------------------------------
    '  If this method is being called by the CW Designer, you need to perform actions that
    '    the OLE Server would preform in the "Initialize" method and other methods (i.e.
    '    create objects, initialize global variables, etc.).
    '--------------------------------------------------------------------------------------
    If vbCalledByDesigner Then



    End If
    
    #If DebugOn Then
        If Not DL Is Nothing Then
            DL.DebugLog Exe, Module, "Properties", "End", DR
        End If
    #End If

    Exit Sub

errhandler:
    '---show the VB error---
    #If DebugOn Then
        If Not DL Is Nothing Then
            DL.DebugLog Exe, Module, "Properties", "ErrHandler", DR
        End If
    #End If
    Call gobjShared.ShowVBError(Err.Number, Err.DESCRIPTION, "Problems.Properties", App.Title, vbExclamation)

    Exit Sub
End Sub

Private Sub Class_Initialize()
#If DebugOn Then
    Set DL = New cDebugLog
    DL.debugStart
#End If
    #If DebugOn Then
        Dim DR As DebugRec
        If Not DL Is Nothing Then
            DL.DebugLog Exe, Module, "Class_Initialize", "Begin", DR
        End If
    #End If

    '---create OLE objects---
    Set gobjShared = New CWShared
    
    '---set the help file for this application---
    App.HelpFile = gobjShared.CWHelpFile

    #If DebugOn Then
        If Not DL Is Nothing Then
            DL.DebugLog Exe, Module, "Class_Initialize", "End", DR
        End If
    #End If

End Sub

Private Sub Class_Terminate()
    #If DebugOn Then
        Dim DR As DebugRec
        If Not DL Is Nothing Then
            DL.DebugLog Exe, Module, "Class_Terminate", "Begin", DR
        End If
    #End If
    '---release reference to object---
    Set gobjShared = Nothing
    Set cwobjConfig = Nothing
    Set cwobjCWComm = Nothing
    Set cwobjLogonEnviron = Nothing
    
    Set mvarOps = Nothing
    
    #If DebugOn Then
        If Not DL Is Nothing Then
            DL.DebugLog Exe, Module, "Class_Terminate", "End", DR
        End If
    #End If

    #If DebugOn Then
        If Not DL Is Nothing Then
            DL.debugStop
            Set DL = Nothing
        End If
    #End If

End Sub
'**************************************************************************************************************
'  Sub Name: PatientSelected
'
'  Description: This method gets called when a new patient has been selected.
'
'  Parameters:
'    none
'**************************************************************************************************************
Public Sub PatientSelected()

End Sub


'******************************************************************************************
'  Sub Name: Initialize2
'
'  Description: This method first first processes the new user parameter, and then
'  executes initialize with the original 5 parameters.
'  Original 5 Parameters:
'    vlMyProgId - Contains the "ProgId" that CW has assigned to this OLE Server.
'    robjComm - Contains a communications object that is used to handle all communication
'               between this OLE Server and the CW core.
'    robjLogonEnviron - Contains the logon environment object.
'    robjConfig - Contains the configuration object used for getting and setting configuration
'                 information for a specific application.
'    robjPatient - Contains the patient used to do patient searches and selections
'  New Parameter
'    robjUser - Contains the user information that is used for role and security purposes
'*******************************************************************************************
Public Sub ICHCSII_CoreAPI2_Initialize2( _
            ByVal vlMyProgId As Long, robjComm As Object, robjLogonEnvironment As Object, _
            robjConfig As Object, robjPatient As Object, robjSecurity As Object)
    On Error GoTo errhandler
    #If DebugOn Then
        Dim DR As DebugRec
        If Not DL Is Nothing Then
            DL.DebugLog Exe, Module, "ICHCII_CorePPI2_Initialize2", "Begin", DR
        End If
    #End If


    '---create a pointer to objects---
    Set cwobjCWComm = robjComm
    Set cwobjLogonEnviron = robjLogonEnvironment
    Set cwobjConfig = robjConfig
    Set cwobjPatient = robjPatient
    Set gobjUser = robjSecurity

    '---set global flags---
    cwlMyProgId = vlMyProgId

    
    Set mvarOps = Nothing
    Set gobjConfig = cwobjConfig
    Set gobjLogon = cwobjLogonEnviron
    Set gobjComm = cwobjCWComm
    
      
    #If DebugOn Then
        If Not DL Is Nothing Then
            DL.DebugLog Exe, Module, "ICHCII_CorePPI2_Initialize2", "End", DR
        End If
    #End If
    Exit Sub

errhandler:
    #If DebugOn Then
        If Not DL Is Nothing Then
            DL.DebugLog Exe, Module, "ICHCII_CorePPI2_Initialize2", "ErrHandler", DR
        End If
    #End If
    Call gobjShared.ShowVBError(Err.Number, Err.DESCRIPTION, "ListMangement.Initialize2", App.Title, vbExclamation)

    Exit Sub
End Sub
