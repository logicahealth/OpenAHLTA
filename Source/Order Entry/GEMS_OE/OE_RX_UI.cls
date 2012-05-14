VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OE_RX_UI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'**********************************************************
'  V A R I A B L E S     S E C T I O N
'**********************************************************

Private mbCalledAfterInit2 As Boolean
Const c_mszModuleName As String = "OE_RX_UI."

'============================================
Dim mVarPrimaryProviderIen As String

Implements ICHCSII_CoreAPI2
Implements ICHCSII_CoreAPI3
Implements ICHCSII_AnP
             
Private ViewOpened As Boolean

Private Sub ICHCSII_AnP_gridDxClick(Row As Long, Col As Long)
    'n/a
End Sub
Private Property Get ICHCSII_AnP_primaryProviderIEN() As String
    ICHCSII_AnP_primaryProviderIEN = mVarPrimaryProviderIen
End Property
Private Property Get ICHCSII_AnP_ShowAllMeds() As Boolean
    ICHCSII_AnP_ShowAllMeds = False
End Property

Private Property Get ICHCSII_AnP_ShowAllRads() As Boolean
    ICHCSII_AnP_ShowAllRads = False
End Property

Private Property Get ICHCSII_AnP_ShowAllLabs() As Boolean
    ICHCSII_AnP_ShowAllLabs = False
End Property
Private Function ICHCSII_CoreAPI3_CloseView2(ByVal viViewType As Integer) As Boolean

    If gbBackgroundBusy = False Then
        CloseView viViewType
        ICHCSII_CoreAPI3_CloseView2 = True  'We will shut down
    Else
        ICHCSII_CoreAPI3_CloseView2 = False
    End If
End Function

Public Sub ICHCSII_CoreAPI2_Initialize2(ByVal vlMyProgId As Long, robjComm As Object, robjLogonEnvironment As Object, robjConfig As Object, robjPatient As Object, robjUser As Object)
    
    oLogger.LogtoFile c_mszModuleName & "ICHCSII_CoreAPI2_Initialize2", en_Entering, "vlMyProgId", vlMyProgId

    mbCalledAfterInit2 = True
    Call CommonInit2(vlMyProgId, robjComm, robjLogonEnvironment, robjConfig, robjPatient, robjUser)
    Call Initialize(vlMyProgId, robjComm, robjLogonEnvironment, robjConfig, robjPatient)
    
    oLogger.LogtoFile c_mszModuleName & "ICHCSII_CoreAPI2_Initialize2", en_Exiting
    
End Sub

Public Sub Show()
    oLogger.LogtoFile c_mszModuleName & "Show", en_Entering
    Frm_OE_RX.Show
    oLogger.LogtoFile c_mszModuleName & "Show", en_Exiting
End Sub

Private Sub Class_Initialize()

     '---create shared object used to log to file---
    If oLogger Is Nothing Then Set oLogger = New LogtoFile
    oLogger.LogtoFile c_mszModuleName & "Class_Initialize", en_Entering
    
    ViewOpened = NO      '#13736
    Set gobjShared = New CWShared
    
    '---set the help file for this application---
    App.HelpFile = gobjShared.CWHelpFile
    
    oLogger.LogtoFile c_mszModuleName & "Class_Initialize", en_Exiting

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
    oLogger.LogtoFile c_mszModuleName & "QueryView", en_Entering, "viViewType", viViewType

          Select Case viViewType
          Case cwiREGULAR_VIEW
               QueryView = True    '--this view is NOT optional--
          Case cwiPATIENT_SUMMARY_VIEW
               QueryView = False
          Case cwiCLINICIAN_SUMMARY_VIEW
               QueryView = False
          Case cwiENCOUNTER_SUMMARY_VIEW
               QueryView = False
          Case cwiPROPERTY_VIEW
               QueryView = False
          
          Case Else   '--unknown "view type", not supported--
               QueryView = False
          End Select
          
    oLogger.LogtoFile c_mszModuleName & "QueryView", en_Exiting
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
    oLogger.LogtoFile c_mszModuleName & "OpenView", en_Entering, "viViewType", viViewType, "vlParenthWnd", vlParenthWnd, "rhMainForm", rhMainForm, _
                    "cwbViewRunning(enView.Rx)", cwbViewRunning(enView.rx)
    
    '---perform actions based on the "view type"---
    If viViewType = cwiREGULAR_VIEW Or viViewType = cwiANP_VIEW Then
        '---set global flags---
        cwbViewRunning(enView.rx) = True
        cwbFirstRegularResizeRx = True
        SetViewMode viViewType
        Load Frm_OE_RX
        SetParent Frm_OE_RX.hwnd, vlParenthWnd
        mbCalledAfterInit2 = False
        
        '---set parameter---
        rhMainForm = Frm_OE_RX.hwnd
        
    Else
            '---set return value---
        OpenView = False
        oLogger.LogtoFile c_mszModuleName & "OpenView", en_Exiting, "OpenView", OpenView, "cwbViewRunning(enView.Rx)", cwbViewRunning(enView.rx)
        Exit Function
    End If
    
    '---set return value---
    OpenView = True
    ViewOpened = True

    Exit Function
    
ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "OE_RX_UI.OpenView", "OE_RX_UI", vbCritical)
    
    '---set return value---
    OpenView = False
    oLogger.LogtoFile c_mszModuleName & "OpenView", en_Exiting, "OpenView", OpenView, "cwbViewRunning(enView.Rx)", cwbViewRunning(enView.rx)
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

    oLogger.LogtoFile c_mszModuleName & "CloseView", en_Entering, "viViewType", viViewType, _
                  "cwbViewRunning(enView.Rx)", cwbViewRunning(enView.rx)
                  
    If ViewOpened = True Then
        ViewOpened = False
        cwbCloseViewMyselfRx = False
        cwbViewRunning(enView.rx) = False
        Unload Frm_OE_RX
    End If
    Set oAnP = Nothing
    
    oLogger.LogtoFile c_mszModuleName & "CloseView", en_Exiting, _
                "cwbViewRunning(enView.Rx)", cwbViewRunning(enView.rx)
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
    oLogger.LogtoFile c_mszModuleName & "MenuItemSelected", en_Entering, "vlMenuId", vlMenuId
    oLogger.LogtoFile c_mszModuleName & "MenuItemSelected", en_Exiting
    Exit Sub

ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "OE_RX_UI.MenuItemSelected", "OE_RX_UI", vbExclamation)
    oLogger.LogtoFile c_mszModuleName & "MenuItemSelected", en_Exiting
    Exit Sub
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
    oLogger.LogtoFile c_mszModuleName & "ResizeView", en_Entering, "viViewType", viViewType, "vlLeft", vlLeft, "vlTop", vlTop, "vlWidth", vlWidth, "vlHeight", vlHeight
    
     On Error GoTo ErrHandler

     cwlLeftRx = vlLeft
     cwlTopRx = vlTop
     cwlWidthRx = vlWidth
     cwlHeightRx = vlHeight

     SetViewMode viViewType

     If Not cwbFirstRegularResizeRx Then
     
        Frm_OE_RX.WindowState = vbNormal
        
        On Error Resume Next
        If Frm_OE_RX.Visible = True Then
             If UCase$(Frm_OE_RX.ActiveControl.Name) = "LST_ORDERSBOTTOM" Then
                 Frm_OE_RX.TxT_Keyword.SetFocus
             End If
        End If
        On Error GoTo ErrHandler
        
        Frm_OE_RX.Visible = False
        Frm_OE_RX.Move cwlLeftRx, cwlTopRx, cwlWidthRx, cwlHeightRx
        Frm_OE_RX.Visible = True
          
    Else    '--first "regular" resize, need to perform "initailization" procedures--
                               
        cwbFirstRegularResizeRx = False
        Frm_OE_RX.Display
        Call Frm_OE_RX.MatchViewMode
        RefreshOrdersOnPage Frm_OE_RX.Lst_OrdersBottom, hlRX_ORDERS_CODE
    
    End If
    
    Exit Sub
    
ErrHandler:

    '---show the VB error---
    If Not (gobjShared Is Nothing) Then
        Call gobjShared.ShowVBError(Err.Number, Err.Description, "OE_RX_UI.ResizeView", "OE_RX_UI", vbExclamation)
    End If
    Exit Sub
    
End Sub

Private Sub Class_Terminate()
    oLogger.LogtoFile c_mszModuleName & "Class_Terminate", en_Entering
    
    '---release reference to objects---
    Set gobjShared = Nothing
    Set cwobjCWCommRx = Nothing
    Set cwobjLogonEnviron = Nothing
    Set cwobjConfig = Nothing
    Set cwobjPatient = Nothing
    Set oAnP = Nothing
    oLogger.LogtoFile c_mszModuleName & "Class_Terminate", en_Exiting
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
'
Public Sub Message(ByVal viMessageType As Integer, ByVal vsMessageText As String, ByVal vsSourceServerNCID As String)
    
    oLogger.LogtoFile c_mszModuleName & "Message", en_Entering, "viMessageType", viMessageType, "vsMessageText", vsMessageText, "vsSourceServerNCID", vsSourceServerNCID
    '-------------------------------------------------------------------------------
    '  Take corresponding actions based upon the message "type".
    '  NOTE: This method is required, even if no messages were utilized
    '-------------------------------------------------------------------------------
    On Error GoTo ErrHandler

    Select Case viMessageType
        Case cwiREFRESH_DATA:
'            goOESession.RefreshPatientInfo
            If Frm_OE_RX.Visible Then
                Call Frm_OE_RX.Btn_RefreshOrders_Click
            End If
            
        Case cwiOK_TO_CHANGE_PATIENT
            If gbBackgroundBusy = True Then
                 vsMessageText = "N"
            Else
                  vsMessageText = "Y"
            End If
            cwobjCWCommRx.Message cwiOK_TO_CHANGE_PATIENT, vsMessageText, vsSourceServerNCID, NCID_OE_MED_APPLICATION
             
        Case cwiOK_TO_CLOSE_VIEW
            If gbBackgroundBusy = True Then    'we are BUSY and cannot allow closing or patient change
                 vsMessageText = "N"
            Else
                  vsMessageText = "Y"
            End If
            cwobjCWCommRx.Message cwiOK_TO_CLOSE_VIEW, vsMessageText, vsSourceServerNCID, NCID_OE_MED_APPLICATION
             
        Case cwiCHANGE_USER
            If Not (goOESession Is Nothing) Then
                goOESession.Logout
            End If
                    
        Case cwiGENERIC
            If Trim$(vsMessageText = "EXIT_APP") Then
                If AllViewsClosed Then
                    If Not (goOESession Is Nothing) Then
                       goOESession.Logout
                    End If
                    If Not (oLogger Is Nothing) Then
                       Set oLogger = Nothing
                    End If
                    DoEvents
                    Exit Sub
                End If
            End If
               
          Case Else
          End Select
          
          oLogger.LogtoFile c_mszModuleName & "Message", en_Exiting
          Exit Sub
    
ErrHandler:
    '---show the VB error---
          Call gobjShared.ShowVBError(Err.Number, Err.Description, "OE_RX_UI.Message", "OE_RX_UI", vbExclamation)
          oLogger.LogtoFile c_mszModuleName & "Message", en_Exiting
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
Public Sub Initialize(ByVal vlMyProgId As Long, robjComm As Object, robjLogonEnvironment As Object, robjConfig As Object, robjPatient As Object)
    
    On Error GoTo ErrHandler
    oLogger.LogtoFile c_mszModuleName & "Initialize", en_Entering, "vlMyProgId", vlMyProgId, _
                        "typename(robjComm)", TypeName(robjComm), "typename(robjLogonEnvironment)", TypeName(robjLogonEnvironment), _
                         "typename(robjConfig)", TypeName(robjConfig), "typename(robjPatient)", TypeName(robjPatient)
    '---save the "ProgId" assigned to me by the CW Core---
    cwlMyProgIdRx = vlMyProgId
    
    '---create references to passed in objects---
    Set cwobjCWCommRx = robjComm
    Set cwobjLogonEnviron = robjLogonEnvironment
    Set cwobjConfig = robjConfig
    Set cwobjPatient = robjPatient
    
    '---set global flags---
    '****************************************************************************
    If goOESession Is Nothing Then      'testing for nothing not required but good practice
          Set goOEGetSessionMgr = New OE_GetSessionManager
          goOESession = goOEGetSessionMgr.GetSessionManager
          Set goOEGetSessionMgr = Nothing
    End If
    
    oLogger.LogtoFile c_mszModuleName & "Initialize", en_Exiting
    Exit Sub

ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "OE_RX_UI.Initialize", "OE_RX_UI", vbExclamation)
    oLogger.LogtoFile c_mszModuleName & "Initialize", en_Exiting
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
Public Sub Properties(robjLogonEnvironment As Object, robjConfig As Object, ByVal vsOwnerNCID As String, ByVal vbCalledByDesigner As Boolean)
    On Error GoTo ErrHandler
    oLogger.LogtoFile c_mszModuleName & "Properties", en_Entering, _
                        "typename(robjLogonEnvironment)", TypeName(robjLogonEnvironment), _
                        "typename(robjConfig)", TypeName(robjConfig), _
                        "vsOwnerNCID", vsOwnerNCID, "vbCalledByDesigner", vbCalledByDesigner
    
    '*****************************************************************
    '  TO DO:
    '    Add code here to invoke this OLE Server's properties page.
    '*****************************************************************
    


    Exit Sub
    oLogger.LogtoFile c_mszModuleName & "Properties", en_Exiting
ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "OE_RX_UI.Properties", "OE_RX_UI", vbExclamation)
    oLogger.LogtoFile c_mszModuleName & "Properties", en_Exiting
    Exit Sub
End Sub


Public Function ICHCSII_AnP_SetAnPParent(objAnPParent As Object) As Boolean
    CommonSetAnP objAnPParent
End Function

