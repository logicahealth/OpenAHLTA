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
Implements ICHCSII_CoreAPI3      'PR292
Implements ICHCSII_AnP               'PR286
             
Private ViewOpened As Boolean '7-24-01; AAB #13736

Private Sub ICHCSII_AnP_DisplayEncounterLockMessage()

End Sub

Private Property Get ICHCSII_AnP_EncounterApptIEN() As String

End Property

Private Property Get ICHCSII_AnP_EncounterClinicNCID() As String

End Property

Private Sub ICHCSII_AnP_gridDxClick(Row As Long, Col As Long)
    'n/a
End Sub

Private Property Get ICHCSII_AnP_IsEncounterLocked() As Boolean

End Property

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
    On Error Resume Next
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

'Public Property Get bCalledAfterInit2() As Boolean
'    bCalledAfterInit2 = mbCalledAfterInit2
'End Property
'Public Property Let bCalledAfterInit2(ByVal vData As Boolean)
'    mbCalledAfterInit2 = bCalledAfterInit2
'End Property

'Dim oeSearchEngine As OE_SearchEngine
'Public Function SearchRX(szSearchText As String, Optional srcSearchWhat As srcSynonyms = srcSimple) As srcSearchResult
'    Dim Res As srcSearchResult
'    Res = oeSearchEngine.Search(frmOE_RX.cmbRxList, hlRX_ORDERS_CODE, UCase(szSearchText), srcSearchWhat)
'End Function


Public Sub Show()
    oLogger.LogtoFile c_mszModuleName & "Show", en_Entering

    Frm_OE_RX.Show
   ' ShowOrderList Frm_OE_RX
    
    oLogger.LogtoFile c_mszModuleName & "Show", en_Exiting
End Sub

Private Sub Class_Initialize()

     '---create shared object used to log to file---
    If oLogger Is Nothing Then Set oLogger = New LogtoFile
    oLogger.LogtoFile c_mszModuleName & "Class_Initialize", en_Entering
    
 '   Set oeSearchEngine = oOESession.oeSearchEngine
 
          ViewOpened = NO      '#13736

'<<<<< From Template
     '---create shared object used to show VB errors, etc.---
            Set gobjShared = New CWShared
   
    '---set the help file for this application---
            App.HelpFile = gobjShared.CWHelpFile
    
    '------------------------------------
    '  This statement is needed to debug
    '    this OLE Server.
    '------------------------------------
    
    #If ciDEBUG = 1 Then
        If cwbDebugLogon = False Then
            oLogger.LogtoFile c_mszModuleName & "Class_Initialize", en_Exiting
            Exit Sub
        End If
    #End If
    
'<<<<< From Template
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
          
          'new 2-22-01; AAB
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
    
'    'If view already running, then exit with return of false
'    If cwbViewRunning(enView.Rx) Then
'        OpenView = False
'        oLogger.LogtoFile c_mszModuleName & "OpenView", en_Exiting, "OpenView", OpenView, "cwbViewRunning(enView.Rx)", cwbViewRunning(enView.Rx)
'        Exit Function
'    End If
    
    '---perform actions based on the "view type"---
    If viViewType = cwiREGULAR_VIEW Or viViewType = cwiAnP_VIEW Or viViewType = cwiOrderSet_VIEW Then
        '---set global flags---
        cwbViewRunning(enView.rx) = True
        cwbFirstRegularResizeRx = True
        If cwbViewRunning(enView.lab) = False And cwbViewRunning(enView.Rad) = False Then
            goOESession.oeSearchEngine.SearchClinicTableforNCID
        End If
        
        '31315  Order Entry - Appt IEN needs to be sent to CHCS
        If (goAnP Is Nothing) Then
            goOESession.ApptIen = ""
        Else
            goOESession.ApptIen = goAnP.EncounterApptIEN
        End If

        SetViewMode viViewType
        
        Debug.Print "RX handle = " & Frm_OE_RX.hwnd
        Debug.Print "RX client calling handle = " & vlParenthWnd
        
        '''''''''''''''''''''''
        SetParent Frm_OE_RX.hwnd, vlParenthWnd

            If (Not goOESession.RefreshPatientInfo) Then
            'If (Not goOESession.RefreshPatientInfo) And (viViewType <> cwiOrderSet_VIEW) Then ' 02-26-00; AAB
                If mbCalledAfterInit2 = False And bHL7SessionCreated = False Then
                    RefreshEachUIsOrdersOnPage
                ElseIf bHL7SessionCreated = True Then
                    RefreshEachUIsOrdersOnPage
                End If
            End If

         Call DoSplash(SPLASH_UNLOAD)
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
    ViewOpened = True    '7-24-01; AAB #13736

    'Tracker 20191 clear the transaction logs
    If Len(gobjShared.CmdLineSwitches("GREEN")) > 0 Then
        Set goOESession.goRxTransactionCol = Nothing
    
        Set goOESession.goRxTransactionCol = New RxTransactionCollection
    End If
    'goOESession.oeSearchEngine.SearchClinicTableforNCID Frm_OE_RX.Cbo_RxLocation
    oLogger.LogtoFile c_mszModuleName & "OpenView", en_Exiting, "OpenView", OpenView, "cwbViewRunning(enView.Rx)", cwbViewRunning(enView.rx)
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
    On Error Resume Next

    oLogger.LogtoFile c_mszModuleName & "CloseView", en_Entering, "viViewType", viViewType, _
                                   "cwbViewRunning(enView.Rx)", cwbViewRunning(enView.rx)
                                   
    If ViewOpened = True Then     '7-24-01; AAB #13736
        ViewOpened = False       'toggle this one-shot flag
            
        '---this indicates that the CW Core is closing the "view", NOT this OLE Server---
        cwbCloseViewMyselfRx = False
               
        'Only one viewtype will ever be running for each view, this flag ensures that
        cwbViewRunning(enView.rx) = False
               
        '---close the form associate with this view ---
        Unload Frm_OE_RX
               
    End If
    'Tracker 22618
    Set goAnP = Nothing
    
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
    
    '----------------------------------------------------------------------------------
    '  NOTE: You don't have to create menu items, these are optional, that is why
    '        the code is commented out.
    '----------------------------------------------------------------------------------
'    '---perform the actions associated with the menu item that was selected---
'    Select Case vlMenuId
'        Case glOleMenuItem
'            '******************************************************************************
'            '  TO DO:
'            '    Take actions needed to process this menu item.
'            '******************************************************************************
'
'
'
'        Case Else   '--do nothing--
'    End Select
    oLogger.LogtoFile c_mszModuleName & "MenuItemSelected", en_Exiting
    Exit Sub

ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "OE_RX_UI.MenuItemSelected", "OE_RX_UI", vbExclamation)
    oLogger.LogtoFile c_mszModuleName & "MenuItemSelected", en_Exiting
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
'    goOESession.ReleaseAndSelectPatient
'    Frm_OE_RX.ClearRx HARD_RESET
'    RefreshFormCaption hlRX_ORDERS_CODE, Frm_OE_RX, hlACTION_NEW
'    RefreshOrdersOnPage Frm_OE_RX.Lst_OrdersBottom, hlRX_ORDERS_CODE
    oLogger.LogtoFile c_mszModuleName & "PatientSelected", en_Entering

         goOESession.CommonPatientSelected
    oLogger.LogtoFile c_mszModuleName & "PatientSelected", en_Exiting
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

                    '---reset global "client area" parameters---
                         cwlLeftRx = vlLeft
                         cwlTopRx = vlTop
                         cwlWidthRx = vlWidth
                         cwlHeightRx = vlHeight
    
                    'commented out 11-29-99 AKM
                              ''If viViewType = cwiREGULAR_VIEW Then
         
                         SetViewMode viViewType           'added 02-26-00; AAB; moved up 03-14-00; AAB
         
                         If Not cwbFirstRegularResizeRx Then
                              '--------------------------------------------------------------------------------------------------------
                              '  Reset the "WindowState" to "Normal", so the form can be resized.  This is
                              '    needed because a window CANNOT be moved or resized if it is "minimized" or
                              '    "maximized".
                              '
                              '  NOTE: This generates a form "resize" event.
                              '--------------------------------------------------------------------------------------------------------
                                   Frm_OE_RX.WindowState = vbNormal
                                 
                              'Bug fix to allow Frm_OE_Rx.Lst_OrdersBottom to repaint itself
                                   On Error Resume Next
                                         If Frm_OE_RX.Visible = True Then
                                              If UCase$(Frm_OE_RX.ActiveControl.Name) = "LST_ORDERSBOTTOM" Then
                                                  Frm_OE_RX.cmdAction(hlACTION_NEW).SetFocus
                                              End If
                                         End If
                                   On Error GoTo ErrHandler
                 
                              '---hide the form to so that annyoing "screen blinking" will not be seen while controls are moved and resized---
                                   Frm_OE_RX.Visible = False
                                        '---resize form to fill entire "client area", NOTE: This generates a "resize" event---
                                             Frm_OE_RX.Move cwlLeftRx, cwlTopRx, cwlWidthRx, cwlHeightRx
                                                  ''SetViewMode viViewType          'added 02-26-00; AAB; moved up 03-14-00; AAB
                                                  
                                             'amw 4/24/00 moved this to after endif as part of pr213
                                             'Call Frm_OE_RX.MatchViewMode       '03-20-00; AAB
                                             
                                             
                                   Frm_OE_RX.Visible = True
                                   
                         Else    '--first "regular" resize, need to perform "initailization" procedures--
                                
                              
                              '---reset flag so this code will not be executed again---
                                   cwbFirstRegularResizeRx = False
                         
                              '---------------------------------------------------------------------------------------------------
                              '  NOTE: Here is where you "Show" the "Regular View" form for the FIRST time.
                              '---------------------------------------------------------------------------------------------------
                                   Frm_OE_RX.Display
                                   
                                   
                         End If
                         
                        'amw 4/24/00 moved this to after endif as part of pr213
                         Call Frm_OE_RX.MatchViewMode       '03-20-00; AAB
                         
                              ''End If
                                   ''SetViewMode viViewType           'added 02-26-00; AAB; moved up 03-14-00
                        oLogger.LogtoFile c_mszModuleName & "ResizeView", en_Exiting
                         Exit Sub
    
ErrHandler:

                    '---show the VB error---
                         If Not (gobjShared Is Nothing) Then
                              Call gobjShared.ShowVBError(Err.Number, Err.Description, "OE_RX_UI.ResizeView", "OE_RX_UI", vbExclamation)
                         End If
                         oLogger.LogtoFile c_mszModuleName & "ResizeView", en_Exiting
                         Exit Sub
    oLogger.LogtoFile c_mszModuleName & "ResizeView", en_Exiting
End Sub



Private Sub Class_Terminate()
    oLogger.LogtoFile c_mszModuleName & "Class_Terminate", en_Entering
    
    '---release reference to objects---
    Set gobjShared = Nothing
    Set cwobjCWCommRx = Nothing
    Set cwobjLogonEnviron = Nothing
    Set cwobjConfig = Nothing
    Set cwobjPatient = Nothing

    '********************************************************************
    '  TO DO:
    '    Add code here to release reference to all other global objects.
    '********************************************************************
    
    
    
    
    '----------------------------------------
    '  If debugging, make sure to "log off".
    '----------------------------------------
    #If ciDEBUG = 1 Then
        cwbDebugLogOff
    #End If
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
    
'KELLY DEAD 7-27-01 Dim szMsg As String
    
          On Error GoTo ErrHandler

          Select Case viMessageType
          Case cwiREFRESH_DATA:
               goOESession.RefreshPatientInfo
                             
               '======================================================================
               ''Public Const cwiOK_TO_CHANGE_PATIENT As Integer = 42
               ''Public Const cwiOK_TO_CLOSE_VIEW As Integer = 43
                         'Basically we can do either as long as there are no HL7 or other background messages-in-motion
                         'Therefore, use the gbBackgroundBusy flag for control
               
          Case cwiOK_TO_CHANGE_PATIENT
                         If gbBackgroundBusy = True Then
                              vsMessageText = "N"
                         Else
                               vsMessageText = "Y"
                         End If
                         'Added 1-19-01; AAB
                         cwobjCWCommRx.Message cwiOK_TO_CHANGE_PATIENT, vsMessageText, vsSourceServerNCID, NCID_OE_MED_APPLICATION
               
          Case cwiOK_TO_CLOSE_VIEW
               ''if val(vsMessageText) =                 'view type; 0=all
               'Our global flag is universal, not separate, so we say NO even if another type order is busy, as we don't know any better
               
                         If gbBackgroundBusy = True Then    'we are BUSY and cannot allow closing or patient change
                              vsMessageText = "N"
                         Else
                               vsMessageText = "Y"
                         End If
                         'Added 1-19-01; AAB
                         cwobjCWCommRx.Message cwiOK_TO_CLOSE_VIEW, vsMessageText, vsSourceServerNCID, NCID_OE_MED_APPLICATION
               ''End If
               
          Case cwiCHANGE_USER      'Added 6-23-01; AAB
                     If Not (goOESession Is Nothing) Then
                        goOESession.Logout
                    End If
                    
          Case cwiGENERIC
               If Trim$(vsMessageText = "EXIT_APP") Then
                    '1-29-01, AAB; core is closing ... clean house and go away also
                    
                    'If all the views are closed then logout     7-24-01; AAB #13736
     
                         If AllViewsClosed Then
     
                              If Not (goOESession Is Nothing) Then
                                 goOESession.Logout
                                 'Set goOESession = Nothing      '06-13-00; AAB PR#263
                              End If
                              If Not (oLogger Is Nothing) Then
                                 Set oLogger = Nothing      '06-13-00; AAB PR#263
                              End If
                              DoEvents
                              Exit Sub
                         End If

               End If
               

          Case Else   '--do nothing--
               '04-20-00 per USI (see PR#241; DTE #3553) remove all 'safety-net' message-boxes; AAB
                    ''szMsg = "This application does not support this action" & vbCrLf
                    ''szMsg = szMsg & "Message Type: " & viMessageType & " Message Text: " & vsMessageText
                    ''MsgBxARMd szMsg, vbInformation, "OE_RX_UI"
          End Select
          
          oLogger.LogtoFile c_mszModuleName & "Message", en_Exiting
          Exit Sub
    
ErrHandler:
    '---show the VB error---
   'CHCS II GIAT Issues: VB Error Ordering MEDS MN#50687/PCR
   '     Call gobjShared.ShowVBError(Err.Number, Err.Description, "OE_RX_UI.Message", "OE_RX_UI", vbExclamation)
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
          
        'Set goOESession = New OE_SessionManager
        'goOESession.CreateHL7Session 'moved to SessionManager Class Initialize
        
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

