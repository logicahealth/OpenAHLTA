VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Procedure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Decription:  Interface class used by core to load procedure module.
Option Explicit

Private mfrmProc As frmProcedure       'Form used to select Procedures

#If DEBUGON Then
  Private Const Module = "Procedure.cls"
#End If

Implements ICHCSII_CoreAPI2
Implements ICHCSII_AnP

Private Sub Class_Initialize()

  On Error Resume Next
  
  #If DEBUGON Then
    Set DL = GetObject("", "dlog.cdebuglog")
    DL.debugStart
    Dim DR  As DebugRec
  
    DL.DebugLog Exe, Module, "Class_Initialize", "Begin", DR
  #End If
   
  Set gobjShared = New CWShared
  App.HelpFile = gobjShared.CWHelpFile
  
  #If DEBUGON Then
'    If cwbDebugLogon = False Then Exit Sub
    DL.DebugLog Exe, Module, "Class_Initialize", "End", DR
  #End If

End Sub

Private Sub Class_Terminate()

  On Error Resume Next
  
  #If DEBUGON Then
    Dim DR  As DebugRec
    DL.DebugLog Exe, Module, "Class_Terminate", "Begin", DR
  #End If
   
  Set gobjShared = Nothing
  Set cwobjCWComm = Nothing
  Set cwobjLogonEnviron = Nothing
  Set cwobjConfig = Nothing
  Set cwobjPatient = Nothing
  
  #If DEBUGON Then
    DL.DebugLog Exe, Module, "Class_Terminate", "End", DR
    Set DL = Nothing
  #End If

End Sub

Private Sub ICHCSII_AnP_gridDxClick(row As Long, col As Long)
' not used
End Sub

'Get a reference to AnP Interface object.
Private Function ICHCSII_AnP_SetAnPParent(objAnPParent As Object) As Boolean
  Set gobjAnPParent = objAnPParent
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
   Initialize vlMyProgId, robjComm, robjLogonEnvironment, robjConfig, robjPatient
   
   Exit Sub
   
ErrHandler:
   gobjShared.ShowVBError Err.Number, Err.Description, "Procedure.Initialize_CoreAPI2_Initialize", "Procedure", vbExclamation

End Sub

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
    
  If viViewType = cwiALL_VIEWS Or viViewType = cwiREGULAR_VIEW Then
    If cwbMainViewRunning Then
      cwbMainViewRunning = False
      Unload mfrmProc
      Set mfrmProc = Nothing
      
      Set gobjAnPParent = Nothing
    End If
  End If

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
   
   '-- -save the "ProgId" assigned to me by the CW Core---
   cwlMyProgId = vlMyProgId
   
   '-- -create references to passed in objects---
   Set cwobjCWComm = robjComm
   Set cwobjLogonEnviron = robjLogonEnvironment
   Set cwobjConfig = robjConfig
   Set cwobjPatient = robjPatient
     
   cwbMainViewRunning = False
   
   gbRefreshingTemplates = False       'JT SCR 37028 - use flag to avoid refresh issue
   
   If Len(gobjShared.CmdLineSwitches("GREEN")) > 0 Then
      gbGreenList = True
   Else
      gbGreenList = False
   End If
    
   Exit Sub
   
ErrHandler:
   gobjShared.ShowVBError Err.Number, Err.Description, "Procedure.Initialize", "Procedure", vbExclamation

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
' No menu items
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
  
  Select Case viMessageType
      Case cwiBACKGROUND_APPLICATION
            
        If gbGreenList = True Then
            mfrmProc.ucTemplateNavigation1.PersistHistory
        End If
        
      Case cwiFOREGROUND_APPLICATION
        If gbGreenList = True Then
             mfrmProc.ucTemplateNavigation1.RetrieveHistory
        End If
    
    Case cwiGENERIC
      Select Case vsMessageText
      
         Case "REFRESH TEMPLATES"
        
            gbRefreshingTemplates = True    'JT SCR 37028 - use flag to avoid refresh issue
        
            If gbGreenList = False Then
                If Not mfrmProc.DiagListOpt(0).Value Then
                  mfrmProc.SelectList 0     'scr 31342 Phil Crowder 1/13/02
                Else
                  mfrmProc.load_template_data
                End If
            Else
                  mfrmProc.load_template_favs
            End If
                        
            gbRefreshingTemplates = False    'JT SCR 37028 - use flag to avoid refresh issue
               
      End Select
      
  End Select
  
  Exit Sub
  
ErrHandler:
   gobjShared.ShowVBError Err.Number, Err.Description, "Procedure.Message", "Procedure", vbExclamation
   
   gbRefreshingTemplates = False    'JT SCR 37028 - use flag to avoid refresh issue
   
End Sub

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
   
  If viViewType = cwiREGULAR_VIEW Then
    '-- -set global flags---
    cwbFirstRegularResize = True
    cwbMainViewRunning = True
    
    Set mfrmProc = New frmProcedure
    SetParent mfrmProc.hwnd, vlParenthWnd
    
    '-- -set parameter---
    rhMainForm = mfrmProc.hwnd
    
    OpenView = True
  Else                                '--unsupported view "type"--
    OpenView = False
  End If
  
  Exit Function
  
ErrHandler:
  gobjShared.ShowVBError Err.Number, Err.Description, "Procedure.OpenView", "Procedure", vbCritical
  OpenView = False

End Function

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
      Case Else                        '--unknown "view type", not supported--
         QueryView = False
   End Select

End Function

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
' Core API 1, not used
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
' Core API 1, not used
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
  
  '-- -reset global "client area" parameters---
  cwlLeft = vlLeft
  cwlTop = vlTop
  cwlWidth = vlWidth
  cwlHeight = vlHeight
  
  If viViewType = cwiREGULAR_VIEW Then
    If Not cwbFirstRegularResize Then
      With mfrmProc
        .WindowState = vbNormal
        .AutoRedraw = False
        .Move cwlLeft, cwlTop, cwlWidth, cwlHeight
        .AutoRedraw = False
        .ZOrder
      End With
    Else                             '--first "regular" resize, need to perform "initailization" procedures--
      '-- -reset flag so this code will not be executed again---
      cwbFirstRegularResize = False
      mfrmProc.Display
    End If
  End If
  
  Exit Sub
  
ErrHandler:
  gobjShared.ShowVBError Err.Number, Err.Description, "Procedure.ResizeView", "Procedure", vbExclamation

End Sub
