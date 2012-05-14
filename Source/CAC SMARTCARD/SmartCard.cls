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
Attribute VB_Name = "SmartCard"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ICHCSII_CoreAPI2
Implements ICHCSII_CoreAPI3

Private cwlMyProgId As Long 'this is the progId assigned to me by Core
Private oCac As Object

Public Sub Initialize(ByVal ProgId As Long, CommObject As Object, LogonEnvironment As Object, Config As Object, Patient As Object, User As Object)
On Error GoTo ErrorHandler
    
    cwlMyProgId = ProgId '---save the "ProgId" assigned to me by CW---b
    
    Set cwobjCWComm = CommObject
    Set cwobjLogonEnviron = LogonEnvironment
    Set cwobjConfig = Config
    Set cwobjPatient = Patient
    
    Exit Sub

ErrorHandler:
    Call goShared.ShowVBError(Err.Number, Err.Description, "CAC_SCM.Initialize", "Smart Card", vbExclamation)
 End Sub

Public Function OpenView(ByVal ViewType As Integer, ByVal ParenthWnd As Long, ByRef MainFormhWnd As Long) As Boolean
On Error GoTo ErrHandler
    
    If ViewType = cwiREGULAR_VIEW Then
        cw32Global.cwbFirstRegularResize = True
        cw32Global.cwbMainViewRunning = True
        
            Screen.MousePointer = vbDefault
             
            ' If Len(goShared.CmdLineSwitches("CAC")) = 0 Then Exit Sub
            
            cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_CAC_SCM_APPLICATION
            frmEnterPIN.Show vbModal ', frmMain
            cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_CAC_SCM_APPLICATION

            CloseView cwiREGULAR_VIEW
            cwobjCWComm.CloseView cwiREGULAR_VIEW, cwlMyProgId

            If Len(gsXML) Then
             
              Set oCac = cwobjCWComm.InitializeOLEServer(NCID_CAC_OPTIONS_APPLICATION)
              If oCac Is Nothing Then
                MsgBxARMd "Error initializing CAC Object", vbExclamation, "CAC SmartCard"
              Else
                  oCac.StartCACProcessing gsXML
              End If
            End If
        
        
    Else
        Screen.MousePointer = vbDefault
        '*** MsgBox Fix - Code_Validator auto change made 3/21/2003
        MsgBxARMd "Unsupported view type: " & CStr(ViewType) & " in CAC_SCM.OpenView..", vbCritical, "CAC_SCM"
        
        OpenView = False
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    
    'Just need to show a Modal form then close, so setting OpenView = False
    OpenView = False 'True
    
    Screen.MousePointer = vbDefault
    Exit Function
ErrHandler:
    Call goShared.ShowVBError(Err.Number, Err.Description, "CAC_SCM.OpenView", "CAC SCM", vbCritical)
    OpenView = False
    Screen.MousePointer = vbDefault
End Function

Public Sub ResizeView(ByVal ViewType As Integer, ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long)
On Error GoTo ErrorHandler
    
    '---perform different actions based upon the "view type"---
    If ViewType = cwiREGULAR_VIEW Then
    Else
        Screen.MousePointer = vbDefault
        '*** MsgBox Fix - Code_Validator auto change made 3/21/2003
        MsgBxARMd "Unsupported resize view type: " & CStr(ViewType) & " in CAC_SCM.ResizeView.", vbExclamation, "CAC SmartCard"
    End If

    Exit Sub
ErrorHandler:
    Call goShared.ShowVBError(Err.Number, Err.Description, "CAC_SCM.ResizeView", "CAC SmartCard", vbExclamation)
End Sub

Public Sub CloseView(ByVal ViewType As Integer)
    'not implemented - look at closeview2
End Sub

Public Sub Message(ByVal viMessageType As Integer, ByVal vsMessageText As String, ByVal vsSourceServerNCID As String)
'
End Sub

Public Function QueryView(ByVal viViewType As Integer) As Boolean
    Select Case viViewType
        Case cwiREGULAR_VIEW
            QueryView = True    '--this view is NOT optional--
        Case cwiPATIENT_SUMMARY_VIEW
            QueryView = False
        Case cwiPROPERTY_VIEW
            QueryView = False
        Case Else   '--unknown "view type", not supported--
            QueryView = False
    End Select
End Function

'currently this window will not implement any menus
Public Sub MenuItemSelected(ByVal vlMenuId As Long)
On Error GoTo ErrHandler
    
    Screen.MousePointer = vbHourglass
    Select Case vlMenuId
        Case Else
            '---do nothing
    End Select
    Screen.MousePointer = vbDefault
    
    Exit Sub
ErrHandler:
    goShared.ShowVBError Err.Number, Err.Description, "CAC_SCM.MenuItemSelected", "CAC SmartCard", vbExclamation
    Screen.MousePointer = vbDefault
End Sub

Public Sub Properties(LogonEnviron As Object, Config As Object, ByVal OwnerNCID As String, ByVal CalledByDesigner As Boolean)
    'not implemented
End Sub

Private Sub Class_Initialize()
    '---create shared object used to show VB errors, etc.---
    If goShared Is Nothing Then
      Set goShared = New CWShared
    End If
    
    '---set the help file for this application---
    App.HelpFile = goShared.CWHelpFile
End Sub

Private Sub Class_Terminate()
    Set goShared = Nothing
    Set cwobjCWComm = Nothing
    Set cwobjLogonEnviron = Nothing
    Set cwobjConfig = Nothing
    Set cwobjPatient = Nothing
End Sub

Public Sub ICHCSII_CoreAPI2_Initialize2(ByVal ProgId As Long, Comm As Object, LogonEnvironment As Object, Config As Object, Patient As Object, User As Object)
    
    Set goUser = User
    
    Initialize ProgId, Comm, LogonEnvironment, Config, Patient, User
    
End Sub


Private Function ICHCSII_CoreAPI3_CloseView2(ByVal viViewType As Integer) As Boolean
    Dim i As Long
On Error Resume Next

    ICHCSII_CoreAPI3_CloseView2 = True
    
    'RaiseEvent ShutDown
    
    If Err.Number <> 0 Then

    End If

End Function
