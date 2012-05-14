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
Attribute VB_Name = "clsOrderAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"colWellness"
Attribute VB_Ext_KEY = "Member1" ,"WellnessReminder"
Option Explicit


'**********************************************************
'  V A R I A B L E S     S E C T I O N
'**********************************************************

Implements ICHCSII_CoreAPI2

Public Enum enumOEOrderStatus
    OE_ORDER_INACTIVE = 0
    OE_ORDER_ACTIVE = 1
    OE_ORDER_DELETED = 2
End Enum

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
    Case Else   '--unknown "view type", not supported--
      QueryView = False
  End Select
End Function

Public Function OpenView(ByVal viViewType As Integer, ByVal vlParenthWnd As Long, rhMainForm As Long) As Boolean
    On Error GoTo ErrHandler
   
    '---perform actions based on the "view type"---
    If viViewType = cwiREGULAR_VIEW Then
        
        cwbFirstRegularResize = True
        cwbMainViewRunning = True
        AddMenus
        SetParent frmOrderAdmin.hwnd, vlParenthWnd
        rhMainForm = frmOrderAdmin.hwnd
        
        OpenView = True
   
    Else    '--unsupported view "type"--
        '---set return value---
        OpenView = False
        Exit Function
        
    End If
    
    Exit Function
    
ErrHandler:
    '---show the VB error---
     goBB.SharedObj.ShowVBError Err.Number, Err.Description, "clsOrdeAdmin.OpenView", "clsOrdeAdmin", vbCritical

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
        If cwbMainViewRunning Then
            '---set global flags---
            cwbMainViewRunning = False
            
            '---this indicates that the CW Core is closing the "view", NOT this OLE Server---
            cwbCloseViewMyself = False
            
            '---close ALL forms that are associated with all "views"---
            Unload frmOrderAdmin
        End If
    ElseIf viViewType = cwiREGULAR_VIEW Then
        '---close the "main view"---
        If cwbMainViewRunning Then
            '---set global flags---
            cwbMainViewRunning = False
            
            '---this indicates that the CW Core is closing the "view", NOT this OLE Server---
            cwbCloseViewMyself = False
            
            '---close ALL forms that are associated with all "views"---
            Unload frmOrderAdmin
        End If
    End If
    
    Set frmOrderAdmin = Nothing

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
Public Sub MenuItemSelected(ByVal vlMenuId)
    On Error GoTo ErrHandler
    
    Select Case vlMenuId
        
        Case glMenuNew
            frmOrderAdmin.AddOrder
              
        Case glMenuDelete
            frmOrderAdmin.DeleteOrder
           
        Case glMenuRefresh
            frmOrderAdmin.RefreshFormOrders
            
        Case glMenuImport
            frmOrderAdmin.ImportData
            
        Case glMenuExport
            frmOrderAdmin.ExportData
            
        Case glMenuActivate
            
            frmOrderAdmin.ActivateOtherOrders
            
    End Select
    
    Exit Sub

ErrHandler:
    '---show the VB error---
     goBB.SharedObj.ShowVBError Err.Number, Err.Description, "clsOrdeAdmin.MenuItemSelected", "clsOrdeAdmin", vbExclamation
    
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
    cwlLeft = vlLeft
    cwlTop = vlTop
    cwlWidth = vlWidth
    cwlHeight = vlHeight
    
    '---perform different actions based upon the "view type"---
    If viViewType = cwiREGULAR_VIEW Then
        If goBB.Comm.Message(cwiMODAL_FORM_SHOWING, "", NCID_CW_APPLICATION, NCID_ORDER_ADMIN_APPLICATION) Then
           Exit Sub
        End If
        If Not cwbFirstRegularResize Then
            With frmOrderAdmin
              .WindowState = vbNormal
              .Visible = False
              .Move cwlLeft, cwlTop, cwlWidth, cwlHeight
              .Visible = True
            End With
        Else    '--first "regular" resize, need to perform "initailization" procedures--
            '---reset flag so this code will not be executed again---
            cwbFirstRegularResize = False
        
            If Not frmOrderAdmin.display Then
                Unload frmOrderAdmin
                cwbMainViewRunning = False
                Exit Sub
            End If
            
        End If
    Else
        Screen.MousePointer = vbDefault
        Beep
        MsgBxARMd "Unsupported resize view type: " & CStr(viViewType) & " in clsOrdeAdmin.ResizeView.", vbExclamation, "clsOrdeAdmin"
    End If
    
    Exit Sub
ErrHandler:
     goBB.SharedObj.ShowVBError Err.Number, Err.Description, "clsOrderAdmin.ResizeView", "clsOrdeAdmin", vbExclamation
End Sub

Private Sub Class_Initialize()
    
    If goBB Is Nothing Then Set goBB = New Blackboard
      
    App.HelpFile = goBB.SharedObj.CWHelpFile
    
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
Public Sub Message(ByVal viMessageType As Integer, ByVal vsMessageText As String, ByVal vsSourceServerNCID As String)
    
    On Error GoTo ErrHandler
    
    Select Case viMessageType
        Case cwiREFRESH
            Call frmOrderAdmin.RefreshFormOrders
        Case cwiOK_TO_CLOSE_VIEW
            
            If cwbMainViewRunning Then
                If frmOrderAdmin.SaveOrder Then
                    goBB.Comm.Message cwiOK_TO_CLOSE_VIEW, "Y", vsSourceServerNCID, NCID_ORDER_ADMIN_APPLICATION
                Else
                    goBB.Comm.Message cwiOK_TO_CLOSE_VIEW, "N", vsSourceServerNCID, NCID_ORDER_ADMIN_APPLICATION
                End If
            End If

       
    End Select
    
    Exit Sub
    
ErrHandler:
    '---show the VB error---
     goBB.SharedObj.ShowVBError Err.Number, Err.Description, "clsOrderAdmin.Message", "clsOrdeAdmin", vbExclamation
    
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
    
    '---save the "ProgId" assigned to me by the CW Core---
    cwlMyProgId = vlMyProgId
    
    '---create references to passed in objects---
    Set goBB.Comm = robjComm
    Set goBB.Logon = robjLogonEnvironment
    Set goBB.Config = robjConfig
    
    '---set global flags---
    cwbMainViewRunning = False
           
    '---Instantiate Help-file object
    App.HelpFile = goBB.SharedObj.CWHelpFile
    
    Exit Sub
ErrHandler:
    goBB.SharedObj.ShowVBError Err.Number, Err.Description, "clsOrdeAdmin.Initialize", "clsOrdeAdmin", vbExclamation
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
    
    '--------------------------------------------------------------------------------------
    '  If this method is being called by the CW Designer, you need to perform actions that
    '    the OLE Server would preform in the "Initialize" method and other methods (i.e.
    '    create objects, initialize global variables, etc.).
    '--------------------------------------------------------------------------------------
    If vbCalledByDesigner Then
        '---create pointers to objects---
        Set goBB.Logon = robjLogonEnvironment
        Set goBB.Config = robjConfig
       

    End If  '--if called by the CW Designer--
    
    
    Exit Sub
ErrHandler:
     goBB.SharedObj.ShowVBError Err.Number, Err.Description, "clsOrdeAdmin.Properties", "clsOrdeAdmin", vbExclamation
End Sub




Private Sub Class_Terminate()
    Set goBB = Nothing
End Sub

Public Sub ICHCSII_CoreAPI2_Initialize2(ByVal vlMyProgId As Long, ByRef robjComm As Object, ByRef robjLogonEnvironment As Object, ByRef robjConfig As Object, ByRef robjPatient As Object, ByRef robjUser As Object)
    
    On Error GoTo ErrHandler
    
    '---save the "ProgId" assigned to me by CW---
    cwlMyProgId = vlMyProgId
    
    '---create pointers to objects---
    Set goBB.Comm = robjComm
    Set goBB.Logon = robjLogonEnvironment
    Set goBB.Config = robjConfig
    Set goBB.User = robjUser

    '---set global flags---
    cwbMainViewRunning = False
        
    Exit Sub

ErrHandler:
    goBB.SharedObj.ShowVBError Err.Number, Err.Description, "clsOrdeAdmin.Initialize", "clsOrdeAdmin", vbExclamation

End Sub
