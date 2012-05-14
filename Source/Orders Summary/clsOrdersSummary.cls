VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOrdersSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ICHCSII_CoreAPI2
Implements ICHCSII_AnP

Private Sub ICHCSII_AnP_gridDxClick(Row As Long, Col As Long)

End Sub

Private Function ICHCSII_AnP_SetAnPParent(objAnPParent As Object) As Boolean
    
    Set goBB.AnPParent = objAnPParent
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
Private Sub ICHCSII_CoreAPI2_Initialize2(ByVal vlMyProgId As Long, _
                                         ByRef robjComm As Object, _
                                         ByRef robjLogonEnvironment As Object, _
                                         ByRef robjConfig As Object, _
                                         ByRef robjPatient As Object, _
                                         ByRef robjUser As Object)

On Error GoTo ErrHandler
    
    '---Save the "ProgId" assigned to me by the CW Core---
    cwlMyProgId = vlMyProgId
    
    If goBB Is Nothing Then
        Set goBB = New Blackboard
    End If
    
    '---create references to passed in objects---
    Set goBB.Comm = robjComm
    Set goBB.Logon = robjLogonEnvironment
    Set goBB.Config = robjConfig
    Set goBB.Patient = robjPatient
    Set goBB.User = robjUser
    
    App.HelpFile = goBB.SharedObj.CWHelpFile 'SCR-HELP
    
    '---set global flags---
    cwbMainViewRunning = False

    Exit Sub
ErrHandler:
    Call goBB.SharedObj.ShowVBError(Err.Number, Err.Description, "clsTasking.ICHCSII_CoreAPI2_Initialize2", "Orders Summary", vbExclamation)
End Sub


'**************************************************************************************************
'  Function Name: OpenView
'**************************************************************************************************
Public Function OpenView(ByVal viViewType As Integer, ByVal vlParenthWnd As Long, rhMainForm As Long) As Boolean

On Error GoTo ErrHandler
    If CheckUserPrivilege = False Then
        MsgBxARMd "You don't have privilege to read the Inpatient Orders Summary.", vbCritical, "Orders Summary"
        OpenView = False
        Exit Function
    End If

    '---perform actions based on the "view type"---
    
    Select Case viViewType
        Case cwiREGULAR_VIEW
            '---set global flags---
            cwbFirstRegularResize = True
            cwbMainViewRunning = True
            SetParent frmOrdersSummary.hwnd, vlParenthWnd
            '---set parameter---
            rhMainForm = frmOrdersSummary.hwnd
            
        Case Else    '--unsupported view "type"--
            '---set return value---
            OpenView = False
            Screen.MousePointer = vbDefault
            Exit Function
    End Select
        
    '---set return value---
    OpenView = True
    Screen.MousePointer = vbDefault
    Exit Function
ErrHandler:
    Call goBB.SharedObj.ShowVBError(Err.Number, Err.Description, "clsOrdersSummary.OpenView", "Orders Summary", vbCritical)
    OpenView = False
    Screen.MousePointer = vbDefault
End Function


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
    
    'If the app is not running then ignore this call from core..
    If Not cwbMainViewRunning Then Exit Sub          'SCR 1196
    
    '---perform different actions based upon the "view type"---
    If viViewType = cwiREGULAR_VIEW Then
        If Not cwbFirstRegularResize Then
            frmOrdersSummary.WindowState = vbNormal
            frmOrdersSummary.Visible = False
            frmOrdersSummary.Move cwlLeft, cwlTop, cwlWidth, cwlHeight
            frmOrdersSummary.Visible = True
        Else
            '--first "regular" resize, need to perform "initailization" procedures--
            '---reset flag so this code will not be executed again---
            cwbFirstRegularResize = False
        End If
    Else
        Screen.MousePointer = vbDefault
        Beep
        MsgBxARMd "Unsupported resize view type: " & CStr(viViewType) & " in clsOrdersSummary.ResizeView.", vbExclamation, "Problems"
    End If
    
    goBB.RefreshRequired = True
    frmOrdersSummary.display
    
    Exit Sub
ErrHandler:
    Call goBB.SharedObj.ShowVBError(Err.Number, Err.Description, "clsOrdersSummary.ResizeView", "Orders Summary", vbCritical)
    
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
         QueryView = True              '--this view is NOT optional--
'      Case cwiPATIENT_SUMMARY_VIEW
'         QueryView = False
'      Case cwiCLINICIAN_SUMMARY_VIEW
'         QueryView = False
'      Case cwiENCOUNTER_SUMMARY_VIEW
'         QueryView = False
'      Case cwiPROPERTY_VIEW
'         QueryView = True
      Case Else                        '--unknown "view type", not supported--
         QueryView = False
   End Select

    Exit Function
ErrHandler:
    Call goBB.SharedObj.ShowVBError(Err.Number, Err.Description, "clsOrdersSummary.QueryView", "Orders Summary", vbCritical)
End Function

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
'      Case cwiOE_NEW, cwiOE_RENEW, cwiOE_ORDER_SET
'            Call frmReminders.BatchOrderStatus(vsMessageText)
      Case cwiGENERIC
            If goBB.SharedObj.IsAppMode(modeTheater) Then
                Select Case vsMessageText
                    Case "print orders summary"             'SCR 84204
                        Call frmOrdersSummary.PrintList
                    Case "print orders summary - file menu" 'SCR 84207
                        Call frmOrdersSummary.PrintList(True)
                End Select
            End If
            
      Case cwiREFRESH_DATA
'            objForm.RefreshPatientData
'      Case cwiPROPERTIES
            '---allow the user to change the "properties" of this application--
'            frmProperties.Display cwobjLogonEnviron.UserNCID, False
      Case Else
            'Stop
   End Select
   
   Exit Sub
   
ErrHandler:
   '-- -show the VB error---
    Call goBB.SharedObj.ShowVBError(Err.Number, Err.Description, "clsOrdersSummary.Message", "Orders Summary", vbCritical)

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

    If cwbMainViewRunning Then
    
        cwbMainViewRunning = False
        
        '---this indicates that the CW Core is closing the "view", NOT this OLE Server---
        cwbCloseViewMyself = False
        
        Unload frmOrdersSummary
        
        '---SetPriority Data
'        If frmTasking.CloseForm() Then
'            '---close ALL forms that are associated with all "views"---
'            Unload frmTasking
'        Else
'            ICHCSII_CoreAPI3_CloseView2 = False
'        End If
        
    End If
    
    Exit Sub

ErrHandler:
   '-- -show the VB error---
    Call goBB.SharedObj.ShowVBError(Err.Number, Err.Description, "clsOrdersSummary.Message", "Orders Summary", vbCritical)
End Sub

'Private Function M_bSubscribeForMessages(ByVal bSubscribe As Boolean) As Boolean
''''************************************************************************
''''Called by:     Initialize
''''               Class_Terminate
''''               OpenView
''''               CloseView
''''Arguments:
'''' bMed          Request to remove all queued Med Items
''''
''''Description:   Based on the flag request or cancel required messages
''''
''''
''''Note:
''''Uses Globals: None
''''
''''Uses Modules:
''''
''''Returns:       True  - If successful
''''               False - else
''''
''''Programmer:    Alexander Alberstein
''''
''''************************************************************************
'   Const sPROC_NAME  As String = "M_bSubscribeForMessages"
'
'   Dim bResult       As Boolean
'
'On Error GoTo ErrHnd
'
'   If Not cwobjCWComm Is Nothing Then
'      bResult = M_bSubscribeForMsg(bSubscribe, hlACTION_NEW)
'      bResult = M_bSubscribeForMsg(bSubscribe, hlACTION_ORDER_SET) And bResult
'   End If
'
'   M_bSubscribeForMessages = bResult
'
'Exit Function
'ErrHnd:
'   M_bSubscribeForMessages = False
'   Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")
'
'End Function
'
'Private Function M_bSubscribeForMsg _
'                (ByVal bSubscribe As Boolean, ByVal lMessageType As Long) As Boolean
''''************************************************************************
''''Called by:     M_bSubscribeForMessages
''''
''''Arguments:
'''' bMed          Request to remove all queued Med Items
''''
''''Description:   Based on the flag request or cancel specified message
''''               from Med ,Rad and Lab modules
''''
''''Note:
''''Uses Globals: None
''''
''''Uses Modules:
''''
''''Returns:       True  - If successful
''''               False - else
''''
''''Programmer:    Alexander Alberstein
''''
''''************************************************************************
'   Const sPROC_NAME  As String = "M_bSubscribeForMsg"
'
'On Error GoTo ErrHnd
'
'
'   With cwobjCWComm
'    If bSubscribe Then
'       .AddListenerForMessage Me, NCID_OE_LAB_APPLICATION, lMessageType, NCID_ORDER_SETS_APPLICATION
'       .AddListenerForMessage Me, NCID_OE_RAD_APPLICATION, lMessageType, NCID_ORDER_SETS_APPLICATION
'       .AddListenerForMessage Me, NCID_OE_MED_APPLICATION, lMessageType, NCID_ORDER_SETS_APPLICATION
'    Else
'       .RemoveListenerForMessage NCID_OE_LAB_APPLICATION, lMessageType, NCID_ORDER_SETS_APPLICATION
'       .RemoveListenerForMessage NCID_OE_RAD_APPLICATION, lMessageType, NCID_ORDER_SETS_APPLICATION
'       .RemoveListenerForMessage NCID_OE_MED_APPLICATION, lMessageType, NCID_ORDER_SETS_APPLICATION
'    End If
'   End With
'
'   M_bSubscribeForMsg = True
'
'Exit Function
'ErrHnd:
'   M_bSubscribeForMsg = False
'   Call LibApp_Error _
'      (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsMESSAGE & " " & lMessageType)
'
'End Function
'

Public Sub ModifyOtherOrder(ByVal oInpatOrder As DTOs.InPatientOrder)
    Call frmOrdersSummary.ctrlInPatientOrder_Modify(oInpatOrder)
End Sub

Private Function CheckUserPrivilege() As Boolean
On Error GoTo ErrHandler

    CheckUserPrivilege = False

    If goBB.User.HasPrivilegeEx(Priv_Orders_Laboratory_Orders, Priv_Read) Or _
        goBB.User.HasPrivilegeEx(Priv_Other_Therapies, Priv_Read) Or _
        goBB.User.HasPrivilegeEx(Priv_Orders_Pharmacy_Orders, Priv_Read) Or _
        goBB.User.HasPrivilegeEx(Priv_Orders_Radiology_Orders, Priv_Read) Or _
        goBB.User.HasPrivilegeEx(Priv_Procedure, Priv_Read) Or _
        goBB.User.HasPrivilegeEx(Priv_Order_Sets_Clinic_Order_Set, Priv_Read) Or _
        goBB.User.HasPrivilegeEx(Priv_Diagnosis, Priv_Read) Then
        CheckUserPrivilege = True
    End If

    Exit Function
    
ErrHandler:
  goBB.SharedObj.ShowVBError Err.Number, Err.Description, ".CheckUserPrivilege", "Orders Summary", vbExclamation

End Function


