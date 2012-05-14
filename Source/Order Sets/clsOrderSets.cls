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
Attribute VB_Name = "OrderSets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''************************************************************************
'''Module:           OrderSets/clsOrderSets.cls
'''
'''Description:      Interface to CHCSII - CHCSII Core API compliant OLE server
'''
'''
'''Modules Required:
'''
'''Author         :  Alexander Alberstein
'''Modified       :
'''               :
'''************************************************************************
Option Explicit

Private Const msMODULE_NAME   As String = "OrderSets."

Private mobjOrderSets   As New clsOrdSets

Implements ICHCSII_CoreAPI2
Implements ICHCSII_User
Implements ICHCSII_AnP
Private Function M_bSubscribeForMessages(ByVal bSubscribe As Boolean) As Boolean
'''************************************************************************
'''Called by:     Initialize
'''               Class_Terminate
'''               OpenView
'''               CloseView
'''Arguments:
''' bMed          Request to remove all queued Med Items
'''
'''Description:   Based on the flag request or cancel required messages
'''
'''
'''Note:
'''Uses Globals: None
'''
'''Uses Modules:
'''
'''Returns:       True  - If successful
'''               False - else
'''
'''Programmer:    Alexander Alberstein
'''
'''************************************************************************
   Const sPROC_NAME  As String = "M_bSubscribeForMessages"
   
   Dim bResult       As Boolean
   
On Error GoTo ErrHnd

   If Not cwobjCWComm Is Nothing Then
      bResult = M_bSubscribeForMsg(bSubscribe, hlACTION_NEW)
      bResult = M_bSubscribeForMsg(bSubscribe, hlACTION_ORDER_SET) And bResult
   End If
   
   M_bSubscribeForMessages = bResult
   
Exit Function
ErrHnd:
   M_bSubscribeForMessages = False
   Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")
   
End Function

Private Function M_bSubscribeForMsg _
                (ByVal bSubscribe As Boolean, ByVal lMessageType As Long) As Boolean
'''************************************************************************
'''Called by:     M_bSubscribeForMessages
'''
'''Arguments:
''' bMed          Request to remove all queued Med Items
'''
'''Description:   Based on the flag request or cancel specified message
'''               from Med ,Rad and Lab modules
'''
'''Note:
'''Uses Globals: None
'''
'''Uses Modules:
'''
'''Returns:       True  - If successful
'''               False - else
'''
'''Programmer:    Alexander Alberstein
'''
'''************************************************************************
   Const sPROC_NAME  As String = "M_bSubscribeForMsg"
   
On Error GoTo ErrHnd


   With cwobjCWComm
    If bSubscribe Then
       .AddListenerForMessage Me, NCID_OE_LAB_APPLICATION, lMessageType, NCID_ORDER_SETS_APPLICATION
       .AddListenerForMessage Me, NCID_OE_RAD_APPLICATION, lMessageType, NCID_ORDER_SETS_APPLICATION
       .AddListenerForMessage Me, NCID_OE_MED_APPLICATION, lMessageType, NCID_ORDER_SETS_APPLICATION
    Else
       .RemoveListenerForMessage NCID_OE_LAB_APPLICATION, lMessageType, NCID_ORDER_SETS_APPLICATION
       .RemoveListenerForMessage NCID_OE_RAD_APPLICATION, lMessageType, NCID_ORDER_SETS_APPLICATION
       .RemoveListenerForMessage NCID_OE_MED_APPLICATION, lMessageType, NCID_ORDER_SETS_APPLICATION
    End If
   End With
   
   M_bSubscribeForMsg = True
   
Exit Function
ErrHnd:
   M_bSubscribeForMsg = False
   Call LibApp_Error _
      (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsMESSAGE & " " & lMessageType)

End Function

Private Sub ProcessCommandLine()
'''************************************************************************
'''Called by:     Initialize
'''
'''Arguments:
'''
'''Description:   Complete initialization by processing current command line switches
'''
'''Note:
'''Uses Globals:
'''Uses Modules:  cwobjCWComm
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''
'''************************************************************************
   Const sPROC_NAME  As String = "ProcessCommandLine"
   
   Dim sSwithes      As String
   
On Error GoTo ErrHnd

   '''ERROR LOGS
   sSwithes = UCase$(cwobjCWComm.CmdLineSwitch("ERRLOG"))
   '''Enable or disable Order Sets error log
   pbErrLog = InStr(1, sSwithes, "ORDERSETS") > 0
   '''Enable or disable global error log
   pbErrLogGlobal = InStr(1, sSwithes, "GLOBAL") > 0
   
     If Len(gobjShared.CmdLineSwitches("GREEN")) > 0 Then
        gbGreenList = True
     Else
        gbGreenList = False
     End If
   
Exit Sub
ErrHnd:
   Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")

End Sub

Public Function Mtd_bPushOrder _
      (objType As Object, ByVal sItemKey As String, Optional ByRef sErrMsg As String) As Boolean
'''************************************************************************
'''Called by:     Any
'''
'''Arguments:
''' objType       Order Item(type Rx, Lab or Med) requested to be added to the queue
''' sItemKey      Key of the specified queue Item - used for Mtd_bRemoveOrder
''' sErrMsg       Error information if any - to be returned byref
'''
'''Description:   If specified object is one of the valid type and queue key is specified
'''                  Add it to the corresponding queue
'''                  If Main form is loaded
'''                     Add items to the grid to allow the user to
'''                     see it as soon as order sets tab is activated.
'''                     By default - select this item for submition
'''               Else
'''                  Report an error
'''                  Return false
'''
'''Note:
'''Uses Globals: None
'''
'''Uses Modules:
'''
'''Returns:       True  - If successful
'''               False - else
'''
'''Programmer:    Alexander Alberstein
'''
'''************************************************************************
   Const sPROC_NAME  As String = "Mtd_bPushOrder"
   
On Error GoTo ErrHnd
   
   Mtd_bPushOrder = mobjOrderSets.P_bAddOrdItem(objType, sItemKey, True, sErrMsg)
   
Exit Function
ErrHnd:
   Mtd_bPushOrder = False
   sErrMsg = Err.Description
   Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsKEY & ": " & sItemKey)

End Function

Public Function Mtd_bSubmitAll(ByRef sErrMsg As String) As Boolean
'''************************************************************************
'''Called by: Any
'''
'''Arguments:
''' sErrMsg       Error information if any - to be returned byref
'''
'''Description:   If Order sets form is up
'''                  Submit all items for processing
'''               Else
'''                  Report an error and return false
'''Note:
'''Uses Globals: None
'''
'''Uses Modules:
'''
'''Returns:       True  - If successful
'''               False - else
'''
'''Programmer:    Alexander Alberstein
'''
'''************************************************************************
   Const sPROC_NAME  As String = "Mtd_bSubmitAll"
   
On Error GoTo ErrHnd
   
   Mtd_bSubmitAll = mobjOrderSets.P_bSubmitOrders(sErrMsg)
   
Exit Function
ErrHnd:
   Mtd_bSubmitAll = False
   sErrMsg = Err.Description
   Call LibApp_Error(Err.Number, sErrMsg, msMODULE_NAME, sPROC_NAME, "")

End Function

Public Function QueryView(ByVal viViewType As Integer) As Boolean
'''*******************************************************************************************
'''  Function Name: QueryView
'''
'''  Description: This method is called when a client wants to know if this OLE Server
'''               supports a specific "view type".
'''
'''  Parameters:
'''    viViewType - Contains the "type" of view being queried for.
'''
'''  Return Value:
'''    Returns TRUE if this OLE Server supports the passed in "view type".
'''    Returns FALSE if this OLE Server does not support the passed in "view type".
'*******************************************************************************************
   Select Case viViewType
      Case cwiREGULAR_VIEW
         QueryView = True              '''--this view is NOT optional--
      Case cwiPATIENT_SUMMARY_VIEW
         QueryView = False
      Case cwiCLINICIAN_SUMMARY_VIEW
         QueryView = False
      Case cwiENCOUNTER_SUMMARY_VIEW
         QueryView = False
      Case Else                        '''--unknown "view type", not supported--
         QueryView = False
   End Select

End Function

Public Function OpenView _
(ByVal viViewType As Integer, ByVal vlParenthWnd As Long, rhMainForm As Long) _
 As Boolean
'''**************************************************************************************************
'''  Function Name: OpenView
'''
'''  Description: This method is called when the CW Core attempts to open a "view" provided by this
'''               OLE server.
'''
'''  Parameters:
'''    viViewType - Contains the "type" of view being requested by the CW Core to open.
'''    vlParenthWnd - Contains the hWnd of the client area window to be used by this OLE Server.
'''    rhMainForm - If "viViewType" is a "regular view", upon return this will contain
'''                 the window "handle" of the main "regular" form for this OLE Server.
'''
'''  Return Value:
'''    Returns TRUE if the "view" was opened successfully.
'''    Returns FALSE if an error occurred openning the "view".
'''**************************************************************************************************
   Const sPROC_NAME     As String = "OpenView"
   
   Dim bResult          As Boolean
   
On Error GoTo ErrHnd
   
   '-- -perform actions based on the "view type"---
   If viViewType = cwiREGULAR_VIEW Then
   
      '-- -set global flags---
      cwbFirstRegularResize = True
      
      If M_bSubscribeForMessages(True) Then
         bResult = mobjOrderSets.P_bOpen(vlParenthWnd, rhMainForm)
      End If
      '''Else
      ''''--unsupported view "type"--
      
   
   End If

   OpenView = bResult
   
Exit Function
ErrHnd:
   OpenView = False
   Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")
   
End Function

Public Sub CloseView(ByVal viViewType As Integer)
'*******************************************************************************************
'  Sub Name: CloseView
'
'  Description: This method is called when the CW Core closes a "view" provided by this OLE
'               Server.
'
'  Parameters:
'    viViewType - Contains the "type" of view that the CW Core is closing.
'*******************************************************************************************
   Const sPROC_NAME     As String = "CloseView"
   
On Error GoTo ErrHnd

   If viViewType = cwiALL_VIEWS Or viViewType = cwiREGULAR_VIEW Then
      Call M_bSubscribeForMessages(False)
      Call mobjOrderSets.P_bClose
   End If
   
   Set gobjAnPParent = Nothing
   
Exit Sub
ErrHnd:
   Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")

End Sub

Public Sub MenuItemSelected(ByVal vlMenuId As Long)
'''*******************************************************************************************
'''  Sub Name: MenuItemSelected
'''
'''  Description: This method is called when the user has selected a menu item (from the CW
'''               Core's main menu), which is "owned" by this OLE Server.
'''
'''               NOTE: This is an OPTIONAL method.
'''
'''  Parameters:
'''    vlMenuId - Contains the menu id of the menu item the user selected.
'*******************************************************************************************
   Const sPROC_NAME     As String = "MenuItemSelected"
   
On Error GoTo ErrHnd
   
   
Exit Sub
ErrHnd:
   Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")
   
End Sub

Public Sub PatientSelected()
'''*****************************************************************************
'''  Sub Name: PatientSelected
'''
'''  Description: This method gets called when a new patient has been selected from the CW Core.
'''
'''               NOTE: This is an OPTIONAL method.
'''
'''  Parameters:
'''    none
'''*****************************************************************************
   Const sPROC_NAME     As String = "PatientSelected"

On Error GoTo ErrHnd
   
   Call mobjOrderSets.P_bClose
   
Exit Sub
ErrHnd:
   Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")

End Sub

Public Sub ResizeView(ByVal viViewType As Integer, _
                       ByVal vlLeft As Long, _
                       ByVal vlTop As Long, _
                       ByVal vlWidth As Long, _
                       ByVal vlHeight As Long)
'''*******************************************************************************************
'''  Sub Name: ResizeView
'''
'''  Description: This method is called when a change has occurred in the "client area"
'''              used by this OLE Server, resulting in a need to resize a particular "view".
'''
'''  Parameters:
'''    viViewType - Contains the "type" of view being requested to resize by the CW Core.
'''    vlLeft - Contains the "left" coordinate for the OLE server.
'''    vlTop - Contains the "top" coordinate for the OLE server.
'''    vlWidth - Contains the "width" for the OLE server.
'''    vlHeight - Contains the "height" for the OLE server.
'''*******************************************************************************************
   Const sPROC_NAME     As String = "ResizeView"

On Error GoTo ErrHnd
   
   If viViewType = cwiREGULAR_VIEW Then
   
      If Not cwbFirstRegularResize Then
         
         With mobjOrderSets.OrderSetsForm
         
            '''Avoid screen flickering
            .Visible = False
            
            '''  Make sure that "WindowState" is "Normal", so the form can be resized.  This is
            '''    needed because a window CANNOT be moved or resized if it is "minimized" or "maximized".
            .WindowState = vbNormal
                                                                                          
            '''Resize form to fill entire "client area"
            ''''''NOTE: This generates a "resize" event
            .Move vlLeft, vlTop, vlWidth, vlHeight
         
            .Visible = True
            
         End With
         
      Else
         '-- -reset flag so this code will not be executed again---
         cwbFirstRegularResize = False
      
         Call mobjOrderSets.P_ResizeView(vlLeft, vlTop, vlWidth, vlHeight, "")
      End If
      
   End If
   
Exit Sub
ErrHnd:
   Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")
   
End Sub

Private Sub Class_Initialize()

   Const sPROC_NAME     As String = "Class_Initialize"
   
   Dim objDLog          As New mclsLibDebug
   
On Error GoTo ErrHnd

'#If DEBUGON Then
'   Set DL = GetObject("", "dlog.cdebuglog")
'   DL.debugStart
'   Dim DR  As DebugRec
'
'   DL.DebugLog Exe, msMODULE, "Class_Initialize", "Begin", DR
'#End If
   Call objDLog.Mtd_bLogMessage(msMODULE_NAME, sPROC_NAME)
   '-- -create shared object used to show VB errors, etc.---
   Set gobjShared = New CWShared
   '-- -set the help file for this application---
   App.HelpFile = gobjShared.CWHelpFile

   psExeName = App.EXEName
   psErrLogName = "OrdSetErr.txt"
   psErrLogGlobalName = "CHError.log"
   pbErrLogGlobal = True
   Call LibApp_bSetClientProps(App, NCID_ORDER_SETS_APPLICATION, Nothing)
   

'#If ciDEBUG = 1 Then
'   If cwbDebugLogon = False Then
'      Exit Sub
'   End If
'#End If
'
'#If DEBUGON Then
'   DL.DebugLog Exe, msMODULE, "Class_Initialize", "End", DR
'#End If
'
'#If coiDebug Then
'     gCOI_DL.debugStart
'#End If

Exit Sub
ErrHnd:
   Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", True)
   
End Sub

Private Sub Class_Terminate()

On Error Resume Next

   Call M_bSubscribeForMessages(False)
   Set gobjShared = Nothing
   Set cwobjCWComm = Nothing
   Set cwobjLogonEnviron = Nothing
   Set cwobjConfig = Nothing
   Set cwobjPatient = Nothing
   Set gobjUser = Nothing

   Set mobjOrderSets = Nothing
   Set gobjAnPParent = Nothing

End Sub

Public Sub Message _
(ByVal viMessageType As Integer, ByVal vsMessageText As String, ByVal vsSourceServerNCID As String)
'''************************************************************************
'''Called By:           any
'''
'''Arguments:
''' viMessageType       Contains the "type" of message to process.
''' vsMessageText       Contains any "data" assoicated with this message.
''' vsSourceServerNCID  Contains the NCID of the OLE Server that "sent" the message.
'''
'''  Description:       This method is used by the CW Core to send messages to this OLE Server.
'''                     The messages can be from the CW Core or from another OLE Server.
'''
'''Note:
'''Warning:
'''Updates Globals:
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************

   On Error GoTo ErrHnd
   
   Const sPROC_NAME     As String = "Message"
   
   Select Case viMessageType
      Case cwiBACKGROUND_APPLICATION
      
        If gbGreenList = True Then
            mobjOrderSets.P_bPersistHistory
        End If
        
      Case cwiFOREGROUND_APPLICATION
        If gbGreenList = True Then
            mobjOrderSets.P_bRetrieveHistory
        End If
      
      '''hlACTION_NEW Or hlACTION_RENEW or hlACTION_ORDER_SET
      Case cwiOE_NEW, cwiOE_RENEW, cwiOE_ORDER_SET
         Call mobjOrderSets.P_OrderStatus(vsMessageText)
         
      Case cwiGENERIC
        Select Case vsMessageText
        
        'Refresh template data if new template is chosen SCR-23260
        Case "REFRESH TEMPLATES"
        If gbGreenList = False Then
            If mobjOrderSets.P_bLoadTemplateItems Then
               mobjOrderSets.PopulateOrdersGrid
            End If
        Else
            If mobjOrderSets.P_bLoad_Template_Favs Then
               mobjOrderSets.PopulateOrdersGrid
            End If
            mobjOrderSets.PopulateStandaloneGrid           'SCRs 44213 and 44214
        End If
      End Select
         
   End Select
     
Exit Sub
ErrHnd:
   Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, vsMessageText)

End Sub

Public Sub Initialize(ByVal vlMyProgId As Long, _
                       robjComm As Object, _
                       robjLogonEnvironment As Object, _
                       robjConfig As Object, _
                       robjPatient As Object)
'''****************************************************************************************************
'''  Sub Name: Initialize
'''
'''  Description: This method is called when the CW Core creates an "instance" of this OLE Server.  It
'''               is used to create and initialize global variables that are shared across all
'''               "views" supported by this OLE Server, i.e. variables that only need to be
'''               initialized ONCE, upon object creation.
'''
'''  Parameters:
'''    vlMyProgId - Contains the "ProgId" that the CW Core has assigned to this OLE Server.
'''    robjComm - Contains a communications object that is used to handle all communication
'''               between this OLE Server and the CW Core.
'''    robjLogonEnvironment - Contains the logon environment object.  Contains information such
'''                           as who logged on, where they logged on, etc.
'''    robjConfig - Contains the configuration object that is used to get and set configuration
'''                 information for a specific application.
'''    robjPatient - Contains the patient object containing the currently selected patient.
'''****************************************************************************************************
   Const sPROC_NAME     As String = "Initialize"
   
On Error GoTo ErrHnd

   '---Instantiate Help-file object
   App.HelpFile = gobjShared.CWHelpFile
   '-- -save the "ProgId" assigned to me by the CW Core---
   cwlMyProgId = vlMyProgId
   
   '-- -create references to passed in objects---
   Set cwobjCWComm = robjComm
   Set cwobjLogonEnviron = robjLogonEnvironment
   Set cwobjConfig = robjConfig
   Set cwobjPatient = robjPatient
   Call LibApp_bSetClientProps(App, NCID_ORDER_SETS_APPLICATION, cwobjCWComm)
   
   Call M_bSubscribeForMessages(True)

   Call ProcessCommandLine
   
Exit Sub
ErrHnd:
   '-- -show the VB error---
   Call gobjShared.ShowVBError _
   (Err.Number, Err.Description, msMODULE_NAME & "Initialize", msMODULE_NAME, vbExclamation)

End Sub

Private Sub ICHCSII_AnP_gridDxClick(row As Long, Col As Long)
   
End Sub

Public Sub ICHCSII_CoreAPI2_Initialize2(ByVal vlMyProgId As Long, _
                                        robjComm As Object, _
                                        robjLogonEnvironment As Object, _
                                        robjConfig As Object, _
                                        robjPatient As Object, _
                                        robjUser As Object)
'''******************************************************************************************
'''  Sub Name: Initialize2
'''
'''  Description: This method is called when CW "creates" an instance of this OLE Server.  It
'''               is used to create and initialize global variables that are shared across all
'''               "views" supported by this OLE Server, i.e. variables that only need to be
'''               initialized ONCE, upon object creation.
'''
'''  Parameters:
'''    vlMyProgId - Contains the "ProgId" that CW has assigned to this OLE Server.
'''    robjComm - Contains a communications object that is used to handle all communication
'''               between this OLE Server and the CW core.
'''    robjLogonEnviron - Contains the logon environment object.
'''    robjConfig - Contains the configuration object used for getting and setting configuration
'''                 information for a specific application.
'''    robjPatient - Contains the patient used to do patient searches and selections
'''*******************************************************************************************
   Const sPROC_NAME     As String = "ICHCSII_CoreAPI2_Initialize2"

On Error GoTo ErrHnd
   
   Set gobjUser = robjUser
   
   Dim objRoles      As ICHCSII_Roles
   Dim objResources  As ICHCSII_Resources
   
   Call Initialize(ByVal vlMyProgId, robjComm, robjLogonEnvironment, robjConfig, robjPatient)
   
Exit Sub
ErrHnd:
   Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")

End Sub

Public Sub Properties(robjLogonEnvironment As Object, _
                       robjConfig As Object, _
                       ByVal vsOwnerNCID As String, _
                       ByVal vbCalledByDesigner As Boolean)
'''*************************************************************************************************
'''  Sub Name: Properties
'''
'''  Description: This method is called by a client when they want to modify the "properties", or the
'''               configuration of this OLE Server.
'''
'''               NOTE: This is an OPTIONAL method.
'''
'''  Parameters:
'''    robjLogonEnvironment - Contains the logon environment object.
'''    robjConfig - Contains the configuration object that is used to get and set configuration
'''                 information for a specific application.
'''    vsOwnerNCID - Contains the NCID of the "owner" to use in storing modifications to the
'''                  configuration.
'''    vbCalledByDesigner - TRUE if this method is being called by the CW Designer, FALSE otherwise.
'''*****************************************************************************************************
   Const sPROC_NAME     As String = "Properties"
   
On Error GoTo ErrHnd
   
   '-- *****************************************************************
   '--  TO DO:
   '--    Add code here to invoke this OLE Server's properties page.
   '-- *****************************************************************
   
Exit Sub
ErrHnd:
   Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")

End Sub

Private Function ICHCSII_AnP_SetAnPParent(objAnPParent As Object) As Boolean

   Set gobjAnPParent = objAnPParent

End Function

Private Function ICHCSII_User_CACPresent() As Boolean

End Function

Private Function ICHCSII_User_CACSignText(ByVal PIN As String, ByVal Text As String, Signature As String, Optional ErrCode As Long, Optional ErrMsg As String) As Boolean

End Function

Private Function ICHCSII_User_CACValidatePIN(ByVal PIN As String) As Boolean

End Function

Private Function ICHCSII_User_GetAttribute(AttributeName As String) As String
   
End Function

Private Function ICHCSII_User_GetAttributeNames() As Collection
   
End Function

Private Function ICHCSII_User_HasPrivilege(NCID As String) As Boolean
   
End Function

Private Function ICHCSII_User_HasPrivilegeEx(PrivilegeArea As String, _
                                              PrivilegeType As CHCSII_ClientSecurity.CHCSII_PrivilegeType, _
                                              Optional OtherType As Integer) _
                                              As Boolean
   
End Function

Private Property Get ICHCSII_User_LoginMode() As CHCSII_ClientSecurity.CHCSII_LoginMode

End Property

Private Property Get ICHCSII_User_Resources() As CHCSII_ClientSecurity.ICHCSII_Resources
   
End Property

Private Property Get ICHCSII_User_Roles() As CHCSII_ClientSecurity.ICHCSII_Roles
   
End Property

Private Property Get ICHCSII_User_UserId() As String
   
End Property

Private Property Get ICHCSII_User_UserName() As String
   
End Property

Private Function ICHCSII_User_ValidatePassword(Password As String) As Boolean
   
End Function

Public Function Mtd_bRemoveOrder _
   (lType As enumItemType, ByVal sItemKey As String, Optional sErrMsg As String) As Boolean
'''************************************************************************
'''Called by:     Any
'''
'''Arguments:
''' lType         Order Type identifier (meds, rads, labs)
''' sItemKey      Key of the specified queue Item - used for Mtd_bRemoveOrder
''' sErrMsg       Error info if any - to be returned byref
'''
'''Description:   If Valid Item type and Item Key are specified
'''                  Remove specified Order Item from the queue
'''                  If main form is loaded
'''                     Remove this item from the grid
'''               else
'''                  Report an error and return false
'''Note:
'''Uses Globals:
'''Uses Modules:
'''
'''Returns:       True  - If successful
'''               False - else
'''
'''Programmer:    Alexander Alberstein
'''
'''************************************************************************
   Const sPROC_NAME  As String = "Mtd_bRemoveOrder"
   
On Error GoTo ErrHnd
   
   Mtd_bRemoveOrder = mobjOrderSets.P_bRemoveOrdItem(sItemKey, sErrMsg)
   
Exit Function
ErrHnd:
   Mtd_bRemoveOrder = False
   sErrMsg = Err.Description
   Call LibApp_Error(Err.Number, sErrMsg, msMODULE_NAME, sPROC_NAME, "Item Key: " & sItemKey)
   
End Function

Public Function Mtd_bClearQueue _
      (Optional ByVal bMed As Boolean = True, _
       Optional ByVal bLab As Boolean = True, _
       Optional ByVal bRad As Boolean = True _
      ) As Boolean
'''************************************************************************
'''Called by: Any
'''
'''Arguments:
''' bMed          Request to remove all queued Med Items
''' bLab          Request to remove all queued Lab Items
''' bRad          Request to remove all queued Rad Items
'''
'''Description:   Remove all requested items from the queue and if mainform is loaded
'''               from the screen
'''
'''Note:
'''Uses Globals: None
'''
'''Uses Modules:
'''
'''Returns:       True  - If successful
'''               False - else
'''
'''Programmer:    Alexander Alberstein
'''
'''************************************************************************
   Const sPROC_NAME  As String = "Mtd_bClearQueue"
   
On Error GoTo ErrHnd

   Mtd_bClearQueue = mobjOrderSets.P_bClearQueue(bMed, bLab, bRad)
   
Exit Function
ErrHnd:
   Mtd_bClearQueue = False
   Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")

End Function


