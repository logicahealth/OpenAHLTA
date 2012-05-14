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
Attribute VB_Name = "OrderMgt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ICHCSII_CoreAPI2
Implements ICHCSII_CoreAPI3

Private mlMenuRefresh       As Long
Private mlMedRefill         As Long
Private mlMedNewPrescription As Long
Private mlSigned            As Long
Private mlMenuPrint         As Long
Private gInPatientViewRunning As Boolean


Private Sub Class_Initialize()
    If goBB Is Nothing Then Set goBB = New Blackboard
End Sub

Public Sub ICHCSII_CoreAPI2_Initialize2(ByVal vlMyProgId As Long, ByRef robjComm As Object, ByRef robjLogonEnvironment As Object, ByRef robjConfig As Object, ByRef robjPatient As Object, ByRef robjUser As Object)
    
    On Error GoTo ErrHandler
    
    '---save the "ProgId" assigned to me by CW---
    cwlMyProgId = vlMyProgId
    
    '---create pointers to objects---
    Set goBB.Comm = robjComm
    Set goBB.Logon = robjLogonEnvironment
    Set goBB.Config = robjConfig
    Set goBB.Patient = robjPatient
    Set goBB.User = robjUser

    '---set global flags---
    gInPatientViewRunning = False
    
    If gobjOrderMgtInpt Is Nothing Then
        Set gobjOrderMgtInpt = New OrderMgtOpsInpt
    End If

    
    '---Instantiate Help-file object
    App.HelpFile = goBB.SharedObj.CWHelpFile
    
    Call goBB.Options.GetUserOptionsFromConfig
    
    Exit Sub

ErrHandler:
    goBB.SharedObj.ShowVBError Err.Number, Err.Description, "OrderMgt.Initialize", "OrderMgt", vbExclamation

End Sub
Private Function ICHCSII_CoreAPI3_CloseView2(ByVal viViewType As Integer) As Boolean
  Call CloseView(viViewType)
  ICHCSII_CoreAPI3_CloseView2 = True
End Function

Public Sub CloseView(ByVal viViewType As Integer)
    If viViewType = cwiALL_VIEWS Then
        '---close the "main view"---
        If gInPatientViewRunning Then
            '---set global flags---
            gInPatientViewRunning = False
            
            '---this indicates that the CW Core is closing the "view", NOT this OLE Server---
            cwbCloseViewMyself = False
            
            '---close ALL forms that are associated with all "views"---
            Unload frmOrderMgt
            Set goCoreMenuID = Nothing
        End If
    ElseIf viViewType = cwiREGULAR_VIEW Then
        '---close the "main view"---
        If gInPatientViewRunning Then
            '---set global flags---
            gInPatientViewRunning = False
            
            '---this indicates that the CW Core is closing the "view", NOT this OLE Server---
            cwbCloseViewMyself = False
            
            '---close ALL forms that are associated with all "views"---
            Unload frmOrderMgt
            Set goCoreMenuID = Nothing
        End If
    End If
    
End Sub

Public Sub ResizeView(ByVal viViewType As Integer, ByVal vlLeft As Long, ByVal vlTop As Long, ByVal vlWidth As Long, ByVal vlHeight As Long)
    On Error GoTo ErrHandler

    cwlLeft = vlLeft
    cwlTop = vlTop
    cwlWidth = vlWidth
    cwlHeight = vlHeight
    
    '---perform different actions based upon the "view type"---
    If viViewType = cwiREGULAR_VIEW Then
        If goBB.Comm.Message(cwiMODAL_FORM_SHOWING, "", NCID_CW_APPLICATION, NCID_ORDER_MANAGER_INPATIENT_APPLICATION) Then
           Exit Sub
        End If
        If Not cwbFirstRegularResize Then
            With frmOrderMgt
              .WindowState = vbNormal
              .Visible = False
              .Move cwlLeft, cwlTop, cwlWidth, cwlHeight
              .Visible = True
              'SCR 106955 - LCI - 09/26/2006
              .grdMgt.Row = 0 'Deselect current selection
            End With
            'SCR 106955 - LCI - 09/26/2006
            Call SetMenuButton(mlDispHistory, False)
            Call frmOrderMgt.RefreshData 'scr 56665 -1/4/2005
        Else    '--first "regular" resize, need to perform "initailization" procedures--
            cwbFirstRegularResize = False
            
            If Not frmOrderMgt.Display() Then
'                DeleteMenus
                Unload frmOrderMgt
                gInPatientViewRunning = False
                Exit Sub
            End If
        End If
    End If
    
    Exit Sub
    
ErrHandler:
    If Err.Number = 401 Then
      Resume Next
    Else
      goBB.SharedObj.ShowVBError Err.Number, Err.Description, "OrderMgt.ResizeView", "OrderMgt", vbExclamation
    End If

End Sub

Public Function OpenView(ByVal viViewType As Integer, ByVal vlParenthWnd As Long, rhMainForm As Long) As Boolean
    
    On Error GoTo ErrHandler
    
    If viViewType = cwiREGULAR_VIEW Then
        cwbFirstRegularResize = True
        gInPatientViewRunning = True
        AddMenus
        SetParent frmOrderMgt.hwnd, vlParenthWnd
        rhMainForm = frmOrderMgt.hwnd
        OpenView = True
    Else    '--unsupported view "type"--
        OpenView = False
        Exit Function
    End If
    
    Exit Function
    
ErrHandler:
    goBB.SharedObj.ShowVBError Err.Number, Err.Description, "OrderMgt.OpenView", "OrderMgt", vbCritical
    
End Function
Public Function QueryView(ByVal viViewType As Integer) As Boolean
    Select Case viViewType
        Case cwiREGULAR_VIEW
            QueryView = True    '--this view is NOT optional--
        Case cwiPATIENT_SUMMARY_VIEW
            QueryView = False
        Case cwiENCOUNTER_SUMMARY_VIEW
            QueryView = False
        Case cwiPROPERTY_VIEW
            If IsInpatient Then
                QueryView = True
            Else
                QueryView = False
            End If
        Case Else   '--unknown "view type", not supported--
            QueryView = False
    End Select
End Function
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
            goBB.Comm.Message cwiOK_TO_CHANGE_PATIENT, "Y", vsSourceServerNCID, NCID_ORDER_MANAGER_INPATIENT_APPLICATION
        
        Case cwiOK_TO_CLOSE_VIEW
            goBB.Comm.Message cwiOK_TO_CLOSE_VIEW, "Y", vsSourceServerNCID, NCID_ORDER_MANAGER_INPATIENT_APPLICATION
    
        Case cwiPROPERTIES
            frmOrderMgt.EditSelections
        
        Case Else   '--do nothing--
    End Select

    Exit Sub
    
ErrHandler:
    '---show the VB error---
    Call goBB.SharedObj.ShowVBError(Err.Number, Err.Description, "DataManager.Message", "DataManager", vbExclamation)
    
    Exit Sub
End Sub
Public Sub MenuItemSelected(ByVal vlMenuId As Long)
    On Error GoTo ErrHandler

    Select Case vlMenuId
        Case mlMenuRefresh
            Call frmOrderMgt.RefreshData
        Case mlViewEdit
                    
            Call frmOrderMgt.ViewEditOrder
        Case mlMedRefill
            '** Deselect the Hilight Row
            frmOrderMgt.grdMgt.Row = 0
            If goBB.Patient.SearchPatient Then
                Call frmActiveMeds.Display
                Call frmOrderMgt.RefreshData
            End If
        
        Case mlMedNewPrescription
        
            Dim bDispenseMedInProgress As Boolean
        
            '** SCR - 7/6/05 - Frank; Add New Rx button
            '** Deselect the Hilight Row
            frmOrderMgt.grdMgt.Row = 0
            
            '** 7/14/05 - Frank
            '** Still goto Label Printing after user CANCEL from Patient Search screen
            '** Current selected patient is used in this case
            goBB.Patient.SearchPatient
            
            'SCR 73556 - Frank -- 7/14/05
            If goBB.Patient Is Nothing Then
                Exit Sub
            ElseIf goBB.Patient.Name = "" Then
                Exit Sub
            End If
            
            '* 7/6/05 - Label Printing takes a long time closing
            '*           saving info to XML file
            '*           User double-click another med with Label Printing is closing
            '*           causes the newly created object to be destroyed
            '**** Add flag so that until Label Printing finish closing, before allowing
            '*    user to view a new Label Printing
            If Not bDispenseMedInProgress Then
            
                bDispenseMedInProgress = True
                Call frmOrderMgt.PrintLabor(True, True) '*** CalledFromOrderMgmt=true, NewRx=true
                bDispenseMedInProgress = False
                
            End If
        
            Call frmOrderMgt.RefreshData
            Screen.MousePointer = vbNormal  '** Hourglass is set int frmDispense.Close_Click()
            
        ' Quan 78339
        Case mlCreateAppt
            goBB.Comm.Message cwiGENERIC, "CREATE_APPT", NCID_CLINIC_SCHEDULE_APPLICATION, NCID_ORDER_MANAGER_INPATIENT_APPLICATION
'            'frmOrderMgt.RefreshData
'        ' Quan 78339
'        Case mlCreateEnc
'            ' goBB.Comm.Message cwiGENERIC, "CREATE_APPT", NCID_CLINIC_SCHEDULE_APPLICATION, NCID_ORDER_MANAGER_INPATIENT_APPLICATION

        ' Quan 80054
        Case mlDispHistory
            Call frmOrderMgt.ShowDispHistory
        Case mlSigned
            'SCR 83976 -- 12/13/05 - Frank
            '** Requires user to have Signing Current Encoutner privilege
            If goBB.User.HasPrivilegeEx(Priv_Current_Encounter, Priv_Sign) Then
                Call frmOrderMgt.SetSignOrders
            Else
                MsgBxARMd "You do not have sign privilege.", vbOKOnly, "Order Manager"
            End If
        Case mlDiscontinue
            Call frmOrderMgt.SetDiscontinueOrders
        Case mlMenuPrint
            Call frmOrderMgt.PrintList
    
        Case goCoreMenuID(orderActiveORPend)
            goBB.Options.ResetSelectionStatus (orderActiveORPend)
            frmOrderMgt.RefreshData
            
        Case goCoreMenuID(orderComplete)
            goBB.Options.ResetSelectionStatus (orderComplete)
            frmOrderMgt.RefreshData
            
        Case goCoreMenuID(orderDiscontinued)
            goBB.Options.ResetSelectionStatus (orderDiscontinued)
            frmOrderMgt.RefreshData
            
        Case goCoreMenuID(orderPreliminary)
            goBB.Options.ResetSelectionStatus (orderPreliminary)
            frmOrderMgt.RefreshData
            
        Case goCoreMenuID(orderAddressed)
            goBB.Options.ResetSelectionStatus (orderAddressed)
            frmOrderMgt.RefreshData
            
        Case goCoreMenuID(orderUnAddressed)
            goBB.Options.ResetSelectionStatus (orderUnAddressed)
            frmOrderMgt.RefreshData
            
        Case goCoreMenuID(orderDispensed)
            goBB.Options.ResetSelectionStatus (orderDispensed)
            frmOrderMgt.RefreshData
            
        Case goCoreMenuID(orderSigned)
            goBB.Options.ResetSelectionStatus (orderSigned)
            frmOrderMgt.RefreshData
        
        Case goCoreMenuID(allStatus)
            goBB.Options.ResetSelectionStatus (allStatus)
            frmOrderMgt.RefreshData
            
        Case mlPrintLabel
            Call frmOrderMgt.PrintLabor
            
        Case Else   '--do nothing--
 
    End Select
    
    '*** 4/15/05 - Frank
    Screen.MousePointer = vbNormal
    Exit Sub

ErrHandler:
    '---show the VB error---
    Call goBB.SharedObj.ShowVBError(Err.Number, Err.Description, "OrderMgt.MenuItemSelected", "DataManager", vbExclamation)
    
    Exit Sub
End Sub
Private Sub AddMenus()

    On Error GoTo ErrHandler
    
    If goCoreMenuID Is Nothing Then
      Set goCoreMenuID = New cCoreMenuID
    End If
  
    mlMenuRefresh = goBB.Comm.AddActionMenuItem("&Refresh", "refresh.bmp", cwlMyProgId, "Refresh Screen")
    
    '*** SCR 73361 - 7/13/05 - Frank
    '*** Only User with can sign pharmacy Orders privilege can create a new med order
    If goBB.User.HasPrivilegeEx(Priv_Orders_Pharmacy_Orders, Priv_Sign) Then
        mlMedNewPrescription = goBB.Comm.AddActionMenuItem("New Rx", "Meds16_New_Rx.bmp", cwlMyProgId, "New Prescription")
    End If
    
'     Quan 78339
    mlCreateAppt = goBB.Comm.AddActionMenuItem("&New Appt", "new.bmp", cwlMyProgId, "Create Appointment")
''     Quan 78338
'    mlCreateEnc = goBB.Comm.AddActionMenuItem("&New Enc", "encounter.bmp", cwlMyProgId, "Create Encounter")
    ' Quan 80054
    mlDispHistory = goBB.Comm.AddActionMenuItem("Disp History", "new.bmp", cwlMyProgId, "Dispense History")
        
        
    mlMedRefill = goBB.Comm.AddActionMenuItem("MedRefill", "Meds16.bmp", cwlMyProgId, "Medication Refill")
    mlViewEdit = goBB.Comm.AddActionMenuItem("View/Edit", "note09.ico", cwlMyProgId, "View or Edit Order")

    If IsInpatient Then
        mlSigned = goBB.Comm.AddActionMenuItem("Sign", "sign.bmp", cwlMyProgId, "Sign Order")
        mlDiscontinue = goBB.Comm.AddActionMenuItem("Discontinue", "delete.bmp", cwlMyProgId, "Discontinue Order")
    End If
    mlMenuPrint = goBB.Comm.AddActionMenuItem("Print List", "print.bmp", cwlMyProgId, "Print Orders")
    mlPrintLabel = goBB.Comm.AddActionMenuItem("Print Label", "print.bmp", cwlMyProgId, "Print Label")
    
    '--- MENU ITEMS SETTING SELECTION FILTER CRITERIA
    goCoreMenuID.Add goBB.Comm.AddViewMenuItem("Any Status", cwlMyProgId, "Any Status"), allStatus
    goCoreMenuID.Add goBB.Comm.AddViewMenuItem("Active/Pending", cwlMyProgId, "Active/Pending"), orderActiveORPend
    goCoreMenuID.Add goBB.Comm.AddViewMenuItem("Complete", cwlMyProgId, "Complete"), orderComplete
    goCoreMenuID.Add goBB.Comm.AddViewMenuItem("Discontinued", cwlMyProgId, "Discontinued"), orderDiscontinued
    goCoreMenuID.Add goBB.Comm.AddViewMenuItem("Preliminary", cwlMyProgId, "Preliminary"), orderPreliminary
    goCoreMenuID.Add goBB.Comm.AddViewMenuItem("Completed/Addressed", cwlMyProgId, "Completed/Addressed"), orderAddressed
    goCoreMenuID.Add goBB.Comm.AddViewMenuItem("Completed/Unaddressed", cwlMyProgId, "Completed/Unaddressed"), orderUnAddressed
    goCoreMenuID.Add goBB.Comm.AddViewMenuItem("Dispensed", cwlMyProgId, "Dispensed"), orderDispensed
    goCoreMenuID.Add goBB.Comm.AddViewMenuItem("UnSigned", cwlMyProgId, "UnSigned"), orderSigned
    
    Exit Sub
    
ErrHandler:
    goBB.SharedObj.ShowVBError Err.Number, Err.Description, "OrderMgt.AddMenus", "OrderMgt", vbExclamation
    
End Sub

Private Sub Class_Terminate()
    Set goBB = Nothing
End Sub

Public Property Get OrderID() As String
    OrderID = frmOrderMgt.OrderID
End Property
Public Property Get OrderType() As String
    OrderType = frmOrderMgt.OrderType
End Property


