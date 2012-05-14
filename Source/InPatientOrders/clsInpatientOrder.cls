VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInpatientOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ICHCSII_CoreAPI2
Implements ICHCSII_CoreAPI3

Private mbMainViewRunning As Boolean

Private mlMenuRefresh       As Long
Private mlMenuNew           As Long
Private mlMenuPrint         As Long ' Quan 81806
Private mlMenuPrintOrder    As Long
Private gInPatientViewRunning As Boolean

Private Sub AddMenuItems()
    
    On Error GoTo ErrHandler
    ' Quan 81699
    mlMenuRefresh = goBB.Comm.AddActionMenuItem("Refresh", "refresh.bmp", cwlMyProgId, "Refresh Screen")
    
    mlMenuNew = goBB.Comm.AddActionMenuItem("New Orders", "new.bmp", cwlMyProgId, "Add a New Order")
    mlMenuPrint = goBB.Comm.AddActionMenuItem("Print List", "print.bmp", cwlMyProgId, "Print Orders")
    ' Quan 81806
    mlMenuPrintOrder = goBB.Comm.AddPrintMenuItem("Inpatient Order", cwlMyProgId, "Print Selected Orders")
          
    Exit Sub
  
ErrHandler:
  goBB.SharedObj.ShowVBError Err.Number, Err.Description, ".AddMenuItems", "InpatientOrder", vbExclamation
    
    
End Sub

Public Sub RemoveMenus()

    If goBB Is Nothing Then Exit Sub
    If goBB.Comm Is Nothing Then Exit Sub
    
    If mlMenuRefresh > 0 Then
        goBB.Comm.DeleteMenuItem mlMenuRefresh
        mlMenuRefresh = 0
    End If

    If mlMenuNew > 0 Then
        goBB.Comm.DeleteMenuItem mlMenuNew
        mlMenuNew = 0
    End If
    
    If mlMenuPrint > 0 Then
        goBB.Comm.DeleteMenuItem mlMenuPrint
        mlMenuPrint = 0
    End If
    
    'Quan 81806
    If mlMenuPrintOrder > 0 Then
        goBB.Comm.DeleteMenuItem mlMenuPrintOrder
        mlMenuPrintOrder = 0
    End If

End Sub

Public Sub RefreshMenus()
On Error GoTo ErrHandler
    
    RemoveMenus
    AddMenuItems

    Exit Sub
ErrHandler:
    goBB.SharedObj.ShowVBError Err.Number, Err.Description, TypeName(Me) & ".RefreshMenus()", "InpateintOrder", vbExclamation
End Sub
Public Sub MenuItemSelected(ByVal vlMenuId As Long)

    On Error GoTo ErrHandler
    
    Screen.MousePointer = vbHourglass
    Select Case vlMenuId
        Case mlMenuRefresh
            Call frmInpatientOrder.RefreshData ' Quan 81699
        Case mlMenuNew
            Call frmInpatientOrder.NewOrder
        Case mlMenuPrint
            Call frmInpatientOrder.PrintList(True)   'Quan 81805
        Case mlMenuPrintOrder ' Quan 81806
            Call frmInpatientOrder.PrintList
    End Select
    
    RefreshMenus
    
    Screen.MousePointer = vbDefault
    Exit Sub
ErrHandler:
    goBB.SharedObj.ShowVBError Err.Number, Err.Description, TypeName(Me) & ".MenuItemSelected()", "InpateintOrder", vbExclamation
    Screen.MousePointer = vbDefault
End Sub

Public Function QueryView(ByVal viViewType As Integer) As Boolean
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
End Function

Public Sub PatientSelected()
    If glPrevPat <> goBB.Patient.UnitNumber Then
        If mbMainViewRunning Then
            Call frmInpatientOrder.Refresh
        End If
        glPrevPat = goBB.Patient.UnitNumber
    End If
End Sub

Public Sub Message(ByVal viMessageType As Integer, ByVal vsMessageText As String, ByVal vsSourceServerNCID As String)
    On Error GoTo ErrHandler
    
    If vsSourceServerNCID <> NCID_CW_APPLICATION And vsSourceServerNCID <> NCID_INPATIENT_ORDERS_APPLICATION Then
        Exit Sub
    End If
    
    Select Case viMessageType
        Case cwiREFRESH_DATA
            If mbMainViewRunning Then
                Call frmInpatientOrder.Refresh
            End If
    End Select

    Exit Sub
    
ErrHandler:
    Call goBB.SharedObj.ShowVBError(Err.Number, Err.Description, ".Message", "InpateintOrder", vbExclamation)
End Sub

Public Sub ResizeView(ByVal viViewType As Integer, ByVal vlLeft As Long, ByVal vlTop As Long, ByVal vlWidth As Long, ByVal vlHeight As Long)
    On Error GoTo ErrHandler
    '---reset global "client area" parameters---
    cwlLeft = vlLeft
    cwlTop = vlTop
    cwlWidth = vlWidth
    cwlHeight = vlHeight
    
    If viViewType = cwiREGULAR_VIEW Then
        If goBB.Comm.Message(cwiMODAL_FORM_SHOWING, "", NCID_CW_APPLICATION, NCID_INPATIENT_ORDERS_APPLICATION) Then
            Exit Sub
        End If
        If Not cwbFirstRegularResize Then
            With frmInpatientOrder
                .WindowState = vbNormal
                .Visible = False
                .Move cwlLeft, cwlTop, cwlWidth, cwlHeight
                .Visible = True
            End With
            Call frmInpatientOrder.RefreshData 'scr 56638 -1/12/2005
        Else
            cwbFirstRegularResize = False
            If Not frmInpatientOrder.Display() Then
                ICHCSII_CoreAPI3_CloseView2 cwiREGULAR_VIEW
            End If
        End If
    Else
        Screen.MousePointer = vbDefault
        Beep
        MsgBxARMd "Unsupported resize view type: " & CStr(viViewType) & " in Inpatient.ResizeView.", vbExclamation, "Problems"
    End If
   
    Exit Sub
ErrHandler:
    Call goBB.SharedObj.ShowVBError(Err.Number, Err.Description, ".ResizeView", "InpateintOrder", vbExclamation)
End Sub


Public Function OpenView(ByVal viViewType As Integer, ByVal vlParenthWnd As Long, rhMainForm As Long) As Boolean
    On Error GoTo ErrHandler
    
    If CheckUserPrivilege = False Then
        MsgBxARMd "You don't have privilege to read the Inpatient Order.", vbCritical, "InPatient Order"
        OpenView = False
        Exit Function
    End If

    If goBB.Patient.UnitNumber = 0 Then
        MsgBxARMd "A Patient Number is needed", vbCritical, "InPatient Order"
        OpenView = False
        Exit Function
    End If
    
    If viViewType = cwiREGULAR_VIEW Then
        cwbFirstRegularResize = True
        gInPatientViewRunning = True
        SetParent frmInpatientOrder.hwnd, vlParenthWnd
        rhMainForm = frmInpatientOrder.hwnd
        OpenView = True
        AddMenuItems
    Else    '--unsupported view "type"--
        OpenView = False
    End If
    
    Exit Function
    
ErrHandler:
    goBB.SharedObj.ShowVBError Err.Number, Err.Description, "OrderMgt.OpenView", "OrderMgt", vbCritical
End Function


Private Sub Class_Initialize()
    
    If goBB Is Nothing Then Set goBB = New BlackBorad

End Sub

Private Sub Class_Terminate()
    
    Set goBB = Nothing
    
End Sub

Public Sub ICHCSII_CoreAPI2_Initialize2(ByVal vlMyProgId As Long, robjComm As Object, robjLogonEnvironment As Object, robjConfig As Object, robjPatient As Object, robjUser As Object)
    On Error GoTo ErrHandler
    
    cwlMyProgId = vlMyProgId
    '---create pointers to objects---
    Set goBB.Comm = robjComm
    Set goBB.Logon = robjLogonEnvironment
    Set goBB.Config = robjConfig
    Set goBB.Patient = robjPatient
    Set goBB.User = robjUser
    
    glPrevPat = goBB.Patient.UnitNumber
    
    App.HelpFile = goBB.SharedObj.CWHelpFile

    '---set global flags---
    gInPatientViewRunning = False
        
    Exit Sub
ErrHandler:
    Call goBB.SharedObj.ShowVBError(Err.Number, Err.Description, "clsInpatientOrder.ICHCSII_CoreAPI2_Initialize2", "InpateintOrder", vbExclamation)
    
End Sub



Public Function ICHCSII_CoreAPI3_CloseView2(ByVal viViewType As Integer) As Boolean
On Error GoTo ErrHandler

    ICHCSII_CoreAPI3_CloseView2 = True

    If gInPatientViewRunning Then
        gInPatientViewRunning = False
        '---this indicates that the CW Core is closing the "view", NOT this OLE Server---
        cwbCloseViewMyself = False

        Unload frmInpatientOrder
    End If
    
    Exit Function
    
ErrHandler:
  ICHCSII_CoreAPI3_CloseView2 = False
  goBB.SharedObj.ShowVBError Err.Number, Err.Description, ".ICHCSII_CoreAPI3_CloseView2", "InpateintOrder", vbExclamation

End Function

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
  goBB.SharedObj.ShowVBError Err.Number, Err.Description, ".CheckUserPrivilege", "InpateintOrder", vbExclamation

End Function





