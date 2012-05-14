VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AnPOLEServers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Collection class for AnPOLEServer objects.  It's main funtion is to display an AnPOLEServer.

Option Explicit

'Child modules in order of display on tabs
Public Enum enumAnPModules
  modDiagnosis = 0
  modOrderSummary = 1
  modOrderSets = 2
  modProcedure = 3
  modReminders = 4
  modConsults = 5
  modLab = 6
  modRad = 7
  modMed = 8
  modPI = 9
  modSpec = 10
End Enum

Private Const DIAG As String = "Diagnoses"
Private Const OSETS As String = "Order Sets"
Private Const PROC As String = "Procedures"
Private Const Consults As String = "Consult Orders"
Private Const Labs As String = "Order Entry Lab"
Private Const Rads As String = "Order Entry Rad"
Private Const Meds As String = "Order Entry Med"
Private Const PI As String = "Patient Instructions"
Private Const REMINDERS As String = "Reminders"
Private Const SRTS As String = "Order Eyeware"
Private Const ORDERSUMMARY As String = "Order Summary"

Private mServers As Collection

Private mStatusForm As StatusDialog

Private Sub Class_Initialize()
  Set mServers = New Collection
  
  LoadOLEServers
End Sub

Public Property Get Item(ByVal Module As enumAnPModules) As AnPOLEServer
  On Error Resume Next
  Set Item = mServers.Item(CStr(Module))
End Property

Private Sub AddServer(ByVal ModuleID As enumAnPModules, ByVal NCID As String, _
                      ByVal ServerName As String, ByVal ViewType As Integer, _
                      ByVal Privilege As String)
Dim OLESvr As AnPOLEServer

  On Error GoTo ErrHandler

  Set OLESvr = New AnPOLEServer

  With OLESvr
    .ModuleID = ModuleID
    .NCID = NCID
    .ServerName = ServerName
    .ViewType = ViewType
    .Privilege = Privilege
  End With
  
  'SCR 35637  5/13/03 Phil Crowder (Init here)
  'Get the A/P child module's form default size
  Select Case OLESvr.ModuleID
    Case modDiagnosis, modProcedure, modOrderSets, modPI, modSpec, modConsults, modReminders, modOrderSummary
      OLESvr.Height = 5100
    Case modLab
      OLESvr.Height = 5700
    Case modRad
      OLESvr.Height = 5600
    Case modMed
      'JT SCR 47426 - Height was larger than meds module default height contributing to shrinking with each tab click eventually cutting off the buttons
      OLESvr.Height = 5800
  End Select
  
  'SCR 35637  5/13/03 Phil Crowder (Init here)
  'Get the A/P child module's form default size
  Select Case OLESvr.ModuleID
    Case modDiagnosis, modProcedure, modOrderSets, modPI, modSpec, modConsults, modReminders, modOrderSummary
      OLESvr.Width = 11000
    Case modLab, modRad, modMed
      OLESvr.Width = 11280
  End Select
  
  mServers.Add OLESvr, CStr(ModuleID)
  
  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnPOLEServers.AddServer", "A/P", vbExclamation
  
End Sub

Public Sub CloseOLEServers()
Dim OLESvr As AnPOLEServer

  On Error Resume Next
  
  Set OLESvr = mServers.Item(modOrderSets)
  If Not (OLESvr.Server Is Nothing) Then OLESvr.Server.Mtd_bClearQueue
  
  For Each OLESvr In mServers
    OLESvr.CloseView
  Next
  
End Sub

Public Function Count() As Long
  On Error Resume Next
  
  Count = mServers.Count
End Function

Public Function DisplayServer(ByVal ModuleID As enumAnPModules, ByVal ContainerHwnd As Long) As Boolean
Dim OLESvr As AnPOLEServer
Dim Success As Boolean

  On Error GoTo ErrHandler
      
  Set OLESvr = mServers.Item(ModuleID + 1)    'Tab order is not likely to change but if it
                                              'is then the enumAnPModules must be changed
  
  If OLESvr.Server Is Nothing Then
    Set mStatusForm = New StatusDialog
    Success = InitServer(OLESvr)
    If Success Then Success = OpenView(OLESvr, ContainerHwnd)
  Else
    Success = True
  End If
  
  If Success Then Success = ResizeView(OLESvr)
  
  UnloadStatusForm
  
  DisplayServer = Success
  
  Exit Function
  
ErrHandler:
  UnloadStatusForm
  gShared.ShowVBError Err.Number, Err.Description, "AnPOLEServers.DisplayServer", "A/P", vbExclamation

End Function

Public Sub ForwardMessage(ByVal MessageType As Integer, ByVal Message As String, ByVal SourceNCID As String, Optional ByVal DestinationNCID As String = vbNullString)
Dim OLESvr As AnPOLEServer

  On Error GoTo ErrHandler
  
  If Len(DestinationNCID) <> 0 Then Exit Sub
  
  For Each OLESvr In mServers
    OLESvr.Message MessageType, Message, SourceNCID
  Next
  
  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnPOLEServers.ForwardMessage", "A/P", vbExclamation
  
End Sub

Private Function InitServer(ByRef OLESvr As AnPOLEServer) As Boolean
  Dim Success  As Boolean
  On Error GoTo ErrHandler
    
  If OLESvr.ModuleID = modMed Then    'AnP can't open Meds mod if already open
    Success = Not IsMedsLoaded        '38155 MSolano 6/26/2003
  Else
    Success = True
  End If
  
  If Success Then
    mStatusForm.Display "Initializing Server ", OLESvr.ServerName
    
    InitServer = OLESvr.InitView
    
    If Not InitServer Then
      MsgBxARMd "Unable to Initialize Module.  Module was not loaded.", , "A/P"
    End If
  End If

  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnPOLEServers.InitServer", "A/P", vbExclamation

End Function

'Make sure the meds module is not loaded by another module
Private Function IsMedsLoaded() As Boolean
Dim CommInterface As Comm
  
  On Error GoTo ErrHandler
  
  Set CommInterface = gBB.Comm
  
  If CommInterface.OE_MEDS_STATE <> 0 And CommInterface.OE_MEDS_STATE <> NCID_ASSESMENT_AND_PLAN_APPLICATION Then
    MsgBxARMd "The Order Entry Medication module is currently opened by another module.  Please complete your work before continuing.", , "Assessment and Planning"
    IsMedsLoaded = True
  Else
    CommInterface.OE_MEDS_STATE = NCID_ASSESMENT_AND_PLAN_APPLICATION   '38155 MSolano 6/26/2003
    IsMedsLoaded = False
  End If

  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnPOLEServers.IsMedsLoaded", "A/P", vbExclamation

End Function

Private Sub LoadOLEServers()
'Constants for Security Privileges
Const Priv_Orders_Consults As String = "Orders/Consults"
Const Priv_Orders_Laboratory_Orders As String = "Orders/Laboratory_Orders"
Const Priv_Orders_Pharmacy_Orders As String = "Orders/Pharmacy_Orders"
Const Priv_Orders_Radiology_Orders As String = "Orders/Radiology_Orders"
Const Priv_SRTS_OE As String = "SRTS/oe"
Const Priv_Diagnosis As String = "Diagnosis"
Const Priv_Procedure As String = "Procedure"
Const Priv_Other_Therapies As String = "Other_Therapies"
Const Priv_Order_Sets_Orders As String = "Orders/Order_Sets_Orders"
Const Priv_Wellness As String = "Wellness"
Const Priv_OrderSummary As String = "OrderSummary"
  
  On Error GoTo ErrHandler
  
  AddServer modDiagnosis, NCID_ANP_DIAGNOSIS_APPLICATION, DIAG, cwiREGULAR_VIEW, Priv_Diagnosis
  If gShared.AppMode = modeTheater Then
        AddServer modOrderSummary, NCID_ORDERS_SUMMARY_APPLICATION, ORDERSUMMARY, cwiREGULAR_VIEW, Priv_OrderSummary
  End If
  AddServer modOrderSets, NCID_ORDER_SETS_APPLICATION, OSETS, cwiREGULAR_VIEW, Priv_Order_Sets_Orders
  AddServer modProcedure, NCID_ANP_PROCEDURES_APPLICATION, PROC, cwiREGULAR_VIEW, Priv_Procedure
  AddServer modReminders, NCID_REMINDERS_APPLICATION, REMINDERS, cwiREGULAR_VIEW, Priv_Wellness
  AddServer modConsults, NCID_CONSULT_TRACKING_APPLICATION, Consults, cwiANP_VIEW, Priv_Orders_Consults
  
  Select Case gShared.AppMode
    Case modeCDR, modeCHCSI_GUI
      AddServer modLab, NCID_OE_LAB_APPLICATION, Labs, cwiANP_VIEW, Priv_Orders_Laboratory_Orders
      AddServer modRad, NCID_OE_RAD_APPLICATION, Rads, cwiANP_VIEW, Priv_Orders_Radiology_Orders
      AddServer modMed, NCID_OE_MED_APPLICATION, Meds, cwiANP_VIEW, Priv_Orders_Pharmacy_Orders
    '<SCR 37848 06/16/03 Alexander Alberstein - replaced And with Or to comply with new ITT AppMode settings
    Case modeITT Or modeCHCSI_GUI, modeITT Or modeCDR, modeTheater
    '>
      AddServer modLab, NCID_GEMS_OE_LAB_APPLICATION, Labs, cwiANP_VIEW, Priv_Orders_Laboratory_Orders
      AddServer modRad, NCID_GEMS_OE_RAD_APPLICATION, Rads, cwiANP_VIEW, Priv_Orders_Radiology_Orders
      AddServer modMed, NCID_GEMS_OE_MED_APPLICATION, Meds, cwiANP_VIEW, Priv_Orders_Pharmacy_Orders
  End Select
  
  AddServer modPI, NCID_PATIENTINSTRUCTION_APPLICATION, PI, cwiREGULAR_VIEW, Priv_Other_Therapies
  AddServer modSpec, NCID_SRTS_OE_APPLICATION, SRTS, cwiANP_VIEW, Priv_SRTS_OE                      'SCR 33135  3/19/03  Phil Crowder
  
  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnPOLEServers.LoadOLEServers", "A/P", vbExclamation

End Sub

Private Function OpenView(ByRef OLESvr As AnPOLEServer, ByVal ContainerHwnd As Long) As Boolean
  
  On Error GoTo ErrHandler
  
  mStatusForm.Display "Opening Server ", OLESvr.ServerName
  
  OpenView = OLESvr.OpenView(ContainerHwnd)
  
  'DoEvents

  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnPOLEServers.OpenView", "A/P", vbExclamation

End Function

Private Function ResizeView(ByRef OLESvr As AnPOLEServer) As Boolean
  On Error GoTo ErrHandler

  ResizeView = OLESvr.ResizeView

  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnPOLEServers.ResizeView", "A/P", vbExclamation

End Function

Private Sub UnloadStatusForm()
  On Error Resume Next
  
  If Not (mStatusForm Is Nothing) Then
    Unload mStatusForm
    DoEvents
    Set mStatusForm = Nothing
  End If
End Sub
