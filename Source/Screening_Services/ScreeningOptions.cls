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
Attribute VB_Name = "ScreeningOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private frmScreeningOptions As frmOptions
Private mcSelectedReminders As Collection
Private iSelectionMode As ePCMSearch
Private oPCMInformation As PCMItem
Private cDMISInformation As Collection
Private iServiceType As eServiceType

Private bFilterByReminderDate As Boolean
Private bFilterByNotificationDate As Boolean
Private bIncludeNotifiedPatients As Boolean

Private dFromDate As Date
Private dToDate As Date
Private dNotifyFromDate As Date
Private dNotifyToDate As Date

Private sNotificationMethod As String

Public Enum ePCMSearch
    MTFWide = 0
    MTF_DMIS
    PCM_Specific
End Enum

Public Enum eServiceType
    eScreening = 0
    eCPG
End Enum

Public Function CloneOptions() As Object

    Dim oTempOptions As ScreeningOptions
    Dim oOption As Object
    
    Set oTempOptions = New ScreeningOptions
    With oTempOptions
        .FilterByNotificationDate = FilterByNotificationDate
        .FilterByReminderDate = FilterByReminderDate
        .FromDate = FromDate
        .IncludeNotifiedPatients = IncludeNotifiedPatients
        .NotificationMethod = NotificationMethod
        .NotifyFromDate = NotifyFromDate
        .NotifyToDate = NotifyToDate
        .PCMInformation = PCMInformation
        For Each oOption In mcSelectedReminders
            .SelectedReminders.Add oOption, oOption.NCID
        Next
        .SelectionMode = SelectionMode
        .ServiceType = ServiceType
        .ToDate = ToDate
    End With
    
    Set CloneOptions = oTempOptions
    
End Function

Public Function CheckForDuplicate(oInputOptions As ScreeningOptions) As Boolean

    Dim oItem As Object
    Dim oCheckItem As Object
    
    CheckForDuplicate = False
    
    With oInputOptions
        If .SelectionMode = PCM_Specific Then
           If Not .PCMInformation.PCMCode = oPCMInformation.PCMCode Then Exit Function
           If SelectedReminders Is Nothing And mcSelectedReminders Is Nothing Then Exit Function
           If Not .SelectedReminders.Count = mcSelectedReminders.Count Then Exit Function
           For Each oItem In .SelectedReminders
               For Each oCheckItem In mcSelectedReminders
                    If Trim(oCheckItem.NCID) = Trim(oItem.NCID) Then GoTo ItemValidated
               Next
               Exit Function
ItemValidated:
           Next
        ElseIf .SelectionMode = MTFWide Then
           If SelectedReminders Is Nothing And mcSelectedReminders Is Nothing Then Exit Function
           If .SelectedReminders.Count = 1 Then
              If Not Trim(.SelectedReminders(1).NCID) = Trim(mcSelectedReminders(1).NCID) Then Exit Function
           Else
              Exit Function
           End If
        Else
        End If
    End With
    
    CheckForDuplicate = True
    Exit Function

ErrHandler:
    '---show the VB error---
    Call gobjShared.ShowVBError(Err.Number, Err.Description, "ScreeningOptions.CheckForDuplicate", "Screening_Services", vbExclamation)
    'Resume Next

End Function

Public Property Let NotifyToDate(RHV As Date)
    dNotifyToDate = RHV
End Property

Public Property Let FilterByNotificationDate(RHV As Boolean)
    bFilterByNotificationDate = RHV
End Property

Public Property Let FilterByReminderDate(RHV As Boolean)
    bFilterByReminderDate = RHV
End Property

Public Property Let FromDate(RHV As Date)
    dFromDate = RHV
End Property

Public Property Let IncludeNotifiedPatients(RHV As Boolean)
    bIncludeNotifiedPatients = RHV
End Property

Public Property Let NotificationMethod(RHV As String)
    sNotificationMethod = RHV
End Property

Public Property Let NotifyFromDate(RHV As Date)
    dNotifyFromDate = RHV
End Property

Friend Function ParseOptions() As Boolean

    If frmScreeningOptions.optLocation(MTFWide).Value = True Then
       iSelectionMode = MTFWide
    ElseIf frmScreeningOptions.optLocation(MTF_DMIS).Value = True Then
       iSelectionMode = MTF_DMIS
    Else
       iSelectionMode = PCM_Specific
    End If

    If frmScreeningOptions.optServiceType(eScreening).Value = True Then
       iServiceType = eScreening
    Else
       iServiceType = eCPG
    End If

    Set oPCMInformation = frmScreeningOptions.PCMSelected
    Set cDMISInformation = frmScreeningOptions.DMISItems
    
    Set mcSelectedReminders = frmScreeningOptions.SelectedReminders
    If mcSelectedReminders.Count = 0 Then
       MsgBox "Please select a screening service to report on.", vbOKOnly + vbInformation, "Screening Notification"
       ParseOptions = False
       Exit Function
    End If
    
    If frmScreeningOptions.chkReminderDateFilter.Value = vbChecked Then
       bFilterByReminderDate = True
    Else
       bFilterByReminderDate = False
    End If
    dFromDate = frmScreeningOptions.dtpFrom.Value
    dToDate = frmScreeningOptions.dtpTo.Value
    
    If frmScreeningOptions.chkIncludeNotifiedPatients.Value = vbChecked Then
       bIncludeNotifiedPatients = True
       If frmScreeningOptions.chkNotificationDateFilter.Value = vbChecked Then
          bFilterByNotificationDate = True
       Else
          bFilterByNotificationDate = False
       End If
       sNotificationMethod = frmScreeningOptions.cboNotificationMethod.Text
    Else
       bIncludeNotifiedPatients = False
       bFilterByNotificationDate = False
       sNotificationMethod = ""
    End If
    dNotifyFromDate = frmScreeningOptions.dtpNotificationStart.Value
    dNotifyToDate = frmScreeningOptions.dtpNotificationEnd.Value
    
    ParseOptions = True

End Function

Public Property Let PCMInformation(RHV As PCMItem)
    Set oPCMInformation = RHV
End Property

Public Property Let SelectedReminders(RHV As Collection)
    Set mcSelectedReminders = RHV
End Property

Public Property Get SelectedRemindersSQL() As String
    SelectedRemindersSQL = "''"
End Property

Public Sub ShowOptions()
    Set frmScreeningOptions = New frmOptions
    frmScreeningOptions.Visible = False
    frmScreeningOptions.ParentObject = Me
    frmScreeningOptions.Display
End Sub

Private Function GetDefaultMailSettings() As Boolean
    
    Dim sKeyName    As String
    Dim objReg      As RegObj.Registry
    Dim objKey      As RegObj.RegKey
    Dim objValue As RegObj.RegValue
    
    On Error Resume Next
    gsProfileName = ""
    sKeyName = "\HKEY_CURRENT_USER\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles"
    Set objReg = New RegObj.Registry
    
    Set objKey = objReg.RegKeyFromString(sKeyName)
    If objKey Is Nothing Then GoTo CleanUp
    
    For Each objValue In objKey.Values
        If objValue.Name = "DefaultProfile" Then gsProfileName = objValue.Value
    Next

CleanUp:
    Set objValue = Nothing
    Set objKey = Nothing
    Set objReg = Nothing
        
End Function

Public Property Let ToDate(RHV As Date)
    dToDate = RHV
End Property

Private Sub Class_Initialize()
    
    iSelectionMode = PCM_Specific
    Set oPCMInformation = New PCMItem
    
    bFilterByReminderDate = False
    bFilterByNotificationDate = False
    bIncludeNotifiedPatients = True
    
    dFromDate = DateAdd("m", -1, Now)
    dToDate = Now
    dNotifyFromDate = DateAdd("m", -1, Now)
    dNotifyToDate = Now
    
    sNotificationMethod = "All"
    
    GetDefaultMailSettings
    
End Sub


Private Sub Class_Terminate()
    Set frmScreeningOptions = Nothing
    Set oPCMInformation = Nothing
    Set mcSelectedReminders = Nothing
End Sub



Public Property Get SelectionMode() As ePCMSearch
    SelectionMode = iSelectionMode
End Property

Public Property Let SelectionMode(ByVal RHV As ePCMSearch)
    iSelectionMode = RHV
End Property

Public Property Get PCMInformation() As PCMItem
    Set PCMInformation = oPCMInformation
End Property
Public Sub NewPCMInformation()
    Set oPCMInformation = New PCMItem
End Sub


Public Property Get FilterByReminderDate() As Boolean
    FilterByReminderDate = bFilterByReminderDate
End Property



Public Property Get FilterByNotificationDate() As Boolean
    FilterByNotificationDate = bFilterByNotificationDate
End Property

Public Property Get IncludeNotifiedPatients() As Boolean
    IncludeNotifiedPatients = bIncludeNotifiedPatients
End Property


Public Property Get FromDate() As Date
    FromDate = dFromDate
End Property

Public Property Get ToDate() As Date
    ToDate = dToDate
End Property


Public Property Get NotifyFromDate() As Date
    NotifyFromDate = dNotifyFromDate
End Property

Public Property Get NotifyToDate() As Date
    NotifyToDate = dNotifyToDate
End Property


Public Property Get NotificationMethod() As String
    NotificationMethod = sNotificationMethod
End Property

Public Property Get SelectedReminders() As Collection
    If mcSelectedReminders Is Nothing Then Set mcSelectedReminders = New Collection
    Set SelectedReminders = mcSelectedReminders
End Property


Public Property Get ServiceType() As eServiceType
    ServiceType = iServiceType
End Property

Public Property Let ServiceType(ByVal RHV As eServiceType)
    iServiceType = RHV
End Property

Public Property Get DMISInformation() As Collection
    If cDMISInformation Is Nothing Then Set cDMISInformation = New Collection
    Set DMISInformation = cDMISInformation
End Property

Public Property Let DMISInformation(ByVal RHV As Collection)
    Set cDMISInformation = RHV
End Property
