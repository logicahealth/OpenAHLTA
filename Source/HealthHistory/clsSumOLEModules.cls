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
Attribute VB_Name = "clsSumOLEModules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsSumOLEModule"
Attribute VB_Ext_KEY = "Member0" ,"clsSumOLEModule"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private mcol As Collection
Private mobjParent As HealthHistory


Public Property Set Parent(RHV As HealthHistory)
    Set mobjParent = RHV
End Property

Public Function AddModule(oNewModule As clsSumOLEModule)
    mcol.Add oNewModule, oNewModule.sApplicationNCID
End Function

Public Function RetrievOLEServerInfo() As Boolean
    Dim sTemp As Variant    '- holds the array of info in the apvalue field
    Dim objMod  As clsSumOLEModule
    Dim sbuf    As String
    Dim i       As Integer
    Dim iNum    As Integer
    Dim oSvr    As Variant
    Dim colSvrs As Collection
    Dim colDefaults As Collection
    Dim bFound As Boolean
    
    Set mcol = New Collection  '-- start the new main cls
    Dim sObjectName As String
    Set colSvrs = mobjParent.CWComm.GetOLEServers
    
    If gobjShared.IsAppMode(modeCHCSI_GUI) Then
      Set colDefaults = New Collection
      
      colDefaults.Add "NO;0;0;0;0;", "CHCSIIEncounterPrevious.EncPrevious"
      colDefaults.Add "NO;0;0;0;0;", "CouplerLauncher.Coupler"
      colDefaults.Add "NO;0;0;0;0;", "MMMHISDemogrph.Demogrph"
      colDefaults.Add "YES;0;0;32;24;", "CHCSII_Appointments.Appointments"
      colDefaults.Add "YES;34;0;31;24;", "MMMHISProblems.Problem"
      colDefaults.Add "YES;67;0;31;25;", "MMMHISAllergy.Allergy"
      colDefaults.Add "YES;0;26;98;25;", "MMMHISMeds.Meds"
      colDefaults.Add "YES;0;52;98;29;", "CHCSII_LabResults.Lab"
      colDefaults.Add "YES;0;82;98;18;", "CHCSII_RadResults.RadResults"
    End If

    For Each oSvr In colSvrs
       ' Debug.Print oSvr.sservername
        
        'SCR #32793  SL 2-21-03... need to implement security for HH
        'If oSvr.bSupportsPatSummary Then
        If oSvr.Config.HasPatientSummary And oSvr.UserHasAccess _
           And oSvr.ApplicationNCID <> NCID_INPATIENT_ORDERS_APPLICATION Then 'SCR 64214 - Jane Shen
            Set objMod = New clsSumOLEModule
            objMod.sObjectName = oSvr.Config.servername
            objMod.sApplicationNCID = oSvr.ApplicationNCID
'<< SCR 14113
            objMod.sApplicationName = oSvr.Config.buttoncaption

'<< Gems Meds temp workaround
            If objMod.sObjectName = "Gems_MMMHISMeds.Meds" Then
                sObjectName = "MMMHISMeds.Meds"
            Else
                sObjectName = objMod.sObjectName
            End If
                
            'SCR #28777-  SL 10-28-02
            '  moving the AddModule call outside the mobjParent.Config.getValue
            '  block below... this change was made in Daily Build by BS, but
            '  not in Source
            Call AddModule(objMod)
            
            bFound = mobjParent.Config.getValue(NCID_CW_APPLICATION, "support patient summary view", sObjectName, sbuf)
            
            If bFound Then
              sTemp = Split(sbuf, ";")
            ElseIf gobjShared.IsAppMode(modeCHCSI_GUI) Then
              sTemp = Split(colDefaults.Item(sObjectName), ";")
            Else
              sTemp = vbNullString
            End If
            
            If IsArray(sTemp) Then
              objMod.bOnSummary = (sTemp(0) = "YES")
              objMod.lSumLeftPercent = sTemp(1)
              objMod.lSumTopPercent = sTemp(2)
              objMod.lSumWidthPercent = sTemp(3)
              objMod.lSumHeightPercent = sTemp(4)
            End If
            
            Set objMod = Nothing
         End If
    Next oSvr
    Set oSvr = Nothing
    Set colSvrs = Nothing
    
    mobjParent.CWComm.SetStatusText ""
End Function


'*******************************************************************************************
'  Description: This function starts up all of the selected summary OLE Servers for the
'               specified view "type".
'
'  Parameters:
'    vlParenthWnd - Contains the parent "hWnd" of the summary view to use.
'    vlParentScaleWidth - Contains the "scale width" of the parent window.
'    vlParentScaleHeight - Contains the "scale height" of the parent window.
'
'  Return Value:
'    Returns TRUE if the summary OLE Servers were started successfully.
'    Returns FALSE if an error occurred in starting the summary OLE Servers.
'*******************************************************************************************
Public Function StartSummaryOLEServers(ByVal vlParenthWnd As Long, _
                                       ByVal vlParentScaleWidth As Long, _
                                       ByVal vlParentScaleHeight As Long) As Boolean
    Dim iNumItems As Integer
    Dim iNumOpened As Integer
    Dim i As Integer
    Dim objOLEServer As Object
    Dim bNeededToCreateOLEServer As Boolean
    Dim sbuf As String
    Dim lDummyMenuId As Long
    Dim lLeft As Long
    Dim lTop As Long
    Dim lWidth As Long
    Dim lHeight As Long
    Dim lDummy As Long
    Dim iSummaryInfoIndex As Integer
    Dim bErrorCreatingOLEServer As Boolean
    Dim objOLEServer2 As ICHCSII_CoreAPI2
    

    If Not DebugMode Then On Error GoTo errhandler
    
Dim objSumMod As clsSumOLEModule
    iNumOpened = 0
    For Each objSumMod In mcol
'<< only the servers that are "on"
        If objSumMod.bOnSummary Then
            Call OpenSummaryOLEServer(objSumMod, vlParenthWnd, vlParentScaleWidth, vlParentScaleHeight)
            iNumOpened = iNumOpened + 1
        End If   '- objSumMod.bOnSummary Then
    Next objSumMod
    
    StartSummaryOLEServers = True
    
    If iNumOpened = 0 Then
      sbuf = "Health History"
      MsgBox "No Modules have been selected for " & sbuf & ". Use 'Tools' - 'Properties - Health History' to select modules for display.", vbOKOnly, sbuf
    
    End If
    mobjParent.CWComm.SetStatusText ""
    Exit Function
    
errhandler:
    '---show the VB error---
    gobjShared.ShowVBError Err.Number, Err.Description, "StartSummaryOLEServers", App.Title, vbExclamation
    
    StartSummaryOLEServers = False
    mobjParent.CWComm.SetStatusText ""
    Exit Function

End Function

Public Function OpenSummaryOLEServer(objSumMod As clsSumOLEModule, _
                               ByVal vlParenthWnd As Long, _
                               ByVal vlParentScaleWidth As Long, _
                               ByVal vlParentScaleHeight As Long) As Boolean

On Error Resume Next
    Dim lDummy As Long
    Dim sbuf   As String
    
    mobjParent.CWComm.SetStatusText "Starting: " & objSumMod.sApplicationName
    If objSumMod.Module Is Nothing Then
    '-- start it
                
        Set objSumMod.Module = mobjParent.CWComm.InitializeOLEServer(objSumMod.sApplicationNCID)
        
        If objSumMod.Module Is Nothing Then
            Screen.MousePointer = vbDefault
            sbuf = "Unable to open a Summary view for the clinical module: " & objSumMod.sApplicationName
            MsgBox sbuf, vbExclamation, App.Title
            Exit Function
        End If
        objSumMod.MenuItems = New Collection
        objSumMod.bSummaryViewRunning = False
    End If
    
    If Not objSumMod.bSummaryViewRunning Then
    '-- open it and resize it
            With objSumMod
                '---open the specified view "type" for the OLE Server---
                If .Module.OpenView(cwiPATIENT_SUMMARY_VIEW, vlParenthWnd, lDummy) Then
                    objSumMod.bSummaryViewRunning = True
                    Call .Resize(vlParentScaleWidth, vlParentScaleHeight)
                    'If .sApplicationNCID = "77613" Then
                    '   Call objSumMod.Resize(vlParentScaleWidth, vlParentScaleHeight)
                    'End If
                Else    '--unable to open a summary view for the OLE Server--
                    Screen.MousePointer = vbDefault
                    sbuf = "Unable to open a Summary view for the clinical module: " & objSumMod.sApplicationName
                    MsgBox sbuf, vbExclamation, App.Title
                End If  '- objSumMod.Module.OpenView(cwiPATIENT_SUMMARY_VIEW, vlParenthWnd, lDummy)
            End With    '- objSumMod.
        
    Else
    '- just resize it
       Call objSumMod.Resize(vlParentScaleWidth, vlParentScaleHeight)
    End If
End Function

'*******************************************************************************************
'  Description: This subroutine resizes all of the selected summary OLE Servers for the
'               specified view "type".
'
'  Parameters:
'    viSummaryViewType - Contains the "type" of summary view.
'    vlParentScaleWidth - Contains the "scale width" of the parent window.
'    vlParentScaleHeight - Contains the "scale height" of the parent window.
'*******************************************************************************************
Public Sub ResizeSummaryOLEServers(ByVal vlParentScaleWidth As Long, _
                                    ByVal vlParentScaleHeight As Long)
    Dim iNumItems As Integer
    Dim i As Integer
    Dim objOLEServer As Object
    Dim lLeft As Long
    Dim lTop As Long
    Dim lWidth As Long
    Dim lHeight As Long
    Dim iSummaryInfoIndex As Integer
    Dim sbuf As String
    Dim oSumMod As clsSumOLEModule

    On Error Resume Next
    
    Screen.MousePointer = vbHourglass
    For Each oSumMod In mcol
        If oSumMod.bSummaryViewRunning Then
            Call oSumMod.Resize(vlParentScaleWidth, vlParentScaleHeight)
        End If
    Next oSumMod
    
    
    
    Screen.MousePointer = vbDefault
End Sub

'*******************************************************************************************
'  Description: This subroutine prints the summary report for the specified view "type".
'
'  Parameters:
'    vbPrintPreview - TRUE if print preview is desired, FALSE otherwise.
'*******************************************************************************************
Public Sub PrintSummaryOLEServers(ByVal vbPrintPreview As Boolean)
    Dim objOLEServer As Object
    Dim sReportName As String
    Dim iPrintSummaryMessage As Integer
    Dim i As Integer
    Dim iNumItems As Integer
    Dim objPrint As CWPrint

    If Not DebugMode Then On Error GoTo errhandler

    iPrintSummaryMessage = cwiPRINT_PATIENT_SUMMARY_VIEW
    sReportName = "PatSum.rpt"

    Screen.MousePointer = vbHourglass
    
    Set objPrint = New CWPrint
    
    If Not objPrint.printHeaderandFooter(mobjParent.LogonEnviron, mobjParent.Patient) Then
        Screen.MousePointer = vbDefault
        Beep
        MsgBox "Unable to print the standard header and footer.  The Health History report will not be printed.", vbExclamation, App.Title
        
        Exit Sub
    End If
    
    '---need to close the exclusive access to the "shared" print database---
    objPrint.PrintDB.Close
    Set objPrint.PrintDB = Nothing
    
    Dim oSumMod As clsSumOLEModule
    
    For Each oSumMod In mcol
        If oSumMod.bSummaryViewRunning Then
            oSumMod.Module.Message iPrintSummaryMessage, objPrint.DatabasePath, ""
        End If
    Next oSumMod
    
    If Not objPrint.printSpoolJob(sReportName, mobjParent.CWComm.ReportsPath, vbPrintPreview) Then
        Screen.MousePointer = vbDefault
        Beep
        MsgBox "An error occurred in printing the summary report.", vbExclamation, App.Title
    End If
    
    GoTo Cleanup
    
errhandler:
    gobjShared.ShowVBError Err.Number, Err.Description, "gPrintSummaryOLEServers", App.Title, vbExclamation
    
Cleanup:
    Set objPrint = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

'*******************************************************************************************
'  Description: This subroutine shuts-down the specified "summary" view of all currently
'               running OLE Servers.

'*******************************************************************************************
Public Function ShutDownSummaryOLEServers() As Boolean
    Dim iNumItems As Integer
    Dim i As Integer
    Dim objOLEServer As Object
    Dim iSummaryInfoIndex As Integer
    Dim sbuf As String
    Dim objoleserver3 As ICHCSII_CoreAPI3
    Dim bFailedToClose    As Boolean
    
    On Error Resume Next
    
    Dim objSumMod As clsSumOLEModule
        
    For Each objSumMod In mcol
        If Not objSumMod.Module Is Nothing Then
            Call objSumMod.Module.Message(cwiOK_TO_CLOSE_VIEW, cwiPATIENT_SUMMARY_VIEW, NCID_PATIENT_SUMMARY_APPLICATION)
            DoEvents
            If Left$(UCase$(objSumMod.sOkTocloseView), 1) = "N" Then
                MsgBox "Unable to shutdown module (" & objSumMod.sApplicationName & ")"
                Exit Function
            End If
        End If  '- Not objSumMod.Module Is Nothing
    Next objSumMod
    
    For Each objSumMod In mcol
        If Not objSumMod.Module Is Nothing Then
            objSumMod.CloseModule
            
        End If   '- Not objSumMod.Module Is Nothing
    Next objSumMod
    
    ShutDownSummaryOLEServers = True
End Function

Public Property Get Item(vntIndexKey As Variant) As clsSumOLEModule
Attribute Item.VB_UserMemId = 0
  Set Item = mcol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mcol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)

    mcol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mcol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mcol = New Collection
End Sub


Private Sub Class_Terminate()
    Set mcol = Nothing
End Sub

