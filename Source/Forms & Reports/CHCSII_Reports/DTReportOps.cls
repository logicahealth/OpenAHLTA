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
Attribute VB_Name = "DTReportOps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-- Wilson9
Option Explicit
'***********************************************************************
'Name:          DisplayReport
'Description:   Display Report on the form which selected by user
'Parameters:
'   poReport                    - Report Designer holder
'   pbDisplayInSeparateWindow   - Weither display on separate window or not
'Return Value:  True, sucessful; False, otherwise
'History:
'   Date    Developer       Description
'========== =============== ============================================
'12/13/01   Wilson Sun      Added Perio Patient Summary item
'10/23/01   Wilson Sun      Added Unsigned Encounters item
'12/24/00   Wilson Sun      Created
'***********************************************************************
Private Function DisplayReport(ByVal pRepItem As ReportItem, _
                                ByVal poReport As CRAXDRT.Report, _
                                ByVal pbDisplayInSeparateWindow As Boolean) As Boolean
    
    If Not DebugMode Then On Error GoTo ErrHandler
    DisplayReport = False
    
    On Error Resume Next
    '-- Dispaly status text
    Select Case pRepItem.Report
        Case DailyProvider
            cwobjCWComm.setStatusText "Printing Daily Provider Workload"
        Case MonthlyProvider
            cwobjCWComm.setStatusText "Printing Monthly Provider Workload"
        Case UnsignedEncounters
           cwobjCWComm.setStatusText "Printing Unsigned Encounters"
        Case PerioPatientSummary
           cwobjCWComm.setStatusText "Printing Perio Patient Summary"
    End Select
        
    Static oForm As Form
    '-- Discard any saved data
    poReport.DiscardSavedData
    
    If Not oForm Is Nothing Then
        Unload oForm
        Set oForm = Nothing
    End If
    
    If pbDisplayInSeparateWindow Then
        Set oForm = New frmCrystalView
    Else
        Set oForm = New frmCrystalViewEmb
        oForm.BorderStyle = 0
        oForm.ClipControls = False
        oForm.WindowState = 2
        SetParent oForm.hwnd, frmReports.picView.hwnd
    End If
    
    '-- Display report
    With oForm.CRViewer1
        .Refresh
        .ReportSource = poReport
        .DisplayGroupTree = False
        .EnableGroupTree = False
        .DisplayTabs = False
        .EnableDrillDown = False
        .EnableRefreshButton = False
        .EnableExportButton = False 'True
        DoEvents
        .ViewReport
        DoEvents
        .Zoom 1
        DoEvents
    End With
    
    oForm.cmdExport.Visible = False
    oForm.Show
    DisplayReport = True
    
    If pRepItem.Report = DailyProvider Then
        '--- SCR 22420 - ARM Dental Provider Workflow Report - RWalker 10/21/02
        goARM.StopFarTx gsARMTx_RunReport_Dental_Provider, enmARMTxStatusOk, enmARMStop_MeOnly
    ElseIf pRepItem.Report = MonthlyProvider Then
        '--- SCR 25228 - ARM Dental Monthly Provider Workflow Report - RWalker 10/21/02
        goARM.StopFarTx gsARMTx_RunReport_Dental_Monthly_Provider, enmARMTxStatusOk, enmARMStop_MeOnly
    Else
        '10/21/02 - RWalker - Constant moved into \Core\CWShared\modARMSharedTxNames.bas and renamed
        'goARM.StopFarTx gsARM_Tx_RunReport_Generic, enmARMTxStatusOk, enmARMStop_MeOnly
        goARM.StopFarTx gsARMTx_RunReport_Generic, enmARMTxStatusOk, enmARMStop_MeOnly
    End If
    
    GoTo CleanUp

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "DTReportOps.DisplayReport", "CHCSII_Reports", vbExclamation)
CleanUp:
    Set poReport = Nothing
    Exit Function
End Function

'******************************************************************************************************
'  Sub Name: RunDentalReport
'
'  Description: Manages the data extraction and printing of a report based on
'               the report criteria object parameter values
'
'  Parameters:
'       pRepItem - Object that describes report and criteria for the report
'       poPatient - Object that used for Demorgraphic
'       pbDisplayInSeparateWindow - Determine report display seprate window or not
'  Return Value:
'       False - if an error occured
'       True  - successful
'  History
'Date       Developer   Description
'========   ==========  ===========================================
'12/13/01   Wilson Sun  Added Perio Patient Summary report. SCR 17228
'10/19/01   Wilson Sun  Added Unsigned Encounters report. SCR 13724
'12/14/00   Wilson Sun  Created
'******************************************************************************************************
Friend Function RunDentalReport(ByVal pRepItem As ReportItem, _
                                ByVal poPatient As Patient, _
                                ByVal pbDisplayInSeparateWindow As Boolean) As Boolean
    Dim oReport As CRAXDRT.Report
    Dim oFormula As FormulaFieldDefinition
    'Dim oClsDaily As New DTRptProvider
    'Dim oClsUnsignedEnc As New DTUnsignedEncounters
    Dim oClsTemplate As Object

    Dim rsRptData As ADODB.Recordset

    If Not DebugMode Then On Error GoTo ErrHandler
        
    Screen.MousePointer = vbHourglass
    RunDentalReport = False
    Set rsRptData = Nothing
    Select Case pRepItem.Report
        Case DailyProvider
            Set oReport = New CRDailyProvider
            Set oClsTemplate = New DTRptProvider
        Case MonthlyProvider
            Set oReport = New CRMonthlyProvider
            Set oClsTemplate = New DTRptProvider
            oReport.fldRptRange.SetText Format(frmReports.dtpFrom(4).Value, "MM/01/YYYY") & " - " & _
                                    Format(frmReports.dtpTo(4).Value, "MM/DD/YYYY")
                                    
'            oReport.fldRptRange.SetText "TEST"
        Case UnsignedEncounters
            Set oReport = New CRUnsignedEncounters
            Set oClsTemplate = New DTUnsignedEncounters
        Case PerioPatientSummary
            Set oReport = New CRPerioPatientSummary
            Set oClsTemplate = New DTPerioPatientSummary
    End Select
    
    If Not gobjshared.UseSQLServer Or gobjshared.IsAppMode(modeCHCSI_GUI) Then
        For Each oFormula In oReport.FormulaFields
            If oFormula.FormulaFieldName = "TrainingUseOnly" Then
               oFormula.Text = ""
            End If
        Next
    End If
    
    '-----------------------------------------------------------------
'    Select Case pRepItem.Report
'        Case ReportType.DailyProvider, ReportType.MonthlyProvider
'            If Not oClsDaily.BuildReport(pRepItem, rsRptData) Then GoTo CleanUp
'        Case ReportType.UnsignedEncounters
'            If Not oClsTemplate.BuildReport(pRepItem, rsRptData) Then GoTo CleanUp
'        Case ReportType.PerioPatientSummary
'            'If Not oClsPerioPatientSummary.BuildReport(pRepItem, rsRptData) Then GoTo CleanUp
'    End Select
    If Not oClsTemplate.BuildReport(pRepItem, rsRptData) Then GoTo CleanUp
    
    If rsRptData Is Nothing Then GoTo CleanUp
    oReport.Database.SetDataSource rsRptData
    oReport.ReadRecords
    
    '-----------------------------------------------------------------
    'Show report
    If Not DisplayReport(pRepItem, oReport, pbDisplayInSeparateWindow) Then GoTo CleanUp
    
    RunDentalReport = True

GoTo CleanUp
ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "DTReportOps.RunDentalReport", "CHCSII_Reports", vbExclamation)
CleanUp:
    cwobjCWComm.setStatusText ""
    'Set oDentalForms = Nothing
    'Set oReportOnForm = Nothing
    Set oReport = Nothing
    'Set oClsDaily = Nothing
    Set oClsTemplate = Nothing
    Set rsRptData = Nothing
    Screen.MousePointer = vbDefault
    Exit Function
End Function

Private Sub Class_Terminate()
    On Error Resume Next
    Dim fs As New Scripting.FileSystemObject
    '-- Delete VB*.tmp files
    fs.DeleteFile App.Path & "\VB*.tmp"
    Set fs = Nothing
End Sub
