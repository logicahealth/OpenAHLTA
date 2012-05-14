VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Crystal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const sMODULE_NAME As String = "Crystal.cls"
Private moReportApp As CRAXDRT.Application
Public Function ReportOpen(sFormFileName As String, oReport As CRAXDRT.Report) As Boolean

    Dim oFormula As FormulaFieldDefinition
    On Error GoTo ErrHandler
    ReportOpen = False
    
    '-- Open Report File ---------------------------------------------------------------------
    Set oReport = moReportApp.OpenReport(gsFormPath & sFormFileName, 1)
    If Not gobjshared.UseSQLServer Or gobjshared.IsAppMode(modeCHCSI_GUI) Then
        For Each oFormula In oReport.FormulaFields
            If oFormula.FormulaFieldName = "TrainingUseOnly" Then
               oFormula.Text = ""
            End If
        Next
    End If
    
    ReportOpen = True
    
    GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & "ReportOpen", App.EXEName, vbExclamation
CleanUp:

End Function

Public Function AssignFormulas(oCollFormulas As Collection, _
                                oReport As CRAXDRT.Report) As Boolean

On Error GoTo ErrHandler
AssignFormulas = False

Dim i As Integer

    '-- Set Formula Fields
    '-----------------------------------------------------------------------
    If Not oCollFormulas Is Nothing Then
        For i = 1 To oReport.FormulaFields.Count
            On Error Resume Next
            oReport.FormulaFields.Item(i).Text = _
                oCollFormulas.Item(oReport.FormulaFields.Item(i).FormulaFieldName)
               If Err.Number <> 0 Then
                  If Err <> 5 Then
                     On Error GoTo ErrHandler
                     Err.Raise Err.Number
                  Else
                     Err.Clear
                  End If
               End If
        Next
    End If

AssignFormulas = True
GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & "AssignFormulas", App.EXEName, vbExclamation
CleanUp:

End Function

Public Function AssignRecordSets(oCollRecSet As Collection, _
                                oReport As CRAXDRT.Report) As Boolean

On Error GoTo ErrHandler
AssignRecordSets = False

Dim oSubReport As CRAXDRT.SubreportObject

Dim i As Integer
Dim i2 As Integer
Dim i3 As Integer


'-- Set RecordSets in Report
'-----------------------------------------------------------------------
If oCollRecSet.Count > 0 Then

    '-- Set RecordSets in Main part of Report
    '----------------------------------------
    For i = 1 To oReport.Database.Tables.Count
        On Error Resume Next
        oReport.Database.SetDataSource oCollRecSet(oReport.Database.Tables(i).Name), 3, i
        If Err <> 5 And Err <> 0 Then
            On Error GoTo ErrHandler
            Err.Raise Err
        End If
    Next

    On Error GoTo ErrHandler

    '-- Set RecordSets in subreports
    '-------------------------------
    For i = 1 To oReport.Sections.Count
        For i2 = 1 To oReport.Sections.Item(i).ReportObjects.Count
            If TypeOf oReport.Sections.Item(i).ReportObjects.Item(i2) Is SubreportObject Then
                Set oSubReport = oReport.Sections.Item(i).ReportObjects.Item(i2)
                For i3 = 1 To oReport.OpenSubreport(oSubReport.SubreportName).Database.Tables.Count
                    On Error Resume Next
                    oReport.OpenSubreport(oSubReport.SubreportName).Database.SetDataSource _
                        oCollRecSet(oReport.OpenSubreport(oSubReport.SubreportName).Database.Tables(i3).Name), 3, i3
                    If Err <> 5 And Err <> 0 Then
                        On Error GoTo ErrHandler
                        Err.Raise Err
                    End If
                Next
                On Error GoTo ErrHandler
            End If
        Next
    Next
    
End If

AssignRecordSets = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & "AssignRecordSets", App.EXEName, vbExclamation
CleanUp:
    Set oSubReport = Nothing
End Function

Public Function PrintReport(oReport As CRAXDRT.Report) As Boolean

    On Error GoTo ErrHandler
    PrintReport = False
    
    '-- Print Form ----------------------------------------------------------------------------
    If Printers.Count = 0 Then
       MsgBxARMd "No printers are installed. To install a printer, point to Settings on the Windows Start menu, " & _
                    "click Printers, and then double-click Add Printer. Follow the instructions in the wizard.", vbOKOnly + vbExclamation
    Else
       oReport.PrintOut False
       PrintReport = True
    End If

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & "PrintReport", App.EXEName, vbExclamation
CleanUp:

End Function
Public Function DisplayReport(oReport As CRAXDRT.Report) As Boolean

    Dim oForm As frmCrystalView
    
    On Error GoTo ErrHandler
    DisplayReport = False
    
    '-- Preview Form    '----------------------------------------------------------------------------
    Set oForm = New frmCrystalView
    oForm.CRViewer1.ReportSource = oReport
    oForm.CRViewer1.DisplayGroupTree = False
    oForm.CRViewer1.EnableGroupTree = False
    oForm.CRViewer1.DisplayTabs = False
    oForm.CRViewer1.EnableDrillDown = False
    oForm.CRViewer1.EnableRefreshButton = True
    oForm.CRViewer1.EnableCloseButton = False
    oForm.CRViewer1.ViewReport
    
    Screen.MousePointer = vbDefault
    DoEvents
    oForm.Timer1.Enabled = True
    DoEvents
    oForm.Show vbModal
    DoEvents
    
    DisplayReport = True
    
    GoTo CleanUp
ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & "DisplayReport", App.EXEName, vbExclamation
CleanUp:
    Set oForm = Nothing  'hwcc 2/19/2003 - SCR 32681, someone forget to release this reference
End Function

Private Sub Class_Initialize()
    Set moReportApp = New CRAXDRT.Application
End Sub

Private Sub Class_Terminate()
    Set moReportApp = Nothing
End Sub


