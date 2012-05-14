VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Forms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private Const sMODULE_NAME As String = "Forms.cls"

'-- Public
'-------------------------------------------------
Public Enum PrintOptions
    PrintOnly = 1
    PrintPreview = 2
    ReturnReportObject = 3
End Enum

Public ReturnObject As Object

'-- Private
'-------------------------------------------------
Private moCrystal As Crystal
Private moReport As CRAXDRT.Report

Private mbDatabaseInitialized As Boolean '-- Out

Private Const cDBName = "CHCSII.MDB" '-- Report database Name (Access)
Private Const RepTuxServ = "DSQL_REP"
Private Const sDefault_Caption As String = "AHLTA-T Print"

Public Enum LabelTabs
    LabSpecimen = 0
    RadFileFolder = 1
    RadFilm = 2
    PharmacyIV = 3
    PharmacyOutPat = 4
    GenericLabel = 5
    LastKnown = 99
End Enum

' Created when the form class is created,
'  destroyed when the form cls is destroyed

Private Sub CreateDBAccessObjs()
On Error GoTo CreateDBAccessObjsError
    'Clr everything
    Call DestroyDBAccessObjs
    
    Set goDBConn = New CHCSII_CONN.Conn
    Set goDBSqlOps = goDBConn.CHCSII_SQLOPS(Auto)
    Set goDBDAS = goDBConn.CHCSII_DAS(Auto)
    Set goDBSqlOpsEx = goDBConn.CHCSII_SQLOPS_EX(Auto)
    
    GoTo ExitCreateDBAccessObjs
    
CreateDBAccessObjsError:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".CreateDBAccessObjs", App.EXEName, vbExclamation
    'Clean up if there is an error
    Call DestroyDBAccessObjs
ExitCreateDBAccessObjs:
End Sub

Private Sub DestroyDBAccessObjs()
    'Clean house
    Set goDBConn = Nothing
    Set goDBSqlOps = Nothing
    Set goDBDAS = Nothing
    Set goDBSqlOpsEx = Nothing
End Sub

Private Function CheckCWCommObj() As Boolean
    CheckCWCommObj = False
    If cwobjCWComm Is Nothing Then
       'Err.Raise vbObjectError, sMODULE_NAME, "Function caller does not set MMMHISComm.Comm object required by Forms."
       Exit Function
    End If
    CheckCWCommObj = True
End Function

Private Sub InformCoreShowModal()
    If cwobjCWComm Is Nothing Then Exit Sub
    cwobjCWComm.Message cwiSHOWING_MODAL_FORM, vbNullString, NCID_CW_APPLICATION, NCID_GENERAL_REPORTS_APPLICATION
End Sub

Private Sub InformCoreDoneShowModal()
    If cwobjCWComm Is Nothing Then Exit Sub
    cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, vbNullString, NCID_CW_APPLICATION, NCID_GENERAL_REPORTS_APPLICATION
End Sub

Public Function IsCommObjSet() As Boolean
    IsCommObjSet = False
    If cwobjCWComm Is Nothing Then Exit Function
    IsCommObjSet = True
End Function

' moved objUser to 3rd arg from last to make it required (not optional)
Public Function printCHCSII_Labels(iPrintOption As PrintOptions, _
                            sPatientUnitNumber As String, _
                            objUser As ICHCSII_User, _
                            Optional StartingTab As LabelTabs = LastKnown, _
                            Optional sStartListValue As String = vbNullString, _
                            Optional CallingFromInptOrdrMgt As Boolean = False, _
                            Optional PrintLabelOnly As Boolean = False, _
                            Optional CreateNewRx As Boolean = False) As Boolean
  
    Dim oLabel As FormCHCSII_Labels
    
    On Error GoTo ErrHandler
    
    printCHCSII_Labels = False
    
    'Call CheckCWCommObj
    
    Set oLabel = New FormCHCSII_Labels
    
    'Set Patient unit number
    oLabel.PatientUnitNumber = sPatientUnitNumber
    Set oLabel.User = objUser
    
    'Set the report object
    Set oLabel.ReportObject = moReport
    
    'Unless the tab needs to be set to a particular place, set this as the 'LastKnown' and it will
    'default to the last place it was when the user exited the form
    oLabel.SetStartingTab StartingTab
    

    
    'Set this value to a nullstring unless you want to set it to a particular list entry. The lists are current populated
    'by the Order_name value from the OE_ tables.
    oLabel.SetStartingListValue sStartListValue
    
    'Build the labels.
    If Not oLabel.initLabels(, CallingFromInptOrdrMgt, CreateNewRx) Then GoTo Cleanup
      
     If PrintLabelOnly Then
        '*** PRINT LABEL AFTER DISPENSING MED FROM INPATIENT ORDER MGMT
        '*** 4/14/05 - FRANK
        
     Else
        'Display the label form
        Call InformCoreShowModal      'Notice Core to stop processing messages
        oLabel.Display
    End If
    
    printCHCSII_Labels = True
    
    GoTo Cleanup
    
ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".printCHCSII_Labels", App.EXEName, vbExclamation
Cleanup:
    'On Error Resume Next
    Call InformCoreDoneShowModal
    Set oLabel = Nothing
End Function


Public Function PrintMedLabelOnly(iPrintOption As PrintOptions, _
                            sPatientUnitNumber As String, Optional StartingTab As LabelTabs = LastKnown, _
                            Optional sStartListValue As String = vbNullString, _
                            Optional CallingFromInptOrdrMgt As Boolean = False, _
                            Optional PrintLabelOnly As Boolean = False, _
                            Optional LabelPrinter As String, _
                            Optional objUser As ICHCSII_User) As Boolean
  
    Dim oLabel As FormCHCSII_Labels
    
    On Error GoTo ErrHandler
    
    PrintMedLabelOnly = False
    
    'Call CheckCWCommObj
    
    Set oLabel = New FormCHCSII_Labels
    
    'Set Patient unit number
    oLabel.PatientUnitNumber = sPatientUnitNumber
    
    'Set the report object
    Set oLabel.ReportObject = moReport
    Set oLabel.User = objUser
    
    'Unless the tab needs to be set to a particular place, set this as the 'LastKnown' and it will
    'default to the last place it was when the user exited the form
    oLabel.SetStartingTab StartingTab
    
    
    'Set this value to a nullstring unless you want to set it to a particular list entry. The lists are current populated
    'by the Order_name value from the OE_ tables.
    oLabel.SetStartingListValue sStartListValue
    
    'Build the labels.
    If Not oLabel.initLabels(, CallingFromInptOrdrMgt) Then GoTo Cleanup
      
     If PrintLabelOnly Then
        '*** PRINT LABEL AFTER DISPENSING MED FROM INPATIENT ORDER MGMT
        '*** 4/14/05 - FRANK
        If Len(LabelPrinter) > 0 Then
            Call frmCHCSII_Labels.SetPrinterDefault(LabelPrinter)
        End If
        frmCHCSII_Labels.cmdPrint.Value = True
     Else
        'Display the label form
        Call InformCoreShowModal      'Notice Core to stop processing messages
        oLabel.Display
    End If
    
    PrintMedLabelOnly = True
    
    GoTo Cleanup
    
ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".PrintMedLabelOnly", App.EXEName, vbExclamation
Cleanup:
    'On Error Resume Next
    Call InformCoreDoneShowModal
    Set oLabel = Nothing

End Function


Public Function GetPatientIEN(pUnit_Number As String) As String
        
    Dim sSQL As String
    
    GetPatientIEN = vbNullString
    If Not gobjshared.AppMode = modeCHCSI_GUI Then Exit Function
    sSQL = "SELECT MEDICAL_RECORD_NUMBER from mmi_site where unit_number = " & pUnit_Number
    goDBSqlOps.Execute sSQL
    If Not (goDBSqlOps.BOF And goDBSqlOps.EOF) Then
       goDBSqlOps.MoveFirst
       GetPatientIEN = goDBSqlOps("MEDICAL_RECORD_NUMBER")
    End If
    
End Function

Private Function InitializeCrystal() As Boolean
'-- This routine is executed for the exclusive purpose of handling a Crystal 8 Bug

    Dim oCrystal As New Crystal
    Dim oForm As frmCrystalView
    
    On Error GoTo ErrHandler
    If bCrystalInitialized Then
       InitializeCrystal = True
       Exit Function
    End If
    
    bCrystalInitialized = True
    InitializeCrystal = False
               
    Set moReport = Nothing
    If Not oCrystal.ReportOpen("Initialize.rpt", moReport) Then GoTo Cleanup
    Set oForm = New frmCrystalView
    
    On Error Resume Next
    oForm.CRViewer1.Refresh
    
    On Error GoTo ErrHandler
    
    oForm.CRViewer1.ReportSource = moReport
    oForm.CRViewer1.DisplayGroupTree = False
    oForm.CRViewer1.EnableGroupTree = False
    oForm.CRViewer1.DisplayTabs = False
    oForm.CRViewer1.EnableDrillDown = False
    oForm.CRViewer1.EnableRefreshButton = False
    oForm.CRViewer1.EnableExportButton = True
    
    DoEvents
    oForm.CRViewer1.ViewReport
    DoEvents
    oForm.CRViewer1.Zoom 1
    DoEvents
        
    InitializeCrystal = True
    
    GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".InitializeCrystal", "CHCSII_Reports", vbExclamation
Cleanup:
    Set oForm = Nothing
    Set oCrystal = Nothing

End Function

'This function is under construction,
' so we are not exposing the interface
Private Function PrintDD2322(ByRef poPatient As Patient, _
                            ByVal psLocalCaseNo As String, _
                            ByVal psClinicName As String, _
                            ByVal psClinicNCID As String, _
                            ByVal psProviderName As String, _
                            ByRef prsSubReport As ADODB.Recordset, _
                            Optional ByVal pbPreview As Boolean = True) As Boolean
'Exposing Dental function
On Error GoTo ErrHandler
    PrintDD2322 = False
    Dim ofrm As CHCSII_Forms.DentalForms
    Set ofrm = New CHCSII_Forms.DentalForms
    PrintDD2322 = ofrm.PrintDD2322(poPatient, psLocalCaseNo, psClinicName, psClinicNCID, psProviderName, prsSubReport, pbPreview)
    GoTo Cleanup:
ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".PrintDD2322", App.EXEName, vbExclamation
Cleanup:
    Set ofrm = Nothing
End Function

'This function is under construction,
' so we are not exposing the interface
Private Function PrintSF603A(ByRef poPatient As Patient, _
                            ByRef poEncounter As Object, _
                            ByRef prsReport As ADODB.Recordset, _
                            ByRef pcolTeethDx As Collection, _
                            ByRef pcolTeethTx As Collection, _
                            Optional ByVal pbPreview As Boolean = True) As Boolean
'Exposing Dental function
On Error GoTo ErrHandler
    PrintSF603A = False
    Dim ofrm As CHCSII_Forms.DentalForms
    Set ofrm = New CHCSII_Forms.DentalForms
    PrintSF603A = ofrm.PrintSF603A2(poPatient, poEncounter, prsReport, pcolTeethDx, pcolTeethTx, pbPreview)
    GoTo Cleanup:
ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".PrintSF603A", App.EXEName, vbExclamation
Cleanup:
    Set ofrm = Nothing
End Function
    
Public Function printInsuranceForm(iPrintOption As PrintOptions, _
                                   sPatientUnitNumber As String, _
                                   sFacilityNCID As String, Optional pEncounterNumber As String = vbNullString) As Boolean

    Dim oInsurance As FormInsurance
    
On Error GoTo ErrHandler
    printInsuranceForm = False
    'Screen.MousePointer = vbHourglass
    
    Set oInsurance = New FormInsurance
    If Not oInsurance.BuildInsurance(sPatientUnitNumber, sFacilityNCID, moReport, pEncounterNumber) Then GoTo Cleanup
    
    Call InformCoreShowModal
    If iPrintOption = PrintOnly Then
        If Not moCrystal.PrintReport(moReport) Then GoTo Cleanup
    ElseIf iPrintOption = PrintPreview Then
        If Not moCrystal.DisplayReport(moReport) Then GoTo Cleanup
    End If

    printInsuranceForm = True
    GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".printInsuranceForm", App.EXEName, vbExclamation
Cleanup:
    'Screen.MousePointer = vbDefault
    Call InformCoreDoneShowModal
    Set oInsurance = Nothing
End Function

Public Function printSF600(oEnc As Object, _
                           sRTF As String, _
                           Optional bPreview As Boolean, Optional bSkipPrinterDialog As Boolean = False) As Boolean

    Dim sSQL As String
    Dim sPatientIEN As String
    Dim oFormSF600 As FormSF600
    
On Error GoTo ErrHandler
    
    printSF600 = False
    
    '
    'Call CheckCWCommObj

    Set oFormSF600 = New FormSF600
    
    'hwcc 2/11/2003 - this is throw away code when GREEN switch no longer valid
    If oEnc.GetCommObj.CmdLineSwitch("GREEN") <> vbNullString Then
        oFormSF600.mbGreenSwitch = True
    Else
        oFormSF600.mbGreenSwitch = False
    End If
    'end throw away code
    
    If gobjshared.AppMode = modeCHCSI_GUI Then
        If oEnc.Patient_Ien = vbNullString Then
           sPatientIEN = GetPatientIEN(oEnc.PatientID)
        Else
            sPatientIEN = oEnc.Patient_Ien
        End If
        Call InformCoreShowModal      'Notice Core to stop processing messages
        If Not oFormSF600.BuildSF600(sPatientIEN, oEnc.FacilityNCID, sRTF, bPreview, bSkipPrinterDialog, oEnc) Then GoTo Cleanup
    Else
        Call InformCoreShowModal      'Notice Core to stop processing messages
        If Not oFormSF600.BuildSF600(oEnc.PatientID, oEnc.FacilityNCID, sRTF, bPreview, bSkipPrinterDialog, oEnc) Then GoTo Cleanup
    End If
    printSF600 = True

    GoTo Cleanup
ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".printSF600", App.EXEName, vbExclamation
Cleanup:
    Call InformCoreDoneShowModal     'Notice Core to accept messages again
    Set oFormSF600 = Nothing
End Function

Public Function printQuestionnaire(sRTF As String, Optional bPreview As Boolean, _
                           Optional iNumberOfCopies As Integer = 1) As Boolean

On Error GoTo ErrHandler
    printQuestionnaire = False
    
    Dim oFormQ As FormQuestionnaire
    
    Set oFormQ = New FormQuestionnaire
    Call InformCoreShowModal
    If Not oFormQ.Build(sRTF, "Questionnaire Preview", bPreview, iNumberOfCopies) Then GoTo Cleanup
    
    printQuestionnaire = True
    GoTo Cleanup
    
ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".printQuestionnaire", App.EXEName, vbExclamation
Cleanup:
    Call InformCoreDoneShowModal
    Set oFormQ = Nothing
End Function

Public Function printRTF(sRTF As String, Optional sTitle As String = vbNullString, Optional bPreview As Boolean = False, Optional iNumberOfCopies As Integer = 1) As Boolean

    Dim iScreenMousepointer As Integer
    Dim strPrinterName As String
    
    On Error GoTo ErrHandler
    printRTF = False
    iScreenMousepointer = Screen.MousePointer
    Screen.MousePointer = vbDefault
    
    Dim oFormRTF As FormRTF
    
    '*** 6/11/05 - Frank
    '*** Ensure that the Default Printer is used (not the Label Printer)
    If LenB(gobjshared.CmdLineSwitches("INPATIENT2")) > 0 Then
        strPrinterName = GetDefaultPrinterDeviceName
        If strPrinterName = "" Then
            'No Default printer Specified.
            frmCHCSII_Labels.FindPrinter
        Else
            Call frmCHCSII_Labels.SetPrinterDefault(GetDefaultPrinterDeviceName)
        End If

    End If
    
    Set oFormRTF = New FormRTF
    Call InformCoreShowModal
    
    
    If Not oFormRTF.Build(sRTF, sTitle, bPreview, iNumberOfCopies) Then GoTo Cleanup
    
    printRTF = True
    GoTo Cleanup
    
ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".printRTF", App.EXEName, vbExclamation
Cleanup:
    Call InformCoreDoneShowModal
    Set oFormRTF = Nothing
    Screen.MousePointer = iScreenMousepointer
End Function

Public Function printDD2161(iPrintOption As PrintOptions, _
                           oConsultFormItem As ConsultFormItem, _
                           sPatientUnitNumber As String) As Boolean
    Dim oFormDD2161 As FormDD2161
On Error GoTo ErrHandler
    printDD2161 = False
    
    Screen.MousePointer = vbHourglass
    Set oFormDD2161 = New FormDD2161
    
    If Not oFormDD2161.BuildDD2161(oConsultFormItem, sPatientUnitNumber, moReport) Then GoTo Cleanup
    
    Call InformCoreShowModal
    If iPrintOption = PrintOnly Then
        If Not moCrystal.PrintReport(moReport) Then GoTo Cleanup
    ElseIf iPrintOption = PrintPreview Then
        If Not moCrystal.DisplayReport(moReport) Then GoTo Cleanup
    End If
    
    printDD2161 = True
    GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".printDD2161", App.EXEName, vbExclamation
Cleanup:
    Call InformCoreDoneShowModal
    Set oFormDD2161 = Nothing
End Function

Public Function printDD2161Complete(iPrintOption As PrintOptions, _
                                   Consult_Id As String, _
                                   ConsultationReportText As String) As Boolean
    Dim oFormDD2161 As FormDD2161

On Error GoTo ErrHandler

    printDD2161Complete = False
    
    Screen.MousePointer = vbHourglass
    
    Set oFormDD2161 = New FormDD2161
    
    If Not oFormDD2161.BuildDD2161Complete(Consult_Id, ConsultationReportText, moReport) Then GoTo Cleanup
    
    Call InformCoreDoneShowModal
    If iPrintOption = PrintOnly Then
        If Not moCrystal.PrintReport(moReport) Then GoTo Cleanup
    ElseIf iPrintOption = PrintPreview Then
        If Not moCrystal.DisplayReport(moReport) Then GoTo Cleanup
    End If
    
    printDD2161Complete = True
    GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".printDD2161Complete", App.EXEName, vbExclamation
Cleanup:
    Call InformCoreDoneShowModal
    Set oFormDD2161 = Nothing
End Function


Public Function printCPG_POC(iPrintOption As PrintOptions, sPatientUnitNumber As String, sFacilityNumber As String) As Boolean

    Dim oFormRTF As FormSF600
    Dim sRTF As String
    Dim sRTF_Header As String
    
On Error GoTo ErrHandler
    Screen.MousePointer = vbHourglass
    printCPG_POC = False
    
    Set oFormRTF = New FormSF600
    
    Dim oPatientPOC As Object
    
    Set oPatientPOC = cwobjCWComm.InitializeOLEServer(NCID_PATIENT_PLANOFCARE)
    sRTF = oPatientPOC.GetEncounterContRTF
        
    sRTF_Header = "{\rtf1\ansi\ucl\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}"
    sRTF_Header = sRTF_Header & "{\f1\froman\fcharset2 Symbol;}"
    sRTF_Header = sRTF_Header & "{\f2\ftech\fprq2 Courier New;}{\f3\froman\fprq2 Times New Roman;}}"
    sRTF_Header = sRTF_Header & "{\colortbl\red0\green0\blue0;}\deflang1033\horzdoc{\*\fchars }{\*\lchars }\plain\f2\fs16 \b\ul"
    sRTF = sRTF_Header & sRTF & " \pard\par }"
    
    Call InformCoreShowModal
    If iPrintOption = PrintOnly Then
       oFormRTF.BuildSF600 sPatientUnitNumber, sFacilityNumber, sRTF, False, False
    ElseIf iPrintOption = PrintPreview Then
       oFormRTF.BuildSF600 sPatientUnitNumber, sFacilityNumber, sRTF, True, False
    End If
    
    printCPG_POC = True
    GoTo Cleanup
    
ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".printCPG_POC", App.EXEName, vbExclamation
Cleanup:
    Call InformCoreDoneShowModal
    Set oFormRTF = Nothing
    Screen.MousePointer = vbDefault
End Function

Public Function printScreening_Notifications(iPrintOption As PrintOptions, _
                             cRecordsets As Collection) As Boolean
    
    Dim oForm As FormScreeningNotifications
    
On Error GoTo ErrHandler
    Screen.MousePointer = vbHourglass
    printScreening_Notifications = False
    gbDocumentPrinted = False
    
    Set oForm = New FormScreeningNotifications
    If Not oForm.Build(moReport, cRecordsets) Then GoTo Cleanup
    
    Call InformCoreShowModal
    If iPrintOption = PrintOnly Then
        If Not moCrystal.PrintReport(moReport) Then GoTo Cleanup
    ElseIf iPrintOption = PrintPreview Then
        gbBypassDialog = True
        If Not moCrystal.DisplayReport(moReport) Then GoTo Cleanup
        If gbDocumentPrinted Then
           If Not moCrystal.PrintReport(moReport) Then GoTo Cleanup
        End If
    End If
    
    printScreening_Notifications = True
    GoTo Cleanup
    
ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".printScreening_Notifications", App.EXEName, vbExclamation
Cleanup:
    Call InformCoreDoneShowModal
    gbBypassDialog = False
    Set oForm = Nothing
    Screen.MousePointer = vbDefault
End Function

Public Function printScreening_PCMList(iPrintOption As PrintOptions, _
                             rsPCMCounts As ADODB.Recordset) As Boolean
    
    Dim oForm As FormScreeningNotifications
    
On Error GoTo ErrHandler
    Screen.MousePointer = vbHourglass
    printScreening_PCMList = False
    
    Set oForm = New FormScreeningNotifications
    If Not oForm.BuildMTF_PCM_Due_List(moReport, rsPCMCounts) Then GoTo Cleanup
    
    Call InformCoreShowModal
    If iPrintOption = PrintOnly Then
        If Not moCrystal.PrintReport(moReport) Then GoTo Cleanup
    ElseIf iPrintOption = PrintPreview Then
        If Not moCrystal.DisplayReport(moReport) Then GoTo Cleanup
    End If
    
    printScreening_PCMList = True
    GoTo Cleanup
    
ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".printScreening_PCMList", App.EXEName, vbExclamation

Cleanup:
    Call InformCoreDoneShowModal
    Set oForm = Nothing
    Screen.MousePointer = vbDefault
End Function

Public Function printScreening_Notif_Envelopers(iPrintOption As PrintOptions, _
                             rsDemographics As ADODB.Recordset) As Boolean
    
    Dim oForm As FormScreeningNotifications
    
On Error GoTo ErrHandler
    Screen.MousePointer = vbHourglass
    printScreening_Notif_Envelopers = False
    
    Set oForm = New FormScreeningNotifications
    If Not oForm.BuildEnvelopes(moReport, rsDemographics) Then GoTo Cleanup
    
    Call InformCoreShowModal
    If iPrintOption = PrintOnly Then
        If Not moCrystal.PrintReport(moReport) Then GoTo Cleanup
    ElseIf iPrintOption = PrintPreview Then
        If Not moCrystal.DisplayReport(moReport) Then GoTo Cleanup
    End If
    
    printScreening_Notif_Envelopers = True
    GoTo Cleanup
    
ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".printScreening_Notif_Envelopers", App.EXEName, vbExclamation

Cleanup:
    Call InformCoreDoneShowModal
    Set oForm = Nothing
    Screen.MousePointer = vbDefault
        
End Function

Public Function printScreening_Notif_Labels(iPrintOption As PrintOptions, _
                             rsDemographics As ADODB.Recordset) As Boolean
    Dim oForm As FormScreeningNotifications
On Error GoTo ErrHandler
    Screen.MousePointer = vbHourglass
    printScreening_Notif_Labels = False
    
    Set oForm = New FormScreeningNotifications
    If Not oForm.BuildAddressLabels(moReport, rsDemographics) Then GoTo Cleanup
    
    Call InformCoreShowModal
    If iPrintOption = PrintOnly Then
        If Not moCrystal.PrintReport(moReport) Then GoTo Cleanup
    ElseIf iPrintOption = PrintPreview Then
        If Not moCrystal.DisplayReport(moReport) Then GoTo Cleanup
    End If
    
    printScreening_Notif_Labels = True
    GoTo Cleanup
    
ErrHandler:
        gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".printScreening_Notif_Labels", App.EXEName, vbExclamation

Cleanup:
    Call InformCoreDoneShowModal
    Set oForm = Nothing
    Screen.MousePointer = vbDefault
        
End Function

Public Function printDD2766(iPrintOption As PrintOptions, _
                            sPatientUnitNumber As String) As Boolean
On Error GoTo ErrHandler
      
    printDD2766 = False
    
    Dim oDD2766 As FormDD2766
    
    Set oDD2766 = New FormDD2766
    
    If Not oDD2766.BuildDD2766(sPatientUnitNumber, moReport) Then GoTo Cleanup
        
    Call InformCoreShowModal
    If iPrintOption = PrintOnly Then
        If Not moCrystal.PrintReport(moReport) Then GoTo Cleanup
    ElseIf iPrintOption = PrintPreview Then
        If Not moCrystal.DisplayReport(moReport) Then GoTo Cleanup
    End If
    
    printDD2766 = True
    GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".printDD2766", App.EXEName, vbExclamation
Cleanup:
    Call InformCoreDoneShowModal
    Set oDD2766 = Nothing
    Set moReport = Nothing
    Screen.MousePointer = vbDefault
End Function

Public Function printDD2813(iPrintOption As PrintOptions, _
                            sPatientUnitNumber As String) As Boolean
On Error GoTo ErrHandler

    printDD2813 = False
    
    Dim oDD2813 As Formdd2813
    
    Set oDD2813 = New Formdd2813
    If Not oDD2813.BuildDD2813(sPatientUnitNumber, moReport) Then GoTo Cleanup
    
    Call InformCoreShowModal
    If iPrintOption = PrintOnly Then
        If Not moCrystal.PrintReport(moReport) Then GoTo Cleanup
    ElseIf iPrintOption = PrintPreview Then
        If Not moCrystal.DisplayReport(moReport) Then GoTo Cleanup
    End If
    
    printDD2813 = True
    GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".printDD2813", App.EXEName, vbExclamation
Cleanup:
    Call InformCoreDoneShowModal
    Screen.MousePointer = vbDefault
    Set oDD2813 = Nothing
End Function

Public Function printDD689(iPrintOption As PrintOptions, _
                           sPatientUnitNumber As String) As Boolean
On Error GoTo ErrHandler

    printDD689 = False
    
    Dim oDD689 As FormDD689
    
    Set oDD689 = New FormDD689
    If Not oDD689.BuildDD689(sPatientUnitNumber, moReport) Then GoTo Cleanup
    
    Call InformCoreShowModal
    If iPrintOption = PrintOnly Then
        If Not moCrystal.PrintReport(moReport) Then GoTo Cleanup
    ElseIf iPrintOption = PrintPreview Then
        If Not moCrystal.DisplayReport(moReport) Then GoTo Cleanup
    End If
    
    printDD689 = True
    GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".printDD689", App.EXEName, vbExclamation
Cleanup:
    Call InformCoreDoneShowModal
    Screen.MousePointer = vbDefault
    Set oDD689 = Nothing
End Function

Public Function printDD741(iPrintOption As PrintOptions, _
                           sPatientUnitNumber As String) As Boolean
On Error GoTo ErrHandler

    printDD741 = False
    
    Dim oDD741 As FormDD741
    
    Set oDD741 = New FormDD741
    If Not oDD741.BuildDD741(sPatientUnitNumber, moReport) Then GoTo Cleanup
    
    Call InformCoreShowModal
    If iPrintOption = PrintOnly Then
        If Not moCrystal.PrintReport(moReport) Then GoTo Cleanup
    ElseIf iPrintOption = PrintPreview Then
        If Not moCrystal.DisplayReport(moReport) Then GoTo Cleanup
    End If
    
    printDD741 = True
    GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".printDD741", App.EXEName, vbExclamation
Cleanup:
    Call InformCoreDoneShowModal
    Screen.MousePointer = vbDefault
    Set oDD741 = Nothing
End Function

Public Function printDD771(iPrintOption As PrintOptions, _
                           collOrderIds As Collection) As Boolean
On Error GoTo ErrHandler

    printDD771 = False
    
    Dim oDD771 As FormDD771
    
    Set oDD771 = New FormDD771
    If Not oDD771.BuildDD771(collOrderIds, "dd771.rpt", moReport) Then GoTo Cleanup
    
    Call InformCoreShowModal
    If iPrintOption = PrintOnly Then
        If Not moCrystal.PrintReport(moReport) Then GoTo Cleanup
    ElseIf iPrintOption = PrintPreview Then
        If Not moCrystal.DisplayReport(moReport) Then GoTo Cleanup
    End If
    
    printDD771 = True
    GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".printDD771", App.EXEName, vbExclamation
Cleanup:
    Call InformCoreDoneShowModal
    Screen.MousePointer = vbDefault
    Set oDD771 = Nothing
End Function

Public Function printDD771Double(iPrintOption As PrintOptions, _
                                 collOrderIds As Collection) As Boolean
On Error GoTo ErrHandler

    printDD771Double = False
    
    Dim oDD771 As FormDD771
    
    Set oDD771 = New FormDD771
    If Not oDD771.BuildDD771(collOrderIds, "dd771_double.rpt", moReport) Then GoTo Cleanup
    
    Call InformCoreShowModal
    If iPrintOption = PrintOnly Then
        If Not moCrystal.PrintReport(moReport) Then GoTo Cleanup
    ElseIf iPrintOption = PrintPreview Then
        If Not moCrystal.DisplayReport(moReport) Then GoTo Cleanup
    End If
    
    printDD771Double = True
    GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".printDD771Double", App.EXEName, vbExclamation
Cleanup:
    Call InformCoreDoneShowModal
    Screen.MousePointer = vbDefault
    Set oDD771 = Nothing
End Function

Public Function printSF522(iPrintOption As PrintOptions, _
                           sPatientUnitNumber As String) As Boolean

On Error GoTo ErrHandler

    printSF522 = False
    
    Dim oSF522 As FormSF522
    
    Set oSF522 = New FormSF522
    If Not oSF522.BuildSF522(sPatientUnitNumber, moReport) Then GoTo Cleanup
    
    Call InformCoreShowModal
    If iPrintOption = PrintOnly Then
        If Not moCrystal.PrintReport(moReport) Then GoTo Cleanup
    ElseIf iPrintOption = PrintPreview Then
        If Not moCrystal.DisplayReport(moReport) Then GoTo Cleanup
    End If
    
    printSF522 = True
    GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".printSF522", App.EXEName, vbExclamation
Cleanup:
    Call InformCoreDoneShowModal
    Screen.MousePointer = vbDefault
    Set oSF522 = Nothing
End Function

Public Function printPatientCheckOut(iPrintOption As PrintOptions, _
                                     sEncounterNumber As String, _
                                    sPatientUnitNumber As String) As Boolean
On Error GoTo ErrHandler

    printPatientCheckOut = False
    
    Screen.MousePointer = vbHourglass
    
    Dim oFormCheckOut As FormCheckOut
    
    Set oFormCheckOut = New FormCheckOut
    
    If Not oFormCheckOut.BuildPatientCheckOut(sEncounterNumber, sPatientUnitNumber, moReport) Then GoTo Cleanup
    
    Call InformCoreShowModal
    If iPrintOption = PrintOnly Then
        If Not moCrystal.PrintReport(moReport) Then GoTo Cleanup
    ElseIf iPrintOption = PrintPreview Then
        If Not moCrystal.DisplayReport(moReport) Then GoTo Cleanup
    End If
    printPatientCheckOut = True
    GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".printPatientCheckOut", App.EXEName, vbExclamation
Cleanup:
    Call InformCoreDoneShowModal
    Set oFormCheckOut = Nothing
End Function

Public Function PrintHospitalSetup(Title As String, colHospitalSetup As Collection) As Boolean

    PrintHospitalSetup = False
On Error GoTo ErrHandler
    
    Dim ofrmHospitalSetup As FormHospitalSetup
    
    Set ofrmHospitalSetup = New FormHospitalSetup
    If Not ofrmHospitalSetup.BuildHospitalSetupReport(moReport, Title, colHospitalSetup) Then GoTo Cleanup
    Call InformCoreShowModal
    If Not moCrystal.DisplayReport(moReport) Then GoTo Cleanup
    
    PrintHospitalSetup = True
    GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".PrintHospitalSetup", App.EXEName, vbExclamation
Cleanup:
    Call InformCoreDoneShowModal
    Set ofrmHospitalSetup = Nothing
    Screen.MousePointer = vbDefault
End Function

'Use for testing printHospitalSetup
'Private Sub CreatMarkupDataForHospitalSetup(colHospSU As Collection)
'    Dim i As Integer
'    Dim oMarkUp As HospitalSetupData
'    Dim toggle As Boolean
'    toggle = True
'    For i = 1 To 20
'        Set oMarkUp = New HospitalSetupData
'        oMarkUp.Specialty = "MarkupSpecialty" & i
'        oMarkUp.BedsAvailable = i + 1
'        oMarkUp.BedsOccupied = i
'        oMarkUp.TotalBeds = i + i + 1
'        oMarkUp.BedsOpen = "1"
'        oMarkUp.Location = i & "LLLLLLooonnnnnnnnnnnnnggggggggggggggggggMarkupLocation" & i
'        oMarkUp.StatusActive = toggle
'        toggle = Not toggle
'        oMarkUp.StatusOpen = toggle
'        colHospSU.Add oMarkUp
'        Set oMarkUp = Nothing
'    Next i
'End Sub

Public Function print_vsFlexGrid(gridHWnd As Long, _
                                 sHeader As String, _
                                 sFooter As String, _
                                 Optional sFrmCaption As String = sDefault_Caption, _
                                 Optional oHeaderProperties As PrintGridHeaderProperties = Nothing, _
                                 Optional oFooterProperties As PrintGridFooterProperties = Nothing) As Boolean
'This is generic printing for vsFlexGrid
' PrintGridHeaderProperties & PrintGridFooterProperties are for future improvement
    print_vsFlexGrid = False
On Error GoTo ErrHandler
    Dim ofrmPrtGrid As frmPrtFlexGrid7
    Set ofrmPrtGrid = New frmPrtFlexGrid7
    
    Call ofrmPrtGrid.SetFlexGrid(gridHWnd, sHeader, sFooter)
    
    '--- SCR 136208  Sherry Wang  7/20/2007
    ofrmPrtGrid.Caption = sDefault_Caption
    
    Call InformCoreShowModal
    ofrmPrtGrid.Show vbModal
    print_vsFlexGrid = True
    GoTo Cleanup
ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".print_vsFlexGrid", App.EXEName, vbExclamation
Cleanup:
    Call InformCoreDoneShowModal
    Call DestroyObj(ofrmPrtGrid)
End Function

Public Function printInpatient(gridHWnd As Long, sHeader As String, sFooter As String) As Boolean
    printInpatient = False
On Error GoTo ErrHandler
    Dim ofrmPrtGrid As frmPrtFlexGrid7
    Set ofrmPrtGrid = New frmPrtFlexGrid7
    
    Call ofrmPrtGrid.SetFlexGrid(gridHWnd, sHeader, sFooter)
    
    '--- SCR 136208  Sherry Wang   7/20/2007
    ofrmPrtGrid.Caption = sDefault_Caption
    
    Call InformCoreShowModal
    ofrmPrtGrid.Show vbModal
    printInpatient = True
    GoTo Cleanup
ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".printInpatient", App.EXEName, vbExclamation
Cleanup:
    Call InformCoreDoneShowModal
    Call DestroyObj(ofrmPrtGrid)
End Function

Public Function printAppointments(Title As String, colAppts As Collection) As Boolean

On Error GoTo ErrHandler

    printAppointments = False
    
    Dim ofrmAppointment As FormAppointments
    
    Set ofrmAppointment = New FormAppointments
    If Not ofrmAppointment.BuildAppointmentsReport(moReport, Title, colAppts) Then GoTo Cleanup
    Call InformCoreShowModal
    If Not moCrystal.DisplayReport(moReport) Then GoTo Cleanup
    
    printAppointments = True
    GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".printAppointments", App.EXEName, vbExclamation
Cleanup:
    Call InformCoreDoneShowModal
    Set ofrmAppointment = Nothing
    Screen.MousePointer = vbDefault
End Function

Private Sub GemsUpdateCheckOutInfo(ByVal strClinicNCID As String)

On Error GoTo Proc_Exit


    Dim rsCheckOut          As ADODB.Recordset
    Dim strSQL              As String
    Dim objCheckOutInfoItem As CheckOutInfoItem
    Dim objFormCheckOut     As FormCheckOut
    'Dim objSQL              As ICHCSII_SQL.ISqlOpsEx
    
    Set objFormCheckOut = New FormCheckOut
    
    Screen.MousePointer = vbHourglass
    If objFormCheckOut.GetCheckOutInfo(strClinicNCID, True, rsCheckOut) Then
    
       Set objCheckOutInfoItem = New CheckOutInfoItem
       
       If Not rsCheckOut.EOF Then
          
           With objCheckOutInfoItem
             .ClinicInfo = rsCheckOut("ClinicInfo") & vbNullString
             .ConsultInfo = rsCheckOut("ConsultInfo") & vbNullString
             .HeaderInfo = rsCheckOut("HeaderInfo") & vbNullString
             .LabInfo = rsCheckOut("LabInfo") & vbNullString
             .MedicationInfo = rsCheckOut("MedicationInfo") & vbNullString
             .XRAYInfo = rsCheckOut("XRAYInfo") & vbNullString
          End With
       End If
    
    
       Load frmCheckOutInfo
       frmCheckOutInfo.CheckOutInfo = objCheckOutInfoItem
       Set frmCheckOutInfo.poMyParent = Me
    
       Screen.MousePointer = vbDefault
       frmCheckOutInfo.Show vbModal
     
       Set frmCheckOutInfo.poMyParent = Nothing
       Unload frmCheckOutInfo
    
       If Not ReturnObject Is Nothing Then
    
         strSQL = "SELECT MedicationInfo,XRayInfo,LabInfo,ConsultInfo,HeaderInfo,ClinicInfo,Clinic_NCID " & _
                  "FROM CheckOutInfo WHERE  Clinic_NCID = " & strClinicNCID
         
         goDBSqlOpsEx.Execute strSQL      'hwcc 12/10/2003 - use a shared CHCSII_CONN, SQLOPS and DAS object
         
         If (goDBSqlOpsEx.BOF And goDBSqlOpsEx.EOF) Then
             goDBSqlOpsEx.AddNew
             goDBSqlOpsEx("Clinic_NCID") = strClinicNCID
         End If
    
         goDBSqlOpsEx("MedicationInfo") = objCheckOutInfoItem.MedicationInfo & vbNullString
         goDBSqlOpsEx("XRayInfo") = objCheckOutInfoItem.XRAYInfo & vbNullString
         goDBSqlOpsEx("LabInfo") = objCheckOutInfoItem.LabInfo & vbNullString
         goDBSqlOpsEx("ConsultInfo") = objCheckOutInfoItem.ConsultInfo & vbNullString
         goDBSqlOpsEx("HeaderInfo") = objCheckOutInfoItem.HeaderInfo & vbNullString
         goDBSqlOpsEx("ClinicInfo") = objCheckOutInfoItem.ClinicInfo & vbNullString
          
          
         goDBSqlOpsEx.Update
          
     End If
    
    End If
      
Proc_Exit:
    Set objFormCheckOut = Nothing
    Set objCheckOutInfoItem = Nothing
    Set rsCheckOut = Nothing
    If Err.Number <> 0 Then
       gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".GemsUpdateCheckOutInfo", App.EXEName, vbExclamation
    End If
    Screen.MousePointer = vbDefault

End Sub

Public Sub UpdateCheckOutInfo(ByVal strClinicNCID As String)

    On Error GoTo ErrHandler
    Screen.MousePointer = vbHourglass

    GemsUpdateCheckOutInfo strClinicNCID  'change the way this sub is currently working - tchikaya brooks 09-06-2001
    GoTo Cleanup

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "TableOps.GetDistinct", "CustomReportOps", vbExclamation)
    
Cleanup:
   Screen.MousePointer = vbDefault

End Sub

Private Sub Class_Initialize()

On Error GoTo ErrHandler


Dim sErrMess As String
Dim sReportsPath As String

#If coiDebug Then
     gCOI_DL.debugStart
#End If
    
'-- Iniatialized shared global object
'-----------------------------------------------------
    Set gobjshared = New CWShared
    
'-- Set Global Path Variables
'-----------------------------------------------------
    sReportsPath = gobjshared.ReportsPath
    If Mid$(sReportsPath, Len(sReportsPath), 1) <> "\" Then
        sReportsPath = sReportsPath & "\"
    End If
    gsFormPath = sReportsPath           '-- Forms File Path
    
'-- Open connection to the Access database
'-----------------------------------------------------
    Set moCrystal = New Crystal
    
    Call CreateDBAccessObjs   'hwcc 12/10/2003
   
    '-- This routine is executed for the exclusive purpose of handling a Crystal 8 Bug
    If Not InitializeCrystal() Then GoTo Cleanup
    GoTo Cleanup

ErrHandler:
    sErrMess = Err.Description & vbCrLf & vbCrLf & "Path:" & gsFormPath & vbCrLf & vbCrLf & "Database:" & gsDatabase
    gobjshared.ShowVBError Err.Number, sErrMess, "Forms.Class_Initialize", App.EXEName, vbExclamation
Cleanup:
    Set moReport = Nothing
    Exit Sub
End Sub

Private Sub Class_Terminate()

    #If coiDebug Then
         gCOI_DL.debugStop
         Set gCOI_DL = Nothing
    #End If
    
    Set moCrystal = Nothing
    Set moReport = Nothing

    Call DestroyDBAccessObjs

End Sub

Public Function printSF513(iPrintOption As PrintOptions, _
                           oConsultFormItem As ConsultFormItem, _
                           sPatientUnitNumber As String) As Boolean


On Error GoTo ErrHandler
    printSF513 = False
    
    
    Dim oFormSF513 As FormSF513
    
    Screen.MousePointer = vbHourglass
    Set oFormSF513 = New FormSF513
    
    If Not oFormSF513.BuildSF513(oConsultFormItem, sPatientUnitNumber, moReport) Then GoTo Cleanup
    
    Call InformCoreShowModal
    If iPrintOption = PrintOnly Then
        If Not moCrystal.PrintReport(moReport) Then GoTo Cleanup
    ElseIf iPrintOption = PrintPreview Then
        If Not moCrystal.DisplayReport(moReport) Then GoTo Cleanup
    End If
    
    printSF513 = True
    GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".printSF513", App.EXEName, vbExclamation
Cleanup:
    Call InformCoreDoneShowModal
    Set oFormSF513 = Nothing
End Function

Public Function printSF513Complete(iPrintOption As PrintOptions, _
                                   Consult_Id As String, _
                                   ConsultationReportText As String) As Boolean

On Error GoTo ErrHandler
    printSF513Complete = False
    
    Dim oFormSF513 As FormSF513
    
    Screen.MousePointer = vbHourglass
    Set oFormSF513 = New FormSF513
    
    If Not oFormSF513.BuildSF513Complete(Consult_Id, ConsultationReportText, moReport) Then GoTo Cleanup
    
    Call InformCoreShowModal
    If iPrintOption = PrintOnly Then
        If Not moCrystal.PrintReport(moReport) Then GoTo Cleanup
    ElseIf iPrintOption = PrintPreview Then
        If Not moCrystal.DisplayReport(moReport) Then GoTo Cleanup
    End If
    
    printSF513Complete = True
    GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".printSF513Complete", App.EXEName, vbExclamation
Cleanup:
    Call InformCoreDoneShowModal
    Set oFormSF513 = Nothing
End Function
' Quan 80054
Public Sub printDispHistory(ByVal OrderId As String, ByVal UnitNumber As String)
    frmMedHistory.Display OrderId, UnitNumber
End Sub

Public Property Get ReportObject() As Object
    Set ReportObject = moReport
End Property

Public Property Set Comm(ByVal RHV As Object)
    Set cwobjCWComm = RHV
End Property

Public Property Set Logon(ByVal RHV As Object)
    Set cwobjLogonEnviron = RHV
End Property

Public Property Get DocumentPrinted() As Boolean
    DocumentPrinted = gbDocumentPrinted
End Property

