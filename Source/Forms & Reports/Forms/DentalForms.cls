VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DentalForms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private moConn As CHCSII_CONN.Conn
Private moDAS As ICHCSII_DataAccess.ICHCSII_DAS
Private mrsMain As ADODB.Recordset
Private oSF As New SVGTeeth.SF
Public oTeethSVG As SVGTeeth.clsTeeth

Const sMODULE_NAME As String = "CHCSII_Forms."

Public Property Set Comm(ByVal RHV As Object)
    Set cwobjCWComm = RHV
End Property
'********************************************************************************
' Name:         LoadGraphics
' Description:  Load DX/TX Images
' Parameter:    plUnitNumber - Patient Unit Number
'               plColor:
'                   vbButtonFace    - Graphics for DataView
'                   vbWhite         - Graphics for SF603A Report
' Return:       True - sucessful, False - fail
' Called By:    PrintSF603A2()
' Histary:
'   Developer:  Date:       Description
'   ----------  ---------   -----------------------------------------------
'   Wilson Sun  07/12/01    Add check calling getDxView and getTxView order(SCR 14381)
'   Wilson Sun  05/16/01    Add new code for retrieve graphics
'   Wilson Sun  05/08/01    Commented out since no images to print
'   Wilson Sun  04/10/01    Created
'********************************************************************************
Public Function LoadGraphics(ByRef poPatient As Object, _
                            ByRef poEncounter As Object, _
                            Optional ByVal plColor As Long) As Boolean
'
'    Const sPROC_NAME As String = "LoadGraphics"
'    Dim i As Integer
'    Dim objCurrentGCOps As CHCSII_Dental_BigTooth.GCOps
'    Dim objPicture As IPictureDisp
'
'    On Error GoTo ErrHandler
'
'    'Screen.MousePointer = vbHourglass
'
'    Set objCurrentGCOps = New CHCSII_Dental_BigTooth.GCOps
'
'    If Not objCurrentGCOps.InitializeEx(poPatient, poEncounter) Then
'        MsgBox ("You cannot access SF603A for now because the dental encounter of this patient does not exist!")
'        LoadGraphics = False
'        GoTo CleanUp
'    End If
'
'    If objCurrentGCOps.IsDxView Then
'        Set gcolTeethTx = objCurrentGCOps.getTxView(plColor)
'        Set gcolTeethDx = objCurrentGCOps.getDxView(plColor)
'    Else
'        Set gcolTeethDx = objCurrentGCOps.getDxView(plColor)
'        Set gcolTeethTx = objCurrentGCOps.getTxView(plColor)
'    End If
'
'    LoadGraphics = True
'    GoTo CleanUp
'
'ErrHandler:
'    Call gobjshared.ShowVBError(Err.Number, Err.Description, sMODULE_NAME & sPROC_NAME, App.Title, vbCritical)
'    LoadGraphics = False
'    Set gcolTeethDx = Nothing
'    Set gcolTeethTx = Nothing
'
'CleanUp:
'    Set objCurrentGCOps = Nothing
'   ' cwobjCWComm.setStatusText ""
'    'Screen.MousePointer = vbDefault
End Function

'********************************************************************************
' Name:         LoadSF603Note
' Description:  Load DX/TX Images
' Parameter:    poEncounter - Encounter object
'               prsReport
' Return:       True - sucessful, False - fail
' Called By:    PrintSF603A2()
' Note:         This method is called only prsReport parameter is nothing
' Histary:
'   Developer:      Date:       Description
'   ----------      ---------   -----------------------------------------------
'   Wilson Sun      08/08/01    1. Took DENTAL schema away from SQL statement.
'                               2. Modifiey DAS object declareation to fit theater needs
'   Wilson Sun      05/17/01    Created
'********************************************************************************
Public Function LoadSF603Note(ByRef poEncounter As Object, _
                            ByRef prsReport As ADODB.Recordset) As Boolean

    Const sPROC_NAME As String = "LoadSF603Note"
    Dim rsDRC As ADODB.Recordset
    Dim sSQL As String
    Dim blnReturn As Boolean
    Dim blnCancel As Boolean
    
    On Error GoTo ErrHandler

    'Screen.MousePointer = vbHourglass
    Set prsReport = New ADODB.Recordset 'Becasue prsReport is nothing)
    'Retrieve DRC
    sSQL = "select DRC from DENTAL_ENCOUNTER " & vbCrLf & _
        "Where PATIENT_UNIT_ID = " & poEncounter.PatientID & vbCrLf & _
        "And   ENCOUNTER_FACILITY_NCID = " & poEncounter.FacilityNCID & vbCrLf & _
        "And   ENCOUNTER_NUMBER = " & poEncounter.encounterID
    Set rsDRC = moDAS.OpenRecordset(sSQL)

    'KDunne 6/20/2002
    'SCR 25020 - If there is no DRC information, we need to raise an error message and cancel
    '      the print job.
    '
    If rsDRC Is Nothing Then
        blnCancel = True
    Else
        If rsDRC.BOF And rsDRC.EOF Then
            blnCancel = True
        End If
    End If
    
    If blnCancel Then
        MsgBxARMd "Dental information not available for current encounter. Dental Encounter must be " & _
            "initialized prior to printing SF603A.", , "CHCSII_Dental"
        blnReturn = False
        GoTo FuncExit
    End If
    'End SCR  25020
    
    'Build Record set to be print
    prsReport.CursorLocation = adUseClient
    prsReport.Fields.Append "ENCOUNTER_DATE", adDate
    prsReport.Fields.Append "NOTE", adBSTR
    prsReport.Fields.Append "DRC", adBSTR
    prsReport.Open

    With prsReport
        .AddNew
        .Fields!ENCOUNTER_DATE = poEncounter.startdts
        .Fields!Note = poEncounter.CurrentRTF
        .Fields!DRC = rsDRC.Fields!DRC
    End With

    blnReturn = True
    
FuncExit:
    LoadSF603Note = blnReturn
    
    Set gcolTeethDx = Nothing
    Set gcolTeethTx = Nothing
    Set rsDRC = Nothing
    
    Exit Function

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, sMODULE_NAME & sPROC_NAME, App.Title, vbCritical)
    blnReturn = False
    
    GoTo FuncExit
End Function


'***********************************************************************
'Function name: PrintDD2322
'Description:   This is the main interface for calling to print DD2322
'Parameters:
'   poPatient               - Patient object for Demographic Info
'   psLocalCaseNo           - Case number
'   psClinicName            - Clinic Name
'   psClinicNCID            - ClinicNCID
'   psProviderName          - Provider Name
'   prsSubReport            - For subreport data
'   pbPreview               - Weither need Preview or not
'Return Value:  Ture, sucessful; False, otherwise
'History
'Date       Developer   Description
'========   ==========  =========================================================
'11/15/01   Wilson Sun  Added Sponsor's SSN field
'07/16/01   Wilson Sun  Added subreport for SCR: 13107
'03/03/01   Wilson Sun  Created
'***********************************************************************
Public Function PrintDD2322(ByRef poPatient As Patient, _
                            ByVal psLocalCaseNo As String, _
                            ByVal psClinicName As String, _
                            ByVal psClinicNCID As String, _
                            ByVal psProviderName As String, _
                            ByRef prsSubReport As ADODB.Recordset, _
                            Optional ByVal pbPreview As Boolean = True) As Boolean
    
    Dim oForm As New frmDentalViewer
    Dim oReport As New CRDD2322

    Dim sSQL As String
    
    'Variable for Report Formula
    Dim sDateOfBirth As String
    Dim sFMP As String
    Dim sOrg As String
    Dim sPatientName As String
    Dim sSex As String
    Dim sSSN As String
    Dim sSponsorName As String
    Dim sSponsorRank As String
    Dim sSponsorPatientCategory As String
    Dim sFMPSponsorSSN  As String
    Dim sSponsorStatus As String
    
    On Error GoTo ErrHandler
    
    'ITT Watermark
    Dim oFormula As FormulaFieldDefinition
    If Not gobjshared.UseSQLServer Or gobjshared.IsAppMode(modeCHCSI_GUI) Then
        For Each oFormula In oReport.FormulaFields
            If oFormula.FormulaFieldName = "TrainingUseOnly" Then
               oFormula.Text = ""
            End If
        Next
    End If
    
    PrintDD2322 = False
    'Load oForm
    Screen.MousePointer = vbHourglass

    'Set oForm = New frmDentalViewer
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Assign object variable to local variable
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''Patient Info
    With poPatient
        sPatientName = .Name
        sSex = .Sex
        sSSN = .SSN
        sDateOfBirth = Format(CDate(.DateOfBirth), "dd mmm yyyy")
        sFMP = .FMP
        sFMPSponsorSSN = .FMPSponsorSSN
    End With
    '''Sponsor Info
    With poPatient.Sponsor
        sSponsorName = .Name
        sSponsorRank = .Rank
        sSponsorPatientCategory = .PatientCategory
        sSponsorStatus = .MilitaryStatus
        sOrg = .UnitIdentificationCode
    End With
        
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Set formulas
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    With oReport
        '--0 Set Subreport record set
        .OpenSubreport("CRDD2322_Sub").Database.SetDataSource prsSubReport

        '--1 Case No
        .fldLocalCaseNo1.Field.Text = "'" & psLocalCaseNo & "'"
        
        '--2 Clinic Info
        .fldTreatmentName2.Field.Text = "'" & psClinicName & "'"
        .fldMaillingAddress2.Field.Text = "'" & GetClinicAddress(psClinicNCID) & "'"
        
        '--4 Patient Name
        .fldPatientName4.Field.Text = "'" & FirstLastMiName(sPatientName) & "'"
        
        '--5 SSN
        .FormulaFields(5).Text = "'" & sSSN & "'" 'SSN formated from Crystal Designer
        
        '--6 Grade/Rank
        .fldGrade6.Field.Text = "'" & sSponsorRank & "'"
        
        '--7 Age
        .fldAge7.Field.Text = "'" & CStr(Round((Date - poPatient.DateOfBirth) / 365)) & "'"
        
        '--Beneficiary Type
        .fldBeneficiaryType9.Field.Text = "'" & GetDepartService(sSponsorPatientCategory) & "'"
        
        '--Organization
        .fldOrganization10.Field.Text = "'" & sOrg & "'"
        
        'SCR 40565 RWalker 9/11/2003 - Form DD 2322 not completely printing out.
        'and the print out does not print the Grade of the Dental Officer in box 27
        'The calling code will append the retrieved rang/grade to the end of the
        'Provider's Name value separated by a slash
        '--Provider Name
        .fldProviderName27.Field.Text = "'" & psProviderName & "'"
        
        '--11/15/01
        .fldFMPSponsorSSN.Field.Text = "'" & sFMPSponsorSSN & "'"
        
    End With
        
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Set formula of Clinic
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    If pbPreview Then
        '''''''''''''''''''''''''''''''''''''''''''''
        ' Display Report
        '''''''''''''''''''''''''''''''''''''''''''''
        With oForm.CRViewer1
            .ReportSource = oReport
            .DisplayGroupTree = False
            .EnableGroupTree = False
            .DisplayTabs = False
            .EnableDrillDown = False
            .EnableRefreshButton = False
            DoEvents
            .ViewReport
            DoEvents
            .Zoom 2
            End With
        cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "True", NCID_CW_APPLICATION, NCID_DENTAL_LABS_APPLICATION
        Screen.MousePointer = vbDefault
        oForm.Show vbModal 'scr 27292
        cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "True", NCID_CW_APPLICATION, NCID_DENTAL_LABS_APPLICATION
       
    Else
        '''''''''''''''''''''''''''''''''''''''''''''
        ' Print Report directly, no display
        '''''''''''''''''''''''''''''''''''''''''''''
        oReport.PrintOut False
    End If
    
    PrintDD2322 = True

    GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "DentalForms.PrintDD2322", "CHCSII_Forms", vbExclamation
Cleanup:
    Set oReport = Nothing
    Set oForm = Nothing
    Screen.MousePointer = vbDefault
End Function
'***********************************************************************
'Function name: PrintSF603A2
'Description:   This is the main interface for calling to print SF603A
'Parameters:
'   poPatient       - Patient object for Demographic Info
'   prsReport       - Record set
'   pcolTeethDx     - Dx graphics collection (if nothing, means pass from outside dental)
'   pcolTeethTx     - Tx graphics collection (if nothing, means pass from outside dental)
'   pbPreview       - Preview or Print
'   Return Value:   - Ture, sucessful; False, otherwise
'Notes: Parameters pcolTeethDx and pcolTeethTx are dimmy now
'       since Teeth Images coming from SVG file.
'History:
'   Date       Developer   Description
'   =========  ==========  ================================================
'   01/14/02   Wilson Sun  Created. Clone from PrintSF603A2_Old
'***********************************************************************
Public Function PrintSF603A2(ByRef poPatient As Patient, _
                            ByRef poEncounter As Object, _
                            ByRef prsReport As ADODB.Recordset, _
                            ByRef pcolTeethDx As Collection, _
                            ByRef pcolTeethTx As Collection, _
                            Optional ByVal pbPreview As Boolean = True) As Boolean
    
'    Dim oForm As frmDentalViewer
    Dim oViewer As CRViewer
    Dim oReport As CRSF603A2
    Dim oSect As Section
    Dim oRptObject As Object
    Dim sSQL As String
    
    'Variable for Report Formula
    Dim sPatientName As String
    Dim sSex As String
    Dim sSSN As String
    Dim sDateOfBirth As String
    Dim sFMP As String
    Dim sSponsorName As String
    Dim sSponsorRank As String
    Dim sSponsorPatientCategory As String
    Dim sSponsorStatus As String
    Dim sSponsorSSN As String
    Dim sOrg As String
    Dim strStatus As String
    
    On Error GoTo ErrHandler
    PrintSF603A2 = False
    Screen.MousePointer = vbHourglass
    
    Set cwobjPatient = poPatient

    Load frmTeethPictures
    Set frmTeethPictures.TeethSVG = oTeethSVG
    cwobjCWComm.Message cwiSHOWING_MODAL_FORM, "True", NCID_CW_APPLICATION, NCID_DDSA_APPLICATION
    frmTeethPictures.Show vbModal
    cwobjCWComm.Message cwiDONE_SHOWING_MODAL_FORM, "True", NCID_CW_APPLICATION, NCID_DDSA_APPLICATION

    'Make sure we didn't encounter an error generating the graphic
    If frmTeethPictures.ErrorOccurred Then GoTo Cleanup
    
    '---------------------------------------------------------
    'Retrieve Notes if print from Encounter screen
    '---------------------------------------------------------
    If prsReport Is Nothing Then
        If Not LoadSF603Note(poEncounter, prsReport) Then GoTo Cleanup
    End If
   
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Assign object variable to local variable
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Patient Info
    With poPatient
        sPatientName = .Name
        sSex = .Sex
        sSSN = .SSN
        sDateOfBirth = Format(CDate(.DateOfBirth), "dd mmm yyyy")
        sFMP = .FMP
        sSponsorSSN = .SponsorSSN
    End With
    'Sponsor Info
    With poPatient.Sponsor
        sSponsorName = .Name
        sSponsorRank = .Rank
        sSponsorPatientCategory = .PatientCategory
        sSponsorStatus = .MilitaryStatus
        sOrg = .UnitIdentificationCode
    End With
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Set formula of Demorgraphic for report page foot
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set oReport = New CRSF603A2
    
    'ITT Watermark
    Dim oFormula As FormulaFieldDefinition
    If Not gobjshared.UseSQLServer Or gobjshared.IsAppMode(modeCHCSI_GUI) Then
        For Each oFormula In oReport.FormulaFields
            If oFormula.FormulaFieldName = "TrainingUseOnly" Then
               oFormula.Text = ""
            End If
        Next
    End If
    
    Debug.Print "Set oReport = New CRSF603A2 "
    'KDunne 4/15/2002
    'SCR 23038
    strStatus = cwobjCWComm.GetStatusText
    cwobjCWComm.setStatusText "Loading Teeth Pictures..."
    'End SCR 23038
    With oReport
        Set .PicDx = frmTeethPictures.PicDx ' oTeethSVG.getTeethPicture(1)
        Set .PicTx = frmTeethPictures.PicTx
        .DiscardSavedData
        .Database.SetDataSource prsReport
        Debug.Print ".Database.SetDataSource prsReport"
        .fldPatientName.Field.Text = "'" & FirstLastMiName(sPatientName) & "'"
        .fldSex.Field.Text = "'" & sSex & "'"
        .FormulaFields(10).Text = "'" & sSSN & "'" 'SSN formated from Crystal Designer
        .fldDateOfBirth.Field.Text = "'" & sDateOfBirth & "'"
        .fldFMP.Field.Text = "'" & sFMP & "'"
        .fldSponsorSSN.Field.Text = "'" & sSponsorSSN & "'" 'formated from Crystal Designer
        'Relationship to sponsor only patient and sponsor's name are different
        If sPatientName <> sSponsorName Then
            .fldRelationshipToSponsor.Field.Text = "'" & GetRelationshipToSponsor(sFMP) & "'"
        End If
        .fldSponsorName.Field.Text = "'" & FirstLastMiName(sSponsorName) & "'"
        .fldRank.Field.Text = "'" & sSponsorRank & "'"
        .fldDepartService.Field.Text = "'" & GetDepartService(sSponsorPatientCategory) & "'"
        .fldComponentStatus.Field.Text = "'" & sSponsorStatus & "'"
        .fldOrganization.Field.Text = "'" & sOrg & "'"
    End With
    cwobjCWComm.setStatusText strStatus 'SCR 23038
    'Set oForm = New frmDentalViewer
    Load frmDentalViewer
    
    Set oViewer = frmDentalViewer.CRViewer1
    If pbPreview Then
        '''''''''''''''''''''''''''''''''''''''''''''
        ' Display Report
        '''''''''''''''''''''''''''''''''''''''''''''
        With oViewer
            .ReportSource = oReport
            .DisplayGroupTree = False
            .EnableGroupTree = False
            .DisplayTabs = False
            .EnableDrillDown = False
            .EnableRefreshButton = False
'            DoEvents
            .ViewReport
            Debug.Print ".ViewReport"
'            DoEvents
'            .Zoom 2
        End With
        frmDentalViewer.Show
    Else
        '''''''''''''''''''''''''''''''''''''''''''''
        ' Print Report directly, no display
        '''''''''''''''''''''''''''''''''''''''''''''
        oReport.PrintOut False
    End If

    PrintSF603A2 = True

    GoTo Cleanup

ErrHandler:
    Unload frmTeethPictures
    Set frmTeethPictures = Nothing
    gobjshared.ShowVBError Err.Number, Err.Description, "DentalForms.PrintSF603A2", "CHCSII_Forms", vbExclamation
Cleanup:
'    DoEvents
    Unload frmTeethPictures
    Set frmTeethPictures = Nothing
    Set oReport = Nothing
'    Set oForm = Nothing
    Set prsReport = Nothing
    Screen.MousePointer = vbDefault
End Function

'***********************************************************************
'Function name: PrintSF603A2_Old
'Description:   This is the main interface for calling to print SF603A
'Parameters:
'   poPatient       - Patient object for Demographic Info
'   prsReport       - Record set
'   pcolTeethDx     - Dx graphics collection (if nothing, means pass from outside dental)
'   pcolTeethTx     - Tx graphics collection (if nothing, means pass from outside dental)
'   pbPreview       - Preview or Print
'   Return Value:   - Ture, sucessful; False, otherwise
'History
'Date       Developer   Description
'=========  ==========  ================================================
'10/23/01   Wilson Sun  Added FMP/SponsorSSN to the report. SCR 17425
'05/14/01   Wilson Sun  Modify images collect retrieve method
'04/10/01   Wilson Sun  Created
'***********************************************************************
Public Function PrintSF603A2_Old(ByRef poPatient As Patient, _
                            ByRef poEncounter As Object, _
                            ByRef prsReport As ADODB.Recordset, _
                            ByRef pcolTeethDx As Collection, _
                            ByRef pcolTeethTx As Collection, _
                            Optional ByVal pbPreview As Boolean = True) As Boolean
    
''    Dim rs As New ADODB.Recordset
'    Dim oForm As New frmDentalViewer
'    Dim oReport As CRSF603A2
'    Dim oSect As Section
'    Dim oRptObject As Object
'    Dim sSQL As String
'
'    'Variable for Report Formula
'    Dim sPatientName As String
'    Dim sSex As String
'    Dim sSSN As String
'    Dim sDateOfBirth As String
'    Dim sFMP As String
'    Dim sSponsorName As String
'    Dim sSponsorRank As String
'    Dim sSponsorPatientCategory As String
'    Dim sSponsorStatus As String
'    Dim sSponsorSSN As String
'    Dim sOrg As String
'
'    On Error GoTo ErrHandler
'    PrintSF603A2_Old = False
'    Screen.MousePointer = vbHourglass
'
'    '---------------------------------------------------------
'    'Retrieve Teeth Images
'    '---------------------------------------------------------
'    Set gcolTeethDx = pcolTeethDx
'    Set gcolTeethTx = pcolTeethTx
'
'    If gcolTeethDx Is Nothing Then
'        ''''''''''''''''''''''''''''''''''''''''''''''''''
'        'Load Teeth Graphics
'        ''''''''''''''''''''''''''''''''''''''''''''''''''
'        If Not LoadGraphics(poPatient, poEncounter, vbWhite) Then GoTo CleanUp
'    End If
'
'    '---------------------------------------------------------
'    'Retrieve Notes if print outside of dental module
'    '---------------------------------------------------------
'    If prsReport Is Nothing Then
'        If Not LoadSF603Note(poEncounter, prsReport) Then GoTo CleanUp
'    End If
'
'
'    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    'Assign object variable to local variable
'    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    '''Patient Info
'    With poPatient
'        sPatientName = .Name
'        sSex = .Sex
'        sSSN = .SSN
'        sDateOfBirth = Format(CDate(.DateOfBirth), "dd mmm yyyy")
'        sFMP = .FMP
'        sSponsorSSN = .SponsorSSN
'    End With
'    '''Sponsor Info
'    With poPatient.Sponsor
'        sSponsorName = .Name
'        sSponsorRank = .Rank
'        sSponsorPatientCategory = .PatientCategory
'        sSponsorStatus = .MilitaryStatus
'        sOrg = .UnitIdentificationCode
'    End With
'
'    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    ' Set formula of Demorgraphic for report page foot
'    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    Set oReport = New CRSF603A2
'    Debug.Print "Set new CRSF603A2 "
'    With oReport
'        .DiscardSavedData
'        .Database.SetDataSource prsReport
'        .fldPatientName.Field.Text = "'" & FirstLastMiName(sPatientName) & "'"
'        .fldSex.Field.Text = "'" & sSex & "'"
'        .FormulaFields(10).Text = "'" & sSSN & "'" 'SSN formated from Crystal Designer
'        .fldDateOfBirth.Field.Text = "'" & sDateOfBirth & "'"
'        .fldFMP.Field.Text = "'" & sFMP & "'"
'        .fldSponsorSSN.Field.Text = "'" & sSponsorSSN & "'" 'formated from Crystal Designer
'        'Relationship to sponsor only patient and sponsor's name are different
'        If sPatientName <> sSponsorName Then
'            .fldRelationshipToSponsor.Field.Text = "'" & GetRelationshipToSponsor(sFMP) & "'"
'        End If
'        .fldSponsorName.Field.Text = "'" & FirstLastMiName(sSponsorName) & "'"
'        .fldRank.Field.Text = "'" & sSponsorRank & "'"
'        .fldDepartService.Field.Text = "'" & GetDepartService(sSponsorPatientCategory) & "'"
'        .fldComponentStatus.Field.Text = "'" & sSponsorStatus & "'"
'        .fldOrganization.Field.Text = "'" & sOrg & "'"
'    End With
'
'    Load oForm
'    If pbPreview Then
'        '''''''''''''''''''''''''''''''''''''''''''''
'        ' Display Report
'        '''''''''''''''''''''''''''''''''''''''''''''
'        With oForm.CRViewer1
'            .ReportSource = oReport
'            .DisplayGroupTree = False
'            .EnableGroupTree = False
'            .DisplayTabs = False
'            .EnableDrillDown = False
'            .EnableRefreshButton = False
'            DoEvents
'            .ViewReport
'            DoEvents
'            .Zoom 2
'            End With
'        oForm.Show
'    Else
'        '''''''''''''''''''''''''''''''''''''''''''''
'        ' Print Report directly, no display
'        '''''''''''''''''''''''''''''''''''''''''''''
'        oReport.PrintOut False
'    End If
'
'    PrintSF603A2_Old = True
'
'    GoTo CleanUp
'
'ErrHandler:
'  gobjshared.ShowVBError Err.Number, Err.Description, "DentalForms.PrintSF603A2_Old", "CHCSII_Forms", vbExclamation
'CleanUp:
'    Set oReport = Nothing
'    Set oForm = Nothing
''    Set rs = Nothing
'    Set prsReport = Nothing
'    Screen.MousePointer = vbDefault
End Function


'***********************************************************************
'Function name: PrintSF603A
'Description:   This is the main interface for calling to print SF603A
'Parameters:
'   poPatient    - Patient object for Demographic Info
'   psRTF        - RTF format string
'   pbPreview    - Weither need Preview or not
'Return Value:  Ture, sucessful; False, otherwise
'History
'Date       Developer   Description
'=========  =========   ==================================================
'10/05/01   Wilson Sun  Since we use PrintSF603A2 instead.
'***********************************************************************
Public Function PrintSF603A(ByVal poPatient As Object, _
                            ByVal psRTF As String, _
                            Optional ByVal pbPreview As Boolean) As Boolean
    
'    Dim rs As New ADODB.Recordset
'    Dim oForm As New frmDentalViewer
'    Dim oReport As New CRSF603A
'    Dim sSQL As String
'
'    'Variable for Report Formula
'    Dim sPatientName As String
'    Dim sSex As String
'    Dim sSSN As String
'    Dim sDateOfBirth As String
'    Dim sFMP As String
'    Dim sSponsorName As String
'    Dim sSponsorRank As String
'    Dim sSponsorPatientCategory As String
'    Dim sSponsorStatus As String
'    Dim sOrg As String
'    On Error GoTo ErrHandler
'
'    PrintSF603A = False
'    'Load oForm
'    Screen.MousePointer = vbHourglass
'
'    ''''''''''''''''''''''''''''''''''''''''''''''''''
'    'Load Teeth Graphics
'    ''''''''''''''''''''''''''''''''''''''''''''''''''
'    'If Not LoadGraphics(poPatient.UnitNumber, vbWhite) Then GoTo CleanUp
'
'    ''''''''''''''''''''''''''''''''''''''''''''''''''
'    ' Get main body record set
'    ''''''''''''''''''''''''''''''''''''''''''''''''''
'    'sSQL = "select T2.CREATE_DATE_TIME, T2.TREATMENT_TECHNIQUE_COMMENT, T1.DRC " & _
'        "from goSqlOpsDENTAL_ENCOUNTER T1, goSqlOpsNOTE_TREATMENT T2 " & _
'        "Where T1.DENTAL_ENCOUNTER_ID = T2.DENTAL_ENCOUNTER_ID " & _
'        "and T1.PATIENT_UNIT_ID = " & poPatient.UnitNumber & _
'        " order by 1 desc"
'    'sSQL = "select T2.CREATE_DATE_TIME, T2.TREATMENT_TECHNIQUE_COMMENT, T1.DRC " & _
'        "from goSqlOpsDENTAL_ENCOUNTER T1, goSqlOpsNOTE_TREATMENT T2 " & _
'        "Where T1.DENTAL_ENCOUNTER_ID = T2.DENTAL_ENCOUNTER_ID " & _
'        " order by 1 desc"
'
'    'Set mrsMain = moDAS.OpenRecordset(sSQL)
'    'oReport.Database.SetDataSource mrsMain
'    'Set mrsMain = Nothing
'
'
'    rs.CursorLocation = adUseClient
'
'    rs.Fields.Append "NOTES", adBSTR
'    rs.Open
'    rs.AddNew
'    rs!NOTES = psRTF
'    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    'Assign object variable to local variable
'    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    '''Patient Info
'    With poPatient
'        sPatientName = .Name
'        sSex = .Sex
'        sSSN = .SSN
'        sDateOfBirth = Format(CDate(.DateOfBirth), "dd mmm yyyy")
'        sFMP = .FMP
'    End With
'    '''Sponsor Info
'    With poPatient.Sponsor
'        sSponsorName = .Name
'        sSponsorRank = .Rank
'        sSponsorPatientCategory = .PatientCategory
'        sSponsorStatus = .MilitaryStatus
'        sOrg = .UnitIdentificationCode
'    End With
'
'    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    ' Set formula of Demorgraphic for report page foot
'    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    With oReport
'        .DiscardSavedData
'        .Database.SetDataSource rs
'        .fldPatientName.Field.Text = "'" & FirstLastMiName(sPatientName) & "'"
'        .fldSex.Field.Text = "'" & sSex & "'"
'        .FormulaFields(10).Text = "'" & sSSN & "'" 'SSN formated from Crystal Designer
'        .fldDateOfBirth.Field.Text = "'" & sDateOfBirth & "'"
'        'Relationship to sponsor only patient and sponsor's name are different
'        If sPatientName <> sSponsorName Then
'            .fldRelationshipToSponsor.Field.Text = "'" & GetRelationshipToSponsor(sFMP) & "'"
'        End If
'        .fldSponsorName.Field.Text = "'" & FirstLastMiName(sSponsorName) & "'"
'        .fldRank.Field.Text = "'" & sSponsorRank & "'"
'        .fldDepartService.Field.Text = "'" & GetDepartService(sSponsorPatientCategory) & "'"
'        .fldComponentStatus.Field.Text = "'" & sSponsorStatus & "'"
'        .fldOrganization.Field.Text = "'" & sOrg & "'"
'    End With
'
'
'    Load oForm
'    If pbPreview Then
'        '''''''''''''''''''''''''''''''''''''''''''''
'        ' Display Report
'        '''''''''''''''''''''''''''''''''''''''''''''
'        With oForm.CRViewer1
'            .ReportSource = oReport
'            .DisplayGroupTree = False
'            .EnableGroupTree = False
'            .DisplayTabs = False
'            .EnableDrillDown = False
'            .EnableRefreshButton = False
'            DoEvents
'            .ViewReport
'            DoEvents
'            .Zoom 2
'            End With
'        oForm.Show
'    Else
'        '''''''''''''''''''''''''''''''''''''''''''''
'        ' Print Report directly, no display
'        '''''''''''''''''''''''''''''''''''''''''''''
'        oReport.PrintOut False
'    End If
'
'    PrintSF603A = True
'
'    GoTo CleanUp
'
'ErrHandler:
'  gobjshared.ShowVBError Err.Number, Err.Description, "DentalForms.PrintSF603A", "CHCSII_Forms", vbExclamation
'CleanUp:
'    Set oReport = Nothing
'    Set oForm = Nothing
'    Set rs = Nothing
'    Set gcolTeethDx = Nothing
'    Set gcolTeethTx = Nothing
'    Screen.MousePointer = vbDefault
End Function


Private Sub Class_Initialize()
'-- Iniatialized shared global object
    Set gobjshared = New CWShared
    
   Set moConn = New CHCSII_CONN.Conn
   Set moDAS = moConn.CHCSII_DAS(Auto)
    
    Set oTeethSVG = oSF.Singleton
'''    With frmTeethPictures
'''        .tmrTeeth.Enabled = True
'''        Set .Parent = Me
'''    End With
End Sub

Private Sub Class_Terminate()
    Set moDAS = Nothing
    Set mrsMain = Nothing
End Sub
'******************************************************************************************
'Name: GetDepartService
'Description: This function return service name based on first letter of PatientCategory
'******************************************************************************************
Public Function GetDepartService(ByVal vsPatientCategory As String) As String
    Select Case UCase(left(vsPatientCategory, 1))
        Case "A"
            GetDepartService = "Army"
        Case "B"
            GetDepartService = "NOAA (National Oceanic and Atmospheric Administration)"
        Case "C"
            GetDepartService = "Coast Guard"
        Case "F"
            GetDepartService = "Air Force"
        Case "K"
            GetDepartService = "Civilian"
        Case "M"
            GetDepartService = "Marines"
        Case "N"
            GetDepartService = "Navy"
        Case "P"
            GetDepartService = "Public Health Service"
        Case "R"
            GetDepartService = "NATO Reciprocal Agreement"
        Case Else
            GetDepartService = "Unknown"
    End Select

End Function
'**************************************************************************
'Name: GetRelationshipToSponsor
'Description:   Access database to retrieve patient relationship to sponsor
'Parameter:     vsFMP As String
'Call By:       PrintSF603A2
'Return:        String
'History
'Date       Developer   Description
'========   ==========  ===================================================
'10/05/01   Wilson Sun  Added SQL for LDDB SCR: 16997, 17024
'11/29/00   Wilson Sun  Created
'**************************************************************************
Private Function GetRelationshipToSponsor(ByVal vsFMP As String) As String
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    
    On Error GoTo ErrHandler
    sSQL = "select T1.DESCRIPTION as RELATIONSHIP" & vbNewLine & _
            "from NCIDS T1, NCIDS T2" & vbNewLine & _
            "where T1.NCID = T2.NCID" & vbNewLine & _
            "and T1.DOMAINNCID = 110627 " & vbNewLine & _
            "and T2.DESCRIPTION = '" & vsFMP & "'"
    
    Set rs = moDAS.OpenRecordset(sSQL)
    If Not (rs.EOF Or rs.BOF) Then
        GetRelationshipToSponsor = rs.Fields.Item("RELATIONSHIP").Value
    Else
        GetRelationshipToSponsor = ""
    End If
    GoTo Cleanup

ErrHandler:
  gobjshared.ShowVBError Err.Number, Err.Description, "DentalForms.GetRelationshipToSponsor", "CHCSII_Forms", vbExclamation
Cleanup:
    Set rs = Nothing
End Function
'**************************************************************************
' Name:         GetClinicAddress
' Description:  Access database to retrieve clinic name and address
' Parameter:    psClinicNCID
' Return:       Clinic Address
'**************************************************************************
Private Function GetClinicAddress(ByVal psClinicNCID As String) As String
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    
    On Error GoTo ErrHandler
    
    If gobjshared.IsAppMode(modeITT) = False Then
        sSQL = "select STREET_ADDRESS || CITY || STATE || ZIP as ADDRESS " & vbCrLf & _
            "from CLINIC " & vbCrLf & _
            "where ncid = " & psClinicNCID
    Else
        sSQL = "select STREET_ADDRESS + CITY + STATE + ZIP as ADDRESS " & vbCrLf & _
            "from CLINIC " & vbCrLf & _
            "where ncid = " & psClinicNCID
    End If
    
'    sSQL = "select STREET_ADDRESS ||', ' || CITY ||', ' || STATE ||', ' || ZIP as ADDRESS " & vbCrLf & _
'        "from CLINIC " & vbCrLf & _
'        "where ncid = " & psClinicNCID
        
    Set rs = moDAS.OpenRecordset(sSQL)
    
    GetClinicAddress = rs.Fields.Item("ADDRESS").Value
    GoTo Cleanup

ErrHandler:
  gobjshared.ShowVBError Err.Number, Err.Description, "DentalForms.GetClinicAddress", "CHCSII_Forms", vbExclamation
Cleanup:
    Set rs = Nothing
End Function
'*************************************************************************
'Function Name: FirstLastMiName
'Description: Convert full name as Last, First, Middle initial format
'*************************************************************************
Public Function FirstLastMiName(ByVal pFullName As String) As String
    Dim vName As Variant
    
    vName = Split(pFullName, " ", , vbTextCompare)
    If UBound(vName, 1) < 2 Then
        'Return original full name
        FirstLastMiName = pFullName
    Else
        'Return formatted name
        FirstLastMiName = vName(0) & " " & vName(1) & " " & left(vName(2), 1)
    End If

End Function
