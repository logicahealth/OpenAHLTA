VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DTRptDD2322"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Wilson23
'********************************************************
'DTRpt2322.cls is only for Report DD2322 functionality
'********************************************************
Option Explicit
Private moPatientSearch As CHCSII_Search.PatientSearch
Private moDAS As CHCSII_DAS
Private mrsMain As ADODB.Recordset
'***********************************************************************
'Function name: BuildReport
'Description:   This is the main interface for calling to print SF603A
'Parameters:
'   poPatient               - Patient object for Demographic Info
'   plEncounterFacilityNCID - Retrieve related data
'   plEncounterNumber       - Retrieve related data
'Return Value:  Ture, sucessful; False, otherwise
'***********************************************************************
Public Function BuildReport(ByVal poReport As CRDD2322, _
                            Optional ByVal poPatient As Patient, _
                            Optional ByVal plEncounterFacilityNCID As Long = 0, _
                            Optional ByVal plEncounterNumber As Long = 0) As Boolean
    
    Dim oPatientItem As PatientItem
    Dim oClinicOps As CHCSII_ClinicClient.ClinicOps
    Dim oClinic As Clinic
    Dim sSQL As String

    '''''''''''''''''''''''''''''''
    'Variable for Report Formula
    '''''''''''''''''''''''''''''''
    ' For patient
    Dim sPatientName As String
    Dim sSex As String
    Dim sSSN As String
    Dim sDateOfBirth As String
    Dim sFMP As String
    Dim sOrg As String
    Dim sAge As String
    Dim sHomePhoneNo As String
    Dim sPatientCategory As String
    
    'For Clinic
    Dim sTreatmentFacilityName As String
    Dim sMaillingAddress As String
      
    'To make GetDepartService() available
    Dim oDentalForms As New CHCSII_Forms.DentalForms
    
    On Error GoTo ErrHandler

    BuildReport = False

    ''''''''''''''''''''''''''''''''''''''''''''''''''
    'Setup teeth surface
    ''''''''''''''''''''''''''''''''''''''''''''''''''

    ''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Get main body record set
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    sSQL = "select T2.CREATE_DATE_TIME, T2.TREATMENT_TECHNIQUE_COMMENT, T1.DRC " & _
        "from DENTAL.DENTAL_ENCOUNTER T1, DENTAL.NOTE_TREATMENT T2 " & _
        "Where T1.DENTAL_ENCOUNTER_ID = T2.DENTAL_ENCOUNTER_ID " & _
        "and T1.PATIENT_UNIT_ID = " & poPatient.UnitNumber & _
        " order by 1 desc"
'    Set mrsMain = moDAS.OpenRecordset(sSQL)
'    poReport.Database.SetDataSource mrsMain
'    Set mrsMain = Nothing

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Assign object variable to local variable
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''Patient Info
    Set oPatientItem = moPatientSearch.GetBy(poPatient.UnitNumber, "")
    With poPatient
        'Since Patient only returns FULL name, so we need reconstruction
        sPatientName = oPatientItem.Last & ", " & oPatientItem.First & " " & Left(oPatientItem.Middle, 1)
        sSex = .Sex
        sSSN = .SSN
        sDateOfBirth = Format(CDate(.DateOfBirth), "dd mmm yyyy")
        sFMP = .FMP
        sAge = Format(CDate(Now - CDate(.DateOfBirth)), "yy")
        sHomePhoneNo = .HomePhone
        sPatientCategory = .PatientCategory
        sOrg = IIf(.UnitIdentificationCode <> "", .UnitIdentificationCode, "N/A")
    End With
    
    '''Clinic Info
    Set oClinicOps = New CHCSII_ClinicClient.ClinicOps
    Set oClinic = oClinicOps.GetClinic(cwobjLogonEnviron.NursingDivisionNCID)
    With oClinic
        sTreatmentFacilityName = .Name
        sMaillingAddress = .StreetAddress & ", " & .City & ", " & .State & " " & .ZipCode
    End With
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Set formula value report
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    With poReport
        .fldPatientName.Field.Text = "'" & sPatientName & "'"
        .FormulaFields(2).Text = "'" & sSSN & "'" 'SSN formated from Crystal Designer
        .fldGrade.Field.Text = "'" & oPatientItem.RankGrade & "'"
        .fldAge.Field.Text = "'" & sAge & "'"
        .fldDentalOfficer.Field.Text = "'" & cwobjLogonEnviron.UserName & "'"
        .fldOrgDutyHomePhoneNo.Field.Text = "'" & sOrg & ", " & _
            oDentalForms.GetDepartService(sPatientCategory) & ", " & sHomePhoneNo & "'"
        .fldTreatmentFacilityNmae.Field.Text = "'" & sTreatmentFacilityName & "'"
        .fldMaillingAddress.Field.Text = "'" & sMaillingAddress & "'"
        .fldAutovonNo.Field.Text = "" 'Not know yet
    End With


    BuildReport = True

    GoTo CleanUp

ErrHandler:
  gobjshared.ShowVBError Err.Number, Err.Description, "DTRpt2322.BuildReport", "CHCSII_Reports", vbExclamation
CleanUp:
    Set mrsMain = Nothing
    Set moDAS = Nothing
    Set oDentalForms = Nothing
    Set oClinicOps = Nothing
    Set oClinic = Nothing

End Function
Private Sub Class_Initialize()
    Set moDAS = New CHCSII_DAS
    Set mrsMain = New ADODB.Recordset
    Set moPatientSearch = New CHCSII_Search.PatientSearch
End Sub
Private Sub Class_Terminate()
    Set moDAS = Nothing
    Set mrsMain = Nothing
    Set moPatientSearch = Nothing
End Sub


