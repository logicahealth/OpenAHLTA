VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DTRptMontlyProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Wilson25
'********************************************************
'DTRpt2322.cls is only for Report DD2322 functionality
'********************************************************
Option Explicit

Private moDAS As CHCSII_DAS
Private mrsMain As ADODB.Recordset

'***********************************************************************
'Function name: BuildReport
'Description:   This is the main interface for calling to print SF603A
'Parameters:
'   poPatient             - Patient object for Demographic Info
'   plEncounterFacilityNCID - Retrieve related data
'   plEncounterNumber       - Retrieve related data
'Return Value:  Ture, sucessful; False, otherwise
'***********************************************************************
Public Function BuildReport(ByVal poReport As CRMonthlyProvider, _
                            Optional ByVal poPatient As Patient, _
                            Optional ByVal plEncounterFacilityNCID As Long = 0, _
                            Optional ByVal plEncounterNumber As Long = 0) As Boolean
    BuildReport = True
    Exit Function
    
    'Empty function above
    'To do the following

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
    Dim sOrg As String
    On Error GoTo ErrHandler

    BuildReport = False

    ''''''''''''''''''''''''''''''''''''''''''''''''''
    'Setup teeth surface
    ''''''''''''''''''''''''''''''''''''''''''''''''''

    ''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Get main body record set
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    'TO DO
    sSQL = ""
    Set mrsMain = moDAS.OpenRecordset(sSQL)
    poReport.Database.SetDataSource mrsMain
    Set mrsMain = Nothing

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
    End With
    ''' Sponsor Info
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
    Dim oDentalForms As New CHCSII_Forms.DentalForms
    With poReport
'        .fldPatientName.Field.Text = "'" & oDentalForms.FirstLastMiName(sPatientName) & "'"
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
    End With


    BuildReport = True

    GoTo CleanUp

ErrHandler:
  gobjshared.ShowVBError Err.Number, Err.Description, "DTRpt2322.BuildReport", "CHCSII_Reports", vbExclamation
CleanUp:
    Set mrsMain = Nothing
    Set moDAS = Nothing
    Set oDentalForms = Nothing
End Function
Private Sub Class_Initialize()
    Set moDAS = New CHCSII_DAS
    Set mrsMain = New ADODB.Recordset
End Sub
Private Sub Class_Terminate()
    Set moDAS = Nothing
    Set mrsMain = Nothing
End Sub


