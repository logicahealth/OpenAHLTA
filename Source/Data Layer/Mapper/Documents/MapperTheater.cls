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
Attribute VB_Name = "MapperTheater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements DL_Interfaces.iMapperDocuments

Private Const MODULE_NAME As String = "DM_OrderEntry.MapperTheater"

Private Sub iMapperDocuments_Init(Support As DL_Support.iParameters, Cache As DL_Support.CacheMapper)

    On Error GoTo ErrHandler
    
    Set mCacheMapper = Cache
    Set mConn = Support.Item("DBConn").Value
    
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub


Private Function iMapperDocuments_RetrieveInpatientAdmissions(PatientUnitNumber As String, FacilityNCID As String, Optional ApptID As String = "") As DTOs.DocAdmissions
    
    'Load the inpatient admissions and encounters from the db and return a DocAdmissions Object.
    'Optional ApptID allows user to filter by ApptID
    
On Error GoTo ErrHandler

Dim oDAS As ICHCSII_DAS
Dim sSQL As String
Dim sDiagSQL As String
Dim oRs As Recordset
Dim oDiagRs As Recordset
Dim oDocAdmissions As DocAdmissions
Dim oDocAdmission As DocAdmission
Dim sPrevInpatientID As Long
Dim oComm As MMMHISComm.Comm
    
    Set oComm = New MMMHISComm.Comm
    
    Set oDAS = mConn.CHCSII_DAS(Auto, NoTimeZoneConversion)
    
    sSQL = "SELECT INPATIENT_ID, ADMISSION_DATE, DISCHARGE_DATE, " _
            & " PROV_NAME, WARD_LOCATION_NAME, Specialty" _
            & " From V_Inpatient_Admissions" _
            & " WHERE (UNIT_NUMBER = " & PatientUnitNumber & ")" ' AND (FACILITY_NCID = " & FacilityNCID & ")"
    If ApptID <> "" Then
        sSQL = sSQL & " AND (APPT_ID = " & ApptID & ")"
    End If
    sSQL = sSQL & " ORDER BY INPATIENT_ID DESC"
    Set oRs = oDAS.OpenRecordset(sSQL)
    
    sDiagSQL = "SELECT INPATIENT_DIAGNOSIS.DESCRIPTION, INPATIENT_DIAGNOSIS.DIAGNOSIS_TYPE," _
        & " INPATIENT_DIAGNOSIS.INPATIENT_ID, INPATIENT_ADMISSION.UNIT_NUMBER" _
        & " From INPATIENT_ADMISSION, INPATIENT_DIAGNOSIS" _
        & " Where INPATIENT_ADMISSION.INPATIENT_ID = INPATIENT_DIAGNOSIS.INPATIENT_ID" _
        & " AND (INPATIENT_ADMISSION.UNIT_NUMBER = " & PatientUnitNumber & ")"
    Set oDiagRs = oDAS.OpenRecordset(sDiagSQL)
     
    Set oDocAdmissions = New DocAdmissions

    Do Until oRs.EOF
        Set oDocAdmission = New DocAdmission
        With oDocAdmission
            .InpatientID = oRs("Inpatient_ID")
            .IsInpatient = True
            
            'SF 62547
            If IsNull(oRs("Admission_Date")) Then
                .AdmissionDate = 0
            Else
                'The patient is admitted
                .AdmissionDate = oComm.GmtToLocal(oRs("Admission_Date"))
                oDiagRs.Filter = ""
                oDiagRs.Filter = "Inpatient_ID = " & .InpatientID & " and Diagnosis_Type = 'ADMISSION'"
                If Not oDiagRs.EOF Then
                    .PrimaryDiagnosis = "Admitting Dx=" & oDiagRs("DESCRIPTION")
                Else
                    .PrimaryDiagnosis = "Admitting Dx=NONE"
                End If
            End If
            
            If IsNull(oRs("Discharge_Date")) Then
                .DischargeDate = 0
            Else
                'The patient is discharged
                .DischargeDate = oComm.GmtToLocal(oRs("Discharge_Date"))
                oDiagRs.Filter = ""
                oDiagRs.Filter = "Inpatient_ID = " & .InpatientID & " and Diagnosis_Type = 'DISCHARGE'"
                If Not oDiagRs.EOF Then
                    .PrimaryDiagnosis = "Discharge Dx=" & oDiagRs("DESCRIPTION")
                Else
                    .PrimaryDiagnosis = "Discharge Dx=NONE"
                End If
            End If
            
            .ProviderName = oRs("Prov_Name") & vbNullString
            .Ward = oRs("Ward_Location_Name") & vbNullString
            .Specialty = oRs("Specialty") & vbNullString
        End With
        
        'Don't load duplicate admissions
        If sPrevInpatientID <> oDocAdmission.InpatientID Then
            oDocAdmissions.Add oDocAdmission, CStr(oDocAdmission.InpatientID)
            
            'Load the encounters for this admission
            If Not oDocAdmission Is Nothing Then
                Set oDocAdmission.DocEncounters = LoadInpatientEncounters(oRs("Inpatient_ID"), oDocAdmission)
            End If
        End If
        
        sPrevInpatientID = oDocAdmission.InpatientID
        
        oRs.MoveNext
    Loop
    
    oDiagRs.Close
    Set oDiagRs = Nothing
    
    oRs.Close
    Set oRs = Nothing
    
    Set iMapperDocuments_RetrieveInpatientAdmissions = oDocAdmissions
    
    Set oDocAdmissions = Nothing
    
Exit Function

ErrHandler:
    Set iMapperDocuments_RetrieveInpatientAdmissions = New DocAdmissions
    Err.Raise Err.Number, MODULE_NAME & ".RetrieveOtherOrderTypeCategories", Err.Description
    
    Exit Function
    Resume
End Function


Private Function LoadInpatientEncounters(InpatientID As Long, Admission As DocAdmission) As DocEncounters

'Loads all the encounters for an inpatient admission

Dim oDAS As ICHCSII_DAS
Dim sSQL As String
Dim oRs As Recordset
Dim oDocEncounters As DocEncounters
Dim oDocEncounter As DocEncounter

On Error GoTo ErrHnd

    Set oDAS = mConn.CHCSII_DAS(Auto, NoTimeZoneConversion)
        
    sSQL = "SELECT FACILITYNCID, ENCOUNTERNUMBER, ENCOUNTER_TYPE, CREATEDON, NAME, ENC_STATUS, WARD_LOCATION_NAME, Inpt_Hosp_Day, CreatedBy" _
        & " From V_INPATIENT_ENCOUNTERS" _
        & " Where Inpatient_ID = " & InpatientID _
        & " Order by CREATEDON DESC"
        
    Set oRs = oDAS.OpenRecordset(sSQL)
    
    Set oDocEncounters = New DocEncounters
    
    Do Until oRs.EOF
        Set oDocEncounter = New DocEncounter
        With oDocEncounter
            .IsInpatient = True
            .EncounterNumber = oRs("EncounterNumber")
            .FacilityNCID = oRs("FacilityNCID")
            .DocumentType = oRs("Encounter_Type") & vbNullString
            .DateTime = IIf(IsNull(oRs("CreatedOn")), 0, oRs("CreatedOn"))
            .ProviderName = oRs("Name") & vbNullString
            .Status = oRs("Enc_Status") & vbNullString
            .Ward = oRs("Ward_Location_Name") & vbNullString
            .HospDay = Val(oRs("Inpt_Hosp_Day") & vbNullString)
            .CreatedBy = oRs("CreatedBy") & vbNullString
            '.UpdateLock = CBool(Val(oRs("UpdateLock") & vbNullString))
            Set .Admission = Admission
        End With
        
        oDocEncounters.Add oDocEncounter, CStr(oDocEncounter.EncounterNumber)
                
        oRs.MoveNext
    Loop
    
    Set LoadInpatientEncounters = oDocEncounters
    
    Set oDocEncounters = Nothing
    
Exit Function

ErrHnd:

    Err.Raise Err.Number, Err.Source, "Error in EncounterPrevious.DocMapper.LoadInpatientEncounters: " & Err.Description
    Set LoadInpatientEncounters = New DocEncounters
    
End Function


