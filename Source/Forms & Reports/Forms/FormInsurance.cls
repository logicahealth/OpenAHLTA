VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FormInsurance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const sMODULE_NAME As String = "FormInsurance.cls"
Private moCollRecSet As Collection
Private moCollFormulas As Collection
Private moPatientDemographics As PatientDemographics
'

Private Function GetInsurance(sUnitNumber, _
                             sFacilityNumber As String, _
                             oRec As ADODB.Recordset, pEncounterNumber As String) As Boolean

Dim sStmt As String
Dim i As Integer
Dim i2 As Integer

GetInsurance = False
On Error GoTo ErrHandler

'-- Create SQL'-----------------------------------------------------------------
    If gobjshared.UseSQLServer Then
        sStmt = "select p.INSCOMPANYNAME as InsuranceName," & vbCrLf & _
            "ADDRESS1 + ' ' + ADDRESS2 + ' ' + ADDRESS3 as Address," & vbCrLf & _
        "INSCOMPANYPHONE as Phone," & vbCrLf & _
        "p.GROUPNUMBER as InsuredGroupNumber," & vbCrLf & _
        "p.GROUPNAME as InsuredGroupName," & vbCrLf & _
        "p.SUBSCRIBERNAME as InsuredGroupEmpName," & vbCrLf & _
        "p.INSIDNUMBER as PolicyNumber, ' ' as ToPatientRel," & vbCrLf & _
        "CONVERT(VARCHAR(25), p.STARTDATE,106) as PlanEffectiveTime," & vbCrLf & _
        "CONVERT(VarChar(25), p.STOPDATE, 106) As PlanExpirationTime" & vbCrLf & _
        "From PATIENT_INSURANCE p"
        If pEncounterNumber = "" Then
            sStmt = sStmt & vbCrLf & "Where p.UNIT_NUMBER = " & sUnitNumber
        Else
            sStmt = sStmt & vbCrLf & "Where p.ENCOUNTER_NUMBER = " & pEncounterNumber
        End If
'        e.hems_encounter_id = em.hems_encounter_id and
'        er.role_cid = decode(e.priority,1,1800000001,2,67580,3,67581,4,67582,5,67583,6,
'        67584,7,14516082,8,
'        14516083,9,14516084,10,14516085,11,14614816) and
'        er.role_cid=
'        any(1800000001,67580,67581,67582,67583,67584,14516082,14516083,
'        14516084,14516085,14614816) and
'        er.hems_encounter_id = em.hems_encounter_id and
'        em.facility_cid =  ms.facility_cid and
'        em.encounter_number = 'E' || ms.medical_record_number and
'        ms.facility_cid = 1046961 and
'        ms.unit_number = 99990077
    Else
        sStmt = "select substr(rs.representation,1,60) as InsuranceName, " & vbNewLine & _
                "er.Street1||' '||er.Street2||' '||er.City||', '||er.State||' '||e.Post_Code as Address, " & vbNewLine & _
                "rtrim(er.area_code)|| decode(Length(er.area_code),Null,'','-') ||substr(er.local_number,1,3)||decode(Length(er.local_number),Null,'','-') ||substr(er.local_number,4,4) as Phone, " & vbNewLine & _
                "e.insured_group_number as InsuredGroupNumber, " & vbNewLine & _
                "e.insured_group_name as InsuredGroupName, " & vbNewLine & _
                "e.insured_group_emp_name as InsuredGroupEmpName, " & vbNewLine & _
                "e.policy_number as PolicyNumber, " & vbNewLine & _
                "substr(r2.representation,1,30) as ToPatientRel, " & vbNewLine & _
                "TO_CHAR(e.plan_effective_time,'DD MON YYYY') as PlanEffectiveTime, " & vbNewLine & _
                "TO_CHAR(e.plan_expiration_time,'DD MON YYYY') as PlanExpirationTime " & vbNewLine
        
        sStmt = sStmt & "From rsform rs, " & vbNewLine & _
                "rsform_context rsc, " & vbNewLine & _
                "rsform r2, " & vbNewLine & _
                "rsform_context r2c, " & vbNewLine & _
                "enc_insurance e, " & vbNewLine & _
                "enc_role er, " & vbNewLine & _
                "enc_main em, " & vbNewLine & _
                "mmi_site ms " & vbNewLine
                
        sStmt = sStmt & "Where " & vbNewLine & _
                "rsc.preferred_score = 0 and " & vbNewLine & _
                "rsc.context_ncid = 2000 and " & vbNewLine & _
                "rsc.rsform_id = rs.rsform_id and " & vbNewLine & _
                "rs.ncid = e.insurance_cid and " & vbNewLine & _
                "r2c.preferred_score = 0 and " & vbNewLine & _
                "r2c.context_ncid = 367 and " & vbNewLine & _
                "r2c.rsform_id = r2.rsform_id and " & vbNewLine & _
                "r2.ncid = e.to_patient_rel_cid and " & vbNewLine & _
                "e.hems_encounter_id = em.hems_encounter_id and " & vbNewLine & _
                "er.role_cid = decode(e.priority,1,1800000001,2,67580,3,67581,4,67582,5,67583,6, " & vbNewLine & _
                "67584,7,14516082,8, " & vbNewLine & _
                "14516083,9,14516084,10,14516085,11,14614816) and " & vbNewLine & _
                "er.role_cid= " & vbNewLine & _
                "any(1800000001,67580,67581,67582,67583,67584,14516082,14516083, " & vbNewLine & _
                "14516084,14516085,14614816) and " & vbNewLine & _
                "er.hems_encounter_id = em.hems_encounter_id and " & vbNewLine & _
                "em.facility_cid =  ms.facility_cid and " & vbNewLine & _
                "em.encounter_number = 'E' || ms.medical_record_number and " & vbNewLine & _
                "ms.facility_cid = " & sFacilityNumber & " and " & vbNewLine & _
                "ms.unit_number = " & sUnitNumber
    End If
    
    If gobjshared.UseSQLServer Then
        Set oRec = goDBDAS.OpenRecordset(sStmt, adOpenDynamic, adLockOptimistic)
    Else
        Set oRec = goDBDAS.OpenRecordset(sStmt)
    End If
    
    If Not ReplaceNulls(oRec) Then GoTo CleanUp
    Set oRec.ActiveConnection = Nothing

    GetInsurance = True
    GoTo CleanUp

ErrHandler:
    'Resume Next
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".GetInsurance", App.EXEName, vbExclamation
CleanUp:
    Exit Function
End Function

Public Function BuildInsurance(sPatientUnitNumber As String, _
                               sFacilityNCID As String, _
                               oReport As CRAXDRT.Report, pEncounterNumber As String) As Boolean

On Error GoTo ErrHandler
BuildInsurance = False

Dim oCrystal As Crystal
Dim oRec As ADODB.Recordset
DoEvents

'-- Get Data
'----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection

    '-- Get Insurance
    '----------------------------------------------------------
    If Not GetInsurance(sPatientUnitNumber, sFacilityNCID, oRec, pEncounterNumber) Then GoTo CleanUp
    moCollRecSet.Add oRec, "Insurance"

    '-- Get Patient Demographics
    '----------------------------------------------------------
    If Not moPatientDemographics.RecordGetDemographics(sPatientUnitNumber, oRec) Then GoTo CleanUp
    moCollRecSet.Add oRec, "demographics"
    DoEvents

    '-- Get Patient Address
    '----------------------------------------------------------
    If Not moPatientDemographics.RecordGetAddress(sPatientUnitNumber, oRec) Then GoTo CleanUp
    moCollRecSet.Add oRec, "Address"
    DoEvents

'-- Set Formulas
'-----------------------------------------------------------------------------
    Set moCollFormulas = Nothing
    Set moCollFormulas = New Collection
    moCollFormulas.Add "'" & Format(Now, gsDefaultFormat & "  HHMM") & "'", "TodaysDateTime"

'-- Create Report Object
'-----------------------------------------------------------------------------
    Set oCrystal = New Crystal
    Set oReport = Nothing
    If Not oCrystal.ReportOpen("InsuranceForm.rpt", oReport) Then GoTo CleanUp
    If Not oCrystal.AssignRecordSets(moCollRecSet, oReport) Then GoTo CleanUp
    If Not oCrystal.AssignFormulas(moCollFormulas, oReport) Then GoTo CleanUp

    BuildInsurance = True
    GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".BuildInsurance", App.EXEName, vbExclamation
CleanUp:
    Set oRec = Nothing
    Exit Function
End Function

Private Sub Class_Initialize()
    Set moPatientDemographics = New PatientDemographics
End Sub

Private Sub Class_Terminate()
    Set moPatientDemographics = Nothing
End Sub


