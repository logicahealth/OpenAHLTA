VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RepCPGGoals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private moCollRecSet As Collection
Private moCollFormulas As Collection

Private Const cHGB_A1C_NCID = "8571"
Private Const cMICROALBUMIN = "31959"
Private Const cLDL = "6094"
Private Const cHDL = "6090"
Private Const cTriglycerides = "21552"
Private Const cWEIGHT_NCID = "2178"
Private Const cHEIGHT_NCID = "110675"
Private Const cBP_SYSTOLIC = "1985"
Private Const cBP_DIASTOLIC = "1976"
Private Const cBMI_NCID = "16752178"

Private moRecDemographics As ADODB.Recordset
Private moRecHGB As ADODB.Recordset
Private moRecMicro As ADODB.Recordset
Private moRecLipid As ADODB.Recordset
Private moRecBloodSys As ADODB.Recordset
Private moRecBloodDia As ADODB.Recordset
Private moRecBodyMass As ADODB.Recordset
Private moRecEye As ADODB.Recordset
Private moRecFoot As ADODB.Recordset
Private moRecHealth As ADODB.Recordset


Private msFacilityLabContextNCID As String
Private msFacilityNCID As String

Public Function BuildReport(sReportName As String, _
                            oReportItem As ReportItem, _
                            oReport As CRAXDRT.Report, _
                            oCollFormulas As Collection) As Boolean

On Error GoTo ErrHandler

BuildReport = False

Dim oCrystal As Crystal
Dim oRec As ADODB.Recordset
Dim sSponsorUnitNumber As String
Dim sUnitNumber As Variant
Dim oPatientsColl As Collection
Dim oLabRec As ADODB.Recordset

If Not LabResultSearch("LDL", oLabRec) Then GoTo CleanUp

'-- Get Data
'----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection

    '-- Initialize Recordsets
    '----------------------------------------------------------
    If Not Initialize_RecordSets() Then GoTo CleanUp

    '-- Get Patients
    '----------------------------------------------------------
    If Not GetPatients(oPatientsColl, oReportItem) Then GoTo CleanUp

 '   Set oPatientsColl = Nothing
 '   Set oPatientsColl = New Collection
 '   oPatientsColl.Add "204110"
    For Each sUnitNumber In oPatientsColl

        
        '-- Get Demographics
        '----------------------------------------------------------
        If Not GetDemographics(CStr(sUnitNumber)) Then GoTo CleanUp

        '-- Get HGB
        '----------------------------------------------------------
        If Not GetHGB(CStr(sUnitNumber)) Then GoTo CleanUp

        '-- Get Microalbumin
        '----------------------------------------------------------
        If Not GetMicroalbumin(CStr(sUnitNumber)) Then GoTo CleanUp
        
        '-- Get Lipid
        '----------------------------------------------------------
        'If Not GetLipid(CStr(sUnitNumber)) Then GoTo CleanUp

        '-- Get Body Mass
        '----------------------------------------------------------
        If Not GetBodyMassIndex(CStr(sUnitNumber)) Then GoTo CleanUp

        '-- Get Blood Pressure Systolic
        '----------------------------------------------------------
        If Not GetBloodPressureSystolic(CStr(sUnitNumber)) Then GoTo CleanUp
        
        '-- Get Blood Pressure Diastolic
        '----------------------------------------------------------
        If Not GetBloodPressureDiastolic(CStr(sUnitNumber)) Then GoTo CleanUp

        '-- Get Eye Exam
        '----------------------------------------------------------
        If Not GetEyeExam(CStr(sUnitNumber)) Then GoTo CleanUp

        '-- Get Health Status
        '----------------------------------------------------------
        If Not GetHealthStatus(CStr(sUnitNumber)) Then GoTo CleanUp

        '-- Get Foot Exam
        '----------------------------------------------------------
        If Not GetFootExam(CStr(sUnitNumber)) Then GoTo CleanUp


    Next

    moCollRecSet.Add moRecDemographics, "demographics"
    moCollRecSet.Add moRecHGB, "hgb"
    moCollRecSet.Add moRecMicro, "microalbumin"
    moCollRecSet.Add moRecLipid, "lipid"
    moCollRecSet.Add moRecBodyMass, "bmi"
    moCollRecSet.Add moRecBloodSys, "bp"
    moCollRecSet.Add moRecBloodDia, "bpd"
    moCollRecSet.Add moRecEye, "eyeexam"
    moCollRecSet.Add moRecHealth, "healthstatus"
    moCollRecSet.Add moRecFoot, "footexam"

'-- Create Report Object
'-----------------------------------------------------------------------------
    If Not CreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo CleanUp

BuildReport = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepCPGGoals.BuildReport", "CHCSII_Reports", vbExclamation
CleanUp:
    Set oCrystal = Nothing
    Set oRec = Nothing
    
End Function

Private Function GetCount(sNCID As String) As Integer

Dim sStmt As String

sStmt = "select count(*) from rsform where ncid = " & sNCID

goDBSqlOps.Execute sStmt

GetCount = goDBSqlOps.Value(1)

End Function

Private Function GetHGB(sUnitNumber As String) As Boolean

On Error GoTo ErrHandler
 
Dim sStmt As String
Dim sGoalValue As String
Dim sCurrentValue As String
Dim sCurrentDate As String
Dim sAbnormal As String
Dim sGoalOpType As String
Dim sGoal As String
Dim oRec As ADODB.Recordset

Dim sFacility As String
Dim sIEN As String


'-- Test Cases
'-- 156953 JACKSON, JEANETTE
'-- 157412 MINSER, GERALD PEPPEL
'-- 157518 FANNING, DOROTHY

'Definitions for NCIDs
'----------------------------------------------
'-- 8571 - Hemoglobin A1C/Hemoglobin Total
'-- 485 -
'-- 1040 -
'-- 1037 -

GetHGB = False


'-- Get HGB A1C latest lab result
'-----------------------------------------------
sStmt = "Select " & _
            "TO_CHAR(e.unit_number) as Unit_Number, " & vbNewLine & _
            "To_Char(e.event_start_gmtime, 'DD MON YYYY') as Current_Date, " & vbNewLine & _
            "o.numeric as Current_Value, " & vbNewLine & _
            "decode(o.abnormal_flag_ncid,1040,'L',1037,'H',NULL,'N','U') as Abnormal_Flag " & vbNewLine & _
        "From " & vbNewLine & _
            "observation o, " & vbNewLine & _
            "event e " & vbNewLine & _
        "Where " & vbNewLine & _
            "o.observation_ncid = 8571 and " & vbNewLine & _
            "o.event_id = e.event_id and " & vbNewLine & _
            "e.event_type_ncid = 485 and " & vbNewLine & _
            "e.event_subtype_ncid in (" & cHGB_A1C_NCID & ") and " & vbNewLine & _
            "e.unit_number = " & sUnitNumber & " " & vbNewLine & _
        "order by e.event_start_gmtime desc"
            
Set oRec = goDBDAS.OpenRecordset(sStmt)

If oRec.RecordCount > 0 Then
    sCurrentValue = oRec.Fields("Current_Value")
    sCurrentDate = oRec.Fields("Current_Date")
    sAbnormal = oRec.Fields("Abnormal_Flag")
Else
    sCurrentValue = ""
    sCurrentDate = ""
    sAbnormal = ""
End If

'-- Get Goal Value
'-----------------------------------------------
If Not GetPatientMatchingGoals(oRec, sUnitNumber, "4") Then GoTo CleanUp

If oRec.RecordCount <> 0 Then

    sIEN = oRec.Fields("Code")

    Do While Not oRec.EOF
        If IsLAB_IEN(sIEN, cHGB_A1C_NCID) Then
            sGoalValue = oRec.Fields("value")
            sGoalOpType = oRec.Fields("Goal_Op_Type")
            Exit Do
        End If
        oRec.MoveNext
    Loop
End If

moRecHGB.AddNew
moRecHGB.Fields("Unit_Number") = sUnitNumber
moRecHGB.Fields("Goal_Value") = sGoalValue
moRecHGB.Fields("Current_Value") = sCurrentValue
moRecHGB.Fields("Current_Date") = sCurrentDate
moRecHGB.Fields("ABNORMAL") = sAbnormal
moRecHGB.Fields("OP_TYPE") = sGoalOpType
moRecHGB.Update

GetHGB = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepCPGGoals.GetHGB", "CHCSII_Reports", vbExclamation
CleanUp:
    Exit Function
End Function
Private Function GetLabRecordSet(sUnitNumber As String, _
                                 oRec As ADODB.Recordset) As Boolean

On Error GoTo ErrHandler

GetLabRecordSet = False

Dim sStmt As String

sStmt = "Select " & _
            "To_Char(e.event_start_gmtime, 'dd mon yyyy') as Current_Date, " & vbNewLine & _
            "o.numeric as Current_Value, e.event_subtype_ncid as Lab_Type_NCID, o.OBSERVATION_NCID, " & vbNewLine & _
            "decode(o.abnormal_flag_ncid,1040,'L',1037,'H',NULL,'N','U') as Abnormal_Flag " & vbNewLine & _
        "From " & vbNewLine & _
            "observation o, " & vbNewLine & _
            "event e " & vbNewLine & _
        "Where " & vbNewLine & _
            " " & vbNewLine & _
            "o.event_id = e.event_id and " & vbNewLine & _
            "e.event_type_ncid = 485 and " & vbNewLine & _
            "e.unit_number = " & sUnitNumber & " " & vbNewLine & _
        "order by e.event_start_gmtime desc"

Set oRec = goDBDAS.OpenRecordset(sStmt)

GetLabRecordSet = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepCPGGoals.GetLabRecordSet", "CHCSII_Reports", vbExclamation
CleanUp:
    Exit Function
End Function

Private Function GetPatientMatchingGoals(oRec As ADODB.Recordset, _
                                         sUnitNumber As String, _
                                         sGoalType As String, _
                                         Optional sId As String = "") As Boolean

Dim sStmt As String

GetPatientMatchingGoals = False
                                         
sStmt = "Select " & _
            "p.value, " & vbNewLine & _
            "substr(p.references, 1, Instr(p.references,CHR(9)) - 1 ) as Code," & vbNewLine & _
            "p.Goal_Op_Type, p.date_created " & vbNewLine & _
        "From " & vbNewLine & _
            "pocgoal p  " & vbNewLine & _
        "Where " & vbNewLine
            
            If sId <> "" Then
                sStmt = sStmt & "substr(p.references, 1, Instr(p.references,CHR(9)) - 1 ) = '" & sId & "' and " & vbNewLine
            End If

sStmt = sStmt & _
            "goal_type = " & sGoalType & " and " & vbNewLine & _
            "p.unit_number = " & sUnitNumber & " " & vbNewLine

sStmt = sStmt & "Union " & vbNewLine

sStmt = sStmt & _
        "Select " & _
            "p.value, " & vbNewLine & _
            "substr(p.references, 1, Instr(p.references,CHR(9)) - 1 ) as NCID," & vbNewLine & _
            "p.Goal_Op_Type, p.date_created " & vbNewLine & _
        "From " & vbNewLine & _
            "pocgoal p,  " & vbNewLine & _
            "pocdef pd, " & vbNewLine & _
            "cpgregistry c " & vbNewLine & _
        "Where " & vbNewLine
        
            If sId <> "" Then
                sStmt = sStmt & "substr(p.references, 1, Instr(p.references,CHR(9)) - 1 ) = '" & sId & "' and " & vbNewLine
            End If
            
sStmt = sStmt & _
            "p.goal_type = " & sGoalType & " and " & vbNewLine & _
            "p.poc_id = pd.pocdef_id and " & vbNewLine & _
            "pd.cpg_id = c.cpg_id and " & vbNewLine & _
            "c.cpg_id is not null and " & vbNewLine & _
            "c.unit_number = " & sUnitNumber & " " & vbNewLine & _
        "order by date_created desc "

Set oRec = goDBDAS.OpenRecordset(sStmt)

GetPatientMatchingGoals = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepCPGGoals.GetPatientMatchingGoals", "CHCSII_Reports", vbExclamation
CleanUp:
    Exit Function
End Function

Private Function IsLAB_IEN(sIEN As String, _
                           sLAB_NCID As String) As Boolean

IsLAB_IEN = False

Dim sStmt As String
Dim sFacility As String
Dim sContextNCID As String


sFacility = msFacilityNCID

If Not GetContextNCID(sFacility, "LAB", sContextNCID) Then GoTo CleanUp

sStmt = "Select " & vbNewLine & _
            "r.NCID " & vbNewLine & _
        "From " & vbNewLine & _
            "rsform_context rc, " & vbNewLine & _
            "rsform r " & vbNewLine & _
        "where " & vbNewLine & _
            "rc.context_ncid = " & sContextNCID & " and " & vbNewLine & _
            "rc.rsform_id = r.rsform_id and " & vbNewLine & _
            "r.representation = '" & sIEN & "' and " & vbNewLine & _
            "r.ncid in (" & sLAB_NCID & ")"

goDBSqlOps.Execute sStmt

If goDBSqlOps.EOF = False Then IsLAB_IEN = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepCPGGoals.IsLAB_IEN", "CHCSII_Reports", vbExclamation
CleanUp:
    Exit Function
End Function


Private Function GetLAB_IEN(sLAB_NCID As String, _
                            sLAB_IEN As String, _
                            sFacilityContextNCID As String) As Boolean
                           
GetLAB_IEN = False

Dim sStmt As String

sStmt = "select " & vbNewLine & _
            "r.representation as IEN " & vbNewLine & _
        "from " & vbNewLine & _
            "rsform_context rc,  " & vbNewLine & _
            "rsform r,  " & vbNewLine & _
        "where " & vbNewLine & _
            "rc.context_ncid = " & sFacilityContextNCID & " and " & vbNewLine & _
            "rc.rsform_id = r.rsform_id and " & vbNewLine & _
            "r.ncid = " & sLAB_NCID

goDBSqlOps.Execute sStmt

If goDBSqlOps.EOF Then
    sLAB_IEN = "0"
Else
    sLAB_IEN = goDBSqlOps.Value("IEN")
End If


GetLAB_IEN = True
GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepCPGGoals.GetLAB_IEN", "CHCSII_Reports", vbExclamation
CleanUp:
    Exit Function
End Function

Private Function GetContextNCID(sFACILITY_NCID As String, _
                                sType As String, _
                                sFacilityContextNCID As String) As Boolean
                           
GetContextNCID = False

Dim sStmt As String

sStmt = "select " & vbNewLine & _
            "fc.context_ncid " & vbNewLine & _
        "from " & vbNewLine & _
            "facility_ncid_correlation fc  " & vbNewLine & _
        "where " & vbNewLine & _
            "fc.Type = '" & sType & "' and " & vbNewLine & _
            "fc.facility_ncid = " & sFACILITY_NCID

goDBSqlOps.Execute sStmt

If goDBSqlOps.EOF Then
    sFacilityContextNCID = "0"
Else
    sFacilityContextNCID = goDBSqlOps.Value("context_ncid")
End If


GetContextNCID = True
GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepCPGGoals.GetContextNCID", "CHCSII_Reports", vbExclamation
CleanUp:
    Exit Function
End Function


Private Function GetMicroalbumin(sUnitNumber As String) As Boolean

On Error GoTo ErrHandler
 
Dim sStmt As String
Dim sGoalValue As String
Dim sCurrentValue As String
Dim sCurrentDate As String
Dim sAbnormal As String
Dim sGoalOpType As String
Dim sGoal As String
Dim oRec As ADODB.Recordset

Dim sFacility As String
Dim sIEN As String


'-- Test Cases
'-- 156953 JACKSON, JEANETTE
'-- 157412 MINSER, GERALD PEPPEL
'-- 157518 FANNING, DOROTHY

'Definitions for NCIDs
'----------------------------------------------
'-- 31959 - Albumin
'-- 1040 -
'-- 1037 -

GetMicroalbumin = False


'-- Get HGB A1C latest lab result
'-----------------------------------------------
sStmt = "Select " & _
            "TO_CHAR(e.unit_number) as Unit_Number, " & vbNewLine & _
            "To_Char(e.event_start_gmtime, 'dd mon yyyy') as Current_Date, " & vbNewLine & _
            "o.numeric as Current_Value, " & vbNewLine & _
            "decode(o.abnormal_flag_ncid,1040,'L',1037,'H',NULL,'N','U') as Abnormal_Flag " & vbNewLine & _
        "From " & vbNewLine & _
            "observation o, " & vbNewLine & _
            "event e " & vbNewLine & _
        "Where " & vbNewLine & _
            "o.observation_ncid = 31959 and " & vbNewLine & _
            "o.event_id = e.event_id and " & vbNewLine & _
            "e.event_type_ncid = 485 and " & vbNewLine & _
            "e.event_subtype_ncid in (" & cMICROALBUMIN & ") and " & vbNewLine & _
            "e.unit_number = " & sUnitNumber & " " & vbNewLine & _
        "order by e.event_start_gmtime desc"
            
Set oRec = goDBDAS.OpenRecordset(sStmt)

If oRec.RecordCount > 0 Then
    sCurrentValue = oRec.Fields("Current_Value")
    sCurrentDate = oRec.Fields("Current_Date")
    sAbnormal = oRec.Fields("Abnormal_Flag")
Else
    sCurrentValue = ""
    sCurrentDate = ""
    sAbnormal = ""
End If

'-- Get Goal Value
'-----------------------------------------------
If Not GetPatientMatchingGoals(oRec, sUnitNumber, "4") Then GoTo CleanUp


If oRec.RecordCount <> 0 Then

    Do While Not oRec.EOF
    
        sIEN = oRec.Fields("Code")

        If IsLAB_IEN(sIEN, cMICROALBUMIN) Then
            sGoalValue = oRec.Fields("value")
            sGoalOpType = oRec.Fields("Goal_Op_Type")
            Exit Do
        End If
        oRec.MoveNext
    Loop

End If

moRecMicro.AddNew
moRecMicro.Fields("Unit_Number") = sUnitNumber
moRecMicro.Fields("Goal_Value") = sGoalValue
moRecMicro.Fields("Current_Value") = sCurrentValue
moRecMicro.Fields("Current_Date") = sCurrentDate
moRecMicro.Fields("ABNORMAL") = sAbnormal
moRecMicro.Fields("OP_TYPE") = sGoalOpType
moRecMicro.Update

GetMicroalbumin = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepCPGGoals.GetMicroalbumin", "CHCSII_Reports", vbExclamation
CleanUp:
    Exit Function
End Function



Private Function GetBodyMassIndex(sUnitNumber As String) As Boolean

On Error GoTo ErrHandler
 
Dim sStmt As String
Dim sGoalValue As String
Dim sCurrentValue As String
Dim sCurrentDate As String
Dim sAbnormal As String
Dim sGoalOpType As String
Dim sGoal As String
Dim oRec As ADODB.Recordset

Dim sFacility As String
Dim sIEN As String
Dim sWeight As String
Dim sHeight As String

GetBodyMassIndex = False


'-- Get BMI
'-----------------------------------------------
sStmt = "Select " & _
            "TO_CHAR(e.unit_number) as Unit_Number, " & vbNewLine & _
            "To_Char(e.event_start_gmtime, 'dd mon yyyy') as Current_Date, " & vbNewLine & _
            "ov.value_string as Current_Value " & vbNewLine & _
        "From " & vbNewLine & _
            "observation_value ov, " & vbNewLine & _
            "clinical_observation co, " & vbNewLine & _
            "event e " & vbNewLine & _
        "Where " & vbNewLine & _
            "ov.observation_seq = co.observation_seq and " & vbNewLine & _
            "ov.event_id = co.event_id and " & vbNewLine & _
            "co.observation_ncid = " & cBMI_NCID & " and " & vbNewLine & _
            "co.event_id = e.event_id and " & vbNewLine & _
            "e.event_subtype_ncid =  110669 and " & vbNewLine & _
            "e.event_type_ncid = 90782 and " & vbNewLine & _
            "e.unit_number = " & sUnitNumber & " " & vbNewLine & _
        "order by e.event_start_gmtime desc"

Set oRec = goDBDAS.OpenRecordset(sStmt)

sAbnormal = ""

'If oRec.RecordCount > 0 Then
'    sBMI = oRec.Fields("BMI")
'    sCurrentDate = oRec.Fields("Current_Date")
'End If


''-- Get Height
''-----------------------------------------------
'sStmt = "Select " & _
'            "TO_CHAR(e.unit_number) as Unit_Number, " & vbNewLine & _
'            "To_Char(e.event_start_gmtime, 'dd mon yyyy') as Current_Date, " & vbNewLine & _
'            "ov.value_number as Height " & vbNewLine & _
'        "From " & vbNewLine & _
'            "observation_value ov, " & vbNewLine & _
'            "clinical_observation co, " & vbNewLine & _
'            "event e " & vbNewLine & _
'        "Where " & vbNewLine & _
'            "ov.observation_seq = co.observation_seq and " & vbNewLine & _
'            "ov.event_id = co.event_id and " & vbNewLine & _
'            "co.observation_ncid = " & cHEIGHT_NCID & " and " & vbNewLine & _
'            "co.event_id = e.event_id and " & vbNewLine & _
'            "e.event_subtype_ncid =  110675 and " & vbNewLine & _
'            "e.event_type_ncid = 90782 and " & vbNewLine & _
'            "e.unit_number = " & sUnitNumber & " " & vbNewLine & _
'        "order by e.event_start_gmtime desc"
'
'Set oRec = goDBDAS.OpenRecordset(sStmt)
'
'sAbnormal = ""
'
'If oRec.RecordCount > 0 Then
'    sHeight = oRec.Fields("Height")
'End If

'-- Calculate BMI
'----------------------------------------------------------------------------------

If oRec.RecordCount > 0 Then
    sCurrentValue = oRec.Fields("Current_Value")
    sCurrentDate = oRec.Fields("Current_Date")
    sAbnormal = ""
Else
    sCurrentValue = ""
    sCurrentDate = ""
    sAbnormal = ""
End If



'If sBMI = "" Then
'    sCurrentValue = ""
'    sCurrentDate = ""
'    sAbnormal = ""
'Else
'
'    sCurrentValue = CStr(Round((Round(CSng(sWeight) * 0.453592, 2)) / ((Round(CSng(sHeight) * (91.44 / 36#), 2)) / 100) ^ 2, 2))
'End If


'-- Get Weight/Height Goal Value
'-----------------------------------------------
'sGoalValue = ""
'sAbnormal = ""
'sGoalOpType = ""

'-- Get Goal Value
'-----------------------------------------------

If Not GetPatientMatchingGoals(oRec, sUnitNumber, "3", cBMI_NCID) Then GoTo CleanUp

If oRec.RecordCount <> 0 Then
    sGoal = oRec.Fields("value")
    sGoalOpType = oRec.Fields("Goal_Op_Type")
End If


moRecBodyMass.AddNew
moRecBodyMass.Fields("Unit_Number") = sUnitNumber
moRecBodyMass.Fields("Goal_Value") = sGoalValue
moRecBodyMass.Fields("Current_Value") = sCurrentValue
moRecBodyMass.Fields("Current_Date") = sCurrentDate
moRecBodyMass.Fields("ABNORMAL") = sAbnormal
moRecBodyMass.Fields("OP_TYPE") = sGoalOpType
moRecBodyMass.Update

GetBodyMassIndex = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepCPGGoals.GetBodyMassIndex", "CHCSII_Reports", vbExclamation
CleanUp:
    Exit Function
End Function

Private Function GetHealthStatus(sUnitNumber As String) As Boolean

On Error GoTo ErrHandler

GetHealthStatus = False

Dim sStmt As String
Dim oRec As ADODB.Recordset
Dim sGoalValue As String
Dim sCurrentValue As String
Dim sCurrentDate As String

Const SequenceNumber As String = "1"

'-- Select
    sStmt = "Select " & vbNewLine & _
                "to_char(qr.start_date_time,'DD MON YYYY') as Start_Date, " & vbNewLine & _
                "ar.answer, " & vbNewLine & _
                "ad.adef_type, " & vbNewLine & _
                "ad.Params, " & _
                "'  ' as Final_Answer " & vbNewLine & _
                "ad.QDEF_ID, " & vbNewLine & _
                "ad.Version " & vbNewLine

'-- From
    sStmt = sStmt & _
            "From " & vbNewLine & _
                "adef ad, " & vbNewLine & _
                "CPGDEF cd, " & vbNewLine & _
                "aresp ar, " & vbNewLine & _
                "qresp qr, " & vbNewLine & _
                "mmi_generic_id fs, " & vbNewLine & _
                "mmi_generic_id pm, " & vbNewLine & _
                "provider p, " & vbNewLine & _
                "mmi_name mn, " & vbNewLine & _
                "cpgregistry cr, " & vbNewLine & _
                "mmi_id id " & vbNewLine

'-- Where
    sStmt = sStmt & _
        "Where " & vbNewLine & _
            "ad.seq = ar.seq and " & vbNewLine & _
            "ad.version =  qr.version and " & vbNewLine & _
            "ad.qdef_id = qr.qdef_id and " & vbNewLine & _
            "cd.name = '" & CPG_DIABETES_DEF_NAME & "' and " & vbNewLine & _
            "cd.cpg_id = cr.cpg_id and cr.unit_number = id.unit_number and " & vbNewLine & _
            "ar.seq = " & SequenceNumber & " and  " & vbNewLine & _
            "ar.qresp_id = qr.qresp_id and " & vbNewLine & _
            "qr.qresp_id in " & vbNewLine & _
            "(select max(qr2.qresp_id) " & vbNewLine & _
            "from qdef qd2, qresp qr2 " & vbNewLine & _
            "where " & vbNewLine & _
            "qd2.NAME in ('" & CPG_DIABETES_ANNUAL_Q_NAME & "','" & CPG_DIABETES_EACH_Q_NAME & "') and " & vbNewLine & _
            "qd2.qdef_id = qr2.qdef_id and " & vbNewLine & _
            "qr2.unit_number = id.unit_number) and " & vbNewLine & _
            "qr.unit_number = id.unit_number and " & vbNewLine & _
            "fs.id_type_ncid = 14511309 and " & vbNewLine & _
            "fs.unit_number = id.unit_number and " & vbNewLine & _
            "p.pcm_code = pm.id_value and " & vbNewLine & _
            "pm.id_type_ncid = 14501724 and " & vbNewLine & _
            "pm.unit_number = id.unit_number and " & vbNewLine & _
            "mn.name_type_cid = 187 and " & vbNewLine & _
            "mn.unit_number = id.unit_number and " & vbNewLine & _
            "id.unit_number = " & sUnitNumber

Set oRec = goDBDAS.OpenRecordset(sStmt)

'-- Get Text for Answer (shouldn't have to do this)
'-------------------------------------------------------------


If oRec.RecordCount > 0 Then
    If Not UpdateQuestionnaireFields(oRec) Then GoTo CleanUp
    oRec.MoveFirst
    sCurrentValue = oRec.Fields("Final_Answer")
    sCurrentDate = oRec.Fields("Start_Date")
Else
    sCurrentValue = ""
    sCurrentDate = ""
End If

moRecHealth.AddNew
moRecHealth.Fields("Unit_Number") = sUnitNumber
moRecHealth.Fields("Start_Date") = sCurrentDate
moRecHealth.Fields("Answer") = sCurrentValue
moRecHealth.Update


GetHealthStatus = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepCPGGoals.GetHealthStatus", "CHCSII_Reports", vbExclamation
CleanUp:
    Exit Function
End Function


Private Function GetBloodPressureSystolic(sUnitNumber As String) As Boolean

On Error GoTo ErrHandler
 
Dim sStmt As String
Dim sGoalValue As String
Dim sCurrentValue As String
Dim sCurrentDate As String
Dim sAbnormal As String
Dim sGoalOpType As String
Dim sGoal As String
Dim oRec As ADODB.Recordset

Dim sFacility As String
Dim sIEN As String


GetBloodPressureSystolic = False


'-- Get Blood Pressure Systolic
'-----------------------------------------------
sStmt = "Select " & _
            "TO_CHAR(e.unit_number) as Unit_Number, " & vbNewLine & _
            "To_Char(e.event_start_gmtime, 'dd mon yyyy') as Current_Date, " & vbNewLine & _
            "ov.value_number as Current_Value " & vbNewLine & _
        "From " & vbNewLine & _
            "observation_value ov, " & vbNewLine & _
            "clinical_observation co, " & vbNewLine & _
            "event e " & vbNewLine & _
        "Where " & vbNewLine & _
            "ov.observation_seq = co.observation_seq and " & vbNewLine & _
            "ov.event_id = co.event_id and " & vbNewLine & _
            "co.observation_ncid = " & cBP_SYSTOLIC & " and " & vbNewLine & _
            "co.event_id = e.event_id and " & vbNewLine & _
            "e.event_subtype_ncid =  110669 and " & vbNewLine & _
            "e.event_type_ncid = 90782 and " & vbNewLine & _
            "e.unit_number = " & sUnitNumber & " " & vbNewLine & _
        "order by e.event_start_gmtime desc"
            
Set oRec = goDBDAS.OpenRecordset(sStmt)

If oRec.RecordCount > 0 Then
    sCurrentValue = oRec.Fields("Current_Value")
    sCurrentDate = oRec.Fields("Current_Date")
    sAbnormal = "" 'oRec.Fields("Abnormal_Flag")
Else
    sCurrentValue = ""
    sCurrentDate = ""
    sAbnormal = ""
End If

'-- Get Goal Value
'-----------------------------------------------

If Not GetPatientMatchingGoals(oRec, sUnitNumber, "3", cBP_SYSTOLIC) Then GoTo CleanUp

If oRec.RecordCount <> 0 Then

    sGoalValue = oRec.Fields("value")
    sGoalOpType = oRec.Fields("Goal_Op_Type")

End If

moRecBloodSys.AddNew
moRecBloodSys.Fields("Unit_Number") = sUnitNumber
moRecBloodSys.Fields("Goal_Value") = sGoalValue
moRecBloodSys.Fields("Current_Value") = sCurrentValue
moRecBloodSys.Fields("Current_Date") = sCurrentDate
moRecBloodSys.Fields("ABNORMAL") = sAbnormal
moRecBloodSys.Fields("OP_TYPE") = sGoalOpType
moRecBloodSys.Update

GetBloodPressureSystolic = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepCPGGoals.GetBloodPressureSystolic", "CHCSII_Reports", vbExclamation
CleanUp:
    Exit Function
End Function

Private Function GetBloodPressureDiastolic(sUnitNumber As String) As Boolean

On Error GoTo ErrHandler
 
Dim sStmt As String
Dim sGoalValue As String
Dim sCurrentValue As String
Dim sCurrentDate As String
Dim sAbnormal As String
Dim sGoalOpType As String
Dim sGoal As String
Dim oRec As ADODB.Recordset

Dim sFacility As String
Dim sIEN As String


GetBloodPressureDiastolic = False


'-- Get Blood Pressure Systolic
'-----------------------------------------------
sStmt = "Select " & _
            "TO_CHAR(e.unit_number) as Unit_Number, " & vbNewLine & _
            "To_Char(e.event_start_gmtime, 'dd mon yyyy') as Current_Date, " & vbNewLine & _
            "ov.value_number as Current_Value " & vbNewLine & _
        "From " & vbNewLine & _
            "observation_value ov, " & vbNewLine & _
            "clinical_observation co, " & vbNewLine & _
            "event e " & vbNewLine & _
        "Where " & vbNewLine & _
            "ov.observation_seq = co.observation_seq and " & vbNewLine & _
            "ov.event_id = co.event_id and " & vbNewLine & _
            "co.observation_ncid = " & cBP_DIASTOLIC & " and " & vbNewLine & _
            "co.event_id = e.event_id and " & vbNewLine & _
            "e.event_subtype_ncid =  110669 and " & vbNewLine & _
            "e.event_type_ncid = 90782 and " & vbNewLine & _
            "e.unit_number = " & sUnitNumber & " " & vbNewLine & _
        "order by e.event_start_gmtime desc"
            
Set oRec = goDBDAS.OpenRecordset(sStmt)

If oRec.RecordCount > 0 Then
    sCurrentValue = oRec.Fields("Current_Value")
    sCurrentDate = oRec.Fields("Current_Date")
    sAbnormal = "" 'oRec.Fields("Abnormal_Flag")
Else
    sCurrentValue = ""
    sCurrentDate = ""
    sAbnormal = ""
End If

'-- Get Goal Value
'-----------------------------------------------

If Not GetPatientMatchingGoals(oRec, sUnitNumber, "3", cBP_DIASTOLIC) Then GoTo CleanUp

If oRec.RecordCount <> 0 Then

    sGoal = oRec.Fields("value")
    sGoalOpType = oRec.Fields("Goal_Op_Type")

End If

moRecBloodDia.AddNew
moRecBloodDia.Fields("Unit_Number") = sUnitNumber
moRecBloodDia.Fields("Goal_Value") = sGoalValue
moRecBloodDia.Fields("Current_Value") = sCurrentValue
moRecBloodDia.Fields("Current_Date") = sCurrentDate
moRecBloodDia.Fields("ABNORMAL") = sAbnormal
moRecBloodDia.Fields("OP_TYPE") = sGoalOpType
moRecBloodDia.Update

GetBloodPressureDiastolic = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepCPGGoals.GetBloodPressureSystolic", "CHCSII_Reports", vbExclamation
CleanUp:
    Exit Function
End Function

Private Function GetEyeExam(sUnitNumber As String) As Boolean

On Error GoTo ErrHandler

GetEyeExam = False

Dim sStmt As String
Dim oRec As ADODB.Recordset
Dim sGoalValue As String
Dim sCurrentValue As String
Dim sCurrentDate As String

Const SequenceNumber As String = "3"

'-- Select
    sStmt = "Select " & vbNewLine & _
                "to_char(qr.start_date_time,'DD MON YYYY') as Start_Date, " & vbNewLine & _
                "ar.answer, " & vbNewLine & _
                "ad.adef_type, " & vbNewLine & _
                "ad.Params, " & _
                "'  ' as Final_Answer " & vbNewLine & _
                "ad.QDEF_ID, " & vbNewLine & _
                "ad.Version " & vbNewLine

'-- From
    sStmt = sStmt & _
            "From " & vbNewLine & _
                "adef ad, " & vbNewLine & _
                "CPGDEF cd, " & vbNewLine & _
                "aresp ar, " & vbNewLine & _
                "qresp qr, " & vbNewLine & _
                "mmi_generic_id fs, " & vbNewLine & _
                "mmi_generic_id pm, " & vbNewLine & _
                "provider p, " & vbNewLine & _
                "mmi_name mn, " & vbNewLine & _
                "cpgregistry cr, " & vbNewLine & _
                "mmi_id id " & vbNewLine

'-- Where
    sStmt = sStmt & _
        "Where " & vbNewLine & _
            "ad.seq = ar.seq and " & vbNewLine & _
            "ad.version =  qr.version and " & vbNewLine & _
            "ad.qdef_id = qr.qdef_id and " & vbNewLine & _
            "cd.name = '" & CPG_DIABETES_DEF_NAME & "' and " & vbNewLine & _
            "cd.cpg_id = cr.cpg_id and cr.unit_number = id.unit_number and " & vbNewLine & _
            "ar.seq = " & SequenceNumber & " and  " & vbNewLine & _
            "ar.qresp_id = qr.qresp_id and " & vbNewLine & _
            "qr.qresp_id in " & vbNewLine & _
            "(select max(qr2.qresp_id) " & vbNewLine & _
            "from qdef qd2, qresp qr2 " & vbNewLine & _
            "where " & vbNewLine & _
            "qd2.NAME = '" & CPG_DIABETES_ANNUAL_Q_NAME & "' and " & vbNewLine & _
            "qd2.qdef_id = qr2.qdef_id and " & vbNewLine & _
            "qr2.unit_number = id.unit_number) and " & vbNewLine & _
            "qr.unit_number = id.unit_number and " & vbNewLine & _
            "fs.id_type_ncid = 14511309 and " & vbNewLine & _
            "fs.unit_number = id.unit_number and " & vbNewLine & _
            "p.pcm_code = pm.id_value and " & vbNewLine & _
            "pm.id_type_ncid = 14501724 and " & vbNewLine & _
            "pm.unit_number = id.unit_number and " & vbNewLine & _
            "mn.name_type_cid = 187 and " & vbNewLine & _
            "mn.unit_number = id.unit_number and " & vbNewLine & _
            "id.unit_number = " & sUnitNumber

Set oRec = goDBDAS.OpenRecordset(sStmt)

'-- Get Text for Answer (shouldn't have to do this)
'-------------------------------------------------------------


If oRec.RecordCount > 0 Then
    If Not UpdateQuestionnaireFields(oRec) Then GoTo CleanUp
    oRec.MoveFirst
    sCurrentValue = oRec.Fields("Final_Answer")
    sCurrentDate = oRec.Fields("Start_Date")
Else
    sCurrentValue = ""
    sCurrentDate = ""
End If

moRecEye.AddNew
moRecEye.Fields("Unit_Number") = sUnitNumber
moRecEye.Fields("Start_Date") = sCurrentDate
moRecEye.Fields("Answer") = sCurrentValue
moRecEye.Update


GetEyeExam = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepCPGGoals.GetEyeExam", "CHCSII_Reports", vbExclamation
CleanUp:
    Exit Function
End Function


Private Function GetFootExam(sUnitNumber As String) As Boolean

On Error GoTo ErrHandler

GetFootExam = False

Dim sStmt As String
Dim oRec As ADODB.Recordset
Dim sGoalValue As String
Dim sCurrentValue As String
Dim sCurrentDate As String

Const SequenceNumber As String = "2"

'-- Select
    sStmt = "Select " & vbNewLine & _
                "to_char(qr.start_date_time,'DD MON YYYY') as Start_Date, " & vbNewLine & _
                "ar.answer, " & vbNewLine & _
                "ad.adef_type, " & vbNewLine & _
                "ad.Params, " & _
                "'  ' as Final_Answer " & vbNewLine & _
                "ad.QDEF_ID, " & vbNewLine & _
                "ad.Version " & vbNewLine

'-- From
    sStmt = sStmt & _
            "From " & vbNewLine & _
                "adef ad, " & vbNewLine & _
                "CPGDEF cd, " & vbNewLine & _
                "aresp ar, " & vbNewLine & _
                "qresp qr, " & vbNewLine & _
                "mmi_generic_id fs, " & vbNewLine & _
                "mmi_generic_id pm, " & vbNewLine & _
                "provider p, " & vbNewLine & _
                "mmi_name mn, " & vbNewLine & _
                "cpgregistry cr, " & vbNewLine & _
                "mmi_id id " & vbNewLine

'-- Where
    sStmt = sStmt & _
        "Where " & vbNewLine & _
            "ad.seq = ar.seq and " & vbNewLine & _
            "ad.version =  qr.version and " & vbNewLine & _
            "ad.qdef_id = qr.qdef_id and " & vbNewLine & _
            "cd.name = '" & CPG_DIABETES_DEF_NAME & "' and " & vbNewLine & _
            "cd.cpg_id = cr.cpg_id and cr.unit_number = id.unit_number and " & vbNewLine & _
            "ar.seq = " & SequenceNumber & " and  " & vbNewLine & _
            "ar.qresp_id = qr.qresp_id and " & vbNewLine & _
            "qr.qresp_id in " & vbNewLine & _
            "(select max(qr2.qresp_id) " & vbNewLine & _
            "from qdef qd2, qresp qr2 " & vbNewLine & _
            "where " & vbNewLine & _
            "qd2.NAME = '" & CPG_DIABETES_ANNUAL_Q_NAME & "' and " & vbNewLine & _
            "qd2.qdef_id = qr2.qdef_id and " & vbNewLine & _
            "qr2.unit_number = id.unit_number) and " & vbNewLine & _
            "qr.unit_number = id.unit_number and " & vbNewLine & _
            "fs.id_type_ncid = 14511309 and " & vbNewLine & _
            "fs.unit_number = id.unit_number and " & vbNewLine & _
            "p.pcm_code = pm.id_value and " & vbNewLine & _
            "pm.id_type_ncid = 14501724 and " & vbNewLine & _
            "pm.unit_number = id.unit_number and " & vbNewLine & _
            "mn.name_type_cid = 187 and " & vbNewLine & _
            "mn.unit_number = id.unit_number and " & vbNewLine & _
            "id.unit_number = " & sUnitNumber

Set oRec = goDBDAS.OpenRecordset(sStmt)

'-- Get Text for Answer (shouldn't have to do this)
'-------------------------------------------------------------


If oRec.RecordCount > 0 Then
    If Not UpdateQuestionnaireFields(oRec) Then GoTo CleanUp
    oRec.MoveFirst
    sCurrentValue = oRec.Fields("Final_Answer")
    sCurrentDate = oRec.Fields("Start_Date")
Else
    sCurrentValue = ""
    sCurrentDate = ""
End If

moRecFoot.AddNew
moRecFoot.Fields("Unit_Number") = sUnitNumber
moRecFoot.Fields("Start_Date") = sCurrentDate
moRecFoot.Fields("Answer") = sCurrentValue
moRecFoot.Update


GetFootExam = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepCPGGoals.GetFootExam", "CHCSII_Reports", vbExclamation
CleanUp:
    Exit Function
End Function
Private Function GetPatients(oPatientsColl As Collection, _
                             oReportItem As ReportItem) As Boolean

On Error GoTo ErrHandler
 
GetPatients = False

'Dim sPCM As String
'Dim oPCMCollection As Collection


Set oPatientsColl = Nothing
Set oPatientsColl = New Collection

If oReportItem.Criteria1 = Patient Then

    oPatientsColl.Add oReportItem.CriteriaSel1.Item(1).UnitNumber
    
ElseIf oReportItem.Criteria1 = PCM Then

    If Not GetUnitNumbersForPCMs(oReportItem.CriteriaSel1, oPatientsColl) Then GoTo CleanUp
    
ElseIf oReportItem.Criteria1 = PCMMulti Then
    
    If Not GetUnitNumbersForPCMs(oReportItem.CriteriaSel1, oPatientsColl) Then GoTo CleanUp
    
End If

GetPatients = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepCPGGoals.GetPatients", "CHCSII_Reports", vbExclamation
CleanUp:
   'Set oCrystal = Nothing
    'Set oRec = Nothing
End Function

Private Function GetDemographics(sUnitNumber As String) As Boolean

On Error GoTo ErrHandler
 
Dim sStmt As String
Dim i As Integer
Dim i2 As Integer
Dim oRec As ADODB.Recordset


GetDemographics = False

sStmt = "Select " & _
            "TO_CHAR(id.unit_number) as Unit_Number, " & _
            "Decode(p.name,Null,'TEST',p.name) as PCM, " & vbNewLine & _
            "fmp.id_value as FMP_SSN, " & vbNewLine & _
            "Trim(mn.last_name || ', ' || mn.first_name || ' ' || mn.middle_name) as full_name " & vbNewLine & _
        "From " & vbNewLine & _
            "provider p, " & vbNewLine & _
            "mmi_generic_id pcm, " & vbNewLine & _
            "mmi_generic_id fmp, " & vbNewLine & _
            "mmi_name mn, " & vbNewLine & _
            "mmi_id id " & vbNewLine & _
        "Where " & vbNewLine & _
            "p.pcm_code (+) = pcm.id_value and " & vbNewLine & _
            "pcm.id_type_ncid (+) = 14501724 and " & vbNewLine & _
            "pcm.unit_number (+) = id.unit_number and " & vbNewLine & _
            "fmp.id_type_ncid = 14511309 and " & vbNewLine & _
            "fmp.unit_number = id.unit_number and " & vbNewLine & _
            "mn.name_type_cid = 187 and  " & vbNewLine & _
            "mn.unit_number = id.unit_number and " & vbNewLine & _
            "id.unit_number = " & sUnitNumber
            
Set oRec = goDBDAS.OpenRecordset(sStmt)

For i = 1 To oRec.RecordCount
    moRecDemographics.AddNew
    For i2 = 0 To moRecDemographics.Fields.Count - 1
        moRecDemographics.Fields(i2) = oRec.Fields(i2).Value
    Next
    moRecDemographics.Update
Next


GetDemographics = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepCPGGoals.GetDemographics", "CHCSII_Reports", vbExclamation
CleanUp:
    Set oRec = Nothing
    Exit Function
End Function


Private Function Initialize_RecordSets() As Boolean

On Error GoTo ErrHandler
 
Initialize_RecordSets = False

'-- Demographics
'------------------------------------------------
Set moRecDemographics = New ADODB.Recordset
moRecDemographics.CursorLocation = adUseClient

moRecDemographics.Fields.Append "Unit_Number", adVarChar, 10
moRecDemographics.Fields.Append "PCM", adVarChar, 30
moRecDemographics.Fields.Append "FMP_SSN", adVarChar, 20
moRecDemographics.Fields.Append "Full_NAME", adVarChar, 61
moRecDemographics.Open

'-- HGB A1C
'------------------------------------------------
Set moRecHGB = New ADODB.Recordset
moRecHGB.CursorLocation = adUseClient

moRecHGB.Fields.Append "Unit_Number", adVarChar, 10
moRecHGB.Fields.Append "Goal_Value", adVarChar, 30
moRecHGB.Fields.Append "Current_Value", adVarChar, 20
moRecHGB.Fields.Append "Current_Date", adVarChar, 61
moRecHGB.Fields.Append "ABNORMAL", adVarChar, 1
moRecHGB.Fields.Append "OP_TYPE", adVarChar, 3
moRecHGB.Open

'-- Microalbumin
'------------------------------------------------
Set moRecMicro = New ADODB.Recordset
moRecMicro.CursorLocation = adUseClient

moRecMicro.Fields.Append "Unit_Number", adVarChar, 10
moRecMicro.Fields.Append "Goal_Value", adVarChar, 30
moRecMicro.Fields.Append "Current_Value", adVarChar, 20
moRecMicro.Fields.Append "Current_Date", adVarChar, 61
moRecMicro.Fields.Append "ABNORMAL", adVarChar, 1
moRecMicro.Fields.Append "OP_TYPE", adVarChar, 3
moRecMicro.Open

'-- Lipid Profile
'------------------------------------------------
Set moRecLipid = New ADODB.Recordset
moRecLipid.CursorLocation = adUseClient

moRecLipid.Fields.Append "Unit_Number", adVarChar, 10
moRecLipid.Fields.Append "Goal_Value", adVarChar, 30
moRecLipid.Fields.Append "Current_Value", adVarChar, 20
moRecLipid.Fields.Append "Current_Date", adVarChar, 61
moRecLipid.Fields.Append "ABNORMAL", adVarChar, 1
moRecLipid.Fields.Append "OP_TYPE", adVarChar, 3
moRecLipid.Open


'-- Body Mass Index (BMI)
'------------------------------------------------
Set moRecBodyMass = New ADODB.Recordset
moRecBodyMass.CursorLocation = adUseClient

moRecBodyMass.Fields.Append "Unit_Number", adVarChar, 10
moRecBodyMass.Fields.Append "Goal_Value", adVarChar, 30
moRecBodyMass.Fields.Append "Current_Value", adVarChar, 20
moRecBodyMass.Fields.Append "Current_Date", adVarChar, 61
moRecBodyMass.Fields.Append "ABNORMAL", adVarChar, 1
moRecBodyMass.Fields.Append "OP_TYPE", adVarChar, 3
moRecBodyMass.Open


'-- Blood Pressure Systolic
'------------------------------------------------
Set moRecBloodSys = New ADODB.Recordset
moRecBloodSys.CursorLocation = adUseClient

moRecBloodSys.Fields.Append "Unit_Number", adVarChar, 10
moRecBloodSys.Fields.Append "Goal_Value", adVarChar, 30
moRecBloodSys.Fields.Append "Current_Value", adVarChar, 20
moRecBloodSys.Fields.Append "Current_Date", adVarChar, 61
moRecBloodSys.Fields.Append "ABNORMAL", adVarChar, 1
moRecBloodSys.Fields.Append "OP_TYPE", adVarChar, 3
moRecBloodSys.Open


'-- Blood Pressure Diastolic
'------------------------------------------------
Set moRecBloodDia = New ADODB.Recordset
moRecBloodDia.CursorLocation = adUseClient

moRecBloodDia.Fields.Append "Unit_Number", adVarChar, 10
moRecBloodDia.Fields.Append "Goal_Value", adVarChar, 30
moRecBloodDia.Fields.Append "Current_Value", adVarChar, 20
moRecBloodDia.Fields.Append "Current_Date", adVarChar, 61
moRecBloodDia.Fields.Append "ABNORMAL", adVarChar, 1
moRecBloodDia.Fields.Append "OP_TYPE", adVarChar, 3
moRecBloodDia.Open



'-- Eye EXAMS
'------------------------------------------------
Set moRecEye = New ADODB.Recordset
moRecEye.CursorLocation = adUseClient

moRecEye.Fields.Append "Unit_Number", adVarChar, 10
moRecEye.Fields.Append "Start_Date", adVarChar, 15
moRecEye.Fields.Append "Answer", adVarChar, 50
moRecEye.Open


'-- Health Status
'------------------------------------------------
Set moRecHealth = New ADODB.Recordset
moRecHealth.CursorLocation = adUseClient

moRecHealth.Fields.Append "Unit_Number", adVarChar, 10
moRecHealth.Fields.Append "Start_Date", adVarChar, 15
moRecHealth.Fields.Append "Answer", adVarChar, 50
moRecHealth.Open


'-- Foot Exam
'------------------------------------------------
Set moRecFoot = New ADODB.Recordset
moRecFoot.CursorLocation = adUseClient

moRecFoot.Fields.Append "Unit_Number", adVarChar, 10
moRecFoot.Fields.Append "Start_Date", adVarChar, 15
moRecFoot.Fields.Append "Answer", adVarChar, 50
moRecFoot.Open


Initialize_RecordSets = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepCPGGoals.Initialize_RecordSets", "CHCSII_Reports", vbExclamation
CleanUp:
    Exit Function
End Function

Private Function AddRecords(oRecSource As ADODB.Recordset, _
                            oRecDest As ADODB.Recordset) As Boolean

On Error GoTo ErrHandler
Dim i As Integer
Dim i2 As Integer

AddRecords = False

For i = 1 To oRecSource.RecordCount
    oRecDest.AddNew
    For i2 = 0 To oRecDest.Fields.Count - 1
        oRecDest.Fields(i2) = oRecSource.Fields(i2).Value
    Next
    oRecDest.Update
Next

AddRecords = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepCPGGoals.AddRecords", "CHCSII_Reports", vbExclamation
CleanUp:
    Exit Function
End Function


Private Function LabResultSearch(sSearchText As String, oRec As ADODB.Recordset) As Boolean
    
LabResultSearch = False

Dim sStmt As String
Dim oRec2 As ADODB.Recordset
Dim icount As Integer

'-- Query
'---------------------------------------
sStmt = "select distinct" & vbNewLine & _
            "r.ncid," & vbNewLine & _
            "r.representation as Lab_Result" & vbNewLine & _
        "From" & vbNewLine & _
            "rsform_context rc," & vbNewLine & _
            "rsform_context rc2," & vbNewLine & _
            "rsform r2," & vbNewLine & _
            "concept_relation cr," & vbNewLine & _
            "rsform r" & vbNewLine & _
        "Where" & vbNewLine & _
            "rc.preferred_score = 0 and" & vbNewLine & _
            "rc.context_ncid = 21916 and" & vbNewLine & _
            "rc.rsform_id = r.rsform_id and" & vbNewLine & _
            "rc2.context_ncid = 54347 and" & vbNewLine & _
            "rc2.rsform_id = r2.rsform_id and" & vbNewLine & _
            "r2.ncid = cr.concept_ncid and" & vbNewLine & _
            "cr.relationship_ncid = 363 and" & vbNewLine & _
            "cr.concept_relation_ncid = 552 and" & vbNewLine & _
            "cr.concept_ncid = r.ncid and" & vbNewLine & _
            "UPPER(r.representation) like '%" & sSearchText & "%'" & vbNewLine & _
        "Order by 2"

Set oRec = goDBDAS.OpenRecordset(sStmt)


'-- Remove Duplicates
'---------------------------------------
If oRec.RecordCount > 1 Then

    Set oRec2 = oRec.Clone
    oRec2.MoveFirst
    oRec.MoveFirst

    Do
    
        oRec2.MoveNext

        If oRec.Fields("Lab_Result") = oRec2.Fields("Lab_Result") Then
            If GetCount(oRec.Fields("NCID")) >= GetCount(oRec2.Fields("NCID")) Then
                oRec.MoveNext
                oRec.Delete
                oRec2.MoveNext
            Else
                oRec.Delete
            End If
        End If

        oRec.MoveNext
        

    Loop While Not oRec2.EOF

End If

LabResultSearch = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "LabResultSearch ", "", vbExclamation
CleanUp:
    Set oRec2 = Nothing
    Exit Function
End Function

Private Sub Class_Initialize()
    '--- SCR-ARM-16148 MSolano 9/10/01; COMMON TO MULTIPLE ARM-SCRs
    '--- Init routine should be called from first top level creatable coclass
    '--- It can be called multiple times, but will actually run only once.
    '--- Instrument all creatable classes in case dll's multiple entry
    '--- points are actually used
     Call ARM_InitAllTransactionsForCHCSIIReports

    msFacilityNCID = cwobjLogonEnviron.FacilityNCID
End Sub


