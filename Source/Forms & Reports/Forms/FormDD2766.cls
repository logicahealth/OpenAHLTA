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
Attribute VB_Name = "FormDD2766"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const cMODULE_NAME As String = "FormDD2766.cls"
Private Const clFieldAttr As Long = ADODB.FieldAttributeEnum.adFldLong + _
                    ADODB.FieldAttributeEnum.adFldUpdatable + _
                    ADODB.FieldAttributeEnum.adFldIsNullable
Private Const cNO_ALLERGIES As String = "No Allergies Specified"
Private Const cNO_KNOWN_ALLERGIES As String = "No Known Allergies"
Private mMedicalRecordNumber As String
Private moPatient As Patient

Private moCollRecSet As Collection
Private moCollFormulas As Collection

Private moPatientDemographics As PatientDemographics

'*****************************************************************************************
'Description: Convert Hex string to ASICC Character
'Parameter: psHex: Hex string
'Reutrn:    String
'History:
'   Date    Developer       Description
'========== ============    ==========================================================
'08/22/01   Wilson Sun      Created
'*****************************************************************************************
Private Function HexToStr(ByVal psHex As String) As String

On Error GoTo ErrHandler

    Dim i As Integer
    Dim j As Integer
    Dim x() As Byte
        
    ReDim x(Len(psHex))
    j = 0
    For i = 1 To Len(psHex) Step 2
        x(j) = FromHexToDecimal(Mid(psHex, i, 2))
        j = j + 1
    Next i
    
    HexToStr = StrConv(x, vbUnicode)
    
    GoTo Cleanup
ErrHandler:
    Err.Raise Err.Number, App.EXEName & "." & cMODULE_NAME & "." & "HexToStr:" & Err.Source, _
       Err.Description, Err.HelpFile, Err.HelpContext
Cleanup:
    
End Function

Public Function FromHexToDecimal(ByVal vsHexVal As String) As String
On Error GoTo ErrHandler
    Dim lVal As Long
    Dim lMultVal As Long
    Dim i As Long
    Dim sChar As String
    
    lVal = 0
    For i = 1 To Len(vsHexVal)
        sChar = Mid$(vsHexVal, Len(vsHexVal) - i + 1, 1)
        If (i > 1) Then
            lMultVal = 16 ^ (i - 1)
        Else
            lMultVal = 1
        End If
        Select Case LCase$((sChar))
            Case "1" To "9": lVal = lVal + (CLng(sChar) * lMultVal)
            Case "a": lVal = lVal + (10 * lMultVal)
            Case "b": lVal = lVal + (11 * lMultVal)
            Case "c": lVal = lVal + (12 * lMultVal)
            Case "d": lVal = lVal + (13 * lMultVal)
            Case "e": lVal = lVal + (14 * lMultVal)
            Case "f": lVal = lVal + (15 * lMultVal)
            Case Else:
        End Select
    Next i
    FromHexToDecimal = CStr(lVal)
    GoTo Cleanup
ErrHandler:
    Err.Raise Err.Number, App.EXEName & "." & cMODULE_NAME & "." & "FromHexToDecimal:" & Err.Source, _
       Err.Description, Err.HelpFile, Err.HelpContext
Cleanup:
    
End Function

'*****************************************************************************************
'Description: Get Counseleing data from OBSERVATION_VALUE and CLINICAL_OBSERVATION table
'             Clone from GetCounseling(). SCR 16475
'Parameter: strUnitNumber,   String,
'           objRecordset,    ADODB.Recordset,
'Reutrn:    Strin
'History:
'   Date    Developer       Description
'========== ============    ==========================================================
'11/13/01   Wilson Sun      SCR: 16475. Coded based on attached file in SCR 16475
'                           "Mapping to Wellness Reminders to Counseling DD2766 Ver 2_Release 1 Plus.xls"
'09/24/01   Wilson Sun      Created
'*****************************************************************************************
Private Function GetCounseling(ByRef objRecordset As ADODB.Recordset) As Boolean

On Error GoTo ErrHandler
    
    Dim intIndex As Integer
    Dim strSQL As String
    Dim strTopic As String
    Dim objRecordsetFrom As ADODB.Recordset
    
    GetCounseling = False
    
    '-- Create Template Recordset
    If Not CreateCounselingTemplate(objRecordset) Then
        GetCounseling = False
        GoTo ErrHandler
    End If
    
    '-- Copy first 20 results to new Recordset
    objRecordset.Open
    objRecordset.AddNew
    
    If gobjshared.IsAppMode(modeTheater) Then 'GetCounseling is not support in Theater mode
        GetCounseling = True
        GoTo Cleanup
    End If
    
    strSQL = strSQL & " e.event_start_gmtime as Event_Date" & vbNewLine
      '-- from
    strSQL = strSQL & _
        "From Event e," & vbNewLine & _
        "clinical_observation co," & vbNewLine & _
        "observation_value ov " & vbNewLine
    '-- where
    strSQL = strSQL & _
        "Where ov.observation_seq = co.observation_seq" & vbNewLine & _
        "and ov.event_id = co.event_id" & vbNewLine & _
        "and co.observation_ncid = 162045" & vbNewLine & _
        "and ov.value_ncid in (" & vbNewLine & _
        "207659,207660,207661,207662,207663,207664,207665,207666,207667,207668," & vbNewLine & _
        "207672,207673,207674,207675,207676,207677,207678,207687,207692,207693," & vbNewLine & _
        "207694,207695,207700,207701,207702,207703,207704,207705,207706)" & vbNewLine & _
        "and co.event_id = e.event_id" & vbNewLine & _
        "and e.event_subtype_ncid = 162175" & vbNewLine & _
        "and e.event_type_ncid=90782" & vbNewLine & _
        "and e.unit_number=" & moPatient.UnitNumber & vbNewLine & _
        "order by e.event_start_gmtime desc"
    
    Set objRecordsetFrom = goDBDAS.OpenRecordset(strSQL, adOpenDynamic, adLockOptimistic)
    
    If Not (objRecordsetFrom.BOF And objRecordsetFrom.EOF) Then
        objRecordsetFrom.MoveFirst
        intIndex = 1
        objRecordset.Fields("TOPIC_DT1") = objRecordsetFrom.Fields("TOPIC_DT")
        objRecordset.Fields("AGE" & CStr(intIndex)) = GetAge(objRecordsetFrom.Fields("EVENT_DATE"))
        Do While Not objRecordsetFrom.EOF And intIndex < 20
            objRecordset.Fields("TOPIC_DT" & CStr(intIndex)) = objRecordsetFrom.Fields("TOPIC_DT")
            objRecordset.Fields("AGE" & CStr(intIndex)) = GetAge(objRecordsetFrom.Fields("EVENT_DATE"))
            objRecordset.Fields("TOPIC" & CStr(intIndex)) = objRecordsetFrom.Fields("TOPIC") & ","
            intIndex = intIndex + 1
            objRecordset.Fields("TOPIC" & CStr(intIndex)) = StrConv(objRecordset.Fields("TOPIC" & CStr(intIndex)), vbProperCase)
            objRecordsetFrom.MoveNext
        Loop
    End If
    
    '--remove last camma(,) in TOPIC field
    If Not (objRecordset.BOF And objRecordset.EOF) Then
        objRecordset.MoveFirst
        For intIndex = 1 To 20
            If IsNull(objRecordset.Fields("TOPIC" & CStr(intIndex))) Then
               strTopic = NULLStr
            Else
               strTopic = objRecordset.Fields("TOPIC" & CStr(intIndex))
            End If
            If strTopic <> NULLStr Then
                objRecordset.Fields("TOPIC" & CStr(intIndex)) = VBA.left$(strTopic, VBA.Len(strTopic) - 1)
            End If
        Next
    End If
    
    GetCounseling = True
    GoTo Cleanup
    
ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, cMODULE_NAME & ".GetCounseling", App.EXEName, vbExclamation
Cleanup:
    Set objRecordsetFrom = Nothing

End Function

Private Function GetPatientData(ByRef objRecordset As ADODB.Recordset) As Boolean

    Dim sFirst_name As String
    Dim sMiddle_name As String
    Dim sLast_name As String
    Dim sValue As String
    Dim lCount As Long
    
    On Error GoTo Proc_Exit
        
    GetPatientData = False

    '-- Create Template Recordset
    If Not CreatePatientTemplate(objRecordset) Then GoTo Proc_Exit
    objRecordset.Open
    With objRecordset
         .AddNew
         .Fields("job_id") = "1"
        .Fields("UNIT_NUMBER") = moPatient.UnitNumber
        .Fields("SSN") = moPatient.SSN
        .Fields("Sex") = moPatient.Sex
        .Fields("DOB") = moPatient.DateOfBirth
        .Fields("DOB_FORMATTED") = Format(moPatient.DateOfBirth, gsDefaultFormat)
        .Fields("Age") = gobjshared.CalcAge(CDbl(moPatient.DateOfBirth))
        .Fields("MARITAL_STATUS") = moPatient.MaritalStatus
        
        If moPatient.CommandInterest.Count > 0 Then .Fields("COMMAND_INTEREST") = moPatient.CommandInterest(1)
        If moPatient.CommandSecurity.Count > 0 Then .Fields("COMMAND_SECURITY") = moPatient.CommandSecurity(1)
        
        .Fields("ALTERNATIVE_CARE") = moPatient.ACV_EnrollmentStatus
        
        .Fields("FLYING_STATUS") = "0"
        .Fields("JUMPING_STATUS") = "0"
        .Fields("DIVING_STATUS") = "0"
        .Fields("ON_MOBILITY_STATUS") = "0"
        .Fields("SUBMARINE_STATUS") = "0"
        If "Y" = moPatient.GetProperty("217201", False) Then .Fields("FLYING_STATUS") = "1"
        If "Y" = moPatient.GetProperty("217202", False) Then .Fields("JUMPING_STATUS") = "1"
        If "Y" = moPatient.GetProperty("217203", False) Then .Fields("DIVING_STATUS") = "1"
        If "Y" = moPatient.GetProperty("217204", False) Then .Fields("ON_MOBILITY_STATUS") = "1"
        If "Y" = moPatient.GetProperty("217205", False) Then .Fields("SUBMARINE_STATUS") = "1"
        
        .Fields("PATIENT_CATEGORY") = moPatient.PatientCategory
        .Fields("UIC") = moPatient.UnitIdentificationCode
        .Fields("SPONSOR_SSN") = moPatient.SponsorSSN
        .Fields("FMP") = moPatient.FMP
        .Fields("VIP") = moPatient.VIP
        .Fields("PCM_ID") = moPatient.PrimaryCareManager
        .Fields("HOME_PHONE") = moPatient.HomePhone
        .Fields("WORK_PHONE") = moPatient.WorkPhone
        '--- SCR 134102 Sherry Wang   7/3/2007
        '--- Service is stored to mmi_genericdata table as a seperate record.
        .Fields("SERVICE") = moPatient.BranchOfService
        .Fields("Rank") = moPatient.Rank
        .Fields("RECORDS_MAINTAINED_AT") = moPatient.RecordsMaintainedAt
        .Fields("RELATIONSHIP_TO_SPONSOR") = GetRelationshipToSponsor(moPatient.FMP)
        
        GetPatientName sFirst_name, sMiddle_name, sLast_name
        .Fields("FIRST_NAME") = sFirst_name
        .Fields("MIDDLE_NAME") = sMiddle_name
        .Fields("LAST_NAME") = sLast_name
        
        .Fields("FULL_NAME") = moPatient.Name
        .Fields("Religion") = moPatient.Religion
        .Fields("Race") = moPatient.Race
        .Fields("ACTIVE_DUTY") = IsActiveDuty(moPatient.UnitNumber)
    End With
    
    GetPatientData = True

Proc_Exit:
    
    If Err.Number <> 0 Then
        'Resume Next
        GetPatientData = False
        gobjshared.ShowVBError Err.Number, Err.Description, cMODULE_NAME & ".GetPatientData", App.EXEName, vbExclamation
    End If
    
End Function

Private Function GetSponsorPatientData(ByRef objRecordset As ADODB.Recordset) As Boolean

    Dim lCount As Long
    
On Error GoTo Proc_Exit
        
    GetSponsorPatientData = False
    
    '-- Create Template Recordset
    If Not CreatePatientTemplate(objRecordset) Then GoTo Proc_Exit
    objRecordset.Open

    If (Not moPatient.PatientOps.Sponsor Is Nothing) Then
        With moPatient.PatientOps.Sponsor
            objRecordset.AddNew
            objRecordset.Fields("job_id") = "1"
            objRecordset.Fields("UNIT_NUMBER") = .UnitNumber
            objRecordset.Fields("FULL_NAME") = .FullName
            objRecordset.Fields("ACTIVE_DUTY") = IsActiveDuty(.UnitNumber)
        End With
    End If
    
    GetSponsorPatientData = True
    
    If Err.Number <> 0 Then
        'Resume Next
        GetSponsorPatientData = False
        gobjshared.ShowVBError Err.Number, Err.Description, cMODULE_NAME & ".GetSponsorPatientData", App.EXEName, vbExclamation
    End If
    Exit Function
    
Proc_Exit:

    Err.Clear
    Resume
    
End Function

Private Function IsActiveDuty(sUnitNumber As String) As String

    Dim sStmt As String
    
    On Error GoTo ErrHandler
    IsActiveDuty = "N"

    sStmt = "select Trim(b.value_text) UIC " & _
            "from gems.mmi_genericdata b " & _
            "where (b.value_ncid in " & _
                "(select a.ncid from gems.ncids a " & _
                " Where a.domainncid = 14501721 " & _
                " and substr(a.description,2,2) < '15' " & _
                " and substr(a.description,2,2) > '00') " & _
                " and unit_number = " & sUnitNumber & ") "
    
    goDBSqlOps.Execute sStmt
    
    If Not (goDBSqlOps.BOF And goDBSqlOps.EOF) Then
        IsActiveDuty = "Y"
    End If

    Exit Function

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, cMODULE_NAME & ".IsActiveDuty", App.EXEName, vbExclamation
    
End Function

'*****************************************************************************************
'Name: GetChronicIllness
'Description: Get Chronic data from Problem list
'Parameter: strUnitNumber,   String,
'           objRecordset,    ADODB.Recordset,
'           objRecordsetSub, ADODB.Recordset
'Reutrn:    String
'History:
'   Date    Developer       Description
'========== ============    ==========================================================
'10/16/01   Wilson Sun      Use Problem module to get data instead of using SQL. SCR (11492)
'??/??/??   Brain Webb      Created
'*****************************************************************************************
Private Function GetChronicIllness(ByRef oProblemOps As Object, ByRef objRecordset As ADODB.Recordset) As Boolean

    On Error GoTo Proc_Exit
    Dim objOneProblem As Object 'OneProblem

    GetChronicIllness = False

    'Create record set to hold data
    Set objRecordset = New ADODB.Recordset
    With objRecordset
        .CursorLocation = adUseClient
        .Fields.Append "Chronic", adBSTR
        .Fields.Append "ChronicDT", adBSTR
        .Fields.Append "ChronicDate", adDate
        .Open
    End With
    
    For Each objOneProblem In oProblemOps.Problems
        With objOneProblem
            If .Status = "Active" And .Chronicity = "Chronic" Then
                objRecordset.AddNew
                objRecordset.Fields!Chronic = StrConv(.Problem, vbProperCase)
                objRecordset.Fields!ChronicDT = UCase(Format(.OnsetDate, "dd MMM yyyy"))
                If IsDate(.OnsetDateTime) Then
                    objRecordset.Fields!ChronicDate = CDate(.OnsetDateTime)
                End If
            End If
        End With
    Next
    
    objRecordset.Sort = "ChronicDate Desc"
    GetChronicIllness = True
    
Proc_Exit:
    Set objOneProblem = Nothing
    
    If Err.Number <> 0 Then
        GetChronicIllness = False
        gobjshared.ShowVBError Err.Number, Err.Description, cMODULE_NAME & ".GetChronicIllness", App.EXEName, vbExclamation
    End If

End Function


'*****************************************************************************************
'Description: Get Historical Procecedures data from Problem module.
'             Clone from GetHospitalizations().
'Parameter: strUnitNumber,   String,
'           objRecordset,    ADODB.Recordset,
'           objRecordsetSub, ADODB.Recordset
'Reutrn:    String
'History:
'   Date    Developer       Description
'========== ============    ==========================================================
'11/30/01   Wilson Sun      SCR 19214. Added extra field (HospitalDate) to recordset for sorting.
'09/12/01   Wilson Sun      Created
'*****************************************************************************************
Private Function GetHospitalizations(ByRef oProblemOps As Object, _
                                   ByRef objRecordset As ADODB.Recordset, _
                                   ByRef objRecordsetSub As ADODB.Recordset) As Boolean

    On Error GoTo Proc_Exit
    
    Dim sCPT As String
    Dim objRecordsetFrom As ADODB.Recordset
    Dim objProcedure As Object 'OneProcedure '

    GetHospitalizations = False

    '-- Create Template Recordset
    If Not CreateHospitalTemplate(objRecordset) Then
        GoTo Proc_Exit
    End If
            
    'Create record set to hold data
    Set objRecordsetSub = New ADODB.Recordset
    With objRecordsetSub
        .CursorLocation = adUseClient
        .Fields.Append "HOSPITAL", adBSTR
        .Fields.Append "HOSPITALDT", adBSTR
        .Fields.Append "HospitalDate", adDate
        .Open
    End With
    
    '--- Inpatient Admissions
    Dim sqlStatement  As String
    
    sqlStatement = "select B.DESCRIPTION AS HOSPITAL_NAME, A.ADMISSION_DATE - A.TIMEBIAS/60/24 AS HOSPITAL_DATE_TIME " & vbCrLf & _
            "from INPATIENT_ADMISSION A, NCIDS B " & vbCrLf & _
            "Where A.UNIT_NUMBER = " & moPatient.UnitNumber & vbCrLf & _
            " AND A.FACILITY_NCID = B.NCID"
    goDBSqlOps.Execute sqlStatement
    
    Do While Not goDBSqlOps.EOF
        objRecordsetSub.AddNew
        objRecordsetSub.Fields!HOSPITAL = goDBSqlOps.Value("HOSPITAL_NAME")
        objRecordsetSub.Fields!HospitalDate = CDate(goDBSqlOps.Value("HOSPITAL_DATE_TIME"))
        objRecordsetSub.Fields!HOSPITALDT = UCase(Format(goDBSqlOps.Value("HOSPITAL_DATE_TIME"), "dd MMM yyyy"))
        
        goDBSqlOps.MoveNext
    Loop
        
    '--- Procedures
    For Each objProcedure In oProblemOps.Procedures
        With objProcedure
            sCPT = .ProcedureCode
            If .ProblemListIndicator = "Historical Procedure" Or _
               .ProblemListIndicator = "Encounter Procedure" Then
                'Exclude CPT codes (SCR 16747)
                If Not (sCPT = "10060" Or sCPT = "10120" Or sCPT = "11055" Or sCPT = "11056" Or sCPT = "11200" Or sCPT = "11719" Or sCPT = "11750" Or sCPT = "11765" Or (sCPT >= "11975" And sCPT <= "11977") Or (sCPT >= "12001" And sCPT <= "12002") Or sCPT = "16000" Or _
                    sCPT = "21310" Or sCPT = "21315" Or sCPT = "24650" Or sCPT = "24670" Or sCPT = "25500" Or sCPT = "25530" Or sCPT = "26600" Or sCPT = "26641" Or sCPT = "26700" Or sCPT = "26770" Or sCPT = "27750" Or sCPT = "27780" Or sCPT = "28470" Or sCPT = "28510" Or _
                    sCPT = "30110" Or sCPT = "30901" Or sCPT = "36000" Or sCPT = "45330" Or sCPT = "54150" Or sCPT = "56420" Or sCPT = "57452" Or sCPT = "59000" Or sCPT = "59020" Or sCPT = "59025" Or sCPT = "59030" Or sCPT = "59050" Or sCPT = "59051" Or sCPT = "69090" Or _
                    (sCPT >= "70010" And sCPT <= "76499") Or (sCPT >= "76505" And sCPT <= "76999") Or (sCPT >= "76805" And sCPT <= "76816") Or sCPT = "76818" Or sCPT = "76819" Or sCPT = "76825" Or (sCPT >= "78300" And sCPT <= "78366") Or (sCPT >= "80048" And sCPT <= "80090") Or _
                    (sCPT >= "80100" And sCPT <= "80440") Or (sCPT >= "81000" And sCPT <= "88167") Or (sCPT >= "90476" And sCPT <= "90477") Or (sCPT >= "99201" And sCPT <= "99205") Or (sCPT >= "99211" And sCPT <= "99215") Or (sCPT >= "99381" And sCPT <= "99397")) Then
                    objRecordsetSub.AddNew
                    objRecordsetSub.Fields!HOSPITAL = StrConv(.Procedure, vbProperCase)
                    objRecordsetSub.Fields!HOSPITALDT = UCase(Format(.OnsetDate, "dd MMM yyyy"))
                    If IsDate(.OnsetDate) Then
                        objRecordsetSub.Fields!HospitalDate = CDate(.OnsetDate)
                    End If
                    'Debug.Print .ProcedureNCID, .Procedure, .OnsetDate, Format(.OnsetDate, "dd MMM yyyy"), .ProcedureCode, .ModifyTime
                End If
            End If
        End With
    Next
    
    objRecordsetSub.Sort = "HospitalDate Desc"
    GetHospitalizations = True

Proc_Exit:
    Set objRecordsetFrom = Nothing
    Set objProcedure = Nothing
    
    If Err.Number <> 0 Then
        GetHospitalizations = False
        gobjshared.ShowVBError Err.Number, Err.Description, cMODULE_NAME & ".GetHospitalizations", App.EXEName, vbExclamation
    End If

End Function
'*****************************************************************************************
'Description: Define a new record
'Parameter: rs, ADODB.Recordset
'Reutrn:    rs
'History:
'   Date    Developer       Description
'========== ============    ==========================================================
'09/18/01   Wilson Sun      Created
'*****************************************************************************************
Public Sub DefineRecordSet(rs As ADODB.Recordset)
On Error GoTo ErrHandler
    rs.CursorLocation = adUseClient
    rs.Fields.Append "FAMILY", adBSTR, -1, clFieldAttr
    rs.Fields.Append "RELATION", adBSTR, -1, clFieldAttr
    rs.Fields.Append "ONSETDATE", adBSTR, -1, clFieldAttr
    rs.Fields.Append "SnoID", adBSTR, -1, clFieldAttr
    rs.Open
    
    GoTo Cleanup
ErrHandler:
    Err.Raise Err.Number, App.EXEName & "." & cMODULE_NAME & "." & "DefineRecordSet:" & Err.Source, _
       Err.Description, Err.HelpFile, Err.HelpContext
Cleanup:
End Sub
'*****************************************************************************************
'Description: Add a new row to record set
'Parameter: rs,         ADODB.Recordset
'Reutrn:    objFamily,  Object
'History:
'   Date    Developer       Description
'========== ============    ==========================================================
'09/18/01   Wilson Sun      Created
'*****************************************************************************************
Public Sub AddNewRecordSet(rs As ADODB.Recordset, objFamily As Object)  ' OneFamilyHistory)
On Error GoTo ErrHandler
    Dim sRelation As String
    rs.AddNew
    With objFamily
'hwcc 11/20/2003 - List of current relationship
'            Other-see comments
'            Daughter
'            Maternal Aunt
'            Maternal Uncle
'            Maternal-Grand Mother
'            Paternal Uncle
'            Sister
'            Brother
'            Father
'            Mother
'            Paternal Aunt
'            Paternal-Grand Mother
'            Maternal -Grandfather
'            Mother
'            Son
'            Paternal -Grandfather
        Select Case Trim(.Relation)
            Case "Other-see comments"
                sRelation = " - (Other)"
            Case "Daughter"
                sRelation = " - (D)"
            Case "Maternal Aunt"
                sRelation = " - (MA)"
            Case "Maternal Uncle"
                sRelation = " - (MU)"
            Case "Maternal-Grand Mother"
                sRelation = " - (MGM)"
            Case "Paternal Uncle"
                sRelation = " - (PU)"
            Case "Sister"
                sRelation = " - (S)"
            Case "Brother"
                sRelation = " - (B)"
            Case "Father"
                sRelation = " - (F)"
            Case "Mother"
                sRelation = " - (M)"
            Case "Paternal Aunt"
                sRelation = " - (PA)"
            Case "Paternal-Grand Mother"
                sRelation = " - (PGM)"
            Case "Maternal-Grandfather"
                sRelation = " - (MGF)"
            Case "Mother"
                sRelation = " - (M)"
            Case "Son"
                sRelation = " - (S)"
            Case "Paternal-Grandfather"
                sRelation = " - (PGF)"
        End Select
        'Debug.Print Trim(.Relation)
        rs.Fields("FAMILY").Value = .FamilyHistory & sRelation
        rs.Fields("RELATION").Value = .Relation
        rs.Fields("ONSETDATE").Value = .OnsetDate
        rs.Fields("SnoID").Value = .SnoID
    End With
    GoTo Cleanup
ErrHandler:
    Err.Raise Err.Number, App.EXEName & "." & cMODULE_NAME & "." & "AddNewRecordSet:" & Err.Source, _
       Err.Description, Err.HelpFile, Err.HelpContext
Cleanup:
End Sub
'*****************************************************************************************
'Description: Create Record set for family history report
'Parameter: rsSource As ADODB.Recordset,
'           objRecordset As ADODB.Recordset,
'           objRecordsetSub As ADODB.Recordset
'Reutrn:    String
'History:
'   Date    Developer       Description
'========== ============    ==========================================================
'09/19/01   Wilson Sun      Created
'*****************************************************************************************

Public Sub GetFamilyHistroyRS( _
    ByRef rsSource As ADODB.Recordset, _
    ByRef objRecordset As ADODB.Recordset, _
    ByRef objRecordsetSub As ADODB.Recordset, _
    ByVal sSurfax As String, _
    ByVal sSectionTitle As String)
On Error GoTo ErrHandler
    Dim intCount As Integer
    Dim intIndex As Integer
 
    intCount = rsSource.RecordCount
    If intCount > 0 Then
        'objRecordset.Open
        If objRecordset.RecordCount = 0 Then
            objRecordset.AddNew
        End If
        If intCount = 1 Then
            objRecordset.Fields("FAMILY" & sSurfax).Value = vbNullString
            objRecordset.Fields("FAMILYDT" & sSurfax).Value = vbNullString
        Else
            objRecordset.Fields("FAMILY" & sSurfax).Value = "(See continuation page)"
            objRecordset.Fields("FAMILYDT" & sSurfax).Value = vbNullString
        End If
        rsSource.MoveFirst
        
        intIndex = 0
        Do While Not rsSource.EOF
            intIndex = intIndex + 1
            objRecordsetSub.AddNew
            objRecordsetSub.Fields("SECTIONTITLE").Value = sSectionTitle
            objRecordsetSub.Fields("FAMILY").Value = StrConv(rsSource.Fields("FAMILY").Value, vbProperCase)
            objRecordsetSub.Fields("FAMILYDT").Value = rsSource.Fields("ONSETDATE").Value
            rsSource.MoveNext
        Loop
    End If
    Set rsSource = Nothing
    GoTo Cleanup
ErrHandler:
    Err.Raise Err.Number, App.EXEName & "." & cMODULE_NAME & "." & "GetFamilyHistoryRS:" & Err.Source, _
       Err.Description, Err.HelpFile, Err.HelpContext
Cleanup:
End Sub
'*****************************************************************************************
'Description: Get Family History data from Problem module.
'             Clone from GetHospitalizations().
'Parameter: strUnitNumber,   String,
'           objRecordset,    ADODB.Recordset,
'           objRecordsetSub, ADODB.Recordset
'Reutrn:    String
'History:
'   Date    Developer       Description
'========== ============    ==========================================================
'09/18/01   Wilson Sun      Created
'*****************************************************************************************
Private Function GetFamilyHistory( _
    ByRef oProblemOps As Object, _
    ByRef objRecordset As ADODB.Recordset, _
    ByRef objRecordsetSub As ADODB.Recordset) As Boolean

    On Error GoTo Proc_Exit
    
    Dim objRecordsetFrom As ADODB.Recordset
    Dim objFamily As Object 'OneFamilyHistory '
    Dim rsFamily_a As ADODB.Recordset   'a. Cancer
                                        'b. Cardiovascuilar disease
                                        'c. Diabetes
                                        'd. Mental Illness
                                        'o. Other Items
    
    Set rsFamily_a = New ADODB.Recordset

    GetFamilyHistory = False

'    'Create record set to hold data
    DefineRecordSet rsFamily_a
    
    For Each objFamily In oProblemOps.FamilyHistories
        With objFamily
            'hwcc - SCR 48908, we only consider Family History with active status
            If .Status = "Active" Then AddNewRecordSet rsFamily_a, objFamily
        End With
    Next
    
    'Define Recordset for Report
    Set objRecordset = New ADODB.Recordset
    objRecordset.CursorLocation = adUseClient
    objRecordset.Fields.Append "FAMILY1", adBSTR, -1, clFieldAttr
    objRecordset.Fields.Append "FAMILY2", adBSTR, -1, clFieldAttr
    objRecordset.Fields.Append "FAMILY3", adBSTR, -1, clFieldAttr
    objRecordset.Fields.Append "FAMILY4", adBSTR, -1, clFieldAttr
    objRecordset.Fields.Append "FAMILY5", adBSTR, -1, clFieldAttr
    objRecordset.Fields.Append "FAMILYDT1", adBSTR, -1, clFieldAttr
    objRecordset.Fields.Append "FAMILYDT2", adBSTR, -1, clFieldAttr
    objRecordset.Fields.Append "FAMILYDT3", adBSTR, -1, clFieldAttr
    objRecordset.Fields.Append "FAMILYDT4", adBSTR, -1, clFieldAttr
    objRecordset.Fields.Append "FAMILYDT5", adBSTR, -1, clFieldAttr
    objRecordset.Open

    'Define Recordset for SubReport
    Set objRecordsetSub = New ADODB.Recordset
    objRecordsetSub.CursorLocation = adUseClient
    objRecordsetSub.Fields.Append "SECTIONTITLE", adBSTR, -1, clFieldAttr
    objRecordsetSub.Fields.Append "FAMILY", adBSTR, -1, clFieldAttr
    objRecordsetSub.Fields.Append "FAMILYDT", adBSTR, -1, clFieldAttr
    objRecordsetSub.Open

      Call GetFamilyHistroyRS(rsFamily_a, objRecordset, objRecordsetSub, "1", "a. CANCER")
    
    GetFamilyHistory = True
    GoTo Cleanup
Proc_Exit:
    If Err.Number <> 0 Then
        GetFamilyHistory = False
        gobjshared.ShowVBError Err.Number, Err.Description, cMODULE_NAME & ".GetFamilyHistory", App.EXEName, vbExclamation
    
    End If
    
Cleanup:
    Set rsFamily_a = Nothing
    Set objRecordsetFrom = Nothing

End Function

Public Function GetAge(vDate As Variant) As String
    
    Dim iAge As Integer
    Dim iMonths As Integer
    Dim dDOB As Date
    
    On Error GoTo ErrHandler
    dDOB = moPatient.DateOfBirth
    iAge = Int(((vDate - dDOB) / 365))
    
    If iAge < 1 Then
        iMonths = Int((vDate - dDOB) / 30) 'Wilson added Int() on 10/02/01
        If iMonths < 0 Then
            GetAge = vbNullString
        ElseIf iMonths < 1 Then
            GetAge = "1 Month"
        Else
            GetAge = CStr(iMonths) & " Month(s)"
        End If
    Else
        GetAge = CStr(iAge)
    End If
    
    Exit Function
    
ErrHandler:
        gobjshared.ShowVBError Err.Number, Err.Description, cMODULE_NAME & ".GetAge", App.EXEName, vbExclamation

End Function



Private Function GetHear(ByRef objRecordset As ADODB.Recordset) As Boolean

    Dim strSQL As String
    Dim objRecordsetFrom As ADODB.Recordset
    Dim strYear1 As String
    Dim strYear2 As String

    On Error GoTo Proc_Exit
    GetHear = False
    
    strSQL = GetVitalsSQL("Hear", 33971)
    
    Set objRecordset = goDBDAS.OpenRecordset(strSQL)
    If Not (objRecordset.BOF And objRecordset.EOF) Then
        objRecordset.MoveFirst
        If IsNull(objRecordset.Fields("Date_Format").Value) Then
            strYear1 = NULLStr
        Else
            strYear1 = Mid$(objRecordset.Fields("Date_Format").Value, 8)
        End If
        objRecordset.MoveNext
        Do While Not objRecordset.EOF
            If IsNull(objRecordset.Fields("Date_Format").Value) Then
                strYear2 = NULLStr
            Else
                strYear2 = Mid$(objRecordset.Fields("Date_Format").Value, 8)
            End If
            
            If strYear1 = strYear2 And strYear2 <> NULLStr Then
                objRecordset.Delete
                objRecordset.Update
            Else
                strYear1 = strYear2
            End If
            objRecordset.MoveNext
        Loop
    End If
    
    GetHear = True

Proc_Exit:
    Set objRecordsetFrom = Nothing
    
    If Err.Number <> 0 Then
        gobjshared.ShowVBError Err.Number, Err.Description, cMODULE_NAME & ".GetHEAR", App.EXEName, vbExclamation
    End If
    
End Function

Private Function GetBP(ByRef objRecordset As ADODB.Recordset) As Boolean

    On Error GoTo Proc_Exit

    Dim strSQL              As String
    Dim objRecordsetFrom    As ADODB.Recordset
    Dim strYear1            As String
    Dim strYear2            As String

    GetBP = False

    'BP S - 1985
    'BP D - 1976

    '-- Create SQL
    strSQL = "SELECT 'BP' AS CATEGORY, " & _
                    "((select vs3.vitalsvalue from vitals v3, vitalsignsvalue vs3 " & _
                    "where vs3.dataid = v3.dataid and " & _
                    "vs3.vitalstype = 1985 and " & _
                    "v3.unit_number = " & moPatient.UnitNumber & " and " & _
                    "v3.entrydate = v.entrydate) || '/' || " & _
                    "(select vs3.vitalsvalue from vitals v3, vitalsignsvalue vs3 " & _
                    "where vs3.dataid = v3.dataid and " & _
                    "vs3.vitalstype = 1976 and " & _
                    "v3.unit_number =  " & moPatient.UnitNumber & " and " & _
                    "v3.entrydate = v.entrydate) || ' mmHg') as Result, " & _
                    "TO_NUMBER(TO_CHAR(V.ENTRYDATE, 'YYYY')) AS Result_Year, " & _
                    "TO_CHAR(V.ENTRYDATE,'DD MON YYYY') as Date_Format, " & _
                    "V.ENTRYDATE as Event_Date " & _
               "FROM VITALSIGNSVALUE VS, VITALS V " & _
               "WHERE " & _
                   "(VS.DataId = V.DATAID) AND " & _
                   "(VS.VITALSTYPE = 1976 OR VS.VITALSTYPE = 1985) AND " & _
                   "(V.UNIT_NUMBER = " & moPatient.UnitNumber & ") " & _
                   "order by event_date DESC "
    
    Set objRecordsetFrom = goDBDAS.OpenRecordset(strSQL)

    '--Copy Recordset
    If Not DuplicateDefinition(objRecordsetFrom, objRecordset, True) Then
        GetBP = False
        GoTo Proc_Exit
    End If

    If objRecordset.RecordCount > 1 Then
        objRecordset.MoveFirst
        strYear1 = Mid$(objRecordset.Fields("Date_Format").Value, 8)
        objRecordset.MoveNext
        Do While Not objRecordset.EOF
    
            strYear2 = Mid$(objRecordset.Fields("Date_Format").Value, 8)
    
            If strYear1 = strYear2 Then
                objRecordset.Delete
                objRecordset.Update
            Else
                strYear1 = strYear2
            End If
            objRecordset.MoveNext
        Loop
    End If

    GetBP = True

Proc_Exit:
    Set objRecordsetFrom = Nothing
    If Err.Number <> 0 Then
        GetBP = False
        gobjshared.ShowVBError Err.Number, Err.Description, cMODULE_NAME & ".GetBP", App.EXEName, vbExclamation
    End If
    Exit Function
End Function
'*****************************************************************************************
'Description: GetWeight()
'       Get Weight value from  observation_value ov, clinical_observation co, event e table
'Parameter: objRecordset,           ADODB.Recordset,
'Reutrn:    Boolean
'History:
'   Date    Developer       Description
'========== ============    ==========================================================
'11/16/01   Wilson Sun      Check OBSERVATION_VALUE.UNITS_NCID value to determine
'                           lbs or kg for Weight; ins or cm for Hight
'                           UNITS_NCID  Display value   for What
'                           ==========  ==============  ===========
'                           1636        lb              WEIGHT
'                           1631        kg              WEIGHT
'??/??/??   Brain Webb      Created.
'*****************************************************************************************
Private Function GetWeight(ByRef objRecordset As ADODB.Recordset) As Boolean

    On Error GoTo Proc_Exit
    
    Dim strSQL              As String
    Dim objRecordsetFrom    As ADODB.Recordset
    Dim strYear1            As String
    Dim strYear2            As String

    GetWeight = False

    strSQL = GetVitalsSQL("WEIGHT", 2178)
    
    Set objRecordsetFrom = goDBDAS.OpenRecordset(strSQL)

    '-- Copy Recordset
    If Not DuplicateDefinition(objRecordsetFrom, objRecordset, True) Then
        GetWeight = False
        GoTo Proc_Exit
    End If
    

    If objRecordset.RecordCount > 1 Then
        objRecordset.MoveFirst
        If IsNull(objRecordset.Fields("Date_Format").Value) Then
            strYear1 = NULLStr
        Else
            strYear1 = Mid$(objRecordset.Fields("Date_Format").Value, 8)
        End If
        objRecordset.MoveNext
        Do While Not objRecordset.EOF
    
            If IsNull(objRecordset.Fields("Date_Format").Value) Then
                strYear2 = NULLStr
            Else
                strYear2 = Mid$(objRecordset.Fields("Date_Format").Value, 8)
            End If
    
            If strYear1 = strYear2 And strYear2 <> NULLStr Then
                objRecordset.Delete
                objRecordset.Update
            Else
                strYear1 = strYear2
            End If
            objRecordset.MoveNext
        Loop
    End If


    GetWeight = True

Proc_Exit:
    Set objRecordsetFrom = Nothing

    If Err.Number <> 0 Then
        gobjshared.ShowVBError Err.Number, Err.Description, cMODULE_NAME & ".GetWeight", App.EXEName, vbExclamation
    End If

End Function
'*****************************************************************************************
'Description: GetHeight()
'       Get Height value from  observation_value ov, clinical_observation co, event e table
'Parameter: objRecordset,           ADODB.Recordset,
'Reutrn:    Boolean
'History:
'   Date    Developer       Description
'========== ============    ==========================================================
'11/16/01   Wilson Sun      Check OBSERVATION_VALUE.UNITS_NCID value to determine
'                           lbs or kg for Weight; ins or cm for Hight
'                           UNITS_NCID  Display value   for What
'                           ==========  ==============  ===========
'                           1606        inc             HEIGHT
'                           1617        cm              HEIGHT
'??/??/??   Brain Webb      Created.
'*****************************************************************************************
Private Function GetHeight(ByRef objRecordset As ADODB.Recordset) As Boolean

    On Error GoTo Proc_Exit
    
    GetHeight = False

    'Height 110675

    Dim strSQL As String
    Dim objRecordsetFrom As ADODB.Recordset
    Dim strYear1 As String
    Dim strYear2 As String

    '-- Create SQL
    ' This is a temporary fix for one of the item Height
    '  displayed as 5'4" yd in SCR 3/4/2004. The cause is that
    '   vital module puts in ncid 1608 as ft, in. But in rsform table
    '    and ncids table, 1608 is defined as yd instead!
    strSQL = "SELECT 'HEIGHT' AS CATEGORY, " & _
                    "(VS.VITALSVALUE || ' ' || DECODE(NCIDS.ncid, 1608, ' ', 1606, 'in')) AS Result, " & _
                    "TO_NUMBER(TO_CHAR(V.ENTRYDATE, 'YYYY')) AS Result_Year, " & _
                    "TO_CHAR(V.ENTRYDATE,'DD MON YYYY') as Date_Format," & _
                    "V.ENTRYDATE as Event_Date " & _
               "FROM VITALSIGNSVALUE VS, NCIDS, VITALS V " & _
               "WHERE (VS.Units = NCIDS.NCID) AND " & _
                   "(VS.DataId = V.DATAID) AND " & _
                   "(VS.VITALSTYPE = " & CStr(110675) & ") AND " & _
                   "(V.UNIT_NUMBER = " & moPatient.UnitNumber & ")" & " Order by Event_Date DESC "
    
    Set objRecordset = goDBDAS.OpenRecordset(strSQL)
    If Not (objRecordset.BOF And objRecordset.EOF) Then
        objRecordset.MoveFirst
        '--- Begin:SCR #24523  Developer: Sherry Wang 06/11/2002 11:49 AM
        'Fixed the VB 94 error - "Invalid use of NULL".
        If IsNull(objRecordset.Fields("Date_Format").Value) Then
            strYear1 = NULLStr
        Else
            strYear1 = Mid$(objRecordset.Fields("Date_Format").Value, 8)
        End If
        '--- End: SCR #24523
        
        objRecordset.MoveNext
        Do While Not objRecordset.EOF
            If IsNull(objRecordset.Fields("Date_Format").Value) Then
                strYear2 = NULLStr
            Else
                strYear2 = Mid$(objRecordset.Fields("Date_Format").Value, 8)
            End If
    
            If strYear1 = strYear2 And strYear2 <> NULLStr Then
            Else
                strYear1 = strYear2
            End If
            objRecordset.MoveNext
        Loop
    End If


    GetHeight = True

Proc_Exit:
    Set objRecordsetFrom = Nothing
    If Err.Number <> 0 Then
        GetHeight = False
        gobjshared.ShowVBError Err.Number, Err.Description, cMODULE_NAME & ".GetHeight", App.EXEName, vbExclamation
    End If
    
End Function

Private Function GetRubella(ByRef objRecordset As ADODB.Recordset) As Boolean

    'Rubella 22394

    Dim strSQL As String
    Dim objRecordsetFrom As ADODB.Recordset
    Dim strYear1 As String
    Dim strYear2 As String

    On Error GoTo Proc_Exit
    GetRubella = False

    Set objRecordset = New ADODB.Recordset
    objRecordset.Fields.Append "Category", adLongVarWChar, -1, clFieldAttr
    objRecordset.Fields.Append "Result_Year", adLongVarWChar, -1, clFieldAttr
    objRecordset.Fields.Append "Result", adLongVarWChar, -1, clFieldAttr
    objRecordset.Fields.Append "Date_Format", adLongVarWChar, -1, clFieldAttr
    objRecordset.Fields.Append "Event_Date", adLongVarWChar, -1, clFieldAttr
    objRecordset.Open
    GetRubella = True
    
    Exit Function

Proc_Exit:
    Set objRecordsetFrom = Nothing
    If Err.Number <> 0 Then
        GetRubella = False
        gobjshared.ShowVBError Err.Number, Err.Description, cMODULE_NAME & ".GetRubella", App.EXEName, vbExclamation
    End If

End Function

Private Function GetTesticular(ByRef objRecordset As ADODB.Recordset) As Boolean

    Dim strSQL As String
    Dim objRecordsetFrom As ADODB.Recordset
    Dim strYear1 As String
    Dim strYear2 As String

    On Error GoTo Proc_Exit
    GetTesticular = False

    GetTesticular = True

    Exit Function
Proc_Exit:
    Set objRecordsetFrom = Nothing
    If Err.Number <> 0 Then
        GetTesticular = False
        gobjshared.ShowVBError Err.Number, Err.Description, cMODULE_NAME & ".GetTesticular", App.EXEName, vbExclamation
    End If

End Function

Private Function GetMammogram(ByRef objRecordset As ADODB.Recordset) As Boolean

    'Mammogram 125339

    Dim strSQL As String
    Dim objRecordsetFrom As ADODB.Recordset
    Dim strYear1 As String
    Dim strYear2 As String

    On Error GoTo Proc_Exit
    GetMammogram = False

    Set objRecordset = New ADODB.Recordset
    objRecordset.Fields.Append "Category", adLongVarWChar, -1, clFieldAttr
    objRecordset.Fields.Append "Result_Year", adLongVarWChar, -1, clFieldAttr
    objRecordset.Fields.Append "Result", adLongVarWChar, -1, clFieldAttr
    objRecordset.Fields.Append "Date_Format", adLongVarWChar, -1, clFieldAttr
    objRecordset.Fields.Append "Event_Date", adLongVarWChar, -1, clFieldAttr
    objRecordset.Open
    GetMammogram = True
    Exit Function

Proc_Exit:
    Set objRecordsetFrom = Nothing
    
    If Err.Number <> 0 Then
        GetMammogram = False
        gobjshared.ShowVBError Err.Number, Err.Description, cMODULE_NAME & ".GetMammogram", App.EXEName, vbExclamation
    End If

End Function



Private Function GetCholesterol(ByRef objRecordset As ADODB.Recordset) As Boolean


    'Cholesterol 6097
    Dim strSQL As String
    Dim strYear1 As String
    Dim strYear2 As String

    On Error GoTo Proc_Exit
    GetCholesterol = False
    
    Set objRecordset = New ADODB.Recordset
    objRecordset.Fields.Append "Category", adLongVarWChar, -1, clFieldAttr
    objRecordset.Fields.Append "Result_Year", adLongVarWChar, -1, clFieldAttr
    objRecordset.Fields.Append "Result", adLongVarWChar, -1, clFieldAttr
    objRecordset.Fields.Append "Date_Format", adLongVarWChar, -1, clFieldAttr
    objRecordset.Fields.Append "Event_Date", adLongVarWChar, -1, clFieldAttr
    objRecordset.Open
    GetCholesterol = True
    Exit Function

Proc_Exit:
    If Err.Number <> 0 Then
        GetCholesterol = False
        gobjshared.ShowVBError Err.Number, Err.Description, cMODULE_NAME & ".GetCholesterol", App.EXEName, vbExclamation
    End If

End Function

Private Function GetHearingExam(ByRef objRecordset As ADODB.Recordset) As Boolean

    Dim strSQL As String
    Dim objRecordsetFrom As ADODB.Recordset
    Dim strYear1 As String
    Dim strYear2 As String

    On Error GoTo Proc_Exit
    GetHearingExam = False
    
    '-- Create SQL
    strSQL = GetBlock689SQL("HEARING", "HEAR_EXAM_DT")
    Set objRecordset = goDBDAS.OpenRecordset(strSQL)
    If Not (objRecordset.BOF And objRecordset.EOF) Then
        objRecordset.MoveFirst
        strYear1 = Mid$(objRecordset.Fields("Date_Format").Value, 8)
        objRecordset.MoveNext
        Do While Not objRecordset.EOF
            strYear2 = Mid$(objRecordset.Fields("Date_Format").Value, 8)
            If strYear1 = strYear2 Then
            Else
                strYear1 = strYear2
            End If
            objRecordset.MoveNext
        Loop
    End If
    
    GetHearingExam = True

Proc_Exit:
    Set objRecordsetFrom = Nothing
    If Err.Number <> 0 Then
        GetHearingExam = False
        gobjshared.ShowVBError Err.Number, Err.Description, cMODULE_NAME & ".GetHearingExam", App.EXEName, vbExclamation
    End If
    
End Function
Private Function GetDentalExam(ByRef objRecordset As ADODB.Recordset) As Boolean

    On Error GoTo Proc_Exit
    
    GetDentalExam = False

    Dim strSQL As String
    Dim objRecordsetFrom As ADODB.Recordset
    Dim strYear1 As String
    Dim strYear2 As String

    '-- Create SQL
    strSQL = GetBlock689SQL("DENTAL", "LAST_DENTAL_EXAM_DT")
    Set objRecordsetFrom = goDBDAS.OpenRecordset(strSQL)

    '-- Copy Recordset
    If Not DuplicateDefinition(objRecordsetFrom, objRecordset, True) Then
        GetDentalExam = False
        GoTo Proc_Exit
    End If

    If objRecordset.RecordCount > 1 Then

        objRecordset.MoveFirst
        strYear1 = Mid$(objRecordset.Fields("Date_Format").Value, 8)
        objRecordset.MoveNext
        Do While Not objRecordset.EOF
    
            strYear2 = Mid$(objRecordset.Fields("Date_Format").Value, 8)
    
            If strYear1 = strYear2 Then
                objRecordset.Delete
                objRecordset.Update
            Else
                strYear1 = strYear2
            End If
            objRecordset.MoveNext
        Loop
    End If

    GetDentalExam = True

Proc_Exit:
    Set objRecordsetFrom = Nothing
    
    If Err.Number <> 0 Then
        GetDentalExam = False
        gobjshared.ShowVBError Err.Number, Err.Description, cMODULE_NAME & ".GetDentalExam", App.EXEName, vbExclamation
    End If

End Function

Private Function GetVisionExam(ByRef objRecordset As ADODB.Recordset) As Boolean

    Dim strSQL As String
    Dim objRecordsetFrom As ADODB.Recordset
    Dim strYear1 As String
    Dim strYear2 As String

    On Error GoTo Proc_Exit
    GetVisionExam = False
    
    '-- Create SQL
    strSQL = GetBlock689SQL("VISION", "LAST_VISUAL_ACUITY_DT")

    Set objRecordsetFrom = goDBDAS.OpenRecordset(strSQL)

    '-- Copy Recordset
    If Not DuplicateDefinition(objRecordsetFrom, objRecordset, True) Then
        GetVisionExam = False
        GoTo Proc_Exit
    End If

    If objRecordset.RecordCount > 1 Then
        objRecordset.MoveFirst
        strYear1 = Mid$(objRecordset.Fields("Date_Format").Value, 8)
        objRecordset.MoveNext
        Do While Not objRecordset.EOF
    
            strYear2 = Mid$(objRecordset.Fields("Date_Format").Value, 8)
    
            If strYear1 = strYear2 Then
                objRecordset.Delete
                objRecordset.Update
            Else
                strYear1 = strYear2
            End If
            objRecordset.MoveNext
        Loop
    End If

    GetVisionExam = True

Proc_Exit:
    
    Set objRecordsetFrom = Nothing
    
    If Err.Number <> 0 Then
        GetVisionExam = False
        gobjshared.ShowVBError Err.Number, Err.Description, cMODULE_NAME & ".GetVisionExam", App.EXEName, vbExclamation
    End If

End Function
Private Function GetPAP(ByRef objRecordset As ADODB.Recordset) As Boolean

    On Error GoTo Proc_Exit
    
    GetPAP = False

    'PAP 20758

    GetPAP = True
    Exit Function

Proc_Exit:
    If Err.Number <> 0 Then
        GetPAP = False
        gobjshared.ShowVBError Err.Number, Err.Description, cMODULE_NAME & ".GetPAP", App.EXEName, vbExclamation
    End If
    Exit Function
End Function



Private Function GetFecalOccultBlood(ByRef objRecordset As ADODB.Recordset) As Boolean

    Dim strSQL As String
    Dim objRecordsetFrom As ADODB.Recordset
    Dim strYear1 As String
    Dim strYear2 As String

    On Error GoTo Proc_Exit
    GetFecalOccultBlood = False
    
    Set objRecordset = New ADODB.Recordset
    objRecordset.Fields.Append "Category", adLongVarWChar, -1, clFieldAttr
    objRecordset.Fields.Append "Result_Year", adLongVarWChar, -1, clFieldAttr
    objRecordset.Fields.Append "Result", adLongVarWChar, -1, clFieldAttr
    objRecordset.Fields.Append "Date_Format", adLongVarWChar, -1, clFieldAttr
    objRecordset.Fields.Append "Event_Date", adLongVarWChar, -1, clFieldAttr
    objRecordset.Open
    GetFecalOccultBlood = True
    Exit Function

Proc_Exit:
    Set objRecordsetFrom = Nothing
    If Err.Number <> 0 Then
        GetFecalOccultBlood = False
        gobjshared.ShowVBError Err.Number, Err.Description, cMODULE_NAME & ".GetFecalOccultBlood", App.EXEName, vbExclamation
    End If

End Function

Private Function GetMaxYear(coResults As Collection, _
                            iYear As Integer) As Boolean

GetMaxYear = False
On Error GoTo ErrHandler

Dim oRec As ADODB.Recordset
Dim i As Integer

iYear = 0
For i = 1 To coResults.Count
    
    Set oRec = coResults.Item(i)
    
    If Not oRec Is Nothing Then
        If oRec.RecordCount > 0 Then
            oRec.MoveFirst
            'If a 9999 year means Result_Year is Null
            If iYear < oRec.Fields("Result_Year") And oRec.Fields("Result_Year") <> "9999" Then
                iYear = oRec.Fields("Result_Year")
            End If
        End If
    End If

Next

GetMaxYear = True
GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, cMODULE_NAME & ".GetMaxYear", App.EXEName, vbExclamation
    Resume Next
Cleanup:
    Exit Function
End Function

Private Function GetMatchingYear(iYear As Integer, _
                                 oRecSource As ADODB.Recordset, _
                                 sResult As String, _
                                 bMatch As Boolean) As Boolean

GetMatchingYear = False
On Error GoTo ErrHandler

Dim i As Integer

sResult = vbNullString
bMatch = False

If oRecSource.RecordCount > 0 Then
    oRecSource.MoveFirst
    
    For i = 1 To oRecSource.RecordCount
        oRecSource.MoveFirst
        
        If oRecSource.Fields("Result_Year") = iYear Then
            sResult = oRecSource.Fields("Result").Value
            bMatch = True
            Exit For
        End If
        
        oRecSource.MoveNext
    Next
End If

GetMatchingYear = True
GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, cMODULE_NAME & ".GetMatchingYear", App.EXEName, vbExclamation
Cleanup:
    Exit Function
    
    Err.Clear
    Resume
    
End Function
'*****************************************************************************************
'Description: GetImmunizations()
'  Get Immunication data from
'  (Imm_People, Imm_People_Vaccine, tblVaccines, tblVacc_Exempt, tblVaccine_Mfg, tblService) Tables
'Parameter: oRec,           ADODB.Recordset,
'           oRecSub,        ADODB.Recordset
'Reutrn:    Boolean
'History:
'   Date    Developer       Description
'========== ============    ==========================================================
'11/08/01   Wilson Sun      HEP A (52), HEP B (43, 44), TD (9, 35)
'                           POLIO (2, 10), INFLUENZA(15, 16), MMR (3)
'04/11/01   Wilson Sun      Created. Clone from previous GetImmunizations. SCR 11496
'*****************************************************************************************
Private Function GetImmunizations(oRec As ADODB.Recordset, _
                                 oRecSub As ADODB.Recordset) As Boolean

    Dim sStmt As String
    Dim oRec1 As ADODB.Recordset
    
    GetImmunizations = False
On Error GoTo ErrHandler
    
    '-- Create Template Recordset -------------------------------------------------------
    If Not CreateImmunizationTemplate(oRec) Then GoTo Cleanup
    oRec.Open
    oRec.AddNew

    '-- Create SQL -----------------------------------------------------------------
    If gobjshared.UseSQLServer Then
        sStmt = "SELECT tblVaccines.VACCINEID as id," & vbNewLine & _
                "tblVaccines.GENERICNAME," & vbNewLine & _
                "tblVaccines.Vaccine," & vbNewLine & _
                "Imm_People_Vaccine.Vacc_Date," & vbNewLine & _
                "CONVERT(VARCHAR(12), Imm_People_Vaccine.Vacc_Date, 106) as Imm_Date," & vbNewLine & _
                "Imm_People_Vaccine.Vacc_Series," & vbNewLine & _
                "Imm_People_Vaccine.Vacc_NextDue," & vbNewLine & _
                "Imm_People_Vaccine.Lot_Nbr," & vbNewLine & _
                "tblVaccine_Mfg.Vaccine_Mfg," & vbNewLine & _
                "Imm_People.ExemptionCode," & vbNewLine & _
                "Imm_People.ExemptionExpirationDate," & vbNewLine & _
                "tblVacc_Exempt.Vacc_Exempt_Text," & vbNewLine & _
                "Imm_People_Vaccine.Vacc_Dosage," & vbNewLine & _
                "Imm_People_Vaccine.Result," & vbNewLine & _
                "Imm_People_Vaccine.RXN_Size" & vbNewLine
    Else
        sStmt = "SELECT tblVaccines.VACCINEID as id," & vbNewLine & _
                "tblVaccines.GENERICNAME," & vbNewLine & _
                "tblVaccines.Vaccine," & vbNewLine & _
                "Imm_People_Vaccine.Vacc_Date," & vbNewLine & _
                "TO_CHAR(Imm_People_Vaccine.Vacc_Date,'DD MON YYYY') as Imm_Date," & vbNewLine & _
                "Imm_People_Vaccine.Vacc_Series," & vbNewLine & _
                "Imm_People_Vaccine.Vacc_NextDue," & vbNewLine & _
                "Imm_People_Vaccine.Lot_Nbr," & vbNewLine & _
                "tblVaccine_Mfg.Vaccine_Mfg," & vbNewLine & _
                "Imm_People.ExemptionCode," & vbNewLine & _
                "Imm_People.ExemptionExpirationDate," & vbNewLine & _
                "tblVacc_Exempt.Vacc_Exempt_Text," & vbNewLine & _
                "Imm_People_Vaccine.Vacc_Dosage," & vbNewLine & _
                "Imm_People_Vaccine.Result," & vbNewLine & _
                "Imm_People_Vaccine.RXN_Size" & vbNewLine
    End If
    
    sStmt = sStmt & "From Imm_People, Imm_People_Vaccine, tblVaccines, tblVacc_Exempt, tblVaccine_Mfg, tblService" & vbNewLine
    
    If gobjshared.UseSQLServer Then
        sStmt = sStmt & "Where tblService.PT_Service =* Imm_People.PT_Service AND" & vbNewLine & _
                "tblVaccine_Mfg.Mfg_Code =* Imm_People_Vaccine.Mfg_Code AND" & vbNewLine & _
                "tblVacc_Exempt.Vacc_Exempt =* Imm_People_Vaccine.Vacc_Exempt AND" & vbNewLine & _
                "tblVaccines.VaccineID = Imm_People_Vaccine.VaccineID AND" & vbNewLine & _
                "Imm_People_Vaccine.SPON_SSAN_FMP = Imm_People.SPON_SSAN_FMP AND" & vbNewLine & _
                "(Imm_People.SPON_SSAN_FMP in" & vbNewLine & _
                "(select sponsor_ssn.id_value + fmp.id_value" & vbNewLine & _
                "From mmi_generic_id fmp, mmi_generic_id sponsor_ssn" & vbNewLine & _
                "Where fmp.id_type_ncid = 204708 and" & vbNewLine & _
                    "fmp.unit_number = sponsor_ssn.unit_number and" & vbNewLine & _
                    "sponsor_ssn.id_type_ncid = 186334 and" & vbNewLine & _
                    "sponsor_ssn.unit_number = " & moPatient.UnitNumber & "))" & vbNewLine & _
            "Order By tblVaccines.VACCINEID, Imm_People_Vaccine.Vacc_Series, Imm_People_Vaccine.Vacc_Date desc"
    Else
        sStmt = sStmt & "Where tblService.PT_Service (+) = Imm_People.PT_Service AND" & vbNewLine & _
                "tblVaccine_Mfg.Mfg_Code (+)= Imm_People_Vaccine.Mfg_Code AND" & vbNewLine & _
                "tblVacc_Exempt.Vacc_Exempt (+) = Imm_People_Vaccine.Vacc_Exempt AND" & vbNewLine & _
                "tblVaccines.VaccineID = Imm_People_Vaccine.VaccineID AND" & vbNewLine & _
                "Imm_People_Vaccine.SPON_SSAN_FMP = Imm_People.SPON_SSAN_FMP AND" & vbNewLine & _
                "Imm_People.SPON_SSAN_FMP in" & vbNewLine & _
                "(select Trim(sponsor_ssn.id_value) || Trim(fmp.id_value)" & vbNewLine & _
                "From mmi_generic_id fmp, mmi_generic_id sponsor_ssn" & vbNewLine & _
                "Where fmp.id_type_ncid = 204708 and" & vbNewLine & _
                    "fmp.unit_number = sponsor_ssn.unit_number and" & vbNewLine & _
                    "sponsor_ssn.id_type_ncid = 186334 and" & vbNewLine & _
                    "sponsor_ssn.unit_number = " & moPatient.UnitNumber & ")" & vbNewLine & _
            "Order By tblVaccines.VACCINEID, Imm_People_Vaccine.Vacc_Series, Imm_People_Vaccine.Vacc_Date desc"
    End If
    Set oRecSub = goDBDAS.OpenRecordset(sStmt)
    
    If oRecSub.RecordCount = 0 Then
        GetImmunizations = True
        GoTo Cleanup
    End If
    
    '-- Create copy of Recordset    '-- Hep A (Series)    '52 (Hep A, adult)
    oRecSub.Filter = "id = 52"
    Call SetNextImmSeries2(1, 2, "HEPA", oRec, oRecSub)
    
    '-- Hep B (Series)    '43 (Hep B, adult)    '44 (Hep B, dialysis)
    oRecSub.Filter = "id=43 or id=44"
    Call SetNextImmSeries2(1, 3, "HEPB", oRec, oRecSub)
    
    '-- MMR (Series)    '3 (MMR)
    oRecSub.Filter = "id = 3"
    Call SetNextImmSeries2(1, 2, "MMR", oRec, oRecSub)
    
    '-- PNEUMOCOCCUS (One shot only)    '33 (Pneumococal)
    oRecSub.Filter = "id = 33"
    Call SetOneShotOnly("PNEUM_DT", oRec, oRecSub)
    
    '-- POLIO OPV (One shot only)    '2 OPV    '10 IPV
    oRecSub.Filter = "id=10 or id=2"
    Call SetOneShotOnly("POLIO_OPV", oRec, oRecSub)
    
    '-- TD  (One shot only)    '9 (Td adult)    '35 (tetanus toxoid)
    oRecSub.Filter = "id=9 or id = 35"
    Call SetOneShotOnly("TD_LAST_DT", oRec, oRecSub, "TD_DUE_DT")
    
    '-- YELLOW FEVER  (One shot only)    '37 yellow fever
    oRecSub.Filter = "id=37"
    Call SetOneShotOnly("YELLOW_DT", oRec, oRecSub, "YELLOW_DUE_DT")
    
    '-- TYPHOID (Multiple shots but series)    '25 (typhoid, oral)    '41 (typhoid, parenteral)    '53 (typhoid, parenteral, AKD (U.S. military))    '101 (typhoid, ViCPs)
    oRecSub.Filter = "id=25 or id=53 or id=41 or id=101"
    Call SetMoreShots("TYPHOID_DT", oRec, oRecSub, 6)
    
    '-- ANTHRAX (Series)    '---------------------------------------------------------------------------------
    oRecSub.Filter = "id=24"
    Call SetNextImmSeries2(1, 6, "ANTHRAX_", oRec, oRecSub)
    
    '-- IPPD  (Multiple shots but series)    '96 TST-PPD intradermal
    oRecSub.Filter = "id=96 and RXN_SIZE <> ''"
    Call SetMoreShots("PPD_DT", oRec, oRecSub, 7)
    
    '-- INFLUENZA  (Multiple shots but series)    '15 influenza, split    '16 influenza, whole
    oRecSub.Filter = "id=15 or id=16"
    Call SetMoreShots("INFLUENZA_DT", oRec, oRecSub, 7)
    
    '-- VARICELLA (Series)
    oRecSub.Filter = "id=21"
    Call SetNextImmSeries2(1, 2, "VARICELLA_DT", oRec, oRecSub)
    
    '-- MENINGO  (Multiple shots but series)
    oRecSub.Filter = "id=32"
    Call SetMoreShots("MENINGO_DT", oRec, oRecSub, 2)
    
    '-- ADENO  (Multiple shots but series)    '54 adenovirus, type 4    '55 adenovirus, type 7
    oRecSub.Filter = "id=54 or id=55"
    Call SetMoreShots("ADENO_DT", oRec, oRecSub, 2)
    
    '-- JAPANESE B ENCEPAHLITIS (Series)
    oRecSub.Filter = "id=39"
    Call SetNextImmSeries2(1, 4, "JAPANESE_DT", oRec, oRecSub)
    
    oRecSub.Filter = vbNullString
    GetImmunizations = True
    GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, cMODULE_NAME & ".GetImmunizations", App.EXEName, vbExclamation
    'Resume Next
Cleanup:
    Set oRec1 = Nothing
    
End Function
Private Function GetScreeningTemplates(coTemplates As Collection) As Boolean

    Dim oRec As ADODB.Recordset
    Dim oRec1 As ADODB.Recordset
    Dim i As Integer
    
    GetScreeningTemplates = False
    On Error GoTo ErrHandler
    
    '-- Create Template Recordset '-------------------------------------------------------
    If Not CreateScreeningTemplate(oRec) Then GoTo Cleanup
    
    Set coTemplates = New Collection
    For i = 1 To 6
        If Not DuplicateDefinition(oRec, oRec1) Then GoTo Cleanup
        oRec1.AddNew
        coTemplates.Add oRec1, CStr(i)
    Next
    
    GetScreeningTemplates = True
    
    GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, cMODULE_NAME & ".GetScreeningTemplates", App.EXEName, vbExclamation
Cleanup:
    Set oRec = Nothing
    Set oRec1 = Nothing

End Function

'*****************************************************************************************
'Description: GetMedications2()
'       Get Medication data from (rsform_context, rsform, drug, active_meds_out,
'                                       patient_order, event) tables
'Parameter: oRec,           ADODB.Recordset,
'           oRecSub,        ADODB.Recordset
'Reutrn:    Boolean
'History:
'   Date    Developer       Description
'========== ============    ==========================================================
'11/11/01   Wilson Sun      Data should not be retrieved from "INACTIVE_MEDS_OUT" table
'10/10/01   Wilson Sun      Created. Clone from GetAllergy. SCR 11493
'*****************************************************************************************
Private Function GetMedications2(oRec As ADODB.Recordset, oRecSub As ADODB.Recordset) As Boolean

    Dim sStmt As String
    
    GetMedications2 = False
    On Error GoTo ErrHandler
    
    '-- Create SQL
    '-----------------------------------------------------------------
    sStmt = "SELECT Initcap(Lower(rxl.DESCDISPLAY)) As RX, Initcap(Lower(meds.SIG)) AS SIG, " & _
        "DECODE(meds.ENDDATE,NULL,'',TO_CHAR(meds.ENDDATE,'DD MON YYYY')) AS RXEXPDT," & _
        "meds.ENDDATE as EXPDT " & vbCrLf & _
        "FROM MEDICATIONS meds, RX_LIST_FDB rxl" & vbCrLf & _
        "WHERE meds.MEDICATIONNCID = rxl.MEDID AND " & vbCrLf & _
        "meds.UNIT_NUMBER = " & moPatient.UnitNumber & " AND " & vbCrLf & _
        "(meds.STATUSNCID = " & NCID_ACTIVE & " OR meds.STATUSNCID = 14516641) "

    Set oRecSub = goDBDAS.OpenRecordset(sStmt, adOpenDynamic, adLockOptimistic)

   'Create new record set for first page
    Set oRec = New ADODB.Recordset
    oRec.Fields.Append ("Medication"), adBSTR, -1, clFieldAttr
    oRec.Fields.Append ("SIG"), adBSTR, -1, clFieldAttr
    oRec.Fields.Append ("RXEXPDT"), adBSTR, -1, clFieldAttr
    oRec.Fields.Append ("Continuation_Msg"), adBSTR, -1, clFieldAttr
    oRec.Open
    oRec.AddNew
       
    If oRecSub.RecordCount = 0 Then
        'Assign value to first page of the report
        oRec.Fields("Medication").Value = NULLStr
    End If
''    ElseIf oRecSub.RecordCount > 1 Then
''        'Assign value to first page of the report
''        oRec.Fields("Medication").Value = "(See continuation page)"
'    Else
'        oRecSub.MoveFirst
'        Do Until oRecSub.EOF
'            oRecSub("RX") = StrConv(oRecSub("RX"), vbProperCase)
'            oRecSub("SIG") = StrConv(oRecSub("SIG"), vbProperCase)
'            oRecSub.MoveNext
'        Loop
'    End If
 
    GetMedications2 = True
    Exit Function

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, cMODULE_NAME & ".GetMedications2", App.EXEName, vbExclamation
    'Resume Next
End Function

'*****************************************************************************************
'Name:        GetAllergy
'Description: Get Allergy data from ALLERGEN, and ALLERGY tables.
'             Clone from GetHospitalizations().
'Parameter: objRecordset,    ADODB.Recordset,
'           objRecordsetSub, ADODB.Recordset
'Reutrn:    String
'History:
'   Date    Developer       Description
'========== ============    ==========================================================
'11/02/01   Wilson Sun      Calling Allergy module to get NKA and NAS (No Allergies Specified)
'10/10/01   Wilson Sun      Clone from GetAllergy() to reduce the code and improve the performance.
'10/09/01   Wilson Sun      SCR 17776. Set oRecSub = oRec1 if no allergy found.
'09/26/01   Wilson Sun      Added third parameter naming oRecSub and modifed second
'                           parameter oRec to hold first reocrd set if "NO KNOWN ALLERGIES" found
'??/??/??   Brian Webb      Created
'*****************************************************************************************
Private Function GetAllergy(oRec As ADODB.Recordset, oRecSub As ADODB.Recordset) As Boolean

    Dim sStmt As String
    Dim sAllergenNCID As String
    Dim sAllergyDesc As String
    Dim sReaction As String
    Dim i As Integer
    Dim oTempRec As ADODB.Recordset
    Dim oTmpRs2 As ADODB.Recordset
    
    Const IEN_NKA            As String = "999001"
    Const IEN_NKA_2          As String = "999000"
    Const IEN_NKA_3          As String = "NKA"
    Const IEN_NKDA           As String = "999003"
    Const IEN_NKDA_2         As String = "999004"
    Const IEN_NKDA_3         As String = "NKDA"
    Const IEN_NKA_NONE       As String = "999002"
    'Const IEN_OTHER          As String = "998000"
    
    '-----------------------------------------------------------------
    '-- Create SQL
    '-- Select
    On Error GoTo ErrHandler
    GetAllergy = False
    
    'hwcc 4/15/2004
    'Check for 'No Known Allegies'(NKA) and 'No Know Drug Allergies'(NKDA),
    ' if we can't find any active NKA or NKDA in Allergies table, we will use "No Allergies Specified
    '  else we will use "No known Allergies"
    Set oTempRec = Nothing
    'Theater or ITT mode
    sStmt = "Select count(*) As NKADACnt from allergies" & vbNewLine & _
        " where allergenncid in (" & IEN_NKA & "," & IEN_NKDA & ")" & vbNewLine & _
        " and unit_number = " & moPatient.UnitNumber & vbNewLine & _
        " and updateflag <> 'D'" & vbNewLine
            
    Set oTempRec = goDBDAS.OpenRecordset(sStmt)
    
    Dim sAllergiesTxt As String
    sAllergiesTxt = cNO_ALLERGIES                'Default to no allergies found
    If Not (oTempRec.BOF And oTempRec.EOF) Then
        oTempRec.MoveFirst
        If oTempRec("NKADACnt") <> "0" Then
            sAllergiesTxt = cNO_KNOWN_ALLERGIES  'Change the txt to No known allergies
        End If
    End If
    Set oTempRec = Nothing
    'Check 'No Known Allergies' ends
    
        sStmt = "select a.DataID as event_id, a.allergen as allergy, " & vbCrLf
        sStmt = sStmt & "      To_Char(a.NOTEDDATE, 'DD MON YYYY') As ONSET_DATE, a.ALLERGENNCID, " & vbCrLf
        sStmt = sStmt & "      ncids.description as INFO_SOURCE, ac.ALLERGYCOMMENT AS COMMENTS " & vbCrLf
        sStmt = sStmt & "  From  NCIDS, ALLERGYCOMMENT ac, ALLERGIES a " & vbCrLf
        sStmt = sStmt & "  WHERE ncids.DOMAINNCID = 1462 and " & vbCrLf
        sStmt = sStmt & "    NCIDS.NCID = a.INFOSOURCENCID and " & vbCrLf
        sStmt = sStmt & "    ac.allergenncid = a.allergenncid and " & vbCrLf
        sStmt = sStmt & "    ac.unit_number = a.unit_number and " & vbCrLf
        sStmt = sStmt & "    a.UPDATEFLAG <> 'D' AND " & vbCrLf
        sStmt = sStmt & "    a.UNIT_NUMBER = " & moPatient.UnitNumber
        
        Set oTempRec = Nothing
        Set oTempRec = goDBDAS.OpenRecordset(sStmt, adOpenDynamic, adLockOptimistic)
        
        Set oRecSub = Nothing
        Set oRecSub = New ADODB.Recordset
        oRecSub.Fields.Append "Event_ID", adLongVarWChar, -1, clFieldAttr
        oRecSub.Fields.Append "ALLERGY", adLongVarWChar, -1, clFieldAttr
        oRecSub.Fields.Append "Onset_Date", adLongVarWChar, -1, clFieldAttr
        oRecSub.Fields.Append "reaction", adLongVarWChar, -1, clFieldAttr
        oRecSub.Fields.Append "info_source", adLongVarWChar, -1, clFieldAttr
        oRecSub.Fields.Append "comments", adLongVarWChar, -1, clFieldAttr
        oRecSub.Fields.Append "ALLERGENNCID", adLongVarWChar, -1, clFieldAttr
        oRecSub.Open
        oRecSub.ActiveConnection = Nothing
        
        If Not (oTempRec.BOF And oTempRec.EOF) Then
            oTempRec.MoveFirst
            With oRecSub
            Do Until oTempRec.EOF
                .AddNew
                .Fields("Event_ID") = oTempRec("Event_ID")
                .Fields("ALLERGY") = oTempRec("ALLERGY")
                .Fields("Onset_Date") = oTempRec("Onset_Date")
                .Fields("info_source") = oTempRec("info_source")
                .Fields("comments") = oTempRec("comments")
                .Fields("ALLERGENNCID") = oTempRec("ALLERGENNCID")
                
                sStmt = "SELECT Reaction, ReactionNCID FROM ALLERGYREACTION " & vbCrLf
                sStmt = sStmt & "WHERE UNIT_NUMBER = " & moPatient.UnitNumber & vbCrLf
                sStmt = sStmt & "AND AllergenNCID = " & .Fields("ALLERGENNCID") & vbCrLf
                sStmt = sStmt & "AND (UPDATEFLAG <> 'D' OR UPDATEFLAG IS NULL) "
                
                Set oTmpRs2 = Nothing
                Set oTmpRs2 = goDBDAS.OpenRecordset(sStmt, adOpenDynamic, adLockOptimistic)
                                
                If Not (oTmpRs2.BOF And oTmpRs2.EOF) Then
                    oTmpRs2.MoveFirst
                    Do Until oTmpRs2.EOF
                        If oTmpRs2("ReactionNCID") = "76281" Then
                            .Fields("reaction") = .Fields("reaction") & "Other: " & oTmpRs2("Reaction")
                        Else
                            .Fields("reaction") = .Fields("reaction") & oTmpRs2("Reaction")
                        End If
                        oTmpRs2.MoveNext
                        If Not oTmpRs2.EOF Then
                            .Fields("reaction") = .Fields("reaction") & ", "
                        End If
                    Loop
                Else
                    .Fields("reaction") = NULLStr
                End If
                
                .Update
                oTempRec.MoveNext
            Loop
            End With
        End If
        Set oTempRec = Nothing
        Set oTmpRs2 = Nothing
    
    '-- Get Reactions    '--------------------------------------------------------------------------
    Set oRec = New ADODB.Recordset 'Create new record set
    oRec.Fields.Append ("ALLERGY"), adBSTR, -1, clFieldAttr
    oRec.Open
    oRec.AddNew
    oRec.Fields("ALLERGY").Value = cNO_ALLERGIES
    
    'added for checking 'NO KNOWN ALLERGIES'. If it is, delete it.
    If oRecSub.BOF And oRecSub.EOF Then
       oRecSub.AddNew
       oRecSub.Fields("ALLERGY").Value = sAllergiesTxt
    Else
 
        oRecSub.MoveFirst
        Do While Not oRecSub.EOF
            If InStr(1, oRecSub("COMMENTS"), "***DELETE***", vbTextCompare) > 0 Then
               oRecSub.Delete adAffectCurrent
            ElseIf oRecSub("ALLERGENNCID") = IEN_NKA Or _
               oRecSub("ALLERGENNCID") = IEN_NKA_2 Or _
               oRecSub("ALLERGENNCID") = IEN_NKA_3 Or _
               oRecSub("ALLERGENNCID") = IEN_NKA_NONE Then
               oRec.Fields("ALLERGY").Value = cNO_KNOWN_ALLERGIES
               oRecSub.Delete adAffectCurrent
               If oRecSub.RecordCount = 0 Then
                  oRecSub.AddNew
                  oRecSub.Fields("ALLERGY").Value = cNO_KNOWN_ALLERGIES
               End If
            ElseIf oRecSub("ALLERGENNCID") = IEN_NKDA Or _
               oRecSub("ALLERGENNCID") = IEN_NKDA_2 Or _
               oRecSub("ALLERGENNCID") = IEN_NKDA_3 Then
               oRec.Fields("ALLERGY").Value = cNO_KNOWN_ALLERGIES  'SCR 27600
               oRecSub.Delete adAffectCurrent
               If oRecSub.RecordCount = 0 Then
                  oRecSub.AddNew
                  oRecSub.Fields("ALLERGY").Value = cNO_KNOWN_ALLERGIES
               End If
            Else
                oRecSub("ALLERGY") = StrConv(oRecSub("ALLERGY"), vbProperCase)
            End If
            oRecSub.MoveNext
        Loop
        
        If oRecSub.RecordCount = 0 Then
            oRecSub.AddNew
            oRecSub.Fields("ALLERGY").Value = oRec.Fields("ALLERGY").Value
        End If
        
   End If
     
   If Not (oRecSub.BOF And oRecSub.EOF) Then
       oRecSub.MoveFirst
   End If
   If Not (oRec.BOF And oRec.EOF) Then
       oRec.MoveFirst
   End If
   
   GetAllergy = True
   GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, cMODULE_NAME & ".GetAllergy", App.EXEName, vbExclamation
Cleanup:
    Set oTempRec = Nothing
    Set oTmpRs2 = Nothing
End Function


Private Function GetScreeningExams(coAllResults As Collection) As Boolean

    Dim i As Integer
    Dim oRec As ADODB.Recordset
    Dim oRec1 As ADODB.Recordset
    Dim oRecTemp As ADODB.Recordset
    Dim coScreen As Collection
    Dim coResults As Collection
    Dim oRec2 As ADODB.Recordset
    Dim iYear As Integer
    Dim bMatch As Boolean
    Dim sResult As String
    
    GetScreeningExams = False
    On Error GoTo ErrHandler
    
    '-- Create Template Recordset
    '-------------------------------------------------------
    If Not GetScreeningTemplates(coScreen) Then Exit Function
    Set coResults = New Collection
    
    '-- HEAR Results    '-----------------------------------------------------------------
    If Not GetHear(oRec) Then Exit Function
    coResults.Add oRec, "HEAR"
    
    '-- HEIGHT
    'Wilson uncommented following call to match calling number of GetMatchingYear()
    '-----------------------------------------------------------------
    If Not GetHeight(oRec) Then Exit Function
    coResults.Add oRec, "HEIGHT"
    
    '-- WEIGHT
    '-----------------------------------------------------------------
    If Not GetWeight(oRec) Then Exit Function
    coResults.Add oRec, "WEIGHT"
    
    '-- BP
    '-----------------------------------------------------------------
    If Not GetBP(oRec) Then Exit Function
    coResults.Add oRec, "BP"
    
    '-- CHOLESTEROL
    '-----------------------------------------------------------------
    If Not GetCholesterol(oRec) Then Exit Function
    coResults.Add oRec, "CHOLESTEROL"
    
    '-- HEARING
    '-----------------------------------------------------------------
    If Not GetHearingExam(oRec) Then Exit Function
    coResults.Add oRec, "HEARING"
    
    '-- DENTAL EXAM
    '-----------------------------------------------------------------
    If Not GetDentalExam(oRec) Then Exit Function
    coResults.Add oRec, "DENTAL"
    
    '-- VISION EXAM
    '-----------------------------------------------------------------
    If Not GetVisionExam(oRec) Then Exit Function
    coResults.Add oRec, "VISION"
    
    '-- MAMMOGRAM
    '-----------------------------------------------------------------
    If Not GetMammogram(oRec) Then Exit Function
    coResults.Add oRec, "MAMMOGRAM"
    
    ''-- PAP 'Wilson: Waiting for the range for display
    ''-----------------------------------------------------------------
    'If Not GetPAP(oRec) Then exit function
    'coResults.Add oRec, "PAP"
    
    '-- FECAL OCCULT BLOOD
    '-----------------------------------------------------------------
    If Not GetFecalOccultBlood(oRec) Then Exit Function
    coResults.Add oRec, "FECAL"
    
    '-- TESTICULAR
    '-----------------------------------------------------------------
    If Not GetTesticular(oRec) Then Exit Function
    coResults.Add oRec, "TESTICULAR"
    
    '-- RUBELLA
    '-----------------------------------------------------------------
    If Not GetRubella(oRec) Then Exit Function
    coResults.Add oRec, "RUBELLA"
    
    '-- Get Year to start this section
    If Not GetMaxYear(coResults, iYear) Then Exit Function
    
    '---------------------------------------------------------------------------------
    'Wilson Begin:
    '10/01/01 SCR 16953
    'Check iYear value first, if there is no any screening entered, this value is 0.
    'So, do not do anything after Wilson End
    '---------------------------------------------------------------------------------
    If iYear = 0 Then
        For i = 1 To 2 ' Most recently years only. 6
            Set oRec1 = coScreen.Item(i)
            coAllResults.Add oRec1, "screen" & CStr(i)
        Next
        GetScreeningExams = True
        Exit Function
    End If
    'Wilson End
    
    For i = 1 To 2 ' Most recently years only. 6
    
        Set oRec1 = coScreen.Item(i)
        oRec1.Fields("YEAR").Value = CStr(iYear)
    
        '-- Age        '-------------------------------------------------------------
        If Year(Now) = iYear Then
            oRec1.Fields("Age").Value = GetAge(Now)
        Else
            oRec1.Fields("Age").Value = GetAge(CDate("12/31/" + CStr(iYear)))
        End If
    
        '-- HEAR        '-------------------------------------------------------------
        If Not GetMatchingYear(iYear, coResults("HEAR"), sResult, bMatch) Then Exit Function
        If bMatch Then oRec1.Fields("HEAR").Value = sResult
    
        '-- HEIGHT        '-------------------------------------------------------------
        If Not GetMatchingYear(iYear, coResults("HEIGHT"), sResult, bMatch) Then Exit Function
        If bMatch Then oRec1.Fields("HEIGHT").Value = sResult
    
        '-- WEIGHT        '-------------------------------------------------------------
        If Not GetMatchingYear(iYear, coResults("WEIGHT"), sResult, bMatch) Then Exit Function
        If bMatch Then oRec1.Fields("WEIGHT").Value = sResult
    
        '-- BP        '-------------------------------------------------------------
        If Not GetMatchingYear(iYear, coResults("BP"), sResult, bMatch) Then Exit Function
        If bMatch Then oRec1.Fields("BP").Value = sResult
    
        '-- CHOLESTEROL        '-------------------------------------------------------------
        If Not GetMatchingYear(iYear, coResults("CHOLESTEROL"), sResult, bMatch) Then Exit Function
        If bMatch Then oRec1.Fields("CHOLESTEROL").Value = sResult
    
        '-- HEARING        '-------------------------------------------------------------
        If Not GetMatchingYear(iYear, coResults("HEARING"), sResult, bMatch) Then Exit Function
        If bMatch Then oRec1.Fields("HEARING").Value = sResult
    
        '-- DENTAL EXAM        '-------------------------------------------------------------
        If Not GetMatchingYear(iYear, coResults("DENTAL"), sResult, bMatch) Then Exit Function
        If bMatch Then oRec1.Fields("DENTAL").Value = sResult
        
        '-- VISION EXAM        '-------------------------------------------------------------
        If Not GetMatchingYear(iYear, coResults("VISION"), sResult, bMatch) Then Exit Function
        If bMatch Then oRec1.Fields("VISION").Value = sResult
        
        '-- MAMMOGRAM        '-------------------------------------------------------------
        If Not GetMatchingYear(iYear, coResults("MAMMOGRAM"), sResult, bMatch) Then Exit Function
        If bMatch Then
            '************************************************************************************
            'Special handling for sResult.
            'Since sResult is Hex string, we should convert them to
            'Decimal string, then use StrConv() to convert to ASCII string
            '************************************************************************************
            sResult = HexToStr(sResult)
            If InStr(1, sResult, " NORMAL", vbTextCompare) = 0 Then 'NORMAL not found
                oRec1.Fields("MAMMOGRAM").Value = "X"
            Else 'NORMAL found
                oRec1.Fields("MAMMOGRAM").Value = "N"
            End If
        End If
    
           
        '-- PAP 'Wilson: Waiting for the range for display
        '-------------------------------------------------------------
    '    If Not GetMatchingYear(iYear, coResults("PAP"), sResult, bMatch) Then exit function
    '    If bMatch Then oRec1.Fields("PAP").Value = sResult
    
        '-- FECAL OCCULT BLOOD        '-------------------------------------------------------------
        If Not GetMatchingYear(iYear, coResults("FECAL"), sResult, bMatch) Then Exit Function
    
        If bMatch Then
            If sResult = "NEGATIVE" Then
                oRec1.Fields("FECAL").Value = "N"
            Else
                oRec1.Fields("FECAL").Value = "X"
            End If
        End If
        
        '-- TESTICULAR
        '-------------------------------------------------------------
        If Not GetMatchingYear(iYear, coResults("TESTICULAR"), sResult, bMatch) Then Exit Function
    
        If bMatch Then
            '************************************************************************************
            'Special handling for sResult.
            'Since sResult is Hex string, we should convert them to
            'Decimal string, then use StrConv() to convert to ASCII string
            '************************************************************************************
            sResult = HexToStr(sResult)
            Debug.Print sResult
            
            If InStr(1, sResult, " NORMAL", vbTextCompare) = 0 Then
                'NORMAL not found
                oRec1.Fields("TESTICULAR").Value = "X"
            Else
                'NORMAL found
                oRec1.Fields("TESTICULAR").Value = "N"
            End If
        End If
        
        '-- RUBELLA        '-------------------------------------------------------------
        If Not GetMatchingYear(iYear, coResults("RUBELLA"), sResult, bMatch) Then Exit Function
        If bMatch Then oRec1.Fields("RUBELLA").Value = sResult
    
        '-- Add Screening Column to Collection        '--------------------------------------------------------------
        coAllResults.Add oRec1, "screen" & CStr(i)
        iYear = iYear - 1
    Next
    
    '''-- SKIN EXAM
    '''-----------------------------------------------------------------
    ''If Not GetSkinExam(sUnitNumber, oRecSkin) Then GoTo CleanUp
    ''
    '''-- BREAST EXAM
    '''-----------------------------------------------------------------
    ''If Not GetBreastExam(sUnitNumber, oRecBreast) Then GoTo CleanUp
    ''
    '''-- SIGMOID
    '''-----------------------------------------------------------------
    ''If Not GetSigmoid(sUnitNumber, oRecSigmoid) Then GoTo CleanUp
    ''
    '''-- Colonoscopy
    '''-----------------------------------------------------------------
    ''If Not GetColonoscopy(sUnitNumber, oRecColonoscopy) Then GoTo CleanUp
    ''
    '''-- PROSTATE
    '''-----------------------------------------------------------------
    ''If Not GetProstate(sUnitNumber, oRecProstate) Then GoTo CleanUp
    ''
    '''-- Occupational
    '''-----------------------------------------------------------------
    ''If Not GetOccupational(sUnitNumber, oRecOccupational) Then GoTo CleanUp
    
    GetScreeningExams = True
    GoTo CleanUpAndExit
    
ErrHandler:
        gobjshared.ShowVBError Err.Number, Err.Description, cMODULE_NAME & ".GetScreeningExams", App.EXEName, vbExclamation
CleanUpAndExit:
    Set oRec = Nothing
    Set oRec1 = Nothing
    Set oRecTemp = Nothing
    Set coScreen = Nothing
    Set coResults = Nothing
    Set oRec2 = Nothing

End Function

Private Function GetReadiness(oRec As ADODB.Recordset) As Boolean

    Dim sStmt As String
    Dim i As Integer
    Dim oRec1 As ADODB.Recordset
    Dim oRec2 As ADODB.Recordset
    Dim sDate As String
    Dim sClass As String
    
    GetReadiness = False
    On Error GoTo ErrHandler
    
    '-- Create Template Recordset '-------------------------------------------------------
    If Not CreateReadinessTemplate(oRec) Then GoTo Cleanup
    
    oRec.Open
    oRec.AddNew
    oRec.Fields("job_id").Value = 1
    
    '-- Get DNA    '---------------------------------------------------------------------
    If gobjshared.UseSQLServer Then
        sStmt = "select CONVERT(VARCHAR(12), DNA_TEST_DT, 106) as DNA_DT" & vbNewLine & _
            "from readiness" & vbNewLine & _
            "where DNA_TEST_DT is not Null and" & vbNewLine & _
            "patient_unit_number =" & moPatient.UnitNumber & vbNewLine & _
            "order by DNA_TEST_DT desc"
    Else
        sStmt = "select TO_CHAR(DNA_TEST_DT,'DD MON YYYY') as DNA_DT" & vbNewLine & _
            "from readiness" & vbNewLine & _
            "where DNA_TEST_DT is not Null and" & vbNewLine & _
            "patient_unit_number =" & moPatient.UnitNumber & vbNewLine & _
            "order by DNA_TEST_DT desc"
    End If
    
    Set oRec1 = goDBDAS.OpenRecordset(sStmt, adOpenDynamic, adLockOptimistic)
    
    oRec.Fields("DNA_DT").Value = NULLStr
    If oRec1.RecordCount > 0 Then
       If Not (CheckNULLDt(oRec1.Fields("DNA_DT").Value)) Then
           oRec.Fields("DNA_DT").Value = oRec1.Fields("DNA_DT").Value
       End If
    End If
    
    '-- Get BLOOD TYPE    '---------------------------------------------------------------------
    If gobjshared.UseSQLServer Then
        sStmt = "select distinct CONVERT(VARCHAR(12), BLOOD_TYPE_TEST_DT, 106) as BLOOD_TYPE_DT," & vbNewLine & _
            "BLOOD_TYPE_TEST_DT as Exam_DT, BLOOD_TYPE_TEST_RSLT" & vbNewLine & _
            "from readiness" & vbNewLine & _
            "where BLOOD_TYPE_TEST_DT is not Null and" & vbNewLine & _
            "patient_unit_number =" & moPatient.UnitNumber & vbNewLine & _
            "order by Exam_DT desc"
    Else
        sStmt = "select distinct TO_CHAR(BLOOD_TYPE_TEST_DT,'DD MON YYYY') as BLOOD_TYPE_DT," & vbNewLine & _
            "BLOOD_TYPE_TEST_DT as Exam_DT, BLOOD_TYPE_TEST_RSLT" & vbNewLine & _
            "from readiness" & vbNewLine & _
            "where BLOOD_TYPE_TEST_RSLT is not Null and" & vbNewLine & _
            "patient_unit_number =" & moPatient.UnitNumber & vbNewLine & _
            "order by Exam_DT desc"
    End If
    
    Set oRec1 = goDBDAS.OpenRecordset(sStmt, adOpenDynamic, adLockOptimistic)
    
    oRec.Fields("BLOOD_TYPE_DT").Value = NULLStr
    oRec.Fields("BLOOD_TYPE_RESULT").Value = NULLStr
    If oRec1.RecordCount > 0 Then
        oRec.Fields("BLOOD_TYPE_RESULT").Value = oRec1.Fields("BLOOD_TYPE_TEST_RSLT").Value
        If Not oRec1.Fields("BLOOD_TYPE_DT").Value = Null Then
            oRec.Fields("BLOOD_TYPE_DT").Value = oRec1.Fields("BLOOD_TYPE_DT").Value
        End If
    End If
    
    '-- Get G6PD    '---------------------------------------------------------------------
    If gobjshared.UseSQLServer Then
        sStmt = "select CONVERT(VARCHAR(12), G6PD_TEST_DT, 106) as G6PD_DT, G6PD_TEST_RSLT" & vbNewLine & _
            "from readiness" & vbNewLine & _
            "where G6PD_TEST_DT is not Null and" & vbNewLine & _
            "patient_unit_number =" & moPatient.UnitNumber & vbNewLine & _
            "order by G6PD_TEST_DT desc"
    Else
        sStmt = "select TO_CHAR(G6PD_TEST_DT,'DD MON YYYY') as G6PD_DT, G6PD_TEST_RSLT" & vbNewLine & _
            "from readiness" & vbNewLine & _
            "where G6PD_TEST_DT is not Null and" & vbNewLine & _
            "patient_unit_number =" & moPatient.UnitNumber & vbNewLine & _
            "order by G6PD_TEST_DT desc"
    End If
    Set oRec1 = goDBDAS.OpenRecordset(sStmt, adOpenDynamic, adLockOptimistic)
    
    oRec.Fields("G6PD_DT").Value = NULLStr
    oRec.Fields("G6PD_RESULT").Value = NULLStr
    If oRec1.RecordCount > 0 Then
      If Not (CheckNULLDt(oRec1.Fields("G6PD_DT").Value)) Then
        oRec.Fields("G6PD_DT").Value = oRec1.Fields("G6PD_DT").Value
        oRec.Fields("G6PD_RESULT").Value = oRec1.Fields("G6PD_TEST_RSLT").Value
      End If
    End If
    
    '-- Get SICKLE CELL     '---------------------------------------------------------------------
    If gobjshared.UseSQLServer Then
        sStmt = "select CONVERT(VARCHAR(12), SICKLE_CELL_TEST_DT, 106) as SICKLE_CELL_DT, SICKLE_CELL_TEST_RSLT" & vbNewLine & _
            "from readiness" & vbNewLine & _
            "where SICKLE_CELL_TEST_DT is not Null and" & vbNewLine & _
            "patient_unit_number =" & moPatient.UnitNumber & vbNewLine & _
            "order by SICKLE_CELL_TEST_DT desc"
    Else
        sStmt = "select TO_CHAR(SICKLE_CELL_TEST_DT,'DD MON YYYY') as SICKLE_CELL_DT, SICKLE_CELL_TEST_RSLT" & vbNewLine & _
            "from readiness" & vbNewLine & _
            "where SICKLE_CELL_TEST_DT is not Null and" & vbNewLine & _
            "patient_unit_number =" & moPatient.UnitNumber & vbNewLine & _
            "order by SICKLE_CELL_TEST_DT desc"
    End If
    
    Set oRec1 = goDBDAS.OpenRecordset(sStmt, adOpenDynamic, adLockOptimistic)
    
    oRec.Fields("SICKLE_CELL_DT").Value = NULLStr
    oRec.Fields("SICKLE_CELL_RESULT").Value = NULLStr
    If oRec1.RecordCount > 0 Then
      If Not (CheckNULLDt(oRec1.Fields("SICKLE_CELL_DT").Value)) Then
        oRec.Fields("SICKLE_CELL_DT").Value = oRec1.Fields("SICKLE_CELL_DT").Value
        oRec.Fields("SICKLE_CELL_RESULT").Value = oRec1.Fields("SICKLE_CELL_TEST_RSLT").Value
      End If
    End If
    
    '-- Get GLASSES/GAS MASK     '--------------------------------------------------------------------------
    If gobjshared.UseSQLServer Then
        sStmt = "select distinct CONVERT(VARCHAR(12), LAST_VISUAL_ACUITY_DT, 106) as GLASS_DT," & vbNewLine & _
                "LAST_VISUAL_ACUITY_DT as Exam_DT," & vbNewLine & _
                "'OD 20/' + CORRECTED_OD + 'OS 20/' + CORRECTED_OS + 'OD 20/' + CORRECTED_OU as GLASS_CORRECTED_RX" & vbNewLine & _
            "from readiness" & vbNewLine & _
            "where LAST_VISUAL_ACUITY_DT is not Null and patient_unit_number =" & moPatient.UnitNumber & vbNewLine & _
            "union" & vbNewLine & _
            "select distinct CONVERT(VARCHAR(12), PROTECT_INSERT_EXAM_DT,106) as GLASS_DT," & vbNewLine & _
                "PROTECT_INSERT_EXAM_DT as Exam_DT," & vbNewLine & _
                "'OD 20/' + CORRECTED_OD + 'OS 20/' + CORRECTED_OS + 'OD 20/' + CORRECTED_OU as GLASS_CORRECTED_RX" & vbNewLine & _
            "from readiness" & vbNewLine & _
            "where PROTECT_INSERT_EXAM_DT is not Null and" & vbNewLine & _
            "patient_unit_number =" & moPatient.UnitNumber & vbNewLine & _
            "order by Exam_DT desc"
    Else
        sStmt = "select distinct TO_CHAR(LAST_VISUAL_ACUITY_DT,'DD MON YYYY') as GLASS_DT," & vbNewLine & _
                "LAST_VISUAL_ACUITY_DT as Exam_DT," & vbNewLine & _
                "'OD 20/' || CORRECTED_OD || 'OS 20/' || CORRECTED_OS || 'OD 20/' || CORRECTED_OU as GLASS_CORRECTED_RX" & vbNewLine & _
            "from readiness" & vbNewLine & _
            "where LAST_VISUAL_ACUITY_DT is not Null and patient_unit_number =" & moPatient.UnitNumber & vbNewLine & _
            "union" & vbNewLine & _
            "select distinct TO_CHAR(PROTECT_INSERT_EXAM_DT,'DD MON YYYY') as GLASS_DT," & vbNewLine & _
                "PROTECT_INSERT_EXAM_DT as Exam_DT," & vbNewLine & _
                "'OD 20/' || CORRECTED_OD || 'OS 20/' || CORRECTED_OS || 'OD 20/' || CORRECTED_OU as GLASS_CORRECTED_RX" & vbNewLine & _
            "from readiness" & vbNewLine & _
            "where PROTECT_INSERT_EXAM_DT is not Null and" & vbNewLine & _
            "patient_unit_number =" & moPatient.UnitNumber & vbNewLine & _
            "order by Exam_DT desc"
    End If
    Set oRec1 = goDBDAS.OpenRecordset(sStmt, adOpenDynamic, adLockOptimistic)
    
    
    If oRec1.RecordCount > 0 Then
        i = 0
        sDate = vbNullString
        oRec1.MoveFirst
        Do While Not oRec1.EOF
            If sDate <> oRec1.Fields("GLASS_DT") And Not (CheckNULLDt(oRec1.Fields("GLASS_DT").Value)) Then
                i = i + 1
                sDate = oRec1.Fields("GLASS_DT")
                oRec.Fields("GLASS_DT" & CStr(i)).Value = oRec1.Fields("GLASS_DT")
                If i = 1 Then
                    oRec.Fields("GLASS_CORRECTED_RX").Value = oRec1.Fields("GLASS_CORRECTED_RX")
                End If
            End If
            If i = 6 Then Exit Do
            oRec1.MoveNext
        Loop
    End If
    
    '-- Get DENTAL EXAM    '--------------------------------------------------------------------------
'    If gobjshared.UseSQLServer Then
'        sStmt = "select distinct CONVERT(VARCHAR(12), DENTAL_READINESS_DT, 106) as DENTAL_DT," & vbNewLine & _
'            "DENTAL_READINESS_DT as Exam_DT," & vbNewLine & _
'            "DENTAL_READINESS_CD as Dental_Class" & vbNewLine & _
'            "from readiness" & vbNewLine & _
'            "where DENTAL_READINESS_DT is not Null and" & vbNewLine & _
'            "patient_unit_number =" & moPatient.UnitNumber & vbNewLine & _
'            "order by Exam_DT desc"
'    Else
'        sStmt = "select distinct TO_CHAR(DENTAL_READINESS_DT,'DD MON YYYY') as DENTAL_DT," & vbNewLine & _
'            "DENTAL_READINESS_DT as Exam_DT," & vbNewLine & _
'            "DENTAL_READINESS_CD as Dental_Class" & vbNewLine & _
'            "from readiness" & vbNewLine & _
'            "where DENTAL_READINESS_DT is not Null and" & vbNewLine & _
'            "patient_unit_number =" & moPatient.UnitNumber & vbNewLine & _
'            "order by Exam_DT desc"
'    End If
    If gobjshared.UseSQLServer Then
        sStmt = "select distinct CONVERT(VARCHAR(12), LAST_DENTAL_EXAM_DT, 106) as DENTAL_DT," & vbNewLine & _
            "DENTAL_READINESS_DT as Exam_DT," & vbNewLine & _
            "DENTAL_READINESS_CD as Dental_Class" & vbNewLine & _
            "from readiness" & vbNewLine & _
            "where LAST_DENTAL_EXAM_DT is not Null and" & vbNewLine & _
            "patient_unit_number =" & moPatient.UnitNumber & vbNewLine & _
            "order by Exam_DT desc"
    Else
        sStmt = "select distinct TO_CHAR(LAST_DENTAL_EXAM_DT,'DD MON YYYY') as DENTAL_DT," & vbNewLine & _
            "DENTAL_READINESS_DT as Exam_DT," & vbNewLine & _
            "DENTAL_READINESS_CD as Dental_Class" & vbNewLine & _
            "from readiness" & vbNewLine & _
            "where LAST_DENTAL_EXAM_DT is not Null and" & vbNewLine & _
            "patient_unit_number =" & moPatient.UnitNumber & vbNewLine & _
            "order by Exam_DT desc"
    End If
    
    Set oRec1 = goDBDAS.OpenRecordset(sStmt, adOpenDynamic, adLockOptimistic)
    
    If oRec1.RecordCount > 0 Then
        i = 0
        sDate = NULLStr
        sClass = NULLStr
        oRec1.MoveFirst
        Do While Not oRec1.EOF
            If sDate <> oRec1.Fields("DENTAL_DT") And Not (CheckNULLDt(oRec1.Fields("DENTAL_DT").Value)) Then  'sClass <> oRec1.Fields("Dental_Class") Then
                i = i + 1
                sDate = oRec1.Fields("DENTAL_DT")
                If Not IsNull(oRec1.Fields("Dental_Class")) Then
                    sClass = oRec1.Fields("Dental_Class")
                End If
                oRec.Fields("DENTAL_DT" & CStr(i)).Value = sDate
                oRec.Fields("DENTAL_C" & CStr(i)).Value = sClass
            End If
            If i = 1 Then Exit Do
            oRec1.MoveNext
        Loop
    End If
    
    '-- Get HIV    '--------------------------------------------------------------------------
    If gobjshared.UseSQLServer Then
        sStmt = "select distinct CONVERT(VARCHAR(12), HIV_TEST_DT, 106) as HIV_DT, HIV_TEST_DT as Exam_DT" & vbNewLine & _
            "from readiness" & vbNewLine & _
            "where HIV_TEST_DT is not Null and" & vbNewLine & _
            "patient_unit_number =" & moPatient.UnitNumber & vbNewLine & _
            "order by Exam_DT desc"
    Else
        sStmt = "select distinct TO_CHAR(HIV_TEST_DT,'DD MON YYYY') as HIV_DT, HIV_TEST_DT as Exam_DT" & vbNewLine & _
            "from readiness" & vbNewLine & _
            "where HIV_TEST_DT is not Null and" & vbNewLine & _
            "patient_unit_number =" & moPatient.UnitNumber & vbNewLine & _
            "order by Exam_DT desc"
    End If
    Set oRec1 = goDBDAS.OpenRecordset(sStmt, adOpenDynamic, adLockOptimistic)
    
    
    If oRec1.RecordCount > 0 Then
        i = 0
        sDate = NULLStr
        oRec1.MoveFirst
        Do While Not oRec1.EOF
            i = i + 1
            If Not (CheckNULLDt(oRec1.Fields("HIV_DT").Value)) Then
                oRec.Fields("HIV_DT" & CStr(i)).Value = oRec1.Fields("HIV_DT")
            End If
            If i = 6 Then Exit Do
            oRec1.MoveNext
        Loop
    End If
        
    oRec.Update
        
    GetReadiness = True
    GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, cMODULE_NAME & ".GetReadiness", App.EXEName, vbExclamation
    'Resume Next
Cleanup:
    Set oRec1 = Nothing
    Set oRec2 = Nothing
    
End Function

Private Function CheckNULLDt(Dt As String) As Boolean
On Error GoTo ErrHandler
   If Dt <> "31 DEC 9999" Then
       CheckNULLDt = False
   Else
       CheckNULLDt = True
   End If
    GoTo Cleanup
ErrHandler:
    Err.Raise Err.Number, App.EXEName & "." & cMODULE_NAME & "." & "CheckNULLDt:" & Err.Source, _
       Err.Description, Err.HelpFile, Err.HelpContext
Cleanup:
    
End Function

'*****************************************************************************************
'Description: Create Record set fields
'             Clone from CreateChronicTemplate().
'Parameter: strUnitNumber,   String,
'           oRec, ADODB.Recordset
'Reutrn:    Boolean
'History:
'   Date    Developer       Description
'========== ============    ==========================================================
'09/22/01   Wilson Sun      Created
'*****************************************************************************************
Public Function CreateCounselingTemplate(oRec As ADODB.Recordset) As Boolean

    Dim i As Integer
    CreateCounselingTemplate = False
    On Error GoTo ErrHandler

    Set oRec = New ADODB.Recordset
    For i = 1 To 20
        oRec.Fields.Append "TOPIC" & CStr(i), adLongVarWChar, -1, clFieldAttr
        oRec.Fields.Append "TOPIC_DT" & CStr(i), adLongVarWChar, -1, clFieldAttr
        oRec.Fields.Append "AGE" & CStr(i), adLongVarWChar, -1, clFieldAttr
    Next
    CreateCounselingTemplate = True
    Exit Function

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, cMODULE_NAME & ".CreateCounselingTemplate", App.EXEName, vbExclamation

End Function


'
'*****************************************************************************************
'Description: Create Record set fields
'             Clone from CreateChronicTemplate().
'Parameter: strUnitNumber,   String,
'           oRec, ADODB.Recordset
'Reutrn:    Boolean
'History:
'   Date    Developer       Description
'========== ============    ==========================================================
'09/18/01   Wilson Sun      Created
'*****************************************************************************************
Private Function CreateFamilyTemplate(oRec As ADODB.Recordset) As Boolean

    CreateFamilyTemplate = False
    On Error GoTo ErrHandler

    Set oRec = New ADODB.Recordset
    oRec.Fields.Append "job_id", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "FAMILY1", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "FAMILY2", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "FAMILY3", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "FAMILY4", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "FAMILYDT1", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "FAMILYDT2", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "FAMILYDT3", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "FAMILYDT4", adLongVarWChar, -1, clFieldAttr
    
    CreateFamilyTemplate = True
    Exit Function

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, cMODULE_NAME & ".CreateFamilyTemplate", App.EXEName, vbExclamation

End Function


'*****************************************************************************************
'Description: Create Record set fields
'             Clone from CreateHospitalTemplate().
'Parameter: strUnitNumber,  String,
'           oRec,           ADODB.Recordset
'Reutrn:    Boolean
'History:
'   Date    Developer       Description
'========== ============    ==========================================================
'09/12/01   Wilson Sun      Created
'*****************************************************************************************
Private Function CreateHospitalTemplate(oRec As ADODB.Recordset) As Boolean

    CreateHospitalTemplate = False
    On Error GoTo ErrHandler

    Set oRec = New ADODB.Recordset
    oRec.Fields.Append "job_id", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "Hospital1", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "Hospital2", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "Hospital3", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "Hospital4", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "Hospital5", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "Hospital6", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "Hospital7", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "Hospital8", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "Hospital9", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "Hospital10", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "Hospital11", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "Hospital12", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "HospitalDT1", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "HospitalDT2", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "HospitalDT3", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "HospitalDT4", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "HospitalDT5", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "HospitalDT6", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "HospitalDT7", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "HospitalDT8", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "HospitalDT9", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "HospitalDT10", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "HospitalDT11", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "HospitalDT12", adLongVarWChar, -1, clFieldAttr
    
    CreateHospitalTemplate = True
    Exit Function

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, cMODULE_NAME & ".CreateHospitalTemplate", App.EXEName, vbExclamation

End Function

Private Function CreatePatientTemplate(oRec As ADODB.Recordset) As Boolean

    CreatePatientTemplate = False
    On Error GoTo ErrHandler

    Set oRec = New ADODB.Recordset
    With oRec.Fields
        .Append "job_id", adLongVarWChar, -1, clFieldAttr
        .Append "UNIT_NUMBER", adLongVarWChar, -1, clFieldAttr
        .Append "SSN", adLongVarWChar, -1, clFieldAttr
        .Append "Sex", adLongVarWChar, -1, clFieldAttr
        .Append "DOB", adLongVarWChar, -1, clFieldAttr
        .Append "DOB_FORMATTED", adLongVarWChar, -1, clFieldAttr
        .Append "Age", adLongVarWChar, -1, clFieldAttr
        .Append "MARITAL_STATUS", adLongVarWChar, -1, clFieldAttr
        .Append "COMMAND_INTEREST", adLongVarWChar, -1, clFieldAttr
        .Append "COMMAND_SECURITY", adLongVarWChar, -1, clFieldAttr
        .Append "ALTERNATIVE_CARE", adLongVarWChar, -1, clFieldAttr
        .Append "FLYING_STATUS", adLongVarWChar, -1, clFieldAttr
        .Append "JUMPING_STATUS", adLongVarWChar, -1, clFieldAttr
        .Append "DIVING_STATUS", adLongVarWChar, -1, clFieldAttr
        .Append "ON_MOBILITY_STATUS", adLongVarWChar, -1, clFieldAttr
        .Append "SUBMARINE_STATUS", adLongVarWChar, -1, clFieldAttr
        .Append "PATIENT_CATEGORY", adLongVarWChar, -1, clFieldAttr
        .Append "UIC", adLongVarWChar, -1, clFieldAttr
        .Append "SPONSOR_SSN", adLongVarWChar, -1, clFieldAttr
        .Append "FMP", adLongVarWChar, -1, clFieldAttr
        .Append "VIP", adLongVarWChar, -1, clFieldAttr
        .Append "PCM_ID", adLongVarWChar, -1, clFieldAttr
        .Append "HOME_PHONE", adLongVarWChar, -1, clFieldAttr
        .Append "WORK_PHONE", adLongVarWChar, -1, clFieldAttr
        .Append "SERVICE", adLongVarWChar, -1, clFieldAttr
        .Append "Rank", adLongVarWChar, -1, clFieldAttr
        .Append "RECORDS_MAINTAINED_AT", adLongVarWChar, -1, clFieldAttr
        .Append "RELATIONSHIP_TO_SPONSOR", adLongVarWChar, -1, clFieldAttr
        .Append "FIRST_NAME", adLongVarWChar, -1, clFieldAttr
        .Append "MIDDLE_NAME", adLongVarWChar, -1, clFieldAttr
        .Append "LAST_NAME", adLongVarWChar, -1, clFieldAttr
        .Append "FULL_NAME", adLongVarWChar, -1, clFieldAttr
        .Append "Religion", adLongVarWChar, -1, clFieldAttr
        .Append "Race", adLongVarWChar, -1, clFieldAttr
        .Append "ACTIVE_DUTY", adLongVarWChar, -1, clFieldAttr
    End With
    
    CreatePatientTemplate = True
    Exit Function

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, cMODULE_NAME & ".CreatePatientTemplate", App.EXEName, vbExclamation

End Function

Private Function CreateImmunizationTemplate(oRec As ADODB.Recordset) As Boolean

    CreateImmunizationTemplate = False
    On Error GoTo ErrHandler
    
    Set oRec = New ADODB.Recordset
    oRec.Fields.Append "job_id", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "HEPA1_DT", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "HEPA2_DT", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "HEPB1_DT", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "HEPB2_DT", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "HEPB3_DT", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "MMR1_DT", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "MMR2_DT", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "PNEUM_DT", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "POLIO_OPV", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "TD_LAST_DT", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "TD_DUE_DT", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "YELLOW_DT", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "YELLOW_DUE_DT", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "TYPHOID_DT1", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "TYPHOID_C1", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "TYPHOID_DT2", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "TYPHOID_C2", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "TYPHOID_DT3", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "TYPHOID_C3", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "TYPHOID_DT4", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "TYPHOID_C4", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "TYPHOID_DT5", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "TYPHOID_C5", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "TYPHOID_DT6", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "TYPHOID_C6", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "ANTHRAX_1", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "ANTHRAX_2", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "ANTHRAX_3", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "ANTHRAX_4", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "ANTHRAX_5", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "ANTHRAX_6", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "PPD_MM1", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "PPD_DT1", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "PPD_MM2", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "PPD_DT2", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "PPD_MM3", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "PPD_DT3", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "PPD_MM4", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "PPD_DT4", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "PPD_MM5", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "PPD_DT5", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "PPD_MM6", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "PPD_DT6", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "PPD_MM7", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "PPD_DT7", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "INFLUENZA_DT1", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "INFLUENZA_DT2", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "INFLUENZA_DT3", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "INFLUENZA_DT4", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "INFLUENZA_DT5", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "INFLUENZA_DT6", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "INFLUENZA_DT7", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "VARICELLA_DT1", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "VARICELLA_DT2", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "MENINGO_DT1", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "MENINGO_DT2", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "ADENO_DT1", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "ADENO_DT2", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "JAPANESE_DT1", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "JAPANESE_DT2", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "JAPANESE_DT3", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "JAPANESE_DT4", adLongVarWChar, -1, clFieldAttr

    CreateImmunizationTemplate = True
    Exit Function

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, cMODULE_NAME & ".CreateImmunizationTemplate", App.EXEName, vbExclamation

End Function


Private Function CreateScreeningTemplate(oRec As ADODB.Recordset) As Boolean

    CreateScreeningTemplate = False
    On Error GoTo ErrHandler

    Set oRec = New ADODB.Recordset
    oRec.Fields.Append "job_id", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "YEAR", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "AGE", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "HEAR", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "WEIGHT", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "HEIGHT", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "BP", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "CHOLESTEROL", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "HEARING", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "SKIN", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "DENTAL", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "VISION", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "BREAST", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "MAMMOGRAM", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "PAP", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "FECAL", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "SIGMOID", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "COLONOSCOPY", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "TESTICULAR", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "PROSTATE", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "RUBELLA", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "OCCUPATIONAL", adLongVarWChar, -1, clFieldAttr
    
    CreateScreeningTemplate = True
    Exit Function

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, cMODULE_NAME & ".CreateScreeningTemplate", App.EXEName, vbExclamation

End Function

Private Function CreateReadinessTemplate(oRec As ADODB.Recordset) As Boolean

    CreateReadinessTemplate = False
    On Error GoTo ErrHandler

    Set oRec = New ADODB.Recordset
    oRec.Fields.Append "job_id", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "DNA_DT", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "BLOOD_TYPE_DT", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "BLOOD_TYPE_RESULT", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "G6PD_DT", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "G6PD_RESULT", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "SICKLE_CELL_DT", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "SICKLE_CELL_RESULT", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "Profile_Date", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "Profile_P", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "Profile_U", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "Profile_L", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "Profile_H", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "Profile_E", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "Profile_S", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "GLASS_DT1", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "GLASS_DT2", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "GLASS_DT3", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "GLASS_DT4", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "GLASS_DT5", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "GLASS_DT6", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "GLASS_CORRECTED_RX", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "DENTAL_DT1", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "DENTAL_C1", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "DENTAL_DT2", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "DENTAL_C2", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "DENTAL_DT3", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "DENTAL_C3", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "DENTAL_DT4", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "DENTAL_C4", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "DENTAL_DT5", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "DENTAL_C5", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "DENTAL_DT6", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "DENTAL_C6", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "HIV_DT1", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "HIV_DT2", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "HIV_DT3", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "HIV_DT4", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "HIV_DT5", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "HIV_DT6", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "FIT_DT1", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "FIT_C1", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "FIT_DT2", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "FIT_C2", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "FIT_DT3", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "FIT_C3", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "FIT_DT4", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "FIT_C4", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "FIT_DT5", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "FIT_C5", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "FIT_DT6", adLongVarWChar, -1, clFieldAttr
    oRec.Fields.Append "FIT_C6", adLongVarWChar, -1, clFieldAttr
    
    CreateReadinessTemplate = True
    Exit Function

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, cMODULE_NAME & ".CreateReadinessTemplate", App.EXEName, vbExclamation

End Function



'*****************************************************************************************
'Description: Main drive for building report.
'Parameter: sPatientUnitNumber, String,
'           oReport,            CRAXDRT.Report
'Reutrn:    String
'History:
'   Date    Developer       Description
'========== ============    ==========================================================
'09/24/01   Wilson Sun      Added Counseling function
'09/17/01   Wilson Sun      Added Family History function
'09/12/01   Wilson Sun      Added Historical Procecure function
'??/??/??   Brain Webb      Created
'*****************************************************************************************
Public Function BuildDD2766(sPatientUnitNumber As String, _
                            oReport As CRAXDRT.Report) As Boolean

    On Error GoTo ErrHandler
    BuildDD2766 = False
    Screen.MousePointer = vbHourglass
    
    Dim oCrystal As Crystal
    Dim oRec As ADODB.Recordset
    Dim oRecSub As ADODB.Recordset
    Dim oProblemOps As Object
    Dim lLoop As Long
    
    DoEvents

    '-- Get Data
    '----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection
    
    ' Get Patient
    Set moPatient = LoadPatientObject(sPatientUnitNumber)
    Set oProblemOps = LoadProblemOps
    
    '-- Get Counseling
    '----------------------------------------------------------
    cwobjCWComm.setStatusText "Getting counseling data..."
    If Not GetCounseling(oRec) Then GoTo Cleanup
    moCollRecSet.Add oRec, "counseling"
    
    '-- Get Family History
    '----------------------------------------------------------
    cwobjCWComm.setStatusText "Getting family history data..."
    If Not GetFamilyHistory(oProblemOps, oRec, oRecSub) Then GoTo Cleanup
    moCollRecSet.Add oRec, "familyhistory"
    moCollRecSet.Add oRecSub, "sub_familyhistory"

    '-- Get hospitalizations/surgeries
    '----------------------------------------------------------
    cwobjCWComm.setStatusText "Getting hospitalizations data..."
    If Not GetHospitalizations(oProblemOps, oRec, oRecSub) Then GoTo Cleanup
    moCollRecSet.Add oRec, "hospitalizations"
    moCollRecSet.Add oRecSub, "sub_hospitalizations"

    '-- Get Chronic Illnesses
    '----------------------------------------------------------
    cwobjCWComm.setStatusText "Getting chronic illness data..."
    If Not GetChronicIllness(oProblemOps, oRecSub) Then GoTo Cleanup
    'moCollRecSet.Add oRec, "chronic"
    moCollRecSet.Add oRecSub, "sub_chronic"

    '-- Get Medications
    '------------------------------------------------------------
    cwobjCWComm.setStatusText "Getting medications data..."
    If Not GetMedications2(oRec, oRecSub) Then GoTo Cleanup
    moCollRecSet.Add oRec, "medications"
    moCollRecSet.Add oRecSub, "sub_medications"

    '-- Get Allergy
    '----------------------------------------------------------
    'Wilson 9/26/01 added oRecSub parameter
    cwobjCWComm.setStatusText "Getting allergy data..."
    If Not GetAllergy(oRec, oRecSub) Then GoTo Cleanup
    moCollRecSet.Add oRec, "allergy"
    moCollRecSet.Add oRecSub, "sub_allergy"

    '-- Get Screening Exams
    '----------------------------------------------------------
    cwobjCWComm.setStatusText "Getting screening exam data..."
    If Not GetScreeningExams(moCollRecSet) Then GoTo Cleanup
    
    '-- Get Readiness
    '----------------------------------------------------------
    cwobjCWComm.setStatusText "Getting readiness data..."
    If Not GetReadiness(oRec) Then GoTo Cleanup
    moCollRecSet.Add oRec, "readiness"

    '-- Immunization
    cwobjCWComm.setStatusText "Getting immunizations data..."
    If Not GetImmunizations(oRec, oRecSub) Then GoTo Cleanup
    moCollRecSet.Add oRec, "immunizations"
    moCollRecSet.Add oRecSub, "sub_immunizations"
    
    '-- Patient Demographics
    '----------------------------------------------------------
    cwobjCWComm.setStatusText "Getting patient demographics data..."
   
    DoEvents
    
    If Not GetPatientData(oRec) Then
        GoTo Cleanup
    End If
    
    moCollRecSet.Add oRec, "demographics"

    '-- Sponsor Demographics    '----------------------------------------------------------
    cwobjCWComm.setStatusText "Getting sponsor demographics data..."
'SCR 16688 - 3/13/02 - Walker - SponsorSSN will only have a value when patient is a dependent of sponsor
'                               Otherwise, this will cause an Error 91 to be generated for the patient.
'SCR 16688 - 3/13/02    If moPatient.Sponsor.UnitNumber = moPatient.UnitNumber Then
    If (moPatient.SponsorSSN = moPatient.SSN) Then  'SCR 16688 - 3/13/02
        DuplicateDefinition oRec, oRecSub, True
    Else
        If Not GetSponsorPatientData(oRecSub) Then
            GoTo Cleanup
        End If
    End If
    moCollRecSet.Add oRecSub, "sponsor"

    '-- Set Formulas '-----------------------------------------------------------------------------
    Set moCollFormulas = Nothing
    Set moCollFormulas = New Collection
    moCollFormulas.Add "'" & Format(Now, gsDefaultFormat & "  HHMM") & "'", "TodaysDateTime"

    '-- Create Report Object '-----------------------------------------------------------------------------
    cwobjCWComm.setStatusText "Opening DD2766..."
    Set oCrystal = New Crystal
    Set oReport = Nothing
    If Not oCrystal.ReportOpen("dd2766.rpt", oReport) Then GoTo Cleanup
    If Not oCrystal.AssignRecordSets(moCollRecSet, oReport) Then GoTo Cleanup
    If Not oCrystal.AssignFormulas(moCollFormulas, oReport) Then GoTo Cleanup

    BuildDD2766 = True

    GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, cMODULE_NAME & ".Build2766", App.EXEName, vbExclamation
    'Resume Next
Cleanup:
    Set oRec = Nothing
    Set oRecSub = Nothing
    Set moPatient = Nothing
    Set oProblemOps = Nothing
    
    If Not moCollRecSet Is Nothing Then
        For lLoop = moCollRecSet.Count To 1 Step -1
              moCollRecSet.Remove lLoop
        Next
        Set moCollRecSet = Nothing
    End If
    
    If Not moCollFormulas Is Nothing Then
        For lLoop = moCollFormulas.Count To 1 Step -1
              moCollFormulas.Remove lLoop
        Next
        Set moCollFormulas = Nothing
    End If
    
    cwobjCWComm.setStatusText vbNullString
    Screen.MousePointer = vbDefault
        
End Function

Private Function GetRelationshipToSponsor(sFMP As String) As String

    On Error GoTo ErrHandler
    GetRelationshipToSponsor = vbNullString
    
    Dim sStmt  As String
    
    sStmt = "select T1.DESCRIPTION as RELATIONSHIP" & vbNewLine & _
            "from NCIDS T1, NCIDS T2" & vbNewLine & _
            "where T1.NCID = T2.NCID" & vbNewLine & _
            "and T1.DOMAINNCID = 110627 " & vbNewLine & _
            "and T2.DESCRIPTION = '" & sFMP & "'"
    
    goDBSqlOps.Execute sStmt
    If goDBSqlOps.EOF = True Then
        Exit Function
    End If
    
    GetRelationshipToSponsor = goDBSqlOps.Value("Relationship")
    
    Exit Function

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, cMODULE_NAME & ".GetRelationshipToSponsor", App.EXEName, vbExclamation
    Exit Function

    Err.Clear
    Resume
    
End Function



Private Sub GetPatientName(sFirst As String, sMiddle As String, sLast As String)

    On Error GoTo ErrHandler
    
    Dim sStmt  As String
    
    sStmt = "select first_name, middle_name, last_name " & vbCrLf & _
            "From mmi_name " & _
            "Where unit_number = " & moPatient.UnitNumber & " and NAME_TYPE_CID = 187"
    goDBSqlOps.Execute sStmt
    If goDBSqlOps.EOF = True Then
        Exit Sub
    End If
    
    sFirst = goDBSqlOps.Value("First_Name")
    sMiddle = goDBSqlOps.Value("Middle_Name")
    sLast = goDBSqlOps.Value("Last_Name")
    
    Exit Sub

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, cMODULE_NAME & ".GetPatientName", App.EXEName, vbExclamation

End Sub

Private Function LoadPatientObject(sUnitNumber As String) As Patient
    
    Dim strSQL As String
    Dim oRS As ADODB.Recordset
    Dim oPatient As Patient
    
    On Error GoTo ErrHandler
    
    Set oRS = New ADODB.Recordset
    Set oPatient = New Patient
    
    'Get MEDICAL_RECORD_NUMBER
    strSQL = "select MEDICAL_RECORD_NUMBER from MMI_SITE where UNIT_NUMBER = " & sUnitNumber
    Set oRS = goDBDAS.OpenRecordset(strSQL)
'SCR 16688 - 3/13/02 - Walker - Check for NULL in MEDICAL_RECORD_NUMBER
'SCR 32383
    If oRS.Fields(0).Name = "MEDICAL_RECORD_NUMBER" Then
        mMedicalRecordNumber = oRS.Fields!MEDICAL_RECORD_NUMBER
    End If
    oRS.Close
        
'SCR 16688 - 3/13/02 - Walker - This prevents required data from being retrieved
'which is then referenced in code which fails when accessing the UnitNumber field.
    Call oPatient.LoadPatient(sUnitNumber, mMedicalRecordNumber)
    Set LoadPatientObject = oPatient
    GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, cMODULE_NAME & ".LoadPatientObject", App.EXEName, vbExclamation
'Resume Next
Cleanup:
    Set oRS = Nothing
    Set oPatient = Nothing
    
End Function

Private Function LoadProblemOps() As Object
On Error GoTo ErrHandler
    Dim objProblem As Object
    Set objProblem = cwobjCWComm.InitializeOLEServer(NCID_PROBLEMS_APPLICATION)
    Set LoadProblemOps = objProblem.ProblemOps(moPatient)
    GoTo Cleanup
ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, cMODULE_NAME & ".LoadProblemOps", App.EXEName, vbExclamation
Cleanup:

End Function


'*****************************************************************************************
'Description: SetNextImmSeries2()
'       Set related values
'Parameter: iSeries,    - Series number
'           sDestField, - report field
'           oRecDest,   - main report
'           oRecSource  - sub report
'Used for:  "HEPA", "HEPB", "MMR", "ANTHRAX_", "VARICELLA_DT", "JAPANESE_DT"
'Reutrn:    Boolean
'History:
'   Date    Developer       Description
'========== ============    ==========================================================
'11/04/01   Wilson Sun      Created.
'*****************************************************************************************
Private Function SetNextImmSeries2(ByVal iSeriesStart As Integer, _
                                  ByVal iSeriesEnd As Integer, _
                                  ByVal sType As String, _
                                  ByRef oRecDest As ADODB.Recordset, _
                                  ByRef oRecSource As ADODB.Recordset) As Boolean
    Dim i As Integer
    Dim sDestField As String
    On Error GoTo ErrHandler
    SetNextImmSeries2 = False
        
    If oRecSource.RecordCount > 0 Then
        oRecSource.MoveFirst
        For i = iSeriesStart To oRecSource.RecordCount
            'Series should not be zero.
            If oRecSource.Fields("Vacc_Series").Value = 0 Then
                oRecSource.Fields("Vacc_Series").Value = 1
            End If
            
            If i = oRecSource.Fields("Vacc_Series").Value Then
                Select Case sType
                    Case "HEPA", "HEPB", "MMR"
                        sDestField = sType & CStr(i) & "_DT"
                    Case "ANTHRAX_", "VARICELLA_DT", "JAPANESE_DT"
                        sDestField = sType & CStr(i)
                End Select
                oRecDest.Fields(sDestField).Value = oRecSource.Fields("Imm_Date").Value
            End If
            If i = iSeriesEnd Then Exit For
            If Not oRecSource.EOF Then oRecSource.MoveNext
        Next
    End If
    SetNextImmSeries2 = True
    Exit Function
ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, cMODULE_NAME & ".SetNextImmSeries2 in " & sType, App.EXEName, vbExclamation
End Function

'*****************************************************************************************
'Description: SetOneShotOnly() to display most recently date shot and due date
'Parameter: sField,
'           oRecDest,   - main report
'           oRecSource,  - sub report
'           bDueDate,
'           sDueDateField
'Used for:  "PNEUM_DT","POLIO_OPV",TD_LAST_DT,"YELLOW_DT"
'Reutrn:    Boolean
'History:
'   Date    Developer       Description
'========== ============    ==========================================================
'11/09/01   Wilson Sun      Created.
'*****************************************************************************************
Private Function SetOneShotOnly(ByVal sField As String, _
                                ByRef oRec As ADODB.Recordset, _
                                ByRef oRecSub As ADODB.Recordset, _
                                Optional ByVal sDueDateField As String = vbNullString) As Boolean
    Dim sClass As String
    
    On Error GoTo ErrHandler
    SetOneShotOnly = False

    If oRecSub.RecordCount > 0 Then
        oRecSub.Sort = "VACC_DATE desc"
        oRecSub.MoveFirst
        Select Case oRecSub.Fields("id").Value
            Case 2 'POLIO OPV
                sClass = "O" & vbCrLf
            Case 10 'POLIO IPV
                sClass = "I" & vbCrLf
        End Select
        oRec.Fields(sField).Value = sClass & oRecSub.Fields("Imm_Date").Value
        If sDueDateField > vbNullString Then
            oRec.Fields(sDueDateField).Value = UCase(Format(oRecSub.Fields("Vacc_NextDue").Value, "dd mmm yyyy"))
        End If
    End If
       
    SetOneShotOnly = True
    Exit Function
ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, cMODULE_NAME & ".SetOneShotOnly in " & sField, App.EXEName, vbExclamation
End Function


'*****************************************************************************************
'Description: SetMoreShots() to display most recently dates and set (mm or class)
'Parameter: sField,
'           oRecDest,   - main report
'           oRecSource,  - sub report
'           iHowMany
'           sDueDateField
'Used for:  "TYPHOID_DT", "PPD_DT"
'Reutrn:    Boolean
'History:
'   Date    Developer       Description
'========== ============    ==========================================================
'11/15/01   Wilson Sun      Changed Vacc_Dosage to RXN_Size to assign to PPD_MM field
'11/09/01   Wilson Sun      Created.
'*****************************************************************************************
Private Function SetMoreShots(ByVal sField As String, _
                                ByRef oRec As ADODB.Recordset, _
                                ByRef oRecSub As ADODB.Recordset, _
                                ByVal iHowMany As Integer) As Boolean
    Dim i As Integer
    Dim sClass As String    'For TYPHOID only
    Dim bMM As Boolean      'For IPPD only
    
    On Error GoTo ErrHandler
    SetMoreShots = False
    If oRecSub.RecordCount >= 0 Then
        If oRecSub.RecordCount <= iHowMany Then
            oRecSub.Sort = "VACC_DATE asc"
        Else
            oRecSub.Sort = "VACC_DATE desc"
        End If
    
        For i = 1 To oRecSub.RecordCount
            '-- Get Class
            Select Case oRecSub.Fields("id").Value
                Case 25: sClass = "0"
                Case 101: sClass = "1"
                Case 41, 53: sClass = "2"
                Case 96: bMM = True
            End Select
            
            If oRecSub.RecordCount <= iHowMany Then
                oRec.Fields(sField & CStr(i)).Value = oRecSub.Fields("Imm_Date").Value
                If sClass <> vbNullString Then oRec.Fields("TYPHOID_C" & CStr(i)).Value = sClass
                If bMM Then oRec.Fields("PPD_MM" & CStr(i)).Value = oRecSub.Fields("RXN_Size").Value
            Else
                oRec.Fields(sField & CStr(iHowMany + 1 - i)).Value = oRecSub.Fields("Imm_Date").Value
                If sClass <> vbNullString Then oRec.Fields("TYPHOID_C" & CStr(iHowMany + 1 - i)).Value = sClass
                If bMM Then oRec.Fields("PPD_MM" & CStr(iHowMany + 1 - i)).Value = oRecSub.Fields("RXN_Size").Value
            End If
            oRecSub.MoveNext
            If i = iHowMany Then Exit For
        Next
    End If
 
       
    SetMoreShots = True
    Exit Function
ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, cMODULE_NAME & ".SetMoreShots in " & sField, App.EXEName, vbExclamation
End Function


Private Sub Class_Initialize()
On Error GoTo ErrHandler
    
    Set moPatientDemographics = New PatientDemographics

    GoTo Cleanup
ErrHandler:
    Err.Raise Err.Number, App.EXEName & "." & cMODULE_NAME & "." & "Class_Initialize:" & Err.Source, _
       Err.Description, Err.HelpFile, Err.HelpContext
Cleanup:
    
End Sub


Private Sub Class_Terminate()
On Error GoTo ErrHandler
    
    Set moCollRecSet = Nothing
    Set moCollFormulas = Nothing

    Set moPatientDemographics = Nothing

    GoTo Cleanup
ErrHandler:
    Err.Raise Err.Number, App.EXEName & "." & cMODULE_NAME & "." & "Class_Terminate:" & Err.Source, _
       Err.Description, Err.HelpFile, Err.HelpContext
Cleanup:
End Sub


'*****************************************************************************************
'Description:
'   This method to get Screeing Exams block 6, 8, 9 result
'   HEARING, ORAL/DENTAL, and EYE/VISION
'History
'   Date        Developer       Description
'   =======     ===========     ========================================================
'   11/16/01    Wilson Sun      Changed Result value from 'Y' to Formated date
'                               Block 9 EYE/VISION is deffered. Field invisibal in DD2766.rpt
'   08/18/01    Wilson Sun      Build Sql for Screeng Exams block 6, 8, and 9
'*****************************************************************************************
Private Function GetBlock689SQL(ByVal strCategory As String, _
                              ByVal strColumnName As String) As String
On Error GoTo ErrHandler
    '-- select
    If gobjshared.UseSQLServer Then
        GetBlock689SQL = "SELECT DISTINCT '" & strCategory & "' as CATEGORY," & vbNewLine & _
            "YEAR(" & strColumnName & ") as Result_Year, " & vbNewLine & _
            "CONVERT(VARCHAR(12), " & strColumnName & ", 106) as Result, " & vbNewLine & _
            "CONVERT(VARCHAR(12), " & strColumnName & ", 106) as Date_Format, " & vbNewLine & _
            strColumnName & " as Event_Date " & vbNewLine
    Else
        GetBlock689SQL = "SELECT DISTINCT '" & strCategory & "' as CATEGORY," & vbNewLine & _
            "To_Number(TO_CHAR(" & strColumnName & ",'YYYY')) as Result_Year, " & vbNewLine & _
            "TO_CHAR(" & strColumnName & ",'DD MON YYYY') as Result, " & vbNewLine & _
            "TO_CHAR(" & strColumnName & ",'DD MON YYYY') as Date_Format, " & vbNewLine & _
            strColumnName & " as Event_Date " & vbNewLine
    End If
    
    '-- from
    GetBlock689SQL = GetBlock689SQL & _
        "FROM READINESS " & vbNewLine
        
    '-- where
    GetBlock689SQL = GetBlock689SQL & _
        "WHERE " & strColumnName & " IS NOT NULL " & vbNewLine & _
        "AND   PATIENT_UNIT_NUMBER = " & moPatient.UnitNumber & vbNewLine & _
        "ORDER BY Event_Date DESC"
    GoTo Cleanup
ErrHandler:
    Err.Raise Err.Number, App.EXEName & "." & cMODULE_NAME & "." & "GetBlock689SQL:" & Err.Source, _
       Err.Description, Err.HelpFile, Err.HelpContext
Cleanup:
End Function

'*****************************************************************************************
'Description:
'   This method to get Lab results for:
'   CHOLESTEROL     (6097)
'   PAP             (20758)
'   OCCULT BLOOD    (6339)
'   RUBELLA SCREEN  (22394)
'History
'   Date        Developer       Description
'   =======     ===========     ========================================================
'   11/16/01    Wilson Sun      For RUBELLA SCREEN (22394), display formated date as result
'   08/18/01    Wilson Sun      Build Sql for lab results
'*****************************************************************************************
Private Function GetLabSQL(ByVal strCategory As String, _
                              ByVal lObservationNcid As Long) As String
On Error GoTo ErrHandler
    Dim sResult As String
    
    If lObservationNcid = 22394 Then
        sResult = "TO_CHAR(e.EVENT_START_GMTIME,'DD MON YYYY')"
    Else
        sResult = "o.TEXT_VALUE"
    End If
    
    '-- select
    GetLabSQL = "SELECT /*+ ORDERED USE_NL(o) USE_NL(oc) INDEX(o OBSERVATION_X1) */ '" & strCategory & "' as CATEGORY," & vbNewLine & _
        "To_Number(TO_CHAR(e.EVENT_START_GMTIME,'YYYY')) as Result_Year, " & vbNewLine & _
        sResult & " as Result, " & vbNewLine & _
        "TO_CHAR(e.EVENT_START_GMTIME,'DD MON YYYY') as Date_Format, " & vbNewLine & _
        "e.EVENT_START_GMTIME as Event_Date " & vbNewLine
    
    '-- from
    GetLabSQL = GetLabSQL & _
        "FROM LAB_EVENT_TYPE e, OBSERVATION o, OBSERVATION_COMMENT oc " & vbNewLine
        
    '-- where
    GetLabSQL = GetLabSQL & _
        "WHERE substr(oc.text_comment,1,13) = 'DATE_RESULTED' " & vbNewLine & _
        "AND   oc.observation_id = o.observation_id " & vbNewLine & _
        "AND   oc.event_id = o.event_id" & vbNewLine
    
    GetLabSQL = GetLabSQL & "AND  observation_ncid = " & CStr(lObservationNcid) & vbNewLine
    GetLabSQL = GetLabSQL & "AND   o.event_id = e.event_id " & vbNewLine & _
        "AND   e.event_type_ncid = '485' " & vbNewLine & _
        "AND   e.unit_number = " & moPatient.UnitNumber
        
    '-- order by
    GetLabSQL = GetLabSQL & vbCrLf & " ORDER BY e.EVENT_START_GMTIME desc "
    
    GoTo Cleanup
ErrHandler:
    Err.Raise Err.Number, App.EXEName & "." & cMODULE_NAME & "." & "GetLabSQL:" & Err.Source, _
       Err.Description, Err.HelpFile, Err.HelpContext
Cleanup:
    
End Function
'*****************************************************************************************
'Description:
'   This method to get Lab results for:
'   CHOLESTEROL     (6097)
'   PAP             (20758)
'   OCCULT BLOOD    (6339)
'   RUBELLA SCREEN  (22394)
'History
'   Date        Developer       Description
'   =======     ===========     ========================================================
'   11/16/01    Wilson Sun      For RUBELLA SCREEN (22394), display formated date as result
'   08/18/01    Wilson Sun      Build Sql for lab results
'*****************************************************************************************
Private Function GetLabSQL2(ByVal strCategory As String, _
                              Optional sTestType As String) As String
On Error GoTo ErrHandler
    
    Dim sResult As String
    
    sResult = "o.TEXT_VALUE"
    
    '-- select
    GetLabSQL2 = "SELECT /*+ ORDERED USE_NL(oc) INDEX(o OBSERVATION_X1) */ '" & strCategory & "' as CATEGORY," & vbNewLine & _
        "To_Number(TO_CHAR(e.EVENT_START_GMTIME,'YYYY')) as Result_Year, " & vbNewLine & _
        sResult & " as Result, " & vbNewLine & _
        "TO_CHAR(e.EVENT_START_GMTIME,'DD MON YYYY') as Date_Format, " & vbNewLine & _
        "e.EVENT_START_GMTIME as Event_Date " & vbNewLine
    
    '-- from
    GetLabSQL2 = GetLabSQL2 & _
        "FROM LAB_EVENT_TYPE e, OBSERVATION o, OBSERVATION_COMMENT oc " & vbNewLine
        
    '-- where
    GetLabSQL2 = GetLabSQL2 & _
        "WHERE substr(oc.text_comment,1,13) = 'DATE_RESULTED' " & vbNewLine & _
        "AND   oc.observation_id = o.observation_id " & vbNewLine & _
        "AND   oc.event_id = o.event_id" & vbNewLine
    
    GetLabSQL2 = GetLabSQL2 & "AND  observation_ncid IN " '& CStr(lObservationNcid) & vbNewLine
    GetLabSQL2 = GetLabSQL2 & "(SELECT TEST_NCID FROM DD2766_TEST_LOOKUP " & vbCrLf
    GetLabSQL2 = GetLabSQL2 & "WHERE TEST_IDENTIFIER = " & goDBConn.SQLQuote(sTestType) & " AND MICRO = 0)" & vbCrLf
    
    GetLabSQL2 = GetLabSQL2 & "AND   o.event_id = e.event_id " & vbNewLine & _
        "AND   e.event_type_ncid = '485' " & vbNewLine & _
        "AND   e.unit_number = " & moPatient.UnitNumber
        
    '-- order by
    GetLabSQL2 = GetLabSQL2 & "UNION" & vbCrLf
    GetLabSQL2 = GetLabSQL2 & "SELECT /*+ ORDERED RULE no_index(e event_x4) */ '" & strCategory & "' as CATEGORY," & vbNewLine & _
        "To_Number(TO_CHAR(e.EVENT_START_GMTIME,'YYYY')) as Result_Year, " & vbNewLine & _
        "ci.item_text as Result, " & vbNewLine & _
        "TO_CHAR(e.EVENT_START_GMTIME,'DD MON YYYY') as Date_Format, " & vbNewLine & _
        "e.EVENT_START_GMTIME as Event_Date " & vbNewLine
    GetLabSQL2 = GetLabSQL2 & "FROM EVENT e, COMMENT_ITEM ci" & vbCrLf
    GetLabSQL2 = GetLabSQL2 & "Where ci.group_ordinal = 7 and ci.event_id = e.event_id and" & vbCrLf
    GetLabSQL2 = GetLabSQL2 & "e.event_subtype_ncid IN (SELECT TEST_NCID FROM DD2766_TEST_LOOKUP " & vbCrLf
    'FACILITY_NCID = " & cwobjLogonEnviron.FacilityNCID & " AND
    GetLabSQL2 = GetLabSQL2 & "WHERE TEST_IDENTIFIER = " & goDBConn.SQLQuote(sTestType) & " AND MICRO = 1)" & vbCrLf
    GetLabSQL2 = GetLabSQL2 & "AND e.event_type_ncid = '75412' AND e.unit_number = " & moPatient.UnitNumber
    GetLabSQL2 = GetLabSQL2 & vbCrLf & " ORDER BY EVENT_DATE desc "
    
GoTo Cleanup
ErrHandler:
    Err.Raise Err.Number, App.EXEName & "." & cMODULE_NAME & "." & "GetLabSQL2:" & Err.Source, _
       Err.Description, Err.HelpFile, Err.HelpContext
Cleanup:
    
End Function

'*****************************************************************************************
'Description:
'   This method to build SQL statement for NEW_RESULTS.CRITICAL_ABNORMAL_VALUE
'Note:
'   if DOCUMENT and NEW_RESULTS tables' data not sync. it will not work
'History
'   Date        Developer       Description
'   =======     ===========     ========================================================
'   08/20/01    Wilson Sun      Build Sql for NEW_RESULTS results
'*****************************************************************************************
Private Function GetNewResultsSQL(ByVal strUnitNumber As String) As String
On Error GoTo ErrHandler
    GetNewResultsSQL = "SELECT CRITICAL_ABNORMAL_VALUE  as RESULT" & vbNewLine & _
            "FROM NEW_RESULTS" & vbNewLine & _
            "WHERE RESULT_TYPE = 'RAD'" & vbNewLine & _
            "AND PATIENT_UNIT_NUMBER = '" & strUnitNumber & "'" & vbNewLine & _
            "ORDER BY ORDER_NUMBER DESC"
GoTo Cleanup
ErrHandler:
    Err.Raise Err.Number, App.EXEName & "." & cMODULE_NAME & "." & "GetNewResultsSQL:" & Err.Source, _
       Err.Description, Err.HelpFile, Err.HelpContext
Cleanup:
    
End Function

'*****************************************************************************************
'Description:
'   This method to get Radiology results for:
'   MAMMOGRAM     (125339)
'   TESTICULAR    (126033)
'History
'   Date        Developer       Description
'   =======     ===========     ========================================================
'   08/20/01    Wilson Sun      Build Sql for Radiology results
'*****************************************************************************************
Private Function GetRadSQL(ByVal strCategory As String, _
                              ByVal lObservationNcid As Long, _
                              Optional sTestType As String = vbNullString) As String
On Error GoTo ErrHandler
    '-- select
    GetRadSQL = "SELECT /*+ ORDERED RULE no_index(e event_x4) */ '" & strCategory & "' as CATEGORY," & vbNewLine & _
        "To_Number(TO_CHAR(E.EVENT_START_GMTIME,'YYYY')) as Result_Year, " & vbNewLine & _
        "'     ' as Result, " & vbNewLine & _
        "TO_CHAR(e.EVENT_START_GMTIME,'DD MON YYYY') as Date_Format, " & vbNewLine & _
        "e.EVENT_START_GMTIME as Event_Date " & vbNewLine
    
    '-- from
    GetRadSQL = GetRadSQL & _
        "FROM EVENT E, DOCUMENT D " & vbNewLine
    
    '-- where
    GetRadSQL = GetRadSQL & "WHERE D.EVENT_ID = E.EVENT_ID " & vbNewLine
    If Len(Trim(sTestType)) = 0 Then
        GetRadSQL = GetRadSQL & "AND   REPORT_TYPE_NCID = " & CStr(lObservationNcid) & vbNewLine
    Else
        GetRadSQL = GetRadSQL & "AND REPORT_TYPE_NCID IN (SELECT TEST_NCID FROM DD2766_TEST_LOOKUP " & vbNewLine
        'FACILITY_NCID = " & cwobjLogonEnviron.FacilityNCID & " AND
        GetRadSQL = GetRadSQL & "WHERE  TEST_IDENTIFIER = " & goDBConn.SQLQuote(sTestType) & ")" & vbCrLf
    End If
    
    GetRadSQL = GetRadSQL & "AND   E.UNIT_NUMBER = " & moPatient.UnitNumber & vbNewLine
        
    '-- order by
    GetRadSQL = GetRadSQL & _
        "ORDER BY E.EVENT_START_GMTIME desc "

GoTo Cleanup
ErrHandler:
    Err.Raise Err.Number, App.EXEName & "." & cMODULE_NAME & "." & "GetRadSQL:" & Err.Source, _
       Err.Description, Err.HelpFile, Err.HelpContext
Cleanup:
    
End Function

Private Function GetVitalsSQL(ByVal strCategory As String, _
                              ByVal lngVitalsTypeID As Long) As String
On Error GoTo ErrHandler
    If gobjshared.UseSQLServer Then
        GetVitalsSQL = "SELECT '" & strCategory & "' AS CATEGORY, " & _
                            "(VITALSIGNSVALUE.VITALSVALUE + ' ' + NCIDS.DESCRIPTION) AS Result, " & _
                            "YEAR(VITALSIGNSVALUE.CreatedOn) AS Result_Year, " & _
                            "CONVERT(VARCHAR(12), VITALSIGNSVALUE.CreatedOn,106) as Date_Format," & _
                            "VITALSIGNSVALUE.CreatedOn as Event_Date " & _
                       "FROM VITALSIGNSVALUE, NCIDS, VITALS " & _
                       "WHERE (VITALSIGNSVALUE.Units = NCIDS.NCID) AND " & _
                           "(VITALSIGNSVALUE.DataId = VITALS.DATAID) AND " & _
                           "(VITALSIGNSVALUE.VITALSTYPE = " & CStr(lngVitalsTypeID) & ") AND " & _
                           "(VITALS.UNIT_NUMBER = " & moPatient.UnitNumber & ")" & " Order by Event_Date DESC "
    Else
        GetVitalsSQL = "SELECT '" & strCategory & "' AS CATEGORY, " & _
                            "(VS.VITALSVALUE || ' ' || NCIDS.DESCRIPTION) AS Result, " & _
                            "TO_NUMBER(TO_CHAR(V.ENTRYDATE, 'YYYY')) AS Result_Year, " & _
                            "TO_CHAR(V.ENTRYDATE,'DD MON YYYY') as Date_Format," & _
                            "V.ENTRYDATE as Event_Date " & _
                       "FROM VITALSIGNSVALUE VS, NCIDS, VITALS V " & _
                       "WHERE (VS.Units = NCIDS.NCID) AND " & _
                           "(VS.DataId = V.DATAID) AND " & _
                           "(VS.VITALSTYPE = " & CStr(lngVitalsTypeID) & ") AND " & _
                           "(V.UNIT_NUMBER = " & moPatient.UnitNumber & ")" & " Order by Event_Date DESC "
    End If
    
GoTo Cleanup
ErrHandler:
    Err.Raise Err.Number, App.EXEName & "." & cMODULE_NAME & "." & "GetVitalsSQL:" & Err.Source, _
       Err.Description, Err.HelpFile, Err.HelpContext
Cleanup:
    
                       
End Function

Public Function LayoutFamilyHistoryRS( _
    ByRef objRecordsetFrom As ADODB.Recordset, _
    ByRef objRecordset As ADODB.Recordset, _
    ByRef objRecordsetSub As ADODB.Recordset) As Boolean
    
    Const MAX_ROW_DISPLAIED As Integer = 2
    Dim intIndex As Integer
    
    On Error GoTo Proc_Exit
    
    LayoutFamilyHistoryRS = False
    '-- Copy first 2 results to new Recordset
    objRecordset.Open
    objRecordset.AddNew
    objRecordset.Fields("job_id") = "1"

    If Not (objRecordsetFrom.BOF And objRecordsetFrom.EOF) Then
        objRecordsetFrom.MoveFirst
        intIndex = 0
        Do While Not (objRecordsetFrom.EOF Or intIndex >= MAX_ROW_DISPLAIED) 'Wilson modified to '>='
            intIndex = intIndex + 1
            objRecordset.Fields("PROBLEM" & CStr(intIndex)).Value = objRecordsetFrom.Fields("PROBLEM").Value
            objRecordset.Fields("RELATION" & CStr(intIndex)).Value = objRecordsetFrom.Fields("RELATION").Value
            objRecordset.Fields("ONSETDATE" & CStr(intIndex)).Value = objRecordsetFrom.Fields("ONSETDATE").Value
            objRecordset.Fields("SnoID" & CStr(intIndex)).Value = objRecordsetFrom.Fields("SnoID").Value
            objRecordsetFrom.MoveNext
        Loop
        If objRecordsetFrom.RecordCount > MAX_ROW_DISPLAIED Then
            'Overwrite the last index (MAX_ROW_DISPLAIED) value
            objRecordset.Fields("PROBLEM" & CStr(intIndex)).Value = "(See continuation page)"
            objRecordset.Fields("RELATION" & CStr(intIndex)).Value = vbNullString
            objRecordset.Fields("ONSETDATE" & CStr(intIndex)).Value = vbNullString
            objRecordset.Fields("SnoID" & CStr(intIndex)).Value = vbNullString
       End If
    End If

    objRecordset.Update

    'Copy recordset
    Set objRecordsetSub = New ADODB.Recordset

    'Wilson modified following
    'DuplicateDefinition objRecordset, objRecordsetSub, False
    DuplicateDefinition objRecordsetFrom, objRecordsetSub, False

    '-- If records 3 and greater (if exists) to new recordset for subreport
    If objRecordsetFrom.RecordCount > MAX_ROW_DISPLAIED Then
        objRecordsetFrom.Move MAX_ROW_DISPLAIED - 1, 1
        Do While Not objRecordsetFrom.EOF
            objRecordsetSub.AddNew
            For intIndex = 0 To objRecordsetFrom.Fields.Count - 1
                objRecordsetSub.Fields(intIndex).Value = objRecordsetFrom.Fields(intIndex).Value
            Next
            objRecordsetSub.Update
            objRecordsetFrom.MoveNext
        Loop
    End If
    LayoutFamilyHistoryRS = True

Proc_Exit:
    If Err.Number <> 0 Then
        LayoutFamilyHistoryRS = False
        gobjshared.ShowVBError Err.Number, Err.Description, cMODULE_NAME & ".LayoutFamilyHistoryRS", App.EXEName, vbExclamation
    End If
    Exit Function

End Function
