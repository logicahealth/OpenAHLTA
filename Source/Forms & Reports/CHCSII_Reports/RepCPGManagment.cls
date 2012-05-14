VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RepCPGManagement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private moCollRecSet As Collection
Private moCollFormulas As Collection

'-- Goals NCIDs
'-------------------------------------------
'-- Labs
Private Const cHGB_A1C_NCID = "8571"
Private Const cMICROALBUMIN = "31961,31968,31959,31958,45555"
Private Const cLDL = "6094"
Private Const cHDL = "6090"
Private Const cTriglycerides = "21552"
Private Const cCholesterol = "6097"

'-- Vitals
Private Const cBP = "1985:1976"
Private Const cBP_SYSTOLIC = "1985"
Private Const cBP_DIASTOLIC = "1976"
Private Const cBMI_NCID = "16752178"

'-- Reports Recordsets
Private moRecDemographics As ADODB.Recordset
Private moRecHGB As ADODB.Recordset
Private moRecMicro As ADODB.Recordset
Private moRecHDL As ADODB.Recordset
Private moRecLDL As ADODB.Recordset
Private moRecTri As ADODB.Recordset
Private moRecChol As ADODB.Recordset

'Private moRecLipid As ADODB.Recordset
Private moRecBloodSys As ADODB.Recordset
Private moRecBloodDia As ADODB.Recordset
Private moRecBodyMass As ADODB.Recordset
Private moRecEye As ADODB.Recordset
Private moRecFoot As ADODB.Recordset
Private moRecHealth As ADODB.Recordset
Private moRecEd As ADODB.Recordset


Private msFacilityLabContextNCID As String
Private msFacilityNCID As String

'-- Not used anymore
Private Const cWEIGHT_NCID = "2178"
Private Const cHEIGHT_NCID = "110675"

'-- Questionnaire
Private Const cQuestSequences = "1,2,4,30"


Private Function BuildFilter(sField As String, sFieldData, Optional sField2 As String = "x", Optional sFieldData2) As String
    
On Error GoTo ErrHandler
    Dim cData As Collection
    Dim vData As Variant
    Dim iLastPos As Integer
    Dim iPos As Integer
    Dim iLength As Integer
    
    '-- Put String Data into Collection
    Set cData = New Collection
    iLength = Len(sFieldData)
    iLastPos = 1
    Do
        iPos = InStr(iLastPos, sFieldData, ",")
        If iPos = 0 Then
            cData.Add Trim(Mid(sFieldData, iLastPos))
            Exit Do
        Else
           cData.Add Trim(Mid(sFieldData, iLastPos, iPos - iLastPos))
           iLastPos = iPos + 1
        End If
    Loop
    
    '-- Build Filter
    BuildFilter = ""
    For Each vData In cData
        If BuildFilter <> "" Then
           BuildFilter = BuildFilter & " or "
        End If

        If sField2 = "x" Then
            BuildFilter = BuildFilter & "(" & sField & " = " & vData & ")"
        Else
            BuildFilter = BuildFilter & "(" & sField & " = " & vData & " and " & _
                                              sField2 & " = " & sFieldData2 & ")"
        End If
    Next
    
GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepCPGGoals.BuildFilter", "CHCSII_Reports", vbExclamation
Cleanup:
    Set cData = Nothing
    Exit Function
    
End Function


Public Function BuildReport(sReportName As String, _
                            oReportItem As ReportItem, _
                            oReport As CRAXDRT.Report, _
                            oCollFormulas As Collection) As Boolean

On Error GoTo ErrHandler

BuildReport = False

Const csQuestionnaireName As String = "'DIABETES ANNUAL QUESTIONNAIRE'"
Const csDIABETES_CPG As String = "'DIABETES'"

Dim sDIABESTES_CPG As String

Dim oCrystal As Crystal
Dim oRec As ADODB.Recordset
Dim sSponsorUnitNumber As String
Dim sUnitNumber As Variant
Dim oPatientsColl As Collection
Dim sPatients As String

Dim oRecGoals As ADODB.Recordset
Dim oRecQuestions As ADODB.Recordset
Dim oRecAnswers As ADODB.Recordset
Dim oRecResponses As ADODB.Recordset

Dim oRecLab As ADODB.Recordset
Dim oRecVitals As ADODB.Recordset

Dim sUnitNumberList As String
Dim sVitalsList As String
Dim sLabsList As String

Dim sQDEF_IDs As String
Dim sCPG_ID As String
Dim i As Integer

'-- Get Data
'----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection

    '-- Initialize Recordsets
    '----------------------------------------------------------
    If Not Initialize_RecordSets() Then GoTo Cleanup
       
'-- Build Lists of Patients
    If Not GetPatients(oReportItem, oPatientsColl) Then GoTo Cleanup
    If Not GetUnitNumberList(oPatientsColl, sUnitNumberList) Then GoTo Cleanup
        
'-- Get Labs
    sLabsList = cHGB_A1C_NCID & "," & cMICROALBUMIN & "," & cLDL & "," & cHDL & _
                                "," & cTriglycerides & "," & cCholesterol
    If Not GetLabResults(oRecLab, sUnitNumberList, sLabsList) Then GoTo Cleanup
    
'-- Get Vitals
    sVitalsList = cBP_SYSTOLIC & "," & cBP_DIASTOLIC & "," & cBMI_NCID
    If Not GetVitals(oRecVitals, sUnitNumberList, sVitalsList) Then GoTo Cleanup

'-- Get Questionaire Information
    If Not GetQDEF_ID(csQuestionnaireName, sQDEF_IDs) Then GoTo Cleanup
       
    '-- Build CPG list
    '--------------------------------------------------------
    For i = 1 To oReportItem.CriteriaSel2.Count
        If i = 1 Then
            sCPG_ID = oReportItem.CriteriaSel2.Item(i).CPG_ID
        Else
            sCPG_ID = sCPG_ID & "," & oReportItem.CriteriaSel2.Item(i).CPG_ID
        End If
    Next

    
    For Each sUnitNumber In oPatientsColl
        
        If IsInRegistry(CStr(sUnitNumber), sCPG_ID) Then
            
            '-- Get Goals
            '----------------------------------------------------------
            If Not GetGoals(oRecGoals, CStr(sUnitNumber), sCPG_ID) Then GoTo Cleanup
        
            '-- Get Demographics
            '----------------------------------------------------------
            If Not GetDemographics(CStr(sUnitNumber)) Then GoTo Cleanup

            '-- Get HGB
            '----------------------------------------------------------
            If Not GetPatientLabStatus(CStr(sUnitNumber), cHGB_A1C_NCID, oRecLab, oRecGoals, moRecHGB) Then GoTo Cleanup

            '-- Get Microalbumin
            '----------------------------------------------------------
            If Not GetPatientLabStatus(CStr(sUnitNumber), cMICROALBUMIN, oRecLab, oRecGoals, moRecMicro) Then GoTo Cleanup
    
            '-- Get LDL
            '----------------------------------------------------------
            If Not GetPatientLabStatus(CStr(sUnitNumber), cLDL, oRecLab, oRecGoals, moRecLDL) Then GoTo Cleanup
    
            '-- Get HDL
            '----------------------------------------------------------
            If Not GetPatientLabStatus(CStr(sUnitNumber), cHDL, oRecLab, oRecGoals, moRecHDL) Then GoTo Cleanup
    
            '-- Get Cholesterol
            '----------------------------------------------------------
            If Not GetPatientLabStatus(CStr(sUnitNumber), cCholesterol, oRecLab, oRecGoals, moRecChol) Then GoTo Cleanup
    
            '-- Get Triglycerides
            '----------------------------------------------------------
            If Not GetPatientLabStatus(CStr(sUnitNumber), cTriglycerides, oRecLab, oRecGoals, moRecTri) Then GoTo Cleanup
    
            '-- Get Body Mass
            '----------------------------------------------------------
            If Not GetPatientVitalStatus(CStr(sUnitNumber), cBMI_NCID, oRecVitals, oRecGoals, moRecBodyMass) Then GoTo Cleanup
    
            '-- Get Blood Pressure Systolic
            '----------------------------------------------------------
            If Not GetPatientBPStatus(CStr(sUnitNumber), oRecVitals, oRecGoals, moRecBloodSys) Then GoTo Cleanup
       
            '-- Get Health Status
            '----------------------------------------------------------
            If Not GetPatientAnswer(CStr(sUnitNumber), "1", sQDEF_IDs, moRecHealth) Then GoTo Cleanup
    
            '-- Get Foot Exam
            '----------------------------------------------------------
            If Not GetPatientAnswer(CStr(sUnitNumber), "2", sQDEF_IDs, moRecFoot) Then GoTo Cleanup
            
            '-- Get Eye Exam
            '----------------------------------------------------------
            'If Not GetPatientAnswerStatus(CStr(sUnitNumber), "4", oRecResponses, oRecAnswers, moRecEye) Then GoTo CleanUp
            If Not GetPatientAnswer(CStr(sUnitNumber), "4", sQDEF_IDs, moRecEye) Then GoTo Cleanup

            '-- Get Education
            '----------------------------------------------------------
            'If Not GetPatientAnswerStatus(CStr(sUnitNumber), "30", oRecResponses, oRecAnswers, moRecEd) Then GoTo CleanUp
            If Not GetPatientAnswer(CStr(sUnitNumber), "30", sQDEF_IDs, moRecEd) Then GoTo Cleanup

        End If
    Next

    moCollRecSet.Add moRecDemographics, "demographics"
    moCollRecSet.Add moRecHGB, "hgb"
    moCollRecSet.Add moRecMicro, "microalbumin"
    moCollRecSet.Add moRecHDL, "hdl"
    moCollRecSet.Add moRecLDL, "ldl"
    moCollRecSet.Add moRecChol, "cholesterol"
    moCollRecSet.Add moRecTri, "triglycerides"
    moCollRecSet.Add moRecBodyMass, "bmi"
    moCollRecSet.Add moRecBloodSys, "bps"
    moCollRecSet.Add moRecEye, "eyeexam"
    moCollRecSet.Add moRecHealth, "healthstatus"
    moCollRecSet.Add moRecFoot, "footexam"
    moCollRecSet.Add moRecEd, "education"

'-- Create Report Object
'-----------------------------------------------------------------------------
    If Not CreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo Cleanup

BuildReport = True

GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepCPGGoals.BuildReport", "CHCSII_Reports", vbExclamation
Cleanup:
    Set oCrystal = Nothing
    Set oRec = Nothing
    
End Function
Private Function GetGoals(oRecGoals As ADODB.Recordset, _
                         sUnitNumber As String, _
                         CPG_ID As String) As Boolean

On Error GoTo ErrHandler

GetGoals = False

Dim sStmt As String
Dim sRefCode As String
Dim sRefText As String
Dim iFrequency As Integer
Dim oRec As ADODB.Recordset
Dim iPrecendence As Integer


Set oRecGoals = Nothing
Set oRecGoals = New ADODB.Recordset
oRecGoals.Fields.Append "UNIT_NUMBER", adUnsignedBigInt
oRecGoals.Fields.Append "REFCODE", adLongVarWChar, 50 'adUnsignedBigInt
oRecGoals.Fields.Append "FREQUENCY", adSingle
oRecGoals.Fields.Append "VALUE", adLongVarWChar, 50
oRecGoals.Fields.Append "GOAL_OP_TYPE", adLongVarWChar, 10
oRecGoals.Fields.Append "GoalOpText", adLongVarWChar, 50
oRecGoals.Fields.Append "PRECEDENCE", adInteger
oRecGoals.Open

'-- Select
'-----------------------------------------------------
sStmt = _
    "Select " & vbNewLine & _
        "p.unit_number, " & vbNewLine & _
        "p.poc_type, " & vbNewLine & _
        "p.poc_id, " & vbNewLine & _
        "p.name, " & vbNewLine & _
        "p.category, " & vbNewLine & _
        "p.goal_type, " & vbNewLine & _
        "p.references," & vbNewLine & _
        "p.frequency, " & vbNewLine & _
        "p.goal_op_type, " & vbNewLine & _
        "p.value, " & vbNewLine & _
        "p.creator_ncid, " & vbNewLine & _
        "p.date_created, " & vbNewLine & _
        "p.modifier_ncid, " & vbNewLine & _
        "p.date_modified, " & vbNewLine

sStmt = sStmt & _
        "pd.Name as PlanOfCareName," & vbNewLine
        
sStmt = sStmt & _
        "Decode(p.date_modified,Null,p.date_created,p.date_modified) as LastModDate, " & vbNewLine & _
        "Decode(p.category,1,'Monitered Items',2,'Referrals',3,'PCM Considerations',4,'Other','Unknown') as CategoryName," & vbNewLine & _
        "Decode(p.frequency,1,'One Time',2,'Every Visit',3,'Annually',4,'Semi-Annually',5,'Quarterly',6,'Monthly',7,'Semi-Monthly',8,'Weekly',9,'Daily','Unknown') as FrequencyText," & vbNewLine & _
        "Decode(p.goal_op_type,1,'=',2,'<>',3,'>',4,'>=',5,'<',6,'<=','') as GoalOpText," & vbNewLine & _
        "'        ' as RefCode, " & vbNewLine & _
        "'                                                                 ' as RefText " & vbNewLine & _
        ""
        
'-- From
'-----------------------------------------------------
sStmt = sStmt & _
    "From" & vbNewLine & _
        "pocgoal p," & vbNewLine & _
        "pocdef pd" & vbNewLine

'-- Where
'-----------------------------------------------------
sStmt = sStmt & _
    "Where" & vbNewLine

sStmt = sStmt & _
        "not exists " & vbNewLine & _
        "(select 'x'" & vbNewLine & _
        "from pocgoal po" & vbNewLine & _
        "where po.poc_type=3 and p.category=99 and" & vbNewLine & _
        "po.unit_number = " & sUnitNumber & " and" & vbNewLine & _
        "po.name = p.name and" & vbNewLine & _
        "po.poc_id = p.poc_id) and" & vbNewLine & _
        "p.unit_number = 0 and" & vbNewLine & _
        "p.poc_type=1 and" & vbNewLine & _
        "p.poc_id = pd.pocdef_id and " & vbNewLine & _
        "pd.cpg_id in (select c.cpg_id from cpgregistry c where c.unit_number = " & sUnitNumber & " and c.cpg_id in (" & CPG_ID & ")) " & vbNewLine
        

sStmt = sStmt & "Union" & vbNewLine

'-- Select
'-----------------------------------------------------
sStmt = sStmt & _
    "Select " & vbNewLine & _
        "p.unit_number, " & vbNewLine & _
        "p.poc_type, " & vbNewLine & _
        "p.poc_id, " & vbNewLine & _
        "p.name, " & vbNewLine & _
        "p.category, " & vbNewLine & _
        "p.goal_type, " & vbNewLine & _
        "p.references," & vbNewLine & _
        "p.frequency, " & vbNewLine & _
        "p.goal_op_type, " & vbNewLine & _
        "p.value, " & vbNewLine & _
        "p.creator_ncid, " & vbNewLine & _
        "p.date_created, " & vbNewLine & _
        "p.modifier_ncid, " & vbNewLine & _
        "p.date_modified, " & vbNewLine

sStmt = sStmt & _
        "'   ' as PlanOfCareName," & vbNewLine
    
        
sStmt = sStmt & _
        "Decode(p.date_modified,Null,p.date_created,p.date_modified) as LastModDate, " & vbNewLine & _
        "Decode(p.category,1,'Monitered Items',2,'Referrals',3,'PCM Considerations',4,'Other','Unknown') as CategoryName," & vbNewLine & _
        "Decode(p.frequency,1,'One Time',2,'Every Visit',3,'Annually',4,'Semi-Annually',5,'Quarterly',6,'Monthly',7,'Semi-Monthly',8,'Weekly',9,'Daily','Unknown') as FrequencyText," & vbNewLine & _
        "Decode(p.goal_op_type,1,'=',2,'<>',3,'>',4,'>=',5,'<',6,'<=','') as GoalOpText," & vbNewLine & _
        "'        ' as RefCode, " & vbNewLine & _
        "'                                                                 ' as RefText " & vbNewLine & _
        ""

'-- From
'-----------------------------------------------------
sStmt = sStmt & _
    "From" & vbNewLine & _
        "pocgoal p" & vbNewLine
        
'-- Where
'-----------------------------------------------------
sStmt = sStmt & _
    "Where" & vbNewLine

sStmt = sStmt & _
        "(p.poc_type = 2 or (p.poc_type=3 and p.category<>99)) and" & vbNewLine & _
        "p.unit_number in (" & sUnitNumber & ") " & vbNewLine



Set oRec = goDBDAS.OpenRecordset(sStmt)

'-- Update Reference Information
'-------------------------------------------------------------------
If oRec.RecordCount <> 0 Then
    oRec.MoveFirst
    Do While Not oRec.EOF
        oRecGoals.AddNew
        On Error Resume Next
        sRefCode = Mid$(oRec.Fields("References").Value, 1, InStr(oRec.Fields("References").Value, vbTab) - 1)
        If Err <> 0 Then
            sRefCode = "0"
            sRefText = oRec.Fields("References").Value
        Else
            sRefText = Mid$(oRec.Fields("References").Value, Len(sRefCode) + 2)
            If Mid(sRefText, Len(sRefText), 1) = "|" Then
                sRefText = Mid$(sRefText, 1, Len(sRefText) - 1)
            End If
        End If
        oRecGoals("REFCODE").Value = sRefCode
        oRec.Fields("RefCode").Value = sRefCode
        oRec.Fields("RefText").Value = sRefText
        
        If Not GetFrequencyInDays(oRec.Fields("frequency"), iFrequency) Then GoTo Cleanup
        oRec.Fields("FrequencyInDays").Value = iFrequency
        oRecGoals("FREQUENCY").Value = iFrequency

        If Not GetFrequencyByPrecedence(oRec.Fields("frequency"), iPrecendence) Then GoTo Cleanup
        oRecGoals("PRECEDENCE").Value = iPrecendence

        oRecGoals("UNIT_NUMBER").Value = oRec.Fields("UNIT_NUMBER")
        oRecGoals("GOAL_OP_TYPE").Value = oRec.Fields("GOAL_OP_TYPE")
        oRecGoals("VALUE").Value = oRec.Fields("VALUE")
        oRecGoals("GoalOpText").Value = oRec.Fields("GoalOpText")
        
        oRecGoals.Update
        oRec.Update
        oRec.MoveNext
    Loop
End If

GetGoals = True
GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "CPGGlobal.GetGoals", "CHCSII_Reports", vbExclamation
Cleanup:
    Exit Function
End Function




Private Function DuplicateDefinition(oRecSource As ADODB.Recordset, _
                                     oRecDest As ADODB.Recordset) As Boolean
                                     
On Error GoTo ErrHandler
Dim i As Integer

DuplicateDefinition = False

Set oRecDest = New ADODB.Recordset

For i = 0 To oRecSource.Fields.Count - 1
    oRecDest.Fields.Append oRecSource.Fields(i).Name, oRecSource.Fields(i).Type, oRecSource.Fields(i).DefinedSize
Next

oRecDest.Open

DuplicateDefinition = True

GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepCPGGoals.DuplicateDefinition", "CHCSII_Reports", vbExclamation
Cleanup:
    Exit Function
End Function



Private Function GetCount(sNCID As String) As Integer

Dim sStmt As String

sStmt = "select count(*) from rsform where ncid = " & sNCID

goDBSqlOps.Execute sStmt

GetCount = goDBSqlOps.Value(1)

End Function


Private Function GetPatientAnswer(sUnitNumber As String, _
                                  sSequence As String, _
                                  sQDEF_IDs As String, _
                                  oRecResult As ADODB.Recordset) As Boolean
                                  
Dim sStmt As String
Dim oRec As ADODB.Recordset

GetPatientAnswer = False

'-- Find latest answer to question
'---------------------------------------------------------------------

'-- Select
    sStmt = "Select " & vbNewLine & _
                "Q.UNIT_NUMBER, " & vbNewLine & _
                "QUESTION_TEXT  ANSWER, " & vbNewLine & _
                "TO_CHAR(END_DATE_TIME, 'DD MON YYYY') END_DATE_TIME" & vbNewLine

'-- From
    sStmt = sStmt & _
            "From " & vbNewLine & _
                "QRESP Q, " & vbNewLine & _
                "ARESP A, " & vbNewLine & _
                "QDEF_CHOICES QDC " & vbNewLine
                
'-- Where
    sStmt = sStmt & _
        "Where " & vbNewLine & _
            "QDC.VERSION = Q.VERSION AND" & vbNewLine & _
            "QDC.SEQ = A.SEQ AND" & vbNewLine & _
            "QDC.CHOICE = A.ANSWER AND" & vbNewLine & _
            "QDC.QDEF_ID = Q.QDEF_ID AND" & vbNewLine & _
            "A.SEQ = " & sSequence & " AND " & vbNewLine & _
            "A.QRESP_ID = Q.QRESP_ID AND" & vbNewLine

sStmt = sStmt & _
           "Q.QRESP_ID IN" & vbNewLine & _
           "(SELECT MAX(Q.QRESP_ID)" & vbNewLine & _
            "FROM QRESP Q" & vbNewLine & _
            "WHERE " & vbNewLine & _
            "Q.STATUS IN (3,4,5) AND Q.QDEF_ID IN (" & sQDEF_IDs & ") AND " & vbNewLine & _
            "Q.UNIT_NUMBER = " & sUnitNumber & " AND Q.END_DATE_TIME IN " & vbNewLine & _
            "(SELECT MAX(Q.END_DATE_TIME)" & vbNewLine & _
            "FROM QRESP Q" & vbNewLine & _
            "WHERE " & vbNewLine & _
            "Q.STATUS IN (3,4,5) AND Q.QDEF_ID IN (" & sQDEF_IDs & ") AND " & vbNewLine & _
            "Q.UNIT_NUMBER = " & sUnitNumber & "))"
            
Set oRec = goDBDAS.OpenRecordset(sStmt)

'-- Add New Record
'---------------------------------------------------------------------


oRecResult.AddNew
oRecResult.Fields("Unit_Number") = sUnitNumber

If oRec.RecordCount > 0 Then
    oRec.MoveFirst
    oRecResult.Fields("Start_Date") = oRec("END_DATE_TIME").Value 'Format(oRec("START_DATE_TIME").Value, "DD MMM YYYY")
    oRecResult.Fields("Answer") = oRec("ANSWER").Value
Else
    oRecResult.Fields("Start_Date") = "None"
    oRecResult.Fields("Answer") = "None"
End If

oRecResult.Update

GetPatientAnswer = True


GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepCPGManagement.GetPatientAnswer", "CHCSII_Reports", vbExclamation
Cleanup:
    Set oRec = Nothing
    Exit Function
End Function


Private Function GetPatientLabStatus(sUnitNumber As String, _
                                     sLabId As String, _
                                     oRecLabs As ADODB.Recordset, _
                                     oRecGoals As ADODB.Recordset, _
                                     oRecResult As ADODB.Recordset) As Boolean

On Error GoTo ErrHandler

GetPatientLabStatus = False

Dim iFrequency As Integer
Dim vFrequency As Variant
Dim vAbnormal As Variant
Dim bAbnormal As Boolean
Dim vGoal As Variant
Dim vCurrent As Variant
Dim vOperator As Variant
Dim vDate As Variant
Dim sOverDue As String

'-- Filter out only goals for specified lab
oRecGoals.Filter = ""
If oRecGoals.RecordCount <> 0 Then
    oRecGoals.Filter = BuildFilter("RefCode", sLabId)
    oRecGoals.Sort = "unit_number desc, frequency asc"
End If



'-- Filter out specific lab
oRecLabs.Filter = ""
If oRecLabs.RecordCount <> 0 Then
    oRecLabs.Filter = BuildFilter("Lab_Id", sLabId, "unit_number", sUnitNumber)
End If

If oRecGoals.RecordCount <> 0 Then
    vGoal = oRecGoals("value").Value
    vOperator = oRecGoals("goal_op_type").Value
    vFrequency = oRecGoals("frequency").Value
End If

If oRecLabs.RecordCount <> 0 Then
    vAbnormal = oRecLabs("abnormal_flag_ncid").Value
    If IsEmpty(oRecLabs("Result_Value").Value) Then
        vCurrent = oRecLabs("Text_Value").Value
    Else
        vCurrent = oRecLabs("Result_Value").Value
    End If

    vDate = oRecLabs("Collect_Date").Value
End If

'-- Check if Abnormal
'---------------------------------------------------
If Not VarType(vGoal) = vbEmpty And Not VarType(vCurrent) = vbEmpty And Not VarType(vOperator) = vbEmpty Then
    On Error Resume Next
    Select Case vOperator
        Case 1 'Equal
            bAbnormal = Not (CSng(vGoal) = CSng(vCurrent))
        Case 2 'Not Equal
            bAbnormal = Not (CSng(vGoal) <> CSng(vCurrent))
        Case 3 'Greater Than
            bAbnormal = Not (CSng(vGoal) < CSng(vCurrent))
        Case 4 'Greater Than Equal To
            bAbnormal = Not (CSng(vGoal) <= CSng(vCurrent))
        Case 5 'Less Than
            bAbnormal = Not (CSng(vGoal) > CSng(vCurrent))
        Case 6 'Less Than Equal To
            bAbnormal = Not (CSng(vGoal) >= CSng(vCurrent))
        Case Else
            bAbnormal = True
    End Select
    If Err.Number <> 0 Then
        bAbnormal = False
    End If
    On Error GoTo ErrHandler
    
Else
    bAbnormal = False
End If


'-- Check if Overdue or Coming Due
'---------------------------------------------------
If Not GetOverDue(vFrequency, vDate, sUnitNumber, sOverDue) Then GoTo Cleanup

'If Not VarType(vFrequency) = vbEmpty And Not VarType(vDate) = vbEmpty Then
'
'    If vFrequency = -1 Then
'        If VarType(vDate) = vbEmpty Then
'            sOverDue = "O"
'        Else
'            sOverDue = ""
'        End If
'    ElseIf vFrequency = -2 Then
'
'    ElseIf vFrequency = 1 Then
'        If CompletedLastVisit(sUnitNumber, oRecLabs("Collect_Date").Value) Then
'            sOverDue = ""
'        Else
'            sOverDue = "O"
'        End If
'    Else
'        If (Now > vDate + vFrequency) Then
'            sOverDue = "O"
'        ElseIf (Now - (vDate + vFrequency)) > -60 Then
'            sOverDue = "C"
'        End If
'    End If
'ElseIf VarType(vFrequency) = vbEmpty Then
'    sOverDue = ""
'Else
'    sOverDue = "O"
'End If

'    If (Now > vDate + vFrequency) Then
'        sOverDue = "O"
'    ElseIf (Now - (vDate + vFrequency)) > -60 Then
'        sOverDue = "C"
'    End If
'ElseIf VarType(vFrequency) = vbEmpty Then
'    sOverDue = ""
'Else
'    sOverDue = "O"
'End If
    
oRecResult.AddNew
oRecResult.Fields("Unit_Number") = sUnitNumber
oRecResult.Fields("Goal_Value") = vGoal
oRecResult.Fields("Current_Value") = vCurrent
If Not VarType(vDate) = vbEmpty Then
    oRecResult.Fields("Current_Date") = Format(vDate, "DD MMM YYYY")
Else
    oRecResult.Fields("Current_Date") = "None"
End If
If bAbnormal Then
    oRecResult.Fields("ABNORMAL") = "Y"
Else
    oRecResult.Fields("ABNORMAL") = "N"
End If
oRecResult.Fields("Due") = sOverDue
If Not VarType(vOperator) = vbEmpty Then
    oRecResult.Fields("OP_TYPE") = oRecGoals("GoalOpText").Value
Else
    oRecResult.Fields("OP_TYPE") = ""
End If

If oRecGoals.RecordCount <> 0 Then
    oRecResult("GOAL").Value = "Y"
Else
    oRecResult("GOAL").Value = "N"
End If
oRecResult.Update

GetPatientLabStatus = True

GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepCPGGoals.GetPatientLabStatus", "CHCSII_Reports", vbExclamation
Cleanup:
    Exit Function
End Function
Private Function GetOverDue(vFrequency As Variant, _
                          vDate As Variant, _
                          sUnitNumber As String, _
                          sOverDue As String) As Boolean

On Error GoTo ErrHandler
GetOverDue = False

'-- Check if Overdue or Coming Due
'---------------------------------------------------
If Not VarType(vFrequency) = vbEmpty And Not VarType(vDate) = vbEmpty Then

    If vFrequency = 3000 Then
        If VarType(vDate) = vbEmpty Then
            sOverDue = "O"
        Else
            sOverDue = ""
        End If

    ElseIf vFrequency = 2 Then
        If CompletedLastVisit(sUnitNumber, vDate) Then
            sOverDue = ""
        Else
            sOverDue = "O"
        End If

    ElseIf vFrequency = 1 Then
        
        If (CDate(CStr(Date) & " 0:00:00 AM") < vDate And _
            vDate < CDate(CStr(Date) & " 11:59:59 PM")) Then
            sOverDue = ""
        Else
            sOverDue = "O"
        End If

    Else
        If (Now > vDate + vFrequency) Then
            sOverDue = "O"
        ElseIf (Now - (vDate + vFrequency)) > -60 Then
            sOverDue = "C"
        Else
            sOverDue = ""
        End If

    End If

ElseIf VarType(vFrequency) = vbEmpty Then
    sOverDue = ""
Else
    sOverDue = "O"
End If

GetOverDue = True
GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "CPGGlobal.GetOverdue", "CHCSII_Reports", vbExclamation
Cleanup:
    Exit Function
End Function




Private Function GetPatientBPStatus(sUnitNumber As String, _
                                    oRecVitals As ADODB.Recordset, _
                                    oRecGoals As ADODB.Recordset, _
                                    oRecResult As ADODB.Recordset) As Boolean

On Error GoTo ErrHandler

GetPatientBPStatus = False

Dim iFrequency As Integer
Dim vFrequency As Variant
Dim vAbnormal As Variant
Dim bAbnormalS As Boolean
Dim bAbnormalD As Boolean
Dim bAbnormal As Boolean
Dim vGoal As Variant
Dim vCurrent As Variant
Dim vOperator As Variant
Dim vDate As Variant
Dim sOverDue As String

Dim vG_SYS As Variant
Dim vG_DIA As Variant
Dim vGOAL_BP As Variant

Dim vSYS As Variant
Dim vDIA As Variant
Dim vCURRENT_BP As Variant


'-- Filter Goals for specified vital
oRecGoals.Filter = ""
If oRecGoals.RecordCount <> 0 Then
    oRecGoals.Filter = "REFCODE = " & cBP
    If oRecGoals.RecordCount <> 0 Then
        vG_SYS = CLng(Mid$(oRecGoals("value"), 1, InStr(1, oRecGoals("value"), "/") - 1))
        vG_DIA = CLng(Mid$(oRecGoals("value"), InStr(1, oRecGoals("value"), "/") + 1))
        vGOAL_BP = CSng(vG_SYS / vG_DIA)
    End If
End If

'-- Filter specified vital
oRecVitals.Filter = ""
If oRecVitals.RecordCount <> 0 Then
    oRecVitals.Filter = "unit_number = " & sUnitNumber & " and Vital_Id = " & cBP_SYSTOLIC
    If oRecVitals.RecordCount <> 0 Then
        vSYS = CLng(oRecVitals("Value"))
        oRecVitals.Filter = "unit_number = " & sUnitNumber & " and Vital_Id = " & cBP_DIASTOLIC
        If oRecVitals.RecordCount <> 0 Then
            vDIA = CLng(oRecVitals("Value"))
            vCURRENT_BP = CSng(vSYS / vDIA)
            vDate = oRecVitals("Collect_Date").Value
        End If
    End If
End If

If oRecGoals.RecordCount <> 0 Then
    vGoal = oRecGoals("value").Value
    vOperator = oRecGoals("goal_op_type").Value
    vFrequency = oRecGoals("frequency").Value
End If

'-- Check if Abnormal
'---------------------------------------------------
If Not VarType(vGOAL_BP) = vbEmpty And Not VarType(vCURRENT_BP) = vbEmpty And Not VarType(vOperator) = vbEmpty Then
    
    On Error Resume Next
    Select Case vOperator
        Case 1 'Equal
            bAbnormalS = Not (vG_SYS = vSYS)
        Case 2 'Not Equal
            bAbnormalS = Not (vG_SYS <> vSYS)
        Case 3 'Greater Than
            bAbnormalS = Not (vG_SYS < vSYS)
        Case 4 'Greater Than Equal To
            bAbnormalS = Not (vG_SYS <= vSYS)
        Case 5 'Less Than
            bAbnormalS = Not (vG_SYS > vSYS)
        Case 6 'Less Than Equal To
            bAbnormalS = Not (vG_SYS >= vSYS)
        Case Else
            bAbnormalS = True
    End Select
    If Err.Number <> 0 Then
        bAbnormalS = False
    End If
    On Error GoTo ErrHandler


    On Error Resume Next
    Select Case vOperator
        Case 1 'Equal
            bAbnormalD = Not (vG_DIA = vDIA)
        Case 2 'Not Equal
            bAbnormalD = Not (vG_DIA <> vDIA)
        Case 3 'Greater Than
            bAbnormalD = Not (vG_DIA < vDIA)
        Case 4 'Greater Than Equal To
            bAbnormalD = Not (vG_DIA <= vDIA)
        Case 5 'Less Than
            bAbnormalD = Not (vG_DIA > vDIA)
        Case 6 'Less Than Equal To
            bAbnormalD = Not (vG_DIA >= vDIA)
        Case Else
            bAbnormalD = True
    End Select
    If Err.Number <> 0 Then
        bAbnormalD = False
    End If
    On Error GoTo ErrHandler

    If bAbnormalS Or bAbnormalD Then
        bAbnormal = True
    Else
        bAbnormal = False
    End If

Else
    bAbnormal = False
End If


'-- Check if Overdue or Coming Due
'---------------------------------------------------
If Not GetOverDue(vFrequency, vDate, sUnitNumber, sOverDue) Then GoTo Cleanup

'If Not VarType(vFrequency) = vbEmpty And Not VarType(vDate) = vbEmpty Then
'    If Not GetFrequencyInDays(CInt(vFrequency), iFrequency) Then GoTo CleanUp
'
'    If vFrequency = -1 Then
'        If VarType(vDate) = vbEmpty Then
'            sOverDue = "O"
'        Else
'            sOverDue = ""
'        End If
'    ElseIf vFrequency = -2 Then
'
'    ElseIf iFrequency = 0 Then
'        sOverDue = ""
'    ElseIf iFrequency = 1 Then
'        If CompletedLastVisit(sUnitNumber, oRecVitals("Collect_Date").Value) Then
'            sOverDue = ""
'        Else
'            sOverDue = "O"
'        End If
'    Else
'        If (Now > vDate + iFrequency) Then
'            sOverDue = "O"
'        ElseIf (Now - (vDate + iFrequency)) > -60 Then
'            sOverDue = "C"
'        End If
'    End If
'ElseIf VarType(vFrequency) = vbEmpty Then
'    sOverDue = ""
'Else
'    sOverDue = "O"
'End If


oRecResult.AddNew
oRecResult.Fields("Unit_Number") = sUnitNumber
oRecResult.Fields("Goal_Value") = vGoal

If VarType(vSYS) = vbEmpty Then
    oRecResult.Fields("Current_Value") = "None"
Else
    oRecResult.Fields("Current_Value") = CStr(vSYS) & "/" & CStr(vDIA)
End If

If Not VarType(vDate) = vbEmpty Then
    oRecResult.Fields("Current_Date") = Format(vDate, "DD MMM YYYY")
Else
    oRecResult.Fields("Current_Date") = "None"
End If

If bAbnormal Then
    oRecResult.Fields("ABNORMAL") = "Y"
Else
    oRecResult.Fields("ABNORMAL") = "N"
End If
oRecResult.Fields("Due") = sOverDue
If Not VarType(vOperator) = vbEmpty Then
    oRecResult.Fields("OP_TYPE") = oRecGoals("GoalOpText").Value
Else
    oRecResult.Fields("OP_TYPE") = ""
End If

If oRecGoals.RecordCount <> 0 Then
    oRecResult("GOAL").Value = "Y"
Else
    oRecResult("GOAL").Value = "N"
End If

oRecResult.Update

GetPatientBPStatus = True

GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepCPGGoals.GetPatientBPStatus", "CHCSII_Reports", vbExclamation
    'Resume Next
Cleanup:
    Exit Function
End Function
Private Function GetPatientVitalStatus(sUnitNumber As String, _
                                       sVitalId As String, _
                                       oRecVitals As ADODB.Recordset, _
                                       oRecGoals As ADODB.Recordset, _
                                       oRecResult As ADODB.Recordset) As Boolean

On Error GoTo ErrHandler

GetPatientVitalStatus = False

Dim iFrequency As Integer
Dim vFrequency As Variant
Dim vAbnormal As Variant
Dim bAbnormal As Boolean
Dim vGoal As Variant
Dim vCurrent As Variant
Dim vOperator As Variant
Dim vDate As Variant
Dim sOverDue As String

'-- Filter Goals for specified vital
oRecGoals.Filter = ""
If oRecGoals.RecordCount <> 0 Then
    oRecGoals.Filter = "RefCode = '" & sVitalId & "'"
    oRecGoals.Sort = "unit_number desc, precedence asc"
End If

'-- Filter specified vital
oRecVitals.Filter = ""
If oRecVitals.RecordCount <> 0 Then
    oRecVitals.Filter = "unit_number = " & sUnitNumber & " and Vital_Id = " & sVitalId
End If

If oRecGoals.RecordCount <> 0 Then
    vGoal = oRecGoals("value").Value
    vOperator = oRecGoals("goal_op_type").Value
    vFrequency = oRecGoals("frequency").Value
End If

If oRecVitals.RecordCount <> 0 Then
    If oRecVitals("Value_String") <> "" Then
        If IsNumeric(oRecVitals("Value_String")) Then vCurrent = CSng(oRecVitals("Value_String"))
    Else
        vCurrent = oRecVitals("Value")
    End If
    vDate = oRecVitals("Collect_Date").Value
End If

'-- Check if Abnormal
'---------------------------------------------------
If Not VarType(vGoal) = vbEmpty And Not VarType(vCurrent) = vbEmpty And Not VarType(vOperator) = vbEmpty Then
    On Error Resume Next
    Select Case vOperator
        Case 1 'Equal
            bAbnormal = Not (CSng(vGoal) = CSng(vCurrent))
        Case 2 'Not Equal
            bAbnormal = Not (CSng(vGoal) <> CSng(vCurrent))
        Case 3 'Greater Than
            bAbnormal = Not (CSng(vGoal) < CSng(vCurrent))
        Case 4 'Greater Than Equal To
            bAbnormal = Not (CSng(vGoal) <= CSng(vCurrent))
        Case 5 'Less Than
            bAbnormal = Not (CSng(vGoal) > CSng(vCurrent))
        Case 6 'Less Than Equal To
            bAbnormal = Not (CSng(vGoal) >= CSng(vCurrent))
        Case Else
            bAbnormal = True
    End Select
    If Err.Number <> 0 Then
        bAbnormal = False
    End If
    On Error GoTo ErrHandler
Else
    bAbnormal = False
End If


'-- Check if Overdue or Coming Due
'---------------------------------------------------
If Not GetOverDue(vFrequency, vDate, sUnitNumber, sOverDue) Then GoTo Cleanup
    
oRecResult.AddNew
oRecResult.Fields("Unit_Number") = sUnitNumber
oRecResult.Fields("Goal_Value") = vGoal
oRecResult.Fields("Current_Value") = vCurrent
If Not VarType(vDate) = vbEmpty Then
    oRecResult.Fields("Current_Date") = Format(vDate, "DD MMM YYYY")
Else
    oRecResult.Fields("Current_Date") = "None"
End If

If bAbnormal Then
    oRecResult.Fields("ABNORMAL") = "Y"
Else
    oRecResult.Fields("ABNORMAL") = "N"
End If
oRecResult.Fields("Due") = sOverDue
If Not VarType(vOperator) = vbEmpty Then
    oRecResult.Fields("OP_TYPE") = oRecGoals("GoalOpText").Value
Else
    oRecResult.Fields("OP_TYPE") = ""
End If

If oRecGoals.RecordCount <> 0 Then
    oRecResult("GOAL").Value = "Y"
Else
    oRecResult("GOAL").Value = "N"
End If

oRecResult.Update

GetPatientVitalStatus = True

GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepCPGGoals.GetPatientVitalStatus", "CHCSII_Reports", vbExclamation
    'Resume Next
Cleanup:
    Exit Function
End Function

Private Function GetPatientAnswerStatus(sUnitNumber As String, _
                                        sQuestionSeq As String, _
                                        oRecResponses As ADODB.Recordset, _
                                        oRecAnswers As ADODB.Recordset, _
                                        oRecResult As ADODB.Recordset) As Boolean

On Error GoTo ErrHandler

GetPatientAnswerStatus = False

Dim vAnswer As Variant
Dim vDate As Variant

'-- Responses
'-- unit_number, qresp_id, qdef_id, version, start_date_time, qresp_type. seq, answer
'-- Answers
'-- qdef_id, version, seq, text, answer_number
'-- cHGB_A1C_NCID

oRecResponses.Filter = ""
If oRecResponses.RecordCount <> 0 Then
    oRecResponses.Filter = "unit_number = " & sUnitNumber & " and seq = " & sQuestionSeq
End If

If oRecResponses.RecordCount <> 0 Then
    oRecAnswers.Filter = ""
    If oRecAnswers.RecordCount <> 0 Then
        On Error Resume Next
        oRecAnswers.Filter = "seq = " & sQuestionSeq & " and answer_number = " & oRecResponses.Fields("answer")
        'If Err <> 0 Then
        '    MsgBxARMd "5"
        'End If
    End If
    If oRecAnswers.RecordCount <> 0 Then
       vAnswer = oRecAnswers.Fields("text")
       vDate = oRecResponses("start_date_time").Value
    End If
End If

oRecResult.AddNew
oRecResult.Fields("Unit_Number") = sUnitNumber
If Not VarType(vDate) = vbEmpty Then
    oRecResult.Fields("Start_Date") = Format(vDate, "DD MMM YYYY")
Else
    oRecResult.Fields("Start_Date") = "None"
End If
'oRecResult.Fields("Start_Date") = vDate
oRecResult.Fields("Answer") = vAnswer
moRecHGB.Update

GetPatientAnswerStatus = True

GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepCPGGoals.GetPatientAnswerStatus", "CHCSII_Reports", vbExclamation
    'Resume Next
Cleanup:
    Exit Function
End Function


Private Function GetUnitNumberList(oPatientsColl As Collection, _
                                   sUnitNumberList As String) As Boolean

On Error GoTo ErrHandler

GetUnitNumberList = False

Dim i As Integer

sUnitNumberList = ""

For i = 1 To oPatientsColl.Count
    If i = oPatientsColl.Count Then
        sUnitNumberList = sUnitNumberList & oPatientsColl(i)
    Else
        sUnitNumberList = sUnitNumberList & oPatientsColl(i) & ","
    End If
Next

If sUnitNumberList = "" Then sUnitNumberList = "Null"

GetUnitNumberList = True
GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "CPGGlobal.GetUnitNumberList", "CHCSII_Reports", vbExclamation
Cleanup:
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
GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepCPGGoals.GetContextNCID", "CHCSII_Reports", vbExclamation
Cleanup:
    Exit Function
End Function


Private Function GetPatients(oReportItem As ReportItem, _
                             oPatientsColl As Collection) As Boolean

On Error GoTo ErrHandler
 
GetPatients = False


Set oPatientsColl = Nothing
Set oPatientsColl = New Collection

If oReportItem.Criteria1 = Patient Then

    oPatientsColl.Add oReportItem.CriteriaSel1.Item(1).UnitNumber
    
ElseIf oReportItem.Criteria1 = PCM Then

    If Not GetUnitNumberForPCMs(oReportItem.CriteriaSel1, oPatientsColl) Then GoTo Cleanup
    
ElseIf oReportItem.Criteria1 = PCMMulti Then
    
    If Not GetUnitNumberForPCMs(oReportItem.CriteriaSel1, oPatientsColl) Then GoTo Cleanup
    
End If

GetPatients = True

GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepCPGGoals.GetPatients", "CHCSII_Reports", vbExclamation
Cleanup:
   'Set oCrystal = Nothing
    'Set oRec = Nothing
End Function
Private Function GetUnitNumberForPCMs(oPCMCodes As Collection, oPatientUnitNumbers As Collection) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler
GetUnitNumberForPCMs = False

Dim sStmt As String
Dim icount As Integer
Dim oSQL As ICHCSII_SQL.ISqlOps
Dim sPCMCodes As String
Dim sPCMCode As String
Dim i As Integer

'-- Initialize
'-----------------------------------------------------------
Set oPatientUnitNumbers = Nothing
Set oPatientUnitNumbers = New Collection

'-- Put PCM Codes into a SQL Clause
'-----------------------------------------------------------
sPCMCodes = "("

For i = 1 To oPCMCodes.Count
    sPCMCodes = sPCMCodes & "'" & oPCMCodes.Item(1).PCMCode & "'"
    If i <> oPCMCodes.Count Then
        sPCMCodes = sPCMCodes & ","
    End If
Next

sPCMCodes = sPCMCodes & ")"


'-- Put PCM Codes into a SQL Clause
'-----------------------------------------------------------
Set oSQL = goDBConn.CHCSII_SQLOPS(Auto)

sStmt = "select mg.unit_number " & _
        " from mmi_generic_id mg " & _
        " where " & _
        " exists " & vbNewLine & _
        " (select 'x'" & vbNewLine & _
        " from CPGREGISTRY c" & vbNewLine & _
        " where" & vbNewLine & _
        " c.unit_number = mg.unit_number) and" & vbNewLine & _
        " mg.id_type_ncid = 14501724 and " & _
        " mg.id_value in " & sPCMCodes

oSQL.Execute sStmt

If oSQL.BOF And oSQL.EOF Then
    GetUnitNumberForPCMs = True
    GoTo Cleanup
End If

icount = 0
Do While Not oSQL.EOF
    oPatientUnitNumbers.Add oSQL.Value("unit_number")
    oSQL.MoveNext
Loop


GetUnitNumberForPCMs = True
GoTo Cleanup

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "RepCPGManagement.GetUnitNumberForPCMs", "CHCSII_Reports", vbExclamation)
Cleanup:
    Set oSQL = Nothing
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

GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepCPGGoals.GetDemographics", "CHCSII_Reports", vbExclamation
Cleanup:
    Set oRec = Nothing
    Exit Function
End Function


Private Function IsInRegistry(sUnitNumber As String, _
                              Optional sCPG_ID As String) As Boolean

On Error GoTo ErrHandler
IsInRegistry = False
 
Dim sStmt As String

sStmt = "Select " & _
            "'x'" & vbNewLine & _
        "From " & vbNewLine & _
            "cpgregistry c" & vbNewLine & _
        "Where " & vbNewLine
        
        If Not IsMissing(sCPG_ID) And sCPG_ID <> "" Then
            sStmt = sStmt & "c.cpg_id in (" & sCPG_ID & ") And " & vbNewLine
        End If
        
        sStmt = sStmt & "c.unit_number = " & sUnitNumber

            
goDBSqlOps.Execute sStmt

If Not goDBSqlOps.EOF Then
    IsInRegistry = True
End If

GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepCPGGoals.IsInRegistry", "CHCSII_Reports", vbExclamation
Cleanup:
    Exit Function
End Function



Private Function Initialize_RecordSets() As Boolean

On Error GoTo ErrHandler
 
Initialize_RecordSets = False

'-- Demographics
'------------------------------------------------
Set moRecDemographics = New ADODB.Recordset
moRecDemographics.CursorLocation = adUseClient

moRecDemographics.Fields.Append "Unit_Number", adVarChar, -1
moRecDemographics.Fields.Append "PCM", adVarChar, -1
moRecDemographics.Fields.Append "FMP_SSN", adVarChar, -1
moRecDemographics.Fields.Append "Full_NAME", adVarChar, -1
moRecDemographics.Open

'-- HGB A1C
'------------------------------------------------
Set moRecHGB = New ADODB.Recordset
moRecHGB.CursorLocation = adUseClient

moRecHGB.Fields.Append "Unit_Number", adVarChar, -1
moRecHGB.Fields.Append "Goal_Value", adVarChar, -1
moRecHGB.Fields.Append "Current_Value", adVarChar, -1
moRecHGB.Fields.Append "Current_Date", adVarChar, -1
moRecHGB.Fields.Append "ABNORMAL", adVarChar, -1
moRecHGB.Fields.Append "Due", adVarChar, -1
moRecHGB.Fields.Append "OP_TYPE", adVarChar, -1
moRecHGB.Fields.Append "GOAL", adVarChar, -1
moRecHGB.Open

If Not DuplicateDefinition(moRecHGB, moRecMicro) Then GoTo Cleanup
If Not DuplicateDefinition(moRecHGB, moRecHDL) Then GoTo Cleanup
If Not DuplicateDefinition(moRecHGB, moRecLDL) Then GoTo Cleanup
If Not DuplicateDefinition(moRecHGB, moRecTri) Then GoTo Cleanup
If Not DuplicateDefinition(moRecHGB, moRecChol) Then GoTo Cleanup
If Not DuplicateDefinition(moRecHGB, moRecBodyMass) Then GoTo Cleanup
If Not DuplicateDefinition(moRecHGB, moRecBloodSys) Then GoTo Cleanup
If Not DuplicateDefinition(moRecHGB, moRecBloodDia) Then GoTo Cleanup

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

'-- Education
'------------------------------------------------
Set moRecEd = New ADODB.Recordset
moRecEd.CursorLocation = adUseClient

moRecEd.Fields.Append "Unit_Number", adVarChar, 10
moRecEd.Fields.Append "Start_Date", adVarChar, 15
moRecEd.Fields.Append "Answer", adVarChar, 50
moRecEd.Open



Initialize_RecordSets = True

GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepCPGGoals.Initialize_RecordSets", "CHCSII_Reports", vbExclamation
Cleanup:
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

GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepCPGGoals.AddRecords", "CHCSII_Reports", vbExclamation
Cleanup:
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

