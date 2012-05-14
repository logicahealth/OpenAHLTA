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
Attribute VB_Name = "RepCPGFollowUpCareLetter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private moCollRecSet As Collection
Private moCollFormulas As Collection

Private Const cHGB_A1C_NCID = "8571"
Private Const cMICROALBUMIN = "31961,31968,31959,45555" '"31959"

Private Const cLDL = "6094"
Private Const cHDL = "6090"
Private Const cCholesterol = "21552"
Private Const cTrig = "6575,21552,7053"

',14622087)" '11146 ?'"14622262,14622263","31959" ','Private Const cLIPID_PROFILE = "14501603" '14503826,14503825'"6097"

Private msHGB_IEN As String
Private msLipid_IEN As String
Private msMicro_IEN As String

Private moRecHGB As ADODB.Recordset
Private moRecMicro As ADODB.Recordset
Private moRecLipid As ADODB.Recordset





Public Function GetGoals(oRecGoals As ADODB.Recordset, _
                         sUnitNumber As String, _
                         CPG_ID As String) As Boolean

On Error GoTo ErrHandler

GetGoals = False

Dim sStmt As String
Dim sRefCode As String
Dim sRefText As String
Dim iFrequency As Integer
Dim iPrecendence As Integer
Dim oRec As ADODB.Recordset


Set oRecGoals = Nothing
Set oRecGoals = New ADODB.Recordset
oRecGoals.Fields.Append "UNIT_NUMBER", adUnsignedBigInt
oRecGoals.Fields.Append "REFCODE", adUnsignedBigInt
oRecGoals.Fields.Append "FREQUENCY", adUnsignedBigInt
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
        "where po.poc_type=3 and po.category=99 and" & vbNewLine & _
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
        oRecGoals("REFCODE").Value = CLng(sRefCode)
        oRec.Fields("RefCode").Value = sRefCode
        oRec.Fields("RefText").Value = sRefText
        
        If Not GetFrequencyInDays(oRec.Fields("frequency"), iFrequency) Then GoTo CleanUp
        oRec.Fields("FrequencyInDays").Value = iFrequency
        oRecGoals("FREQUENCY").Value = iFrequency
        
        If Not GetFrequencyByPrecedence(oRec.Fields("frequency"), iPrecendence) Then GoTo CleanUp
        oRecGoals("PRECEDENCE").Value = iPrecendence
        
        oRecGoals("UNIT_NUMBER").Value = oRec.Fields("UNIT_NUMBER")
        
        oRecGoals.Update
        oRec.Update
        oRec.MoveNext
    Loop
End If

GetGoals = True
GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "CPGGlobal.GetGoals", "CHCSII_Reports", vbExclamation
CleanUp:
    Exit Function
End Function
Private Function CustomCreateReport(sReportName As String, _
                             oReport As CRAXDRT.Report, _
                             oCollFormulas As Collection, _
                             oCollRecSet As Collection, _
                             Optional sErrorMsg As String) As Boolean
    
    Dim oCrystal As Crystal
    Dim oCrystalText As TextObject
    
    On Error GoTo ErrHandler
    CustomCreateReport = False

    Set oCrystal = New Crystal
    
    If Not oCrystal.ReportOpen(sReportName, oReport) Then GoTo CleanUp
    If Not oCrystal.AssignRecordSets(oCollRecSet, oReport) Then GoTo CleanUp
    If Not oCrystal.AssignFormulas(oCollFormulas, oReport) Then GoTo CleanUp
    
    If TypeOf oReport.Sections.Item(13).ReportObjects.Item(1) Is TextObject Then
       Set oCrystalText = oReport.Sections.Item(13).ReportObjects.Item(1)
       oCrystalText.SetText sErrorMsg
    End If
    
    CustomCreateReport = True
    GoTo CleanUp

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ReportOps.GetMTForPCM_SQL", "ReportOps", vbExclamation)
CleanUp:
    Set oCrystalText = Nothing
    Set oCrystal = Nothing
End Function
Public Function BuildReport(sReportName As String, _
                            oReportItem As ReportItem, _
                            oReport As CRAXDRT.Report, _
                            oCollFormulas As Collection) As Boolean

    On Error GoTo ErrHandler
     
    BuildReport = False
    
    Dim oCrystal As Crystal
    Dim oPatientsRec As ADODB.Recordset
    Dim sSponsorUnitNumber As String
    Dim sUnitNumber As String
    Dim sErrorMsg As String
    Dim sStmt As String
    Dim sCPG_ID As String
    Dim i As Integer
    DoEvents


'-- Get Data
'----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection

    If Not Initialize_RecordSets() Then GoTo CleanUp

    '-- Build CPG list
    '--------------------------------------------------------
    For i = 1 To oReportItem.CriteriaSel2.Count
        If i = 1 Then
            sCPG_ID = oReportItem.CriteriaSel2.Item(i).CPG_ID
        Else
            sCPG_ID = sCPG_ID & "," & oReportItem.CriteriaSel2.Item(i).CPG_ID
        End If
    Next
    

    '-- Get Patient(s)
    '----------------------------------------------------------
    If Not GetPatientInfo(oReportItem, sCPG_ID, oPatientsRec) Then GoTo CleanUp
    If oPatientsRec.RecordCount = 0 Then
        If oReportItem.Criteria1 = PCM Then
            sErrorMsg = "The selected PCM has no patients included on the registry"
        Else
            sErrorMsg = "No Letters Need To Be Generated At This Time." '"Patient Must Have A PCM And Be Included On The Registry"
        End If
    Else
       sErrorMsg = "No Letters Need To Be Generated At This Time."
    End If
    
    oCollFormulas.Add CrystalQuote(sErrorMsg), "ErrorMessage"
    

    If Not GetFollowLabResults(sCPG_ID, oPatientsRec) Then GoTo CleanUp

    moCollRecSet.Add oPatientsRec, "letter"
    moCollRecSet.Add moRecHGB, "hgb"
    moCollRecSet.Add moRecMicro, "micro"
    moCollRecSet.Add moRecLipid, "lipid"

'-- Create Report Object
'-----------------------------------------------------------------------------

    If Not CustomCreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet, sErrorMsg) Then GoTo CleanUp

BuildReport = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepRegistryCandidates.BuildReport", "CHCSII_Reports", vbExclamation
    Resume Next
CleanUp:
    Set oCrystal = Nothing
    Set oPatientsRec = Nothing
End Function

Private Function GetFollowLabResults(sCPG_ID As String, oPatientsRec As ADODB.Recordset) As Boolean

On Error GoTo ErrHandler

GetFollowLabResults = False

Dim sStmt As String
Dim i As Integer
Dim vUnitNumber As Variant
Dim sHGB_Date As String
Dim sMicro_Date As String
Dim sLipid_Date As String
Dim oRecLab As ADODB.Recordset
Dim oRecGoal As ADODB.Recordset
Dim bNeedsHGBTest As Boolean
Dim bNeedsMicro As Boolean
Dim bNeedsLipid As Boolean
Dim sUnitNumberList As String
Dim sLabNCIDList As String


If Not GetUnitNumberList(oPatientsRec, "Unit_Number", sUnitNumberList) Then GoTo CleanUp

sLabNCIDList = cHGB_A1C_NCID & "," & cMICROALBUMIN & "," & cLDL & "," & cHDL & "," & cCholesterol

If Not GetLabResultSummary(oRecLab, sUnitNumberList, sLabNCIDList) Then GoTo CleanUp

If Not oPatientsRec.RecordCount = 0 Then
    oPatientsRec.MoveFirst
    Do While Not oPatientsRec.EOF
        If Not GetGoals(oRecGoal, oPatientsRec("unit_number").Value, sCPG_ID) Then GoTo CleanUp
        
        If oRecGoal.RecordCount = 0 Then
            oPatientsRec.Delete
        Else
    
            '--HGB A1C
            '----------------------------------
            If Not SetLastLabResultDate(oPatientsRec, oRecGoal, oRecLab, cHGB_A1C_NCID, moRecHGB, bNeedsHGBTest) Then GoTo CleanUp
    
            '--Microalbumin
            '----------------------------------
            If Not SetLastLabResultDate(oPatientsRec, oRecGoal, oRecLab, cMICROALBUMIN, moRecMicro, bNeedsMicro) Then GoTo CleanUp
    
            '--LIPID
            '----------------------------------
            If Not SetLastLabResultDateForLipid(oPatientsRec, oRecGoal, oRecLab, moRecLipid, bNeedsLipid) Then GoTo CleanUp
    
            If Not (bNeedsHGBTest Or bNeedsMicro Or bNeedsLipid) Then
                oPatientsRec.Delete
            End If
            
            oRecGoal.Filter = ""
        End If
        oPatientsRec.MoveNext
    Loop
    
    If oPatientsRec.RecordCount = 0 Then
        Set oPatientsRec = Nothing
        If Not GetEmptyRecordSet(oPatientsRec) Then GoTo CleanUp
    End If

End If

GetFollowLabResults = True
GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepCPGFollowUpCareLetter.GetFollowLabResults", "CHCSII_Reports", vbExclamation
CleanUp:
    Exit Function
End Function


Private Function BuildFilter(sField As String, sFieldData, Optional sField2 As String = "x", Optional sFieldData2) As String
'-- sField - Name of field in recordset
'-- sFieldData - Data to filter field by
'-- sField2 - Name of field in recordset (assumed to be unit_number)
'-- sFieldData2 - Data to filter field by

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
    
GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepCPGGoals.BuildFilter", "CHCSII_Reports", vbExclamation
CleanUp:
    Set cData = Nothing
    Exit Function
    
End Function

Private Function GetOverDue(oRecGoal As ADODB.Recordset, _
                            oRecLab As ADODB.Recordset, _
                            oRecPatient As ADODB.Recordset, _
                            sLAB_NCIDs As String, _
                            sLastDate As String, _
                            bOverDue As Boolean) As Boolean

bOverDue = True

Dim vDiffDate As Variant

oRecGoal.Filter = ""
If oRecGoal.RecordCount <> 0 Then
    oRecGoal.Filter = BuildFilter("RefCode", sLAB_NCIDs)
    oRecGoal.Sort = "unit_number desc, precedence asc"
End If

If oRecGoal.RecordCount = 0 Then
    bOverDue = False
ElseIf oRecLab.RecordCount <> 0 Then
    oRecLab.Filter = BuildFilter("Lab_Id", sLAB_NCIDs, "unit_number", oRecPatient.Fields("unit_number"))
    oRecLab.Sort = "Collect_Date"

    If oRecLab.RecordCount = 0 Then
        sLastDate = "Never"
    Else
        If oRecGoal.Fields("frequency") = 32767 Then
            bOverDue = False
        ElseIf oRecGoal.Fields("frequency") = 1 Then
            If (CDate(CStr(Date) & " 0:00:00 AM") < oRecLab.Fields("Collect_Date") And _
                oRecLab.Fields("Collect_Date") < CDate(CStr(Date) & " 11:59:59 PM")) Then
                bOverDue = False
            Else
                sLastDate = Format(oRecLab.Fields("Collect_Date"), "DD MMM YYYY")
            End If
        ElseIf oRecGoal.Fields("frequency") = 2 Then
            If CompletedLastVisit(oRecPatient("Unit_Number"), oRecLab.Fields("Collect_Date")) Then
                bOverDue = False
            Else
                sLastDate = Format(oRecLab.Fields("Collect_Date"), "DD MMM YYYY")
            End If
        Else
            vDiffDate = oRecLab.Fields("Collect_Date") + oRecGoal.Fields("frequency") + 60
            If vDiffDate > Now Then
                bOverDue = False
            Else
                sLastDate = Format(oRecLab.Fields("Collect_Date"), "DD MMM YYYY")
            End If
        End If
    End If
Else
    sLastDate = "Never"
End If

GetOverDue = True

End Function

Private Function GetPatientInfo(oReportItem As ReportItem, _
                               sCPGs As String, _
                               oRec As ADODB.Recordset) As Boolean
On Error GoTo ErrHandler

GetPatientInfo = False

Dim sStmt As String
Dim sFromSQL As String
Dim sWhereSQL As String

    If Not GetMTForPCM_SQL(oReportItem, sFromSQL, sWhereSQL) Then GoTo CleanUp

'-- Select
    sStmt = "Select distinct " & vbNewLine & _
                "1 as job_id, " & vbNewLine & _
                "mn.unit_number, " & vbNewLine & _
                "ma.street_1, " & vbNewLine & _
                "ma.street_2, " & vbNewLine & _
                "ma.city, " & vbNewLine & _
                "ma.state, " & vbNewLine & _
                "ma.post_code, " & vbNewLine & _
                "mn.first_name, " & vbNewLine & _
                "mn.middle_name, " & vbNewLine & _
                "mn.last_name, " & vbNewLine & _
                "'               ' as hgb_Date, " & vbNewLine & _
                "'               ' as micro_Date, " & vbNewLine & _
                "'               ' as lipid_Date, " & vbNewLine & _
                "p2.Name " & vbNewLine

'-- From
    sStmt = sStmt & _
            "From " & vbNewLine & _
                "mmi_address ma, " & vbNewLine & _
                "mmi_name mn, " & vbNewLine & _
                "provider p2, " & vbNewLine & _
                "mmi_generic_id id2, " & vbNewLine & _
                "mmi_id id " & vbNewLine

    '-- Additional From Clause
    sStmt = sStmt & sFromSQL


'-- Where
    sStmt = sStmt & _
        "Where " & vbNewLine & _
            "ma.unit_number (+) = id.unit_number and " & vbNewLine & _
            "mn.name_type_cid = 187 and " & vbNewLine & _
            "mn.unit_number = id.unit_number and " & vbNewLine & _
            "p2.PCM_Code = id2.id_value and " & vbNewLine & _
            "id2.id_type_ncid = 14501724 and " & vbNewLine & _
            "id2.unit_number = id.unit_number and " & vbNewLine & _
            "exists " & vbNewLine & _
            "(select 'x' from cpgregistry c where c.cpg_id in (" & sCPGs & ") AND  c.unit_number = id.unit_number) " & vbNewLine

    '-- Check if patient has records at specified MTF or is patient of specified PCM
    sStmt = sStmt & " and " & sWhereSQL

Set oRec = goDBDAS.OpenRecordset(sStmt)

GetPatientInfo = True
GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepRegistryCandidates.GetPatientInfo", "CHCSII_Reports", vbExclamation
CleanUp:
    Exit Function
End Function

Private Function GetLabResultSummary(oRec As ADODB.Recordset, _
                                     sUnitNumber As String, _
                                     Optional sLabNCID As String) As Boolean


On Error GoTo ErrHandler

GetLabResultSummary = False

Dim sStmt As String


'-- SQL
'-----------------------------------------------------
'-- Fix SCR 27378
    sStmt = "select /*+ ORDERED */ /*+ INDEX (e event_x1) (o observation_pk) */ " & vbNewLine & _
                "e.unit_number, " & _
                "e.event_start_gmtime as Collect_Date, " & _
                "o.observation_ncid as Lab_Id " & vbNewLine & _
            "from " & _
                "event e, " & _
                "observation o " & vbNewLine & _
            "where " & vbNewLine & _
                "o.observation_ncid in (" & sLabNCID & ") and " & vbNewLine & _
                "o.event_id = e.event_id and " & vbNewLine & _
                "e.event_type_ncid = 485 and " & vbNewLine & _
                "e.unit_number in (" & sUnitNumber & ") " & vbNewLine & _
            "order by 2, 1 desc"
    

Set oRec = goDBDAS.OpenRecordset(sStmt)

GetLabResultSummary = True
GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepCPGFollowUpCareLetter.GetLabResultSummary", "CHCSII_Reports", vbExclamation
CleanUp:
    Exit Function
End Function

 
Private Function Initialize_RecordSets() As Boolean

On Error GoTo ErrHandler
Initialize_RecordSets = False


'-- HGB A1C
'------------------------------------------------
Set moRecHGB = New ADODB.Recordset
moRecHGB.CursorLocation = adUseClient

moRecHGB.Fields.Append "Unit_Number", adLongVarWChar, 12
moRecHGB.Fields.Append "lab_date", adVarChar, 12
moRecHGB.Open

'-- Micro
'------------------------------------------------
Set moRecMicro = New ADODB.Recordset
moRecMicro.CursorLocation = adUseClient

moRecMicro.Fields.Append "Unit_Number", adLongVarWChar, 12
moRecMicro.Fields.Append "lab_date", adVarChar, 12
moRecMicro.Open

'-- Lipid
'------------------------------------------------
Set moRecLipid = New ADODB.Recordset
moRecLipid.CursorLocation = adUseClient

moRecLipid.Fields.Append "Unit_Number", adLongVarWChar, 12
moRecLipid.Fields.Append "lab_date", adVarChar, 12
moRecLipid.Open



Initialize_RecordSets = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepCPGFollowUpCareLetter.Initialize_RecordSets", "CHCSII_Reports", vbExclamation
CleanUp:
End Function

Private Function SetLastLabResultDate(oRecPatient As ADODB.Recordset, _
                                      oRecGoal As ADODB.Recordset, _
                                      oRecLab As ADODB.Recordset, _
                                      sLAB_NCID As String, _
                                      oRecResult As ADODB.Recordset, _
                                      bNeedsTest As Boolean) As Boolean

On Error GoTo ErrHandler

SetLastLabResultDate = False

Dim iFrequency As Integer
Dim vDiffDate As Variant
Dim sLastDate As String
Dim bOverDue As Boolean
                                      

If Not GetOverDue(oRecGoal, oRecLab, oRecPatient, sLAB_NCID, sLastDate, bOverDue) Then GoTo CleanUp
                                    
If bOverDue Then
    oRecResult.AddNew
    oRecResult.Fields("unit_number").Value = oRecPatient.Fields("unit_number")
    oRecResult.Fields("lab_date").Value = sLastDate
End If


bNeedsTest = bOverDue

SetLastLabResultDate = True
GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepCPGFollowUpCareLetter.SetLastLabResultDate", "CHCSII_Reports", vbExclamation
CleanUp:
    Exit Function
End Function

Private Function SetLastLabResultDateForLipid(oRecPatient As ADODB.Recordset, _
                                              oRecGoal As ADODB.Recordset, _
                                              oRecLab As ADODB.Recordset, _
                                              oRecResult As ADODB.Recordset, _
                                              bNeedsTest As Boolean) As Boolean

On Error GoTo ErrHandler

SetLastLabResultDateForLipid = False

Dim iFrequency As Integer
Dim vDiffDate As Variant
Dim vLDL_Date As Variant
Dim sLDL_Date As String
Dim bLDL_Overdue As Boolean
Dim vTotalChol_Date As Variant
Dim sTotalChol_Date As String
Dim bTotalChol_Overdue As Boolean
Dim vHDL_Date As Variant
Dim sHDL_Date As String
Dim bHDL_Overdue As Boolean
Dim vOldestDueDate As Variant
Dim vCurrentDate As Variant
Dim sDateFieldName As String
Dim sLastDate As String
Dim vTrig_Date As Variant
Dim sTrig_Date As String
Dim bTrig_Overdue As Boolean

'-- Initialize Dates
'-----------------------------------------------------------

bNeedsTest = False
vLDL_Date = CVDate("01/01/1970")
vTotalChol_Date = vLDL_Date
vHDL_Date = vLDL_Date

'-- Check LDL first
'-----------------------------------------------------------

If Not GetOverDue(oRecGoal, oRecLab, oRecPatient, cLDL, sLDL_Date, bLDL_Overdue) Then GoTo CleanUp

'-- Check Cholesterol
'-----------------------------------------------------------

If Not GetOverDue(oRecGoal, oRecLab, oRecPatient, cCholesterol, sTotalChol_Date, bTotalChol_Overdue) Then GoTo CleanUp

'-- Check HDL
'-------------------------------------------
If Not GetOverDue(oRecGoal, oRecLab, oRecPatient, cHDL, sHDL_Date, bHDL_Overdue) Then GoTo CleanUp

'-- Check Trig
'-------------------------------------------
If Not GetOverDue(oRecGoal, oRecLab, oRecPatient, cTrig, sTrig_Date, bTrig_Overdue) Then GoTo CleanUp

'-- Determine if Lipid Profile is Overdue
'---------------------------------------------
If Not (bLDL_Overdue = False And bTotalChol_Overdue = False And bHDL_Overdue = False And bTrig_Overdue = False) Then

    If sLDL_Date = "Never" Or sHDL_Date = "Never" Or sTotalChol_Date = "Never" Or sTrig_Date = "Never" Then
        sLastDate = "Never"
    Else
        If sLDL_Date <> "" Then
            sLastDate = sLDL_Date
        ElseIf sHDL_Date <> "" Then
            sLastDate = sHDL_Date
        ElseIf sTotalChol_Date <> "" Then
            sLastDate = sTotalChol_Date
        ElseIf sTrig_Date <> "" Then
            sLastDate = sTrig_Date
        End If
    End If
    oRecResult.AddNew
    oRecResult.Fields("unit_number").Value = oRecPatient.Fields("unit_number")
    oRecResult.Fields("lab_date").Value = sLastDate
    oRecResult.Update
    bNeedsTest = True
End If


SetLastLabResultDateForLipid = True
GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepCPGFollowUpCareLetter.SetLastLabResultDateForLipid", "CHCSII_Reports", vbExclamation
CleanUp:
    Exit Function
End Function




Private Sub Class_Initialize()
    '--- SCR-ARM-16148 MSolano 9/10/01; COMMON TO MULTIPLE ARM-SCRs
    '--- Init routine should be called from first top level creatable coclass
    '--- It can be called multiple times, but will actually run only once.
    '--- Instrument all creatable classes in case dll's multiple entry
    '--- points are actually used
     Call ARM_InitAllTransactionsForCHCSIIReports
End Sub


