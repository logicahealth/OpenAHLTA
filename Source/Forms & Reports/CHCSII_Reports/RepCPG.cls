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
Attribute VB_Name = "RepCPG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private moCollRecSet As Collection
Private moCollFormulas As Collection

Public Function BuildReport(sReportName As String, _
                            oReportItem As ReportItem, _
                            oReport As CRAXDRT.Report, _
                            oCollFormulas As Collection) As Boolean
 
On Error GoTo ErrHandler
 
BuildReport = False

Select Case oReportItem.Report
    Case CPG_QuestionnaireData
        If Not BuildCPGQuestionnaireData(sReportName, oReportItem, oReport, oCollFormulas) Then GoTo CleanUp
End Select

BuildReport = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepRegistryCandidates.BuildReport", "CHCSII_Reports", vbExclamation
CleanUp:
    
End Function


Public Function BuildCPGQuestionnaireData(sReportName As String, _
                                          oReportItem As ReportItem, _
                                          oReport As CRAXDRT.Report, _
                                          oCollFormulas As Collection) As Boolean
 
On Error GoTo ErrHandler
 
BuildCPGQuestionnaireData = False

Dim oRec As ADODB.Recordset
Dim sStmt As String

'-- Get Data
'----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection


'-- Select
'------------
    sStmt = "Select " & vbNewLine & _
                "/*+ ORDERED INDEX (id MMI_ID_PK) (ar ARESP_PK) (qd qdef_pk) (mn mmi_name_x5) (pm mmi_generic_id_x3) (P PROVIDER_PK) */ " & vbNewLine & _
                "p.NAME as PCM_Name, " & vbNewLine & _
                "Trim(mn.last_name ||', ' || mn.first_name || ' ' ||" & vbNewLine & _
                "mn.middle_name) as Full_Name," & vbNewLine & _
                "cr.research_id as RegistryId," & vbNewLine & _
                "ar.seq as QuestionNumber," & vbNewLine & _
                "'                                    ' as FinaL_Answer," & vbNewLine & _
                "qd.name quest_name," & vbNewLine & _
                "qr.start_date_time," & vbNewLine & _
                "ar.answer," & vbNewLine & _
                "ad.qdef_id," & vbNewLine & _
                "ad.version," & vbNewLine & _
                "ad.adef_type," & vbNewLine & _
                "'A' SelectType" & vbNewLine


'-- From
'------------
    sStmt = sStmt & _
            "From " & vbNewLine & _
                "mmi_id id," & vbNewLine & _
                "cpgregistry cr," & vbNewLine & _
                "qresp qr," & vbNewLine & _
                "aresp ar," & vbNewLine & _
                "qdef qd," & vbNewLine & _
                "adef ad," & vbNewLine & _
                "mmi_name mn," & vbNewLine & _
                "mmi_generic_id pm," & vbNewLine & _
                "Provider p" & vbNewLine

'-- Where
'------------
    sStmt = sStmt & _
        "Where " & vbNewLine & _
            "p.pcm_code = pm.id_value and" & vbNewLine & _
            "pm.id_type_ncid = 14501724 and" & vbNewLine & _
            "pm.unit_number = id.unit_number and" & vbNewLine & _
            "mn.name_type_cid = 187 and" & vbNewLine & _
            "mn.unit_number = id.unit_number and" & vbNewLine & _
            "ad.seq = ar.seq and" & vbNewLine & _
            "ad.version = qr.version and" & vbNewLine & _
            "ad.qdef_id =  qr.qdef_id and" & vbNewLine & _
            "qd.version = qr.version and " & vbNewLine & _
            "qd.qdef_id = qr.qdef_id and" & vbNewLine & _
            "ar.qresp_id = qr.qresp_id and" & vbNewLine

    sStmt = sStmt & _
            "qr.qresp_id in" & vbNewLine & _
            "(select max(qr2.qresp_id)" & vbNewLine & _
            "from qresp qr2" & vbNewLine & _
            "Where" & vbNewLine & _
            "qr2.end_date_time in" & vbNewLine & _
            "(select max(qr3.end_date_time) " & vbNewLine & _
            "from qresp qr3" & vbNewLine & _
            "where " & vbNewLine & _
            "qr3.status in (3,4,5) and" & vbNewLine & _
            "qr3.version = " & CStr(oReportItem.CriteriaSel2.Item(1).Version) & " and" & vbNewLine & _
            "qr3.qdef_id = " & CStr(oReportItem.CriteriaSel2.Item(1).QDefID) & " and" & vbNewLine & _
            "qr3.unit_number = qr2.unit_number) and" & vbNewLine & _
            "qr2.status in (3,4,5) and" & vbNewLine & _
            "qr2.version = " & CStr(oReportItem.CriteriaSel2.Item(1).Version) & " and" & vbNewLine & _
            "qr2.qdef_id = " & CStr(oReportItem.CriteriaSel2.Item(1).QDefID) & " and" & vbNewLine & _
            "qr2.unit_number = id.unit_number) and" & vbNewLine & _
            "qr.unit_number = id.unit_number and" & vbNewLine & _
            "cr.unit_number = id.unit_number and" & vbNewLine & _
            "id.unit_number in" & vbNewLine & _
            "(select /*+ index (mgi mmi_generic_id_x3) */ mgi.unit_number" & vbNewLine & _
            "from  mmi_generic_id mgi" & vbNewLine & _
            "where mgi.id_type_ncid = 14501724 and" & vbNewLine & _
            "mgi.id_value = '" & oReportItem.CriteriaSel1.Item(1).PCMCode & "')" & vbNewLine

 sStmt = sStmt & "Union" & vbNewLine
 
'-- Select
'------------
    sStmt = sStmt & "Select " & vbNewLine & _
                "/*+ ORDERED */ " & vbNewLine & _
                "/*+ INDEX (id MMI_ID_PK) (ar ARESP_PK) (qd qdef_pk) " & vbNewLine & _
                "(mn mmi_name_x5) (pm mmi_generic_id_x3) (P PROVIDER_PK) */ " & vbNewLine & _
                "p.NAME as PCM_Name, " & vbNewLine & _
                "Trim(mn.last_name ||', ' || mn.first_name || ' ' ||" & vbNewLine & _
                "mn.middle_name) as Full_Name," & vbNewLine & _
                "cr.research_id as RegistryId," & vbNewLine & _
                "ar.seq as QuestionNumber," & vbNewLine & _
                "'                                    ' as FinaL_Answer," & vbNewLine & _
                "qd.name quest_name," & vbNewLine & _
                "qr.start_date_time," & vbNewLine & _
                "TO_CHAR(ar.choice)," & vbNewLine & _
                "ad.qdef_id, " & vbNewLine & _
                "ad.version," & vbNewLine & _
                "ad.adef_type," & vbNewLine & _
                "'Q' SelectType" & vbNewLine


   
    
'-- From
'------------
    sStmt = sStmt & _
            "From " & vbNewLine & _
                "mmi_id id," & vbNewLine & _
                "cpgregistry cr," & vbNewLine & _
                "qresp qr," & vbNewLine & _
                "qresp_choices ar," & vbNewLine & _
                "qdef qd," & vbNewLine & _
                "adef ad," & vbNewLine & _
                "mmi_name mn," & vbNewLine & _
                "mmi_generic_id pm," & vbNewLine & _
                "Provider p" & vbNewLine

'-- Where
'------------
    sStmt = sStmt & _
        "Where " & vbNewLine & _
            "p.pcm_code = pm.id_value and" & vbNewLine & _
            "pm.id_type_ncid = 14501724 and" & vbNewLine & _
            "pm.unit_number = id.unit_number and" & vbNewLine & _
            "mn.name_type_cid = 187 and" & vbNewLine & _
            "mn.unit_number = id.unit_number and" & vbNewLine & _
            "ad.seq = ar.seq and" & vbNewLine & _
            "ad.version = qr.version and" & vbNewLine & _
            "ad.qdef_id =  qr.qdef_id and" & vbNewLine & _
            "qd.version = qr.version and " & vbNewLine & _
            "qd.qdef_id = qr.qdef_id and" & vbNewLine & _
            "ar.qresp_id = qr.qresp_id and" & vbNewLine

    sStmt = sStmt & _
            "qr.qresp_id in" & vbNewLine & _
            "(select max(qr2.qresp_id)" & vbNewLine & _
            "from qresp qr2" & vbNewLine & _
            "Where" & vbNewLine & _
            "qr2.end_date_time in" & vbNewLine & _
            "(select max(qr3.end_date_time) " & vbNewLine & _
            "from qresp qr3" & vbNewLine & _
            "where " & vbNewLine & _
            "qr3.status in (3,4,5) and" & vbNewLine & _
            "qr3.version = " & CStr(oReportItem.CriteriaSel2.Item(1).Version) & " and" & vbNewLine & _
            "qr3.qdef_id = " & CStr(oReportItem.CriteriaSel2.Item(1).QDefID) & " and" & vbNewLine & _
            "qr3.unit_number = qr2.unit_number) and" & vbNewLine & _
            "qr2.status in (3,4,5) and" & vbNewLine & _
            "qr2.version = " & CStr(oReportItem.CriteriaSel2.Item(1).Version) & " and" & vbNewLine & _
            "qr2.qdef_id = " & CStr(oReportItem.CriteriaSel2.Item(1).QDefID) & " and" & vbNewLine & _
            "qr2.unit_number = id.unit_number) and" & vbNewLine & _
            "qr.unit_number = id.unit_number and" & vbNewLine & _
            "cr.unit_number = id.unit_number and" & vbNewLine & _
            "id.unit_number in" & vbNewLine & _
            "(select /*+ index (mgi mmi_generic_id_x3) */ mgi.unit_number" & vbNewLine & _
            "from  mmi_generic_id mgi" & vbNewLine & _
            "where mgi.id_type_ncid = 14501724 and" & vbNewLine & _
            "mgi.id_value = '" & oReportItem.CriteriaSel1.Item(1).PCMCode & "')" & vbNewLine
            
'-- Order By
'------------
    sStmt = sStmt & "Order by 1, 2, 3, 4  "

    Set oRec = goDBDAS.OpenRecordset(sStmt)
    
    
'-- Verify Records were returned
'----------------------------------------------------------------------------
If oRec.RecordCount = 0 Then
    MsgBxARMd "No records were returned that matched patients for this PCM and Questionnaire.", _
           vbOKOnly + vbInformation, "No Data Returned"
    GoTo CleanUp
End If

'-- Add Recordset to list
'----------------------------------------------------------------------------
    moCollRecSet.Add oRec, "answers"

    Set goRecQuestData = oRec

    If Not UpdateQuestionnaireRecordSet(oRec) Then GoTo CleanUp

'-- Get Text for Answer
'-------------------------------------------------------------
    'If Not UpdateQuestionnaireFields(oRec) Then GoTo CleanUp
    'If Not UpdateAnswerFields(oRec) Then GoTo CleanUp

    
'-- Create Report Object
'-----------------------------------------------------------------------------
    If Not CreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo CleanUp

BuildCPGQuestionnaireData = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepCPG.BuildCPGQuestionnaireData", "CHCSII_Reports", vbExclamation
CleanUp:
    Set oRec = Nothing
    
End Function


Private Function GetChoiceText(sQDEF_ID As String, sVersion As String, sSeq As String, sChoice As String) As String

On Error GoTo ErrHandler

    Dim sStmt As String
    Dim oRec As ADODB.Recordset

    
    If Trim(sChoice) = "" Then
        GetChoiceText = ""
        GoTo CleanUp
    End If
    sStmt = "select" & vbNewLine & _
                "Question_Text" & vbNewLine & _
            "from" & vbNewLine & _
                "qdef_choices" & vbNewLine & _
            "where" & vbNewLine & _
                "choice=" & sChoice & " and" & vbNewLine & _
                "seq=" & sSeq & " and" & vbNewLine & _
                "version=" & sVersion & " and" & vbNewLine & _
                "qdef_id=" & sQDEF_ID
    
    Set oRec = goDBDAS.OpenRecordset(sStmt)
    
    If Not oRec.RecordCount = 0 Then
        GetChoiceText = oRec("Question_Text").Value
    Else
        GetChoiceText = ""
    End If

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepCPG.GetChoiceText", "CHCSII_Reports", vbExclamation
CleanUp:
   Set oRec = Nothing
End Function

Private Function UpdateQuestionnaireRecordSet(oRec As ADODB.Recordset) As Boolean

On Error GoTo ErrHandler
UpdateQuestionnaireRecordSet = False

If oRec.RecordCount <> 0 Then
    oRec.MoveFirst
    
    While oRec.EOF <> True

        If oRec("SelectType").Value = "A" Then
            If oRec("adef_type").Value = 3 Or oRec("adef_type").Value = 4 Then
                oRec("Final_Answer").Value = oRec("answer")
            Else
                oRec("Final_Answer").Value = GetChoiceText(oRec("qdef_id"), oRec("version"), oRec("QuestionNumber"), oRec("answer"))
            End If
        Else
            oRec("Final_Answer").Value = GetChoiceText(oRec("qdef_id"), oRec("version"), oRec("QuestionNumber"), oRec("answer"))
        End If
        oRec.Update
        oRec.MoveNext
    Wend
    
End If

UpdateQuestionnaireRecordSet = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepCPG.UpdateQuestionnaireRecordSet", "CHCSII_Reports", vbExclamation
CleanUp:
    Set oRec = Nothing
End Function

Private Sub Class_Initialize()
    '--- SCR-ARM-16148 MSolano 9/10/01; COMMON TO MULTIPLE ARM-SCRs
    '--- Init routine should be called from first top level creatable coclass
    '--- It can be called multiple times, but will actually run only once.
    '--- Instrument all creatable classes in case dll's multiple entry
    '--- points are actually used
     Call ARM_InitAllTransactionsForCHCSIIReports
End Sub


