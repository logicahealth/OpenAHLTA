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
Attribute VB_Name = "FormDD2161"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private moCollRecSet As Collection
Private moCollFormulas As Collection

Private moPatientDemographics As PatientDemographics

Private Sub Class_Initialize()
    Set moPatientDemographics = New PatientDemographics
End Sub

Private Sub Class_Terminate()
    Set moPatientDemographics = Nothing
End Sub


Public Function BuildDD2161(oConsultFormItem As ConsultFormItem, _
                           sPatientUnitNumber As String, _
                           oReport As CRAXDRT.Report) As Boolean

On Error GoTo ErrHandler

BuildDD2161 = False

Dim oCrystal As Crystal
Dim oRec As ADODB.Recordset
Dim sSponsorUnitNumber As String
Dim sUnitNumber As String
DoEvents

'-- Get Data
'----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection
    
    '-- Patient Demographics
    '----------------------------------------------------------
    If Not moPatientDemographics.RecordGetDemographics(sPatientUnitNumber, oRec) Then GoTo CleanUp
    moCollRecSet.Add oRec, "demographics"

    '-- Sponsor Demographics
    '----------------------------------------------------------
    If oRec.Fields.Item("FMP").Value = "20" Then
        sSponsorUnitNumber = sPatientUnitNumber
    Else
        If gobjshared.IsAppMode(modeCHCSI_GUI) Then
            sSponsorUnitNumber = oRec.Fields.Item("SPONSOR_UNIT_NUMBER").Value
        Else
            sSponsorUnitNumber = moPatientDemographics.GetUnitNumber(oRec.Fields.Item("Sponsor_SSN").Value)
        End If
    End If
    If Not moPatientDemographics.RecordGetDemographics(sSponsorUnitNumber, oRec) Then GoTo CleanUp
    moCollRecSet.Add oRec, "sponsor"

    '-- Create DD2161 Data Record
    '----------------------------------------------------------
    If Not CreateDD2161Template(oRec) Then GoTo CleanUp
    oRec.AddNew
    oRec.Fields("Job_Id").Value = 1
    oRec.Fields("TodaysDateTime").Value = Format(Now, gsDefaultFormat & "  HHMM")
    oRec.Fields("ToProvider").Value = oConsultFormItem.ConsultTo
    oRec.Fields("DateOfRequest").Value = Format$(oConsultFormItem.DateOfRequest, gsDefaultFormat)
    oRec.Fields("Facility").Value = oConsultFormItem.Facility
    oRec.Fields("FromProvider").Value = oConsultFormItem.From
    oRec.Fields("ProvisionalDiagnosis").Value = oConsultFormItem.ProvisionalDiagnosis
    oRec.Fields("ReasonForRequest").Value = oConsultFormItem.ReasonForRequest
    oRec.Fields("Urgency").Value = oConsultFormItem.Urgency
    oRec.Update
    moCollRecSet.Add oRec, "dd2161"

'-- Create Report Object
'-----------------------------------------------------------------------------
    Set oCrystal = New Crystal
    Set oReport = Nothing
    If Not oCrystal.ReportOpen("DD2161.rpt", oReport) Then GoTo CleanUp
    If Not oCrystal.AssignRecordSets(moCollRecSet, oReport) Then GoTo CleanUp

BuildDD2161 = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "FormDD2161.BuildDD2161", "CHCSII_Forms", vbExclamation
CleanUp:
    Set oCrystal = Nothing
    Set oRec = Nothing
    
End Function


Private Function CreateDD2161Template(oRec As ADODB.Recordset) As Boolean

Dim sStmt As String
Dim i As Integer
Dim i2 As Integer
Dim oRec1 As ADODB.Recordset

    CreateDD2161Template = False
    On Error GoTo ErrHandler

    If gobjshared.UseSQLServer Or gobjshared.IsAppMode(modeCHCSI_GUI) Then
       Set oRec = New ADODB.Recordset
       oRec.Fields.Append "job_id", adLongVarWChar, -1, 132
       oRec.Fields.Append "TodaysDateTime", adLongVarWChar, -1, 132
       oRec.Fields.Append "ToProvider", adLongVarWChar, -1, 132
       oRec.Fields.Append "DateOfRequest", adLongVarWChar, -1, 132
       oRec.Fields.Append "Facility", adLongVarWChar, -1, 132
       oRec.Fields.Append "FromProvider", adLongVarWChar, -1, 132
       oRec.Fields.Append "ProvisionalDiagnosis", adLongVarWChar, -1, 132
       oRec.Fields.Append "ReasonForRequest", adLongVarWChar, -1, 132
       oRec.Fields.Append "Urgency", adLongVarWChar, -1, 132
       oRec.Open
       CreateDD2161Template = True
       GoTo CleanUp
    End If

    '-- Select
    sStmt = _
        "select" & vbNewLine & _
            "1 as job_id," & vbNewLine & _
            "varchar_fld as TodaysDateTime," & vbNewLine & _
            "varchar_fld as ToProvider," & vbNewLine & _
            "varchar_fld as DateOfRequest," & vbNewLine & _
            "varchar_fld as Facility," & vbNewLine & _
            "varchar_fld as FromProvider," & vbNewLine & _
            "varchar_fld as ProvisionalDiagnosis," & vbNewLine & _
            "varchar_fld as ReasonForRequest," & vbNewLine & _
            "varchar_fld as Urgency" & vbNewLine
    
    '-- From
    sStmt = sStmt & _
            "from Form_Template" & vbNewLine
    
    Set oRec1 = goDBDAS.OpenRecordset(sStmt)
    
    Set oRec = New ADODB.Recordset
    For i = 0 To oRec1.Fields.Count - 1
        oRec.Fields.Append oRec1(i).Name, oRec1(i).Type, oRec1(i).DefinedSize, oRec1(i).Attributes
    Next
    oRec.Open
    

CreateDD2161Template = True
GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "FormDD2161.CreateDD2161Template", "CHCSII_Forms", vbExclamation
CleanUp:
    Set oRec1 = Nothing
    Exit Function
End Function


Public Function BuildDD2161Complete(sConsult_Id As String, _
                                   sConsultationReportText As String, _
                                   oReport As CRAXDRT.Report) As Boolean

On Error GoTo ErrHandler

BuildDD2161Complete = False

Dim oCrystal As Crystal
Dim oRec As ADODB.Recordset
Dim sSponsorUnitNumber As String
Dim sClinicNCID As String
Dim sUnitNumber As String
Dim bDefault As Boolean
DoEvents

'-- Get Data
'----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection

    '-- Get Consult Record
    '----------------------------------------------------------
    If Not GetConsultRecord(sConsult_Id, sConsultationReportText, oRec) Then GoTo CleanUp

    '-- Check if consult is complete
    If oRec.Fields.Count <= 1 Then
        MsgBxARMd "Consult not found in database." & vbCrLf & "Consult Id: " & sConsult_Id, vbExclamation
        GoTo CleanUp
    End If
    moCollRecSet.Add oRec, "dd2161"
    
    '-- Patient Demographics
    '----------------------------------------------------------
    sUnitNumber = oRec.Fields.Item("unit_number").Value

    If Not moPatientDemographics.RecordGetDemographics(sUnitNumber, oRec) Then GoTo CleanUp
    moCollRecSet.Add oRec, "demographics"

    '-- Sponsor Demographics    '----------------------------------------------------------
    If oRec.Fields.Item("FMP").Value = "20" Then
        sSponsorUnitNumber = sUnitNumber
    Else
        If gobjshared.IsAppMode(modeCHCSI_GUI) Then
            sSponsorUnitNumber = oRec.Fields.Item("SPONSOR_UNIT_NUMBER").Value
        Else
            sSponsorUnitNumber = moPatientDemographics.GetUnitNumber(oRec.Fields.Item("Sponsor_SSN").Value)
        End If
    End If
    
    If Not moPatientDemographics.RecordGetDemographics(sSponsorUnitNumber, oRec) Then GoTo CleanUp

    
    moCollRecSet.Add oRec, "sponsor"
    
''-- Set Formulas
''-----------------------------------------------------------------------------
'    Set moCollFormulas = Nothing
'    Set moCollFormulas = New Collection
'    moCollFormulas.Add "'" & Format(Now, gsDefaultFormat & "  HHMM") & "'", "TodaysDateTime"

'-- Create Report Object
'-----------------------------------------------------------------------------
    Set oCrystal = New Crystal
    Set oReport = Nothing
    If Not oCrystal.ReportOpen("DD2161.rpt", oReport) Then GoTo CleanUp
    If Not oCrystal.AssignRecordSets(moCollRecSet, oReport) Then GoTo CleanUp
    If Not oCrystal.AssignFormulas(moCollFormulas, oReport) Then GoTo CleanUp

BuildDD2161Complete = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "FormDD2161.BuildDD2161Complete", "CHCSII_Forms", vbExclamation
CleanUp:
    Set oCrystal = Nothing
    Set oRec = Nothing
End Function


Private Function GetConsultRecord(sConsultId As String, _
                                  sConsultReportText As String, _
                                  oRec As ADODB.Recordset) As Boolean
                    
    Dim sStmt As String
    Dim sUrgency As String

    On Error GoTo ErrHandler
    GetConsultRecord = False
    
    '-- Get Consult Information '--------------------------------------------------------------------------
    If gobjshared.UseSQLServer Then
        sStmt = "Select 1 as job_id, CONVERT(VARCHAR(25), GETDATE(),106) as TodaysDateTime, CONVERT(VARCHAR(25), co.order_date,106) as DateOfRequest," & vbNewLine & _
            "co.CIVILIAN_PROVIDER as ToProvider, co.enc_num, co.result_id, co.CIVILIAN_CLINIC as Facility, co.unit_number, " & vbNewLine
        sStmt = sStmt & _
            "CASE co.urgency " & vbNewLine & _
                      " WHEN 0 THEN 'Routine'" & vbNewLine & _
                      " WHEN 1 THEN '24 Hours'" & vbNewLine & _
                      " WHEN 2 THEN '48 Hours'" & vbNewLine & _
                      " WHEN 3 THEN '72 Hours'" & vbNewLine & _
                      " WHEN 4 THEN 'PreOp'" & vbNewLine & _
                      " WHEN 5 THEN 'Today'" & vbNewLine & _
                      " WHEN 6 THEN 'STAT'" & vbNewLine & _
                      " WHEN 7 THEN 'ASAP'" & vbNewLine & _
                  " END as Urgency, " & vbNewLine & _
            "co.primary_dx_text AS ProvisionalDiagnosis, p1.name as FromProvider," & vbNewLine & _
            "co.reason_for_request AS ReasonForRequest, " & vbNewLine & _
            "co.order_status " & vbNewLine & _
            "from provider p1, consult_order co " & vbNewLine
         
        sStmt = sStmt & "WHERE p1.ncid =* co.ref_provider_ncid and " & _
                    "co.consult_id = " & sConsultId
    Else
        sStmt = "Select " & _
                    "1 as job_id, " & _
                    "to_char(sysdate,'DD MON YYYY') as TodaysDateTime, " & vbCrLf & _
                    "co.CIVILIAN_PROVIDER as ToProvider, " & vbCrLf & _
                    "to_char(co.order_date,'DD MON YYYY') as DateOfRequest, " & vbCrLf & _
                    "r.representation as Facility, " & vbCrLf & _
                    "r2.representation as FromProvider, " & vbCrLf & _
                    "co.primary_dx_text as ProvisionalDiagnosis, " & vbCrLf & _
                    "co.reason_for_request as ReasonForRequest, " & vbCrLf & _
                    "decode(co.urgency,0, 'Routine', 1,'24 Hours', 2,'48 Hours', 3,'72 Hours',4, 'PreOp', " & _
                    "5, 'Today', 6, 'STAT', 7, 'ASAP') as Urgency, " & vbCrLf & _
                    "co.unit_number " & vbCrLf & _
                    "from " & vbCrLf & _
                    "rsform_context rc2, rsform r2, " & vbCrLf & _
                    "rsform_context rc, rsform r, " & vbCrLf & _
                    "consult_order co " & vbCrLf
         
            sStmt = sStmt & _
                "where " & vbCrLf & _
                    "rc2.context_ncid = 2000 and " & vbCrLf & _
                    "rc2.rsform_id = r2.rsform_id and " & vbCrLf & _
                    "r2.ncid = co.ref_provider_ncid and " & vbCrLf & _
                    "rc.context_ncid = 2000 and " & vbCrLf & _
                    "rc.rsform_id = r.rsform_id and " & vbCrLf & _
                    "r.ncid = co.enc_fac_ncid and " & vbCrLf & _
                    "co.consult_id = " & sConsultId
    End If
    
    Set oRec = goDBDAS.OpenRecordset(sStmt)
    Set oRec.ActiveConnection = Nothing
    
    GetConsultRecord = True
    GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "FormDD2161.GetConsultRecord", "CHCSII_Forms", vbExclamation
CleanUp:
    Exit Function
End Function


