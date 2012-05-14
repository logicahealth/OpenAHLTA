VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FormSF513"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const sMODULE_NAME As String = "FormSF513.cls"
Private moCollRecSet As Collection
Private moCollFormulas As Collection

Private moPatientDemographics As PatientDemographics

Public Function BuildSF513(oConsultFormItem As ConsultFormItem, _
                           sPatientUnitNumber As String, _
                           oReport As CRAXDRT.Report) As Boolean

On Error GoTo ErrHandler

BuildSF513 = False

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
    If Not moPatientDemographics.RecordGetDemographics(sPatientUnitNumber, oRec) Then GoTo Cleanup
    moCollRecSet.Add oRec, "demographics"

    '-- Sponsor Demographics
    '----------------------------------------------------------
    
    If gobjshared.IsAppMode(modeCHCSI_GUI) Then
       sSponsorUnitNumber = oRec.Fields.Item("SPONSOR_UNIT_NUMBER").Value
    Else
       sSponsorUnitNumber = moPatientDemographics.GetUnitNumber(oRec.Fields.Item("Sponsor_SSN").Value)
    End If
    
    If Not moPatientDemographics.RecordGetDemographics(sSponsorUnitNumber, oRec) Then GoTo Cleanup
    moCollRecSet.Add oRec, "sponsor"
    
    '-- Create SF513 Data Record
    '----------------------------------------------------------
    If Not CreateSF513Template(oRec) Then GoTo Cleanup
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
    moCollRecSet.Add oRec, "SF513"
    
    
'-- Create Report Object
'-----------------------------------------------------------------------------
    Set oCrystal = New Crystal
    Set oReport = Nothing
    If Not oCrystal.ReportOpen("SF513.rpt", oReport) Then GoTo Cleanup
    If Not oCrystal.AssignRecordSets(moCollRecSet, oReport) Then GoTo Cleanup


BuildSF513 = True

GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".BuildSF513", App.EXEName, vbExclamation
    'Resume Next
Cleanup:
    Set oCrystal = Nothing
    Set oRec = Nothing
    
End Function
Public Function BuildSF513Complete(sConsult_Id As String, _
                                   sConsultationReportText As String, _
                                   oReport As CRAXDRT.Report) As Boolean

On Error GoTo ErrHandler

BuildSF513Complete = False

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
    If Not GetConsultRecord(sConsult_Id, sConsultationReportText, oRec) Then GoTo Cleanup

    '-- Check if consult is complete
    If oRec.Fields.Count <= 1 Then
       If Not GetConsultRecord2(sConsult_Id, sConsultationReportText, oRec) Then GoTo Cleanup
       If oRec.Fields.Count <= 1 Then
          MsgBxARMd "Consult not found in database." & vbCrLf & "Consult Id: " & sConsult_Id, vbExclamation
          GoTo Cleanup
       End If
    End If
    moCollRecSet.Add oRec, "SF513"
    
    '-- Patient Demographics
    '----------------------------------------------------------
    sUnitNumber = oRec.Fields.Item("unit_number").Value

    If Not moPatientDemographics.RecordGetDemographics(sUnitNumber, oRec) Then GoTo Cleanup
    moCollRecSet.Add oRec, "demographics"

    '-- Sponsor Demographics
    '----------------------------------------------------------
    If oRec.Fields.Item("FMP").Value = "20" Then
        sSponsorUnitNumber = sUnitNumber
    Else
        If gobjshared.IsAppMode(modeCHCSI_GUI) Then
            sSponsorUnitNumber = oRec.Fields.Item("SPONSOR_UNIT_NUMBER").Value
        Else
            sSponsorUnitNumber = moPatientDemographics.GetUnitNumber(oRec.Fields.Item("Sponsor_SSN").Value)
        End If
    End If
    
    If Not moPatientDemographics.RecordGetDemographics(sSponsorUnitNumber, oRec) Then GoTo Cleanup

    
    moCollRecSet.Add oRec, "sponsor"
    
'-- Set Formulas
'-----------------------------------------------------------------------------
    Set moCollFormulas = Nothing
    Set moCollFormulas = New Collection
    moCollFormulas.Add "'" & Format(Now, gsDefaultFormat & "  HHMM") & "'", "TodaysDateTime"

'-- Create Report Object
'-----------------------------------------------------------------------------
    Set oCrystal = New Crystal
    Set oReport = Nothing
    If Not oCrystal.ReportOpen("SF513Complete.rpt", oReport) Then GoTo Cleanup
    If Not oCrystal.AssignRecordSets(moCollRecSet, oReport) Then GoTo Cleanup
    If Not oCrystal.AssignFormulas(moCollFormulas, oReport) Then GoTo Cleanup

BuildSF513Complete = True

GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".BuildSF513Complete", App.EXEName, vbExclamation
Cleanup:
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

    '-- Get Consult Information'--------------------------------------------------------------------------
    If gobjshared.UseSQLServer Then
        sStmt = "Select 1 as job_id, " & goDBConn.SQLQuote(sConsultReportText) & " AS ReportText," & vbNewLine & _
            "CONVERT(VARCHAR(25), co.order_date,106) as Order_Date," & vbNewLine & _
            "co.enc_num, co.result_id, c.name as Facility, co.unit_number," & vbNewLine
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
            "co.primary_dx_text, p1.name as recv_entity," & vbNewLine & _
            "co.reason_for_request, p2.name as ref_provider, " & vbNewLine & _
            "co.order_status " & vbNewLine & _
            "from provider p1, provider p2, clinic c, consult_order co " & vbNewLine
         
        sStmt = sStmt & "where c.ncid = co.RECV_CLINIC_NCID AND c.FACILITY_NCID = co.ENC_FAC_NCID AND " & vbNewLine & _
                    "p1.ncid =* co.recv_provider_ncid and " & _
                    "p2.ncid =* co.ref_provider_ncid and " & _
                    "co.consult_id = " & sConsultId
    Else
        sStmt = "Select 1 as job_id, SF513.ReportText, " & _
                    "to_char(co.order_date,'DD MON YYYY') as Order_Date, " & _
                    "co.enc_num, co.result_id, r.representation as Facility, " & _
                    "co.unit_number, " & _
                    "decode(co.urgency,0, 'Routine', 1,'24 Hours', 2,'48 Hours', 3,'72 Hours',4, 'PreOp', " & _
                    "5, 'Today', 6, 'STAT', 7, 'ASAP') as Urgency, " & _
                    "co.primary_dx_text, r2.representation as recv_entity, co.reason_for_request, r3.representation as ref_provider, " & _
                    "co.order_status " & _
                    "from SF513, rsform_context rc3, rsform r3, rsform_context rc2, rsform r2, rsform_context rc, rsform r, consult_order co "
         
            sStmt = sStmt & _
                "where " & _
                    "SF513.Job_Id = 1 and " & _
                    "rc3.context_ncid = 2000 and " & _
                    "rc3.rsform_id = r3.rsform_id and " & _
                    "r3.ncid = co.ref_provider_ncid and " & _
                    "rc2.context_ncid = 2000 and " & _
                    "rc2.rsform_id = r2.rsform_id and " & _
                    "r2.ncid = co.recv_entity_ncid and " & _
                    "rc.context_ncid = 2000 and " & _
                    "rc.rsform_id = r.rsform_id and " & _
                    "r.ncid = co.enc_fac_ncid and " & _
                    "co.consult_id = " & sConsultId
    End If

    Set oRec = goDBDAS.OpenRecordset(sStmt)
    Set oRec.ActiveConnection = Nothing

    If oRec.Fields.Count > 1 Then
        If Not gobjshared.UseSQLServer Then
            oRec.MoveFirst
            oRec.Fields.Item("ReportText").Value = sConsultReportText
            oRec.Update
        End If
    End If
    
    GetConsultRecord = True
    GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".GetConsultRecord", App.EXEName, vbExclamation
Cleanup:
    Exit Function
End Function




Private Function GetConsultRecord2(sConsultId As String, _
                                  sConsultReportText As String, _
                                  oRec As ADODB.Recordset) As Boolean
                    
    Dim oTempResults As ADODB.Recordset
    
    On Error GoTo ErrHandler
    GetConsultRecord2 = False
    
    Dim sStmt As String
    Dim sUrgency As String

    '-- Get Consult Information --------------------------------------------------------------------------
    If gobjshared.UseSQLServer Then
        sStmt = "Select 1 as job_id, " & goDBConn.SQLQuote(sConsultReportText) & " AS ReportText," & vbNewLine & _
            "CONVERT(VARCHAR(25), co.order_date,106) as Order_Date," & vbNewLine & _
            "co.enc_num, co.result_id, c.name as Facility, co.unit_number," & vbNewLine
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
            "co.primary_dx_text, p1.name as recv_entity," & vbNewLine & _
            "co.reason_for_request, p2.name as ref_provider, " & vbNewLine & _
            "co.order_status " & vbNewLine & _
            "from provider p1, provider p2, clinic c, consult_order co " & vbNewLine
         
        sStmt = sStmt & "where c.ncid = co.RECV_CLINIC_NCID AND c.FACILITY_NCID = co.ENC_FAC_NCID AND " & vbNewLine & _
                    "p1.ncid =* co.recv_provider_ncid and " & _
                    "p2.ncid =* co.ref_provider_ncid and " & _
                    "co.consult_id = " & sConsultId
        Set oRec = goDBDAS.OpenRecordset(sStmt, adOpenDynamic)
    Else
       
       sStmt = "Select co.RECV_CLINIC_NCID, co.ref_provider_ncid " & _
               "from consult_order co " & _
               "where co.consult_id = " & sConsultId
        Set oTempResults = goDBDAS.OpenRecordset(sStmt)
       If Not (oTempResults.BOF And oTempResults.EOF) Then
          oTempResults.MoveFirst
          sStmt = "Select 1 as job_id, SF513.ReportText, " & _
                    "to_char(co.order_date,'DD MON YYYY') as Order_Date, " & _
                    "co.enc_num, co.result_id, r.representation as Facility, " & vbCrLf & _
                    "co.unit_number, " & _
                    "decode(co.urgency,0, 'Routine', 1,'24 Hours', 2,'48 Hours', 3,'72 Hours',4, 'PreOp', " & _
                    "5, 'Today', 6, 'STAT', 7, 'ASAP') as Urgency, " & vbCrLf & _
                    "co.primary_dx_text, "
          If Not oTempResults("RECV_CLINIC_NCID") = "" Then
            sStmt = sStmt & "'' as recv_entity, "
          End If
          
          sStmt = sStmt & "co.reason_for_request, "
          If Not oTempResults("RECV_CLINIC_NCID") = "" Then
             sStmt = sStmt & "r3.representation as ref_provider, "
          Else
             sStmt = sStmt & "'' as ref_provider, "
          End If
          sStmt = sStmt & "co.order_status " & vbCrLf
       
          sStmt = sStmt & "from SF513, "
          If Not oTempResults("ref_provider_ncid") = "" Then sStmt = sStmt & "rsform_context rc3, rsform r3, "
          If Not oTempResults("RECV_CLINIC_NCID") = "" Then sStmt = sStmt & "rsform_context rc2, rsform r2, "
          sStmt = sStmt & "rsform_context rc, rsform r, consult_order co " & vbCrLf
    
          sStmt = sStmt & "where SF513.Job_Id = 1 and " & vbCrLf
          If Not oTempResults("ref_provider_ncid") = "" Then
               sStmt = sStmt & "rc3.context_ncid = 2000 and " & vbCrLf & _
                    "rc3.rsform_id = r3.rsform_id and " & vbCrLf & _
                    "r3.ncid = co.ref_provider_ncid and " & vbCrLf
          End If
            
          If Not oTempResults("RECV_CLINIC_NCID") = "" Then
               sStmt = sStmt & "rc2.context_ncid = 2000 and " & vbCrLf & _
                    "rc2.rsform_id = r2.rsform_id and " & vbCrLf & _
                    "r2.ncid = co.RECV_CLINIC_NCID and " & vbCrLf
          End If
            
          sStmt = sStmt & "rc.context_ncid = 2000 and " & vbCrLf & _
                 "rc.rsform_id = r.rsform_id and " & vbCrLf & _
                 "r.ncid = co.enc_fac_ncid and " & vbCrLf & _
                 "co.consult_id = " & sConsultId
             Set oRec = goDBDAS.OpenRecordset(sStmt)
       End If
    End If
    
    Set oRec.ActiveConnection = Nothing
        
    If Not gobjshared.UseSQLServer Then
        If oRec.Fields.Count > 1 Then
            oRec.MoveFirst
            oRec.Fields.Item("ReportText").Value = sConsultReportText
            oRec.Update
        End If
    End If
    
    GetConsultRecord2 = True
    GoTo Cleanup

ErrHandler:
    'Resume Next
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".GetConsultRecord2", App.EXEName, vbExclamation
Cleanup:
    Set oTempResults = Nothing
    
End Function

Private Function CreateSF513Template(oRec As ADODB.Recordset) As Boolean

Dim sStmt As String
Dim i As Integer
Dim i2 As Integer
Dim oRec1 As ADODB.Recordset

CreateSF513Template = False
On Error GoTo ErrHandler

    '-- Select
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
       CreateSF513Template = True
       GoTo Cleanup
    End If
    
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

    CreateSF513Template = True
    GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".CreateSF513Template", App.EXEName, vbExclamation
    'Resume Next
Cleanup:
    Set oRec1 = Nothing
    Exit Function
End Function


Private Sub Class_Initialize()
    Set moPatientDemographics = New PatientDemographics
End Sub


Private Sub Class_Terminate()
    Set moPatientDemographics = Nothing
End Sub


