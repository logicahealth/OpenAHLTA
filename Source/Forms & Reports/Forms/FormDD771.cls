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
Attribute VB_Name = "FormDD771"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const sMODULE_NAME As String = "FormDD771.cls"
Private moCollRecSet As Collection
Private moCollFormulas As Collection

Private moPatientDemographics As PatientDemographics


Public Function GetDD771Data2(sOrderIds As String, oRec As ADODB.Recordset) As Boolean

Dim sStmt As String

GetDD771Data2 = False
On Error GoTo ErrHandler

'-- DUE TO POOR SYSTEM DESIGN BY 3M THIS SQL STATEMENT HAS BEEN
'-- SPLIT INTO TWO AND THEN THE TWO RECORDSETS ARE MERGED
'-----------------------------------------------------------------

    
'-- Create SQL
'-----------------------------------------------------------------

'-- Select
'HWCC SCR47422 2/12/2004 - Merge changes Foster made for ITT 2.1.832 ITT 2Ic
'R. Foster SCR #40701; 40702; 40731; 40732; 40733; 40734; 40735; 40736; 40737; 40746; 40753; 40754; 40759; 40767; 40768; 40770; 40773; 40774;   Developer: Robert Foster 10/27/2003 12:31 PM
'R. Foster Note: Function Changed:CHCSII_FORMS.dll - GetDasOBJ, GetSQLOBJ, (Declarations), GetDD771Data2, GetDD771Data3, GetDD771Data1, Initialize    Reports.dll - BuildOrderedLensesReport, BuildOrderedFramesReport, BuildRejectionReport, BuildOverdueReport, BuildOverdueRangeReport, Initialize, GetDasOBJ, GetSQLOBJ,       Impact of Change: Allow the CHCSII SRTS II Reports and Forms moodules to operate in ITT mode.
'ITT 832 UPDATES BLOCK II
If gobjshared.IsAppMode(modeITT) Then
    sStmt = "SELECT * FROM vwDD771Data2 WHERE CLINICAL_ORDER_ID IN(" & sOrderIds & ")"
Else
    sStmt = _
            "select" & vbNewLine & _
                "sco.right_add," & vbNewLine & _
                "sco.right_seght," & vbNewLine & _
                "sco.left_add," & vbNewLine & _
                "sco.left_seght," & vbNewLine & _
                "sco.priority," & vbNewLine & _
                "sco.user_initials," & vbNewLine & _
                "sp.description as priority_description," & vbNewLine & _
                "sco.fund_cite_id," & vbNewLine & _
                "sco.clinical_order_comment," & vbNewLine & _
                "sco.clinician_name," & vbNewLine & _
                "To_Char(sco.script_expire_date,'DD MON YYYY') as Expire_Date," & _
                "sfe.b_val" & vbNewLine
    
    '-- From
    sStmt = sStmt & _
                "From" & vbNewLine & _
                    "srts_frame_eye_xrefs sfe," & vbNewLine & _
                    "srts_priority sp," & vbNewLine & _
                    "srts_facility lab," & vbNewLine & _
                    "srts_facility clinic," & vbNewLine & _
                    "srts_clinical_order sco" & vbNewLine
    
    '-- Where
    sStmt = sStmt & _
            "Where" & vbNewLine & _
                "sfe.eye (+) = sco.eye and" & vbNewLine & _
                "sfe.frame_type (+) = sco.frame_type and" & vbNewLine & _
                "sp.priority (+) = sco.priority and" & vbNewLine & _
                "lab.facility_id (+) = sco.lab_facility_id and" & vbNewLine & _
                "clinic.facility_id (+) = sco.clinic_facility_id and" & vbNewLine & _
                "sco.clinical_order_id in (" & sOrderIds & ")"
End If
'-- Execute SQL
'-----------------------------------------------------------------
    
    Set oRec = goDBDAS.OpenRecordset(sStmt)

GetDD771Data2 = True
GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".GetDD771Data2", App.EXEName, vbExclamation
CleanUp:
    Exit Function
End Function

Public Function GetDD771Data3(sOrderIds As String, oRec As ADODB.Recordset) As Boolean

Dim sStmt As String

GetDD771Data3 = False
On Error GoTo ErrHandler
    
'-- Create SQL
'-----------------------------------------------------------------

'-- Select
'HWCC SCR47422 2/12/2004 - Merge changes Foster made for ITT 2.1.832 ITT 2Ic
'R. Foster SCR #40701; 40702; 40731; 40732; 40733; 40734; 40735; 40736; 40737; 40746; 40753; 40754; 40759; 40767; 40768; 40770; 40773; 40774;   Developer: Robert Foster 10/27/2003 12:31 PM
'R. Foster Note: Function Changed:CHCSII_FORMS.dll - GetDasOBJ, GetSQLOBJ, (Declarations), GetDD771Data2, GetDD771Data3, GetDD771Data1, Initialize    Reports.dll - BuildOrderedLensesReport, BuildOrderedFramesReport, BuildRejectionReport, BuildOverdueReport, BuildOverdueRangeReport, Initialize, GetDasOBJ, GetSQLOBJ,       Impact of Change: Allow the CHCSII SRTS II Reports and Forms moodules to operate in ITT mode.
'ITT 832 UPDATES BLOCK II
If gobjshared.IsAppMode(modeITT) Then
    sStmt = "SELECT * FROM vwDD771Data3 WHERE CLINICAL_ORDER_ID IN(" & sOrderIds & ")"
Else
    sStmt = _
            "select" & vbNewLine & _
                "sco.clinical_order_id, To_Char(sot.dt_transaction,'DD MON YYYY') as Trans_Date" & vbNewLine
    
    '-- From
    sStmt = sStmt & _
                "From" & vbNewLine & _
                    "srts_order_transaction sot," & vbNewLine & _
                    "srts_clinical_order sco" & vbNewLine
    
    '-- Where
    sStmt = sStmt & _
            "Where" & vbNewLine & _
                "sot.dt_transaction in" & vbNewLine & _
                    "(select max(dt_transaction)" & vbNewLine & _
                    "from srts_order_transaction sot2" & vbNewLine & _
                    "Where" & vbNewLine & _
                    "sot2.order_status = 3 and" & vbNewLine & _
                    "sot2.clinical_order_id = sot.clinical_order_id) and" & vbNewLine & _
                "sot.clinical_order_id = sco.clinical_order_id and" & vbNewLine & _
                "sco.clinical_order_id in (" & sOrderIds & ")"
End If

'-- Execute SQL
'-----------------------------------------------------------------
    
    Set oRec = goDBDAS.OpenRecordset(sStmt)

GetDD771Data3 = True
GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".GetDD771Data3", App.EXEName, vbExclamation
CleanUp:
    Exit Function
End Function


Public Function GetDD771Data1(sOrderIds As String, oRec As ADODB.Recordset) As Boolean

Dim sStmt As String

GetDD771Data1 = False
On Error GoTo ErrHandler

'-- DUE TO POOR SYSTEM DESIGN BY 3M THIS SQL STATEMENT HAS BEEN
'-- SPLIT INTO TWO AND THEN THE TWO RECORDSETS ARE MERGED
'-----------------------------------------------------------------

'-- Create SQL
'-----------------------------------------------------------------

'-- Select
'HWCC SCR47422 2/12/2004 - Merge changes Foster made for ITT 2.1.832 ITT 2Ic
'R. Foster SCR #40701; 40702; 40731; 40732; 40733; 40734; 40735; 40736; 40737; 40746; 40753; 40754; 40759; 40767; 40768; 40770; 40773; 40774;   Developer: Robert Foster 10/27/2003 12:31 PM
'R. Foster Note: Function Changed:CHCSII_FORMS.dll - GetDasOBJ, GetSQLOBJ, (Declarations), GetDD771Data2, GetDD771Data3, GetDD771Data1, Initialize    Reports.dll - BuildOrderedLensesReport, BuildOrderedFramesReport, BuildRejectionReport, BuildOverdueReport, BuildOverdueRangeReport, Initialize, GetDasOBJ, GetSQLOBJ,       Impact of Change: Allow the CHCSII SRTS II Reports and Forms moodules to operate in ITT mode.
'ITT 832 UPDATES BLOCK II
If gobjshared.IsAppMode(modeITT) Then
    sStmt = "SELECT * FROM vwDD771Data1 WHERE CLINICAL_ORDER_ID IN (" & sOrderIds & ")"
Else
    sStmt = _
            "select" & vbNewLine & _
                "sco.clinical_order_id," & vbNewLine & _
                "substr(sco.clinical_order_id,-6) as order_number," & vbNewLine & _
                "sco.clinic_facility_id as account_number," & vbNewLine & _
                "lab.facility_name as lab_name," & vbNewLine & _
                "lab.Address1 as lab_address1," & vbNewLine & _
                "lab.Address2 as lab_address2," & vbNewLine & _
                "lab.City as lab_city," & vbNewLine & _
                "lab.State as lab_state," & vbNewLine & _
                "lab.Zip as lab_zip," & vbNewLine & _
                "lab.Phone_Voice_Commercial as lab_phone_voice," & vbNewLine & _
                "lab.Phone_Voice_DSN as lab_phone_dsn," & vbNewLine & _
                "lab.Phone_FAX_Commercial as lab_phone_fax," & vbNewLine & _
                "clinic.facility_name as clinic_name," & vbNewLine & _
                "clinic.Address1 as clinic_address1," & vbNewLine & _
                "clinic.Address2 as clinic_address2," & vbNewLine & _
                "clinic.City as clinic_city," & vbNewLine & _
                "clinic.State as clinic_state," & vbNewLine & _
                "clinic.Zip as clinic_zip," & vbNewLine & _
                "clinic.Phone_Voice_Commercial as clinic_phone_voice," & vbNewLine & _
                "clinic.Phone_Voice_DSN as clinic_phone_dsn," & vbNewLine & _
                "clinic.Phone_FAX_Commercial as clinic_phone_fax," & vbNewLine & _
                "Trim(sco.last_name || ', ' || sco.first_name || ' ' || sco.middle_name) as patient_name," & vbNewLine
    
    sStmt = sStmt & _
                "sco.ssn," & vbNewLine & _
                "sco.rank," & vbNewLine & _
                "sco.street1," & vbNewLine & _
                "sco.street2," & vbNewLine & _
                "sco.city," & vbNewLine & _
                "sco.state," & vbNewLine & _
                "sco.zipcode," & vbNewLine & _
                "sco.home_phone," & vbNewLine & _
                "sco.ship_to," & vbNewLine & _
                "sco.frame_type," & vbNewLine & _
                "sco.eye," & vbNewLine & _
                "sco.bridge," & vbNewLine & _
                "sco.temple," & vbNewLine & _
                "sco.frame_color," & vbNewLine & _
                "sco.distpd," & vbNewLine & _
                "sco.nearpd," & vbNewLine & _
                "sco.lens_type," & vbNewLine & _
                "sco.tint," & vbNewLine & _
                "sco.lens_material," & vbNewLine & _
                "sco.pair," & vbNewLine & _
                "sco.cases," & vbNewLine & _
                "sco.right_sphere," & vbNewLine & _
                "sco.right_cylinder," & vbNewLine & _
                "sco.right_axis," & vbNewLine
    
                
    sStmt = sStmt & _
                "sco.right_hprism," & vbNewLine & _
                "sco.right_hbase," & vbNewLine & _
                "sco.right_vprism," & vbNewLine & _
                "sco.right_vbase," & vbNewLine & _
                "sco.left_sphere," & vbNewLine & _
                "sco.left_cylinder," & vbNewLine & _
                "sco.left_axis," & vbNewLine & _
                "sco.left_hprism," & vbNewLine & _
                "sco.left_hbase," & vbNewLine & _
                "sco.left_vprism," & vbNewLine & _
                "sco.left_vbase," & vbNewLine
    
    
    sStmt = sStmt & _
                "decode(sco.patient_status,null,'',substr(sco.patient_status,1,1)) as service," & vbNewLine & _
                "decode(sco.patient_status,null,'',substr(sco.patient_status,2,1)) as status1," & vbNewLine & _
                "decode(sco.patient_status,null,'',substr(sco.patient_status,3,1)) as status2" & vbNewLine
    
    
    '-- From
    sStmt = sStmt & _
                "From" & vbNewLine & _
                "srts_frame_eye_xrefs sfe," & vbNewLine & _
                "srts_priority sp," & vbNewLine & _
                "srts_facility lab," & vbNewLine & _
                "srts_facility clinic," & vbNewLine & _
                "srts_clinical_order sco" & vbNewLine
    
    '-- Where
    sStmt = sStmt & _
            "Where" & vbNewLine & _
                "sfe.eye (+) = sco.eye and" & vbNewLine & _
                "sfe.frame_type (+) = sco.frame_type and" & vbNewLine & _
                "sp.priority (+) = sco.priority and" & vbNewLine & _
                "lab.facility_id (+) = sco.lab_facility_id and" & vbNewLine & _
                "clinic.facility_id (+) = sco.clinic_facility_id and" & vbNewLine & _
                "sco.clinical_order_id in (" & sOrderIds & ")"
End If

'-- Execute SQL
'-----------------------------------------------------------------
    
    Set oRec = goDBDAS.OpenRecordset(sStmt)
    
GetDD771Data1 = True
GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".GetDD771Data1", App.EXEName, vbExclamation
CleanUp:
    Exit Function
End Function

Public Function BuildDD771(coOrderIds As Collection, _
                           sReportName As String, _
                           oReport As CRAXDRT.Report) As Boolean

On Error GoTo ErrHandler

BuildDD771 = False

Dim oCrystal As Crystal
Dim oRec As ADODB.Recordset
Dim sOrderIds As String
DoEvents

'-- Create Orders List
'----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection
        
    If Not CreateOrderIdList(coOrderIds, sOrderIds) Then GoTo CleanUp
    
'-- Get DD771 Data
'----------------------------------------------------------
    If Not GetDD771Data(sOrderIds, oRec) Then GoTo CleanUp
    moCollRecSet.Add oRec, "DD771"
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
    If Not oCrystal.ReportOpen(sReportName, oReport) Then GoTo CleanUp
'HWCC SCR47422 2/12/2004 - Merge changes Foster made for ITT 2.1.832 ITT 2Ic
'R. Foster SCR #46299;   Developer: Robert Foster 12/03/2003 05:38 PM
'R. Foster Note: Function Changed: CreateReport, Build Report,  Build DD771      Impact of Change: Modify Reports and code to place a watermark ("Training Use Only") when operating in ITT mode
'   ITT 832 Block 2 Updates
    Dim oFormula As FormulaFieldDefinition
    If Not gobjshared.IsAppMode(modeITT) Then     'Not in Training mode
        For Each oFormula In oReport.FormulaFields
            If oFormula.FormulaFieldName = "TrainingUseOnly" Then
               oFormula.Text = vbNullString
            End If
        Next
    End If
'   ITT 832 Block 2 Updates End
    If Not oCrystal.AssignRecordSets(moCollRecSet, oReport) Then GoTo CleanUp
    If Not oCrystal.AssignFormulas(moCollFormulas, oReport) Then GoTo CleanUp
    oReport.ReadRecords

BuildDD771 = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".BuildDD771", App.EXEName, vbExclamation
CleanUp:
    Set oCrystal = Nothing
    Set oRec = Nothing
    
End Function

Public Function GetDD771Data(sOrderIds As String, oRecIN As ADODB.Recordset) As Boolean

Dim sStmt As String
Dim i As Integer
Dim i2 As Integer
Dim oRec1 As ADODB.Recordset
Dim oRec2 As ADODB.Recordset
Dim oRec3 As ADODB.Recordset
'HWCC SCR47422 2/12/2004 - Merge changes Foster made for ITT 2.1.832 ITT 2Ic
'R. Foster SCR 45720;   Developer: Robert Foster 12/19/2003 02:41 PM
'R. Foster Note: GetDD771Data (Forms Module) @FreeComments (DD771 Reports)     Impact of Change: modify Comments before processing in Crystal Report File
Dim oRec As ADODB.Recordset
Dim oRecTemp As ADODB.Recordset
' SCR 47422 Ends
Dim sComment As String
'R. Foster SCR 45720;   Developer: Robert Foster 12/19/2003 02:41 PM
'R. Foster Note: GetDD771Data (Forms Module) @FreeComments (DD771 Reports)     Impact of Change: modify Comments before processing in Crystal Report File
Dim sDupReason As String
GetDD771Data = False
On Error GoTo ErrHandler

'-- DUE TO POOR SYSTEM DESIGN BY 3M THIS SQL STATEMENT HAS BEEN
'-- SPLIT INTO TWO AND THEN THE TWO RECORDSETS ARE MERGED
'-----------------------------------------------------------------

If Not GetDD771Data1(sOrderIds, oRec1) Then GoTo CleanUp

If Not GetDD771Data2(sOrderIds, oRec2) Then GoTo CleanUp

If Not GetDD771Data3(sOrderIds, oRec3) Then GoTo CleanUp


'-- Build Recordset Template
'-----------------------------------------------------------------

On Error Resume Next
Set oRec = New ADODB.Recordset
Set oRecTemp = New ADODB.Recordset
'R. Foster SCR #40701; 40702; 40731; 40732; 40733; 40734; 40735; 40736; 40737; 40746; 40753; 40754; 40759; 40767; 40768; 40770; 40773; 40774;   Developer: Robert Foster 10/27/2003 12:31 PM
'R. Foster Note: Function Changed:CHCSII_FORMS.dll - GetDasOBJ, GetSQLOBJ, (Declarations), GetDD771Data2, GetDD771Data3, GetDD771Data1, Initialize    Reports.dll - BuildOrderedLensesReport, BuildOrderedFramesReport, BuildRejectionReport, BuildOverdueReport, BuildOverdueRangeReport, Initialize, GetDasOBJ, GetSQLOBJ,       Impact of Change: Allow the CHCSII SRTS II Reports and Forms moodules to operate in ITT mode.
'ITT 832 UPDATES BLOCK II
With oRec
    .ActiveConnection = Nothing
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockBatchOptimistic
End With
'R. Foster SCR #40701; 40702; 40731; 40732; 40733; 40734; 40735; 40736; 40737; 40746; 40753; 40754; 40759; 40767; 40768; 40770; 40773; 40774;   Developer: Robert Foster 10/27/2003 12:31 PM
'R. Foster Note: Function Changed:CHCSII_FORMS.dll - GetDasOBJ, GetSQLOBJ, (Declarations), GetDD771Data2, GetDD771Data3, GetDD771Data1, Initialize    Reports.dll - BuildOrderedLensesReport, BuildOrderedFramesReport, BuildRejectionReport, BuildOverdueReport, BuildOverdueRangeReport, Initialize, GetDasOBJ, GetSQLOBJ,       Impact of Change: Allow the CHCSII SRTS II Reports and Forms moodules to operate in ITT mode.
'ITT 832 UPDATES BLOCK II
For i = 0 To oRec1.Fields.Count - 1
    oRec.Fields.Append oRec1(i).Name, oRec1(i).Type, oRec1(i).DefinedSize, oRec1(i).Attributes
    If oRec1(i).Type = adNumeric Then
        oRec.Fields(i).Precision = oRec1.Fields(i).Precision
    End If
    
Next

For i = 0 To oRec2.Fields.Count - 1
    'R. Foster SCR #40701; 40702; 40731; 40732; 40733; 40734; 40735; 40736; 40737; 40746; 40753; 40754; 40759; 40767; 40768; 40770; 40773; 40774;   Developer: Robert Foster 10/27/2003 12:31 PM
    'R. Foster Note: Function Changed:CHCSII_FORMS.dll - GetDasOBJ, GetSQLOBJ, (Declarations), GetDD771Data2, GetDD771Data3, GetDD771Data1, Initialize    Reports.dll - BuildOrderedLensesReport, BuildOrderedFramesReport, BuildRejectionReport, BuildOverdueReport, BuildOverdueRangeReport, Initialize, GetDasOBJ, GetSQLOBJ,       Impact of Change: Allow the CHCSII SRTS II Reports and Forms moodules to operate in ITT mode.
    'ITT 832 UPDATES BLOCK II
    If oRec2(i).Type = adNumeric Or oRec2(i).Type = adVarChar Then
        oRec.Fields.Append oRec2(i).Name, oRec2(i).Type, oRec2(i).DefinedSize, oRec2(i).Attributes
        oRec.Fields(oRec2.Fields(i).Name).Precision = oRec2.Fields(i).Precision
    Else
        oRec.Fields.Append oRec2(i).Name, oRec2(i).Type, oRec2(i).DefinedSize, oRec2(i).Attributes
    End If
Next

If Err = 0 Then
    oRec.Fields.Append "Trans_Date", 203, -1, 132
End If
oRec.Fields.Append "Trans_Date", 203, -1, 132

On Error GoTo ErrHandler

'-- Merge Data from multiple recordsets into one recordset
'-----------------------------------------------------------------
oRec1.MoveFirst
oRec2.MoveFirst
oRec.ActiveConnection = Nothing
oRec.CursorLocation = adUseClient
oRec.Open

Do While Not oRec1.EOF
    oRec.AddNew
    For i = 0 To oRec1.Fields.Count - 1
        oRec.Fields(oRec1.Fields(i).Name).Value = oRec1.Fields(i).Value
    Next
    
    For i = 0 To oRec2.Fields.Count - 1
        oRec.Fields(oRec2.Fields(i).Name).Value = oRec2.Fields(i).Value
    Next
    
    oRec1.MoveNext
    oRec2.MoveNext
    oRec.Update
Loop


'-- If their is a transaction date then update the recordset
'-- with the appopriate date
'-----------------------------------------------------------------
If oRec3.RecordCount > 0 Then
oRec.MoveFirst
Do While Not oRec.EOF

    Do While Not oRec3.EOF
        
        If oRec.Fields("clinical_order_id").Value = oRec3.Fields("clinical_order_id").Value Then
            oRec.Fields("TRANS_DATE").Value = oRec3.Fields("TRANS_DATE").Value
            oRec.Update
        End If
        oRec3.MoveNext
    Loop
    'R. Foster SCR #46354; 46355;   Developer: Robert Foster 12/05/2003 11:06 AM
    'R. Foster Note: Function Changed: GetDD771Data      Impact of Change: We need to return the Recordset that contains the Transaction Date to the beginning after each search.
    oRec3.MoveFirst
    oRec.MoveNext
Loop
End If

'-- Update/Modify Comments Field for each recordset
'-- Remove ";" characters and carriage return where appropriate
'----------------------------------------------------------------
If oRec.RecordCount > 0 Then
    oRec.MoveFirst
    Do While Not oRec.EOF

        'R. Foster SCR 45720;   Developer: Robert Foster 12/19/2003 02:41 PM
        'R. Foster Note: GetDD771Data (Forms Module) @FreeComments (DD771 Reports)      Impact of Change: modify Comments before processing in Crystal Report File
        sComment = oRec.Fields("clinical_order_comment").Value
        If InStr(1, sComment, "DUPLICATE OVERRIDE REASON:") > 0 Then
            sDupReason = Mid(sComment, InStr(1, sComment, "DUPLICATE OVERRIDE REASON:"))
            sComment = left(sComment, InStr(1, sComment, "DUPLICATE OVERRIDE REASON:") - 1)
        End If
        sComment = Replace(sComment, "SECOND FOC REASON:", vbCrLf & "SECOND FOC REASON:")
        sComment = Replace(sComment, "Patient eligibility verified", vbCrLf & "Patient eligibility verified")
        sComment = sDupReason & vbCrLf & sComment
        'If Not ModifyCommentsField(sComment) Then GoTo CleanUp
        oRec.Fields("clinical_order_comment").Value = sComment
        oRec.Update
        oRec.MoveNext
    Loop
End If

Set oRecIN = oRec
GetDD771Data = True
GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".GetDD771Data", App.EXEName, vbExclamation
CleanUp:
    Set oRec1 = Nothing
    Set oRec2 = Nothing
    Set oRec3 = Nothing
    Set oRec = Nothing
    Set oRecTemp = Nothing
End Function


Public Function CreateOrderIdList(coOrderIds As Collection, _
                                  sOrderIds As String) As Boolean

On Error GoTo ErrHandler
CreateOrderIdList = False

Dim i As Integer

If coOrderIds Is Nothing Then
    gobjshared.ShowVBError 0, "The list of Orders to print is empty,", sMODULE_NAME & ".CreateOrderIdList", App.EXEName, vbExclamation
    Exit Function
End If

If coOrderIds.Count = 0 Then
    sOrderIds = "NULL"
Else
    For i = 1 To coOrderIds.Count
        sOrderIds = sOrderIds & "'" & coOrderIds(i) & "'"
        
        If i <> coOrderIds.Count Then
            sOrderIds = sOrderIds & ","
        End If
        
        If (i Mod 3) = 0 Then
            sOrderIds = sOrderIds & vbNewLine
        End If
    Next
End If

CreateOrderIdList = True
GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".CreateOrderIdList", App.EXEName, vbExclamation
CleanUp:
    Exit Function
End Function



Private Function ModifyCommentsField(ByRef sComments As String) As Boolean

On Error GoTo ErrHandler
ModifyCommentsField = False

Dim iPos As Integer
Dim iLen As Integer

If sComments = Null Then
    ModifyCommentsField = True
    GoTo CleanUp
ElseIf sComments = vbNullString Then
    ModifyCommentsField = True
    GoTo CleanUp
Else
    iPos = InStr(1, sComments, ";")
    iLen = Len(sComments)
    
    If iPos = 0 Then
        ModifyCommentsField = True
        GoTo CleanUp
    ElseIf iLen = 1 Then
        sComments = vbNullString
    ElseIf iPos = 1 Then
        sComments = Mid(sComments, 2)
    ElseIf Mid(sComments, iLen, 1) = ";" And iLen <> iPos Then
        sComments = Mid(sComments, 1, iPos - 1) + Chr(10) + Mid(sComments, iPos + 1, iLen - (iPos + 1))
    Else
        sComments = Mid(sComments, 1, iPos - 1) + Chr(10) + Mid(sComments, iPos + 1)
    End If
End If


ModifyCommentsField = True
GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".ModifyCommentsField", App.EXEName, vbExclamation
CleanUp:
    Exit Function
End Function

Private Sub Class_Initialize()
    Set moPatientDemographics = New PatientDemographics
End Sub


Private Sub Class_Terminate()
    Set moPatientDemographics = Nothing
End Sub






