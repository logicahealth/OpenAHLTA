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
Attribute VB_Name = "RepScreeningServDue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private moCollRecSet As Collection
Private moCollFormulas As Collection
Private msMachineName As String

Public Function BuildReport(sReportName As String, _
                            oReportItem As ReportItem, _
                            oReport As CRAXDRT.Report, _
                            oCollFormulas As Collection) As Boolean
 
    On Error GoTo ErrHandler
     
    BuildReport = False
    
    Dim oCrystal As Crystal
    Dim oRec As ADODB.Recordset
    Dim sSponsorUnitNumber As String
    Dim sUnitNumber As String
    DoEvents
    
    '-- Get Data
    '----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection

    '-- Get Candidates
    '----------------------------------------------------------
    If Not GetData(oReportItem, oRec) Then GoTo CleanUp
    moCollRecSet.Add oRec, "Reminders"

    On Error Resume Next
    oCollFormulas.Add "PCM: " & oReportItem.CriteriaSel1(1).Name, "Header2"
    
'-- Create Report Object
'-----------------------------------------------------------------------------
    If Not CreateReport(sReportName & ".rpt", oReport, oCollFormulas, moCollRecSet) Then GoTo CleanUp

    BuildReport = True

GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepScreeningServDue.BuildReport", "CHCSII_Reports", vbExclamation
CleanUp:
    Set oCrystal = Nothing
    Set oRec = Nothing
    
End Function

Private Function GetData(oReportItem As ReportItem, _
                           oRec As ADODB.Recordset) As Boolean

    Dim sSql As String
    Dim sServiceFilter As String
    Dim sMTFFilter As String
    Dim sPCMFilter As String
    Dim sSortFilter As String
    Dim sTestType As String
    Dim sTestName As String
    Dim sAdditionalReminder As String
    
    Dim oItem As WellnessItem
    
    Dim oIntermediate As New ADODB.Recordset
    Dim oReturnSet As New ADODB.Recordset
    
    On Error GoTo ErrHandler
    GetData = False
           
    sServiceFilter = ""
    For Each oItem In oReportItem.CriteriaSel2
        If Not sServiceFilter = "" Then
            sServiceFilter = sServiceFilter & "," & CStr(oItem.NCID)
        Else
            sServiceFilter = CStr(oItem.NCID)
        End If
        
        sAdditionalReminder = CheckForAltReadinessNCID(CStr(oItem.NCID))
        If Not sAdditionalReminder = "" Then sServiceFilter = sServiceFilter & "," & sAdditionalReminder
    
    Next
    If sServiceFilter = "" Then Exit Function
    
    Set oReturnSet.ActiveConnection = Nothing
    oReturnSet.CursorLocation = adUseClient
    oReturnSet.Fields.Append "PCM_Name", adChar, 30
    oReturnSet.Fields.Append "Full_Name", adChar, 65
    oReturnSet.Fields.Append "Phone", adChar, 35
    oReturnSet.Fields.Append "Street_1", adChar, 30
    oReturnSet.Fields.Append "City_State_post_code", adChar, 60
    oReturnSet.Fields.Append "SSN", adChar, 16
    oReturnSet.Fields.Append "Date_Of_Birth", adChar, 15
    oReturnSet.Fields.Append "Unit_Number", adChar, 10
    oReturnSet.Fields.Append "PATCAT", adChar, 50
    oReturnSet.Fields.Append "email_address", adChar, 50
    
    oReturnSet.Fields.Append "Sponsor_SSN", adChar, 20
    oReturnSet.Fields.Append "Sponsor_Rank", adChar, 50
    oReturnSet.Fields.Append "Sponsor_Full_Name", adChar, 50
    
    oReturnSet.Fields.Append "Accounting_Date", adChar, 12
    oReturnSet.Fields.Append "Accounting_Method", adChar, 15
    oReturnSet.Fields.Append "Test_Results", adChar, 749
    
    oReturnSet.Open
    
    msMachineName = GetLocalComputerName
    
    sSql = "DELETE FROM REMINDER_REPORT_DATA WHERE COMPUTER_NAME = '" & msMachineName & "'"
    goDBSqlOps.Execute sSql

    cwobjCWComm.setStatusText "Loading Matrix..."
    DoEvents
    If Not Populate_Table(sServiceFilter, _
                          oReportItem.CriteriaSel1(1).PCMCode, _
                          oReportItem.CriteriaSel1(1).Name) Then GoTo CleanUp
    
'    cwobjCWComm.setStatusText "Load Patient PCM Matrix..."
'    DoEvents
'    Screening_Load_PCM_Data oReportItem
    
    cwobjCWComm.setStatusText "Load Patient Demographics..."
    DoEvents
    Screening_Load_Demographics oReturnSet, oReportItem
    
    cwobjCWComm.setStatusText "Checking email address..."
    DoEvents
    Screening_Load_Emails oReturnSet
            
'    cwobjCWComm.setStatusText "Finding Reminder(s) For Patients..."
'    DoEvents
'    Screening_Load_Event_Data
'
'    cwobjCWComm.setStatusText "Checking Reminder(s) Status..."
'    DoEvents
'    Load_Screening_Event_Status
'
'    cwobjCWComm.setStatusText "Checking Reminder(s) Dates..."
'    DoEvents
'    Screening_Load_Event_Dates
    
    cwobjCWComm.setStatusText "Checking notifications..."
    DoEvents
    Screening_Load_Notifications oReturnSet
        
    Screening_Load_Dates_Into_Rs oReturnSet, oReportItem
    
    With oReturnSet
        If .BOF And .EOF Then
           Set oRec = oReturnSet
           GetData = True
           GoTo CleanUp
        End If
    End With
            
    sSortFilter = ""
    If Not oReportItem.SortItemSelected = "" Then sSortFilter = oReportItem.SortItemSelected
    If Not oReportItem.SortDirection = "" Then sSortFilter = sSortFilter & " " & oReportItem.SortDirection
    oReturnSet.Sort = sSortFilter
    
    Set oRec = oReturnSet
    GetData = True
    GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "RepScreeningServDue.GetData", "CHCSII_Reports", vbExclamation
    'Resume Next
CleanUp:
    Set oItem = Nothing
    Set oIntermediate = Nothing
    Set oReturnSet = Nothing
End Function


Private Function Populate_Table(sServiceList As String, sPCMCode As String, sPCMName As String) As Boolean

On Error GoTo ErrHandler
Populate_Table = False

Dim oPCMsRec As ADODB.Recordset
Dim oRec As ADODB.Record
Dim sSql As String
Dim i As Integer

    sSql = "INSERT INTO REMINDER_REPORT_DATA" & vbCrLf & _
                "(COMPUTER_NAME, PROCESS_STATUS, UNIT_NUMBER, PCM_CODE," & vbCrLf & _
                "PCM_NAME, ITEM_COUNT, DUE_DATE)" & vbCrLf & _
           "SELECT DISTINCT '" & msMachineName & "',4, e.unit_number," & vbCrLf & _
                  "'" & sPCMCode & "'," & SQLQuote(sPCMName) & "," & vbCrLf & _
                  "e.event_subtype_ncid, ov.VALUE_DATETIME" & vbCrLf & _
           "FROM" & vbCrLf & _
                "EVENT E, CLINICAL_EVENT_HEADER C, CLINICAL_OBSERVATION CO, OBSERVATION_VALUE OV" & vbCrLf & _
           "Where" & vbCrLf & _
                "OV.OBSERVATION_SEQ = CO.OBSERVATION_SEQ AND" & vbCrLf & _
                "OV.EVENT_ID = CO.EVENT_ID AND" & vbCrLf & _
                "CO.EVENT_ID = E.EVENT_ID AND" & vbCrLf & _
                "C.PRODUCTION_STATUS_NCID = 101984 AND" & vbCrLf & _
                "C.STATUS_NCID = 1024 AND" & vbCrLf & _
                "C.EVENT_ID = E.EVENT_ID AND" & vbCrLf & _
                "E.EVENT_SUBTYPE_NCID IN (" & sServiceList & ") AND" & vbCrLf & _
                "(E.EVENT_TYPE_NCID = 162156 OR E.EVENT_TYPE_NCID = 162158) AND" & vbCrLf & _
                "E.UNIT_NUMBER IN" & vbCrLf & _
                    "(SELECT UNIT_NUMBER" & vbCrLf & _
                    "From PATIENT_PCMCODE p" & vbCrLf & _
                    "Where p.PCM_CODE = " & SQLQuote(sPCMCode) & ")"
                        
    goDBSqlOps.Execute sSql

    sSql = "INSERT INTO REMINDER_REPORT_DATA" & vbCrLf & _
                "(COMPUTER_NAME, PROCESS_STATUS, UNIT_NUMBER, PCM_CODE," & vbCrLf & _
                "PCM_NAME, ITEM_COUNT, DUE_DATE)" & vbCrLf & _
           "SELECT DISTINCT '" & msMachineName & "',4, id.unit_number," & vbCrLf & _
                  "'" & sPCMCode & "'," & SQLQuote(sPCMName) & "," & vbCrLf & _
                  "NULL, NULL" & vbCrLf & _
           "FROM" & vbCrLf & _
                "MMI_ID ID" & vbCrLf & _
           "Where" & vbCrLf & _
                "NOT EXISTS" & vbCrLf & _
                    "(select 'x' from reminder_report_data rrd" & vbCrLf & _
                    "where rrd.unit_number = id.unit_number and" & vbCrLf & _
                    "rrd.computer_name = '" & msMachineName & "') AND" & vbCrLf & _
                "ID.UNIT_NUMBER IN" & vbCrLf & _
                    "(SELECT UNIT_NUMBER" & vbCrLf & _
                    "From PATIENT_PCMCODE p" & vbCrLf & _
                    "Where p.PCM_CODE = " & SQLQuote(sPCMCode) & ")"

    goDBSqlOps.Execute sSql
    
    Populate_Table = True
    GoTo CleanUp

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "Populate_Table", "CHCSII_Reports", vbExclamation)
CleanUp:
    
End Function

Private Function Load_Screening_Event_Status() As Boolean
        
    Dim sSql As String
    
    sSql = "UPDATE REMINDER_REPORT_DATA r" & vbCrLf
    sSql = sSql & "SET PROCESS_STATUS = 3" & vbCrLf
    sSql = sSql & "WHERE r.UNIT_NUMBER IN ("
    sSql = sSql & "SELECT DISTINCT /*+ ORDERED index(c CLINICAL_EVENT_HEADER_PK)  */ r.UNIT_NUMBER " & vbCrLf
    sSql = sSql & "FROM REMINDER_REPORT_DATA r, clinical_event_header c " & vbCrLf
    sSql = sSql & "WHERE c.event_id = r.event_id and " & vbCrLf
    sSql = sSql & "c.status_ncid = 1024 and c.PRODUCTION_STATUS_NCID = 101984 and" & vbCrLf
    sSql = sSql & "r.PROCESS_STATUS = 2 AND "
    sSql = sSql & "r.COMPUTER_NAME = " & SQLQuote(msMachineName) & ")"
    
    goDBSqlOps.Execute sSql
        
    sSql = "DELETE FROM REMINDER_REPORT_DATA " & vbCrLf
    sSql = sSql & "WHERE COMPUTER_NAME = " & SQLQuote(msMachineName) & vbCrLf
    sSql = sSql & "AND PROCESS_STATUS = 2"
    goDBSqlOps.Execute sSql
    
End Function

Private Function Screening_Load_Event_Dates() As Boolean
        
    Dim sSql As String
    
    sSql = "INSERT INTO REMINDER_REPORT_DATA" & vbCrLf
    sSql = sSql & "(COMPUTER_NAME, PROCESS_STATUS, UNIT_NUMBER, PCM_CODE, PCM_NAME, ITEM_COUNT, EVENT_ID, DUE_DATE)" & vbCrLf
    sSql = sSql & "SELECT /*+ ORDERED index(ov OBSERVATION_VALUE_PK)  */r.COMPUTER_NAME, 4, r.UNIT_NUMBER, r.PCM_CODE, r.PCM_NAME, r.ITEM_COUNT, r.EVENT_ID, ov.VALUE_DATETIME " & vbCrLf
    sSql = sSql & "From reminder_report_data r, clinical_observation co, observation_value ov" & vbCrLf
    sSql = sSql & "WHERE ov.OBSERVATION_SEQ (+) = co.OBSERVATION_SEQ and" & vbCrLf
    sSql = sSql & "ov.event_id (+) = co.event_id and" & vbCrLf
    sSql = sSql & "co.observation_ncid = 162168 and" & vbCrLf
    sSql = sSql & "co.event_id (+) = r.event_id and" & vbCrLf
    sSql = sSql & "r.PROCESS_STATUS = 3 AND "
    sSql = sSql & "r.COMPUTER_NAME = " & SQLQuote(msMachineName)
    goDBSqlOps.Execute sSql
    
    sSql = "DELETE FROM REMINDER_REPORT_DATA " & vbCrLf
    sSql = sSql & "WHERE COMPUTER_NAME = " & SQLQuote(msMachineName) & vbCrLf
    sSql = sSql & "AND PROCESS_STATUS = 3"
    goDBSqlOps.Execute sSql
    
End Function

Private Function Screening_Load_Event_Data() As Boolean
        
    Dim sSql As String
    
    sSql = "INSERT INTO REMINDER_REPORT_DATA" & vbCrLf
    sSql = sSql & "(COMPUTER_NAME, PROCESS_STATUS, UNIT_NUMBER, PCM_CODE, PCM_NAME, ITEM_COUNT, EVENT_ID)" & vbCrLf
    sSql = sSql & "SELECT /*+ ORDERED index(e EVENT_X1)  */r.COMPUTER_NAME, 2, r.UNIT_NUMBER, r.PCM_CODE, r.PCM_NAME, r.ITEM_COUNT, e.event_id " & vbCrLf
    sSql = sSql & "FROM REMINDER_REPORT_DATA r, EVENT e" & vbCrLf
    sSql = sSql & "WHERE e.unit_number = r.unit_number and " & vbCrLf & _
                "r.PROCESS_STATUS = 1 AND " & _
                "r.COMPUTER_NAME = " & SQLQuote(msMachineName) & " and " & _
                "e.EVENT_SUBTYPE_NCID = r.ITEM_COUNT and e.event_type_ncid = 162156"
    
    goDBSqlOps.Execute sSql
    
    sSql = "DELETE FROM REMINDER_REPORT_DATA " & vbCrLf
    sSql = sSql & "WHERE COMPUTER_NAME = " & SQLQuote(msMachineName) & vbCrLf
    sSql = sSql & "AND PROCESS_STATUS = 1"
    goDBSqlOps.Execute sSql
    
End Function

Private Sub Screening_Load_Demographics(oPatientRecords As ADODB.Recordset, oReportItem As ReportItem)

    Dim sSql As String
    sSql = "select Trim(mn.last_name ||', ' || mn.first_name || ' ' || mn.middle_name) as Full_Name," & vbCrLf & _
            "mn.first_name || ' ' || mn.middle_name || ' ' ||Trim(mn.last_name ) as Print_Name," & vbCrLf & _
            "mp.area_code, mp.local_number, ma.street_1, ma.city, ma.state, ma.post_code," & vbCrLf & _
            "mi.SSN, mi.Birth_Date, Trim(tSSN.id_value) as SPONSOR_SSN, trim(tPC.value_text) as PATCAT, mi.unit_number, " & vbCrLf & _
            "Trim(smn.last_name ||', ' || smn.first_name || ' ' || smn.middle_name) as Sponsor_Full_Name " & vbCrLf & _
            "From reminder_report_data r, mmi_generic_id tSSN,  mmi_genericdata tPC, mmi_phone mp, mmi_address ma, mmi_name mn, mmi_name smn, mmi_id mi" & vbCrLf & _
            "Where mp.phone_type_cid (+)= 193 And mp.unit_number (+)= mi.unit_number AND" & vbCrLf & _
            "ma.address_type_cid (+)= 199 and ma.unit_number (+)= mi.unit_number and" & vbCrLf & _
            "mn.name_type_cid = 187 and mn.unit_number = mi.unit_number and" & vbCrLf & _
            "smn.name_type_cid = 187 and smn.unit_number=tSSN.unit_number and " & vbCrLf & _
            "tSSN.id_type_ncid(+)=186334 and tSSN.unit_number(+)=mi.unit_number and" & vbCrLf & _
            "tPC.type_ncid(+)=14501721 and tPC.unit_number(+)=mi.unit_number and" & vbCrLf & _
            "mi.unit_number = r.unit_number and r.COMPUTER_NAME = " & SQLQuote(msMachineName)
    
    goDBSqlOps.Execute sSql
    If goDBSqlOps.BOF And goDBSqlOps.EOF Then Exit Sub
    goDBSqlOps.MoveFirst
    
    With oPatientRecords
        Do Until goDBSqlOps.EOF
             On Error Resume Next
             .Filter = "Unit_Number = '" & goDBSqlOps("UNIT_NUMBER") & "'"
             If .BOF And .EOF Then
               .AddNew
               .Fields("Unit_Number") = Trim(goDBSqlOps("UNIT_NUMBER"))
               .Fields("Street_1") = Trim(goDBSqlOps("STREET_1"))
                
                If Trim(goDBSqlOps("city")) = "" Or Trim(goDBSqlOps("state")) = "" Then
                   .Fields("City_State_post_code") = goDBSqlOps("city") & goDBSqlOps("state")
                Else
                    .Fields("City_State_post_code") = goDBSqlOps("city") & ", " & goDBSqlOps("state")
                End If
                
                .Fields("City_State_post_code") = Trim(.Fields("City_State_post_code")) & " " & Trim(goDBSqlOps("post_code"))
               
               .Fields("Area_Code") = goDBSqlOps("AREA_CODE")
               .Fields("Home_Phone") = goDBSqlOps("LOCAL_NUMBER")
               
                If Not (Trim(goDBSqlOps("area_code")) = "" And Trim(goDBSqlOps("local_number")) = "") Then
                   .Fields("Phone") = goDBSqlOps("area_code") & "-"
                   If Len(Trim(goDBSqlOps("local_number"))) = 7 Then
                      .Fields("Phone") = Trim(.Fields("Phone")) & Left(Trim(goDBSqlOps("local_number")), 3) & "-" & Right(Trim(goDBSqlOps("local_number")), 4)
                   Else
                      .Fields("Phone") = Trim(.Fields("Phone")) & Trim(goDBSqlOps("local_number"))
                   End If
                Else
                   .Fields("Phone") = Trim(goDBSqlOps("area_code")) & " " & Trim(goDBSqlOps("local_number"))
                End If
                              
               .Fields("Full_Name") = goDBSqlOps("FULL_NAME")
               .Fields("SSN") = Format(goDBSqlOps("SSN"), "000-00-0000")
               .Fields("Date_Of_Birth") = "DOB " & Format(goDBSqlOps("Birth_Date"), "dd mmm yyyy")
               .Fields("PCM_Name") = oReportItem.CriteriaSel1(1).Name
               .Fields("Sponsor_SSN") = Format(goDBSqlOps("SPONSOR_SSN"), "000-00-0000")
               .Fields("Sponsor_Rank") = GetRank(goDBSqlOps("SPONSOR_SSN"))
               .Fields("Sponsor_Full_Name") = goDBSqlOps("Sponsor_Full_Name")
               .Fields("PATCAT") = goDBSqlOps("PATCAT")
             End If
             
             .Filter = ""
             goDBSqlOps.MoveNext
        Loop
    End With
    
End Sub

Private Function Screening_Load_Dates_Into_Rs(oPatientRecords As ADODB.Recordset, oReportItem As ReportItem) As Boolean
    
    Dim sSql As String
    Dim sStartDate As String
    Dim sEndDate As String
    Dim sTestName As String
    Dim sStdNCID As String
    Dim lReminder As Long
    
    sSql = "select UNIT_NUMBER, ITEM_COUNT, DUE_DATE " & vbCrLf
    sSql = sSql & "From reminder_report_data r" & vbCrLf
    sSql = sSql & "WHERE "
    sSql = sSql & "r.COMPUTER_NAME = " & SQLQuote(msMachineName) & " AND" & vbCrLf
    sSql = sSql & "r.PROCESS_STATUS = 4" & vbCrLf
    
    goDBSqlOps.Execute sSql
    If goDBSqlOps.BOF And goDBSqlOps.EOF Then GoTo CleanUpAndExit
    goDBSqlOps.MoveFirst
    With oPatientRecords
        Do Until goDBSqlOps.EOF
             On Error Resume Next
             .Filter = "Unit_Number = '" & goDBSqlOps("UNIT_NUMBER") & "'"
             If Not (.BOF And .EOF) Then
                
                If Not .Fields("Test_Results").Value = "" Then
                   .Fields("Test_Results") = .Fields("Test_Results") & vbCrLf
                End If
                
                sTestName = ""
                For lReminder = 1 To oReportItem.CriteriaSel2.Count
                    sStdNCID = goDBSqlOps("ITEM_COUNT")
                    sStdNCID = CheckForStdReadinessNCID(sStdNCID)
                    If oReportItem.CriteriaSel2(lReminder).NCID = sStdNCID Then
                       sTestName = oReportItem.CriteriaSel2(lReminder).Name
                       Exit For
                    End If
                Next
                
                If sTestName <> "" Then
                    .Fields("Test_Results") = .Fields("Test_Results") & sTestName & ": " & Format(goDBSqlOps("DUE_DATE"), "dd mmm yyyy")
                    If goDBSqlOps("DUE_DATE") < Format(Now, "mm/dd/yyyy") Then .Fields("Test_Results") = .Fields("Test_Results") & " *"
                End If
             End If
             oPatientRecords.Filter = ""
             goDBSqlOps.MoveNext
        Loop
    End With
    
    If Not (oPatientRecords.BOF And oPatientRecords.EOF) Then
        oPatientRecords.MoveFirst
        Do Until oPatientRecords.EOF
            If oPatientRecords.Fields("Test_Results") = "" Then
                oPatientRecords.Delete adAffectCurrent
            End If
            oPatientRecords.MoveNext
        Loop
    End If
    
    GoTo CleanUpAndExit

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ScreeningServices.Screening_Results", "Screening_Results", vbExclamation)
CleanUpAndExit:

End Function

Private Function Screening_Load_Emails(oPatientRecords As ADODB.Recordset) As Boolean
    
    Dim sSql As String
    Const NCID_Email As String = "67549"
    
    sSql = "select distinct mgd.Unit_Number, mgd.Value_Text " & vbCrLf
    sSql = sSql & "From reminder_report_data r, mmi_genericdata mgd" & vbCrLf
    sSql = sSql & "Where mgd.unit_number = r.unit_number and mgd.Type_NCID = " & NCID_Email & " AND" & vbCrLf
    sSql = sSql & "r.COMPUTER_NAME = " & SQLQuote(msMachineName)
    
    goDBSqlOps.Execute sSql
    If goDBSqlOps.BOF And goDBSqlOps.EOF Then GoTo CleanUpAndExit
    goDBSqlOps.MoveFirst
    With oPatientRecords
        Do Until goDBSqlOps.EOF
             On Error Resume Next
             .Filter = "Unit_Number = '" & goDBSqlOps("UNIT_NUMBER") & "'"
             If Not (.BOF And .EOF) Then
                .Fields("emailAddress").Value = goDBSqlOps("Value_Text")
             End If
             oPatientRecords.Filter = ""
             goDBSqlOps.MoveNext
        Loop
    End With
    
    GoTo CleanUpAndExit

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ScreeningServices.Screening_Load_Emails", "Screening_Results", vbExclamation)
CleanUpAndExit:

End Function

Private Function Screening_Load_Notifications(oPatientRecords As ADODB.Recordset) As Boolean
    
    Dim sSql As String
    
    sSql = "select distinct r.Unit_Number, ssn.* " & vbCrLf
    sSql = sSql & "From reminder_report_data r, SCREEN_SERV_NOTIF ssn" & vbCrLf
    sSql = sSql & "Where ssn.unit_number = r.unit_number AND ssn.TYPE = 0 AND" & vbCrLf
    sSql = sSql & "r.COMPUTER_NAME = " & SQLQuote(msMachineName) & " AND" & vbCrLf
    sSql = sSql & "r.PROCESS_STATUS = 4" & vbCrLf
    
    goDBSqlOps.Execute sSql
    If goDBSqlOps.BOF And goDBSqlOps.EOF Then GoTo CleanUpAndExit
    goDBSqlOps.MoveFirst
    With oPatientRecords
        Do Until goDBSqlOps.EOF
             On Error Resume Next
             .Filter = "Unit_Number = '" & goDBSqlOps("UNIT_NUMBER") & "'"
             If Not (.BOF And .EOF) Then
                If Not goDBSqlOps("Last_Notification_Date") = 0 Then
                  .Fields("Accounting_Date") = goDBSqlOps("LAST_NOTIFICATION_DATE")
                  If goDBSqlOps("LAST_NOTIFICATION_METHOD") = 0 Then
                    .Fields("Accounting_Method") = "eMail"
                  ElseIf goDBSqlOps("LAST_NOTIFICATION_METHOD") = 1 Then
                    .Fields("Accounting_Method") = "Letter"
                 Else
                    .Fields("Accounting_Method") = ""
                End If
                End If
             End If
             oPatientRecords.Filter = ""
             goDBSqlOps.MoveNext
        Loop
    End With
    
    GoTo CleanUpAndExit

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "ScreeningServices.Screening_Load_Notifications", "Screening_Results", vbExclamation)
CleanUpAndExit:

End Function

Private Function Screening_Load_PCM_Data(oReportItem As ReportItem) As Boolean
        
    Dim sSql As String
    Dim sAdditionalReminder As String
    Dim lCount As Long
    
    For lCount = 1 To oReportItem.CriteriaSel2.Count
        sSql = "INSERT INTO REMINDER_REPORT_DATA" & vbCrLf
        sSql = sSql & "(COMPUTER_NAME, PROCESS_STATUS, UNIT_NUMBER, PCM_CODE, PCM_NAME, ITEM_COUNT)" & vbCrLf
        sSql = sSql & "SELECT /*+ ORDERED index(p provider_pk) */  " & SQLQuote(msMachineName) & ", 1, mg.UNIT_NUMBER, p.PCM_CODE, p.NAME, " & oReportItem.CriteriaSel2(lCount).NCID & vbCrLf
        sSql = sSql & "FROM PROVIDER p , PATIENT_PCMCODE pp" & vbCrLf
        sSql = sSql & "WHERE pp.PCM_CODE = p.pcm_code and" & vbCrLf
        sSql = sSql & "p.pcm_code  = " & SQLQuote(oReportItem.CriteriaSel1(1).PCMCode)
        goDBSqlOps.Execute sSql
        
        sAdditionalReminder = CheckForAltReadinessNCID(oReportItem.CriteriaSel2(lCount).NCID)
        If Not sAdditionalReminder = "" Then
           sSql = "INSERT INTO REMINDER_REPORT_DATA " & vbCrLf
           sSql = sSql & "(COMPUTER_NAME, PROCESS_STATUS, UNIT_NUMBER, PCM_CODE, PCM_NAME, ITEM_COUNT)" & vbCrLf
           sSql = sSql & "SELECT /*+ ORDERED index(p provider_pk) */  " & SQLQuote(msMachineName) & ", 1, pp.UNIT_NUMBER, pp.PCM_CODE, p.NAME, " & sAdditionalReminder & vbCrLf
           sSql = sSql & "FROM PROVIDER p , PATIENT_PCMCODE pp" & vbCrLf
           sSql = sSql & "WHERE pp.PCM_CODE = p.pcm_code and" & vbCrLf
           sSql = sSql & "p.pcm_code  = " & SQLQuote(oReportItem.CriteriaSel1(1).PCMCode)
           goDBSqlOps.Execute sSql
        End If
    Next
        
End Function

Private Function CheckForAltReadinessNCID(pReminderNCID As String) As String
    
    Select Case pReminderNCID
        Case "14592298"
            CheckForAltReadinessNCID = "14592299"
        Case "14592296"
            CheckForAltReadinessNCID = "14592297"
        Case "14592294"
            CheckForAltReadinessNCID = "14592295"
        Case "14592292"
            CheckForAltReadinessNCID = "14592292"
        Case "14592290"
            CheckForAltReadinessNCID = "14592291"
        Case "14592288"
            CheckForAltReadinessNCID = "14592289"
        Case "14592286"
            CheckForAltReadinessNCID = "14592287"
        Case "14592284"
            CheckForAltReadinessNCID = "14592285"
        Case "14592282"
            CheckForAltReadinessNCID = "14592283"
        Case "14592280"
            CheckForAltReadinessNCID = "14592281"
        Case "14592278"
            CheckForAltReadinessNCID = "14592279"
        Case "14592276"
            CheckForAltReadinessNCID = "14592277"
        Case "14592274"
            CheckForAltReadinessNCID = "14592275"
        Case "14592272"
            CheckForAltReadinessNCID = "14592273"
        Case "14592270"
            CheckForAltReadinessNCID = "14592271"
        Case "14592268"
            CheckForAltReadinessNCID = "14592269"
        Case "14592266"
            CheckForAltReadinessNCID = "14592267"
        Case "14592264"
            CheckForAltReadinessNCID = "14592265"
        Case "14592262"
            CheckForAltReadinessNCID = "14592263"
        Case "14592260"
            CheckForAltReadinessNCID = "14592261"
        Case "14592258"
            CheckForAltReadinessNCID = "14592259"
        Case "14592256"
            CheckForAltReadinessNCID = "14592257"
        Case "14592254"
            CheckForAltReadinessNCID = "14592255"
        Case "14592252"
            CheckForAltReadinessNCID = "14592253"
        Case "14592250"
            CheckForAltReadinessNCID = "14592251"
        Case "14592300"
            CheckForAltReadinessNCID = "14510330"
        Case Else
            CheckForAltReadinessNCID = ""
    End Select

End Function

Private Function CheckForStdReadinessNCID(pReminderNCID As String) As String
    
    Select Case pReminderNCID
        Case "14592299"
            CheckForStdReadinessNCID = "14592298"
        Case "14592297"
            CheckForStdReadinessNCID = "14592296"
        Case "14592295"
            CheckForStdReadinessNCID = "14592294"
        Case "14592293"
            CheckForStdReadinessNCID = "14592292"
        Case "14592291"
            CheckForStdReadinessNCID = "14592290"
        Case "14592289"
            CheckForStdReadinessNCID = "14592288"
        Case "14592287"
            CheckForStdReadinessNCID = "14592286"
        Case "14592285"
            CheckForStdReadinessNCID = "14592284"
        Case "14592283"
            CheckForStdReadinessNCID = "14592282"
        Case "14592281"
            CheckForStdReadinessNCID = "14592280"
        Case "14592279"
            CheckForStdReadinessNCID = "14592278"
        Case "14592277"
            CheckForStdReadinessNCID = "14592276"
        Case "14592275"
            CheckForStdReadinessNCID = "14592274"
        Case "14592273"
            CheckForStdReadinessNCID = "14592272"
        Case "14592271"
            CheckForStdReadinessNCID = "14592270"
        Case "14592269"
            CheckForStdReadinessNCID = "14592268"
        Case "14592267"
            CheckForStdReadinessNCID = "14592266"
        Case "14592265"
            CheckForStdReadinessNCID = "14592264"
        Case "14592263"
            CheckForStdReadinessNCID = "14592262"
        Case "14592261"
            CheckForStdReadinessNCID = "14592260"
        Case "14592259"
            CheckForStdReadinessNCID = "14592258"
        Case "14592257"
            CheckForStdReadinessNCID = "14592256"
        Case "14592255"
            CheckForStdReadinessNCID = "14592254"
        Case "14592253"
            CheckForStdReadinessNCID = "14592252"
        Case "14592251"
            CheckForStdReadinessNCID = "14592250"
        Case "14510330"
            CheckForStdReadinessNCID = "14592300"
        Case Else
            CheckForStdReadinessNCID = pReminderNCID
    End Select

End Function


Private Function GetRank(pSSN As String) As String
    
    Dim sStmt As String
    Dim sUnitNumber As String
    
    sStmt = "select mg.value_ncid as ResultNCID, r.representation as ResultString " & vbCrLf & _
            "from rsform_context rc, rsform r, mmi_genericdata mg, mmi_id mi " & vbCrLf & _
            "where rc.preferred_score = 0 and rc.context_ncid = 2000 and " & vbCrLf & _
            "rc.rsform_id = r.rsform_id and r.ncid = mg.value_ncid and " & vbCrLf & _
            "mg.type_ncid = 1015 and mg.unit_number = mi.unit_number  and " & vbCrLf & _
            "mi.ssn = '" & pSSN & "'"
   goDBSqlOps.Execute sStmt
   If Not (goDBSqlOps.BOF And goDBSqlOps.EOF) Then
      GetRank = goDBSqlOps("ResultString")
   End If
End Function

