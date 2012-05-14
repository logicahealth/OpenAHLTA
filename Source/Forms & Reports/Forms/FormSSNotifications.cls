VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FormScreeningNotifications"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private moCollRecSet As Collection
Private moCollFormulas As Collection

Private Const gsMainDelim As String = "|"
Private Const msEscChr As String = "\"

Public Function Build(oReport As CRAXDRT.Report, cRecordsets As Collection) As Boolean

    On Error GoTo ErrHandler
    Build = False

    Dim oCrystal As Crystal
    Dim oRec As ADODB.Recordset
    Dim sSponsorUnitNumber As String
    Dim sUnitNumber As String
    Dim oGoals As Object
    Dim sSelfMgmtText As String
    DoEvents


    '-- Get Data          ----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection
    
    moCollRecSet.Add cRecordsets(1), "demographics"
    moCollRecSet.Add cRecordsets(2), "Reminders"
    
    '-- Set Formulas   -----------------------------------------------------------------------------
    Set moCollFormulas = Nothing
    Set moCollFormulas = New Collection
    moCollFormulas.Add "'" & Format(Now, gsDefaultFormat) & "'", "TodaysDateTime"
    
    '-- Create Report Object   -----------------------------------------------------------------------------
    Set oCrystal = New Crystal
    Set oReport = Nothing
    If Not oCrystal.ReportOpen("ScreeningServicesReminder.rpt", oReport) Then GoTo CleanUp
    If Not oCrystal.AssignRecordSets(moCollRecSet, oReport) Then GoTo CleanUp
    If Not oCrystal.AssignFormulas(moCollFormulas, oReport) Then GoTo CleanUp
    
    Build = True
    GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "FormScreeningNotifications.Build", "CHCSII_Forms", vbExclamation
CleanUp:
    Set oCrystal = Nothing
    Set oRec = Nothing
    
End Function
Public Function BuildAddressLabels(oReport As CRAXDRT.Report, _
                      rsDemographics As ADODB.Recordset) As Boolean

    On Error GoTo ErrHandler
    BuildAddressLabels = False

    Dim oCrystal As Crystal
    Dim oRec As ADODB.Recordset
    Dim sSponsorUnitNumber As String
    Dim sUnitNumber As String
    Dim oGoals As Object
    Dim sSelfMgmtText As String
    DoEvents


    '-- Get Data          ----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection
    
    Set oRec = GetLabelRSFromDemographics(rsDemographics)
    moCollRecSet.Add oRec, "Address"
    
    '-- Set Formulas   -----------------------------------------------------------------------------
    Set moCollFormulas = Nothing
    Set moCollFormulas = New Collection
    moCollFormulas.Add "'" & Format(Now, gsDefaultFormat) & "'", "TodaysDateTime"
    
    '-- Create Report Object   -----------------------------------------------------------------------------
    Set oCrystal = New Crystal
    Set oReport = Nothing
    If Not oCrystal.ReportOpen("MailingAddressLabel.rpt", oReport) Then GoTo CleanUp
    If Not oCrystal.AssignRecordSets(moCollRecSet, oReport) Then GoTo CleanUp
    If Not oCrystal.AssignFormulas(moCollFormulas, oReport) Then GoTo CleanUp
    
    BuildAddressLabels = True
    GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "FormScreeningNotifications.BuildAddressLabels", "CHCSII_Forms", vbExclamation
CleanUp:
    Set oCrystal = Nothing
    Set oRec = Nothing
End Function

Public Function BuildEnvelopes(oReport As CRAXDRT.Report, _
                      rsDemographics As ADODB.Recordset) As Boolean

    On Error GoTo ErrHandler
    BuildEnvelopes = False

    Dim oCrystal As Crystal
    Dim oRec As ADODB.Recordset
    Dim sSponsorUnitNumber As String
    Dim sUnitNumber As String
    Dim oGoals As Object
    Dim sSelfMgmtText As String
    DoEvents


    '-- Get Data          ----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection
    
    Set oRec = GetLabelRSFromDemographics(rsDemographics)
    moCollRecSet.Add oRec, "Address"
    
    '-- Set Formulas   -----------------------------------------------------------------------------
    Set moCollFormulas = Nothing
    Set moCollFormulas = New Collection
    moCollFormulas.Add "'" & Format(Now, gsDefaultFormat) & "'", "TodaysDateTime"
    
    '-- Create Report Object   -----------------------------------------------------------------------------
    Set oCrystal = New Crystal
    Set oReport = Nothing
    If Not oCrystal.ReportOpen("MailingAddress.rpt", oReport) Then GoTo CleanUp
    If Not oCrystal.AssignRecordSets(moCollRecSet, oReport) Then GoTo CleanUp
    If Not oCrystal.AssignFormulas(moCollFormulas, oReport) Then GoTo CleanUp
    
    oReport.PaperSource = crPRBinEnvelope
    DoEvents
    
    BuildEnvelopes = True
    GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "FormScreeningNotifications.BuildEnvelopes", "CHCSII_Forms", vbExclamation
CleanUp:
    Set oCrystal = Nothing
    Set oRec = Nothing
End Function

Public Function BuildMTF_PCM_Due_List(oReport As CRAXDRT.Report, _
                      rsPCMCounts As ADODB.Recordset) As Boolean

    On Error GoTo ErrHandler
    BuildMTF_PCM_Due_List = False

    Dim oCrystal As Crystal
    Dim oRec As ADODB.Recordset

    '-- Get Data          ----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection
    
    Set oRec = GetPCMListFromRS(rsPCMCounts)
    moCollRecSet.Add oRec, "PCMList"
    
    '-- Set Formulas   -----------------------------------------------------------------------------
    Set moCollFormulas = Nothing
    Set moCollFormulas = New Collection
    moCollFormulas.Add "'" & Format(Now, gsDefaultFormat) & "'", "TodaysDateTime"
    
    '-- Create Report Object   -----------------------------------------------------------------------------
    Set oCrystal = New Crystal
    Set oReport = Nothing
    If Not oCrystal.ReportOpen("PCM_ScreenServ.rpt", oReport) Then GoTo CleanUp
    If Not oCrystal.AssignRecordSets(moCollRecSet, oReport) Then GoTo CleanUp
    If Not oCrystal.AssignFormulas(moCollFormulas, oReport) Then GoTo CleanUp
    
    DoEvents
    
    BuildMTF_PCM_Due_List = True
    GoTo CleanUp

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, "FormScreeningNotifications.BuildMTF_PCM_Due_List", "CHCSII_Forms", vbExclamation
    Resume Next
CleanUp:
    Set oCrystal = Nothing
    Set oRec = Nothing
End Function

Private Function GetLabelRSFromDemographics(pSource As ADODB.Recordset) As ADODB.Recordset
    
    Dim oReturnSet As New ADODB.Recordset
    Set GetLabelRSFromDemographics = New ADODB.Recordset
    
    Set oReturnSet.ActiveConnection = Nothing
    oReturnSet.CursorLocation = adUseClient
    oReturnSet.Fields.Append "Full_Name", adChar, 65
    oReturnSet.Fields.Append "street_1", adChar, 30
    oReturnSet.Fields.Append "city", adChar, 34
    oReturnSet.Fields.Append "state", adChar, 10
    oReturnSet.Fields.Append "post_code", adChar, 12
    oReturnSet.Fields.Append "UnitNumber", adChar, 10
    oReturnSet.Open
    
    With pSource
        If .BOF And .EOF Then Exit Function
        .MoveFirst
        Do Until .EOF
              oReturnSet.AddNew
              oReturnSet.Fields("Full_Name") = .Fields("Full_Name")
              oReturnSet.Fields("street_1") = .Fields("street_1")
              oReturnSet.Fields("city") = .Fields("city")
              oReturnSet.Fields("state") = .Fields("state")
              oReturnSet.Fields("post_code") = .Fields("post_code")
              oReturnSet.Fields("UnitNumber") = .Fields("UnitNumber")
              .MoveNext
        Loop
    End With
    
    Set GetLabelRSFromDemographics = oReturnSet
    
End Function

Private Function GetPCMListFromRS(pSource As ADODB.Recordset) As ADODB.Recordset
    
    Dim oReturnSet As New ADODB.Recordset
    Set GetPCMListFromRS = New ADODB.Recordset
    
    Set oReturnSet.ActiveConnection = Nothing
    oReturnSet.CursorLocation = adUseClient
    oReturnSet.Fields.Append "PCM_CODE", adChar, 30
    oReturnSet.Fields.Append "NAME", adChar, 30
    oReturnSet.Fields.Append "SERVICENAME", adChar, 30
    oReturnSet.Fields.Append "COMINGDUE_COUNT", adChar, 20
    oReturnSet.Fields.Append "OVERDUE_COUNT", adChar, 20
    oReturnSet.Open
    
    With pSource
        If .BOF And .EOF Then Exit Function
        .MoveFirst
        Do Until .EOF
              oReturnSet.AddNew
              oReturnSet.Fields("PCM_CODE") = .Fields("PCM_CODE")
              oReturnSet.Fields("NAME") = .Fields("PCM_NAME")
              oReturnSet.Fields("SERVICENAME") = .Fields("event_subtype_ncid")
              oReturnSet.Fields("COMINGDUE_COUNT") = .Fields(Trim(.Fields("event_subtype_ncid")) & "_Comingdue_Count")
              oReturnSet.Fields("OVERDUE_COUNT") = .Fields(Trim(.Fields("event_subtype_ncid")) & "_Overdue_Count")
              .MoveNext
        Loop
    End With
    
    Set GetPCMListFromRS = oReturnSet
    
End Function


