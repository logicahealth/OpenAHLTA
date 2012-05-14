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
Attribute VB_Name = "TableOps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ReturnObject As Object

Private Const mcUIC As String = "14501722"

'**************************************************************************************
'  Function Name: OpenTable
'
'  Description: Opens a table for input
'
'  Parameters:
'       sTableName - Name of table to open
'       rs - Recordset
'  Return Value:
'       True if successufl
'       False if error
'***************************************************************************************
Public Function OpenTable(sTableName As String, rs As ADODB.Recordset) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler

OpenTable = False

Set rs = Nothing
Set rs = New ADODB.Recordset

rs.CursorLocation = adUseClient
'rs.CursorType = adOpenKeyset
'rs.LockType = adLockOptimistic
rs.Open sTableName, poConnection, adOpenDynamic, adLockOptimistic, adCmdTable
OpenTable = True
GoTo CleanUp

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "TableOps.OpenTable", "OpenTable", vbExclamation)
CleanUp:
    Rem Set rs = Nothing
    Exit Function
End Function
'********************************************************************************
'  Function Name: DeleteFromTable
'
'  Description: Deletes all the data from a table in an access database.
'
'  Parameters:
'    TableName - Name of table to be deleted
'
'  Return Value:
'    True - Successful
'    False - Error occured
'**********************************************************************************
Public Function DeleteFromTable(TableName As String) As Boolean

    If Not DebugMode Then On Error GoTo ErrHandler

    DeleteFromTable = False
    
    'Dim rsOutput As New ADODB.Recordset
    Dim sStmt As String
    
    sStmt = "Delete from " & TableName
    poConnection.Execute sStmt
    'rsOutput.CursorLocation = adUseClient
    'rsOutput.CursorType = adOpenStatic
    'rsOutput.ActiveConnection = poConnection
    'rsOutput.Open sStmt, , , adLockBatchOptimistic, adCmdText
    DeleteFromTable = True
    
GoTo CleanUp

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "TableOps.DeleteFromTable", "TableOps", vbExclamation)
CleanUp:
    'Set rsOutput = Nothing
End Function




Public Function Initialize(oTableRec As Object) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler

Initialize = False



Initialize = True
GoTo CleanUp

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "TableOps.Initialize", "TableOps", vbExclamation)
CleanUp:
    Exit Function
End Function
'**************************************************************************************
'  Function Name: SaveRec
'
'  Description: Insert a record into a Report Table
'
'  Parameters:
'
'  Return Value:
'***************************************************************************************
Public Function SaveRec(oTableRec As Object, rs As ADODB.Recordset) As Boolean

    If Not DebugMode Then On Error GoTo ErrHandler
    
    Dim iVal As Integer
    Dim sModifier As String
    
    SaveRec = False
    
    rs.AddNew
    
    '-- Event Table
    '----------------------------------------------------------------
    
    If TypeOf oTableRec Is MTFItem Then
        Dim oMTF As MTFItem
        Set oMTF = oTableRec
        rs("MTF_NCID") = oMTF.NCID
        rs("MTFName") = Left$(oMTF.Name, rs("MTFName").DefinedSize)
    ElseIf TypeOf oTableRec Is DiagnosisItem Then
        Dim oDiagnosis As DiagnosisItem
        Set oDiagnosis = oTableRec
        rs("DiagnosisNCID") = oDiagnosis.NCID
        rs("ICD9") = oDiagnosis.ICD9
        rs("DiagnosisName") = Left$(oDiagnosis.Name, rs("DiagnosisName").DefinedSize)
    ElseIf TypeOf oTableRec Is ClinicItem Then
        Dim oClinic As ClinicItem
        Set oClinic = oTableRec
        rs("ClinicName") = Left$(oClinic.Name, rs("ClinicName").DefinedSize)
        rs("ClinicNCID") = oClinic.NCID
        rs("MeprsCode") = Left$(oClinic.MeprsCode, rs("MeprsCode").DefinedSize)
        rs("DMISIdCode") = Left$(oClinic.DMISIdCode, rs("DMISIdCode").DefinedSize)
    ElseIf TypeOf oTableRec Is EandMItem Then
        Dim oEandM As EandMItem
        Set oEandM = oTableRec
        rs("EandMDesc") = Left$(oEandM.Name, rs("EandMDesc").DefinedSize)
         rs("EandMNCID") = oEandM.NCID
         rs("EandMCode") = Left$(oEandM.Code, rs("EandMCode").DefinedSize)
    ElseIf TypeOf oTableRec Is DispositionItem Then
        Dim oDisposition As DispositionItem
        Set oDisposition = oTableRec
        rs("DispositionDesc") = Left$(oDisposition.Name, rs("DispositionDesc").DefinedSize)
        rs("DispositionNCID") = oDisposition.NCID
    ElseIf TypeOf oTableRec Is ProviderItem Then
        Dim oProvider As ProviderItem
        Set oProvider = oTableRec
        rs("ProviderName") = Left$(oProvider.Name, rs("ProviderName").DefinedSize)
        rs("ProviderNCID") = oProvider.NCID
    ElseIf TypeOf oTableRec Is PatientItem Then
        Dim oPatient As PatientItem
        Set oPatient = oTableRec
    
        rs("PatientId") = Left$(oPatient.UnitNumber, rs("PatientId").DefinedSize)
        rs("FirstName") = Left$(oPatient.First, rs("firstname").DefinedSize)
        rs("LastName") = Left$(oPatient.Last, rs("LastName").DefinedSize)
        rs("MiddleName") = Left$(oPatient.Middle, rs("MiddleName").DefinedSize)
        rs("FMP") = oPatient.FMP
        rs("SSN") = oPatient.SSN
        rs("Sex") = oPatient.Sex
       If oPatient.DOB <> "" Then
            iVal = DateDiff("yyyy", oPatient.DOB, Now)
            sModifier = ""
            If iVal = 0 Then
               iVal = DateDiff("m", oPatient.DOB, Now)
               sModifier = " months"
               If iVal = 0 Then
                 iVal = DateDiff("d", oPatient.DOB, Now)
                sModifier = " days"
               End If
            End If
            rs("Age") = CStr(iVal) & sModifier
            If Not DebugMode Then On Error GoTo ErrHandler
        End If
        rs("DOB") = Format(oPatient.DOB, gsDefaultFormat)
        rs("HomePhone") = oPatient.HomePhone '"(" & oPatient.HomeAreaCode & ") " &
        rs("WorkPhone") = oPatient.WorkPhone '"(" & oPatient.WorkAreaCode & ") " &
        rs("Street1") = oPatient.Street1
        rs("Street2") = oPatient.Street2
        rs("City") = oPatient.City
        rs("State") = oPatient.State
        rs("ZIP") = oPatient.Zip
        rs("Category") = oPatient.Category
        rs("SponsorSSN") = oPatient.SponsorSSN
        '--- SCR 18327, added data for Rank, Status, Department and Organization
        rs("Rank") = oPatient.RankGrade                                           '--- Rank
        rs("Status") = GetStatus(oPatient.UnitNumber, oPatient.FMP)               '--- Status
        '--- Department
        rs("Department") = GetDepartment(oPatient.UnitNumber, oPatient.FMP, oPatient.Category)
        rs("Organization") = GetOrganization(oPatient.UnitNumber, oPatient.FMP)    '--- Organization
        
    ElseIf TypeOf oTableRec Is PreventiveItem Then
        Dim oPreventive As PreventiveItem
        Set oPreventive = oTableRec
        rs("PatientId") = oPreventive.PatientId
        rs("Desc") = oPreventive.Desc
        rs("Status") = oPreventive.Status
        rs("DateDue") = oPreventive.DateDue
        rs("DateAdministered") = oPreventive.DateAdministered
        rs("Category") = oPreventive.Category
    ElseIf TypeOf oTableRec Is AppointmentItem Then
        Dim oAppointment As CHCSII_Reports.AppointmentItem
        Set oAppointment = oTableRec
        rs("Clinic_NCID") = oAppointment.Clinic_NCID
        rs("Clinician_NCID") = oAppointment.Clinician_NCID
        rs("Appt_Status") = oAppointment.Appt_Status
        rs("Patient_Unit_Number") = oAppointment.Patient_Unit_Number
    ElseIf TypeOf oTableRec Is ConsultItem Then
        Dim oConsultItem As ConsultItem
        Set oConsultItem = oTableRec
        rs("Consult_Id") = oConsultItem.Consult_Id
        rs("Enc_Fac_NCID") = oConsultItem.Enc_Fac_NCID
        rs("Enc_Num") = oConsultItem.Enc_Num
        rs("Name") = oConsultItem.Name
        rs("Ref_Provider_NCID") = oConsultItem.Ref_Provider_NCID
        rs("Unit_Number") = oConsultItem.Unit_Number
    ElseIf TypeOf oTableRec Is ProcedureItem Then
        Dim oProcedureItem As ProcedureItem
        Set oProcedureItem = oTableRec
        rs("desc") = oProcedureItem.Name
        rs("CPT4") = oProcedureItem.CPT4
        rs("ProcedureNCID") = oProcedureItem.NCID
    End If
    
    rs.Update
        
    SaveRec = True
    
    GoTo CleanUp

ErrHandler:
    'Call gobjshared.ShowVBError(Err.Number, Err.Description, "TableOps.SaveRec", "TableOps", vbExclamation)
    '---SCR 29242
    '---If there is one data error, ignore it and continue to save other data.
    Resume Next
CleanUp:
    Exit Function
End Function




Private Function GetStatus(sUnitNumber As String, FMP As String) As String
    Dim sIsActiveDuty As String
    Dim oPatient As PatientDemographics
    
    Set oPatient = New PatientDemographics
    
    sIsActiveDuty = oPatient.IsActiveDuty(sUnitNumber)
    
    If sIsActiveDuty = "N" Or FMP = "20" Then
        GetStatus = "Not Active Duty"
    Else
        GetStatus = "Active Duty"
    End If
    
    Set oPatient = Nothing
    
End Function


Private Function GetDepartment(sUnitNumber As String, FMP As String, sCategory As String) As String
    Dim oPatient As PatientDemographics
    
    Set oPatient = New PatientDemographics
    
    If FMP <> "20" Then
        GetDepartment = "N/A"
    Else
        GetDepartment = oPatient.GetService(sCategory)
    End If
    
    Set oPatient = Nothing
        
End Function




Private Function GetOrganization(sUnitNumber As String, FMP As String) As String
    Dim sSQL As String
    
    GetOrganization = ""
    
    If FMP <> "20" Then
        GetOrganization = "N/A"
    Else
        sSQL = "SELECT VALUE_NCID, VALUE_TEXT" & vbCrLf & _
               "FROM MMI_GENERICDATA" & vbCrLf & _
               "WHERE UNIT_NUMBER = " & sUnitNumber & vbCrLf & _
                    "AND TYPE_NCID = " & mcUIC
        goDBSqlOps.Execute sSQL
                    
        If Not goDBSqlOps.BOF And Not goDBSqlOps.EOF Then
            GetOrganization = goDBSqlOps.Value("VALUE_TEXT")
        End If
    End If
        
End Function

Public Function CloseTable(rs As ADODB.Recordset) As Boolean
    
    If Not DebugMode Then On Error GoTo ErrHandler
    CloseTable = False
    
    If Not rs Is Nothing Then
       If rs.State = adStateOpen Then rs.Close
       Set rs = Nothing
    End If
    
    CloseTable = True
    Exit Function
    
ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "TableOps.OpenTable", "OpenTable", vbExclamation)

End Function

Friend Function GetMTFTableData(sFromTable As String, sFromField As String) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler
Dim rs As ADODB.Recordset
Dim oMTF As MTFItem
Dim oColl As Collection
Dim oMTFSearch As New MTFSearch
Dim i

    GetMTFTableData = False
    If Not DeleteFromTable("MTF") Then GoTo CleanUp
    If Not GetDistinct("Encounter", "MTF_Ncid", oColl) Then GoTo CleanUp
    If Not oColl Is Nothing Then
        If Not OpenTable("MTF", rs) Then GoTo CleanUp
        For i = 1 To oColl.Count
            If oColl.Item(i) <> "" Then
                Set oMTF = oMTFSearch.GetByNCID(oColl.Item(i))
                If Not SaveRec(oMTF, rs) Then GoTo CleanUp
            End If
        Next
        If Not CloseTable(rs) Then GoTo CleanUp
    End If
    GetMTFTableData = True
    
GoTo CleanUp

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "TableOps.GetMTFTableData", "TableOps", vbExclamation)
CleanUp:
    Set oMTFSearch = Nothing
    Set oMTF = Nothing
    Set rs = Nothing
    Exit Function
End Function
Friend Function GetClinicTableData(sFromTable As String, sFromField As String) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler
Dim rs As ADODB.Recordset
Dim oClinic As CHCSII_Search.ClinicItem
Dim oColl As Collection
Dim oClinicSearch As New ClinicSearch
Dim i As Integer

    GetClinicTableData = False
    If Not DeleteFromTable("Clinic") Then GoTo CleanUp
    If Not GetDistinct(sFromTable, sFromField, oColl) Then GoTo CleanUp
    If Not oColl Is Nothing Then
        If Not OpenTable("Clinic", rs) Then GoTo CleanUp
        For i = 1 To oColl.Count
            If oColl.Item(i) <> "" Then
                Set oClinic = oClinicSearch.GetByNCID(oColl.Item(i))
                If Not oClinic Is Nothing Then
                    If Not SaveRec(oClinic, rs) Then GoTo CleanUp
                End If
            End If
        Next
'        rs.Save   'SCR 29271
        If Not CloseTable(rs) Then GoTo CleanUp
    End If
    GetClinicTableData = True

GoTo CleanUp

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "TableOps.GetClinicTableData", "TableOps", vbExclamation)
CleanUp:
    Set oClinicSearch = Nothing
    Set oClinic = Nothing
    Set rs = Nothing
    Exit Function
End Function

Friend Function GetDiagnosisTableData(sFromTable As String, sFromField As String) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler
Dim rs As ADODB.Recordset
Dim oDiagnosis As DiagnosisItem
Dim oColl As Collection
Dim oDiagnosisSearch As New DiagnosisSearch
Dim i As Integer

    GetDiagnosisTableData = False
    
    If Not DeleteFromTable("Diagnosis") Then GoTo CleanUp
    If Not GetDistinct(sFromTable, sFromField, oColl) Then GoTo CleanUp
    If Not oColl Is Nothing Then
        If Not OpenTable("Diagnosis", rs) Then GoTo CleanUp
        For i = 1 To oColl.Count
            If oColl.Item(i) <> "" Then
                Set oDiagnosis = oDiagnosisSearch.GetByNCID(oColl.Item(i))
                If Not SaveRec(oDiagnosis, rs) Then GoTo CleanUp
            End If
        Next
        If Not CloseTable(rs) Then GoTo CleanUp
    End If

    GetDiagnosisTableData = True
    
GoTo CleanUp

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "TableOps.GetDiagnosisTableData", "TableOps", vbExclamation)
CleanUp:
    Set oDiagnosisSearch = Nothing
    Set oDiagnosis = Nothing
    Set rs = Nothing
    Exit Function
End Function

Friend Function GetProcedureTableData(sFromTable As String, sFromField As String) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler
Dim rs As ADODB.Recordset
Dim oProcedure As ProcedureItem
Dim oColl As Collection
Dim oProcedureSearch As New ProcedureSearch
Dim i As Integer

    GetProcedureTableData = False
    
    If Not DeleteFromTable("Procedures") Then GoTo CleanUp
    If Not GetDistinct(sFromTable, sFromField, oColl) Then GoTo CleanUp
    If Not oColl Is Nothing Then
        If Not OpenTable("Procedures", rs) Then GoTo CleanUp
        For i = 1 To oColl.Count
            If oColl.Item(i) <> "" Then
                Set oProcedure = oProcedureSearch.GetByNCID(oColl.Item(i))
                If Not SaveRec(oProcedure, rs) Then GoTo CleanUp
            End If
        Next
        If Not CloseTable(rs) Then GoTo CleanUp
    End If

    GetProcedureTableData = True
    
GoTo CleanUp

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "TableOps.GetProcedureTableData", "TableOps", vbExclamation)
CleanUp:
    Set oProcedureSearch = Nothing
    Set oProcedure = Nothing
    Set rs = Nothing
    Exit Function
End Function


Friend Function GetProviderTableData(sFromTable As String, sFromField As String) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler
Dim rs As ADODB.Recordset
Dim oProvider As ProviderItem
Dim oColl As Collection
Dim oProviderSearch As New ProviderSearch
Dim i As Integer

    GetProviderTableData = False
    
    If Not DeleteFromTable("Provider") Then GoTo CleanUp
    If Not GetDistinct(sFromTable, sFromField, oColl) Then GoTo CleanUp
    If Not oColl Is Nothing Then
        If Not OpenTable("Provider", rs) Then GoTo CleanUp
        For i = 1 To oColl.Count
            If oColl.Item(i) <> "" And oColl.Item(i) <> "0" Then
                Set oProvider = oProviderSearch.GetByNCID(oColl.Item(i))
                If Not SaveRec(oProvider, rs) Then GoTo CleanUp
            End If
        Next
'        rs.Save
        If Not CloseTable(rs) Then GoTo CleanUp
    End If

    DoEvents
    GetProviderTableData = True

GoTo CleanUp

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "TableOps.GetProviderTableData", "TableOps", vbExclamation)
CleanUp:
    Set oProviderSearch = Nothing
    Set oProvider = Nothing
    Set rs = Nothing
    Exit Function
End Function


Friend Function GetPatientTableData(sFromTable As String, sFromField As String) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler
Dim rs As ADODB.Recordset
Dim oPatient As PatientItem
Dim oColl As Collection
Dim oPatientSearch As New PatientSearch
Dim i As Integer

    GetPatientTableData = False
    
    If Not DeleteFromTable("Patient") Then GoTo CleanUp
    If Not GetDistinct(sFromTable, sFromField, oColl) Then GoTo CleanUp
    If Not oColl Is Nothing Then
        If Not OpenTable("Patient", rs) Then GoTo CleanUp
        For i = 1 To oColl.Count
            If oColl.Item(i) <> "" Then
                Set oPatient = oPatientSearch.GetBy(oColl.Item(i), "")
                If Not oPatient Is Nothing Then
                    If Not SaveRec(oPatient, rs) Then GoTo CleanUp
                End If
            End If
            DoEvents
        Next
        If Not CloseTable(rs) Then GoTo CleanUp
    End If

    GetPatientTableData = True

GoTo CleanUp

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "TableOps.GetPatientTableData", "CHCSII_Reports", vbExclamation)
CleanUp:
    Set oPatientSearch = Nothing
    Set oPatient = Nothing
    Set rs = Nothing
    Exit Function
End Function
Friend Function GetSponsorTableData(sFromTable As String, sFromField As String) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler
Dim rs As ADODB.Recordset
Dim oPatient As PatientItem
Dim oColl As Collection
Dim oPatientSearch As New PatientSearch
Dim i As Integer

    GetSponsorTableData = False
    
    If Not DeleteFromTable("Sponsor") Then GoTo CleanUp
    If Not GetDistinct(sFromTable, sFromField, oColl) Then GoTo CleanUp
    If Not oColl Is Nothing Then
        If Not OpenTable("Sponsor", rs) Then GoTo CleanUp
        For i = 1 To oColl.Count
            If oColl.Item(i) <> "" And oColl.Item(i) <> "0" Then
                Set oPatient = oPatientSearch.GetBy("", oColl.Item(i), "")
                If Not oPatient Is Nothing Then
                    If Not SaveRec(oPatient, rs) Then GoTo CleanUp
                End If
            End If
        Next
        'If Not (rs.BOF And rs.EOF) Then rs.Save
        If Not CloseTable(rs) Then GoTo CleanUp
    End If

    DoEvents
    GetSponsorTableData = True

GoTo CleanUp

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "TableOps.GetSponsorTableData", "TableOps", vbExclamation)
CleanUp:
    Set oPatientSearch = Nothing
    Set oPatient = Nothing
    Set rs = Nothing
    Exit Function
End Function




Friend Function GetEandMTableData(sFromTable As String, sFromField As String) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler
Dim rs As ADODB.Recordset
Dim oEandM As EandMItem
Dim oColl As Collection
Dim i As Integer
Dim sStmt As String
Dim sItem As Variant


    GetEandMTableData = False

    If Not DeleteFromTable("EandM") Then GoTo CleanUp

    'sStmt = "select EM_CODE, CODE_DESCRIPTION FROM EM_HELP"

    'moSQL.Execute sStmt
    '    GetEandMTableData = True
    '    GoTo CleanUp
    'End If

    If Not GetDistinct(sFromTable, sFromField, oColl) Then GoTo CleanUp
    
    If Not oColl Is Nothing Then
        If Not OpenTable("EandM", rs) Then GoTo CleanUp
            For Each sItem In oColl
                If sItem <> "" Then
                
                    sStmt = "select RCODE.NCID, " & _
                                "TRIM(RCODE.REPRESENTATION) EMCODE, " & _
                                "TRIM(RDESC.REPRESENTATION) Description " & _
                            "from RSFORM RCODE, RSFORM RDESC, RSFORM_CONTEXT RCCODE, RSFORM_CONTEXT RCDESC " & _
                            "Where " & _
                            "RCDESC.CONTEXT_NCID = 17276 AND " & _
                            "RCDESC.RSFORM_ID = RDESC.RSFORM_ID AND " & _
                            "RCCODE.CONTEXT_NCID = 17274 AND " & _
                            "RCCODE.RSFORM_ID = RCODE.RSFORM_ID AND " & _
                            "RDESC.NCID = " & sItem & " AND " & _
                            "RCODE.NCID = " & sItem
        
                    goDBSqlOps.Execute sStmt
                
                    If Not goDBSqlOps.BOF And Not goDBSqlOps.EOF Then
                        Set oEandM = New EandMItem
                        oEandM.NCID = goDBSqlOps("NCID")
                        oEandM.Code = goDBSqlOps("EMCODE")
                        oEandM.Name = goDBSqlOps("Description")
                        If Not SaveRec(oEandM, rs) Then GoTo CleanUp
                        Set oEandM = Nothing
                    End If
                    
                End If
            Next
        End If
        If Not CloseTable(rs) Then GoTo CleanUp
 

    GetEandMTableData = True

GoTo CleanUp

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "TableOps.GetEandMTableData", "TableOps", vbExclamation)
CleanUp:
    Set oEandM = Nothing
    Set rs = Nothing
    Exit Function
End Function


Friend Function GetAppointmentTableData(sFromTable As String, sFromField As String) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler
Dim rs As ADODB.Recordset
Dim oAppointment As AppointmentItem
Dim oColl As Collection
Dim i As Integer

    GetAppointmentTableData = False
    
    If Not DeleteFromTable("Appointment") Then GoTo CleanUp
    If Not GetDistinct(sFromTable, sFromField, oColl) Then GoTo CleanUp
    If Not oColl Is Nothing Then
        If Not OpenTable("Appointment", rs) Then GoTo CleanUp
        For i = 1 To oColl.Count
            If oColl.Item(i) <> "" Then
            
                Set oAppointment = New AppointmentItem
Rem WBW                oAppointment.NCID = oColl.Item(i)
Rem WBW                oAppointment.Name = HDDLookUp(oColl.Item(i), "2000")
                If Not SaveRec(oAppointment, rs) Then GoTo CleanUp
                Set oAppointment = Nothing

            End If
        Next
        If Not CloseTable(rs) Then GoTo CleanUp
    End If

    GetAppointmentTableData = True

GoTo CleanUp

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "TableOps.GetAppointmentTableData", "TableOps", vbExclamation)
CleanUp:

    Set oAppointment = Nothing
    Set rs = Nothing
    Exit Function
End Function

Friend Function UpdateCheckOutInfoTableData() As Boolean

If Not DebugMode Then On Error GoTo ErrHandler

UpdateCheckOutInfoTableData = False

    Dim rs As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim sStmt As String
    Dim oCheckOutInfoItem As New CheckOutInfoItem

'-- Retrieve Check Out table information
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    rs.LockType = adLockBatchOptimistic
    rs.Open "CheckOutInfo", poConnection, adOpenStatic, adLockBatchOptimistic, adCmdTable
    rs.MoveFirst
    If Not rs.EOF Then
        oCheckOutInfoItem.ClinicInfo = rs("ClinicInfo")
        oCheckOutInfoItem.ConsultInfo = rs("ConsultInfo")
        oCheckOutInfoItem.HeaderInfo = rs("HeaderInfo")
        oCheckOutInfoItem.LabInfo = rs("LabInfo")
        oCheckOutInfoItem.MedicationInfo = rs("MedicationInfo")
        oCheckOutInfoItem.XRAYInfo = rs("XRAYInfo")
    End If

'-- Retrieve Check Out table information

    Load frmCheckOutInfo
    frmCheckOutInfo.CheckOutInfo = oCheckOutInfoItem
    Set frmCheckOutInfo.poMyParent = Me
    frmCheckOutInfo.Show vbModal
    If Not ReturnObject Is Nothing Then
        If Not OpenTable("CheckOutInfo", rs2) Then GoTo CleanUp
        rs2.AddNew
        rs2!ClinicInfo = oCheckOutInfoItem.ClinicInfo
        rs2!ConsultInfo = oCheckOutInfoItem.ConsultInfo
        rs2!HeaderInfo = oCheckOutInfoItem.HeaderInfo
        rs2!LabInfo = oCheckOutInfoItem.LabInfo
        rs2!MedicationInfo = oCheckOutInfoItem.MedicationInfo
        rs2!XRAYInfo = oCheckOutInfoItem.XRAYInfo

        If Not DeleteFromTable("CheckOutInfo") Then GoTo CleanUp
        rs2.Update
        rs2.Close
    End If
'-- Retrieve Check Out table information
    rs.Close

UpdateCheckOutInfoTableData = True
GoTo CleanUp

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "TableOps.GetDistinct", "CustomReportOps", vbExclamation)
CleanUp:
    Set rs = Nothing
    Exit Function
End Function



Friend Function GetDispositionTableData(sFromTable As String, sFromField As String) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler
Dim rs As ADODB.Recordset
Dim oDisposition As DispositionItem
Dim oColl As Collection
Dim i As Integer

    GetDispositionTableData = False
    
    If Not DeleteFromTable("Disposition") Then GoTo CleanUp
    If Not GetDistinct(sFromTable, sFromField, oColl) Then GoTo CleanUp
    If Not oColl Is Nothing Then
        If Not OpenTable("Disposition", rs) Then GoTo CleanUp
        For i = 1 To oColl.Count
            If oColl.Item(i) <> "" Then
            
                Set oDisposition = New DispositionItem
                oDisposition.NCID = oColl.Item(i)
                oDisposition.Name = HDDLookUp(oColl.Item(i), "2000")
                If Not SaveRec(oDisposition, rs) Then GoTo CleanUp
                Set oDisposition = Nothing

            End If
        Next
        If Not CloseTable(rs) Then GoTo CleanUp
    End If

    GetDispositionTableData = True

GoTo CleanUp

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "TableOps.GetAppointmentTableData", "TableOps", vbExclamation)
CleanUp:

    Set oDisposition = Nothing
    Set rs = Nothing
    Exit Function
End Function



'******************************************************************************************************
'  Sub Name: GetDistinct
'
'  Description: Returns distinct values from a column in an Access table.  The table and column
'               values are passed in parameters
'
'  Parameters:
'       sTableName - Name of table query will be run against
'       sFieldName - Name of table column that will be checked for distinct values
'       oColl (Output) - Collection of strings returned with the distinct values from the table
'
'  Return Value:
'       False - if an error occured
'       True  - successful
'******************************************************************************************************
Public Function GetDistinct(sTableName As String, sFieldName As String, oColl As Collection) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler

GetDistinct = False

    Dim rsInput As ADODB.Recordset
    Dim sStmt As String

    Set oColl = Nothing
    Set oColl = New Collection

    Set rsInput = New ADODB.Recordset
    rsInput.CursorLocation = adUseClient
    rsInput.CursorType = adOpenStatic
    rsInput.ActiveConnection = poConnection
    sStmt = "Select Distinct " & sFieldName & " " _
          & "From " & sTableName
    rsInput.Open sStmt, , , adLockBatchOptimistic, adCmdText
    
    If rsInput.EOF = True Then
        Set oColl = Nothing
    Else
        While (Not rsInput.EOF)
            oColl.Add rsInput.Fields(0).Value
            rsInput.MoveNext
        Wend
    End If
    
    rsInput.Close

GetDistinct = True
GoTo CleanUp

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "TableOps.GetDistinct", "CustomReportOps", vbExclamation)
    Set oColl = Nothing
CleanUp:
    Set rsInput = Nothing
    Exit Function
End Function




Public Function HDDLookUp(sNCID As String, sCONTEXT) As String

If Not DebugMode Then On Error GoTo ErrHandler
    
    
    Dim oConcept As New ConceptCtrl
    
    On Error Resume Next
    oConcept.UniqueId = sNCID
    HDDLookUp = oConcept.PrefRep(sCONTEXT).Representation
    If Err <> 0 Then
        HDDLookUp = "Unknown"
    End If

GoTo CleanUp

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "TableOps.HDDLookUp", "CHCSII_DiagnosisSearch", vbExclamation)
CleanUp:
    Set oConcept = Nothing
    Exit Function
End Function

Public Function RetrieveReportInfo(eReportId As ReportType, oReportInfo As ReportInformation) As Boolean

If Not DebugMode Then On Error GoTo ErrHandler

RetrieveReportInfo = False

    Dim rsInput As ADODB.Recordset
    Dim sStmt As String

    Set rsInput = New ADODB.Recordset
    rsInput.CursorLocation = adUseClient
    rsInput.CursorType = adOpenStatic
    rsInput.ActiveConnection = poConnection
    sStmt = "Select ReportId, Name, LongName, Header, DateFrom, DateTo, ADate from reports where reportId = " & CStr(eReportId)
    rsInput.Open sStmt, , , adLockBatchOptimistic, adCmdText
    
    If rsInput.EOF = True Then
        RetrieveReportInfo = False
        GoTo CleanUp
    Else
        Set oReportInfo = Nothing
        Set oReportInfo = New ReportInformation
        oReportInfo.ReportTypeId = rsInput.Fields(0).Value
        oReportInfo.ReportFileName = rsInput.Fields(1).Value
        oReportInfo.LongName = rsInput.Fields(2).Value
        oReportInfo.Header = rsInput.Fields(3).Value
        oReportInfo.DateFrom = rsInput.Fields(4).Value
        oReportInfo.DateTo = rsInput.Fields(5).Value
        oReportInfo.ADate = rsInput.Fields(6).Value
    End If
    
    rsInput.Close

GoTo CleanUp

ErrHandler:
    Call gobjshared.ShowVBError(Err.Number, Err.Description, "TableOps.RetrieveReportOptions", "CHCSII_Reports", vbExclamation)
CleanUp:
    Set rsInput = Nothing
    Exit Function

End Function

Private Sub Class_Terminate()
    Set ReturnObject = Nothing
End Sub

