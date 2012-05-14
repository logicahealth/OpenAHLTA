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
Attribute VB_Name = "FormCheckOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const sMODULE_NAME As String = "FormCheckOut.cls"
Private Const sNoRcd As String = "_CHCSII_No_Record_Found"
Private moCollRecSet As Collection
Private moCollFormulas As Collection

Private moPatientDemographics As PatientDemographics

Public Function BuildPatientCheckOut(sEncounterNumber As String, _
                                     sPatientUnitNumber As String, _
                                     oReport As CRAXDRT.Report) As Boolean

On Error GoTo ErrHandler
    BuildPatientCheckOut = False

    Dim oCrystal As Crystal
    Dim oRec As ADODB.Recordset
    Dim sClinicNCID As String
    Dim sActiveDuty As String
    Dim bDefault As Boolean
    DoEvents

'-- Get Data
'----------------------------------------------------------------------------
    Set moCollRecSet = Nothing
    Set moCollRecSet = New Collection

    '-- Get Rads
    '----------------------------------------------------------
    If Not GetRads(sEncounterNumber, oRec) Then GoTo Cleanup
    moCollRecSet.Add oRec, "rads"
    
    '-- Get Meds
    '----------------------------------------------------------
    If Not GetMeds(sEncounterNumber, oRec) Then GoTo Cleanup
    moCollRecSet.Add oRec, "meds"
    
    '-- Get Labs
    '----------------------------------------------------------
    If Not GetLabs(sEncounterNumber, oRec) Then GoTo Cleanup
    moCollRecSet.Add oRec, "labs"
    
    '-- Get Consults
    '----------------------------------------------------------
    If Not GetConsults(sEncounterNumber, oRec) Then GoTo Cleanup
    moCollRecSet.Add oRec, "consults"
    
    '-- Get CheckOutInfo
    '----------------------------------------------------------
    If Not GetClinicNCID(sEncounterNumber, sClinicNCID) Then GoTo Cleanup
    If Not GetCheckOutInfo(sClinicNCID, bDefault, oRec) Then GoTo Cleanup
    moCollRecSet.Add oRec, "checkoutinfo"

    '-- Patient Demographics
    '----------------------------------------------------------
    If Not moPatientDemographics.RecordGetDemographics(sPatientUnitNumber, oRec) Then GoTo Cleanup
    moCollRecSet.Add oRec, "demographics"

    sActiveDuty = ""
    If Not (oRec.BOF And oRec.EOF) Then
       sActiveDuty = oRec("Active_Duty")
    End If
    '-- Get Preventive
    '----------------------------------------------------------
    If Not GetCPSServicesDue(sPatientUnitNumber, oRec, sActiveDuty) Then GoTo Cleanup
    moCollRecSet.Add oRec, "servicesdue"

'-- Set Formulas
'-----------------------------------------------------------------------------
    Set moCollFormulas = Nothing
    Set moCollFormulas = New Collection
    moCollFormulas.Add "'" & Format(Now, gsDefaultFormat & "  HHMM") & "'", "TodaysDateTime"

'-- Create Report Object
'-----------------------------------------------------------------------------
    Set oCrystal = New Crystal
    Set oReport = Nothing
    If Not oCrystal.ReportOpen("CheckOutForm.rpt", oReport) Then GoTo Cleanup
    If Not oCrystal.AssignRecordSets(moCollRecSet, oReport) Then GoTo Cleanup
    If Not oCrystal.AssignFormulas(moCollFormulas, oReport) Then GoTo Cleanup

    BuildPatientCheckOut = True

GoTo Cleanup

ErrHandler:
    Call HandleError(Err, sMODULE_NAME & "." & "BuildPatientCheckOut")
Cleanup:
End Function



Private Function GetClinicName(sClinicNCID As String, sClinicName As String) As Boolean
On Error GoTo ErrHandler
    GetClinicName = False
    
    Dim sStmt As String
    
    sStmt = "select c.Name " & _
            "from Clinic c " & _
            "where c.ncid = " & sClinicNCID
    
    goDBSqlOps.Execute sStmt
    
    If Not goDBSqlOps.EOF Then
        sClinicName = goDBSqlOps.Value("Name")
    Else
        sClinicName = vbNullString
    End If
    
    GetClinicName = True
    
    GoTo Cleanup

ErrHandler:
    Call HandleError(Err, sMODULE_NAME & "." & "GetClinicName")
Cleanup:
End Function

Private Function GetClinicNCID(sEncounterNumber As String, sClinicNCID As String) As Boolean
    
On Error GoTo ErrHandler
    Dim sStmt As String
    
    GetClinicNCID = False
    
    sStmt = "select a.clinic_ncid " & _
            "from appointment a " & _
            "where a.encounter_number = " & goDBConn.SQLQuote(sEncounterNumber)
    
    goDBSqlOps.Execute sStmt
    
    If Not goDBSqlOps.EOF Then
        sClinicNCID = goDBSqlOps.Value("clinic_ncid")
    Else
        sClinicNCID = "0"
    End If
    
    GetClinicNCID = True
    GoTo Cleanup

ErrHandler:
    Call HandleError(Err, sMODULE_NAME & "." & "GetClinicNCID")
Cleanup:
End Function

Private Function GetRads(sEncounterNumber As String, _
                         oRec As ADODB.Recordset) As Boolean

On Error GoTo ErrHandler

    GetRads = False
    Dim sStmt As String
    
    '-- Get Radiology Orders
    '----------------------------------------------------------------------------------------
    
    '--- SCR 75487   Sherry Wang  12/9/2005
    If gobjshared.IsAppMode(modeITT) Or gobjshared.IsAppMode(modeLocalCache) _
        Or gobjshared.IsAppMode(modeTheater) Or gobjshared.IsAppMode(modeCDR) Then
       sStmt = "select order_name " _
           & " From oe_rad " _
           & " Where enc_num = '" & sEncounterNumber & "'"
       
       Set oRec = goDBDAS.OpenRecordset(sStmt)
    Else
       CreateGemsStubRecordset oRec, "order_name", "oe_rad"
    End If
    GetRads = True
    GoTo Cleanup

ErrHandler:
    Call HandleError(Err, sMODULE_NAME & "." & "GetRads")
Cleanup:
End Function

Private Sub CreateGemsStubRecordset(ByRef objRecordset As ADODB.Recordset, _
                                    ByVal strField1Name As String, _
                                    ByVal strField2Name As String, _
                                    Optional ByVal strField3Name As String = vbNullString)
On Error GoTo ErrHandler
 
   Set objRecordset = New ADODB.Recordset
    
   objRecordset.Fields.Append strField1Name, adLongVarWChar, -1, 132
   objRecordset.Fields.Append strField2Name, adLongVarWChar, -1, 132
   
   If Trim(strField3Name) <> vbNullString Then
      objRecordset.Fields.Append strField3Name, adLongVarWChar, -1, 132
   End If
   objRecordset.Open
   objRecordset.AddNew
   
   objRecordset.Fields(strField1Name).Value = sNoRcd
   objRecordset.Fields(strField2Name).Value = sNoRcd
   
   If Trim(strField3Name) <> vbNullString Then
      objRecordset.Fields(strField3Name).Value = sNoRcd
   End If
    
   GoTo Cleanup
   
ErrHandler:
    Call HandleError(Err, sMODULE_NAME & "." & "CreateGemsStubRecordset")
Cleanup:
End Sub

Private Function GetMeds(ByVal strEncounterNumber As String, _
                         ByRef objRecordset As ADODB.Recordset) As Boolean

   On Error GoTo ErrHandler
   
   Dim strSQL As String
   
   GetMeds = False
   
   '-- Get Medication Orders
       
   '--- SCR 75487   Sherry Wang  12/9/2005
   If gobjshared.IsAppMode(modeITT) Or gobjshared.IsAppMode(modeLocalCache) _
        Or gobjshared.IsAppMode(modeTheater) Or gobjshared.IsAppMode(modeCDR) Then
          strSQL = "select quantity, refills, order_name " _
              & " From oe_rx " _
              & " Where " _
              & " enc_num = '" & strEncounterNumber & "'"
   Else 'PGUI mode
         strSQL = "SELECT QUANTITY, REFILLS, ORDERNUMBER AS order_name FROM MEDICATIONS " & _
                  "WHERE UNIT_NUMBER = '" & strEncounterNumber & "'"
   End If
    
   Set objRecordset = goDBDAS.OpenRecordset(strSQL)

   GetMeds = True
   
   GoTo Cleanup
ErrHandler:
    Call HandleError(Err, sMODULE_NAME & "." & "GetMeds")
Cleanup:
End Function

Private Function GetConsults(sEncounterNumber As String, _
                             oRec As ADODB.Recordset) As Boolean

    Dim sStmt As String
    
    On Error GoTo ErrHandler
    GetConsults = False
    
    '-- Get Consult Orders'----------------------------------------------------------------------------------------
    CreateGemsStubRecordset oRec, "Name", "diagnosis_text"
    
    GetConsults = True
    GoTo Cleanup
    
ErrHandler:
    Call HandleError(Err, sMODULE_NAME & "." & "GetConsults")
Cleanup:
End Function
Public Function GetCheckOutInfo(sClinicNCID As String, _
                                bDefault As Boolean, _
                                oRec As ADODB.Recordset) As Boolean

    Dim sStmt As String
    Dim sClinicName As String
    
    On Error GoTo ErrHandler
    
    GetCheckOutInfo = False
    
    '-- Get CheckOut Info    '----------------------------------------------------------------------------------------
    sStmt = "select 1 as Job_Id, c.MedicationInfo, c.XRayInfo, c.LabInfo, c.ConsultInfo, c.HeaderInfo, c.ClinicInfo " & _
            "from CheckOutInfo c " & _
            "where c.Clinic_NCID = " & sClinicNCID
    Set oRec = goDBDAS.OpenRecordset(sStmt)
    If oRec.RecordCount = 0 Then
        Set oRec = New ADODB.Recordset
        oRec.Fields.Append "job_id", adLongVarWChar, -1, 132
        oRec.Fields.Append "MedicationInfo", adLongVarWChar, -1, 132
        oRec.Fields.Append "XRayInfo", adLongVarWChar, -1, 132
        oRec.Fields.Append "LabInfo", adLongVarWChar, -1, 132
        oRec.Fields.Append "ConsultInfo", adLongVarWChar, -1, 132
        oRec.Fields.Append "HeaderInfo", adLongVarWChar, -1, 132
        oRec.Fields.Append "ClinicInfo", adLongVarWChar, -1, 132
        oRec.Open
        oRec.AddNew
        oRec("job_id") = "1"
        GetCheckOutInfo = True
        bDefault = True  '--- SCR 29659
        Exit Function
    End If
    '-- Get Default CheckOut Info'----------------------------------------------------------------------------------------
    If oRec.Fields.Count < 2 Then
        bDefault = True
        If Not GetClinicName(sClinicNCID, sClinicName) Then GoTo Cleanup
        
        sStmt = "select 1 as Job_Id, c.MedicationInfo, c.XRayInfo, c.LabInfo, c.ConsultInfo, c.HeaderInfo, c.ClinicInfo " & _
                "from CheckOutInfo c " & _
                "where c.Clinic_NCID = 0"
        
        Set oRec = goDBDAS.OpenRecordset(sStmt)
        Set oRec.ActiveConnection = Nothing
        
        oRec.Fields("HeaderInfo").Value = Replace(oRec.Fields("HeaderInfo").Value, "?", sClinicName)
        oRec.Fields("ClinicInfo").Value = Replace(oRec.Fields("ClinicInfo").Value, "?", sClinicName)
    Else
        bDefault = False
    End If
    
    GetCheckOutInfo = True
    GoTo Cleanup

ErrHandler:
    Call HandleError(Err, sMODULE_NAME & "." & "GetCheckOutInfo")
Cleanup:
End Function

Private Function GetCPSServicesDue(sPatientUnitNumber As String, _
                                  oRec As ADODB.Recordset, _
                                  ActiveDuty As String) As Boolean
On Error GoTo ErrHandler
    Const NCID_CANCELED As String = "11981"
    GetCPSServicesDue = False
    
    Dim sStmt As String
    
    CreateGemsStubRecordset oRec, "RecDate", "status", "Description"
   
    GetCPSServicesDue = True
    GoTo Cleanup
ErrHandler:
    Call HandleError(Err, sMODULE_NAME & "." & "GetCPSServicesDue")
Cleanup:
End Function
    
Private Function GetLabs(sEncounterNumber As String, _
                         oRec As ADODB.Recordset) As Boolean

On Error GoTo ErrHandler

    GetLabs = False
    
    Dim sStmt As String

'-- Get Lab Orders
'----------------------------------------------------------------------------------------
    '--- SCR 75487   Sherry Wang  12/9/2005
    If gobjshared.IsAppMode(modeITT) Or gobjshared.IsAppMode(modeLocalCache) _
        Or gobjshared.IsAppMode(modeTheater) Or gobjshared.IsAppMode(modeCDR) Then
        
      sStmt = "select order_name " _
          & " From oe_lab " _
          & " Where " _
          & " enc_num = '" & sEncounterNumber & "'"
   
      Set oRec = goDBDAS.OpenRecordset(sStmt)
   Else
      CreateGemsStubRecordset oRec, "order_name", "oe_lab"
   End If
   
   
    GetLabs = True
    GoTo Cleanup

ErrHandler:
    Call HandleError(Err, sMODULE_NAME & "." & "GetLabs")
Cleanup:
End Function
Private Sub Class_Initialize()
On Error GoTo ErrHandler
   If gobjshared Is Nothing Then Set gobjshared = New CWShared
    Set moPatientDemographics = New PatientDemographics
GoTo Cleanup
ErrHandler:
    Call HandleError(Err, sMODULE_NAME & "." & "Class_Initialize")
Cleanup:
End Sub

Private Sub Class_Terminate()
On Error GoTo ErrHandler
    Set moCollRecSet = Nothing
    Set moCollFormulas = Nothing
    Set moPatientDemographics = Nothing
GoTo Cleanup
ErrHandler:
    Call HandleError(Err, sMODULE_NAME & "." & "Class_Terminate")
Cleanup:
End Sub


