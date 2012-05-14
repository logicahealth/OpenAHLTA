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
Attribute VB_Name = "AppointmentOps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' copied from patient demographics--should be in ncidcons.bas
 Private Const NCID_SPONSOR_SSN As String = "186334"
 Private Const NCID_PATIENT_FMP As String = "204708"
 Private Const NCID_CURRENT_NAME As String = "187"

'SCR 13738, Phil Crowder 7/23/01  Added Oracle Hint to QuerySelect and reordered QueryWhere1 and QueryWhere2 and added QueryWhere3
'--- MSolano 11/29/01 corrected Begining / End of hint syntax: BOH: /*+   EOH: */
'/*+ INDEX(n mmi_name_x5) index(fmpx mmi_generic_id_x5) use_hash(a) */
Private Const sOracleSelect = _
      "Select " & vbCrLf _
    & " a.Facility_NCID, a.Appt_Id, a.Appt_IEN, a.Appt_Date_Time, a.CheckIn_Time, a.CheckOut_Time, " & vbCrLf _
    & " a.Appt_Comment, a.TELCON_URGENCY, a.Appt_Reason, t.Description as Appt_Type, " & vbCrLf _
    & " a.Appt_Classification, a.Appt_Status, a.Appt_Cancel_Reason, " & vbCrLf _
    & " a.Encounter_Number, a.Encounter_Status, a.Patient_Unit_Number, a.Order_Number, " & vbCrLf _
    & " a.Clinic_NCID, c.Name as Clinic_Name, a.Clinician_NCID, a.INJ_ILL_CATEGORY, a.INJ_ILL_CAUSE, " & vbCrLf _
    & " cp.Provider_Name as Provider_Name, c.MEPRS_Code, " & vbCrLf _
    & " m.Description as MEPRS_Description, a.Date_Created, a.Date_Modified, " & vbCrLf _
    & " trim(n.last_name || ',' || n.first_name || ' ' || n.middle_name) as Patient_Name, " & vbCrLf _
    & " site.medical_record_number as patient_ien, " & vbCrLf _
    & " id.ssn as patient_ssn, " & vbCrLf _
    & " trim(wp.area_code || ' ' || wp.local_number || ' ' || wp.internal_extension) " & vbCrLf _
    & " as patient_work_phone, " & vbCrLf _
    & " trim(hp.area_code || ' ' || hp.local_number || ' ' || hp.internal_extension) " & vbCrLf _
    & " as patient_home_phone , GetApptPendingComplaints(a.Appt_ID) AS Pending_Complaints, " & vbCrLf _
    & " fmp.id_value as patient_fmp, sssn.id_value as sponsor_ssn"

Private Const sMySQLSelect = _
      "Select " & vbCrLf _
    & " a.Facility_NCID, a.Appt_Id, a.Appt_IEN, a.Appt_Date_Time, a.CheckIn_Time, a.CheckOut_Time, " & vbCrLf _
    & " a.Appt_Comment, a.TELCON_URGENCY, a.Appt_Reason, t.Description as Appt_Type, " & vbCrLf _
    & " a.Appt_Classification, a.Appt_Status, a.Appt_Cancel_Reason, " & vbCrLf _
    & " a.Encounter_Number, a.Encounter_Status, a.Patient_Unit_Number, a.Order_Number, " & vbCrLf _
    & " a.Clinic_NCID, c.Name as Clinic_Name, a.Clinician_NCID, a.INJ_ILL_CATEGORY, a.INJ_ILL_CAUSE, " & vbCrLf _
    & " cp.Provider_Name as Provider_Name, c.MEPRS_Code, " & vbCrLf _
    & " m.Description as MEPRS_Description, a.Date_Created, a.Date_Modified, " & vbCrLf _
    & " concat(coalesce(n.last_name,''), ',', coalesce(n.first_name,''), ' ', coalesce(n.middle_name,'')) as Patient_Name, " & vbCrLf _
    & " site.medical_record_number as patient_ien, " & vbCrLf _
    & " id.ssn as patient_ssn, " & vbCrLf _
    & " concat(coalesce(wp.area_code,''), ' ', coalesce(wp.local_number,''), ' ', coalesce(wp.internal_extension,'')) " & vbCrLf _
    & " as patient_work_phone, " & vbCrLf _
    & " concat(coalesce(hp.area_code,''), ' ', coalesce(hp.local_number,''), ' ', coalesce(hp.internal_extension,'')) " & vbCrLf _
    & " as patient_home_phone , GetApptPendingComplaints(a.Appt_ID) AS Pending_Complaints, " & vbCrLf _
    & " fmp.id_value as patient_fmp, sssn.id_value as sponsor_ssn"

Private Const sWhereOracle = _
      " From Clinic c, Meprs_Code m, Appt_Type t, " & vbCrLf _
    & "    Clinic_Provider cp, mmi_name n, mmi_site site, mmi_id id, " & vbCrLf _
    & "    mmi_phone wp, mmi_phone hp, mmi_generic_id fmp, mmi_generic_id sssn, " & vbCrLf _
    & "    Appointment a " & vbCrLf _
    & "Where n.unit_number = wp.unit_number(+) and wp.phone_type_cid(+)= " + NCID_WORK_PHONE & vbCrLf _
    & "  and n.unit_number = hp.unit_number(+) and hp.phone_type_cid(+)= " + NCID_HOME_PHONE _
    & "  and n.unit_number = fmp.unit_number(+) and fmp.id_type_ncid(+) = " & NCID_PATIENT_FMP & vbCrLf _
    & "  and n.unit_number = sssn.unit_number(+) and sssn.id_type_ncid(+) = " & NCID_SPONSOR_SSN & vbCrLf _

Private Const sWhereMySQL = "  FROM    (   (   (   (   (   (   (   (   (   (   gems.mmi_site site" _
                                               & " Inner Join gems.Appointment a" _
                                               & " ON (site.FACILITY_CID = a.FACILITY_NCID))" _
                                           & " Inner Join gems.meprs_code m" _
                                           & " ON (m.FACILITY_NCID = a.FACILITY_NCID))" _
                                        & " Inner Join gems.mmi_name n" _
                                       & " ON (n.UNIT_NUMBER = a.PATIENT_UNIT_NUMBER) AND (site.UNIT_NUMBER = n.UNIT_NUMBER) and n.NAME_TYPE_CID = 187)" _
                                   & " LEFT OUTER JOIN gems.mmi_phone wp" _
                                   & " ON (n.UNIT_NUMBER = wp.UNIT_NUMBER) and wp.PHONE_TYPE_CID = 194)" _
                               & " LEFT OUTER JOIN gems.mmi_phone hp" _
                               & " ON (n.UNIT_NUMBER = hp.UNIT_NUMBER) and hp.PHONE_TYPE_CID = 193)" _
                           & " LEFT OUTER JOIN gems.mmi_generic_id FMP" _
                           & " ON (n.UNIT_NUMBER = fmp.UNIT_NUMBER) and fmp.ID_TYPE_NCID = 204708)" _
                       & " LEFT OUTER JOIN gems.mmi_generic_id sssn" _
                       & " ON (n.UNIT_NUMBER = sssn.UNIT_NUMBER) and sssn.ID_TYPE_NCID = 186334)" _
                   & " Inner Join gems.Clinic c" _
                   & " ON (c.NCID = a.CLINIC_NCID) AND (c.MEPRS_IEN = m.IEN))" _
               & " Inner Join gems.APPT_TYPE t" _
               & " ON (t.CODE = a.APPT_TYPE) AND (t.FACILITY_NCID = a.FACILITY_NCID))" _
           & " Inner Join gems.clinic_provider cp" _
           & " ON (cp.CLINIC_NCID = a.CLINIC_NCID) AND (cp.CLINICIAN_NCID = a.CLINICIAN_NCID))" _
       & " Inner Join gems.mmi_id ID" _
       & " ON (id.UNIT_NUMBER = n.UNIT_NUMBER)"
       
Private Const sQueryWhere2MySQL = "" _
    & "  Where (fmp.create_audit_num = (select max(create_audit_num)" & vbCrLf _
    & "                                from mmi_generic_id fmpx" & vbCrLf _
    & "                                Where FMP.unit_number = fmpx.unit_number" & vbCrLf _
    & "                                and fmpx.id_type_ncid = " & NCID_PATIENT_FMP & ")" & vbCrLf _
    & "         or not exists (select 'x'" & vbCrLf _
    & "                        from mmi_generic_id fmpx" & vbCrLf _
    & "                        Where FMP.unit_number = fmpx.unit_number" & vbCrLf _
    & "                        and fmpx.id_type_ncid = " & NCID_PATIENT_FMP & "))" & vbCrLf _
    & "    and (sssn.create_audit_num = (select max(create_audit_num)" & vbCrLf _
    & "                                 from mmi_generic_id sssnx" & vbCrLf _
    & "                                 Where sssn.unit_number = sssnx.unit_number" & vbCrLf _
    & "                                 and sssnx.id_type_ncid = " & NCID_SPONSOR_SSN & ")" & vbCrLf _
    & "         or not exists (select 'x'" & vbCrLf _
    & "                        from mmi_generic_id sssnx" & vbCrLf _
    & "                        Where sssn.unit_number = sssnx.unit_number" & vbCrLf _
    & "                        and sssnx.id_type_ncid = " & NCID_SPONSOR_SSN & "))" & vbCrLf
   
Private Const sQueryWhere2Oracle = "  and c.NCID = a.Clinic_NCID " & vbCrLf _
    & "  And t.Code = a.Appt_Type  and t.facility_ncid = a.facility_ncid" & vbCrLf _
    & "  And c.Meprs_IEN = m.IEN and m.facility_ncid = a.facility_ncid" & vbCrLf _
    & "  and cp.Clinic_NCID = a.Clinic_NCID " & vbCrLf _
    & "  and cp.Clinician_NCID = a.Clinician_NCID " & vbCrLf _
    & "  and n.unit_number = a.patient_unit_number and n.name_type_cid = " & NCID_CURRENT_NAME & vbCrLf _
    & "  and site.facility_cid = a.facility_ncid and site.unit_number = n.unit_number " & vbCrLf _
    & "  and id.unit_number = n.unit_number " & vbCrLf _
    & "  and (fmp.create_audit_num = (select max(create_audit_num)" & vbCrLf _
    & "                                from mmi_generic_id fmpx" & vbCrLf _
    & "                                Where FMP.unit_number = fmpx.unit_number" & vbCrLf _
    & "                                and fmpx.id_type_ncid = " & NCID_PATIENT_FMP & ")" & vbCrLf _
    & "         or not exists (select 'x'" & vbCrLf _
    & "                        from mmi_generic_id fmpx" & vbCrLf _
    & "                        Where FMP.unit_number = fmpx.unit_number" & vbCrLf _
    & "                        and fmpx.id_type_ncid = " & NCID_PATIENT_FMP & "))" & vbCrLf _
    & "    and (sssn.create_audit_num = (select max(create_audit_num)" & vbCrLf _
    & "                                 from mmi_generic_id sssnx" & vbCrLf _
    & "                                 Where sssn.unit_number = sssnx.unit_number" & vbCrLf _
    & "                                 and sssnx.id_type_ncid = " & NCID_SPONSOR_SSN & ")" & vbCrLf _
    & "         or not exists (select 'x'" & vbCrLf _
    & "                        from mmi_generic_id sssnx" & vbCrLf _
    & "                        Where sssn.unit_number = sssnx.unit_number" & vbCrLf _
    & "                        and sssnx.id_type_ncid = " & NCID_SPONSOR_SSN & "))" & vbCrLf
    
Private Const sQueryWhere3 = "  and a.appt_classification <> 3  and a.appt_classification <> 1 "

Public Enum ApptTransTypeEnum
    ApptTransInsert = 1
    ApptTransProvider = 2
    ApptTransClinic = 3
    ApptTransCancel = 4
    ApptTransKept = 5
    ApptTransLWOBS = 6
    ApptTransTelconComplete = 7
    ApptTransInpatient = 8
    ApptTransUndoCancel = 5
End Enum
Public Enum AppointmentPriority  'SCR-42772
    ApptNone = 0
    ApptMinor = 1
    ApptUrgent = 2
    ApptEmergent = 3
End Enum

Public Enum AppointmentClassifications
    ApptOutpatient = 0
    ApptInpatient = 1
    ApptAmbulatoryProcedure = 2
    ApptTelCon = 3 ' if this should change, need to change sQueryWhere2Oracle above and in TelconOps
End Enum

'--- OTHER MENUS SET UP BY CORE THAT
'--- SET APPT SELECTION CRITERIA
Public Enum AppointmentFilters
    apptDefault = 0
    apptAll = 1
    apptToday = 2
    apptTodayPlusIncomplete = 4
    apptPending = 8
    apptCheckedIn = 16
    apptWaiting = 32
    apptInProgress = 64
    apptNeedsCoSignature = 128
    apptUpdated = 256
    apptCompleted = 512
    apptIncomplete = 1024
    apptCheckedInPending = 2048
    apptOneDate = 4096
    apptDateRange = 8192
    
End Enum

Public Enum CHCS_GUI_ApptFilters
    CHCS_GUIapptDefault = 0
    CHCS_GUIapptAnyStatus = 1
    CHCS_GUIapptPending = 8
    CHCS_GUIapptCheckedIn = 16
    CHCS_GUIapptCancelled = 32
    'CHCS_GUIhealthHistory = 256
End Enum


Public Enum AppointmentMenuActions
    '--- arbitrary constants - values choosen such that
    '--- they don 't line up w/ AppointmentFilters
    '--- so they can exist side by side in goCoreMenuID
    apptMnuActRefreshMenuId = 33
    apptMnuActNewMenuId = 34
    apptMnuActCheckinMenuId = 35
    apptMnuActCheckoutMenuId = 36
    apptMnuActCancelMenuId = 37
    apptMnuActTransferMenuId = 38
    apptMnuActAddProvider = 39
    apptMnuActPrint = 40
    apptMnuActComments = 41
    apptMnuActPrintInsuranceForm = 42
    apptMnuActUndoCancel = 43
    'SCR-31309
    apptMnuActOpenMenuId = 44
End Enum
'<<< SCR-18843

Private sSQLStatement                               As String
Private msIBWAClinicianFacilityNCID                 As String           'SCR-RNDS
Private msCriteriaEncounterNumber                   As String
Private mdCriteriaStartDate                         As Date
Private mdCriteriaEndDate                           As Date
Private mlCriteriaUnitNumber                        As Long
Private mlProgId                                    As Long
Private mbRNDS_On_and_NonTheater_mode               As Boolean
Private mbTelcons_Enabled_CHCS_GUI                  As Boolean 'SCR-39944
Private mbGEMSUserHasPriv_Data_VIP                  As Boolean

Private mobjEncParent                               As CHCSIIEncounterCurrent.EncounterParent
'Private mobjSQL                                     As ICHCSII_SQL.ISqlOps
Private mobjLogon                                   As MMMHISLogon.Logon
Private mobjConfig                                  As MMMHISConfiguration.Config
Private mobjComm                                    As MMMHISComm.Comm
Private mobjPatient                                 As MMMHISPatient.Patient
Private mobjUser                                    As CHCSII_ClientSecurity.ICHCSII_User
Private mobjShared                                  As CWShared
Private gobjClinicOps                               As CHCSII_ClinicClient.ClinicOpsEx
Private menumCriteriafilter                         As AppointmentFilters
'Private menumCHCS_GUICriteriafilter                 As CHCS_GUI_ApptFilters

Private mcClinicianNCIDs                            As Collection
Private mcClinicNCIDs                               As Collection
Private mcIBWAClinicianNCID                         As Collection       'SCR-RNDS
Private mcIBWAClinicNCIDs                           As Collection       'SCR-RNDS

Private mcolAppointmentsCache                       As Collection       'SCR-32352
                                                    'Any place where we are doing a "GetAppointment" we will check the
                                                    'cAppointments collection first instead of pulling from the database.
                                                    'mcolAppointmentCache is a dup of mcolAppontments in Appointments.frm
                                                    'Applies only to CDRmode


Public Property Get SQLStatement() As String
    SQLStatement = sSQLStatement
End Property

Public Property Let SQLStatement(ByVal sNewValue As String)
    sSQLStatement = sNewValue
End Property

Public Property Get AppointmentsCache() As Collection
    Set AppointmentsCache = mcolAppointmentsCache
End Property

Public Property Get SelectedClinicCount() As Long
    If mcClinicNCIDs Is Nothing Then
        SelectedClinicCount = 0
    Else
        SelectedClinicCount = mcClinicNCIDs.Count
    End If
End Property
Public Property Get SelectedIBWAClinicCount() As Long
    If mcIBWAClinicNCIDs Is Nothing Then
        SelectedIBWAClinicCount = 0
    Else
        SelectedIBWAClinicCount = mcIBWAClinicNCIDs.Count
    End If
End Property

Public Property Get SelectedClinicianCount() As Long
    If mcClinicianNCIDs Is Nothing Then
        SelectedClinicianCount = 0
    Else
        SelectedClinicianCount = mcClinicianNCIDs.Count
    End If
End Property
Public Property Get SelectedIBWAClinicianCount() As Long
    If mcIBWAClinicianNCID Is Nothing Then
        SelectedIBWAClinicianCount = 0
    Else
        SelectedIBWAClinicianCount = mcIBWAClinicianNCID.Count
    End If
End Property

Public Property Let IBWAClinicianFacilityNCID(sNCID As String)  'SCR-RNDS
    msIBWAClinicianFacilityNCID = sNCID
End Property
Public Property Get IBWAClinicianFacilityNCID() As String       'SCR-RNDS
    IBWAClinicianFacilityNCID = msIBWAClinicianFacilityNCID
End Property
Public Property Let CriteriaEncounterNumber(rsEncounterNumber As String)
    msCriteriaEncounterNumber = rsEncounterNumber
End Property

Public Property Let CriteriaStartDate(rdDate As Date)
    mdCriteriaStartDate = rdDate
End Property

Public Property Let CriteriaEndDate(rdDate As Date)
    mdCriteriaEndDate = rdDate
End Property
'Public Property Let CHCS_GUICriteriaFilter(renumfilter As CHCS_GUI_ApptFilters)
'    menumCHCS_GUICriteriafilter = renumfilter
'End Property
Public Property Let CriteriaFilter(renumfilter As AppointmentFilters)
    menumCriteriafilter = renumfilter
End Property
Public Property Let CriteriaUnitNumber(rlUnitNumber As Long)
    mlCriteriaUnitNumber = rlUnitNumber
End Property
Public Sub ClearCriteria_CHCS_GUI()
    Set mcIBWAClinicianNCID = New Collection
    Set mcIBWAClinicNCIDs = New Collection
    msIBWAClinicianFacilityNCID = vbNullString
End Sub
Public Sub CriteriaAddIBWAClinician(rsClinicianNCID As String) 'SCR-RNDS
On Error Resume Next
    'There should only be one IBWA Clinician
    mcIBWAClinicianNCID.Add rsClinicianNCID, rsClinicianNCID
End Sub
Public Sub CriteriaAddIBWAClinic(rsClinicNCID As String)        'SCR-RNDS
On Error Resume Next
    mcIBWAClinicNCIDs.Add rsClinicNCID, rsClinicNCID
End Sub
Public Sub CriteriaAddClinician(rsClinicianNCID As String)
On Error Resume Next
    mcClinicianNCIDs.Add rsClinicianNCID
End Sub

Public Sub CriteriaAddClinic(rsClinicNCID As String)
    On Error Resume Next
    mcClinicNCIDs.Add rsClinicNCID, rsClinicNCID
End Sub

Public Sub ClearCriteria()
On Error GoTo ErrHandler
    Set mcClinicianNCIDs = New Collection
    Set mcClinicNCIDs = New Collection
    Set mcIBWAClinicianNCID = New Collection
    Set mcIBWAClinicNCIDs = New Collection
    msIBWAClinicianFacilityNCID = vbNullString
    mdCriteriaStartDate = 0
    mdCriteriaEndDate = 0
    msCriteriaEncounterNumber = vbNullString
    mlCriteriaUnitNumber = 0
    menumCriteriafilter = apptAll
Exit Sub
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "AppointmentsOps.cls - ClearCriteria. ", App.Title, vbCritical

End Sub

Public Function GetInpatientAdmissionsData(sPatientIEN As String, sFacilityNCID As String) As ADODB.Recordset
Dim rs As ADODB.Recordset
On Error GoTo ErrHandler
    If mobjShared.IsAppMode(modeITT) Then
        Set rs = GetInpatientAdmissionsData_SQL(sPatientIEN, sFacilityNCID)
    ElseIf mobjShared.IsAppMode(modeCHCSI_GUI) Then
        Set rs = GetInpatientAdmissionsData_CHCS_GUI(sPatientIEN)
    Else
        Set rs = GetInpatientAdmissionsData_SP(sPatientIEN, sFacilityNCID)
    End If
    If rs Is Nothing Then Exit Function
    If rs.EOF And rs.BOF Then Exit Function
    If rs.RecordCount > 0 Then
        ApplyInpatientAdmissionFilter rs
    End If
    Set GetInpatientAdmissionsData = rs
    Set rs = Nothing
Exit Function
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "AppointmentsOps.cls - GetInpatientAdmissionsData. ", App.Title, vbCritical
    Set rs = Nothing
End Function
Private Sub ApplyInpatientAdmissionFilter(ByRef rs As ADODB.Recordset)
'  --c_RON                              CONSTANT VARCHAR(16):= 'REMAIN OVERNIGHT';
'  --c_CRO                              CONSTANT VARCHAR(22):= 'CARDED FOR RECORD ONLY';
Const c_RON As String = "REMAIN OVERNIGHT"
Const c_CRO As String = "CARDED FOR RECORD ONLY"
Dim sAdmissType As Variant
Dim dDisDate As Variant
Dim dMedHoldDisDate As Variant
Dim dCurrDate As Date
Dim sRegNum As Variant
Dim dAdmissDate As Variant
Dim NewRS As ADODB.Recordset

On Error GoTo ErrHandler
    dCurrDate = CDate(Format$(Now, "mm/dd/yyyy"))
    Set NewRS = CreateNewInpatientRS
    rs.MoveFirst
    Do While rs.EOF <> True
        sAdmissType = rs("ADMISSION_TYPE")
        dDisDate = rs("DISPOSITION_DATE")
        dMedHoldDisDate = rs("MEDICAL_HOLD_DISPOSITION_DATE")
        sRegNum = rs("REGISTER_NUMBER")
        dAdmissDate = rs("ADMISSION_DATE")
        SetInpatientAdmissionVars sRegNum, sAdmissType, dDisDate, dMedHoldDisDate, dAdmissDate
        If sRegNum = vbNullString Then GoTo NextRecord
        If sAdmissType = c_RON Or sAdmissType = c_CRO Then GoTo NextRecord
        If DateDiff("d", dDisDate, dCurrDate) > 30 Then
            If dMedHoldDisDate <= dCurrDate Then GoTo NextRecord
        End If
        With NewRS
            .AddNew
            .Fields("REGISTER_NUMBER") = sRegNum
            .Fields("ADMISSION_DATE") = GetDateAsString(rs("ADMISSION_DATE"))
            .Fields("DISPOSITION_DATE") = GetDateAsString(rs("DISPOSITION_DATE"))
            .Fields("ADMISSION_TYPE") = rs("ADMISSION_TYPE") & vbNullString
            .Fields("ADMISSION_SYNONYM") = rs("ADMISSION_SYNONYM") & vbNullString
            .Fields("DISPOSITION_TYPE") = rs("DISPOSITION_TYPE") & vbNullString
            .Fields("ABSENT") = rs("ABSENT") & vbNullString
            .Fields("MEDICAL_HOLD_DISPOSITION_DATE") = GetDateAsString(rs("MEDICAL_HOLD_DISPOSITION_DATE"))
            If IsNull(rs("MEPRS_CODE")) Then
                .Fields("MEPRS_CODE") = vbNullString
            Else
                .Fields("MEPRS_CODE") = rs("MEPRS_CODE")
            End If
            If IsNull(rs("MEPRS_IEN")) Then
                .Fields("MEPRS_IEN") = vbNullString
            Else
                .Fields("MEPRS_IEN") = rs("MEPRS_IEN")
            End If
            
        End With
NextRecord:
        rs.MoveNext
    Loop
    If NewRS.RecordCount > 0 Then
        Set rs = NewRS
    Else
        Set rs = Nothing
    End If
    Set NewRS = Nothing
    
Exit Sub
ErrHandler:
    Set rs = Nothing
    Set NewRS = Nothing
    mobjShared.ShowVBError Err.Number, Err.Description, "AppointmentsOps.cls - ApplyInpatientAdmissionFilter. ", App.Title, vbCritical

End Sub
Private Sub SetInpatientAdmissionVars(ByRef sRegNum As Variant, _
                                      ByRef sAdmissType As Variant, _
                                      ByRef dDisDate As Variant, _
                                      ByRef dMedHoldDisDate As Variant, _
                                      ByRef dAdmissDate As Variant)
On Error GoTo ErrHandler
        If IsNull(sRegNum) Then
            sRegNum = vbNullString
        ElseIf IsNumeric(sRegNum) = False Then
            sRegNum = vbNullString
        End If
        If IsNull(sAdmissType) Then
            sAdmissType = vbNullString
        End If
        If IsNull(dDisDate) Then
            dDisDate = CDate(Format$(Now, "mm/dd/yyyy"))
            'The reason we set to "Now" instead of "0" is that when doing the date calculation,
            'the datediff function does evaluate our business rule correctly.
        ElseIf CDate(Format$(dDisDate, "mm/dd/yyyy")) = CDate("12/31/9999") Then
            'the SP class in CHCS Conn seems to be returning this date when it is really null.
            dDisDate = CDate(Format$(Now, "mm/dd/yyyy"))
        ElseIf IsDate(dDisDate) Then
            dDisDate = CDate(Format$(dDisDate, "mm/dd/yyyy"))
        Else
            dDisDate = 0
        End If
        If IsNull(dMedHoldDisDate) Then
            dMedHoldDisDate = CDate(Format$(Now, "mm/dd/yyyy"))
        ElseIf CDate(Format$(dMedHoldDisDate, "mm/dd/yyyy")) = CDate("12/31/9999") Then
            'the SP class in CHCS Conn seems to be returning this date when it is really null.
            dMedHoldDisDate = CDate(Format$(Now, "mm/dd/yyyy"))
        ElseIf IsDate(dMedHoldDisDate) Then
            dMedHoldDisDate = CDate(Format$(dMedHoldDisDate, "mm/dd/yyyy"))
        Else
            dMedHoldDisDate = 0
        End If
        If IsNull(dAdmissDate) Then
            dAdmissDate = 0
        ElseIf CDate(Format$(dAdmissDate, "mm/dd/yyyy")) = CDate("12/31/9999") Then
            'the SP class in CHCS Conn seems to be returning this date when it is really null.
            dAdmissDate = 0
        ElseIf IsDate(dAdmissDate) Then
            dAdmissDate = CDate(Format$(dAdmissDate, "mm/dd/yyyy"))
        Else
            dAdmissDate = 0
        End If
        
Exit Sub
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "AppointmentsOps.cls - SetInpatientAdmissionVars. ", App.Title, vbCritical

End Sub
Private Function CreateNewInpatientRS() As ADODB.Recordset
On Error GoTo ErrHandler
    Set CreateNewInpatientRS = New ADODB.Recordset
    DoEvents
    With CreateNewInpatientRS
      Set .ActiveConnection = Nothing
      .CursorLocation = adUseClient
      .LockType = adLockBatchOptimistic
    
        With .Fields
            .Append "REGISTER_NUMBER", adBSTR, 25
            .Append "ADMISSION_DATE", adBSTR, 25
            .Append "DISPOSITION_DATE", adBSTR, 25
            .Append "ADMISSION_TYPE", adBSTR, 110
            .Append "ADMISSION_SYNONYM", adBSTR, 110
            .Append "DISPOSITION_TYPE", adBSTR, 110
            .Append "ABSENT", adBSTR, 3
            .Append "MEDICAL_HOLD_DISPOSITION_DATE", adBSTR, 25
            .Append "MEPRS_CODE", adBSTR, 10
            .Append "MEPRS_IEN", adBSTR, 20
         End With
        If .State = 0 Then .Open
    End With
    
Exit Function
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "AppointmentsOps.cls - ApplyInpatientAdmissionFilter. ", App.Title, vbCritical

End Function

Private Function GetInpatientAdmissionsData_CHCS_GUI(sPatientIEN As String) As ADODB.Recordset
'DataLayer Item
Dim oCHCS_AdmissionData As CHCS_Interface.ICHCSPatient
On Error GoTo ErrHandler
    
    Set oCHCS_AdmissionData = mobjShared.CHCSConnection
    Set GetInpatientAdmissionsData_CHCS_GUI = New ADODB.Recordset
    Set GetInpatientAdmissionsData_CHCS_GUI = oCHCS_AdmissionData.RetrieveInpatientAdmissions(sPatientIEN)
Exit Function

ErrHandler:
        mobjShared.ShowVBError Err.Number, Err.Description + vbCrLf + _
        "Error in retrieving inpatient admissions data from CHCS. ", _
        "AppointmentsOps.cls - GetInpatientAdmissionsData_CHCS_GUI. ", App.Title, vbCritical

End Function
Private Function GetInpatientAdmissionsData_SQL(ByVal sPatientIEN As String, ByVal sFacilityNCID As String) As ADODB.Recordset
    Dim strSQL As String
    Dim objConn As CHCSII_CONN.Conn
  On Error GoTo ErrHandler
    strSQL = "SELECT REGISTER_NUMBER, ADMISSION_DATE, ADMISSION_TYPE, " & _
            "ADMISSION_SYNONYM, DISPOSITION_DATE, DISPOSITION_TYPE, ABSENT, " & _
            "MEDICAL_HOLD_DISPOSITION_DATE, INPATIENT_MEPRS_CODE AS MEPRS_CODE, " & _
            "INPATIENT_MEPRS_IEN As MEPRS_Ien " & _
            " From PATIENT_ADMISSIONS  Where PATIENT_IEN = " & sPatientIEN & _
            "  AND FACILITY_NCID = " & sFacilityNCID

    Set objConn = New CHCSII_CONN.Conn
    
    Set GetInpatientAdmissionsData_SQL = objConn.CHCSII_DAS.ExecuteSQL(strSQL)
Cleanup:
    Set objConn = Nothing
    Exit Function
ErrHandler:
        mobjShared.ShowVBError Err.Number, Err.Description & vbCrLf & _
        "Error in retrieving inpatient admissions data", "AppointmentsOps.cls - " & _
        "GetInpatientAdmissionsData_SQL. ", App.Title, vbCritical
 
    GoTo Cleanup
End Function
Private Function GetInpatientAdmissionsData_SP(sPatientIEN As String, sFacilityNCID As String) As ADODB.Recordset
'Returns a collection of mAdmission_Data; from the Patient_Admissions table.
On Error GoTo ErrHandler
Dim oDatabase As CHCSII_CONN.Conn
Dim objCmd As CHCSII_Command.CHCSII_Cmd

    If mobjDASsp Is Nothing Then Set mobjDASsp = Set_CHCSII_CONN_SP_connection()
    Set objCmd = New CHCSII_Command.CHCSII_Cmd
    objCmd.SPName = CURR_SP_VER + ".Get_InpatientAdmissionData"
    DoEvents
    
    If Not Set_Parms_for_SP(objCmd, , , , , CLng(sFacilityNCID), , , , CLng(sPatientIEN)) Then
        Set GetInpatientAdmissionsData_SP = Nothing
        GoTo Cleanup
    End If
    Set GetInpatientAdmissionsData_SP = mobjDASsp.OpenRecordsetSP(objCmd)
Cleanup:
    Set objCmd = Nothing
    Set oDatabase = Nothing
    Exit Function
ErrHandler:
        mobjShared.ShowVBError Err.Number, Err.Description + vbCrLf + _
        "Error in retrieving inpatient admissions data from Stored " + _
        "Procedure: Appointments_Vr3_0_Pkg.Get_InpatientAdmissionsData ", _
        "AppointmentsOps.cls - GetInpatientAdmissionsData_SP. ", App.Title, vbCritical
 
    GoTo Cleanup

End Function

Public Function Insert_PatientAdmissionRecord_SQL(sApptID As String, sRegisterNumber As String) As Boolean
    Dim strSQL As String
    On Error GoTo ErrHandler
    'SCR-53141
    strSQL = "INSERT INTO APPT_INPATIENT_ADMISSIONS VALUES(" & sApptID & ", " & sRegisterNumber & ")"
    mobjSQL.Execute strSQL
    Exit Function
    
ErrHandler:
        mobjShared.ShowVBError Err.Number, Err.Description + vbCrLf + _
        "Error in inserting inpatient admissions data.", "AppointmentsOps.cls - Insert_PatientAdmissionRecord_SQL. ", App.Title, vbCritical
End Function
Public Function Insert_PatientAdmissionRecord_SP(sApptID As String, sRegisterNumber As String) As Boolean
'DataLayer Item
On Error GoTo ErrHandler
Dim oDatabase As CHCSII_CONN.Conn
Dim objCmd As CHCSII_Command.CHCSII_Cmd
Dim rs As ADODB.Recordset
Dim vData As Variant

    If mobjDASsp Is Nothing Then Set mobjDASsp = Set_CHCSII_CONN_SP_connection()
    Set objCmd = New CHCSII_Command.CHCSII_Cmd
    objCmd.SPName = CURR_SP_VER + ".Insert_PatientAdmissionRecord"
    DoEvents
    
    Call objCmd.AddParam(False, False, "textString", sApptID)
    Call objCmd.AddParam(False, False, "textString", sRegisterNumber)
    'This Parameter is defined as a cursor, used for output
    Call objCmd.AddParam(True, True, vbNullString, vbNullString)
    
    Set rs = mobjDASsp.OpenRecordsetSP(objCmd)
    If rs.EOF = True Then
        Insert_PatientAdmissionRecord_SP = False
    Else
        If IsNull(rs("RETURN_VALUE")) = False Then
            If InStr(1, rs("RETURN_VALUE"), "TRUE", vbTextCompare) > 0 Then
                Insert_PatientAdmissionRecord_SP = True
            Else
                Insert_PatientAdmissionRecord_SP = False
                vData = Split(rs("RETURN_VALUE"), "|", , vbTextCompare)
                If UBound(vData) = 1 Then
                    MsgBxARMd "Unable to insert the Inpatient Admission record for ApptID: " + sApptID + _
                        ". Record already exisits with Register number: " + vData(1)
                End If
            End If
        Else
            Insert_PatientAdmissionRecord_SP = False
        End If
    End If

Cleanup:
    Set objCmd = Nothing
    Set oDatabase = Nothing
    Exit Function
ErrHandler:
        mobjShared.ShowVBError Err.Number, Err.Description + vbCrLf + _
        "Error in inserting inpatient admissions data from Stored Procedure: Appointments_Vr3_0_Pkg.Insert_PatientAdmissionRecord ", "AppointmentsOps.cls - Insert_PatientAdmissionRecord_SP. ", App.Title, vbCritical
 
    GoTo Cleanup


End Function

Public Function DoSearch_GetIBWA_NCIDS_CHCS_GUI(sFacilityNCID As String) As Collection
'DataLayer Item
Dim sSql As String

On Error GoTo ErrHandler
    
    Set DoSearch_GetIBWA_NCIDS_CHCS_GUI = New Collection
    sSql = "SELECT NCID FROM CLINIC WHERE  name LIKE 'IBWA%' AND Facility_NCID = " + sFacilityNCID
    If mobjSQL Is Nothing Then Set mobjSQL = SelectSqlOps()
    mobjSQL.Execute sSql
    If mobjSQL.BOF And mobjSQL.EOF Then
        Exit Function
    End If
    mobjSQL.MoveFirst
    Do While mobjSQL.EOF = False
        If IsNull(mobjSQL("NCID")) = False Then
            On Error Resume Next
            DoSearch_GetIBWA_NCIDS_CHCS_GUI.Add CStr(mobjSQL("NCID")), CStr(mobjSQL("NCID"))
        End If
        mobjSQL.MoveNext
    Loop

    Exit Function
    
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "AppointmentsOps.cls - DoSearch_GetIBWA_NCIDS_CHCS_GUI. ", App.Title, vbCritical


End Function

Public Function DoSearch_GetIBWA_NCIDS_SP(sFacilityNCID As String) As Collection

On Error GoTo ErrHandler

Dim oDatabase As CHCSII_CONN.Conn
Dim objCmd As CHCSII_Command.CHCSII_Cmd
Dim rs As ADODB.Recordset
    
    
    Set DoSearch_GetIBWA_NCIDS_SP = New Collection
    If mobjDASsp Is Nothing Then Set mobjDASsp = Set_CHCSII_CONN_SP_connection()
    Set objCmd = New CHCSII_Command.CHCSII_Cmd
    objCmd.SPName = CURR_SP_VER + ".Get_NCIDs_For_IBWA_Clinics"
    DoEvents
    
    Call objCmd.AddParam(False, False, "textString", sFacilityNCID)
    'This Parameter is defined as a cursor, used for output
    Call objCmd.AddParam(True, True, vbNullString, vbNullString)
    
    Set rs = mobjDASsp.OpenRecordsetSP(objCmd)
    'If successful, SP Returns:
    'TRUE|<1>'
    If rs Is Nothing Then GoTo Cleanup
    If rs.EOF Then GoTo Cleanup
    rs.MoveFirst
    Do While rs.EOF = False
        If IsNull(rs("NCID")) = False Then
            On Error Resume Next
            DoSearch_GetIBWA_NCIDS_SP.Add CStr(rs("NCID")), CStr(rs("NCID"))
        End If
        rs.MoveNext
    Loop

Cleanup:
    Set objCmd = Nothing
    Set oDatabase = Nothing
    Set rs = Nothing
    Exit Function
ErrHandler:
        mobjShared.ShowVBError Err.Number, Err.Description + vbCrLf + _
        "Error in retrieving IBWA Clinics from Stored Procedure: Appointments_Vr3_0_Pkg.Get_NCIDs_For_IBWA_Clinics ", "AppointmentsOps.cls - DoSearch_GetIBWA_NCIDS_SP. ", App.Title, vbCritical
 
    
    GoTo Cleanup
End Function
Public Function Appts_Check_PKG_Version_SP() As Boolean
    On Error GoTo ErrHandler
    
    If Not Check_PKG_Version_SP() Then
        Appts_Check_PKG_Version_SP = False
    Else
        Appts_Check_PKG_Version_SP = True
    End If
Exit Function

ErrHandler:
        mobjShared.ShowVBError Err.Number, Err.Description + vbCrLf + _
        "Error from Stored Procedure: Appointments_Vr3_0_Pkg.Appt_PKG_Version_Check ", "AppointmentsOps.cls - Appts_Check_PKG_Version_SP. ", App.Title, vbCritical

End Function


Public Sub Initialize(ByVal vlMyProgId As Long, ByRef robjComm As Object, ByRef robjLogonEnvironment As Object, ByRef robjConfig As Object, ByRef robjPatient As Object, ByRef robjUser As Object)
    On Error GoTo ErrHandler
    
    Set mobjLogon = robjLogonEnvironment
    Set mobjConfig = robjConfig
    Set mobjUser = robjUser
    Set mobjComm = robjComm
    Set mobjPatient = robjPatient
    
    mlProgId = vlMyProgId
    
    If mobjSQL Is Nothing Then Set mobjSQL = SelectSqlOps()
    If mobjDASsp Is Nothing Then Set mobjDASsp = Set_CHCSII_CONN_SP_connection()

    Set mcClinicianNCIDs = New Collection
    Set mcClinicNCIDs = New Collection
    Set mcIBWAClinicianNCID = New Collection 'SCR-RNDS
    Set mcIBWAClinicNCIDs = New Collection
    Set mcolAppointmentsCache = New Collection
    If mobjShared.IsAppMode(modeCHCSI_GUI) Then 'SCR-39944
        mbTelcons_Enabled_CHCS_GUI = mobjUser.GetAttribute("TelconsEnabled")
    End If
    If LenB(mobjShared.CmdLineSwitches.Item("RNDS")) > 0 And Not mobjShared.IsAppMode(modeTheater) Then       'SCR-RNDS
        mbRNDS_On_and_NonTheater_mode = True
    End If
    
    'ITT convergence - Builds the SQL SERVER query for appts.
    If mobjShared.IsAppMode(modeITT) Then Build_gsSelectSQLServer
    
    'Placing here instead of the DoSearch_GEM function so that it will only
    'need to be called once, on each call appt query.
    If mobjShared.IsAppMode(modeTheater) Then
        If mobjUser.HasPrivilegeEx(Priv_Data_VIP, Priv_Read) Then 'SCR-42772
            mbGEMSUserHasPriv_Data_VIP = True
        End If
    End If
Exit Sub

ErrHandler:
        MsgBxARMd "Error initializing AppointmentOps. " + CStr(Err.Number) + ", " + Err.Description
        
End Sub
Public Function ApptsInLastThreeYears(sPatientUnitNumber As String, sClinicNCID As String, sCurrApptID As String) As Boolean
Dim sSql As String
Dim sDateWhere As String
Const sSQLServerDate As String = " DATEADD(month, -36, GetDate()) "
Const sOracleDate As String = " ADD_MONTHS(SYSDATE, -36)"
Const sMySQLDate As String = " Date_ADD(sysdate() - INTERVAL 3 MONTH)"
On Error GoTo ErrHandler
    If LenB(sPatientUnitNumber) = 0 Or IsNumeric(sPatientUnitNumber) = False Then
        MsgBxARMd "Error. Invalid Patient Unit Number:" + sPatientUnitNumber
        ApptsInLastThreeYears = False
        Exit Function
    End If
    If LenB(sClinicNCID) = 0 Or IsNumeric(sClinicNCID) = False Then
        MsgBxARMd "Error. Invalid Clinic NCID:" + sClinicNCID
        ApptsInLastThreeYears = False
        Exit Function
    End If

    If geDBUsed = edb_MySQL Then
        sDateWhere = sMySQLDate
    ElseIf geDBUsed = edb_sqlserver Then
        sDateWhere = sSQLServerDate
    ElseIf geDBUsed = edb_Oracle Then
        sDateWhere = sOracleDate
    Else
        MsgBxARMd "Error. Function ApptsInLastThreeYears not supported in this database mode."
        ApptsInLastThreeYears = False
        Exit Function
    End If
    
    
    sSql = "SELECT DISTINCT APPT_STATUS " + vbCrLf
    sSql = sSql + " From Appointment " + vbCrLf
    sSql = sSql + " Where" + vbCrLf
    sSql = sSql + " PATIENT_UNIT_NUMBER = " + sPatientUnitNumber + " AND" + vbCrLf
    sSql = sSql + " CLINIC_NCID = " + sClinicNCID + " AND" + vbCrLf
    sSql = sSql + " APPT_STATUS = 'KEPT' AND" + vbCrLf
    sSql = sSql + " APPT_CLASSIFICATION <> 3 AND" + vbCrLf
    sSql = sSql + " APPT_DATE_TIME > " + sDateWhere
    
    If LenB(sCurrApptID) <> 0 Then
        sSql = sSql + " AND Appt_ID <> " + sCurrApptID
    End If
    If mobjSQL Is Nothing Then Set mobjSQL = SelectSqlOps()
    mobjSQL.Execute sSql
    If mobjSQL.BOF And mobjSQL.EOF Then
        ApptsInLastThreeYears = False
    Else
        ApptsInLastThreeYears = True
    End If

    Exit Function
    
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "AppointmentsOps.cls - ApptsInLastThreeYears. ", App.Title, vbCritical


End Function

Private Function BuildWhereDateRange() As String

Dim sWhere As String
On Error GoTo ErrHandler
    
    If mobjShared.UseMySQLDB Then
      sWhere = "Appt_Date_Time Between date(sysdate()) and date(date_add(sysdate(),INTERVAL 1 DAY)) "
    ElseIf mobjShared.UseLocalDB Then
      sWhere = " Appt_Date_Time >= TRUNC(SYSDATE) AND Appt_Date_Time < TRUNC(SYSDATE+1) "
    ElseIf mobjShared.UseSQLServer Then
      sWhere = "Appt_Date_Time Between '" & Format$(Now, "mm/dd/yyyy") & "' and '" & Format$(DateAdd("d", 1, Now), "mm/dd/yyyy") & "' "
    End If
    If menumCriteriafilter = apptTodayPlusIncomplete Then
        sWhere = "(" & sWhere & vbCrLf & " OR (Appt_Status = 'KEPT' OR Appt_Status = 'PENDING')  AND " & _
                  "(Encounter_Status not in (" & EncCompleted & ", " & EncUpdated & ", " & EncCancel & ") and Appt_Status = 'KEPT' OR Appt_Status = 'PENDING')  )"
                  'Note: The Encounter_Status of 8 means cancel.
        
    ElseIf mdCriteriaStartDate <> 0 Then
        If mdCriteriaEndDate = 0 Then
            If mobjShared.UseMySQLDB Then
              sWhere = "Appt_Date_Time Between date(sysdate()) And date(date_add(sysdate(),INTERVAL 1 DAY)) "
            ElseIf mobjShared.UseSQLServer Then
              sWhere = "Appt_Date_Time Between '" & Format$(mdCriteriaStartDate, "mm/dd/yyyy") & _
                       "' And '" & Format$(DateAdd("d", 1, mdCriteriaStartDate), "mm/dd/yyyy") & "' "
            Else
              sWhere = " Appt_Date_Time >= TRUNC(" & oApptConn.SQLDate(mdCriteriaStartDate) & ") AND Appt_Date_Time < TRUNC(" & oApptConn.SQLDate(mdCriteriaStartDate) & " +1) "
            End If
        Else
            sWhere = "Appt_Date_Time >= " & oApptConn.SQLDate(mdCriteriaStartDate) & _
                      " and Appt_Date_Time <= " & oApptConn.SQLDate(mdCriteriaEndDate)
        End If
    End If

    BuildWhereDateRange = " and " & sWhere
    Exit Function
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "AppointmentsOps.cls - BuildWhereDateRange. ", App.Title, vbCritical
    
End Function

Private Function BuildWhereClinicians() As String
    Dim sClinician As String
    Dim vVar As Variant
On Error GoTo ErrHandler
    If mcClinicianNCIDs.Count > 0 Then
        For Each vVar In mcClinicianNCIDs
            If LenB(sClinician) = 0 Then
                sClinician = " and a.Clinician_NCID in (" & vVar
            Else
                sClinician = sClinician & ", " & vVar
            End If
        Next vVar
        sClinician = sClinician & ")"
    End If

    BuildWhereClinicians = sClinician
    Exit Function
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "AppointmentsOps.cls - BuildWhereClinicians. ", App.Title, vbCritical
End Function
Private Function BuildWhereIBWAClinicians() As String
    Dim sClinician As String
    Dim vVar As Variant
On Error GoTo ErrHandler
    If mcIBWAClinicianNCID.Count > 0 Then
        For Each vVar In mcIBWAClinicianNCID
            If LenB(sClinician) = 0 Then
                sClinician = " and a.Clinician_NCID in (" & vVar
            Else
                sClinician = sClinician & ", " & vVar
            End If
        Next vVar
        sClinician = sClinician & ")"
    End If

    BuildWhereIBWAClinicians = sClinician
Exit Function
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "AppointmentsOps.cls - BuildWhereIBWAClinicians. ", App.Title, vbCritical

End Function

Private Function BuildWhereClinics() As String
    Dim sClinic As String
    Dim vVar As Variant
On Error GoTo ErrHandler
    'SCR 13738, Phil Crowder 7/23/01  Changed ORing to In
    If mcClinicNCIDs.Count > 0 Then
        sClinic = vbNullString
        For Each vVar In mcClinicNCIDs
            If LenB(sClinic) = 0 Then
                sClinic = "and a.Clinic_NCID in (" & vVar
            Else
                sClinic = sClinic & ", " & vVar
            End If
        Next vVar
        sClinic = sClinic & ")"
    End If

    BuildWhereClinics = sClinic
Exit Function
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "AppointmentsOps.cls - BuildWhereClinics. ", App.Title, vbCritical
End Function
Private Function BuildWhereIBWAClinics() As String
    Dim sClinic As String
    Dim vVar As Variant
On Error GoTo ErrHandler
    'SCR 13738, Phil Crowder 7/23/01  Changed ORing to In
    If mcIBWAClinicNCIDs.Count > 0 Then
        sClinic = vbNullString
        For Each vVar In mcIBWAClinicNCIDs
            If LenB(sClinic) = 0 Then
                sClinic = "and a.Clinic_NCID in (" & vVar
            Else
                sClinic = sClinic & ", " & vVar
            End If
        Next vVar
        sClinic = sClinic & ")"
    End If

    BuildWhereIBWAClinics = sClinic
Exit Function
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "AppointmentsOps.cls - BuildWhereIBWAClinics. ", App.Title, vbCritical
End Function
Private Function BuildWhereStatus() As String
    'KDunne 9/17/2001 scr 16559
    Dim sStatusWhere As String
On Error GoTo ErrHandler
    'KDunne SCR 16559 9/17/2001
    'User's can select multiple Status' to search on now.
    sStatusWhere = Empty
    '--- SCR-18843 MSolano 11/16/01 ``
    If menumCriteriafilter And apptPending Then
      sStatusWhere = sStatusWhere & " or Appt_Status = 'PENDING'"
    End If
    
    If menumCriteriafilter And apptCheckedIn Then
        sStatusWhere = sStatusWhere & " or Encounter_Status = " & EncCheckedIn & _
        " OR ( Encounter_number IS NULL AND encounter_status IS NULL AND APPT_Status = 'KEPT')"
    End If
    If menumCriteriafilter And apptWaiting Then
        sStatusWhere = sStatusWhere & " or Encounter_Status = " & EncWaiting
    End If
    If menumCriteriafilter And apptInProgress Then
        sStatusWhere = sStatusWhere & " or Encounter_Status = " & EncInProgress
    End If
    If menumCriteriafilter And apptNeedsCoSignature Then
        sStatusWhere = sStatusWhere & " or Encounter_Status = " & EncNeedsCosignature
    End If
    If menumCriteriafilter And apptUpdated Then
        sStatusWhere = sStatusWhere & " or Encounter_Status = " & EncUpdated
    End If
    If menumCriteriafilter And apptCompleted Then
        sStatusWhere = sStatusWhere & " or Encounter_Status = " & EncCompleted
    End If
    
    If menumCriteriafilter And apptIncomplete Then
        'SCR 6585 Requirement to add CHSC I pending to incomplete list.
        'Phil Crowder 7/5/01
        sStatusWhere = sStatusWhere & " or (Appt_Status = 'PENDING' OR (Appt_Status = 'KEPT' and not Encounter_Status in (" _
            & EncCompleted & ", " & EncUpdated & ")) OR (APPT_STATUS = 'KEPT' and Encounter_Status is Null) )"
    End If
    
    If LenB(sStatusWhere) <> 0 Then
        sStatusWhere = " And (" & Replace(sStatusWhere, " or ", vbNullString, , 1) & ")"
    End If
    'End SCR 16559
   BuildWhereStatus = sStatusWhere
Exit Function
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "AppointmentsOps.cls - BuildWhereStatus. ", App.Title, vbCritical
    
End Function

Private Function BuildWhereEncounterNumber() As String
    Dim sWhere As String
On Error GoTo ErrHandler
    If LenB(msCriteriaEncounterNumber) <> 0 Then
        sWhere = " and Encounter_Number = '" & msCriteriaEncounterNumber & "'"
    End If
    
    If mlCriteriaUnitNumber <> 0 Then
        sWhere = sWhere & " and Patient_Unit_Number = " & mlCriteriaUnitNumber
    End If

    BuildWhereEncounterNumber = sWhere

Exit Function
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "AppointmentsOps.cls - BuildWhereEncounterNumber. ", App.Title, vbCritical
    
End Function



Public Function DoCancelAppt_CHCS_GUI(sApptIEN As String, lApptClassification As Long, _
                    sApptCancelReason As String, sApptStatus As String, dtCancelDateTime As Date, _
                    sFacilityNCID As String, sPatientIEN As String, ByRef sErrMsg As String) As Boolean
Dim oCHCSInterface_Appt As CHCS_Interface.ICHCSAppointments
Dim objEncounter As CHCSIIEncounterOps.Encounter
Dim eNewEncStatus As ENCOUNTER_INTERFACES.EncStatus
Dim bLockedEncounter As Boolean
Dim bStatus As Boolean

On Error GoTo ErrHandler
    sErrMsg = vbNullString
    
    Set objEncounter = GetEncounter_CHCS_GUI(sApptIEN, sFacilityNCID, sPatientIEN)
    If objEncounter Is Nothing Then GoTo SkipEncounterCancel
    If Not objEncounter.LockingObject.SectionIsLocked(eEntireEncounter, 0) Then
        bLockedEncounter = objEncounter.LockingObject.LockSection(EncounterSectionEnum.eEntireEncounter, 0, LockReasonEnum.SectionLock, "LWOBS")
        If Not bLockedEncounter Then
            MsgBxARMd "Appointment not cancelled." + vbCrLf + _
                        "ApptIEN: " + sApptIEN + "," + vbCrLf + _
                        "ApptStatus: " + sApptStatus + "," + vbCrLf + _
                        "CancelReason: " + sApptCancelReason
            GoTo Cleanup
        End If
    End If
    
    Select Case UCase$(sApptStatus)
        Case "CANCEL"       'should never get to this case
            Select Case UCase$(sApptCancelReason)
                Case "PATIENT CANCELLED"
                    eNewEncStatus = CancelByPatient
                Case "FACILITY CANCELLED"
                    eNewEncStatus = CancelByFacility
                Case "NO-SHOW"
                    eNewEncStatus = NoShow
                
                Case "LWOBS"
                    eNewEncStatus = PatientLeftWithoutBeingSeen
                Case "PROVIDER CANCELLED"
                    eNewEncStatus = CancelByFacility 'CancelByProvider
                Case Else
                    MsgBxARMd "Invalid Cancellation Reason Code in DoCancelAppt_CHCS_GUI. " + vbCrLf + _
                        "ApptIEN: " + sApptIEN + "," + vbCrLf + _
                        "ApptStatus: " + sApptStatus + "," + vbCrLf + _
                        "CancelReason: " + sApptCancelReason + vbCrLf + "Appointment not cancelled."
                    GoTo Cleanup
            End Select
        Case "NO-SHOW"
            eNewEncStatus = NoShow
'        Case "KEPT"  'SCR-31946
'            If sApptCancelReason = "UNDOCANCEL" Then
'                bStatus = objEncounter.UndoCancel
'                GoTo SkipCancel
'            End If
        Case Else
            eNewEncStatus = PatientLeftWithoutBeingSeen
    End Select
    
    bStatus = objEncounter.Cancel(eNewEncStatus)
    If Not bStatus Then
        sErrMsg = "Could not cancel the Encounter document. Cancellation failed for ApptIEN: " + sApptIEN + ", EncID: " + CStr(objEncounter.EncounterID)
        GoTo Cleanup
    End If
SkipEncounterCancel:
    On Error Resume Next
    Set oCHCSInterface_Appt = mobjShared.CHCSConnection
    oCHCSInterface_Appt.AppointmentCancel sApptIEN, lApptClassification, sApptCancelReason, sApptStatus, dtCancelDateTime
    If Err Then
        sErrMsg = "Cancellation failed for ApptIEN: " + sApptIEN + vbCrLf + Err.Description
        DoCancelAppt_CHCS_GUI = False
        Err.Clear
    Else
        DoCancelAppt_CHCS_GUI = True
    End If
    
Cleanup:
    If bLockedEncounter Then
        Call objEncounter.LockingObject.UnlockSection(EncounterSectionEnum.eEntireEncounter, 0)
    End If
    Set oCHCSInterface_Appt = Nothing
    Set objEncounter = Nothing
    Exit Function
ErrHandler:
    Set oCHCSInterface_Appt = Nothing
    mobjShared.ShowVBError Err.Number, Err.Description, "  ApptIEN: " + sApptIEN + " AppointmentsOps.cls - DoCancelAppt_CHCS_GUI.", App.Title, vbCritical
End Function

Public Function DoTransferAppt_CHCS_GUI(sApptIEN As String, sNewProviderIEN As String, _
                    ByRef sErrMsg As String) As Boolean
Dim oCHCSInterface_Appt As CHCS_Interface.ICHCSAppointments
On Error GoTo ErrHandler
    sErrMsg = vbNullString
    DoTransferAppt_CHCS_GUI = True
    Set oCHCSInterface_Appt = mobjShared.CHCSConnection
    On Error Resume Next
    oCHCSInterface_Appt.AppointmentProviderTransfer sApptIEN, sNewProviderIEN
    If Err Then
        If InStr(1, Err.Description, "Provider is not assigned to this clinic or is inactive", vbTextCompare) > 0 Then
            Err.Clear
            sErrMsg = " Provider is not assigned to this clinic or is inactive."
        Else
            sErrMsg = Err.Description
            DoTransferAppt_CHCS_GUI = False
        End If
    End If
    
    
    Set oCHCSInterface_Appt = Nothing
    Exit Function
ErrHandler:
    Set oCHCSInterface_Appt = Nothing
    mobjShared.ShowVBError Err.Number, Err.Description, "AppointmentsOps.cls - DoTransferAppt_CHCS_GUI.", App.Title, vbCritical
End Function

Public Function RetrievePatientUnitNumber_CHCS_GUI(sPatientIEN As String, Optional sApptIEN As String) As String
'Function created for SCR-39491 but as of 8/5/03 it is not used by appointments.

        Dim oCHCSInterface_Patient As CHCS_Interface.ICHCSPatient
        Dim rsPatient As ADODB.Recordset
        On Error GoTo ErrHandler

        If LenB(sPatientIEN) = 0 Or sPatientIEN = "0" Then
            RetrievePatientUnitNumber_CHCS_GUI = vbNullString
            GoTo Cleanup
        End If
        Set oCHCSInterface_Patient = mobjShared.CHCSConnection
        Set rsPatient = oCHCSInterface_Patient.RetrieveDemographics(sPatientIEN)
        If Not rsPatient Is Nothing Then
            If rsPatient.EOF = False Then
                RetrievePatientUnitNumber_CHCS_GUI = rsPatient("Patient_UnitNumber") & vbNullString
                
                If IsNumeric(RetrievePatientUnitNumber_CHCS_GUI) = False Then
                    RetrievePatientUnitNumber_CHCS_GUI = vbNullString
                Else
                    'Update this (and any other appt for the same patient)
                    'appt obj in the collection of appts if the apptIEN is provided.
                    Dim sPatName As String
                    Dim sPatIEN As String
                    Dim i As Integer
                    
                    If IsNumeric(sApptIEN) = True Then
                        mcolAppointmentsCache.Item(sApptIEN).PatientUnitNumber = RetrievePatientUnitNumber_CHCS_GUI
                        sPatName = mcolAppointmentsCache.Item(sApptIEN).PatientName
                        sPatIEN = mcolAppointmentsCache.Item(sApptIEN).PatientIEN
                        If LenB(sPatIEN) = 0 Or LenB(sPatName) = 0 Then
                            GoTo Cleanup
                        End If
                        Debug.Print "PatientUnitNumber: " + RetrievePatientUnitNumber_CHCS_GUI + " updated for ApptIEN: " + sApptIEN
                        For i = 1 To mcolAppointmentsCache.Count
                            If mcolAppointmentsCache.Item(i).PatientName = sPatName And _
                                mcolAppointmentsCache.Item(i).PatientIEN = sPatIEN Then
                                mcolAppointmentsCache.Item(i).PatientUnitNumber = RetrievePatientUnitNumber_CHCS_GUI
                            End If
                        Next
                    End If
                End If
            Else
                RetrievePatientUnitNumber_CHCS_GUI = vbNullString
            End If
        Else
            RetrievePatientUnitNumber_CHCS_GUI = vbNullString
        End If

Cleanup:
    Set oCHCSInterface_Patient = Nothing
    Set rsPatient = Nothing

Exit Function
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "AppointmentsOps.cls - RetrievePatientUnitNumber_CHCS_GUI.", App.Title, vbCritical
    GoTo Cleanup
        
End Function

Public Function RetrieveInpatientStatus_CHCS_GUI(sPatientIEN As String, ByRef bIsCurrentInpatient As Boolean, ByRef sInpatientLabel As String, ByRef sInpatient_MEPRS_Code As String, _
                ByRef sInpatient_MEPRS_IEN As String, ByRef sInpatient_MEPRS_Description As String) As Boolean
'sPatientIEN is only input arg, rest are return values.
        Dim oCHCSInterface_Patient As CHCS_Interface.ICHCSPatient
        On Error GoTo ErrHandler

        If LenB(sPatientIEN) = 0 Or sPatientIEN = "0" Then
            RetrieveInpatientStatus_CHCS_GUI = True
            GoTo Cleanup
        End If
        Set oCHCSInterface_Patient = mobjShared.CHCSConnection
        sInpatient_MEPRS_Code = vbNullString
        oCHCSInterface_Patient.RetrieveInpatientStatus sPatientIEN, sInpatient_MEPRS_IEN, _
            sInpatient_MEPRS_Code, sInpatient_MEPRS_Description, True
        If LenB(sInpatient_MEPRS_IEN) <> 0 Then
            bIsCurrentInpatient = True
            sInpatientLabel = "Patient found as InPatient"
            sInpatientLabel = sInpatientLabel + "(MEPRS Code:" + sInpatient_MEPRS_Code + ")"
        Else
            bIsCurrentInpatient = False
        End If
        
        RetrieveInpatientStatus_CHCS_GUI = True
Cleanup:
    Set oCHCSInterface_Patient = Nothing
    Exit Function
ErrHandler:
    RetrieveInpatientStatus_CHCS_GUI = False
    mobjShared.ShowVBError Err.Number, Err.Description, "AppointmentsOps.cls - RetrieveInpatientStatus_CHCS_GUI.", App.Title, vbCritical
    GoTo Cleanup
End Function

Public Function DoSearch_CHCS_GUI(Optional bRefresh As Boolean = False, Optional sProviderIEN As String, _
                    Optional sClinicIEN As String, Optional dSearchDate As Date, Optional sPatientIEN As String, _
                    Optional bIncludeTelcons As Boolean, Optional sApptIEN As String, _
                    Optional sIBWAClinicIEN As String, Optional sIBWAProviderIEN As String) As Collection

Dim cAppointments As Collection
Dim rs As ADODB.Recordset
Dim rsIBWA As ADODB.Recordset
Dim dStartDate As Date
Dim dEndDate As Date
Dim oCHCSInterface_Appt As CHCS_Interface.ICHCSAppointments

    


        On Error GoTo ErrHandler

        If Not mobjShared.IsAppMode(modeCHCSI_GUI) Then GoTo Cleanup

        If Not mbTelcons_Enabled_CHCS_GUI Then bIncludeTelcons = False 'SCR-39944
        If sProviderIEN = "0" Then
            MsgBxARMd "Error. CHCS appointment search could not be completed. The " + _
                      "selected provider's IEN was 0 (zero)."
            GoTo Cleanup
        End If
        '**********************************************************************************
        If LenB(sPatientIEN) > 0 Then 'CALL ONLY FOR HEALTH HISTORY
            dStartDate = DateAdd("m", -15, Now)
            dEndDate = DateAdd("d", 1, CDate(Format$(Now, "mm/dd/yyyy")))
            Set oCHCSInterface_Appt = mobjShared.CHCSConnection
            Set rs = oCHCSInterface_Appt.RetrieveAppointments(, , , sPatientIEN, , dStartDate, dEndDate, False, mbTelcons_Enabled_CHCS_GUI)  'SCR-39944
            Set oCHCSInterface_Appt = Nothing
            Set cAppointments = New Collection
            LoadAppointments_CHCS_GUI cAppointments, rs
        
        Else
            'Get appts from CHCS
            If Not DoSearch_CHCS_GUI_1CallCHCS(rs, rsIBWA, bRefresh, sProviderIEN, sClinicIEN, _
                                              dSearchDate, bIncludeTelcons, sApptIEN, sIBWAClinicIEN, sIBWAProviderIEN) Then GoTo Cleanup
            
            'This will combine both recordsets into one collection of appointments.
            Set cAppointments = DoSearch_CHCS_GUI_2ProcessRS(rs, rsIBWA)
            
            'This will load cAppointments into our global collection of
            'mcolAppointmentsCache cache appts. cAppointments will be displayed on the grid.
            If Not DoSearch_CHCS_GUI_3LoadCache(sApptIEN, cAppointments) Then GoTo Cleanup

        End If
        Set DoSearch_CHCS_GUI = cAppointments
        

        'For i = 1 To mcolAppointmentsCache.Count
        ' debug.Print mcolAppointmentsCache.Item(i).
        'Next
Cleanup:
        Set rs = Nothing
        Set rsIBWA = Nothing
        Set cAppointments = Nothing
        If DoSearch_CHCS_GUI Is Nothing Then Set DoSearch_CHCS_GUI = New Collection
        Exit Function
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "AppointmentsOps.cls - DoSearch_CHCS_GUI", App.Title, vbCritical
    Set oCHCSInterface_Appt = Nothing
    GoTo Cleanup
End Function
Private Function DoSearch_CHCS_GUI_1CallCHCS(rs As ADODB.Recordset, rsIBWA As ADODB.Recordset, _
                    Optional bRefresh As Boolean = False, Optional sProviderIEN As String, _
                    Optional sClinicIEN As String, Optional dSearchDate As Date, _
                    Optional bIncludeTelcons As Boolean, Optional sApptIEN As String, _
                    Optional sIBWAClinicIEN As String, Optional sIBWAProviderIEN As String) As Boolean
                    'rs and rsIBWA are return values
'DataLayer Item
                    
            'must have value in sIBWAProviderIEN or don't call
            'sIBWAProviderIEN will only have one value
            'sIBWAClinicIEN will only have one value or be blank
            
            'Must have a value in one or the other, both can not be blank
            'sProviderIEN - can be empty or one single value
            'sClinicIEN - can be empty or one single value
                    
                    
'Dim rs As ADODB.Recordset
'Dim rsIBWA As ADODB.Recordset
'SCR-38679 Provier list can now be comma delimited string.
Dim oCHCSInterface_Appt As CHCS_Interface.ICHCSAppointments
Dim eRNDSOpt As eIncludeRounds
On Error GoTo ErrHandler
    Set oCHCSInterface_Appt = mobjShared.CHCSConnection
             
             
    If LenB(sApptIEN) <> 0 Then 'SINGLE APPT QUERY FROM GetAppointment_CHCS_GUI
        Set rs = oCHCSInterface_Appt.RetrieveAppointments(, , , , sApptIEN, , , bRefresh, mbTelcons_Enabled_CHCS_GUI, , Include) 'SCR-39944
    
    Else  'QUERY FOR APPTS GRID
        'NON-RNDS
        Debug.Print "#############################################################"
        Debug.Print "Start CHCS Query: " + CStr(Now)
        Debug.Print "sClinicIEN= " + sClinicIEN + ", sProviderIEN= " + sProviderIEN + ", ForceRefresh= " + CStr(bRefresh)
        If sClinicIEN <> vbNullString Or sProviderIEN <> vbNullString Then
            eRNDSOpt = Exclude
            Set rs = oCHCSInterface_Appt.RetrieveAppointments(sProviderIEN, sClinicIEN, CDate(Format$(dSearchDate, "mm/dd/yyyy")), , , , , bRefresh, bIncludeTelcons, , Exclude)
        End If
        'Debug.Print "RetAppts: sProvIEN:" + sProviderIEN + ", sClinIEN:" + sClinicIEN + ", Date:" + Format$(dSearchDate, "mm/dd/yyyy") + ", FRefresh:True  , IncludeTelcons:" + CStr(bIncludeTelcons)
        'Wide Open search:
        'Set rs = oCHCSInterface_Appt.RetrieveAppointments(, , , , , CDate("1/9/2004"), CDate("1/16/2004"), True, True)
        Debug.Print CStr(Now)
        'RNDS appts
        If mbRNDS_On_and_NonTheater_mode And sIBWAProviderIEN <> vbNullString Then
            eRNDSOpt = IncludeOnly
            Debug.Print "sIBWAClinicIEN= " + sIBWAClinicIEN + ", sIBWAProviderIEN= " + sIBWAProviderIEN + ", ForceRefresh= " + CStr(bRefresh)
            Set rsIBWA = oCHCSInterface_Appt.RetrieveAppointments(sIBWAProviderIEN, sIBWAClinicIEN, _
                CDate(Format$(dSearchDate, "mm/dd/yyyy")), , , , , bRefresh, False, , eRNDSOpt)
            
        End If
        Debug.Print "END CHCS Query: " + CStr(Now)
        Debug.Print "#############################################################"
    End If
    If Not rs Is Nothing Then
        Debug.Print "Appt Records Returned:" + CStr(rs.RecordCount)
    End If
    If Not rsIBWA Is Nothing Then
        Debug.Print "IBWA Appt Records Returned:" + CStr(rsIBWA.RecordCount)
    End If
    DoSearch_CHCS_GUI_1CallCHCS = True
Cleanup:
    Set oCHCSInterface_Appt = Nothing

    Exit Function
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "AppointmentsOps.cls - DoSearch_CHCS_GUI_1CallCHCS", App.Title, vbCritical
    GoTo Cleanup

End Function
Private Function DoSearch_CHCS_GUI_2ProcessRS(rs As ADODB.Recordset, rsIBWA As ADODB.Recordset) As Collection     'of Appts
Dim cAppts As Collection
On Error GoTo ErrHandler
   
     Set cAppts = New Collection
    If rs Is Nothing Then
        Set mcolAppointmentsCache = New Collection  'SCR-32645
    ElseIf rs.BOF And rs.EOF Then
        Set mcolAppointmentsCache = New Collection  'SCR-32645
    Else
        LoadAppointments_CHCS_GUI cAppts, rs
    End If
    
    If mbRNDS_On_and_NonTheater_mode Then
        If Not rsIBWA Is Nothing Then
            If Not rsIBWA.EOF And Not rsIBWA.BOF Then
                LoadAppointments_CHCS_GUI cAppts, rsIBWA
            End If
        End If
    
    End If
    Set DoSearch_CHCS_GUI_2ProcessRS = cAppts
Cleanup:
    Set cAppts = Nothing
Exit Function
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "AppointmentsOps.cls - DoSearch_CHCS_GUI_2ProcessRS", App.Title, vbCritical
    GoTo Cleanup
                    
End Function

Private Function DoSearch_CHCS_GUI_3LoadCache(sApptIEN As String, cAppointments As Collection) As Boolean
                                             
Dim oAppt As CHCSII_AppointmentClient.Appointment
Dim i As Integer
                    
On Error GoTo ErrHandler
        
        'If we have an ApptIEN this means that we are making a query for a single
        'appt which was made when creating a new appt. Here we want to simply add to
        'the collection and not over write.
        'mcolAppointmentsCache is our cache of appts that is searched when
        'we (or a calling app) does a GetAppointment_CHCS_GUI
        If LenB(sApptIEN) <> 0 Then
                If cAppointments Is Nothing Then
                    Set mcolAppointmentsCache = New Collection
                End If
                If cAppointments.Count = 1 Then
                    Set oAppt = cAppointments.Item(1)
                    On Error Resume Next
                    If Not mcolAppointmentsCache Is Nothing Then
                        mcolAppointmentsCache.Remove oAppt.ApptIEN
                        mcolAppointmentsCache.Add oAppt, oAppt.ApptIEN
                        
                    Else
                        Set mcolAppointmentsCache = New Collection
                        mcolAppointmentsCache.Add oAppt, oAppt.ApptIEN
                    End If
                    If Err Then Err.Clear
                End If
        Else
            If cAppointments Is Nothing Then Set cAppointments = New Collection
            Set mcolAppointmentsCache = New Collection
            For i = 1 To cAppointments.Count
                mcolAppointmentsCache.Add cAppointments.Item(i), cAppointments.Item(i).ApptIEN
            Next
            Debug.Print "Appt Cache reset with " + CStr(cAppointments.Count) + " new appts."
        End If
   
                 
   DoSearch_CHCS_GUI_3LoadCache = True
Cleanup:
   'Set cAppointments = Nothing
Exit Function
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "AppointmentsOps.cls - DoSearch_CHCS_GUI_3LoadCache", App.Title, vbCritical
    GoTo Cleanup
                    
End Function



Public Function CheckForInpatientStatus(objSQL As Object) As Boolean  'ICHCSII_SQL.ISqlOps
Dim sql As String

'Need to type as object here so that the calling function will have a
'generic recordset to parse. mobjSQL is not the same type as returned
'by SP method.
On Error GoTo ErrHandler
    'PERF STORED PROC
    If mobjShared.IsAppMode(modeCDR) And Not mobjShared.IsAppMode(modeITT) Then
        Set objSQL = CheckForInpatientStatus_SP(mobjPatient.UnitNumber, CLng(mobjLogon.FacilityNCID))
    Else
        sql = "SELECT DISTINCT INPATIENT_MEPRS_CODE,INPATIENT_MEPRS_IEN, INPATIENT_MEPRS_DESCRIPTION  "
        sql = sql & "  FROM INPATIENT_MEPRS_INFO_V WHERE PATIENT_UNIT_NUMBER = " + CStr(mobjPatient.UnitNumber) + " AND FACILITY_NCID = " + CStr(mobjLogon.FacilityNCID)
        If mobjSQL Is Nothing Then Set mobjSQL = SelectSqlOps()
        mobjSQL.Execute sql
        Set objSQL = mobjSQL
        'Set CheckForInpatientStatus = mobjSQL
    End If
    CheckForInpatientStatus = True
Exit Function
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "AppointmentsOps.cls - CheckForInpatientStatus", App.Title, vbCritical

End Function
Public Function DoSearch(Optional lRecordLimit As Long, Optional bRefresh As Boolean = False, Optional sPatientIEN As String) As Collection  ' of Appointment
    
    Dim sWhere As String
    Dim sClinic As String
    Dim sClinician As String
    Dim sIBWAClinic As String 'SCR-RNDS
    Dim sIBWAClinician As String 'SCR-RNDS
    Dim i As Integer
    
    On Error GoTo ErrHandler
    
    If Not mobjShared.IsAppMode(modeCHCSI_GUI) Then
        sWhere = sWhere & BuildWhereDateRange
        sWhere = sWhere & BuildWhereEncounterNumber
        sWhere = sWhere & BuildWhereStatus
        sClinic = BuildWhereClinics
        sClinician = BuildWhereClinicians
    
        If mbRNDS_On_and_NonTheater_mode Then 'SCR-RNDS
            sIBWAClinic = BuildWhereIBWAClinics
            sIBWAClinician = BuildWhereIBWAClinicians
        End If
    End If
    
    '#######################################################################################
    '#######################################################################################
    'To query appointent records we have basically four modes to deal with here.
    
    If mobjShared.IsAppMode(modeCDR) And Not mobjShared.IsAppMode(modeITT) Then
        '1.) modeCDR - Uses Oracle data source; Uses latest Appts Package/Stored Procedures
        '              referenced by Public Const CURR_SP_VER.
        Set DoSearch = DoSearch_CDR(lRecordLimit, sWhere, sClinic, sClinician, sIBWAClinic, sIBWAClinician)
    
    ElseIf mobjShared.IsAppMode(modeITT) Then
        '2.)  modeITT - Uses SQL Server data source; No stored procedures. Main SQL script built when
        '               AppointmentOps is initialized: gsSelectSQLServer.
        '               modeITT uses both Appts & Telcons module.
        Set DoSearch = DoSearch_ITT(lRecordLimit, sWhere, sClinic, sClinician, sIBWAClinic, sIBWAClinician)
    
    ElseIf mobjShared.IsAppMode(modeTheater) Then
        '3.) modeTheater - Uses Oracle data source; No stored procedures. Main SQL script built
        '                  as Const sSelect + sWhereOracle + sQueryWhere2Oracle.
        '                  modeTheater uses Appts module only.
        Set DoSearch = DoSearch_GEMS(lRecordLimit, sWhere, sClinic, sClinician)
    
    ElseIf mobjShared.IsAppMode(modeCHCSI_GUI) Then
        Set DoSearch = DoSearch_CHCS_GUI(bRefresh, , , Now, sPatientIEN)
        '4.) modeCHCS_GUI - Uses VB interface to CHCS1. The calling module (frmAppointments)
        '                   makes calls directly to DoSearch_CHCS_GUI, but condition left here to
        '                   cover all modes.
    Else
        MsgBxARMd "The CHCSII Appointments search engine does not support this mode. " + _
                  "Unable to complete query selection."
                  Exit Function
    End If
    
    Set mcolAppointmentsCache = New Collection  'SCR-32352
    On Error Resume Next
    For i = 1 To DoSearch.Count
        mcolAppointmentsCache.Add DoSearch.Item(i), DoSearch.Item(i).ApptId
    Next


Cleanup:
    Exit Function
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "AppointmentsOps.cls - DoSearch", App.Title, vbCritical
End Function
Private Function DoSearch_CDR(lRecordLimit As Long, sWhere As String, sClinic As String, _
                              sClinician As String, sIBWAClinic As String, sIBWAClinician As String) As Collection
                              
    Dim rs As ADODB.Recordset

    On Error GoTo ErrHandler

    Set DoSearch_CDR = New Collection
    '***************************************************************************************
    If sClinic <> vbNullString Or sClinician <> vbNullString Then 'if only doing a RNDS search, sClinic will be empty.
        Set rs = New ADODB.Recordset
        Set rs = DoSearch_Oracle_SP(APPT_QUERY, sClinic, sClinician, sWhere, lRecordLimit, mobjLogon.FacilityNCID)
        If Not rs Is Nothing Then
            If Not rs.BOF And Not rs.EOF Then
                LoadAppointments_SP DoSearch_CDR, rs
                'SCR-38720 'GetDetailCodesforApptQuery
                
                'GetDetailCodesforApptQuery updates rs with the correct detail codes.
                'GetDetailCodesforApptQuery DoSearch_CDR, sClinic, sClinician, sWhere, lRecordLimit, mobjLogon.FacilityNCID, "0, 1, 2"
            End If
        End If
    Else
        sSQLStatement = vbNullString
    End If
    '***************************************************************************************
    If Not mbRNDS_On_and_NonTheater_mode Then GoTo Cleanup 'SCR-RNDS
    If sIBWAClinic <> vbNullString Or sIBWAClinician <> vbNullString Then
        Set rs = New ADODB.Recordset
        If msIBWAClinicianFacilityNCID = vbNullString Then
            msIBWAClinicianFacilityNCID = mobjLogon.FacilityNCID
        End If
        Set rs = DoSearch_Oracle_SP(APPT_QUERY, sIBWAClinic, sIBWAClinician, sWhere, lRecordLimit, CLng(msIBWAClinicianFacilityNCID))
        If rs Is Nothing Then GoTo Cleanup
            
        If Not rs.BOF And Not rs.EOF Then
            LoadAppointments_SP DoSearch_CDR, rs
            'SCR-38720
            'GetDetailCodesforApptQuery updates rs with the correct detail codes.
            'GetDetailCodesforApptQuery DoSearch_CDR, sClinic, sClinician, sWhere, lRecordLimit, mobjLogon.FacilityNCID, "0, 1, 2"
        End If
    End If
    
                            
Exit Function
Cleanup:
    Set rs = Nothing
    Exit Function
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "AppointmentsOps.cls - DoSearch_CDR", App.Title, vbCritical
    GoTo Cleanup
                              
End Function
Private Function DoSearch_ITT(lRecordLimit As Long, sWhere As String, sClinic As String, _
                              sClinician As String, sIBWAClinic As String, sIBWAClinician As String) As Collection
Dim sRecordLimit As String

On Error GoTo ErrHandler

    If lRecordLimit <> 0 Then
      sRecordLimit = "TOP " + CStr(lRecordLimit) + " "
    End If
    
    '***************************************************************************************
    If sClinic <> vbNullString Then 'if only doing a RNDS search, sClinic will be empty.
        sSQLStatement = "Select " + sRecordLimit + gsSelectSQLServer
        sSQLStatement = sSQLStatement & sClinic & sClinician & sWhere & sQueryWhere3 & vbCrLf _
            & " and a.facility_ncid = " & mobjLogon.FacilityNCID & vbCrLf _
            & " and mc.FACILITY_NCID = " & mobjLogon.FacilityNCID & vbCrLf
        sSQLStatement = sSQLStatement & " order by appt_date_time, patient_name, patient_ssn"
        Set DoSearch_ITT = New Collection
        If mobjSQL Is Nothing Then Set mobjSQL = SelectSqlOps()
        mobjSQL.Execute sSQLStatement
        If Not mobjSQL.BOF And Not mobjSQL.EOF Then
            LoadAppointments DoSearch_ITT
        End If
    Else
        sSQLStatement = vbNullString
    End If
    '***************************************************************************************
    If Not mbRNDS_On_and_NonTheater_mode Then GoTo Cleanup 'SCR-RNDS
    If sIBWAClinic <> vbNullString Or sIBWAClinician <> vbNullString Then
    
        sSQLStatement = "Select " + sRecordLimit + gsSelectSQLServer
        sSQLStatement = sSQLStatement & sIBWAClinic & sIBWAClinician & sQueryWhere3 & sWhere & sQueryWhere3 & vbCrLf _
            & " and a.facility_ncid = " & mobjLogon.FacilityNCID & vbCrLf _
            & " and mc.FACILITY_NCID = " & mobjLogon.FacilityNCID & vbCrLf
        sSQLStatement = sSQLStatement & " order by appt_date_time, patient_name, patient_ssn"
        If mobjSQL Is Nothing Then Set mobjSQL = SelectSqlOps()
        mobjSQL.Execute sSQLStatement
        If mobjSQL.BOF And mobjSQL.EOF Then
            GoTo Cleanup
        End If
        LoadAppointments DoSearch_ITT
    End If
    
    Exit Function
Cleanup:
    Exit Function
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "AppointmentsOps.cls - DoSearch_ITT", App.Title, vbCritical
    GoTo Cleanup
                              
End Function
Private Function DoSearch_GEMS(lRecordLimit As Long, sWhere As String, sClinic As String, _
                              sClinician As String) As Collection
                              
 On Error GoTo ErrHandler
    If LenB(mobjShared.CmdLineSwitches.Item("INPATIENT2")) = 0 Then
        If mobjShared.UseMySQLDB Then
            sSQLStatement = sMySQLSelect & vbNewLine & sWhereMySQL
        Else
            sSQLStatement = sOracleSelect & vbNewLine & sWhereOracle
        End If
    Else
        If mobjShared.UseMySQLDB Then
            sSQLStatement = sMySQLSelect & ", GETAPPTDISPOSITION(a.Appt_ID) AS DISPOSITION" & vbNewLine & sWhereMySQL
        Else
            sSQLStatement = sOracleSelect & ", GETAPPTDISPOSITION(a.Appt_ID) AS DISPOSITION" & vbNewLine & sWhereOracle
        End If
    End If
    If Not mbGEMSUserHasPriv_Data_VIP Then 'SCR-42772
        sSQLStatement = sSQLStatement _
                    & "   and not exists (select gd.unit_number from MMI_Generic_id gd " & vbCrLf _
                    & "         where gd.id_type_ncid=14501719 and " & vbCrLf _
                    & "         gd.id_value='Y' and " & vbCrLf _
                    & "         gd.unit_number = id.unit_number)"
                    
    End If
    If mobjShared.UseMySQLDB Then
        sSQLStatement = sSQLStatement & sClinic & sClinician & sQueryWhere2MySQL & sWhere & _
                        sQueryWhere3 & vbCrLf & " and a.facility_ncid = " & _
                        mobjLogon.FacilityNCID & vbCrLf
    Else
        sSQLStatement = sSQLStatement & sClinic & sClinician & sQueryWhere2Oracle & sWhere & _
                        sQueryWhere3 & vbCrLf & " and a.facility_ncid = " & _
                        mobjLogon.FacilityNCID & vbCrLf
    End If
    
    'SCR-18127 S.McAvoy 7/22/02 Added record limit option.
    If lRecordLimit <> 0 Then
        If mobjShared.UseMySQLDB Then
            sSQLStatement = sSQLStatement & " LIMIT " & CStr(lRecordLimit + 1) & " "
        Else
            sSQLStatement = sSQLStatement & " AND ROWNUM < " & CStr(lRecordLimit + 1) & " "
        End If
    End If
    
    sSQLStatement = sSQLStatement & " order by appt_date_time, patient_name, patient_ssn"
                              
    Set DoSearch_GEMS = New Collection
    If mobjSQL Is Nothing Then Set mobjSQL = SelectSqlOps()
    mobjSQL.Execute sSQLStatement
    
    If mobjSQL.BOF And mobjSQL.EOF Then
        GoTo Cleanup
    End If
    LoadAppointments DoSearch_GEMS
                            
Exit Function
Cleanup:
    Exit Function
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "AppointmentsOps.cls - DoSearch_GEMS", App.Title, vbCritical
    GoTo Cleanup
                              
End Function



'SCR-27715 S.McAvoy 9/24/02
Public Sub CheckForConsultOrder(Appt_Id As String, Provider_NCID As String, OldProvider_Name As String, NewProvider_Name As String)
Dim sSql As String
Dim cConsult_IDs As Collection
Dim i As Integer
Dim lconsult_trans_id As Long
Dim curDate As String
Dim sTemp As String

On Error GoTo ErrHandler
    If mobjShared.IsAppMode(modeTheater) Then Exit Sub
    If mobjSQL Is Nothing Then Set mobjSQL = SelectSqlOps()
    sSql = "Select CONSULT_ID, RECV_PROVIDER_NCID from CONSULT_ORDER where APPT_ID = " + Appt_Id
    mobjSQL.Execute sSql
    If mobjSQL.BOF And mobjSQL.EOF Then
        'No consults found
        Exit Sub
    End If
    Set cConsult_IDs = New Collection
    mobjSQL.MoveFirst
    
    'Should be only one consult found, but add them to collection...
    Do While Not mobjSQL.EOF
        If CStr(mobjSQL.Value("RECV_PROVIDER_NCID")) + vbNullString <> Provider_NCID Then
            cConsult_IDs.Add mobjSQL.Value("CONSULT_ID") & vbNullString
        End If
        mobjSQL.MoveNext
    Loop
    
    'For each consult found, update its provider in the consult_order table.
    For i = 1 To cConsult_IDs.Count
        
        sSql = "Update CONSULT_ORDER Set RECV_PROVIDER_NCID = " + Provider_NCID + _
               " Where CONSULT_ID = " + cConsult_IDs.Item(i)
        If mobjSQL Is Nothing Then Set mobjSQL = SelectSqlOps()
        mobjSQL.Execute (sSql)
        'Also for each consult record changed update the consult log table...
        mobjSQL.Execute "select max(consult_trans_id)as CONSULT_TRANS_ID from consult_tracking_log where consult_id = " & cConsult_IDs.Item(i)
        If mobjSQL.BOF And mobjSQL.EOF Then 'No entries in the log found.
            lconsult_trans_id = 1
        Else
            mobjSQL.MoveFirst
            sTemp = CStr(mobjSQL.Value("CONSULT_TRANS_ID") & vbNullString)
            If LenB(sTemp) = 0 Then
                lconsult_trans_id = 0
            Else
                lconsult_trans_id = CLng(sTemp)
            End If
            lconsult_trans_id = lconsult_trans_id + 1
        End If
        
        sSql = "Insert into CONSULT_TRACKING_LOG Values( " + cConsult_IDs.Item(i) + ", " + _
            CStr(lconsult_trans_id) + ", " & oApptConn.SQLDate(curDate) & ", " + _
            "'Provider changed FROM" + OldProvider_Name + " TO " + NewProvider_Name + " via Appointments.', " + mobjShared.CurrentUserNCID + ", NULL)"
        
        mobjSQL.Execute sSql
    
    
    Next
    Set cConsult_IDs = Nothing
    Exit Sub
ErrHandler:
    Set cConsult_IDs = Nothing
    mobjShared.ShowVBError Err.Number, Err.Description, "AppointmentsOps.cls - CheckForConsultOrder. ", App.Title, vbCritical
    
End Sub

Private Sub LoadAppointments(ByRef Appointments As Collection)
    Dim lNumber As Long
    Dim sSource As String
    Dim sDescription As String
    Dim objAppt As CHCSII_AppointmentClient.Appointment

    On Error Resume Next
    If mobjSQL Is Nothing Then Exit Sub
    If Appointments Is Nothing Then Set Appointments = New Collection 'SCR-53141
    mobjSQL.MoveFirst
    
    Do While Not mobjSQL.EOF
        Set objAppt = New CHCSII_AppointmentClient.Appointment
        
        objAppt.Deserialize mobjSQL  '
        
        'DEM - Added error handling to ignore duplicate key error messages
        Appointments.Add objAppt, objAppt.ApptId
        
        If Err <> 0 And Err <> 457 Then
            'Raise error
            lNumber = Err.Number
            sSource = Err.Source
            sDescription = Err.Description
            On Error GoTo 0
            Err.Raise lNumber, sSource, sDescription
        End If
        
        Set objAppt = Nothing
        mobjSQL.MoveNext
    Loop

End Sub
Private Sub LoadAppointments_SP(ByRef Appointments As Collection, ByRef rs As Recordset)
    Dim lNumber As Long
    Dim sSource As String
    Dim sDescription As String
    Dim objAppt As CHCSII_AppointmentClient.Appointment

    On Error Resume Next
    If rs.EOF = True Then Exit Sub
    
    rs.MoveFirst
    
    Do While Not rs.EOF
        Set objAppt = New CHCSII_AppointmentClient.Appointment
        
        objAppt.DeserializeAppts_SP rs
        
        'DEM - Added error handling to ignore duplicate key error messages
        Appointments.Add objAppt, objAppt.ApptId
        
        If Err <> 0 And Err <> 457 Then
            'Raise error
            lNumber = Err.Number
            sSource = Err.Source
            sDescription = Err.Description
            On Error GoTo 0
            Err.Raise lNumber, sSource, sDescription
        End If
        
        Set objAppt = Nothing
        rs.MoveNext
    Loop

End Sub
Private Sub LoadAppointments_CHCS_GUI(ByRef Appointments As Collection, ByRef rs As Recordset)
    Dim lNumber As Long
    Dim sSource As String
    Dim sDescription As String
    Dim objAppt As CHCSII_AppointmentClient.Appointment
    
    On Error Resume Next
'/* SCR 29616 - the recordset returned from CHCS_GUI is always at the end of file, need to check BOF to determin if it's empty
'    If rs.EOF = True Then Exit Sub
    If rs Is Nothing Then Exit Sub
    If rs.BOF And rs.EOF = True Then Exit Sub
'*/ end SCR 29616

    rs.MoveFirst
    
    Do While Not rs.EOF
        Set objAppt = New CHCSII_AppointmentClient.Appointment
        objAppt.CurrentCHCS_GUI_Provider = mobjUser.UserName
        objAppt.Deserialize_CHCS_GUI rs, mobjLogon.FacilityNCID
        
        'DEM - Added error handling to ignore duplicate key error messages
            
        Appointments.Add objAppt, objAppt.ApptIEN 'in PGUI mode apptID & apptIEN are the same.
        'Debug.Print CStr(Err.Number)
        If Err <> 0 And Err <> 457 Then
            'Raise error
            lNumber = Err.Number
            sSource = Err.Source
            sDescription = Err.Description
            On Error GoTo 0
            Err.Raise lNumber, sSource, sDescription
        End If
        
        Set objAppt = Nothing
        rs.MoveNext
    Loop

End Sub
Public Function GetWorkLoad_Options(sApptID As String, Optional sClinicNCID As String) As String
'If the above returns NC then enable Count AND Non-Count
'If the above returns C then enable ONLY Count… and so forth.

Dim oAppt As CHCSII_AppointmentClient.Appointment
Dim sErrMsg As String
Dim sClinicianNCID As String
Dim sClinNCID As String

On Error GoTo ErrHandler
    If LenB(sApptID) = 0 Then
        MsgBxARMd "Could not obtain valid Workload Options. No ApptID sent when calling GetWorkLoad_Options.", , "Error"
        GetWorkLoad_Options = vbNullString
        GoTo Cleanup
    End If
    Set oAppt = GetAppointment(sApptID)
    If oAppt Is Nothing Then
        Set oAppt = GetAppointment(sApptID, False)
        If oAppt Is Nothing Then
            MsgBxARMd "Could not obtain valid Workload Options. Could not retrieve a valid Appointment object for the " + _
                "given ApptID: " + sApptID, , "Error"
            GetWorkLoad_Options = vbNullString
            GoTo Cleanup
        End If
    End If
    If LenB(sClinicNCID) <> 0 Then
        sClinNCID = sClinicNCID
    Else
        If mobjShared.IsAppMode(modeCHCSI_GUI) Then
            sClinNCID = mobjLogon.NursingDivisionNCID
        Else
            sClinNCID = oAppt.ClinicNCID
        End If
    End If
    If LenB(sClinNCID) = 0 Then
        MsgBxARMd "Could not obtain valid Workload Options. Could not retrieve a valid ClinicNCID for the " + _
            "given ApptID: " + sApptID, , "Error"
        GetWorkLoad_Options = vbNullString
        GoTo Cleanup
    End If
    
    'SCR-38036>>> When calling GetClinicWorkloadType, we need to use the apptType Code not apptType Description.
    'oAppt.Appt_Type contains the appt_type description as displayed in the grid.
    If gobjClinicOps Is Nothing Then
        Set gobjClinicOps = New CHCSII_ClinicClient.ClinicOpsEx
    End If
    If mobjShared.IsAppMode(modeCHCSI_GUI) Then  'SCR-40111
        'In PGUI mode the oAppt.ClinicianNCID holds the ClinicianIEN so need to get from
        'another source.
        sClinicianNCID = gobjClinicOps.GetProviderNCID(oAppt.FacilityNCID, oAppt.ClinicianNCID)
        GetWorkLoad_Options = gobjClinicOps.GetClinicWorkloadType(oAppt.FacilityNCID, sClinNCID, sClinicianNCID, oAppt.ApptType_Code, sErrMsg)
    Else
        GetWorkLoad_Options = gobjClinicOps.GetClinicWorkloadType(oAppt.FacilityNCID, sClinNCID, oAppt.ClinicianNCID, oAppt.ApptType_Code, sErrMsg)
    End If
    If oAppt.RelatedToInpatient = True And GetWorkLoad_Options = "NC" Then

        GetWorkLoad_Options = "N"
    End If
    
Cleanup:
 Set oAppt = Nothing
Exit Function
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "AppointmentsOps.cls - GetWorkLoad_Options. ", App.Title, vbCritical
End Function
Public Function GetEncounterFilingStatus_CHCS_GUI(sAppt_IEN As String, ByRef bHasBeenFiled As Boolean) As Boolean
'SCR-34640
Dim objAppt As CHCSII_AppointmentClient.Appointment
On Error GoTo ErrHandler

    If Not mobjShared.IsAppMode(modeCHCSI_GUI) Then Exit Function
    If LenB(sAppt_IEN) = 0 Then GoTo Cleanup
    If IsNumeric(sAppt_IEN) = False Then GoTo Cleanup
    Set objAppt = GetAppointment_CHCS_GUI(sAppt_IEN)
    If objAppt Is Nothing Then GoTo Cleanup
    If LenB(objAppt.ADM_Encounter) = 0 Then
        bHasBeenFiled = False
    ElseIf UCase$(objAppt.ADM_Encounter) = "COMPLETE" Then
        bHasBeenFiled = True
    Else
        GoTo Cleanup
    End If
    GetEncounterFilingStatus_CHCS_GUI = True
Cleanup:
    Set objAppt = Nothing
    
Exit Function
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "AppointmentsOps.cls - GetEncounterFilingStatus_CHCS_GUI. ", App.Title, vbCritical
End Function
Public Function GetAppointmentStatus_CHCS_GUI(sAppt_IEN As String, ByRef sApptStatus As String) As Boolean
'SCR-34640
Dim objAppt As CHCSII_AppointmentClient.Appointment
On Error GoTo ErrHandler
'Possible Appointment Status values:
'· Pending
'· KEPT
'· Cancel
'· NO - Show
'· WALK-IN
'· S-CALL
'· TEL - CON
'· LWOBS
'· ADMIN
'· OCC - SVC
'· WAITREQ
'· BOOKED

    If Not mobjShared.IsAppMode(modeCHCSI_GUI) Then Exit Function
    If LenB(sAppt_IEN) = 0 Then GoTo Cleanup
    If IsNumeric(sAppt_IEN) = False Then GoTo Cleanup
    Set objAppt = GetAppointment_CHCS_GUI(sAppt_IEN)
    If objAppt Is Nothing Then GoTo Cleanup
    sApptStatus = UCase$(objAppt.ApptStatus)
    
    GetAppointmentStatus_CHCS_GUI = True
Cleanup:
    Set objAppt = Nothing
    
Exit Function
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "AppointmentsOps.cls - GetAppointmentStatus_CHCS_GUI. ", App.Title, vbCritical
End Function

Public Function GetAppointment(rsApptid As String, Optional bUseCache As Boolean = True, Optional sFacilityNCID As String) As CHCSII_AppointmentClient.Appointment
Dim sSQLStatement As String
Dim rs As ADODB.Recordset
Dim lFacNCID As Long 'SCR-42490

    If LenB(rsApptid) = 0 Then GoTo Cleanup
    If rsApptid = "0" Then GoTo Cleanup
    If IsNumeric(rsApptid) = False Then GoTo Cleanup
    If mobjShared.IsAppMode(modeCHCSI_GUI) Then
        Set GetAppointment = GetAppointment_CHCS_GUI(rsApptid, bUseCache)
        GoTo Cleanup
    End If
    If mobjShared.IsAppMode(modeCDR) Or mobjShared.IsAppMode(modeTheater) Then
    
        'SCR-32352
        If bUseCache = True Then
            'We will first search the cache of appts, if found we will return that appt object.
            'If not found then we will call stored procedure.
            If mcolAppointmentsCache Is Nothing Then
                Set mcolAppointmentsCache = New Collection
                If mobjShared.IsAppMode(modeTheater) Or mobjShared.IsAppMode(modeITT) Then
                    GoTo CallSQL
                Else
                    GoTo CallStoredProc
                End If
            End If
            If mcolAppointmentsCache.Count = 0 Then
                If mobjShared.IsAppMode(modeTheater) Or mobjShared.IsAppMode(modeITT) Then
                    GoTo CallSQL
                Else
                    GoTo CallStoredProc
                End If
            End If
            On Error Resume Next
            Set GetAppointment = mcolAppointmentsCache.Item(rsApptid)
            If Err Then
                Err.Clear
                Debug.Print "ApptID: " + rsApptid + " NOT found in cache"
                If mobjShared.IsAppMode(modeTheater) Or mobjShared.IsAppMode(modeITT) Then
                    GoTo CallSQL
                Else
                    GoTo CallStoredProc
                End If
            Else
                Debug.Print "ApptID: " + rsApptid + " found in cache"
                If GetAppointment.RelatedToInpatient = True And Not mobjShared.IsAppMode(modeTheater) Then
                    GetAppointment.MeprsCode = GetAppointment.Inpatient_MEPRS_Code
                End If
                GoTo Cleanup
            End If
        Else
            If mobjShared.IsAppMode(modeTheater) Or mobjShared.IsAppMode(modeITT) Then GoTo CallSQL
        End If
        
CallStoredProc:
        Set rs = New ADODB.Recordset
        If LenB(sFacilityNCID) = 0 Then
            lFacNCID = mobjLogon.FacilityNCID
        Else
            lFacNCID = CLng(sFacilityNCID)
        End If
        Set rs = DoSearch_Oracle_SP(APPT_QUERY, , , , , lFacNCID, rsApptid)  'SCR-42490
        Debug.Print "ApptID: " + rsApptid + " calling Stored Procedure "
        If rs Is Nothing Then GoTo Cleanup
        If Not rs.BOF And Not rs.EOF Then
            Set GetAppointment = New CHCSII_AppointmentClient.Appointment
            GetAppointment.DeserializeAppts_SP rs
'            If bUseCache = False Then 'SCR-42490
                If Not GetAppointment Is Nothing Then
                    If GetAppointment.RelatedToInpatient = True Then
                        GetAppointment.MeprsCode = GetAppointment.Inpatient_MEPRS_Code
                    End If
                    On Error Resume Next
                    mcolAppointmentsCache.Remove GetAppointment.ApptId 'Clear from cache.
                    Err.Clear
                    mcolAppointmentsCache.Add GetAppointment, GetAppointment.ApptId
                    Debug.Print "ApptID: " + rsApptid + " refreshed to cache"
                End If
'            End If 'SCR-42490
        End If
        GoTo Cleanup
    Else
CallSQL:
        If mobjShared.IsAppMode(modeITT) Then
          sSQLStatement = "Select " & gsSelectSQLServer & " and Appt_Id = " & rsApptid
        ElseIf mobjShared.IsAppMode(modeTheater) Then
            If mobjShared.UseMySQLDB Then
                sSQLStatement = sMySQLSelect & sWhereMySQL & sQueryWhere2MySQL & " and Appt_Id = " & rsApptid
            Else
                sSQLStatement = sOracleSelect & sWhereOracle & sQueryWhere2Oracle & " and Appt_Id = " & rsApptid
            End If
        Else
          MsgBxARMd "Invalid app mode called in GetAppointment. Unable to retrieve appointment for" + _
          " ApptID: " + rsApptid
          GoTo Cleanup
        End If
    
        mobjSQL.Execute sSQLStatement
        If Not mobjSQL.EOF Then
            Set GetAppointment = New CHCSII_AppointmentClient.Appointment
            GetAppointment.Deserialize mobjSQL
            
                If Not GetAppointment Is Nothing Then  'SCR-42772
                    On Error Resume Next
                    mcolAppointmentsCache.Remove GetAppointment.ApptId 'Clear from cache.
                    If Err Then Err.Clear
                    mcolAppointmentsCache.Add GetAppointment, GetAppointment.ApptId
                    Debug.Print "ApptID: " + rsApptid + " refreshed to cache"
                End If
            
            
        End If
    End If
'SetRelatedToInpatient:
'    If GetAppointment.RelatedToInpatient = True Then
'        GetAppointment.MeprsCode = GetAppointment.Inpatient_MEPRS_Code
'    End If
Cleanup:
    Set rs = Nothing
End Function
Public Function GetAllAppointmentsForPatient(sPatientUnitNumber As String, Optional sApptStatus As String) As Collection
Dim sStatusCriteria As String
Dim rs As ADODB.Recordset
Dim cAppointments As Collection
    On Error GoTo ErrHandler
    If LenB(mobjShared.CmdLineSwitches.Item("CAC")) = 0 Then Exit Function
    If LenB(sPatientUnitNumber) = 0 Then Exit Function
    
    If LenB(sApptStatus) <> 0 Then
        Select Case UCase$(sApptStatus)
            Case "PENDING"
                sStatusCriteria = " AND (Appt_Status = 'PENDING') "
                
            Case "CHECKEDIN" '(Checked-In or InProgress or Waiting)
                sStatusCriteria = " AND (Encounter_Status = 1 OR ( Encounter_number IS NULL AND encounter_status IS NULL AND APPT_Status = 'KEPT')" + _
                                   " OR Encounter_Status = 3 OR Encounter_Status = 2) "
            Case "INCOMPLETE"
                sStatusCriteria = " AND ((Appt_Status = 'PENDING' OR (Appt_Status = 'KEPT' AND NOT Encounter_Status IN (4, 6)))) "
                
            Case "ALL"
                sStatusCriteria = vbNullString
            Case Else
                Exit Function
                
        End Select

    End If
    sStatusCriteria = sStatusCriteria + " AND PATIENT_UNIT_NUMBER = " + sPatientUnitNumber + " "
    If mobjShared.IsAppMode(modeTheater) Then 'SCR-52771
        Set GetAllAppointmentsForPatient = DoSearch_GEMS(0, sStatusCriteria, vbNullString, vbNullString)
    
    Else
        Set rs = New ADODB.Recordset
        Set rs = DoSearch_Oracle_SP(APPT_QUERY, , , sStatusCriteria, 0, mobjLogon.FacilityNCID)
        
        If rs Is Nothing Then GoTo Cleanup
        If rs.BOF And rs.EOF Then
            GoTo Cleanup
        End If
        Set cAppointments = New Collection
        LoadAppointments_SP cAppointments, rs
        Set GetAllAppointmentsForPatient = cAppointments
    End If
Cleanup:
    Set cAppointments = Nothing
    Set rs = Nothing
    Exit Function
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "AppointmentsOps.cls - GetAllAppointmentsForPatient", App.Title, vbCritical
    Set rs = Nothing
    Set cAppointments = Nothing

End Function
Public Function GetAppointment_CHCS_GUI(rsApptid As String, Optional bUseCache As Boolean = True) As CHCSII_AppointmentClient.Appointment
Dim colTemp As Collection
On Error GoTo ErrHandler
    
    If Not mobjShared.IsAppMode(modeCHCSI_GUI) Then Exit Function
    If LenB(rsApptid) = 0 Then Exit Function
    If IsNumeric(rsApptid) = False Then Exit Function
    If mcolAppointmentsCache Is Nothing Then
        Set mcolAppointmentsCache = New Collection
    End If
    If bUseCache = True Then
        'Try to pull from cache
        On Error Resume Next
        Set GetAppointment_CHCS_GUI = mcolAppointmentsCache.Item(rsApptid)
        If Err Then
            Err.Clear
            Debug.Print "ApptID: " + rsApptid + " NOT found in cache"
            GoTo CallchcsInterface
        
        Else
            Debug.Print "ApptID: " + rsApptid + " found in cache"
            GoTo Cleanup
        End If
    End If
    
CallchcsInterface:
    'if not found in cache the make new single search for appt. In DoSearch_CHCS_GUI it will be added to cache.
        
        Set colTemp = DoSearch_CHCS_GUI(True, , , , , True, rsApptid)
        If Not colTemp Is Nothing Then
            If colTemp.Count = 1 Then
                Set GetAppointment_CHCS_GUI = colTemp.Item(1)
                On Error Resume Next
                mcolAppointmentsCache.Remove GetAppointment_CHCS_GUI.ApptIEN 'Clear from cache.
                Err.Clear
                mcolAppointmentsCache.Add GetAppointment_CHCS_GUI, GetAppointment_CHCS_GUI.ApptIEN
                Debug.Print "ApptID: " + rsApptid + " refreshed to cache"
            End If
        End If
        
        
Cleanup:
    Set colTemp = Nothing
Exit Function
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "AppointmentsOps.cls - GetAppointment_CHCS_GUI", App.Title, vbCritical
End Function


Public Function PatientHasOpenEncounters(rlPatientUnitNumber As Long) As Boolean
On Error GoTo ErrHandler

    If rlPatientUnitNumber = 0 Then Exit Function
    'PERF STORED PROC
    If mobjShared.IsAppMode(modeCDR) And Not mobjShared.IsAppMode(modeITT) Then
        If PatientHasOpenEncounters_SP(CStr(rlPatientUnitNumber)) Then
           PatientHasOpenEncounters = True
        End If
    Else
        mobjSQL.Execute "select appt_id from appointment where patient_unit_number = " & rlPatientUnitNumber _
            & " and appt_status = 'KEPT' and encounter_number is not null and encounter_status <> " & EncCompleted
        If Not mobjSQL.EOF And Not mobjSQL.BOF Then
            If mobjSQL.RecordCount > 0 Then
                PatientHasOpenEncounters = True
            End If
        End If
    End If
Exit Function
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "AppointmentsOps.cls - PatientHasOpenEncounters", App.Title, vbCritical
    
End Function

Public Function GetNewAppointment() As CHCSII_AppointmentClient.Appointment
    Set GetNewAppointment = New CHCSII_AppointmentClient.Appointment
End Function

Private Function AddApptTransaction(ByVal ApptTransID As String, ByVal rsApptid As String, ByVal TransType As ApptTransTypeEnum) As String
'DataLayer Item
    Dim sSeq As String
On Error GoTo ErrHandler
        
    If TransType = ApptTransInsert Then
        AddApptTransaction = "insert into appt_transaction(appt_trans_id, appt_id, trans_type," _
            & " trans_date) select " & ApptTransID & ", " & rsApptid _
            & ", " & TransType & ", date_created from appointment where appt_id = " & rsApptid
    ElseIf TransType = ApptTransInpatient Then
        AddApptTransaction = "insert into appt_transaction(appt_trans_id, appt_id, trans_type," _
            & " trans_date) select " & ApptTransID & ", " & rsApptid _
            & ", " & TransType & ", date_created from appointment where appt_id = " & rsApptid
    Else
        AddApptTransaction = "insert into appt_transaction(appt_trans_id, appt_id, trans_type," _
            & " trans_date) select " & ApptTransID & ", " & rsApptid _
            & ", " & TransType & ", date_modified from appointment where appt_id = " & rsApptid
    End If
    Exit Function
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "AppointmentsOps.cls - AddApptTransaction", App.Title, vbCritical
End Function

Public Function SaveNewAppointment(robjAppt As CHCSII_AppointmentClient.Appointment) As String
    Dim colStmts As Collection
    Dim sApptID As String
    Dim sApptTransID As String
    
    On Error GoTo ErrHandler
    
    If Not robjAppt Is Nothing Then
        If Len(robjAppt.ApptReason) > 39999 Then
            MsgBxARMd "Unable to save the this new visit.  The appointment reason is too long " + _
            "to be saved to the datasource."
            Exit Function
        End If
    End If
    
    sApptID = mobjSQL.GetNextID(esAPPT_ID)
    
    Set colStmts = New Collection
    colStmts.Add robjAppt.Serialize(sApptID, mobjLogon.FacilityNCID)
    sApptTransID = mobjSQL.GetNextID(esAPPT_TRANS_ID)
    colStmts.Add AddApptTransaction(sApptTransID, sApptID, ApptTransInsert)
    mobjSQL.ExecuteMulti colStmts
    SaveNewAppointment = sApptID
    robjAppt.ApptId = sApptID
    'SCR-RNDS
    If robjAppt.ApptType = "RNDS*" And robjAppt.Inpatient_Register_Number <> vbNullString Then
        If mobjShared.IsAppMode(modeITT) Then
             Insert_PatientAdmissionRecord_SQL sApptID, robjAppt.Inpatient_Register_Number
        Else
            Insert_PatientAdmissionRecord_SP sApptID, robjAppt.Inpatient_Register_Number
        End If
    End If
    'SCR-51307
    If mobjShared.IsAppMode(modeTheater) Then
        'If this is not a future/pending appt then robjAppt.PendingComplaints will be nothing and nothing will be saved.
        If robjAppt.ApptDateTime > Now Then
            GEMS_SavePendingComplaints robjAppt
        End If
    End If
    Set colStmts = Nothing
    
    'Per Brian Mowbray this is no longer needed. S.McAvoy 2/20/04
    'If Not mobjComm.CmdLineSwitch("CITA") = vbNullString Then
      'Added this function back and changed it. Will use for DMDC DEERS Immunization Inquires for appoinments.
      'It was added before but never used so code was commented out. Brian Mowbray 5/4/01
     ' AddPatientForImmunizationPull mobjPatient.UnitNumber, robjAppt.ApptDateTime, mobjSQL
    'End If
Exit Function

ErrHandler:
    Dim sTxt As String
    Dim i As Integer
    If Not colStmts Is Nothing Then
        For i = 1 To colStmts.Count
            sTxt = sTxt + CStr(colStmts.Item(i)) + vbCrLf
        Next
    
    End If
    mobjShared.ShowVBError Err.Number, Err.Description & vbCrLf + sTxt, "AppointmentsOps.cls - SaveNewAppointment", App.Title, vbCritical


End Function
Private Sub GEMS_SavePendingComplaints(robjAppt As CHCSII_AppointmentClient.Appointment)
    'SCR-51307
    Dim colStmts As Collection
    Dim sql As String
    Dim i As Integer
    Dim sErrTxt As String
    On Error GoTo ErrHandler
    
    If robjAppt Is Nothing Then Exit Sub
    With robjAppt
        If .PendingComplaints Is Nothing Then Exit Sub
        If .PendingComplaints.Count = 0 Then Exit Sub
        Set colStmts = New Collection
        For i = 1 To .PendingComplaints.Count
            sql = "Insert into Appt_Pending_Complaints Values(" + CStr(.ApptId) + ", " + oApptConn.SQLQuote(.PendingComplaints.Item(i)) + ")"
            colStmts.Add sql
        Next
    End With
    mobjSQL.ExecuteMulti colStmts
    Set colStmts = Nothing
    Exit Sub
ErrHandler:
    For i = 1 To robjAppt.PendingComplaints.Count
        sErrTxt = sErrTxt + robjAppt.PendingComplaints.Item(i) + vbCrLf
    Next
        mobjShared.ShowVBError Err.Number, "Unable to save complaints for this appointment." & Err.Description & vbCrLf & vbCrLf & _
            sErrTxt, "AppointmentsOps.cls - GEMS_SavePendingComplaints", App.Title, vbCritical

End Sub
Public Function SavePendingCHCSAppointment(sql As String) As Boolean
'SCR-35075
    
    If LenB(mobjShared.CmdLineSwitches.Item("WAM")) > 0 And mobjShared.IsAppMode(modeCDR) Then
        mobjSQL.Execute sql
        SavePendingCHCSAppointment = True
    Else
        SavePendingCHCSAppointment = False
    End If
     
End Function
Public Function UnlockApptRecord(objAppt As CHCSII_AppointmentClient.Appointment) As Boolean
On Error Resume Next

    If objAppt Is Nothing Then Exit Function
    If objAppt.ApptStatus = "PENDING" And objAppt.EncounterNumber = "-1" Then
        mobjSQL.Execute "Update Appointment Set ENCOUNTER_NUMBER = Null, " + _
        "Appt_Lock_Info = Null " + _
        " Where Appt_ID = " + objAppt.ApptId
        
        If Err Then Err.Clear
    End If
    UnlockApptRecord = True
End Function

'DEM Modified 01/05/2000
'MAG 05/21/2001 Added optional roAppointment Object to speed up setting objAppointment Object
'--- SCR-9485 MSolano 9/20/01 added bArrivedByAmbulance
'SCR 24368 Phil Crowder 6/6/02 Added optional parameter to allow the check in of an appointment
'without sending a Select Patient message to Core.  This should only be used if the patient
'does not need to be displayed in the folder list.
'SCR 24368, 'SCR-24498
'38127  Open  Goodfellow - Phantom Encounter / lost data
'38295  Open  DB change: Add column CDRPlus.Appointment.Appt_Lock_Info
'39178  Added enhancement to locking process.
'DB change: Add column CDRPlus.Appointment.Appt_Lock_Info
'Other SCRs noted in this function:
'SCR-9485
'SCR-17624
'SCR-31315
'SCR-12372
'SCR-34683
'SCR-34683
'
'SCR-43626 & 38479
'Implemented Appt locking within Stored Procedure: Appointments_Vr3_0_Pkg.
'CheckIn_1_SetApptVariables = False
'    FAILED_TO_SET_APPT_VARS --> to Cleanup
'
'CheckIn_2_Init_LOCK_Process = True
'    APPT_LOCK_SUCCESS
'
'CheckIn_2_Init_LOCK_Process = False
'    ENC_NUM_FOUND
'    APPT_LOCKED_BY_OTHER_USER  --> msg displayed --> to Cleanup"
'    APPT_LOCK_FAILED --> errhnd msg displayed --> to Cleanup
'    CONSULT_COMPLETED_IN_CHCSI --> to Cleanup
'
'CheckIn_3_CreateEncounter = True
'    ENC_CREATE_OK
'
'CheckIn_3_CreateEncounter = False
'    ENC_FAILED --> to Cleanup
'    ENC_ID_FAILED --> to Cleanup
'    ENC_ID_ZERO --> to Cleanup
'    ENC_FAILED --> errhnd w/ msg dispalyed --> to Cleanup
'
'CheckIn_4_UpdateApptRecord = False
'    UPDATE_APPT_REC_FAILED w/ msg displayed  --> to Cleanup
'
'CheckIn = True
'    APPT_REC_UPDATE_OK
'
'CheckIn = False --> errhnd w/msg displayed
'   UPDATE_APPT_REC_FAILED

Public Function CheckIn(rsApptid As String, _
                        Optional rsNoteTitle As String, _
                        Optional rsNoteText As String, _
                        Optional rsEncounterNumber As String, _
                        Optional roAppointment As CHCSII_AppointmentClient.Appointment, _
                        Optional bArrivedByAmbulance As Boolean = False, _
                        Optional bSelectPatient As Boolean = True, _
                        Optional sWorkLoadTypeOption As String = "N", _
                        Optional sReturnStatus As String, _
                        Optional objHIPAAclt As CHCSIIEncounterOps.EncAccidentData) _
                    As Boolean
     
        
    Dim objEncounter As CHCSIIEncounterOps.Encounter
    Dim objAppointment As CHCSII_AppointmentClient.Appointment
    Dim eEncType As EncType
    Dim bIsConsultOrder As Boolean
    Dim bIsObservation As Boolean
    Dim sEncNumber As String
    Dim sAppt_Lock_Info As String
    Dim sApptDB_Lock_Info As String
    Dim sTempEncNum As String 'Used as a double check to alert user if Phantom occurs.
                              'Directly after the encounter is created we set this var.
                              'At the end of this function when we do a "GetAppoitment",
                              'we look at the encounterNum attached to the appt obj.
                              'That num should match sTempEncNum.
    On Error GoTo ErrHandler
    If IsMissing(roAppointment) Or roAppointment Is Nothing Then
        Set objAppointment = GetAppointment(rsApptid, True)
    Else
        Set objAppointment = roAppointment
    End If
    
    If objAppointment Is Nothing Then
        sReturnStatus = "APPT_OBJECT_FAIL"
        GoTo Cleanup
    End If
    With objAppointment
        If mobjPatient.UnitNumber <> .PatientUnitNumber And bSelectPatient Then
            mobjComm.Message cwiSELECT_PATIENT, .PatientUnitNumber & ";" & .PatientIEN, NCID_CW_APPLICATION, mlProgId
        End If
    End With
    rsNoteTitle = Trim$(rsNoteTitle)
    rsNoteText = Trim$(rsNoteText)
    
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    If Not CheckIn_1_SetApptVariables(objAppointment, rsNoteTitle, bIsConsultOrder, bIsObservation, eEncType) Then
        sReturnStatus = "FAILED_TO_SET_APPT_VARS"
        GoTo Cleanup
    End If
    
    sAppt_Lock_Info = UCase$(Left$(Format$(Now, "mm/dd/yyyy hh:mm:ss") + "|" + CStr(mobjUser.UserId) + "|" + gsComputerName, 99))
    '   "06/30/03 11:43:34|UserNCID|gsComputerName"  ---> Gets Inserted to the Appt_Lock_Info field.
    
    
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    If Not CheckIn_2_Init_LOCK_Process(sEncNumber, sApptDB_Lock_Info, sAppt_Lock_Info, objAppointment.ApptId, sReturnStatus) Then
        GoTo Cleanup 'If we fail here it means that the appt could not be locked.
                     ' e.g. -1 and Appt_Lock_Code inserted.
    Else
        objAppointment.EncounterNumber = sEncNumber
    End If
    
    
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    If LenB(objAppointment.OrderNumber) <> 0 Then
        'Check to see if Appointment is associated to a consult
        'Need to find associated consult and get info for note
        rsNoteText = CreateConsultOrder(objAppointment)
        rsNoteTitle = "Consult Order"
        bIsConsultOrder = True
    End If
    
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    If Not CheckIn_3_CreateEncounter(objEncounter, objAppointment, eEncType, bArrivedByAmbulance, bIsConsultOrder, sWorkLoadTypeOption, bIsObservation, rsNoteTitle, rsNoteText, sReturnStatus, objHIPAAclt) Then
        CheckIn_LOCK_or_UNLOCK objAppointment.ApptId, sAppt_Lock_Info, sApptDB_Lock_Info, sEncNumber, "UNLOCK"
        GoTo Cleanup
    Else
        'Only a valid Encouter object returned and EncID must be number > 0
        sTempEncNum = CStr(objEncounter.EncounterID)
        sEncNumber = sTempEncNum
        If Not IsMissing(rsEncounterNumber) Then
            rsEncounterNumber = CStr(objEncounter.EncounterID)
        End If
    End If
    
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    If Not CheckIn_4_UpdateApptRecord(objAppointment, objEncounter) Then
        Set objAppointment = GetAppointment(objAppointment.ApptId, False)
        CheckIn_LOCK_or_UNLOCK objAppointment.ApptId, sAppt_Lock_Info, sApptDB_Lock_Info, sEncNumber, "UNLOCK"
        sReturnStatus = "UPDATE_APPT_REC_FAILED"
        GoTo Cleanup
    End If

    sReturnStatus = "APPT_REC_UPDATE_OK" '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    
    CheckIn = True
Cleanup:
    If IsMissing(roAppointment) = False Then 'GetAppointment is always called.
        If objAppointment Is Nothing Then Exit Function
        Set roAppointment = GetAppointment(objAppointment.ApptId, False)
        DoEvents
        If sReturnStatus = "APPT_REC_UPDATE_OK" Then 'Check for Phantom encounter.
            If Not roAppointment Is Nothing Then
                If sTempEncNum <> roAppointment.EncounterNumber Then
                    MsgBxARMd "Possible duplicate encounter created for this appointment." & vbCrLf + vbCrLf + _
                        "Appointment ID = " + roAppointment.ApptId + vbCrLf + _
                        "Appointment Time = " + CStr(roAppointment.ApptDateTime) + vbCrLf + _
                        "Current Enc# Assigned = " + CStr(roAppointment.EncounterNumber) + vbCrLf + _
                        "Duplicate Enc# Assigned = " + sTempEncNum + vbCrLf + _
                        "Facility NCID = " + roAppointment.FacilityNCID + vbCrLf + _
                        "Patient Name = " + roAppointment.PatientName + vbCrLf + _
                        "User ID = " + CStr(mobjUser.UserId), , "WARNING - Duplicate Encounter Created"
                        
                End If
            End If
        
        End If
    End If
    Set objAppointment = Nothing
    Set objEncounter = Nothing
    
    Exit Function
ErrHandler:
    sReturnStatus = "UPDATE_APPT_REC_FAILED"
    If Not objAppointment Is Nothing Then
        MsgBxARMd "Error in " & "CHCSII_AppointmentClient.AppointmentOps.CheckIn: " & Err.Description & vbCrLf + vbCrLf + _
            "Appointment ID = " + objAppointment.ApptId + vbCrLf + _
            "Appointment Time = " + CStr(objAppointment.ApptDateTime) + vbCrLf + _
            "Encounter Number = " + sEncNumber + vbCrLf + _
            "Facility NCID = " + objAppointment.FacilityNCID + vbCrLf + _
            "Patient Name = " + objAppointment.PatientName + vbCrLf + _
            "User ID = " + CStr(mobjUser.UserId)

    Else
        mobjShared.ShowVBError Err.Number, Err.Description, "AppointmentsOps.cls - CheckIn", App.Title, vbCritical
    End If
    GoTo Cleanup
End Function

Private Function CheckIn_2_Init_LOCK_Process(sEncNumber As String, sApptDB_Lock_Info As String, sAppt_Lock_Info As String, sApptID As String, sReturnStatus As String) As Boolean
Dim CurrentDBLockData() As String
On Error GoTo ErrHandler


    If mobjShared.IsAppMode(modeTheater) Or mobjShared.IsAppMode(modeITT) Then
        'In theater mode, no need to lock appt.
        'SCR-46394
        'But we do want to at least check to see if valid appt record before
        'we create an encounter.
        Dim objAppt As CHCSII_AppointmentClient.Appointment
        Set objAppt = GetAppointment(sApptID, False)
        If objAppt Is Nothing Then
            sReturnStatus = "APPT_NOT_FOUND"
            Exit Function
        ElseIf objAppt.EncounterNumber <> vbNullString Then
            sReturnStatus = "ENC_NUM_FOUND"
            Exit Function
        Else
            sReturnStatus = "APPT_LOCK_SUCCESS"
        End If
        Set objAppt = Nothing
    Else
        If Not CheckIn_LOCK_or_UNLOCK(sApptID, sAppt_Lock_Info, sApptDB_Lock_Info, sEncNumber, "LOCK") Then
            If sEncNumber = "0" Then
                sReturnStatus = "CONSULT_COMPLETED_IN_CHCSI"
                Exit Function
            End If
            If LenB(sEncNumber) <> 0 And sEncNumber <> "0" And sEncNumber <> "-1" And IsNumeric(sEncNumber) Then
                sReturnStatus = "ENC_NUM_FOUND" 'Means we could not lock and another user
                                                'has already assigned an encounter# to this
                                                'appt record.
                Exit Function
            End If
            If LenB(sApptDB_Lock_Info) = 0 Then Exit Function
            If InStr(1, sApptDB_Lock_Info, "|", vbTextCompare) = 0 Then Exit Function
            CurrentDBLockData = Split(sApptDB_Lock_Info, "|", , vbTextCompare)
            If UBound(CurrentDBLockData) < 2 Then Exit Function
            MsgBxARMd "The selected Appointment record (" + sApptID + ") is currently " + _
                  "being opened by user: " + CurrentDBLockData(1) + ", Workstation: " + CurrentDBLockData(2) + "." + vbCrLf + _
                  "Please refresh the appointment grid reopen appointment."
            sReturnStatus = "APPT_LOCKED_BY_OTHER_USER"
            Exit Function
        Else
            sReturnStatus = "APPT_LOCK_SUCCESS"
        End If
    End If
    
    CheckIn_2_Init_LOCK_Process = True
    Erase CurrentDBLockData
    Exit Function
ErrHandler:
        MsgBxARMd "Error in " & "CHCSII_AppointmentClient.AppointmentOps.CheckIn_2_Init_LOCK_Process: " & Err.Description & vbCrLf + vbCrLf + _
            "Appointment ID = " + sApptID + vbCrLf + _
            "Appt_Lock_Info = " + sAppt_Lock_Info + vbCrLf + _
            "ApptDB_Lock_Info = " + sApptDB_Lock_Info + vbCrLf + _
            "User ID = " + CStr(mobjUser.UserId)
        sReturnStatus = "APPT_LOCK_FAILED"

End Function

Private Function CheckIn_LOCK_or_UNLOCK(sApptID As String, sAppt_Lock_Info As String, _
        sApptDB_Lock_Info As String, sEncNumber As String, sLockType As String) As Boolean

On Error GoTo ErrHandler

Dim oDatabase As CHCSII_CONN.Conn
Dim objCmd As CHCSII_Command.CHCSII_Cmd
Dim rs_ApptLock_SP As ADODB.Recordset
Dim aAppt_Lock_Info() As String
Dim sLockStatus As Boolean
Dim sReturnCode As String
    If mobjShared.IsAppMode(modeTheater) Or mobjShared.IsAppMode(modeITT) Then
        CheckIn_LOCK_or_UNLOCK = True
        GoTo Cleanup
    End If

    aAppt_Lock_Info = Split(sAppt_Lock_Info, "|", , vbTextCompare)
    If UBound(aAppt_Lock_Info) <> 2 Then
        CheckIn_LOCK_or_UNLOCK = False
        GoTo Cleanup
    End If
    sLockType = UCase$(Trim$(sLockType))
    If sLockType <> "UNLOCK" And sLockType <> "LOCK" Then
        CheckIn_LOCK_or_UNLOCK = False
        GoTo Cleanup
    End If
    If sLockType = "UNLOCK" Then
        sApptDB_Lock_Info = vbNullString
    End If
    If mobjDASsp Is Nothing Then Set mobjDASsp = Set_CHCSII_CONN_SP_connection()
    Set objCmd = New CHCSII_Command.CHCSII_Cmd
    objCmd.SPName = CURR_SP_VER + ".Lock_UNLock_Appt"
    DoEvents

    
    Call objCmd.AddParam(False, False, "textString", sApptID)
    Call objCmd.AddParam(False, False, "textString", aAppt_Lock_Info(0))
    Call objCmd.AddParam(False, False, "textString", aAppt_Lock_Info(1))
    Call objCmd.AddParam(False, False, "textString", aAppt_Lock_Info(2))
    Call objCmd.AddParam(False, False, "textString", sLockType)
    'This Parameter is defined as a cursor, used for output
    Call objCmd.AddParam(True, True, vbNullString, vbNullString)
    
    Set rs_ApptLock_SP = mobjDASsp.OpenRecordsetSP(objCmd)
    'If successful, SP Returns:
    'TRUE|APPT_LOCK_SUCCESS|-1|11/06/2003 17:33:11|BLK2SJTST.COM@18FRIEND|MCAVOY_2K
    If rs_ApptLock_SP Is Nothing Then GoTo Cleanup
        If rs_ApptLock_SP.EOF Then GoTo Cleanup
        If IsNull(rs_ApptLock_SP("RETURN_VALUE")) = False Then
            If LenB(rs_ApptLock_SP("RETURN_VALUE")) <> 0 Then
                aAppt_Lock_Info = Split(rs_ApptLock_SP("RETURN_VALUE"), "|", , vbTextCompare)
                If UBound(aAppt_Lock_Info) >= 0 Then
                    sLockStatus = aAppt_Lock_Info(0)
                End If
                If UBound(aAppt_Lock_Info) >= 1 Then
                    sReturnCode = aAppt_Lock_Info(1)
                End If
                If UBound(aAppt_Lock_Info) >= 2 Then
                    sEncNumber = aAppt_Lock_Info(2)
                End If
                If UBound(aAppt_Lock_Info) >= 5 Then
                    sApptDB_Lock_Info = aAppt_Lock_Info(3) + "|" + aAppt_Lock_Info(4) + "|" + aAppt_Lock_Info(5)
                End If
                If sLockType = "LOCK" Then
                    If sLockStatus = "TRUE" And sReturnCode = "APPT_LOCK_SUCCESS" And _
                        sEncNumber = "-1" Then
                        CheckIn_LOCK_or_UNLOCK = True
                    End If
                ElseIf sLockType = "UNLOCK" Then
                    If sLockStatus = "TRUE" And sReturnCode = "APPT_UNLOCK_SUCCESS" And _
                        LenB(sEncNumber) = 0 And LenB(sApptDB_Lock_Info) = 0 Then
                        CheckIn_LOCK_or_UNLOCK = True
                    End If
                Else
                    CheckIn_LOCK_or_UNLOCK = False
                End If
            End If
        End If
Cleanup:
    Set objCmd = Nothing
    Set oDatabase = Nothing
    Set rs_ApptLock_SP = Nothing
    Erase aAppt_Lock_Info
    Exit Function
ErrHandler:
    If Err.Description = "Service returned an error." Then
        MsgBxARMd "The database stored procedure " + CStr(objCmd.SPName) + " called from CHCSII_AppointmentClient.AppointmentOps.CheckIn_LOCK_or_UNLOCK caused an error: " & Err.Description & vbCrLf + vbCrLf + _
            "The appointment could not be locked for the checkin process. " + vbCrLf + _
            "Appointment ID = " + sApptID + vbCrLf + _
            "Appt_Lock_Info = " + sAppt_Lock_Info + vbCrLf + _
            "       User ID = " + CStr(mobjUser.UserId)
    
    Else
        MsgBxARMd "Error in " & "CHCSII_AppointmentClient.AppointmentOps.CheckIn_LOCK_or_UNLOCK: " & Err.Description & vbCrLf + vbCrLf + _
            "Appointment ID = " + sApptID + vbCrLf + _
            "Appt_Lock_Info = " + sAppt_Lock_Info + vbCrLf + _
            "       User ID = " + CStr(mobjUser.UserId)
    End If
    sEncNumber = vbNullString
    GoTo Cleanup
End Function
Private Function CheckIn_4_UpdateApptRecord(objAppointment As CHCSII_AppointmentClient.Appointment, objEncounter As CHCSIIEncounterOps.Encounter) As Boolean
'DataLayer Item
Dim sEncNumber As String
Dim colStmts As Collection
Dim sRetCode As String
Dim sGemsComplaintErrTxt As String
'SCR-40190
Dim sRelatedToInpat As String
Dim sApptTransID As String

    On Error GoTo ErrHandler
    
    
    If objAppointment Is Nothing Then Exit Function
    If objEncounter Is Nothing Then Exit Function
    If CStr(objEncounter.EncounterID) = vbNullString Then Exit Function
    If IsNumeric(objEncounter.EncounterID) = False Then Exit Function
    If CStr(objEncounter.EncounterID) = "0" Then Exit Function
    sEncNumber = CStr(objEncounter.EncounterID)
    

    Set colStmts = New Collection
    'SCR-40190 for ITT only.FormatRelatedToInpatient
    colStmts.Add "update Appointment set Date_Modified = " & oApptConn.SQLDate(Now) & ", " _
        & " Appt_Status = 'KEPT'," _
        & FormatRelatedToInpatient(objAppointment.RelatedToInpatient) _
        & " Encounter_Number = " & objEncounter.EncounterID & "," _
        & " Encounter_Status = " & EncCheckedIn & ", Appt_Lock_Info = Null, " _
        & " CheckIn_Time = " & oApptConn.SQLDate(Now) _
        & " where Appt_ID = " & objAppointment.ApptId
    sApptTransID = mobjSQL.GetNextID(esAPPT_TRANS_ID)
    colStmts.Add AddApptTransaction(sApptTransID, objAppointment.ApptId, ApptTransKept)
    mobjSQL.ExecuteMulti colStmts
    
    If mobjShared.IsAppMode(modeTheater) Then
        If Not CheckIn_GEMS_SaveComplaints(objAppointment, sGemsComplaintErrTxt) Then
            MsgBxARMd "The appointment was saved, and the encounter was created but application could not save the chief complaints." + vbCrLf + vbCrLf + _
            "complaints and/or appointment comment. " + vbCrLf + sGemsComplaintErrTxt, , "AppointmentOps.Checkin"
        End If
    End If
  
    
    Set colStmts = Nothing
    CheckIn_4_UpdateApptRecord = True
    Exit Function
ErrHandler:
    If Not objAppointment Is Nothing Then
        MsgBxARMd "Error in " & "CHCSII_AppointmentClient.AppointmentOps.CheckIn_4_UpdateApptRecord: " & Err.Description & ", ErrCode: " & sRetCode & vbCrLf + vbCrLf + _
            "Appointment ID = " + objAppointment.ApptId + vbCrLf + _
            "Appointment Time = " + CStr(objAppointment.ApptDateTime) + vbCrLf + _
            "Encounter Number = " + sEncNumber + vbCrLf + _
            "Facility NCID = " + objAppointment.FacilityNCID + vbCrLf + _
            "Patient Name = " + objAppointment.PatientName + vbCrLf + _
            "User ID = " + CStr(mobjUser.UserId)

    Else
        mobjShared.ShowVBError Err.Number, Err.Description, "AppointmentsOps.cls - CheckIn_4_UpdateApptRecord", App.Title, vbCritical
    End If
    Set colStmts = Nothing
End Function
Private Function FormatRelatedToInpatient(bValue As Boolean) As String
    
    On Error GoTo ErrHandler
    If mobjShared.IsAppMode(modeTheater) Then Exit Function
    If Not mobjShared.IsAppMode(modeITT) Then Exit Function
    FormatRelatedToInpatient = " RELATED_TO_INPATIENT = '" + UCase(CStr(bValue)) + "', "
    Exit Function
ErrHandler:
 mobjShared.ShowVBError Err.Number, Err.Description, "AppointmentsOps.cls - FormatRelatedToInpatient", App.Title, vbCritical
End Function
Private Function CheckIn_UpdateApptRecord_SP(sApptID As String, sEncNum As String, sEncStatus As String, _
            sRelatedToInpatient As String, sRetCode As String) As Boolean
'SCR-40190 for ITT only
On Error GoTo ErrHandler

Dim oDatabase As CHCSII_CONN.Conn
Dim objCmd As CHCSII_Command.CHCSII_Cmd
Dim rs_ApptUpdate_SP As ADODB.Recordset
Dim aSP_Return() As String
Dim sReturnStatus As String
'Dim sReturnCode As String
Dim sCurrDateTime As String

    
    If LenB(sApptID) = 0 Or LenB(sEncNum) = 0 Or IsNumeric(sApptID) = False Or IsNumeric(sEncNum) = False Then
        sRetCode = "Invalid arguments in CheckIn_UpdateApptRecord_SP."
        CheckIn_UpdateApptRecord_SP = False
        GoTo Cleanup
    End If
    If mobjDASsp Is Nothing Then Set mobjDASsp = Set_CHCSII_CONN_SP_connection()
    Set objCmd = New CHCSII_Command.CHCSII_Cmd
    objCmd.SPName = CURR_SP_VER + ".CheckIn_Pending_Appt"
    DoEvents

    If LenB(sRelatedToInpatient) = 0 Then sRelatedToInpatient = "NULL"
    sRelatedToInpatient = UCase$(sRelatedToInpatient)
    sCurrDateTime = Format$(Now, "DDMMMYYYY HH:MM:SS")
    
    Call objCmd.AddParam(False, False, "textString", sApptID)
    Call objCmd.AddParam(False, False, "textString", sEncNum)
    Call objCmd.AddParam(False, False, "textString", sEncStatus)
    Call objCmd.AddParam(False, False, "textString", sCurrDateTime)
    Call objCmd.AddParam(False, False, "textString", sRelatedToInpatient)
    Call objCmd.AddParam(False, False, "textString", CStr(ApptTransKept))
    'This Parameter is defined as a cursor, used for output
    Call objCmd.AddParam(True, True, vbNullString, vbNullString)
    If mobjDASsp Is Nothing Then Set mobjDASsp = Set_CHCSII_CONN_SP_connection()
    Set rs_ApptUpdate_SP = mobjDASsp.OpenRecordsetSP(objCmd)
    'If successful, SP Returns:
    'TRUE|<1>'
    If rs_ApptUpdate_SP Is Nothing Then GoTo Cleanup
        If rs_ApptUpdate_SP.EOF Then GoTo Cleanup
        If IsNull(rs_ApptUpdate_SP("RETURN_VALUE")) = False Then
            If LenB(rs_ApptUpdate_SP("RETURN_VALUE")) <> 0 Then
                sRetCode = rs_ApptUpdate_SP("RETURN_VALUE")
                aSP_Return = Split(rs_ApptUpdate_SP("RETURN_VALUE"), "|", , vbTextCompare)
                If UBound(aSP_Return) >= 0 Then
                    sReturnStatus = aSP_Return(0)
                End If
'                If UBound(aSP_Return) >= 1 Then
'                    sReturnCode = aSP_Return(1)
'                End If
                If sReturnStatus = "TRUE" Then
                    CheckIn_UpdateApptRecord_SP = True
                End If
            End If
        End If
Cleanup:
    Set objCmd = Nothing
    Set oDatabase = Nothing
    Set rs_ApptUpdate_SP = Nothing
    Erase aSP_Return
    Exit Function
ErrHandler:
    If Err.Description = "Service returned an error." Then
        MsgBxARMd "The database stored procedure " + CStr(objCmd.SPName) + " called from CHCSII_AppointmentClient.AppointmentOps.CheckIn_UpdateApptRecord_SP caused an error: " & Err.Description & vbCrLf + vbCrLf + _
            "The appointment could not be updated for the checkin process. " + vbCrLf + _
            "Appointment ID = " + sApptID + vbCrLf + _
            "EncNumber = " + sEncNum + vbCrLf + _
            "User ID = " + CStr(mobjUser.UserId)
    
    Else
        MsgBxARMd "Error in " & "CHCSII_AppointmentClient.AppointmentOps.CheckIn_UpdateApptRecord_SP: " & Err.Description & vbCrLf + vbCrLf + _
            "Appointment ID = " + sApptID + vbCrLf + _
            "EncNumber = " + sEncNum + vbCrLf + _
            "User ID = " + CStr(mobjUser.UserId)
    End If
    
    GoTo Cleanup



End Function

Private Function CreateConsultOrder(objAppointment As CHCSII_AppointmentClient.Appointment) As String
    Dim objConsultOrderOps As Object
    Dim objConsultOrder As Object
    Const ConsultOrder_RTFView As Integer = 2
        Set objConsultOrderOps = CreateObject("CHCSII_Consult.ConsultOrderOps")
        
        On Error Resume Next
        With objConsultOrderOps
            Set .Config = mobjConfig
            Set .Comm = mobjComm
            Set .Logon = mobjLogon
            Set .Patient = mobjPatient
            Set .User = mobjUser
            Set objConsultOrder = objConsultOrderOps.GetConsultOrderByOrderNumber(objAppointment.OrderNumber, objAppointment.FacilityNCID)
            If Err Then
                Err.Clear
                MsgBxARMd "Unable to get consult order information during the appointment checkin process." & vbCrLf & "ERROR: " & Err.Description, vbExclamation + vbOKOnly, "Appointment Checkin"
                CreateConsultOrder = "<Information Not Available>"
            Else
                CreateConsultOrder = objConsultOrder.GetDataView(ConsultOrder_RTFView)
            End If
        End With

    Set objConsultOrder = Nothing
    Set objConsultOrderOps = Nothing


End Function
Private Function CheckIn_1_SetApptVariables(objAppointment As CHCSII_AppointmentClient.Appointment, _
            rsNoteTitle As String, bIsConsultOrder As Boolean, bIsObservation As Boolean, _
            eEncType As EncType) As Boolean

On Error GoTo ErrHandler

    Select Case objAppointment.ApptClassEnum
        Case ApptInpatient
            eEncType = Inpatient
        Case ApptAmbulatoryProcedure
            eEncType = Ambulatory
        Case Else
            eEncType = Outpatient
    End Select
    If rsNoteTitle = "Consult Order" Then
        bIsConsultOrder = True
    Else
        bIsConsultOrder = False
    End If
    If objAppointment.Observation = "1" Then
        bIsObservation = True
    ElseIf objAppointment.Observation = "0" Then
        bIsObservation = False
    Else
        bIsObservation = False
    End If
    CheckIn_1_SetApptVariables = True
    Exit Function
ErrHandler:
    If Not objAppointment Is Nothing Then
        MsgBxARMd "Error in " & "CHCSII_AppointmentClient.AppointmentOps.CheckIn_1_SetApptVariables: " & Err.Description & vbCrLf + vbCrLf + _
            "Appointment ID = " + objAppointment.ApptId + vbCrLf + _
            "Appointment Time = " + CStr(objAppointment.ApptDateTime) + vbCrLf + _
            "Facility NCID = " + objAppointment.FacilityNCID + vbCrLf + _
            "Patient Name = " + objAppointment.PatientName + vbCrLf + _
            "User ID = " + CStr(mobjUser.UserId)

    Else
            mobjShared.ShowVBError Err.Number, Err.Description, "AppointmentsOps.cls - CheckIn_1_SetApptVariables", App.Title, vbCritical
    End If

End Function
Private Function CheckIn_3_CreateEncounter(objEncounter As CHCSIIEncounterOps.Encounter, _
                objAppointment As CHCSII_AppointmentClient.Appointment, _
                eEncType As EncType, _
                bArrivedByAmbulance As Boolean, _
                bIsConsultOrder As Boolean, _
                sWorkLoadTypeOption As String, bIsObservation As Boolean, _
                rsNoteTitle As String, _
                rsNoteText As String, _
                sReturnStatus As String, _
                Optional objHIPAAclt As CHCSIIEncounterOps.EncAccidentData) As Boolean

On Error GoTo ErrHandler
    If mobjEncParent Is Nothing Then
        Set mobjEncParent = mobjComm.InitializeOLEServer(NCID_ENCOUNTER_APPLICATION)
    End If
    With objAppointment
        If mobjShared.IsAppMode(modeTheater) Then 'SCR-42772
            'SCR-46148 All notes from appointments are created by complaints
            'and not when we are creating the encounter.
            Set objEncounter = mobjEncParent.CreateEncounter( _
                                      CLng(.ApptId), _
                                      .ApptType, _
                                      eEncType, _
                                      .FacilityNCID, _
                                      .ClinicNCID, _
                                      .ClinicianNCID, _
                                      , _
                                      bArrivedByAmbulance, _
                                      .ApptDateTime)
                    
        
        Else
        
            Set objEncounter = mobjEncParent.CreateEncounter( _
                                      CLng(.ApptId), _
                                      .ApptType, _
                                      eEncType, _
                                      .FacilityNCID, _
                                      .ClinicNCID, _
                                      .ClinicianNCID, _
                                      , _
                                      bArrivedByAmbulance, _
                                      .ApptDateTime, , bIsConsultOrder, _
                                      .ApptReason, .ApptComment, _
                                      CLng(.PatientUnitNumber), _
                                      sWorkLoadTypeOption, _
                                      .Workload_Type, _
                                      bIsObservation, _
                                      .Inpatient_MEPRS_Code, _
                                      .MeprsCode, _
                                      .ApptIEN, _
                                      , _
                                      objHIPAAclt)
                                  
        End If
    End With

    If objEncounter Is Nothing Then
        sReturnStatus = "ENC_FAILED"    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$#
        Exit Function
    ElseIf IsNumeric(objEncounter.EncounterID) = False Then
        sReturnStatus = "ENC_ID_FAILED" '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$#
        Exit Function
    ElseIf CLng(objEncounter.EncounterID) = 0 Then
        sReturnStatus = "ENC_ID_ZERO"   '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$#
        Exit Function
    Else
        objAppointment.EncounterNumber = CStr(objEncounter.EncounterID)
        sReturnStatus = "ENC_CREATE_OK" '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$#
        CheckIn_3_CreateEncounter = True
    End If
    
    If (LenB(rsNoteTitle) <> 0 And LenB(rsNoteText) <> 0) Then
        If Not (objEncounter.AddTextNote(rsNoteTitle, rsNoteText)) Then
            MsgBxARMd "Encounter was created but application could not add initial Encounter note." + vbCrLf + vbCrLf + _
                       "NoteTitle: " + rsNoteTitle + vbCrLf + ", NoteText: " + rsNoteText, , "AppointmentOps.Checkin"
        End If
    End If

    Exit Function

ErrHandler:
    sReturnStatus = "ENC_FAILED"
    If Not objAppointment Is Nothing Then
        MsgBxARMd "Error in " & "CHCSII_AppointmentClient.AppointmentOps.CheckIn_3_CreateEncounter: " & Err.Description & vbCrLf + vbCrLf + _
            "Appointment ID = " + objAppointment.ApptId + vbCrLf + _
            "Appointment Time = " + CStr(objAppointment.ApptDateTime) + vbCrLf + _
            "Facility NCID = " + objAppointment.FacilityNCID + vbCrLf + _
            "Patient Name = " + objAppointment.PatientName + vbCrLf + _
            "User ID = " + CStr(mobjUser.UserId)

    Else
        mobjShared.ShowVBError Err.Number, Err.Description, "AppointmentsOps.cls - CheckIn_3_CreateEncounter", App.Title, vbCritical
    End If


End Function
Private Function CheckIn_GEMS_SaveComplaints(objAppointment As CHCSII_AppointmentClient.Appointment, _
                                            sGemsComplaintErrTxt As String) As Boolean

Dim cInjury As Collection
Dim i As Integer
On Error GoTo ErrHandler
'Complaints are saved on to the appt object as a collection of complaints in string format.
'Here we turn the string format in to a collection of BodyImage.Injury objects
'That will be saved to the newly created encounter object.

    If objAppointment Is Nothing Then
        sGemsComplaintErrTxt = "objAppointment is nothing"
        Exit Function
    ElseIf objAppointment.EncounterNumber = vbNullString Then
        sGemsComplaintErrTxt = "No Encounter Number set."
        Exit Function
    ElseIf objAppointment.PendingComplaints Is Nothing Then
        CheckIn_GEMS_SaveComplaints = True
        Exit Function
    ElseIf objAppointment.PendingComplaints.Count = 0 Then
        CheckIn_GEMS_SaveComplaints = True
        Exit Function
    Else
        'GEMS_LoadComplaints with an empty collection (cInjury)
        'Saves the Appt comments on the encounter under "Appointment Note:"
        Set cInjury = GEMS_GetPendingApptComplaints(objAppointment)
        If cInjury Is Nothing Then   'no saved complaints
           Set cInjury = New Collection
        End If
        
        GEMS_LoadComplaints objAppointment.EncounterNumber, cInjury
           
    End If

Set cInjury = Nothing
CheckIn_GEMS_SaveComplaints = True
Exit Function
ErrHandler:

    For i = 1 To objAppointment.PendingComplaints.Count
        If i = 15 Then Exit For
        sGemsComplaintErrTxt = sGemsComplaintErrTxt + Replace(objAppointment.PendingComplaints.Item(i), "###", ", ", , , vbTextCompare) + vbCrLf
    Next
    mobjShared.ShowVBError Err.Number, "Unable to set complaints. " & Err.Description & vbCrLf & _
    sGemsComplaintErrTxt, "AppointmentsOps.cls - CheckIn_3_CreateEncounter", App.Title, vbCritical

End Function
Public Function CheckOut(rsApptid As String, rdDateTime As Date) As Boolean
    mobjSQL.Execute "update appointment set checkout_time = " & oApptConn.SQLDate(rdDateTime) & " where appt_id = " & rsApptid
    CheckOut = True
End Function

Public Function SetEncounterStatus(rsApptid As String, rnStatus As EncounterStatusEnum, rbWorkload As Boolean) As Boolean
    Dim bTelcon As Boolean
    Dim sWorkload As String
    Dim colStmts As Collection
    Dim sApptTransID As String
    
    On Error GoTo ErrHandler
    DoEvents
    If rsApptid = "0" Or IsNumeric(rsApptid) = False Then
        MsgBxARMd "Error in " & "CHCSII_AppointmentClient.AppointmentOps.SetEncounterStatus: " & _
          "Invalid ApptID. ApptID:" & rsApptid & ", Status:" & CStr(rnStatus) & ", Workload:" & CStr(rbWorkload)
        Exit Function
    End If
    
    
    ' as a convenience ot callers of this method, we will connect to the server if necessary
    If mobjSQL Is Nothing Then
        Set mobjSQL = SelectSqlOps()
    End If
    
    If mobjSQL Is Nothing Then
        MsgBxARMd "Error in " & "CHCSII_AppointmentClient.AppointmentOps.SetEncounterStatus: " & _
            "Could not set mobjSQL. ApptID:" & rsApptid & ", Status:" & CStr(rnStatus) & ", Workload:" & CStr(rbWorkload)
        SetEncounterStatus = False
        Exit Function
    End If
    
    mobjSQL.Execute "select appt_classification from appointment where appt_id = " & rsApptid
    DoEvents
    If mobjSQL.EOF = True Then
        MsgBxARMd "Error in " & "CHCSII_AppointmentClient.AppointmentOps.SetEncounterStatus: " & _
            "Could not get Appt_Classification. ApptID:" & rsApptid & ", Status:" & CStr(rnStatus) & ", Workload:" & CStr(rbWorkload)
        SetEncounterStatus = False
        Exit Function
    End If
    
    If mobjSQL("APPT_CLASSIFICATION") = AppointmentClassifications.ApptTelCon Then
        bTelcon = True
        If rbWorkload Then
            sWorkload = ", telcon_workload = 1"
        Else
            sWorkload = ", telcon_workload = 0"
        End If
    Else
        sWorkload = vbNullString
    End If
    Set colStmts = New Collection
    colStmts.Add "update appointment set date_modified = " & SystemDateString & ", encounter_status = " & rnStatus _
        & sWorkload & " where appt_id = " & rsApptid
    If bTelcon And rnStatus = EncCompleted Then
            sApptTransID = mobjSQL.GetNextID(esAPPT_TRANS_ID)
            colStmts.Add AddApptTransaction(sApptTransID, rsApptid, ApptTransTelconComplete)
            DoEvents
    End If
    mobjSQL.ExecuteMulti colStmts
    DoEvents
    GoTo Cleanup
ErrHandler:
    MsgBxARMd "Error in " & "CHCSII_AppointmentClient.AppointmentOps.SetEncounterStatus: " & Err.Description & _
    "  ApptID:" & rsApptid & ", Status:" & CStr(rnStatus) & ", Workload:" & CStr(rbWorkload)
    Set colStmts = Nothing
    SetEncounterStatus = False
    Exit Function
Cleanup:
    Set colStmts = Nothing
    SetEncounterStatus = True
End Function

Public Function TransferAppointmentProvider(robjAppointment As CHCSII_AppointmentClient.Appointment, oNewDTOsProvider As DTOs.Provider) As Boolean
    Dim objEncParent As CHCSIIEncounterCurrent.EncounterParent
    Dim objEncounter As CHCSIIEncounterOps.Encounter
    Dim bLockedEncounter As Boolean
    On Error Resume Next
    
    If Not mobjShared.IsAppMode(modeCHCSI_GUI) Then
        If mobjPatient.UnitNumber <> robjAppointment.PatientUnitNumber Then
            mobjComm.Message cwiSELECT_PATIENT, robjAppointment.PatientUnitNumber & ";" & robjAppointment.PatientIEN, NCID_CW_APPLICATION, mlProgId
        End If
        If LenB(robjAppointment.EncounterNumber) = 0 Then GoTo UpdateAppointmentOnly
        If CLng(robjAppointment.EncounterNumber) < 1 Then GoTo UpdateAppointmentOnly
    
        Set objEncParent = mobjComm.InitializeOLEServer(NCID_ENCOUNTER_APPLICATION)
        If Err Or objEncParent Is Nothing Then
            MsgBxARMd "Could not initialize encounter document object for AppointmentID: " + robjAppointment.ApptId & vbCrLf _
                & "Error: " & Err.Description, vbCritical, "Appointments"
            GoTo Cleanup
        End If
        Set objEncounter = objEncParent.GetEncounter(robjAppointment.PatientUnitNumber, robjAppointment.FacilityNCID, robjAppointment.EncounterNumber)
    Else
        Set objEncounter = GetEncounter_CHCS_GUI(robjAppointment.ApptId, robjAppointment.FacilityNCID, robjAppointment.PatientIEN)
        If objEncounter Is Nothing Then
            TransferAppointmentProvider = True
            GoTo Cleanup
        End If
    End If
    
    
    If objEncounter Is Nothing Then
        MsgBxARMd "Could not get encounter object for AppointmentID: " + robjAppointment.ApptId, vbCritical, "Appointments"
        GoTo Cleanup
    End If
    If Not objEncounter.LockingObject.SectionIsLocked(eHeader, 0) Then
        bLockedEncounter = objEncounter.LockingObject.LockSection(EncounterSectionEnum.eHeader, 0, LockReasonEnum.SectionLock, "Updating Provider")
        If Not bLockedEncounter Then
            GoTo Cleanup
        End If
    End If
    
    
    If LenB(mobjShared.CmdLineSwitches.Item("HIPAA837")) > 0 And Not mobjShared.IsAppMode(modeTheater) Then
      objEncounter.TransferToProvider oNewDTOsProvider.NCID, oNewDTOsProvider.Name
    Else

        'Transfer the primary provider on the encounter object.
        '1st we need to see if (and remove if) the provider to be transfered to is an Additional Provider.
        'SCR-44276, 43459
        'Note:  PrimaryProv=item1 in the provider collection is saved with NCID as an NCID.
        'Note:  Prov#1=item2 in the provider collection is saved with NCID as an IEN.
        'Note:  Prov#2=item3 in the provider collection is saved with NCID as an IEN.
        Dim i As Integer
        'SCR-53278 Only remove the new provider if he is an additional provider; Item(2) or Item(3)
        For i = 2 To objEncounter.Providers.Count
            If objEncounter.Providers.Item(i).NCID = oNewDTOsProvider.NCID Then
                 objEncounter.Providers.Remove (i)
                 Exit For
            End If
        Next
        'Now transfer the primary provider.
        objEncounter.PrimaryProvider.FullName = oNewDTOsProvider.Name
        objEncounter.PrimaryProvider.NCID = oNewDTOsProvider.NCID
        'objEncounter.PrimaryProvider.State = rsModified
    End If
    'Then Save the encounter.
    If Not objEncounter.Save(True) Then
        MsgBxARMd "Could not update encounter database." & vbCrLf & Err.Description, vbCritical, "AppointmentOps"
        GoTo Cleanup
    End If
    
    If mobjShared.IsAppMode(modeCHCSI_GUI) Then
        TransferAppointmentProvider = True
        GoTo Cleanup
    End If
    
UpdateAppointmentOnly:
    'Created new public function so EncounterOps can update appointment provider externally SCR-23893
    If Not UpdateAppointmentProvider(robjAppointment.ApptId, oNewDTOsProvider.NCID) Then
        GoTo Cleanup
    End If
    
    robjAppointment.ClinicianNCID = oNewDTOsProvider.NCID
    TransferAppointmentProvider = True
Cleanup:
    On Error Resume Next
    If bLockedEncounter Then
        Call objEncounter.LockingObject.UnlockSection(EncounterSectionEnum.eHeader, 0)
    End If
    Set objEncParent = Nothing
    Set objEncounter = Nothing
End Function

Public Function TransferAppointmentClinic(rsApptid As String, rsNewClinicNCID As String) As Boolean
    Dim colStmts As Collection
    Dim sApptTransID As String
    
    Set colStmts = New Collection
    colStmts.Add "update appointment set clinic_ncid = " & rsNewClinicNCID & "," _
        & " date_modified = " & SystemDateString _
        & " where appt_id = " & rsApptid
    
    sApptTransID = mobjSQL.GetNextID(esAPPT_TRANS_ID)
    colStmts.Add AddApptTransaction(sApptTransID, rsApptid, ApptTransClinic)
    mobjSQL.ExecuteMulti colStmts
    Set colStmts = Nothing
    TransferAppointmentClinic = True
End Function

Public Function UpdateAppointmentComment(rsApptid As String, rsComment As String) As Boolean
    Dim objAppointment As CHCSII_AppointmentClient.Appointment
    On Error GoTo ErrHandler
    Set objAppointment = GetAppointment(rsApptid)
    If objAppointment Is Nothing Then
        Err.Raise vbObjectError + 1, , "Could not retrieve apptid " & rsApptid
    End If

    mobjSQL.Execute "update appointment set appt_comment = " & oApptConn.SQLQuote(rsComment) & ", " _
        & " date_modified = " & oApptConn.SQLDate(Now) _
        & " where appt_id = " & rsApptid
    UpdateAppointmentComment = True
    GoTo Cleanup
ErrHandler:
    MsgBxARMd "Error in " & "CHCSII_AppointmentClient.AppointmentOps.UpdateAppointmentComment: " & Err.Description
Cleanup:
    Set objAppointment = Nothing
End Function

Public Function UpdateAppointmentClassification(rsApptid As String, reClassification As AppointmentClassifications) As Boolean
    Dim objAppointment As CHCSII_AppointmentClient.Appointment
    On Error GoTo ErrHandler
    Set objAppointment = GetAppointment(rsApptid)
    If objAppointment Is Nothing Then
        Err.Raise vbObjectError + 1, , "Could not retrieve apptid " & rsApptid
    End If
    mobjSQL.Execute "update appointment set appt_classification = " & reClassification & ", " _
        & " date_modified = " & SystemDateString _
        & " where appt_id = " & rsApptid
    UpdateAppointmentClassification = True
    GoTo Cleanup
ErrHandler:
    MsgBxARMd "Error in " & "CHCSII_AppointmentClient.AppointmentOps.UpdateAppointmentClassification: " & Err.Description
Cleanup:
    Set objAppointment = Nothing
End Function

Public Function UpdateAppointmentObservation(rsApptid As String, reClassification As AppointmentClassifications, rbObservation As Boolean) As Boolean
    Dim objAppointment As CHCSII_AppointmentClient.Appointment
    On Error GoTo ErrHandler
    Set objAppointment = GetAppointment(rsApptid)
    If objAppointment Is Nothing Then
        Err.Raise vbObjectError + 1, , "Could not retrieve apptid " & rsApptid
    End If
    
    Call UpdateAppointmentClassification(rsApptid, reClassification)
    
    If rbObservation = True Then
            mobjSQL.Execute "update appointment set observation = 1, " _
                & " date_modified = " & SystemDateString _
                & " where appt_id = " & rsApptid
    End If
    
    UpdateAppointmentObservation = True
    GoTo Cleanup
ErrHandler:
    MsgBxARMd "Error in " & "CHCSII_AppointmentClient.AppointmentOps.UpdateAppointmentObservation: " & Err.Description
Cleanup:
    Set objAppointment = Nothing
End Function

Public Function UpdateAppointmentStatus(rsApptid As String, rsStatus As String, rsCancelReason As String) As Boolean
    Dim objEncParent As CHCSIIEncounterCurrent.EncounterParent
    Dim objEncounter As CHCSIIEncounterOps.Encounter
    Dim objAppointment As CHCSII_AppointmentClient.Appointment
    Dim colStmts As Collection
    Dim bLockedEncounter As Boolean
    Dim bStatus As Boolean
    Dim sApptTransID As String
    On Error Resume Next
    
    'UpdateAppointmentStatus is called from two different locations:
    '   1.) From the frmAppointments - UndoCancel when we want to change the status
    '       back to "PENDING" or "KEPT"
    '
    '   2.) From the dlgCancel - OKbutton when we want to cancel the appointment by setting the
    '       status to CANCEL, LWOBS, or NO-SHOW.
    
    
    'If called from dlgCancel and the rsStatus is CANCEL or NO-SHOW then there will be no
    'encounter for this appointment because the previous status was PENDING. Also, if the rsStatus is
    'CANCEL or NO-SHOW then the objAppointment.EncounterNumber will be "" because an
    'encounter is not created for PENDING appointments.  For these statuses we want to skip
    'the encounter object stuff and just update the appointment tables.
    
    'If rsStatus is CANCEL, LWOBS, or NO-SHOW then we want the message:
    '
    '  MsgBxARMd "The change made to this appointment requires....
    
'SCR-31974 remove EOD message
'    If IsMissing(mbDoNotShowEODmsg) Then
'        bDoNotShowEODmsg = False
'    Else
'        If mbDoNotShowEODmsg = True Then
'            bDoNotShowEODmsg = True
'        Else
'            bDoNotShowEODmsg = False
'        End If
'    End If
    
    
    Set objAppointment = GetAppointment(rsApptid)
    If objAppointment Is Nothing Then
        MsgBxARMd "Could not retrieve appointment from database."
        GoTo Cleanup
    End If
    
    'SCR-26030 S.McAvoy 8/5/02
    If rsStatus <> "NO-SHOW" And rsStatus <> "LWOBS" And rsStatus <> "CANCEL" And _
        rsStatus <> "PENDING" And rsStatus <> "KEPT" Then
        MsgBxARMd "Unknown or invalid appointment status: " + rsStatus + ", sent to UpdateAppointmentStatus in " + _
        "AppointmentOps.cls.  Could not update appointment status."
        UpdateAppointmentStatus = False
        GoTo Cleanup
    End If
    
    'if NO-SHOW or CANCEL then there is no EncounterNumber
    If (LenB(objAppointment.EncounterNumber) <> 0 And objAppointment.EncounterNumber <> "0" And objAppointment.EncounterNumber <> "-1") _
    And (rsStatus = "NO-SHOW" Or rsStatus = "LWOBS" Or rsStatus = "CANCEL") Then 'ONLY LWOBS will have an EncounterNumber?
        Dim eNewEncStatus As ENCOUNTER_INTERFACES.EncStatus
        Set objEncParent = mobjComm.InitializeOLEServer(NCID_ENCOUNTER_APPLICATION)
        If Err Or objEncParent Is Nothing Then
            MsgBxARMd "Could not initialize encounter document object." & vbCrLf _
                & "Error: " & Err.Description, vbCritical, "Appointments"
            GoTo Cleanup
        End If
        Set objEncounter = objEncParent.CurrentEncounter
        If Not objEncounter Is Nothing Then
            If objEncounter.EncounterID <> objAppointment.EncounterNumber _
            Or objEncounter.FacilityNCID <> objAppointment.FacilityNCID Then
                Set objEncounter = Nothing
            End If
        End If
        If objEncounter Is Nothing Then
            Set objEncounter = objEncParent.GetEncounter(objAppointment.PatientUnitNumber, objAppointment.FacilityNCID, objAppointment.EncounterNumber)
        End If
        If objEncounter Is Nothing Then
            MsgBxARMd "Could not get encounter object. Appointment status not updated.", vbCritical, "Appointments"
            GoTo Cleanup
        End If
        If Not objEncounter.LockingObject.SectionIsLocked(eEntireEncounter, 0) Then
            bLockedEncounter = objEncounter.LockingObject.LockSection(EncounterSectionEnum.eEntireEncounter, 0, LockReasonEnum.SectionLock, "LWOBS")
            If Not bLockedEncounter Then
                MsgBxARMd "Could not lock the encounter. Appointment status not updated.", vbCritical, "Appointments"
                GoTo Cleanup
            End If
        End If
        Select Case UCase$(rsStatus)
            Case "CANCEL"       'should never get to this case
                Select Case UCase$(rsCancelReason)
                    Case "PATIENT CANCELLED"
                        eNewEncStatus = CancelByPatient
                    Case "FACILITY CANCELLED"
                        eNewEncStatus = CancelByFacility
                    Case "NO-SHOW"
                        eNewEncStatus = NoShow
                    Case "PROVIDER CANCELLED"
                        eNewEncStatus = CancelByProvider
                    Case "LWOBS"
                        eNewEncStatus = PatientLeftWithoutBeingSeen
                    Case Else
                        MsgBxARMd "Invalid Cancellation Reason Code in UpdateAppointmentStatus. " + vbCrLf + _
                            "Appointment status not set." + vbCrLf + _
                            "ApptID: " + CStr(rsApptid) + "," + vbCrLf + _
                            "rsStatus: " + rsStatus + "," + vbCrLf + _
                            "rsCancelReason: " + rsCancelReason + vbCrLf + "Appointment status not updated."
                        GoTo Cleanup
                End Select
            Case "NO-SHOW"
                eNewEncStatus = NoShow
            Case "KEPT"  'SCR-31946
                If rsCancelReason = "UNDOCANCEL" Then
                    bStatus = objEncounter.UndoCancel
                    GoTo SkipCancel
                End If
            Case Else
                eNewEncStatus = PatientLeftWithoutBeingSeen
        End Select
        bStatus = objEncounter.Cancel(eNewEncStatus)
SkipCancel:
        If Err Then
            MsgBxARMd "Error from encounter module: " & Err.Description & ". Appointment status not updated."
            GoTo Cleanup
        ElseIf Not bStatus Then
            ' message is presented by encounter
            GoTo Cleanup
        End If
    End If
    
    
    Set colStmts = New Collection
    
    'SCR-34717  Checked in on chcsi
    If rsStatus = "PENDING" And rsCancelReason = "UNDOCANCEL" And LenB(objAppointment.EncounterNumber) = 0 And _
        LenB(objAppointment.EncounterStatus) = 0 And (objAppointment.Appt_Status_CHCS = "KEPT" Or _
        objAppointment.Appt_Status_CHCS = "WALK-IN" Or objAppointment.Appt_Status_CHCS = "S-CALL") Then
        colStmts.Add "update appointment set appt_status = 'KEPT', " _
            & " appt_cancel_reason = ''," _
            & " date_modified = " & SystemDateString _
            & " where appt_id = " & rsApptid
    
    'SCR-31946
    ElseIf (rsStatus = "KEPT" Or rsStatus = "PENDING") And rsCancelReason = "UNDOCANCEL" Then
        colStmts.Add "update appointment set appt_status = '" & rsStatus & "', " _
            & " appt_cancel_reason = ''," _
            & " date_modified = " & SystemDateString _
            & " where appt_id = " & rsApptid
    
    Else
        If objAppointment.EncounterNumber = "-1" Then
            colStmts.Add "update appointment set appt_status = '" & rsStatus & "', " _
                & " appt_cancel_reason = '" & rsCancelReason & "'," _
                & " encounter_number = Null, " _
                & " date_modified = " & SystemDateString _
                & " where appt_id = " & rsApptid
        
        Else
            colStmts.Add "update appointment set appt_status = '" & rsStatus & "', " _
                & " appt_cancel_reason = '" & rsCancelReason & "'," _
                & " date_modified = " & SystemDateString _
                & " where appt_id = " & rsApptid
        End If
    End If
    Select Case rsStatus
        Case "CANCEL", "NO-SHOW"
            sApptTransID = mobjSQL.GetNextID(esAPPT_TRANS_ID)
            colStmts.Add AddApptTransaction(sApptTransID, rsApptid, ApptTransCancel)
        Case "LWOBS"
            sApptTransID = mobjSQL.GetNextID(esAPPT_TRANS_ID)
            colStmts.Add AddApptTransaction(sApptTransID, rsApptid, ApptTransLWOBS)
        Case "KEPT", "PENDING"
            If rsCancelReason = "UNDOCANCEL" Then
                sApptTransID = mobjSQL.GetNextID(esAPPT_TRANS_ID)
                colStmts.Add AddApptTransaction(sApptTransID, rsApptid, ApptTransUndoCancel)
            End If
    End Select
    mobjSQL.ExecuteMulti colStmts
        
    UpdateAppointmentStatus = True
    
Cleanup:
    On Error Resume Next
    If bLockedEncounter Then
        Call objEncounter.LockingObject.UnlockSection(EncounterSectionEnum.eEntireEncounter, 0)
    End If
    Set objAppointment = Nothing
    Set objEncParent = Nothing
    Set objEncounter = Nothing
    Set colStmts = Nothing
End Function

'SCR 24097  Phil Crowder 6/4/02
Public Function UpdateAppointmentRFV(ByVal ApptId As String, ByVal RFV As String) As Boolean
Dim objAppointment As CHCSII_AppointmentClient.Appointment
Dim lStrLen As Long
Dim ApptReasonFV As String

    On Error GoTo ErrHandler
    
    Set objAppointment = GetAppointment(ApptId)
    If objAppointment Is Nothing Then
        MsgBxARMd "Error in " & "CHCSII_AppointmentClient.AppointmentOps.UpdateAppointmentRFV: Could not retrieve Appointment with ApptID: " & ApptId & "  Reason for Visit will not be updated."
        Exit Function
    End If
    
    'SCR 105643 - LCI - 09/14/2006: Do not update RFV is one is already present.
    mobjSQL.Execute "select appt_reason " _
          & " from appointment" _
          & " where appt_id = " & ApptId
    
    ApptReasonFV = mobjSQL("appt_reason")
    If ApptReasonFV <> "" Then
        Exit Function
    End If
    
    'SCR-50577
    If Trim$(RFV) = vbNullString Then
      mobjSQL.Execute "update appointment set appt_reason = Null, " _
          & " date_modified = " & SystemDateString _
          & " where appt_id = " & ApptId
    
    Else
    If mobjShared.IsAppMode(modeTheater) Then
        lStrLen = 3995
    Else
        lStrLen = 80
    End If
      mobjSQL.Execute "update appointment set appt_reason = " & oApptConn.SQLQuote(RFV, lStrLen) & ", " _
          & " date_modified = " & SystemDateString _
          & " where appt_id = " & ApptId
    End If
    
    UpdateAppointmentRFV = True
    Set objAppointment = Nothing

    Exit Function
    
ErrHandler:
    Set objAppointment = Nothing
    MsgBxARMd "Error in " & "CHCSII_AppointmentClient.AppointmentOps.UpdateAppointmentRFV: " & Err.Description

End Function

Public Function UpdateADMStatus_CHCS_GUI(sAppt_IEN As String, sADMstatus As String) As Boolean
Dim objAppt As CHCSII_AppointmentClient.Appointment
Dim bAllOk As Boolean
Dim i As Integer
Dim eNewEncStatus As CHCSII_AppointmentStatus.EncounterStatusEnum

On Error GoTo ErrHandler


    If Not mobjShared.IsAppMode(modeCHCSI_GUI) Then GoTo Cleanup
    If LenB(sAppt_IEN) = 0 Then Exit Function
    If IsNumeric(sAppt_IEN) = False Then GoTo Cleanup
    If mcolAppointmentsCache Is Nothing Then GoTo Cleanup
    If mcolAppointmentsCache.Count = 0 Then GoTo Cleanup
    
    For i = 1 To mcolAppointmentsCache.Count
        If mcolAppointmentsCache.Item(i).ApptIEN = sAppt_IEN Then
            Set objAppt = mcolAppointmentsCache.Item(i)
            
            
            If IsNumeric(sADMstatus) = True Then
                eNewEncStatus = CInt(sADMstatus)
                sADMstatus = TranslateEncStatus(eNewEncStatus)
            End If
            
            If UCase$(sADMstatus) = "COMPLETE" Then
                objAppt.ADM_Encounter = "Complete"
            Else
                objAppt.ADM_Encounter = vbNullString 'sADMstatus 'SCR-40435
            End If
            mcolAppointmentsCache.Remove (i)
            mcolAppointmentsCache.Add objAppt, sAppt_IEN
            bAllOk = True
            Exit For
        End If
    Next
Cleanup:
    If bAllOk = True Then
        UpdateADMStatus_CHCS_GUI = True
    Else
        UpdateADMStatus_CHCS_GUI = False
    End If
    Set objAppt = Nothing
    Exit Function

ErrHandler:
    Set objAppt = Nothing
    UpdateADMStatus_CHCS_GUI = False
    MsgBxARMd "Error in " & "CHCSII_AppointmentClient.AppointmentOps.UpdateADMStatus_CHCS_GUI: " & Err.Description

End Function

'UpdateAppointmentWorkload
Public Function UpdateAppointmentWorkload(ByVal ApptId As String, ByVal WrkLoadType As String) As Boolean
Dim objAppointment As CHCSII_AppointmentClient.Appointment
Dim oCHCSInterface_Appt As ICHCSAppointments
    
    On Error GoTo ErrHandler
    
    Set objAppointment = GetAppointment(ApptId, True)
    If objAppointment Is Nothing Then
        MsgBxARMd "Error in " & "CHCSII_AppointmentClient.AppointmentOps.UpdateAppointmentWorkload: Could not retrieve Appointment with ApptID: " & ApptId & ".  Appt Workload indicator will not be updated."
        GoTo Cleanup
    End If
    WrkLoadType = UCase$(WrkLoadType)
    If LenB(WrkLoadType) = 0 Or (WrkLoadType <> "N" And WrkLoadType <> "C") Then
        MsgBxARMd "Error in " & "CHCSII_AppointmentClient.AppointmentOps.UpdateAppointmentWorkload: Invalid workload type. " & ApptId & ".  Appt Workload indicator will not be updated."
        GoTo Cleanup
    End If
    With objAppointment
        If .Workload_Type = WrkLoadType Then  'no need to update record.
            UpdateAppointmentWorkload = True
            GoTo Cleanup
        End If
        If mobjShared.IsAppMode(modeCHCSI_GUI) Then  'SCR-40847
            Set oCHCSInterface_Appt = mobjShared.CHCSConnection
                oCHCSInterface_Appt.AppointmentUpdate .ApptIEN, .USV_TYPE, .ApptType, .ApptStatus, .CheckInTime, WrkLoadType, .APPT_MEPRS_IEN, .ApptClassEnum
        Else
        mobjSQL.Execute "update appointment set WORKLOAD_TYPE = '" & WrkLoadType & "', " _
              & " date_modified = " & SystemDateString _
              & " where appt_id = " & ApptId
              
        End If
    
        Set objAppointment = GetAppointment(ApptId, False)
        If WrkLoadType <> objAppointment.Workload_Type Then
            MsgBxARMd "Unable to update the work load data in appointments for ApptID: " & ApptId & "."
            GoTo Cleanup
        End If
    End With
    UpdateAppointmentWorkload = True

Cleanup:
    
    Set objAppointment = Nothing
    Set oCHCSInterface_Appt = Nothing

    Exit Function
    
ErrHandler:
    If LenB(Err.Description) <> 0 Then
        If InStr(1, Err.Description, "Unable to lock Appointment", vbTextCompare) > 0 Then
            UpdateAppointmentWorkload = True
            Err.Clear
            GoTo Cleanup
        End If
    End If
    Set objAppointment = Nothing
    MsgBxARMd "Error in " & "CHCSII_AppointmentClient.AppointmentOps.UpdateAppointmentWorkload: " & Err.Description
    GoTo Cleanup
    
    Resume
End Function
Public Function GetPatientIEN(rsMTFFaciltiyNCID As String, rsPatientUnitNumber As String) As String
    On Error GoTo ExitFunction
    
    
    
    mobjSQL.Execute "select medical_record_number from mmi_site" _
        & " where facility_cid = " & rsMTFFaciltiyNCID _
        & "    and unit_number = " & rsPatientUnitNumber
    GetPatientIEN = mobjSQL("MEDICAL_RECORD_NUMBER")
ExitFunction:
End Function

Private Sub Class_Initialize()
    
    On Error GoTo ErrHnd
    
    Set mobjShared = New CWShared
    
    If mobjShared.UseMySQLDB Then
        geDBUsed = edb_MySQL
    ElseIf mobjShared.UseSQLServer Then
      geDBUsed = edb_sqlserver
    Else
      geDBUsed = edb_Oracle
    End If
Exit Sub

ErrHnd:
    MsgBox "Application Error: " & Err.Description & " Occurred in CHCSII_AppointmentClient.AppointmentOps.Class_Initialize."
    
End Sub

Private Sub Class_Terminate()
On Error GoTo ErrHandler
    Set mobjShared = Nothing
    Set mcClinicianNCIDs = Nothing
    Set mcClinicNCIDs = Nothing
    'Set mobjSQL = Nothing
    Set mobjComm = Nothing
    Set mobjPatient = Nothing
    Set mobjLogon = Nothing
    Set mobjConfig = Nothing
    Set mobjUser = Nothing
    Set mobjEncParent = Nothing
    'Set mcolAppointments_CHCS_GUI = Nothing
    'SCR-32352
    Set mcolAppointmentsCache = Nothing
    Set mcIBWAClinicNCIDs = Nothing 'SCR-RNDS
    Set mcIBWAClinicianNCID = Nothing
    Exit Sub
ErrHandler:
 MsgBxARMd "Error in " & "CHCSII_AppointmentClient.AppointmentOps.Class_Terminate: " & Err.Description

End Sub


' Function also called by EncountersOps when transfering Encounters during Sign SCR-23893
Public Function UpdateAppointmentProvider(ByVal sAppointmentID As String, sNewProviderNCID As String) As Boolean

    Dim colStmts As Collection
    Dim sApptTransID As String
    
    On Error Resume Next
    
    UpdateAppointmentProvider = True
    
    Set colStmts = New Collection
    colStmts.Add "update appointment set clinician_ncid = " & sNewProviderNCID & "," _
        & " date_modified = " & SystemDateString _
        & " where appt_id = " & sAppointmentID
    sApptTransID = mobjSQL.GetNextID(esAPPT_TRANS_ID)
    colStmts.Add AddApptTransaction(sApptTransID, sAppointmentID, ApptTransProvider)
    mobjSQL.ExecuteMulti colStmts
    If Err Then
        MsgBxARMd "Error: could not update appointment database." & vbCrLf & Err.Description, vbCritical, "AppointmentOps"
        UpdateAppointmentProvider = False
    End If

    Set colStmts = Nothing
End Function
Public Function GetEncounter_CHCS_GUI(sApptID As String, sFacilityNCID As String, sPatientIEN As String) As CHCSIIEncounterOps.Encounter
    
    Dim objEncParent As CHCSIIEncounterCurrent.EncounterParent
    Dim objEncOps As CHCSIIEncounterOps.EncounterOps
    Dim objEncounter As CHCSIIEncounterOps.Encounter
    Dim sSql As String
    Dim sEncNum As String
    Dim sPatientUnitNumber As String
    
    On Error GoTo ErrHandler
    
    If Not mobjShared.IsAppMode(modeCHCSI_GUI) Then GoTo Cleanup
    If LenB(sApptID) = 0 Or LenB(sFacilityNCID) = 0 Then GoTo Cleanup
    sSql = "SELECT ENCOUNTERNUMBER, UNIT_NUMBER FROM ENCOUNTERS WHERE APPTID = '" + sApptID + "' AND FACILITYNCID = '" + sFacilityNCID + "'"
    mobjSQL.Execute sSql
    If mobjSQL.BOF And mobjSQL.EOF Then
        Set GetEncounter_CHCS_GUI = Nothing
        GoTo Cleanup
    Else
        sEncNum = mobjSQL("ENCOUNTERNUMBER") & vbNullString
        If IsNumeric(sEncNum) = False Then
            Set GetEncounter_CHCS_GUI = Nothing
            GoTo Cleanup
        End If
        
        sPatientUnitNumber = mobjSQL("UNIT_NUMBER")
        If LenB(sPatientUnitNumber) = 0 Or sPatientUnitNumber = "0" Then
            If IsNumeric(sPatientIEN) = False Then
                Set GetEncounter_CHCS_GUI = Nothing
                GoTo Cleanup
            End If
            sPatientUnitNumber = RetrievePatientUnitNumber_CHCS_GUI(sPatientIEN, sApptID)
        End If
        If IsNumeric(sPatientUnitNumber) = False Then
            Set GetEncounter_CHCS_GUI = Nothing
            GoTo Cleanup
        End If
        
    End If
    
    Set objEncParent = mobjComm.InitializeOLEServer(NCID_ENCOUNTER_APPLICATION)
    Set objEncOps = objEncParent.EncounterOps
    Set objEncounter = objEncOps.GetEncounter(sPatientUnitNumber, sFacilityNCID, sEncNum)
    If Not objEncounter Is Nothing Then
        Set GetEncounter_CHCS_GUI = objEncounter
    Else
        Set GetEncounter_CHCS_GUI = Nothing
        GoTo Cleanup
    End If

Cleanup:
    Set objEncParent = Nothing
    Set objEncOps = Nothing
    Set objEncounter = Nothing
 Exit Function
ErrHandler:
 mobjShared.ShowVBError Err.Number, Err.Description, "AppointmentsOps.cls - GetEncounter_CHCS_GUI", App.Title, vbCritical
 GoTo Cleanup
End Function
Public Function UpdateAppointmentMEPRS(ByVal sAppointmentID As String, ByVal bRelatedToInpatient As Boolean) As Boolean

    Dim objAppointment As CHCSII_AppointmentClient.Appointment
    Dim sRelatedToInpatient As String
    Dim sSql As String
    Dim oCHCSInterface_Appt As CHCS_Interface.ICHCSAppointments
    
    'SCR-34769
    On Error GoTo ErrHandler
    
    If mobjShared.IsAppMode(modeCHCSI_GUI) Then '##########################################################################################
        If Not UpdateAppointmentMEPRS_CHCS_GUI(sAppointmentID, bRelatedToInpatient) Then
            UpdateAppointmentMEPRS = False
            GoTo Cleanup
        End If
        
    Else  '##########################################################################################
    
        If mobjSQL Is Nothing Then
            Set mobjSQL = SelectSqlOps()
        End If
        
        If bRelatedToInpatient = True Then  '$$$$$$$$$$$$$$$$$$$$$$$$$$$$
            sRelatedToInpatient = "TRUE"
            
            sSql = "update appointment set RELATED_TO_INPATIENT = '" & sRelatedToInpatient & "'," _
            & " date_modified = " & SystemDateString & ", Appt_MEPRS_IEN = INPATIENT_MEPRS_IEN " _
            & " where appt_id = " & sAppointmentID
            
            'Set Apt_Meprs_IEN = INPATIENT_MEPRS_IEN
            
        Else
            sRelatedToInpatient = "FALSE"
            Set objAppointment = GetAppointment(sAppointmentID)
            sSql = "update appointment set RELATED_TO_INPATIENT = '" & sRelatedToInpatient & "'," _
            & " date_modified = " & SystemDateString & ", Appt_MEPRS_IEN = (Select MEPRS_IEN from Clinic c Where c.ncid = " + objAppointment.ClinicNCID + " and c.facility_NCID = " + objAppointment.FacilityNCID + ") " _
            & " where appt_id = " & sAppointmentID
            
        End If  '$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        
        On Error Resume Next
        mobjSQL.Execute sSql
        
        If Err Then
            MsgBxARMd "Error: could not update appointment database." & vbCrLf & Err.Description, vbCritical, "AppointmentOps"
            UpdateAppointmentMEPRS = False
            Err.Clear
            GoTo Cleanup
        End If
        Set objAppointment = GetAppointment(sAppointmentID, False) 'To refresh cache with new value.
    End If
    UpdateAppointmentMEPRS = True
Cleanup:
    Set objAppointment = Nothing
    Set oCHCSInterface_Appt = Nothing
 Exit Function
ErrHandler:
 mobjShared.ShowVBError Err.Number, Err.Description, "AppointmentsOps.cls - UpdateAppointmentMEPRS", App.Title, vbCritical
 GoTo Cleanup
End Function
Public Function UpdateAppointmentMEPRS_CHCS_GUI(ByVal sAppointmentID As String, ByVal bRelatedToInpatient As Boolean) As Boolean

    Dim objAppointment As CHCSII_AppointmentClient.Appointment
    Dim sSql As String
    Dim oCHCSInterface_Appt As CHCS_Interface.ICHCSAppointments
    
    'SCR-34769
    Dim sPatientIEN As String
    Dim bIsCurrentInpatient As Boolean
    Dim sInpatientLabel As String
    Dim sInpatient_MEPRS_Code As String
    Dim sInpatient_MEPRS_IEN As String
    Dim sInpatient_MEPRS_Description As String
    
    On Error GoTo ErrHandler
    
    If Not mobjShared.IsAppMode(modeCHCSI_GUI) Then '##########################################################################################
        GoTo Cleanup
    End If
    
    
    Set objAppointment = GetAppointment_CHCS_GUI(sAppointmentID)
    
    'If relatedToInpatientStay = True then simply set the appt.MEPRS_Code to the "A" level
    'This is found from RetrieveInpatientStatus_CHCS_GUI
    If bRelatedToInpatient = True Then
        sPatientIEN = objAppointment.PatientIEN
        If LenB(sPatientIEN) = 0 Then
            MsgBxARMd "Could not update appoinment MEPRS code. Invalid Patient IEN found.", , "AppointmentOps.UpdateAppointmentMEPRS_CHCS_GUI"
            UpdateAppointmentMEPRS_CHCS_GUI = False
            GoTo Cleanup
        End If
        'SCR-34769
        If Not RetrieveInpatientStatus_CHCS_GUI(sPatientIEN, bIsCurrentInpatient, _
            sInpatientLabel, sInpatient_MEPRS_Code, sInpatient_MEPRS_IEN, _
            sInpatient_MEPRS_Description) Then
            MsgBxARMd "Could not retrieve Inpatient MEPRS code.", , "AppointmentOps.UpdateAppointmentMEPRS_CHCS_GUI"
            UpdateAppointmentMEPRS_CHCS_GUI = False
            GoTo Cleanup
        End If
    
    Else
    'If relatedToInpatientStay = False then simply set the appt.MEPRS_Code to the non"A" level
    'This is found from sSql = "Select m.Code, m.IEN from MEPRS_Code m, Clinic c " + _

    
    If mobjShared.UseSQLServer = True Then  'IsAppMode(modeITT )
        sSql = "Select m.Code, m.IEN from MEPRS_Code m, Clinic c " + _
                "Where c.ien = " + objAppointment.ClinicNCID + " and " + _
                "c.facility_NCID = " + objAppointment.FacilityNCID + " and " + _
                "c.facility_NCID = m.facility_NCID AND " + _
                "c.MEPRS_Ien = m.IEN"
    
    Else
        sSql = "Select m.Code, m.IEN from MEPRS_Code m, Clinic c " + _
                "Where c.ien = " + objAppointment.ClinicNCID + " and " + _
                "c.facility_NCID = " + objAppointment.FacilityNCID + " and " + _
                "c.facility_NCID = m.facility_NCID (+) AND " + _
                "c.MEPRS_Ien = m.IEN"
    End If

        mobjSQL.Execute sSql
        If mobjSQL.BOF And mobjSQL.EOF Then
            MsgBxARMd "Could not retrieve Clinic MEPRS code.", , "AppointmentOps.UpdateAppointmentMEPRS_CHCS_GUI"
            UpdateAppointmentMEPRS_CHCS_GUI = False
            GoTo Cleanup
        Else
            sInpatient_MEPRS_IEN = mobjSQL("IEN") & vbNullString
            sInpatient_MEPRS_Code = mobjSQL("CODE") & vbNullString
        End If
        If LenB(sInpatient_MEPRS_IEN) = 0 Then
            MsgBxARMd "Could not retrieve Clinic MEPRS CODE.", , "AppointmentOps.UpdateAppointmentMEPRS_CHCS_GUI"
            UpdateAppointmentMEPRS_CHCS_GUI = False
            GoTo Cleanup
        End If
    
    End If
    
    Set oCHCSInterface_Appt = mobjShared.CHCSConnection
    With objAppointment
        oCHCSInterface_Appt.AppointmentUpdate sAppointmentID, .USV_TYPE, .ApptType, _
            .ApptStatus, Now, .Workload_Type, sInpatient_MEPRS_IEN, .ApptClassEnum
    End With
    Set objAppointment = GetAppointment_CHCS_GUI(sAppointmentID, False)
    
    UpdateAppointmentMEPRS_CHCS_GUI = True
Cleanup:
    Set objAppointment = Nothing
    Set oCHCSInterface_Appt = Nothing
 Exit Function
ErrHandler:
 mobjShared.ShowVBError Err.Number, Err.Description, "AppointmentsOps.cls - UpdateAppointmentMEPRS_CHCS_GUI", App.Title, vbCritical
 GoTo Cleanup
End Function

'''Public Function Validate_Appt_Enc_Status(sApptID As String, sEncNum As String, _
'''        enumEnc_Status As ENCOUNTER_INTERFACES.EncStatus, enumApptEnc_Status As CHCSII_AppointmentStatus.EncounterStatusEnum, sReturnMsg As String) As Boolean
''''SCR-36381, 42360
'''Dim bOutofSync As Boolean
'''On Error GoTo ErrHandler
'''    sReturnMsg = ""
'''    If sApptID = "" Or sEncNum = "" Or CInt(enumEnc_Status) < 1 Or CInt(enumApptEnc_Status) < 1 Then
'''        sReturnMsg = "Unable to validate appointment status to encounter status. Invalid arguments"
'''        Exit Function
'''    End If
'''
'''    Select Case enumEnc_Status
'''
'''        Case CancelByFacility, PatientLeftWithoutBeingSeen, CancelByPatient, NoShow
'''            If enumApptEnc_Status <> EncCancel Then bOutofSync = True
'''
'''        Case CheckedIn
'''            If enumApptEnc_Status <> EncCheckedIn Then bOutofSync = True
'''
'''        Case complete
'''            If enumApptEnc_Status <> EncCompleted Then bOutofSync = True
'''
'''        Case InProgress
'''            If enumApptEnc_Status <> EncInProgress Then bOutofSync = True
'''
'''        Case NeedsCoSignature
'''            If enumApptEnc_Status <> EncNeedsCosignature Then bOutofSync = True
'''
'''        Case Updated
'''            If enumApptEnc_Status <> EncUpdated Then bOutofSync = True
'''
'''        Case Updating
'''            If enumApptEnc_Status <> EncUpdating Then bOutofSync = True
'''
'''        Case Waiting
'''            If enumApptEnc_Status <> EncWaiting Then bOutofSync = True
'''
'''        Case NeedsSignature, NoSave, Pending, Incomplete
'''            'No equivilent appt status
'''            bOutofSync = True
'''
'''        Case Else
'''            bOutofSync = True
'''            sReturnMsg = "Unknown Encounter Status: " + CStr(enumEnc_Status) + ". "
'''
'''    End Select
'''    If bOutofSync = True Then
'''        Validate_Appt_Enc_Status = False
''''        sReturnMsg = sReturnMsg + "The appointment status and the encounter status are out of sync for ApptID/EncID " + sApptID + "/" + sEncNum + "." + _
''''                    vbCrLf + vbCrLf + "Appt Status = " + TranslateApptStatus(enumApptEnc_Status) + ", Encounter Status = " + TranslateEncounterStatus(enumEnc_Status) + "."
'''    Else
'''        Validate_Appt_Enc_Status = True
'''    End If
'''
''' Exit Function
'''ErrHandler:
''' mobjShared.ShowVBError Err.Number, Err.Description, "AppointmentsOps.cls - Validate_Appt_Enc_Status", App.Title, vbCritical
'''
'''End Function

'**********************************************************************************
'**********************************************************************************
'******* Functions Added as a part of Theater to CurrentBuild SCR-42772 **********
'**********************************************************************************
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Public Function GetApptTypes() As ICHCSII_SQL.ISqlOps
Dim sql As String
    On Error GoTo ErrHandler
        If mobjShared.AppMode <> modeTheater Then Exit Function
    sql = "select DESCRIPTION, CODE from appt_type order by Description "
          '"facility_ncid  = " + CStr(mobjLogon.FacilityNCID) + " order by Description"
    If mobjSQL Is Nothing Then
        Set mobjSQL = SelectSqlOps()
    End If
    mobjSQL.Execute sql
    Set GetApptTypes = mobjSQL
Exit Function
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "AppointmentsOps.cls - GetApptTypes", App.Title, vbCritical
End Function





Public Function GEMS_WorkStatus_UpdateIsQualified(pNCID As String, _
                                              pChecked As Boolean, _
                                              pQualified As String) _
                                              As Boolean
    
    Dim sSql            As String
    Dim sWhere          As String
    Dim sChecked As String
    Dim oEncDoc         As Encounter
    
    Dim gobjEncParent As ENCOUNTER_INTERFACES.IParent ' - the 'parent node'  Use it to do things like shut down this child

        
    ' 7/08/2002
    'WR IParent.CurrentEncounter() now returns IEncDoc interface SCR 25225
    Set oEncDoc = gobjEncParent.CurrentEncounter
    
    On Error Resume Next
    
    GEMS_WorkStatus_UpdateIsQualified = False
    If mobjSQL Is Nothing Then
        Set mobjSQL = SelectSqlOps()
    End If
    
    sSql = "SELECT WORK_STATUS_QUALIFIED FROM WORK_STATUS "
    
    With oEncDoc
         sWhere = "WHERE PATIENT_NCID = " & cwobjPatient.UnitNumber & " " & _
                "  AND ENCOUNTER_NUMBER = '" & .EncounterID & "' " & _
                "  AND FACILITY_ID = " & .FacilityNCID & " " & _
                "  AND WORK_STATUS_NCID = " & pNCID & " "
    End With
    
    sSql = sSql & sWhere
    
    mobjSQL.Execute sSql
    If Err Then
       GoTo LookupError
    Else
        With mobjSQL
            If pChecked Then sChecked = "-1" Else sChecked = 0
            
            If (.BOF And .EOF) Then
                With oEncDoc
                    
                    sSql = "INSERT INTO WORK_STATUS ( " & vbLf _
                        & "  PATIENT_NCID, ENCOUNTER_NUMBER, WORK_STATUS_NCID, " _
                        & "  WORK_STATUS_SELECTED, WORK_STATUS_QUALIFIED, " _
                        & "  FACILITY_ID, PROVIDER_NCID) " _
                        & "VALUES(" & CStr(cwobjPatient.UnitNumber) & ", '" _
                        & CStr(.EncounterID) & "', " _
                        & pNCID & ", " & sChecked & ", " _
                        & pQualified & ", " & .FacilityNCID & ", " _
                        & .PrimaryProvider.NCID & ")"
                End With 'oEncDoc
            Else
                '.MoveFirst     'Do not need to do this.
       
                sSql = "UPDATE WORK_STATUS " & vbLf _
                     & "SET WORK_STATUS_SELECTED = " & sChecked & ", " _
                     & "    WORK_STATUS_QUALIFIED = " & pQualified & ", " _
                     & "    PROVIDER_NCID = " & oEncDoc.PrimaryProvider.NCID & " "
        
                If Len(sWhere$) Then
                    sSql = sSql & sWhere$
                End If
            End If
        End With 'objRecordset
    End If
    
    mobjSQL.Execute sSql
    
    If Err Then
       GoTo CleanUpAndExit 'Should goto ErrHandler.
    End If
    
    GEMS_WorkStatus_UpdateIsQualified = True
    GoTo CleanUpAndExit
    
LookupError: '<--Replace with standard ErrHandler (below).
    
    'Replace with call to standard error logging routine:
    With Err
        If Len(sSql) Then
            .Description = .Description & vbLf & "SQL = " & sSql
        End If
             
        MsgBxARMd "Error #: " & .Number & " has occurred." & vbCrLf & "Description: " & .Description
             
    End With
    GoTo CleanUpAndExit
    
ErrHandler:
    
    Screen.MousePointer = vbDefault
    
    'CALL THE GENERIC ERROR HANDLER:
    With Err
        If Len(sSql) Then
            .Description = .Description & vbLf & "SQL = " & sSql
        End If
         mobjShared.ShowVBError Err.Number, Err.Description, "AppointmentsOps.cls - GEMS_WorkStatus_UpdateIsQualified", App.Title, vbCritical

    End With 'Err
    
CleanUpAndExit:
    'Set objRecordset = Nothing

End Function 'GEMS_WorkStatus_UpdateIsQualified
Private Function GEMS_BuildComplaintsFromString(colComplaints As Collection) As Collection
Dim ComplaintRow As Variant
Dim ComplaintItem As Variant
Dim j As Integer
Dim oInj As Object

    On Error GoTo ErrHandler

        If colComplaints Is Nothing Then GoTo Cleanup
        If colComplaints.Count = 0 Then GoTo Cleanup
        Set GEMS_BuildComplaintsFromString = New Collection
        For j = 1 To colComplaints.Count
            ComplaintRow = colComplaints.Item(j)
            Set oInj = CreateObject("CHCSII_BodyImage.cInjury")
            ComplaintItem = Split(ComplaintRow, "###", , vbTextCompare)
            If UBound(ComplaintItem) >= 0 Then _
                oInj.Sno.SnoID = ComplaintItem(0)
            If UBound(ComplaintItem) >= 1 Then _
                oInj.Sno.Desc = ComplaintItem(1)
            If UBound(ComplaintItem) >= 2 Then _
                oInj.RFVcomment = ComplaintItem(2)
            If UBound(ComplaintItem) >= 3 Then _
                oInj.NewFollowup = ComplaintItem(3)
            GEMS_BuildComplaintsFromString.Add oInj
        Next
    GoTo Cleanup
ErrHandler:
    Set GEMS_BuildComplaintsFromString = Nothing
    Screen.MousePointer = vbDefault
    Call mobjShared.ShowVBError(Err.Number, Err.Description, "AppointmentOps.GEMS_BuildComplaintsFromString", "CHCSII_AppointmentClient", vbExclamation)
Cleanup:
    Set oInj = Nothing

End Function

Private Function GEMS_GetPendingApptComplaints(objAppointment As CHCSII_AppointmentClient.Appointment) As Collection
    On Error GoTo ErrHandler

    'SCR-51307 - no need to query db... get values from appt obj.
        If objAppointment Is Nothing Then Exit Function
        With objAppointment
            If .PendingComplaints Is Nothing Then Exit Function
            If .PendingComplaints.Count = 0 Then Exit Function
            Set GEMS_GetPendingApptComplaints = GEMS_BuildComplaintsFromString(.PendingComplaints)
        End With
        Exit Function
ErrHandler:
    Set GEMS_GetPendingApptComplaints = Nothing
    Call mobjShared.ShowVBError(Err.Number, Err.Description, "AppointmentOps.GEMS_GetPendingApptComplaints", "CHCSII_AppointmentClient", vbExclamation)
End Function
Public Function GEMS_DeleteSavedPendingComplaints(sApptID As String) As Boolean
 On Error GoTo ErrHandler
    If Not mobjShared.IsAppMode(modeITT) Then Exit Function
    If IsNumeric(sApptID) = False Then Exit Function
    'Remove cached data from appt_Pending_Complaints table
    'We will only have records in this table if this was a
    'future/scheduled appt.
    mobjSQL.Execute "Delete APPT_Pending_Complaints WHERE APPT_ID = " + sApptID
    GEMS_DeleteSavedPendingComplaints = True
    Exit Function
ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description, "AppointmentOps.GEMS_DeleteSavedPendingComplaints", "CHCSII_AppointmentClient", vbExclamation)
End Function

Public Function GEMS_LoadComplaints(sEncounterNumber As String, cInjury As Collection) As Boolean
    Dim DataID          As Long
    Dim sRTF            As String
    Dim aBytes()        As Byte
    Dim objZlib As CHCSII_Zlib
    Dim oInj            As Object
    Dim oSQL            As ICHCSII_SQL.ISqlOpsEx
    Dim oRFVComplaints  As MMMHISReasonForVisit.Complaints
    Dim oRFV            As MMMHISReasonForVisit.ReasonForVisit
    Dim encVisitTyp As enVisitType
    Dim oDatabase           As New CHCSII_CONN.Conn
    Dim sSql As String

    On Error GoTo ErrHandler
  
    Set oSQL = oDatabase.CHCSII_SQLOPS(Auto)
    Set objZlib = New CHCSII_Zlib
    Set oInj = CreateObject("CHCSII_BodyImage.cInjury")
    Set oRFV = mobjComm.InitializeOLEServer(NCID_REASON_FOR_VISIT)
    If oRFV Is Nothing Then
        MsgBxARMd "Unable to create/save Reason for Visit data to encounter: " + sEncounterNumber, , "AppointmentOps - GEMS_LoadComplaints"
        GoTo Cleanup
    End If
    Set oRFVComplaints = oRFV.getComplaintsClass
    If oRFVComplaints Is Nothing Then
        MsgBxARMd "Unable to create/save Reason for Visit (complaints) data to encounter: " + sEncounterNumber, , "AppointmentOps - GEMS_LoadComplaints"
        GoTo Cleanup
    End If
    
    
    oRFVComplaints.Load CLng(sEncounterNumber)
    DoEvents
    
'<< SCR 29839 non-complaint medcin ids need to have descriptions sent to rfv as comments
    For Each oInj In cInjury
            If UCase$(oInj.NewFollowup) = "NEW" Then
                encVisitTyp = enVisitTypeNew
            Else
                encVisitTyp = enVisitTypeFollowup
            End If
            Call oRFVComplaints.Add(" ", "2|" & oInj.RFVcomment, encVisitTyp, oInj.Sno.SnoID, _
            14516241, mobjLogon.FacilityNCID, sEncounterNumber, CStr(oInj.lblIndex))
    Next oInj

    oRFVComplaints.SaveComplaints
    
    '__ delete old section info if any
    'SCR-42772
    
               
    sRTF = oRFVComplaints.GetRTF(CLng(sEncounterNumber))
    
    If LenB(sRTF) <> 0 Then
        '__ compress it and save it.
        
        oSQL.Execute "Select * from enc_sections where " + _
                " EncounterNumber = " + sEncounterNumber + _
                " and enc_sectionsindex = " + CStr(NCID_REASON_FOR_VISIT)
                        
        If oSQL.EOF Then
            oSQL.AddNew
            oSQL!DataID = oSQL.GetNextID(exMAINSEQ)
        End If
        
        
        oSQL("facilityncid") = mobjLogon.FacilityNCID
        'oSQL("dataid") = DataID
        oSQL("encounternumber") = sEncounterNumber
        oSQL("enc_sectionsindex") = NCID_REASON_FOR_VISIT '   184916  '- screening app ncid
        oSQL("DTS") = Now
        oSQL("ownerncid") = mobjLogon.StoringUserNCID
        oSQL("ownername") = mobjLogon.StoringUserName
        oSQL("unit_number") = mobjPatient.UnitNumber
        oSQL("Sensitivitylevel") = 0
        oSQL("status") = 0
        oSQL("createdby") = mobjLogon.UserNCID
        oSQL("createdon") = Now
            aBytes = StrConv(Replace(sRTF, Chr(0), vbNullString), vbFromUnicode)
            objZlib.CompressData aBytes
        oSQL.Value("Doc") = aBytes
        oSQL.Value("OriginalSize") = Len(sRTF)
        
        oSQL.Update
    End If
    GEMS_LoadComplaints = True
    GoTo Cleanup
ErrHandler:
    GEMS_LoadComplaints = False
    Call mobjShared.ShowVBError(Err.Number, Err.Description, "GEMS_LoadComplaints", "CHCSII_AppointmentClient", vbExclamation)
Cleanup:
    Set objZlib = Nothing
    Set oInj = Nothing
    'Set oRFVComplaint = Nothing
    Set oRFVComplaints = Nothing
    Set oRFV = Nothing
    Set oDatabase = Nothing
    Erase aBytes
End Function

Public Function UpdateAppointmentPriority(rsApptid As String, reApptPriority As CHCSII_AppointmentClient.AppointmentPriority) As Boolean 'GEMS
    Dim objAppointment As CHCSII_AppointmentClient.Appointment
    On Error GoTo ErrHandler
    If Not mobjShared.IsAppMode(modeTheater) Then GoTo Cleanup
    Set objAppointment = GetAppointment(rsApptid)
    If objAppointment Is Nothing Then
        Err.Raise vbObjectError + 1, , "Could not retrieve apptid " & rsApptid
    End If
    mobjSQL.Execute "update appointment set TELCON_URGENCY  = " & reApptPriority & ", " _
        & " date_modified = " & SystemDateString _
        & " where appt_id = " & rsApptid
    UpdateAppointmentPriority = True
    GoTo Cleanup
ErrHandler:
    MsgBxARMd "Error in " & "CHCSII_AppointmentClient.AppointmentOps.UpdateAppointmentPriority: " & Err.Description
Cleanup:
    Set objAppointment = Nothing
End Function

Public Function UpdateAppointmentInjury_Illness(rsApptid As String, sINJ_ILL_CATEGORY As String, sINJ_ILL_CAUSE As String) As Boolean
    Dim objAppointment As CHCSII_AppointmentClient.Appointment
    On Error GoTo ErrHandler
    
    If Not mobjShared.IsAppMode(modeTheater) Then GoTo Cleanup
'<< SCR 29799 These fields can't be added to the save script if not in gems_army mode.
    If Not LenB(mobjShared.CmdLineSwitches.Item("GEMS_ARMY")) > 0 Then GoTo Cleanup
'>>
    If rsApptid = vbNullString Then
        MsgBxARMd "Invalid ApptID sent while attempting to update " + _
                   " the appointment injury and illness data."
        GoTo Cleanup
    End If
    mobjSQL.Execute "update appointment set INJ_ILL_CATEGORY  = " & oApptConn.SQLQuote(sINJ_ILL_CATEGORY, 1998) & ", " _
        & "INJ_ILL_CAUSE  = " & oApptConn.SQLQuote(sINJ_ILL_CAUSE, 1998) & " ," _
        & " date_modified = " & SystemDateString _
        & " where appt_id = " & rsApptid
        
    Set objAppointment = GetAppointment(rsApptid, False) 'Update Cache data.
    If objAppointment Is Nothing Then
        MsgBxARMd "Could not retrieve appointment data for ApptID: " & rsApptid & " while attempting to update " + _
                   " the appointment injury and illness data."
        GoTo Cleanup
    End If
        
    UpdateAppointmentInjury_Illness = True
    GoTo Cleanup
ErrHandler:
    MsgBxARMd "Error in " & "CHCSII_AppointmentClient.AppointmentOps.UpdateAppointmentInjury_Illness: " & Err.Description
Cleanup:
    Set objAppointment = Nothing
End Function

