VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Appointment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private msFacilityNCID As String
Private msApptId As String
Private msApptIEN As String
Private mdApptDateTime As Date
Private mdCheckInTime As Date
Private mdCheckOutTime As Date
Private msApptComment As String
Private msApptReason As String
Private mnApptClassification As AppointmentClassifications
Private msApptType As String
Private msApptType_Code As String 'SCR-38036
Private msApptStatus As String
Private msApptCancelReason As String
Private msEncounterNumber As String
Private msEncounterStatus As String
Private msPatientUnitNumber As String
Private msPatientIEN As String
Private msPatientName As String
Private msPatientSSN As String
Private msPatientFMP As String
Private msPatientHomePhone As String
Private msPatientWorkPhone As String
Private msSponsorSSN As String
Private msClinicNCID As String
Private msClinicName As String
Private msClinicianNCID As String
Private msProviderName As String
Private msOrderNumber As String
Private mdDateCreated As Date
Private mdDateModified As Date
Private miPriority As Integer
Private mbArrivedByAmbulance As Boolean   'SCR 9485 MSolano 9/24/01
Private msWorkload_Type As String  'N, C, or ""
Private msDisposition As String
Private mobjShared As CWShared
Private mvarCurrentCHCS_GUI_Provider As String
'APPT_STATUS_CHCS SCR-32517
Private msAppt_Status_CHCS As String
Private msAppt_Duration As String 'SCR-28213

'WAM requirement change - added props for OBSERVATION , INPATIENT_MEPRS_CODE
Private msObservation As String
Private mbRelatedToInpatient As Boolean
Private msAPPT_TYPE_CHCS As String
Private msUSV_TYPE As String

Private msPatientCallBackPhone As String

Private msMeprsCode As String  'NOT IN APPT TABLE  'Join in main appt query, Clinic Meprs Code
Private msMeprsDescription As String   'NOT IN APPT TABLE  'Join in main appt query, Clinic Meprs Description

'Only if inpatient will we have these which are in the appt table:
Private msAPPT_MEPRS_IEN As String  'INPATIENT_MEPRS_IEN
Private msINPATIENT_MEPRS_IEN As String
Private msInpatient_MEPRS_Code As String

Private msInPatient_MEPRS_Description As String

Private msAppt_Detail_Codes As String

Private msADM_Encounter As String 'In chcs_gui mode, used to capture the
                                    '"adm_status" returned back from CHCS_GUI Client.
                                    'Will show up on the appt grid and in Healt Hist form
                                    'in chcs_gui mode only... a Rich Lowell tracker.
                                    'This value will be populated in the "Encounter" column of the grid.


'SCR-42772 Theater to CurrentBuild >>>>
Private msInj_Ill_Category As String
Private msInj_Ill_Cause As String
'<<<<<<
Private msInpatient_Register_Number As String 'SCR-RNDS

'SCR-51307 Only used in GEMS mode to store the Pending Complaints.
'When a user creates a "Scheduled/Future" appt (using the /WH switch,
'we are creating a Pending appt; no encounter is created at this point.
'The appts module needs a place to store the users selected
'Reason for Visit (Chief Complaint) as selected from the
'dlgGEMS_NewApptArmy screen. These values are stored here only while
'the appt is pending. When the appt is double clicked to checkin
'and create the encounter we then rebuild the Complaints object using
'the BodyImage Injury object. This injury object is then saved to Encounters
'in the Global.bas - NewArmyApp_LoadComplaints_GEMS.  When the Appt DataLayer
'is implemented, it will be saved via Encounters DataLayer.
Private mcolPendingComplaints As Collection
Public Property Set PendingComplaints(ByVal vData As Collection)
    Set mcolPendingComplaints = vData
End Property
Public Property Get PendingComplaints() As Collection
    Set PendingComplaints = mcolPendingComplaints
End Property
'<<< end SCR-51307


'Private Type TIME_ZONE_INFORMATION
'        Bias As Long
'        StandardName(31) As Integer
'        StandardDate As SYSTEMTIME
'        StandardBias As Long
'        DaylightName(31) As Integer
'        DaylightDate As SYSTEMTIME
'        DaylightBias As Long
'End Type
'Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
'Private lTZinfo As TIME_ZONE_INFORMATION


Public Property Let Inpatient_Register_Number(ByVal vData As String)
    msInpatient_Register_Number = vData
End Property
Public Property Get Inpatient_Register_Number() As String
    Inpatient_Register_Number = msInpatient_Register_Number
End Property
Public Property Let Inj_Ill_Cause(ByVal vData As String)
    msInj_Ill_Cause = vData
End Property
Public Property Get Inj_Ill_Cause() As String
    Inj_Ill_Cause = msInj_Ill_Cause
End Property

Public Property Let Inj_Ill_Category(ByVal vData As String)
    msInj_Ill_Category = vData
End Property

Public Property Get Inj_Ill_Category() As String
    Inj_Ill_Category = msInj_Ill_Category
End Property


Public Property Let Appt_Detail_Codes(ByVal vData As String)
    msAppt_Detail_Codes = vData
End Property


Public Property Get Appt_Detail_Codes() As String
    Appt_Detail_Codes = msAppt_Detail_Codes
End Property
Public Property Let PatientCallBackPhone(ByVal vsPatientCallBackPhone As String)
    msPatientCallBackPhone = vsPatientCallBackPhone
End Property

Public Property Get PatientCallBackPhone() As String
    PatientCallBackPhone = msPatientCallBackPhone
End Property


Public Property Let INPATIENT_MEPRS_IEN(ByVal vData As String)
    msINPATIENT_MEPRS_IEN = vData
End Property


Public Property Get INPATIENT_MEPRS_IEN() As String
     INPATIENT_MEPRS_IEN = msINPATIENT_MEPRS_IEN
End Property



Public Property Let APPT_MEPRS_IEN(ByVal vData As String)
    msAPPT_MEPRS_IEN = vData
End Property


Public Property Get APPT_MEPRS_IEN() As String
     APPT_MEPRS_IEN = msAPPT_MEPRS_IEN
End Property



Public Property Let USV_TYPE(ByVal vData As String)
    msUSV_TYPE = vData
End Property


Public Property Get USV_TYPE() As String
     USV_TYPE = msUSV_TYPE
End Property



Public Property Let APPT_TYPE_CHCS(ByVal vData As String)
    msAPPT_TYPE_CHCS = vData
End Property


Public Property Get APPT_TYPE_CHCS() As String
     APPT_TYPE_CHCS = msAPPT_TYPE_CHCS
End Property




Public Property Let Appt_Duration(ByVal vData As String)
    msAppt_Duration = vData
End Property


Public Property Get Appt_Duration() As String
    Appt_Duration = msAppt_Duration
End Property



Public Property Let RelatedToInpatient(ByVal vData As Boolean)
    mbRelatedToInpatient = vData
End Property
Public Property Get RelatedToInpatient() As Boolean
    RelatedToInpatient = mbRelatedToInpatient
End Property

'SCR-32517
Public Property Let Appt_Status_CHCS(ByVal vData As String)
    msAppt_Status_CHCS = vData
End Property

'SCR-32517
Public Property Get Appt_Status_CHCS() As String
    Appt_Status_CHCS = msAppt_Status_CHCS
End Property
Public Property Let InPatient_MEPRS_Description(ByVal vData As String)
    msInPatient_MEPRS_Description = vData
End Property

Public Property Get InPatient_MEPRS_Description() As String
    InPatient_MEPRS_Description = msInPatient_MEPRS_Description
End Property
Public Property Let Inpatient_MEPRS_Code(ByVal vData As String)
    msInpatient_MEPRS_Code = vData
End Property

Public Property Get Inpatient_MEPRS_Code() As String
    Inpatient_MEPRS_Code = msInpatient_MEPRS_Code
End Property

Public Property Let Observation(ByVal vData As String)
    msObservation = vData
End Property

Public Property Get Observation() As String
    Observation = msObservation
End Property

Public Property Let Workload_Type(ByVal vData As String)
    msWorkload_Type = vData
End Property


Public Property Get Workload_Type() As String
    Workload_Type = msWorkload_Type
End Property

'--- SCR 14798   Sherry Wang 12/7/2004
Public Property Let Disposition(ByVal vData As String)
    msDisposition = vData
End Property

Public Property Get Disposition() As String
    Disposition = msDisposition
End Property

Public Property Let ADM_Encounter(ByVal vData As String)
    msADM_Encounter = vData
End Property
Public Property Get ADM_Encounter() As String
    ADM_Encounter = msADM_Encounter
End Property

Public Property Let CurrentCHCS_GUI_Provider(ByVal vData As String)
    mvarCurrentCHCS_GUI_Provider = vData
End Property

Public Property Get CurrentCHCS_GUI_Provider() As String
    CurrentCHCS_GUI_Provider = mvarCurrentCHCS_GUI_Provider
End Property


Public Property Get FacilityNCID() As String
    FacilityNCID = msFacilityNCID
End Property

Public Property Let FacilityNCID(rsValue As String)
    msFacilityNCID = rsValue
End Property

Public Property Get ApptId() As String
    ApptId = msApptId
End Property

Public Property Let ApptId(rsValue As String)
    msApptId = rsValue
End Property
Public Property Let ApptIEN(rsValue As String)
    msApptIEN = rsValue
End Property

Public Property Get ApptIEN() As String
    ApptIEN = msApptIEN
End Property

Public Property Get ApptDateTime() As Date
    ApptDateTime = mdApptDateTime
End Property
Public Property Let ApptDateTime(ByVal vdApptDateTime As Date)
    mdApptDateTime = vdApptDateTime
End Property
Public Property Get Priority() As Integer    '--JCB
    Priority = miPriority
End Property
Public Property Let Priority(ByVal viPriority As Integer)
    miPriority = viPriority
End Property
Public Property Get CheckInTime() As Date
    CheckInTime = mdCheckInTime
End Property

Public Property Let CheckInTime(ByVal vdCheckInTime As Date)
    mdCheckInTime = vdCheckInTime
End Property

Public Property Get CheckOutTime() As Date
    CheckOutTime = mdCheckOutTime
End Property

Public Property Let CheckOutTime(ByVal vdCheckOutTime As Date)
    mdCheckOutTime = vdCheckOutTime
End Property

Public Property Get ApptComment() As String
    ApptComment = msApptComment
End Property

Public Property Let ApptComment(ByVal vsApptComment As String)
    msApptComment = vsApptComment
End Property

Public Property Get ApptReason() As String
    ApptReason = msApptReason
End Property

Public Property Let ApptReason(ByVal vsApptReason As String)
    msApptReason = vsApptReason
End Property

Public Property Get ApptClassEnum() As AppointmentClassifications
    ApptClassEnum = mnApptClassification
End Property

Public Property Get ApptClassification() As String
    Select Case mnApptClassification
    Case ApptInpatient
        ApptClassification = "Inpatient"
    Case ApptAmbulatoryProcedure
        ApptClassification = "AmbulatoryProcedure"
    Case ApptOutpatient
        ApptClassification = "OutPatient"
    Case Else
        ApptClassification = ""
    End Select
End Property

Public Property Let ApptClassification(ByVal vsApptClassification As String)
    Select Case vsApptClassification
    Case "Inpatient"
        mnApptClassification = ApptInpatient
    Case "AmbulatoryProcedure"
        mnApptClassification = ApptAmbulatoryProcedure
    Case Else
        mnApptClassification = ApptOutpatient
    End Select
End Property

Public Property Get ApptType() As String
    ApptType = msApptType
End Property

Public Property Let ApptType(ByVal vsApptType As String)
    msApptType = vsApptType
End Property


'SCR-38036
Public Property Get ApptType_Code() As String
    ApptType_Code = msApptType_Code
End Property

Public Property Let ApptType_Code(ByVal vsApptType_Code As String)
    msApptType_Code = vsApptType_Code
End Property



Public Property Get ApptStatus() As String
    ApptStatus = msApptStatus
End Property

Public Property Let ApptStatus(ByVal vsApptStatus As String)
    msApptStatus = vsApptStatus
End Property

Public Property Get ApptCancelReason() As String
    ApptCancelReason = msApptCancelReason
End Property

Public Property Let ApptCancelReason(ByVal vsApptCancelReason As String)
    msApptCancelReason = vsApptCancelReason
End Property

Public Property Get EncounterNumber() As String
    EncounterNumber = msEncounterNumber
End Property

Public Property Let EncounterNumber(ByVal vsEncounterNumber As String)
    msEncounterNumber = vsEncounterNumber
End Property
Public Property Let EncounterStatus(ByVal p_oEncounterStatus As String)
    msEncounterStatus = p_oEncounterStatus
End Property
Public Property Get EncounterStatus() As String
    If msEncounterStatus <> "" And msEncounterStatus <> "0" Then
        Dim x As EncounterStatusEnum
        x = msEncounterStatus
        EncounterStatus = TranslateEncStatus(x)
    Else
        EncounterStatus = ""
    End If
End Property

Public Property Get PatientUnitNumber() As String
    PatientUnitNumber = msPatientUnitNumber
End Property

Public Property Let PatientUnitNumber(ByVal vsPatientUnitNumber As String)
    Dim objPatient As MMMHISPatient.Patient
    msPatientUnitNumber = vsPatientUnitNumber
'    Set objPatient = New MMMHISPatient.Patient
'    Set objPatient.Config = cwobjConfig
'    objPatient.LoadPatient vsPatientUnitNumber, ""
'    msPatientName = objPatient.Name
'    msPatientSSN = objPatient.SSN
'    msPatientHomePhone = objPatient.HomePhone
'    msPatientWorkPhone = objPatient.WorkPhone
'    msPatientIEN = objPatient.MedicalRecordNumber
'todo: finish populating read-only patient info--if it makes sense to do it
'we are missing the fmp and the sponsor ssn
    Set objPatient = Nothing
End Property

Public Property Get PatientIEN() As String
    PatientIEN = msPatientIEN
End Property
Public Property Let PatientIEN(vsPatientIEN As String)
    msPatientIEN = vsPatientIEN
End Property

Public Property Get PatientName() As String
    PatientName = msPatientName
End Property
Public Property Let PatientName(vsPatientName As String)
    msPatientName = vsPatientName
End Property
Public Property Get PatientSSN() As String
    PatientSSN = msPatientSSN
End Property
Public Property Let PatientSSN(vsPatientSSN As String)
    msPatientSSN = vsPatientSSN
End Property
Public Property Get PatientFMP() As String
    PatientFMP = msPatientFMP
End Property
Public Property Let PatientFMP(vsPatientFMP As String)
    msPatientFMP = vsPatientFMP
End Property
Public Property Get PatientHomePhone() As String
    PatientHomePhone = msPatientHomePhone
End Property

Public Property Get PatientWorkPhone() As String
    PatientWorkPhone = msPatientWorkPhone
End Property
Public Property Let PatientWorkPhone(vsPatientWorkPhone As String)
    msPatientWorkPhone = vsPatientWorkPhone
End Property
Public Property Let PatientHomePhone(vsPatientHomePhone As String)
    msPatientHomePhone = vsPatientHomePhone
End Property

Public Property Get SponsorSSN() As String
    SponsorSSN = msSponsorSSN
End Property
Public Property Let SponsorSSN(vsSponsorSSN As String)
    msSponsorSSN = vsSponsorSSN
End Property
Public Property Get ClinicNCID() As String
    ClinicNCID = msClinicNCID
End Property

Public Property Let ClinicNCID(vsClinicNCID As String)
    msClinicNCID = vsClinicNCID
End Property

Public Property Get ClinicName() As String
    ClinicName = msClinicName
End Property
Public Property Let ClinicName(ByVal vsClinicName As String)
    msClinicName = vsClinicName
End Property
Public Property Get ClinicianNCID() As String
    ClinicianNCID = msClinicianNCID
End Property

Public Property Let ClinicianNCID(ByVal vsClinicianNCID As String)
    msClinicianNCID = vsClinicianNCID
End Property

Public Property Get ProviderName() As String
    ProviderName = msProviderName
End Property
Public Property Let ProviderName(vsProviderName As String)
    msProviderName = vsProviderName
End Property
Public Property Get MeprsCode() As String
    MeprsCode = msMeprsCode
End Property

Public Property Let MeprsCode(vsMeprsCode As String)
    msMeprsCode = vsMeprsCode
End Property

Public Property Get MeprsDescription() As String
    MeprsDescription = msMeprsDescription
End Property
Public Property Let MeprsDescription(vsMeprsDescription As String)
    msMeprsDescription = vsMeprsDescription
End Property



'DEM 01/03/2000
Public Property Get OrderNumber() As String
    OrderNumber = msOrderNumber
End Property

'DEM 01/03/2000
Public Property Let OrderNumber(sOrderNumber As String)
    msOrderNumber = sOrderNumber
End Property

'SCR-9485 MSolano 10/8/01
Public Property Get ArrivedByAmbulance() As Boolean
  ArrivedByAmbulance = mbArrivedByAmbulance
End Property
'SCR-9485 MSolano 10/8/01
Public Property Let ArrivedByAmbulance(bData As Boolean)
  mbArrivedByAmbulance = bData
End Property

Public Function Deserialize(robjSQL As ICHCSII_SQL.ISqlOps) As Boolean
'code review note Solano 3/24/04:
'recommend replacing on err resume next w/ err handler
'(public class; public method; references objects)

    If robjSQL Is Nothing Then
        Exit Function
    End If
    On Error GoTo ErrHandler

    msFacilityNCID = robjSQL("Facility_NCID")
    msApptId = robjSQL("Appt_Id")
    msApptIEN = robjSQL("Appt_IEN")
    
    
   '******* Appt_Date_Time **********************************************
    If Not IsNull(robjSQL("Appt_Date_Time")) Then
        If IsDate(robjSQL("Appt_Date_Time")) Then
            If robjSQL("Appt_Date_Time") > #12/31/2100# Or robjSQL("Appt_Date_Time") < #12/31/1990# Or _
                robjSQL("Appt_Date_Time") = 0 Then
                mdApptDateTime = 0
            Else
                mdApptDateTime = robjSQL("Appt_Date_Time")
            End If
        Else
            mdApptDateTime = 0
        End If
    Else
        mdApptDateTime = 0
    End If
   '********* CheckIn_Time ****mdCheckInTime****************************************
    If Not IsNull(robjSQL("CheckIn_Time")) Then
        If IsDate(robjSQL("CheckIn_Time")) Then
            If robjSQL("CheckIn_Time") > #12/31/2100# Or robjSQL("CheckIn_Time") < #12/31/1990# Or _
                robjSQL("CheckIn_Time") = 0 Then
                mdCheckInTime = 0
            Else
                mdCheckInTime = robjSQL("CheckIn_Time")
            End If
        Else
            mdCheckInTime = 0
        End If
    Else
        mdCheckInTime = 0
    End If
    '*********** CheckOut_Time ****mdCheckOutTime**************************************
    If Not IsNull(robjSQL("CheckOut_Time")) Then
        If IsDate(robjSQL("CheckOut_Time")) Then
            If robjSQL("CheckOut_Time") > #12/31/2100# Or robjSQL("CheckOut_Time") < #12/31/1990# Or _
                robjSQL("CheckOut_Time") = 0 Then
                mdCheckOutTime = 0
            Else
                mdCheckOutTime = robjSQL("CheckOut_Time")
            End If
        Else
            mdCheckOutTime = 0
        End If
    Else
        mdCheckOutTime = 0
    End If
    '*********** Date_Created ****mdDateCreated**************************************
    If Not IsNull(robjSQL("Date_Created")) Then
        If IsDate(robjSQL("Date_Created")) Then
            If robjSQL("Date_Created") > #12/31/2100# Or robjSQL("Date_Created") < #12/31/1990# Or _
                robjSQL("Date_Created") = 0 Then
                mdDateCreated = 0
            Else
                mdDateCreated = robjSQL("Date_Created")
            End If
        Else
            mdDateCreated = 0
        End If
    Else
        mdDateCreated = 0
    End If
    
    '*********** Date_Modified ****mdDateModified**************************************
    If Not IsNull(robjSQL("Date_Modified")) Then
        If IsDate(robjSQL("Date_Modified")) Then
            If robjSQL("Date_Modified") > #12/31/2100# Or robjSQL("Date_Modified") < #12/31/1990# Or _
                robjSQL("Date_Modified") = 0 Then
                mdDateModified = 0
            Else
                mdDateModified = robjSQL("Date_Modified")
            End If
        Else
            mdDateModified = 0
        End If
    Else
        mdDateModified = 0
    End If
    '*****************************************************
    
    msApptComment = robjSQL("Appt_Comment")
    msApptReason = robjSQL("Appt_Reason")
    If IsNull(robjSQL("Appt_Classification")) Then
        mnApptClassification = 0
    Else
        mnApptClassification = robjSQL("Appt_Classification")
    End If
    
    msApptType = robjSQL("Appt_Type")
    If msApptType = "RNDS*" Or UCase(msApptType) = "INPATIENT WARD APPT" Then
        msApptType = "RNDS"
    End If
    msApptStatus = robjSQL("Appt_Status")
    
    msApptCancelReason = robjSQL("Appt_Cancel_Reason")
    msEncounterNumber = robjSQL("Encounter_Number")
    If msEncounterNumber = "0" Then
        msEncounterNumber = ""
    End If
    msEncounterStatus = robjSQL("Encounter_Status")
    msPatientUnitNumber = robjSQL("Patient_Unit_Number")
    msPatientIEN = robjSQL("Patient_IEN")
    msPatientName = robjSQL("Patient_Name")
    msPatientSSN = robjSQL("Patient_SSN")
    msPatientFMP = robjSQL("Patient_FMP")
    msPatientHomePhone = robjSQL("Patient_Home_Phone")
    msPatientWorkPhone = robjSQL("Patient_Work_Phone")
    msSponsorSSN = robjSQL("Sponsor_SSN")
    msClinicNCID = robjSQL("Clinic_NCID")
    msClinicName = robjSQL("Clinic_Name")
    msClinicianNCID = robjSQL("Clinician_NCID")
    msProviderName = robjSQL("Provider_Name")
    msMeprsCode = robjSQL("Meprs_Code")
    'jrm:Added for WAM to distinguish between Clinic Level MEPRS and Inpatient MEPRS (B Level or A Level)
    msMeprsDescription = robjSQL("Meprs_Description")
    msOrderNumber = robjSQL("Order_Number")
    
    'KDunne 3/2/2004
    'ITT Convergence
    If mobjShared.IsAppMode(modeCDR) And Not mobjShared.IsAppMode(modeITT) Then
        msWorkload_Type = robjSQL("Workload_Type")
        msObservation = robjSQL("Observation")
        msInpatient_MEPRS_Code = robjSQL("Inpatient_MEPRS_CODE")
    'R. Foster SCR #40191; 40140;   Developer: Robert Foster 10/21/2003 11:36 PM
    'R. Foster Note: Function Changed: CHCSII_Appointments.dll\Appointments.frm: FillGrid ; CHCSII_AppointmentClient.dll\ AppointmentOps.cls: Declarations, GetAppointment; Appointment.cls: Deserialize, AddAllFields      Impact of Change: Modifiy SQL statements to return InPatient and Outpatient states
    'ITT UPDATES
    ElseIf mobjShared.IsAppMode(modeITT) Then
        'SCR-40190
        msInPatient_MEPRS_Description = robjSQL("InPatient_MEPRS_Description")
        msAPPT_MEPRS_IEN = robjSQL("APPT_MEPRS_IEN") & ""
        msWorkload_Type = robjSQL("Workload_Type") & ""
        msObservation = robjSQL("Observation")
        msInpatient_MEPRS_Code = robjSQL("Inpatient_MEPRS_CODE")
        If UCase(robjSQL("RELATED_TO_INPATIENT") & "") = "TRUE" Then
            mbRelatedToInpatient = True
        Else
            mbRelatedToInpatient = False
        End If
        msAppt_Duration = robjSQL("Appt_Duration") 'SCR-40351
    End If
    
    If mobjShared.IsAppMode(modeTheater) Then
        If IsNull(robjSQL("Telcon_Urgency")) = False Then
            If robjSQL("Telcon_Urgency") = "" Then
                miPriority = 0
            Else
                miPriority = CInt(robjSQL("Telcon_Urgency"))
            End If
        Else
            miPriority = 0
        End If
        '<< SCR 29799 These fields can't be added to the save script if not in gems_army mode.
        If IsNull(robjSQL("Inj_Ill_Category")) = False Then 'SCR-34147
            'Need to leave this in to handle any data previously saved prior to SCR-51307
            'Can remove after a next release.
            If InStr(1, robjSQL("Inj_Ill_Category"), "###", vbTextCompare) = 0 Then 'not used for complaint cache
                msInj_Ill_Category = robjSQL("Inj_Ill_Category") & ""
            End If
        End If
        
        msInj_Ill_Cause = robjSQL("Inj_Ill_Cause") & ""
        
        If IsNull(robjSQL("Pending_Complaints")) = False Then
            If robjSQL("Pending_Complaints") <> vbNullString Then
                Set mcolPendingComplaints = GEMS_Parse_Pending_Complaints(robjSQL("Pending_Complaints"))
            End If
        End If
        
        If LenB(mobjShared.CmdLineSwitches.Item("INPATIENT2")) > 0 Then
            msDisposition = robjSQL("DISPOSITION")
        End If
    End If
    Deserialize = True
    
    Exit Function

ErrHandler:
    Resume Next
End Function
Private Function GEMS_Parse_Pending_Complaints(sComplaints As String) As Collection
'SCR-51307
Dim ComplaintRow As Variant
Dim i As Integer
On Error GoTo ErrHandler

    If Trim(sComplaints) = vbNullString Then Exit Function
    ComplaintRow = Split(sComplaints, "|||", , vbTextCompare)
    If UBound(ComplaintRow) >= 0 Then
        Set GEMS_Parse_Pending_Complaints = New Collection
    Else
        Exit Function
    End If
    For i = 0 To UBound(ComplaintRow)
        GEMS_Parse_Pending_Complaints.Add ComplaintRow(i)
    Next
Exit Function
ErrHandler:
mobjShared.ShowVBError Err.Number, Err.Description, "Appointment.cls.GEMS_Parse_Pending_Complaints", "CHCSII_AppointmentClient", vbExclamation
End Function
'Private Function AdjustDate(ByVal dDate As Date) As Date
'    Dim lYear As Integer
'    Dim dStandard As Date
'    Dim dDaylight As Date
'    Dim lTimeBias As Long
'
'    If Year(dDate) <> lYear Then
'        lYear = Year(dDate)
'        dStandard = GetTZInfoDate(lTZinfo.StandardDate, lYear)
'        dDaylight = GetTZInfoDate(lTZinfo.DaylightDate, lYear)
'    End If
'
'    If dStandard > dDaylight Then
'        If dDate >= dDaylight And dDate < dStandard Then
'            lTimeBias = lTZinfo.Bias + lTZinfo.DaylightBias
'        Else
'            lTimeBias = lTZinfo.Bias + lTZinfo.StandardBias
'        End If
'    Else
'        If dDate >= dStandard And dDate < dDaylight Then
'            lTimeBias = lTZinfo.Bias + lTZinfo.StandardBias
'        Else
'            lTimeBias = lTZinfo.Bias + lTZinfo.DaylightBias
'        End If
'    End If
'    AdjustDate = DateAdd("n", lTimeBias * -1, dDate)
'End Function

'Private Function GetTZInfoDate(TZDate As SYSTEMTIME, TZYear As Integer) As Date
'    With TZDate
'        If .wDay = 5 Then
'            ' means last .wDayOfWeek in month
'            GetTZInfoDate = DateSerial(TZYear, .wMonth + 1, 0) + TimeSerial(.wHour, .wMinute, .wSecond)
'            Do While Weekday(GetTZInfoDate) <> .wDayOfWeek + 1
'                GetTZInfoDate = DateAdd("d", -1, GetTZInfoDate)
'            Loop
'        Else
'            ' means .wday'TH .wdayofweek in month
'            GetTZInfoDate = DateSerial(TZYear, .wMonth, 1) + TimeSerial(.wHour, .wMinute, .wSecond)
'            Do While Weekday(GetTZInfoDate) <> .wDayOfWeek + 1
'                GetTZInfoDate = DateAdd("d", 1, GetTZInfoDate)
'            Loop
'            GetTZInfoDate = DateAdd("d", (.wDay - 1) * 7, GetTZInfoDate)
'        End If
'    End With
'End Function
Public Function DeserializeAppts_SP(rs As ADODB.Recordset) As Boolean
    If rs Is Nothing Then
        Exit Function
    End If
    On Error Resume Next
    
    msFacilityNCID = rs("Facility_NCID")
    msApptId = rs("Appt_Id")
    msApptIEN = rs("Appt_IEN")
    
    'dCurrDate = DateAdd("n", (lTZinfo.Bias * -1), Now())
    
    
   '******* Appt_Date_Time **********************************************
    If Not IsNull(rs("Appt_Date_Time")) Then
        If IsDate(rs("Appt_Date_Time")) Then
            If rs("Appt_Date_Time") > #12/31/2100# Or rs("Appt_Date_Time") < #12/31/1990# Or _
                rs("Appt_Date_Time") = 0 Then
                mdApptDateTime = 0
            Else
                mdApptDateTime = rs("Appt_Date_Time")
                'mdApptDateTime = DateAdd("n", (lTZinfo.Bias * -1), rs("Appt_Date_Time"))
            End If
        Else
            mdApptDateTime = 0
        End If
    Else
        mdApptDateTime = 0
    End If
   '********* CheckIn_Time ****mdCheckInTime****************************************
    If Not IsNull(rs("CheckIn_Time")) Then
        If IsDate(rs("CheckIn_Time")) Then
            If rs("CheckIn_Time") > #12/31/2100# Or rs("CheckIn_Time") < #12/31/1990# Or _
                rs("CheckIn_Time") = 0 Then
                mdCheckInTime = 0
            Else
                mdCheckInTime = rs("CheckIn_Time")
            End If
        Else
            mdCheckInTime = 0
        End If
    Else
        mdCheckInTime = 0
    End If
    '*********** CheckOut_Time ****mdCheckOutTime**************************************
    If Not IsNull(rs("CheckOut_Time")) Then
        If IsDate(rs("CheckOut_Time")) Then
            If rs("CheckOut_Time") > #12/31/2100# Or rs("CheckOut_Time") < #12/31/1990# Or _
                rs("CheckOut_Time") = 0 Then
                mdCheckOutTime = 0
            Else
                 mdCheckOutTime = rs("CheckOut_Time")
            End If
        Else
            mdCheckOutTime = 0
        End If
    Else
        mdCheckOutTime = 0
    End If
    '*********** Date_Created ****mdDateCreated**************************************
    If Not IsNull(rs("Date_Created")) Then
        If IsDate(rs("Date_Created")) Then
            If rs("Date_Created") > #12/31/2100# Or rs("Date_Created") < #12/31/1990# Or _
                rs("Date_Created") = 0 Then
                mdDateCreated = 0
            Else
                mdDateCreated = rs("Date_Created")
            End If
        Else
            mdDateCreated = 0
        End If
    Else
        mdDateCreated = 0
    End If
    
    '*********** Date_Modified ****mdDateModified**************************************
    If Not IsNull(rs("Date_Modified")) Then
        If IsDate(rs("Date_Modified")) Then
            If rs("Date_Modified") > #12/31/2100# Or rs("Date_Modified") < #12/31/1990# Or _
                rs("Date_Modified") = 0 Then
                mdDateModified = 0
            Else
                mdDateModified = rs("Date_Modified")
            End If
        Else
            mdDateModified = 0
        End If
    Else
        mdDateModified = 0
    End If
    '*****************************************************
    'SCR-32517
    msAppt_Status_CHCS = rs("APPT_STATUS_CHCS") & ""
    msApptComment = rs("Appt_Comment")
    msApptReason = rs("Appt_Reason")
    If IsNull(rs("Appt_Classification")) Then
        mnApptClassification = 0
    Else
        mnApptClassification = rs("Appt_Classification")
    End If
    
    msApptType = rs("Appt_Type")  'This is the appt type description as displayed in grid.
    If msApptType = "RNDS*" Or UCase(msApptType) = "INPATIENT WARD APPT" Then
        msApptType = "RNDS"
    End If
    msApptType_Code = rs("Appt_Type_Code") 'SCR-38036
    msApptStatus = rs("Appt_Status")
    msApptCancelReason = rs("Appt_Cancel_Reason")
    msEncounterNumber = rs("Encounter_Number")
    msEncounterStatus = rs("Encounter_Status")
    msPatientUnitNumber = rs("Patient_Unit_Number")
    msPatientIEN = rs("Patient_IEN")
    msPatientName = rs("Patient_Name")
    msPatientSSN = rs("Patient_SSN")
    msPatientFMP = rs("Patient_FMP")
    msPatientHomePhone = rs("Patient_Home_Phone")
    msPatientWorkPhone = rs("Patient_Work_Phone")
    msSponsorSSN = rs("Sponsor_SSN")
    msClinicNCID = rs("Clinic_NCID")
    msClinicName = rs("Clinic_Name")
    msClinicianNCID = rs("Clinician_NCID")
    msProviderName = rs("Provider_Name")
    
    
    
    
    msOrderNumber = rs("Order_Number") 'DEM 01/03/2000
    msWorkload_Type = rs("Workload_Type")
    msObservation = rs("Observation")
    
    
    '****************************************************************
    If UCase(rs("RELATED_TO_INPATIENT") & "") = "TRUE" Then
        mbRelatedToInpatient = True
    Else
        mbRelatedToInpatient = False
    End If
    
    msInpatient_MEPRS_Code = rs("Inpatient_MEPRS_CODE")
    msMeprsCode = rs("Meprs_Code")
    msAPPT_MEPRS_IEN = rs("APPT_MEPRS_IEN") & ""
    msInPatient_MEPRS_Description = rs("InPatient_MEPRS_Description")
    msMeprsDescription = rs("Meprs_Description")
    If IsNull(rs("INPATIENT_MEPRS_IEN")) Then
        msINPATIENT_MEPRS_IEN = ""
    Else
        msINPATIENT_MEPRS_IEN = CStr(rs("INPATIENT_MEPRS_IEN"))
    End If
    
    '****************************************************************
    
    msAppt_Duration = rs("Appt_Duration") 'SCR-28213
    msAPPT_TYPE_CHCS = rs("APPT_TYPE_CHCS") & ""
    msUSV_TYPE = rs("USV_TYPE") & ""
    
    'SCR-38720
    msAppt_Detail_Codes = rs("Appt_Detail_Codes")
        
    DeserializeAppts_SP = True
End Function
Public Function Deserialize_CHCS_GUI(rs As ADODB.Recordset, sFacilityNCID As String) As Boolean
'Dim i As Integer

'All possible field in recordset as of 7/17/03
'Appt_Type: ROUT$
'APPT_STATUS: WALK-IN
'adm_status: 0
'Appt_Reason: 'flowsheets1'
'CheckIn_Time: 7/16/2003 12:01:00 PM
'Appt_Classification: 0
'Clinic_Name: BLUE MTF
'MEPRS_CODE: BHAA
'Appt_Comment: 'CHCSII'
'Appt_IEN: 1090457
'APPT_MEPRS_IEN: 617
'USV_TYPE: WALK_IN
'FACILITY_NCID: 0
'Appt_Id: 1090457
'Telcon_IEN:
'CheckOut_Time:
'Description:
'Appt_Cancel_Reason:
'Encounter_Number:
'Encounter_Status:
'Patient_IEN: 85339
'Order_Number:
'Clinic_NCID:
'Clinic_IEN: 1623
'Clinician_NCID:
'Clinician_IEN: 9436
'Provider_Name: TEN , DOCTOR
'MEPRS_Description:
'Date_Created:
'Date_Modified:
'Patient_Name: GARCIA,QUENTIN M JR
'patient_ssn: 137-32-6424
'patient_work_phone: 9192178460
'patient_home_phone: 9192178460
'patient_fmp: 20
'sponsor_ssn: 137-32-6424
'workload_type: C
'PATIENT_STATUS: Outpatient
'Inpatient_MEPRS_CODE:
'DURATION: 30


    If rs Is Nothing Then
        Exit Function
    End If
    On Error GoTo ErrorTrap
    msFacilityNCID = rs("Facility_NCID") & ""
    If msFacilityNCID = "0" Or msFacilityNCID = "" Then
        msFacilityNCID = sFacilityNCID
    End If
    msApptId = rs("Appt_IEN") & ""
    msApptIEN = rs("Appt_IEN") & ""
   '******* Appt_Date_Time **********************************************
    If Not IsNull(rs("Appt_Date_Time")) Then
        If IsDate(rs("Appt_Date_Time")) Then
            If rs("Appt_Date_Time") > #12/31/2100# Or rs("Appt_Date_Time") < #12/31/1990# Or _
                rs("Appt_Date_Time") = 0 Then
                mdApptDateTime = 0
            Else
                mdApptDateTime = Format(rs("Appt_Date_Time"), "mm/dd/yyyy hh:mm")
            End If
        Else
            mdApptDateTime = 0
        End If
    Else
        mdApptDateTime = 0
    End If
    'Format(rs("Date_Created"), "mm/dd/yyyy hh:mm")
   '********* CheckIn_Time ****mdCheckInTime****************************************
    If Not IsNull(rs("CheckIn_Time")) Then
        If IsDate(rs("CheckIn_Time")) Then
            If rs("CheckIn_Time") > #12/31/2100# Or rs("CheckIn_Time") < #12/31/1990# Or _
                rs("CheckIn_Time") = 0 Then
                mdCheckInTime = 0
            Else
                mdCheckInTime = Format(rs("CheckIn_Time"), "mm/dd/yyyy hh:mm")
            End If
        Else
            mdCheckInTime = 0
        End If
    Else
        mdCheckInTime = 0
    End If
    '*********** CheckOut_Time ****mdCheckOutTime**************************************
    If Not IsNull(rs("CheckOut_Time")) Then
        If IsDate(rs("CheckOut_Time")) Then
            If rs("CheckOut_Time") > #12/31/2100# Or rs("CheckOut_Time") < #12/31/1990# Or _
                rs("CheckOut_Time") = 0 Then
                mdCheckOutTime = 0
            Else
                mdCheckOutTime = Format(rs("CheckOut_Time"), "mm/dd/yyyy hh:mm")
            End If
        Else
            mdCheckOutTime = 0
        End If
    Else
        mdCheckOutTime = 0
    End If
    '*********** Date_Created ****mdDateCreated**************************************
    If Not IsNull(rs("Date_Created")) Then
        If IsDate(rs("Date_Created")) Then
            If rs("Date_Created") > #12/31/2100# Or rs("Date_Created") < #12/31/1990# Or _
                rs("Date_Created") = 0 Then
                mdDateCreated = 0
            Else
                mdDateCreated = Format(rs("Date_Created"), "mm/dd/yyyy hh:mm")
            End If
        Else
            mdDateCreated = 0
        End If
    Else
        mdDateCreated = 0
    End If
    
    '*********** Date_Modified ****mdDateModified**************************************
    If Not IsNull(rs("Date_Modified")) Then
        If IsDate(rs("Date_Modified")) Then
            If rs("Date_Modified") > #12/31/2100# Or rs("Date_Modified") < #12/31/1990# Or _
                rs("Date_Modified") = 0 Then
                mdDateModified = 0
            Else
                mdDateModified = Format(rs("Date_Modified"), "mm/dd/yyyy hh:mm")
            End If
        Else
            mdDateModified = 0
        End If
    Else
        mdDateModified = 0
    End If
    '*****************************************************
    
   
    msApptComment = rs("Appt_Comment") & ""
    msApptReason = rs("Appt_Reason") & ""
    If IsNull(rs("Appt_Classification")) Then
        mnApptClassification = ApptOutpatient
    Else
        If IsNumeric(rs("Appt_Classification")) Then
            Select Case CInt(rs("Appt_Classification"))
                Case 0
                    mnApptClassification = ApptOutpatient
                Case 1
                    mnApptClassification = ApptInpatient
                
                Case 2
                    mnApptClassification = ApptAmbulatoryProcedure
                Case 3
                    mnApptClassification = ApptTelCon
                Case Else
                    mnApptClassification = ApptOutpatient
            End Select
        
        Else
            If InStr(1, UCase(CStr(rs("Appt_Classification"))), "OUTPATIENT", vbTextCompare) > 0 Then
                mnApptClassification = ApptOutpatient
            ElseIf InStr(1, UCase(CStr(rs("Appt_Classification"))), "INPATIENT", vbTextCompare) > 0 Then
                mnApptClassification = ApptInpatient
            Else
                mnApptClassification = ApptOutpatient
            End If
        End If
    End If
    
    'If InStr(1, "T-CON* TEL-CON TCON TCON$ TCONX OCC-SVR", msApptType, vbTextCompare) > 0 Then
    msApptStatus = rs("Appt_Status") & ""

    msApptType = rs("Appt_Type") & ""
    If msApptType = "" Then
        If msApptStatus = "OCC-SVC" Or msApptStatus = "TEL-CON" Then
            msApptType = "TCON"
        Else
            msApptType = msApptStatus
        End If
    End If
    If InStr(1, msApptType, "RNDS", vbTextCompare) > 0 Then
        msApptType = "RNDS"
    End If
    msApptType_Code = msApptType 'SCR-38036
    msApptCancelReason = rs("Appt_Cancel_Reason") & ""
    If Not IsNull(rs("adm_status")) Then
        If CStr(rs("adm_status")) = "0" Then
            msADM_Encounter = ""
        ElseIf CStr(rs("adm_status")) = "1" Then
            msADM_Encounter = "Complete"
        Else
            msADM_Encounter = ""
        End If
    Else
        msADM_Encounter = ""
    End If
    msUSV_TYPE = rs("USV_Type") & ""
    msEncounterNumber = rs("Encounter_Number") & ""
        
    'msAPPT_MEPRS_IEN = rs("APPT_MEPRS_IEN") & ""
    msEncounterStatus = rs("Encounter_Status") & ""
    
    'SCR-39491 msPatientUnitNumber retieved from chcs_interface.Patient object when opening encounter.
    msPatientUnitNumber = "" 'rs("Patient_IEN") & "" '****************
    msPatientIEN = rs("Patient_IEN") & ""
    msPatientName = rs("Patient_Name") & ""
    msPatientSSN = rs("Patient_SSN") & ""
    msPatientFMP = rs("Patient_FMP") & ""
    
    'SCR-39058
    msPatientHomePhone = Reformat_Phone(rs("Patient_Home_Phone") & "")
    msPatientWorkPhone = Reformat_Phone(rs("Patient_Work_Phone") & "")
    msSponsorSSN = rs("Sponsor_SSN") & ""
    msClinicNCID = rs("Clinic_IEN") & ""
    msClinicName = rs("Clinic_Name") & ""
    msClinicianNCID = rs("Clinician_IEN") & "" 'per Jason
'/* SCR 30790 - use the provider name from Rs
'    msProviderName = CurrentCHCS_GUI_Provider   'rs("Provider_Name") & ""
    msProviderName = rs("Provider_Name") & ""
'*/ end SCR 30790

    msMeprsCode = rs("Meprs_Code") & ""
    
    If msMeprsCode <> "" Then
        If UCase(Mid(msMeprsCode, 1, 1)) = "A" Then
            mbRelatedToInpatient = True
        Else
            mbRelatedToInpatient = False
        End If
    End If
    
    msMeprsDescription = rs("Meprs_Description") & ""
    msInpatient_MEPRS_Code = rs("Inpatient_MEPRS_CODE") & ""
    If IsNull(rs("Workload_Type")) Then
        msWorkload_Type = ""
    ElseIf rs("Workload_Type") = "" Then
        msWorkload_Type = ""
    Else
        msWorkload_Type = rs("Workload_Type")
    End If
    msOrderNumber = rs("Order_Number") & ""
    msInpatient_Register_Number = rs("REGISTER_NUMBER") & ""
    If msInpatient_Register_Number <> vbNullString Then
        Debug.Print msInpatient_Register_Number
    End If
'    Dim i As Integer
'    For i = 1 To rs.Fields.Count
'        With rs
'            'Debug.Print CStr(rs.Fields.Item(i).Name) + ": " + CStr(rs.Fields.Item(i).Value & "")
'        End With
'    Next

    
    Deserialize_CHCS_GUI = True
    Exit Function
    
ErrorTrap:
    If Err.Number = 3265 Then Resume Next
    mobjShared.ShowVBError Err.Number, Err.Description, "Appointment.cls.Deserialize_CHCS_GUI", "CHCSII_AppointmentClient", vbExclamation
End Function

Public Function Serialize(ByVal sApptID As String, ByVal sFacilityNCID As String) As String
    On Error GoTo ErrHnd
    
    Dim sColumns As String
    Dim sValues As String
    'SCR-26928 S.McAvoy 8/29/02  added the Date_Modified field
    sColumns = "Facility_NCID, Date_Created, Date_Modified"
    sValues = sFacilityNCID & ", " & SystemDateString & ", " & SystemDateString
    AddAllFields sColumns, AddColumn
    AddAllFields sValues, AddValue
    Serialize = "insert into appointment(appt_id, " & sColumns & ") values (" & sApptID & ", " & sValues & " )"
Exit Function

ErrHnd:
    MsgBxARMd "Application Error: " & Err.Description & " Occurred in CHCSII_AppointmentClient.Appointment.Serialize."
    
End Function

Private Sub AddAllFields(rStr As String, ByVal vWhat As AddWhatEnum)
'    AddFieldIfPresent rStr, "Appt_Id", msApptId, vWhat
    AddFieldIfPresent rStr, "Appt_IEN", msApptIEN, vWhat
    AddFieldIfPresent rStr, "Appt_Date_Time", mdApptDateTime, vWhat
    AddFieldIfPresent rStr, "CheckIn_Time", mdCheckInTime, vWhat
    AddFieldIfPresent rStr, "CheckOut_Time", mdCheckOutTime, vWhat
    AddFieldIfPresent rStr, "Appt_Comment", msApptComment, vWhat
    AddFieldIfPresent rStr, "Appt_Reason", msApptReason, vWhat
    AddFieldIfPresent rStr, "Appt_Classification", mnApptClassification, vWhat
    AddFieldIfPresent rStr, "Appt_Type", msApptType, vWhat
    AddFieldIfPresent rStr, "Appt_Status", msApptStatus, vWhat
    AddFieldIfPresent rStr, "Cancel_Reason", msApptCancelReason, vWhat
    AddFieldIfPresent rStr, "Encounter_Number", msEncounterNumber, vWhat
    AddFieldIfPresent rStr, "Encounter_Status", msEncounterStatus, vWhat
    AddFieldIfPresent rStr, "Patient_Unit_Number", msPatientUnitNumber, vWhat
    AddFieldIfPresent rStr, "Clinic_NCID", msClinicNCID, vWhat
    AddFieldIfPresent rStr, "Clinician_NCID", msClinicianNCID, vWhat
    AddFieldIfPresent rStr, "Order_Number", msOrderNumber, vWhat 'DEM 01/03/2000
    AddFieldIfPresent rStr, "Workload_Type", msWorkload_Type, vWhat
    AddFieldIfPresent rStr, "Observation", msObservation, vWhat
    
    'KDunne 3/2/2004
    'ITT Convergence
    'R. Foster SCR #40191; 40140;   Developer: Robert Foster 10/21/2003 11:36 PM
    'R. Foster Note: Function Changed: CHCSII_Appointments.dll\Appointments.frm: FillGrid ; CHCSII_AppointmentClient.dll\ AppointmentOps.cls: Declarations, GetAppointment; Appointment.cls: Deserialize, AddAllFields      Impact of Change: Modifiy SQL statements to return InPatient and Outpatient states
    'ITT UPDATES
    If mobjShared.IsAppMode(modeITT) Then
        AddFieldIfPresent rStr, "Inpatient_MEPRS_CODE", msInpatient_MEPRS_Code, vWhat
    End If
    
    'End ITT Convergence.
    
    AddFieldIfPresent rStr, "Appt_Duration", msAppt_Duration, vWhat
    
    AddFieldIfPresent rStr, "APPT_TYPE_CHCS", msAPPT_TYPE_CHCS, vWhat
    AddFieldIfPresent rStr, "USV_TYPE", msUSV_TYPE, vWhat
    AddFieldIfPresent rStr, "APPT_MEPRS_IEN", msAPPT_MEPRS_IEN, vWhat
    AddFieldIfPresent rStr, "INPATIENT_MEPRS_IEN", msINPATIENT_MEPRS_IEN, vWhat
    AddFieldIfPresent rStr, "PATIENT_CALL_BACK_PHONE", msPatientCallBackPhone, vWhat
    If mobjShared.IsAppMode(modeTheater) Then
        
        'PATIENT_FMP, PATIENT_HOME_PHONE, PATIENT_IEN, PATIENT_NAME, PATIENT_SSN, PATIENT_WORK_PHONE, SPONSOR_SSN,
        AddFieldIfPresent rStr, "PATIENT_FMP", msPatientFMP, vWhat
        AddFieldIfPresent rStr, "PATIENT_HOME_PHONE", msPatientHomePhone, vWhat
        AddFieldIfPresent rStr, "PATIENT_IEN", msPatientIEN, vWhat
        AddFieldIfPresent rStr, "PATIENT_NAME", msPatientName, vWhat
        AddFieldIfPresent rStr, "PATIENT_SSN", msPatientSSN, vWhat
        AddFieldIfPresent rStr, "PATIENT_WORK_PHONE", msPatientWorkPhone, vWhat
        AddFieldIfPresent rStr, "SPONSOR_SSN", msSponsorSSN, vWhat
        AddFieldIfPresent rStr, "TELCON_URGENCY", miPriority, vWhat
        'SCR-51307
        AddFieldIfPresent rStr, "Inj_Ill_Category", msInj_Ill_Category, vWhat
        AddFieldIfPresent rStr, "Inj_Ill_Cause", msInj_Ill_Cause, vWhat
    End If
'APPT_TYPE_CHCS,
'USV_TYPE,
'APPT_MEPRS_IEN ,
'INPATIENT_MEPRS_IEN
        
        
    
End Sub

Private Sub Class_Initialize()
Dim lrv As Long
  Set mobjShared = New CWShared
  'lrv = GetTimeZoneInformation(lTZinfo) ' lTZinfo.Bias = # min between local time and GMT.

End Sub
