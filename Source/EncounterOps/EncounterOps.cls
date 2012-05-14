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
Attribute VB_Name = "EncounterOps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "The main controing class for Encounter Summary."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"EncProperties"
Attribute VB_Ext_KEY = "Member1" ,"EncRTF"
Attribute VB_Ext_KEY = "Member2" ,"EncProperties"


Option Explicit

Private mobjUser       As ICHCSII_User       '-- The Security User Object
Private mobjPatient    As MMMHISPatient.Patient             '-- patient object from core
Private mobjLogon As Logon        '-- logon env obj from core
Private mobjComm       As Object             '-- comm obj from core
Private mobjProperties As EncProperties      '-- the autocite an other stuff from the OSH
Private mobjShared     As MMMHISShared.CWShared
Private mobjConfig     As MMMHISConfiguration.Config
Private mbUseDental    As Boolean
Private mobjAudit      As Audit
Private meSource        As IDMBE.DMLocations
Private mobjCurrEncounter As Encounter
Private mcolTemplateHx As Collection

Private mobjCHCSAppt    As Object
 
Private Enum EncToReturn
  etrAll = 1
  etrOpen = 2
  etrPrev = 3
  etrInpatientNotes = 4  ' EncType = 5
End Enum

' Disconnected Record sets to
' Hold DAS and Encounter info.
Private m_oRSDASEvents As ADODB.Recordset
Private m_oRSEncounters As ADODB.Recordset

'-- implements the dlog.dll timing module
#If debugon Then
  Private DR As DebugRec
  Private DL As cDebugLog
  Private Const Exe = "EncounterOps"
  Private Const Module = "EncounterOps.cls"
#End If
Public Property Get TemplateHx() As Collection
    Set TemplateHx = mcolTemplateHx
End Property
Public Property Set TemplateHx(colHx As Collection)
   Set mcolTemplateHx = colHx
End Property

Public Property Get UseDental() As Boolean:                 UseDental = mbUseDental:                    End Property

Public Property Get Comm() As Object:                       Set Comm = mobjComm:                         End Property
Public Property Get EncProperties() As EncProperties:       Set EncProperties = mobjProperties:          End Property
Public Property Get LogonEnvironment() As Object:           Set LogonEnvironment = mobjLogon: End Property
Public Property Get Patient() As Object:                    Set Patient = mobjPatient:                   End Property
Public Property Get User() As ICHCSII_User:                 Set User = mobjUser:                         End Property

Private Sub TransferClearBadCosigner(objEnc As Encounter)
    If Not objEnc.WhoCoSigned Is Nothing Then 'cosigner object exists
        If objEnc.WhoCoSigned.NCID = 0 Then 'but the data is invalid
            Set objEnc.WhoCoSigned = Nothing
        End If
    End If
End Sub

Private Function TransferCopyEncounter(objSource As Encounter, objDest As Encounter)
'copies encounter properties for transfer from LDDB to CDR
On Error GoTo ErrHandler

    With objDest
        .AlleryVerifiedByDATE = objSource.AlleryVerifiedByDATE
        .AlleryVerifiedByNCID = objSource.AlleryVerifiedByNCID
        'Set .AllListNoteRecords = objsource.AllListNoteRecords
        'Set .AllListToolRecords = objsource.AllListToolRecords
        '.AppointmentID = objsource.AppointmentID
        .APStatus = objSource.APStatus
        '.ArrivedByAmbulance = objsource.ArrivedByAmbulance
        .BypassSODialog = objSource.BypassSODialog
        '.Category = objsource.Category
        .Class = objSource.Class
        '.Clinic = objsource.Clinic
        '.ClinicNCID = objsource.ClinicNCID
        '.Comm = objsource.Comm
        '.Config = objsource.Config
        Set .Cosigner = objSource.Cosigner
        '.CurrentRTF = objsource.CurrentRTF
        'Set .CurrentSONote = objsource.CurrentSONote
        '.CurrSection = objsource.CurrSection
        '.CWShared = objsource.CWShared
        '.DataId = objsource.DataId
        .DentalDispComments = objSource.DentalDispComments
        .DentalDispOptions = objSource.DentalDispOptions
        .DentalDispStatus = objSource.DentalDispStatus
        .DentalDispStatusDesc = objSource.DentalDispStatusDesc
        '.Dirty = objSource.Dirty
        .Dirty = True
        .DispAdminOption = objSource.DispAdminOption
        .DispositionNCID = objSource.DispositionNCID
        .DispositionText = objSource.DispositionText
        '.DispStatus = objsource.DispStatus
        .EandMCalc = objSource.EandMCalc
        .EandMCodeReview = objSource.EandMCodeReview
        .EandMNCID = objSource.EandMNCID
        .EandMReviewDate = objSource.EandMReviewDate
        .EandMReviewed = objSource.EandMReviewed
        .EandMReviewedBy = objSource.EandMReviewedBy
        .EMCalcOverrideNote = objSource.EMCalcOverrideNote
        .EncounterID = objSource.EncounterID
        '.EncounterParent = .EncounterParent
        '.EncounterType = objsource.EncounterType
        '.EncProperties = objsource.EncProperties
        '.EncRTFs = objsource.EncRTFs
        '.EndDTS = objsource.EndDTS
        '.EnterpriseNCID = objsource.EnterpriseNCID
        '.Facility = objsource.Facility
        '.FacilityNCID = objsource.FacilityNCID
        '.FlagNew = objsource.FlagNew
        .FollowupComments = objSource.FollowupComments
        .FollowupTimeFrame = objSource.FollowupTimeFrame
        '.FullyLoaded = objsource.FullyLoaded
        '.IsOpen = objsource.IsOpen
        .DiscussedComments = objSource.DiscussedComments
        .ItemsDiscussed = objSource.ItemsDiscussed
        '.Loading = objsource.Loading
        '.LockingObject = objsource.LockingObject
        '.Logon = objsource.Logon
        '.Patient = objsource.Patient
        .PatientID = objSource.PatientID
        .PrimaryDiagnosisNCID = objSource.PrimaryDiagnosisNCID
        '.PrimaryProvider = objsource.PrimaryProvider
        '.Providers = objsource.Providers
        '.RelatedDiagnoses = objsource.RelatedDiagnoses
        .SADRStatus = objSource.SADRStatus
        '.SectionIsEditable = objsource.SectionIsEditable
        '.SectionParent = objsource.SectionParent
        '.SectionRTF = objsource.SectionRTF
        .Sensitivity = objSource.Sensitivity
        '.ShowDentalSection = objsource.ShowDentalSection
        '.StartDTS = objsource.StartDTS
        .Status = objSource.Status
        '.StorageKey = objsource.StorageKey
        .TemplateData = objSource.TemplateData
        '.TextNotes = objsource.TextNotes
        '.TimeZone = objsource.TimeZone
        '.UpdateClinicianNCID = objsource.UpdateClinicianNCID
        '.UpdateDate = objsource.UpdateDate
        '.User = objsource.User
        '.WasUpdated = objsource.WasUpdated
        .WhenCoSigned = objSource.WhenCoSigned
        .WhenSigned = objSource.WhenSigned
        '.WhoCoSigned = objsource.WhoCoSigned
        '.WhoSigned = objsource.WhoSigned
        .WorkStatusDisposition = objSource.WorkStatusDisposition
        .WorkStatusScreening = objSource.WorkStatusScreening
    End With

    TransferCopyEncounter = True
    Exit Function
ErrHandler:
    TransferCopyEncounter = False
    Resume
End Function



Public Function GetEncounter(ByVal PatientUnitNumber As String, _
                             ByVal FacilityNCID As String, _
                             ByVal EncounterNumber As String) As Encounter
Attribute GetEncounter.VB_Description = "Returns a single (existing) encounter based on patient and encounter numbers. "
    Dim objEncounter As Encounter
    On Error GoTo ErrHandler
    #If debugon Then
        Dim DR As DebugRec
        If Not DL Is Nothing Then DL.DebugLog Exe, Module, "GetEncounter", "Begin", DR
    #End If
    
    '>>> SCR 26339 MSolano 10/10/2002
    goARM.StartCmplxTx gsARMTx_EncID, EncounterNumber, EncounterNumber
    goARM.StopCmplxTx gsARMTx_EncID, enmARMTxStatusOk
    '<<< SCR 26339
    
    If Len(EncounterNumber) = 0 Then Exit Function 'Otherwise a random encounter will be returned SCR-39408
    
    If mobjPatient.UnitNumber <> PatientUnitNumber Then 'SCR-37623, SCR-38132, SCR-38135
        Call mobjComm.Message(cwiSELECT_PATIENT, PatientUnitNumber & ";", NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION)
    End If
    
    'PERF - Check if Encounter to retrieve is the encounter cached SCR-31509
    If Not mobjCurrEncounter Is Nothing Then
        If mobjCurrEncounter.EncounterID = EncounterNumber Then
            Set objEncounter = mobjCurrEncounter
'            If objEncounter.Patient.UnitNumber <> PatientUnitNumber Then
'                objEncounter.Patient.LoadPatient PatientUnitNumber, vbNullString 'Wrong function called SCR-25257
'            End If
             'Set GetEncounter = objEncounter
        End If
    End If
    
    If objEncounter Is Nothing Then
        ResetCurrLoadedEncounter 'SCR#51375, 45212
        Set objEncounter = New Encounter
        objEncounter.Init mobjUser, mobjComm, mobjShared, mobjLogon, mobjPatient, mobjProperties
        
        If Not objEncounter.Load(PatientUnitNumber, FacilityNCID, EncounterNumber) Then
            'Encounter Load Failed SCR#49482
            Set objEncounter = Nothing
        End If
    End If
    
    If Not objEncounter Is Nothing Then
        If mobjShared.IsAppMode(modeCHCSI_GUI) Then
            objEncounter.Patient_IEN = mobjPatient.IEN_number  'SCR-42656
        End If
        
        Set GetEncounter = objEncounter
    End If
    
    GoTo CleanUp
ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description, "EncounterOps.GetEncounter ", "EncounterOps", vbCritical)
CleanUp:
    Set objEncounter = Nothing
#If debugon Then
    If Not DL Is Nothing Then DL.DebugLog Exe, Module, "GetEncounter", "End", DR
#End If
End Function

Private Sub Class_Initialize()
#If debugon Then
    Set DL = GetObject("", "dlog.cdebuglog")
    DL.debugStart
    Dim DR As DebugRec
    If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Class_Initialize", "Begin", DR
#End If
     
    '---create shared object used to show VB errors, etc.---
    Set mobjShared = New CWShared
    Set mobjProperties = New EncProperties
    
    '--- SCR 26339 MSolano 8/8/2002
    Call ARM_InitAllTransactionsForEncounterOps
    
    '---set the help file for this application---
    App.HelpFile = mobjShared.CWHelpFile
    Call InitializeGlobalStrings
    #If debugon Then
        If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Class_Initialize", "End", DR
    #End If
End Sub

Private Sub Class_Terminate()
    Set mobjPatient = Nothing
    Set mobjLogon = Nothing
    Set mobjComm = Nothing
    Set mobjProperties = Nothing
    Set mobjShared = Nothing
    Set mobjConfig = Nothing
    Set mobjAudit = Nothing
    Set mobjUser = Nothing
    Set mcolTemplateHx = Nothing
    
#If debugon Then
    Dim DR As DebugRec
    If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Class_Terminate", "End", DR
    Set DL = Nothing
#End If
    
End Sub

Public Sub Initialize(ByRef robjComm As Object, _
                            robjLogonEnvironment As Object, _
                            robjConfig As Object, _
                            robjPatient As Object, _
                            robjUser As Object)
    Dim sMsg As String
    If Not DebugMode Then On Error GoTo ErrHandler
    Set mobjComm = robjComm
    Set mobjLogon = robjLogonEnvironment
    Set mobjConfig = robjConfig
    Set mobjPatient = robjPatient
    Set mobjUser = robjUser
    Call mobjProperties.Initialize(mobjUser, mobjLogon, mobjComm, mobjShared, mobjConfig)
    mobjProperties.RetrieveProperties
    
    'Change dental priv check SCR#49296
    mbUseDental = mobjUser.HasPrivilegeEx("Dental/Create_Dental_Encounter", Priv_Write)
    
'    sMsg = "USER_HAS_APP_ACCESS|" & NCID_DDSA_APPLICATION
'    If mobjComm.Message(cwiGENERIC, sMsg, NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION) Then
'        mbUseDental = sMsg = "TRUE"
'    End If
    
    'Get global to which DB to use SCR-24819
    If mobjShared.UseSQLServer Then
        geDBUsed = edb_sqlserver
    Else
        geDBUsed = edb_Oracle
    End If
    
    Exit Sub
ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description, "EncounterOps.Initialize", "EncounterOps", vbExclamation)
End Sub

Public Function GetOpenEncounters(ByVal vlUnitNumber As Long) As Collection
    Dim colEncounters As Collection
    Dim i As Integer
    On Error Resume Next
    Set colEncounters = GetPatientEncountersInternal(vlUnitNumber, True)
    For i = colEncounters.Count To 1 Step -1
        Select Case colEncounters(i).Status
        Case complete, NeedsCoSignature, Updated, PatientLeftWithoutBeingSeen, _
        CancelByPatient, CancelByFacility, NoShow, CancelByProvider
             ' remove it
            colEncounters.Remove i
        End Select
    Next i
    Set GetOpenEncounters = colEncounters
End Function

Public Function GetPreviousEncounters(ByVal vlUnitNumber As Long) As Collection
    Dim colEncounters As Collection
    Dim i As Integer
    On Error Resume Next
    Set colEncounters = GetPatientEncountersInternal(vlUnitNumber, True, , etrPrev)
    For i = colEncounters.Count To 1 Step -1
        Select Case colEncounters(i).Status
        Case complete, NeedsCoSignature, Updated, PatientLeftWithoutBeingSeen, CancelByFacility, CancelByPatient, NoShow, Updating, CancelByProvider
            ' we want these
        Case Else
            ' remove it
            colEncounters.Remove i
        End Select
    Next i
    Set GetPreviousEncounters = colEncounters
End Function

Public Function GetEncountersToCosign(Optional ByRef ProviderNCID As String) As Collection
    
    On Error GoTo ErrHandler
    
    If IsMissing(ProviderNCID) Or ProviderNCID = "" Then
        ProviderNCID = mobjLogon.UserNCID
    End If
    
    Set GetEncountersToCosign = GEMS_GetEncountersToCosign(ProviderNCID)
    
    GoTo CleanUp

ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "EncounterOps.GetPreviousEncounters", "EncounterOps", vbCritical

CleanUp:

End Function

Public Function GetPatientEncounters(ByVal vlUnitNumber As Long) As Collection
    Set GetPatientEncounters = GetPatientEncountersInternal(vlUnitNumber, False)
End Function

Public Function GetPreviousEncountersLite(ByVal vlUnitNumber As Long, Optional ByVal ReturnMultiMedia As Boolean = True) As Collection
    Set GetPreviousEncountersLite = GetPatientEncountersInternal(vlUnitNumber, True, ReturnMultiMedia, etrPrev)
End Function

Public Function GetPatientEncountersLite(ByVal vlUnitNumber As Long, Optional ByVal ReturnMultiMedia As Boolean = True) As Collection
    Set GetPatientEncountersLite = GetPatientEncountersInternal(vlUnitNumber, True, ReturnMultiMedia)
End Function
Public Function GetInpatientNoteEncounters(ByVal vlUnitNumber As Long, Optional ByVal ReturnMultiMedia As Boolean = True) As Collection
    Set GetInpatientNoteEncounters = GetPatientEncountersInternal(vlUnitNumber, True, ReturnMultiMedia, etrInpatientNotes)
End Function

Private Function GetPatientEncountersInternal(ByVal vlUnitNumber As Long, ByVal Lite As Boolean, Optional ByVal bReturnMultiMedia As Boolean = True, Optional enumEncToReturn As EncToReturn = etrAll) As Collection

    On Error GoTo ErrHandler
    
        Set GetPatientEncountersInternal = GEMS_GetPatientEncountersInternal(vlUnitNumber, Lite, enumEncToReturn)
        Exit Function

    GoTo CleanUp
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "EncounterOps.GetPreviousEncounters", "EncounterOps", vbCritical

CleanUp:
    
End Function

'-----------------------------------------------------
'Purpose:  Creates a new encounter based on information supplied
'
'Parameters:  AppID - Called by appt so this is pointer back to appt
'             Appttype - In/Out/Amb/Telecons/Dental
'             Facilityncid -
'             ClinicNCID -
'             Clerknote - any note passed by a clerk (used for telecons)
'Returns:  The encounter created
'
'Comments: NOTE: it gets an encounter id from the add method
'----------------------------------------------------
Public Function CreateEncounter(ApptID As String, _
                                ApptType As String, _
                                Enc_Type As EncType, _
                                FacilityNCID As String, _
                                ClinicNCID As String, _
                                PrimaryProviderNCID As String, _
                       Optional ClerkNote As String = "", _
                       Optional bArrivedByAmbulance As Boolean = False, _
                       Optional dApptDTS As Date, _
                       Optional bCheckPriv As Boolean = True, _
                       Optional bResultedConsultOrder As Boolean = False, _
                       Optional sApptReason As String = "", _
                       Optional sApptComment As String = "", _
                       Optional lUnitNumber As Long = 0, _
                       Optional sWorkLoadOptions As String = "", _
                       Optional sWorkLoad As String = "", _
                       Optional bObservation As Boolean = False, _
                       Optional sAMEPRSCode As String = "", _
                       Optional sBMEPRSCode As String = "", _
                       Optional sApptIEN As String = "", _
                       Optional sInpatientService As String = "", _
                       Optional oEncAccidentData As EncAccidentData = Nothing, _
                       Optional bTransferToLDDB As Boolean = False, _
                       Optional InptNoteTypeNCID As String = "", _
                       Optional InptHospDay As Long = 0) As Encounter
                                'SF 56557, 56571

    Dim oOcc As Occurrence
    Dim objEncounter As Encounter
    Dim bCreated As Boolean
    Dim sSql As String

    On Error GoTo ErrHandler

    If mobjShared.IsAppMode(modeCHCSI_GUI) Then
        'In PGUI Appts will always call CreateEncounter since it does not know whether an encounter has been created or not
        'lUnitNumber = 0 when using tree view to open encounter
        If lUnitNumber = 0 Then
            lUnitNumber = mobjPatient.UnitNumber
        End If

        Set objEncounter = GetIfEncounterExists(ApptID, FacilityNCID, mobjPatient.UnitNumber, PrimaryProviderNCID)
        
        If Not objEncounter Is Nothing Then
            Set CreateEncounter = objEncounter
            Set mobjCurrEncounter = objEncounter
            Exit Function
        End If
    End If
    
    ResetCurrLoadedEncounter 'SCR#51375, 45212
    
    ' check that primary provider is legit
    'SCR 38475  Phil Crowder  7/9/03  Skip check for PGUI mode
    If bCheckPriv And Not mobjShared.IsAppMode(modeCHCSI_GUI) Then 'Dont check for transfer to LDDB dont have provider privileges
        If Not UserCanBePrimaryProvider(PrimaryProviderNCID) Then
            MsgBxARMd "The designated primary provider is not authorized to sign encounters."
            Exit Function
        End If
    End If
    
    'Check to make sure the Patient Unit Number for the Appointment matches the Current Patient Unit Number
    'This is to prevent an Encounter being created for the wrong patient SCR-33211
    If lUnitNumber <> 0 Then
        If mobjPatient.UnitNumber <> lUnitNumber Then
            Call mobjComm.Message(cwiSELECT_PATIENT, lUnitNumber & ";", NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION)
            'Check to make sure Select Patient succeeded SCR#51375, 52338, 52339
            If mobjPatient.UnitNumber <> lUnitNumber Then
                'Select Patient Failed  SCR#51375
                MsgBxARMd "Encounter creation failed. Unable to change patients. Please make sure the patient is selected and try again.", vbOKOnly + vbCritical, "Encounters"
                Exit Function
            End If
            'Load Patient does not change core's patient object SCR-33211
            'mobjPatient.LoadPatient lUnitNumber, vbNullString
        End If
    End If
    
    'fixed scr 16105, GEMS has dental
    If Enc_Type <> TeleConsult Then
        If mbUseDental Then
            Enc_Type = Dental
        End If
    End If
    
    Set objEncounter = New Encounter
    objEncounter.Init mobjUser, mobjComm, mobjShared, mobjLogon, mobjPatient, mobjProperties
    
    'date:01/09/2004 SCR #:47870 developer: jrm Description: Added for ITT mode
    'Auto is needed to determine ITT, or Theatre
    
    If mobjShared.IsAppMode(modeCDR) And mobjShared.IsAppMode(modeITT) = False Then 'SCR 36493 SCR-
        Set oOcc = New Occurrence
        oOcc.UnitNumber = mobjPatient.UnitNumber
        oOcc.Facility = FacilityNCID
        oOcc.ArrivalTime = Now
        objEncounter.EncounterID = Val(oOcc.Add)
    Else
        Dim oDas As ICHCSII_SQL.ISqlOps
        Set oDas = gobjCHCSIIConn.CHCSII_SQLOPS(Auto)  '''<SCR 36493
        objEncounter.EncounterID = oDas.GetNextID(esENCOUNTER_ID)
        Set oDas = Nothing
'        'Use Appt ID which is acutally Appt IEN
'        objEncounter.EncounterID = ApptID
    End If
    
    '>>> SCR 26339 MSolano 10/10/2002
    goARM.StartCmplxTx gsARMTx_EncID, objEncounter.EncounterID, objEncounter.EncounterID
    goARM.StopCmplxTx gsARMTx_EncID, enmARMTxStatusOk
    '<<< SCR 26339

    If objEncounter.EncounterID = 0 Then
        Set objEncounter = Nothing
        MsgBxARMd "Adding Encounter Occurrence to CDR produced no error, but generated no encounter id.", vbCritical, "EncounterOps"
        GoTo CleanUp
    End If
    objEncounter.PatientID = mobjPatient.UnitNumber
    objEncounter.ArrivedByAmbulance = bArrivedByAmbulance
    objEncounter.ApptDTS = dApptDTS 'SCR-17624
    objEncounter.ResultedConsultOrder = bResultedConsultOrder
    objEncounter.ApptIEN = sApptIEN 'SCR-31315
    objEncounter.InpatientService = sInpatientService
    
    If bTransferToLDDB Then
        'Transferred from LDDB 'SCR#53547
        objEncounter.Status = complete
    End If
    
    If Len(mobjComm.CmdLineSwitch("WAM")) Then
        objEncounter.WorkLoad = sWorkLoad
        objEncounter.WorkLoadOptions = sWorkLoadOptions
        If bObservation Then
            objEncounter.PatientStatus = "Outpatient Observation"
        ElseIf Enc_Type = Ambulatory Then
            objEncounter.PatientStatus = "Outpatient APV"  'SCR-35186
        ElseIf Enc_Type = Inpatient Then
            objEncounter.PatientStatus = "Inpatient"
        Else
            objEncounter.PatientStatus = "Outpatient"
        End If
        objEncounter.MEPRSCodes = sAMEPRSCode & "|" & sBMEPRSCode
    End If
    
    'Add Default template functionality SCR-22258 Eanble for PG mode SCR-35673
    If Len(mobjComm.CmdLineSwitch("GREEN")) = 0 Or mobjShared.IsAppMode(modeCHCSI_GUI) Then
        'When the Green switch is removed remove the code in this block.
        objEncounter.TemplateData = GetDefaultTemplateData
    End If
    
    'date:01/09/2004 SCR #:47870 developer: jrm Description: Added for ITT mode
    'Auto is needed to determine ITT, or Theatre
    
    If mobjShared.IsAppMode(modeCDR) And mobjShared.IsAppMode(modeITT) = False Then 'SCR 36493
        If Not oEncAccidentData Is Nothing Then objEncounter.AccidentData.InitFromCopy oEncAccidentData
        If Not objEncounter.Create(mobjPatient, ApptID, ApptType, Enc_Type, FacilityNCID, ClinicNCID, PrimaryProviderNCID, ClerkNote) Then
            Set objEncounter = Nothing
            GoTo CleanUp
        Else
            bCreated = True
        End If
    Else
        If Not objEncounter.GEMS_Create(mobjPatient, ApptID, ApptType, _
                                        Enc_Type, FacilityNCID, ClinicNCID, _
                                        PrimaryProviderNCID, ClerkNote, dApptDTS, , InptNoteTypeNCID, InptHospDay) Then
            
            Set objEncounter = Nothing
            GoTo CleanUp
        End If
    End If
    
    If Not mobjShared.IsAppMode(modeCHCSI_GUI) Then 'SCR 36493
        Call objEncounter.SaveApptNotes(sApptReason, sApptComment) 'SCR-26496
        'PERF - Cache newly created Encounter SCR-31509
        Set mobjCurrEncounter = objEncounter
    End If
    
    GoTo CleanUp
ErrHandler:
   Call mobjShared.ShowVBError(Err.Number, Err.Description, "EncounterOps.CreateEncounter", "EncounterOps", vbExclamation)
CleanUp:
    Set CreateEncounter = objEncounter
    Set objEncounter = Nothing
    Set oOcc = Nothing
End Function

'-----------------------------------------------------
'Purpose:  Creates a new encounter based on information supplied
'
'Parameters:  AppID - Called by appt so this is pointer back to appt
'             Appttype - In/Out/Amb/Telecons/Dental
'             Facilityncid -
'             ClinicNCID -
'             Clerknote - any note passed by a clerk (used for telecons)
'Returns:  The encounter created
'
'Comments: NOTE: it gets an encounter id from the add method
'----------------------------------------------------
Public Function GEMS_CreateEncounter(ApptID As String, _
                                ApptType As String, _
                                Enc_Type As EncType, _
                                FacilityNCID As String, _
                                ClinicNCID As String, _
                                PrimaryProviderNCID As String, _
                       Optional ClerkNote As String = "", _
                       Optional ApptDTS As Date, _
                       Optional sApptReason As String = "", _
                       Optional sApptComment As String = "") As Encounter
    Dim objEncounter As Encounter
    Dim oIns As Object
    Dim oDas As ICHCSII_SQL.ISqlOps
    Set oDas = gobjCHCSIIConn.CHCSII_SQLOPS(Auto)  '''<SCR 36493
    
    On Error GoTo ErrHandler
    
    Set objEncounter = New Encounter
    objEncounter.Init mobjUser, mobjComm, mobjShared, mobjLogon, mobjPatient, mobjProperties
    'If mobjShared.IsAppMode(modeCHCSI_GUI) Then
    '    objEncounter.EncounterID = ApptID
    'Else
        objEncounter.EncounterID = oDas.GetNextID(esENCOUNTER_ID)
    'End If
    objEncounter.PatientID = mobjPatient.UnitNumber
    
    If Not objEncounter.GEMS_Create(mobjPatient, ApptID, ApptType, _
                                    Enc_Type, FacilityNCID, ClinicNCID, _
                                    PrimaryProviderNCID, ClerkNote, ApptDTS) Then
        
        Set objEncounter = Nothing
        GoTo CleanUp
    End If
    
    Call objEncounter.SaveApptNotes(sApptReason, sApptComment)
        
    GoTo CleanUp
ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description, "EncounterOps.CreateEncounter", "EncounterOps", vbExclamation)
CleanUp:
    Set oDas = Nothing
    Set GEMS_CreateEncounter = objEncounter
    Set objEncounter = Nothing
    Set oIns = Nothing

End Function

Public Function UserCanBePrimaryProvider(ByRef UserNCID As String) As Boolean
    
    Dim sLogin As String
    
    If mobjShared.IsAppMode(modeTheater) Then
       sLogin = GEMS_GetLogonFromNCID(UserNCID)
    End If
    
    'Allow UNASSIGNED to create encounters
    'Check if user has Need CoSign Priv SCR44923
    If UserHasSignPrivilege(mobjComm, UserNCID) Or UserHasNeedCoSignPrivilege(mobjComm, UserNCID) Or InStr(1, UCase(sLogin), "UNASSIGNED", vbTextCompare) Then
        UserCanBePrimaryProvider = True
    End If
    
End Function

Private Function GEMS_GetPatientEncountersInternal(ByVal vlUnitNumber As Long, ByVal Lite As Boolean, Optional enumEncToReturn As EncToReturn = etrAll) As Collection
    Dim colEncounters   As Collection
    Dim objEncounter    As Encounter         '- an encounter object
    Dim colDataStrings  As Collection
    Dim objDAS          As ICHCSII_DAS
    Dim sSql            As String
    Dim objRS           As ADODB.Recordset
    Dim oEncProv        As EncProvider
    Dim sMsgText        As String
    Dim sApptIEN        As String
    Dim bFound          As Boolean
    Dim sbuf()          As String
    Dim sClinicianNCID  As String
    Dim objAppts        As Object
    Dim objClinicOps    As Object
    
    
    On Error GoTo ErrHandler

    'date:01/09/2004 SCR #:47870 developer: jrm Description: Added for ITT mode
    'Auto is needed to determine ITT, or Theatre

    Set objDAS = gobjCHCSIIConn.CHCSII_DAS(Auto) '''<SCR 36493
    
    If enumEncToReturn = etrPrev Then
        'PERF Return Prev Enocunters only SCR-40046
        ' Statuses needed by Prev Encounter
        'Case complete, NeedsCoSignature, Updated, PatientLeftWithoutBeingSeen, _
        '     CancelByFacility, CancelByPatient, NoShow, Updating, CancelByProvider
        sSql = "select EncounterNumber, FacilityNCID from Encounters where unit_number = " & vlUnitNumber _
                  & " AND STATUS IN ( 14510306, 14510309, 205657, 14510310, 14508309, 14505856, 14613990, 665, 160595 ) and TYPE <> " & InpatientNote
    ElseIf enumEncToReturn = etrInpatientNotes Then
        sSql = "select EncounterNumber, FacilityNCID from Encounters where unit_number = " & vlUnitNumber & " and TYPE = " & InpatientNote
    Else
        'Return all encounters SCR#43002
        sSql = "select EncounterNumber, FacilityNCID from Encounters where unit_number = " & vlUnitNumber
    End If
    
    If mobjShared.IsAppMode(modeCHCSI_GUI) Then
        'PERF do not fully load encounters in PGUI mode SCR-40046
        Lite = True
    End If
    
    Set objRS = objDAS.OpenRecordset(sSql)
    
    Set colEncounters = New Collection
    Do While Not objRS.EOF
            Set objEncounter = New Encounter
            objEncounter.Init mobjUser, mobjComm, mobjShared, mobjLogon, mobjPatient, mobjProperties
            If objEncounter.GEMS_LoadEncounter(vlUnitNumber, objRS("facilityncid"), objRS("EncounterNumber"), Lite) Then
                On Error Resume Next
                colEncounters.Add objEncounter, CStr(objEncounter.StorageKey)
                On Error GoTo ErrHandler
            End If
        
        objRS.MoveNext
    Loop   ' while not objsql.eof
    
    If mobjShared.IsAppMode(modeCHCSI_GUI) Then
        'see if appts is open
        If mobjComm.Message(cwiOLE_SERVER_RUNNING, NCID_CLINIC_SCHEDULE_APPLICATION, NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION) Then
        
            'Get Appt ID from Appointments
            'Send Generic Msg to Appts to obtain Appt IEN, Appt Type and Appt DTS SCR-31496
            sMsgText = "CHCS_GUI_APPT_GRIDINFO"
            mobjComm.Message cwiGENERIC, sMsgText, NCID_CLINIC_SCHEDULE_APPLICATION, NCID_ENCOUNTER_APPLICATION
            '    0       1          2         3          4                5                     6
            'ApptIEN|ApptType|ApptDateTime|ClinicName|ClinicIEN|Inpatient_MEPRS_Code|CStr(enumEncounterType)
            If Len(sMsgText) Then
                sbuf = Split(sMsgText, "|")
    
                sApptIEN = sbuf(0)
    
    
                'See if currently selected appt is in the collection
                For Each objEncounter In colEncounters
                    If objEncounter.AppointmentID = sApptIEN Then
                        'Found encounter
                        bFound = True
                        Exit For
                    End If
                Next
        
                Set objEncounter = Nothing
        
                'Not Found
                If Not bFound Then
                    'Get CHCSAppointment
                    Set objAppts = mobjComm.InitializeOLEServer(NCID_CLINIC_SCHEDULE_APPLICATION)
        
                    Set mobjCHCSAppt = objAppts.AppointmentOps.GetAppointment_CHCS_GUI(sApptIEN)
                    
                    'Only add dummy encounter if completed SCR-43463
                    If InStr(1, mobjCHCSAppt.ADM_ENCOUNTER, "COMPLETE", vbTextCompare) Then
                        Set objClinicOps = CreateObject("CHCSII_ClinicClient.ClinicOps")
        
                        sClinicianNCID = objClinicOps.GetProviderNCID(mobjCHCSAppt.FacilityNCID, mobjCHCSAppt.ClinicianNCID)
            
                        'Create dummy encounter for Prev Encounters
                        Set objEncounter = New Encounter
                        objEncounter.Init mobjUser, mobjComm, mobjShared, mobjLogon, mobjPatient, mobjProperties
                        'objEncounter.EncounterID = 99999999
                        'Enc # start at 99990287  SCR#46162
                        objEncounter.EncounterID = 11111111
                        
                        objEncounter.GEMS_Create mobjPatient, sApptIEN, mobjCHCSAppt.ApptType, Outpatient, mobjCHCSAppt.FacilityNCID, _
                                                mobjLogon.NursingDivisionNCID, sClinicianNCID, , mobjCHCSAppt.ApptDateTime, True
        
                        ' Add Dummy RTF
                        Dim objEncRTF As EncRTF
                        Set objEncRTF = New EncRTF
                        'objEncRTF.msUserNCID = sClinicianNCID
                        objEncRTF.msRTF = "{\rtf1\ansi\ucl\deff0\deflang1033\deflangfe1033" _
                            & "{\fonttbl{\f0\fswiss\fcharset0\fprq2 Arial;}}" _
                            & "{\colortbl;\red0\green0\blue0;}" _
                            & "\pard\plain\f0\fs16\b " _
                            & " This Encounter was not completed in Provider GUI." & "\b0\par } "
                        objEncounter.EncRTFs.Add objEncRTF
                        
                        objEncounter.Status = complete
                        
                        On Error Resume Next
                        colEncounters.Add objEncounter, objEncounter.StorageKey
                        On Error GoTo ErrHandler
                    End If
                End If
            End If
        End If

    End If
    GoTo CleanUp
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "EncounterOps.GetPreviousEncounters", "EncounterOps", vbCritical
    Set colEncounters = New Collection
CleanUp:
    Set GEMS_GetPatientEncountersInternal = colEncounters
    Set colEncounters = Nothing
    Set objEncounter = Nothing
    If objRS.State > 0 Then objRS.Close
    Set objRS = Nothing
    Set objDAS = Nothing
End Function


Private Function GEMS_GetEncountersToCosign(ByRef ProviderNCID As String) As Collection
    Dim colEncounters   As Collection
    Dim objEncounter    As Encounter    '- an encounter object
    Dim objDAS          As ICHCSII_DAS
    Dim objConcept      As GEMS_ConceptCtrl  '- converter of ncids to prepresentation 2000
    Dim colDataStrings  As Collection
    Dim objRS           As ADODB.Recordset
    Dim sSql            As String
    On Error GoTo ErrHandler
    If IsMissing(ProviderNCID) Or ProviderNCID = "" Then
        ProviderNCID = mobjLogon.UserNCID
    End If
    Set objConcept = New GEMS_ConceptCtrl
    Set objDAS = gobjCHCSIIConn.CHCSII_DAS(Auto) '''<SCR 36493
    sSql = "Select unit_number, facilityNCID, encounterNumber from encounters " _
         & " where status = " & StatusEnumToNCID(NeedsCoSignature) _
         & " and CosignerNCID = " & ProviderNCID
    Set objRS = objDAS.ExecuteSQL(sSql)
    
    Set colEncounters = New Collection
    Do While Not objRS.EOF
            Set objEncounter = New Encounter
            objEncounter.Init mobjUser, mobjComm, mobjShared, mobjLogon, mobjPatient, mobjProperties
            If objEncounter.GEMS_LoadEncounter(objRS("Unit_Number"), objRS("facilityncid"), objRS("EncounterNumber"), True) Then
                On Error Resume Next
                colEncounters.Add objEncounter, CStr(objEncounter.StorageKey)
                On Error GoTo ErrHandler
            End If
        
        objRS.MoveNext
    Loop   ' while not objsql.eof
    GoTo CleanUp
    
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "EncounterOps.GetPreviousEncounters", "EncounterOps", vbCritical
    Set colEncounters = New Collection
CleanUp:
    Set GEMS_GetEncountersToCosign = colEncounters
    Set colEncounters = Nothing
    Set objEncounter = Nothing
    If objRS.State > 0 Then objRS.Close
    Set objRS = Nothing
    Set objRS = Nothing
    Set objDAS = Nothing
    
End Function

'Check to see if user has access to view senstive encounter
'Primary Provider and Cosigner can view regardless of access
Public Function ViewSensitiveEncounter(Encounter As Encounter, Optional bSupressMsg As Boolean = False, Optional sPriv As String = vbNullString) As Boolean

    Dim sMsg As String
    Dim sSecurityObject As String
    
    On Error GoTo ErrHandler

    ViewSensitiveEncounter = True
    'Check if encounter is sensitive
    sMsg = "Checking Encounter Sensitivity"
    If Encounter.Sensitivity Then
        'Check to see if user is Primary Provider
        sMsg = "Checking to see if user if Primary Provider"
        If Not Encounter.UserIsPrimaryProvider Then
            'Check to see if user is CoSigner
            sMsg = "Checking to see if User is Cosigner"
            If Not Encounter.UserIsCosigner Then
    
                If Len(sPriv) = 0 Then
                    sSecurityObject = Priv_Data_Sensitive_Encounters
                    'Get sensitive data security object 'SCR-35365
                    Select Case Encounter.Sensitivity
                        Case esrStandard
                            sSecurityObject = Priv_Data_Sensitive_Encounters
                        Case esrBehavioralHealth
                            sSecurityObject = Priv_Data_Behavioral_Health_Record
                        Case esrFamilyAdvocacy
                            sSecurityObject = Priv_Data_Family_Advocacy_Record
                    End Select
                Else
                    sSecurityObject = sPriv
                End If
                
                'Check if user has "Break the Glass" priv
                sMsg = "Checking user privileges"
                If Not mobjUser.HasPrivilegeEx(sSecurityObject, Priv_Break_The_Glass) Then
                
                    'Check if user has read privilege
                    If Not mobjUser.HasPrivilegeEx(sSecurityObject, Priv_Read) Then
                    
                        'User does not have access to sensitive encounters
                        If Not bSupressMsg Then
                            MsgBxARMd "You do not have access to view this sensitive encounter.", vbOKOnly, "Security"
                        End If
                        ViewSensitiveEncounter = False
                   End If
                Else
                    'User has "Break the Glass" privilege
                    If Not bSupressMsg Then
                        If MsgBxARMd("You have requested to view sensitive data. You will be audited if you wish to proceed. Proceed?", vbYesNo, "Security") = vbYes Then
                            'Reload encounter to audit viewing of encounter
                            If mobjAudit Is Nothing Then
                                Set mobjAudit = New Audit
                            End If
                            
                            sMsg = "Turning on Audit"
                            mobjAudit.TurnAuditOn
                            GetEncounter Encounter.PatientID, Encounter.FacilityNCID, Encounter.EncounterID
                            mobjAudit.TurnAuditOff
                        Else
                            'User chose not to be audited
                            ViewSensitiveEncounter = False
                        End If
                    Else
                        ViewSensitiveEncounter = False
                    End If
                    
                End If
                
            End If
        End If
    End If

    Exit Function
ErrHandler:
    
    Call mobjShared.ShowVBError(Err.Number, sMsg & " - " & Err.Description, "EncounterParent.ViewSensitiveEncounter", "EncounterOps", vbCritical)
    Exit Function
    
    Resume
End Function
 
Public Function TransferPIC2LDDB(oDMBE As IDMBE.I_DMBE) As Boolean

Dim Rs          As ADODB.Recordset
Dim RSObject    As IDMBE.I_RSObjects
Dim RSEncounters As IDMBE.I_RSObjects
Dim RSEncSections As IDMBE.I_RSObjects
Dim RSTextNotes As IDMBE.I_RSObjects
Dim RSListToolRecord As IDMBE.I_RSObjects
Dim RSListToolRecords As IDMBE.I_RSObjects
Dim RSEncDisposition As IDMBE.I_RSObjects
Dim rsInpatientAdmission As IDMBE.I_RSObjects
Dim oEncounter As Encounter
Dim oSQL       As ICHCSII_SQL.ISqlOpsEx
Dim Field As ADODB.Field
Dim sPICEncID     As String
Dim sLDDBEncID      As String
Dim sFacilityNCID   As String
Dim sUnitNumber As String
Dim oProblems As I_RSTransfer
Dim oVitals As I_RSTransfer
Dim oMeds As I_RSTransfer
Dim oPAD As I_RSTransfer
Dim cCol As Collection
Dim lngRet As Long
Dim objEncParent As Object
Dim sProvNCID As String
Dim sTempNCID As String
Dim dEncDate As Date
Dim sSql As String
Dim sNewApptId As String

    On Error GoTo ErrHandler
        
    meSource = oDMBE.Source
    
    Set cCol = New Collection
    
    Call oDMBE.Execute(cCol, lngRet)
    
    If cCol.Count = 0 Then
        Exit Function
    End If
    
    Set oSQL = gobjCHCSIIConn.CHCSII_SQLOPS(Auto)
    
    For Each RSObject In cCol
        With RSObject
            If RSObject.DBRecordSet.RecordCount > 0 Then
                RSObject.DBRecordSet.MoveFirst
            End If
            
            If .TableName = "ENCOUNTERS" Then
                Set RSEncounters = RSObject
            ElseIf .TableName = "ENC_SECTIONS" Then
                Set RSEncSections = RSObject
            ElseIf .TableName = "ENC_RTFS" Then
                Set RSTextNotes = RSObject
            ElseIf .TableName = "LISTTOOLRECORD" Then
                Set RSListToolRecord = RSObject
            ElseIf .TableName = "LISTTOOLRECORDS" Then
                Set RSListToolRecords = RSObject
            ElseIf .TableName = "INPATIENT_ADMISSION" Then
                Set rsInpatientAdmission = RSObject
            
            End If
        End With
    Next
    
    '******************************************************
    'SF 6-16-05
    'Go through each Inpatient_Admission record and determine:
    '1. Does the record exist in the destination db
    '2. If it exists, then map the incoming ApptID to the existing AppID
    '3. If it does not exist, get a new ApptID
    If Not rsInpatientAdmission Is Nothing Then
        Do Until rsInpatientAdmission.DBRecordSet.EOF
            
            'Does the InpatientAdmission exist in the dest DB?
            sSql = "Select * from Inpatient_Admission" _
                & " where Facility_NCID = " & rsInpatientAdmission.DBRecordSet.Fields("Facility_NCID") _
                & " and Unit_Number = " & mobjPatient.UnitNumber _
                & " and Admission_Date = " & gobjCHCSIIConn.SQLDate(rsInpatientAdmission.DBRecordSet.Fields("Admission_Date"))
            
            oSQL.Execute sSql
            If oSQL.RecordCount = 0 Then
                  'Get a new Appt_ID
                  sNewApptId = oSQL.GetNextID(exAPPT_ID)
            Else
                  sNewApptId = oSQL("Appt_ID")
            End If
                          
            'Add the mapping
            oDMBE.IDMappers.Add NCID_ENCOUNTER_APPLICATION, "APPT_ID", rsInpatientAdmission.DBRecordSet("Appt_ID"), sNewApptId
            
            rsInpatientAdmission.DBRecordSet.MoveNext
        Loop
        sNewApptId = "" 'Clean up
    End If
    '*****************************************************
    
    '<< Begin:SCR #36826;   Developer: Brian Mowbray 06/04/2003 11:28 PM
    'This will prevent error when there are no encounter records. This only occurs when
    'the DISK option is used and several patients are selected at once and all Data Manager Options
    'are selected even when some patients may not have any encounter data
    If Not RSEncounters Is Nothing Then
      

      Do Until RSEncounters.DBRecordSet.EOF
      
          'Save encounter data
          With RSEncounters
              'Validate Facility & Clinic NCID
              Call oSQL.ValidateFacility(.DBRecordSet("FACILITYNCID"), True)
              
              Call oSQL.ValidateFacility(.DBRecordSet("FACILITYNCID"), True, .DBRecordSet("CLINICNCID"))
                                        
              sProvNCID = .DBRecordSet("PRIMARYPROVIDERNCID") & vbNullString
              If oDMBE.IDMappers.Exists(NCID_PROVIDER_SERVICES, "NCID", sProvNCID) Then
                sProvNCID = oDMBE.IDMappers.Find(NCID_PROVIDER_SERVICES, "NCID", sProvNCID)
              End If
              
              'The Provider should have already been created with Valid Data (ie. First Name, Last Name, SSN)
              Call oSQL.ValidateProvider2(sProvNCID, "", "", "", "", True)
              
              If .DBRecordSet("PRIMARYPROVIDERNCID") & vbNullString <> sProvNCID Then
                oDMBE.IDMappers.Add NCID_PROVIDER_SERVICES, "NCID", .DBRecordSet("PRIMARYPROVIDERNCID") & vbNullString, sProvNCID
              End If
              
              sPICEncID = "" & .DBRecordSet("ENCOUNTERNUMBER")
              sFacilityNCID = "" & .DBRecordSet("FACILITYNCID")
              sUnitNumber = mobjPatient.UnitNumber     '   " & .DBRecordSet("UNIT_NUMBER")
              
              'test to see if this enncounter has already been added
              dEncDate = .DBRecordSet("STARTDTS").value
              sSql = "select EncounterNumber from encounters where STARTDTS = " & gobjCHCSIIConn.SQLDate(dEncDate) & " AND UNIT_NUMBER = " & mobjPatient.UnitNumber
                   
              oSQL.Execute sSql
              
              If oSQL.RecordCount > 0 Then
                'This encounter already exist for this person
                oDMBE.IDMappers.Add NCID_ENCOUNTER_APPLICATION, "DATA_ID", sPICEncID, oSQL("EncounterNumber")
                oDMBE.IDMappers.Add NCID_ENCOUNTER_APPLICATION, "ENCOUNTER_ID", sPICEncID, oSQL("EncounterNumber")
                'SCR 136325
                sLDDBEncID = oSQL("EncounterNumber")
                oSQL.Execute ("Select * from Encounters where EncounterNumber = " & sLDDBEncID)
                GoTo Update
                'GoTo NextRecord
              End If
  
                'SF 6-16-05 **********************
                'Get the new appt id if it is mapped from above
                If .DBRecordSet("Type") = 5 Then 'Inpatient Note
                    'Get the appointment id translation
                    If oDMBE.IDMappers.Exists(NCID_ENCOUNTER_APPLICATION, "APPT_ID", .DBRecordSet("ApptID")) Then
                        sNewApptId = oDMBE.IDMappers.Find(NCID_ENCOUNTER_APPLICATION, "APPT_ID", .DBRecordSet("ApptID"))
                    Else
                        'This will happen if an admission does not exist for the note. "broken admission link"
                        'Therefore, don't import the note because the user can't view it without the admission
                        GoTo NextRecord
                    End If
                Else
                    Dim catType As String
                    'Get a new Appt_ID
                    sNewApptId = oSQL.GetNextID(exAPPT_ID)
                    
                    
                    '--- SCR 133860  Sherry Wang   7/3/2007
                    '--- Create a New Appointment
                    sSql = "insert into appointment " & vbNewLine
                    sSql = sSql & "( CLINICIAN_NCID, FACILITY_NCID, " & vbNewLine
                    sSql = sSql & "APPT_ID , CLINIC_NCID, PATIENT_UNIT_NUMBER, " & vbNewLine
                    sSql = sSql & "APPT_DATE_TIME, APPT_TYPE, CHECKIN_TIME, APPT_COMMENT, " & vbNewLine
                    sSql = sSql & "APPT_CLASSIFICATION , ENCOUNTER_STATUS, DATE_CREATED, " & vbNewLine
                    sSql = sSql & "PATIENT_NAME,  APPT_STATUS) values (" & vbNewLine
                    sSql = sSql & sProvNCID & ", " & mobjLogon.FacilityNCID & ", " & sNewApptId & ", " & vbNewLine
                    sSql = sSql & mobjLogon.NursingDivisionNCID & ", " & vbNewLine
                    sSql = sSql & mobjPatient.UnitNumber & ", " & gobjCHCSIIConn.SQLDate(dEncDate) & ", " & vbNewLine
                    sSql = sSql & "'WI', sysdate, 'Rapid Patient Generation', 0, 4, sysdate, " & vbNewLine
                    sSql = sSql & oSQL.SQLQuote(mobjPatient.Name) & ", 'KEPT')"
                                    'SCR 163399
                    oSQL.Execute sSql
                    
                    'Add the mapping
                    oDMBE.IDMappers.Add NCID_ENCOUNTER_APPLICATION, "APPT_ID", .DBRecordSet("ApptID"), sNewApptId
                End If
                
                '*************************
              
              'Create Encounter in LDDB
              'Creates row in Encounters
              'If Appt Date exists then pass to Create Encounter SCR#53602
              If IsNull(.DBRecordSet("APPTDTS")) Then
                'Does not exist
                Set oEncounter = CreateEncounter(sNewApptId, .DBRecordSet("APPTTYPE"), .DBRecordSet("TYPE"), _
                                                .DBRecordSet("FACILITYNCID"), .DBRecordSet("CLINICNCID"), sProvNCID, , , , False, , , , , , , , , , , , , True)
                                                
              Else
                'Exists
                Set oEncounter = CreateEncounter(sNewApptId, .DBRecordSet("APPTTYPE"), .DBRecordSet("TYPE"), _
                                                .DBRecordSet("FACILITYNCID"), .DBRecordSet("CLINICNCID"), sProvNCID, , , .DBRecordSet("APPTDTS"), False, , , , , , , , , , , , , True)
                                                            
              End If
              
              oEncounter.Save
              oSQL.Execute ("Select * from Encounters where EncounterNumber = " & oEncounter.EncounterID)
  
              sLDDBEncID = oEncounter.EncounterID
              'SCR 136325
Update:
              'Update all the other fields in the encounter table
              For Each Field In .DBRecordSet.Fields
                  If Field.Name <> "ENCOUNTERNUMBER" And Field.Name <> "FACILITYNCID" _
                    And Field.Name <> "UNIT_NUMBER" And Field.Name <> "PRIMARYPROVIDERNCID" And Field.Name <> "APPTID" Then
                      oSQL.value(Field.Name) = .DBRecordSet(Field.Name)
                  End If
              Next
              
              'MAP THE NCID FIELDS
              'AdditionalProvider1NCID
              sTempNCID = oSQL.value("ADDITIONALPROVIDER1NCID") & vbNullString
              If (sTempNCID <> vbNullString) And (sTempNCID <> "0") Then
                If oDMBE.IDMappers.Exists(NCID_PROVIDER_SERVICES, "NCID", sTempNCID) Then
                  oSQL.value("ADDITIONALPROVIDER1NCID") = oDMBE.IDMappers.Find(NCID_PROVIDER_SERVICES, "NCID", sTempNCID)
                End If
              End If
              
              'AdditionalProvider2NCID
              sTempNCID = oSQL.value("ADDITIONALPROVIDER2NCID") & vbNullString
              If (sTempNCID <> vbNullString) And (sTempNCID <> "0") Then
                If oDMBE.IDMappers.Exists(NCID_PROVIDER_SERVICES, "NCID", sTempNCID) Then
                  oSQL.value("ADDITIONALPROVIDER2NCID") = oDMBE.IDMappers.Find(NCID_PROVIDER_SERVICES, "NCID", sTempNCID)
                End If
              End If
              
              'WhoSignedNCID
              sTempNCID = oSQL.value("WHOSIGNEDNCID") & vbNullString
              If (sTempNCID <> vbNullString) And (sTempNCID <> "0") Then
                If oDMBE.IDMappers.Exists(NCID_PROVIDER_SERVICES, "NCID", sTempNCID) Then
                  oSQL.value("WHOSIGNEDNCID") = oDMBE.IDMappers.Find(NCID_PROVIDER_SERVICES, "NCID", sTempNCID)
                End If
              End If

              'WhoCoSignedNCID
              sTempNCID = oSQL.value("WHOCOSIGNEDNCID") & vbNullString
              If (sTempNCID <> vbNullString) And (sTempNCID <> "0") Then
                If oDMBE.IDMappers.Exists(NCID_PROVIDER_SERVICES, "NCID", sTempNCID) Then
                  oSQL.value("WHOCOSIGNEDNCID") = oDMBE.IDMappers.Find(NCID_PROVIDER_SERVICES, "NCID", sTempNCID)
                End If
              End If

              'CoSignerNCID
              sTempNCID = oSQL.value("COSIGNERNCID") & vbNullString
              If (sTempNCID <> vbNullString) And (sTempNCID <> "0") Then
                If oDMBE.IDMappers.Exists(NCID_PROVIDER_SERVICES, "NCID", sTempNCID) Then
                  oSQL.value("COSIGNERNCID") = oDMBE.IDMappers.Find(NCID_PROVIDER_SERVICES, "NCID", sTempNCID)
                End If
              End If
              
              'CreatedBy
              sTempNCID = oSQL.value("CREATEDBY") & vbNullString
              If (sTempNCID <> vbNullString) And (sTempNCID <> "0") Then
                If oDMBE.IDMappers.Exists(NCID_PROVIDER_SERVICES, "NCID", sTempNCID) Then
                  oSQL.value("CREATEDBY") = oDMBE.IDMappers.Find(NCID_PROVIDER_SERVICES, "NCID", sTempNCID)
                End If
              End If
              
              'UpdatedBy
              sTempNCID = oSQL.value("UPDATEDBY") & vbNullString
              If (sTempNCID <> vbNullString) And (sTempNCID <> "0") Then
                If oDMBE.IDMappers.Exists(NCID_PROVIDER_SERVICES, "NCID", sTempNCID) Then
                  oSQL.value("UPDATEDBY") = oDMBE.IDMappers.Find(NCID_PROVIDER_SERVICES, "NCID", sTempNCID)
                End If
              End If
              
              oSQL.value("UPDATELOCK") = CInt(True)
              'oSQL.value("") = .DBRecordSet("")
  
              oSQL.Update
              
              'SCR 136324 - LCI - 07/21/2007: Need to sync APPT_TYPE with Enc. Category after mapping complete
              'SCR 136325
              oSQL.Execute ("Select Category from Encounters where EncounterNumber = " & sLDDBEncID)
              catType = oSQL.value("Category")
              If Not sNewApptId = "" Then
                oSQL.Execute ("Select APPT_TYPE from Appointment where Appt_ID = " & sNewApptId)
                'SCR 163399
                If oSQL.RecordCount > 0 Then
                    oSQL.value("APPT_TYPE") = catType
                    oSQL.Update
                End If
              End If
              Set oEncounter = Nothing
              
          End With 'RSEncounters
          
          Set oEncounter = GetEncounter(sUnitNumber, sFacilityNCID, sLDDBEncID)
          'Set current encounter in EncounterParent so Problems has access to it.
  '        If objEncParent Is Nothing Then
  '            Set objEncParent = mobjComm.InitializeOLEServer(NCID_ENCOUNTER_APPLICATION)
  '        End If
          'Call objEncParent.SetCurrentEncounter(sUnitNumber, sFacilityNCID, sLDDBEncID)
          
          'Set oEncounter = objEncParent.CurrentEncounter
          
          oDMBE.IDMappers.Add NCID_ENCOUNTER_APPLICATION, "DATA_ID", sPICEncID, oEncounter.EncounterID
          oDMBE.IDMappers.Add NCID_ENCOUNTER_APPLICATION, "ENCOUNTER_ID", sPICEncID, oEncounter.EncounterID
          
          'Transfer Encounter Sections
          ' SCR #30970;   Developer: Mag 12/19/2002 08:12 AM
          ' Added If Not RS Is Nothing
          If Not RSEncSections Is Nothing Then
              Call TransferSectionsPIC2LDDB(RSEncSections, cCol, oEncounter, sLDDBEncID, sPICEncID, oDMBE.IDMappers)
          End If
          
          'Transfer Text Notes
          If Not RSTextNotes Is Nothing Then
            Call TransferTextNotes(RSTextNotes, oEncounter, sPICEncID, oDMBE.IDMappers)
          End If
          
          'transfer SO Notes
          If Not RSListToolRecord Is Nothing Then
              Call TransferSONotes(RSListToolRecord, RSListToolRecords, oEncounter, sPICEncID, oDMBE.IDMappers)
          End If
          
          For Each RSObject In cCol
              With RSObject
          
                  If .TableName = "ENC_DISPOSITION" Then
                       Call TransferData2LDDB(RSObject, oEncounter, sPICEncID)
                  ElseIf .TableName = "WORK_STATUS" Then
                      Call TransferWorkStatus(RSObject, oEncounter, sPICEncID, oDMBE.IDMappers)
                  ElseIf .TableName = "ENC_DIAGNOSIS" Then
                      Call TransferData2LDDB(RSObject, oEncounter, sPICEncID)
                  ElseIf .TableName = "ENC_SIGBLOCK" Then
                      Call TransferData2LDDB(RSObject, oEncounter, sPICEncID)
                  'CLINNOTE not part of Encounters
                  'Also, TransferData2LDDB is looking for Column "ENCOUNTERNUMBER"
                  ' CLINNOTE Table has Column Name "ENCOUNTER"
  '                ElseIf .TableName = "CLINNOTE" Then
  '                    Call TransferData2LDDB(RSObject, oEncounter, sPICEncID, True)
                  ElseIf .TableName = "SONOTES" Then
                      Call TransferData2LDDB(RSObject, oEncounter, sPICEncID, True)
                  ElseIf UCase(.TableName) = "NCIDS" Then
                      'Transfer the NICDS table
                      Call TransferNCIDS2LDDB(RSObject)
                      
                  ElseIf .TableName = "ENCTEXTNOTES" Then
                        If RSObject.DBRecordSet.RecordCount <> 0 Then
                            RSObject.DBRecordSet.MoveFirst
                            Do While Not RSObject.DBRecordSet.EOF
                                sTempNCID = RSObject.DBRecordSet("CREATECLINICIANNCID") & vbNullString
                                If sTempNCID <> vbNullString Then
                                    If oDMBE.IDMappers.Exists(NCID_PROVIDER_SERVICES, "NCID", sTempNCID) Then
                                        sTempNCID = oDMBE.IDMappers.Find(NCID_PROVIDER_SERVICES, "NCID", sTempNCID)
                                    End If
                                End If
                                RSObject.DBRecordSet("CREATECLINICNCID") = sTempNCID
                                
                                sTempNCID = RSObject.DBRecordSet("MODIFYCLINICNCID") & vbNullString
                                If sTempNCID <> vbNullString Then
                                    If oDMBE.IDMappers.Exists(NCID_PROVIDER_SERVICES, "NCID", sTempNCID) Then
                                        sTempNCID = oDMBE.IDMappers.Find(NCID_PROVIDER_SERVICES, "NCID", sTempNCID)
                                    End If
                                End If
                                RSObject.DBRecordSet("MODIFYCLINICNCID") = sTempNCID
                                RSObject.DBRecordSet.Update
                            Loop
                            RSObject.DBRecordSet.MoveFirst
                            Call TransferData2LDDB(RSObject, oEncounter, sPICEncID, True)
                        End If
                  End If
              End With
          Next
          
NextRecord:
      
          RSEncounters.DBRecordSet.MoveNext
      Loop
    
      'Let problems do transfer of Dx, Px and Problems
      Set oProblems = mobjComm.InitializeOLEServer(NCID_PROBLEMS_APPLICATION)
      Call oProblems.Transfer(oDMBE)
      
      Set oVitals = mobjComm.InitializeOLEServer(NCID_VITAL_SIGNS_APPLICATION)
      Call oVitals.Transfer(oDMBE)
      
      'Transfer Meds SCR-28326
      Set oMeds = mobjComm.InitializeOLEServer(NCID_MEDICATIONS_APPLICATION)
      Call oMeds.Transfer(oDMBE)
    
      'Transfer Inpatient Data
      Set oPAD = mobjComm.InitializeOLEServer(NCID_INPATIENT_APPLICATION)
      Call oPAD.Transfer(oDMBE)
        
    End If
    '>> End: SCR #36826;

    
    Set Rs = Nothing
    Set RSObject = Nothing
    Set RSEncounters = Nothing
    Set RSEncSections = Nothing
    Set RSTextNotes = Nothing
    Set RSListToolRecord = Nothing
    Set RSListToolRecords = Nothing
    Set RSEncDisposition = Nothing
    Set oEncounter = Nothing
    Set oSQL = Nothing
    Set Field = Nothing
    Set oProblems = Nothing
    Set oVitals = Nothing
    Set oMeds = Nothing
    
    Exit Function
    
ErrHandler:

    Call mobjShared.ShowVBError(Err.Number, Err.Description, "EncounterOps.TransferPIC2LDDB", "EncounterOps", vbCritical)
    Exit Function
    Resume
End Function


Private Sub TransferSectionsPIC2LDDB(ByRef RSEncSections As IDMBE.I_RSObjects, ByRef cCol As Collection, ByRef oEncounter As Encounter, ByVal sLDDBEncID As String, ByVal sPICEncID As String, oIDMappers As I_IDMappers)

Dim sMsg As String

    On Error GoTo ErrHandler

    ' Begin:SCR #30970;   Developer: Mag 12/19/2002 08:11 AM
    If RSEncSections.DBRecordSet.BOF And RSEncSections.DBRecordSet.EOF Then
        Exit Sub
    End If
    ' End: SCR #30970;
    
    RSEncSections.DBRecordSet.MoveFirst
    
    Do Until RSEncSections.DBRecordSet.EOF
            'Look for section with this encounter number
            If RSEncSections.DBRecordSet("ENCOUNTERNUMBER") = CLng(sPICEncID) Then
                TransferSectionRTF RSEncSections, RSEncSections.DBRecordSet("ENC_SECTIONSINDEX"), oEncounter, oIDMappers
            End If
        RSEncSections.DBRecordSet.MoveNext
    Loop
    
    Exit Sub
    
ErrHandler:

    Call mobjShared.ShowVBError(Err.Number, sMsg & " - " & Err.Description, "EncounterOps.TransferSectionsPIC2LDDB", "EncounterOps", vbCritical)

End Sub


Private Sub TransferSectionRTF(RSSection As IDMBE.I_RSObjects, lSection As Long, oEncounter As Encounter, oIDMappers As I_IDMappers)

Dim oSQL       As ICHCSII_SQL.ISqlOpsEx
Dim sOwnerNCID As String
    On Error GoTo ErrHandler
    
    Set oSQL = gobjCHCSIIConn.CHCSII_SQLOPS_EX(Auto)
    
    oSQL.Execute ("Select * from ENC_SECTIONS where ENCOUNTERNUMBER = " & oEncounter.EncounterID & _
                    " AND ENC_SECTIONSINDEX = " & lSection)
    
    If oSQL.EOF Then
        oSQL.AddNew
        oSQL!DataId = oSQL.GetNextID(exMAINSEQ)   'GEMS_Seq_NextVal("MainSEQ")
    End If
    
    With RSSection
        oSQL!Doc = .DBRecordSet("DOC")
        oSQL!WSRFV = .DBRecordSet("WSRFV")
        oSQL!DTS = .DBRecordSet("DTS")
        sOwnerNCID = .DBRecordSet("OWNERNCID") & vbNullString
        If sOwnerNCID <> vbNullString Then
          If oIDMappers.Exists(NCID_PROVIDER_SERVICES, "NCID", sOwnerNCID) Then
            sOwnerNCID = oIDMappers.Find(NCID_PROVIDER_SERVICES, "NCID", sOwnerNCID)
          End If
        End If
        oSQL!OwnerNCID = sOwnerNCID
        oSQL!OwnerName = .DBRecordSet("OWNERNAME")
        oSQL!OriginalSize = .DBRecordSet("ORIGINALSIZE") 'SCR-22830
        oSQL!EncounterNumber = oEncounter.EncounterID
        oSQL!FacilityNCID = oEncounter.FacilityNCID
        oSQL!enc_sectionsindex = lSection
    End With
    
    oSQL.Update
     Exit Sub
    
ErrHandler:

    Call mobjShared.ShowVBError(Err.Number, Err.Description, "EncounterOps.TransferSectionRTF", "EncounterOps", vbCritical)

End Sub


Private Sub TransferTextNotes(RSTextNotes As IDMBE.I_RSObjects, oEncounter As Encounter, ByVal sPICEncID As String, oIDMappers As I_IDMappers)
Dim oNote As TextNote
Dim RSObject    As IDMBE.I_RSObjects
Dim objTextNotes As TextNotes
Dim objEncRTF As EncRTF
Dim oSQL As ICHCSII_SQL.ISqlOpsEx
Set oSQL = gobjCHCSIIConn.CHCSII_SQLOPS_EX(Auto)
Dim sUserNCID As String

    On Error GoTo ErrHandler
    
    Set objTextNotes = oEncounter.SectionParent(eNotes).mobjSection
    objTextNotes.OpenEncounter oEncounter.SectionParent(eNotes), oEncounter

    If RSTextNotes.DBRecordSet.RecordCount > 0 Then
      RSTextNotes.DBRecordSet.MoveFirst
    End If
        
    Do Until RSTextNotes.DBRecordSet.EOF
        'Find Notes for this Encounter
        With RSTextNotes
            If .DBRecordSet("ENCOUNTERNUMBER") = CLng(sPICEncID) And _
               .DBRecordSet("RTFTYPE") = "NOTE" Then
                Set oNote = New TextNote
                If IsDate(.DBRecordSet("DTS")) Then
                    oNote.mdDTS = .DBRecordSet("DTS")
                Else
                    oNote.mdDTS = Now
                End If
                
                ' Begin:SCR #35666;   Developer: Mag 05/07/2003 04:50 PM
                oNote.msOwnerName = .DBRecordSet("USERNAME") & vbNullString
                sUserNCID = .DBRecordSet("USERNCID") & vbNullString
                If oIDMappers.Exists(NCID_PROVIDER_SERVICES, "NCID", sUserNCID) Then
                  sUserNCID = oIDMappers.Find(NCID_PROVIDER_SERVICES, "NCID", sUserNCID)
                End If
                oNote.msOwnerNCID = sUserNCID
                ' End: SCR #35666;
                
                If .DBRecordSet("Complete") = "Y" Then
                    oNote.mbIsComplete = True
                End If
                
                ' Begin:SCR #35666;   Developer: Mag 05/07/2003 04:51 PM
                oNote.msCategory = .DBRecordSet("CATEGORY") & vbNullString
                oNote.msTitle = .DBRecordSet("TITLE") & vbNullString
                oNote.msRTF = .DBRecordSet("DOC") & vbNullString
                ' End: SCR #35666;
                
                'Lock Section before saving SCR#44787
                oEncounter.LockingObject.LockSection eNotes, 0, SectionLock
                objTextNotes.Save oNote
                oEncounter.LockingObject.UnlockSection eNotes, 0
            ElseIf .DBRecordSet("ENCOUNTERNUMBER") = CLng(sPICEncID) And _
                   .DBRecordSet("RTFTYPE") = "RTF" Then
               
               'SCR 164074
               oSQL.Execute "Select * from enc_rtfs where encounternumber = " & oEncounter.EncounterID & " and " _
                    & " FacilityNCID = " & oEncounter.FacilityNCID & " and DTS = " & gobjCHCSIIConn.SQLDate(.DBRecordSet("DTS")) & " And " _
                    & " RTFType = 'RTF'"
               'oSQL.Execute "enc_rtfs"
               If oSQL.EOF Then
                   oSQL.AddNew
               End If
               oSQL("FacilityNCID") = oEncounter.FacilityNCID
               oSQL("EncounterNumber") = oEncounter.EncounterID
               If IsDate(.DBRecordSet("DTS")) Then
                   oSQL("DTS") = .DBRecordSet("DTS")
                Else
                    oSQL("DTS") = Now
                End If
                
                ' SCR #47334;   Developer: Mag 02/12/2004 02:24 PM
                sUserNCID = .DBRecordSet("USERNCID") & vbNullString
                If oIDMappers.Exists(NCID_PROVIDER_SERVICES, "NCID", sUserNCID) Then
                  sUserNCID = oIDMappers.Find(NCID_PROVIDER_SERVICES, "NCID", sUserNCID)
                End If
                oSQL("USERNCID") = sUserNCID
               
               oSQL("RTFTYPE") = "RTF"
               oSQL("ENC_RTFSINDEX") = oSQL.GetNextID(exMAINSEQ)
               Call oSQL.AppendAsChunk("DOC", .DBRecordSet("DOC"))
               oSQL("ORIGINALSIZE") = .DBRecordSet("ORIGINALSIZE") 'SCR-22830
               oSQL.Update
            End If
            
            .DBRecordSet.MoveNext
        End With
    Loop
    
    Set objTextNotes = Nothing
                        
    Exit Sub
    
ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description, "EncounterOps.TransferTextNotes", "EncounterOps", vbCritical)
    
    Exit Sub
    
    Resume
End Sub

Private Sub TransferSONotes(RSListToolRecord As IDMBE.I_RSObjects, RSListToolRecords As IDMBE.I_RSObjects, oEncounter As Encounter, ByVal sPICEncID As String, oIDMappers As I_IDMappers)

Dim objSO As SO
Dim oSONOte As SONote
Dim oLTR As ListToolRecord
Dim sOwnerNCID As String
    On Error GoTo ErrHandler
    
    Set objSO = oEncounter.SectionParent(eSO).mobjSection
    objSO.OpenEncounter oEncounter.SectionParent(eSO), oEncounter

    If RSListToolRecord.DBRecordSet.RecordCount > 0 Then
      RSListToolRecord.DBRecordSet.MoveFirst
    End If
        
    Do Until RSListToolRecord.DBRecordSet.EOF
        'Look for SO list tool record with PIC enc id
        If RSListToolRecord.DBRecordSet("ENCOUNTERNUMBER") = CLng(sPICEncID) And _
            RSListToolRecord.DBRecordSet("SOURCENCID") = SectionEnumToNCID(eSO) Then
        
            Set oSONOte = New SONote
            
            Select Case meSource
                Case eDM_EIC, eDM_PIC, eDM_AGGPIC
                    sOwnerNCID = RSListToolRecord.DBRecordSet("CLINICIANNCID") & vbNullString
                    If sOwnerNCID <> vbNullString Then
                      If oIDMappers.Exists(NCID_PROVIDER_SERVICES, "NCID", sOwnerNCID) Then
                        sOwnerNCID = oIDMappers.Find(NCID_PROVIDER_SERVICES, "NCID", sOwnerNCID)
                      End If
                    End If
                    oSONOte.msOwnerNCID = sOwnerNCID
'                    'PROVIDER_NCID = SSN look up NCID from Providers table
'                    oSQL2.Execute ("SELECT NCID FROM PROVIDER WHERE SSN = '" & RSListToolRecord.DBRecordSet("CLINICIANNCID")) & "'"
'                    ' Begin:SCR #35666;   Developer: Mag 05/07/2003 04:51 PM
'                    oSONOte.msOwnerNCID = oSQL2("NCID") & vbNullString
'                    ' End: SCR #35666;
                Case Else
                   ' Begin:SCR #35666;   Developer: Mag 05/07/2003 04:51 PM
                   oSONOte.msOwnerNCID = RSListToolRecord.DBRecordSet("CREATEDBY") & vbNullString
                   ' End: SCR #35666;
            End Select
            
            
            oSONOte.msRTF = RSListToolRecord.DBRecordSet("RTF")
            
            If RSListToolRecords.DBRecordSet.RecordCount > 0 Then
              RSListToolRecords.DBRecordSet.MoveFirst
            End If
            
            'Get the list tool records
            Do Until RSListToolRecords.DBRecordSet.EOF
                With RSListToolRecords
                    'Find List tools records for the matching DATAID of the list tool record
                    If .DBRecordSet("DATAID") = RSListToolRecord.DBRecordSet("DATAID") Then
                    
                        Set oLTR = New ListToolRecord
                        
                        oLTR.SnoID = .DBRecordSet("SnoID")
                        If .DBRecordSet("DURATION") <> Null Then oLTR.Duration = .DBRecordSet("DURATION")
                        If .DBRecordSet("MODIFIER") <> Null Then oLTR.Modifier = .DBRecordSet("MODIFIER")
                        If .DBRecordSet("NOTE") <> Null Then oLTR.Note = .DBRecordSet("NOTE")
                        If .DBRecordSet("ONSET") <> Null Then oLTR.Onset = .DBRecordSet("ONSET")
                        If .DBRecordSet("PREFIX") <> Null Then oLTR.prefix = .DBRecordSet("PREFIX")
                        If .DBRecordSet("RESULT") <> Null Then oLTR.result = .DBRecordSet("RESULT")
                        If .DBRecordSet("STATUS") <> Null Then oLTR.Status = .DBRecordSet("STATUS")
                        If .DBRecordSet("UNIT") <> Null Then oLTR.Unit = .DBRecordSet("UNIT")
                        If .DBRecordSet("VALUE") <> Null Then oLTR.value = .DBRecordSet("VALUE")
                        oLTR.RangeNormalHigh = .DBRecordSet("RANGENORMALHIGH")
                        oLTR.RangeNormalLow = .DBRecordSet("RANGENORMALLOW")
                        oLTR.RangeScale = .DBRecordSet("RANGESCALE")
                        oLTR.Referid = .DBRecordSet("REFERID")
                    
                        If oSONOte.mcolListRecords Is Nothing Then
                            Set oSONOte.mcolListRecords = New Collection
                        End If
                        oSONOte.mcolListRecords.Add oLTR
                    End If
                End With
            
                RSListToolRecords.DBRecordSet.MoveNext
            Loop
            
            objSO.SaveNote oSONOte, True
            
        End If
        
        RSListToolRecord.DBRecordSet.MoveNext
    
    Loop
    
    Exit Sub
    
ErrHandler:

    Call mobjShared.ShowVBError(Err.Number, Err.Description, "EncounterOps.TransferSONotes", "EncounterOps", vbCritical)
Exit Sub
    Resume
End Sub



Private Sub TransferWorkStatus(ByRef RSWorkStatus As IDMBE.I_RSObjects, oEncounter As Encounter, ByVal sPICEncID As String, oIDMappers As I_IDMappers)

Dim oSQL       As ICHCSII_SQL.ISqlOpsEx
Dim Field As ADODB.Field
Dim sUserNCID As String
    On Error GoTo ErrHandler
    
    Set oSQL = gobjCHCSIIConn.CHCSII_SQLOPS_EX(Auto)
    
    If RSWorkStatus.DBRecordSet.RecordCount > 0 Then
      RSWorkStatus.DBRecordSet.MoveFirst
    End If
    
    Do Until RSWorkStatus.DBRecordSet.EOF
        With RSWorkStatus
            If .DBRecordSet("ENCOUNTER_NUMBER") = CLng(sPICEncID) Then
                oSQL.Execute ("SELECT * FROM WORK_STATUS WHERE ENCOUNTER_NUMBER = " & oEncounter.EncounterID _
                                & " AND FACILITY_ID = " & oEncounter.FacilityNCID)
                
                If oSQL.EOF Then
                    oSQL.AddNew
                End If
                
                For Each Field In .DBRecordSet.Fields
                    If Field.Name <> "ENCOUNTER_NUMBER" Then
                        oSQL.value(Field.Name) = .DBRecordSet(Field.Name)
                    End If
                Next
                    
                Select Case meSource
                    Case eDM_EIC, eDM_PIC, eDM_AGGPIC
                        sUserNCID = .DBRecordSet("PROVIDER_NCID") & vbNullString
                        If sUserNCID <> vbNullString Then
                          If oIDMappers.Exists(NCID_PROVIDER_SERVICES, "NCID", sUserNCID) Then
                            sUserNCID = oIDMappers.Find(NCID_PROVIDER_SERVICES, "NCID", sUserNCID)
                          End If
                        End If
                        oSQL!Provider_NCID = sUserNCID
                        
'                        'PROVIDER_NCID = SSN look up NCID from Providers table
'                        Dim oSQL2       As ICHCSII_SQL.ISqlOpsEx
'                        oSQL2.Execute ("SELECT NCID FROM PROVIDER WHERE SSN = '" & .DBRecordSet("PROVIDER_NCID")) & "'"
'                        oSQL!Provider_NCID = "" & oSQL2("PROVIDER_NCID")
                End Select
                
                'SCR 163399
                oSQL.value("Encounter_Number") = oEncounter.EncounterID
                oSQL.Update
    
            End If
            .DBRecordSet.MoveNext
        End With
    Loop
    
    Exit Sub
    
ErrHandler:

    Call mobjShared.ShowVBError(Err.Number, Err.Description, "EncounterOps.TransferWorkStatus", "EncounterOps", vbCritical)
    Exit Sub
    Resume
End Sub


Private Sub TransferData2LDDB(ByRef RSObject As IDMBE.I_RSObjects, oEncounter As Encounter, ByVal sPICEncID As String, Optional ByVal bDataID As Boolean = False)

Dim oSQL       As ICHCSII_SQL.ISqlOpsEx
Dim Field As ADODB.Field

    On Error GoTo ErrHandler
    Set oSQL = gobjCHCSIIConn.CHCSII_SQLOPS_EX(Auto)
    'SCR 129936
    If RSObject.DBRecordSet.RecordCount > 0 Then
      RSObject.DBRecordSet.MoveFirst
    End If
    
    Do Until RSObject.DBRecordSet.EOF
        With RSObject
            If .DBRecordSet("ENCOUNTERNUMBER") = CLng(sPICEncID) Then
                oSQL.Execute ("SELECT * FROM " & RSObject.TableName & " WHERE ENCOUNTERNUMBER = " & oEncounter.EncounterID _
                                & " AND FACILITYNCID = " & oEncounter.FacilityNCID)
                
                If oSQL.EOF Then
                    oSQL.AddNew
                    If bDataID Then
                        oSQL!DataId = oSQL.GetNextID(exMAINSEQ)  'GEMS_Seq_NextVal("MainSEQ")
                    End If
                End If
                
                For Each Field In .DBRecordSet.Fields
                    If Field.Name <> "ENCOUNTERNUMBER" Then
                        oSQL.value(Field.Name) = .DBRecordSet(Field.Name)
                    End If
                Next

                oSQL!EncounterNumber = oEncounter.EncounterID
                oSQL.Update
    
            End If
            
            .DBRecordSet.MoveNext
            
        End With
                
    Loop
    Exit Sub
    
ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description, "EncounterOps.TransferEncDisposition", "EncounterOps", vbCritical)

End Sub


Private Sub TransferNCIDS2LDDB(ByRef RSObject As IDMBE.I_RSObjects)

'Merges the data in the NCIDS table.

Dim oSQL       As ICHCSII_SQL.ISqlOpsEx
Dim Field As ADODB.Field

    On Error GoTo ErrHandler
    Set oSQL = gobjCHCSIIConn.CHCSII_SQLOPS_EX(Auto)
    
    Do Until RSObject.DBRecordSet.EOF
        With RSObject
            oSQL.Execute ("SELECT * FROM " & RSObject.TableName & " WHERE DOMAINNCID = 119 and NCID = " & RSObject.DBRecordSet("NCID"))
            
            If oSQL.EOF Then
                oSQL.AddNew
            End If
            
            For Each Field In .DBRecordSet.Fields
                oSQL.value(Field.Name) = .DBRecordSet(Field.Name)
            Next

            oSQL.Update

            .DBRecordSet.MoveNext
            
        End With
                
    Loop
    Exit Sub
    
ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description, "EncounterOps.TransferNCIDS2LDDB", "EncounterOps", vbCritical)

End Sub






' Reset cached Encounter, called by EncounterParent.cls CloseEncounterPrivate SCR-31509
Public Sub ResetCurrLoadedEncounter()

    Set mobjCurrEncounter = Nothing

End Sub


'Added for PGUI mode SCR-40019
Public Function GetEncounterStatus(ByVal ApptID As String, ByRef EncounterNumber As String) As EncStatus

    Dim sSql As String
    Dim oSQL As ICHCSII_SQL.ISqlOps
    Set oSQL = gobjCHCSIIConn.CHCSII_SQLOPS(Auto)
        
    On Error GoTo ErrHandler

        sSql = "select ApptID, Status from Encounters where EncounterNumber = '" & EncounterNumber & "'" _
            & " and FacilityNCID = " & mobjLogon.FacilityNCID & " and Unit_Number = " & mobjPatient.UnitNumber

        oSQL.Execute sSql

        If Not oSQL.EOF Then
            ApptID = oSQL("ApptID")
            GetEncounterStatus = StatusNCIDToEnum(oSQL("Status"))
        End If

    Set oSQL = Nothing
    
    Exit Function
    
ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description, "EncounterOps.GetEncounterStatus", "EncounterOps", vbCritical)
End Function

Public Function CreateCHCSEncounter(ByVal ApptID As String, ByVal ClinicianNCID As String) As Encounter
    Dim EncType As EncType
    On Error GoTo ErrHandler
    
    'Set correct EncType
    Select Case mobjCHCSAppt.ApptType
        Case "T-CON*", "T-CON", "TCON", "TEL-CON", "OCC-SVR"
            EncType = TeleConsult
        Case Else
            If mobjCHCSAppt.ApptClassification = 1 Then
                EncType = Inpatient
            Else
                EncType = Outpatient
            End If
    End Select
                
    Set CreateCHCSEncounter = CreateEncounter(ApptID, mobjCHCSAppt.ApptType, EncType, mobjCHCSAppt.FacilityNCID, mobjLogon.NursingDivisionNCID, ClinicianNCID, _
            , , mobjCHCSAppt.ApptDateTime, , , , , , , , , , , CStr(ApptID))
            
    Exit Function
    
ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description, "EncounterOps.CreateCHCSEncounter", "EncounterOps", vbCritical)

End Function

'SCR#45056,45066,45074
Private Function GetIfEncounterExists(ByVal sApptId As String, ByVal lFacilityNCID As Long, ByVal lUnitNumber As Long, ByVal PrimaryProviderNCID As String) As CHCSIIEncounterOps.Encounter

    Dim sSql As String
    Dim oSQL As ICHCSII_SQL.ISqlOps
    Dim oShared As CWShared
    Dim sEncNum As String
    Dim objEncounter As CHCSIIEncounterOps.Encounter
    Dim objConcept As GEMS_ConceptCtrl
    Dim sProviderName As String
    
    On Error GoTo ErrHandler
    
    Set objConcept = New GEMS_ConceptCtrl
    Set oSQL = gobjCHCSIIConn.CHCSII_SQLOPS(Auto)
    Set oShared = New CWShared
    
    'date:01/09/2004 SCR #:47870 developer: jrm Description: Added for ITT mode
    'Auto is needed to determine ITT, or Theatre
    
    If oShared.IsAppMode(modeCDR) = False Or oShared.IsAppMode(modeITT) = True Then
        sSql = "select * from Encounters where APPTID = '" & sApptId & "'" _
            & " and FacilityNCID = " & lFacilityNCID & " and Unit_Number = " & lUnitNumber
        
        oSQL.Execute sSql
        If Not oSQL.EOF Then
            sEncNum = oSQL("EncounterNumber")
        End If
        
        If Len(sEncNum) Then
            'If Enc Num exists then get encounter
            Set objEncounter = GetEncounter(lUnitNumber, lFacilityNCID, sEncNum)
            If Not objEncounter Is Nothing Then
                'Make sure PP is correct since transfer may have been done in CHCS SCR#51564
                If objEncounter.PrimaryProvider.NCID <> PrimaryProviderNCID Then
                    objConcept.UniqueId = PrimaryProviderNCID
                    sProviderName = objConcept.PrefRep("2000").Representation
                    If objEncounter.TransferToProvider(PrimaryProviderNCID, sProviderName) Then
                        objEncounter.LockingObject.LockSection eHeader, 0, SectionLock
                        objEncounter.Save True
                        objEncounter.LockingObject.UnlockSection eHeader, 0
                    End If
                End If
                
                Set GetIfEncounterExists = objEncounter
            End If
        End If
    Else
        MsgBxARMd "Function is not supported in CDR mode"
    End If
    
    Set objConcept = Nothing
    Set objEncounter = Nothing
    Set oShared = Nothing
    Set oSQL = Nothing
    
    Exit Function

ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description, "EncounterOps.GetIfEncounterExists", "EncounterOps", vbCritical)
 
End Function
'Called by Encounter.CreateEncounter
Private Function GetDefaultTemplateData() As String

    Dim objTemplateMgmt As Object
    
    On Error GoTo ErrHandler
    
    Set objTemplateMgmt = mobjComm.InitializeOLEServer(NCID_TEMPLATE_MANAGEMENT)
    If Not objTemplateMgmt Is Nothing Then
       GetDefaultTemplateData = objTemplateMgmt.DefaultEncounterTemplateData(mobjLogon.UserNCID)   'SCR 26839 RBell 10/10/2002
    End If
    Set objTemplateMgmt = Nothing
    
    Exit Function
    
ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description, "EncounterOps.GetDefaultTemplateData", "EncounterOps", vbCritical)

End Function
