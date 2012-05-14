VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Encounter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "The Actual Encounter."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"EncNote"
Option Explicit

Private Const MODULE_NAME = "Encounter"
Private Const EMPTY_LOCATION_IEN = 0

Public Appt_IEN As String
Public Patient_IEN As String
'---------------------------
'The individual encounter.  Identified by Patient UnitNumber and Encounter Number.

Private Const NCID_INPATIENT_DOCUMENTS_APPLICATION  As String = "1999938"
Private Const NCID_Admitted As String = "204849" '"208849"  'SCR#43518 'Came from MMMHISDisposition.modDisposition
Private Const NCID_Admission_HP_Note As String = "15639660" ' "15119311" 'SCR 134416, JQ

'- presisted properties
Private mnEncounterID        As Long        '- Encounter ID
Private msEnterpriseNCID    As String      '- Enterprise NCID
Private msFacilityNCID      As String      '- Facility NCID
Private msFacility          As String      '- Facility long name
Private msClinicNCID        As String      '- Clinic NCID
Private msClinic            As String      '- clinic long name
Private meStatus            As EncStatus   '- Current Status of Enc
Private msCategory          As String      '- Really the appt type as a string
Private meClass             As EncClass    '- Encounter Class (Followup / New / NoCount)
Private meType              As EncType     '- Enc Category (misnamed) one of {Inpatient, outpatient, ambulatory}
Private mdEndDTS            As Date        '- End Date/Time
Private msApptID            As String      '- the related appointment id or Telecon ID
Private mnPatientID         As Long        '- unit number from the patient obj
Private msPatientName       As String
Private msPrimaryDiagnosisNCID As String
Private mdStartDTS          As Date        '- Start Date/Time
Private mcolEncRTFs         As Collection  '- collection of complete, signed RTF representations
Private mcolSecParents      As Collection  '- collection of SectionParent objects
Private mcolProviders       As EncProviders   '- collection providers for this encounter
Private msSADRStatus        As String
'Private mbSensitivity       As Boolean
Private mlSensitivity       As Long     '- are we sensitive? SCR-35365
Private msTemplateData      As String      '- place to store template information from template mgr
Private msTemplateID      As String      '- place to store template information from template mgr
Private msTimeZone          As String      '- 3-letter timezone designation
Private mobjCosigner        As EncProvider
Private mobjWhoSigned       As EncProvider
Private mdWhenSigned        As Date
Private mobjWhoCosigned     As EncProvider
Private mdWhenCosigned      As Date
Private mbWasUpdated        As Boolean
Private msAllergyVerifiedByNCID  As String
Private mdAllergyVerifiedByDate  As Date
Private mbBypassSODialog    As Boolean
Private mbArrivedByAmbulance As Boolean
Private mblnUpdateLock      As Boolean
Private mblnUpdateLockPending As Boolean
Private mdApptDTS           As Date 'SCR-17624
Private mnPrimaryDxSnoID As String 'SCR-19371
Private msPrimaryDxMedcinPrefix As String
Private msRTF As String 'SCR-24498
Private mbResultedConsultOrder As Boolean
Private msApptNotesRTF      As String
Private msApptRFV           As String
Private msApptNote          As String
Private mnApptNotesDataID   As Long
Private msApptIEN           As String
Private msApptType          As String
Private meApptClass         As ApptClassificationEnum
Private msAPStatusComment   As String  'SCR-17942
Private mobjWAM             As WAMData.WAM
Private mbFiledEncounter    As Boolean
Private msWorkLoadOptions   As String
Private msWorkLoad          As String
Private msMEPRSCodes          As String
Private msPatientStatus       As String   'Same as Disp Appt Class
Private msInpatientService  As String
Private mobjAccidentData   As EncAccidentData        'HIPAA Accident Data Object
Private mobjPregnancyData  As EncPregnancy           'HIPAA Pregnancy Data Object
Private msCreateClinicianNCID As String 'SCR#42996
Private msProvHIPAACode     As String

'SF Added Inpatient Note Type
Private msInptNoteTypeNCID As String
Private mlInptHospDay As Long

'Extra properties for E & M Code Review
Private msEandMCodeReview   As String

'- Non-persisted propterties
Private mbOpen              As Boolean
Private mobjParent          As Object      ' can only be EncounterParent for current, previous or cosign
Private mbFullyLoaded       As Boolean     '- was this enc filled with a full read of db or just a create/getinfo load?
Private mbFlagNew           As Boolean     '- Is this a new enc?
Private mbDirty             As Boolean     '- has this Enc Changed?
Private mnDataID            As Long        '- from Das control update of nonCDR data
Private mobjLocking         As Locking     '- object for locking sections of this encounter
Private mobjRTF             As EncRTF      '- the current/most recent rtf representation of the encounter
Private mobjUser            As ICHCSII_User
Private mobjComm            As MMMHISComm.Comm
Private mobjShared          As MMMHISShared.CWShared
Private mobjLogon           As MMMHISLogon.Logon
Private mobjPatient         As MMMHISPatient.Patient
Private mnUpdateCount       As Long        '- count in objDas.Events.ActionsInfo.Count
Private msUpdateDate        As Date        '- date of last update
Private msUpdateClinicianNCID As String      '- NCID of last updated by
Private mbLoading             As Boolean     '-flag to stop side effects during initial data load
Private mobjCurrSection     As SectionParent ' currently open section
Private mobjProperties      As EncProperties
Private mbSignEncbyEncParent As Boolean  'SCR 22247 flag ARM to time sign from Enc only (not children)  MSolano 4/24/02
Private mbIsBMISTEnc As Boolean ' SCR 128313

' 7/08/2002
'WR moved to Encops_itf.tlb SCR 25225
'Public Enum EncounterSectionEnum
'    eAutoCite = 1
'    eRFV = 2
'    eVitals = 3
'    eSO = 4
'    eLabs = 5
'    eRads = 6
'    eQuestionnaire = 7
'    eap = 8
'    eDisp = 9
'    eNotes = 10
'    eDental = 11
'    eEducation = 12
'    eSOOpen = 96
'    eHistory = 97
'    eHeader = 98
'    eEntireEncounter = 99
'End Enum

Public Enum EncType     '-- this is really enc category, for the first 3, and category is appt 'type'
  Inpatient = 0
  Outpatient = 1
  Ambulatory = 2
  TeleConsult = 3
  Dental = 4
  InpatientNote = 5
End Enum

Public Enum EncClass
  ec_Followup = 1
  ec_New = 2
  ec_NoCount = 3
End Enum

Public Enum EncSensitivity
  None = 0
  VIP = 1
  Adolescent = 2
  Medicine = 3
  Psychiatry = 4
  HIV = 5
  FAP = 6
End Enum

Public Enum EncSensitiveRecord 'SCR-35365
  esrNone = 0
  esrStandard = 1
  esrBehavioralHealth = 2
  esrFamilyAdvocacy = 4
End Enum

' 7/08/2002
'WR moved to Encops_itf.tlb SCR 25225
'Public Enum EncStatus
'  Incomplete = 1
'  NeedsSignature = 2
'  NeedsCoSignature = 3
'  Pending = 4
'  CheckedIn = 5
'  NoSave = 6
'  Waiting = 7
'  InProgress = 8
'  Updated = 10
'  complete = 11
'  Updating = 12
'  PatientLeftWithoutBeingSeen = 13
'  CancelByFacility = 14
'  CancelByPatient = 15
'  NoShow = 16
'  CancelByProvider = 17
'End Enum

Public Enum ChangeHistoryEnum
    Deletion = 1
    Overwritten = 2
    LockedOut = 3
    EncAmended = 4
End Enum

Public Type Encounter_Key_Information
        ID                                   As String
        Facility                             As String
        Unit_NUmber                          As Long
        Category                             As String  'SCR#44005
End Type

Public Enum Provider_Role_Type_NCIDs
   rtnAssistingProvider = 209680
   rtnAttendingProvider = 209679
   rtnSupervisingProvider = 209681
   rtnNurse = 67941
   rtnParaProfessional = 209682
   rtnOperatingProvider1 = 15009420    'SCR 50671 - Updated NCID value
End Enum

Public Enum Provider_Priority
   pcPRIMARY = 1
   pcFIRST_ADDITIONAL
   pcSECOND_ADDITIONAL
End Enum

Private Const Exe = "CHCSIIEncounterOps"
Private Const Module = "Encounter"
Private Const ExeModule = Exe & "." & Module
#If debugon Then
    Private DR As DebugRec
    Private DL As cDebugLog
#End If
'Public Event EncStatusChanged(meStatus As ENCOUNTER_INTERFACES.EncStatus)
Public Event EncStatusChanged(meStatus As EncStatus)

Implements IEncDoc
Public Function GetCommObj() As MMMHISComm.Comm  'expose the Comm object for Form to get to the Green Switch
    Set GetCommObj = mobjComm
End Function

Public Function CopyForwardSONote(pEncounter As Encounter)
    
    Dim objSO As SO
    Dim oSONOte As SONote
    Dim oNewSONote As SONote
    Dim colListRecords As Collection
    Dim oLRecord As ListToolRecord
    Dim oFindItem As ListToolRecord
    Dim sNote1 As String
    Dim sNote2 As String
    
    Set colListRecords = New Collection
    
    With Me 'If multiple encounters are concatenated then get last copy forward note to append to
        Set objSO = .SectionParent(eSO).mobjSection
        If Not objSO.CopyForwardNote Is Nothing Then
            If objSO.CopyForwardNote.mcolListRecords.Count > 0 Then
               Set colListRecords = objSO.CopyForwardNote.mcolListRecords
            End If
        End If
    End With
    
    With pEncounter
        .Load .PatientID, .FacilityNCID, .EncounterID
        Set objSO = .SectionParent(eSO).mobjSection
    End With
    
    On Error Resume Next
    For Each oSONOte In objSO.Notes
        For Each oLRecord In oSONOte.mcolListRecords
            Set oFindItem = Nothing
            Set oFindItem = colListRecords(CStr(oLRecord.SnoID))
            If Not oFindItem Is Nothing Then
               colListRecords.Remove CStr(oLRecord.SnoID)
               sNote1 = oFindItem.Note
               sNote2 = oLRecord.Note
               If sNote1 = sNote2 Or InStr(1, sNote2, sNote1, vbTextCompare) > 0 Then
                  'Do nothign eithers notes are equal or the new item note is contained in the old note.
               Else
                  If Len(sNote1) > 0 And Len(sNote2) > 0 Then
                     oLRecord.Note = Trim(sNote1) & ". " & Trim(sNote2)
                  Else
                     oLRecord.Note = Trim(sNote1) & Trim(sNote2)
                  End If
               End If
            End If
            colListRecords.Add oLRecord, CStr(oLRecord.SnoID)
        Next
    Next
    
    Set oNewSONote = New SONote
    With oNewSONote
        .msOwnerNCID = mobjLogon.UserNCID
        .mdDTS = Now
        .meType = ListNoteType
        .msOwnerName = mobjLogon.UserName
        .msTitle = "Copied Forward Note"
        Set .mcolListRecords = colListRecords
    End With
    
    With Me
        Set objSO = .SectionParent(eSO).mobjSection
        objSO.CopyForwardNote = oNewSONote
    End With
    
    Set objSO = Nothing
    Set oSONOte = Nothing
    Set oNewSONote = Nothing
    Set colListRecords = Nothing
    Set oLRecord = Nothing
    Set oFindItem = Nothing

End Function

Public Property Get InptHospDay() As Long
    
    InptHospDay = mlInptHospDay
    
End Property

Public Property Get InptNoteTypeNCID() As String
    
    InptNoteTypeNCID = msInptNoteTypeNCID
    
End Property

Public Property Get InptNoteTypeText() As String

'Convert the msInptNoteTypeNCID to a string

Dim objConcept As GEMSConceptCtrl.GEMS_ConceptCtrl

    Set objConcept = New GEMS_ConceptCtrl
    
    objConcept.UniqueId = msInptNoteTypeNCID
    InptNoteTypeText = objConcept.PrefRep("").Representation
    
End Property


Public Property Get IsBMISTEnc() As Boolean
    'SCR 128313
    IsBMISTEnc = mbIsBMISTEnc
End Property

'<<< SCR 22247 flag ARM to time sign from Enc only (not children)  MSolano 4/24/02
Public Property Let SignByEncParent(bSignByParent As Boolean)
  mbSignEncbyEncParent = bSignByParent
End Property
Public Property Get SignByEncParent() As Boolean
  SignByEncParent = mbSignEncbyEncParent
End Property
'>>> SCR 22247
Public Property Get ResultedConsultOrder() As Boolean
    ResultedConsultOrder = mbResultedConsultOrder
End Property

Public Property Let ResultedConsultOrder(ByVal value As Boolean)
    mbResultedConsultOrder = value
End Property

Public Property Get ArrivedByAmbulance() As Boolean
    ArrivedByAmbulance = mbArrivedByAmbulance
End Property

Public Property Let ArrivedByAmbulance(ByVal value As Boolean)
    mbArrivedByAmbulance = value
End Property

Public Property Get BypassSODialog() As Boolean
    BypassSODialog = mbBypassSODialog
End Property

Public Property Let BypassSODialog(ByVal RHV As Boolean)
    mbBypassSODialog = RHV
End Property

Public Property Get DentalDispStatus() As String
    DentalDispStatus = mcolSecParents(CStr(eDisp)).mobjSection.DentalDispStatus
End Property

Public Property Let DentalDispStatus(ByVal RHV As String)
    mcolSecParents(CStr(eDisp)).mobjSection.DentalDispStatus = RHV
End Property

Public Property Get DentalDispStatusDesc() As String
    DentalDispStatusDesc = mcolSecParents(CStr(eDisp)).mobjSection.DentalDispStatusDesc
End Property

Public Property Let DentalDispStatusDesc(ByVal RHV As String)
    mcolSecParents(CStr(eDisp)).mobjSection.DentalDispStatusDesc = RHV
End Property

Public Property Get DentalDispComments() As String
    DentalDispComments = mcolSecParents(CStr(eDisp)).mobjSection.DentalDispComments
End Property

Public Property Let DentalDispComments(ByVal RHV As String)
    mcolSecParents(CStr(eDisp)).mobjSection.DentalDispComments = RHV
End Property

Public Property Get DentalDispOptions() As String
    DentalDispOptions = mcolSecParents(CStr(eDisp)).mobjSection.DentalDispOptions
End Property

Public Property Let DentalDispOptions(ByVal RHV As String)
    mcolSecParents(CStr(eDisp)).mobjSection.DentalDispOptions = RHV
End Property



Friend Property Get EncounterParent() As Object:      Set EncounterParent = mobjParent:             End Property
Public Property Get LockingObject() As Locking:       Set LockingObject = mobjLocking:              End Property
Friend Property Get Comm() As MMMHISComm.Comm:        Set Comm = mobjComm:                          End Property
Friend Property Get CWShared() As MMMHISShared.CWShared: Set CWShared = mobjShared:                 End Property
Friend Property Get User() As CHCSII_ClientSecurity.ICHCSII_User: Set User = mobjUser:              End Property
Friend Property Get Logon() As MMMHISLogon.Logon:     Set Logon = mobjLogon:                        End Property
Friend Property Get SectionParent(ByVal Section As EncounterSectionEnum) As SectionParent: Set SectionParent = mcolSecParents(CStr(Section)): End Property
Friend Property Get Patient():                        Set Patient = mobjPatient:                    End Property
Friend Property Get Config():                        Set Config = mobjParent.Config:                    End Property

Public Property Get AppointmentID() As String:        AppointmentID = msApptID:                     End Property
Public Property Get Category() As String:             Category = msCategory:                        End Property
Public Property Get Class() As EncClass:              Class = meClass:                              End Property
Public Property Get Clinic() As String:               Clinic = msClinic:                            End Property
Public Property Get ClinicNCID() As String:           ClinicNCID = msClinicNCID:                    End Property
Public Property Get Cosigner() As EncProvider:        Set Cosigner = mobjCosigner:                  End Property
Public Property Get DataId() As Long:                 DataId = mnDataID:                            End Property
Public Property Let DataId(lngDataID As Long)
    mnDataID = lngDataID
End Property

Public Property Get EncounterID() As Long:            EncounterID = mnEncounterID:                  End Property
Public Property Get EncounterType() As EncType:       EncounterType = meType:                       End Property
Public Property Get EncProperties() As EncProperties: Set EncProperties = mobjProperties:           End Property
Public Property Get EncRTFs() As Collection:          Set EncRTFs = mcolEncRTFs:                    End Property
Public Property Get EndDTS() As Date:                 EndDTS = mdEndDTS:                            End Property
Public Property Get EnterpriseNCID() As String:       EnterpriseNCID = msEnterpriseNCID:            End Property
Public Property Get Facility() As String:             Facility = msFacility:                        End Property
Public Property Get FacilityNCID() As String:         FacilityNCID = msFacilityNCID:                End Property
Public Property Get Dirty() As Boolean:               Dirty = mbDirty:                              End Property
Public Property Get FullyLoaded() As Boolean:         FullyLoaded = mbFullyLoaded:                  End Property
Public Property Get FlagNew() As Boolean:             FlagNew = mbFlagNew:                          End Property
Public Property Get IsOpen() As Boolean:              IsOpen = mbOpen:                              End Property
Public Property Get Loading() As Boolean:             Loading = mbLoading:                          End Property
Public Property Get PatientID() As Long:              PatientID = mnPatientID:                      End Property
Public Property Get Providers() As Collection:        Set Providers = mcolProviders.Providers:      End Property
Public Property Get AssociatedProviders() As EncProviders:        Set AssociatedProviders = mcolProviders:      End Property
Public Property Get SADRStatus() As String:           SADRStatus = msSADRStatus:                    End Property
Public Property Get StartDTS() As Date:               StartDTS = mdStartDTS:                        End Property
Public Property Get Status() As EncStatus:            Status = meStatus:                            End Property
Public Property Get StorageKey() As String:           StorageKey = Trim$(msFacilityNCID) & "|" & mnEncounterID: End Property
Public Property Get TemplateData() As String:         TemplateData = msTemplateData:                End Property
Public Property Get TextNotes() As TextNotes:         Set TextNotes = SectionParent(eNotes).mobjSection: End Property
Public Property Get TimeZone() As String:             TimeZone = msTimeZone:                        End Property
Public Property Get UpdateDate() As Date:             UpdateDate = msUpdateDate:                    End Property
Public Property Get UpdateClinicianNCID() As String:  UpdateClinicianNCID = msUpdateClinicianNCID:  End Property
Public Property Get WasUpdated() As Boolean:          WasUpdated = mbWasUpdated:                    End Property
Public Property Get WhenCoSigned() As Date:           WhenCoSigned = mdWhenCosigned:                End Property
Public Property Get WhoCoSigned() As EncProvider:     Set WhoCoSigned = mobjWhoCosigned:            End Property
Public Property Get WhenSigned() As Date:             WhenSigned = mdWhenSigned:                    End Property
Public Property Get WhoSigned() As EncProvider:       Set WhoSigned = mobjWhoSigned:                End Property
Public Property Get ApptDTS() As Date
    ApptDTS = mdApptDTS
End Property
Public Property Let ApptDTS(dApptDTD As Date)
    mdApptDTS = dApptDTD
End Property
'don't allow direct access to change UpdateLock
Public Property Get UpdateLock() As Boolean
    UpdateLock = mblnUpdateLock
End Property
Public Sub SetUpdateLock()
    If mblnUpdateLock = False Then
        mblnUpdateLock = True
        mblnUpdateLockPending = True
    End If
End Sub
Public Property Get Sensitivity() As Boolean
    Sensitivity = mlSensitivity
End Property
Public Property Let Sensitivity(Sensitive As Boolean)
    mlSensitivity = Sensitive
End Property
'Public Property Get Sensitivity() As Long
'    Sensitivity = mlSensitivity
'End Property
'Public Property Let Sensitivity(Sensitive As Long)
'    mlSensitivity = Sensitive
'End Property

Public Property Get EandMReviewed() As Boolean
    EandMReviewed = mcolSecParents(CStr(eDisp)).mobjSection.EandMReviewed
End Property

Public Property Let EandMReviewed(RHV As Boolean)
    mcolSecParents(CStr(eDisp)).mobjSection.EandMReviewed = RHV
End Property

Public Property Get EandMReviewedBy() As String
    EandMReviewedBy = mcolSecParents(CStr(eDisp)).mobjSection.EandMReviewedBy
End Property

Public Property Let EandMReviewedBy(RHV As String)
    mcolSecParents(CStr(eDisp)).mobjSection.EandMReviewedBy = RHV
End Property

Public Property Get EandMReviewDate() As Date
    EandMReviewDate = mcolSecParents(CStr(eDisp)).mobjSection.EandMReviewDate
End Property

Public Property Let EandMReviewDate(RHV As Date)
    mcolSecParents(CStr(eDisp)).mobjSection.EandMReviewDate = RHV
End Property

Public Property Get DBNICategory() As String
    DBNICategory = mcolSecParents(CStr(eDisp)).mobjSection.DBNICategory
End Property

Public Property Let DBNICategory(RHV As String)
    mcolSecParents(CStr(eDisp)).mobjSection.DBNICategory = RHV
End Property

Public Property Get DBNICause() As String
    DBNICause = mcolSecParents(CStr(eDisp)).mobjSection.DBNICause
End Property

Public Property Let DBNICause(RHV As String)
    mcolSecParents(CStr(eDisp)).mobjSection.DBNICause = RHV
End Property

Public Property Get EandMCodeReview() As String
    EandMCodeReview = msEandMCodeReview
End Property

Public Property Let EandMCodeReview(RHV As String)
    If msEandMCodeReview <> RHV Then
        msEandMCodeReview = RHV
        mbDirty = True
    End If
End Property

Public Property Get AlleryVerifiedByNCID() As String
    AlleryVerifiedByNCID = msAllergyVerifiedByNCID
End Property

Public Property Let AlleryVerifiedByNCID(ByRef RHV As String)
    If msAllergyVerifiedByNCID <> RHV Then
        msAllergyVerifiedByNCID = RHV
        mbDirty = True
    End If
End Property

Public Property Get AlleryVerifiedByDATE() As Date
    AlleryVerifiedByDATE = mdAllergyVerifiedByDate
End Property

Public Property Let AlleryVerifiedByDATE(ByRef RHV As Date)
    If mdAllergyVerifiedByDate <> RHV Then
        mdAllergyVerifiedByDate = RHV
        mbDirty = True
    End If
End Property
'Need to Save Primary Dx Medcin ID for Prev Enc look up of description using Medcin 'SCR-19371
Public Property Let PrimaryDxSnoID(ByRef value As String)
    If mnPrimaryDxSnoID <> value Then
        mbDirty = True
        mnPrimaryDxSnoID = value
    End If
End Property
Public Property Get PrimaryDxSnoID() As String
    PrimaryDxSnoID = mnPrimaryDxSnoID
End Property
'Need to Save Primary Dx Medcin Prefix for Prev Enc look up of description using Medcin 'SCR-19371
Public Property Let PrimaryDxMedcinPrefix(ByRef value As String)
    If msPrimaryDxMedcinPrefix <> value Then
        mbDirty = True
        msPrimaryDxMedcinPrefix = value
    End If
End Property
Public Property Get PrimaryDxMedcinPrefix() As String
    PrimaryDxMedcinPrefix = msPrimaryDxMedcinPrefix
End Property

Public Property Get APStatusComment() As String
    APStatusComment = msAPStatusComment  'SCR-17942
End Property
Public Property Let APStatusComment(ByRef value As String)
    mbDirty = True
    msAPStatusComment = value  'SCR-17942
End Property

Public Property Get FiledEncounter() As Boolean
    FiledEncounter = mbFiledEncounter
End Property

Public Property Get WorkLoadOptions() As String
    WorkLoadOptions = msWorkLoadOptions
End Property

Public Property Let WorkLoadOptions(ByRef value As String)
    msWorkLoadOptions = value
End Property

Public Property Get WorkLoad() As String
    If Len(msWorkLoad) = 0 Then
        'Use SADR Status to determine WL SCR-35050
        'SADR Statuses
        'Count: 209662, 209660, 209661
        'No Count: 14864181, 14510311
        If msSADRStatus = NCID_SADR_NO_COUNT Or msSADRStatus = NCID_SADR_NO_COUNT_PROCESSED Then
            msWorkLoad = "N"
        Else
            msWorkLoad = "C"
        End If
    End If
    
    WorkLoad = msWorkLoad

End Property

Public Property Let WorkLoad(ByRef value As String)

    If Len(msWorkLoad) And StrComp(msWorkLoad, value, vbTextCompare) <> 0 Then
        mbDirty = True
        If meType = TeleConsult Then
            'Call Telcons SCR#49377
            Dim objTelCons As Object
            
            Set objTelCons = Comm.InitializeOLEServer(NCID_TELCON_APPLICATION)
            
            objTelCons.TelconOps.UpdateTelconWorkload msApptID, value
            
            Set objTelCons = Nothing
            
        Else
            Dim objAppts As Object
            Set objAppts = Comm.InitializeOLEServer(NCID_CLINIC_SCHEDULE_APPLICATION)
            
            'Update appt workload SCR-34837
            objAppts.AppointmentOps.UpdateAppointmentWorkload msApptID, value
            Set objAppts = Nothing
        End If
    End If
    
    msWorkLoad = value
    
    'Set Class to set SADR Status SCR-34852
    If msWorkLoad = "N" Then
        meClass = ec_NoCount
    Else
        meClass = ec_New
    End If
    
End Property

Public Property Let PatientStatus(ByRef value As String)
    Dim objAppts As Object

    If Len(msPatientStatus) And StrComp(msPatientStatus, value, vbTextCompare) <> 0 Then 'SCR-28118
    
        If meType = Dental Or meType = TeleConsult Then Exit Property  'SCR-35350
        
        'no outpatient observation in theater mode #55849
        
        mbDirty = True
    End If
    msPatientStatus = value
End Property
Public Property Get PatientStatus() As String
    PatientStatus = msPatientStatus
End Property
Public Property Get MEPRSCode() As String
    Dim i As Integer
    Dim objAppts As Object
    Dim objAppt As Object
    
    If Not mobjShared.IsAppMode(modeCHCSI_GUI) Then
        If Len(msMEPRSCodes) < 2 Then
            Exit Property
        End If
        
        i = InStr(1, msMEPRSCodes, "|", vbTextCompare)
        If meType = Inpatient Then
            'Return A MEPRS Code
            MEPRSCode = Left(msMEPRSCodes, i - 1)
        Else
            'Return B MEPRS Code
            MEPRSCode = Right(msMEPRSCodes, Len(msMEPRSCodes) - i)
        End If
    Else
        'PGUI Mode Get MEPRS Codes from Appt Object
        Set objAppts = Comm.InitializeOLEServer(NCID_CLINIC_SCHEDULE_APPLICATION)
        Set objAppt = objAppts.AppointmentOps.GetAppointment_CHCS_GUI(msApptIEN)
        
        If meType = Inpatient Then
            MEPRSCode = objAppt.Inpatient_MEPRS_Code
        Else
            MEPRSCode = objAppt.MEPRSCode
        End If
    End If
End Property

Public Property Let MEPRSCodes(ByRef value As String)
    msMEPRSCodes = value
End Property
Public Property Get InpatientService() As String
    InpatientService = msInpatientService
End Property
Public Property Let InpatientService(ByRef value As String)
    msInpatientService = value
End Property
Public Property Get CreateClinicianNCID() As String 'SCR#42996
    CreateClinicianNCID = msCreateClinicianNCID
End Property
'Called by Disposition.OpenSection to refresh WAM data SCR#47301
Friend Sub ClearWAM()

    Set mobjWAM = Nothing

End Sub


' ****************************************************************
' Provides list of SnoIDs from S/O, A/P and Vitals that will be used to calculate E&M Code
'
' Modification:
' SCR-11863  C Atencio  8/16/01  If user clicked on "Reviewed Vitals" in S/O then load Vitals SnoIDs to
'                                 be calculated in E&M Calculator
                                 

Public Property Get AllListToolRecords() As Collection
    Dim objAp As AP
    Dim objSO As SO
    Dim objSONote As SONote
    Dim objVitals As Vitals
    Dim objLTR As ListToolRecord
    Dim bReviewedVitals As Boolean
    Dim i As Integer
    
    On Error Resume Next
    
    Set AllListToolRecords = New Collection
    Set objSO = mcolSecParents(CStr(eSO)).mobjSection
    If Not objSO Is Nothing Then
       For i = 1 To objSO.Count
           Set objSONote = objSO.SONote(i)
           If objSONote.meType = ListNoteType Then
               For Each objLTR In objSONote.mcolListRecords
                   AllListToolRecords.Add objLTR

                   'Check to see if provider clicked on reviewed vitals signs in S/O
                   If objLTR.SnoID = REVIEWED_VITALS_SnoID Or (objLTR.SnoID = REVIEWED_VITALS2_SnoID And UCase$(objLTR.Note) = REVIEWED_VITALS_KEYWORD) Or (objLTR.SnoID = REVIEWED_VITALS3_SnoID And Len(Trim$(objLTR.result)) = 0) Then
                       bReviewedVitals = True
                   End If
                   
               Next objLTR
           End If
       Next i
    End If
    Set objSONote = Nothing
    Set objSO = Nothing
    
    Set objAp = mcolSecParents(CStr(eap)).mobjSection
    If Not objAp Is Nothing Then
       For Each objLTR In objAp.SnoIDs
            AllListToolRecords.Add objLTR
       Next objLTR
    End If
    Set objAp = Nothing
    
    If bReviewedVitals Then
        'Get vitals signs SnoIDs and add them to the ListToolRecords to be calculated by E&M calculator
        Set objVitals = mcolSecParents(CStr(eVitals)).mobjSection
        If Not objVitals Is Nothing Then
           For Each objLTR In objVitals.SnoIDs
                AllListToolRecords.Add objLTR
           Next objLTR
        End If
        Set objVitals = Nothing
    End If
    
    Set objLTR = Nothing
    
End Property

Public Property Get AllListNoteRecords() As Collection
    Dim objSO As SO
    Dim objSONote As SONote
    Dim objLTR As ListToolRecord
    Dim i As Integer
    On Error Resume Next
    Set AllListNoteRecords = New Collection
    Set objSO = mcolSecParents(CStr(eSO)).mobjSection
    If Not objSO Is Nothing Then
       For i = 1 To objSO.Count
           Set objSONote = objSO.SONote(i)
           If objSONote.meType = ListNoteType Then
               For Each objLTR In objSONote.mcolListRecords
                   AllListNoteRecords.Add objLTR
               Next objLTR
           End If
       Next i
    End If
    Set objSONote = Nothing
    Set objSO = Nothing
End Property

Public Property Get APStatus() As EncStatus
    On Error Resume Next
    APStatus = mcolSecParents(CStr(eap)).mobjSection.Status
End Property

Public Property Get CurrentSONote() As SONote
    Set CurrentSONote = mcolSecParents(CStr(eSO)).mobjSection.CurrentNote
End Property

Public Property Get CurrentRTF(Optional bPOC_Continuation As Boolean = False, Optional bIncludeSignature As Boolean = False) As String
'Public Property Get CurrentRTF(Optional ByVal bForce As Boolean = False) As String
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '  SCR 11246 Printing Errors.                                '
    '                                                            '
    ' change check to look for mestatus = canceled               '
    ' by patient or facility as well.                            '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Added PLOWBS SCR-31581
    'Added Updating  SCR-32803
    If meStatus = complete _
    Or meStatus = NeedsCoSignature _
    Or meStatus = Updated _
    Or meStatus = CancelByFacility _
    Or meStatus = CancelByPatient _
    Or meStatus = CancelByProvider _
    Or meStatus = PatientLeftWithoutBeingSeen _
    Or meStatus = Updating Then
        If Not mcolEncRTFs Is Nothing Then
            If mcolEncRTFs.Count > 0 Then
                CurrentRTF = mcolEncRTFs(1).msRTF
            End If
        End If
        
        If Len(CurrentRTF) = 0 Then
            CurrentRTF = BuildRTF(Me, bPOC_Continuation, bIncludeSignature) 'SCR-24203
        End If
        
        Exit Property 'SCR-24635  Moved out of End..If
    End If

'   Caching causing signatures not to appear SCR-24727
'
'    msRTF = ""
'
'    'Use cached RTF, cleared in Save and LoadDataStrings 'SCR-24498
'    If Len(msRTF) = 0 Then
'        msRTF = BuildRTF(Me)
'    End If
    
    CurrentRTF = BuildRTF(Me, bPOC_Continuation, bIncludeSignature)

End Property

Private Function GetPOCContinuation() As String

    Dim sRTF As String
    Dim oPatientPOC As Object
    
    GetPOCContinuation = ""
    
    If mobjComm.CmdLineSwitch("CPG") <> "" Then Exit Function
    Set oPatientPOC = mobjComm.InitializeOLEServer(NCID_PATIENT_PLANOFCARE)
    sRTF = oPatientPOC.GetEncounterContRTF

    GetPOCContinuation = sRTF
    GoTo CleanUp

ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "Encounter.GetPOCContinuation", "EncounterOps", vbExclamation
CleanUp:
    Set oPatientPOC = Nothing
    
End Function

Public Property Get DispStatus() As EncStatus
    On Error Resume Next
    DispStatus = SectionParent(eDisp).mobjSection.Status
End Property
Public Property Let DispStatus(ByVal RHV As EncStatus)
    'Called by AP
    Call mcolSecParents(CStr(eDisp)).mobjSection.SetStatus(RHV)
    'Need to save since AP is setting to incomplete
    Call mcolSecParents(CStr(eDisp)).mobjSection.Save
End Property


Public Property Get DispositionText() As String
    DispositionText = mcolSecParents(CStr(eDisp)).mobjSection.DispositionText
End Property

Public Property Get DispositionNCID() As String
    DispositionNCID = mcolSecParents(CStr(eDisp)).mobjSection.DispositionNCID
End Property

Public Property Get DispAdminOption() As String
    DispAdminOption = mcolSecParents(CStr(eDisp)).mobjSection.DispAdminOption
End Property

Public Property Get EandMNCID() As String
    EandMNCID = mcolSecParents(CStr(eDisp)).mobjSection.EandMNCID
End Property

Public Property Get EandMCalc() As String
    EandMCalc = mcolSecParents(CStr(eDisp)).mobjSection.EandMCalc
End Property

Public Property Get EMCalcOverrideNote() As String
    EMCalcOverrideNote = mcolSecParents(CStr(eDisp)).mobjSection.EMCalcOverrideNote
End Property

Public Property Get FollowupTimeFrame() As String
    FollowupTimeFrame = mcolSecParents(CStr(eDisp)).mobjSection.FollowupTimeFrame
End Property

Public Property Get FollowupComments() As String
    FollowupComments = mcolSecParents(CStr(eDisp)).mobjSection.FollowupComments
End Property

Public Property Get ItemsDiscussed() As String
    ItemsDiscussed = mcolSecParents(CStr(eDisp)).mobjSection.ItemsDiscussed
End Property

Public Property Get PrimaryDiagnosisNCID() As String
    PrimaryDiagnosisNCID = msPrimaryDiagnosisNCID
End Property

Public Property Get PrimaryProvider() As EncProvider
    If Not mcolProviders Is Nothing Then
        If mcolProviders.Count > 0 Then
            Set PrimaryProvider = mcolProviders(1)
        End If
    End If
End Property

' 4/8/2 C Atencio Added NPOE Security
' 7/15/2 C Atencio Added Optional Parameter bCheckReadPriv to Check Read Priv instead of Write SCR-25679
' 8/2/2 SCR-26169 No longer check priv for Current_Encounter just Section
' 4/4/04 C Atencio Dental now used NPOE security objects SCR#49090
Public Property Get SectionIsEditable(ByVal Section As EncounterSectionEnum, Optional bCheckReadPriv As Boolean = False) As Boolean
    
    Dim lPriv As Long
    
    ' Testing scr 10878''''
    On Error Resume Next  '
    '''''''''''''''''''''''
    If meStatus = complete _
    Or meStatus = PatientLeftWithoutBeingSeen _
    Or meStatus = CancelByFacility _
    Or meStatus = CancelByPatient _
    Or meStatus = CancelByProvider _
    Or meStatus = NoShow _
    Or meStatus = Updated Then
        Exit Property   ' Or meStatus = NeedsCoSignature _  (Removed) SCR-14496
    End If

    lPriv = Priv_Write
    
    If bCheckReadPriv Then
        lPriv = Priv_Read
    End If
    
    Select Case Section
    
    Case eAutoCite
    
        SectionIsEditable = mobjUser.HasPrivilegeEx(Priv_Current_Encounter, lPriv)
                        '''''''''''''''''''''''''''''''''''''''''
                        ' SCR 5049 Turn on autocite for         '
                        ' all rather than simply the primary.   '
                        'UserIsPrimaryProvider                  '
                        '''''''''''''''''''''''''''''''''''''''''
                        
    Case eRFV
         'Theater not using NPOE Security
        SectionIsEditable = (Not meType = TeleConsult) _
                        And mobjUser.HasPrivilegeEx(SectionEnumToSecurityName(Section), lPriv)
    
    Case eVitals
        SectionIsEditable = (Not EncounterType = TeleConsult) _
                        And mobjUser.HasPrivilegeEx(SectionEnumToSecurityName(Section), lPriv)
                        
    Case eSO, eNotes
        'Theater not using NPOE Security
        SectionIsEditable = mobjUser.HasPrivilegeEx(SectionEnumToSecurityName(Section), lPriv)
                        
    Case eap, eDisp, eEducation
        SectionIsEditable = mobjUser.HasPrivilegeEx(SectionEnumToSecurityName(Section), lPriv)
        
'    Case eDental
'        SectionIsEditable = EncounterType = Dental _
'                        And mobjParent.EncounterOps.UseDental _
'                        And mobjUser.HasPrivilegeEx(Priv_Current_Encounter, lPriv)
                        
    Case eDental
        SectionIsEditable = EncounterType = Dental _
                        And mobjUser.HasPrivilegeEx(Priv_Dental_Access, lPriv) _
                        And mobjUser.HasPrivilegeEx(Priv_Current_Encounter, lPriv)
    End Select

End Property


Public Property Get SectionRTF(ByVal eSection As EncounterSectionEnum, Optional bBypassSecurityCheck As Boolean = False) As String
    On Error Resume Next
    
    Dim bShowRTF As Boolean
    
    ' bBypassSecurityCheck is
    'Used by Appts to check if Disp exists for Encounter
    'needed b/c user may not have access to Disp RTF
    ' and when building RTF SCR-41052
    
    bShowRTF = True
     
     If Not mobjShared.IsAppMode(modeTheater) Or Not (eSection = eRads) Then
        'Check NPOE Security
        
        If CheckNPOESecurity(eSection) Or bBypassSecurityCheck Then
            Select Case eSection
                Case eSO
                    'If questionnaires selected, place above objective portion SCR-12617
                    If mobjProperties.AutoCites(CStr(NCID_PATIENT_QUESTIONNAIRES)).Selected = vbChecked Then
                        'Questionnaire may not have RTF SCR-24623
                        'SCR 34161 - Further reduce blank line if S/O entries is blank
                        If Len(mcolSecParents(CStr(eQuestionnaire)).RTF) > 0 Then
                            SectionRTF = mcolSecParents(CStr(eQuestionnaire)).RTF
                            If Len(mcolSecParents(CStr(eSection)).RTF) > 0 Then
                               SectionRTF = SectionRTF & "\par " & mcolSecParents(CStr(eSection)).RTF
                            End If
                        Else
                            SectionRTF = mcolSecParents(CStr(eSection)).RTF
                        End If
''                        If Len(mcolSecParents(CStr(eQuestionnaire)).RTF) > 0 Then
''                            'Reduce spacing if SO RTF is blank SCR-33397
''                            If Len(mcolSecParents(CStr(eSection)).RTF) > 0 Then
''                                SectionRTF = mcolSecParents(CStr(eQuestionnaire)).RTF & "\par " & mcolSecParents(CStr(eSection)).RTF
''                            Else
''                                SectionRTF = mcolSecParents(CStr(eQuestionnaire)).RTF
''                            End If
''                        Else
''                            SectionRTF = mcolSecParents(CStr(eSection)).RTF
''                        End If
                    Else
                        SectionRTF = mcolSecParents(CStr(eSection)).RTF
                    End If
                
                Case eQuestionnaire
                    
                    'Questionnaire RTF is already included in SO RTF SCR-25294
                    'Questionnaire RTF  is it's own seperate section that appears to be part of the SO section but is not. SCR-31096
                    SectionRTF = "" 'mcolSecParents(CStr(eSection)).RTF
                Case eap, eDisp, eVitals
                    'Only display RTF if user has read privileges for SCR-22462, 22463, 22464, 25790
                    If mobjUser.HasPrivilegeEx(SectionEnumToSecurityName(eSection), Priv_Read) Or bBypassSecurityCheck Then
                        SectionRTF = mcolSecParents(CStr(eSection)).RTF
                    Else
                        bShowRTF = False
                    End If
                Case Else
                    SectionRTF = mcolSecParents(CStr(eSection)).RTF
            End Select
        Else
            bShowRTF = False
        End If
        
        If mobjComm.CmdLineSwitch("GREEN") <> "" Then
            'Add a blank line between section
            'Debug.Print "Section: " & eSection
            If SectionRTF <> "" Then
               SectionRTF = SectionRTF & " \par"
               'Debug.Print SectionRTF
            End If
        End If
        
        
        If Not bShowRTF And Not bBypassSecurityCheck Then
            'Only display for NPOE sections
            SectionRTF = "{\rtf1\ansi\ucl\deff0\deflang1033\deflangfe1033" _
                & "{\fonttbl{\f0\fswiss\fcharset0\fprq2 Arial;}}" _
                & "{\colortbl;\red0\green0\blue0;}" _
                & "\pard\plain\f0\fs16\b " _
                & " You do not have the access privileges to read this section." & "\b0\par } "
        End If
      
    End If
End Property

Public Function CheckNPOESecurity(ByVal Section As EncounterSectionEnum) As Boolean
    
    CheckNPOESecurity = True
    
    'NPOE Security not implemented in Theater
    If Not mobjShared.AppMode = modeTheater Then
        Select Case Section
        
            Case eSO, eAutoCite, eNotes, eRFV
                CheckNPOESecurity = mobjUser.HasPrivilegeEx(SectionEnumToNPOESecurityName(Section), Priv_Write) _
                                    Or mobjUser.HasPrivilegeEx(SectionEnumToNPOESecurityName(Section), Priv_Read)

        End Select
    End If
    
End Function
Public Property Get WorkStatusDisposition() As String
    WorkStatusDisposition = mcolSecParents(CStr(eDisp)).mobjSection.WorkStatusDisposition
End Property

Public Property Get WorkStatusScreening() As String
    On Error Resume Next
    WorkStatusScreening = mcolSecParents(CStr(eRFV)).mobjSection.WorkStatusScreening
End Property

Public Property Let APStatus(ByVal RHV As EncStatus)
    mcolSecParents(CStr(eap)).mobjSection.APStatus = RHV
End Property

Public Property Let Class(ByVal RHV As EncClass)
    If meClass <> RHV Then
        mbDirty = True
        meClass = RHV
'<< SCR 10086
        If meClass = ec_NoCount Then
            msSADRStatus = NCID_SADR_NO_COUNT
        ElseIf meClass = ec_New Then
            If mbWasUpdated Then
                msSADRStatus = NCID_SADR_NEEDS_UPDATE_PROCESSING
            Else
                msSADRStatus = NCID_SADR_NEEDS_NEW_PROCESSING
            End If
        End If
        
    End If
End Property

Public Property Get ShowDentalSection() As Boolean
    ShowDentalSection = (meType = Dental _
                    And mobjComm.CmdLineSwitch("DENTAL") <> "")
End Property

Friend Property Set Cosigner(ByRef RHV As EncProvider)
    Set mobjCosigner = RHV
    mbDirty = True
End Property
'<< SCR 10768
Public Property Get CurrSection() As SectionParent
    Set CurrSection = mobjCurrSection
End Property
'>>
Public Property Set CurrSection(ByRef RHV As SectionParent)
    Set mobjCurrSection = RHV
End Property

Friend Property Let Dirty(ByVal RHV As Boolean)
    mbDirty = RHV
End Property

Public Property Let DispositionText(RHV As String)
    mcolSecParents(CStr(eDisp)).mobjSection.DispositionText = RHV
End Property

Public Property Let DispositionNCID(RHV As String)
    mcolSecParents(CStr(eDisp)).mobjSection.DispositionNCID = RHV
End Property

Public Property Get DispMetadata() As String
    DispMetadata = mcolSecParents(CStr(eDisp)).mobjSection.DispMetadata
End Property
Public Property Let DispMetadata(RHV As String)
    mcolSecParents(CStr(eDisp)).mobjSection.DispMetadata = RHV
End Property


Public Property Let DispAdminOption(RHV As String)
    mcolSecParents(CStr(eDisp)).mobjSection.DispAdminOption = RHV
End Property

Public Property Let EandMNCID(RHV As String)
    mcolSecParents(CStr(eDisp)).mobjSection.EandMNCID = RHV
End Property

Public Property Let EandMCalc(RHV As String)
    mcolSecParents(CStr(eDisp)).mobjSection.EandMCalc = RHV
End Property

Public Property Let EMCalcOverrideNote(RHV As String)
    mcolSecParents(CStr(eDisp)).mobjSection.EMCalcOverrideNote = RHV
End Property

Friend Property Let EncounterID(ByVal RHV As Long)
    mnEncounterID = RHV
End Property

Public Property Let EncounterType(ByVal RHV As EncType)

    Dim sDestApp As String
    
    'Cannot change EncounterType for Dental or TeleCon 'SCR-29439 SCR-30768 SCR#42995
    If meType = Dental Or meType = TeleConsult Or meType = InpatientNote Then
        Exit Property
    End If
    
    If meType <> RHV Then

        mbDirty = True
        meType = RHV
    
        'Send message to Appointments that Classification has changed. SCR-2235
        sDestApp = NCID_CLINIC_SCHEDULE_APPLICATION
        
        If Not mobjComm.Message(cwiOLE_SERVER_RUNNING, sDestApp, NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION) Then
            'if it's not running initialize it
            Call mobjComm.Message(cwiSTART_OLE_SERVER, sDestApp, NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION)
            
        End If
         
        '- tell it to update the single appt
        mobjComm.Message cwiGENERIC, "UPDATE_APPT_CLASSIFICATION|" & msApptID & "|" & gsEnctype(RHV), sDestApp, NCID_ENCOUNTER_APPLICATION
    
    End If
End Property

Public Property Let FollowupTimeFrame(RHV As String)
    mcolSecParents(CStr(eDisp)).mobjSection.FollowupTimeFrame = RHV
End Property

Public Property Let FollowupComments(RHV As String)
    mcolSecParents(CStr(eDisp)).mobjSection.FollowupComments = RHV
End Property

Public Property Let ItemsDiscussed(RHV As String)
    mcolSecParents(CStr(eDisp)).mobjSection.ItemsDiscussed = RHV
End Property

Friend Property Let PatientID(RHV As Long)
 mnPatientID = RHV
End Property

Public Property Let PrimaryDiagnosisNCID(ByRef RHV As String)
    If msPrimaryDiagnosisNCID <> RHV Then
        mbDirty = True
        msPrimaryDiagnosisNCID = RHV
    End If
End Property

Public Property Set RelatedDiagnoses(RHV As Collection)
    Set mcolSecParents(CStr(eap)).mobjSection.RelatedDiagnoses = RHV
End Property

Public Property Let SADRStatus(RHV As String)
    If msSADRStatus <> RHV Then
        mbDirty = True
        
        'All Dental encounters are considered no counts SCR-20682
        If meType = Dental Then
            msSADRStatus = NCID_SADR_NO_COUNT
        Else
            msSADRStatus = RHV
        End If
    End If
End Property

Public Property Let TemplateData(RHV As String)
    If msTemplateData <> RHV Then
        mbDirty = True
        msTemplateData = RHV
    End If
End Property

Friend Property Let WhenCoSigned(ByVal RHV As Date)
    mdWhenCosigned = RHV
    mbDirty = True
End Property

Friend Property Set WhoCoSigned(ByRef RHV As EncProvider)
    Set mobjWhoCosigned = RHV
    mbDirty = True
End Property

Friend Property Let WhenSigned(ByVal RHV As Date)
    mdWhenSigned = RHV
    mdEndDTS = RHV  'SCR-32291
    mbDirty = True
End Property

Friend Property Set WhoSigned(ByRef RHV As EncProvider)

    Set mobjWhoSigned = RHV
    mbDirty = True
    
    'When Doing Print Preview this is being set but should not be saved since encounter is not yet signed
    If RHV Is Nothing Then
        mbDirty = False
    End If
End Property

Public Property Let WorkStatusDisposition(RHV As String)
    mcolSecParents(CStr(eDisp)).mobjSection.WorkStatusDisposition = RHV
End Property

Public Property Let WorkStatusScreening(RHV As String)
    mcolSecParents(CStr(eRFV)).mobjSection.WorkStatusScreening = RHV
End Property

Public Property Let Category(ByVal RHV As String)
    Dim oApptStat As ApptStatus
    On Error Resume Next
    If mbLoading Or RHV = "" Or msCategory = RHV Then Exit Property
    msCategory = RHV
    mbDirty = True
    If Not (msApptID = "") Then
        '- this is  a change to the category, tell appointments
        Set oApptStat = New ApptStatus
        If Not oApptStat.SetApptClassification(msApptID, CInt(RHV)) Then
            MsgBxARMd "Could not update appointment classification. " & Err.Description
        End If
        Set oApptStat = Nothing
    End If
End Property

Public Property Let Status(ByVal RHV As EncStatus)
    Dim oApptStat As ApptStatus
    Dim oConsOrdStat As Object
    Dim strContextNCID As String
    Dim dDate As Date
    Dim eOldStatus As EncStatus
    Dim lConsultOrderStatus As Integer
     'SCR-24203
    'variable added to hold resulted consult note
    Dim strEncounterNotes As String
    
'    On Error Resume Next
On Error GoTo ErrHnd

    #If debugon Then
        Dim DR As DebugRec
        If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Let Status", "Begin", DR, "Enc", mnEncounterID
    #End If
    '<SCR SCR 36493
    If mobjShared.IsAppMode(modeCHCSI_GUI) Then
        'Do not set status to Inprogress or Waiting SCR#49815
        If RHV = InProgress Or RHV = Waiting Then
            Exit Property
        End If
    End If
    '>
    If meStatus <> Updated Then
        If meStatus = RHV Then Exit Property
    End If
    
    eOldStatus = meStatus
    
    meStatus = RHV
    mbDirty = True
   #If debugon Then
        If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Let Status", "Raiseevent", DR
    #End If
    
'-- the event handler back in enctr will tell a/p.  DDSA also uses this to trigger their actions
    'Moved to save function SCR-21322
    'RaiseEvent EncStatusChanged(RHV)
    If Err Then
        Call mobjShared.ShowVBError(Err.Number, Err.Description, "Encounter.PrintEncounter", "Encounter", vbExclamation)
        Err.Clear
    End If
    If mbLoading Then Exit Property
    Select Case meStatus
    Case complete
'<< SCR 10086
        If Not meClass = ec_NoCount Then
            msSADRStatus = NCID_SADR_NEEDS_NEW_PROCESSING
        Else
            msSADRStatus = NCID_SADR_NO_COUNT
        End If
    Case Updated
        mbWasUpdated = True
        If Not meClass = ec_NoCount Then
            msSADRStatus = NCID_SADR_NEEDS_UPDATE_PROCESSING 'SCR-27708
        Else
            msSADRStatus = NCID_SADR_NO_COUNT
        End If
        
'>>
   #If debugon Then
        If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Let Status", "AMS Send", DR
    #End If

        'msSADRStatus = NCID_SADR_NEEDS_UPDATE_PROCESSING
        Dim oAMS As AlertsClient
        If mcolProviders.Count > 0 Then
            Set oAMS = New AlertsClient
            'Use ApptDTS SCR-28950 'ApptDTS may not exist for earlier encounters SCR-32567
            If mdApptDTS <> 0 Then
                dDate = mdApptDTS
            Else
                dDate = mdStartDTS
            End If
            Call oAMS.Send(mcolProviders.Item(1).NCID, "{7EFF358C-8A26-11d3-8304-0050DA0C5BB9}", "The encounter (dated: " & dDate & ") for " & msPatientName & " has been updated.", 0, 0, CStr(mnPatientID), "", CStr(mnEncounterID))
                                                        
            If Err Then
                Call mobjShared.ShowVBError(Err.Number, Err.Description, Module, Exe, vbExclamation)
                Err.Clear
            End If
        End If
        
    Case Updating 'Set WasUpdated flag SCR-24727
        mbWasUpdated = True
        
    End Select
    If meClass = ec_NoCount Then
        msSADRStatus = NCID_SADR_NO_COUNT
    End If
    #If debugon Then
        If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Let Status", "oApptStat.SetEncounterStatus", DR
    #End If
    
    If msApptID <> "" And giEncApptStatus(RHV) <> 0 Then
    
        'Complete and updated status will we updated in during SignAndSave process SCR-39893
        If Not (RHV = complete Or RHV = Updated) Then
            UpdateApptEncounterStatus
        End If

        #If debugon Then
            If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Let Status", "oConsOrdStat.SetEncounterStatus", DR
        #End If
    End If
   #If debugon Then
        If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Let Status", "end", DR
    #End If

Exit Property
ErrHnd:
    MsgBox "Application error in: " & Err.Description & " Occurred in Encounter.Status Property Let"
    Exit Property
    Resume
End Property

Private Sub Class_Initialize()
    #If debugon Then
        Set DL = GetObject("", "dlog.cdebuglog")
        DL.debugStart
        Dim DR As DebugRec
        If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Class_Initialize", "Begin/End", DR
    #End If
    
    mbSignEncbyEncParent = False 'SCR 22247 flag ARM to time sign from Enc only (not children)  MSolano 5/17/02
    '--- SCR-ARM-16126 MSolano 8/23/01
    Call ARM_InitAllTransactionsForEncounterOps
    
    Set mcolEncRTFs = New Collection
    Set mcolProviders = New EncProviders
    Set mcolSecParents = New Collection
    meStatus = Incomplete
    mbFlagNew = True
    mbDirty = False
    Set mobjAccidentData = New EncAccidentData
    Set mobjPregnancyData = New EncPregnancy
End Sub

Private Sub Class_Terminate()
    Set mobjProperties = Nothing
    Set mcolSecParents = Nothing
    Set mcolEncRTFs = Nothing
    Set mcolProviders = Nothing
    Set mobjCosigner = Nothing
    Set mobjWhoSigned = Nothing
    Set mobjWhoCosigned = Nothing
    Set mobjUser = Nothing
    Set mobjComm = Nothing
    Set mobjShared = Nothing
    Set mobjLogon = Nothing
    Set mobjPatient = Nothing
    Set mobjLocking = Nothing
    Set mobjAccidentData = Nothing
    Set mobjPregnancyData = Nothing
    
    #If debugon Then
        Dim DR As DebugRec
        If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Class_Terminate", "", DR, "Enc #", mnEncounterID
        Set DL = Nothing
    #End If
End Sub

Public Sub Init(User As ICHCSII_User, Comm As MMMHISComm.Comm, CWShared As MMMHISShared.CWShared, Logon As MMMHISLogon.Logon, Patient As MMMHISPatient.Patient, Properties As EncProperties)
    Set mobjUser = User
    Set mobjComm = Comm
    Set mobjShared = CWShared
    Set mobjLogon = Logon
    Set mobjPatient = Patient
    Set mobjProperties = Properties
    msPatientName = mobjPatient.Name
End Sub

Friend Function AddEncRTF(ByRef RTF As EncRTF) As Boolean
    If mcolEncRTFs Is Nothing Then
        Set mcolEncRTFs = New Collection
    End If
    ' add new rtfs at the beginning of the collection
    If mcolEncRTFs.Count = 0 Then
        mcolEncRTFs.Add RTF
    Else
        mcolEncRTFs.Add RTF, , 1
    End If
    mbDirty = True
End Function
'Alllow ability to replace enc RTF ' SCR-21227
'Called by A/P StartEncounterSave
Public Function ReplaceEncRTF() As Boolean

    Dim oEncRTF As EncRTF
    Dim sNewRTF As String
    Dim oLocking As Locking
    For Each oEncRTF In mcolEncRTFs
    
        'Find the most recently signed encounter
        If (oEncRTF.msUserNCID = Me.WhoSigned.NCID) And (oEncRTF.mdDTS = Me.WhenSigned) Then
            sNewRTF = BuildRTF(Me)
            oEncRTF.msRTF = sNewRTF
            Me.LockingObject.LockSection eHeader, 0, SectionLock
            Save True
            ReplaceEncRTF = True
            Exit Function
        End If
    Next
    
End Function
Public Function AddToSection(ByRef AppNCID As String, ByRef RTF As String) As Boolean
    Dim eSection As EncounterSectionEnum
    Dim objLabsRads As LabsRads
    Dim objEducation As Education
    Dim objSO As SO
    Dim objSONote As SONote
    eSection = SectionNCIDToEnum(AppNCID)
    Select Case eSection
    Case eLabs, eRads
        Set objLabsRads = mcolSecParents(CStr(eSection)).mobjSection
        AddToSection = objLabsRads.AddRTF(RTF)
        Set objLabsRads = Nothing
    Case eQuestionnaire
        Set objLabsRads = mcolSecParents(CStr(eSection)).mobjSection
        AddToSection = objLabsRads.AddRTF(RTF)
        Set objLabsRads = Nothing
    Case eEducation 'For Education Groups
        Set objEducation = mcolSecParents(CStr(eSection)).mobjSection
        objEducation.OpenEncounter Me.SectionParent(eEducation), Me
        objEducation.SetRTF RTF
        AddToSection = objEducation.Save
        Set objEducation = Nothing
    Case eSO
        Set objSO = mcolSecParents(CStr(eSO)).mobjSection
        objSO.OpenEncounter Me.SectionParent(eSO), Me
        Set objSONote = New SONote
        objSONote.mbCouplerData = True
        objSONote.meType = ListNoteType
        objSONote.mbReadOnly = True
        objSONote.msRTF = RTF
        objSONote.mdDTS = Now
        objSONote.msOwnerNCID = PrimaryProvider.NCID
        objSONote.msOwnerName = PrimaryProvider.FullName
        
        AddToSection = objSO.SaveNote(objSONote)
        Set objSO = Nothing
    Case Else
        MsgBxARMd "Internal Error: An application module (" & AppNCID & ") attempted to add RTF data to an encounter section where this capability is not supported.", _
            vbInformation, "Encounter"
    End Select
End Function

Public Function Amend() As Boolean
    Dim objEncParent As Object
    Dim sRTF As String
    On Error GoTo ErrHandler
    
    If mobjLocking Is Nothing Then
       Set mobjLocking = New Locking
       mobjLocking.msFacilityNCID = FacilityNCID
       mobjLocking.msEncounterNumber = mnEncounterID
       mobjLocking.msUserNCID = mobjLogon.UserNCID
    End If
    
    If Not mobjLocking.LockSection(eEntireEncounter, 0, AmendLock) Then
        Exit Function
    End If
    
    Call Me.Refresh
    
    If Len(mobjComm.CmdLineSwitch("INPATIENT2")) > 0 Then
        'RQT 42999
        If Me.EncounterType = InpatientNote And Me.CreateClinicianNCID <> Logon.UserNCID Then
            MsgBxARMd "Only the owner can amend an inpatient note.", vbInformation
            GoTo CleanUp
        End If
    End If
    
    'Added CheckedIn in PGUI Mode SCR#45833
    If Not (meStatus = NeedsCoSignature Or meStatus = complete Or meStatus = Updated) And Not ((meStatus = Incomplete Or meStatus = CheckedIn) And mobjShared.IsAppMode(modeCHCSI_GUI)) Then
        MsgBxARMd "You cannot amend an encounter in " & EncStatusName(meStatus) & " status.", vbInformation, "Encounter"
        GoTo CleanUp
    End If
    If Not (Me.UserIsPrimaryProvider Or Me.UserIsCosigner) Then
        MsgBxARMd "Only the signer or co-signer of an encounter can amend it.", vbInformation, "Encounter"
        GoTo CleanUp
    End If
    
    'if cosigner is amending an encounter with NeedsCoSignature status
    'do not "delete" PP signature and leave status in NeedsCoSignature status SCR-27403
    If Not (Me.UserIsCosigner And Me.Status = NeedsCoSignature) Then
        'If PP is amending Completed encounter then "delete" PP & CS signatures
        'If CS is amending Completed encounters then only delete CS signature and set status to NeedsCoSignature SCR-27403
        If Me.UserIsPrimaryProvider Then
            If Not Me.WhoSigned Is Nothing Then
                sRTF = sRTF & FormatSignature("Signed", Me.WhoSigned, Me.WhenSigned)
            End If
            
            Me.WhenSigned = 0
            Set Me.WhoSigned = Nothing
            
            'Need to delete signatures from Encounters table SCR-573 ITT
            Dim sSql  As String
            Dim objDAS           As ICHCSII_DAS
                        
            Set objDAS = gobjCHCSIIConn.CHCSII_DAS(Auto) '''<SCR 36493
            
            sSql = "UPDATE ENCOUNTERS SET WHOSIGNEDNCID = NULL WHERE ENCOUNTERNUMBER = " & Me.EncounterID
            objDAS.ExecuteSQL (sSql)
            Set objDAS = Nothing
        End If
        
        If Not Me.WhoCoSigned Is Nothing Then
          If Not Me.WhoCoSigned.NCID = 0 Then
            sRTF = sRTF & FormatSignature("Co-Signed", Me.WhoCoSigned, Me.WhenCoSigned)
          End If
        End If
        
        Me.WhenCoSigned = 0
        Set Me.WhoCoSigned = Nothing
        
        Call AddHistoryItem(EncAmended, "", sRTF)
        
        'Change status to Updating when amending SCR-130800
        Me.Status = Updating
        'Change status to Updating if user is PP SCR-14496, 25920
'        If Me.UserIsPrimaryProvider Then
'            Me.Status = Updating
'        Else 'User is CS
'            Me.Status = NeedsCoSignature
'        End If
    
    End If
    
    'Cosigner is amending doc, send alert to orignal PP 'SCR-24835
    If Me.UserIsCosigner Then
        Dim oAMS As AlertsClient
        If mcolProviders.Count > 0 Then
            'Need to send alert now since a user may take over ownership
            Set oAMS = New AlertsClient
            'Use ApptDTS SCR-28950
            Call oAMS.Send(mcolProviders.Item(1).NCID, "{7EFF358C-8A26-11d3-8304-0050DA0C5BB9}", "The encounter (dated: " & mdApptDTS & ") for " & msPatientName & " has been updated.", 0, 0, CStr(mnPatientID), "", CStr(mnEncounterID))
            If Err Then
                Call mobjShared.ShowVBError(Err.Number, Err.Description, Module, Exe, vbExclamation)
                Err.Clear
            End If
        End If
        'Clear cosigner
        'Set Me.Cosigner = Nothing COMMENTED OUT SCR-25920
    End If
    
    'HIPAA - Determine if we need to update pregnancy values - only update from patient if no encounters have since been completed and this encounter had pregnancy indicated
    If mobjShared.IsAppMode(modeCHCSI_GUI) And Not mobjShared.IsAppMode(modeITT) And Len(mobjComm.CmdLineSwitch("HIPAA837")) > 0 Then
      If Not PregnancyData.PregnancyIndicator Then
         'SCR 46947 - No previous pregnancy data so clear
         PregnancyData.Clear
      End If
    End If
    
    'Set Was Updated flag SCR-28744
    mbWasUpdated = True
    
    If Not Me.Save Then
        MsgBxARMd "Could not save encounter as 'Inprogress/Updating'.", vbCritical, "Encounter"
        GoTo CleanUp
    End If
    'End If
    Amend = True
    If mobjPatient.UnitNumber <> Me.PatientID Then
        Call mobjComm.Message(cwiSELECT_PATIENT, Me.PatientID & ";", NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION)
    End If
    
    SyncADMEncounter
    
    Set objEncParent = mobjComm.InitializeOLEServer(NCID_ENCOUNTER_APPLICATION)        'Changed to false SCR 27495
    Call objEncParent.SetCurrentEncounter(Me.PatientID, Me.FacilityNCID, Me.EncounterID, True, "", True)
    GoTo CleanUp
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "Encounter.Amend", App.Title, vbCritical
CleanUp:
    'Resume Next
    Call mobjLocking.UnlockSection(eEntireEncounter, 0)
    Set objEncParent = Nothing
    
    Exit Function
    Resume
End Function

Private Sub SyncADMEncounter()
    Dim objADMEncounter As CHCS_Interface.ICHCSEncounters

    On Error GoTo ErrHandler
    
    If Not mobjShared.IsAppMode(modeCHCSI_GUI) Then Exit Sub
    
    Set objADMEncounter = mobjShared.CHCSConnection
    objADMEncounter.SyncEncounter ApptIEN, EncounterID, PrimaryProvider.NCID, mobjPatient.UnitNumber
    
    Exit Sub
    
ErrHandler:
    If InStr(Err.Description, "ADM record locked in CHCS") Then 'SCR#41172
        MsgBxARMd "The appointment record for this appointment is open in CHCS. Please close the open record in CHCS before attempting to update.", vbExclamation, App.Title
    Else
        mobjShared.ShowVBError Err.Number, Err.Description, "Encounter.SyncADMEncounter", App.Title, vbCritical
    End If
End Sub

Public Function AppendNote() As Boolean
    Dim objSignOps As EncSignOps
    On Error GoTo ErrHandler
    If Not mbOpen Then
        MsgBxARMd "You cannot append a note to an encounter that is not open.", vbInformation, "Encounter"
        Exit Function
    End If
    If Not mobjLocking.LockSection(eEntireEncounter, 0, AppendLock) Then
        Exit Function
    End If
    Set objSignOps = New EncSignOps
    AppendNote = objSignOps.AppendNote(Me)
    GoTo CleanUp
ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description, "Encounter.AppendNote", Exe, vbCritical)
CleanUp:
    Call mobjLocking.UnlockSection(eEntireEncounter, 0)
    Set objSignOps = Nothing
End Function

Friend Function Create(CurPatient As MMMHISPatient.Patient, _
                        ApptID As String, _
                        ApptType As String, _
                        Enc_Type As EncType, _
                        FacilityNCID As String, _
                        ClinicNCID As String, _
                        PrimaryProviderNCID As String, _
               Optional ClerkNote As String = "") As Boolean
               
    Dim mobjConcept As ConceptCtrl
    Dim objProvider As EncProvider
    If mnEncounterID = 0 Then Exit Function
    On Error GoTo ErrHandler
    mbLoading = True
    Set mobjLocking = New Locking
    mobjLocking.msFacilityNCID = FacilityNCID
    mobjLocking.msEncounterNumber = mnEncounterID
    mobjLocking.msUserNCID = mobjLogon.UserNCID
    mobjLocking.LockSection 99, 0, SectionLock
    meType = Enc_Type
    msCategory = ApptType
    msApptID = ApptID
    mdStartDTS = Now

    Set mobjConcept = New ConceptCtrl
    msFacilityNCID = FacilityNCID
    mobjConcept.UniqueId = FacilityNCID
    msFacility = mobjConcept.PrefRep("2000").Representation
    If Trim$(ClinicNCID) <> "" Then
        msClinicNCID = ClinicNCID
        mobjConcept.UniqueId = ClinicNCID
        msClinic = mobjConcept.PrefRep("2000").Representation
    End If
    Set objProvider = New EncProvider
    objProvider.FacilityNCID = FacilityNCID
    objProvider.NCID = PrimaryProviderNCID
    mobjConcept.UniqueId = PrimaryProviderNCID
    objProvider.FullName = mobjConcept.PrefRep("2000").Representation
    mobjConcept.UniqueId = FacilityNCID
    objProvider.Facility = mobjConcept.PrefRep("2000").Representation
    objProvider.RoleNCID = NCID_Attending_Provider   'Set default role for HIPAA
    mcolProviders.Add objProvider, CStr(objProvider.NCID)
    
   If Len(mobjShared.CmdLineSwitches("HIPAA837")) > 0 And mobjShared.IsAppMode(modeCDR) And Not mobjShared.IsAppMode(modeITT) Then
      If AccidentData.IsDirty Then AccidentData.CDRSave CurPatient.UnitNumber, Me.FacilityNCID, Me.EncounterID
      'SCR 51265 - Always initialize the Pregnancy Object but only force a save of data for females
      PregnancyData.InitFromPatient CurPatient, True
      If CurPatient.Sex = "F" Then
         PregnancyData.CDRPlusSave Me.FacilityNCID, Me.EncounterID, Me.Status
      End If
      'Get Provider HIPAA Code SCR#43283
      msProvHIPAACode = GetProviderHIPAACode(PrimaryProviderNCID, FacilityNCID)
   End If
   


    msTimeZone = GetLocalTimeZone()
    mnPatientID = CurPatient.UnitNumber
    msPatientName = CurPatient.Name
    If meType = TeleConsult Then
        meStatus = InProgress
    Else
        meStatus = CheckedIn
    End If
    Call CreateSections
    If ClerkNote <> "" Then
        Me.AddTextNote "Appointment Note", ClerkNote
    End If
    mbLoading = False
    If Save() Then
        mbDirty = False
        mbFlagNew = False
        Create = True
    Else
        MsgBxARMd "Creation of encounter failed."
    End If
    mobjLocking.UnlockSection 99, 0
    Set mobjConcept = Nothing
    Set objProvider = Nothing
    Exit Function
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "Encounter.Create", App.Title, vbCritical
    Set mobjConcept = Nothing
    Set objProvider = Nothing
End Function

Friend Function GEMS_Create(CurPatient As MMMHISPatient.Patient, _
                             ApptID As String, _
                             ApptType As String, _
                             Enc_Type As EncType, _
                             FacilityNCID As String, _
                             ClinicNCID As String, _
                             PrimaryProviderNCID As String, _
                    Optional ClerkNote As String = "", _
                    Optional ApptDTS As Date, _
                    Optional Dummy As Boolean = False, _
                    Optional InptNoteTypeNCID As String = "", _
                    Optional InptHospDay As Long = 0) As Boolean
    
    Dim mobjConcept As New GEMS_ConceptCtrl
    Dim objProvider As EncProvider
    Dim szSQL As String
    Dim szHIPPASqlFields As String
    Dim szHIPAASqlValues As String
    
    If mnEncounterID = 0 Then Exit Function
    On Error GoTo ErrHandler
    mbLoading = True
    Dim oDas As ICHCSII_DAS
    
    'date:01/09/2004 SCR #:47870 developer: jrm Description: Added for ITT mode
    'Auto is needed to determine ITT, or Theatre
    
    Set oDas = gobjCHCSIIConn.CHCSII_DAS(Auto) '''<SCR 36493
    
   If Len(mobjComm.CmdLineSwitch("HIPAA837")) > 0 Then
      If mobjShared.IsAppMode(modeITT) Then
         AccidentData.InitFromSQL ClinicNCID, sApptId:=ApptID
         PregnancyData.InitFromPatient CurPatient, True
      ElseIf mobjShared.IsAppMode(modeCHCSI_GUI) Then
         'Load up any accident data
         AccidentData.InitFromSQL ClinicNCID, sApptId:=ApptID
         'SCR 46943 - Load patient pregnancy data for the encounter and prepare for encounter creation
         PregnancyData.InitFromPatient CurPatient
         szHIPPASqlFields = PregnancyData.GetSQLFields
         szHIPAASqlValues = PregnancyData.GetSQLValues
      End If
      
      'Get Provider HIPAA Code SCR#43283
      'msProvHIPAACode = GetProviderHIPAACode(PrimaryProviderNCID, FacilityNCID)
   End If
   
    meType = Enc_Type
    
    msCategory = ApptType
    
    msApptID = ApptID
    
    mdStartDTS = Now
    
    msFacilityNCID = FacilityNCID
    mobjConcept.UniqueId = FacilityNCID
    If mobjConcept.UniqueId <> "" Then msFacility = mobjConcept.PrefRep("2000").Representation
    If Trim$(ClinicNCID) <> "" Then
        msClinicNCID = ClinicNCID
        mobjConcept.UniqueId = ClinicNCID
        If mobjConcept.UniqueId <> "" Then msClinic = mobjConcept.PrefRep("2000").Representation
    End If
    
    Set objProvider = New EncProvider
    objProvider.FacilityNCID = FacilityNCID
    objProvider.NCID = PrimaryProviderNCID
    mobjConcept.UniqueId = PrimaryProviderNCID
    If mobjConcept.UniqueId <> "" Then objProvider.FullName = mobjConcept.PrefRep("2000").Representation
    mobjConcept.UniqueId = FacilityNCID
    If mobjConcept.UniqueId <> "" Then objProvider.Facility = mobjConcept.PrefRep("2000").Representation
    mcolProviders.Add objProvider, CStr(objProvider.NCID)
    
    'SCR 69422
    If ApptDTS <> Empty Then
        msTimeZone = GetTimeZoneBasedOnDate(ApptDTS)
    Else
        msTimeZone = GetLocalTimeZone()
    End If
    
    mnPatientID = CurPatient.UnitNumber
    
    msPatientName = CurPatient.Name
    
    mdApptDTS = ApptDTS
    
    If Not mobjShared.IsAppMode(modeCHCSI_GUI) Then
        If meType = TeleConsult Or meType = InpatientNote Then
            meStatus = InProgress
        Else
            If meStatus <> complete Then 'SCR#53547
                meStatus = CheckedIn
            End If
        End If
    Else
        'set default to incomplete SCR#49815
        meStatus = Incomplete
    End If
    If Not Dummy Then
        Set mobjLocking = New Locking
        mobjLocking.msFacilityNCID = FacilityNCID
        mobjLocking.msEncounterNumber = mnEncounterID
        mobjLocking.msUserNCID = mobjLogon.UserNCID
        ' cant lock since encounter must exists to create lock
        ' mobjLocking.LockSection 99, 0, SectionLock
         
        ' TODO: Need to rewrite this insert statement
        ' SCR #42987;   Developer:  01/06/2004
        ' Added Inpatient Service to Insert statement
        szSQL = "INSERT INTO ENCOUNTERS ( "
        szSQL = szSQL & " ENCOUNTERNUMBER,"
        szSQL = szSQL & " TYPE,"
        szSQL = szSQL & " CATEGORY,"
        szSQL = szSQL & " APPTID,"
        szSQL = szSQL & " STARTDTS,"
        szSQL = szSQL & " STARTDTS_GMT," '141844
        szSQL = szSQL & " FACILITYNCID,"
        szSQL = szSQL & " CLINICNCID,"
        szSQL = szSQL & " PRIMARYPROVIDERNCID,"
        szSQL = szSQL & " TIMEZONE,"
        szSQL = szSQL & " UNIT_NUMBER,"
        szSQL = szSQL & " APPTDTS,"
        szSQL = szSQL & " APPTDTS_GMT," '141844
        szSQL = szSQL & " STATUS,"
        szSQL = szSQL & " PATIENT_STATUS,"
                
        szSQL = szSQL & " INPATIENT_SERVICE," 'SCR-47912
        
        If Len(mobjShared.CmdLineSwitches("INPATIENT2")) > 0 Then
            'SF
            szSQL = szSQL & " INPT_NOTE_TYPE_NCID," 'SCR-56557
            szSQL = szSQL & " INPT_HOSP_DAY," 'SCR-56557
        End If
        
        szSQL = szSQL & szHIPPASqlFields
        
        szSQL = szSQL & " UPDATELOCK)"
        
        szSQL = szSQL & " VALUES ( " & vbCrLf
        szSQL = szSQL & " %ENCOUNTERNUMBER,"
        szSQL = szSQL & " %TYPE,"
        szSQL = szSQL & " '%CATEGORY',"
        szSQL = szSQL & " %APPTID,"
        szSQL = szSQL & " %STARTDTS,"
        szSQL = szSQL & " %STARTDTS_GMT," '141844
        szSQL = szSQL & " %FACILITYNCID,"
        szSQL = szSQL & " %CLINICNCID,"
        szSQL = szSQL & " %PRIMARYPROVIDERNCID,"
        szSQL = szSQL & " '%TIMEZONE',"
        szSQL = szSQL & " %UNIT_NUMBER,"
        szSQL = szSQL & " %APPTDTS,"
        szSQL = szSQL & " %APPTDTS_GMT," '141844
        szSQL = szSQL & " %STATUS,"
        szSQL = szSQL & " '%PATIENT_STATUS',"
        
        'szSQL = szSQL & " %CREATEDBY," 55608
        szSQL = szSQL & " %INPATIENT_SERVICE,"
        
        If Len(mobjShared.CmdLineSwitches("INPATIENT2")) > 0 Then
            'SF
            szSQL = szSQL & " '%INPT_NOTE_TYPE_NCID'," 'SCR-56557
            szSQL = szSQL & " %INPT_HOSP_DAY," 'SCR-56571
        End If
        
        szSQL = szSQL & szHIPAASqlValues
        
        szSQL = szSQL & " %UPDATELOCK)"
        
        szSQL = Replace(szSQL, "%ENCOUNTERNUMBER", mnEncounterID)
        szSQL = Replace(szSQL, "%TYPE", Enc_Type)
        szSQL = Replace(szSQL, "%CATEGORY", ApptType)
        szSQL = Replace(szSQL, "%APPTID", ApptID)
        szSQL = Replace(szSQL, "%STARTDTS_GMT", gobjCHCSIIConn.SQLDate(mobjComm.LocalToGmt(Now))) '141844
        szSQL = Replace(szSQL, "%STARTDTS", gobjCHCSIIConn.SQLDate(Now))
        szSQL = Replace(szSQL, "%FACILITYNCID", FacilityNCID)
        szSQL = Replace(szSQL, "%CLINICNCID", ClinicNCID)
        szSQL = Replace(szSQL, "%PRIMARYPROVIDERNCID", PrimaryProviderNCID)
        szSQL = Replace(szSQL, "%TIMEZONE", msTimeZone) 'SF SCR 69422 ' GetLocalTimeZone())
        szSQL = Replace(szSQL, "%UNIT_NUMBER", mobjPatient.UnitNumber)
        szSQL = Replace(szSQL, "%APPTDTS_GMT", gobjCHCSIIConn.SQLDate(mobjComm.LocalToGmt(ApptDTS))) '141844
        szSQL = Replace(szSQL, "%APPTDTS", gobjCHCSIIConn.SQLDate(ApptDTS))
        szSQL = Replace(szSQL, "%STATUS", StatusEnumToNCID(meStatus))
        szSQL = Replace(szSQL, "%PATIENT_STATUS", msPatientStatus)
        szSQL = Replace(szSQL, "%UPDATELOCK", CInt(False))
        
        szSQL = Replace(szSQL, "%INPATIENT_SERVICE", SQLQuote(Me.InpatientService))
        
        If Len(mobjShared.CmdLineSwitches("INPATIENT2")) > 0 Then
            'OA
            If InptNoteTypeNCID = "" Then
                szSQL = Replace(szSQL, "'%INPT_NOTE_TYPE_NCID'", "NULL")
            Else
                szSQL = Replace(szSQL, "%INPT_NOTE_TYPE_NCID", InptNoteTypeNCID)
            End If
            szSQL = Replace(szSQL, "%INPT_HOSP_DAY", InptHospDay)
            'Update the variable
            msInptNoteTypeNCID = InptNoteTypeNCID
            mlInptHospDay = InptHospDay
        End If
                
        oDas.ExecuteSQL szSQL
        
        'Error is not raised if insert fails 'SCR#45056,45066,45074
        If Len(CheckIfEncounterExists(ApptID, FacilityNCID, mobjPatient.UnitNumber)) = 0 Then
            'Unable to create encounter
            GEMS_Create = False
            Exit Function
        End If
        
        'Save the Provider Info
        If Len(mobjComm.CmdLineSwitch("HIPAA837")) > 0 Then
            If mobjShared.IsAppMode(modeITT) Then
               SetProviderInfo pcPRIMARY, objProvider.NCID, rtnAttendingProvider, objProvider.IEN, objProvider.FullName
               SaveHIPAAData bForceCreate:=True
            ElseIf mobjShared.IsAppMode(modeCHCSI_GUI) Then
               SetProviderInfo pcPRIMARY, objProvider.NCID, rtnAttendingProvider, objProvider.IEN, objProvider.FullName
               GEMS_Create = objProvider.GEMS_SaveEncounterProvider(mnEncounterID, FacilityNCID, 1)
            End If
        End If
        
    End If
    
    Dim oSQLEx As ICHCSII_SQL.ISqlOpsEx
    Dim sTemplate As String
        
    Set oSQLEx = gobjCHCSIIConn.CHCSII_SQLOPS(Auto) '''<SCR 36493

        'Add Default template functionality SCR-22258
    If mobjComm.CmdLineSwitch("GREEN") = "" Then
       'When the Green switch is removed remove the code in this block.
        Dim objTemplateMgmt As Object
        Set objTemplateMgmt = mobjComm.InitializeOLEServer(NCID_TEMPLATE_MANAGEMENT)
        If Not objTemplateMgmt Is Nothing Then
           sTemplate = objTemplateMgmt.DefaultEncounterTemplateData(mobjLogon.UserNCID)
        End If
        Set objTemplateMgmt = Nothing
        If Len(sTemplate) > 0 Then
           oSQLEx.Execute "Select TemplateData from Encounters " & vbCrLf & _
              "where ENCOUNTERNUMBER = " & CStr(mnEncounterID) & vbCrLf & _
              "AND FACILITYNCID = " & FacilityNCID
           If Not (oSQLEx.BOF And oSQLEx.EOF) Then
              oSQLEx.AppendAsChunk "TemplateData", sTemplate
              oSQLEx.Update
           End If
        End If
    End If
    
    If Not Dummy Then
        'need to create sections as well SCR-40076
        Call GEMS_CreateSections(mobjPatient.UnitNumber, _
                             FacilityNCID, _
                             mnEncounterID, _
                             oDas)
    End If
    
    If ClerkNote <> "" Then
        Me.AddTextNote "Appointment Note", ClerkNote
    End If
    mbLoading = False
    
    'If GEMS_Save() Then
        mbDirty = False
        mbFlagNew = False
        GEMS_Create = True
    ' Else
    '     MsgBxARMd "Creation of encounter failed."
    ' End If
    If Not Dummy Then
        mobjLocking.UnlockSection 99, 0
    Else
        'So Prev Enc wont try to load the dummy encounters
        mbFullyLoaded = True
    End If

'    Set rs = Nothing
    Set mobjConcept = Nothing
    Set objProvider = Nothing
    Set oDas = Nothing
    
    Exit Function
    
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "Encounter.GEMS_Create", App.Title, vbCritical

    'If rs.State > 0 Then rs.Close
    'Set rs = Nothing
    Set mobjConcept = Nothing
    Set objProvider = Nothing
End Function


Private Sub CreateSections(Optional ByRef colDataStrings As Collection = Nothing)
    Dim colData As Collection
    If colDataStrings Is Nothing Then
        Set colData = New Collection
    Else
        Set colData = colDataStrings
    End If
    
    If mobjShared.IsAppMode(modeCHCSI_GUI) And Not (mobjWAM Is Nothing) Then
      colData.Add mobjWAM.Encounter.RTFs.Item(1), CStr(eap)
      colData.Add mobjWAM.Encounter.RTFs.Item(2), CStr(eDisp)
      colData.Add mobjWAM.Encounter.AnPStatus, "AnPStatus"
      colData.Add mobjWAM.Encounter.DispositionStatus, "DispositionStatus"
    End If
    
    Call LoadApptNotes(colData)
    CreateSection eAutoCite, colData
    CreateSection eRFV, colData
    CreateSection eVitals, colData
    CreateSection eSO, colData
    CreateSection eLabs, colData
    CreateSection eRads, colData
    CreateSection eQuestionnaire, colData
    CreateSection eEducation, colData
    CreateSection eap, colData
    If EncounterType = Dental Then
       CreateSection eDental, colData
    End If
    CreateSection eDisp, colData
    CreateSection eNotes, colData
    CreateSection eHistory, colData
    Set colData = Nothing
End Sub

Private Sub CreateSection(ByVal eSection As EncounterSectionEnum, ByRef colDataStrings As Collection)
    Dim objSecParent As SectionParent
    On Error Resume Next
    Set objSecParent = mcolSecParents(CStr(eSection))
    If objSecParent Is Nothing Then
        Set objSecParent = New SectionParent
        'If mobjShared.AppMode <> modeCHCSI_GUI Then
            Call objSecParent.Load(Me, eSection, colDataStrings)
        'End If
        mcolSecParents.Add objSecParent, CStr(eSection)
    Else
        ' we get here from Refresh->Load->CreateSections->CreateSection
        'If mobjShared.AppMode <> modeCHCSI_GUI Then
            Call objSecParent.Load(Me, eSection, colDataStrings)
        'End If
    End If
    'TODO - Atecnio Need to find out where error coming from (causing issues with Coding Review)
    Err.Clear
End Sub

'*****************************************************************************
'History
' Developer     Modified Date   Description
'-----------    --------------  -------------------------------------------
' Wilson Sun    05/22/01        Added bPreview in PrintSF603A2
' Wilson Sun    05/14/01        Change PrintSF603A to PrintSF603A2 and build record set
' Wilson Sun    03/08/01        Change PrintSF603A method parameters
'*****************************************************************************
Public Function PrintEncounter(bPreview As Boolean, _
                                                Optional bPOC_Continuation As Boolean = False, _
                                                Optional bIncludeSignature As Boolean = False, _
                                                Optional bRefreshData As Boolean = True) As Boolean
    Dim objForms As Object
    Dim sSql As String
    On Error Resume Next
    
    ''''''''''''''''''''
    ' SCR #11246.      '
    ' Printing errors. '
    ''''''''''''''''''''
    Dim sCurrentText As String
    Screen.MousePointer = vbHourglass
    
    'Refresh to get all the latest data for the encounter SCR#47397
    If bRefreshData Then 'SCR#48882
        Refresh
    End If
    
    'For cancelled encounters used RTF saved in CDR since logon user will be documented as cancelling provider in BuildRTF SCR-39045
    If (meStatus = CancelByFacility Or meStatus = CancelByPatient Or meStatus = NoShow Or meStatus = PatientLeftWithoutBeingSeen Or meStatus = CancelByProvider) And mcolEncRTFs.Count > 0 Then
        sCurrentText = mcolEncRTFs(mcolEncRTFs.Count).msRTF
    Else
        sCurrentText = BuildRTF(Me, bPOC_Continuation, bIncludeSignature) 'SCR-33863 3/20/03/ RBELL Rebuild RTF whenever printing
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    ' SCR #11246. Printing Errrors                    '
    ' Setting PrintEncounter = False will trigger a   '
    ' msg indicating encounter document not found.    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    If sCurrentText = "" Then
        PrintEncounter = False
        Exit Function
    End If
    
    'Wilson1 Begin
    'KDunne 6/14/2002
    'SCR 8900 - Only people with dental access print the SF603A, everyone else prints the SF600.

    If Me.EncounterType = Dental And mobjUser.HasPrivilegeEx(Priv_Dental_Access, Priv_Read) Then
        Set objForms = CreateObject("CHCSII_Forms.DentalForms")
    Else
        Set objForms = CreateObject("CHCSII_Forms.Forms")
    End If
    
    Set objForms.Comm = mobjComm
    
    'Wilson1 end
    If objForms Is Nothing Then
        MsgBxARMd "Could not create CHCSII_Forms object.", vbExclamation, "Encounter.Print"
        Exit Function
    End If
    'Wilson2 begin
            
    'SCR 45291 RWalker 11/25/03
    goARM.StartCmplxTx gsARM_Tx_CorrUT_PrintAndPreview, , , , , , , , eAnyFamily
    
    'KDunne 6/14/2002
    'SCR 8900 - Only people with dental access print the SF603A, everyone else prints the SF600.
    If Me.EncounterType = Dental And mobjUser.HasPrivilegeEx(Priv_Dental_Access, Priv_Read) Then
        If bPreview Then
            mobjComm.SetStatusText "Building SF603A Form"
        Else
            mobjComm.SetStatusText "Printing SF603A Form"
        End If
        
        If objForms.PrintSF603A2(Me.Patient, Me, Nothing, Nothing, Nothing, bPreview) Then
            PrintEncounter = True
        ElseIf Not bPreview Then
            'KDunne 6/20/2002
            'SCR 25020 - PrintSF603A already displays approriate message. We would be
            '            displaying two messages for the same error.
            '
'            MsgBxARMd "An error occurred during printing.", vbCritical, "Print Error"
            PrintEncounter = False
        End If
    Else 'Wilson2 end
        Screen.MousePointer = vbDefault

        If objForms.PrintSF600(Me, sCurrentText, bPreview) Then
            PrintEncounter = True
        ElseIf Not bPreview Then
            'MsgBxARMd "An error occurred during printing.", vbCritical, "Print Error"
            PrintEncounter = False
        End If
    End If
    mobjComm.SetStatusText ""
    Set objForms = Nothing
            
    '--- SCR-ARM-45291 RWalker 11/25/2003
    goARM.StopCmplxTx gsARM_Tx_CorrUT_PrintAndPreview, enmARMTxStatusOk, enmARMStop_MeOnly

    Screen.MousePointer = vbDefault


   
End Function


Public Function PrintPlanOfCare(bPreview As Boolean) As Boolean

    Dim objForms As Object

On Error Resume Next
    If mobjComm.CmdLineSwitch("CPG") = vbNullString Then Exit Function

    Set objForms = CreateObject("CHCSII_Forms.Forms")
    Set objForms.Comm = mobjComm
    If objForms Is Nothing Then
       MsgBxARMd "Could not create CHCSII_Forms object.", vbExclamation, "Encounter.PrintPlanOfCare"
       Exit Function
    End If
    If objForms.printCPG_POC(1, mobjPatient.UnitNumber) Then
       PrintPlanOfCare = True
    End If
    mobjComm.SetStatusText vbNullString
    Set objForms = Nothing

End Function

Public Function Save(Optional ByVal bForce As Boolean = False) As Boolean
    Dim bResult          As Boolean
    Dim i                As Integer
    Dim objData          As DataString
    Dim objEvent         As ClinicalEvent
    Dim objHeader        As ClinObsHeader
    Dim objClinObs       As ClinicalObservation
    Dim objSectionParent As SectionParent
    Dim oEmh             As EmhCtrl
    Dim objProvider      As EncProvider
    Dim objEncRTF        As EncRTF
    Dim bRaiseEvent      As Boolean
    Dim lDocumentNumber  As Long
    
    On Error GoTo ErrHandler
    
    If Not (mbDirty Or mbFlagNew Or bForce Or AccidentData.IsDirty Or PregnancyData.IsDirty Or AssociatedProviders.IsDirty) Then
      Save = True
      Exit Function
    End If
    
    'If UpdateLock is set and wasn't just set, then don't perform the save
    If mblnUpdateLock And Not mblnUpdateLockPending Then
        Save = True
        Exit Function
    End If
    
    #If debugon Then
        Dim DR As DebugRec
        If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Save", "Begin", DR, "Enc", mnEncounterID
    #End If
    

    If Not mobjLocking.StartSectionUpdate(eHeader, 0) Then
        MsgBxARMd "Your lock on this encounter was broken by another user, and your changes cannot be saved.", vbInformation, "Encounter"
        '--- SCR-ARM-25468 RWalker 4/10/03 Must manually stop a correlated tx
        goARM.StopCmplxTx gsARM_Tx_CorrUT_SignFamily, enmARMTxStatusFailed, enmARMStop_MeOnly
        GoTo CleanUp
    End If

    If meStatus = complete Or meStatus = NeedsCoSignature Or meStatus = Updated Then 'SCR-34229
        'If there is no E&M Code then SADR Status has to be NO_COUNT SCR-26152
        If Len(SectionParent(eDisp).mobjSection.EandMNCID) = 0 And (msSADRStatus = NCID_SADR_NEEDS_NEW_PROCESSING Or msSADRStatus = NCID_SADR_NEEDS_UPDATE_PROCESSING) Then
            msSADRStatus = NCID_SADR_NO_COUNT
        End If
        'Also
        'If there is no Primary Dx then SADR Status has to be NO_COUNT SCR-26152
        If mnPrimaryDxSnoID = "" And (msSADRStatus = NCID_SADR_NEEDS_NEW_PROCESSING Or msSADRStatus = NCID_SADR_NEEDS_UPDATE_PROCESSING) Then
            msSADRStatus = NCID_SADR_NO_COUNT
        End If
    End If

    'Reset msRTF
    msRTF = ""
    
    Save = GEMS_Save(bForce)
    Exit Function

ErrHandler:
   Call mobjShared.ShowVBError(Err.Number, Err.Description, "Enctr.Save ", "Enc", vbCritical)

CleanUp:
    Exit Function
    Resume
    
End Function

Friend Function InitializeDataString(ByVal eSection As EncounterSectionEnum, _
                                     ByVal nDataID As Long, _
                                     ByRef dStartDTS As Date, _
                                     ByRef dEndDTS As Date) As DATASTRATLLib.DataString
    Dim objDAS As DASATLLib.DasCtrl
    Dim objData As DATASTRATLLib.DataString
    Dim objEvent As CLINTYPEATLLib.ClinicalEvent
    Dim objHeader As CLINTYPEATLLib.ClinObsHeader
    Dim sDTS                   As String
    Dim i                      As Integer
    On Error GoTo ErrHandler
    If nDataID = 0 Then
        Set objData = New DATASTRATLLib.DataString
        objData.ClearAll
        objData.DataTypeNCID = NCID_ENCEVENT
        If mobjPatient.UnitNumber <> 0 Then
            objData.UnitNumber = mobjPatient.UnitNumber  'mnPatientID SCR-12436 make sure unit number is correct
        Else
            objData.UnitNumber = mnPatientID  'If Co-Signing patient obj not used SCR-13658
        End If
        If Val(msClinicNCID) > 0 Then
          objData.PointOfCare.PhysicalLocation.InfoType = "nursingDivision"
          objData.PointOfCare.PhysicalLocation.NursingDivision.ValueConcept = msClinicNCID
        End If
    Else
        '- Just like the new one
        Set objData = New DATASTRATLLib.DataString
        objData.ClearAll
        objData.DataTypeNCID = NCID_ENCEVENT
        objData.UnitNumber = mobjPatient.UnitNumber  'mnPatientID SCR-12436 make sure unit number is correct
        If Val(msClinicNCID) > 0 Then
          objData.PointOfCare.PhysicalLocation.InfoType = "nursingDivision"
          objData.PointOfCare.PhysicalLocation.NursingDivision.ValueConcept = msClinicNCID
        End If
    
    End If
    objData.Application.ValueConcept = NCID_ENCOUNTER_APPLICATION
    objData.Facility.ValueConcept = msFacilityNCID
    objData.EncounterNumber = mnEncounterID
    objData.EncounterFacility.ValueConcept = msFacilityNCID
    objData.Clinician.ValueConcept.UniqueId = mobjLogon.UserNCID
    objData.SemanticLinks.ClearAll
    Set objEvent = objData.PatientData
    objEvent.ClinObs.ClearAll
    Set objHeader = objEvent.Header
    With objHeader
        .ObsBatId.InfoType = "att"
        .ObsBatId.value.InfoType = "coded"
        .ObsBatId.value.Coded.ValueConcept.UniqueId = SectionEnumToNCID(eSection)
        '-----------------------------------------
        ' Date/Time of session to .ObsDateTime
        '-----------------------------------------
        '-- can't store a zero here
        If dStartDTS > 0 Then
            sDTS = dStartDTS
        Else
            sDTS = CDate(0)
        End If
        .ObsDateTime.InfoType = "timeIntAtt"
        .ObsDateTime.StartTime.InfoType = "att"
        .ObsDateTime.StartTime.value.InfoType = "dateTime"
        .ObsDateTime.StartTime.value.DateTime = sDTS
        '-- can't store a zero here
        If dEndDTS > 0 Then
            sDTS = dEndDTS
        Else
            sDTS = CDate(0)
        End If
        .ObsDateTime.EndTime.InfoType = "att"
        .ObsDateTime.EndTime.value.InfoType = "dateTime"
        .ObsDateTime.EndTime.value.DateTime = sDTS
        .Comments.ClearAll
        .ObservedBy.InfoType = "att"
        .ObservedBy.value.InfoType = "codedWOSform"
        .ObservedBy.value.CodedWOSform.ValueConcept.UniqueId = mobjLogon.UserNCID
    End With
    Set InitializeDataString = objData
    GoTo CleanUp
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "Encounter.InitializeDataString()", Exe, vbCritical
CleanUp:
    Set objData = Nothing
    Set objEvent = Nothing
    Set objHeader = Nothing
End Function


'Refactor Phase 1 SCR#48785
Public Function SignEncounter(Optional AppToBringToForeground As String = "") As Boolean
    Dim objSignOps As EncSignOps
    Dim objWAMErr As ErrObject
    Dim objErr As WAMData.ErrObject
    Dim objWAMEncounter As CHCS_Interface.ICHCSEncounters
    Dim sTemp As String
    Dim sMsg As String
    Dim sCoSignerNCID As String
    
    On Error GoTo ErrHandler
    
    If Len(mobjComm.CmdLineSwitch("INPATIENT2")) > 0 Then
        'RQT 56576                                'SF 4-11-05 SCR 68832
        If Not Me.Cosigner Is Nothing Then
            sCoSignerNCID = Me.Cosigner.NCID
        End If
        If Me.EncounterType = InpatientNote And Not (Me.CreateClinicianNCID = Logon.UserNCID Or sCoSignerNCID = Logon.UserNCID) Then
            MsgBxARMd "Only the provider that created the note can sign it.", vbInformation, "EncounterOps"
            goARM.StopCmplxTx gsARM_Tx_CorrUT_SignFamily, enmARMTxStatusFailed, enmARMStop_MeOnly
            Exit Function
        End If
    End If
    
    If Not mbOpen Then  'bmstop
        MsgBxARMd "Cannot sign an encounter that is not opened.", vbInformation, "EncounterOps"
        '--- SCR-ARM-25468 RWalker 4/10/03 Must manually stop a correlated tx
        goARM.StopCmplxTx gsARM_Tx_CorrUT_SignFamily, enmARMTxStatusFailed, enmARMStop_MeOnly
        Exit Function
    End If
    
    sMsg = "Compare Patient"
    If CStr(mobjPatient.UnitNumber) <> mnPatientID Then
        mobjComm.Message cwiSELECT_PATIENT, mnPatientID & ";", NCID_CW_APPLICATION, 0
    End If
    
    sMsg = "Lock Entire Encounter"
    If Not mobjLocking.LockSection(eEntireEncounter, 0, SignLock) Then
        '--- SCR-ARM-25468 RWalker 4/10/03 Must manually stop a correlated tx
        goARM.StopCmplxTx gsARM_Tx_CorrUT_SignFamily, enmARMTxStatusFailed, enmARMStop_MeOnly
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass  'SCR#45993
    
    sMsg = "Create EncSignOps"
    Set objSignOps = New EncSignOps
        
    sMsg = "Sign Encounter"
    SignEncounter = objSignOps.SignEncounter(Me)
    
    '--- SCR-ARM-25468 RWalker 4/10/03 Set ARM Status based upon Sign Encounter status
    If SignEncounter Then
        goARM.StopCmplxTx gsARM_Tx_CorrUT_SignFamily, enmARMTxStatusOk, enmARMStop_MeOnly
    Else
        goARM.StopCmplxTx gsARM_Tx_CorrUT_SignFamily, enmARMTxStatusFailed, enmARMStop_MeOnly
    End If

    If SignEncounter And Not mobjParent Is Nothing Then
        'Unlock the encounter because it was signed (reverse SCR 64240).
        Call mobjLocking.UnlockSection(eEntireEncounter, 0)
        
        'SF 3-10-05 allow the calling app to override the default logic in DefaultParentApplication
        If AppToBringToForeground = Empty Then
            'DefaultParentApplication will return correct app to return to after signing encounter SCR#47795
            AppToBringToForeground = DefaultParentApplication
        End If
        
        'If server is already running, bring it to the foreground SCR-26464, 26453
        If Not mobjComm.Message(cwiOLE_SERVER_RUNNING, AppToBringToForeground, NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION) Then
            Call mobjComm.Message(cwiSTART_OLE_SERVER, AppToBringToForeground, NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION)
        Else
            Call mobjComm.Message(cwiBRING_SERVER_TO_FOREGROUND, AppToBringToForeground, NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION)
        End If
        
        If Me.EncounterType = TeleConsult Then 'SCR#48729
            Call mobjComm.Message(cwiREFRESH_DATA, "TELCONID|" & AppointmentID, NCID_TELCON_APPLICATION, NCID_ENCOUNTER_APPLICATION)
        ElseIf Me.EncounterType <> InpatientNote Then 'SCR 67667
            'SCR#49319
            Call mobjComm.Message(cwiREFRESH_DATA, "APPTID|" & AppointmentID, NCID_CLINIC_SCHEDULE_APPLICATION, NCID_ENCOUNTER_APPLICATION)
        End If
        
    Else
        Call Me.RefreshDisplay
    End If
                                                                                                      
    If Len(mobjComm.CmdLineSwitch("INPATIENT2")) > 0 Then
        If SignEncounter And Me.EncounterType <> InpatientNote Then
            'Admit an outpatient
            'Determine if the disposition is "Admitted"
            If Me.DispositionNCID = NCID_Admitted Then
                AdmitPatient
            Else
                'Close the encounter
                'SCR 64991
                'SCR 67747 'Don't call close current encounter if the parent is enccosign
                'This is a kludge but it is late bound so use the type name instead of the object type.
                If TypeName(mobjParent) <> "EncCosign" Then
                    mobjParent.CloseCurrentEncounter
                End If
            End If
        ElseIf SignEncounter Then
            'Close the encounter
            'SCR 64991
            'SCR 67747 'Don't call close current encounter if the parent is enccosign
            'This is a kludge but it is late bound so use the type name instead of the object type.
            If TypeName(mobjParent) <> "EncCosign" Then
                mobjParent.CloseCurrentEncounter
            End If
        End If
    ElseIf SignEncounter Then
        mobjParent.CloseCurrentEncounter
    Else
        'Do nothing because the user did not sign the encounter.
    End If
    
    Screen.MousePointer = vbDefault
    Set objSignOps = Nothing

Exit Function

ErrHandler:
    If InStr(Err.Description, "Cannot obtain a lock on the ADM record") Then
        MsgBxARMd "The appointment record for this appointment is open in CHCS. Please close the open record in CHCS before attempting to File.", vbExclamation, App.Title
        '--- SCR-ARM-25468 RWalker 4/10/03 Must manually stop a correlated tx
        goARM.StopCmplxTx gsARM_Tx_CorrUT_SignFamily, enmARMTxStatusFailed, enmARMStop_MeOnly
        Exit Function
    End If
    
    Call mobjShared.ShowVBError(Err.Number, Err.Description & " - " & sMsg, "Encounter.SignEncounter", Exe, vbCritical)
    
    Exit Function
    Resume
End Function
Private Function AdmitPatient() As Boolean

    Dim oPAD As Object
    Dim oInpatient As Object
    Dim oDocuments As Object
'    Dim objEnc As CHCSIIEncounterOps.Encounter
    Dim sErrMsg As String
    Dim sEncounterNumberFrom As String
    Dim lMousePointer As Long
    Dim objEncParent As Object
    Dim lInpatientID As Long
    
    On Error GoTo ErrHandler
    
    lMousePointer = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    
    sEncounterNumberFrom = Me.EncounterID
    
'    LogAction "AdmitPatient - Start"

'    Set objEnc = gobjEncParent.CurrentEncounter
    mobjComm.Message cwiSTART_OLE_SERVER, NCID_INPATIENT_APPLICATION, NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION
    Set oPAD = mobjComm.InitializeOLEServer(NCID_INPATIENT_APPLICATION)
    If Not (oPAD Is Nothing) Then 'SCR#48990
        If oPAD.AdmitPatient(mobjPatient, Logon.UserNCID, Me.PrimaryProvider.NCID, Now, Me.EncounterID, sErrMsg, lInpatientID, sEncounterNumberFrom) = False Then
            If Len(sErrMsg) Then
                'Display all other err msgs 'SCR#48605
                'If StrComp(sErrMsg, "Patient is already Admitted", vbTextCompare) <> 0 Then
                    MsgBxARMd sErrMsg, vbInformation, "Encounter" '48248
                'End If
            End If
           Set objEncParent = mobjComm.InitializeOLEServer(NCID_ENCOUNTER_APPLICATION)
           objEncParent.CloseCurrentEncounter
           'CloseCurrentEncounter
        Else
            'prompt for completing admission...
           Set objEncParent = mobjComm.InitializeOLEServer(NCID_ENCOUNTER_APPLICATION)
           
           objEncParent.CloseCurrentEncounter
           'CloseCurrentEncounter
            If MsgBxARMd("Do you want to complete an Admission History and Physical now?", vbYesNo + vbQuestion) = vbYes Then
                'Create the note
                Set oInpatient = oPAD.SelectedPatients.getInPatientByID(lInpatientID)
                'SCR 105943
                If oInpatient Is Nothing Then
                    Set oInpatient = oPAD.SelectedPatients.GetInpatientFromDB(CStr(lInpatientID), FacilityNCID)
                End If
                'Select that patient
                Set oPAD.SelectedPatients.Inpatient = oInpatient
                
                'Add the note
                oPAD.AddNote oInpatient, NCID_Admission_HP_Note
                'Lock the encounter, save it, unlock the encounter
                'Don't call this instance of the encounter because this one is not the current one.
                'Get the current one from encounter parent.
                objEncParent.CurrentEncounter.LockingObject.LockSection eEntireEncounter, 0, SectionLock
                objEncParent.CurrentEncounter.Save
                objEncParent.CurrentEncounter.LockingObject.UnlockSection eEntireEncounter, 0
'                Me.LockingObject.LockSection eEntireEncounter, 0, SectionLock
'                Me.Save
'                Me.LockingObject.UnlockSection eEntireEncounter, 0
                DoEvents
                
            End If
        End If
    Else
        'SCR 64991
        objEncParent.CloseCurrentEncounter
    End If
    
    'LogAction "AdmitPatient - End"
    
    Set objEncParent = Nothing
    Set oPAD = Nothing
'    Set objEnc = Nothing
    
    Screen.MousePointer = lMousePointer

    Exit Function
    
ErrHandler:
    MsgBox Err.Description & " in EncounterParent.AdmitPatient"
    Exit Function
    Resume
    
End Function




'Called by SignEncounter to return correct application ncid to return to after signing an encounter SCR#47795
Private Function DefaultParentApplication() As String
    
    If Me.EncounterType = TeleConsult Then
        DefaultParentApplication = NCID_TELCON_APPLICATION
    ElseIf Me.EncounterType = InpatientNote Then
        DefaultParentApplication = NCID_INPATIENT_APPLICATION
    Else
        DefaultParentApplication = NCID_CLINIC_SCHEDULE_APPLICATION
    End If
    
End Function

Public Function AddHistoryItem(ByVal Reason As ChangeHistoryEnum, ByRef ItemName As String, ByRef RTF As String, _
                                Optional ByVal OldOwnerName As String = "", Optional ByVal OldDTS As Date, Optional ByVal SectionName As String = "") As Boolean
    Dim objSecParent As SectionParent
    Dim objHistory As History
    Dim bDidOpen As Boolean
    
    On Error GoTo ErrHandler
    
    Set objSecParent = mcolSecParents(CStr(eHistory))
    Set objHistory = objSecParent.mobjSection
    If Not mbOpen Then
        bDidOpen = True
        Call objSecParent.OpenEncounter(Me)
    End If
    
    If Reason = Overwritten And Len(OldOwnerName) Then  'SCR-39592
        'Add overwrite message to RTF
        'Removed space after last \par SCR#39592
        RTF = "\pard\plain\f0\fs16" & _
                "The " & SectionName & " section was last updated by " & Logon.UserName & " @ " & FormatDTG(Now) & " - see above." & _
                "Previous Version of " & SectionName & " section was entered/updated by " & OldOwnerName & " @ " & FormatDTG(OldDTS) & "." & _
                " \par" & RTF
    End If
    
    AddHistoryItem = objHistory.AddHistoryItem(Reason, ItemName, RTF)
    If bDidOpen Then
        Call objSecParent.CloseEncounter
    End If
    Set objSecParent = Nothing
    Set objHistory = Nothing
    
    Exit Function
    
ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description, "Encounter.AddHistoryItem ", Exe, vbCritical)
    
End Function

Public Sub AddProvider()

Dim bReturn As Boolean
Dim bDidLock As Boolean
    
'     'SCR #11001 Unauthorized user shouldn't add provider.
'     If Not User.HasPrivilegeEx(Priv_Current_Encounter, Priv_Read) Then
'        MsgBxARMd "You do not have permission to view this information." & Chr$(10) _
'          & "Please contact your security administrator if you believe you have received this message in error.", vbExclamation
'        Exit Sub
'     End If
     
    If Not mobjLocking.SectionIsLocked(eHeader, 0) Then
        If Not mobjLocking.LockSection(eHeader, 0, SectionLock, "Providers") Then
            Exit Sub
        End If
        bDidLock = True
    End If
    
    If (mobjShared.IsAppMode(modeCDR) Or mobjShared.IsAppMode(modeCHCSI_GUI)) And Len(mobjComm.CmdLineSwitch("HIPAA837")) > 0 Then
        bReturn = frmHIPAAProviders.Display(Me)
    Else
        frmProviderList.Display Me
        bReturn = True
    End If
    
    If bDidLock Then
        Call mobjLocking.UnlockSection(eHeader, 0)
    End If
    
    If bReturn Then Call RefreshDisplay
    
End Sub
' SCR-22188 Added ReadOnly optional parameter
Public Function AddTextNote(ByRef Title As String, ByRef Text As String, Optional ByRef DataId As Long, Optional ByRef bReadOnly As Boolean = False, Optional ByVal eRelatedSection As EncounterSectionEnum = eNotes) As Boolean
    Dim objSecParent As SectionParent
    Dim objTextNotes As TextNotes
    Dim objNote As TextNote
    Dim bDidOpen As Boolean
    Dim lLockNumber As Long
    
    On Error Resume Next
    
    If mcolSecParents Is Nothing Then
        Refresh 'refresh to reload sections 'SCR#47049
    End If
    
    Set objSecParent = mcolSecParents(CStr(eNotes))
    
    If objSecParent Is Nothing Then
        Refresh 'refresh to reload sections 'SCR#47049
        Set objSecParent = mcolSecParents(CStr(eNotes))
        If objSecParent Is Nothing Then
            MsgBxARMd "Unable to retrieve section data to Add Text Note for Encounter:" & mnEncounterID, vbCritical, "Encounter - Add Text Note"
            Exit Function
        End If
    End If
    
    Set objTextNotes = objSecParent.mobjSection
    If Not mbOpen Then
        bDidOpen = True
        Call objSecParent.OpenEncounter(Me)
    End If
    Set objNote = New TextNote
    objNote.mbIsComplete = True
    objNote.mdDTS = Now
'<< shaw not the correct owner
    'Add name in all cases except when PGUI mode and adding to SO section SCR-38241
    If Not (mobjShared.IsAppMode(modeCHCSI_GUI) And eRelatedSection = eSO) Then '<SCR SCR 36493
        objNote.msOwnerName = mobjLogon.UserName
    End If
    objNote.msOwnerNCID = mobjLogon.UserNCID
    objNote.msRTF = ConvertTextToRTF(Text)
    objNote.msTitle = Title
    objNote.mnDataID = DataId
    objNote.ReadOnly = bReadOnly
    objNote.RelatedSection = eRelatedSection 'SCR-38142
'<<< shaw need this save feature to add to the collection in chcs_gui mode
'    If Not (mobjShared.IsAppMode(modeCHCSI_GUI) And eRelatedSection = eNotes) Then
        'Need to lock section SCR-38102
        If mobjLocking.LockSection(eNotes, objNote.DataId, SectionLock) Then
            'Record # since it may change from 0 to ###### SCR-35669
            lLockNumber = objNote.DataId
            AddTextNote = objTextNotes.Save(objNote)
            mobjLocking.UnlockSection eNotes, lLockNumber
        End If
'    ElseIf mobjShared.IsAppMode(modeCHCSI_GUI) Then
'        'Add non SO notes to AppendNote in PGUI mode SCR-38142
'        WAM.Encounter.AppendNote objNote.RTF
'    End If
    
    If AddTextNote And Not IsMissing(DataId) Then
        DataId = objNote.mnDataID
    End If
    Set objNote = Nothing
    If bDidOpen Then
        Call objSecParent.CloseEncounter
    Else
        Call RefreshDisplay
    End If
    Set objSecParent = Nothing
    Set objTextNotes = Nothing
End Function

Public Function DeleteTextNote(ByVal Index As Integer) As Boolean
    DeleteTextNote = mcolSecParents(CStr(eNotes)).mobjSection.DeleteTextNote(Index)
End Function

Public Sub EditNewTextNote(Optional ByVal eRelatedSection As EncounterSectionEnum = eNotes)
    Dim objNote As TextNote
    On Error Resume Next
    Set objNote = New TextNote
    objNote.msOwnerNCID = mobjLogon.UserNCID
    objNote.msOwnerName = mobjLogon.UserName
    objNote.mdDTS = Now
    objNote.RelatedSection = eRelatedSection
    Call EditTextNote(objNote)
    Set objNote = Nothing
End Sub

'Use mobjLogon.UserNCID as EventID since multiple notes may be open SCR#46335
Public Sub EditTextNote(ByRef Note As TextNote)
    Dim objNewNote As TextNote
    
    If mobjShared.IsAppMode(modeCHCSI_GUI) Then GoTo DoEdit
        
    If Note.mnDataID = 0 Then
        'Add lock for text note
        If Not mobjLocking.LockSection(eNotes, mobjLogon.UserNCID, SectionLock) Then 'SCR-35244
            Exit Sub
        End If
        Note.LockID = mobjLogon.UserNCID 'SCR#47028
        GoTo DoEdit
    End If
    If Note.msOwnerNCID = mobjLogon.UserNCID Then
        If Not mobjLocking.LockSection(eNotes, Note.DataId, SectionLock) Then
            Exit Sub
        End If
        Note.LockID = Note.DataId
        '>>> SCR-18260 S.McAvoy 7/09/02; When called from the TelconNoteFunctionality.cls - cmdSaveClick
        'and the Note title is "Appointment Note" or "Provider Telcon Note"
        'the only thing needed is to just
        'update the text note…  you don't want to open up the frmTextNote.
        If (Note.msTitle = "Appointment Note" Or Note.msTitle = "Provider Telecon Note") And meStatus = Updating Then
            Note.msOwnerNCID = mobjLogon.UserNCID
            Note.msOwnerName = mobjLogon.UserName
            Note.mdDTS = Now
            If Not mcolSecParents(CStr(eNotes)).mobjSection.Save(Note) Then
                GoTo CleanUp
            End If
            GoTo CleanUp
        End If
        '<<< SCR-18260
        If meStatus = Updating Then
            GoTo SaveToHistory
        End If
        GoTo DoEdit
    End If
    
    Select Case MsgBxARMd("This note belongs to " & Note.msOwnerName & "." & vbCrLf & vbCrLf _
                    & "Do you want to create your own note starting with the contents of this one?", _
                    vbYesNoCancel, "Encounter Note Ownership")
    Case vbYes
        Set objNewNote = New TextNote
        With objNewNote
            .msCategory = Note.msCategory
            .msTitle = Note.msTitle
            .msRTF = Note.msRTF
            .msOwnerNCID = mobjLogon.UserNCID
            .msOwnerName = mobjLogon.UserName
            .mdDTS = Now
        End With
        Set Note = objNewNote
        Set objNewNote = Nothing
        
        If Not mobjLocking.LockSection(eNotes, mobjLogon.UserNCID, SectionLock) Then 'SCR-35244
            Exit Sub
        End If
        Note.LockID = mobjLogon.UserNCID  'SCR#47028
        GoTo DoEdit
    Case vbNo
        Select Case MsgBxARMd("Do you want to copy the existing note to Change History and take over ownership of it?", _
                        vbYesNo, "Encounter Note Ownership")
        Case vbNo
            Exit Sub
        Case vbYes
            'Lock using DataID SCR-35317
            If Not mobjLocking.LockSection(eNotes, Note.DataId, SectionLock) Then
                Exit Sub
            End If
            Note.LockID = Note.DataId
            GoTo SaveToHistory
        End Select
    Case vbCancel
        Exit Sub
    End Select
MsgBxARMd "Internal error in Encounter.EditTextNote: unexpected execution flow."
Exit Sub
SaveToHistory:
    If Not AddHistoryItem(Overwritten, "Text Note", Note.RTF) Then
        MsgBxARMd "Unable to copy section to change history.", vbCritical, "Encounter"
        GoTo CleanUp
    End If
    Note.msOwnerNCID = mobjLogon.UserNCID
    Note.msOwnerName = mobjLogon.UserName
    Note.mdDTS = Now
    If Not mcolSecParents(CStr(eNotes)).mobjSection.Save(Note) Then
        GoTo CleanUp
    End If
DoEdit:
    'Added config and logon to initialize SCR - 12323
    Call frmTextNote.Initialize(Me, Note, Not Note.mbIsComplete, , mobjParent.Config, mobjLogon)
    mobjComm.Message cwiSHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION
    frmTextNote.Show vbModal
    mobjComm.Message cwiDONE_SHOWING_MODAL_FORM, "", NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION
CleanUp:
    'Unlock always since user may cancel
    Call mobjLocking.UnlockSection(eNotes, Note.LockID)
End Sub
Public Function LeftWithoutBeingSeen() As Boolean
    MsgBxARMd "DLL configuration error: Encounter.LeftWithoutBeingSeen() was replaced by Encounter.Cancel()."
    LeftWithoutBeingSeen = False
End Function

Public Function Cancel(ByVal NewStatus As EncStatus) As Boolean
'Rules: if anything has been done in A&P then return FALSE
'else, Set Disposition status to;;;;;
' and  set enc SADRStatus to needs coding
' and  Set Enc status to NewStatus
' and build rtf w/o signature block (but with user who did this)
    On Error GoTo ErrHandler
    Dim objSecParent As SectionParent
    Dim objEncRTF As EncRTF
    Dim bDidOpen As Boolean
'  COMMENTED OUT Code SCR-31783
'SCR 19750 - Cannot Cancel Appointments - Walker - 1/8/2002
'   If msPrimaryDiagnosisNCID <> "" AND Then
'    If msPrimaryDiagnosisNCID <> "" And msPrimaryDiagnosisNCID <> "0" Then
'        MsgBxARMd "This encounter cannot be cancelled because a diagnosis has been assigned."
'        GoTo CleanUp
'    End If
    'Set status before opening encounter SCR#52837
    
    Status = NewStatus
    If Not mbOpen Then
        Call Me.OpenEncounter(Nothing)
        bDidOpen = True
    End If

'<< SCR 10086
    If Not meClass = ec_NoCount Then
        SADRStatus = NCID_SADR_NEEDS_NEW_PROCESSING
    End If
'>>
    Set objSecParent = mcolSecParents(CStr(EncounterSectionEnum.eDisp))
    With objSecParent.mobjSection
        .DispositionNCID = 546
        .DispositionText = ""
        .DispAdminOption = ""
    End With
    Set objEncRTF = New EncRTF
    objEncRTF.msRTF = BuildRTF(Me)
    If objEncRTF.msRTF = "" Then
        GoTo CleanUp
    End If
    objEncRTF.mdDTS = Now
    objEncRTF.msUserNCID = mobjLogon.UserNCID
    Call AddEncRTF(objEncRTF)
    mbDirty = True
    Cancel = Save()
    If bDidOpen Then
        Call CloseEncounter
    Else 'If LeftWithoutBeingSeen Then SCR-16218
        'User canceling the current encounter
        If Not mobjParent Is Nothing Then
            mobjParent.CloseCurrentEncounter
        End If
    End If
    GoTo CleanUp
ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description, "Encounter.Cancel ", Exe, vbCritical)
CleanUp:
    Set objEncRTF = Nothing
    Set objSecParent = Nothing
End Function

Friend Function Load(ByVal PatientUnitNumber As String, _
                             ByVal FacilityNCID As String, _
                             ByVal EncounterNumber As String) As Boolean
    Dim colDataStrings As Collection
    On Error GoTo ErrHandler
    
    Load = GEMS_LoadEncounter(PatientUnitNumber, FacilityNCID, EncounterNumber)
    If Load Then
        'Reset RTF String SCR-24498
        msRTF = ""
    End If
    
    GoTo CleanUp
ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description, "Encounter.Load ", "EncounterOps", vbCritical)
CleanUp:
    Set colDataStrings = Nothing
End Function
'Refresh autocites w/o closing any modules SCR-37596
Public Sub RefreshAutocites()
    If Not mobjParent Is Nothing Then
        SectionParent(eAutoCite).mobjSection.Refresh
        mobjParent.RefreshDisplay False
    End If
End Sub

Public Sub RefreshDisplay()
    If mobjParent Is Nothing Then
        Call Refresh
    Else
        Call mobjParent.RefreshDisplay(True)
    End If
End Sub

Public Function Refresh(Optional ByVal StatusOnly As Boolean = False) As Boolean
    Dim colDataStrings As Collection
    Dim objWAMEncounter As CHCS_Interface.ICHCSEncounters
    Dim sCurrentProviderIEN  As String
    Dim sCurrentProviderName As String
    Dim dtDateLocked As Date
    
    On Error Resume Next
'<<GEMS

    'date:01/14/2004 SCR #:47870 developer: jrm Description: Added for ITT mode
    'Auto is needed to determine ITT, or Theatre; Mode needed is theatre logic gate

    If mobjShared.IsAppMode(modeCDR) And mobjShared.IsAppMode(modeITT) = False Then
        
        Set colDataStrings = GetDataStrings(mnPatientID, msFacilityNCID, mnEncounterID)
        If Not colDataStrings Is Nothing Then
            Refresh = LoadDataStrings(colDataStrings, , StatusOnly)
            Set colDataStrings = Nothing
            'Resume
        End If
    Else
                'Call Lite when refreshing status only SCR#49489
        Refresh = GEMS_LoadEncounter(mnPatientID, _
                                  msFacilityNCID, _
                                  mnEncounterID, _
                                  StatusOnly)
    End If
'>>
End Function

Private Function GetDataStrings(ByVal PatientUnitNumber As String, _
                               ByVal FacilityNCID As String, _
                               ByVal EncounterNumber As String) As Collection
    Dim objDAS As DasCtrl
    Dim colDataStrings As Collection
    
    If Not DebugMode Then On Error GoTo ErrHandler
    
    #If debugon Then
        Dim DR As DebugRec
        If Not DL Is Nothing Then DL.DebugLog Exe, Module, "GetDataStrings", "Begin", DR
    #End If
    
    Set objDAS = New DasCtrl
    objDAS.ClearAll
    objDAS.GeneralCriteria.ClearAll
    objDAS.GeneralCriteria.UnitNumber = PatientUnitNumber
    objDAS.GeneralCriteria.Facility.ValueConcept = FacilityNCID
    objDAS.GeneralCriteria.EncounterNumber = EncounterNumber
    objDAS.GeneralCriteria.QueryQualifier = "EVENT_TYPE_NCID in (" & NCID_ENCEVENT & "," & NCID_DAS_CLINICAL_TEXT & ")"
    objDAS.GeneralCriteria.ReturnMultiMedia = True

    Call mobjParent.EncounterOps.GEMS_DoSearch(objDAS, 0)

    If objDAS.EOF Then
        GoTo CleanUp
    End If
    Set colDataStrings = New Collection
    Do While Not objDAS.EOF
        colDataStrings.Add objDAS.Events
        objDAS.GoToNext
    Loop
    Set GetDataStrings = colDataStrings
    GoTo CleanUp
ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description, "Encounter.GetDataStrings ", "EncounterOps", vbCritical)
CleanUp:
    Set objDAS = Nothing
    Set colDataStrings = Nothing
#If debugon Then
    If Not DL Is Nothing Then DL.DebugLog Exe, Module, "GetDataStrings", "End", DR
#End If
End Function

Friend Function LoadDataStrings(ByRef colDataStrings As Collection, Optional ByVal Lite As Boolean = False, Optional ByVal StatusOnly As Boolean = False) As Boolean
    Dim objData         As DataString
    Dim objEvent        As ClinicalEvent
    Dim objCommentGroup As CommentGroup
    Dim objCommentItem  As CommentItem
    Dim objObs          As ClinicalObservation
    Dim objObsMod       As ClinicalObservation
    Dim objProvider     As EncProvider
    Dim objEncRTF       As EncRTF
    Dim objConcept      As ConceptCtrl
    Dim sName           As String
    Dim sValue          As String
    Dim lDocumentNumber As Long
    Dim i As Long
    Dim j As Long

    On Error GoTo ErrHandler

    If mobjShared.IsAppMode(modeCHCSI_GUI) Then '<SCR SCR 36493
        Exit Function
    End If
    '
    ' Find the datastring for header information and load it
    'Using the For...Each method with 3M objects causes memory leak
    For i = 1 To colDataStrings.Count
    'For Each objData In colDataStrings
        Set objData = colDataStrings(i)
        If objData.DataTypeNCID = NCID_ENCEVENT Then
            Set objEvent = objData.PatientData
            If objEvent.Header.ObsBatId.value.Coded.ValueConcept.UniqueId = SectionEnumToNCID(eHeader) Then
                Exit For
            End If
        End If
    Next
    If objData Is Nothing Then
        MsgBxARMd "Encounter data does not contain primary event.", vbExclamation, ExeModule
        GoTo CleanUp
    End If
    
    If StatusOnly Then
        meStatus = StatusNCIDToEnum(objEvent.Header.TestStatus.value.CodedWOSform.ValueConcept)
        Exit Function
    End If
    
    mbLoading = True
    mbFlagNew = False
    Set objConcept = New ConceptCtrl 'Not called in PG mode
    mbWasUpdated = False
    mdWhenSigned = 0
    mdWhenCosigned = 0
    Set mobjCosigner = Nothing
    Set mobjWhoSigned = Nothing
    Set mobjWhoCosigned = Nothing
    Set mcolEncRTFs = New Collection
    Set mcolProviders = New EncProviders
    lDocumentNumber = 0
    
    mnDataID = objData.DataId
    mnEncounterID = objData.EncounterNumber
    If mnEncounterID = 0 Then
        'MsgBxARMd "Encounter data contains a zero encounter id.", vbExclamation, ExeModule
        GoTo CleanUp
    End If
    mnPatientID = objData.UnitNumber
    msFacilityNCID = objData.EncounterFacility.ValueConcept
    msFacility = objData.EncounterFacility.ValueConcept.PrefRep("2000").Representation
    msClinicNCID = objData.CreatePointOfCare.PhysicalLocation.NursingDivision.ValueConcept
    msClinic = objData.CreatePointOfCare.PhysicalLocation.NursingDivision.ValueConcept.PrefRep("2000").Representation
'   Get PP from Comments SCR-43638
'    Set objProvider = New EncProvider
'    If objData.ModifyClinician Is Nothing Then
'        objProvider.NCID = objData.CreateClinician.ValueConcept.UniqueId
'        objProvider.FullName = objData.CreateClinician.ValueConcept.PrefRep("2000").Representation
'    Else
'        objProvider.NCID = objData.ModifyClinician.ValueConcept.UniqueId
'        objProvider.FullName = objData.ModifyClinician.ValueConcept.PrefRep("2000").Representation
'    End If
'    mcolProviders.Add objProvider, CStr(objProvider.NCID)
'    Set objProvider = Nothing
    mnUpdateCount = objData.ActionsInfo.Count
    msUpdateDate = objData.ActionsInfo.Item(mnUpdateCount).EnteredTime
    msUpdateClinicianNCID = objData.ActionsInfo.Item(mnUpdateCount).EnteredBy.ValueConcept.UniqueId
    With objEvent.Header
        mdStartDTS = .ObsDateTime.StartTime.value.DateTime
        mdEndDTS = .ObsDateTime.EndTime.value.DateTime
        If CDbl(mdEndDTS) < 1000 Then
            mdEndDTS = 0
        End If
        '.TestStatus.value
        
        meStatus = StatusNCIDToEnum(.TestStatus.value.CodedWOSform.ValueConcept)
        'Using the For...Each method with 3M objects causes memory leak
        For i = 1 To .Comments.Count
        'For Each objCommentGroup In .Comments
            For j = 1 To .Comments(i).Comments.Count
            'For Each objCommentItem In objCommentGroup.Comments
                Set objCommentItem = .Comments(i).Comments(j)
                Call ParseCommentItem(objCommentItem, sName, sValue)
                Select Case sName
                Case "WASUPDATED"
                    mbWasUpdated = sValue
                Case "COSIGNER"
'jf todo: take this out after a while, as correct way of storing cosigner ncid takes hold
                    Set mobjCosigner = New EncProvider
                    mobjCosigner.NCID = sValue
                    objConcept.UniqueId = sValue
                    mobjCosigner.FullName = objConcept.PrefRep("2000").Representation
                Case "WHENCOSIGNED"
                    mdWhenCosigned = sValue
                Case "WHOCOSIGNED"
                    Set mobjWhoCosigned = New EncProvider
                    mobjWhoCosigned.NCID = sValue
                    objConcept.UniqueId = sValue
                    mobjWhoCosigned.FullName = objConcept.PrefRep("2000").Representation
                Case "COSIGNERLINE2"
                    If Not mobjWhoCosigned Is Nothing Then
                        mobjWhoCosigned.SigLine2 = sValue
                    End If
                Case "COSIGNERLINE3"
                    If Not mobjWhoCosigned Is Nothing Then
                        mobjWhoCosigned.SigLine3 = sValue
                    End If
                Case "WHENSIGNED"
                    mdWhenSigned = CDate(sValue)
                Case "WHOSIGNED"
                    Set mobjWhoSigned = New EncProvider
                    mobjWhoSigned.NCID = sValue
                    objConcept.UniqueId = sValue
                    mobjWhoSigned.FullName = objConcept.PrefRep("2000").Representation
                Case "SIGNERLINE2"
                    If Not mobjWhoSigned Is Nothing Then
                        mobjWhoSigned.SigLine2 = sValue
                    End If
                Case "SIGNERLINE3"
                    If Not mobjWhoSigned Is Nothing Then
                        mobjWhoSigned.SigLine3 = sValue
                    End If
                Case "ENCCAT"
                    msCategory = sValue
                Case "TIMEZONE"
                    msTimeZone = sValue
                Case "ENCTYPE"
                    meType = Val(sValue)
                Case "ENCNOCOUNT"
'                    ec_Followup = 1
'                    ec_New = 2
'                    ec_NoCount = 3
                    Select Case Val(sValue)
                        Case 1 'ec_Followup = 1
                            meClass = ec_Followup
                        Case 2 'ec_New = 2
                            meClass = ec_New
                        Case 3 'ec_NoCount = 3
                            meClass = ec_NoCount
                            '<< SCR 10086
                            msSADRStatus = NCID_SADR_NO_COUNT
                        Case Else
                    End Select
                Case "EANDMCODEREVIEW"          'DEM - 05/04/00
                    msEandMCodeReview = sValue
                Case "ALLERGYVERIFIEDBYDATE"
                    mdAllergyVerifiedByDate = sValue
                Case "ALLERGYVERIFIEDBYNCID"
                    msAllergyVerifiedByNCID = sValue
                Case "SENSITIVE" 'SCR-35365
                    If StrComp(sValue, "T") = 0 Then
                        'mbSensitivity = True
                        mlSensitivity = esrStandard
                    ElseIf IsNumeric(sValue) Then
                        mlSensitivity = CLng(sValue)
                    End If
                    
                Case "ARRIVEDBYAMBULANCE"
                    If StrComp(sValue, "T") = 0 Then
                        mbArrivedByAmbulance = True
                    End If
                Case "APPTDTS"
                    mdApptDTS = CDate(sValue)
                Case "UPDATELOCK"
                    mblnUpdateLock = CBool(sValue)
                Case "DXSnoID"
                    mnPrimaryDxSnoID = CLng(sValue)
                Case "DXMEDCINPREFIX"
                    msPrimaryDxMedcinPrefix = sValue
                Case "RESULTEDCONSULTORDER"
                    If StrComp(sValue, "T") = 0 Then
                        mbResultedConsultOrder = True
                    End If
                Case "APSTATUSCOMMENT"  'SCR-17942
                    msAPStatusComment = sValue
                Case "WORKLOADOPTIONS"
                    msWorkLoadOptions = sValue
                Case "WORKLOAD"
                    msWorkLoad = sValue
                Case "PATIENTSTATUS"
                    msPatientStatus = sValue
                Case "MEPRSCODES"
                    msMEPRSCodes = sValue
                Case "APPTIEN"
                    msApptIEN = sValue
                Case "TEMPLATEID"
                    msTemplateID = sValue 'SCR-44706 RBELL 11/3/03 retrieve template id data from db text object
                Case "PRIMPROV"
                    'Modify Clinician is unreliable SCR-43638
                    Set objProvider = New EncProvider
                    objProvider.NCID = sValue
                    objConcept.UniqueId = sValue
                    objProvider.FullName = objConcept.PrefRep("2000").Representation
                    objProvider.RoleNCID = NCID_Attending_Provider      'Default value when loading encounters saved with previous version so old encounters will work with HIPAA
                    mcolProviders.Add objProvider, CStr(objProvider.NCID)
                    Set objProvider = Nothing
                Case "HIPAA_CODE" 'SCR#43283
                    msProvHIPAACode = sValue
                End Select
                
                
            Next 'objCommentItem
        Next 'objCommentGroup
    End With
    
    lDocumentNumber = 1
    
    For i = 1 To objEvent.ClinObs.Count
    'For Each objObs In objEvent.ClinObs
        Set objObs = objEvent.ClinObs(i)
        
        Select Case objObs.ObsId.value.CodedWOSform.ValueConcept
        Case NCID_NEEDS_COSIGN
            Set mobjCosigner = New EncProvider
            mobjCosigner.NCID = objObs.ObsValue.value.CodedWOSform.ValueConcept
            mobjCosigner.FullName = objObs.ObsValue.value.CodedWOSform.ValueConcept.PrefRep("2000").Representation
        Case NCID_ENCOUNTERRTF
'            If Not Lite Then
                Dim objInfo As EventActionInfo
                
                If objObs.ObsMods.Count > 2 Then
                    'Encounter RTF may be saved in sections 'SCR#42412
                    'added to build 11/4/03
                    'Checks to make sure RTF is stored in this ClinObs
                    If StrComp(objObs.ObsMods(3).ObsValue.value.InfoType, "basicString", vbTextCompare) = 0 Then
                        If StrComp(objObs.ObsMods(3).ObsValue.value.BasicString, "RTF", vbTextCompare) = 0 Then
                            'Only need to extract RTF once per document
                            If objObs.ObsMods(1).ObsValue.value.ULong = lDocumentNumber Then
                                Set objEncRTF = New EncRTF
                                objEncRTF.msRTF = ExtractRTF2(objEvent, lDocumentNumber)

                                lDocumentNumber = lDocumentNumber + 1
                            End If
                        End If
                    End If
                Else
                    'Encounter old way
                    Set objEncRTF = New EncRTF
                    objEncRTF.msRTF = ExtractRTF(objObs)
                End If
                
                If Not objEncRTF Is Nothing Then
                    If objObs.ActionsInfo.Count <> 0 Then
                    Set objInfo = objObs.ActionsInfo.Item(1)
                    objEncRTF.mdDTS = objObs.ActionsInfo.Item(1).EnteredTime
                    objEncRTF.msUserNCID = objInfo.Clinician.ValueConcept.UniqueId
                    objEncRTF.mdDTS = objInfo.EffectiveTime
                    
                    Set objInfo = Nothing
                    mcolEncRTFs.Add objEncRTF ' let these remain in descending date order
                    End If
                    Set objEncRTF = Nothing
                End If
'            End If
        Case NCID_ADDITIONALPROVIDER
            If Len(mobjShared.CmdLineSwitches("HIPAA837")) > 0 Then
               Set objProvider = New EncProvider
               objProvider.CDRLoad objObs
            Else
               Set objProvider = New EncProvider
               objProvider.NCID = objObs.ObsValue.value.CodedWOSform.ValueConcept
               objProvider.FullName = objObs.ObsValue.value.CodedWOSform.ValueConcept.PrefRep("2000").Representation
               For j = 1 To objObs.ObsMods.Count
               'For Each objObsMod In objObs.ObsMods
                   Set objObsMod = objObs.ObsMods(j)
                   Select Case objObsMod.ObsId.value.CodedWOSform.ValueConcept
                   Case NCID_PROVIDER_ROLE
                       objProvider.RoleNCID = objObsMod.ObsValue.value.CodedWOSform.ValueConcept
                       objProvider.Role = objObsMod.ObsValue.value.CodedWOSform.ValueConcept.PrefRep("2000").Representation
                   End Select
               Next 'objObsMod
            End If
            
            Dim bAddProvider As Boolean

            bAddProvider = True

            'Check if provider alredy exists in collection SCR-25124
            For j = 1 To mcolProviders.Count
                If mcolProviders(j).NCID = objProvider.NCID Then
                    bAddProvider = False
                End If
            Next
            
            If bAddProvider Then
                mcolProviders.Add objProvider, CStr(objProvider.NCID)
            ElseIf mcolProviders.Count = 1 And Len(mobjShared.CmdLineSwitches("HIPAA837")) > 0 Then
               'when HIPAA is enabled check to see if this is the same NCID as the primary provider
               'if it is we should replace the one loaded with the provider retrieved from the new CDR storage location
               mcolProviders.Remove 1
               mcolProviders.Add objProvider, CStr(objProvider.NCID)
            End If
            
        Case NCID_DIAGNOSIS
            msPrimaryDiagnosisNCID = objObs.ObsValue.value.CodedWOSform.ValueConcept.UniqueId
        Case NCID_SADR_STATUS
            msSADRStatus = objObs.ObsValue.value.CodedWOSform.ValueConcept
                        
            If msSADRStatus = NCID_SADR_NO_COUNT Then 'SCR-21981
                meClass = ec_NoCount
            End If
            
        Case NCID_APPOINTMENT_ID
            msApptID = objObs.ObsValue.value.BasicString
        Case NCID_APPOINTMENT_TYPE:
        Case NCID_SECURITY_LABEL:
        Case NCID_ENC_CATEGORY:
        Case NCID_ENC_TEMPLATE_ID:
            msTemplateData = StrConv(objObs.ObsValue.value.MultiMedia.Data, vbUnicode)
        End Select
    Next 'objObs
    
    If Not Lite Then
        Set mobjLocking = New Locking
        mobjLocking.msFacilityNCID = FacilityNCID
        mobjLocking.msEncounterNumber = mnEncounterID
        mobjLocking.msUserNCID = mobjLogon.UserNCID
        Call CreateSections(colDataStrings)
        mbFullyLoaded = True
    End If

    LoadDataStrings = True
    GoTo CleanUp
ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description, "Encounter.Load ", "EncounterOps", vbCritical)
CleanUp:
    mbLoading = False
    'Resume
    mbDirty = False
    Set objData = Nothing
    Set objEvent = Nothing
    Set objCommentGroup = Nothing
    Set objCommentItem = Nothing
    Set objObs = Nothing
    Set objObsMod = Nothing
    Set objProvider = Nothing
    Set objConcept = Nothing
    Set objEncRTF = Nothing
End Function

Public Function CloseEncounter() As Boolean
    #If debugon Then
        Dim DR As DebugRec
        If Not DL Is Nothing Then DL.DebugLog Exe, Module, "CloseEncounter", "Begin", DR
    #End If

    Dim objSecParent As SectionParent
'<< SCR 10768  check with curr section if it's up, ask if it can close
    If Not CurrSection Is Nothing Then
        'CurrSection.msOKToCloseView = ""
        ' SCR 11207
        '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
        ' New Message Parameter of False indicates not to create objects if they are not instantiated.
    #If debugon Then
        If Not DL Is Nothing Then DL.DebugLog Exe, Module, "CloseEncounter", "Message cwiOK_TO_CLOSE_VIEW", DR, "CurrSection.msAppNCID", CurrSection.msAppNCID
    #End If
        
        'mobjComm.Message cwiOK_TO_CLOSE_VIEW, "CurrSection.msOKToCloseView", CurrSection.msAppNCID, NCID_ENCOUNTER_APPLICATION, False
        'mobjComm.Message cwiOK_TO_CLOSE_VIEW, "", CurrSection.msAppNCID, NCID_ENCOUNTER_APPLICATION, False
        mobjComm.Message cwiOK_TO_CLOSE_VIEW, CurrSection.msOKToCloseView, CurrSection.msAppNCID, NCID_ENCOUNTER_APPLICATION, False
        'DoEvents
        '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    #If debugon Then
        If Not DL Is Nothing Then DL.DebugLog Exe, Module, "CloseEncounter", "After message ok-to-close", DR
    #End If
        If Not CurrSection Is Nothing Then 'SCR -24514
            If UCase$(CurrSection.msOKToCloseView) = "N" Then
                CloseEncounter = False
                Exit Function
            End If
        End If
    End If
'>>
    
    On Error Resume Next
    Call CloseSection
    #If debugon Then
        If Not DL Is Nothing Then DL.DebugLog Exe, Module, "after close section", "End", DR
    #End If

    For Each objSecParent In mcolSecParents
        Call objSecParent.CloseSection
        Call objSecParent.CloseEncounter
    Next objSecParent
    Set mobjParent = Nothing
    mbOpen = False
    CloseEncounter = True
    
    #If debugon Then
        If Not DL Is Nothing Then DL.DebugLog Exe, Module, "CloseEncounter", "End", DR
    #End If

End Function

Public Function OpenEncounter(ByRef Parent As Object) As Boolean
    Dim objSecParent As SectionParent
    On Error Resume Next

    Set mobjParent = Parent
    If Not mobjProperties Is Nothing Then mobjProperties.RetrievedProperties = False
    For Each objSecParent In mcolSecParents
        Call objSecParent.OpenEncounter(Me)
    Next objSecParent
    'Re-init the Template History so it's fresh upon opening the encounter
    Set mobjParent.EncounterOps.TemplateHx = New Collection
    OpenEncounter = True
    mbOpen = True


End Function

Friend Sub SectionClosed()
    Dim bWasLocked As Boolean
    Dim bResult As Boolean
    Set mobjCurrSection = Nothing
    If Not mbDirty Then
        Exit Sub
    End If
    bWasLocked = mobjLocking.SectionIsLocked(eHeader, 0)
    If Not bWasLocked Then
        If Not mobjLocking.LockSection(eHeader, 0, SectionLock) Then
            If MsgBxARMd("Warning: the encounter header needs to be updated in the database." _
            & vbCrLf & vbCrLf & "Do you want to try again?", vbYesNo) = vbNo Then
                GoTo CleanUp
            End If
            If Not mobjLocking.LockSection(eHeader, 0, SectionLock) Then
                GoTo CleanUp
            End If
        End If
    End If
    bResult = Save()
CleanUp:
    If bResult Then
        Call mobjComm.SetStatusText("Encounter " & mnEncounterID & " was updated.")
    Else
        Call mobjComm.SetStatusText("Encounter " & mnEncounterID & " was not saved.")
    End If
    If Not bWasLocked Then
        Call mobjLocking.UnlockSection(eHeader, 0)
    End If
End Sub

Public Function CloseSection() As Boolean
    On Error Resume Next
    #If debugon Then
        Dim DR As DebugRec
        If Not DL Is Nothing Then DL.DebugLog Exe, Module, "CloseSection", "Begin", DR
    #End If
    
    If Not mobjCurrSection Is Nothing Then
        If mobjCurrSection.CloseSection Then
            Set mobjCurrSection = Nothing
            CloseSection = True
        End If
    End If
    #If debugon Then
        If Not DL Is Nothing Then DL.DebugLog Exe, Module, "CloseSection", "End", DR
    #End If

End Function

Public Function OpenSection(ByVal eSection As EncounterSectionEnum) As Boolean

    Dim bAlreadyOpen As Boolean
    On Error Resume Next
    If Not mobjParent Is Nothing Then
        Call mobjParent.AllowUI(False)
        If Not mobjCurrSection Is Nothing Then
            If mobjCurrSection.meSection <> eSection Then
                If Not mobjCurrSection.CloseSection Then
                    'SCR 102788
                    If Not mobjCurrSection Is Nothing Then
                        If mobjCurrSection.meSection = eSO Then
                            If Not mobjCurrSection.CloseSection Then
                                Call mobjParent.AllowUI(True)
                                Exit Function
                            End If
                        Else
                            Call mobjParent.AllowUI(True)
                            Exit Function
                        End If
                    Else
                        Call mobjParent.AllowUI(True)
                        Exit Function
                    End If
                End If
            Else
            
                'Send refresh to core to update menu items SCR-23986
                mobjComm.Message cwiGENERIC, "REFRESH", NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION
                
                'If the curr section is the section we want to open
                'then don't close and reopen section
                bAlreadyOpen = True
            End If
        End If
        
        If Not bAlreadyOpen Then
            Set mobjCurrSection = mcolSecParents(CStr(eSection))
            OpenSection = mobjCurrSection.OpenSection
            If eSection = eAutoCite Or (Not OpenSection) Then
                Set mobjCurrSection = Nothing
                mobjParent.hidechild 'SCR-22897
            End If
        Else
            OpenSection = True
        End If
        Call mobjParent.AllowUI(True)
    End If
End Function

Public Sub MenuItemSelected(ByVal MenuID As Long)
    On Error Resume Next
    Call mobjCurrSection.MenuItemSelected(MenuID)
End Sub

Public Function UndoCancel() As Boolean
    Dim bDidOpen As Boolean
    On Error GoTo ErrHandler
    If Not mbOpen Then
        bDidOpen = OpenEncounter(Nothing)
        If Not mbOpen Then
            Exit Function
        End If
    End If
    If Not mobjLocking.LockSection(eEntireEncounter, 0, SectionLock, "Undo Cancel") Then
        GoTo CleanUp
    End If
    If SectionRTF(eap) = "" _
    And SectionRTF(eDisp) = "" _
    And SectionRTF(eSO) = "" Then
        If SectionRTF(eVitals) = "" _
        And SectionRTF(eRFV) = "" Then
            Status = CheckedIn
        Else
            Status = Waiting
        End If
    Else
        Status = InProgress
    End If
    UndoCancel = Save(True)
    GoTo CleanUp
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "Encounter.UndoCancel", App.Title, vbCritical
CleanUp:
    If mobjLocking.SectionIsLocked(eEntireEncounter, 0) Then
        Call mobjLocking.UnlockSection(eEntireEncounter, 0)
    End If
    If bDidOpen Then
        Call CloseEncounter
    End If
End Function

Public Function UserIsPrimaryProvider() As Boolean
    If mcolProviders.Count > 0 Then
        If CLng(mobjLogon.UserNCID) = mcolProviders(1).NCID Then
            'Also do a name check
            If mobjLogon.UserName = mcolProviders(1).FullName Then
                UserIsPrimaryProvider = True
            End If
        End If
    End If
End Function

Public Function UserIsCosigner() As Boolean

Dim objDAS As ICHCSII_DAS
Dim oRS As Recordset

    If Not mobjCosigner Is Nothing Then
        If mobjCosigner.NCID = mobjLogon.UserNCID Then
            UserIsCosigner = True
        End If
    Else
        'SF 14858 Go to the DB to find the cosigner.
        'For some reason in GEMS_LoadEncounter line 4280, they don't want to load the cosigner if he/she already signed.
        If Len(mobjComm.CmdLineSwitch("INPATIENT2")) > 0 Then
            Set objDAS = gobjCHCSIIConn.CHCSII_DAS
            Set oRS = objDAS.OpenRecordset("Select CosignerNCID from Encounters where EncounterNumber = " & Me.EncounterID & " and CosignerNCID = " & mobjLogon.UserNCID)
            If oRS.RecordCount > 0 Then
                UserIsCosigner = True
            End If
            oRS.Close
            Set oRS = Nothing
        End If
    End If
End Function


Friend Function GEMS_LoadEncounter(ByVal PatientUnitNumber As String, _
                                   ByVal FacilityNCID As String, _
                                   ByVal EncounterNumber As String, _
                                   Optional ByVal Lite As Boolean = False, _
                                   Optional ByVal ApptID As String = "") As Boolean
'- fill the das ctrl with information requested in the criteria sections
    Dim objDAS As ICHCSII_DAS
    
    Set objDAS = gobjCHCSIIConn.CHCSII_DAS(Auto) '''<SCR 36493
    
    Dim objRS As ADODB.Recordset
    Dim sSql As String
    Dim objProvider     As EncProvider

    Dim objConcept      As GEMSConceptCtrl.GEMS_ConceptCtrl
    Dim sName           As String
    Dim sValue          As String
    
    Dim oSQLEx As ICHCSII_SQL.ISqlOps
    Dim oSQLEx2 As ICHCSII_SQL.ISqlOpsEx
    

    Set oSQLEx = gobjCHCSIIConn.CHCSII_SQLOPS(Auto) '''<SCR 36493

    On Error Resume Next
    mbLoading = True
    mbFlagNew = False
    Set objConcept = New GEMSConceptCtrl.GEMS_ConceptCtrl
    mbWasUpdated = False
    mdWhenSigned = 0
    mdWhenCosigned = 0
    Set mobjCosigner = Nothing
    Set mobjWhoSigned = Nothing
    Set mobjWhoCosigned = Nothing
    Set mcolEncRTFs = New Collection
    Set mcolProviders = New EncProviders

   'Build the encounters select statement
   If mobjShared.IsAppMode(modeTheater) Or Len(ApptID) = 0 Then
      sSql = "select E.* from Encounters E where E.EncounterNumber = " & EncounterNumber _
            & " and E.FacilityNCID = " & FacilityNCID & " and E.Unit_Number = " & PatientUnitNumber
   Else
      'USE Apptid to see if encounter has already been created
      sSql = "select E.* from Encounters E where E.APPTID = '" & ApptID & "'" _
            & " and E.FacilityNCID = " & FacilityNCID & " and E.Unit_Number = " & PatientUnitNumber
   End If
   
    oSQLEx.Execute sSql
    If oSQLEx.EOF Then
        If mobjShared.IsAppMode(modeTheater) Then
            MsgBxARMd "Encounter data does not contain primary event.", vbExclamation, ExeModule
        End If
        GoTo CleanUp
    End If
    
    mnDataID = oSQLEx("EncounterNumber")
    mnEncounterID = oSQLEx("EncounterNumber")
    If mnEncounterID = 0 Then
        MsgBxARMd "Encounter data contains a zero encounter id.", vbExclamation, ExeModule
        GoTo CleanUp
    End If

    mnPatientID = "" & oSQLEx("Unit_Number")
    msFacilityNCID = "" & oSQLEx("FacilityNCID")
    objConcept.UniqueId = msFacilityNCID
    
    msFacility = objConcept.PrefRep("2000").Representation
    
    msClinicNCID = "" & oSQLEx("ClinicNCID")
    objConcept.UniqueId = msClinicNCID
    
    msClinic = objConcept.PrefRep("2000").Representation
    
    'SCR 55063
    mlSensitivity = oSQLEx("Sensitive")
    
    Set objProvider = New EncProvider
    objProvider.NCID = "" & oSQLEx("primaryproviderncid")
    objConcept.UniqueId = objProvider.NCID
    
    objProvider.FullName = objConcept.PrefRep("2000").Representation
    objProvider.RoleNCID = rtnAttendingProvider             'Set default provider role in case this encounter was created with a non HIPPA version
    mcolProviders.Add objProvider, CStr(objProvider.NCID)

   'Load HIPAA Data Elements from SQL object
   If Len(mobjComm.CmdLineSwitch("HIPAA837")) > 0 Then
      If mobjShared.IsAppMode(modeITT) Then
         mobjAccidentData.InitFromSQL ClinicNCID, oSQLEx
         mobjPregnancyData.CDRPlusLoad Me.FacilityNCID, Me.EncounterID, mobjPatient
      ElseIf mobjShared.IsAppMode(modeCHCSI_GUI) Then
         mobjAccidentData.InitFromSQL ClinicNCID, oSQLEx
         mobjPregnancyData.InitFromSQL oSQLEx, mobjPatient
      End If
   End If
   
    If "" & oSQLEx("UpdatedBy") = "" Then
        Set objProvider = Nothing
        Set objProvider = New EncProvider
        objProvider.NCID = "" & oSQLEx("CreatedBy")
        objConcept.UniqueId = objProvider.NCID
        'objProvider.FullName = objConcept.PrefRep("2000").Representation
    Else
        Set objProvider = Nothing
        Set objProvider = New EncProvider
        objProvider.NCID = "" & oSQLEx("UpdatedBy")
        objConcept.UniqueId = objProvider.NCID
        'objProvider.FullName = objConcept.PrefRep("2000").Representation
    End If

    If objProvider.NCID <> "" & oSQLEx("primaryproviderncid") And objProvider.NCID <> 0 Then
    
        objProvider.FullName = objConcept.PrefRep("2000").Representation
        
        mcolProviders.Add objProvider, CStr(objProvider.NCID)
    End If
    Set objProvider = Nothing


    '<< tfb change to handle null value 07-11-2001
    If oSQLEx("UpdatedOn") <> "" Then
        msUpdateDate = "" & oSQLEx("UpdatedOn")  'note:  msUpdateDate is a date data type
    Else
        msUpdateDate = CDate(0)
    End If

    If Not oSQLEx("UpdatedBy") = 0 Then
        msUpdateClinicianNCID = "" & oSQLEx("UpdatedBy")
    Else
        msUpdateClinicianNCID = vbNullString
    End If

    If Not oSQLEx("CreatedBy") = 0 Then 'SCR#42996
        msCreateClinicianNCID = "" & oSQLEx("CreatedBy")
    Else
        msCreateClinicianNCID = vbNullString
    End If
    
    If oSQLEx("StartDTS") <> "" Then
        mdStartDTS = oSQLEx("StartDTS")
    Else
        mdStartDTS = CDate(0)
    End If

    If oSQLEx("EndDTS") <> "" Then
        mdEndDTS = oSQLEx("EndDTS")
    Else
        mdEndDTS = CDate(0)
    End If
    
    If oSQLEx("APPTDTS") <> "" Then
        mdApptDTS = oSQLEx("APPTDTS")
    Else
        mdApptDTS = CDate(0)
    End If
    
    '>> tfb change end

    If CDbl(mdEndDTS) < 1000 Then
        mdEndDTS = 0
    End If
'<< SCR  10086 moved from below to have this set before status is updated
    If Val("" & oSQLEx("NOCOUNT")) > 0 Then
        meClass = ec_NoCount
    End If
'>>
    meStatus = StatusNCIDToEnum("" & oSQLEx("Status"))
    mbWasUpdated = (oSQLEx("WasUpdated") = "Y")
    'Amending does not clear who signed NCID so do not load instead
    'If meStatus <> Updating Then
        If oSQLEx("WHENCOSIGNED") <> "" Then
            mdWhenCosigned = oSQLEx("WHENCOSIGNED")
        End If
        Set mobjWhoCosigned = New EncProvider
        mobjWhoCosigned.NCID = "" & oSQLEx("WHOCOSIGNEDNCID")
        If Not mobjWhoCosigned.NCID = 0 Then
            objConcept.UniqueId = mobjWhoCosigned.NCID
            mobjWhoCosigned.FullName = objConcept.PrefRep("2000").Representation
            mobjWhoCosigned.SigLine2 = "" & oSQLEx("COSIGNERLINE2")
            mobjWhoCosigned.SigLine3 = "" & oSQLEx("COSIGNERLINE3")
        End If
        mdWhenSigned = CDate("" & oSQLEx("WhenSigned"))
        Set mobjWhoSigned = New EncProvider
        mobjWhoSigned.NCID = "" & oSQLEx("WHOSIGNEDNCID")
        If Not mobjWhoSigned.NCID = 0 Then
            objConcept.UniqueId = mobjWhoSigned.NCID
            If mobjWhoSigned.NCID = PrimaryProvider.NCID Then
                
                mobjWhoSigned.FullName = PrimaryProvider.FullName
            Else
                mobjWhoSigned.FullName = objConcept.PrefRep("2000").Representation
            End If
            mobjWhoSigned.SigLine2 = "" & oSQLEx("SIGNERLINE2")
            mobjWhoSigned.SigLine3 = "" & oSQLEx("SIGNERLINE3")
        Else
            Set mobjWhoSigned = Nothing
        End If
    'End If
    msCategory = "" & oSQLEx("Category")
    msTimeZone = "" & oSQLEx("TIMEZONE")
    meType = "" & oSQLEx("TYPE")
    
    If Len(mobjComm.CmdLineSwitch("INPATIENT2")) > 0 Then
        msInptNoteTypeNCID = oSQLEx("INPT_NOTE_TYPE_NCID") & vbNullString
        mlInptHospDay = oSQLEx("INPT_HOSP_DAY") & vbNullString
    End If
    
    If IsDate(oSQLEx("ALLERGYVERIFIEDBYDATE")) Then
        mdAllergyVerifiedByDate = oSQLEx("ALLERGYVERIFIEDBYDATE")
    End If
    msAllergyVerifiedByNCID = "" & oSQLEx("ALLERGYVERIFIEDBYNCID")

    If Not oSQLEx("CoSignerNCID") = 0 Then
    '- if the one who is to cosign has cosigned, don't add him again to the list
        If Not "" & oSQLEx("CoSignerNCID") = "" & oSQLEx("WHOCOSIGNEDNCID") Then
            Set mobjCosigner = New EncProvider
            mobjCosigner.NCID = "" & oSQLEx("CoSignerNCID")
            objConcept.UniqueId = mobjCosigner.NCID
            mobjCosigner.FullName = objConcept.PrefRep("2000").Representation
        End If
    End If

    If Not Val(oSQLEx("AdditionalProvider1NCID")) = 0 Then
        Set objProvider = New EncProvider
        objProvider.NCID = "" & oSQLEx("AdditionalProvider1NCID")
        objProvider.RoleNCID = "" & oSQLEx("APROLE1NCID") 'SCR-19027 'Wrong column name oSQLEx("APROLE1")
        
        If Not mobjShared.IsAppMode(modeCHCSI_GUI) Then
            objConcept.UniqueId = objProvider.NCID
            objProvider.FullName = objConcept.PrefRep("2000").Representation
            objConcept.UniqueId = objProvider.RoleNCID
            objProvider.Role = objConcept.PrefRep("2000").Representation
        Else
            objProvider.FullName = GetProviderNameCHCS(objProvider.NCID)
            
            Select Case Trim$(objProvider.RoleNCID)
                Case "4"
                    objProvider.Role = "Nurse"
                Case "2"
                    objProvider.Role = "Assisting"
                Case "3"
                    objProvider.Role = "Supervising"
                Case "5"
                    objProvider.Role = "Para-professional"
                Case "1"          '
                    objProvider.Role = "Attending"
            End Select
        End If

        mcolProviders.Add objProvider, CStr(objProvider.NCID)
    End If
    If Not Val(oSQLEx("AdditionalProvider2NCID")) = 0 Then
        Set objProvider = New EncProvider
        objProvider.NCID = "" & oSQLEx("AdditionalProvider2NCID")
        objProvider.RoleNCID = "" & oSQLEx("APROLE2NCID") 'SCR-19027 'Wrong column name oSQLEx("APROLE2")

        If Not mobjShared.IsAppMode(modeCHCSI_GUI) Then
            objConcept.UniqueId = objProvider.NCID
            objProvider.FullName = objConcept.PrefRep("2000").Representation
            objConcept.UniqueId = objProvider.RoleNCID
            objProvider.Role = objConcept.PrefRep("2000").Representation
        Else
            objProvider.FullName = GetProviderNameCHCS(objProvider.NCID)
            Select Case Trim$(objProvider.RoleNCID)
                Case "4"
                    objProvider.Role = "Nurse"
                Case "2"
                    objProvider.Role = "Assisting"
                Case "3"
                    objProvider.Role = "Supervising"
                Case "5"
                    objProvider.Role = "Para-professional"
                Case "1"          '
                    objProvider.Role = "Attending"
            End Select
        End If
        mcolProviders.Add objProvider, CStr(objProvider.NCID)
    End If

    msPrimaryDiagnosisNCID = "" & oSQLEx("PrimarydiagnosisNCID")
    mnPrimaryDxSnoID = "" & oSQLEx("DXSnoID")   ' Load Medcin ID & Prefix SCR-19435
    msPrimaryDxMedcinPrefix = "" & oSQLEx("DXMEDCINPREFIX")
    msSADRStatus = "" & oSQLEx("SADRStatus")
    msApptID = "" & oSQLEx("APPTID")
    
    'SCR 128313
    mbIsBMISTEnc = Val(oSQLEx("BMIST_ENC") & "")

    'Get Template Data
    If mobjShared.AppMode = modeCHCSI_GUI Then
      msTemplateData = mobjShared.CHCSConnection.ReadChunk("ENCOUNTERS", "ENCOUNTERNUMBER = " & mnEncounterID, "TEMPLATEDATA")
      msApptIEN = msApptID
    Else
        Set oSQLEx2 = gobjCHCSIIConn.CHCSII_SQLOPS_EX(Auto)
          
        oSQLEx2.Execute "Select TemplateData from Encounters " & vbCrLf & _
        "where ENCOUNTERNUMBER = " & CStr(mnEncounterID) & vbCrLf & _
        "AND FACILITYNCID = " & msFacilityNCID
        
        If Not (oSQLEx2.BOF And oSQLEx2.EOF) Then
            oSQLEx2.GetAsChunk "TemplateData", msTemplateData
        End If
        
        Set oSQLEx2 = Nothing
    End If
    
    If mobjShared.IsAppMode(modeCDR) Then
      'Application is in CDR mode, set GEMS encounter object UpdateLock value = false
      mblnUpdateLock = False
    Else
      mblnUpdateLock = CBool(oSQLEx("UpdateLock"))
    End If
    
    msPatientStatus = "" & oSQLEx("Patient_Status")
    
    If mobjShared.IsAppMode(modeTheater) Then
        msInpatientService = "" & oSQLEx("INPATIENT_SERVICE")
    End If
    
 '==========================
 '--RTF
   Call GEMS_LoadRTFs(FacilityNCID, _
                      EncounterNumber)
 '==========
    
    If Not Lite Then
        Set mobjLocking = New Locking
        mobjLocking.msFacilityNCID = FacilityNCID
        mobjLocking.msEncounterNumber = mnEncounterID
        mobjLocking.msUserNCID = mobjLogon.UserNCID
        Call GEMS_CreateSections(PatientUnitNumber, _
                                 FacilityNCID, _
                                 EncounterNumber, _
                                 objDAS)
        mbFullyLoaded = True
    End If
    
    GEMS_LoadEncounter = True
    GoTo CleanUp
    
ErrHandler:
    Call mobjShared.ShowVBError(Err.Number, Err.Description, "Encounter.GEMS_LoadEncounter ", "Encounter", vbCritical)

CleanUp:
    mbLoading = False
    mbDirty = False
    Set objProvider = Nothing
    Set objConcept = Nothing
   ' If objRS.State > 0 Then objRS.Close
    Set objRS = Nothing
    Set objDAS = Nothing
    'Screen.MousePointer = vbDefault 'Commented out for SCR-23678
End Function

Private Function GEMS_LoadRTFs(ByVal FacilityNCID As String, _
                               ByVal EncounterNumber As String) As Boolean
    Dim objEncRTF As EncRTF
    Dim objSQL    As ICHCSII_SQL.ISqlOps

    'date:01/14/2004 SCR #:47870 developer: jrm Description: Added for ITT mode
    'Auto is needed to determine ITT, or Theatre; Mode needed is theatre logic gate

    Set objSQL = gobjCHCSIIConn.CHCSII_SQLOPS(Auto) '''<SCR 36493
    
    Dim sSql      As String
    Dim aDoc      As Variant
    
    Dim oShared   As CWShared
    
    Set oShared = New CWShared
    
    sSql = "Select * from Enc_RTFs where EncounterNumber = " & EncounterNumber _
        & " and FacilityNCID = " & FacilityNCID _
        & " and rtftype =  'RTF' " _
        & " Order by DTS DESC"
    objSQL.Execute sSql
    
    Do While Not objSQL.EOF
        Set objEncRTF = New EncRTF
        
        objEncRTF.mlDataID = objSQL("ENC_RTFSINDEX")
        'Call objSQL.GetAsChunk("Doc", aDoc)

        If oShared.IsAppMode(modeCHCSI_GUI) Then
            objEncRTF.msRTF = oShared.CHCSConnection.ReadChunk("Enc_RTFs", "EncounterNumber = " & EncounterNumber & " and FacilityNCID = " & FacilityNCID & " and rtftype =  'RTF' and ENC_RTFSINDEX = " & objSQL("ENC_RTFSINDEX"), "DOC")
        Else
            'Decompress if compressed SCR-24683
            If CVar(objSQL("OriginalSize")) > 0 Then
                Dim objZlib As EncZLib
                Set objZlib = New EncZLib
                Dim aBytes() As Byte
                
                aBytes = objSQL("Doc")
                Call objZlib.DecompressData(aBytes, objSQL("OriginalSize"))
                objEncRTF.msRTF = StrConv(aBytes, vbUnicode)
                'objEncRTF.msRTF = DecompressRTF(CStr(aDoc), objSQL("OriginalSize"))
            Else
                objEncRTF.msRTF = objSQL("DOC")
            End If
        End If

        
        objEncRTF.mdDTS = objSQL("DTS")
        objEncRTF.msUserNCID = objSQL("UserNCID")
 
        mcolEncRTFs.Add objEncRTF
        Set objEncRTF = Nothing
        objSQL.MoveNext
    Loop
    
    Set objSQL = Nothing
    
End Function

Friend Function GEMS_CreateSections(ByVal PatientUnitNumber As String, _
                                    ByVal FacilityNCID As String, _
                                    ByVal EncounterNumber As String, _
                                          objDAS As ICHCSII_DAS) As Boolean
                                          
On Error GoTo ErrHnd

'-- use the same objdas to speed up the retrieves.
    Dim colData As Collection
    Dim objRS As ADODB.Recordset
    Dim sSQL1 As String
    Dim oSQLEx As ICHCSII_SQL.ISqlOps
    
    Set oSQLEx = gobjCHCSIIConn.CHCSII_SQLOPS(Auto) '''<SCR 36493
    
    sSQL1 = "Select * from Enc_Sections where encounternumber = " & EncounterNumber _
          & " and facilityNCID = " & FacilityNCID & " and enc_sectionsIndex = "
          
    Set objRS = objDAS.OpenRecordset(sSQL1 & SectionEnumToNCID(eAutoCite))
    GEMS_CreateSection eAutoCite, objRS
    objRS.Close
    
    'Phillip Crowder 8/20/03  Performance Improvements PGUI
    If Not mobjShared.IsAppMode(modeCHCSI_GUI) Then
      Set objRS = objDAS.OpenRecordset(sSQL1 & SectionEnumToNCID(eRFV))
      GEMS_CreateSection eRFV, objRS
      objRS.Close
    
      Set objRS = objDAS.OpenRecordset(sSQL1 & SectionEnumToNCID(eVitals))
      GEMS_CreateSection eVitals, objRS
      objRS.Close
    End If
    
    Set objRS = objDAS.OpenRecordset(sSQL1 & SectionEnumToNCID(eSO))
    GEMS_CreateSection eSO, objRS
    objRS.Close
    
'    Moved to below these sections are stores as notes 'SCR-737, 739
'    Set objRS = objDAS.ExecuteSQL(sSQL1 & SectionEnumToNCID(eLabs))
'    GEMS_CreateSection eLabs, objRS
'    objRS.Close
'
'    Set objRS = objDAS.ExecuteSQL(sSQL1 & SectionEnumToNCID(eRads))
'    GEMS_CreateSection eRads, objRS
'    objRS.Close
'
'    Set objRS = objDAS.ExecuteSQL(sSQL1 & SectionEnumToNCID(eQuestionnaire))
'    GEMS_CreateSection eQuestionnaire, objRS
'    objRS.Close
    
    Set objRS = objDAS.OpenRecordset(sSQL1 & SectionEnumToNCID(eap))
    GEMS_CreateSection eap, objRS
    objRS.Close
    
    If EncounterType = Dental Then
        Set objRS = objDAS.OpenRecordset(sSQL1 & SectionEnumToNCID(eDental))
        GEMS_CreateSection eDental, objRS
        objRS.Close
    End If
    
    Set objRS = objDAS.OpenRecordset(sSQL1 & SectionEnumToNCID(eDisp))
    GEMS_CreateSection eDisp, objRS
    objRS.Close
    
    Set objRS = objDAS.OpenRecordset(sSQL1 & SectionEnumToNCID(eApptNotes)) 'SCR-26496
    GEMS_LoadApptNotes objRS
    objRS.Close
    
    sSQL1 = "Select * from Enc_rtfs where encounternumber = " & EncounterNumber _
      & " and facilityNCID = " & FacilityNCID & " and rtftype = "

    Set objRS = objDAS.OpenRecordset(sSQL1 & "'NOTE'")
    GEMS_CreateSection eNotes, objRS
    objRS.Close
    
    Set objRS = objDAS.OpenRecordset(sSQL1 & "'TextNote'")
    GEMS_CreateSection eHistory, objRS
    objRS.Close
    
    'Added for labs and rads 'SCR-737, 739
    Set objRS = objDAS.OpenRecordset(sSQL1 & "'LABNOTE'")
    GEMS_CreateSection eLabs, objRS
    objRS.Close
    
    Set objRS = objDAS.OpenRecordset(sSQL1 & "'RADNOTE'")
    GEMS_CreateSection eRads, objRS
    objRS.Close
    
    Set objRS = objDAS.OpenRecordset(sSQL1 & "'QNRNOTE'")
    GEMS_CreateSection eQuestionnaire, objRS
    objRS.Close
    
    Set objRS = Nothing
    
Exit Function

ErrHnd:
    MsgBxARMd "Error in Gems_CreateSections: " & Err.Description & ": " & Err.Number

End Function

Friend Function GEMS_CreateSection(ByVal eSection As EncounterSectionEnum, _
                                         objRS As ADODB.Recordset) As Boolean
    Dim objSecParent As SectionParent
    
    On Error Resume Next
    Set objSecParent = mcolSecParents(CStr(eSection))
    If objSecParent Is Nothing Then
        Set objSecParent = New SectionParent
        Call objSecParent.GEMS_Load(Me, eSection, objRS)
        mcolSecParents.Add objSecParent, CStr(eSection)
    Else
        ' we get here from Refresh->Load->CreateSections->CreateSection
        Call objSecParent.GEMS_Load(Me, eSection, objRS)
    End If
    
    'clean up
    If objSecParent Is Nothing Then
      Set objSecParent = Nothing
    End If

End Function
Public Function GEMS_LoadSection(ByVal eSection As EncounterSectionEnum) As Boolean

    Dim objRS As ADODB.Recordset
    Dim sSQL1 As String
    Dim objSecParent As SectionParent
    Dim objDAS As ICHCSII_DAS
    
    Set objDAS = gobjCHCSIIConn.CHCSII_DAS(Auto) '''<SCR 36493
    
    sSQL1 = "Select * from Enc_Sections where encounternumber = " & EncounterID _
          & " and facilityNCID = " & FacilityNCID & " and enc_sectionsIndex = "
          
    Set objRS = objDAS.ExecuteSQL(sSQL1 & SectionEnumToNCID(eSection))
    
    If mcolSecParents Is Nothing Then
        Refresh 'refresh to reload sections 'SCR#47049
    End If
    
    Set objSecParent = mcolSecParents(CStr(eSection))
    
    If objSecParent Is Nothing Then
        Refresh 'refresh to reload sections 'SCR#47049
        Set objSecParent = mcolSecParents(CStr(eSection))
        If objSecParent Is Nothing Then
            MsgBxARMd "Unable to retrieve section data for Encounter:" & mnEncounterID, vbCritical, "Encounter - Add Text Note"
            Exit Function
        End If
    End If
    
    Call objSecParent.GEMS_Load(Me, eSection, objRS)
    
    objRS.Close
    
    Set objRS = Nothing
    Set objDAS = Nothing
    
End Function





Public Function GEMS_Save(Optional ByVal bForce As Boolean = False) As Boolean
    Dim bResult          As Boolean
    Dim i                As Integer
    Dim objSectionParent As SectionParent
    Dim objEncRTF        As EncRTF
    Dim objDAS           As ICHCSII_DAS
    Dim sql              As String
    Dim sql2             As String
    Dim sTemp            As String
    Dim oSQLEx As ICHCSII_SQL.ISqlOpsEx
    
    On Error GoTo ErrHandler
    
    If Not (mbDirty Or mbFlagNew Or bForce Or AccidentData.IsDirty Or PregnancyData.IsDirty Or AssociatedProviders.IsDirty) Then
      GEMS_Save = True
      Exit Function
    End If
    
    #If debugon Then
        Dim DR As DebugRec
        If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Save", "Begin", DR, "Enc", mnEncounterID
    #End If
    
    If Not mobjLocking.StartSectionUpdate(eHeader, 0) Then
        MsgBxARMd "Your lock on this encounter was broken by another user, and your changes cannot be saved.", vbInformation, "Encounter"
        GoTo CleanUp
    End If
    
    Set objDAS = gobjCHCSIIConn.CHCSII_DAS(Auto) '''<SCR 36493
    Set oSQLEx = gobjCHCSIIConn.CHCSII_SQLOPS(Auto)
    
    oSQLEx.Execute "select * from encounters where encounternumber = " & mnEncounterID
    
    If Not oSQLEx.EOF Then
        oSQLEx("UNIT_NUMBER") = mnPatientID
        oSQLEx("FACILITYNCID") = msFacilityNCID
        oSQLEx("ENCOUNTERNUMBER") = mnEncounterID
        oSQLEx("CLINICNCID") = msClinicNCID
        oSQLEx("STATUS") = StatusEnumToNCID(meStatus)
        oSQLEx("CATEGORY") = msCategory
        oSQLEx("TYPE") = meType
        ' not used SCR-25084
        'oSQLEx("ENDDTS") = mdEndDTS
        oSQLEx("APPTID") = msApptID
        
        oSQLEx("STARTDTS") = mdStartDTS
        '141844
        oSQLEx("STARTDTS_GMT") = mobjComm.LocalToGmt(mdStartDTS)
        
        oSQLEx("SADRSTATUS") = msSADRStatus
        oSQLEx.AppendAsChunk "TEMPLATEDATA", msTemplateData
        oSQLEx("TIMEZONE") = msTimeZone
        oSQLEx("UPDATELOCK") = CInt(mblnUpdateLock)
        
        oSQLEx("APPTDTS") = mdApptDTS
        '141844
        oSQLEx("APPTDTS_GMT") = mobjComm.LocalToGmt(mdApptDTS)
        
        'SCR 55063
        oSQLEx("Sensitive") = mlSensitivity
        
        If Trim$(msPrimaryDiagnosisNCID) <> vbNullString Then
            oSQLEx("PRIMARYDIAGNOSISNCID") = msPrimaryDiagnosisNCID
        End If
        
        ' Save Medcin ID & Prefix SCR-19435
        oSQLEx("DXSnoID") = mnPrimaryDxSnoID
        
        If Trim$(msPrimaryDxMedcinPrefix) <> vbNullString Then
            oSQLEx("DXMEDCINPREFIX") = msPrimaryDxMedcinPrefix
        End If

        If mcolProviders.Count > 0 Then
            oSQLEx("PRIMARYPROVIDERNCID") = mcolProviders(1).NCID
            If mcolProviders.Count > 1 Then
                oSQLEx("ADDITIONALPROVIDER1NCID") = mcolProviders(2).NCID
                oSQLEx("APROLE1NCID") = mcolProviders(2).RoleNCID
                If mcolProviders.Count > 2 Then
                    oSQLEx("ADDITIONALPROVIDER2NCID") = mcolProviders(3).NCID
                    oSQLEx("APROLE2NCID") = mcolProviders(3).RoleNCID
                Else  '--- SCR 50359   Sherry Wang   1/19/2005
                    oSQLEx("ADDITIONALPROVIDER2NCID") = Null
                    oSQLEx("APROLE2NCID") = Null
                End If
            Else '--- SCR 50359   Sherry Wang   1/19/2005
                oSQLEx("ADDITIONALPROVIDER1NCID") = Null
                oSQLEx("APROLE1NCID") = Null
                oSQLEx("ADDITIONALPROVIDER2NCID") = Null
                oSQLEx("APROLE2NCID") = Null
            End If
    
            If Not mobjCosigner Is Nothing Then
                oSQLEx("COSIGNERNCID") = mobjCosigner.NCID
                    'SCR 81090 - Jane Shen 1/11/2006.
                    ' If whenCosigned is NULL, not set default value into database
                    If CDbl(mdWhenCosigned) > 0 Then
                        oSQLEx("WHENCOSIGNED") = mdWhenCosigned
                        '141844
                        oSQLEx("WHENCOSIGNED_GMT") = mobjComm.LocalToGmt(mdWhenCosigned)
                    End If
                    If Not mobjWhoCosigned Is Nothing Then
                        If mobjWhoCosigned.NCID <> 0 Then
                            oSQLEx("WHOCOSIGNEDNCID") = mobjWhoCosigned.NCID
                            oSQLEx("COSIGNERLINE2") = mobjWhoCosigned.SigLine2
                            oSQLEx("COSIGNERLINE3") = mobjWhoCosigned.SigLine3
                        End If
                    End If
            Else
                'Need to clear NCID in case we are amending
                'oSQLEx("COSIGNERNCID") = vbEmpty
            End If
    
            If Not mobjWhoSigned Is Nothing Then
                If mdWhenSigned <> 0 Then  'SCR-21188
                    oSQLEx("WHENSIGNED") = mdWhenSigned  ' gobjCHCSIIConn.sqldate(mdWhenSigned)
                    '141844s
                    oSQLEx("WHENSIGNED_GMT") = mobjComm.LocalToGmt(mdWhenSigned)
                End If
                oSQLEx("WHOSIGNEDNCID") = mobjWhoSigned.NCID
                oSQLEx("SIGNERLINE2") = mobjWhoSigned.SigLine2
                oSQLEx("SIGNERLINE3") = mobjWhoSigned.SigLine3
            Else
                'Need to clear NCID in case we are amending
                'oSQLEx("WHOSIGNEDNCID") = Null
            End If
            
            If Not Val(msAllergyVerifiedByNCID) = 0 Then
                oSQLEx("ALLERGYVERIFIEDBYNCID") = msAllergyVerifiedByNCID   'WRogers removed the ticks, field is number not character string
            End If
            oSQLEx("ALLERGYVERIFIEDBYDATE") = mdAllergyVerifiedByDate  ' gobjCHCSIIConn.sqldate(mdAllergyVerifiedByDate)
        Else
            MsgBxARMd "Primary Provider data missing in GEMS_Save for Encounter Number: " & mnEncounterID, vbCritical, "EncounterOps"
            GoTo CleanUp
        End If
        
        oSQLEx("INPATIENT_SERVICE") = msInpatientService
      
        '****************************************
        'SCR 64993
        'Reset the TMIP_STATUS.
        oSQLEx("TMIP_STATUS") = Null
        '****************************************
      
        oSQLEx.Update
        
        GEMS_Save = SaveHIPAAData(False)

    End If
    
    If EandMReviewed And EandMReviewDate <> NULL_DATE And Val(EandMReviewedBy) > 0 Then
        If mobjLocking.LockSection(eDisp, 0, SectionLock) Then
            Set objSectionParent = mcolSecParents(CStr(eDisp))
            Call objSectionParent.OpenEncounter(Me)
            If Not mcolSecParents(CStr(eDisp)).mobjSection.Save Then
                Err.Raise vbObjectError, "Encounter.Save", "Error saving E & M Coding Information"
            End If
            Call mobjLocking.UnlockSection(eDisp, 0)
        End If
    End If
    
    Dim oSQLOps   As ICHCSII_SQL.ISqlOps
    
    Set oSQLOps = gobjCHCSIIConn.CHCSII_SQLOPS(Auto) '''<SCR 36493
    
    Dim rtf_index As Long
    Dim objNote   As TextNote
    
    objDAS.ExecuteSQL "Delete from enc_rtfs where " _
                    & " ENCOUNTERNUMBER = " & mnEncounterID _
                    & "  and FACILITYNCID = " & msFacilityNCID _
                    & " and rtfType = 'RTF'"
                    
    For Each objEncRTF In mcolEncRTFs
        rtf_index = oSQLOps.GetNextID()
        sql2 = ""
        sql2 = "insert into enc_RTFs (FACILITYNCID,ENCOUNTERNUMBER,rtftype,DTS,ENC_RTFSINDEX ) "
        sql2 = sql2 & " values ( " & msFacilityNCID & ", " & mnEncounterID & ", "
'        sql2 = sql2 & " 'RTF', " & gobjCHCSIIConn.sqldate(objEncRTF.mdDTS) & "," & rtf_index & " )"
        'sql2 = sql2 & " 'RTF', '" & Format$(objEncRTF.mdDTS, DATE_FORMAT_VB) & "'," & rtf_index & " )"
        If mobjShared.UseSQLServer Then 'ITT Change
            sql2 = sql2 & " 'RTF', " & gobjCHCSIIConn.SQLDate(objEncRTF.mdDTS) & "," & rtf_index & " )"
        Else
            'Format does not work when transferring LDDB->CDR SCR-22939
            sql2 = sql2 & " 'RTF', " & gobjCHCSIIConn.SQLDate(objEncRTF.mdDTS) & "," & rtf_index & " )"
            'Removed single quotes; SCR 25196
            'sql2 = sql2 & " 'RTF', '" & gobjCHCSIIConn.sqldate(objEncRTF.mdDTS) & "'," & rtf_index & " )"
        End If
        
        Call objDAS.ExecuteSQL(sql2)
        
        Set objNote = New TextNote
        objNote.mnDataID = rtf_index
        objNote.msOwnerName = mobjLogon.UserName
        objNote.msOwnerNCID = objEncRTF.msUserNCID
        objNote.mdDTS = objEncRTF.mdDTS
        objNote.msRTF = objEncRTF.msRTF
        
        objNote.GEMS_WriteToDataString
        Set objNote = Nothing
        
    Next objEncRTF
    
    'RaiseEvent when Enc Status is saved SCR-21322
    'Do not raise event for Dental encounters SCR-23470, 24189
    If meType <> Dental Then
        RaiseEvent EncStatusChanged(meStatus)
    End If
    
    Set oSQLOps = Nothing
    GEMS_Save = True
    mbDirty = False
    
    GoTo CleanUp
ErrHandler:
   Call mobjShared.ShowVBError(Err.Number, Err.Description, "EnctrGems.Save ", "Enc", vbCritical)
CleanUp:
    Call mobjLocking.EndSectionUpdate(eHeader, 0)
    
    #If debugon Then
        If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Save", "End", DR
    #End If
    
   
    If GEMS_Save Then
        mobjComm.SetStatusText "Encounter " & mnEncounterID & " was saved."
    Else
        mobjComm.SetStatusText ""
    End If
    
    Set objDAS = Nothing
    
End Function

Public Property Get DiscussedComments() As String
    DiscussedComments = mcolSecParents(CStr(eDisp)).mobjSection.DiscussedComments
End Property

Public Property Let DiscussedComments(ByVal RHV As String)
    mcolSecParents(CStr(eDisp)).mobjSection.DiscussedComments = RHV
End Property


Private Property Get IEncDoc_EncounterID() As Long
    IEncDoc_EncounterID = mnEncounterID
End Property

Private Property Let IEncDoc_EncounterType(ByVal RHS As ENCOUNTER_INTERFACES.EncType)

   EncounterType = RHS
   
End Property

Private Property Get IEncDoc_EncounterType() As ENCOUNTER_INTERFACES.EncType

    IEncDoc_EncounterType = EncounterType()

End Property

Private Property Get IEncDoc_EndDTS() As Date
    
    IEncDoc_EndDTS = mdEndDTS

End Property

Private Property Get IEncDoc_FacilityNCID() As String

    IEncDoc_FacilityNCID = msFacilityNCID

End Property

Private Property Get IEncDoc_SectionRTF(ByVal eSection As EncounterSectionEnum) As String

    IEncDoc_SectionRTF = SectionRTF(eSection)

End Property

Private Property Get IEncDoc_StartDTS() As Date

    IEncDoc_StartDTS = mdStartDTS

End Property

Private Property Let IEncDoc_Status(ByVal RHS As EncStatus)

    Status = RHS

End Property

Private Property Get IEncDoc_Status() As EncStatus

    IEncDoc_Status = Status

End Property

'Only return SO RTF SCR-28020
Public Property Get SORTFOnly() As String

    SORTFOnly = mcolSecParents(CStr(eSO)).RTF

End Property
'Save Appt Notes to display on Encounter document SCR-26496
Public Function SaveApptNotes(ByVal sApptRFV As String, ByVal sApptNotes As String)

    Dim sRTF        As String
    Dim objData     As DataString
    Dim objEvent    As ClinicalEvent
    Dim objLTR      As ListToolRecord
    Dim objSQL      As ICHCSII_SQL.ISqlOpsEx
    Dim strSQL      As String
    Dim sVitals     As String
    Dim sApptType   As String
    Dim sRFVNotes   As String
    
    On Error GoTo ErrHandler
    
    'Create String
    sRFVNotes = sApptRFV & "|*|" & sApptNotes
    msApptRFV = sApptRFV
    msApptNote = sApptNotes
    
    Set objSQL = gobjCHCSIIConn.CHCSII_SQLOPS(Auto) '''<SCR 36493
    
    strSQL = "SELECT DataID, FacilityNCID, EncounterNumber, Enc_SectionsIndex, Doc, OwnerName, OwnerNCID, OriginalSize " & _
         "FROM Enc_Sections " & _
         "WHERE EncounterNumber = " & Str(Me.EncounterID) & " " & _
         "AND Enc_SectionsIndex = " & Str(SectionEnumToNCID(eApptNotes)) & " " & _
         "AND FacilityNCID = " & Str(Me.FacilityNCID)
        
    objSQL.Execute strSQL
    
    If (objSQL.BOF And objSQL.EOF) Then
        objSQL.AddNew
        mnDataID = objSQL.GetNextID(exMAINSEQ)
        objSQL("DataID") = mnDataID
        objSQL("FacilityNCID") = Me.FacilityNCID
        objSQL("EncounterNumber") = Me.EncounterID
        objSQL("OwnerName") = Me.Logon.UserName
        objSQL("OwnerNCID") = Me.Logon.UserNCID
        objSQL("enc_sectionsindex") = SectionEnumToNCID(eApptNotes)
    End If
    
    Dim objZlib As EncZLib
    Dim aBytes() As Byte
    Set objZlib = New EncZLib
    
    If Len(sRFVNotes) <> 0 Then
        aBytes = StrConv(Replace(sRFVNotes, Chr(0), ""), vbFromUnicode)
        objZlib.CompressData aBytes
    End If
    objSQL.value("Doc") = aBytes
    
    objSQL("OriginalSize") = Len(sRFVNotes)
    objSQL.Update
    
    Exit Function
    
ErrHandler:
    
End Function
'Save Appt Notes to display on Encounter Document SCR-26496
Public Function LoadApptNotes(ByRef colDataStrings As Collection) As Boolean
    Dim objData As DATASTRATLLib.DataString
    Dim objEvent As CLINTYPEATLLib.ClinicalEvent
    Dim objClinObs As CLINTYPEATLLib.ClinicalObservation
    Dim i As Long
    Dim j As Long
    Dim sTemp As String
    Dim sbuf() As String
    
    If mobjShared.IsAppMode(modeCHCSI_GUI) Then '<SCR SCR 36493
        Exit Function
    End If
 
    mnApptNotesDataID = 0
    'Using the For...Each method with 3M objects causes memory leak
    For i = 1 To colDataStrings.Count
        Set objData = colDataStrings(i)
        If objData.DataTypeNCID = NCID_ENCEVENT Then
            Set objEvent = objData.PatientData
            If objEvent.Header.ObsBatId.value.Coded.ValueConcept.UniqueId = SectionEnumToNCID(eApptNotes) Then
                mnApptNotesDataID = objData.DataId
                For j = 1 To objEvent.ClinObs.Count
                    Set objClinObs = objEvent.ClinObs.Item(j)
                    Select Case objClinObs.ObsId.value.CodedWOSform.ValueConcept.UniqueId
                    Case NCID_ENCOUNTERRTF
                        sTemp = ExtractRTF(objClinObs)
                    End Select
                Next j

            End If
        End If
    Next 'objData
    
    sbuf = Split(sTemp, "|*|")
    
    If UBound(sbuf) = 1 Then
        msApptRFV = sbuf(0)
        msApptNote = sbuf(1)
    End If
End Function
Public Function GEMS_LoadApptNotes(objRS As ADODB.Recordset) As Boolean

    Dim sTemp As String
    Dim sbuf() As String

    If Not objRS.EOF Then
        
        'Decompress if compressed SCR-24683
        'If Original Size = 0 then dont load any RTF SCR-854 ITT
        If CVar(objRS("OriginalSize")) > 0 Then
            Dim objZlib As EncZLib
            Set objZlib = New EncZLib
            Dim aBytes() As Byte
            
            aBytes = objRS("Doc")
            Call objZlib.DecompressData(aBytes, objRS("OriginalSize"))
            sTemp = StrConv(aBytes, vbUnicode)
        End If
    End If

    sbuf = Split(sTemp, "|*|")
    
    If UBound(sbuf) = 1 Then
        msApptRFV = sbuf(0)
        msApptNote = sbuf(1)
    End If
    
    Exit Function
    
ErrHandler:

End Function
Public Property Get ApptNoteComment() As String

'    ApptNoteComment = "{\rtf1\ansi\ucl\deff0\deflang1033\deflangfe1033" _
'        & "{\fonttbl{\f0\fswiss\fcharset0\fprq2 Arial;}}" _
'        & "{\colortbl;\red0\green0\blue0;}" _
'        & "\pard\plain\f0\fs16 " & ConvertTextToRTF(msApptNote) & "\par " _
'        & "}"
     ApptNoteComment = ConvertTextToRTF(msApptNote)
End Property

' Added space after \fs16 for msApptRFV SCR-33748
' Suppress Appt Comments if none entered SCR-37551
' Concatenate Reason for Appt SCR-37552
Public Property Get ApptNotesRTF() As String
    
    Dim sApptType As String
    
    If meType = TeleConsult Then
        sApptType = "Reason for Telephone Consult:"
    Else
        sApptType = "Reason for Appointment:"
    End If
    
    If mobjComm.CmdLineSwitch("GREEN") <> "" Then
        msApptNotesRTF = "{\rtf1\ansi\ucl\deff0\deflang1033\deflangfe1033" _
            & "{\fonttbl{\f0\fswiss\fcharset0\fprq2 Arial;}}" _
            & "{\colortbl;\red0\green0\blue0;}"
            
        'Add Reason For Visit
        msApptNotesRTF = msApptNotesRTF & "\pard\plain\f0\fs16\b\ul " & sApptType & " \b0 " & "\pard \plain \f0 \fs16  " & msApptRFV & "\par "

        
        'Suppress Appt Comments if none entered
        If Len(msApptNote) Then
            'Add Appt Comments
            msApptNotesRTF = msApptNotesRTF & "\pard\plain\f0\fs16\b\ul " & "Appointment Comments:" & " \b0 " & "\par \pard"
            If Len(msApptNote) > 0 Then
                msApptNotesRTF = msApptNotesRTF & "\plain\f0\fs16\b0\ul0 " & ConvertTextToRTF(msApptNote) & " \par"
            Else
                msApptNotesRTF = msApptNotesRTF & "\plain\f0\fs16\b0\ul0 " & " "
            End If
        End If
        
        msApptNotesRTF = msApptNotesRTF & "\par }"
    Else
        msApptNotesRTF = "{\rtf1\ansi\ucl\deff0\deflang1033\deflangfe1033" _
            & "{\fonttbl{\f0\fswiss\fcharset0\fprq2 Arial;}}" _
            & "{\colortbl;\red0\green0\blue0;}" _
            & "\pard\plain\f0\fs16\b\ul " & sApptType & " \b0 " & "\par " _
            & "\pard\plain\f0\fs16 " & msApptRFV & "\par " _
            & "\pard\plain\f0\fs16\b\ul " & "Appointment Comments: " & " \b0 " & "\par " _
            & "\pard\plain\f0\fs16 " & ConvertTextToRTF(msApptNote) & "\par " _
            & "}"
    End If

    ApptNotesRTF = msApptNotesRTF

End Property
'If current Primary Provider is .Unassigned and current user has sign privileges then
'transfer ownership to current user.  Called by GEMS_Save in each Encounter Child section in EncounterOps project
' SCRs 29117, SCR-24349
Public Sub CheckEncounterOwnership()
    Dim oEncProvider As EncProvider
    Dim objAppts As Object
    Dim AppointmentOps As Object
    
    If CWShared.IsAppMode(modeTheater) Then
        'Check if owner is .Unassigned and user has sign privilege
        If InStr(1, UCase(PrimaryProvider.FullName), "UNASSIGNED", vbTextCompare) And mobjUser.HasPrivilegeEx(Priv_Current_Encounter, Priv_Sign) Then
            'Transfer ownership to current user
            'Need to remove current provider if on Provider list as Addl Provider
            'Skip if no Addl Providers exist SCR#54266
            If Providers.Count > 1 Then
                On Error Resume Next
                Providers.Remove (Logon.UserNCID)
                On Error GoTo ErrHandler
            End If
            
            'Change Primary Provider in Encounter
            Set oEncProvider = Providers(1)
            oEncProvider.FullName = Logon.UserName
            oEncProvider.NCID = Logon.UserNCID
            
            Call mobjLocking.LockSection(eHeader, 0, SectionLock)
'            Call mobjLocking.StartSectionUpdate(eHeader, 0)
            If Not Save(True) Then
                MsgBxARMd "Could not update encounter database." & vbCrLf & Err.Description, vbCritical, "AppointmentOps"
                Exit Sub
            End If
            Call mobjLocking.UnlockSection(eHeader, 0)
            'Change Primary Provider in Appointments
            'Set AppointmentOps = CreateObject("CHCSII_AppointmentClient.AppointmentOps")
            'AppointmentOps.Initialize NCID_ENCOUNTER_APPLICATION, Comm, Logon, Config, Patient, User
            ' Use Comm.InitializeOLEServer to obtain Appointment object SCR#41181
            Set objAppts = mobjComm.InitializeOLEServer(NCID_CLINIC_SCHEDULE_APPLICATION)
            
            If Not objAppts.AppointmentOps.UpdateAppointmentProvider(AppointmentID, Logon.UserNCID) Then
                Exit Sub
            End If
            
            'PASS APPTID|xxxxx 'SCR#49319
            mobjComm.Message cwiREFRESH_DATA, "APPTID|" & AppointmentID, NCID_CLINIC_SCHEDULE_APPLICATION, NCID_ENCOUNTER_APPLICATION
                        
            'Update provider name in encouter document
            RefreshDisplay
        End If
    End If
    
    Exit Sub
    
ErrHandler:

End Sub

Public Property Let ApptIEN(ByVal Val As String)
    msApptIEN = Val
End Property

Public Property Get ApptType() As String
'SCR 44501 - Added to support RNDS and HIPAA

    Dim objAppts As Object
    Dim objAppt As Object
    
    On Error GoTo ErrHandler
    
    If Len(msApptType) = 0 Then
        'Try to get Appt Type from Appt object
        Set objAppts = Comm.InitializeOLEServer(NCID_CLINIC_SCHEDULE_APPLICATION)
        Set objAppt = objAppts.AppointmentOps.GetAppointment(msApptID, True) 'SCR#52439
        
        If Not objAppt Is Nothing Then
           msApptType = objAppt.ApptType
        Else
            msApptType = vbNullString
        End If
        
        Set objAppts = Nothing
        Set objAppt = Nothing
    
    End If
    
    ApptType = msApptType
    
    Exit Property
    
ErrHandler:
    
    mobjShared.ShowVBError Err.Number, Err.Description, "Encounter.ApptIEN (Get)", "Encounters", vbCritical
       
    
End Property

Public Property Get ApptClass() As Long

    Dim objAppts As Object
    Dim objAppt As Object
    
    On Error GoTo ErrHandler
    
    If meApptClass - 1 Then
       'Try to get Appt Type from Appt object
       Set objAppts = Comm.InitializeOLEServer(NCID_CLINIC_SCHEDULE_APPLICATION)
       Set objAppt = objAppts.AppointmentOps.GetAppointment(msApptID, True) 'SCR#52439
        
       If Not objAppt Is Nothing Then
          meApptClass = objAppt.ApptClassEnum
       Else
          meApptClass = -1
       End If
       
       Set objAppts = Nothing
       Set objAppt = Nothing
       
    End If
    
    ApptClass = meApptClass
    
    Exit Property
    
ErrHandler:
    
    mobjShared.ShowVBError Err.Number, Err.Description, "Encounter.ApptIEN (Get)", "Encounters", vbCritical
       
    
End Property

Public Property Get ApptIEN() As String

    Dim objAppts As Object
    Dim objAppt As Object
    
    On Error GoTo ErrHandler
    
    If Len(msApptIEN) = 0 Then
        'Try to get Appt IEN from Appt object SCR-31315
        Set objAppts = Comm.InitializeOLEServer(NCID_CLINIC_SCHEDULE_APPLICATION)
        Set objAppt = objAppts.AppointmentOps.GetAppointment(msApptID, True) 'SCR#52439
        
        If Not objAppt Is Nothing Then 'SCR-42397, 40952
            If Len(objAppt.ApptIEN) Then
                msApptIEN = objAppt.ApptIEN
            Else
                msApptIEN = vbNullString
                'Need to know what to do in this scenario SCR-31315
            End If
        Else
            msApptIEN = vbNullString
        End If
        
        Set objAppts = Nothing
        Set objAppt = Nothing
    End If
    
    ApptIEN = msApptIEN
    
    Exit Property
    
ErrHandler:
    
    mobjShared.ShowVBError Err.Number, Err.Description, "Encounter.ApptIEN (Get)", "Encounters", vbCritical
    
End Property
'Public Function AddCouplerData(ByVal oCol As Collection, ByVal sRTF As String, ByVal bNewCouplerData As Boolean) As Boolean
Public Function AddCouplerData(ByVal colLTR As Collection, ByVal sRTF As String, ByVal bNewCouplerData As Boolean) As Boolean

    Dim objSO As SO
    Dim oLTR As ListToolRecord
    'Dim oConvert As New actxSearchMedcin.IConvert
    'Dim sRTF As String
    Dim objSONote As SONote
    Dim i As Long
    Dim bFound As Boolean
    Dim Note As SONote

    Set objSO = SectionParent(eSO).mobjSection
    
    For Each Note In objSO.Notes
        If Note.mbNewCouplerData Then
            Set objSONote = Note
            bFound = True
        End If
    Next
    
    If Not bFound Then
        Set objSONote = New SONote
    End If

    objSONote.mbNewCouplerData = bNewCouplerData
    objSONote.mbCouplerData = True
    objSONote.meType = ListNoteType
    objSONote.msRTF = sRTF
    objSONote.mdDTS = Now
    objSONote.msOwnerNCID = PrimaryProvider.NCID
    objSONote.msOwnerName = PrimaryProvider.FullName

    'Add SnoIDs
    Set objSONote.mcolListRecords = New Collection
    If Not colLTR Is Nothing Then
        For Each oLTR In colLTR
            objSONote.mcolListRecords.Add oLTR
        Next
    End If
'    Dim eOpenSection As Long
'    'We will now open SO module to display selected items
'    'First see if any modules are open
'    If Not mobjCurrSection Is Nothing Then
'        eOpenSection = mobjCurrSection.meSection
'        CloseSection
'    End If
'
'    'OpenSection eSO
'    'Dim objSO As SO
'    Set mobjCurrSection = mcolSecParents(CStr(eSO))
'    objSO.EditNote objSONote

    'Do not Lock section if new SO Note SCR-31457
    If bFound Then
        LockingObject.LockSection eSO, objSONote.DataId, SectionLock
    End If

    Call objSO.SaveNote(objSONote)

    'For new SO Note LockSection is called in SaveNote CleanUp code
    LockingObject.UnlockSection eSO, objSONote.DataId

    'Refresh Encounter document
    RefreshDisplay

'    If bNewCouplerData And Not bFound Then
'        If MsgBox("Would you like to include this coupler data into the note?", vbQuestion + vbYesNo, "Encounter") = vbNo Then
'            Call objSO.DeleteNote(0, objSONote.DataID)
'            RefreshDisplay
'        End If
'    End If
    
    Set objSONote = Nothing
    Set objSO = Nothing

End Function

Public Property Get GetCouplerData() As SONote
    
    Dim objSO As SO
    Dim objSONote As SONote
    
    Set objSO = SectionParent(eSO).mobjSection
    
    For Each objSONote In objSO.Notes
        If objSONote.mbNewCouplerData Then
            Set GetCouplerData = objSONote
            Exit Property
        End If
    Next

    'No Coupler data. Set up dummy data SCR-30929
    Set objSONote = New SONote
    Set objSONote.mcolListRecords = New Collection
    Set GetCouplerData = objSONote
    
    Set objSO = Nothing
    Set objSONote = Nothing
End Property


Public Function AppendNoteToSignedEncounter(ByRef Note As TextNote) As Boolean

    Dim sRTF As String
    Dim objNotes As TextNotes
    Dim objEncRTF As EncRTF
    
    On Error GoTo ErrHandler
    
    Set objNotes = SectionParent(eNotes).mobjSection
    If objNotes.Save(Note) Then

        Set objEncRTF = New EncRTF
        sRTF = BuildRTF(Me)
        If sRTF = "" Then
            Exit Function
        End If
        objEncRTF.mdDTS = Now
        objEncRTF.msRTF = sRTF
        objEncRTF.msUserNCID = Note.msOwnerNCID
        AddEncRTF objEncRTF
        Set objEncRTF = Nothing

        AppendNoteToSignedEncounter = Save
    End If
    
    Exit Function
    
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "Encounter.AppendNoteToSignedEncounter", App.Title, vbCritical
    
End Function
'Called by Appointments when cancelling an appointment
Public Sub ClearAddlProviders()

    Dim i As Integer

    For i = mcolProviders.Count To 2 Step -1
        mcolProviders.Remove (i)
    Next

End Sub
'Called in EncSignOps.SignEncounter SCR-35067
Public Function APAfterDisp() As Boolean

    On Error GoTo ErrHandler
    
    APAfterDisp = False
    
    'Check if both AP and Disposition have been completed
    If APStatus = complete And DispStatus = complete Then 'SCR#52811
        'Check if AP data changed after Disposition data entered
        If mcolSecParents(CStr(eap)).mobjSection.DTS > mcolSecParents(CStr(eDisp)).mobjSection.DTS Then
            APAfterDisp = True
        End If
    End If
    
    Exit Function
    
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "Encounter.APAfterDisp", App.Title, vbCritical
    
End Function

Public Function SavePGSONote() As Boolean
'<< shaw - need to save the s/o note text
    If mobjShared.IsAppMode(modeCHCSI_GUI) Then '<SCR SCR 36493
        If TextNotes.Count > 0 Then
            With TextNotes.TextNote(TextNotes.Count)
                If Not .ReadOnly Then
                
                    '__save the new stuff below the fold
                    Dim NewNote As String
                    
                    If TextNotes.Count > 1 Then
                        'Add space between notes
                        NewNote = "\par "
                    End If
                    
                    NewNote = NewNote & "{\rtf1\ansi\ucl\deff0\deflang1033\deflangfe1033" _
                            & "{\fonttbl{\f999\fswiss\fcharset0\fprq2 Arial;}}" _
                            & "{\colortbl;\red0\green0\blue0;}" _
                            & "\pard\plain\f999\fs16\b\ul Note \b0 Written by " _
                            & .msOwnerName & " @ " & FormatDTG(.mdDTS) & "\par\ul0\b " _
                            & .msRTF & "}"
                            
                    Call mobjShared.CHCSConnection.AppendSoNote(CStr(mnEncounterID), NewNote)
                    
                    If Not (mobjWAM Is Nothing) Then
                      With mobjWAM.Encounter.RTFs.Item(3)
                        .RTF = .RTF & NewNote
                      End With
                    End If
                End If
            End With
        End If
    End If
End Function

Public Sub DisplayDetailedError(ByVal Section As String, ByVal Msg As String)

    Dim sErrMsg As String
    
    sErrMsg = Msg & vbCrLf & Section
    
    sErrMsg = sErrMsg & vbCrLf & "(" & PatientID & " - " & FacilityNCID & " - " & EncounterID & ")"
    
    sErrMsg = sErrMsg & vbCrLf & FormatDTG(Now)

    MsgBxARMd sErrMsg, vbCritical, App.Title
    
End Sub
Private Function M_sSqlVal(ByVal sVal As String, ByVal bQuate As Boolean) As String
   Dim sResult As String
   If Len(sVal) = 0 Then
      sResult = "NULL"
   Else
      sResult = sVal
      If bQuate Then
         sResult = "'" & sResult & "'"
      End If
   End If
   M_sSqlVal = sResult
End Function

' Created to fix ADM Writeback error SCR-39893
' Call by SetCurrentEncounter when statuses are out of sync SCR#51097, 52460
Public Sub UpdateApptEncounterStatus()

    Dim sDestApp As String
    Dim oApptStat As ApptStatus
    
    On Error GoTo ErrHandler
    
    If Not mobjShared.IsAppMode(modeCHCSI_GUI) Then
        Set oApptStat = New ApptStatus
        If Not oApptStat.SetEncounterStatus(msApptID, giEncApptStatus(meStatus), (meClass <> ec_NoCount)) Then
            MsgBxARMd "Internal processing error in Encounter.Status(let). Unable to update appointment status.", vbOKOnly, "Encounter communication Error"
        End If
        Set oApptStat = Nothing
    End If
    
    #If debugon Then
        If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Let Status", "msg to appt screen to update", DR
    #End If
    
    '<< SCR 10812 refresh of appt on enc status changes
    '<< SCR 10916 Diferentiate between appt and telecon
    If meType = TeleConsult Then
        sDestApp = NCID_TELCON_APPLICATION
    Else
        sDestApp = NCID_CLINIC_SCHEDULE_APPLICATION
    End If
    
    If mobjComm.Message(cwiOLE_SERVER_RUNNING, sDestApp, NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION) Then
        '- tell it to update the single appt updates grid
        'mobjComm.Message cwiGENERIC, "UPDATE_ENCOUNTER_STATUS|" & msApptID & "|" & giEncApptStatus(meStatus), sDestApp, NCID_ENCOUNTER_APPLICATION
        
        'SCR 67667
        If Me.EncounterType <> InpatientNote Then
            'Use cwiRefresh_Data 'SCR#49319
            mobjComm.Message cwiREFRESH_DATA, "APPTID|" & msApptID, sDestApp, NCID_ENCOUNTER_APPLICATION
        End If
    End If
    '>>
    '>>
    Exit Sub
    
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "Encounter.UpdateApptEncounterStatus", App.Title, vbCritical
End Sub

Public Property Get TemplateID() As String
    TemplateID = msTemplateID
End Property

Public Property Let TemplateID(ByVal RHV As String)
    msTemplateID = RHV
End Property

'Used by Previous Encounters to get current RTF for Inpatient Notes SCR#42988
Public Function GetCurrentRTF() As String
    Refresh
    GetCurrentRTF = BuildRTF(Me)
End Function

Public Property Get AccidentData() As EncAccidentData
   Set AccidentData = mobjAccidentData
End Property

Public Property Get PregnancyData() As EncPregnancy
   Set PregnancyData = mobjPregnancyData
End Property

Public Property Get AdditionalProvider(ByVal lIndex As Long) As EncProvider
    If Not mcolProviders Is Nothing Then
        If mcolProviders.Count > 0 And (mcolProviders.Count >= (lIndex + 1)) Then
            Set AdditionalProvider = mcolProviders(lIndex + 1)
        End If
    End If
End Property


Friend Function GEMS_SaveProviders(ByVal sFacilityNCID As String, _
                               ByVal sEncounterNumber As String) As Boolean
                               
On Error GoTo ErrorHandler

Dim oProvider As EncProvider
Dim oSQL As ICHCSII_SQL.ISqlOpsEx
Dim sSql As String
Dim lIdx As Long

Dim lPriority As Long

   'Perform the deletes first to avoid constraint violations
   For lIdx = 1 To mcolProviders.Count
      Set oProvider = mcolProviders(lIdx)
      If oProvider.State = rsDelete Then
         'Just call method to save the providers - procedures will be deleted as part of the save
         GEMS_SaveProviders = oProvider.GEMS_SaveEncounterProvider(sEncounterNumber, sFacilityNCID, lPriority)
      End If
   Next
   
   'Now Process everything else
   For lIdx = 1 To mcolProviders.Count
      Set oProvider = mcolProviders(lIdx)
      If oProvider.State <> rsRemove Then lPriority = lPriority + 1
      If oProvider.State <> rsDelete Then
         'Save Provider
         If oProvider.GEMS_SaveEncounterProvider(sEncounterNumber, sFacilityNCID, lPriority) Then
            'Save the procedure associations
            GEMS_SaveProviders = oProvider.GEMS_SaveProviderProcedures(sEncounterNumber, sFacilityNCID)
         End If
      End If
   Next
   
   'Clean out the providers collection
   mcolProviders.RemoveProviders
   
ErrorHandler:
   Set oProvider = Nothing
   Set oSQL = Nothing
   
   If CBool(Err) Then HandleErrors Err, MODULE_NAME, "GEMS_SaveProviders"
   
End Function
Friend Function GEMS_LoadProviders(ByVal lFacilityNCID As String, _
                               ByVal lEncounterNumber As String) As Boolean
                               
'NOT USED IN THEATER MODE

On Error GoTo ErrorHandler

Dim sSql As String
Dim oSQL As ICHCSII_SQL.ISqlOps
Dim oProvider As EncProvider
Dim oProcedure As EncProviderProcedure
Dim cProviders As EncProviders

   'Initialize new collection
   Set cProviders = New EncProviders

   'Build the SQL - This recordset will be the combination of providers and procedures
   If mobjShared Is Nothing Then Set mobjShared = New CWShared
   If mobjShared.UseSQLServer Then
      'Build SQL using SQLServer Syntax
      sSql = "SELECT P.Name, EP.Provider_NCID, P.IEN as Provider_IEN, EP.Role_Type_NCID, PP.DataId as ProcedureDataId, " & _
            " PR.SnoID, RT.Role_Type_Description as Role_Desc, EP.Priority_Indicator " & _
            " FROM Encounter_Provider EP " & _
            " INNER JOIN Provider_Role_Type RT On EP.Role_Type_NCID = RT.Role_Type_NCID " & _
            " INNER JOIN Provider P ON EP.Provider_NCID = P.NCID " & _
            " LEFT OUTER JOIN Encounter_Procedure PP on EP.EncounterNumber = PP.EncounterNumber " & _
               " AND EP.FacilityNCID = PP.FacilityNCID AND EP.Provider_NCID = PP.Provider_NCID " & _
            " LEFT OUTER JOIN Procedures PR on PP.DataId = PR.DataId " & _
            " WHERE EP.EncounterNumber = " & lEncounterNumber & " AND EP.FacilityNCID = " & lFacilityNCID & _
            " ORDER BY EP.Priority_Indicator ASC"
   Else
      sSql = "SELECT P.Name, EP.Provider_NCID, P.IEN as Provider_IEN, EP.Role_Type_NCID, PP.DataId as ProcedureDataId, " & _
            " PR.SnoID, RT.Role_Type_Description as Role_Desc, EP.Priority_Indicator " & _
            " FROM Encounter_Provider EP, Encounter_Procedure PP, Provider_Role_Type RT, Provider P, Procedures PR " & _
            " WHERE EP.EncounterNumber = " & lEncounterNumber & " AND EP.FacilityNCID = " & lFacilityNCID & _
            " AND EP.EncounterNumber = PP.EncounterNumber(+) AND EP.FacilityNCID = PP.FacilityNCID(+) " & _
            " AND EP.Provider_NCID = PP.Provider_NCID(+) " & _
            " AND EP.Role_Type_NCID = RT.Role_Type_NCID " & _
            " AND EP.Provider_ncid = P.ncid " & _
            " AND PP.DataId = PR.DataId(+) " & _
            " ORDER BY Priority_Indicator ASC "
   End If
      
   Set oSQL = gobjCHCSIIConn.CHCSII_SQLOPS(Auto)
   oSQL.Execute sSql
    
   Do While Not oSQL.EOF
   
      'Create and load the provider
      Set oProvider = New EncProvider
      
      'Load the provider
      oProvider.PrevNCID = oSQL("Provider_NCID") & ""
      oProvider.NCID = oSQL("Provider_NCID") & ""
      oProvider.IEN = oSQL("Provider_IEN") & ""
      oProvider.FullName = oSQL("Name") & ""
      oProvider.RoleNCID = oSQL("Role_Type_NCID") & ""
      oProvider.Role = oSQL("Role_Desc") & ""
      oProvider.State = rsSaved
      
      'Create and load the procedure associations
      Do
         If Not IsNull(oSQL("ProcedureDataId")) Then
            'Add a procedure to the provider - make sure we have a valid data id since ITT mode a null DB value will be returned as a zero
            If Val(oSQL("ProcedureDataId")) > 0 Then
               Set oProcedure = New EncProviderProcedure
               oProcedure.Load oSQL("ProcedureDataId"), oSQL("SnoID"), oSQL("Provider_NCID")
               oProvider.Procedures.Add oProcedure
               Set oProcedure = Nothing
            End If
         End If
         oSQL.MoveNext
         If oSQL.EOF Then Exit Do
      Loop Until oSQL("Provider_NCID") & "" <> CStr(oProvider.NCID)
      
      'Add the provider to the collection
      cProviders.Add oProvider, CStr(oProvider.NCID)
            
      Set oProvider = Nothing
      
   Loop
      
   If cProviders.Count > 0 Then
      'For backwards compatibility - only update encounter collection if we found provider information in new structure
      Set mcolProviders = Nothing
      Set mcolProviders = cProviders
   End If
   
ErrorHandler:
   Set oSQL = Nothing
   Set oProvider = Nothing
   Set oProcedure = Nothing
   Set cProviders = Nothing
   
   If CBool(Err) Then HandleErrors Err, MODULE_NAME, "GEMS_LoadProviders"
           
End Function

Public Function SetProviderInfo(ByVal lPriority As Provider_Priority, ByVal sProviderNCID As String, _
   ByVal sRoleNCID As String, ByVal sProviderIEN As String, ByVal sProviderName As String) As EncProvider

On Error GoTo ErrorHandler

Dim lIdx As Long
Dim oNewProvider As EncProvider
Dim cProcedures As Collection

   If sProviderNCID = "0" Or sProviderNCID = vbNullString Then Exit Function
   
   'Check if the provider already exists in the collection
   If lPriority > mcolProviders.Count Then
      'Doesn't exist so create the provider and add to the collection
      Set oNewProvider = New EncProvider
      oNewProvider.State = rsNew
      mcolProviders.Add oNewProvider, CStr(sProviderNCID)
   Else
      Set oNewProvider = mcolProviders(lPriority)
      If oNewProvider.State = rsUnknown Then
         oNewProvider.State = rsNew
      Else
         If oNewProvider.NCID <> sProviderNCID Or oNewProvider.RoleNCID <> sRoleNCID Then
            oNewProvider.State = rsModified
         End If
      End If
   End If
   
   oNewProvider.NCID = sProviderNCID
   oNewProvider.RoleNCID = sRoleNCID
   oNewProvider.IEN = sProviderIEN
   oNewProvider.FullName = sProviderName
   
   Set SetProviderInfo = oNewProvider
   
ErrorHandler:
   If CBool(Err) Then HandleErrors Err, MODULE_NAME, "SetProviderInfo"
   
End Function

Public Sub ClearProvider(ByVal lPriority As Long)

Dim oProvider As EncProvider
Dim cProviders As Collection
Dim lIdx As Long

   If lPriority >= 1 And lPriority <= mcolProviders.Count Then
      Set oProvider = mcolProviders.Item(lPriority)
      If oProvider.State <> rsRemove Then
         oProvider.State = rsDelete
         oProvider.NCID = 0
         oProvider.FullName = ""
         oProvider.RoleNCID = 0
         With oProvider.Procedures
            For lIdx = 1 To .Count
               oProvider.UnAssociateProcedure .Item(lIdx).DataId, .Item(lIdx).SnoID
            Next
         End With
      End If
   End If
   
End Sub

'Transfer this encounter to another provider and re-prioritize additional providers
Public Function TransferToProvider(ByVal sProviderNCID As String, ByVal sProviderName As String) As Boolean

On Error GoTo ErrorHandler

Dim oProvider As EncProvider
Dim lIdx As Long
Dim lPosition As Long
   
   'Locate the specified provider if they are already associated with the encounter
   Set oProvider = mcolProviders.Find(sProviderNCID, lPosition)
   
   'Make sure the provider isn't already the primary provider
   If mcolProviders(1).NCID = sProviderNCID Then Exit Function
   
   'Remove the primary Provider
   mcolProviders.RemoveProvider mcolProviders(1).NCID
      
   If Not oProvider Is Nothing Then
                  
      'Move the provider to the primary provider in the collection
      mcolProviders.Remove CStr(oProvider.NCID)
      mcolProviders.Add oProvider, CStr(oProvider.NCID), 1
      
      'Update the state of each provider since the priority may have changed to ensure they will be updated during save
      For lIdx = 1 To mcolProviders.Count
         If mcolProviders(lIdx).State = rsSaved Then
            mcolProviders(lIdx).State = rsModified
         End If
      Next lIdx

   Else
      'Add the provider as the primary since it is not already associated
      Set oProvider = New EncProvider
      oProvider.NCID = sProviderNCID
      oProvider.FullName = sProviderName
      oProvider.RoleNCID = rtnAttendingProvider
      oProvider.State = rsNew
      mcolProviders.Add oProvider, CStr(oProvider.NCID), 1
   End If

   'Processed transfer ok
   TransferToProvider = True
   
ErrorHandler:
   Set oProvider = Nothing
   
   If CBool(Err) Then HandleErrors Err, MODULE_NAME, "TransferToProvider"
         
End Function

Public Function SaveHIPAAData(Optional ByVal bForceCreate As Boolean = False, _
   Optional ByRef oEvent As ClinicalEvent) As Boolean

On Error GoTo ErrorHandler:

Dim bSuccess            As Boolean        'Indicates whether the save was successful
Dim bRefreshAutoCites   As Boolean        'Indicates whether the autocites RTF needs to be refreshed (PGUI only)
Dim oProvider           As EncProvider

   'Initialize Shared if not already initialized
   If mobjShared Is Nothing Then Set mobjShared = New CWShared
   
   'Initialize the return value
   bSuccess = True
   
   'Only process HIPAA data when the switch is on
   If Len(mobjShared.CmdLineSwitches("HIPAA837")) > 0 Then
                  
      'Process Accident Data
      If Me.AccidentData.IsDirty And (mobjShared.IsAppMode(modeITT) Or mobjShared.IsAppMode(modeCHCSI_GUI)) Then
         bSuccess = bSuccess And Me.AccidentData.GemsSave(vbNullString, Me.FacilityNCID, Me.EncounterID)
      ElseIf Me.AccidentData.IsDirty And mobjShared.IsAppMode(modeCDR) Then
         bSuccess = bSuccess And Me.AccidentData.CDRSave(Me.Patient.UnitNumber, Me.FacilityNCID, Me.EncounterID)
      End If
                  
      'Process Pregnancy Data and update the visual pregnancy indicator if changed
      If (Me.PregnancyData.IsDirty Or (bForceCreate And Me.PregnancyData.PregnancyIndicator)) _
         And (mobjShared.IsAppMode(modeITT) Or mobjShared.IsAppMode(modeCDR)) Then
      
         'CDR and all ITT modes use the same db structures for pregnancy data
         bSuccess = bSuccess And Me.PregnancyData.CDRPlusSave(Me.FacilityNCID, Me.EncounterID, Me.Status)
         If Not Me.Comm Is Nothing Then Me.Comm.Message cwiRESET_VISUAL_INDICATORS, "", NCID_CW_APPLICATION, NCID_ENCOUNTER_APPLICATION
         
      ElseIf Me.PregnancyData.IsDirty And mobjShared.IsAppMode(modeCHCSI_GUI) Then
         'PGUI mode has a different db structure
         
      End If
      
      'Process Associated Providers
      If Me.AssociatedProviders.IsDirty And (mobjShared.IsAppMode(modeITT) Or mobjShared.IsAppMode(modeCHCSI_GUI)) Then
         bSuccess = bSuccess And GEMS_SaveProviders(Me.FacilityNCID, Me.EncounterID)
      ElseIf mobjShared.IsAppMode(modeCDR) And Not mobjShared.IsAppMode(modeITT) Then     'SCR 51520 - Always save in CDR mode otherwise the data will get blown away
         
         'Process each of the associated providers
         For Each oProvider In Me.AssociatedProviders
            bSuccess = bSuccess And oProvider.CDRSave(oEvent)
         Next
         
         'Clean out the providers collection
         mcolProviders.RemoveProviders
         
      End If
      
   End If
   
   SaveHIPAAData = bSuccess
   
ErrorHandler:
   'Clean Up Objects
   Set oProvider = Nothing
   
   If CBool(Err) Then HandleErrors Err, MODULE_NAME, "SaveHIPAAData"
   
End Function
