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
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--- begin DTO Builder; DTO=Appointment
'--- 3/17/2004 4:40:52 PM
'---
'--- NOTE: Any time you modify this class, increment the version number
'
'

Option Explicit

Implements iObjectPersist

Private Const OBJECT_VERSION = "1.0"
Private Const UNIQUE_IDENTIFIER = "Appointment"

Public Enum ApptSaveOp
    All
    CheckIn
    CheckOut
    Register
    Transfer
End Enum

Private mDateTime                   As Date
Private mCheckInTime                As Date
Private mCheckOutTime               As Date
Private mPriority                   As Integer
Private mClassification             As DTOs.ApptClassification
Private mADMEncounter               As String
Private mDetailCodes                As String
Private mDuration                   As String
Private mCHCSStatus                 As String
Private mCHCSType                   As String
Private mCancelReason               As String
Private mComment                    As String
Private mID                         As String
Private mIEN                        As String
Private mReason                     As String
Private mStatus                     As String
Private mType                       As String
Private mTypeCode                   As String
Private mClinicianNCID              As String
Private mClinicName                 As String
Private mClinicNCID                 As String
Private mEncounterNumber            As String
Private mEncounterStatus            As String
Private mFacilityNCID               As String
Private mInjuryIllnessCategory      As String
Private mInjuryIllnessCause         As String
Private mOrderNumber                As String
Private mPatient                    As DTOs.Patient
Private mProviderName               As String
Private mSponsorSSN                 As String
Private mUSVType                    As String
Private mWorkloadType               As String
Private mWorkloadTypeOption         As String
Private mPendingComplaints          As DTOs.ApptComplaints
Private mRFV                        As String  '  Requested for Encounter
Private mNote                       As String  '  Requested for Encounter
Private mNotesDataID                As String  '  Requested for Encounter
Private mWorkload                   As String  '  Requested for Encounter
Private mResultedConsultOrder       As Boolean
Private mArrivedByAmbulance         As Boolean
Private mRelatedToInpatient         As Boolean

Public Property Get ArrivedByAmbulance() As Boolean
    ArrivedByAmbulance = mArrivedByAmbulance
End Property
Public Property Let ArrivedByAmbulance(ByVal value As Boolean)
    mArrivedByAmbulance = value
End Property
Public Property Get RelatedToInpatient() As Boolean
    RelatedToInpatient = mRelatedToInpatient
End Property
Public Property Let RelatedToInpatient(ByVal value As Boolean)
    mRelatedToInpatient = value
End Property
Public Property Get DateTime() As Date
    DateTime = mDateTime
End Property
Public Property Let DateTime(ByVal value As Date)
    mDateTime = value
End Property
Public Property Get CheckInTime() As Date
    CheckInTime = mCheckInTime
End Property
Public Property Let CheckInTime(ByVal value As Date)
    mCheckInTime = value
End Property
Public Property Get CheckOutTime() As Date
    CheckOutTime = mCheckOutTime
End Property
Public Property Let CheckOutTime(ByVal value As Date)
    mCheckOutTime = value
End Property
Public Property Get PendingComplaints() As DTOs.ApptComplaints
    'SCR-52097
    If mPendingComplaints Is Nothing Then Set mPendingComplaints = New DTOs.ApptComplaints
    Set PendingComplaints = mPendingComplaints
End Property
Public Property Set PendingComplaints(ByRef value As ApptComplaints)
    Set mPendingComplaints = value
End Property
Public Property Get Priority() As Integer
    Priority = mPriority
End Property
Public Property Let Priority(ByVal value As Integer)
    mPriority = value
End Property
Public Property Get Classification() As DTOs.ApptClassification
    'SCR-52097
    If mClassification Is Nothing Then Set mClassification = New DTOs.ApptClassification
    Classification = mClassification
End Property
Public Property Let Classification(ByRef value As ApptClassification)
    mClassification = value
End Property
Public Property Get ADMEncounter() As String
    ADMEncounter = mADMEncounter
End Property
Public Property Let ADMEncounter(ByVal value As String)
    mADMEncounter = value
End Property
Public Property Get DetailCodes() As String
    DetailCodes = mDetailCodes
End Property
Public Property Let DetailCodes(ByVal value As String)
    mDetailCodes = value
End Property
Public Property Get Duration() As String
    Duration = mDuration
End Property
Public Property Let Duration(ByVal value As String)
    mDuration = value
End Property
Public Property Get CHCSStatus() As String
    CHCSStatus = mCHCSStatus
End Property
Public Property Let CHCSStatus(ByVal value As String)
    mCHCSStatus = value
End Property
Public Property Get CHCSType() As String
    CHCSType = mCHCSType
End Property
Public Property Let CHCSType(ByVal value As String)
    mCHCSType = value
End Property
Public Property Get CancelReason() As String
    CancelReason = mCancelReason
End Property
Public Property Let CancelReason(ByVal value As String)
    mCancelReason = value
End Property
Public Property Get Comment() As String
    Comment = mComment
End Property
Public Property Let Comment(ByVal value As String)
    mComment = value
End Property
Public Property Get ID() As String
    ID = mID
End Property
Public Property Let ID(ByVal value As String)
    mID = value
End Property
Public Property Get IEN() As String
    IEN = mIEN
End Property
Public Property Let IEN(ByVal value As String)
    mIEN = value
End Property
Public Property Get Reason() As String
    Reason = mReason
End Property
Public Property Let Reason(ByVal value As String)
    mReason = value
End Property
Public Property Get Status() As String
    Status = mStatus
End Property
Public Property Let Status(ByVal value As String)
    mStatus = value
End Property
Public Property Get TypeAppt() As String
    TypeAppt = mType
End Property
Public Property Let TypeAppt(ByVal value As String)
    mType = value
End Property
Public Property Get TypeCode() As String
    TypeCode = mTypeCode
End Property
Public Property Let TypeCode(ByVal value As String)
    mTypeCode = value
End Property
Public Property Get ClinicianNCID() As String
    ClinicianNCID = mClinicianNCID
End Property
Public Property Let ClinicianNCID(ByVal value As String)
    mClinicianNCID = value
End Property
Public Property Get ClinicName() As String
    ClinicName = mClinicName
End Property
Public Property Let ClinicName(ByVal value As String)
    mClinicName = value
End Property
Public Property Get ClinicNCID() As String
    ClinicNCID = mClinicNCID
End Property
Public Property Let ClinicNCID(ByVal value As String)
    mClinicNCID = value
End Property
Public Property Get EncounterNumber() As String
    EncounterNumber = mEncounterNumber
End Property
Public Property Let EncounterNumber(ByVal value As String)
    mEncounterNumber = value
End Property
Public Property Get EncounterStatus() As String
    EncounterStatus = mEncounterStatus
End Property
Public Property Let EncounterStatus(ByVal value As String)
    mEncounterStatus = value
End Property
Public Property Get FacilityNCID() As String
    FacilityNCID = mFacilityNCID
End Property
Public Property Let FacilityNCID(ByVal value As String)
    mFacilityNCID = value
End Property
Public Property Get InjuryIllnessCategory() As String
    InjuryIllnessCategory = mInjuryIllnessCategory
End Property
Public Property Let InjuryIllnessCategory(ByVal value As String)
    mInjuryIllnessCategory = value
End Property
Public Property Get InjuryIllnessCause() As String
    InjuryIllnessCause = mInjuryIllnessCause
End Property
Public Property Let InjuryIllnessCause(ByVal value As String)
    mInjuryIllnessCause = value
End Property
Public Property Get OrderNumber() As String
    OrderNumber = mOrderNumber
End Property
Public Property Let OrderNumber(ByVal value As String)
    mOrderNumber = value
End Property
Public Property Get Patient() As DTOs.Patient
    'SCR-52097
    If mPatient Is Nothing Then Set mPatient = New DTOs.Patient
    Set Patient = mPatient
End Property
Public Property Set Patient(ByRef value As Patient)
    Set mPatient = value
End Property
Public Property Get ProviderName() As String
    ProviderName = mProviderName
End Property
Public Property Let ProviderName(ByVal value As String)
    mProviderName = value
End Property
Public Property Get SponsorSSN() As String
    SponsorSSN = mSponsorSSN
End Property
Public Property Let SponsorSSN(ByVal value As String)
    mSponsorSSN = value
End Property
Public Property Get USVType() As String
    USVType = mUSVType
End Property
Public Property Let USVType(ByVal value As String)
    mUSVType = value
End Property
Public Property Get WorkloadType() As String
    WorkloadType = mWorkloadType
End Property
Public Property Let WorkloadType(ByVal value As String)
    mWorkloadType = value
End Property
Public Property Get WorkloadTypeOption() As String
    WorkloadTypeOption = mWorkloadTypeOption
End Property
Public Property Let WorkloadTypeOption(ByVal value As String)
    mWorkloadTypeOption = value
End Property
Public Property Get RFV() As String
    RFV = mRFV
End Property
Public Property Let RFV(ByVal value As String)
    mRFV = value
End Property
Public Property Get Note() As String
    Note = mNote
End Property
Public Property Let Note(ByVal value As String)
    mNote = value
End Property
Public Property Get NotesDataID() As String
    NotesDataID = mNotesDataID
End Property
Public Property Let NotesDataID(ByVal value As String)
    mNotesDataID = value
End Property
Public Property Get Workload() As String
    Workload = mWorkload
End Property
Public Property Let Workload(ByVal value As String)
    mWorkload = value
End Property
Public Property Get ResultedConsultOrder() As Boolean
    ResultedConsultOrder = mResultedConsultOrder
End Property
Public Property Let ResultedConsultOrder(ByVal value As Boolean)
    mResultedConsultOrder = value
End Property




Private Function iObjectPersist_Copy() As Object
    Dim NewCopy As Appointment
    Dim Persist As iObjectPersist
    
    Set NewCopy = New Appointment
    
    With NewCopy
        .ArrivedByAmbulance = mArrivedByAmbulance
        .RelatedToInpatient = mRelatedToInpatient
        .DateTime = mDateTime
        .CheckInTime = mCheckInTime
        .CheckOutTime = mCheckOutTime
        
        Set Persist = mPendingComplaints
        Set .PendingComplaints = Persist.Copy
        
        .Priority = mPriority
        
        Set Persist = mClassification
        Set .Classification = Persist.Copy
        
        .ADMEncounter = mADMEncounter
        .DetailCodes = mDetailCodes
        .Duration = mDuration
        .CHCSStatus = mCHCSStatus
        .CHCSType = mCHCSType
        .CancelReason = mCancelReason
        .Comment = mComment
        .ID = mID
        .IEN = mIEN
        .Reason = mReason
        .Status = mStatus
        .TypeAppt = mType
        .TypeCode = mTypeCode
        .ClinicianNCID = mClinicianNCID
        .ClinicName = mClinicName
        .ClinicNCID = mClinicNCID
        .EncounterNumber = mEncounterNumber
        .EncounterStatus = mEncounterStatus
        .FacilityNCID = mFacilityNCID
        .InjuryIllnessCategory = mInjuryIllnessCategory
        .InjuryIllnessCause = mInjuryIllnessCause
        .OrderNumber = mOrderNumber
        
        Set Persist = mPatient
        Set .Patient = Persist.Copy
        
        .ProviderName = mProviderName
        .SponsorSSN = mSponsorSSN
        .USVType = mUSVType
        .WorkloadType = mWorkloadType
        .WorkloadTypeOption = mWorkloadTypeOption
        .RFV = mRFV
        .Note = mNote
        .NotesDataID = mNotesDataID
        .Workload = mWorkload
        .ResultedConsultOrder = mResultedConsultOrder
    End With

    Set iObjectPersist_Copy = NewCopy

End Function




Private Sub iObjectPersist_Persist(Writer As iWriter)
    With Writer
        .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION
        .WriteProperty "ArrivedByAmbulance", mArrivedByAmbulance
        .WriteProperty "RelatedToInpatient", mRelatedToInpatient
        .WriteProperty "DateTime", mDateTime
        .WriteProperty "CheckInTime", mCheckInTime
        .WriteProperty "CheckOutTime", mCheckOutTime
        .WriteObject "PendingComplaints", mPendingComplaints
        .WriteProperty "Priority", mPriority
        .WriteObject "Classification", mClassification
        .WriteProperty "ADMEncounter", mADMEncounter
        .WriteProperty "DetailCodes", mDetailCodes
        .WriteProperty "Duration", mDuration
        .WriteProperty "CHCSStatus", mCHCSStatus
        .WriteProperty "CHCSType", mCHCSType
        .WriteProperty "CancelReason", mCancelReason
        .WriteProperty "Comment", mComment
        .WriteProperty "ID", mID
        .WriteProperty "IEN", mIEN
        .WriteProperty "Reason", mReason
        .WriteProperty "Status", mStatus
        .WriteProperty "Type", mType
        .WriteProperty "TypeCode", mTypeCode
        .WriteProperty "ClinicianNCID", mClinicianNCID
        .WriteProperty "ClinicName", mClinicName
        .WriteProperty "ClinicNCID", mClinicNCID
        .WriteProperty "EncounterNumber", mEncounterNumber
        .WriteProperty "EncounterStatus", mEncounterStatus
        .WriteProperty "FacilityNCID", mFacilityNCID
        .WriteProperty "InjuryIllnessCategory", mInjuryIllnessCategory
        .WriteProperty "InjuryIllnessCause", mInjuryIllnessCause
        .WriteProperty "OrderNumber", mOrderNumber
        .WriteObject "Patient", mPatient
        .WriteProperty "ProviderName", mProviderName
        .WriteProperty "SponsorSSN", mSponsorSSN
        .WriteProperty "USVType", mUSVType
        .WriteProperty "WorkloadType", mWorkloadType
        .WriteProperty "WorkloadTypeOption", mWorkloadTypeOption
        .WriteProperty "RFV", mRFV
        .WriteProperty "Note", mNote
        .WriteProperty "NotesDataID", mNotesDataID
        .WriteProperty "Workload", mWorkload
        .WriteProperty "ResultedConsultOrder", mResultedConsultOrder
    End With
End Sub




Private Function iObjectPersist_Restore(Reader As iReader) As Boolean
Dim Failed As Boolean

    On Error GoTo ErrHandler
  
    With Reader
        mArrivedByAmbulance = .ReadProperty("ArrivedByAmbulance")
        mRelatedToInpatient = .ReadProperty("RelatedToInpatient")
        mDateTime = .ReadProperty("DateTime")
        mCheckInTime = .ReadProperty("CheckInTime")
        mCheckOutTime = .ReadProperty("CheckOutTime")
        .ReadObject "PendingComplaints", PendingComplaints
        mPriority = .ReadProperty("Priority")
        .ReadObject "Classification", Classification
        mADMEncounter = .ReadProperty("ADMEncounter")
        mDetailCodes = .ReadProperty("DetailCodes")
        mDuration = .ReadProperty("Duration")
        mCHCSStatus = .ReadProperty("CHCSStatus")
        mCHCSType = .ReadProperty("CHCSType")
        mCancelReason = .ReadProperty("CancelReason")
        mComment = .ReadProperty("Comment")
        mID = .ReadProperty("ID")
        mIEN = .ReadProperty("IEN")
        mReason = .ReadProperty("Reason")
        mStatus = .ReadProperty("Status")
        mType = .ReadProperty("Type")
        mTypeCode = .ReadProperty("TypeCode")
        mClinicianNCID = .ReadProperty("ClinicianNCID")
        mClinicName = .ReadProperty("ClinicName")
        mClinicNCID = .ReadProperty("ClinicNCID")
        mEncounterNumber = .ReadProperty("EncounterNumber")
        mEncounterStatus = .ReadProperty("EncounterStatus")
        mFacilityNCID = .ReadProperty("FacilityNCID")
        mInjuryIllnessCategory = .ReadProperty("InjuryIllnessCategory")
        mInjuryIllnessCause = .ReadProperty("InjuryIllnessCause")
        mOrderNumber = .ReadProperty("OrderNumber")
        .ReadObject "Patient", Patient
        mProviderName = .ReadProperty("ProviderName")
        mSponsorSSN = .ReadProperty("SponsorSSN")
        mUSVType = .ReadProperty("USVType")
        mWorkloadType = .ReadProperty("WorkloadType")
        mWorkloadTypeOption = .ReadProperty("WorkloadTypeOption")
        mRFV = .ReadProperty("RFV")
        mNote = .ReadProperty("Note")
        mNotesDataID = .ReadProperty("NotesDataID")
        mWorkload = .ReadProperty("Workload")
        mResultedConsultOrder = .ReadProperty("ResultedConsultOrder")
    End With

    iObjectPersist_Restore = Not Failed
  
    Exit Function
  
ErrHandler:
    Failed = True
    Resume Next

End Function


Private Property Get iObjectPersist_UniqueIdentifier() As String
    iObjectPersist_UniqueIdentifier = UNIQUE_IDENTIFIER
End Property

Private Property Get iObjectPersist_Version() As String
    iObjectPersist_Version = OBJECT_VERSION
End Property
'--- end DTO Builder

