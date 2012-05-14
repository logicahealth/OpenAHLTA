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
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--- begin DTO Builder; DTO=Encounter
'--- 3/16/04 7:40:29 PM
'---
'--- NOTE: Any time you modify this class, increment the version number
'
'  Medical Record Encounter

Option Explicit

Implements iObjectPersist
Implements iState

Private Const OBJECT_VERSION = "1.0"
Private Const UNIQUE_IDENTIFIER = "Encounter"

Private mState As enumObjectState
Private mID As Long '  Encounter ID
Private mFacilityNCID As String '
Private mFacilityName As String '
Private mClinicNCID As String   '
Private mClinicName As String   '
Private mStatus As String   '
Private mClass As String    '
Private mEncType As String '
Private mWhenSignedDate As Date '
Private mPatientName As String  '
Private mPatientUnitNum As String   '
Private mCreatedDate As Date    '
Private mPrevSignedRTFs As Documents    '
Private mProviders As EncProviders  '
Private mSADRStatus As String   '
Private mSensitivity As Long    '
Private mTemplateData As String '
Private mTemplatedID As String  '
Private mCosigner As EncProvider    '
Private mWhoSigned As EncProvider   '
Private mWhoCosigned As EncProvider '
Private mWhenCosigned As Date   '
Private mWasUpdated As Boolean  '
Private mUpdateLock As Boolean  '
Private mInpatientService As String '
Private mAccidentData As EncAccident    '
Private mPregnancyData As EncPregnancy  '
Private mSections As Sections
Private mAppointment As Appointment
Private mDataID As Long
Private mPrimaryDxSnoID As String
Private mUpdateDate As Date
Private mUpdateClinicianNCID As String
Private mCreateClinicianNCID As String

Public Property Get CreateClinicianNCID() As String
    CreateClinicianNCID = mCreateClinicianNCID
End Property
Public Property Let CreateClinicianNCID(value As String)
    mCreateClinicianNCID = value
End Property
Public Property Get UpdateClinicianNCID() As String
    UpdateClinicianNCID = mUpdateClinicianNCID
End Property
Public Property Let UpdateClinicianNCID(value As String)
    mUpdateClinicianNCID = value
End Property
Public Property Get UpdateDate() As Date
    UpdateDate = mUpdateDate
End Property
Public Property Let UpdateDate(value As Date)
    mUpdateDate = value
End Property
Public Property Get PrimaryDxSnoID() As String
    PrimaryDxSnoID = mPrimaryDxSnoID
End Property
Public Property Let PrimaryDxSnoID(value As String)
    mPrimaryDxSnoID = value
End Property
Public Property Get DataID() As Long
    DataID = mDataID
End Property
Public Property Let DataID(value As Long)
    mDataID = value
End Property
Public Property Get Sections() As Sections
    If mSections Is Nothing Then Set mSections = New Sections
    Set Sections = mSections
End Property
Public Property Set Sections(value As Sections)
    Set mSections = value
End Property

Public Property Get Appointment() As Appointment
    If mAppointment Is Nothing Then Set mAppointment = New Appointment
    Set Appointment = mAppointment
End Property
Public Property Set Appointment(value As Appointment)
    Set mAppointment = value
End Property

Public Property Get ID() As Long
    ID = mID
End Property
Public Property Let ID(value As Long)
    mID = value
End Property
Public Property Get FacilityNCID() As String
    FacilityNCID = mFacilityNCID
End Property
Public Property Let FacilityNCID(value As String)
    mFacilityNCID = value
End Property
Public Property Get FacilityName() As String
    FacilityName = mFacilityName
End Property
Public Property Let FacilityName(value As String)
    mFacilityName = value
End Property
Public Property Get ClinicNCID() As String
    ClinicNCID = mClinicNCID
End Property
Public Property Let ClinicNCID(value As String)
    mClinicNCID = value
End Property
Public Property Get ClinicName() As String
    ClinicName = mClinicName
End Property
Public Property Let ClinicName(value As String)
    mClinicName = value
End Property
Public Property Get Status() As String
    Status = mStatus
End Property
Public Property Let Status(value As String)
    mStatus = value
End Property
Public Property Get Class() As String
    Class = mClass
End Property
Public Property Let Class(value As String)
    mClass = value
End Property

Public Property Get EncType() As String
    EncType = mEncType
End Property
Public Property Let EncType(value As String)
    mEncType = value
End Property
Public Property Get WhenSignedDate() As Date
    WhenSignedDate = mWhenSignedDate
End Property
Public Property Let WhenSignedDate(value As Date)
    mWhenSignedDate = value
End Property
Public Property Get PatientName() As String
    PatientName = mPatientName
End Property
Public Property Let PatientName(value As String)
    mPatientName = value
End Property
Public Property Get PatientUnitNum() As String
    PatientUnitNum = mPatientUnitNum
End Property
Public Property Let PatientUnitNum(value As String)
    mPatientUnitNum = value
End Property
Public Property Get CreatedDate() As Date
    CreatedDate = mCreatedDate
End Property
Public Property Let CreatedDate(value As Date)
    mCreatedDate = value
End Property
Public Property Get PrevSignedRTFs() As Documents
    If mPrevSignedRTFs Is Nothing Then Set mPrevSignedRTFs = New Documents
    Set PrevSignedRTFs = mPrevSignedRTFs
End Property
Public Property Set PrevSignedRTFs(value As Documents)
    Set mPrevSignedRTFs = value
End Property
Public Property Get Providers() As EncProviders
    If mProviders Is Nothing Then Set mProviders = New EncProviders
    Set Providers = mProviders
End Property
Public Property Set Providers(value As EncProviders)
    Set mProviders = value
End Property
Public Property Get SADRStatus() As String
    SADRStatus = mSADRStatus
End Property
Public Property Let SADRStatus(value As String)
    mSADRStatus = value
End Property
Public Property Get Sensitivity() As Long
    Sensitivity = mSensitivity
End Property
Public Property Let Sensitivity(value As Long)
    mSensitivity = value
End Property
Public Property Get TemplateData() As String
    TemplateData = mTemplateData
End Property
Public Property Let TemplateData(value As String)
    mTemplateData = value
End Property
Public Property Get TemplatedID() As String
    TemplatedID = mTemplatedID
End Property
Public Property Let TemplatedID(value As String)
    mTemplatedID = value
End Property
Public Property Get Cosigner() As EncProvider
    If mCosigner Is Nothing Then Set mCosigner = New EncProvider
    Set Cosigner = mCosigner
End Property
Public Property Set Cosigner(value As EncProvider)
    Set mCosigner = value
End Property
Public Property Get WhoSigned() As EncProvider
    If mWhoSigned Is Nothing Then Set mWhoSigned = New EncProvider
    Set WhoSigned = mWhoSigned
End Property
Public Property Set WhoSigned(value As EncProvider)
    Set mWhoSigned = value
End Property
Public Property Get WhoCosigned() As EncProvider
    If mWhoCosigned Is Nothing Then Set mWhoCosigned = New EncProvider
    Set WhoCosigned = mWhoCosigned
End Property
Public Property Set WhoCosigned(value As EncProvider)
    Set mWhoCosigned = value
End Property
Public Property Get WhenCosigned() As Date
    WhenCosigned = mWhenCosigned
End Property
Public Property Let WhenCosigned(value As Date)
    mWhenCosigned = value
End Property
Public Property Get WasUpdated() As Boolean
    WasUpdated = mWasUpdated
End Property
Public Property Let WasUpdated(value As Boolean)
    mWasUpdated = value
End Property
Public Property Get UpdateLock() As Boolean
    UpdateLock = mUpdateLock
End Property
Public Property Let UpdateLock(value As Boolean)
    mUpdateLock = value
End Property
Public Property Get InpatientService() As String
    InpatientService = mInpatientService
End Property
Public Property Let InpatientService(value As String)
    mInpatientService = value
End Property
Public Property Get AccidentData() As EncAccident
    If mAccidentData Is Nothing Then Set mAccidentData = New EncAccident
    Set AccidentData = mAccidentData
End Property
Public Property Set AccidentData(value As EncAccident)
    Set mAccidentData = value
End Property
Public Property Get PregnancyData() As EncPregnancy
    If mPregnancyData Is Nothing Then Set mPregnancyData = New EncPregnancy
    Set PregnancyData = mPregnancyData
End Property
Public Property Set PregnancyData(value As EncPregnancy)
    Set mPregnancyData = value
End Property

Private Function iObjectPersist_Copy() As Object
    Dim NewCopy As DTOs.Encounter
    Dim State As iState
    
    Set NewCopy = New DTOs.Encounter
    
    With NewCopy
        .ID = mID
        .FacilityNCID = mFacilityNCID
        .FacilityName = mFacilityName
        .ClinicNCID = mClinicNCID
        .ClinicName = mClinicName
        .Status = mStatus
        .Class = mClass
        .EncType = mEncType
        .WhenSignedDate = mWhenSignedDate
        .PatientName = mPatientName
        .PatientUnitNum = mPatientUnitNum
        .CreatedDate = mCreatedDate
        Set .PrevSignedRTFs = mPrevSignedRTFs
        Set .Providers = mProviders
        .SADRStatus = mSADRStatus
        .Sensitivity = mSensitivity
        .TemplateData = mTemplateData
        .TemplatedID = mTemplatedID
        .Cosigner = mCosigner
        .WhoSigned = mWhoSigned
        .WhoCosigned = mWhoCosigned
        .WhenCosigned = mWhenCosigned
        .WasUpdated = mWasUpdated
        .UpdateLock = mUpdateLock
        .InpatientService = mInpatientService
        .AccidentData = mAccidentData
        .PregnancyData = mPregnancyData
        .DataID = mDataID
        .UpdateDate = mUpdateDate
        .UpdateClinicianNCID = mUpdateClinicianNCID
        .CreateClinicianNCID = mCreateClinicianNCID
        
        'iState interface
        Set State = NewCopy
        State.State = mState
    End With

    Set iObjectPersist_Copy = NewCopy

End Function




Private Sub iObjectPersist_Persist(Writer As iWriter)
    With Writer
        .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION
        .WriteProperty "ID", mID
        .WriteProperty "FacilityNCID", mFacilityNCID
        .WriteProperty "FacilityName", mFacilityName
        .WriteProperty "ClinicNCID", mClinicNCID
        .WriteProperty "ClinicName", mClinicName
        .WriteProperty "Status", mStatus
        .WriteProperty "Class", mClass
        .WriteProperty "EncType", mEncType
        .WriteProperty "WhenSignedDate", mWhenSignedDate
        .WriteProperty "PatientName", mPatientName
        .WriteProperty "PatientUnitNum", mPatientUnitNum
        .WriteProperty "CreatedDate", mCreatedDate
        .WriteObject "PrevSignedRTFs", mPrevSignedRTFs
        .WriteObject "Providers", mProviders
        .WriteProperty "SADRStatus", mSADRStatus
        .WriteProperty "Sensitivity", mSensitivity
        .WriteProperty "TemplateData", mTemplateData
        .WriteProperty "TemplatedID", mTemplatedID
        .WriteProperty "Cosigner", mCosigner
        .WriteProperty "WhoSigned", mWhoSigned
        .WriteProperty "WhoCosigned", mWhoCosigned
        .WriteProperty "WhenCosigned", mWhenCosigned
        .WriteProperty "WasUpdated", mWasUpdated
        .WriteProperty "UpdateLock", mUpdateLock
        .WriteProperty "InpatientService", mInpatientService
        .WriteProperty "AccidentData", mAccidentData
        .WriteProperty "PregnancyData", mPregnancyData
        .WriteProperty "DataID", mDataID
        .WriteProperty "UpdateDate", mUpdateDate
        .WriteProperty "UpdateClinicianNCID", mUpdateClinicianNCID
        .WriteProperty "CreateClinicianNCID", mCreateClinicianNCID
        
        'iState interface
        .WriteProperty "iState.State", mState
    End With
End Sub




Private Function iObjectPersist_Restore(Reader As iReader) As Boolean
Dim Failed As Boolean

    On Error GoTo ErrHandler
  
    With Reader
        mID = .ReadProperty("ID")
        mFacilityNCID = .ReadProperty("FacilityNCID")
        mFacilityName = .ReadProperty("FacilityName")
        mClinicNCID = .ReadProperty("ClinicNCID")
        mClinicName = .ReadProperty("ClinicName")
        mStatus = .ReadProperty("Status")
        mClass = .ReadProperty("Class")
        mEncType = .ReadProperty("EncType")
        mWhenSignedDate = .ReadProperty("WhenSignedDate")
        mPatientName = .ReadProperty("PatientName")
        mPatientUnitNum = .ReadProperty("PatientUnitNum")
        mCreatedDate = .ReadProperty("CreatedDate")
        .ReadObject "PrevSignedRTFs", PrevSignedRTFs
        .ReadObject "Providers", Providers
        mSADRStatus = .ReadProperty("SADRStatus")
        mSensitivity = .ReadProperty("Sensitivity")
        mTemplateData = .ReadProperty("TemplateData")
        mTemplatedID = .ReadProperty("TemplatedID")
        mCosigner = .ReadProperty("Cosigner")
        mWhoSigned = .ReadProperty("WhoSigned")
        mWhoCosigned = .ReadProperty("WhoCosigned")
        mWhenCosigned = .ReadProperty("WhenCosigned")
        mWasUpdated = .ReadProperty("WasUpdated")
        mUpdateLock = .ReadProperty("UpdateLock")
        mInpatientService = .ReadProperty("InpatientService")
        mAccidentData = .ReadProperty("AccidentData")
        mPregnancyData = .ReadProperty("PregnancyData")
        mDataID = .ReadProperty("DataID")
        mUpdateDate = .ReadProperty("UpdateDate")
        mUpdateClinicianNCID = .ReadProperty("UpdateClinicianNCID")
        mCreateClinicianNCID = .ReadProperty("CreateClinicianNCID")
        
        'iState interface
        mState = .ReadProperty("iState.State")
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
Private Property Let iState_State(ByVal NewValue As enumObjectState)
    mState = NewValue
End Property

Private Property Get iState_State() As enumObjectState
    iState_State = mState
End Property

Private Property Let iState_ID(ByVal RHS As String)

End Property

Private Property Get iState_ID() As String
    iState_ID = UNIQUE_IDENTIFIER
End Property
