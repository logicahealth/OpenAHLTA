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
Attribute VB_Name = "DocEncounter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

'--- begin XML DTO Builder; DTO=DocEncounter
'--- [Insert DateTime Stamp here]
'---
'--- NOTE: Any time you modify this class, increment the version number
'
'

Implements iObjectPersist
Implements iState

Private Const OBJECT_VERSION = "1.0"
Private Const UNIQUE_IDENTIFIER = "DocEncounter"

Private mDocumentType As String 'The document type
Private mDateTime As Date 'Date of the encounter
Private mProviderName As String 'The attending provider
Private mHospDay As Long 'The hospital day that the encounter was created on.

'This is the NCID status not the regular enum status
Private mStatus As String 'The NCID status of the encounter
Private meStatus As EncStatus 'The regular enum status of the encounter

Private mWard As String 'The ward the patient is in
Private mEncounterNumber As String '
Private mFacilityNCID As String '
Private mIsInpatient As Boolean 'Is the encounter an inpatient or outpatient encounter
Private mState As enumObjectState
Private mCreatedBy As String
Private mPrimaryDiag As String
Private mUpdateLock As Boolean
Private mSensitivity As Boolean

'Reference it's parent
Private mAdmission As DocAdmission


Public Property Let CreatedBy(vData As String)
    
    mCreatedBy = vData
    
End Property

Public Property Get CreatedBy() As String
    
    CreatedBy = mCreatedBy
    
End Property


Public Property Get Days() As Integer
'Remove this when pushing to the build.
End Property


Public Property Let PrimaryDiagnosis(vData As String)
    
    mPrimaryDiag = vData
    
End Property

Public Property Get PrimaryDiagnosis() As String
    
    PrimaryDiagnosis = mPrimaryDiag
    
End Property

Private Function StatusEnumToNCID(ByVal eStatus As EncStatus) As String
    Select Case eStatus
    Case Incomplete
        StatusEnumToNCID = "75657"
    Case NeedsSignature
        StatusEnumToNCID = "14510305"
    Case NeedsCoSignature
        StatusEnumToNCID = "14510306"
    Case Pending
        StatusEnumToNCID = "662"
    Case CheckedIn
        StatusEnumToNCID = "14510307"
    Case Waiting
        StatusEnumToNCID = "14510308"
    Case InProgress
        StatusEnumToNCID = "11992"
    Case Updated
        StatusEnumToNCID = "14510309"
    Case complete
        StatusEnumToNCID = "205657"
    Case Updating
        StatusEnumToNCID = "14510310"
    Case PatientLeftWithoutBeingSeen
        StatusEnumToNCID = "14508309"
    Case CancelByPatient
        StatusEnumToNCID = "14505856"
    Case CancelByFacility
        StatusEnumToNCID = "14613990"
    Case NoShow
        StatusEnumToNCID = "665"
    Case CancelByProvider
        StatusEnumToNCID = "160595"
    End Select
End Function

Private Function StatusNCIDToEnum(ByRef sNCID As String) As EncStatus
    Select Case sNCID
    Case "75657"
        StatusNCIDToEnum = Incomplete
    Case "14510305"
        StatusNCIDToEnum = NeedsSignature
    Case "14510306"
        StatusNCIDToEnum = NeedsCoSignature
    Case "662"
        StatusNCIDToEnum = Pending
    Case "14510307"
        StatusNCIDToEnum = CheckedIn
    Case "14510308"
        StatusNCIDToEnum = Waiting
    Case "11992"
        StatusNCIDToEnum = InProgress
    Case "14510309"
        StatusNCIDToEnum = Updated
    Case "205657"
        StatusNCIDToEnum = complete
    Case "14510310"
        StatusNCIDToEnum = Updating
    Case "14508309"
        StatusNCIDToEnum = PatientLeftWithoutBeingSeen
    Case "14505856"
        StatusNCIDToEnum = CancelByPatient
    Case "14613990"
        StatusNCIDToEnum = CancelByFacility
    Case "665"
        StatusNCIDToEnum = NoShow
    Case "160595"
        StatusNCIDToEnum = CancelByProvider
    End Select
End Function

Private Function StatusEnumToText(eStatus As EncStatus)

    Select Case eStatus
        Case Incomplete
            StatusEnumToText = "Incomplete"
        Case NeedsSignature
            StatusEnumToText = "Needs Signature"
        Case NeedsCoSignature
            StatusEnumToText = "Needs Co-Signature"
        Case Pending
            StatusEnumToText = "Pending"
        Case CheckedIn
            StatusEnumToText = "Checked-In"
        Case NoSave
            StatusEnumToText = "Waiting For VS"
        Case Waiting
            StatusEnumToText = "Waiting"
        Case InProgress
            StatusEnumToText = "In Progress"
        Case Updated
            StatusEnumToText = "Updated"
        Case complete
            StatusEnumToText = "Complete"
        Case Updating
            StatusEnumToText = "Updating"
        Case PatientLeftWithoutBeingSeen
            StatusEnumToText = "Left Without Being Seen"
        Case CancelByFacility
            StatusEnumToText = "Cancelled by Facility"
        Case CancelByPatient
            StatusEnumToText = "Cancelled By Patient"
        Case NoShow
            StatusEnumToText = "No Show"
    End Select

End Function

Public Property Set Admission(vData As DocAdmission)
    
    Set mAdmission = vData
    
End Property

Public Property Get Admission() As DocAdmission
    
    Set Admission = mAdmission
    
End Property


Public Property Get DocumentType() As String
    DocumentType = mDocumentType
End Property
Public Property Let DocumentType(ByVal value As String)
    mDocumentType = value
End Property
                        
Public Property Get DateTime() As Date
    DateTime = mDateTime
End Property
Public Property Let DateTime(ByVal value As Date)
    mDateTime = value
End Property
Public Property Get HospDay() As Long

    HospDay = mHospDay
    
End Property
                       

Public Property Let HospDay(Val As Long)

    mHospDay = Val
    
End Property
                       


Public Property Get EncounterKey() As String
    
    EncounterKey = mFacilityNCID & "|" & mEncounterNumber
    
End Property

Public Property Get ProviderName() As String
    ProviderName = mProviderName
End Property
Public Property Let ProviderName(ByVal value As String)
    mProviderName = value
End Property
                        
Public Property Get Status() As String
    Status = mStatus
End Property

Public Property Let Status(ByVal value As String)
    mStatus = value
End Property
                        
Public Property Get StatusEnum() As EncStatus
    
    If mIsInpatient = True Then
        StatusEnum = StatusNCIDToEnum(mStatus)
    Else
        StatusEnum = mStatus
    End If
    
End Property

Public Property Get StatusText() As String
    
    'SF This is crazy, we need a better solution across the projects.
    Dim eStatus As EncStatus

    If mIsInpatient = True Then
        eStatus = StatusNCIDToEnum(mStatus)
        
'        Select Case mStatus
'            Case "75657"
'                eStatus = Incomplete
'            Case "14510305"
'                eStatus = NeedsSignature
'            Case "14510306"
'                eStatus = NeedsCoSignature
'            Case "662"
'                eStatus = Pending
'            Case "14510307"
'                eStatus = CheckedIn
'            Case "14510308"
'                eStatus = Waiting
'            Case "11992"
'                eStatus = InProgress
'            Case "14510309"
'                eStatus = Updated
'            Case "205657"
'                eStatus = complete
'            Case "14510310"
'                eStatus = Updating
'            Case "14508309"
'                eStatus = PatientLeftWithoutBeingSeen
'            Case "14505856"
'                eStatus = CancelByPatient
'            Case "14613990"
'                eStatus = CancelByFacility
'            Case "665"
'                eStatus = NoShow
'            Case "160595"
'                eStatus = CancelByProvider
'        End Select
    
    ElseIf mIsInpatient = False Then
        eStatus = mStatus
    End If

    StatusText = StatusEnumToText(eStatus)
    
End Function

Public Property Let UpdateLock(vData As Boolean)
    
    mUpdateLock = vData
    
End Property

Public Property Get UpdateLock() As Boolean
    
    UpdateLock = mUpdateLock
    
End Property

Public Property Let Sensitivity(vData As Boolean)
    mSensitivity = vData
End Property

Public Property Get Sensitivity() As Boolean
    Sensitivity = mSensitivity
End Property

Public Property Get Ward() As String
    Ward = mWard
End Property
Public Property Let Ward(ByVal value As String)
    mWard = value
End Property
                        
Public Property Get EncounterNumber() As String
    EncounterNumber = mEncounterNumber
End Property
Public Property Let EncounterNumber(ByVal value As String)
    mEncounterNumber = value
End Property
                        
Public Property Get FacilityNCID() As String
    FacilityNCID = mFacilityNCID
End Property
Public Property Let FacilityNCID(ByVal value As String)
    mFacilityNCID = value
End Property
                        
Public Property Get IsInpatient() As Boolean
    IsInpatient = mIsInpatient
End Property
Public Property Let IsInpatient(ByVal value As Boolean)
    mIsInpatient = value
End Property
                        
Private Sub Class_Terminate()
    
    Set mAdmission = Nothing
    
End Sub


Private Function iObjectPersist_Copy() As Object
    Dim NewCopy As DocEncounter
    Dim PersistHelper As ObjectPersist.PersistHelper
    Dim State As iState
    Dim i As Long 'Iterator if we need to copy a collection
    
    Set NewCopy = New DocEncounter
    Set PersistHelper = New ObjectPersist.PersistHelper

    With NewCopy
                .DocumentType = mDocumentType
                .DateTime = mDateTime
                .ProviderName = mProviderName
                .Status = mStatus
                .Ward = mWard
                .EncounterNumber = mEncounterNumber
                .FacilityNCID = mFacilityNCID
                .IsInpatient = mIsInpatient
                .HospDay = mHospDay
                .CreatedBy = mCreatedBy
                .UpdateLock = mUpdateLock
                .Sensitivity = mSensitivity
                
        'iState interface
        Set State = NewCopy
        State.State = mState
    End With

    Set iObjectPersist_Copy = NewCopy
        Set NewCopy = Nothing
        Set State = Nothing
End Function

Private Sub iObjectPersist_Persist(Writer As iWriter)
    With Writer
        .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION
        
        .WriteProperty "DocumentType", mDocumentType
        .WriteProperty "DateTime", mDateTime
        .WriteProperty "ProviderName", mProviderName
        .WriteProperty "Status", mStatus
        .WriteProperty "Ward", mWard
        .WriteProperty "EncounterNumber", mEncounterNumber
        .WriteProperty "FacilityNCID", mFacilityNCID
        .WriteProperty "IsInpatient", mIsInpatient
        .WriteArray "HospDay", mHospDay
        .WriteArray "CreatedBy", mCreatedBy
        .WriteArray "UpdateLock", mUpdateLock
        .WriteArray "Sensitivity", mSensitivity
        
        'iState interface
        .WriteProperty "iState.State", mState

    End With
End Sub


Private Function iObjectPersist_Restore(Reader As iReader) As Boolean
Dim Failed As Boolean

    On Error GoTo ErrHandler

    With Reader
        
        mDocumentType = .ReadProperty("DocumentType")
        mDateTime = .ReadProperty("DateTime")
        mProviderName = .ReadProperty("ProviderName")
        mStatus = .ReadProperty("Status")
        mWard = .ReadProperty("Ward")
        mEncounterNumber = .ReadProperty("EncounterNumber")
        mFacilityNCID = .ReadProperty("FacilityNCID")
        mIsInpatient = .ReadProperty("IsInpatient")
        mHospDay = .ReadArray("HospDay")
        mCreatedBy = .ReadArray("CreatedBy")
        mUpdateLock = .ReadArray("UpdateLock")
        mSensitivity = .ReadArray("Sensitivity")
        
        'iState interface
        mState = .ReadProperty("iState.State")

    End With

    iObjectPersist_Restore = Not Failed

    Exit Function

ErrHandler:
    Failed = True
    Resume Next

End Function


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



Private Property Get iObjectPersist_UniqueIdentifier() As String
    iObjectPersist_UniqueIdentifier = UNIQUE_IDENTIFIER
End Property

Private Property Get iObjectPersist_Version() As String
    iObjectPersist_Version = OBJECT_VERSION
End Property
'--- end DTO Builder
