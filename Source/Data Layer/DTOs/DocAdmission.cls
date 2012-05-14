VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DocAdmission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

'--- begin XML DTO Builder; DTO=DocAdmission
'--- [Insert DateTime Stamp here]
'---
'--- NOTE: Any time you modify this class, increment the version number
'
'

Implements iObjectPersist
Implements iState

Private Const OBJECT_VERSION = "1.0"
Private Const UNIQUE_IDENTIFIER = "DocAdmission"

Private mPrimaryDiagnosis As String 'Primary Diagnosis for the Inpatient Admission
Private mIsInpatient As Boolean 'Is the admission an inpatient or outpatient
Private mAdmissionDate As Date 'Date the admission was created
Private mDischargeDate As Date 'Date the patient was discharged
Private mProviderName As String 'The admitting provider
Private mWard As String 'The ward the patient is in
Private mSpecialty As String 'The specialty of the admission
Private mDocEncounters As DocEncounters 'The DocEncounters object
Private mStatus As String 'Status of the admission
Private mInpatientID As Long 'ID of the Inpatient Admission
'Private mDays As Integer 'Elapsed days from the admission date
Private mState As enumObjectState



Public Property Get PrimaryDiagnosis() As String
    PrimaryDiagnosis = mPrimaryDiagnosis
End Property
Public Property Let PrimaryDiagnosis(ByVal value As String)
    mPrimaryDiagnosis = value
End Property
                        
Public Property Get IsInpatient() As Boolean
    IsInpatient = mIsInpatient
End Property
Public Property Let IsInpatient(ByVal value As Boolean)
    mIsInpatient = value
End Property
                        
Public Property Get AdmissionDate() As Date
    AdmissionDate = mAdmissionDate
End Property
Public Property Let AdmissionDate(ByVal value As Date)
    mAdmissionDate = value
End Property
 Public Property Get DischargeDate() As Date
    DischargeDate = mDischargeDate
End Property
Public Property Let DischargeDate(ByVal value As Date)
    mDischargeDate = value
End Property
                       
Public Property Get AdmissionDateRange() As String
    
    AdmissionDateRange = mAdmissionDate & " - " & IIf(mDischargeDate = 0, "", mDischargeDate)
    
End Property

Public Property Get ProviderName() As String
    ProviderName = mProviderName
End Property
Public Property Let ProviderName(ByVal value As String)
    mProviderName = value
End Property
                        
Public Property Get Ward() As String
    Ward = mWard
End Property
Public Property Let Ward(ByVal value As String)
    mWard = value
End Property
                        
Public Property Get Specialty() As String
    Specialty = mSpecialty
End Property
Public Property Let Specialty(ByVal value As String)
    mSpecialty = value
End Property
                        
Public Property Get DocEncounters() As DocEncounters
    If mDocEncounters Is Nothing Then
        Set mDocEncounters = New DocEncounters
    End If
    Set DocEncounters = mDocEncounters
End Property
Public Property Set DocEncounters(ByRef value As DocEncounters)
    Set mDocEncounters = value
End Property
                        
'Public Property Get Status() As String
'    Status = mStatus
'End Property
'Public Property Let Status(ByVal value As String)
'    mStatus = value
'End Property
                        
Public Property Get StatusText() As String
        
    If mDischargeDate = 0 Then
        StatusText = "Admitted"
    Else
        StatusText = "Discharged"
    End If
    
End Property

Public Property Get InpatientID() As Long
    InpatientID = mInpatientID
End Property
Public Property Let InpatientID(ByVal value As Long)
    mInpatientID = value
End Property
                        
Public Property Get Days() As Integer

'ToDO: Calulate days
'if discharge date=0 then days=now-admit date +1
'else days = discharge date - admit date +1
    If DischargeDate = 0 Then
        Days = DateDiff("d", mAdmissionDate, Now) + 1
    Else
        Days = DateDiff("d", mAdmissionDate, mDischargeDate) + 1
    End If
    
End Property
Public Property Let Days(ByVal value As Integer)
'    mDays = value
End Property
                        
Private Function iObjectPersist_Copy() As Object
    Dim NewCopy As DocAdmission
    Dim PersistHelper As ObjectPersist.PersistHelper
    Dim State As iState
    Dim i As Long 'Iterator if we need to copy a collection
    
    Set NewCopy = New DocAdmission
    Set PersistHelper = New ObjectPersist.PersistHelper

    With NewCopy
                .PrimaryDiagnosis = mPrimaryDiagnosis
                .IsInpatient = mIsInpatient
                .AdmissionDate = mAdmissionDate
                .DischargeDate = mDischargeDate
                .ProviderName = mProviderName
                .Ward = mWard
                .Specialty = mSpecialty
                Set .DocEncounters = PersistHelper.CopyObject(mDocEncounters)
                '.Status = mStatus
                .InpatientID = mInpatientID
                '.Days = mDays

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
        
        .WriteProperty "PrimaryDiagnosis", mPrimaryDiagnosis
        .WriteProperty "IsInpatient", mIsInpatient
        .WriteProperty "AdmissionDate", mAdmissionDate
        .WriteProperty "DischargeDate", mDischargeDate
        .WriteProperty "ProviderName", mProviderName
        .WriteProperty "Ward", mWard
        .WriteProperty "Specialty", mSpecialty
                .WriteObject "DocEncounters", mDocEncounters
        .WriteProperty "Status", mStatus
        .WriteProperty "InpatientID", mInpatientID
        '.WriteProperty "Days", mDays

        'iState interface
        .WriteProperty "iState.State", mState

    End With
End Sub


Private Function iObjectPersist_Restore(Reader As iReader) As Boolean
Dim Failed As Boolean

    On Error GoTo ErrHandler

    With Reader
        
        mPrimaryDiagnosis = .ReadProperty("PrimaryDiagnosis")
        mIsInpatient = .ReadProperty("IsInpatient")
        mAdmissionDate = .ReadProperty("AdmissionDate")
        mDischargeDate = .ReadProperty("DischargeDate")
        mProviderName = .ReadProperty("ProviderName")
        mWard = .ReadProperty("Ward")
        mSpecialty = .ReadProperty("Specialty")
                .ReadObject "DocEncounters", mDocEncounters
        mStatus = .ReadProperty("Status")
        mInpatientID = .ReadProperty("InpatientID")
        'mDays = .ReadProperty("Days")

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
