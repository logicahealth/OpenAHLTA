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
Attribute VB_Name = "Patient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True


Option Explicit

Implements iObjectPersist
Implements iState

Private Const OBJECT_VERSION = "1.0"
Private Const UNIQUE_IDENTIFIER = "Patient"


Private mACV_EnrollmentStatus As String
Private mAdopted As String
Private mAge As String
Private mAssigningAuthority As String
Private mBranchOfService As String
Private mCommandInterest As PatientGenericDatas
Private mCommandSecurity As PatientGenericDatas
Private mDateOfBirth As Double
Private mDeersEligibility As DeersEligibility
Private mDeployUnit As String
Private mDeployLocation As String
Private mDNBI As String
Private mEmailAddress As String
Private mEnrollmentFacility As String
Private mExpired As String
Private mFacility As String
Private mFMP As String
Private mFMPSponsorSSN As String
Private mForeignID_Flag As Boolean
Private mFormatted_SponsorSSN As String
Private mFormatted_SSN As String
Private mGenericDatas As PatientGenericDatas
Private mGuarantor As String
Private mMaidenName As String
Private mMaritalStatus As String
Private mMedicalRecordNumber As String
Private mMedicare As String
Private mMedicareCode As String
Private mMilitaryStatus As String
Private mMothersMaidenName As String
Private mHomeAddress As Address
Private mHomePhone As Phone
Private mWorkPhone As Phone
Private mPhones As Phones
Private mOrganDonor As String
Private mPatientCategory As String
Private mName As Name
Private mPayGrade As String
Private mPCM_IEN As String
Private mPregnancy As Pregnancy
Private mPrimaryCareManager As String
Private mRace As String
Private mRank As String
Private mRecordsMaintainedAt As String
Private mReligion As String
Private mSex As String
Private mSponsor As Patient
Private mSponsorSSN As String
Private mSponsorName As String
Private mSSN As String
Private mSpecialWorkStatus As SpecialWorkStatus
Private mUnitIdentificationCode As String
Private mUnitNumber As Long
Private mVIP As Boolean
Private mComments As String
Private mSponsor_IEN As String
Private mPCM_Code As String
Private mFacilityNCID As String
Private mBloodType As String
Private mRankNCID As String
Private mPatientFlag As String
Private mUnidentifiedFlag As String
Private mSponsorSrhString As String
Private mPatientSrhString As String
Private mAddresses As Addresses
Private mInsurances As Insurances
Private mState As enumObjectState
Private mEnterpiseUnitNumber As String
Private mDeersID As String
Private mADSFlag As Boolean

Private mVerifiedFlag As Boolean
Private mVerifiedSource As String
Private mVerifiedDate As Date
Private mVerifiedVersion As String

'******************************************************************************
'*  Patient Profile information
'******************************************************************************
Private m_ProfileImageFile            As String
Private m_ProfileThumbnailFile        As String
Private m_ProfileImageChanged         As Boolean
Public Property Get DeersID() As String
    DeersID = mDeersID
End Property
Public Property Let DeersID(ByVal value As String)
    mDeersID = value
End Property


Public Property Get EnterpiseUnitNumber() As String
    EnterpiseUnitNumber = mEnterpiseUnitNumber
End Property
Public Property Let EnterpiseUnitNumber(ByVal value As String)
    mEnterpiseUnitNumber = value
End Property


Public Property Get Insurances() As Insurances
    If mInsurances Is Nothing Then
        Set mInsurances = New Insurances
    End If
    
    Set Insurances = mInsurances
End Property
Public Property Set Insurances(ByRef value As Insurances)
    Set mInsurances = value
End Property


Public Property Get SponsorSrhString() As String
    SponsorSrhString = mSponsorSrhString
End Property
Public Property Let SponsorSrhString(ByVal value As String)
    mSponsorSrhString = value
End Property


Public Property Get PatientSrhString() As String
    PatientSrhString = mPatientSrhString
End Property
Public Property Let PatientSrhString(ByVal value As String)
    mPatientSrhString = value
End Property



Public Property Get UnidentifiedFlag() As String
    UnidentifiedFlag = mUnidentifiedFlag
End Property
Public Property Let UnidentifiedFlag(ByVal value As String)
    mUnidentifiedFlag = value
End Property


Public Property Get PatientFlag() As String
    PatientFlag = mPatientFlag
End Property
Public Property Let PatientFlag(ByVal value As String)
    mPatientFlag = value
End Property

Public Property Get RankNCID() As String
    RankNCID = mRankNCID
End Property
Public Property Let RankNCID(ByVal value As String)
    mRankNCID = value
End Property


Public Property Get BloodType() As String
    BloodType = mBloodType
End Property
Public Property Let BloodType(ByVal value As String)
    mBloodType = value
End Property


Public Property Get ACV_EnrollmentStatus() As String
    ACV_EnrollmentStatus = mACV_EnrollmentStatus
End Property
Public Property Let ACV_EnrollmentStatus(ByVal value As String)
    mACV_EnrollmentStatus = value
End Property
Public Property Get Adopted() As String
    Adopted = mAdopted
End Property
Public Property Let Adopted(ByVal value As String)
    mAdopted = value
End Property
Public Property Get Age() As String
    Age = mAge
End Property
Public Property Let Age(ByVal value As String)
    mAge = value
End Property
Public Property Get AssigningAuthority() As String
    AssigningAuthority = mAssigningAuthority
End Property
Public Property Let AssigningAuthority(ByVal value As String)
    mAssigningAuthority = value
End Property
Public Property Get BranchOfService() As String
    BranchOfService = mBranchOfService
End Property
Public Property Let BranchOfService(ByVal value As String)
    mBranchOfService = value
End Property
Public Property Get CommandInterest() As PatientGenericDatas
    If mCommandInterest Is Nothing Then
        Set mCommandInterest = New PatientGenericDatas
    End If
    
    Set CommandInterest = mCommandInterest
End Property
Public Property Set CommandInterest(ByVal value As PatientGenericDatas)
    Set mCommandInterest = value
End Property
Public Property Get CommandSecurity() As PatientGenericDatas
    If mCommandSecurity Is Nothing Then
        Set mCommandSecurity = New PatientGenericDatas
    End If
    
    Set CommandSecurity = mCommandSecurity
End Property
Public Property Set CommandSecurity(ByVal value As PatientGenericDatas)
    Set mCommandSecurity = value
End Property
Public Property Get DateOfBirth() As Double
    DateOfBirth = mDateOfBirth
End Property
Public Property Let DateOfBirth(ByVal value As Double)
    mDateOfBirth = value
End Property
Public Property Get DeersEligibility() As DeersEligibility
    If mDeersEligibility Is Nothing Then
        Set mDeersEligibility = New DeersEligibility
    End If
    
    Set DeersEligibility = mDeersEligibility
End Property
Public Property Set DeersEligibility(ByVal value As DeersEligibility)
    Set mDeersEligibility = value
End Property
Public Property Get DeployUnit() As String
    DeployUnit = mDeployUnit
End Property
Public Property Let DeployUnit(ByVal value As String)
    mDeployUnit = value
End Property
Public Property Get DeployLocation() As String
    DeployLocation = mDeployLocation
End Property
Public Property Let DeployLocation(ByVal value As String)
    mDeployLocation = value
End Property
Public Property Get DNBI() As String
    DNBI = mDNBI
End Property
Public Property Let DNBI(ByVal value As String)
    mDNBI = value
End Property
Public Property Get EmailAddress() As String
    EmailAddress = mEmailAddress
End Property
Public Property Let EmailAddress(ByVal value As String)
    mEmailAddress = value
End Property
Public Property Get EnrollmentFacility() As String
    EnrollmentFacility = mEnrollmentFacility
End Property
Public Property Let EnrollmentFacility(ByVal value As String)
    mEnrollmentFacility = value
End Property
Public Property Get Expired() As String
    Expired = mExpired
End Property
Public Property Let Expired(ByVal value As String)
    mExpired = value
End Property
Public Property Get Facility() As String
    Facility = mFacility
End Property
Public Property Let Facility(ByVal value As String)
    mFacility = value
End Property
Public Property Get FMP() As String
    FMP = mFMP
End Property
Public Property Let FMP(ByVal value As String)
    mFMP = value
End Property
Public Property Get FMPSponsorSSN() As String
    FMPSponsorSSN = mFMPSponsorSSN
End Property
Public Property Let FMPSponsorSSN(ByVal value As String)
    mFMPSponsorSSN = value
End Property
Public Property Get ForeignID_Flag() As Boolean
    ForeignID_Flag = mForeignID_Flag
End Property
Public Property Let ForeignID_Flag(ByVal value As Boolean)
    mForeignID_Flag = value
End Property
Public Property Get Formatted_SponsorSSN() As String
    Formatted_SponsorSSN = mFormatted_SponsorSSN
End Property
Public Property Let Formatted_SponsorSSN(ByVal value As String)
    mFormatted_SponsorSSN = value
End Property
Public Property Get Formatted_SSN() As String
    Formatted_SSN = mFormatted_SSN
End Property
Public Property Let Formatted_SSN(ByVal value As String)
    mFormatted_SSN = value
End Property
Public Property Get GenericDatas() As PatientGenericDatas
    If mGenericDatas Is Nothing Then
        Set mGenericDatas = New PatientGenericDatas
    End If
    
    Set GenericDatas = mGenericDatas
End Property
Public Property Set GenericDatas(ByRef value As PatientGenericDatas)
    Set mGenericDatas = value
End Property
Public Property Get Guarantor() As String
    Guarantor = mGuarantor
End Property
Public Property Let Guarantor(ByVal value As String)
    mGuarantor = value
End Property
Public Property Get MaidenName() As String
    MaidenName = mMaidenName
End Property
Public Property Let MaidenName(ByVal value As String)
    mMaidenName = value
End Property
Public Property Get MaritalStatus() As String
    MaritalStatus = mMaritalStatus
End Property
Public Property Let MaritalStatus(ByVal value As String)
    mMaritalStatus = value
End Property
Public Property Get MedicalRecordNumber() As String
    MedicalRecordNumber = mMedicalRecordNumber
End Property
Public Property Let MedicalRecordNumber(ByVal value As String)
    mMedicalRecordNumber = value
End Property
Public Property Get Medicare() As String
    Medicare = mMedicare
End Property
Public Property Let Medicare(ByVal value As String)
    mMedicare = value
End Property
Public Property Get MedicareCode() As String
    MedicareCode = mMedicareCode
End Property
Public Property Let MedicareCode(ByVal value As String)
    mMedicareCode = value
End Property
Public Property Get MilitaryStatus() As String
    MilitaryStatus = mMilitaryStatus
End Property
Public Property Let MilitaryStatus(ByVal value As String)
    mMilitaryStatus = value
End Property
Public Property Get MothersMaidenName() As String
    MothersMaidenName = mMothersMaidenName
End Property
Public Property Let MothersMaidenName(ByVal value As String)
    mMothersMaidenName = value
End Property
Public Property Get HomeAddress() As Address
    If mHomeAddress Is Nothing Then
        Set mHomeAddress = New Address
    End If
    
    Set HomeAddress = mHomeAddress
End Property
Public Property Set HomeAddress(ByVal value As Address)
    Set mHomeAddress = value
End Property

Public Property Get Addresses() As Addresses
    If mAddresses Is Nothing Then
        Set mAddresses = New Addresses
    End If
    
    Set Addresses = mAddresses
End Property
Public Property Set Addresses(ByVal value As Addresses)
    Set mAddresses = value
End Property

Public Property Get Phones() As Phones
    If mPhones Is Nothing Then
        Set mPhones = New Phones
    End If
    
    Set Phones = mPhones
End Property
Public Property Set Phones(ByRef value As Phones)
    Set mPhones = value
End Property

Public Property Get HomePhone() As Phone
    If mHomePhone Is Nothing Then
        Set mHomePhone = New Phone
    End If

    Set HomePhone = mHomePhone
End Property
Public Property Set HomePhone(ByRef value As Phone)
    Set mHomePhone = value
End Property
Public Property Get WorkPhone() As Phone
    If mWorkPhone Is Nothing Then
        Set mWorkPhone = New Phone
    End If
    
    Set WorkPhone = mWorkPhone
End Property
Public Property Set WorkPhone(ByRef value As Phone)
    Set mWorkPhone = value
End Property
Public Property Get OrganDonor() As String
    OrganDonor = mOrganDonor
End Property
Public Property Let OrganDonor(ByVal value As String)
    mOrganDonor = value
End Property
Public Property Get PatientCategory() As String
    PatientCategory = mPatientCategory
End Property
Public Property Let PatientCategory(ByVal value As String)
    mPatientCategory = value
End Property
Public Property Get Name() As Name
    If mName Is Nothing Then
        Set mName = New Name
    End If
    
    Set Name = mName
End Property
Public Property Set Name(ByRef value As Name)
    Set mName = value
End Property
Public Property Get PayGrade() As String
    PayGrade = mPayGrade
End Property
Public Property Let PayGrade(ByVal value As String)
    mPayGrade = value
End Property
Public Property Get PCM_IEN() As String
    PCM_IEN = mPCM_IEN
End Property
Public Property Let PCM_IEN(ByVal value As String)
    mPCM_IEN = value
End Property
Public Property Get Pregnancy() As Pregnancy
    If mPregnancy Is Nothing Then
        Set mPregnancy = New Pregnancy
    End If
    
    Set Pregnancy = mPregnancy
End Property
Public Property Set Pregnancy(ByVal value As Pregnancy)
    Set mPregnancy = value
End Property
Public Property Get PrimaryCareManager() As String
    PrimaryCareManager = mPrimaryCareManager
End Property
Public Property Let PrimaryCareManager(ByVal value As String)
    mPrimaryCareManager = value
End Property
Public Property Get Race() As String
    Race = mRace
End Property
Public Property Let Race(ByVal value As String)
    mRace = value
End Property
Public Property Get Rank() As String
    Rank = mRank
End Property
Public Property Let Rank(ByVal value As String)
    mRank = value
End Property
Public Property Get RecordsMaintainedAt() As String
    RecordsMaintainedAt = mRecordsMaintainedAt
End Property
Public Property Let RecordsMaintainedAt(ByVal value As String)
    mRecordsMaintainedAt = value
End Property
Public Property Get Religion() As String
    Religion = mReligion
End Property
Public Property Let Religion(ByVal value As String)
    mReligion = value
End Property
Public Property Get Sex() As String
    Sex = mSex
End Property
Public Property Let Sex(ByVal value As String)
    mSex = value
End Property
Public Property Get Sponsor() As Patient
    If mSponsor Is Nothing Then
        Set mSponsor = New Patient
    End If
    
    Set Sponsor = mSponsor
End Property
Public Property Set Sponsor(ByRef value As Patient)
    Set mSponsor = value
End Property
Public Property Get SponsorSSN() As String
    SponsorSSN = mSponsorSSN
End Property
Public Property Let SponsorSSN(ByVal value As String)
    mSponsorSSN = value
End Property
Public Property Get SponsorName() As String
    SponsorName = mSponsorName
End Property
Public Property Let SponsorName(ByVal value As String)
    mSponsorName = value
End Property
Public Property Get SSN() As String
    SSN = mSSN
End Property
Public Property Let SSN(ByVal value As String)
    mSSN = value
End Property
Public Property Get SpecialWorkStatus() As SpecialWorkStatus
    If mSpecialWorkStatus Is Nothing Then
        Set mSpecialWorkStatus = New SpecialWorkStatus
    End If
    
    Set SpecialWorkStatus = mSpecialWorkStatus
End Property
Public Property Set SpecialWorkStatus(ByRef value As SpecialWorkStatus)
    Set mSpecialWorkStatus = value
End Property
Public Property Get UnitIdentificationCode() As String
    UnitIdentificationCode = mUnitIdentificationCode
End Property
Public Property Let UnitIdentificationCode(ByVal value As String)
    mUnitIdentificationCode = value
End Property
Public Property Get UnitNumber() As Long
    UnitNumber = mUnitNumber
End Property
Public Property Let UnitNumber(ByVal value As Long)
    mUnitNumber = value
End Property
Public Property Get VIP() As Boolean
    VIP = mVIP
End Property
Public Property Let VIP(ByVal value As Boolean)
    mVIP = value
End Property

Public Property Get Comments() As String
    Comments = mComments
End Property
Public Property Let Comments(ByVal value As String)
    mComments = value
End Property


Public Property Get Sponsor_IEN() As String
    Sponsor_IEN = mSponsor_IEN
End Property
Public Property Let Sponsor_IEN(ByVal value As String)
    mSponsor_IEN = value
End Property

Public Property Get PCM_Code() As String
    PCM_Code = mPCM_Code
End Property
Public Property Let PCM_Code(ByVal value As String)
    mPCM_Code = value
End Property


Public Property Get FacilityNCID() As String
    FacilityNCID = mFacilityNCID
End Property
Public Property Let FacilityNCID(ByVal value As String)
    mFacilityNCID = value
End Property

Public Property Get ADSFlag() As Boolean
    ADSFlag = mADSFlag
End Property
Public Property Let ADSFlag(ByVal value As Boolean)
    mADSFlag = value
End Property

Public Property Get VerifiedFlag() As Boolean
    VerifiedFlag = mVerifiedFlag
End Property
Public Property Let VerifiedFlag(ByVal value As Boolean)
    mVerifiedFlag = value
End Property

Public Property Get VerifiedDate() As Date
    VerifiedDate = mVerifiedDate
End Property
Public Property Let VerifiedDate(ByVal value As Date)
    mVerifiedDate = value
End Property

Public Property Get VerifiedSource() As String
    VerifiedSource = mVerifiedSource
End Property
Public Property Let VerifiedSource(ByVal value As String)
    mVerifiedSource = value
End Property
 
Public Property Get VerifiedVersion() As String
    VerifiedVersion = mVerifiedVersion
End Property
Public Property Let VerifiedVersion(ByVal value As String)
    mVerifiedVersion = value
End Property

'Adding Patient Profile picture
Public Property Get ProfileImageFile() As String
    ProfileImageFile = m_ProfileImageFile
End Property
Public Property Let ProfileImageFile(value As String)
    m_ProfileImageFile = value
End Property

Public Property Get ProfileThumbnailFile() As String
    ProfileThumbnailFile = m_ProfileThumbnailFile
End Property
Public Property Let ProfileThumbnailFile(value As String)
    m_ProfileThumbnailFile = value
End Property

Public Property Get ProfileImageChanged() As Boolean
    ProfileImageChanged = m_ProfileImageChanged
End Property
Public Property Let ProfileImageChanged(value As Boolean)
    m_ProfileImageChanged = value
End Property

Private Function iObjectPersist_Copy() As Object
    Dim NewCopy As Patient
    Dim State As iState
    
    Dim Persist As ObjectPersist.iObjectPersist
    
    On Error GoTo ErrHandler
    
    Set NewCopy = New Patient
    
    With NewCopy
        .ACV_EnrollmentStatus = mACV_EnrollmentStatus
        .Adopted = mAdopted
        .Age = mAge
        .AssigningAuthority = mAssigningAuthority
        .BranchOfService = mBranchOfService
        .Sponsor_IEN = mSponsor_IEN
        
        Set Persist = mCommandInterest
        Set .CommandInterest = Persist.Copy

        Set Persist = mCommandSecurity
        Set .CommandSecurity = Persist.Copy

        .DateOfBirth = mDateOfBirth
        
        Set Persist = mDeersEligibility
        Set .DeersEligibility = Persist.Copy
        
        Set Persist = mInsurances
        Set .Insurances = Persist.Copy
        
        .DeployUnit = mDeployUnit
        .DeployLocation = mDeployLocation
        .DNBI = mDNBI
        .EmailAddress = mEmailAddress
        .EnrollmentFacility = mEnrollmentFacility
        .Expired = mExpired
        .Facility = mFacility
        .FMP = mFMP
        .FMPSponsorSSN = mFMPSponsorSSN
        .ForeignID_Flag = mForeignID_Flag
        .Formatted_SponsorSSN = mFormatted_SponsorSSN
        .Formatted_SSN = mFormatted_SSN
        
        Set Persist = mGenericDatas
        Set .GenericDatas = Persist.Copy

        .Guarantor = mGuarantor
        .MaidenName = mMaidenName
        .MaritalStatus = mMaritalStatus
        .MedicalRecordNumber = mMedicalRecordNumber
        .Medicare = mMedicare
        .MedicareCode = mMedicareCode
        .MilitaryStatus = mMilitaryStatus
        .MothersMaidenName = mMothersMaidenName
        
        Set Persist = mHomeAddress
        Set .HomeAddress = Persist.Copy
        
        Set Persist = mAddresses
        Set .Addresses = Persist.Copy
        
        Set Persist = mHomePhone
        Set .HomePhone = Persist.Copy
        
        Set Persist = mWorkPhone
        Set .WorkPhone = Persist.Copy
        
        Set Persist = mPhones
        Set .Phones = Persist.Copy
                
        .OrganDonor = mOrganDonor
        .PatientCategory = mPatientCategory
        
        Set Persist = mName
        Set .Name = Persist.Copy

        .PayGrade = mPayGrade
        .PCM_IEN = mPCM_IEN
        
        Set Persist = mPregnancy
        Set .Pregnancy = Persist.Copy

        .PrimaryCareManager = mPrimaryCareManager
        .Race = mRace
        .Rank = mRank
        .RecordsMaintainedAt = mRecordsMaintainedAt
        .Religion = mReligion
        .Sex = mSex
        
        Set Persist = mSponsor
        Set .Sponsor = Persist.Copy

        .SponsorSSN = mSponsorSSN
        .SponsorName = mSponsorName
        .SSN = mSSN
        
        Set Persist = mSpecialWorkStatus
        Set .SpecialWorkStatus = Persist.Copy

        .UnitIdentificationCode = mUnitIdentificationCode
        .UnitNumber = mUnitNumber
        .VIP = mVIP
        .Comments = mComments
        .PCM_Code = mPCM_Code
        .FacilityNCID = mFacilityNCID
        .BloodType = mBloodType
        .RankNCID = mRankNCID
        .PatientFlag = mPatientFlag
        .UnidentifiedFlag = mUnidentifiedFlag
        .PatientSrhString = mPatientSrhString
        .SponsorSrhString = mSponsorSrhString
        .EnterpiseUnitNumber = mEnterpiseUnitNumber
        .DeersID = mDeersID
        .ADSFlag = mADSFlag
        .VerifiedDate = mVerifiedDate
        .VerifiedFlag = mVerifiedFlag
        .VerifiedSource = mVerifiedSource
        .VerifiedVersion = mVerifiedVersion
        'iState interface
        Set State = NewCopy
        State.State = mState
    End With

    Set iObjectPersist_Copy = NewCopy

    Set NewCopy = Nothing
    Set State = Nothing
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "DTOsPatient.iObjectPersist_Persist", Err.Description, Err.HelpFile, Err.HelpContext

End Function




Private Sub iObjectPersist_Persist(ByRef Writer As iWriter)
    On Error GoTo ErrHandler

    With Writer
        .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION
        .WriteProperty "ACV_EnrollmentStatus", mACV_EnrollmentStatus
        .WriteProperty "Adopted", mAdopted
        .WriteProperty "Age", mAge
        .WriteProperty "AssigningAuthority", mAssigningAuthority
        .WriteProperty "BranchOfService", mBranchOfService
        .WriteProperty "Sponsor_IEN", mSponsor_IEN
        .WriteObject "CommandInterest", mCommandInterest
        .WriteObject "CommandSecurity", mCommandSecurity
        .WriteProperty "DateOfBirth", mDateOfBirth
        .WriteObject "DeersEligibility_", mDeersEligibility
        .WriteObject "Insurances_", mInsurances
        .WriteProperty "DeployUnit", mDeployUnit
        .WriteProperty "DeployLocation", mDeployLocation
        .WriteProperty "DNBI", mDNBI
        .WriteProperty "EmailAddress", mEmailAddress
        .WriteProperty "EnrollmentFacility", mEnrollmentFacility
        .WriteProperty "Expired", mExpired
        .WriteProperty "Facility", mFacility
        .WriteProperty "FMP", mFMP
        .WriteProperty "FMPSponsorSSN", mFMPSponsorSSN
        .WriteProperty "ForeignID_Flag", mForeignID_Flag
        .WriteProperty "Formatted_SponsorSSN", mFormatted_SponsorSSN
        .WriteProperty "Formatted_SSN", mFormatted_SSN
        .WriteObject "GenericDatas", mGenericDatas
        .WriteProperty "Guarantor", mGuarantor
        .WriteProperty "MaidenName", mMaidenName
        .WriteProperty "MaritalStatus", mMaritalStatus
        .WriteProperty "MedicalRecordNumber", mMedicalRecordNumber
        .WriteProperty "Medicare", mMedicare
        .WriteProperty "MedicareCode", mMedicareCode
        .WriteProperty "MilitaryStatus", mMilitaryStatus
        .WriteProperty "MothersMaidenName", mMothersMaidenName
        .WriteObject "HomeAddress", mHomeAddress
        .WriteObject "Addresses", mAddresses
        .WriteObject "HomePhone", mHomePhone
        .WriteObject "WorkPhone", mWorkPhone
        .WriteObject "Phones_", mPhones
        .WriteProperty "OrganDonor", mOrganDonor
        .WriteProperty "PatientCategory", mPatientCategory
        .WriteObject "Name_", mName
        .WriteProperty "PayGrade", mPayGrade
        .WriteProperty "PCM_IEN", mPCM_IEN
        .WriteObject "Pregnancy_", mPregnancy
        .WriteProperty "PrimaryCareManager", mPrimaryCareManager
        .WriteProperty "Race", mRace
        .WriteProperty "Rank", mRank
        .WriteProperty "RecordsMaintainedAt", mRecordsMaintainedAt
        .WriteProperty "Religion", mReligion
        .WriteProperty "Sex", mSex
        If FMP <> "20" Then
            .WriteObject "Sponsor", mSponsor
        End If
        .WriteProperty "SponsorSSN", mSponsorSSN
        .WriteProperty "SponsorName", mSponsorName
        .WriteProperty "SSN", mSSN
        .WriteObject "SpecialWorkStatus_", mSpecialWorkStatus
        .WriteProperty "UnitIdentificationCode", mUnitIdentificationCode
        .WriteProperty "UnitNumber", mUnitNumber
        .WriteProperty "VIP", mVIP
        .WriteProperty "Comments", mComments
        .WriteProperty "PCM_Code", mPCM_Code
        .WriteProperty "FacilityNCID", mFacilityNCID
        .WriteProperty "BloodType", mBloodType
        .WriteProperty "RankNCID", mRankNCID
        .WriteProperty "PatientFlag", mPatientFlag
        .WriteProperty "UnidentifiedFlag", mUnidentifiedFlag
        .WriteProperty "PatientSrhString", mPatientSrhString
        .WriteProperty "SponsorSrhString", mSponsorSrhString
        .WriteProperty "EnterpiseUnitNumber", mEnterpiseUnitNumber
        .WriteProperty "DeersID", mDeersID
        .WriteProperty "ADSFlag", mADSFlag
        .WriteProperty "VerifiedFlag", mVerifiedFlag
        .WriteProperty "VerifiedSource", mVerifiedSource
        .WriteProperty "VerifiedDate ", mVerifiedDate
        .WriteProperty "VerifiedVersion ", mVerifiedVersion
        
        'iState interface
        .WriteProperty "iState.State", mState
    End With
    
    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, "DTOsPatient.iObjectPersist_Persist", Err.Description, Err.HelpFile, Err.HelpContext
End Sub




Private Function iObjectPersist_Restore(ByRef Reader As iReader) As Boolean
Dim Failed As Boolean

    On Error GoTo ErrHandler
  
    With Reader
        mACV_EnrollmentStatus = .ReadProperty("ACV_EnrollmentStatus")
        mAdopted = .ReadProperty("Adopted")
        mAge = .ReadProperty("Age")
        mAssigningAuthority = .ReadProperty("AssigningAuthority")
        mBranchOfService = .ReadProperty("BranchOfService")
        mSponsor_IEN = .ReadProperty("Sponsor_IEN")
        .ReadObject "CommandInterest", CommandInterest
        .ReadObject "CommandSecurity", CommandSecurity
        mDateOfBirth = Val(.ReadProperty("DateOfBirth"))
        .ReadObject "DeersEligibility_", DeersEligibility
        .ReadObject "Insurances_", Insurances
        mDeployUnit = .ReadProperty("DeployUnit")
        mDeployLocation = .ReadProperty("DeployLocation")
        mDNBI = .ReadProperty("DNBI")
        mEmailAddress = .ReadProperty("EmailAddress")
        mEnrollmentFacility = .ReadProperty("EnrollmentFacility")
        mExpired = .ReadProperty("Expired")
        mFacility = .ReadProperty("Facility")
        mFMP = .ReadProperty("FMP")
        mFMPSponsorSSN = .ReadProperty("FMPSponsorSSN")
        mForeignID_Flag = .ReadProperty("ForeignID_Flag")
        mFormatted_SponsorSSN = .ReadProperty("Formatted_SponsorSSN")
        mFormatted_SSN = .ReadProperty("Formatted_SSN")
        .ReadObject "GenericDatas", GenericDatas
        mGuarantor = .ReadProperty("Guarantor")
        mMaidenName = .ReadProperty("MaidenName")
        mMaritalStatus = .ReadProperty("MaritalStatus")
        mMedicalRecordNumber = .ReadProperty("MedicalRecordNumber")
        mMedicare = .ReadProperty("Medicare")
        mMedicareCode = .ReadProperty("MedicareCode")
        mMilitaryStatus = .ReadProperty("MilitaryStatus")
        mMothersMaidenName = .ReadProperty("MothersMaidenName")
        .ReadObject "HomeAddress", HomeAddress
        .ReadObject "Addresses", Addresses
        .ReadObject "HomePhone", HomePhone
        .ReadObject "WorkPhone", WorkPhone
        .ReadObject "Phones_", Phones
        mOrganDonor = .ReadProperty("OrganDonor")
        mPatientCategory = .ReadProperty("PatientCategory")
        .ReadObject "Name_", Name
        mPayGrade = .ReadProperty("PayGrade")
        mPCM_IEN = .ReadProperty("PCM_IEN")
        .ReadObject "Pregnancy_", Pregnancy
        mPrimaryCareManager = .ReadProperty("PrimaryCareManager")
        mRace = .ReadProperty("Race")
        mRank = .ReadProperty("Rank")
        mRecordsMaintainedAt = .ReadProperty("RecordsMaintainedAt")
        mReligion = .ReadProperty("Religion")
        mSex = .ReadProperty("Sex")
        If FMP <> "20" Then
            .ReadObject "Sponsor", Sponsor
        Else
            Set mSponsor = Me
        End If
        mSponsorSSN = .ReadProperty("SponsorSSN")
        mSponsorName = .ReadProperty("SponsorName")
        mSSN = .ReadProperty("SSN")
        .ReadObject "SpecialWorkStatus_", SpecialWorkStatus
        mUnitIdentificationCode = .ReadProperty("UnitIdentificationCode")
        mUnitNumber = .ReadProperty("UnitNumber")
        mVIP = .ReadProperty("VIP")
        mComments = .ReadProperty("Comments")
        mPCM_Code = .ReadProperty("PCM_Code")
        mFacilityNCID = .ReadProperty("FacilityNCID")
        mBloodType = .ReadProperty("BloodType")
        mRankNCID = .ReadProperty("RankNCID")
        mPatientFlag = .ReadProperty("PatientFlag")
        mUnidentifiedFlag = .ReadProperty("UnidentifiedFlag")
        mPatientSrhString = .ReadProperty("PatientSrhString")
        mSponsorSrhString = .ReadProperty("SponsorSrhString")
        mEnterpiseUnitNumber = .ReadProperty("EnterpiseUnitNumber")
        mDeersID = .ReadProperty("DeersID")
        mADSFlag = .ReadProperty("ADSFlag")
        mVerifiedFlag = .ReadProperty("VerifiedFlag ")
        mVerifiedSource = .ReadProperty("VerifiedSource ")
        mVerifiedDate = .ReadProperty("VerifiedDate ")
        mVerifiedVersion = .ReadProperty("VerifiedVersion ")
        
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
