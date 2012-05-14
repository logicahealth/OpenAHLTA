VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Allergy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--- begin DTO Builder; DTO=Allergy
'--- 1/21/2004 6:07:03 PM
'---
'--- NOTE: Any time you modify this class, increment the version number
'
'

Implements iObjectPersist

Private Const OBJECT_VERSION = "1.0"
Private Const UNIQUE_IDENTIFIER = "Allergy"


Private mIENNumber As String
Private mApplication As String  '
Private mNotedDate As String    '
Private mClinician As String    '
Private mClinicianNCID As String    '
Private mPointOfCareFacility As String  '
Private mPointOfCareFacilityNCID As String  '
Private mEncounterFacility As String    '
Private mEncounterFacilityNCID As String    '
Private mEncounterNumber As String  '
Private mAllergyType As String  '
Private mAllergyTypeNCID As String  '
Private mAllergen As String '
Private mAllergenNCID As String '
Private mAllergyReactions As AllergyReactions '
Private mInfoSource As String   '
Private mInfoSourceNCID As String   '
Private mComment As String  '
Private mCreateUser As String   '
Private mCreateUserNCID As String   '
Private mCreateTime As String   '
Private mModified As Integer    '
Private mModifyUser As String   '
Private mModifyUserNCID As String   '
Private mModifyTime As String   '
Private mFacility As String '
Private mFacilityNCID As String '
Private mDataID As String   '
Private mNewRec As Integer  '
Private mModsFlg As Integer '
Private mCHCS_ID As String  '
Private mCHCSII_ENTRY As String '
Private mCHCS_FACILITIES() As String   '
Private mCDRDataID As String    '
Private mUnitNumber As String   '
Private mUpDateFlag As String   '
Private mTmip_Status As String  '

'--- end init / terminate procedures

'--- begin property lets/gets
Public Property Get Application() As String
    Application = mApplication
End Property
Public Property Let Application(value As String)
    mApplication = value
End Property
Public Property Get NotedDate() As String
    NotedDate = mNotedDate
End Property
Public Property Let NotedDate(value As String)
    mNotedDate = value
End Property
Public Property Get Clinician() As String
    Clinician = mClinician
End Property
Public Property Let Clinician(value As String)
    mClinician = value
End Property
Public Property Get ClinicianNCID() As String
    ClinicianNCID = mClinicianNCID
End Property
Public Property Let ClinicianNCID(value As String)
    mClinicianNCID = value
End Property
Public Property Get PointOfCareFacility() As String
    PointOfCareFacility = mPointOfCareFacility
End Property
Public Property Get IENNumber() As String
    IENNumber = mIENNumber
End Property
Public Property Let IENNumber(value As String)
    mIENNumber = value
End Property
Public Property Let PointOfCareFacility(value As String)
    mPointOfCareFacility = value
End Property
Public Property Get PointOfCareFacilityNCID() As String
    PointOfCareFacilityNCID = mPointOfCareFacilityNCID
End Property
Public Property Let PointOfCareFacilityNCID(value As String)
    mPointOfCareFacilityNCID = value
End Property
Public Property Get EncounterFacility() As String
    EncounterFacility = mEncounterFacility
End Property
Public Property Let EncounterFacility(value As String)
    mEncounterFacility = value
End Property
Public Property Get EncounterFacilityNCID() As String
    EncounterFacilityNCID = mEncounterFacilityNCID
End Property
Public Property Let EncounterFacilityNCID(value As String)
    mEncounterFacilityNCID = value
End Property
Public Property Get EncounterNumber() As String
    EncounterNumber = mEncounterNumber
End Property
Public Property Let EncounterNumber(value As String)
    mEncounterNumber = value
End Property
Public Property Get AllergyType() As String
    AllergyType = mAllergyType
End Property
Public Property Let AllergyType(value As String)
    mAllergyType = value
End Property
Public Property Get AllergyTypeNCID() As String
    AllergyTypeNCID = mAllergyTypeNCID
End Property
Public Property Let AllergyTypeNCID(value As String)
    mAllergyTypeNCID = value
End Property
Public Property Get Allergen() As String
    Allergen = mAllergen
End Property
Public Property Let Allergen(value As String)
    mAllergen = value
End Property
Public Property Get AllergenNCID() As String
    AllergenNCID = mAllergenNCID
End Property
Public Property Let AllergenNCID(value As String)
    mAllergenNCID = value
End Property
Public Property Get AllergyReactions() As AllergyReactions
    Set AllergyReactions = mAllergyReactions
End Property
Public Property Set AllergyReactions(value As AllergyReactions)
    Set mAllergyReactions = value
End Property
Public Property Get InfoSource() As String
    InfoSource = mInfoSource
End Property
Public Property Let InfoSource(value As String)
    mInfoSource = value
End Property
Public Property Get InfoSourceNCID() As String
    InfoSourceNCID = mInfoSourceNCID
End Property
Public Property Let InfoSourceNCID(value As String)
    mInfoSourceNCID = value
End Property
Public Property Get Comment() As String
    Comment = mComment
End Property
Public Property Let Comment(value As String)
    mComment = value
End Property
Public Property Get CreateUser() As String
    CreateUser = mCreateUser
End Property
Public Property Let CreateUser(value As String)
    mCreateUser = value
End Property
Public Property Get CreateUserNCID() As String
    CreateUserNCID = mCreateUserNCID
End Property
Public Property Let CreateUserNCID(value As String)
    mCreateUserNCID = value
End Property
Public Property Get CreateTime() As String
    CreateTime = mCreateTime
End Property
Public Property Let CreateTime(value As String)
    mCreateTime = value
End Property
Public Property Get Modified() As Integer
    Modified = mModified
End Property
Public Property Let Modified(value As Integer)
    mModified = value
End Property
Public Property Get ModifyUser() As String
    ModifyUser = mModifyUser
End Property
Public Property Let ModifyUser(value As String)
    mModifyUser = value
End Property
Public Property Get ModifyUserNCID() As String
    ModifyUserNCID = mModifyUserNCID
End Property
Public Property Let ModifyUserNCID(value As String)
    mModifyUserNCID = value
End Property
Public Property Get ModifyTime() As String
    ModifyTime = mModifyTime
End Property
Public Property Let ModifyTime(value As String)
    mModifyTime = value
End Property
Public Property Get Facility() As String
    Facility = mFacility
End Property
Public Property Let Facility(value As String)
    mFacility = value
End Property
Public Property Get FacilityNCID() As String
    FacilityNCID = mFacilityNCID
End Property
Public Property Let FacilityNCID(value As String)
    mFacilityNCID = value
End Property
Public Property Get DataID() As String
    DataID = mDataID
End Property
Public Property Let DataID(value As String)
    mDataID = value
End Property
Public Property Get NewRec() As Integer
    NewRec = mNewRec
End Property
Public Property Let NewRec(value As Integer)
    mNewRec = value
End Property
Public Property Get ModsFlg() As Integer
    ModsFlg = mModsFlg
End Property
Public Property Let ModsFlg(value As Integer)
    mModsFlg = value
End Property
Public Property Get CHCS_ID() As String
    CHCS_ID = mCHCS_ID
End Property
Public Property Let CHCS_ID(value As String)
    mCHCS_ID = value
End Property
Public Property Get CHCSII_ENTRY() As String
    CHCSII_ENTRY = mCHCSII_ENTRY
End Property
Public Property Let CHCSII_ENTRY(value As String)
    mCHCSII_ENTRY = value
End Property
Public Property Get CHCS_FACILITIES(iIndex As Long) As String
    CHCS_FACILITIES = mCHCS_FACILITIES(iIndex)
End Property
Public Property Let CHCS_FACILITIES(iIndex As Long, value As String)
    On Error GoTo ErrHandler
    If (iIndex) > UBound(mCHCS_FACILITIES) Then
ErrHandler:
        ReDim Preserve mCHCS_FACILITIES(0 To iIndex)
    End If
    mCHCS_FACILITIES(iIndex) = value
End Property
Public Property Get CDRDataID() As String
    CDRDataID = mCDRDataID
End Property
Public Property Let CDRDataID(value As String)
    mCDRDataID = value
End Property
Public Property Get UnitNumber() As String
    UnitNumber = mUnitNumber
End Property
Public Property Let UnitNumber(value As String)
    mUnitNumber = value
End Property
Public Property Get UpDateFlag() As String
    UpDateFlag = mUpDateFlag
End Property
Public Property Let UpDateFlag(value As String)
    mUpDateFlag = value
End Property
Public Property Get Tmip_Status() As String
    Tmip_Status = mTmip_Status
End Property
Public Property Let Tmip_Status(value As String)
    mTmip_Status = value
End Property

'--- end property lets/gets

Public Function CHCS_FACILITIES_Count() As Long
    CHCS_FACILITIES_Count = 0
    On Error GoTo ErrHandler
    CHCS_FACILITIES_Count = UBound(mCHCS_FACILITIES) + 1
ErrHandler:
    Err.Clear
End Function


Private Sub Class_Initialize()
    Erase mCHCS_FACILITIES
    Set mAllergyReactions = New AllergyReactions
End Sub

Private Function iObjectPersist_Copy() As Object
    Dim iIndex As Long
    Dim NewCopy As Allergy
    Set NewCopy = New Allergy
    Dim Persist As ObjectPersist.iObjectPersist
    With NewCopy
        .IENNumber = mIENNumber
        .Application = mApplication
        .NotedDate = mNotedDate
        .Clinician = mClinician
        .ClinicianNCID = mClinicianNCID
        .PointOfCareFacility = mPointOfCareFacility
        .PointOfCareFacilityNCID = mPointOfCareFacilityNCID
        .EncounterFacility = mEncounterFacility
        .EncounterFacilityNCID = mEncounterFacilityNCID
        .EncounterNumber = mEncounterNumber
        .AllergyType = mAllergyType
        .AllergyTypeNCID = mAllergyTypeNCID
        .Allergen = mAllergen
        .AllergenNCID = mAllergenNCID
        
        Set Persist = mAllergyReactions
        Set .AllergyReactions = Persist.Copy
        
        .InfoSource = mInfoSource
        .InfoSourceNCID = mInfoSourceNCID
        .Comment = mComment
        .CreateUser = mCreateUser
        .CreateUserNCID = mCreateUserNCID
        .CreateTime = mCreateTime
        .Modified = mModified
        .ModifyUser = mModifyUser
        .ModifyUserNCID = mModifyUserNCID
        .ModifyTime = mModifyTime
        .Facility = mFacility
        .FacilityNCID = mFacilityNCID
        .DataID = mDataID
        .NewRec = mNewRec
        .ModsFlg = mModsFlg
        .CHCS_ID = mCHCS_ID
        .CHCSII_ENTRY = mCHCSII_ENTRY
        
        On Error GoTo ErrHandler
        For iIndex = 0 To UBound(mCHCS_FACILITIES)
            .CHCS_FACILITIES(iIndex) = mCHCS_FACILITIES(iIndex)
        Next
ErrHandler:
        Err.Clear
        On Error Resume Next
        
        .CDRDataID = mCDRDataID
        .UnitNumber = mUnitNumber
        .UpDateFlag = mUpDateFlag
        .Tmip_Status = mTmip_Status
    End With

    Set iObjectPersist_Copy = NewCopy

End Function
'--- end copy procedure




Private Sub iObjectPersist_Persist(Writer As iWriter)
    With Writer
        .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION
        .WriteProperty "Application", mApplication
        .WriteProperty "NotedDate", mNotedDate
        .WriteProperty "Clinician", mClinician
        .WriteProperty "ClinicianNCID", mClinicianNCID
        .WriteProperty "PointOfCareFacility", mPointOfCareFacility
        .WriteProperty "PointOfCareFacilityNCID", mPointOfCareFacilityNCID
        .WriteProperty "EncounterFacility", mEncounterFacility
        .WriteProperty "EncounterFacilityNCID", mEncounterFacilityNCID
        .WriteProperty "EncounterNumber", mEncounterNumber
        .WriteProperty "IENNumber", mIENNumber
        .WriteProperty "AllergyType", mAllergyType
        .WriteProperty "AllergyTypeNCID", mAllergyTypeNCID
        .WriteProperty "Allergen", mAllergen
        .WriteProperty "AllergenNCID", mAllergenNCID
        .WriteObject "AllergyReactions", mAllergyReactions
        .WriteProperty "InfoSource", mInfoSource
        .WriteProperty "InfoSourceNCID", mInfoSourceNCID
        .WriteProperty "Comment", mComment
        .WriteProperty "CreateUser", mCreateUser
        .WriteProperty "CreateUserNCID", mCreateUserNCID
        .WriteProperty "CreateTime", mCreateTime
        .WriteProperty "Modified", mModified
        .WriteProperty "ModifyUser", mModifyUser
        .WriteProperty "ModifyUserNCID", mModifyUserNCID
        .WriteProperty "ModifyTime", mModifyTime
        .WriteProperty "Facility", mFacility
        .WriteProperty "FacilityNCID", mFacilityNCID
        .WriteProperty "DataId", mDataID
        .WriteProperty "NewRec", mNewRec
        .WriteProperty "ModsFlg", mModsFlg
        .WriteProperty "CHCS_ID", mCHCS_ID
        .WriteProperty "CHCSII_ENTRY", mCHCSII_ENTRY
        .WriteArray "CHCS_FACILITIES", mCHCS_FACILITIES
        .WriteProperty "CDRDataID", mCDRDataID
        .WriteProperty "UnitNumber", mUnitNumber
        .WriteProperty "UpDateFlag", mUpDateFlag
        .WriteProperty "Tmip_Status", mTmip_Status
    End With
End Sub
'--- end persist procedure




Private Function iObjectPersist_Restore(Reader As iReader) As Boolean
Dim Failed As Boolean

    On Error GoTo ErrHandler
  
    With Reader
        mIENNumber = .ReadProperty("IENNumber")
        mApplication = .ReadProperty("Application")
        mNotedDate = .ReadProperty("NotedDate")
        mClinician = .ReadProperty("Clinician")
        mClinicianNCID = .ReadProperty("ClinicianNCID")
        mPointOfCareFacility = .ReadProperty("PointOfCareFacility")
        mPointOfCareFacilityNCID = .ReadProperty("PointOfCareFacilityNCID")
        mEncounterFacility = .ReadProperty("EncounterFacility")
        mEncounterFacilityNCID = .ReadProperty("EncounterFacilityNCID")
        mEncounterNumber = .ReadProperty("EncounterNumber")
        mAllergyType = .ReadProperty("AllergyType")
        mAllergyTypeNCID = .ReadProperty("AllergyTypeNCID")
        mAllergen = .ReadProperty("Allergen")
        mAllergenNCID = .ReadProperty("AllergenNCID")
        .ReadObject "AllergyReactions", AllergyReactions
        'mAllergyReactions = .ReadObject("AllergyReactions")
        mInfoSource = .ReadProperty("InfoSource")
        mInfoSourceNCID = .ReadProperty("InfoSourceNCID")
        mComment = .ReadProperty("Comment")
        mCreateUser = .ReadProperty("CreateUser")
        mCreateUserNCID = .ReadProperty("CreateUserNCID")
        mCreateTime = .ReadProperty("CreateTime")
        mModified = .ReadProperty("Modified")
        mModifyUser = .ReadProperty("ModifyUser")
        mModifyUserNCID = .ReadProperty("ModifyUserNCID")
        mModifyTime = .ReadProperty("ModifyTime")
        mFacility = .ReadProperty("Facility")
        mFacilityNCID = .ReadProperty("FacilityNCID")
        mDataID = .ReadProperty("DataId")
        mNewRec = .ReadProperty("NewRec")
        mModsFlg = .ReadProperty("ModsFlg")
        mCHCS_ID = .ReadProperty("CHCS_ID")
        mCHCSII_ENTRY = .ReadProperty("CHCSII_ENTRY")
        mCHCS_FACILITIES = .ReadArray("CHCS_FACILITIES")
        mCDRDataID = .ReadProperty("CDRDataID")
        mUnitNumber = .ReadProperty("UnitNumber")
        mUpDateFlag = .ReadProperty("UpDateFlag")
        mTmip_Status = .ReadProperty("Tmip_Status")
    End With

    iObjectPersist_Restore = Not Failed
  
    Exit Function
  
ErrHandler:
    Failed = True
    Resume Next

End Function
'--- end restore procedure


Private Property Get iObjectPersist_UniqueIdentifier() As String
    iObjectPersist_UniqueIdentifier = UNIQUE_IDENTIFIER
End Property

Private Property Get iObjectPersist_Version() As String
    iObjectPersist_Version = OBJECT_VERSION
End Property
'--- end DTO Builder


