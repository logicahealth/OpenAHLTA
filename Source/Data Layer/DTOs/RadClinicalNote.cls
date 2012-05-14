VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RadClinicalNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--- begin DTO Builder; DTO=RadClinicalNote
'--- 2/25/2004 10:08:52 AM
'---
'--- NOTE: Any time you modify this class, increment the version number
'
'

Implements iObjectPersist

Private Const OBJECT_VERSION = "1.0"
Private Const UNIQUE_IDENTIFIER = "RadClinicalNote"


Private mEventID As String  '
Private mApplicationNCID As Long    '
Private mDataTypeNCID As String '
Private mDataType As String '
Private mReportNCID As Long '
Private mPlacerNumber As String '
Private mReportNameOne As String    '
Private mReportNameTwo As String    '
Private mEventDate As Date  '
Private mExamDate As Date   '
Private mCreateDate As Date '
Private mModifyDate As Date '
Private mStatus As String   '
Private mStatusNCID As Long '
Private mAmended As Boolean '
Private mResultName As String   '
Private mResultCode As Long '
Private mProviderName As String '
Private mProviderIEN As Long    '
Private mProviderNCID As Long   '
Private mEncounterNumber As String  '
Private mEncounterFacilityNCID As Long  '
Private mEncounterFacility As String    '
Private mCreatedBy As String    '
Private mModifiedBy As String   '
Private mPriorityNCID As Long   '
Private mPriorityName As String '
Private mCaller As String   '
Private mClinician As String    '
Private mPOCFacility As String  '
Private mHTMLText As String '
Private mResultDisplay As Collection    '



Public Property Get EventID() As String
    EventID = mEventID
End Property
Public Property Let EventID(value As String)
    mEventID = value
End Property
Public Property Get ApplicationNCID() As Long
    ApplicationNCID = mApplicationNCID
End Property
Public Property Let ApplicationNCID(value As Long)
    mApplicationNCID = value
End Property
Public Property Get DataTypeNCID() As String
    DataTypeNCID = mDataTypeNCID
End Property
Public Property Let DataTypeNCID(value As String)
    mDataTypeNCID = value
End Property
Public Property Get DataType() As String
    DataType = mDataType
End Property
Public Property Let DataType(value As String)
    mDataType = value
End Property
Public Property Get ReportNCID() As Long
    ReportNCID = mReportNCID
End Property
Public Property Let ReportNCID(value As Long)
    mReportNCID = value
End Property
Public Property Get PlacerNumber() As String
    PlacerNumber = mPlacerNumber
End Property
Public Property Let PlacerNumber(value As String)
    mPlacerNumber = value
End Property
Public Property Get ReportNameOne() As String
    ReportNameOne = mReportNameOne
End Property
Public Property Let ReportNameOne(value As String)
    mReportNameOne = value
End Property
Public Property Get ReportNameTwo() As String
    ReportNameTwo = mReportNameTwo
End Property
Public Property Let ReportNameTwo(value As String)
    mReportNameTwo = value
End Property
Public Property Get EventDate() As Date
    EventDate = mEventDate
End Property
Public Property Let EventDate(value As Date)
    mEventDate = value
End Property
Public Property Get ExamDate() As Date
    ExamDate = mExamDate
End Property
Public Property Let ExamDate(value As Date)
    mExamDate = value
End Property
Public Property Get CreateDate() As Date
    CreateDate = mCreateDate
End Property
Public Property Let CreateDate(value As Date)
    mCreateDate = value
End Property
Public Property Get ModifyDate() As Date
    ModifyDate = mModifyDate
End Property
Public Property Let ModifyDate(value As Date)
    mModifyDate = value
End Property
Public Property Get Status() As String
    Status = mStatus
End Property
Public Property Let Status(value As String)
    mStatus = value
End Property
Public Property Get StatusNCID() As Long
    StatusNCID = mStatusNCID
End Property
Public Property Let StatusNCID(value As Long)
    mStatusNCID = value
End Property
Public Property Get Amended() As Boolean
    Amended = mAmended
End Property
Public Property Let Amended(value As Boolean)
    mAmended = value
End Property
Public Property Get ResultName() As String
    ResultName = mResultName
End Property
Public Property Let ResultName(value As String)
    mResultName = value
End Property
Public Property Get ResultCode() As Long
    ResultCode = mResultCode
End Property
Public Property Let ResultCode(value As Long)
    mResultCode = value
End Property
Public Property Get ProviderName() As String
    ProviderName = mProviderName
End Property
Public Property Let ProviderName(value As String)
    mProviderName = value
End Property
Public Property Get ProviderIEN() As Long
    ProviderIEN = mProviderIEN
End Property
Public Property Let ProviderIEN(value As Long)
    mProviderIEN = value
End Property
Public Property Get ProviderNCID() As Long
    ProviderNCID = mProviderNCID
End Property
Public Property Let ProviderNCID(value As Long)
    mProviderNCID = value
End Property
Public Property Get EncounterNumber() As String
    EncounterNumber = mEncounterNumber
End Property
Public Property Let EncounterNumber(value As String)
    mEncounterNumber = value
End Property
Public Property Get EncounterFacilityNCID() As Long
    EncounterFacilityNCID = mEncounterFacilityNCID
End Property
Public Property Let EncounterFacilityNCID(value As Long)
    mEncounterFacilityNCID = value
End Property
Public Property Get EncounterFacility() As String
    EncounterFacility = mEncounterFacility
End Property
Public Property Let EncounterFacility(value As String)
    mEncounterFacility = value
End Property
Public Property Get CreatedBy() As String
    CreatedBy = mCreatedBy
End Property
Public Property Let CreatedBy(value As String)
    mCreatedBy = value
End Property
Public Property Get ModifiedBy() As String
    ModifiedBy = mModifiedBy
End Property
Public Property Let ModifiedBy(value As String)
    mModifiedBy = value
End Property
Public Property Get PriorityNCID() As Long
    PriorityNCID = mPriorityNCID
End Property
Public Property Let PriorityNCID(value As Long)
    mPriorityNCID = value
End Property
Public Property Get PriorityName() As String
    PriorityName = mPriorityName
End Property
Public Property Let PriorityName(value As String)
    mPriorityName = value
End Property
Public Property Get Caller() As String
    Caller = mCaller
End Property
Public Property Let Caller(value As String)
    mCaller = value
End Property
Public Property Get Clinician() As String
    Clinician = mClinician
End Property
Public Property Let Clinician(value As String)
    mClinician = value
End Property
Public Property Get POCFacility() As String
    POCFacility = mPOCFacility
End Property
Public Property Let POCFacility(value As String)
    mPOCFacility = value
End Property
Public Property Get HTMLText() As String
    HTMLText = mHTMLText
End Property
Public Property Let HTMLText(value As String)
    mHTMLText = value
End Property
Public Property Get ResultDisplay() As Collection

  If mResultDisplay Is Nothing Then
    Set mResultDisplay = New Collection
  End If
  
  Set ResultDisplay = mResultDisplay

End Property
Public Property Set ResultDisplay(value As Collection)
    Set mResultDisplay = value
End Property


Private Function iObjectPersist_Copy() As Object
    Dim NewCopy As RadClinicalNote
    Dim i As Long
    
    Set NewCopy = New RadClinicalNote
    
    With NewCopy
        .EventID = mEventID
        .ApplicationNCID = mApplicationNCID
        .DataTypeNCID = mDataTypeNCID
        .DataType = mDataType
        .ReportNCID = mReportNCID
        .PlacerNumber = mPlacerNumber
        .ReportNameOne = mReportNameOne
        .ReportNameTwo = mReportNameTwo
        .EventDate = mEventDate
        .ExamDate = mExamDate
        .CreateDate = mCreateDate
        .ModifyDate = mModifyDate
        .Status = mStatus
        .StatusNCID = mStatusNCID
        .Amended = mAmended
        .ResultName = mResultName
        .ResultCode = mResultCode
        .ProviderName = mProviderName
        .ProviderIEN = mProviderIEN
        .ProviderNCID = mProviderNCID
        .EncounterNumber = mEncounterNumber
        .EncounterFacilityNCID = mEncounterFacilityNCID
        .EncounterFacility = mEncounterFacility
        .CreatedBy = mCreatedBy
        .ModifiedBy = mModifiedBy
        .PriorityNCID = mPriorityNCID
        .PriorityName = mPriorityName
        .Caller = mCaller
        .Clinician = mClinician
        .POCFacility = mPOCFacility
        .HTMLText = mHTMLText
        
        For i = 1 To mResultDisplay.Count
          .ResultDisplay.Add mResultDisplay.Item(i)
        Next
    End With

    Set iObjectPersist_Copy = NewCopy

End Function




Private Sub iObjectPersist_Persist(Writer As iWriter)
    With Writer
        .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION
        .WriteProperty "EventID", mEventID
        .WriteProperty "ApplicationNCID", mApplicationNCID
        .WriteProperty "DataTypeNCID", mDataTypeNCID
        .WriteProperty "DataType", mDataType
        .WriteProperty "ReportNCID", mReportNCID
        .WriteProperty "PlacerNumber", mPlacerNumber
        .WriteProperty "ReportNameOne", mReportNameOne
        .WriteProperty "ReportNameTwo", mReportNameTwo
        .WriteProperty "EventDate", mEventDate
        .WriteProperty "ExamDate", mExamDate
        .WriteProperty "CreateDate", mCreateDate
        .WriteProperty "ModifyDate", mModifyDate
        .WriteProperty "Status", mStatus
        .WriteProperty "StatusNCID", mStatusNCID
        .WriteProperty "Amended", mAmended
        .WriteProperty "ResultName", mResultName
        .WriteProperty "ResultCode", mResultCode
        .WriteProperty "ProviderName", mProviderName
        .WriteProperty "ProviderIEN", mProviderIEN
        .WriteProperty "ProviderNCID", mProviderNCID
        .WriteProperty "EncounterNumber", mEncounterNumber
        .WriteProperty "EncounterFacilityNCID", mEncounterFacilityNCID
        .WriteProperty "EncounterFacility", mEncounterFacility
        .WriteProperty "CreatedBy", mCreatedBy
        .WriteProperty "ModifiedBy", mModifiedBy
        .WriteProperty "PriorityNCID", mPriorityNCID
        .WriteProperty "PriorityName", mPriorityName
        .WriteProperty "Caller", mCaller
        .WriteProperty "Clinician", mClinician
        .WriteProperty "POCFacility", mPOCFacility
        .WriteProperty "HTMLText", mHTMLText
        .WriteCollection "ResultDisplay", mResultDisplay
    End With
End Sub




Private Function iObjectPersist_Restore(Reader As iReader) As Boolean
Dim Failed As Boolean

    On Error GoTo ErrHandler
  
    With Reader
        mEventID = .ReadProperty("EventID")
        mApplicationNCID = .ReadProperty("ApplicationNCID")
        mDataTypeNCID = .ReadProperty("DataTypeNCID")
        mDataType = .ReadProperty("DataType")
        mReportNCID = .ReadProperty("ReportNCID")
        mPlacerNumber = .ReadProperty("PlacerNumber")
        mReportNameOne = .ReadProperty("ReportNameOne")
        mReportNameTwo = .ReadProperty("ReportNameTwo")
        mEventDate = .ReadProperty("EventDate")
        mExamDate = .ReadProperty("ExamDate")
        mCreateDate = .ReadProperty("CreateDate")
        mModifyDate = .ReadProperty("ModifyDate")
        mStatus = .ReadProperty("Status")
        mStatusNCID = .ReadProperty("StatusNCID")
        mAmended = .ReadProperty("Amended")
        mResultName = .ReadProperty("ResultName")
        mResultCode = .ReadProperty("ResultCode")
        mProviderName = .ReadProperty("ProviderName")
        mProviderIEN = .ReadProperty("ProviderIEN")
        mProviderNCID = .ReadProperty("ProviderNCID")
        mEncounterNumber = .ReadProperty("EncounterNumber")
        mEncounterFacilityNCID = .ReadProperty("EncounterFacilityNCID")
        mEncounterFacility = .ReadProperty("EncounterFacility")
        mCreatedBy = .ReadProperty("CreatedBy")
        mModifiedBy = .ReadProperty("ModifiedBy")
        mPriorityNCID = .ReadProperty("PriorityNCID")
        mPriorityName = .ReadProperty("PriorityName")
        mCaller = .ReadProperty("Caller")
        mClinician = .ReadProperty("Clinician")
        mPOCFacility = .ReadProperty("POCFacility")
        mHTMLText = .ReadProperty("HTMLText")
        Set mResultDisplay = .ReadCollection("ResultDisplay")
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

