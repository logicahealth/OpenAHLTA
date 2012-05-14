VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Medication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--- begin DTO Builder; DTO=Medication
'--- 4/1/2004 3:33:57 PM
'---
'--- NOTE: Any time you modify this class, increment the version number
'
'
Implements iObjectPersist

Private Const OBJECT_VERSION = "1.0"
Private Const UNIQUE_IDENTIFIER = "Medication"


Private mEventID As String  '
Private mDataID As String   '
Private mUnitNumber As String   '
Private mDataType As Integer    '
Private mEncounterFacility As ConceptValue  '
Private mEncounterNumber As String  '
Private mOrderStatus As ConceptValue    '
Private mFacility As ConceptValue   '
Private mStoringApplicationNCID As String   '
Private mPointOfCare As ConceptValue    '
Private mStoringFacilityNCID As String  '
Private mStoringUserNCID As String  '
Private mEventDate As Date  '
Private mPlacerNumber As String '
Private mOutPatient As Boolean  '
Private mCurrent As Boolean '
Private mExpired As Boolean '
Private mIVMed As Boolean    '
Private mOTC As Boolean '
Private mOrderingProvider As Provider   '
Private mEnteringProvider As Provider   '
Private mVerifyingProvider As Provider  '
Private mOrderStartDate As Date '
Private mOrderExpirationDate As Date    '
Private mTextComment As String  '
Private mOrderComment As String '
Private mCodedComment As String '
Private mProviderAdminInstructions As String    '
Private mPharmacySpecialInstructions As String  '
Private mDrugs As Drugs '
Private mRoute As ConceptValue  '
Private mQuantityUnits As ConceptValue  '
Private mRefills As Integer '
Private mRefillsRemaining As Integer    '
Private mRefillsActual As String    '
Private mSubstitutionStatus As ConceptValue '
Private mLastFilledDate As Date '
Private mDispensingLocation As ConceptValue '
Private mDaysOfSupply As Integer    '
Private mRXClinic As ConceptValue '
Private mFrequencies As ConceptValues    '
Private mTimingQuantityDisplay As String    '
Private mIVVolumeRate As String  '
Private mIVSpecialRate As String '
Private mOrderStatusDisplay As String   '
Private mValid As Boolean   '

Private Sub Class_Initialize()

On Error GoTo ErrHandler

  Set mEncounterFacility = New ConceptValue
  Set mOrderStatus = New ConceptValue
  Set mFacility = New ConceptValue
  Set mRoute = New ConceptValue
  Set mPointOfCare = New ConceptValue
  Set mQuantityUnits = New ConceptValue
  Set mSubstitutionStatus = New ConceptValue
  Set mDispensingLocation = New ConceptValue
  Set mFrequencies = New ConceptValues
  Set mOrderingProvider = New Provider
  Set mEnteringProvider = New Provider
  Set mVerifyingProvider = New Provider
  Set mRXClinic = New ConceptValue
  Set mDrugs = New Drugs
  
  Exit Sub
ErrHandler:
    Err.Raise Err.Number, "DTOsMedication.Class_Initialize", Err.Description, Err.HelpFile, Err.HelpContext

  
End Sub

Public Property Get EventID() As String
    EventID = mEventID
End Property
Public Property Let EventID(ByVal value As String)
    mEventID = value
End Property
Public Property Get DataID() As String
    DataID = mDataID
End Property
Public Property Let DataID(ByVal value As String)
    mDataID = value
End Property
Public Property Get UnitNumber() As String
    UnitNumber = mUnitNumber
End Property
Public Property Let UnitNumber(ByVal value As String)
    mUnitNumber = value
End Property
Public Property Get DataType() As Integer
    DataType = mDataType
End Property
Public Property Let DataType(ByVal value As Integer)
    mDataType = value
End Property
Public Property Get EncounterFacility() As ConceptValue
    Set EncounterFacility = mEncounterFacility
End Property
Public Property Set EncounterFacility(ByRef value As ConceptValue)
    Set mEncounterFacility = value
End Property
Public Property Get EncounterNumber() As String
    EncounterNumber = mEncounterNumber
End Property
Public Property Let EncounterNumber(ByVal value As String)
    mEncounterNumber = value
End Property
Public Property Get OrderStatus() As ConceptValue
    Set OrderStatus = mOrderStatus
End Property
Public Property Set OrderStatus(ByRef value As ConceptValue)
    Set mOrderStatus = value
End Property
Public Property Get Facility() As ConceptValue
    Set Facility = mFacility
End Property
Public Property Set Facility(ByRef value As ConceptValue)
    Set mFacility = value
End Property
Public Property Get StoringApplicationNCID() As String
    StoringApplicationNCID = mStoringApplicationNCID
End Property
Public Property Let StoringApplicationNCID(ByVal value As String)
    mStoringApplicationNCID = value
End Property
Public Property Get PointOfCare() As ConceptValue
    Set PointOfCare = mPointOfCare
End Property
Public Property Set PointOfCare(ByRef value As ConceptValue)
    Set mPointOfCare = value
End Property
Public Property Get StoringFacilityNCID() As String
    StoringFacilityNCID = mStoringFacilityNCID
End Property
Public Property Let StoringFacilityNCID(ByVal value As String)
    mStoringFacilityNCID = value
End Property
Public Property Get StoringUserNCID() As String
    StoringUserNCID = mStoringUserNCID
End Property
Public Property Let StoringUserNCID(ByVal value As String)
    mStoringUserNCID = value
End Property
Public Property Get EventDate() As Date
    EventDate = mEventDate
End Property
Public Property Let EventDate(ByVal value As Date)
    mEventDate = value
End Property
Public Property Get PlacerNumber() As String
    PlacerNumber = mPlacerNumber
End Property
Public Property Let PlacerNumber(ByVal value As String)
    mPlacerNumber = value
End Property
Public Property Get OutPatient() As Boolean
    OutPatient = mOutPatient
End Property
Public Property Let OutPatient(ByVal value As Boolean)
    mOutPatient = value
End Property
Public Property Get Current() As Boolean
    Current = mCurrent
End Property
Public Property Let Current(ByVal value As Boolean)
    mCurrent = value
End Property
Public Property Get Expired() As Boolean
    Expired = mExpired
End Property
Public Property Let Expired(ByVal value As Boolean)
    mExpired = value
End Property
Public Property Get IVMed() As Boolean
    IVMed = mIVMed
End Property
Public Property Let IVMed(ByVal value As Boolean)
    mIVMed = value
End Property
Public Property Get OTC() As Boolean
    OTC = mOTC
End Property
Public Property Let OTC(ByVal value As Boolean)
    mOTC = value
End Property
Public Property Get OrderingProvider() As Provider
    Set OrderingProvider = mOrderingProvider
End Property
Public Property Set OrderingProvider(ByRef value As Provider)
    Set mOrderingProvider = value
End Property
Public Property Get EnteringProvider() As Provider
    Set EnteringProvider = mEnteringProvider
End Property
Public Property Set EnteringProvider(ByRef value As Provider)
    Set mEnteringProvider = value
End Property
Public Property Get VerifyingProvider() As Provider
    Set VerifyingProvider = mVerifyingProvider
End Property
Public Property Set VerifyingProvider(ByRef value As Provider)
    Set mVerifyingProvider = value
End Property
Public Property Get OrderStartDate() As Date
    OrderStartDate = mOrderStartDate
End Property
Public Property Let OrderStartDate(ByVal value As Date)
    mOrderStartDate = value
End Property
Public Property Get OrderExpirationDate() As Date
    OrderExpirationDate = mOrderExpirationDate
End Property
Public Property Let OrderExpirationDate(ByVal value As Date)
    mOrderExpirationDate = value
End Property
Public Property Get TextComment() As String
    TextComment = mTextComment
End Property
Public Property Let TextComment(ByVal value As String)
    mTextComment = value
End Property
Public Property Get OrderComment() As String
    OrderComment = mOrderComment
End Property
Public Property Let OrderComment(ByVal value As String)
    mOrderComment = value
End Property
Public Property Get CodedComment() As String
    CodedComment = mCodedComment
End Property
Public Property Let CodedComment(ByVal value As String)
    mCodedComment = value
End Property
Public Property Get ProviderAdminInstructions() As String
    ProviderAdminInstructions = mProviderAdminInstructions
End Property
Public Property Let ProviderAdminInstructions(ByVal value As String)
    mProviderAdminInstructions = value
End Property
Public Property Get PharmacySpecialInstructions() As String
    PharmacySpecialInstructions = mPharmacySpecialInstructions
End Property
Public Property Let PharmacySpecialInstructions(ByVal value As String)
    mPharmacySpecialInstructions = value
End Property
Public Property Get Drugs() As Drugs
    Set Drugs = mDrugs
End Property
Public Property Set Drugs(ByRef value As Drugs)
    Set mDrugs = value
End Property
Public Property Get Route() As ConceptValue
    Set Route = mRoute
End Property
Public Property Set Route(ByRef value As ConceptValue)
    Set mRoute = value
End Property
Public Property Get QuantityUnits() As ConceptValue
    Set QuantityUnits = mQuantityUnits
End Property
Public Property Set QuantityUnits(ByRef value As ConceptValue)
    Set mQuantityUnits = value
End Property
Public Property Get Refills() As Integer
    Refills = mRefills
End Property
Public Property Let Refills(ByVal value As Integer)
    mRefills = value
End Property
Public Property Get RefillsRemaining() As Integer
    RefillsRemaining = mRefillsRemaining
End Property
Public Property Let RefillsRemaining(ByVal value As Integer)
    mRefillsRemaining = value
End Property
Public Property Get RefillsActual() As String
    RefillsActual = mRefillsActual
End Property
Public Property Let RefillsActual(ByVal value As String)
    mRefillsActual = value
End Property
Public Property Get SubstitutionStatus() As ConceptValue
    Set SubstitutionStatus = mSubstitutionStatus
End Property
Public Property Set SubstitutionStatus(ByRef value As ConceptValue)
    Set mSubstitutionStatus = value
End Property
Public Property Get LastFilledDate() As Date
    LastFilledDate = mLastFilledDate
End Property
Public Property Let LastFilledDate(ByVal value As Date)
    mLastFilledDate = value
End Property
Public Property Get DispensingLocation() As ConceptValue
    Set DispensingLocation = mDispensingLocation
End Property
Public Property Set DispensingLocation(ByRef value As ConceptValue)
    Set mDispensingLocation = value
End Property
Public Property Get DaysOfSupply() As Integer
    DaysOfSupply = mDaysOfSupply
End Property
Public Property Let DaysOfSupply(ByVal value As Integer)
    mDaysOfSupply = value
End Property

Public Property Get RXClinic() As ConceptValue
    Set RXClinic = mRXClinic
End Property
Public Property Set RXClinic(ByRef value As ConceptValue)
    Set mRXClinic = value
End Property

Public Property Get Frequencies() As DTOs.ConceptValues
    Set Frequencies = mFrequencies
End Property
Public Property Set Frequencies(ByRef value As DTOs.ConceptValues)
    Set mFrequencies = value
End Property

Public Property Get TimingQuantityDisplay() As String
    TimingQuantityDisplay = mTimingQuantityDisplay
End Property
Public Property Let TimingQuantityDisplay(ByVal value As String)
    mTimingQuantityDisplay = value
End Property
Public Property Get IVVolumeRate() As String
    IVVolumeRate = mIVVolumeRate
End Property
Public Property Let IVVolumeRate(ByVal value As String)
    mIVVolumeRate = value
End Property
Public Property Get IVSpecialRate() As String
    IVSpecialRate = mIVSpecialRate
End Property
Public Property Let IVSpecialRate(ByVal value As String)
    mIVSpecialRate = value
End Property
Public Property Get OrderStatusDisplay() As String
    OrderStatusDisplay = mOrderStatusDisplay
End Property
Public Property Let OrderStatusDisplay(ByVal value As String)
    mOrderStatusDisplay = value
End Property
Public Property Get Valid() As Boolean
    Valid = mValid
End Property
Public Property Let Valid(ByVal value As Boolean)
    mValid = value
End Property

Private Function iObjectPersist_Copy() As Object

On Error GoTo ErrHandler

    Dim NewCopy As Medication
    Dim Persist As iObjectPersist
    
    Set NewCopy = New Medication
    
    With NewCopy
        .EventID = mEventID
        .DataID = mDataID
        .UnitNumber = mUnitNumber
        .DataType = mDataType
        Set Persist = mEncounterFacility
        .EncounterFacility = Persist.Copy
        .EncounterNumber = mEncounterNumber
        Set Persist = mOrderStatus
        .OrderStatus = Persist.Copy
        Set Persist = mFacility
        .Facility = Persist.Copy
        .StoringApplicationNCID = mStoringApplicationNCID
        Set Persist = mPointOfCare
        .PointOfCare = Persist.Copy
        .StoringFacilityNCID = mStoringFacilityNCID
        .StoringUserNCID = mStoringUserNCID
        .EventDate = mEventDate
        .PlacerNumber = mPlacerNumber
        .OutPatient = mOutPatient
        .Current = mCurrent
        .Expired = mExpired
        .IVMed = mIVMed
        .OTC = mOTC
        Set Persist = mOrderingProvider
        .OrderingProvider = Persist.Copy
        Set Persist = mEnteringProvider
        .EnteringProvider = Persist.Copy
        Set Persist = mVerifyingProvider
        .VerifyingProvider = Persist.Copy
        .OrderStartDate = mOrderStartDate
        .OrderExpirationDate = mOrderExpirationDate
        .TextComment = mTextComment
        .OrderComment = mOrderComment
        .CodedComment = mCodedComment
        .ProviderAdminInstructions = mProviderAdminInstructions
        .PharmacySpecialInstructions = mPharmacySpecialInstructions
        Set Persist = mDrugs
        Set .Drugs = Persist.Copy
        Set Persist = mRoute
        Set .Route = Persist.Copy
        Set Persist = mQuantityUnits
        Set .QuantityUnits = Persist.Copy
        .Refills = mRefills
        .RefillsRemaining = mRefillsRemaining
        .RefillsActual = mRefillsActual
        Set Persist = mSubstitutionStatus
        Set .SubstitutionStatus = Persist.Copy
        .LastFilledDate = mLastFilledDate
        Set Persist = mDispensingLocation
        Set .DispensingLocation = Persist.Copy
        .DaysOfSupply = mDaysOfSupply
        Set Persist = mRXClinic
        Set Persist = mFrequencies
        Set .Frequencies = Persist.Copy
        .TimingQuantityDisplay = mTimingQuantityDisplay
        .IVVolumeRate = mIVVolumeRate
        .IVSpecialRate = mIVSpecialRate
        .OrderStatusDisplay = mOrderStatusDisplay
        .Valid = mValid
    End With

    Set iObjectPersist_Copy = NewCopy

Exit Function
ErrHandler:
    Err.Raise Err.Number, "DTOsMedication.iObjectPersist_Copy", Err.Description, Err.HelpFile, Err.HelpContext

End Function

Private Sub iObjectPersist_Persist(Writer As iWriter)

On Error GoTo ErrHandler

    With Writer
        .InitHeader UNIQUE_IDENTIFIER, OBJECT_VERSION
        .WriteProperty "EventID", mEventID
        .WriteProperty "DataID", mDataID
        .WriteProperty "UnitNumber", mUnitNumber
        .WriteProperty "DataType", mDataType
        .WriteObject "EncounterFacility", mEncounterFacility
        .WriteProperty "EncounterNumber", mEncounterNumber
        .WriteObject "OrderStatus", mOrderStatus
        .WriteObject "Facility", mFacility
        .WriteProperty "StoringApplicationNCID", mStoringApplicationNCID
        .WriteObject "PointOfCare", mPointOfCare
        .WriteProperty "StoringFacilityNCID", mStoringFacilityNCID
        .WriteProperty "StoringUserNCID", mStoringUserNCID
        .WriteProperty "EventDate", mEventDate
        .WriteProperty "PlacerNumber", mPlacerNumber
        .WriteProperty "OutPatient", mOutPatient
        .WriteProperty "Current", mCurrent
        .WriteProperty "Expired", mExpired
        .WriteProperty "IVMed", mIVMed
        .WriteProperty "OTC", mOTC
        .WriteObject "OrderingProvider", mOrderingProvider
        .WriteObject "EnteringProvider", mEnteringProvider
        .WriteObject "VerifyingProvider", mVerifyingProvider
        .WriteProperty "OrderStartDate", mOrderStartDate
        .WriteProperty "OrderExpirationDate", mOrderExpirationDate
        .WriteProperty "TextComment", mTextComment
        .WriteProperty "OrderComment", mOrderComment
        .WriteProperty "CodedComment", mCodedComment
        .WriteProperty "ProviderAdminInstructions", mProviderAdminInstructions
        .WriteProperty "PharmacySpecialInstructions", mPharmacySpecialInstructions
        .WriteObject "Drugs", mDrugs
        .WriteObject "Route", mRoute
        .WriteObject "QuantityUnits", mQuantityUnits
        .WriteProperty "Refills", mRefills
        .WriteProperty "RefillsRemaining", mRefillsRemaining
        .WriteProperty "RefillsActual", mRefillsActual
        .WriteObject "SubstitutionStatus", mSubstitutionStatus
        .WriteProperty "LastFilledDate", mLastFilledDate
        .WriteObject "DispensingLocation", mDispensingLocation
        .WriteProperty "DaysOfSupply", mDaysOfSupply
        .WriteObject "RXClinic", mRXClinic
        .WriteObject "Frequencies", mFrequencies
        .WriteProperty "TimingQuantityDisplay", mTimingQuantityDisplay
        .WriteProperty "IVVolumeRate", mIVVolumeRate
        .WriteProperty "IVSpecialRate", mIVSpecialRate
        .WriteProperty "OrderStatusDisplay", mOrderStatusDisplay
        .WriteProperty "Valid", mValid
    End With
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, "DTOsMedication.iObjectPersist_Persist", Err.Description, Err.HelpFile, Err.HelpContext

End Sub




Private Function iObjectPersist_Restore(Reader As iReader) As Boolean
Dim Failed As Boolean

    On Error GoTo ErrHandler
  
    With Reader
        mEventID = .ReadProperty("EventID")
        mDataID = .ReadProperty("DataID")
        mUnitNumber = .ReadProperty("UnitNumber")
        mDataType = .ReadProperty("DataType")
        .ReadObject "EncounterFacility", EncounterFacility
        mEncounterNumber = .ReadProperty("EncounterNumber")
        .ReadObject "OrderStatus", OrderStatus
        .ReadObject "Facility", Facility
        mStoringApplicationNCID = .ReadProperty("StoringApplicationNCID")
        .ReadObject "PointOfCare", PointOfCare
        mStoringFacilityNCID = .ReadProperty("StoringFacilityNCID")
        mStoringUserNCID = .ReadProperty("StoringUserNCID")
        mEventDate = .ReadProperty("EventDate")
        mPlacerNumber = .ReadProperty("PlacerNumber")
        mOutPatient = .ReadProperty("OutPatient")
        mCurrent = .ReadProperty("Current")
        mExpired = .ReadProperty("Expired")
        mIVMed = .ReadProperty("IVMed")
        mOTC = .ReadProperty("OTC")
        .ReadObject "OrderingProvider", OrderingProvider
        .ReadObject "EnteringProvider", EnteringProvider
        .ReadObject "VerifyingProvider", VerifyingProvider
        mOrderStartDate = .ReadProperty("OrderStartDate")
        mOrderExpirationDate = .ReadProperty("OrderExpirationDate")
        mTextComment = .ReadProperty("TextComment")
        mOrderComment = .ReadProperty("OrderComment")
        mCodedComment = .ReadProperty("CodedComment")
        mProviderAdminInstructions = .ReadProperty("ProviderAdminInstructions")
        mPharmacySpecialInstructions = .ReadProperty("PharmacySpecialInstructions")
        .ReadObject "Drugs", Drugs
        .ReadObject "Route", Route
        .ReadObject "QuantityUnits", QuantityUnits
        mRefills = .ReadProperty("Refills")
        mRefillsRemaining = .ReadProperty("RefillsRemaining")
        mRefillsActual = .ReadProperty("RefillsActual")
        .ReadObject "SubstitutionStatus", SubstitutionStatus
        mLastFilledDate = .ReadProperty("LastFilledDate")
        .ReadObject "DispensingLocation", DispensingLocation
        mDaysOfSupply = .ReadProperty("DaysOfSupply")
        .ReadObject "RXClinic", RXClinic
        .ReadObject "Frequencies", Frequencies
        mTimingQuantityDisplay = .ReadProperty("TimingQuantityDisplay")
        mIVVolumeRate = .ReadProperty("IVVolumeRate")
        mIVSpecialRate = .ReadProperty("IVSpecialRate")
        mOrderStatusDisplay = .ReadProperty("OrderStatusDisplay")
        mValid = .ReadProperty("Valid")
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

