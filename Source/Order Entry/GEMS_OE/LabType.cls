VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LabType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

'Used to implement 2 interfaces - the basetablet and its own labtype
Implements BaseTablet

'Internal class object used in aggregation
Private oBaseTablet As BaseTablet
Private mvarApprovalRequired As hlAPPROVALREQUIRED 'local copy
Private mvarMaxPriority As hlLABPRIORITY 'local copy
Private mvarSingleSpecimenIsComplete As hlSingleSpecimenIsComplete 'local copy
Private mvarPriority As hlLABPRIORITY 'local copy
Private mvarCollectMethod As hlLABCOLLECT  'local copy
Private mvarSpecimen As LabSpecimenType 'local copy
Private mvarAPSpecimenCollection As APSpecimenCollection 'local copy
Private mvarAPAutopsy As AutopsyType 'local copy
Private mvarAPComments As APCommentsType 'local copy
Private mvarAPCytGyn As CytologyGyn 'local copy

Private mvarLabWillCollect_Times As String 'local copy
Private mvarLabWillCollect_Specimen As String 'local copy
Private mvarLabWillCollect_SpecimenIen As String 'local copy
Private mvarSpecimenCollectiondate As String 'local copy
Private mvarLabCollectPriority As hlLABPRIORITY  'local copy
Private mvarSubscript As hlLABSUBSCRIPT 'local copy
Private mvarModifySpecimenZocIndex As Integer 'local copy
Private mvarOrderSetHasReqComments As Boolean 'local copy

Private mInPatientType As String
Private mProviderComment As String
Private mResultAssessment As Inpatient_Utilities.enumItemCritical

Public Property Let InPatientType(ByVal vData As String)
    mInPatientType = vData
End Property
Public Property Get InPatientType() As String
    InPatientType = mInPatientType
End Property

Public Property Let ProviderComment(ByVal vData As String)
    mProviderComment = vData
End Property
Public Property Get ProviderComment() As String
    ProviderComment = mProviderComment
End Property

Public Property Let ResultAssessment(ByVal vData As Inpatient_Utilities.enumItemCritical)
    mResultAssessment = vData
End Property
Public Property Get ResultAssessment() As Inpatient_Utilities.enumItemCritical
    ResultAssessment = mResultAssessment
End Property

Friend Property Let ModifySpecimenZocIndex(ByVal vData As Integer)
    mvarModifySpecimenZocIndex = vData
End Property
Public Property Get ModifySpecimenZocIndex() As Integer
    ModifySpecimenZocIndex = mvarModifySpecimenZocIndex
End Property
Friend Property Let OrderSetHasReqComments(ByVal vData As Boolean)
    mvarOrderSetHasReqComments = vData
End Property
Public Property Get OrderSetHasReqComments() As Boolean
    OrderSetHasReqComments = mvarOrderSetHasReqComments
End Property

Public Property Let SpecimenCollectionDate(ByVal vData As String)
    mvarSpecimenCollectiondate = vData
End Property

Public Property Get SpecimenCollectionDate() As String
    SpecimenCollectionDate = mvarSpecimenCollectiondate
End Property

Friend Property Let LabWillCollect_SpecimenIen(ByVal vData As String)
    mvarLabWillCollect_SpecimenIen = vData
End Property

Public Property Get LabWillCollect_SpecimenIen() As String
    LabWillCollect_SpecimenIen = mvarLabWillCollect_SpecimenIen
End Property
 
Friend Property Let LabWillCollect_Specimen(ByVal vData As String)
    mvarLabWillCollect_Specimen = vData
End Property

Public Property Get LabWillCollect_Specimen() As String
    LabWillCollect_Specimen = mvarLabWillCollect_Specimen
End Property

Friend Property Let LabWillCollect_Times(ByVal vData As String)
    mvarLabWillCollect_Times = vData
End Property

Public Property Get LabWillCollect_Times() As String
    LabWillCollect_Times = mvarLabWillCollect_Times
End Property

Public Property Set Specimen(ByVal vData As LabSpecimenType)
    Set mvarSpecimen = vData
End Property

Public Property Get Specimen() As LabSpecimenType
    Set Specimen = mvarSpecimen
End Property
Public Property Get APSpecimenColl() As APSpecimenCollection
    Set APSpecimenColl = mvarAPSpecimenCollection
End Property
Public Property Get APAutopsy() As AutopsyType
    Set APAutopsy = mvarAPAutopsy
End Property
Public Property Get APComments() As APCommentsType
    Set APComments = mvarAPComments
End Property
Public Property Set APSpecimenColl(ByVal vData As APSpecimenCollection)
    Set mvarAPSpecimenCollection = vData
End Property
Public Property Set APAutopsy(ByVal vData As AutopsyType)
    Set mvarAPAutopsy = vData
End Property
Public Property Set APComments(ByVal vData As APCommentsType)
    Set mvarAPComments = vData
End Property
Public Property Set APCytGyn(ByVal vData As CytologyGyn)
    Set mvarAPCytGyn = vData
End Property
Public Property Get APCytGyn() As CytologyGyn
    Set APCytGyn = mvarAPCytGyn
End Property



Public Property Let CollectMethod(ByVal vData As hlLABCOLLECT)
    mvarCollectMethod = vData
End Property

Public Property Get CollectMethod() As hlLABCOLLECT
    CollectMethod = mvarCollectMethod
End Property

Public Property Let ProcessPriority(ByVal vData As hlLABPRIORITY)
    mvarPriority = vData
End Property

Public Property Get ProcessPriority() As hlLABPRIORITY
    ProcessPriority = mvarPriority
End Property

Friend Property Let SingleSpecimenIsComplete(ByVal vData As hlSingleSpecimenIsComplete)
    mvarSingleSpecimenIsComplete = vData
End Property

Public Property Get SingleSpecimenIsComplete() As hlSingleSpecimenIsComplete
    SingleSpecimenIsComplete = mvarSingleSpecimenIsComplete
End Property

Friend Property Let MaxPriorityAllowed(ByVal vData As hlLABPRIORITY)
    mvarMaxPriority = vData
End Property
'This appears to apply to processing priority but not collection priority AMW 10/18/99
Public Property Get MaxPriorityAllowed() As hlLABPRIORITY
    MaxPriorityAllowed = mvarMaxPriority
End Property

Friend Property Let ApprovalRqd(ByVal vData As hlAPPROVALREQUIRED)
    mvarApprovalRequired = vData
End Property

Public Property Get ApprovalRqd() As hlAPPROVALREQUIRED
    ApprovalRqd = mvarApprovalRequired
End Property
Public Property Let LabCollectPriority(ByVal vData As hlLABPRIORITY)
    mvarLabCollectPriority = vData
End Property

Public Property Get LabCollectPriority() As hlLABPRIORITY
    LabCollectPriority = mvarLabCollectPriority
End Property
Public Property Let Subscript(ByVal vData As hlLABSUBSCRIPT)
   mvarSubscript = vData
End Property
Public Property Get Subscript() As hlLABSUBSCRIPT
    Subscript = mvarSubscript
End Property

'****************************************************************************
'Provide the basetablet interface for the labtype interface as well
'   This will expose these properties to the client
'   while delegating to oBaseTablet object which is created by this class'
'   initialization code.
'   Read-only properties declare Let/Set properties as Friend
'****************************************************************************

Public Property Let ActionReason(ByVal vData As String)
    oBaseTablet.ActionReason = vData
End Property

Public Property Get ActionReason() As String
    ActionReason = oBaseTablet.ActionReason
End Property

Public Property Let Comment(ByVal vData As String)
    oBaseTablet.Comment = vData
End Property

Public Property Get Comment() As String
    Comment = oBaseTablet.Comment
End Property

Friend Property Let ErrorIsMessage(ByVal vData As Boolean)
    oBaseTablet.ErrorIsMessage = vData
End Property

Public Property Get ErrorIsMessage() As Boolean
    ErrorIsMessage = oBaseTablet.ErrorIsMessage
End Property

Friend Property Let OrderId(ByVal vData As String)
    oBaseTablet.OrderId = vData
End Property

Public Property Get OrderId() As String
    OrderId = oBaseTablet.OrderId
End Property
Public Property Let OrderIen(ByVal vData As String)
    oBaseTablet.OrderIen = vData
End Property

Public Property Get OrderIen() As String
    OrderIen = oBaseTablet.OrderIen
End Property

Public Property Let OrderName(ByVal vData As String)
    oBaseTablet.OrderName = vData
End Property

Public Property Get OrderName() As String
    OrderName = oBaseTablet.OrderName
End Property
Friend Property Let OrderingProviderLname(ByVal vData As String)
    oBaseTablet.OrderingProviderLname = vData
End Property

Public Property Get OrderingProviderLname() As String
    OrderingProviderLname = oBaseTablet.OrderingProviderLname
End Property

Friend Property Let OrderingProviderIen(ByVal vData As String)
    oBaseTablet.OrderingProviderIen = vData
End Property

Public Property Get OrderingProviderIen() As String
    OrderingProviderIen = oBaseTablet.OrderingProviderIen
End Property

Friend Property Let OrderingProviderFname(ByVal vData As String)
    oBaseTablet.OrderingProviderFname = vData
End Property

Public Property Get OrderingProviderFname() As String
    OrderingProviderFname = oBaseTablet.OrderingProviderFname
End Property

Friend Property Set Timing(oTiming As TimingType)
    Set oBaseTablet.Timing = oTiming
End Property

Public Property Get Timing() As TimingType
    Set Timing = oBaseTablet.Timing
End Property

Public Property Let WarningReason(ByVal vData As String)
    oBaseTablet.WarningReason = vData
End Property
Public Property Let WarningOverriden(ByVal vData As Boolean)
    oBaseTablet.WarningOverriden = vData
End Property

Public Property Get WarningOverriden() As Boolean
    WarningOverriden = oBaseTablet.WarningOverriden
End Property
Public Property Get WarningReason() As String
    WarningReason = oBaseTablet.WarningReason
End Property

Friend Property Let WarningHeader(ByVal vData As String)
    oBaseTablet.WarningHeader = vData
End Property

Public Property Get WarningHeader() As String
    WarningHeader = oBaseTablet.WarningHeader
End Property

Friend Property Let LastSentNameWas(ByVal vData As String)
    oBaseTablet.LastSentNameWas = vData
End Property

Public Property Get LastSentNameWas() As String
    LastSentNameWas = oBaseTablet.LastSentNameWas
End Property

Friend Property Let LastSentIenWas(ByVal vData As String)
    oBaseTablet.LastSentIenWas = vData
End Property

Public Property Get LastSentIenWas() As String
    LastSentIenWas = oBaseTablet.LastSentIenWas
End Property

Friend Property Let LastPreverifyNameWas(ByVal vData As String)
    oBaseTablet.LastPreverifyNameWas = vData
End Property

Public Property Get LastPreverifyNameWas() As String
    LastPreverifyNameWas = oBaseTablet.LastPreverifyNameWas
End Property

Friend Property Let LastPreverifyNamesMatch(ByVal vData As Boolean)
    oBaseTablet.LastPreverifyNamesMatch = vData
End Property

Public Property Get LastPreverifyNamesMatch() As Boolean
    LastPreverifyNamesMatch = oBaseTablet.LastPreverifyNamesMatch
End Property

Friend Property Let LastPreverifyIenWas(ByVal vData As String)
    oBaseTablet.LastPreverifyIenWas = vData
End Property

Public Property Get LastPreverifyIenWas() As String
    LastPreverifyIenWas = oBaseTablet.LastPreverifyIenWas
End Property
Friend Property Let EnteredByIEN(ByVal vData As String)
    oBaseTablet.EnteredByIEN = vData
End Property

Public Property Get EnteredByIEN() As String
    EnteredByIEN = oBaseTablet.EnteredByIEN
End Property
Friend Property Let EnteredByName(ByVal vData As String)
    oBaseTablet.EnteredByName = vData
End Property

Public Property Get EnteredByName() As String
    EnteredByName = oBaseTablet.EnteredByName
End Property
'Added 12/17/2002 GG SCR 29495
Public Property Let ReactivationDate(ByVal vData As String)
    oBaseTablet.ReactivationDate = vData
End Property
'Added 12/17/2002 GG SCR 29495
Public Property Get ReactivationDate() As String
    ReactivationDate = oBaseTablet.ReactivationDate
End Property
'****************************************************************************
'Implementing the basetablet Class Interface
'   All properties are Friend to hide from client, but can be used within
'   the project.
'****************************************************************************

Friend Property Let basetablet_ActionReason(ByVal vData As String)
    oBaseTablet.ActionReason = vData
End Property

Friend Property Get basetablet_ActionReason() As String
    basetablet_ActionReason = oBaseTablet.ActionReason
End Property

Friend Property Let basetablet_Comment(ByVal vData As String)
    oBaseTablet.Comment = vData
End Property

Friend Property Get basetablet_Comment() As String
    basetablet_Comment = oBaseTablet.Comment
End Property

Friend Property Let basetablet_ErrorIsMessage(ByVal vData As Boolean)
    oBaseTablet.ErrorIsMessage = vData
End Property

Friend Property Get basetablet_ErrorIsMessage() As Boolean
    basetablet_ErrorIsMessage = oBaseTablet.ErrorIsMessage
End Property

Friend Property Let basetablet_OrderId(ByVal vData As String)
    oBaseTablet.OrderId = vData
End Property

Friend Property Get basetablet_OrderId() As String
    basetablet_OrderId = oBaseTablet.OrderId
End Property
Friend Property Let basetablet_OrderIen(ByVal vData As String)
    oBaseTablet.OrderIen = vData
End Property

Friend Property Get basetablet_OrderIen() As String
    basetablet_OrderIen = oBaseTablet.OrderIen
End Property

Friend Property Let basetablet_OrderName(ByVal vData As String)
    oBaseTablet.OrderName = vData
End Property

Friend Property Get basetablet_OrderName() As String
    basetablet_OrderName = oBaseTablet.OrderName
End Property
Friend Property Let basetablet_OrderingProviderLname(ByVal vData As String)
    oBaseTablet.OrderingProviderLname = vData
End Property

Friend Property Get basetablet_OrderingProviderLname() As String
    basetablet_OrderingProviderLname = oBaseTablet.OrderingProviderLname
End Property

Friend Property Let basetablet_OrderingProviderIen(ByVal vData As String)
    oBaseTablet.OrderingProviderIen = vData
End Property

Friend Property Get basetablet_OrderingProviderIen() As String
    basetablet_OrderingProviderIen = oBaseTablet.OrderingProviderIen
End Property

Friend Property Let basetablet_OrderingProviderFname(ByVal vData As String)
    oBaseTablet.OrderingProviderFname = vData
End Property

Friend Property Get basetablet_OrderingProviderFname() As String
    basetablet_OrderingProviderFname = oBaseTablet.OrderingProviderFname
End Property

Private Property Let BaseTablet_OrderTypeCode(ByVal RHS As hlORDERTYPECODE)
    oBaseTablet.OrderTypeCode = RHS
End Property

Private Property Get BaseTablet_OrderTypeCode() As hlORDERTYPECODE
    BaseTablet_OrderTypeCode = oBaseTablet.OrderTypeCode
End Property

Friend Property Set basetablet_Timing(oTiming As TimingType)
    Set oBaseTablet.Timing = oTiming
End Property

Friend Property Get basetablet_Timing() As TimingType
    Set basetablet_Timing = oBaseTablet.Timing
End Property

Friend Property Let basetablet_WarningReason(ByVal vData As String)
    oBaseTablet.WarningReason = vData
End Property

Friend Property Let basetablet_WarningOverriden(ByVal vData As Boolean)
    oBaseTablet.WarningOverriden = vData
End Property

Friend Property Get basetablet_WarningOverriden() As Boolean
    basetablet_WarningOverriden = oBaseTablet.WarningOverriden
End Property
Friend Property Get basetablet_WarningReason() As String
    basetablet_WarningReason = oBaseTablet.WarningReason
End Property

Friend Property Let basetablet_WarningHeader(ByVal vData As String)
    oBaseTablet.WarningHeader = vData
End Property

Friend Property Get basetablet_WarningHeader() As String
    basetablet_WarningHeader = oBaseTablet.WarningHeader
End Property

Friend Property Let basetablet_LastSentNameWas(ByVal vData As String)
    oBaseTablet.LastSentNameWas = vData
End Property

Public Property Get basetablet_LastSentNameWas() As String
    basetablet_LastSentNameWas = oBaseTablet.LastSentNameWas
End Property

Friend Property Let basetablet_LastSentIenWas(ByVal vData As String)
    oBaseTablet.LastSentIenWas = vData
End Property

Public Property Get basetablet_LastSentIenWas() As String
    basetablet_LastSentIenWas = oBaseTablet.LastSentIenWas
End Property

Friend Property Let basetablet_LastPreverifyNameWas(ByVal vData As String)
    oBaseTablet.LastPreverifyNameWas = vData
End Property

Friend Property Get basetablet_LastPreverifyNameWas() As String
    basetablet_LastPreverifyNameWas = oBaseTablet.LastPreverifyNameWas
End Property

Friend Property Let basetablet_LastPreverifyNamesMatch(ByVal vData As Boolean)
    oBaseTablet.LastPreverifyNamesMatch = vData
End Property

Friend Property Get basetablet_LastPreverifyNamesMatch() As Boolean
    basetablet_LastPreverifyNamesMatch = oBaseTablet.LastPreverifyNamesMatch
End Property

Friend Property Let basetablet_LastPreverifyIenWas(ByVal vData As String)
    oBaseTablet.LastPreverifyIenWas = vData
End Property

Friend Property Get basetablet_LastPreverifyIenWas() As String
    basetablet_LastPreverifyIenWas = oBaseTablet.LastPreverifyIenWas
End Property


Friend Property Let basetablet_EnteredByName(ByVal vData As String)
    oBaseTablet.EnteredByName = vData
End Property

Friend Property Get basetablet_EnteredByName() As String
    basetablet_EnteredByName = oBaseTablet.EnteredByName
End Property
Friend Property Let basetablet_EnteredByIEN(ByVal vData As String)
    oBaseTablet.EnteredByIEN = vData
End Property

Friend Property Get basetablet_EnteredByIEN() As String
    basetablet_EnteredByIEN = oBaseTablet.EnteredByIEN
End Property
'Added 12/17/2002 GG SCR 29495
Friend Property Let basetablet_ReactivationDate(ByVal vData As String)
    oBaseTablet.ReactivationDate = vData
End Property
'Added 12/17/2002 GG SCR 29495
Friend Property Get basetablet_ReactivationDate() As String
    basetablet_ReactivationDate = oBaseTablet.ReactivationDate
End Property
Private Sub Class_Initialize()
  Set oBaseTablet = New BaseTablet
  Set mvarSpecimen = New LabSpecimenType
  Set mvarAPComments = New APCommentsType
  Set mvarAPSpecimenCollection = New APSpecimenCollection
  Set mvarAPAutopsy = New AutopsyType
  Set mvarAPCytGyn = New CytologyGyn

  mvarSubscript = hlCHEMISTRY
End Sub



