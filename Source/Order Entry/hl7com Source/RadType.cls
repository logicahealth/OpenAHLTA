VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RadType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

'Used to implement 2 interfaces - the basetablet and its own radtype
Implements BaseTablet

'Internal class object used in aggregation
Private oBaseTablet As BaseTablet


Private mvarDefaultAdminTime As String 'local copy
Private mvarDefaultDuration As Single 'local copy
Private mvarCounterSignatureRqd As hlCOUNTERSIGNREQUIRED 'local copy
Private mvarApprovalRqd As hlAPPROVALREQUIRED 'local copy
Private mvarAppointmentRqd As hlAPPOINTMENTREQUIRED 'local copy
Private mvarLocationIen As String 'local copy
Private mvarProcLocation As String 'local copy
Private mvarPortable As hlRADPORTABLE 'local copy
Private mvarContinuousAllowed As hlCONTINUOUSALLOWED 'local copy
Private mvarAllergyRqd As hlALLERGYREQUIRED  'local copy
Private mvarClinicalImpressionText As String 'local copy
Private mvarTransportMode As hlRADXPORT 'local copy
Private mvarWetReading As hlWETREADING 'local copy
          
Private mvarProcessPriority As hlLABPRIORITY 'local copy
 
Private mInPatientType As String
Public Property Let InPatientType(ByVal vData As String)
    mInPatientType = vData
End Property
Public Property Get InPatientType() As String
    InPatientType = mInPatientType
End Property

'****************************************************************************
'Properties that are RadType only
'   Read-only properties declare Let/Set properties as Friend
'****************************************************************************


Friend Property Let AllergyRqd(ByVal vData As hlALLERGYREQUIRED)
    mvarAllergyRqd = vData
End Property

Public Property Get AllergyRqd() As hlALLERGYREQUIRED
    AllergyRqd = mvarAllergyRqd
End Property

Friend Property Let AppointmentRqd(ByVal vData As hlAPPOINTMENTREQUIRED)
    mvarAppointmentRqd = vData
End Property

Public Property Get AppointmentRqd() As hlAPPOINTMENTREQUIRED
    AppointmentRqd = mvarAppointmentRqd
End Property
Friend Property Let ApprovalRqd(ByVal vData As hlAPPROVALREQUIRED)
    mvarApprovalRqd = vData
End Property

Public Property Get ApprovalRqd() As hlAPPROVALREQUIRED
    ApprovalRqd = mvarApprovalRqd
End Property
Public Property Let ClinicalImpressionText(ByVal vData As String)
    mvarClinicalImpressionText = vData
End Property

Public Property Get ClinicalImpressionText() As String
    ClinicalImpressionText = mvarClinicalImpressionText
End Property


Friend Property Let ContinuousAllowed(ByVal vData As hlCONTINUOUSALLOWED)
    mvarContinuousAllowed = vData
End Property

Public Property Get ContinuousAllowed() As hlCONTINUOUSALLOWED
    ContinuousAllowed = mvarContinuousAllowed
End Property
Friend Property Let CounterSignatureRqd(ByVal vData As hlCOUNTERSIGNREQUIRED)
    mvarCounterSignatureRqd = vData
End Property

Public Property Get CounterSignatureRqd() As hlCOUNTERSIGNREQUIRED
    CounterSignatureRqd = mvarCounterSignatureRqd
End Property
Friend Property Let DefaultDuration(ByVal vData As Single)
    mvarDefaultDuration = vData
End Property

Public Property Get DefaultDuration() As Single
    DefaultDuration = mvarDefaultDuration
End Property
Public Property Let DefaultAdminTime(ByVal vData As String)
    mvarDefaultAdminTime = vData
End Property

Public Property Get DefaultAdminTime() As String
    DefaultAdminTime = mvarDefaultAdminTime
End Property
Friend Property Let PortableAllowed(ByVal hlData As hlRADPORTABLE)
    mvarPortable = hlData
End Property

Public Property Get PortableAllowed() As hlRADPORTABLE
    PortableAllowed = mvarPortable
End Property

Public Property Let ProcedureLocation(ByVal vData As String)
    mvarProcLocation = vData
End Property

Public Property Get ProcedureLocation() As String
    ProcedureLocation = mvarProcLocation
End Property

Public Property Let ProcedureLocationIen(ByVal vData As String)
    mvarLocationIen = vData
End Property

Public Property Get ProcedureLocationIen() As String
    ProcedureLocationIen = mvarLocationIen
End Property
Public Property Let ProcessPriority(ByVal vData As hlLABPRIORITY)
    mvarProcessPriority = vData
End Property

Public Property Get ProcessPriority() As hlLABPRIORITY
    ProcessPriority = mvarProcessPriority
End Property
Public Property Let TransportMode(ByVal vData As hlRADXPORT)
    mvarTransportMode = vData
End Property

Public Property Get TransportMode() As hlRADXPORT
    TransportMode = mvarTransportMode
End Property
Public Property Let WetReadingWanted(ByVal vData As hlWETREADING)
    mvarWetReading = vData
End Property

Public Property Get WetReadingWanted() As hlWETREADING
    WetReadingWanted = mvarWetReading
End Property


'****************************************************************************
'Provide the basetablet interface for the radtype interface as well
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
Public Property Get LastSentNameWas() As String
    LastSentNameWas = oBaseTablet.LastSentNameWas
End Property
Public Property Let LastSentNameWas(szData As String)
    oBaseTablet.LastSentNameWas = szData
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
    BaseTablet_OrderTypeCode = oBaseTablet.OrderingProviderFname
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

Friend Property Get basetablet_WarningReason() As String
    basetablet_WarningReason = oBaseTablet.WarningReason
End Property
Friend Property Let basetablet_WarningOverriden(ByVal vData As Boolean)
    oBaseTablet.WarningOverriden = vData
End Property

Friend Property Get basetablet_WarningOverriden() As Boolean
    basetablet_WarningOverriden = oBaseTablet.WarningOverriden
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
End Sub

