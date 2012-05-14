VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RxType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

'Used to implement 2 interfaces - the basetablet and its own rxtype
Implements BaseTablet

'Internal class object used in aggregation
Private oBaseTablet As BaseTablet

'local variable(s) to hold property value(s)
Private mvarDrugIsValid As Boolean  'local copy
Private mvarPharmacyName As String 'local copy
Private mvarPharmacyIen As String 'local copy
Private mvarMaxRefills As String 'local copy
Private mvarDefaultSig As String 'local copy
Private mvarDefaultUnit As String 'local copy
Private mvarMaxQuantity As String 'local copy
Private mvarMaxDays As String 'local copy
Private mvarNoteTo As String 'local copy
Private mvarCost As String 'local copy
Private mvarDispenseFull As hlDISPENSEFULL 'local copy
Private mvarDefaultQuantity As String 'local copy
Private mvarDefaultExpire As String 'local copy
Private mvarDefaultDaysSupply As String 'local copy
Private mvarSigCode As String 'local copy
Private mvarQuantity As String 'local copy
Private mvarRefills As String 'local copy
Private mvarChildCap As Long 'local copy
Private mvarOrderDate As String 'local copy
Private mvarExpandedSig As String 'local copy
Private mvarUnit As String 'local copy
Private mvarStatus As String 'local copy
Private mvarDefaultPharmacyIen As String 'local copy
Private mvarDefaultPharmacyName As String 'local copy
Private mvarLastSentNamesMatch As Boolean
Private mvarWarningPDTSOutcome As hlPDTSOUTCOMES
Private mvarPriority As Integer
Private mvarDispensingLocationNCID As String

'****************************************************************************
'Properties that are rxType only
'   Read-only properties declare Let/Set properties as Friend
'****************************************************************************

Public Property Let ChildCap(ByVal vData As hlCHILDCAP)
    mvarChildCap = vData
End Property

Public Property Get ChildCap() As hlCHILDCAP
    ChildCap = mvarChildCap
End Property

Friend Property Let Cost(ByVal vData As String)
    mvarCost = vData
End Property

Public Property Get Cost() As String
    Cost = mvarCost
End Property

Friend Property Let DefaultDaysSupply(ByVal vData As String)
    mvarDefaultDaysSupply = vData
End Property

Public Property Get DefaultDaysSupply() As String
    DefaultDaysSupply = mvarDefaultDaysSupply
End Property

Friend Property Let DefaultExpire(ByVal vData As String)
    mvarDefaultExpire = vData
End Property

Public Property Get DefaultExpire() As String
    DefaultExpire = mvarDefaultExpire
End Property

Friend Property Let DefaultQuantity(ByVal vData As String)
    mvarDefaultQuantity = vData
End Property

Public Property Get DefaultQuantity() As String
    DefaultQuantity = mvarDefaultQuantity
End Property
Friend Property Let DefaultPharmacyName(ByVal vData As String)
    mvarDefaultPharmacyName = vData
End Property

Public Property Get DefaultPharmacyName() As String
    DefaultPharmacyName = mvarDefaultPharmacyName
End Property

Friend Property Let DefaultPharmacyIen(ByVal vData As String)
    mvarDefaultPharmacyIen = vData
End Property

Public Property Get DefaultPharmacyIen() As String
    DefaultPharmacyIen = mvarDefaultPharmacyIen
End Property
Friend Property Let DefaultUnit(ByVal vData As String)
    mvarDefaultUnit = vData
End Property

Public Property Get DefaultUnit() As String
    DefaultUnit = mvarDefaultUnit
End Property

Friend Property Let DefaultSig(ByVal vData As String)
    mvarDefaultSig = vData
End Property

Public Property Get DefaultSig() As String
    DefaultSig = mvarDefaultSig
End Property
Friend Property Let DispenseFull(ByVal vData As hlDISPENSEFULL)
    mvarDispenseFull = vData
End Property

Public Property Get DispenseFull() As hlDISPENSEFULL
    DispenseFull = mvarDispenseFull
End Property
Friend Property Let DrugIsValid(ByVal vData As Boolean)
    mvarDrugIsValid = vData
End Property

Public Property Get DrugIsValid() As Boolean
    DrugIsValid = mvarDrugIsValid
End Property
Friend Property Let ExpandedSig(ByVal vData As String)
    mvarExpandedSig = vData
End Property

Public Property Get ExpandedSig() As String
    ExpandedSig = mvarExpandedSig
End Property
Friend Property Let MaxDaysSupply(ByVal vData As String)
    mvarMaxDays = vData
End Property

Public Property Get MaxDaysSupply() As String
    MaxDaysSupply = mvarMaxDays
End Property

Friend Property Let MaxQuantity(ByVal vData As String)
    mvarMaxQuantity = vData
End Property

Public Property Get MaxQuantity() As String
    MaxQuantity = mvarMaxQuantity
End Property

Friend Property Let MaxRefills(ByVal vData As String)
    mvarMaxRefills = vData
End Property

Public Property Get MaxRefills() As String
    MaxRefills = mvarMaxRefills
End Property
Friend Property Let NoteToProvider(ByVal vData As String)
    mvarNoteTo = vData
End Property

Public Property Get NoteToProvider() As String
    NoteToProvider = mvarNoteTo
End Property
Public Property Let OrderDate(ByVal szData As String)
     On Error GoTo ErrorHandler:
     If IsDate(szData) Then
          mvarOrderDate = Format$(szData, "yyyymmdd")
     Else
          mvarOrderDate = szData
     End If
     Exit Property
ErrorHandler:
     mvarOrderDate = szData
End Property

Public Property Get OrderDate() As String
    OrderDate = mvarOrderDate
End Property
Public Property Let PharmacyIenSent(ByVal vData As String)
    mvarPharmacyIen = vData
End Property

Public Property Get PharmacyIenSent() As String
    PharmacyIenSent = mvarPharmacyIen
End Property

Public Property Let PharmacyNameSent(ByVal vData As String)
    mvarPharmacyName = vData
End Property

Public Property Get PharmacyNameSent() As String
    PharmacyNameSent = mvarPharmacyName
End Property
Public Property Let Quantity(ByVal vData As String)
    mvarQuantity = vData
End Property

Public Property Get Quantity() As String
    Quantity = mvarQuantity
End Property
Public Property Let Refills(ByVal vData As String)
    mvarRefills = vData
End Property

Public Property Get Refills() As String
    Refills = mvarRefills
End Property
Public Property Let SigCode(ByVal szData As String)
    mvarSigCode = szData
End Property

Public Property Get SigCode() As String
    SigCode = mvarSigCode
End Property
Friend Property Let Status(ByVal vData As String)
    mvarStatus = vData
End Property
Public Property Get Status() As String
    Status = mvarStatus
End Property

Friend Property Let Unit(ByVal vData As String)
    mvarUnit = vData
End Property

Public Property Get Unit() As String
    Unit = mvarUnit
End Property

'Only RX gives the orderid/ordername on ACK from new order (RXO-1)
Friend Property Let LastSentNamesMatch(ByVal vData As Boolean)
    mvarLastSentNamesMatch = vData
End Property
Public Property Get LastSentNamesMatch() As Boolean
    LastSentNamesMatch = mvarLastSentNamesMatch
End Property
Public Function CopyTablet() As RxType
     
     Set CopyTablet = New RxType

     With CopyTablet
          .ActionReason = Me.ActionReason
          .ChildCap = Me.ChildCap
          .Comment = Me.Comment
          .Cost = Me.Cost
          .DefaultDaysSupply = Me.DefaultDaysSupply
          .DefaultExpire = Me.DefaultExpire
          .DefaultPharmacyIen = Me.DefaultPharmacyIen
          .DefaultPharmacyName = Me.DefaultPharmacyName
          .DefaultQuantity = Me.DefaultQuantity
          .DefaultSig = Me.DefaultSig
          .DefaultUnit = Me.DefaultUnit
          .DispenseFull = Me.DispenseFull
          .DrugIsValid = Me.DrugIsValid
          .ErrorIsMessage = Me.ErrorIsMessage
          .ExpandedSig = Me.ExpandedSig
          .LastPreverifyIenWas = Me.LastPreverifyIenWas
          .LastPreverifyNamesMatch = Me.LastPreverifyNamesMatch
          .LastPreverifyNameWas = Me.LastPreverifyNameWas
          .LastSentIenWas = Me.LastSentIenWas
          .LastSentNamesMatch = Me.LastSentNamesMatch
          .LastSentNameWas = Me.LastSentNameWas
          .MaxDaysSupply = Me.MaxDaysSupply
          .MaxQuantity = Me.MaxQuantity
          .MaxRefills = Me.MaxRefills
          .NoteToProvider = Me.NoteToProvider
          .OrderDate = Me.OrderDate
          .OrderId = Me.OrderId
          .OrderIen = Me.OrderIen
          .OrderingProviderFname = Me.OrderingProviderFname
          .OrderingProviderIen = Me.OrderingProviderIen
          .OrderingProviderLname = Me.OrderingProviderLname
          .EnteredByIEN = Me.EnteredByIEN
          .EnteredByName = Me.EnteredByName
          .OrderName = Me.OrderName
          .PharmacyIenSent = Me.PharmacyIenSent
          .PharmacyNameSent = Me.PharmacyNameSent
          .Quantity = Me.Quantity
          .Refills = Me.Refills
          .SigCode = Me.SigCode
          .Status = Me.Status
          With .Timing
               .AdminTime = Me.Timing.AdminTime
               .Duration = Me.Timing.Duration
               '.DurationUnits = Me.Timing.DurationUnits    'Currently readonly property
               .Frequency = Me.Timing.Frequency
               .OnexOrContinued = Me.Timing.OnexOrContinued
               .StartDate = Me.Timing.StartDate
               .StartTime = Me.Timing.StartTime
               .StopDate = Me.Timing.StopDate
               .StopTime = Me.Timing.StopTime
               .UnexpandedTime = Me.Timing.UnexpandedTime
          End With
          .Unit = Me.Unit
          .WarningHeader = Me.WarningHeader
          .WarningReason = Me.WarningReason
          .WarningOverriden = Me.WarningOverriden
          .WarningPDTSOutcome = Me.WarningPDTSOutcome
          .Priority = Me.Priority
          .DispensingLocationNCID = Me.DispensingLocationNCID
     End With
     
End Function

Public Sub DebugPrint()
          Debug.Print ; "*******************************************"
          
          Debug.Print "ActionReason: " & Me.ActionReason
          Debug.Print "ChildCap: " & Me.ChildCap
          Debug.Print "Comment: " & Me.Comment
          Debug.Print "Cost: " & Me.Cost
          Debug.Print "DefaultDaysSupply: " & Me.DefaultDaysSupply
          Debug.Print "DefaultExpire: " & Me.DefaultExpire
          Debug.Print "DefaultPharmacyIen: " & Me.DefaultPharmacyIen
          Debug.Print "DefaultPharmacyName: " & Me.DefaultPharmacyName
          Debug.Print "DefaultQuantity: " & Me.DefaultQuantity
          Debug.Print "DefaultSig: " & Me.DefaultSig
          Debug.Print "DefaultUnit: " & Me.DefaultUnit
          Debug.Print "DispenseFull: ", Me.DispenseFull
          Debug.Print "DrugIsValid: " & Me.DrugIsValid
          Debug.Print "ErrorIsMessage: " & Me.ErrorIsMessage
          Debug.Print "ExpandedSig: " & Me.ExpandedSig
          Debug.Print "LastPreverifyIenWas: " & Me.LastPreverifyIenWas
          Debug.Print "LastPreverifyNamesMatch: " & Me.LastPreverifyNamesMatch
          Debug.Print "LastPreverifyNameWas: " & Me.LastPreverifyNameWas
          Debug.Print "LastSentIenWas: " & Me.LastSentIenWas
          Debug.Print "LastSentNamesMatch: " & Me.LastSentNamesMatch
          Debug.Print "LastSentNameWas: " & Me.LastSentNameWas
          Debug.Print "MaxDaysSupply: " & Me.MaxDaysSupply
          Debug.Print "MaxQuantity: " & Me.MaxQuantity
          Debug.Print "MaxRefills: " & Me.MaxRefills
          Debug.Print "NoteToProvider: " & Me.NoteToProvider
          Debug.Print "OrderDate: " & Me.OrderDate
          Debug.Print "OrderId: " & Me.OrderId
          Debug.Print "OrderIen: " & Me.OrderIen
          Debug.Print "OrderingProviderFname: " & Me.OrderingProviderFname
          Debug.Print "OrderingProviderIen: " & Me.OrderingProviderIen
          Debug.Print "OrderingProviderLname: " & Me.OrderingProviderLname
          Debug.Print "OrderName: " & Me.OrderName
          Debug.Print "PharmacyIenSent: " & Me.PharmacyIenSent
          Debug.Print "PharmacyNameSent: " & Me.PharmacyNameSent
          Debug.Print "Quantity: " & Me.Quantity
          Debug.Print "Refills: " & Me.Refills
          Debug.Print "SigCode: " & Me.SigCode
          Debug.Print "Status: " & Me.Status
          With Me.Timing
               Debug.Print "Timing."
               Debug.Print vbTab & "AdminTime: " & Me.Timing.AdminTime
               Debug.Print vbTab & "Duration: " & Me.Timing.Duration
               Debug.Print vbTab & "DurationUnits: " & Me.Timing.DurationUnits
               Debug.Print vbTab & "Frequency: " & Me.Timing.Frequency
               Debug.Print vbTab & "OnexOrContinued: " & Me.Timing.OnexOrContinued
               Debug.Print vbTab & "OnexOrContinued: " & Me.Timing.StartDate
               Debug.Print vbTab & "StartTime: " & Me.Timing.StartTime
               Debug.Print vbTab & "StopDate: " & Me.Timing.StopDate
               Debug.Print vbTab & "StopTime: " & Me.Timing.StopTime
               Debug.Print vbTab & "UnexpandedTime: " & Me.Timing.UnexpandedTime
          End With
          Debug.Print "Unit: " & Me.Unit
          Debug.Print "WarningHeader: " & Me.WarningHeader
          Debug.Print "WarningReason: " & Me.WarningReason
          Debug.Print "Priority: " & Me.Priority
          Debug.Print "Dispensing Location NCID" & Me.DispensingLocationNCID
          Debug.Print ; "*******************************************"

End Sub

'****************************************************************************
'Provide the basetablet interface for the rxtype interface as well
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

Public Property Let WarningPDTSOutcome(ByVal vData As hlPDTSOUTCOMES)
    mvarWarningPDTSOutcome = vData
End Property
Public Property Get WarningPDTSOutcome() As hlPDTSOUTCOMES
    WarningPDTSOutcome = mvarWarningPDTSOutcome
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

Friend Property Get basetablet_LastSentNameWas() As String
    basetablet_LastSentNameWas = oBaseTablet.LastSentNameWas
End Property

Friend Property Let basetablet_LastSentIenWas(ByVal vData As String)
    oBaseTablet.LastSentIenWas = vData
End Property

Friend Property Get basetablet_LastSentIenWas() As String
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

Friend Property Let Priority(ByVal NewValue As Integer)
    mvarPriority = NewValue
End Property

Public Property Get Priority() As Integer
    Priority = mvarPriority
End Property

Friend Property Let DispensingLocationNCID(ByVal NewValue As String)
    mvarDispensingLocationNCID = NewValue
End Property

Public Property Get DispensingLocationNCID() As String
    DispensingLocationNCID = mvarDispensingLocationNCID
End Property

Private Sub Class_Initialize()
  Set oBaseTablet = New BaseTablet
End Sub
