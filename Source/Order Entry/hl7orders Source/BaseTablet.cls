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
Attribute VB_Name = "BaseTablet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'This is an abstract class used to declare functions that need to accept/use
'either LabType, RxType, or RadType objects without declaring these objects as
'an object type.  This class will implement the members that are common to the Rx,Lab and RadType
'classes.
'
'The drive behind eliminating declarations as Object type is that Friend Properties
'cannot be used with a late-bound (declared as object) object.  Friend properties are
'desired to allow us to make read-only properties.
'
'Now, instead of using the 'as object' declaration, we will use 'as BaseTablet' declaration.
'So, the objects passed must implement the BaseTablet interface (use the 'implements' keyword).
'This will be the Rad,Rx,and LabType objects.  The objects can then be used as a BaseTablet
'object, where the common members can be accessed, or assigned to an Rx, Lab, or RadTablet
'object to access all its members.
'AMW 9/6/99



Private mvarOrderId As String 'local copy
Private mvarWarningHeader As String 'local copy
Private mvarWarningOverriden As Boolean
Private mvarErrorIsMessage As Boolean 'local copy
Private mvarWarningReason As String 'local copy
Private mvarTiming As TimingType 'local copy
Private mvarTestName As String 'local copy
Private mvarTestIen As String 'local copy
Private mvarComment As String 'local copy
Private mvarOrderingProvider As String 'local copy
Private mvarOrderingProviderIen As String 'local copy
Private mvarEnteredByName As String 'local copy
Private mvarEnteredByIen As String 'local copy

Private mvarActionReason As String 'local copy
Private mvarOrderingProviderLname As String 'local copy
Private mvarLastSentIenWas As String 'local copy
Private mvarLastSentNameWas As String 'local copy
Private mvarLastPreverifyIenWas As String 'local copy
Private mvarLastPreverifyNamesMatch As Boolean 'local copy
Private mvarLastPreverifyNameWas As String 'local copy
Private mvarOrderTypeCode As hlORDERTYPECODE 'needed for cancelling non-labs, rads, rx in signing orders module
Private mvarReactivationDate As String 'local copy added 12/17/2002 by GG for SCR 29495


'*************************************
'The BaseTablet Class Interface
'*************************************

Public Property Let ActionReason(ByVal vData As String)
    mvarActionReason = vData
End Property

Public Property Get ActionReason() As String
    ActionReason = mvarActionReason
End Property

Public Property Let Comment(ByVal vData As String)
    mvarComment = vData
End Property

Public Property Get Comment() As String
    Comment = mvarComment
End Property

Public Property Let ErrorIsMessage(ByVal vData As Boolean)
    mvarErrorIsMessage = vData
End Property

Public Property Get ErrorIsMessage() As Boolean
    ErrorIsMessage = mvarErrorIsMessage
End Property

Public Property Let OrderId(ByVal vData As String)
    mvarOrderId = vData
End Property

Public Property Get OrderId() As String
    OrderId = mvarOrderId
End Property
Public Property Let OrderIen(ByVal vData As String)
    mvarTestIen = vData
End Property

Public Property Get OrderIen() As String
    OrderIen = mvarTestIen
End Property

Public Property Let OrderName(ByVal vData As String)
    mvarTestName = vData
End Property

Public Property Get OrderName() As String
    OrderName = mvarTestName
End Property
Public Property Let OrderingProviderLname(ByVal vData As String)
    mvarOrderingProviderLname = vData
End Property

Public Property Get OrderingProviderLname() As String
    OrderingProviderLname = mvarOrderingProviderLname
End Property

Public Property Let OrderingProviderIen(ByVal vData As String)
    mvarOrderingProviderIen = vData
End Property

Public Property Get OrderingProviderIen() As String
    OrderingProviderIen = mvarOrderingProviderIen
End Property

Public Property Let OrderingProviderFname(ByVal vData As String)
    mvarOrderingProvider = vData
End Property

Public Property Get OrderingProviderFname() As String
    OrderingProviderFname = mvarOrderingProvider
End Property

Public Property Set Timing(oTiming As TimingType)
    Set mvarTiming = oTiming
End Property

Public Property Get Timing() As TimingType
    Set Timing = mvarTiming
End Property


Public Property Let WarningReason(ByVal vData As String)
    mvarWarningReason = Trim$(vData)
End Property

Public Property Get WarningReason() As String
    WarningReason = mvarWarningReason
End Property
Public Property Let WarningOverriden(ByVal vData As Boolean)
    mvarWarningOverriden = vData
End Property

Public Property Get WarningOverriden() As Boolean
    WarningOverriden = mvarWarningOverriden
End Property

Public Property Let WarningHeader(ByVal vData As String)
    mvarWarningHeader = vData
End Property

Public Property Get WarningHeader() As String
    WarningHeader = mvarWarningHeader
End Property

Public Property Let LastSentNameWas(ByVal vData As String)
    mvarLastSentNameWas = vData
End Property

Public Property Get LastSentNameWas() As String
    LastSentNameWas = mvarLastSentNameWas
End Property

Public Property Let LastSentIenWas(ByVal vData As String)
    mvarLastSentIenWas = vData
End Property

Public Property Get LastSentIenWas() As String
    LastSentIenWas = mvarLastSentIenWas
End Property
Public Property Let LastPreverifyNameWas(ByVal vData As String)
    mvarLastPreverifyNameWas = vData
End Property

Public Property Get LastPreverifyNameWas() As String
    LastPreverifyNameWas = mvarLastPreverifyNameWas
End Property

Public Property Let LastPreverifyNamesMatch(ByVal vData As Boolean)
    mvarLastPreverifyNamesMatch = vData
End Property

Public Property Get LastPreverifyNamesMatch() As Boolean
    LastPreverifyNamesMatch = mvarLastPreverifyNamesMatch
End Property

Public Property Let LastPreverifyIenWas(ByVal vData As String)
    mvarLastPreverifyIenWas = vData
End Property

Public Property Get LastPreverifyIenWas() As String
    LastPreverifyIenWas = mvarLastPreverifyIenWas
End Property
Public Property Let EnteredByIEN(ByVal vData As String)
    mvarEnteredByIen = vData
End Property

Public Property Get EnteredByIEN() As String
    EnteredByIEN = mvarEnteredByIen
End Property
Public Property Let EnteredByName(ByVal vData As String)
    mvarEnteredByName = vData
End Property

Public Property Get EnteredByName() As String
    EnteredByName = mvarEnteredByName
End Property
Public Property Let OrderTypeCode(ByVal vData As hlORDERTYPECODE)
    mvarOrderTypeCode = vData
End Property

Public Property Get OrderTypeCode() As hlORDERTYPECODE
    OrderTypeCode = mvarOrderTypeCode
End Property
Public Property Let ReactivationDate(ByVal vData As String)
    mvarReactivationDate = vData
End Property

Public Property Get ReactivationDate() As String
    ReactivationDate = mvarReactivationDate
End Property
Private Sub Class_Initialize()
     Set mvarTiming = New TimingType
End Sub
