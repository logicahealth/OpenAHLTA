VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RxRadLabBaseType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private mvarOrderId As String 'local copy
Private mvarWarningHeader As String 'local copy
Private mvarErrorIsMessage As Boolean 'local copy
Private mvarWarningReason As String 'local copy
Private mvarTiming As TimingType 'local copy
Private mvarTestName As String 'local copy
Private mvarTestIen As String 'local copy
Private mvarComment As String 'local copy
Private mvarOrderingProvider As String 'local copy
Private mvarOrderingProviderIen As String 'local copy
Private mvarActionReason As String 'local copy
Private mvarOrderingProviderLname As String 'local copy

'*************************************
'The RxRadLabBaseType Class Interface
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

Friend Property Let ErrorIsMessage(ByVal vData As Boolean)
    mvarErrorIsMessage = vData
End Property

Public Property Get ErrorIsMessage() As Boolean
    ErrorIsMessage = mvarErrorIsMessage
End Property

Friend Property Let OrderId(ByVal vData As String)
    mvarOrderId = vData
End Property

Public Property Get OrderId() As String
    OrderId = mvarOrderId
End Property
Friend Property Let OrderIen(ByVal vData As String)
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
Friend Property Let OrderingProviderLname(ByVal vData As String)
    mvarOrderingProviderLname = vData
End Property

Public Property Get OrderingProviderLname() As String
    OrderingProviderLname = mvarOrderingProviderLname
End Property

Friend Property Let OrderingProviderIen(ByVal vData As String)
    mvarOrderingProviderIen = vData
End Property

Public Property Get OrderingProviderIen() As String
    OrderingProviderIen = mvarOrderingProviderIen
End Property

Friend Property Let OrderingProviderFname(ByVal vData As String)
    mvarOrderingProvider = vData
End Property

Public Property Get OrderingProviderFname() As String
    OrderingProviderFname = mvarOrderingProvider
End Property

Friend Property Set Timing(oTiming As TimingType)
    Set mvarTiming = oTiming
End Property

Public Property Get Timing() As TimingType
    Set Timing = mvarTiming
End Property


Public Property Let WarningReason(ByVal vData As String)
    mvarWarningReason = vData
End Property

Public Property Get WarningReason() As String
    WarningReason = mvarWarningReason
End Property

Friend Property Let WarningHeader(ByVal vData As String)
    mvarWarningHeader = vData
End Property

Public Property Get WarningHeader() As String
    WarningHeader = mvarWarningHeader
End Property


