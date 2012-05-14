VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RFV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private msCHCSReasonForVisit As String


Public Property Get CHCSReasonForVisit() As String
  CHCSReasonForVisit = msCHCSReasonForVisit
End Property
Friend Property Let CHCSReasonForVisit(vdata As String)
  msCHCSReasonForVisit = vdata
End Property


Private Sub Class_Initialize()
    '--- SCR-ARM-16107 MSolano 8/27/01
    '--- Init routine should be called from first top level creatable coclass
    '--- It can be called multiple times, but will actually run only once.
    '--- Instrument all creatable classes in case dll's multiple entry
    '--- points are actually used
    Call ARM_InitAllTransactionsForReasonForVisit

    msCHCSReasonForVisit = "" 'change to txtReasonForVisit.Text
End Sub
