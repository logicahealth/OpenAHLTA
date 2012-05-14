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
Attribute VB_Name = "ReportItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private eReportType As ReportType
Private eReportCategory As ReportCategory
Private eReportCriteria1 As CriteriaType
Private eReportCriteria2 As CriteriaType
Private eReportCriteria3 As CriteriaType
Private sReportSelection1 As String
Private sReportSelection2 As String
Private sReportSelection3 As String
Private eReportGroupBy1 As CriteriaType
Private eReportGroupBy2 As CriteriaType
Private eReportGroupBy3 As CriteriaType
Private ReportSelCollection1 As Collection
Private ReportSelCollection2 As Collection
Private ReportSelCollection3 As Collection
Private vDateFrom As Variant
Private vDateTo As Variant
Private sSortItemSelected As String
Private sSortDirection As String
Private sCustomDateFormat As String
Private bSecondaryFuncSelected As Boolean
Public Property Get ReportCat() As ReportCategory
    ReportCat = eReportCategory
End Property

Public Property Let ReportCat(ByVal eNewValue As ReportCategory)
    eReportCategory = eNewValue
End Property
Public Property Get Report() As ReportType
    Report = eReportType
End Property

Public Property Let Report(ByVal eNewValue As ReportType)
    eReportType = eNewValue
End Property
Public Property Get Criteria1() As CriteriaType
    Criteria1 = eReportCriteria1
End Property

Public Property Let Criteria1(ByVal eNewValue As CriteriaType)
    eReportCriteria1 = eNewValue
End Property
Public Property Get Criteria2() As CriteriaType
    Criteria2 = eReportCriteria2
End Property

Public Property Let Criteria2(ByVal eNewValue As CriteriaType)
    eReportCriteria2 = eNewValue
End Property

Public Property Get Criteria3() As CriteriaType
    Criteria3 = eReportCriteria3
End Property

Public Property Let Criteria3(ByVal eNewValue As CriteriaType)
    eReportCriteria3 = eNewValue
End Property

Public Property Get GroupBy1() As CriteriaType
    GroupBy1 = eReportGroupBy1
End Property

Public Property Let GroupBy1(ByVal eNewValue As CriteriaType)
    eReportGroupBy1 = eNewValue
End Property

Public Property Get GroupBy2() As CriteriaType
    GroupBy2 = eReportGroupBy2
End Property

Public Property Let GroupBy2(ByVal eNewValue As CriteriaType)
    eReportGroupBy2 = eNewValue
End Property

Public Property Get GroupBy3() As CriteriaType
    GroupBy3 = eReportGroupBy3
End Property

Public Property Let GroupBy3(ByVal eNewValue As CriteriaType)
    eReportGroupBy3 = eNewValue
End Property


Private Sub Class_Initialize()
    '--- SCR-ARM-16148 MSolano 9/10/01; COMMON TO MULTIPLE ARM-SCRs
    '--- Init routine should be called from first top level creatable coclass
    '--- It can be called multiple times, but will actually run only once.
    '--- Instrument all creatable classes in case dll's multiple entry
    '--- points are actually used
     Call ARM_InitAllTransactionsForCHCSIIReports

    eReportType = DiagnosisRep
    eReportCriteria1 = NoCriteria
    eReportCriteria2 = NoCriteria
    eReportCriteria3 = NoCriteria
    sReportSelection1 = ""
    sReportSelection2 = ""
    sReportSelection3 = ""
    eReportGroupBy1 = NoCriteria
    eReportGroupBy2 = NoCriteria
    eReportGroupBy3 = NoCriteria
End Sub

Public Property Get DateFrom() As Variant
    DateFrom = vDateFrom
End Property

Public Property Let DateFrom(ByVal vNewValue As Variant)
    vDateFrom = vNewValue
End Property

Public Property Get DateTo() As Variant
    DateTo = vDateTo
End Property

Public Property Let DateTo(ByVal vNewValue As Variant)
    vDateTo = vNewValue
End Property

Public Property Get CriteriaSel1() As Collection
    Set CriteriaSel1 = ReportSelCollection1
End Property

Public Property Let CriteriaSel1(ByVal oNewValue As Collection)
    Set ReportSelCollection1 = oNewValue
End Property
Public Property Get CriteriaSel2() As Object
    Set CriteriaSel2 = ReportSelCollection2
End Property

Public Property Let CriteriaSel2(ByVal oNewValue As Object)
    Set ReportSelCollection2 = oNewValue
End Property
Public Property Get CriteriaSel3() As Object
    Set CriteriaSel3 = ReportSelCollection3
End Property

Public Property Let CriteriaSel3(ByVal oNewValue As Object)
    Set ReportSelCollection3 = oNewValue
End Property


Public Property Get SortItemSelected() As String
    SortItemSelected = sSortItemSelected
End Property

Public Property Let SortItemSelected(ByVal RHV As String)
    sSortItemSelected = RHV
End Property

Public Property Get SortDirection() As String
    SortDirection = sSortDirection
End Property

Public Property Let SortDirection(ByVal RHV As String)
    sSortDirection = RHV
End Property

Public Property Get CustomDateFormat() As String
    CustomDateFormat = sCustomDateFormat
End Property

Public Property Let CustomDateFormat(ByVal RHV As String)
    sCustomDateFormat = RHV
End Property

Public Property Get SecondaryFuncSelected() As Boolean
    SecondaryFuncSelected = bSecondaryFuncSelected
End Property

Public Property Let SecondaryFuncSelected(ByVal RHV As Boolean)
    bSecondaryFuncSelected = RHV
End Property
