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
Attribute VB_Name = "ReportInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ReportTypeId As ReportType
Public ReportFileName As String
Public LongName As String
Public AlternativeHeader As String
Public Header As Boolean
Public AppendCriteriaToHeader As Boolean
Public DateFrom As Boolean
Public DateTo As Boolean
Public ADate As Boolean
Public CustomDateFormat As String
Public Criteria1 As Collection
Public Criteria2 As Collection
Public Criteria3 As Collection
Public Groupby As Collection
Public ReportCategory As ReportCategory
Public IsGrouping As Boolean
Public IsCriteria1 As Boolean
Public IsCriteria2 As Boolean
Public IsCriteria3 As Boolean
Public IsCriteria1Required As Boolean
Public IsCriteria2Required As Boolean
Public IsCriteria3Required As Boolean
Public IsDisplayed As Boolean
Public Description As String
Public ClassName As String
Public SortReport As Boolean
Public DefaultSortColumn As Boolean
Public SortItems As Collection
Public Switch As String
Public EnableSecondaryFunction As Boolean
Public SecondaryFunctionTtitle As String
Public MaximumDuration As Integer

Public Sub DebugPrintReportInformation()
    Debug.Print "ReportFileName/LongName: " & ReportFileName & "/" & LongName
    Debug.Print "     ReportTypeId: " & ReportTypeId
    Debug.Print "     AlternativeHeader: " & AlternativeHeader
    Debug.Print "     Header: " & Header
    Debug.Print "     AppendCriteriaToHeader: " & AppendCriteriaToHeader
    Debug.Print "     DateFrom: " & DateFrom
    Debug.Print "     DateTo: " & DateTo
    Debug.Print "     ADate: " & ADate
    Debug.Print "     CustomDateFormat: " & CustomDateFormat
    Debug.Print "     Criteria1: " & Criteria1.Count
    Debug.Print "     Criteria2: " & Criteria2.Count
    Debug.Print "     Criteria3: " & Criteria3.Count
    Debug.Print "     Groupby: " & Groupby.Count
    Debug.Print "     ReportCategory: " & ReportCategory
    Debug.Print "     IsGrouping: " & IsGrouping
    Debug.Print "     IsCriteria1: " & IsCriteria1
    Debug.Print "     IsCriteria2: " & IsCriteria2
    Debug.Print "     IsCriteria3: " & IsCriteria3
    Debug.Print "     IsCriteria1Required: " & IsCriteria1Required
    Debug.Print "     IsCriteria2Required: " & IsCriteria2Required
    Debug.Print "     IsCriteria3Required: " & IsCriteria3Required
    Debug.Print "     IsDisplayed: " & IsDisplayed
    Debug.Print "     Description: " & Description
    Debug.Print "     ClassName: " & ClassName
    Debug.Print "     SortReport: " & SortReport
    Debug.Print "     DefaultSortColumn: " & DefaultSortColumn
    Debug.Print "     SortItems: " & SortItems.Count
    Debug.Print "     Switch: " & Switch
    Debug.Print "     EnableSecondaryFunction: " & EnableSecondaryFunction
    Debug.Print "     SecondaryFunctionTtitle: " & SecondaryFunctionTtitle
    Debug.Print "     MaximumDuration: " & MaximumDuration
End Sub
