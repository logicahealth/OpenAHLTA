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
Attribute VB_Name = "Report_Sort_Item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim lReportId As Long
Dim sSortField As String
Dim sDisplayValue As String

Public Property Get ReportId() As Long
    ReportId = lReportId
End Property

Public Property Let ReportId(ByVal RHV As Long)
    lReportId = RHV
End Property

Public Property Get SortField() As String
    SortField = sSortField
End Property

Public Property Let SortField(ByVal RHV As String)
    sSortField = RHV
End Property

Public Property Get DisplayValue() As String
    DisplayValue = sDisplayValue
End Property

Public Property Let DisplayValue(ByVal RHV As String)
    sDisplayValue = RHV
End Property
