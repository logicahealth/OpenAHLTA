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
Attribute VB_Name = "CatchmentItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mbFilterByCatchment As Boolean
Private mbIsNew As Boolean
Private mlReportID As Long
Private msItemStatus As String
Private msReportTitle As String
Public Property Get ReportID() As Long
    ReportID = mlReportID
End Property

Public Property Let ReportID(ByVal RHV As Long)
    mlReportID = RHV
End Property

Public Property Get ReportTitle() As String
    ReportTitle = msReportTitle
End Property

Public Property Let ReportTitle(ByVal RHV As String)
    msReportTitle = RHV
End Property

Public Property Get FilterByCatchment() As Boolean
    FilterByCatchment = mbFilterByCatchment
End Property

Public Property Let FilterByCatchment(ByVal RHV As Boolean)
    mbFilterByCatchment = RHV
End Property

Public Property Get ItemStatus() As String
    ItemStatus = msItemStatus
End Property

Public Property Let ItemStatus(ByVal RHV As String)
    msItemStatus = RHV
End Property

Public Property Get IsNew() As Boolean
    IsNew = mbIsNew
End Property

Public Property Let IsNew(ByVal RHV As Boolean)
    mbIsNew = RHV
End Property
