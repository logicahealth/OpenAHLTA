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
Attribute VB_Name = "Inpatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'hwcc 1/14/2003 - This object is use for testing PrintInpatient
'

Public RegNum As String
Public Specialty As String
Public Capacity As String
Public Bed As String
Public Rank As String
Public Patient As String
Public FMPSSN As String
Public Unit As String
Public PatService As String
Public AdmittingDiagnosis As String
Public AdmissionDate As String
Public AdmissionComments As String
Public Day As String
Public Location As String
Public Status As String
Public DischargeDate As String
Public Disposition As String
Public DischargeDiagnosis As String
Public AdmittingProvider As String
Public AttendingProvider As String
Public TransferDate As String
Public Service As String

