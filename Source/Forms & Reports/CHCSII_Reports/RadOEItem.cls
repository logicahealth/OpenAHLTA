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
Attribute VB_Name = "PatientItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public UnitNumber As String
Public First As String
Public Last As String
Public Middle As String
Public Suffix As String
Public Title As String
Public Degree As String
Public SSN As String
Public Sex As String
Public DOB As Variant
Public Status As String
Public RankGrade As String
Public Organization As String
Public Department As String
Public Category As String
Public UnitId As String
Public EnrollmentStatus As String

'-- Sponsor
'Public SponsorFullName As String
'Public SponsorFirst As String
'Public SponsorLast As String
'Public SponsorMiddle As String
Public SponsorSSN As String
Public FMP As String
Rem Public Relationship As String

'-- Street Address
Public Street1 As String
Public Street2 As String
Public City As String
Public State As String
Public Zip As String
Public Country As String

'-- Phone Home
Public HomePhone As String
Public HomeAreaCode As String

'-- Phone Work
Public WorkPhone As String
Public WorkAreaCode As String

'-- PCM
Public PCM As String


'-- Other Info
Public Race As String
Public Religion As String
Public MaritalStatus As String

'-- Insurance

