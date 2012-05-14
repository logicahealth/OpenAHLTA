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
Attribute VB_Name = "WellnessReminder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ReminderText As String '-- sTypeName
Public Status As String '-- sStatusName
Public ReminderDocNCID As String '-- sASN1TypeNCID
Public LastServiceDate As Double  '-- dLastDoneDate
Public RecommendedDate As Double '-- dRecDate
Public EarliestDate As Double '-- dEarlyDate
Public Comment As String '-- Reminder Comment

Public TypeNCID As String '-- sTypeNCID
Public EventId As String '-- sEventId
Public SequenceNumberNCID As String '-- SeqNumberNCID
Public MarkedAsDeleted As Boolean '-- sMarkedAsDeleted
Public StatusNCID As String '-- sStatusNCID

Public Category As TypeOfReminder
Public PatientId As String
