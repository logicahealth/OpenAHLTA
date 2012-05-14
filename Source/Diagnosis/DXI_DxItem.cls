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
Attribute VB_Name = "clsDXI_DxItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'''Properties required for adding Dx Item to A&P
Public psDescription   As String
Public psNCID          As String
Public psSnoID      As String
Public psICD9          As String
Public psPrefix        As String
Public psExtICD_Name   As String
Public psExtICD_Code   As String
Public psExtICD_NCID   As String
Public psNotFromMedcin As Boolean   'JT SCR 42399 - track if we have a valid medcin id

