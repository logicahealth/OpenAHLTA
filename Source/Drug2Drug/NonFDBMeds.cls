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
Attribute VB_Name = "NonFDBMed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mDescription As String
Private mLastFilled As String
Private mRefills As String
Private mStatus As String
Private mOrderNumber As String

Public Property Get Description() As String
  Description = mDescription
End Property

Public Property Let Description(sValue As String)
  mDescription = sValue
End Property

Public Property Get LastFilled() As String
  LastFilled = mLastFilled
End Property
Public Property Let LastFilled(sValue As String)
  mLastFilled = sValue
End Property

Public Property Get Refills() As String
  Refills = mRefills
End Property
Public Property Let Refills(sValue As String)
  mRefills = sValue
End Property

Public Property Get Status() As String
  Status = mStatus
End Property
Public Property Let Status(sValue As String)
  mStatus = sValue
End Property

Public Property Get OrderNumber() As String
  OrderNumber = mOrderNumber
End Property
Public Property Let OrderNumber(sValue As String)
  mOrderNumber = sValue
End Property


