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
Attribute VB_Name = "clsAllergyData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mID As String
Private mDescription As String
Private mProspective As Boolean
Private mAllergyType As Integer ' fdbAllergyType
Private mInternalAllergyID As Long

Public Property Let ID(sValue As String)
  mID = sValue
End Property
Public Property Get ID() As String
  ID = mID
End Property

Public Property Let Description(sValue As String)
  mDescription = sValue
End Property
Public Property Get Description() As String
  Description = mDescription
End Property

Public Property Let Prospective(bValue As Boolean)
  mProspective = bValue
End Property
Public Property Get Prospective() As Boolean
  Prospective = mProspective
End Property

Public Property Let AllergyType(sValue As Integer)
  mAllergyType = sValue
End Property
Public Property Get AllergyType() As Integer
  AllergyType = mAllergyType
End Property

Public Property Let InternalAllergyID(sValue As Long)
  mInternalAllergyID = sValue
End Property
Public Property Get InternalAllergyID() As Long
  InternalAllergyID = mInternalAllergyID
End Property

