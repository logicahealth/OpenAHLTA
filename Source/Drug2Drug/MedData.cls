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
Attribute VB_Name = "MedData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private mID As String
Private mDescription As String
Private mProspective As Boolean
Private mRoute As String
Private mRouteAbbrev As String
Private mStrength As String
Private mStrengthUnit As String
Private mIsActive As Boolean

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

Public Property Let IsActive(bValue As Boolean)
  mIsActive = bValue
End Property
Public Property Get IsActive() As Boolean
  IsActive = mIsActive
End Property
Public Property Let Prospective(bValue As Boolean)
  mProspective = bValue
End Property
Public Property Get Prospective() As Boolean
  Prospective = mProspective
End Property

Public Property Let Route(sValue As String)
  mRoute = sValue
End Property
Public Property Get Route() As String
  Route = mRoute
End Property

Public Property Let RouteAbbrev(sValue As String)
  mRouteAbbrev = sValue
End Property
Public Property Get RouteAbbrev() As String
  RouteAbbrev = mRouteAbbrev
End Property

Public Property Let Strength(sValue As String)
  mStrength = sValue
End Property
Public Property Get Strength() As String
  Strength = mStrength
End Property

Public Property Let StrengthUnit(sValue As String)
  mStrengthUnit = sValue
End Property
Public Property Get StrengthUnit() As String
  StrengthUnit = mStrengthUnit
End Property


