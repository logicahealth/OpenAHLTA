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
Attribute VB_Name = "GEMS_Address"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarType      As String
Private mvarStreet1   As String
Private mvarStreet2   As String
Private mvarCity      As String
Private mvarState     As String
Private mvarZip       As String
Private mvarCounty    As String
Private mvarCountry   As String
Private mvarUpdateFlag As String

Public Property Get UpdateFlag() As String:           UpdateFlag = mvarUpdateFlag:  End Property
Public Property Let UpdateFlag(ByVal RHV As String):  mvarUpdateFlag = RHV:         End Property

Public Property Get AddrType() As String:           AddrType = mvarType:    End Property
Attribute AddrType.VB_MemberFlags = "200"
Public Property Let AddrType(ByVal RHV As String):  mvarType = RHV:         End Property

Public Property Get Street1() As String:            Street1 = mvarStreet1:  End Property
Public Property Let Street1(ByVal RHV As String):   mvarStreet1 = RHV:      End Property

Public Property Get Street2() As String:            Street2 = mvarStreet2:  End Property
Public Property Let Street2(ByVal RHV As String):   mvarStreet2 = RHV:      End Property

Public Property Get City() As String:               City = mvarCity:        End Property
Public Property Let City(ByVal RHV As String):      mvarCity = RHV:         End Property

Public Property Get State() As String:              State = mvarState:      End Property
Public Property Let State(ByVal RHV As String):     mvarState = RHV:        End Property

Public Property Get Zip() As String:                Zip = mvarZip:          End Property
Public Property Let Zip(ByVal RHV As String):       mvarZip = RHV:          End Property

Public Property Get County() As String:             County = mvarCounty:    End Property
Public Property Let County(ByVal RHV As String):    mvarCounty = RHV:       End Property

Public Property Get Country() As String:            Country = mvarCountry:  End Property
Public Property Let Country(ByVal RHV As String):   mvarCountry = RHV:      End Property

