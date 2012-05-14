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
Attribute VB_Name = "PassWordRules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public LifeSpan                As Long
Public ExpDate                 As Long
Public MinLength               As Long
Public MaxLength               As Long
Public MinNumeric              As Long
Public MixedCase               As Long
Public PercentChanged          As Long
Public PercentDuplicate        As Long
Public DictionaryCheck         As Long
Public NumberNonAlphanumeric   As Long
Public AllowChanges            As Long
Public KeepHistory             As Long
Public Remind                  As Long



Public Function setProperty(PropertyName, PropertyValue) As Boolean

    Select Case PropertyName
    Case "passwdlifespan":                  LifeSpan = PropertyValue
    Case "passwdexpdate":                   ExpDate = PropertyValue
    Case "passwdminlength":                 MinLength = PropertyValue
    Case "passwdmaxlength":                 MaxLength = PropertyValue
    Case "passwdminnumeric":                MinNumeric = PropertyValue
    Case "passwdmixedcase":                 MixedCase = PropertyValue
    Case "passwdpercentchanged":            PercentChanged = PropertyValue
    Case "passwdpercentduplicate":          PercentDuplicate = PropertyValue
    Case "passwddictionarycheck":           DictionaryCheck = PropertyValue
    Case "passwdnumbernonalphanumeric":     NumberNonAlphanumeric = PropertyValue
    Case "passwdallowchanges":              AllowChanges = PropertyValue
    Case "passwdkeephistory":               KeepHistory = PropertyValue
    Case "passwdremind":                    Remind = PropertyValue
    End Select

    
End Function

Public Function SetDefaultProperties() As Boolean
'-- if we don't have a connection, make these the defaults
 LifeSpan = 90
 ExpDate = 0
 MinLength = 2
 MaxLength = 16
 MinNumeric = 1
 MixedCase = 1
 PercentChanged = 0
 PercentDuplicate = 100
 DictionaryCheck = 0
 AllowChanges = 0
 NumberNonAlphanumeric = 1
 KeepHistory = 999999
 Remind = 7
End Function
