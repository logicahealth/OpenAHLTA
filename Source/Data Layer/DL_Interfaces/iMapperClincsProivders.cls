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
Attribute VB_Name = "iMapperClincsProivders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Searches for Clinics based upon the Params, which should be
'an instance of ClinicSearchParam
'Public Function GetClinics(Params As iParameters) As DTOs.Clinics
'
'End Function

'Searches for Providers based upon the Params, which should be
'an instance of ProviderSearchParam
Public Function GetProviders(Params As iParameters) As DTOs.Providers

End Function
