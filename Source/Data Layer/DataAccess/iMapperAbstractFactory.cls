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
Attribute VB_Name = "iMapperAbstractFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Abstract Factory Interface for mode specific data mapper factories.

Option Explicit

Public Function DocumentsMapper() As iMapperDocuments
End Function

'Add support components used by mappers.
Public Sub Init(ByRef Support As iParameters, ByRef Cache As Cache)
End Sub

Public Function AllergiesMapper() As iMapperAllergies
End Function

Public Function AppointmentsMapper() As iMapperAppointments
End Function

Public Function ClinicClientMapper() As iMapperClinicClient
End Function

Public Function CoreMapper() As iMapperCore
End Function

Public Function EncounterMapper() As iMapperEncounter
End Function

Public Function InpatientOrderMapper() As iMapperInpatientOrder
End Function

Public Function MedsMapper() As iMapperMeds
End Function

Public Function NonProviderListMapper() As iMapperNonProviderList
End Function

Public Function OrderEntryMapper() As iMapperOrderEntry
End Function


Public Function PatientMapper() As iMapperPatient
End Function

Public Function RadsMapper() As iMapperRads
End Function

Public Function FillMedMapper() As iMapperFillMed
End Function
