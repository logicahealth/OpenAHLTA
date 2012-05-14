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
Attribute VB_Name = "FMapUnitTesting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Creates instances of Mappers used for Unit Testing.

Option Explicit

Implements iMapperAbstractFactory

Private mSupport As iParameters
Private mCache As Cache

Private Function iMapperAbstractFactory_AllergiesMapper() As DL_Interfaces.iMapperAllergies
Dim Mapper As iMapperAllergies
Dim Cache As CacheMapper

  Set Cache = InitCacheSupport(mCache, "Allergies")
  Set Mapper = New DM_Allergies.MapperUnitTest
  
  Mapper.Init mSupport, Cache
  
  Set iMapperAbstractFactory_AllergiesMapper = Mapper

End Function

Private Function iMapperAbstractFactory_AppointmentsMapper() As DL_Interfaces.iMapperAppointments

End Function

Private Function iMapperAbstractFactory_ClinicClientMapper() As DL_Interfaces.iMapperClinicClient

End Function

Private Function iMapperAbstractFactory_CoreMapper() As DL_Interfaces.iMapperCore
Dim Mapper As iMapperCore
Dim Cache As CacheMapper

  Set Cache = InitCacheSupport(mCache, "CoreMapper")
  
  Set Mapper = New CoreUnitTest
  Mapper.Init mSupport, Cache
  
  Set iMapperAbstractFactory_CoreMapper = Mapper
  
End Function

Private Function iMapperAbstractFactory_DocumentsMapper() As DL_Interfaces.iMapperDocuments

End Function

Private Function iMapperAbstractFactory_EncounterMapper() As DL_Interfaces.iMapperEncounter
'Dim Mapper As iMapperEncounter
'Dim Cache As CacheMapper
'
'  Set Cache = InitCacheSupport(mCache, "EncounterMapper")
'
'  Set Mapper = New DM_Encounters.MapperUnitTesting
'  Mapper.Init mSupport, Cache
'
'  Set iMapperAbstractFactory_NonProviderListMapper = Mapper
End Function

Private Sub iMapperAbstractFactory_Init(Support As DL_Support.iParameters, Cache As DL_Support.Cache)
  Set mSupport = Support
  Set mCache = Cache
End Sub

Private Function iMapperAbstractFactory_InpatientOrderMapper() As DL_Interfaces.iMapperInpatientOrder

End Function

Private Function iMapperAbstractFactory_MedsMapper() As DL_Interfaces.iMapperMeds

End Function

Private Function iMapperAbstractFactory_NonProviderListMapper() As DL_Interfaces.iMapperNonProviderList
Dim Mapper As iMapperNonProviderList
Dim Cache As CacheMapper

  Set Cache = InitCacheSupport(mCache, "NonProviderListMapper")
  
  Set Mapper = New DM_NonProviderList.MapperUnitTesting
  Mapper.Init mSupport, Cache
  
  Set iMapperAbstractFactory_NonProviderListMapper = Mapper

End Function


Private Function iMapperAbstractFactory_OrderEntryMapper() As DL_Interfaces.iMapperOrderEntry

End Function

Private Function iMapperAbstractFactory_PatientMapper() As DL_Interfaces.iMapperPatient
Dim Mapper As DL_Interfaces.iMapperPatient
Dim Cache As CacheMapper

  Set Cache = InitCacheSupport(mCache, "PatientMapper")
  
  Set Mapper = New DM_Patient.MapperUnitTesting
  Mapper.Init mSupport, Cache
  
  Set iMapperAbstractFactory_PatientMapper = Mapper

End Function
Private Function iMapperAbstractFactory_RadsMapper() As DL_Interfaces.iMapperRads

End Function

Private Function iMapperAbstractFactory_FillMedMapper() As DL_Interfaces.iMapperFillMed

End Function
