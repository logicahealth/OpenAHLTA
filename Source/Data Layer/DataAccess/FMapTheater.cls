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
Attribute VB_Name = "FMapTheater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Creates instances of Mappers used to access Oracle database (GEMS Schema).

Option Explicit

Implements iMapperAbstractFactory

Private mSupport As iParameters
Private mCache As Cache

Private Function iMapperAbstractFactory_AllergiesMapper() As DL_Interfaces.iMapperAllergies
Dim Mapper As iMapperAllergies
Dim Cache As CacheMapper

  Set Cache = InitCacheSupport(mCache, "Allergies")
  Set Mapper = New DM_Allergies.MapperTheater
  
  Mapper.Init mSupport, Cache
  
  Set iMapperAbstractFactory_AllergiesMapper = Mapper

End Function

Private Function iMapperAbstractFactory_AppointmentsMapper() As DL_Interfaces.iMapperAppointments

End Function

Private Function iMapperAbstractFactory_ClinicClientMapper() As DL_Interfaces.iMapperClinicClient
Dim Mapper As iMapperClinicClient
Dim Cache As CacheMapper

  Set Cache = InitCacheSupport(mCache, "ClinicsProviders")
  Set Mapper = New DM_ClinicClient.MapperTheater

  Mapper.Init mSupport, Cache

  Set iMapperAbstractFactory_ClinicClientMapper = Mapper

End Function

Private Function iMapperAbstractFactory_CoreMapper() As DL_Interfaces.iMapperCore
Dim Mapper As iMapperCore
Dim Cache As CacheMapper

  Set Cache = InitCacheSupport(mCache, "CoreMapper")
  
  Set Mapper = New DM_Core.MapperTheater
  Mapper.Init mSupport, Cache
  
  Set iMapperAbstractFactory_CoreMapper = Mapper
  
End Function

Private Function iMapperAbstractFactory_DocumentsMapper() As DL_Interfaces.iMapperDocuments
    Dim Mapper As iMapperDocuments
    Dim Cache As CacheMapper
    
    Set Cache = InitCacheSupport(mCache, "DocumentsMapper")
    Set Mapper = New DM_Documents.MapperTheater
    Mapper.Init mSupport, Cache
    
    Set iMapperAbstractFactory_DocumentsMapper = Mapper

End Function

Private Function iMapperAbstractFactory_EncounterMapper() As DL_Interfaces.iMapperEncounter
Dim Mapper As iMapperEncounter
Dim Cache As CacheMapper

  Set Cache = InitCacheSupport(mCache, "EncounterMapper")
  Set Mapper = New DM_Encounters.MapperTheater
  Mapper.Init mSupport, Cache

  Set iMapperAbstractFactory_CoreMapper = Mapper
End Function

Private Sub iMapperAbstractFactory_Init(Support As DL_Support.iParameters, Cache As DL_Support.Cache)
  Set mSupport = Support
  Set mCache = Cache
End Sub

Private Function iMapperAbstractFactory_InpatientOrderMapper() As DL_Interfaces.iMapperInpatientOrder
Dim Mapper As iMapperInpatientOrder
Dim Cache As CacheMapper

  Set Cache = InitCacheSupport(mCache, "InpatientOrderMapper")
  Set Mapper = New DM_InPatOrder.MapperTheater
  Mapper.Init mSupport, Cache

  Set iMapperAbstractFactory_InpatientOrderMapper = Mapper

End Function

Private Function iMapperAbstractFactory_MedsMapper() As DL_Interfaces.iMapperMeds

End Function

Private Function iMapperAbstractFactory_NonProviderListMapper() As DL_Interfaces.iMapperNonProviderList
Dim Mapper As iMapperNonProviderList
Dim Cache As CacheMapper

  Set Cache = InitCacheSupport(mCache, "NonProviderListMapper")
  
  Set Mapper = New DM_NonProviderList.MapperTheater
  Mapper.Init mSupport, Cache
  
  Set iMapperAbstractFactory_NonProviderListMapper = Mapper

End Function

Private Function iMapperAbstractFactory_OrderEntryMapper() As DL_Interfaces.iMapperOrderEntry
    Dim Mapper As iMapperOrderEntry
    Dim Cache As CacheMapper
    
    Set Cache = InitCacheSupport(mCache, "OrderEntryMapper")
    Set Mapper = New DM_OrderEntry.MapperTheater
    Mapper.Init mSupport, Cache
    
    Set iMapperAbstractFactory_OrderEntryMapper = Mapper
End Function

Private Function iMapperAbstractFactory_PatientMapper() As DL_Interfaces.iMapperPatient
Dim Mapper As DL_Interfaces.iMapperPatient
Dim Cache As CacheMapper

  Set Cache = InitCacheSupport(mCache, "PatientMapper")
  Set Mapper = New DM_Patient.MapperTheater
  
  Mapper.Init mSupport, Cache
  
  Set iMapperAbstractFactory_PatientMapper = Mapper

End Function
Private Function iMapperAbstractFactory_RadsMapper() As DL_Interfaces.iMapperRads
Dim Mapper As DL_Interfaces.iMapperRads
Dim Cache As CacheMapper

  Set Cache = InitCacheSupport(mCache, "RadsMapper")
  Set Mapper = New DM_Radiology.MapperTheater
  
  Mapper.Init mSupport, Cache
  
  Set iMapperAbstractFactory_RadsMapper = Mapper

End Function

Private Function iMapperAbstractFactory_FillMedMapper() As DL_Interfaces.iMapperFillMed
Dim Mapper As DL_Interfaces.iMapperFillMed
Dim Cache As CacheMapper

  Set Cache = InitCacheSupport(mCache, "FillMedMapper")
  Set Mapper = New DM_FillMedHistory.MapperTheater
  
  Mapper.Init mSupport, Cache
  
  Set iMapperAbstractFactory_FillMedMapper = Mapper

End Function


