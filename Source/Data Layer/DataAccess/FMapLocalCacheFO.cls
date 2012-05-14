VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FMapLocalCacheFO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Creates instances of Mappers used in Local Cache mode.  In Local Cahce mode, data is received and saved
'from a Web Service.

Option Explicit

Implements iMapperAbstractFactory

Private mSupport As iParameters
Private mCache As Cache

Private Function iMapperAbstractFactory_AllergiesMapper() As DL_Interfaces.iMapperAllergies
Dim Mapper As iMapperAllergies
Dim Cache As CacheMapper

  Set Cache = InitCacheSupport(mCache, "Allergies")
  Set Mapper = New DM_Allergies.MapperPGUI
  
  Mapper.Init mSupport, Cache
  
  Set iMapperAbstractFactory_AllergiesMapper = Mapper

End Function

Private Function iMapperAbstractFactory_AppointmentsMapper() As DL_Interfaces.iMapperAppointments

End Function

Private Function iMapperAbstractFactory_ClinicClientMapper() As DL_Interfaces.iMapperClinicClient
Dim Mapper As iMapperClinicClient
Dim Cache As CacheMapper

  Set Cache = InitCacheSupport(mCache, "ClinicsProviders")
  Set Mapper = New DM_ClinicClient.MapperPGUI

  Mapper.Init mSupport, Cache

  Set iMapperAbstractFactory_ClinicClientMapper = Mapper

End Function

Private Function iMapperAbstractFactory_CoreMapper() As DL_Interfaces.iMapperCore
Dim Mapper As iMapperCore
Dim Cache As CacheMapper

  Set Cache = InitCacheSupport(mCache, "CoreMapper")
  
  Set Mapper = New DM_Core.MapperPGUI
  Mapper.Init mSupport, Cache
  
  Set iMapperAbstractFactory_CoreMapper = Mapper
  
End Function

Private Function iMapperAbstractFactory_DocumentsMapper() As DL_Interfaces.iMapperDocuments

End Function

Private Function iMapperAbstractFactory_EncounterMapper() As DL_Interfaces.iMapperEncounter

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
  
  Set Mapper = New DM_NonProviderList.MapperPGUI
  Mapper.Init mSupport, Cache
  
  Set iMapperAbstractFactory_NonProviderListMapper = Mapper
  
End Function

Private Function iMapperAbstractFactory_OrderEntryMapper() As DL_Interfaces.iMapperOrderEntry

End Function

Private Function iMapperAbstractFactory_PatientMapper() As DL_Interfaces.iMapperPatient
Dim Mapper As DL_Interfaces.iMapperPatient
Dim Cache As CacheMapper

  Set Cache = InitCacheSupport(mCache, "PatientMapper")
  
  Set Mapper = New DM_Patient.MapperPGUI
  Mapper.Init mSupport, Cache
  
  Set iMapperAbstractFactory_PatientMapper = Mapper
  
End Function

Private Function iMapperAbstractFactory_RadsMapper() As DL_Interfaces.iMapperRads
Dim Mapper As DL_Interfaces.iMapperRads
Dim Cache As CacheMapper

  Set Cache = InitCacheSupport(mCache, "RadsMapper")
  Set Mapper = New DM_Radiology.MapperPGUI
  
  Mapper.Init mSupport, Cache
  
  Set iMapperAbstractFactory_RadsMapper = Mapper

End Function

Private Function iMapperAbstractFactory_FillMedMapper() As DL_Interfaces.iMapperFillMed

End Function
