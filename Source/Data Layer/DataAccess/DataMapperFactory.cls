VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataMapperFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Data Layer service interface.  It is used to retrieve mapper objects for a given mode of operation (ie. backend).

Option Explicit

Implements iDataMapperFactory

Private mDefaultMode As enumDataAcessMode

Private mCacheConfigFile As String                  'File that holds cache settings
Private mDataCachPath As String                     'Location of local cache files

Private mMapperFactory As iMapperAbstractFactory    'Mapper factory for a given mode

Private mSupport As iParameters                     'May include references to: Cache, CHCSII_CONN.Conn, MObjects.MDB, LGSGateway
Private mCache As Cache

Private Sub Class_Initialize()
  Set mSupport = New ParametersGeneric
  Set mCache = New Cache
  
  SetDefaultMode
End Sub

Private Function iDataMapperFactory_Allergies() As DL_Interfaces.iMapperAllergies
    Set iDataMapperFactory_Allergies = mMapperFactory.AllergiesMapper
End Function

Private Function iDataMapperFactory_AppointmentsMapper() As DL_Interfaces.iMapperAppointments
    Set iDataMapperFactory.AppointmentsMapper = mMapperFactory.AppointmentsMapper
End Function

'Reset mode of operation back to default
Private Sub iDataMapperFactory_ClearOverride()
  SetCurrentMode mDefaultMode
End Sub

Private Property Let iDataMapperFactory_CacheConfigFile(ByVal RHS As String)
  mCacheConfigFile = RHS
End Property

Private Property Get iDataMapperFactory_CacheConfigFile() As String
  iDataMapperFactory_CacheConfigFile = mCacheConfigFile
End Property

Private Function iDataMapperFactory_ClinicClientMapper() As DL_Interfaces.iMapperClinicClient
  Set iDataMapperFactory_ClinicClientMapper = mMapperFactory.ClinicClientMapper
End Function

Private Function iDataMapperFactory_CoreMapper() As DL_Interfaces.iMapperCore
  Set iDataMapperFactory_CoreMapper = mMapperFactory.CoreMapper
End Function

Private Property Let iDataMapperFactory_DataCachePath(ByVal RHS As String)
  mDataCachPath = RHS
End Property

Private Property Get iDataMapperFactory_DataCachePath() As String
  iDataMapperFactory_DataCachePath = mDataCachPath
End Property

Private Property Let iDataMapperFactory_DefaultDataAccessMode(ByVal RHS As DL_Interfaces.enumDataAcessMode)
  If mDefaultMode = 0 Then
    mDefaultMode = RHS
    
    SetCurrentMode mDefaultMode
  End If
End Property

Private Property Get iDataMapperFactory_DefaultDataAccessMode() As DL_Interfaces.enumDataAcessMode
  iDataMapperFactory_DefaultDataAccessMode = mDefaultMode
End Property

Private Function iDataMapperFactory_DocumentsMapper() As DL_Interfaces.iMapperDocuments
    Set iDataMapperFactory_DocumentsMapper = mMapperFactory.DocumentsMapper
End Function

Private Function iDataMapperFactory_EncounterMapper() As DL_Interfaces.iMapperEncounter
  Set iDataMapperFactory_EncounterMapper = mMapperFactory.EncounterMapper
End Function
Private Function iDataMapperFactory_InpatientOrderMapper() As DL_Interfaces.iMapperInpatientOrder
    Set iDataMapperFactory_InpatientOrderMapper = mMapperFactory.InpatientOrderMapper
End Function

Private Function iDataMapperFactory_MedsMapper() As DL_Interfaces.iMapperMeds

End Function

Private Function iDataMapperFactory_NonProviderList() As DL_Interfaces.iMapperNonProviderList
    Set iDataMapperFactory_NonProviderList = mMapperFactory.NonProviderListMapper
End Function

Private Function iDataMapperFactory_OrderEntryMapper() As DL_Interfaces.iMapperOrderEntry
    Set iDataMapperFactory_OrderEntryMapper = mMapperFactory.OrderEntryMapper
End Function

Private Function iDataMapperFactory_PatientMapper() As DL_Interfaces.iMapperPatient
    Set iDataMapperFactory_PatientMapper = mMapperFactory.PatientMapper
End Function
Private Function iDataMapperFactory_RadsMapper() As DL_Interfaces.iMapperRads
    Set iDataMapperFactory_RadsMapper = mMapperFactory.RadsMapper
End Function
Private Function iDataMapperFactory_FillMedMapper() As DL_Interfaces.iMapperFillMed
    Set iDataMapperFactory_FillMedMapper = mMapperFactory.FillMedMapper
End Function
'Override default mode
Private Sub iDataMapperFactory_OverrideDataAccessMode(ByVal NewDataAccessMode As DL_Interfaces.enumDataAcessMode)
  SetCurrentMode NewDataAccessMode
End Sub

'Add this object to Support collection object.
Private Sub AddDataMapperFactory()
  If Not mSupport.Exists("DataMapperFactory") Then
    mSupport.Add "DataMapperFactory", Me
  End If
End Sub

'Add Database Connection object to Support collection object.
Private Sub AddDBConnSupport(ByVal DataAccessMode As enumDataAcessMode)
Dim DBConn As CHCSII_CONN.Conn

  If Not mSupport.Exists("DBConn") Then
    Set DBConn = New CHCSII_CONN.Conn
    
    Select Case DataAccessMode
      Case damCDRTraining, damProviderGUITraining
        DBConn.InitDefaultMode ITT
        
      Case damTheater, damCHCS
        DBConn.InitDefaultMode GEMS
      
      Case damMySQL
        DBConn.InitDefaultMode MYSQL
        
      Case damProviderGUI, damLocalCacheFailover
        DBConn.InitDefaultMode LGS
      
      Case Else
        DBConn.InitDefaultMode CDR
    End Select
    
    mSupport.Add "DBConn", DBConn
  End If

End Sub

'Add LGS Gateway object to Support collection object.
Private Sub AddLGSSupport()
Dim LGSGateway As GatewayLGS

  If Not mSupport.Exists("LGSGateway") Then
    Set LGSGateway = New GatewayLGS

    mSupport.Add "LGSGateway", LGSGateway
  End If

End Sub

'Init Support Collection and Mode Mapper Factory.
Private Sub SetCurrentMode(ByVal DataAccessMode As enumDataAcessMode)
  
  AddDataMapperFactory
  AddDBConnSupport DataAccessMode

  Select Case DataAccessMode
    Case damCDRTraining
      Set mMapperFactory = New FMapCDRTraining
          
    Case damTheater, damMySQL
      Set mMapperFactory = New FMapTheater
    
    Case damProviderGUI
      AddLGSSupport
'      Set mMapperFactory = New FMapPGUI
    
    Case damProviderGUITraining
      Set mMapperFactory = New FMapPGUITraining
    
    Case damCHCS
      
    Case damUnitTesting
      Set mMapperFactory = New FMapUnitTesting
    
    Case damLocalCache
      AddLGSSupport
'      Set mMapperFactory = New FMapLocalCache
    
    Case damLocalCacheServer
'      Set mMapperFactory = New FMapLocalCacheSvr
      
    Case damLocalCacheFailover
      AddLGSSupport
'      Set mMapperFactory = New FMapLocalCacheFO
    
    Case Else   '    damCDR
'      Set mMapperFactory = New FMapCDR
        MsgBox "SetCurrentMode CDR?"
  End Select
  
  mMapperFactory.Init mSupport, mCache
  
End Sub

Private Sub SetDefaultMode()
Dim Reader As ReaderXML
Dim Persist As iObjectPersist
Dim ConfigFile As String
Dim Config As iParameters

  On Error GoTo ErrHandler
    
  ConfigFile = App.Path & "\DataAccess.xml"
  
  If Len(Dir$(ConfigFile)) = 0 Then Exit Sub
  
  Set Reader = New ReaderXML
  
  Reader.Load ConfigFile
  
  Set Config = New ParametersGeneric
  Set Persist = Config
  
  Persist.Restore Reader
  
  If Config.Exists("DefaultMode") Then
    mDefaultMode = Config.Item("DefaultMode").Value
  End If
  
  Exit Sub
  
ErrHandler:
  'Unable to load config info so don't use

End Sub

