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
Attribute VB_Name = "iDataMapperFacotry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Astract interface that is implemented by the DataMapperFactory class in the DataAccess component.  Using this interface allows
'us to pass the DataMapperFactory instance to Mapper class instance (prevents DLL circular references which VB does not allow).

Option Explicit

'All Supported Data Access Modes
Public Enum enumDataAcessMode
  damCDR = 1
  damCDRTraining = 2
  damTheater = 10
  damProviderGUI = 20
  damProviderGUITraining = 21
  damCHCS = 30
  damUnitTesting = 40
End Enum

Public CacheConfigFile As String                            'Path and File name of the Configuration File for the Cache object (in DL_Support component)
Public DataCachePath As String                              'Path where the Cache object should save cache entries in the file system  (in DL_Support component)
Public DefaultDataAccessMode As enumDataAcessMode           'Default Data Access Mode

'Resets data access mappers to Default Mode
Public Sub ClearOverride()
End Sub

'Overriders the default data access mode to a new mode which changes the mappers used.
Public Sub OverrideDataAccessMode(ByVal NewDataAccessMode As enumDataAcessMode)
End Sub

