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
Attribute VB_Name = "MapperTheater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements DL_Interfaces.iMapperEncounter

Private mConn As CHCSII_CONN.Conn
Private mCache As CacheMapper

Private Sub iMapperEncounter_Init(Support As DL_Support.iParameters, Cache As DL_Support.CacheMapper)
  Set mConn = Support.Item("DBConn").Value
  Set mCache = Cache
End Sub

Private Function iMapperEncounter_EncounterList(Params As DL_Support.iParameters) As DTOs.Encounters

    Dim EncMapper As iMapperEncounter
    Dim EncMap As EncMapper
    
    Set EncMap = New EncMapper
    
    If mConn.UseMySQL Then
        EncMap.Init mConn, mCache, MYSQL
    Else
        EncMap.Init mConn, mCache, GEMS
    End If
    
    Set EncMapper = EncMap
    
    Set iMapperEncounter_EncounterList = EncMapper.EncounterList(Params)
    
End Function

Private Function iMapperEncounter_LockSection(Params As DL_Support.iParameters) As EncLocks
    Dim EncMapper As iMapperEncounter
    Dim EncMap As EncMapper
    
    Set EncMap = New EncMapper
    
    If mConn.UseMySQL Then
        EncMap.Init mConn, mCache, MYSQL
    Else
        EncMap.Init mConn, mCache, GEMS
    End If
    
    Set EncMapper = EncMap
    
    Set iMapperEncounter_LockSection = EncMapper.LockSection(Params)

    
End Function

Private Function iMapperEncounter_Retrieve(Params As DL_Support.iParameters) As DTOs.Encounter
    Dim EncMapper As iMapperEncounter
    Dim EncMap As EncMapper
    
    Set EncMap = New EncMapper
    
    If mConn.UseMySQL Then
        EncMap.Init mConn, mCache, MYSQL
    Else
        EncMap.Init mConn, mCache, GEMS
    End If
    
    Set EncMapper = EncMap
    
    Set iMapperEncounter_Retrieve = EncMapper.Retrieve(Params)
End Function

Private Function iMapperEncounter_Save(Encounter As DTOs.Encounter) As Boolean
    Dim EncMapper As iMapperEncounter
    Dim EncMap As EncMapper
    
    Set EncMap = New EncMapper
    
    If mConn.UseMySQL Then
        EncMap.Init mConn, mCache, MYSQL
    Else
        EncMap.Init mConn, mCache, GEMS
    End If
    
    Set EncMapper = EncMap
    
    iMapperEncounter_Save = EncMapper.Save(Encounter)
End Function

Private Function iMapperEncounter_UnlockSection(Params As DL_Support.iParameters) As Boolean
    Dim EncMapper As iMapperEncounter
    Dim EncMap As EncMapper
    
    Set EncMap = New EncMapper
    
    If mConn.UseMySQL Then
        EncMap.Init mConn, mCache, MYSQL
    Else
        EncMap.Init mConn, mCache, GEMS
    End If
    
    Set EncMapper = EncMap
    
    iMapperEncounter_UnlockSection = EncMapper.UnlockSection(Params)
End Function

