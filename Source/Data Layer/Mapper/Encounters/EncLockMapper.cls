VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EncLockMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements iMapperEncounter


Private mConn As CHCSII_CONN.Conn
Private mCache As CacheMapper

Private mMode As CHCSII_CONN.eBackend       'Data Access mode used for DAS object
Private mUseCache As Boolean

Private Sub iEncLock_Init(Conn As CHCSII_CONN.Conn, Cache As DL_Support.CacheMapper, ByVal DataAccessMode As CHCSII_CONN.eBackend, Optional ByVal CacheConfigItems As Boolean = False)
  Set mConn = Conn
  Set mCache = Cache
  
  mMode = DataAccessMode
  mUseCache = CacheConfigItems
End Sub

Private Function iMapperEncounter_EncounterList(Params As DL_Support.iParameters) As DTOs.EncounterDTOs

End Function

Private Function iMapperEncounter_LockSection(Params As DL_Support.iParameters) As DTOs.EncLocks

    Dim oEncLock As iEncLock
    Dim oSQL As ISqlOps
    
    If mMode = CDR Then
        'Use stored procedures
        Set oEncLock = New EncLockStoredProc
    Else
        'Use SQL
        Set oEncLock = New EncLockSQL
    End If
    
    oEncLock.Init mConn, mCache, mMode, mUseCache
    
    Set oSQL = oEncLock.LockSection(Params)
    
    
    
    
End Function

Private Function iMapperEncounter_Retrieve(Params As DL_Support.iParameters) As DTOs.EncounterDTO

End Function

Private Function iMapperEncounter_Save(Encounter As DTOs.EncounterDTO) As Boolean

End Function

Private Function iMapperEncounter_UnlockSection(Params As DL_Support.iParameters) As Boolean

End Function
