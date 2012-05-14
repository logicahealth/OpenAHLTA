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
Attribute VB_Name = "EncMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements iMapperEncounter

Public Enum EncClass
  ec_Followup = 1
  ec_New = 2
  ec_NoCount = 3
End Enum

Private mConn As CHCSII_CONN.Conn
Private mCache As CacheMapper

Private mMode As CHCSII_CONN.eBackend       'Data Access mode used for DAS object
Private mUseCache As Boolean

Public Sub Init(Conn As CHCSII_CONN.Conn, Cache As DL_Support.CacheMapper, ByVal DataAccessMode As CHCSII_CONN.eBackend, Optional ByVal CacheConfigItems As Boolean = False)
  Set mConn = Conn
  Set mCache = Cache
  
  mMode = DataAccessMode
  mUseCache = CacheConfigItems
End Sub

Private Function iMapperEncounter_EncounterList(Params As DL_Support.iParameters) As DTOs.Encounters
    Dim oEncDataMapper As iEncDataMapper
    
    Set oEncDataMapper = New EncDataSQL
    
    oEncDataMapper.Init mConn, mCache, mMode
    
    Set iMapperEncounter_EncounterList = oEncDataMapper.EncounterList(Params)
End Function

Private Sub iMapperEncounter_Init(Support As DL_Support.iParameters, Cache As DL_Support.CacheMapper)

End Sub

Private Function iMapperEncounter_LockSection(Params As DL_Support.iParameters) As DTOs.EncLocks

    Dim oEncLock As iEncLock
    Dim oSQL As ISqlOps
    Dim oEncLockInfo As EncLock
    Dim oEncLockInfos As EncLocks
    Dim nMyRetries As Long
    
    Dim EncLockParams As ParamsEncLock
    
    Set EncLockParams = Params
    
    If mMode = CDR And mMode <> ITT Then
        'Use stored procedures
        Set oEncLock = New EncLockStoredProc
    Else
        'Use SQL
        Set oEncLock = New EncLockSQL
    End If
    
    oEncLock.Init mConn, mCache, mMode, mUseCache
    
TryAgain:

    Set oSQL = oEncLock.LockSection(Params)
    
    oSQL.MoveFirst
    If oSQL.EOF Then
        If nMyRetries > 10 Then
            'Raise Unable to Lock Error
            'ERR_UNABLE_TO_LOCK_ENCOUNTER
            Err.Raise 9809809, "LockSection", "Unable to Lock Encounter"
            Exit Function 'SCR-26095
        Else
            nMyRetries = nMyRetries + 1
            GoTo TryAgain
        End If
    End If
    
    'Scott Foster - 4/14/03 SCR 32981 Moved this code outside of the loop
    oSQL.MoveFirst
    '*********
    Do While Not oSQL.EOF
        ' We consider a lock ours if it is for our particular section or for the whole encounter
        If oSQL("USER_NCID") = EncLockParams.UserNCID _
            And ((oSQL("SECTION_NUMBER") = EncLockParams.SectionNumber _
                    And oSQL("EVENT_ID") = EncLockParams.EventID) _
                        Or (oSQL("SECTION_NUMBER") = 99 And oSQL("EVENT_ID") = 0)) _
            And oSQL("WORKSTATION") = EncLockParams.WorkStation Then
                'Lock successful
                Exit Function
        Else
            'Multiple SO note can be open for an encounter
            If EncLockParams.SectionNumber <> eSOOpen Then
                'Mulitple new text notes can be open for an encounter 'SCR-35244
                If Not (EncLockParams.SectionNumber = eNotes And EncLockParams.EventID = 0) Then

                    Set oEncLockInfo = New EncLock
                    
                    oEncLockInfo.EventID = oSQL("EVENT_ID")
                    oEncLockInfo.LockDate = oSQL("DATE_TIME")
                    oEncLockInfo.LockStatus = oSQL("STATUS")
                    oEncLockInfo.OwnerNCID = oSQL("USER_NCID")
                    oEncLockInfo.Section = oSQL("SECTION_NUMBER")
                    oEncLockInfo.WorkStation = oSQL("WORKSTATION")
                    oEncLockInfo.Unbreakable = (oSQL("STATUS") = STATUS_UPDATING And oSQL("LOCK_MINUTES") < MAX_UPDATE_MINUTES)
                    
                    If oEncLockInfos Is Nothing Then
                        Set oEncLockInfos = New EncLocks
                    End If
                    
                    oEncLockInfos.Add oEncLockInfo
                End If
            End If
        End If
        oSQL.MoveNext
    Loop

    
    Set iMapperEncounter_LockSection = oEncLockInfos
    
End Function

Private Function iMapperEncounter_Retrieve(Params As DL_Support.iParameters) As DTOs.Encounter

    Dim oEncDataMapper As iEncDataMapper
    
    Set oEncDataMapper = New EncDataSQL
    
    oEncDataMapper.Init mConn, mCache, mMode
    
    Set iMapperEncounter_Retrieve = oEncDataMapper.Retrieve(Params)

End Function

Private Function iMapperEncounter_Save(Encounter As DTOs.Encounter) As Boolean
    Dim oEncDataMapper As iEncDataMapper
    Dim oSectionDataMapper As iSectionDataMapper
    Dim State As iState
    Dim Section As iSection
    Dim i As Long
    
    Set State = Encounter
    If State.State = osModified Or State.State = osNew Then
        Set oEncDataMapper = New EncDataSQL
        
        oEncDataMapper.Init mConn, mCache, mMode
        
        iMapperEncounter_Save = oEncDataMapper.Save(Encounter)
    End If
    
    For i = 1 To Encounter.Sections.Count
        Set State = Encounter.Sections(i)
    
        If State.State = osNew Or State.State = osModified Then
        
            Set oSectionDataMapper = New SectionDataSQL
            
            'Save Section Data
            Set Section = Encounter.Sections(i)
            
            Select Case Section.Id
            
            Case "AnP"
                oSectionDataMapper.SaveAP Encounter
            Case "Disposition"
                oSectionDataMapper.SaveDispo Encounter
            
            End Select
            
        End If
    Next
    
    

End Function

Private Function iMapperEncounter_UnlockSection(Params As DL_Support.iParameters) As Boolean
    Dim oEncLock As iEncLock
    
    If mMode = CDR And mMode <> ITT Then
        'Use stored procedures
        Set oEncLock = New EncLockStoredProc
    Else
        'Use SQL
        Set oEncLock = New EncLockSQL
    End If
    
    oEncLock.Init mConn, mCache, mMode, mUseCache

    iMapperEncounter_UnlockSection = oEncLock.UnlockSection(Params)
    
End Function

