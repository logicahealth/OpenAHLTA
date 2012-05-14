VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EncLockSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements iEncLock

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

Private Function iEncLock_LockSection(Params As DL_Support.iParameters) As ISqlOps
    Dim param As DM_Encounters.ParamsEncLock
    Dim oSQL As ICHCSII_SQL.ISqlOps
    Dim sSelect As String
    
    Set param = Params
    
    If mMode = MYSQL Then
        sSelect = "select encounter_lock.*, time_to_sec(timediff(sysdate(),date_time)) /60 as lock_minutes" _
                & " from encounter_lock"
    ElseIf mMode = ITT Then
        'SQL Server
        sSelect = "select lock_minutes.*, 24 * 60 * ( DATEDIFF(d, GETDATE(),lock_minutes.date_time) )" _
                & " from encounter_lock as lock_minutes"
    Else
        sSelect = "select encounter_lock.*, 24 * 60 * (sysdate - date_time) as lock_minutes" _
                & " from encounter_lock"
    End If

    With param
    
        sSelect = sSelect & " where facility_ncid = " & .FacilityNCID _
            & " and encounter_number = " & .EncounterNumber
     
        If .SectionNumber <> 99 Then
            sSelect = sSelect & " and ((section_number = 99 and event_id = 0) or" _
                & " (section_number = " & .SectionNumber & " and event_id = " & .EventID & "))"
        End If
    
        'not opening SO
        If .SectionNumber <> 96 Then
            sSelect = sSelect & " and section_number <> 96"
        End If
    
        oSQL.Execute "insert into Encounter_Lock(facility_ncid, encounter_number, section_number," _
            & " event_id, user_ncid, item_name, workstation, date_time, reason, status)" _
            & " select " & .FacilityNCID & ", " & .EncounterNumber & ", " & .SectionNumber & ", " _
            & .EventID & ", " & .UserNCID & ", '" & .ItemName & "', '" & .WorkStation & "', " & mConn.SQLDate(Now) & " , " _
            & .Reason & ", '" & STATUS_LOCKED & "' from dual where not exists (" & sSelect & ")"
    
        oSQL.Execute sSelect

    End With

    Set iEncLock_LockSection = oSQL
    
End Function

Private Function iEncLock_UnlockSection(Params As DL_Support.iParameters) As Boolean
    Dim param As DM_Encounters.ParamsEncLock
    Dim oSQL As ICHCSII_SQL.ISqlOps
    Dim sSelect As String
    
    On Error GoTo ErrHandler
    
    Set param = Params
    
    With param
        oSQL.Execute "delete from encounter_lock" _
            & " where facility_ncid = " & .FacilityNCID _
            & " and encounter_number = " & .EncounterNumber _
            & " and section_number = " & .SectionNumber _
            & " and event_id = " & .EventID _
            & " and workstation = '" & .WorkStation _
            & "' and user_ncid = " & .UserNCID
    End With
    
    iEncLock_UnlockSection = True

    Exit Function
    
ErrHandler:
    
End Function

'Public Function StartSectionUpdate(ByVal veSectionNumber As EncSectionEnum, ByVal vlEventID As Long, Optional ByVal bLocked As Boolean = False) As Boolean
'    Dim updateStr As String
'    Dim objCmd As CHCSII_Command.CHCSII_Cmd
'    Dim objGEMSQL As ICHCSII_SQL.ISqlOps
'
'    Dim retVal As String
'    On Error Resume Next
'
'TryAgain:
'
'        Set objGEMSQL = mconn.CHCSII_SQLOPS(Auto) '''<SCR 36493
'
'        If Not bLocked Then
'            If SectionIsLocked(veSectionNumber, vlEventID) Then
'                bLocked = True
'            End If
'        End If
'
'        If bLocked Then
'            updateStr = "update encounter_lock set status = '" & StatusUpdating & "'" _
'                & " where facility_ncid = " & msFacilityNCID _
'                & " and encounter_number = " & msEncounterNumber _
'                & " and ((section_number = 99 and event_id = 0)" _
'                & "   or (section_number = " & veSectionNumber & " and event_id = " & vlEventID & "))"
'
'            objGEMSQL.Execute updateStr
'
'            Set objGEMSQL = Nothing
'
'            If Err Then
'                Err.RAISE 100000, "Encounters DL",  "Error setting start of encounter update: " & Err.Description, vbCritical, "Encounter Ops"
'            Else
'                StartSectionUpdate = True
'            End If
'        End If
'
'    If Err Then
'        'DB error, 27032 & -2147220604 error causes Shared to attempt to reconnect to CDR SCR-24083
'        Call oShared.ShowVBError(-2147220604, "Database error.", "EncounterLocking", "EncounterOps", vbExclamation)
'
'        'Check for Successful Reconnect
'        If Err = 0 Then
'            GoTo TryAgain
'        End If
'
'        Err.Clear
'    End If
'
'End Function
'
'Public Function SectionIsLocked(ByVal veSectionNumber As EncSectionEnum, ByVal vlEventID As Long) As Boolean
'    Dim sSelect As String
'    Dim sLockUserNcid As String
'    Dim sLockUserName As String
'    Dim sLockSectionNumber As String
'    Dim sLockEventID As String
'    Dim sLockWorkstation As String
'    Dim objCmd As CHCSII_Command.CHCSII_Cmd
'
'    On Error Resume Next
'
'
'    'date:01/09/2004 SCR #:47870 developer: jrm Description: Added for ITT mode
'    'Auto is needed to determine ITT, or Theatre
'
'
'    sSelect = "select USER_NCID, SECTION_NUMBER, EVENT_ID, WORKSTATION" _
'        & " from encounter_lock" _
'        & " where facility_ncid = " & msFacilityNCID _
'        & " and encounter_number = " & msEncounterNumber
'    If veSectionNumber <> 99 Then
'        sSelect = sSelect & " and (section_number = 99 or" _
'            & " (section_number = " & veSectionNumber & " and event_id = " & vlEventID & "))"
'    End If
'    mobjSQL.Execute sSelect
'
'    'If mobjSQL.RecordCount = 1 Then
'    If Not mobjSQL.EOF Then
'        sLockUserNcid = mobjSQL("USER_NCID")
'        sLockSectionNumber = mobjSQL("SECTION_NUMBER")
'        sLockEventID = mobjSQL("EVENT_ID")
'        sLockWorkstation = mobjSQL("WORKSTATION")
'        ' We consider a lock ours if it is for our particular section or for the whole encounter
'        If sLockUserNcid = msUserNCID _
'        And ((sLockSectionNumber = veSectionNumber And sLockEventID = vlEventID) _
'                Or (sLockSectionNumber = 99 And sLockEventID = 0)) _
'        And sLockWorkstation = msWorkstation Then
'            SectionIsLocked = True
'        End If
'    End If
'
'End Function
'
'
'Public Function EndSectionUpdate(ByVal veSectionNumber As EncSectionEnum, ByVal vlEventID As Long, Optional ByVal bLocked As Boolean = False) As Boolean
'
'    Dim objCmd As CHCSII_Command.CHCSII_Cmd
'    Dim objGEMSQL As ICHCSII_SQL.ISqlOps
'
'    'date:01/09/2004 SCR #:47870 developer: jrm Description: Added for ITT mode
'    'Auto is needed to determine ITT, or Theatre
'
'    Set objGEMSQL = mconn.CHCSII_SQLOPS(Auto) '''<SCR 36493
'
'    If Not bLocked Then
'        If SectionIsLocked(veSectionNumber, vlEventID) Then
'            bLocked = True
'        End If
'    End If
'    Dim updateStr As String
'
'    If bLocked Then
'        updateStr = "update encounter_lock set status = '" & StatusLocked & "'" _
'            & " where facility_ncid = " & msFacilityNCID _
'            & " and encounter_number = " & msEncounterNumber _
'            & " and ((section_number = 99 and event_id = 0)" _
'            & "   or (section_number = " & veSectionNumber & " and event_id = " & vlEventID & "))"
'
'        objGEMSQL.Execute updateStr
'
'        Set objGEMSQL = Nothing
'
'        If Err Then
'            Err.RAISE 100000, "Encounters DL",  "Error setting end of encounter update: " & Err.Description, vbCritical, "Encounter Ops"
'        Else
'            EndSectionUpdate = True
'        End If
'    End If
'End Function
'
