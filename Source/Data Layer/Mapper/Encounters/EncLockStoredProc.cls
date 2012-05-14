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
Attribute VB_Name = "EncLockStoredProc"
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
    Dim objCmd As CHCSII_Command.CHCSII_Cmd
    Dim objSQL As ICHCSII_SQL.ISqlOps

    On Error GoTo ErrHandler
    
    Set param = Params

    'Use Stored Procs
    Set objCmd = New CHCSII_Command.CHCSII_Cmd

    objCmd.SPName = "ENCLOCK_LOCKSECTION"

    With param
    
        'We must add paramaters in the same order as the appear in the Stored Proc Signature
        Call objCmd.AddParam(False, False, "textString", .EncounterNumber)
        Call objCmd.AddParam(False, False, "textString", CLng(.FacilityNCID))
        Call objCmd.AddParam(False, False, "textString", .SectionNumber)
        Call objCmd.AddParam(False, False, "textString", .EventID)
        Call objCmd.AddParam(False, False, "textString", CLng(.UserNCID))
        Call objCmd.AddParam(False, False, "textString", .WorkStation)
        'Call objCmd.AddParam(False, False, "textString", "T")

    End With

    'This Parameter is defined as a cursor, used for output
    Call objCmd.AddParam(True, True, "", "")


    Set objSQL = mConn.CHCSII_SQLOPS_SP
    Call objSQL.ExecuteSP(objCmd)
    
    Set iEncLock_LockSection = objSQL
    
    Exit Function
    
ErrHandler:

End Function

Private Function iEncLock_UnlockSection(Params As DL_Support.iParameters) As Boolean
    Dim param As DM_Encounters.ParamsEncLock
    Dim objCmd As CHCSII_Command.CHCSII_Cmd
    Dim objSQL As ICHCSII_SQL.ISqlOps

    On Error GoTo ErrHandler
    
    Set param = Params

    'Use Stored Procs
    Set objCmd = New CHCSII_Command.CHCSII_Cmd
    
    objCmd.SPName = "ENCLOCK_UNLOCKSECTION"
    
    With param
    
        'We must add paramaters in the same order as the appear in the Stored Proc Signature
        Call objCmd.AddParam(False, False, "textString", .EncounterNumber)
        Call objCmd.AddParam(False, False, "textString", CLng(.FacilityNCID))
        Call objCmd.AddParam(False, False, "textString", .SectionNumber)
        Call objCmd.AddParam(False, False, "textString", .EventID)
        Call objCmd.AddParam(False, False, "textString", CLng(.UserNCID))
        Call objCmd.AddParam(False, False, "textString", .WorkStation)
        
    End With
    'This Parameter is defined as a cursor, used for output
    'Call objCmd.AddParam(True, True, "", "")
    
    Set objSQL = mConn.CHCSII_SQLOPS_SP
    Call objSQL.ExecuteSP(objCmd)

    iEncLock_UnlockSection = True
    
    Exit Function
    
ErrHandler:

'    Err.RAISE 100000, "Encounters DL",  "Error in EncounterOps UnlockSection: " & Err.Description, vbCritical
 '   E

End Function
'Public Function StartSectionUpdate(ByVal veSectionNumber As EncSectionEnum, ByVal vlEventID As Long, Optional ByVal bLocked As Boolean = False) As Boolean
'    Dim updateStr As String
'    Dim objCmd As CHCSII_Command.CHCSII_Cmd
'    Dim objGEMSQL As ICHCSII_SQL.ISqlOps
'
'    Dim retVal As String
'    On Error Resume Next
'
'
'TryAgain:
'
'    Set objCmd = New CHCSII_Command.CHCSII_Cmd
'
'    objCmd.SPName = "ENCLOCK_STARTUPDATE"
'
'    'We must add paramaters in the same order as the appear in the Stored Proc Signature
'    Call objCmd.AddParam(False, False, "textString", msEncounterNumber)
'    Call objCmd.AddParam(False, False, "textString", CLng(msFacilityNCID))
'    Call objCmd.AddParam(False, False, "textString", veSectionNumber)
'    Call objCmd.AddParam(False, False, "textString", vlEventID)
'    Call objCmd.AddParam(False, False, "textString", CLng(msUserNCID))
'    'Call objCmd.AddParam(False, False, "textString", retVal)
'
'    'This Parameter is defined as a cursor, used for output
'    Call objCmd.AddParam(True, True, "", "")
'
'    'Call mobjDASsp.ExecuteSP(objCmd)
'    Set mobjSQL = mconn.CHCSII_SQLOPS_SP     'SCR 36780 Phil Crowder 6/2/03
'    Call mobjSQL.ExecuteSP(objCmd)
'    If Err Then
'        'Err.RAISE 100000, "Encounters DL",  "Error setting start of encounter update: " & Err.Description, vbCritical, "Encounter Ops"
'    Else
'        mobjSQL.MoveFirst
'        If mobjSQL("ROW_COUNT") > 0 Then
'            StartSectionUpdate = True
'        End If
'    End If
'
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
'    'date:01/09/2004 SCR #:47870 developer: jrm Description: Added for ITT mode
'    'Auto is needed to determine ITT, or Theatre
'
'    Set objCmd = New CHCSII_Command.CHCSII_Cmd
'
'    objCmd.SPName = "ENCLOCK_SECTIONISLOCKED"
'
'    'We must add paramaters in the same order as the appear in the Stored Proc Signature
'    Call objCmd.AddParam(False, False, "textString", msEncounterNumber)
'    Call objCmd.AddParam(False, False, "textString", CLng(msFacilityNCID))
'    Call objCmd.AddParam(False, False, "textString", veSectionNumber)
'    Call objCmd.AddParam(False, False, "textString", vlEventID)
'
'    'This Parameter is defined as a cursor, used for output
'    Call objCmd.AddParam(True, True, "", "")
'
'    Set mobjSQL = mconn.CHCSII_SQLOPS_SP     'SCR 36780 Phil Crowder 6/2/03
'    Call mobjSQL.ExecuteSP(objCmd)
'
'    mobjSQL.MoveFirst
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
'    Set objCmd = New CHCSII_Command.CHCSII_Cmd
'
'    objCmd.SPName = "ENCLOCK_ENDUPDATE"
'
'    'We must add paramaters in the same order as the appear in the Stored Proc Signature
'    Call objCmd.AddParam(False, False, "textString", msEncounterNumber)
'    Call objCmd.AddParam(False, False, "textString", CLng(msFacilityNCID))
'    Call objCmd.AddParam(False, False, "textString", veSectionNumber)
'    Call objCmd.AddParam(False, False, "textString", vlEventID)
'    Call objCmd.AddParam(False, False, "textString", CLng(msUserNCID))
'
'    'This Parameter is defined as a cursor, used for output
'    'Call objCmd.AddParam(True, True, "", "")
'
'    Call mobjDASsp.ExecuteSP(objCmd)
'
'    If Err Then
'        Err.RAISE 100000, "Encounters DL",  "Error setting start of encounter update: " & Err.Description, vbCritical, "Encounter Ops"
'    Else
'        EndSectionUpdate = True
'    End If
'
'End Function
'
