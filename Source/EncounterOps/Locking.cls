VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Locking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum LockReasonEnum
    SectionLock = 1
    SignLock = 2
    CosignLock = 3
    AmendLock = 4
    AppendLock = 5
End Enum

Private Const MaxUpdateMinutes As Double = 5#
Private Const StatusLocked As String = "L"
Private Const StatusUpdating As String = "U"

Public msFacilityNCID As String
Public msEncounterNumber As String
Public msUserNCID As String

Private msWorkstation As String
'<<GEMS change
Private mobjSQL As ICHCSII_SQL.ISqlOps
Private mobjSQL2 As ICHCSII_SQL.ISqlOps
'>>
Private gemsConcept As GEMS_ConceptCtrl
Private oShared As New CWShared
Private mobjDASsp As ICHCSII_DataAccess.ICHCSII_DAS


Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Sub Class_Initialize()
    Dim nLen As Integer
    
    Set mobjSQL = gobjCHCSIIConn.CHCSII_SQLOPS(Auto) '''<SCR 36493
    
    Set gemsConcept = New GEMS_ConceptCtrl
    Set mobjSQL2 = gobjCHCSIIConn.CHCSII_SQLOPS(Auto)
   
    msWorkstation = Space(128)
    If GetComputerName(msWorkstation, Len(msWorkstation)) Then
        nLen = InStr(msWorkstation, Chr(0))
        If nLen > 0 Then
            msWorkstation = Left$(msWorkstation, nLen - 1)
        End If
    End If
    msWorkstation = Trim$(msWorkstation)
    If msWorkstation = "" Then
        msWorkstation = "Unknown"
    End If
End Sub

Private Sub Class_Terminate()
    Set mobjSQL = Nothing
    Set mobjSQL2 = Nothing
End Sub

Public Function LockSection(ByVal veSectionNumber As EncounterSectionEnum, _
                            ByVal vlEventID As Long, _
                            ByVal veReason As LockReasonEnum, _
                            Optional ByVal vsItemName As String = "", _
                            Optional ByVal vbOpeningSO As Boolean = False) As Boolean
    Dim sSelect As String
    Dim sMsg As String
    Dim sDeleteWhere As String
    Dim sLockUserNcid As String
    Dim sLockUserName As String
    Dim sLockSectionNumber As String
    Dim sLockEventID As String
    Dim sLockItemName As String
    Dim sLockWorkstation As String
    Dim nMyRetries As Integer
    Dim bUnbreakable As Boolean
    Dim nLocks As Integer
    Dim bDeleteLock As Boolean
    Dim objCmd As CHCSII_Command.CHCSII_Cmd
    
    On Error Resume Next
    
    If oShared.UseMySQLDB Then
        sSelect = "select encounter_lock.*,time_to_sec(timediff(sysdate(), date_time))/60 as lock_minutes" _
                & " from encounter_lock"
    ElseIf oShared.UseSQLServer Then
        sSelect = "select lock_minutes.*, 24 * 60 * ( DATEDIFF(d, GETDATE(),lock_minutes.date_time) )" _
                & " from encounter_lock as lock_minutes"
    Else
        sSelect = "select encounter_lock.*, 24 * 60 * (sysdate - date_time) as lock_minutes" _
                & " from encounter_lock"
    End If

    
    sSelect = sSelect & " where facility_ncid = " & msFacilityNCID _
        & " and encounter_number = " & msEncounterNumber
        
        
    If veSectionNumber <> 99 Then
        sSelect = sSelect & " and ((section_number = 99 and event_id = 0) or" _
            & " (section_number = " & veSectionNumber & " and event_id = " & vlEventID & "))"
    End If
    


    If Not vbOpeningSO Then
        sSelect = sSelect & " and section_number <> 96"
    End If

TryAgain:

    'date:01/09/2004 SCR #:47870 developer: jrm Description: Added for ITT mode
    'Auto is needed to determine ITT, or Theatre

    If oShared.IsAppMode(modeCDR) And oShared.IsAppMode(modeITT) = False Then
        Set mobjSQL = gobjCHCSIIConn.CHCSII_SQLOPS_SP     'SCR 36780 Phil Crowder 6/2/03
        Call mobjSQL.ExecuteSP(objCmd)
        'Call mobjDASsp.ExecuteSP(objCmd)
    Else
        Set mobjSQL = gobjCHCSIIConn.CHCSII_SQLOPS(Auto)
        mobjSQL.Execute "insert into Encounter_Lock(facility_ncid, encounter_number, section_number," _
            & " event_id, user_ncid, item_name, workstation, date_time, reason, status)" _
            & " select " & msFacilityNCID & ", " & msEncounterNumber & ", " & veSectionNumber & ", " _
            & vlEventID & ", " & msUserNCID & ", '" & vsItemName & "', '" & msWorkstation & "', " & gobjCHCSIIConn.SQLDate(Now) & " , " _
            & veReason & ", '" & StatusLocked & "' from dual where not exists (" & sSelect & ")"
            
        mobjSQL.Execute sSelect
    End If
    mobjSQL.MoveFirst
    If mobjSQL.EOF Then
        If nMyRetries > 10 Then
            sMsg = "Could not lock "
            If sLockSectionNumber = 99 Then
                sMsg = sMsg & "this"
            ElseIf vsItemName <> "" Then
                sMsg = sMsg & SectionEnumToName(veSectionNumber) & ":" & vsItemName & " in this"
            Else
                sMsg = sMsg & "the " & SectionEnumToName(veSectionNumber) & " section of this"
            End If
            MsgBxARMd sMsg & " encounter.", vbInformation, "Encounter"
            Exit Function 'SCR-26095
        Else
            nMyRetries = nMyRetries + 1
            GoTo TryAgain
        End If
    End If

    If Not vbOpeningSO Then
        sMsg = ""
        sDeleteWhere = ""
        bUnbreakable = False
        nLocks = 0
        LockSection = False
        
        'Scott Foster - 4/14/03 SCR 32981 Moved this code outside of the loop
        mobjSQL.MoveFirst
        '*********
        Do While Not mobjSQL.EOF
            sLockUserNcid = mobjSQL("USER_NCID")
            sLockSectionNumber = mobjSQL("SECTION_NUMBER")
            sLockEventID = mobjSQL("EVENT_ID")
            sLockItemName = mobjSQL("ITEM_NAME")
            sLockWorkstation = mobjSQL("WORKSTATION")
            ' We consider a lock ours if it is for our particular section or for the whole encounter
            If sLockUserNcid = msUserNCID _
            And ((sLockSectionNumber = veSectionNumber And sLockEventID = vlEventID) _
                    Or (sLockSectionNumber = 99 And sLockEventID = 0)) _
            And sLockWorkstation = msWorkstation Then
                LockSection = True
            Else
                'Multiple SO note can be open for an encounter
                If sLockSectionNumber <> eSOOpen Then
                    'Mulitple new text notes can be open for an encounter 'SCR-35244
                    If Not (veSectionNumber = eNotes And vlEventID = 0) Then
                        nLocks = nLocks + 1
                        gemsConcept.UniqueId = sLockUserNcid
                        sLockUserName = gemsConcept.PrefRep("2000").Representation
                        If sLockSectionNumber = 99 Then
                            sMsg = sMsg & "This"
                        ElseIf sLockItemName <> "" Then
                            sMsg = sMsg & SectionEnumToName(sLockSectionNumber) & ":" & sLockItemName & " in this"
                        Else
                            sMsg = sMsg & "The " & SectionEnumToName(sLockSectionNumber) & " section of this"
                        End If
                        sMsg = sMsg & " encounter was locked by " & sLockUserName & " on workstation '" _
                            & sLockWorkstation & "' at " & Format$(mobjSQL("DATE_TIME"), gsDefaultDateTime) _
                            & "." & vbCrLf
                        If mobjSQL("STATUS") = StatusUpdating And mobjSQL("LOCK_MINUTES") < MaxUpdateMinutes Then
                            bUnbreakable = True
                        Else
                            If sDeleteWhere <> "" Then
                                sDeleteWhere = sDeleteWhere & " or "
                            End If
                            sDeleteWhere = sDeleteWhere & "(section_number = " & sLockSectionNumber _
                                & " and event_id = " & sLockEventID & ")"
                        End If
                    End If
                End If
            End If
            mobjSQL.MoveNext
        Loop
        
        Set mobjSQL = Nothing
        
        If nLocks > 0 Then
            If bUnbreakable Then
                If nLocks = 1 Then
                    sMsg = sMsg & vbCrLf & "This lock"
                Else
                    sMsg = sMsg & vbCrLf & "One or more of these locks"
                End If
                MsgBxARMd sMsg & " is for an update in progress and cannot be broken.", vbInformation, "Encounter"
            Else
                'Show custom lock message particular sections
                Select Case veSectionNumber
                    Case eap, eHeader
                        'SCR 51459 - Strengthen the lock warning message received for the provider header section so the user knows other's data may be affected by breaking the lock
                        sMsg = sMsg & vbCrLf & "You may take over ownership, but the other person's work will be adversely affected." & _
                                vbCrLf & vbCrLf & "Do you want to take over ownership?"
                        
                        Screen.MousePointer = vbDefault
                        
                        frmEncLock.txtMessage.Text = sMsg
                        goARM.StopAllCurrTx enmARMTxStatusFailed, enmCHCSIIDlgDetect_CustDlg, "CHCSIIEncounterOps.Locking.LockSection()"  'SCR 23017 MSolano 4/24/2002
                        frmEncLock.Show vbModal
                        Screen.MousePointer = vbHourglass
                        
                        If frmEncLock.bTakeOwnerShip Then
                            bDeleteLock = True
                        End If
                        
                    Case Else
                        If MsgBxARMd(sMsg & vbCrLf & "Do you want to break the lock(s)?", vbYesNo, "Encounter Locking") = vbYes Then
                            If MsgBxARMd(sMsg & vbCrLf & "Please confirm that you want to break the lock(s).", vbOKCancel, "Encounter Locking") = vbOK Then
                                bDeleteLock = True
                            End If
                        Else
                            bDeleteLock = False
                        End If
                End Select
                
                If bDeleteLock Then
                    'Need to delete any lock using Enc # SCR-29725
                    'Ergo cannot use Stored Proc
                    'Call UnlockSection(veSectionNumber, vlEventID)
                    mobjSQL2.Execute "delete from encounter_lock" _
                        & " where facility_ncid = " & msFacilityNCID _
                        & " and encounter_number = " & msEncounterNumber _
                        & " and (" & sDeleteWhere & ")"
                    GoTo TryAgain
                End If
            End If
        Else
            If Not LockSection Then
                GoTo TryAgain
            End If
        End If
    End If
End Function

Public Function SectionIsLocked(ByVal veSectionNumber As EncounterSectionEnum, ByVal vlEventID As Long) As Boolean
    Dim sSelect As String
    Dim sLockUserNcid As String
    Dim sLockUserName As String
    Dim sLockSectionNumber As String
    Dim sLockEventID As String
    Dim sLockWorkstation As String
    Dim objCmd As CHCSII_Command.CHCSII_Cmd

    On Error Resume Next
    
'    If oShared.IsAppMode(modeCHCSI_GUI) Then 'SCR 36493
'        SectionIsLocked = True
'        Exit Function
'    End If

    'date:01/09/2004 SCR #:47870 developer: jrm Description: Added for ITT mode
    'Auto is needed to determine ITT, or Theatre

    If oShared.IsAppMode(modeCDR) And oShared.IsAppMode(modeITT) = False Then
        Set objCmd = New CHCSII_Command.CHCSII_Cmd
        
        objCmd.SPName = "ENCLOCK_SECTIONISLOCKED"
        
        'We must add paramaters in the same order as the appear in the Stored Proc Signature
        Call objCmd.AddParam(False, False, "textString", msEncounterNumber)
        Call objCmd.AddParam(False, False, "textString", CLng(msFacilityNCID))
        Call objCmd.AddParam(False, False, "textString", veSectionNumber)
        Call objCmd.AddParam(False, False, "textString", vlEventID)
        
        'This Parameter is defined as a cursor, used for output
        Call objCmd.AddParam(True, True, "", "")
        
        Set mobjSQL = gobjCHCSIIConn.CHCSII_SQLOPS_SP     'SCR 36780 Phil Crowder 6/2/03
        Call mobjSQL.ExecuteSP(objCmd)

        mobjSQL.MoveFirst
    Else
        
        Set mobjSQL = gobjCHCSIIConn.CHCSII_SQLOPS(Auto)
        
        sSelect = "select USER_NCID, SECTION_NUMBER, EVENT_ID, WORKSTATION" _
            & " from encounter_lock" _
            & " where facility_ncid = " & msFacilityNCID _
            & " and encounter_number = " & msEncounterNumber
        If veSectionNumber <> 99 Then
            sSelect = sSelect & " and (section_number = 99 or" _
                & " (section_number = " & veSectionNumber & " and event_id = " & vlEventID & "))"
        End If
        mobjSQL.Execute sSelect
    End If
    'If mobjSQL.RecordCount = 1 Then
    If Not mobjSQL.EOF Then
        sLockUserNcid = mobjSQL("USER_NCID")
        sLockSectionNumber = mobjSQL("SECTION_NUMBER")
        sLockEventID = mobjSQL("EVENT_ID")
        sLockWorkstation = mobjSQL("WORKSTATION")
        ' We consider a lock ours if it is for our particular section or for the whole encounter
        If sLockUserNcid = msUserNCID _
        And ((sLockSectionNumber = veSectionNumber And sLockEventID = vlEventID) _
                Or (sLockSectionNumber = 99 And sLockEventID = 0)) _
        And sLockWorkstation = msWorkstation Then
            SectionIsLocked = True
        End If
    End If
    
    Set mobjSQL = Nothing
End Function

Public Function StartSectionUpdate(ByVal veSectionNumber As EncounterSectionEnum, ByVal vlEventID As Long, Optional ByVal bLocked As Boolean = False) As Boolean
    Dim updateStr As String
    Dim objCmd As CHCSII_Command.CHCSII_Cmd
    Dim objGEMSQL As ICHCSII_SQL.ISqlOps

    Dim retVal As String
    On Error Resume Next

'    If oShared.IsAppMode(modeCHCSI_GUI) Then 'SCR 36493
'        StartSectionUpdate = True
'        Exit Function
'    End If
    
TryAgain:
    
    
'date:01/09/2004 SCR #:47870 developer: jrm Description: Added for ITT mode
'Auto is needed to determine ITT, or Theatre
    
    Set objGEMSQL = gobjCHCSIIConn.CHCSII_SQLOPS(Auto) '''<SCR 36493

    If Not bLocked Then
        If SectionIsLocked(veSectionNumber, vlEventID) Then
            bLocked = True
        End If
    End If

    If bLocked Then
        updateStr = "update encounter_lock set status = '" & StatusUpdating & "'" _
            & " where facility_ncid = " & msFacilityNCID _
            & " and encounter_number = " & msEncounterNumber _
            & " and ((section_number = 99 and event_id = 0)" _
            & "   or (section_number = " & veSectionNumber & " and event_id = " & vlEventID & "))"
        
        objGEMSQL.Execute updateStr
        
        Set objGEMSQL = Nothing
        
        If Err Then
            MsgBxARMd "Error setting start of encounter update: " & Err.Description, vbCritical, "Encounter Ops"
        Else
            StartSectionUpdate = True
        End If
    End If
    
    If Err Then
        'DB error, 27032 & -2147220604 error causes Shared to attempt to reconnect to CDR SCR-24083
        Call oShared.ShowVBError(-2147220604, "Database error.", "EncounterLocking", "EncounterOps", vbExclamation)
        
        'Check for Successful Reconnect
        If Err = 0 Then
            GoTo TryAgain
        End If
        
        Err.Clear
    End If
    
    Set mobjSQL = Nothing
    
End Function

Public Function EndSectionUpdate(ByVal veSectionNumber As EncounterSectionEnum, ByVal vlEventID As Long, Optional ByVal bLocked As Boolean = False) As Boolean
    
    Dim objCmd As CHCSII_Command.CHCSII_Cmd
    Dim objGEMSQL As ICHCSII_SQL.ISqlOps
    
    'date:01/09/2004 SCR #:47870 developer: jrm Description: Added for ITT mode
    'Auto is needed to determine ITT, or Theatre
        
    Set objGEMSQL = gobjCHCSIIConn.CHCSII_SQLOPS(Auto) '''<SCR 36493

    If Not bLocked Then
        If SectionIsLocked(veSectionNumber, vlEventID) Then
            bLocked = True
        End If
    End If
    Dim updateStr As String
    
    If bLocked Then
        updateStr = "update encounter_lock set status = '" & StatusLocked & "'" _
            & " where facility_ncid = " & msFacilityNCID _
            & " and encounter_number = " & msEncounterNumber _
            & " and ((section_number = 99 and event_id = 0)" _
            & "   or (section_number = " & veSectionNumber & " and event_id = " & vlEventID & "))"
            
        objGEMSQL.Execute updateStr
        
        Set objGEMSQL = Nothing
            
        If Err Then
            MsgBxARMd "Error setting end of encounter update: " & Err.Description, vbCritical, "Encounter Ops"
        Else
            EndSectionUpdate = True
        End If
    End If
    
End Function

Public Function UnlockSection(ByVal veSectionNumber As EncounterSectionEnum, ByVal vlEventID As Long) As Boolean

    Dim objCmd As CHCSII_Command.CHCSII_Cmd

    On Error Resume Next
    
'    If oShared.IsAppMode(modeCHCSI_GUI) Then 'SCR 36493
'        UnlockSection = True
'        Exit Function
'    End If
    
    'date:01/09/2004 SCR #:47870 developer: jrm Description: Added for ITT mode
    'Auto is needed to determine ITT, or Theatre
    
    If oShared.IsAppMode(modeCDR) And oShared.IsAppMode(modeITT) = False Then
        Set objCmd = New CHCSII_Command.CHCSII_Cmd
        
        objCmd.SPName = "ENCLOCK_UNLOCKSECTION"
        
        'We must add paramaters in the same order as the appear in the Stored Proc Signature
        Call objCmd.AddParam(False, False, "textString", msEncounterNumber)
        Call objCmd.AddParam(False, False, "textString", CLng(msFacilityNCID))
        Call objCmd.AddParam(False, False, "textString", veSectionNumber)
        Call objCmd.AddParam(False, False, "textString", vlEventID)
        Call objCmd.AddParam(False, False, "textString", CLng(msUserNCID))
        Call objCmd.AddParam(False, False, "textString", msWorkstation)
        
        'This Parameter is defined as a cursor, used for output
        'Call objCmd.AddParam(True, True, "", "")
        
        Call mobjDASsp.ExecuteSP(objCmd)
    Else
            Set mobjSQL = gobjCHCSIIConn.CHCSII_SQLOPS(Auto)
            
            mobjSQL.Execute "delete from encounter_lock" _
            & " where facility_ncid = " & msFacilityNCID _
            & " and encounter_number = " & msEncounterNumber _
            & " and section_number = " & veSectionNumber _
            & " and event_id = " & vlEventID _
            & " and workstation = '" & msWorkstation _
            & "' and user_ncid = " & msUserNCID
    End If
    If Err Then
        MsgBxARMd "Error in EncounterOps UnlockSection: " & Err.Description, vbCritical
    Else
        UnlockSection = True
    End If
    
    Set mobjSQL = Nothing

End Function



