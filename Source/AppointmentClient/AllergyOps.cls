VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AllergyOps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mobjShared As CWShared
Private mobjComm As MMMHISComm.Comm

Public Enum AllergyTransactionType
    RefreshRequest = 0
    Synchronize = 1
End Enum
Private Function AddAllergyTransaction_SP(sPatientUnitNum As String) As Boolean
    Dim oDatabase As CHCSII_CONN.Conn
    Dim objCmd As CHCSII_Command.CHCSII_Cmd
    Dim rs As ADODB.Recordset

    On Error GoTo ErrHandler

    If mobjDASsp Is Nothing Then
        Set oDatabase = New CHCSII_CONN.Conn
        Set mobjDASsp = oDatabase.CHCSII_DAS
    End If
    
    Set objCmd = New CHCSII_Command.CHCSII_Cmd
    objCmd.SPName = "MT_ALLERGY_PKG" + ".PRC_REQUEST_ALLERGY_SYNC_CW"
    If Not Set_Parms_for_SP(objCmd, , , , , , , CLng(sPatientUnitNum)) Then
        AddAllergyTransaction_SP = False
        GoTo Cleanup
    End If
    
    Set rs = mobjDASsp.OpenRecordsetSP(objCmd)
    If Not rs Is Nothing Then
        If Not rs.BOF And Not rs.EOF Then
            rs.MoveFirst
        End If
    End If


    AddAllergyTransaction_SP = True
Cleanup:
Set oDatabase = Nothing
Set objCmd = Nothing
Set rs = Nothing

Exit Function
ErrHandler:
mobjShared.ShowVBError Err.Number, Err.Description, "AllergyOps.AddAllergyTransaction_SP", "CHCSII_AppointmentClient", vbExclamation
GoTo Cleanup
End Function


Private Function RefreshAllergiesFromCHCS_SP(sPatientUnitNum As String, lTransType As Long) As ADODB.Recordset
    Dim oDatabase As CHCSII_CONN.Conn
    Dim objCmd As CHCSII_Command.CHCSII_Cmd

    On Error GoTo ErrHandler

    If mobjDASsp Is Nothing Then
        Set oDatabase = New CHCSII_CONN.Conn
        Set mobjDASsp = oDatabase.CHCSII_DAS
    End If
    
    Set objCmd = New CHCSII_Command.CHCSII_Cmd
    objCmd.SPName = "MT_ALLERGY_PKG.PRC_REFRESH_ALLERGY_TRANS"
    If Not Set_Parms_for_SP(objCmd, , , , , , , CLng(sPatientUnitNum), lTransType) Then
        GoTo Cleanup
    End If
    
    Set RefreshAllergiesFromCHCS_SP = mobjDASsp.OpenRecordsetSP(objCmd)
Cleanup:
    Set oDatabase = Nothing
    Set objCmd = Nothing

    Exit Function
ErrHandler:
    mobjShared.ShowVBError Err.Number, Err.Description, "AllergyOps.RefreshAllergiesFromCHCS_SP", "CHCSII_AppointmentClient", vbExclamation
    GoTo Cleanup
End Function





Public Function AddAllergyTransaction(eTransactionType As AllergyTransactionType, _
                         ByRef p_oPatient As Object) As Boolean

    Const sSUBR_NAME$ = "AddAllergyTransaction"

    Dim sSql                As String
    Dim sSQL2               As String
    Dim objRecordset        As ADODB.Recordset
    Dim objRecordset2       As ADODB.Recordset
    Dim objCHCSII_DAS       As ICHCSII_DAS         'CHCS-CONN
    Dim oDatabase           As New CHCSII_CONN.Conn
    Dim sAllergyTransId     As String
    Dim i                   As Integer
    Dim objSQL              As ICHCSII_SQL.ISqlOps
    
    On Error GoTo ErrHandler
    
    If mobjShared.UseSQLServer Then
      AddAllergyTransaction = True
      Exit Function
    End If
    If p_oPatient Is Nothing Then
        AddAllergyTransaction = False
        MsgBxARMd "Invalid patient unit number sent to add allergy transaction."
        GoTo CleanUpAndExit
    End If
    'PERF STORED PROC
    If mobjShared.IsAppMode(modeCDR) Then
        If AddAllergyTransaction_SP(p_oPatient.UnitNumber) Then
           AddAllergyTransaction = True
        Else
           AddAllergyTransaction = False
        End If
        GoTo CleanUpAndExit
    End If
    'Get all patient ien's and allergy transaction record trans_id's
    sSql = "SELECT at.allergy_trans_id, ms.medical_record_number, ms.facility_cid, at.update_status " _
         & " FROM mmi_site ms, allergy_transaction at" _
         & " WHERE ms.unit_number = " & p_oPatient.UnitNumber _
         & " AND to_number(ms.medical_record_number) = at.patient_ien(+)" _
         & " AND ms.facility_cid = at.facility_ncid(+)" _
         & " AND " & eTransactionType & " = at.trans_type(+)"
         
    If eTransactionType = Synchronize Then
        sSql = sSql & " AND 0 = at.update_status(+)"
    End If
    'CHCS-CONN
    Set objCHCSII_DAS = oDatabase.CHCSII_DAS
    Set objRecordset = objCHCSII_DAS.OpenRecordset(sSql)
    
    If objRecordset Is Nothing Then
        GoTo CleanUpAndExit
    End If
    
    sSQL2$ = "SELECT cdrplus.allergy_trans_id_sequence.nextval FROM dual"
    Set objSQL = SelectSqlOps()
    
    With objRecordset
        For i = 1 To .RecordCount
        
            If .Fields("update_status") = "1" Then
                MsgBxARMd "Another user is currently refreshing the allergy record for this patient. The allergy " & _
                        "information you are viewing may not be up to date. Please wait several minutes then " & _
                        "refresh your screen to display the most recent allergy information.", vbOKOnly + vbInformation
                GoTo CleanUpAndExit
            End If
            
            If .Fields("allergy_trans_id") = 0 Then
            
                sAllergyTransId = objSQL.GetNextID(exMAINSEQ)
                
                If sAllergyTransId = "" Or sAllergyTransId = "0" Then
                    Err.Raise vbObjectError + 146, _
                             sSUBR_NAME$, _
                             "Next allergy transaction id returned: " & sAllergyTransId
                End If
                
                sSql = "INSERT INTO allergy_transaction (allergy_trans_id, unit_number, patient_ien, " _
                     & "facility_ncid, trans_type, trans_date, update_status) " _
                     & "VALUES (" & sAllergyTransId & ", " & p_oPatient.UnitNumber & ", " _
                     & .Fields("medical_record_number") & ", " _
                     & .Fields("facility_cid") & ", " _
                     & eTransactionType & ", '" & Format(Now, "YYYYMMDDHHNNSS") & "', 0)"
                objCHCSII_DAS.ExecuteSQL sSql
            Else
                sSql = "UPDATE allergy_transaction " + vbLf _
                     + "SET trans_date = '" & Format$(Now, "YYYYMMDDHHNNSS") _
                        + "', update_status = 0 " _
                     + "WHERE allergy_trans_id = " _
                        + .Fields("allergy_trans_id")
                
                objCHCSII_DAS.ExecuteSQL sSql
            End If
            
            .MoveNext
        Next
        .Close
    End With 'objRecordset
    
    AddAllergyTransaction = True
    GoTo CleanUpAndExit
    Exit Function
    
ErrHandler:
    Screen.MousePointer = vbDefault
    
    'TO DO: Change this to call the standard error message & logging routine?
    'Also add sSQL$ to error message.
    '-------------------------------------------------------------------------
    MsgBxARMd "In subroutine """ + sSUBR_NAME$ + """," + vbCrLf _
         + "Unable to add allergy transaction. " + vbCrLf _
         + "Synchronization with CHCS will not be performed." _
         + vbCrLf + "Error: " & Err.Description, _
         vbExclamation, _
         "Allergies"
    MsgBxARMd "SQL = " + sSql$
         
    AddAllergyTransaction = False
    
CleanUpAndExit:
    Set objRecordset = Nothing
    Set objRecordset2 = Nothing
    Set objCHCSII_DAS = Nothing
    Set oDatabase = Nothing
    Set objSQL = Nothing
End Function 'AddAllergyTransaction

Public Sub CancelRefresh()
    gbAllergySyncCancel = True
End Sub


Public Function RefreshAllergiesFromCHCS(ByVal bSuppressMsg As Boolean, _
                                          ByRef p_oPatient As Object, _
                                          ByVal p_lMyProgId As Long, _
                                          ByRef p_oLoginEnvironment As Object, _
                                          Optional sOutPutText As String) As Boolean

    Dim sSql                As String
    Dim sMsg                As String
    Dim i                   As Integer
    Dim iNumChecks          As Integer
    Dim bDone               As Boolean
    Dim sStatus             As String
    Dim mobjSQL2            As Object
    'Need to type as object here so that the calling function will have a
    'generic recordset to parse. mobjSQL is not the same type as returned
    'by SP method.

    
    
    On Error GoTo ErrHandler
    
    If mobjShared.UseSQLServer Then
      RefreshAllergiesFromCHCS = True
      GoTo CleanUpAndExit
    End If
    
    RefreshAllergiesFromCHCS = False
    sMsg = "The allergy refresh process is still not complete. Do wish to continue waiting?" & vbCrLf & vbCrLf & _
            "If this message continues to display please have your system administrator check that " & _
            "the allergy synchronization process is running on the application server."
    sOutPutText = sOutPutText + " Phase1 complete. " + vbCrLf
    If Not bSuppressMsg Then
       If MsgBxARMd("This operation can take up to one minute to complete. Do you wish to continue?", vbQuestion + vbYesNo, "Allergy Synchronization") = vbNo Then
            Screen.MousePointer = vbDefault
            GoTo CleanUpAndExit
       End If
    End If
    sOutPutText = sOutPutText + " Phase2 complete. " + vbCrLf
    'Poll every 5 seconds
    bDone = True
    gbAllergySyncCancel = False

    'frmAllergySyncStatus.Show 'SCR 9816
    
    Do
        Do:
            'Kdunne 9/12/2001 SCR 16425
            'Added DoEvents. Before, when user clicked cancel event,
            'system would wait until after this loop had processed
            'before it would process the cancel button.
            '
            DoEvents
            If gbAllergySyncCancel Then GoTo CleanUpAndExit
            Pause 1
            sStatus = sStatus & "."
            'Use the status information as a counter to determine when to exit from loop and check status
            If Len(sStatus) Mod 10 = 0 Then Exit Do
        Loop
        sOutPutText = sOutPutText + " Phase3 complete. " + vbCrLf
        
        
        'PERF STORED PROC
        If mobjShared.IsAppMode(modeCDR) Then
            Set mobjSQL2 = RefreshAllergiesFromCHCS_SP(p_oPatient.UnitNumber, CLng(RefreshRequest))
        Else
        
            sSql = "SELECT at.facility_ncid, at.update_status, at.err_text" _
                 & " FROM mmi_site ms, allergy_transaction at" _
                 & " WHERE ms.unit_number = " & p_oPatient.UnitNumber _
                 & " AND to_number(ms.medical_record_number) = at.patient_ien(+)" _
                 & " AND ms.facility_cid = at.facility_ncid(+)" _
                 & " AND " & RefreshRequest & " = at.trans_type(+)"
            'CHCS-CONN
            
            Set mobjSQL2 = SelectSqlOps()
            If mobjSQL2 Is Nothing Then
                sOutPutText = sOutPutText + vbCrLf + " (mobjSQL2 not set) "
            End If
            If p_oLoginEnvironment Is Nothing Then
                sOutPutText = sOutPutText + vbCrLf + " (p_oLoginEnvironment not set) "
            End If
            mobjSQL2.Execute sSql
        End If
        
        If mobjSQL2 Is Nothing Then
            sOutPutText = sOutPutText + vbCrLf + " (: mobjSQL2 not set) "
            GoTo ErrHandler
        End If
        
        
        sOutPutText = sOutPutText + " Phase4 complete. " + vbCrLf
        If mobjSQL2.BOF = False And mobjSQL2.EOF = False Then
            sOutPutText = sOutPutText + " Phase5 complete. " + vbCrLf
            mobjSQL2.MoveFirst
            For i = 1 To mobjSQL2.RecordCount
                'Only check refresh requests handled by the mid-tier for your facility
                If IsNull(mobjSQL2("update_status")) Then
                    sOutPutText = sOutPutText + "update_status is Null" + vbCrLf
                    Exit For
                End If
                If IsNull(mobjSQL2("Facility_NCID")) Then
                    sOutPutText = sOutPutText + "Facility_NCID is Null" + vbCrLf
                    Exit For
                End If
                If IsNumeric(mobjSQL2("update_status")) = False Then
                    sOutPutText = sOutPutText + "update_status is non-numeric" + vbCrLf
                    Exit For
                End If
                
                If (CLng(mobjSQL2("update_status")) < 2) And _
                    CStr(mobjSQL2("Facility_NCID")) = CStr(p_oLoginEnvironment.FacilityNCID) Then
                    bDone = False
                    Exit For
                End If
                mobjSQL2.MoveNext
            Next
            sOutPutText = sOutPutText + " Phase6 complete. " + vbCrLf
        End If
        iNumChecks = iNumChecks + 1
        
        If bDone Then
           RefreshAllergiesFromCHCS = True
           Exit Do
        End If
        sOutPutText = sOutPutText + " Phase7 complete. " + vbCrLf
        If iNumChecks = 9 Then
           'Prompt do wish to continue waiting
           If MsgBxARMd(sMsg, vbQuestion + vbYesNo, "Allergy Synchronization") = vbNo Then
              Exit Do
           End If
           sStatus = ""
           iNumChecks = 0
        End If
        sOutPutText = sOutPutText + " Phase8 complete. " + vbCrLf
        bDone = True 'Flag must be reset before looping otherwise refresh will never complete
    Loop

    'Send refresh to Core when refreshing allergies SCR-35556 & 38254
    'needs to be called to update visual indicators
    sOutPutText = sOutPutText + " Phase9 complete. " + vbCrLf
    mobjComm.Message cwiREFRESH_ALLERGIES, "", NCID_CW_APPLICATION, p_lMyProgId
    sOutPutText = sOutPutText + " Phase10 complete. " + vbCrLf
    GoTo CleanUpAndExit
    
ErrHandler:
    
    MsgBxARMd "Unable to Refresh allergies." & vbCrLf _
         + "Synchronization with CHCS will not be performed." _
         + vbCrLf + "Debug Listing: " + vbCrLf + sOutPutText + vbCrLf & "Error: " & Err.Description, _
         vbExclamation, "Allergies"
    
CleanUpAndExit:
    sOutPutText = sOutPutText + " Phase11 complete. " + vbCrLf
    Set mobjSQL2 = Nothing
        
End Function 'RefreshAllergiesFromCHCS

Private Sub Pause(nSeconds As Single)

    'Code taken from an MSDN article
    
    Dim t0 As Single
    
    t0 = Timer
    
    Do While Timer - t0 < nSeconds
        DoEvents
        If Timer < t0 Then  'Back up 1 day
            t0 = t0 - CLng(24) * CLng(60) * CLng(60)
        End If
    Loop

End Sub

Private Sub Class_Initialize()
  Set mobjShared = New CWShared
End Sub

Private Sub Class_Terminate()
  Set mobjShared = Nothing
  Set mobjComm = Nothing
End Sub
Public Sub Initialize(ByRef robjComm As Object)
    Set mobjComm = robjComm
End Sub

