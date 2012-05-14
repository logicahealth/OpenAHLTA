VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MYSQL_DAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'-- Facility workaround is not included since this dll will only be used in the case
'   of a single facility.

Implements ICHCSII_DataAccess.ICHCSII_DAS
Implements ISec_SQL.ISecSqlOps
Implements ICHCSII_DataAccess.I_DAS_Direct
Implements ICHCSII_DataAccess.IADOShape
Implements IADO_DAS

Private msDBKey As String

Private m_saRelevantTables() As String
'Private goMySQLCnn                As ADODB.Connection

'3-24-04 These have been replaced by functions called CurrentUserNCID and CurrentUserName
'Private msUserNCID           As String
'Private msUserName           As String

Private mNoAudit             As Collection
Private Const Module       As String = "MYSQL_DAS Class"

#If DEBUGON Then
  Public DL                  As cDebugLog
  Private Const Exe          As String = "MYSQL_DAS"
#End If

Private Property Get I_DAS_Direct_ActiveConnection() As ADODB.Connection
    Set I_DAS_Direct_ActiveConnection = goMySQLCnn
End Property

Private Function I_DAS_Direct_BeginTrans() As Boolean
    goMySQLCnn.BeginTrans
End Function

Private Function I_DAS_Direct_CommitTrans() As Boolean
    goMySQLCnn.CommitTrans
End Function

Private Function I_DAS_Direct_ExecuteSQL(ByVal SQL As String, Optional ByVal UserID As String, Optional ByVal RecordDateTime As Date) As Variant
    I_DAS_Direct_ExecuteSQL = ExecuteSQL(SQL, UserID, RecordDateTime)
End Function

Private Function I_DAS_Direct_RollbackTrans() As Boolean
    goMySQLCnn.RollbackTrans
End Function

Private Property Get IADO_DAS_ActiveConnection() As ADODB.Connection

On Error GoTo ErrHnd

    'Call ICHCSII_DAS
    Set IADO_DAS_ActiveConnection = ICHCSII_DAS_ActiveConnection

Exit Property

ErrHnd:
    Err.Raise Err.Number, Err.Source, Err.Description

End Property

Private Function IADO_DAS_BeginTrans() As Boolean

On Error GoTo ErrHnd

    'Call ICHCSII_DAS
    IADO_DAS_BeginTrans = ICHCSII_DAS_BeginTrans
    
Exit Function

ErrHnd:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function


Private Function IADO_DAS_CommitTrans() As Boolean

On Error GoTo ErrHnd

    'Call ICHCSII_DAS
    IADO_DAS_CommitTrans = ICHCSII_DAS_CommitTrans
    
Exit Function

ErrHnd:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function


Private Function IADO_DAS_CreateRecordsetFromFile(sFileName As String, sDelimiter As String) As ADODB.Recordset

On Error GoTo ErrHnd

    'Call ICHCSII_DAS
    Set IADO_DAS_CreateRecordsetFromFile = ICHCSII_DAS_CreateRecordsetFromFile(sFileName, sDelimiter)

Exit Function

ErrHnd:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function


Private Function IADO_DAS_ExecuteSP(oCommand As ADODB.Command) As Long
On Error GoTo ErrHnd

'3-09-04 SF
'This function takes the ADO command object and executes it

    oCommand.ActiveConnection = goMySQLSPCnn
    
    oCommand.Execute
    
Exit Function

ErrHnd:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function


Private Function IADO_DAS_ExecuteSQL(SQL As String) As Variant

On Error GoTo ErrHnd

    'Call ICHCSII_DAS
    Set IADO_DAS_ExecuteSQL = ICHCSII_DAS_ExecuteSQL(SQL)

Exit Function

ErrHnd:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function


Private Property Let IADO_DAS_FacilityWorkaround(ByVal RHS As Boolean)

On Error GoTo ErrHnd

    Err.Raise 1001, "IADO_DAS", "FacilityWorkaround is not supported in ICD DAS."

Exit Property

ErrHnd:
    Err.Raise Err.Number, Err.Source, Err.Description

End Property

Private Function IADO_DAS_GetNextID(Optional ByVal eSeqName As eSeq = 0&) As Variant

On Error GoTo ErrHnd

    'Call ICHCSII_DAS
    Set IADO_DAS_GetNextID = ICHCSII_DAS_GetNextID(eSeqName)

Exit Function

ErrHnd:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function IADO_DAS_GetNextNcid() As Variant

On Error GoTo ErrHnd

    'Call ICHCSII_DAS
    IADO_DAS_GetNextNcid = ICHCSII_DAS_GetNextNcid

Exit Function

ErrHnd:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function


Private Function IADO_DAS_OpenRecordset(SQL As String, Optional ByVal CursorType As ADODB.CursorTypeEnum = 0&, Optional ByVal LockType As ADODB.LockTypeEnum = 1&) As ADODB.Recordset

On Error GoTo ErrHnd

    'Call ICHCSII_DAS
    Set IADO_DAS_OpenRecordset = ICHCSII_DAS_OpenRecordset(SQL, CursorType, LockType)

Exit Function

ErrHnd:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function


Private Function IADO_DAS_OpenRecordsetSP(oCommand As ADODB.Command) As ADODB.Recordset

On Error GoTo ErrHnd

'3-09-04 SF
'This function takes the ADO command object and executes it

    oCommand.ActiveConnection = goMySQLSPCnn
    
    Set IADO_DAS_OpenRecordsetSP = oCommand.Execute
    
Exit Function

ErrHnd:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function


Private Function IADO_DAS_RollbackTrans() As Boolean
    
On Error GoTo ErrHnd

    'Call ICHCSII_DAS
    IADO_DAS_RollbackTrans = ICHCSII_DAS_RollbackTrans

Exit Function

ErrHnd:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function


Private Property Let IADO_DAS_UserNCID(RHS As String)

On Error GoTo ErrHnd

    'Call ICHCSII_DAS
    ICHCSII_DAS_UserNCID = RHS

Exit Property

ErrHnd:
    Err.Raise Err.Number, Err.Source, Err.Description

End Property


Private Function IADOShape_ShapeConnect() As String
    
    IADOShape_ShapeConnect = GetConnectionString(True)
    
End Function

Private Property Get ICHCSII_DAS_ActiveConnection() As ADODB.Connection
    
    Set ICHCSII_DAS_ActiveConnection = goMySQLCnn

End Property

Private Function ICHCSII_DAS_CreateRecordsetFromFile(sFileName As String, sDelimiter As String) As ADODB.Recordset

Dim iFileNum As Integer
Dim i As Integer
Dim lRow As Long
Dim sData As String
Dim sHeaderRow As String
Dim sValues() As String
Dim sColumnNames() As String
Dim bIsOpen As Boolean
Dim objRecordset As New ADODB.Recordset

iFileNum = FreeFile

Open sFileName For Input As #iFileNum

Do While Not EOF(iFileNum)
    Line Input #iFileNum, sData
    If Left(sData, 3) <> "***" Then
        'Assume first good row is column headers
        If lRow = 0 Then
            sHeaderRow = sData
            lRow = lRow + 1
        Else
            sValues = Split(sData, sDelimiter)
            If lRow = 1 Then
                'Do headers
                sColumnNames = Split(sHeaderRow, sDelimiter)
                
                For i = LBound(sValues) To UBound(sValues)
                    If sColumnNames(i) <> "" Then
                        If IsDate(sValues(i)) = True And InStr(1, sValues(i), "-") < 1 Then
                            objRecordset.fields.Append sColumnNames(i), adDate
                        Else
                            objRecordset.fields.Append sColumnNames(i), adBSTR
                        End If
                    End If
                Next
                
            End If
            
            'Do Data
            If Not (bIsOpen) Then
                objRecordset.Open
                bIsOpen = True
            End If
            
            objRecordset.AddNew
            For i = 0 To objRecordset.fields.Count - 1
                If objRecordset.fields.Item(i).Type = adDate Then
                    objRecordset.fields.Item(i).Value = CDate(sValues(i))
                Else
                    If sValues(i) <> "" Then
                        objRecordset.fields.Item(i).Value = CStr(sValues(i))
                    End If
                End If
            Next
            
            lRow = lRow + 1
        End If
    End If
Loop

If bIsOpen Then
    objRecordset.MoveFirst
End If

Close #iFileNum

Set ICHCSII_DAS_CreateRecordsetFromFile = objRecordset

End Function

Private Property Let ICHCSII_DAS_FacilityWorkaround(ByVal RHS As Boolean)

    Err.Raise 1001, "MYSQL_DAS", "FacilityWorkaround is not supported in GEMS DAS."
    
End Property


Private Function ICHCSII_DAS_GetNextID(Optional ByVal eSeqName As ICHCSII_DataAccess.eSeq = esMAINSEQ) As Variant
'-- pulls the next value from the MainSeq sequence

On Error GoTo ErrHandler
    
Dim oRs As New ADODB.Recordset
Dim sSeqName As String
Dim lSeqNum As Long
    
    sSeqName = SeqEnumToString(eSeqName)
    goMySQLCnn.Execute "UPDATE " & sSeqName & " SET id=LAST_INSERT_ID(id+1)"
    oRs.Open "SELECT LAST_INSERT_ID()", goMySQLCnn, adOpenForwardOnly, adLockOptimistic, adCmdText

    If Not oRs.EOF Then
        ICHCSII_DAS_GetNextID = oRs(0)
    End If
    
Func_Exit:
  On Error Resume Next
  oRs.Close
  Set oRs = Nothing
  Exit Function
  
ErrHandler:
  '<< SCR #44407;   Developer: Brian Mowbray 01/13/2004 02:07 PM
  Call HandleDBError(Err.Number, Err.Description, Module, goMySQLCnn)
  'This was added to trap for error raised by the HandleDBError if the error is
  'not related to the loss of connectivity.
  If Err.Number <> 0 Then
    Err.Raise Err.Number, Module, Err.Description
  End If
  Resume Func_Exit
  
End Function

Private Function ICHCSII_DAS_GetNextNcid() As Variant

'-- pulls the next value from the NCID sequence
    Dim oRs As ADODB.Recordset
    Set oRs = New ADODB.Recordset
    Dim lSeqNum As Long
    
    On Error GoTo Func_Err
    
'<< SRC 29784     the NCIDs and MainSeq sequencers were colliding.  So drop the NCIDS one.
    
    goMySQLCnn.Execute "UPDATE seq_main SET id=LAST_INSERT_ID(id+1)"
    oRs.Open "SELECT LAST_INSERT_ID()", goMySQLCnn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not oRs.EOF Then
        ICHCSII_DAS_GetNextNcid = oRs(0)
    End If
    
    
Func_Exit:
  On Error Resume Next
  oRs.Close
  Set oRs = Nothing
  Exit Function
  
Func_Err:
  '<< SCR #44407;   Developer: Brian Mowbray 01/13/2004 02:07 PM
  Call HandleDBError(Err.Number, Err.Description, Module, goMySQLCnn)
  'This was added to trap for error raised by the HandleDBError if the error is
  'not related to the loss of connectivity.
  If Err.Number <> 0 Then
    Err.Raise Err.Number, Module, Err.Description
  End If
  Resume Func_Exit
End Function

Private Function SeqEnumToString(ByVal eiSeq As eSeq) As String
    Select Case eiSeq
    Case esMAINSEQ:         SeqEnumToString = "SEQ_MAIN" ' "MainSeq"
    Case esAPPT_ID:         SeqEnumToString = "SEQ_APPT_ID"
    Case esAPPT_TRANS_ID:   SeqEnumToString = "SEQ_APPT_TRANS_ID"
    Case esENC_TEMPLATE_ID: SeqEnumToString = "SEQ_ENC_TEMPLATE_ID"
    Case esENCOUNTER_ID:    SeqEnumToString = "SEQ_ENCOUNTER_ID"
    Case esIEN:             SeqEnumToString = "SEQ_IEN"
    Case esPCM_CODE:        SeqEnumToString = "SEQ_PCM_CODE"
    Case esPI_DATA_ID:      SeqEnumToString = "SEQ_PI_DATA_ID"
    Case esUNIT_NUMBER:     SeqEnumToString = "SEQ_UNIT_NUMBER"
    Case esUSER_NUMBER:     SeqEnumToString = "SEQ_USER_NUMBER"
    Case esJDOE_NUMBER:     SeqEnumToString = "SEQ_JDOE_NUMBER"
    Case esTEMP_SSN:        SeqEnumToString = "SEQ_TEMP_SSN"
    Case esDENTAL_ANESTHETICS_SEQ:        SeqEnumToString = "SEQ_DENTAL_ANESTHETICS"
    Case esDENTAL_ENCOUNTER_DXS_SS_SEQ:        SeqEnumToString = "SEQ_DENTAL_ENCOUNTER_DXS_SS"
    Case esDENTAL_ENCOUNTER_DX_SEQ:        SeqEnumToString = "SEQ_DENTAL_ENCOUNTER_DX"
    Case esDENTAL_ENCOUNTER_SEQ:        SeqEnumToString = "SEQ_DENTAL_ENCOUNTER"
    Case esDENTAL_ENCOUNTER_TEETH_SEQ:        SeqEnumToString = "SEQ_DENTAL_ENCOUNTER_TEETH"
    Case esDENTAL_ENCOUNTER_TXS_SEQ:        SeqEnumToString = "SEQ_DENTAL_ENCOUNTER_TXS"
    Case esDENTAL_ENCOUNTER_TXS_SS_SEQ:        SeqEnumToString = "SEQ_DENTAL_ENCOUNTER_TXS_SS"
    Case esDENTAL_ENCOUNTER_TX_SEQ:        SeqEnumToString = "SEQ_DENTAL_ENCOUNTER_TX"
    Case esDENTAL_ENC_SURFACE_SET_SEQ:        SeqEnumToString = "SEQ_DENTAL_ENC_SURFACE_SET"
    Case esDENTAL_IMAGING_RADS_SEQ:        SeqEnumToString = "SEQ_DENTAL_IMAGING_RADS"
    Case esDENTAL_IMAGING_VERSION_GRP_SEQ:        SeqEnumToString = "SEQ_DENTAL_IMAGING_VERSION_GRP"
    Case esDENTAL_IMPRESSIONS_SEQ:        SeqEnumToString = "SEQ_DENTAL_IMPRESSIONS"
    Case esDENTAL_RADIOGRAPHS_SEQ:        SeqEnumToString = "SEQ_DENTAL_RADIOGRAPHS"
    Case esDX_ICD_CODE_DEFINITION_SEQ:        SeqEnumToString = "SEQ_DX_ICD_CODE_DEFINITION"
    Case esGRAFT_ID_SEQ:        SeqEnumToString = "SEQ_GRAFT_ID"
    Case esGROUP_ID_SEQ:        SeqEnumToString = "SEQ_GROUP_ID"
    Case esLAB_ADA_CLV_SEQ:        SeqEnumToString = "SEQ_LAB_ADA_CLV"
    Case esLAB_CASE_SEQ:        SeqEnumToString = "SEQ_LAB_CASE"
    Case esLAB_MATERIALS_SEQ:        SeqEnumToString = "SEQ_LAB_MATERIALS"
    Case esLAB_MATERIAL_ELEMENTS_SEQ:        SeqEnumToString = "SEQ_LAB_MATERIAL_ELEMENTS"
    Case esLAB_NOTE_SEQ:        SeqEnumToString = "SEQ_LAB_NOTE"
    Case esLAB_PART_SPECIFICATIONS_SEQ:        SeqEnumToString = "SEQ_LAB_PART_SPECIFICATIONS"
    Case esLAB_TASK_CATEGORY_DEF_SEQ:        SeqEnumToString = "SEQ_LAB_TASK_CATEGORY_DEF"
    Case esLAB_TASK_DEFINITION_SEQ:        SeqEnumToString = "SEQ_LAB_TASK_DEFINITION"
    Case esLAB_TX_SPECIFICATIONS_SEQ:        SeqEnumToString = "SEQ_LAB_TX_SPECIFICATIONS"
    Case esLAB_TX_TASK_SPECS_SEQ:        SeqEnumToString = "SEQ_LAB_TX_TASK_SPECS"
    Case esNOTE_CODES_DEFINITION_SEQ:        SeqEnumToString = "SEQ_NOTE_CODES_DEFINITION"
    Case esNOTE_CODES_TYPE_DEFINITION_SEQ:        SeqEnumToString = "SEQ_NOTE_CODES_TYPE_DEFINITION"
    Case esNOTE_DXS_LOCATION_SEQ:        SeqEnumToString = "SEQ_NOTE_DXS_LOCATION"
    Case esNOTE_DX_SEQ:        SeqEnumToString = "SEQ_NOTE_DX"
    Case esNOTE_DX_SUB_LOCATION_SEQ:        SeqEnumToString = "SEQ_NOTE_DX_SUB_LOCATION"
    Case esNOTE_TXS_LOCATION_SEQ:        SeqEnumToString = "SEQ_NOTE_TXS_LOCATION"
    Case esNOTE_TXS_SEQ:        SeqEnumToString = "SEQ_NOTE_TXS"
    Case esNOTE_TX_SUB_LOCATION_SEQ:        SeqEnumToString = "SEQ_NOTE_TX_SUB_LOCATION"
    Case esPREVENTIVE_DEBRIDEMENT_SEQ:        SeqEnumToString = "SEQ_PREVENTIVE_DEBRIDEMENT"
    Case esPREVENTIVE_DELIVERY_SEQ:        SeqEnumToString = "SEQ_PREVENTIVE_DELIVERY"
    Case esPREVENT_SCALING_SUPRAGIN_SEQ:        SeqEnumToString = "SEQ_PREVENT_SCALING_SUPRAGIN"
    Case esSURFACE_DEFINITION_SEQ:        SeqEnumToString = "SEQ_SURFACE_DEFINITION"
    Case esTEMPL_ANESTH_SEQ:        SeqEnumToString = "SEQ_TEMPL_ANESTH"
    Case esTEMPL_CARIES_RISK_PROTOCOL_SEQ:        SeqEnumToString = "SEQ_TEMPL_CARIES_RISK_PROTOCOL"
    Case esTEMPL_CARIES_RISK_RECALL_V_SEQ:        SeqEnumToString = "SEQ_TEMPL_CARIES_RISK_RECALL_V"
    Case esTEMPL_EXAM_ENDO_ADD_F_SEQ:        SeqEnumToString = "SEQ_TEMPL_EXAM_ENDO_ADD_F"
    Case esTEMPL_EXAM_ENDO_COMMENTS_SEQ:        SeqEnumToString = "SEQ_TEMPL_EXAM_ENDO_COMMENTS"
    Case esTEMPL_EXAM_ENDO_ENDONTIC_T_SEQ:        SeqEnumToString = "SEQ_TEMPL_EXAM_ENDO_ENDONTIC_T"
    Case esTEMPL_EXAM_ENDO_FINDINGS_SEQ:        SeqEnumToString = "SEQ_TEMPL_EXAM_ENDO_FINDINGS"
    Case esTEMPL_EXAM_ENDO_PAIN_CHARS_SEQ:        SeqEnumToString = "SEQ_TEMPL_EXAM_ENDO_PAIN_CHARS"
    Case esTEMPL_EXAM_ENDO_SUBJECTIVE_SEQ:        SeqEnumToString = "SEQ_TEMPL_EXAM_ENDO_SUBJECTIVE"
    Case esTEMPL_EXAM_PROS_DENTURE_F_SEQ:        SeqEnumToString = "SEQ_TEMPL_EXAM_PROS_DENTURE_F"
    Case esTEMPL_EXAM_PROS_DENTURE_SEQ:        SeqEnumToString = "SEQ_TEMPL_EXAM_PROS_DENTURE"
    Case esTEMPL_EXAM_XML_SEQ:        SeqEnumToString = "SEQ_TEMPL_EXAM_XML"
    Case esTEMPL_GENPROS_IMPRESS_SEQ:        SeqEnumToString = "SEQ_TEMPL_GENPROS_IMPRESS"
    Case esTEMPL_NARRATIVE_SUMMARY_SEQ:        SeqEnumToString = "SEQ_TEMPL_NARRATIVE_SUMMARY"
    Case esTEMPL_PREVENTIVE_SEQ:        SeqEnumToString = "SEQ_TEMPL_PREVENTIVE"
    Case esTEMPL_RADIOLOGY_FINDINGS_SEQ:        SeqEnumToString = "SEQ_TEMPL_RADIOLOGY_FINDINGS"
    Case esTEMPL_TREAT_NSE_CANAL_INFO_SEQ:        SeqEnumToString = "SEQ_TEMPL_TREAT_NSE_CANAL_INFO"
    Case esTEMPL_TREAT_NSE_ENCOUNTER_SEQ:        SeqEnumToString = "SEQ_TEMPL_TREAT_NSE_ENCOUNTER"
    Case esTEMPL_TREAT_NSE_TOOTH_INFO_SEQ:        SeqEnumToString = "SEQ_TEMPL_TREAT_NSE_TOOTH_INFO"
    Case esTEMPL_WORKLOAD_SEQ:        SeqEnumToString = "SEQ_TEMPL_WORKLOAD"
    Case esTOOTH_DEFINITION_SEQ:        SeqEnumToString = "SEQ_TOOTH_DEFINITION"
    Case esTOOTH_ORIENT_DEFINITION_SEQ:        SeqEnumToString = "SEQ_TOOTH_ORIENT_DEFINITION"
    Case esTREATMENT_ANESTHETIC_SEQ:        SeqEnumToString = "SEQ_TREATMENT_ANESTHETIC"
    Case esTX_ADA_CODE_DEFINITION_SEQ:        SeqEnumToString = "SEQ_TX_ADA_CODE_DEFINITION"
    End Select

End Function



Public Function ICHCSII_DAS_ExecuteSQL(SqlStatement As String) As Variant
    If UBound(Split(SqlStatement, " ")) = 1 Or UCase$(Left$(SqlStatement, 1)) = "S" Then
        Set ICHCSII_DAS_ExecuteSQL = ExecuteSQL(SqlStatement)
    Else
        ICHCSII_DAS_ExecuteSQL = ExecuteSQL(SqlStatement)
    End If
End Function


Public Function ExecuteSQL(SqlStatement As String, _
         Optional ByVal UserID As String, _
         Optional ByVal DTS As Date) As Variant
#If DEBUGON Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "ExecuteSQL", "Begin", DR, "SQL", SqlStatement
#End If

    Dim FirstInsPt  As Long
    Dim SecondInsPt As Long
    Dim ValuesStrPt As Long                 '*-- Coleman 040401
    Dim lRecCnt     As Long
    Dim words()     As String
    Dim WhereClause As String
    Dim sTableName As String
    
    On Error GoTo Func_Err

    
'    If msUserNCID = "" Then
'        Dim objShared As New CWShared
'        msUserNCID = objShared.CurrentUserNCID
'        Set objShared = Nothing
'    End If
    
    If UserID = "" Then UserID = CurrentUserNCID

'SCR 21048 - 1/30/02 - Walker - This code doesn't detect a NULL date.
'Consequently, a date is never placed into it for NULLs.
    If Val(DTS) = 0 Then DTS = Now
    If CDbl(DTS) = 0 Then DTS = Now
    
    'SqlStatement = FixParans(SqlStatement)
    
    'Call RedirectEncrypted(SqlStatement, Me) '__ this is a byref call
    
    words = Split(SqlStatement, " ")
        
'<< SCR 33533 encrypt db.  chaned from 'don't audit' to audit.
If UBound(words) > 1 Then
    Select Case UCase$(Left$(SqlStatement, 6))
    Case "DELETE"
        sTableName = words(2)
        '- is as "Delete from tablename where field...
        If AuditedTable(words(2)) Then
            If InStr(LCase$(SqlStatement), " where ") Then
                WhereClause = Mid$(SqlStatement, InStr(LCase$(SqlStatement), " where "))
            Else
                WhereClause = ""
            End If
            goMySQLCnn.Execute "update " & words(2) & " set UpdatedBy = " & SQLQuote(UserID) _
                        & ", UpdatedOn = " & SQLDate(DTS) & WhereClause, lRecCnt
        End If
        goMySQLCnn.Execute SqlStatement, lRecCnt
        ExecuteSQL = lRecCnt
        
    Case "INSERT"    '- may be "Insert into Tablename ( field1....) values (value1...)"
                '- or     "Insert into TableName (field1...) select ....."
        sTableName = words(2)
        If AuditedTable(sTableName) Then
            FirstInsPt = InStr(SqlStatement, "(")
            '*-- Coleman 040401: Handles case where no space before parentheses - "values("
            '*-- If InStr(LCase(SqlStatement), " values ") > 0 Then
            '*--    SecondInsPt = InStr(FirstInsPt + 1, SqlStatement, "(")
            ValuesStrPt = InStr(LCase(SqlStatement), " values")     '*-- Coleman 040401
            If ValuesStrPt > 0 Then                                 '*-- Coleman 040401
                SecondInsPt = InStr(ValuesStrPt, SqlStatement, "(") '*-- Coleman 040401
            Else
                SecondInsPt = InStr(LCase(SqlStatement), " select ") + 7
            End If
       
            SqlStatement = Stuff(SqlStatement, SecondInsPt, SQLQuote(UserID) & ", " & SQLDate(DTS) & ", ")
            SqlStatement = Stuff(SqlStatement, FirstInsPt, "CreatedBy, CreatedOn, ")
        End If
        
        goMySQLCnn.Execute SqlStatement, lRecCnt
        ExecuteSQL = lRecCnt
        
    Case "UPDATE"    '- must be "Update TableName Set Field1 = ..."
    '- get the table name
        sTableName = words(1)
        If AuditedTable(sTableName) Then
            FirstInsPt = InStr(UCase$(SqlStatement), "SET")
            SqlStatement = Stuff(SqlStatement, FirstInsPt + 3, " UpdatedBy = " & SQLQuote(UserID) & ", UpdatedOn = " & SQLDate(DTS) & ", ")
        End If
        goMySQLCnn.Execute SqlStatement, lRecCnt
        ExecuteSQL = lRecCnt
        
    Case "SELECT"
        'SCR 83806
        Dim iFrom As Integer
        Dim iTable As Integer
        Dim iEndTable As Integer
    
        iFrom = InStr(UCase(SqlStatement), "FROM")
        If iFrom > 0 Then
            iEndTable = InStr(iFrom + Len("FROM "), UCase(SqlStatement), " ")
            sTableName = Mid(UCase(SqlStatement), iFrom + Len("FROM "), iEndTable - (iFrom + Len("FROM ")))
        End If
        
        Set ExecuteSQL = goMySQLCnn.Execute(SqlStatement)
        
    End Select
    
  Else  '- ubound(words) > 1
     Set ExecuteSQL = goMySQLCnn.Execute(SqlStatement)
  End If
  
'<< SCR 48680  Need to move TMIP audit code to Current Build
  Call TMIPAudit("C2T_DAS", "MYSQL_DAS.ExecuteSQL", CurrentUserName, sTableName, SqlStatement, Err.Description, goMySQLCnn)

#If DEBUGON Then
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "ExecuteSQL", "End", DR
#End If

  Exit Function

Func_Err:
  '<< SCR #44407;   Developer: Brian Mowbray 01/13/2004 02:08 PM
  Call HandleDBError(Err.Number, Err.Description, Module, goMySQLCnn)
  If Err <> 0 Then
    Err.Raise Err, Module, Err.Description
  End If
  Exit Function
  Resume
End Function

Private Function SQLCase(ByVal NewCase As eCase, ByVal sText As String) As String
    Select Case NewCase
    Case ec_lower
        
        SQLCase = "lower('" & sText & "')"
        
    Case ec_UPPER
        SQLCase = "upper('" & sText & "')"
    
    End Select

End Function



Private Sub Class_Initialize()
        
    Call OpenDatabase
    
    #If DEBUGON Then
      Set DL = GetObject("", "dlog.cdebuglog")
      If Not DL Is Nothing Then DL.debugStart
      Dim DR As DebugRec
      If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Initialize", "", DR
    #End If

End Sub

Private Sub Class_Terminate()
#If DEBUGON Then
  Dim DR As DebugRec
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Class Terminate", "Begin", DR
#End If

    CloseDatabase
    
#If DEBUGON Then
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Class Terminate", "End", DR
#End If
    
End Sub

'<<SCR 33533 encrypt db
Private Sub OpenDatabase()
'- Initiate a connection to the Oracle db

    On Error GoTo Sub_Err

    'Connection for sql statements
    glMySQLCnnCnt = glMySQLCnnCnt + 1
    If goMySQLCnn Is Nothing Then
        Set goMySQLCnn = New ADODB.Connection
        
        goMySQLCnn.ConnectionString = GetConnectionString
        
        goMySQLCnn.ConnectionTimeout = 30
        goMySQLCnn.CursorLocation = adUseClient
    End If
    If goMySQLCnn.State <> 1 Then
        goMySQLCnn.Open
    End If
    
    'TODO: Convert to MySQL connection
    'Use the microsoft provider for stored procedures to work with cursor refs
    If goMySQLSPCnn Is Nothing Then
        Set goMySQLSPCnn = New ADODB.Connection
        
        goMySQLSPCnn.ConnectionString = ReplaceParameter(GetConnectionString, "PROVIDER=", "MSDAORA")
        
        goMySQLSPCnn.ConnectionTimeout = 30
        goMySQLSPCnn.CursorLocation = adUseClient
    End If
    If goMySQLSPCnn.State <> 1 Then
        goMySQLSPCnn.Open
    End If

    If goMySQLCnn.State = 1 Then
        If msDBKey <> "" Then
            goMySQLCnn.Execute "delete from mmi_x"
            goMySQLCnn.Execute "Insert into mmi_x (x0) values ('" & msDBKey & "')"
        End If
    End If
    
    Exit Sub
    
Sub_Err:
  '<< SCR #44407;   Developer: Brian Mowbray 01/13/2004 02:08 PM
  Call HandleDBError(Err.Number, Err.Description, Module, goMySQLCnn)
  'This was added to trap for error raised by the HandleDBError if the error is
  'not related to the loss of connectivity.
  If Err.Number <> 0 Then
    Err.Raise Err.Number, Module, Err.Description
  End If
  Exit Sub
  Resume
End Sub

Private Function GetConnectionString(Optional bShape As Boolean = False) As String
    Dim sServer As String
    Dim sUid As String
    Dim sPWD As String
    Dim oShared As New CWShared
    Dim sDBKey As String
    
    If oShared.GetLocalLogon(sUid, sPWD) Then
        If bShape Then
            GetConnectionString = MySQLShapeConnect
        ElseIf sUid = "secadmin" Then
            GetConnectionString = MYSQL_SEC_Conn
        Else
            GetConnectionString = MYSQL_GEMS_Conn
        End If
        GetConnectionString = ReplaceParameter(GetConnectionString, "SERVER=", oShared.MySQLDataSource)
        GetConnectionString = ReplaceParameter(GetConnectionString, "UID=", sUid)
        GetConnectionString = ReplaceParameter(GetConnectionString, "PWD=", sPWD)
        If sDBKey <> "" Then
            msDBKey = sDBKey
        End If
    End If
    Set oShared = Nothing
    
End Function

'return database vendor specific Trim function
Private Function SQLTrim(inputStr As String) As String

    SQLTrim = "trim(" & inputStr & ")"

End Function


Private Sub CloseDatabase()

    glMySQLCnnCnt = glMySQLCnnCnt - 1
    If glMySQLCnnCnt <= 0 Then
        If goMySQLCnn Is Nothing Then Exit Sub
        If goMySQLCnn.State > 0 Then goMySQLCnn.Close
        Set goMySQLCnn = Nothing
        
        If goMySQLSPCnn Is Nothing Then Exit Sub
        If goMySQLSPCnn.State > 0 Then goMySQLSPCnn.Close
        Set goMySQLSPCnn = Nothing
    End If
    
End Sub

Private Function ICHCSII_DAS_OpenRecordset(SQL As String, _
                            Optional ByVal CursorType As ADODB.CursorTypeEnum = 0&, _
                            Optional ByVal LockType As ADODB.LockTypeEnum = 1&) As ADODB.Recordset

    Dim objRecordset As New ADODB.Recordset
    Dim CmdType As Long
    Dim sTableName As String
    Dim words() As String
    Dim iCounter As Long
    Dim nCounter As Long
    
    On Error GoTo Func_Err
    
    'SF SCR 89066 3/2/06
    SQL = Replace(SQL, vbCrLf, " ")
    
    words = Split(SQL, " ")
    
    '-- only allow select's and tables
    Select Case UCase$(Left$(SQL, 6))
    Case "INSERT", "DELETE", "UPDATE"
        Exit Function
    Case "SELECT", "SHAPE ", "SHAPE{"
        For iCounter = 0 To UBound(words())
            If Trim(UCase(words(iCounter))) = "FROM" Then
                'SF SCR 89066 3/2/06
                nCounter = 1
                Do
                    sTableName = words(iCounter + nCounter)
                    nCounter = nCounter + 1
                Loop Until sTableName <> ""
                Exit For
            End If
        Next
        
        CmdType = adCmdText
    Case Else
        sTableName = SQL
        CmdType = adCmdTable
    End Select
    objRecordset.CursorLocation = adUseClient
'    Dim i As Integer
'    Dim v As Variant
'    For Each v In goMySQLCnn.Properties
'        Debug.Print v.Name, v.Type, v.Value
'    Next v
    objRecordset.Open SQL, goMySQLCnn, CursorType, LockType, CmdType
    
    If Not objRecordset.EOF Then
        objRecordset.MoveFirst
    End If
    
'<< SCR 48680  Need to move TMIP audit code to Current Build
  Call TMIPAudit("C2T_DAS", "MYSQL_DAS.OpenRecordset", CurrentUserName, sTableName, SQL, Err.Description, goMySQLCnn)
    
Func_Exit:
  On Error Resume Next
  Set ICHCSII_DAS_OpenRecordset = objRecordset
  Set objRecordset = Nothing
  Exit Function

Func_Err:
  '<< SCR #44407;   Developer: Brian Mowbray 01/13/2004 02:06 PM
  Call HandleDBError(Err.Number, Err.Description, Module, goMySQLCnn)
  'This was added to trap for error raised by the HandleDBError if the error is
  'not related to the loss of connectivity.
  If Err.Number <> 0 Then
    Err.Raise Err.Number, Module, Err.Description
  End If
  Resume Func_Exit
  
  Exit Function
  Resume
End Function

Private Function ICHCSII_DAS_OpenRecordsetMSP(Parms As Collection) As Collection

    Err.Raise 1001, "MYSQL_DAS", "OpenRecordsetMSP is not supported in GEMS DAS."

End Function

Private Function ICHCSII_DAS_OpenRecordsetSP(Parms As CHCSII_Command.CHCSII_Cmd) As ADODB.Recordset

    Err.Raise 1001, "MYSQL_DAS", "OpenRecordsetSP is not supported in GEMS DAS."

End Function

Public Property Let ICHCSII_DAS_UserNCID(RHS As String)
    
    CurrentUserNCID RHS
    
End Property
 
Public Function ICHCSII_DAS_BeginTrans() As Boolean
    goMySQLCnn.BeginTrans
End Function
Public Function ICHCSII_DAS_CommitTrans() As Boolean
    goMySQLCnn.CommitTrans
End Function
Public Function ICHCSII_DAS_RollbackTrans() As Boolean
    goMySQLCnn.RollbackTrans
End Function

Private Function ISecSqlOps_Connect() As Boolean
    
    Err.Raise 1001, "CDR_DAS", "ISecSqlOps_Connect is no longer supported."
    
End Function
Public Function SecDisConnect() As Boolean
'__ 9.11.02 Shaw: added back to this module so we can run without tmip
'- terminate a "secure" connection to the Oracle db
    
    If Not goMySQLCnn Is Nothing Then
        If goMySQLCnn.State > 0 Then goMySQLCnn.Close
        Set goMySQLCnn = Nothing
    End If

    Call OpenDatabase
    
    SecDisConnect = True
ErrHandler:

End Function

Private Function ICHCSII_DAS_ExecuteSP(Params As CHCSII_Command.CHCSII_Cmd) As Variant
 
    Err.Raise 1001, "MYSQL_DAS", "ExecuteSP is not supported in GEMS DAS."
 
End Function
Private Function ICHCSII_DAS_ExecuteMSP(Parms As Collection) As Collection
 
End Function


