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
Attribute VB_Name = "ITT_DAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'THIS WORKS FOR LOCAL DBS (ORACLE AND SQL SERVER)
'AND IS STILL DEPENDENT ON CWSHARED TO DETERMINE THE DB USED

'-- Facility workaround is not included since this dll will only be used in the case
'   of a single facility.

Implements ICHCSII_DataAccess.ICHCSII_DAS
Implements ISec_SQL.ISecSqlOps
Implements ICHCSII_DataAccess.I_DAS_Direct
Implements ICHCSII_DataAccess.IADOShape
Implements IADO_DAS

'Private m_saRelevantTables() As String
'Private goITTCnn                As ADODB.Connection
Private msUserNCID           As String
Private mNoAudit             As Collection
#If DEBUGON Then
  Public DL                  As cDebugLog
  Private Const Exe          As String = "ITT_Das"
  Private Const Module       As String = "ITT_DAS Cls"
#End If



Private Property Get I_DAS_Direct_ActiveConnection() As ADODB.Connection
    Set I_DAS_Direct_ActiveConnection = goITTCnn
End Property

Private Function I_DAS_Direct_BeginTrans() As Boolean
    goITTCnn.BeginTrans
End Function

Private Function I_DAS_Direct_CommitTrans() As Boolean
    goITTCnn.CommitTrans
End Function

Private Function I_DAS_Direct_ExecuteSQL(ByVal SQL As String, Optional ByVal UserID As String, Optional ByVal RecordDateTime As Date) As Variant
    I_DAS_Direct_ExecuteSQL = ExecuteSQL(SQL, UserID, RecordDateTime)
End Function

Private Function I_DAS_Direct_RollbackTrans() As Boolean
    goITTCnn.RollbackTrans
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

    oCommand.ActiveConnection = goITTCnn
    
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

    Err.Raise 1001, "IADO_DAS", "FacilityWorkaround is not supported in ITT DAS."

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

    oCommand.ActiveConnection = goITTCnn
    
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
    Dim sNSN As String
    Dim iDS As Integer
    Dim iSC As Integer
    Dim sShapeConnect As String
    Dim SearchStr As String
    Dim sDSN As String
    Dim sInitCat As String
    Dim sUid As String
    Dim sPWD As String
    Dim oShared As New CWShared
    
    sNSN = oShared.NetServiceName
    sDSN = oShared.DSN
    sInitCat = oShared.InitialCatalog
    
    Const pADOShapeConnect   As String = _
        "Provider=MSDataShape.1;Persist Security Info=False;" & _
        "Data Source=TAN_TEST;Initial Catalog=cdrsqlbs;" & _
        "User ID=sa;Password=;" & _
        "Data Provider=SQLOLEDB.1" '-- ITT
        
    sShapeConnect = pADOShapeConnect
    
    If sDSN <> "" Then
        sShapeConnect = ReplaceParameter(sShapeConnect, "DATA SOURCE=", sDSN)
    End If
    
    If sInitCat <> "" Then
        sShapeConnect = ReplaceParameter(sShapeConnect, "INITIAL CATALOG=", sInitCat)
    End If

'-- get the local uid and pwd if available
    If oShared.GetLocalLogon(sUid, sPWD) Then
        '__ replace the user id and password
         sShapeConnect = ReplaceParameter(sShapeConnect, "USER ID=", sUid)
         sShapeConnect = ReplaceParameter(sShapeConnect, "PASSWORD=", sPWD)
    End If

    Set oShared = Nothing
    
    IADOShape_ShapeConnect = sShapeConnect
    
End Function

Private Property Get ICHCSII_DAS_ActiveConnection() As ADODB.Connection
    
    Set ICHCSII_DAS_ActiveConnection = goITTCnn

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

    Err.Raise 1001, "ITT_DAS", "FacilityWorkaround is not supported in ITT DAS."
    
End Property


Private Function ICHCSII_DAS_GetNextID(Optional ByVal eSeqName As ICHCSII_DataAccess.eSeq = esMAINSEQ) As Variant
'-- pulls the next value from the MainSeq sequence

On Error GoTo ErrHandler
    
Dim oRs As New ADODB.Recordset
Dim sSeqName As String
Dim lSeqNum As Long
    
    sSeqName = SeqEnumToString(eSeqName)
    
    If eSeqName = esMAINSEQ Then sSeqName = "PROVIDER_NCID"
    
    oRs.Open "Select Value from Sequences where seq_type = '" & sSeqName & "'", goITTCnn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not oRs.EOF Then
        lSeqNum = Val(oRs(0))
    Else
        MsgBox "ITT_DAS.GetNextID: Could not find a 'sequence' called " & sSeqName & " in the Sequences table."
        GoTo ErrHandler
    End If
    
    goITTCnn.Execute "update Sequences set value = " & lSeqNum + 1 & " where seq_type = '" & sSeqName & "'"
    ICHCSII_DAS_GetNextID = lSeqNum
    
ErrHandler:
    If oRs.State > 0 Then oRs.Close
    Set oRs = Nothing
    
End Function

Private Function ICHCSII_DAS_GetNextNcid() As Variant

'-- pulls the next value from the NCID sequence
    Dim oRs As ADODB.Recordset
    Set oRs = New ADODB.Recordset
    Dim lSeqNum As Long
    
    oRs.Open "Select Value from Sequences where seq_type = 'PROVIDER_NCID'", goITTCnn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not oRs.EOF Then
        lSeqNum = Val(oRs(0))
    End If
    goITTCnn.Execute "update Sequences set value = " & lSeqNum + 1 & " where seq_type = 'PROVIDER_NCID'"
    ICHCSII_DAS_GetNextNcid = lSeqNum
    
    oRs.Close
    Set oRs = Nothing
    
End Function

Private Function SeqEnumToString(ByVal eiSeq As eSeq) As String
    Select Case eiSeq
    Case esMAINSEQ:         SeqEnumToString = "MainSeq"
    Case esAPPT_ID:         SeqEnumToString = "APPT_ID"
    Case esAPPT_TRANS_ID:   SeqEnumToString = "APPT_TRANS_ID"
    Case esENC_TEMPLATE_ID: SeqEnumToString = "ENC_TEMPLATE_ID"
    Case esENCOUNTER_ID:    SeqEnumToString = "ENCOUNTER_ID"
    Case esIEN:             SeqEnumToString = "IEN"
    Case esPCM_CODE:        SeqEnumToString = "PCM_CODE"
    Case esPI_DATA_ID:      SeqEnumToString = "PI_DATA_ID"
    Case esUNIT_NUMBER:     SeqEnumToString = "UNIT_NUMBER"
    Case esUSER_NUMBER:     SeqEnumToString = "USER_NUMBER"
    Case esJDOE_NUMBER:     SeqEnumToString = "JDOE_NUMBER"
    Case esTEMP_SSN:        SeqEnumToString = "TEMP_SSN"
    Case esDENTAL_ANESTHETICS_SEQ:        SeqEnumToString = "DENTAL_ANESTHETICS_SEQ"
    Case esDENTAL_ENCOUNTER_DXS_SS_SEQ:        SeqEnumToString = "DENTAL_ENCOUNTER_DXS_SS_SEQ"
    Case esDENTAL_ENCOUNTER_DX_SEQ:        SeqEnumToString = "DENTAL_ENCOUNTER_DX_SEQ"
    Case esDENTAL_ENCOUNTER_SEQ:        SeqEnumToString = "DENTAL_ENCOUNTER_SEQ"
    Case esDENTAL_ENCOUNTER_TEETH_SEQ:        SeqEnumToString = "DENTAL_ENCOUNTER_TEETH_SEQ"
    Case esDENTAL_ENCOUNTER_TEETH_SEQ:        SeqEnumToString = "DENTAL_ENCOUNTER_TEETH_SEQ"
    Case esDENTAL_ENCOUNTER_TXS_SEQ:        SeqEnumToString = "DENTAL_ENCOUNTER_TXS_SEQ"
    Case esDENTAL_ENCOUNTER_TXS_SS_SEQ:        SeqEnumToString = "DENTAL_ENCOUNTER_TXS_SS_SEQ"
    Case esDENTAL_ENCOUNTER_TX_SEQ:        SeqEnumToString = "DENTAL_ENCOUNTER_TX_SEQ"
    Case esDENTAL_ENC_SURFACE_SET_SEQ:        SeqEnumToString = "DENTAL_ENC_SURFACE_SET_SEQ"
    Case esDENTAL_IMAGING_RADS_SEQ:        SeqEnumToString = "DENTAL_IMAGING_RADS_SEQ"
    Case esDENTAL_IMAGING_VERSION_GRP_SEQ:        SeqEnumToString = "DENTAL_IMAGING_VERSION_GRP_SEQ"
    Case esDENTAL_IMPRESSIONS_SEQ:        SeqEnumToString = "DENTAL_IMPRESSIONS_SEQ"
    Case esDENTAL_RADIOGRAPHS_SEQ:        SeqEnumToString = "DENTAL_RADIOGRAPHS_SEQ"
    Case esDX_ICD_CODE_DEFINITION_SEQ:        SeqEnumToString = "DX_ICD_CODE_DEFINITION_SEQ"
    Case esGRAFT_ID_SEQ:        SeqEnumToString = "GRAFT_ID_SEQ"
    Case esGROUP_ID_SEQ:        SeqEnumToString = "GROUP_ID_SEQ"
    Case esLAB_ADA_CLV_SEQ:        SeqEnumToString = "LAB_ADA_CLV_SEQ"
    Case esLAB_CASE_SEQ:        SeqEnumToString = "LAB_CASE_SEQ"
    Case esLAB_MATERIALS_SEQ:        SeqEnumToString = "LAB_MATERIALS_SEQ"
    Case esLAB_MATERIAL_ELEMENTS_SEQ:        SeqEnumToString = "LAB_MATERIAL_ELEMENTS_SEQ"
    Case esLAB_NOTE_SEQ:        SeqEnumToString = "LAB_NOTE_SEQ"
    Case esLAB_PART_SPECIFICATIONS_SEQ:        SeqEnumToString = "LAB_PART_SPECIFICATIONS_SEQ"
    Case esLAB_TASK_CATEGORY_DEF_SEQ:        SeqEnumToString = "LAB_TASK_CATEGORY_DEF_SEQ"
    Case esLAB_TASK_DEFINITION_SEQ:        SeqEnumToString = "LAB_TASK_DEFINITION_SEQ"
    Case esLAB_TX_SPECIFICATIONS_SEQ:        SeqEnumToString = "LAB_TX_SPECIFICATIONS_SEQ"
    Case esLAB_TX_TASK_SPECS_SEQ:        SeqEnumToString = "LAB_TX_TASK_SPECS_SEQ"
    Case esNOTE_CODES_DEFINITION_SEQ:        SeqEnumToString = "NOTE_CODES_DEFINITION_SEQ"
    Case esNOTE_CODES_TYPE_DEFINITION_SEQ:        SeqEnumToString = "NOTE_CODES_TYPE_DEFINITION_SEQ"
    Case esNOTE_DXS_LOCATION_SEQ:        SeqEnumToString = "NOTE_DXS_LOCATION_SEQ"
    Case esNOTE_DX_SEQ:        SeqEnumToString = "NOTE_DX_SEQ"
    Case esNOTE_DX_SUB_LOCATION_SEQ:        SeqEnumToString = "NOTE_DX_SUB_LOCATION_SEQ"
    Case esNOTE_TXS_LOCATION_SEQ:        SeqEnumToString = "NOTE_TXS_LOCATION_SEQ"
    Case esNOTE_TXS_SEQ:        SeqEnumToString = "NOTE_TXS_SEQ"
    Case esNOTE_TX_SUB_LOCATION_SEQ:        SeqEnumToString = "NOTE_TX_SUB_LOCATION_SEQ"
    Case esPREVENTIVE_DEBRIDEMENT_SEQ:        SeqEnumToString = "PREVENTIVE_DEBRIDEMENT_SEQ"
    Case esPREVENTIVE_DELIVERY_SEQ:        SeqEnumToString = "PREVENTIVE_DELIVERY_SEQ"
    Case esPREVENT_SCALING_SUPRAGIN_SEQ:        SeqEnumToString = "PREVENT_SCALING_SUPRAGIN_SEQ"
    Case esSURFACE_DEFINITION_SEQ:        SeqEnumToString = "SURFACE_DEFINITION_SEQ"
    Case esTEMPL_ANESTH_SEQ:        SeqEnumToString = "TEMPL_ANESTH_SEQ"
    Case esTEMPL_CARIES_RISK_PROTOCOL_SEQ:        SeqEnumToString = "TEMPL_CARIES_RISK_PROTOCOL_SEQ"
    Case esTEMPL_CARIES_RISK_RECALL_V_SEQ:        SeqEnumToString = "TEMPL_CARIES_RISK_RECALL_V_SEQ"
    Case esTEMPL_EXAM_ENDO_ADD_F_SEQ:        SeqEnumToString = "TEMPL_EXAM_ENDO_ADD_F_SEQ"
    Case esTEMPL_EXAM_ENDO_COMMENTS_SEQ:        SeqEnumToString = "TEMPL_EXAM_ENDO_COMMENTS_SEQ"
    Case esTEMPL_EXAM_ENDO_ENDONTIC_T_SEQ:        SeqEnumToString = "TEMPL_EXAM_ENDO_ENDONTIC_T_SEQ"
    Case esTEMPL_EXAM_ENDO_FINDINGS_SEQ:        SeqEnumToString = "TEMPL_EXAM_ENDO_FINDINGS_SEQ"
    Case esTEMPL_EXAM_ENDO_PAIN_CHARS_SEQ:        SeqEnumToString = "TEMPL_EXAM_ENDO_PAIN_CHARS_SEQ"
    Case esTEMPL_EXAM_ENDO_SUBJECTIVE_SEQ:        SeqEnumToString = "TEMPL_EXAM_ENDO_SUBJECTIVE_SEQ"
    Case esTEMPL_EXAM_PROS_DENTURE_F_SEQ:        SeqEnumToString = "TEMPL_EXAM_PROS_DENTURE_F_SEQ"
    Case esTEMPL_EXAM_PROS_DENTURE_SEQ:        SeqEnumToString = "TEMPL_EXAM_PROS_DENTURE_SEQ"
    Case esTEMPL_EXAM_XML_SEQ:        SeqEnumToString = "TEMPL_EXAM_XML_SEQ"
    Case esTEMPL_GENPROS_IMPRESS_SEQ:        SeqEnumToString = "TEMPL_GENPROS_IMPRESS_SEQ"
    Case esTEMPL_NARRATIVE_SUMMARY_SEQ:        SeqEnumToString = "TEMPL_NARRATIVE_SUMMARY_SEQ"
    Case esTEMPL_PREVENTIVE_SEQ:        SeqEnumToString = "TEMPL_PREVENTIVE_SEQ"
    Case esTEMPL_RADIOLOGY_FINDINGS_SEQ:        SeqEnumToString = "TEMPL_RADIOLOGY_FINDINGS_SEQ"
    Case esTEMPL_TREAT_NSE_CANAL_INFO_SEQ:        SeqEnumToString = "TEMPL_TREAT_NSE_CANAL_INFO_SEQ"
    Case esTEMPL_TREAT_NSE_ENCOUNTER_SEQ:        SeqEnumToString = "TEMPL_TREAT_NSE_ENCOUNTER_SEQ"
    Case esTEMPL_TREAT_NSE_TOOTH_INFO_SEQ:        SeqEnumToString = "TEMPL_TREAT_NSE_TOOTH_INFO_SEQ"
    Case esTEMPL_WORKLOAD_SEQ:        SeqEnumToString = "TEMPL_WORKLOAD_SEQ"
    Case esTOOTH_DEFINITION_SEQ:        SeqEnumToString = "TOOTH_DEFINITION_SEQ"
    Case esTOOTH_ORIENT_DEFINITION_SEQ:        SeqEnumToString = "TOOTH_ORIENT_DEFINITION_SEQ"
    Case esTREATMENT_ANESTHETIC_SEQ:        SeqEnumToString = "TREATMENT_ANESTHETIC_SEQ"
    Case esTX_ADA_CODE_DEFINITION_SEQ:        SeqEnumToString = "TX_ADA_CODE_DEFINITION_SEQ"
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
    
    If msUserNCID = "" Then
        Dim objshared As New CWShared
        msUserNCID = objshared.CurrentUserNCID
        Set objshared = Nothing
    End If
    If UserID = "" Then UserID = msUserNCID
'SCR 21048 - 1/30/02 - Walker - This code doesn't detect a NULL date.
'Consequently, a date is never placed into it for NULLs.
    If Val(DTS) = 0 Then DTS = Now
    If CDbl(DTS) = 0 Then DTS = Now
    
    SqlStatement = FixParans(SqlStatement)
    
    words = Split(SqlStatement, " ")
        
 If UBound(words) > 1 Then
    
    Select Case UCase$(Left$(SqlStatement, 1))
    
    Case "D"
        '- is as "Delete from tablename where field...
        If Not DoNoAuditing(words(2)) Then
            If InStr(LCase$(SqlStatement), " where ") Then
                WhereClause = Mid$(SqlStatement, InStr(LCase$(SqlStatement), " where "))
            Else
                WhereClause = ""
            End If
            goITTCnn.Execute "update " & words(2) & " set UpdatedBy = " & SQLQuote(UserID) _
                        & ", UpdatedOn = " & SQLDate(DTS) & WhereClause, lRecCnt
        End If
        goITTCnn.Execute SqlStatement, lRecCnt
        ExecuteSQL = lRecCnt
        
    Case "I"    '- may be "Insert into Tablename ( field1....) values (value1...)"
                '- or     "Insert into TableName (field1...) select ....."
        If Not DoNoAuditing(words(2)) Then
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
        
        goITTCnn.Execute SqlStatement, lRecCnt
        ExecuteSQL = lRecCnt
        
    Case "U"    '- must be "Update TableName Set Field1 = ..."
    '- get the table name
        If Not DoNoAuditing(words(1)) Then
            FirstInsPt = InStr(UCase$(SqlStatement), "SET")
            SqlStatement = Stuff(SqlStatement, FirstInsPt + 3, " UpdatedBy = " & SQLQuote(UserID) & ", UpdatedOn = " & SQLDate(DTS) & ", ")
        End If
        goITTCnn.Execute SqlStatement, lRecCnt
        ExecuteSQL = lRecCnt
        
    Case "S"
       
        Set ExecuteSQL = goITTCnn.Execute(SqlStatement)
        
    End Select
    
  Else  '- ubound(words) > 1
     Set ExecuteSQL = goITTCnn.Execute(SqlStatement)
  End If
#If DEBUGON Then
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "ExecuteSQL", "End", DR
#End If
    
End Function

Private Sub Class_Initialize()
    'LoadRelevantWorkaroundTables
    
    Dim objshared As New CWShared
    msUserNCID = objshared.CurrentUserNCID

    Call OpenDatabase
    
    Set objshared = Nothing

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

Private Sub OpenDatabase()
'- Initiate a connection to the Oracle db
    glITTCnnCnt = glITTCnnCnt + 1
    If goITTCnn Is Nothing Then
        Set goITTCnn = New ADODB.Connection
        
        goITTCnn.ConnectionString = GetConnectionString
        
        goITTCnn.ConnectionTimeout = 30
        goITTCnn.CursorLocation = adUseClient
    End If
    If goITTCnn.State <> 1 Then
         goITTCnn.Open
    End If
    
    Dim i As Integer
'    For i = 0 To goITTCnn.Properties.Count - 1
'    Debug.Print goITTCnn.Properties(i).Name
'    Next i
    
End Sub

Private Function SQLDate(ByVal dDate As Date) As String
    
    'SCR 143 training Tools -  WR
    'SQL server style 13 is dd mon yyyy hh:mm:ss:mmm(24h)
    SQLDate = "convert(datetime,'" & Format(dDate, "dd/mmm/yyyy hh:nn:ss") & "', 13)"

End Function

Private Function SQLCase(ByVal NewCase As eCase, ByVal sText As String) As String
    Select Case NewCase
    Case ec_lower
        'Access is lcase()
        'Sql server is lower()
        SQLCase = "lower('" & sText & "')"
        
    Case ec_UPPER
        'Access ucase()
        'Sql Server upper()
        SQLCase = "upper('" & sText & "')"
    
    End Select

End Function



Private Sub CloseDatabase()

    glITTCnnCnt = glITTCnnCnt - 1
    If glITTCnnCnt <= 0 Then
        If goITTCnn Is Nothing Then Exit Sub
        If goITTCnn.State > 0 Then goITTCnn.Close
        Set goITTCnn = Nothing
    End If
    
End Sub

'return database vendor specific Trim function
Private Function SQLTrim(inputStr As String) As String

    SQLTrim = "Ltrim(Rtrim(" & inputStr & "))"

End Function






Private Function ICHCSII_DAS_OpenRecordset(SQL As String, _
                            Optional ByVal CursorType As ADODB.CursorTypeEnum = 0&, _
                            Optional ByVal LockType As ADODB.LockTypeEnum = 1&) As ADODB.Recordset

    Dim objRecordset As New ADODB.Recordset
    Dim CmdType As Long
    
    '-- only allow select's and tables
    Select Case UCase$(Left$(SQL, 6))
    Case "INSERT", "DELETE", "UPDATE"
        Exit Function
    Case "SELECT", "SHAPE ", "SHAPE{"
        CmdType = adCmdText
    Case Else
        CmdType = adCmdTable
    End Select
    objRecordset.CursorLocation = adUseClient
'    Dim i As Integer
'    Dim v As Variant
'    For Each v In goITTCnn.Properties
'        Debug.Print v.Name, v.Type, v.Value
'    Next v
    objRecordset.Open SQL, goITTCnn, CursorType, LockType, CmdType
    
    If Not objRecordset.EOF Then
        objRecordset.MoveFirst
    End If
    
    Set ICHCSII_DAS_OpenRecordset = objRecordset
    Set objRecordset = Nothing

End Function

Private Function ICHCSII_DAS_OpenRecordsetMSP(Parms As Collection) As Collection

    Err.Raise 1001, "ITT_DAS", "OpenRecordsetMSP is not supported in ITT DAS."

End Function

Private Function ICHCSII_DAS_OpenRecordsetSP(Parms As CHCSII_Command.CHCSII_Cmd) As ADODB.Recordset

    Err.Raise 1001, "ITT_DAS", "OpenRecordsetSP is not supported in ITT DAS."

End Function

Public Property Let ICHCSII_DAS_UserNCID(RHS As String)
    msUserNCID = RHS
End Property
 
Public Function ICHCSII_DAS_BeginTrans() As Boolean
    goITTCnn.BeginTrans
End Function
Public Function ICHCSII_DAS_CommitTrans() As Boolean
    goITTCnn.CommitTrans
End Function
Public Function ICHCSII_DAS_RollbackTrans() As Boolean
    goITTCnn.RollbackTrans
End Function

Private Function DoNoAuditing(sTable As String) As Boolean
    Dim sMember As Variant
    If gbITTSecConn Then
        DoNoAuditing = True
        Exit Function
    End If
    
    If mNoAudit Is Nothing Then LoadNoAudit
    
    For Each sMember In mNoAudit
        If LCase(sMember) = LCase(sTable) Then
            DoNoAuditing = True
            Exit For
        End If
    Next sMember
End Function

Private Sub LoadNoAudit()
    
    If gbITTSecConn Then Exit Sub
    Set mNoAudit = New Collection
    Dim oRs As ADODB.Recordset
    Dim oDAS As ICHCSII_DAS
    Set oDAS = New ITT_DAS
    
    Set oRs = oDAS.OpenRecordset("DoNotAudit")
    Do While Not oRs.EOF
        mNoAudit.Add "" & oRs(0)
        oRs.MoveNext
    Loop
    oRs.Close
    Set oRs = Nothing
    Set oDAS = Nothing
    
End Sub

Private Function ISecSqlOps_Connect() As Boolean

    Err.Raise 1001, "CDR_DAS", "ISecSqlOps_Connect is not supported in ITT_DAS."
    
    Exit Function
    
'-- functionality is moved to TMIP_Security.dll
'__ 9.11.02 Shaw: added back to this module so we can run without tmip
'- Initiate a "secure" connection to the Oracle db
    

End Function
Public Function SecDisConnect() As Boolean

    Err.Raise 1001, "CDR_DAS", "SecDisConnect is not supported in ITT_DAS."
    
    Exit Function


'May not be applicable for ITT (carried over from GEMS_DAS)

'__ 9.11.02 Shaw: added back to this module so we can run without tmip
    
    If Not goITTCnn Is Nothing Then
        If goITTCnn.State > 0 Then goITTCnn.Close
        Set goITTCnn = Nothing
    End If

    Call OpenDatabase
    
    SecDisConnect = True
ErrHandler:

End Function

Private Function GetConnectionString() As String

    'SF 2-2-04 Use GetSQLServerConnectionString when in ITT mode.
    GetConnectionString = SQLSVR_Conn
    
    GetConnectionString = GetSQLServerConnectionString(GetConnectionString)
   
End Function
Private Function ICHCSII_DAS_ExecuteSP(Params As CHCSII_Command.CHCSII_Cmd) As Variant
 
    Err.Raise 1001, "ITT_DAS", "ExecuteSP is not supported in ITT DAS."
 
End Function
Private Function ICHCSII_DAS_ExecuteMSP(Parms As Collection) As Collection
 
End Function




