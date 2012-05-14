VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICD_DAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'SCR 48849
'3-08-04 SF Added support for ICD

'-- Facility workaround is not included since this dll will only be used in the case
'   of a single facility.

Implements ICHCSII_DataAccess.ICHCSII_DAS
Implements ISec_SQL.ISecSqlOps
Implements ICHCSII_DataAccess.I_DAS_Direct
Implements ICHCSII_DataAccess.IADOShape
Implements IADO_DAS

Private msDBKey As String

'3-24-04 These have been replaced by functions called CurrentUserNCID and CurrentUserName
'Private msUserNCID           As String
'Private msUserName           As String

Private Const Module       As String = "ICD_DAS Class"

'Vars for ICD DAS
Public moICDSPCnn                 As ADODB.Connection
Public moICDCnn                   As ADODB.Connection
Public mlICDCnnCnt                As Long        '__ number of connections


#If DEBUGON Then
  Public DL                  As cDebugLog
  Private Const Exe          As String = "ICD_Das"
#End If

Private Property Get I_DAS_Direct_ActiveConnection() As ADODB.Connection
    Set I_DAS_Direct_ActiveConnection = moICDCnn
End Property

Private Function I_DAS_Direct_BeginTrans() As Boolean
    moICDCnn.BeginTrans
End Function

Private Function I_DAS_Direct_CommitTrans() As Boolean
    moICDCnn.CommitTrans
End Function

Private Function I_DAS_Direct_ExecuteSQL(ByVal SQL As String, Optional ByVal UserID As String, Optional ByVal RecordDateTime As Date) As Variant
    I_DAS_Direct_ExecuteSQL = ExecuteSQL(SQL, UserID, RecordDateTime)
End Function

Private Function I_DAS_Direct_RollbackTrans() As Boolean
    moICDCnn.RollbackTrans
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


Private Function IADO_DAS_ExecuteSP(oCommand As Command) As Long

On Error GoTo ErrHnd

'3-09-04 SF
'This function takes the ADO command object and executes it

    oCommand.ActiveConnection = moICDSPCnn
    
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


Private Function IADO_DAS_OpenRecordsetSP(oCommand As Command) As ADODB.Recordset

On Error GoTo ErrHnd

'3-09-04 SF
'This function takes the ADO command object and executes it

    oCommand.ActiveConnection = moICDSPCnn
    
    ' Modified 12/13/2004 - Tom Winn SCR 62575 this property is required for 9i stored procedures returning a ref cursor.
    oCommand.Properties("PLSQLRSet") = True
    
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
    Dim sShapeConnect As String
    Dim sDSN As String
    Dim sInitCat As String
    Dim sUid As String
    Dim sPWD As String
    Dim oShared As CWShared
    
    On Error GoTo ErrHnd
    
    Set oShared = New CWShared
    sNSN = oShared.NetServiceName
    sDSN = oShared.DSN
    sInitCat = oShared.InitialCatalog
    
    Const pADOShapeConnect   As String = _
        "Provider=MSDataShape.1;Persist Security Info=False;" & _
        "Data Source=TAN_TEST;Initial Catalog=cdrsqlbs;" & _
        "User ID=sa;Password=;" & _
        "Data Provider=SQLOLEDB.1" '-- ITT
        
    If Not oShared.UseSQLServer Then
        
        If sNSN = "" Then
            sShapeConnect = ORAShapeConnect
        Else
            sShapeConnect = ReplaceParameter(ORAShapeConnect, "DATA SOURCE=", sNSN)
        End If
    
    Else        '- yes, use sql server db
        sShapeConnect = pADOShapeConnect
        
        If sDSN <> "" Then
            sShapeConnect = ReplaceParameter(sShapeConnect, "DATA SOURCE=", sDSN)
        End If
        
        If sInitCat <> "" Then
            sShapeConnect = ReplaceParameter(sShapeConnect, "INITIAL CATALOG=", sInitCat)
        End If
    
    End If

    '-- get the local uid and pwd if available
    If oShared.GetLocalLogon(sUid, sPWD) Then
        '__ replace the user id and password
        sShapeConnect = ReplaceParameter(sShapeConnect, "USER ID=", sUid)
        sShapeConnect = ReplaceParameter(sShapeConnect, "PASSWORD=", sPWD)
    End If

    Set oShared = Nothing
    
    IADOShape_ShapeConnect = sShapeConnect
    
Exit Function

ErrHnd:
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

Private Property Get ICHCSII_DAS_ActiveConnection() As ADODB.Connection
    
    Set ICHCSII_DAS_ActiveConnection = moICDCnn

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
Dim objRecordset As ADODB.Recordset

On Error GoTo ErrHnd

iFileNum = FreeFile

Set objRecordset = New ADODB.Recordset

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

Exit Function

ErrHnd:
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

Private Property Let ICHCSII_DAS_FacilityWorkaround(ByVal RHS As Boolean)

    Err.Raise 1001, "ICD_DAS", "FacilityWorkaround is not supported in ICD DAS."
    
End Property


Private Function ICHCSII_DAS_GetNextID(Optional ByVal eSeqName As ICHCSII_DataAccess.eSeq = esMAINSEQ) As Variant
'-- pulls the next value from the MainSeq sequence

On Error GoTo ErrHandler
    
Dim oRs As ADODB.Recordset
Dim sSeqName As String
    
    Set oRs = New ADODB.Recordset
    sSeqName = SeqEnumToString(eSeqName)
    oRs.Open "Select " & sSeqName & ".nextval from dual", moICDCnn, adOpenForwardOnly, adLockOptimistic, adCmdText

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
  Call HandleDBError(Err.Number, Err.Description, Module, moICDCnn)
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
    
    On Error GoTo Func_Err
    
'<< SRC 29784     the NCIDs and MainSeq sequencers were colliding.  So drop the NCIDS one.
    oRs.Open "Select MainSeq.nextval from dual", moICDCnn, adOpenForwardOnly, adLockOptimistic, adCmdText
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
  Call HandleDBError(Err.Number, Err.Description, Module, moICDCnn)
  'This was added to trap for error raised by the HandleDBError if the error is
  'not related to the loss of connectivity.
  If Err.Number <> 0 Then
    Err.Raise Err.Number, Module, Err.Description
  End If
  Resume Func_Exit
End Function

Private Function SeqEnumToString(ByVal eiSeq As eSeq) As String
    
On Error GoTo ErrHnd

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
    Case esDENTAL_IMAGING_RAD_ORDERS_SEQ:        SeqEnumToString = "DENTAL_IMAGING_RAD_ORDERS_SEQ"
    Case esDENTAL_IMG_RAD_ORDER_ACC_SEQ:        SeqEnumToString = "DENTAL_IMG_RAD_ORDER_ACC_SEQ"
    
    End Select

Exit Function

ErrHnd:
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function



Public Function ICHCSII_DAS_ExecuteSQL(SqlStatement As String) As Variant
    
On Error GoTo ErrHnd

    If UBound(Split(SqlStatement, " ")) = 1 Or UCase$(Left$(SqlStatement, 1)) = "S" Then
        Set ICHCSII_DAS_ExecuteSQL = ExecuteSQL(SqlStatement)
    Else
        ICHCSII_DAS_ExecuteSQL = ExecuteSQL(SqlStatement)
    End If

Exit Function

ErrHnd:
    Err.Raise Err.Number, Err.Source, Err.Description
    
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
    
    On Error GoTo Func_Err

    
'    If msUserNCID = "" Then
'        Dim objShared As CWShared
'        Set objShared = New CWShared
'        msUserNCID = objShared.CurrentUserNCID
'        Set objShared = Nothing
'    End If

    If UserID = "" Then UserID = CurrentUserNCID
    
'SCR 21048 - 1/30/02 - Walker - This code doesn't detect a NULL date.
'Consequently, a date is never placed into it for NULLs.
    If Val(DTS) = 0 Then DTS = Now
    If CDbl(DTS) = 0 Then DTS = Now
    
    SqlStatement = FixParans(SqlStatement)
    
    Call RedirectEncrypted(SqlStatement, Me)  '__ this is a byref call
    
    words = Split(SqlStatement, " ")
        
'<< SCR 33533 encrypt db.  chaned from 'don't audit' to audit.
If UBound(words) > 1 Then
    Select Case UCase$(Left$(SqlStatement, 6))
    Case "DELETE"
        '- is as "Delete from tablename where field...
        
'        SF - 4-21-04 Per Jason Ray - Since auditing doesn't work now, no reason to make it work because we don't have any requirements for it.
'        If AuditedTable(words(2)) Then
'            If InStr(LCase$(SqlStatement), " where ") Then
'                WhereClause = Mid$(SqlStatement, InStr(LCase$(SqlStatement), " where "))
'            Else
'                WhereClause = ""
'            End If
'            moICDCnn.Execute "update " & words(2) & " set UpdatedBy = " & SQLQuote(UserID) _
'                        & ", UpdatedOn = " & OracleDate(DTS) & WhereClause, lRecCnt
'        End If
        moICDCnn.Execute SqlStatement, lRecCnt
        ExecuteSQL = lRecCnt
        
    Case "INSERT"    '- may be "Insert into Tablename ( field1....) values (value1...)"
                '- or     "Insert into TableName (field1...) select ....."
'        SF - 4-21-04 Per Jason Ray - Since auditing doesn't work now, no reason to make it work because we don't have any requirements for it.
'        If AuditedTable(words(2)) Then
'            FirstInsPt = InStr(SqlStatement, "(")
'            '*-- Coleman 040401: Handles case where no space before parentheses - "values("
'            '*-- If InStr(LCase(SqlStatement), " values ") > 0 Then
'            '*--    SecondInsPt = InStr(FirstInsPt + 1, SqlStatement, "(")
'            ValuesStrPt = InStr(LCase(SqlStatement), " values")     '*-- Coleman 040401
'            If ValuesStrPt > 0 Then                                 '*-- Coleman 040401
'                SecondInsPt = InStr(ValuesStrPt, SqlStatement, "(") '*-- Coleman 040401
'            Else
'                SecondInsPt = InStr(LCase(SqlStatement), " select ") + 7
'            End If
'
'            SqlStatement = Stuff(SqlStatement, SecondInsPt, SQLQuote(UserID) & ", " & OracleDate(DTS) & ", ")
'            SqlStatement = Stuff(SqlStatement, FirstInsPt, "CreatedBy, CreatedOn, ")
'        End If
        
        moICDCnn.Execute SqlStatement, lRecCnt
        ExecuteSQL = lRecCnt
        
    Case "UPDATE"    '- must be "Update TableName Set Field1 = ..."
    '- get the table name
'        SF - 4-21-04 Per Jason Ray - Since auditing doesn't work now, no reason to make it work because we don't have any requirements for it.
'        If AuditedTable(words(1)) Then
'            FirstInsPt = InStr(UCase$(SqlStatement), "SET")
'            SqlStatement = Stuff(SqlStatement, FirstInsPt + 3, " UpdatedBy = " & SQLQuote(UserID) & ", UpdatedOn = " & OracleDate(DTS) & ", ")
'        End If
        moICDCnn.Execute SqlStatement, lRecCnt
        ExecuteSQL = lRecCnt
        
    Case "SELECT"
       
        Set ExecuteSQL = moICDCnn.Execute(SqlStatement)
        
    End Select
    
  Else  '- ubound(words) > 1
     Set ExecuteSQL = moICDCnn.Execute(SqlStatement)
  End If
  
'<< SCR 48680  Need to move TMIP audit code to Current Build
'SF Commented out because the SP is not in CDR which the user may connect to using ICD
'  Call TMIPAudit("C2T_DAS", "ExecuteSQL", UserID, CurrentUserName, SqlStatement, Err.Description, moICDCnn)

#If DEBUGON Then
  If Not DL Is Nothing Then DL.DebugLog Exe, Module, "ExecuteSQL", "End", DR
#End If

  Exit Function

Func_Err:
  '<< SCR #44407;   Developer: Brian Mowbray 01/13/2004 02:08 PM
  Call HandleDBError(Err.Number, Err.Description, Module, moICDCnn)
  If Err <> 0 Then
    Err.Raise Err, Module, Err.Description
  End If
End Function

Private Sub Class_Initialize()
    'LoadRelevantWorkaroundTables
    
'    Dim objShared As CWShared
'
'    Set objShared = New CWShared
'    msUserNCID = objShared.CurrentUserNCID
'    msUserName = objShared.CurrentUserName
       
'Commented out per Phil 3-25-04. The developer must call the opendatabase method
'    Call OpenDatabase
    
'    Set objShared = Nothing
    
    'msDBKey = "8Cy2A4AmB5m97f_111AeB6_1"
    
    #If DEBUGON Then
      Set DL = GetObject("", "dlog.cdebuglog")
      If Not DL Is Nothing Then DL.debugStart
      Dim DR As DebugRec
      If Not DL Is Nothing Then DL.DebugLog Exe, Module, "Initialize", "", DR
    #End If

End Sub

Private Sub Class_Terminate()

On Error Resume Next

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
Public Sub OpenDatabase(Optional ConnectionString As String = "", Optional DBKey As String = "8Cy2A4AmB5m97f_111AeB6_1")
'- Initiate a connection to the Oracle db

Dim oRs As Recordset

    On Error GoTo Sub_Err
    
    msDBKey = DBKey
    
    'Close the existing connection so we can re-open it with a new connection string if needed.
    Set moICDCnn = Nothing
    
    mlICDCnnCnt = mlICDCnnCnt + 1
    If moICDCnn Is Nothing Then
        Set moICDCnn = New ADODB.Connection
        
        If ConnectionString = "" Then
            moICDCnn.ConnectionString = GetConnectionString
        Else
            'moICDCnn.ConnectionString = "Provider=OraOLEDB.Oracle.1;User Id=ico;Password=ico;Data Source=PGUI-CHCSIIbad"
            moICDCnn.ConnectionString = ConnectionString
        End If
        
        moICDCnn.ConnectionTimeout = 30
        moICDCnn.CursorLocation = adUseClient
    End If
    If moICDCnn.State <> 1 Then
        moICDCnn.Open
    End If
    
    'Use the microsoft provider for stored procedures to work with cursor refs
    'Close the existing connection so we can re-open it with a new connection string if needed.
    Set moICDSPCnn = Nothing
    
    If moICDSPCnn Is Nothing Then
        Set moICDSPCnn = New ADODB.Connection
        
'        If ConnectionString = "" Then
'            moICDCnn.ConnectionString = GetConnectionString
'        Else
        moICDSPCnn.ConnectionString = ReplaceParameter(ConnectionString, "PROVIDER=", "MSDAORA")
'        End If
        
        moICDSPCnn.ConnectionTimeout = 30
        moICDSPCnn.CursorLocation = adUseClient
    End If
    If moICDSPCnn.State <> 1 Then
        moICDSPCnn.Open
    End If
    
'    Dim i As Integer
'    For i = 0 To moICDCnn.Properties.Count - 1
'    Debug.Print moICDCnn.Properties(i).Name
'    Next i
    
    If moICDCnn.State = 1 Then
        'SF not used for ICD mode
        'Call SetORADBDateFormat(moICDCnn)
        If msDBKey <> "" Then
            'SF 4-21-04 Make sure the mmi_x table exists before executing the delete and insert statements
            Set oRs = moICDCnn.Execute("select * from all_tables where table_name = 'MMI_X'")
            If oRs.RecordCount > 0 Then
                moICDCnn.Execute "delete from mmi_x"
                moICDCnn.Execute "Insert into mmi_x (x0) values ('" & msDBKey & "')"
                
                '5-20-04 Do this for the stored procedure connection also
                moICDSPCnn.Execute "delete from mmi_x"
                moICDSPCnn.Execute "Insert into mmi_x (x0) values ('" & msDBKey & "')"
            End If
            oRs.Close
            Set oRs = Nothing
        End If
    End If
    
    Exit Sub
    
Sub_Err:
  '<< SCR #44407;   Developer: Brian Mowbray 01/13/2004 02:08 PM
  Call HandleDBError(Err.Number, Err.Description, Module, moICDCnn)
  'This was added to trap for error raised by the HandleDBError if the error is
  'not related to the loss of connectivity.
  If Err.Number = -2147467259 Then 'TNS:could not resolve service name
    'Do nothing
  ElseIf Err.Number <> 0 Then
    Err.Raise Err.Number, Module, Err.Description
  End If
  Exit Sub
  Resume
End Sub

'<< SCR 33533 encrypt db
Private Function GetConnectionString() As String
    
On Error GoTo ErrHnd

    Dim sNSN As String
    Dim sUid As String
    Dim sPWD As String
    Dim oShared As CWShared
    Dim sDBKey As String
    
    Set oShared = New CWShared
    sNSN = oShared.NetServiceName
    GetConnectionString = gcICD_Conn
    If sNSN <> "" Then
        GetConnectionString = ReplaceParameter(GetConnectionString, "DATA SOURCE=", sNSN)
    End If
       
'-- get the local uid and pwd if available
    If oShared.GetLocalLogon(sUid, sPWD, sDBKey) Then
        '__ replace the user id and password
        GetConnectionString = ReplaceParameter(GetConnectionString, "USER ID=", sUid)
        GetConnectionString = ReplaceParameter(GetConnectionString, "PASSWORD=", sPWD)
        If sDBKey <> "" Then
            msDBKey = sDBKey
        End If
    End If
    Set oShared = Nothing
    
Exit Function

ErrHnd:
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

Private Sub CloseDatabase()

On Error GoTo ErrHnd

    mlICDCnnCnt = mlICDCnnCnt - 1
    If mlICDCnnCnt <= 0 Then
        If moICDCnn Is Nothing Then Exit Sub
        If moICDCnn.State > 0 Then moICDCnn.Close
        Set moICDCnn = Nothing
        
        If moICDSPCnn Is Nothing Then Exit Sub
        If moICDSPCnn.State > 0 Then moICDSPCnn.Close
        Set moICDSPCnn = Nothing
    End If
    
Exit Sub

ErrHnd:
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub

Private Function ICHCSII_DAS_OpenRecordset(SQL As String, _
                            Optional ByVal CursorType As ADODB.CursorTypeEnum = 0&, _
                            Optional ByVal LockType As ADODB.LockTypeEnum = 1&) As ADODB.Recordset

    Dim objRecordset As ADODB.Recordset
    Dim CmdType As Long
    
    On Error GoTo Func_Err
    
    Set objRecordset = New ADODB.Recordset
    
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
'    For Each v In moICDCnn.Properties
'        Debug.Print v.Name, v.Type, v.Value
'    Next v
    objRecordset.Open SQL, moICDCnn, CursorType, LockType, CmdType
    
    If Not objRecordset.EOF Then
        objRecordset.MoveFirst
    End If
    
    '5-20-04 Ed Clark requested to set the active connection to nothing
    Set objRecordset.ActiveConnection = Nothing
    
Func_Exit:
  On Error Resume Next
  Set ICHCSII_DAS_OpenRecordset = objRecordset
  Set objRecordset = Nothing
  Exit Function

Func_Err:
  '<< SCR #44407;   Developer: Brian Mowbray 01/13/2004 02:06 PM
  Call HandleDBError(Err.Number, Err.Description, Module, moICDCnn)
  'This was added to trap for error raised by the HandleDBError if the error is
  'not related to the loss of connectivity.
  If Err.Number <> 0 Then
    Err.Raise Err.Number, Module, Err.Description
  End If
  Resume Func_Exit

End Function

Private Function ICHCSII_DAS_OpenRecordsetMSP(Parms As Collection) As Collection

    Err.Raise 1001, "ICD_DAS", "OpenRecordsetMSP is not supported in ICD DAS."

End Function

Private Function ICHCSII_DAS_OpenRecordsetSP(Parms As CHCSII_Command.CHCSII_Cmd) As ADODB.Recordset

    Err.Raise 1001, "ICD_DAS", "OpenRecordsetSP is not supported in ICD DAS. Please use the ADO_DAS method in CHCSII_CONN instead."

End Function

Public Property Let ICHCSII_DAS_UserNCID(RHS As String)

On Error GoTo ErrHnd

    CurrentUserNCID RHS
    
Exit Property

ErrHnd:
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Property
 
Public Function ICHCSII_DAS_BeginTrans() As Boolean
    moICDCnn.BeginTrans
End Function
Public Function ICHCSII_DAS_CommitTrans() As Boolean
    
    moICDCnn.CommitTrans

End Function
Public Function ICHCSII_DAS_RollbackTrans() As Boolean
    moICDCnn.RollbackTrans
End Function

Private Function ISecSqlOps_Connect() As Boolean
    
    Err.Raise 1001, "CDR_DAS", "ISecSqlOps_Connect is no longer supported."
    
End Function
Public Function SecDisConnect() As Boolean
'__ 9.11.02 Shaw: added back to this module so we can run without tmip
'- terminate a "secure" connection to the Oracle db
    
On Error GoTo ErrHnd

    If Not moICDCnn Is Nothing Then
        If moICDCnn.State > 0 Then moICDCnn.Close
        Set moICDCnn = Nothing
    End If

    Call OpenDatabase
    
    SecDisConnect = True

Exit Function

ErrHnd:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function ICHCSII_DAS_ExecuteSP(Params As CHCSII_Command.CHCSII_Cmd) As Variant
 
    Err.Raise 1001, "ICD_DAS", "ExecuteSP is not supported in the ICHCSII_DAS interface. Please use the ADO_DAS method of CHCSII_CONN."
 
End Function
Private Function ICHCSII_DAS_ExecuteMSP(Parms As Collection) As Collection
 
End Function


