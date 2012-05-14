VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDR_DAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'This class is used to access data from/ to the CDR via the DSql control
'provided by 3M.

Option Explicit

Implements ICHCSII_DataAccess.ICHCSII_DAS

'3-09-04 SF Added support for the ADO Command object
Implements IADO_DAS

'SF 1-22-04 Not used anymore
'DAS
'Private m_objDSQL As DSQLATLLib.DSqlConnect

'SPs
Private moSqlConnect As SQLSERVATLLib.SqlServConnect
Private msServiceName As String
'Private moADOCommand As ADODB.Command

'Commented out by SF 1-29-04, did not find any references to it.
'Private Const cnst_MAX_NUM_OF_RECORDS = 5000 '* Add by j.ray 31AUG2001 to fix SCR #16462

Private bSkipFacilityWorkAround As Boolean     '- so mid-tier can function...

'SF 1-22-04 Indicates if dates should be converted for the local time zone.
Private meTimeZoneConversion As eTimeZoneConversion



Public Sub Init(TimeZoneConversion As eTimeZoneConversion, ServiceName As eServiceNameMode)

    'SF 1-22-04 Indicates if dates should be converted for the local time zone.
    meTimeZoneConversion = TimeZoneConversion
    
    'SCR 46007
    'Developer: Scott Foster 12-6-03
    'Created to expose the service name without modifying ICHCSII_DAS
    msServiceName = TranslateServiceName(ServiceName)

End Sub


Private Sub Class_Initialize()
  
    bSkipFacilityWorkAround = False
    
'1-22-04 SF No longer used
'    'CDR
'    Set m_objDSQL = New DSQLATLLib.DSqlConnect
    
    'CDR_SP
    Set moSqlConnect = New SQLSERVATLLib.SqlServConnect
    
    LoadRelevantWorkaroundTables
    GetTimeZoneInformation gTZInfo

End Sub

Private Sub Class_Terminate()
    
    'CDR
    'Set m_objDSQL = Nothing
        
    'CDR_SP
    Set moSqlConnect = Nothing

End Sub


Private Sub CloseDatabase()

    Set moSqlConnect = Nothing

End Sub

Private Property Get IADO_DAS_ActiveConnection() As ADODB.Connection
    
    Set IADO_DAS_ActiveConnection = Nothing
    
End Property

Private Function IADO_DAS_BeginTrans() As Boolean

    Err.Raise 1001, "CDR_DAS", "BeginTrans is not supported in CDR DAS."

End Function



Private Function IADO_DAS_CommitTrans() As Boolean

    Err.Raise 1001, "CDR_DAS", "Transaction control is not supported in CDR DAS."

End Function


Private Function IADO_DAS_CreateRecordsetFromFile(sFileName As String, sDelimiter As String) As ADODB.Recordset
    
    'Use the ICHCSII_DAS_CreateRecordsetFromFile function to process this
    Set IADO_DAS_CreateRecordsetFromFile = ICHCSII_DAS_CreateRecordsetFromFile(sFileName, sDelimiter)
    
End Function


Private Function IADO_DAS_ExecuteSP(oCommand As Command) As Long
    
'3-09-04 SF
'This function takes the ADO command object, translates it to SQLServ bind variables and executes it through SQLServ

On Error GoTo ErrHnd

Dim oBindVar As SQLSERVATLLib.SqlServBindVar
Dim oStatement As SqlServStatement
Dim oEh As EMHATLLib.EmhCtrl

    moSqlConnect.ClearAll
    
    moSqlConnect.ServiceName = msServiceName
    
    Set oStatement = moSqlConnect.Transaction.Add
    
    ConvertADOCommandToSQLServStatement oCommand, oStatement
    
    moSqlConnect.ExecuteTransaction
    
    'IADO_DAS_ExecuteSP = moSqlConnect.Transaction(1).
    
Exit Function

ErrHnd:
    If Err.Number = -2147220887 Then
        'Generic service message so get the specific error message
        Set oEh = New EMHATLLib.EmhCtrl
        Err.Raise oEh.Number, oEh.Origin, oEh.Message
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If

End Function


Private Function IADO_DAS_ExecuteSQL(SQL As String) As Variant

    'Use the ICHCSII_DAS_ExecuteSQL function to process this
    Set IADO_DAS_ExecuteSQL = ICHCSII_DAS_ExecuteSQL(SQL)
    
End Function


Private Property Let IADO_DAS_FacilityWorkaround(ByVal RHS As Boolean)

    bSkipFacilityWorkAround = Not RHS

End Property

Private Function IADO_DAS_GetNextID(Optional ByVal eSeqName As eSeq = 0&) As Variant

    'Use the ICHCSII_DAS_GetNextID function to process this
    IADO_DAS_GetNextID = ICHCSII_DAS_GetNextID(eSeqName)
    
End Function


Private Function IADO_DAS_GetNextNcid() As Variant

    Err.Raise 1001, "CDR_DAS", "GetNextNcid is not supported in CDR DAS."

End Function


Private Function IADO_DAS_OpenRecordset(SQL As String, Optional ByVal CursorType As ADODB.CursorTypeEnum = 0&, Optional ByVal LockType As ADODB.LockTypeEnum = 1&) As ADODB.Recordset

    'Use the ICHCSII_DAS_OpenRecordset function to process this
    Set IADO_DAS_OpenRecordset = ICHCSII_DAS_OpenRecordset(SQL, CursorType, LockType)
    
End Function


Private Function IADO_DAS_OpenRecordsetSP(oCommand As ADODB.Command) As ADODB.Recordset
    
'3-09-04 SF
'This function takes the ADO command object, translates it to SQLServ bind variables and executes it through SQLServ

Dim oStatement As SqlServStatement
Dim oEh As EMHATLLib.EmhCtrl

On Error GoTo ErrHnd

    moSqlConnect.ClearAll
    
    moSqlConnect.ServiceName = msServiceName
    
    Set oStatement = moSqlConnect.Transaction.Add
    
    ConvertADOCommandToSQLServStatement oCommand, oStatement
    
    moSqlConnect.ExecuteTransaction
    
    Set IADO_DAS_OpenRecordsetSP = SServConvertStatementToRecordset(moSqlConnect.Transaction(1), True, meTimeZoneConversion)
    
Exit Function

ErrHnd:
    If Err.Number = -2147220887 Then
        'Generic service message so get the specific error message
        Set oEh = New EMHATLLib.EmhCtrl
        Err.Raise oEh.Number, oEh.Origin, oEh.Message
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Function


Private Function IADO_DAS_RollbackTrans() As Boolean

    Err.Raise 1001, "CDR_DAS", "RollbackTrans is not supported in CDR DAS."

End Function


Private Property Let IADO_DAS_UserNCID(RHS As String)

    Err.Raise 1001, "CDR_DAS", "UserNCID Let is not supported in CDR DAS."

End Property


'Private Function getColumnNames(oColumnNames As SQLSERVATLLib.SqlServColumnNames) As String
'
'    Dim iLoop As Integer
'    Dim sColumnNames As String
'
'    sColumnNames = ""
'    With oColumnNames
'    For iLoop = 1 To .Count
'        If (iLoop = 1) Then
'            sColumnNames = .Item(iLoop)
'        Else
'            sColumnNames = sColumnNames & " | " & .Item(iLoop)
'        End If
'    Next
'    End With
'
'    getColumnNames = sColumnNames
'
'End Function

Private Property Get ICHCSII_DAS_ActiveConnection() As ADODB.Connection
    
    Set ICHCSII_DAS_ActiveConnection = Nothing
    
End Property

Private Function ICHCSII_DAS_BeginTrans() As Boolean

    Err.Raise 1001, "CDR_DAS", "BeginTrans is not supported in CDR DAS."

End Function


Private Function ICHCSII_DAS_CommitTrans() As Boolean

    Err.Raise 1001, "CDR_DAS", "Transaction control is not supported in CDR DAS."

End Function


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
                            objRecordset.Fields.Append sColumnNames(i), adDate
                        Else
                            objRecordset.Fields.Append sColumnNames(i), adBSTR
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
            For i = 0 To objRecordset.Fields.Count - 1
                If objRecordset.Fields.Item(i).Type = adDate Then
                    objRecordset.Fields.Item(i).Value = CDate(sValues(i))
                Else
                    If sValues(i) <> "" Then
                        objRecordset.Fields.Item(i).Value = CStr(sValues(i))
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


Private Function ExecutMultiSP(mParams As Collection) As SqlServTransaction
   
'This function takes a collection of command objects,
'creates one or more SQLServATL statements, executes them and returns the transaction

'1/21/04 Modified by SF to return the transaction

    Dim vParam  As CHCSII_SPparam
    Dim sParam  As CHCSII_Cmd
    Dim oBindVar As SQLSERVATLLib.SqlServBindVar
    Dim i As Integer
    Dim oStatement As SqlServStatement
    
    moSqlConnect.ClearAll
    
    moSqlConnect.ServiceName = msServiceName
    
    For Each sParam In mParams
        Set oStatement = moSqlConnect.Transaction.Add
        
        With oStatement
            .ProcedureName = sParam.SPName
            For i = 1 To sParam.Count
                Set vParam = sParam.Item(i)
                Set oBindVar = .BindVariables.Add
                If vParam.InfoType <> "" Then
                    oBindVar.Value.InfoType = vParam.InfoType
                    If vParam.InfoType = "date" Then
                        oBindVar.Value.DateNoTZConversion = vParam.Param
                    ElseIf vParam.InfoType = "textString" Then
                        oBindVar.Value.TextString = vParam.Param
                    End If
                End If
                oBindVar.IsCursor = vParam.IsCursor
                oBindVar.IsOutput = vParam.IsOutput
            Next
        End With
    Next
    
    moSqlConnect.ExecuteTransaction
    
    Set ExecutMultiSP = moSqlConnect.Transaction

End Function

Private Function ICHCSII_DAS_ExecuteSP(Parms As CHCSII_Cmd) As Variant

    ExecuteSP Parms

End Function

Private Function ICHCSII_DAS_ExecuteSQL(SQL As String) As Variant

Dim oStatement As SqlServStatement

'For now, we raise the error to the client code.
On Error GoTo ErrHnd
    
    moSqlConnect.ClearAll
    
    moSqlConnect.ServiceName = msServiceName
    
    Set oStatement = moSqlConnect.Transaction.Add

    oStatement.SqlId.ValueConcept.UniqueId = gSQLQueryNCID

    oStatement.SqlStatement.Text = SQL

    ExecuteSQLServ moSqlConnect
    
    ICHCSII_DAS_ExecuteSQL = True

Exit Function

ErrHnd:
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function


Private Property Let ICHCSII_DAS_FacilityWorkaround(ByVal Enabled As Boolean)

    bSkipFacilityWorkAround = Not Enabled

End Property


Private Function ICHCSII_DAS_GetNextID(Optional ByVal eSeqName As ICHCSII_DataAccess.eSeq = esMAINSEQ) As Variant

'SF This was taken from GEMS_DAS
On Error GoTo ErrHandler
    
Dim oRs As New ADODB.Recordset
Dim sSeqName As String
    
    sSeqName = SeqEnumToString(eSeqName)
    Set oRs = ICHCSII_DAS_OpenRecordset("Select " & sSeqName & ".nextval from dual")

    If Not oRs.EOF Then
        ICHCSII_DAS_GetNextID = oRs(0)
    End If
    
ErrHandler:
    If oRs.State > 0 Then oRs.Close
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




Private Function ICHCSII_DAS_GetNextNcid() As Variant

    Err.Raise 1001, "CDR_DAS", "GetNextNcid is not supported in CDR DAS."

End Function


Private Function ICHCSII_DAS_OpenRecordset(SQL As String, Optional ByVal CursorType As ADODB.CursorTypeEnum = 0&, Optional ByVal LockType As ADODB.LockTypeEnum = 1&) As ADODB.Recordset

'Modified by: SF
'1-21-04 SF - This used to operate under DSQL, converted it to use SQLServATL

On Error GoTo ErrHnd

Dim oStatement As SqlServStatement
Dim oRs As Recordset
Dim sSQLTemp As String
Dim bError As Boolean
Dim bUseSQL As Boolean

    sSQLTemp = DoFacilityNCIDWorkaround(SQL, bSkipFacilityWorkAround)

    moSqlConnect.ClearAll

    'Set the service name
    moSqlConnect.ServiceName = msServiceName

    Set oStatement = moSqlConnect.Transaction.Add

    oStatement.SqlId.ValueConcept.UniqueId = gSQLQueryNCID
    
    oStatement.SqlStatement = sSQLTemp
    
'        'Bind variables
'        Dim oVars As SqlServBindVars, oVar As SqlServBindVar
''        Set oVars = oStatement.BindVariables
''        Set oVar = oVars.Add
'        Set oVar = oStatement.BindVariables.Add
'        oVar.Value.InfoType = "textString"
'        oVar.Value.TextString = "DEFAULT"
'        oStatement.SqlStatement = "Select * from MT_Options where setting = :X1"
'        moSqlConnect.ExecuteTransaction
        
    ExecuteSQLServ moSqlConnect
'    moSqlConnect.ExecuteTransaction
    
    Set oStatement = moSqlConnect.Transaction(1)

    'Convert the statement into a recordset
    Set oRs = SServConvertStatementToRecordset(oStatement, False, meTimeZoneConversion)
    
    Set ICHCSII_DAS_OpenRecordset = oRs
    
    Set oRs = Nothing
    Set oStatement = Nothing
    
Exit Function

ErrHnd:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function


Private Function ICHCSII_DAS_OpenRecordsetMSP(Parms As Collection) As Collection

    Err.Raise 1001, "CDR_DAS", "OpenRecordsetMSP is not supported in CDR DAS."

End Function


Private Function ICHCSII_DAS_OpenRecordsetSP(Parms As CHCSII_Cmd) As ADODB.Recordset

'1-21-04 SF Modified to remove the global statement object

Dim oStatement As SqlServStatement

    Set oStatement = ExecuteSP(Parms)
    
    Set ICHCSII_DAS_OpenRecordsetSP = SServConvertStatementToRecordset(oStatement, True, meTimeZoneConversion)
    
End Function


Private Function ICHCSII_DAS_RollbackTrans() As Boolean

    Err.Raise 1001, "CDR_DAS", "RollbackTrans is not supported in CDR DAS."
    
End Function


Private Property Let ICHCSII_DAS_UserNCID(RHS As String)

    Err.Raise 1001, "CDR_DAS", "UserNCID Let is not supported in CDR DAS."
    
End Property


Private Function ExecuteSP(Params As CHCSII_Cmd) As SqlServStatement

'This sub takes a command object
'adds it to a new collection and passes it to ExecuteMultiSP which returns the transaction
'1/21/04 Modified by SF to return the SQLServ Statement

    Dim mSPCols As Collection
    Set mSPCols = New Collection
    
    mSPCols.Add Params
    
    'Return the first and only transaction
    Set ExecuteSP = ExecutMultiSP(mSPCols)(1)
        
End Function

Private Function OpenRecordsetMSP(mSPCols As Collection) As Collection

    Dim Params As CHCSII_Cmd
    Dim mRS As Collection
    Dim rs As ADODB.Recordset
    Dim i As Integer
    Dim oTransaction As SqlServTransaction
    
    Set oTransaction = ExecutMultiSP(mSPCols)
    
    Set mRS = New Collection
 
    For i = 1 To oTransaction.Count
        Set rs = SServConvertStatementToRecordset(oTransaction(i), True, meTimeZoneConversion)
        mRS.Add rs
    Next
    
End Function

Private Function OpenRecordsetSP(Params As CHCSII_Cmd) As ADODB.Recordset

    Call ExecuteSP(Params)
    
    Set OpenRecordsetSP = SServConvertStatementToRecordset(moSqlConnect.Transaction(1), True, meTimeZoneConversion)

End Function

'Private Function getValue(sqlValue As SQLSERVATLLib.SqlServValue, IsUsingSP As Boolean) As Variant
'
''SF 1-22-04 Modified heavily to support backwards compatibility with the old DSQL and to fix
''time zone conversions
''SF 1-22-04 This has been replaced by SServGetValue

'Dim sInfoType As String
'Dim sStringValue As String
'
'    sInfoType = sqlValue.InfoType
'    Select Case sInfoType
'
'        Case "textString"
'            sStringValue = Trim(sqlValue.TextString)
'            If Not (sStringValue = "") Then
'                getValue = CStr(sStringValue)
'            End If
'        Case "binaryString"
'            getValue = sqlValue.BinaryString
'        Case "date"
'            If sqlValue.DateNoTZConversion <> #12/31/9999# Then
'
'                If meTimeZoneConversion = BadTimeZoneConversion Then
'                    If IsUsingSP = True Then
'                        'This is for backwards compatibility for the way time zone conversion in SPs used to be implemented
'                        'This is the old incorrect time zone conversion
'                        getValue = AdjustDate(sqlValue.DateNoTZConversion)
'                    Else
'                        'This is for backwards compatibility for the old DSQL time zone conversion
'                        getValue = sqlValue.DateNoTZConversion
'                    End If
'                ElseIf meTimeZoneConversion = NoTimeZoneConversion Then
'                    getValue = sqlValue.DateNoTZConversion
'
'                ElseIf meTimeZoneConversion = YesTimeZoneConversion Then
'
'                    'If the date does not have a date, but only a time, then don't do a time zone conversion.
'                    'Note: This includes dates of 12/30/1899 12:00:00 AM through 12/31/1899 11:59:59 AM
'                    'A birth date is the only date that could be in this range and birth dates should not
'                    'be converted to a local time zone anyway.
'                    If CLng(sqlValue.DateNoTZConversion) = 0 Or CLng(sqlValue.DateNoTZConversion) = 1 Then
'                        getValue = sqlValue.DateNoTZConversion
'
''Commented out for now, until we decide to use it
''                    'If the date does not have a time, but does have a date, then don't do a time zone conversion. (BirthDays)
''                    ElseIf Hour(sqlValue.DateNoTZConversion) = 0 And Minute(sqlValue.DateNoTZConversion) = 0 And Second(sqlValue.DateNoTZConversion) = 0 Then
''                        getValue = sqlValue.DateNoTZConversion
'
'                    Else
'                        getValue = sqlValue.Date
'                    End If
'
'                End If
'            Else
'                getValue = sqlValue.Date
'            End If
'
''After discussing with Kieran, we have determined that this did not need to be commented out on 10/29/2002.
''' Begin:SCR #28638;   Developer: Jane Shen 10/29/2002 06:01 PM
'''            If sqlValue.Date <> #12/31/9999# Then
''                getValue = AdjustDate(sqlValue.DateNoTZConversion)
'''            Else
'''                getValue = sqlValue.Date
'''            End If
''' End: SCR #28638;
'
'    End Select
'
'End Function
'
Private Function DoFacilityNCIDWorkaround(sSQL As String, ByVal SkipFacilityWorkAround As Boolean) As String

Dim iFromIndex As Integer
Dim iWhereIndex As Integer
Dim iOrderByIndex As Integer
Dim iUnionIndex As Integer
Dim sTableNames As String
Dim saTableNames() As String
Dim bIsTableRelevant As Boolean
Dim bDoLabTestFix As Boolean
Dim sTableToCheck As String
Dim i As Integer
Dim j As Integer
Dim sAdditionalWhereClause As String
Dim sBeginClause As String
Dim sEndClause As String
Dim sNewSQL As String
Dim sFromClause As String
Dim sFacilityNCID As String

    If Len(sSQL) < 1 Or SkipFacilityWorkAround Then
        DoFacilityNCIDWorkaround = sSQL
        Exit Function
    End If

    Dim oShared As MMMHISShared.CWShared
    Set oShared = New MMMHISShared.CWShared
    sFacilityNCID = oShared.CurrentFacilityNCID
    Set oShared = Nothing

    'Check for UNION
    iUnionIndex = InStr(1, sSQL, "UNION ", vbTextCompare)
    If (iUnionIndex > 0) Then
        DoFacilityNCIDWorkaround = HandleUnion(sSQL, SkipFacilityWorkAround)
        Exit Function
    End If
    
    'Parse SQL statement looking for FROM, WHERE, ORDER BY
    iFromIndex = InStr(1, sSQL, "FROM ", vbTextCompare)
    If (iFromIndex < 1) Then
        DoFacilityNCIDWorkaround = sSQL
        Exit Function
    End If
    
    iWhereIndex = InStr(iFromIndex, sSQL, "WHERE ", vbTextCompare)
    If (iWhereIndex > 0) Then
        iOrderByIndex = InStr(iWhereIndex, sSQL, "ORDER BY ", vbTextCompare)
        sTableNames = Mid(sSQL, iFromIndex + 5, iWhereIndex - (iFromIndex + 5))
    Else
        iOrderByIndex = InStr(iFromIndex, sSQL, "ORDER BY ", vbTextCompare)
        If (iOrderByIndex > 0) Then
            sTableNames = Mid(sSQL, iFromIndex + 5, iOrderByIndex - (iFromIndex + 5))
        Else
            sTableNames = Right(sSQL, Len(sSQL) - (iFromIndex + 4))
        End If
    End If
    
    saTableNames = Split(sTableNames, ",")
    For i = 0 To UBound(saTableNames)
        saTableNames(i) = Trim(saTableNames(i))
    Next
    
    'Figure out if we care about these tables
    bIsTableRelevant = False
    bDoLabTestFix = False
    For i = 0 To UBound(saTableNames)
        sTableToCheck = UCase(saTableNames(i))
        
        If (sTableToCheck = "LAB_TEST") Then
            bDoLabTestFix = True
        End If
        
        For j = 1 To UBound(gRelevantTables)
            If (sTableToCheck = gRelevantTables(j)) Then
                bIsTableRelevant = True
                Exit For
            End If
        Next
        
        If (bIsTableRelevant) Then
            Exit For
        End If
    Next
    
    If (bIsTableRelevant) Then
        'Now that we care about the table(s), make the Facility_NCID fix
        For i = 0 To UBound(saTableNames)
            If (i > 0) Then
                sAdditionalWhereClause = sAdditionalWhereClause & " And " & saTableNames(i) & ".facility_ncid = " & sFacilityNCID
            Else
                sAdditionalWhereClause = saTableNames(i) & ".facility_ncid = " & sFacilityNCID
            End If
        Next
        
        If (iWhereIndex > 0) Then
            'Have WHERE clause
            sBeginClause = Left(sSQL, iWhereIndex + 5)
            sEndClause = Right(sSQL, Len(sSQL) - (iWhereIndex + 5))
            sNewSQL = sBeginClause & sAdditionalWhereClause & " AND " & sEndClause
        ElseIf (iWhereIndex < 1 And iOrderByIndex > 0) Then
            'Only have ORDER BY clause
            sBeginClause = Left(sSQL, iOrderByIndex - 1)
            sEndClause = Right(sSQL, Len(sSQL) - (iOrderByIndex - 1))
            sNewSQL = sBeginClause & " WHERE " & sAdditionalWhereClause & " " & sEndClause
        Else
            'Only FROM clause
            sNewSQL = sSQL & " WHERE " & sAdditionalWhereClause
        End If
        
        'Do special fix for lab_test table
        If (bDoLabTestFix) Then
            sBeginClause = Left(sNewSQL, iFromIndex + 4)
            
            iWhereIndex = InStr(iFromIndex, sNewSQL, "WHERE ", vbTextCompare)
            sEndClause = Right(sNewSQL, Len(sNewSQL) - (iWhereIndex - 1))
            
            For i = 0 To UBound(saTableNames)
                If (UCase(saTableNames(i)) = "LAB_TEST") Then
                    If (i > 0) Then
                        sFromClause = sFromClause & ", lab_test_oe lab_test"
                    Else
                        sFromClause = "lab_test_oe lab_test"
                    End If
                Else
                    If (i > 0) Then
                        sFromClause = sFromClause & ", " & saTableNames(i)
                    Else
                        sFromClause = saTableNames(i)
                    End If
                End If
            Next
            
            sNewSQL = sBeginClause & sFromClause & " " & sEndClause
            
        End If
        
        If Right(sNewSQL, 1) = ";" Then
            sNewSQL = Left(sNewSQL, Len(sNewSQL) - 1)
        End If
    
        DoFacilityNCIDWorkaround = sNewSQL
    Else
        DoFacilityNCIDWorkaround = sSQL
    End If

End Function

Private Function HandleUnion(sSQL As String, ByVal SkipFacilityWorkAround As Boolean) As String
Dim sSQLSegments() As String
Dim iUnionIndex As Integer
Dim sNewSQL As String
Dim i As Integer

    iUnionIndex = InStr(1, sSQL, " UNION ", vbTextCompare)
    If (iUnionIndex > 0) Then
        sSQLSegments = Split(UCase$(sSQL), " UNION ")
        For i = 0 To UBound(sSQLSegments)
            If (i > 0) Then
                sNewSQL = sNewSQL & " UNION " & DoFacilityNCIDWorkaround(sSQLSegments(i), SkipFacilityWorkAround)
            Else
                sNewSQL = DoFacilityNCIDWorkaround(sSQLSegments(i), SkipFacilityWorkAround)
            End If
        Next
            
        HandleUnion = sNewSQL
        Exit Function
    Else
        HandleUnion = sSQL
    End If

End Function

Private Sub LoadRelevantWorkaroundTables()

    ReDim gRelevantTables(1 To 15)
    
    gRelevantTables(1) = "LAB_PANEL"
    gRelevantTables(2) = "LAB_SPECIMEN"
    gRelevantTables(3) = "LAB_TEST"
    gRelevantTables(4) = "LAB_TEST_SYNONYM"
    gRelevantTables(5) = "LAB_TEST_TYPE"
    gRelevantTables(6) = "LAB_TOPO"
    gRelevantTables(7) = "PHARMACY"
    gRelevantTables(8) = "RAD_TEST"
    gRelevantTables(9) = "RAD_TEST_SYNONYM"
    gRelevantTables(10) = "RAD_TEST_TYPE"
    gRelevantTables(11) = "RX_LIST"
    gRelevantTables(12) = "RX_SYNONYM"
    gRelevantTables(13) = "LAB_TEST_WTYPE"
    gRelevantTables(14) = "MEPRS"
    '<< SCR 19158  Added:
    gRelevantTables(15) = "LAB_TEST_OE"
    'gRelevantTables(16) = "LAB_TEST_OE"

End Sub



