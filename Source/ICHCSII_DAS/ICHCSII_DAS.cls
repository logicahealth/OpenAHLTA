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
Attribute VB_Name = "ICHCSII_DAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Enum eSeq
    esMAINSEQ = 0
    esAPPT_ID
    esAPPT_TRANS_ID
    esENC_TEMPLATE_ID
    esENCOUNTER_ID
    esIEN
    esPCM_CODE
    esPI_DATA_ID
    esUNIT_NUMBER
    esUSER_NUMBER
    esJDOE_NUMBER
    esTEMP_SSN
    esDENTAL_ANESTHETICS_SEQ
    esDENTAL_ENCOUNTER_DXS_SS_SEQ
    esDENTAL_ENCOUNTER_DX_SEQ
    esDENTAL_ENCOUNTER_SEQ
    esDENTAL_ENCOUNTER_TEETH_SEQ
    esDENTAL_ENCOUNTER_TXS_SEQ
    esDENTAL_ENCOUNTER_TXS_SS_SEQ
    esDENTAL_ENCOUNTER_TX_SEQ
    esDENTAL_ENC_SURFACE_SET_SEQ
    esDENTAL_IMAGING_RADS_SEQ
    esDENTAL_IMAGING_VERSION_GRP_SEQ
    esDENTAL_IMPRESSIONS_SEQ
    esDENTAL_RADIOGRAPHS_SEQ
    esDX_ICD_CODE_DEFINITION_SEQ
    esGRAFT_ID_SEQ
    esGROUP_ID_SEQ
    esLAB_ADA_CLV_SEQ
    esLAB_CASE_SEQ
    esLAB_MATERIALS_SEQ
    esLAB_MATERIAL_ELEMENTS_SEQ
    esLAB_NOTE_SEQ
    esLAB_PART_SPECIFICATIONS_SEQ
    esLAB_TASK_CATEGORY_DEF_SEQ
    esLAB_TASK_DEFINITION_SEQ
    esLAB_TX_SPECIFICATIONS_SEQ
    esLAB_TX_TASK_SPECS_SEQ
    esNOTE_CODES_DEFINITION_SEQ
    esNOTE_CODES_TYPE_DEFINITION_SEQ
    esNOTE_DXS_LOCATION_SEQ
    esNOTE_DX_SEQ
    esNOTE_DX_SUB_LOCATION_SEQ
    esNOTE_TXS_LOCATION_SEQ
    esNOTE_TXS_SEQ
    esNOTE_TX_SUB_LOCATION_SEQ
    esPREVENTIVE_DEBRIDEMENT_SEQ
    esPREVENTIVE_DELIVERY_SEQ
    esPREVENT_SCALING_SUPRAGIN_SEQ
    esSURFACE_DEFINITION_SEQ
    esTEMPL_ANESTH_SEQ
    esTEMPL_CARIES_RISK_PROTOCOL_SEQ
    esTEMPL_CARIES_RISK_RECALL_V_SEQ
    esTEMPL_EXAM_ENDO_ADD_F_SEQ
    esTEMPL_EXAM_ENDO_COMMENTS_SEQ
    esTEMPL_EXAM_ENDO_ENDONTIC_T_SEQ
    esTEMPL_EXAM_ENDO_FINDINGS_SEQ
    esTEMPL_EXAM_ENDO_PAIN_CHARS_SEQ
    esTEMPL_EXAM_ENDO_SUBJECTIVE_SEQ
    esTEMPL_EXAM_PROS_DENTURE_F_SEQ
    esTEMPL_EXAM_PROS_DENTURE_SEQ
    esTEMPL_EXAM_XML_SEQ
    esTEMPL_GENPROS_IMPRESS_SEQ
    esTEMPL_NARRATIVE_SUMMARY_SEQ
    esTEMPL_PREVENTIVE_SEQ
    esTEMPL_RADIOLOGY_FINDINGS_SEQ
    esTEMPL_TREAT_NSE_CANAL_INFO_SEQ
    esTEMPL_TREAT_NSE_ENCOUNTER_SEQ
    esTEMPL_TREAT_NSE_TOOTH_INFO_SEQ
    esTEMPL_WORKLOAD_SEQ
    esTOOTH_DEFINITION_SEQ
    esTOOTH_ORIENT_DEFINITION_SEQ
    esTREATMENT_ANESTHETIC_SEQ
    esTX_ADA_CODE_DEFINITION_SEQ
End Enum
Public Property Let FacilityWorkaround(ByVal Enabled As Boolean):       End Property
Public Function GetNextID(Optional ByVal eSeqName As eSeq = esMAINSEQ): End Function
Public Function GetNextNcid():                                          End Function

Public Property Let UserNCID(RHV As String):                     End Property
Public Function OpenRecordset(SQL As String, _
               Optional ByVal CursorType As ADODB.CursorTypeEnum = adOpenForwardOnly, _
               Optional ByVal LockType As ADODB.LockTypeEnum = adLockReadOnly) As ADODB.Recordset: End Function
Public Function ExecuteSQL(SQL As String) As Variant:            End Function
Public Function CreateRecordsetFromFile(sFileName As String, _
                                        sDelimiter As String) As ADODB.Recordset: End Function
Public Function BeginTrans() As Boolean:             End Function
Public Function CommitTrans() As Boolean:            End Function
Public Function RollbackTrans() As Boolean:          End Function
Public Property Get ActiveConnection() As ADODB.Connection: End Property
Public Function OpenRecordsetSP(Parms As CHCSII_Command.CHCSII_Cmd) As ADODB.Recordset: End Function

Public Function ExecuteSP(Parms As CHCSII_Command.CHCSII_Cmd) As Variant: End Function



