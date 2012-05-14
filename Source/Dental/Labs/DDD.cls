VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDDD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''************************************************************************
'''Module:        clsDDD/DDD.cls
'''
'''Description:   Dental Data Defenition object.
'''               Names of the Dental Tables and associated fields and Sequences
'''
'''Modules Required: basDDD (allows to use single array of names for multiple objects)
'''
'''Note:
'''
'''Warning:       pbIncludeTableAlias switch is ignored by psATable property
'''
'''Usage:
'''
''' Add DDD.cls and DDD.bas modules to your project
''' Declare object(s) of clsDDD class
''' Use it to build SQL statements
'''
''' Example:
'''   Global gobjDDD    As New clsDDD
'''   Dim sSql          As String
'''
'''With gobjDDD
'''   .pbIncludeSchema = True
'''   .pbIncludeTable = True
'''   .pbIncludeTableAlias = True
'''   .pbIncludeColumnAlias = False
'''
'''   sSql = "SELECT " _
'''        & .psDX_ICD_CODE_DEFINITION(lDICD_ICD_Category) & "," _
'''        & .psDX_ICD_CODE_DEFINITION(lDICD_ICD_Code) & "," _
'''        & .psDX_ICD_CODE_DEFINITION(lDICD_ICD_Definition) & "," _
'''        & .psDX_ICD_CODE_DEFINITION(lDICD_Dx_ICD_Code_ID)
'''   sSql = sSql _
'''        & " FROM " _
'''        & .psDX_ICD_CODE_DEFINITION & _
'''          " WHERE " _
'''        & .psDX_ICD_CODE_DEFINITION(lDICD_Chart_Related) & "=0 " & _
'''          " ORDER BY " _
'''        & .psDX_ICD_CODE_DEFINITION(lDICD_ICD_Category) & "," _
'''        & .psDX_ICD_CODE_DEFINITION(lDICD_ICD_Definition)
'''End With
'''
'''As result sSql variable will contain following string:
'''        SELECT DICD.ICD_Category,DICD.ICD_Code,DICD.ICD_Definition,DICD.Dx_ICD_Code_ID
'''        FROM DENTAL.DX_ICD_CODE_DEFINITION DICD
'''        WHERE DICD.Chart_Related=0
'''        ORDER BY DICD.ICD_Category,DICD.ICD_Definition
'''
'''Author         : Alexander Alberstein
'''Modified       :
'''               :
'''************************************************************************
Option Explicit
'Public Property Get plALL_FIELDS()
''''''''''''''' Constanst Section '''''''''''''''''''''''''''''''''''''''''

Private Const msMODULE_NAME      As String = "clsDDD."

Private Const msSCHEMA_DENTAL    As String = "DENTAL."

Private Const mlMAX_COLUMNS      As Long = 20
Private Const plDDD_TABLE_NAME        As Long = 0
Private Const mlIDX_NAME         As Long = 0
Private Const mlIDX_ALIAS        As Long = 1

''''''''''''''' Variables Section '''''''''''''''''''''''''''''''''''''''''

'''Indicates whether schema name will preface table name. Defaults to true
Public pbIncludeSchema        As Boolean
'''Indicates whether Table name will preface field name.
Public pbIncludeTable         As Boolean '''Defaults to false
'''Indicates whether Alias will be associated with table name
Public pbIncludeTableAlias    As Boolean '''Defaults to false
'''Indicates whether Column Heading will be associated with column name
Public pbIncludeColumnAlias  As Boolean '''Defaults to false



Public Function Mtd_sName(ByVal lTable As pnumDTables, Optional lColumn As Long = plDDD_TABLE_NAME) As String
'''************************************************************************
'''Called By:     All Properties except psATable and psSequence
'''
'''
'''Arguments:
''' lTable        Table Index
''' lColumn       Column index
'''
'''Description:   Return name of the requested entity
'''
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If successful  - name of the requested entity
'''               Else           - Empty string
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Mtd_sName"
   
   Dim sResult          As String
   Dim sHeader          As String
   Dim bValid           As Boolean
   
On Error GoTo ErrHnd
   
   ''''''If column
   If lColumn > plDDD_TABLE_NAME Then
      If lColumn > plDDD_ALL_FIELDS Then
         sResult = Trim$(DDD_sName(lTable, lColumn, mlIDX_NAME))
         bValid = Not sResult = ""
         If pbIncludeColumnAlias And bValid Then
            sHeader = Trim$(DDD_sName(lTable, lColumn, mlIDX_ALIAS))
            If sHeader = "" Then
               '''Use Column name
               sHeader = sResult
            Else
               sHeader = "'" & sHeader & "'"
            End If
            '''Append column header
            sResult = sResult & " " & sHeader
         End If
      Else
         sResult = "*"
         bValid = True
      End If
      If bValid Then
         '''If alias is used
         If pbIncludeTableAlias Then
            sResult = DDD_sName(lTable, plDDD_TABLE_NAME, mlIDX_ALIAS) & "." & sResult
         ElseIf pbIncludeTable Then
            sResult = DDD_sName(lTable, plDDD_TABLE_NAME, mlIDX_NAME) & "." & sResult
            
            If pbIncludeSchema Then
               sResult = msSCHEMA_DENTAL & sResult
            End If
         End If
      Else
         ''' Use LibApp_Error if all dental developers will use Lib modules
         '''Call LibApp_Error _
         '''     (1, gsUNDEFINED & gsCOLUMN, msMODULE_NAME, sPROC_NAME, gsTABLE & lTable & ", " & gsCOLUMN & lColumn)
         Call gobjShared.ShowVBError _
            (1, "Undefined Column" & vbCrLf & "TABLE: " & lTable & ", " & "COLUMN: " & lColumn, _
             msMODULE_NAME & sPROC_NAME, App.Title, vbCritical _
            )
      End If
   Else
      '''Table name
      If pbIncludeSchema Then
         sResult = msSCHEMA_DENTAL & DDD_sName(lTable, plDDD_TABLE_NAME, mlIDX_NAME)
      Else
         sResult = DDD_sName(lTable, plDDD_TABLE_NAME, mlIDX_NAME)
      End If
      If pbIncludeTableAlias Then
         sResult = sResult & " " & DDD_sName(lTable, plDDD_TABLE_NAME, mlIDX_ALIAS)
      End If
   End If
   
   Mtd_sName = sResult
   
Exit Function
ErrHnd:
   Mtd_sName = ""
   ''' Use LibApp_Error if all dental developers will use Lib modules
   '''Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsTABLE & lTable & ", " & gsCOLUMN & lColumn)
   Call gobjShared.ShowVBError _
      (Err.Number, Err.Description & vbCrLf & "TABLE: " & lTable & ", " & "COLUMN: " & lColumn, _
       msMODULE_NAME & sPROC_NAME, App.Title, vbCritical _
      )
Exit Function
End Function
Public Property Get plALL_FIELDS()
   plALL_FIELDS = plDDD_ALL_FIELDS
End Property

Public Property Get psDENTAL_ENC_SURFACE_SET(Optional ByVal lDef As pnumDESS_DENTAL_ENC_SURFACE_SET = lDESS_aTable) As String
   psDENTAL_ENC_SURFACE_SET = Mtd_sName(lDT_DENTAL_ENC_SURFACE_SET, lDef)
End Property
Public Property Get psDENTAL_ENCOUNTER(Optional ByVal lDef As pnumDE_DENTAL_ENCOUNTER = lDE_aTable) As String
   psDENTAL_ENCOUNTER = Mtd_sName(lDT_DENTAL_ENCOUNTER, lDef)
End Property
Public Property Get psDENTAL_ENCOUNTER_DX(Optional ByVal lDef As pnumDEDX_DENTAL_ENCOUNTER_DX = lDEDX_aTable) As String
   psDENTAL_ENCOUNTER_DX = Mtd_sName(lDT_DENTAL_ENCOUNTER_DX, lDef)
End Property
Public Property Get psDENTAL_ENCOUNTER_DXS(Optional ByVal lDef As pnumDEXS_DENTAL_ENCOUNTER_DXS = lDEXS_aTable) As String
   psDENTAL_ENCOUNTER_DXS = Mtd_sName(lDT_DENTAL_ENCOUNTER_DXS, lDef)
End Property
Public Property Get psDENTAL_ENCOUNTER_TEETH(Optional ByVal lDef As pnumDET_DENTAL_ENCOUNTER_TEETH = lDET_aTable) As String
   psDENTAL_ENCOUNTER_TEETH = Mtd_sName(lDT_DENTAL_ENCOUNTER_TEETH, lDef)
End Property
Public Property Get psDENTAL_ENCOUNTER_TX(Optional ByVal lDef As pnumDETX_DENTAL_ENCOUNTER_TX = lDETX_aTable) As String
   psDENTAL_ENCOUNTER_TX = Mtd_sName(lDT_DENTAL_ENCOUNTER_TX, lDef)
End Property
Public Property Get psDENTAL_ENCOUNTER_TXS(Optional ByVal lDef As pnumDETXS_DENTAL_ENCOUNTER_TXS = lDETXS_aTable) As String
   psDENTAL_ENCOUNTER_TXS = Mtd_sName(lDT_DENTAL_ENCOUNTER_TXS, lDef)
End Property
Public Property Get psDENTAL_RADIOGRAPHS(Optional ByVal lDef As pnumDR_DENTAL_RADIOGRAPHS = lDR_aTable) As String
   psDENTAL_RADIOGRAPHS = Mtd_sName(lDT_DENTAL_RADIOGRAPHS, lDef)
End Property
Public Property Get psDX_ICD_CODE_DEFINITION(Optional ByVal lDef As pnumDICD_DX_ICD_CODE_DEFINITION = lDICD_aTable) As String
   psDX_ICD_CODE_DEFINITION = Mtd_sName(lDT_DX_ICD_CODE_DEFINITION, lDef)
End Property

Public Property Get psATable(ByVal lTable As pnumDTables) As String
'''************************************************************************
'''Called By:     Any
'''
'''
'''Arguments:
''' lTable        Table Index
'''
'''Description:   Return name of the requested table
'''
'''Note:
'''
'''Warning:       pbIncludeTableAlias switch is ignored by this property
'''
'''Updates Globals:
'''
'''Returns:       If successful  - name of the requested Sequence
'''               Else           - Empty string
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "psATable"
   
   Dim sResult          As String
   
On Error GoTo ErrHnd
      
   sResult = DDD_sName(lTable, plDDD_TABLE_NAME, mlIDX_NAME)
   
   If pbIncludeSchema Then
      sResult = msSCHEMA_DENTAL & sResult
   End If
   
   psATable = sResult
   
Exit Property
ErrHnd:
   psATable = ""
   ''' Use LibApp_Error if all dental developers will use Lib modules
   '''Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsTABLE & lTable & ", " & gsCOLUMN & lColumn)
   Call gobjShared.ShowVBError _
   (Err.Number, Err.Description & vbCrLf & "TABLE: " & lTable, msMODULE_NAME & sPROC_NAME, App.Title, vbCritical)
Exit Property
End Property


Public Property Get psASequence(ByVal lTable As pnumDTables, Optional bAddNextVal As Boolean = False) As String
'''************************************************************************
'''Called By:     Any
'''
'''
'''Arguments:
''' lTable        Table Index
''' bAddNextVal   Indicates whether to append .NEXTVAL to Sequence Name
'''Description:   Return name of the Sequence corresponding to requested table
'''
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If successful  - name of the requested Sequence
'''               Else           - Empty string
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "psASequence"
   
   Dim sResult          As String
   
On Error GoTo ErrHnd
   
   psASequence = psATable(lTable) & "_SEQ" & IIf(bAddNextVal, ".NEXTVAL", "")
   
Exit Property
ErrHnd:
   psASequence = ""
   ''' Use LibApp_Error if all dental developers will use Lib modules
   '''Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsTABLE & lTable & ", " & gsCOLUMN & lColumn)
   Call gobjShared.ShowVBError _
   (Err.Number, Err.Description & vbCrLf & "TABLE: " & lTable, msMODULE_NAME & sPROC_NAME, App.Title, vbCritical)
Exit Property
End Property

Private Sub Class_Initialize()
'''************************************************************************
'''Called By:     Mtd_sReset
'''
'''Arguments
'''
'''Description:   Initialize object
'''
'''
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Class_Initialize"

On Error GoTo ErrHnd

   pbIncludeSchema = True
      
   Call DDD_Init '''Populate data names array
   
Exit Sub
ErrHnd:
   Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME & sPROC_NAME, App.Title, vbCritical)

Exit Sub
End Sub
Public Property Get psEVALUATION_SUBJECTIVE(Optional ByVal lDef As pnumDES_EVALUATION_SUBJECTIVE = lDES_aTable) As String
   psEVALUATION_SUBJECTIVE = Mtd_sName(lDT_EVALUATION_SUBJECTIVE, lDef)
End Property

Public Property Get psEVALUATION_TYPE(Optional ByVal lDef As pnumDEVT_EVALUATION_TYPE = lDEVT_aTable) As String
   psEVALUATION_TYPE = Mtd_sName(lDT_EVALUATION_TYPE, lDef)
End Property
Public Property Get psLAB_CASE(Optional ByVal lDef As pnumDLC_LAB_CASE = lDLC_aTable) As String
   psLAB_CASE = Mtd_sName(lDT_LAB_CASE, lDef)
End Property
Public Property Get psLAB_MATERIAL_ELEMENTS(Optional ByVal lDef As pnumDLME_LAB_MATERIAL_ELEMENTS = lDLME_aTable) As String
   psLAB_MATERIAL_ELEMENTS = Mtd_sName(lDT_LAB_MATERIAL_ELEMENTS, lDef)
End Property
Public Property Get psLAB_MATERIALS(Optional ByVal lDef As pnumDLM_LAB_MATERIALS = lDLM_aTable) As String
   psLAB_MATERIALS = Mtd_sName(lDT_LAB_MATERIALS, lDef)
End Property
Public Property Get psLAB_NOTE(Optional ByVal lDef As pnumDLN_LAB_NOTE = lDLN_aTable) As String
   psLAB_NOTE = Mtd_sName(lDT_LAB_NOTE, lDef)
End Property
Public Property Get psLAB_NOTES(Optional ByVal lDef As pnumDLNS_LAB_NOTES = lDLNS_aTable) As String
   psLAB_NOTES = Mtd_sName(lDT_LAB_NOTES, lDef)
End Property
Public Property Get psLAB_PART_SPECIFICATIONS(Optional ByVal lDef As pnumDLPS_LAB_PART_SPECIFICATIONS = lDLPS_aTable) As String
   psLAB_PART_SPECIFICATIONS = Mtd_sName(lDT_LAB_PART_SPECIFICATIONS, lDef)
End Property
Public Property Get psLAB_TASK_CATEGORY_DEF(Optional ByVal lDef As pnumDLTCD_LAB_TASK_CATEGORY_DEF = lDLTCD_aTable) As String
   psLAB_TASK_CATEGORY_DEF = Mtd_sName(lDT_LAB_TASK_CATEGORY_DEF, lDef)
End Property
Public Property Get psLAB_TASK_DEFINITION(Optional ByVal lDef As pnumDLTD_LAB_TASK_DEFINITION = lDLTD_aTable) As String
   psLAB_TASK_DEFINITION = Mtd_sName(lDT_LAB_TASK_DEFINITION, lDef)
End Property
Public Property Get psLAB_TX_SPECIFICATIONS(Optional ByVal lDef As pnumDLTS_LAB_TX_SPECIFICATIONS = lDLTS_aTable) As String
   psLAB_TX_SPECIFICATIONS = Mtd_sName(lDT_LAB_TX_SPECIFICATIONS, lDef)
End Property
Public Property Get psLAB_TX_TASK_SPECS(Optional ByVal lDef As pnumDLTTS_LAB_TX_TASK_SPECS = lDLTTS_aTable) As String
   psLAB_TX_TASK_SPECS = Mtd_sName(lDT_LAB_TX_TASK_SPECS, lDef)
End Property
Public Property Get psLAB_TX_TASKS(Optional ByVal lDef As pnumDLTT_LAB_TX_TASKS = lDLTT_aTable) As String
   psLAB_TX_TASKS = Mtd_sName(lDT_LAB_TX_TASKS, lDef)
End Property
Public Property Get psNOTE_CODES_DEFINITION(Optional ByVal lDef As pnumDNCD_NOTE_CODES_DEFINITION = lDNCD_aTable) As String
   psNOTE_CODES_DEFINITION = Mtd_sName(lDT_NOTE_CODES_DEFINITION, lDef)
End Property
Public Property Get psNOTE_CODES_TYPE_DEFINITION(Optional ByVal lDef As pnumDNCT_NOTE_CODES_TYPE_DEFINITION = lDNCT_aTable) As String
   psNOTE_CODES_TYPE_DEFINITION = Mtd_sName(lDT_NOTE_CODES_TYPE_DEFINITION, lDef)
End Property
Public Property Get psNOTE_EVALUATION(Optional ByVal lDef As pnumDNE_NOTE_EVALUATION = lDNE_aTable) As String
   psNOTE_EVALUATION = Mtd_sName(lDT_NOTE_EVALUATION, lDef)
End Property
Public Property Get psNOTE_MATERIALS(Optional ByVal lDef As pnumDNM_NOTE_MATERIALS = lDNM_aTable) As String
   psNOTE_MATERIALS = Mtd_sName(lDT_NOTE_MATERIALS, lDef)
End Property
Public Property Get psNOTE_MATERIALS_TREATMENT(Optional ByVal lDef As pnumDMT_NOTE_MATERIALS_TREATMENT = lDMT_aTable) As String
   psNOTE_MATERIALS_TREATMENT = Mtd_sName(lDT_NOTE_MATERIALS_TREATMENT, lDef)
End Property
Public Property Get psNOTE_PERIODIC(Optional ByVal lDef As pnumDNP_NOTE_PERIODIC = lDNP_aTable) As String
   psNOTE_PERIODIC = Mtd_sName(lDT_NOTE_PERIODIC, lDef)
End Property
Public Property Get psNOTE_PREVENTIVE(Optional ByVal lDef As pnumDNPR_NOTE_PREVENTIVE = lDNPR_aTable) As String
   psNOTE_PREVENTIVE = Mtd_sName(lDT_NOTE_PREVENTIVE, lDef)
End Property
Public Property Get psNOTE_RADIOGRAPH(Optional ByVal lDef As pnumDNR_NOTE_RADIOGRAPH = lDNR_aTable) As String
   psNOTE_RADIOGRAPH = Mtd_sName(lDT_NOTE_RADIOGRAPH, lDef)
End Property
Public Property Get psNOTE_TREATMENT(Optional ByVal lDef As pnumDNT_NOTE_TREATMENT = lDNT_aTable) As String
   psNOTE_TREATMENT = Mtd_sName(lDT_NOTE_TREATMENT, lDef)
End Property
Public Property Get psPREVENTIVE_COUNSELING(Optional ByVal lDef As pnumDTPC_PREVENTIVE_COUNSELING = lDTPC_aTable) As String
   psPREVENTIVE_COUNSELING = Mtd_sName(lDT_PREVENTIVE_COUNSELING, lDef)
End Property
Public Property Get psPREVENTIVE_DEMONSTRATION(Optional ByVal lDef As pnumDPD_PREVENTIVE_DEMONSTRATION = lDPD_aTable) As String
   psPREVENTIVE_DEMONSTRATION = Mtd_sName(lDT_PREVENTIVE_DEMONSTRATION, lDef)
End Property
Public Property Get psPREVENTIVE_FLUORIDE(Optional ByVal lDef As pnumDPF_PREVENTIVE_FLUORIDE = lDPF_aTable) As String
   psPREVENTIVE_FLUORIDE = Mtd_sName(lDT_PREVENTIVE_FLUORIDE, lDef)
End Property
Public Property Get psPREVENTIVE_POLISH(Optional ByVal lDef As pnumDPP_PREVENTIVE_POLISH = lDPP_aTable) As String
   psPREVENTIVE_POLISH = Mtd_sName(lDT_PREVENTIVE_POLISH, lDef)
End Property
Public Property Get psPREVENTIVE_SCALING(Optional ByVal lDef As pnumDPS_PREVENTIVE_SCALING = lDPS_aTable) As String
   psPREVENTIVE_SCALING = Mtd_sName(lDT_PREVENTIVE_SCALING, lDef)
End Property
Public Property Get psDENTAL_ANESTHETICS(Optional ByVal lDef As pnumDTA_DENTAL_ANESTHETICS = lDTA_aTable) As String
   psDENTAL_ANESTHETICS = Mtd_sName(lDT_DENTAL_ANESTHETICS, lDef)
End Property
Public Property Get psTREATMENT_INSULATION(Optional ByVal lDef As pnumDTI_TREATMENT_INSULATION = lDTI_aTable) As String
   psTREATMENT_INSULATION = Mtd_sName(lDT_TREATMENT_INSULATION, lDef)
End Property
Public Property Get psTX_ADA_CODE_DEFINITION(Optional ByVal lDef As pnumDTXD_TX_ADA_CODE_DEFINITION = lDTXD_aTable) As String
   psTX_ADA_CODE_DEFINITION = Mtd_sName(lDT_TX_ADA_CODE_DEFINITION, lDef)
End Property
Public Property Get psDENTAL_IMPRESSIONS(Optional ByVal lDef As pnumDTDI_DENTAL_IMPRESSIONS = lDTDI_aTable) As String
   psDENTAL_IMPRESSIONS = Mtd_sName(lDT_DENTAL_IMPRESSIONS, lDef)
   
End Property
Public Property Get psPREVENTIVE_DEBRIDEMENT(Optional ByVal lDef As pnumDTPD_PREVENTIVE_DEBRIDEMENT = lDTPD_aTable) As String
   psPREVENTIVE_DEBRIDEMENT = Mtd_sName(lDT_PREVENTIVE_DEBRIDEMENT, lDef)
End Property
Public Property Get psPREVENTIVE_SCALING_SUPRAGIN(Optional ByVal lDef As pnumDTPG_PREVENTIVE_SCALING_SUPRAGIN = lDTPG_aTable) As String
   psPREVENTIVE_SCALING_SUPRAGIN = Mtd_sName(lDT_PREVENTIVE_SCALING_SUPRAGIN, lDef)
End Property
Public Property Get psPREVENTIVE_DELIVERY(Optional ByVal lDef As pnumDTD_PREVENTIVE_DELIVERY = lDTD_aTable) As String
   psPREVENTIVE_DELIVERY = Mtd_sName(lDT_PREVENTIVE_DELIVERY, lDef)
End Property
Public Property Get psHYGIENE_MEDICAMENTS(Optional ByVal lDef As pnumDTHM_HYGIENE_MEDICAMENTS = lDTHM_aTable) As String
   psHYGIENE_MEDICAMENTS = Mtd_sName(lDT_HYGIENE_MEDICAMENTS, lDef)
End Property

