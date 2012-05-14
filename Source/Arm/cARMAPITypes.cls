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
Attribute VB_Name = "cARMAPITypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit



'=====================================================================================
'                   ARM DATA BUFFER COMPONENT STRUCTS / MISC CONSTANTS
'=====================================================================================

Public Enum BitMask_Enum
                        'BIT POSITION
                        '76543210
  eBitMask7 = gbyBitMask7 '10000000
  eBitMask6 = gbyBitMask6 '01000000
  eBitMask5 = gbyBitMask5 '00100000
  eBitMask4 = gbyBitMask4 '00010000
  eBitMask3 = gbyBitMask3 '00001000
  eBitMask2 = gbyBitMask2 '00000100
  eBitMask1 = gbyBitMask1 '00000010
  eBitMask0 = gbyBitMask0 '00000001
End Enum



Public Enum ARM_Metrics_Flag_Enum
  'REMEMBER TO CONVERT LONG ENUM TO BYTE
  'FOR READING / SETTING BIT FLAGS IN BYTE
                                'BIT POSITIONS:
                                '76543210
  eMetric1Present = gbyBitMask7 '10000000
  eMetric2Present = gbyBitMask6 '01000000
  eMetric3Present = gbyBitMask5 '00100000
  eMetric4Present = gbyBitMask4 '00010000
  eMetric5Present = gbyBitMask3 '00001000
  eMetric6Present = gbyBitMask2 '00000100
  eString1Present = gbyBitMask1 '00000010
End Enum
Public Enum ARM_Metrics_ArrayIndex_Enum
  eMetricInvalidIndex = -1      'SCR 28568 MSolano 10/16/02
  eMetric1Index = 1
  eMetric2Index = 2
  eMetric3Index = 3
  eMetric4Index = 4
  eMetric5Index = 5
  eMetric6Index = 6
  eMetricStringIndex = 7
End Enum



'=====================================================================================
'                   ARM DATATYPES
'=====================================================================================

'             --- ARM DATATYPES ENUM ID'S ---

'--- ARMDataType4MetricsFld_Enum
'--- rem'd values are valid ARM Data Types that may
'--- be supported in future versions of CHCSII ARM
Public Enum ARMDataType4MetricsFld_Enum
  eARM_InvalidType = -1
  'eARM_Counter64 = 2
  'eARM_CntrDivr32 = 3
  'eARM_Gauge32 = 4
  'eARM_Gauge64 = 5
  'eARM_GaugeDivr32 = 6
  eARM_NumericID32 = 7
  'eARM_NumericID64 = 8
  eARM_String8 = 9
  '--- special CHCSII ARM datatypes
  eARM_Date = 11
  eARM_Counter32_AutoIncrement = 12
End Enum
Public Enum ARMDataType4StringFld_Enum
  eARM_String32Only = 10
End Enum


'             --- ARM DATATYPES ENUM STRUCTS ---

'--- ARM_COUNTER: doesn't decreases except when
'--- counter increases to lCounterMaxVal (resets to 0)
'--- compare ARM_GAUGE
Public Type ARM_Counter32_Type
  lCounterMaxVal          As Long
  lUnused                 As Long
End Type

Public Type ARM_Counter64_Type
  byCounterMaxVal(64)      As Byte
End Type

'--- representation for decimal val's
'--- using floating point (Single, Double)
Public Type ARM_CntrDivr32_Type
  lCounterNumerator       As Long
  lDivisorDenominator     As Long
End Type

'--- ARM_GAUGE: Can increase or decrease
'--- compare ARM_COUNTER
Public Type ARM_Gauge32_Type
  lGuage                  As Long
  lUnused                 As Long
End Type

Public Type ARM_Gauge64_Type
  byGuage(64)              As Byte
End Type

'--- representation for decimal val's
'--- using floating point (Single, Double)
Public Type ARM_GaugeDivr32_Type
  lGuageNumerator           As Long
  lDivisorDenominator       As Long
End Type

'--- ARM_NumericID: math operations are meaningless
Public Type ARM_NumericID32_Type
  lID                       As Long
  lUnused                   As Long
End Type

Public Type ARM_NumericID64_Type
  byNumericID(64)            As Byte
End Type

Public Type ARM_String8_Type
  'sNzString8                As String * 8           'sNz = STRING *NOT* NULL TERMINATED; PAD W/ BLANKS
  bysNzString8(1 To 8)           As Byte               'sNz = STRING *NOT* NULL TERMINATED; PAD W/ BLANKS
End Type

Public Type ARM_String32_Type
  'sNzString32               As String * 32         'sNz = STRING *NOT* NULL TERMINATED; PAD W/ BLANKS
  bysNzString32(1 To 32)           As Byte          'sNz = STRING *NOT* NULL TERMINATED; PAD W/ BLANKS
End Type
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



'=====================================================================================
'       COMPLEX TRANSACTIONS: CLASS ENUMS & NAMES
'
' MODIFICATIONS REQUIRED FOR EACH NEW COMPLEX TRANSACTION CLASS
'
' cARMAPITypes Declarations section:
'   - Add an enum
'
' CmplxTxClassName
'   - add local const
'   - map enum to const in SELECT statement
'=====================================================================================
'Public Enum CmplxTx_PolymorphFunct_Enum  '
'  eCmplxTxCls_PolymorphFunct_Invalid
'  eCmplxTxCls_PolymorphFunct_Open
'  eCmplxTxCls_PolymorphFunct_Close
'End Enum

'Public Enum CmplxTxCls_Enum
'  eCmplxTxCls_SRTSOM_OpenOMFromCore = NCID_SRTS_OM_APPLICATION 'SCR 24431 RWalker 9/13/2002
'  eCmplxTxCls_SRTSOE_OpenOEFromCore = NCID_SRTS_OE_APPLICATION 'SCR 24821 RWalker 9/13/2002

'  eCmplxTxCls_SRTSOM_OpenOE                 'SCR 24821 RWalker 9/12/2002
'  eCmplxTxCls_SRTSOM_OpenAdmin              'SCR 24434 Rwalker  9/13/2002
  
'  eCmplxTxCls_SRTSOE_OpenAdmin              'SCR 24825 RWalker 10/2/2002
  
'  eCmplxTxCls_ARMFlagMetricCoerced_Value
'  eCmplxTxCls_ARMFlagMetricCoerced_DataType
'  eCmplxTxCls_EncounterID

'  eCmplxTxCls_SRTSOM_SearchByLabClinicDate  'SCR 24823 RWalker 9/12/2002
'  eCmplxTxCls_SRTSOM_SearchByOrderNo        'SCR 24432 MSolano 9/4/2002
'  eCmplxTxCls_SRTSOM_SearchBySSN            'SCR 24822 MSolano 9/4/2002
'  eCmplxTxCls_SRTSOM_ViewDetail             'SCR 24831 RWalker 9/12/2002
  
'  eCmplxTxCls_Test1
'  eCmplxTxCls_Test2
'  eCmplxTxCls_Test3
'  eCmplxTxCls_Test4
'  eCmplxTxCls_Test5
'  eCmplxTxCls_Test6
'  eCmplxTxCls_Test7
'  eCmplxTxCls_Test8
'  eCmplxTxCls_Test9
'  eCmplxTxCls_Test10
'  eCmplxTxCls_Test11
'  eCmplxTxCls_Test12
'End Enum

'Const msCmplxTxName_Unknown                       As String = "UNKNOWN TX CLASS"
'Const msCmplxTxName_ARMFlagMetricCoerced_Value    As String = "ARM FLAG - Complex Tx: Coerced Metric value"
'Const msCmplxTxName_ARMFlagMetricCoerced_DataType As String = "ARM FLAG - Complex Tx: Coerced Metric value and data type"
'Const msCmplxTxName_SRTSOM_OpenAdmin              As String = "SRTS OM Open SRTS Admin"                     'SCR 24434 MSolano 9/13/2002
'Const msCmplxTxName_SRTSOM_OpenOE                 As String = "SRTS OM Open SRTS OE"                        'SCR 244433 MSolano 9/12/2002
'Const msCmplxTxName_SRTSOM_OpenOM                 As String = "SRTS OM Open SRTS OM"                        'SCR 24431 MSolano 9/13/2002
'Const msCmplxTxName_EncID                         As String = "EncounterID"                                 'SCR 26651 MSolano 8/15/2002
'Const msCmplxTxName_SRTSOM_SearchByLabClinicDate  As String = "SRTS OM Search by Lab, Clinic, [Date Range]" 'SCR 24823 MSolano 9/4/2002
'Const msCmplxTxName_SRTSOM_SearchByOrderNo        As String = "SRTS OM Search by Order Number"              'SCR 24432 MSolano 9/4/2002
'Const msCmplxTxName_SRTSOM_SearchBySSN            As String = "SRTS OM Search by SNN"                       'SCR 24432 MSolano 9/4/2002
'Const msCmplxTxName_SRTSOM_ViewDetail             As String = "SRTS OM View Order Detail"                   'SCR 24831 MSolano 9/12/2002


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



'Friend Function CmplxTxClassName( _
                         eCmplxTxCls As CmplxTxCls_Enum, _
                Optional ePolymorphicAppFunct As CmplxTx_PolymorphFunct_Enum = eCmplxTxCls_PolymorphFunct_Invalid _
                ) As String
'Friend Function CmplxTxClassName( _
'                         eCmplxTxCls As CmplxTxCls_Enum _
'                ) As String
'
'    Dim sRetVal                                      As String
'
'    '------------------------------------------------------------------------------------------------------------------------------------------------
'    ' MODIFICATIONS REQUIRED FOR EACH NEW COMPLEX TRANSACTION CLASS
'    '
'    ' cARMAPITypes Declarations section:
'    '   - Add an enum
'    '
'    ' CmplxTxClassName
'    '   - add local const
'    '   - map enum to const in SELECT statement
'    '------------------------------------------------------------------------------------------------------------------------------------------------
'    Const sCmplxTxName_Unknown                       As String = "UNKNOWN TX CLASS"
'    Const sCmplxTxName_ARMFlagMetricCoerced_Value    As String = "ARM FLAG - Complex Tx: Coerced Metric value"
'    Const sCmplxTxName_ARMFlagMetricCoerced_DataType As String = "ARM FLAG - Complex Tx: Coerced Metric value and data type"
'    Const sCmplxTxName_SRTSOM_OpenAdmin              As String = "SRTS OM Open SRTS Admin"                     'SCR 24434 MSolano 9/13/2002
'    Const sCmplxTxName_SRTSOM_OpenOE                 As String = "SRTS OM Open SRTS OE"                        'SCR 244433 MSolano 9/12/2002
'    Const sCmplxTxName_SRTSOM_OpenOM                 As String = "SRTS OM Open SRTS OM"                        'SCR 24431 MSolano 9/13/2002
'    Const sCmplxTxName_EncID                         As String = "EncounterID"                                 'SCR 26651 MSolano 8/15/2002
'    Const sCmplxTxName_SRTSOM_SearchByLabClinicDate  As String = "SRTS OM Search by Lab, Clinic, [Date Range]" 'SCR 24823 MSolano 9/4/2002
'    Const sCmplxTxName_SRTSOM_SearchByOrderNo        As String = "SRTS OM Search by Order Number"              'SCR 24432 MSolano 9/4/2002
'    Const sCmplxTxName_SRTSOM_SearchBySSN            As String = "SRTS OM Search by SNN"                       'SCR 24432 MSolano 9/4/2002
'    Const sCmplxTxName_SRTSOM_ViewDetail             As String = "SRTS OM View Order Detail"                   'SCR 24831 MSolano 9/12/2002
'    Const sCmplxTxName_SRTSOE_OpenOE                 As String = "SRTS OE Open SRTS OE"                        'SCR 24821 MSolano 9/20/2002
'
'    On Error GoTo ErrorHandler
'
'    Select Case eCmplxTxCls
'
'      Case eCmplxTxCls_SRTSOM_OpenOMFromCore
'      'If ePolymorphicAppFunct = eCmplxTxCls_PolymorphFunct_Open Then
'      '  sRetVal = sCmplxTxName_SRTSOM_OpenOM
'      'End If
'
'      Case eCmplxTxCls_SRTSOE_OpenOEFromCore
'      'If ePolymorphicAppFunct = eCmplxTxCls_PolymorphFunct_Open Then
'      '  sRetVal = sCmplxTxName_SRTSOE_OpenOE
'      'End If
'
'      Case eCmplxTxCls_ARMFlagMetricCoerced_Value
'      sRetVal = sCmplxTxName_ARMFlagMetricCoerced_Value
'
'      Case eCmplxTxCls_ARMFlagMetricCoerced_DataType
'      sRetVal = sCmplxTxName_ARMFlagMetricCoerced_DataType
'
'      Case eCmplxTxCls_EncounterID
'      sRetVal = sCmplxTxName_EncID
'
'      Case eCmplxTxCls_SRTSOM_OpenAdmin
'      sRetVal = sCmplxTxName_SRTSOM_OpenAdmin
'
'      Case eCmplxTxCls_SRTSOM_OpenOE
'      sRetVal = sCmplxTxName_SRTSOM_OpenOE
'
'
'      Case eCmplxTxCls_SRTSOM_SearchByLabClinicDate
'      sRetVal = sCmplxTxName_SRTSOM_SearchByLabClinicDate
'
'      Case eCmplxTxCls_SRTSOM_SearchByOrderNo
'      sRetVal = sCmplxTxName_SRTSOM_SearchByOrderNo
'
'      Case eCmplxTxCls_SRTSOM_SearchBySSN
'      sRetVal = sCmplxTxName_SRTSOM_SearchBySSN
'
'      Case eCmplxTxCls_SRTSOM_ViewDetail
'      sRetVal = sCmplxTxName_SRTSOM_ViewDetail
'
'
'
'
'      Case Else
'      'sRetVal = sCmplxTxName_Unknown
'      sRetVal = "TestTxClsName_" & CStr(eCmplxTxCls)
'
'    End Select
'
'    CmplxTxClassName = sRetVal
'    Exit Function
'ErrorHandler:
'  CmplxTxClassName = sCmplxTxName_Unknown
'  Exit Function
'End Function



