VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCmplxTxCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=====================================================================================
'       cCmplxTxCls -  COMPLEX TRANSACTIONS CLASS

' PURPOSE:
'   To abstract construction of the data buffer in the
'   ARM API 's arm_getid() function
'
' COMPLEX TRANSACTION CLASSES
'   A term used in the ARM.dll adapter to the API to describe
'   transaction classes that involve either
'     - correlation
'     - user defined metrics
'=====================================================================================

Option Explicit




Public Type ARM_GetID_Metric_Type
  eMetricType               As ARMDataType4MetricsFld_Enum
  'szMetricName              As String * glARM_GetID_MetricNameBufSize 'sz = NULL TERMINATED
  byszMetricName(1 To glARM_GetID_MetricNameBufSize) As Byte            'sz = NULL TERMINATED
End Type

Public Type ARM_GetID_String_Type
  eMetricType             As ARMDataType4StringFld_Enum
  'szMetricName              As String * glARM_GetID_MetricNameBufSize  '*IS* NULL TERM
  byszMetricName(1 To glARM_GetID_MetricNameBufSize) As Byte             '*IS* NULL TERM
End Type


Public Enum ARM_GetID_Format_Enum
  eGetIDFormat = 101
End Enum





'=====================================================================================
'                   ARM_GetID DATA BUFFER STRUCT
'=====================================================================================

Public Type ARM_GetID_DataBuffer_Type
  l4Format                                        As Long 'const = 101
  by4Flags(1 To 4)                                As Byte 'byte 1, 3,4 = 0; byte 2 abcdefg0:
                                                          ' bits 2a-2f set if udtMetric(1-6) present;
                                                          ' bit 2g set if udtString present
  'udtMetric(1 To 6)                              As ARM_GetID_Metric_Type
  'udtString                                      As ARM_GetID_String_Type
  Metric1Type                                     As Long
  Metric1Name(1 To glARM_GetID_MetricNameBufSize) As Byte
  Metric2Type                                     As Long
  Metric2Name(1 To glARM_GetID_MetricNameBufSize) As Byte
  Metric3Type                                     As Long
  Metric3Name(1 To glARM_GetID_MetricNameBufSize) As Byte
  Metric4Type                                     As Long
  Metric4Name(1 To glARM_GetID_MetricNameBufSize) As Byte
  Metric5Type                                     As Long
  Metric5Name(1 To glARM_GetID_MetricNameBufSize) As Byte
  Metric6Type                                     As Long
  Metric6Name(1 To glARM_GetID_MetricNameBufSize) As Byte
  String1Type                                     As Long
  String1Name(1 To glARM_GetID_MetricNameBufSize) As Byte
End Type
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Private moMetricsProfile(eMetric1Index To eMetricStringIndex) As New cARMMetricsProfile

'Private Type MetricProfile_Type
'  bMetricPresent                As Boolean
'  eMetricType                   As ARMDataType4MetricsFld_Enum
'End Type
'Private mMetricsProfile(1 To 6) As MetricProfile_Type
'Private mbStringFieldPresent    As Boolean

Private mudtCmplxTxClsBuf       As ARM_GetID_DataBuffer_Type







'10/3/2002 Private meCmplx_Tx_Cls    As CmplxTxCls_Enum


Private msTxClassName                 As String
Private mlTxClassID                   As Long
Private mlTxClassAutoCount            As Long 'SCR 28568 MSolano 10/16/02
Private meTxClassAutoCountMetricPos   As ARM_Metrics_ArrayIndex_Enum 'SCR 28568 MSolano 10/16/02

'=====================================================================================
'       TRANSACTION FIELD TITLES FOR USER DEFINED METRICS IN COMPLEX TRANSACTIONS
'
'--- DO NOT USE EMBEDDED BLANKS IN FIELD TITLES
'--- NOTE: ARM API SAYS THAT THE STRING#1 FIELD IN THE CLASS
'--- DATA BUFFER IS NOT NULL TERMINATED
'--- BUT IS TERMINATED WITH A BLANK (ASCII 32)
'=====================================================================================

Private Const msMetricNotUsed                       As String = "MetricNotUsed"
Private Const msMetricInvalidString                 As String = "BadStr"
Private Const msMetricInvalidNumber                 As Long = 0

'Private mdictCorrelatorParentTxClasses              As Scripting.Dictionary 'SCR 28779 MSOLANO 11/4/02
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
















'<<< SCR 28568 MSolano 10/16/02
Public Property Let AutoCount(lVal As Long)
  On Error GoTo ErrorHandler
  
  mlTxClassAutoCount = lVal
  
  Exit Property
ErrorHandler:
  mlTxClassAutoCount = 0
  AutoCount = 0
  Call ARMDebugAssert(Err.Number = 0)    'stop in IDE
  Exit Property
End Property
Public Property Get AutoCount() As Long
  On Error GoTo ErrorHandler
  
  AutoCount = mlTxClassAutoCount
  
  Exit Property
ErrorHandler:
  AutoCount = glIGNORE_ME
  Call ARMDebugAssert(Err.Number = 0)    'stop in IDE
  Exit Property
End Property
Public Property Get AutoCountMetricPosition() As ARM_Metrics_ArrayIndex_Enum
  On Error GoTo ErrorHandler
  
  AutoCountMetricPosition = meTxClassAutoCountMetricPos
  
  Exit Property
ErrorHandler:
  AutoCountMetricPosition = eMetricInvalidIndex
  Call ARMDebugAssert(Err.Number = 0)    'stop in IDE
  Exit Property
End Property
'>>> SCR 28568 MSolano 10/16/02

Public Function DataBuffer( _
                   ByVal sReadableTxClsName As String, _
                Optional sString31FieldName As String = vbNullString, _
                Optional sMetric1Of6Name As String = vbNullString, _
                Optional eMetric1Of6Type As ARMDataType4MetricsFld_Enum = eARM_InvalidType, _
                Optional sMetric2Of6Name As String = vbNullString, _
                Optional eMetric2Of6Type As ARMDataType4MetricsFld_Enum = eARM_InvalidType, _
                Optional sMetric3Of6Name As String = vbNullString, _
                Optional eMetric3Of6Type As ARMDataType4MetricsFld_Enum = eARM_InvalidType, _
                Optional sMetric4Of6Name As String = vbNullString, _
                Optional eMetric4Of6Type As ARMDataType4MetricsFld_Enum = eARM_InvalidType, _
                Optional sMetric5Of6Name As String = vbNullString, _
                Optional eMetric5Of6Type As ARMDataType4MetricsFld_Enum = eARM_InvalidType, _
                Optional sMetric6Of6Name As String = vbNullString, _
                Optional eMetric6Of6Type As ARMDataType4MetricsFld_Enum = eARM_InvalidType _
                         ) As ARM_GetID_DataBuffer_Type
  
  Dim bContinue     As Boolean
  On Error GoTo ErrorHandler

  bContinue = Initialize( _
                  sReadableTxClsName, _
                  sString31FieldName, _
                  sMetric1Of6Name, _
                  eMetric1Of6Type, _
                  sMetric2Of6Name, _
                  eMetric2Of6Type, _
                  sMetric3Of6Name, _
                  eMetric3Of6Type, _
                  sMetric4Of6Name, _
                  eMetric4Of6Type, _
                  sMetric5Of6Name, _
                  eMetric5Of6Type, _
                  sMetric6Of6Name, _
                  eMetric6Of6Type)

    Call ARMDebugAssert(bContinue)
    '--- ? if it fails, return the udt anyway
    '--- BMTODO BMTEST: what will API do?
    DataBuffer = mudtCmplxTxClsBuf
    
    Exit Function
ErrorHandler:
   'stop in IDE
  Exit Function

End Function


Public Property Get DataBufferLength() As Long
  DataBufferLength = LenB(mudtCmplxTxClsBuf)
End Property
Public Property Get TxClassName() As String
  On Error GoTo ErrorHandler
  
  TxClassName = msTxClassName
  
  Exit Property
ErrorHandler:
  TxClassName = vbNullString
  Exit Property
End Property
Public Property Get TxClassID() As Long

  On Error GoTo ErrorHandler
  
  TxClassID = mlTxClassID
  
  Exit Property
ErrorHandler:
  TxClassID = glIGNORE_TX
  Exit Property
End Property
Public Property Let TxClassID(lData As Long)
  On Error GoTo ErrorHandler
  
  'Dbug.Assert False
  
  If lData > 0 Then
    mlTxClassID = lData
  Else
    mlTxClassID = glIGNORE_TX
  End If
  
  Exit Property
ErrorHandler:
  TxClassID = glIGNORE_TX
  Exit Property
End Property

Public Function Initialize( _
                     ByVal sReadableTxClsName As String, _
            Optional ByVal sString31FieldName As String = vbNullString, _
            Optional ByVal sMetric1Of6Name As String = vbNullString, _
            Optional ByVal eMetric1Of6Type As ARMDataType4MetricsFld_Enum = eARM_InvalidType, _
            Optional ByVal sMetric2Of6Name As String = vbNullString, _
            Optional ByVal eMetric2Of6Type As ARMDataType4MetricsFld_Enum = eARM_InvalidType, _
            Optional ByVal sMetric3Of6Name As String = vbNullString, _
            Optional ByVal eMetric3Of6Type As ARMDataType4MetricsFld_Enum = eARM_InvalidType, _
            Optional ByVal sMetric4Of6Name As String = vbNullString, _
            Optional ByVal eMetric4Of6Type As ARMDataType4MetricsFld_Enum = eARM_InvalidType, _
            Optional ByVal sMetric5Of6Name As String = vbNullString, _
            Optional ByVal eMetric5Of6Type As ARMDataType4MetricsFld_Enum = eARM_InvalidType, _
            Optional ByVal sMetric6Of6Name As String = vbNullString, _
            Optional ByVal eMetric6Of6Type As ARMDataType4MetricsFld_Enum = eARM_InvalidType _
                     ) As Boolean
  
  Dim bContinue         As Boolean
  Static bDoneThis      As Boolean
  On Error GoTo ErrorHandler
    
  If bDoneThis Then
    '---
    Call ARMDebugAssert(bDoneThis = False)
  Else
    bDoneThis = True
  End If
  
  Initialize = False
  msTxClassName = sReadableTxClsName
  
  bContinue = SetAutoCountMetricPosition( _
                  eMetric1Of6Type, _
                  eMetric2Of6Type, _
                  eMetric3Of6Type, _
                  eMetric4Of6Type, _
                  eMetric5Of6Type, _
                  eMetric6Of6Type)
  Call ARMDebugAssert(bContinue)
  If Not bContinue Then Exit Function

  bContinue = SetARMAPIBuffer( _
                  sString31FieldName, _
                  sMetric1Of6Name, _
                  eMetric1Of6Type, _
                  sMetric2Of6Name, _
                  eMetric2Of6Type, _
                  sMetric3Of6Name, _
                  eMetric3Of6Type, _
                  sMetric4Of6Name, _
                  eMetric4Of6Type, _
                  sMetric5Of6Name, _
                  eMetric5Of6Type, _
                  sMetric6Of6Name, _
                  eMetric6Of6Type)
  Call ARMDebugAssert(bContinue)
  If Not bContinue Then Exit Function
    
  bContinue = SetMetricsProfile( _
                  sString31FieldName, _
                  eMetric1Of6Type, _
                  eMetric2Of6Type, _
                  eMetric3Of6Type, _
                  eMetric4Of6Type, _
                  eMetric5Of6Type, _
                  eMetric6Of6Type)
  Call ARMDebugAssert(bContinue)
  If Not bContinue Then Exit Function
 
  
  Initialize = True
  Exit Function
ErrorHandler:
  Initialize = False
  Call ARMDebugAssert(Initialize)  'stop in IDE
  Exit Function

End Function


Public Function ValidateOrCoerceMetricValByExpectType( _
                           ByVal eWhichMetric As ARM_Metrics_ArrayIndex_Enum, _
                  Optional ByVal vInVal As Variant _
                  ) As Variant
  Dim vRetVal As Variant

  On Error GoTo ErrorHandler
  
  'BM4/
  If moMetricsProfile(eWhichMetric).DataType = eARM_Counter32_AutoIncrement Then
    '--- Tx class stores auto count of it's instances
    vRetVal = mlTxClassAutoCount
  Else
    '--- cARMMetricsProfile.Value will coerce input
    '--- data according to expected data type
    vRetVal = moMetricsProfile(eWhichMetric).ValidateOrCoerceMetricValByExpectType(vInVal)
  End If
  
  ValidateOrCoerceMetricValByExpectType = vRetVal
  Exit Function
ErrorHandler:
  Call ARMDebugAssert(Err.Number = 0)
  ValidateOrCoerceMetricValByExpectType = vInVal
End Function

'Public Function ValidateOrCoerceAllMetricValsByExpectType( _
'                  Optional ByRef vInOutStr31 As Variant, _
'                  Optional ByRef vInOutMetric1 As Variant, _
'                  Optional ByRef vInOutMetric2 As Variant, _
'                  Optional ByRef vInOutMetric3 As Variant, _
'                  Optional ByRef vInOutMetric4 As Variant, _
'                  Optional ByRef vInOutMetric5 As Variant, _
'                  Optional ByRef vInOutMetric6 As Variant _
'                  ) As Boolean
'
'  On Error GoTo ErrorHandler
'  ValidateOrCoerceAllMetricValsByExpectType = False
'
'  '
'  ' BMToDoTest TEST  TEST  TEST  TEST  TEST  TEST  TEST  TEST
'  '
'  ' send in a string argument when a long expected
'  ' and confirm coercion
'  '
'  ' TEST  TEST  TEST  TEST  TEST  TEST  TEST  TEST
'
'  '--- cARMMetricsProfile.Value will coerce input
'  '--- data according to expected data type
'  moMetricsProfile(eMetric1Index).MetricValue = vInOutMetric1
'  moMetricsProfile(eMetric2Index).MetricValue = vInOutMetric2
'  moMetricsProfile(eMetric3Index).MetricValue = vInOutMetric3
'  moMetricsProfile(eMetric4Index).MetricValue = vInOutMetric4
'  moMetricsProfile(eMetric5Index).MetricValue = vInOutMetric5
'  moMetricsProfile(eMetric6Index).MetricValue = vInOutMetric6
'
'  moMetricsProfile(eMetricStringIndex).MetricValue = vInOutStr31
'
'  ValidateOrCoerceAllMetricValsByExpectType = True
'  Exit Function
'ErrorHandler:
'  ValidateOrCoerceAllMetricValsByExpectType = False
'  Call ARMDebugAssert(ValidateOrCoerceAllMetricValsByExpectType)
'End Function



Private Function SetMetricsProfile( _
                  Optional ByVal sString31FieldName As String = vbNullString, _
                  Optional ByVal eMetric1Of6Type As ARMDataType4MetricsFld_Enum = eARM_InvalidType, _
                  Optional ByVal eMetric2Of6Type As ARMDataType4MetricsFld_Enum = eARM_InvalidType, _
                  Optional ByVal eMetric3Of6Type As ARMDataType4MetricsFld_Enum = eARM_InvalidType, _
                  Optional ByVal eMetric4Of6Type As ARMDataType4MetricsFld_Enum = eARM_InvalidType, _
                  Optional ByVal eMetric5Of6Type As ARMDataType4MetricsFld_Enum = eARM_InvalidType, _
                  Optional ByVal eMetric6Of6Type As ARMDataType4MetricsFld_Enum = eARM_InvalidType _
                  ) As Boolean
  Dim i As ARM_Metrics_ArrayIndex_Enum

  On Error GoTo ErrorHandler
  SetMetricsProfile = True
  
  'dbug.Assert False
  
  moMetricsProfile(eMetric1Index).DataType = eMetric1Of6Type
  moMetricsProfile(eMetric2Index).DataType = eMetric2Of6Type
  moMetricsProfile(eMetric3Index).DataType = eMetric3Of6Type
  moMetricsProfile(eMetric4Index).DataType = eMetric4Of6Type
  moMetricsProfile(eMetric5Index).DataType = eMetric5Of6Type
  moMetricsProfile(eMetric6Index).DataType = eMetric6Of6Type
  
  If sString31FieldName = vbNullString Then
    moMetricsProfile(eMetricStringIndex).DataType = eARM_InvalidType
  Else
    moMetricsProfile(eMetricStringIndex).DataType = eARM_String32Only
  End If
  
  Exit Function
ErrorHandler:
  SetMetricsProfile = False
  Call ARMDebugAssert(SetMetricsProfile)

End Function
'<<< SCR 28568 MSolano 10/16/02
Private Function SetAutoCountMetricPosition( _
                  Optional ByVal eMetric1Of6Type As ARMDataType4MetricsFld_Enum = eARM_InvalidType, _
                  Optional ByRef eInOutMetric2Of6Type As ARMDataType4MetricsFld_Enum = eARM_InvalidType, _
                  Optional ByRef eInOutMetric3Of6Type As ARMDataType4MetricsFld_Enum = eARM_InvalidType, _
                  Optional ByRef eInOutMetric4Of6Type As ARMDataType4MetricsFld_Enum = eARM_InvalidType, _
                  Optional ByRef eInOutMetric5Of6Type As ARMDataType4MetricsFld_Enum = eARM_InvalidType, _
                  Optional ByRef eInOutMetric6Of6Type As ARMDataType4MetricsFld_Enum = eARM_InvalidType _
                  ) As Boolean
  
  On Error GoTo ErrorHandler
  SetAutoCountMetricPosition = False
  
  '--- IF MORE THAN ONE METRIC DEFINED AS AUTOCOUNT,
  '--- FIRST ONE WINS THE SPOT
  '--- OTHERS COERCED TO INVALID DATA TYPE
  'call ARMDebugAssert (False)
  
  If eMetric1Of6Type = eARM_Counter32_AutoIncrement Then
    meTxClassAutoCountMetricPos = eMetric1Index
  End If
  
  If eInOutMetric2Of6Type = eARM_Counter32_AutoIncrement Then
    If meTxClassAutoCountMetricPos = eMetricInvalidIndex Then
      meTxClassAutoCountMetricPos = eMetric2Index
    Else 'you tried, but other metric position got there first; coerce data type so this metric's ignored
      eInOutMetric2Of6Type = eARM_InvalidType
    End If
  End If
  
  If eInOutMetric3Of6Type = eARM_Counter32_AutoIncrement Then
    If meTxClassAutoCountMetricPos = eMetricInvalidIndex Then
      meTxClassAutoCountMetricPos = eMetric3Index
    Else 'you tried, but other metric position got there first; coerce data type so this metric's ignored
      eInOutMetric3Of6Type = eARM_InvalidType
    End If
  End If
  
  If eInOutMetric4Of6Type = eARM_Counter32_AutoIncrement Then
    If meTxClassAutoCountMetricPos = eMetricInvalidIndex Then
      meTxClassAutoCountMetricPos = eMetric4Index
    Else 'you tried, but other metric position got there first; coerce data type so this metric's ignored
      eInOutMetric4Of6Type = eARM_InvalidType
    End If
  End If
  
  If eInOutMetric5Of6Type = eARM_Counter32_AutoIncrement Then
    If meTxClassAutoCountMetricPos = eMetricInvalidIndex Then
      meTxClassAutoCountMetricPos = eMetric5Index
    Else 'you tried, but other metric position got there first; coerce data type so this metric's ignored
      eInOutMetric5Of6Type = eARM_InvalidType
    End If
  End If
  
  If eInOutMetric6Of6Type = eARM_Counter32_AutoIncrement Then
    If meTxClassAutoCountMetricPos = eMetricInvalidIndex Then
      meTxClassAutoCountMetricPos = eMetric6Index
    Else 'you tried, but other metric position got there first; coerce data type so this metric's ignored
      eInOutMetric6Of6Type = eARM_InvalidType
    End If
  End If
 
  
  SetAutoCountMetricPosition = True
  Exit Function
ErrorHandler:
  meTxClassAutoCountMetricPos = False
  Call ARMDebugAssert(Err.Number = 0)    'stop in IDE
  Exit Function

End Function
'>>> SCR 28568










Private Sub Class_Initialize()
  msTxClassName = vbNullString
  mlTxClassID = glIGNORE_TX
  mudtCmplxTxClsBuf.l4Format = eGetIDFormat   'CONST
  
  mlTxClassAutoCount = 0
  meTxClassAutoCountMetricPos = eMetricInvalidIndex
End Sub

'--- set which data field is present
'--- (none to 6 metric fields, or string field)
'--- set string field's name
Private Function SetStringFldInBuf( _
                      ByVal sMetricName As String _
                      ) As Boolean
  Dim bContinue       As Boolean
  
  On Error GoTo ErrorHandler
  SetStringFldInBuf = False
 
  'dbug.Assert False
  
  '--- SetMetricsFlags() REPLACED BY CALL TO
  '--- TAPMBugAllMetrics_SetFlags IN InitUserDefMetrics_EncounterID
  '--- AS (TEMP?) WORKAROUND FOR TAPM ARM BUG
  '--- Set Bit Flag to note which Metric field is present
  'bContinue = SetMetricsFlags(eWhichMetricByByteFlag)
  'call ARMDebugAssert(bContinue)
  'If Not bContinue Then Exit Function
  
  bContinue = ANSIStringWSpecifiedTerminator( _
              mudtCmplxTxClsBuf.String1Name, _
              sMetricName, _
              gbyBYTE_CHAR_NULL)
    Call ARMDebugAssert(bContinue)
    If Not bContinue Then Exit Function
    
    '--- coerce ARM data type to the only
    '--- valid type for the string descriptor field
    mudtCmplxTxClsBuf.String1Type = eARM_String32Only
  
  SetStringFldInBuf = True
  Exit Function
ErrorHandler:
  SetStringFldInBuf = False
  Call ARMDebugAssert(SetStringFldInBuf)  'stop in IDE
  Exit Function

End Function

Private Function SetARMAPIBuffer( _
                  Optional ByVal sString31Field As String = vbNullString, _
                  Optional ByVal sMetric1Of6Name As String = vbNullString, _
                  Optional ByVal eMetric1Of6Type As ARMDataType4MetricsFld_Enum = eARM_InvalidType, _
                  Optional ByVal sMetric2Of6Name As String = vbNullString, _
                  Optional ByVal eMetric2Of6Type As ARMDataType4MetricsFld_Enum = eARM_InvalidType, _
                  Optional ByVal sMetric3Of6Name As String = vbNullString, _
                  Optional ByVal eMetric3Of6Type As ARMDataType4MetricsFld_Enum = eARM_InvalidType, _
                  Optional ByVal sMetric4Of6Name As String = vbNullString, _
                  Optional ByVal eMetric4Of6Type As ARMDataType4MetricsFld_Enum = eARM_InvalidType, _
                  Optional ByVal sMetric5Of6Name As String = vbNullString, _
                  Optional ByVal eMetric5Of6Type As ARMDataType4MetricsFld_Enum = eARM_InvalidType, _
                  Optional ByVal sMetric6Of6Name As String = vbNullString, _
                  Optional ByVal eMetric6Of6Type As ARMDataType4MetricsFld_Enum = eARM_InvalidType _
                  ) As Boolean

  Dim bContinue   As Boolean
  On Error GoTo ErrorHandler
  
  SetARMAPIBuffer = False

  bContinue = TAPMBugAllMetrics_SetFlags()
  Call ARMDebugAssert(bContinue)
  'If Not bContinue Then Exit Function
  
  If sString31Field = vbNullString Then
    sString31Field = msMetricNotUsed
  End If
  bContinue = SetStringFldInBuf(sString31Field)
  Call ARMDebugAssert(bContinue)
  
  '--- eARM_NumericID32 default in ARM API Buffer for missing metrics
  
  If sMetric1Of6Name = vbNullString _
  Or eMetric1Of6Type = eARM_InvalidType Then
    sMetric1Of6Name = msMetricNotUsed
    eMetric1Of6Type = eARM_NumericID32
  ElseIf eMetric1Of6Type = eARM_Date Then
    eMetric1Of6Type = eARM_NumericID32
  End If
  bContinue = SetMetricFldInBuf( _
              sMetric1Of6Name, _
              eMetric1Index, _
              eMetric1Of6Type)
  Call ARMDebugAssert(bContinue)
    
  If sMetric2Of6Name = vbNullString _
  Or eMetric2Of6Type = eARM_InvalidType Then
    sMetric2Of6Name = msMetricNotUsed
    eMetric2Of6Type = eARM_NumericID32
  ElseIf eMetric2Of6Type = eARM_Date Then
    eMetric2Of6Type = eARM_NumericID32
  End If
  bContinue = SetMetricFldInBuf( _
              sMetric2Of6Name, _
              eMetric2Index, _
              eMetric2Of6Type)
  Call ARMDebugAssert(bContinue)
  
  If sMetric3Of6Name = vbNullString _
  Or eMetric3Of6Type = eARM_InvalidType Then
    sMetric3Of6Name = msMetricNotUsed
    eMetric3Of6Type = eARM_NumericID32
  ElseIf eMetric3Of6Type = eARM_Date Then
    eMetric3Of6Type = eARM_NumericID32
  End If
  bContinue = SetMetricFldInBuf( _
              sMetric3Of6Name, _
              eMetric3Index, _
              eMetric3Of6Type)
  Call ARMDebugAssert(bContinue)
  
  If sMetric4Of6Name = vbNullString _
  Or eMetric4Of6Type = eARM_InvalidType Then
    sMetric4Of6Name = msMetricNotUsed
    eMetric4Of6Type = eARM_NumericID32
  ElseIf eMetric4Of6Type = eARM_Date Then
    eMetric4Of6Type = eARM_NumericID32
  End If
  bContinue = SetMetricFldInBuf( _
              sMetric4Of6Name, _
              eMetric4Index, _
              eMetric4Of6Type)
  Call ARMDebugAssert(bContinue)
    
  If sMetric5Of6Name = vbNullString _
  Or eMetric5Of6Type = eARM_InvalidType Then
    sMetric5Of6Name = msMetricNotUsed
    eMetric5Of6Type = eARM_NumericID32
  ElseIf eMetric5Of6Type = eARM_Date Then
    eMetric5Of6Type = eARM_NumericID32
  End If
  bContinue = SetMetricFldInBuf( _
              sMetric5Of6Name, _
              eMetric5Index, _
              eMetric5Of6Type)
  Call ARMDebugAssert(bContinue)
    
  If sMetric6Of6Name = vbNullString _
  Or eMetric6Of6Type = eARM_InvalidType Then
    sMetric6Of6Name = msMetricNotUsed
    eMetric6Of6Type = eARM_NumericID32
  ElseIf eMetric6Of6Type = eARM_Date Then
    eMetric6Of6Type = eARM_NumericID32
  End If
  bContinue = SetMetricFldInBuf( _
              sMetric6Of6Name, _
              eMetric6Index, _
              eMetric6Of6Type)
  Call ARMDebugAssert(bContinue)
    
  Call ARMDebugAssert(LenB(mudtCmplxTxClsBuf) = glARM_GetID_DataBuffer_Size)
    
  SetARMAPIBuffer = True

  Exit Function
ErrorHandler:
  SetARMAPIBuffer = False
  Call ARMDebugAssert(SetARMAPIBuffer)

End Function

Private Function SetMetricFldInBuf( _
                       ByVal sMetricName As String, _
                       ByVal eWhichMetricByArrayPos As ARM_Metrics_ArrayIndex_Enum, _
              Optional ByVal eMetricType As ARMDataType4MetricsFld_Enum = eARM_NumericID32 _
                       ) As Boolean
              
  Dim bContinue        As Boolean

  On Error GoTo ErrorHandler
  SetMetricFldInBuf = False
  
  'Dbug.Assert False
   
  '--- SetMetricsFlags() REPLACED BY CALL TO
  '--- TAPMBugAllMetrics_SetFlags IN InitUserDefMetrics_EncounterID
  '--- AS (TEMP?) WORKAROUND FOR TAPM ARM BUG
  '--- Set Bit Flag to note which Metric field is present
  'bContinue = SetMetricsFlags(eWhichMetricByByteFlag)
  'call ARMDebugAssert(bContinue)
  'If Not bContinue Then Exit Function

  '--- convert CHCSII ARM data type to ARM API datatype
  If eMetricType = eARM_Counter32_AutoIncrement Then
    eMetricType = eARM_NumericID32
  End If

  Select Case eWhichMetricByArrayPos
  
  Case eMetric1Index
    mudtCmplxTxClsBuf.Metric1Type = eMetricType
    bContinue = ANSIStringWSpecifiedTerminator( _
                    mudtCmplxTxClsBuf.Metric1Name(), _
                    sMetricName, _
                    gbyBYTE_CHAR_NULL)
                    
  Case eMetric2Index
    mudtCmplxTxClsBuf.Metric2Type = eMetricType
    bContinue = ANSIStringWSpecifiedTerminator( _
                    mudtCmplxTxClsBuf.Metric2Name(), _
                    sMetricName, _
                    gbyBYTE_CHAR_NULL)

  Case eMetric3Index
    mudtCmplxTxClsBuf.Metric3Type = eMetricType
    bContinue = ANSIStringWSpecifiedTerminator( _
                    mudtCmplxTxClsBuf.Metric3Name(), _
                    sMetricName, _
                    gbyBYTE_CHAR_NULL)

  Case eMetric4Index
    mudtCmplxTxClsBuf.Metric4Type = eMetricType
    bContinue = ANSIStringWSpecifiedTerminator( _
                    mudtCmplxTxClsBuf.Metric4Name(), _
                    sMetricName, _
                    gbyBYTE_CHAR_NULL)
                    
  Case eMetric5Index
    mudtCmplxTxClsBuf.Metric5Type = eMetricType
    bContinue = ANSIStringWSpecifiedTerminator( _
                    mudtCmplxTxClsBuf.Metric5Name(), _
                    sMetricName, _
                    gbyBYTE_CHAR_NULL)
                    
  Case eMetric6Index
    mudtCmplxTxClsBuf.Metric6Type = eMetricType
    bContinue = ANSIStringWSpecifiedTerminator( _
                    mudtCmplxTxClsBuf.Metric6Name(), _
                    sMetricName, _
                    gbyBYTE_CHAR_NULL)

  End Select

  SetMetricFldInBuf = True

  Exit Function
ErrorHandler:
  SetMetricFldInBuf = False
  Call ARMDebugAssert(SetMetricFldInBuf)
End Function



'--- "ALL OR NONE" TAPM ARM BUG regarding metrics - per Tivoli / Itally
'--- giampiero catasta giampiero.catasta@ it.ibm.com
'--- L3 Customer Support Tivoli Application Performance Management
'--- either all flags for metrics are set
'--- or none are set
'--- either all metrics are sent
'--- or none are sent
Private Function TAPMBugAllMetrics_SetFlags() As Boolean
  Dim byFlags         As Byte
  
  On Error GoTo ErrorHandler
  TAPMBugAllMetrics_SetFlags = False
  
  'dbug.Assert False
  
  '--- SET ALL BITS SHOWING ALL METRICS ARE PRESENT
  '--- ... LET TIVOLI FIGURE OUT WHICH ARE
  byFlags = byFlags Or gbyBitMask1
  byFlags = byFlags Or gbyBitMask2
  byFlags = byFlags Or gbyBitMask3
  byFlags = byFlags Or gbyBitMask4
  byFlags = byFlags Or gbyBitMask5
  byFlags = byFlags Or gbyBitMask6
  byFlags = byFlags Or gbyBitMask7
  
  mudtCmplxTxClsBuf.by4Flags(glARM_GetID_FlagsByte) = byFlags
    
  TAPMBugAllMetrics_SetFlags = True
  Exit Function
ErrorHandler:
  TAPMBugAllMetrics_SetFlags = False
  Call ARMDebugAssert(TAPMBugAllMetrics_SetFlags)
End Function

'Private Function SetTxClsName( _
'                    eCmplxTxCls As CmplxTxCls_Enum _
'                    ) As Boolean
'  On Error GoTo ErrorHandler
'  SetTxClsName = True  'assume optimistic; case else returns false
'
'  Select Case eCmplxTxCls
'
'    Case eCmplxTxCls_ARMFlagMetricCoerced_Value
'      msTxClassName = msCmplxTxName_ARMFlagMetricCoerced_Value
'
'    Case eCmplxTxCls_ARMFlagMetricCoerced_DataType
'      msTxClassName = msCmplxTxName_ARMFlagMetricCoerced_DataType
'
'    Case eCmplxTxCls_EncounterID
'      msTxClassName = msCmplxTxName_EncID
'
'    Case eCmplxTxCls_SRTSOM_OpenAdmin
'      msTxClassName = msCmplxTxName_SRTSOM_OpenAdmin
'
'    Case eCmplxTxCls_SRTSOM_OpenOE
'      msTxClassName = msCmplxTxName_SRTSOM_OpenOE
'
'    Case eCmplxTxCls_SRTSOM_OpenOM
'      msTxClassName = msCmplxTxName_SRTSOM_OpenOM
'
'    Case eCmplxTxCls_SRTSOM_SearchByLabClinicDate
'      msTxClassName = msCmplxTxName_SRTSOM_SearchByLabClinicDate
'
'    Case eCmplxTxCls_SRTSOM_SearchByOrderNo
'      msTxClassName = msCmplxTxName_SRTSOM_SearchByOrderNo
'
'    Case eCmplxTxCls_SRTSOM_SearchBySSN
'      msTxClassName = msCmplxTxName_SRTSOM_SearchBySSN
'
'    Case eCmplxTxCls_SRTSOM_ViewDetail
'      msTxClassName = msCmplxTxName_SRTSOM_ViewDetail
'
'    Case Else
'      msTxClassName = msCmplxTxName_Unknown
'      SetTxClsName = False
'      Call ARMDebugAssert(SetTxClsName)
'  End Select
'
'  Exit Function
'
'ErrorHandler:
'  SetTxClsName = False
'  Call ARMDebugAssert(SetTxClsName)
'  Exit Function
'
'
'End Function

Private Function ValidateStringInVariant( _
                      ByRef sRetVal As String, _
                      Optional ByVal vData31Char As Variant _
                      ) As Boolean
  
  Dim bContinue     As Boolean
  
  On Error GoTo ErrorHandler
  ValidateStringInVariant = False
  
  bContinue = Not (IsMissing(vData31Char))
  Call ARMDebugAssert(bContinue)
  If Not bContinue Then Exit Function
  
  bContinue = (VarType(vData31Char) = vbString)
  Call ARMDebugAssert(bContinue)
  If Not bContinue Then Exit Function
  
  sRetVal = CStr(vData31Char)
  
  If Not Len(sRetVal) > 0 Then sRetVal = gsNO_DATA
  
  ValidateStringInVariant = True
  Exit Function
ErrorHandler:
  ValidateStringInVariant = False
  Call ARMDebugAssert(ValidateStringInVariant)    'stop in IDE
  Exit Function

End Function



'--- set which data field is present
'--- (none to 6 metric fields, or string field)
'--- set data field's name
'--- set data field's data type
'Private Function SetDataField( _
'                       ByVal sMetricName As String, _
'                       ByVal eWhichMetricByByteFlag As ARM_Metrics_Flag_Enum, _
'                    Optional eWhichMetricByArrayPos As ARM_Metrics_ArrayIndex_Enum, _
'              Optional ByVal eMetricType As ARMDataType4MetricsFld_Enum = eARM_NumericID32 _
'              ) As Boolean
'  Dim bContinue       As Boolean
'
'  On Error GoTo ErrorHandler
'  SetDataField = False
'
'  'dbug.Assert False
'
'  '--- SetMetricsFlags() REPLACED BY CALL TO
'  '--- TAPMBugAllMetrics_SetFlags IN InitUserDefMetrics_EncounterID
'  '--- AS (TEMP?) WORKAROUND FOR "ALL OR NONE" TAPM ARM BUG
'  '--- Set Bit Flag to note which Metric field is present
'  'bContinue = SetMetricsFlags(eWhichMetricByByteFlag)
'  'call ARMDebugAssert(bContinue)
'  'If Not bContinue Then Exit Function
'
'  '--- Fill in the title of the field
'  '--- Both of these:
'  '---    - All 6 metrics fields
'  '---    - String description
'  '---  are terminated by a 1 byte representation (ANSI, not 2 byte Unicode)
'  '---  of one character
'  If eWhichMetricByByteFlag = eString1Present Then
'    ' --- String terminated by Space
'    bContinue = ANSIStringWSpecifiedTerminator( _
'                mudtCmplxTxClsBuf.String1Name, _
'                sMetricName, _
'                gbyBYTE_CHAR_NULL)
'    call ARMDebugAssert(bContinue)
'    If Not bContinue Then Exit Function
'    '--- coerce ARM data type to the only
'    '--- valid type for the string descriptor field
'    mudtCmplxTxClsBuf.String1Type = eARM_String32Only
'
'  Else    'One of 6 Metrics Fields
'
'    bContinue = SetMetricFldInBuf( _
'                    sMetricName, _
'                    eWhichMetricByArrayPos, _
'                    eMetricType)
'
'    call ARMDebugAssert(bContinue)
'    If Not bContinue Then Exit Function
'
'  End If
'
'  SetDataField = True
'  Exit Function
'ErrorHandler:
'  SetDataField = False
'  call ARMDebugAssert(SetDataField)  'stop in IDE
'  Exit Function
'
'End Function


'Private Function SetMetricsFlags( _
'                    ByVal eWhichFlag As ARM_Metrics_Flag_Enum _
'                    ) As Boolean
'  Dim byFlags        As Byte
'
'  On Error GoTo ErrorHandler
'  SetMetricsFlags = False
'
'  If eWhichFlag < 0 Or eWhichFlag > 255 Then
'    Exit Function
'  End If
'
'
'  '--- "ALL OR NONE" TAPM ARM BUG regarding metrics - per Tivoli / Itally
'  '--- giampiero catasta giampiero.catasta@ it.ibm.com
'  '--- L3 Customer Support Tivoli Application Performance Management
'  '--- either all flags for metrics are set
'  '--- or none are set
'  '--- either all metrics are sent
'  '--- or none are sent
'
'  '--- THIS IS THE THE WAY IT SHOULD BE ...
'  byFlags = mudtCmplxTxClsBuf.by4Flags(glARM_GetID_FlagsByte)
'  '--- set desired bit
'  mudtCmplxTxClsBuf.by4Flags(glARM_GetID_FlagsByte) = byFlags Or CByte(eWhichFlag)
'
'  SetMetricsFlags = True
'  Exit Function
'ErrorHandler:
'  SetMetricsFlags = False
'  call ARMDebugAssert(SetMetricsFlags    'stop in IDE
'  Exit Function
'End Function





'Private Function MetricName( _
'                      ByRef byInOutArray() As Byte, _
'                      ByVal sMetricName As String _
'                      ) As Boolean
''Private Function MetricName( _
'                      ByVal sMetricName As String _
'                      ) As Byte()
'  Dim byRet()     As Byte
'  Dim bContinue   As Boolean
'
'  On Error GoTo ErrorHandler
'  'ReDim byRet(1 To glARM_GetID_MetricNameBufSize)
'
'  '--- Other metrics terminated by Null char
'  bContinue = ANSIStringWSpecifiedTerminator( _
'                  byInOutArray(), _
'                  sMetricName, _
'                  gbyBYTE_CHAR_NULL)
'  call ARMDebugAssert(bContinue
'  If Not bContinue Then Exit Function
'
'  MetricName = byRet()
'  Exit Function
'ErrorHandler:
'  bContinue = False
'  call ARMDebugAssert(bContinue
'
'End Function







'--- set which data field is present
'--- (none to 6 metric fields, or string field)
'--- set data field's name
'--- set data field's data type
'Private Function SetDataField_20020819( _
'                       ByVal sMetricName As String, _
'                       ByVal eWhichMetricByByteFlag As ARM_Metrics_Flag_Enum, _
'                    Optional eWhichMetricByArrayPos As ARM_Metrics_ArrayIndex_Enum, _
'              Optional ByVal eMetricType As ARMDataType4MetricsFld_Enum = eARM_NumericID32 _
'              ) As Boolean
'  Dim bContinue       As Boolean
'
'  On Error GoTo ErrorHandler
'  SetDataField_20020819 = False
'
'
'
'  '--- Set Flag to note which Metric field is present
'  bContinue = SetMetricsFlags(eWhichMetricByByteFlag)
'  call ARMDebugAssert(bContinue
'  If Not bContinue Then Exit Function
'
'  '--- Fill in the title of the field
'  '--- Both of these:
'  '---    - All 6 metrics fields
'  '---    - String description
'  '---  are terminated by a 1 byte representation (ANSI, not 2 byte Unicode)
'  '---  of one character
'  If eWhichMetricByByteFlag = eString1Present Then
'    ' --- String terminated by Space
'    'bContinue = ANSIStringWSpecifiedTerminator( _
'                mudtCmplxTxClsBuf.udtString.byszMetricName, _
'                sMetricName, _
'                gbyBYTE_CHAR_NULL)
'    call ARMDebugAssert(bContinue
'    If Not bContinue Then Exit Function
'    '--- coerce ARM data type to the only
'    '--- valid type for the string descriptor field
'    'mudtCmplxTxClsBuf.udtString.eMetricType = eARM_String32Only
'
'  Else    'One of 6 Metrics Fields
'
'    'eWhichMetricByArrayPos MUST BE PRESENT
'    'bcontinue = (
'
'    '--- Other metrics terminated by Null char
'    'bContinue = ANSIStringWSpecifiedTerminator( _
'                mudtCmplxTxClsBuf.udtMetric(eWhichMetricByArrayPos).byszMetricName, _
'                sMetricName, _
'                gbyBYTE_CHAR_NULL)
'    call ARMDebugAssert(bContinue
'    If Not bContinue Then Exit Function
'    '--- Declare datatype (const if String1 present)
'    'mudtCmplxTxClsBuf.udtMetric(eWhichMetricByArrayPos).eMetricType = eMetricType
'
'  End If
'
'  SetDataField_20020819 = True
'  Exit Function
'ErrorHandler:
'  SetDataField_20020819 = False
'  call ARMDebugAssert(SetDataField_20020819  'stop in IDE
'  Exit Function
'
'End Function



