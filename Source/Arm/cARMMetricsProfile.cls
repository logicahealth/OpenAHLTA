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
Attribute VB_Name = "cARMMetricsProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'--- ARMDataTypeInProfile_Enum combines data type
'--- enums for metrics fields and String field - combination of:
'---    ARMDataType4MetricsFld_Enum
'---    ARMDataType4StringFld_Enum
Public Enum ARMDataTypeInProfile_Enum
  eProfileMetricsType_InvalidType = -1
  'eProfileMetricsType_Counter64 = 2
  'eProfileMetricsType_CntrDivr32 = 3
  'eProfileMetricsType_Gauge32 = 4
  'eProfileMetricsType_Gauge64 = 5
  'eProfileMetricsType_GaugeDivr32 = 6
  eProfileMetricsType_NumericID32 = 7
  'eProfileMetricsType_NumericID64 = 8
  eProfileMetricsType_String8 = 9
  eProfileMetricsType_String32 = 10
  '--- special CHCSII ARM datatypes
  eProfileMetricsType_Date = 11
  eProfileMetricsType_Counter32_AutoIncrement = 12  'SCR 28568 MSolano 10/17/02
End Enum


Private meMetricDataType              As ARMDataTypeInProfile_Enum
Private mvMetricDefaultValue          As Variant
Private mvMetricValue                 As Variant

Private mbMetricPassedIn_Coerced      As Boolean
Private mvMetricPassedIn_Value        As Variant
Private meMetricPassedIn_Type         As VbVarType
Private mbMetricPrevSet                As Boolean

Private Const msMetricInvalidString   As String = "MISSING"  '8 char or less
'--- has to be zero, the bytes in buffer default to zero
Private Const mlMetricInvalidNumber   As Long = 0


Private Sub Class_Initialize()
  meMetricDataType = eARM_InvalidType
  mvMetricValue = mlMetricInvalidNumber
  mbMetricPrevSet = False
End Sub
Private Function ValidateOrCoerce_DateToLong( _
                      ByRef vInMetricValue As Variant _
                          ) As Long
  Dim lDateAsLong           As Long
  Dim sDate                 As String
  Dim dDate                 As Date
  On Error GoTo ErrorHandler
  
  'BMTODO: CHANGE LOGIC TO STORE DATE IN STR8 FIELD
  
  If VarType(vInMetricValue) = vbDate Or IsDate(vInMetricValue) Then
    dDate = CDate(vInMetricValue)
    sDate = Format(dDate, "YYYYMMDD")
    ValidateOrCoerce_DateToLong = CLng(sDate)
  
  Else
  
    'Call ARMDebugAssert(mvMetricDefaultValue = mlMetricInvalidNumber)
    ValidateOrCoerce_DateToLong = mlMetricInvalidNumber
    
  End If
  
  Exit Function
ErrorHandler:
  'Call ARMDebugAssert(Err.Number = 0)     'stop in IDE
  ValidateOrCoerce_DateToLong = mvMetricDefaultValue
  mbMetricPassedIn_Coerced = True

End Function

Private Function ValidateOrCoerce_Long( _
                      ByVal vInVal As Variant _
                          ) As Long
  Dim lCoerceVariant2Long   As Long
  On Error GoTo SetDefaultsForLong
  
  If VarType(vInVal) = vbLong Then
    ValidateOrCoerce_Long = vInVal
  
  Else
    lCoerceVariant2Long = CLng(vInVal)
    ValidateOrCoerce_Long = lCoerceVariant2Long
  End If
  
  Exit Function
  
SetDefaultsForLong:
  'Call ARMDebugAssert(Err.Number = 0)     'stop in IDE
  ValidateOrCoerce_Long = mlMetricInvalidNumber
  mbMetricPassedIn_Coerced = True

End Function

Private Function ValidateOrCoerce_String( _
                      ByVal vInMetricValue As Variant _
                      ) As String
  On Error GoTo ErrorHandler
  
  If VarType(vInMetricValue) = vbString Then
    ValidateOrCoerce_String = vInMetricValue
  
  Else
    
    mbMetricPassedIn_Coerced = True
    'Call ARMDebugAssert((mbMetricPassedIn_Coerced = False))      'stop IDE on purpose
    ValidateOrCoerce_String = CStr(vInMetricValue)
  
  End If
    
  Exit Function
ErrorHandler:
  'Call ARMDebugAssert(Err.Number = 0)     'stop in IDE
  ValidateOrCoerce_String = msMetricInvalidString
  mbMetricPassedIn_Coerced = True

End Function



















Public Property Get DataType() As ARMDataTypeInProfile_Enum
  DataType = meMetricDataType
End Property

Public Property Let DataType(ByVal lNewDataType As Long)
  Dim eNewDataType As ARMDataTypeInProfile_Enum
  
  On Error GoTo ErrorHandler
  
  '--- type check composite enum for both Metric Fields & String Field
  eNewDataType = ConvertDataTypeToProfileEnum(lNewDataType)
  
  meMetricDataType = eNewDataType
  mbMetricPassedIn_Coerced = False    'reset
  
  '--- data type dictates default values
  '--- default values used as actual values if values
  '--- w/ unexpected data types are passed in
  
  Select Case meMetricDataType
  
  Case eProfileMetricsType_InvalidType
    mvMetricDefaultValue = mlMetricInvalidNumber
  
  Case eProfileMetricsType_NumericID32
    mvMetricDefaultValue = mlMetricInvalidNumber
    
  Case eProfileMetricsType_String8
    mvMetricDefaultValue = msMetricInvalidString
  
  Case eProfileMetricsType_String32
    mvMetricDefaultValue = msMetricInvalidString
 
  Case eProfileMetricsType_Date
    mvMetricDefaultValue = mlMetricInvalidNumber
 
  Case eProfileMetricsType_Counter32_AutoIncrement
    mvMetricDefaultValue = mlMetricInvalidNumber
    
  'Case eARM_Counter64
  'Case eARM_CntrDivr32
  'Case eARM_Gauge32
  'Case eARM_Gauge64
  'Case eARM_GaugeDivr32
  
  Case Else
    '--- coerce to data type & default valuesto flag that the metric is missing
    meMetricDataType = eARM_InvalidType
    mvMetricDefaultValue = mlMetricInvalidNumber
  
  End Select
  
  Exit Property
ErrorHandler:
  'Call ARMDebugAssert(Err.Number = 0)           'stop in IDE on purpose
  meMetricDataType = eARM_InvalidType
  mvMetricDefaultValue = mlMetricInvalidNumber
  Exit Property

End Property

Public Property Get Value() As Variant
  Set Value = mvMetricValue
End Property

Public Function ValidateOrCoerceMetricValByExpectType( _
                    ByVal vInVal As Variant _
                    ) As Variant
  
  On Error GoTo ErrorHandler
  
  '--- record what was originally passed in for ARM flag reporting
  mvMetricPassedIn_Value = vInVal
  meMetricPassedIn_Type = VarType(vInVal)
  mbMetricPassedIn_Coerced = False
  
  ' TEST CASES  TEST CASES  TEST CASES  TEST CASES  TEST CASES  TEST CASES
  'TRY CHANGING THE RET VAL
  'TYPE=INVALID; DEFAULT VAL MISSING
  'TYPE<>INVALID; METRIC NOT MISSING
  'TYPE<>INVALID; METRIC IS MISSING; METRIC NOT SET BEFORE
  'TYPE<>INVALID; METRIC IS MISSING; METRIC *WAS* SET BEFORE
  ' TEST CASES  TEST CASES  TEST CASES  TEST CASES  TEST CASES  TEST CASES
  'BM5/
  If IsMissing(vInVal) Then
    If mbMetricPrevSet Then
      '--- Metric missing this time BUT PREVIOUSLY PASSED IN
      '--- assume it validated or was coerced
      ValidateOrCoerceMetricValByExpectType = mvMetricValue
    Else
      '--- Metric missing this time and NOT YET PASSED IN
      ValidateOrCoerceMetricValByExpectType = CoerceInMetircVal2DefaultVal()
    End If
    Exit Function
  End If
  
  '--- METRIC PASSED IN
  mbMetricPrevSet = True
  
  '--- data type dictates default values
  '--- default values used as actual values if values
  '--- w/ unexpected data types are passed in
  
  Select Case meMetricDataType
  
  Case eProfileMetricsType_Counter32_AutoIncrement
    '--- Tx class coerced w/ count of instances
    
  Case eProfileMetricsType_NumericID32
    mvMetricValue = ValidateOrCoerce_Long(vInVal)
    
  Case eProfileMetricsType_String8, eProfileMetricsType_String32
    mvMetricValue = ValidateOrCoerce_String(vInVal)
  
  Case eProfileMetricsType_Date
    mvMetricValue = ValidateOrCoerce_DateToLong(vInVal)
  
  Case eProfileMetricsType_InvalidType
    '--- this will happen if client code said in the InitCmplxTxCls method that
    '--- this metric didn't exist but passed one in in the StartCmplxTx method
    '--- Return metric value coerced to default by Tx class expected data type
    mvMetricValue = Coerce2Invalid_DataType_DefaultVal_Val
    
  Case Else
    '--- Return metric value coerced to default by Tx class expected data type
    mvMetricValue = Coerce2Invalid_DataType_DefaultVal_Val

  End Select


  ValidateOrCoerceMetricValByExpectType = mvMetricValue
  Exit Function

ErrorHandler:
  'Call ARMDebugAssert(Err.Number = 0)       'stop in IDE on purpose
  mvMetricValue = mvMetricDefaultValue
  ValidateOrCoerceMetricValByExpectType = mvMetricDefaultValue
  mbMetricPassedIn_Coerced = True
  Exit Function

End Function

Private Function CoerceInMetircVal2DefaultVal() As Variant
  On Error GoTo ErrorHandler
  
  mbMetricPassedIn_Coerced = True
  'Call ARMDebugAssert((mbMetricPassedIn_Coerced = False))      'stop IDE on purpose
  
  mvMetricValue = mvMetricDefaultValue
  '--- Return metric value coerced to default by Tx class expected data type
  CoerceInMetircVal2DefaultVal = mvMetricDefaultValue
  
  Exit Function
ErrorHandler:
  '--- THIS SHOULD NEVER HAPPEN
  'Call ARMDebugAssert(Err.Number = 0)       'stop in IDE on purpose
  CoerceInMetircVal2DefaultVal = mvMetricDefaultValue
  mvMetricValue = mvMetricDefaultValue
  mbMetricPassedIn_Coerced = True
  
End Function

Private Function Coerce2Invalid_DataType_DefaultVal_Val( _
                          ) As Long
  
  '**********************************
  '   THIS SHOULD NEVER HAPPEN ???
  '**********************************

  mbMetricPassedIn_Coerced = True
  'Call ARMDebugAssert((mbMetricPassedIn_Coerced = False))      'stop IDE on purpose

  meMetricDataType = eARM_InvalidType
  mvMetricDefaultValue = mlMetricInvalidNumber
  mvMetricValue = mlMetricInvalidNumber
  
  Coerce2Invalid_DataType_DefaultVal_Val = mlMetricInvalidNumber
    
End Function

Private Function ConvertDataTypeToProfileEnum( _
                          ByVal lDataType As Long _
                          ) As ARMDataTypeInProfile_Enum
  Dim eRetVal             As ARMDataTypeInProfile_Enum
  On Error GoTo ErrorHandler
  
  Select Case lDataType
    
  Case eARM_InvalidType
    eRetVal = eProfileMetricsType_InvalidType
  
  Case eARM_Counter32_AutoIncrement
    eRetVal = eProfileMetricsType_Counter32_AutoIncrement
  
  'Case eARM_Counter64
  'Case eARM_CntrDivr32
  'Case eARM_Gauge32
  'Case eARM_Gauge64
  'Case eARM_GaugeDivr32
  
  Case eARM_NumericID32
    eRetVal = eProfileMetricsType_NumericID32
  
  'Case eARM_NumericID64
  
  Case eARM_String8
    eRetVal = eProfileMetricsType_String8
  
  Case eARM_String32Only
    eRetVal = eProfileMetricsType_String32
  
  Case eARM_Date
    eRetVal = eProfileMetricsType_Date
  
  Case Else
    eRetVal = eProfileMetricsType_InvalidType

  End Select
  
  ConvertDataTypeToProfileEnum = eRetVal
  Exit Function
  
ErrorHandler:
  ConvertDataTypeToProfileEnum = eProfileMetricsType_InvalidType
  Exit Function

End Function

'Public Property Get DefaultValue() As Variant
'  Select Case meMetricDataType
'
'  Case eARM_NumericID32, eARM_InvalidType
'      mvMetricDefaultValue = mlMetricInvalidNumber
'
'  Case eARM_String8
'      mvMetricDefaultValue = msMetricInvalidString
'
'  Case Else
'
'  End Select
'End Property


'Public Property Get Present() As Boolean
'  Present = mbMetricPresent
'End Property
'
'Public Property Let Present(ByVal bNewValue As Boolean)
'  mbMetricPresent = bNewValue
'End Property


