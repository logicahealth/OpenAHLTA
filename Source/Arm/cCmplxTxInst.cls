VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCmplxTxInst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Public Event CmplxTxInstanceTerminated( _
                      lTxClsID As Long, _
                      lTxInstID As Long)

Public Type Start_Metric_Type
  byMetric(1 To 8)                As Byte         'format/content depends on
                                                  'ARM_GetID_DataBuffer_Type.udtMetric().eMetricType
End Type


Public Type Correlator_Type
  byCorrelator(1 To 168)  As Byte
End Type

Public Enum Start_Format_Enum
  eStartFormat = 1
End Enum

'--- All poss. flags that Correlator field can contain
Public Enum ARM_Start_CorrelatorActAgent_Enum
  'REMEMBER TO CONVERT LONG ENUM TO BYTE
  'FOR READING / SETTING BIT FLAGS IN BYTE
                                        'BIT POSITIONS:
                                        '76543210
  'eAppPassCorrelator2SubTx = gbyBitMask7 '10000000
  'eAppAskAgent4Correlator = gbyBitMask6  '01000000
  eAgentRetCorrelator2App = gbyBitMask5   '00100000
  'eAppAskAgentTrace = gbyBitMask4        '00010000
End Enum

'--- Subset of Correlator Actions that Application can initiate
Public Enum ARM_Start_CorrelatorActApp_Enum
  eAppActNone = 0
  eAppActPassCorrelator2SubTx = gbyBitMask7   '10000000
  eAppActAskAgent4Correlator = gbyBitMask6    '01000000
  'eAppActAgentReturnCorrelator = gbyBitMask5 '00100000
  eAppActAskAgentTrace = gbyBitMask4          '00010000
End Enum

'Public Enum CorrelatedFamilies_Enum
'  eNoFamily
'  eAnyFamily
'  eTestFamily
'End Enum

Public Enum ARM_Start_MetricsFlagByte_Enum
  'REMEMBER TO CONVERT LONG ENUM TO BYTE
  'FOR READING / SETTING BIT FLAGS IN BYTE
                                      'BIT POSITIONS:
                                      '76543210
  eStartMetric1Present = gbyBitMask7  '10000000
  eStartMetric2Present = gbyBitMask6  '01000000
  eStartMetric3Present = gbyBitMask5  '00100000
  eStartMetric4Present = gbyBitMask4  '00010000
  eStartMetric5Present = gbyBitMask3  '00001000
  eStartMetric6Present = gbyBitMask2  '00000100
  eStartString1Present = gbyBitMask1  '00000010
End Enum



'=====================================================================================
'                   ARM_Start / ARM_Stop / ARM_Update DATA BUFFER STRUCT
'=====================================================================================
Public Type ARM_Start_DataBuffer_Type
  'eFormat               As Start_Format_Enum      'const = 1
  eFormat               As Long                   'const = 1
  byFlags(1 To 4)       As Byte                   'Byte(1) = Correlator flags: abcd0000
                                                  ' 1a=1 if app passing Correlator from parent Tx
                                                  ' 1b=1 if app requesting agent generate Correlator
                                                  ' 1C=1 if agent returning Correlator
                                                  ' 1d=1 if app requesting agent to trace the Tx
                                                  'Byte(2) = Metric Flags: abcdefg0
                                                  ' a-f=1 if udtMetric(1 to 6) passed
                                                  ' 2g=1 if sNzString32 passed
  udtMetric(1 To 6)      As Start_Metric_Type
  'byMetric1(1 To 8)       As Byte
  'byMetric2(1 To 8)       As Byte
  'byMetric3(1 To 8)       As Byte
  'byMetric4(1 To 8)       As Byte
  'byMetric5(1 To 8)       As Byte
  'byMetric6(1 To 8)       As Byte
  bysNzString32(1 To 32)  As Byte                 'sNz = STRING *NOT* NULL TERMINATED; PAD W/ BLANKS
  udtCorrelator           As Correlator_Type      '1st 2 bytes = length & ARE part of Correlator
  'byCorrelator(1 To 168)  As Byte     'all data, incl 2 byte length
End Type
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



'*-- SC 01/16/03
Private mlCreationTimeStamp                 As Long
Private mudtCmplxTxInstBuf                  As ARM_Start_DataBuffer_Type
Private moCmplxTxCls                        As ARM.cCmplxTxCls
Private mlTxInstanceID                      As Long
Private meCorrelatorAction                  As ARM_Start_CorrelatorActApp_Enum
Private meCorrelatedFamily                  As CorrelatedFamilies_Enum           'SCR 28779 MSolano 11/21/2002
'Private mbCorrelated                       As Boolean                           'SCR 28779 MSolano 11/21/2002
Private mbCorrelatorSetByPropertyLet        As Boolean                           'SCR 28779 MSolano 11/21/2002
Private mbConcurrentInstancesAllowed        As Boolean
Private mbOrphaned                          As Boolean
Private mbCloned                            As Boolean

Private moCmplxTransactnsCallBk             As ARM.cCmplxTransactions
Private Const msModule                      As String = "cCmplxTxInst"

Private Const msField_TimeStamp             As String = "mlCreationTimeStamp"
'Private Const msField_Buffer                As String = "msField_Buffer"
'Private Const msField_CTCls                 As String = "msField_CTCls"
Private Const msField_InstID                As String = "msField_InstID"
'Private Const msField_CorrAct               As String = "msField_CorrAct"
'Private Const msField_CorrFamily            As String = "msField_CorrFamily"
'Private Const msField_CorrSetByPropLet      As String = "msField_CorrSetByPropLet"
'Private Const msField_Concurrency           As String = "msField_Concurrency"
'Private Const msField_Orphaned             As String = "msField_Orphaned"

Private Sub Class_Initialize()
  mudtCmplxTxInstBuf.eFormat = eStartFormat
  
  '--- default to flag no correlation
  mbCorrelatorSetByPropertyLet = False
  'mbCorrelated = False
  meCorrelatorAction = eAppActNone
  meCorrelatedFamily = eNoFamily
  mlTxInstanceID = glIGNORE_ME
  mbConcurrentInstancesAllowed = False
  
  '*-- SC 011603
  mlCreationTimeStamp = Timer
  Debug.Print "Cls_Init Instance/TS: "; CStr(mlTxInstanceID) & " / " & CStr(mlCreationTimeStamp)
  
End Sub


'--- set which data field is present
'--- (none to 6 metric fields, or string field)
'--- set data field's name
'--- set data field's data type
Private Function SetMetricField( _
                       ByVal vMetricVal As Variant, _
                       ByVal eWhichMetricByArrayPos As ARM_Metrics_ArrayIndex_Enum _
                       ) As Boolean
                       'ByVal eMetricType As ARMDataType4MetricsFld_Enum _

  Dim sData                                 As String
  Dim lCoerce2Long                          As Long
  Dim bContinue                             As Boolean
  Dim lNdx                                  As Long
  
  On Error GoTo ErrorHandler
  SetMetricField = False
  
  'Select Case eMetricType
  Select Case VarType(vMetricVal)
  
  'Case eARM_NumericID32
  Case vbLong  '
    '--- don't bother w/ VarType(vMetricVal) = vbInteger
    '--- it should have already been coerced literally to a long,
    '--- not just a numeric type
    
    '--- don't bother packing zeros (the bytes are initialized to zero)
    If vMetricVal = 0 Then
      SetMetricField = True
      Exit Function
    End If
    
    lCoerce2Long = CLng(vMetricVal)
    bContinue = InsertLongInMetric( _
                    eWhichMetricByArrayPos, _
                    lCoerce2Long)
    Call ARMDebugAssert(bContinue)
    If Not bContinue Then Exit Function
  
  
  'Case eARM_String8
  Case vbString
    sData = CStr(vMetricVal)
    
    '--- both Metric & String#1 fields - SPACE term'd, NOT NULL term'd
    
    bContinue = ANSIStringWSpecifiedTerminator( _
                mudtCmplxTxInstBuf.udtMetric(eWhichMetricByArrayPos).byMetric(), _
                sData, _
                gbyBYTE_CHAR_SPACE)
    Call ARMDebugAssert(bContinue)
    If Not bContinue Then Exit Function
  
  Case Else
    SetMetricField = False
    Call ARMDebugAssert(SetMetricField)
    Exit Function
  
  End Select
    
  SetMetricField = True
  Exit Function
ErrorHandler:
  SetMetricField = False
  Call ARMDebugAssert(SetMetricField)    'stop in IDE
  Resume
  Exit Function

End Function

'------------------------------------------------
'--- Procedure: SetStringField
'---
'--- set string field as present
'--- set data field's name
'--- set data field's data type
'------------------------------------------------
Private Function SetStringField( _
                      ByVal vMetricVal As Variant _
                      ) As Boolean
              
  Dim sMetricVal      As String
  Dim bContinue       As Boolean
  
  On Error GoTo ErrorHandler
  SetStringField = False
  
  sMetricVal = CStr(vMetricVal)
      
  '--- String#1 field is SPACE terminated, NOT NULL term'd
  bContinue = ANSIStringWSpecifiedTerminator( _
                  mudtCmplxTxInstBuf.bysNzString32, _
                  sMetricVal, _
                  gbyBYTE_CHAR_SPACE)
  Call ARMDebugAssert(bContinue)
  If Not bContinue Then Exit Function

  SetStringField = True
  Exit Function
ErrorHandler:
  SetStringField = False
  Call ARMDebugAssert(SetStringField)    'stop in IDE
  Exit Function

End Function

Private Function InsertLongInMetric( _
                    ByVal eWhichMetricByArrayPos As ARM_Metrics_ArrayIndex_Enum, _
                    ByVal lData As Long _
                    ) As Boolean
  Dim byTemp()      As Byte
  Dim bContinue     As Boolean
  
  On Error GoTo ErrorHandler
  InsertLongInMetric = False
  
  'byTemp = LongTo4Bytes_A(lData)
  byTemp = LongTo4Bytes_B(lData)
   
  bContinue = Pack4Bytes(byTemp(), mudtCmplxTxInstBuf.udtMetric(eWhichMetricByArrayPos).byMetric())

  InsertLongInMetric = bContinue
  Exit Function
ErrorHandler:
  InsertLongInMetric = False
  Call ARMDebugAssert(InsertLongInMetric)    'stop in IDE
  Exit Function

End Function

Private Function Pack4Bytes( _
                    ByRef by4BytesOfData() As Byte, _
                    ByRef byInOut() As Byte _
                    ) As Boolean
  On Error GoTo ErrorHandler
  Pack4Bytes = False
    
  byInOut(1) = by4BytesOfData(1)
  byInOut(2) = by4BytesOfData(2)
  byInOut(3) = by4BytesOfData(3)
  byInOut(4) = by4BytesOfData(4)
  
  Pack4Bytes = True
  Exit Function
ErrorHandler:
  Pack4Bytes = False
  Exit Function
End Function



'--- "ALL OR NONE" TAPM ARM BUG regarding metrics - per Tivoli / Itally
'--- giampiero catasta giampiero.catasta@ it.ibm.com
'--- L3 Customer Support Tivoli Application Performance Management
'--- either all flags for metrics are set
'--- or none are set
'--- either all metrics are sent
'--- or none are sent
Private Function TAPMBugAllMetrics_SetFlags( _
                      ) As Boolean
  Dim byFlags        As Byte
  
  On Error GoTo ErrorHandler
  TAPMBugAllMetrics_SetFlags = False
  
  '--- SET 'EM ALL ... LET TIVOLI FIGURE IT OUT
  byFlags = byFlags Or gbyBitMask1
  byFlags = byFlags Or gbyBitMask2
  byFlags = byFlags Or gbyBitMask3
  byFlags = byFlags Or gbyBitMask4
  byFlags = byFlags Or gbyBitMask5
  byFlags = byFlags Or gbyBitMask6
  byFlags = byFlags Or gbyBitMask7
  
  mudtCmplxTxInstBuf.byFlags(glARM_Start_FlagsByteMetrics) = byFlags
  
  TAPMBugAllMetrics_SetFlags = True
  Exit Function
ErrorHandler:
  TAPMBugAllMetrics_SetFlags = False
  Call ARMDebugAssert(TAPMBugAllMetrics_SetFlags)    'stop in IDE
  Exit Function
End Function
  
Private Function ValidateCorrelatorAction( _
                      eNewActionVal As ARM_Start_CorrelatorActApp_Enum _
                      ) As ARM_Start_CorrelatorActApp_Enum
  '--- these properties have to agree that there IS, or IS NOT correlation
  '---    meCorrelatorAction
  '---    meCorrelatedFamily
  '---    mbCorrelated
  On Error GoTo ErrorHandler

  If eNewActionVal = eAppActNone Then                       'act prop says there ISN'T correlation
    'mbCorrelated = False                                   'bool prop says there ISN'T correlation
    meCorrelatedFamily = eNoFamily                          'family prop says there ISN'T correlation
  
  Else                                                      'act prop says there IS correlation
  
    'mbCorrelated = True                                    'bool prop says there IS correlation
    If meCorrelatedFamily = eNoFamily Then
      Call ARMDebugAssert(meCorrelatedFamily <> eNoFamily)  'stop IDE intentionally
      meCorrelatedFamily = eAnyFamily                       'family prop COERCED says there IS correlation
    End If
  End If

  ValidateCorrelatorAction = eNewActionVal
  Exit Function
ErrorHandler:
  ValidateCorrelatorAction = eAppActNone
  Call ARMDebugAssert(Err.Number = 0)    'stop in IDE
  Exit Function
End Function

Private Function ValidateCorrelatedFamily( _
                      eNewValue As CorrelatedFamilies_Enum _
                      ) As ARM_Start_CorrelatorActApp_Enum
  '--- these properties have to agree that there IS, or IS NOT correlation
  '---    meCorrelatorAction
  '---    meCorrelatedFamily
  '---    mbCorrelated
  On Error GoTo ErrorHandler

  '--- these properties have to agree that there IS, or IS NOT correlation
  '---    meCorrelatorAction
  '---    meCorrelatedFamily
  '---    mbCorrelated
  If eNewValue = eNoFamily Then                          'family prop. says there ISN'T correlation
    'mbCorrelated = False                                'bool prop. says there ISN'T correlation
    meCorrelatorAction = eAppActNone                     'act prop. says there ISN'T correlation
  
  Else                                                   'family prop. says there IS correlation
  
    'mbCorrelated = True                                 'bool prop. says there IS correlation
    If meCorrelatorAction = eAppActNone Then
      If mbCorrelatorSetByPropertyLet = True Then
        'assumption: outside client code set correlator for Tx instance to another in its family
        meCorrelatorAction = eAppActPassCorrelator2SubTx  'act prop. COERCED says there IS correlation
      Else
        'assumption: this is the first instance in its family, OR
        'client code will later set correlator for curr. Tx instance to another in its family
        meCorrelatorAction = eAppActAskAgent4Correlator
      End If
    End If
  End If

  ValidateCorrelatedFamily = eNewValue

  Exit Function
ErrorHandler:
  Call ARMDebugAssert(Err.Number = 0)    'stop in IDE
  Exit Function
End Function

Private Sub Class_ReadProperties(PropBag As PropertyBag)
  On Error GoTo ErrorHandler
  With PropBag
    Debug.Print "Cls_ReadProp1 Instance/TS: "; CStr(mlTxInstanceID) & " / " & CStr(mlCreationTimeStamp)
    
    mlCreationTimeStamp = .ReadProperty(msField_TimeStamp)
    mlTxInstanceID = .ReadProperty(msField_InstID)
    
    Debug.Print "Cls_ReadProp2 Instance/TS: "; CStr(mlTxInstanceID) & " / " & CStr(mlCreationTimeStamp)
  End With
  Exit Sub
ErrorHandler:
  Call ARMDebugAssert(Err.Number = 0)
  Exit Sub
End Sub
Private Sub Class_WriteProperties(PropBag As PropertyBag)
  On Error GoTo ErrorHandler
  
  With PropBag
    Debug.Print "Cls_WriteProp1 Instance/TS: "; CStr(mlTxInstanceID) & " / " & CStr(mlCreationTimeStamp)
  
    .WriteProperty msField_TimeStamp, mlCreationTimeStamp
    '.WriteProperty msField_, mudtCmplxTxInstBuf
    '.WriteProperty msField_, moCmplxTxCls
    .WriteProperty msField_InstID, mlTxInstanceID
    '.WriteProperty msField_, meCorrelatorAction
    '.WriteProperty msField_, meCorrelatedFamily
    '.WriteProperty msField_, mbCorrelatorSetByPropertyLet
    '.WriteProperty msField_, mbConcurrentInstancesAllowed
    
    Debug.Print "Cls_WriteProp2 Instance/TS: "; CStr(mlTxInstanceID) & " / " & CStr(mlCreationTimeStamp)
  End With
  Exit Sub
ErrorHandler:
  Call ARMDebugAssert(Err.Number = 0)
  Exit Sub
End Sub

Private Sub Class_Terminate()
  On Error GoTo ErrorHandler

  '--- notify cCmplxTransactions Tx Instance destroyed
  '*-- SCR 31901 Coleman 2/20/03 - Changed signature to post more info on GUI
  Call moCmplxTransactnsCallBk.CallBk_CTInstTerminated( _
                                moCmplxTxCls.TxClassName, _
                                moCmplxTxCls.TxClassID, _
                                mlTxInstanceID, _
                                meCorrelatedFamily, _
                                mlCreationTimeStamp)

  Set moCmplxTxCls = Nothing
  Set moCmplxTransactnsCallBk = Nothing
  Exit Sub
ErrorHandler:
  Call ARMDebugAssert(Err.Number = 0)
  Exit Sub
  Resume
End Sub









Public Sub CallBkRegistration(oCmplxTxs As ARM.cCmplxTransactions)
  On Error GoTo ErrorHandler
  
  Call ARMDebugAssert(Not (oCmplxTxs Is Nothing), msModule, "CallBkRegistration", "A1", 1)
  
  Set moCmplxTransactnsCallBk = oCmplxTxs
  
  Exit Sub
ErrorHandler:
  Call ARMDebugAssert(Err.Number = 0)
  Exit Sub
End Sub

'*-- SC 011603
Public Property Get CreationTimeStamp() As Long
    CreationTimeStamp = mlCreationTimeStamp
End Property

Public Property Get TxClassName() As String
  On Error GoTo ErrorHandler
  
  If moCmplxTxCls Is Nothing Then
    TxClassName = "ERROR"
  Else
    TxClassName = moCmplxTxCls.TxClassName
  End If
  
  Exit Function
ErrorHandler:
  TxClassName = "ERROR"
  Call ARMDebugAssert(Err.Number = 0)    'stop in IDE
  Exit Function

End Property
Public Property Get TxClassID() As String
  On Error GoTo ErrorHandler
  
  If moCmplxTxCls Is Nothing Then
    TxClassID = glIGNORE_TX
  Else
    TxClassID = moCmplxTxCls.TxClassID
  End If
  
  Exit Property
ErrorHandler:
  TxClassID = glIGNORE_TX
  Call ARMDebugAssert(Err.Number = 0)    'stop in IDE
  Exit Property

End Property


Public Property Get TxInstanceID() As Long
  On Error GoTo ErrorHandler
  
  TxInstanceID = mlTxInstanceID
  
  Exit Property
ErrorHandler:
  mlTxInstanceID = glIGNORE_TX
  TxInstanceID = glIGNORE_TX
  Call ARMDebugAssert(Err.Number = 0)    'stop in IDE
  Exit Property

End Property

Public Property Let TxInstanceID(ByVal lData As Long)
  On Error GoTo ErrorHandler
  
  If lData > 0 Then
    mlTxInstanceID = lData
  Else
    mlTxInstanceID = glIGNORE_TX
    Call ARMDebugAssert(mlTxInstanceID <> glIGNORE_TX)
  End If
    
  Exit Property
ErrorHandler:
  mlTxInstanceID = glIGNORE_TX
  Call ARMDebugAssert(Err.Number = 0)    'stop in IDE
  Exit Property
End Property

Public Property Set CmplxTxClass(ByRef oTxCls As ARM.cCmplxTxCls)
  On Error GoTo ErrorHandler
  
  If moCmplxTxCls Is Nothing Then
    Set moCmplxTxCls = oTxCls
    '--- SCR 28568 Tx Class AutoCount Feature MSolano 10/16/02
    moCmplxTxCls.AutoCount = moCmplxTxCls.AutoCount + 1
    
  Else
  
    Call ARMDebugAssert(moCmplxTxCls Is Nothing)
  End If
  
  Exit Property
ErrorHandler:
  Set moCmplxTxCls = Nothing
  Call ARMDebugAssert(Err.Number = 0)    'stop in IDE
  Exit Property

End Property
Public Property Get CmplxTxClass() As ARM.cCmplxTxCls
  On Error GoTo ErrorHandler
  
  Set CmplxTxClass = moCmplxTxCls
  
  Exit Function
ErrorHandler:
  Set CmplxTxClass = Nothing
  Call ARMDebugAssert(Err.Number = 0)
  Exit Function
End Property

Public Function SetMetric( _
                           eWhichMetric As ARM_Metrics_ArrayIndex_Enum, _
                  Optional ByRef vMetricVal As Variant _
                      ) As Variant
  Dim bContinue As Boolean
  On Error GoTo ErrorHandler
  
  If moCmplxTxCls Is Nothing Then
    Call ARMDebugAssert(Not (moCmplxTxCls Is Nothing))   'stop IDE
    Exit Function
  End If
  
  'BM3/
  vMetricVal = moCmplxTxCls.ValidateOrCoerceMetricValByExpectType( _
                  eWhichMetric, _
                  vMetricVal)
  
  'populate the buffer
  ' use the same buffer in start() / update() / stop
  If eWhichMetric = eMetricStringIndex Then
    bContinue = SetStringField(vMetricVal)
    Call ARMDebugAssert(bContinue)
  
  Else
  
    bContinue = SetMetricField(vMetricVal, eWhichMetric)
    Call ARMDebugAssert(bContinue)
  End If

  
  Exit Function
ErrorHandler:
  Call ARMDebugAssert(Err.Number = 0)    'stop in IDE
  Exit Function

End Function

Public Function DataBuffer( _
                  Optional ByVal vData31Char As Variant, _
                  Optional ByVal vMetric1 As Variant, _
                  Optional ByVal vMetric2 As Variant, _
                  Optional ByVal vMetric3 As Variant, _
                  Optional ByVal vMetric4 As Variant, _
                  Optional ByVal vMetric5 As Variant, _
                  Optional ByVal vMetric6 As Variant _
                  ) As ARM_Start_DataBuffer_Type
  
  Dim bContinue       As Boolean
  Dim i As Long
  On Error GoTo ErrorHandler
  
  bContinue = TAPMBugAllMetrics_SetFlags()
  
  '...BMTODO BMTEST
  ' > TEST CASE WHERE IsMissing(which input val?) = true - WHAT DOES INITIALIZE() DO?
  ' > If IsMissing() Then
    
'  bContinue = SetStringField(vData31Char)
'  Call ARMDebugAssert(bContinue)
'
'  bContinue = SetMetricField(vMetric1, eMetric1Index)
'  Call ARMDebugAssert(bContinue)
'
'  bContinue = SetMetricField(vMetric2, eMetric2Index)
'  Call ARMDebugAssert(bContinue)
'
'  bContinue = SetMetricField(vMetric3, eMetric3Index)
'  Call ARMDebugAssert(bContinue)
'
'  bContinue = SetMetricField(vMetric4, eMetric4Index)
'  Call ARMDebugAssert(bContinue)
'
'  bContinue = SetMetricField(vMetric5, eMetric5Index)
'  Call ARMDebugAssert(bContinue)
'
'  bContinue = SetMetricField(vMetric6, eMetric6Index)
'  Call ARMDebugAssert(bContinue)

  Call ARMDebugAssert(LenB(mudtCmplxTxInstBuf) = glARM_Start_DataBuffer_Size)
    
  DataBuffer = mudtCmplxTxInstBuf
  
  Exit Function
ErrorHandler:
  Exit Function

End Function

Public Property Get DataBufferLength() As Long
  On Error GoTo ErrorHandler
  
  DataBufferLength = LenB(mudtCmplxTxInstBuf)
  
  Exit Property
ErrorHandler:
  Call ARMDebugAssert(Err.Number = 0)    'stop in IDE
  Exit Property
End Property

Public Property Let Correlator(udtCorr As Correlator_Type)
  On Error GoTo ErrorHandler
  
  mudtCmplxTxInstBuf.udtCorrelator = udtCorr
  
  '--- flag that correlator value explicitly passed in
  mbCorrelatorSetByPropertyLet = True
  
  '--- only legit val of 3 if correlator explicitly set
  '--- side effect: fires Prop Let CorrelatorAction ()
  Me.CorrelatorAction = eAppActPassCorrelator2SubTx
  
  '--- flag this Tx instance as correlated
  'mbCorrelated = True
  
  Exit Property
ErrorHandler:
  Call ARMDebugAssert(Err.Number = 0)    'stop in IDE
  Exit Property
End Property

Public Property Get Correlator() As Correlator_Type
  On Error GoTo ErrorHandler
  
  Correlator = mudtCmplxTxInstBuf.udtCorrelator
  
  Exit Property
ErrorHandler:
  Call ARMDebugAssert(Err.Number = 0)    'stop in IDE
  Exit Property
End Property

'>>> SCR 28779 MSolano 11/21/2002
Public Property Get CorrelatedFamily() As CorrelatedFamilies_Enum
  On Error GoTo ErrorHandler
  
  If meCorrelatedFamily = eAnyFamily Then
    mbConcurrentInstancesAllowed = True
  End If

  CorrelatedFamily = meCorrelatedFamily

  Exit Property
ErrorHandler:
  Call ARMDebugAssert(Err.Number = 0)    'stop in IDE
  Exit Property

End Property

Public Property Let CorrelatedFamily(ByVal eNewValue As CorrelatedFamilies_Enum)
  On Error GoTo ErrorHandler
  
  meCorrelatedFamily = ValidateCorrelatedFamily(eNewValue)

  Exit Property
ErrorHandler:
  Call ARMDebugAssert(Err.Number = 0)    'stop in IDE
  Exit Property

End Property

Public Property Let CorrelatorAction( _
                        eNewVal As ARM_Start_CorrelatorActApp_Enum)
  Dim byFlags           As Byte
  On Error GoTo ErrorHandler
  
  meCorrelatorAction = ValidateCorrelatorAction(eNewVal)
  
  '--- set the correlator flags field in the buffer
  byFlags = mudtCmplxTxInstBuf.byFlags(glARM_Start_FlagsByteCorrelator)
  
  '--- OH BOY ... trusting the ARM 2.0 Documentation here
  '--- These val's are MUTUALLY EXCLUSIVELY settable by Application
  '---    eAppActPassCorrelator2SubTx
  '---    eAppActAskAgent4Correlator
  '--- TAPM Agent can set
  '---    eAgentRetCorrelator2App
  If byFlags And eAgentRetCorrelator2App Then
    '--- TAPM Agent set the correlator returned bit - keep it
    '*-- 010303 - Coleman - Change flag calculation
    'byFlags = byFlags Or meCorrelatorAction
    byFlags = meCorrelatorAction Or eAgentRetCorrelator2App
  Else
    byFlags = meCorrelatorAction
  End If
  
  mudtCmplxTxInstBuf.byFlags(glARM_Start_FlagsByteCorrelator) = byFlags
 
  Exit Property
ErrorHandler:
  Call ARMDebugAssert(Err.Number = 0)    'stop in IDE
  Exit Property
End Property
'<<< SCR 28779

Public Function StartCmplxTxInstance( _
                  Optional ByVal vData31Char As Variant, _
                  Optional ByVal vMetric1 As Variant, _
                  Optional ByVal vMetric2 As Variant, _
                  Optional ByVal vMetric3 As Variant, _
                  Optional ByVal vMetric4 As Variant, _
                  Optional ByVal vMetric5 As Variant, _
                  Optional ByVal vMetric6 As Variant _
                  ) As Boolean
  On Error GoTo ErrorHandler
  
  'BM1/
  Call SetMetric(eMetric1Index, vMetric1)
  Call SetMetric(eMetric2Index, vMetric2)
  Call SetMetric(eMetric3Index, vMetric3)
  Call SetMetric(eMetric4Index, vMetric4)
  Call SetMetric(eMetric5Index, vMetric5)
  Call SetMetric(eMetric6Index, vMetric6)
  Call SetMetric(eMetricStringIndex, vData31Char)
  
  mlTxInstanceID = goARMAPI.StartCmplxTxInstance( _
                          moCmplxTxCls.TxClassID, _
                          mudtCmplxTxInstBuf, _
                          LenB(mudtCmplxTxInstBuf))

  '--- definition of success is a positive ret val from API:
  StartCmplxTxInstance = (mlTxInstanceID > glIGNORE_TX)
  
  'RWalker 12/19/02 Need to record by sReadableTxClsName versus sSomeTxHandle
  #If DEBUGON Then
    Call D_bugLog("cARM", "StartCplxTx", moCmplxTxCls.TxClassName, moCmplxTxCls.TxClassID, mlTxInstanceID, CLng(StartCmplxTxInstance))
    Call D_bugLogMetrics("cARM", "Tx Metrics", moCmplxTxCls.TxClassName, _
                           CStr(vData31Char), _
                           "", vMetric1, _
                           "", vMetric2, _
                           "", vMetric3, _
                           "", vMetric4, _
                           "", vMetric5, _
                           "", vMetric6)
  #End If
  
  Call ARMDebugAssert(StartCmplxTxInstance)
  
  Exit Function
ErrorHandler:
  Call ARMDebugAssert(Err.Number = 0)    'stop in IDE
  Exit Function
    
End Function
Public Function StopCmplxTxInstance( _
                  Optional ByVal eTxStatus As enmARMTxStatus = enmARMTxStatusOk, _
                  Optional ByVal vData31Char As Variant, _
                  Optional ByVal vMetric1 As Variant, _
                  Optional ByVal vMetric2 As Variant, _
                  Optional ByVal vMetric3 As Variant, _
                  Optional ByVal vMetric4 As Variant, _
                  Optional ByVal vMetric5 As Variant, _
                  Optional ByVal vMetric6 As Variant _
                  ) As Boolean

  On Error GoTo ErrorHandler

  'BM1/
  'USE cCmplxTxCls PROP'S TO COERCE
  Call SetMetric(eMetric1Index, vMetric1)
  Call SetMetric(eMetric2Index, vMetric2)
  Call SetMetric(eMetric3Index, vMetric3)
  Call SetMetric(eMetric4Index, vMetric4)
  Call SetMetric(eMetric5Index, vMetric5)
  Call SetMetric(eMetric6Index, vMetric6)
  Call SetMetric(eMetricStringIndex, vData31Char)

  StopCmplxTxInstance = goARMAPI.StopCmplxTxInstance( _
                              mlTxInstanceID, _
                              eTxStatus, _
                              mudtCmplxTxInstBuf)

  'RWalker 12/30/02 - ARM - Log sReadableTxClsName and metrics
  #If DEBUGON Then
  Call D_bugLogMetrics("cARM", "Tx Metrics", TxClassName, _
                         CStr(vData31Char), _
                         "", vMetric1, _
                         "", vMetric2, _
                         "", vMetric3, _
                         "", vMetric4, _
                         "", vMetric5, _
                         "", vMetric6)
    Call D_bugLog("cARM", "StopCplxTx", TxClassName, TxClassID, mlTxInstanceID, CLng(eTxStatus))
  #End If

  Call ARMDebugAssert(StopCmplxTxInstance)

  Exit Function
ErrorHandler:
  Call ARMDebugAssert(Err.Number = 0)    'stop in IDE
  Exit Function

End Function

Public Property Get ConcurrentInstancesAllowed() As Boolean
  ConcurrentInstancesAllowed = mbConcurrentInstancesAllowed
End Property
Public Property Let Orphaned(bNewVal As Boolean)
  mbOrphaned = bNewVal
End Property
Public Property Get Orphaned() As Boolean
  Orphaned = mbOrphaned
End Property
Public Property Let Cloned(bNewVal As Boolean)
  mbCloned = bNewVal
End Property
Public Property Get Cloned() As Boolean
  Cloned = mbCloned
End Property










'Private Function SetMetricsFlags(ByVal eWhichFlag As ARM_Metrics_Flag_Enum) As Boolean
'  Dim byFlags        As Byte
'
'  On Error GoTo ErrorHandler
'  SetMetricsFlags = False
'
'  If eWhichFlag < 0 Or eWhichFlag > 255 Then
'
'    Exit Function
'  End If
'
'  byFlags = mudtCmplxTxInstBuf.byFlags(glARM_Start_FlagsByteMetrics)
'  '--- set desired bit
'  mudtCmplxTxInstBuf.byFlags(glARM_Start_FlagsByteMetrics) = byFlags Or CByte(eWhichFlag)
'
'  SetMetricsFlags = True
'  Exit Function
'ErrorHandler:
'  SetMetricsFlags = False
'  call ARMDebugAssert(SetMetricsFlags    'stop in IDE
'  Exit Function
'End Function



'Private Function InsertLongInMetric_20020820( _
'                    udtByteArray As Start_Metric_Type, _
'                    lData As Long _
'                    ) As Boolean
'  Dim byTemp() As Byte
'
'  On Error GoTo ErrorHandler
'  InsertLongInMetric = False
'
'  byTemp = LongTo4Bytes_A(lData)
'
'  udtByteArray.byMetric(1) = byTemp(1)
'  udtByteArray.byMetric(2) = byTemp(2)
'  udtByteArray.byMetric(3) = byTemp(3)
'  udtByteArray.byMetric(4) = byTemp(4)
'
'  InsertLongInMetric = True
'  Exit Function
'ErrorHandler:
'  udtByteArray.byMetric(1) = 0
'  udtByteArray.byMetric(2) = 0
'  udtByteArray.byMetric(3) = 0
'  udtByteArray.byMetric(4) = 0
'  InsertLongInMetric_20020820 = False
'  call ARMDebugAssert(InsertLongInMetric_20020820    'stop in IDE
'  Exit Function
'
'End Function





'Private Function SetMetric( _
'                       ByVal sMetricName As String, _
'                       ByVal eWhichMetricByArrayPos As ARM_Metrics_ArrayIndex_Enum, _
'              Optional ByVal eMetricType As ARMDataType4MetricsFld_Enum = eARM_NumericID32 _
'              ) As Boolean
'  Dim MetricType       As ARMDataType4MetricsFld_Enum
'
'  On Error GoTo ErrorHandler
'  SetMetric = False
'
'
'
'  Select Case eWhichMetricByArrayPos
'  Case eMetric1Index
'    mudtCmplxTxInstBuf.byMetric1 = MetricStringValue(sMetricName)
'  Case eMetric2Index
'    mudtCmplxTxInstBuf.Metric2 = MetricStringValue(sMetricName)
'  Case eMetric3Index
'    mudtCmplxTxInstBuf.Metric3 = MetricStringValue(sMetricName)
'  Case eMetric4Index
'    mudtCmplxTxInstBuf.Metric4 = MetricStringValue(sMetricName)
'  Case eMetric5Index
'    mudtCmplxTxInstBuf.Metric5 = MetricStringValue(sMetricName)
'  Case eMetric6Index
'    mudtCmplxTxInstBuf.Metric6 = MetricStringValue(sMetricName)
'  End Select
'
'  SetMetric = True
'
'  Exit Function
'ErrorHandler:
'  SetMetric = False
'  call ARMDebugAssert(SetMetric
'End Function
'Private Function MetricStringValue( _
'                    ByVal sMetricName As String _
'                    ) As Byte()
'  Dim byRet()     As Byte
'  Dim bContinue   As Boolean
'
'  On Error GoTo ErrorHandler
'  ReDim byRet(1 To glARM_GetID_MetricNameBufSize)
'
'  '--- Other metrics terminated by Null char
'  bContinue = ANSIStringWSpecifiedTerminator( _
'                  byRet(), _
'                  sMetricName, _
'                  gbyBYTE_CHAR_SPACE)
'  call ARMDebugAssert(bContinue
'  If Not bContinue Then Exit Function
'
'  MetricStringValue = byRet()
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
'Private Function SetDataField_20020820( _
'                       ByVal vMetricVal As Variant, _
'                       ByVal eWhichMetricByByteFlag As ARM_Metrics_Flag_Enum, _
'                    Optional eWhichMetricByArrayPos As ARM_Metrics_ArrayIndex_Enum, _
'              Optional ByVal eMetricType As ARMDataType4MetricsFld_Enum = eARM_NumericID32 _
'              ) As Boolean
'
'  Dim lMetricType     As Long
'  Dim sData           As String
'  Dim lData           As Long
'  Dim bContinue       As Boolean
'
'  On Error GoTo ErrorHandler
'  SetDataField = False
'
'
'
'  '--- Set Flag to note which Metric field is present
'  bContinue = SetMetricsFlags(eWhichMetricByByteFlag)
'  call ARMDebugAssert(bContinue
'  If Not bContinue Then Exit Function
'
'  If eWhichMetricByByteFlag = eString1Present Or eMetricType = eARM_String8 Then
'    '--- We have string data either for:
'    '---  - one of the Metric(1 to 6) fields of string type
'    '---  - Or for the String#1 field
'    bContinue = (VarType(vMetricVal) = vbString)
'    call ARMDebugAssert(bContinue
'    If Not bContinue Then Exit Function
'
'    sData = CStr(vMetricVal)
'
'    '--- both Metric & String#1 fields - SPACE term'd, NOT NULL term'd
'   If eWhichMetricByByteFlag = eString1Present Then
'      bContinue = ANSIStringWSpecifiedTerminator( _
'                      mudtCmplxTxInstBuf.bysNzString32, _
'                      sData, _
'                      gbyBYTE_CHAR_SPACE)
'      call ARMDebugAssert(bContinue
'      If Not bContinue Then Exit Function
'
'   Else 'eMetricType = eARM_String8
'      bContinue = ANSIStringWSpecifiedTerminator( _
'                      mudtCmplxTxInstBuf.udtMetric(eWhichMetricByArrayPos).byMetric, _
'                      sData, _
'                      gbyBYTE_CHAR_SPACE)
'      call ARMDebugAssert(bContinue
'      If Not bContinue Then Exit Function
'
'   End If
'  End If
'
'  '
'
'  If eMetricType = eARM_NumericID32 Then
'    '--- CONVERT LONG TO FIRST 4 BYTES OF
'    bContinue = (VarType(vMetricVal) = vbLong)
'    call ARMDebugAssert(bContinue
'    lData = CLng(CStr(vMetricVal))
'    bContinue = (lData > 0)
'    call ARMDebugAssert(bContinue
'    If Not bContinue Then Exit Function
'    'InsertLongInMetric1
'    bContinue = InsertLongInMetric( _
'                  eWhichMetricByArrayPos, _
'                  lData)
'    call ARMDebugAssert(bContinue
'    If Not bContinue Then Exit Function
'  End If
'
'  SetDataField = True
'  Exit Function
'ErrorHandler:
'  SetDataField = False
'  call ARMDebugAssert(SetDataField    'stop in IDE
'  Exit Function
'
'End Function



'--- set which data field is present
'--- (none to 6 metric fields, or string field)
'--- set data field's name
'--- set data field's data type
'Private Function SetMetricField_V1( _
'                       ByVal vMetricVal As Variant, _
'                       ByVal eWhichMetricByArrayPos As ARM_Metrics_ArrayIndex_Enum _
'                       ) As Boolean
'                       'ByVal eMetricType As ARMDataType4MetricsFld_Enum _
'
'  Dim sData           As String
'  Dim lCoerce2Long    As Long
'  Dim bContinue       As Boolean
'
'  On Error GoTo ErrorHandler
'  SetMetricField = False
'
'
'  '--- THIS BLOCK REPLACED BY CALL TO
'  '--- TAPMBugAllMetrics_SetFlags IN Init_EncounterID
'  '--- AS (TEMP?) WORKAROUND FOR "ALL OR NONE" TAPM ARM BUG
'  '  '--- Set Flag to note which Metric field is present
'  '  bContinue = SetMetricsFlags(eWhichMetricByByteFlag)
'  '  call ARMDebugAssert(bContinue
'  '  If Not bContinue Then Exit Function
'
'  'Select Case eMetricType
'  Select Case VarType(vMetricVal)
'
'  'Case eARM_NumericID32
'  Case vbLong  '
'    '--- don't bother w/ VarType(vMetricVal) = vbInteger
'    '--- it should have already been coerced literally to a long,
'    '--- not just a numeric type
'
'    '--- don't bother packing zeros (the bytes are initialized to zero)
'    If vMetricVal = 0 Then
'      SetMetricField = True
'      Exit Function
'    End If
'
'
'
'    lCoerce2Long = CLng(vMetricVal)
'    bContinue = InsertLongInMetric( _
'                    eWhichMetricByArrayPos, _
'                    lCoerce2Long)
'    Call ARMDebugAssert(bContinue)
'    If Not bContinue Then Exit Function
'
'
'  'Case eARM_String8
'  Case vbString
'    sData = CStr(vMetricVal)
'
'    '--- both Metric & String#1 fields - SPACE term'd, NOT NULL term'd
'
'      Select Case eWhichMetricByArrayPos
'
'        Case eMetric1Index
'          bContinue = ANSIStringWSpecifiedTerminator( _
'                          mudtCmplxTxInstBuf.byMetric1, _
'                          sData, _
'                          gbyBYTE_CHAR_SPACE)
'          Call ARMDebugAssert(bContinue)
'          If Not bContinue Then Exit Function
'
'        Case eMetric2Index
'          bContinue = ANSIStringWSpecifiedTerminator( _
'                          mudtCmplxTxInstBuf.byMetric2, _
'                          sData, _
'                          gbyBYTE_CHAR_SPACE)
'          Call ARMDebugAssert(bContinue)
'          If Not bContinue Then Exit Function
'
'        Case eMetric3Index
'          bContinue = ANSIStringWSpecifiedTerminator( _
'                          mudtCmplxTxInstBuf.byMetric3, _
'                          sData, _
'                          gbyBYTE_CHAR_SPACE)
'          Call ARMDebugAssert(bContinue)
'          If Not bContinue Then Exit Function
'
'        Case eMetric4Index
'          bContinue = ANSIStringWSpecifiedTerminator( _
'                          mudtCmplxTxInstBuf.byMetric4, _
'                          sData, _
'                          gbyBYTE_CHAR_SPACE)
'          Call ARMDebugAssert(bContinue)
'          If Not bContinue Then Exit Function
'
'        Case eMetric5Index
'          bContinue = ANSIStringWSpecifiedTerminator( _
'                          mudtCmplxTxInstBuf.byMetric5, _
'                          sData, _
'                          gbyBYTE_CHAR_SPACE)
'          Call ARMDebugAssert(bContinue)
'          If Not bContinue Then Exit Function
'
'        Case eMetric6Index
'          bContinue = ANSIStringWSpecifiedTerminator( _
'                          mudtCmplxTxInstBuf.byMetric6, _
'                          sData, _
'                          gbyBYTE_CHAR_SPACE)
'          Call ARMDebugAssert(bContinue)
'          If Not bContinue Then Exit Function
'
'      End Select
'
'
'
'  Case Else
'    SetMetricField = False
'    Call ARMDebugAssert(SetMetricField)
'    Exit Function
'
'  End Select
'
'  SetMetricField = True
'  Exit Function
'ErrorHandler:
'  SetMetricField = False
'  Call ARMDebugAssert(SetMetricField)    'stop in IDE
'  Exit Function
'
'End Function



'Private Function InsertLongInMetric( _
'                    ByVal eWhichMetricByArrayPos As ARM_Metrics_ArrayIndex_Enum, _
'                    ByVal lData As Long _
'                    ) As Boolean
'  Dim byTemp()      As Byte
'  Dim bContinue     As Boolean
'
'  On Error GoTo ErrorHandler
'  InsertLongInMetric = False
'
'  'byTemp = LongTo4Bytes_A(lData)
'  byTemp = LongTo4Bytes_B(lData)
'
'  Select Case eWhichMetricByArrayPos
'  Case eMetric1Index
'    bContinue = Pack4Bytes(byTemp(), mudtCmplxTxInstBuf.byMetric1())
'  Case eMetric2Index
'    bContinue = Pack4Bytes(byTemp(), mudtCmplxTxInstBuf.byMetric2())
'  Case eMetric3Index
'    bContinue = Pack4Bytes(byTemp(), mudtCmplxTxInstBuf.byMetric3())
'  Case eMetric4Index
'    bContinue = Pack4Bytes(byTemp(), mudtCmplxTxInstBuf.byMetric4())
'  Case eMetric5Index
'    bContinue = Pack4Bytes(byTemp(), mudtCmplxTxInstBuf.byMetric5())
'  Case eMetric6Index
'    bContinue = Pack4Bytes(byTemp(), mudtCmplxTxInstBuf.byMetric6())
'  End Select
'
'  InsertLongInMetric = bContinue
'  Exit Function
'ErrorHandler:
'  InsertLongInMetric = False
'  Call ARMDebugAssert(InsertLongInMetric)    'stop in IDE
'  Exit Function
'
'End Function
