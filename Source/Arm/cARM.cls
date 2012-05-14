VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cARM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------
'--- CLASS:         cARM
'---                Original design 8/2001
'---                by Phil Crowder & Mike Solano
'---
'--- PURPOSE:
'---    Private worker class.
'---    Encapsulates Application Response Measurement (ARM) API
'---
'--- INSTANCING:    Private
'---
'--- NOTATION:
'---        Tx = Transaction (throughout)
'--- HISTORY:
'--- * MSolano    8/1/2001  Initial Version SCR-ARM-1600
'------------------------------------------------------------
Option Explicit









'<<< SCR 22454 MSolano 3/25/02
#If DEBUGON Then
  Private Const msModule = "cARM"
#End If
'>>> SCR 22454 MSolano 3/25/02
#Const mbDebugMode = False

'--- For ARM_INIT, help client describe if they're passing a specific
'--- user or want one of the defaults (ignore user; use NT logon)
Public Enum enmUserType
    UserSpecified
    IgnoreUser
    UseNTUser
End Enum

'Private Const mlIGNORE_TX                     As Long = -1

Private mdicTxClassIDs                        As Scripting.Dictionary
Private mlAppID                               As Long          'Application ID from arm_init() API
Private mOverrideTxInstance                   As Long

'>>> SCR 31198 MSolano 3/25/03
'Private Type CorrelatorKernel_Type
'  byCorrelator(1 To 166)  As Byte
'End Type
'
'Private Type ARM_Start_DataBufferEmpty_Type
'  eFormat               As Long                     'const = 1
'  byFlags(1 To 4)       As Byte                     'Byte(1) = Correlator flags: abcd0000
'                                                    ' 1a=1 if app passing Correlator from parent Tx
'                                                    ' 1b=1 if app requesting agent generate Correlator
'                                                    ' 1C=1 if agent returning Correlator
'                                                    ' 1d=1 if app requesting agent to trace the Tx
'                                                    'Byte(2) = Metric Flags: abcdefg0
'                                                    ' a-f=1 if udtMetric(1 to 6) passed
'                                                    ' 2g=1 if sNzString32 passed
'  udtMetric(1 To 6)      As Start_Metric_Type
'  bysNzString32(1 To 32) As Byte                    'sNz = STRING *NOT* NULL TERMINATED; PAD W/ BLANKS
'  iCorrelatorLen         As Integer
'  udtCorrelatorKernel    As CorrelatorKernel_Type   '1st 2 bytes = length & ARE part of Correlator
'End Type
'<<< SCR 31198

Private msAppModePrefix                       As String       'SCR 36345 MSolano 5/29/03

Public Property Let AppID(lData As Long)
    mlAppID = lData
End Property

'------------------------------------------------------------
'--- PROCEDURE: Class_Initialize
'---
'--- PURPOSE:   Init application level variables
'---            for ARM API
'--- INPUTS / OUTPUTS / RETURNS: None
'--- HISTORY:
'--- * MSolano    8/3/2001  Initial Version SCR-ARM-1600
'------------------------------------------------------------
Private Sub Class_Initialize()
    
  Set mdicTxClassIDs = New Scripting.Dictionary
  msAppModePrefix = vbNullString                    'SCR 36345 MSolano 5/29/03
    
End Sub

'------------------------------------------------------------
'--- PROCEDURE: Class_Terminate
'---
'--- PURPOSE:
'---   Garbage collection
'---
'--- INPUTS / OUTPUTS / RETURN: None
'--- HISTORY:
'--- * MSolano  8/3/2001  Initial Version SCR-ARM-1600
'------------------------------------------------------------
Private Sub Class_Terminate()
    On Error Resume Next
       
    '--- GARBAGE COLLECTION FOR ENTIRE ARM COMPONENT
    Set mdicTxClassIDs = Nothing
        
End Sub
'------------------------------------------------------------
'--- PROCEDURE: StartTxInstance
'---
'--- PURPOSE:   Start a specific instance of a transaction
'---            of a named Tx class (type)
'--- INPUTS:
'---    sReadableTxClsName:
'---        unit of work denoting a Tx class
'---        named by the developer instrumenting the application
'---    lTxInstanceFlags:
'---        ARM API V2.0 - future use
'---        defaults to 0
'---    sTxInstanceData:
'---        buffer for additional data
'---        defaults to null string
'---        NOTE SPECIAL FORMAT IN ARM API DOCUMENT
'---        ARM 2.0 SDK USER's GUIDE
'---
'--- RETURN: the ID of a transaction instance
'---        used by StopTxInstance (lTxInstanceID)
'--- HISTORY:
'--- * MSolano    8/2/2001  Initial Version SCR-ARM-1600
'------------------------------------------------------------
Public Function StartTxInstance(sReadableTxClsName As String, _
                                Optional lTxInstanceFlags As Long = 0, _
                                Optional sTxInstanceData As String = "0" _
                                ) As Long
    
    'Dim udtEmpty                As ARM_Start_DataBufferEmpty_Type  'SCR 31198 MSolano 3/25/2003
    Dim lTxInstanceID           As Long                             'the ret. val.
    Dim lTxClassID              As Long
    Dim lTxInstanceDataSize     As Long
    Dim curTime1                As Currency                         'SCR 32099 Coleman 2/3/03
    Dim curTime2                As Currency                         'SCR 32099 Coleman 2/3/03
    
    '---------  Initialization & Test Basic Exit Conditions  -------
    '>>> SCR 31198 MSolano 3/25/2003
    On Error GoTo ErrorHandler
    StartTxInstance = glIGNORE_TX
    
    'udtEmpty = GetSimpleTxBuffer()
    'lTxInstanceDataSize = LenB(udtEmpty)
    lTxInstanceDataSize = 0
    
    If Not mdicTxClassIDs.Exists(sReadableTxClsName) Then
      Exit Function
    Else
      lTxClassID = mdicTxClassIDs.Item(sReadableTxClsName)
    End If
    If Not lTxClassID > 0 Then Exit Function
    'SCR 40974 - ARM - Remove logging of Pre- and Post- API Call transactions
    'Call D_bugLog(msModule, "StartTxInstance - Before", sReadableTxClsName, lTxClassID)
    '<<< SCR 31198
    '---------  End Initialization & Test Basic Exit Conditions  -------
    
        
    QueryPerformanceCounter curTime1                                 'SCR 32099 Coleman 2/3/03
    If gbARMInitialized Then
        '--- SCR 31198 vbNullString MSolano 3/25/2003
        lTxInstanceID = arm_start(lTxClassID, _
                                lTxInstanceFlags, _
                                vbNullString, _
                                lTxInstanceDataSize)
        'SCR 40974 - ARM - Remove logging of Pre- and Post- API Call transactions
        'Call D_bugLog(msModule, "StartTxInstance - After", sReadableTxClsName, lTxClassID, lTxInstanceID)
    '<<< SCR 22541  MSolano 3/29/2002
    ElseIf gbARMOverride Then
        mOverrideTxInstance = mOverrideTxInstance + 1
        lTxInstanceID = mOverrideTxInstance
        'SCR 40974 - ARM - Remove logging of Pre- and Post- API Call transactions
        'Call D_bugLog(msModule, "StartTxInstance (Override) - After", sReadableTxClsName, lTxClassID, lTxInstanceID)
    '>>> SCR 22541
    End If

    '<<< SCR 32099 Coleman 2/3/03
    QueryPerformanceCounter curTime2
    Call GUI_API_StartTime(curTime1, curTime2)
    '>>> SCR 32099 Coleman 2/3/03
    
    'Dbug.Print "arm_start:", lTxInstanceID

    'HUMMM - NOTE TO SELF: IS 0 ALLOWED *AND*
    'IS IT A GOOD TX? AN ARM'D WORKSTATION? W/ PROFILE WORKING?
    'Call ARMDebugAssert(lTxInstanceID > 0)
    
    #If mbDebugMode = True Then
        MsgBox "arm_start: " & lTxInstanceID '
    #End If

    If lTxInstanceID >= 0 Then
      StartTxInstance = lTxInstanceID
    Else
      StartTxInstance = glIGNORE_TX
    End If
    
    'StackTrace "cARM.StartTxInstance", False
Exit Function
ErrorHandler:
    '--- DOESN'T MAKE SENSE TO INTERRUPT USER FOR EACH
    '--- FAILURE OF AN INSTANCE OF A TRANSACTION
    'StackTrace "cARM.StartTxInstance", False
    StartTxInstance = glIGNORE_TX  'SCR 31198 MSolano 3/25/2003
    Exit Function

End Function

'------------------------------------------------------------
'--- PROCEDURE: StartCmplxTxInstance
'---
'--- PURPOSE:   Start a specific instance of a transaction
'---            of a named Tx class (type)
'--- INPUTS:
'---    sReadableTxClsName:
'---        unit of work denoting a Tx class
'---        named by the developer instrumenting the application
'---    lTxInstanceFlags:
'---        ARM API V2.0 - future use
'---        defaults to 0
'---    udtTxInstData:
'---        buffer for additional data
'---        defaults to null string
'---        NOTE SPECIAL FORMAT IN ARM API DOCUMENT
'---        ARM 2.0 SDK USER's GUIDE
'---
'--- RETURN: the ID of a transaction instance
'---        used by StopTxInstance (lTxInstanceID)
'--- HISTORY:
'--- * MSolano    8/2/2001  Initial Version SCR-ARM-1600
'--- * MSolano    8/12/2002 SCR 26339 Signature changed to also
'---                                  accomodate complex transactions
'------------------------------------------------------------
Public Function StartCmplxTxInstance( _
                           ByVal lTxClassID As Long, _
                           ByRef udtTxInstData As ARM_Start_DataBuffer_Type, _
                           ByVal lTxInstanceDataSize As Long, _
                  Optional lTxInstanceFlags As Long = 0 _
                  ) As Long
    
  Dim lTxInstanceID           As Long  'the ret. val.
  Dim curTime1                As Currency       'SCR 32099 Coleman 2/3/03
  Dim curTime2                As Currency       'SCR 32099 Coleman 2/3/03
  
  On Error GoTo ErrorHandler
  
  'StackTrace "cARM.StartCmplxTxInstance", True
  
  '--- test before
  #If gbModUnitTest = True Then
    Call ARMUnitTestAgentFlaggedThatItRetrunedACorrelator( _
            udtTxInstData.byFlags(glARM_Start_FlagsByteCorrelator), _
            False)
  #End If
  
  'SCR 32099 Coleman 2/3/03
  QueryPerformanceCounter curTime1
  If gbARMInitialized Then
      lTxInstanceID = arm_start(lTxClassID, _
                              lTxInstanceFlags, _
                              udtTxInstData, _
                              lTxInstanceDataSize)
  ElseIf gbARMOverride Then
      mOverrideTxInstance = mOverrideTxInstance + 1
      lTxInstanceID = mOverrideTxInstance
  End If
  
  '<<< SCR 32099 Coleman 2/3/03
  QueryPerformanceCounter curTime2
  Call GUI_API_StartTime(curTime1, curTime2)
  '>>> SCR 32099 Coleman 2/3/03
 
  '--- test after
  '#If gbModUnitTest = True Then
    'Call ARMUnitTestAgentFlaggedThatItRetrunedACorrelator( _
            udtTxInstData.byFlags(glARM_Start_FlagsByteCorrelator), _
            True)
  '#End If

  'HUMMM - NOTE TO SELF: IS 0 ALLOWED *AND*
  'IS IT A GOOD TX? AN ARM'D WORKSTATION? W/ PROFILE WORKING?
  Call ARMDebugAssert(lTxInstanceID > 0)
  
  #If mbDebugMode = True Then
      MsgBox "arm_start: " & lTxInstanceID '
  #End If

  If lTxInstanceID >= 0 Then
    StartCmplxTxInstance = lTxInstanceID
  Else
    StartCmplxTxInstance = glIGNORE_TX
  End If
  
  'StackTrace "cARM.StartCmplxTxInstance", False
  Exit Function
ErrorHandler:
    StartCmplxTxInstance = glIGNORE_TX
    '--- DOESN'T MAKE SENSE TO INTERRUPT USER FOR EACH
    '--- FAILURE OF AN INSTANCE OF A TRANSACTION
    'StackTrace "cARM.StartCmplxTxInstance", False
    Exit Function

End Function
'------------------------------------------------------------
'--- PROCEDURE: StopTxInstance
'---
'--- PURPOSE:   Stop a specific instance of a transaction
'---            ID'd by lTxInstanceID returned from StartTxInstance()
'--- INPUTS:
'---    sTxClassName
'---    lTxInstanceFlags
'---    sTxInstanceData
'---
'--- RETURN:    Boolean success / failure
'---
'--- HISTORY:
'--- * MSolano    8/2/2001  Initial Version SCR-ARM-1600
'--- * MSolano    12/17/2001 SCR-19804
'---              lTxInstanceStatus As enmARMTxStatus = enmARMTxStatusOk
'------------------------------------------------------------
Public Function StopTxInstance(lTxInstanceID As Long, _
                        Optional lTxInstanceStatus As enmARMTxStatus = enmARMTxStatusOk, _
                        Optional lTxInstanceFlags As Long = 0, _
                        Optional sTxInstanceData As String = "0") _
                    As Boolean
    
    Dim lRetVal                 As Long
    Dim lTxInstanceDataSize     As Long
    Const lAPI_FAILURE As Long = -1
    Dim curTime1                        As Currency     'SCR 32099 Coleman 2/3/03
    Dim curTime2                        As Currency     'SCR 32099 Coleman 2/3/03
    On Error GoTo ErrorHandler

    'StackTrace "cARM.StopTxInstance", True
    
    '--- MUST PASS lTxInstanceID - POSITIVE VAL
    '--- RETURNED FROM StartTxInstance
    Call ARMDebugAssert(lTxInstanceID > 0)
    
    'lTxInstanceDataSize = Len(sTxInstanceData)
    lTxInstanceDataSize = 0
    
    'SCR 32099 Coleman 2/3/03
    QueryPerformanceCounter curTime1
    '--- FORCE API CALL TO EXPLICITLY CHANGE lRetVal
    lRetVal = lAPI_FAILURE
    lRetVal = arm_stop(lTxInstanceID, _
                        lTxInstanceStatus, _
                        lTxInstanceFlags, _
                        sTxInstanceData, _
                        lTxInstanceDataSize)
    
    '<<< SCR 32099 Coleman 2/3/03
    QueryPerformanceCounter curTime2
    Call GUI_API_StopTime(curTime1, curTime2)
    '>>> SCR 32099 Coleman 2/3/03
    
    '--- ARM 2.0 API 0 = success
    Call ARMDebugAssert(lRetVal = 0)
    If lRetVal = 0 Then
        StopTxInstance = True
    Else
        StopTxInstance = False
    End If
    
    #If mbDebugMode = True Then
        MsgBox "arm_stop: " & lRetVal '
    #End If
                                
  'StackTrace "cARM.StopTxInstance", False
Exit Function
ErrorHandler:
    '--- DOESN'T MAKE SENSE TO INTERRUPT USER FOR EACH
    '--- FAILURE OF AN INSTANCE OF A TRANSACTION
    'StackTrace "cARM.StopTxInstance", False
    Exit Function

End Function

Public Function StopCmplxTxInstance( _
                      lTxInstanceID As Long, _
                      lTxInstanceStatus As Long, _
                      udtTxInstanceData As ARM_Start_DataBuffer_Type, _
             Optional lTxInstanceFlags As Long = 0 _
              ) As Boolean
    
    Dim lRetVal                         As Long
    Dim lTxInstanceDataSize             As Long
    Dim curTime1                        As Currency     'SCR 32099 Coleman 2/3/03
    Dim curTime2                        As Currency     'SCR 32099 Coleman 2/3/03
    Const lAPI_FAILURE                  As Long = -1
    
    On Error GoTo ErrorHandler
    StopCmplxTxInstance = False
    
    'StackTrace "cARM.StopCmplxTxInstance", True
    
    '--- MUST PASS lTxInstanceID - POSITIVE VAL
    '--- RETURNED FROM StartTxInstance
    Call ARMDebugAssert(lTxInstanceID > 0)
    
    lTxInstanceDataSize = LenB(udtTxInstanceData)
    
    '--- FORCE API CALL TO EXPLICITLY CHANGE lRetVal
    lRetVal = lAPI_FAILURE
    'SCR 32099 Coleman 2/3/03
    QueryPerformanceCounter curTime1
    lRetVal = arm_stop( _
                  lTxInstanceID, _
                  lTxInstanceStatus, _
                  lTxInstanceFlags, _
                  udtTxInstanceData, _
                  lTxInstanceDataSize)
    '<<< SCR 32099 Coleman 2/3/03
    QueryPerformanceCounter curTime2
    Call GUI_API_StopTime(curTime1, curTime2)
    '>>> SCR 32099 Coleman 2/3/03
    
    '--- ARM 2.0 API 0 = success
    StopCmplxTxInstance = (lRetVal = 0)
    Call ARMDebugAssert(StopCmplxTxInstance)
    
    #If mbDebugMode = True Then
        MsgBox "arm_stop: " & lRetVal '
    #End If
                                
  'StackTrace "cARM.StopCmplxTxInstance", False
  Exit Function
ErrorHandler:
    StopCmplxTxInstance = False
    'StackTrace "cARM.StopCmplxTxInstance", False
    Exit Function

End Function

'------------------------------------------------------------
'--- PROCEDURE:     InitTxClass
'--- PURPOSE:
'---    Return unique ID of what ARM API terms a "transaction class":
'---    a body of code, arbitrarily defined, representing some
'---    unit of work, and named by the developer ARMing the app.
'---    (the Tx "class" is contrasted w/ a Tx "instance",
'---    or runtime episode of the named Tx class)
'---
'--- INPUTS:
'---    sTxClassName: name of "class" (type) of the transaction
'---                  (vs. a specific instance of that Tx)
'---    sTxDetail: Transaction detail - ARM V2.0 API SPEC:
'---        * 128 byte free form text, incl. null terminator
'---        * if contents in subsequent calls change using same
'---          AppID / sTxClassName pair, but sTxDetail changes,
'---          new sTxDetail *IGNORED*
'---    lTxClassFlags: ARM API V2.0 - future use
'---    vTxClassData:
'---        Additional data that can be passed to Tx instance
'---        API calls arm_start(), arm_update(), arm_stop()
'---
'--- RETURN:    ID for Transaction Class - long
'--- HISTORY:
'--- * MSolano  8/1/2001  Initial Version SCR 1600
'--- * MSolano  8/8/2002  SCR 26339
'---            Changed signature, RetVal
'------------------------------------------------------------
'>>> SCR 26339 NEW / PREV SIGNATURE
Public Function InitTxClass( _
                               sTxClassName As String, _
                      Optional vTxClassData As Variant = "0", _
                      Optional lTxClassDataSize As Long = 0, _
                      Optional sTxClassDetail As String = "0", _
                      Optional lTxClassFlags As Long = 0 _
                      ) As Long
'<<< SCR 26339
    
    Dim lTxClassID            As Long
    Dim bComplxTx             As Boolean
    Dim udtTxClassData        As ARM_GetID_DataBuffer_Type
    Dim curTime1              As Currency                           'SCR 32099 Coleman 2/3/03
    Dim curTime2              As Currency                           'SCR 32099 Coleman 2/3/03
    Dim sTxClassNamePrefixed  As String                             'SCR 36345 MSolano 5/29/2003
    
    On Error GoTo ErrorHandler
    
    'Call ARMDebugAssert(False)
        
    '<<< SCR 26339 MSolano 8/8/2002
    'Dbug.Assert Not (TypeOf vTxClassData Is ARM_GetID_DataBuffer_Type)  'for debugging Complex Tx
    If TypeOf vTxClassData Is ARM_GetID_DataBuffer_Type Then
      udtTxClassData = vTxClassData
      bComplxTx = True
        
    ElseIf VarType(vTxClassData) = vbString Then
      vTxClassData = Trim(vTxClassData)
    End If
    '>>> SCR 26339
       
    
    
      '--- SEE IF Tx CLASS ALREADY INIT'd & IN DICTIONARY
      If Not mdicTxClassIDs.Exists(sTxClassName) Then
      
          'SCR 32099 Coleman 2/3/03
          QueryPerformanceCounter curTime1
          
          sTxClassNamePrefixed = msAppModePrefix & sTxClassName     'SCR 36345 MSolano 5/29/2003
          
          '<<< SCR 22541 if ARM init'd, use API
          If gbARMInitialized Then
            If Not bComplxTx Then
                lTxClassID = arm_getid( _
                                    mlAppID, _
                                    sTxClassNamePrefixed, _
                                    sTxClassDetail, _
                                    lTxClassFlags, _
                                    vTxClassData, _
                                    lTxClassDataSize)
            Else
                lTxClassID = arm_getid( _
                                    mlAppID, _
                                    sTxClassNamePrefixed, _
                                    sTxClassDetail, _
                                    lTxClassFlags, _
                                    udtTxClassData, _
                                    lTxClassDataSize)
            End If
          
          ElseIf gbARMOverride Then
              '--- Fake ARM API to force DLog out
              mOverrideTxInstance = mOverrideTxInstance + 1
              lTxClassID = mOverrideTxInstance
          End If
          '>>> SCR 22541
          
          '<<< SCR 32099 Coleman 2/3/03
          QueryPerformanceCounter curTime2
          Call GUI_API_InitTime(curTime1, curTime2)
          '>>> SCR 32099 Coleman 2/3/03
          
          mdicTxClassIDs.Add sTxClassName, lTxClassID
      Else
          lTxClassID = mdicTxClassIDs.Item(sTxClassName)
      End If
    
      
    '<<< SCR 22454 MSolano 3/25/02
    #If DEBUGON Then
    'RWalker 12/19/2002 Differentiate class type
      If Not bComplxTx Then
        Call D_bugLog(msModule, "InitTxClass", sTxClassName, lTxClassID)
      Else
        Call D_bugLog(msModule, "InitCplxTx", sTxClassName, lTxClassID)
      End If
    #End If
    '>>> SCR 22454

    '--- ARM 2.0 API: ERR IF arm_getid() < 0
    If lTxClassID > 0 Then
        InitTxClass = lTxClassID    'SCR 26339 MSolano 8/8/2002
    Else
        InitTxClass = glIGNORE_TX   'SCR 26339 MSolano 8/8/2002
    End If
    
    Call ARMDebugAssert(lTxClassID > 0)
    
    #If mbDebugMode = True Then
        MsgBox "arm_getid sTxClassName / lTxClassID: " & sTxClassName & "/" & CStr(lTxClassID)
    #End If

Exit Function
ErrorHandler:
    InitTxClass = glIGNORE_TX   'SCR 26339 MSolano 8/8/2002
    
    '>>> SCR 22128 MSolano 3/13/2002
    '--- circular references not allowed
    '--- Report error only once
    'If Not bErrReported Then
    '    bErrReported = True
    '    goShared.ShowVBError Err.Number, Err.Description, sRoutine, App.Title, vbExclamation
    'End If
    '<<< SCR 22128
    Exit Function
    Resume
End Function

'>>> SCR 36345 MSolano 5/29/2003
Public Property Let AppModePrefix(sVal As String)
  On Error GoTo ErrorHandler

  msAppModePrefix = sVal & " - "

  Exit Property
ErrorHandler:
  msAppModePrefix = vbNullString
  Call ARMDebugAssert(Err.Number = 0)     'stop in IDE
  Exit Property

End Property

Public Property Get AppModePrefix() As String
  On Error GoTo ErrorHandler

  AppModePrefix = msAppModePrefix

  Exit Property
ErrorHandler:
  AppModePrefix = vbNullString
  Call ARMDebugAssert(Err.Number = 0)     'stop in IDE
  Exit Sub

End Property
'<<< SCR 36345

'>>> SCR 31198 MSolano 3/25/03
'Private Function GetSimpleTxBuffer() As ARM_Start_DataBufferEmpty_Type
'  Dim udtEmpty As ARM_Start_DataBufferEmpty_Type
'  On Error GoTo ErrorHandler
'
'  udtEmpty.eFormat = 1
'  udtEmpty.iCorrelatorLen = 168
'  GetSimpleTxBuffer = udtEmpty
'
'  Exit Function
'ErrorHandler:
'  Call ARMDebugAssert(Err.Number = 0)
'  Exit Function
'End Function
'<<< SCR 31198


