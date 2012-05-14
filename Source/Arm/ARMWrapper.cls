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
Attribute VB_Name = "cARMWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------
'--- CLASS:         cARMWrapper
'---
'--- PURPOSE:   Wrapper class for cARM worker class
'---    * cARMWrapper tracks which transactions are ARM'd
'---    * cARMWrapper is implemented as Singleton:
'---        - Creatable only by cARMFactory
'---        - Only single object instance allowed:
'---            * ARM API initialized only once
'---            * ARM App ID maintained
'---
'--- INSTANCING:    PublicNotCreatable
'---                (enforces singleton)
'--- NOTATION:
'---        Tx = Transaction (throughout)
'--- HISTORY:
'--- * MSolano    8/7/2001  Initial Version SCR-ARM-1600
'------------------------------------------------------------

Option Explicit

'<<< SCR 22454 MSolano 3/25/02
#If DEBUGON Then
  Private Const msModule = "cARMWrapper"
#End If
#Const mbDebugMode = False
'>>> SCR 22454

Private mlTimeStartMS   As Long   'milliseconds aft midnight
Private mlTimeStopMS    As Long   'milliseconds aft midnight
Private mlTimeDiffMS    As Long   'milliseconds delta

Public Enum enmArmUserLogMode
    ARMUserAsCWUserAndComputername = 0  'DEFAULT use CW app user & NT computer name
    ARMUserAsCWUser                     'use CW app user
    ARMUserAsNetworkUser                'NT user
    ARMUserAsNetworkUserAndCWUser       'NT user & CW app user
End Enum

'--- MSolano 12/17/2001 SCR-19804
'--- VALUES ARE PER ARM 2.0 API
Public Enum enmARMTxStatus
    enmARMTxStatusOk = 0
    enmARMTxStatusAborted = 1
    enmARMTxStatusFailed = 2
End Enum


Public Enum CorrelatedFamilies_Enum
  eNoFamily
  eAnyFamily
  eModalFamily
  eSelectPatientFamily
  eSignEncounterFamily
  eTestFamily
  eTestFamily2
  eTestFamily3
End Enum

Public Enum enmARMStopWhichTx  '--- SCR 22142 MSolano 3/17/02
  enmARMStop_AllCurrTx = 0
  enmARMStop_MeOnly = 1
End Enum

'--- SCR 22241 MSolano 3/16/02
Public Enum enmCHCSIIDlg
    enmCHCSIIDlgDetect_None = 0  'none detected, rather than asserting none existed
    enmCHCSIIDlgDetect_VBMsgBx = 1
    enmCHCSIIDlgDetect_CustDlg = 2
    enmCHCSIIDlgDetect_WinAPI = 3
End Enum

'Private Const mlIGNORE_TX                     As Long = -1

'Private moARM                                 As cARM
Private moCmplxTxs                            As cCmplxTransactions   'SCR 26339 MSolano 8/8/2002

Private mdictReadableTxClassName              As Scripting.Dictionary
Private mdictCachFarTxIDs                     As Scripting.Dictionary
Private mdictNearTxsByHandle                  As Scripting.Dictionary  'SCR 22034 MSOLANO 3/8/2002
Private mdictNearTxsByID                      As Scripting.Dictionary  'SCR 22034 MSOLANO 3/8/2002

'--- SCR 22541 changed mbARMInitialized to global MSolano 3/29/2002
'Private mbARMInitialized                      As Boolean      'flags if ARM app init'd

Private msAppName                             As String       'Application name

Private msAppUserName                         As String       'Name of the Application user
Private menmAppUserLogMode                    As enmArmUserLogMode   'What of several possible user names to
                                                        'consider the application user for ARM API
Private mlAppFlags                            As Long         'Reserved = 0
Private msAppData                             As String       'Reserved = 0
Private mlAppDataSize                         As Long         'Reseverd = 0

Private mlAppID                               As Long         'Application ID from arm_init() API

'SCR-36503 RWalker 7/8/03 ARM: CORE - Instrument transactions for timer and multifunction processes
Private mAppUnitTestLOG                       As Boolean

'SCR-48611 RWalker 1/23/2004 Transactions removed
'SCR 22033 Msolano 3/8/2002
'SCR 22034 Msolano 3/16/2002
'SCR 22241 Msolano 3/16/2002

'--- ARM API V2.0 MAGIC NUMBERS
Private Const lARM_API_USER_NAME_LEN = 127  '128 incl. null terminator
Private Const lARM_API_APP_NAME_LEN = 127   '128 incl. null terminator
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
    On Error GoTo AFTER_DEBUG
    #If DEBUGON Then
        Dim DR As DebugRec
        Set goDL = GetObject("", "DLog.cDebugLog")
        
        goDL.debugStart
        
        '--- SCR 22541  MSolano 3/29/2002
        If goDL.ARMOverride Then gbARMOverride = True
        
        '--- COL HEADINGS / EMPTY DATA
        goDL.DebugLog Exe, msModule, "Class_Initialize", "", DR, _
                  "DATE", Format(Now, "mm/dd/yy"), _
                  "TIME", Format(Now, "Hh:Nn:Ss"), _
                  "MILLISEC", "", _
                  "CLASS NAME", "", _
                  "CLASS ID", "", _
                  "INSTANCE ID", "", _
                  "INSTANCE STATUS (0=Gd;1=ABORT;2=FAIL)", ""
    #End If

AFTER_DEBUG:
    On Error GoTo ErrorHandler
    Set mdictReadableTxClassName = New Scripting.Dictionary
    Set mdictCachFarTxIDs = New Scripting.Dictionary
    Set mdictNearTxsByHandle = New Scripting.Dictionary
    Set mdictNearTxsByID = New Scripting.Dictionary
    
    Set goARMAPI = New cARM
    
    Set moCmplxTxs = New cCmplxTransactions    'SCR 26339 MSolano 8/8/2002

'SCR 48872 RWalker 2/4/2004 Added new parameter containing the INI Path.
'Since we will not have this value until the InitARM method is executed,
'move the next 2 lines into that procedure.
'2/4/2004    'SCR 32099 Coleman 2/3/03
'2/4/2004    ARM_GUIUnitTestCheck moCmplxTxs
'2/4/2004
'2/4/2004    'SCR-36503 RWalker 7/8/03 ARM: CORE - Instrument transactions for
'2/4/2004    'timer and multifunction processes
'2/4/2004    AppUnitTestLOG = ARM_LOGUnitTestCheck
        
    '--- ARM 2.0 - Default val's for variables reserved for future
    mlAppFlags = 0
    msAppData = "0"
    mlAppDataSize = 0
ErrorHandler:
    Exit Sub
End Sub

'<<< SCR 24288 MSolano 8/13/02
Private Function NumAllCurrentInstancesAnyTxType() As Long
  On Error GoTo ErrorHandler
  
  'SCR 28539 MSolano 10/15/02 added moCmplxTxs.CountCurrentInstances
  NumAllCurrentInstancesAnyTxType = _
        mdictNearTxsByHandle.Count + mdictNearTxsByID.Count _
        + mdictCachFarTxIDs.Count _
        + moCmplxTxs.CountCurrentInstances
  
  Exit Function
ErrorHandler:
  'On error, assume "pessimistic" - i.e. try to stop curr Tx's
  NumAllCurrentInstancesAnyTxType = 1
  Call ARMDebugAssert(Err.Number = 0)    'stop in IDE
  Exit Function
  
End Function
'>>> SCR 24288









'--- Read only property published to type lib
Public Property Get IgnoreTx() As Long
  IgnoreTx = glIGNORE_TX
End Property
'------------------------------------------------------------
'--- PROCEDURE: ARM_InitAllTransactionsForARM
'---
'--- PURPOSE:
'---  MESSAGE MAP REGISTERING ARM TRANSACTIONS
'---  THAT ARM ITSELF IS RESPONSIBLE FOR
'---  (ARM Tx's INSTRUMENTED IN ARM ITSELF)
'---
'--- ASSUMPTION:
'---  CALLED ONCE PER PROJECT / COMPONENT
'---
'--- HISTORY:
'--- * MSolano    3/8/2002 Initial Version
'------------------------------------------------------------
Public Function ARM_InitAllTransactionsForARM() As Boolean
    Static bDoneThis    As Boolean
    Dim oARMFactory     As New ARM.cARMFactory
    On Error GoTo ErrorHandler
    
    ARM_InitAllTransactionsForARM = False
        
    '--- Init Tx classes only once regardless of where
    '--- ARM_InitAllTransactionsForARM called from or how many times
    If bDoneThis Then
        Exit Function
    Else
        bDoneThis = True
    End If
    
    '*-- InitTransaction will post INIT times to GUI, so go ahead and show
    '*-- it here. Make sure gcurFreq is set so no division by zero errors...
    If gbGUIUnitTest Then
        QueryPerformanceFrequency gcurFreq
        frmCorrTrace.Show
    End If
    
    '---------------------------------------------------
    '   MESSAGE MAP OF TRANSACTIONS TO ARM
    '---------------------------------------------------

    'Call ARMDebugAssert(False)

    'SCR 44601 RWalker 12/04/2003 Only log and generate transactions in DebugLog mode
    'SCR 48611 RWalker 1/23/2004 Transactions Removed
    'If gARM_LOGDebugMode Then
    'End If
    
    '>>> SCR 33025  MSolano 3/28/03
    'Call InitCmplxTxCls(gsARMTx_ARMMsgBxObjInit, , , _
                        "Instance ID", eARM_NumericID32, _
                        "Instance Count", eARM_NumericID32) 'SCR 33025  MSolano 2/26/2003
    'Call InitCmplxTxCls(gsARMTx_ARMMsgBxObjTerm, , , _
                        "Instance ID", eARM_NumericID32, _
                        "Instance Count", eARM_NumericID32) 'SCR 33025  MSolano 2/26/2003
    
    'Call InitCmplxTxCls(gsARMTx_ARMMsgBxMeth, , , _
                        "Instance ID", eARM_NumericID32, _
                        "Instance Count", eARM_NumericID32) 'SCR 32152, 33025  MSolano 2/26/2003
    'Call InitTransaction(gsARMTx_ARMMsgBxMeth)              'SCR 32152, 33025  MSolano 2/26/2003
    '<<< SCR 33025
    
    '--- SCR 32152 MSolano 2/7/2003
    'Call InitCmplxTxCls( _
              gsARMTx_DoEvents, , "Module.Proc.Line", _
              "StrOverflow1", eARM_String8, _
              "StrOverflow2", eARM_String8, _
              "StrOverflow3", eARM_String8, _
              "StrOverflow4", eARM_String8, _
              "StrOverflow5", eARM_String8, _
              "StrOverflow6", eARM_String8)
    
    #If gbModUnitTest = True Then
      Call ARM_DoUnitTests
    #End If
    
    ARM_InitAllTransactionsForARM = True
    Exit Function
ErrorHandler:
    Call ARMDebugAssert(Err.Number = 0)
    Exit Function
End Function

'------------------------------------------------------------
'--- PROCEDURE: Class_Terminate
'---
'--- PURPOSE:
'---   Shut down ARM API
'---   Garbage collection for the entire ARM component.
'---   Other object(s) are support infrastructure.
'---
'--- INPUTS / OUTPUTS / RETURN: None
'--- HISTORY:
'--- * MSolano  8/3/2001  Initial Version SCR-ARM-1600
'------------------------------------------------------------
Private Sub Class_Terminate()
    Dim bEnd As Boolean
    On Error Resume Next
    
    bEnd = EndARMApp()
    
    '--- GARBAGE COLLECTION FOR ENTIRE ARM COMPONENT
    Set goARMAPI = Nothing
    Set mdictReadableTxClassName = Nothing
    Set mdictCachFarTxIDs = Nothing
    Set mdictNearTxsByHandle = Nothing
    Set mdictNearTxsByID = Nothing
    
    Set moCmplxTxs = Nothing   'SCR 26339 MSolano 8/8/2002
    
    'Set goShared = Nothing 'SCR 22128 MSolano 3/13/2002 circular references not allowed
    
    Set goARMApp = Nothing  'singleton, which is terminating right here
End Sub
Public Property Let AppName(sData As String)
   sData = Left(Trim(sData), lARM_API_APP_NAME_LEN)
   msAppName = sData
End Property

Public Property Get AppName() As String
    AppName = msAppName
End Property
'------------------------------------------------------------
'--- PROCEDURE: UserName
'--- ASSUMPTIONS:
'---    If UserLoggingMode property used to modify
'---    the ARM user, that property must be used
'---    before UserName property
'--- PURPOSE:   What is considered the ARM API user
'---            is settable using UserLoggingMode property
'---
'--- HISTORY:
'--- * MSolano    8/3/2001  Initial Version SCR-ARM-1600
'------------------------------------------------------------
Public Property Let UserName(sCWAppUser As String)
    'Dim sCompName   As String
    Dim sUserName   As String
    On Error GoTo ErrorHandler
        
    sCWAppUser = Trim(sCWAppUser)
    
    Select Case menmAppUserLogMode
    Case ARMUserAsCWUserAndComputername
        '--- APPEND WINDOWS COMPUTER NAME
        sUserName = sCWAppUser & "-" & gsComputerName()
    
    Case ARMUserAsCWUser
        sUserName = sCWAppUser
    
    Case ARMUserAsNetworkUser
        'TODO - IMPLEMENT API CALL TO GETUSER
        sUserName = sCWAppUser
    
    Case ARMUserAsNetworkUserAndCWUser
        'TODO - IMPLEMENT API CALL TO GETUSER  & APPEND
        sUserName = sCWAppUser
    End Select
   
   sUserName = Left(sUserName, lARM_API_USER_NAME_LEN)
    
    msAppUserName = sUserName

Exit Property
ErrorHandler:
    Exit Property
End Property
Public Property Get UserName() As String
    UserName = msAppUserName
End Property
Public Property Get AppFlags() As Long
    AppFlags = mlAppFlags
End Property
Public Property Get AppData() As String
    AppData = msAppData
End Property
Public Property Get AppDataSize() As Long
    AppDataSize = mlAppDataSize
End Property

Public Property Get AppID() As Long
    AppID = mlAppID
End Property

'------------------------------------------------------------
'--- PROCEDURE:     InitARMApp
'--- PURPOSE:       Initialize ARM API
'---                ENFORCE THAT IT'S ONLY INIT'd ONCE
'--- ASSUMPTIONS / KEY NOTES:
'---    V2.0 of the ARM API
'---    Arm_Init() returns a different mlAppID
'---    based on unique combination of
'---    application & application user
'---
'--- INPUTS:
'---    sUserName:  can be
'---                CW user (application user) or NT Login
'---                See above note in ASSUMPTIONS.
'---
'--- OUTPUT:        Success - boolean
'--- HISTORY:
'--- * MSolano    8/1/2001  Initial Version SCR-ARM-1600
'------------------------------------------------------------
Public Function InitARMApp(sAppTitle As String, _
                           sUserName As String, _
                           enmArmUserLogMode As enmArmUserLogMode, _
                           sINIPath) _
                    As Boolean

'SCR 48872 RWalker 2/4/2004 Added new parameter containing the INI Path.

    Static bAlreadyInitialized  As Boolean
    Static bErrReported         As Boolean
    
    On Error GoTo ErrorHandler
    
    '--- Semaphore - stop re-entrance, re-initialization
    If bAlreadyInitialized Then
        InitARMApp = gbARMInitialized
        Exit Function
    Else
        bAlreadyInitialized = True
    End If
    
    Me.AppName = sAppTitle
    
    '--- UserLoggingMode property set must preceede
    '--- UserName property set
    Me.UserLoggingMode = enmArmUserLogMode
    Me.UserName = sUserName
    
'SCR 48872 RWalker 2/4/2004 Added new parameter containing the INI Path.
    gsINIPath = sINIPath

'Since we didn't have this value until now, we moved the next 2 lines into here.
    ARM_GUIUnitTestCheck moCmplxTxs
    AppUnitTestLOG = ARM_LOGUnitTestCheck
    
    mlAppID = arm_init(msAppName, _
                        msAppUserName, _
                        mlAppFlags, _
                        msAppData, _
                        mlAppDataSize)
    
    goARMAPI.AppID = mlAppID
    
    '--- ARM API 2.0: error if arm_init() < 0
    '--- that val CAN be passed to arm_getid, which functions as NULL operation
    
    Call ARMDebugAssert(mlAppID > 0)
    #If mbDebugMode = True Then
        MsgBox "arm_init: " & mlAppID '
    #End If
    
    If mlAppID < 0 Then
        '>>> SCR 22541 MSolano 3/29/2002
        gbARMInitialized = False
        #If DEBUGON Then
          Call D_bugLog(msModule, "InitARMApp", "ARM API DID NOT INITIALIZE")
        #End If
        '<<< SCR 22541
    Else
        gbARMInitialized = True
    End If
        
    InitARMApp = gbARMInitialized
Exit Function
ErrorHandler:
    '--- PROBABLY WORKSTATION ISN'T CONFIGURED FOR ARM
    '--- ARM API (LibARM32.dll) PROB NOT INSTALLED
    '<<< SCR 22541 MSolano 3/29/2002
    gbARMInitialized = False
    InitARMApp = gbARMInitialized
    
    #If DEBUGON Then
      Call D_bugLog(msModule, "InitARMApp", "ARM API DID NOT INITIALIZE")
    #End If
    '>>> SCR 22541
    
    Exit Function
End Function
'------------------------------------------------------------
'--- PROCEDURE: EndARMApp
'---
'--- PURPOSE:
'---   Shut down ARM API
'---
'--- INPUTS / OUTPUTS / RETURN: None
'--- HISTORY:
'--- * MSolano  8/3/2001  Initial Version SCR-ARM-1600
'------------------------------------------------------------
Public Function EndARMApp() As Boolean
    Dim lRetVal             As Long
    Dim lAppDataSize        As Long
    Static bDoneThis        As Boolean
    On Error Resume Next
    
    '<<< SCR 22454 MSolano 3/25/02
    #If DEBUGON Then
    If Not goDL Is Nothing Then
      Dim DR As DebugRec  '- starts timer
      goDL.DebugLog Exe, msModule, "EndARMApp", "", DR, _
              "TIME", Format(Now, "mm/dd/yy"), _
              "", "", _
              "", "", _
              "", "", _
              "", "", _
              "", "", _
              "", ""
      goDL.debugStop
      Set goDL = Nothing
    End If
    #End If
    '>>> SCR 22454
    
    '--- Leave if ARM API never initialized
    '--- e.g. workstation not configured for ARM (LibARM32.dll not installed)
    If Not gbARMInitialized Then
        '--- Important to enforce singleton
        Set goARMApp = Nothing
        Exit Function
    End If
        
    '--- Semaphore - avoid re-entrance
    '--- can be called from exposed method, or class terminate
    If bDoneThis Then
        '--- if second call, assume first call
        '--- either worked or err was handled
        EndARMApp = True
        Exit Function
    Else
        bDoneThis = True
    End If
    
    lAppDataSize = Len(Trim(msAppData))  'bmstop
    
    lRetVal = arm_end(mlAppID, _
                      mlAppFlags, _
                      msAppData, _
                      lAppDataSize)
    
    '--- ARM 2.0 API - Error if lRetVal < 0
    Call ARMDebugAssert(lRetVal >= 0)
    #If mbDebugMode = True Then
        MsgBox "arm_end: " & lRetVal '
    #End If
        
    If lRetVal >= 0 Then
      EndARMApp = True
    Else
      EndARMApp = False
    End If
    
    '--- Important to enforce singleton
    Set goARMApp = Nothing
End Function
'------------------------------------------------------------
'--- PROCEDURE: UserLoggingMode (Prop Let)
'---
'--- PURPOSE:   Mode setting for what to consider the ARM API user:
'---            Options documented in enmArmUserLogMode enumeration.
'--- INPUTS:    Mode option
'--- HISTORY:
'--- * MSolano  8/1/2001  Initial Version SCR-ARM-1600
'------------------------------------------------------------
Public Property Let UserLoggingMode(enmData As enmArmUserLogMode)
    menmAppUserLogMode = enmData
End Property
Public Property Let AppUnitTestLOG(inData As Boolean)
'SCR-36503 RWalker 7/8/03 ARM: CORE - Instrument transactions for
'timer and multifunction processes
    mAppUnitTestLOG = inData
End Property
Public Property Get AppUnitTestLOG() As Boolean
'SCR-36503 RWalker 7/8/03 ARM: CORE - Instrument transactions for
'timer and multifunction processes
    AppUnitTestLOG = mAppUnitTestLOG
End Property
Public Property Get UserLoggingMode() As enmArmUserLogMode
    UserLoggingMode = menmAppUserLogMode
End Property

'------------------------------------------------------------
'--- PROCEDURE: InitTransaction
'---
'--- PURPOSE:
'---        Initialize transaction classes to be ARMd.
'---
'--- DOCUMENTATION:
'---    ARM documentation in VSS:
'---    $/CHCS2 INC2/SOURCE/ARM/ReadMe - Using the  CHCSII ARM component.doc
'---
'--- INPUTS:
'---        * sReadableTxClsName
'---          If ARMing a transaction at an atomic level, the
'---          developer can choose a readable name for it, in which case
'---          we don't need sSomeTxHandle
'---        * sSomeTxHandle:
'---          Some unreadable handle - If ARMing a transaction launched from common
'---          code area, developer may have an unreadable TxID like an NCID
'--- OUTPUT:  None
'--- RETURN:  success (boolean)
'--- HISTORY:
'--- * MSolano    8/7/2001  Initial Version SCR-ARM-1600
'------------------------------------------------------------
Public Function InitTransaction( _
                             sReadableTxClsName As String, _
                    Optional sSomeTxHandle As String = "" _
                    ) As Boolean
                    
    Dim bARMd               As Boolean
    
    On Error GoTo ErrorHandler
    InitTransaction = True  'assume happy
    If Not ApplicationARMd Then Exit Function
    
    If sSomeTxHandle = "" Then
        '--- There's no second unreadable handle, there's just the readable one
        sSomeTxHandle = sReadableTxClsName
    Else
        '--- There are two different handles:
        '--- a readable one sReadableTxClsName & an unreadable one sSomeTxHandle
        '--- Do Nothing
    End If
    
    '--- Init each class only once
    If mdictReadableTxClassName.Exists(sSomeTxHandle) Then
        Exit Function
    End If
        
    'mdictReadableTxClassName.Add Key, Item
    mdictReadableTxClassName.Add sSomeTxHandle, sReadableTxClsName
    'Dbug.Print "cARMWrapper.InitTransaction - mdictReadableTxClassName key, val: ", sSomeTxHandle, mdictReadableTxClassName.Item(sSomeTxHandle) 'BMStop
        
    bARMd = goARMAPI.InitTxClass(sReadableTxClsName)
    
    InitTransaction = bARMd
    
    Exit Function
ErrorHandler:
    InitTransaction = False
    Call ARMDebugAssert(InitTransaction)
    Exit Function
End Function




'>>> SCR 26339 MSolano 8/8/2002
Public Function InitCmplxTxCls( _
                         sReadableTxClsName As String, _
                Optional sSomeTxHandle As String = vbNullString, _
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
                         ) As Boolean
  
  Dim curTime1  As Currency         'SCR 32099 Coleman 2/3/03
  Dim curTime2  As Currency         'SCR 32099 Coleman 2/3/03
  
  On Error GoTo ErrorHandler
  
  Call ARMDebugAssert(((moCmplxTxs Is Nothing) = False))
  
  'SCR 32099 Coleman 2/3/03
  QueryPerformanceCounter curTime1
  InitCmplxTxCls = moCmplxTxs.InitCmplxTxCls( _
                                sReadableTxClsName, _
                                sSomeTxHandle, _
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
  '<<< SCR 32099 Coleman 2/3/03
  QueryPerformanceCounter curTime2
  GUI_ARM_InitTime curTime1, curTime2
  '>>> SCR 32099 Coleman 2/3/03
  Exit Function
ErrorHandler:
  InitCmplxTxCls = False
  Exit Function
End Function
'<<< SCR 26339


'------------------------------------------------------------
'--- PROCEDURE: StartTx
'---
'--- PURPOSE:
'---    Start a transaction instance, if transaction class is ARM'd
'---
'--- DOCUMENTATION:
'---    ARM documentation in VSS:
'---    $/CHCS2 INC2/SOURCE/ARM/ReadMe - Using the  CHCSII ARM component.doc
'---
'--- ASSUMPTIONS:
'---    ARMWrapper tracks if API is initialized (workstation ARM'd), otherwise exits.
'---    SOME TRANSACTIONS ARE ARM'd AT COMMON POINTS
'---    E.G. CORE gobjServers.StartCWOLEServer(oSvr.sAppNCID)
'---    ONLY SOME OF THE ACTIONS (TRANSACTIONS) LAUNCHED
'---    AT THESE COMMON CODE POINTS ARE ARM'D
'---
'--- INPUTS:
'---    External name of a transaction class
'---    May be an unreadable identifier like an NCID
'---    E.G. common code like Core's
'---    gobjServers.StartCWOLEServer(oSvr.sAppNCID)
'---    This name is arbitrary - chosen by developer as
'---    s/he defines a transaction class to be ARM'd
'---
'--- RETURN:
'---    Transaction instance ID USED BY StopTx(),
'---    or FLAG FOR StopTx TO IGNORE THIS TRANSACTION.
'---
'--- HISTORY:
'--- * MSolano    8/7/2001  Initial Version SCR-ARM-16000
'--- * MSolano    3/8/2002  SCR 22034
'------------------------------------------------------------
Public Function StartTx(sSomeTxHandle As String) As Long

    Dim sReadableTxClsName        As String
    Dim lTxInstID                 As Long
    Dim lTxPrevInstID             As Long
    Dim lTxIDARMAbort             As Long
    Dim lDBugCountNearTxByID      As Long     'SCR 22034 local var for Mutek spying
    Dim lDBugCountNearTxByHandle  As Long     'SCR 22034 local var for Mutek spying
    
    '---------  Initialization & Test Basic Exit Conditions  -------
    On Error GoTo ErrorHandler
    StartTx = glIGNORE_TX
    'StackTrace "StartTx", True
    
    If Not ApplicationARMd Then
      'StackTrace "StartTx", False
      Exit Function
    End If

    '--- DEFAULT - FLAG StopTx TO IGNORE
    '--- THE CASE WHEN COMMON CODE LAUNCHING DIFF. TXs IS INSTRUMENTED
    '--- e.g. Core doing gobjServers.StartCWOLEServer(oSvr.sAppNCID)
    lTxInstID = glIGNORE_TX
    '---------  End Initialization & Test Basic Exit Conditions  -------
    
    
    '--- check list of all registered Tx classes/types
    If mdictReadableTxClassName.Exists(sSomeTxHandle) Then
        '--- Developer ARMing the code may have an unreadable
        '--- identifier like an NCID (stored in sSomeTxHandle)
        sReadableTxClsName = mdictReadableTxClassName.Item(sSomeTxHandle)
    Else    '--- Tx class/type never registered - leave
        'StackTrace "StartTx", False
        Exit Function
    End If
        
    '<<< SCR 22034 MSOLANO 3/8/2002
    '--- EVEN ARM NEAR CALLS ARE SUSCEPTIBLE TO PREV. UNRESOLVED INSTANCES
    '--- E.G. ANY TIME A COMPONENT TALKS TO Core's gobjMiddleManComm (MMMHISComm.Comm)
    '--- terminate Tx class denoted by sSomeTxHandle
    Call ARMDebugAssert(mdictNearTxsByHandle.Exists(sSomeTxHandle) = False)
    If mdictNearTxsByHandle.Exists(sSomeTxHandle) Then

    '--- SCR 22033 flag that prev instance being aborted by ARM MSOLANO 3/8/2002
    'SCR 44601 RWalker 12/04/2003 Only log and generate transactions in DebugLog mode
    'SCR 48611 RWalker 1/23/2004 Transactions Removed
                
        '--- stop ONLY prev unresolved instance (enmARMStop_MeOnly)
        '--- no reason to suspect user dlg (where we stop all)
        'dbug.print "call StopTx for ONE (concurrent)"
        lTxPrevInstID = mdictNearTxsByHandle.Item(sSomeTxHandle)
        Call StopTx(lTxPrevInstID, enmARMTxStatusAborted, enmARMStop_MeOnly)
        
        '--- proceede to cache readable class name of Tx
    Else
        '--- everything OK - no unresolved instance
        '--- proceede to cache readable class name of Tx
    End If
    '>>> SCR 22034
           
    '--- FINALLY, START THE TARGET Tx INSTANCE
    lTxInstID = goARMAPI.StartTxInstance(sReadableTxClsName)
    
    'RWalker 12/19/02 Need to record by sReadableTxClsName versus sSomeTxHandle
    #If DEBUGON Then
        Call D_bugLog("cARM", "StartNearTx", sReadableTxClsName, , lTxInstID)
    #End If
    
    '<<< SCR 22034 MSOLANO 3/8/2002
    '--- ADD TO DICTIONARIES TRACKING UNRESOLVED Tx INSTANCES
    Call ARMDebugAssert(lTxInstID > glIGNORE_TX)
    If lTxInstID > glIGNORE_TX Then
        '--- Dict.Add *KEY*, item
        mdictNearTxsByHandle.Add sSomeTxHandle, lTxInstID
        mdictNearTxsByID.Add lTxInstID, sSomeTxHandle
        
        Call GUI_CurrTxNearFarAdd(sReadableTxClsName, lTxInstID)
        
        '--- Local var's defined for Mutex Appsight debugging
        lDBugCountNearTxByHandle = mdictNearTxsByHandle.Count
        lDBugCountNearTxByID = mdictNearTxsByID.Count
        Call ARMDebugAssert(lDBugCountNearTxByHandle = lDBugCountNearTxByID)
    Else
        'nothing
    End If
    '>>> SCR 22034
        
    StartTx = lTxInstID
    
    'StackTrace "StartTx", False
Exit Function
ErrorHandler:
    '--- DOESN'T MAKE SENSE TO INTERRUPT USER FOR EACH
    '--- FAILURE OF AN INSTANCE OF A TRANSACTION
    StartTx = glIGNORE_TX
    Exit Function

End Function



'------------------------------------------------------------
'--- PROCEDURE: StopTx
'---
'--- PURPOSE:   Stop a specific instance of a transaction
'---            ID'd by lTxInstanceID returned from StartTxInstance()
'---
'--- DOCUMENTATION:
'---    ARM documentation in VSS:
'---    $/CHCS2 INC2/SOURCE/ARM/ReadMe - Using the  CHCSII ARM component.doc
'---
'--- INPUTS:
'---      lTxInstID: unique Tx INSTANCE ID returned by StartTx()
'---      eTxStatus: ok aborted or failed; Optional
'---
'--- RETURN:    Boolean success / failure
'---
'--- HISTORY:
'--- * MSolano    12/17/2001 SCR-19804
'---              lTxInstanceStatus As enmARMTxStatus = enmARMTxStatusOk
'------------------------------------------------------------
Public Function StopTx(lTxInstID As Long, _
                  Optional eTxStatus As enmARMTxStatus = enmARMTxStatusOk, _
                  Optional eStopWhichTx As enmARMStopWhichTx = enmARMStop_AllCurrTx _
                  ) As Boolean
    Dim sSomeTxHandle             As String
    Dim sTxName                   As String
    Dim lDBugCountNearTxByID      As Long     'local var for Mutek spying
    Dim lDBugCountNearTxByHandle  As Long     'local var for Mutek spying
        
    '---------  Initialization & Test Basic Exit Conditions  -------
    On Error GoTo ErrorHandler
    'StackTrace "StopTx", True
    
    StopTx = True   'assume happy
    
    '--- Leave if ARM API never initialized or calling routine not ARM'd
    '--- e.g. workstation not configured for ARM (LibARM32.dll not installed)
    If lTxInstID = glIGNORE_TX Or Not ApplicationARMd() Then
      Exit Function
    End If
    
    'SCR 45176 RWalker 11/19/2003 If doesn't exist, has been stopped by another process
    If Not mdictNearTxsByID.Exists(lTxInstID) Then
        Exit Function
    End If
    
    '---------  End Initialization & Test Basic Exit Conditions  -------
    
    '--- SCR 22241 MSolano 3/16/02
    '--- STOP ALL if both:
    '---  status not good - ASSUMES USER DLG RUINED TIME
    '---  did not override default behavior to stop all
    If eTxStatus <> enmARMTxStatusOk _
        And eStopWhichTx = enmARMStop_AllCurrTx Then
            
      '*******************
      ' ABORT ALL FEATURE
      '*******************
      '--- STOP ALL CURR ONGOING Tx's: NEAR & FAR
      '--- IF ANY ONE Tx INSTANCE ABORTED OR FAILED,
      '--- PROB. INVOLVED USER DLG - SKEWS TIMES
      Call StopAllCurrTx(enmARMTxStatusAborted)
    
    Else  'STOP A SINGLE INSTANCE
    
      StopTx = goARMAPI.StopTxInstance(lTxInstID, eTxStatus)
    
      If mdictNearTxsByID.Exists(lTxInstID) Then
          sSomeTxHandle = mdictNearTxsByID.Item(lTxInstID)
      End If
      If mdictReadableTxClassName.Exists(sSomeTxHandle) Then
          sTxName = mdictReadableTxClassName.Item(sSomeTxHandle)
      Else
          sTxName = sSomeTxHandle
      End If
    
      'RWalker 1/21/03 Record sReadableTxClsName if possible
      #If DEBUGON Then
        'SCR 40974 RWalker 9/18/2003 Allow logging regardless of status of API call
          Call D_bugLog("cARM", "StopNearTx", sTxName, , lTxInstID, CLng(eTxStatus))
      #End If
      
      '<<< SCR 22034 MSOLANO 3/8/2002
      '--- DICTIONARIES TRACK UNRESOLVED NEAR CALLS
      Call ARMDebugAssert(mdictNearTxsByID.Exists(lTxInstID))
      If mdictNearTxsByID.Exists(lTxInstID) Then
          sSomeTxHandle = mdictNearTxsByID.Item(lTxInstID)
          mdictNearTxsByID.Remove (lTxInstID)
          Call ARMDebugAssert(mdictNearTxsByHandle.Exists(sSomeTxHandle))
          If mdictNearTxsByHandle.Exists(sSomeTxHandle) Then
              mdictNearTxsByHandle.Remove (sSomeTxHandle)
          End If
          
          Call GUI_CurrTxRemove(CStr(lTxInstID), sTxName)
          
          '--- Local var's defined for Mutex Appsight debugging
          lDBugCountNearTxByHandle = mdictNearTxsByHandle.Count
          lDBugCountNearTxByID = mdictNearTxsByID.Count
          Call ARMDebugAssert(lDBugCountNearTxByHandle = lDBugCountNearTxByID)
      End If
      '>>> SCR 22034
    
    End If  '(If eTxStatus <> enmARMTxStatusOk)
    
    'StackTrace "StopTx", False
    Exit Function
ErrorHandler:
    '--- DOESN'T MAKE SENSE TO INTERRUPT USER FOR EACH
    '--- FAILURE OF AN INSTANCE OF A TRANSACTION
    'StackTrace "StopTx", False
    Exit Function

End Function

'------------------------------------------------------------
'--- PROCEDURE: StartFarTx
'---
'--- PURPOSE:
'---    if transaction class is ARM'd, start a transaction instance
'---    to be resolved in a remote procedure
'--- DOCUMENTATION:
'---    ARM documentation in VSS:
'---    $/CHCS2 INC2/SOURCE/ARM/ReadMe - Using the  CHCSII ARM component.doc
'--- ASSUMPTIONS:
'---    ARMWrapper tracks if API is initialized (workstation ARM'd), otherwise exits.
'---    SOME TRANSACTIONS ARE ARM'd AT COMMON POINTS
'---    E.G. CORE gobjServers.StartCWOLEServer(oSvr.sAppNCID)
'---    ONLY SOME OF THE ACTIONS (TRANSACTIONS) LAUNCHED
'---    AT THESE COMMON CODE POINTS ARE ARM'D
'--- INPUTS:
'---    sSomeTxHandle:
'---    External name of a transaction class
'---    May be an unreadable identifier like an NCID
'---    E.G. common code like Core's
'---    gobjServers.StartCWOLEServer(oSvr.sAppNCID)
'---    This name is arbitrary - chosen by developer as
'---    s/he defines a transaction class to be ARM'd
'--- RETURN:
'---    Transaction instance ID USED BY StopTx(),
'---    or FLAG FOR StopTx TO IGNORE THIS TRANSACTION.
'--- HISTORY:
'--- * MSolano    8/7/2001  Initial Version SCR-ARM-16000
'--- * MSolano    3/8/2002  SCR 22034
'------------------------------------------------------------
Public Function StartFarTx(sSomeTxHandle As String) As Boolean
    Dim sReadableTxClsName  As String
    Dim lIDFarTxInstance    As Long
        
    '---------  Initialization & Test Basic Exit Conditions  -------
    On Error GoTo ErrorHandler
    'StackTrace "StartFarTx", True
    
    '--- DEFAULT - FLAG StopTx TO IGNORE
    '--- THE CASE WHEN COMMON CODE LAUNCHING DIFF. TXs IS INSTRUMENTED
    '--- e.g. Core doing gobjServers.StartCWOLEServer(oSvr.sAppNCID)
    lIDFarTxInstance = glIGNORE_TX
    StartFarTx = False      'Assume failure
    
    If Not CheckOk2FarStartAndGetReadableTxClsName( _
                  sSomeTxHandle, sReadableTxClsName) Then Exit Function
    '---------  End Initialization & Test Basic Exit Conditions  -------
        
    Call CheckForAndStopPrevUnresolvedInstances(sSomeTxHandle)
    
    lIDFarTxInstance = goARMAPI.StartTxInstance(sReadableTxClsName)
    
    'RWalker 12/19/02 Record sReadableTxClsName versus sSomeTxHandle
    #If DEBUGON Then
      Call D_bugLog("cARM", "StartFarTx", sReadableTxClsName, , lIDFarTxInstance)
    #End If
    
    #If mbDebugMode = True Then
      GoSub DEBUG_SECTION
    #End If

    Call Add2CacheOfCurrentTxInstances(sSomeTxHandle, lIDFarTxInstance)
    
    StartFarTx = True
    'StackTrace "StartFarTx", False
Exit Function
ErrorHandler:
    StartFarTx = False
    Call ARMDebugAssert(StartFarTx)
    'StackTrace "StartFarTx", False
    Exit Function

'--- Put conditionally compiled debug stuff in a subroutine so
'--- it doesn 't clutter the real code
DEBUG_SECTION:
    mlTimeStartMS = Timer * 1000
    'bug.Print "StartFarTx", Now
    Return

End Function
'------------------------------------------------------------
'--- PROCEDURE: StopFarTx
'---
'--- PURPOSE:   Stop a specific instance of a transaction
'---            started by StartFarTx()
'--- INPUTS:
'---  * sSomeTxHandle
'---    transaction name - may be the UNReadable version
'---    See "Specific" vs. "Generic" Procedures in associsted design doc
'---  * eTxStatus: ok aborted or failed; Optional
'---
'--- RETURN:    Boolean success / failure
'---
'--- HISTORY:
'--- * MSolano    12/17/2001 SCR-19804
'---              lTxInstanceStatus As enmARMTxStatus = enmARMTxStatusOk
'------------------------------------------------------------
Public Function StopFarTx( _
                           sSomeTxHandle As String, _
                  Optional eTxStatus As enmARMTxStatus = enmARMTxStatusOk, _
                  Optional eStopWhichTx As enmARMStopWhichTx = enmARMStop_AllCurrTx _
                  ) As Boolean
    
    '---------  Initialization & Test Basic Exit Conditions  -------
    On Error GoTo ErrorHandler
    StopFarTx = True   'assume happy
    'StackTrace "StopFarTx", True
    
    If Not ApplicationARMd Then
      'StackTrace "StopFarTx", False
      Exit Function
    End If
    '-------- End Initialization & Test Basic Exit Conditions  -------
    
    '--- SCR 22241 MSolano 3/16/02
    '--- STOP ALL if both:
    '---  status not good - ASSUMES USER DLG RUINED TIME
    '---  did not override default behavior to stop all
    If eTxStatus <> enmARMTxStatusOk _
        And eStopWhichTx = enmARMStop_AllCurrTx Then
            
      '*******************
      ' ABORT ALL FEATURE
      '*******************
      '--- STOP ALL CURR ONGOING Tx's: NEAR & FAR
      '--- IF ANY ONE Tx INSTANCE ABORTED OR FAILED,
      '--- PROB. INVOLVED USER DLG - SKEWS TIMES
      Call StopAllCurrTx(enmARMTxStatusAborted)
    
    Else
    
      Call StopOneFarTxInstance(sSomeTxHandle, eTxStatus)
    
    End If '(If eTxStatus <> enmARMTxStatusOk)
    
    
    #If mbDebugMode = True Then
      GoSub DEBUG_SECTION
    #End If
    
    'StackTrace "StopFarTx", False
    Exit Function
ErrorHandler:
    StopFarTx = False
    Call ARMDebugAssert(StopFarTx)
    'StackTrace "StopFarTx", False
    Exit Function

'--- Put conditionally compiled debug stuff in a subroutine so
'--- it doesn 't clutter the real code
DEBUG_SECTION:
      mlTimeStopMS = Timer * 1000
      mlTimeDiffMS = mlTimeStopMS - mlTimeStartMS
      'bug.Print "StopFarTx", (mlTimeDiffMS / 1000)
      'bug.Print "---"
      Return

End Function

'>>> SCR 26339 MSolano 8/8/2002
Public Function StartCmplxTx( _
                           ByVal sSomeTxHandle As String, _
                  Optional ByVal vData31Char As Variant, _
                  Optional ByVal vMetric1 As Variant, _
                  Optional ByVal vMetric2 As Variant, _
                  Optional ByVal vMetric3 As Variant, _
                  Optional ByVal vMetric4 As Variant, _
                  Optional ByVal vMetric5 As Variant, _
                  Optional ByVal vMetric6 As Variant, _
                  Optional ByRef eCorrelatedFamily As CorrelatedFamilies_Enum = eNoFamily _
                  ) As Boolean

  Dim curTime1  As Currency         'SCR 32099 Coleman 2/3/03
  Dim curTime2  As Currency         'SCR 32099 Coleman 2/3/03
  
  '---------  Initialization & Test Basic Exit Conditions  -------
  On Error GoTo ErrorHandler
    
  If Not ApplicationARMd Then
    'StackTrace "StartTx", False
    Exit Function
  End If
  
  '--- DEFAULT - FLAG StopTx TO IGNORE
  '--- THE CASE WHEN COMMON CODE LAUNCHING DIFF. TXs IS INSTRUMENTED
  '--- e.g. Core doing gobjServers.StartCWOLEServer(oSvr.sAppNCID)
  StartCmplxTx = False
  '---------  End Initialization & Test Basic Exit Conditions  -------
 
  'SCR 32099 Coleman 2/3/03
  QueryPerformanceCounter curTime1
  StartCmplxTx = moCmplxTxs.StartCmplxTx( _
                                sSomeTxHandle, _
                                vData31Char, _
                                vMetric1, _
                                vMetric2, _
                                vMetric3, _
                                vMetric4, _
                                vMetric5, _
                                vMetric6, _
                                eCorrelatedFamily)
  'SCR 45003 RWalker 11/14/2003 Modal Form without an active Family should stop all transactions
  If ((Not StartCmplxTx) And (eCorrelatedFamily <> eNoFamily) And (InStr(1, UCase(sSomeTxHandle), "CORRUT") = 1)) Then
  '12/12/2003      Call StopAllCurrTx(enmARMTxStatusAborted, enmCHCSIIDlgDetect_CustDlg)
  End If
                                
  '<<< SCR 32099 Coleman 2/3/03
  QueryPerformanceCounter curTime2
  GUI_ARM_StartTime curTime1, curTime2
  '>>> SCR 32099 Coleman 2/3/03
  Exit Function
ErrorHandler:
  StartCmplxTx = False
  Call ARMDebugAssert(StartCmplxTx)
  Exit Function

End Function
'<<< SCR 26339
'>>> SCR 26339 MSolano 8/8/2002
Public Function StopCmplxTx( _
                           ByVal sSomeTxHandle As String, _
                  Optional ByVal eTxStatus As enmARMTxStatus = enmARMTxStatusOk, _
                  Optional ByVal eStopWhichTx As enmARMStopWhichTx = enmARMStop_AllCurrTx, _
                  Optional ByVal vData31Char As Variant, _
                  Optional ByVal vMetric1 As Variant, _
                  Optional ByVal vMetric2 As Variant, _
                  Optional ByVal vMetric3 As Variant, _
                  Optional ByVal vMetric4 As Variant, _
                  Optional ByVal vMetric5 As Variant, _
                  Optional ByVal vMetric6 As Variant _
                  ) As Boolean
  
  Dim curTime1  As Currency         'SCR 32099 Coleman 2/3/03
  Dim curTime2  As Currency         'SCR 32099 Coleman 2/3/03
  
  '---------  Initialization & Test Basic Exit Conditions  -------
  On Error GoTo ErrorHandler
  'StackTrace "StopTx", True
  
  StopCmplxTx = True    'assume happy
  
  If Not ApplicationARMd Then
    'StackTrace "StopTx", False
    Exit Function
  End If
  '---------  End Initialization & Test Basic Exit Conditions  -------
  
  'SCR 32099 Coleman 2/3/03
  QueryPerformanceCounter curTime1
  StopCmplxTx = moCmplxTxs.StopCmplxTx( _
                              sSomeTxHandle, _
                              eTxStatus, _
                              eStopWhichTx, _
                              vData31Char, _
                              vMetric1, _
                              vMetric2, _
                              vMetric3, _
                              vMetric4, _
                              vMetric5, _
                              vMetric6)
  '<<< SCR 32099 Coleman 2/3/03
  QueryPerformanceCounter curTime2
  GUI_ARM_StopTime curTime1, curTime2
  '>>> SCR 32099 Coleman 2/3/03
  Exit Function
ErrorHandler:
  StopCmplxTx = False
  Call ARMDebugAssert(StopCmplxTx)    'stop in IDE
  Exit Function

End Function
'<<< SCR 26339

'----------------------------------------------------------------------
'--- PROCEDURE:
'---
'--- PURPOSE:
'---  Stop all currently ongoing Tx instances
'---  both near and far calls
'---
'--- DOCUMENTATION:
'---    ARM documentation in VSS:
'---    $/CHCS2 INC2/SOURCE/ARM/ReadMe - Using the  CHCSII ARM component.doc
'--- INPUTS:
'---  eTxStatus (optional):
'---    Ok; Aborted or Failed
'---  eDlg (optional):
'---    flags if a user CHCSII dlg was detected (ruins timings)
'---  sDetectedInModule (optional):
'---    flags what msModule the rougue dlg detected in
'---
'--- BYREFOUTPUT: None
'--- RETURN:      Boolean
'---
'--- CHANGE HISTORY:
'--- * MSolano 3/16/2002 Initial Version SCR 22241
'----------------------------------------------------------------------
Public Function StopAllCurrTx( _
                    Optional eTxStatus As enmARMTxStatus = enmARMTxStatusOk, _
                    Optional eDlg As enmCHCSIIDlg = enmCHCSIIDlgDetect_None, _
                    Optional sDetectedInModule As String = "" _
                    ) As Boolean
  Dim oCmplxTransactions          As ARM.cCmplxTransactions
  
  '---------  Initialization & Test Basic Exit Conditions  -------
  On Error GoTo ErrorHandler
     
  StopAllCurrTx = True   'assume happy
  
  If Not ApplicationARMd Then
    'StackTrace "StopAllCurrTx", False
    Exit Function
  End If
  
  If Not (NumAllCurrentInstancesAnyTxType > 0) Then Exit Function  'SCR 24288, 28539
  '-------- End Initialization & Test Basic Exit Conditions  -------
   
  
  'SCR 44601 RWalker 12/04/2003 Only log and generate transactions in DebugLog mode
  'SCR 48611 RWalker 1/23/2004 Transactions Removed

  '--- do the actual forced stops
  Call StopAllCurrFarTx(eTxStatus)
  Call StopAllCurrNearTx(eTxStatus)
  Call moCmplxTxs.StopCurrCmplxTxsNOTCrlatd(eTxStatus, eDlg, sDetectedInModule) 'SCR 28539 MSolano 10/15/02
  
  
  'SCR 44601 RWalker 12/04/2003 Only log and generate transactions in DebugLog mode
  'SCR 48611 RWalker 1/23/2004 Transactions Removed

  Exit Function
ErrorHandler:
  StopAllCurrTx = False

  'SCR 44601 RWalker 12/04/2003 Only log and generate transactions in DebugLog mode
  'SCR 48611 RWalker 1/23/2004 Transactions Removed

End Function



Private Function StopAllCurrFarTx( _
                    Optional eTxStatus As enmARMTxStatus = enmARMTxStatusOk) _
                    As Boolean
  Dim lNoTxInstances    As Long
  Dim lIDFarTxInstance  As Long
  Dim vDictKey          As Variant
  Dim sTxName           As String
      
  '---------  Initialization & Test Basic Exit Conditions  -------
  On Error GoTo ErrorHandler
  'StackTrace "StopAllCurrFarTx", True
  
  StopAllCurrFarTx = True   'assume happy
  '-------- End Initialization & Test Basic Exit Conditions  -------

  lNoTxInstances = mdictCachFarTxIDs.Count 'use local var for Mutek spying
  For Each vDictKey In mdictCachFarTxIDs
    'mdictCachFarTxIDs.Add sSomeTxHandle, lIDFarTxInstance
    lIDFarTxInstance = mdictCachFarTxIDs.Item(vDictKey)
    'RWalker 1/9/03 Log stopped transactions
    If mdictReadableTxClassName.Exists(vDictKey) Then
      sTxName = mdictReadableTxClassName.Item(vDictKey)
    Else
      sTxName = CStr(vDictKey)
    End If
    #If DEBUGON Then
      Call D_bugLog("cARM", "StopAllCurrFarTx", sTxName, , lIDFarTxInstance, CLng(eTxStatus))
    #End If
    goARMAPI.StopTxInstance lIDFarTxInstance, eTxStatus
    Call GUI_CurrTxRemove(CStr(lIDFarTxInstance), sTxName)
  Next
  
  '--- clear cach of current ongoing FAR Tx instances
  mdictCachFarTxIDs.RemoveAll

  'StackTrace "StopAllCurrFarTx", False
  Exit Function
ErrorHandler:
  'StackTrace "StopAllCurrFarTx", False
  Exit Function

End Function
                    
Private Function StopAllCurrNearTx( _
                      Optional eTxStatus As enmARMTxStatus = enmARMTxStatusOk) _
                          As Boolean
  Dim lTxInstIDToStop                 As Long
  Dim vKeyDictByHndl                  As Variant
  Dim vKeyDictByID                    As Variant
  Dim lARMTxIDsToStop()               As Long
  Dim lTxIDFlag_ByHandleDictBroken    As Long
  Dim lCntDictByID                    As Long
  Dim lCntDictByHandle                As Long
  
  '---------  Initialization & Test Basic Exit Conditions  -------
  On Error GoTo ErrorHandler
  'StackTrace "StopAllCurrNearTx", True
  
  StopAllCurrNearTx = True   'assume happy
  '-------- End Initialization & Test Basic Exit Conditions  -------
  
  
  'mdictNearTxsByHandle.Add Key=sSomeTxHandle, Item=lARMTxIDToStop
  For Each vKeyDictByHndl In mdictNearTxsByHandle
    lTxInstIDToStop = mdictNearTxsByHandle(vKeyDictByHndl)
  
    If mdictNearTxsByID.Exists(lTxInstIDToStop) Then
      '--- ALL'S WELL: parallel dictionaries are synch'd
      mdictNearTxsByID.Remove (lTxInstIDToStop)
    Else
      
      '--- UH-OH!: flag Tivoli that 2 dictionaries out
      '--- of synch for the next Tx about to be stopped
      '--- NULL TIME FLAG - force flag out now, above target Tx stop
      'SCR 48611 RWalker 1/23/2004 Transactions Removed
    End If
    
    'RWalker 1/9/03 Log stopped transactions
    #If DEBUGON Then
    If mdictReadableTxClassName.Exists(vKeyDictByHndl) Then
      Call D_bugLog("cARM", "StopAllCurrNearTx", mdictReadableTxClassName.Item(vKeyDictByHndl), , lTxInstIDToStop, CLng(eTxStatus))
    Else
      Call D_bugLog("cARM", "StopAllCurrNearTx", CStr(vKeyDictByHndl), , lTxInstIDToStop, CLng(eTxStatus))
    End If
    #End If
    
    '--- NOW - after flag, stop the target Tx instance
    goARMAPI.StopTxInstance lTxInstIDToStop, eTxStatus
  Next
  
  '--- purge first dictionary
  mdictNearTxsByHandle.RemoveAll
  
  '--- shouldn't be more left in 2nd dictionary
  lTxInstIDToStop = glIGNORE_TX
  For Each vKeyDictByID In mdictNearTxsByID
    'mdictNearTxsByID.Add Key=lARMTxID, Item=sSomeTxHandle
    lTxInstIDToStop = CLng(vKeyDictByID)
          
    '--- UH-OH!: flag Tivoli that 2 dictionaries out
    '--- of synch for the next Tx about to be stopped
    '--- NULL TIME FLAG - force flag out now, above target Tx stop
    'SCR 48611 RWalker 1/23/2004 Transactions Removed
    
    '--- NOW - after flag, stop the target instance
    goARMAPI.StopTxInstance lTxInstIDToStop, eTxStatus
    
    'RWalker 1/9/03 Log stopped transactions
    #If DEBUGON Then
    If mdictReadableTxClassName.Exists(vKeyDictByID) Then
      Call D_bugLog("cARM", "StopAllCurrNearTx", mdictReadableTxClassName.Item(vKeyDictByID), , lTxInstIDToStop, CLng(eTxStatus))
    Else
      Call D_bugLog("cARM", "StopAllCurrNearTx", CStr(vKeyDictByID), , lTxInstIDToStop, CLng(eTxStatus))
    End If
    #End If
  
  Next
  
  '--- purge second dictionary
  mdictNearTxsByID.RemoveAll
  
  'StackTrace "StopAllCurrNearTx", False
Exit Function
ErrorHandler:
  mdictNearTxsByHandle.RemoveAll
  mdictNearTxsByID.RemoveAll
  'StackTrace "StopAllCurrNearTx", False
  
End Function

Private Function CheckOk2FarStartAndGetReadableTxClsName( _
                    ByVal sSomeTxHandle As String, _
                    ByRef sInOutReadableTxClsName As String _
                    ) As Boolean
  
  On Error GoTo ErrorHandler
  CheckOk2FarStartAndGetReadableTxClsName = False  'assume pessimistic
  
  If Not ApplicationARMd Then
    'StackTrace "StartFarTx", False
    Exit Function
  End If
  
  '--- mdictReadableTxClassName REPRESENTS ALL REGISTERED Tx CLASSES
  '--- Developer ARMing the code may have an unreadable identifier
  '--- like an NCID - e.g. common code sections that launch many acts
  If mdictReadableTxClassName.Exists(sSomeTxHandle) Then
      sInOutReadableTxClsName = mdictReadableTxClassName.Item(sSomeTxHandle)
  Else  '--- Tx class/type never registered - leave
      'StackTrace "StartFarTx", False
      Exit Function
  End If
  
  CheckOk2FarStartAndGetReadableTxClsName = True
  Exit Function
ErrorHandler:
  CheckOk2FarStartAndGetReadableTxClsName = False
  Call ARMDebugAssert(CheckOk2FarStartAndGetReadableTxClsName)     'stop in IDE
  Exit Function


End Function

Private Sub CheckForAndStopPrevUnresolvedInstances( _
                ByVal sSomeTxHandle As String)
  
  On Error GoTo ErrorHandler
  
  '--- mdictCachFarTxIDs REPRESENTS ALL FAR Tx's CURRENTLY ONGOING
  '--- PRESENCE OF sSomeTxHandle IN mdictCachFarTxIDs AT THIS POINT
  '--- MEANS UNRESOLVED PAIR OF StartFarTx() / StopFarTx()
  '--- EARLIER RUN TIME INSTANCE TERMINATED W/ EITHER
  '--- FAIL (BUSINESS RULES) OR ABORTED (RUN TIME ERROR)
  
  Call ARMDebugAssert(mdictCachFarTxIDs.Exists(sSomeTxHandle) = False)
  If mdictCachFarTxIDs.Exists(sSomeTxHandle) Then
      '--- SCR 22033 flag that prev instance being aborted by ARM MSOLANO 3/8/2002
      'SCR 48611 RWalker 1/23/2004 Transactions Removed

      'SCR 44601 RWalker 12/04/2003 Only log and generate transactions in DebugLog mode
      'SCR 48611 RWalker 1/23/2004 Transactions Removed
      
      '--- stop ONLY prev unresolved instance (enmARMStop_MeOnly)
      '--- no reason to suspect user dlg (where we stop all)
      '--- MSolano 12/13/01 Added enmARMTxStatusAborted to avoid
      '--- counting this as a good Tx
      '--- SCR 22027 changed sReadableTxClsName to sSomeTxHandle MSOLANO 3/8/2002
      'dbug.print "call StopFarTx for ONE (concurrent)"
      Call StopFarTx(sSomeTxHandle, enmARMTxStatusAborted, enmARMStop_MeOnly)
      
      '--- proceede to cache readable class name of far Tx
  Else
      '--- proceede to cache readable class name of far Tx
  End If
  
  Exit Sub
ErrorHandler:
  Call ARMDebugAssert(Err.Number = 0)    'stop in IDE
  Exit Sub

End Sub
'--- cache / map readable class name of far Tx to Tx instance handle
'--- in StartFarTx() / StopFarTx() pair
'--- Tx terminator [StopFarTx()] is instrumented in code that's
'--- REMOTE to that which started the transaction; since it's not
'--- in the same routine, the instrumented code won't cache
'--- transaction instance handle (lIDFarTxInstance) - we do here:
'--- mdictCachFarTxIDs.Add *KEY*, Item
Private Sub Add2CacheOfCurrentTxInstances( _
              sSomeTxHandle As String, _
              lIDFarTxInstance As Long)
    
  On Error GoTo ErrorHandler
   
  Call ARMDebugAssert(lIDFarTxInstance > 0)
  
  If mdictCachFarTxIDs.Exists(sSomeTxHandle) Then
    '--- SHOULD NEVER HAPPEN
    Call ARMDebugAssert(mdictCachFarTxIDs.Exists(sSomeTxHandle) = False)
    mdictCachFarTxIDs.Remove (sSomeTxHandle)
  Else    'nothing
  End If
  
  If lIDFarTxInstance = glIGNORE_TX Then Exit Sub
  
  mdictCachFarTxIDs.Add sSomeTxHandle, lIDFarTxInstance
  
  Call GUI_CurrTxNearFarAdd(sSomeTxHandle, lIDFarTxInstance)
  
  Exit Sub
ErrorHandler:
  Call ARMDebugAssert(Err.Number = 0)    'stop in IDE
  Exit Sub

End Sub

Private Function StopOneFarTxInstance( _
                           ByVal sSomeTxHandle As String, _
                  Optional ByVal eTxStatus As enmARMTxStatus = enmARMTxStatusOk _
                  ) As Boolean
  Dim lIDFarTxInstance    As Long
  Dim sTxName             As String
  
  On Error GoTo ErrorHandler
  
  StopOneFarTxInstance = True     'assume optimistic
  lIDFarTxInstance = glIGNORE_TX
  
  '--- StopFarTx can be thrown w/out StartFarTx() being thrown
  '--- (e.g.when StopFarTx instrumented in
  '--- common code w/ multiple entry points)
  '--- in this case just leave
  If Not mdictCachFarTxIDs.Exists(sSomeTxHandle) Then
    'StackTrace "StopFarTx", False
    Exit Function
  End If
  
  '--- get cached TxID using Tx class handle
  lIDFarTxInstance = mdictCachFarTxIDs.Item(sSomeTxHandle)
  
  '--- now that we've found the transaction instance
  '--- we 're about to resolve w/ a stop, remove it
  '--- From cache of CURRENTLY ONGOING Far calls
  mdictCachFarTxIDs.Remove (sSomeTxHandle)
  
  '--- Leave if ARM_Start API never initialized
  If lIDFarTxInstance = glIGNORE_TX Then
    Call ARMDebugAssert(lIDFarTxInstance <> glIGNORE_TX)
    'StackTrace "StopFarTx", False
    Exit Function
  End If
  
  ''dbug.print "call to goARMAPI.StopTxInstance"
  StopOneFarTxInstance = goARMAPI.StopTxInstance(lIDFarTxInstance, eTxStatus)

  If mdictReadableTxClassName.Exists(sSomeTxHandle) Then
    sTxName = mdictReadableTxClassName.Item(sSomeTxHandle)
  Else
    sTxName = sSomeTxHandle
  End If
  'If StopOneFarTxInstance Then
    Call GUI_CurrTxRemove(CStr(lIDFarTxInstance), sTxName)
  'End If
  
  'RWalker 1/23/03 Need to record by sReadableTxClsName versus sSomeTxHandle
  #If DEBUGON Then
      'SCR 40974 RWalker 9/18/2003 Allow logging regardless of status of API call
      Call D_bugLog("cARM", "StopFarTx", sTxName, , lIDFarTxInstance, CLng(eTxStatus))
  #End If
  
  Exit Function
ErrorHandler:
  StopOneFarTxInstance = False
  Call ARMDebugAssert(StopOneFarTxInstance)     'stop in IDE
  Exit Function

End Function

'<<< SCR 36345 MSolano 5/29/2003
Public Sub SetTxNamePrefixes( _
                      lAppMode As Long, _
                      sAppMode As String)
  
  On Error GoTo ErrorHandler

  'Call ARMDebugAssert(False)
    
  goARMAPI.AppModePrefix = sAppMode

  Exit Sub
ErrorHandler:
  Call ARMDebugAssert(Err.Number = 0)     'stop in IDE
  Exit Sub

End Sub

