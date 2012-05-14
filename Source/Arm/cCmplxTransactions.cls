VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCmplxTransactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event ErrorRecoveryAttempt(sMsg As String)

Public Event FamilyCreated( _
                  eCorrelatedFamily As CorrelatedFamilies_Enum, _
                  lTimeOut As Long)

Public Event FamilyTerminated( _
                  eCorrelatedFamily As CorrelatedFamilies_Enum, _
                  lFirstTxClsID As Long, _
                  lFirstTxInstID As Long)

Public Event FamilyTxInstAdd( _
                  sReadableName As String, _
                  lTxClsID As Long, _
                  lTxInstID As Long, _
                  eCorrelatedFamily As CorrelatedFamilies_Enum, _
                  lTimeStamp As Long)
                  
'<<< SCR 31901 Coleman 2/20/03
'*-- Changed signatures to post more info on GUI
Public Event FamilyTxInstRemove( _
                  sReadableName As String, _
                  lTxClsID As Long, _
                  lTxInstID As Long, _
                  eCorrelatedFamily As CorrelatedFamilies_Enum, _
                  lTimeStamp As Long)

Public Event FamilyTxInstOrphaned( _
                  sReadableName As String, _
                  lTxClsID As Long, _
                  lTxInstID As Long, _
                  eCorrelatedFamily As CorrelatedFamilies_Enum, _
                  lTimeStamp As Long)

Public Event CurrInstancesAdd( _
                  sReadableName As String, _
                  lTxClsID As Long, _
                  lTxInstID As Long, _
                  eCorrelatedFamily As CorrelatedFamilies_Enum, _
                  lTimeStamp As Long)

'Public Event CurrInstancesRemoveOne( _
                  sSomeTxHandle As String, _
                  oCTInst As ARM.cCmplxTxInst)
'Public Event CurrInstancesRemoveAll()

Public Event CmplxTxInstanceTerminated( _
                  sReadableName As String, _
                  lTxClsID As Long, _
                  lTxInstID As Long, _
                  eCorrelatedFamily As CorrelatedFamilies_Enum, _
                  lTimeStamp As Long)
'>>> SCR 31901 Coleman 2/20/03

Private mdictRegisteredCmplxTxClasses           As Scripting.Dictionary  'SCR 26339 MSOLANO 8/8/2002
Private mdictCmplxTx_CurrInstances              As Scripting.Dictionary  'SCR 26339 MSOLANO 8/8/2002
Private mdictCorrelatedTxFamilies               As Scripting.Dictionary  'SCR 28779 MSOLANO 11/4/02
Private mdictCloneTxClasses                     As Scripting.Dictionary
Private mbClassTerminating                      As Boolean

Private WithEvents moCorrFamilyCallBk           As ARM.cCorrelatedFamily
Attribute moCorrFamilyCallBk.VB_VarHelpID = -1
'Private WithEvents moCTInstCallBk               As ARM.cCmplxTxInst

Private Const msModule                          As String = "cCmplxTransactions"



Private Sub Class_Initialize()
    
    Set mdictRegisteredCmplxTxClasses = New Scripting.Dictionary  'SCR 26339 MSolano 8/8/2002
    Set mdictCmplxTx_CurrInstances = New Scripting.Dictionary     'SCR 26339 MSolano 8/8/2002
    Set mdictCorrelatedTxFamilies = New Scripting.Dictionary      'SCR 28779 MSolano 11/21/2002
    Set mdictCloneTxClasses = New Scripting.Dictionary
    
End Sub

Private Sub Class_Terminate()
    mbClassTerminating = True
    Set mdictRegisteredCmplxTxClasses = Nothing   'SCR 26339 MSolano 8/8/2002
    Set mdictCmplxTx_CurrInstances = Nothing      'SCR 26339 MSolano 8/8/2002
    Set mdictCorrelatedTxFamilies = Nothing       'SCR 28779 MSolano 11/21/2002
    Set mdictCloneTxClasses = Nothing
End Sub

Private Function StopOneCmplxTx( _
                         ByVal sSomeTxHandle As String, _
                Optional ByVal eTxStatus As enmARMTxStatus = enmARMTxStatusOk, _
                Optional ByVal vData31Char As Variant, _
                Optional ByVal vMetric1 As Variant, _
                Optional ByVal vMetric2 As Variant, _
                Optional ByVal vMetric3 As Variant, _
                Optional ByVal vMetric4 As Variant, _
                Optional ByVal vMetric5 As Variant, _
                Optional ByVal vMetric6 As Variant _
                        ) As ARM.cCmplxTxInst

  Dim oCmplxTxInst                      As ARM.cCmplxTxInst
  Dim bContinue                         As Boolean
  
  '---------  Initialization & Test Basic Exit Conditions  -------
  On Error GoTo ErrorHandler
  
  Set oCmplxTxInst = GetCurrInstAndDropFromCaches(sSomeTxHandle)

  If oCmplxTxInst Is Nothing Then
    Call ARMDebugAssert(Not oCmplxTxInst Is Nothing, msModule, "StopOneCmplxTx", "A1")
    Exit Function
  End If
  '-------  End Initialization & Test Basic Exit Conditions  -----
  
  Call oCmplxTxInst.StopCmplxTxInstance( _
                          eTxStatus, _
                          vData31Char, _
                          vMetric1, _
                          vMetric2, _
                          vMetric3, _
                          vMetric4, _
                          vMetric5, _
                          vMetric6)
      
  Set StopOneCmplxTx = oCmplxTxInst
  Set oCmplxTxInst = Nothing
  Exit Function
ErrorHandler:
  Call ARMDebugAssert(Err.Number = 0, msModule, "StopOneCmplxTx", "A2")
  Set oCmplxTxInst = Nothing

End Function
Private Function Add2CacheOfRegisteredClasses( _
                    ByVal sSomeTxHandle As String, _
                    lTxClsID As Long, _
                    ByRef oCmplxTxCls As ARM.cCmplxTxCls _
                    ) As Boolean
  On Error GoTo ErrorHandler
  
  If lTxClsID > 0 Then
  
    oCmplxTxCls.TxClassID = lTxClsID
    
    mdictRegisteredCmplxTxClasses.Add sSomeTxHandle, oCmplxTxCls
    Add2CacheOfRegisteredClasses = True
  
  Else
  
    Add2CacheOfRegisteredClasses = False
  End If
  
  
  Exit Function
ErrorHandler:
  Add2CacheOfRegisteredClasses = False
  Call ARMDebugAssert(Add2CacheOfRegisteredClasses, msModule, "Add2CacheOfRegisteredClasses", "A1")   'stop in IDE
  Exit Function

End Function

'--- ADD TO DICTIONARIES TRACKING CURRENTLY ONGOING Tx INSTANCES
Private Function CacheCurrInstances_Add( _
                    ByVal sSomeTxHandle As String, _
                    ByVal oCmplxTxInst As ARM.cCmplxTxInst _
                    ) As Boolean
  
  On Error GoTo ErrorHandler
  
  If mdictCmplxTx_CurrInstances.Exists(sSomeTxHandle) Then
    '--- THIS SHOULD NEVER HAPPEN
    Call ARMDebugAssert(mdictCmplxTx_CurrInstances.Exists(sSomeTxHandle) = False, msModule, "CacheCurrInstances_Add", "Attempt to add duplicate Tx Instance", 1)
    Call CacheCurrInstances_GetAndDel(sSomeTxHandle, False)
  End If
  
  If oCmplxTxInst.TxInstanceID = glIGNORE_TX Then Exit Function
  
  '--- Dict.Add *KEY*, item
  mdictCmplxTx_CurrInstances.Add sSomeTxHandle, oCmplxTxInst
  '*-- SCR 31901 Coleman 2/20/03 - Changed signature to pass more info to GUI
  RaiseEvent CurrInstancesAdd(oCmplxTxInst.TxClassName, _
                                oCmplxTxInst.TxClassID, _
                                oCmplxTxInst.TxInstanceID, _
                                oCmplxTxInst.CorrelatedFamily, _
                                oCmplxTxInst.CreationTimeStamp)
  
  CacheCurrInstances_Add = True
  Exit Function
ErrorHandler:
  CacheCurrInstances_Add = False
  Call ARMDebugAssert(CacheCurrInstances_Add, msModule, "CacheCurrInstances_Add", "A2")
End Function

Private Function CacheCurrInstances_GetAndDel( _
                      ByVal sSomeTxHandle As String, _
                      ByVal bInstanceExpected As Boolean _
                      ) As ARM.cCmplxTxInst
  Dim oCTInst           As ARM.cCmplxTxInst
  On Error GoTo ErrorHandler
  
  If mdictCmplxTx_CurrInstances.Exists(sSomeTxHandle) Then
    Set oCTInst = mdictCmplxTx_CurrInstances.Item(sSomeTxHandle)
    mdictCmplxTx_CurrInstances.Remove (sSomeTxHandle)
  End If
  
  If oCTInst Is Nothing And bInstanceExpected Then
    Call ARMDebugAssert(Not bInstanceExpected, msModule, "CacheCurrInstances_GetAndDel", "Expected entry in mdictCmplxTx_CurrInstances", 1)
  End If
  
  Set CacheCurrInstances_GetAndDel = oCTInst
  Exit Function
ErrorHandler:
  Set CacheCurrInstances_GetAndDel = Nothing
  'Set moCTInstCallBk = Nothing
  Call ARMDebugAssert(Not (CacheCurrInstances_GetAndDel Is Nothing), msModule, "CacheCurrInstances_GetAndDel", "A2", 1)
  Exit Function
End Function
'--- EVEN ARM NEAR CALLS ARE SUSCEPTIBLE TO PREV. UNRESOLVED INSTANCES
'--- E.G. ANY TIME A COMPONENT TALKS TO Core's gobjMiddleManComm (MMMHISComm.Comm)
'--- terminate prev instances of complex Tx class
'--- for class eCmplxTxCls
Private Sub StopHungTxInstances( _
                ByVal sSomeTxHandle As String)
  
  Dim lTxInstIDAbortFlg           As Long
  Dim oCmplxTxInst                As ARM.cCmplxTxInst
  
  '---------  Initialization & Test Basic Exit Conditions  -------
  On Error GoTo ErrorHandler

  If Not mdictCmplxTx_CurrInstances.Exists(sSomeTxHandle) Then
    '--- All's well - no hung instances
    Exit Sub
  End If
  '---------  END Initialization & Test Basic Exit Conditions  -------
  
  'BMConcurrency
  
  'SCR 44601 RWalker 12/04/2003 Only log and generate transactions in DebugLog mode
  If gARM_LOGDebugMode Then
    '--- flag that prev instance being aborted by ARM MSOLANO 3/8/2002
    lTxInstIDAbortFlg = goARMAPI.StartTxInstance(gsARMTx_StopUnresolvedCmplxTx)
    goARMAPI.StopTxInstance (lTxInstIDAbortFlg)
  End If
          
  '--- stop ONLY prev unresolved instance (enmARMStop_MeOnly)
  '--- no reason to suspect user dlg (where we stop all)
  Call StopCmplxTx(sSomeTxHandle, enmARMTxStatusAborted, enmARMStop_MeOnly)

  Exit Sub
ErrorHandler:
  Exit Sub

End Sub

Private Function RegisteredCmplxTxClass( _
                    ByVal sSomeTxHandle As String _
                    ) As ARM.cCmplxTxCls

  On Error GoTo ErrorHandler

  If Not mdictRegisteredCmplxTxClasses.Exists(sSomeTxHandle) Then
      Call ARMDebugAssert(mdictRegisteredCmplxTxClasses.Exists(sSomeTxHandle), msModule, "RegisteredCmplxTxClass", "mdictRegisteredCmplxTxClasses: Expected entry MISSING", 1)
      '--- Tx class/type not registered - leave
      'StackTrace "StartTx", False
      Exit Function
  Else
    '--- Get metrics profile for this Tx CLASS
    Set RegisteredCmplxTxClass = mdictRegisteredCmplxTxClasses.Item(sSomeTxHandle)
  End If

  If RegisteredCmplxTxClass Is Nothing Then
    Call ARMDebugAssert(Not (RegisteredCmplxTxClass Is Nothing), msModule, "RegisteredCmplxTxClass", "mdictRegisteredCmplxTxClasses: EMPTY entry", 1)
    Exit Function
  End If

  Exit Function
ErrorHandler:
  Set RegisteredCmplxTxClass = Nothing
  Call ARMDebugAssert(Not RegisteredCmplxTxClass Is Nothing, msModule, "RegisteredCmplxTxClass", "Error Handler", 1)    'stop in IDE
  Exit Function

End Function

'--- StopOneCmplxTx can be thrown w/out StartFarTx() being thrown
'--- e.g.when StopFarTx instrumented in common code w/ multiple entry points
Private Function GetCurrInstAndDropFromCaches( _
                      ByVal sSomeTxHandle As String _
                      ) As ARM.cCmplxTxInst
  
  Dim oCmplxTxInst                      As ARM.cCmplxTxInst
  On Error GoTo ErrorHandler
  
  '--- drop from Current Instances Cache:
  Set oCmplxTxInst = CacheCurrInstances_GetAndDel(sSomeTxHandle, True)
  
  '--- drop from Correlated Family Cache:
  Call Crlatn_RemoveInstanceFromFamily(oCmplxTxInst)
  
  Set GetCurrInstAndDropFromCaches = oCmplxTxInst
  Set oCmplxTxInst = Nothing
  Exit Function
ErrorHandler:
  Set GetCurrInstAndDropFromCaches = Nothing
  Set oCmplxTxInst = Nothing
  Call ARMDebugAssert(Not GetCurrInstAndDropFromCaches Is Nothing, msModule, "GetCurrInstAndDropFromCaches", "A3")    'stop in IDE
  Exit Function

End Function

Private Function Crlatn_Start( _
                           ByRef oCmplxTxInst As ARM.cCmplxTxInst, _
                  Optional ByVal vData31Char As Variant, _
                  Optional ByVal vMetric1 As Variant, _
                  Optional ByVal vMetric2 As Variant, _
                  Optional ByVal vMetric3 As Variant, _
                  Optional ByVal vMetric4 As Variant, _
                  Optional ByVal vMetric5 As Variant, _
                  Optional ByVal vMetric6 As Variant, _
                  Optional ByVal eCorrFamily_Input As CorrelatedFamilies_Enum = eNoFamily _
                      ) As Boolean
  Dim bJoinedExistingFamilies       As Boolean
  Dim bStartInstForONEFamily        As Boolean
  Dim eCorrFamilyVar                As CorrelatedFamilies_Enum
  
  '---------  Initialization & Test Basic Exit Conditions  -------
  On Error GoTo ErrorHandler
  
  '--- Main test for correlation - Exit if none
  If eCorrFamily_Input = eNoFamily Then
    Crlatn_Start = False
    Exit Function
  End If
  '---------  End Initialization & Test Basic Exit Conditions  -------
  
  If eCorrFamily_Input = eAnyFamily Then
    '--- Instance is a "wildcard" - gets attached to ANY & ALL FAMILIES
    '--- Clone wildcard instance into all currently existing families
    bJoinedExistingFamilies = _
        Crlatn_StartInstancesForALLFamilies( _
            oCmplxTxInst, vData31Char, vMetric1, vMetric2, _
            vMetric3, vMetric4, vMetric5, vMetric6, eCorrFamily_Input)
  End If
    
  If eCorrFamily_Input <> eAnyFamily Then
    '--- instance belongs to a specific family
    eCorrFamilyVar = eCorrFamily_Input
    bStartInstForONEFamily = True
  ElseIf eCorrFamily_Input = eAnyFamily And bJoinedExistingFamilies = False Then
    '<<< SCR 32152 MSolano 2/14/2003
    'eCorrFamilyVar = eAnyFamily
    'bStartInstForONEFamily = True
    Crlatn_Start = True
    '>>> SCR 32152
  ElseIf eCorrFamily_Input = eAnyFamily And bJoinedExistingFamilies = True Then
    Crlatn_Start = True
  End If

  If bStartInstForONEFamily Then
    '--- Configure Tx Instance's Correlation properties according
    '--- to the ONE SPECIFIC family to which it belongs
    Crlatn_Start = _
    Crlatn_StartInstForONEFamily( _
        True, oCmplxTxInst, vData31Char, vMetric1, vMetric2, _
        vMetric3, vMetric4, vMetric5, vMetric6, eCorrFamilyVar)
  End If
  
  Exit Function
ErrorHandler:
  Set moCorrFamilyCallBk = Nothing
  Call ARMDebugAssert(Err.Number = 0, msModule, "Crlatn_Start", "A1")    'stop in IDE
  Exit Function

End Function
Private Function Crlatn_StartInstancesForALLFamilies( _
                           ByRef oCTInstOriginal As ARM.cCmplxTxInst, _
                  Optional ByVal vData31Char As Variant, _
                  Optional ByVal vMetric1 As Variant, _
                  Optional ByVal vMetric2 As Variant, _
                  Optional ByVal vMetric3 As Variant, _
                  Optional ByVal vMetric4 As Variant, _
                  Optional ByVal vMetric5 As Variant, _
                  Optional ByVal vMetric6 As Variant, _
                  Optional ByVal eCorrelatedFamily As CorrelatedFamilies_Enum = eNoFamily _
                      ) As Boolean
  Dim vFamilyKey        As Variant
  Dim oCTInstVar        As ARM.cCmplxTxInst
  Dim iNoClones         As Long       'No. of CLONES does NOT include original
  Dim bTryToStart       As Boolean
  On Error GoTo ErrorHandler
  
  Set oCTInstVar = oCTInstOriginal
  
  For Each vFamilyKey In mdictCorrelatedTxFamilies
      
      '*AFTER* Original means we're REALLY CLONING
      If iNoClones > 0 Then
        Set oCTInstVar = Crlatn_CopyInstance(oCTInstOriginal)
      End If
      
      If Crlatn_StartInstForONEFamily( _
            False, oCTInstVar, vData31Char, vMetric1, _
            vMetric2, vMetric3, vMetric4, vMetric5, _
            vMetric6, vFamilyKey) _
      Then
          iNoClones = iNoClones + 1
          If iNoClones > 1 Then
            '--- *AFTER* Original means we're REALLY CLONING
            '--- ONLY cloned copies go in clone cache
            Call Crlatn_AddToCloneCaches(oCTInstVar)
          End If
      End If  'Crlatn_StartInstForONEFamily
  
  Next vFamilyKey
  
  If iNoClones = 1 Then
    Crlatn_StartInstancesForALLFamilies = True
  ElseIf iNoClones > 1 Then
    Crlatn_StartInstancesForALLFamilies = True
    oCTInstOriginal.Cloned = True
  End If
  
  Set oCTInstVar = Nothing
  Exit Function
ErrorHandler:
  Call ARMDebugAssert(Err.Number = 0, msModule, "Crlatn_StartInstancesForALLFamilies", "A2")
  Exit Function
End Function
Private Function Crlatn_AddToCloneCaches( _
                    ByRef oCmplxTxInst As ARM.cCmplxTxInst _
                    ) As Boolean
  Dim odictTxInst     As Scripting.Dictionary
  Dim bOkToAdd        As Boolean
  
  '---------  Initialization & Test Basic Exit Conditions  -------
  On Error GoTo ErrorHandler

  Set odictTxInst = Crlatn_GetCloneTxInstanceDict(oCmplxTxInst)
  If odictTxInst Is Nothing Then Exit Function
  '---------  END Initialization & Test Basic Exit Conditions  -------
  
  bOkToAdd = CBool(Not (odictTxInst.Exists(oCmplxTxInst.TxInstanceID)))
  If bOkToAdd Then
    odictTxInst.Add oCmplxTxInst.TxInstanceID, oCmplxTxInst
    '*-- SCR 31901 Coleman 2/20/03 - Changed signature to pass more info to GUI
    RaiseEvent CurrInstancesAdd(oCmplxTxInst.TxClassName, _
                                oCmplxTxInst.TxClassID, _
                                oCmplxTxInst.TxInstanceID, _
                                oCmplxTxInst.CorrelatedFamily, _
                                oCmplxTxInst.CreationTimeStamp)
    Crlatn_AddToCloneCaches = True
  Else
    '--- should never happen
    Call ARMDebugAssert(bOkToAdd, msModule, "Crlatn_AddToCloneCaches", "UNEXPECTED instance entry already in INSTANCE dictionary inside mdictCloneTxClasses", 1)
    '--- try to fix
    odictTxInst.Remove oCmplxTxInst.TxInstanceID
    odictTxInst.Add oCmplxTxInst.TxInstanceID, oCmplxTxInst
  End If
 
  Exit Function
ErrorHandler:
  Call ARMDebugAssert(Err.Number = 0, msModule, "Crlatn_AddToCloneCaches", "Error Handler")
  Exit Function
End Function

'--- Return either EXISTING or NEW dictionary
'--- of Tx Instance clones for a given Tx Class
Private Function Crlatn_GetCloneTxInstanceDict( _
                    ByRef oCmplxTxInst As ARM.cCmplxTxInst _
                    ) As Scripting.Dictionary
  Dim odictTxInst   As Scripting.Dictionary
  Dim bSuccess      As Boolean
  Dim bMakeNewDict  As Boolean

  '---------  Initialization & Test Basic Exit Conditions  -------
  On Error GoTo ErrorHandler
  
  If Not oCmplxTxInst.TxClassID > 0 Then
    Call ARMDebugAssert(bSuccess, msModule, "Crlatn_GetCloneTxInstanceDict", "ClassID <= 0", 1)
    Exit Function
  End If
  '---------  END Initialization & Test Basic Exit Conditions  -------
  
  If Not mdictCloneTxClasses.Exists(oCmplxTxInst.TxClassID) Then
    bMakeNewDict = True
  
  Else  'something's there, but is it good?
  
    Set odictTxInst = mdictCloneTxClasses.Item(oCmplxTxInst.TxClassID)
    bSuccess = CBool(Not (odictTxInst Is Nothing))
    If Not bSuccess Then
      '--- should never happen
      Call ARMDebugAssert(Not bSuccess, msModule, "Crlatn_GetCloneTxInstanceDict", "mdictCloneTxClasses: Expected entry MISSING for Tx Class " & CStr(oCmplxTxInst.TxClassID), 1)
      mdictCloneTxClasses.Remove oCmplxTxInst.TxClassID
      bMakeNewDict = True
    End If
  End If
  
  If bMakeNewDict Then
    Set odictTxInst = New Scripting.Dictionary
    mdictCloneTxClasses.Add oCmplxTxInst.TxClassID, odictTxInst
  End If
  
  Set Crlatn_GetCloneTxInstanceDict = odictTxInst
  Exit Function
ErrorHandler:
  Set Crlatn_GetCloneTxInstanceDict = Nothing
  Call ARMDebugAssert(Not (Crlatn_GetCloneTxInstanceDict Is Nothing), msModule, "Crlatn_GetCloneTxInstanceDict", "Error handler", 1)
  
End Function
Private Function Crlatn_StartInstForONEFamily( _
                           ByVal bForceCreatnOfNewFam As Boolean, _
                           ByRef oCmplxTxInst As ARM.cCmplxTxInst, _
                  Optional ByVal vData31Char As Variant, _
                  Optional ByVal vMetric1 As Variant, _
                  Optional ByVal vMetric2 As Variant, _
                  Optional ByVal vMetric3 As Variant, _
                  Optional ByVal vMetric4 As Variant, _
                  Optional ByVal vMetric5 As Variant, _
                  Optional ByVal vMetric6 As Variant, _
                  Optional ByVal eCorrelatedFamily As CorrelatedFamilies_Enum = eNoFamily _
                      ) As Boolean
  Dim bContinue         As Boolean
  '---------  Initialization & Test Basic Exit Conditions  -------
  On Error GoTo ErrorHandler
  
  Call ARMDebugAssert(Not (eCorrelatedFamily = eNoFamily), msModule, "Crlatn_StartInstForONEFamily", "A1")

  'oCmplxTxInst.CorrelatedFamily = eCorrelatedFamily

  If bForceCreatnOfNewFam Then
    Set moCorrFamilyCallBk = Crlatn_GetOrMakeOpenFamily(eCorrelatedFamily)
    bContinue = CBool(Not (moCorrFamilyCallBk Is Nothing))
    If Not bContinue Then
      '--- should never happen
      Call ARMDebugAssert(Not (bContinue), msModule, "Crlatn_StartInstForONEFamily", "Expected family not returned from Crlatn_GetOrMakeOpenFamily", 1)
      Exit Function
    End If
    
  Else
  
    Set moCorrFamilyCallBk = Crlatn_GetExistingFamily(eCorrelatedFamily)
    If Not moCorrFamilyCallBk Is Nothing Then
      If Not moCorrFamilyCallBk.Open_ToNewTxInstances Then
        Set moCorrFamilyCallBk = Nothing
        Exit Function
      End If
    Else
      Exit Function
    End If
  End If
  '---------  END Initialization & Test Basic Exit Conditions  -------


  If Not moCorrFamilyCallBk Is Nothing Then
    '--- cCorrelatedFamily knows how to prepare Correlation related
    '--- properties of cCmplxTxInst (instance) for API Start call
    bContinue = moCorrFamilyCallBk.SetCorrelationPropertiesForAPI(oCmplxTxInst)
    If bContinue And oCmplxTxInst.StartCmplxTxInstance( _
                      vData31Char, vMetric1, vMetric2, _
                      vMetric3, vMetric4, vMetric5, vMetric6) _
    Then
        If Crlatn_AddInstanceToFamily(oCmplxTxInst) Then
          Crlatn_StartInstForONEFamily = True
        End If
    End If
    
#If DEBUGON Then
    'RWalker 2/13/2003 Readables for logging
    Dim sFamilyName As String
    Select Case eCorrelatedFamily
        Case 0
            sFamilyName = "NoFamily"
        Case 1
            sFamilyName = "AnyFamily"
        Case 2
            sFamilyName = "ModalFamily"
        Case 3
            sFamilyName = "SelectPatientFamily"
        Case 4
            sFamilyName = "SignEncounterFamily"
        Case 5
            sFamilyName = "TestFamily"
        Case 6
            sFamilyName = "TestFamily2"
        Case 7
            sFamilyName = "TestFamily3"
    End Select
    Call D_bugLog("cARM", "Corr_UT For " & sFamilyName, oCmplxTxInst.TxClassName, oCmplxTxInst.TxClassID, oCmplxTxInst.TxInstanceID, CLng(Crlatn_StartInstForONEFamily))
#End If
  End If
  
  Set moCorrFamilyCallBk = Nothing
  Exit Function
ErrorHandler:
  Crlatn_StartInstForONEFamily = False
  Call ARMDebugAssert(Err.Number = 0, msModule, "Crlatn_StartInstForONEFamily", "A2")
  Exit Function
End Function
Private Function Crlatn_CopyInstance( _
                      ByRef oCTInstOriginal As ARM.cCmplxTxInst _
                      ) As ARM.cCmplxTxInst
  Dim oCTInstCpy        As ARM.cCmplxTxInst
  On Error GoTo ErrorHandler

  Set oCTInstCpy = New ARM.cCmplxTxInst
  
  Call oCTInstCpy.CallBkRegistration(Me)
  Set oCTInstCpy.CmplxTxClass = oCTInstOriginal.CmplxTxClass
  
  Set Crlatn_CopyInstance = oCTInstCpy
  Exit Function
ErrorHandler:
  Set oCTInstCpy = Nothing
  Call ARMDebugAssert(Err.Number = 0, msModule, "Crlatn_CopyInstance", "A1")
  Exit Function
End Function

'Private Function Crlatn_CopyInstance_1( _
'                      ByRef oCTInstOriginal As ARM.cCmplxTxInst, _
'                      sInstTag As String _
'                      ) As ARM.cCmplxTxInst
'  Dim byStream()        As Byte
'  Dim oCTInstCpy        As ARM.cCmplxTxInst
'  Dim oCopier           As ARM.cObjectCopier
'  On Error GoTo ErrorHandler
'
'  Set oCopier = New ARM.cObjectCopier
'  'Set oCTInstCpy = New ARM.cCmplxTxInst
'
'  Debug.Print
'  Debug.Print "Original Instance/TS: "; CStr(oCTInstOriginal.TxInstanceID) & " / " & CStr(oCTInstOriginal.CreationTimeStamp)
'  byStream = oCopier.ObjectToStream(sInstTag, oCTInstOriginal)
'
'  Set oCTInstCpy = oCopier.StreamToObject(sInstTag, byStream)
'  Debug.Print "Copy Instance/TS: "; CStr(oCTInstCpy.TxInstanceID) & " / " & CStr(oCTInstCpy.CreationTimeStamp)
'
'  Set Crlatn_CopyInstance_1 = oCTInstCpy
'  Set oCopier = Nothing
'  Exit Function
'ErrorHandler:
'  Set oCopier = Nothing
'  Set oCTInstCpy = Nothing
'  Call ARMDebugAssert(Err.Number = 0)
'  Exit Function
'
'
'End Function


'--- Get a correlated family - either
'---    - an EXISTING one that's still open for membership by new Tx instances; or
'---    - a NEW one
Private Function Crlatn_GetOrMakeOpenFamily( _
                    eCorrelatedFamily As CorrelatedFamilies_Enum _
                    ) As ARM.cCorrelatedFamily
  Dim bCreateNewFamily          As Boolean
  On Error GoTo ErrorHandler

  Set moCorrFamilyCallBk = Crlatn_GetExistingFamily(eCorrelatedFamily)
  
  If Not moCorrFamilyCallBk Is Nothing Then
    If Not moCorrFamilyCallBk.Open_ToNewTxInstances Then
      '--- family exists, but it's being closed down
      Set moCorrFamilyCallBk = Crlatn_GetNewFamily(eCorrelatedFamily)
    End If
  Else
    Set moCorrFamilyCallBk = Crlatn_GetNewFamily(eCorrelatedFamily)
  End If
  
  Set Crlatn_GetOrMakeOpenFamily = moCorrFamilyCallBk
  Exit Function
ErrorHandler:
  Call ARMDebugAssert(Err.Number = 0, msModule, "Crlatn_GetOrMakeOpenFamily", "A1")
  Exit Function
End Function
Private Function Crlatn_GetNewFamily( _
                    eCorrelatedFamily As CorrelatedFamilies_Enum _
                    ) As ARM.cCorrelatedFamily

  On Error GoTo ErrorHandler

  Set moCorrFamilyCallBk = New ARM.cCorrelatedFamily
  Call moCorrFamilyCallBk.CallBkRegistration(Me)
  moCorrFamilyCallBk.Family = eCorrelatedFamily
  
  mdictCorrelatedTxFamilies.Add eCorrelatedFamily, moCorrFamilyCallBk
  '*-- Note: this event is ignored by GUI since no valid TimeStamp is passed.
  RaiseEvent FamilyCreated(eCorrelatedFamily, -1)
  
  Set Crlatn_GetNewFamily = moCorrFamilyCallBk
  Exit Function
ErrorHandler:
  Set Crlatn_GetNewFamily = Nothing
  Call ARMDebugAssert(Err.Number = 0, msModule, "Crlatn_GetNewFamily", "Error Handler", 1)
  Exit Function

  
  
End Function

Private Function Crlatn_AddInstanceToFamily( _
                        oCmplxTxInst As ARM.cCmplxTxInst _
                        ) As Boolean
  
  '---------  Initialization & Test Basic Exit Conditions  -------
  On Error GoTo ErrorHandler
  
  Set moCorrFamilyCallBk = Crlatn_GetExistingFamily(oCmplxTxInst.CorrelatedFamily)
  
  If moCorrFamilyCallBk Is Nothing Then
    '--- SHOULD NEVER HAPPEN
    Crlatn_AddInstanceToFamily = False
    Call ARMDebugAssert(Crlatn_AddInstanceToFamily, msModule, "Crlatn_AddInstanceToFamily", "Expected family missing from mdictCorrelatedTxFamilies", 1)
    Exit Function
  End If
  '---------  END Initialization & Test Basic Exit Conditions  -------

  Crlatn_AddInstanceToFamily = moCorrFamilyCallBk.Add(oCmplxTxInst)
  
  Set moCorrFamilyCallBk = Nothing
  Exit Function
ErrorHandler:
  Set moCorrFamilyCallBk = Nothing
  Crlatn_AddInstanceToFamily = False
  Call ARMDebugAssert(Crlatn_AddInstanceToFamily, msModule, "Crlatn_AddInstanceToFamily", "A2")
  Exit Function

End Function

Private Sub Crlatn_RemoveInstanceFromFamily( _
                oCmplxTxInst As ARM.cCmplxTxInst)
  
  '---------  Initialization & Test Basic Exit Conditions  -------
  On Error GoTo ErrorHandler
  
  If oCmplxTxInst Is Nothing Then
    Exit Sub
  ElseIf oCmplxTxInst.CorrelatedFamily = eNoFamily Then
    Exit Sub
  ElseIf oCmplxTxInst.Orphaned Then   'Flags that Family died
    '*-- SCR 31901 Coleman 2/20/03 - Changed signature to post more info on GUI
    With oCmplxTxInst
        RaiseEvent FamilyTxInstRemove(.TxClassName, .TxClassID, .TxInstanceID, _
                                        .CorrelatedFamily, .CreationTimeStamp)
    End With
    Exit Sub
  Else
    Set moCorrFamilyCallBk = Crlatn_GetExistingFamily(oCmplxTxInst.CorrelatedFamily)
    If moCorrFamilyCallBk Is Nothing Then
      '--- SHOULD NEVER HAPPEN
      Call ARMDebugAssert(Not moCorrFamilyCallBk Is Nothing, msModule, "Crlatn_RemoveInstanceFromFamily", "A1")
      Exit Sub
    End If
  End If
  '---------  END Initialization & Test Basic Exit Conditions  -------
  
  '--- family will raise an event if TimedOut or lost parent
  moCorrFamilyCallBk.Remove oCmplxTxInst
  
  Set moCorrFamilyCallBk = Nothing
  Exit Sub
ErrorHandler:
  Set moCorrFamilyCallBk = Nothing
  Call ARMDebugAssert(Err.Number = 0, msModule, "Crlatn_RemoveInstanceFromFamily", "A2")
  Exit Sub

End Sub
'--- CLONE is #2, after original which ...
'---    - is stored in mdictCmplxTx_CurrInstances
'---    ... has .Cloned prop set
Private Function Crlatn_StopClones( _
                               oCTInstOriginal As ARM.cCmplxTxInst, _
                Optional ByVal eTxStatus As enmARMTxStatus = enmARMTxStatusOk, _
                Optional ByVal vData31Char As Variant, _
                Optional ByVal vMetric1 As Variant, _
                Optional ByVal vMetric2 As Variant, _
                Optional ByVal vMetric3 As Variant, _
                Optional ByVal vMetric4 As Variant, _
                Optional ByVal vMetric5 As Variant, _
                Optional ByVal vMetric6 As Variant _
                    ) As Boolean
  Dim odictInst         As Scripting.Dictionary
  Dim vTxInstanceKey    As Variant
  Dim oCTInstClone      As ARM.cCmplxTxInst
  Dim bGoodInst         As Boolean
  On Error GoTo ErrorHandler
    
  Set odictInst = Crlatn_GetCloneTxInstanceDict(oCTInstOriginal)
  
  For Each vTxInstanceKey In odictInst
    Set oCTInstClone = odictInst.Item(vTxInstanceKey)
    bGoodInst = Not (oCTInstClone Is Nothing)
    If Not bGoodInst Then
      Call ARMDebugAssert(bGoodInst, msModule, "Crlatn_StopClones", "Bad instance from dictionary in mdictCloneTxClasses", 1)
    Else
      Crlatn_StopClones = _
      oCTInstClone.StopCmplxTxInstance(eTxStatus, vData31Char, vMetric1, vMetric2, _
                                       vMetric3, vMetric4, vMetric5, vMetric6)
      Call Crlatn_RemoveInstanceFromFamily(oCTInstClone)
    End If
  Next vTxInstanceKey
  
  mdictCloneTxClasses.Remove oCTInstOriginal.TxClassID
  
  Exit Function
ErrorHandler:
  Call ARMDebugAssert(Err.Number = 0)
  Exit Function
End Function
Private Function Crlatn_GetExistingFamily( _
                    eCorrelatedFamily As CorrelatedFamilies_Enum _
                    ) As ARM.cCorrelatedFamily
  On Error GoTo ErrorHandler
  
  If mdictCorrelatedTxFamilies.Exists(eCorrelatedFamily) Then
    Set Crlatn_GetExistingFamily = mdictCorrelatedTxFamilies.Item(eCorrelatedFamily)
    
    If Crlatn_GetExistingFamily Is Nothing Then
      '--- SHOULD NEVER HAPPEN ... FIX BAD DICT ENTRY
      Call ARMDebugAssert(Not Crlatn_GetExistingFamily Is Nothing, msModule, "Crlatn_GetExistingFamily", "mdictCorrelatedTxFamilies: Empty entry", 1)
      mdictCorrelatedTxFamilies.Remove (eCorrelatedFamily)
      RaiseEvent FamilyTerminated(eCorrelatedFamily, -1, -1)
    End If
  
  Else
    Set Crlatn_GetExistingFamily = Nothing
  End If
  
  Exit Function
ErrorHandler:
  Set Crlatn_GetExistingFamily = Nothing
  Call ARMDebugAssert(Not Crlatn_GetExistingFamily Is Nothing, msModule, "Crlatn_GetExistingFamily", "Error Handler")
  Exit Function

End Function







'--- called from cCorrelatedFamily Class_Terminate
Public Sub CallBk_CrlatdFamilyTerminated( _
                  eCorrelatedFamily As CorrelatedFamilies_Enum, _
                  lFirstTxClsID As Long, _
                  lFirstTxInstID As Long)
  Dim bUNEXPECTEDRef2Family As Boolean
  '---------  Initialization & Test Basic Exit Conditions  -------
  On Error GoTo ErrorHandler
  
  If mbClassTerminating Then Exit Sub
  '---------  END Initialization & Test Basic Exit Conditions  -------
  
  bUNEXPECTEDRef2Family = mdictCorrelatedTxFamilies.Exists(eCorrelatedFamily)
  If bUNEXPECTEDRef2Family Then
    Call ARMDebugAssert(bUNEXPECTEDRef2Family, msModule, "CallBk_CrlatdFamilyTerminated", "UNEXPECTED family reference persists in mdictCorrelatedTxFamilies", 1)
    '--- Delete the correlated Tx family
    mdictCorrelatedTxFamilies.Remove (eCorrelatedFamily)
  Else
    'nothing
  End If
  
  RaiseEvent FamilyTerminated(eCorrelatedFamily, lFirstTxClsID, lFirstTxInstID)
  
  Exit Sub
ErrorHandler:
  Call ARMDebugAssert(Err.Number = 0, msModule, "CallBk_CrlatdFamilyTerminated", "A1")
  Exit Sub
  Resume
End Sub

'*-- SCR 31901 Coleman 2/20/03 - Changed signature to post more info on GUI
Public Sub CallBk_CTInstTerminated( _
                  sReadableName As String, _
                  lTxClsID As Long, _
                  lTxInstID As Long, _
                  eCorrelatedFamily As CorrelatedFamilies_Enum, _
                  lTimeOut As Long)
  '---------  Initialization & Test Basic Exit Conditions  -------
  On Error GoTo ErrorHandler
  
  If mbClassTerminating Then Exit Sub
  '---------  END Initialization & Test Basic Exit Conditions  -------
  
  '*-- SCR 31901 Coleman 2/20/03 - Changed signature to post more info on GUI
  RaiseEvent CmplxTxInstanceTerminated(sReadableName, _
                                        lTxClsID, _
                                        lTxInstID, _
                                        eCorrelatedFamily, _
                                        lTimeOut)
  
  Exit Sub
ErrorHandler:
  Call ARMDebugAssert(Err.Number = 0)
  Exit Sub
End Sub
Public Property Get CountCurrentInstances() As Long
  On Error GoTo ErrorHandler
  CountCurrentInstances = mdictCmplxTx_CurrInstances.Count
  Exit Property
ErrorHandler:
  CountCurrentInstances = -1
End Property


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
  
  Dim oCmplxTxCls             As ARM.cCmplxTxCls
  Dim udtTxClsBuf             As ARM_GetID_DataBuffer_Type
  Dim lTxClassDataSize        As Long
  Dim lTxClsID                As Long
  
  '---------  Initialization & Test Basic Exit Conditions  -------
  On Error GoTo ErrorHandler
  InitCmplxTxCls = True
  If Not ApplicationARMd Then Exit Function

  If sSomeTxHandle = vbNullString Then
    '--- There's no second unreadable handle, there's just the readable one
    sSomeTxHandle = sReadableTxClsName
  Else
    '--- There are two different handles:
    '--- a readable one sReadableTxClsName & an unreadable one sSomeTxHandle
    '--- Do Nothing
  End If

  '--- Init each class only once
  If mdictRegisteredCmplxTxClasses.Exists(sSomeTxHandle) Then
      Exit Function  'assume earlier init worked
  End If
  '---------  END Initialization & Test Basic Exit Conditions  -------
  
  '--- Enum eCmplxTxCls determines properties of Complex Tx Class
  '---    Tx Class Data Buffer
  '---    Lenght of Tx Class Data Buffer
  '---    Tx Class Name
  Set oCmplxTxCls = New ARM.cCmplxTxCls
      
  '--- DataBuffer Method initializes the object
 udtTxClsBuf = oCmplxTxCls.DataBuffer( _
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
  
  lTxClassDataSize = oCmplxTxCls.DataBufferLength
  'sReadableTxClsName = oCmplxTxCls.TxClassName
  
  lTxClsID = goARMAPI.InitTxClass( _
                      sReadableTxClsName, _
                      udtTxClsBuf, _
                      lTxClassDataSize)

  InitCmplxTxCls = Add2CacheOfRegisteredClasses( _
                      sSomeTxHandle, _
                      lTxClsID, _
                      oCmplxTxCls)
  
  #If DEBUGON Then
  'RWalker 12/30/2002 - ARM - Log metrics D_bugLogMetrics
  Call D_bugLogMetrics("cARM", "Tx Metrics", sReadableTxClsName, _
                       sString31FieldName, _
                       sMetric1Of6Name, _
                       D_bugMetricTypes(eMetric1Of6Type), _
                       sMetric2Of6Name, _
                       D_bugMetricTypes(eMetric2Of6Type), _
                       sMetric3Of6Name, _
                       D_bugMetricTypes(eMetric3Of6Type), _
                       sMetric4Of6Name, _
                       D_bugMetricTypes(eMetric4Of6Type), _
                       sMetric5Of6Name, _
                       D_bugMetricTypes(eMetric5Of6Type), _
                       sMetric6Of6Name, _
                       D_bugMetricTypes(eMetric6Of6Type))
  #End If
  
  Set oCmplxTxCls = Nothing
  Exit Function

ErrorHandler:
  Set oCmplxTxCls = Nothing
  InitCmplxTxCls = False
  Call ARMDebugAssert(InitCmplxTxCls, msModule, "InitCmplxTxCls", "A1")
  Exit Function

End Function
'<<< SCR 26339

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
                  Optional ByVal eCorrelatedFamily As CorrelatedFamilies_Enum = eNoFamily _
                  ) As Boolean

  Dim oCmplxTxCls               As ARM.cCmplxTxCls
  Dim oCmplxTxInst              As ARM.cCmplxTxInst
  Dim bCorrelated               As Boolean
  
  On Error GoTo ErrorHandler
  StartCmplxTx = False
  
  Call StopHungTxInstances(sSomeTxHandle)
  
  Set oCmplxTxCls = RegisteredCmplxTxClass(sSomeTxHandle)
  If oCmplxTxCls Is Nothing Then Exit Function
    
  Set oCmplxTxInst = New ARM.cCmplxTxInst
  Call oCmplxTxInst.CallBkRegistration(Me)
  Set oCmplxTxInst.CmplxTxClass = oCmplxTxCls

  '--- Correlation Logic entry point
  If Not eCorrelatedFamily = eNoFamily Then
      Call Crlatn_Start( _
                    oCmplxTxInst, _
                    vData31Char, _
                    vMetric1, _
                    vMetric2, _
                    vMetric3, _
                    vMetric4, _
                    vMetric5, _
                    vMetric6, _
                    eCorrelatedFamily)
  
  Else
      Call oCmplxTxInst.StartCmplxTxInstance( _
                                vData31Char, _
                                vMetric1, _
                                vMetric2, _
                                vMetric3, _
                                vMetric4, _
                                vMetric5, _
                                vMetric6)
  End If
  
  If oCmplxTxInst.TxInstanceID > glIGNORE_TX Then
    StartCmplxTx = CacheCurrInstances_Add(sSomeTxHandle, oCmplxTxInst)
  End If
  
  Exit Function
ErrorHandler:
    '--- DOESN'T MAKE SENSE TO INTERRUPT USER FOR EACH
    '--- FAILURE OF AN INSTANCE OF A TRANSACTION
    StartCmplxTx = False
    Call ARMDebugAssert(StartCmplxTx, msModule, "StartCmplxTx", "A1")
    Exit Function

End Function
'<<< SCR 26339
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
  Dim oCTInst     As ARM.cCmplxTxInst
  Dim bContinue   As Boolean

  On Error GoTo ErrorHandler
  StopCmplxTx = False
  
  '--- attempt to store in buffer the metrics (if) passed in
  '--- for one specific Tx instance
  '--- regardless of conditions being met for ABORT ALL FEATURE
  Set oCTInst = StopOneCmplxTx( _
                  sSomeTxHandle, eTxStatus, vData31Char, vMetric1, _
                  vMetric2, vMetric3, vMetric4, vMetric5, vMetric6)
    
  '12/1/2003 RWalker
  If Not oCTInst Is Nothing Then
  If oCTInst.Cloned Then
      Call Crlatn_StopClones( _
            oCTInst, eTxStatus, vData31Char, vMetric1, _
            vMetric2, vMetric3, vMetric4, vMetric5, vMetric6)
  End If
  End If

  '---  ABORT ALL FEATURE
  '---  Assumes user dlg ruined Tx time ... stop ALL currently
  '---  ongoing complex Tx instances (if any), if both:
  '---  * status not good
  '---  * Client code NOT specifically instrumented to override default behavior to stop all
  '---    (no override of eStopWhichTx parameter w/ enmARMStop_MeOnly)
  If eTxStatus <> enmARMTxStatusOk _
    And eStopWhichTx = enmARMStop_AllCurrTx Then
    
    bContinue = StopCurrCmplxTxsNOTCrlatd(enmARMTxStatusAborted)

  'Else
    '--- Nothing
  End If
  
  StopCmplxTx = True
  Exit Function
ErrorHandler:
  StopCmplxTx = False
  Call ARMDebugAssert(StopCmplxTx, msModule, "StopCmplxTx", "A1")
  Exit Function

End Function

Public Function StopCurrCmplxTxsNOTCrlatd( _
                    Optional lTxStatus As enmARMTxStatus = enmARMTxStatusOk, _
                    Optional enmDlg As enmCHCSIIDlg = enmCHCSIIDlgDetect_None, _
                    Optional sDetectedInModule As String = "" _
                    ) As Boolean

  Dim bContinue                   As Boolean
  Dim vDictKey                    As Variant
  Dim oCmplxTxInst                As ARM.cCmplxTxInst

  On Error GoTo ErrorHandler
  StopCurrCmplxTxsNOTCrlatd = False
  
  If Not (mdictCmplxTx_CurrInstances.Count > 0) Then Exit Function
  
  '--- do the actual forced stops
  For Each vDictKey In mdictCmplxTx_CurrInstances
    
    'mdictCmplxTx_CurrInstances.Add KEY = eCmplxTxCls, ITEM = lTxInstID
    Set oCmplxTxInst = mdictCmplxTx_CurrInstances.Item(vDictKey)
    
    If Not oCmplxTxInst Is Nothing Then
      If oCmplxTxInst.CorrelatedFamily = eNoFamily Then
          goARMAPI.StopTxInstance oCmplxTxInst.TxInstanceID, lTxStatus
      
          'RWalker 1/14/03 Log stopped transactions
          #If DEBUGON Then
            Call D_bugLog("cARM", "StopCurrCmplxTxsNOTCrlatd", oCmplxTxInst.TxClassName, oCmplxTxInst.TxClassID, oCmplxTxInst.TxInstanceID, CLng(lTxStatus))
          #End If
          
          '--- release obj ref to evoke call back from Tx instance class_terminate
          Set oCmplxTxInst = Nothing
          mdictCmplxTx_CurrInstances.Remove (vDictKey)
      End If  'oCmplxTxInst.CorrelatedFamily = eNoFamily
    End If    'oCmplxTxInst Is Nothing
  
  Next vDictKey
  
  Set oCmplxTxInst = Nothing
  StopCurrCmplxTxsNOTCrlatd = True
  Exit Function
ErrorHandler:
  Set oCmplxTxInst = Nothing
  StopCurrCmplxTxsNOTCrlatd = False
  Call ARMDebugAssert(StopCurrCmplxTxsNOTCrlatd, msModule, "StopCurrCmplxTxsNOTCrlatd", "A1")    'stop in IDE
  Exit Function

End Function


Private Sub moCorrFamilyCallBk_ErrorRecoveryAttempted(sMsg As String)
    RaiseEvent ErrorRecoveryAttempt(sMsg)
End Sub

Private Sub moCorrFamilyCallBk_FamilyCreated( _
                  oFamily As cCorrelatedFamily, _
                  oParentCTInst As cCmplxTxInst)
  On Error GoTo ErrorHandler
  
  Debug.Print "RaiseEvent cCmplxTransactions:moCorrFamilyCallBk_FamilyCreated"
  RaiseEvent FamilyCreated(oFamily.Family, oParentCTInst.CreationTimeStamp)
    
  Exit Sub
ErrorHandler:
  Call ARMDebugAssert(Err.Number = 0)
  Exit Sub
End Sub

Private Sub moCorrFamilyCallBk_FamilyClosing( _
                meCorrelatedFamily As CorrelatedFamilies_Enum)

  On Error GoTo ErrorHandler
  
  '--------------------------------------------------------------
  '---    V1.0 CORRELATION CANDIDATE BUSINESS RULE RETIRED ---
  '--- Kill the instance of the correlated family object
  '--- NOT Tx Instances
  '--------------------------------------------------------------
  'Call Crlatn_StopAllTxInstancesInFamily
  '--------------------------------------------------------------
  
  If mdictCorrelatedTxFamilies.Exists(meCorrelatedFamily) Then
    '--- Delete the correlated Tx family
    mdictCorrelatedTxFamilies.Remove (meCorrelatedFamily)
    
    ' GUI updated in cCorrelatedFamily terminate event
    'RaiseEvent FamilyTerminated(meCorrelatedFamily)
  End If
  
  '--- should go into cCorrelatedFamily Class_Terminate
  Set moCorrFamilyCallBk = Nothing
  Exit Sub
ErrorHandler:
  Set moCorrFamilyCallBk = Nothing
  Call ARMDebugAssert(Err.Number = 0, msModule, "moCorrFamilyCallBk_FamilyClosing", "A1")
  Exit Sub
End Sub
Private Sub moCorrFamilyCallBk_FamilyTxInstAdd(oCTInst As cCmplxTxInst)
  On Error GoTo ErrorHandler

  Debug.Print "cCmplxTransactions:FamilyTxInstAdd "
  RaiseEvent FamilyTxInstAdd( _
                    oCTInst.TxClassName, _
                    oCTInst.TxClassID, _
                    oCTInst.TxInstanceID, _
                    oCTInst.CorrelatedFamily, _
                    oCTInst.CreationTimeStamp)
    
  Exit Sub
ErrorHandler:
  Call ARMDebugAssert(Err.Number = 0)
  Exit Sub
End Sub

Private Sub moCorrFamilyCallBk_FamilyTxInstRemove(oCTInst As cCmplxTxInst)
  On Error GoTo ErrorHandler
  
  '*-- SCR 31901 Coleman 2/20/03 - Changed signature to post more info on GUI
  RaiseEvent FamilyTxInstRemove( _
                    oCTInst.TxClassName, _
                    oCTInst.TxClassID, _
                    oCTInst.TxInstanceID, _
                    oCTInst.CorrelatedFamily, _
                    oCTInst.CreationTimeStamp)
  
  Exit Sub
ErrorHandler:
  Call ARMDebugAssert(Err.Number = 0)
  Exit Sub
End Sub

Private Sub moCorrFamilyCallBk_FamilyTxInstOrphaned(oCTInst As cCmplxTxInst)
  On Error GoTo ErrorHandler
  
  '*-- SCR 31901 Coleman 2/20/03 - Changed signature to post more info on GUI
  RaiseEvent FamilyTxInstOrphaned( _
                    oCTInst.TxClassName, _
                    oCTInst.TxClassID, _
                    oCTInst.TxInstanceID, _
                    oCTInst.CorrelatedFamily, _
                    oCTInst.CreationTimeStamp)

  Exit Sub
ErrorHandler:
  Call ARMDebugAssert(Err.Number = 0)
  Exit Sub
End Sub


