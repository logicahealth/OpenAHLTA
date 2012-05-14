VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCorrelatedFamily"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event ErrorRecoveryAttempted(sMsg As String)

Public Event FamilyCreated(oFamily As ARM.cCorrelatedFamily, _
                           oParentCTInst As ARM.cCmplxTxInst)
                    
Public Event FamilyClosing(meCorrelatedFamily As CorrelatedFamilies_Enum)

Public Event FamilyTxInstOrphaned(oCTInst As ARM.cCmplxTxInst)

Public Event FamilyTxInstAdd(oCTInst As ARM.cCmplxTxInst)

Public Event FamilyTxInstRemove(oCTInst As ARM.cCmplxTxInst)


Private moCmplxTransactnsCallBk       As ARM.cCmplxTransactions

Private mdictTxInstances              As Scripting.Dictionary

Private mudtFamilyCorrelator          As ARM.Correlator_Type
Private meCorrelatorAction            As ARM_Start_CorrelatorActApp_Enum
Private bCorrelatorHasBeen            As Boolean

Private meCorrelatedFamily            As CorrelatedFamilies_Enum
Private mlFirstTxInstID               As Long
Private mlFirstTxClsID                As Long
Private mbFirstTxInstIDSet            As Boolean
Private mlLastTxInstID                As Long
Private mlFamilyStartupTime           As Long               'seconds after midnight

Private Const mlTIME_OUT_DEFAULT      As Long = 100         '100 second cap on family lifespan
Private mlTimeOutDuration             As Long

Private Const msModule                As String = "cCorrelatedFamily"


Private Sub Class_Initialize()
  On Error GoTo ErrorHandler
  
  mlFamilyStartupTime = Timer  'seconds after midnight
  mlTimeOutDuration = mlTIME_OUT_DEFAULT
  
  Set mdictTxInstances = New Scripting.Dictionary
  mlLastTxInstID = glIGNORE_TX
  meCorrelatedFamily = eNoFamily
  
  Exit Sub
ErrorHandler:
  Call ARMDebugAssert(Err.Number = 0)    'stop in IDE
  Exit Sub

End Sub

Private Sub Class_Terminate()
  On Error GoTo ErrorHandler
  
  Call moCmplxTransactnsCallBk.CallBk_CrlatdFamilyTerminated( _
                                            meCorrelatedFamily, _
                                            mlFirstTxClsID, _
                                            mlFirstTxInstID)
  
  Set moCmplxTransactnsCallBk = Nothing
  Set mdictTxInstances = Nothing
  
  Exit Sub
ErrorHandler:
  Call ARMDebugAssert(Err.Number = 0)    'stop in IDE
  Exit Sub

End Sub
Private Function GetFirstFamilyTxInstance() As ARM.cCmplxTxInst
  Dim vDictKey As Variant
  On Error GoTo ErrorHandler
  
  Set GetFirstFamilyTxInstance = Nothing
  
  '--- called when a Tx instance is removed
  '--- and it happens to be the first
  '--- assumes either that:
  '---  * the physical order of the dictionary preserves instance order
  '---  * the properties we want don't change
  '---    - correlator
  '---    - correlator action
  For Each vDictKey In mdictTxInstances
    Set GetFirstFamilyTxInstance = mdictTxInstances.Item(vDictKey)
    Exit Function
  Next
    
  Exit Function
ErrorHandler:
  Call ARMDebugAssert(Err.Number = 0)    'stop in IDE
  Set GetFirstFamilyTxInstance = Nothing

  Exit Function

End Function




Private Function Exists( _
                      lTxInstanceID As Long _
                      ) As Boolean
  On Error GoTo ErrorHandler
  
  If mdictTxInstances.Exists(lTxInstanceID) Then
    Exists = True
    'Call ARMDebugAssert(mdictTxInstances.Exists(lTxInstanceID) = False)
    Exit Function
  End If

  Exit Function
ErrorHandler:
  Exists = False
  Call ARMDebugAssert(Err.Number = 0)    'stop in IDE
  Exit Function

End Function

Private Sub ProcessFamilyParent( _
                  ByRef oCmplxTxInst As ARM.cCmplxTxInst)
  Static bDoneThis          As Boolean
  On Error GoTo ErrorHandler
   
  If bDoneThis Then
    Exit Sub                '--- EXIT
  End If
  
  '--- Family parent is the first Tx instance in the family instance
  '--- (first Tx instance among instances forming a family)
  mlFirstTxInstID = oCmplxTxInst.TxInstanceID
  mbFirstTxInstIDSet = True
  mlFirstTxClsID = oCmplxTxInst.TxClassID
  
  RaiseEvent FamilyCreated(Me, oCmplxTxInst)
  
  '--- Each family Tx instance subsequent to first
  '--- uses the correlator of the first instance
  mudtFamilyCorrelator = oCmplxTxInst.Correlator 'udtNewCorrelator
  
  bDoneThis = True
  Exit Sub
ErrorHandler:
  Call ARMDebugAssert(Err.Number = 0)    'stop in IDE
  Exit Sub

End Sub
'--- BUSINESS RULES GOVERNING WHEN A NEW TX INSTANCE CAN JOIN A FAMILY
Private Function Open_ToTHISTxInstance( _
                      lCurrTxInstanceID As Long _
                      ) As Boolean
  
  Static bClosedToALLNewInstances As Boolean
  On Error GoTo ErrorHandler
  
  Open_ToTHISTxInstance = True
  
  If bClosedToALLNewInstances Then
    Open_ToTHISTxInstance = False
  
  ElseIf Exists(lCurrTxInstanceID) Then
    '--- SHOULD NEVER HAPPEN
    Open_ToTHISTxInstance = False
    Call ARMDebugAssert(Open_ToTHISTxInstance, msModule, "Open_ToTHISTxInstance", "Duplicate entry in mdictTxInstances", 1)
  
  ElseIf lCurrTxInstanceID < 1 Then
    Open_ToTHISTxInstance = False

  Else
    Open_ToTHISTxInstance = Open_ToNewTxInstances()
    bClosedToALLNewInstances = Not (Open_ToTHISTxInstance)
  End If
  
  
  Exit Function
ErrorHandler:
  Open_ToTHISTxInstance = False
  Call ARMDebugAssert(Err.Number = 0)    'stop in IDE
  Exit Function

End Function

Private Function Closed_ParentStopped() As Boolean
  Dim oTxInst As ARM.cCmplxTxInst
  On Error GoTo ErrorHandler
  
  If mbFirstTxInstIDSet And Not Exists(mlFirstTxInstID) Then
    Closed_ParentStopped = True
    Call Closed_OrphanAllTxInstances
  End If
  
  Exit Function
ErrorHandler:
  Closed_ParentStopped = True
  Call ARMDebugAssert(Closed_ParentStopped)
  Exit Function

End Function






Public Function Open_ToNewTxInstances() As Boolean
  On Error GoTo ErrorHandler
  
  Open_ToNewTxInstances = True
  
  If Closed_TimeOut() Then
    '--- Family persisted longer than allowed
    Open_ToNewTxInstances = False
  
  ElseIf Closed_ParentStopped() Then
    '--- new Tx instance can't join family if parent Tx instance ended
    Open_ToNewTxInstances = False
  End If
  
  Exit Function
ErrorHandler:
  Open_ToNewTxInstances = False
  Call ARMDebugAssert(Err.Number = 0)
  Exit Function

End Function

Public Property Let Family(eNewVal As CorrelatedFamilies_Enum)
  On Error GoTo ErrorHandler
  
  meCorrelatedFamily = eNewVal
  TimeOutDuration = mlTIME_OUT_DEFAULT
  
  Exit Property
ErrorHandler:
  meCorrelatedFamily = eNoFamily
  Call ARMDebugAssert(Err.Number = 0)    'stop in IDE
  Exit Property

End Property

Public Property Get Family() As CorrelatedFamilies_Enum
  On Error GoTo ErrorHandler
  
  Family = meCorrelatedFamily
  
  Exit Property
ErrorHandler:
  Family = eNoFamily
  Call ARMDebugAssert(Err.Number = 0)    'stop in IDE
  Exit Property

End Property

Public Property Get LastItem() As ARM.cCmplxTxInst
  On Error GoTo ErrorHandler
  
  If mdictTxInstances.Exists(mlLastTxInstID) Then
    Set LastItem = mdictTxInstances.Item(mlLastTxInstID)
  Else
    Set LastItem = Nothing
  End If
  
  Exit Property
ErrorHandler:
  Set LastItem = Nothing
  Call ARMDebugAssert(Err.Number = 0)    'stop in IDE
  Exit Property

End Property

Public Property Get Count() As Long
  On Error GoTo ErrorHandler
  
  Count = mdictTxInstances.Count
  
  Exit Property
ErrorHandler:
  Count = glIGNORE_ME
  Call ARMDebugAssert(Count <> glIGNORE_ME)    'stop in IDE
  Exit Property
End Property

Public Sub CallBkRegistration(oCmplxTxs As ARM.cCmplxTransactions)
  On Error GoTo ErrorHandler
  
  Call ARMDebugAssert(Not (oCmplxTxs Is Nothing), msModule, "CallBkRegistration", "A1", 1)
  
  Set moCmplxTransactnsCallBk = oCmplxTxs
  
  Exit Sub
ErrorHandler:
  Call ARMDebugAssert(Err.Number = 0)
  Exit Sub
End Sub

Public Function Add( _
                  ByRef oCmplxTxInst As ARM.cCmplxTxInst _
                  ) As Boolean
  
  '---------  Initialization & Test Basic Exit Conditions  -------
  On Error GoTo ErrorHandler
  
  If Not Open_ToTHISTxInstance(oCmplxTxInst.TxInstanceID) Then
    Exit Function
  End If
  '---------  END Initialization & Test Basic Exit Conditions  -------
  
  '--- only the very first instance ever in this
  '--- family object instance should set correlator
  Call ProcessFamilyParent(oCmplxTxInst)

  mdictTxInstances.Add oCmplxTxInst.TxInstanceID, oCmplxTxInst
  mlLastTxInstID = oCmplxTxInst.TxInstanceID
  
  RaiseEvent FamilyTxInstAdd(oCmplxTxInst)
  
  Add = True
  
  Debug.Print mdictTxInstances.Count
  
  Exit Function
ErrorHandler:
  Add = False
  Call ARMDebugAssert(Add)    'stop in IDE
  Exit Function
End Function

Public Function SetCorrelationPropertiesForAPI( _
                    ByRef oCmplxTxInst As ARM.cCmplxTxInst _
                    ) As Boolean
  Dim bNewInstAlreadyExists   As Boolean
  '---------  Initialization & Test Basic Exit Conditions  -------
  On Error GoTo ErrorHandler
  
  '--- Assume calling code calls Open_ToNewTxInstances()
  '--- Do NOT call Open_ToNewTxInstances() here
  '--- That call in both places
  '--- (client code before this routine + this routine)
  '--- would create small time window during which
  '--- new Tx can 1st be accepted, then rejected
  '--- (Family open @ call #1 / times out @ call #2)
  bNewInstAlreadyExists = Exists(oCmplxTxInst.TxInstanceID)
  If bNewInstAlreadyExists Then
    '--- SHOULD NEVER HAPPEN
    Call ARMDebugAssert(Not bNewInstAlreadyExists, msModule, "SetCorrelationPropertiesForAPI", "Duplicate entry in mdictTxInstances", 1)
    Exit Function
  End If

  '---------  END Initialization & Test Basic Exit Conditions  -------

  oCmplxTxInst.CorrelatedFamily = meCorrelatedFamily

  If mdictTxInstances.Count < 1 Then
    '--- First family Tx instance uses a blank correlator
    '--- don't set Correlator prop. ON PURPOSE
    'oCmplxTxInst.Correlator =
    oCmplxTxInst.CorrelatorAction = eAppActAskAgent4Correlator
  
  Else

    oCmplxTxInst.Correlator = mudtFamilyCorrelator
    oCmplxTxInst.CorrelatorAction = eAppActPassCorrelator2SubTx
  End If
  
  SetCorrelationPropertiesForAPI = True
  Exit Function
ErrorHandler:
  SetCorrelationPropertiesForAPI = False
  Call ARMDebugAssert(SetCorrelationPropertiesForAPI)    'stop in IDE
  Exit Function

End Function

Public Function Remove( _
                  ByRef oCmplxTxInst As ARM.cCmplxTxInst _
                  ) As Boolean
  
  '---------  Initialization & Test Basic Exit Conditions  -------
  On Error GoTo ErrorHandler
  
  If Not mdictTxInstances.Exists(oCmplxTxInst.TxInstanceID) Then
    Exit Function
  End If
  '---------  END Initialization & Test Basic Exit Conditions  -------

  mdictTxInstances.Remove (oCmplxTxInst.TxInstanceID)
  RaiseEvent FamilyTxInstRemove(oCmplxTxInst)
  
  '*-- See if family timed out - if so orphan all remaining children
  Call Closed_TimeOut
  
  '--- check if parent stopped - if so close family / raise event
  Call Closed_ParentStopped
  
  Remove = True
  
  Debug.Print mdictTxInstances.Count
  Exit Function
ErrorHandler:
  Remove = False
  Call ARMDebugAssert(Remove)     'stop in IDE
  Exit Function

End Function

Public Function Closed_TimeOut() As Boolean
  Dim mlNow               As Long
  Const lSECONDS_PER_DAY  As Long = 86400
  On Error GoTo ErrorHandler
      
  '--- Wrapped around midnight?
  mlNow = Timer
  If mlNow < mlFamilyStartupTime Then
    mlNow = lSECONDS_PER_DAY + mlNow
  End If
  
  Closed_TimeOut = ((mlNow - mlFamilyStartupTime) > mlTimeOutDuration)
  
  Closed_TimeOut = GUI_CorrFamilyForceTimeOutToFalse(Closed_TimeOut)
  
  '--- CALLBACK
  If Closed_TimeOut Then
    Call Closed_OrphanAllTxInstances
  End If
  
  Exit Function
ErrorHandler:
  Call ARMDebugAssert(Err.Number = 0)    'stop in IDE
  Exit Function

End Function
Private Sub Closed_OrphanAllTxInstances()
  Dim vKey                As Variant
  Dim oCmplxTxInst        As ARM.cCmplxTxInst
  On Error GoTo ErrorHandler
  
  For Each vKey In mdictTxInstances
    Set oCmplxTxInst = mdictTxInstances.Item(vKey)
    If Not oCmplxTxInst Is Nothing Then
      If Not oCmplxTxInst.Orphaned Then
        oCmplxTxInst.Orphaned = True
        RaiseEvent FamilyTxInstOrphaned(oCmplxTxInst)
      Else
        Call ARMDebugAssert(Not oCmplxTxInst.Orphaned, msModule, "Closed_OrphanAllTxInstances", "Redundant set of Orphan status", 1)
      End If
    Else
      '--- should never happen
      Call ARMDebugAssert(Not oCmplxTxInst Is Nothing, msModule, "Closed_OrphanAllTxInstances", "Empty entry in mdictTxInstances", 1)
    End If
  Next vKey
      
  RaiseEvent FamilyClosing(meCorrelatedFamily)

  Exit Sub
ErrorHandler:
  Call ARMDebugAssert(Err.Number = 0, msModule, "Closed_OrphanAllTxInstances", "A2")
  Exit Sub
End Sub


Public Property Let TimeOutDuration(lTimeOut As Long)
  On Error GoTo ErrorHandler
    
  mlTimeOutDuration = lTimeOut
  '*-- Pass the Timeout value to the GUI and he will pass it back or change
  '*-- it depending on the GUI settings.
  mlTimeOutDuration = GUI_CorrFamilySetTimeoutDuration(lTimeOut, meCorrelatedFamily)
  Exit Property
ErrorHandler:
  Call ARMDebugAssert(Err.Number = 0)
  Exit Property
End Property

Public Property Get TimeOutDuration() As Long
  On Error GoTo ErrorHandler
    
  TimeOutDuration = mlTimeOutDuration
    
  Exit Property
ErrorHandler:
  Call ARMDebugAssert(Err.Number = 0)
  Exit Property
End Property
