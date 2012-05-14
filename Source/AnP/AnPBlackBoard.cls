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
Attribute VB_Name = "AnPBlackBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This class is used to hold references to objects used throughout A/P.

Option Explicit

Private mOLEServers As AnPOLEServers
Private mChildInterface As ChildInterface

Private mComm As Object
Private mConfig As Object
Private mLogon As Object
Private mPatient As Object
Private mUser As Object

Private mEncParent As Object

Private mModules As Collection     'Module NCID is the key, all module Initialized through Core

Private mAllOrders As AnPOrders
Private mAllDiagnoses As AnPDiagnoses

'External Facades
Private mProxyAMD As ProxyADM
Private mProxyCDRPlus As ProxyCDRPlus
'Private mProxyConsult As ProxyConsult
Private mProxyEncounter As ProxyEncounter
Private mProxyOE As ProxyOE
Private mProxyProblems As ProxyProblems

Private mNullGridRow As iGridRow    'Single instance of Null Grid Row used by AnPMainView

Private Sub Class_Initialize()
  Set mModules = New Collection
End Sub

Public Property Get ADMInterface() As ProxyADM
  If mProxyAMD Is Nothing Then
    Set mProxyAMD = New ProxyADM
  End If
  
  Set ADMInterface = mProxyAMD
End Property

Public Property Get AllDiagnoses() As AnPDiagnoses
  Set AllDiagnoses = mAllDiagnoses
End Property

Public Property Get AllOrders() As AnPOrders
  Set AllOrders = mAllOrders
End Property

Public Property Get AnPChild(ByVal NCID As String) As AnPOLEServer
  On Error Resume Next
  Set AnPChild = mOLEServers.Item(NCID)
End Property

Public Property Get AnPChildInterface() As ChildInterface
  Set AnPChildInterface = mChildInterface
End Property

Public Property Get AnPChildren() As AnPOLEServers
  Set AnPChildren = mOLEServers
End Property

Public Property Get Comm() As Object
  Set Comm = mComm
End Property

Public Property Set Comm(ByRef NewValue As Object)
  Set mComm = NewValue
End Property

Public Property Get Config() As Object
  Set Config = mConfig
End Property

Public Property Set Config(ByRef NewValue As Object)
  Set mConfig = NewValue
End Property

Public Property Get CDRPlusInterface() As ProxyCDRPlus
  If mProxyCDRPlus Is Nothing Then
    Set mProxyCDRPlus = New ProxyCDRPlus
  End If
  
  Set CDRPlusInterface = mProxyCDRPlus
End Property

Public Property Get ConsultInterface() As Object 'Delete me
'  If mProxyConsult Is Nothing Then
'    Set mProxyConsult = New ProxyConsult
'  End If
'
'  Set ConsultInterface = mProxyConsult
End Property

Public Property Get Diagnosis(ByVal Key As String) As AnPDiagnosis
  On Error Resume Next
  
  Set Diagnosis = mAllDiagnoses.Item(Key)
End Property

'12/10/04 Fix Type mismatch error on build box
Public Property Get CurrentEncounter() As CHCSIIEncounterOps.Encounter
    'SF SCR 63686
    '  Set CurrentEncounter = mobjParentEncounter.CurrentEncounter
    Set CurrentEncounter = mEncParent.CurrentEncounter
End Property

Public Property Get EncParent() As ENCOUNTER_INTERFACES.IParent  'Object
  Set EncParent = mEncParent
End Property

Public Property Set EncParent(ByRef NewValue As ENCOUNTER_INTERFACES.IParent) 'Object)

  'SF SCR 63686
  Set mEncParent = NewValue
  
    'SF commented all this out. It should use the encounter that is passed in.
    'Set mobjParentEncounter = NewValue.CurrentEncounter
    '  '12/10/04 Fix Type mismatch error on build box
    '  If mobjParentEncounter Is Nothing Then
    '    Set mobjParentEncounter = New CHCSIIEncounterCurrent.EncounterParent
    '  End If
    '  Set mobjParentEncounter = gBB.Comm.InitializeOLEServer(NCID_ENCOUNTER_APPLICATION)

End Property

Public Property Get EncounterInterface() As ProxyEncounter
  If mProxyEncounter Is Nothing Then
    Set mProxyEncounter = New ProxyEncounter
  End If
  
  Set EncounterInterface = mProxyEncounter
End Property

Public Property Get Logon() As Object
  Set Logon = mLogon
End Property

Public Property Set Logon(ByRef NewValue As Object)
  Set mLogon = NewValue
End Property

Public Property Get NullGridRow() As iGridRow
  If mNullGridRow Is Nothing Then
    Set mNullGridRow = New NullGridRow
  End If
  
  Set NullGridRow = mNullGridRow
End Property

Public Property Get OEInterface() As ProxyOE
  If mProxyOE Is Nothing Then
    Set mProxyOE = New ProxyOE
  End If
  
  Set OEInterface = mProxyOE
End Property

Public Property Get Order(ByVal OrderType As enumOrderType, ByVal Key As String) As iOrder
  On Error Resume Next
  
  Set Order = mAllOrders.Item(OrderType, Key)
End Property

Public Property Get Patient() As Object
  Set Patient = mPatient
End Property

Public Property Set Patient(ByRef NewValue As Object)
  Set mPatient = NewValue
End Property

Public Property Get ProblemInterface() As ProxyProblems
  If mProxyProblems Is Nothing Then
    Set mProxyProblems = New ProxyProblems
  End If
  
  Set ProblemInterface = mProxyProblems
End Property

Public Property Get User() As Object
  Set User = mUser
End Property

Public Property Set User(ByRef NewValue As Object)
  Set mUser = NewValue
End Property

'Clear all module initialized through Core by A/P and the Encounter Parent object.
Public Sub ClearBlackBoard()
  On Error Resume Next
  
  Set mModules = Nothing
  Set mEncParent = Nothing
  Set mOLEServers = Nothing
  Set mChildInterface = Nothing
  
  Set mProxyProblems = Nothing
'  Set mProxyConsult = Nothing
  Set mProxyEncounter = Nothing
  Set mProxyOE = Nothing
  Set mProxyCDRPlus = Nothing
  Set mProxyAMD = Nothing
  
  Set mNullGridRow = Nothing
  
  Set mAllDiagnoses = Nothing
  Set mAllOrders = Nothing
  
End Sub

Public Sub Init()
  Set mModules = New Collection
  Set mOLEServers = New AnPOLEServers
  Set mChildInterface = New ChildInterface

  If mAllDiagnoses Is Nothing Then
    Set mAllDiagnoses = New AnPDiagnoses
  End If
  
  If mAllOrders Is Nothing Then
    Set mAllOrders = New AnPOrders
  End If
End Sub

'Get a reference to any module that can be created by core.
Public Function ReadModule(ByVal NCID As String) As Object
Dim obj As Object

  On Error Resume Next
  Set obj = mModules.Item(NCID)
  
  On Error GoTo ErrHandler
  
  If obj Is Nothing Then
    If Not (mComm Is Nothing) Then
      Set obj = mComm.InitializeOLEServer(NCID)
      
      If Not (obj Is Nothing) Then mModules.Add obj, NCID
    End If
  End If
  
  Set ReadModule = obj
  
  Exit Function
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnPBlackBoard.ReadModule", "A/P", vbExclamation

End Function
