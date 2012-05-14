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
Attribute VB_Name = "cARMTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------
'--- CLASS:         cARMTransaction
'--- PURPOSE:
'---    Encapsulates Application Response Measurement (ARM) API
'--- INSTANCING: PublicNotCreatable
'---    cARM.InitTxClass is the factory
'---    trasaction class name set there
'--- HISTORY:
'--- * MSolano    8/1/2001  Initial Version
'------------------------------------------------------------
Option Explicit

Private Const msMODULE = "cARMTransaction"

'--- Application Variables
Private mlAppID             As Long     'Application ID from goARMApp

'--- TRANSACTION CLASS VARIABLES
'--- A transaction class is a chunk of code or
'--- unit of work representing a transaction,
'--- & is named by the developer ARMing the target test code.
'--- Compare w/ Transaction Instance
Private msTxClassName       As String   'Name of Tx Class vs. instance
Private mlTxClassID         As Long     'ID of Tx Class vs. Instance; Internal
Private msTxClassDetail     As String
Private mlTxClassFlags      As Long
Private msTxClassData       As String
Private mlTxClassDataSize   As Long     'Internal

'--- TRANSACTION INSTANCE VARIABLES
Private mlTxInstanceID      As Long     'Unique handle of a specific instance of a named Tx
Public Property Get TxClassName() As String
    TxClassName = msTxClassName
End Property
Public Property Let TxClassName(sData As String)
    msTxClassName = sData
End Property
Public Property Get TxDetail() As String
    TxDetail = msTxClassDetail
End Property
Public Property Let TxDetail(sData As String)
    msTxClassDetail = sData
End Property
Public Property Get TxData() As String
    msTxClassData = Trim(msTxClassData)
    mlTxClassDataSize = Len(msTxClassData)
    TxData = msTxClassData
End Property
Public Property Let TxData(sData As String)
    msTxClassData = sData
End Property
Public Property Get TxFlags() As Long
    TxFlags = mlTxClassFlags
End Property
Public Property Let TxFlags(lData As Long)
    mlTxClassFlags = lData
End Property
Private Sub Class_Initialize()
    
    If goARMApp Is Nothing Then
    '--- This code not necessary if
    '--- original design of cARM.InitTxClass
    '--- kept as the factory for cARMTransaction
        'Raise Error  '
    End If
    
    '--- INIT APPLICATION LEVEL VAR'S
    mlAppID = goARMApp.AppID
    
    '--- INIT TRANSACTION CLASS LEVEL VAR'S
    mlTxClassID = lTransactionClassID()
    
    '--- DEFAULT VALUES
    '--- ARM V2.0: THESE AREN'T USED YET
    msTxClassDetail = ""
    mlTxClassFlags = 0
    msTxClassData = ""
    mlTxClassDataSize = 0

End Sub

'------------------------------------------------------------
'--- PROCEDURE:     Prop Get TxClassID()
'--- PURPOSE:
'---    Return unique ID of what ARM API terms
'---    a "transaction class" -
'---    a named unit of code, named by the developer,
'---    representing one trasaction "type" or "class"
'---    (VS. a specific instance of that Tx)
'---
'--- INPUTS:
'---    sTxClassName: name of "class" (type) of the transaction
'---                  (vs. a specific instance of that Tx)
'---    sTxDetail: Transaction detail - ARM V2.0 API SPEC:
'---        * typical use: arguments
'---        * 128 byte free form text, incl. null terminator
'---        * if contents in subsequent calls change using same
'---          mlAppID / sTxClassName pair, but sTxDetail changes,
'---          new sTxDetail ignored
'--- RETURN:    Transaction ID - long
'--- HISTORY:
'--- * MSolano    8/1/2001  Initial Version
'------------------------------------------------------------
Private Function lTransactionClassID() As Long
    Dim lTxClsID            As Long
    Static bErrReported     As Boolean
    Static bInitialized     As Boolean
    Dim sRoutine            As String
    On Error GoTo EH
    
    '--- DO API CALL FOR TX CLASS LEVEL ID ONLY ONCE
    '--- Semaphore - stop re-entrance, re-initialization
    If bInitialized Then
        Exit Function
    Else
        bInitialized = True
    End If
    
    sRoutine = msMODULE & ".TxClassID"
    
    lTxClsID = arm_getid(mlAppID, _
                        msTxClassName, _
                        msTxClassDetail, _
                        mlTxClassFlags, _
                        msTxClassData, _
                        mlTxClassDataSize)
    
    lTransactionClassID = lTxClsID
    
    'If lTxClsID < 0 Then ERR.Raise
Exit Function
EH:
    '--- Report error only once
    If Not bErrReported Then
        bErrReported = True
        goShared.ShowVBError Err.Number, Err.Description, sRoutine, App.Title, vbExclamation
    End If
    Exit Function

End Function
