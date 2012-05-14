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
Attribute VB_Name = "cARMFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------
'--- CLASS:         cARMFactory
'---                Original design 8/2001
'---                by Phil Crowder & Mike Solano

'--- PURPOSE:
'---    Class factory for cARMWrapper class.
'---    Guarantees only one object instance of cARMWrapper
'---    based on singleton design pattern.
'---    The factory cARMFactory is pulic creatable
'---    (instancing prop. = multiuse).
'---    CW client creates an instance of the factory;
'---    The factory creates one instance of cARMWrapper.
'---
'--- INSTANCING:    Multiuse
'---
'--- HISTORY:
'--- * MSolano    7/30/2001  Initial Version SCR-ARM-1600
'------------------------------------------------------------
Option Explicit

'<<< SCR 22454 MSolano 3/25/02
#If DEBUGON Then
  Private Const msModule = "cARMFactory"
#End If
'>>> SCR 22454

'------------------------------------------------------------
'--- PROCEDURE:         CreateARMApp
'---                    Original design 8/2001
'---                    by Phil Crowder & Mike Solano
'--- PURPOSE:           Object factory for cARMWrapper class.
'---                    Enforces cARMWrapper as a Singleton
'--- INPUTS / OUTPUT:   None
'--- RETURN:            cARMWrapper singleton
'--- HISTORY:
'--- * MSolano    7/30/2001  Initial Version SCR-ARM-1600
'--- * MSolano    12/11/01   SCR-ARM-19645
'---                         Changed default enmArmUserLogMode from
'---                         ARMUserAsCWUserAndComputername to
'---                         ARMUserAsCWUser
'------------------------------------------------------------
Public Function CreateARMApp(sAppTitle As String, _
                            Optional sUserName As String = "", _
                            Optional enmUserLogMode As enmArmUserLogMode = ARMUserAsCWUser, _
                            Optional sINIPath As String = "") _
                    As cARMWrapper

    'SCR 48872 RWalker 2/4/2004 Added new parameter containing the INI Path.
    
    Dim bInitialized            As Boolean
    
    On Error GoTo ErrorHandler
        
    '--- ENFORCE SINGLETON HERE: create only one instance
    If goARMApp Is Nothing Then
    
        'Call ARMDebugAssert(False)
    
        '--- FIRST INSTANCE OF ARM CLASS FACTORY
        '--- FIRST TRY TO CreateARMApp SHOULD INSTANTIATE ...
        Set goARMApp = New cARMWrapper
        'SCR 48872 RWalker 2/4/2004 Added new parameter containing the INI Path.
        bInitialized = goARMApp.InitARMApp(sAppTitle, sUserName, enmUserLogMode, sINIPath)
        
        '--- FOR PURPOSE OF RETURNING REF TO SINGLETON
        '--- DO NOT TEST FOR bInitialized = TRUE
        '--- IF WORKSTATION ISN'T ARMed (MISSING LIBARM32.DLL)
        '--- ARM COMPONENT SHOULD STILL RETURN A GOOD OBJECT REFERENCE
        '--- AND JUST TRACK INTERNALLY THAT API IS NOT ACTUALLY INITIALIZED
        If Not (goARMApp Is Nothing) Then
            Set CreateARMApp = goARMApp
            
            '--- SCR 22033 Init the Tx's w/in ARM itself MSolano 3/8/2002
            If bInitialized Or gbARMOverride Then  'SCR 22541 MSolano 3/29/2002
              goARMApp.ARM_InitAllTransactionsForARM
            End If
        Else
            Set CreateARMApp = Nothing
        End If
    Else
        '--- ASSUME goARM is created / initialized
        '--- ... SUBSEQUENT INSTANCES OF CLASS FACTORY
        '--- OR SUBSEQUENT ATTEMPTS (SAME INSTANCE) TO CreateARMApp
        '--- SHOULD RETURN SINGLETON
        Set CreateARMApp = goARMApp
    End If
    
    Exit Function
    
ErrorHandler:
    '>>> SCR 22128 MSolano 3/13/2002
    '--- circular references not allowed
    '--- Report error only once
    'If Not bErrReported Then
    '    bErrReported = True
    '    goShared.ShowVBError Err.Number, Err.Description, sRoutine, App.Title, vbExclamation
    'End If
    '<<< SCR 22128
    Exit Function

End Function
'------------------------------------------------------------
'--- PROCEDURE: Class_Initialize
'--- PURPOSE:
'---    Since the factory's created first,
'---    and it sets up cARMWrapper
'---    the factory can set up infrastructure
'---    like the object instance of MMMHISShared.CWShared
'---    error handler.
'--- INPUTS / OUTPUTS / RETURNS:    None
'--- RAMIFICATIONS:                 CWShared created
'--- HISTORY:
'--- * MSolano    8/3/2001  Initial Version SCR-ARM-1600
'------------------------------------------------------------
Private Sub Class_Initialize()
    '>>> SCR 22128 MSolano 3/13/2002
    '--- circular references not allowed
    'If goShared Is Nothing Then
        'Set goShared = New MMMHISShared.CWShared
    'End If
    '<<< SCR 22128 MSolano 3/13/2002

End Sub

