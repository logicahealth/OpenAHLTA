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
Attribute VB_Name = "cARMMsgBx"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'----------------------------------------------------------------------
'--- CLASS:   cARMMsgBx
'---
'--- PURPOSE:
'---  ARM'd wrapper for VB MsgBox
'---
'--- INSTANCING:
'---  Global Multiuse so appears as keyword
'---
'--- CHANGE HISTORY:
'--- * MSolano 3/17/2002 Initial Version
'----------------------------------------------------------------------

'<<< SCR 22454 MSolano 3/25/02
#If DEBUGON Then
  Private Const msModule = "cARMMsgBx"
#End If
'>>> SCR 22454

'<<< SCR 23633 MSolano 5/7/02
Private Type udtErrClone
  Description    As String
  '--- Documentation wrongly says HelpContext is a STRING
  '--- in IDE, VarType(Err.HelpContext) = 3, i.e. vbLONG
  HelpContext    As Long
  HelpFile       As String
  'Err.LastDllError 'read only property
  Number         As Long
  Source         As String
End Type

Private mudtErrClone As udtErrClone
Private Const NO_ERROR As Long = 0
'>>> SCR 23633


Private Const msDefaultTitle As String = "OpenAHLTA Clinical Desktop"   'SCR 28330 MSolano 10/9/2002
Private Const mi32Char As Integer = 1
Private Const miMetric6 As Integer = 7

Private mlMsgBxObjID As Long  'SCR 33025  MSolano 2/26/2003




'--- PROCEDURE:   MsgBxARMd
'---
'--- PURPOSE:
'---  Stop all CURR ongoing Tx instances due to
'---  msgbx dlg w/ user.  ARM times CHCSII, not people.
'---
'--- ASSUMPTIONS:
'---  * Function Name
'---    Named it MsgBxARMd for 2 reasons:
'---    - to closely parallel MsgBox, so it shows in Intellisense
'---    (for proj's referencing ARM)
'---    - To simplify putting it in or taking out w/ global search & replace.
'---    Neither name, taken as a whole, is a substring of the other
'---    don 't have to remember to check "Find whole word only" chkbox
'---    on search / replace dlg.
'---
'--- DOCUMENTATION:
'---  VSS/Source $/CHCS2 INC2/Source/ARM
'---
'--- INPUTS:
'---  * Title
'---    Default MsgBox title when missing is App.Tittle
'---    (This is the DLL's title as listed in the .VBP
'---    not Core's or CW's title [unless MsgBox used in CW.exe])
'---    Default title for MsgBxARMd changes this behavior.
'---
'--- BYREFOUTPUT: None
'---
'--- RETURN:      enum of VbMsgBoxResult
'---
'--- CHANGE HISTORY:
'--- * MSolano 3/17/2002  Initial Version SCR-22128
'--- * MSolano 5/7/2002   SCR 23633
'--- * MSolano 10/9/2002  SCR 28330
'----------------------------------------------------------------------
Public Function MsgBxARMd( _
                  Optional vPrompt_String As Variant = vbNullString, _
                  Optional vButtons_VbMsgBoxStyle As Variant = vbOKOnly, _
                  Optional vTitle_String As Variant = msDefaultTitle, _
                  Optional vHelp_String As Variant = vbNullString, _
                  Optional vContext_Long As Variant = 0) _
                  As VbMsgBoxResult
  
  Dim bCachedErrObj       As Boolean
  Dim bStartedCmplxTx     As Boolean  'SCR 32152 MSolano 2/14/2003
    
  '<<< SCR 23633 MSolano 5/7/02
  '--- If MsgBxARMd used in error handler, cache Err obj's state
  '--- so code after MsgBxARMd can use meaningful Err obj
  If Err.Number <> NO_ERROR Then
    bCachedErrObj = True
    Call ErrState_Save
  End If
  '>>> SCR 23633
  '--- Only AFTER caching ERR state can we set err handler - clears Err obj
  On Error GoTo ErrorHandler

  '>>> SCR 32928 RWalker MSolano 2/25/2003
  If Not (goARMApp Is Nothing) And ApplicationARMd() Then
    Call StopAllCurrentTx
    'bStartedCmplxTx = goARMApp.StartCmplxTx(gsARMTx_ARMMsgBxMeth, , , , , , , , eAnyFamily)      'SCR 33025  MSolano 3/28/03
  End If
  '<<< SCR 32928
    
  '--- sHelp & lContext are both optional,
  '--- but if one is present, other must be PRESENT to avoid run time error
  '--- But does not have to be MEANINGFUL to avoid err - i.e. "" & 0 work
  '--- Unit tests of the possible combinations
  '--- show that default values cover the possibilities

  MsgBxARMd = MsgBox( _
                  vPrompt_String, _
                  vButtons_VbMsgBoxStyle, _
                  vTitle_String, _
                  vHelp_String, _
                  vContext_Long)
  
  If bStartedCmplxTx Then
    'goARMApp.StopCmplxTx gsARMTx_ARMMsgBxMeth, enmARMTxStatusOk, enmARMStop_MeOnly  'SCR 33025  MSolano 3/28/03
  End If
  
  '<<< SCR 23633 MSolano 5/7/02
  If bCachedErrObj Then
    Call ErrState_Set   'Restore Err obj's state
  End If
  '>>> SCR 23633

  Exit Function
ErrorHandler:
  '--- this should never happen:
  Dim sErr                As String
  MsgBxARMd = vbCritical
  sErr = "The following error occurred: " & vbCrLf
  sErr = sErr & Err.Description & vbCrLf
  sErr = sErr & "When displaying a message box "
  MsgBox sErr
  'sErr = sErr & "with following message:" & vbCrLf
  'MsgBox sErr & sCoercedErrorHandlerPrompt
  
End Function
Public Sub DoEventsARMd( _
                    Optional vAppTittle As Variant = vbNullString, _
                    Optional vModuleName As Variant = vbNullString, _
                    Optional vProcName As Variant = vbNullString, _
                    Optional vLineLocatn As Variant = vbNullString _
                    )
  Dim sArr(mi32Char To miMetric6) As String
  Dim bStartedCmplxTx             As Boolean  'SCR 32152 MSolano 2/14/2003
  On Error GoTo ErrorHandler
    
  '>>> SCR 32928 MSolano 2/25/2003
  If Not (goARMApp Is Nothing) And ApplicationARMd() Then
  
    '>>> SCR 32152 MSolano 3/28/03
    'Call DoEventsARMd_ParseMetrics( _
                sArr(), vAppTittle, vModuleName, _
                vProcName, vLineLocatn)
    

    'bStartedCmplxTx = goARMApp.StartCmplxTx(gsARMTx_DoEvents, sArr(1), _
                          sArr(2), sArr(3), sArr(4), sArr(5), _
                          sArr(6), sArr(7), eAnyFamily)
    '<<< SCR 32152
  End If
  '<<< SCR 32928
  
  DoEvents
  
  If bStartedCmplxTx Then
    'goARMApp.StopCmplxTx gsARMTx_DoEvents, enmARMTxStatusOk, enmARMStop_MeOnly  'SCR 32152 MSolano 3/28/03
  End If
    
  Exit Sub
ErrorHandler:
  Call ARMDebugAssert(Err.Number = 0)
  Exit Sub
End Sub
Private Sub DoEventsARMd_ParseMetrics( _
                             ByRef sArr() As String, _
                    Optional ByVal vAppTittle As Variant = vbNullString, _
                    Optional ByVal vModuleName As Variant = vbNullString, _
                    Optional vProcName As Variant = vbNullString, _
                    Optional vLineLocatn As Variant = vbNullString _
                    )
  Dim s             As String
  Dim sModProcLine  As String
  Dim i As Integer
  Dim j As Integer
  On Error GoTo ErrorHandler
  
  
  s = Trim(CStr(vAppTittle))
  If Len(s) Then sModProcLine = s
  s = Trim(CStr(vModuleName))
  If Len(s) Then sModProcLine = sModProcLine & "." & s
  s = Trim(CStr(vProcName))
  If Len(s) Then sModProcLine = sModProcLine & "." & s
  s = Trim(CStr(vLineLocatn))
  If Len(s) Then sModProcLine = sModProcLine & "." & s
  
  i = Len(sModProcLine)
  
  '123...32 33-40 41-48 49-56
  
  If i <= 32 Then
    sArr(1) = sModProcLine
  Else
    sArr(1) = Left(sModProcLine, 32)
    s = Mid(sModProcLine, 33)
    i = Len(s)
    Debug.Print "Mod.Proc.Line: "; s
    Do While Len(s) > 0 And j < 7
      j = j + 1
      s = Left(s, 8)
      sArr(j) = s
      Debug.Print "Mod.Proc.Line: "; s
      s = Mid(s, 9)
    Loop
  End If
  
  
  Exit Sub
ErrorHandler:
  Call ARMDebugAssert(Err.Number = 0)
  Exit Sub
End Sub
'--- STOP all current Tx's w/ FAIL flag indicating user dlg encountered ...
Private Sub StopAllCurrentTx()
  On Error GoTo ErrorHandler
  
  If Not goARMApp Is Nothing Then
    goARMApp.StopAllCurrTx enmARMTxStatusFailed, enmCHCSIIDlgDetect_VBMsgBx, "ARM"
  End If
    
  'Just leave on failure
ErrorHandler:
  Exit Sub

End Sub
'----------------------------------------------------------------------
'--- PROCEDURE:  ErrState_Save
'---
'--- PURPOSE:
'--- Cache state of err object for cases where
'--- MsgBxARMd is used in error handler in client code
'---
'--- ASSUMPTIONS:
'---  NO ERROR HANDLER
'---    > On purpose - it would clear state of err object being cached
'---    > If we err here (famous last words: should never happen):
'---      - we blow out of here
'---      - we blow out of MsgBxARMd (err handler not yet set)
'---      - we blow out of last client call (already in err handler)
'---      - IF there's a prior call, with an err handler, we're OK
'---      - Otherwise, uh!
'---
'--- CHANGE HISTORY:
'--- * MSolano 5/7/2002 Initial Version SCR 23633
'----------------------------------------------------------------------
Private Sub ErrState_Save()

  mudtErrClone.Description = Err.Description
  mudtErrClone.HelpContext = Err.HelpContext
  mudtErrClone.HelpFile = Err.HelpFile
  'Err.LastDllError 'read only property
  mudtErrClone.Number = Err.Number
  mudtErrClone.Source = Err.Source
  
End Sub
'----------------------------------------------------------------------
'--- PROCEDURE:  ErrState_Set
'---
'--- PURPOSE:
'--- Restore cached state of err object if MsgBxARMd
'--- Used in an error handler in client code
'---
'--- CHANGE HISTORY:
'--- * MSolano 5/7/2002 Initial Version SCR 23633
'----------------------------------------------------------------------
Private Sub ErrState_Set()

  On Error GoTo ErrorHandler

  If mudtErrClone.Number = NO_ERROR Then
    Err.Clear
  
  Else
  
    Err.Description = mudtErrClone.Description
    Err.HelpContext = mudtErrClone.HelpContext
    Err.HelpFile = mudtErrClone.HelpFile
    'Err.LastDllError 'read only
    Err.Number = mudtErrClone.Number
    Err.Source = mudtErrClone.Source
  End If
    
  Exit Sub
ErrorHandler:
  Err.Clear
  mudtErrClone.Number = NO_ERROR
  Exit Sub

End Sub



'----------------------------------------------------------------------
'           CANDIDATE SOLUTION FOR SCR 28330 . . .
'----------------------------------------------------------------------

'Private Function CoerceString(ByVal vInString As Variant) As String
'  On Error GoTo ErrorHandler
'
'  CoerceString = CStr(vInString)
'
'  Exit Function
'ErrorHandler:
'  CoerceString = vbNullString
'  Call ARMDebugAssert(CoerceString <> vbNullString)    'stop in IDE
'  Exit Function
'
'End Function
'
'Private Function CoerceVBMsgBoxStyleEnum(ByVal vInStyle As Variant) As VbMsgBoxStyle
'  On Error GoTo ErrorHandler
'  Dim eRetVal   As VbMsgBoxStyle
'
'  eRetVal = CLng(vInStyle) 'throw an error, if you must
'
'  CoerceVBMsgBoxStyleEnum = eRetVal
'
'  Exit Function
'ErrorHandler:
'  CoerceVBMsgBoxStyleEnum = vbOKOnly
'  Call ARMDebugAssert(CoerceVBMsgBoxStyleEnum <> vbOKOnly)    'stop in IDE
'  Exit Function
'
'End Function
'
'
''--- sHelp & lContext are both optional,
''--- but if one is present, other must be PRESENT to avoid run time error
''--- But does not have to be MEANINGFUL to avoid err - i.e. "" & 0 work
''--- Unit tests of the possible combinations
''--- show that default values cover the possibilities
'Private Sub CoerceHelpFileAndContext( _
'                Optional ByRef vHelp As Variant = "", _
'                Optional ByRef vContext As Variant = 0)
'  Dim sCoercedHelp    As String
'  Dim lCoercedContext As Long
'
'  On Error GoTo ErrorHandler
'
'  If IsMissing(vHelp) Or IsMissing(vContext) Then
'    vHelp = ""
'    vContext = 0
'    Exit Sub
'
'  Else
'
'    sCoercedHelp = CStr(vHelp)
'    vHelp = sCoercedHelp
'
'    lCoercedContext = CLng(vContext)
'    vContext = lCoercedContext
'  End If
'
'  Exit Sub
'ErrorHandler:
'  vHelp = ""
'  vContext = 0
'  Call ARMDebugAssert(Err.Number = 0)    'stop in IDE
'  Exit Sub
'
'End Sub
'----------------------------------------------------------------------
'           . . . END CANDIDATE SOLUTION FOR SCR 28330
'----------------------------------------------------------------------

'>>> SCR 33025 MSolano 2/26/2003
Private Sub Class_Initialize()
  On Error GoTo ErrorHandler
  
  Exit Sub    'SCR 33025  MSolano 3/28/03
  
  goARMApp.StartCmplxTx gsARMTx_ARMMsgBxObjInit
  
  glMsgBxObjCnt = glMsgBxObjCnt + 1
  mlMsgBxObjID = glMsgBxObjCnt
  
  'MsgBox "INSTANCE ID: " & CStr(mlMsgBxObjID) & vbCrLf & _
         "INSTANCE COUNT (after increment): " & CStr(glMsgBxObjCnt), vbOKOnly, _
         "cARMMsgBx Class INITIALIZE"
  
  goARMApp.StopCmplxTx gsARMTx_ARMMsgBxObjInit, enmARMTxStatusOk, enmARMStop_MeOnly, , mlMsgBxObjID, glMsgBxObjCnt
  
ErrorHandler:
  Exit Sub
End Sub
'<<< SCR 33025

'>>> SCR 33025 MSolano 2/26/2003
Private Sub Class_Terminate()
  On Error GoTo ErrorHandler
  
  Exit Sub    'SCR 33025  MSolano 3/28/03
  
  goARMApp.StartCmplxTx gsARMTx_ARMMsgBxObjTerm
  glMsgBxObjCnt = glMsgBxObjCnt - 1
  
  'MsgBox "INSTANCE ID: " & CStr(mlMsgBxObjID) & vbCrLf & _
         "INSTANCE COUNT (after decrement): " & CStr(glMsgBxObjCnt), vbOKOnly, _
         "cARMMsgBx Class TERMINATE"

  goARMApp.StopCmplxTx gsARMTx_ARMMsgBxObjTerm, enmARMTxStatusOk, enmARMStop_MeOnly, , mlMsgBxObjID, glMsgBxObjCnt

ErrorHandler:
  Exit Sub
End Sub
'<<< SCR 33025
