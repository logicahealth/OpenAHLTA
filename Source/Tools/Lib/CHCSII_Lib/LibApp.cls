VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "gclsLibApp"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''************************************************************************
'''Module:        gclsLibApp/LibApp.cls
'''
'''Description:   General Application utilities: Error handling,
'''               logs
'''
'''Modules Required:
''' Global
''' LibMsg
'''
'''Author         : Alexander Alberstein
'''Modified       :
'''               :
'''************************************************************************
Option Explicit
''''''''''''''' Constanst Section '''''''''''''''''''''''''''''''''''''''''

Const msMODULE_NAME     As String = "gclsLibApp."

Public Enum gnumDODataOperations
   lDO_Select
   lDO_Insert
   lDO_Update
   lDO_Delete
   lDO_Validate
End Enum

Public Enum gnumTRN_Transaction
   lTRN_Begin
   lTRN_Commit
   lTRN_RollBack
End Enum

Public Enum pnumYNA_YesNoAny
   lYNA_Any
   lYNA_Yes
   lYNA_No
End Enum

Private Const mlHWND_TOP         As Long = -1
Private Const mlHWND_NO_TOP      As Long = -2
Private Const mlSWP_NO_ACTIVATE  As Long = &H10

''''''''''''''' Variables Section '''''''''''''''''''''''''''''''''''''''''

Private msClientExeName             As String
Public pbDebugPrint                 As Boolean
Public pbDebugLog                   As Boolean
Private msClientVersionMajor        As String
Private msClientVersionMinor        As String
Private msClientVersionRevision     As String

Private mbSystemVersion             As Boolean  '''Used to evaluate whether to add system version to error message
Private msSystemVersionAppPathName  As String
Private msSystemVersionMajor        As String
Private msSystemVersionMinor        As String
Private msSystemVersionRevision     As String

Private msNCID_AppClient            As String '<DSN
Private msClientTitle               As String

''''''''''''''' API Section '''''''''''''''''''''''''''''''''''''''''''''''

Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function GetShortPathName Lib "Kernel32" _
       Alias "GetShortPathNameA" _
            (ByVal lpszLongPath As String, ByVal lpszShortPath As String, _
             ByVal cchBuffer As Long _
            ) As Long
Private Declare Function API_lGetWindowLong Lib "user32" _
       Alias "GetWindowLongA" (ByVal lHWnd As Long, ByVal lIndex As Long) As Long

Private Declare Function API_SetWindowPos Lib "user32" _
      Alias "SetWindowPos" _
         (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
          ByVal x As Long, ByVal y As Long, _
          ByVal cx As Long, ByVal cy As Long, _
          ByVal wFlags As Long _
         ) As Long
         
Private Declare Function API_SetParent Lib "user32" _
      Alias "SetParent" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
      
''''''''''''''' Objects Section '''''''''''''''''''''''''''''''''''''''''''''''

Private mobjShared         As CWShared
Private mobjConn           As CHCSII_CONN.Conn
Private mobjLibLogErr      As New mclsLibLog
Public pobjCWComm          As Object
Private mobjLibMsg         As gclsLibMsg
Public Property Get AppMode(ByVal lMode As enumCSCSIIAppMode) As Boolean
   AppMode = mobjShared.IsAppMode(lMode)
End Property


Public Sub LibApp_AlwaysOnTop _
(ByVal lHWnd As Long, ByVal lTop As Long, ByVal lHeight As Long, ByVal lWidth As Long, _
 Optional bOnTop = True _
)
'''************************************************************************
'''Called By:     Any
'''
'''Arguments
''' lHWnd         Window handle of the form requested to be on top
''' lTop          of the form requested to be on top
''' lHeight       of the form requested to be on top
''' lWidth        of the form requested to be on top
''' bOnTop        Optional indicates whether set form on top or cancel it
'''
'''Description:   Set specified window to always be top most window and optionaly allows
'''               to move the form
'''
'''Note:          see also LibApp_bFormAlwaysOnTop
'''Warning:
'''Updates Globals:
'''
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME           As String = "LibApp_AlwaysOnTop"
   
   Dim dXCompns               As Double
   Dim dYCompns               As Double
   Dim lOnTop                 As Long
   
On Error GoTo ErrHnd

   dXCompns = Screen.TwipsPerPixelX
   dYCompns = Screen.TwipsPerPixelY

   If bOnTop Then
      lOnTop = mlHWND_TOP
   Else
      lOnTop = mlHWND_NO_TOP
   End If
   
   Call API_SetWindowPos _
      (lHWnd, lOnTop, _
      ((Screen.Width - lWidth) / dXCompns) / 2, lTop / dYCompns, _
      lWidth / dXCompns, lHeight / dXCompns, _
      mlSWP_NO_ACTIVATE _
      )
   
Exit Sub
ErrHnd:
   Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")
   
End Sub



Public Function LibApp_bFormAlwaysOnTop(objForm As Object, Optional bOnTop = True) As Boolean
'''************************************************************************
'''Called By:     Any
'''
'''Arguments
''' objForm       Reference to requested form
''' bOnTop        Optional indicates whether set form on top or cancel it
'''
'''Description:   Set specified window to always be top most window
'''
'''Note:          See also LibApp_AlwaysOnTop
'''Warning:
'''Updates Globals:
'''
'''Returns:       True  - if successful
'''               False - else
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME           As String = "LibApp_AlwaysOnTop"
   
   Dim dXCompns               As Double
   Dim dYCompns               As Double
   Dim lOnTop                 As Long
   
On Error GoTo ErrHnd

   dXCompns = Screen.TwipsPerPixelX
   dYCompns = Screen.TwipsPerPixelY

   If bOnTop Then
      lOnTop = mlHWND_TOP
   Else
      lOnTop = mlHWND_NO_TOP
   End If
   
   Call API_SetWindowPos _
   (objForm.hwnd, lOnTop, _
    ((Screen.Width - objForm.Width) / dXCompns) / 2, _
    objForm.Top / dYCompns, _
    objForm.Width / dXCompns, objForm.Height / dXCompns, _
    mlSWP_NO_ACTIVATE _
   )
   
   LibApp_bFormAlwaysOnTop = True
   
Exit Function
ErrHnd:
   LibApp_bFormAlwaysOnTop = False
   Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")

End Function

Public Function LibApp_bPositionOnForm _
      (objControl As Object, ByRef lLeft As Long, ByRef lTop As Long) As Boolean
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
''' objControl
'''   Reference of to control position of which is requested to be evaluated
''' lLeft
'''   Distance between the internal left edge of an object and parent form
'''   requested to be calculated. Returned byref
''' lTop
'''   Distance between the internal top edge of an object and parent form
'''   requested to be calculated. Returned byref
'''
'''Description:   Determine left and right position of the control in reference
'''               to parent form
'''
'''
'''Updates Globals:
'''
'''Returns:       If Successful
'''                  True
'''               Else
'''                  False
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "LibApp_bPositionOnForm"
   Const lMAX_NEST      As Long = 20
   
   Dim bResult          As Boolean
   Dim ctlControl       As Control
   Dim objContainer     As Object
   Dim lNest            As Long
   
On Error GoTo ErrHnd
   
20    Set ctlControl = objControl
40    Set objContainer = ctlControl.Container
60    lLeft = ctlControl.Left
80    lTop = ctlControl.Top

On Error GoTo 0
On Error Resume Next
   
   Do While (Not objContainer Is Nothing) And (Not TypeOf objContainer Is Form) And (Err.Number = 0) And (lNest < lMAX_NEST)
      lNest = lNest + 1
      lLeft = lLeft + objContainer.Left
      lTop = lTop + objContainer.Top
      Set objContainer = objContainer.Container
   Loop
  
   LibApp_bPositionOnForm = bResult
   
Exit Function
ErrHnd:
   LibApp_bPositionOnForm = False
   Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl)

End Function

Public Function LibApp_bSetParent _
      (ByVal lHWndChild As Long, ByVal lHWndNewParent As Long, _
       Optional ByRef lHWndOldParent As Long, Optional ByRef sErrMsg As String _
      ) As Boolean
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
''' lHWndChild       Handle of the child window
''' lHWndNewParent   Handle of the Parent window
''' lHWndOldParent   Handle of the previous parent returned byref. 0 if error
''' sErrMsg          Error info if any returned byref
'''
'''Description:      Assign new parent window to specified child window
'''
'''Note:
''' sErrMsg may return message even if return value indicates success
'''
'''Updates Globals:
'''
'''Returns:       If Successful
'''                  True
'''               Else
'''                  False
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "LibApp_bSetParent"
   
   Dim bResult          As Boolean
   
On Error GoTo ErrHnd
   
20    sErrMsg = ""
40    lHWndOldParent = 0
100   Err.Clear
200   If Not (Err.LastDllError = 0) Then
'''<DEV make it log only for release
220      Call LibApp_Error(Err.LastDllError, "Persistent DLL error", msMODULE_NAME, sPROC_NAME, "")
      End If
      
300   If lHWndChild > 0 Then
320      lHWndOldParent = API_SetParent(lHWndChild, lHWndNewParent)
400      If Not (lHWndOldParent = 0) Then
420         bResult = True
500         If Not Err.LastDllError = 0 Then
               '''Residual error (conflicting with return value)
520            sErrMsg = "SetParent Error(rsd) " & Err.LastDllError
            End If
         Else
700         sErrMsg = "SetParent Error " & Err.LastDllError
         End If
      Else
800      sErrMsg = "Invalid child window handle(0)"
      End If
   
999   LibApp_bSetParent = bResult
  
Exit Function
ErrHnd:
   sErrMsg = Err.Description
   LibApp_bSetParent = False
   Call LibApp_Error(Err.Number, sErrMsg, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl)

End Function

Public Sub LibApp_DebugLog(ByVal sModuleName As String, sProcName As String, ByVal sMsg As String)
'''************************************************************************
'''Called By:     Any
'''
'''Arguments
''' sModuleName   Name of the caller module
''' sProcName     Name of the caller procedure
''' sMsg          Specified message
'''
'''Description:   If debug mode is ON - log specified debug message
'''
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
On Error Resume Next
   If pbDebugLog Then
      Static bStarted_static   As Boolean
      
   
   
      Static objDebugLog_static  As New DLog.cDebugLog
      Static uDR_static          As DebugRec
      
      If Not bStarted_static Then
         bStarted_static = True
         objDebugLog_static.debugStart
      End If
      
      If Not objDebugLog_static Is Nothing Then
         uDR_static.TimerTime = Timer
         Call objDebugLog_static.DebugLog(App.EXEName, sModuleName, sProcName, sMsg, uDR_static)
      End If
   End If

End Sub


Public Sub LibApp_Error _
      (ByVal lErrNo As Long, ByVal sErrMsg As String, ByVal sModuleName As String, _
       ByVal sProcName As String, ByVal sInfo As String, _
       Optional ByVal bRaise As Boolean = False, _
       Optional ByVal bDisplayErr As Boolean = True, _
       Optional ByVal bLogErr As Boolean = True _
      )
'''************************************************************************
'''Called By:     Any
'''
'''Arguments
''' lErrNo        Error Number
''' sErrMsg       Error Description
''' sModuleName   Name of the colling module (bas, class, form name...)
''' sProcName     Name of the caller procedure
''' sInfo         Additional Information
''' bRaise        Indicates whether to Raise an error after displaying message.
'''               Note: used for subs and properties to inform user about error and its location and
'''                     in the same time interrupt execution of the calling procedure when dependent
'''                     procedure fails
''' bDisplayErr   Optional request to Display error message. Defaults to true
''' bLogErr       Optional request to Log error message. Defaults to true
'''
'''Description:   Display and optionally log specified error information
'''
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME                 As String = msMODULE_NAME & "LibApp_Error"
   Const dMIN_INTERVAL              As Double = 0.0006
   Const sMIN_INTERVAL_ERR          As String = "Duplicate errors within short(40 sec) time interval." _
                                              & vbCrLf
   Const lMAX_DUPLICATES            As Long = 2

   Dim sErrReport                   As String
   Dim sSourceProc                  As String
   Static sSource_static            As String '''error source
   Static lLastErr_static           As Long
   Static sLastErrMsg_static        As String
   Static dLastTime_static          As Double
   Static lNoOfDuplicates_static    As Long
   Static bPreventDuplicate_static  As Boolean

   '''Save source before Error object is cleared
   sSource_static = Trim$(Err.Source)

On Error GoTo ErrHnd
   
   '''If same error and less then approximately 40-50 seconds since last one was recorded
20 If ((CDbl(Now) - dLastTime_static) < dMIN_INTERVAL) And _
      (lLastErr_static = lErrNo) And (sLastErrMsg_static = sErrMsg) Then
      '''If not waiting already
      If Not bPreventDuplicate_static Then
         '''If repeated more often then specified
         If (lNoOfDuplicates_static >= lMAX_DUPLICATES) Then
            bPreventDuplicate_static = True
         Else
            '''Allow a few repetitions to show user how frequently an error occurs
            '''Note then by the time this counter is incremented for first time one error
            '''already will be recorded so actual number of errors recorded wil be one more
            '''then it specified in the lMAX_DUPLICATES
            lNoOfDuplicates_static = lNoOfDuplicates_static + 1
         End If
      End If
   Else
      '''No duplicates or interval is greater then minimum
      bPreventDuplicate_static = False
   End If
   If Not bPreventDuplicate_static Then
100   dLastTime_static = CDbl(Now)
      lLastErr_static = lErrNo
      sLastErrMsg_static = sErrMsg
      '''Accumulate Procedure names to trace the stack of calls during error raise sequence.
      ''' Store most recent one on the left in the case if its trancated
200   sSourceProc = sModuleName & Trim$(sProcName) & IIf(sSource_static = "", "", "><" & sSource_static)
      
      If lErrNo = 0 Then
         lErrNo = 1
      End If
      If pbDebugPrint Or bLogErr Then
300      sErrReport = _
         gsNUMBER & vbTab & vbTab & ": " & CStr(lErrNo) & vbCrLf _
         & gsDESCRIPTION & vbTab & ": " & Trim$(sErrMsg) & vbCrLf _
         & gsSOURCE & vbTab & vbTab & ": " & sSourceProc & vbCrLf _
         & gobjLibStr.LibStr_sBuildTitle(gsPROGRAM, gsNAME) & vbTab & ": " & msClientExeName & vbCrLf _
         & gobjLibStr.LibStr_sBuildTitle(gsPROGRAM, gsVERSION) & vbTab & ": " & LibApp_sVersion(False, False) & vbCrLf _
         & IIf(mbSystemVersion, gobjLibStr.LibStr_sBuildTitle(gsSYSTEM, gsVERSION) & vbTab & ": " & LibApp_sVersion(False, True) & vbCrLf, "") _
         & gsDETAILS & vbTab & vbTab & ": " & Trim$(sInfo)
      End If
      If bLogErr Then
400      Call LibApp_DebugLog _
            (sModuleName, sProcName, "<<>> " & gsERROR & ": " & sErrReport & " - " & vbCrLf & sInfo)
420      Call mobjLibLogErr.Mtd_Message(sSourceProc, sErrReport)
440      Call gobjLibLogErr.Mtd_Message(sSourceProc, sErrReport)
         If pbDebugPrint Then
460         Debug.Print sErrReport
         End If
      End If
      If bDisplayErr Then
500      Call M_DisplayErr(lErrNo, sErrReport, sSourceProc)
         '''If requested and not already printed
         If pbDebugPrint And (Not bLogErr) Then
580         Debug.Print sErrReport
         End If
      End If
      #If LIB_LOG Then
         Call gobjLog.Mtd_Message(sSourceProc, sErrReport)
      #End If
      If bRaise Then
On Error GoTo 0
600      Call M_RaiseErr(lErrNo, sErrMsg, sModuleName, sProcName, sInfo)
      End If
   Else
      '''Log this error only ones
      If (lNoOfDuplicates_static >= lMAX_DUPLICATES) Then
700      Call mobjLibLogErr.Mtd_Message(sPROC_NAME, sMIN_INTERVAL_ERR)
720      Call gobjLibLogErr.Mtd_Message(sPROC_NAME, sMIN_INTERVAL_ERR)
         lNoOfDuplicates_static = 0
      End If
   End If
   
Exit Sub
ErrHnd:
   sErrReport = "ERROR: " & Err.Number & ", " & Err.Description & vbCrLf _
        & "LINE: " & Erl & vbCrLf _
        & "MODULE: " & sModuleName & vbCrLf _
        & "PROCEDURE: " & sProcName & vbCrLf _
        & "ERROR: " & lErrNo & ", " & sErrMsg _
        & vbCrLf & sInfo
On Error Resume Next
   Call MsgBox(sErrReport, vbOKOnly, sPROC_NAME)
   
   
End Sub




Public Function LibApp_sDataOperationName _
   (lOperation As gnumDODataOperations, Optional ByVal bIncludeLabel As Boolean = False) As String
'''************************************************************************
'''Called By:     LibMsg_lOKCancel or any
'''
'''Arguments:
''' lOperation - Data Operation identifier
''' bIncludeLabel - Indicates whether to preface name with "Data Operation: " label
'''
'''Description:   Return name of the specified data operation
'''
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If successful
'''                  name of the specified data operation
'''               else
'''                  Error message
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "LibApp_sDataOperationName"

   Dim sResult          As String
   
On Error GoTo ErrHnd

   Select Case lOperation
      Case lDO_Select
         sResult = gsSELECT
      Case lDO_Insert
         sResult = gsINSERT
      Case lDO_Update
         sResult = gsUPDATE
      Case lDO_Delete
         sResult = gsDELETE
      Case lDO_Validate
         sResult = gsVALIDATE
      Case Else
         sResult = gsUNDEFINED & CStr(lOperation)
   End Select
   
   If bIncludeLabel Then
      sResult = gobjLibStr.LibStr_sBuildSentence(gsDATA, gsOPERATION, ": ", sResult)
   Else
      sResult = gobjLibStr.LibStr_sBuildSentence(sResult)
   End If
   
   '''Return Result
   LibApp_sDataOperationName = gobjLibStr.LibStr_sBuildSentence(sResult)
   
Exit Function
ErrHnd:
   sResult = Err.Description
   LibApp_sDataOperationName = sResult
   Call LibApp_Error(Err.Number, sResult, msMODULE_NAME, sPROC_NAME, gsOPERATION & ": " & CStr(lOperation))
   
End Function


Public Function LibApp_sInitErr _
      (ByVal bInitRequired As Boolean, ByVal bCurrentlyInit As Boolean, _
       Optional bRaise As Boolean = True, Optional sInfo As String _
      ) As String
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
''' bInitRequired If true  - Not initialized is an error
'''               Else     - Already Initialized is an error
''' bCurrentlyInit - Indicates whether object is currently initialized
''' bRaise - optional request to raise an error
''' sInfo - Optional additional information
'''
'''Description:   Evaluate whether object was initialized and depending on bInitRequired
'''               determine whether current state is an error. Optionaly raise error if any
'''
'''Note:
'''
'''Warning:       If pbAcceptInvalid is true Value will be accepted without validation
'''
'''Updates Globals:
'''
'''Returns:       Empty          - if object was initialized
'''               Error Message  - Else
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "LibApp_sInitErr"

   Dim sResult          As String
   
On Error GoTo ErrHnd
   
   '''If must be but not initialized
   If bInitRequired And Not bCurrentlyInit Then
      sResult = gobjLibStr.LibStr_sBuildSentence(gsNOT, gsINITIALIZED)
   '''If must not be initialized and it is
   ElseIf Not bInitRequired And bCurrentlyInit Then
      sResult = gobjLibStr.LibStr_sBuildSentence(gsALREADY, gsINITIALIZED)
   End If
   '''If error
   If Not sResult = "" Then
      If bRaise Then
         On Error GoTo 0
         Call M_RaiseErr(1, sResult, msMODULE_NAME, sPROC_NAME, sInfo)
      End If
   End If
   
   LibApp_sInitErr = sResult
   
Exit Function
ErrHnd:
   sResult = Err.Description
   LibApp_sInitErr = sResult
   Call LibApp_Error(Err.Number, sResult, msMODULE_NAME, sPROC_NAME, "")

End Function

Public Function LibApp_sNextKey() As String
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
'''
'''Description:   Generate and return component level unique key
'''
'''
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If successful  -  Key value unique within component application session
'''               Else           - empty string
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "LibApp_sNextKey"

   Static dKey_static   As Double
   
On Error GoTo ErrHnd
   
   dKey_static = dKey_static + 1
   
   LibApp_sNextKey = CStr(dKey_static)
   
Exit Function
ErrHnd:
   LibApp_sNextKey = ""
   Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")

End Function

Sub LibApp_Wait(ByVal dSecondsToWait As Double, ByVal bDoEvents As Boolean, _
                Optional ByVal bNoMax As Boolean = False _
               )
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
''' dSecondsToWait        Number for seconds to wait
''' bDoEvents     Allow to do other events while waiting.
''' bNoMax        Do not limit wait time
'''
'''Description:   Wait specified number of seconds but no more then maximum
'''               allowed.
'''
'''Note:          If Long Delay is needed it must be specificaly requested
'''               by setting optional bNoMax switch to true
'''
'''Updates Globals:
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME           As String = "LibApp_Wait"
   Const sLONG_DELAY          As String = " - Requested Delay exceeds "
   '''24hr * 60min * 60sec
   Const lNO_OF_SECONDS_DAY   As Long = 86400
   
   Dim lMaxWaitTime           As Long
   Dim dStartTime             As Double
   Dim dCurrentTime           As Double

   On Error GoTo ErrHnd
   
   If bNoMax Then
      lMaxWaitTime = 999
   Else
      lMaxWaitTime = 10
   End If

   If dSecondsToWait > lMaxWaitTime Then
      dSecondsToWait = lMaxWaitTime
      '''Make sure it's not a Typo
      Debug.Print sPROC_NAME & sLONG_DELAY & CStr(lMaxWaitTime)
   End If

   If bDoEvents Then
      dStartTime = Timer
      Do
         DoEvents
         dCurrentTime = Timer
         '''If passed midnight
         If dCurrentTime < dStartTime Then
            dStartTime = dStartTime - lNO_OF_SECONDS_DAY
         End If
      Loop While (dCurrentTime - dStartTime) < dSecondsToWait
   Else
      Call LibApp_Sleep(dSecondsToWait)
   End If
   
Exit Sub
ErrHnd:
   Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")

End Sub


Public Sub LibApp_Sleep(ByVal dSeconds As Double)
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
''' dSeconds      Requestee number of seconds to wait
'''
'''Description:   Convert specified number of seconds into milliseconds
'''               and wait for this time
'''
'''Updates Globals:
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "LibApp_Sleep"
   
   Dim lMilliseconds    As Long
   
On Error GoTo ErrHnd

   lMilliseconds = CLng(dSeconds * 1000)
   Call Sleep(lMilliseconds)
   
Exit Sub
ErrHnd:
   Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")
   
Exit Sub
End Sub

Private Sub M_DisplayErr(ByVal lErrNo As Long, ByVal sErrReport As String, ByVal sProcName As String)
'''************************************************************************
'''Called By:     LibApp_Error
'''
'''Arguments
''' lErrNo        Error number
''' sErrReport    Error message
''' sProcName     Error source
'''
'''
'''
'''Description:   Display specified error information
'''
'''Note:
'''Warning:
'''Updates Globals:
'''
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME           As String = "M_DisplayErr"

On Error GoTo ErrHnd


      If Not mobjShared Is Nothing Then
200      Call mobjShared.ShowVBError(lErrNo, vbCrLf & sErrReport, sProcName, msClientExeName, vbCritical)
      Else
300      Call MsgBox(sErrReport, vbOKOnly + vbCritical, msClientExeName)
      End If

Exit Sub
ErrHnd:
   sErrReport = "ERROR: " & Err.Number & ", " & Err.Description & vbCrLf _
              & "LINE: " & Erl & vbCrLf _
              & sErrReport
On Error GoTo 0
On Error Resume Next
   Call MsgBox(sErrReport, vbOKOnly, msMODULE_NAME & sPROC_NAME)

End Sub

Private Sub M_ProcessCommandLine()
'''************************************************************************
'''Called by:     Let psExeName
'''
'''Arguments:
'''
'''Description:   If CHCSII Shared is available extract and process current command line switches
'''
'''Note:
'''Uses Globals:
'''Uses Modules:  CWShared
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''
'''************************************************************************
   Const sPROC_NAME  As String = "M_ProcessCommandLine"
   
   Dim sSwithes      As String
   
On Error GoTo ErrHnd

      '''If command line is available and not Library it self(currently Lib can be tracked in the global log only)
20    If Not gobjLibStr.LibStr_bTxtEql(msClientExeName, App.EXEName) Then '<DEV
'''<DSN currently Let only 70          msSystemVersionMajor = mobjShared.AppVersion
'74          msSystemVersionRevision = mobjShared.AppBuild
'80          mbSystemVersion = True
'''>DSN
'''<DSN causes an error sometime  watch for 2 weeks   08/07/02
         '''ERROR LOGS
'On Error GoTo 0
'On Error Resume Next
100      sSwithes = UCase$(mobjShared.CmdLineSwitches("ERRLOG"))  '''
'On Error GoTo 0
'On Error GoTo ErrHnd
'''>DSN
         '''If name of the client is set
120      If Len(msClientExeName) > 0 Then
140         psErrLogName = msClientExeName & "Err.txt"   '''Setup name of the component error log
            '''If component level log is requested
160         If gobjLibStr.LibStr_lInStrPlus(sSwithes, "= ,", msClientExeName, " ,", vbTextCompare) > 0 Then
200            pbErrLog = True   '''Enable component error log
            End If
         End If
         '''If global error log is requested
300      If gobjLibStr.LibStr_lInStrPlus(sSwithes, "= ,", "GLOBAL", " ,", vbTextCompare) > 0 Or _
            gobjLibStr.LibStr_lInStrPlus(sSwithes, "= ,", "CHCSII", " ,", vbTextCompare) > 0 _
         Then
320         pbErrLogGlobal = True   '''Enable global error log
         End If
      End If
      
Exit Sub
ErrHnd:
   Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl, , False)
   

End Sub

Private Sub M_RaiseErr _
      (ByVal lErrNo As Long, ByVal sErrMsg As String, _
       ByVal sModuleName As String, ByVal sProcName As String, ByVal sInfo As String _
      )
'''************************************************************************
'''Called By:     LibApp_Error
'''               LibApp_sInitErr
'''
'''Arguments
''' lErrNo        Error Number
''' sErrMsg       Error Description
''' sModuleName   Name of the colling module (bas, class, form name...)
''' sProcName     Name of the caller procedure
''' sInfo         Additional Information
'''
'''Description:   Reformat error information and raise error
'''
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   sProcName = gobjLibStr.LibStr_sEndWith(sModuleName, ".", True, True) & sProcName
      
      Err.Raise Number:=lErrNo, _
             Source:=sProcName, _
             Description:=sErrMsg & vbCrLf & IIf(sInfo = "", "", "<i>" & sInfo)


End Sub
Public Function LibApp_objCreateInformModal() As mclsLibModalMsg
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
'''
'''Description:   Create and return inform core object with preset
'''               client specific properties: App NCID...
'''
'''Note:          LibApp_bSetClientProps must executed prior to calling this procedure
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If successful
'''                  New object of the mclsLibModalMsg class
'''                  with populated psNCID_AppSource property
'''               Else
'''                  Nothing
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "LibApp_objCreateInformModal"

   Dim objResult        As mclsLibModalMsg
   
On Error GoTo ErrHnd
   
20    Set objResult = New mclsLibModalMsg
40    objResult.psNCID_AppSource = msNCID_AppClient
60    Set objResult.pobjCWComm = pobjCWComm
900   Set LibApp_objCreateInformModal = objResult
920   Set objResult = Nothing

Exit Function
ErrHnd:
   Set LibApp_objCreateInformModal = Nothing
   Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl)

End Function

Public Property Get pbErrLog() As Boolean
   pbErrLog = mobjLibLogErr.bActive
End Property

Public Property Let pbErrLog(ByVal bActivate As Boolean)
   mobjLibLogErr.bActive = bActivate
End Property
Public Function LibApp_bSetClientProps _
      (objApp As Object, ByVal sNCID_AppClient As String, _
       Optional objCWComm As Object = Nothing, Optional ByVal sTitle As String _
      ) As Boolean
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
''' objApp           reference to App object of the client application - used to extract
'''                  ExeName version ...
''' sNCID_AppClient  NCID of the client application - used for vsSourceServerNCID parameter
'''                  when sending messages to core
''' objCWComm        Core communitaion object - used for sending messages to core
'''                  Note: Ignored if not set
'''
'''Description:      Store client information
'''
'''Updates Globals:
'''
'''Returns:       If Successful  - true
'''               Else           - false
'''
'''Programmer:    Alexander Alberstein
'''Modified:
'''************************************************************************
   Const sPROC_NAME     As String = "LibApp_bSetClientProps"
   
On Error GoTo ErrHnd
    
#If LIB_LOG Then
   Dim sReport As String
   Call gobjLog.Mtd_Message(msMODULE_NAME & sPROC_NAME, ">>> " & gsENTER)
   sReport = "Initiating client Properties. <current msClientExeName: " & msClientExeName & vbCrLf _
           & "objApp is nothing: " & gobjLibValid.LibValid_bEvaluateBoolean(objApp Is Nothing) & vbCrLf _
           & "sNCID_AppClient  : " & sNCID_AppClient & vbCrLf _
           & "sTitle           : " & sTitle & vbCrLf _
           & "objCWComm is nothing : " & gobjLibValid.LibValid_bEvaluateBoolean(objCWComm Is Nothing)
   Call gobjLog.Mtd_Message(msMODULE_NAME & sPROC_NAME, sReport)
#End If
      '''Client Version
100   msClientVersionMajor = objApp.Major
120   msClientVersionMinor = objApp.Minor
140   msClientVersionRevision = objApp.Revision
160   psExeName = objApp.EXEName
'<DSN

200   Set pobjCWComm = objCWComm
220   Set mobjLibMsg.pobjCWComm = objCWComm
240   If Not objCWComm Is Nothing Then
260      Set gobjCWComm = objCWComm
      End If

300   msNCID_AppClient = sNCID_AppClient
320   mobjLibMsg.psNCID_AppClient = sNCID_AppClient
      
400   If Trim$(sTitle) = "" Then
420      msClientTitle = objApp.Title
      Else
440      msClientTitle = sTitle
      End If
460   mobjLibMsg.psClientTitle = msClientTitle
   
#If LIB_LOG Then
   sReport = "Result Setting" & vbCrLf _
           & "msClientExeName       : " & msClientExeName & vbCrLf _
           & "msNCID_AppClient      : " & msNCID_AppClient & vbCrLf _
           & "msClientTitle         : " & msClientTitle & vbCrLf _
           & "pobjCWComm is nothing : " & gobjLibValid.LibValid_bEvaluateBoolean(pobjCWComm Is Nothing)
           
   Call gobjLog.Mtd_Message(msMODULE_NAME & sPROC_NAME, sReport)
#End If
'>DSN
   
   LibApp_bSetClientProps = True
   
#If LIB_LOG Then
   Call gobjLog.Mtd_Message(msMODULE_NAME & sPROC_NAME, "<<< " & gsEXIT)
#End If
   
Exit Function
ErrHnd:
   LibApp_bSetClientProps = False
   Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, _
                     gsLINE_ & Erl & gsPROGRAM & " " & msClientExeName _
                    )
End Function

Public Property Get pobjConn() As CHCSII_CONN.Conn
   Set pobjConn = mobjConn
End Property

Public Property Get pobjCWShared() As CWShared
   '''<DSN may not be needed - instantiated in the initialize event
   'If mobjShared Is Nothing Then
   '   Set mobjShared = New CWShared
   'End If
   '''>
   Set pobjCWShared = mobjShared
   
End Property

Public Property Get pobjMsgBox() As gclsLibMsg
   Set pobjMsgBox = mobjLibMsg
End Property

Public Property Get pobjNewPickList() As mclsLibPickList
'''************************************************************************
'''Called By:
'''
'''Arguments
'''
'''Description:   Create and return new Pick List object
'''               with preset reference to comm object, Client app NCID and Title
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If Successful
'''                  New Pick List object
'''               else
'''                  nothing
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Get pobjNewPickList"
   
   Dim objResult        As mclsLibPickList
   
On Error GoTo ErrHnd

20    Set objResult = New mclsLibPickList
100   Set objResult.pobjCWComm = gobjCWComm
120   objResult.psNCID_AppClient = msNCID_AppClient
140   objResult.psTitle = msClientTitle

999   Set pobjNewPickList = objResult

Exit Property
ErrHnd:
   Set pobjNewPickList = Nothing
   Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl)

End Property

Public Property Get psErrLogName() As String
   psErrLogName = mobjLibLogErr.sFileName
End Property

Public Property Let psErrLogName(ByVal sFullPathName As String)
   If Not Trim$(sFullPathName) = "" Then
      mobjLibLogErr.sFileName = sFullPathName
   End If
End Property
Public Property Let psExeName(ByVal sName As String)
   msClientExeName = Trim$(sName)
   Call M_ProcessCommandLine
End Property

Public Property Let psSystemVersionAppPathName(ByVal sVal As String)
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
''' sVal
'''   Full Path name of the application version of which will be used as system version
'''
'''Description:   Access specified file and set system version to its version
'''
'''
'''Note:          If not set - system verion will not be included in error reports
'''
'''
'''Updates Globals:
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME           As String = "psSystemVersionAppPathName"

   Dim sErrMsg                As String
   
On Error GoTo ErrHnd
   
   mbSystemVersion = False '''Used to evaluate whether to use system version
   
   Call gobjLibFile.LibFile_sVersion _
   (sVal, msSystemVersionMajor, msSystemVersionMinor, msSystemVersionRevision, , sErrMsg)
   
   If sErrMsg = "" Then
      mbSystemVersion = True
   Else
      Call LibApp_Error _
         (1, gobjLibStr.LibStr_sBuildSentence(gsFAILED, gsTO, gsSET, gsSYSTEM, gsVERSION), _
          msMODULE_NAME, sPROC_NAME, sErrMsg _
         )
   End If
   
Exit Property
ErrHnd:
   Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")

End Property

Private Sub Class_Initialize()
   Const sPROC_NAME     As String = "Class_Initialize"
   
On Error GoTo ErrHnd

20    Call Lib_Load  '''Use count
40    Set mobjLibMsg = New gclsLibMsg
   
100   With mobjLibLogErr
120      .lMaxSize = 20000
140      .lKeepSize = 10000
160      .bDateStamp = True
180      .bTimeStamp = True
200      .bProcStamp = True
220      .bComputerStamp = True
240      .bSuppressErrors = True
260      .lStampOffset = 40
280      .sFileName = "ErrorLog.txt"
      End With

500   Set mobjShared = New CWShared
520   Set mobjConn = New CHCSII_CONN.Conn

Exit Sub
ErrHnd:
   Call MsgBox(gsERROR & Err.Number & ", " & Err.Description & vbCrLf & "Line: " & Erl, vbCritical + vbOKOnly, sPROC_NAME)
   
End Sub


Private Sub Class_Terminate()

   Set mobjShared = Nothing
   Set mobjLibLogErr = Nothing
   Set pobjCWComm = Nothing
   Set mobjLibMsg = Nothing
   Set mobjConn = Nothing
   Call Lib_Unload   '''Use count
   
End Sub
Public Function LibApp_sVersion _
      (Optional ByVal bLabel As Boolean = False, Optional bSystemVersion As Boolean = False) As String
'''************************************************************************
'''Called By:     LibApp_Error
'''
'''Arguments:
''' bLabel Request to Preface version with Label
''' bSystemVersion Indicates whether to return system or components version
'''
'''Description:   Build Version in the Mag.Min.Rev format
'''
'''Note:          components version is set in LibApp_bSetClientProps
'''Warning:
'''Updates Globals:
'''
'''Returns:       if successful
'''                  Application Version
'''               Else
'''                  Empty string
'''
'''Programmer: Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME              As String = "LibApp_sVersion"
   
   Dim sResult                   As String
   Dim sMajor       As String
   Dim sMinor       As String
   Dim sRevision    As String
   
On Error GoTo ErrHnd
   
      If bSystemVersion Then
         sMajor = msSystemVersionMajor
         sMinor = msSystemVersionMinor
         sRevision = msSystemVersionRevision
      Else
         sMajor = msClientVersionMajor
         sMinor = msClientVersionMinor
         sRevision = msClientVersionRevision
      End If

100   sResult = Replace(sMajor & "." & sMinor & "." & sRevision, "..", ".")
      
      If bLabel Then
         sResult = gsVERSION & ": " & sResult
      End If

      LibApp_sVersion = sResult
   
Exit Function
ErrHnd:
   LibApp_sVersion = ""
   Call LibApp_Error _
   (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl & vbCrLf & gsCLIENT & " " & msClientExeName)
   
End Function

Public Property Get pbErrLogGlobal() As Boolean
   pbErrLogGlobal = gobjLibLogErr.bActive
End Property
Public Property Let pbErrLogGlobal(ByVal bActivate As Boolean)
   gobjLibLogErr.bActive = bActivate
End Property
Public Property Get psErrLogGlobalName() As String
   psErrLogGlobalName = gobjLibLogErr.sFileName
End Property
Public Property Let psErrLogGlobalName(ByVal sFullPathName As String)
   If Not Trim$(sFullPathName) = "" Then
      gobjLibLogErr.sFileName = sFullPathName
   End If
End Property

