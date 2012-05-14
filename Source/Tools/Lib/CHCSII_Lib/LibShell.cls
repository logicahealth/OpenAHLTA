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
Attribute VB_Name = "gclsLibShell"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''************************************************************************
'''Module:        gclsLibShell/LibShell.cls
'''
'''Description:   Various shell functions
'''
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

Const msMODULE_NAME        As String = "gclsLibShell."

Private Const mlSTR_LEN    As Long = 512  '''Used to init API string buffers

'''LibShell_lSendMessage lwMsgID
Public Enum pnumSMT_SendMessageTypes
   SMT_ShowDropDown = &H14F
   SMT_GetDropDownWidth = &H15F
   SMT_SetDropDownWidth = &H160
   SMT_SetTabStops = &H192
   SMT_LineScroll = &HB6         '''SendMessage(TextBox.hwnd, SMT_LineScroll, 0, ByVal <lNoOfLines>)
End Enum

''''''''''''''' Variables Section '''''''''''''''''''''''''''''''''''''''''

Private msSysPath          As String
Private msWinPath          As String

''''''''''''''' Records Section '''''''''''''''''''''''''''''''''''''''''

Private Type mudtPINF_ProcessInfo
    lProcessHnd      As Long
    lThreadHnd       As Long
    lProcessID       As Long
    lThreadID        As Long
End Type

Public Type mudtSINF_StartupInfo
    lRecLen          As Long
    lReserved        As Long
    lDesktop         As Long
    lTitle           As Long
    lX               As Long
    lY               As Long
    lXSize           As Long
    lYSize           As Long
    lXCountChars     As Long
    lYCountChars     As Long
    lFillAttribute   As Long
    lFlags           As Long
    iShowWindow      As Integer
    ibReserved2      As Integer
    lReserved2       As Long
    lStdInput        As Long
    lStdOutput       As Long
    lStdError        As Long
End Type

''''''''''''''' API Section '''''''''''''''''''''''''''''''''''''''''''''''

Private Declare Function API_ShellExecute Lib "shell32.dll" _
       Alias "ShellExecuteA" (ByVal hwnd As Long, _
       ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, _
       ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function API_WaitForSingleObject Lib "Kernel32" _
      Alias "WaitForSingleObject" _
      (ByVal hProcess As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function API_WaitForInputIdle Lib "user32" _
      Alias "WaitForInputIdle" _
      (ByVal hProcess As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function API_CreateProcessA Lib "Kernel32" _
      Alias "CreateProcessA" _
      (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, _
       ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
       ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
       ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
       lpStartupInfo As mudtSINF_StartupInfo, lpProcessInformation As mudtPINF_ProcessInfo _
      ) As Long

Private Declare Function API_CloseHandle Lib "Kernel32" _
      Alias "CloseHandle" _
      (ByVal hObject As Long) As Long

Private Declare Function API_GetKeyState Lib "user32" _
       Alias "GetKeyState" (ByVal nVirtKey As Long) As Integer

Private Declare Function API_GetSystemDirectory Lib "kernel32.dll" _
      Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Declare Function API_GetWindowsDirectory Lib "kernel32.dll" _
      Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Declare Function API_GetTempPath Lib "kernel32.dll" _
      Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Declare Function API_GetTempFileName Lib "kernel32.dll" _
      Alias "GetTempFileNameA" _
      (ByVal lpszPath As String, ByVal lpPrefixString As String, _
       ByVal wUnique As Long, ByVal lpTempFileName As String _
      ) As Long

Private Declare Function API_GetComputerName Lib "Kernel32" _
      Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Declare Function API_SendMessage Lib "user32" _
      Alias "SendMessageA" _
      (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Function LibShell_bCloseHandle _
      (ByVal lProcessHnd As Long, Optional ByRef sErrMsg As String) As Boolean
'''************************************************************************
'''Called By:     LibShell_bSyncExecute or any
'''
'''Arguments:
'''
'''
''' sErrMsg
'''   Error i
'''
'''Description:   Close specified handle (proccess)
'''
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If successful
'''                  true
'''               else
'''                  false
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME                 As String = "LibShell_bCloseHandle"
   
On Error GoTo ErrHnd
  
   sErrMsg = ""
   Call Err.Clear
   Call API_CloseHandle(lProcessHnd)

   LibShell_bCloseHandle = Err.LastDllError = 0
   
Exit Function
ErrHnd:
   LibShell_bCloseHandle = False
   sErrMsg = Err.Description
   Call gobjLibApp.LibApp_Error(Err.Number, sErrMsg, msMODULE_NAME, sPROC_NAME, lProcessHnd)

End Function

Public Function LibShell_lSendMessage _
      (ByVal lHWnd As Long, ByVal lwMsgID As pnumSMT_SendMessageTypes, _
       ByVal lwParam As Long, Optional vParam As Variant = 0& _
      ) As Long
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
''' lHWnd
''' lwMsgID
''' lwParam
''' vParam
'''
'''Description:   Wraper for SendMessage API
'''
'''
'''
'''Updates Globals:
'''
'''Returns:       If Successful
'''                  Depends on the message
'''               Else
'''                  Zero
'''
'''Programmer:    Alexander Alberstein
'''Modified:
'''************************************************************************
   Const sPROC_NAME     As String = "LibShell_lSendMessage"
    
   Dim lResult          As Long
   
On Error GoTo ErrHnd

100   Select Case VarType(vParam)
         Case vbNull, vbEmpty
200         lResult = API_SendMessage(lHWnd, lwMsgID, lwParam, ByVal 0&)
         Case vbLong, vbInteger
300         lResult = API_SendMessage(lHWnd, lwMsgID, lwParam, ByVal CLng(vParam))
         Case Else
400         lResult = API_SendMessage(lHWnd, lwMsgID, lwParam, vParam)
      End Select
   
999   LibShell_lSendMessage = lResult
   
Exit Function
ErrHnd:
   LibShell_lSendMessage = 0
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, lwMsgID)

End Function

Public Function LibShell_sExecute _
      (ByRef lResult As Long, _
       ByVal sOperation As String, ByVal sFile As String, ByVal sParameters As String, _
       ByVal sDirectory As String, ByVal lCommand As Long, _
       Optional lHWnd As Long = 0 _
      ) As String
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
''' lResult       Return Code - > 32 is success
'''
'''Description:   Wraper for ShellExecute Windows API function.
'''               This function launches a file with its associated program
'''               as listed in the Windows registry.
'''
'''Updates Globals:
'''
'''Returns:       If Successful  - Empty string
'''               Else           - Error Message
'''
'''Programmer:    Alexander Alberstein
'''Modified:
'''************************************************************************
   Const sPROC_NAME     As String = "LibShell_sExecute"
    
   Dim sResult          As String
   Dim objLibMSPnt      As New mclsLibMsPnt
   
On Error GoTo ErrHnd
    
   Call objLibMSPnt.Mtd_SetMousePointer(Screen)
   
   If Trim$(sOperation) = "" Then
      sOperation = vbNullString
   End If
   If Trim$(sParameters) = "" Then
      sParameters = vbNullString
   End If
   If Trim$(sDirectory) = "" Then
      sDirectory = vbNullString
   End If
   
   lResult = API_ShellExecute _
            (lHWnd, _
             sOperation, _
             sFile, _
             sParameters, _
             sDirectory, _
             lCommand _
            )
   Select Case lResult
       Case Is > 32
          ''' Return value > 32 = success
       Case 0, 11
           sResult = gobjLibStr.LibStr_sBuildSentence(gsINVALID, gsFILE, gsFORMAT)
       Case 2
          sResult = gobjLibStr.LibStr_sBuildSentence(gsFILE, gsNOT, gsFOUND)
       Case 3
          sResult = gobjLibStr.LibStr_sBuildSentence(gsPATH, gsNOT, gsFOUND)
       Case 8
          sResult = gobjLibStr.LibStr_sBuildSentence(gsOUT, gsOF, gsMEMORY)
      Case 26
         sResult = gobjLibStr.LibStr_sBuildSentence(gsSHARING, gsVIOLATION)
      Case 32
         sResult = gobjLibStr.LibStr_sBuildSentence(gsUNRECOGNIZED, gsFILE, gsFORMAT)

      Case Else
         sResult = gsERROR
   End Select
   
   LibShell_sExecute = sResult
   
Exit Function
ErrHnd:
   sResult = Err.Description
   lResult = Err.Number
   LibShell_sExecute = sResult
   Call gobjLibApp.LibApp_Error(lResult, sResult, msMODULE_NAME, sPROC_NAME, gsFILE & sFile)

End Function

Public Function LibShell_bStartProcess _
      (ByVal sCommand As String, ByRef lProcessHnd As Long, _
       Optional bHide As Boolean, Optional sErrMsg As String _
      ) As Boolean
'''************************************************************************
'''Called By:     LibShell_bSyncExecute or any
'''
'''Arguments:
''' sCommand
'''   Start process (app,program...) command. E.g. name of the executable
''' lProcessHnd   Process handle to be returned byref
'''
''' bHide         Request to hide window
'''
''' sErrMsg       Error information if any. Returned byref
'''
'''Description:   Start process and is successful return its handle
'''
'''Note:
'''
'''Warning:       Do not send empty string to CreateProcess
'''               It will not error out - it will crash, hard
'''
'''Updates Globals:
'''
'''Returns:       If successful
'''                  true
'''               else
'''                  false
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME                 As String = "LibShell_bStartProcess"

   Const lSTARTF_USESHOWWINDOW      As Long = &H1
   Const lSW_HIDE                   As Long = 0
   Const lNORMAL_PRIORITY_CLASS     As Long = &H20&
   
   Dim uPINF                        As mudtPINF_ProcessInfo
   Dim uSINF                        As mudtSINF_StartupInfo
   
On Error GoTo ErrHnd
  
    sErrMsg = ""
    lProcessHnd = -1
    sCommand = Trim$(sCommand)
    
    If Len(sCommand) > 0 Then
      With uSINF
        .lRecLen = Len(uSINF)
        If bHide Then
           .lFlags = lSTARTF_USESHOWWINDOW
           .iShowWindow = lSW_HIDE
        End If
      End With
      Call Err.Clear
      '''Start the application
      Call API_CreateProcessA _
           (0&, sCommand, 0&, 0&, 1&, lNORMAL_PRIORITY_CLASS, 0&, 0&, uSINF, uPINF)
           
      If Err.LastDllError <> 0 Then
         sErrMsg = "DLL Error: " & Err.LastDllError
      End If
      lProcessHnd = uPINF.lProcessHnd
   Else
      sErrMsg = gobjLibStr.LibStr_sBuildSentence(gsPROCESS, gsIS, gsNOT, gsSPECIFIED)
   End If
   
   LibShell_bStartProcess = Len(sErrMsg) = 0
   
Exit Function
ErrHnd:
   LibShell_bStartProcess = False
   sErrMsg = Err.Description
   Call gobjLibApp.LibApp_Error(Err.Number, sErrMsg, msMODULE_NAME, sPROC_NAME, sCommand)

End Function

Public Function LibShell_bSyncExecute _
(ByVal sCommand As String, _
 Optional ByVal lTimeOut As Long, _
 Optional ByVal bHide As Boolean, _
 Optional ByVal bWaitUntilLoaded As Boolean, _
 Optional ByRef sErrMsg As String _
) As Boolean
'''************************************************************************
'''Called By:     any
'''
'''Arguments:
''' sCommand
'''   Start process (app,program...) command. E.g. name of the executable
''' bHide
'''   Indicates whether to run the process hidden
''' lTimeOut
'''   Period time before retorning without waiting for process to complete.
'''   Positive values are interpreted as seconds, negative as miliseconds.
'''   Infinite wait if zero.
''' bWaitUntilLoaded
'''   If True
'''      Wait until process(object) is initialized - ready to be used
'''
'''   else (default)
'''      Wait until process(object) terminates
'''
''' sErrMsg
'''
'''Description:
'''
'''Note:
'''
'''Warning:
'''   if bWaitUntilLoaded is true and lTimeOut
'''   Do not Hide the process if it requires user input in order to terminate
'''
'''Updates Globals:
'''
'''Returns:       If successful
'''                  true
'''               else
'''                  false
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME                 As String = "LibShell_bSyncExecute"
   Const lPROCESS_FINISHED          As Long = &H0
   Const lTIMEOUT_NEVER             As Long = -1&

   Dim lResult                      As Long
   Dim lProcessHnd                  As Long
   Dim lMilliseconds                As Long
   
On Error GoTo ErrHnd
  
   sErrMsg = ""
   
   If lTimeOut = 0 Then
      lTimeOut = lTIMEOUT_NEVER
   ElseIf lTimeOut > 0 Then
      '''convert to miliseconds
      lTimeOut = lTimeOut * 1000
   '''Negative values are milliseconds - use it as is
   End If
   
   '''If was able to start the process
   If LibShell_bStartProcess(sCommand, lProcessHnd, bHide, sErrMsg) Then
      Call Err.Clear
      '''If requested to wait for process to initialize - not terminate
      '''e.g. Wait until MSWord is completely loaded
      If bWaitUntilLoaded Then
         lResult = API_WaitForInputIdle(lProcessHnd, lTimeOut)
      Else
         '''Wait for process to finish
         lResult = API_WaitForSingleObject(lProcessHnd, lTimeOut)
      End If
      
      If Err.LastDllError <> 0 Then
         sErrMsg = "DLL Error: " & Err.LastDllError
      End If
      
      '''Make sure process is closed
      Call LibShell_bCloseHandle(lProcessHnd)
      
      Select Case lResult
         Case lPROCESS_FINISHED
            '''Completed successfully - Do nothing
         Case Else   '''258
            '''Time out - most likely
            sErrMsg = gobjLibStr.LibStr_sBuildSentence(gsTIME, gsOUT)
      End Select
   Else
      sErrMsg = gobjLibStr.LibStr_sBuildSentence(gsFAILED, gsTO, gsINITIATE, gsPROCESS) & ". " & sErrMsg
   End If
   
   LibShell_bSyncExecute = Len(sErrMsg) = 0
   
Exit Function
ErrHnd:
   LibShell_bSyncExecute = False
   sErrMsg = Err.Description
   Call gobjLibApp.LibApp_Error(Err.Number, sErrMsg, msMODULE_NAME, sPROC_NAME, sCommand)

End Function

Public Function LibShell_lGetKeyState(ByVal lVirtKey As Long, Optional ByRef bErr As Boolean) As Long
'''************************************************************************
'''Called By:     LibApp_KeyPressed or any
'''
'''Arguments:
''' lVirtKey      Virtual Key value state of which requested to be evaluated
''' bErr          Indicates whether error occured
'''
'''Description:   Wraper for
'''
'''Note:          If the high-order bit is 1, the key is down; otherwise, it is up.
'''               If the low-order bit is 1, the key is toggled.
'''               A key, such as the CAPS LOCK key, is toggled if it is turned on.
'''               The key is off and untoggled if the low-order bit is 0.
'''
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If successful
'''                  State of the the requsted Key
'''               else
'''                  0
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "LibShell_lGetKeyState"

On Error GoTo ErrHnd
   
   '''Return Result
   LibShell_lGetKeyState = API_GetKeyState(lVirtKey)
   
Exit Function
ErrHnd:
   bErr = True
   LibShell_lGetKeyState = 0
   Call gobjLibApp.LibApp_Error _
   (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsKEY & ": " & CStr(lVirtKey))

End Function

Public Function LibShell_bKeyPressed _
               (ByVal lVirtKey As KeyCodeConstants, Optional ByRef bErr As Boolean) As Boolean
'''************************************************************************
'''Called By:     any
'''
'''Arguments:
''' lVirtKey      Virtual Key value state of which requested to be evaluated
''' bErr          Indicates whether error occured
'''
'''Description:   Evalaute specified Key
'''Note:
'''Warning:
'''Returns:       If successful
'''                  If key was pressed(Up or down)   - true
'''                  Else                             - false
'''               else
'''                  False
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "LibShell_bKeyPressed"

On Error GoTo ErrHnd
   
   '''Return Result
   LibShell_bKeyPressed = LibShell_lGetKeyState(lVirtKey) <= -127
   
Exit Function
ErrHnd:
   bErr = True
   LibShell_bKeyPressed = False
   Call gobjLibApp.LibApp_Error _
   (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsKEY & ": " & CStr(lVirtKey))

End Function

Public Function LibShell_sLocalComputerName() As String
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
'''
'''Returns:       If Successful  - Name of the local computer
'''               Else           - Empty string
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "LibShell_sLocalComputerName"
   
   Dim sResult          As String
   Dim lSize            As Long
   Static bIn_static    As Boolean
   
On Error GoTo ErrHnd
   
100   If Not bIn_static Then
         bIn_static = True
120      sResult = gobjLibStr.LibStr_sBuildSentence(gsFAILED, gsTO, gsREAD, gsCOMPUTER, gsNAME)
         '''Make sure there is enough room in the input buffer
140      lSize = 256
160      sResult = Space(lSize)
         '''If was able to get name of the local computer
500      If Not API_GetComputerName(sResult, lSize) = 0 Then
520         sResult = Left$(sResult, lSize)  '''Remove spaces and null
         End If
         bIn_static = False
      End If
      
999   LibShell_sLocalComputerName = sResult
  
Exit Function
ErrHnd:
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE & Erl)
   bIn_static = False
   
End Function

Public Property Get psSystemDirectory(Optional ByVal bBackSlash As Boolean) As String
'''************************************************************************
'''Called By:     any
'''
'''Arguments:
''' bBackSlash
'''   Optional request to return result ending with backslash. Defaults to false
'''
'''Description:   If first time retrieve, store and return system path
'''               else return stored path
'''Note:
'''Warning:
'''Returns:       If successful
'''                  Path of the System Directory not terminated with Backslash
'''               else
'''                  Empty String
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Get psSystemDirectory"

   Dim lLen             As Long
   
On Error GoTo ErrHnd
   
   '''If not already set
   If Len(msSysPath) = 0 Then
      msSysPath = Space(mlSTR_LEN)
      lLen = API_GetSystemDirectory(msSysPath, mlSTR_LEN)
      msSysPath = Trim$(Left(msSysPath, lLen))
   End If
   If Len(msSysPath) > 0 Then
      msSysPath = gobjLibStr.LibStr_sEndWith(msSysPath, "\", bBackSlash, True)
   End If
   
   psSystemDirectory = msSysPath
   
Exit Property
ErrHnd:
   psSystemDirectory = ""
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")

End Property

Public Property Get psWindowsDirectory(Optional ByVal bBackSlash As Boolean) As String
'''************************************************************************
'''Called By:     any
'''
'''Arguments:
''' bBackSlash
'''   Optional request to return result ending with backslash. Defaults to false
'''
'''Description:   If first time retrieve, store and return windows path
'''               else return stored path
'''Note:
'''Warning:
'''Returns:       If successful
'''                  Path of the Windows Directory not terminated with Backslash
'''               else
'''                  Empty String
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "psWindowsDirectory"

   Dim lLen             As Long
   
On Error GoTo ErrHnd
   
   '''If not already set
   If msWinPath = "" Then
      msWinPath = Space(mlSTR_LEN)
      lLen = API_GetWindowsDirectory(msWinPath, mlSTR_LEN)
      msWinPath = Trim$(Left(msWinPath, lLen))
   End If
   If Len(msWinPath) > 0 Then
      msWinPath = gobjLibStr.LibStr_sEndWith(msWinPath, "\", bBackSlash, True)
   End If
   
   psWindowsDirectory = msWinPath
   
Exit Property
ErrHnd:
   psWindowsDirectory = ""
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")

End Property

Public Property Get psTempDirectory(Optional ByVal bBackSlash As Boolean) As String
'''************************************************************************
'''Called By:     any
'''
'''Arguments:
''' bBackSlash
'''   Optional request to return result ending with backslash. Defaults to false
'''
'''Description:   Retrieve and return path of the Temporary directory
'''
'''Note:          See also psNextTempFile
'''
'''
'''Warning:       Do Not cache - Temp Path can be changed during User session
'''
'''Returns:       If successful
'''                  path of the Temporary directory
'''               else
'''                  Empty String
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Get psTempDirectory"

   Dim lLen             As Long
   Dim sResult          As String
   
On Error GoTo ErrHnd
   
   sResult = Space(mlSTR_LEN)
   lLen = API_GetTempPath(mlSTR_LEN, sResult)
   sResult = Trim$(Left(sResult, lLen))
   If Len(sResult) > 0 Then
      sResult = gobjLibStr.LibStr_sEndWith(sResult, "\", bBackSlash, True)
   End If
   psTempDirectory = sResult
   
Exit Property
ErrHnd:
   psTempDirectory = ""
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")

End Property

Public Property Get psNextTempFile(Optional ByVal sPath As String, Optional sPrefix As String) As String
'''************************************************************************
'''Called By:     any
'''
'''Arguments:
''' sPath
'''   Optional The path to put the temporary file into.
'''   If not specified Windows temp direcory will be used
''' sPrefix
'''   Optional The first three characters of the filename.
'''   If not specified defaults to Lib
'''
'''Description:   Generates unique filename for a temporary file and creates it.
'''               The format of the generated filename is <Path>\<FileName>.TMP
'''
'''Note:          See also psTempDirectory
'''
'''Warning:
'''Returns:       If successful
'''                  Full filename, including the path
'''               else
'''                  Empty String
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Get psNextTempFile"

   Dim lLen             As Long
   Dim sResult          As String
   
On Error GoTo ErrHnd

   '''If directory is not specified
   If sPath = "" Then
      '''Use windows default
      sPath = psTempDirectory
   End If
   
   '''If was able to get Temp path
   If Not sPath = "" Then
      If sPrefix = "" Then
         sPrefix = "Lib"
      End If
      
      sResult = Space(mlSTR_LEN)
      
      '''Returns zero if error
      If Not API_GetTempFileName(sPath, sPrefix, 0, sResult) = 0 Then
         '''Filename is returned null terminated
         sResult = Left$(sResult, InStr(sResult, vbNullChar) - 1)
      Else
         sResult = ""
         Call gobjLibApp.LibApp_Error(0, sPROC_NAME & " " & gsERROR, msMODULE_NAME, sPROC_NAME, "")
      End If
   End If
   
   psNextTempFile = sResult
   
Exit Property
ErrHnd:
   psNextTempFile = ""
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")

End Property

