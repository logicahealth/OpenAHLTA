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
Attribute VB_Name = "gclsLibFile"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''************************************************************************
'''Module:        gclsLibFile/LibFile.cls
'''
'''Description:   Generic File operation procedures
'''
'''Modules Required: LibErr
'''
'''Author         : Alexander Alberstein
'''Modified       :
'''               :
'''************************************************************************
Option Explicit

''''''''''''''' Constanst Section '''''''''''''''''''''''''''''''''''''''''

Const msMODULE_NAME              As String = "gclsLibFile."

Public Enum pnumFileMode
   [lFM_undefined]
   lFM_APPEND
   lFM_RANDOM
   lFM_INPUT
   lFM_BINARY
   lFM_OUTPUT
End Enum

Private Const mlPATH_NOT_FOUND   As Long = 76

''''''''''''''' Records Section '''''''''''''''''''''''''''''''''''''''''

''''''''''''''' API

Private Type mudtFVI_FileVersionInfo
    '''There is data in the following two dwords, but it is for Windows internal
    '''   use and we should ignore it
    yIgnore(1 To 8)        As Byte
    '''Signature As Long
    '''StrucVersion As Long
    iFileVerPart2          As Integer
    iFileVerPart1          As Integer
    iFileVerPart4          As Integer
    iFileVerPart3          As Integer
    iProductVerPart2       As Integer
    iProductVerPart1       As Integer
    iProductVerPart4       As Integer
    iProductVerPart3       As Integer
    lFileFlagsMask         As Long     '''VersionFileFlags
    lFileFlags             As Long     '''VersionFileFlags
    lFileOS                As Long     '''VersionOperatingSystemTypes
    lFileType              As Long
    lFileSubtype           As Long     '''VersionFileSubTypes
    '''Ignore following two dwords
    '''DateHighPart As Long
    '''DateLowPart As Long
    yIgnored(1 To 8)       As Byte
End Type

''''''''''''''' API Section '''''''''''''''''''''''''''''''''''''''''

Private Declare Function GetShortPathName Lib "Kernel32" _
                  Alias "GetShortPathNameA" _
                  (ByVal lpszLongPath As String, ByVal lpszShortPath As String, _
                   ByVal cchBuffer As Long _
                  ) As Long
Private Declare Function API_lGetFileVersionInfoSize Lib "version.dll" _
                  Alias "GetFileVersionInfoSizeA" (ByVal sFile As String, lpLen As Long _
                  ) As Long
Private Declare Function API_lGetFileVersionInfo Lib "version.dll" _
                  Alias "GetFileVersionInfoA" _
                  (ByVal sFile As String, ByVal lpIgnored As Long, ByVal lpSize As Long, _
                   ByVal lpBuf As Long _
                  ) As Long
Private Declare Function API_lVerQueryValue Lib "version.dll" _
                  Alias "VerQueryValueA" _
                  (ByVal lpBuf As Long, ByVal szReceive As String, lpBufPtr As Long, lLen As Long _
                  ) As Long
Private Declare Sub API_CopyMemory Lib "Kernel32" _
                  Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)


Public Function LibFile_bCloseIfOpen _
      (ByVal lFileNumber As Long, Optional bPopErr As Boolean = True) As Boolean
'''************************************************************************
'''Called by: Any
'''
'''Arguments:
''' lFileNumber   file number of the file requested to be closed
''' bPopErr       Optional Request to Pop message if unexpected error occures
'''
'''Description:   If file number is more then zero
'''                  close it, reset number to zero and return true
'''               Else
'''                  Return True
'''
'''Note:          File Number Zero is interpreted as file is already closed
'''
'''
'''Uses Globals: None
'''
'''Uses Modules: Error
'''
'''Returns:       True  - If specified file is or was closed
'''               False - if error
'''
'''Programmer:    Alexander Alberstein
'''
'''************************************************************************
   Const sPROC_NAME  As String = "LibFile_bCloseIfOpen"
   
On Error GoTo ErrHnd

   If lFileNumber > 0 Then
      Close lFileNumber
      lFileNumber = 0
   End If
   
   LibFile_bCloseIfOpen = True
   
Exit Function
ErrHnd:
   LibFile_bCloseIfOpen = False
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, _
                                gsFILE & gsNUMBER & ": " & lFileNumber, , bPopErr, False _
                               )
End Function


Public Function LibFile_bDelete _
       (ByVal sFilePathName As String, Optional ByVal bPopErr As Boolean = True) As Boolean
'''************************************************************************
'''Called by: Any
'''
'''Arguments:
''' sFilePathName Full path Name of the file
''' bPopErr       Optional Request to Pop message if unexpected error occures
'''
'''Description:   Delete specified file
'''
'''
'''Uses Globals: None
'''
'''Uses Modules: Error
'''
'''Returns:       If successful  - true
'''               else           - false
'''
'''Programmer:    Alexander Alberstein
'''
'''************************************************************************
   Const sPROC_NAME  As String = "LibFile_bDelete"
   
   
On Error GoTo ErrHnd

   Kill sFilePathName
   
   LibFile_bDelete = True
   
Exit Function
ErrHnd:
   LibFile_bDelete = False
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, _
                                gsFILE & ": " & sFilePathName, , bPopErr, False _
                               )
End Function

Public Function LibFile_bPathExist _
      (ByVal sPathName As String, _
       Optional ByVal bPopErr As Boolean = False, _
       Optional ByVal bLogErr As Boolean = False, _
       Optional ByRef sErrMsg As String _
      ) As Boolean
'''************************************************************************
'''Called By:     Any
'''
'''Arguments
''' sFilePathName Full or relative path name
''' bPopErr       Optional Request to Pop message if unexpected error occures
''' bLogErr       Optional Request to Log message if unexpected error occures
''' sErrMsg       Used to return error information if value is invalid
'''
'''Description:   Evaluate whether specified is valid directory
'''
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:      True   - If the directory path exists
'''              False  - Else
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = msMODULE_NAME & "LibFile_bPathExist"

   Dim bResult          As Boolean
   Dim lAttr            As Long
   
On Error GoTo ErrHnd
   
   sErrMsg = ""
   
   '''If path is specified
   If Trim$(sPathName) <> "" Then
      
      lAttr = GetAttr(sPathName)
      If lAttr And vbDirectory Then
         bResult = True
      Else
         sErrMsg = gobjLibStr.LibStr_sBuildSentence(gsPATH, gsMUST, gsBE, "a", gsDIRECTORY)
      End If
   Else
      sErrMsg = gobjLibStr.LibStr_sBuildSentence(gsPATH, gsIS, gsNOT, gsSPECIFIED)
   End If
    
   If Not sErrMsg = "" Then
      Call gobjLibApp.LibApp_Error _
         (1, sErrMsg, msMODULE_NAME, sPROC_NAME, sPathName, , bPopErr, bLogErr)
   End If
   
   LibFile_bPathExist = bResult
   
Exit Function
ErrHnd:
   sErrMsg = Err.Description
   LibFile_bPathExist = False
   Call gobjLibApp.LibApp_Error _
      (Err.Number, sErrMsg, msMODULE_NAME, sPROC_NAME, sPathName, , bPopErr, False)
   
End Function

Public Function LibFile_lOpenWithLock _
   (ByVal sFilePathName As String, ByVal lMode As pnumFileMode, _
    Optional ByVal bPopErr As Boolean = True _
   ) As Long
'''************************************************************************
'''Called by: Any
'''
'''Arguments:
''' sFilePathName Full path Name of the file
''' lMode         File Mode. Use pnumFileMode
''' bPopErr       Optional Request to Pop message if unexpected error occures
'''
'''Description:   Open specified file in requested mode with lock
'''
'''
'''Uses Globals: None
'''
'''Uses Modules: Error
'''
'''Returns:       If successful  - File number
'''               else           - Zero
'''
'''Programmer:    Alexander Alberstein
'''
'''************************************************************************
   Const sPROC_NAME  As String = "LibFile_lOpenWithLock"
   
   Dim lResult       As Long
   
On Error GoTo ErrHnd

   lResult = FreeFile
   If lResult > 0 Then
      Select Case lMode
         Case lFM_BINARY
            Open sFilePathName For Binary Lock Read Write As #lResult
         Case lFM_INPUT
            Open sFilePathName For Input Lock Read Write As #lResult
         Case lFM_OUTPUT
            Open sFilePathName For Output Access Write Lock Read Write As #lResult
         Case lFM_RANDOM
            Open sFilePathName For Random Access Read Lock Read Write As #lResult
         Case lFM_APPEND
            Open sFilePathName For Append Access Read Write Lock Read Write As #lResult
         Case Else
            lResult = 0
            Call gobjLibApp.LibApp_Error _
               (0, gsUNDEFINED & gsMODE, msMODULE_NAME, sPROC_NAME, _
                gsFILE & ": " & sFilePathName & ", " & gsMODE & ": " & lMode, , bPopErr, False _
               )
      End Select
   End If

   LibFile_lOpenWithLock = lResult
   
Exit Function
ErrHnd:
   LibFile_lOpenWithLock = 0
   Call gobjLibApp.LibApp_Error _
      (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, _
       gsFILE & ": " & sFilePathName & ", " & gsMODE & ": " & lMode, , bPopErr, False _
      )
End Function

Public Function LibFile_bReOpen _
   (ByRef lFileNumber As Long, ByVal sFilePathName As String, ByVal lMode As pnumFileMode, _
    Optional ByVal bPopErr As Boolean = True _
   ) As Boolean
'''************************************************************************
'''Called by: Any
'''
'''Arguments:
''' lFileNumber   File number of the file requested to be reopened
''' sFilePathName Full path Name of the file to be reopened
''' lMode         File Mode. Use pnumFileMode
''' bPopErr       Optional Request to Pop message if unexpected error occures
'''
'''Description:   Open specified file is opened close it
'''               Open specified file in the requested mode
'''
'''Note:          If file is not currently opened(File Number Zero)
'''               The file will be opened
'''
'''
'''Uses Globals: None
'''
'''Uses Modules: Error
'''
'''Returns:       If successful  - True
'''               else           - False
'''
'''Programmer:    Alexander Alberstein
'''
'''************************************************************************
   Const sPROC_NAME  As String = "LibFile_bReOpen"
   
   Dim bResult       As Boolean
   
On Error GoTo ErrHnd

   If LibFile_bCloseIfOpen(lFileNumber, bPopErr) Then
      lFileNumber = LibFile_lOpen(sFilePathName, lMode, bPopErr)
      bResult = (lFileNumber > 0)
   End If

   LibFile_bReOpen = bResult
   
Exit Function
ErrHnd:
   LibFile_bReOpen = False
   Call gobjLibApp.LibApp_Error _
      (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, _
           gsFILE & ": " & sFilePathName & vbCrLf & _
           gsMODE & ": " & lMode & vbCrLf & _
           gsFILE & gsNUMBER & ": " & lFileNumber, , bPopErr, False _
      )
End Function

Public Function LibFile_lOpen(ByVal sFilePathName As String, ByVal lMode As pnumFileMode, _
                           Optional ByVal bPopErr As Boolean = True _
                           ) As Long
'''************************************************************************
'''Called by: Any
'''
'''Arguments:
''' sFilePathName Full path Name of the file
''' lMode         File Mode. Use pnumFileMode
''' bPopErr       Optional Request to Pop message if unexpected error occures
'''
'''Description:   Open specified file in requested mode
'''
'''
'''Uses Globals: None
'''
'''Uses Modules: Error
'''
'''Returns:       If successful  - File number
'''               else           - Zero
'''
'''Programmer:    Alexander Alberstein
'''
'''************************************************************************
   Const sPROC_NAME  As String = "LibFile_lOpen"
   
   Dim lResult       As Long
   
On Error GoTo ErrHnd

   lResult = FreeFile
   If lResult > 0 Then
      Select Case lMode
         Case lFM_BINARY
            Open sFilePathName For Binary As #lResult
         Case lFM_INPUT
            Open sFilePathName For Input As #lResult
         Case lFM_OUTPUT
            Open sFilePathName For Output As #lResult
         Case lFM_RANDOM
            Open sFilePathName For Random As #lResult
         Case lFM_APPEND
            Open sFilePathName For Append As #lResult
         Case Else
            lResult = 0
            Call gobjLibApp.LibApp_Error _
               (0, gsUNDEFINED & gsMODE, msMODULE_NAME, sPROC_NAME, _
                gsFILE & ": " & sFilePathName & ", " & gsMODE & ": " & lMode, , bPopErr, False _
               )
      End Select
   End If

   LibFile_lOpen = lResult
   
Exit Function
ErrHnd:
   LibFile_lOpen = 0
   Call gobjLibApp.LibApp_Error _
      (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, _
       gsFILE & ": " & sFilePathName & ", " & gsMODE & ": " & lMode, , bPopErr, False _
      )
End Function


Function LibFile_bExist _
        (ByVal sFilePathName As String, _
         Optional ByVal bPopErr As Boolean = False, _
         Optional ByVal bLogErr As Boolean = False, _
         Optional ByRef sErrMsg As String _
        ) As Boolean
'''************************************************************************
'''Called By:     LibFile_bOpenForInput or Any
'''
'''Arguments:
''' sFilePathName Full or relative path name of the file
''' bPopErr       Optional Request to Pop message if unexpected error occures
''' bLogErr       Optional Request to Log message if unexpected error occures
''' sErrMsg       Used to return error information if value is invalid
'''
'''Description:   This procedure checks if sFilePathName exists.
'''
'''Note:
'''
'''Updates Globals:
'''
'''Returns:       True  - if file exists
'''               False - if file does not exist
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME        As String = "LibFile_bExist"

   Const lFILE_NOT_FOUND   As Long = 53
   
   Static bIn_static       As Boolean
   
On Error GoTo ErrHnd

   If Not bIn_static Then
      bIn_static = True
      sErrMsg = ""
      If Len(Trim$(sFilePathName)) > 0 Then
         '''Attempt to Get the file date
         Call FileDateTime(sFilePathName)
      Else
         sErrMsg = gobjLibStr.LibStr_sBuildSentence(gsFILE, gsIS, gsNOT, gsSPECIFIED)
      End If
      bIn_static = False
   Else
      sErrMsg = "Reentry error"
   End If
   
   ''''No Error - Return True
   LibFile_bExist = Len(sErrMsg) = 0

Exit Function
ErrHnd:
   '''Return False in any Error
   LibFile_bExist = False
   sErrMsg = Err.Description
   '''Determine the type of error
   Select Case Err.Number
      Case lFILE_NOT_FOUND, mlPATH_NOT_FOUND
         '''Do Nothing
         '''File not found and Path not found are
         '''expected errors for this function.
         'If bPopErr Then
         '   Call gobjLibMsg.LibMsg_lOKCancel(lMSGOK_FILE_NOT_FOUND, , , sFilePathName)
         'End If
         
         Call gobjLibApp.LibApp_Error _
            (Err.Number, sErrMsg, msMODULE_NAME, sPROC_NAME, _
             gsFILE & ": " & sFilePathName, , bPopErr, False _
            )
      Case Else
         '''Unexpected Error - Inform User - always log unexpected error
         Call gobjLibApp.LibApp_Error _
            (Err.Number, sErrMsg, msMODULE_NAME, sPROC_NAME, _
             gsFILE & ": " & sFilePathName, , bPopErr, False _
            )
   End Select
   bIn_static = False
   
End Function


Function LibFile_bOpenForInput _
        (ByVal sFilePathName As String, ByRef lFileNumber As Long) _
         As Boolean
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
''' sFilePathName Full Path name of the file to be opened
''' lFileNo       file handle to be returned byref
'''
'''Description:   Open specified file for input, Return  handle to it byref.
'''
'''Note:
'''Warning:
'''Updates Globals:
'''
'''Returns:       True  - If Successful
'''               False - Else
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME           As String = "LibFile_bOpenForInput"

On Error GoTo ErrHnd

   lFileNumber = LibFile_lOpen(sFilePathName, lFM_INPUT, True)
   
   LibFile_bOpenForInput = (lFileNumber > 0)
   
Exit Function
ErrHnd:
   LibFile_bOpenForInput = False
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, _
                                gsFILE & sFilePathName, , True, False _
                               )
End Function

Public Function LibFile_sAbsolutePath _
      (ByVal sRelativePath As String, Optional ByVal sBasePath As String) As String
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
''' sRelativePath Path relative to base path
''' sBasePath     Optional base path if not specified current system path (CurDir) will be used
'''
'''
'''Description:   Based on information in sRelativePath and base paths return absolute path
'''
'''Note:
'''
'''Example:       Assuming the current directory is c:\test
'''               LibFile_sAbsolutePath("..\") or ("..")                   will return C:\
'''               LibFile_sAbsolutePath("..\..", "D:\A\B\C")               will return D:\A
'''               LibFile_sAbsolutePath("\Other") or ("Other")             will return C:\test\Other
'''               LibFile_sAbsolutePath("..\Other") or ("..Other")         will return C:\Other
'''
'''Updates Globals:
'''Returns:       If Successful  - Absolute path
'''               Else           - Empty string
'''
'''Programmer:    Alexander Alberstein
'''Modified:
'''************************************************************************
   Const sPROC_NAME        As String = "LibFile_sAbsolutePath"
   
   Dim sResult             As String
   Dim lNoOfLevels         As Long
   Dim bErr                As Boolean
   Dim lTargetLevelPos     As Long
   Dim sRelPathReminder    As String
   
On Error GoTo ErrHnd
   
   If sBasePath = "" Then
      sBasePath = CurDir
   End If
   
   '''If starts with drive leter
   If Mid$(sRelativePath, 2, 1) = ":" Then
      '''it is absolute path
      sResult = sRelativePath
   Else
      With gobjLibStr
         '''Determine how many levels do we need to go up the path if any and get reminder of the path
         '''Count the dots and get everything after it last occurance
         sRelPathReminder = .LibStr_sAfterLastDelimiter(sRelativePath, "..", , lNoOfLevels, bErr)
         
         If Not bErr Then
            '''If relative path indicates moving up the base path
            If lNoOfLevels > 0 Then
               '''Keep the name part of the path
               sRelativePath = sRelPathReminder
               '''Adjust Base path
               '''Make sure that first occurance of the folder delimiter(level indicater)
               '''is before the name of the folder. So when we are moving up we actually moving
               '''number of folders - not number of folders minus trailing indicator
               sBasePath = .LibStr_sEndWith(sBasePath, "\", False, True)
               '''Find position pointer at required level
               lTargetLevelPos = .LibStr_lOccurrencePos(sBasePath, "\", lNoOfLevels, False)
               '''If no errors
               If lTargetLevelPos > 0 Then
                  '''Keep remaining part of the base path
                  sBasePath = Left$(sBasePath, lTargetLevelPos - 1)
               Else
                  bErr = True
               End If
            '''Else
               '''keep relative and base paths as is - contains only name part of the path
            End If
            If Not bErr Then
               '''Build absolute path
               sResult = .LibStr_sEndWith(sBasePath, "\", True, True) _
                       & .LibStr_sStartWith(sRelativePath, "\", False, True)
            End If
         End If
      End With
   End If
   
   LibFile_sAbsolutePath = sResult
   
Exit Function
ErrHnd:
   LibFile_sAbsolutePath = ""
   Call gobjLibApp.LibApp_Error _
      (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, sRelativePath & vbCrLf & sBasePath, , , False)

End Function

Function LibFile_sOpenWithAssociatedExe _
        (ByRef lResult As Long, ByVal sFileName As String, _
         Optional sFolder As String = "", Optional ByVal sParameters As String = "", _
         Optional lHWnd As Long = 0 _
        ) As String
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
''' lResult       Return Code - > 32 is success
''' sFileName     Name of the file
''' sFolder       Optional file folder
''' sParameters   Optional file command line Parameters
''' lHWnd         Optional Window Handle
'''
'''Description:   This Procedure
'''
'''Note:
'''
'''Updates Globals:
'''Returns:       If Successful  - Empty string
'''               Else           - Error Message
'''
'''Programmer:    Alexander Alberstein
'''Modified:
'''************************************************************************
   Const sPROC_NAME        As String = "LibFile_sOpenWithAssociatedExe"

On Error GoTo ErrHnd
   
   LibFile_sOpenWithAssociatedExe = gobjLibShell.LibShell_sExecute _
                                    (lResult, "", sFileName, sParameters, sFolder, 1, lHWnd)
Exit Function
ErrHnd:
   LibFile_sOpenWithAssociatedExe = "* " & Err.Description
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, _
                                gsFILE & sFileName, , True, False _
                               )
End Function

Public Function LibFile_sShortFileName(ByVal sLongPathName As String) As String
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
''' sLongPathName - Long path name of the file or directory
'''
'''Description:   Get short name of the file or directory specified in
'''               the long path name
'''
'''Note:          This procedure will remove Path information e.g.
'''               LibFile_sShortFileName _
'''               ("\\Ucsgsd1\Public\Software\McAfee Updates\12_21_99\4056updt.exe")
'''               Will return: 4056updt.exe
'''
'''Updates Globals:
'''Returns:
'''   if successful -  short Name of the file or directory
'''   Else          - Empty string
'''
'''Programmer:       Alexander Alberstein
'''***********************************************************************
   Const sPROC_NAME        As String = "LibFile_sShortFileName"
   
   Dim sShortPath       As String * 255
   Dim lBufferLen       As Long
   Dim lLastBackSlash   As Long
   Dim sResult          As String
   Dim lNameLen         As Long
   Dim bFound           As Boolean
  
On Error GoTo ErrHnd

   lBufferLen = Len(sShortPath) + 1
   lBufferLen = GetShortPathName(sLongPathName, sShortPath, lBufferLen)
   
   If lBufferLen > 0 Then
      '''Extract full path
      sResult = Left$(sShortPath, lBufferLen)
      '''Find begining of the file name
      '''Start from the first character
      lLastBackSlash = lBufferLen
      Do
         bFound = Mid$(sResult, lLastBackSlash, 1) = "\"
         
         If Not bFound Then
            lLastBackSlash = lLastBackSlash - 1
            lNameLen = lNameLen + 1
         Else
            lLastBackSlash = lLastBackSlash + 1
         End If
         
      Loop Until (lLastBackSlash < 1) Or (lNameLen > 12) Or bFound
      '''Loop until found or file name becames more the 12 characters or
      '''And of the string
      If bFound Then
         sResult = Mid$(sResult, lLastBackSlash, lNameLen)
      Else
         sResult = ""
      End If
   Else
      sResult = ""
   End If

   '''Returm Result
   LibFile_sShortFileName = sResult

Exit Function
ErrHnd:
   LibFile_sShortFileName = ""
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, _
                                gsFILE & sLongPathName, , True, False _
                               )
End Function

Public Function LibFile_sShortPathName(ByVal sLongPathName As String) As String
'''************************************************************************
'''Called By:     Any
'''
'''Arguments:
''' sLongPathName - Long path name of the file or directory
'''
'''Description:      Get short Full path name name of the file or directory
'''                  specified in the long  path name
'''
'''Note:
'''
'''Updates Globals:
'''Returns:
'''   if successful - Short full path Name of the file or directory
'''   Else          - Empty string
'''
'''Programmer:       Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME        As String = "LibFile_sShortPathName"
   
   Dim sShortPath          As String * 255
   Dim lBufferLen          As Long
   Dim sResult             As String
   
On Error GoTo ErrHnd

   lBufferLen = Len(sShortPath) + 1
   lBufferLen = GetShortPathName(sLongPathName, sShortPath, lBufferLen)
   
   If lBufferLen > 0 Then
      sResult = Left$(sShortPath, lBufferLen)
   Else
      sResult = ""
   End If

   '''Returm Result
   LibFile_sShortPathName = sResult

Exit Function
ErrHnd:
   LibFile_sShortPathName = ""
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, _
                                sLongPathName, , True, False _
                               )
End Function
Public Function LibFile_sVersion _
      (ByVal sFileName As String, _
       Optional ByRef sMajor As String, Optional ByRef sMinor As String, _
       Optional ByRef sBuild As String, Optional ByRef sVerPart4 As String, _
       Optional ByRef sErrMsg As String _
      ) As String
'''************************************************************************
'''Called by:     Any
'''
'''Arguments:
''' sFileName     Full Path name of the requested file
''' sMajor,
''' sMinor,
''' sBuild        Optional - separate parts of the version returned by ref
''' sErrMsg       Error information if any to be returned byref
'''
'''Description:   Return Version of the specified file
'''
'''Uses Globals:  None
'''
'''Uses Modules:  None
'''
'''Returns:
'''               If successful - Version of the specified file
'''               Else          - Empty string
'''************************************************************************

   Const sPROC_NAME     As String = "LibFile_sVersion"
   
   Dim sResult          As String
   Dim uFVIRec          As mudtFVI_FileVersionInfo
   
On Error GoTo ErrHnd

   sErrMsg = M_sGetFileVersionInfo(sFileName, uFVIRec)
   
   If sErrMsg = "" Then
      sMajor = Trim$(uFVIRec.iFileVerPart1)
      sMinor = Trim$(uFVIRec.iFileVerPart2)
      sBuild = Trim$(uFVIRec.iFileVerPart3)
      sVerPart4 = Trim$(uFVIRec.iFileVerPart4)
      
      sResult = gobjLibStr.LibStr_sBuildDelimited(".", sMajor, sMinor, sBuild) _
              & IIf(sVerPart4 = "", "", "." & sVerPart4)
   End If
   
   LibFile_sVersion = sResult
   
Exit Function
ErrHnd:
   sErrMsg = Err.Description
   LibFile_sVersion = ""
   Call gobjLibApp.LibApp_Error(Err.Number, sErrMsg, msMODULE_NAME, sPROC_NAME, "", , False, False)

End Function

Private Function M_sGetFileVersionInfo _
               (ByVal sFileName As String, ByRef uFVIRec As mudtFVI_FileVersionInfo) _
               As String
'''************************************************************************
'''Called by:     LibFile_sVersion
'''
'''Arguments:
''' sFileName     Full Path name of the requested file
''' uFVIRec       file information record
'''
'''Description:   Get Version information of the requested file into
'''               specified record
'''
'''Uses Globals:  None
'''
'''Uses Modules:  None
'''
'''Returns:
'''               If successful - Empty string
'''               Else          - Error message
'''************************************************************************
   Const sPROC_NAME     As String = "M_sGetFileVersionInfo"
   Const sEXE           As String = "\"
   
   Dim sResult          As String
   Dim lVerSize         As Long
   Dim lTemp            As Long
   Dim lRet             As Long
   Dim bInfo()          As Byte
   Dim lpBuffer         As Long
   
On Error GoTo ErrHnd

   If LibFile_bExist(sFileName) Then
      lVerSize = API_lGetFileVersionInfoSize(sFileName, lTemp)
      ReDim bInfo(lVerSize)
      If lVerSize > 0 Then
         lRet = API_lGetFileVersionInfo(sFileName, lTemp, lVerSize, VarPtr(bInfo(0)))
         If lRet <> 0 Then
            lRet = API_lVerQueryValue(VarPtr(bInfo(0)), sEXE, lpBuffer, lVerSize)
            If lRet <> 0 Then
               Call API_CopyMemory(uFVIRec, ByVal lpBuffer, lVerSize)
            Else
               sResult = "VerQueryValue"
            End If
         Else
           sResult = "GetFileVersionInfo"
         End If
      Else
         sResult = "GetFileVersionInfoSize"
      End If
      If Not sResult = "" Then
         sResult = sResult & " - failed"
      End If
   Else
      sResult = "File not found"
   End If

   M_sGetFileVersionInfo = sResult
   
Exit Function
ErrHnd:
   sResult = Err.Description
   M_sGetFileVersionInfo = sResult
   Call gobjLibApp.LibApp_Error(Err.Number, sResult, msMODULE_NAME, sPROC_NAME, "", , False, False)

End Function

Private Function M_bGetFileVerStruct(ByVal sFileName As String, _
                                     ByRef uFVIRec As mudtFVI_FileVersionInfo, _
                                     ByRef sErrMsg As String _
                                    ) As Boolean
'''************************************************************************
'''Called by:     Any
'''
'''Arguments:
''' sFileName     Full Path name of the requested file
''' uFVIRec       File version record to be returned byref
''' sErrMsg       Error information if any to be returned byref
'''
'''Description:   Get and return file version information for the specified file
'''
'''
'''Uses Globals:  None
'''
'''Uses Modules:  None
'''
'''Returns:
'''               if successful - true
'''               Else          - false
'''************************************************************************
   Const sPROC_NAME     As String = "M_bGetFileVerStruct"
   
   Dim bResult          As Boolean
   
On Error GoTo ErrHnd
   
   
   M_bGetFileVerStruct = bResult
   
Exit Function
ErrHnd:
   M_bGetFileVerStruct = False
   Call gobjLibApp.LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", , True, False)

End Function



Public Function LibFile_bReduceFileSize _
      (ByVal sFileName As String, ByVal lKeepSize As Long, _
       Optional ByVal sOppNotes As String, Optional ByVal bPopErr As Boolean = True _
      ) As Boolean
'''************************************************************************
'''Called by:     Any
'''
'''Arguments:
''' sFileName     Full Path name of the requested file
''' lKeepSize     Number of the last bytes requested to be preserved
''' sOppNotes     Optional comments to be saved in the modified file
''' bPopErr       Optional Request to Pop message if unexpected error occures
'''
'''Description:   Reduce size of the requested file to the specified size
'''               Keep lKeepSize bytes of data starting from the EOF
'''
'''Uses Globals:  None
'''
'''Uses Modules:  None
'''
'''Returns:
'''               if successful - true
'''               Else          - false
'''************************************************************************
   Const sPROC_NAME     As String = "LibFile_bReduceFileSize"
   
   Dim bResult          As Boolean
   Dim lFileChannel     As Long
   Dim lCurrentSize     As Long
   Dim sKeepSize        As String
   
On Error GoTo ErrHnd
   
   '''Get the Size of the specified file
   lCurrentSize = FileLen(sFileName)
   
   '''If reduction is needed
   If lCurrentSize > lKeepSize Then
      '''If there is anything to keep
      If lKeepSize > 0 Then
         '''Size the buffer to store all data intended to be preserved
         sKeepSize = Space$(lKeepSize)
         lFileChannel = LibFile_lOpen(sFileName, lFM_BINARY)
         If lFileChannel > 0 Then
            '''Set Position to read last lKeepSize bytes
            Seek #lFileChannel, (lCurrentSize - lKeepSize)
            '''Read lKeepSize bytes
            Get #lFileChannel, , sKeepSize
         End If
      End If
      If LibFile_bCloseIfOpen(lFileChannel, bPopErr) Then
         Call Kill(sFileName)
         '''Save data requested to be preserved
         lFileChannel = LibFile_lOpen(sFileName, lFM_OUTPUT, bPopErr)
         If lFileChannel > 0 Then
            Print #lFileChannel, sOppNotes & sKeepSize
            bResult = LibFile_bCloseIfOpen(lFileChannel, bPopErr)
         End If
      End If
   Else
      bResult = True
   End If
   
   LibFile_bReduceFileSize = bResult
   
Exit Function
ErrHnd:
   LibFile_bReduceFileSize = False
   Call gobjLibApp.LibApp_Error _
   (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", , bPopErr, False)

End Function
