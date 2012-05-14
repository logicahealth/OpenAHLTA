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
Attribute VB_Name = "mclsLibLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''************************************************************************
'''Module: LibLog.cls - mclsLibLog Log Maintenance class
'''
'''Programer: Alexander Alberstein
'''
'''Description:
''' Allows to Create and Maintain multiple Log files for any application
'''
'''Suggested Use:
'''
''' .sFileName = sFileName
'''   Full Path Name of the Log File to be mainteined
'''   If file is not set
'''      Will Default to <APP.ExeName>.log (Name of the executable with LOG extension) in the
'''      Current Directory
'''
''' call Mtd_Open
'''   Will open Log file and keep it open until call Mtd_Close is made
'''   If currently open will close and then open it
'''   Note:
'''      If file is not opened call to Mtd_Message (sProcNames, sMsg) will open and close
'''      the file on each call
'''
''' .bDateStamp = True/False
'''   Sets or Cancels Request for adding Date info for each msg
'''   If True
'''      Current Date Info will be added to each Message
'''
''' .bTime = True/False
'''   Sets or Cancels Request for adding Time info for each msg
'''   If True
'''      Current Time Info will be added to each Message
'''
''' .bProcStamp
'''   Sets or Cancels Request for adding Procedure Name info for each msg
'''   If True
'''      Procedure Name will be added to each Message
'''
''' .bComputerStamp
'''   Sets or Cancels Request for adding Computer Name info for each msg
'''   If True
'''      Computer Name will be added to each Message
'''
''' .lMaxSize
'''   Sets Maximum allowed size of the log file. If not set will default to 50K
'''   If you want to have unlimited size set this property to lNO_LIMIT Property
'''   e.g. .lMaxSize = .lNO_LIMIT
'''   Note: File size will be reduced to lKeepSize if it's size is more then maximum allowed
'''         but only during set name procedure and ones per 50 Mtd_Message calls.
'''         This is done to optimize performance - therefore physical file size may exceed
'''         size specified in lMaxSize if it's reached this size before evaluation occurred.
'''
''' .lKeepSize
'''   Size to keep when maximum size is reached and file size is reduced
'''   The size will be reduced to the size specified in this property.
'''   Last (lKeepSize) bytes are preserved
'''   If lKeepSize is greater then lMaxSize lKeepSize is preserved
'''   Note: Value lKeepSize must be in the range 0 to 50000 bytes.
'''         If it is set to the value greater then 50000 it will be reset to 50000
'''         If it is set to the value less    then 0     it will be reset to 0
'''   If this Property is not set it will default to 10K or MaxSize whichever is less
'''
''' .bActive = True/False
'''   If Set to False
'''      Mtd_Message will do nothing until bActive = True
'''  else
'''      Mtd_Message will Log the Message until bActive = False
'''
''' .bSuppressErrors
'''   If Set to true
'''      Run time Errors will be ignored
'''   else
'''      Run time errors are reported in the message box but not loged
'''
'''  lStampOffset
'''   Determines how many spaces to shift Stamp Information to the right
'''   valid range from 0 to 80
'''
''' call Mtd_Message (sProcNames, sMsg)
'''    Depending on Active Mode will either Log a Massage to the Log file
'''    that was set by Log_SetFileName or do Nothing
'''
'''Note: File Will be Created if not Exist. If Mtd_Open is used -
'''      Mtd_Close must be used in order to close the Log file
'''
'''************************************************************************
Option Explicit
''''''''''''''''''''''''''''' Constants '''''''''''''''''''''''''''''''''''

Const msMODULE_NAME     As String = "mclsLibLog."

Const mlNO_LIMIT        As Long = -99
Const mlDEFAULT_SIZE    As Long = 50000
Const mlDEFAULT_KEEP    As Long = 10000
Const mlMAX_KEEP_SIZE   As Long = 200000

''''''''''''''''''''''''''''' Variables '''''''''''''''''''''''''''''''''''

Dim msFileName       As String
Dim miLogChannel     As Integer  '''file channel of the Log file
Dim mbActive         As Boolean  '''Determines whether Log Message will be written
'''Set to true by Mtd_Open and set to False by Mtd_Close
Dim mbKeepOpen       As Boolean
Dim mbDateStamp      As Boolean  '''Determines whether Date will be added to each msg
Dim mbTimeStamp      As Boolean  '''Determines whether Time will be added to each msg
Dim mbProcStamp      As Boolean  '''Determines whether Procedure Name will be added to each msg
Dim mbComputerStamp  As Boolean  '''Determines whether Computer Name will be added to each msg
Dim mlMaxSize        As Long     '''Upper Limit of the file size before it will be deleted
Dim mlKeepSize       As Long     ''''Size to keep when maximum size is reached and file size is reduced
Dim mbKeepSizeSet    As Boolean
Dim mbSuppressErrors As Boolean
Dim mlStampOffset    As Long
Dim msComputerName   As String
'''Character that will be used to draw the line between messages. defaults to "-"
Public psSeparatorChar  As String

Public Property Let bActive(bValue As Boolean)
   mbActive = bValue
   
End Property


Public Property Get bActive() As Boolean
   bActive = mbActive
   
End Property

Public Property Let bComputerStamp(ByVal bValue As Boolean)
'''************************************************************************
'''Called by:    Any
'''
'''Description:  Set local Computer Name flag and initialize
'''              required data
'''
'''Uses Globals: None
'''
'''Uses Modules: None
'''
'''************************************************************************
   Const sPROC_NAME     As String = "bComputerStamp"
   
On Error GoTo ErrHnd

   If bValue Then
      msComputerName = gobjLibShell.LibShell_sLocalComputerName
      mbComputerStamp = Not (msComputerName = "")
   Else
      mbComputerStamp = bValue
   End If
   
Exit Property
ErrHnd:
   Call gobjLibApp.LibApp_Error _
   (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", , Not mbSuppressErrors, False)
End Property

Public Property Get bComputerStamp() As Boolean
   bComputerStamp = mbComputerStamp
   
End Property

Public Property Let bDateStamp(ByVal bValue As Boolean)
   mbDateStamp = bValue
   
End Property

Public Property Get bDateStamp() As Boolean
   bDateStamp = mbDateStamp
   
End Property

Public Property Let bProcStamp(ByVal bValue As Boolean)
   mbProcStamp = bValue
   
End Property

Public Property Get bProcStamp() As Boolean
   bProcStamp = mbProcStamp
   
End Property

Public Property Let bSuppressErrors(ByVal bValue As Boolean)
   mbSuppressErrors = bValue
   
End Property

Public Property Let bTimeStamp(ByVal bValue As Boolean)
   mbTimeStamp = bValue
   
End Property

Public Property Get bTimeStamp() As Boolean
   bTimeStamp = mbTimeStamp
   
End Property

Public Property Get lKeepSize() As Long
   lKeepSize = mlKeepSize
End Property

Public Property Let lKeepSize(ByVal lValue As Long)
'''************************************************************************
'''Called by:    Any
'''
'''Description:
'''   Size to keep when maximum size is reached and file size is reduced
'''   The size will be reduced to the size specified ib this property.
'''   Last lKeepSize bytes are preserved
'''   If lKeepSize is greater then lMaxSize - KeepSize is preserved
'''   Note: Value lKeepSize must be in the range 0 to mlMAX_KEEP_SIZE bytes.
'''         If it is set to the value greater then mlMAX_KEEP_SIZE
'''            it will be reset to mlMAX_KEEP_SIZE
'''         If it is set to the value less    then 0     it will be reset to 0
'''
'''Uses Globals: None
'''
'''Uses Modules: None
'''
'''************************************************************************
   If lValue > 0 Then
      If lValue < mlMAX_KEEP_SIZE Then
         mlKeepSize = lValue
      Else
         mlKeepSize = mlMAX_KEEP_SIZE
      End If
   Else
      mlKeepSize = 0
   End If
      
   mbKeepSizeSet = True
   
End Property

Public Property Let lMaxSize(ByVal lValue As Long)
   mlMaxSize = lValue
   
End Property

Public Property Get lMaxSize() As Long
   lMaxSize = mlMaxSize
   
End Property

Public Property Get lNO_LIMIT() As Long
   lNO_LIMIT = mlNO_LIMIT
   
End Property

Public Property Get lStampOffset() As Long
   lStampOffset = mlStampOffset
   
End Property

Public Property Let lStampOffset(ByVal lValue As Long)
   On Error Resume Next
   
   mlStampOffset = lValue
   
   If mlStampOffset < 0 Then
      mlStampOffset = 0
   ElseIf lStampOffset > 80 Then
      mlStampOffset = 80
   End If
   
End Property

Private Sub M_EvaluateMaxSize()
'''************************************************************************
'''Called by:    sFileName Let
'''              Mtd_Message
'''
'''Description:  If file size is more then maximum allowed - delete file
'''              Maximum size = mlNO_LIMIT - means: No Size Limit
'''
'''Note:         If file name is not set Will Default to APP.ExeName & .log in the
'''              App Directory
'''
'''Uses Globals: None
'''
'''Uses Modules: None
'''
'''************************************************************************
   Const sPROC_NAME     As String = "M_EvaluateMaxSize"
   
   Dim lCurrentSize     As Long
   
On Error GoTo ErrHnd
   
   '''If Unlimited size is not specificaly requested
   If Not (mlMaxSize = mlNO_LIMIT) Then
      '''If programmer forget to set the size or set it to 0
      If mlMaxSize = 0 Then
         lMaxSize = mlDEFAULT_SIZE
      End If
      If Not mbKeepSizeSet Then
         lKeepSize = mlDEFAULT_KEEP
      End If
      '''If file is Set
      If (msFileName <> "") Then
         If gobjLibFile.LibFile_bExist(msFileName) Then
            
            '''Get Size of the current log
            lCurrentSize = FileLen(msFileName)
            '''If current the file size kept on the disk is more the maximum specified And more
            '''then the size specified to be kept (if somebody for some reason set Keep>Max)
            If (lCurrentSize > mlMaxSize) And (lCurrentSize > mlKeepSize) Then
               
               
               Call M_ReduceFileSize
               
            End If
         End If
      End If
   End If

Exit Sub
ErrHnd:
   Call gobjLibApp.LibApp_Error _
   (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", , Not mbSuppressErrors, False)
End Sub



Private Sub M_ReduceFileSize()
'''************************************************************************
'''Called by:    M_EvaluateMaxSize
'''
'''Description:  Reduce size of the log file to the size specified in KeepSize
'''
'''Uses Globals: None
'''
'''Uses Modules: None
'''
'''************************************************************************
   Const sPROC_NAME     As String = "M_ReduceFileSize"
   
   Dim bFileWasOpen     As Boolean
'   Dim lCurrentSize     As Long
'   Dim sKeepSize        As String
   Dim sComments        As String
On Error GoTo ErrHnd
   
   '''If file is currently open
   If mbKeepOpen Then
      '''close the file to flash all data to the disk
      Call Mtd_Close
      bFileWasOpen = True '''mbKeepOpen will be set to false in Mtd_Close
   End If
   
   sComments = "File Size has reached(lMaxSize): " & CStr(mlMaxSize) & " bytes" _
                & " and was reduced to (lKeepSize): " & CStr(mlKeepSize) & " bytes" _
                & vbNewLine & vbNewLine
   sComments = String$(mlStampOffset, psSeparatorChar) _
             & Format$(Time, "HH:MM:SS") & " " & Format$(Date, "MM/DD/YY") & vbCrLf _
             & sComments

   Call gobjLibFile.LibFile_bReduceFileSize(msFileName, mlKeepSize, sComments, Not mbSuppressErrors)
   
   If bFileWasOpen Then
      Call Mtd_Open   '''reopen the file
   End If
   
Exit Sub
ErrHnd:
   Call gobjLibApp.LibApp_Error _
   (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", , Not mbSuppressErrors, False)
End Sub

Public Property Let sFileName(ByVal sValue As String)
'''************************************************************************
'''Called by:    Any
'''
'''Description:  Set Full Path Name of the log file to the specified value
'''              If file size is more then maximum allowed - delete file
'''
'''Uses Globals: None
'''
'''Uses Modules: None
'''
'''************************************************************************
   msFileName = sValue
   Call M_EvaluateMaxSize
   
End Property

Public Property Get sFileName() As String
   sFileName = msFileName
   
End Property

Public Sub Mtd_Message(ByVal sProcName As String, ByVal sMsg As String)
'''************************************************************************
'''Called by: Any
'''
'''Description:
''' Depending on Active Mode will either Log a Massage to the Log file
''' or do Nothing
''' If mbKeepOpen is True - Do not Open and Close the file on Each Call
'''
'''Uses Globals: None
'''
'''Uses Modules: Error
'''
'''************************************************************************
   Const sPROC_NAME        As String = "Mtd_Message"
   Const sDEFAULT_EXT      As String = ".log"
   Static slCount          As Long
   
   Dim sDate               As String * 8
   Dim sTime               As String * 8
   Dim sComputerName       As String * 20
   Dim sProcedureName      As String
   
   On Error GoTo ErrHnd
   
   If mbActive Then
      '''If file name is not set
      
      If Trim$(msFileName) = "" Then
         '''Default to APP.ExeName & .log in the App Directory - call property procedure
         sFileName = gobjLibStr.LibStr_sEndWith(App.Path, "\", True, True) _
                   & Trim$(App.EXEName) & sDEFAULT_EXT
      End If
      
      If mbDateStamp Then
         sDate = Format$(Date, "MM/DD/YY")
      End If
      If mbTimeStamp Then
         sTime = Format$(Time, "HH:MM:SS")
      End If
      If mbProcStamp Then
         sProcedureName = sProcName
      End If
      If mbComputerStamp Then
         sComputerName = msComputerName
      End If
      If slCount > 5 Then
         slCount = 0
         Call M_EvaluateMaxSize
      Else
         slCount = slCount + 1
      End If
      
      sMsg = String$(mlStampOffset, psSeparatorChar) _
           & sTime & " " _
           & sDate & " " _
           & sComputerName & " " _
           & sProcedureName _
           & vbCrLf _
           & sMsg
           
      If Not mbKeepOpen Then
         miLogChannel = gobjLibFile.LibFile_lOpen(msFileName, lFM_APPEND, Not mbSuppressErrors)
      End If
      
      Print #miLogChannel, sMsg
         
      If Not mbKeepOpen Then
         Call gobjLibFile.LibFile_bCloseIfOpen(miLogChannel, Not mbSuppressErrors)
      End If
   End If
   
Exit Sub
ErrHnd:
   Call gobjLibApp.LibApp_Error _
   (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", , Not mbSuppressErrors, False)
   Call gobjLibFile.LibFile_bCloseIfOpen(miLogChannel, False)
End Sub



Public Sub Mtd_Close()
'''************************************************************************
'''Called by: Any
'''
'''Description:
''' close Current Log file and indicate that it is Closed. Works in
''' conjunction with Mtd_Open
'''
'''Uses Globals:
'''
'''Uses Modules: Error
'''
'''************************************************************************
   Const sPROC_NAME As String = "Mtd_Close"
   
On Error GoTo ErrHnd
   
   mbKeepOpen = Not gobjLibFile.LibFile_bCloseIfOpen(miLogChannel, Not mbSuppressErrors)
   
Exit Sub
ErrHnd:
   Call gobjLibApp.LibApp_Error _
   (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", , Not mbSuppressErrors, False)
End Sub

Public Sub Mtd_Open()
'''************************************************************************
'''Called by: Any
'''
'''Description:
''' Open Current Log file and indicate that it is open. Works in
''' conjunction with Mtd_Close
'''
'''Uses Globals: None
'''
'''Uses Modules: Error
'''
'''************************************************************************
   Const sPROC_NAME As String = "Mtd_Open"
   
On Error GoTo ErrHnd

   '''If currently opened
   If mbKeepOpen Then
      Call Mtd_Close
   End If
   
   miLogChannel = gobjLibFile.LibFile_lOpen(msFileName, lFM_APPEND, Not mbSuppressErrors)
   mbKeepOpen = miLogChannel > 0
   
Exit Sub
ErrHnd:
   Call gobjLibApp.LibApp_Error _
   (Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "", , Not mbSuppressErrors, False)
End Sub






Private Sub Class_Initialize()
   mbDateStamp = True
   mbTimeStamp = True
   mbProcStamp = True
   mlStampOffset = 40
   psSeparatorChar = "-"
End Sub


