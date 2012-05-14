VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LogtoFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"LogtoFile"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' Module Level Variables
    Private fso As FileSystemObject
    
    Private m_lLoggingOn As Long
    Private m_lMaxFileSize As Long
    Private szLogFileName As String
    Private mszConfigFile As String
    Private fLogFile As File
    Private tsLogStream As TextStream
    Private szComputerName As String
    Private szUserName As String
   
    Public Enum enLogIndent
         en_Entering = 1
         en_Exiting = -1
         en_NoChange = 0
    End Enum
        
    Const szFile1 As String = "C:\\SIGN_A.LOG"
    Const szFile2 As String = "C:\\SIGN_B.LOG"
    Const c_lMaxFileSize As Long = 100000


Private Sub Class_Initialize()
    
    Set fso = New FileSystemObject
    'Get windows directory and create logfile pathname
     mszConfigFile = fso.GetSpecialFolder(WindowsFolder) & "\hl7serv.ini"
    
    'Get max log file size and logging on/off from ini file
    If fso.FileExists(mszConfigFile) Then
        m_lMaxFileSize = GetPrivateProfileInt("Logging", "MaxLogSize", 100000, mszConfigFile)
        m_lLoggingOn = GetPrivateProfileInt("Logging", "Logging", 0, mszConfigFile)
    End If
     
     If m_lLoggingOn = False Then Exit Sub
     If m_lMaxFileSize < c_lMaxFileSize Then m_lMaxFileSize = c_lMaxFileSize
     
    'Get File name
    szLogFileName = szGetFileName
    
    'If file exists, set textstream object to it, otherwise create one
    If fso.FileExists(szLogFileName) Then
        Set fLogFile = fso.GetFile(szLogFileName)
        Set tsLogStream = fLogFile.OpenAsTextStream(ForAppending)
    Else
        Set tsLogStream = fso.CreateTextFile(szLogFileName, True)
        Set fLogFile = fso.GetFile(szLogFileName)
    End If

     'Get Computer and User names
    Dim lLength As Long
    Dim lResult As Long
    Const MAX_COMPUTERNAME_LENGTH As Integer = 256
    Const MAX_USERNAME_LENGTH As Integer = 256
        
    szComputerName = String$(MAX_COMPUTERNAME_LENGTH + 1, 0)
    lLength = MAX_COMPUTERNAME_LENGTH + 1
    lResult = GetComputerName(szComputerName, lLength)
    szComputerName = Left$(szComputerName, lLength)
    
    szUserName = String$(MAX_USERNAME_LENGTH + 1, 0)
    lLength = MAX_USERNAME_LENGTH + 1
    lResult = GetUserName(szUserName, lLength)
    szUserName = Left$(szUserName, lLength)
    'Check file size, switching files if necessary
    CheckFileSize

     tsLogStream.WriteLine "*****************************************"
     tsLogStream.WriteLine Now & vbTab & "Start Logging"
     With App
     tsLogStream.WriteLine .Title & " " & .Major & "." & .Minor & "." & .Revision
     End With
     tsLogStream.WriteLine "ComputerName: " & szComputerName & ", UserName: " & szUserName
     
     tsLogStream.WriteLine "*****************************************"
     tsLogStream.WriteBlankLines (1)
     
     tsLogStream.Close
     
End Sub

Public Sub LogtoFile(szProcedureName As String, lIndent As enLogIndent, _
                Optional szParamName1 As String, Optional Param1 As Variant, _
                Optional szParamName2 As String, Optional Param2 As Variant, _
                Optional szParamName3 As String, Optional Param3 As Variant, _
                Optional szParamName4 As String, Optional Param4 As Variant, _
                Optional szParamName5 As String, Optional Param5 As Variant, _
                Optional szParamName6 As String, Optional Param6 As Variant _
                )
                
    If m_lLoggingOn = False Then Exit Sub
     
    Static iIndent As Integer
    Static szIndent As String
    'KELLY DEAD 7-27-01 Dim dtDateTime As Date
    Dim szPrintLine As String
    'KELLY DEAD 7-27-01 Dim iCount As Integer
    Dim szParamDelim(1 To 6) As String
    Dim szParam(1 To 6) As String
    Dim szPrintDelim(1 To 6) As String
    Const c_szPrintDelim As String = ", "
    Dim szInorOut As String
    Static iCheckFileSize As Integer
    Const c_iCheckFileFreq As Integer = 10
    
    'Open textstream back up
    'First check that file still exists (user did not delete it)
    If Not fso.FileExists(szLogFileName) Then
        'File has been deleted, recreate it
        Set tsLogStream = fso.CreateTextFile(szLogFileName, True)
        Set fLogFile = fso.GetFile(szLogFileName)
        'Not sure why this next line is required, but if I don't open for reading before opening
        'forappending, then I get "Permission denied (Error 70)"
        Set tsLogStream = fLogFile.OpenAsTextStream(ForReading)
    End If
    Set tsLogStream = fLogFile.OpenAsTextStream(ForAppending)
    
        
    'Check file size every c_iCheckFileFreq times
     If iCheckFileSize >= c_iCheckFileFreq Then
          iCheckFileSize = 0
          CheckFileSize
     Else
          iCheckFileSize = iCheckFileSize + 1
     End If
                
    'Increase indent level
    If lIndent > 0 Then
          iIndent = iIndent + 1
          If iIndent >= 2 Then szIndent = szIndent & vbTab
          szInorOut = "Entering "
     ElseIf lIndent < 0 Then
          szInorOut = "Leaving  "
    End If
    
    ' Set up some of the print values base on what params were sent
    'If lIndent <> en_Exiting Then
        szPrintDelim(1) = "  "
        If Not IsMissing(Param1) Then
            szParamDelim(1) = "="
            szParam(1) = CStr(Param1)
            If Not IsMissing(Param2) Then
                szParamDelim(2) = "="
                szParam(2) = CStr(Param2)
                szPrintDelim(2) = c_szPrintDelim
                If Not IsMissing(Param3) Then
                    szParamDelim(3) = "="
                    szParam(3) = CStr(Param3)
                    szPrintDelim(3) = c_szPrintDelim
                    
                    If Not IsMissing(Param4) Then
                        szParamDelim(4) = "="
                        szParam(4) = CStr(Param4)
                        szPrintDelim(4) = c_szPrintDelim
                        
                        If Not IsMissing(Param5) Then
                            szParamDelim(5) = "="
                            szParam(5) = CStr(Param5)
                            szPrintDelim(5) = c_szPrintDelim
                            
                            If Not IsMissing(Param6) Then
                                szParamDelim(6) = "="
                                szParam(6) = CStr(Param6)
                                szPrintDelim(6) = c_szPrintDelim
                                
                            End If
                        End If
                    End If
                End If
            End If
        End If
    'End If
        
        
    
    'write to the logfile
    szPrintLine = szIndent & Now & vbTab & _
                 szInorOut & szProcedureName & szPrintDelim(1) & _
                 szParamName1 & szParamDelim(1) & szParam(1) & szPrintDelim(2) & _
                 szParamName2 & szParamDelim(2) & szParam(2) & szPrintDelim(3) & _
                 szParamName3 & szParamDelim(3) & szParam(3) & szPrintDelim(4) & _
                 szParamName4 & szParamDelim(4) & szParam(4) & szPrintDelim(5) & _
                 szParamName5 & szParamDelim(5) & szParam(5) & szPrintDelim(6) & _
                 szParamName6 & szParamDelim(6) & szParam(6)
        
    tsLogStream.WriteLine szPrintLine
    
    'Close the textstream until next log
    tsLogStream.Close
    
    'Decrease indent level
    If lIndent < 0 Then
          iIndent = iIndent - 1
          If Len(szIndent) > 0 Then szIndent = Left$(szIndent, Len(szIndent) - 1)
     End If
                
End Sub



'Checks the size of the log file, if it is above maxsize, then we switch files
Private Function CheckFileSize()
    Dim szOldLogName As String

    If fLogFile.Size > m_lMaxFileSize Then
        szOldLogName = szLogFileName
        'switch file names
        If szLogFileName = szFile1 Then
            szLogFileName = szFile2
        Else
            szLogFileName = szFile1
        End If
        
        'Tell full log that we are switching over
        tsLogStream.WriteBlankLines (1)
        tsLogStream.WriteLine "*****************************************"
        tsLogStream.WriteLine Now & vbTab & "Switching to log " & szLogFileName
        tsLogStream.WriteLine "*****************************************"
        
        'Remove our own lock on the file
        tsLogStream.Close
        Set tsLogStream = Nothing
        Set fLogFile = Nothing
        
        'create (clear if already exists) the file now to be used
        Set tsLogStream = fso.CreateTextFile(szLogFileName, True)
        Set fLogFile = fso.GetFile(szLogFileName)
        
        'Tell new log that we are switching over to it
        tsLogStream.WriteBlankLines (1)
        tsLogStream.WriteLine "*****************************************"
        tsLogStream.WriteLine Now & vbTab & "Switching over from log " & szOldLogName
        tsLogStream.WriteLine "*****************************************"
        
        
    End If

End Function

'Determines which of the two files to use and returns the filename.
Private Function szGetFileName() As String

    Dim fFile1 As File
    Dim fFile2 As File
    
    Dim bFileAExists As Boolean
    Dim bFileBExists As Boolean
    
    Dim dtFile1Time As Date
    Dim dtfile2Time As Date
    
    Dim lDateDiff As Long
    
    Set fso = New FileSystemObject
    bFileAExists = fso.FileExists(szFile1)
    bFileBExists = fso.FileExists(szFile2)
   

    'Determine which file to use
    If Not bFileAExists And Not bFileBExists Then
        szLogFileName = szFile1 'Neither file exists, so  use file1
    ElseIf bFileAExists And Not bFileBExists Then
        szLogFileName = szFile1 'file1 exists and file2 does not, so use file 1
    ElseIf Not bFileAExists And bFileBExists Then
        szLogFileName = szFile2 'file2 exists and file1 does not, so use file 1
    Else    'Both files exist, so we will have to compare the files
        Set fFile1 = fso.GetFile(szFile1)
        dtFile1Time = fFile1.DateLastModified
        
        Set fFile2 = fso.GetFile(szFile2)
        dtfile2Time = fFile2.DateLastModified
        
        lDateDiff = DateDiff("s", dtfile2Time, dtFile1Time)
        
        If lDateDiff > 0 Then
            szLogFileName = szFile1 'File1 was logged to last, so use file1
        ElseIf lDateDiff < 0 Then
            szLogFileName = szFile2 'File2 was logged to last, so use file2
        Else
            If fFile1.Size < fFile2.Size Then   'Both dated the same (who knows why), use smaller one
                szLogFileName = szFile1
            Else
                szLogFileName = szFile2
            End If
        End If
            
    End If

    szGetFileName = szLogFileName

End Function
    
Public Property Get ActiveLogFile() As String
     If m_lLoggingOn <> False Then Exit Property
     
     ActiveLogFile = szLogFileName
End Property

Public Property Get InactiveLogFile() As String
    If m_lLoggingOn = False Then Exit Property
 
    If szLogFileName = szFile1 Then
        InactiveLogFile = szFile2
    Else
        InactiveLogFile = szFile1
    End If
End Property

Private Sub Class_Terminate()
     If Not tsLogStream Is Nothing Then
          Set tsLogStream = fLogFile.OpenAsTextStream(ForAppending)
          tsLogStream.WriteBlankLines (1)
          tsLogStream.WriteLine "*****************************************"
          tsLogStream.WriteLine Now & vbTab & "Stop Logging"
          tsLogStream.WriteLine "*****************************************"
          tsLogStream.WriteBlankLines (1)
          
          tsLogStream.Close
          Set fso = Nothing
          Set fLogFile = Nothing
          Set tsLogStream = Nothing
     
     End If
     
End Sub
