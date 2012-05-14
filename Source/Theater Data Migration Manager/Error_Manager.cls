VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Error_Manager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mbDisplayMessage As Boolean
Private mbLogMessage As Boolean
Private msLogFile As String
Private msErrorPath As String
Private mobjLogManager As Log_Manager
Private msErrorMess As String
Private mobjErrorLog As Log_Manager
Public Sub WriteLine_Error_To_Log(ErrorMessage As String)
On Error GoTo ErrHandler

    mobjLogManager.OpenLog msLogFile, False
    mobjLogManager.WriteLine ErrorMessage
    mobjLogManager.CloseLog

GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "Error_Manager", "Write_Error_To_Log", Err.Number, Err.Description
CleanUp:
End Sub

Public Sub Write_Error_To_Log(ErrorMessage As String)
On Error GoTo ErrHandler

    mobjLogManager.OpenLog msLogFile, False
    mobjLogManager.WriteText ErrorMessage
    mobjLogManager.CloseLog

GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "Error_Manager", "Write_Error_To_Log", Err.Number, Err.Description
CleanUp:
End Sub

Public Sub Err_Log(Module As String, Procedure, Error_Number As Long, Error_Message As String, Optional Raise_Error As Boolean = True, Optional Comments As String = "")

    Dim sError As String

    glExitCode = Exit_Code_Enum.General_Failure
    
    If Error_Number <> Error_Enum.Application_Error Then  '-2147211505 Then 'bObjectError + 9999 = -2147211505

        msErrorMess = ""
        
'Or Error_Number <> Error_Enum.Database_Connect_Error Or _
'        Error_Number <> Error_Enum.Export_Error Or Error_Number <> Error_Enum.Import_Error Or _
'        Error_Number <> Error_Enum.SQL_Error
        msErrorPath = ""

        sError = "Module: " & Module & vbNewLine & _
                   "Procedure: " & Procedure & vbNewLine & _
                   "Error Number: " & CStr(Error_Number) & vbNewLine & _
                   "Error Message: " & Error_Message
                   
        msErrorMess = sError
                   
        If mbDisplayMessage = True Then
            MsgBox sError, vbCritical, "Database Patch Utility"
        End If

        If mbLogMessage = True Then
            WriteLine_Error_To_Log sError
            mobjErrorLog.WriteField sError
        End If

    End If
    
    If Error_Number = Error_Enum.Application_Error Then

        If msErrorPath = "" Then
            msErrorPath = Module & "." & Procedure
        Else
            msErrorPath = Module & "." & Procedure & " --> " & msErrorPath
        End If

    End If
    
    '-2147211505
    If Raise_Error = False Then 'Error_Number = Application_Error Or
        
        If Not mobjErrorLog Is Nothing Then
            mobjErrorLog.WriteField "Error Path: " & msErrorPath
            mobjErrorLog.WriteField "Time: " & Now
        End If
        
        WriteLine_Error_To_Log "Error Path: " & msErrorPath
        WriteLine_Error_To_Log "Time: " & Now
        If Comments <> "" Then
            WriteLine_Error_To_Log "Comments: " & Comments
            If Not mobjErrorLog Is Nothing Then
                mobjErrorLog.WriteField "Comments: " & Comments
            End If
        End If
        WriteLine_Error_To_Log "---------------------------------------------------------------------------------------"
        If Not mobjErrorLog Is Nothing Then
             mobjErrorLog.WriteField "---------------------------------------------------------------------------------------"
        End If
    End If
    
    If Raise_Error Then
        Err.Raise Number:=Application_Error
    Else
        Err.Number = Application_Error
    End If

End Sub

Public Property Get Display_Message() As Boolean
    Display_Message = mbDisplayMessage
End Property

Public Property Let Display_Message(ByVal vNewValue As Boolean)
    mbDisplayMessage = vNewValue
End Property

Public Property Get Log_Message() As Boolean
    Log_Message = mbLogMessage
End Property

Public Property Let Log_Message(ByVal vNewValue As Boolean)
    mbLogMessage = vNewValue
End Property

Public Property Get Log_File() As String
    Log_File = msLogFile
End Property

Public Property Let Log_File(ByVal vNewValue As String)
    msLogFile = vNewValue
End Property

Private Sub Class_Initialize()
    mbDisplayMessage = True
    mbLogMessage = True
    Set mobjLogManager = New Log_Manager
End Sub

Public Property Get Error_Path() As String
    Error_Path = msErrorPath
End Property
Public Property Let Error_Path(ByVal vNewValue As String)
    If vNewValue = "" Then
        msErrorPath = ""
    Else
        msErrorPath = vNewValue & vbNewLine & msErrorPath
    End If
End Property


Private Sub Class_Terminate()
    Set mobjLogManager = Nothing
    Set mobjErrorLog = Nothing
End Sub

Public Property Get Error_Log() As Variant
    Set Error_Log = mobjErrorLog
End Property

Public Property Let Error_Log(ByVal vNewValue As Variant)
    Set mobjErrorLog = vNewValue
End Property
