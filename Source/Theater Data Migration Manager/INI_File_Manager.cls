VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "INI_File_Manager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Application Key Value
'[Application]
'Key=Value

Private Declare Function WritePrivateProfileString _
    Lib "kernel32" Alias "WritePrivateProfileStringA" _
        (ByVal lpApplicationname As String, ByVal lpKeyName As Any, _
            ByVal lsString As Any, ByVal lplFilename As String) As Long
Private Declare Function GetPrivateProfileInt Lib _
    "kernel32" Alias "GetPriviteProfileIntA" (ByVal lpApplicationname As String, _
        ByVal lpKeyName As String, ByVal nDefault As Long, _
            ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias _
    "GetPrivateProfileStringA" (ByVal lpApplicationname As String, _
        ByVal lpKeyName As String, ByVal lpDefault As String, _
            ByVal lpReturnedString As String, ByVal nSize As Long, _
                ByVal lpFileName As String) As Long


Private msFileName As String
Private msApplicationName As String ' [Application]

Public Function Get_Value(ByVal Key_Name As String, Optional ByVal Default_Return_Value As String = "") As String
On Error GoTo ErrHandler

    Dim strVal As String * 1000
    Dim lret As Long

    lret = GetPrivateProfileString(msApplicationName, Key_Name, Default_Return_Value, strVal, Len(strVal), msFileName)
    If lret <> 0 Then
        Get_Value = ConvertString(strVal)
    Else
        'MsgBox "Bad return value!", vbCritical
        'Err.Raise Application_Error, "", "PrivateProfileString return error."
    End If

GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "INI_File_Manager", "Get_Value", Err.Number, Err.Description
CleanUp:
End Function



Public Function Set_Value(ByVal Key_Name As String, ByVal Value As String) As Boolean
On Error GoTo ErrHandler

    Dim ret As Long
    
    ret = WritePrivateProfileString(msApplicationName, Key_Name, Value, msFileName)
    If ret = 0 Then
        Set_Value = False
        Exit Function
    End If
    Set_Value = True

GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "INI_File_Manager", "Get_Value", Err.Number, Err.Description
CleanUp:
End Function



Private Function ConvertString(str As String) As String
On Error GoTo ErrHandler

Dim i As Integer
Dim intAsc As Integer

    str = Trim(str)
    
    For i = 1 To Len(str)
        intAsc = Asc(Mid(str, i, 1))
        If intAsc > 31 And intAsc < 127 Then
            ConvertString = ConvertString & Chr(intAsc)
        End If
    Next
    
GoTo CleanUp
ErrHandler:
    gobjErrManager.Err_Log "INI_File_Manager", "ConvertString", Err.Number, Err.Description
CleanUp:
End Function

Public Property Get Application_Name() As String
    Application_Name = msApplicationName
End Property

Public Property Let Application_Name(ByVal vNewValue As String)
    msApplicationName = vNewValue
End Property

Public Property Get File_Name() As String
    File_Name = msFileName
End Property

Public Property Let File_Name(ByVal vNewValue As String)
    msFileName = vNewValue
End Property
