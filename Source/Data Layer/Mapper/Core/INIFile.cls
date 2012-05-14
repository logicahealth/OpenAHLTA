VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "INIFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Used to read from and write to INI files.

Option Explicit

Private mFileName As String
Private mFullFileName As String

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpSectionName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpSectionName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Property Get FullFileName() As String
  FullFileName = mFullFileName
End Property

Public Sub Init(ByVal FileName As String)
Const FILE_EXTENTION As String = ".INI"
  
  If Right$(FileName, Len(FILE_EXTENTION)) <> FILE_EXTENTION Then FileName = FileName & FILE_EXTENTION
  
  mFileName = FileName
  mFullFileName = InitTempDir & FileName
  
End Sub

Public Function GetINIValue(ByRef Item As OshBinItem) As Boolean
Const MAX_VALUE_STRING_LEN As Integer = 2000

Dim NumBytesReturned As Long
Dim ReturnBuffer As String
    
  On Error GoTo ErrHandler

  ReturnBuffer = Space$(MAX_VALUE_STRING_LEN)

  NumBytesReturned = GetPrivateProfileString(Item.Section, Item.Key, vbNullString, ReturnBuffer, MAX_VALUE_STRING_LEN, mFullFileName)
    
  If NumBytesReturned = MAX_VALUE_STRING_LEN - 2 Then
    Screen.MousePointer = vbDefault
    Err.Raise 60060, "MapperCDR.INIFile", "Buffer overflow occurred in Config.mbGetFileValueString.  The buffer size: " & CStr(MAX_VALUE_STRING_LEN) & " bytes was not large enough to hold all of the returned data."
  End If

  If NumBytesReturned > 0 Then
    If NumBytesReturned <> MAX_VALUE_STRING_LEN - 1 Then
      Item.Value = Mid$(ReturnBuffer, 1, NumBytesReturned)
      GetINIValue = True
    Else
      Err.Raise 60061, "MapperCDR.INIFile", "The value string retrieved from the file: " & mFullFileName & _
                                         " exceeded the maximum length of " & CStr(MAX_VALUE_STRING_LEN) & " bytes."
    End If
  Else
    Item.Value = vbNullString    'In case someone is passing something in the Value parameter (By Ref)
  End If

  Exit Function

ErrHandler:
  Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Function SetINIValue(ByVal Section As String, ByVal Key As String, _
                            ByVal Value As String) As Boolean
Dim rc As Long

  rc = WritePrivateProfileString(Section, Key, Value, mFullFileName)
  
  If rc <> 0 Then
    '  VERY IMPORTANT: To ensure the modifications REALLY get written to disk, you need
    '                  to "flush" the file buffers.  If the first three parameters are
    '                  "vbNullString", the API call will "flush" the file buffers.
    rc = WritePrivateProfileString(vbNullString, vbNullString, vbNullString, mFullFileName)
    
    SetINIValue = True
  End If
                         
End Function

'All ini files are saved in user's temp directory
Private Function InitTempDir() As String
Dim TempDir As String

  TempDir = Environ$("temp")
    
  If Right$(TempDir, 1) <> "\" Then
    InitTempDir = TempDir & "\"
  End If
  
End Function

