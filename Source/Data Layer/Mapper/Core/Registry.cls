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
Attribute VB_Name = "Registry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private Const REG_SZ As Long = 1
Private Const REG_DWORD As Long = 4

Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003

Private Const ERROR_NONE = 0
Private Const ERROR_BADDB = 1
Private Const ERROR_BADKEY = 2
Private Const ERROR_CANTOPEN = 3
Private Const ERROR_CANTREAD = 4
Private Const ERROR_CANTWRITE = 5
Private Const ERROR_OUTOFMEMORY = 6
Private Const ERROR_ARENA_TRASHED = 7
Private Const ERROR_ACCESS_DENIED = 8
Private Const ERROR_INVALID_PARAMETERS = 87
Private Const ERROR_NO_MORE_ITEMS = 259

Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_ALL_ACCESS = &H3F

Private Const REG_OPTION_NON_VOLATILE = 0

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long

Public Function QueryValue(ByVal PrimaryKey As String, ByVal SubKey As String, ByVal ValueKey As String) As String
Dim Key As String
Dim Val As String
Dim KeyHandle As Long
Dim Success As Long

  On Error GoTo ErrHandler

  Key = "SOFTWARE\CH2CW" & "\" & PrimaryKey & "\" & SubKey
  
  Success = RegOpenKeyEx(HKEY_LOCAL_MACHINE, Key, 0, KEY_QUERY_VALUE, KeyHandle)
  Success = QueryValueEx(KeyHandle, ValueKey, Val)

  RegCloseKey KeyHandle
  
  QueryValue = CStr(Val)

  Exit Function
  
ErrHandler:
  Err.Raise Err.Number, "Registry.GetValue", "Err.Description"

End Function

Private Function QueryValueEx(ByVal KeyHandle As Long, ByVal ValueKey As String, ByRef Value As String) As Long
Dim cch As Long
Dim rc As Long
Dim ValueType As Long
Dim Val As String

  On Error GoTo ErrHandler

  rc = RegQueryValueExNULL(KeyHandle, ValueKey, 0&, ValueType, 0&, cch)
       
  If rc <> ERROR_NONE Then Error 5

  Val = String(cch, 0)

  rc = RegQueryValueExString(KeyHandle, ValueKey, 0&, REG_SZ, Val, cch)
   
  If rc = ERROR_NONE Then
      Value = Left$(Val, cch - 1)
  Else
      Value = vbNullString
  End If

ErrHandler:
  QueryValueEx = rc
       
End Function

Public Function SetValue(ByVal PrimaryKey As String, ByVal SubKey As String, ByVal ValueKey As String, ByVal Value As Variant) As Boolean
Dim Key As String
Dim KeyHandle As Long
Dim Success As Long
    
  On Error GoTo ErrHandler
  
  Key = "SOFTWARE\C2CW" & "\" & PrimaryKey & "\" & SubKey
  
  Success = RegCreateKeyEx(HKEY_LOCAL_MACHINE, Key, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, KeyHandle, Success)
  
  Success = SetValueEx(KeyHandle, ValueKey, CStr(Value))
  
  RegCloseKey KeyHandle
  
  SetValue = True
  
  Exit Function
  
ErrHandler:
  Err.Raise Err.Number, "Registry.SetValue", "Err.Description"

End Function

Private Function SetValueEx(ByVal KeyHandle As Long, ValueKey As String, Value As String) As Long
Dim Val As String
  
  Val = Value & Chr$(0)
  SetValueEx = RegSetValueExString(KeyHandle, ValueKey, 0&, REG_SZ, Val, Len(Val))
  
End Function

