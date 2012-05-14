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
Attribute VB_Name = "CRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Class Name:  CRegistry
'
' Description: The CRegistry class provides access to the Windows
'              Registry data. It encapsulates all the requisite API
'              calls to make it easier for programmers to get what they
'              need and not worry about how it happens.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' API Declarations
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Any, hKeyHandle As Long, lpdwDisposition As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, hKeyHandle As Long) As Long
Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Constant Declarations
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const ERROR_SUCCESS = 0&

Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_CONFIG = &H80000005
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_DYN_DATA = &H80000006
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_PERFORMANCE_DATA = &H80000004
Private Const HKEY_USERS = &H80000003

Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const REG_OPTION_NON_VOLATILE = 0
Private Const REG_SZ = 1
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const SYNCHRONIZE = &H100000
Private Const REG_CREATED_NEW_KEY = &H1
Private Const REG_OPENED_EXISTING_KEY = &H2

Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or _
   KEY_QUERY_VALUE Or _
   KEY_SET_VALUE Or _
   KEY_CREATE_SUB_KEY Or _
   KEY_ENUMERATE_SUB_KEYS Or _
   KEY_NOTIFY Or _
   KEY_CREATE_LINK) And (Not SYNCHRONIZE))


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Type Declarations
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Private Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Private m_blnSelected As Boolean
Private m_hCurrentKey As Long
Private m_lngCurrentSection As RegistrySection
Private m_strCurrentKey As String

Private m_strKeyClass As String
Private m_lngSubKeyCount As Long
Private m_lngMaxSubKeyLen As Long
Private m_lngMaxClassLen As Long
Private m_lngNumOfValues As Long
Private m_lngMaxValueNameLen As Long
Private m_lngMaxValueLen As Long
Private m_lngSecurityDescriptor As Long
Private m_typLastWriteTime As FILETIME

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' This public enum allows calling code to pick one of these values
' using the IntelliSense technology. It also saves a ton of code since
' we don't have to match up a constant/boolean flag with the predefined
' constants.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Enum RegistrySection
   ClassesRoot = HKEY_CLASSES_ROOT
   CurrentConfig = HKEY_CURRENT_CONFIG
   CurrentUserInfo = HKEY_CURRENT_USER
   DynamicData = HKEY_DYN_DATA
   LocalMachine = HKEY_LOCAL_MACHINE
   PerformaceData = HKEY_PERFORMANCE_DATA
   Users = HKEY_USERS
End Enum

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Public Sub OpenKey
'
' This subroutine allows the caller to select a key in a
' particular hive. This key then stays open until 1) CloseKey
' is called, 2) OpenKey is called again, or 3) the object is
' terminated.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub OpenKey(Section As RegistrySection, strKey As String, Optional blnCreateNew As Boolean = False)
   Dim lngReturn As Long
   Dim lngAction As Long
   
   If m_blnSelected = True Then
      CloseKey
   End If
   
   '
   ' Remove any leading backslashes, since they cause the function to fail
   '
   If Left(strKey, 1) = "\" Then
      strKey = Mid(strKey, 2)
   End If
   
   '
   ' Allow forward slashes, as they are a logical way to separate the keys
   '
   strKey = Replace(strKey, "/", "\")
   
   lngReturn = RegOpenKeyEx(Section, strKey, 0, KEY_ALL_ACCESS, m_hCurrentKey)
   
   If lngReturn = ERROR_SUCCESS Then
      m_blnSelected = True
      m_strCurrentKey = strKey
      m_lngCurrentSection = Section
   Else
      If blnCreateNew Then
         lngReturn = RegCreateKeyEx(Section, strKey, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal (0&), m_hCurrentKey, lngAction)
         If lngReturn = ERROR_SUCCESS Then
            m_blnSelected = True
            m_strCurrentKey = strKey
            m_lngCurrentSection = Section
         Else
            Err.Raise vbObjectError + 2, "CRegistry:OpenKey", "Failed to create key " & strKey & "."
         End If
      Else
         Err.Raise vbObjectError + 2, "CRegistry:OpenKey", "Failed to open key " & strKey & "."
         Exit Sub
      End If
   End If
   
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Public Sub DeleteKey
'
' This subroutine deletes the current key, if it has no subkeys.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DeleteKey()
   Dim lngReturn As Long
   Dim strKey As String
   
   strKey = m_strCurrentKey ' Mid(m_strCurrentKey, InStrRev(m_strCurrentKey, "\") + 1)
   CloseKey
   lngReturn = RegDeleteKey(CLng(m_lngCurrentSection), strKey)
   If lngReturn <> ERROR_SUCCESS Then
      Err.Raise vbObjectError + 2, "CRegistry:DeleteKey", "Failed to delete key " & strKey & "."
      Exit Sub
   End If
   
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Public Sub CloseKey
'
' This subroutine closes the current key, if necessary.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CloseKey()
   If m_hCurrentKey <> 0 Then
      RegCloseKey m_hCurrentKey
      m_blnSelected = False
      m_strCurrentKey = ""
   End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Public Function ReadValue
'
' This function returns a value from the selected key and
' value specified. If no value is specified, the routine/API
' call will return the default value in the selected key.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ReadValue(Optional strValueName As String = "") As String
   
   Dim strBuffer As String
   Dim lngLength As Long
   Dim lngReturn As Long
   
   If Not m_blnSelected Then
      Err.Raise vbObjectError + 2, "CRegistry:ReadValue", "No key is currently open."
      Exit Function
   End If
   
   strBuffer = Space(255)
   lngLength = 255
   lngReturn = RegQueryValueExString(m_hCurrentKey, strValueName, 0, 0, strBuffer, lngLength)
   If lngReturn = ERROR_SUCCESS Then
      ReadValue = Left(strBuffer, lngLength - 1)
   Else
      Err.Raise vbObjectError + 2, "CRegistry:ReadValue", "Failed to read value " & strValueName & " from Registry."
      Exit Function
   End If
   
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Public Sub WriteValue
'
' This function saves a value to the selected key and
' value specified.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub WriteValue(strValueName As String, strNewValue As String)
   
   Dim lngReturn As Long
   
   If Not m_blnSelected Then
      Err.Raise vbObjectError + 2, "CRegistry:WriteValue", "No key is currently open."
      Exit Sub
   End If
   
   lngReturn = RegSetValueEx(m_hCurrentKey, strValueName, 0&, REG_SZ, ByVal (strNewValue), Len(strNewValue))
   If lngReturn <> ERROR_SUCCESS Then
      Err.Raise vbObjectError + 2, "CRegistry:WriteValue", "Failed to write new value to " & strValueName & " from Registry."
      Exit Sub
   End If
   
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Public Sub DeleteValue
'
' This function deletes a value from the selected key.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DeleteValue(strValueName As String)

   Dim lngReturn As Long
   
   If Not m_blnSelected Then
      Err.Raise vbObjectError + 2, "CRegistry:DeleteValue", "No key is currently open."
      Exit Sub
   End If
   
   lngReturn = RegDeleteValue(m_hCurrentKey, strValueName)
   If lngReturn <> ERROR_SUCCESS Then
      Err.Raise vbObjectError + 2, "CRegistry:DeleteValue", "Failed to delete value " & strValueName & " from Registry."
      Exit Sub
   End If

End Sub

Private Sub Class_Terminate()
   CloseKey
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Private Sub RefreshKeyInfo
'
' This subroutine calls the RegQueryInfoKey to retrieve information about
' the current key. This information is then read through this object's
' properties. This routine is called whenever a property is accessed.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RefreshKeyInfo()

   Dim lngReturn As Long
   Dim lngKeyClassSize As Long
   Dim strLongBinary As String
   
   lngKeyClassSize = 255
   m_strKeyClass = Space(lngKeyClassSize)
  
   lngReturn = RegQueryInfoKey(m_hCurrentKey, m_strKeyClass, lngKeyClassSize, 0&, _
      m_lngSubKeyCount, m_lngMaxSubKeyLen, m_lngMaxClassLen, _
      m_lngNumOfValues, m_lngMaxValueNameLen, m_lngMaxValueLen, _
      m_lngSecurityDescriptor, m_typLastWriteTime)
   
   If lngReturn <> ERROR_SUCCESS Then
      Err.Raise vbObjectError + 2, "CRegistry:RefreshKeyInfo", "Failed to retrieve key information."
      Exit Sub
   End If
   
   If lngKeyClassSize > 0 Then
      m_strKeyClass = Left(m_strKeyClass, lngKeyClassSize - 1)
   End If
   
  
   
End Sub

Public Property Get CurrentKey() As String
   CurrentKey = m_strCurrentKey
End Property

Public Property Get KeyClass() As String
   RefreshKeyInfo
   KeyClass = m_strKeyClass
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Public Property LastWriteTime
'
' This property returns the last time that data was written to this
' key. It has to do some conversions to change the number provided
' by the API into a normal time value that is legible.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get LastWriteTime() As Date
   Dim typSystemTime As SYSTEMTIME
   Dim typLocalTime As FILETIME
   Dim lngReturn As Long
   
   On Error Resume Next ' This is in case something bad
                        ' is passed to the CDate function
   
   RefreshKeyInfo
   
   '
   ' Convert number to GMT time format
   '
   lngReturn = FileTimeToLocalFileTime(m_typLastWriteTime, typLocalTime)
   
   '
   ' Convert GMT time to local time
   '
   lngReturn = FileTimeToSystemTime(typLocalTime, typSystemTime)
   
   '
   ' Convert time structure to VB date format
   '
   With typSystemTime
      LastWriteTime = CDate(.wMonth & "/" & .wDay & "/" & .wYear _
         & " " _
         & .wHour & ":" & .wMinute & ":" & .wSecond)
   End With
End Property

Public Property Get NumOfSubKeys() As Long
   RefreshKeyInfo
   NumOfSubKeys = m_lngSubKeyCount
End Property

Public Property Get MaxSubkeyNameLength() As Long
   RefreshKeyInfo
   MaxSubkeyNameLength = m_lngMaxSubKeyLen
End Property

Public Property Get MaxClassNameLength() As Long
   RefreshKeyInfo
   MaxClassNameLength = m_lngMaxClassLen
End Property

Public Property Get NumOfValues() As Long
   RefreshKeyInfo
   NumOfValues = m_lngNumOfValues
End Property

Public Property Get MaxValueNameLength() As Long
   RefreshKeyInfo
   MaxValueNameLength = m_lngMaxValueNameLen
End Property

Public Property Get MaxValueLength() As Long
   RefreshKeyInfo
   MaxValueLength = m_lngMaxValueLen
End Property

Public Property Get SecurityDescriptor() As Long
   RefreshKeyInfo
   SecurityDescriptor = m_lngSecurityDescriptor
End Property

