VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RegistryObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'========================================================================
' The RegistryObject class exposes methods that utilize RegObj.dll,
' a freeware utility from the VB Owners Area website, to read, write
' and delete registry settings.
'
' NOTE: RegObj.dll must be registered and you must add a reference
' to "Registration Manipulation Classes" in your project in order
' for these methods to work.
'
' The entire registry can be manipulated through the 3 public methods
' in this class. To retrieve error messages, read the .LastRegError
' property immediately after an error code is returned.
'
' .RegRead returns the read value or the passed default if an error
' occurs (Empty is returned if there is no default and an error occurs).
' For .RegWrite and .RegDelete, a return value of 0 (ERROR_SUCCESS) is
' successful; all other return values are errors.
'========================================================================
' Examples:
'
'  '*** Create a registry object
'  Dim ro As New RegistryObject
'
'  ' Write a value:
'  ro.RegWrite HKEY_CURRENT_USER, "Sub\Key\Registry\Path", "ValueName", "write value", rvString
'
'  ' To report write errors, do it this way:
'  If ro.RegWrite(HKEY_CURRENT_USER, "Sub\Key\Registry\Path", "ValueName", "write value", rvString) <> ERROR_SUCCESS Then
'    MsgBox ro.LastRegError, vbExclamation, "RegWrite Error"
'  End If
'
'  '*** Read a value:
'  Dim Value As Variant
'  Value = ro.RegRead(HKEY_CURRENT_USER, "Sub\Key\Registry\Path", "ValueName", "default value")
'
'  ' To report read errors, do it this way:
'  ' (Value must be a Variant in order to hold an Empty value)
'  Value = ro.RegRead(HKEY_CURRENT_USER, "Sub\Key\Registry\Path", "ValueName")
'  If Value = Empty Then
'    MsgBox ro.LastRegError, vbExclamation, "RegRead Error"
'  Else
'    MsgBox Value, vbInformation, "RegRead Value"
'  End If
'
'  '*** Delete a registry key:
'  ro.RegDelete HKEY_CURRENT_USER, "Sub\Key\Registry\Path", "ValueName"
'
'  ' To report delete errors, do it this way:
'  If ro.RegDelete(HKEY_CURRENT_USER, "Sub\Key\Registry\Path", "ValueName") <> ERROR_SUCCESS Then
'    MsgBox ro.LastRegError, vbExclamation, "RegDelete Error"
'  End If
'
'  '*** Delete a registry key with all its sub keys:
'  ' (to remove all sub keys, don't pass a value name)
'  ro.RegDelete HKEY_CURRENT_USER, "Sub"
'========================================================================



Public Enum HKeys
  HKEY_CLASSES_ROOT = RegObj.HKEY_CLASSES_ROOT
  HKEY_CURRENT_USER = RegObj.HKEY_CURRENT_USER
  HKEY_LOCAL_MACHINE = RegObj.HKEY_LOCAL_MACHINE
  HKEY_USERS = RegObj.HKEY_USERS
  HKEY_PERFORMANCE_DATA = RegObj.HKEY_PERFORMANCE_DATA
End Enum

Public Enum RegReturnCodes
  ERROR_SUCCESS = 0
  ' Any other return value is an error
End Enum

Private m_objRegKey As RegObj.RegKey
Private m_strLastError  As String


' Public access for m_strLastError variable
' -----------------------------------------
Public Property Get LastRegError() As String
  LastRegError = m_strLastError
End Property


' Set private m_strLastError variable
' -----------------------------------
Private Sub SetLastError(Optional ByVal Reset As Boolean = False)
  If Reset Then
    m_strLastError = "No error"
  Else
    If InStr(Err.Description, "name is not in use") Then
      ' Create a more meaningful message for this error
      m_strLastError = "Sub key or named value not found"
    Else
      m_strLastError = Err.Description
    End If
  End If
End Sub


' Read a registry key value or return a default
' ---------------------------------------------
' The optional ReturnedValueType returns the ValueName type;
' it does not set it. You can set or change the value type
' using .RegWrite.
Public Function RegRead(ByVal hKey As HKeys, _
                        ByVal SubKey As String, _
                        ByVal ValueName As String, _
                        Optional ByVal Default As Variant, _
                        Optional ReturnedValueType As Variant) _
                        As Variant
                        
  On Error GoTo Handler
  
  SetLastError True    'Reset error message
  
  ' Set a reference to the root key
  Set m_objRegKey = RegObj.RegKeyFromHKey(hKey)
  
  ' Set a reference to the sub key
  Set m_objRegKey = m_objRegKey.ParseKeyName(SubKey)
  
  RegRead = m_objRegKey.Values.Item(ValueName).Value
  ReturnedValueType = m_objRegKey.Values.Item(ValueName).Type
Exit Function

Handler:
  If IsMissing(Default) Then
    SetLastError
    RegRead = Empty   'If no default, return Empty
  Else
    RegRead = Default
  End If
End Function


' Write a new registry key or modify an existing value
' ----------------------------------------------------
' If the optional ValueName and corresponding Value are not passed,
' only the the SubKey will be created. If the optional ValueType
' is not passed, 'rvString' is the default.
Public Function RegWrite(ByVal hKey As HKeys, _
                         ByVal SubKey As String, _
                         Optional ByVal ValueName As Variant, _
                         Optional ByVal Value As Variant, _
                         Optional ByVal ValueType As RegObj.RegValueType = rvString) _
                         As RegReturnCodes
                         
  On Error GoTo Handler
  
  SetLastError True    'Reset error message
  
  ' Set a reference to the root key
  Set m_objRegKey = RegObj.RegKeyFromHKey(hKey)
  
  ' Add the sub key
  ' (ignore error here in case the sub key already exists)
  On Error Resume Next
  m_objRegKey.SubKeys.Add SubKey

  ' See if we can set a reference to the sub key
  On Error GoTo Handler
  Set m_objRegKey = m_objRegKey.ParseKeyName(SubKey)
    
  If Not IsMissing(ValueName) And Not IsMissing(Value) Then
    ' Add the value
    ' (ignore error here in case the value name already exists)
    On Error Resume Next
    m_objRegKey.Values.Add ValueName, Value, ValueType
    
    If Err.Number Then
      ' If an error occurred, the value name may
      ' have already existed -- try to change it
      On Error GoTo Handler
      m_objRegKey.Values.Item(ValueName).ChangeType ValueType, Value
    End If
  End If
  
  RegWrite = ERROR_SUCCESS
Exit Function

Handler:
  SetLastError
  RegWrite = Err.Number
End Function


' Delete a sub key or value name
' ------------------------------
' If the optional ValueName is not passed, the SubKey and all its
' values will be removed.  If ValueName is passed, only the ValueName
' will be removed.
Public Function RegDelete(ByVal hKey As HKeys, _
                          ByVal SubKey As String, _
                          Optional ByVal ValueName As Variant) _
                          As RegReturnCodes
                          
  On Error GoTo Handler
  
  SetLastError True    'Reset error message
  
  ' Set a reference to the root key
  Set m_objRegKey = RegObj.RegKeyFromHKey(hKey)
  
  ' Set a reference to the sub key
  Set m_objRegKey = m_objRegKey.ParseKeyName(SubKey)
    
  If IsMissing(ValueName) Then
    ' Remove the sub key and all its sub values
    m_objRegKey.Parent.SubKeys.Remove m_objRegKey.Name
  Else
    ' Remove the value name only
    m_objRegKey.Values.Remove m_objRegKey.Values.Item(ValueName).Name
  End If
  
  RegDelete = ERROR_SUCCESS
Exit Function

Handler:
  SetLastError
  RegDelete = Err.Number
End Function


Private Sub Class_Terminate()
  ' Release reference to private reg object
  Set m_objRegKey = Nothing
End Sub


