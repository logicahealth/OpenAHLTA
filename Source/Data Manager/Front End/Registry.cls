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

Private m_hKey As Long
'local variable(s) to hold property value(s)
Private mvarRootKey As Long 'local copy

Public Property Let RootKey(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RootKey = 5
    mvarRootKey = vData
End Property

Public Property Get RootKey() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RootKey
    RootKey = mvarRootKey
End Property

Public Sub OpenKey(strSubKey As String, bCanCreate As Boolean)
Dim lRet As Long
Dim dwDisposition As Long

lRet = RegOpenKeyEx( _
            RootKey, _
            strSubKey, _
            0, _
            KEY_READ Or KEY_WRITE, _
            m_hKey& _
            )
                        
End Sub

Public Sub CloseKey()
RegCloseKey (m_hKey)
End Sub

Private Sub Class_Initialize()
m_hKey = 0
End Sub

Private Sub Class_Terminate()
If m_hKey <> 0 Then
    CloseKey
    m_hKey = 0
End If
End Sub

Public Sub ReadString(strName As String, Value As String, Optional defValue As String)
Dim lType As Long
Dim nSize As Long
Dim lRet As Long
Dim str As String * 256

Value = defValue

nSize = 256
lRet = RegQueryStrValueEx( _
    m_hKey, _
    strName, _
    0&, _
    lType, _
    str, _
    nSize)
    
If lRet = 0 Then
    'GRIZ 07.21.98 fixed blow up if String was blank
    If nSize <= 0 Then
        Value = ""
    Else
        Value = Left$(str, nSize - 1)
    End If
Else
    WriteString strName, Value
End If
       
End Sub

Public Sub ReadInteger(strName As String, Value As Long, Optional defValue As Long)
Dim lType As Long
Dim nSize As Long
Dim lRet As Long
Dim iVal As Long

Value = defValue

nSize = 8
lRet = RegQueryValueEx( _
    m_hKey, _
    strName, _
    0&, _
    lType, _
    iVal, _
    nSize)
    
If lRet = 0 Then
    If lType = REG_DWORD Then Value = iVal
Else
    WriteInteger strName, Value
End If
    
End Sub

Public Sub ReadBool(strName As String, Value As Boolean, Optional defValue As Boolean)
Dim lType As Long
Dim nSize As Long
Dim lRet As Long
Dim iVal As Long
Dim iDefault As Long

If defValue = True Then
    iDefault = 1
Else
    iDefault = 0
End If
    
nSize = 4
ReadInteger strName, iVal, iDefault
   
If iVal = 0 Then
    Value = False
Else
    Value = True
End If
    
End Sub

Public Sub WriteInteger(strName As String, Value As Long)

lRet = RegSetValueEx( _
    m_hKey, _
    strName, _
    0&, _
    REG_DWORD, _
    Value, _
    4)
           
End Sub


Public Sub WriteBool(strName As String, Value As Boolean)
Dim iVal As Long

If Value = True Then
    iVal = 1
Else
    iVal = 0
End If

WriteInteger strName, iVal
    
End Sub

Public Sub WriteString(strName As String, Value As String)

lRet = RegSetStrValueEx( _
    m_hKey, _
    strName, _
    0&, _
    REG_SZ, _
    Value, _
    Len(Value))
           
End Sub

