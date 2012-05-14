VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AnPLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This support class is used to lock A/P during lengthy operations.  It
'makes sure that once the module is locked that it stays locked until the
'operation is complete.

Option Explicit

Private mMethod As String      'Used to identify where the lock occurred
Private mLocked As Boolean     'Current stat of lock

Public Event Disable()
Public Event Enable()

Public Property Get Locked() As Boolean
  Locked = mLocked
End Property

Public Property Get Method() As String
  Method = mMethod
End Property

Public Sub StartLock(ByVal MethodName As String)

  On Error GoTo ErrHandler
  
  If Len(MethodName) = 0 Then Exit Sub
  
  If Len(mMethod) = 0 Or mMethod = MethodName Then
    mLocked = True
    
    Screen.MousePointer = vbHourglass
        
    mMethod = MethodName
    
    RaiseEvent Disable
  End If
  
  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnPLock.StartLock", "A/P", vbExclamation
  
End Sub

Public Sub StopLock(ByVal MethodName As String)
  
  On Error Resume Next
  
  If mMethod = MethodName Then
    mMethod = vbNullString
    
    RaiseEvent Enable
        
    Screen.MousePointer = vbDefault
    
    mLocked = False
  End If
  
End Sub

'Allows code to unlock the module even if it was not the source of the lock.
Public Sub ForceUnlock()
  StopLock mMethod
End Sub

