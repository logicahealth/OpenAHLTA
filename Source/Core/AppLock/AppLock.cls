VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AppLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const SW_SHOWMINIMIZED = 2
Private Const ASFW_ANY = -1

Private Type RECT
   Left    As Long
   Top     As Long
   Right   As Long
   Bottom  As Long
End Type

Private Type POINTAPI
   x       As Long
   y       As Long
End Type

Private Type WINDOWPLACEMENT
   length            As Long
   flags             As Long
   showCmd           As Long
   ptMinPosition     As POINTAPI
   ptMaxPosition     As POINTAPI
   rcNormalPosition  As RECT
End Type

Private WithEvents mProxy  As AppLockProxy
Attribute mProxy.VB_VarHelpID = -1
Private mWindowsHandles As Collection

Private mExternalProcess As Collection

Private mLocked As Boolean

Private mProcessID As Long

Private mLogoPath As String

Public Event Restore()
Public Event CloseApp()
Public Event Validate(ByVal Password As String, ByRef IsValid As Boolean, ByRef Cancel As Boolean)

Private Declare Function GetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Private Declare Function SetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Private Declare Function AllowSetForegroundWindow Lib "user32.dll" (ByVal dwProcessId As Long) As Long

Private Sub Class_Terminate()
  Set mProxy = Nothing
  Set mWindowsHandles = Nothing
  Set mExternalProcess = Nothing
End Sub

Private Sub mProxy_CloseApp()
  mLocked = False
  RaiseEvent CloseApp
End Sub

Private Sub mProxy_Validate(ByVal Password As String, IsValid As Boolean)
Dim Cancel As Boolean

  On Error GoTo ErrHandler

  RaiseEvent Validate(Password, IsValid, Cancel)
  
  If IsValid Then
    mLocked = False
    
    ShowApp
    
    RaiseEvent Restore
    
  ElseIf Cancel Then
    TimeoutDialog.Abort = True
    
    MsgBox "You have cancelled re-authentication.  Exiting Application."
    RaiseEvent CloseApp
  End If
  
  Exit Sub
  
ErrHandler:
  MsgBox Err.Description & " (" & Err.Number & ", AppLock.mProxy_Validate)"
  
End Sub

Public Property Let AlternateLogoPath(ByVal NewValue As String)
  mLogoPath = NewValue
End Property

Public Property Let ProcessID(ByVal NewValue As Long)
  mProcessID = NewValue
End Property

Public Sub AddExternalProcess(ByVal WindowTitle As String, ByVal Handle As Long)
Dim ProcessID As Long

  On Error Resume Next
  
  If mExternalProcess Is Nothing Then Set mExternalProcess = New Collection

  If Handle = 0 Then
    Handle = FindWindowLike(0, WindowTitle, "*")
  End If
  
  If Handle > 0 Then
    ProcessID = ProcIDFromWnd(Handle)
  End If
  
  If ProcessID > 0 Then
    mExternalProcess.Add ProcessID
  End If
  
End Sub

Private Sub AllowSetForeGround()
Dim rc As Long

  On Error Resume Next
  'AllowSetForegroundWindow mProcessID
  rc = AllowSetForegroundWindow(-1)     'scr 34363  Phil Crowder 4/21/03

End Sub

Public Function IsLocked() As Boolean
  IsLocked = mLocked
End Function

Private Sub HideApp(ByVal ProcessID As Long)
Dim ParentHandle As Long
  
  If mWindowsHandles Is Nothing Then Set mWindowsHandles = New Collection

  'Retrieve all the parent (top level) window handles owned by this process
  ParentHandle = ProcessToParentWnd(ProcessID)

  Do While ParentHandle

    ShowWindow ParentHandle, SW_HIDE

    mWindowsHandles.Add ParentHandle

    ParentHandle = ProcessToParentWnd(ProcessID)
  Loop

  'Reset internal statics in function
  ProcessToParentWnd 0
  
End Sub

'SCR 46301, 44332, 43676, 43647 Phil Crowder 12/12/03 Added DoEvents, appears to be a timing issue.
Public Sub LockApp(ByVal UserID As String, ByVal DialogTitle As String, ByVal Instructions As String)
Dim i As Long
Dim rc As Long

  On Error GoTo ErrHandler
    
  DoEvents
  
  HideApp mProcessID
  
  DoEvents
  
  If Not (mExternalProcess Is Nothing) Then
    For i = 1 To mExternalProcess.Count
      HideApp mExternalProcess.Item(i)
      DoEvents
    Next
  End If
  
  DoEvents
  AllowSetForeGround
  
  If mProxy Is Nothing Then Set mProxy = New AppLockProxy
    
  DoEvents
  ShowTimeoutDialog UserID, DialogTitle, Instructions
  
  DoEvents
  mLocked = True
  
  Exit Sub
  
ErrHandler:
  MsgBox Err.Description & " (" & Err.Number & ", AppLock.LockApp)"
  
End Sub

Public Sub Minimize()
  mProxy.Minimize
End Sub

Public Sub Restore()
  mProxy.Restore
End Sub

Private Function ProcIDFromWnd(ByVal hWnd As Long) As Long
Dim idProc As Long
    
  GetWindowThreadProcessId hWnd, idProc
  ProcIDFromWnd = idProc
  
End Function

Private Sub ShowApp()
Dim i As Long
Dim WStyle As Long
      
  For i = mWindowsHandles.Count To 1 Step -1
    WStyle = GetWindowLong(mWindowsHandles.Item(i), GWL_STYLE)
    
    If WStyle And WS_MAXIMIZE Then
      ShowWindow mWindowsHandles.Item(i), SW_SHOWMAXIMIZED
    Else
      ShowWindow mWindowsHandles.Item(i), SW_RESTORE
    End If
  Next

  Set mWindowsHandles = Nothing

End Sub

'Need to use API to show dialog because the dialog does not gain focus correctly
'the 1st time the dialog is show using the VB Show method.
Private Sub ShowTimeoutDialog(ByVal UserID As String, ByVal DialogTitle As String, ByVal Instructions As String)
Dim WndToActivate As Long
Dim currWinP As WINDOWPLACEMENT
Dim TDialog As TimeoutDialog
    
  Set TDialog = New TimeoutDialog

  TDialog.Init UserID, DialogTitle, Instructions, mProxy, mLogoPath

  WndToActivate = TDialog.hWnd

  With currWinP
     .length = Len(currWinP)
     GetWindowPlacement WndToActivate, currWinP

     .length = Len(currWinP)
     .flags = 0&
     .showCmd = SW_SHOWMINIMIZED
  End With

  SetWindowPlacement WndToActivate, currWinP

  Set TDialog = Nothing
  
End Sub
