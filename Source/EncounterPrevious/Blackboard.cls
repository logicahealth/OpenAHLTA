VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Blackboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mComm               As MMMHISComm.Comm
Private mConfig             As MMMHISConfiguration.Config
Private mLogon              As MMMHISLogon.Logon
Private mUser               As ICHCSII_User
Private moShared            As CWShared
Private mDAS                As ICHCSII_DAS
Private moDatabase          As CHCSII_CONN.Conn
Private mPatient            As MMMHISPatient.Patient
Private mAnPParent          As Object
Private mRefreshRequired            As Boolean

Private Sub Class_Terminate()
 
    Set mComm = Nothing
    Set mConfig = Nothing
    Set mLogon = Nothing
    Set mUser = Nothing
    Set moShared = Nothing
    Set mDAS = Nothing
    Set mAnPParent = Nothing
End Sub

Public Property Get AnPParent() As Object
  Set AnPParent = mAnPParent
End Property

Public Property Set AnPParent(ByRef NewValue As Object)
  Set mAnPParent = NewValue
End Property


Public Property Get Comm() As MMMHISComm.Comm
  Set Comm = mComm
End Property

Public Property Set Comm(ByRef NewValue As Object)
  Set mComm = NewValue
End Property

Public Property Get Config() As MMMHISConfiguration.Config
  Set Config = mConfig
End Property

Public Property Set Config(ByRef NewValue As Object)
  Set mConfig = NewValue
End Property

Public Property Get Logon() As MMMHISLogon.Logon
  Set Logon = mLogon
End Property

Public Property Set Logon(ByRef NewValue As Object)
  Set mLogon = NewValue
End Property

Public Property Get User() As ICHCSII_User
  Set User = mUser
End Property

Public Property Set User(ByRef NewValue As ICHCSII_User)
  Set mUser = NewValue
End Property

Public Property Get SharedObj() As CWShared
    If moShared Is Nothing Then Set moShared = New CWShared
    Set SharedObj = moShared
End Property

Public Sub ClearBlackBoard()
    Set moShared = Nothing
End Sub

Public Property Get I_DAS() As ICHCSII_DAS
On Error GoTo ErrHandler
  
    If mDAS Is Nothing Then
        Set mDAS = DataConn.CHCSII_DAS
    End If
    Set I_DAS = mDAS
    
    Exit Property
  
ErrHandler:
    moShared.ShowVBError Err.Number, Err.Description, "BlackBoard.", "OrderAdmin", vbExclamation

End Property

Public Property Get DataConn() As CHCSII_CONN.Conn
    If moDatabase Is Nothing Then
        Set moDatabase = New CHCSII_CONN.Conn
    End If
    Set DataConn = moDatabase
End Property


Public Property Get Patient() As MMMHISPatient.Patient
  Set Patient = mPatient
End Property

Public Property Set Patient(ByRef NewValue As MMMHISPatient.Patient)
  Set mPatient = NewValue
End Property



Public Property Get RefreshRequired() As Boolean
    RefreshRequired = mRefreshRequired
End Property

Public Property Let RefreshRequired(ByVal vNewValue As Boolean)
    mRefreshRequired = vNewValue
End Property
