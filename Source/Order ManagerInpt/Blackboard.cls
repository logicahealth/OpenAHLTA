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

Private mComm   As MMMHISComm.Comm
Private mConfig As MMMHISConfiguration.Config
Private mLogon  As MMMHISLogon.Logon
Private mPatient As MMMHISPatient.Patient
Private mUser   As ICHCSII_User
Private mI_DAS  As ICHCSII_DataAccess.ICHCSII_DAS
Private mShared As CWShared
Private mDatabase   As CHCSII_CONN.Conn
Private mClinicOps  As CHCSII_ClinicClient.ClinicOpsEx
Private mOptions    As clsOptions
Private mFrequencyDetail As Collection

Private Sub Class_Terminate()
 
    Set mComm = Nothing
    Set mConfig = Nothing
    Set mLogon = Nothing
    Set mPatient = Nothing
    Set mUser = Nothing
    Set mI_DAS = Nothing
    Set mShared = Nothing
    Set mClinicOps = Nothing
    Set mOptions = Nothing
    Set mFrequencyDetail = Nothing
    
End Sub

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

Public Property Get Patient() As MMMHISPatient.Patient
  Set Patient = mPatient
End Property

Public Property Set Patient(ByRef NewValue As Object)
  Set mPatient = NewValue
End Property

Public Property Get User() As Object
  Set User = mUser
End Property

Public Property Set User(ByRef NewValue As Object)
  Set mUser = NewValue
End Property

Public Property Get SharedObj() As CWShared
    If mShared Is Nothing Then Set mShared = New CWShared
    Set SharedObj = mShared
End Property

Public Property Get ClinicOps() As CHCSII_ClinicClient.ClinicOpsEx
    If mClinicOps Is Nothing Then
        Set mClinicOps = New CHCSII_ClinicClient.ClinicOpsEx
    End If
    Set ClinicOps = mClinicOps
End Property

Public Property Get Options() As clsOptions
    If mOptions Is Nothing Then
        Set mOptions = New clsOptions
    End If
    Set Options = mOptions
End Property

Public Property Set Options(ByRef NewValue As clsOptions)
  Set mOptions = NewValue
End Property


Public Property Get DataConn() As CHCSII_CONN.Conn
    If mDatabase Is Nothing Then
        Set mDatabase = New CHCSII_CONN.Conn
    End If
    Set DataConn = mDatabase
End Property


Public Property Get I_DAS() As ICHCSII_DAS
    If mI_DAS Is Nothing Then
        Set mI_DAS = DataConn.CHCSII_DAS()
    End If
    Set I_DAS = mI_DAS
End Property

Public Property Get FrequencyDetail() As Collection
    If mFrequencyDetail Is Nothing Then
        Dim clsFreq As clsFrequency
        Set clsFreq = New clsFrequency
        Set mFrequencyDetail = clsFreq.GetFrequencys
    End If
    Set FrequencyDetail = mFrequencyDetail
End Property

