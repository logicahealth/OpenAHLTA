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
'Class Name:    BlackBoard.cls
'Author:        Mark Broaddus
'Created:       9/20/2002
'Description:   This is the class factory for the App.  It also holds the references to the 5 objects passed in by CORE.
'Revisions:

Option Explicit

Private mComm As MMMHISComm.Comm
Private mConfig As MMMHISConfiguration.Config
Private mLogon As MMMHISLogon.Logon
Private mPatient As Object
Private mUser As Object
Private moShared As CWShared
Private moCData As CardData
Private moAppt As CHCSII_Appointments.Appointments
Private moPatientList As Object

Private Sub Class_Terminate()
 
    Set mComm = Nothing
    Set mConfig = Nothing
    Set mLogon = Nothing
    Set mPatient = Nothing
    Set mUser = Nothing
    Set moShared = Nothing
    Set moAppt = Nothing
    
End Sub

Public Property Get Appts() As Appointments
    If moAppt Is Nothing Then
         Set moAppt = mComm.InitializeOLEServer(NCID_CLINIC_SCHEDULE_APPLICATION)
    End If
    
    Set Appts = moAppt
End Property

Public Property Get PatList() As Object
    If moPatientList Is Nothing Then
         Set moPatientList = mComm.InitializeOLEServer(NCID_PATIENT_LIST_APPLICATION)
    End If
    
    Set PatList = moPatientList
End Property

Public Property Get CData() As CardData
    
    If moCData Is Nothing Then
        Set moCData = New CardData
    End If
    
    Set CData = moCData
    
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

Public Property Get Patient() As Object
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
    If moShared Is Nothing Then Set moShared = New CWShared
    Set SharedObj = moShared
End Property

'Clear all module initialized through Core by A/P and the Encounter Parent object.
Public Sub ClearBlackBoard()
On Error Resume Next
  
    Set moShared = Nothing

End Sub

