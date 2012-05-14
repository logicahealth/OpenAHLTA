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
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Complaints"
'Class Name:    BlackBoard.cls
'Author:        Mark Broaddus
'Created:       9/20/2002
'Description:   This is the class factory for the App.  It also holds the references to the 5 objects passed in by CORE.
'Revisions:

Option Explicit

Private mComm As MMMHISComm.Comm
Private mConfig As MMMHISConfiguration.Config
Private mLogon As mmmhislogon.Logon
Private mPatient As MMMHISPatient.Patient
Private mUser As Object
Private moDatabase As CHCSII_CONN.Conn
Private moSQLOps As ICHCSII_SQL.ISqlOps ' .ISqlOpsEx
Private moGEMS_Concept As GEMS_ConceptCtrl
Private moARMWrapper As ARM.cARMWrapper
Private moShared As CWShared
Private mEncParent As Object
Private moRFVForm As frmReasonForVisit
Private moRFVClass As ReasonForVisit
Private moListManagement As ListManagement
Private moListManagementOps As ListManagementOps
Private moAllergy As Allergy
Private moSWSForm As frmSWS
Private moComplaints As Complaints

Public Property Get Complaints() As Complaints
    If moComplaints Is Nothing Then
        Set moComplaints = New Complaints
    End If
    
    Set Complaints = moComplaints
End Property

Private Sub Class_Initialize()
  Set moDatabase = New CHCSII_CONN.Conn
End Sub

Private Sub Class_Terminate()
 
    Set mComm = Nothing
    Set mConfig = Nothing
    Set mLogon = Nothing
    Set mPatient = Nothing
    Set mUser = Nothing
    Set moSQLOps = Nothing
    Set moGEMS_Concept = Nothing
    Set moARMWrapper = Nothing
    Set moShared = Nothing
    Set mEncParent = Nothing
    Set moRFVForm = Nothing
    Set moAllergy = Nothing
    Set moSWSForm = Nothing
    Set moDatabase = Nothing
End Sub

Public Property Get Comm() As MMMHISComm.Comm
  Set Comm = mComm
End Property

Public Property Set Comm(ByRef NewValue As MMMHISComm.Comm)
  Set mComm = NewValue
End Property

Public Property Get Config() As MMMHISConfiguration.Config
  Set Config = mConfig
End Property

Public Property Set Config(ByRef NewValue As MMMHISConfiguration.Config)
  Set mConfig = NewValue
End Property

Public Property Get EncParent() As ENCOUNTER_INTERFACES.IParent
  Set EncParent = mEncParent
End Property

Public Property Set EncParent(ByRef NewValue As ENCOUNTER_INTERFACES.IParent)
  Set mEncParent = NewValue
End Property

Public Property Get CurrentEncounter() As Encounter
If mEncParent Is Nothing Then Exit Property
    Set CurrentEncounter = mEncParent.CurrentEncounter
End Property

Public Property Get Logon() As mmmhislogon.Logon
  Set Logon = mLogon
End Property

Public Property Set Logon(ByRef NewValue As mmmhislogon.Logon)
  Set mLogon = NewValue
End Property

Public Property Get Patient() As MMMHISPatient.Patient
  Set Patient = mPatient
End Property

Public Property Set Patient(ByRef NewValue As MMMHISPatient.Patient)
  Set mPatient = NewValue
End Property

Public Property Get User() As Object
  Set User = mUser
End Property

Public Property Set User(ByRef NewValue As Object)
  Set mUser = NewValue
End Property

Public Property Get SQLOps() As ICHCSII_SQL.ISqlOps
    If moSQLOps Is Nothing Then
        Set moSQLOps = moDatabase.CHCSII_SQLOPS(Auto)
    End If
    
    Set SQLOps = moSQLOps
End Property

Public Property Get GEMS_Concept() As GEMS_ConceptCtrl
    If moGEMS_Concept Is Nothing Then Set moGEMS_Concept = New GEMS_ConceptCtrl
    Set GEMS_Concept = moGEMS_Concept
End Property

'Public Property Get DAS() As DasCtrl
'    If moDAS Is Nothing Then Set moDAS = New DasCtrl
'    Set DAS = moDAS
'End Property

Public Property Get ARM() As ARM.cARMWrapper
    Dim oARMFactory As ARM.cARMFactory
    If moARMWrapper Is Nothing Then
        Set oARMFactory = New ARM.cARMFactory
        Set moARMWrapper = oARMFactory.CreateARMApp("")
    End If
    Set ARM = moARMWrapper
End Property

Public Property Get SharedObj() As CWShared
    If moShared Is Nothing Then Set moShared = New CWShared
    Set SharedObj = moShared
End Property

Public Property Set RFVClass(ClassReference As ReasonForVisit)
    Set moRFVClass = ClassReference
End Property

Public Property Get RFVClass() As ReasonForVisit
    Set RFVClass = moRFVClass
End Property

Public Property Get AllergyObj() As Allergy
    If moAllergy Is Nothing Then
        Set moAllergy = goBB.Comm.InitializeOLEServer(NCID_ALLERGIES_APPLICATION)
    End If
    Set AllergyObj = moAllergy
End Property

Public Property Get ListMgmntOps() As ListManagementOps
    
    If moListManagementOps Is Nothing Then
        Set moListManagement = goBB.Comm.InitializeOLEServer(NCID_LIST_MGMT_APPLICATION)
        Set moListManagementOps = moListManagement.ListManagementOps
        Set moListManagement = Nothing
    End If
    
    Set ListMgmntOps = moListManagementOps
End Property

Friend Property Get RFVForm() As frmReasonForVisit
    If moRFVForm Is Nothing Then
        Set moRFVForm = New frmReasonForVisit
        Set moRFVForm.Parent = moRFVClass
    End If
    Set RFVForm = moRFVForm
End Property

Friend Property Get SWSForm() As frmSWS
    If moSWSForm Is Nothing Then
        Set moSWSForm = New frmSWS
    End If
    Set SWSForm = moSWSForm
End Property

Friend Function ReleaseRFVForm() As Boolean
    Set moRFVForm = Nothing
End Function

Public Sub ClearBlackBoard()
On Error Resume Next
  
    Set moSQLOps = Nothing
    Set moGEMS_Concept = Nothing
    Set moComplaints = Nothing
    
End Sub

