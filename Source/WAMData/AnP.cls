VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AnP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IPersist

Private mvarValidation As Validation

Private mvarCPT4s As CPT4s
Private mvarICD9s As ICD9s
Private mvarConsults As Consults
Private mvarLabs As Labs
Private mvarMeds As Meds
Private mvarRads As Rads
Private mvarPIs As PatientInstructions

Public Property Get PatientInstructions() As PatientInstructions
Set PatientInstructions = mvarPIs
End Property

Public Property Let PatientInstructions(vData As PatientInstructions)
  Set mvarPIs = vData
End Property

Public Property Get Rads() As Rads
    If mvarRads Is Nothing Then
        Set mvarRads = New Rads
    End If

    Set Rads = mvarRads
End Property

Public Property Set Rads(vData As Rads)
    Set mvarRads = vData
End Property

Public Property Get Meds() As Meds
    If mvarMeds Is Nothing Then
        Set mvarMeds = New Meds
    End If

    Set Meds = mvarMeds
End Property

Public Property Set Meds(vData As Meds)
    Set mvarMeds = vData
End Property

Public Property Get Labs() As Labs
    If mvarLabs Is Nothing Then
        Set mvarLabs = New Labs
    End If

    Set Labs = mvarLabs
End Property

Public Property Set Labs(vData As Labs)
    Set mvarLabs = vData
End Property

Public Property Get Consults() As Consults
    If mvarConsults Is Nothing Then
        Set mvarConsults = New Consults
    End If

    Set Consults = mvarConsults
End Property

Public Property Set Consults(vData As Consults)
    Set mvarConsults = vData
End Property

Public Property Get ICD9s() As ICD9s
    Set ICD9s = mvarICD9s
End Property

Public Property Set ICD9s(vData As ICD9s)
    Set mvarICD9s = vData
End Property

Public Property Get CPT4s() As CPT4s
    Set CPT4s = mvarCPT4s
End Property

Public Property Set CPT4s(vData As CPT4s)
    Set mvarCPT4s = vData
End Property

Private Sub Class_Initialize()
    Set mvarValidation = New Validation
    
    Set mvarICD9s = New ICD9s
    mvarICD9s.Init mvarValidation
    
    Set mvarCPT4s = New CPT4s
    mvarCPT4s.Init mvarValidation
    
    Set mvarPIs = New PatientInstructions
    mvarPIs.Init mvarValidation
    
    Set mvarRads = New Rads
    mvarRads.Init mvarValidation
    
    Set mvarMeds = New Meds
    mvarMeds.Init mvarValidation
    
    Set mvarLabs = New Labs
    mvarLabs.Init mvarValidation
    
    Set mvarConsults = New Consults
    mvarConsults.Init mvarValidation
    
    mvarValidation.Init mvarICD9s, mvarCPT4s, Nothing
    
End Sub
Private Sub Class_Terminate()
    
    mvarValidation.Release          'Possible Circular Reference
    Set mvarValidation = Nothing
    
    Set mvarPIs = Nothing
    Set mvarICD9s = Nothing
    Set mvarCPT4s = Nothing
    Set mvarValidation = Nothing
    
    Set mvarRads = Nothing
    Set mvarMeds = Nothing
    Set mvarLabs = Nothing
    Set mvarConsults = Nothing
    
End Sub

Private Function IPersist_ChildFactory() As Object
'
End Function

Private Function IPersist_Persist() As String
    IPersist_Persist = Persist(Me)
End Function

Private Function IPersist_Restore(ByVal Persist As String) As Boolean
    IPersist_Restore = Restore(Me, Persist)
End Function

