VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProxyADM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Sub UpdateCPTPriorty(ByVal CPT As String, ByVal Priority As String, ByRef CPT4Codes As WAMData.CPT4s)
Dim CPTCode As WAMData.CPT4
  
  On Error GoTo ErrHandler
  
  For Each CPTCode In CPT4Codes
    If CPT = CPTCode.CPT4_CODE Then
      If InStr(1, CPTCode.CPT4_DIAGNOSIS_PRIORITY, Priority) = 0 Then
        CPTCode.CPT4_DIAGNOSIS_PRIORITY = CPTCode.CPT4_DIAGNOSIS_PRIORITY & Priority
      End If
      
      Exit For
    End If
  Next
  
  Exit Sub

ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyADM.UpdateCPTPriorty", "A/P", vbExclamation
  
End Sub

Private Sub DisplayWarning(ByVal Message As String)
Dim Warning As WarningDialog

  On Error GoTo ErrHandler
  
  Set Warning = New WarningDialog
  
  Warning.Display Message, True

  Set Warning = Nothing
      
  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyADM.DisplayWarning", "A/P", vbExclamation
      
End Sub

Private Function FindExistingICD9Diagnosis(ByVal ICD9 As String, ByRef ICD9Codes As WAMData.ICD9s) As String
Dim ICD9Code As WAMData.ICD9

  For Each ICD9Code In ICD9Codes
    If ICD9 = ICD9Code.ICD_CODE Then
      FindExistingICD9Diagnosis = ICD9Code.ICD_DIAGNOSIS_PRIORITY
      Exit For
    End If
  Next
  
End Function

'BR (SCR 29571) Roll Up ICD9 Code (Don't Increment Priority if AddDiagnosis call fails)  1/17/03 Phil Crowder
'
'BR (SCR 29572) Roll Up CPTs to Highest Dup. ICD Code (Diagnosis are pulled in priority
'order and added based on the NextPriority var.)  1/17/03 Phil Crowder
Public Sub Save()
Dim Enc As CHCSIIEncounterOps.Encounter
Dim Diagnoses As Collection
Dim Diagnosis As AnPDiagnosis
Dim NextPriority As String
Dim AssociationPriority As String
Dim i As Long
Dim Success As Boolean

  On Error GoTo ErrHandler
  
  Set Diagnoses = gBB.AllDiagnoses.Filter(fltSaved)

  Set Enc = gBB.CurrentEncounter
  
  Exit Sub
  
ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "ProxyADM.Save", "A/P", vbExclamation

End Sub
