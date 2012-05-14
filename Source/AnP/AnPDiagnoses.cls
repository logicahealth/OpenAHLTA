'***************************************************************************************
'
'  Copyright (c) 2007-2012 Northrop Grumman Corporation
'
'  Licensed by Tricare Management Activity under license from the Copyright owner.
'
'  This text file must be included in all Derivative Works of the licensed Source Code.
'
'***************************************************************************************

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AnPDiagnoses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Collection class for AnPDiagnosis objects.

Option Explicit

Private mKeyIndex As Long         'Used to assign a unique key to each added diagnosis.
                                  'Loaded diagnosis use their dataid as the Key.

Private mDiagnoses As Collection

Private Sub Class_Initialize()
  Set mDiagnoses = New Collection
End Sub

Public Function AddDiagnosis() As AnPDiagnosis
Dim Diagnosis As AnPDiagnosis

  On Error GoTo ErrHandler

  Set Diagnosis = New AnPDiagnosis
  
  Diagnosis.Priority = NextPriority
  AddDiagnosisToCollection Diagnosis
  
  Set AddDiagnosis = Diagnosis

  Exit Function

ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnPDiagnoses.AddDiagnosis", "A/P", vbExclamation

End Function

'Gets the next available key and adds the diagnosis to the collection using that key.
Private Sub AddDiagnosisToCollection(ByRef Diagnosis As AnPDiagnosis)

  On Error GoTo ErrHandler

Retry:
  mKeyIndex = mKeyIndex + 1

  mDiagnoses.Add Diagnosis, CStr(mKeyIndex)

  Diagnosis.Key = mKeyIndex

  Exit Sub

ErrHandler:
  If Err.Number <> 6 Then Resume Retry  '6 = Overflow

End Sub

Public Sub AssociateAllQueued(ByVal OrderType As enumOrderType, ByVal OrderKey As String)
Dim Diagnosis As AnPDiagnosis
  
  For Each Diagnosis In mDiagnoses
    Diagnosis.AssociateQueuedOrders OrderType, OrderKey
  Next

End Sub

Public Function Count() As Long
  On Error Resume Next
  
  Count = mDiagnoses.Count
End Function

Public Function Filter(Optional ByVal FilterType As enumFilterType = 0) As Collection

  On Error GoTo ErrHandler

  If FilterType = fltNone Then
    Set Filter = mDiagnoses
  Else
    Set Filter = FilterDiagnoses(FilterType)
  End If

  Exit Function

ErrHandler:
  Set Filter = New Collection
  gShared.ShowVBError Err.Number, Err.Description, "AnPDiagnoses.Filter", "A/P", vbExclamation

End Function

Private Function FilterDiagnoses(ByVal FilterType As enumFilterType) As Collection
Dim i As Long
Dim Filtered As Collection
Dim Associations As Collection

  On Error GoTo ErrHandler

  Set Filtered = New Collection

  Select Case FilterType
    Case fltActive
      For i = 1 To mDiagnoses.Count
        Select Case mDiagnoses.Item(i).State
          Case dsNew, dsSaved, dsModified
            Filtered.Add mDiagnoses.Item(i), CStr(mDiagnoses.Item(i).Priority)
        End Select
      Next

    Case fltDirty
      For i = 1 To mDiagnoses.Count
        Select Case mDiagnoses.Item(i).State
          Case dsNew, dsDelete, dsModified
            Filtered.Add mDiagnoses.Item(i)
            
          Case Else
            Set Associations = mDiagnoses.Item(i).Filter(, fltDirty)
            
            If Associations.Count > 0 Then
              Filtered.Add mDiagnoses.Item(i)
            End If
        End Select
      Next

    Case fltSaved
      For i = 1 To mDiagnoses.Count
        Select Case mDiagnoses.Item(i).State
          Case dsSaved
            Filtered.Add mDiagnoses.Item(i), CStr(mDiagnoses.Item(i).Priority)
        End Select
      Next
  End Select

  Set FilterDiagnoses = Filtered

  Exit Function

ErrHandler:
  Set FilterDiagnoses = Filtered
  gShared.ShowVBError Err.Number, Err.Description, "AnPDiagnoses.FilterDiagnoses", "A/P", vbExclamation

End Function

'Check all active diagnoses to see if an Order is associated to a diagnosis.
Public Function IsOrderAssociated(ByVal OrderType As enumOrderType, ByVal OrderKey As String, Optional ByVal SavedOnly As Boolean = False) As Boolean
Dim Diagnosis As AnPDiagnosis

  On Error GoTo ErrHandler
  
  For Each Diagnosis In mDiagnoses
   Select Case Diagnosis.State
    Case dsNew, dsSaved, dsModified
      If Diagnosis.IsOrderAssociated(OrderType, OrderKey, SavedOnly) Then
        IsOrderAssociated = True
        Exit For
      End If
    End Select
  Next
  
  Exit Function

ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnPDiagnoses.IsOrderAssociated", "A/P", vbExclamation
  
End Function

Public Function Item(ByVal Key As Variant) As AnPDiagnosis
  On Error Resume Next
  Set Item = mDiagnoses.Item(Key)
End Function

Public Function Load() As Boolean
Dim EncounterID As String
Dim Proxy As ProxyProblems

  On Error GoTo ErrHandler

  EncounterID = CStr(gBB.CurrentEncounter.EncounterID)

  Set Proxy = gBB.ProblemInterface
    
  Set mDiagnoses = Proxy.LoadDiagnoses(EncounterID)
  
  Load = True

  Exit Function

ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnPDiagnoses.Load", "A/P", vbExclamation

End Function

'Returns the next highest priority available for a diagnosis.
Private Function NextPriority() As Long
Dim Highest As Long
Dim Diagnosis As AnPDiagnosis

  On Error GoTo ErrHandler

  For Each Diagnosis In mDiagnoses
    Select Case Diagnosis.State
      Case dsModified, dsNew, dsSaved
        If Diagnosis.Priority > Highest Then
          Highest = Diagnosis.Priority
        End If
      End Select
  Next
  
  Highest = Highest + 1
  
  NextPriority = Highest
  
  Exit Function

ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnPDiagnoses.NextPriority", "A/P", vbExclamation
  
End Function

Private Sub RemoveDiagnoses()
Dim i As Long

  On Error GoTo ErrHandler

  For i = mDiagnoses.Count To 1 Step -1
    If mDiagnoses.Item(i).State = dsRemove Then
      mDiagnoses.Remove i
    End If
  Next

  Exit Sub

ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnPDiagnoses.RemoveDiagnoses", "A/P", vbExclamation
  Resume Next

End Sub

Public Sub Reset()
  Set mDiagnoses = Nothing
  Set mDiagnoses = New Collection
End Sub

'Returns the RTF string for all diagnoses and their associations.
Public Function RTF(Optional ByVal SavedOnly As Boolean = False) As String
Dim Diagnosis As AnPDiagnosis
Dim Diagnoses As Collection
Dim AddDiagnosis As Boolean
Dim DiagnosesRTF() As String
Dim i As Long

  On Error GoTo ErrHandler

  Set Diagnoses = FilterDiagnoses(fltActive)
  
  ReDim DiagnosesRTF(Diagnoses.Count + 1)
  
  For Each Diagnosis In Diagnoses
    With Diagnosis
      If SavedOnly Then
        AddDiagnosis = (.State = dsSaved)
      Else
        AddDiagnosis = (.State = dsModified Or .State = dsNew Or .State = dsSaved)
      End If
    
      If AddDiagnosis Then
        DiagnosesRTF(.Priority) = .RTF(SavedOnly)
      End If
    End With
  Next
 
  For i = 0 To Diagnoses.Count
    RTF = RTF & DiagnosesRTF(i)
  Next
  
  Exit Function

ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnPDiagnoses.RTF", "A/P", vbExclamation
  
End Function

'Save all diagnoses.
Public Function Save() As Boolean

Dim j As Long
Dim i As Long
Dim Success As Boolean

   On Error GoTo ErrHandler

   Save = True
  
   'JT BEGIN SCR 43293 - Added an outer loop so we can process deletes first.  This way if an error occurs during save
   'we might be able to avoid duplicate diagnosis priorities
   For j = 0 To 1
      For i = 1 To mDiagnoses.Count
         If (j = 0 And mDiagnoses.Item(i).State = dsDelete) Or _
            (j = 1 And mDiagnoses.Item(i).State <> dsDelete) Then
            Success = mDiagnoses.Item(i).Save
         End If
         Save = Save And Success
      Next i
  Next j
  'JT SCR END 43293
  
  RemoveDiagnoses

  Exit Function

ErrHandler:
  gShared.ShowVBError Err.Number, Err.Description, "AnPDiagnoses.Save", "A/P", vbExclamation

End Function

'Unassociate an order from all diagnoses if associated.
Public Sub UnassociateFromAll(ByVal OrderType As enumOrderType, ByVal OrderKey As String)
Dim Diagnosis As AnPDiagnosis
  
  For Each Diagnosis In mDiagnoses
    Diagnosis.UnassociateOrder OrderType, OrderKey
  Next

End Sub

'Called when a diagnosis' state is set to delete.   This method shifts the diagnoses'
'priorities.
Public Sub UpdatePriorities(ByVal PriorityDeleted As Long)
Dim Diagnosis As AnPDiagnosis
  
  For Each Diagnosis In mDiagnoses
    If Diagnosis.Priority > PriorityDeleted Then
      Diagnosis.Priority = Diagnosis.Priority - 1
    End If
  Next

End Sub

'Order data id is not set until the order is saved.  After the order is saved, then all
'order associations are updated with the order data id.  The order data id is used when
'saving the order association.
Public Sub UpdateOrderDataID(ByVal OrderType As enumOrderType, ByVal OrderKey As String, ByVal OrderDataID As String)
Dim Diagnosis As AnPDiagnosis
  
  For Each Diagnosis In mDiagnoses
    Diagnosis.UpdateOrderDataID OrderType, OrderKey, OrderDataID
  Next

End Sub
