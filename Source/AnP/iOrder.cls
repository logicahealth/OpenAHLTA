VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum enumOrderState
  osUnknown
  osNew
  osDelete
  osRemove
  osSaved
  osQueued
  osModified
End Enum

Private mDataID As String
Private mKey As String
Private mName As String
Private mCurrentState As enumOrderState
Private mPrevState As enumOrderState

Public Property Let DataID(ByVal NewValue As String)
  If Len(mDataID) = 0 Then
    mDataID = NewValue
  End If
End Property

Public Property Get DataID() As String
  DataID = mDataID
End Property

Public Property Let Key(ByVal NewValue As String)
  If Len(mKey) = 0 Then
    mKey = NewValue
  End If
End Property

Public Property Get Key() As String
  Key = mKey
End Property

Public Property Let Name(ByVal NewValue As String)
  mName = NewValue
End Property

Public Property Get Name() As String
  Name = mName
End Property

Public Property Get RTF(Optional ByRef Diagnosis As AnPDiagnosis) As String

End Property

Public Property Let State(ByVal NewValue As enumOrderState)
Dim ValidNewState As Boolean
  
  ValidNewState = IsValidState(NewValue)
  
  If ValidNewState Then
    If mCurrentState = osUnknown Then
      mPrevState = NewValue
    Else
      mPrevState = mCurrentState
    End If
  
    mCurrentState = NewValue
  End If
  
End Property

Public Property Get State() As enumOrderState
  State = mCurrentState
End Property

Public Function CreateAssociation() As OrderAssociation
  
End Function

Public Sub Init(ByVal DataID As String, ByVal Name As String)
  mKey = DataID
  mDataID = DataID
  mName = Name
  
  mCurrentState = osSaved
  mPrevState = osSaved
End Sub

Public Function IsAssociated(Optional ByVal SavedOnly As Boolean = False) As Boolean

End Function

Private Function IsValidState(ByVal NewValue As enumOrderState) As Boolean

  Select Case mCurrentState
    Case osUnknown
      IsValidState = True
      
    Case osNew
      IsValidState = (NewValue = osDelete Or NewValue = osRemove)
    
    Case osQueued
      IsValidState = (NewValue = osNew Or NewValue = osDelete Or NewValue = osRemove _
                      Or (NewValue = osSaved And gShared.IsAppMode(modeCHCSI_GUI)))
      '@@@@  Check to see of last check needed
    Case osDelete, osRemove
      IsValidState = False
      
    Case osSaved
      IsValidState = (NewValue = osModified Or NewValue = osDelete Or NewValue = osRemove)
      
    Case osModified
      IsValidState = (NewValue = osDelete Or NewValue = osRemove)
  End Select

End Function

Public Sub RollbackState()
  mCurrentState = mPrevState
End Sub

Public Function Save() As Boolean
  
  Select Case mCurrentState
    Case osNew, osModified
      mCurrentState = osSaved
      mPrevState = osSaved
      
    Case osDelete
      mCurrentState = osRemove
      mPrevState = osRemove
  End Select
  
  Save = True
  
End Function

Public Sub Unassociate()

End Sub

Public Function UserHasPriv(ByVal PrivType As CHCSII_PrivilegeType) As Boolean

End Function
