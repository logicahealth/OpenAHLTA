VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iMapperEncounter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Sub Init(ByRef Support As iParameters, ByRef Cache As CacheMapper)
End Sub
'Returns a sub set of encounters based on search criteria such as Encounter Type (Inpatient, Outpatient, etc.),
'Grouping (Open, Cosign, Completed, etc.), include section data or include documents.
Public Function EncounterList(ByRef Params As iParameters) As DTOs.Encounters
End Function

'Locks an encounter section.  Only one user can edit a section of an encounter at a time.
Public Function LockSection(ByRef Params As iParameters) As EncLocks
End Function

'Returns a partical or complete encounter record depending on passed in parameters.
'A complete encounter record includes all section data.
Public Function Retrieve(ByRef Params As iParameters) As DTOs.Encounter
End Function

'Saves any changed data in the encounter record.
Public Function Save(ByRef Encounter As DTOs.Encounter) As Boolean
End Function

'Unlocks an encounter section.  Only one user can edit a section of an encounter at a time.
Public Function UnlockSection(ByRef Params As iParameters) As Boolean
End Function
