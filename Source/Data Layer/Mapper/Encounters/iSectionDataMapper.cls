VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iSectionDataMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Public Sub Init(ByRef Conn As CHCSII_CONN.Conn, ByRef Cache As DL_Support.CacheMapper, _
                ByVal DataAccessMode As CHCSII_CONN.eBackend, Optional ByVal CacheConfigItems As Boolean = False)
        
End Sub
Public Function RetrieveAutocites(ByRef Encounter As DTOs.Encounter, ByRef Params As iParameters) As Boolean
End Function

Public Function SaveAutocites(ByRef Encounter As DTOs.Encounter) As Boolean
End Function

Public Function RetrieveRFV(ByRef Encounter As DTOs.Encounter, ByRef Params As iParameters) As Boolean
End Function

Public Function SaveRFV(ByRef Encounter As DTOs.Encounter) As Boolean
End Function

Public Function RetrieveVitals(ByRef Encounter As DTOs.Encounter, ByRef Params As iParameters) As Boolean
End Function

Public Function SaveVitals(ByRef Encounter As DTOs.Encounter) As Boolean
End Function

Public Function RetrieveSO(ByRef Encounter As DTOs.Encounter, ByRef Params As iParameters) As Boolean
End Function

Public Function SaveSO(ByRef Encounter As DTOs.Encounter) As Boolean
End Function

Public Function RetrieveAP(ByRef Encounter As DTOs.Encounter, ByRef Params As iParameters) As Boolean
End Function

Public Function SaveAP(ByRef Encounter As DTOs.Encounter) As Boolean
End Function

Public Function RetrieveDispo(ByRef Encounter As DTOs.Encounter, ByRef Params As iParameters) As Boolean
End Function

Public Function SaveDispo(ByRef Encounter As DTOs.Encounter) As Boolean
End Function

Public Function RetrieveHistory(ByRef Encounter As DTOs.Encounter, ByRef Params As iParameters) As Boolean
End Function

Public Function SaveHistory(ByRef Encounter As DTOs.Encounter) As Boolean
End Function

Public Function RetrieveEducation(ByRef Encounter As DTOs.Encounter, ByRef Params As iParameters) As Boolean
End Function

Public Function SaveEducation(ByRef Encounter As DTOs.Encounter) As Boolean
End Function

Public Function RetrieveDental(ByRef Encounter As DTOs.Encounter, ByRef Params As iParameters) As Boolean
End Function

Public Function SaveDental(ByRef Encounter As DTOs.Encounter) As Boolean
End Function

Public Function RetrieveEncNotes(ByRef Encounter As DTOs.Encounter, ByRef Params As iParameters) As Boolean
End Function

Public Function SaveEncNotes(ByRef Encounter As DTOs.Encounter) As Boolean
End Function
