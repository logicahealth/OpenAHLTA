VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iEncDataMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Public Function Retrieve(ByRef Params As iParameters) As DTOs.Encounter
End Function

Public Function Save(ByRef Encounter As DTOs.Encounter) As Boolean
End Function

Public Sub Init(ByRef Conn As CHCSII_CONN.Conn, ByRef Cache As DL_Support.CacheMapper, _
                ByVal DataAccessMode As CHCSII_CONN.eBackend, Optional ByVal CacheConfigItems As Boolean = False)
        
End Sub

Public Function EncounterList(Params As DL_Support.iParameters) As DTOs.Encounter

End Function
