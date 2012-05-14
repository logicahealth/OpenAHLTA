VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EncProviderProcedure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const MODULE_NAME = "EncProviderProcedure"

Private msPrevDataId As String   'Value in the database
Private msPrevProvNCID As String 'Value in the database
Private msDataId As String          'Number(20) - too big for long
Private msSnoID As String
Private msDescription As String
Private meCurrentState As enumRecordState

Public Sub Init(ByVal sDataId As String, ByVal sSnoID As String)
   msDataId = sDataId
   msSnoID = sSnoID
   meCurrentState = rsNew
End Sub

Friend Sub Load(ByVal sDataId As String, ByVal sSnoID As String, _
   Optional ByVal sProvNCID As String = vbNullString, Optional ByVal sDesc As String = vbNullString)
   
   msPrevDataId = sDataId
   msPrevProvNCID = sProvNCID
   msDataId = sDataId
   msSnoID = sSnoID
   msDescription = sDesc
   meCurrentState = rsSaved
   
End Sub

Public Property Get Description() As String

On Error GoTo Error_Handler:

Dim oIConvert As IConvert

   If Len(msDescription) = 0 And IsNumeric(Left(msSnoID, 1)) Then
      msDescription = oIConvert.SnoDescription(msSnoID)
   End If
   Description = msDescription
   
Error_Handler:
   Set oIConvert = Nothing
   
   If CBool(Err) Then HandleErrors Err, MODULE_NAME, "GEMS_SaveEncounterProvider"
  
End Property

Public Property Let DataId(ByVal sNewValue As String)
   msDataId = sNewValue
End Property

Public Property Get DataId() As String
   DataId = msDataId
End Property

Public Property Get SnoID() As String
   SnoID = msSnoID
End Property

Public Property Let SnoID(ByVal sNewValue As String)
   msSnoID = sNewValue
End Property

Public Property Get State() As enumRecordState
  State = meCurrentState
End Property

Public Property Let State(ByVal eNewValue As enumRecordState)
  meCurrentState = eNewValue
End Property

Friend Function GEMS_SaveProcedure(ByVal sEncounterNumber As String, ByVal sFacilityNCID As String, _
         ByVal sProviderNCID As String) As Boolean

On Error GoTo Error_Handler

Dim sSql As String
Dim oProc As EncProviderProcedure
Dim oSQL As ICHCSII_SQL.ISqlOps
         
   Select Case meCurrentState
   
   Case rsNew
   
      sSql = "INSERT INTO Encounter_Procedure " & _
            " ( EncounterNumber, FacilityNCID, Provider_NCID, DataID) " & _
            " VALUES " & _
            " ( " & sEncounterNumber & ", " & sFacilityNCID & ", " & sProviderNCID & ", " & _
               msDataId & " )"
               
   Case rsModified

      sSql = "UPDATE Encounter_Procedure " & _
            " SET Provider_NCID = " & NCID & ", " & _
               " DataId = " & DataId & ", " & _
            " WHERE EncounterNumber = " & sEncounterNumber & _
               " AND FacilityNCID = " & sFacilityNCID & _
               " AND Provider_NCID = " & msPrevProvNCID & _
               " AND DataID = " & msPrevDataId
               
   Case rsDelete
            
      sSql = "DELETE FROM Encounter_Procedure " & _
            " WHERE EncounterNumber = " & sEncounterNumber & _
            "     AND FacilityNCID = " & sFacilityNCID & _
            "     AND Provider_NCID = " & msPrevProvNCID & _
            "     AND DataID = " & msPrevDataId
            
   End Select
   
   If Len(sSql) > 0 Then
      Set oSQL = gobjCHCSIIConn.CHCSII_SQLOPS(Auto)
      oSQL.Execute sSql
      
      If Me.State = rsDelete Then
         Me.State = rsRemove
      Else
         Me.State = rsSaved
      End If
      
      msPrevDataId = msDataId
      msPrevProvNCID = sProviderNCID
      
   End If
   
   GEMS_SaveProcedure = True
   
Error_Handler:
   Set oSQL = Nothing
   
   If CBool(Err) Then HandleErrors Err, MODULE_NAME, "GEMS_SaveProcedure"

End Function

Public Property Get IsDirty() As Boolean
   If State = rsNew Or State = rsDelete Or State = rsModified Then
      IsDirty = True
   Else
      IsDirty = False
   End If
End Property
