VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MapperUnitTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements DL_Interfaces.iMapperAllergies

Private Const MODULE As String = "DM_Allergies.MapperUnitTest"    'string to identify this module when throwing errors
Private Const ERROR_NUMBER As Long = 12345              'arbitrary error number for "method not supported" errors

Private Function iMapperAllergies_DeleteAllergy(ByRef DelAllergy As DTOs.Allergy) As Boolean
    Err.Raise ERROR_NUMBER, MODULE & ".iMapperAllergies_DeleteAllergy", "This method has not been implemented in this data mode."
End Function

Private Sub iMapperAllergies_Init(ByRef Support As DL_Support.iParameters, ByRef Cache As DL_Support.CacheMapper)
    Err.Raise ERROR_NUMBER, MODULE & ".iMapperAllergies_Init", "This method has not been implemented in this data mode."
End Sub

Private Function iMapperAllergies_RetrieveAllergies(ByRef params As DL_Support.iParameters) As DTOs.Allergies
    Err.Raise ERROR_NUMBER, MODULE & ".iMapperAllergies_RetrieveAllergies", "This method has not been implemented in this data mode."
End Function

Private Function iMapperAllergies_RetrieveAllergens(ByRef params As DL_Support.iParameters) As DTOs.Allergens
    Err.Raise ERROR_NUMBER, MODULE & ".iMapperAllergies_RetrieveAllergens", "This method has not been implemented in this data mode."
End Function

Private Function iMapperAllergies_SaveAllergy(ByRef SaveAllergy As DTOs.Allergy) As Boolean
    Err.Raise ERROR_NUMBER, MODULE & ".iMapperAllergies_SaveAllergy", "This method has not been implemented in this data mode."
End Function

Private Function iMapperAllergies_RetrieveAllergyReactions(ByRef params As iParameters) As DTOs.AllergyReactions
    Err.Raise ERROR_NUMBER, MODULE & ".iMapperAllergies_RetrieveAllergyReactions", "This method has not been implemented in this data mode."
End Function

Private Function iMapperAllergies_SaveAllergyTransaction(ByRef AllergyTrans As DTOs.AllergyTransaction) As Boolean
    Err.Raise ERROR_NUMBER, MODULE & ".iMapperAllergies_SaveAllergyTransaction", "This method has not been implemented in this data mode."
End Function

Private Function iMapperAllergies_RetrieveAllergyTransactions(ByRef params As iParameters) As DTOs.AllergyTransactions
    Err.Raise ERROR_NUMBER, MODULE & ".iMapperAllergies_RetrieveAllergyTransactions", "This method has not been implemented in this data mode."
End Function

Private Function iMapperAllergies_RecordsetTransferToLDDB(ByRef col As Collection) As Long
    Err.Raise ERROR_NUMBER, MODULE & ".iMapperAllergies_RecordsetTransferToLDDB", "This method has not been implemented in this data mode."
End Function

Private Function iMapperAllergies_SaveAllergies(ByRef SavAllergies As DTOs.Allergies) As Boolean
    Err.Raise ERROR_NUMBER, MODULE & ".iMapperAllergies_SaveAllergies", "This method has not been implemented in this data mode."
End Function

