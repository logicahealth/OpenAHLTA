VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iMapperAllergies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'This should only be called from the DataAccess library
Public Sub Init(ByRef Support As iParameters, ByRef Cache As CacheMapper)
End Sub

'Gets Allergy information, Params is DM_Allergies.AllergySearchParam
Public Function RetrieveAllergies(Params As iParameters) As dtos.Allergies
End Function

'Saves the allergy to the appropriate data store
Public Function SaveAllergy(SavAllergy As dtos.Allergy) As Boolean
End Function

'Marks the Allergy for deletion
Public Function DeleteAllergy(DelAllergy As dtos.Allergy) As Boolean
End Function

'Gets Allergen information.  Params is DM_Allergies.AllergenSearchParam
Public Function RetrieveAllergens(Params As iParameters) As dtos.Allergens
End Function

'Gets Reaction information.  Params is DM_Allergies.ReactionSearchParams
Public Function RetrieveAllergyReactions(Params As iParameters) As dtos.AllergyReactions
End Function

'Saves the Transaction
Public Function SaveAllergyTransaction(AllergyTrans As dtos.AllergyTransaction) As Boolean
End Function

'Gets the TransactionInformation.  Params is DM_Allergies.TransactionSearchParams
Public Function RetrieveAllergyTransactions(Params As iParameters) As dtos.AllergyTransactions
End Function

'Support for Data Transfer
Public Function RecordsetTransferToLDDB(col As Collection) As Long
End Function

'All previously existing Allergies are deleted and replaced with the new Allergies
Public Function SaveAllergies(SavAllergies As dtos.Allergies) As Boolean
End Function


