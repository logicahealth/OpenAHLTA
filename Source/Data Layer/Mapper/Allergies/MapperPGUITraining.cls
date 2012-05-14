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
Attribute VB_Name = "MapperPGUITraining"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements DL_Interfaces.iMapperAllergies

Private Const MODULE As String = "DM_Allergies.MapperPGUITraining"

Private Function iMapperAllergies_DeleteAllergy(ByRef DelAllergy As DTOs.Allergy) As Boolean
    Err.Raise 12345, MODULE & ".iMapperAllergies_DeleteAlelrgy", "This method is not implemented in this database mode."
End Function

Private Sub iMapperAllergies_Init(ByRef Support As DL_Support.iParameters, ByRef Cache As DL_Support.CacheMapper)
    Err.Raise 12345, MODULE & ".iMapperAllergies_Init", "This method is not implemented in this database mode."
End Sub

Private Function iMapperAllergies_RetrieveAllergies(ByRef params As DL_Support.iParameters) As DTOs.Allergies
    Err.Raise 12345, MODULE & ".iMapperAllergies_RetrieveAllergies", "This method is not implemented in this database mode."
End Function

Private Function iMapperAllergies_RetrieveAllergens(ByRef params As DL_Support.iParameters) As DTOs.Allergens
    Err.Raise 12345, MODULE & ".iMapperAllergies_RetrieveAllergens", "This method is not implemented in this database mode."
End Function

Private Function iMapperAllergies_SaveAllergy(ByRef SaveAllergy As DTOs.Allergy) As Boolean
    Err.Raise 12345, MODULE & ".iMapperAllergies_SaveAllergy", "This method is not implemented in this database mode."
End Function

Private Function iMapperAllergies_RetrieveAllergyReactions(ByRef params As iParameters) As DTOs.AllergyReactions
    Err.Raise 12345, MODULE & ".iMapperAllergies_RetrieveAllergyReactions", "This method is not implemented in this database mode."
End Function

Private Function iMapperAllergies_SaveAllergyTransaction(ByRef AllergyTrans As DTOs.AllergyTransaction) As Boolean
    Err.Raise 12345, MODULE & ".iMapperAllergies_SaveAllergyTransaction", "This method is not implemented in this database mode."
End Function

Private Function iMapperAllergies_RetrieveAllergyTransactions(ByRef params As iParameters) As DTOs.AllergyTransactions
    Err.Raise 12345, MODULE & ".iMapperAllergies_RetrieveAllergytTransactions", "This method is not implemented in this database mode."
End Function

Private Function iMapperAllergies_RecordsetTransferToLDDB(ByRef col As Collection) As Long
    Err.Raise 12345, MODULE & ".RecordsetTransferToLDDB", "This method is not implemented in this database mode."
End Function

Private Function iMapperAllergies_SaveAllergies(ByRef SavAllergies As DTOs.Allergies) As Boolean
    Err.Raise 12345, MODULE & ".iMapperAllergies_SaveAllergies", "This method is not implemented in this database mode."
End Function

