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
Attribute VB_Name = "MapperTraining"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements DL_Interfaces.iMapperAllergies

Private Const MODULE As String = "DM_Allergies.MapperTraining"     'used to identify this module when throwing errors
Private Const ERROR_NUMBER As Long = 123456             'arbitrary number for "method not supported" errors

Private mConn As CHCSII_CONN.Conn                       'CHCSII Connection

Private Function iMapperAllergies_DeleteAllergy(ByRef DelAllergy As DTOs.Allergy) As Boolean
Dim SQLHelper As AllergiesSQL

    Set SQLHelper = New AllergiesSQL
    SQLHelper.Init mConn, ITT
    
    iMapperAllergies_DeleteAllergy = SQLHelper.DeleteAllergy(DelAllergy)
    Set SQLHelper = Nothing

End Function

Private Sub iMapperAllergies_Init(ByRef Support As DL_Support.iParameters, ByRef Cache As DL_Support.CacheMapper)
    On Error GoTo ErrHandler

    'Save objects given to us by DataAccess Factory
    Set mConn = Support.Item("DBConn").Value
    
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, MODULE & ".iMapperAllergies_Init", Err.Description

End Sub

Private Function iMapperAllergies_RetrieveAllergies(ByRef params As DL_Support.iParameters) As DTOs.Allergies
Dim SQLHelper As AllergiesSQL

    Set SQLHelper = New AllergiesSQL
    SQLHelper.Init mConn, ITT
    
    Set iMapperAllergies_RetrieveAllergies = SQLHelper.RetrieveAllergies(params)
    Set SQLHelper = Nothing

End Function

Private Function iMapperAllergies_RetrieveAllergens(ByRef params As DL_Support.iParameters) As DTOs.Allergens
Dim SQLHelper As AllergiesSQL

    Set SQLHelper = New AllergiesSQL
    SQLHelper.Init mConn, ITT
    
    Set iMapperAllergies_RetrieveAllergens = SQLHelper.RetrieveAllergens(params)
    Set SQLHelper = Nothing

End Function

Private Function iMapperAllergies_SaveAllergy(ByRef SavAllergy As DTOs.Allergy) As Boolean
Dim SQLHelper As AllergiesSQL

    Set SQLHelper = New AllergiesSQL
    SQLHelper.Init mConn, ITT
    
    iMapperAllergies_SaveAllergy = SQLHelper.SaveAllergy(SavAllergy)
    Set SQLHelper = Nothing

End Function

Private Function iMapperAllergies_RetrieveAllergyReactions(ByRef params As iParameters) As DTOs.AllergyReactions
Dim SQLHelper As AllergiesSQL

    Set SQLHelper = New AllergiesSQL
    SQLHelper.Init mConn, ITT
    
    Set iMapperAllergies_RetrieveAllergyReactions = SQLHelper.RetrieveAllergyReactions(params)
    Set SQLHelper = Nothing

End Function

Private Function iMapperAllergies_SaveAllergyTransaction(ByRef AllergyTrans As DTOs.AllergyTransaction) As Boolean
Dim SQLHelper As AllergiesSQL

    Set SQLHelper = New AllergiesSQL
    SQLHelper.Init mConn, ITT
    
    iMapperAllergies_SaveAllergyTransaction = SQLHelper.SaveAllergyTransaction(AllergyTrans)
    Set SQLHelper = Nothing
    
End Function

Private Function iMapperAllergies_RetrieveAllergyTransactions(ByRef params As iParameters) As DTOs.AllergyTransactions
    Err.Raise ERROR_NUMBER, MODULE & ".iMapperAllergies_RetrieveAllergyTransactions", "This method is not implemented in this database mode."
End Function

Private Function iMapperAllergies_RecordsetTransferToLDDB(ByRef col As Collection) As Long
    Err.Raise ERROR_NUMBER, MODULE & ".iMapperAllergies_RecordsetTransferToLDDB", "This method is not implemented in this database mode."
End Function

Private Function iMapperAllergies_SaveAllergies(ByRef SavAllergies As DTOs.Allergies) As Boolean
Dim SQLHelper As AllergiesSQL

    Set SQLHelper = New AllergiesSQL
    SQLHelper.Init mConn, ITT
    
    iMapperAllergies_SaveAllergies = SQLHelper.SaveAllergies(SavAllergies)
    Set SQLHelper = Nothing

End Function

