VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MapperPGUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'03-16-04 - Scott Lehman
'  this class implements the iMapperAllergies interface for PGUI

Implements DL_Interfaces.iMapperAllergies

Private Const MODULE As String = "DM_Allergies.MapperPGUI"  'identifies this module for error messages
Private Const ERROR_NUMBER As Long = 12345                'arbitrary number to use when raising errors

Private mConn As CHCSII_CONN.Conn                 'CHCSII Connection supplied during initialization

Private Function iMapperAllergies_DeleteAllergy(ByRef DelAllergy As DTOs.Allergy) As Boolean
    On Error GoTo ErrHandler
    
    Dim oCHCS_Patient As ICHCSPatient  'ICHCSPatient allows us to manipulate "legacy" Allergies
       
    Set oCHCS_Patient = GetCHCSPatient
    oCHCS_Patient.AllergyDelete DelAllergy.IENNumber, DelAllergy.AllergenNCID, DelAllergy.Comment
    iMapperAllergies_DeleteAllergy = True
    GoTo CleanUp
    
ErrHandler:
    iMapperAllergies_DeleteAllergy = False
    Err.Raise Err.Number, MODULE & ".iMapperAllergies_DeleteAllergy", Err.Description

CleanUp:
    On Error Resume Next
    Set oCHCS_Patient = Nothing
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
    On Error GoTo ErrHandler
    
    Dim oRecordset As ADODB.Recordset               'contains the Allergies retrieved from legacy
    Dim oCHCS_Patient As ICHCSPatient               'allows us access to legacy Allergies
    Dim dtoReturn As DTOs.Allergies                 'the DTOs.Allergies object this function returns
    Dim dtoAllergy As DTOs.Allergy                  'a single DTOs.Allergy object used to build dtoReturn
    Dim param As DM_Allergies.AllergySearchParam    'AllergySearchParam object used typecast the params argument
    
    'instantiate local objects
    Set dtoReturn = New DTOs.Allergies
    Set oCHCS_Patient = GetCHCSPatient
    Set param = params
    
    'retrieve patient allergies from legacy
    Set oRecordset = oCHCS_Patient.RetrieveAllergies(param.PatientIENNumber)
    
    '.RetrieveAllergies method doesn't reliable position the cursor on the first record...
    If oRecordset.RecordCount > 0 Then
        oRecordset.MoveFirst
    End If

    'build the return DTO
    Do Until oRecordset.EOF
        Set dtoAllergy = New DTOs.Allergy
        
        With dtoAllergy
            .IENNumber = param.PatientIENNumber
            .Allergen = IsNullValue(oRecordset.Fields("ALLERGY_SELECTION"), 0)
            .AllergenNCID = IsNullValue(oRecordset.Fields("ALLERGY_SELECTION_IEN"), 0)
            .Comment = IsNullValue(oRecordset.Fields("COMMENT"), 0)
            .Clinician = "CHCS"
        End With
        
        dtoReturn.Add dtoAllergy
        oRecordset.MoveNext
    Loop
     
    'cleanup, and get out
    GoTo CleanUp

ErrHandler:

    'raise an error, and resume with cleanup
    Err.Raise Err.Number, MODULE & ".iMapperAllergies_RetrieveAllergies", Err.Description

CleanUp:
    On Error Resume Next
    
    'set return for this function
    Set iMapperAllergies_RetrieveAllergies = dtoReturn
    
    'release local objects
    If oRecordset.State = adStateOpen Then
        oRecordset.Close
    End If
    Set oRecordset = Nothing
    Set oCHCS_Patient = Nothing
    
End Function

Private Function iMapperAllergies_RetrieveAllergens(ByRef params As DL_Support.iParameters) As DTOs.Allergens
Dim SQLHelper As AllergiesSQL

    Set SQLHelper = New AllergiesSQL
    SQLHelper.Init mConn, LGS
    
    Set iMapperAllergies_RetrieveAllergens = SQLHelper.RetrieveAllergens(params)
    Set SQLHelper = Nothing

End Function

Private Function iMapperAllergies_SaveAllergy(ByRef SavAllergy As DTOs.Allergy) As Boolean
    On Error GoTo ErrHandler
    
    Dim oCHCS_Patient As ICHCSPatient                      'allows us to access "legacy" Allergies
    Dim dtoAllergies As DTOs.Allergies                     'used to retrieve existing Allergy list
    Dim SearchParams As DM_Allergies.AllergySearchParam    'used to hold the parameters when retrieving existing Allergies
    Dim bAdded As Boolean                                  'flag
    Dim i As Long                                          'loop index/counter
       
    Set oCHCS_Patient = GetCHCSPatient
    
    Set SearchParams = New DM_Allergies.AllergySearchParam
    With SearchParams
        .PatientIENNumber = SavAllergy.IENNumber
    End With
    
    'SCR #50803 - need to check whether we're adding or updating an Allergy
    bAdded = False
    Set dtoAllergies = iMapperAllergies_RetrieveAllergies(SearchParams)
    For i = 1 To dtoAllergies.count
        With dtoAllergies.Item(i)
            If .AllergenNCID = SavAllergy.AllergenNCID Then
                oCHCS_Patient.AllergyUpdate SavAllergy.IENNumber, SavAllergy.AllergenNCID, .Comment, SavAllergy.Comment
                bAdded = True
                Exit For
            End If
        End With
    Next
    
    If Not (bAdded) Then
        oCHCS_Patient.AllergyAdd SavAllergy.IENNumber, SavAllergy.AllergenNCID, SavAllergy.Comment
    End If
    
    iMapperAllergies_SaveAllergy = True
    GoTo CleanUp
    
ErrHandler:
    iMapperAllergies_SaveAllergy = False
    Err.Raise Err.Number, MODULE & ".iMapperAllergies_SaveAllergy", Err.Description

CleanUp:
    On Error Resume Next
    Set oCHCS_Patient = Nothing

End Function

Private Function iMapperAllergies_RetrieveAllergyReactions(ByRef params As iParameters) As DTOs.AllergyReactions
Dim SQLHelper As AllergiesSQL

    Set SQLHelper = New AllergiesSQL
    SQLHelper.Init mConn, LGS
    
    Set iMapperAllergies_RetrieveAllergyReactions = SQLHelper.RetrieveAllergyReactions(params)
    Set SQLHelper = Nothing
    
End Function

Private Function iMapperAllergies_SaveAllergyTransaction(ByRef AllergyTrans As DTOs.AllergyTransaction) As Boolean
Dim SQLHelper As AllergiesSQL

    Set SQLHelper = New AllergiesSQL
    SQLHelper.Init mConn, LGS
    
    iMapperAllergies_SaveAllergyTransaction = SQLHelper.SaveAllergyTransaction(AllergyTrans)
    Set SQLHelper = Nothing

End Function

Private Function iMapperAllergies_RetrieveAllergyTransactions(ByRef params As iParameters) As DTOs.AllergyTransactions
    Err.Raise ERROR_NUMBER, MODULE & ".RetrieveAllergyTransactions", "This method is not implemented in this database mode."
End Function

Private Function iMapperAllergies_RecordsetTransferToLDDB(ByRef col As Collection) As Long
    Err.Raise ERROR_NUMBER, MODULE & ".RecordsetTransferToLDDB", "This method is not implemented in this database mode."
End Function

Private Function iMapperAllergies_SaveAllergies(ByRef SavAllergies As DTOs.Allergies) As Boolean
    Err.Raise ERROR_NUMBER, MODULE & ".iMapperAllergies_SaveAllergies", "This method is not implemented in this database mode."
End Function

