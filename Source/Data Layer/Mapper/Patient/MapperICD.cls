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
Attribute VB_Name = "MapperICD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements DL_Interfaces.iMapperPatient

Private mobjMapperGeneric As DL_Interfaces.iMapperPatient


Private Sub Class_Initialize()
Dim objMapper As MapperGeneric
Dim oSQLFix As iSQLFix
Dim oPicklist As iPickList

On Error GoTo ErrHandler

  Set mobjMapperGeneric = New MapperGeneric
  Set oSQLFix = New SQLOracle
  Set oPicklist = New PickListGEMS
  
  Set objMapper = mobjMapperGeneric
  objMapper.BackEndDB = ICD
  Set objMapper.SQLFix = oSQLFix
  Set objMapper.PickList = oPicklist
  Exit Sub
  
ErrHandler:
  Err.Raise Err.Number, "DM_Patient.MapperICD.Initialize", Err.Description, Err.HelpFile, Err.HelpContext


End Sub

Private Function iMapperPatient_GetRankNCID(ByVal sRankCode As String, ByVal sBranch As String) As String
    Err.Raise Err.Number, "MapperICD.iMapperPatient_GetRankNCID", "GetRankNCID is not allowed in ICD mode.", Err.HelpFile, Err.HelpContext
End Function

Private Sub iMapperPatient_Init(Support As DL_Support.iParameters, Cache As DL_Support.CacheMapper)
    mobjMapperGeneric.Init Support, Cache
End Sub

Private Function iMapperPatient_Retrieve(Params As DL_Support.iParameters) As DTOs.Patients

On Error GoTo ErrHandler
    
    Set iMapperPatient_Retrieve = mobjMapperGeneric.Retrieve(Params)
    
    GoTo CleanUpAndExit
ErrHandler:
    Err.Raise Err.Number, "MapperICD.iMapperPatient_Retrieve", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
End Function

Private Function iMapperPatient_RetrieveSelectStrings(Params As DL_Support.iParameters) As DTOs.PatientSelectStrings
On Error GoTo ErrHandler
    
    Set iMapperPatient_RetrieveSelectStrings = mobjMapperGeneric.RetrieveSelectStrings(Params)
    
    GoTo CleanUpAndExit
ErrHandler:
    Err.Raise Err.Number, "MapperICD.iMapperPatient_RetrieveSelectStrings", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:

End Function

Private Function iMapperPatient_RetrieveRepresentation(ByRef Params As DL_Support.iParameters) As String
    On Error GoTo ErrHandler
    
    iMapperPatient_RetrieveRepresentation = mobjMapperGeneric.RetrieveRepresentation(Params)
    
    GoTo CleanUpAndExit
ErrHandler:
    Err.Raise Err.Number, "MapperICD.iMapperPatient_RetrieveRepresentation", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
End Function

Private Function iMapperPatient_Save(ByRef Patient As DTOs.Patient) As Boolean
    Dim oState As iState
On Error GoTo ErrHandler

    Set oState = Patient
    
    If oState.State = osNew Or oState.State = osUnknown Then
        iMapperPatient_Save = CreateNewPatient(Patient)
    Else
        iMapperPatient_Save = mobjMapperGeneric.Save(Patient)
    End If
    
    GoTo CleanUpAndExit
ErrHandler:
    Err.Raise Err.Number, "MapperICD.iMapperPatient_RetrieveSelectStrings", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:

End Function


Private Function iMapperPatient_TransferCHCSPatient(ByRef Params As DL_Support.iParameters) As DTOs.PatientTransferResult
    Err.Raise Err.Number, "MapperICD.iMapperPatient_RetrieveSelectStrings", "TransferCHCSPatient is not allowed in ICD mode.", Err.HelpFile, Err.HelpContext
End Function


Private Function CreateNewPatient(ByRef oPatient As DTOs.Patient) As Boolean
    Dim oState As iState
    Dim lUnitNumber As Long
    Dim oConn As CHCSII_CONN.Conn
    Dim oDAS As CHCSII_CONN.IADO_DAS
    Dim oCmd As ADODB.Command
    Dim oParam As ADODB.Parameter


On Error GoTo ErrHandler

    CreateNewPatient = False

    Set oConn = New Conn
    Set oDAS = oConn.ADO_DAS(ICD)
    Set oCmd = New ADODB.Command

    With oCmd
        .CommandText = "Mirror_LGS_Patient"
        .CommandType = adCmdStoredProc

        Set oParam = .CreateParameter("p_Enterpise_Unit_Number", adNumeric, adParamInput, 20, oPatient.UnitNumber)
        .Parameters.Append oParam
        If LenB(oPatient.DeersID) = 0 Then
            oPatient.DeersID = oPatient.UnitNumber
        End If
        Set oParam = .CreateParameter("p_DeersID", adVarChar, adParamInput, 100, oPatient.DeersID)
        .Parameters.Append oParam
        Set oParam = .CreateParameter("p_FacilityNCID", adNumeric, adParamInput, 20, CLng(oPatient.FacilityNCID))
        .Parameters.Append oParam
        Set oParam = .CreateParameter("p_PatientIEN", adVarChar, adParamInput, 100, oPatient.MedicalRecordNumber)
        .Parameters.Append oParam
        Set oParam = .CreateParameter("p_LastName", adVarChar, adParamInput, 200, oPatient.Name.Last_Name)
        .Parameters.Append oParam
        Set oParam = .CreateParameter("p_FirstName", adVarChar, adParamInput, 200, oPatient.Name.First_Name)
        .Parameters.Append oParam
        Set oParam = .CreateParameter("p_MiddleName", adVarChar, adParamInput, 200, oPatient.Name.Middle_Name)
        .Parameters.Append oParam
        Set oParam = .CreateParameter("p_SSN", adVarChar, adParamInput, 30, oPatient.SSN)
        .Parameters.Append oParam
        Set oParam = .CreateParameter("p_DOB", adDate, adParamInput, 50, CDate(oPatient.DateOfBirth))
        .Parameters.Append oParam
        Set oParam = .CreateParameter("p_Gender", adVarChar, adParamInput, 100, oPatient.SEX)
        .Parameters.Append oParam
        Set oParam = .CreateParameter("p_FMP", adVarChar, adParamInput, 10, oPatient.FMP)
        .Parameters.Append oParam
        Set oParam = .CreateParameter("p_Sponsor_SSN", adVarChar, adParamInput, 30, oPatient.SponsorSSN)
        .Parameters.Append oParam
        Set oParam = .CreateParameter("p_MothersMaidenName", adVarChar, adParamInput, 300, oPatient.MothersMaidenName)
        .Parameters.Append oParam
        Set oParam = .CreateParameter("p_Marital_Status", adVarChar, adParamInput, 200, oPatient.MaritalStatus)
        .Parameters.Append oParam
        Set oParam = .CreateParameter("p_Race", adVarChar, adParamInput, 200, oPatient.Race)
        .Parameters.Append oParam
        Set oParam = .CreateParameter("p_Religion", adVarChar, adParamInput, 200, oPatient.Religion)
        .Parameters.Append oParam
        Set oParam = .CreateParameter("p_Address_1", adVarChar, adParamInput, 500, oPatient.HomeAddress.Street1)
        .Parameters.Append oParam
        Set oParam = .CreateParameter("p_Address_2", adVarChar, adParamInput, 500, oPatient.HomeAddress.Street2)
        .Parameters.Append oParam
        Set oParam = .CreateParameter("p_City", adVarChar, adParamInput, 100, oPatient.HomeAddress.City)
        .Parameters.Append oParam
        Set oParam = .CreateParameter("p_State", adVarChar, adParamInput, 100, oPatient.HomeAddress.State)
        .Parameters.Append oParam
        Set oParam = .CreateParameter("p_Zip_Code", adVarChar, adParamInput, 100, oPatient.HomeAddress.Zip)
        .Parameters.Append oParam
        Set oParam = .CreateParameter("p_Phone", adVarChar, adParamInput, 100, "(" & oPatient.HomePhone.AreaCode & ")" & oPatient.HomePhone.Phone)
        .Parameters.Append oParam
        Set oParam = .CreateParameter("p_Office_Phone", adVarChar, adParamInput, 100, "(" & oPatient.WorkPhone.AreaCode & ")" & oPatient.WorkPhone.Phone)
        .Parameters.Append oParam
        Set oParam = .CreateParameter("p_Application", adVarChar, adParamInput, 500, "CHCSII")
        .Parameters.Append oParam
        Set oParam = .CreateParameter("p_ApplicationVersion", adVarChar, adParamInput, 500, "1.0")
        .Parameters.Append oParam

        Set oParam = .CreateParameter("Local_Unit_Number", adNumeric, adParamOutput, 20)
        .Parameters.Append oParam
    End With

    oDAS.ExecuteSP oCmd

    lUnitNumber = oCmd.Parameters("Local_Unit_Number").Value


    '--- After the new patient is created
    oPatient.EnterpiseUnitNumber = CStr(oPatient.UnitNumber)
    oPatient.UnitNumber = lUnitNumber
    
    Set oState = oPatient
    oState.State = osSaved
    CreateNewPatient = True
    
    GoTo CleanUpAndExit
ErrHandler:
    Err.Raise Err.Number, "MapperICD.CreateNewPatient", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
    Set oParam = Nothing
    Set oDAS = Nothing
    Set oCmd = Nothing
    Set oConn = Nothing
End Function

