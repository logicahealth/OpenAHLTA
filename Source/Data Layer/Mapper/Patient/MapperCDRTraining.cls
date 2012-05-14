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
Attribute VB_Name = "MapperCDRTraining"
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
  Set oSQLFix = New SQLSQLServer
  Set oPicklist = New PickListGEMS

  Set objMapper = mobjMapperGeneric
  objMapper.BackEndDB = ITT
  Set objMapper.SQLFix = oSQLFix
  Set objMapper.PickList = oPicklist
  Exit Sub
  
ErrHandler:
  Err.Raise Err.Number, "DM_Patient.MapperCDRTraining.Initialize", Err.Description, Err.HelpFile, Err.HelpContext


End Sub

Private Function iMapperPatient_GetRankNCID(ByVal sRankCode As String, ByVal sBranch As String) As String
    Err.Raise Err.Number, "MapperCDRTraining.iMapperPatient_GetRankNCID", "GetRankNCID is not allowed in Training mode.", Err.HelpFile, Err.HelpContext
End Function

Private Sub iMapperPatient_Init(Support As DL_Support.iParameters, Cache As DL_Support.CacheMapper)
    mobjMapperGeneric.Init Support, Cache
End Sub

Private Function iMapperPatient_Retrieve(Params As DL_Support.iParameters) As DTOs.Patients
Dim oPatParams As PatientSearchParam
Dim oPatients As DTOs.Patients

On Error GoTo ErrHandler

    Set oPatParams = Params
    
    Set oPatients = mobjMapperGeneric.Retrieve(Params)
    
     '--- Only load the insurance info if the retrieve is for one patient
    If LenB(oPatParams.UnitNumber) > 0 And Val(oPatParams.UnitNumber) <> 0 Then
        Set oPatients(1).Insurances = GetInsurances(oPatParams.UnitNumber)
    End If
   
    Set iMapperPatient_Retrieve = oPatients
    
    GoTo CleanUpAndExit
ErrHandler:
    Err.Raise Err.Number, "MapperCDRTraining.iMapperPatient_Retrieve", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
    Set oPatParams = Nothing
    Set oPatients = Nothing
End Function

Private Function iMapperPatient_RetrieveSelectStrings(Params As DL_Support.iParameters) As DTOs.PatientSelectStrings
On Error GoTo ErrHandler
    
    Set iMapperPatient_RetrieveSelectStrings = mobjMapperGeneric.RetrieveSelectStrings(Params)
    
    GoTo CleanUpAndExit
ErrHandler:
    Err.Raise Err.Number, "MapperCDRTraining.iMapperPatient_RetrieveSelectStrings", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:

End Function


Private Function iMapperPatient_RetrieveRepresentation(ByRef Params As DL_Support.iParameters) As String
    On Error GoTo ErrHandler
    
    iMapperPatient_RetrieveRepresentation = mobjMapperGeneric.RetrieveRepresentation(Params)
    
    GoTo CleanUpAndExit
ErrHandler:
    Err.Raise Err.Number, "MapperCDRTraining.iMapperPatient_RetrieveRepresentation", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
End Function


Private Function iMapperPatient_Save(Patient As DTOs.Patient) As Boolean
On Error GoTo ErrHandler
    
    iMapperPatient_Save = mobjMapperGeneric.Save(Patient)
    
    GoTo CleanUpAndExit
ErrHandler:
    Err.Raise Err.Number, "MapperCDRTraining.iMapperPatient_RetrieveSelectStrings", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:

End Function

Private Function iMapperPatient_TransferCHCSPatient(ByRef Params As DL_Support.iParameters) As DTOs.PatientTransferResult
    Err.Raise Err.Number, "MapperCDRTraining.iMapperPatient_RetrieveSelectStrings", "TransferCHCSPatient is not allowed in ITT mode.", Err.HelpFile, Err.HelpContext
End Function



Private Function GetInsurances(ByVal sUnitNumber As String) As DTOs.Insurances

Dim objRS      As ADODB.Recordset
Dim sSQL       As String
Dim mConn      As CHCSII_CONN.Conn

    On Error GoTo ErrHandler
    
    Set mConn = New CHCSII_CONN.Conn
    
    sSQL = "SELECT  "
    sSQL = sSQL & "     INSCOMPANYNAME       InsuranceName," & vbNewLine
    sSQL = sSQL & "     INSCOMPANYPHONE      Phone," & vbNewLine
    sSQL = sSQL & "     ADDRESS1             Address," & vbNewLine
    sSQL = sSQL & "     GROUPNAME            InsuredGroupName," & vbNewLine
    sSQL = sSQL & "     GROUPNUMBER          InsuredGroupNumber," & vbNewLine
    sSQL = sSQL & "     SUBSCRIBERNAME       InsuredGroupEmpName," & vbNewLine
    sSQL = sSQL & "     STARTDATE            PlanEffectiveTime," & vbNewLine
    sSQL = sSQL & "     STOPDATE             PlanExpirationTime," & vbNewLine
    sSQL = sSQL & "     INSIDNUMBER          PolicyNumber," & vbNewLine
    sSQL = sSQL & "     RELATIONTOSPONSOR    ToPatientRel" & vbNewLine
    sSQL = sSQL & "  from patient_insurance " & vbNewLine
    sSQL = sSQL & "  where unit_number =  " & sUnitNumber & vbNewLine

    Set objRS = mConn.CHCSII_DAS(ITT).OpenRecordset(sSQL)
    
    LoadInsuranceObject GetInsurances, objRS
    
    GoTo CleanUpAndExit
    
ErrHandler:
    Err.Raise Err.Number, "MapperCDRTraining.GetInsurances", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
    Set objRS = Nothing
End Function



