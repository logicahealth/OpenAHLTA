VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MapperUnitTesting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements DL_Interfaces.iMapperPatient

Private Function iMapperPatient_GetRankNCID(ByVal sRankCode As String, ByVal sBranch As String) As String
    Err.Raise Err.Number, "MapperUnitTesting.iMapperPatient_GetRankNCID", "GetRankNCID is not allowed in UnitTesting mode.", Err.HelpFile, Err.HelpContext
End Function

Private Sub iMapperPatient_Init(ByRef Support As DL_Support.iParameters, ByRef Cache As DL_Support.CacheMapper)

End Sub

Private Function iMapperPatient_Retrieve(ByRef Params As DL_Support.iParameters) As DTOs.Patients
Dim objPatientParams As PatientSearchParam
Dim oPat As DTOs.Patient
    On Error GoTo ErrHandler
    
    Set iMapperPatient_Retrieve = GetPatients()
    
    Set objPatientParams = Params
    If LenB(objPatientParams.UnitNumber) > 0 Then
        For Each oPat In iMapperPatient_Retrieve
            If oPat.UnitNumber <> objPatientParams.UnitNumber Then
                Exit For
            End If
        Next
        
        Set iMapperPatient_Retrieve = New DTOs.Patients
        
        iMapperPatient_Retrieve.Add oPat, oPat.UnitNumber
        
        Set oPat = Nothing
        Set objPatientParams = Nothing
    End If
    
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "MapperUnitTesting.iMapperPatient_Save", Err.Description, Err.HelpFile, Err.HelpContext
End Function

Private Function iMapperPatient_RetrieveSelectStrings(ByRef Params As DL_Support.iParameters) As DTOs.PatientSelectStrings
    Dim objSelectStrParams As SelectStringSearchParam
    
    On Error GoTo ErrHandler
    
    Set objSelectStrParams = Params
    
    Set iMapperPatient_RetrieveSelectStrings = GetSelectStrings(objSelectStrParams.SelectStringCategory)
    
    GoTo CleanUpAndExit
ErrHandler:
    Err.Raise Err.Number, "MapperUnitTesting.iMapperPatient_RetrieveSelectStrings", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
    Set objSelectStrParams = Nothing
End Function

Private Function iMapperPatient_RetrieveRepresentation(ByRef Params As DL_Support.iParameters) As String
Dim oParams As SelectStringSearchParam
Dim oSelectStrings As DTOs.PatientSelectStrings
Dim sNCID As String
Dim i As Integer

    On Error GoTo ErrHandler
    
    iMapperPatient_RetrieveRepresentation = vbNullString
    Set oParams = Params
    sNCID = oParams.NCID
    
    Set oSelectStrings = GetSelectStrings(oParams.SelectStringCategory)
    
    For i = 1 To oSelectStrings.Count
        If sNCID = oSelectStrings.Item(i).ID Then
            iMapperPatient_RetrieveRepresentation = oSelectStrings.Item(i).Value
            Exit For
        End If
    Next i
    
    GoTo CleanUpAndExit
ErrHandler:
    Err.Raise Err.Number, "MapperUnitTesting.iMapperPatient_RetrieveRepresentation", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
    Set oParams = Nothing
    Set oSelectStrings = Nothing
End Function

Private Function iMapperPatient_Save(ByRef Patient As DTOs.Patient) As Boolean
    Dim oPatients As DTOs.Patients
    Dim i As Integer

    On Error GoTo ErrHandler
    iMapperPatient_Save = False
    
    Set oPatients = GetPatients()
    
    For i = 1 To oPatients.Count
        If oPatients.Item(i).UnitNumber = Patient.UnitNumber Then
            oPatients.Remove i
            Exit For
        End If
    Next i
    
    oPatients.Add Patient, Patient.UnitNumber
    
    SavePatients oPatients

    iMapperPatient_Save = True
    
    Set oPatients = Nothing
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "MapperUnitTesting.iMapperPatient_Save", Err.Description, Err.HelpFile, Err.HelpContext
End Function

Private Function iMapperPatient_TransferCHCSPatient(ByRef Params As DL_Support.iParameters) As DTOs.PatientTransferResult

End Function


Private Sub SavePatients(ByRef Patients As DTOs.Patients)
    Dim Writer As ObjectPersist.WriterXML
    Dim Persist As iObjectPersist
    
    On Error GoTo ErrHandler
    
    Set Writer = New WriterXML
    Set Persist = Patients
    Persist.Persist Writer

    Writer.Save App.Path & "\PatientsUnitTest.xml", False
    
    GoTo CleanUp
    
ErrHandler:
    Err.Raise Err.Number, "MapperUnitTesting.SavePatients", Err.Description, Err.HelpFile, Err.HelpContext
CleanUp:
    Set Writer = Nothing
    Set Persist = Nothing
End Sub



Private Function GetPatients() As DTOs.Patients
Dim Reader As ObjectPersist.ReaderXML
Dim Persist As iObjectPersist
Dim iReader As iReader

    On Error GoTo ErrHandler
    
    Set Persist = New DTOs.Patients
    Set Reader = New ObjectPersist.ReaderXML
    Set iReader = Reader
    iReader.DocumentVersion = Persist.Version
    Reader.Load App.Path & "\PatientsUnitTest.xml"

    Persist.Restore Reader
    
    Set GetPatients = Persist
    
    GoTo CleanUp
    
ErrHandler:
    Err.Raise Err.Number, "MapperUnitTesting.GetPatients", Err.Description, Err.HelpFile, Err.HelpContext
CleanUp:
    Set Reader = Nothing
    Set Persist = Nothing
    Set iReader = Nothing
End Function

'************************************
'--- This is for future reference
'************************************
'Private Sub SaveSelectStrings(ByRef oSelectStrings As DTOs.PatientSelectStrings, ByVal eCat As DTOs.enumSelectStringCategory)
'    Dim Writer As ObjectPersist.WriterXML
'    Dim Persist As iObjectPersist
'
'    On Error GoTo ErrHandler
'
'    Set Writer = New WriterXML
'    Set Persist = oSelectStrings
'    Persist.Persist Writer
'
'    Writer.Save App.Path & "\SelectString" & CStr(eCat) & ".xml", False
'
'    GoTo CleanUp
'
'ErrHandler:
'    Err.Raise Err.Number, "MapperUnitTesting.SaveSelectStrings", Err.Description, Err.HelpFile, Err.HelpContext
'CleanUp:
'    Set Writer = Nothing
'    Set Persist = Nothing
'End Sub




Private Function GetSelectStrings(ByVal eCategory As DTOs.enumSelectStringCategory) As DTOs.PatientSelectStrings
Dim Reader As ObjectPersist.ReaderXML
Dim Persist As iObjectPersist
Dim iReader As iReader

    On Error GoTo ErrHandler
    
    Set Persist = New DTOs.PatientSelectStrings
    Set Reader = New ObjectPersist.ReaderXML
    Set iReader = Reader
    iReader.DocumentVersion = Persist.Version
    Reader.Load App.Path & "\SelectString" & CStr(eCategory) & ".xml"

    Persist.Restore Reader
    
    Set GetSelectStrings = Persist
    
    GoTo CleanUp
    
ErrHandler:
    Err.Raise Err.Number, "MapperUnitTesting.GetPatients", Err.Description, Err.HelpFile, Err.HelpContext
CleanUp:
    Set Reader = Nothing
    Set Persist = Nothing
    Set iReader = Nothing
End Function

