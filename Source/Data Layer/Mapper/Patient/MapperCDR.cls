VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MapperCDR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements DL_Interfaces.iMapperPatient

Private mobjMapperGeneric As DL_Interfaces.iMapperPatient
Private mobjMapper As MapperGeneric
Private mConn      As CHCSII_CONN.Conn

Private Sub Class_Initialize()
Dim oSQLFix As iSQLFix
Dim oPicklist As iPickList

On Error GoTo ErrHandler

  Set mobjMapperGeneric = New MapperGeneric
  Set mConn = New CHCSII_CONN.Conn
  
  Set oSQLFix = New SQLOracle
  Set oPicklist = New PickListCDR
  
  Set mobjMapper = mobjMapperGeneric
  mobjMapper.BackEndDB = CDR
  Set mobjMapper.SQLFix = oSQLFix
  Set mobjMapper.PickList = oPicklist
  
  Exit Sub
  
ErrHandler:
  Err.Raise Err.Number, "DM_Patient.MapperCDR.Initialize", Err.Description, Err.HelpFile, Err.HelpContext

End Sub

Private Sub Class_Terminate()
    Set mConn = Nothing
    Set mobjMapper = Nothing
    Set mobjMapperGeneric = Nothing
End Sub

Private Function iMapperPatient_GetRankNCID(ByVal sRankCode As String, ByVal sBranch As String) As String
    Err.Raise Err.Number, "MapperCDR.iMapperPatient_GetRankNCID", "GetRankNCID is not allowed in CDR mode.", Err.HelpFile, Err.HelpContext
End Function

Private Sub iMapperPatient_Init(Support As DL_Support.iParameters, Cache As DL_Support.CacheMapper)
    mobjMapperGeneric.Init Support, Cache
End Sub


Private Function iMapperPatient_Retrieve(Params As DL_Support.iParameters) As DTOs.Patients

Dim oPatParams As PatientSearchParam
Dim oPatients As DTOs.Patients

On Error GoTo ErrHandler

    Set oPatParams = Params
    
    If oPatParams.SearchCHCS Then
        Set iMapperPatient_Retrieve = GetPatients_CHCS(oPatParams)
    Else
        Set oPatients = mobjMapperGeneric.Retrieve(Params)
        
        '--- Only load the insurance info if the retrieve is for one patient
        If LenB(oPatParams.UnitNumber) > 0 And Val(oPatParams.UnitNumber) <> 0 Then
            Set oPatients(1).Insurances = GetInsurances(oPatParams.UnitNumber)
        End If
        
        Set iMapperPatient_Retrieve = oPatients
    End If
    
    GoTo CleanUpAndExit
ErrHandler:
    Err.Raise Err.Number, "MapperCDR.iMapperPatient_Retrieve", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
    Set oPatParams = Nothing
    Set oPatients = Nothing
End Function

Private Function iMapperPatient_RetrieveRepresentation(ByRef Params As DL_Support.iParameters) As String
    On Error GoTo ErrHandler
    
    iMapperPatient_RetrieveRepresentation = mobjMapperGeneric.RetrieveRepresentation(Params)
    
    GoTo CleanUpAndExit
ErrHandler:
    Err.Raise Err.Number, "MapperCDR.iMapperPatient_RetrieveRepresentation", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
End Function

Private Function iMapperPatient_RetrieveSelectStrings(ByRef Params As DL_Support.iParameters) As DTOs.PatientSelectStrings
On Error GoTo ErrHandler
    
    Set iMapperPatient_RetrieveSelectStrings = mobjMapperGeneric.RetrieveSelectStrings(Params)
    
    GoTo CleanUpAndExit
ErrHandler:
    Err.Raise Err.Number, "MapperCDR.iMapperPatient_RetrieveSelectStrings", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:


End Function


Private Function iMapperPatient_Save(ByRef oPatient As DTOs.Patient) As Boolean
    On Error GoTo ErrHandler
    
    iMapperPatient_Save = SaveToCDR(oPatient)
    
    GoTo CleanUpAndExit
ErrHandler:
    Err.Raise Err.Number, "MapperCDR.iMapperPatient_TransferCHCSPatient", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:

End Function

Private Function iMapperPatient_TransferCHCSPatient(ByRef Params As DL_Support.iParameters) As DTOs.PatientTransferResult

Dim sUnitNumber As String
Dim objResult As DTOs.PatientTransferResult
Dim oPatParams As PatientSearchParam

    On Error GoTo ErrHandler
    
    Set oPatParams = Params
    Set objResult = New DTOs.PatientTransferResult
    sUnitNumber = CheckIfNeedsLoaded(oPatParams.MedicalRecordNumber, oPatParams.FacilityNCID)
    
    If Val(sUnitNumber) = 0 Then
        sUnitNumber = Transfer_CHCS_Patient(oPatParams.MedicalRecordNumber, oPatParams.FacilityNCID)
        If IsNumeric(sUnitNumber) Then
            objResult.TransferSuccess = True
        Else
            objResult.TransferSuccess = False
        End If
        
        objResult.PatientInCDR = False
    Else
        objResult.PatientInCDR = True
        objResult.TransferSuccess = True
    End If
    
    objResult.UnitNumber = sUnitNumber
    
    Set iMapperPatient_TransferCHCSPatient = objResult
    
    GoTo CleanUpAndExit
    
ErrHandler:
    Err.Raise Err.Number, "MapperCDR.iMapperPatient_TransferCHCSPatient", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
    Set objResult = Nothing
    Set oPatParams = Nothing
End Function

Private Function GetPatients_CHCS(ByRef objPatientParams As PatientSearchParam) As DTOs.Patients
Dim oResultRS               As ADODB.Recordset
Dim objDASsp                As ICHCSII_DataAccess.ICHCSII_DAS
Dim objCmd                  As CHCSII_Command.CHCSII_Cmd
Dim objPatientSearchResults As CHCS_Utils_Helper.PatientSearchResults
Dim objSearchResultRecord   As CHCS_Utils_Helper.SearchResultRecord
Dim sName                   As String
Dim objPatient              As DTOs.Patient
Dim objPatients             As DTOs.Patients
Dim sResults                As String
Dim i                       As Integer

    On Error GoTo ErrHandler
       
    Set objDASsp = mConn.CHCSII_DAS
    Set objCmd = New CHCSII_Command.CHCSII_Cmd
    
    objCmd.SPName = "CDRPLUS.PT_SRCH_I_OR_ADD_PT_II"
    
    With objPatientParams
        If LenB(.LastName) > 0 Then
            sName = .LastName
        End If
        
        If LenB(.FirstName) > 0 Then
            sName = sName & "," & .FirstName
        End If
        
        objCmd.AddParam False, False, "textString", .MaximumNumberOfRecords
        objCmd.AddParam False, False, "textString", .LastRecordID
        objCmd.AddParam False, False, "textString", .MedicalRecordNumber
        objCmd.AddParam False, False, "textString", sName
        objCmd.AddParam False, False, "textString", .SSN
        objCmd.AddParam False, False, "textString", .SponsorSSN
        objCmd.AddParam False, False, "textString", .FMP
        objCmd.AddParam False, False, "textString", .DOB
        objCmd.AddParam False, False, "textString", .FacilityNCID
        objCmd.AddParam False, False, "textString", "CDRPLUS.PT_SRCH_RQST_Q" 'queue request name
        objCmd.AddParam False, False, "textString", "CDRPLUS.PT_SRCH_RSLT_Q" 'queue result name
                
        '--- Output Parameter
        objCmd.AddParam True, True, "textString", "initialize"
    End With
    
    Set oResultRS = objDASsp.OpenRecordsetSP(objCmd)
    
    sResults = oResultRS.Fields(0)
    If LenB(sResults) > 0 Then
        Set objPatientSearchResults = New CHCS_Utils_Helper.PatientSearchResults
        objPatientSearchResults.Deserialize sResults
        
        With objPatientSearchResults
            If LenB(.ErrorText) > 0 Then
                Err.Raise vbObjectError + Err.Number, "CHCSRetrieveCDR.iCHCSRetrieve_SearchCHCS", .ErrorText
                GoTo CleanUpAndExit
            Else
                Set objPatients = New DTOs.Patients
                
                For i = 1 To .Results.Count
                    Set objSearchResultRecord = .Results.Item(i)
                    Set objPatient = Nothing
                    Set objPatient = New DTOs.Patient
                    
                    With objPatient
                        '--- Name
                        sName = objSearchResultRecord.Name
                        
                        If InStr(1, sName, ",") = 0 Then
                            .Name.Last_Name = sName
                        Else
                            .Name.Last_Name = Mid$(sName, 1, InStr(1, sName, ",") - 1)
                            sName = Mid$(sName, InStr(1, sName, ",") + 1)
                            If InStr(1, sName, " ") > 0 Then
                                .Name.First_Name = Trim$(Mid$(sName, 1, InStr(1, sName, " ") - 1))
                                .Name.Middle_Name = Trim$(Mid$(sName, InStr(1, sName, " ") + 1))
                            Else
                                .Name.First_Name = sName
                                .Name.Middle_Name = vbNullString
                            End If
                        End If
                        
                        '--- Other
                        .FMP = objSearchResultRecord.FMP
                        .SponsorSSN = objSearchResultRecord.SponsorSSN
                        .MedicalRecordNumber = objSearchResultRecord.PatientIEN
                        .UnitNumber = objSearchResultRecord.PatientIEN
                        If LenB(objSearchResultRecord.DOB) > 0 Then
                            .DateOfBirth = CDate(objSearchResultRecord.DOB)
                        End If
                    End With
                    
                    objPatients.Add objPatient, objPatient.UnitNumber
                Next
                
                objPatients.LastRecordID = .LastRecordID
            End If
        End With
    End If

    Set GetPatients_CHCS = objPatients
    
    GoTo CleanUpAndExit
ErrHandler:
    Err.Raise Err.Number, "MapperCDR.GetPatients_CHCS", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
    Set oResultRS = Nothing
    Set objDASsp = Nothing
    Set objCmd = Nothing
    Set objPatientSearchResults = Nothing
    Set objSearchResultRecord = Nothing

End Function






Private Function CheckIfNeedsLoaded(ByVal sPatientIEN As String, ByVal sFacilityNCID As String) As String

Dim objResults As ADODB.Recordset
Dim sSQL       As String

    On Error GoTo Error_Handler
    
    sSQL = "select unit_number from mmi_site " _
        & " where facility_cid = " & sFacilityNCID _
        & " and medical_record_number = '" & sPatientIEN & "'"
                   
    Set objResults = mConn.CHCSII_DAS(CDR).OpenRecordset(sSQL)
                   
    If objResults.RecordCount > 0 Then
        CheckIfNeedsLoaded = objResults("UNIT_NUMBER")
    Else
        CheckIfNeedsLoaded = vbNullString
    End If
    
    GoTo CleanUpAndExit

Error_Handler:
    Err.Raise Err.Number, "MapperCDR.CheckIfNeedsLoaded", Err.Description
CleanUpAndExit:
    Set objResults = Nothing
End Function




Private Function Transfer_CHCS_Patient(ByVal sPatientIEN As String, ByVal sFacilityNCID As String) As String

Dim objDASsp                As ICHCSII_DataAccess.ICHCSII_DAS
Dim objCmd                  As CHCSII_Command.CHCSII_Cmd
Dim objResultRS             As ADODB.Recordset
Dim sUnitNumber             As String

    On Error GoTo ErrHandler
       
    Set objDASsp = mConn.CHCSII_DAS
    Set objCmd = New CHCSII_Command.CHCSII_Cmd
    
    objCmd.SPName = "CDRPLUS.PT_SRCH_I_OR_ADD_PT_II"
    
    objCmd.AddParam False, False, "textString", vbNullString
    objCmd.AddParam False, False, "textString", vbNullString
    objCmd.AddParam False, False, "textString", sPatientIEN
    objCmd.AddParam False, False, "textString", vbNullString
    objCmd.AddParam False, False, "textString", vbNullString
    objCmd.AddParam False, False, "textString", vbNullString
    objCmd.AddParam False, False, "textString", vbNullString
    objCmd.AddParam False, False, "textString", vbNullString
    objCmd.AddParam False, False, "textString", sFacilityNCID
    objCmd.AddParam False, False, "textString", "CDRPLUS.PT_ADD_RQST_Q"
    objCmd.AddParam False, False, "textString", "CDRPLUS.PT_ADD_RSLT_Q"

    '--- This Parameter is used for output
    objCmd.AddParam True, True, "textString", "initialize"

    Set objResultRS = objDASsp.OpenRecordsetSP(objCmd)
    
    '--- stored procedure timed out before getting search results
    If InStr(1, objResultRS.Fields(0), "25228") = 0 Then
        sUnitNumber = objResultRS.Fields(0)
    Else
        Err.Raise 25228, "MapperCDR.Transfer_CHCS_Patient", "Add Patient Timed Out Before Returning Unit Number."
        GoTo CleanUpAndExit
    End If
    
    Transfer_CHCS_Patient = sUnitNumber
    
    GoTo CleanUpAndExit
    
ErrHandler:
    Err.Raise Err.Number, "MapperCDR.Transfer_CHCS_Patient", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
    Set objResultRS = Nothing
    Set objDASsp = Nothing
    Set objCmd = Nothing
End Function



Private Function SaveToCDR(ByRef oPatient As DTOs.Patient) As Boolean
    
    Dim objMMI           As MmiLib.MmiDoc
    Dim objMember        As MmiLib.Member
    Dim objName          As MmiLib.Name
    Dim objFacility      As MmiLib.Facility
    Dim objPhone         As MmiLib.Phone
    Dim lResult          As Long
    Dim bResult          As Boolean
    Dim oGenData         As DTOs.PatientGenericData
    Dim bHomePhoneExists As Boolean
    Dim bWorkPhoneExists As Boolean
    
    On Error GoTo CleanUpAndExit
    

    '*-- Check for UnitNumber match
    '*-- To Do: Add code for record validation with SSN, LastName and DOB.
    Set objMMI = New MmiDoc
    Set objMember = New Member
    objMember.UnitNumber = oPatient.UnitNumber


    objMMI.DoSearch objMember

    '*-- Turn errors off for GET(), check the ERR return and then turn
    '*-- errors back on.
    On Error Resume Next
    bResult = objMMI.Get(True)
    On Error GoTo CleanUpAndExit
    
    If Err = 1000 Or bResult = False Then
        '*-- Not found so add a new record...
        Set objMember = New Member
        With objMember
            .BirthDate = oPatient.DateOfBirth
            .SSN = ReturnNumbers(oPatient.SSN)
            If LenB(.SSN) > 0 Then
                .SSNCorrect = "U"
            End If

            .Name = NCID_CURRENT_NAME
            .Name.Last = oPatient.Name.Last_Name
            .Name.First = oPatient.Name.First_Name
            .Name.Middle = oPatient.Name.Middle_Name

            .GenericId = NCID_SPONSOR_SSN
            .GenericId.Type = NCID_SPONSOR_SSN
            .GenericId.Value = ReturnNumbers(oPatient.SponsorSSN)
            .Add
            
            .GenericId = NCID_PATIENT_FMP
            .GenericId.Type = NCID_PATIENT_FMP
            .GenericId.Value = oPatient.FMP
            .Add
            
            .Facility = oPatient.FacilityNCID
            .Facility.Type = oPatient.FacilityNCID
            .Facility.MedRecNum = oPatient.MedicalRecordNumber
            lResult = .Add
            If lResult = 0 Then
                GoTo CleanUpAndExit
            End If
        End With

        bResult = objMMI.WriteUpdate
        If bResult = False Then
            Err.Raise vbObjectError, vbNullString, "WriteUpdate failed."
        Else
            SaveToCDR = True
        End If

    ElseIf Err = 0 And bResult = True Then
        '*-- Patient found, now update object and write back.
        Set objMember = objMMI.Member
        If Not objMember Is Nothing Then
            objMember.GotoFirst
    
            With objMember
                'Need to add facility info for this patient
                .Facility = oPatient.FacilityNCID
                .Facility.Type = oPatient.FacilityNCID
                .Facility.MedRecNum = oPatient.MedicalRecordNumber
    
                If Not oPatient.HomeAddress Is Nothing Then
                    .Address = NCID_HOME_ADDRESS
                    .Address.Street = oPatient.HomeAddress.Street1
                    .Address.Street2 = oPatient.HomeAddress.Street2
                    .Address.City = oPatient.HomeAddress.City
                    .Address.State = oPatient.HomeAddress.State
                    .Address.PostCode = oPatient.HomeAddress.Zip
                End If
                
                '---Handle and Save Phone Info
                If Not .Phone Is Nothing Then
                    .Phone.GotoFirst
                End If
    
                Set objPhone = .Phone
                bHomePhoneExists = False
                bWorkPhoneExists = False
    
                If Not objPhone Is Nothing Then     '--make sure a "phone" object exists--
                  Do
                    'only show phone numbers that are different than listed
                    If objPhone.Type = NCID_HOME_PHONE Then
                        If Not oPatient.HomePhone Is Nothing Then
                            objPhone.AreaCode = oPatient.HomePhone.AreaCode
                            objPhone.LocalNumber = oPatient.HomePhone.Phone
                            bHomePhoneExists = True
                        End If
                    ElseIf objPhone.Type = NCID_WORK_PHONE Then
                        If Not oPatient.WorkPhone Is Nothing Then
                            objPhone.AreaCode = oPatient.WorkPhone.AreaCode
                            objPhone.LocalNumber = oPatient.WorkPhone.Phone
                            bWorkPhoneExists = True
                        End If
                    End If
            
                    '---go to the next object---
                    objPhone.GotoNext

                    '---point at this object---
                    Set objPhone = .Phone
                  Loop Until objPhone.EOF
                End If
                
                '---If phone number objects don't exist, create them---
                If bHomePhoneExists = False And Not oPatient.HomePhone Is Nothing Then
                    .Phone = NCID_HOME_PHONE
                    .Phone.AreaCode = oPatient.HomePhone.AreaCode
                    .Phone.LocalNumber = oPatient.HomePhone.Phone
                    bHomePhoneExists = True
                End If
                If bWorkPhoneExists = False And Not oPatient.WorkPhone Is Nothing Then
                    .Phone = NCID_WORK_PHONE
                    .Phone.AreaCode = oPatient.WorkPhone.AreaCode
                    .Phone.LocalNumber = oPatient.WorkPhone.Phone
                    bWorkPhoneExists = True
                End If
                
                .SEX = oPatient.SEX
                .Race = oPatient.Race
                .Religion = oPatient.Religion
                .UnitNumber = oPatient.UnitNumber
                .SSN = oPatient.SSN
                .MomMaidenName = oPatient.MothersMaidenName
                .MaritalStatus = oPatient.MaritalStatus
                .BirthDate = oPatient.DateOfBirth

                '--- MMI_GENERIC_ID
                If oPatient.FMP <> vbNullString Then
                    UpdateMMIGenID objMember, NCID_FMP, oPatient.FMP, "0"
                End If
                If oPatient.SponsorSSN <> vbNullString Then
                    UpdateMMIGenID objMember, NCID_SPONSOR_SSN, oPatient.SponsorSSN, "0"
                End If
                If oPatient.VIP Then
                    UpdateMMIGenID objMember, NCID_VIP, "Y", "0"
                Else
                    UpdateMMIGenID objMember, NCID_VIP, "N", "0"
                End If
                If oPatient.FMPSponsorSSN <> vbNullString Then
                    UpdateMMIGenID objMember, NCID_FMP_SPONSOR_SSN, oPatient.FMPSponsorSSN, "0"
                End If
                
                If oPatient.SponsorSrhString = vbNullString Then
                    oPatient.SponsorSrhString = Left$(oPatient.Name.Last_Name, 1) & Right$(oPatient.SponsorSSN, 4)
                End If
                If oPatient.SponsorSrhString <> vbNullString Then
                    UpdateMMIGenID objMember, NCID_FIRST_LETTER_LAST_NAME_LAST_FOUR_SPONSOR_SSN, oPatient.SponsorSrhString, "0"
                End If
                
                If oPatient.PatientSrhString = vbNullString Then
                    oPatient.PatientSrhString = Left$(oPatient.Name.Last_Name, 1) & Right$(oPatient.SSN, 4)
                End If
                If oPatient.PatientSrhString <> vbNullString Then
                    UpdateMMIGenID objMember, NCID_FIRST_LETTER_LAST_NAME_LAST_FOUR_PATIENT_SSN, oPatient.PatientSrhString, "0"
                End If
                
                '--- MMI_GENERICDATA
                If Not oPatient.GenericDatas Is Nothing Then
                    For Each oGenData In oPatient.GenericDatas
                        UpdateMMIGenData objMember, oGenData.TypeNcid, oGenData.ValueText, oGenData.ValueRsid, oGenData.ValueNcid, oGenData.ErrorText
                    Next oGenData
                End If

            End With

            bResult = objMMI.WriteUpdate
            If bResult = False Then
                Err.Raise Err.Number, "MapperCDR.SaveToCDR", Err.Description, Err.HelpFile, Err.HelpContext
            Else
                SaveToCDR = True
            End If
        End If
    End If
    
    GoTo CleanUpAndExit
    
ErrHandler:
    Err.Raise Err.Number, "MapperCDR.SaveToCDR", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
    Set objMMI = Nothing
    Set objMember = Nothing
    Set objName = Nothing
    Set oGenData = Nothing
    Set objFacility = Nothing
    Exit Function
    Resume
End Function



'*-- Update the MMI GenericID collection. No WriteUpdate done.
Private Function UpdateMMIGenID(ByRef oMember As Member, _
                                ByVal sType As String, _
                                ByVal sValue As String, _
                                ByVal sAssigningAuthority As String) As Boolean
    Dim bFound As Boolean
    
    With oMember
        bFound = False
        .GenericId.GotoFirst
        Do
            If sType = .GenericId.Type Then
                '*-- Found matching GenericId, modify it
                .GenericId.Value = sValue
                .GenericId.AssigningAuthority = sAssigningAuthority
                bFound = True
                Exit Do
            Else
                .GenericId.GotoNext
            End If
        Loop While Not .GenericId.EOF
        If Not bFound Then
            '*-- Add new item...can only be done at GenericId.EOF.
            .GenericId = sType         '*-- Implicit create
            .GenericId.Type = sType
            .GenericId.Value = sValue
            .GenericId.AssigningAuthority = sAssigningAuthority
        End If
    End With
    UpdateMMIGenID = True
End Function


'*-- Update the MMI GenericData collection. No WriteUpdate done.
Private Function UpdateMMIGenData(ByRef oMember As Member, _
                                    ByVal sType_NCID As String, _
                                    ByVal sValue_Text As String, _
                                    ByVal sValue_RSID As String, _
                                    ByVal sValue_NCID As String, _
                                    ByVal sError As String) As Boolean
    Dim bFound As Boolean
    
    With oMember
        bFound = False
        If Not .GenericData Is Nothing Then
            .GenericData.GotoFirst
            Do
                If sType_NCID = .GenericData.Type Then
                    '*-- Found matching GenericData, modify it
                    .GenericData.ValueNcid = sValue_NCID
                    .GenericData.ValueRsid = sValue_RSID
                    .GenericData.ValueText = sValue_Text
                    .GenericData.ErrorText = sError
                    bFound = True
                    Exit Do
                Else
                    .GenericData.GotoNext
                End If
            Loop While Not .GenericData.EOF
        End If
        If Not bFound Then
            '*-- Add new item...can only be done at GenericData.EOF.
            .GenericData = sType_NCID         '*-- Implicit create
            .GenericData.Type = sType_NCID
            .GenericData.ValueNcid = sValue_NCID
            .GenericData.ValueRsid = sValue_RSID
            .GenericData.ValueText = sValue_Text
            .GenericData.ErrorText = sError
        End If
    End With
    UpdateMMIGenData = True
End Function


Private Function GetInsurances(ByVal sUnitNumber As String) As DTOs.Insurances

Dim objDASsp                As ICHCSII_DataAccess.ICHCSII_DAS
Dim objCmd                  As CHCSII_Command.CHCSII_Cmd
Dim objResultRS             As ADODB.Recordset

    On Error GoTo ErrHandler
       
    Set objDASsp = mConn.CHCSII_DAS(CDR)
    Set objCmd = New CHCSII_Command.CHCSII_Cmd
    
    objCmd.SPName = "DEMOGRAPHIC_PERFORMANCE_PKG.prc_Load_Insurance"
    
    objCmd.AddParam False, False, "textString", sUnitNumber

    '--- This Parameter is used for output
    objCmd.AddParam True, True, vbNullString, vbNullString

    Set objResultRS = objDASsp.OpenRecordsetSP(objCmd)
    
    LoadInsuranceObject GetInsurances, objResultRS
    
    GoTo CleanUpAndExit
    
ErrHandler:
    Err.Raise Err.Number, "MapperCDR.GetInsurances", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
    Set objResultRS = Nothing
    Set objDASsp = Nothing
    Set objCmd = Nothing
End Function


