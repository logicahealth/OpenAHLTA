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

Implements DL_Interfaces.iMapperPatient

Private mCache As CacheMapper
Private mConn As CHCSII_CONN.Conn


Private Function iMapperPatient_GetRankNCID(ByVal sRankCode As String, ByVal sBranch As String) As String
    Err.Raise Err.Number, "MapperPGUI.iMapperPatient_GetRankNCID", "GetRankNCID is not allowed in PGUI mode.", Err.HelpFile, Err.HelpContext
End Function

Private Sub iMapperPatient_Init(Support As DL_Support.iParameters, Cache As DL_Support.CacheMapper)
    Set mConn = Support.Item("DBConn").Value
    Set mCache = Cache
End Sub

Private Function iMapperPatient_Retrieve(Params As DL_Support.iParameters) As DTOs.Patients
Dim objParams As PatientSearchParam

    On Error GoTo ErrHandler
    
    Set objParams = Params
    
    Set iMapperPatient_Retrieve = GetPatients(objParams)
    
    GoTo CleanUpAndExit
ErrHandler:
    Err.Raise Err.Number, "MapperPGUI.iMapperPatient_Retrieve", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
    Set objParams = Nothing
End Function

Private Function iMapperPatient_RetrieveSelectStrings(Params As DL_Support.iParameters) As DTOs.PatientSelectStrings

End Function

Private Function iMapperPatient_RetrieveRepresentation(ByRef Params As DL_Support.iParameters) As String
    iMapperPatient_RetrieveRepresentation = vbNullString
End Function

Private Function iMapperPatient_Save(Patient As DTOs.Patient) As Boolean

End Function

Private Function iMapperPatient_TransferCHCSPatient(ByRef Params As DL_Support.iParameters) As DTOs.PatientTransferResult

End Function

Private Function GetPatients(ByRef objPatientParams As PatientSearchParam) As DTOs.Patients

Dim rsPatients As ADODB.Recordset
Dim rsSponsor As ADODB.Recordset
Dim objPatient As DTOs.Patient
Dim objSponsor As DTOs.Patient
Dim objPatients As DTOs.Patients
Dim objCHCS_Patient As CHCS_Interface.ICHCSPatient
Dim sName As String
Dim bBasicOnly As Boolean

    On Error GoTo ErrHandler
    
    bBasicOnly = objPatientParams.IncludeBasicOnly
    
    Set objCHCS_Patient = GetCHCSPatient
    
    If bBasicOnly Then
        With objPatientParams
            If LenB(.LastName) > 0 Then
                sName = .LastName
            End If
            
            If LenB(.FirstName) > 0 Then
                sName = sName & "," & .FirstName
            End If
            
            Set rsPatients = objCHCS_Patient.PatientSearch(sName, .SSN, .SponsorSSN, _
                    .DOB, .FMP, .LastRecordID, .MaximumNumberOfRecords)
        End With

    Else
        Set rsPatients = objCHCS_Patient.RetrieveDemographics(objPatientParams.MedicalRecordNumber)
    End If
    
    
    If rsPatients.RecordCount > 0 Then
      
      rsPatients.MoveFirst
      Set objPatients = New DTOs.Patients
      
      Do While Not rsPatients.EOF
      
          Set objPatient = Nothing
          Set objPatient = New DTOs.Patient
        
          If LoadPatientObject(objPatient, rsPatients, bBasicOnly) Then
        
            If Not bBasicOnly Then
            
              If objPatient.FMP = "20" Then
                Set objPatient.Sponsor = objPatient
              Else
                '--- Load Sponsor
                If LenB(objPatient.Sponsor_IEN) > 0 Then
                    Set rsSponsor = objCHCS_Patient.RetrieveDemographics(objPatient.Sponsor_IEN)
                    If rsSponsor.RecordCount > 0 Then
                        Set objSponsor = New DTOs.Patient
                         If LoadSponsorObject(objSponsor, rsSponsor) Then
                            Set objPatient.Sponsor = objSponsor
                         End If
                    End If
                End If
              End If
            End If
            
            '--- Add Patient
            If Not objPatients.Exists(objPatient.UnitNumber) Then
                objPatients.Add objPatient, objPatient.UnitNumber
            End If
            
          End If
          
          If rsPatients.Bookmark = objPatientParams.MaximumNumberOfRecords Then
             objPatients.LastRecordID = rsPatients("LAST_RECORD_ID")
          End If
            
          rsPatients.MoveNext
      Loop
                
    End If
   
    Set GetPatients = objPatients
    
    GoTo CleanUpAndExit
    
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.GetPatientsBySearchParams", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
    Set rsPatients = Nothing
    Set objPatient = Nothing
    Set objPatients = Nothing
End Function

'**********************************************************
' Load Patient Object with the CHCS data in Recordset
'**********************************************************

Private Function LoadPatientObject(ByRef objPatient As DTOs.Patient, ByVal oRS As ADODB.Recordset, ByVal bBasicOnly As Boolean) As Boolean
  
Dim sTemp As String
Dim sData() As String
Dim i As Integer
Dim oGenericData As DTOs.PatientGenericData
Dim objPregnancy As DTOs.Pregnancy

  On Error Resume Next
    
  With objPatient
    sTemp = ReadStringField(oRS("NAME"))
    
    .Name.Last_Name = Mid$(sTemp, 1, InStr(1, sTemp, ",") - 1)
    sTemp = Mid$(sTemp, InStr(1, sTemp, ",") + 1)
    If InStr(1, sTemp, " ") > 0 Then
        .Name.First_Name = Trim$(Mid$(sTemp, 1, InStr(1, sTemp, " ") - 1))
        .Name.Middle_Name = Trim$(Mid$(sTemp, InStr(1, sTemp, " ") + 1))
    Else
        .Name.First_Name = sTemp
        .Name.Middle_Name = vbNullString
    End If
    
    .UnitNumber = ReadStringField(oRS("PATIENT_UNITNUMBER"))
    .MedicalRecordNumber = ReadStringField(oRS("Patient_IEN"))
    .SEX = ReadStringField(oRS("SEX"))
    .SSN = ReadStringField(oRS("Patient_ssn"))
    .Formatted_SSN = FormatSSN(.SSN)
    
    If Not IsNull(oRS("DOB")) Then
        If oRS("DOB") <> NULL_DATE And LenB(oRS("DOB")) > 0 Then
            .DateOfBirth = CDate(oRS("DOB"))
        End If
    End If
    
    .FMP = ReadStringField(oRS("FMP"))
    .SponsorSSN = ReadStringField(oRS("Sponsor_SSN"))
    .FMPSponsorSSN = FormatFMP_SSN(.FMP & "/" & .SponsorSSN)
    
    If bBasicOnly Then
        LoadPatientObject = True
        Exit Function
    End If
   
   '--- The following is for full patient information load
    .HomeAddress.Street1 = ReadStringField(oRS("STREET_ADDRESS"))
    .HomeAddress.Street2 = ReadStringField(oRS("STREET_ADDRESS_2"))
    .HomeAddress.City = ReadStringField(oRS("CITY"))
    .HomeAddress.State = ReadStringField(oRS("STATE"))
    .HomeAddress.Zip = ReadStringField(oRS("ZIP_CODE"))
    
    If Not IsNull(oRS("PHONE")) Then
        sTemp = Reformat_Phone(oRS("PHONE"))
        If InStr(1, sTemp, "(") > 0 Then
            .HomePhone.AreaCode = Mid$(sTemp, 2, 3)
            .HomePhone.Phone = ReturnNumbers(Mid$(sTemp, 6))
        Else
            .HomePhone.AreaCode = vbNullString
            .HomePhone.Phone = sTemp
        End If
    End If
    
    If Not IsNull(oRS("OFFICE_PHONE")) Then
        sTemp = Reformat_Phone(oRS("OFFICE_PHONE"))
        If InStr(1, sTemp, "(") > 0 Then
            .WorkPhone.AreaCode = Mid$(sTemp, 2, 3)
            .WorkPhone.Phone = ReturnNumbers(Mid$(sTemp, 6))
        Else
            .WorkPhone.AreaCode = vbNullString
            .WorkPhone.Phone = sTemp
        End If
    End If
    
    .Race = ReadStringField(oRS("Race"))
    .Rank = ReadStringField(oRS("Rank"))
    .Religion = ReadStringField(oRS("Religion"))
    .MaritalStatus = ReadStringField(oRS("MARITAL_STATUS"))
    .Medicare = ReadStringField(oRS("Medicare"))
    .MedicareCode = ReadStringField(oRS("MedicareCode"))
    .PatientCategory = ComposePATCAT(oRS)
    .PrimaryCareManager = ReadStringField(oRS("PCM"))
    .PCM_IEN = ReadStringField(oRS("PCM_IEN"))
    .ACV_EnrollmentStatus = ReadStringField(oRS("ACV"))
    .UnitIdentificationCode = ReadStringField(oRS("Duty Unit"))
    
    .DeersEligibility.Description = ReadStringField(oRS("DEERS_ELIGIBILITY"))
    .DeersEligibility.StartDate = ReadStringField(oRS("DEERS_ELIGIBILITY_START_DATE"))
    .DeersEligibility.EndDate = ReadStringField(oRS("DEERS_ELIGIBILITY_END_DATE"))
    
    If Not IsNull(oRS("Command_Security")) Then
      If LenB(oRS("Command_Security")) > 0 Then
        sData = Split(oRS("Command_Security"), "|")
        For i = 0 To UBound(sData)
                            
            Set oGenericData = New DTOs.PatientGenericData
            oGenericData.TypeNcid = NCID_COMMAND_SECURITY
            oGenericData.ValueText = sData(i)
            oGenericData.ValueNcid = ConvertCommandSecurityToNCID(sData(i))

            .CommandSecurity.Add oGenericData
        Next i
      End If
    End If
    
    If Not IsNull(oRS("Command_Interest")) Then
      If LenB(oRS("Command_Interest")) > 0 Then
        sData = Split(oRS("Command_Interest"), "|")
        For i = 0 To UBound(sData)
                            
            Set oGenericData = New DTOs.PatientGenericData
            oGenericData.TypeNcid = NCID_COMMAND_INTEREST
            oGenericData.ValueText = sData(i)

            .CommandInterest.Add oGenericData
        Next i
      End If
    End If
    
    '--- Pregnancy Data
    If LoadPregnancyObject(.UnitNumber, objPregnancy) Then
        If Not objPregnancy Is Nothing Then
            Set .Pregnancy = objPregnancy
        End If
    End If
    
    .Sponsor_IEN = ReadStringField(oRS("SPONSOR_IEN"))
    
  End With
        
  LoadPatientObject = True
    
End Function






'**********************************************************
' Load Sponsor Object with the CHCS data in Recordset
'**********************************************************

Private Function LoadSponsorObject(ByRef objSponsor As DTOs.Patient, ByVal oRS As ADODB.Recordset) As Boolean
  
Dim sTemp As String

  On Error Resume Next

  LoadSponsorObject = False
  
  With objSponsor
    
    sTemp = ReadStringField(oRS("NAME"))
    
    .Name.Last_Name = Mid$(sTemp, 1, InStr(1, sTemp, ",") - 1)
    sTemp = Mid$(sTemp, InStr(1, sTemp, ",") + 1)
    If InStr(1, sTemp, " ") > 0 Then
        .Name.First_Name = Trim$(Mid$(sTemp, 1, InStr(1, sTemp, " ") - 1))
        .Name.Middle_Name = Trim$(Mid$(sTemp, InStr(1, sTemp, " ") + 1))
    Else
        .Name.First_Name = sTemp
        .Name.Middle_Name = vbNullString
    End If
    
    .PatientCategory = ComposePATCAT(oRS)
    .UnitNumber = ReadStringField(oRS("PATIENT_UNITNUMBER"))
    .MedicalRecordNumber = ReadStringField(oRS("Patient_IEN"))
    .SEX = ReadStringField(oRS("SEX"))
    .SSN = ReadStringField(oRS("Patient_ssn"))
    .Formatted_SSN = FormatSSN(.SSN)
    .Rank = ReadStringField(oRS("Rank"))
    .UnitIdentificationCode = ReadStringField(oRS("Duty Unit"))
    
    
  End With
        
  LoadSponsorObject = True
    
End Function




'**********************************************************
' Load Pregnancy Object with the data in the big pregnancy string
'**********************************************************

Private Function LoadPregnancyObject(ByVal sUnitNumber As String, ByRef objPregancy As DTOs.Pregnancy) As Boolean
    
    Dim sPregnancy() As String
    Dim sSQL As String
    Dim oRS As ADODB.Recordset
    
    On Error GoTo ErrHandler
    
    sSQL = "SELECT VALUE_TEXT FROM MMI_GENERICDATA " & vbNewLine
    sSQL = sSQL & "WHERE UNIT_NUMBER = " & sUnitNumber & vbNewLine
    sSQL = sSQL & " AND TYPE_NCID = " & NCID_PREGNANCY
    
    Set oRS = mConn.CHCSII_DAS(LGS).OpenRecordset(sSQL)
    
    If oRS.RecordCount > 0 Then
        sPregnancy = Split(oRS("VALUE_TEXT"), "~")
        
        If objPregancy Is Nothing Then
            Set objPregancy = New DTOs.Pregnancy
        End If
        
        With objPregancy
            
            If UBound(sPregnancy) >= 0 Then
                .PregnancyIndicator = CBool(sPregnancy(0))
            Else
                .PregnancyIndicator = False
            End If
            
            If UBound(sPregnancy) > 0 Then
                .LastMenstrualPeriod = CDate(sPregnancy(1))
            Else
                .LastMenstrualPeriod = NULL_DATE
            End If
            
            If UBound(sPregnancy) > 1 Then
                .EstimatedDOB = CDate(sPregnancy(2))
            Else
                .EstimatedDOB = NULL_DATE
            End If
               
        End With
    End If
    
    LoadPregnancyObject = True
    
    GoTo CleanUpAndExit
    
ErrHandler:
    Err.Raise Err.Number, "MapperPGUI.LoadPregnancyObject", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
End Function





