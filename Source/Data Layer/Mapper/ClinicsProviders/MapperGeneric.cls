VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MapperGeneric"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements DL_Interfaces.iMapperClinicClient

Private mConn As CHCSII_CONN.Conn
Private mCache As CacheMapper        'Core Cache manager

Private mobjSQLFix As iSQLFix
Private meDataAccessMode As CHCSII_CONN.eBackend

Private Function iMapperClinicClient_GetClinics(NPParams As DL_Support.iParameters) As DTOs.Clinics
  Set iMapperClinicClient_GetClinics = GetClinics(NPParams)
End Function

Private Function iMapperClinicClient_GetProviders(NPParams As iParameters) As DTOs.Providers
  Set iMapperClinicClient_GetProviders = GetProviders(NPParams)
End Function

Private Function iMapperClinicClient_GetClinicsWorkload(NPParams As iParameters, ocolClinics As DTOs.Clinics) As Boolean
  iMapperClinicClient_GetClinicsWorkload = GetClinicsWorkload(NPParams, ocolClinics)
End Function

Private Function iMapperClinicClient_GetProvidersWorkload(NPParams As iParameters, ocolProviders As DTOs.Providers) As Boolean
  iMapperClinicClient_GetProvidersWorkload = GetProvidersWorkload(NPParams, ocolProviders)
End Function

Private Sub iMapperClinicClient_Init(Support As DL_Support.iParameters, ByRef Cache As CacheMapper)
  Set mConn = Support.Item("DBConn").Value
  Set mCache = Cache
End Sub

Private Function GetProviders(ByRef NPParams As DL_Support.iParameters) As DTOs.Providers
Dim objProvidersParams As ProviderSearchParam
Dim rsProviders As ADODB.Recordset
Dim objProvider As DTOs.Provider
Dim objProviders As DTOs.Providers
Dim sSQL As String
Dim bIncludeLogonName As Boolean
    
    On Error GoTo errHandler
    
    Set objProvidersParams = NPParams
    
    Set objProviders = GetCachedProviders(objProvidersParams) 'GetProvidersList

    If objProviders.Count > 0 Then
      'Cached Providers found return
      Set GetProviders = objProviders
      Exit Function
    End If
    
    If Len(objProvidersParams.ProviderRole) > 0 Or objProvidersParams.IncludeRoles Then
      'Get Providers by Roles
      sSQL = GetSQLProvidersRoles(objProvidersParams)
      bIncludeLogonName = True
    ElseIf objProvidersParams.UseClinicProviderTable Then
      'Get Providers from Clinic_Provider table
      sSQL = GetSQLClinicProviders(objProvidersParams)
    Else
      'Get Providers from Provider table
      sSQL = GetSQLProviders(objProvidersParams)
    End If
    
    Set rsProviders = mConn.CHCSII_DAS.OpenRecordset(sSQL)
    
    Set objProviders = New DTOs.Providers
    
    Do While Not rsProviders.EOF
      Set objProvider = New DTOs.Provider
      
      'Need to initialize Workloads for Each provider
      Set objProvider.ProviderWorkloads = New Workloads
      
      If Not objProviders.Exists(CStr(rsProviders.Fields("ClINICIAN_NCID"))) Then
        objProviders.Add objProvider, CStr(rsProviders.Fields("ClINICIAN_NCID"))
      End If
      If objProvidersParams.IncludeExtendedData Then
        Set objProvider.ProviderExtendedData = GetProviderExtendedData(objProvidersParams)
      End If
    
      DeserializeProvider objProvider, rsProviders, bIncludeLogonName
      rsProviders.MoveNext
    Loop
    
    If objProvidersParams.IncludeWorkload Then
      'No Theater Workloads
      Call GetProvidersWorkload(objProvidersParams, objProviders)
    End If
        
    If objProvidersParams.IncludeRoles Then
      'TODO: Need to call Security Object to get the User Object.Roles
'      Call GetProviderRoles(objProvidersParams, objProviders)
    End If
    
    SaveCachedProviders objProviders, objProvidersParams

    Set GetProviders = objProviders
    
    Exit Function
    
errHandler:
    Err.Raise Err.Number, "MapperGeneric.GetProviders", Err.Description, Err.HelpFile, Err.HelpContext

End Function

Private Function GetClinics(ByRef NPParams As DL_Support.iParameters) As DTOs.Clinics

Dim objClinicParams As ClinicSearchParam
Dim rsClinics As ADODB.Recordset
Dim objClinic As DTOs.Clinic
Dim objClinics As DTOs.Clinics
Dim sSQL As String
    
    On Error GoTo errHandler
    
    Set objClinicParams = NPParams
    
    Set objClinics = GetCachedClinics(objClinicParams)

    If objClinics.Count > 0 Then
      'Cached Clinics found return
      Set GetClinics = objClinics
      Exit Function
    End If
    
    sSQL = GetSQLClinics(objClinicParams)
    
    Set objClinics = New DTOs.Clinics
    
    'No Parameters Set
    If sSQL = vbNullString Then Exit Function
    
    Set rsClinics = mConn.CHCSII_DAS.OpenRecordset(sSQL)
    
    If rsClinics.RecordCount > 0 Then
    
      Do While Not rsClinics.EOF
        Set objClinic = New DTOs.Clinic
        
        'Need to initialize Workloads for Each Clinic
        Set objClinic.ClinicWorkloads = New Workloads
        
        objClinics.Add objClinic, rsClinics.Fields("NCID")
      
        If objClinicParams.IncludeWorkload Then
            'No Theater Workloads
            objClinic.Default_Workload_Type = GetClinicWorkloadType(objClinicParams)
        End If
        
        DeserializeClinic objClinic, rsClinics
        rsClinics.MoveNext
      Loop
          
      'Includes Workloads for all Clinics
      If objClinicParams.IncludeWorkload Then
        'No Theater Workloads
        Call GetClinicsWorkload(objClinicParams, objClinics)
      End If
    End If
    
    SaveCachedClinics objClinics, objClinicParams
    
    Set GetClinics = objClinics
    
    Exit Function
    
errHandler:
    Err.Raise Err.Number, "MapperGeneric.GetClinics", Err.Description, Err.HelpFile, Err.HelpContext

End Function

Private Function GetSQLClinics(ByRef params As ClinicSearchParam) As String
Dim sSelect As String
Dim sSQL As String

    On Error GoTo errHandler
    
    sSelect = "Select c.Facility_NCID, c.NCID, c.IEN, c.Name, c.Meprs_IEN, c.Meprs_Code,"
    sSelect = sSelect & " c.WALK_IN, c.SICK_CALL, c.STREET_ADDRESS, c.BUILDING_NUMBER, c.City,"
    sSelect = sSelect & " c.State, c.ZIP, c.DMIS_ID_CODE, m.Description as MeprsDescription,"
    sSelect = sSelect & " c.Location_Type"
    
    
    'NOT IN THEATER (c.BDM_DEF_PHARMACY_ID, c.DEFAULT_WORKLOAD_TYPE)
    If gobjMiscParam.PharmacyID Then
      sSelect = sSelect & ", c.BDM_DEF_PHARMACY_ID"
    End If
    
    If gobjMiscParam.WorkloadType Then
      sSelect = sSelect & ", c.DEFAULT_WORKLOAD_TYPE"
    End If
    
    'Only In CDR Mode (l.Location_IEN)
    If gobjMiscParam.Location_IEN Then
      sSelect = sSelect & ", l.Location_IEN"
    End If
    
    sSQL = sSelect
    
    'Get User Clinics
    If Len(params.UserNCID) > 0 Then
      If params.IncludeUserClinicsOnly Then
        sSQL = sSQL & " From Clinic c, User_Clinic p, Meprs_Code m"
        sSQL = sSQL & GetGeographic_LocationTable(gobjMiscParam.Location_IEN)
        sSQL = sSQL & GetClauseUser_NCID(params.UserNCID, True, "p.")
      ElseIf params.UseClinicProviderTable Then
        sSQL = sSQL & " From Clinic c, Clinic_Provider p, Meprs_Code m"
        sSQL = sSQL & GetGeographic_LocationTable(gobjMiscParam.Location_IEN)
        sSQL = sSQL & GetClauseClinician_NCID(params.UserNCID, True, "p.")
      Else
        sSQL = sSQL & " From Clinic c, Provider p, Meprs_Code m"
        sSQL = sSQL & GetGeographic_LocationTable(gobjMiscParam.Location_IEN)
        sSQL = sSQL & GetClauseNCID(params.UserNCID, True, "p.")
      End If
        
      'Check ClinicNCID Value
      sSQL = sSQL & GetClauseNCID(params.ClinicNCID, False, "c.")
      
      'Check ClinicIEN Value
      sSQL = sSQL & GetClauseIEN(params.ClinicIEN, False, "c.")
      
      'Check FacilityNCID Value
      sSQL = sSQL & GetClauseFacility_NCID(params.FacilityNCID, False, "c.")
      
      'Check Include NonMTF Location Type
      sSQL = sSQL & GetClauseIncludeNonMTF(params.IncludeNonMTF, False, "c.")
      
      sSQL = sSQL & " and c.NCID = p.Clinic_NCID"
      sSQL = sSQL & " and c.Meprs_IEN = m.IEN"
      sSQL = sSQL & " and (c.InActive <> 1 or InActive is NULL) "
      sSQL = sSQL & " and c.Facility_NCID = m.Facility_NCID"
        
      'Clinic Division State IEN
      If gobjMiscParam.Location_IEN Then
        sSQL = sSQL & " and " & mobjSQLFix.GetSubStrOuterJoin("c.Division_State", "l.Location_Abbreviation", 1, 2)
      End If
      
      sSQL = sSQL & " Order by c.Name"
  
      'Include IBWA Clinics (SQL UNION STATEMENT)
      sSQL = sSQL & GetIBWAClinics(params, sSelect)
      
    'Get Single Clinic
    ElseIf (Len(params.ClinicNCID) > 0) Or (Len(params.ClinicIEN) > 0) Then
      sSQL = sSQL & " From Clinic c, Meprs_Code m"
      sSQL = sSQL & GetGeographic_LocationTable(gobjMiscParam.Location_IEN)
      sSQL = sSQL & " Where " & mobjSQLFix.GetOuterJoin("c.Facility_NCID", "m.Facility_NCID")  '" Where c.facility_ncid = m.facility_ncid(+)"
      sSQL = sSQL & " And " & mobjSQLFix.GetOuterJoin("c.MEPRS_IEN", "m.IEN")  '" and c.MEPRS_IEN = m.IEN(+)"
      
      'Check ClinicNCID Value
      sSQL = sSQL & GetClauseNCID(params.ClinicNCID, False, "c.")
      
      'Check ClinicIEN Value
      sSQL = sSQL & GetClauseIEN(params.ClinicIEN, False, "c.")
      
      'Check FacilityNCID Value
      sSQL = sSQL & GetClauseFacility_NCID(params.FacilityNCID, False, "c.")
      
      'Check Include NonMTF Location Type
      sSQL = sSQL & GetClauseIncludeNonMTF(params.IncludeNonMTF, False, "c.")
    
      'Clinic Division State IEN
      If gobjMiscParam.Location_IEN Then
        sSQL = sSQL & " and " & mobjSQLFix.GetSubStrOuterJoin("c.Division_State", "l.Location_Abbreviation", 1, 2)
      End If
      
      sSQL = sSQL & " Order by c.Name"
    
      'Include IBWA Clinics (SQL UNION STATEMENT)
      sSQL = sSQL & GetIBWAClinics(params, sSelect)
    
    'Get Clinics based on FacilityNCID
    ElseIf Len(params.FacilityNCID) > 0 Then
      sSQL = sSQL & " From Clinic c, Meprs_Code m"
      sSQL = sSQL & GetGeographic_LocationTable(gobjMiscParam.Location_IEN)
      sSQL = sSQL & " Where m.facility_ncid = c.facility_ncid"
      sSQL = sSQL & " And m.IEN = c.Meprs_IEN"
      sSQL = sSQL & " And c.facility_ncid = " & params.FacilityNCID
        
      'Check Include NonMTF Location Type
      sSQL = sSQL & GetClauseIncludeNonMTF(params.IncludeNonMTF, False, "c.")
        
      sSQL = sSQL & " and (c.InActive <> 1 or InActive is NULL)"
      
      'Clinic Division State IEN
      If gobjMiscParam.Location_IEN Then
        sSQL = sSQL & " and " & mobjSQLFix.GetSubStrOuterJoin("c.Division_State", "l.Location_Abbreviation", 1, 2)
      End If
      
      sSQL = sSQL & " Order by c.Name"
      
      'Include IBWA Clinics (SQL UNION STATEMENT)
      sSQL = sSQL & GetIBWAClinics(params, sSelect)
      
    Else
      sSQL = vbNullString
    End If
    
    GetSQLClinics = sSQL
    
    Exit Function
    
errHandler:
    Err.Raise Err.Number, "MapperGeneric.GetSQLClinics", Err.Description, Err.HelpFile, Err.HelpContext
    
End Function

Private Function DeserializeClinic(ByRef objClinic As DTOs.Clinic, ByVal oRS As ADODB.Recordset) As Boolean
  On Error Resume Next
    
  With objClinic
    .FacilityNCID = oRS("FACILITY_NCID") & vbNullString
    .NCID = oRS("NCID") & vbNullString
    .IEN = oRS("IEN") & vbNullString
    .Name = oRS("Name") & vbNullString
    .MeprsIEN = oRS("Meprs_IEN") & vbNullString
    .MeprsCode = oRS("Meprs_Code") & vbNullString
    .MeprsDescription = oRS("MeprsDescription") & vbNullString
    .LocationType = oRS("LOCATION_TYPE") & vbNullString
    .WalkInAllowed = (oRS("WALK_IN") = 1)
    .SickCallAllowed = (oRS("SICK_CALL") = 1)
    .StreetAddress = oRS("STREET_ADDRESS") & vbNullString
    .BuildingNumber = oRS("BUILDING_NUMBER") & vbNullString
    .City = oRS("CITY") & vbNullString
    .State = oRS("STATE") & vbNullString
    .ZipCode = oRS("ZIP") & vbNullString
    .DmisIdCode = oRS("DMIS_ID_CODE") & vbNullString
    
    If gobjMiscParam.PharmacyID Then
      .DefPharmacyBDM = oRS("BDM_DEF_PHARMACY_ID") & vbNullString
    Else
      .DefPharmacyBDM = vbNullString
    End If
    
    If gobjMiscParam.WorkloadType Then
      .Default_Workload_Type = oRS("DEFAULT_WORKLOAD_TYPE") & vbNullString
    Else
      .Default_Workload_Type = vbNullString
    End If
    
    If .Default_Workload_Type = vbNullString Then
       .Default_Workload_Type = "N"
    End If
    
    If gobjMiscParam.Location_IEN Then
      .DivisionStateIEN = oRS("Location_IEN")
    Else
      .DivisionStateIEN = vbNullString
    End If
    
  End With
  
  DeserializeClinic = True
    
End Function

Private Function GetClinicWorkloadType(ByRef params As ClinicSearchParam) As String
    Dim rst                 As ADODB.Recordset
    Dim sProviderType       As String
    Dim sClinicType         As String
    Dim sDefaultType        As String
    Dim sSQL       As String
    Dim bResult             As Boolean
    
    On Error GoTo errHandler
    
    GetClinicWorkloadType = "N"
    
    sSQL = "SELECT DEFAULT_WORKLOAD_TYPE AS WORKLOAD_TYPE, 'DEFAULT' AS LEVELType"
    sSQL = sSQL & " From Clinic"
    sSQL = sSQL & " Where FACILITY_NCID = " & params.FacilityNCID
    sSQL = sSQL & " And NCID = " & params.ClinicNCID
    sSQL = sSQL & " Union"
    sSQL = sSQL & " Select WORKLOAD_TYPE,  'CLINIC' AS LEVELType"
    sSQL = sSQL & " From CLINIC_APPT_TYPE"
    sSQL = sSQL & " WHERE FACILITY_NCID=" & params.FacilityNCID
    sSQL = sSQL & " AND CLINIC_NCID=" & params.ClinicNCID
    sSQL = sSQL & " AND APPT_CODE= '" & params.ApptCode & "'"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " SELECT WORKLOAD_TYPE,  'PROVIDER' AS LEVELType"
    sSQL = sSQL & " From PROVIDER_APPT_TYPE "
    sSQL = sSQL & " WHERE FACILITY_NCID=" & params.FacilityNCID
    sSQL = sSQL & " AND CLINIC_NCID=" & params.ClinicNCID
    sSQL = sSQL & " AND PROVIDER_NCID= " & params.UserNCID '??ProviderNCID
    sSQL = sSQL & " AND APPT_CODE= '" & params.ApptCode & "'"
    
    Set rst = mConn.CHCSII_DAS.OpenRecordset(sSQL)
    If rst.RecordCount > 0 Then
        rst.MoveFirst 'Get the Providers Workload
        rst.Find "LEVELType='PROVIDER'"
        If Not rst.EOF Then
            sProviderType = rst.Fields("WORKLOAD_TYPE").Value & vbNullString
        Else
            sProviderType = vbNullString
        End If
        rst.MoveFirst 'Get the Clinic's Appointment Workload
        rst.Find "LEVELType='CLINIC'"
        If Not rst.EOF Then
            sClinicType = rst.Fields("WORKLOAD_TYPE").Value & vbNullString
        Else
            sClinicType = vbNullString
        End If
        rst.MoveFirst 'Get the Clinic's Default Workload
        rst.Find "LEVELType='DEFAULT'"
        If Not rst.EOF Then
            sDefaultType = rst.Fields("WORKLOAD_TYPE").Value & vbNullString
        Else
            sDefaultType = "N"
        End If
        
        bResult = (UCase$(sDefaultType) = "C")
        
        If bResult Then
        
            If Len(sProviderType) = 0 Then
                bResult = bResult And (Len(sClinicType) = 0 Or (UCase$(sClinicType) = "C"))
            ElseIf bResult And (UCase$(sProviderType) = "C") Then
                bResult = True
            Else
                bResult = False
            End If
        End If
        
        If bResult Then
          GetClinicWorkloadType = "NC"
        Else
          GetClinicWorkloadType = "N"
        End If
        
    End If
    Exit Function
    
errHandler:
    Err.Raise Err.Number, "MapperGeneric.GetClinicWorkloadType", Err.Description, Err.HelpFile, Err.HelpContext

End Function


Private Function GetClinicsWorkload(ByRef params As ClinicSearchParam, ByRef ocolClinics As DTOs.Clinics) As Boolean
Dim rst  As ADODB.Recordset
Dim objClinic As DTOs.Clinic
Dim objWorkload As DTOs.WorkLoad
Dim sSQL As String
Dim sClause As String
Dim sOldClinicNCID As String
Dim sNewClinicNCID As String

  On Error GoTo errHandler
  
  If ocolClinics Is Nothing Then
    GetClinicsWorkload = False
    Exit Function
  End If
  
  'If No Clinics then return True to GetClinicsWorkload
  If ocolClinics.Count = 0 Then
    GetClinicsWorkload = True
    Exit Function
  End If

  
  sSQL = "Select Distinct c.Clinic_NCID, c.Appt_Code, a.Description, c.Workload_Type, c.Appt_Duration"
  sSQL = sSQL & " From CLINIC_APPT_TYPE c, APPT_TYPE a"
  sSQL = sSQL & " Where c.Appt_Code = a.Code"
  sSQL = sSQL & " And c.FACILITY_NCID = a.FACILITY_NCID"
  sSQL = sSQL & " And c.Facility_NCID = " & params.FacilityNCID
            
  '''If limited subset of clinics
  If ocolClinics.Count < 21 Then
    '''Implement In clause to improve performance
    For Each objClinic In ocolClinics
        sClause = sClause & objClinic.NCID & ","
    Next objClinic
    sClause = Left$(sClause, Len(sClause) - 1)
    sSQL = sSQL & " And c.Clinic_NCID In (" & sClause & ")"
  End If
            
  sSQL = sSQL & " Order By c.Clinic_NCID, c.Appt_Code"

  Set rst = mConn.CHCSII_DAS.OpenRecordset(sSQL)
  
  Do While Not rst.EOF
    sNewClinicNCID = rst.Fields("Clinic_NCID").Value & vbNullString
    If ocolClinics.Exists(sNewClinicNCID) Then
      Set objWorkload = New DTOs.WorkLoad
      Call DeserializeWorkload(objWorkload, rst)
      If sNewClinicNCID <> sOldClinicNCID Then
        Set objClinic = ocolClinics.Item(sNewClinicNCID)
        If objClinic.ClinicWorkloads Is Nothing Then
          Set objClinic.ClinicWorkloads = New DTOs.Workloads
        End If
        sOldClinicNCID = sNewClinicNCID
      End If
      objClinic.ClinicWorkloads.Add objWorkload, objWorkload.Appt_Code
    End If
    rst.MoveNext
  Loop
  
  GetClinicsWorkload = True
  
  Exit Function
    
errHandler:
    Err.Raise Err.Number, "MapperGeneric.GetClinicsWorkload", Err.Description, Err.HelpFile, Err.HelpContext
  
End Function

Private Function GetSQLProviders(ByRef params As ProviderSearchParam) As String
Dim sSQL As String
Dim bFirstClause As Boolean
  On Error GoTo errHandler

  bFirstClause = True
  
  sSQL = "Select Distinct p.NCID AS Clinician_NCID, p.IEN as Provider_IEN,"
  sSQL = sSQL & " p.Name as Provider_Name, p.User_IEN, p.Signature_Class"
  
  If params.ClinicNCID <> vbNullString And params.ClinicNCID <> "0" Then
    sSQL = sSQL & " From Provider p, Clinic_Provider cp"
    sSQL = sSQL & " Where cp.clinician_ncid = p.ncid"
    sSQL = sSQL & " And ( cp.Clinic_NCID = " & params.ClinicNCID
    sSQL = sSQL & " OR " & " p.Clinic_NCID = " & params.ClinicNCID & ")"
    bFirstClause = False
  Else
    sSQL = sSQL & " From Provider p"
  End If

  'Check ProviderIEN value
  sSQL = sSQL & GetClauseIEN(params.ProviderIEN, bFirstClause, "p.")
  
  'Check ProviderNCID value
  sSQL = sSQL & GetClauseNCID(params.ProviderNCID, bFirstClause, "p.")
  
  'Check SignableOnly Value
  sSQL = sSQL & GetClauseSignableOnly(params.SignableOnly, bFirstClause, "p.")
  
  'Check LastName Value
  sSQL = sSQL & GetClauseNameLike(params.ProviderLastName, bFirstClause, "p.")
  
  'Check ClinicUsersOnly Value
  sSQL = sSQL & GetClauseClinicUsersOnly(params.ClinicUsersOnly, False, bFirstClause, "p.")
  
  'Check FacilityNCID value
  sSQL = sSQL & GetClauseFacility_NCID(params.FacilityNCID, bFirstClause, "p.")
  
  sSQL = sSQL & " Order By Provider_name"
  
  GetSQLProviders = sSQL
  
  Exit Function
    
errHandler:
    Err.Raise Err.Number, "MapperGeneric.GetSQLProviders", Err.Description, Err.HelpFile, Err.HelpContext

End Function

Private Function GetSQLProvidersRoles(ByRef params As ProviderSearchParam) As String
Dim sSQL As String
Dim bFirstClause As Boolean

    On Error GoTo errHandler
    
        If meDataAccessMode = MYSQL Then
            sSQL = "Select distinct c.Clinician_NCID, c.Provider_IEN, c.Provider_Name, c.User_IEN," _
              & "h.LOGON_NAME, c.SIGNATURE_CLASS " _
              & "From Clinic_Provider c LEFT JOIN hems_users h ON c.clinician_ncid = h.user_ncid "
              bFirstClause = True
        Else
            sSQL = "Select Distinct c.Clinician_NCID, c.Provider_IEN, c.Provider_Name,"
            sSQL = sSQL & " c.User_IEN, h.Logon_Name, c.Signature_Class"
            sSQL = sSQL & " From Hems_Users h, Clinic_Provider c"
            sSQL = sSQL & " Where " & mobjSQLFix.GetOuterJoin("c.Clinician_NCID", "h.User_NCID") '" c.Clinician_NCID = h.User_NCID(+)"
            bFirstClause = False
        End If
    
    If params.ClinicNCID <> vbNullString And params.ClinicNCID <> "0" Then
      sSQL = sSQL & GetClauseClinic_NCID(params.ClinicNCID, False, "c.")
    End If
        
    'Check ProviderIEN Value
    sSQL = sSQL & GetClauseProvider_IEN(params.ProviderIEN, bFirstClause, "c.")
    
    'Check ProviderNCID Value
    sSQL = sSQL & GetClauseClinician_NCID(params.ProviderNCID, bFirstClause, "c.")
    
    'Check SignableOnly Value
    sSQL = sSQL & GetClauseSignableOnly(params.SignableOnly, bFirstClause, "c.")
    
    'Check Provider_Name Value
    sSQL = sSQL & GetClauseProvider_NameLike(params.ProviderLastName, bFirstClause, "c.")
    
    'Check Clinic Users Only
     sSQL = sSQL & GetClauseClinicUsersOnly(params.ClinicUsersOnly, True, bFirstClause, "c.")
    
    'Check FacilityNCID value
    sSQL = sSQL & GetClauseFacility_NCID(params.FacilityNCID, bFirstClause, "c.")
     
    sSQL = sSQL & "Order By Provider_Name"
    
    GetSQLProvidersRoles = sSQL
    
    Exit Function
    
errHandler:
    Err.Raise Err.Number, "MapperGeneric.GetSQLProvidersRoles", Err.Description, Err.HelpFile, Err.HelpContext

End Function

Private Function GetSQLClinicProviders(ByRef params As ProviderSearchParam) As String
Dim sSQL As String
Dim bFirstClause As Boolean
  On Error GoTo errHandler

    If meDataAccessMode = MYSQL Then
        sSQL = "Select distinct c.Clinician_NCID, c.Provider_IEN, c.Provider_Name, c.User_IEN," _
            & "h.LOGON_NAME, c.SIGNATURE_CLASS " _
            & "From Clinic_Provider c LEFT JOIN hems_users h ON c.clinician_ncid = h.user_ncid "
        bFirstClause = True
    ElseIf meDataAccessMode = GEMS Then
        sSQL = "Select Distinct c.Clinician_NCID, c.Provider_IEN, c.Provider_Name, c.User_IEN," _
                    & "h.LOGON_NAME, c.SIGNATURE_CLASS " _
                    & "From hems_users h, Clinic_Provider c "
        sSQL = sSQL & " Where " & mobjSQLFix.GetOuterJoin("c.Clinician_NCID", "h.User_NCID") '" c.Clinician_NCID = h.User_NCID(+)"
        bFirstClause = False
    End If
    
    If params.ClinicNCID <> vbNullString And params.ClinicNCID <> "0" Then
      sSQL = sSQL & GetClauseClinic_NCID(params.ClinicNCID, False, "c.")
      bFirstClause = False
    End If
    
    sSQL = sSQL & GetClauseProvider_NameLike(params.ProviderLastName, bFirstClause)
    sSQL = sSQL & GetClauseClinicUsersOnly(params.ClinicUsersOnly, True, bFirstClause, "c.")
    sSQL = sSQL & GetClauseFacility_NCID(params.FacilityNCID, bFirstClause)
    sSQL = sSQL & GetClauseClinic_NCID(params.ClinicNCID, bFirstClause)
    sSQL = sSQL & GetClauseSignableOnly(params.SignableOnly, bFirstClause)
    
    'Check ProviderIEN value
    sSQL = sSQL & GetClauseProvider_IEN(params.ProviderIEN, bFirstClause, "c.")
    
    'Check ProviderNCID value
    sSQL = sSQL & GetClauseClinician_NCID(params.ProviderNCID, bFirstClause, "c.")
     
    sSQL = sSQL & " Order By Provider_name"
    
    GetSQLClinicProviders = sSQL
    
    Exit Function
    
errHandler:
    Err.Raise Err.Number, "MapperGeneric.GetSQLClinicProviders", Err.Description, Err.HelpFile, Err.HelpContext
  
End Function

Private Function GetIBWAClinics(ByRef params As ClinicSearchParam, sSelect As String) As String
Dim sSQL As String

On Error GoTo errHandler

  If params.IncludeIBWAClinics Then
    sSelect = Replace(sSelect, "l.Location_IEN", "0")
    sSQL = "Union " & sSelect & " From Clinic c, Meprs_Code m"
    sSQL = sSQL & " Where " & mobjSQLFix.GetOuterJoin("c.Facility_NCID", "m.Facility_NCID")  '" Where c.facility_ncid = m.facility_ncid(+)"
    sSQL = sSQL & " And " & mobjSQLFix.GetOuterJoin("c.MEPRS_IEN", "m.IEN")  '" and c.MEPRS_IEN = m.IEN(+)"
    sSQL = sSQL & " And c.Name Like 'IBWA%'"
    
    If Len(params.FacilityNCID) > 0 Then
      'Check FacilityNCID Value
      sSQL = sSQL & GetClauseFacility_NCID(params.FacilityNCID, False, "c.")
    End If
  End If

  GetIBWAClinics = sSQL
Exit Function
    
errHandler:
    Err.Raise Err.Number, "MapperGeneric.GetIBWAClinics", Err.Description, Err.HelpFile, Err.HelpContext

End Function

Private Function GetProviderExtendedData(ByRef params As ProviderSearchParam) As ProviderExt
Dim objExtended As ProviderExt
Dim sSQL As String
Dim rst As ADODB.Recordset
  
  On Error GoTo errHandler

  Set objExtended = Nothing
  sSQL = "Select Distinct NCID, NAME, DUTY_PHONE_1, DUTY_PHONE_2, IEN, SIGNATURE_CLASS,"
  sSQL = sSQL & " USER_IEN, SSN, SPECIALTY_CODE, CLINIC_NCID, FACILITY_NCID"
  sSQL = sSQL & " From Provider"
  sSQL = sSQL & " Where NCID=" & params.ProviderNCID
         
  Set rst = mConn.CHCSII_DAS.OpenRecordset(sSQL)
  
  If Not rst.EOF Then
    Set objExtended = New ProviderExt
    Call DeserializeProvExtended(objExtended, rst)
  End If
  
  Set GetProviderExtendedData = objExtended
  
  Exit Function
    
errHandler:
    Err.Raise Err.Number, "MapperGeneric.GetProviderExtendedData", Err.Description, Err.HelpFile, Err.HelpContext
  
End Function

Private Function DeserializeProvExtended(ByRef objExtended As ProviderExt, ByVal rst As ADODB.Recordset) As Boolean
  
  On Error GoTo errHandler
  
  With objExtended
    .NCID = rst("NCID") & vbNullString
    .Phone1 = rst("DUTY_PHONE_1") & vbNullString
    .Phone2 = rst("DUTY_PHONE_2") & vbNullString
    .ProviderIEN = rst("IEN") & vbNullString
    .SignatureClass = rst("SIGNATURE_CLASS") & vbNullString
    .UserIEN = rst("USER_IEN") & vbNullString
    .SSN = rst("SSN") & vbNullString
    .SpecialityCode = rst("SPECIALTY_CODE") & vbNullString
    .DefaultClinicNCID = rst("CLINIC_NCID") & vbNullString
    .FacilityNCID = rst("FACILITY_NCID") & vbNullString
  End With
  
  DeserializeProvExtended = True
  Exit Function
    
errHandler:
    Err.Raise Err.Number, "MapperGeneric.DeserializeProvExtended", Err.Description, Err.HelpFile, Err.HelpContext

End Function

Private Function GetProvidersWorkload(ByRef params As ProviderSearchParam, ByRef ocolProviders As DTOs.Providers) As Boolean
Dim rst  As ADODB.Recordset
Dim objProvider As DTOs.Provider
Dim objWorkload As DTOs.WorkLoad
Dim sSQL As String
Dim sClause As String
Dim sOldProviderNCID As String
Dim sNewProviderNCID As String

  On Error GoTo errHandler
  
  If ocolProviders Is Nothing Then
    GetProvidersWorkload = False
    Exit Function
  End If
  
  'If No Providers then return True to GetProvidersWorkload
  If ocolProviders.Count = 0 Then
    GetProvidersWorkload = True
    Exit Function
  End If
  
  sSQL = "Select Distinct p.Appt_Code, a.Description, p.Workload_Type, p.Appt_Duration, p.PROVIDER_NCID"
  sSQL = sSQL & " From PROVIDER_APPT_TYPE p, APPT_TYPE a"
  sSQL = sSQL & " Where p.Appt_Code = a.Code"
  sSQL = sSQL & " And p.Facility_NCID = a.FACILITY_NCID"
  sSQL = sSQL & " And p.Facility_NCID = " & params.FacilityNCID
  sSQL = sSQL & " And p.Clinic_NCID = " & params.ClinicNCID
            
  '''If limited subset of providers
  If ocolProviders.Count < 21 Then
    '''Implement In clause to improve performance
    For Each objProvider In ocolProviders
        sClause = sClause & objProvider.NCID & ","
    Next objProvider
    sClause = Left$(sClause, Len(sClause) - 1)
    sSQL = sSQL & " And p.PROVIDER_NCID In (" & sClause & ")"
  End If
  
  sSQL = sSQL & " Order By p.PROVIDER_NCID, p.Appt_Code"

  Set rst = mConn.CHCSII_DAS.OpenRecordset(sSQL)
  
  Do While Not rst.EOF
    sNewProviderNCID = rst.Fields("Provider_NCID").Value & vbNullString
    If ocolProviders.Exists(sNewProviderNCID) Then
      Set objWorkload = New DTOs.WorkLoad
      Call DeserializeWorkload(objWorkload, rst)
      If sNewProviderNCID <> sOldProviderNCID Then
        Set objProvider = ocolProviders.Item(sNewProviderNCID)
        If objProvider.ProviderWorkloads Is Nothing Then
          Set objProvider.ProviderWorkloads = New DTOs.Workloads
        End If
        sOldProviderNCID = sNewProviderNCID
      End If
      objProvider.ProviderWorkloads.Add objWorkload, objWorkload.Appt_Code
    End If
    rst.MoveNext
  Loop
  
  GetProvidersWorkload = True
  Exit Function
    
errHandler:
    Err.Raise Err.Number, "MapperGeneric.GetProvidersWorkload", Err.Description, Err.HelpFile, Err.HelpContext

End Function

Private Function DeserializeWorkload(ByRef objWorkload As DTOs.WorkLoad, ByVal oRS As ADODB.Recordset) As Boolean
  With objWorkload
    .Appt_Code = oRS("Appt_Code") & vbNullString
    .Appt_Description = oRS("Description") & vbNullString
    .Appt_Duration = oRS("Appt_Duration") & vbNullString
    .Workload_Type = oRS("Workload_Type") & vbNullString
  End With
  
  DeserializeWorkload = True
End Function

Private Function DeserializeProvider(ByRef objProvider As DTOs.Provider, ByVal oRS As ADODB.Recordset, _
          Optional ByVal bIncludeLogonName As Boolean = False) As Boolean
  On Error Resume Next
  
  With objProvider
    .NCID = oRS("Clinician_NCID") & vbNullString
    .IEN = oRS("Provider_IEN") & vbNullString
    .Name = oRS("Provider_Name") & vbNullString
    .SignatureClass = oRS("Signature_Class") & vbNullString
    If bIncludeLogonName Then
      .LogonName = oRS("Logon_Name") & vbNullString
    End If
  End With
  
  DeserializeProvider = True
    
End Function

Public Property Set SQLFix(ByVal NewValue As iSQLFix)
  Set mobjSQLFix = NewValue
End Property

Public Property Get DataAccessMode() As CHCSII_CONN.eBackend
  DataAccessMode = meDataAccessMode
End Property

Public Property Let DataAccessMode(ByVal NewValue As CHCSII_CONN.eBackend)
  meDataAccessMode = NewValue
End Property

Private Function GetClauseProvider_NameLike(ByVal sLastName As String, ByRef bFirstClause As Boolean, Optional ByVal sAlias As String = vbNullString) As String
    
  If Trim$(sLastName) = vbNullString Then Exit Function
  
  GetClauseProvider_NameLike = ClauseConcat(bFirstClause) & " UPPER(" & sAlias & "Provider_Name) LIKE " & mConn.SQLQuote(UCase$(sLastName) & "%")
End Function

Private Function GetClauseNameLike(ByVal sLastName As String, ByRef bFirstClause As Boolean, Optional ByVal sAlias As String = vbNullString) As String
    
  If Trim$(sLastName) = vbNullString Then Exit Function
  
  GetClauseNameLike = ClauseConcat(bFirstClause) & " UPPER(" & sAlias & "Name) LIKE " & mConn.SQLQuote(UCase$(sLastName) & "%")
End Function

Private Function GetClauseClinicUsersOnly(ByVal bClinicUsersOnly As Boolean, ByVal bUseClinicProviderTable As Boolean, ByRef bFirstClause As Boolean, Optional ByVal sAlias As String = vbNullString) As String
    
  If Not bClinicUsersOnly Then Exit Function
  If bUseClinicProviderTable Then
    GetClauseClinicUsersOnly = ClauseConcat(bFirstClause) & " not (h.Logon_Name Is NULL and h.Logon_Name = '') "
  Else
    GetClauseClinicUsersOnly = ClauseConcat(bFirstClause) & " Exists (select 'x' from Hems_Users H where H.User_NCID = " & sAlias & "NCID)"
  End If
End Function

Private Function GetClauseFacility_NCID(ByVal sFacilityNCID As String, ByRef bFirstClause As Boolean, Optional ByVal sAlias As String = vbNullString) As String

  If Trim$(sFacilityNCID) = vbNullString Then Exit Function
  
  GetClauseFacility_NCID = ClauseConcat(bFirstClause) & sAlias & "Facility_NCID = " & sFacilityNCID
End Function

Private Function GetClauseClinic_NCID(ByVal sClinicNCID As String, ByRef bFirstClause As Boolean, Optional ByVal sAlias As String = vbNullString) As String
  If Trim$(sClinicNCID) = vbNullString Or Trim$(sClinicNCID) = "0" Then Exit Function
  
  GetClauseClinic_NCID = ClauseConcat(bFirstClause) & sAlias & "Clinic_NCID = " & sClinicNCID
End Function

Private Function GetClauseNCID(ByVal sNCID As String, ByRef bFirstClause As Boolean, Optional ByVal sAlias As String = vbNullString) As String
  
  If Trim$(sNCID) = vbNullString Or Trim$(sNCID) = "0" Then Exit Function
  
  GetClauseNCID = ClauseConcat(bFirstClause) & sAlias & "NCID = " & sNCID
End Function

Private Function GetClauseIEN(ByVal sIEN As String, ByRef bFirstClause As Boolean, Optional ByVal sAlias As String = vbNullString) As String
  
  If Trim$(sIEN) = vbNullString Or Trim$(sIEN) = "0" Then Exit Function
    
    ' SCR #50235;   Developer: MAG 03/04/2004
    'PGUI Mode. IEN column is a string in PGUI, added single quotes for performance.
    If DataAccessMode = LGS Then
      sIEN = "'" & sIEN & "'"
    End If
    
    GetClauseIEN = ClauseConcat(bFirstClause) & sAlias & "IEN = " & sIEN
End Function

Private Function GetClauseSignableOnly(ByVal bSignableOnly As Boolean, ByRef bFirstClause As Boolean, Optional ByVal sAlias As String = vbNullString) As String
  
  If Not bSignableOnly Then Exit Function
  
  GetClauseSignableOnly = ClauseConcat(bFirstClause) & sAlias & "SIGNATURE_CLASS IN(2,3,4)"
End Function

Private Function GetClauseUser_NCID(ByVal sNCID As String, ByRef bFirstClause As Boolean, Optional ByVal sAlias As String = vbNullString) As String
  
  If Trim$(sNCID) = vbNullString Then Exit Function
  
  GetClauseUser_NCID = ClauseConcat(bFirstClause) & sAlias & "User_NCID = " & sNCID
End Function

Private Function GetClauseIncludeNonMTF(ByVal bIncludeNonMTF As Boolean, ByRef bFirstClause As Boolean, Optional ByVal sAlias As String = vbNullString) As String
  
  If bIncludeNonMTF Then Exit Function
  
  GetClauseIncludeNonMTF = ClauseConcat(bFirstClause) & sAlias & "LOCATION_TYPE <> 'MCP NON-MTF'"
End Function

Private Function GetClauseProvider_IEN(ByVal sIEN As String, ByRef bFirstClause As Boolean, Optional ByVal sAlias As String = vbNullString) As String
  
  If Trim$(sIEN) = vbNullString Then Exit Function
  
  ' SCR #50235;   Developer: MAG 03/04/2004
  'PGUI Mode. IEN column is a string in PGUI, added single quotes for performance.
  If DataAccessMode = LGS Then
    sIEN = "'" & sIEN & "'"
  End If
  
  GetClauseProvider_IEN = ClauseConcat(bFirstClause) & sAlias & "Provider_IEN = " & sIEN
End Function

Private Function GetClauseClinician_NCID(ByVal sNCID As String, ByRef bFirstClause As Boolean, Optional ByVal sAlias As String = vbNullString) As String
  
  If Trim$(sNCID) = vbNullString Then Exit Function
  
  GetClauseClinician_NCID = ClauseConcat(bFirstClause) & sAlias & "Clinician_NCID = " & sNCID
End Function

Private Function GetGeographic_LocationTable(ByVal bInclude As Boolean) As String
  
  If bInclude Then
    GetGeographic_LocationTable = ", Geographic_Location l"
  End If
End Function

Private Function ClauseConcat(ByRef bFirstClause As Boolean) As String
  
  If bFirstClause Then
    ClauseConcat = " Where "
    bFirstClause = False
  Else
    ClauseConcat = " And "
  End If
End Function

Private Sub SaveCachedClinics(ByRef Clinics As DTOs.Clinics, ByRef params As iParameters)
Dim ClinicParams As ClinicSearchParam
    
    Set ClinicParams = params
    
    With ClinicParams
      If Len(.ApptCode) <> 0 Or .IncludeWorkload Or .IncludeNonMTF Then
          'Do not Cache when these values are set
          Exit Sub
      End If
    End With
    
    mCache.Save CACHE_NAME, Clinics, params

End Sub

Private Sub SaveCachedProviders(ByRef Providers As DTOs.Providers, ByRef params As iParameters)
Dim ProvParams As ProviderSearchParam

    Set ProvParams = params
    
    With ProvParams
      If .IncludeExtendedData Or .IncludeRoles Or .IncludeWorkload _
        Or Len(.ProviderLastName) <> 0 Or Len(.ProviderRole) <> 0 _
        Or .SignableOnly Then
          'Do not Cache when these values are set
          Exit Sub
      End If
    End With
    
    mCache.Save CACHE_NAME, Providers, params

End Sub

Private Function GetCachedProviders(ByRef params As iParameters) As DTOs.Providers
Dim CacheEntry As iCacheEntry
Dim objParams As iParameters
Dim objProvs As DTOs.Providers

    On Error GoTo errHandler
    
    Set objParams = params

    'Get the list if it exists
    Set CacheEntry = mCache.Retrieve(CACHE_NAME, objParams)
    If CacheEntry.Valid Then
        Set GetCachedProviders = CacheEntry.Value
    Else
        If objProvs Is Nothing Then
            Set objProvs = New DTOs.Providers
        End If
        Set GetCachedProviders = objProvs
    End If
    Exit Function
    
errHandler:
    Err.Raise Err.Number, "MapperGeneric.GetCachedProviders", Err.Description, Err.HelpFile, Err.HelpContext

End Function

Private Function GetCachedClinics(ByRef params As iParameters) As DTOs.Clinics
Dim CacheEntry As iCacheEntry
Dim objParams As ClinicSearchParam
Dim objClinics As DTOs.Clinics

    On Error GoTo errHandler
    
    Set objParams = params

    'Get the list if it exists
    Set CacheEntry = mCache.Retrieve(CACHE_NAME, objParams)
    If CacheEntry.Valid Then
        Set GetCachedClinics = CacheEntry.Value
    Else
        If objClinics Is Nothing Then
            Set objClinics = New DTOs.Clinics
        End If
        Set GetCachedClinics = objClinics
    End If

    Exit Function
    
errHandler:
    Err.Raise Err.Number, "MapperGeneric.GetCachedClinics", Err.Description, Err.HelpFile, Err.HelpContext

End Function

'Private Function CreateClinicCacheParams(ByRef Params As iParameters) As iParameters
'Dim objSearchNPParam As ClinicSearchParam
'Dim sFacilityNCID As String
'Dim sUserNCID As String
'Dim nUseClinicProviderTable As Integer
'Dim objParams As iParameters
'
'    On Error GoTo errHandler
'
'    'Generate key to determine whose frequently used list to retrieve
'    Set objSearchNPParam = Params
'    sFacilityNCID = objSearchNPParam.FacilityNCID
'    sUserNCID = objSearchNPParam.UserNCID
'    nUseClinicProviderTable = CInt(objSearchNPParam.UseClinicProviderTable)
'
'    Set objParams = New ParametersGeneric
'    objParams.Add "USERNCID" & sUserNCID, sUserNCID
'    If objSearchNPParam.IncludeUserClinicsOnly Then
'      objParams.Add "CACHE_KEY", CACHE_KEY_USERCLINICS
'    Else
'      objParams.Add "CACHE_KEY", CACHE_KEY_PROVIDERCLINICS
'    End If
'    objParams.Add "USRNCID" & sUserNCID, sUserNCID
'    objParams.Add "FACNCID" & sFacilityNCID, sFacilityNCID
'    objParams.Add "TABLE" & nUseClinicProviderTable, nUseClinicProviderTable
'    Set CreateClinicCacheParams = objParams
'
'    Set objSearchNPParam = Nothing
'
'    Exit Function
'
'errHandler:
'    Err.Raise Err.Number, "MapperGeneric.CreateClinicCacheParams", Err.Description, Err.HelpFile, Err.HelpContext
'
'End Function

'Private Function CreateProviderCacheParams(ByRef Params As iParameters) As iParameters
'Dim objSearchNPParam As ProviderSearchParam
'Dim sClinicNCID As String
'Dim nUsersOnly As Integer
'Dim objParams As iParameters
'
'    On Error GoTo errHandler
'
'    'Generate key to determine whose frequently used list to retrieve
'    Set objSearchNPParam = Params
'    sClinicNCID = objSearchNPParam.ClinicNCID
'    nUsersOnly = CInt(objSearchNPParam.ClinicUsersOnly)
'
'    Set objParams = New ParametersGeneric
'    objParams.Add "CLINICNCID" & sClinicNCID, sClinicNCID
'    objParams.Add "USERSONLY" & nUsersOnly, nUsersOnly
'    objParams.Add "CACHE_KEY", CACHE_KEY_PROVIDERS
'    Set CreateProviderCacheParams = objParams
'
'    Set objSearchNPParam = Nothing
'
'    Exit Function
'
'errHandler:
'    Err.Raise Err.Number, "MapperGeneric.CreateProviderCacheParams", Err.Description, Err.HelpFile, Err.HelpContext
'
'End Function


