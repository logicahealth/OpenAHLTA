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
Attribute VB_Name = "SecOps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum tsPrivAccess
    ts_Read = 1
    ts_Write = 2
    ts_Delete = 3
    ts_BreadGlass = 4
End Enum
    
Public mcolAllUsers      As Collection
Private mcolAllResources  As Collection
Private mcolAllPrivileges As Collection
Private mcolAllRoles      As Collection
Private mcolAllFacilities As Collection
Public mcolAllClinics    As Collection
Private mobjCrypto        As CCrypto
Public sFacilityName    As String


Public Property Get Crypto() As CCrypto: Set Crypto = mobjCrypto: End Property

Private Function GetAllUsers(oSQL As ICHCSII_SQL.ISqlOps) As Boolean

    Dim oUser As TheaterAdmin.User
    Set mcolAllUsers = New Collection
    'Dim oRes As User_Res
    'Dim oFac As User_Fac
    'Dim oPriv As privilege
    
    oSQL.Execute "select * from hems_users "
    Do While Not oSQL.EOF
        Set oUser = New TheaterAdmin.User
        oUser.NCID = "" & oSQL("User_Ncid")
        With oUser
            .User_Number = "" & oSQL("USER_NUMBER")
            .NCID = "" & oSQL("USER_NCID")
            .PrimaryRole = "" & oSQL("ROLE_NCID")
            .Unit_Number = "" & oSQL("UNIT_NUMBER")
            .LOGON_NAME = "" & oSQL("LOGON_NAME")
            .LOGON_PASSWD = "" & oSQL("LOGON_PASSWD")
            .USER_INITIALS = "" & oSQL("USER_INITIALS")
            .ACTIVE_FLAG = "" & oSQL("ACTIVE_FLAG")
            .SpecialyCode = Val("" & oSQL("SPECIALTY_NCID"))
            
            If "" & oSQL("signature_class") = "" Then
                .Signature_Class = 4
            Else
                .Signature_Class = Val("" & oSQL("signature_class"))
            End If
        End With
        
        
        mcolAllUsers.Add oUser, oUser.NCID

        oSQL.MoveNext
    Loop
    
    For Each oUser In mcolAllUsers
        Call GetProvider(oUser, oSQL)
        Call GetUserFacility(oUser, oSQL)
    Next oUser
End Function


Private Function GetProvider(oUser As User, oSQL As ICHCSII_SQL.ISqlOps) As Boolean
    On Error GoTo ErrHandler
    oSQL.Execute "select pcm_code, ien, Name, SSN,  clinic_ncid, signature_class, duty_phone_1, duty_phone_2, email" _
    & " from provider" _
    & " where ncid = " & oUser.NCID
    If Not oSQL.EOF Then
        oUser.PCM_Code = "" & oSQL("pcm_code")
        oUser.IEN = "" & oSQL("ien")
        oUser.ProviderName = "" & oSQL("name")
        oUser.SSN = oSQL("SSN")
        oUser.HomePhone = "" & oSQL("duty_phone_2")
        oUser.WorkPhone = "" & oSQL("duty_phone_1")
        oUser.EMail = "" & oSQL("email")
        oUser.ClinicNCID = oSQL("clinic_ncid")
        If "" & oSQL("signature_class") = "" Then
            oUser.Signature_Class = 4
        Else
            oUser.Signature_Class = Val("" & oSQL("signature_class"))
        End If
    Else
'        MsgBox "Unexpected Error: The Security Tables are not synchronized with the Clinical Table 'Provider'." _
'            & vbCrLf & "Please have the database administrator correct the error before running this application.", vbCritical, "Internal Error"
        '-- End JCB
    End If
ErrHandler:
End Function

Private Function GetAllFacilities(oSQL As ICHCSII_SQL.ISqlOps) As Boolean
    On Error GoTo haveerror
    
    Dim ofac As Facility
    Set mcolAllFacilities = New Collection
    Dim bFoundIt As Boolean
    
'__ need all the facilities from all sources:  Osh_ini_settings, ncids, facility_ncid_correlation,
    
    Do While Not oSQL.EOF
        Select Case oSQL("apkey")
        Case "FACILITY_NAME":        ofac.FacName = oSQL("APVALUE")
        Case "FACILITY_NCID":        ofac.FacNcid = oSQL("APVALUE")
        End Select
        oSQL.MoveNext
    Loop
    If ofac.FacName & ofac.FacNcid <> "" Then
        mcolAllFacilities.Add ofac, CStr(oSQL("fncid"))
    End If
    Set ofac = Nothing
    
    oSQL.Execute "select description as name, ncid from ncids where domainncid = " & NCID_FACILITY
    Do While Not oSQL.EOF
        On Error Resume Next
        Err.Clear
        Set ofac = mcolAllFacilities(oSQL("ncid"))
        If Err.Number <> 0 Then
            Set ofac = New Facility
            ofac.FacName = oSQL("Name")
            ofac.FacNcid = oSQL("ncid")
            mcolAllFacilities.Add ofac, CStr(oSQL("ncid"))
            Err.Clear
        Else
            '__ we found it, is it the osh one?
            If oSQL("ncid") = gobjLogon.FacilityNCID Then
                bFoundIt = True
            End If
        End If
        Set ofac = Nothing
        oSQL.MoveNext
    Loop
    
    oSQL.Execute "select apkey, apvalue " & _
                  " from osh_ini_settings " & _
                  " WHERE apsection = 'LOCAL_ENV'"
   Set ofac = New Facility
    
    
    oSQL.Execute "select facility_ncid " & _
                  " from FACILITY_NCID_Correlation "
    
    GetAllFacilities = True
    
haveerror:
        
End Function

Private Function GetDefaultFacility(oSQL As ICHCSII_SQL.ISqlOps) As Boolean
    oSQL.Execute "select apvalue from osh_ini_settings where apsection = 'LOCAL_ENV' and apkey = 'FACILITY_NAME'"
    If Not oSQL.EOF Then
        sFacilityName = "" & oSQL(0)
    End If
End Function

Private Function GetAllClinics(oSQL As ICHCSII_SQL.ISqlOps) As Boolean
 
    Dim oclinic As clinic
    Set mcolAllClinics = New Collection
    
    oSQL.Execute "Select * from Clinic where facility_ncid = " & cwobjLogonEnviron.FacilityNCID
    
    Do While Not oSQL.EOF
        If Val("" & oSQL("ncid")) > 0 Then
        On Error Resume Next
        Set oclinic = New clinic
        With oclinic
            .clinicName = oSQL("Name")
            '.FacilityName = oSQL("description")
            '.FacNcid = CStr(oSQL("fncid"))
            .ClinicNCID = CStr(oSQL("ncid"))
            .Addr = oSQL("STREET_ADDRESS")
            .City = oSQL("CITY")
            .State = oSQL("STATE")
            .Zip = oSQL("ZIP")
            .Phone = oSQL("TELEPHONE_NUMBER")
            .Dmis = oSQL("DMIS_ID_CODE")
            .meprs_code = oSQL("MEPRS_CODE")
            .meprs_ien = oSQL("MEPRS_IEN")
            .CType = oSQL("LOCATION_TYPE")
            .SickCall = oSQL("SICK_CALL")
            .WalkIn = oSQL("WALK_IN")
            .Inactive = Val("" & oSQL("INACTIVE"))
            .CIen = oSQL("IEN")
        End With
        mcolAllClinics.Add oclinic, oclinic.ClinicNCID
        oclinic.bDirty = False
        oclinic.bNew = False
        Set oclinic = Nothing
        End If
        oSQL.MoveNext
    Loop
    GetAllClinics = True

End Function

Public Sub Initialize()
    
    gobjComm.SetStatusText "Loading Users..."
        Call GetAllUsers(oSQL)

    gobjComm.SetStatusText "Loading Clinics..."
        Call GetAllClinics(oSQL)
End Sub

Public Function GetCollectionOfClinics() As Collection
  Set GetCollectionOfClinics = mcolAllClinics
End Function
Public Function GetCollectionOfFacilities() As Collection
  Set GetCollectionOfFacilities = mcolAllFacilities
End Function

Public Function GetAUserCollection(ByVal UserNCID As Long) As User
 Set GetAUserCollection = mcolAllUsers.Item(UserNCID)
End Function

Public Function SaveUser(oUser As User) As Boolean
    Dim SQL  As String
    Dim mCol As Collection
    Dim sFac As String
    Dim sDom As String
    Dim sTemp() As String
    Dim FirstName As String
    Dim LastName As String
    With oUser
    
    'SCR 101231 - LCI - 09/20/2006: Split first and last names in order to update MMI_NAME table
    sTemp = Split(.ProviderName, ", ", , vbTextCompare)
    LastName = sTemp(0)
    If UBound(sTemp) > 1 Then
        FirstName = sTemp(1)
    End If
    
    If .NewUser Then
            On Error Resume Next  ' - may have alread added to hems, provider
            Set mCol = New Collection
            SQL = "Insert into Hems_Users (user_number, unit_number, specialty_ncid,USER_INITIALS, user_ncid, role_ncid) values (" & _
                      .User_Number & ", " & _
                      .Unit_Number & ", " & _
                      SQLQuote(.SpecialyCode) & ", " & _
                      SQLQuote(.USER_INITIALS) & ", " & _
                      .NCID & ",322 )"
            mCol.Add SQL
            
            SQL = "Insert into Provider (PCM_Code, IEN, Unit_number, name, duty_Phone_1, duty_phone_2, ssn, specialty_code, email, ncid, signature_class) values (" & _
                   SQLQuote(.PCM_Code) & ", " & _
                            .IEN & ", " & _
                            .Unit_Number & ", " & _
                   SQLQuote(.ProviderName) & ", " & _
                   SQLQuote(.WorkPhone) & ", " & _
                   SQLQuote(.HomePhone) & ", " & _
                   SQLQuote(CleanSSN(.SSN)) & ", " & _
                   SQLQuote(.SpecialyCode) & ", " & _
                   SQLQuote(.EMail) & ", " & _
                            .NCID & "," & _
                            .Signature_Class & ")"
            mCol.Add SQL
            
            'SCR 101231 - LCI - 09/20/2006
            SQL = "Insert into MMI_NAME (last_name, first_name) values (" & _
                      LastName & ", " & _
                      FirstName & ")"
            mCol.Add SQL
            
            oSQL.ExecuteMulti mCol
            Set mCol = Nothing
            
            mcolAllUsers.Add oUser, CStr(oUser.NCID)
          
          
    Else  '- not new
         oSQL.Execute "Update Hems_Users  set " & _
                      "user_number = " & .User_Number & ", " & _
                      "unit_number = " & .Unit_Number & ", " & _
                      "specialty_ncid = " & .SpecialyCode & ", " & _
                      "USER_INITIALS = " & SQLQuote(.USER_INITIALS) & _
                      " Where " & _
                      "user_ncid = " & .NCID
                  
        sFac = ""
        
        oSQL.Execute "Update Provider  set " & _
                     "PCM_Code = '" & .PCM_Code & "', " & _
                     "IEN = " & .IEN & ", " & _
                     "Unit_Number = " & .Unit_Number & ", " & _
                     "Name = " & SQLQuote(.ProviderName) & ", " & _
                     "Duty_Phone_1 = " & SQLQuote(.WorkPhone) & ", " & _
                     "Duty_Phone_2 = " & SQLQuote(.HomePhone) & ", " & _
                     "SSN = " & SQLQuote(CleanSSN(.SSN)) & ", " & _
                     sFac & _
                     "Specialty_code = " & SQLQuote(.SpecialyCode) & ", " & _
                     "Email = " & SQLQuote(.EMail) & ", " & _
                     "Signature_class = " & .Signature_Class & " " & _
                     " Where " & _
                     "ncid = " & .NCID
'<< SCR 24175  change provider name and also change in clinic table
       oSQL.Execute "Update clinic_Provider set provider_name = " & SQLQuote(.ProviderName) _
                & ", Signature_class = " & .Signature_Class _
                 & " where clinician_ncid = " & .NCID
                     
       'SCR 101231 - LCI - 09/20/2006
       oSQL.Execute "Update mmi_name set last_name = " & SQLQuote(LastName) _
                & ", first_name = " & SQLQuote(FirstName) _
                 & " where unit_number = " & .Unit_Number
                                           
    End If
    
    End With
    SaveUser = True
End Function

Public Function GetHEMSUser(ByVal sNCID As String) As User
    Dim oUser As User

    'Dim oRes  As User_Res
    'Dim oFac  As User_Fac
    Dim count As Integer
    'Dim oPriv As privilege
    
    oSQL.Execute "select User_Number , Unit_Number, Specialty_Ncid,user_initials" _
    & " from hems_users" _
    & " where  user_ncid = " & sNCID
    If oSQL.EOF Then
        Exit Function
    End If
    
    Set oUser = New TheaterAdmin.User
    oUser.NCID = sNCID
    oUser.Timer = 30
    oUser.Unit_Number = "" & oSQL("unit_number")
    oUser.User_Number = "" & oSQL("user_number")
    oUser.USER_INITIALS = "" & oSQL("USER_INITIALS")
    oUser.SpecialyCode = IIf(IsNull(oSQL("Specialty_Ncid")), 0, oSQL("Specialty_Ncid"))


    Call GetProvider(oUser, oSQL)
    Call GetUserFacility(oUser, oSQL)
    

    Set GetHEMSUser = oUser
    Set oUser = Nothing

End Function

Private Function GetUserFacility(oUser As User, oSQL As ICHCSII_SQL.ISqlOps) As Boolean
    Dim ofac As User_Fac
'<< 8.8.02 SHAW:  added constraint of fac ncid
    oSQL.Execute "select distinct facility_ncid, clinic_ncid, f.description as facility_name, c.description as clinic_name" _
    & " from user_clinic u, ncids f, ncids c" _
    & " where f.ncid = facility_ncid" _
    & " and c.ncid = clinic_ncid " _
    & " and u.user_ncid = " & oUser.NCID _
    & " and u.facility_ncid = " & cwobjLogonEnviron.FacilityNCID
    Do While Not oSQL.EOF
        On Error Resume Next
        Set ofac = New User_Fac
        ofac.clinicName = oSQL("clinic_name")
        ofac.FacilityName = oSQL("facility_name")
        ofac.ClinicNCID = oSQL("clinic_ncid")
        ofac.FacilityNCID = oSQL("facility_ncid")
        oUser.Facilities.Add ofac, CStr(ofac.ClinicNCID)
        oSQL.MoveNext
    Loop

End Function

Public Function DuplicateUID(ByVal sUID As String, ByVal resId As Long, ByVal sUserNCID As String) As Boolean
'- be sure there is not duplicate uid
' return true if one is found
    oSQL.Execute "select count(*) from user_resources where res_id = " & resId _
      & " and user_name = '" & sUID & "' and not user_ncid = " & sUserNCID
    If oSQL(0) > 0 Then
        DuplicateUID = True
        MsgBox "The logon name is already in use. Please select another.", vbCritical, "User Logon Error"
    End If
End Function

Private Sub Class_Initialize()
    Set mobjCrypto = New CCrypto
End Sub

'/// special call for maint use only
Public Sub RefreshAllUsers()
    Dim U As User
    For Each U In mcolAllUsers
        SaveUser U
    Next U
End Sub


Public Function SaveClinic(oClin As clinic) As Boolean
    Dim SQL As String
    Dim bNewOne As Boolean
    On Error GoTo ErrHandler
    
    With oClin
        If .ClinicNCID = "" Then
                    
        '12/5/01 - 19412 - Walker - Use standard interface to create provider clinics
            'Dim oSQL2 As ICHCSII_SQL.ISqlOpsEx
            'Set oSQL2 = New GEMS_SQL.SqlOps
                        
            '104841 Use the same sequence as the providers and patients.
            .ClinicNCID = oSQL.GetNextID(exUNIT_NUMBER)
            '.ClinicNCID = oSQL.GetNextNcid()
            
            .CIen = oSQL.GetNextID(esIEN)
            .meprs_code = oSQL.GetNextID()
            .meprs_ien = .ClinicNCID
         '- this should create the clinic and put it in the facility (tables: clinic, facilityncidcorr, ncids)
            If oSQL.ValidateFacility(.FacNcid, True, .ClinicNCID) Then
            End If
            bNewOne = True
        Else
            '- update this one
        End If
    
        SQL = "update clinic set " _
        & " Name               = " & SQLQuote(.clinicName) _
        & ",STREET_ADDRESS     = " & SQLQuote(.Addr) _
        & ",CITY               = " & SQLQuote(.City) _
        & ",STATE              = " & SQLQuote(.State) _
        & ",ZIP                = " & SQLQuote(.Zip) _
        & ",TELEPHONE_NUMBER   = " & SQLQuote(.Phone) _
        & ",DMIS_ID_CODE       = " & SQLQuote(.Dmis) _
        & ",MEPRS_CODE         = " & SQLQuote(.meprs_code) _
        & ",MEPRS_IEN          = " & SQLQuote(.meprs_ien) _
        & ",LOCATION_TYPE      = " & SQLQuote(.CType) _
        & ",SICK_CALL          = " & SQLQuote(.SickCall) _
        & ",WALK_IN            = " & SQLQuote(.WalkIn) _
        & ",INACTIVE           = " & SQLQuote(.Inactive) _
        & ",IEN                = " & SQLQuote(.CIen) _
        & " where ncid =          " & .ClinicNCID
        
        oSQL.Execute SQL
        
        '__ update the entry in ncids, too
        oSQL.Execute "update ncids set description = " & SQLQuote(.clinicName) & " where ncid = " & .ClinicNCID
        
        
    End With
    
'<< special case: this is the base clinic (the one in osh_ini_settings table)
'   so we have to update that name, too
    If oClin.ClinicNCID = gobjLogon.NursingDivisionNCID Then
        oSQL.Execute "update osh_ini_settings set apvalue = " & SQLQuote(oClin.clinicName) _
                    & " where apsection = 'LOCAL_ENV' and apkey = 'CLINIC_NAME'"
        'gobjLogon.NursingDivisionName = oclin.clinicName   '___ read only
    End If
    
    oClin.bDirty = False
    
    If oClin.bNew Then
        mcolAllClinics.Add oClin, oClin.ClinicNCID
'<< SCR #25474;   Developer: Shaw 07/10/2002 09:08 AM

        '___ This is a new clinic so associate the Unassigned provider to it
'<< scr
        Call AssignUnassignedToClinic(oClin.ClinicNCID)
'>>
    
    End If
    
    SaveClinic = True
    Exit Function
ErrHandler:

End Function

Private Function AssignUnassignedToClinic(ByVal ClinicNCID As String) As Boolean
    Dim sUassignedNCID  As String       '__ the ncid for pseudo-provider ".Unassgined"
    Dim sIEN            As String
    On Error GoTo ErrHandler
    
'__ get the .unassigned ncid from hems_users.
    oSQL.Execute "select user_ncid from hems_users where lower(logon_name) = '.unassigned'"
    If Not oSQL.EOF Then
        sUassignedNCID = "" & oSQL("User_Ncid")
    Else
        AssignUnassignedToClinic = False
        Exit Function
    End If
    
    oSQL.Execute "Select ien from provider where ncid =" & sUassignedNCID
    If Not oSQL.EOF Then
        sIEN = "" & oSQL("ien")
        oSQL.Execute "update provider set clinic_ncid = " & ClinicNCID _
            & ", facility_ncid = " & gobjLogon.FacilityNCID _
            & ", facilityncid = " & gobjLogon.FacilityNCID _
            & ", name = '.Unassigned'" _
            & " WHERE NCID = " & sUassignedNCID
    Else
        sIEN = oSQL.GetNextID(esIEN)
        oSQL.Execute "Insert into provider (ncid,pcm_code,ien,unit_number,name,clinic_ncid,facility_ncid,facilityncid,specialty_code,ssn,signature_class)" _
        & "values(" & sUassignedNCID & ",'.5-A0101'," & sIEN & "," & sUassignedNCID & ",'.UNASSIGNED'," & ClinicNCID & "," & gobjLogon.FacilityNCID & "," & gobjLogon.FacilityNCID & ",322,'001010001',4)"
    End If

'__ 4. be sure he's in the user_clinic and clinic_provider tables for current clinic
    On Error Resume Next
    oSQL.Execute "insert into user_clinic  (CLINIC_NCID, FACILITY_NCID, USER_NCID )" _
            & " Values (" & ClinicNCID & "," & gobjLogon.FacilityNCID & "," & sUassignedNCID & ")"

    oSQL.Execute "insert into clinic_provider (CLINICIAN_NCID,CLINIC_NCID,FACILITY_NCID,PROVIDER_IEN,PROVIDER_NAME,USER_IEN) values ( " _
        & sUassignedNCID & "," & ClinicNCID & "," & gobjLogon.FacilityNCID & "," & sIEN & ",'.Unassigned'," & sIEN & ")"
        'everyone defaults to this clinic
    Exit Function
ErrHandler:
    gobjShared.ShowVBError Err.Number, Err.Description, "SecOps.AssignUnassignedToClinic", "Error in Theater Admin Module", vbExclamation
    Resume Next

End Function



Friend Sub UpdateUserClinics(oUser As User, oUFac As User_Fac, ByVal sAction As String)
    On Error GoTo ErrHandler
    
    Select Case sAction
    Case "ADD"
        '--- delete
        oSQL.Execute "delete from user_clinic where user_ncid = " & oUser.NCID _
         & " and clinic_ncid = " & oUFac.ClinicNCID & " and FACILITY_NCID = " & oUFac.FacilityNCID
         '--- insert
        oSQL.Execute "insert into User_Clinic (user_ncid, Facility_ncid, Clinic_Ncid) values (" _
         & oUser.NCID & ", " & oUFac.FacilityNCID & ", " & oUFac.ClinicNCID & ")"

        '--- delete
        oSQL.Execute "delete from clinic_provider where clinician_ncid = " & oUser.NCID _
         & " and clinic_ncid = " & oUFac.ClinicNCID & " AND FACILITY_NCID = " & oUFac.FacilityNCID
        '--- insert
        oSQL.Execute "insert into clinic_provider (facility_ncid ,Clinic_NCID,clinician_ncid, " _
        & "provider_name,provider_ien,user_ien) values (" _
        & oUFac.FacilityNCID & ", " & oUFac.ClinicNCID & ", " & oUser.NCID & ", " _
        & SQLQuote(oUser.ProviderName) & "," & oUser.IEN & "," & oUser.User_Number & ")"

        If Val(oUser.ClinicNCID) = 0 Then
            '- add this as the first/primary clinic
           oSQL.Execute "update provider set clinic_ncid = " & oUFac.ClinicNCID _
                & " where ncid = " & oUser.NCID
        End If

    Case "REMOVE"
        
        oSQL.Execute "delete from user_clinic where user_ncid = " & oUser.NCID _
         & " and clinic_ncid = " & oUFac.ClinicNCID
        oSQL.Execute "delete from clinic_provider where clinician_ncid = " & oUser.NCID _
         & " and clinic_ncid = " & oUFac.ClinicNCID
        oSQL.Execute "update provider set clinic_ncid = NULL " _
         & " where clinic_ncid = " & oUFac.ClinicNCID _
         & " and ncid = " & oUser.NCID

    End Select
'<< scr 49320 tell everyone to update their clinic lists
'<< scr 50496  forgot to add the 'false' as the final param so it was opening the apps
    gobjComm.Message cwiGENERIC, "UPDATECLINICS", NCID_CLINIC_SCHEDULE_APPLICATION, NCID_THEATER_ADMIN_APPLICATION, False
    gobjComm.Message cwiGENERIC, "UPDATECLINICS", NCID_TEMPLATE_MANAGEMENT, NCID_THEATER_ADMIN_APPLICATION, False

    Exit Sub
ErrHandler:
    gobjShared.ShowVBError Err.Number, Err.Description, "SecOps.UpdateUserclinics", "Error in Theater Admin Module", vbExclamation
    Resume Next
End Sub

