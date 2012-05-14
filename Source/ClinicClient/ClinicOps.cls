VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClinicOps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'''****************************************************************************************
'-- if you've got it, don't go get it again.
'''Note:       Provider procedures can operate on three tables
'''            Provider
'''               Contains all providers.
'''               Generally should be unique but may have duplicates because it is
'''               populated based on FACILITY_NCID and IEN
'''            Clinic_Provider
'''               Provider 1-n Clinics
'''            Hems_users
'''               Contains providers who are also system users - can login to CHSCII
'''
'''****************************************************************************************
''''''''''''''' Constants Section '''''''''''''''''''''''''''''''''''''''''
'Maint:
'WRogers 9/25/2003
' Provider theater only path for SQL string in GetUserClinics() and GetProviderClinics().
'
'WRogers 10/1/2003
' SCR #39383 - IB/WAM and PGUI both use this workload calculation.
' The workload type for Provider Clinic Appt can some times be blank, then we use
' the clinic_appt type to decide.


Private Const msMODULE_NAME         As String = "ClinicOps."

Private Enum penmColType
   enmProvider
   enmClinic
End Enum

''''''''''''''' Variables Section '''''''''''''''''''''''''''''''''''''''''

Public mbUseClinicProviderTable     As Boolean

Private bAddingToCollection                     As Boolean  '''<bjr113001 SCR18815

''''''''''''''' Objects Section '''''''''''''''''''''''''''''''''''''''''

Private mrst                                    As adodb.Recordset

Public Function GetClinic(rsNCID As String) As Clinic
   Const sPROC_NAME  As String = "GetClinic"
   
   Dim sStmt         As String
   Dim colClinic     As Collection
   
On Error GoTo errHandler

    '''SCR #31046;   Developer: Alexander Alberstein 12/26/2002 03:32 PM
    If Not (rsNCID = "") Then
      sStmt = "Select c.Facility_NCID, c.NCID, c.IEN, c.Name, c.Meprs_IEN, m.Code as Meprs_Code, " _
          & " c.WALK_IN, c.SICK_CALL, c.STREET_ADDRESS, c.BUILDING_NUMBER," _
          & " c.City, c.State, c.ZIP, c.DMIS_ID_CODE," _
          & " m.Description as MeprsDescription, c.Location_Type, c.BDM_DEF_PHARMACY_ID, c.DEFAULT_WORKLOAD_TYPE" _
          & " From Clinic c, Meprs_Code m" _
          & " Where c.NCID = " & rsNCID _
          & " and m.IEN = c.Meprs_IEN" _
          & " and m.facility_ncid = c.facility_ncid" _
          & " Order by c.Name"
      '''<SCR #30152;   Developer: Alexander Alberstein 01/10/2003 9:00 PM
      Set colClinic = ClinicClient_colGetClinics(sStmt)
      If colClinic.Count > 0 Then
         Set GetClinic = colClinic(1)
      End If
   End If
   
Exit Function
errHandler:

    Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, vbExclamation)
   'Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, rsNCID)
   
End Function

Public Function GetAllClinics(rsMTFFacilityNCID As String, Optional boolIncludeNonMTF As Boolean = True) As Collection
   Const sPROC_NAME  As String = "GetAllClinics"
   
   Dim sStmt           As String
    
On Error GoTo errHandler

   '''SCR #31046;   Developer: Alexander Alberstein 12/26/2002 03:32 PM
   sStmt = "Select c.Facility_NCID, c.NCID, c.IEN, c.Name, c.Meprs_IEN, m.Code as Meprs_Code, " _
        & " c.WALK_IN, c.SICK_CALL, c.STREET_ADDRESS, c.BUILDING_NUMBER," _
        & " c.City, c.State, c.ZIP, c.DMIS_ID_CODE," _
        & " m.Description as MeprsDescription, c.Location_Type, c.BDM_DEF_PHARMACY_ID, c.DEFAULT_WORKLOAD_TYPE" _
        & " From Clinic c, Meprs_Code m" _
        & " Where c.facility_ncid = " & rsMTFFacilityNCID _
        & " and m.facility_ncid = c.facility_ncid" _
        & " and m.IEN = c.Meprs_IEN"
        'RFoster Begin:SCR #24157;   Developer: Robert Foster 06/20/2002 12:58 PM
        'RFoster Note: Function Changed: in ClinicClient -GetAllClinics; in Clinicam - mbLoadClinics      Impact of Change: If the additional input parameter is false, then the non-mtf clinics are not returned.
        If Not boolIncludeNonMTF Then
            sStmt = sStmt & " and c.LOCATION_TYPE <> 'MCP NON-MTF'"
        End If
        'RFoster End: SCR #24157;
        sStmt = sStmt & " and (c.inactive <> 1 or inactive is null)" _
        & " Order by c.Name"
   
   '''<SCR #30152;   Developer: Alexander Alberstein 01/10/2003 9:00 PM
   Set GetAllClinics = ClinicClient_colGetClinics(sStmt)
   
Exit Function
errHandler:

    Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, vbExclamation)
   'Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")
   
   Set GetAllClinics = New Collection
   
End Function

Public Function GetClinicProvidersByRole _
      (ByVal vsClinicNCID As String, Optional ByVal sRole As String = "") As Collection
      
   Set GetClinicProvidersByRole = ClinicClient_colGetClinicProviders(vsClinicNCID, True, sRole, True)
   
End Function

Public Function GetClinicsWorkload _
      (ByVal sMTFFacilityNCID As String, ByRef colClinics As Collection) As Boolean
'''************************************************************************
'''Called By:
'''
'''Arguments:
'''   sMTFFacilityNCID
'''      NCID of the Facility associated with specified clinics
'''   colClinics
'''      Collection of the clinic objects requested to be populated with workload information
'''
'''Description:      Retrieve workload information for requested clinics and populate
'''                  Workloads properties of all clinic objects in specified collection
'''
'''
'''Note:             'SCR #31046;   Developer: Alexander Alberstein 12/27/2002 16:32 PM
'''
'''Warning:          All specified clinics are assumed to be associated with with specified facility
'''
'''Updates Globals:
'''
'''Returns:       If successful
'''                  true
'''               Else
'''                  false
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "GetClinicsWorkload"

   Dim bErr             As Boolean
   Dim objClinic        As Clinic
   Dim objSQLExp        As SQLText
   Dim sClause          As String
   
On Error GoTo ErrHnd

20    If (colClinics.Count > 0) Then
         '''Build SQL statement for retrieving workload records associated with specified facility
100      Set objSQLExp = New SQLText
200      With objSQLExp
            '''Columns
210         Call .AddSQLSelect("Distinct c.Appt_Code")
220         Call .AddSQLSelect("a.Description")
230         Call .AddSQLSelect("c.Workload_Type")
240         Call .AddSQLSelect("c.Appt_Duration")
250         Call .AddSQLSelect("c.Clinic_NCID")
            '''Tables
300         Call .AddSQLFrom("CLINIC_APPT_TYPE c")
320         Call .AddSQLFrom("APPT_TYPE a")
            '''Joint
400         Call .AddSQLWhere("c.Appt_Code = a.Code")
420         Call .AddSQLWhere("c.FACILITY_NCID = a.FACILITY_NCID")
            '''Filter
430         Call .AddSQLWhere("c.Facility_NCID = " & sMTFFacilityNCID)
            '''If limited subset of clinics
500         If colClinics.Count < 21 Then
               '''Implement In clause to improve performance
520            For Each objClinic In colClinics
530               sClause = sClause & ", " & objClinic.NCID
               Next objClinic
560            Call .AddSQLWhere("c.Clinic_NCID In (" & Mid$(sClause, 2) & ")")
            End If
            '''Sort
700         Call .AddSQLOrder("c.Clinic_NCID")
710         Call .AddSQLOrder("c.Appt_Code")
            '''<SCR #30152;   Developer: Alexander Alberstein 01/13/2003 2:00 PM
800         Set mrst = ClinicClient_objDAS.OpenRecordset(.GetSQLStatement, adOpenStatic)

         End With

         '''Instantiate and populate workload collections of all clinic objects
900      For Each objClinic In colClinics
            '''Always return valid workloads collection even if no workload was found for this clinic
920         Set objClinic.Workloads = New Collection
            '''<DSN Skip Non-MTF clinics - verify: can not have workload
940         'If Not objClinic.LocationType = "MCP NON-MTF" Then
960            If Not M_bPopulateWorkload("Clinic_NCID", objClinic.NCID, objClinic.Workloads) Then
                  bErr = True
               End If
            'End If
            '>
         Next objClinic

990      If Not mrst Is Nothing Then
992         Call mrst.Close
         End If
      End If

999   GetClinicsWorkload = Not bErr
   
Exit Function
ErrHnd:
   GetClinicsWorkload = False
   
   Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, vbExclamation)
   'Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl)

End Function

Public Function GetClinicWorkloadType _
      (ByVal sFacilityNCID, ByVal sClinicNCID As String, ByVal sProviderNCID As String, ByVal sApointmentCode As String, _
       Optional ByRef sErrMsg As String, Optional ByVal bPopMsgOnErr As Boolean = False _
      ) As String
    Const sPROC_NAME     As String = "GetClinicWorkloadType"
    
    Dim sResult             As String
    Dim objSQLExp           As SQLText
    Dim rst                 As adodb.Recordset
    Dim sTest               As String
    Dim sProviderType       As String
    Dim sClinicType         As String
    Dim sDefaultType        As String
    Dim sSQL       As String
    Dim bResult             As Boolean
    On Error GoTo ErrHnd
    sErrMsg = ""
    sResult = "N"
    
    ' Begin:SCR #44837;   Developer:MAG  11/06/2003
    If gobjShared.IsAppMode(modeTheater) Then
      Exit Function
    End If
    ' End: SCR #44837;
    
    'R. Foster SCR #30519;   Developer: Robert Foster 02/27/2003 03:55 PM
    'R. Foster Note: Function Changed: GetClinicWorkloadType      Impact of Change: Added Facility NCID to input param's added logic to capture instances where the provider's workloadType is not present to default up the tree and capture the previous value.
    
    sSQL = "SELECT CLINIC.FACILITY_NCID, CLINIC.NCID AS CLINIC_NCID, 0 AS PROVIDER_NCID, "
    sSQL = sSQL & "'DEFAULT' AS APPT_CODE, CLINIC.DEFAULT_WORKLOAD_TYPE AS WORKLOAD_TYPE, "
    sSQL = sSQL & "'DEFAULT' AS LEVELType "
    sSQL = sSQL & "From Clinic "
    sSQL = sSQL & "Where Clinic.FACILITY_NCID = " & sFacilityNCID & " And Clinic.NCID = " & sClinicNCID & " "
    sSQL = sSQL & "Union "
    sSQL = sSQL & "SELECT CLINIC_APPT_TYPE.FACILITY_NCID, CLINIC_APPT_TYPE.CLINIC_NCID, "
    sSQL = sSQL & "0 AS PROVIDER_NCID, CLINIC_APPT_TYPE.APPT_CODE, "
    sSQL = sSQL & "CLINIC_APPT_TYPE.WORKLOAD_TYPE,  'CLINIC' AS LEVELType "
    sSQL = sSQL & "From CLINIC_APPT_TYPE "
    sSQL = sSQL & "WHERE CLINIC_APPT_TYPE.FACILITY_NCID=" & sFacilityNCID & " AND "
    sSQL = sSQL & "CLINIC_APPT_TYPE.CLINIC_NCID=" & sClinicNCID & " AND "
    sSQL = sSQL & "CLINIC_APPT_TYPE.APPT_CODE= '" & sApointmentCode & "' "
    sSQL = sSQL & "Union "
    sSQL = sSQL & "SELECT PROVIDER_APPT_TYPE.FACILITY_NCID, PROVIDER_APPT_TYPE.CLINIC_NCID, "
    sSQL = sSQL & "PROVIDER_APPT_TYPE.PROVIDER_NCID, PROVIDER_APPT_TYPE.APPT_CODE, "
    sSQL = sSQL & "PROVIDER_APPT_TYPE.WORKLOAD_TYPE,  'PROVIDER' AS LEVELType "
    sSQL = sSQL & "From PROVIDER_APPT_TYPE "
    sSQL = sSQL & "WHERE PROVIDER_APPT_TYPE.FACILITY_NCID=" & sFacilityNCID & " AND "
    sSQL = sSQL & "PROVIDER_APPT_TYPE.CLINIC_NCID=" & sClinicNCID & " "
    sSQL = sSQL & "AND PROVIDER_APPT_TYPE.PROVIDER_NCID= " & sProviderNCID & " "
    sSQL = sSQL & "AND PROVIDER_APPT_TYPE.APPT_CODE= '" & sApointmentCode & "' "
    If ClinicClient_bOpenRecordSet(sSQL, rst) Then
        rst.MoveFirst 'Get the Providers Workload
        rst.Find "LEVELType='PROVIDER'"
        If Not rst.EOF Then
            sProviderType = rst.Fields("WORKLOAD_TYPE").Value & ""
        Else
            sProviderType = ""
        End If
        rst.MoveFirst 'Get the Clinic's Appointment Workload
        rst.Find "LEVELType='CLINIC'"
        If Not rst.EOF Then
            sClinicType = rst.Fields("WORKLOAD_TYPE").Value & ""
        Else
            sClinicType = ""
        End If
        rst.MoveFirst 'Get the Clinic's Default Workload
        rst.Find "LEVELType='DEFAULT'"
        If Not rst.EOF Then
            sDefaultType = rst.Fields("WORKLOAD_TYPE").Value & ""
        Else
            sDefaultType = "N"  'If the record is missing, then default to No-Count
            'This should never happen, you have to be in the clinic profile to schedule appointments
            'and to do that you have to ,ake a choice of Count or No-Count.
        End If
        'If bResult is False here, then it's a No-Count Clinic
        bResult = LibStr_bTxtEql(sDefaultType, "C")
        
        If bResult Then
        
            If Len(sProviderType) = 0 Then
                bResult = bResult And (Len(sClinicType) = 0 Or LibStr_bTxtEql(sClinicType, "C"))
            ElseIf bResult And LibStr_bTxtEql(sProviderType, "C") Then
                bResult = True
            Else
                bResult = False
            End If
        End If
        
        If bResult Then
            sResult = "NC"
        Else
            sResult = "N"
        End If
    End If

GetClinicWorkloadType = sResult
Exit Function
ErrHnd:
   GetClinicWorkloadType = ""
   sErrMsg = Err.Description
   
   Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, vbExclamation)
   
   'Call LibApp_Error _
   (Err.Number, sErrMsg, msMODULE_NAME, sPROC_NAME, _
    sClinicNCID & ", " & sProviderNCID & ", " & sApointmentCode & vbCrLf & gsLINE_ & Erl _
   )
   
End Function

Public Function GetProviderExt(ByVal sProviderNCID As String) As mclsProviderExt
   Const sPROC_NAME  As String = "GetProviderExt"
   
On Error GoTo ErrHnd

   Set GetProviderExt = ClinicClient_objProviderExt(sProviderNCID)
   
Exit Function
ErrHnd:
   Set GetProviderExt = Nothing
   Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, vbExclamation)
'   Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, sProviderNCID)

End Function

Public Function GetProvidersWorkload _
   (ByVal sMTFFacilityNCID As String, ByVal sClinicNCID As String, ByRef colProviders As Collection) _
    As Boolean
'''************************************************************************
'''Called By:
'''
'''Arguments:
''' sMTFFacilityNCID
'''   NCID of the Facility associated with specified providers
''' sClinicNCID
'''   ID of the Clinic associated with specified providers
''' colProviders
'''   Collection of the provider objects requested to be populated with workload information
'''
'''Description:      Retrieve workload information for requested providers and populate
'''                  Workloads properties of all provider objects in specified collection
'''
'''
'''Note:             'SCR #31046;   Developer: Alexander Alberstein 12/28/2002 16:32 PM
'''
'''Warning:          All specified providers are assumed to be associated with with specified facility
'''                  ans Clinic
'''Updates Globals:
'''
'''Returns:       If successful
'''                  true
'''               Else
'''                  false
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "GetProvidersWorkload"

   Dim bResult          As Boolean
   Dim objProvider      As Provider
   Dim objSQLExp        As SQLText
   Dim sClause          As String
   
On Error GoTo ErrHnd

      ' Begin:SCR #44837;   Developer:  11/06/2003
      If gobjShared.IsAppMode(modeTheater) Then
        Exit Function
      End If
      ' End: SCR #44837;
      

20    If colProviders.Count > 0 Then
         '''Build SQL statement for retrieving workload records associated with specified clinic
100      Set objSQLExp = New SQLText
200      With objSQLExp
            '''Columns
210         Call .AddSQLSelect("Distinct p.Appt_Code")
220         Call .AddSQLSelect("a.Description")
230         Call .AddSQLSelect("p.Workload_Type")
240         Call .AddSQLSelect("p.Appt_Duration")
250         Call .AddSQLSelect("p.PROVIDER_NCID")
            '''Tables
300         Call .AddSQLFrom("PROVIDER_APPT_TYPE p")
320         Call .AddSQLFrom("APPT_TYPE a")
            '''Joint
400         Call .AddSQLWhere("p.Appt_Code = a.Code")
420         Call .AddSQLWhere("p.Facility_NCID = a.FACILITY_NCID")
            '''Filter
460         Call .AddSQLWhere("p.Facility_NCID = " & sMTFFacilityNCID)
470         Call .AddSQLWhere("p.Clinic_NCID = " & sClinicNCID)
            '''If limited subset of prividers
500         If colProviders.Count < 21 Then
               '''Implement In clause to improve performance
520            For Each objProvider In colProviders
530               sClause = sClause & ", " & objProvider.ClinicianNCID
               Next objProvider
560            Call .AddSQLWhere("p.PROVIDER_NCID In (" & Mid$(sClause, 2) & ")")
            End If
            '''Sort
700         Call .AddSQLOrder("p.PROVIDER_NCID")
710         Call .AddSQLOrder("p.Appt_Code")
         End With
      
         '''<SCR #30152;   Developer: Alexander Alberstein 01/13/2003 2:00 PM
900      bResult = M_bPopulateProvidersWorkload(objSQLExp.GetSQLStatement, colProviders)
      Else
         bResult = True
      End If
      
999   GetProvidersWorkload = bResult
   
Exit Function
ErrHnd:
   GetProvidersWorkload = False
   Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME & gsLINE & Erl, sPROC_NAME, vbExclamation)
   'Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl)

End Function

Public Function GetUserClinics(ByVal vsMTFFacilityNCID As String, ByVal vsUserNCID As String, Optional ByVal vbGetIBWAClinics As Boolean = False) As Collection
   Const sPROC_NAME  As String = "GetUserClinics"
   
   Dim oCmd         As CHCSII_Command.CHCSII_Cmd
   Dim sStmt        As String
   Dim oRs          As adodb.Recordset
   Dim sProcName    As String
   Dim colResult        As Collection
   Dim oClinic       As Clinic
   
On Error GoTo errHandler
    
    If gobjShared.IsAppMode(enumCSCSIIAppMode.modeTheater) Then
       sStmt = "Select c.Facility_NCID, c.NCID, c.IEN, c.Name, c.Meprs_IEN, m.Code as Meprs_Code, " _
        & " c.WALK_IN, c.SICK_CALL, c.STREET_ADDRESS, c.BUILDING_NUMBER," _
        & " c.City, c.State, c.ZIP, c.DMIS_ID_CODE," _
        & " m.Description as MeprsDescription, c.Location_Type " _
        & " From Clinic c, User_Clinic uc, Meprs_Code m" _
        & " Where uc.User_NCID = " & vsUserNCID _
        & " and c.ncid = uc.Clinic_NCID and m.IEN = c.Meprs_IEN" _
        & " and c.facility_ncid = " & vsMTFFacilityNCID _
        & " and c.facility_ncid = m.facility_ncid" _
        & " and (c.inactive <> 1 or inactive is null) " _
        & " Order by c.Name"

    ElseIf gobjShared.IsAppMode(enumCSCSIIAppMode.modeCDR) And _
        Not gobjShared.IsAppMode(enumCSCSIIAppMode.modeITT) Then
        
' SCR #35883;   Developer: MAG 03/03/2004
'Stored Procedure #2
        Set colResult = New Collection
        
        sProcName = "CLINIC_CLIENT_PERFORMANCE_PKG.Get_User_Clinics"
        
        Set oCmd = New CHCSII_Command.CHCSII_Cmd
        
        With oCmd
            .SPName = sProcName
            .AddParam False, False, "textString", vsUserNCID
            .AddParam False, False, "textString", vsMTFFacilityNCID
            .AddParam True, True, "", ""   '--- for Output Parameter
        End With
        
        Set oRs = ClinicClient_objDAS.OpenRecordsetSP(oCmd)
        
        If Not oRs Is Nothing Then
          If oRs.RecordCount > 0 Then
            With oRs
              .MoveFirst
              Do While Not .EOF
                Set oClinic = New Clinic
                Call oClinic.F_Deserialize(.Fields)
                On Error Resume Next
                colResult.Add oClinic, oClinic.NCID
                On Error GoTo errHandler
                Call .MoveNext
              Loop
              Call .Close
            End With
          End If
        End If
        
        If vbGetIBWAClinics Then
          sStmt = "Select c.Facility_NCID, c.NCID, c.IEN, c.Name, c.Meprs_IEN, m.Code as Meprs_Code, " _
            & " c.WALK_IN, c.SICK_CALL, c.STREET_ADDRESS, c.BUILDING_NUMBER," _
            & " c.City, c.State, c.ZIP, c.DMIS_ID_CODE," _
            & " m.Description as MeprsDescription, c.Location_Type, c.BDM_DEF_PHARMACY_ID, c.DEFAULT_WORKLOAD_TYPE" _
            & " From Clinic c, Meprs_Code m" _
            & " Where m.IEN = c.Meprs_IEN" _
            & " and c.facility_ncid = " & vsMTFFacilityNCID _
            & " and c.facility_ncid = m.facility_ncid" _
            & " and (c.inactive <> 1 or inactive is null) " _
            & " and upper(c.Name) Like 'IBWA%'" _
            & " Order by c.Name"
          
          If ClinicClient_bOpenRecordSet(sStmt, oRs) Then
            With oRs
              .MoveFirst
              Do While Not .EOF
                Set oClinic = New Clinic
                Call oClinic.F_Deserialize(.Fields)
                On Error Resume Next
                colResult.Add oClinic, oClinic.NCID
                On Error GoTo errHandler
                Call .MoveNext
              Loop
              Call .Close
            End With
          End If
        
        End If
        
        Set oRs = Nothing
        Set GetUserClinics = colResult
        
        Exit Function
    
    Else
    
   '''SCR #24273;   Developer: Alexander Alberstein 05/31/2002 02:56 PM
   '''SCR #31046;   Developer: Alexander Alberstein 12/26/2002 03:32 PM
   sStmt = "Select c.Facility_NCID, c.NCID, c.IEN, c.Name, c.Meprs_IEN, m.Code as Meprs_Code, " _
        & " c.WALK_IN, c.SICK_CALL, c.STREET_ADDRESS, c.BUILDING_NUMBER," _
        & " c.City, c.State, c.ZIP, c.DMIS_ID_CODE," _
        & " m.Description as MeprsDescription, c.Location_Type, c.BDM_DEF_PHARMACY_ID, c.DEFAULT_WORKLOAD_TYPE" _
        & " From Clinic c, User_Clinic uc, Meprs_Code m" _
        & " Where uc.User_NCID = " & vsUserNCID _
        & " and c.ncid = uc.Clinic_NCID and m.IEN = c.Meprs_IEN" _
        & " and c.facility_ncid = " & vsMTFFacilityNCID _
        & " and c.facility_ncid = m.facility_ncid" _
        & " and (c.inactive <> 1 or inactive is null) " _
        & " Order by c.Name"
        
   End If
   
'''<TEST
'Call GetClinicsWorkload(vsMTFFacilityNCID, colClinics)
'''>TEST
   '''SCR #30152;   Developer: Alexander Alberstein 01/10/2003 19:52 PM
   Set GetUserClinics = ClinicClient_colGetClinics(sStmt)
   
Exit Function
errHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME & " usererNcid=" & vsUserNCID, sPROC_NAME, vbExclamation)
   'Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, vsUserNCID)
   Set GetUserClinics = New Collection
   
End Function

Public Function GetAllClinicsForUser(ByVal vsUserNCID As String) As Collection
'SCR-49371 RBELL 2/10/2004 This function is used to gather all clinics in the users MTF it does it by subquerying the user_clinic table and joining to the clinic table by facility_ncid
'The requirement for all clinics in the users mtf was defined as part of the template enhancements

   Const sPROC_NAME  As String = "GetAllClinicsForUser"
   Dim sStmt        As String
   Dim sClinicName As String
   Dim rstClinics       As adodb.Recordset
   Dim objClinic As Clinic
   Dim cUserClinics As Collection
   
   On Error GoTo errHandler
   
    If gobjShared.IsAppMode(enumCSCSIIAppMode.modeTheater) Then
       sStmt = "Select c.Facility_NCID, c.NCID, c.IEN, c.Name, c.Meprs_IEN, m.Code as Meprs_Code, " & vbCrLf _
        & " c.WALK_IN, c.SICK_CALL, c.STREET_ADDRESS, c.BUILDING_NUMBER," & vbCrLf _
        & " c.City, c.State, c.ZIP, c.DMIS_ID_CODE," & vbCrLf _
        & " m.Description as MeprsDescription, c.Location_Type " & vbCrLf _
        & " From Clinic c, Meprs_Code m" & vbCrLf _
        & " Where m.IEN = c.Meprs_IEN" & vbCrLf _
        & " and m.facility_ncid = c.facility_ncid " & vbCrLf _
        & " and (c.inactive <> 1 or inactive is null) " & vbCrLf _
        & " and c.facility_ncid IN (SELECT Facility_ncid from User_Clinic Where User_NCID = " & vsUserNCID & ")" & vbCrLf _
        & " Order by c.Name"
       Set GetAllClinicsForUser = ClinicClient_colGetClinics(sStmt)
    Else
        Set cUserClinics = New Collection
        Do: 'Used to capture records if more thean 5k a rereturned from SQL statement
            sStmt = "Select c.Facility_NCID, c.NCID, c.IEN, c.Name, c.Meprs_IEN, m.Code as Meprs_Code, " & vbCrLf _
                 & " c.WALK_IN, c.SICK_CALL, c.STREET_ADDRESS, c.BUILDING_NUMBER," & vbCrLf _
                 & " c.City, c.State, c.ZIP, c.DMIS_ID_CODE," & vbCrLf _
                 & " m.Description as MeprsDescription, c.Location_Type, c.BDM_DEF_PHARMACY_ID, c.DEFAULT_WORKLOAD_TYPE" & vbCrLf _
                 & " From Clinic c, Meprs_Code m" & vbCrLf _
                 & " Where "
            If Len(sClinicName) > 0 Then sStmt = sStmt & "c.NAME > " & SQLQuote(sClinicName) & vbCrLf & "and " 'To handle 5K limit
            sStmt = sStmt & "m.IEN = c.Meprs_IEN" & vbCrLf _
                    & "and m.facility_ncid = c.facility_ncid" & vbCrLf _
                    & "and (c.inactive <> 1 or inactive is null)" & vbCrLf _
                    & "and c.facility_ncid IN (SELECT Facility_ncid from User_Clinic Where User_NCID = " & vsUserNCID & ")" & vbCrLf _
                    & "Order by c.Name"
            If ClinicClient_bOpenRecordSet(sStmt, rstClinics) Then
               With rstClinics
                    .MoveFirst
                    Do While Not .EOF
                        Set objClinic = New Clinic
                        Call objClinic.F_Deserialize(.Fields)
                        sClinicName = objClinic.Name
                        cUserClinics.Add objClinic, objClinic.NCID
                        Call .MoveNext
                    Loop
                    Call .Close
               End With
               Set rstClinics = Nothing
            Else
                Exit Do 'No records returned
            End If
        Loop
        Set GetAllClinicsForUser = cUserClinics
   End If
   
   GoTo CleanUpAndExit
    
errHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME & " usererNcid=" & vsUserNCID, sPROC_NAME, vbExclamation)
    Set GetAllClinicsForUser = New Collection
CleanUpAndExit:
   Set rstClinics = Nothing
   Set objClinic = Nothing
   Set cUserClinics = Nothing

End Function

Public Function GetProviderName _
      (ByVal vsMTFFacilityNCID As String, ByVal vsProviderNCID As String) As String
'<SCR #30152;   Developer: Alexander Alberstein 01/13/2003 10:00 AM
   Dim sStatement    As String
      
   '<bjr120301
   If mbUseClinicProviderTable Then
      sStatement = "Select provider_name from clinic_provider" _
         & " Where clinician_ncid = " & vsProviderNCID _
         & " and facility_ncid = " & vsMTFFacilityNCID
   Else
      sStatement = "Select name as provider_name from provider" _
         & " Where ncid = " & vsProviderNCID _
         & " and facility_ncid = " & vsMTFFacilityNCID
   End If
   
   GetProviderName = ClinicClient_sColumnValue(sStatement, "Provider_Name")
   
End Function

Public Function GetProviderClinics _
      (ByVal vsMTFFacilityNCID As String, ByVal vsClinicianNCID As String) As Collection
   Const sPROC_NAME  As String = "GetProviderClinics"
   
   Dim oCmd          As CHCSII_Command.CHCSII_Cmd
   Dim sStmt         As String
   Dim sProcName     As String
   Dim oRs           As adodb.Recordset
   Dim colResult        As Collection
   Dim oClinic       As Clinic
On Error GoTo errHandler
    If gobjShared.IsAppMode(enumCSCSIIAppMode.modeTheater) Then
        sStmt = "Select c.Facility_NCID,  "
        sStmt = sStmt & "  c.NCID, "
        sStmt = sStmt & "  c.IEN, "
        sStmt = sStmt & "  c.Name, "
        sStmt = sStmt & "  c.Meprs_IEN, "
        sStmt = sStmt & "  m.Code as Meprs_Code, "
        sStmt = sStmt & "  c.WALK_IN, "
        sStmt = sStmt & "  c.SICK_CALL, "
        sStmt = sStmt & "  c.STREET_ADDRESS, "
        sStmt = sStmt & "  c.BUILDING_NUMBER,"
        sStmt = sStmt & "  c.City, "
        sStmt = sStmt & "  c.State, "
        sStmt = sStmt & "  c.ZIP, "
        sStmt = sStmt & "  c.DMIS_ID_CODE,"
        sStmt = sStmt & "  m.Description as MeprsDescription, "
        sStmt = sStmt & "  c.Location_Type"
        sStmt = sStmt & "      From Clinic c, Clinic_Provider cp, Meprs_Code m"
        sStmt = sStmt & "      Where cp.Clinician_NCID = " & vsClinicianNCID
        sStmt = sStmt & "      and c.ncid = cp.Clinic_NCID and m.IEN = c.Meprs_IEN"
        sStmt = sStmt & "      and c.facility_ncid = " & vsMTFFacilityNCID
        sStmt = sStmt & "      and c.facility_ncid = m.facility_ncid"
        sStmt = sStmt & "      and (c.inactive <> 1 or inactive is null)"
        sStmt = sStmt & "      Order by c.Name"

    ElseIf gobjShared.IsAppMode(enumCSCSIIAppMode.modeCDR) And _
        Not gobjShared.IsAppMode(enumCSCSIIAppMode.modeITT) Then
        
        ' SCR #35883;   Developer: MAG 03/03/2004
        'Stored Procedure #1
        Set colResult = New Collection

        If mbUseClinicProviderTable Then
          sProcName = "CLINIC_CLIENT_PERFORMANCE_PKG.Clinics_Via_ClinicProviderTab"
        Else
          sProcName = "CLINIC_CLIENT_PERFORMANCE_PKG.Clinics_Via_Provider_Table"
        End If
        
        Set oCmd = New CHCSII_Command.CHCSII_Cmd
        
        With oCmd
            .SPName = sProcName
            .AddParam False, False, "textString", vsClinicianNCID
            .AddParam False, False, "textString", vsMTFFacilityNCID
            .AddParam True, True, "", ""   '--- for Output Parameter
        End With
        
        Set oRs = ClinicClient_objDAS.OpenRecordsetSP(oCmd)

        If Not oRs Is Nothing Then
          If oRs.RecordCount > 0 Then
            With oRs
              .MoveFirst
              Do While Not .EOF
                Set oClinic = New Clinic
                Call oClinic.F_Deserialize(.Fields)
                On Error Resume Next
                colResult.Add oClinic, oClinic.NCID
                On Error GoTo errHandler
                Call .MoveNext
             Loop
             Call .Close
            End With
          End If
          Set oRs = Nothing
        End If
        
        Set GetProviderClinics = colResult
      
        Exit Function
    Else
        '''SCR #31046;   Developer: Alexander Alberstein 12/26/2002 03:32 PM
        If mbUseClinicProviderTable Then
          sStmt = "Select c.Facility_NCID, c.NCID, c.IEN, c.Name, c.Meprs_IEN, m.Code as Meprs_Code, " _
              & " c.WALK_IN, c.SICK_CALL, c.STREET_ADDRESS, c.BUILDING_NUMBER," _
              & " c.City, c.State, c.ZIP, c.DMIS_ID_CODE," _
              & " m.Description as MeprsDescription, c.Location_Type, c.BDM_DEF_PHARMACY_ID, c.DEFAULT_WORKLOAD_TYPE" _
              & " From Clinic c, Clinic_Provider cp, Meprs_Code m" _
              & " Where cp.Clinician_NCID = " & vsClinicianNCID _
              & " and c.ncid = cp.Clinic_NCID and m.IEN = c.Meprs_IEN" _
              & " and c.facility_ncid = " & vsMTFFacilityNCID _
              & " and c.facility_ncid = m.facility_ncid" _
              & " and (c.inactive <> 1 or inactive is null)" _
              & " Order by c.Name"
        Else
          sStmt = "Select c.Facility_NCID, c.NCID, c.IEN, c.Name, c.Meprs_IEN, m.Code as Meprs_Code, " _
              & " c.WALK_IN, c.SICK_CALL, c.STREET_ADDRESS, c.BUILDING_NUMBER," _
              & " c.City, c.State, c.ZIP, c.DMIS_ID_CODE," _
              & " m.Description as MeprsDescription, c.Location_Type, c.BDM_DEF_PHARMACY_ID, c.DEFAULT_WORKLOAD_TYPE" _
              & " From Clinic c, Provider p, Meprs_Code m" _
              & " Where p.NCID = " & vsClinicianNCID _
              & " and c.ncid = p.Clinic_NCID and m.IEN = c.Meprs_IEN" _
              & " and c.facility_ncid = " & vsMTFFacilityNCID _
              & " and c.facility_ncid = m.facility_ncid" _
              & " and (c.inactive <> 1 or inactive is null)" _
              & " Order by c.Name"
        End If
    End If
    '''<SCR #30152;   Developer: Alexander Alberstein 01/10/2003 9:00 PM
    Set GetProviderClinics = ClinicClient_colGetClinics(sStmt)
    
Exit Function
errHandler:
   Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME & " clinicianNCID=" & vsClinicianNCID, sPROC_NAME, vbExclamation)
   'Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, vsClinicianNCID)
   Set GetProviderClinics = New Collection
   
End Function
Public Function GetClinicProviders _
      (ByVal vsClinicNCID As String, Optional ByVal bUsersOnly As Boolean = False, _
       Optional ByVal bIncludeRoles As Boolean _
      ) As Collection
   Const sPROC_NAME  As String = "GetClinicProviders"
   
   Dim sStmt           As String
    
On Error GoTo errHandler

   If mbUseClinicProviderTable Then
      Set GetClinicProviders = ClinicClient_colGetClinicProviders(vsClinicNCID, bUsersOnly, , bIncludeRoles)
   Else
      '''<bjr110501 SCR4004
      sStmt = "Select NCID as Clinician_NCID, IEN as Provider_IEN, Name as Provider_Name, User_IEN" _
          & " From Provider" _
          & " Where Clinic_NCID = " & vsClinicNCID
      If bUsersOnly Then
         '''<bjr111901 scr18760
          sStmt = sStmt & " and exists (select 'x' from hems_users where user_ncid = ncid)"
      End If
      sStmt = sStmt & " Order by Provider_Name"
      '''<SCR #30152;   Developer: Alexander Alberstein 01/13/2003 2:00 PM
      Set GetClinicProviders = ClinicClient_colGetProviders(sStmt)
   End If
   
Exit Function
errHandler:
    Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, vbExclamation)
   'Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")

End Function
Public Function GetClinicProvidersLikeLastName _
      (ByVal vsMTFFacilityNCID, _
      ByVal vsLikeLastName, _
      ByVal vsClinicNCID As String, _
      Optional ByVal bUsersOnly As Boolean = False) As Collection
''''''''''''''''''''''''''''''''''''
'''Desc:   Returns collection aka GetClinicProviders with the addition of the
'''           LIKE provider_name constraint
'''Note:   <bjr102901 SCR4004
''''''''''''''''''''''''''''''''''
   Const sPROC_NAME  As String = "GetClinicProvidersLikeLastName"
   
   Dim sStmt           As String
   Dim colProviders    As Collection
   Dim objProvider     As Provider
    
On Error GoTo errHandler
   
   If mbUseClinicProviderTable Then
      '''<BUG ClinicClient_colGetClinicProviders currently does not support like operator
      Set colProviders = ClinicClient_colGetClinicProviders _
         (vsClinicNCID, False, "", False)
   Else
      ' Begin:SCR #44789;   Developer:MAG  11/05/2003
      sStmt = "Select NCID AS Clinician_NCID, IEN as Provider_IEN, Name as Provider_Name, User_IEN" _
          & " From Provider" _
          & " Where Clinic_NCID = " & vsClinicNCID _
          & " and UPPER(Name) LIKE " & SQLQuote(UCase$(vsLikeLastName))
      ' End: SCR #44789;
      If bUsersOnly Then
          sStmt = sStmt & " and exists (select 'x' from hems_users where user_ncid = ncid)"
      End If
      sStmt = sStmt & " Order by Name"
      
      '''<SCR #30152;   Developer: Alexander Alberstein 01/13/2003 2:00 PM
      Set colProviders = ClinicClient_colGetProviders(sStmt)
   End If
   
   Set GetClinicProvidersLikeLastName = colProviders
   
Exit Function
errHandler:
   Set GetClinicProvidersLikeLastName = Nothing
   Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, vbExclamation)
   'Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")
   Set colProviders = Nothing
   
End Function

Public Function GetProviders _
   (ByVal sMTFFacilityNCID As String, _
   ByVal sLikeLastName As String, _
   ByVal sClinicNCID As String, _
   Optional ByVal bUsersOnly As Boolean = False, _
   Optional ByVal bSignable As Boolean = False) As Collection
''''''''''''''''''''''''''''''''''''
'''Desc:   Selection builder based on passed criteria
'''        Could replace
'''           Public Function GetClinicProviders _
'''           Public Function GetClinicProvidersLikeLastName _
'''           Public Function GetAllProviders _
'''           Public Function GetAllProvidersLikeLastName _
'''           Public Function GetSignableClinicProviders _
'''Note:   <bjr120301 SCR18815
''''''''''''''''''''''''''''''''''
   Const sPROC_NAME  As String = "GetProviders"
   
   Dim SQLt          As SQLText
   Dim sAlias        As String
   
On Error GoTo errHandler

   Set SQLt = New SQLText
   
   If mbUseClinicProviderTable Then
      Call SQLt.AddSQLSelect("cp.Clinician_NCID")
      Call SQLt.AddSQLSelect("cp.Provider_IEN")
      Call SQLt.AddSQLSelect("cp.Provider_Name")
      Call SQLt.AddSQLSelect("cp.User_IEN")
      Call SQLt.AddSQLSelect("cp.Signature_Class")
      Call SQLt.AddSQLFrom("Clinic_Provider cp")
      sAlias = "cp."
      If sLikeLastName <> "" Then
         ' Begin:SCR #44789;   Developer:  11/05/2003
         Call SQLt.AddSQLWhere("UPPER(Provider_Name) LIKE " & SQLQuote(UCase$(sLikeLastName) & "%"))
        ' End: SCR #44789;
      End If
      If bUsersOnly Then
         Call SQLt.AddSQLWhere("exists (select 'x' from hems_users where user_ncid = cp.clinician_ncid)")
      End If
   Else
      Call SQLt.AddSQLSelect("p.NCID AS Clinician_NCID")
      Call SQLt.AddSQLSelect("p.IEN as Provider_IEN")
      Call SQLt.AddSQLSelect("p.Name as Provider_Name")
      Call SQLt.AddSQLSelect("p.User_IEN")
      Call SQLt.AddSQLSelect("p.Signature_Class")
      Call SQLt.AddSQLFrom("Provider p")
      sAlias = "p."
      If sLikeLastName <> "" Then
         ' Begin:SCR #44789;   Developer:  11/05/2003
         Call SQLt.AddSQLWhere("UPPER(Name) LIKE " & SQLQuote(UCase$(sLikeLastName) & "%"))
         ' End: SCR #44789;
      End If
      If bUsersOnly Then
         Call SQLt.AddSQLWhere("exists (select 'x' from hems_users where user_ncid = p.ncid)")
      End If
   End If
   
   If sMTFFacilityNCID <> "" Then
      Call SQLt.AddSQLWhere(sAlias & "facility_ncid = " & sMTFFacilityNCID)
   End If
   
   If sClinicNCID <> "" And sClinicNCID <> "0" Then
      If mbUseClinicProviderTable Then
         Call SQLt.AddSQLWhere("cp.Clinic_NCID = " & sClinicNCID)
      Else
        '''<< SCR 20223 cosign enc.  - inlcude all clinics (clinic-provider table, too)
        SQLt.AddSQLFrom ("Clinic_Provider cp")
        SQLt.AddSQLWhere ("cp.clinician_ncid = p.ncid")
        SQLt.AddSQLWhere ("(cp.clinic_ncid = " & sClinicNCID & " Or p.clinic_ncid = " & sClinicNCID & ")")
        SQLt.AddSQLWhere (" p.ncid = cp.clinician_ncid")
      End If
   End If
   
   If bSignable Then
      Call SQLt.AddSQLWhere(sAlias & "SIGNATURE_CLASS IN(2,3,4)")
   End If
   
   Call SQLt.AddSQLOrder("Provider_name")
   
   '''<SCR #30152;   Developer: Alexander Alberstein 01/13/2003 2:00 PM
   Set GetProviders = ClinicClient_colGetProviders(SQLt.GetSQLStatement)
   
Exit Function
errHandler:
   Set GetProviders = Nothing
   Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, vbExclamation)
   'Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")
   
End Function
Public Function GetClinicIEN(ByVal vsClinicNCID As String) As String
'<SCR #30152;   Developer: Alexander Alberstein 01/13/2003 10:00 AM
   GetClinicIEN = ClinicClient_sColumnValue _
                 ("Select IEN from clinic where NCID = " & vsClinicNCID, "IEN")
End Function

Public Function GetClinicNCID(ByVal vsMTFFacilityNCID As String, vsClinicIEN As String) As String
'<SCR #30152;   Developer: Alexander Alberstein 01/13/2003 10:00 AM
   GetClinicNCID = ClinicClient_sColumnValue _
   ("Select NCID from clinic where IEN = " & vsClinicIEN & " and facility_ncid = " & vsMTFFacilityNCID, "NCID")
    
End Function

Public Function GetAllProviders _
   (ByVal vsMTFFacilityNCID, _
    Optional ByVal bUsersOnly As Boolean = False, Optional ByVal bSignable As Boolean = False, _
    Optional ByVal sSingleProviderNCID As String = vbNullString, _
    Optional ByVal bSignCosign As Boolean = False) As Collection ''' of Provider Objects
    '104848
    'sSingleProviderNCID added to retrieve only one specific provider
    '12/05/03 S.McAvoy Called by Appointment module to meet IBWA-RNDS requirement.
    
    Const sPROC_NAME     As String = "GetAllProviders"

    Dim sStmt            As String
    
On Error GoTo errHandler
    ''''''''''''''''''''''''
    '<bjr102701 SCR4004
    ' Query from table as directed by flag
    ''''''''''''''''''''''''
    If mbUseClinicProviderTable Then
        sStmt = "Select distinct Clinician_NCID, Provider_IEN, Provider_Name, User_IEN, SIGNATURE_CLASS" _
            & " From Clinic_Provider" _
            & " Where Facility_NCID = " & vsMTFFacilityNCID
        If bUsersOnly Then   '<bjr111901 scr18760 - put inside conditional (from statement affects field naming)
            sStmt = sStmt & " and exists (select 'x' from hems_users where user_ncid = clinician_ncid)"
        End If
    Else 'Use Provider table
        sStmt = "Select distinct NCID as Clinician_NCID, IEN as Provider_IEN, Name as Provider_Name, User_IEN, SIGNATURE_CLASS" _
            & " From Provider" _
            & " Where Facility_NCID = " & vsMTFFacilityNCID
        If bUsersOnly Then   '''<bjr111901 scr18760 - put inside conditional (from statement affects field naming)
            sStmt = sStmt & " and exists (select 'x' from hems_users where user_ncid = ncid)"
        End If
    End If
    If bSignable Then
        sStmt = sStmt & " AND SIGNATURE_CLASS IN(2,3,4)"
    End If
    
    If bSignCosign = True Then
        '104848 Only get providers that have sign and cosign privs.
        sStmt = sStmt & " and Clinician_NCID in (SELECT user_roles.user_ncid" _
                        & " From scadb.user_roles, scadb.role_priv, scadb.available_priv" _
                        & " WHERE ((user_roles.role_id = role_priv.role_id)" _
                        & " AND (role_priv.priv_id = available_priv.priv_id))" _
                        & " and (lower(available_priv.priv_name) = 'current_encounter')" _
                        & " and (role_priv.priv_string like '%S%' or role_priv.priv_string like '%C%'))"

    End If
   
    If sSingleProviderNCID <> vbNullString Then
        If mbUseClinicProviderTable Then
            sStmt = sStmt & " AND Clinician_NCID = " & sSingleProviderNCID
        Else
            sStmt = sStmt & " AND NCID = " & sSingleProviderNCID
        End If
    Else
        sStmt = sStmt & " Order by Provider_Name"
    End If
   
   
   '''<SCR #30152;   Developer: Alexander Alberstein 01/13/2003 2:00 PM
   Set GetAllProviders = ClinicClient_colGetProviders(sStmt)
   
Exit Function
errHandler:
   Set GetAllProviders = Nothing
   
   Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, vbExclamation)
   'Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")
    
End Function
Public Function GetAllProvidersLikeLastName _
   (ByVal vsMTFFacilityNCID, ByVal vsLikeLastName, _
    Optional ByVal bUsersOnly As Boolean = False, Optional ByVal bSignable As Boolean = False _
    ) As Collection     ' of Provider Objects
    
   Const sPROC_NAME     As String = "GetAllProvidersLikeLastName"
   
   Dim sStmt            As String
    
On Error GoTo errHandler

   ''''''''''''''''''''''''
   '<bjr102701 SCR4004
   ' Query from table as directed by flag
   ''''''''''''''''''''''''
   If mbUseClinicProviderTable Then
      ' Begin:SCR #44789;   Developer:  11/05/2003
      sStmt = "Select distinct Clinician_NCID, Provider_IEN, Provider_Name, User_IEN, SIGNATURE_CLASS" _
          & " From Clinic_Provider" _
          & " Where Facility_NCID = " & vsMTFFacilityNCID _
          & " And UPPER(Provider_Name) LIKE " & SQLQuote(UCase$(vsLikeLastName))
      ' End: SCR #44789;
      If bUsersOnly Then   '''<bjr111901 scr18760 - put inside conditional (from statement affects field naming)
          sStmt = sStmt & " and exists (select 'x' from hems_users where user_ncid = clinician_ncid)"
      End If
   Else 'Use Provider table
      ' Begin:SCR #44789;   Developer:  11/05/2003
      sStmt = "Select distinct NCID as Clinician_NCID, IEN as Provider_IEN, Name as Provider_Name, User_IEN, SIGNATURE_CLASS" _
          & " From Provider" _
          & " Where Facility_NCID = " & vsMTFFacilityNCID _
          & " And UPPER(Name) LIKE " & SQLQuote(UCase$(vsLikeLastName))
      ' End: SCR #44789;
      If bUsersOnly Then   '''<bjr111901 scr18760 - put inside conditional (from statement affects field naming)
          sStmt = sStmt & " and exists (select 'x' from hems_users where user_ncid = ncid)"
      End If
   End If
   If bSignable Then
     sStmt = sStmt & " AND SIGNATURE_CLASS IN(2,3,4)"
   End If
   sStmt = sStmt & " Order by Provider_Name"

   '''<SCR #30152;   Developer: Alexander Alberstein 01/13/2003 2:00 PM
   Set GetAllProvidersLikeLastName = ClinicClient_colGetProviders(sStmt)
   
Exit Function
errHandler:
   Set GetAllProvidersLikeLastName = Nothing
   Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, vbExclamation)
   'Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")
    
End Function
Public Function GetProviderIEN(ByVal vsMTFFacilityNCID As String, ByVal vsProviderNCID As String) As String
'<SCR #30152;   Developer: Alexander Alberstein 01/13/2003 10:00 AM
   Dim sStatement    As String
    
   If mbUseClinicProviderTable Then
      sStatement = "Select provider_ien from clinic_provider" _
          & " Where clinician_ncid = " & vsProviderNCID _
          & " and facility_ncid = " & vsMTFFacilityNCID
   Else
      sStatement = "Select ien as provider_ien from provider" _
          & " Where ncid = " & vsProviderNCID _
          & " and facility_ncid = " & vsMTFFacilityNCID
   End If
   GetProviderIEN = ClinicClient_sColumnValue(sStatement, "Provider_IEN")
    
End Function
Public Function GetProviderNCID(ByVal vsMTFFacilityNCID As String, ByVal vsProviderIEN As String) As String
'<SCR #30152;   Developer: Alexander Alberstein 01/13/2003 10:00 AM
   Dim sStatement    As String
    
   If mbUseClinicProviderTable Then
      sStatement = "Select clinician_ncid from clinic_provider" _
          & " Where provider_ien = " & vsProviderIEN _
          & " and facility_ncid = " & vsMTFFacilityNCID
   Else
      sStatement = "Select ncid as clinician_ncid from provider" _
         & " Where ien = " & vsProviderIEN _
         & " and facility_ncid = " & vsMTFFacilityNCID
   End If
   
   GetProviderNCID = ClinicClient_sColumnValue(sStatement, "Clinician_NCID")
   
End Function
Private Function M_bFindRecord(ByVal sColumnName As String, ByVal sColumnValue As String) As Boolean
'''************************************************************************
'''Called By:        M_bPopulateWorkLoad
'''
'''Arguments:
'''   sColumnName
'''      Nof the column being evalauted
'''   sColumnValue
'''      Value being sought
'''
'''Description:      Attempt to find first occurance of the specified value
'''                  for specified column within current mobjSQL recordset
'''
'''
'''Note:             if match is found mobjSQL cursor will be positioned at the matching record
'''                  else it will be at EOF
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If match found
'''                  true
'''               Else
'''                  false
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME        As String = "M_bFindRecord"

   Dim bResult             As Boolean
   
On Error GoTo ErrHnd


100   With mrst
         '''If any records
200      If Not (.BOF And .EOF) Then
220         Call .MoveFirst
300         Do While (Not .EOF) '''Will exit loop on first match
400            If sColumnValue = .Fields(sColumnName).Value Then
                  bResult = True
420               Exit Do
               Else
500               Call .MoveNext
               End If
            Loop
         End If
      End With
      
999   M_bFindRecord = bResult
   
Exit Function
ErrHnd:
   M_bFindRecord = False
   Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, vbExclamation)
   'Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl)

End Function

Private Function M_bGetProviderObj _
        (col As Collection, ByVal sKey As String, ByRef objProvider As Provider) As Boolean
        
On Error GoTo ErrHnd

   Set objProvider = col(sKey)
   M_bGetProviderObj = True
   
Exit Function
ErrHnd:
   M_bGetProviderObj = False
   Set objProvider = Nothing
   
End Function

Private Function M_bPopulateProvidersWorkload _
        (ByVal sSelectStatement As String, ByRef colProviders As Collection) As Boolean
'''************************************************************************
'''Called By:        GetProvidersWorkload
'''
'''Arguments:
''' colProviders
'''   Collection of the provider objects requested to be populated with workload information
'''
'''Description:      Populate workloads properties of all specified provider objects
'''
'''
'''Note:             'SCR #31046;   Developer: Alexander Alberstein 12/28/2002 16:32 PM
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:       If successful
'''                  true
'''               Else
'''                  false
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "M_bPopulateProvidersWorkload"

   Dim objProvider      As Provider
   Dim objWorkload      As Workload
   Dim sBreak           As String
   Dim sControl      As String
   Dim bPopulate        As Boolean
   
On Error GoTo ErrHnd

      '''<SCR #30152;   Developer: Alexander Alberstein 01/13/2003 2:00 PM
      If ClinicClient_bOpenRecordSet(sSelectStatement, mrst) Then
100      With mrst
300         Do '''Populate workload collection
320            sControl = .Fields("PROVIDER_NCID").Value & ""
               '''If next subset
400            If Not sBreak = sControl Then
420               sBreak = sControl
                  '''If was able to find corresponding provider object
500               If M_bGetProviderObj(colProviders, sControl, objProvider) Then
510                  Set objProvider.Workloads = New Collection
520                  bPopulate = True
                  Else
                     bPopulate = False
                  End If
               End If
               '''If subset has corresponding provider object
600            If bPopulate Then
                  '''Populate next provider workload object
620               Set objWorkload = New Workload
640               Call objWorkload.F_Deserialize(.Fields)
                  '''Store workload object
700               Call objProvider.Workloads.Add(objWorkload)
               End If
800            Call .MoveNext
900         Loop Until .EOF
         End With
940      Call mrst.Close
      Else
         '''Always return valid collection of workloads
990      For Each objProvider In colProviders
992         Set objProvider.Workloads = New Collection
         Next objProvider
      End If
   
999   M_bPopulateProvidersWorkload = True
   
Exit Function
ErrHnd:
   M_bPopulateProvidersWorkload = False
   Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, vbExclamation)
   'Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl)

End Function

Private Function M_bPopulateWorkload _
   (ByVal sSubsetColumn As String, ByVal sSubsetColumnValue As String, ByRef colWorkload As Collection) _
    As Boolean
'''************************************************************************
'''Called By:     GetClinicsWorkload
'''
'''Arguments:
''' sSubsetColumn
'''   Name of the column containing sSubsetColumnValue
''' sSubsetColumnValue
'''   Control break value of the subset
''' colWorkload
'''   Collection requested to be populated with workload objects
'''
'''Description:      Attempt to populate requested workloads collection
'''
'''
'''Note:             'SCR #31046;   Developer: Alexander Alberstein 12/27/2002 16:32 PM
'''
'''Warning:          mobjSQL is assumed to contain recordset of workload records associated with
'''                  entity identified by sSubsetColumnValue
'''                  colWorkload is asumed to be Instantiated
'''Updates Globals:
'''
'''Returns:       If successful
'''                  true
'''               Else
'''                  false
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME        As String = "M_bPopulateWorkLoad"

   Dim objWorkload         As Workload
   Dim bEOSubset           As Boolean
   
On Error GoTo ErrHnd

      '''if was able to find subset associated with specified clinic
100   If M_bFindRecord(sSubsetColumn, sSubsetColumnValue) Then
200      With mrst
            Do '''Populate workload collection
500            Set objWorkload = New Workload
520            Call objWorkload.F_Deserialize(.Fields)
700            Call colWorkload.Add(objWorkload)
780            Call .MoveNext
800            If Not .EOF Then
810               bEOSubset = Not sSubsetColumnValue = .Fields(sSubsetColumn).Value
               Else
820               bEOSubset = True
               End If
            '''Until end of the subset
900         Loop Until bEOSubset
         End With
      End If

999   M_bPopulateWorkload = True
   
Exit Function
ErrHnd:
   M_bPopulateWorkload = False
   Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, vbExclamation)
   'Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, gsLINE_ & Erl)

End Function

Private Sub Class_Initialize()
'''************************************************************************
'''Called By:     None
'''
'''Arguments:
'''Description:   Initialize module
'''
'''Note:
'''
'''Warning:
'''
'''Updates Globals:
'''
'''Returns:
'''
'''Programmer:    Alexander Alberstein
'''************************************************************************
   Const sPROC_NAME     As String = "Class_Initialize"

   Call ClinicClient_bLoad
   mbUseClinicProviderTable = True  '''<bjr102701 SCR4004

End Sub

Private Sub Class_Terminate()

On Error Resume Next

   Set mrst = Nothing
   Call ClinicClient_Unload

End Sub

Public Function GetSignableClinicProviders _
   (ByVal vsClinicNCID As String, Optional ByVal bUsersOnly As Boolean = False) As Collection
   
   Const sPROC_NAME  As String = "GetSignableClinicProviders"
   
   Dim sStmt           As String
   
On Error GoTo errHandler

   ''''''''''''''''''''''''
   '<bjr102701 SCR4004
   ' Query from table as directed by flag
   ''''''''''''''''''''''''
   If mbUseClinicProviderTable Then
      sStmt = "Select Clinician_NCID, Provider_IEN, Provider_Name, User_IEN, SIGNATURE_CLASS" _
          & " From Clinic_Provider" _
          & " Where Clinic_NCID = " & vsClinicNCID _
      & " AND SIGNATURE_CLASS IN(2,3,4)"
      If bUsersOnly Then   '''<bjr111901 scr111901 - moved above
          sStmt = sStmt & " and exists (select 'x' from hems_users where user_ncid = clinician_ncid)"
      End If
   Else 'Use Provider table
      sStmt = "Select NCID as Clinician_NCID, IEN as Provider_IEN, Name as Provider_Name, User_IEN, SIGNATURE_CLASS" _
          & " From Provider" _
          & " Where Clinic_NCID = " & vsClinicNCID _
      & " AND SIGNATURE_CLASS IN(2,3,4)"
      If bUsersOnly Then   '''<bjr111901 scr111901 - moved above
          sStmt = sStmt & " and exists (select 'x' from hems_users where user_ncid = ncid)"
      End If
   End If
   
   sStmt = sStmt & " Order by Provider_Name"
   '''<SCR #30152;   Developer: Alexander Alberstein 01/13/2003 2:00 PM
   Set GetSignableClinicProviders = ClinicClient_colGetProviders(sStmt)
   
Exit Function
errHandler:
   Set GetSignableClinicProviders = Nothing
   Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, vbExclamation)
   'Call LibApp_Error(Err.Number, Err.Description, msMODULE_NAME, sPROC_NAME, "")
   
End Function

Public Function GetClinicDivStateIEN(ByVal sMTFFacilityNCID As String, sClinicNCID As String) As String

On Error GoTo errHandler

Dim sSQL As String

   If gobjShared Is Nothing Then Set gobjShared = New CWShared
   
   sSQL = "SELECT l.Location_IEN " & _
          " FROM clinic c, geographic_location l " & _
          " WHERE c.NCID = " & sClinicNCID & _
            " AND c.facility_ncid = " & sMTFFacilityNCID & _
            " AND c.division_state = substr(l.location_abbreviation,1,2) "
   
   If gobjShared.UseSQLServer Then
      'Replace the substr function with SQL Servers substring function
      sSQL = Replace(sSQL, "substr", "substring")
   End If
   
   GetClinicDivStateIEN = ClinicClient_sColumnValue(sSQL, "Location_IEN")
   
   Exit Function
   
errHandler:
   Call gobjShared.ShowVBError(Err.Number, Err.Description, msMODULE_NAME, "GetClinicDivStateIEN", vbExclamation)
   
End Function


