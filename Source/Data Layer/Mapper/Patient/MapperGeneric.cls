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
Attribute VB_Name = "MapperGeneric"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements DL_Interfaces.iMapperPatient

Private mConn As CHCSII_CONN.Conn
Private mCache As CacheMapper
Private mobjSQLFix As iSQLFix
Private mobjPicklist As iPickList
Private mDB As CHCSII_CONN.eBackend
Private mISQLOps As ICHCSII_SQL.ISqlOps
Private mobjDAS As ICHCSII_DAS



Friend Property Set SQLFix(ByVal NewValue As iSQLFix)
  Set mobjSQLFix = NewValue
End Property

Friend Property Set PickList(ByVal NewValue As iPickList)
  Set mobjPicklist = NewValue
End Property

Friend Property Get BackEndDB() As CHCSII_CONN.eBackend
    If LenB(mDB) > 0 Then
        BackEndDB = mDB
    Else
        BackEndDB = Auto
    End If
End Property
Friend Property Let BackEndDB(ByVal NewValue As CHCSII_CONN.eBackend)
    mDB = NewValue
End Property


Private Sub Class_Terminate()
    Set mobjSQLFix = Nothing
    Set mISQLOps = Nothing
End Sub

Private Function iMapperPatient_GetRankNCID(ByVal sRankCode As String, ByVal sBranch As String) As String
    Err.Raise Err.Number, "MapperGeneric.iMapperPatient_GetRankNCID", "GetRankNCID is not allowed in MapperGeneric.", Err.HelpFile, Err.HelpContext
End Function

Private Sub iMapperPatient_Init(Support As DL_Support.iParameters, Cache As DL_Support.CacheMapper)
  Set mConn = Support.Item("DBConn").Value
  Set mCache = Cache
  Set mobjDAS = mConn.CHCSII_DAS(BackEndDB)
End Sub


Private Function iMapperPatient_Retrieve(Params As DL_Support.iParameters) As DTOs.Patients
Dim objPatientParams As PatientSearchParam
    On Error GoTo ErrHandler
    
    Set objPatientParams = Params
    
    If objPatientParams.IncludeBasicOnly = True Then
        Set iMapperPatient_Retrieve = GetPatientsBySearchParams(objPatientParams)
    Else
        If LenB(objPatientParams.UnitNumber) = 0 Then
            Err.Raise ERR_NUMBER, "MapperGeneric.iMapperPatient_Retrieve", "Patient Unit_Number is Missing."
        ElseIf objPatientParams.UnitNumber = "0" Then
            Set iMapperPatient_Retrieve = GetNewPatients(objPatientParams)
        Else
            Set iMapperPatient_Retrieve = GetPatientsByID(objPatientParams.UnitNumber, objPatientParams.FacilityNCID, objPatientParams.IncludeSponsor)
        End If
    End If
    
    GoTo CleanUpAndExit
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.iMapperPatient_Retrieve", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
    Set objPatientParams = Nothing
    Exit Function
    Resume
End Function

Private Function iMapperPatient_RetrieveSelectStrings(ByRef Params As DL_Support.iParameters) As DTOs.PatientSelectStrings

    Dim objSelectStrParams As SelectStringSearchParam
    
    On Error GoTo ErrHandler
    
    Set objSelectStrParams = Params
    
    Select Case objSelectStrParams.SelectStringCategory
        Case osUnknownCategory
            Set iMapperPatient_RetrieveSelectStrings = Nothing
        Case osPatientCategory
            Set iMapperPatient_RetrieveSelectStrings = mobjPicklist.PATCATSelectStrings(NCID_PATIENT_CATEGORY)
        Case osFMP
             Set iMapperPatient_RetrieveSelectStrings = mobjPicklist.FMPSelectStrings(NCID_FMP_DOMAIN_GEMS, NCID_FMP_TEXT_DOMAIN_GEMS)
        Case osMaritalStatus
             Set iMapperPatient_RetrieveSelectStrings = mobjPicklist.ComboxSelectStrings(NCID_MARTIAL_STATUS_DOMAIN)
        Case osRace
             Set iMapperPatient_RetrieveSelectStrings = mobjPicklist.ComboxSelectStrings(NCID_RACE_DOMAIN)
        Case osSpecialtyCode
             Set iMapperPatient_RetrieveSelectStrings = mobjPicklist.ComboxSelectStrings(NCID_RACE_DOMAIN)
        Case osReligion
             Set iMapperPatient_RetrieveSelectStrings = mobjPicklist.ComboxSelectStrings(NCID_RELIGION_DOMAIN)
        Case osEnrollmentFacility
             Set iMapperPatient_RetrieveSelectStrings = mobjPicklist.ComboxSelectStrings(NCID_ENROLLMENT_FACILITY_DOMAIN)
        Case osPhoneType
             Set iMapperPatient_RetrieveSelectStrings = mobjPicklist.ComboxSelectStrings(NCID_PHONE_TYPE_DOMAIN)
        Case osAddressType
             Set iMapperPatient_RetrieveSelectStrings = mobjPicklist.ComboxSelectStrings(NCID_ADDRESS_TYPE_DOMAIN)
        Case osGender
             Set iMapperPatient_RetrieveSelectStrings = mobjPicklist.ComboxSelectStrings(NCID_GENDER_DOMAIN)
        Case osRank
             Set iMapperPatient_RetrieveSelectStrings = mobjPicklist.ComboxSelectStrings(NCID_RANK)
        Case osRecordsMaintainedAt
             Set iMapperPatient_RetrieveSelectStrings = mobjPicklist.ComboxSelectStrings(NCID_RECORDS_MAINTAINED_AT)
        Case osACVEnrollmentStatus
             Set iMapperPatient_RetrieveSelectStrings = mobjPicklist.ComboxSelectStrings(NCID_ACV_ENROLLMENT_STATUS)
    End Select
    
    GoTo CleanUpAndExit
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.iMapperPatient_RetrieveSelectStrings", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
    Set objSelectStrParams = Nothing
End Function

Private Function iMapperPatient_RetrieveRepresentation(ByRef Params As DL_Support.iParameters) As String
Dim oParams As SelectStringSearchParam
Dim sNCID As String

    On Error GoTo ErrHandler
    
    Set oParams = Params
    
    sNCID = oParams.NCID
    
    Select Case oParams.SelectStringCategory
        Case osUnknownCategory
            iMapperPatient_RetrieveRepresentation = vbNullString
        Case osPatientCategory
            iMapperPatient_RetrieveRepresentation = mobjPicklist.ReturnRepFromNCID(NCID_PATIENT_CATEGORY, sNCID)
        Case osMaritalStatus
            iMapperPatient_RetrieveRepresentation = mobjPicklist.ReturnRepFromNCID(NCID_MARTIAL_STATUS_DOMAIN, sNCID)
        Case osRace
            iMapperPatient_RetrieveRepresentation = mobjPicklist.ReturnRepFromNCID(NCID_RACE_DOMAIN, sNCID)
        Case osSpecialtyCode
            iMapperPatient_RetrieveRepresentation = mobjPicklist.ReturnRepFromNCID(NCID_RACE_DOMAIN, sNCID)
        Case osReligion
            iMapperPatient_RetrieveRepresentation = mobjPicklist.ReturnRepFromNCID(NCID_RELIGION_DOMAIN, sNCID)
        Case osEnrollmentFacility
            iMapperPatient_RetrieveRepresentation = mobjPicklist.ReturnRepFromNCID(NCID_ENROLLMENT_FACILITY_DOMAIN, sNCID)
        Case osPhoneType
            iMapperPatient_RetrieveRepresentation = mobjPicklist.ReturnRepFromNCID(NCID_PHONE_TYPE_DOMAIN, sNCID)
        Case osAddressType
            iMapperPatient_RetrieveRepresentation = mobjPicklist.ReturnRepFromNCID(NCID_ADDRESS_TYPE_DOMAIN, sNCID)
        Case osGender
            iMapperPatient_RetrieveRepresentation = mobjPicklist.ReturnRepFromNCID(NCID_GENDER_DOMAIN, sNCID)
        Case osRank
            iMapperPatient_RetrieveRepresentation = mobjPicklist.ReturnRepFromNCID(NCID_RANK, sNCID)
    End Select
    
    GoTo CleanUpAndExit
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.iMapperPatient_ReturnRepFromNCID", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
    Set oParams = Nothing
End Function

Private Function iMapperPatient_Save(ByRef oPatient As DTOs.Patient) As Boolean
    
    Dim oState As iState
    
    On Error GoTo ErrHandler
    
    Set oState = oPatient
    
    If oState.State = osDelete Then
        iMapperPatient_Save = DeletePatient(oPatient.UnitNumber)
    ElseIf oState.State = osNew Then
        iMapperPatient_Save = SavePatient(oPatient, "N")
    Else
        iMapperPatient_Save = SavePatient(oPatient, "U")
    End If
    
    UpdatePatientObject oPatient

    GoTo CleanUpAndExit
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.iMapperPatient_Save", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
    Set oState = Nothing
    Exit Function
    Resume
End Function

Private Function iMapperPatient_TransferCHCSPatient(ByRef Params As DL_Support.iParameters) As DTOs.PatientTransferResult

End Function


Private Function GetPatientsBySearchParams(ByRef objPatientParams As PatientSearchParam) As DTOs.Patients

Dim rsPatients As ADODB.Recordset
Dim objPatient As DTOs.Patient
Dim objPatients As DTOs.Patients
Dim sSQL As String
Dim iPatientCount As Integer  '--- SCR 88312   Sherry Wang       2/17/2006
    
    On Error GoTo ErrHandler
    
    '--- SCR 84903   Jane Shen   01/05/2006
    If objPatientParams.SearchADS = True Then
      sSQL = GetSQLPatientSearchFromADS(objPatientParams)
    Else
      sSQL = GetSQLPatientSearch(objPatientParams)
    End If
   
     'No Parameters Set
    If sSQL = vbNullString Then
        Exit Function
    End If
    
    Set rsPatients = mConn.CHCSII_DAS.OpenRecordset(sSQL)
    
    If rsPatients.RecordCount > 0 Then
      Set objPatients = New DTOs.Patients
      
      '--- SCR 88312   Sherry Wang       2/17/2006
      Do While Not rsPatients.EOF And iPatientCount < MAX_PATIENTS
        
        '--- SCR 88312   Sherry Wang       2/17/2006
        iPatientCount = iPatientCount + 1
        Set objPatient = Nothing
        
        Set objPatient = New DTOs.Patient
        
        If objPatientParams.SearchADS = True Then
            '--- SCR 84903   Jane Shen   01/05/2006
            If LoadADSPatientObject(objPatient, rsPatients) Then
                objPatients.Add objPatient ', objPatient.UnitNumber
            End If
        Else
            '--- SCR 71089   Sherry Wang   5/19/2005
            If LoadPatientObject(objPatient, rsPatients, LenB(objPatientParams.UIC) > 0) Then
                If Not objPatients.Exists(objPatient.UnitNumber) Then
                    objPatients.Add objPatient, objPatient.UnitNumber
                End If
            End If
        End If
      
        rsPatients.MoveNext
      Loop
          
    End If
    
 
    
   
    Set GetPatientsBySearchParams = objPatients
    
    GoTo CleanUpAndExit
    
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.GetPatientsBySearchParams", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
    Set rsPatients = Nothing
    Set objPatient = Nothing
    Set objPatients = Nothing
End Function

Private Function GetSQLPatientSearch(ByRef Params As PatientSearchParam) As String
Dim sSQL As String
Dim sSelect As String
Dim sFrom As String
Dim sWhere As String
Dim sTemp As String
Dim bParamExists As Boolean

    On Error GoTo ErrHandler
    
    bParamExists = False
        
    '--- Select
    If BackEndDB = MYSQL Then
        sSelect = "SELECT DATE_FORMAT(ID.BIRTH_DATE, '%m/%d/%Y') BIRTH_DATE, " & vbNewLine
    ElseIf BackEndDB = GEMS Then
        sSelect = "SELECT TO_CHAR(ID.BIRTH_DATE, 'MM/DD/YYYY') BIRTH_DATE, " & vbNewLine
    ElseIf BackEndDB = ITT Then
        sSelect = Replace(sSelect, "TO_CHAR(ID.BIRTH_DATE, 'MM/DD/YYYY') ", vbNullString)
    End If
    If BackEndDB = GEMS Or BackEndDB = MYSQL Then
        '-- SCR 190220 Webb Deers Id Fix
        sSelect = sSelect & "ID.VERIFIED_FLAG, ID.VERIFIED_SOURCE, ID.VERIFIED_DATE, ID.VERIFIED_VERSION,ID.REFERENCE DEERSID, " & vbNewLine
    End If
    sSelect = sSelect & "ID.UNIT_NUMBER, ID.SSN, ID.SEX, NAME.LAST_NAME, " & vbNewLine
    sSelect = sSelect & " NAME.FIRST_NAME, NAME.MIDDLE_NAME, NAME.NAME_TYPE_CID, GID.ID_VALUE AS FMP_SSN," & vbNewLine
    sSelect = sSelect & " HOME_PH.AREA_CODE AS HOME_AREA_CODE, HOME_PH.LOCAL_NUMBER AS HOME_LOCAL_NUMBER, " & vbNewLine
    sSelect = sSelect & " WORK_PH.AREA_CODE AS WORK_AREA_CODE, WORK_PH.LOCAL_NUMBER AS WORK_LOCAL_NUMBER, " & vbNewLine
    sSelect = sSelect & " ADDR.STREET_1, ADDR.STREET_2, ADDR.CITY, ADDR.STATE, ADDR.POST_CODE " & vbNewLine
    
    '--- From
    If BackEndDB = MYSQL Then
'        sFrom = sFrom & "  MMI_ADDRESS ADDR" & vbNewLine
        sFrom = sFrom & " FROM mmi_id id JOIN mmi_name name ON ID.UNIT_NUMBER=NAME.UNIT_NUMBER AND NAME.NAME_TYPE_CID=187"
        sFrom = sFrom & " LEFT JOIN mmi_generic_id gid ON ID.UNIT_NUMBER=GID.UNIT_NUMBER AND GID.ID_TYPE_NCID=" & NCID_FMP_SPONSOR_SSN & vbNewLine
        sFrom = sFrom & " LEFT JOIN mmi_phone home_ph ON ID.UNIT_NUMBER=HOME_PH.UNIT_NUMBER AND HOME_PH.PHONE_TYPE_CID=" & NCID_HOME_PHONE & vbNewLine
        sFrom = sFrom & " LEFT JOIN mmi_phone work_ph ON ID.UNIT_NUMBER=WORK_PH.UNIT_NUMBER AND WORK_PH.PHONE_TYPE_CID=" & NCID_WORK_PHONE & vbNewLine
        sFrom = sFrom & " LEFT JOIN mmi_address addr ON ID.UNIT_NUMBER=ADDR.UNIT_NUMBER AND ADDR.ADDRESS_TYPE_CID=" & NCID_HOME_ADDRESS & vbNewLine
        
        sWhere = " WHERE ID.PATIENT_FLAG='Y' " & vbNewLine
    Else
        sFrom = sFrom & " FROM MMI_ID ID, MMI_NAME NAME, MMI_GENERIC_ID GID, " & vbNewLine
        sFrom = sFrom & " MMI_PHONE HOME_PH, MMI_PHONE WORK_PH, MMI_ADDRESS ADDR" & vbNewLine
        
        sWhere = sWhere & " WHERE " & mobjSQLFix.GetOuterJoin("ID.UNIT_NUMBER", "GID.UNIT_NUMBER") & vbNewLine
        sWhere = sWhere & " AND " & mobjSQLFix.GetOuterJoin(NCID_FMP_SPONSOR_SSN, "GID.ID_TYPE_NCID") & vbNewLine
        sWhere = sWhere & " AND " & mobjSQLFix.GetOuterJoin("ID.UNIT_NUMBER", "HOME_PH.UNIT_NUMBER") & vbNewLine
        sWhere = sWhere & " AND " & mobjSQLFix.GetOuterJoin(NCID_HOME_PHONE, "HOME_PH.PHONE_TYPE_CID") & vbNewLine
        sWhere = sWhere & " AND " & mobjSQLFix.GetOuterJoin("ID.UNIT_NUMBER", "WORK_PH.UNIT_NUMBER") & vbNewLine
        sWhere = sWhere & " AND " & mobjSQLFix.GetOuterJoin(NCID_WORK_PHONE, "WORK_PH.PHONE_TYPE_CID") & vbNewLine
        sWhere = sWhere & " AND " & mobjSQLFix.GetOuterJoin("ID.UNIT_NUMBER", "ADDR.UNIT_NUMBER") & vbNewLine
        sWhere = sWhere & " AND " & mobjSQLFix.GetOuterJoin(NCID_HOME_ADDRESS, "ADDR.ADDRESS_TYPE_CID") & vbNewLine
        sWhere = sWhere & " AND ID.PATIENT_FLAG='Y' " & vbNewLine
        sWhere = sWhere & " AND NAME.NAME_TYPE_CID=187 AND ID.UNIT_NUMBER=NAME.UNIT_NUMBER" & vbNewLine

    End If
    
    '--- SCR 84911  Sherry Wang 1/6/2006
    If Params.LoadUnverified Then
        sWhere = sWhere & " AND ID.VERIFIED_FLAG = 0"
        GetSQLPatientSearch = sSelect & sFrom & sWhere
        Exit Function
    End If
    
    If Params.GetAllPatients Then
        GetSQLPatientSearch = sSelect & sFrom & sWhere
        Exit Function
    End If
        
    If Not Params.ShowVIP Then
        sWhere = sWhere & " AND  NOT EXISTS "
        sWhere = sWhere & " (SELECT 'X' FROM MMI_GENERIC_ID GD  " & vbNewLine
        sWhere = sWhere & " WHERE GD.ID_TYPE_NCID=" & NCID_VIP & " AND  GD.ID_VALUE='Y' " & vbNewLine
        sWhere = sWhere & " AND GD.UNIT_NUMBER = ID.UNIT_NUMBER) " & vbNewLine
    End If
    
    If Params.QuickSearch = True Then
    
        '--- Fist letter of last name + last 4 digits of SSN
        If LenB(Params.LastName) > 0 And LenB(Params.SSN) > 0 Then
            sTemp = Mid$(UCase$(Params.LastName), 1, 1) & Right$(Params.SSN, 4)
            
            If BackEndDB = MYSQL Then
                sFrom = sFrom & " JOIN MMI_GENERIC_ID L_SSN ON L_SSN.UNIT_NUMBER = ID.UNIT_NUMBER" & vbNewLine
            Else
                sFrom = sFrom & ", MMI_GENERIC_ID L_SSN" & vbNewLine
               sWhere = sWhere & " AND L_SSN.UNIT_NUMBER = ID.UNIT_NUMBER" & vbNewLine
            End If
            sWhere = sWhere & " AND L_SSN.ID_TYPE_NCID = " & NCID_FIRST_LETTER_LAST_NAME_LAST_FOUR_PATIENT_SSN & vbNewLine
            sWhere = sWhere & " AND L_SSN.ID_VALUE = '" & sTemp & "'" & vbNewLine
            bParamExists = True
        End If
        
        '--- Fist letter of last name + last 4 digits of SponsorSSN
        If LenB(Params.LastName) > 0 And LenB(Params.SponsorSSN) > 0 Then
            sTemp = Mid$(UCase$(Params.LastName), 1, 1) & Right$(Params.SponsorSSN, 4)
            
            If BackEndDB = MYSQL Then
               sFrom = sFrom & " JOIN MMI_GENERIC_ID LS_SSN ON LS_SSN.UNIT_NUMBER = ID.UNIT_NUMBER" & vbNewLine
            Else
               sFrom = sFrom & ", MMI_GENERIC_ID LS_SSN" & vbNewLine
            sWhere = sWhere & " AND LS_SSN.UNIT_NUMBER = ID.UNIT_NUMBER" & vbNewLine
            End If
            
            sWhere = sWhere & " AND LS_SSN.ID_TYPE_NCID = " & NCID_FIRST_LETTER_LAST_NAME_LAST_FOUR_SPONSOR_SSN & vbNewLine
            sWhere = sWhere & " AND LS_SSN.ID_VALUE = '" & sTemp & "'" & vbNewLine
            
            bParamExists = True
        End If
        
        '--- FMP/SPONSOR_SSN
        If LenB(Params.FMP) > 0 And LenB(Params.SponsorSSN) > 0 Then
            sTemp = Params.FMP & "-" & FormatSSN(Params.SponsorSSN)
            
            If BackEndDB = MYSQL Then
                sFrom = sFrom & " JOIN MMI_GENERIC_ID FMP_SSN ON FMP_SSN.UNIT_NUMBER = ID.UNIT_NUMBER" & vbNewLine
            Else
               sFrom = sFrom & ", MMI_GENERIC_ID FMP_SSN" & vbNewLine
               sWhere = sWhere & " AND FMP_SSN.UNIT_NUMBER = ID.UNIT_NUMBER" & vbNewLine
            End If
            
            sWhere = sWhere & " AND FMP_SSN.ID_TYPE_NCID = " & NCID_FMP_SPONSOR_SSN & vbNewLine
            sWhere = sWhere & " AND FMP_SSN.ID_VALUE = '" & sTemp & "'" & vbNewLine
            
            bParamExists = True
        End If
    Else
        '--- Last Name
        If LenB(Params.LastName) > 0 Then
            sWhere = sWhere & " AND  NAME.LAST_NAME LIKE '" & Replace(Params.LastName, "'", "''") & "%'" & vbNewLine
            bParamExists = True
        End If
        
        '--- First Name
        If LenB(Params.FirstName) > 0 Then
            sWhere = sWhere & " AND  NAME.FIRST_NAME LIKE '" & Replace(Params.FirstName, "'", "''") & "%'" & vbNewLine
            bParamExists = True
        End If
        
        '--- Patient SSN
        If LenB(Params.SSN) > 0 Then
            If BackEndDB = GEMS Or BackEndDB = MYSQL Then
                '--- SCR 48476
                sWhere = sWhere & " AND  ID.SSN = '" & GetSearchSSN(Params.SSN) & "'" & vbNewLine
            Else
                sWhere = sWhere & " AND  ID.SSN = '" & ReturnNumbers(Params.SSN) & "'" & vbNewLine
            End If
            bParamExists = True
        End If
        
        '--- Patient Sex
        If LenB(Params.SEX) > 0 Then
            sWhere = sWhere & " AND  ID.SEX = '" & UCase$(Left$(Params.SEX, 1)) & "'" & vbNewLine
            bParamExists = True
        End If
        
        '--- Facility NCID
        If LenB(Params.FacilityNCID) > 0 Then
            If BackEndDB = MYSQL Then
                sFrom = sFrom & " JOIN MMI_SITE SITE ON ID.UNIT_NUMBER = SITE.UNIT_NUMBER" & vbNewLine
            Else
                sFrom = sFrom & ", MMI_SITE SITE" & vbNewLine
                sWhere = sWhere & " AND ID.UNIT_NUMBER = SITE.UNIT_NUMBER" & vbNewLine
            End If
            
            sWhere = sWhere & " AND SITE.FACILITY_CID = " & Params.FacilityNCID & vbNewLine
            
            bParamExists = True
        End If
        
        '--- Birth Date
        If LenB(Params.DOB) > 0 Then
            sWhere = sWhere & "AND ID.BIRTH_DATE = " & mConn.SQLDate(Params.DOB) & vbNewLine
            bParamExists = True
        End If
        
        '--- UIC
        '--- SCR 53501: Added UPPER() Function
        If LenB(Params.UIC) > 0 Then
            '--- SCR 71089   Sherry Wang   5/19/2005
            sSelect = sSelect & ", UIC.VALUE_TEXT AS UIC" & vbNewLine
            
            If BackEndDB = MYSQL Then
                sFrom = sFrom & " JOIN MMI_GENERICDATA UIC ON ID.UNIT_NUMBER = UIC.UNIT_NUMBER" & vbNewLine
            Else
                sFrom = sFrom & ", MMI_GENERICDATA UIC" & vbNewLine
                sWhere = sWhere & " AND ID.UNIT_NUMBER = UIC.UNIT_NUMBER" & vbNewLine
            End If
            
            sWhere = sWhere & " AND UIC.TYPE_NCID = " & NCID_UNIT_IDENTIFICATION_CODE & vbNewLine
            'SF Removed "UPPER" because it was too slow. 7/7/05
            sWhere = sWhere & " AND UIC.VALUE_TEXT LIKE '" & Params.UIC & "%'" & vbNewLine
            'sWhere = sWhere & " AND UPPER(UIC.VALUE_TEXT) LIKE '" & Params.UIC & "%'" & vbNewLine
                
            bParamExists = True
        End If
    
        '--- SPONSOR_SSN with or without FMP information
        If LenB(Params.SponsorSSN) > 0 Then
        
            If LenB(Params.FMP) > 0 Then
                sTemp = Params.FMP & "/" & FormatSSN(Params.SponsorSSN)
                
                If BackEndDB = MYSQL Then
                    sFrom = sFrom & " JOIN MMI_GENERIC_ID FMP_SSN ON FMP_SSN.UNIT_NUMBER = ID.UNIT_NUMBER" & vbNewLine
                Else
                    sFrom = sFrom & ", MMI_GENERIC_ID FMP_SSN" & vbNewLine
                    sWhere = sWhere & " AND FMP_SSN.UNIT_NUMBER = ID.UNIT_NUMBER" & vbNewLine
                End If
                
                sWhere = sWhere & " AND FMP_SSN.ID_TYPE_NCID = " & NCID_FMP_SPONSOR_SSN & vbNewLine
                sWhere = sWhere & " AND (FMP_SSN.ID_VALUE = '" & sTemp & "'" & vbNewLine
                '--- SCR 53502
                sWhere = sWhere & " or FMP_SSN.ID_VALUE = '" & Replace(sTemp, "-", vbNullString) & "' " & vbNewLine
                sWhere = sWhere & " or FMP_SSN.ID_VALUE = '" & Replace(sTemp, "/", "-") & "')" & vbNewLine

            Else
                sTemp = ReturnNumbers(Params.SponsorSSN)
                
                If BackEndDB = MYSQL Then
                    sFrom = sFrom & " JOIN MMI_GENERIC_ID SP_SSN ON SP_SSN.UNIT_NUMBER = ID.UNIT_NUMBER" & vbNewLine
                Else
                    sFrom = sFrom & ", MMI_GENERIC_ID SP_SSN" & vbNewLine
                    sWhere = sWhere & " AND SP_SSN.UNIT_NUMBER = ID.UNIT_NUMBER" & vbNewLine
                End If
                
                sWhere = sWhere & " AND SP_SSN.ID_TYPE_NCID = " & NCID_SPONSOR_SSN & vbNewLine
                sWhere = sWhere & " AND SP_SSN.ID_VALUE = '" & sTemp & "'" & vbNewLine
            End If
            
            bParamExists = True
        End If
    End If
        
    '--- If no input parameters exists, return NULL string
    If bParamExists = False Then
        sSQL = vbNullString
    Else
        sSQL = sSelect & sFrom & sWhere
    End If
    
    GetSQLPatientSearch = sSQL
    
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.GetSQLPatientSearch", Err.Description, Err.HelpFile, Err.HelpContext
    
End Function




'**********************************************************
' Load Patient Object with the data in Recordset
'**********************************************************
'--- SCR 71089   Sherry Wang   5/19/2005
'--- Add UIC optional parameter
Private Function LoadPatientObject(ByRef objPatient As DTOs.Patient, ByRef oRS As ADODB.Recordset, Optional ByVal bUIC As Boolean = False) As Boolean
    On Error GoTo ErrHandler
    
    With objPatient
    
        .Name.First_Name = ReadStringField(oRS("FIRST_NAME"))
        .Name.Last_Name = ReadStringField(oRS("LAST_NAME"))
        .Name.Middle_Name = ReadStringField(oRS("MIDDLE_NAME"))
        
        .UnitNumber = ReadStringField(oRS("UNIT_NUMBER"))
        .SEX = ReadStringField(oRS("SEX"))
        .SSN = ReadStringField(oRS("SSN"))
        .Formatted_SSN = FormatSSN(.SSN)
        .FMPSponsorSSN = FormatFMP_SSN(ReadStringField(oRS("FMP_SSN")))
        If BackEndDB = GEMS Or BackEndDB = MYSQL Then
            .VerifiedFlag = CBool(ReadLongField(oRS("VERIFIED_FLAG")))
            .VerifiedSource = ReadStringField(oRS("VERIFIED_SOURCE"))
            .VerifiedVersion = ReadStringField(oRS("VERIFIED_VERSION"))
            If oRS("VERIFIED_DATE") <> NULL_DATE And LenB(oRS("VERIFIED_DATE")) > 0 Then
                .VerifiedDate = CDate(oRS("VERIFIED_DATE"))
            End If
            
            
            '-- SCR 190220 B. Webb Populate Deers Id
            .deersID = GetDeersID(ReadStringField(oRS("DEERSID")))
        End If
        
        If Not IsNull(oRS("BIRTH_DATE")) Then
            If oRS("BIRTH_DATE") <> NULL_DATE And LenB(oRS("BIRTH_DATE")) > 0 Then
                .DateOfBirth = CDate(oRS("BIRTH_DATE"))
            End If
        End If
        
        
        .HomeAddress.Street1 = ReadStringField(oRS("STREET_1"))
        .HomeAddress.Street2 = ReadStringField(oRS("STREET_2"))
        .HomeAddress.City = ReadStringField(oRS("CITY"))
        .HomeAddress.State = ReadStringField(oRS("STATE"))
        .HomeAddress.Zip = ReadStringField(oRS("POST_CODE"))
        
        .HomePhone.AreaCode = ReadStringField(oRS("HOME_AREA_CODE"))
        .HomePhone.Phone = ReadStringField(oRS("HOME_LOCAL_NUMBER"))
        
        .WorkPhone.AreaCode = ReadStringField(oRS("WORK_AREA_CODE"))
        .WorkPhone.Phone = ReadStringField(oRS("WORK_LOCAL_NUMBER"))
        
        If bUIC = True Then
            .UnitIdentificationCode = ReadStringField(oRS("UIC"))
        End If
        
    End With
    
    LoadPatientObject = True
    
    GoTo CleanUpAndExit
    
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.LoadPatientObject", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
End Function




Private Function GetPatientsByID(ByVal sUnitNumber As String, Optional ByVal sFacilityNCID As String, Optional ByVal bSponsor As Boolean = False) As DTOs.Patients

Dim rsMain As ADODB.Recordset
Dim rsAddress As ADODB.Recordset
Dim rsPhone As ADODB.Recordset
Dim rsGeneric As ADODB.Recordset
Dim objPatient As DTOs.Patient
Dim objPatients As DTOs.Patients
Dim sSQL As String
Dim objState As DTOs.iState
    
    On Error GoTo ErrHandler
    
    sSQL = GetSQLMain(sUnitNumber, sFacilityNCID)
    
    If sSQL = vbNullString Then
        Exit Function
    End If
    
    Set rsMain = mConn.CHCSII_DAS(BackEndDB).OpenRecordset(sSQL)
    
    If rsMain.RecordCount > 0 Then
    
      '--- Get Generic Data
      sSQL = GetSQLGeneric(sUnitNumber)
      Set rsGeneric = mConn.CHCSII_DAS(BackEndDB).OpenRecordset(sSQL)
      
      '--- Get Address Data
      sSQL = "SELECT UNIT_NUMBER, ADDRESS_TYPE_CID, STREET_1, STREET_2, CITY, COUNTY, STATE, POST_CODE, COUNTRY" & vbNewLine
      sSQL = sSQL & "FROM MMI_ADDRESS WHERE UNIT_NUMBER = " & sUnitNumber
      Set rsAddress = mConn.CHCSII_DAS(BackEndDB).OpenRecordset(sSQL)
      
      '--- Get Phone Data
      sSQL = "SELECT UNIT_NUMBER, PHONE_TYPE_CID, COUNTRY_CODE, AREA_CODE, LOCAL_NUMBER, INTERNAL_EXTENSION" & vbNewLine
      sSQL = sSQL & "FROM MMI_PHONE WHERE UNIT_NUMBER = " & sUnitNumber
      Set rsPhone = mConn.CHCSII_DAS(BackEndDB).OpenRecordset(sSQL)
      
      Set objPatients = New DTOs.Patients
      Set objPatient = Nothing
      Set objPatient = New DTOs.Patient
        
      If LoadPatientObjectExt(objPatient, rsMain, rsGeneric, rsAddress, rsPhone) Then
        If bSponsor Then
            If objPatient.FMP = "20" Then
                Set objPatient.Sponsor = objPatient
            '--- SCR 53350, 53351, 51955
            ElseIf objPatient.FMP = "99" And (BackEndDB = GEMS Or BackEndDB = MYSQL) Then
                Set objPatient.Sponsor = objPatient
            Else
                Set objPatient.Sponsor = GetSponsorBySSN(objPatient.SponsorSSN)
            End If
        End If
        
        Set objState = objPatient
        objState.State = osSaved
        
        objPatients.Add objPatient, objPatient.UnitNumber
      End If
      
          
    End If
   
    Set GetPatientsByID = objPatients
    
    GoTo CleanUpAndExit
    
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.GetPatientsByID", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
    Set rsMain = Nothing
    Set rsGeneric = Nothing
    Set rsAddress = Nothing
    Set rsPhone = Nothing
    Set objPatient = Nothing
    Set objPatients = Nothing
End Function


Private Function GetSQLMain(ByVal sUnitNumber As String, Optional ByVal sFacilityNCID As String) As String
Dim sSQL As String
Dim sSelect As String
Dim sFrom As String
Dim sWhere As String

    On Error GoTo ErrHandler
    
    
    '--- Select
    If BackEndDB = MYSQL Then
        sSelect = "SELECT DATE_FORMAT(ID.BIRTH_DATE, '%m/%d/%Y') BIRTH_DATE, ID.UNIT_NUMBER, ID.SSN, ID.SEX, NAME.LAST_NAME, " & vbNewLine
    Else
        sSelect = "SELECT TO_CHAR(ID.BIRTH_DATE, 'MM/DD/YYYY') BIRTH_DATE, ID.UNIT_NUMBER, ID.SSN, ID.SEX, NAME.LAST_NAME, " & vbNewLine
    End If
    
    sSelect = sSelect & "ID.VERIFIED_FLAG, ID.VERIFIED_SOURCE, ID.VERIFIED_DATE,  ID.VERIFIED_VERSION," & vbNewLine
    sSelect = sSelect & " ID.ADOPTION_FLAG, ID.EXPIRED_FLAG, ID.GUARANTOR_FLAG," & vbNewLine
    sSelect = sSelect & " ID.FOREIGN_ID_FLAG, ID.DEPLOYLOCATION, ID.DEPLOYUNIT, ID.DNBI, " & vbNewLine
    
    '-- SCR 190220 Webb DEERS ID Fix
    'ElseIf BackEndDB = CDR Or BackEndDB = ICD Then
    If BackEndDB <> ITT Then
        sSelect = sSelect & " ID.REFERENCE AS DEERSID, " & vbNewLine
    End If
    
    If BackEndDB = ICD Then
        sSelect = sSelect & " ID.ENTERPRISE_UNIT_NUMBER, " & vbNewLine
    End If
    sSelect = sSelect & " ID.MAIDEN_NAME, ID.MARITAL_STATUS_CID, ID.MOTHER_NAME," & vbNewLine
    sSelect = sSelect & " ID.ORGAN_DONOR_FLAG, ID.RACE_CID, ID.RELIGION_CID, ID.PATIENT_FLAG, ID.UNIDENTIFIED_FLAG," & vbNewLine
    sSelect = sSelect & " NAME.DEGREE, NAME.TITLE, NAME.SUFFIX," & vbNewLine
    sSelect = sSelect & " NAME.FIRST_NAME, NAME.MIDDLE_NAME, NAME.NAME_TYPE_CID, " & vbNewLine
    sSelect = sSelect & " SITE.FACILITY_CID, SITE.MEDICAL_RECORD_NUMBER" & vbNewLine
    
    '--- From
    If BackEndDB = MYSQL Then
        sFrom = sFrom & " FROM mmi_id id JOIN mmi_name name ON ID.UNIT_NUMBER=NAME.UNIT_NUMBER"
        sFrom = sFrom & " JOIN mmi_site site ON (ID.UNIT_NUMBER=SITE.UNIT_NUMBER)"
        If LenB(sFacilityNCID) = 0 Then
            sFrom = sFrom & " AND (SITE.FACILITY_CID=0)" & vbNewLine
        Else
            sFrom = sFrom & " AND (SITE.FACILITY_CID=" & sFacilityNCID & ")" & vbNewLine
        End If
    Else
        sFrom = sFrom & " FROM MMI_ID ID, MMI_NAME NAME, " & vbNewLine
        sFrom = sFrom & "MMI_SITE SITE"
        sWhere = " Where NAME.NAME_TYPE_CID=187 AND ID.UNIT_NUMBER=NAME.UNIT_NUMBER" & vbNewLine
        sWhere = sWhere & " AND" & mobjSQLFix.GetOuterJoin("ID.UNIT_NUMBER", "SITE.UNIT_NUMBER") & vbNewLine
        If LenB(sFacilityNCID) = 0 Then
            sWhere = sWhere & " AND  " & mobjSQLFix.GetOuterJoin("0", "SITE.FACILITY_CID") & vbNewLine
        Else
            sWhere = sWhere & " AND  " & mobjSQLFix.GetOuterJoin(sFacilityNCID, "SITE.FACILITY_CID") & vbNewLine
        End If
    End If
    
    '--- Where
    sWhere = sWhere & " and ID.UNIT_NUMBER = " & sUnitNumber & vbNewLine
    
    
    sSQL = sSelect & sFrom & sWhere
    
    GetSQLMain = sSQL
    
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.GetSQLMain", Err.Description, Err.HelpFile, Err.HelpContext
    
End Function



Private Function GetSQLGeneric(ByVal sUnitNumber As String) As String
Dim sSQL As String

    On Error GoTo ErrHandler
    
    sSQL = "SELECT UNIT_NUMBER, 0 AS ORDINAL, ID_TYPE_NCID TYPE_NCID, 0 VALUE_NCID, ID_VALUE VALUE_TEXT" & vbNewLine
    sSQL = sSQL & "  FROM MMI_GENERIC_ID" & vbNewLine
    sSQL = sSQL & "  WHERE UNIT_NUMBER = " & sUnitNumber & vbNewLine
    sSQL = sSQL & "  UNION ALL" & vbNewLine
    sSQL = sSQL & "  SELECT MG.UNIT_NUMBER, ORDINAL, TYPE_NCID, VALUE_NCID, VALUE_TEXT" & vbNewLine
    sSQL = sSQL & "  FROM MMI_GENERICDATA MG" & vbNewLine
    sSQL = sSQL & "  WHERE UNIT_NUMBER = " & sUnitNumber & vbNewLine

    GetSQLGeneric = sSQL
    
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.GetSQLGeneric", Err.Description, Err.HelpFile, Err.HelpContext
    
End Function




'**********************************************************
' Load Extended Patient Object with the data in Recordset
'**********************************************************

Private Function LoadPatientObjectExt(ByRef objPatient As DTOs.Patient, _
                                      ByRef oRsMain As ADODB.Recordset, _
                                      ByRef oRsGeneric As ADODB.Recordset, _
                                      ByRef oRsAddress As ADODB.Recordset, _
                                      ByRef oRsPhone As ADODB.Recordset) As Boolean


    On Error GoTo ErrHandler
    
    With objPatient
        With .Name
            .First_Name = ReadStringField(oRsMain("FIRST_NAME"))
            .Last_Name = ReadStringField(oRsMain("LAST_NAME"))
            .Middle_Name = ReadStringField(oRsMain("MIDDLE_NAME"))
            .Degree = ReadStringField(oRsMain("DEGREE"))
            
            .Title = ReadStringField(oRsMain("TITLE"))
            .Suffix = ReadStringField(oRsMain("SUFFIX"))
            .Name_Type_NCID = ReadStringField(oRsMain("NAME_TYPE_CID"))
        End With
        
        If BackEndDB = GEMS Or BackEndDB = MYSQL Then
            .VerifiedFlag = CBool(ReadLongField(oRsMain("VERIFIED_FLAG")))
            .VerifiedSource = ReadStringField(oRsMain("VERIFIED_SOURCE"))
            .VerifiedVersion = ReadStringField(oRsMain("VERIFIED_VERSION"))
            If oRsMain("VERIFIED_DATE") <> NULL_DATE And LenB(oRsMain("VERIFIED_DATE")) > 0 Then
                .VerifiedDate = CDate(oRsMain("VERIFIED_DATE"))
            End If
            
        End If
        
        .UnitNumber = oRsMain("UNIT_NUMBER").Value
        
        .deersID = GetDeersID(ReadStringField(oRsMain("DEERSID"))) '-- SCR 190439 Webb - Include Deers Id (JPU v2 message)
        
        .SEX = ReadStringField(oRsMain("SEX"))
        .SSN = ReadStringField(oRsMain("SSN"))
        
        .Formatted_SSN = FormatSSN(.SSN)
        
        .Adopted = ReadStringField(oRsMain("ADOPTION_FLAG"))
        .Expired = ReadStringField(oRsMain("EXPIRED_FLAG"))
        .OrganDonor = ReadStringField(oRsMain("ORGAN_DONOR_FLAG"))
        .Guarantor = ReadStringField(oRsMain("GUARANTOR_FLAG"))
        .MaidenName = ReadStringField(oRsMain("MAIDEN_NAME"))
        .MothersMaidenName = ReadStringField(oRsMain("MOTHER_NAME"))
        .MaritalStatus = ReadStringField(oRsMain("MARITAL_STATUS_CID"))
        .PatientFlag = ReadStringField(oRsMain("PATIENT_FLAG"))
        .UnidentifiedFlag = ReadStringField(oRsMain("UNIDENTIFIED_FLAG"))
        

        If IsNull(oRsMain("FOREIGN_ID_FLAG")) Then
            .ForeignID_Flag = False
        Else
            .ForeignID_Flag = CBool(oRsMain("FOREIGN_ID_FLAG"))
        End If
        
        .DeployLocation = ReadStringField(oRsMain("DEPLOYLOCATION"))
        .DeployUnit = ReadStringField(oRsMain("DEPLOYUNIT"))
        .DNBI = ReadStringField(oRsMain("DNBI"))
      
        'This If-Else could be reworked to use ReadStringField
        If Not IsNull(oRsMain("FACILITY_CID")) Then
            If Val(oRsMain("FACILITY_CID")) > 0 Then
                .FacilityNCID = oRsMain("FACILITY_CID")
                .Facility = mobjPicklist.ReturnRepFromNCID(NCID_FACILITY_DOMAIN, oRsMain("FACILITY_CID"))
                If LenB(objPatient.Facility) = 0 Then
                    .Facility = GetOSH_INI_SETTINGS("FACILITY_NAME")
                End If
            End If
        Else
            .Facility = vbNullString
        End If
        
        .MedicalRecordNumber = ReadStringField(oRsMain("MEDICAL_RECORD_NUMBER"))
        .Race = ReadStringField(oRsMain("RACE_CID"))
        .Religion = ReadStringField(oRsMain("RELIGION_CID"))
        
        If Not IsNull(oRsMain("BIRTH_DATE")) Then
            If oRsMain("BIRTH_DATE") <> NULL_DATE And LenB(oRsMain("BIRTH_DATE")) > 0 Then
                .DateOfBirth = CDate(oRsMain("BIRTH_DATE"))
            End If
        End If
    End With
    
    '--- Addresses
    LoadAddress objPatient, oRsAddress
    '--- Phones
    LoadPhones objPatient, oRsPhone
    '--- GenericData
    LoadGenericData objPatient, oRsGeneric
                
    LoadPatientObjectExt = True
    
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.LoadPatientObjectExt", Err.Description, Err.HelpFile, Err.HelpContext
End Function


Private Function GetPayGradeFromRank(ByVal sRank As String, ByVal sRankNcid As String) As String
  
Dim oRS As ADODB.Recordset
Dim sSQL As String
    
    On Error GoTo ErrHandler
    
    If Val(sRankNcid) > 0 Then
        sSQL = "SELECT PAYGRADE FROM RANK_ABBREV WHERE NCID = " & sRankNcid
        Set oRS = mConn.CHCSII_DAS.OpenRecordset(sSQL)
        
        If oRS.RecordCount > 0 Then
            GetPayGradeFromRank = ReadStringField(oRS("PAYGRADE"))
        End If
    End If
    
    If LenB(GetPayGradeFromRank) = 0 Then
        If LenB(sRank) > 0 Then
            sSQL = "SELECT PAYGRADE FROM RANK_ABBREV WHERE FULL_RANK LIKE '" & sRank & "%'"
            Set oRS = mConn.CHCSII_DAS.OpenRecordset(sSQL)
            
            If oRS.RecordCount > 0 Then
                GetPayGradeFromRank = ReadStringField(oRS("PAYGRADE"))
            End If
        End If
    End If
    
    GoTo CleanUpAndExit
    
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.GetPayGradeFromRank", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
    Set oRS = Nothing
End Function







'**********************************************************
' Get PCM Information
'**********************************************************
Private Function Get_PCM(ByVal sPCM_Code As String) As ADODB.Recordset

Dim sSQL As String

    On Error GoTo ErrHandler

    sSQL = "SELECT PCM_CODE, NAME, IEN FROM PROVIDER WHERE PCM_CODE = '" & sPCM_Code & "'"
      
    Set Get_PCM = mConn.CHCSII_DAS(BackEndDB).OpenRecordset(sSQL)
   
    GoTo CleanUpAndExit
    
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.Get_PCM", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
End Function



Private Function GetSponsorBySSN(ByVal sSSN As String) As DTOs.Patient

Dim rsMain As ADODB.Recordset
Dim rsGeneric As ADODB.Recordset
Dim objSponsor As DTOs.Patient
Dim sSQL As String
    
    On Error GoTo ErrHandler
    
    sSQL = GetSQLSponsor(sSSN)
    
    If sSQL = vbNullString Then
        Exit Function
    End If
    
    Set rsMain = mConn.CHCSII_DAS(BackEndDB).OpenRecordset(sSQL)
    
    
    If rsMain.RecordCount > 0 Then
    
      '--- Get Generic Data for Sponsor
      sSQL = GetSQLGenericSponsor(rsMain("UNIT_NUMBER"))
      Set rsGeneric = mConn.CHCSII_DAS(BackEndDB).OpenRecordset(sSQL)
             
      Set objSponsor = New DTOs.Patient
        
      If LoadSponsorObject(objSponsor, rsMain, rsGeneric) Then
          Set GetSponsorBySSN = objSponsor
      End If
      
          
    End If
   
    
    GoTo CleanUpAndExit
    
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.GetSponsorBySSN", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
    Set rsMain = Nothing
    Set rsGeneric = Nothing
    Set objSponsor = Nothing
End Function




Private Function GetSQLGenericSponsor(ByVal sUnitNumber As String) As String
Dim sSQL As String

    On Error GoTo ErrHandler
    
    sSQL = "SELECT UNIT_NUMBER, ORDINAL, TYPE_NCID, VALUE_NCID, VALUE_TEXT" & vbNewLine
    sSQL = sSQL & " FROM MMI_GENERICDATA " & vbNewLine
    sSQL = sSQL & " WHERE UNIT_NUMBER = " & sUnitNumber & vbNewLine
    sSQL = sSQL & " AND TYPE_NCID IN (" & NCID_UNIT_IDENTIFICATION_CODE & ", " & NCID_PATIENT_CATEGORY & ", " & NCID_RANK & ")" & vbNewLine

    GetSQLGenericSponsor = sSQL
    
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.GetSQLGenericSponsor", Err.Description, Err.HelpFile, Err.HelpContext
    
End Function


Private Function GetSQLSponsor(ByVal sSSN As String) As String
Dim sSQL As String

    On Error GoTo ErrHandler
    
    sSQL = "SELECT  ID.UNIT_NUMBER, ID.SSN, ID.BIRTH_DATE, " & vbNewLine
    sSQL = sSQL & " NAME.LAST_NAME, NAME.FIRST_NAME, NAME.MIDDLE_NAME, NAME.NAME_TYPE_CID " & vbNewLine
    sSQL = sSQL & " FROM MMI_ID ID, MMI_NAME NAME " & vbNewLine
    sSQL = sSQL & " WHERE  ID.SSN = '" & sSSN & "'" & vbNewLine
    sSQL = sSQL & " AND NAME.NAME_TYPE_CID=187 AND ID.UNIT_NUMBER=NAME.UNIT_NUMBER" & vbNewLine
    
    GetSQLSponsor = sSQL
    
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.GetSQLSponsor", Err.Description, Err.HelpFile, Err.HelpContext
    
End Function





'**********************************************************
' Load Extended Patient Object with the data in Recordset
'**********************************************************

Private Function LoadSponsorObject(ByRef objSponsor As DTOs.Patient, _
                                   ByRef oRsMain As ADODB.Recordset, _
                                   ByRef oRsGeneric As ADODB.Recordset) As Boolean
    
    Dim oGenericData As DTOs.PatientGenericData
    
    On Error GoTo ErrHandler
    
    With objSponsor
        .Name.First_Name = ReadStringField(oRsMain("FIRST_NAME"))
        .Name.Last_Name = ReadStringField(oRsMain("LAST_NAME"))
        .Name.Middle_Name = ReadStringField(oRsMain("MIDDLE_NAME"))
        .Name.Name_Type_NCID = ReadStringField(oRsMain("NAME_TYPE_CID"))
            
        .UnitNumber = oRsMain("UNIT_NUMBER")
        .SSN = ReadStringField(oRsMain("SSN"))
        .Formatted_SSN = FormatSSN(.SSN)
        
        
        If Not IsNull(oRsMain("BIRTH_DATE")) Then
            If oRsMain("BIRTH_DATE") <> NULL_DATE And LenB(oRsMain("BIRTH_DATE")) > 0 Then
                .DateOfBirth = oRsMain("BIRTH_DATE")
            End If
        End If
        
        If oRsGeneric.RecordCount > 0 Then
          oRsGeneric.MoveFirst
          
          Do While Not oRsGeneric.EOF
              Set oGenericData = Nothing
              Set oGenericData = New DTOs.PatientGenericData
                
              With oGenericData
                .Ordinal = ReadLongField(oRsGeneric("ORDINAL"))
                .TypeNcid = ReadLongField(oRsGeneric("TYPE_NCID"))
                .ValueNcid = ReadLongField(oRsGeneric("VALUE_NCID"))
                .ValueText = ReadStringField(oRsGeneric("VALUE_TEXT"))
              End With
                
             .GenericDatas.Add oGenericData, oGenericData.Ordinal
           
           
            Select Case oRsGeneric("TYPE_NCID")
              Case NCID_RANK
                .RankNCID = ReadStringField(oRsGeneric("VALUE_NCID"))
                .Rank = ReadStringField(oRsGeneric("VALUE_TEXT"))
                If LenB(oRsGeneric("VALUE_TEXT")) > 0 Then
                    .PayGrade = GetPayGradeFromRank(oRsGeneric("VALUE_TEXT"), .Rank)
                End If
              Case NCID_UNIT_IDENTIFICATION_CODE
                .UnitIdentificationCode = ReadStringField(oRsGeneric("VALUE_TEXT"))
              Case NCID_PATIENT_CATEGORY
                .PatientCategory = ReadStringField(oRsGeneric("VALUE_TEXT"))
                
                If LenB(oRsGeneric("VALUE_TEXT")) > 0 Then
                    .BranchOfService = GetBranchOfService(.PatientCategory)
                End If
            End Select
        
            oRsGeneric.MoveNext
        Loop
        
      End If
        
    End With
    
    LoadSponsorObject = True
    
    GoTo CleanUpAndExit
    
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.LoadSponsorObject", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
    Set oGenericData = Nothing
End Function



'**********************************************************
' Load Pregnancy Object with the data in the big pregnancy string
'**********************************************************

Private Function LoadPregnancyObject(ByVal sPregData As String, ByRef objPregancy As DTOs.Pregnancy) As Boolean
    
    Dim sPregnancy() As String
    Dim sTemp() As String
    Dim i As Integer
    
    On Error GoTo ErrHandler
    
    sPregnancy = Split(sPregData, ", ")
    
    Set objPregancy = New DTOs.Pregnancy
     
    With objPregancy
        For i = 0 To UBound(sPregnancy)
            sTemp = Split(sPregnancy(i), "=")
            
            If Left$(sTemp(0), 8) = "BCMethod" Then
                .BirthControlMethods.Add sTemp(1), Mid$(sTemp(0), 9)
            Else
              Select Case sTemp(0)
                Case "PregnancyIndicator"
                    .PregnancyIndicator = CBool(sTemp(1))
                Case "EstimatedDOB"
                    .EstimatedDOB = sTemp(1)
                Case "LastMenstrualPeriod"
                    .LastMenstrualPeriod = sTemp(1)
                Case "MenopauseIndicator"
                    .MenopauseIndicator = CBool(sTemp(1))
                Case "HysterectomyIndicator"
                    .HysterectomyIndicator = CBool(sTemp(1))
                Case "TotalAbortus"
                    .TotalAbortus = sTemp(1)
                Case "SpontaneousAbortions"
                    .SpontaneousAbortions = sTemp(1)
                Case "ElectiveAbortions"
                    .ElectiveAbortions = sTemp(1)
                Case "EctopicAbortions"
                    .EctopicAbortions = sTemp(1)
                Case "Gravida"
                    .Gravida = sTemp(1)
                Case "Para"
                    .Para = sTemp(1)
                Case "LivingChildren"
                    .LivingChildren = sTemp(1)
              End Select
            End If
        Next i
    End With
    
    
    LoadPregnancyObject = True
    
    GoTo CleanUpAndExit
    
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.LoadPregnancyObject", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
End Function


Private Function GetNewPatients(ByRef oParams As PatientSearchParam) As DTOs.Patients

Dim objPatient As DTOs.Patient
Dim objPatients As DTOs.Patients
Dim objState As DTOs.iState
    
    On Error GoTo ErrHandler
          
    Set objPatients = New DTOs.Patients
    Set objPatient = New DTOs.Patient
    Set objState = objPatient
      
    objPatient.UnitNumber = GetNewUnitNum
    objPatient.SSN = GetNextSSN
    objPatient.FacilityNCID = oParams.FacilityNCID
    objPatient.Facility = mobjPicklist.ReturnRepFromNCID(NCID_FACILITY_DOMAIN, oParams.FacilityNCID)
    If LenB(objPatient.Facility) = 0 Then
        objPatient.Facility = GetOSH_INI_SETTINGS("FACILITY_NAME")
    End If
    
    objState.State = osNew
    
    objPatients.Add objPatient, objPatient.UnitNumber
    
    Set GetNewPatients = objPatients
    
    GoTo CleanUpAndExit
    
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.GetNewPatients", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
End Function


'*****************************************************************************
'  Function Name: GetNewUnitNum()
'
'  Description: This function retrieves the next Unit_Number from the
'       Sequences table.
'*****************************************************************************
Private Function GetNewUnitNum() As String
    If mISQLOps Is Nothing Then
        Set mISQLOps = mConn.CHCSII_SQLOPS
    End If
    
    GetNewUnitNum = mISQLOps.GetNextID(esUNIT_NUMBER)
End Function



'*****************************************************************************
'  Function Name: sGetNextSSN()
'
'  Description: This function retrieves the NEXT_DEFAULT_SSN from the
'       Sequences table.
'
'*****************************************************************************
Private Function GetNextSSN() As String
    If mISQLOps Is Nothing Then
        Set mISQLOps = mConn.CHCSII_SQLOPS
    End If
    GetNextSSN = mISQLOps.GetNextID(esTEMP_SSN)
End Function


Private Function Patient_Exists(ByVal sUnitNumber As String) As Boolean

    Dim oRS As ADODB.Recordset
    Dim sSQL As String
    
    On Error GoTo ErrHandler
    
    sSQL = "SELECT UNIT_NUMBER FROM MMI_ID WHERE UNIT_NUMBER = " & sUnitNumber
    
    Set oRS = mConn.CHCSII_DAS(BackEndDB).OpenRecordset(sSQL)
    
    If oRS.RecordCount = 0 Then
        Patient_Exists = False
    Else
        Patient_Exists = True
    End If
    
    GoTo CleanUpAndExit
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.Patient_Exists", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
    Set oRS = Nothing
End Function



Private Sub Save_MMI_ID(ByRef oPatient As DTOs.Patient, ByVal sAction As String)

    Dim ActionClause As String
    Dim WhereClause As String
    Dim ParseChar As String
    Dim Values As String
    
    On Error GoTo ErrHandler
    
    ParseChar = Chr$(255)
    
    If IsNull(oPatient.UnitNumber) Then
        oPatient.UnitNumber = GetNewUnitNum
    End If
    
    If Patient_Exists(oPatient.UnitNumber) Then
        ActionClause = "UPDATE"
        WhereClause = "WHERE Unit_Number = " & oPatient.UnitNumber
    Else
        ActionClause = "INSERT INTO"
    End If
    
    If oPatient.DateOfBirth > 0 Then
        Values = Values & "Birth_Date = " & mConn.SQLDate(oPatient.DateOfBirth) & ParseChar
    End If
    
    If LenB(Trim$(oPatient.MaritalStatus)) > 0 Then
        Values = Values & "Marital_Status_CID = " & _
                    RTrim$(oPatient.MaritalStatus) & "" & ParseChar
    End If
    If LenB(Trim$(oPatient.Race)) > 0 Then
        Values = Values & "Race_cid = " & Trim$(oPatient.Race) & "" & ParseChar
    End If
    If LenB(Trim$(oPatient.Religion)) > 0 Then
        Values = Values & "Religion_CID = " & Trim$(oPatient.Religion) & "" & ParseChar
    End If
    If LenB(Trim$(oPatient.SEX)) > 0 Then
        Values = Values & "Sex = '" & Trim$(oPatient.SEX) & "'" & ParseChar
    End If
    If LenB(Trim$(oPatient.SSN)) > 0 Then
        Values = Values & "Foreign_ID_Flag = " & CInt(oPatient.ForeignID_Flag) & ParseChar
        Values = Values & "SSN = '" & Trim$(oPatient.SSN) & "'" & ParseChar
    End If
    
    Values = Values & "DeployUnit = " & mConn.SQLQuote(Trim$(oPatient.DeployUnit)) & ParseChar
    
    If LenB(Trim$(oPatient.DeployLocation)) > 0 Then
        Values = Values & "DeployLocation = " & mConn.SQLQuote(Trim$(oPatient.DeployLocation)) & ParseChar
    End If
    If LenB(Trim$(oPatient.DNBI)) > 0 Then
        Values = Values & "DNBI = " & mConn.SQLQuote(Trim$(oPatient.DNBI)) & ParseChar
    End If
    If LenB(Trim$(oPatient.Adopted)) > 0 Then
        Values = Values & "Adoption_Flag = '" & Trim$(oPatient.Adopted) & "'" & ParseChar
    End If
    If LenB(Trim$(oPatient.Expired)) > 0 Then
        Values = Values & "Expired_Flag = '" & Trim$(oPatient.Expired) & "'" & ParseChar
    End If
    If LenB(Trim$(oPatient.Guarantor)) > 0 Then
        Values = Values & "Guarantor_Flag = '" & Trim$(oPatient.Guarantor) & "'" & ParseChar
    End If
    If LenB(Trim$(oPatient.OrganDonor)) > 0 Then
        Values = Values & "Organ_Donor_Flag = '" & Trim$(oPatient.OrganDonor) & "'" & ParseChar
    End If
    
    If LenB(Trim$(oPatient.PatientFlag)) > 0 Then
        Values = Values & "Patient_Flag = '" & Trim$(oPatient.PatientFlag) & "'" & ParseChar
    Else
        Values = Values & "Patient_Flag = 'Y'" & ParseChar
    End If
    
    If LenB(Trim$(oPatient.UnidentifiedFlag)) > 0 Then
        Values = Values & "Unidentified_Flag = '" & Trim$(oPatient.UnidentifiedFlag) & "'" & ParseChar
    End If
    If LenB(Trim$(oPatient.MothersMaidenName)) > 0 Then
        Values = Values & "Mother_Name = " & mConn.SQLQuote(Trim$(oPatient.MothersMaidenName)) & ParseChar
    End If
    
    If LenB(Trim$(oPatient.MaidenName)) > 0 Then
        Values = Values & "Maiden_Name = " & mConn.SQLQuote(Trim$(oPatient.MaidenName)) & ParseChar
    End If
    
    If oPatient.VerifiedFlag = True Then
        Values = Values & "VERIFIED_FLAG = 1" & ParseChar
        Values = Values & "VERIFIED_DATE = SYSDATE" & ParseChar
        '-- SCR 190220 B. Webb DEERS Id fix
        If LenB(oPatient.deersID) > 0 Then
            Values = Values & "REFERENCE = " & mConn.SQLQuote("D" & oPatient.deersID) & ParseChar
        Else
            Values = Values & "Reference = " & mConn.SQLQuote(oPatient.UnitNumber) & ParseChar
        End If
        Dim oRS As ADODB.Recordset
        Set oRS = GetVerifySourceInfo
        If Not oRS Is Nothing Then
            oRS.MoveFirst
            Values = Values & "VERIFIED_SOURCE = " & mConn.SQLQuote(Trim$(ReadStringField(oRS("DATA_SOURCE")))) & ParseChar
            Values = Values & "VERIFIED_VERSION = " & mConn.SQLQuote(Trim$(ReadStringField(oRS("VERIFY_VERSION")))) & ParseChar
        End If
    Else
        Values = Values & "VERIFIED_FLAG = 0" & ParseChar
        Values = Values & "VERIFIED_DATE = NULL" & ParseChar
        Values = Values & "VERIFIED_SOURCE = NULL" & ParseChar
        Values = Values & "VERIFIED_VERSION = NULL" & ParseChar
        Values = Values & "Reference = " & oPatient.UnitNumber & ParseChar
    End If

    
    Values = Values & "Unit_Number = " & oPatient.UnitNumber & ParseChar
    Values = Values & "Create_Audit_Num = 576" & ParseChar
    
    '*-- No ParseChar at end of last values term.
    Values = Values & "UpdateFlag = '" & sAction & "' "
    
    DoWrite ActionClause, WhereClause, "MMI_ID", Values, ParseChar
    
    GoTo CleanUpAndExit
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.Save_MMI_ID", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
End Sub


Private Function DeletePatient(ByVal sUnitNumber As String) As Boolean

    Dim oRS As ADODB.Recordset
    Dim sSetString As String
    Dim mobjDAS As ICHCSII_DAS
    
    On Error GoTo ErrHandler
    
    Set mobjDAS = mConn.CHCSII_DAS(BackEndDB)
    mobjDAS.BeginTrans
        sSetString = " where unit_number = " & sUnitNumber
        mobjDAS.ExecuteSQL "delete from mmi_name " & sSetString
        mobjDAS.ExecuteSQL "delete from mmi_address " & sSetString
        mobjDAS.ExecuteSQL "delete from mmi_phone " & sSetString
        mobjDAS.ExecuteSQL "delete from mmi_genericdata " & sSetString
        mobjDAS.ExecuteSQL "delete from mmi_generic_id " & sSetString
        mobjDAS.ExecuteSQL "delete from mmi_site " & sSetString
        mobjDAS.ExecuteSQL "delete from mmi_id " & sSetString
    mobjDAS.CommitTrans
    
    DeletePatient = True
    
    GoTo CleanUpAndExit
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.DeletePatient", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
    Set oRS = Nothing
    Set mobjDAS = Nothing
End Function





Private Sub DoWrite(ByVal ActionClause As String, _
                   ByVal WhereClause As String, _
                   ByVal TableName As String, _
                   ByVal Values As String, _
                   ByVal ParseChar As String)


    Dim Parms() As String
    Dim xFields As String
    Dim FieldValues As String
    Dim SelectedFields As String
    Dim xloc As Integer
    Dim sepchar As String
    Dim ParmSet As Variant
     
    Select Case UCase$(ActionClause)
        Case "INSERT INTO"        ' Add a Record
            
            '--- Add a Record...Split Field1=Value1 ParseChar Field2=Value2....
            '--- Into (Field1, Field2...) and  (Value1,Value2...)
            Parms = Split(Values, ParseChar)
                                                    
            For Each ParmSet In Parms
             
               '--- Split Field and it's value...Field1 = Value1
               xloc = InStr(ParmSet, " = ")
               xFields = xFields & sepchar & Trim$(Mid$(ParmSet, 1, xloc - 1))
               FieldValues = FieldValues & sepchar & Trim$(Mid$(ParmSet, xloc + 3))
               sepchar = ", "
               
            Next ParmSet
             
            Debug.Print "ExecuteSQL: " & ActionClause & " " & TableName & " (" & xFields & ")" & " VALUES (" & FieldValues & ")"
            mobjDAS.ExecuteSQL ActionClause & " " & TableName & " (" & xFields & ")" & " VALUES (" & FieldValues & ")"
             
        Case "UPDATE"            ' Update a Record
            
            SelectedFields = Replace(Values, ParseChar, ", ", 1, -1, vbBinaryCompare)
             
            Debug.Print "ExecuteSQL: " & ActionClause & " " & TableName & " SET " & SelectedFields & " " & WhereClause
            mobjDAS.ExecuteSQL ActionClause & " " & TableName & " SET " & SelectedFields & " " & WhereClause
             
        Case Else
            
    End Select
            
            
End Sub



Private Sub Save_MMI_NAME(ByRef oPatient As DTOs.Patient)

    Dim ActionClause As String
    Dim WhereClause As String
    Dim ParseChar As String
    Dim sAction As String
    Dim sSQL As String
    Dim Values As String
    
    On Error GoTo ErrHandler
    
    ParseChar = Chr$(255)
    
    sSQL = "Select count(*) num from MMI_Name Where Unit_Number = " & oPatient.UnitNumber
    sSQL = sSQL & " And Name_Type_cid = " & oPatient.Name.Name_Type_NCID
    
    If mobjDAS.OpenRecordset(sSQL).Fields("num") > 0 Then
       sAction = "U"
    Else
       sAction = "N"
    End If
    
    If sAction = "N" Then
        ActionClause = "INSERT INTO"
    ElseIf sAction = "U" Then
        ActionClause = "UPDATE"
        WhereClause = "WHERE Unit_Number = " & oPatient.UnitNumber
    End If
    
    With oPatient.Name
        If LenB(Trim$(.Degree)) > 0 Then
            Values = Values & "Degree = " & mConn.SQLQuote(Trim$(.Degree)) & ParseChar
        End If
        If LenB(Trim$(.Last_Name)) > 0 Then
            Values = Values & "last_name = " & mConn.SQLQuote(Trim$(.Last_Name)) & ParseChar
        End If
        If LenB(Trim$(.First_Name)) > 0 Then
            Values = Values & "first_name = " & mConn.SQLQuote(Trim$(.First_Name)) & ParseChar
        End If
        If LenB(Trim$(.Middle_Name)) > 0 Then
            Values = Values & "middle_name = " & mConn.SQLQuote(Trim$(.Middle_Name)) & ParseChar
        End If
        If LenB(Trim$(.Suffix)) > 0 Then
            Values = Values & "Suffix = " & mConn.SQLQuote(Trim$(.Suffix)) & ParseChar
        End If
        If LenB(Trim$(.Title)) > 0 Then
            Values = Values & "Title = " & mConn.SQLQuote(Trim$(.Title)) & ParseChar
        End If
        If LenB(.Name_Type_NCID) > 0 Then
            Values = Values & "name_type_cid = " & .Name_Type_NCID & "" & ParseChar
            WhereClause = WhereClause & " and name_type_cid = " & .Name_Type_NCID

        Else
            Values = Values & "name_type_cid = " & NCID_CURRENT_NAME & "" & ParseChar
            WhereClause = WhereClause & " and name_type_cid = " & NCID_CURRENT_NAME
        End If
        
        '*-- Coleman 062101:
        Values = Values & "CREATE_AUDIT_NUM = 576" & ParseChar
        Values = Values & "Unit_Number = " & oPatient.UnitNumber & "" & ParseChar
        Values = Values & "Name_Index = 1" & "" & ParseChar
        Values = Values & "updateflag = '" & sAction & "' "
        
        DoWrite ActionClause, WhereClause, "MMI_Name", Values, ParseChar
    End With
        
    GoTo CleanUpAndExit
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.Save_MMI_NAME", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
End Sub

'PATIENT PROFILE
Private Sub SavePatientProfile(ByRef oPatient As DTOs.Patient)
    Dim sSQL As String
    Dim oRS As ADODB.Recordset
    Dim oStream As ADODB.Stream
    Dim oStreamThumbnail As ADODB.Stream
    
    On Error GoTo ErrHandler
    
    Set oRS = New ADODB.Recordset
    Set oStream = New ADODB.Stream
    oStream.Type = adTypeBinary

    Set oStreamThumbnail = New ADODB.Stream
    oStreamThumbnail.Type = adTypeBinary
    
    'Patient Profile
    If BackEndDB = MYSQL Then
        If oPatient.ProfileImageChanged Then
           
           sSQL = "Select * from MMI_Name Where Unit_Number = " & oPatient.UnitNumber
           sSQL = sSQL & " And Name_Type_cid = " & oPatient.Name.Name_Type_NCID
    
           Set oRS = mobjDAS.OpenRecordset(sSQL, adOpenKeyset, adLockOptimistic)
            
           oStream.Open
           oStream.LoadFromFile oPatient.ProfileImageFile
           oRS.Fields("PROFILE_PIC").Value = oStream.Read

           oStreamThumbnail.Open
           oStreamThumbnail.LoadFromFile oPatient.ProfileThumbnailFile
           oRS.Fields("PROFILE_THUMBIMG").Value = oStreamThumbnail.Read
            
           oRS.Update
           oStream.Close
           oStreamThumbnail.Close
           oRS.Close
        End If
    End If
    
    GoTo CleanUpAndExit
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.Save_MMI_NAME", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
    Set oRS = Nothing
    Set oStream = Nothing
    Set oStreamThumbnail = Nothing
End Sub

Private Sub Save_MMI_ADDRESS(ByRef oPatient As DTOs.Patient, ByVal sAction As String)

    Dim ActionClause As String
    Dim WhereClause As String
    Dim ParseChar As String
    Dim sSQL As String
    Dim Values As String
    Dim oAddr As DTOs.Address
    Dim sSetString As String
    
    On Error GoTo ErrHandler
    
    ParseChar = Chr$(255)
    
    If oPatient.Addresses.Count = 0 Then
        GoTo CleanUpAndExit
    End If
    
    sSetString = " where unit_number = " & oPatient.UnitNumber
    
    For Each oAddr In oPatient.Addresses
        With oAddr
            Values = vbNullString
            
            '--- Delete the addresses for the patient
            sSQL = "delete from MMI_Address " & sSetString & " and Address_Type_cid = " & .AddrType
            mobjDAS.ExecuteSQL sSQL
            
            '--- Do the insert
            sAction = "N"
                
            ActionClause = "INSERT INTO"
            
            If LenB(ActionClause) > 0 Then
                If LenB(Trim$(.AddrType)) > 0 Then
                    Values = Values & "address_type_cid = " & Trim$(.AddrType) & "" & ParseChar
                End If
                If LenB(Trim$(.Street1)) > 0 Then
                    Values = Values & "street_1 = " & mConn.SQLQuote(Trim$(.Street1)) & ParseChar
                End If
                If LenB(Trim$(.Street2)) > 0 Then
                    Values = Values & "street_2 = " & mConn.SQLQuote(Trim$(.Street2)) & ParseChar
                End If
                If LenB(Trim$(.City)) > 0 Then
                    Values = Values & "city = " & mConn.SQLQuote(Trim$(.City)) & ParseChar
                End If
                If LenB(Trim$(.County)) > 0 Then
                    Values = Values & "county = " & mConn.SQLQuote(Trim$(.County)) & ParseChar
                End If
                If LenB(Trim$(.State)) > 0 Then
                    Values = Values & "state = " & mConn.SQLQuote(Trim$(.State)) & ParseChar
                End If
                If LenB(Trim$(.Zip)) > 0 Then
                    Values = Values & "post_code = '" & Trim$(.Zip) & "'" & ParseChar
                End If
                If LenB(Trim$(.Country)) > 0 Then
                    Values = Values & "country = " & mConn.SQLQuote(Trim$(.Country)) & ParseChar
                End If
                
                Values = Values & "CREATE_AUDIT_NUM = 576" & ParseChar
                Values = Values & "Unit_Number = " & oPatient.UnitNumber & ParseChar
                Values = Values & "updateflag = '" & sAction & "' "
                
                DoWrite ActionClause, WhereClause, "MMI_Address", Values, ParseChar
            End If
        End With
    Next oAddr
        
    GoTo CleanUpAndExit
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.Save_MMI_ADDRESS", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
End Sub




Private Sub Save_MMI_PHONE(ByRef oPatient As DTOs.Patient, ByVal sAction As String)
    Dim ActionClause As String
    Dim WhereClause As String
    Dim ParseChar As String
    Dim sSQL As String
    Dim Values As String
    Dim oPhone As DTOs.Phone
    Dim sSetString As String
    
    On Error GoTo ErrHandler
    
    ParseChar = Chr$(255)
    
    If oPatient.Phones.Count = 0 Then
        GoTo CleanUpAndExit
    End If
    
    sSetString = " where unit_number = " & oPatient.UnitNumber
    
    For Each oPhone In oPatient.Phones
        Values = vbNullString
        With oPhone
            '--- Delete
            sSQL = "delete from MMI_Phone " & sSetString & " and phone_type_cid = " & .PhoneType
            mobjDAS.ExecuteSQL sSQL
                
                
            '--- Insert
            sAction = "N"
            ActionClause = "INSERT INTO"
            
            If LenB(ActionClause) > 0 Then
                If LenB(Trim$(.PhoneType)) > 0 Then
                    Values = Values & "phone_type_cid = " & Trim$(.PhoneType) & "" & ParseChar
                End If
                If LenB(Trim$(.CountryCode)) > 0 Then
                    Values = Values & "country_code = '" & Trim$(.CountryCode) & "' " & ParseChar
                End If
                If LenB(Trim$(.AreaCode)) > 0 Then
                    Values = Values & "area_code = '" & Trim$(.AreaCode) & "' " & ParseChar
                End If
                If LenB(Trim$(.Phone)) > 0 Then
                    Values = Values & "local_number = '" & Trim$(.Phone) & "' " & ParseChar
                End If
                If LenB(Trim$(.Extension)) > 0 Then
                    Values = Values & "internal_extension = '" & Trim$(.Extension) & "' " & ParseChar
                End If
                   
                Values = Values & "CREATE_AUDIT_NUM = 576" & ParseChar
                Values = Values & "Unit_Number = " & oPatient.UnitNumber & ParseChar
                Values = Values & "updateflag = '" & sAction & "' "
                
                DoWrite ActionClause, WhereClause, "MMI_Phone", Values, ParseChar
            End If
        End With
      
    Next oPhone
    
    GoTo CleanUpAndExit
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.Save_MMI_PHONE", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
End Sub



Private Sub Save_MMI_GENERIC_ID(ByRef oPatient As DTOs.Patient, ByVal sAction As String)

    Dim ActionClause As String
    Dim WhereClause As String
    Dim ParseChar As String
    Dim sSQL As String
    Dim Values As String
    Dim oGenID As Generic_ID
    Dim oGenIDs As Generic_IDs
    Dim sSetString As String
    
    On Error GoTo ErrHandler
    
    LoadGenericIDs oPatient, oGenIDs
    
    If oGenIDs.Count = 0 Then
        GoTo CleanUpAndExit
    End If
    
    ParseChar = Chr$(255)
    sSetString = "  where unit_number = " & oPatient.UnitNumber
    
    For Each oGenID In oGenIDs
        Values = vbNullString
        With oGenID
            If sAction = "D" Then
                sSQL = "delete from mmi_generic_id " & sSetString _
                    & " and id_Type_ncid = " & CStr(.id_type_ncid)
                mobjDAS.ExecuteSQL sSQL
            Else
                sSQL = "Select count(*) num from mmi_generic_id Where Unit_Number = " & oPatient.UnitNumber _
                        & " and id_Type_ncid = " & CStr(.id_type_ncid)
                        
                If mobjDAS.OpenRecordset(sSQL).Fields("num") > 0 Then
                    sAction = "U"
                Else
                    sAction = "N"
                End If

            
                If sAction = "N" Then
                    ActionClause = "INSERT INTO"
                ElseIf sAction = "U" Then
                    ActionClause = "UPDATE"
                    WhereClause = "WHERE Unit_Number = " & oPatient.UnitNumber & " and id_Type_ncid = " & CStr(.id_type_ncid)
                End If
                
                If LenB(ActionClause) > 0 Then
                    If .assigning_authority > 0 Then
                        Values = Values & "assigning_authority_ncid = " & CStr(.assigning_authority) & "" & ParseChar
                    End If
                    If .id_type_ncid > 0 Then
                        Values = Values & "id_type_ncid = " & CStr(.id_type_ncid) & "" & ParseChar
                    End If
                    If LenB(Trim$(.id_value)) > 0 Then
                        Values = Values & "id_value = " & mConn.SQLQuote(.id_value) & ParseChar
                    End If
                    Values = Values & "CREATE_AUDIT_NUM = 576" & ParseChar
                    Values = Values & "Unit_Number = " & oPatient.UnitNumber & ParseChar
                    Values = Values & "updateflag = '" & sAction & "' "
                    
                    DoWrite ActionClause, WhereClause, "MMI_Generic_ID", Values, ParseChar
                End If
            End If
        End With
    Next oGenID
        
    GoTo CleanUpAndExit
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.Save_MMI_GENERIC_ID", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
    Set oGenID = Nothing
    Set oGenIDs = Nothing
End Sub



Private Sub LoadGenericIDs(ByRef oPatient As DTOs.Patient, ByRef oGenIDs As Generic_IDs)

    Dim oGenID As Generic_ID
    
    On Error GoTo ErrHandler
    
    If oGenIDs Is Nothing Then
        Set oGenIDs = New Generic_IDs
    End If
    
    With oPatient
        
        '--- FMP
        If LenB(.FMP) > 0 Then
            Set oGenID = Nothing
            Set oGenID = New Generic_ID
            oGenID.id_type_ncid = NCID_PATIENT_FMP
            oGenID.id_value = .FMP
            
            oGenIDs.Add oGenID
        End If
                    
        '--- FMP/SponsorSSN
        If LenB(.FMPSponsorSSN) > 0 Then
            Set oGenID = Nothing
            Set oGenID = New Generic_ID
            oGenID.id_type_ncid = NCID_FMP_SPONSOR_SSN
            oGenID.id_value = .FMPSponsorSSN
            
            oGenIDs.Add oGenID
        End If
        
        '--- SponsorSSN
        If LenB(.SponsorSSN) > 0 Then
            Set oGenID = Nothing
            Set oGenID = New Generic_ID
            oGenID.id_type_ncid = NCID_SPONSOR_SSN
            oGenID.id_value = .SponsorSSN
            
            oGenIDs.Add oGenID
        End If
        
        '---VIP
        Set oGenID = Nothing
        Set oGenID = New Generic_ID
        oGenID.id_type_ncid = NCID_VIP
        If .VIP = False Then
            oGenID.id_value = "N"
        Else
            oGenID.id_value = "Y"
        End If
        
        oGenIDs.Add oGenID
       
        '---PCM
         If LenB(.PCM_Code) > 0 Then
            Set oGenID = Nothing
            Set oGenID = New Generic_ID
            oGenID.id_type_ncid = NCID_PRIMARY_CARE_MANAGER
            oGenID.id_value = .PCM_Code
            
            oGenIDs.Add oGenID
        End If
        
        '---Sponsor Search String
          If LenB(.SponsorSrhString) > 0 Then
            Set oGenID = Nothing
            Set oGenID = New Generic_ID
            
            oGenID.id_type_ncid = NCID_FIRST_LETTER_LAST_NAME_LAST_FOUR_SPONSOR_SSN
            oGenID.id_value = .SponsorSrhString
            
            oGenIDs.Add oGenID
        End If
       
        '---Patient Search String
          If LenB(.PatientSrhString) > 0 Then
            Set oGenID = Nothing
            Set oGenID = New Generic_ID
            
            oGenID.id_type_ncid = NCID_FIRST_LETTER_LAST_NAME_LAST_FOUR_PATIENT_SSN
            oGenID.id_value = .PatientSrhString
            
            oGenIDs.Add oGenID
        End If
       
    End With
    
    GoTo CleanUpAndExit
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.Save_MMI_GENERIC_ID", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
End Sub





Private Sub Save_MMI_GENERIC_DATA(ByRef oPatient As DTOs.Patient, ByVal sAction As String)

    Dim ActionClause As String
    Dim WhereClause As String
    Dim ParseChar As String
    Dim sSQL As String
    Dim Values As String
    Dim oGenD As DTOs.PatientGenericData
    Dim sSetString As String
    Dim rs As ADODB.Recordset
    Dim oState As DTOs.iState
    Dim bFound As Boolean
    
    On Error GoTo ErrHandler
    
    If oPatient.GenericDatas.Count = 0 Then
        GoTo CleanUpAndExit
    End If
    
    ParseChar = Chr$(255)
    sSetString = " where unit_number = " & oPatient.UnitNumber
    
    For Each oGenD In oPatient.GenericDatas
        Values = vbNullString
        
        With oGenD
            
            Set oState = oGenD
            
            If oState.State = osDelete Or sAction = "D" Then
                sSQL = "delete from MMI_GenericData " & sSetString _
                    & " and ordinal= " & CStr(.Ordinal)
                mobjDAS.ExecuteSQL sSQL
            Else
                sSQL = "Select count(*) num, ordinal, Value_NCID from mmi_genericdata Where Unit_Number = " & oPatient.UnitNumber _
                                & " and Type_ncid = " & CStr(.TypeNcid) & " group by ordinal, Value_NCID"
                Set rs = mobjDAS.OpenRecordset(sSQL)
                                
                'mmi_genericdata contains possibly more than one command interest and
                '   command security
                'WR 5/1/2002 scr 18847
                If rs.RecordCount Then
                    If .TypeNcid = NCID_COMMAND_INTEREST Or _
                        .TypeNcid = NCID_COMMAND_SECURITY Then
                        
                        bFound = False  '--- SCR 82208   Sherry Wang   11/14/2005
                        Do While Not rs.EOF
                            If .ValueNcid = rs("Value_NCID").Value Then
                                .Ordinal = rs("ordinal").Value
                                bFound = True   '--- SCR 82208   Sherry Wang   11/14/2005
                                Exit Do
                            End If
                            rs.MoveNext
                        Loop
                        
                        '--- SCR 82208   Sherry Wang   11/14/2005
                        If bFound = False Then
                            sAction = "N"
                        End If
                    Else
                        If rs.Fields("num") > 0 Then
                            .Ordinal = rs("ordinal").Value
                            sAction = "U"
                        End If
                    End If
                End If
            
                If sAction = "N" Or rs.RecordCount = 0 Then
                    'need to set a default value for ordinal
                    sSQL = "Select max(ordinal) maximum from mmi_genericdata Where Unit_Number = " & oPatient.UnitNumber & " "
                    Set rs = mobjDAS.OpenRecordset(sSQL)

                    If rs.RecordCount = 0 Then
                        .Ordinal = 1
                    
                    'WR problem occurs on new patients downloaded
                    ElseIf IsNull(rs("maximum")) Then
                        .Ordinal = 1
                    Else
                        .Ordinal = rs("maximum").Value + 1
                    End If
                    
                    ActionClause = "INSERT INTO"
                ElseIf sAction = "U" Then
                    ActionClause = "UPDATE"
                    WhereClause = "WHERE unit_number = " & oPatient.UnitNumber & " and ordinal= " & .Ordinal
                End If
                
                If LenB(ActionClause) > 0 Then
                    If .Ordinal > 0 Then
                        Values = Values & "ordinal = " & CStr(.Ordinal) & "" & ParseChar
                    End If
                    If LenB(Trim$(.ErrorText)) > 0 Then
                        Values = Values & "error_text = " & mConn.SQLQuote(Trim$(.ErrorText)) & ParseChar
                    End If
                    If .Parent > 0 Then
                        Values = Values & "parent = " & CStr(.Parent) & "" & ParseChar
                    End If
                    If .TypeNcid > 0 Then
                        Values = Values & "type_ncid = " & CStr(.TypeNcid) & "" & ParseChar
                    End If
                    '*-- Coleman 112601: SCR 17043 - Facility NCID can be < 0...
                    If .ValueNcid <> 0 Then
                        Values = Values & "Value_NCID = " & CStr(.ValueNcid) & "" & ParseChar
                    End If
                    Values = Values & "Value_text = " & mConn.SQLQuote(Trim$(.ValueText)) & ParseChar
                    
                    If .ValueRsid > 0 Then
                        Values = Values & "Value_RSID = " & CStr(.ValueRsid) & "" & ParseChar
                    End If
                    '*-- Coleman 062101:
                    Values = Values & "CREATE_AUDIT_NUM = 576" & ParseChar
                    Values = Values & "Unit_Number = " & oPatient.UnitNumber & ParseChar
                    Values = Values & "updateflag = '" & sAction & "' "
                    
                    DoWrite ActionClause, WhereClause, "MMI_GenericData", Values, ParseChar
                End If
            End If
        End With
      
    Next oGenD
        
    GoTo CleanUpAndExit
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.Save_MMI_GENERIC_DATA", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
    Set oGenD = Nothing
End Sub

Private Sub Save_MMI_SITE(ByRef oPatient As DTOs.Patient)

    Dim ActionClause As String
    Dim WhereClause As String
    Dim ParseChar As String
    Dim sSQL As String
    Dim Values As String
    Dim sSetString As String
    Dim oSites As MMI_Sites
    Dim oSite As MMI_Site
    Dim lCnt As Long

    On Error GoTo ErrHandler

    ParseChar = Chr$(255)
    sSetString = " where unit_number = " & oPatient.UnitNumber


    If LenB(oPatient.FacilityNCID) = 0 And LenB(oPatient.MedicalRecordNumber) = 0 Then
        GoTo CleanUpAndExit
    End If
    
    '*-- Coleman 101901: Transfer of patients to the LDDB from
    '*-- CDR requires that we put in a SITE record, so make one
    '*-- up if none exist...Needed to create new appointments.
    If SaveSiteObj(oSites, oPatient.FacilityNCID, oPatient.UnitNumber, oPatient.MedicalRecordNumber) Then
    End If

    'SCR 25089 WR
    For lCnt = 1 To oSites.Count
        Values = vbNullString
        Set oSite = oSites.Item(lCnt)

        With oSite
            '--- do the delete
            sSQL = "delete from MMI_Site " & sSetString & " and Facility_cid=" & CLng(.FacilityNCID)
            mobjDAS.ExecuteSQL sSQL
                
            ActionClause = "INSERT INTO"

            If LenB(ActionClause) > 0 Then
                '*-- To Do: Patient.Facility is a text string, we need real CID for laptop...
                If LenB(Trim$(.FacilityNCID)) > 0 Then
                    Values = Values & "FACILITY_CID = " & _
                                CStr(.FacilityNCID) & ParseChar
                End If

                If LenB(Trim$(.MedRecNum)) > 0 Then
                    Values = Values & "MEDICAL_RECORD_NUMBER = " _
                            & mConn.SQLQuote(Trim$(.MedRecNum)) & ParseChar
                End If

                If LenB(Trim$(.RadiologyNum)) > 0 Then
                    Values = Values & "RADIOLOGY_NUMBER = " _
                               & mConn.SQLQuote(Trim$(.RadiologyNum)) & ParseChar
                End If

                '*-- Create_Audit_Num is arbitrary non-null value
                Values = Values & "CREATE_AUDIT_NUM = 576" & ParseChar
                '*-- Last values term added should not have ParseChar at end.
                Values = Values & "UNIT_NUMBER = " & CStr(oPatient.UnitNumber) & ""

                DoWrite ActionClause, WhereClause, "MMI_Site", Values, ParseChar
            End If
        End With
    Next lCnt


    GoTo CleanUpAndExit
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.Save_MMI_SITE", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
    Set oSite = Nothing
    Set oSites = Nothing
    
End Sub


Private Function SavePatient(ByRef oPatient As DTOs.Patient, ByVal sAction As String) As Boolean

    On Error GoTo ErrHandler
    
    mobjDAS.BeginTrans
    
    Save_MMI_ID oPatient, sAction
    Save_MMI_NAME oPatient
    Save_MMI_ADDRESS oPatient, sAction
    Save_MMI_PHONE oPatient, sAction
    Save_MMI_GENERIC_ID oPatient, sAction
    Save_MMI_GENERIC_DATA oPatient, sAction
    Save_MMI_SITE oPatient
    
            
    'Patient Profile
    SavePatientProfile oPatient

    mobjDAS.CommitTrans
    SavePatient = True
    
    GoTo CleanUpAndExit
ErrHandler:
    mobjDAS.RollbackTrans
    Err.Raise Err.Number, "MapperGeneric.SavePatient", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
End Function



Private Function GetOSH_INI_SETTINGS(ByVal sApkey As String) As String
    Dim sSQL As String
    Dim oRS As ADODB.Recordset
    
    On Error GoTo ErrHandler
    sSQL = "Select apvalue from osh_ini_settings where application_ncid = " & _
            "10001 and apsection = 'LOCAL_ENV' and apkey = '" & sApkey & "'"
            
    Set oRS = mConn.CHCSII_DAS(BackEndDB).OpenRecordset(sSQL, adOpenStatic, adLockReadOnly)
    
    If Not oRS Is Nothing Then
        If Not oRS.EOF Then
            If Not IsNull(oRS.Fields("apvalue")) Then
                GetOSH_INI_SETTINGS = oRS.Fields("apvalue")
            End If
        End If
        oRS.Close
        Set oRS = Nothing
    End If
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.GetOSH_INI_SETTINGS", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
    Set oRS = Nothing
End Function



'**********************************************************
' Load Extended Patient Object with the data in Recordset
'**********************************************************

Private Sub LoadGenericData(ByRef objPatient As DTOs.Patient, _
                           ByRef oRsGeneric As ADODB.Recordset)

Dim oGenericData As DTOs.PatientGenericData
Dim oRsPCM As ADODB.Recordset
Dim oPregnacy As DTOs.Pregnancy
Dim sTemp As String

    On Error GoTo ErrHandler
    
    With objPatient
        
        '--- Generic Data in both MMI_GENERIC_ID and MMI_GENERICDATA table
        If oRsGeneric.RecordCount > 0 Then
          oRsGeneric.MoveFirst
          
          Do While Not oRsGeneric.EOF
          
            If oRsGeneric("ORDINAL") > 0 Then
                
                Set oGenericData = Nothing
                Set oGenericData = New DTOs.PatientGenericData
                
                With oGenericData
                    .Ordinal = ReadLongField(oRsGeneric("ORDINAL"))
                    .TypeNcid = ReadLongField(oRsGeneric("TYPE_NCID"))
                    .ValueNcid = ReadLongField(oRsGeneric("VALUE_NCID"))
                    .ValueText = ReadStringField(oRsGeneric("VALUE_TEXT"))
                End With
                
                .GenericDatas.Add oGenericData, oGenericData.Ordinal
            End If
                
          
            Select Case oRsGeneric("TYPE_NCID")
              Case NCID_BLOOD_TYPE
                .BloodType = ReadStringField(oRsGeneric("VALUE_TEXT"))
              
              Case NCID_SPONSOR_SSN
                .SponsorSSN = ReadStringField(oRsGeneric("VALUE_TEXT"))
                .Formatted_SponsorSSN = FormatSSN(.SponsorSSN)
                
              Case NCID_FMP_SPONSOR_SSN
                .FMPSponsorSSN = ReadStringField(oRsGeneric("VALUE_TEXT"))
              
              Case NCID_PATIENT_FMP
                .FMP = ReadStringField(oRsGeneric("VALUE_TEXT"))
              
              Case NCID_VIP
                If IsNull(oRsGeneric("VALUE_TEXT")) Then
                    .VIP = False
                Else
                    If oRsGeneric("VALUE_TEXT") = "Y" Then
                        .VIP = True
                    Else
                        .VIP = False
                    End If
                End If
              
              Case NCID_RECORDS_MAINTAINED_AT
                .RecordsMaintainedAt = ReadStringField(oRsGeneric("VALUE_TEXT"))
              
              Case NCID_RANK
                .RankNCID = ReadStringField(oRsGeneric("VALUE_NCID"))
                .Rank = ReadStringField(oRsGeneric("VALUE_TEXT"))
                
                If LenB(oRsGeneric("VALUE_TEXT")) > 0 Then
                    .PayGrade = GetPayGradeFromRank(oRsGeneric("VALUE_TEXT"), .RankNCID)
                End If
              
              Case NCID_MEDICARE_ELIGIBILITY
                sTemp = ReadStringField(oRsGeneric("VALUE_NCID"))
                If LenB(sTemp) > 0 Then
                    .Medicare = mobjPicklist.ReturnRepFromNCID(NCID_COMMAND_INTEREST, sTemp)
                End If
                
              Case NCID_DEERS_ELIGIBILITY
                .DeersEligibility.Description = ReadStringField(oRsGeneric("VALUE_TEXT"))
              
              Case NCID_DEERS_ELIGIBILITY_START
                .DeersEligibility.StartDate = ReadStringField(oRsGeneric("VALUE_TEXT"))
              
              Case NCID_DEERS_ELIGIBILITY_END
                .DeersEligibility.EndDate = ReadStringField(oRsGeneric("VALUE_TEXT"))
                
              Case NCID_UNIT_IDENTIFICATION_CODE
                .UnitIdentificationCode = ReadStringField(oRsGeneric("VALUE_TEXT"))
                
              Case NCID_ACV_ENROLLMENT_STATUS
                .ACV_EnrollmentStatus = ReadStringField(oRsGeneric("VALUE_TEXT"))
              
              Case NCID_PATIENT_CATEGORY
                .PatientCategory = ReadStringField(oRsGeneric("VALUE_TEXT"))
                
                '--- SCR 134102  Sherry Wang  7/3/2007
                If LenB(oRsGeneric("VALUE_TEXT")) > 0 And LenB(.BranchOfService) = 0 Then
                    .BranchOfService = GetBranchOfService(.PatientCategory)
                End If
                
              Case NCID_SERVICE
                If LenB(oRsGeneric("VALUE_TEXT")) > 0 Then
                    .BranchOfService = ReadStringField(oRsGeneric("VALUE_TEXT"))
                End If
                
              Case NCID_FLYING_STATUS
                .SpecialWorkStatus.FlyingStatus = ReadStringField(oRsGeneric("VALUE_TEXT"))
              
              Case NCID_JUMPING_STATUS
                .SpecialWorkStatus.JumpingStatus = ReadStringField(oRsGeneric("VALUE_TEXT"))
              
              Case NCID_DIVING_STATUS
                .SpecialWorkStatus.DivingStatus = ReadStringField(oRsGeneric("VALUE_TEXT"))
              
              Case NCID_ON_MOBILITY
                .SpecialWorkStatus.OnMobility = ReadStringField(oRsGeneric("VALUE_TEXT"))
                
              Case NCID_MILITARY_POLICE
                .SpecialWorkStatus.MilitaryPolice = ReadStringField(oRsGeneric("VALUE_TEXT"))
              
              Case NCID_SUBMARINE
                .SpecialWorkStatus.Submarine = ReadStringField(oRsGeneric("VALUE_TEXT"))
              
              Case NCID_EMAIL
                .EmailAddress = ReadStringField(oRsGeneric("VALUE_TEXT"))
              
              Case NCID_PREGNANCY
                If Not IsNull(oRsGeneric("VALUE_TEXT")) Then
                    If LoadPregnancyObject(oRsGeneric("VALUE_TEXT"), oPregnacy) Then
                        Set .Pregnancy = oPregnacy
                    End If
                End If
              Case NCID_COMMAND_INTEREST
                If Not IsNull(oRsGeneric("VALUE_NCID")) Then
                    If LenB(oRsGeneric("VALUE_NCID")) > 0 Or LenB(oRsGeneric("VALUE_TEXT")) > 0 Then
                        
                        Set oGenericData = Nothing
                        Set oGenericData = New DTOs.PatientGenericData
                        
                        oGenericData.TypeNcid = NCID_COMMAND_INTEREST
                        oGenericData.ValueNcid = oRsGeneric("VALUE_NCID")
                        If IsNull(oRsGeneric("VALUE_TEXT")) Then
                            oGenericData.ValueText = mobjPicklist.ReturnRepFromNCID(NCID_COMMAND_INTEREST, oRsGeneric("VALUE_NCID"))
                        Else
                            oGenericData.ValueText = oRsGeneric("VALUE_TEXT")
                        End If
                        
                        .CommandInterest.Add oGenericData, oGenericData.ValueNcid
                    End If
                End If
              Case NCID_COMMAND_SECURITY
                If Not IsNull(oRsGeneric("VALUE_NCID")) Then
                    If LenB(oRsGeneric("VALUE_NCID")) > 0 Or LenB(oRsGeneric("VALUE_TEXT")) > 0 Then
                        Set oGenericData = Nothing
                        Set oGenericData = New DTOs.PatientGenericData
                        
                        oGenericData.TypeNcid = NCID_COMMAND_SECURITY
                        oGenericData.ValueNcid = oRsGeneric("VALUE_NCID")
                        If IsNull(oRsGeneric("VALUE_TEXT")) Then
                            oGenericData.ValueText = mobjPicklist.ReturnRepFromNCID(NCID_COMMAND_SECURITY, oRsGeneric("VALUE_NCID"))
                        Else
                            oGenericData.ValueText = oRsGeneric("VALUE_TEXT")
                        End If
                        
                        .CommandSecurity.Add oGenericData, oGenericData.ValueNcid
                    End If
                End If
              Case NCID_ENROLLMENT_FACILITY
                If Not IsNull(oRsGeneric("VALUE_NCID")) Then
                    If LenB(oRsGeneric("VALUE_NCID")) > 0 Then
                        .Facility = mobjPicklist.ReturnRepFromNCID(NCID_ENROLLMENT_FACILITY, oRsGeneric("VALUE_NCID"))
                        If LenB(.Facility) = 0 Then
                            .Facility = mobjPicklist.ReturnRepFromNCID(NCID_FACILITY_DOMAIN, oRsGeneric("VALUE_NCID"))
                        End If
                        
                    End If
                End If
                
                If Not IsNull(oRsGeneric("VALUE_TEXT")) Then
                    .EnrollmentFacility = oRsGeneric("VALUE_TEXT")
                End If
              Case NCID_PRIMARY_CARE_MANAGER
                .PCM_Code = oRsGeneric("VALUE_TEXT")
                Set oRsPCM = Get_PCM(oRsGeneric("VALUE_TEXT"))
                If oRsPCM.RecordCount > 0 Then
                    .PCM_IEN = ReadStringField(oRsPCM("IEN"))
                    .PrimaryCareManager = ReadStringField(oRsPCM("NAME"))
                End If
             Case NCID_COMMENTS
                .Comments = ReadStringField(oRsGeneric("VALUE_TEXT"))
            End Select
        
            oRsGeneric.MoveNext
        Loop
        
      End If
        
    End With
    
    GoTo CleanUpAndExit
    
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.LoadGenericData", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
    Set oRsPCM = Nothing
End Sub



'**********************************************************
' Load Extended Patient Object with the Address data in Recordset
'**********************************************************

Private Sub LoadAddress(ByRef objPatient As DTOs.Patient, _
                           ByRef oRsAddress As ADODB.Recordset)

Dim oAddress As DTOs.Address

    On Error GoTo ErrHandler
    
    With objPatient
        
        If oRsAddress.RecordCount > 0 Then
            oRsAddress.MoveFirst
            
            Do While Not oRsAddress.EOF
                Set oAddress = Nothing
                Set oAddress = New DTOs.Address
                With oAddress
                    .Street1 = ReadStringField(oRsAddress("STREET_1"))
                    .Street2 = ReadStringField(oRsAddress("STREET_2"))
                    .City = ReadStringField(oRsAddress("CITY"))
                    .State = ReadStringField(oRsAddress("STATE"))
                    .Zip = ReadStringField(oRsAddress("POST_CODE"))
                    .County = ReadStringField(oRsAddress("COUNTY"))
                    .Country = ReadStringField(oRsAddress("COUNTRY"))
                    .AddrType = ReadStringField(oRsAddress("ADDRESS_TYPE_CID"))
                End With
                
                .Addresses.Add oAddress, oAddress.AddrType
                
                If oAddress.AddrType = NCID_HOME_ADDRESS Then
                    Set .HomeAddress = oAddress
                End If
                
                oRsAddress.MoveNext
            Loop
        End If
        
    End With
    
    GoTo CleanUpAndExit
    
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.LoadAddress", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
    Set oAddress = Nothing
End Sub




'**********************************************************
' Load Extended Patient Object with the Phone data in Recordset
'**********************************************************

Private Sub LoadPhones(ByRef objPatient As DTOs.Patient, _
                           ByRef oRsPhone As ADODB.Recordset)


Dim oPhone As DTOs.Phone

    On Error GoTo ErrHandler
    
    With objPatient
        
        If oRsPhone.RecordCount > 0 Then
            oRsPhone.MoveFirst
            
            Do While Not oRsPhone.EOF
                Set oPhone = Nothing
                Set oPhone = New DTOs.Phone
                With oPhone
                    
                    .AreaCode = ReadStringField(oRsPhone("AREA_CODE"))
                    .Phone = ReadStringField(oRsPhone("LOCAL_NUMBER"))
                    .CountryCode = ReadStringField(oRsPhone("COUNTRY_CODE"))
                    .Extension = ReadStringField(oRsPhone("INTERNAL_EXTENSION"))
                    .PhoneType = ReadStringField(oRsPhone("PHONE_TYPE_CID"))
                    
                End With
                
                If oPhone.PhoneType = NCID_HOME_PHONE Then
                    Set .HomePhone = oPhone
                ElseIf oPhone.PhoneType = NCID_WORK_PHONE Then
                    Set .WorkPhone = oPhone
                End If
                
                .Phones.Add oPhone, oPhone.PhoneType
                
                oRsPhone.MoveNext
            Loop
        End If
        
    End With
    
    GoTo CleanUpAndExit
    
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.LoadPhones", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
    Set oPhone = Nothing
End Sub



Private Sub UpdatePatientObject(ByRef oPatient As DTOs.Patient)

Dim sService As String

    On Error GoTo ErrHandler
    
    sService = GetBranchOfService(oPatient.PatientCategory)
    
    '--- SCR 52842: update the service with the patient category info.
    If LenB(oPatient.PatientCategory) > 0 Then
        oPatient.BranchOfService = sService
    End If
    
    If oPatient.UnitNumber = oPatient.Sponsor.UnitNumber Then
        oPatient.Sponsor.BranchOfService = sService
    End If
    
    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.UpdatePatient", Err.Description, Err.HelpFile, Err.HelpContext
End Sub

'--- SCR 48476
Private Function GetSearchSSN(ByVal sSSN As String) As String
Dim sTempSSN As String

    On Error GoTo ErrHandler
    
    GetSearchSSN = sSSN
    sTempSSN = Replace(Trim$(sSSN), "-", vbNullString)
    If Len(sTempSSN) = 9 And Len(ReturnNumbers(sTempSSN)) = 9 Then
        GetSearchSSN = sTempSSN
    End If
    
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.GetSearchSSN", Err.Description, Err.HelpFile, Err.HelpContext
End Function

Private Function GetSQLPatientSearchFromADS(ByRef Params As PatientSearchParam) As String
Dim sSQL As String
Dim sSelect As String
Dim sFrom As String
Dim sWhere As String
Dim sTemp As String
Dim bParamExists As Boolean
Dim oShared As CWShared

    On Error GoTo ErrHandler
    
    bParamExists = False
    
    '--- Select
    If BackEndDB = MYSQL Then
        sSelect = "SELECT DATE_FORMAT(ID.DOB, '%m/%d/%Y') BIRTH_DATE, " & vbNewLine
    Else
        sSelect = "SELECT TO_CHAR(ID.DOB, 'MM/DD/YYYY') BIRTH_DATE, " & vbNewLine
    End If
    
    '-- SCR 190220 B. Webb DEERS Id Fix
    sSelect = sSelect & " ID.SSN, ID.GENDER AS SEX, ID.LAST_NAME, ID.EDIPIN AS DEERSID, 0 AS UNIT_NUMBER, " & vbNewLine
    sSelect = sSelect & " ID.FIRST_NAME, ID.MIDDLE_NAME, ID.FMP, ID.SPONSOR_SSN, " & vbNewLine
    sSelect = sSelect & " ID.MARITAL_STATUS_NCID, ID.RACE_NCID," & vbNewLine
    sSelect = sSelect & " ID.RELIGION_NCID, ID.RANK_NCID," & vbNewLine
    sSelect = sSelect & " ID.RANK_TEXT, ID.UIC," & vbNewLine
    sSelect = sSelect & " ID.PATCAT_NCID, ID.PATCAT_TEXT," & vbNewLine
    sSelect = sSelect & " ID.SERVICE , ID.SPECIALTY_CODE" & vbNewLine
    
    '--- From
    If BackEndDB = MYSQL Then
        sFrom = sFrom & " FROM GEMSA.MMI_ADS ID " & vbNewLine
    Else
        sFrom = sFrom & " FROM MMI_ADS ID " & vbNewLine
    End If
    
    If Params.GetAllPatients Then
        GetSQLPatientSearchFromADS = sSelect & sFrom & sWhere
        Exit Function
    End If
    sWhere = " Where "
    '--- Last Name
    If LenB(Params.LastName) > 0 Then
        sWhere = sWhere & "ID.LAST_NAME LIKE '" & Replace(Params.LastName, "'", "''") & "%'" & vbNewLine
        bParamExists = True
    End If
    
    '--- First Name
    If LenB(Params.FirstName) > 0 Then
        '--- SCR 88926   Sherry Wang  3/7/2006
        If bParamExists = True Then
            sWhere = sWhere & " AND  ID.FIRST_NAME LIKE '" & Replace(Params.FirstName, "'", "''") & "%'" & vbNewLine
        Else
             sWhere = sWhere & "ID.FIRST_NAME LIKE '" & Replace(Params.FirstName, "'", "''") & "%'" & vbNewLine
        End If
        bParamExists = True
    End If
    
    '--- Patient SSN
    If LenB(Params.SSN) > 0 Then
        If bParamExists = True Then
            sWhere = sWhere & " AND  ID.SSN = '" & GetSearchSSN(Params.SSN) & "'" & vbNewLine
        Else
            sWhere = sWhere & " ID.SSN = '" & GetSearchSSN(Params.SSN) & "'" & vbNewLine
        End If
        bParamExists = True
    End If
       
    '--- Patient DOB
    If LenB(Params.DOB) > 0 Then
        If bParamExists = True Then
            sWhere = sWhere & " AND  ID.DOB = " & mConn.SQLDate(Params.DOB) & vbNewLine
        Else
            sWhere = sWhere & " ID.DOB = " & mConn.SQLDate(Params.DOB) & vbNewLine
        End If
        bParamExists = True
    End If
    
    '--- Patient Gender
    If LenB(Params.SEX) > 0 Then
        If bParamExists = True Then
            sWhere = sWhere & " AND ID.GENDER = '" & UCase$(Left$(Params.SEX, 1)) & "'" & vbNewLine
        Else
            sWhere = sWhere & " ID.GENDER = '" & UCase$(Left$(Params.SEX, 1)) & "'" & vbNewLine
        End If
        bParamExists = True
    End If
    
    '--- Max Number Of Records
    If LenB(Params.MaximumNumberOfRecords) > 0 Then
        oShared = New CWShared
        If bParamExists = True Then
            If oShared.UseMySQLDB Then
                sWhere = sWhere & " LIMIT '" & Params.MaximumNumberOfRecords + 1 & "'" & vbNewLine
            Else
                sWhere = sWhere & " AND  rownum <= '" & Params.MaximumNumberOfRecords + 1 & "'" & vbNewLine
            End If
        Else
            If oShared.UseMySQLDB Then
                sWhere = sWhere & " LIMIT '" & Params.MaximumNumberOfRecords + 1 & "'" & vbNewLine
            Else
                sWhere = sWhere & " rownum <= '" & Params.MaximumNumberOfRecords + 1 & "'" & vbNewLine
            End If
        End If
        bParamExists = True
    End If
    
    '--- If no input parameters exists, return NULL string
    If bParamExists = False Then
        sSQL = vbNullString
    Else
        sSQL = sSelect & sFrom & sWhere
    End If
    
    GetSQLPatientSearchFromADS = sSQL
    
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.GetSQLPatientSearch", Err.Description, Err.HelpFile, Err.HelpContext
    
End Function

'--- SCR 84903   Jane Shen   01/05/2006
Private Function LoadADSPatientObject(ByRef objPatient As DTOs.Patient, ByRef oRS As ADODB.Recordset) As Boolean
    On Error GoTo ErrHandler
    
    With objPatient
        .ADSFlag = True
        .Name.First_Name = ReadStringField(oRS("FIRST_NAME"))
        .Name.Last_Name = ReadStringField(oRS("LAST_NAME"))
        .Name.Middle_Name = ReadStringField(oRS("MIDDLE_NAME"))
        .SEX = ReadStringField(oRS("SEX"))
        .SSN = ReadStringField(oRS("SSN"))
        .FMP = ReadStringField(oRS("FMP"))
        .SponsorSSN = ReadStringField(oRS("SPONSOR_SSN"))
        .Formatted_SSN = FormatSSN(.SSN)
        .FMPSponsorSSN = ReadStringField(oRS("FMP")) & "/" & (ReadStringField(oRS("SPONSOR_SSN")))
        
        If Not IsNull(oRS("BIRTH_DATE")) Then
            If oRS("BIRTH_DATE") <> NULL_DATE And LenB(oRS("BIRTH_DATE")) > 0 Then
                .DateOfBirth = CDate(oRS("BIRTH_DATE"))
            End If
        End If
        
        .UnitIdentificationCode = ReadStringField(oRS("UIC"))
        .Race = ReadStringField(oRS("RACE_NCID"))
        .Religion = ReadStringField(oRS("RELIGION_NCID"))
        .MaritalStatus = ReadStringField(oRS("MARITAL_STATUS_NCID"))
        .RankNCID = ReadStringField(oRS("RANK_NCID"))
        .Rank = ReadStringField(oRS("RANK_TEXT"))
        .PatientCategory = ReadStringField(oRS("PATCAT_TEXT"))
        .BranchOfService = ReadStringField(oRS("SERVICE"))
        '-- SCR 190220 B. Webb DEERS Id Fix
        .deersID = GetDeersID(ReadStringField(oRS("DEERSID")))
    End With
    
    LoadADSPatientObject = True
    
    GoTo CleanUpAndExit
    
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.LoadADSPatientObject", Err.Description, Err.HelpFile, Err.HelpContext
CleanUpAndExit:
End Function


Private Function GetVerifySourceInfo() As ADODB.Recordset

Dim oRS As ADODB.Recordset
Dim sSQL As String
    
    On Error GoTo ErrHandler
    
    If BackEndDB = MYSQL Then
        sSQL = "SELECT DATA_SOURCE, concat('Data extracted on ', DATE_FORMAT(EXTRACT_DATE, '%m/%d/%Y')) AS VERIFY_VERSION"
    Else
        sSQL = "SELECT DATA_SOURCE, 'Data extracted on ' || TO_CHAR(EXTRACT_DATE, 'MM/DD/YYYY') AS VERIFY_VERSION"
    End If
    
    sSQL = sSQL & " From ADS_IMPORT_HISTORY"
    sSQL = sSQL & " ORDER BY IMPORT_DATE DESC"  '--- SCR 89205 Sherry Wang  3/3/2006
    
    Set oRS = mConn.CHCSII_DAS(BackEndDB).OpenRecordset(sSQL)
    
    If oRS.RecordCount > 0 Then
      Set GetVerifySourceInfo = oRS
    Else
        Set GetVerifySourceInfo = Nothing
    End If
    
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "MapperGeneric.GetVerifySourceInfo", Err.Description, Err.HelpFile, Err.HelpContext
End Function

'-- SCR 190220 B. Webb Deers Id Fix
Private Function GetDeersID(ByVal deersID As String) As String
    Dim id As String

    If LenB(deersID) > 1 Then
        If Left$(deersID, 1) = "D" Then
            id = Mid(deersID, 2)
        End If
    End If

    GetDeersID = id
End Function
