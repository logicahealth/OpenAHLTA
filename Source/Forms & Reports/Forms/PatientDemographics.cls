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
Attribute VB_Name = "PatientDemographics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const sMODULE_NAME As String = "PatientDemographic.cls"

'-- Patient Demographic Category NCIDs
'-----------------------------------------------------------------

'-- MMI_GENERIC_DATA
Private Const mcCOMMAND_INTEREST As String = "204685"
Private Const mcCOMMAND_SECURITY As String = "204687"
Private Const mcALTERNATIVE_CARE As String = "204696"
Private Const mcFLYING_STATUS As String = "217201"
Private Const mcJUMPING_STATUS As String = "217202"
Private Const mcDIVING_STATUS As String = "217203"
Private Const mcON_MOBILITY As String = "217204"
Private Const mcSUBMARINE As String = "217205"
Private Const mcPATIENT_CATEGORY As String = "14501721"
Private Const mcUIC As String = "14501722"
Private Const mcRECORDS_MAINTAINED_AT = "14504166"
Private Const mcRANK = "1015"
Private Const mcENROLLMENT_FACILITY = "14501723"

'-- MMI_GENERIC_ID
Private Const mcSPONSOR_SSN As String = "186334"
Private Const msFMP As String = "204708"
Private Const msVIP As String = "14501719"
Private Const msPCM As String = "14501724"

Public Enum ReturnType
    None = 0
    NCID = 1
    OriginalText = 2
    HDDText = 3
End Enum
Public Function CHCS_Get_Demographics(sUnitNumber As String, oData As ADODB.Recordset) As Boolean
    
    Dim oCHCSInterface As ICHCSPatient
    Dim oPatientInfo As ADODB.Recordset
    
    Dim sTemp As String
    Dim sStmt As String
    Dim sHomePhone As String
    Dim sWorkPhone As String
    Dim sFirstName As String
    Dim sMiddleName As String
    Dim sLastName As String
    Dim cGenericData As Collection
    Dim cGenericID As Collection
    
    On Error GoTo ErrHandler
    CHCS_Get_Demographics = False
    
    Set oCHCSInterface = gobjshared.CHCSConnection
    Set oPatientInfo = oCHCSInterface.RetrieveDemographics(sUnitNumber)
    
    Set oData = Nothing
    GetTimeZoneInformation TZInfo

    '-- Initialize Demographics Record
    '-------------------------------------------------------
    Set oData = New ADODB.Recordset
    
    Init_Patient_Demographics_Rcd oData
    
    '-- Fill Record    '-------------------------------------------------------
    oData.AddNew
    oData.Fields.Item("job_id") = 1
    If sUnitNumber = "" Then GoTo Skip
    With oData.Fields
        .Item("unit_number").Value = sUnitNumber
'        .Item("First_Name").Value = sFirstName
'        .Item("Middle_Name").Value = sMiddleName
'        .Item("Last_Name").Value = sLastName
        .Item("full_Name").Value = oPatientInfo("NAME") & vbNullString
        
        .Item("SSN").Value = oPatientInfo("PATIENT_SSN") & vbNullString
        If Len(oPatientInfo("SEX")) > 0 Then
            .Item("Sex").Value = left(oPatientInfo("SEX"), 1) & vbNullString
        End If
        
        .Item("DOB").Value = oPatientInfo("DOB") & vbNullString
        .Item("dob_formatted").Value = Format(.Item("DOB").Value, gsDefaultFormat) & vbNullString
        .Item("AGE").Value = gobjshared.CalcAge(CDbl(CDate(Format(.Item("DOB").Value, "Short Date"))))
        .Item("Marital_Status").Value = oPatientInfo("MARITAL_STATUS") & vbNullString
    
        '-- Command Security can be multiple values
        On Error Resume Next 'Fields may not be in collection
        '.Item("Command_Interest").Value = OPATIENTINFO("")
        .Item("Command_Security").Value = oPatientInfo("COMMAND_SECURITY") & vbNullString
        .Item("Alternative_Care").Value = oPatientInfo("ACV_CODE") & vbNullString
        
        sTemp = ""
        sTemp = oPatientInfo("MILITARY_FLYING_STATUS")
        If left(sTemp, 1) = "Y" Then
            .Item("Flying_Status").Value = 1
        Else
            .Item("Flying_Status").Value = 0
        End If
        
        sTemp = ""
        'sTemp = oPatientInfo("mcJUMPING_STATUS")
        If Len(sTemp) >= 1 Then
           .Item("Jumping_Status").Value = left(sTemp, 1) = "Y"
        Else
           .Item("Jumping_Status").Value = False
        End If
        
        sTemp = ""
        'sTemp = oPatientInfo("mcDIVING_STATUS")
        If Len(sTemp) >= 1 Then
           .Item("Diving_Status").Value = left(sTemp, 1) = "Y"
        Else
           .Item("Diving_Status").Value = False
        End If
        
        sTemp = ""
        'sTemp = oPatientInfo("mcON_MOBILITY")
        If Len(sTemp) >= 1 Then
           .Item("On_Mobility_Status").Value = left(sTemp, 1) = "Y"
        Else
           .Item("On_Mobility_Status").Value = False
        End If
        
        .Item("Patient_Category").Value = oPatientInfo("PATIENT_CATEGORY") & vbNullString
    
        '-- Militiary Police
        .Item("UIC").Value = oPatientInfo("DUTY UNIT") & vbNullString
        .Item("Sponsor_SSN").Value = oPatientInfo("SPONSOR_SSN") & vbNullString
        .Item("SPONSOR_UNIT_NUMBER").Value = oPatientInfo("SPONSOR_IEN") & vbNullString
        .Item("FMP").Value = oPatientInfo("FMP") & vbNullString
        '.Item("VIP").Value = cGenericID(msVIP) & vbNullString
            
        .Item("PCM_Id").Value = oPatientInfo("NEDS_PCM_NAME") & vbNullString 'PCM
        .Item("Home_Phone").Value = oPatientInfo("PHONE") & vbNullString
        .Item("Work_Phone").Value = oPatientInfo("OFFICE_PHONE") & vbNullString
        .Item("Service").Value = GetService(.Item("Patient_Category").Value & vbNullString)
        .Item("Rank").Value = oPatientInfo("RANK") & vbNullString
        .Item("Records_Maintained_At").Value = oPatientInfo("OUTPATIENT_RECORD_LOCATION") & vbNullString
        If Not IsNull(.Item("FMP").Value) Then
            .Item("Relationship_To_Sponsor").Value = oPatientInfo("FMP") & vbNullString
        End If
        .Item("INSURANCE").Value = oPatientInfo("INSURANCE") & vbNullString
        .Item("TELPCM").Value = oPatientInfo("PCM_PHONE") & vbNullString
        '.Item("Active_Duty").Value = IsActiveDuty(sUnitNumber)
        
        '-- Enrollment Facility  SCR 25810 Sherry Wang 7/19/2002
        '.Item("Enrollment_Facility").Value = GetEnrollmentFacility(sUnitNumber, HDDText)
    End With

Skip:
    oData.Update

    CHCS_Get_Demographics = True
    GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".CHCS_Get_Demographics", App.EXEName, vbExclamation
    'Resume Next
Cleanup:
    Set cGenericData = Nothing
    Set cGenericID = Nothing
    Set oCHCSInterface = Nothing
    Set oPatientInfo = Nothing
End Function

'
Public Function RecordGetDemographics(sUnitNumber As String, oData As ADODB.Recordset) As Boolean

    On Error GoTo ErrHandler
    If gobjshared.IsAppMode(modeCHCSI_GUI) Then
       RecordGetDemographics = CHCS_Get_Demographics(sUnitNumber, oData)
       Exit Function
    End If
    
    RecordGetDemographics = False
    
    Dim sTemp As String
    Dim sStmt As String
    Dim sHomePhone As String
    Dim sWorkPhone As String
    Dim sFirstName As String
    Dim sMiddleName As String
    Dim sLastName As String
    Dim cGenericData As Collection
    Dim cGenericID As Collection
    Dim rsMMI_ID As ADODB.Recordset
    
    GetTimeZoneInformation TZInfo

    '-- Initialize Demographics Record
    '-------------------------------------------------------
    Set oData = Nothing
    Set oData = New ADODB.Recordset
    If Not Init_Patient_Demographics_Rcd(oData) Then GoTo Cleanup

    '-- Fill Record    '-------------------------------------------------------
    oData.AddNew
    oData.Fields.Item("job_id") = 1
    If sUnitNumber = "" Then GoTo Skip
    If Not GetPatientName(sUnitNumber, sFirstName, sMiddleName, sLastName) Then GoTo Cleanup
    Get_MMI_ID_Information sUnitNumber, rsMMI_ID
    
    Set cGenericID = GetGenericIDCollection(sUnitNumber)
    Set cGenericData = GetGenericDataCollection(sUnitNumber)
    
    If Not rsMMI_ID Is Nothing Then rsMMI_ID.MoveFirst
    
    With oData.Fields
        .Item("unit_number").Value = sUnitNumber
        .Item("First_Name").Value = sFirstName
        .Item("Middle_Name").Value = sMiddleName
        .Item("Last_Name").Value = sLastName
        .Item("full_Name").Value = Trim(sLastName & ", " & sFirstName & " " & sMiddleName)
        
        If Not rsMMI_ID Is Nothing Then
            .Item("SSN").Value = rsMMI_ID("SSN")
            .Item("Sex").Value = rsMMI_ID("SEX")
            If Not rsMMI_ID("BIRTH_DATE") = NULLStr Then .Item("DOB").Value = rsMMI_ID("BIRTH_DATE")
            .Item("dob_formatted").Value = Format(.Item("DOB").Value, gsDefaultFormat)
        End If
            
        If Not IsNull(.Item("DOB").Value) Then
           .Item("AGE").Value = Val(gobjshared.CalcAge(CDbl(CDate(Format(.Item("DOB").Value, "Short Date")))))
        End If
        
        .Item("Marital_Status").Value = GetMaritalStatus(sUnitNumber)
    
        '-- Command Security can be multiple values
        On Error Resume Next 'Fields may not be in collection
        .Item("Command_Interest").Value = cGenericData(mcCOMMAND_INTEREST)
        .Item("Command_Security").Value = cGenericData(mcCOMMAND_SECURITY)
        .Item("Alternative_Care").Value = cGenericData(mcALTERNATIVE_CARE)
        
        sTemp = ""
        sTemp = cGenericData(mcFLYING_STATUS)
        If sTemp = "Y" Then
            .Item("Flying_Status").Value = 1
        Else
            .Item("Flying_Status").Value = 0
        End If
        
        sTemp = ""
        sTemp = cGenericData(mcJUMPING_STATUS)
        If Len(sTemp) >= 1 Then
           .Item("Jumping_Status").Value = left(sTemp, 1) = "Y"
        Else
           .Item("Jumping_Status").Value = False
        End If
        
        sTemp = ""
        sTemp = cGenericData(mcDIVING_STATUS)
        If Len(sTemp) >= 1 Then
           .Item("Diving_Status").Value = left(sTemp, 1) = "Y"
        Else
           .Item("Diving_Status").Value = False
        End If
        
        sTemp = ""
        sTemp = cGenericData(mcON_MOBILITY)
        If Len(sTemp) >= 1 Then
           .Item("On_Mobility_Status").Value = left(sTemp, 1) = "Y"
        Else
           .Item("On_Mobility_Status").Value = False
        End If
        
        .Item("Patient_Category").Value = cGenericData(mcPATIENT_CATEGORY)
    
        '-- Militiary Police
        .Item("UIC").Value = cGenericData(mcUIC)
        .Item("Sponsor_SSN").Value = cGenericID(mcSPONSOR_SSN)
        .Item("FMP").Value = cGenericID(msFMP)
            
        .Item("PCM_Id").Value = cGenericID(msPCM)
        Call GetHomeAndWorkPhone(sUnitNumber, sWorkPhone, sHomePhone)
        .Item("Home_Phone").Value = sHomePhone
        .Item("Work_Phone").Value = sWorkPhone
        .Item("Service").Value = GetService(.Item("Patient_Category").Value & "")
        .Item("Rank").Value = cGenericData(mcRANK)
        .Item("Records_Maintained_At").Value = GetRecordsMaintainedAt(sUnitNumber)
        .Item("Active_Duty").Value = IsActiveDuty(sUnitNumber)
        
        '-- Enrollment Facility  SCR 25810 Sherry Wang 7/19/2002
        .Item("Enrollment_Facility").Value = GetEnrollmentFacility(sUnitNumber, HDDText)
    End With

Skip:
    oData.Update

    RecordGetDemographics = True
    GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".RecordGetDemographics", App.EXEName, vbExclamation
Cleanup:
    Set rsMMI_ID = Nothing
    Set cGenericData = Nothing
    Set cGenericID = Nothing
End Function

Public Function RecordGetAddress(sUnitNumber As String, oData As ADODB.Recordset) As Boolean
On Error GoTo ErrHandler
    RecordGetAddress = False
    Dim sStmt As String
    Set oData = Nothing
    
    sStmt = "select " & _
            "1 as job_id, unit_number, street_1, street_2, city, state, post_code, country " & _
        "from " & _
            "mmi_address " & _
        "where " & _
            "unit_number = " & sUnitNumber
           
    If gobjshared.UseSQLServer Then
        Set oData = goDBDAS.OpenRecordset(sStmt, adOpenDynamic, adLockOptimistic)
    Else
        Set oData = goDBDAS.OpenRecordset(sStmt)
    End If
    Set oData.ActiveConnection = Nothing
    
    
    If Not ReplaceNulls(oData) Then GoTo Cleanup
    

RecordGetAddress = True
GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".RecordGetAddress", App.EXEName, vbExclamation
Cleanup:
End Function


Public Function IsActiveDuty(sUnitNumber As String) As String

    Dim sStmt As String
    
    On Error GoTo ErrHandler
    IsActiveDuty = "N"

    sStmt = "select Trim(value_text) UIC " & _
            "from mmi_genericdata " & _
            "where (value_ncid in " & _
                "(99990864, 99990865, 99990866, 99990870, 99990871, 99990935, 99990939, 99991114, 99991117, 99991139, 99991142, 99991143) and " & _
                "(type_ncid = 14501721) and unit_number = " & sUnitNumber & ") "
    
    goDBSqlOps.Execute sStmt
    
    If Not (goDBSqlOps.BOF And goDBSqlOps.EOF) Then
        IsActiveDuty = "Y"
    End If

    Exit Function

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".IsActiveDuty", App.EXEName, vbExclamation
End Function

Function GetSSN(sUnitNumber As String) As String

On Error GoTo ErrHandler
GetSSN = ""

Dim sStmt  As String

sStmt = "select " & _
            "SSN " & _
        "from " & _
            "mmi_id id " & _
        "where " & _
            "id.unit_number = " & sUnitNumber

goDBSqlOps.Execute sStmt

If goDBSqlOps.EOF = True Then
    GoTo Cleanup
End If

GetSSN = goDBSqlOps.Value("SSN")
GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".GetSSN", App.EXEName, vbExclamation
Cleanup:
End Function

Private Function Get_MMI_ID_Information(sUnitNumber As String, rsRcd As ADODB.Recordset) As Boolean
'hwcc 1/27/2004 - Collection does not work well and splitting out VB error in RecordGetDemographics
' I am changing it to use RecordSet Instead

    Get_MMI_ID_Information = False
    On Error GoTo ErrHandler

    Dim sStmt As String
    sStmt = "select * " & _
            "from mmi_id id " & _
            "where id.unit_number = " & sUnitNumber

    If gobjshared.UseSQLServer Then
        Set rsRcd = goDBDAS.OpenRecordset(sStmt, adOpenDynamic, adLockOptimistic)
    Else
        Set rsRcd = goDBDAS.OpenRecordset(sStmt, adOpenDynamic, adLockOptimistic)
    End If
    Set rsRcd.ActiveConnection = Nothing
    
    'sUnitNumber is the primary key for MMI_ID table,
    ' we expect to have one row comes back
    If rsRcd Is Nothing Or rsRcd.RecordCount <> 1 Then
       Exit Function
    End If
    ReplaceNulls rsRcd
   
    Get_MMI_ID_Information = True
    GoTo Cleanup
ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".Get_MMI_ID_Information", App.EXEName, vbExclamation
Cleanup:
    
End Function

Function GetSex(sUnitNumber As String) As String
On Error GoTo ErrHandler
GetSex = NULLStr

Dim sStmt  As String

sStmt = "select " & _
            "SEX " & _
        "from " & _
            "mmi_id id " & _
        "where " & _
            "id.unit_number = " & sUnitNumber

goDBSqlOps.Execute sStmt

If goDBSqlOps.EOF = True Then
    GoTo Cleanup
End If

GetSex = UCase(goDBSqlOps.Value("Sex"))
GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".GetSex", App.EXEName, vbExclamation
Cleanup:
End Function

Function GetBirthDate(sUnitNumber As String) As Variant
    
    Dim sStmt  As String
    On Error GoTo ErrHandler
        
    sStmt = "select id.birth_date " & _
            "from mmi_id id " & _
            "where id.unit_number = " & sUnitNumber
    
    goDBSqlOps.Execute sStmt
    
    If goDBSqlOps.EOF = True Then
        GoTo Cleanup
    End If
    
    GetBirthDate = goDBSqlOps.Value("birth_date")
    GoTo Cleanup
    
ErrHandler:
        gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".GetBirthDate", App.EXEName, vbExclamation
    
Cleanup:

End Function
Function GetUnitNumber(sSSN As String) As String

On Error GoTo ErrHandler
GetUnitNumber = ""

Dim sStmt  As String


sStmt = "select " & _
            "id.unit_number " & _
        "from " & _
            "mmi_id id " & _
        "where " & _
            "id.ssn = '" & sSSN & "'" & _
        "order by 1 desc" 'Wilson on 7/20/01. SCR 12225

goDBSqlOps.Execute sStmt

If goDBSqlOps.EOF = True Then
    GoTo Cleanup
End If

GetUnitNumber = CStr(goDBSqlOps.Value("unit_number"))
GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".GetUnitNumber", App.EXEName, vbExclamation
Cleanup:
End Function

Private Function GetAge(sUnitNumber As String) As Integer

    On Error GoTo ErrHandler
    GetAge = 0
    
    Dim sStmt  As String
    sStmt = "select " & _
                "decode(trunc((sysdate - id.birth_date)/365),null,0,trunc((sysdate - id.birth_date)/365)) as Age " & _
            "from " & _
                "mmi_id id " & _
            "where " & _
                "id.unit_number = " & sUnitNumber
    
    goDBSqlOps.Execute sStmt
    
    If goDBSqlOps.EOF = True Then
        GoTo Cleanup
    End If
    
    GetAge = goDBSqlOps.Value("Age")
    GoTo Cleanup
    
ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".GetAge", App.EXEName, vbExclamation
Cleanup:
End Function

Function GetMaritalStatus(sUnitNumber As String) As String

    On Error GoTo ErrHandler
    
    Dim strSQL  As String
   
    GetMaritalStatus = ""

    strSQL = "SELECT NCIDS.DESCRIPTION AS MaritalStatus " & _
             "FROM MMI_ID, NCIDS " & _
             "WHERE MMI_ID.MARITAL_STATUS_CID = NCIDS.NCID AND MMI_ID.Unit_Number = " & sUnitNumber

    goDBSqlOps.Execute strSQL

    If Not (goDBSqlOps.BOF And goDBSqlOps.EOF) Then
        GetMaritalStatus = UCase(goDBSqlOps.Value("MaritalStatus") & "")
    Else
        GetMaritalStatus = vbNullString
    End If

    GoTo Cleanup
ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".GetMaritalStatus", App.EXEName, vbExclamation
Cleanup:
End Function

Public Function GetGenericData(sUnitNumber As String, _
                               sNCID As String, _
                               eReturnType As ReturnType, _
                               ResultNcid As String, _
                               ResultString As String, _
                               Optional sContextNCID As String) As Boolean
On Error GoTo ErrHandler
GetGenericData = False

Dim sStmt  As String
Dim sCONTEXT As String
    

If sContextNCID = "" Then
    sCONTEXT = "2000"
Else
    sCONTEXT = sContextNCID
End If

If eReturnType = HDDText Then
    sStmt = "select " & _
                "mg.value_ncid as ResultNCID, r.representation as ResultString " & _
            "from " & _
                "rsform_context rc, rsform r, mmi_genericdata mg " & _
            "where " & _
                "rc.preferred_score = 0 and " & _
                "rc.context_ncid = " & sCONTEXT & " and " & _
                "rc.rsform_id = r.rsform_id and " & _
                "r.ncid = mg.value_ncid and " & _
                "mg.type_ncid = " & sNCID & " and " & _
                "mg.unit_number = " & sUnitNumber
Else
    sStmt = "select " & _
                "mg.value_ncid as ResultNCID, mg.value_text as ResultString " & _
            "from " & _
                "mmi_genericdata mg " & _
            "where " & _
                "mg.type_ncid = " & sNCID & " and " & _
                "mg.unit_number = " & sUnitNumber
End If

goDBSqlOps.Execute sStmt

If goDBSqlOps.EOF = True Then
    GoTo Cleanup
End If

ResultNcid = goDBSqlOps.Value("ResultNCID")
ResultString = goDBSqlOps.Value("ResultString")

GetGenericData = True
GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".GetGenericData", App.EXEName, vbExclamation
Cleanup:
    Exit Function
End Function
Public Function GetGenericDataCollection(sUnitNumber As String) As Collection
    
    Dim cValues As New Collection
    Dim sStmt  As String
    
    On Error GoTo ErrHandler
    Set GetGenericDataCollection = New Collection
    
    sStmt = "select * from mmi_genericdata mg " & _
            "where mg.unit_number = " & sUnitNumber

    goDBSqlOps.Execute sStmt
    If goDBSqlOps.EOF = True Then GoTo Cleanup
    On Error Resume Next
    Do Until goDBSqlOps.EOF
         cValues.Add goDBSqlOps.Value("value_text"), CStr(goDBSqlOps.Value("type_ncid"))
         goDBSqlOps.MoveNext
    Loop
    
    Set GetGenericDataCollection = cValues
    GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".GetGenericDataCollection", App.EXEName, vbExclamation
Cleanup:
    Set cValues = Nothing
End Function

Public Function GetGenericIDCollection(sUnitNumber As String) As Collection
    
    Dim cValues As New Collection
    Dim sStmt  As String
    
    On Error GoTo ErrHandler
    Set GetGenericIDCollection = New Collection
    
    sStmt = "select * " & _
            "from mmi_generic_id " & _
            "where unit_number = " & sUnitNumber

    goDBSqlOps.Execute sStmt
    If goDBSqlOps.EOF = True Then GoTo Cleanup
    On Error Resume Next
    Do Until goDBSqlOps.EOF
         cValues.Add goDBSqlOps.Value("id_value"), CStr(goDBSqlOps.Value("ID_TYPE_NCID"))
         goDBSqlOps.MoveNext
    Loop
    
    Set GetGenericIDCollection = cValues
    GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".GetGenericIDCollection", App.EXEName, vbExclamation
Cleanup:
    Set cValues = Nothing
End Function

Public Function GetPatientName(sUnitNumber As String, _
                               sFirstName As String, _
                               sMiddleName As String, _
                               sLastName As String) As Boolean

On Error GoTo ErrHandler
GetPatientName = False

Dim sStmt As String

    sStmt = "select " & _
            "first_name, middle_name, last_name " & _
            "from mmi_name mn " & _
            "where mn.unit_number = " & sUnitNumber & _
                     " AND mn.NAME_TYPE_CID = 187"

    goDBSqlOps.Execute sStmt
    
    If goDBSqlOps.EOF <> True Then
        sFirstName = goDBSqlOps.Value("first_name") & ""
        sMiddleName = goDBSqlOps.Value("middle_name") & ""
        sLastName = goDBSqlOps.Value("last_name") & ""
    End If

GetPatientName = True
GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".GetPatientName", App.EXEName, vbExclamation
Cleanup:
End Function

Public Function GetOSH_INI_SETTINGS_APVALUE(APSECTION As String, _
                                    APKEY As String) As String
On Error GoTo ErrHandler
GetOSH_INI_SETTINGS_APVALUE = NULLStr

Dim sStmt  As String

sStmt = "select " & _
            "APVALUE as ResultString " & _
        "from " & _
            "OSH_INI_SETTINGS " & _
        "where " & _
            "APSECTION = " & "'" & APSECTION & "' and " & _
            "APKEY = " & "'" & APKEY & "'" & vbNewLine


goDBSqlOps.Execute sStmt

If goDBSqlOps.EOF = True Then
    GoTo Cleanup
End If

GetOSH_INI_SETTINGS_APVALUE = goDBSqlOps.Value("ResultString")

GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".GetOSH_INI_SETTINGS_APVALUE", App.EXEName, vbExclamation
Cleanup:
End Function

Public Function GetGenericId(sUnitNumber As String, _
                             sNCID As String, _
                             ResultString As String) As Boolean

On Error GoTo ErrHandler
GetGenericId = False

Dim sStmt  As String

sStmt = "select " & _
            "mi.id_value as ResultString " & _
        "from " & _
            "mmi_generic_id mi " & _
        "where " & _
            "mi.id_type_ncid = " & sNCID & " and " & _
            "mi.unit_number = " & sUnitNumber


goDBSqlOps.Execute sStmt

If goDBSqlOps.EOF = True Then
    GoTo Cleanup
End If

ResultString = goDBSqlOps.Value("ResultString")

GetGenericId = True
GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".GetGenericId", App.EXEName, vbExclamation
Cleanup:
End Function

Public Function GetHomePhone(sUnitNumber As String) As String

On Error GoTo ErrHandler
GetHomePhone = ""

Dim sStmt  As String
Dim sAreaCode As String
Dim sPhone As String


sStmt = "select " & _
            "area_code, local_number " & _
        "from " & _
            "mmi_phone mp " & _
        "where " & _
            "mp.phone_type_cid = 193 and " & _
            "mp.unit_number = " & sUnitNumber


goDBSqlOps.Execute sStmt

If goDBSqlOps.EOF = True Then
    GoTo Cleanup
End If

If goDBSqlOps.Value("area_code") <> "" Then
    sAreaCode = goDBSqlOps.Value("area_code") & "-"
Else
    sAreaCode = ""
End If

If goDBSqlOps.Value("local_number") <> "" Then
    If Len(goDBSqlOps.Value("local_number")) = 7 Then
        sPhone = sAreaCode & Mid(goDBSqlOps.Value("local_number"), 1, 3) & "-" & Mid(goDBSqlOps.Value("local_number"), 4, 4)
    End If
End If

GetHomePhone = sPhone
GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".GetHomePhone", App.EXEName, vbExclamation
Cleanup:
End Function
Public Sub GetHomeAndWorkPhone(sUnitNumber As String, sWorkPhone As String, sHomePhone As String)

    Dim sStmt  As String
    Dim sAreaCode As String
    Dim sPhone As String

    On Error GoTo ErrHandler
    sWorkPhone = ""
    sHomePhone = ""
    
    sStmt = "select area_code, local_number, phone_type_cid, internal_extension " & _
            "from mmi_phone mp " & _
            "where (mp.phone_type_cid = 193 or mp.phone_type_cid = 194) and " & _
                "mp.unit_number = " & sUnitNumber

    goDBSqlOps.Execute sStmt
    If goDBSqlOps.EOF = True Then Exit Sub
    Do Until goDBSqlOps.EOF
        If goDBSqlOps.Value("area_code") <> "" Then
           sAreaCode = goDBSqlOps.Value("area_code") & "-"
        Else
           sAreaCode = ""
        End If
    
        If goDBSqlOps.Value("local_number") <> "" Then
           If Len(goDBSqlOps.Value("local_number")) = 7 Then
              sPhone = sAreaCode & Mid(goDBSqlOps.Value("local_number"), 1, 3) & "-" & Mid(goDBSqlOps.Value("local_number"), 4, 4)
           End If
        End If
        
        If goDBSqlOps.Value("internal_extension") <> "" Then
            sPhone = sPhone & " Ext. " & goDBSqlOps.Value("internal_extension")
        End If
        
        If goDBSqlOps.Value("phone_type_cid") = "193" Then
           sHomePhone = sPhone
        ElseIf goDBSqlOps.Value("phone_type_cid") = "194" Then
           sWorkPhone = sPhone
        End If
        goDBSqlOps.MoveNext
    Loop
    
    GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".GetHomeAndWorkPhone", App.EXEName, vbExclamation
Cleanup:
End Sub

Public Function GetAddress(sUnitNumber As String, _
                           sStreet1 As String, _
                           sStreet2 As String, _
                           sCity As String, _
                           sState As String, _
                           sZip As String, _
                           sCountry As String) As Boolean

    On Error GoTo ErrHandler
    GetAddress = False
    
    Dim sStmt  As String
    sStreet1 = ""
    sStreet2 = ""
    sCity = ""
    sState = ""
    sZip = ""
    sCountry = ""
    
    sStmt = "select " & _
                "street_1, street_2, city, state, post_code, country " & _
            "from " & _
                "mmi_address " & _
            "where " & _
                "unit_number = " & sUnitNumber
    
    
    goDBSqlOps.Execute sStmt
    
    If goDBSqlOps.EOF = True Then
        GoTo Cleanup
    End If
    
    sStreet1 = goDBSqlOps.Value("street_1")
    sStreet2 = goDBSqlOps.Value("street_2")
    sCity = goDBSqlOps.Value("city")
    sState = goDBSqlOps.Value("state")
    sZip = goDBSqlOps.Value("post_code")
    sCountry = goDBSqlOps.Value("country")
    
    GetAddress = True
    GoTo Cleanup
    
ErrHandler:
        gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".GetAddress", App.EXEName, vbExclamation
Cleanup:
        Exit Function
End Function


Public Function GetService(sPatientCategory As String) As String

On Error GoTo ErrHandler

    Select Case UCase(Mid(sPatientCategory, 1, 1))
        Case "A"
            GetService = "Army"
        Case "B"
            GetService = "NOAA (National Oceanic and Atmospheric Administration)"
        Case "C"
            GetService = "Coast Guard"
        Case "F"
            GetService = "Air Force"
        Case "K"
            GetService = "Civilian"
        Case "M"
            GetService = "Marines"
        Case "N"
            GetService = "Navy"
        Case "P"
            GetService = "Public Health Service"
        Case "R"
            GetService = "NATO Reciprocal Agreement"
        Case Else
            GetService = "Unknown"
    End Select
    
GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".GetService", App.EXEName, vbExclamation
Cleanup:
    Exit Function
End Function

Public Function GetWorkPhone(sUnitNumber As String) As String

On Error GoTo ErrHandler
GetWorkPhone = ""

Dim sStmt  As String
Dim sAreaCode As String
Dim sPhone As String



sStmt = "select " & _
            "area_code, local_number, internal_extension " & _
        "from " & _
            "mmi_phone mp " & _
        "where " & _
            "mp.phone_type_cid = 194 and " & _
            "mp.unit_number = " & sUnitNumber


goDBSqlOps.Execute sStmt

If goDBSqlOps.EOF = True Then
    GoTo Cleanup
End If

If goDBSqlOps.Value("area_code") <> "" Then
    sAreaCode = goDBSqlOps.Value("area_code") & "-"
Else
    sAreaCode = ""
End If

If goDBSqlOps.Value("local_number") <> "" Then
    If Len(goDBSqlOps.Value("local_number")) = 7 Then
        sPhone = sAreaCode & Mid(goDBSqlOps.Value("local_number"), 1, 3) & "-" & Mid(goDBSqlOps.Value("local_number"), 4, 4)
    End If
End If

If goDBSqlOps.Value("internal_extension") <> "" Then
    sPhone = sPhone & " Ext. " & goDBSqlOps.Value("internal_extension")
End If


GetWorkPhone = sPhone
GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".GetWorkPhone", App.EXEName, vbExclamation
Cleanup:
    Exit Function
End Function



Function GetCommandInterest(sUnitNumber As String, Optional eReturnType As ReturnType) As String

    On Error GoTo ErrHandler
    GetCommandInterest = ""
    
    Dim sResultNCID As String
    Dim sResultString As String
    Dim eReturn As ReturnType
    
    If eReturnType = None Then
        eReturn = OriginalText
    End If
    
    If GetGenericData(sUnitNumber, mcCOMMAND_INTEREST, eReturn, sResultNCID, sResultString) Then
        If eReturnType <> NCID Then
            GetCommandInterest = sResultString
        Else
            GetCommandInterest = sResultNCID
        End If
    End If

GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".GetCommandInterest", App.EXEName, vbExclamation
Cleanup:
    Exit Function
End Function
Function GetPatientCategory(sUnitNumber As String, Optional eReturnType As ReturnType) As String

    On Error GoTo ErrHandler
    GetPatientCategory = ""

    Dim sResultNCID As String
    Dim sResultString As String
    Dim eReturn As ReturnType

    
    
    If eReturnType = None Then
        eReturn = OriginalText
    End If
    
    If GetGenericData(sUnitNumber, mcPATIENT_CATEGORY, eReturn, sResultNCID, sResultString) Then
        If eReturnType <> NCID Then
            GetPatientCategory = sResultString
        Else
            GetPatientCategory = sResultNCID
        End If
    End If

GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".GetPatientCategory", App.EXEName, vbExclamation
Cleanup:
End Function

Function GetUIC(sUnitNumber As String) As String

    On Error GoTo ErrHandler
    GetUIC = ""

    Dim sResultNCID As String
    Dim sResultString As String
    Dim eReturn As ReturnType

    eReturn = OriginalText

    
    If GetGenericData(sUnitNumber, mcUIC, eReturn, sResultNCID, sResultString) Then
            GetUIC = sResultString
    End If

GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".GetUIC", App.EXEName, vbExclamation
Cleanup:
End Function

Function GetSponsorSSN(sUnitNumber As String) As String

    On Error GoTo ErrHandler
    GetSponsorSSN = ""

    Dim sResultNCID As String
    Dim sResultString As String

    If GetGenericId(sUnitNumber, mcSPONSOR_SSN, sResultString) Then
        GetSponsorSSN = sResultString
    End If

GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".GetSponsorSSN", App.EXEName, vbExclamation
Cleanup:
End Function
Function GetPCMId(sUnitNumber As String) As String

    On Error GoTo ErrHandler
    GetPCMId = ""

    Dim sResultNCID As String
    Dim sResultString As String

    If GetGenericId(sUnitNumber, msPCM, sResultString) Then
        GetPCMId = sResultString
    End If

GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".GetPCMId", App.EXEName, vbExclamation
Cleanup:
End Function

Function GetFMP(sUnitNumber As String) As String

    On Error GoTo ErrHandler
    GetFMP = ""
    
    Dim sResultNCID As String
    Dim sResultString As String
    
    If GetGenericId(sUnitNumber, msFMP, sResultString) Then
        GetFMP = sResultString
    End If

GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".GetFMP", App.EXEName, vbExclamation
Cleanup:
End Function


Function GetVIP(sUnitNumber As String) As Boolean

    On Error GoTo ErrHandler
    GetVIP = False
    
    Dim sResultNCID As String
    Dim sResultString As String

    
    
    If GetGenericId(sUnitNumber, msVIP, sResultString) Then

        If sResultString <> "" Then
            If UCase(Mid(sResultString, 1, 1)) = "Y" Then
                GetVIP = True
            End If
        End If
    End If

GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".GetVIP", App.EXEName, vbExclamation
Cleanup:
End Function
Function GetAlternativeCare(sUnitNumber As String, Optional eReturnType As ReturnType) As String

    On Error GoTo ErrHandler
    GetAlternativeCare = ""

    Dim sResultNCID As String
    Dim sResultString As String
    Dim eReturn As ReturnType
    
    If eReturnType = None Then
        eReturn = OriginalText
    End If
    
    If GetGenericData(sUnitNumber, mcALTERNATIVE_CARE, eReturn, sResultNCID, sResultString) Then
        If eReturnType <> NCID Then
            GetAlternativeCare = sResultString
        Else
            GetAlternativeCare = sResultNCID
        End If
    End If

GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".GetAlternativeCare", App.EXEName, vbExclamation
Cleanup:
End Function
Function GetCommandSecurity(sUnitNumber As String, Optional eReturnType As ReturnType) As String
    
    On Error GoTo ErrHandler
    GetCommandSecurity = ""
    
    Dim sResultNCID As String
    Dim sResultString As String
    Dim eReturn As ReturnType
    
    If eReturnType = None Then
        eReturn = OriginalText
    End If
    
    If GetGenericData(sUnitNumber, mcCOMMAND_SECURITY, eReturn, sResultNCID, sResultString) Then
        If eReturnType <> NCID Then
            GetCommandSecurity = sResultString
        Else
            GetCommandSecurity = sResultNCID
        End If
    End If

GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".GetCommandSecurity", App.EXEName, vbExclamation
Cleanup:
End Function
Function GetRank(sUnitNumber As String, Optional eReturnType As ReturnType) As String
    
    On Error GoTo ErrHandler
    GetRank = ""
    
    Dim sResultNCID As String
    Dim sResultString As String
    Dim eReturn As ReturnType
    
    If eReturnType = None Then
        eReturn = OriginalText
    End If

    If GetGenericData(sUnitNumber, mcRANK, eReturn, sResultNCID, sResultString) Then
        If eReturnType <> NCID Then
            GetRank = sResultString
        Else
            GetRank = sResultNCID
        End If
    End If
    
GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".GetRank", App.EXEName, vbExclamation
Cleanup:
End Function

Function GetRecordsMaintainedAt(sUnitNumber As String, Optional eReturnType As ReturnType) As String
    
    On Error GoTo ErrHandler
    GetRecordsMaintainedAt = ""
    
    
    Dim sResultNCID As String
    Dim sResultString As String
    Dim eReturn As ReturnType
        
    
    If eReturnType = None Then
        eReturn = OriginalText
    End If

    If GetGenericData(sUnitNumber, mcRECORDS_MAINTAINED_AT, eReturn, sResultNCID, sResultString) Then
        If eReturnType <> NCID Then
            GetRecordsMaintainedAt = sResultString
        Else
            GetRecordsMaintainedAt = sResultNCID
        End If
    End If
    
GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".GetRecordsMaintainedAt", App.EXEName, vbExclamation
Cleanup:
End Function
Function GetFlyingStatus(sUnitNumber As String) As Boolean

    On Error GoTo ErrHandler
    GetFlyingStatus = False
    
    Dim sResultNCID As String
    Dim sResultString As String
    Dim eReturn As ReturnType

    eReturn = OriginalText

    If GetGenericData(sUnitNumber, mcFLYING_STATUS, eReturn, sResultNCID, sResultString) Then
        
        If sResultString <> "" Then
            If UCase(Mid(sResultString, 1, 1)) = "Y" Then
                GetFlyingStatus = True
            End If
        End If
    End If

GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".GetFlyingStatus", App.EXEName, vbExclamation
Cleanup:
End Function


Function GetJumpingStatus(sUnitNumber As String) As Boolean

    On Error GoTo ErrHandler
    GetJumpingStatus = False
    
    Dim sResultNCID As String
    Dim sResultString As String
    Dim eReturn As ReturnType


    eReturn = OriginalText

    If GetGenericData(sUnitNumber, mcJUMPING_STATUS, eReturn, sResultNCID, sResultString) Then
        
        If sResultString <> "" Then
            If UCase(Mid(sResultString, 1, 1)) = "Y" Then
                GetJumpingStatus = True
            End If
        End If
    End If

GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".GetJumpingStatus", App.EXEName, vbExclamation
Cleanup:
End Function
Function GetDivingStatus(sUnitNumber As String) As Boolean

    On Error GoTo ErrHandler
    GetDivingStatus = False
    
    Dim sResultNCID As String
    Dim sResultString As String
    Dim eReturn As ReturnType

    eReturn = OriginalText

    If GetGenericData(sUnitNumber, mcDIVING_STATUS, eReturn, sResultNCID, sResultString) Then
        If sResultString <> "" Then
            If UCase(Mid(sResultString, 1, 1)) = "Y" Then
                GetDivingStatus = True
            End If
        End If
    End If

GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".GetDivingStatus", App.EXEName, vbExclamation
Cleanup:
End Function
Function GetOnMobilityStatus(sUnitNumber As String) As Boolean

    On Error GoTo ErrHandler
    GetOnMobilityStatus = False
    
    Dim sResultNCID As String
    Dim sResultString As String
    Dim eReturn As ReturnType

    eReturn = OriginalText

    If GetGenericData(sUnitNumber, mcON_MOBILITY, eReturn, sResultNCID, sResultString) Then
        If sResultString <> "" Then
            If UCase(Mid(sResultString, 1, 1)) = "Y" Then
                GetOnMobilityStatus = True
            End If
        End If
    End If

GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".GetOnMobilityStatus", App.EXEName, vbExclamation
Cleanup:
End Function
Function GetSubmarineStatus(sUnitNumber As String) As Boolean

    On Error GoTo ErrHandler
    GetSubmarineStatus = False

    Dim sResultNCID As String
    Dim sResultString As String
    Dim eReturn As ReturnType

    

    eReturn = OriginalText

    If GetGenericData(sUnitNumber, mcSUBMARINE, eReturn, sResultNCID, sResultString) Then
        If sResultString <> "" Then
            If UCase(Mid(sResultString, 1, 1)) = "Y" Then
                GetSubmarineStatus = True
            End If
        End If
    End If

GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".GetSubmarineStatus", App.EXEName, vbExclamation
Cleanup:
End Function
Public Function HDD(NCID As String, Optional ContextNCID As String) As String

On Error GoTo ErrHandler
HDD = ""

Dim sStmt As String
Dim sCONTEXT As String

If ContextNCID = "" Then
    sCONTEXT = "2000"
Else
    sCONTEXT = ContextNCID
End If


sStmt = "select representation " & _
        "from rsform_context rc, rsform r " & _
        "where " & _
            "rc.preferred_score = 0 and " & _
            "rc.context_ncid = " & sCONTEXT & " and " & _
            "rc.rsform_id = r.rsform_id and " & _
            "r.ncid = " & NCID

goDBSqlOps.Execute sStmt

If goDBSqlOps.EOF = True Then GoTo Cleanup

HDD = goDBSqlOps.Value("representation")

GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".HDD", App.EXEName, vbExclamation
Cleanup:
End Function

Function GetEnrollmentFacility(sUnitNumber As String, Optional eReturnType As ReturnType) As String

    On Error GoTo ErrHandler
    GetEnrollmentFacility = ""
    
    If gobjshared.IsAppMode(modeTheater) Then
       'Enrollment Facility in Theater/Gems is recorded in OSH_INI_SETTINGS
        GetEnrollmentFacility = GetOSH_INI_SETTINGS_APVALUE("LOCAL_ENV", "FACILITY_NAME")
        GoTo Cleanup
    End If
    
    Dim sResultNCID As String
    Dim sResultString As String
    Dim eReturn As ReturnType
    
    If eReturnType = None Then
        eReturn = OriginalText
    Else
        eReturn = eReturnType
    End If
    
    If GetGenericData(sUnitNumber, mcENROLLMENT_FACILITY, eReturn, sResultNCID, sResultString) Then
        If eReturnType <> NCID Then
            GetEnrollmentFacility = sResultString
        Else
            GetEnrollmentFacility = sResultNCID
        End If
    End If

GoTo Cleanup

ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".GetEnrollmentFacility", App.EXEName, vbExclamation
Cleanup:
End Function


Private Function Init_Patient_Demographics_Rcd(oData As ADODB.Recordset) As Boolean
On Error GoTo ErrHandler
    Init_Patient_Demographics_Rcd = False
    With oData.Fields
        .Append "JOB_ID", adLongVarWChar, -1, 132
        .Append "UNIT_NUMBER", adLongVarWChar, -1, 132
        .Append "SSN", adLongVarWChar, -1, 132
        .Append "Sex", adLongVarWChar, -1, 132
        .Append "DOB", adLongVarWChar, -1, 132
        .Append "DOB_FORMATTED", adLongVarWChar, -1, 132
        .Append "Age", adLongVarWChar, -1, 132
        .Append "MARITAL_STATUS", adLongVarWChar, -1, 132
        .Append "COMMAND_INTEREST", adLongVarWChar, -1, 132
        .Append "COMMAND_SECURITY", adLongVarWChar, -1, 132
        .Append "ALTERNATIVE_CARE", adLongVarWChar, -1, 132
        .Append "FLYING_STATUS", adLongVarWChar, -1, 132
        .Append "JUMPING_STATUS", adLongVarWChar, -1, 132
        .Append "DIVING_STATUS", adLongVarWChar, -1, 132
        .Append "ON_MOBILITY_STATUS", adLongVarWChar, -1, 132
        .Append "SUBMARINE_STATUS", adLongVarWChar, -1, 132
        .Append "PATIENT_CATEGORY", adLongVarWChar, -1, 132
        .Append "UIC", adLongVarWChar, -1, 132
        .Append "SPONSOR_SSN", adLongVarWChar, -1, 132
        .Append "FMP", adLongVarWChar, -1, 132
        .Append "VIP", adLongVarWChar, -1, 132
        .Append "PCM_ID", adLongVarWChar, -1, 132
        .Append "HOME_PHONE", adLongVarWChar, -1, 132
        .Append "WORK_PHONE", adLongVarWChar, -1, 132
        .Append "SERVICE", adLongVarWChar, -1, 132
        .Append "Rank", adLongVarWChar, -1, 132
        .Append "RECORDS_MAINTAINED_AT", adLongVarWChar, -1, 132
        .Append "RELATIONSHIP_TO_SPONSOR", adLongVarWChar, -1, 132
        .Append "FIRST_NAME", adLongVarWChar, -1, 132
        .Append "MIDDLE_NAME", adLongVarWChar, -1, 132
        .Append "LAST_NAME", adLongVarWChar, -1, 132
        .Append "FULL_NAME", adLongVarWChar, -1, 132
        .Append "Religion", adLongVarWChar, -1, 132
        .Append "Race", adLongVarWChar, -1, 132
        .Append "ACTIVE_DUTY", adLongVarWChar, -1, 132
        .Append "ENROLLMENT_FACILITY", adLongVarWChar, -1, 132
        .Append "SPONSOR_UNIT_NUMBER", adLongVarWChar, -1, 132
        .Append "INSURANCE", adLongVarWChar, -1, 132
        .Append "TELPCM", adLongVarWChar, -1, 132
    End With
    Set oData.ActiveConnection = Nothing
    oData.Open
    Init_Patient_Demographics_Rcd = True
    GoTo Cleanup
ErrHandler:
    gobjshared.ShowVBError Err.Number, Err.Description, sMODULE_NAME & ".Init_Patient_Demographics_Rcd", App.EXEName, vbExclamation
Cleanup:

End Function
