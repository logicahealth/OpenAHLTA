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
Attribute VB_Name = "cDAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'__ Data Access Lay for Inpatient and Hospital Setup
'__ all sql statements are here, and calle by a 'procedure name' and param array
'__ these may be replaced with stored procedures once chcsii_conn supports sp's in theater mode.

Public Enum edt_Inpatient
    edti_Diagnoses = 0
    edti_Transfer
    edti_Patient_All
    edti_Patient_one
    edti_EncDiagnoses
    edti_Discharge
    edti_UpdateDischargeInfo
    edti_UnassignedUser
    edti_PostDischargeNote
    edti_DischargeData
    edti_HasInptNoteType
End Enum

Public Enum edt_PAD
    edtp_Location = 0
    edtp_Location_New
    edtp_AdmitType
    edtp_DMIS_ID
    edtp_Specialty
    edtp_Ward
    edtp_Ward_One
    edtp_Ward_Open
    edtp_Ward_facility
    edtp_Ward_Bed
    edtp_Ward_Bed_One
    edtp_Ward_Bed_Occupied_Count
    edtp_Is_Bed_Occupied
End Enum
Private oDAS As ICHCSII_DataAccess.ICHCSII_DAS

Public Sub DiscontinueAllInpatientOrders(AdmitDate As Date, DischargeDate As Date, UnitNumber As String)
    
'This sub discontinues all inpatient orders between the AdmitDate and the DischargeDate

On Error GoTo ErrHnd
Dim oConn As CHCSII_CONN.Conn
Dim sSql As String
    
    Set oConn = New Conn
    
    sSql = "Update OE_LAB SET Status = 6 Where Unit_Number = " & UnitNumber & " And" _
        & " TIME_ORDERED >=  " & oConn.FormatDateForDBInsert(AdmitDate) & "" _
        & " And" _
        & " TIME_ORDERED <=  " & oConn.FormatDateForDBInsert(DischargeDate) & "" _
        & " And" _
        & " (Status <> 2 or Status <> 3)" _
        & " And" _
        & " Patient_Type = 1"
    oDAS.ExecuteSQL sSql
    
    sSql = "Update OE_RAD SET Status = 6 Where Unit_Number = " & UnitNumber & " And" _
        & " TIME_ORDERED >=  " & oConn.FormatDateForDBInsert(AdmitDate) & "" _
        & " And" _
        & " TIME_ORDERED <=  " & oConn.FormatDateForDBInsert(DischargeDate) & "" _
        & " And" _
        & " (Status <> 2 or Status <> 3)" _
        & " And" _
        & " Patient_Type = 1"
    oDAS.ExecuteSQL sSql
    
    sSql = "Update OE_OTHER SET Order_Status = 6 Where Unit_Number = " & UnitNumber & " And" _
        & " ORDER_DATE >=  " & oConn.FormatDateForDBInsert(AdmitDate) & "" _
        & " And" _
        & " ORDER_DATE <=  " & oConn.FormatDateForDBInsert(DischargeDate) & ""
    oDAS.ExecuteSQL sSql

    sSql = "Update OE_RX SET STATUS_NCID = 1524 Where Unit_Number = " & UnitNumber & " And" _
        & " TIME_ORDERED >=  " & oConn.FormatDateForDBInsert(AdmitDate) & "" _
        & " And" _
        & " TIME_ORDERED <=  " & oConn.FormatDateForDBInsert(DischargeDate) & "" _
        & " And" _
        & " Patient_Type = 1"
    oDAS.ExecuteSQL sSql
    
    'SCR 64024
    sSql = "Delete from DIAGNOSESRELATIONSHIPS where RELATEDDATAID in (Select oe_rx_data_id from oe_rx" _
        & " Where Unit_Number = " & UnitNumber & " And " _
        & " TIME_ORDERED >=  " & oConn.FormatDateForDBInsert(AdmitDate) & "" _
        & " And" _
        & " TIME_ORDERED <=  " & oConn.FormatDateForDBInsert(DischargeDate) & "" _
        & " And" _
        & " Patient_Type = 1)"
    oDAS.ExecuteSQL sSql

   sSql = "Update medications SET STATUSNCID = 1524 Where Unit_Number = " & UnitNumber & " And" _
        & " orderdate >=  " & oConn.FormatDateForDBInsert(AdmitDate) & "" _
        & " And" _
        & " orderdate <=  " & oConn.FormatDateForDBInsert(DischargeDate) & "" _
        & " And" _
        & " PatientType = 'I'"
    oDAS.ExecuteSQL sSql
    
Exit Sub

ErrHnd:
    gobjShared.ShowVBError Err.Number, Err.Description, "cDal.DiscontinueAllInpatientOrders", "CHCSII_PAD", vbInformation
    
    Exit Sub
    Resume
    
End Sub

Public Function GetInpatientData(datatype As edt_Inpatient, ParamArray ID()) As ADODB.Recordset
On Error GoTo ErrHandler
    Select Case datatype
    Case edti_Diagnoses
        Set GetInpatientData = oDAS.OpenRecordset("select * from inpatient_diagnosis " _
                                                 & " where inpatient_id = " & ID(0) _
                                                 & " and diagnosis_Type = " & SQLQuote(CStr(ID(1))), adOpenForwardOnly, adLockOptimistic)
    Case edti_EncDiagnoses
        Set GetInpatientData = oDAS.OpenRecordset("select DIAGNOSISNCID,SnoID,NEWFOLLOWUP,DIAGNOSESCOMMENT,PREFIX from diagnoses" _
                                  & " where ENCOUNTERNUMBER = " & ID(0) _
                                  & " and FACILITYNCID = " & ID(1), adOpenDynamic, adLockOptimistic)
    
    Case edti_Transfer
        Set GetInpatientData = oDAS.OpenRecordset("select * from inpatient_transfer  " _
                                                 & " where inpatient_id = " & ID(0) _
                                                 & " order by transfer_date", adOpenForwardOnly, adLockOptimistic)
                                                 
    Case edti_Patient_All
        Set GetInpatientData = oDAS.OpenRecordset("select * from inpatient_admission  " _
                                                 & " where discharge_date is null " _
                                                 & " And facility_ncid = " & ID(0), adOpenForwardOnly, adLockOptimistic)
    Case edti_Patient_one
        If ID(1) = Empty Then
            'SCR 67621 For old admissions we may not have the facility in which they were created.
            Set GetInpatientData = oDAS.OpenRecordset("select * from inpatient_admission " _
                                                    & " where inpatient_id = " & ID(0), adOpenDynamic, adLockOptimistic)
        Else
            Set GetInpatientData = oDAS.OpenRecordset("select * from inpatient_admission " _
                                                    & " where inpatient_id = " & ID(0) _
                                                    & " and facility_ncid = " & ID(1), adOpenDynamic, adLockOptimistic)
        End If
        
    Case edti_UnassignedUser
        Set GetInpatientData = oDAS.OpenRecordset("select User_NCID from hems_users " _
                                                & " where upper(logon_name) = '.UNASSIGNED'", adOpenForwardOnly, adLockReadOnly)
    Case edti_DischargeData
'        Set GetInpatientData = oDAS.OpenRecordset("select a.*, b.DESCRIPTION Discharge_Diagnosis" _
'                                & " from inpatient_discharge a, inpatient_diagnosis b" _
'                                & " where a.inpatient_id = " & ID(0) _
'                                & " and a.inpatient_id = b.inpatient_id (+)" _
'                                & " and b.diagnosis_Type(+) = 'DISCHARGE'", adOpenForwardOnly, adLockReadOnly)
        Set GetInpatientData = oDAS.OpenRecordset("select a.*, b.DESCRIPTION Discharge_Diagnosis from inpatient_discharge a LEFT JOIN inpatient_diagnosis b ON a.inpatient_id = b.inpatient_id and b.diagnosis_Type = 'DISCHARGE' Where a.Inpatient_ID = " & ID(0))
        
    '9-1-05 SCR 76436 Allow the app to check if the user has a specific inpatient note type for a specific appointment.
    Case edti_HasInptNoteType
        'Added the updated status: SCR 103875
        'SF signed status = 205657, Updated = 14510309
        If UBound(ID) = 1 Then
            Set GetInpatientData = oDAS.OpenRecordset("select count(*) count  from encounters where apptid = " & ID(0) & " and inpt_note_type_ncid = " & ID(1))
        Else
            'Add the Status to the query.
            Set GetInpatientData = oDAS.OpenRecordset("select count(*) count  from encounters where apptid = " & ID(0) & " and inpt_note_type_ncid = " & ID(1) & " and status in (" & ID(2) & ")")
        End If
    
    End Select
    Exit Function
ErrHandler:
    Select Case Err.Number
    Case -2147217900
        MsgBox Err.Description
        Resume Next
    Case 9    '__ paramarray has insufficient data elements
        Exit Function
    Case Else
        Resume Next
    End Select
    Exit Function
    Resume
End Function

Public Function SetInpatientData(datatype As edt_Inpatient, ParamArray ID()) As ADODB.Recordset

On Error GoTo ErrHandler
    
    Dim txt As String
    Dim rs As Recordset
    Dim sDischargeDate As String
    Dim oConn As New Conn
    
    Select Case datatype
    Case edti_Diagnoses
    '__ this just deletes all of the diag records for a given patient and opens an empty diag rs
        oDAS.ExecuteSQL "Delete from inpatient_diagnosis where inpatient_id = " & ID(0) _
                                                 & " and diagnosis_Type = " & SQLQuote(CStr(ID(1)))
        Set SetInpatientData = oDAS.OpenRecordset("select * from inpatient_diagnosis where inpatient_id = 0", adOpenDynamic, adLockOptimistic)
        
    Case edti_Transfer
        Set SetInpatientData = oDAS.OpenRecordset("select * from inpatient_transfer where inpatient_id = 0", adOpenDynamic, adLockOptimistic)

    Case edti_Patient_All
    
    Case edti_Discharge
    'Discharge the patient. Set the discharge date in the Inpatient_Admission table
    ' and in the Inpatient_Discharge table
    
'       0     Inpatient_ID, _

        Set rs = oDAS.OpenRecordset("select dischargedate from inpatient_discharge where inpatient_id = " & ID(0))
        If rs.EOF Then
            'The discharge date is not in the inpatient_discharge table so default it to today.
            Set rs = oDAS.OpenRecordset("select sysdate from dual")
           sDischargeDate = gobjComm.LocalToGmt(rs("sysdate"))
        Else
            sDischargeDate = rs("dischargedate")
        End If
        
        Call oDAS.ExecuteSQL("update inpatient_admission set discharge_date =" & oConn.SQLDate(sDischargeDate, DD_MON_YYYY) _
            & " where inpatient_id = " & ID(0))
        
    Case edti_UpdateDischargeInfo
    'Delete and Insert the data in the Inpatient_Discharge and Inpatient_Diagnosis tables
    
'       0     Inpatient_ID, _
'       1     gobjComm.LocalToGmt(Discharge_Date), _
'       2     "Discharge Disposition: " & sDisp & vbCrLf & "Discharge Comment: " & Discharge_Comment, _
'       3     Disposition_NCID, _
'       4     Disposition_Comment, _
'       5     Discussed_Items, _
'       6     Discussed_Comment, _
'       7     Discharge_Diagnosis_SnoID, _
' Removed '       8     Discharge_Diagnosis_MedcinPrefix, _
'       8     Discharge_Diagnosis_Text
        
        Call oDAS.ExecuteSQL("Delete from Inpatient_Discharge where Inpatient_ID = " & ID(0))
        
        Call oDAS.ExecuteSQL("Insert into Inpatient_discharge (Inpatient_ID, DischargeDate, DispositionNCID, DispositionTEXT, ItemsDiscussed, DiscussedTEXT,timebias)" _
                 & " values (" & ID(0) & ", " & oConn.SQLDate(ID(1), DD_MON_YYYY_HH_MM_SS) & ", " & SQLQuote(CStr(ID(3))) & ",  " & SQLQuote(CStr(ID(4))) & ", " & SQLQuote(CStr(ID(5))) & ", " & SQLQuote(CStr(ID(6))) & "," & gobjComm.GetTotalTimeBias & " )") '105144
        
        '--- Handle LONG data type for column DischargeText
        Dim oSQLEx As ICHCSII_SQL.ISqlOpsEx
        
        Set oConn = New CHCSII_CONN.Conn
        Set oSQLEx = oConn.CHCSII_SQLOPS_EX(Auto)
        
        If oConn.UseMySQL Then
            oSQLEx.Execute "Update Inpatient_discharge set dischargetext = '" & ID(2) & "' where inpatient_id = " & ID(0)
        Else
            oSQLEx.Execute "Select DISCHARGETEXT from Inpatient_discharge where Inpatient_ID = " & ID(0)
            If Not (oSQLEx.BOF And oSQLEx.EOF) Then
                oSQLEx.AppendAsChunk "DISCHARGETEXT", ID(2)
                oSQLEx.Update
            End If
        End If
        
        Set oSQLEx = Nothing
        
        'If the SnoID is "" then don't try to insert it.
        Call oDAS.ExecuteSQL("Delete from Inpatient_Diagnosis where Inpatient_ID = " & ID(0) & " AND DIAGNOSIS_TYPE = 'DISCHARGE'")
        If ID(7) <> Empty Then
            Call oDAS.ExecuteSQL("insert into inpatient_diagnosis ( INPATIENT_ID, SnoID,   DESCRIPTION, DIAGNOSIS_TYPE)" _
                     & " values (" & ID(0) & "," & ID(7) & "," & SQLQuote(CStr(ID(8))) & ",'DISCHARGE')")
        End If
    
    Case edti_PostDischargeNote
        Dim oRS As ADODB.Recordset
        Set oRS = oDAS.OpenRecordset("select postdischargetext from inpatient_discharge  where inpatient_id = " & ID(0))
        If Not oRS.EOF Then
            txt = vbNullString & oRS(0) & CStr(ID(1))
            Call oDAS.ExecuteSQL("update inpatient_discharge set postdischargetext = " & SQLQuote(txt) _
                & " where inpatient_id = " & ID(0))
        Else  '--- SCR 68731  Sherry Wang  1/27/2006
            Call oDAS.ExecuteSQL("insert into inpatient_discharge (inpatient_id, postdischargetext) " _
                & " values (" & ID(0) & ", " & SQLQuote(CStr(ID(1))) & ")")
        End If
        oRS.Close
    End Select
    Exit Function
ErrHandler:
'    Select Case Err.Number
'    Case 9    '__ paramarray has insufficient data elements
'        Exit Function
'    Case Else
'    End Select
    
    MsgBox "Application Error: " & Err.Description & " Occurred in CHCSII_PAD.cDAL.SetInpatientData."
    Exit Function
    Resume
End Function

Public Function GetPADData(datatype As edt_PAD, ParamArray ID()) As ADODB.Recordset
On Error GoTo ErrHandler
    Select Case datatype
    Case edtp_Location
        Set GetPADData = oDAS.OpenRecordset("select WARD_LOCATION_NCID,WARD_LOCATION_NAME " _
                                            & " from WARD_LOCATION " _
                                            & " where facility_ncid = " & ID(0) _
                                            & " order by WARD_LOCATION_NAME", adOpenDynamic, adLockOptimistic)
    Case edtp_AdmitType
        Set GetPADData = oDAS.OpenRecordset("select ncid, description from ncids " _
                                            & " where domainncid = 99990557 " _
                                            & " order by description", adOpenDynamic, adLockOptimistic)
    Case edtp_DMIS_ID
        Set GetPADData = oDAS.OpenRecordset("select apvalue from osh_ini_settings " _
                                            & " where application_ncid= " & NCID_HOSPITAL_SETUP _
                                            & " and APSECTION='CONFIGURATION'" _
                                            & " and APKEY='DMIS_ID'", adOpenDynamic, adLockOptimistic)
    Case edtp_Specialty
        Set GetPADData = oDAS.OpenRecordset("select ncid, description from ncids " _
                                            & " where domainncid = 99990504", adOpenDynamic, adLockOptimistic)
    Case edtp_Ward
        Set GetPADData = oDAS.OpenRecordset("Select * from ward", adOpenDynamic, adLockOptimistic)
    
    Case edtp_Ward_Open
        Set GetPADData = oDAS.OpenRecordset("select distinct ward_name,ward_ncid from ward where ward_status_code = 'OPEN' and facility_ncid = " & ID(0), adOpenForwardOnly, adLockReadOnly)
        
    Case edtp_Ward_One
        Set GetPADData = oDAS.OpenRecordset("Select * from ward where ward_NCID = " & ID(0), adOpenDynamic, adLockOptimistic)
        
    Case edtp_Ward_facility
        Set GetPADData = oDAS.OpenRecordset("select * from ward where facility_ncid = " & ID(0), adOpenDynamic, adLockOptimistic)
    
    Case edtp_Ward_Bed
        Set GetPADData = oDAS.OpenRecordset("Select * from ward_bed where ward_ncid = " & ID(0) & " order by bed_id", adOpenDynamic, adLockOptimistic)
    
    Case edtp_Ward_Bed_One
        Set GetPADData = oDAS.OpenRecordset("Select * from ward_bed where ward_ncid = " & ID(0) _
                                                & " and bed_id = " & ID(1), adOpenDynamic, adLockOptimistic)
    Case edtp_Ward_Bed_Occupied_Count
        Set GetPADData = oDAS.OpenRecordset("select count(*) from inpatient_admission " _
                                            & " where discharge_date is null " _
                                            & " and ward_ncid = " & ID(0) _
                                            & " and facility_ncid = " & ID(1), adOpenForwardOnly, adLockOptimistic)

    Case edtp_Is_Bed_Occupied
        Set GetPADData = oDAS.OpenRecordset("Select Inpatient_ID from ward_bed where ward_ncid = " & ID(0) & " and Bed_ID = " & ID(1) & " and bed_status = 'OPEN'", adOpenStatic, adLockOptimistic)
        
    End Select
    Exit Function
ErrHandler:
    Select Case Err.Number
    Case 9    '__ paramarray has insufficient data elements
        Exit Function
    Case Else
        Resume Next
    End Select
    Exit Function
    Resume
End Function

Public Function SetPADData(datatype As edt_PAD, ParamArray ID()) As ADODB.Recordset
On Error GoTo ErrHandler
    Select Case datatype
    Case edtp_DMIS_ID
         On Error GoTo 0
         Call oDAS.ExecuteSQL("delete from osh_ini_settings " _
                            & " where application_ncid= " & NCID_HOSPITAL_SETUP _
                            & " and APSECTION='CONFIGURATION'" _
                            & " and APKEY='DMIS_ID'")
         On Error GoTo ErrHandler
         Call oDAS.ExecuteSQL("insert into osh_ini_settings(application_ncid,APSECTION,APKEY,APVALUE) values " _
                            & "(" & NCID_HOSPITAL_SETUP & ",'CONFIGURATION','DMIS_ID'," & SQLQuote(CStr(ID(0))) & ")")
    Case edtp_Location_New
    
         Call oDAS.ExecuteSQL("insert into ward_location (facility_ncid,ward_location_ncid,ward_location_name) values(" _
                            & ID(0) & "," & ID(1) & "," & SQLQuote(CStr(ID(2))) & ")")
    Case edtp_Ward_Bed
        oDAS.ExecuteSQL "Delete from ward_bed where ward_ncid = " & ID(0)
        Set SetPADData = oDAS.OpenRecordset("select * from ward_bed where ward_ncid = 0", adOpenDynamic, adLockOptimistic)

    End Select
    Exit Function
ErrHandler:
    Select Case Err.Number
    Case 9    '__ paramarray has insufficient data elements
        Exit Function
    Case Else
        MsgBox "Application error: " & Err.Description & " Occurred in PAD.cDAL.SetPADData"
        
    End Select
End Function

Private Sub Class_Initialize()
     Set oDAS = GetDAS()
End Sub


Public Function GetNextID(SeqName As ICHCSII_DataAccess.eSeq) As Long
    GetNextID = oDAS.GetNextID(SeqName)
End Function

Public Function GetPatientList(CurrOptions As cOptions) As ADODB.Recordset
    '-- use the curroptions object properties to build an sql statement to get the inpatient records
    '   requested by the user-selected options
    Dim sSql          As String         '- the sql statement
    Dim ProviderList  As String         '-- list of providers to include
    Dim SpecialtyList As String         '-- list of wards to include
    Dim StatusString  As String
    Dim DateString    As String         '-- the dates to choose between
    Dim tmpdate       As Date
    Dim DateFrom      As Date           '-- start of date range to inc
    Dim DateTo        As Date           '-- end of date range to inc
    Dim vItem         As Variant
    Dim oConn As New Conn
    
    On Error GoTo ErrHnd
    
    sSql = "Select * from inpatient_Admission "
    
    With CurrOptions
'__ handle the desired wards
        If Not .Allspecialties Then  '-- dropped this And .SpecialtyNCIDs.Count > 0
            SpecialtyList = " ward_ncid in (select ward_ncid from ward where ward_name in ("
            For Each vItem In .SpecialtyNCIDs
                SpecialtyList = SpecialtyList & vItem & ","
            Next vItem
            If Right$(SpecialtyList, 1) = "," Then
                SpecialtyList = Left$(SpecialtyList, Len(SpecialtyList) - 1)  '-- drop the trailing comma
            End If
            SpecialtyList = SpecialtyList & "))"
            
        End If
        
'__ handle the desired providers
        If Not .AllProviders And Not .JustMe Then    '-- dropped this since they may select nothing 'And .ProviderNCIDs.Count > 0
            ProviderList = " attending_provider_ncid in ("
            For Each vItem In .ProviderNCIDs
                ProviderList = ProviderList & vItem & ","
            Next vItem
            If Right$(ProviderList, 1) = "," Then
                ProviderList = Left$(ProviderList, Len(ProviderList) - 1)  '-- drop the trailing comma
            End If
            ProviderList = ProviderList & ")"
            
        ElseIf .JustMe Then
            ProviderList = " attending_provider_ncid = " & gobjLogon.UserNCID
        End If
        
'__ handle all current inpatients
        If .ShowAllInpatients = True Then
            StatusString = " DISCHARGE_DATE is " & oConn.SQLDate(0) & " "
        Else
'__ handle the selected statuses
            '-- inpatient only = no discharge date
            If (.StatusSelection And os_any) Or _
                ((.StatusSelection And os_inpatient) And _
                (.StatusSelection And os_discharged)) Then
                '-- no op
            Else
                If (.StatusSelection And os_inpatient) Then
                    StatusString = " DISCHARGE_DATE is " & oConn.SQLDate(0) & " "
                ElseIf (.StatusSelection And os_discharged) Then
                    StatusString = " not (DISCHARGE_DATE is " & oConn.SQLDate(0) & ") "
                End If
            End If
    
'__ single date
            If .OneDate And Not CDbl(.SelectedDate = 0) Then
                DateString = " (ADMISSION_DATE >= " & oConn.SQLDate(gobjComm.LocalToGmt(.SelectedDate), DD_MON_YYYY_HH_MM_SS)
                DateString = DateString & " and ADMISSION_DATE < " & oConn.SQLDate(DateAdd("d", 1, gobjComm.LocalToGmt(.SelectedDate)), DD_MON_YYYY_HH_MM_SS) & ")"
            End If
            If .TodayAndInpatient Or .TodayOnly Then
                DateString = " (ADMISSION_DATE >= " & oConn.SQLDate(gobjComm.LocalToGmt(Date), DD_MON_YYYY_HH_MM_SS)
                DateString = DateString & " and ADMISSION_DATE < " & oConn.SQLDate(DateAdd("d", 1, gobjComm.LocalToGmt(Date)), DD_MON_YYYY_HH_MM_SS) & ")"
                Debug.Print DateString
            End If
'__ check the direction of the dates - it should be from oldest to newest
            If .DateRange Then
                If Not CDbl(.DateRangeFrom) = 0 And Not CDbl(.DateRangeTo) = 0 Then
                    If .DateRangeFrom > .DateRangeTo Then
                        '__ we're going backwards, flip them
                        tmpdate = .DateRangeTo
                        .DateRangeTo = .DateRangeFrom
                        .DateRangeFrom = tmpdate
                    End If
                    If Not CDbl(.DateRangeFrom) = 0 And Not CDbl(.DateRangeTo) = 0 Then
                        DateString = " (ADMISSION_DATE >= " & oConn.SQLDate(gobjComm.LocalToGmt(.DateRangeFrom), DD_MON_YYYY_HH_MM_SS)
                        DateString = DateString & " and ADMISSION_DATE <= " & oConn.SQLDate(gobjComm.LocalToGmt(.DateRangeTo), DD_MON_YYYY_HH_MM_SS) & ")"
                    End If
                End If
            End If
        End If
    End With
    

    sSql = sSql & " where ("
    
'__ datestring will Always be something.
    sSql = sSql & DateString
    If CurrOptions.TodayAndInpatient Then
    '__ handle the special case of this 'or' situation
        If gobjShared.UseMySQLDB Then
            sSql = sSql & " or (DISCHARGE_DATE  >= " & oConn.SQLDate(gobjComm.LocalToGmt(DateAdd("n", gobjComm.GetTotalTimeBias, Date)), DD_MON_YYYY) & " and DATE(DISCHARGE_DATE) < " & oConn.SQLDate(DateAdd("d", 1, gobjComm.LocalToGmt(Date)), DD_MON_YYYY) & ") or  (DISCHARGE_DATE is " & oConn.SQLDate(0) & ") )"
        Else
            sSql = sSql & " or (DISCHARGE_DATE  >= " & oConn.SQLDate(gobjComm.LocalToGmt(DateAdd("n", gobjComm.GetTotalTimeBias, Date)), DD_MON_YYYY) & " and trunc(DISCHARGE_DATE) < " & oConn.SQLDate(DateAdd("d", 1, gobjComm.LocalToGmt(Date)), DD_MON_YYYY) & ") or  (DISCHARGE_DATE is " & oConn.SQLDate(0) & ") )"
        End If
    Else
        If StatusString <> vbNullString Then
             If DateString = vbNullString Then
                sSql = sSql & StatusString
             Else
                sSql = sSql & " and " & StatusString
             End If
        End If
        '__ close the paren from the 'where' above
        sSql = sSql & ")"
    End If
    
    If ProviderList <> vbNullString Or SpecialtyList <> vbNullString Then
        sSql = sSql & " and ("
        If ProviderList <> vbNullString Then
            sSql = sSql & ProviderList & " and "
        End If
        
        If SpecialtyList <> vbNullString Then
            sSql = sSql & SpecialtyList
        Else
            '__ get rid of that dangling 'and'
            If Right$(sSql, 5) = " and " Then
                sSql = Left$(sSql, Len(sSql) - 5)
            End If
            
        End If
        '__ close the paren from the 'and' above
        sSql = sSql & ") "
    End If

    '__ restrict to this facility only
    sSql = sSql & " and facility_ncid = " & gobjLogon.FacilityNCID
    
    Set GetPatientList = oDAS.OpenRecordset(sSql)
    '__ save the sql for the print process
    CurrOptions.CurrSQL = sSql
    Exit Function

ErrHnd:
    CurrOptions.CurrSQL = vbNullString
    gobjShared.ShowVBError Err.Number, Err.Description, "Inpatient.DAL", "Improper query", vbExclamation
    
    Exit Function
    Resume
End Function
