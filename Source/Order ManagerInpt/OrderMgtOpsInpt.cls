VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderMgtOpsInpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private sSQLStatement       As String
Private mcClinicianNames    As Collection
Private mcClinicNames       As Collection

Private Const NCID_SPONSOR_SSN As String = "186334"
Private Const NCID_PATIENT_FMP As String = "204708"

Private Const sORAQuerySelectLabs = _
      "Select Distinct trim(n.last_name || ', ' || n.first_name || ' ' || n.middle_name) as Patient_Name, " & vbCrLf _
    & " lab.PROCESS_PRIORITY as priority, " & vbCrLf _
    & " lab.ORDER_NAME as test_name, " & vbCrLf _
    & " lab.status, lab.NORMAL_STATUS, lab.COMPLETED_DATE, lab.TIMING_START_DATE," & vbCrLf _
    & " lab.UNIT_NUMBER, lab.RESULTS, lab.OE_LAB_DATA_ID as data_id, lab.TIMING_FREQUENCY, TIMING_DURATION, " & vbCrLf _
    & " lab.ORDERED_BY, lab.TIME_ORDERED, lab.LAB_COMMENT as prv_comment," & vbCrLf _
    & " lab.SIGN_FLAG, lab.PATIENT_TYPE, lab.NURSE_ACK_FALG, lab.order_id," & vbCrLf _
    & " id.ssn as patient_ssn, " & vbCrLf _
    & " id.SEX as sex, " & vbCrLf _
    & " fmp.id_value as patient_fmp, sssn.id_value as sponsor_ssn," & vbCrLf _
    & " lab.ENC_NUM as enc_num "

Private Const sMySQLQuerySelectLabs = _
      "Select Distinct concat_ws(' ', concat_ws(', ', coalesce(n.last_name,''), coalesce(n.first_name,'')), coalesce(n.middle_name,'')) as Patient_Name, " & vbCrLf _
    & " lab.PROCESS_PRIORITY as priority, " & vbCrLf _
    & " lab.ORDER_NAME as test_name, " & vbCrLf _
    & " lab.status, lab.NORMAL_STATUS, lab.COMPLETED_DATE, lab.TIMING_START_DATE," & vbCrLf _
    & " lab.UNIT_NUMBER, lab.RESULTS, lab.OE_LAB_DATA_ID as data_id, lab.TIMING_FREQUENCY, TIMING_DURATION, " & vbCrLf _
    & " lab.ORDERED_BY, lab.TIME_ORDERED, lab.LAB_COMMENT as prv_comment," & vbCrLf _
    & " lab.SIGN_FLAG, lab.PATIENT_TYPE, lab.NURSE_ACK_FALG, lab.order_id," & vbCrLf _
    & " id.ssn as patient_ssn, " & vbCrLf _
    & " id.SEX as sex, " & vbCrLf _
    & " fmp.id_value as patient_fmp, sssn.id_value as sponsor_ssn," & vbCrLf _
    & " lab.ENC_NUM as enc_num "
    
Private Const sMySQLQueryWhereLabs = _
      " From oe_lab lab JOIN mmi_id id ON lab.unit_number = id.unit_number " _
        & " JOIN mmi_name n ON lab.unit_number = n.unit_number" _
        & " JOIN encounters e ON e.encounternumber = lab.enc_num and e.facilityncid = lab.enc_fac_ncid" _
        & " LEFT JOIN mmi_generic_id fmp ON n.unit_number = fmp.unit_number and fmp.id_type_ncid = " & NCID_PATIENT_FMP & vbCrLf _
        & " LEFT JOIN mmi_generic_id sssn ON n.unit_number = sssn.unit_number and sssn.id_type_ncid = " & NCID_SPONSOR_SSN & vbCrLf _
        & " LEFT JOIN inpatient_admission inp ON n.unit_number = inp.unit_number " _

Private Const sORAQueryWhereLabs = _
      " From  oe_lab lab,  mmi_id id, mmi_name n, mmi_generic_id fmp, " & vbCrLf _
    & "  mmi_generic_id sssn , inpatient_admission inp, encounters e " & vbCrLf _
    & " Where lab.unit_number = id.unit_number " & vbCrLf _
    & " and   lab.unit_number = n.unit_number " & vbCrLf _
    & "  and n.unit_number = fmp.unit_number(+) and fmp.id_type_ncid(+) = " & NCID_PATIENT_FMP & vbCrLf _
    & "  and n.unit_number = sssn.unit_number(+) and sssn.id_type_ncid(+) = " & NCID_SPONSOR_SSN & vbCrLf _
    & "  and n.unit_number = inp.unit_number(+)  " & vbCrLf _
    & "  and e.encounternumber = lab.enc_num " & vbCrLf _
    & "  and e.facilityncid = lab.enc_fac_ncid " & vbCrLf
    
Private Const sORAQuerySelectRads = _
      "Select Distinct trim(n.last_name || ', ' || n.first_name || ' ' || n.middle_name) as Patient_Name, " & vbCrLf _
    & " rad.PROCESS_PRIORITY as priority, " & vbCrLf _
    & " rad.ORDER_NAME as test_name, " & vbCrLf _
    & " rad.status, rad.NORMAL_STATUS,rad.COMPLETED_DATE,rad.TIMING_START_DATE," & vbCrLf _
    & " rad.UNIT_NUMBER, rad.RESULTS, rad.OE_RAD_DATA_ID as data_id, rad.TIMING_FREQUENCY,TIMING_DURATION," & vbCrLf _
    & " rad.ORDERED_BY, rad.TIME_ORDERED,rad.RAD_COMMENT as prv_comment," & vbCrLf _
    & " rad.SIGN_FLAG, rad.PATIENT_TYPE, rad.NURSE_ACK_FALG, rad.order_id," & vbCrLf _
    & " id.ssn as patient_ssn, " & vbCrLf _
    & " id.SEX as sex, " & vbCrLf _
    & " fmp.id_value as patient_fmp, sssn.id_value as sponsor_ssn," & vbCrLf _
    & " rad.ENC_NUM as enc_num "

Private Const sMySQLQuerySelectRads = _
      "Select Distinct concat_ws(' ', concat_ws(', ', coalesce(n.last_name,''), coalesce(n.first_name,'')), coalesce(n.middle_name,'')) as Patient_Name, " & vbCrLf _
    & " rad.PROCESS_PRIORITY as priority, " & vbCrLf _
    & " rad.ORDER_NAME as test_name, " & vbCrLf _
    & " rad.status, rad.NORMAL_STATUS,rad.COMPLETED_DATE,rad.TIMING_START_DATE," & vbCrLf _
    & " rad.UNIT_NUMBER, rad.RESULTS, rad.OE_RAD_DATA_ID as data_id, rad.TIMING_FREQUENCY,TIMING_DURATION," & vbCrLf _
    & " rad.ORDERED_BY, rad.TIME_ORDERED,rad.RAD_COMMENT as prv_comment," & vbCrLf _
    & " rad.SIGN_FLAG, rad.PATIENT_TYPE, rad.NURSE_ACK_FALG, rad.order_id," & vbCrLf _
    & " id.ssn as patient_ssn, " & vbCrLf _
    & " id.SEX as sex, " & vbCrLf _
    & " fmp.id_value as patient_fmp, sssn.id_value as sponsor_ssn," & vbCrLf _
    & " rad.ENC_NUM as enc_num "

Private Const sORAQueryWhereRads = _
      " From  oe_rad rad,  mmi_id id, mmi_name n, mmi_generic_id fmp," & vbCrLf _
    & "  mmi_generic_id sssn, inpatient_admission inp, encounters e " & vbCrLf _
    & " Where rad.unit_number = id.unit_number " & vbCrLf _
    & " and   rad.unit_number = n.unit_number " & vbCrLf _
    & "  and n.unit_number = fmp.unit_number(+) and fmp.id_type_ncid(+) = " & NCID_PATIENT_FMP & vbCrLf _
    & "  and n.unit_number = sssn.unit_number(+) and sssn.id_type_ncid(+) = " & NCID_SPONSOR_SSN & vbCrLf _
    & "  and n.unit_number = inp.unit_number(+)  " & vbCrLf _
    & "  and e.encounternumber = rad.enc_num " & vbCrLf _
    & "  and e.facilityncid = rad.enc_fac_ncid " & vbCrLf

Private Const sMySQLQueryWhereRads = _
      " From oe_rad rad JOIN mmi_id id ON rad.unit_number = id.unit_number " _
        & " JOIN mmi_name n ON rad.unit_number = n.unit_number " _
        & " LEFT JOIN mmi_generic_id fmp ON n.unit_number = fmp.unit_number and fmp.id_type_ncid = " & NCID_PATIENT_FMP & vbCrLf _
        & " LEFT JOIN mmi_generic_id sssn ON n.unit_number = sssn.unit_number and sssn.id_type_ncid = " & NCID_SPONSOR_SSN & vbCrLf _
        & " LEFT JOIN inpatient_admission inp ON n.unit_number = inp.unit_number " _
        & " JOIN encounters e ON e.encounternumber = rad.enc_num and e.facilityncid = rad.enc_fac_ncid " & vbCrLf

Private Const sORAQuerySelectMeds = _
      "Select Distinct trim(n.last_name || ', ' || n.first_name || ' ' || n.middle_name) as Patient_Name, " & vbCrLf _
    & " rx.ORDER_NAME as test_name, " & vbCrLf _
    & " rx.UNIT_NUMBER,  rx.ORDER_ID, rx.STATUS_NCID, " & vbCrLf _
    & " rx.ORDERED_BY, rx.TIME_ORDERED, rx.RX_COMMENT as prv_comment," & vbCrLf _
    & " rx.SIGN_FLAG, rx.PATIENT_TYPE, rx.NURSE_ACK_FALG, SIG_CODE, rx.PROCESS_PRIORITY, " & vbCrLf _
    & " rx.FILL_DATE as COMPLETED_DATE, id.ssn as patient_ssn, " & vbCrLf _
    & " id.SEX as sex, " & vbCrLf _
    & " fmp.id_value as patient_fmp, sssn.id_value as sponsor_ssn," & vbCrLf _
    & " rx.ENC_NUM as enc_num, ORDER_DATE " 'SF SCR 77233, 77234, 77235, 77236
    
Private Const sMySQLQuerySelectMeds = _
      "Select Distinct concat_ws(' ', concat_ws(', ',coalesce(n.last_name,''), coalesce(n.first_name,'')), coalesce(n.middle_name,'')) as Patient_Name, " & vbCrLf _
    & " rx.ORDER_NAME as test_name, " & vbCrLf _
    & " rx.UNIT_NUMBER,  rx.ORDER_ID, rx.STATUS_NCID, " & vbCrLf _
    & " rx.ORDERED_BY, rx.TIME_ORDERED, rx.RX_COMMENT as prv_comment," & vbCrLf _
    & " rx.SIGN_FLAG, rx.PATIENT_TYPE, rx.NURSE_ACK_FALG, SIG_CODE, rx.PROCESS_PRIORITY, " & vbCrLf _
    & " rx.FILL_DATE as COMPLETED_DATE, id.ssn as patient_ssn, " & vbCrLf _
    & " id.SEX as sex, " & vbCrLf _
    & " fmp.id_value as patient_fmp, sssn.id_value as sponsor_ssn," & vbCrLf _
    & " rx.ENC_NUM as enc_num, ORDER_DATE " 'SF SCR 77233, 77234, 77235, 77236
        
Private Const sORAQueryWhereMeds = _
      " From  oe_rx rx,  mmi_id id, mmi_name n, mmi_generic_id fmp, " & vbCrLf _
    & "  mmi_generic_id sssn, inpatient_admission inp, encounters e" & vbCrLf _
    & " Where rx.unit_number = id.unit_number " & vbCrLf _
    & " and   rx.unit_number = n.unit_number " & vbCrLf _
    & "  and n.unit_number = fmp.unit_number(+) and fmp.id_type_ncid(+) = " & NCID_PATIENT_FMP & vbCrLf _
    & "  and n.unit_number = sssn.unit_number(+) and sssn.id_type_ncid(+) = " & NCID_SPONSOR_SSN & vbCrLf _
    & "  and n.unit_number = inp.unit_number(+)  " & vbCrLf _
    & "  and e.encounternumber = rx.enc_num " & vbCrLf _
    & "  and e.facilityncid = rx.enc_fac_ncid " & vbCrLf

Private Const sMySQLQueryWhereMeds = _
      " From oe_rx rx JOIN mmi_id id ON rx.unit_number = id.unit_number " _
        & " JOIN mmi_name n ON rx.unit_number = n.unit_number " _
        & " LEFT JOIN mmi_generic_id fmp ON n.unit_number = fmp.unit_number and fmp.id_type_ncid = " & NCID_PATIENT_FMP & vbCrLf _
        & " LEFT JOIN mmi_generic_id sssn ON n.unit_number = sssn.unit_number and sssn.id_type_ncid = " & NCID_SPONSOR_SSN & vbCrLf _
        & " LEFT JOIN inpatient_admission inp ON n.unit_number = inp.unit_number" _
        & " JOIN encounters e ON e.encounternumber = rx.enc_num and e.facilityncid = rx.enc_fac_ncid " & vbCrLf _

Private Const sORAQuerySelectOthers = _
      "Select Distinct trim(n.last_name || ', ' || n.first_name || ' ' || n.middle_name) as Patient_Name, " & vbCrLf _
    & " ot.ORDER_TYPE, ot.ORDER_TYPE_DESC as test_name, " & vbCrLf _
    & " ot.UNIT_NUMBER,  ot.OE_OTR_DATA_ID, ot. ORDER_STATUS, " & vbCrLf _
    & " ot.ORDERED_BY, ot.ORDER_DATE, ot.ORDER_COMMENTS as prv_comment," & vbCrLf _
    & " ot.SIGN_FLAG,  ot.NURSE_ACK_FLAG ," & vbCrLf _
    & "  id.ssn as patient_ssn, " & vbCrLf _
    & " id.SEX as sex, " & vbCrLf _
    & " fmp.id_value as patient_fmp, sssn.id_value as sponsor_ssn," & vbCrLf _
    & " ot.ENCOUNTER_NUMBER as enc_num "
    
Private Const sMySQLQuerySelectOthers = _
      "Select Distinct concat_ws(' ', concat_ws(', ', coalesce(n.last_name,''), coalesce(n.first_name,'')), coalesce(n.middle_name,'')) as Patient_Name, " & vbCrLf _
    & " ot.ORDER_TYPE, ot.ORDER_TYPE_DESC as test_name, " & vbCrLf _
    & " ot.UNIT_NUMBER,  ot.OE_OTR_DATA_ID, ot. ORDER_STATUS, " & vbCrLf _
    & " ot.ORDERED_BY, ot.ORDER_DATE, ot.ORDER_COMMENTS as prv_comment," & vbCrLf _
    & " ot.SIGN_FLAG,  ot.NURSE_ACK_FLAG ," & vbCrLf _
    & "  id.ssn as patient_ssn, " & vbCrLf _
    & " id.SEX as sex, " & vbCrLf _
    & " fmp.id_value as patient_fmp, sssn.id_value as sponsor_ssn," & vbCrLf _
    & " ot.ENCOUNTER_NUMBER as enc_num "
    
Private Const sORAQueryWhereOthers = _
      " From  oe_other ot,  mmi_id id, mmi_name n, mmi_generic_id fmp,  " & vbCrLf _
    & "  mmi_generic_id sssn, inpatient_admission inp, encounters e " & vbCrLf _
    & " Where ot.unit_number = id.unit_number " & vbCrLf _
    & " and   ot.unit_number = n.unit_number " & vbCrLf _
    & "  and n.unit_number = fmp.unit_number(+) and fmp.id_type_ncid(+) = " & NCID_PATIENT_FMP & vbCrLf _
    & "  and n.unit_number = sssn.unit_number(+) and sssn.id_type_ncid(+) = " & NCID_SPONSOR_SSN & vbCrLf _
    & "  and n.unit_number = inp.unit_number(+)  " & vbCrLf _
    & "  and e.encounternumber = ot.encounter_number " & vbCrLf _
    & "  and e.facilityncid = ot.facility_ncid " & vbCrLf

Private Const sMySQLQueryWhereOthers = _
      " From  oe_other ot JOIN mmi_id id ON ot.unit_number = id.unit_number" _
        & " JOIN mmi_name n ON ot.unit_number = n.unit_number" _
        & " LEFT JOIN mmi_generic_id fmp ON n.unit_number = fmp.unit_number and fmp.id_type_ncid = " & NCID_PATIENT_FMP & vbCrLf _
        & " LEFT JOIN mmi_generic_id sssn ON n.unit_number = sssn.unit_number and sssn.id_type_ncid = " & NCID_SPONSOR_SSN & vbCrLf _
        & " LEFT JOIN inpatient_admission inp ON n.unit_number = inp.unit_number" _
        & " JOIN encounters e ON e.encounternumber = ot.encounter_number and e.facilityncid = ot.facility_ncid " & vbCrLf _













Public Function DoSearch() As Collection ' of Results
Dim cOrderMgt   As Collection
Dim sLabWhere As String
Dim sRadWhere As String
Dim sMedWhere As String

Dim otherOrder  As Collection
    
On Error GoTo ErrHandler

    Set cOrderMgt = New Collection
    Set otherOrder = New Collection
    Set mcClinicianNames = New Collection 'reset clinicanNames
    Set mcClinicNames = New Collection ' reset clinicnames
    
    Call goBB.Options.GetClinicans
    Call goBB.Options.GetClinics
    
    sLabWhere = GetDateWhere("TIMING_START_DATE") & getVIPPrivilege
    sRadWhere = GetDateWhere("TIMING_START_DATE") & getVIPPrivilege
    sMedWhere = GetDateWhere("ORDER_DATE") & getVIPPrivilege 'ORDER_DATE is really the start date time in the gui.

    Set DoSearch = cOrderMgt
    
    'Get Lab data
    If goBB.Options.mAllOrderType Then
        Set cOrderMgt = DoSearch_Lab(cOrderMgt, sLabWhere)
        Set cOrderMgt = DoSearch_Rad(cOrderMgt, sRadWhere)
        Set cOrderMgt = DoSearch_Med(cOrderMgt, sMedWhere)
        Set cOrderMgt = DoSearch_OtherOrders(cOrderMgt, otherOrder)
    ElseIf goBB.Options.mOrderList.Count > 0 Then
        Dim i As Integer
        For i = 1 To goBB.Options.mOrderList.Count
            If goBB.Options.mOrderList(i) = "LAB_ORDERS" Then
                Set cOrderMgt = DoSearch_Lab(cOrderMgt, sLabWhere)
            ElseIf goBB.Options.mOrderList(i) = "RAD_ORDERS" Then
                Set cOrderMgt = DoSearch_Rad(cOrderMgt, sRadWhere)
            ElseIf goBB.Options.mOrderList(i) = "MED_ORDERS" Then
                Set cOrderMgt = DoSearch_Med(cOrderMgt, sMedWhere)
            Else
                otherOrder.Add goBB.Options.mOrderList(i)
            End If
        Next i
        'Get Other Orders
        If otherOrder.Count > 0 Then
            Set cOrderMgt = DoSearch_OtherOrders(cOrderMgt, otherOrder)
        End If
    End If
    
    Set DoSearch = cOrderMgt
Cleanup:
    Set cOrderMgt = Nothing
    Exit Function
    
ErrHandler:
  goBB.SharedObj.ShowVBError Err.Number, Err.Description, "OrderMgt.DoSearch", "OrderMgt", vbExclamation

End Function

Private Function BuildWhereClinics() As String
Dim sClinician As String
Dim vVar As Variant

On Error GoTo ErrHandler

    If mcClinicNames Is Nothing Then Exit Function
    
    If mcClinicNames.Count > 0 Then
        For Each vVar In mcClinicNames
            If LenB(sClinician) = 0 Then
                sClinician = " and e.Clinicncid in (" & "'" & vVar & "'"
            Else
                sClinician = sClinician & ", " & "'" & vVar & "'"
            End If
        Next vVar
        sClinician = sClinician & ")"
    End If

    BuildWhereClinics = sClinician
    
    Exit Function
    
ErrHandler:
    goBB.SharedObj.ShowVBError Err.Number, Err.Description, "OrderMgt.BuildWhereClinics", "OrderMgt", vbExclamation
End Function

Private Function BuildWhereClinicians(ByVal tableName As String) As String
Dim sClinician As String
Dim vVar As Variant

On Error GoTo ErrHandler

    If mcClinicianNames Is Nothing Then Exit Function
    
    If mcClinicianNames.Count > 0 Then
        For Each vVar In mcClinicianNames
            If LenB(sClinician) = 0 Then
                If tableName = "ot" Then
                    sClinician = " and " & tableName & ".ORDERED_BY_NCID in (" & "'" & vVar & "'"
                Else
                    sClinician = " and " & tableName & ".PROVIDER in (" & "'" & vVar & "'"
                End If
            Else
                sClinician = sClinician & ", " & "'" & vVar & "'"
            End If
        Next vVar
        sClinician = sClinician & ")"
    End If

    BuildWhereClinicians = sClinician
    
    Exit Function
    
ErrHandler:
    goBB.SharedObj.ShowVBError Err.Number, Err.Description, "OrderMgt.BuildWhereClinicians", "OrderMgt", vbExclamation
End Function

Public Sub CriteriaAddClinic(rsClinicNCID As String)
On Error GoTo ErrHandler
    
    If mcClinicNames Is Nothing Then
        Set mcClinicNames = New Collection
    End If
    mcClinicNames.Add rsClinicNCID
    
    Exit Sub
    
ErrHandler:
    goBB.SharedObj.ShowVBError Err.Number, Err.Description, "OrderMgt.CriteriaAddClinic", "OrderMgt", vbExclamation
End Sub

Public Sub CriteriaAddClinician(rsClinicNCID As String)
On Error GoTo ErrHandler
    
    If mcClinicianNames Is Nothing Then
        Set mcClinicianNames = New Collection
    End If
    mcClinicianNames.Add rsClinicNCID
    
    Exit Sub
    
ErrHandler:
    goBB.SharedObj.ShowVBError Err.Number, Err.Description, "OrderMgt.CriteriaAddClinician", "OrderMgt", vbExclamation
End Sub
Private Function DoSearch_Lab(ByRef cOrderMgt As Collection, ByVal sWhere As String) As Collection
Dim sClinician   As String
Dim sClinics     As String
Dim rs           As Recordset
Dim objOrderResult As orderResult

On Error GoTo ErrHandler

    sWhere = sWhere & GetStatusWhereLabRad("lab")
    sClinician = BuildWhereClinicians("lab")
    sClinics = BuildWhereClinics
    If goBB.SharedObj.UseMySQLDB Then
        sSQLStatement = sMySQLQuerySelectLabs & sMySQLQueryWhereLabs & " WHERE " & Mid(sWhere, InStr(sWhere, "and ") + Len("and ")) & sClinician & _
                        sClinics & GetWardWhere & GetOccurrence & " order by patient_name"
    Else
        sSQLStatement = sORAQuerySelectLabs & sORAQueryWhereLabs & sWhere & sClinician & _
                        sClinics & GetWardWhere & GetOccurrence & " order by patient_name"
    End If
    
    Set rs = goBB.I_DAS.OpenRecordset(sSQLStatement & ", lab.TIMING_START_DATE")
    
    Do While Not rs.EOF
        Set objOrderResult = New orderResult
        objOrderResult.Deserialize rs, hlLAB_ORDERS_CODE
        cOrderMgt.Add objOrderResult
        Set objOrderResult = Nothing
        rs.MoveNext
    Loop

    Set DoSearch_Lab = cOrderMgt
    
Cleanup:
    Set cOrderMgt = Nothing
    rs.Close
    Exit Function
    
ErrHandler:
  goBB.SharedObj.ShowVBError Err.Number, Err.Description, "OrderMgt.DoSearch_Lab", "OrderMgt", vbExclamation
  Exit Function
  Resume
End Function
Private Function DoSearch_Rad(ByRef cOrderMgt As Collection, ByVal sWhere As String) As Collection
Dim sClinician  As String
Dim sClinics    As String
Dim rs          As Recordset
Dim objOrderResult As orderResult
On Error GoTo ErrHandler

    sWhere = sWhere & GetStatusWhereLabRad("rad")
    sClinician = BuildWhereClinicians("rad")
    sClinics = BuildWhereClinics
    If goBB.SharedObj.UseMySQLDB Then
        sSQLStatement = sMySQLQuerySelectRads & sMySQLQueryWhereRads & " WHERE " & Mid(sWhere, InStr(sWhere, "and ") + Len("and ")) & sClinician & _
                        sClinics & GetWardWhere & GetOccurrence & " order by patient_name"
    Else
        sSQLStatement = sORAQuerySelectRads & sORAQueryWhereRads & sWhere & sClinician & _
                        sClinics & GetWardWhere & GetOccurrence & " order by patient_name"
    End If
    
    Set rs = goBB.I_DAS.OpenRecordset(sSQLStatement & ", rad.TIMING_START_DATE")
    
    Do While Not rs.EOF
        Set objOrderResult = New orderResult
        objOrderResult.Deserialize rs, hlRAD_ORDERS_CODE
        On Error Resume Next
        cOrderMgt.Add objOrderResult
        Set objOrderResult = Nothing
        rs.MoveNext
    Loop

    Set DoSearch_Rad = cOrderMgt
    
Cleanup:
    Set cOrderMgt = Nothing
    rs.Close
    Exit Function
    
ErrHandler:
  goBB.SharedObj.ShowVBError Err.Number, Err.Description, "OrderMgt.DoSearch_Rad", "OrderMgt", vbExclamation
End Function
Public Function DoSearch_Med(ByRef cOrderMgt As Collection, ByVal sWhere As String) As Collection ' of Results
Dim objOrderResult As orderResult
Dim sClinician As String
Dim sClinics As String
Dim sStatusWhere As String
Dim rs     As Recordset

On Error GoTo ErrHandler
    
    sStatusWhere = vbNullString
        
    If goBB.Options.mSelection And allOrder Then
        sStatusWhere = vbNullString
    Else
        If (goBB.Options.mSelection And orderDispensed) Then
          sStatusWhere = sStatusWhere & " or STATUS_NCID = " & NCID_ORDERSTATUS_ACTIVE
        End If
        
        If goBB.Options.mSelection And orderActiveORPend Then
          sStatusWhere = sStatusWhere & " or STATUS_NCID = " & NCID_PARTIAL & " or STATUS_NCID = " & NCID_ORDERSTATUS_ACTIVE
        End If
        
        If goBB.Options.mSelection And orderDiscontinued Then
          sStatusWhere = sStatusWhere & " or STATUS_NCID = " & NCID_ORDERSTATUS_DISCONTINUED
        End If
        
        If goBB.Options.mSelection And orderComplete Then
          sStatusWhere = sStatusWhere & " or STATUS_NCID = " & isCompleted_Order
        End If
        
        If goBB.Options.mSelection And orderPreliminary Then
          sStatusWhere = sStatusWhere & " or STATUS_NCID = " & isPreliminary
        End If
        
        If goBB.Options.mSelection And orderAddressed Then ' scr 55048
            sStatusWhere = sStatusWhere & " or STATUS_NCID = " & isCompleted_Addr
        End If
        
        If goBB.Options.mSelection And orderUnAddressed Then
            sStatusWhere = sStatusWhere & " or STATUS_NCID = " & isCompleted_Unaddr
        End If
                
        If LenB(sStatusWhere) > 0 Then
            sWhere = sWhere & " And (" & Replace(sStatusWhere, " or ", vbNullString, , 1) & ")"
        End If
        
    End If
    
    If LenB(sStatusWhere) > 0 Then
        sWhere = sWhere & " And (" & Replace(sStatusWhere, " or ", vbNullString, , 1) & ")"
    End If
    
    sWhere = sWhere
    
    If goBB.Options.mSelection And orderSigned Then
        '** 1/17/06 -- Frank; Unsign can either be NULL or 0
        sWhere = sWhere & " And ((SIGN_FLAG is NULL ) or (SIGN_FLAG=0))"
    End If
   
  
    Set DoSearch_Med = cOrderMgt
    
    'Get med data
    sClinician = BuildWhereClinicians("rx")
    sClinics = BuildWhereClinics()
    
    If goBB.SharedObj.UseMySQLDB Then
        sSQLStatement = sMySQLQuerySelectMeds & sMySQLQueryWhereMeds & " WHERE " & Mid(sWhere, InStr(sWhere, "and ") + Len("and ")) & sClinician & _
                        sClinics & GetWardWhere & "  order by patient_name"
    Else
        sSQLStatement = sORAQuerySelectMeds & sORAQueryWhereMeds & sWhere & sClinician & _
                        sClinics & GetWardWhere & "  order by patient_name"
    End If
    
    Set rs = goBB.I_DAS.OpenRecordset(sSQLStatement)
    
    Do While Not rs.EOF
        Set objOrderResult = New orderResult
        objOrderResult.DeserializeMed rs
        On Error Resume Next
        cOrderMgt.Add objOrderResult
        Set objOrderResult = Nothing
        rs.MoveNext
    Loop
    
Cleanup:
    Set cOrderMgt = Nothing
    rs.Close
    Exit Function
    
ErrHandler:
    goBB.SharedObj.ShowVBError Err.Number, Err.Description, "OrderMgt.DoSearch_Med", "OrderMgt", vbExclamation
    
    Exit Function
    Resume
End Function
     
Private Function DoSearch_OtherOrders(ByRef cOrderMgt As Collection, ByVal oOrders As Collection) As Collection
Dim sClinician  As String
Dim sClinics    As String
Dim rs          As Recordset
Dim sWhere      As String
Dim i           As Integer
Dim objOrderResult As orderResult

On Error GoTo ErrHandler
    
    With goBB.Options
        If goBB.Options.mSelection = orderTodayIncomORMinus Then
            sWhere = sWhere & " and ORDER_DATE >= " & goBB.DataConn.SQLDate(.StartDate)
            sWhere = sWhere & " and ORDER_DATE <= " & goBB.DataConn.SQLDate(.EndDate)
        ElseIf (goBB.Options.mSelection And orderToday) Or .StartDate = 0 Then
            If DateValue(Now) <> DateValue(.EndDate) Then
                sWhere = sWhere & " and ORDER_DATE >= " & goBB.DataConn.SQLDate(DateValue(Now) + TimeValue(.StartDate))
                sWhere = sWhere & " and ORDER_DATE <= " & goBB.DataConn.SQLDate(DateValue(Now) + TimeValue(.EndDate))
            Else
                sWhere = sWhere & " and ORDER_DATE >= " & goBB.DataConn.SQLDate(.StartDate)
                sWhere = sWhere & " and ORDER_DATE <= " & goBB.DataConn.SQLDate(.EndDate)
            End If
        ElseIf .StartDate <> 0 Then
            If .EndDate = 0 Then
                If goBB.SharedObj.UseMySQLDB Then
                    sWhere = sWhere & " and DATE(ORDER_DATE) = DATE(" & SQLDate(.StartDate) & ")"
                Else
                    sWhere = sWhere & " and trunc(ORDER_DATE) = trunc(" & SQLDate(.StartDate) & ")"
                End If
            Else
                sWhere = sWhere & " and ORDER_DATE >= " & goBB.DataConn.SQLDate(.StartDate)
                sWhere = sWhere & " and ORDER_DATE <= " & goBB.DataConn.SQLDate(.EndDate)
            End If
        End If
    End With
    
    If oOrders.Count > 0 Then
        For i = 1 To oOrders.Count
            If i = 1 Then
                sWhere = sWhere & " and ORDER_TYPE in ( '" & Replace(oOrders.Item(i), "'", "''") & "'"
            Else
                sWhere = sWhere & ",'" & Replace(oOrders.Item(i), "'", "''") & "'"
            End If
        Next i
        sWhere = sWhere & ") "
    End If
    
    sWhere = sWhere & GetStatusWhereOther
    sClinician = BuildWhereClinicians("ot")
    sClinics = BuildWhereClinics
    If goBB.SharedObj.UseMySQLDB Then
        sSQLStatement = sMySQLQuerySelectOthers & sMySQLQueryWhereOthers & sWhere & sClinician & _
                        sClinics & GetWardWhere & "  order by patient_name"
    Else
        sSQLStatement = sORAQuerySelectOthers & sORAQueryWhereOthers & sWhere & sClinician & _
                sClinics & GetWardWhere & "  order by patient_name"
    End If
    Set rs = goBB.I_DAS.OpenRecordset(sSQLStatement)
    
    Do While Not rs.EOF
        Set objOrderResult = New orderResult
        Call objOrderResult.Deserialize_Other(rs)
        On Error Resume Next
        cOrderMgt.Add objOrderResult
        Set objOrderResult = Nothing
        rs.MoveNext
    Loop

    Set DoSearch_OtherOrders = cOrderMgt

Cleanup:
    Set cOrderMgt = Nothing
    rs.Close
    Exit Function
    
ErrHandler:
  goBB.SharedObj.ShowVBError Err.Number, Err.Description, "OrderMgt.DoSearch_OtherOrders", "OrderMgt", vbExclamation
End Function

Private Function getVIPPrivilege() As String
Dim sWhere As String
On Error GoTo ErrHandler

    sWhere = vbNullString
    
    If Not goBB.User.HasPrivilegeEx(Priv_Data_VIP, Priv_Read) Then
        sWhere = "   and not exists (select gd.unit_number from MMI_Generic_id gd " & vbCrLf _
                & "         where gd.id_type_ncid=14501719 and " & vbCrLf _
                & "         gd.id_value='Y' and " & vbCrLf _
                & "         gd.unit_number = id.unit_number)"
    End If

    getVIPPrivilege = sWhere

    Exit Function
    
ErrHandler:
    goBB.SharedObj.ShowVBError Err.Number, Err.Description, "OrderMgt.getVIPPrivilege", "OrderMgt", vbExclamation
End Function

Public Function getWards() As Collection
Dim sSQLStr     As String
Dim oWard       As cWard
Dim colWards    As Collection
Dim rs          As Recordset
On Error GoTo ErrHandler

    sSQLStr = "Select Distinct n.DESCRIPTION, w.WARD_NCID, wl.Ward_Location_name " & vbCrLf & _
              " from WARD w, NCIDS n, Ward_Location wl " & vbCrLf & _
              " Where w.WARD_NAME = n.NCID " & vbCrLf & _
              " And w.WARD_LOCATION_NCID = wl.WARD_LOCATION_NCID " & _
              " And w.facility_ncid = " & goBB.Logon.FacilityNCID & _
              " And DOMAINNCID = 99990504 "
              
    Set rs = goBB.I_DAS.OpenRecordset(sSQLStr)
    
    Set colWards = New Collection
    Do While Not rs.EOF
        Set oWard = New cWard
        oWard.WardName = rs("DESCRIPTION") & "/" & rs("Ward_Location_name")
        oWard.WardNCID = rs("WARD_NCID")
        colWards.Add oWard
        Set oWard = Nothing
        rs.MoveNext
    Loop
    
    Set getWards = colWards
    
Cleanup:
    Set colWards = Nothing
    rs.Close
    Exit Function
    
ErrHandler:
  goBB.SharedObj.ShowVBError Err.Number, Err.Description, "OrderMgt.getWards", "OrderMgt", vbExclamation
End Function

Public Function getOtherOrders() As Collection
Dim searchParams As DM_OrderEntry.OEOtherOrderTypeSP
On Error GoTo ErrHandler
    
    Set searchParams = New DM_OrderEntry.OEOtherOrderTypeSP

    Set getOtherOrders = goBB.SharedObj.DataAccessService.OrderEntryMapper.RetrieveOtherOrderTypeCategories(goBB.Logon.FacilityNCID)

    Exit Function
    
ErrHandler:
    goBB.SharedObj.ShowVBError Err.Number, Err.Description, "OrderMgt.getOtherOrders", "OrderMgt", vbExclamation
End Function
Private Function GetDateWhere(DateField As String) As String
Dim sWhere As String
On Error GoTo ErrHandler
    
'SF Use the TIMING_START_DATE field for LAB and RAD instead of the TIME_ORDERED for MED field.
    
    With goBB.Options
        If goBB.Options.mSelection And orderTodayIncomORMinus Then
            .EndDate = Now
            sWhere = sWhere & " and " & DateField & " >= " & goBB.DataConn.SQLDate(.StartDate)
            sWhere = sWhere & " and " & DateField & " <= " & goBB.DataConn.SQLDate(.EndDate)
            
        ElseIf (goBB.Options.mSelection And orderToday) Or .StartDate = 0 Then
            If DateValue(Now) <> DateValue(.EndDate) Then
                sWhere = sWhere & " and " & DateField & " >= " & goBB.DataConn.SQLDate(DateValue(Now) + TimeValue(.StartDate))
                sWhere = sWhere & " and " & DateField & " <= " & goBB.DataConn.SQLDate(DateValue(Now) + TimeValue(.EndDate))
            Else
                sWhere = sWhere & " and " & DateField & " >= " & goBB.DataConn.SQLDate(.StartDate)
                sWhere = sWhere & " and " & DateField & " <= " & goBB.DataConn.SQLDate(.EndDate)
            End If
            
        ElseIf .StartDate <> 0 Then
            If .EndDate = 0 Then
                If goBB.SharedObj.UseMySQLDB Then
                    sWhere = sWhere & " and DATE(" & DateField & ") = DATE(" & SQLDate(.StartDate) & ")"
                Else
                    sWhere = sWhere & " and trunc(" & DateField & ") = trunc(" & SQLDate(.StartDate) & ")"
                End If
            ElseIf .StartDate = .EndDate Then
                If goBB.SharedObj.UseMySQLDB Then
                    sWhere = sWhere & " and DATE(" & DateField & ") = DATE(sysdate())"
                Else
                    sWhere = sWhere & " and trunc(" & DateField & ") = trunc(sysdate)"
                End If
            Else
                sWhere = sWhere & " and " & DateField & " >= " & goBB.DataConn.SQLDate(.StartDate)
                sWhere = sWhere & " and " & DateField & " <= " & goBB.DataConn.SQLDate(.EndDate)
            End If
        End If
    End With
    
    GetDateWhere = sWhere
    
    Exit Function
    
ErrHandler:
    goBB.SharedObj.ShowVBError Err.Number, Err.Description, ".GetDateWhere", "OrderMgt", vbExclamation
End Function

Private Function GetStatusWhereLabRad(ByVal tableName As String) As String
Dim sStatusWhere As String
Dim sSignWhere As String
On Error GoTo ErrHandler
    
    If goBB.Options.mSelection And allOrder Then
        sStatusWhere = vbNullString
    Else
       
      If goBB.Options.mSelection And orderActiveORPend Then
        sStatusWhere = sStatusWhere & " or " & tableName & ".Status = " & isPending
      End If
      
      If goBB.Options.mSelection And orderPreliminary Then
        sStatusWhere = sStatusWhere & " or " & tableName & ".Status = " & isPreliminary
      End If
    
      If goBB.Options.mSelection And orderAddressed Then ' scr 55048
          sStatusWhere = sStatusWhere & " or " & tableName & ".Status = " & isCompleted_Addr
      End If
      
      If goBB.Options.mSelection And orderUnAddressed Then
          sStatusWhere = sStatusWhere & " or " & tableName & ".Status = " & isCompleted_Unaddr
      End If
          
      If goBB.Options.mSelection And orderDiscontinued Then
          sStatusWhere = sStatusWhere & " or " & tableName & ".Status = " & isDiscontinued
      End If
      
      If goBB.Options.mSelection And orderComplete Then
          sStatusWhere = sStatusWhere & " or " & tableName & ".Status = " & isCompleted_Order
          sStatusWhere = sStatusWhere & " or " & tableName & ".Status = " & isCompleted_Addr
          sStatusWhere = sStatusWhere & " or " & tableName & ".Status = " & isCompleted_Unaddr
      End If
      
      If goBB.Options.mSelection And orderDispensed Then
          sStatusWhere = sStatusWhere & " or " & tableName & ".Status = " & isDispensed
      End If
          
    End If
    
    If LenB(sStatusWhere) > 0 Then
        sStatusWhere = " And (" & Replace(sStatusWhere, " or ", vbNullString, , 1) & ")"
    End If
    
    If goBB.Options.mSelection And orderSigned Then
        '** 1/17/06 - Frank; Unsign can either be NULL or 0
        sSignWhere = " And ((SIGN_FLAG is NULL ) or (SIGN_FLAG=0))"
    End If
    
    GetStatusWhereLabRad = sStatusWhere & sSignWhere
    
    Exit Function
    
ErrHandler:
    goBB.SharedObj.ShowVBError Err.Number, Err.Description, ".GetStatusWhere", "OrderMgt", vbExclamation
End Function
Friend Sub UpdateSignStatus(ByVal SignIds As Collection)
Dim i       As Integer
Dim sSQLStr As String
Dim sOrder  As orderResult

On Error GoTo ErrHandler
    
    For i = 1 To SignIds.Count
        Set sOrder = SignIds.Item(i)
        
        Select Case sOrder.ResultType
            Case hlLAB_ORDERS_CODE
                sSQLStr = "UPDATE OE_LAB Set SIGN_FLAG = 1 Where OE_LAB_DATA_ID  =" & _
                          CLng(sOrder.DataId)
            Case hlRAD_ORDERS_CODE
                sSQLStr = "UPDATE OE_RAD Set SIGN_FLAG = 1 Where OE_RAD_DATA_ID  =" & _
                          CLng(sOrder.DataId)
            Case hlRX_ORDERS_CODE
                sSQLStr = "UPDATE OE_RX Set SIGN_FLAG = 1 Where ORDER_ID = " & _
                   goBB.DataConn.SQLQuote(sOrder.OrderID)
           Case Else 'Other order
                sSQLStr = "UPDATE OE_OTHER Set SIGN_FLAG = 1 Where OE_OTR_DATA_ID = " & _
                          CLng(sOrder.DataId)
        End Select
        goBB.I_DAS.ExecuteSQL sSQLStr
    Next i
    
    Exit Sub
    
ErrHandler:
    goBB.SharedObj.ShowVBError Err.Number, Err.Description, ".UpdateSignStatus", "OrderMgt", vbExclamation
End Sub

Friend Sub DiscontinueOrderStatus(ByVal sOrder As orderResult)
Dim sSQLStr As String
On Error GoTo ErrHandler
            
    Select Case sOrder.ResultType
        Case hlLAB_ORDERS_CODE
            sSQLStr = "UPDATE OE_LAB Set STATUS = " & isDiscontinued & _
                      " Where OE_LAB_DATA_ID = " & goBB.DataConn.SQLQuote(sOrder.DataId)
            goBB.I_DAS.ExecuteSQL sSQLStr

        Case hlRAD_ORDERS_CODE
            sSQLStr = "UPDATE OE_RAD Set STATUS = " & isDiscontinued & _
                      " Where OE_RAD_DATA_ID = " & goBB.DataConn.SQLQuote(sOrder.DataId)
            goBB.I_DAS.ExecuteSQL sSQLStr
        
        Case hlRX_ORDERS_CODE
                sSQLStr = "UPDATE OE_RX Set STATUS_NCID = " & CLng(NCID_ORDERSTATUS_DISCONTINUED) & " Where ORDER_ID = " & _
                        goBB.DataConn.SQLQuote(sOrder.OrderID)
                goBB.I_DAS.ExecuteSQL sSQLStr
    
                '*** 4/12/05 - Frank
                '*** SCR 67437 (Ensure that Med is DISCONTINUE IN GMES_OE also
                sSQLStr = "UPDATE MEDICATIONS Set STATUSNCID = " & CLng(NCID_ORDERSTATUS_DISCONTINUED) & " Where ORDERNUMBER = " & _
                        goBB.DataConn.SQLQuote(sOrder.OrderID)
                goBB.I_DAS.ExecuteSQL sSQLStr
                
                '**** 2/16/05 - Frank; Changed from using Order_IEN to PARENT_DATA_ID
                
                '************* This is also in GEMS_OE: OE_SessionManager.PerformAction
                'SF SCR 76748, 78766 - Added Patient_Type and Unit_Number qualifiers.
                '--Update the medications pending order to discontinued. (for continuous med orders)
                'SCR 79803 If the order status is completed, then don't discontinue any thing else.
                If sOrder.OrderSatus <> NCID_ORDERSTATUS_COMPLETE Then
                    sSQLStr = "update medications set statusncid = " & NCID_ORDERSTATUS_DISCONTINUED & " where ORDERNUMBER in"
                    sSQLStr = sSQLStr & " (select rx2.order_id from oe_rx rx1, oe_rx rx2"
                    sSQLStr = sSQLStr & " WHERE rx1.ORDER_ID = '" & sOrder.OrderID & "' and rx1.patient_type=1 and rx1.unit_number = rx2.unit_number and"
                    sSQLStr = sSQLStr & " rx2.PARENT_DATA_ID = rx1.OE_RX_DATA_ID and rx2.STATUS_NCID = " & NCID_PARTIAL & ")"
                    goBB.I_DAS.ExecuteSQL sSQLStr
                    
                    'SF SCR 76748, 78766 - Added Patient_Type and Unit_Number qualifiers.
                    'Update the oe_rx pending order to discontinued. (for continuous med orders)
                    sSQLStr = "update oe_rx set status_ncid = " & NCID_ORDERSTATUS_DISCONTINUED & " where order_id in"
                    sSQLStr = sSQLStr & " (select rx2.order_id from oe_rx rx1, oe_rx rx2"
                    sSQLStr = sSQLStr & " WHERE rx1.ORDER_ID = '" & sOrder.OrderID & "' and rx1.patient_type=1 and rx1.unit_number = rx2.unit_number and"
                    sSQLStr = sSQLStr & " rx2.PARENT_DATA_ID = rx1.OE_RX_DATA_ID and rx2.STATUS_NCID = " & NCID_PARTIAL & ")"
                    goBB.I_DAS.ExecuteSQL sSQLStr
                End If
                
        Case Else 'Other order
            sSQLStr = "UPDATE OE_OTHER Set ORDER_STATUS = " & isDiscontinued & _
                      " Where OE_OTR_DATA_ID = " & CLng(sOrder.DataId)
            goBB.I_DAS.ExecuteSQL sSQLStr
            
    End Select
    
    Exit Sub
    
ErrHandler:
    goBB.SharedObj.ShowVBError Err.Number, Err.Description, ".UpdateOrderStatus", "OrderMgt", vbExclamation
    Exit Sub
    Resume
End Sub

Friend Function GetFreqencyDetail(ByVal code As String) As String
Dim i As Integer
On Error GoTo ErrHandler
    
    If code = "One Time" Then Exit Function
    
    For i = 1 To goBB.FrequencyDetail.Count
        If code = goBB.FrequencyDetail.Item(i).code Then
            GetFreqencyDetail = goBB.FrequencyDetail.Item(i).Describtion
            Exit For
        End If
    Next i

    Exit Function
    
ErrHandler:
    goBB.SharedObj.ShowVBError Err.Number, Err.Description, ".UpdateOrderStatus", "OrderMgt", vbExclamation
End Function

Private Function GetStatusWhereOther() As String
Dim sStatusWhere As String
Dim sSignWhere As String
On Error GoTo ErrHandler
    
    If goBB.Options.mSelection And allOrder Then
        sStatusWhere = vbNullString
    Else
      If goBB.Options.mSelection And orderActiveORPend Then
        sStatusWhere = sStatusWhere & " or ORDER_STATUS = " & isPending & _
        " or ORDER_STATUS = " & isActive
      End If
      
      If goBB.Options.mSelection And orderDiscontinued Then
        sStatusWhere = sStatusWhere & " or ORDER_STATUS = " & isDiscontinued
      End If
      
      If goBB.Options.mSelection And orderComplete Then '
          sStatusWhere = sStatusWhere & " or ORDER_STATUS = " & isCompleted_Order
      End If
      If goBB.Options.mSelection And orderPreliminary Then
        sStatusWhere = sStatusWhere & " or ORDER_STATUS = " & isPreliminary
      End If
    
      If goBB.Options.mSelection And orderAddressed Then ' scr 55048
          sStatusWhere = sStatusWhere & " or ORDER_STATUS = " & isCompleted_Addr
      End If
      
      If goBB.Options.mSelection And orderUnAddressed Then
          sStatusWhere = sStatusWhere & " or ORDER_STATUS = " & isCompleted_Unaddr
      End If
      
      If goBB.Options.mSelection And orderDispensed Then
          sStatusWhere = sStatusWhere & " or ORDER_STATUS = " & isDispensed
      End If
                
    End If
    
    If LenB(sStatusWhere) > 0 Then
        sStatusWhere = " And (" & Replace(sStatusWhere, " or ", vbNullString, , 1) & ")"
    End If
    
    If goBB.Options.mSelection And orderSigned Then
        sSignWhere = " And (SIGN_FLAG <> -1)"
    End If
    
    GetStatusWhereOther = sStatusWhere & sSignWhere
    
    Exit Function
    
ErrHandler:
    goBB.SharedObj.ShowVBError Err.Number, Err.Description, ".GetStatusWhere", "OrderMgt", vbExclamation
End Function

Private Function GetWardWhere() As String
Dim oWard        As cWard
Dim i            As Integer
On Error GoTo ErrHandler
    
    If goBB.Options.mcolSearchWards.Count > 0 Then
       
        For i = 1 To goBB.Options.mcolSearchWards.Count
            Set oWard = goBB.Options.mcolSearchWards(i)
            If i = 1 Then
                GetWardWhere = GetWardWhere & " and inp.WARD_NCID in ( '" & oWard.WardNCID & "'"
            Else
                GetWardWhere = GetWardWhere & ",'" & oWard.WardNCID & "'"
            End If
        Next i
        GetWardWhere = GetWardWhere & ") and inp.DISCHARGE_DATE is null "
            
    End If
        
    Exit Function
    
ErrHandler:
    goBB.SharedObj.ShowVBError Err.Number, Err.Description, ".GetWardWhere", "OrderMgt", vbExclamation
End Function

Private Function GetOccurrence() As String

    GetOccurrence = " And (IS_ORDER_OCCURRENCE = 1 or (IS_ORDER_OCCURRENCE = 0 and TIMING_DURATION = 0)) "
'    GetOccurrence = " And (TIMING_DURATION is NULL or TIMING_DURATION = 0) "
End Function
